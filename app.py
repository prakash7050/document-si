import os
import io
import json
import re
from typing import Any

from flask import Flask, request, send_file, jsonify
from docx import Document
from google import genai
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

MAX_FILE_MB = 15
ALLOWED_EXTENSIONS = {".docx"}

# Gemini is only used to understand the template and map JSON keys to
# locations in the document. DOCX editing is done locally with python-docx.
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.6-flash")


def get_gemini_client() -> genai.Client:
    key = os.getenv("GEMINI_API_KEY")
    if not key:
        raise RuntimeError(
            "GEMINI_API_KEY is not set. Put it in .env or export it in your shell."
        )
    return genai.Client(api_key=key)


def json_path_get(data: Any, path: str) -> Any:
    """
    Supports:
      name
      customer.name
      experiences[0].position
      experiences.0.position
    """
    path = path.strip()
    if not path:
        return None

    # Convert [0] -> .0
    path = re.sub(r"\[(\d+)\]", r".\1", path)

    current = data
    for part in path.split("."):
        if part == "":
            continue
        if isinstance(current, dict):
            if part not in current:
                return None
            current = current[part]
        elif isinstance(current, list) and part.isdigit():
            idx = int(part)
            if idx >= len(current):
                return None
            current = current[idx]
        else:
            return None
    return current


def value_to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def replace_paragraph_text(paragraph, value: Any) -> None:
    text = value_to_text(value)

    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def replace_cell_text(cell, value: Any) -> None:
    """
    Replace text while keeping the table/cell formatting.
    We write into the first paragraph and clear the other runs/paragraphs.
    """
    text = value_to_text(value)

    if not cell.paragraphs:
        cell.add_paragraph()

    first = cell.paragraphs[0]

    if first.runs:
        first.runs[0].text = text
        for run in first.runs[1:]:
            run.text = ""
    else:
        first.add_run(text)

    # Clear any additional paragraphs so old placeholder text does not remain.
    for p in cell.paragraphs[1:]:
        for run in p.runs:
            run.text = ""


def iter_all_paragraphs(doc):
    # Body paragraphs
    for i, p in enumerate(doc.paragraphs):
        yield {"type": "paragraph", "index": i, "text": p.text}

    # Paragraphs inside tables are represented through their cells, so the
    # table-cell targets below are enough for this service.


def extract_document_structure(doc: Document) -> dict:
    paragraphs = [
        {"index": i, "text": p.text}
        for i, p in enumerate(doc.paragraphs)
    ]

    tables = []
    for ti, table in enumerate(doc.tables):
        rows = []
        for ri, row in enumerate(table.rows):
            cells = []
            for ci, cell in enumerate(row.cells):
                cells.append({
                    "row": ri,
                    "col": ci,
                    "text": cell.text,
                })
            rows.append(cells)
        tables.append({
            "table": ti,
            "rows": rows,
        })

    return {
        "paragraphs": paragraphs,
        "tables": tables,
    }


def extract_placeholder_keys(doc: Document) -> list[str]:
    """
    If a template contains {{name}}, {{loan_id}}, etc., these are used directly.
    This path needs no AI and is the most deterministic option.
    """
    found = []

    def collect(text: str):
        for key in re.findall(r"\{\{\s*([A-Za-z_][A-Za-z0-9_.\[\]]*)\s*\}\}", text):
            if key not in found:
                found.append(key)

    for p in doc.paragraphs:
        collect(p.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                collect(cell.text)

    return found


def replace_placeholders(doc: Document, data: dict) -> tuple[list[str], list[str]]:
    filled = []
    missing = []

    def replace_in_paragraph(p):
        original = p.text
        matches = re.findall(
            r"\{\{\s*([A-Za-z_][A-Za-z0-9_.\[\]]*)\s*\}\}", original
        )
        if not matches:
            return

        new_text = original
        for key in matches:
            value = json_path_get(data, key)
            if value is None:
                missing.append(key)
                continue
            new_text = re.sub(
                r"\{\{\s*" + re.escape(key) + r"\s*\}\}",
                value_to_text(value),
                new_text,
            )
            if key not in filled:
                filled.append(key)

        if new_text != original:
            replace_paragraph_text(p, new_text)

    for p in doc.paragraphs:
        replace_in_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                original = cell.text
                matches = re.findall(
                    r"\{\{\s*([A-Za-z_][A-Za-z0-9_.\[\]]*)\s*\}\}", original
                )
                if not matches:
                    continue

                new_text = original
                for key in matches:
                    value = json_path_get(data, key)
                    if value is None:
                        missing.append(key)
                        continue
                    new_text = re.sub(
                        r"\{\{\s*" + re.escape(key) + r"\s*\}\}",
                        value_to_text(value),
                        new_text,
                    )
                    if key not in filled:
                        filled.append(key)

                if new_text != original:
                    replace_cell_text(cell, new_text)

    return sorted(set(filled)), sorted(set(missing))


def build_mapping_with_gemini(doc_structure: dict, data: dict) -> dict:
    """
    Ask Gemini to map JSON source paths to actual DOCX locations.
    Returns only executable assignments; document editing stays local.
    """
    client = get_gemini_client()

    # google-genai response_schema does not accept JSON Schema unions like
    # ["integer", "null"]. We use -1 as "not applicable".
    schema = {
        "type": "object",
        "properties": {
            "assignments": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "source": {"type": "string"},
                        "target_type": {
                            "type": "string",
                            "enum": ["paragraph", "table_cell"],
                        },
                        "paragraph_index": {
                            "type": "integer",
                            "description": "Paragraph index; use -1 when not applicable.",
                        },
                        "table_index": {
                            "type": "integer",
                            "description": "Table index; use -1 when not applicable.",
                        },
                        "row": {
                            "type": "integer",
                            "description": "Table row index; use -1 when not applicable.",
                        },
                        "col": {
                            "type": "integer",
                            "description": "Table column index; use -1 when not applicable.",
                        },
                    },
                    "required": [
                        "source",
                        "target_type",
                        "paragraph_index",
                        "table_index",
                        "row",
                        "col",
                    ],
                },
            }
        },
        "required": ["assignments"],
    }

    prompt = f"""
You are a DOCX form-field mapping engine.

Map JSON DATA KEYS to locations in the supplied DOCX template.

RULES:
1. Use ONLY locations that actually exist in DOCUMENT STRUCTURE.
2. Source is a JSON path, e.g. name, customer.name,
   professional_qualifications[0].qualification,
   professional_experience[0].position.
3. For arrays/repeated records, map items in order to corresponding repeated
   blocks in the document.
4. Never invent paragraph/table/row/column indexes.
5. Prefer the blank value cell next to a visible label.
6. A row containing "Qualification:" should map qualification data to its
   associated blank value area.
7. A repeated "Position:" block should receive the corresponding position.
8. A row containing "to" can contain from/start and to/end date values.
9. Do not map unrelated JSON keys to arbitrary locations.
10. paragraph_index, table_index, row and col MUST always be integers.
    Use -1 when a field does not apply.
11. Return ONLY executable assignments. No explanation.

DOCUMENT STRUCTURE:
{json.dumps(doc_structure, ensure_ascii=False, indent=2)}

DATA:
{json.dumps(data, ensure_ascii=False, indent=2)}
"""

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config={
            "response_mime_type": "application/json",
            "response_schema": schema,
        },
    )

    if not response.text:
        raise RuntimeError("Gemini returned an empty response.")

    return json.loads(response.text)


def apply_assignments(doc: Document, assignments: list[dict], data: dict):
    filled_sources = []
    unmatched_sources = []

    # De-duplicate because merged Word cells can appear multiple times in
    # python-docx's row.cells representation.
    seen_targets = set()

    for item in assignments:
        source = item["source"]
        value = json_path_get(data, source)

        if value is None:
            if source not in unmatched_sources:
                unmatched_sources.append(source)
            continue

        target_type = item["target_type"]

        if target_type == "paragraph":
            idx = item["paragraph_index"]
            if idx is None or idx < 0 or idx >= len(doc.paragraphs):
                continue

            target_id = ("p", idx)
            if target_id in seen_targets:
                continue

            replace_paragraph_text(doc.paragraphs[idx], value)
            seen_targets.add(target_id)
            filled_sources.append(source)

        elif target_type == "table_cell":
            ti = item["table_index"]
            ri = item["row"]
            ci = item["col"]

            if ti is None or ri is None or ci is None:
                continue
            if ti < 0 or ti >= len(doc.tables):
                continue

            table = doc.tables[ti]
            if ri < 0 or ri >= len(table.rows):
                continue
            if ci < 0 or ci >= len(table.rows[ri].cells):
                continue

            cell = table.rows[ri].cells[ci]

            # Merged cells can repeat the same underlying XML cell.
            target_id = ("c", id(cell._tc))
            if target_id in seen_targets:
                continue

            replace_cell_text(cell, value)
            seen_targets.add(target_id)
            filled_sources.append(source)

    return sorted(set(filled_sources)), sorted(set(unmatched_sources))


def validate_upload(file_storage):
    if not file_storage or not file_storage.filename:
        raise ValueError("file is required")

    filename = file_storage.filename
    if not filename.lower().endswith(".docx"):
        raise ValueError("Only .docx files are supported")

    # Check size without trusting the extension.
    stream = file_storage.stream
    current = stream.tell()
    stream.seek(0, io.SEEK_END)
    size = stream.tell()
    stream.seek(current)

    if size > MAX_FILE_MB * 1024 * 1024:
        raise ValueError(f"File too large. Maximum is {MAX_FILE_MB} MB.")


@app.get("/health")
def health():
    return jsonify({
        "success": True,
        "service": "DOCX Auto Fill API",
        "ai_provider": "gemini",
        "model": GEMINI_MODEL,
        "gemini_key_configured": bool(os.getenv("GEMINI_API_KEY")),
    })


@app.get("/")
def home():
    return jsonify({
        "message": "DOCX Auto Fill API is running",
        "endpoints": {
            "POST /analyze": "multipart/form-data: file=<template.docx>",
            "POST /fill-docx": "multipart/form-data: file=<template.docx>, data=<JSON string>",
        },
    })


@app.post("/analyze")
def analyze():
    try:
        file = request.files.get("file")
        validate_upload(file)

        file.stream.seek(0)
        doc = Document(file.stream)
        structure = extract_document_structure(doc)
        placeholders = extract_placeholder_keys(doc)

        return jsonify({
            "success": True,
            "placeholders": placeholders,
            "document": structure,
            "note": (
                "If placeholders are present, use {{field_name}} in the DOCX "
                "for deterministic filling. Otherwise /fill-docx uses Gemini "
                "to map JSON keys to labels/table locations."
            ),
        })

    except Exception as exc:
        return jsonify({"success": False, "error": str(exc)}), 400


@app.post("/fill-docx")
def fill_docx():
    try:
        file = request.files.get("file")
        validate_upload(file)

        raw_data = request.form.get("data")
        if not raw_data:
            return jsonify({
                "success": False,
                "error": "data field is required and must contain JSON"
            }), 400

        try:
            data = json.loads(raw_data)
        except json.JSONDecodeError as exc:
            return jsonify({
                "success": False,
                "error": f"Invalid JSON in data field: {exc}"
            }), 400

        if not isinstance(data, dict):
            return jsonify({
                "success": False,
                "error": "data must be a JSON object"
            }), 400

        file.stream.seek(0)
        doc = Document(file.stream)

        # 1) Deterministic placeholder mode.
        placeholders = extract_placeholder_keys(doc)
        filled, missing = replace_placeholders(doc, data)

        # 2) AI mapping mode for templates that have labels but no {{...}}.
        mapping_used = False
        if not placeholders:
            structure = extract_document_structure(doc)
            mapping = build_mapping_with_gemini(structure, data)
            ai_filled, ai_unmatched = apply_assignments(
                doc, mapping.get("assignments", []), data
            )
            filled = sorted(set(filled + ai_filled))

            # Only report keys that were actually not mapped.
            all_input_keys = list(data.keys())
            mapped_top_keys = {s.split(".")[0].split("[")[0] for s in ai_filled}
            unmatched = [
                k for k in all_input_keys if k not in mapped_top_keys
            ]
            missing = sorted(set(missing + ai_unmatched + unmatched))
            mapping_used = True

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        response = send_file(
            output,
            as_attachment=True,
            download_name="filled_document.docx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
        )

        response.headers["X-Filled-Fields"] = json.dumps(
            sorted(set(filled)), ensure_ascii=True
        )
        response.headers["X-Unmatched-Fields"] = json.dumps(
            sorted(set(missing)), ensure_ascii=True
        )
        response.headers["X-Mapping-Mode"] = (
            "placeholder" if placeholders else "gemini"
        )
        return response

    except Exception as exc:
        return jsonify({"success": False, "error": str(exc)}), 500


if __name__ == "__main__":
    # Local development only.
    app.run(host="0.0.0.0", port=5000, debug=True)
