#!/usr/bin/env python3
"""SCU Law Review Footnote Formatter — Flask web application."""

import io
import os
import tempfile
import traceback
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from footnote_processor import process_footnotes_to_pdf, process_footnotes_to_xlsx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload limit


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    # ── Validate file ──────────────────────────────────────────────────────────
    if "docx" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["docx"]
    if not file or file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    if not file.filename.lower().endswith(".docx"):
        return jsonify({"error": "File must be a .docx file."}), 400

    # ── Validate footnote range ────────────────────────────────────────────────
    try:
        start_fn = int(request.form["start_fn"])
        end_fn   = int(request.form["end_fn"])
    except (KeyError, ValueError):
        return jsonify({"error": "Footnote numbers must be integers."}), 400

    if start_fn < 1:
        return jsonify({"error": "Starting footnote must be ≥ 1."}), 400
    if end_fn < start_fn:
        return jsonify({"error": "Ending footnote must be ≥ starting footnote."}), 400
    if end_fn - start_fn > 500:
        return jsonify({"error": "Range too large — max 500 footnotes at once."}), 400

    # ── Process ────────────────────────────────────────────────────────────────
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False, dir="/tmp")
    try:
        file.save(tmp.name)
        tmp.close()

        pdf_bytes, _fn_count, _src_count = process_footnotes_to_pdf(
            tmp.name, start_fn, end_fn, file.filename
        )

        stem     = Path(file.filename).stem
        pdf_name = f"{stem}_footnotes_{start_fn}-{end_fn}.pdf"

        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=pdf_name,
        )

    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500

    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


@app.route("/process_xlsx", methods=["POST"])
def process_xlsx():
    if "docx" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["docx"]
    if not file or file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    if not file.filename.lower().endswith(".docx"):
        return jsonify({"error": "File must be a .docx file."}), 400

    try:
        start_fn = int(request.form["start_fn"])
        end_fn   = int(request.form["end_fn"])
    except (KeyError, ValueError):
        return jsonify({"error": "Footnote numbers must be integers."}), 400

    if start_fn < 1:
        return jsonify({"error": "Starting footnote must be ≥ 1."}), 400
    if end_fn < start_fn:
        return jsonify({"error": "Ending footnote must be ≥ starting footnote."}), 400
    if end_fn - start_fn > 500:
        return jsonify({"error": "Range too large — max 500 footnotes at once."}), 400

    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False, dir="/tmp")
    try:
        file.save(tmp.name)
        tmp.close()

        xlsx_bytes, _fn_count, _src_count = process_footnotes_to_xlsx(
            tmp.name, start_fn, end_fn, file.filename
        )

        stem      = Path(file.filename).stem
        xlsx_name = f"{stem}_footnotes_{start_fn}-{end_fn}.xlsx"

        return send_file(
            io.BytesIO(xlsx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=xlsx_name,
        )

    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500

    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
