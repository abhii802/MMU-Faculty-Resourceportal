"""
Faculty Resource Console — generator backend.

Exposes the three private scripts as three HTTP endpoints:
  POST /api/generate/paper   { "topic": "..." }  -> .docx
  POST /api/generate/ppt     { "topic": "..." }  -> .pptx
  POST /api/generate/notes   { "topic": "..." }  -> .docx

The scripts themselves are never sent to the browser — only the
generated file comes back. Keep this app.py + the three .py files
private; only the deployed URL goes into the webpage.

Run locally:
    pip install -r requirements.txt
    export OPENAI_API_KEY="..."          # optional, enables AI-enhanced content
    export ALLOWED_ORIGIN="https://YOUR-USERNAME.github.io"
    python app.py
"""

import os
import uuid
import traceback
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

from research_paper_generator import generate_paper
from cs_pptx_generator import generate_ppt
from cs_notes_generator import generate_notes_docx

app = Flask(__name__)

# Lock this down to your GitHub Pages origin in production.
# Comma-separated list is supported, e.g. "https://a.github.io,https://b.com"
ALLOWED_ORIGIN = os.getenv("ALLOWED_ORIGIN", "*")
CORS(app, resources={r"/api/*": {"origins": ALLOWED_ORIGIN.split(",")}})

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "generated")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def _get_topic():
    data = request.get_json(silent=True) or {}
    topic = (data.get("topic") or "").strip()
    if not topic:
        raise ValueError("Missing 'topic' in request body.")
    return topic


@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "faculty-resource-console-backend"})


@app.route("/api/generate/paper", methods=["POST"])
def api_generate_paper():
    try:
        topic = _get_topic()
        filepath = generate_paper(topic)
        return send_file(
            filepath,
            as_attachment=True,
            download_name=os.path.basename(filepath),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception:
        traceback.print_exc()
        return jsonify({"error": "Generation failed. Please try again."}), 500


@app.route("/api/generate/ppt", methods=["POST"])
def api_generate_ppt():
    try:
        topic = _get_topic()
        safe_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in topic).strip().replace(" ", "_")
        output_path = os.path.join(OUTPUT_DIR, f"{safe_name}_{uuid.uuid4().hex[:6]}.pptx")
        filepath = generate_ppt(topic, output_path=output_path)
        return send_file(
            filepath,
            as_attachment=True,
            download_name=os.path.basename(filepath),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception:
        traceback.print_exc()
        return jsonify({"error": "Generation failed. Please try again."}), 500


@app.route("/api/generate/notes", methods=["POST"])
def api_generate_notes():
    try:
        topic = _get_topic()
        safe_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in topic).strip().replace(" ", "_")
        output_path = os.path.join(OUTPUT_DIR, f"{safe_name}_{uuid.uuid4().hex[:6]}.docx")
        filepath = generate_notes_docx(topic, output_path=output_path)
        return send_file(
            filepath,
            as_attachment=True,
            download_name=os.path.basename(filepath),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception:
        traceback.print_exc()
        return jsonify({"error": "Generation failed. Please try again."}), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
