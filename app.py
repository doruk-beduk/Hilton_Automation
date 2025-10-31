#!/usr/bin/env python3
import os
import tempfile
import uuid
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

# Import your generator
from MOP_Generator import generate_mop

ALLOWED_EXTS = {"txt", "pdf"}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16 MB

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")  # replace in prod
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTS

@app.errorhandler(RequestEntityTooLarge)
def handle_413(_e):
    flash("The uploaded file is too large (max 16 MB).")
    return redirect(url_for("index"))

@app.route("/", methods=["GET"])
def index():
    # Pass limits/allowlist to template for helper text
    return render_template(
        "index.html",
        max_mb=MAX_CONTENT_LENGTH // (1024 * 1024),
        allowed_exts=", ".join(sorted(ALLOWED_EXTS))
    )

@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        flash("No file part in request.")
        return redirect(url_for("index"))

    f = request.files["file"]
    if f.filename == "":
        flash("No file selected.")
        return redirect(url_for("index"))

    if not allowed_file(f.filename):
        flash("Please upload a .txt or .pdf ASA config.")
        return redirect(url_for("index"))

    filename = secure_filename(f.filename)

    with tempfile.TemporaryDirectory() as tmpdir:
        in_path = os.path.join(tmpdir, filename)
        f.save(in_path)

        out_basename = f"MOP_{uuid.uuid4().hex}.docx"
        out_path = os.path.join(tmpdir, out_basename)

        try:
            generate_mop(in_path, out_path)
        except Exception as e:
            app.logger.exception("Failed to generate MOP")
            flash(f"Failed to generate MOP: {e}")
            return redirect(url_for("index"))

        return send_file(
            out_path,
            as_attachment=True,
            download_name=out_basename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5050)), debug=True)
