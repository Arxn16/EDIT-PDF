import os
import shutil
import tempfile
import urllib.parse

from flask import Flask, jsonify, render_template, request, send_file

from converter import DEFAULT_FONT, analyze, convert
from editor import bp as editor_bp

app = Flask(__name__)
app.register_blueprint(editor_bp)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "200")) * 1024 * 1024

FONTS = [
    ("Sarabun TH Compat", "Sarabun TH Compat – แบบ TH Sarabun ฝังในไฟล์ (แนะนำ)"),
    ("TH Sarabun New", "TH Sarabun New"), ("TH SarabunPSK", "TH SarabunPSK"), ("TH SarabunIT๙", "TH SarabunIT๙"),
    ("Tahoma", "Tahoma"), ("Leelawadee UI", "Leelawadee UI"), ("Thonburi", "Thonburi"),
    ("Angsana New", "Angsana New"), ("Cordia New", "Cordia New"),
]


def _save_upload():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".pdf"):
        return None, None, None
    work = tempfile.mkdtemp(prefix="pdf2word_")
    pdf_path = os.path.join(work, "input.pdf")
    f.save(pdf_path)
    return f.filename, work, pdf_path


@app.get("/")
def index():
    return render_template("index.html", fonts=FONTS)


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/analyze")
def do_analyze():
    name, work, pdf_path = _save_upload()
    if not work:
        return jsonify(error="กรุณาเลือกไฟล์ .pdf"), 400
    try:
        return jsonify(analyze(pdf_path))
    except Exception as e:  # noqa: BLE001
        return jsonify(error=f"อ่านไฟล์ไม่ได้: {e}"), 500
    finally:
        shutil.rmtree(work, ignore_errors=True)


@app.post("/convert")
def do_convert():
    name, work, pdf_path = _save_upload()
    if not work:
        return jsonify(error="กรุณาเลือกไฟล์ .pdf"), 400
    mode = request.form.get("mode", "auto")
    font = request.form.get("font", DEFAULT_FONT)
    dict_fix = request.form.get("dict_fix", "1") == "1"
    out_path = os.path.join(work, "output.docx")
    try:
        info = convert(pdf_path, out_path, mode, font, dict_fix)
    except Exception as e:  # noqa: BLE001
        shutil.rmtree(work, ignore_errors=True)
        app.logger.exception("convert failed")
        return jsonify(error=f"แปลงไม่สำเร็จ: {e}"), 500

    resp = send_file(out_path, as_attachment=True,
                     download_name=os.path.splitext(name)[0] + ".docx")
    resp.headers["X-Method"] = info["method"]
    resp.headers["X-Detail"] = urllib.parse.quote(info["detail"])
    resp.headers["Access-Control-Expose-Headers"] = "X-Method, X-Detail"
    resp.call_on_close(lambda: shutil.rmtree(work, ignore_errors=True))
    return resp


@app.errorhandler(413)
def too_large(_):
    return jsonify(error="ไฟล์ใหญ่เกินกำหนด"), 413


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
