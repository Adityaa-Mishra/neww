import os
import io
import zipfile
import tempfile
import logging
from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
from PIL import Image
from pdf2image import convert_from_path
from docx import Document
from pptx import Presentation
from pdf2docx import Converter
# import pythoncom

# Initialize Flask
app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
CONVERTED_FOLDER = "converted"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["CONVERTED_FOLDER"] = CONVERTED_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


def allowed_file(filename):
    allowed_extensions = {
        ".jpg", ".jpeg", ".png", ".pdf", ".docx", ".doc", ".pptx", ".ppt"
    }
    return any(filename.lower().endswith(ext) for ext in allowed_extensions)


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert_file():
    if "file" not in request.files:
        return "No file uploaded.", 400

    uploaded_file = request.files["file"]
    target_format = request.form.get("target")

    if not uploaded_file:
        return "No file selected.", 400

    if not target_format:
        return "Target format not selected.", 400

    if not allowed_file(uploaded_file.filename):
        return "File type not allowed.", 400

    filename = secure_filename(uploaded_file.filename)
    ext = os.path.splitext(filename)[1].lower()
    name = os.path.splitext(filename)[0]
    target_format = target_format.lower()

    temp_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    uploaded_file.save(temp_path)

    # pythoncom.CoInitialize()

    # ==============================
    # IMAGE → PNG/JPG/PDF/DOCX/PPTX
    # ==============================
    if ext in [".jpg", ".jpeg", ".png"]:
        img = Image.open(temp_path).convert("RGB")

        if target_format in ["jpg", "jpeg", "png"]:
            out_path = os.path.join(CONVERTED_FOLDER, f"{name}.{target_format}")
            img.save(out_path, target_format.upper())
            return send_file(out_path, as_attachment=True)

        elif target_format == "pdf":
            out_path = os.path.join(CONVERTED_FOLDER, f"{name}.pdf")
            img.save(out_path, "PDF", resolution=100.0)
            return send_file(out_path, as_attachment=True)

        elif target_format == "docx":
            out_path = os.path.join(CONVERTED_FOLDER, f"{name}.docx")
            doc = Document()
            doc.add_picture(temp_path)
            doc.save(out_path)
            return send_file(out_path, as_attachment=True)

        elif target_format == "pptx":
            out_path = os.path.join(CONVERTED_FOLDER, f"{name}.pptx")
            ppt = Presentation()
            slide = ppt.slides.add_slide(ppt.slide_layouts[5])
            slide.shapes.add_picture(temp_path, 0, 0, width=ppt.slide_width)
            ppt.save(out_path)
            return send_file(out_path, as_attachment=True)

    # ==============================
    # PDF → DOCX/PPTX + (NEW) PNG/JPG CONVERSION
    # ==============================
    elif ext == ".pdf":

        # ------------------------------------------------------
        # ★ NEW LOGIC ADDED HERE EXACTLY WHERE YOU WANTED ★
        # ------------------------------------------------------
        if target_format in ["png", "jpg", "jpeg"]:
            logger.info(f"PDF → {target_format}")

            images = convert_from_path(temp_path)

            # single page PDF → return direct image
            if len(images) == 1:
                fmt = "PNG" if target_format == "png" else "JPEG"
                buffer = io.BytesIO()
                images[0].save(buffer, format=fmt)
                buffer.seek(0)
                return send_file(buffer,
                                 as_attachment=True,
                                 download_name=f"{name}.{target_format}")

            # multi-page → return ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for i, img in enumerate(images):
                    img_buffer = io.BytesIO()
                    fmt = "PNG" if target_format == "png" else "JPEG"
                    img.save(img_buffer, format=fmt)
                    img_buffer.seek(0)
                    zip_file.writestr(
                        f"{name}_page_{i+1}.{target_format}",
                        img_buffer.getvalue()
                    )
            zip_buffer.seek(0)
            return send_file(zip_buffer,
                             as_attachment=True,
                             download_name=f"{name}_{target_format}s.zip",
                             mimetype="application/zip")

        # normal PDF → DOCX/PPTX
        if target_format in ["docx", "doc", "pptx", "ppt"]:
            try:
                out_path = os.path.join(CONVERTED_FOLDER, f"{name}.{target_format}")
                cv = Converter(temp_path)
                cv.convert(out_path)
                cv.close()
                return send_file(out_path, as_attachment=True)
            except Exception:
                return "Error converting PDF. Install poppler correctly.", 500

    # ==============================
    # WORD → PDF/PPT
    # ==============================
    elif ext in [".doc", ".docx"]:
        try:
            doc = Document(temp_path)
        except:
            return "Error reading .doc file. Convert to .docx first.", 400

        if target_format == "pdf":
            images = []
            for p in doc.paragraphs:
                if p.text.strip():
                    img = Image.new("RGB", (800, 200), "white")
                    images.append(img)

            if not images:
                return "Word file has no visible text.", 400

            out_path = os.path.join(CONVERTED_FOLDER, f"{name}.pdf")
            images[0].save(out_path, "PDF", save_all=True, append_images=images[1:])
            return send_file(out_path, as_attachment=True)

        elif target_format in ["ppt", "pptx"]:
            out_path = os.path.join(CONVERTED_FOLDER, f"{name}.pptx")
            ppt = Presentation()
            slide = ppt.slides.add_slide(ppt.slide_layouts[1])
            text_box = slide.shapes.add_textbox(100, 100, 500, 400)
            tf = text_box.text_frame
            tf.text = "\n".join([p.text for p in doc.paragraphs])
            ppt.save(out_path)
            return send_file(out_path, as_attachment=True)

    # ==============================
    # PPT → PDF/WORD
    # ==============================
    elif ext in [".ppt", ".pptx"]:
        ppt = Presentation(temp_path)

        if target_format == "pdf":
            images = [
                Image.new("RGB", (ppt.slide_width, ppt.slide_height), "white")
                for _ in ppt.slides
            ]

            out_path = os.path.join(CONVERTED_FOLDER, f"{name}.pdf")
            images[0].save(out_path, "PDF", save_all=True, append_images=images[1:])
            return send_file(out_path, as_attachment=True)

        elif target_format in ["doc", "docx"]:
            out_path = os.path.join(CONVERTED_FOLDER, f"{name}.docx")
            doc = Document()
            for slide in ppt.slides:
                doc.add_paragraph(
                    "\n".join([
                        shape.text for shape in slide.shapes
                        if hasattr(shape, "text")
                    ])
                )
            doc.save(out_path)
            return send_file(out_path, as_attachment=True)

    # ==============================
    # Unsupported
    # ==============================
    return f"Unsupported conversion: {ext} → {target_format}", 400


if __name__ == "__main__":
    app.run(debug=True)
