from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
from PIL import Image
from docx import Document
from docx.shared import Inches
from pptx import Presentation
from pptx.util import Inches as PPTInches
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io
import os
import sys
import logging

# Additional imports for new conversions
from pdf2docx import Converter
from pdf2image import convert_from_path

# Setup logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

UPLOAD_FOLDER = "uploads"
CONVERTED_FOLDER = "converted"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/convert", methods=["POST", "OPTIONS"], strict_slashes=False)
def convert_file():
    if request.method == "OPTIONS":
        return ('', 200)

    if "file" not in request.files:
        return "No file uploaded", 400

    file = request.files["file"]
    target = request.form.get("target", "").strip().lower()

    if not file or file.filename == "":
        return "No selected file", 400
    if not target:
        return "No target format specified", 400

    filename = secure_filename(file.filename)
    name, ext = os.path.splitext(filename)
    ext = ext.lower()

    logger.info(f"=== NEW CONVERSION REQUEST ===")
    logger.info(f"File: {filename} | Ext: {ext} | Target: {target}")

    temp_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(temp_path)
    logger.info(f"File saved: {temp_path}")

    try:
        # ===== IMAGE CONVERSIONS =====
        if ext in [".jpg", ".jpeg", ".png"]:
            img = Image.open(temp_path)

            # Image → Image
            if target in ["png", "jpg", "jpeg"]:
                logger.info(f"Image {ext} → {target}")
                if target in ["jpg", "jpeg"] and img.mode in ('RGBA', 'LA', 'P'):
                    bg = Image.new('RGB', img.size, (255, 255, 255))
                    if img.mode == 'P':
                        img = img.convert('RGBA')
                    bg.paste(img, mask=img.split()[-1] if 'A' in img.mode else None)
                    img = bg
                out = io.BytesIO()
                fmt = "JPEG" if target in ["jpg", "jpeg"] else target.upper()
                img.save(out, format=fmt)
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.{target}")

            # Image → PDF
            elif target == "pdf":
                logger.info("Image → PDF")
                if img.mode in ('RGBA', 'LA', 'P'):
                    bg = Image.new('RGB', img.size, (255, 255, 255))
                    if img.mode == 'P':
                        img = img.convert('RGBA')
                    bg.paste(img, mask=img.split()[-1] if 'A' in img.mode else None)
                    img = bg
                out = io.BytesIO()
                img.save(out, format="PDF")
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.pdf")

            # Image → Word
            elif target in ["docx", "doc"]:
                logger.info("Image → Word")
                doc = Document()
                doc.add_heading(f'Image: {filename}', 0)
                try:
                    doc.add_picture(temp_path, width=Inches(6))
                except Exception as e:
                    doc.add_paragraph(f"(Image could not be embedded: {e})")
                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.{target}")

            # Image → PowerPoint
            elif target in ["pptx", "ppt"]:
                logger.info("Image → PowerPoint")
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_picture(temp_path, PPTInches(0.5), PPTInches(0.5), width=PPTInches(9))
                out = io.BytesIO()
                prs.save(out)
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.{target}")

        # ===== WORD CONVERSIONS =====
        elif ext in [".docx", ".doc"]:
            doc = Document(temp_path)

            # Word → Word
            if target in ["docx", "doc"]:
                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.{target}")

            # Word → PDF
            elif target == "pdf":
                out = io.BytesIO()
                c = canvas.Canvas(out, pagesize=letter)
                y = 750
                for para in doc.paragraphs:
                    if para.text.strip():
                        c.drawString(50, y, para.text[:100])
                        y -= 15
                        if y < 50:
                            c.showPage()
                            y = 750
                c.save()
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.pdf")

            # Word → PowerPoint
            elif target in ["pptx", "ppt"]:
                prs = Presentation()
                for para in doc.paragraphs:
                    if para.text.strip():
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        slide.shapes.title.text = para.text[:60]
                out = io.BytesIO()
                prs.save(out)
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.{target}")

        # ===== POWERPOINT CONVERSIONS =====
        elif ext in [".pptx", ".ppt"]:
            prs = Presentation(temp_path)

            # PowerPoint → PowerPoint
            if target in ["pptx", "ppt"]:
                out = io.BytesIO()
                prs.save(out)
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.{target}")

            # PowerPoint → PDF
            elif target == "pdf":
                out = io.BytesIO()
                c = canvas.Canvas(out, pagesize=letter)
                for i, slide in enumerate(prs.slides, 1):
                    c.drawString(50, 750, f"Slide {i}")
                    y = 730
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and shape.text.strip():
                            c.drawString(60, y, shape.text[:90])
                            y -= 20
                    c.showPage()
                c.save()
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.pdf")

            # PowerPoint → Word
            elif target in ["docx", "doc"]:
                doc = Document()
                doc.add_heading(f'Converted from {filename}', 0)
                for i, slide in enumerate(prs.slides, 1):
                    doc.add_heading(f'Slide {i}', level=1)
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and shape.text.strip():
                            doc.add_paragraph(shape.text)
                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                return send_file(out, as_attachment=True, download_name=f"{name}.{target}")

        # ===== PDF CONVERSIONS =====
        elif ext == ".pdf":
            # PDF → Word
            if target in ["docx", "doc"]:
                logger.info("PDF → Word")
                word_path = os.path.join(CONVERTED_FOLDER, f"{name}.docx")
                cv = Converter(temp_path)
                cv.convert(word_path, start=0, end=None)
                cv.close()
                return send_file(word_path, as_attachment=True, download_name=f"{name}.docx")

            # PDF → PowerPoint
            elif target in ["pptx", "ppt"]:
                logger.info("PDF → PowerPoint")
                slides = convert_from_path(temp_path)
                prs = Presentation()
                for slide_img in slides:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    temp_img = os.path.join(CONVERTED_FOLDER, "temp_slide.png")
                    slide_img.save(temp_img, 'PNG')
                    slide.shapes.add_picture(temp_img, PPTInches(0.5), PPTInches(0.5), width=PPTInches(9))
                    os.remove(temp_img)
                pptx_path = os.path.join(CONVERTED_FOLDER, f"{name}.pptx")
                prs.save(pptx_path)
                return send_file(pptx_path, as_attachment=True, download_name=f"{name}.pptx")

        # ===== UNSUPPORTED =====
        msg = f"Unsupported conversion: {ext} → {target}"
        logger.warning(msg)
        return msg, 400

    except Exception as e:
        logger.error(f"Conversion failed: {e}", exc_info=True)
        return f"Error: {str(e)}", 500

    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
                logger.info(f"Cleaned: {temp_path}")
            except Exception as e:
                logger.warning(f"Cleanup failed: {e}")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
