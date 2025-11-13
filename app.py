# from flask import Flask, request, send_file, render_template
# from werkzeug.utils import secure_filename
# from PIL import Image
# import os

# app = Flask(__name__)

# UPLOAD_FOLDER = "uploads"
# CONVERTED_FOLDER = "converted"
# os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# os.makedirs(CONVERTED_FOLDER, exist_ok=True)

# @app.route("/")
# def home():
#     return render_template("index.html")  # Your frontend file

# @app.route("/convert", methods=["POST", "OPTIONS"], strict_slashes=False)
# def convert_file():
#     # Allow CORS preflight or other OPTIONS requests to succeed.
#     if request.method == 'OPTIONS':
#         return ('', 200)

#     if "file" not in request.files:
#         return "No file uploaded", 400

#     file = request.files["file"]
#     target = request.form.get("target")
#     # Log incoming request method and target for easier debugging
#     print(f"/convert called with method={request.method}, target={target}")

#     if file.filename == "":
#         return "No selected file", 400

#     filename = secure_filename(file.filename)
#     input_path = os.path.join(UPLOAD_FOLDER, filename)
#     file.save(input_path)

#     name, ext = os.path.splitext(filename)
#     output_filename = f"{name}_converted.{target}"
#     output_path = os.path.join(CONVERTED_FOLDER, output_filename)

#     # Basic conversion: image to another image format
#     try:
#         if target in ["png", "jpg", "jpeg"]:
#             img = Image.open(input_path)
#             img.save(output_path)
#         elif target == "pdf":
#             img = Image.open(input_path)
#             img.save(output_path, "PDF")
#         else:
#             return "Unsupported conversion type", 400

#     except Exception as e:
#         return str(e), 500

#     return send_file(output_path, as_attachment=True)

# if __name__ == "__main__":
#     app.run(debug=True)

from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
from PIL import Image
import io

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")  # Your frontend file

@app.route("/convert", methods=["POST", "OPTIONS"], strict_slashes=False)
def convert_file():
    if request.method == 'OPTIONS':
        return ('', 200)

    if "file" not in request.files:
        return "No file uploaded", 400

    file = request.files["file"]
    target = request.form.get("target")

    if file.filename == "":
        return "No selected file", 400

    filename = secure_filename(file.filename)
    name, ext = filename.rsplit(".", 1)

    try:
        img = Image.open(file.stream)

        img_io = io.BytesIO()
        if target in ["png", "jpg", "jpeg"]:
            img.save(img_io, format=target.upper())
        elif target == "pdf":
            img.save(img_io, format="PDF")
        else:
            return "Unsupported conversion type", 400

        img_io.seek(0)
        output_filename = f"{name}_converted.{target}"
        return send_file(img_io, as_attachment=True, download_name=output_filename)

    except Exception as e:
        return str(e), 500

# if __name__ == "__main__":
#     app.run(debug=True)

# if __name__ == "__main__":
#     import os
#     port = int(os.environ.get("PORT", 5000))
#     app.run(host="0.0.0.0", port=port)
import os

port = int(os.environ.get("PORT", 5000))
app.run(host="0.0.0.0", port=port)
