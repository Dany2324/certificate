from flask import Flask, request, render_template, send_file
import os
from docx import Document
from docx2pdf import convert
import pythoncom
import time
import uuid
from docx.shared import Pt, RGBColor

app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        name = request.form['name']
        document_path = "C:/Users/Abdelrahman Fathi/Downloads/ITP.docx"
        output_path = f"C:/Users/Abdelrahman Fathi/Downloads/gg/{name}_modified.docx"

        if os.path.exists(document_path):
            document = Document(document_path)
            for paragraph in document.paragraphs:
                if "RAMY MAHMOUD MOHAMED" in paragraph.text:
                    new_name_parts = name.split(" ")[:3]
                    new_name_parts = [part.upper() for part in new_name_parts]
                    new_name = " ".join(new_name_parts)
                    paragraph.text = paragraph.text.replace("RAMY MAHMOUD MOHAMED", new_name)

                    # Set the font, size, color, and bold attribute for the paragraph
                    run = paragraph.runs[0]
                    run.font.name = '29LT Kaff Semi Bold'
                    run.font.size = Pt(16)
                    run.font.color.rgb = RGBColor(0x11, 0xBA, 0xB1)
                    run.font.bold = True

            document.save(output_path)

            pythoncom.CoInitialize()  # Initialize COM before using docx2pdf
            convert(output_path)  # Convert the modified DOCX to PDF using docx2pdf
            pythoncom.CoUninitialize()  # Uninitialize COM

            timestamp = int(time.time())
            unique_filename = f"{timestamp}_{str(uuid.uuid4())[:8]}_{name}_modified.pdf"
            unique_file_path = f"C:/Users/Abdelrahman Fathi/Downloads/gg/{unique_filename}"

            os.rename(f"C:/Users/Abdelrahman Fathi/Downloads/gg/{name}_modified.pdf", unique_file_path)

            download_link = f"/download/{unique_filename}"
            return render_template('download.html', download_link=download_link, file_name=f"{name}_modified.pdf")
        else:
            return 'Document not found.'

    return render_template('index.html')


@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    file_path = f"C:/Users/Abdelrahman Fathi/Downloads/gg/{filename}"
    return send_file(file_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
