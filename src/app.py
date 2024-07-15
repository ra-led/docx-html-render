from flask import Flask, render_template, request, redirect, send_file
import os
import tempfile
from utils import docx_to_html
from html_to_json import html_to_json


def create_app():
    app = Flask(__name__)
    app.config['UPLOAD_FOLDER'] = 'uploads/'
    app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max file size

    @app.route('/', methods=['GET', 'POST'])
    def upload_file():
        if request.method == 'POST':
            if 'file' not in request.files:
                return redirect(request.url)
            file = request.files['file']
            if file.filename == '' or not file.filename.endswith(('.doc', '.docx')):
                return redirect(request.url)
            if file:
                # Convert DOC to HTML
                temp_file = tempfile.NamedTemporaryFile(delete=False)
                file.save(temp_file.name)
                html, toc = docx_to_html(temp_file.name)
                os.unlink(temp_file.name)

                # Convert HTML to JSON
                json_data = html_to_json(html)
                json_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.json')
                with open(json_file_path, 'w') as json_file:
                    json_file.write(json_data)

                return render_template('result.html', html_content=html, toc_links=toc, json_file_path=json_file_path)
        return render_template('upload.html')

    @app.route('/download_json')
    def download_json():
        json_file_path = request.args.get('path')
        return send_file(json_file_path, as_attachment=True)

    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    return app
