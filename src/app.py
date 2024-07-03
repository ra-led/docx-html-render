from flask import Flask, render_template, request, redirect
import os
import tempfile
from utils import docx_to_html

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
                temp_file = tempfile.NamedTemporaryFile(delete=False)
                file.save(temp_file.name)
                html, toc = docx_to_html(temp_file.name)
                os.unlink(temp_file.name)
                return render_template('result.html', html_content=html, toc_links=toc)
        return render_template('upload.html')

    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    return app
