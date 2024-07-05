from flask import Flask, render_template, request, redirect, session
import os
import tempfile
from utils import DocHandler
import docx
import threading


def update_progress(progress):
    session['progress'] = progress


def docx_to_html(docx_path, update_progress):
    doc = docx.Document(docx_path)
    handler = DocHandler(doc)
    html_content = []
    toc_links = []
    total_content = len(list(doc.iter_inner_content()))
    processed_content = 0
    
    for content in doc.iter_inner_content():
        if type(content) is docx.text.paragraph.Paragraph:
            html_paragraph, html_links = handler.process_paragraph(content)
            html_content.extend(html_paragraph)
            toc_links.extend(html_links)
        elif type(content) is docx.table.Table:
            html_table, table_links = handler.process_table(content)
            html_content.append(html_table)
            toc_links.extend(table_links)
        else:
            print(type(content), 'missed')
        
        processed_content += 1
        progress = int((processed_content / total_content) * 100)
        update_progress(progress)
    
    return '\n'.join(html_content), '\n'.join([link for link, src in toc_links])


def create_app():
    app = Flask(__name__)
    app.config['UPLOAD_FOLDER'] = 'uploads/'
    app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max file size
    app.secret_key = 'your_secret_key'

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

    @app.route('/progress')
    def progress():
        return str(session.get('progress', 0))

    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    return app
