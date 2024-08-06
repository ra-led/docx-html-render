import json
import os
import tempfile
import traceback
import docx
from flask import Flask, render_template, request, redirect, send_file
from doc_parse import DocHandler, DocHTML, DocJSON
from doc_parse.conf import CONF


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
                # Read DOC
                temp_file = tempfile.NamedTemporaryFile(delete=False)
                file.save(temp_file.name)
                
                doc = docx.Document(temp_file.name)
                handler = DocHandler(doc, **CONF)
                
                # Convert to HTML
                html_converter = DocHTML()
                html_content, toc_links = html_converter.get_html(handler)

                os.unlink(temp_file.name)

                # Convert to JSON
                try:
                    json_converter = DocJSON()
                    json_content = json_converter.get_json(handler)
                except Exception as ex:
                    tb = ''.join(
                        traceback.TracebackException.from_exception(ex).format()
                    )
                    json_content = json.dumps({'result': 'Failed', 'traceback': tb})
                json_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.json')
                with open(json_file_path, 'w') as json_file:
                    json_file.write(json_content)

                return render_template(
                    'result.html',
                    html_content=html_content,
                    toc_links=toc_links,
                    json_file_path=json_file_path
                )
        return render_template('upload.html')

    @app.route('/download_json')
    def download_json():
        json_file_path = request.args.get('path')
        return send_file(json_file_path, as_attachment=True)

    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    return app
