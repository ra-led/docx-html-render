import json
import os
import tempfile
import traceback
import docx
from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from jinja2 import Template
from doc_parse import DocHandler, DocHTML, DocJSON
from doc_parse.conf import CONF

def create_app():
    app = FastAPI()
    upload_folder = 'uploads/'
    max_content_length = 16 * 1024 * 1024  # 16 MB max file size

    @app.get("/", response_class=HTMLResponse)
    async def upload_file_form():
        with open('templates/upload.html', 'r') as f:
            template = Template(f.read())
        return template.render()

    @app.post("/", response_class=HTMLResponse)
    async def upload_file(file: UploadFile = File(...)):
        if not file.filename.endswith(('.doc', '.docx')):
            return RedirectResponse(url="/", status_code=303)

        # Read DOC
        temp_file = tempfile.NamedTemporaryFile(delete=False)
        try:
            contents = await file.read()
            temp_file.write(contents)
            temp_file.flush()
            doc = docx.Document(temp_file.name)
            handler = DocHandler(doc, **CONF)

            # Convert to HTML
            html_converter = DocHTML()
            html_content, toc_links = html_converter.get_html(handler)

            # Convert to JSON
            try:
                json_converter = DocJSON()
                json_content = json_converter.get_json(handler)
            except Exception as ex:
                tb = ''.join(traceback.TracebackException.from_exception(ex).format())
                json_content = json.dumps({'result': 'Failed', 'traceback': tb})

            json_file_path = os.path.join(upload_folder, 'output.json')
            with open(json_file_path, 'w') as json_file:
                json_file.write(json_content)

            with open('templates/result.html', 'r') as f:
                template = Template(f.read())
            return template.render(
                html_content=html_content,
                toc_links=toc_links,
                json_file_path=json_file_path
            )
        finally:
            temp_file.close()
            os.unlink(temp_file.name)

    @app.get("/download_json")
    async def download_json(path: str):
        return FileResponse(path, media_type='application/json', filename='output.json')

    os.makedirs(upload_folder, exist_ok=True)
    return app

app = create_app()

