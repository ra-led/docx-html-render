from flask import Flask, render_template, request, redirect, url_for
import os
import tempfile
import docx

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max file size

STYLE_TAGS = {
    'Title': 'h1',
    'Body Text': 'p',
    'List Paragraph': 'li',
    'Heading 1': 'h2',
    'Heading 2': 'h3',
    'Heading 3': 'h4',
}


def docx_to_html(docx_path):
    doc = docx.Document(docx_path)
    html_content = []
    toc_links = []
    tables_cnt = 0

    for content in doc.iter_inner_content():
        if type(content) is docx.text.paragraph.Paragraph:
            try:
                tag = STYLE_TAGS[content.style.name]
            except KeyError:
                tag = 'p'
            if tag.startswith('h'):  # Check if it's a heading
                anchor = content.text.replace(' ', '_').lower()
                toc_links.append(f'<div class="link"><a href="#{anchor}">{content.text}</a></div><br>')
                html_content.append(f'<{tag} id="{anchor}">{content.text}</{tag}>')
            else:
                html_content.append(f'<{tag}>')
                html_content.append(
                    ''.join([
                        '<span style="font-size:{};">{}</span>'.format(14, run.text)
                        for run in content.runs
                    ])
                )
                html_content.append(f'</{tag}>')
        elif type(content) is docx.table.Table:
            tables_cnt += 1
            anchor = f'table{tables_cnt}'
            toc_links.append(f'<div class="link"><a href="#{anchor}">Таблица {tables_cnt}</a></div><br>')
            html_content.append(f'<table id="{anchor}" class="w3-table-all w3-hoverable">')
            for i, row in enumerate(content.rows):
                html_content.append('<tr>')
                for cell in row.cells:
                    html_content.append('<td>' if i > 0 else '<th>')
                    html_content.append(cell.text)
                    html_content.append('</td>' if i > 0 else '</th>')
                html_content.append('</tr>')
            html_content.append('</table>')
        else:
            print(type(content), 'missed')

    html = ''.join(html_content)
    toc=''.join(toc_links)
    return html, toc


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


if __name__ == '__main__':
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run(debug=True)
