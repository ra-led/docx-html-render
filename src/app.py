from flask import Flask, render_template, request, redirect, url_for
import os
import tempfile
import docx
import re
import xmltodict
import uuid

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max file size

STYLE_TAGS = {
    'Title': 'h1',
    'List Paragraph': 'li',
    'Heading 1': 'h2',
    'Heading 2': 'h3',
    'Heading 3': 'h4',
    'Heading 4': 'h4',
    'Heading 5': 'h4',
    'Heading 6': 'h4',
    'Heading 7': 'h4'
}

ALIGNMENT = {
    'JUSTIFY': 'justify',
    'LEFT': 'left',
    'RIGHT': 'right',
    'CENTER': 'center'
}


def get_numbering_depth(text):
    numbering_pattern = r'^\d+\.'
    
    depth = 0
    while 1:
        if not re.findall(numbering_pattern, text.strip()):
            break
        depth += 1
        text = re.sub(numbering_pattern, '', text)
    
    # Last chance num without dot
    numbering_pattern = r'^\d+'
    if re.findall(numbering_pattern, text.strip()):
        depth += 1
    
    return depth


def get_xml_depth(par):
    depth = 0
    xml = xmltodict.parse(par._p.xml, process_namespaces=False)
    try:
        numeric = xml['w:p']['w:pPr']['w:numPr']['w:numId']['@w:val'] == '2'
        if numeric:
            level = xml['w:p']['w:pPr']['w:numPr']['w:ilvl']['@w:val']
            depth = int(level) + 1
    except KeyError:
        pass
    return depth


def get_heading_depth(par):
    depth = 0
    style = par.style.name
    if style:
        match = re.search(r'Heading (\d+)', style)
        if match:
            depth = int(match.group(1))
    return depth


def make_toc_header(text, depth, max_len=35):
    text = '__' * (depth - 1) + text
    if len(text) > max_len:
        text = text[:max_len] + '...'
    return text


def paragraph_style(par):
    css = ''
    try:
        css += 'text-align: {}'.format(ALIGNMENT[par.alignment.name])
    except (KeyError, AttributeError):
        pass
    if css:
        css = ' style="' + css + '"'
    return css


def docx_to_html(docx_path):
    doc = docx.Document(docx_path)
    html_content = []
    toc_links = []
    tables_cnt = 0
    last_depth = 0

    for content in doc.iter_inner_content():
        if type(content) is docx.text.paragraph.Paragraph:
            # if content.text == 'Приложение 3 (обязательное) – Проект договора подряда.':
            #     break
            style = content.style.name
            depth = get_numbering_depth(content.text) \
                or get_xml_depth(content) \
                or get_heading_depth(content)
            if depth:
                style = f'Heading {min(7, depth)}'
            try:
                tag = STYLE_TAGS[style]
            except KeyError:
                tag = 'p'
            css = paragraph_style(content)
            if tag.startswith('h'):  # Check if it's a heading
                anchor = str(uuid.uuid4())
                toc_links.append(f'<a href="#{anchor}">{make_toc_header(content.text, depth)}</a><br>')
                html_content.append(f'<div{css}><{tag} id="{anchor}">{content.text}</{tag}></div>')
            else:
                html_content.append(f'<div{css}><{tag}>')
                html_content.append(
                    ''.join([
                        '<span style="{bold}">{text}</span>'.format(
                            bold="font-weight: bold;",
                            text=run.text
                        ) if run.bold else run.text
                        for run in content.runs
                    ])
                )
                html_content.append(f'</{tag}></div>')
            last_depth = depth if depth else last_depth
        elif type(content) is docx.table.Table:
            tables_cnt += 1
            anchor = f'table{tables_cnt}'
            toc_links.append(f'<a href="#{anchor}">{make_toc_header("", last_depth + 1)}Таблица {tables_cnt}</a><br>')
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
    return html, ''.join(toc_links)


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
    app.run(host='0.0.0.0', port=5000, debug=True)
