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


class DocNumeration:
    def __init__(self, doc):
        self.xml = xmltodict.parse(
            doc.part.numbering_part.element.xml,
            process_namespaces=False
        )
        self.nums_abstarct = {
            x['w:abstractNumId']['@w:val']: x['@w:numId']
            for x in self.xml['w:numbering']['w:num']
        }
        self.nums_levels = {
            self.nums_abstarct[x['@w:abstractNumId']]: x['w:lvl']
            for x in self.xml['w:numbering']['w:abstractNum']
        }
        self.increment = {
            x['@w:numId'] : {
                i: 0
                for i in range(len(self.nums_levels[x['@w:numId']]))
            }
            for x in self.xml['w:numbering']['w:num']
        }
        
    def numrize(self, par):
        p_xml = xmltodict.parse(par._p.xml, process_namespaces=False)
        try:
            num_id = p_xml['w:p']['w:pPr']['w:numPr']['w:numId']['@w:val']
        except:
            return ''
        level = int(p_xml['w:p']['w:pPr']['w:numPr']['w:ilvl']['@w:val'])
        # Update inc
        self.increment[num_id][level] += 1
        for lvl_next in range(level + 1, len(self.increment[num_id])):
            self.increment[num_id][lvl_next] = 0
        abstarct_levels = self.nums_levels[num_id]
        num_prefix = ''
        for lvl_a, lvl_i in zip(abstarct_levels, self.increment[num_id]):
            if lvl_i > level:
                break
            try:
                num_start = int(lvl_a['w:start']['@w:val'])
            except KeyError:
                num_start = 1
            num = self.increment[num_id][lvl_i] + num_start - 1
            num = max(num, num_start)
            num_prefix += f'{num}.'
        
        return num_prefix


def get_numbering_depth(par):
    text = par.text
    numbering_pattern = r'^\d\.'
    
    depth = 0
    while 1:
        if not re.findall(numbering_pattern, text.strip()):
            break
        depth += 1
        text = re.sub(numbering_pattern, '', text)
    
    # Last chance num without dot
    numbering_pattern = r'^\d'
    if re.findall(numbering_pattern, text.strip()):
        depth += 1
    
    return depth, 'N'


def get_xml_depth(par):
    depth = 0
    xml = xmltodict.parse(par._p.xml, process_namespaces=False)
    try:
        numeric = int(xml['w:p']['w:pPr']['w:numPr']['w:numId']['@w:val'])
        level = xml['w:p']['w:pPr']['w:numPr']['w:ilvl']['@w:val']
        depth = int(level) + 1
    except KeyError:
        numeric = None
    return depth, numeric


def get_heading_depth(par):
    depth = 0
    style = par.style.name
    if style:
        match = re.search(r'Heading (\d+)', style)
        if match:
            depth = int(match.group(1))
    return depth, 'H'


def get_max_depth(par):
    depths = [get_numbering_depth(par), get_xml_depth(par), get_heading_depth(par)]
    max_depth = max(depths, key=lambda x: x[0])
    return max_depth


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


def cell_style(cell, borders, c_xml):
    try:
        borders.update(c_xml['w:tc']['w:tcPr']['w:tcBorders'])
    except KeyError:
        pass
    css = ''
    for k, v in borders.items():
        side = k.replace('w:', '')
        try:
            width = int(float(v['@w:sz']) / 4)
            color = v['@w:color']
        except KeyError:
            width = 0
            color = 'fff'
        css += f'border-{side}: {width}px solid #{color};'
    return ' style="' + css + '"'


def doc_table_to_html(table, anchor, page_w, last_depth, doc_num):
    t_xml = xmltodict.parse(table._element.xml, process_namespaces=False)
    try:
        default_borders = {
            'w:bottom': t_xml['w:tbl']['w:tblPr']['w:tblBorders']['w:insideH'],
            'w:right': t_xml['w:tbl']['w:tblPr']['w:tblBorders']['w:insideV'],
            'w:left': t_xml['w:tbl']['w:tblPr']['w:tblBorders']['w:insideV'],
            'w:top': t_xml['w:tbl']['w:tblPr']['w:tblBorders']['w:insideH']
        }
    except:
        # table without thresholds - most likley TEXT
        default_borders = {}
    table_links = []
    merged = set()
    html_table = f'<table id="{anchor}" class="w3-table w3-hoverable">'
    for i, row in enumerate(table.rows):
        html_table += "<tr>"
        for j, cell in enumerate(row.cells):
            if cell._element in merged:
                continue
            c_xml = xmltodict.parse(cell._element.xml, process_namespaces=False)
            # page width
            cell_w = int(c_xml['w:tc']['w:tcPr']['w:tcW']['@w:w'])
            # Cell is text
            if (cell_w / page_w) >= 0.8:
                text = ''
                for c_par in cell.paragraphs:
                    html_paragraph, depth, last_depth, html_links = process_paragraph(c_par, last_depth, doc_num)
                    table_links += html_links
                    text += ''.join(html_paragraph)
            else:
                text = cell.text
            css = cell_style(cell, default_borders.copy(), c_xml)
            rowspan = 1
            colspan = 1
            for next_row in table.rows[i+1:]:
                if j < len(next_row.cells) and next_row.cells[j]._element == cell._element:
                    rowspan += 1
                else:
                    break
            for next_cell in row.cells[j+1:]:
                if next_cell._element == cell._element:
                    colspan += 1
                else:
                    break
            if rowspan > 1 or colspan> 1:
                html_table += f'<td rowspan="{rowspan}" colspan="{colspan}"{css}>{text}</td>'
                merged.add(cell._element)
            else:
                html_table += f'<td{css}>{text}</td>'
        html_table += "</tr>"
    html_table += "</table>"
    return html_table, table_links, last_depth


def process_paragraph(par, last_depth, doc_num):
    html_paragraph = []
    html_links = []
    style = par.style.name
    depth, source = get_max_depth(par)
    if depth:
        style = f'Heading {min(7, depth)}'
    try:
        tag = STYLE_TAGS[style]
    except KeyError:
        tag = 'p'
    css = paragraph_style(par)
    if tag.startswith('h'):  # Check if it's a heading
        text = doc_num.numrize(par) + par.text
        anchor = str(uuid.uuid4())
        html_links.append((f'<a href="#{anchor}">{make_toc_header(text, depth)}</a><br>', source))
        html_paragraph.append(f'<div{css}><{tag} id="{anchor}">{text}</{tag}></div>')
    else:
        html_paragraph.append(f'<div{css}><{tag}>')
        html_paragraph.append(
            ''.join([
                '<span style="{bold}">{text}</span>'.format(
                    bold="font-weight: bold;",
                    text=run.text
                ) if run.bold else run.text
                for run in par.runs
            ])
        )
        html_paragraph.append(f'</{tag}></div>')
    last_depth = depth if depth else last_depth
    return html_paragraph, depth, last_depth, html_links


def docx_to_html(docx_path):
    doc = docx.Document(docx_path)
    doc_num = DocNumeration(doc)
    d_xml = xmltodict.parse(doc.element.xml, process_namespaces=False)
    # page width
    page_w = int(d_xml['w:document']['w:body']['w:sectPr']['w:pgSz']['@w:w'])
    html_content = []
    toc_links = []
    tables_cnt = 0
    last_depth = 0

    for content in doc.iter_inner_content():
        if type(content) is docx.text.paragraph.Paragraph:
            html_paragraph, depth, last_depth, html_links = process_paragraph(content, last_depth, doc_num)
            html_content.extend(html_paragraph)
            toc_links.extend(html_links)
        elif type(content) is docx.table.Table:
            tables_cnt += 1
            anchor = f'table{tables_cnt}'
            toc_links.append((f'<a href="#{anchor}">{make_toc_header("", last_depth + 1)}Таблица {tables_cnt}</a><br>', 'T'))
            table, table_links, last_depth = doc_table_to_html(content, anchor, page_w, last_depth, doc_num)
            html_content.append(table)
            toc_links.extend(table_links)
        else:
            print(type(content), 'missed')

    html = ''.join(html_content)
    return html, ''.join([link for link, src in toc_links])


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
