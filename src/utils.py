import docx
import re
import xmltodict
import uuid


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

DEFAULT_NUM_LEVELS = [
    {'@w:ilvl': '0', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': '%1'}},
    {'@w:ilvl': '1', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': '%1.%2'}},
    {'@w:ilvl': '2', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': '%1.%2.%3'}},
    {'@w:ilvl': '3', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': '%1.%2.%3.%4'}},
    {'@w:ilvl': '4', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': '%1.%2.%3.%4.%5'}},
    {'@w:ilvl': '5', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': '%1.%2.%3.%4.%5.%6'}},
    {'@w:ilvl': '6', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': '%1.%2.%3.%4.%5.%6.%7'}},
    {'@w:ilvl': '7', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': '%1.%2.%3.%4.%5.%6.%7.%8'}},
    {'@w:ilvl': '8', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': '%1.%2.%3.%4.%5.%6.%7.%8.%9'}}
]


class DocHandler:
    def __init__(self, doc):
        self.xml = xmltodict.parse(doc.element.xml, process_namespaces=False)
        self.num_xml = xmltodict.parse(
            doc.part.numbering_part.element.xml,
            process_namespaces=False
        )
        self.styles_xml = xmltodict.parse(
            doc.part.styles.element.xml,
            process_namespaces=False
        )
        self.styles = {
            x['@w:styleId']: x
            for x in self.styles_xml['w:styles']['w:style']
        }
        self.nums_abstarct = {
            x['w:abstractNumId']['@w:val']: x['@w:numId']
            for x in self.num_xml['w:numbering']['w:num']
        }
        self.nums_levels = {
            self.nums_abstarct[x['@w:abstractNumId']]: x['w:lvl']
            for x in self.num_xml['w:numbering']['w:abstractNum']
        }
        self.style_nums = {}
        self.style_levels = {}
        for numId, levels in self.nums_levels.items():
            for lvl in levels:
                if 'w:pStyle' in lvl:
                    self.style_nums[lvl['w:pStyle']['@w:val']] = numId
                    self.style_levels[lvl['w:pStyle']['@w:val']] = lvl
        self.increment = {
            x['@w:numId']: {
                i: 0
                for i in range(len(self.nums_levels[x['@w:numId']]))
            }
            for x in self.num_xml['w:numbering']['w:num']
        }
        self.depth = 0
        self.source = None
        self.depth_anchor = {}
        self.tables_cnt = 0
        self.width = int(self.xml['w:document']['w:body']['w:sectPr']['w:pgSz']['@w:w'])
        self.last_par = None
        self.max_frame_space = 7
        
    def numrize_by_meta(self, par):
        p_xml = xmltodict.parse(par._p.xml, process_namespaces=False)
        try:
            num_id = p_xml['w:p']['w:pPr']['w:numPr']['w:numId']['@w:val']
        except:
            return '', 0, None
        level = int(p_xml['w:p']['w:pPr']['w:numPr']['w:ilvl']['@w:val'])
        if num_id not in self.nums_levels:
            self.nums_levels[num_id] = DEFAULT_NUM_LEVELS
            self.increment[num_id] = {lvl_i: 0 for lvl_i in range(len(DEFAULT_NUM_LEVELS))}
        # Update inc
        self.increment[num_id][level] += 1
        for lvl_i in self.increment[num_id]:
            if lvl_i > level:
                self.increment[num_id][lvl_i] = 0
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
        
        main_style = []
        for lvl in range(1, level + 2):
            main_style.append(f'%{lvl}')
        main_style = '.'.join(main_style)
        if abstarct_levels[level]['w:lvlText']['@w:val'] != main_style:
            num_id = 'sub'
        return num_prefix, level + 1, num_id
    
    def numrize_by_style_id(self, par):
        style_id = par.style.style_id
        base_id = par.style.base_style.style_id if par.style.base_style else None
        if style_id in self.style_levels:
            level = int(self.style_levels[style_id]['@w:ilvl'])
            num_id = self.style_nums[style_id]
        elif base_id in self.style_levels:
            level = int(self.style_levels[base_id]['@w:ilvl'])
            num_id = self.style_nums[base_id]
        else:
            return '', 0, None
        # Update inc
        self.increment[num_id][level] += 1
        for lvl_i in self.increment[num_id]:
            if lvl_i > level:
                self.increment[num_id][lvl_i] = 0
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
        
        main_style = []
        for lvl in range(1, level + 2):
            main_style.append(f'%{lvl}')
        main_style = '.'.join(main_style)
        if abstarct_levels[level]['w:lvlText']['@w:val'] != main_style:
            num_id = 'sub'
        return num_prefix, level + 1, num_id
    
    def numerize_by_text(self, par):
        depth = 0
        text = par.text
        num_prefix = ''
        numbering_pattern = r'^\d+\.'
        while 1:
            match = re.findall(numbering_pattern, text.strip())
            if not match:
                break
            depth += 1
            text = re.sub(numbering_pattern, '', text)
            num_prefix += match[0]
        
        # Last chance num without dot
        numbering_pattern = r'^\d+\s'
        match = re.findall(numbering_pattern, text.strip())
        if match:
            depth += 1
            num_prefix += match[0]
        return num_prefix, depth, 'N'
    
    def numerize_by_style(self, par):
        depth = 0
        style = par.style.name
        if style:
            match = re.search(r'Heading (\d+)', style)
            if match:
                depth = int(match.group(1))
            elif style == 'Title':
                depth = 1
        return par.text, depth, 'H'
    
    def numerize(self, par):
        numerize_prioritet = [
            self.numrize_by_meta,
            self.numrize_by_style_id,
            self.numerize_by_text,
            self.numerize_by_style
        ]
        for method in numerize_prioritet:
            num_prefix, depth, source = method(par)
            if num_prefix:
                return num_prefix, depth, source
        return '', 0, None
    
    def get_depth_classes(self):
        aa = []
        for k, v in self.depth_anchor.items():
            if k <= self.depth:
                aa.append(v)
        return " ".join(aa)
    
    def process_paragraph(self, par):
        html_paragraph = []
        html_links = []
        style = par.style.name
        num_prefix, depth, source = self.numerize(par)
        
        if depth:
            anchor = 'a' + str(uuid.uuid4())
            if source != 'sub':
                # anchor = str(uuid.uuid4())
                self.depth = depth
                self.source = source
                self.depth_anchor[depth] = anchor
                style = f'Heading {min(7, depth)}'
            else:
                style = 'List Paragraph'
                # depth = self.depth + depth
                # style = f'Heading {min(7, depth)}'
        try:
            tag = STYLE_TAGS[style]
        except KeyError:
            tag = 'p'
        css = paragraph_style(par)
        classes = self.get_depth_classes()
        if tag.startswith('h'):  # Check if it's a heading
            # text = num_prefix + f' [{source}] ' + par.text
            if source not in ('H', 'N'):
                text = num_prefix + par.text
            else:
                text = par.text
            html_links.append((f'<a href="#{anchor}">{make_toc_header(text, depth)}</a><br>', source))
            html_paragraph.append(f'<div{css} class="{classes}"><{tag} id="{anchor}">{text}</{tag}></div>')
        else:
            html_paragraph.append(f'<div{css} class="{classes}"><{tag}>')
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
        return html_paragraph, html_links
    
    def investigate_table(self, table):
        # t_xml = xmltodict.parse(table._element.xml, process_namespaces=False)
        merged = set()
        all_text_cells = []
        for i, row in enumerate(table.rows):
            row_cols = 0
            for j, cell in enumerate(row.cells):
                if cell._element in merged:
                    continue
                c_xml = xmltodict.parse(cell._element.xml, process_namespaces=False)
                # Cell width
                cell_width = int(c_xml['w:tc']['w:tcPr']['w:tcW']['@w:w'])
                text_cell = (cell_width / self.width) >= 0.8
                # detect merged cells
                rowspan = 1
                colspan = 1
                if text_cell:
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
                if rowspan > 1 or colspan > 1:
                    merged.add(cell._element)
                row_cols += 1
                if text_cell:
                    all_text_cells.append({
                        'row_top_space': i,
                        'col_left_space': j,
                        'row_bottom_space': i + rowspan,
                        'col_right_space': j + colspan
                    })
        # Find col space
        try:
            left_space = max(d['col_left_space'] for d in all_text_cells)
            frequent_right_space = max(
                set(d['col_right_space']for d in all_text_cells),
                key=lambda x: list(d['col_right_space'] for d in all_text_cells).count(x)
            )
            right_space = len(row.cells) - frequent_right_space
        except ValueError:
            return None
        
        if len(table.columns) < self.max_frame_space:
            return None
        
        bottom_space = max(d['row_bottom_space'] for d in all_text_cells)
        top_space = min(d['row_top_space'] for d in all_text_cells)
        if left_space > self.max_frame_space:
            left_space = 0
        if (len(table.columns) - right_space) > self.max_frame_space:
            right_space = len(table.columns)
        if top_space > self.max_frame_space:
            top_space = 0
        if (len(table.rows) - bottom_space) > self.max_frame_space:
            bottom_space = len(table.rows)
        text_rows = [list(range(x['row_top_space'], x['row_bottom_space'])) for x in all_text_cells]
        return left_space, right_space, top_space, bottom_space, sum(text_rows, [])

    def process_table(self, table):
        self.tables_cnt += 1
        frame = self.investigate_table(table)
        if frame:
            left_space, right_space, top_space, bottom_space, text_rows = frame
            print(text_rows)
        else:
            text_rows = []
        anchor = f'table{self.tables_cnt}'
        table_links = [
            # (
            #     f'<a href="#{anchor}">'
            #     f'{make_toc_header("", self.depth + 1)}Таблица {self.tables_cnt}'
            #     '</a><br>', 'T'
            # )
        ]
        t_xml = xmltodict.parse(table._element.xml, process_namespaces=False)
        try:
            default_borders = {
                'w:bottom': t_xml['w:tbl']['w:tblPr']['w:tblBorders']['w:insideH'],
                'w:right': t_xml['w:tbl']['w:tblPr']['w:tblBorders']['w:insideV'],
                'w:left': t_xml['w:tbl']['w:tblPr']['w:tblBorders']['w:insideV'],
                'w:top': t_xml['w:tbl']['w:tblPr']['w:tblBorders']['w:insideH']
            }
        except KeyError:
            # table without thresholds - most likley TEXT
            default_borders = {}
        merged = set()
        # html_table = f'<table id="{anchor}"class="w3-table w3-hoverable">'
        html_table = f'<table "class="w3-table w3-hoverable">'
        for i, row in enumerate(table.rows):
            html_table += "<tr>"
            for j, cell in enumerate(row.cells):
                ignore = False
                if frame:
                    if i < top_space or ignore:
                        ignore = True
                    if i >= bottom_space or ignore:
                        ignore = True
                    if j < left_space or ignore:
                        ignore = True
                    if j >= right_space or ignore:
                        ignore = True
                if cell._element in merged:
                    continue
                c_xml = xmltodict.parse(cell._element.xml, process_namespaces=False)
                # Cell width
                cell_width = int(c_xml['w:tc']['w:tcPr']['w:tcW']['@w:w'])
                text_cell = (cell_width / self.width) >= 0.8
                # Cell is text
                if text_cell:
                    text = ''
                    for c_par in cell.paragraphs:
                        html_paragraph, html_links = self.process_paragraph(c_par)
                        table_links += html_links
                        text += ''.join(html_paragraph)
                else:
                    text = cell.text
                css = cell_style(cell, default_borders.copy(), c_xml)
                # detect merged cells
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
                classes = ""
                if not text_cell:
                    classes = self.get_depth_classes()
                if ignore:
                    continue
                    # text = 'IGNORE'
                    # classes = 'ignore'
                if i in text_rows and text_cell:
                    print(html_table[-40:])
                    print(text[:40], '...', text[:-40])
                    html_table += '</tr></table>' + text + '<table "class="w3-table w3-hoverable">'
                    if rowspan > 1 or colspan > 1:
                        merged.add(cell._element)
                    continue
                else:
                    if rowspan > 1 or colspan > 1:
                        html_table += f'<td class="{classes}" rowspan="{rowspan}" colspan="{colspan}"{css}>{text}</td>'
                        merged.add(cell._element)
                    else:
                        html_table += f'<td{css}>{text}</td>'
            html_table += "</tr>"
        html_table += "</table>"
        return html_table, table_links
        
    
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


def docx_to_html(docx_path):
    doc = docx.Document(docx_path)
    handler = DocHandler(doc)
    html_content = []
    toc_links = []
    
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
    return ''.join(html_content), ''.join([link for link, src in toc_links])