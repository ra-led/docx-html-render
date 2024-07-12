import docx
import re
import xmltodict
import uuid
import html


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

DEFAULT_LEVELS = [
    {'@w:ilvl': '0', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default - %1'}},
    {'@w:ilvl': '1', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default - %1.%2'}},
    {'@w:ilvl': '2', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default - %1.%2.%3'}},
    {'@w:ilvl': '3', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default - %1.%2.%3.%4'}},
    {'@w:ilvl': '4', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default - %1.%2.%3.%4.%5'}},
    {'@w:ilvl': '5', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default - %1.%2.%3.%4.%5.%6'}},
    {'@w:ilvl': '6', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default - %1.%2.%3.%4.%5.%6.%7'}},
    {'@w:ilvl': '7', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default - %1.%2.%3.%4.%5.%6.%7.%8'}},
    {'@w:ilvl': '8', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default - %1.%2.%3.%4.%5.%6.%7.%8.%9'}}
]


class NumberingDB:
    def __init__(self, doc):
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
        self.levels = {
            x['@w:abstractNumId']: x['w:lvl']
            for x in self.num_xml['w:numbering']['w:abstractNum']
        }
        self.nums_to_abstarct = {
            x['@w:numId']: x['w:abstractNumId']['@w:val']
            for x in self.num_xml['w:numbering']['w:num']
        }
        self.style_to_abstract = {}
        for absId, lvls in self.levels.items():
            for lvl in lvls:
                if 'w:pStyle' in lvl:
                    self.style_to_abstract[lvl['w:pStyle']['@w:val']] = {
                        'absId': absId,
                        'lvl': int(lvl['@w:ilvl'])
                    }
        self.increment = {
            k: {i: 0 for i in range(len(v))}
            for k, v in self.levels.items()
        }
        
    def get_abs_id(self, numId=None, styleId=None):
        if numId:
            try:
                return self.nums_to_abstarct[numId]
            except KeyError:
                pass
        if styleId:
            try:
                return self.style_to_abstract[styleId]

            except KeyError:
                pass
        if numId:
            absId = str(uuid.uuid4())
            self.levels[absId] = DEFAULT_LEVELS
            self.increment[absId] = {i: 0 for i in range(len(DEFAULT_LEVELS))}
            self.nums_to_abstarct[numId] = absId
            return absId
        return None
    
    def check_heading_style(self, par):
        if re.findall('^таблица', par.text.strip().lower()):
            return False
        if re.findall('^рисунок', par.text.strip().lower()):
            return False
        bold = any([par.style.font.bold] + [run.bold for run in par.runs])
        large_font = par.style.font.size.pt > 12 if par.style.font.size else None
        if bold or large_font:
            return True
        else:
            return False
    
    def count_builtin(self, absId, level, par):
        # find outlined numaration levels
        # try:
        #     out_lvl = [self.styles[par.style.style_id]['w:pPr']['w:outlineLvl']['@w:val']]
        # except KeyError:
        #     out_lvl = []
        # Re-init lower leves
        self.increment[absId][level] += 1
        for lvl_i in self.increment[absId]:
            if lvl_i > level:
                self.increment[absId][lvl_i] = 0
        # Get levels
        abstarct_levels = self.levels[absId]
        # Generate num prefix
        depth = 0
        num_prefix = abstarct_levels[level]['w:lvlText']['@w:val']
        for lvl_a, lvl_i in zip(abstarct_levels, self.increment[absId]):
            if lvl_i > level:
                break
            try:
                num_start = int(lvl_a['w:start']['@w:val'])
            except KeyError:
                num_start = 1
            num = self.increment[absId][lvl_i] + num_start - 1
            num = max(num, num_start)
            if f'%{lvl_i + 1}' in num_prefix:
                depth += 1
                num_prefix = re.sub(f'%{lvl_i + 1}', str(num), num_prefix)
        return num_prefix, depth, absId
    
    def numrize_by_meta(self, par):
        p_xml = xmltodict.parse(par._p.xml, process_namespaces=False)
        try:
            numId = p_xml['w:p']['w:pPr']['w:numPr']['w:numId']['@w:val']
        except:
            return '', 0, None
        level = int(p_xml['w:p']['w:pPr']['w:numPr']['w:ilvl']['@w:val'])
        absId = self.get_abs_id(numId=numId)
        num_prefix, depth, source = self.count_builtin(absId, level, par)
        if self.check_heading_style(par) or depth > 1:
            return num_prefix, depth, source
        else:
            return '', 0, None
    
    def numrize_by_style(self, par):
        style_abs = self.get_abs_id(styleId=par.style.style_id)
        if style_abs is None:
            base_style_id = par.style.base_style.style_id if par.style.base_style else None
            style_abs = self.get_abs_id(styleId=base_style_id)
        if style_abs is None:
            return '', 0, None
        absId, level = style_abs['absId'], style_abs['lvl']
        return self.count_builtin(absId, level, par)
    
    def numerize_by_text(self, par):
        depth = 0
        text = par.text.strip()
        # Numbers with dots at the begin of text
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
        if self.check_heading_style(par) or depth > 1:
            return num_prefix, depth, 'REGEX'
        else:
            return '', 0, None
    
    def numerize_by_heading(self, par):
        depth = 0
        style = par.style.name
        if style:
            match = re.search(r'Heading (\d+)', style)
            if match:
                depth = int(match.group(1))
            elif style == 'Title':
                depth = 1
        if self.check_heading_style(par):
            return par.text if par.text.strip() else '[UNNAMED]', depth, 'HEADING'
        else:
            return '', 0, None
    
    def numerize(self, par):
        numerize_prioritet = [
            self.numrize_by_meta,
            self.numrize_by_style,
            self.numerize_by_text,
            self.numerize_by_heading
        ]
        for method in numerize_prioritet:
            num_prefix, depth, source = method(par)
            if num_prefix:
                return num_prefix, depth, source
        return '', 0, None
        

class DocHandler:
    def __init__(self, doc):
        self.xml = xmltodict.parse(doc.element.xml, process_namespaces=False)
        self.num_db = NumberingDB(doc)
        self.depth = 0
        self.source = None
        self.depth_anchor = {}
        self.tables_cnt = 0
        self.width = int(self.xml['w:document']['w:body']['w:sectPr']['w:pgSz']['@w:w'])
        self.height = int(self.xml['w:document']['w:body']['w:sectPr']['w:pgSz']['@w:h'])
        self.max_frame_space = 7
        self.last_pars = []
    
    def get_depth_classes(self):
        aa = []
        for k, v in self.depth_anchor.items():
            if k <= self.depth:
                aa.append(v)
        return " ".join(aa)
    
    def get_table_title(self):
        regex_title = ' '.join(self.last_pars)
        if "таблица" in regex_title.lower():
            strat_idx = re.search("таблица", regex_title.lower()).start()
            title = regex_title[strat_idx:]
        else:
            title = 'Таблица'
            # try:
            #     title = self.last_pars[-1]
            # except IndexError:
            #     title = ''
        title = html.escape(title if title.strip() else 'Таблица')
        anchor = f'table{self.tables_cnt}'
        return title, anchor
    
    def process_paragraph(self, par):
        html_paragraph = []
        html_links = []
        num_prefix, depth, source = self.num_db.numerize(par)
        
        styled_text = ''.join([
            '<span style="{bold}">{text}</span>'.format(
                bold="font-weight: bold;",
                text=html.escape(run.text)
            ) if run.bold else run.text
            for run in par.runs
        ])
        par_css = paragraph_style(par)
        tag = 'p'
        
        if depth:
            anchor = 'a' + str(uuid.uuid4())
            self.depth = depth
            self.source = source
            self.depth_anchor[depth] = anchor
        
            if source not in ('HEADING', 'REGEX'):
                styled_text = f'<span>{num_prefix} </span>' + styled_text
                text = num_prefix + ' ' + html.escape(par.text) #f' [{source}] ' + 
            else:
                text = html.escape(par.text)
            classes = self.get_depth_classes()
            html_links.append(f'<a href="#{anchor}">{make_toc_header(text, depth)}</a><br>')
            html_paragraph.append(f'<div {par_css} class="{classes}"><{tag} id="{anchor}">{styled_text}</{tag}></div>')
        else:
            classes = self.get_depth_classes()
            text = html.escape(par.text)
            html_paragraph.append(f'<div {par_css} class="{classes}"><{tag}>{styled_text}</{tag}></div>')
        # last to paragraphs for table title
        if text.strip():
            self.last_pars.append(text)
            self.last_pars = self.last_pars[-2:]
        return html_paragraph, html_links
    
    def investigate_table(self, table):
        t_xml = xmltodict.parse(table._element.xml, process_namespaces=False)
        try:
            table_height = sum([
                int(row['w:trPr']['w:trHeight']['@w:val'])
                for row in t_xml['w:tbl']['w:tr']
            ])
        except (KeyError, TypeError):
            table_height = 0
        if (table_height / self.height) < 0.8:
            return None
        merged = set()
        all_text_cells = []
        for i, row in enumerate(table.rows):
            row_cols = 0
            for j, cell in enumerate(row.cells):
                if cell._element in merged:
                    continue
                c_xml = xmltodict.parse(cell._element.xml, process_namespaces=False)
                # Cell width
                try:
                    cell_width = int(c_xml['w:tc']['w:tcPr']['w:tcW']['@w:w'])
                    text_cell = (cell_width / self.width) >= 0.8
                except KeyError:
                    #c_xml['w:tc']['w:tcPr']['w:shd'] # !!! rest of row from other columns
                    text_cell = False
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
        frame = self.investigate_table(table)
        if frame:
            left_space, right_space, top_space, bottom_space, text_rows = frame
        else:
            text_rows = []
        table_links = []
        html_content = ''
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
        self.tables_cnt += 1
        title, anchor = self.get_table_title()
        classes = self.get_depth_classes()
        filled = ''
        html_table = f'<table id="{anchor}"class="w3-table w3-hoverable {classes}" title="{title}">'
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
                try:
                    cell_width = int(c_xml['w:tc']['w:tcPr']['w:tcW']['@w:w'])
                    text_cell = (cell_width / self.width) >= 0.8
                except KeyError:
                    #c_xml['w:tc']['w:tcPr']['w:shd'] # !!! rest of row from other columns
                    text_cell = False
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
                if ignore:
                    continue
                if i in text_rows and text_cell:
                    # close table
                    html_table += '</tr></table>'
                    # check is table filled
                    if filled:
                        html_content += html_table
                        table_links.append(f'<a href="#{anchor}">{make_toc_header(title, self.depth + 1)}</a><br>')
                    html_content += text
                    # new table
                    self.tables_cnt += 1
                    classes = self.get_depth_classes()
                    title, anchor = self.get_table_title()
                    filled = ''
                    html_table = f'<table id="{anchor}"class="w3-table w3-hoverable {classes}" title="{title}">'
                    if rowspan > 1 or colspan > 1:
                        merged.add(cell._element)
                    continue
                else:
                    if rowspan > 1 or colspan > 1:
                        html_table += f'<td rowspan="{rowspan}" colspan="{colspan}"{css}>{text}</td>'
                        merged.add(cell._element)
                    else:
                        html_table += f'<td{css}>{text}</td>'
                    filled += text.strip()
            html_table += "</tr>"
        html_table += "</table>"
        if filled.strip():
            html_content += html_table
            table_links.append(f'<a href="#{anchor}">{make_toc_header(title, self.depth + 1)}</a><br>')
        return html_content, table_links
        
    
def make_toc_header(text, depth, max_len=35):
    text = '__' * (depth - 1) + text
    if len(text) > max_len:
        text = text[:max_len] + '...'
    return text


def paragraph_style(par):
    css = ''
    try:
        css += 'text-align: {};'.format(ALIGNMENT[par.alignment.name])
    except (KeyError, AttributeError):
        pass
    if par.style.font.bold:
        css += 'font-weight: bold;'
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
    return ''.join(html_content), ''.join(toc_links)
