from typing import Union
import docx
import xmltodict


class Node:
    def __init__(self, num_prefix: str = '', depth: int = 0, source: Union[str, None] = None,
                 _id = None):
        self._id = _id
        self.parents = {}
        self.source = source
        self.num_prefix = num_prefix
        self.depth = depth
        
        
class ParHandler:
    def __init__(self, par: docx.text.paragraph.Paragraph):
        self.par = par
        self.ctext = par.text.strip()
        self.node = Node()
        self.font_size = self.get_par_font_size(par)
        self.bold = self.get_par_bold_option(par)
        self.xml = xmltodict.parse(par._p.xml, process_namespaces=False)
        self.style_id = par.style.style_id
        self.base_style = par.style.base_style
        self.base_style_id = self.base_style.style_id if self.base_style else None
        self.style_name = par.style.name
        
    def get_par_font_size(self, par:  docx.text.paragraph.Paragraph):
        font_sizes = []
        if par.style.font.size:
            font_sizes.append(par.style.font.size.pt)
        font_sizes += [run.font.size.pt for run in par.runs if run.font.size]
        return max(font_sizes) if font_sizes else None
    
    def get_par_bold_option(self, par:  docx.text.paragraph.Paragraph):
        runs_bold_frac = sum([run.bold is True for run in par.runs]) / (len(par.runs) + 1)
        return par.style.font.bold is True or runs_bold_frac > 0.6
    
    def get_full_text(self):
        if self.node.source not in ('HEADING', 'REGEX', 'APPENDIX') \
            and self.node.num_prefix \
            and 'default' not in self.node.num_prefix:
            return self.node.num_prefix + ' ' + self.ctext
        else:
            return self.ctext


class TableHandler:
    def __init__(self, table: docx.table.Table, src_page_width: int, src_page_height: int,
                 text_cell_min_width: float = 0.8, frame_table_min_hight: float = 0.8,
                 min_frame_columns: int = 7):
        self.table = table
        self.xml = xmltodict.parse(table._element.xml, process_namespaces=False)
        self.height = self.get_table_height(self.xml)
        self.width = self.get_table_width(self.xml)
        self.cols_count = len(self.table.columns)
        self.rows_count = len(self.table.rows)
        self.src_page_width = src_page_width
        self.src_page_height = src_page_height
        self.text_cell_min_width = text_cell_min_width
        self.frame_table_min_hight = frame_table_min_hight
        self.min_frame_columns = min_frame_columns
        self.merged = set()
        self.rows = []
        self.text_col_starts = self.cols_count
        self.text_col_ends = -1
        self.text_row_starts = self.rows_count
        self.text_row_ends = -1
        
        self.investigate()
        self.has_frame = self.detect_frame()

    def investigate(self):
        for i, row in enumerate(self.table.rows):
            cells = []
            for j, cell in enumerate(row.cells):
                if cell._element in self.merged:
                    continue
                
                # Detect merged cells
                rowspan = 1
                colspan = 1
                for next_row in self.table.rows[i + 1:]:
                    if next_row.cells[j]._element == cell._element:
                        rowspan += 1
                    else:
                        break
                for next_cell in row.cells[j + 1:]:
                    if next_cell._element == cell._element:
                        colspan += 1
                    else:
                        break
                if rowspan > 1 or colspan > 1:
                    self.merged.add(cell._element)
                cell_handler = CellHandler(cell, rowspan, colspan, j, i)
                
                # Detect text (for frames) cell
                cell_handler.is_text = cell_handler.width / self.src_page_width > self.text_cell_min_width
                if cell_handler.is_text:
                    self.text_row_starts = min(self.text_row_starts, i)
                    self.text_row_ends = max(self.text_row_ends, i + rowspan)
                    self.text_col_starts = min(self.text_col_starts, j)
                    self.text_col_ends = max(self.text_col_ends, j + colspan)
            
                # cell_handler.ctext = f'DBG [COLS / MIN_COLS {len(self.table.columns)} / {self.min_frame_columns}; PAGE W = {self.src_page_width}; PAGE H = {self.src_page_height};TABLE W = {self.width}; TABLE H = {self.height}; XML W = {cell_handler.width}; TEXT = {cell_handler.is_text}; {(self.text_row_starts, self.text_row_ends, self.text_col_starts, self.text_col_ends)}, X = {cell_handler.x}, Y = {cell_handler.y}]' + cell_handler.ctext
                cells.append(cell_handler)
            self.rows.append(cells)

    def detect_frame(self):
        # Table hight far from page height and page is portrait
        if (self.height / self.src_page_height) < self.frame_table_min_hight \
            and (self.width < self.src_page_width):
            return False
        # Table hcols count lower than min cols count in frame
        if self.cols_count < self.min_frame_columns:
            return False
        return self.text_row_starts >= 0

    def get_table_height(self, xml):
        try:
            return sum([int(row['w:trPr']['w:trHeight']['@w:val']) for row in xml['w:tbl']['w:tr']])
        except (KeyError, TypeError):
            return 0

    def get_table_width(self, xml):
        try:
            return sum([int(col['@w:w']) for col in xml['w:tbl']['w:tblGrid']['w:gridCol']])
        except (KeyError, TypeError):
            return 0

        
class CellHandler:
    def __init__(self, cell, rowspan: int, colspan: int, x: int, y: int):
        self.x = x
        self.y = y
        self.paragraphs = cell.paragraphs
        self.xml = xmltodict.parse(cell._element.xml, process_namespaces=False)
        try:
            self.width = int(self.xml['w:tc']['w:tcPr']['w:tcW']['@w:w'])
        except KeyError:
            self.width = 0
        self.is_text = False
        self.rowspan = rowspan
        self.colspan = colspan
        self.ctext = '\n'.join([c_par.text.strip() for c_par in self.paragraphs]).strip()
        
        
class TableView:
    def __init__(self, node: Node):
        self.rows = []
        self.node = node
    
    def empty(self):
        n_chars = 0
        for r in self.rows:
            for c in r:
                n_chars += len(c.ctext)
        return n_chars == 0
    

class DocRoot:
    def __init__(self):
        self.node = Node('[Начало документа]', 1, 'ROOT')
        self.node._id = 'default-start-doc'
