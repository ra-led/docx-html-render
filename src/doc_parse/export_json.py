import json
from .core import ParHandler, TableView, DocRoot
from .ooxml import DocHandler

class DocJSON:
    def __init__(self):
        root = DocRoot()
        self.elements = []
        self.indexed_pars = {root.node._id: root}
        
    def paragraph_json(self, par: ParHandler):
        # Title elelment
        if par.node.depth == 1:
            el = {
                'content_type': 'text/title',
                'title': par.get_full_text(),
                'sub_title': '',
                'content': ''
            }
        # Subtitle element
        elif par.node.depth > 1:
            title_par = self.indexed_pars[par.node.parents[1]]
            prefix_ends = 0 if par.node.source in ['HEADING', 'APPENDIX'] else len(par.node.num_prefix)
            el = {
                'content_type': 'text/subtitle',
                'title': title_par.get_full_text(),
                'sub_title': make_title(par.get_full_text()),
                'content': par.get_full_text()[prefix_ends:].strip()
            }
        # Regular text element
        elif par.node.depth == 0:
            title_par = self.indexed_pars[par.node.parents[1]]
            sub_title_par = self.indexed_pars[par.node.parents[max(par.node.parents.keys())]]
            el = {
                'content_type': 'text',
                'title': title_par.get_full_text(),
                'sub_title': make_title(sub_title_par.get_full_text()),
                'content': par.get_full_text()
            }
        self.elements.append(el)

    def table_json(self, table: TableView):
        sub_title_par = self.indexed_pars[table.node.parents[max(table.node.parents.keys())]]
        left_top_cell = table.rows[0][0]
        row_prefix = left_top_cell.ctext + ': ' if left_top_cell.ctext else ''
        content_x_left = left_top_cell.x + left_top_cell.colspan
        content_y_top = left_top_cell.y + left_top_cell.rowspan
        content = []
        i = 0
        for row in table.rows:
            j = 0
            for cell in row:
                # Check if cell is content
                if cell.x >= content_x_left and cell.y >= content_y_top and cell.ctext:
                    # Increment rows, cols counters
                    i += (j == 0)
                    j += 1
                    # Indexing content cell
                    content.append({
                        "row": i,
                        "col": j,
                        "sub-title-row": row_prefix + get_left_index(table, cell, content_x_left),
                        "sub-title-col": get_top_index(table, cell, content_y_top),
                        "value": cell.ctext
                    })
        self.elements.append({
            'content-type': 'table',
            'title': table.node.num_prefix,
            'sub_title': make_title(sub_title_par.get_full_text()),
            'content': content
        })
        # print(self.elements[-1]['title'], '\n', self.elements[-1], '\n', '='*80)
        
    def get_json(self, handler: DocHandler) -> tuple:
        if not handler.processed:
            handler.process()
        for content in handler.processed_content:
            if type(content) is ParHandler:
                if content.node._id:
                    self.indexed_pars[content.node._id] = content
                self.paragraph_json(content)
            elif type(content) is TableView:
                self.table_json(content)
        return json.dumps(self.elements, ensure_ascii=False)
            

def make_title(text: str, max_len: int = 35) -> str:
    """
    Creates header.
    
    Args:
        text (str): The text of the content.
        max_len (int, optional): The maximum length of the title text.
    
    Returns:
        str: The formatted element title.
    """
    if len(text) > max_len:
        text = text[:max_len] + '...'
    return text


def get_left_index(table, cell, content_x_left):
    index = []
    for idx_row in table.rows:
        for idx_cell in idx_row:
            if idx_cell.x >= content_x_left or idx_cell.y > cell.y:
                break
            if idx_cell.y <= cell.y \
                and (cell.y + cell.rowspan) <= (idx_cell.y + idx_cell.rowspan) \
                and idx_cell.ctext:
                index.append(idx_cell.ctext)
    return ': '.join(index)


def get_top_index(table, cell, content_y_top):
    index = []
    for idx_row in table.rows:
        for idx_cell in idx_row:
            if idx_cell.y >= content_y_top or idx_cell.x > cell.x:
                break
            if idx_cell.x <= cell.x \
                and (cell.x + cell.colspan) <= (idx_cell.x + idx_cell.colspan) \
                and idx_cell.ctext:
                index.append(idx_cell.ctext)
    return ': '.join(index)
