import os
from collections import defaultdict
import json
from typing import List, Union
from override.callbacks import custom_callback
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
                'content-type': 'text/title',
                'title': par.get_full_text(),
                'sub-title': '',
                'content': ''
            }
        # Subtitle element
        elif par.node.depth > 1:
            title_par = self.indexed_pars[par.node.parents[1]]
            prefix_ends = 0 if par.node.source in ['HEADING', 'APPENDIX'] else len(par.node.num_prefix)
            el = {
                'content-type': 'text/subtitle',
                'title': title_par.get_full_text(),
                'sub-title': make_title(par.get_full_text()),
                'content': par.get_full_text()[prefix_ends:].strip()
            }
        # Regular text element
        elif par.node.depth == 0:
            title_par = self.indexed_pars[par.node.parents[1]]
            sub_title_par = self.indexed_pars[par.node.parents[max(par.node.parents.keys())]]
            el = {
                'content-type': 'text',
                'title': title_par.get_full_text(),
                'sub-title': make_title(sub_title_par.get_full_text()),
                'content': par.get_full_text()
            }
        self.elements.append(el)
        
    def get_table_content_range(self, table):
        header_cols_count = len(table.rows[0])
        left_top_cell = table.rows[0][0]
        if header_cols_count > 1:
            row_prefix = left_top_cell.ctext + ': ' if left_top_cell.ctext else ''
            content_x_left = left_top_cell.x + left_top_cell.colspan
        else:
            row_prefix = ''
            content_x_left = 0
        content_y_top = left_top_cell.y + left_top_cell.rowspan
        return content_x_left, content_y_top, row_prefix

    def table_json(self, table: TableView):
        # Table parent node
        sub_title_par = self.indexed_pars[table.node.parents[max(table.node.parents.keys())]]
        # Find table index for rows and cols
        content_x_left, content_y_top, row_prefix = self.get_table_content_range(table)
        # Process tavle content
        content = []
        i = 0
        for row in table.rows:
            j = 0
            content_row = []
            for cell in row:
                # Check if cell is content
                if cell.x >= content_x_left and cell.y >= content_y_top and cell.ctext:
                    # Increment rows, cols counters
                    i += (j == 0)
                    j += 1
                    # Indexing content cell
                    content_row.append({
                        "row": i,
                        "col": j,
                        "sub-title-row": row_prefix + get_left_index(table, cell, content_x_left),
                        "sub-title-col": get_top_index(table, cell, content_y_top),
                        "value": cell.ctext
                    })
            # Add cells grouped by index
            grouped_cells = group_by_index(content_row)
            content.extend(grouped_cells)
        self.elements.append({
            'content-type': 'table',
            'title': table.node.num_prefix,
            'sub-title': make_title(sub_title_par.get_full_text()),
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

        # POSTPROCESS WITH CUSTOM CALLBACK
        processed_elements = []
        for element in self.elements:
            action, updated_element = custom_callback(element)
            if action == 'pass':
                processed_elements.append(element)
            elif action == 'update':
                processed_elements.append(updated_element)
            elif action == 'remove':
                continue
            else:
                raise ValueError(f'Recieved invalid action "{action}"')
            
        return json.dumps(processed_elements, ensure_ascii=False)
            

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
            # Index cell out of possible content cell range
            if idx_cell.x >= content_x_left or idx_cell.y > cell.y:
                break
            # Index cell in content cell rows range
            if idx_cell.y <= cell.y \
                and (cell.y + cell.rowspan) <= (idx_cell.y + idx_cell.rowspan) \
                and idx_cell.ctext:
                index.append(idx_cell.ctext)
    return ': '.join(index)


def get_top_index(table, cell, content_y_top):
    index = []
    for idx_row in table.rows:
        for idx_cell in idx_row:
            # Index cell out of possible content cell range
            if idx_cell.y >= content_y_top or idx_cell.x > cell.x:
                break
            # Index cell in content cell cols range
            if idx_cell.x <= cell.x \
                and (cell.x + cell.colspan) <= (idx_cell.x + idx_cell.colspan) \
                and idx_cell.ctext:
                index.append(idx_cell.ctext)
    return ': '.join(index)


def group_by_index(cells: List[dict]) -> List[dict]:
    grouped_cells = []
    group_dict = defaultdict(lambda: {"row": None, "col": float('inf'), "values": []})

    for cell in cells:
        index_row = cell["sub-title-row"]
        index_col = cell["sub-title-col"]
        key = (index_row, index_col)

        if group_dict[key]["row"] is None:
            group_dict[key]["row"] = cell["row"]
        group_dict[key]["col"] = min(group_dict[key]["col"], cell["col"])
        group_dict[key]["values"].append(cell["value"])

    for (index_row, index_col), data in group_dict.items():
        grouped_cells.append({
            "row": data["row"],
            "col": data["col"],
            "sub-title-row": index_row,
            "sub-title-col": index_col,
            "value": data["values"] if len(data["values"]) > 1 else data["values"][0]
        })

    return grouped_cells
