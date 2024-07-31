import re
from typing import List, Union
import docx
from loguru import logger
import xmltodict
from .core import ParHandler, TableHandler, TableView, Node, DocRoot
from .numbering import NumberingDB


class DocHandler:
    """
    Handles the conversion of DOCX document content to HTML.
    """
    def __init__(self, doc: docx.Document, default_width: int = 11907, default_height: int = 16840,
                 max_frame_space: int = 7, max_toc_pages: int = 10, max_doc_pages: int = 2000,
                 avg_page_chars_count: int = 1200):
        """
        Initializes the DocHandler with a DOCX document.
        
        Args:
            doc (docx.Document): The DOCX document to process.
        """
        self.doc = doc
        self.xml = xmltodict.parse(doc.element.xml, process_namespaces=False)
        self.num_db = NumberingDB(doc)
        self.chars_count = 0
        self.last_depth = 1
        self.last_pars = []
        self.processed_content = [DocRoot()]
        self.depth_anchor = {1: self.processed_content[0].node._id}

        try:
            self.width = int(self.xml['w:document']['w:body']['w:sectPr']['w:pgSz']['@w:w'])
        except KeyError:
            self.width = default_width
        try:
            self.height = int(self.xml['w:document']['w:body']['w:sectPr']['w:pgSz']['@w:h'])
        except KeyError:
            self.height = default_height

        self.max_frame_space = max_frame_space
        self.max_toc_pages = max_toc_pages
        self.max_doc_pages = max_doc_pages
        self.avg_page_chars_count = avg_page_chars_count
        self.processed = False
        
    def process(self):
        for content in self.doc.iter_inner_content():
            if type(content) is docx.text.paragraph.Paragraph:
                self.process_paragraph(content)
            elif type(content) is docx.table.Table:
                self.process_table(content)
            else:
                logger.warning(type(content), 'missed')
        self.processed = True
        
    def insert_node(self, node: Node):
        self.last_depth = node.depth
        anchor = f'par{len(self.processed_content)}'
        self.depth_anchor[node.depth] = anchor
        return anchor
        
    def detect_toc_row(self, par: ParHandler) -> bool:
        if '.....' in par.ctext:
            return True
        if (self.chars_count // self.avg_page_chars_count) > self.max_toc_pages:
            return False
        # check if row ends with number
        match = re.search(r'(\d+)$', par.ctext)
        if match:
            # check if number less than max pages count
            return int(match[0]) < self.max_doc_pages
        else:
            return False

    def get_parents(self) -> List:
        return {k: v for k, v in self.depth_anchor.items() if k <= self.last_depth}
    
    def get_table_title(self) -> Node:
        """
        Retrieves the title for a table.
        
        Returns:
            tuple: A tuple containing the title and anchor for the table.
        """
        regex_title = ' '.join(self.last_pars)
        if 'таблица' in regex_title:
            title = regex_title[regex_title.lower().rindex('таблица')]
        elif 'т а б л и ц а' in regex_title:
            title = regex_title[regex_title.lower().rindex('т а б л и ц а')]
        elif regex_title.strip():
            title = regex_title
        else:
            title = 'Таблица'
        depth = self.last_depth  + 1
        table_node = Node(title, depth, 'TABLE')
        table_node._id = f'table{len(self.processed_content)}'
        return table_node

    def process_paragraph(self, par: docx.text.paragraph.Paragraph) -> tuple:
        """
        Processes a text paragraph.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the HTML content and table of contents links.
        """
        # Update doc numeration
        par = self.num_db.numerize(ParHandler(par))
        if par.ctext:
            self.last_pars.append(par.ctext)
            self.last_pars = self.last_pars[-2:]
        else:
            return None
        # Check TOC row
        if self.detect_toc_row(par):
            par.node.depth = 0
        # Link node
        if par.node.depth:
            par.node._id = self.insert_node(par.node)
        par.node.parents = self.get_parents()
        # Count page
        self.chars_count += len(par.ctext)
        # Store processed paragraph
        # par.ctext = f'DBG [{par.node.source}] ' + par.ctext
        self.processed_content.append(par)

    def process_table(self, table: docx.table.Table) -> tuple:
        """
        Processes a table.
        
        Args:
            table (docx.table.Table): The table to process.
        
        Returns:
            tuple: A tuple containing the HTML content and table of contents links.
        """
        table = TableHandler(table, self.width, self.height)
        subtable = TableView(self.get_table_title())
        for i, row in enumerate(table.rows):
            if not any([cell.is_text for cell in row]):
                if table.has_frame:
                    # Find visible cells inside frame borders
                    visible_cells = [
                        cell for cell in row
                        if table.text_col_starts <= cell.x < table.text_col_ends \
                            and table.text_row_starts <= cell.y < table.text_row_ends
                    ]
                else:
                    visible_cells = row
                if visible_cells:
                    subtable.rows.append(visible_cells)
            else:
                # Close table
                self.append_table(subtable)
                # Process cell paragraphes
                text_cell = [cell for cell in row if cell.is_text][0]
                for c_par in text_cell.paragraphs:
                    self.process_paragraph(c_par)
                # Create new table
                subtable = TableView(self.get_table_title())
        # Close last opened table
        self.append_table(subtable)
            
    def append_table(self, table: TableView):
        if table.empty():
            return
        if table_extend(self.processed_content[-1], table):
            self.processed_content[-1] = concat_tables(self.processed_content[-1], table)
        else:
            table.node.parents = self.get_parents()
            self.processed_content.append(table)


def table_extend(prev_element: TableView, next_element: TableView) -> bool:
    # Check if both tags are <table> tags
    if not (isinstance(prev_element, TableView) and isinstance(next_element, TableView)):
        return False
    # Tables are considered one if they have the same number of columns
    return len(prev_element.rows[-1]) == len(next_element.rows[0])


def concat_tables(prev_table: TableView, next_table: TableView) -> str:
    prev_table_header = '\t'.join([cell.ctext for cell in prev_table.rows[0]])
    next_table_header = '\t'.join([cell.ctext for cell in next_table.rows[0]])
    # Remove the header from the next table if it matches the header of the previous table
    if prev_table_header == next_table_header:
        next_table.rows = next_table.rows[1:]
    # Concat tables
    prev_table.rows.extend(next_table.rows)
    return prev_table
