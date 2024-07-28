import asyncio
import os
import string
import uuid
import logging
import statistics
from typing import List, Union
from bs4 import BeautifulSoup

import aspose.words as aw
from aio_pika import Message, connect
import docx
import re
import xmltodict
import html

from ml import BERTTextClassifier


logger = logging.getLogger(__name__)

async def get_connection():
    """
    Establishes a connection to the RabbitMQ server.

    Returns:
        aio_pika.Connection: The connection object to the RabbitMQ server.
    """
    user = os.environ.get('RABBITMQ_USER', default='guest')
    pasw = os.environ.get('RABBITMQ_PASS', default='guest')
    host = os.environ.get('RABBITMQ_HOST', default='rabbitmq')
    port = os.environ.get('RABBITMQ_PORT', default=5672)

    return await connect(f'amqp://{user}:{pasw}@{host}:{port}')

def doc_to_docx(in_stream, out_stream):
    """
    Converts a .doc file to a .docx file using Aspose.Words.

    Args:
        in_stream (io.BytesIO): The input stream containing the .doc file.
        out_stream (io.BytesIO): The output stream to write the .docx file.
    """
    doc = aw.Document(in_stream)
    doc.save(out_stream, aw.SaveFormat.DOCX)

class ConverterProxy:
    """
    A proxy class to handle document conversion requests via RabbitMQ.
    """

    def __init__(self):
        """
        Initializes the ConverterProxy instance.
        """
        self.initialized = False
        self.futures = {}

    async def convert(self, data: bytes):
        """
        Sends a document conversion request to the RabbitMQ queue and waits for the response.

        Args:
            data (bytes): The document data to be converted.

        Returns:
            bytes: The converted document data.
        """
        if not self.initialized:
            self.initialized = True
            self.connection = await get_connection()
            self.channel = await self.connection.channel()
            self.callback_queue = await self.channel.declare_queue(exclusive=True)
            await self.callback_queue.consume(self.on_message, no_ack=True)

        correlation_id = str(uuid.uuid4())
        loop = asyncio.get_running_loop()
        future = loop.create_future()

        self.futures[correlation_id] = future

        await self.channel.default_exchange.publish(
            Message(
                data,
                correlation_id=correlation_id,
                reply_to=self.callback_queue.name
                ),
            routing_key=os.environ.get('CONVERTER_QUEUE', default='convert')
            )
        return await future

    async def on_message(self, message):
        """
        Handles incoming messages from the RabbitMQ callback queue.

        Args:
            message (aio_pika.IncomingMessage): The incoming message from the RabbitMQ queue.
        """
        if message.correlation_id is None:
            print(f"Bad message {message!r}")
            return

        future: asyncio.Future = self.futures.pop(message.correlation_id)
        future.set_result(message.body)


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
    {'@w:ilvl': '0', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default %1'}},
    {'@w:ilvl': '1', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default %1.%2'}},
    {'@w:ilvl': '2', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default %1.%2.%3'}},
    {'@w:ilvl': '3', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default %1.%2.%3.%4'}},
    {'@w:ilvl': '4', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default %1.%2.%3.%4.%5'}},
    {'@w:ilvl': '5', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default %1.%2.%3.%4.%5.%6'}},
    {'@w:ilvl': '6', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default %1.%2.%3.%4.%5.%6.%7'}},
    {'@w:ilvl': '7', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default %1.%2.%3.%4.%5.%6.%7.%8'}},
    {'@w:ilvl': '8', 'w:start': {'@w:val': '1'}, 'w:numFmt': {'@w:val': 'decimal'}, 'w:lvlText': {'@w:val': 'default %1.%2.%3.%4.%5.%6.%7.%8.%9'}}
]


class NumberingDB:
    """
    Handles numbering and styles in a DOCX document.
    """
    def __init__(self, doc: docx.Document, appendix_header_length: int = 40,
                 norm_numeration_model: str = 'model_dir/num_clf',
                 norm_heading_model: str = 'model_dir/word_clf'):
        """
        Initializes the NumberingDB with a DOCX document.
        
        Args:
            doc (docx.Document): The DOCX document to process.
        """
        self.font_size = []
        try:
            self.num_xml = xmltodict.parse(
                doc.part.numbering_part.element.xml,
                process_namespaces=False
            )
            try:
                abstract_levels = self.num_xml['w:numbering']['w:abstractNum']
            except KeyError:
                abstract_levels = []
            if type(abstract_levels) is list:
                self.levels = {
                    x['@w:abstractNumId']: x['w:lvl'] if type(x['w:lvl']) is list else [x['w:lvl']]
                    for x in abstract_levels
                }
            else:
                self.levels = {
                    abstract_levels['@w:abstractNumId']: abstract_levels['w:lvl']
                    if type(abstract_levels['w:lvl']) is list else [abstract_levels['w:lvl']]
                }
            try:
                self.nums_to_abstarct = {
                    x['@w:numId']: x['w:abstractNumId']['@w:val']
                    for x in self.num_xml['w:numbering']['w:num']
                }
            except KeyError:
                self.nums_to_abstarct = {}
        except NotImplementedError:
            self.levels = {}
            self.nums_to_abstract = {}
        self.styles_xml = xmltodict.parse(
            doc.part.styles.element.xml,
            process_namespaces=False
        )
        self.styles = {
            x['@w:styleId']: x
            for x in self.styles_xml['w:styles']['w:style']
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
        self.appendix_header_length = appendix_header_length
        
        self.norm_numeration_clf = BERTTextClassifier(norm_numeration_model)
        self.norm_heading_clf = BERTTextClassifier(norm_numeration_model)
        
    def get_abs_id(self, numId: Union[str, None] = None, styleId: Union[str, None] = None) -> Union[str, None]:
        """
        Retrieves the abstract number ID for a given number ID or style ID.
        
        Args:
            numId (str, optional): The number ID.
            styleId (str, optional): The style ID.
        
        Returns:
            str: The abstract number ID.
        """
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
    
    def check_heading_style(self, par: docx.text.paragraph.Paragraph) -> bool:
        """
        Checks if a paragraph has a heading style.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to check.
        
        Returns:
            bool: True if the paragraph has a heading style, False otherwise.
        """
        if re.findall('^таблица', par.text.strip().lower()):
            return False
        if re.findall('^рисунок', par.text.strip().lower()):
            return False
        bold = any([par.style.font.bold] + [run.bold for run in par.runs])
        regular_font_size = statistics.median(self.font_size) if self.font_size else 12
        font_sizes = []
        if par.style.font.size:
            font_sizes.append(par.style.font.size.pt)
        font_sizes += [run.font.size.pt for run in par.runs if run.font.size]
        large_font = any([x > regular_font_size for x in font_sizes])
        if bold or large_font:
            return True
        else:
            return False
    
    def count_builtin(self, absId: str, level: int, par: docx.text.paragraph.Paragraph) -> tuple:
        """
        Counts the built-in numbering for a given abstract number ID and level.
        
        Args:
            absId (str): The abstract number ID.
            level (int): The level of the numbering.
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        self.increment[absId][level] += 1
        for lvl_i in self.increment[absId]:
            if lvl_i > level:
                self.increment[absId][lvl_i] = 0
        abstarct_levels = self.levels[absId]
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
            try:
                num_fmt = lvl_a['w:numFmt']['@w:val']
                if num_fmt == 'upperLetter':
                    num = string.ascii_uppercase[num - 1]
                elif num_fmt == 'lowerLetter':
                    num = string.ascii_lowercase[num - 1]
                elif num_fmt == 'upperRoman':
                    num = int_to_roman(num)
                elif num_fmt == 'lowerRoman':
                    num = int_to_roman(num).lower()
            except KeyError:
                pass
            if f'%{lvl_i + 1}' in num_prefix:
                depth += 1
                num_prefix = re.sub(f'%{lvl_i + 1}', str(num), num_prefix)
        return num_prefix, depth, absId
    
    def numrize_by_meta(self, par: docx.text.paragraph.Paragraph) -> tuple:
        """
        Processes numbering by metadata.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        p_xml = xmltodict.parse(par._p.xml, process_namespaces=False)
        try:
            numId = p_xml['w:p']['w:pPr']['w:numPr']['w:numId']['@w:val']
        except:
            return '', 0, None
        level = int(p_xml['w:p']['w:pPr']['w:numPr']['w:ilvl']['@w:val'])
        absId = self.get_abs_id(numId=numId)
        num_prefix, depth, source = self.count_builtin(absId, level, par)
        if  not self.check_heading_style(par) and depth == 1:
            depth = 0
        # sublist always ends with ")"
        if any([symb in num_prefix for symb in [')', '-', ':']]):
            depth = 0
        # sublist text always starts with '-'
        try:
            if any([symb in num_prefix for symb in ['-', '–', '—', '−']]) or\
                any([symb == par.text.strip()[0] for symb in ['-', '–', '—', '−']]):
                depth = 0
        except IndexError:
            return '', 0, None
        return num_prefix, depth, source
    
    def numrize_by_style(self, par: docx.text.paragraph.Paragraph) -> tuple:
        """
        Processes numbering by style.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        style_abs = self.get_abs_id(styleId=par.style.style_id)
        if style_abs is None:
            base_style_id = par.style.base_style.style_id if par.style.base_style else None
            style_abs = self.get_abs_id(styleId=base_style_id)
        if style_abs is None:
            return '', 0, None
        absId, level = style_abs['absId'], style_abs['lvl']
        return self.count_builtin(absId, level, par)
    
    def numerize_by_text(self, par: docx.text.paragraph.Paragraph) -> tuple:
        """
        Detect numbering in text prefix.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        depth = 0
        text = par.text.strip()
        num_prefix = ''
        
        letter_pattern = r'^(\w\.)\d'
        match = re.findall(letter_pattern, text)
        if match:
            text = re.sub(r'^\w\.', '', text)
            num_prefix += match[0]
            depth += 1
        
        numbering_pattern = r'^\d+\.'
        while 1:
            match = re.findall(numbering_pattern, text)
            if not match:
                break
            depth += 1
            text = re.sub(numbering_pattern, '', text)
            num_prefix += match[0]
        
        numbering_pattern = r'^\d+\s'
        match = re.findall(numbering_pattern, text.strip())
        if match:
            depth += 1
            text = re.sub(numbering_pattern, '', text)
            num_prefix += match[0]
        # sublist text always starts with '-'
        try:
            if any([symb in num_prefix for symb in ['-', '–', '—', '−']]) or \
                any([symb == text.strip()[0] for symb in ['-', '–', '—', '−']]):
                return '', 0, None
        except IndexError:
            return '', 0, None
        # check style
        if self.check_heading_style(par) or depth > 1:
            if self.norm_numeration_clf(par.text):
                return num_prefix, depth, 'REGEX'
            else:
                return '', 0, None
        else:
            return '', 0, None

    def numerize_by_heading(self, par: docx.text.paragraph.Paragraph) -> tuple:
        """
        Processes numbering by heading.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        depth = 0
        style = par.style.name
        if style:
            match = re.search(r'Heading (\d+)', style)
            if match:
                depth = 1
            elif style == 'Title':
                depth = 1
        if self.check_heading_style(par) and depth > 0:
            if self.norm_heading_clf(par.text):
                return par.text if par.text.strip() else '[UNNAMED]', depth, 'HEADING'
            else:
                return '', 0, None
        else:
            return '', 0, None
        
    def numerize_by_appendix(self, par: docx.text.paragraph.Paragraph) -> tuple:
        """
        Processes numbering by detecting appendix header.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        text = par.text.strip().split('\n')[0]
        match = re.search(r'^приложение', text.lower())
        if match and len(par.text.strip()) < self.appendix_header_length:
            return text, 1, 'APPENDIX'
        else:
            return '', 0, None
    
    def numerize(self, par: docx.text.paragraph.Paragraph) -> tuple:
        """
        Processes numbering for a paragraph.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        numerize_prioritet = [
            self.numrize_by_meta,
            self.numrize_by_style,
            self.numerize_by_text,
            self.numerize_by_heading,
            self.numerize_by_appendix
        ]
        for method in numerize_prioritet:
            num_prefix, depth, source = method(par)
            if num_prefix:
                return num_prefix, depth, source
        return '', 0, None
        

class DocHandler:
    """
    Handles the conversion of DOCX document content to HTML.
    """
    def __init__(self, doc: docx.Document):
        """
        Initializes the DocHandler with a DOCX document.
        
        Args:
            doc (docx.Document): The DOCX document to process.
        """
        self.xml = xmltodict.parse(doc.element.xml, process_namespaces=False)
        self.num_db = NumberingDB(doc)
        self.depth = 1
        self.source = None
        self.depth_anchor = {1: 'default-start-doc'}
        self.tables_cnt = 0
        self.page = 0
        self.chars_count = 0
        self.last_pars = []

        try:
            self.width = int(self.xml['w:document']['w:body']['w:sectPr']['w:pgSz']['@w:w'])
        except KeyError:
            self.width = 11907
        try:
            self.height = int(self.xml['w:document']['w:body']['w:sectPr']['w:pgSz']['@w:h'])
        except KeyError:
            self.height = 16840

        self.max_frame_space = 7
        self.max_toc_pages = 10
        self.max_doc_pages = 1000
        self.avg_page_chars_count = 1200
    
    def detect_toc_row(self, par: docx.text.paragraph.Paragraph) -> bool:
        if '.....' in par.text:
            return True
        if self.page > self.max_toc_pages:
            return False
        text = par.text.strip()
        match = re.search(r'(\d+)$', text)
        if match:
            return int(match[0]) < self.max_doc_pages
        else:
            return False

    def get_depth_classes(self) -> str:
        """
        Retrieves the depth classes for HTML elements.
        
        Returns:
            str: The depth classes as a string.
        """
        aa = []
        for k, v in self.depth_anchor.items():
            if k <= self.depth:
                aa.append(v)
        return " ".join(aa)
    
    def get_table_title(self) -> tuple:
        """
        Retrieves the title for a table.
        
        Returns:
            tuple: A tuple containing the title and anchor for the table.
        """
        regex_title = ' '.join(self.last_pars)
        try:
            strat_idx = max(
                regex_title.lower().rindex('таблица'),
                regex_title.lower().rindex('т а б л и ц а')
            )
            title = regex_title[strat_idx:]
        except ValueError:
            try:
                title = self.last_pars[-1]
            except IndexError:
                title = 'Таблица'
        title = html.escape(title if title.strip() else 'Таблица')
        anchor = f'table{self.tables_cnt}'
        return title, anchor
    
    def process_paragraph(self, par: docx.text.paragraph.Paragraph) -> tuple:
        """
        Processes a paragraph to convert it to HTML.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the HTML content and table of contents links.
        """
            
        html_paragraph = []
        html_links = []
        num_prefix, depth, source = self.num_db.numerize(par)
        
        if par.style.font.size:
            self.num_db.font_size.append(par.style.font.size.pt)
        self.num_db.font_size += [run.font.size.pt for run in par.runs if run.font.size]
        
        tag = 'p'
        if source not in ('HEADING', 'REGEX', 'APPENDIX') and num_prefix and 'default' not in num_prefix:
            text = num_prefix + ' ' + html.escape(par.text)
        else:
            text = html.escape(par.text) 
        # Check TOC row
        if self.detect_toc_row(par):
            depth = 0
        # Render
        if depth:
            anchor = 'a' + str(uuid.uuid4())
            self.depth = depth
            self.source = source
            self.depth_anchor[depth] = anchor
            tag = f'h{min(depth, 9)}'
            classes = self.get_depth_classes()
            html_links.append(f'<a href="#{anchor}">{make_toc_header(text, depth)}</a><br>')
            html_paragraph.append(f'<div class="{classes}"><{tag} id="{anchor}">{text}</{tag}></div>')
        else:
            classes = self.get_depth_classes()
            html_paragraph.append(f'<div class="{classes}"><{tag}>{text}</{tag}></div>')

        if text.strip():
            self.last_pars.append(text)
            self.last_pars = self.last_pars[-2:]
            
        # count page
        self.chars_count += len(par.text)
        self.page = self.chars_count // self.avg_page_chars_count
        return html_paragraph, html_links
    
    def investigate_table(self, table: docx.table.Table) -> Union[tuple, None]:
        """
        Investigates a table to determine its structure. Detects blueprint frame
        
        Args:
            table (docx.table.Table): The table to investigate.
        
        Returns:
            tuple: A tuple containing the blueprint's frame left, right, top, and bottom spaces, and text rows. Return None if no frame detected.
        """
        t_xml = xmltodict.parse(table._element.xml, process_namespaces=False)
        try:
            table_height = sum([
                int(row['w:trPr']['w:trHeight']['@w:val'])
                for row in t_xml['w:tbl']['w:tr']
            ])
            table_width = sum([
                int(col['@w:w'])
                for col in t_xml['w:tbl']['w:tblGrid']['w:gridCol']
            ])
        except (KeyError, TypeError):
            table_height = 0
            table_width = 0
        
        if (table_height / self.height) < 0.8 and (table_width < self.width):
            return None
        merged = set()
        all_text_cells = []
        for i, row in enumerate(table.rows):
            row_cols = 0
            for j, cell in enumerate(row.cells):
                if cell._element in merged:
                    continue
                # count page
                self.chars_count += len(cell.text)
                self.page = self.chars_count // self.avg_page_chars_count
                
                c_xml = xmltodict.parse(cell._element.xml, process_namespaces=False)
                try:
                    cell_width = int(c_xml['w:tc']['w:tcPr']['w:tcW']['@w:w'])
                    text_cell = (cell_width / self.width) >= 0.8
                except KeyError:
                    text_cell = False
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

    def process_table(self, table: docx.table.Table) -> tuple:
        """
        Processes a table to convert it to HTML.
        
        Args:
            table (docx.table.Table): The table to process.
        
        Returns:
            tuple: A tuple containing the HTML content and table of contents links.
        """
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
                try:
                    cell_width = int(c_xml['w:tc']['w:tcPr']['w:tcW']['@w:w'])
                    text_cell = (cell_width / self.width) >= 0.8
                except KeyError:
                    text_cell = False
                if ignore and not text_cell:
                    continue
                if text_cell:
                    text = ''
                    for c_par in cell.paragraphs:
                        html_paragraph, html_links = self.process_paragraph(c_par)
                        table_links += html_links
                        text += ''.join(html_paragraph)
                else:
                    text = ''
                    for c_par in cell.paragraphs:
                        if c_par.text.strip():
                            text += f'<p>{c_par.text}</p>'
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
                if i in text_rows and text_cell:
                    html_table += '</tr></table>'
                    if filled:
                        html_content += html_table
                        table_links.append(f'<a href="#{anchor}">{make_toc_header(title, self.depth + 1)}</a><br>')
                    html_content += text
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
    
    def process_tables_batch(self, contents_batch: List[str]):
        '''
        Function `process_tables_batch` combines several HTML code snippets into one. Sometimes passages may contain the beginning and end of the same table, so you need to combine them into one "<table>" tag.

        Parameters
        ----------
        contents_batch : List[str]
            DESCRIPTION.

        Returns
        -------
        TYPE
            DESCRIPTION.

        '''
        collected_tags = []
        collected_links = []
        
        for xml_content in contents_batch:
            # Parse content and create HTML content
            new_tags, new_links = self.process_table(xml_content)
            # Split HTML content to HTML tags list
            new_tags = split_tags(new_tags)
            
            # If any tags allready exist
            if collected_tags:
                try: # !!! TODO: перенести в функции и проверить new_tags
                    # check if table in first new tag is an extension of table form last collected tag
                    if table_extend(prev_tag=collected_tags[-1], next_tag=new_tags[0]):
                        # concat tables if so
                        collected_tags[-1] = concat_tables(prev_table=collected_tags[-1], next_table=new_tags[0])
                        new_tags = new_tags[1:]
                        new_links = new_links[1:]
                except IndexError:
                    pass
    
            collected_tags.extend(new_tags)
            collected_links.extend(new_links)
            
        return '\n'.join(collected_tags), collected_links


def split_tags(html_content: str) -> List[str]:
    # Split HTML content into individual tags
    tags = [str(tag) for tag in BeautifulSoup(html_content, features="html.parser").children]
    return tags


def table_extend(prev_tag: str, next_tag: str) -> bool:
    # Check if both tags are <table> tags
    if not (prev_tag.startswith('<table') and next_tag.startswith('<table')):
        return False
    
    # Extract the number of columns from the first row of each table
    # try:
    last_row = BeautifulSoup(prev_tag, features="html.parser").findAll('tr')[0]
    first_row = BeautifulSoup(prev_tag, features="html.parser").findAll('tr')[0]
    
    # Count the number of <td> or <th> tags in the first row
    prev_col_count = max(
        len(last_row.findAll('th')),
        len(last_row.findAll('td'))
    )
    next_col_count = max(
        len(first_row.findAll('th')),
        len(first_row.findAll('td'))
    )
    
    # Tables are considered one if they have the same number of columns
    return prev_col_count == next_col_count


def concat_tables(prev_table: str, next_table: str) -> str:
    # Remove the header from the next table if it matches the header of the previous table
    prev_header = str(BeautifulSoup(prev_table, features="html.parser").findAll('tr')[0])
    next_rows = [
        str(row_tag)
        for row_tag in BeautifulSoup(next_table, features="html.parser").findAll('tr')
    ]
    
    if next_rows[0] == prev_header:
        next_rows = next_rows[1:]
    
    # Combine the tables
    combined_table = prev_table.replace('</table>', '') + ''.join(next_rows) + '</table>'
    return combined_table
      
    
def make_toc_header(text: str, depth: int, max_len: int = 35) -> str:
    """
    Creates a table of contents header.
    
    Args:
        text (str): The text of the header.
        depth (int): The depth of the header.
        max_len (int, optional): The maximum length of the header text.
    
    Returns:
        str: The formatted table of contents header.
    """
    text = '__' * (depth - 1) + text
    if len(text) > max_len:
        text = text[:max_len] + '...'
    return text


def paragraph_style(par: docx.text.paragraph.Paragraph) -> str:
    """
    Retrieves the CSS style for a paragraph.
    
    Args:
        par (docx.text.paragraph.Paragraph): The paragraph to process.
    
    Returns:
        str: The CSS style as a string.
    """
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


def cell_style(cell: docx.table._Cell, borders: dict, c_xml: dict) -> str:
    """
    Retrieves the CSS style for a table cell.
    
    Args:
        cell (docx.table.Cell): The cell to process.
        borders (dict): The borders dictionary.
        c_xml (dict): The XML dictionary for the cell.
    
    Returns:
        str: The CSS style as a string.
    """
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


def int_to_roman(num: int) -> str:
    """
    Converts an integer to a Roman numeral.
    
    Args:
        num (int): The integer to convert.
    
    Returns:
        str: The Roman numeral as a string.
    """
    m = ["", "M", "MM", "MMM"]
    c = ["", "C", "CC", "CCC", "CD", "D",
         "DC", "DCC", "DCCC", "CM "]
    x = ["", "X", "XX", "XXX", "XL", "L",
         "LX", "LXX", "LXXX", "XC"]
    i = ["", "I", "II", "III", "IV", "V",
         "VI", "VII", "VIII", "IX"]
    thousands = m[num // 1000]
    hundreds = c[(num % 1000) // 100]
    tens = x[(num % 100) // 10]
    ones = i[num % 10]
    ans = (thousands + hundreds +
           tens + ones)
    return ans


def docx_to_html(docx_path: str) -> tuple:
    """
    Converts a DOCX document to HTML.
    
    Args:
        docx_path (str): The path to the DOCX file.
    
    Returns:
        tuple: A tuple containing the HTML content and table of contents links.
    """
    doc = docx.Document(docx_path)
    handler = DocHandler(doc)
    html_content = [
        # Document start default headrer
        f'<div id="{handler.depth_anchor[1]}"></div>'
    ]
    toc_links = [
        # Document start default link
        f'<a href="#{handler.depth_anchor[1]}">{make_toc_header("[Начало документа]", 1)}</a><br>'
    ]
    tables_to_process = []
    for content in doc.iter_inner_content():
        if type(content) is docx.text.paragraph.Paragraph:
            if not content.text.strip():
                continue
            if tables_to_process:
                html_table, table_links = handler.process_tables_batch(tables_to_process)
                html_content.append(html_table)
                toc_links.extend(table_links)
                tables_to_process = []
                
            html_paragraph, html_links = handler.process_paragraph(content)
            html_content.extend(html_paragraph)
            toc_links.extend(html_links)
        
        elif type(content) is docx.table.Table:
            tables_to_process.append(content)
        else:
            print(type(content), 'missed')
    # if table was not processed
    if tables_to_process:
        html_table, table_links = handler.process_tables_batch(tables_to_process)
        html_content.append(html_table)
        toc_links.extend(table_links)
    return ''.join(html_content), ''.join(toc_links)
