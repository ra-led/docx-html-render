import re
import statistics
import string
from typing import Union
import uuid
import docx
import xmltodict
from .core import ParHandler, Node
from .ml import BERTTextClassifier


class NumberingDB:
    """
    Handles numbering and styles in a DOCX document.
    """     
    def __init__(self, doc: docx.Document, appendix_header_length: int = 40,
                 default_levels: int = 9, default_font: int = 12,
                 norm_numeration_model: str = 'model_dir/num_clf',
                 norm_heading_model: str = 'model_dir/word_clf'):
        """
        Initializes the NumberingDB with a DOCX document.
        
        Args:
            doc (docx.Document): The DOCX document to process.
        """
        self.doc = doc
        self.appendix_header_length = appendix_header_length
        self.default_levels = default_levels
        self.default_font = default_font
        try:
            self.num_xml = xmltodict.parse(
                self.doc.part.numbering_part.element.xml,
                process_namespaces=False
            )
        except NotImplementedError:
            self.num_xml = {}
        self.init_default_abstract(default_levels)
        self.get_levels_abstracts()
        self.link_nums_to_abstracts()
        self.link_styles_to_abstracts()
        self.init_numbering_increment()
        
        self.font_size = []
        
        self.norm_numeration_clf = BERTTextClassifier(norm_numeration_model)
        self.norm_heading_clf = BERTTextClassifier(norm_numeration_model)
        
        self.stop_symbs = [')', ':', '-', '–', '—', '−']

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
            self.levels[absId] = self.default_abstract
            self.increment[absId] = self.default_increment
            self.nums_to_abstarct[numId] = absId
            return absId
        return None
    
    def check_heading_style(self, par: ParHandler) -> bool:
        """
        Checks if a paragraph has a heading style.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to check.
        
        Returns:
            bool: True if the paragraph has a heading style, False otherwise.
        """
        if re.findall('^таблица', par.ctext.lower()):
            return False
        if re.findall('^рисунок', par.ctext.lower()):
            return False
        par_font_size = par.font_size or self.default_font
        if par.bold or par_font_size > self.get_regular_font_size():
            return True
        else:
            return False
    
    def count_builtin(self, absId: str, level: int) -> Node:
        """
        Counts the built-in numbering for a given abstract number ID and level.
        
        Args:
            absId (str): The abstract number ID.
            level (int): The level of the numbering.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        self.inc_levels(absId, level)
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
            # Current level num respect to numbering format
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
            # Inject current level num to num prefix template
            if f'%{lvl_i + 1}' in num_prefix:
                depth += 1
                num_prefix = re.sub(f'%{lvl_i + 1}', str(num), num_prefix)
        return Node(num_prefix, depth, absId)
    
    def numrize_by_meta(self, par: ParHandler) -> ParHandler:
        """
        Processes numbering by metadata.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        try:
            numId = par.xml['w:p']['w:pPr']['w:numPr']['w:numId']['@w:val']
            level = int(par.xml['w:p']['w:pPr']['w:numPr']['w:ilvl']['@w:val'])
        except KeyError:
            return par
        absId = self.get_abs_id(numId=numId)
        node = self.count_builtin(absId, level)
        if not self.check_heading_style(par) and node.depth == 1:
            node.depth = 0
        if self.stop_symbs_in_prefix(node.num_prefix) or self.stop_symbs_in_start(par.ctext):
            node.depth = 0
        par.node = node
        return par
    
    def numrize_by_style(self, par: ParHandler) -> ParHandler:
        """
        Processes numbering by style.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        style_abs = self.get_abs_id(styleId=par.style_id) or self.get_abs_id(styleId=par.base_style_id)
        if style_abs:
            par.node = self.count_builtin(style_abs['absId'], style_abs['lvl'])
        return par
    
    def numerize_by_text(self, par: ParHandler) -> ParHandler:
        """
        Detect numbering in text prefix.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        num_prefix, depth, cleaned_text = find_manual_numbering(par.ctext, self.default_levels)
        if depth:
            if self.stop_symbs_in_start(cleaned_text):
                return par
            if not self.check_heading_style(par) and depth == 1:
                return par
            if not self.norm_numeration_clf(par.ctext):
                return par
            par.node = Node(num_prefix, depth, 'REGEX')
        return par

    def numerize_by_heading(self, par: ParHandler) -> ParHandler:
        """
        Processes numbering by heading.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        if not par.style_name:
            return par
        if not re.search(r'Heading (\d+)', par.style_name) or par.style_name != 'Title':
            return par
        if not self.check_heading_style(par):
            return par
        if not self.norm_heading_clf(par.ctext):
            return par
        par.node = Node(par.ctext, 1, 'HEADING')
        return par
        
    def numerize_by_appendix(self, par: ParHandler) -> ParHandler:
        """
        Processes numbering by detecting appendix header.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        text = par.ctext.split('\n')[0]
        match = re.search(r'^приложение', text.lower())
        if match and len(text) < self.appendix_header_length:
            par.node = Node(text, 1, 'APPENDIX')
        return par
    
    def numerize(self, par: ParHandler) -> ParHandler:
        """
        Processes numbering for a paragraph.
        
        Args:
            par (docx.text.paragraph.Paragraph): The paragraph to process.
        
        Returns:
            tuple: A tuple containing the numbering prefix, depth, and source.
        """
        # Update font size stat
        self.font_size.append(par.font_size or self.default_font)
        # Numeraize paragraph
        numerize_prioritet = [
            self.numrize_by_meta,
            self.numrize_by_style,
            self.numerize_by_text,
            self.numerize_by_heading,
            self.numerize_by_appendix
        ]
        for method in numerize_prioritet:
            par = method(par)
            if par.node.num_prefix:
                break
        return par
    
    def get_levels_abstracts(self):
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
            
    def link_nums_to_abstracts(self):
        try:
            nums_abs = self.num_xml['w:numbering']['w:num']
            if type(nums_abs) is list:
                self.nums_to_abstarct = {
                    x['@w:numId']: x['w:abstractNumId']['@w:val']
                    for x in nums_abs
                }
            elif type(nums_abs) is dict:
                self.nums_to_abstarct = {
                    nums_abs['@w:numId']: nums_abs['w:abstractNumId']['@w:val']
                }
        except KeyError:
            self.nums_to_abstarct = {}
            
    def link_styles_to_abstracts(self):
        self.styles_xml = xmltodict.parse(
            self.doc.part.styles.element.xml,
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
                    
    def init_numbering_increment(self):
        self.increment = {
            k: {i: 0 for i in range(len(v))}
            for k, v in self.levels.items()
        }
        
    def init_default_abstract(self, n_levels: int):
        self.default_abstract = [
            {
                '@w:ilvl': str(i),
                'w:start': {'@w:val': '1'},
                'w:numFmt': {'@w:val': 'decimal'},
                'w:lvlText': {'@w:val': 'default ' + ''.join([f'%{j + 1}' for j in range(n_levels)])}
            }
            for i in range(n_levels)
        ]
        self.default_increment = {i: 0 for i in range(n_levels)}
        
    def inc_levels(self, absId: str, level: int):
        self.increment[absId][level] += 1
        for lvl_i in self.increment[absId]:
            if lvl_i > level:
                self.increment[absId][lvl_i] = 0
        
    def get_regular_font_size(self):
        return statistics.median(self.font_size) if self.font_size else self.default_font
    
    def stop_symbs_in_prefix(self, num_prefix: str):
        return any([symb in num_prefix for symb in self.stop_symbs])
    
    def stop_symbs_in_start(self, text: str):
        try:
            return any([symb == text[0] for symb in self.stop_symbs])
        except IndexError:
            return True


def find_manual_numbering(text: str, max_levels: int) -> tuple:
    depth = 0
    num_prefix = ''
    
    letter_pattern = r'^(\w\.)\d'
    match = re.findall(letter_pattern, text)
    if match:
        text = re.sub(r'^\w\.', '', text)
        num_prefix += match[0]
        depth += 1
    
    numbering_pattern = r'^\d+\.'
    for _ in range(max_levels):
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
    return num_prefix, depth, text.strip()

            
def int_to_roman(num: int) -> str:
    """
    Converts an integer to a Roman numeral.
    
    Args:
        num (int): The integer to convert.
    
    Returns:
        str: The Roman numeral as a string.
    """
    m = ["", "M", "MM", "MMM"]
    c = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM "]
    x = ["", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC"]
    i = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"]
    thousands = m[num // 1000]
    hundreds = c[(num % 1000) // 100]
    tens = x[(num % 100) // 10]
    ones = i[num % 10]
    ans = thousands + hundreds + tens + ones
    return ans
