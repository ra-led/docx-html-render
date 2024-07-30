import html
from .core import ParHandler, TableView, DocRoot
from .ooxml import DocHandler


class DocHTML:
    def __init__(self):
        root = DocRoot()
        self.html_content = [
            # Document start default headrer
            f'<div id="{root.node._id}"></div>'
        ]
        self.toc_links = [
            # Document start default link
            f'<a href="#{root.node._id}">{make_toc_header(root.node.num_prefix, 1)}</a><br>'
        ]
        
    def paragraph_html(self, par: ParHandler):
        classes = ' '.join(par.node.parents.values())
        css = paragraph_style(par)
        par_text = html.escape(par.get_full_text())
        if par.node.depth > 0:
            tag = f'h{par.node.depth}'
            link_text = make_toc_header(par.get_full_text(), par.node.depth)
            anchor = par.node._id
            self.toc_links.append(
                f'<a href="#{anchor}">{link_text}</a><br>'
            )
            self.html_content.append(
                f'<div style="{css}"><{tag} id="{anchor}" class="{classes}">{par_text}</{tag}></div>'
            )
        else:
            self.html_content.append(
                f'<div style="{css}"><p class="{classes}">{par_text}</p></div>'
            )
            
    def table_html(self, table: TableView):
        classes = ' '.join(table.node.parents.values())
        anchor = table.node._id
        title = table.node.num_prefix
        link_text = make_toc_header(title, table.node.depth)
        self.toc_links.append(
            f'<a href="#{anchor}">{link_text}</a><br>'
        )
        html_table = f'<table id="{anchor}" class="w3-table w3-hoverable {classes}" title="{title}">'
        for i, row in enumerate(table.rows):
            html_table += '<tr>'
            for cell in row:
                if i == 0:
                    cell_tag = 'th'
                else:
                    cell_tag = 'td'
                cell_text = '<br>'.join([t for t in cell.ctext.split('\n')])
                html_table += f'<{cell_tag}>{cell_text}</{cell_tag}>'
            html_table += '</tr>'
        html_table += '</table>'
        self.html_content.append(html_table)
        
    def get_html(self, handler: DocHandler) -> tuple:
        if not handler.processed:
            handler.process()
        for content in handler.processed_content:
            if type(content) is ParHandler:
                self.paragraph_html(content)
            elif type(content) is TableView:
                self.table_html(content)
        return ''.join(self.html_content), ''.join(self.toc_links)
            

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
    return html.escape(text)


def paragraph_style(par: ParHandler) -> str:
    """
    Retrieves the CSS style for a paragraph.
    
    Args:
        par (docx.text.paragraph.Paragraph): The paragraph to process.
    
    Returns:
        str: The CSS style as a string.
    """
    css = ''
    try:
        css += 'text-align: {};'.format(par.par.alignment.name.lower())
    except (KeyError, AttributeError):
        pass
    if par.bold:
        css += 'font-weight: bold;'
    return css
    