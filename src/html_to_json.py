import json
import logging

from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)
html_header_tags = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'h7']

def _collect_table_content(table_tag):
    trs = table_tag.find_all('tr')
    cells = []

    row = 0
    for tr in trs:
        col = 0
        for el in tr.contents:
            if el.name == 'td' or el.name == 'th':
                rowspan = int(el.get('rowspan', '1'))
                colspan = int(el.get('colspan', '1'))
                cells.append({
                        'row': row+1, 'col': col+1,
                        'rowspan': rowspan, 'colspan': colspan,
                        'value': el.text
                        })
                col += colspan
        row += rowspan
    return cells

        
def _parse_contents_recursive(tag):
    return [_parse_contents_recursive(el) for el in tag.contents]

def _parse_recursive(tag):
    if tag.name is None:
        return [{
            "content-type": "text",
            "content": tag.text, 
            }]
    elif tag.name == 'table':
        return [{
            "content-type": "table",
            "title": tag.get('title', 'Unnamed table'),
            "content": _collect_table_content(tag)
            }]
    elif tag.name in html_header_tags:
        return [{
            "content-type": "text/"+tag.name,
            "content": tag.text,
            }]
    else:
        result = []
        for el in tag.contents:
            result += _parse_recursive(el)
        return result

def _collect_headers(parsed_html):
    headers = filter(lambda el: el['content-type'].startswith('text') and el['content-type'][5:] in html_header_tags, parsed_html)
    headers = map(lambda el: el['content-type'][5:], headers)
    return sorted(set(headers))

def _build_initial_context(parsed_html):
    headers = _collect_headers(parsed_html)
    logger.debug('headers: '+json.dumps(headers))
    context = {
        "title_tag": headers[0] if len(headers) > 0 else "",
        "subtitle_tag": headers[1] if len(headers) > 1 else "",
        "subsubtitle_tag": headers[2] if len(headers) > 2 else "",
    }
    logger.debug('initial context: '+json.dumps(context))
    return context

def _update_header(el, context):
    logger.debug("el: "+json.dumps(el)+"context: "+json.dumps(context))
    tag_name = el['content-type'][5:]
    el['title'] = ''
    el['sub-title'] = ''

    if tag_name == context['title_tag']:
        el['content-type'] = 'text/title'
        context['last-title'] = el['content']
        context['last-subtitle'] = ''
        context['last-subsubtitle'] = ''
    elif tag_name == context['subtitle_tag']:
        el['content-type'] = 'text/subtitle'
        el['title'] = context['last-title']
        context['last-subtitle'] = el['content']
        context['last-subsubtitle'] = ''
    elif tag_name == context['subsubtitle_tag']:
        el['content-type'] = 'text/subtitle'
        el['title'] = context['last-title']
        el['sub-title'] = context['last-subtitle']
        context['last-subsubtitle'] = el['content']

def _update_text(el, context):
    el['title'] = context.get('last-title', '')
    el['sub-title'] = context.get('last-subsubtitle') if context.get('last-subsubtitle', '') != '' \
        else context.get('last-subtitle', '')

def _contextual_parser(parsed_html):
    context = _build_initial_context(parsed_html)
    for el in parsed_html:
        if el['content-type'].startswith('text') and el['content-type'][5:] in html_header_tags:
            _update_header(el, context)
        elif el['content-type'] == 'text':
            _update_text(el, context)
    return parsed_html


def html_to_json(html):
    with open('test', 'w') as f:
        f.write(html)

    soup = BeautifulSoup(html, features="lxml")
    initial_json = _parse_recursive(soup.contents[0])
    json_with_context = _contextual_parser(initial_json)
    logger.debug(json.dumps(json_with_context, indent=4, ensure_ascii=False))

    return json.dumps(json_with_context)
