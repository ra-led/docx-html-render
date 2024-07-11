import json
import logging

from html_to_json import html_to_json

logger = logging.getLogger(__name__)

def test_header_to_json():
    result = html_to_json("<h1>Sample header</h1>")
    logger.warning(result)
    result = json.loads(result)
    assert (len(result), result[0]['content-type'], result[0]['content']) \
            == (1, 'text/title', 'Sample header')

def test_header_hierarchy():
    result = html_to_json("""<h1>Heading 1</h1><h2>Heading 2</h2><h3>Heading 3</h3>""")
    logger.warning(result)
    result = json.loads(result)

    assert len(result) == 3

    assert (result[0]['content-type'], result[0]['content']) == \
            ('text/title', 'Heading 1')

    assert (result[1]['content-type'], result[1]['content'], result[1]['title']) == \
            ('text/subtitle', 'Heading 2', 'Heading 1')

    assert (result[2]['content-type'], result[2]['content'], result[2]['title'], result[2]['sub-title']) == \
            ('text/subtitle', 'Heading 3', 'Heading 1', 'Heading 2')

def test_text_header_linking():
    result = html_to_json("""
        text before headers
        <h1>Section title</h1>
        section description
        <h2>Subsection title</h2>
        subsection text
        <h3>Sub-subsection title</h3>
        sub-subsection text

        <h1>Another section title</h1>
        another section description
    """)
    logger.warning(result)
    result = json.loads(result)

    assert len(result) == 9
    assert (result[0]['content-type'], result[0]['content'].strip(), result[0]['title'], result[0]['sub-title']) == \
            ('text', 'text before headers', '','')

    assert (result[1]['content-type'], result[1]['content'].strip(), result[1]['title'], result[1]['sub-title']) == \
            ('text/title', 'Section title', '','')
    assert (result[2]['content-type'], result[2]['content'].strip(), result[2]['title'], result[2]['sub-title']) == \
            ('text', 'section description', 'Section title','')

    assert (result[3]['content-type'], result[3]['content'].strip(), result[3]['title'], result[3]['sub-title']) == \
            ('text/subtitle', 'Subsection title', 'Section title','')
    assert (result[4]['content-type'], result[4]['content'].strip(), result[4]['title'], result[4]['sub-title']) == \
            ('text', 'subsection text', 'Section title','Subsection title')

    assert (result[5]['content-type'], result[5]['content'].strip(), result[5]['title'], result[5]['sub-title']) == \
            ('text/subtitle', 'Sub-subsection title', 'Section title','Subsection title')
    assert (result[6]['content-type'], result[6]['content'].strip(), result[6]['title'], result[6]['sub-title']) == \
            ('text', 'sub-subsection text', 'Section title','Sub-subsection title')

    assert (result[7]['content-type'], result[7]['content'].strip(), result[7]['title'], result[7]['sub-title']) == \
            ('text/title', 'Another section title', '','')
    assert (result[8]['content-type'], result[8]['content'].strip(), result[8]['title'], result[8]['sub-title']) == \
            ('text', 'another section description', 'Another section title','')

def _assert_table_cell(cell, row, col, text, colspan = 1, rowspan = 1):
    assert (cell['row'], cell['col'], cell['value'], cell['colspan'], cell['rowspan']) == \
            (row, col, text, colspan, rowspan)

def test_table():
    result = html_to_json("""
        <table>
        <tr><td>cell 1 1</td><td>cell 1 2</td></tr>
        <tr><td>cell 2 1</td><td>cell 2 2</td></tr>
        </table>""")
    logger.warning(result)
    result = json.loads(result)

    assert len(result) == 1
    assert result[0]['content-type'] == 'table'
    assert len(result[0]['content']) == 4

    _assert_table_cell(result[0]['content'][0], 1,1, 'cell 1 1')
    _assert_table_cell(result[0]['content'][1], 1,2, 'cell 1 2')
    _assert_table_cell(result[0]['content'][2], 2,1, 'cell 2 1')
    _assert_table_cell(result[0]['content'][3], 2,2, 'cell 2 2')

def test_table_colspan_rowspan():
    result = html_to_json("""
        <table>
        <tr><td>cell 1 1</td><td colspan="2">cell 1 2</td></tr>
        <tr><td rowspan="3">cell 2 1</td><td>cell 2 2</td></tr>
        <tr><td colspan="5" rowspan="3">cell 3 1</td></tr>
        </table>""")
    logger.warning(result)
    result = json.loads(result)

    assert len(result) == 1
    assert result[0]['content-type'] == 'table'
    assert len(result[0]['content']) == 5

    _assert_table_cell(result[0]['content'][0], 1,1, 'cell 1 1')
    _assert_table_cell(result[0]['content'][1], 1,2, 'cell 1 2', colspan=2)
    _assert_table_cell(result[0]['content'][2], 2,1, 'cell 2 1', rowspan=3)
    _assert_table_cell(result[0]['content'][3], 2,2, 'cell 2 2')
    _assert_table_cell(result[0]['content'][4], 3,1, 'cell 3 1', rowspan=3, colspan=5)

