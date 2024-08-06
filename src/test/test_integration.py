# pytest --host=158.160.39.120 --port=8000 --warning_threshold=0.9 --error_threshold=0.1
import json
from pathlib import Path
import re
import requests
import pandas as pd
import pytest
from loguru import logger

def parse_num(text):
    if re.match(r'\d', text.replace('_', '').split()[0][-1]):
        return text.replace('_', '').split()[0]

def parse_element(element):
    table = element['content-type'] == 'table'
    num_prefix = None
    header = ''
    if element['content-type'] == 'text/title':
        header = element['title']
        num_prefix = parse_num(header)
    elif element['content-type'] == 'text/subtitle':
        header = element['sub-title']
        num_prefix = parse_num(header)
    return {
        'depth': len(re.sub(r'\.$', '', num_prefix).split('.')) if num_prefix else 0,
        'header': header,
        'table': table,
        'num_prefix': num_prefix
    }

labeled_docs = [
    (docx_path, docx_path.with_suffix('.tsv'))
    for docx_path in Path('test/labeled_docs').rglob('*.docx')
]

@pytest.mark.parametrize("docx_path, labels_path", labeled_docs)
def test_doc_structure(docx_path, labels_path, host, port, warning_threshold, error_threshold):
    url = f"http://{host}:{port}/"
    with open(docx_path, 'rb') as file:
        files = {'file': file}
        response = requests.post(url, files=files)
    assert response.status_code == 200, f'Response != 200 for {docx_path}'
    elements = json.loads(response.json())
    df_elements = pd.DataFrame([parse_element(el) for el in elements])
    df_labeled_elements = pd.read_csv(labels_path, sep='\t')

    # Numeration
    true_nums = set(df_labeled_elements['num_prefix'].dropna())
    found_nums = set(df_elements['num_prefix'].dropna())
    num_recall = len(true_nums.intersection(found_nums)) / len(true_nums)
    num_precision = len(true_nums.intersection(found_nums)) / len(found_nums)

    # Tables
    true_tables_count = df_labeled_elements['table'].sum()
    found_tables_count = df_elements['table'].sum()
    tables_error = abs(true_tables_count - found_tables_count) / true_tables_count

    # Logging and Assertions
    logger.info(f'Numeration P: {num_recall:.2f}; {docx_path.name}')
    logger.info(f'Numeration R: {num_precision:.2f}; {docx_path.name}')
    logger.info(f'Tables Error: {tables_error:.2f} {docx_path.name}')

    if num_recall < error_threshold or num_precision < error_threshold or tables_error > error_threshold:
        pytest.fail(f'Metrics below error threshold for {docx_path.name}')
    elif num_recall < warning_threshold or num_precision < warning_threshold or tables_error > warning_threshold:
        logger.warning(f'Metrics below warning threshold for {docx_path.name}')
