#!/usr/bin/env python3
"""Render Word, resolve TOC page numbers, and verify the resulting pagination."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys

from pypdf import PdfReader
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--docx', type=Path, default=ROOT / 'drafts/working-draft.docx')
args = parser.parse_args()
directory = ROOT / 'build/render-current'
directory.mkdir(exist_ok=True)
env = os.environ.copy()
fonts = Path('/workspace/shared/render-fonts/fontconfig.xml')
if fonts.exists():
    env['FONTCONFIG_FILE'] = str(fonts)
    env['XDG_CACHE_HOME'] = '/workspace/shared/render-cache'
renderer = shutil.which('soffice')
if not renderer:
    raise SystemExit('LibreOffice is required to refresh and check contents pages.')
pdf = directory / (args.docx.stem + '.pdf')

def render():
    if pdf.exists():
        pdf.unlink()
    result = subprocess.run([renderer, '--headless', '--convert-to', 'pdf',
                             '--outdir', str(directory), str(args.docx)],
                            env=env, capture_output=True, text=True, timeout=60)
    if result.returncode or not pdf.exists():
        raise RuntimeError(result.stdout + result.stderr)

render()
for attempt in range(3):
    subprocess.run([sys.executable, str(ROOT / 'tools/assemble-thesis.py'),
                    '--output', str(args.docx), '--pagination-pdf', str(pdf)], check=True)
    expected = json.loads((ROOT / 'build/thesis-contents.json').read_text())
    render()
    pdf_pages = PdfReader(pdf).pages
    pages = [p.extract_text() for p in pdf_pages]
    norm = lambda s: re.sub(r'\s+', '', s).casefold()
    if all(norm(e['heading']) in norm(pages[e['page'] - 1]) for e in expected if e['written']):
        break
else:
    raise RuntimeError('TOC pagination failed to converge')
assert len(pages) > 3
assert 'СОДЕРЖАНИЕ' in pages[1] and 'СОДЕРЖАНИЕ' not in pages[0]
assert 'Козырева Никиты Алексеевича' in pages[0] and '2026' in pages[0]
glossary_page = expected[0]['page']
assert 'ПЕРЕЧЕНЬ СОКРАЩЕНИЙ' in pages[glossary_page - 1]
assert 'ВВЕДЕНИЕ' in pages[expected[1]['page'] - 1]
contents_text = norm(' '.join(pages[1:glossary_page - 1]))
assert not re.search(r'\n1\s*$', pages[0]), 'Visible number on title page'
for number, text in enumerate(pages[1:], 2):
    footer_numbers = []
    def visit(value, cm, tm, font, size):
        if value.strip() == str(number) and 0 <= tm[5] <= 72:
            footer_numbers.append((tm[4], tm[5]))
    pdf_pages[number - 1].extract_text(visitor_text=visit)
    assert footer_numbers, 'Missing footer page number: ' + str(number)
    assert len(text.strip()) > 50, 'Empty page: ' + str(number)
for entry in expected:
    if not entry['written']:
        assert entry['page'] is None and norm(entry['label']) in contents_text
        continue
    assert entry['page'] > 0
    pattern = re.escape(norm(entry['label'])) + r'\.*' + str(entry['page'])
    assert re.search(pattern, contents_text), 'Incorrect contents entry: ' + entry['label']
table_checks = []
for table in Document(args.docx).tables:
    headers = [cell.text for cell in table.rows[0].cells]
    locations = [[] for _ in headers]
    for page_number, page in enumerate(pdf_pages, 1):
        def visit_header(value, cm, tm, font, size):
            clean = norm(value)
            for index, header in enumerate(headers):
                target = norm(header)
                matches = clean == target if len(target) <= 15 else clean.startswith(target[:15])
                if clean and matches:
                    locations[index].append((page_number, round(tm[4], 2), round(tm[5], 2)))
        page.extract_text(visitor_text=visit_header)
    assert all(locations), 'Missing rendered table headers: ' + str(headers)
    first = [column[0] for column in locations]
    assert len({point[0] for point in first}) == 1, 'Table headers split across pages'
    assert max(point[2] for point in first) - min(point[2] for point in first) < 1, 'Table headers are not on one row'
    assert all(b[1] > a[1] + 30 for a, b in zip(first, first[1:])), 'Table columns not rendered separately'
    table_page = norm(pages[first[0][0] - 1])
    assert all(norm(cell.text) in table_page for row in table.rows for cell in row.cells), 'Compact table continues without a continuation caption'
    table_checks.append({'headers': headers, 'page': first[0][0], 'coordinates': first,
                         'all_cells_on_one_page': True})
published_pdf = args.docx.with_suffix('.pdf')
shutil.copyfile(pdf, published_pdf)
font_result = subprocess.run(['pdffonts', str(pdf)], capture_output=True, text=True, check=True)
report = {
    'status': 'passed', 'renderer': subprocess.check_output([renderer, '--version'], env=env, text=True).strip(),
    'pages': len(pages), 'contents_page': 2, 'title_page_number_visible': False,
    'contents_page_count': glossary_page - 2,
    'planned_entries': len(expected), 'written_entries': sum(e['written'] for e in expected),
    'page_numbers': '2 through ' + str(len(pages)), 'contents': expected,
    'table_headers': table_checks,
    'font_report': font_result.stdout.strip(),
    'word_sha256': hashlib.sha256(args.docx.read_bytes()).hexdigest(),
    'pdf_sha256': hashlib.sha256(published_pdf.read_bytes()).hexdigest(),
    'original_title_template_sha256': hashlib.sha256((ROOT / 'requirements/originals/title-template.docx').read_bytes()).hexdigest(),
    'limits': 'Pagination verified in the named LibreOffice renderer. Word may repaginate when edited or opened with different fonts.',
}
(ROOT / 'build/thesis-pagination-checks.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
print('Pagination passed:', len(pages), 'pages;', published_pdf)
