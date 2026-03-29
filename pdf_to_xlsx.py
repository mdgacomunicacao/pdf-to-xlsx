#!/usr/bin/env python3
"""
pdf_to_xlsx.py — Converte PDFs de ensaios GENVCE em XLSX formatado.

Uso:
    python pdf_to_xlsx.py <arquivo.pdf> [saida.xlsx] [--logo logo.png]
"""
import sys, re, argparse
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    sys.exit("Erro: pip install pdfplumber")

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Erro: pip install openpyxl")

# ── Estilos ───────────────────────────────────────────────────────────────────
GREEN_HDR  = '388E3C'
GREEN_EVEN = 'F1F8E9'
WHITE      = 'FFFFFF'

def fill(color): return PatternFill('solid', fgColor=color)

def style_header_cell(cell, text):
    cell.value = text
    cell.font = Font(bold=True, size=10, color='FFFFFF', name='Arial')
    cell.fill = fill(GREEN_HDR)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def style_data_cell(cell, value, zebra=False, align='center'):
    cell.value = value
    cell.font = Font(size=10, name='Arial')
    cell.fill = fill(GREEN_EVEN if zebra else WHITE)
    cell.alignment = Alignment(horizontal=align, vertical='center')

def style_stat_cell(cell, value, bold=False, align='left'):
    cell.value = value
    cell.font = Font(size=10, bold=bold, name='Arial')
    cell.alignment = Alignment(horizontal=align, vertical='center')

def style_meta_cell(cell, value, size=9, bold=False):
    cell.value = value
    cell.font = Font(size=size, bold=bold, name='Arial')
    cell.alignment = Alignment(horizontal='left', vertical='center')

# ── Parser por posição de palavras (sem extract_tables) ───────────────────────
def try_numeric(s):
    s = str(s or '').strip()
    try: return int(s)
    except: pass
    try: return float(s.replace(',', '.'))
    except: pass
    return s if s else None

STAT_KEYS = {'media', 'desviación', 'coeficiente', 'diseño'}
META_RE   = re.compile(r'instituto|iriaf|itacyl|fecha de|siembra|cosecha', re.I)
HDR_RE    = re.compile(r'kg/ha|índice|grupo|nascen|espigad|encamad|rendimi|humedad|altura|daños|frío', re.I)

def group_rows(words, y_tol=4):
    if not words: return []
    rows, cur_y, cur = [], None, []
    for w in sorted(words, key=lambda w: (round(w['top']/y_tol), w['x0'])):
        y = round(w['top']/y_tol)
        if cur_y is None or y != cur_y:
            if cur: rows.append(cur)
            cur, cur_y = [w], y
        else:
            cur.append(w)
    if cur: rows.append(cur)
    return rows

def chunk_row(word_row, gap=12):
    if not word_row: return []
    chunks, cur = [], [word_row[0]]
    for w in word_row[1:]:
        if w['x0'] - cur[-1]['x1'] > gap:
            chunks.append(cur); cur = [w]
        else:
            cur.append(w)
    chunks.append(cur)
    return chunks

def chunk_text(chunk):  return ' '.join(w['text'] for w in chunk)
def chunk_center(chunk): return (chunk[0]['x0'] + chunk[-1]['x1']) / 2

def assign_to_cols(value_chunks, col_centers):
    result = [None] * len(col_centers)
    for vc in value_chunks:
        cx = chunk_center(vc)
        nearest = min(range(len(col_centers)), key=lambda i: abs(col_centers[i] - cx))
        result[nearest] = chunk_text(vc)
    return result

def parse_page(page):
    words = page.extract_words(keep_blank_chars=False, x_tolerance=2, y_tolerance=3)
    word_rows = group_rows(words, y_tol=4)
    result = {'title':'', 'section':'', 'instituto':'', 'dates':[],
              'headers':[], 'col_centers':[], 'rows':[], 'stats':[], 'footnote':''}
    state = 'meta'
    VARIETY_X_MAX = 120

    for word_row in word_rows:
        chunks = chunk_row(word_row, gap=12)
        text   = ' '.join(chunk_text(c) for c in chunks)
        x0     = word_row[0]['x0']

        if not result['title'] and x0 > 100 and 'Ensayo' in text:
            result['title'] = text; continue
        if text.strip().startswith('*') and 'testigo' in text.lower():
            result['footnote'] = text.strip(); continue
        if META_RE.search(text):
            if re.search(r'instituto|iriaf|itacyl', text, re.I): result['instituto'] = text
            else: result['dates'].append(text)
            continue
        if (state == 'meta' and not result['section'] and x0 < 80
                and 8 < len(text) < 90
                and not re.search(r'\d{4}|ensayo en|secanos', text, re.I)
                and not text.isupper()):
            result['section'] = text; continue
        if state == 'data' and any(k in text.lower() for k in STAT_KEYS):
            left  = [c for c in chunks if c[0]['x0'] < 250]
            right = [c for c in chunks if c[0]['x0'] >= 250]
            label = ' '.join(chunk_text(c) for c in left).strip()
            val   = try_numeric(' '.join(chunk_text(c) for c in right).strip())
            result['stats'].append((label or text, val if val is not None else ' '.join(chunk_text(c) for c in right)))
            continue
        if state in ('meta','header') and HDR_RE.search(text) and len(chunks) >= 2:
            data_chunks = [c for c in chunks if c[0]['x0'] > VARIETY_X_MAX - 20]
            if not data_chunks: continue
            result['col_centers'] = [chunk_center(c) for c in data_chunks]
            result['headers'] = ['VARIEDAD'] + [chunk_text(c) for c in data_chunks]
            state = 'data'; continue
        if state == 'data' and result['col_centers'] and x0 < VARIETY_X_MAX:
            variety_chunks = [c for c in chunks if c[-1]['x1'] < result['col_centers'][0] - 5]
            data_chunks    = [c for c in chunks if c[0]['x0'] >= result['col_centers'][0] - 15]
            variety = ' '.join(chunk_text(c) for c in variety_chunks).strip()
            if not variety and chunks:
                variety = chunk_text(chunks[0]); data_chunks = chunks[1:]
            values = assign_to_cols(data_chunks, result['col_centers'])
            row = [variety] + [try_numeric(v) for v in values]
            result['rows'].append(row)

    result['dates_str'] = '  |  '.join(result['dates'])
    return result

def extract_doc_meta(first_page):
    words = first_page.extract_words()
    word_rows = group_rows(words)
    meta = {'title':'', 'instituto':'', 'dates_str':''}
    dates = []
    for wr in word_rows:
        text = ' '.join(w['text'] for w in wr)
        if not meta['title'] and 'Ensayo' in text and wr[0]['x0'] > 100:
            meta['title'] = text
        if META_RE.search(text):
            if re.search(r'instituto|iriaf|itacyl', text, re.I): meta['instituto'] = text
            elif re.search(r'fecha|siembra|cosecha', text, re.I): dates.append(text)
    meta['dates_str'] = '  |  '.join(dates)
    return meta

# ── Escrita do XLSX ───────────────────────────────────────────────────────────
def write_sheet(ws, page_data, doc_meta, logo_path=None):
    ws.row_dimensions[1].height = 22
    for r in [2,3,4]: ws.row_dimensions[r].height = 6

    if logo_path and Path(logo_path).exists():
        img = XLImage(str(logo_path))
        img.width = 120; img.height = 23; img.anchor = 'A1'
        ws.add_image(img)

    section    = page_data.get('section', '')
    base_title = doc_meta.get('title', page_data['title'])
    full_title = f"{base_title} – {section}" if section else base_title

    ws.row_dimensions[5].height = 18
    style_meta_cell(ws['A5'], full_title, size=13, bold=True)
    ws.row_dimensions[6].height = 13
    style_meta_cell(ws['A6'], doc_meta.get('instituto', page_data.get('instituto','')), size=9)
    ws.row_dimensions[7].height = 13
    style_meta_cell(ws['A7'], doc_meta.get('dates_str', page_data.get('dates_str','')), size=9)
    if section:
        ws.row_dimensions[8].height = 16
        style_meta_cell(ws['A8'], section, size=11, bold=True)

    headers = page_data['headers']
    if not headers: return

    HDR_ROW = 9
    ws.row_dimensions[HDR_ROW].height = 30
    for c, h in enumerate(headers, 1):
        style_header_cell(ws.cell(HDR_ROW, c), h)

    for i, row in enumerate(page_data['rows']):
        r = HDR_ROW + 1 + i
        zebra = i % 2 != 0
        ws.row_dimensions[r].height = 15
        while len(row) < len(headers): row.append(None)
        for c, val in enumerate(row[:len(headers)]):
            style_data_cell(ws.cell(r, c+1), val, zebra=zebra, align='left' if c==0 else 'center')

    last_data = HDR_ROW + len(page_data['rows'])

    if page_data['stats']:
        r = last_data + 2
        for label, val in page_data['stats']:
            style_stat_cell(ws.cell(r, 1), label, bold=True)
            style_stat_cell(ws.cell(r, 2), val, align='center')
            r += 1

    if page_data.get('footnote'):
        fn_row = last_data + len(page_data['stats']) + 3
        ws.cell(fn_row, 1).value = page_data['footnote']
        ws.cell(fn_row, 1).font = Font(size=9, name='Arial')

    for c in range(1, len(headers)+1):
        lengths = [len(str(headers[c-1]))]
        for row in page_data['rows']:
            if c-1 < len(row) and row[c-1] is not None:
                lengths.append(len(str(row[c-1])))
        ws.column_dimensions[get_column_letter(c)].width = min(max(max(lengths)+3, 12), 45)

# ── Ponto de entrada ──────────────────────────────────────────────────────────
def convert(pdf_path, xlsx_path, logo_path=None):
    pdf_path  = Path(pdf_path)
    xlsx_path = Path(xlsx_path)
    if not pdf_path.exists():
        sys.exit(f"Arquivo não encontrado: {pdf_path}")

    print(f"📄 Lendo: {pdf_path.name}")
    with pdfplumber.open(pdf_path) as pdf:
        doc_meta   = extract_doc_meta(pdf.pages[0])
        pages_data = [parse_page(p) for p in pdf.pages]

    print(f"   Documento: {doc_meta['title']}")
    wb = Workbook()
    wb.remove(wb.active)

    for i, pdata in enumerate(pages_data):
        name = (pdata['section'] or f'Página {i+1}')[:31]
        ws   = wb.create_sheet(title=name)
        write_sheet(ws, pdata, doc_meta, logo_path=logo_path)
        print(f"   ✓ '{name}': {len(pdata['headers'])} colunas, {len(pdata['rows'])} linhas")

    wb.save(xlsx_path)
    print(f"\n✅ Salvo em: {xlsx_path}")

def main():
    p = argparse.ArgumentParser()
    p.add_argument('pdf')
    p.add_argument('xlsx', nargs='?')
    p.add_argument('--logo')
    args = p.parse_args()
    pdf_path  = Path(args.pdf)
    xlsx_path = Path(args.xlsx) if args.xlsx else pdf_path.with_suffix('.xlsx')
    convert(pdf_path, xlsx_path, logo_path=args.logo)

if __name__ == '__main__':
    main()
