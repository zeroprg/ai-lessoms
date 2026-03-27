import re
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import sys
import os

def create_styles(doc):
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.space_before = Pt(2)

    for i in range(1, 5):
        h = doc.styles[f'Heading {i}']
        h.font.name = 'Calibri'
        h.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)
        if i == 1:
            h.font.size = Pt(22)
            h.paragraph_format.space_before = Pt(12)
            h.paragraph_format.space_after = Pt(8)
        elif i == 2:
            h.font.size = Pt(16)
            h.paragraph_format.space_before = Pt(14)
            h.paragraph_format.space_after = Pt(6)
        elif i == 3:
            h.font.size = Pt(13)
            h.paragraph_format.space_before = Pt(10)
            h.paragraph_format.space_after = Pt(4)
        elif i == 4:
            h.font.size = Pt(11)
            h.paragraph_format.space_before = Pt(8)
            h.paragraph_format.space_after = Pt(4)

    code_style = doc.styles.add_style('CodeBlock', 1)
    code_style.font.name = 'Consolas'
    code_style.font.size = Pt(9)
    code_style.font.color.rgb = RGBColor(0x1E, 0x1E, 0x1E)
    code_style.paragraph_format.space_before = Pt(4)
    code_style.paragraph_format.space_after = Pt(4)
    code_style.paragraph_format.left_indent = Cm(0.5)

    quote_style = doc.styles.add_style('BlockQuote', 1)
    quote_style.font.name = 'Calibri'
    quote_style.font.size = Pt(11)
    quote_style.font.italic = True
    quote_style.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    quote_style.paragraph_format.left_indent = Cm(1.0)
    quote_style.paragraph_format.space_before = Pt(4)
    quote_style.paragraph_format.space_after = Pt(4)


def add_formatted_text(paragraph, text):
    """Parse inline markdown (bold, italic, code, emoji) and add runs."""
    parts = re.split(r'(`[^`]+`|\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('***') and part.endswith('***'):
            run = paragraph.add_run(part[3:-3])
            run.bold = True
            run.italic = True
        elif part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        elif part.startswith('`') and part.endswith('`'):
            run = paragraph.add_run(part[1:-1])
            run.font.name = 'Consolas'
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(0xC7, 0x25, 0x4E)
        else:
            paragraph.add_run(part)


def add_table(doc, rows):
    """Add a markdown table to the document."""
    if len(rows) < 2:
        return

    headers = [cell.strip() for cell in rows[0].strip('|').split('|')]
    data_rows = []
    for row in rows[2:]:
        cells = [cell.strip() for cell in row.strip('|').split('|')]
        if cells:
            data_rows.append(cells)

    num_cols = len(headers)
    table = doc.add_table(rows=1 + len(data_rows), cols=num_cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        run = p.add_run(header)
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = 'Calibri'
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E2F3"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    for r_idx, data_row in enumerate(data_rows):
        for c_idx in range(min(len(data_row), num_cols)):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            add_formatted_text(p, data_row[c_idx])
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.name = 'Calibri'

    doc.add_paragraph()


def convert_md_to_docx(md_path, docx_path):
    for enc in ('utf-8-sig', 'utf-8', 'cp1251', 'cp1252'):
        try:
            with open(md_path, 'r', encoding=enc) as f:
                lines = f.readlines()
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    else:
        with open(md_path, 'r', encoding='utf-8', errors='replace') as f:
            lines = f.readlines()

    doc = Document()

    section = doc.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    create_styles(doc)

    i = 0
    while i < len(lines):
        line = lines[i].rstrip('\n')

        if not line.strip():
            i += 1
            continue

        if line.strip().startswith('---') and all(c in '-' for c in line.strip()):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            pPr = p._p.get_or_add_pPr()
            pBdr = parse_xml(
                f'<w:pBdr {nsdecls("w")}>'
                f'  <w:bottom w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/>'
                f'</w:pBdr>'
            )
            pPr.append(pBdr)
            i += 1
            continue

        if line.strip().startswith('```'):
            lang = line.strip()[3:].strip()
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].rstrip('\n').strip().startswith('```'):
                code_lines.append(lines[i].rstrip('\n'))
                i += 1
            i += 1

            if lang:
                p = doc.add_paragraph()
                run = p.add_run(f'[{lang}]')
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
                run.font.name = 'Consolas'
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.space_before = Pt(6)

            for code_line in code_lines:
                p = doc.add_paragraph(style='CodeBlock')
                run = p.add_run(code_line if code_line else ' ')
                run.font.name = 'Consolas'
                run.font.size = Pt(9)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)

            continue

        if line.strip().startswith('|') and '|' in line.strip()[1:]:
            table_rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_rows.append(lines[i].strip())
                i += 1
            add_table(doc, table_rows)
            continue

        heading_match = re.match(r'^(#{1,4})\s+(.*)', line)
        if heading_match:
            level = len(heading_match.group(1))
            text = heading_match.group(2).strip()
            p = doc.add_heading(level=level)
            add_formatted_text(p, text)
            i += 1
            continue

        if line.strip().startswith('>'):
            quote_lines = []
            while i < len(lines) and lines[i].strip().startswith('>'):
                quote_text = re.sub(r'^>\s?', '', lines[i].strip())
                quote_lines.append(quote_text)
                i += 1
            full_quote = '\n'.join(quote_lines)
            for ql in full_quote.split('\n'):
                if ql.strip():
                    p = doc.add_paragraph(style='BlockQuote')
                    add_formatted_text(p, ql)
            continue

        list_match = re.match(r'^(\s*)([-*]|\d+\.)\s+(.*)', line)
        if list_match:
            indent_level = len(list_match.group(1)) // 2
            marker = list_match.group(2)
            text = list_match.group(3)

            if re.match(r'\d+\.', marker):
                p = doc.add_paragraph(style='List Number')
            else:
                p = doc.add_paragraph(style='List Bullet')

            p.paragraph_format.left_indent = Cm(1.0 + indent_level * 0.6)
            add_formatted_text(p, text)
            i += 1
            continue

        p = doc.add_paragraph()
        add_formatted_text(p, line)
        i += 1

    doc.save(docx_path)
    return docx_path


if __name__ == '__main__':
    import glob

    pattern = sys.argv[1] if len(sys.argv) > 1 else None
    if pattern:
        matches = [m for m in glob.glob(pattern) if m.endswith('.md')]
        if matches:
            md_file = matches[0]
        else:
            md_file = pattern
    else:
        files = glob.glob("*.md")
        lesson_files = [f for f in files if "44" in f]
        md_file = lesson_files[0] if lesson_files else files[0]

    base = os.path.splitext(md_file)[0]
    docx_file = base + '.docx'

    print(f"Input exists: {os.path.exists(md_file)}", file=sys.stderr)
    print(f"Input size: {os.path.getsize(md_file) if os.path.exists(md_file) else 'N/A'}", file=sys.stderr)
    try:
        convert_md_to_docx(md_file, docx_file)
    except PermissionError:
        docx_file = base + ' (1).docx'
        convert_md_to_docx(md_file, docx_file)
    print(f"Output exists: {os.path.exists(docx_file)}", file=sys.stderr)
    print("OK")
