import os, re, sys, io, json, math, zipfile, tempfile, traceback, subprocess, shutil
from pathlib import Path
from collections import Counter

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

APP_TITLE = "SAP EWA Action Item Extractor"
SAP_NOTE_URL = "https://me.sap.com/notes/{note}"

HEAD_RE = re.compile(r'\b(?:SAP\s+)?Note\s*[:#]?\s*(\d{5,8})\b', re.I)
SNOTE_RE = re.compile(r'\b(?:SNOTE|SAP\s*Note)\s*(?:#|:)?\s*(\d{5,8})\b', re.I)
NUMBER_RE = re.compile(r'\b(\d{5,8})\b')


def clean(s):
    if s is None:
        return ""
    s = str(s).replace('\u00ad', '').replace('\xa0', ' ')
    s = re.sub(r'[ \t]+', ' ', s)
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()


def normalize(s):
    s = clean(s).lower()
    s = s.replace('*', '').replace('–', '-').replace('—', '-')
    s = re.sub(r'[^a-z0-9]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def extract_note_numbers(text):
    text = clean(text)
    nums = []
    for m in SNOTE_RE.finditer(text):
        nums.append(m.group(1))
    # Also capture phrases such as "SAP Notes 91488 and 2253047" and "Note 1650394".
    for m in re.finditer(r'(?i)(?:SAP\s+Notes?|Notes?|SNOTE)\s*(?:[:#])?\s*([0-9]{5,8}(?:\s*(?:,|and|&)\s*[0-9]{5,8})*)', text):
        for n in re.findall(r'\d{5,8}', m.group(1)):
            nums.append(n)
    # Avoid treating years and numeric IDs as notes unless introduced by Note/SAP Note/SNOTE.
    seen = []
    for n in nums:
        if n not in seen:
            seen.append(n)
    return seen


def image_rating(cell, document):
    ids = [x.get(qn('r:embed')) for x in cell._tc.xpath('.//a:blip')]
    if not ids:
        return ""
    part = document.part.related_parts.get(ids[0])
    if part is None:
        return ""
    try:
        im = Image.open(io.BytesIO(part.blob)).convert('RGBA')
        cnt = Counter((r, g, b) for r, g, b, a in im.getdata() if a > 100)
        red = sum(v for (r, g, b), v in cnt.items() if r > 180 and g < 120 and b < 120)
        yellow = sum(v for (r, g, b), v in cnt.items() if r > 180 and g > 150 and b < 120)
        green = sum(v for (r, g, b), v in cnt.items() if g > 140 and r < 120 and b < 120)
        if red > yellow and red > green and red >= 5:
            return 'Red'
        if yellow > green and yellow >= 5:
            return 'Yellow'
        if green >= 5:
            return 'Green'
    except Exception:
        pass
    return ''


def ordered_blocks(doc):
    blocks = []
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith('}p'):
            blocks.append(('p', Paragraph(child, doc)))
        elif child.tag.endswith('}tbl'):
            table = next((t for t in doc.tables if t._tbl is child), None)
            if table is not None:
                blocks.append(('t', table))
    return blocks


def build_heading_index(blocks):
    headings = []
    number = 0
    for idx, (kind, obj) in enumerate(blocks):
        if kind == 'p' and obj.style.name.startswith('Heading') and clean(obj.text) and 'AUTONUMLGL' in obj._p.xml:
            number += 1
            level = int(re.search(r'(\d+)$', obj.style.name).group(1))
            headings.append({'block': idx, 'level': level, 'number': number, 'title': clean(obj.text)})
    return headings


def heading_for_block(headings, block_idx):
    prev = None
    for h in headings:
        if h['block'] > block_idx:
            break
        prev = h
    return prev


def find_next_heading(headings, block_idx, text):
    target = normalize(text)
    if not target:
        return None
    best = None
    best_score = 0
    for h in headings:
        if h['block'] <= block_idx:
            continue
        ht = normalize(h['title'])
        if target == ht:
            return h
        if target in ht or ht in target:
            score = min(len(target), len(ht)) / max(len(target), len(ht))
            if score > best_score:
                best_score, best = score, h
        else:
            a, b = set(target.split()), set(ht.split())
            if a and b:
                score = len(a & b) / max(1, len(a))
                if score > best_score and score >= 0.65:
                    best_score, best = score, h
    return best


def block_text(blocks, start_idx, end_idx):
    parts = []
    for kind, obj in blocks[start_idx:end_idx]:
        if kind == 'p':
            t = clean(obj.text)
            if t:
                parts.append(t)
        else:
            for row in obj.rows:
                vals = [clean(c.text) for c in row.cells]
                line = ' | '.join(v for v in vals if v)
                if line:
                    parts.append(line)
    return '\n'.join(parts)


def section_content(blocks, headings, h):
    start = h['block'] + 1
    end = len(blocks)
    for nxt in headings:
        if nxt['block'] <= h['block']:
            continue
        if nxt['level'] <= h['level']:
            end = nxt['block']
            break
    text = block_text(blocks, start, end)
    return text


def extract_recommendation(text):
    text = clean(text)
    if not text:
        return ''
    m = re.search(r'(?is)(?:^|\n)\s*Recommendation\s*:\s*(.*?)(?=\n\s*(?:Background|Implementation|Further Information|Note\s*:)|$)', text)
    if m:
        rec = clean(m.group(1))
        return rec
    # Some EWA text uses Recommendation embedded in a paragraph.
    m = re.search(r'(?is)Recommendation\s*:\s*(.*)$', text)
    if m:
        return clean(m.group(1))
    return ''


def extract_analysis(text):
    text = clean(text)
    if not text:
        return ''
    # Prefer paragraphs before Recommendation, excluding table-like metadata.
    before = text.split('Recommendation:', 1)[0].strip() if 'Recommendation:' in text else text
    paras = [clean(x) for x in re.split(r'\n+', before) if clean(x)]
    useful = []
    for p in paras:
        if len(p) < 20:
            continue
        if re.match(r'^(Rating|Check|System ID|Date|Instance|Host|Area|Item|Value|Description|Priority)\b', p, re.I):
            continue
        useful.append(p)
    if not useful:
        return ''
    # Keep the first few meaningful finding paragraphs; source wording is preserved.
    out = ' '.join(useful[:3])
    return out[:2200]


def make_directive(heading_title, check, content):
    analysis = extract_analysis(content)
    rec = extract_recommendation(content)
    parts = [f'Topic: {check}']
    if analysis:
        parts.append('Analysis: ' + analysis)
    if rec:
        parts.append('Recommendation: ' + rec)
    if not parts:
        # Do not invent a recommendation when the source has none.
        parts.append('Check: ' + check)
    return '\n'.join(parts)


def parse_docx(path, include_green=False):
    doc = Document(path)
    blocks = ordered_blocks(doc)
    headings = build_heading_index(blocks)
    heading_by_block = {h['block']: h for h in headings}

    # Identify rating tables and their rows. A row's matching heading is the next heading with the same check/subtopic title.
    candidates = []
    for idx, (kind, obj) in enumerate(blocks):
        if kind != 't' or not obj.rows:
            continue
        header = [normalize(c.text) for c in obj.rows[0].cells]
        if not header or 'rating' not in header[0]:
            continue
        if not any(('check' in h or 'result' in h) for h in header[1:]):
            continue
        for row in obj.rows[1:]:
            if len(row.cells) < 2:
                continue
            rating = image_rating(row.cells[0], doc)
            check = clean(row.cells[1].text)
            if not rating or not check:
                continue
            if rating == 'Green' and not include_green:
                continue
            h = find_next_heading(headings, idx, check)
            if h is None:
                h = heading_for_block(headings, idx)
            if h is None:
                continue
            content = section_content(blocks, headings, h)
            directive = make_directive(h['title'], check, content)
            notes = extract_note_numbers(content)
            candidates.append({
                'section': str(h['number']),
                'section_title': h['title'],
                'rating': rating,
                'check': check,
                'directive': directive,
                'notes': notes,
            })

    # Also capture explicit recommendations under headings that have no rating-table row, but only if a nearby rating row
    # points to the same heading and is missing due to a PDF/formatting limitation. DOCX path normally does not need this.
    # De-duplicate exact section/check pairs.
    out = []
    seen = set()
    for x in candidates:
        key = (x['section'], x['check'])
        if key in seen:
            continue
        seen.add(key)
        out.append(x)
    return out


def parse_pdf(path, include_green=False):
    # Best-effort PDF fallback. Text-based EWA PDFs are supported, but rating icons may not be machine-readable.
    import pdfplumber
    records = []
    with pdfplumber.open(path) as pdf:
        full = '\n'.join((p.extract_text() or '') for p in pdf.pages)
    # Identify numbered headings in the common EWA format.
    heading_pat = re.compile(r'(?m)^\s*(\d+)\s+(.+?)\s+\d+/\d+\s*$')
    matches = list(heading_pat.finditer(full))
    for i, m in enumerate(matches):
        sec, title = m.group(1), clean(m.group(2))
        end = matches[i+1].start() if i+1 < len(matches) else len(full)
        block = full[m.end():end]
        # Split on explicit Recommendation markers; without icon ratings we only emit entries where Red/Yellow is explicitly named.
        rec = extract_recommendation(block)
        if not rec:
            continue
        rating = 'Red' if re.search(r'\bRED\b', block, re.I) else 'Yellow' if re.search(r'\bYELLOW\b', block, re.I) else ''
        if not rating:
            continue
        if rating == 'Green' and not include_green:
            continue
        notes = extract_note_numbers(block)
        records.append({'section': sec, 'section_title': title, 'rating': rating, 'check': title, 'directive': ('Recommendation: ' + rec), 'notes': notes})
    return records


def note_display(notes):
    return ", ".join(f"SAP Note {n}" for n in notes) if notes else ""


def write_excel(records, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'EWA Action Items'
    headers = ['Section', 'EWA Rating', 'EWA Directives', 'SAP note', 'SID', 'Ticket Reference', 'Responsible Team', 'Status', 'Next Actions']
    ws.append(headers)
    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(color='FFFFFF', bold=True)
    thin = Side(style='thin', color='D9E1F2')
    for c in ws[1]:
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = Border(bottom=thin)

    rating_fills = {
        'Red': PatternFill('solid', fgColor='FFC7CE'),
        'Yellow': PatternFill('solid', fgColor='FFEB9C'),
        'Green': PatternFill('solid', fgColor='C6EFCE'),
    }
    for rec in records:
        ws.append([rec['section'], rec['rating'], rec['directive'], '', '', '', '', '', ''])
        row = ws.max_row
        ws.cell(row, 4).value = note_display(rec['notes'])
        if rec['notes']:
            # The main cell links to the first referenced note; all referenced notes are also listed individually on the SAP Notes sheet.
            ws.cell(row, 4).hyperlink = SAP_NOTE_URL.format(note=rec['notes'][0])
            ws.cell(row, 4).style = 'Hyperlink'
        ws.cell(row, 2).fill = rating_fills.get(rec['rating'], PatternFill())
        for col in range(1, 10):
            ws.cell(row, col).alignment = Alignment(vertical='top', wrap_text=True)
            ws.cell(row, col).border = Border(bottom=thin)

    widths = {1: 12, 2: 14, 3: 90, 4: 32, 5: 14, 6: 22, 7: 24, 8: 16, 9: 40}
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    # Status dropdown for the blank tracker column.
    dv = DataValidation(type='list', formula1='"Open,In Progress,Completed,Deferred,Not Applicable"', allow_blank=True)
    ws.add_data_validation(dv)
    if ws.max_row >= 2:
        dv.add(f'H2:H{ws.max_row}')

    # Add an Instructions sheet.
    ins = wb.create_sheet('Instructions')
    instructions = [
        ('Purpose', 'Extract EWA rated checks and source-backed recommendations into an action tracker.'),
        ('Source', 'The program reads the EWA report. DOCX reports are preferred because EWA rating icons can be read directly.'),
        ('Section', 'Uses the report heading sequence number and links each rated check to the corresponding heading/subcategory.'),
        ('EWA Rating', 'Red / Yellow / Green derived from the EWA rating icon in DOCX.'),
        ('EWA Directives', 'Source-backed analysis and recommendation text from the corresponding EWA section. No recommendation is invented when the report does not provide one.'),
        ('SAP note', 'Referenced SAP Notes are shown as hyperlinks to SAP for Me note pages.'),
        ('Tracker columns', 'SID, Ticket Reference, Responsible Team, Status, and Next Actions are intentionally left blank for follow-up.'),
    ]
    for r, (a,b) in enumerate(instructions,1):
        ins.cell(r,1).value=a; ins.cell(r,2).value=b
        ins.cell(r,1).font=Font(bold=True)
        ins.cell(r,1).fill=PatternFill('solid', fgColor='D9EAF7')
        ins.cell(r,1).alignment=Alignment(vertical='top', wrap_text=True)
        ins.cell(r,2).alignment=Alignment(vertical='top', wrap_text=True)
    ins.column_dimensions['A'].width=22; ins.column_dimensions['B'].width=100

    notes_ws = wb.create_sheet('SAP Notes')
    notes_ws.append(['SAP Note', 'SAP Note Link'])
    for c in notes_ws[1]:
        c.fill = header_fill
        c.font = header_font
    all_notes = []
    for rec in records:
        for n in rec['notes']:
            if n not in all_notes:
                all_notes.append(n)
    for n in all_notes:
        r = notes_ws.max_row + 1
        notes_ws.cell(r,1).value = f'SAP Note {n}'
        notes_ws.cell(r,1).hyperlink = SAP_NOTE_URL.format(note=n)
        notes_ws.cell(r,1).style = 'Hyperlink'
        notes_ws.cell(r,2).value = SAP_NOTE_URL.format(note=n)
        notes_ws.cell(r,2).hyperlink = SAP_NOTE_URL.format(note=n)
        notes_ws.cell(r,2).style = 'Hyperlink'
    notes_ws.column_dimensions['A'].width = 24
    notes_ws.column_dimensions['B'].width = 60

    wb.save(output_path)


def process(input_path, output_path, include_green=False):
    ext = Path(input_path).suffix.lower()
    if ext == '.docx':
        records = parse_docx(input_path, include_green)
    elif ext == '.pdf':
        records = parse_pdf(input_path, include_green)
    else:
        raise ValueError('Please select a DOCX or PDF EWA report.')
    if not records:
        raise ValueError('No rated action items were detected. For DOCX, make sure this is an SAP EWA report with rating icons.')
    write_excel(records, output_path)
    return records


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('760x430')
        self.resizable(False, False)
        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.include_green = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value='Ready')
        self._build()

    def _build(self):
        pad = {'padx': 18, 'pady': 8}
        frm = ttk.Frame(self, padding=18)
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text=APP_TITLE, font=('Segoe UI', 18, 'bold')).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0,18))
        ttk.Label(frm, text='EWA Report (DOCX preferred; PDF best-effort):').grid(row=1, column=0, sticky='w', **pad)
        ttk.Entry(frm, textvariable=self.input_var, width=70).grid(row=1, column=1, sticky='ew', **pad)
        ttk.Button(frm, text='Browse…', command=self.browse_input).grid(row=1, column=2, **pad)
        ttk.Label(frm, text='Output Excel file:').grid(row=2, column=0, sticky='w', **pad)
        ttk.Entry(frm, textvariable=self.output_var, width=70).grid(row=2, column=1, sticky='ew', **pad)
        ttk.Button(frm, text='Save As…', command=self.browse_output).grid(row=2, column=2, **pad)
        ttk.Checkbutton(frm, text='Include Green-rated checks', variable=self.include_green).grid(row=3, column=1, sticky='w', padx=18, pady=12)
        ttk.Label(frm, text='Default: Red and Yellow action items only.').grid(row=4, column=1, sticky='w', padx=18)
        self.run_btn = ttk.Button(frm, text='Extract EWA Action Items', command=self.run)
        self.run_btn.grid(row=5, column=1, sticky='w', padx=18, pady=20)
        ttk.Label(frm, textvariable=self.status_var, foreground='#333333', wraplength=650).grid(row=6, column=0, columnspan=3, sticky='w', padx=18, pady=10)
        ttk.Label(frm, text='Output columns: Section | EWA Rating | EWA Directives | SAP note | SID | Ticket Reference | Responsible Team | Status | Next Actions', wraplength=680).grid(row=7, column=0, columnspan=3, sticky='w', padx=18, pady=10)
        frm.columnconfigure(1, weight=1)

    def browse_input(self):
        p = filedialog.askopenfilename(filetypes=[('EWA Reports', '*.docx *.pdf'), ('Word documents', '*.docx'), ('PDF files', '*.pdf')])
        if p:
            self.input_var.set(p)
            if not self.output_var.get():
                self.output_var.set(str(Path(p).with_name(Path(p).stem + '_EWA_Action_Items.xlsx')))

    def browse_output(self):
        p = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Workbook', '*.xlsx')], initialfile='EWA_Action_Items.xlsx')
        if p:
            self.output_var.set(p)

    def run(self):
        inp = self.input_var.get().strip()
        out = self.output_var.get().strip()
        if not inp or not os.path.isfile(inp):
            messagebox.showerror(APP_TITLE, 'Please select a valid EWA DOCX or PDF report.')
            return
        if not out:
            out = str(Path(inp).with_name(Path(inp).stem + '_EWA_Action_Items.xlsx'))
            self.output_var.set(out)
        try:
            self.run_btn.config(state='disabled')
            self.status_var.set('Processing report…')
            self.update_idletasks()
            records = process(inp, out, self.include_green.get())
            red = sum(1 for r in records if r['rating']=='Red')
            yellow = sum(1 for r in records if r['rating']=='Yellow')
            green = sum(1 for r in records if r['rating']=='Green')
            self.status_var.set(f'Completed: {len(records)} items (Red={red}, Yellow={yellow}, Green={green}). Output: {out}')
            if messagebox.askyesno(APP_TITLE, f'Excel created successfully with {len(records)} items.\n\nOpen the Excel file now?'):
                try:
                    os.startfile(out)
                except Exception:
                    pass
        except Exception as e:
            self.status_var.set('Failed.')
            messagebox.showerror(APP_TITLE, str(e))
        finally:
            self.run_btn.config(state='normal')


def main():
    if len(sys.argv) > 1:
        inp = sys.argv[1]
        out = sys.argv[2] if len(sys.argv)>2 else str(Path(inp).with_name(Path(inp).stem + '_EWA_Action_Items.xlsx'))
        process(inp, out, False)
        print(out)
        return
    App().mainloop()

if __name__ == '__main__':
    main()
