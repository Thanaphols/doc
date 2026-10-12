"""Build one thesis chapter from input/chapters/chN.txt, formatted like the
senior thesis that passed review (input/reference/ไฟล์เล่มรุ่นพี่ที่ตรวจผ่านแล้ว.docx).

Usage (from repo root):  python .agent/skills/rmuti-thesis/scripts/thesis_build.py 2 [--no-pdf]

The document starts from input/reference/rmuti_thesis_template.docx, so fonts,
Heading 1-4 styles and their auto-numbering ("บทที่ %1", "%1.%2", ...) are the
template's own. Chapter file format (UTF-8, one paragraph per line):
    line 1         chapter number, e.g. "2"
    line 2         chapter title
    "2.1 ข้อความ"   -> Heading 2 (number is stripped, Word numbers it)
    "2.1.1 ข้อความ" -> chapter 1: Heading 3 (auto number)
                     chapter 2+: Normal, typed number, 1 tab   (as in senior book)
    "2.1.1.1 ..."   -> chapter 1: Heading 4 · chapter 2+: Normal, 2 tabs
    "@toc"          -> list of all X.X topics (senior chapter 2 intro)
    "1) ข้อความ"     -> Normal, 3 tabs (senior's list under X.X.X.X)
    "@fig ชื่อรูป"   -> framed 1x1 box with a placeholder + caption "รูปที่ N.n  ชื่อรูป" inside
    "@table ชื่อ"    -> caption "ตารางที่ N.n  ชื่อ" (left) + Table Grid built from the
                       following "a | b | c" lines; first row is the header
    other text      -> Normal body, 1 tab (2 tabs when under an X.X.X item)
"""
import os
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
TEMPLATE = os.path.join(ROOT, "input", "reference", "rmuti_thesis_template.docx")
NUM_RE = re.compile(r"^(\d+(?:\.\d+)+)\s+")
LIST_RE = re.compile(r"^\d+\)\s")
# column widths (cm) by column count; text width is 14.65 cm
COL_CM = {3: [2.2, 5.6, 6.85], 4: [3.6, 2.9, 4.7, 3.45]}


def add_caption(container, label, chapter, text, align):
    """Caption paragraph "<label> N.n  text" with a SEQ field, so Word's
    Table of Figures can list it. The number is cached so it shows before F9."""
    add = container.add_paragraph
    p = add(style="Caption")
    p.alignment = align
    p.add_run(f"{label} {chapter}.")
    seq_no = container._seq[label] = container._seq.get(label, 0) + 1
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), f" SEQ {label} \\* ARABIC ")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = str(seq_no)
    r.append(t)
    fld.append(r)
    p._p.append(fld)
    p.add_run(f"  {text}")
    return p


def add_figure(doc, chapter, title):
    seq_no = doc._seq.get("รูปที่", 0) + 1
    candidates = [
        os.path.join(ROOT, "input", "figures", f"ch{chapter}", f"{chapter}.{seq_no}.png"),
        os.path.join(ROOT, "input", "figures", f"ch{chapter}", f"{seq_no}.png"),
        os.path.join(ROOT, "input", "figures", f"{chapter}.{seq_no}.png"),
        os.path.join(ROOT, "input", "figures", f"ch{chapter}", f"{chapter}.{seq_no}.jpg"),
    ]
    img_path = next((p for p in candidates if os.path.exists(p)), None)

    box = doc.add_table(rows=1, cols=1)
    tblBorders = OxmlElement("w:tblBorders")
    for b in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{b}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        tblBorders.append(el)
    box._tbl.tblPr.append(tblBorders)

    cell = box.rows[0].cells[0]
    cell.width = Cm(14.65)
    ph = cell.paragraphs[0]
    ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if img_path:
        run = ph.add_run()
        run.add_picture(img_path, width=Cm(13.0))
    else:
        ph.add_run(f"[ใส่รูป: {title}]")
    cell._seq = doc._seq
    p_cap = add_caption(cell, "รูปที่", chapter, title, WD_ALIGN_PARAGRAPH.CENTER)
    p_cap.paragraph_format.space_before = Pt(4)
    p_cap.paragraph_format.space_after = Pt(6)
    doc.add_paragraph()


A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"


def normalize_figures(doc):
    """House style for figures (as the author fixed by hand): the picture has a
    thin black outline, the 1x1 table holding it has no borders at all."""
    for tbl in doc.element.body.iter(qn("w:tbl")):
        if tbl.find(".//" + PIC + "pic") is None:
            continue
        tblPr = tbl.tblPr
        if tblPr.find(qn("w:tblStyle")) is None:
            st = OxmlElement("w:tblStyle")
            st.set(qn("w:val"), "TableGrid")
            tblPr.insert(0, st)
        for old in tblPr.findall(qn("w:tblBorders")):
            tblPr.remove(old)
        b = OxmlElement("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
            b.append(el)
        # schema order inside tblPr: ... tblW, jc, tblCellSpacing, tblInd, tblBorders, shd, tblLayout ...
        after = next((tblPr.find(qn(t)) for t in ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")
                      if tblPr.find(qn(t)) is not None), None)
        (after.addprevious if after is not None else tblPr.append)(b)
        for tcb in tbl.iter(qn("w:tcBorders")):
            tcb.getparent().remove(tcb)
    for spPr in doc.element.body.iter(PIC + "spPr"):
        if spPr.find(A + "ln") is not None:
            continue
        ln = OxmlElement("a:ln")
        fill = OxmlElement("a:solidFill")
        clr = OxmlElement("a:schemeClr")
        clr.set("val", "tx1")
        fill.append(clr)
        ln.append(fill)
        # a:ln comes after xfrm / geometry / fill, before effects
        nxt = next((c for c in spPr if c.tag in (A + "effectLst", A + "effectDag", A + "scene3d", A + "sp3d", A + "extLst")), None)
        (nxt.addprevious if nxt is not None else spPr.append)(ln)


def add_data_table(doc, chapter, title, rows):
    add_caption(doc, "ตารางที่", chapter, title, WD_ALIGN_PARAGRAPH.LEFT)
    cells = [[c.strip() for c in r.split("|")] for r in rows]
    ncol = max(len(r) for r in cells)
    t = doc.add_table(rows=len(cells), cols=ncol)
    t.style = "Table Grid"
    widths = COL_CM.get(ncol)
    for i, row in enumerate(cells):
        for j in range(ncol):
            cell = t.rows[i].cells[j]
            if widths:
                cell.width = Cm(widths[j])
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.add_run(row[j] if j < len(row) else "")
    trPr = t.rows[0]._tr.get_or_add_trPr()  # repeat header row on each page
    trPr.append(OxmlElement("w:tblHeader"))
    doc.add_paragraph()


def mark_thai(run):
    """Tag a run as Thai complex script. Without <w:cs/> Word treats Thai as
    Latin, can't word-break it, and thaiDistribute spreads the letters apart."""
    rPr = run._r.get_or_add_rPr()
    for tag in ("w:rFonts", "w:cs", "w:lang"):
        for old in rPr.findall(qn(tag)):
            rPr.remove(old)
    fonts = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        fonts.set(qn(a), "TH SarabunPSK")
    fonts.set(qn("w:hint"), "cs")
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "en-US")
    lang.set(qn("w:bidi"), "th-TH")
    rPr.insert(0, fonts)
    rPr.append(OxmlElement("w:cs"))
    rPr.append(lang)


def split_docx_table(doc, tbl_index, split_after_row, cont_caption_text, style="Caption"):
    """Split doc.tables[tbl_index] after data row split_after_row. The second
    part starts a new page under "<caption> (ต่อ)" and repeats header row 0."""
    import copy
    t = doc.tables[tbl_index]
    parent = t._tbl.getparent()
    idx = parent.index(t._tbl)
    tbl2 = copy.deepcopy(t._tbl)
    t2 = docx.table.Table(tbl2, doc)
    keep = split_after_row + 1  # header + rows 1..split_after_row
    for _ in range(len(t.rows) - keep):
        t._tbl.remove(t.rows[keep]._tr)
    for _ in range(split_after_row):
        t2._tbl.remove(t2.rows[1]._tr)

    p = docx.text.paragraph.Paragraph(OxmlElement("w:p"), doc)
    p.style = style
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    pf.keep_with_next = True
    pf.page_break_before = True
    r = p.add_run(cont_caption_text)
    r.bold = True
    mark_thai(r)
    parent.insert(idx + 1, p._p)
    parent.insert(idx + 2, tbl2)


def unsplit_tables(doc):
    """Undo earlier split_tables runs: "<caption> (ต่อ)" + its table go back into the
    table above (minus the repeated header row). A chapter's page breaks change once
    it sits in the book, so it must be split again on the final layout, not twice."""
    joined = 0
    for p in list(doc.paragraphs):
        if not (p.text.strip().startswith("ตารางที่") and p.text.strip().endswith("(ต่อ)")):
            continue
        prev, nxt = p._p.getprevious(), p._p.getnext()
        if prev is None or nxt is None or prev.tag != qn("w:tbl") or nxt.tag != qn("w:tbl"):
            continue
        for tr in nxt.findall(qn("w:tr"))[1:]:
            prev.append(tr)
        nxt.getparent().remove(nxt)
        p._p.getparent().remove(p._p)
        joined += 1
    return joined


def _pic_count(doc):
    xml = doc.element.body.xml
    return xml.count("<pic:pic"), xml.count("<w:drawing")


def _captioned_tables(doc):
    """[(table index, "ตารางที่ N.k  ชื่อ")] for data tables (caption paragraph
    right before them). Figure boxes have no caption before them and are skipped.
    N comes from the caption itself, so a merged multi-chapter file works too."""
    out, chapter, k, title = [], None, 0, ""
    for i, t in enumerate(doc.tables):
        prev = t._tbl.getprevious()
        if prev is None or prev.tag != qn("w:p"):
            continue
        text = docx.text.paragraph.Paragraph(prev, doc).text.strip()
        m = re.match(r"ตารางที่ (\d+)\.", text)
        if not m:
            continue
        if text.endswith("(ต่อ)"):  # part of a table already split: same number/title
            out.append((i, f"ตารางที่ {chapter}.{k}  {title}"))
            continue
        if m.group(1) != chapter:
            chapter, k = m.group(1), 0
        k += 1
        title = text.split("  ", 1)[1].strip() if "  " in text else text
        out.append((i, f"ตารางที่ {chapter}.{k}  {title}"))
    return out


def _save(doc, path, tries=10):
    """Word releases a file a moment after Document.Close; retry the save."""
    import time
    for n in range(tries):
        try:
            return doc.save(path)
        except OSError:
            if n == tries - 1:
                raise
            time.sleep(1)


def _open(path, tries=10):
    """Word releases a file a moment after Document.Close; retry the open."""
    import time
    for n in range(tries):
        try:
            return docx.Document(path)
        except OSError:
            if n == tries - 1:
                raise
            time.sleep(1)


def split_tables(docx_path, max_rounds=40):
    """Make every captioned table that runs past a page end continue on the next
    page under "ตารางที่ N.n  ชื่อ (ต่อ)" + repeated header (RMUTI rule).
    Word paginates; python-docx splits one table per round, then re-measures.
    Returns the number of splits. Refuses to save if any picture would be lost."""
    import win32com.client
    path = os.path.abspath(docx_path)
    doc = _open(path)
    pics = _pic_count(doc)
    for i, _ in _captioned_tables(doc):
        t = doc.tables[i]
        for row in t.rows:  # never break a single row across pages
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))
        # caption + header row + first data row stay on one page
        docx.text.paragraph.Paragraph(t._tbl.getprevious(), doc).paragraph_format.keep_with_next = True
        for cell in t.rows[0].cells:
            for p in cell.paragraphs:
                p.paragraph_format.keep_with_next = True
    _save(doc, path)

    word = win32com.client.DispatchEx("Word.Application")  # own instance, never the user's open Word
    word.Visible = False
    word.DisplayAlerts = 0  # no hidden modal dialogs
    splits = 0
    try:
        for _ in range(max_rounds):
            d = word.Documents.Open(path)
            d.Repaginate()
            hit = None
            for i, caption in _captioned_tables(doc):
                tbl = d.Tables(i + 1)
                rows = tbl.Rows
                first = rows(1).Range.Information(3)  # wdActiveEndPageNumber
                cap = tbl.Range.Paragraphs(1).Previous(1).Range.Information(3)
                if cap != first or (rows.Count > 1 and rows(2).Range.Information(3) != first):
                    print(f"WARN {caption}: caption/header not on the same page as the first row")
                for r in range(3, rows.Count + 1):  # keep >= 1 data row on the first page
                    if rows(r).Range.Information(3) > first:
                        hit = (i, r - 2, caption)
                        break
                if hit:
                    break
            d.Close(False)
            if not hit:
                return splits
            i, keep, caption = hit
            cap_p = docx.text.paragraph.Paragraph(doc.tables[i]._tbl.getprevious(), doc)
            if keep < 2 and not cap_p.paragraph_format.page_break_before:
                # one lone row on the page looks broken: move the whole table to the next page
                cap_p.paragraph_format.page_break_before = True
                _save(doc, path)
                print(f"moved {caption} to the next page")
                continue
            split_docx_table(doc, i, keep, caption + " (ต่อ)")
            if _pic_count(doc) != pics:
                raise RuntimeError(f"pictures changed {pics} -> {_pic_count(doc)}, not saved")
            _save(doc, path)
            splits += 1
            print(f"split {caption} after data row {keep}")
        raise RuntimeError(f"tables still overflow after {max_rounds} rounds")
    finally:
        word.Quit()


def start_chapter_numbering(doc, chapter):
    """Heading styles share one numbering list; make it start at this chapter."""
    num_id = doc.styles["Heading 1"].element.pPr.find(qn("w:numPr")).find(qn("w:numId")).get(qn("w:val"))
    numbering = doc.part.numbering_part.element
    num = next(n for n in numbering.findall(qn("w:num")) if n.get(qn("w:numId")) == num_id)
    for old in num.findall(qn("w:lvlOverride")):
        num.remove(old)
    ov = OxmlElement("w:lvlOverride")
    ov.set(qn("w:ilvl"), "0")
    so = OxmlElement("w:startOverride")
    so.set(qn("w:val"), str(chapter))
    ov.append(so)
    num.append(ov)


def build_doc(lines):
    lines = [l.rstrip() for l in lines if l.strip()]
    chapter, title, body = int(lines[0]), lines[1].strip(), lines[2:]
    doc = docx.Document(TEMPLATE)
    for el in list(doc.element.body):  # keep only the template's section settings
        if el.tag != qn("w:sectPr"):
            doc.element.body.remove(el)
    s = doc.sections[0]
    # ponytail: one section at 3.81 cm; first page reaches 5.08 cm via 36pt
    # space-before on Heading 1 (senior book used a manual section break instead).
    s.top_margin = Cm(3.81)
    start_chapter_numbering(doc, chapter)

    h1 = doc.add_paragraph(style="Heading 1")
    h1.paragraph_format.space_before = Pt(36)
    h1.add_run().add_break()  # "บทที่ N" (auto) / line break / title
    h1.add_run(title)

    topics = [NUM_RE.sub("", l) for l in body if (m := NUM_RE.match(l)) and m.group(1).count(".") == 1]
    merged, just_merged = [], False
    for line in (l.strip() for l in body):
        prev = merged[-1] if merged else ""
        pm = NUM_RE.match(prev)
        if (chapter > 1 and not just_merged and pm and pm.group(1).count(".") >= 2
                and not NUM_RE.match(line) and not line.startswith("@")
                and " | " not in line and not LIST_RE.match(line)
                and not prev.endswith("ดังนี้")):
            merged[-1] = f"{prev} {line}"
            just_merged = True  # only the first body line joins the item
        else:
            merged.append(line)
            just_merged = False

    doc._seq = {}
    first_h2, under_item = True, False
    i = 0
    while i < len(merged):
        line = merged[i]
        i += 1
        if line.startswith("@fig "):
            add_figure(doc, chapter, line[5:].strip())
            continue
        if line.startswith("@table "):
            rows = []
            while i < len(merged) and " | " in merged[i]:
                rows.append(merged[i])
                i += 1
            add_data_table(doc, chapter, line[7:].strip(), rows)
            continue
        if LIST_RE.match(line):
            doc.add_paragraph("			" + line)
            continue
        if line == "@toc":
            for topic_idx, t in enumerate(topics, 1):
                p = doc.add_paragraph(f"{chapter}.{topic_idx} {t}")
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.left_indent = Cm(1.27)
            continue
        m = NUM_RE.match(line)
        level = m.group(1).count(".") + 1 if m else 0
        if level == 2:
            if first_h2 and chapter == 1:
                doc.add_paragraph()  # senior: blank line under the chapter title
            elif not first_h2:
                doc.add_paragraph()  # senior: blank line before each X.X
            first_h2, under_item = False, False
            doc.add_paragraph(NUM_RE.sub("", line), style="Heading 2")
        elif level >= 3 and chapter == 1:
            doc.add_paragraph(NUM_RE.sub("", line), style=f"Heading {min(level, 4)}")
        elif level >= 3:
            doc.add_paragraph("\t" * (level - 2) + line)
            under_item = True
        else:
            doc.add_paragraph(("\t\t" if under_item else "\t") + line)
    for p in all_paragraphs(doc):
        for r in p.runs:
            mark_thai(r)
    normalize_figures(doc)
    return doc, chapter, title


def all_paragraphs(doc):
    yield from doc.paragraphs
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def check(doc):
    """Fail loudly if output drifts from the senior-book format."""
    paras = doc.paragraphs
    assert paras[0].style.name == "Heading 1"
    for p in list(all_paragraphs(doc))[1:]:
        if p.style.name == "Heading 2":
            assert not NUM_RE.match(p.text), f"typed number on Heading 2: {p.text[:40]}"
        for r in p.runs:
            assert r._r.rPr.find(qn("w:cs")) is not None, f"run not tagged Thai: {p.text[:40]}"
            assert r.bold is None, f"direct bold (use styles): {p.text[:40]}"
        assert p.paragraph_format.first_line_indent is None, f"use tabs, not indent: {p.text[:40]}"


FONT_FILE = "C:/Windows/Fonts/THSarabunNew.ttf"  # same metrics as TH SarabunPSK
LINE_PT = 415.0  # A4 21 cm - 3.81 - 2.54 margins = 14.65 cm
TAB_PT = 36.0    # template default tab stop 1.27 cm
# where the text starts on line 1 (pt); from the template's numbering indents
FIRST_OFFSET_PT = {"Heading 3": 58.5, "Heading 4": 42.6}


def short_last_lines(doc, min_fill=0.4):
    """Estimate wrapping and list paragraphs whose last line holds only a
    word or two (looks ragged; advisors ask to fix). Approximate: sums glyph
    advances, ignores Word's exact Thai break points."""
    try:
        from PIL import ImageFont
        font = ImageFont.truetype(FONT_FILE, 16)
    except (ImportError, OSError):
        return []
    bad = []
    for p in doc.paragraphs:
        text = p.text.replace("	", "")
        if not text.strip() or p.style.name in ("Heading 1", "Heading 2", "Caption"):
            continue
        width = font.getlength(text)
        if p.style.name in FIRST_OFFSET_PT:
            level = int(p.style.name[-1])
            width += font.getlength(".".join(["0"] * level) + "  ")
            first = LINE_PT - FIRST_OFFSET_PT[p.style.name]
        else:
            first = LINE_PT - TAB_PT * p.text.count("	", 0, 3)
        # >95% full is risky too: estimate error can push one word to a new line
        if width <= first:
            if width > 0.95 * first:
                bad.append((round(width / first * 100), text))
            continue
        last = (width - first) % LINE_PT
        if last < min_fill * LINE_PT or last > 0.95 * LINE_PT:
            bad.append((round(last / LINE_PT * 100), text))
    return bad


def export_pdf(docx_path, pdf_path):
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")  # own instance, never the user's open Word
    word.Visible = False
    word.DisplayAlerts = 0  # no hidden modal dialogs
    try:
        d = word.Documents.Open(docx_path)
        d.SaveAs(pdf_path, FileFormat=17)
        d.Close()
    finally:
        word.Quit()


def render_pngs(pdf_path, stem):
    import pymupdf
    out_dir = os.path.join(ROOT, "process", "temp")
    os.makedirs(out_dir, exist_ok=True)
    pdf = pymupdf.open(pdf_path)
    for i, page in enumerate(pdf):
        page.get_pixmap(dpi=110).save(os.path.join(out_dir, f"{stem}_p{i + 1}.png"))
    return len(pdf)


def build_bib(argv):
    bib_path = os.path.join(ROOT, "input", "chapters", "bib.txt")
    with open(bib_path, encoding="utf-8") as f:
        raw_lines = [l.strip() for l in f if l.strip() and not l.startswith("#")]

    doc = docx.Document(TEMPLATE)
    for el in list(doc.element.body):
        if el.tag != qn("w:sectPr"):
            doc.element.body.remove(el)

    s = doc.sections[0]
    s.top_margin = Cm(3.81)

    # Heading: บรรณานุกรม
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.paragraph_format.space_before = Pt(36)
    h.paragraph_format.space_after = Pt(16)
    r = h.add_run("บรรณานุกรม")
    r.bold = r.font.cs_bold = True  # Thai text is bolded by bCs
    mark_thai(r)

    # Entries
    for line in raw_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        pPr = p._p.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "720")
        ind.set(qn("w:hanging"), "720")
        pPr.append(ind)

        # Parse *italic*
        tokens = line.split("*")
        for idx, token in enumerate(tokens):
            if not token:
                continue
            run = p.add_run(token)
            if idx % 2 == 1:
                run.italic = True
            mark_thai(run)

    stem = "Chatshop_บรรณานุกรม"
    out = os.path.join(ROOT, "output", stem + ".docx")
    try:
        doc.save(out)
    except PermissionError:
        out = os.path.join(ROOT, "output", stem + "_new.docx")
        doc.save(out)
        print("Target open in Word, saved instead to", out)
    print("docx:", out)
    if "--no-pdf" in argv:
        return
    pdf = os.path.join(ROOT, "output", stem + ".pdf")
    export_pdf(out, pdf)
    print(f"pdf: {pdf} ({render_pngs(pdf, stem)} pages, previews in process/temp/)")


def main(argv):
    n = argv[1]
    if n == "bib":
        build_bib(argv)
        return
    with open(os.path.join(ROOT, "input", "chapters", f"ch{n}.txt"), encoding="utf-8") as f:
        doc, chapter, title = build_doc(f.readlines())
    check(doc)
    for fill, text in short_last_lines(doc):
        print(f"WARN last line ~{fill}% full, reword: {text[:50]}")
    stem = f"Chatshop_บทที่{chapter}_{title.replace(' ', '')}"
    out = os.path.join(ROOT, "output", stem + ".docx")
    try:
        doc.save(out)
    except PermissionError:  # file open in Word: don't lose the build
        out = os.path.join(ROOT, "output", stem + "_new.docx")
        doc.save(out)
        print("Target open in Word, saved instead to", out)
    print("docx:", out)
    if "--no-pdf" in argv:
        return
    split_tables(out)
    pdf = os.path.join(ROOT, "output", stem + ".pdf")
    export_pdf(out, pdf)
    print(f"pdf: {pdf} ({render_pngs(pdf, stem)} pages, previews in process/temp/)")


if __name__ == "__main__":
    main(sys.argv)

