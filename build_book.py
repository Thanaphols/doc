"""Assemble the complete thesis: covers .. acknowledgements (from the sample in
input/reference), table of contents / tables / figures with real page numbers,
chapters 1-5 (output/Chatshop_บทที่1_5.docx) and the bibliography.

Usage (repo root, Word closed):  py build_book.py [--rebuild-body]
    --rebuild-body   run build_ch1_ch5 first (slow; otherwise reuse output/Chatshop_บทที่1_5.docx)
"""
import os
import shutil
import sys

import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, ".agent", "skills", "rmuti-thesis", "scripts"))
from thesis_build import _pic_count, _save, _open, export_pdf, mark_thai, render_pngs, split_docx_table, normalize_figures  # noqa: E402

OUT = os.path.join(ROOT, "output")
TMP = os.path.join(ROOT, "process", "temp")
SAMPLE = os.path.join(ROOT, "input", "reference", "01 ปกนอก ถึง กิตติกรรมประกาศ.docx")
BODY_SRC = os.path.join(TMP, "Chatshop_บทที่1_5.docx")  # written by build_ch1_ch5
from build_ch1_ch5 import src  # noqa: E402  output/ or output/old/
BIB_SRC = src("Chatshop_บรรณานุกรม.docx")
FRONT = os.path.join(TMP, "book_front.docx")
BODY = os.path.join(TMP, "book_body.docx")
BOOK = os.path.join(OUT, "Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v5.docx")

# ponytail: estimate until appendix ก is written; then add it to the book and measure instead.
APPENDIX_A_PAGES = 10
APPENDIX_A_TITLE = "ภาคผนวก ก คู่มือการใช้งานแอปพลิเคชัน"
# Word's thaiLetters sequence (skips ฃ ฅ ฆ; checked against our PDF: ... ฏ ฐ ฑ)
THAI = "กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ"
TEXT_CM, PAGE_COL_CM = 14.65, 1.5

TH_TITLE = "แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อ\nสำหรับธุรกิจออนไลน์"
EN_TITLE = "Centralized Chat and Order Management\nApplication for Online Business"
ADVISOR_TH, ADVISOR_EN = "ผู้ช่วยศาสตราจารย์ ดร.อติราช  สุขสวัสดิ์", "Assistant Professor, Atirarj Suksawad, D.Eng"  # thesis advisor
AUTHOR_TH, AUTHOR_EN = "ธนพล  วุฒิธาดา", "Thanaphol  Wuttitada"


def word_app():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")  # own instance, never the user's open Word
    w.Visible = False
    w.DisplayAlerts = 0
    return w


# ---------------------------------------------------------------- front matter
def set_text(p, text):
    """Replace a paragraph's text, keeping the formatting of its first text run."""
    runs = [r for r in p.runs if r.text]
    if not runs:
        r = p.add_run()
        mark_thai(r)
        runs = [r]
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""
    if any("฀" <= ch <= "๿" for ch in text):
        # runs[0] may be a bare tab run without <w:cs/>: Thai in it gets letter-spaced by thaiDistribute
        rPr = runs[0]._r.get_or_add_rPr()
        if rPr.find(qn("w:cs")) is None:
            mark_thai(runs[0])


def drop(p):
    assert p._p.pPr is None or p._p.pPr.find(qn("w:sectPr")) is None, "would delete a section break"
    p._p.getparent().remove(p._p)


def body_paras(src, min_len=40):
    return [p.text.strip() for p in docx.Document(src).paragraphs if len(p.text.strip()) >= min_len]


def fill_front(doc):
    ps = doc.paragraphs
    expect = {3: "ระบบจัดการสินค้า", 14: "ธนพล", 102: "หัวข้อปริญญานิพนธ์", 136: "บทคัดย่อ",
              153: "Abstract", 161: "กิตติกรรมประกาศ"}
    for i, s in expect.items():  # sample changed? stop instead of writing into the wrong lines
        assert ps[i].text.strip().startswith(s), (i, ps[i].text)

    allp = body_paras(src("Chatshop_บทคัดย่อ.docx"))
    abs_th = [t for t in allp if not t.startswith(("คำสำคัญ", "Keywords")) and not t.isascii()]
    abs_en = [t for t in allp if not t.startswith(("คำสำคัญ", "Keywords")) and t.isascii()]
    kw_th = next(t for t in allp if t.startswith("คำสำคัญ")).split(":", 1)[1].strip()
    kw_en = next(t for t in allp if t.startswith("Keywords")).split(":", 1)[1].strip()
    ack = [t.replace("อาจารย์ ดร.ประภาส ผ่องสนาม", ADVISOR_TH)
           for t in body_paras(src("Chatshop_กิตติกรรมประกาศ.docx"))]
    fix = lambda t: t.replace("แชทบอทอัจฉริยะ", "แชทบอท").replace("an intelligent chatbot", "a chatbot")
    abs_th, abs_en = [fix(t) for t in abs_th], [fix(t) for t in abs_en]
    assert len(abs_th) == 2 and len(abs_en) == 2 and len(ack) == 3, (len(abs_th), len(abs_en), len(ack))

    fac = "คณะวิศวกรรมศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน วิทยาเขตขอนแก่น"
    new = {
        # outer cover
        3: TH_TITLE, 4: EN_TITLE.split("\n")[0], 5: EN_TITLE.split("\n")[1],
        14: AUTHOR_TH, 15: "",
        26: "สาขาวิชาวิศวกรรมคอมพิวเตอร์", 27: fac, 28: "พ.ศ. 2569",
        29: "ลิขสิทธิ์ของคณะวิศวกรรมศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน",
        # inner cover TH
        30: TH_TITLE, 41: "\t" + AUTHOR_TH, 42: "",
        62: "สาขาวิชาวิศวกรรมคอมพิวเตอร์", 63: fac, 64: "พ.ศ. 2569",
        65: "ลิขสิทธิ์ของคณะวิศวกรรมศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน",
        # inner cover EN
        66: EN_TITLE.split("\n")[0].upper(), 67: EN_TITLE.split("\n")[1].upper(),
        77: "\t" + AUTHOR_EN, 78: "",
        97: "the Degree of Bachelor of Engineering", 100: "2026",
        101: "© Faculty of Engineering, Rajamangala University of Technology Isan",
        # approval page
        102: "หัวข้อปริญญานิพนธ์ \t:\t" + TH_TITLE.replace("\n", ""),
        103: "จัดทำโดย \t:\tธนพล วุฒิธาดา", 105: "ประธานที่ปรึกษา \t:\t" + ADVISOR_TH,
        109: fac,
        # Thai abstract
        127: "หัวข้อปริญญานิพนธ์\t" + TH_TITLE.replace("\n", ""),
        128: "จัดทำโดย\tธนพล วุฒิธาดา", 129: "ปีที่ปริญญานิพนธ์สำเร็จ\tพ.ศ. 2569",
        131: "ประธานที่ปรึกษา\t" + ADVISOR_TH,
        138: "\t" + abs_th[0], 139: "\t" + abs_th[1], 141: "คำสำคัญ\t\t:  " + kw_th,
        # English abstract
        144: "Project Title\t" + EN_TITLE.split("\n")[0],
        145: "                                 " + EN_TITLE.split("\n")[1],
        146: "Proposed by\tThanaphol  Wuttitada", 147: "Year\t2026",
        149: "Advisor\t" + ADVISOR_EN,
        155: "\t" + abs_en[0], 156: "\t" + abs_en[1], 158: "Keywords\t\t:  " + kw_en,
        # acknowledgements
        163: "\t" + ack[0], 164: "\t" + ack[1], 165: "\t" + ack[2],
        169: "\t\t\t" + AUTHOR_TH, 170: "",
    }
    for i, t in new.items():
        set_text(ps[i], t)
    assert ps[112].text.rstrip().endswith("คณบดีคณะวิศวกรรมศาสตร"), ps[112].text
    set_text(ps[112], ps[112].text.rstrip() + "์")  # sample typo: missing the final ์
    # committee: one signature, right-aligned (the advisor is the whole committee)
    assert "ประธานกรรมการสอบ" in ps[118].text and ps[125].text.strip().startswith("("), (ps[118].text, ps[125].text)
    for i, t in ((118, "................................................ประธานที่ปรึกษา"), (119, f"({ADVISOR_TH})")):
        set_text(ps[i], t)
        ps[i].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ps[i].paragraph_format.left_indent = ps[i].paragraph_format.first_line_indent = Cm(0)
    # bottom-up keeps earlier indexes valid: extra abstract/ack paragraphs, then committee lines 120-125
    for i in (166, 157, 140, 125, 124, 123, 122, 121, 120):
        drop(ps[i])


def _szcs(run, half_points):
    rPr = run._r.get_or_add_rPr()
    for old in rPr.findall(qn("w:szCs")):
        rPr.remove(old)
    el = OxmlElement("w:szCs")
    el.set(qn("w:val"), str(half_points))
    sz = rPr.find(qn("w:sz"))
    (sz.addnext if sz is not None else rPr.append)(el)


def pin_sample_formatting(doc):
    """The book is assembled inside the body document, whose Normal style has
    tab stops at 1.8 cm.., single spacing, thaiDistribute and 16 pt. Word drops
    direct formatting that is redundant in the source, so pinning values per
    paragraph does not survive InsertFile. Instead the sample paragraphs get their
    own style holding the sample's defaults (Word 2016: 11/14 pt, after 8 pt,
    line 1.08, left, no tabs); a style the body lacks is imported as-is."""
    st = doc.styles.add_style("FrontNormal", WD_STYLE_TYPE.PARAGRAPH)
    st.base_style = None
    pf = st.paragraph_format
    pf.space_after, pf.line_spacing, pf.alignment = Pt(8), 1.08, WD_ALIGN_PARAGRAPH.LEFT
    rPr = st.element.get_or_add_rPr()
    f = OxmlElement("w:rFonts")
    for a, v in (("asciiTheme", "minorHAnsi"), ("eastAsiaTheme", "minorHAnsi"),
                 ("hAnsiTheme", "minorHAnsi"), ("cstheme", "minorBidi")):
        f.set(qn("w:" + a), v)
    rPr.insert(0, f)
    st.font.size = Pt(11)
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), "28")
    rPr.find(qn("w:sz")).addnext(szcs)
    for p in doc.paragraphs:
        p.style = st


# ---------------------------------------------------------------- TOC tables
def _no_borders(tbl):
    tblPr = tbl._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{e}")
        el.set(qn("w:val"), "none")
        b.append(el)
    tblPr.append(b)
    mar = OxmlElement("w:tblCellMar")  # text starts at the margin like normal paragraphs
    for side in ("left", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def _cell(cell, text, align, indent_cm=0.0, bold=False):
    p = cell.paragraphs[0]
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before, pf.space_after = Pt(0), Pt(4)  # senior "toc 1" uses 5 pt; 4 pt keeps a lone last line off its own page
    pf.line_spacing = 1.0
    pf.left_indent = Cm(indent_cm)
    for r in p.runs:
        r._r.getparent().remove(r._r)
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(16)
    mark_thai(r)
    _szcs(r, 32)


def add_list(doc, title, left_head, rows):
    """One TOC list: centered bold title + borderless 2-column table. The page
    number sits in its own column, so it stays on the first line of a long title."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(title)
    r.bold = r.font.cs_bold = True  # Thai text is bolded by bCs, not b
    r.font.size = Pt(16)
    mark_thai(r)
    _szcs(r, 32)
    t = doc.add_table(rows=1 + len(rows), cols=2)
    _no_borders(t)
    widths = [Cm(TEXT_CM - PAGE_COL_CM), Cm(PAGE_COL_CM)]
    for i, (text, page, level) in enumerate([(left_head, "หน้า", 0)] + rows):
        row = t.rows[i]
        for j, c in enumerate(row.cells):
            c.width = widths[j]
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        _cell(row.cells[0], text, WD_ALIGN_PARAGRAPH.LEFT, 0.56 * level)
        _cell(row.cells[1], page, WD_ALIGN_PARAGRAPH.RIGHT)
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        if i == 0:
            trPr.append(OxmlElement("w:tblHeader"))
    return t


def build_front(entries, front_pages):
    """entries: dict with 'toc', 'tables', 'figures' -> [(text, page, level)].
    front_pages: letters for บทคัดย่อ/กิตติกรรมประกาศ/สารบัญตาราง/สารบัญรูป."""
    shutil.copy2(SAMPLE, FRONT)
    doc = docx.Document(FRONT)
    pics = _pic_count(doc)
    fill_front(doc)
    pin_sample_formatting(doc)
    toc = [("บทคัดย่อ", front_pages.get("บทคัดย่อ", "ง"), 0),
           ("กิตติกรรมประกาศ", front_pages.get("กิตติกรรมประกาศ", "ฉ"), 0),
           ("สารบัญตาราง", front_pages.get("สารบัญตาราง", "ญ"), 0),
           ("สารบัญรูป", front_pages.get("สารบัญรูป", "ฎ"), 0)] + entries["toc"]
    for title, head, rows in (("สารบัญ", "", toc), ("สารบัญตาราง", "ตาราง", entries["tables"]),
                              ("สารบัญรูป", "รูป", entries["figures"])):
        doc.add_section(WD_SECTION.NEW_PAGE)  # copies the acknowledgements section: thaiLetters, 3.81 cm
        add_list(doc, title, head, rows)
    assert _pic_count(doc) == pics
    _save(doc, FRONT)


def fit_one_page_sections(d, sections=range(1, 5)):
    """Cover, inner covers and approval page must stay one page each. Our title is
    longer than the sample's, so drop blank spacer lines (the one just above the
    next text block) until each fits."""
    changed = False
    for si in sections:
        for _ in range(15):
            rng = d.Sections(si).Range
            if rng.Information(3) == rng.Paragraphs(1).Range.Information(3):
                break
            paras = rng.Paragraphs
            for k in range(paras.Count - 1, 1, -1):
                if paras(k).Range.Text == chr(13) and paras(k + 1).Range.Text.strip():
                    paras(k).Range.Delete()
                    changed = True
                    break
            else:
                raise RuntimeError(f"section {si} overflows and has no blank line left to drop")
            d.Repaginate()
    return changed


def paginate_front():
    """Split TOC tables at page ends under "<title> (ต่อ)" and return the page
    letters of the front-matter headings. Each list is its own section, so lists
    are split independently; loop until nothing overflows."""
    word = word_app()
    try:
        for _ in range(20):
            d = word.Documents.Open(FRONT)
            d.Repaginate()
            if fit_one_page_sections(d):
                d.Save()
            n_front_tables = d.Tables.Count
            hits = []
            for ti in range(1, n_front_tables + 1):
                rows = d.Tables(ti).Rows
                first = rows(1).Range.Information(3)
                for r in range(3, rows.Count + 1):
                    if rows(r).Range.Information(3) > first:
                        hits.append((ti - 1, r - 2))
                        break
            letters = {}
            for p in d.Paragraphs:
                t = p.Range.Text.strip()
                if t in ("บทคัดย่อ", "Abstract", "กิตติกรรมประกาศ", "สารบัญ", "สารบัญตาราง", "สารบัญรูป") \
                        and p.Range.Information(12) is False:  # wdWithInTable
                    letters.setdefault(t, THAI[p.Range.Information(1) - 1])
            d.Close(False)
            if not hits:
                return letters
            doc = _open(FRONT)
            for ti, keep in reversed(hits):  # bottom-up keeps earlier table indexes valid
                title = docx.text.paragraph.Paragraph(doc.tables[ti]._tbl.getprevious(), doc).text.strip()
                title = title.replace(" (ต่อ)", "")
                split_docx_table(doc, ti, keep, f"{title} (ต่อ)", style="Normal")
                cont = docx.text.paragraph.Paragraph(doc.tables[ti]._tbl.getnext(), doc)
                cont.style = doc.styles["Normal"]
                cont.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in cont.runs:
                    r.font.cs_bold = True
                    r.font.size = Pt(16)
                    _szcs(r, 32)
                cont.paragraph_format.space_after = Pt(12)
            _save(doc, FRONT)
        raise RuntimeError("TOC still overflows")
    finally:
        word.Quit()


# ---------------------------------------------------------------- body
def build_body():
    """Chapters 1-5 + bibliography in one file; returns its TOC entries and last page."""
    shutil.copy2(BODY_SRC, BODY)
    word = word_app()
    try:
        d = word.Documents.Open(BODY)
        rng = d.Content
        rng.Collapse(0)
        rng.InsertBreak(2)  # section break, next page
        rng.Collapse(0)
        rng.InsertFile(BIB_SRC)
        d.Repaginate()
        toc, tables, figures = [], [], []
        for p in d.Paragraphs:
            lvl = p.OutlineLevel
            text = p.Range.Text.replace("\x0b", " ").replace("\r", "").replace("\x07", "").strip()
            if lvl in (1, 2) and text:
                num = p.Range.ListFormat.ListString.strip()
                page = str(p.Range.Information(3))
                if lvl == 1:
                    toc.append((f"{num} {text}".strip(), page, 0))
                else:
                    toc.append((f"{num} {text}".strip(), page, 1))
            elif text == "บรรณานุกรม":
                toc.append(("บรรณานุกรม", str(p.Range.Information(3)), 0))
            elif text.startswith(("ตารางที่", "รูปที่")) and not text.endswith("(ต่อ)"):
                (tables if text.startswith("ตาราง") else figures).append(
                    (text.replace("  ", " "), str(p.Range.Information(3)), 0))
        last = d.ComputeStatistics(2)
        d.Save()
        d.Close(False)
    finally:
        word.Quit()
    L = last
    toc += [("ภาคผนวก", str(L + 1), 0), (APPENDIX_A_TITLE, str(L + 2), 1),
            ("ประวัติผู้เขียน", str(L + 2 + APPENDIX_A_PAGES), 0)]
    return {"toc": toc, "tables": tables, "figures": figures}, last


# ---------------------------------------------------------------- final merge
def set_header(word, sec, kind, show):
    h = sec.Headers(kind)  # 1 = primary, 2 = first page
    h.LinkToPrevious = False
    h.Range.Text = ""
    if show:
        h.Range.Fields.Add(h.Range, 33)  # wdFieldPage
    h.Range.ParagraphFormat.Alignment = 2  # right
    h.Range.Font.Name = "TH SarabunPSK"
    h.Range.Font.NameBi = "TH SarabunPSK"
    h.Range.Font.Size = 16
    h.Range.Font.SizeBi = 16


def merge_and_number(n_front_sections):
    shutil.copy2(BODY, BOOK)  # body is the container: InsertFile keeps the container's styles
    word = word_app()
    try:
        d = word.Documents.Open(BOOK)
        d.Range(0, 0).InsertBreak(2)  # empty first section to receive the front matter
        d.Range(0, 0).InsertFile(FRONT)
        assert d.Sections.Count > n_front_sections
        thai_style = d.Sections(5).Headers(1).PageNumbers.NumberStyle  # abstract section: thaiLetters
        for i in range(1, d.Sections.Count + 1):
            s = d.Sections(i)
            s.PageSetup.DifferentFirstPageHeaderFooter = True
            if s.PageSetup.TopMargin > 100:  # not on the 2.54 cm cover: the header would push the text down
                s.PageSetup.HeaderDistance = 72  # 2.54 cm from the top edge
            pn = s.Headers(1).PageNumbers
            if i <= 7:  # sample: cover none, ก-ค hidden, ง จ ฉ shown
                show_first = show_rest = i >= 5
            elif i <= n_front_sections:  # TOC lists: first page hidden
                show_first, show_rest = False, True
                pn.NumberStyle = thai_style
                pn.RestartNumberingAtSection = False
            else:
                first = s.Range.Paragraphs(1).Range
                starts_chapter = first.ParagraphFormat.OutlineLevel == 1 or first.Text.strip() == "บรรณานุกรม"
                show_first, show_rest = not starts_chapter, True
                pn.NumberStyle = 0  # arabic
                pn.RestartNumberingAtSection = i == n_front_sections + 1
                pn.StartingNumber = 1
            set_header(word, s, 2, show_first)
            set_header(word, s, 1, show_rest)
        d.Repaginate()
        fit_one_page_sections(d)  # again in the final document; covers carry no page numbers
        for si in range(1, 5):
            rng = d.Sections(si).Range
            assert rng.Information(3) == rng.Paragraphs(1).Range.Information(3), f"section {si} > 1 page"
        d.Save()
        d.Close(False)
    finally:
        word.Quit()

    # Clean up empty heading paragraph created at junction of front matter & Chapter 1
    doc = _open(BOOK)
    for p in doc.paragraphs:
        if not p.text.strip() and p.style.name in ("Heading 1", "FrontNormal"):
            p.style = "Normal"
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
    normalize_figures(doc)
    _save(doc, BOOK)


def main(argv):
    if "--rebuild-body" in argv:
        import build_ch1_ch5
        build_ch1_ch5.main()
    os.makedirs(TMP, exist_ok=True)
    entries, last = build_body()
    print(f"body: {last} pages, {len(entries['toc'])} toc / {len(entries['tables'])} tables / {len(entries['figures'])} figures")
    letters = {}
    for _ in range(3):  # list lengths decide later letters; settle in 2 rounds
        build_front(entries, letters)
        new = paginate_front()
        if new == letters:
            break
        letters = new
    print("front letters:", letters)
    n_front = len(_open(FRONT).sections)
    merge_and_number(n_front)
    pdf = BOOK[:-5] + ".pdf"
    export_pdf(BOOK, pdf)
    print("pdf:", pdf, render_pngs(pdf, "book"), "pages")


if __name__ == "__main__":
    main(sys.argv)
