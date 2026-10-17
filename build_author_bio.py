"""Builder for Chatshop Author Bio (output/Chatshop_ประวัติผู้เขียน.docx & .pdf).
Follows the format of the senior thesis reference.
"""
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import win32com.client

ROOT = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(ROOT, "input", "reference", "rmuti_thesis_template.docx")
OUT_DIR = os.path.join(ROOT, "output")
os.makedirs(OUT_DIR, exist_ok=True)


def mark_thai(run, font_size=16, bold=False, italic=False):
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    for tag in ("w:rFonts", "w:cs", "w:sz", "w:szCs", "w:lang"):
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

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(font_size * 2)))
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(font_size * 2)))

    rPr.insert(0, fonts)
    rPr.append(sz)
    rPr.append(szCs)
    rPr.append(OxmlElement("w:cs"))
    rPr.append(lang)


def create_base_doc():
    doc = docx.Document(TEMPLATE)
    for el in list(doc.element.body):
        if el.tag != qn("w:sectPr"):
            doc.element.body.remove(el)
    s = doc.sections[0]
    s.different_first_page_header_footer = True
    s.top_margin = Cm(5.08)  # first page of bio = 2 inches
    s.bottom_margin = Cm(2.54)
    s.left_margin = Cm(3.81)
    s.right_margin = Cm(2.54)
    return doc


def add_p(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, size=16, space_b=0, space_a=0, line_spacing=1.0):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_b)
    pf.space_after = Pt(space_a)
    pf.line_spacing = line_spacing
    if text:
        r = p.add_run(text)
        mark_thai(r, font_size=size, bold=bold, italic=italic)
    return p


def build_author_bio():
    doc = create_base_doc()

    # Heading
    add_p(doc, "ประวัติผู้เขียน", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_b=0, space_a=18)

    # 2-column table for info (left) and photo frame (right)
    tbl = doc.add_table(rows=1, cols=2)
    tblPr = tbl._tbl.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for b in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{b}")
        el.set(qn("w:val"), "none")
        tblBorders.append(el)
    tblPr.append(tblBorders)

    row = tbl.rows[0]
    cell_info, cell_photo = row.cells[0], row.cells[1]
    cell_info.width = Cm(10.5)
    cell_photo.width = Cm(4.15)

    # Personal info
    info_items = [
        ("ชื่อ-นามสกุล", "นายธนพล  วุฒิธาดา"),
        ("วัน-เดือน-ปีเกิด", "4 สิงหาคม 2545"),
        ("ที่อยู่", "73 หมู่ 1 ต.ปากตม อ.เชียงคาน จ.เลย 42110"),
        ("อีเมล", "thanaphol.wu@rmuti.ac.th"),
        ("เบอร์โทรศัพท์", "096-878-3234"),
    ]
    for idx, (label, val) in enumerate(info_items):
        p = cell_info.paragraphs[0] if idx == 0 else cell_info.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        r1 = p.add_run(f"{label} : ")
        mark_thai(r1, font_size=16, bold=True)
        r2 = p.add_run(val)
        mark_thai(r2, font_size=16)

    # Photo placeholder box (1.5 x 2 inches: 3.6 cm x 4.9 cm)
    p_ph = cell_photo.paragraphs[0]
    p_ph.paragraph_format.space_before = Pt(0)
    p_ph.paragraph_format.space_after = Pt(0)
    photo_tbl = cell_photo.add_table(rows=1, cols=1)
    photo_tbl_pr = photo_tbl._tbl.tblPr
    photo_jc = OxmlElement("w:jc")
    photo_jc.set(qn("w:val"), "right")
    photo_tbl_pr.append(photo_jc)
    photo_borders = OxmlElement("w:tblBorders")
    for b in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{b}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        photo_borders.append(el)
    for b in ("insideH", "insideV"):
        el = OxmlElement(f"w:{b}")
        el.set(qn("w:val"), "none")
        photo_borders.append(el)
    photo_tbl_pr.append(photo_borders)

    photo_row = photo_tbl.rows[0]
    trPr = photo_row._tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), "2778")  # ~4.9 cm
    trHeight.set(qn("w:hRule"), "atLeast")
    trPr.append(trHeight)
    p_cell = photo_row.cells[0].paragraphs[0]
    p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cell.paragraph_format.space_before = Pt(0)
    p_cell.paragraph_format.space_after = Pt(0)
    photo_row.cells[0].width = Cm(3.6)

    # Education history
    add_p(doc, "ประวัติการศึกษา", align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, size=16, space_b=18, space_a=6)

    edu_items = [
        ("มัธยมศึกษาตอนต้น", "โรงเรียนเลยพิทยาคม ปีการศึกษา 2558"),
        ("ประกาศนียบัตรวิชาชีพ", "วิทยาลัยเทคนิคเลย ปีการศึกษา 2561"),
        ("ประกาศนียบัตรวิชาชีพชั้นสูง", "วิทยาลัยเทคนิคเลย ปีการศึกษา 2564"),
        ("ปริญญาวิศวกรรมศาสตรบัณฑิต", "สาขาวิชาวิศวกรรมคอมพิวเตอร์ คณะวิศวกรรมศาสตร์\nมหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน วิทยาเขตขอนแก่น ปีการศึกษา 2569"),
    ]

    for level, school in edu_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Cm(1.8)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        r_lvl = p.add_run(f"{level} : ")
        mark_thai(r_lvl, font_size=16, bold=True)
        r_sch = p.add_run(school)
        mark_thai(r_sch, font_size=16)

    out_docx = os.path.join(OUT_DIR, "Chatshop_ประวัติผู้เขียน.docx")
    out_pdf = os.path.join(OUT_DIR, "Chatshop_ประวัติผู้เขียน.pdf")
    doc.save(out_docx)
    export_pdf(out_docx, out_pdf)
    print(f"Saved: {out_docx} and {out_pdf}")
    return out_docx


def export_pdf(docx_path, pdf_path):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        d = word.Documents.Open(os.path.abspath(docx_path))
        d.Repaginate()
        d.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)
        d.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    build_author_bio()
