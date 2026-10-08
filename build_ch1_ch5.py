import os, sys, shutil, copy
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import win32com.client
import pymupdf

sys.stdout.reconfigure(encoding='utf-8')

ROOT = os.path.abspath(os.path.dirname(__file__))
TEMP_DIR = os.path.join(ROOT, "process", "temp")
os.makedirs(TEMP_DIR, exist_ok=True)

OUT_DOCX = os.path.join(ROOT, "output", "Chatshop_บทที่1_5.docx")
OUT_PDF = os.path.join(ROOT, "output", "Chatshop_บทที่1_5.pdf")

def mark_thai(run):
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

def split_docx_table(doc, tbl_index, split_after_row, cont_caption_text):
    """
    Splits table at doc.tables[tbl_index]:
    Part 1 keeps row 0 (header) and rows 1 to split_after_row.
    Between Part 1 and 2, inserts a continuation caption paragraph with page_break_before.
    Part 2 keeps row 0 (header) and rows split_after_row + 1 to end.
    """
    t = doc.tables[tbl_index]
    parent = t._tbl.getparent()
    idx = parent.index(t._tbl)

    tbl2_elem = copy.deepcopy(t._tbl)
    t2 = docx.table.Table(tbl2_elem, doc)

    # In Part 1, remove rows from (split_after_row + 1) to end
    rows_to_keep_part1 = split_after_row + 1  # header (0) + rows (1..split_after_row)
    for _ in range(len(t.rows) - rows_to_keep_part1):
        t._tbl.remove(t.rows[rows_to_keep_part1]._tr)

    # In Part 2, remove data rows 1..split_after_row (keep row 0 as repeated header)
    for _ in range(split_after_row):
        t2._tbl.remove(t2.rows[1]._tr)

    # Continuation paragraph
    p_cont = docx.text.paragraph.Paragraph(OxmlElement("w:p"), doc)
    p_cont.style = "Caption"
    p_cont.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_cont.paragraph_format.space_before = Pt(0)
    p_cont.paragraph_format.space_after = Pt(0)
    p_cont.paragraph_format.keep_with_next = True
    p_cont.paragraph_format.page_break_before = True

    r = p_cont.add_run(cont_caption_text)
    r.bold = True
    mark_thai(r)

    # Insert continuation paragraph and Part 2 table
    parent.insert(idx + 1, p_cont._p)
    parent.insert(idx + 2, tbl2_elem)
    print(f"Table {tbl_index} split: Part 1={len(t.rows)} rows, Part 2={len(t2.rows)} rows")

def prepare_ch3():
    src = os.path.join(ROOT, "output", "Chatshop_บทที่3_วิธีดำเนินงาน.docx")
    dst = os.path.join(TEMP_DIR, "ch3_prepared.docx")
    print(f"Preparing Chapter 3 from {src}...")
    doc = docx.Document(src)

    # Table 3.1 is doc.tables[3] (18 rows). Page 32 fits header + 3 data rows (tenants, tenant_users, tenant_roles)
    # Part 1: rows 0..3 (4 rows). Part 2: header + rows 4..17 (15 rows)
    split_docx_table(doc, 3, 3, "ตารางที่ 3.1  ตารางในฐานข้อมูล MySQL (ต่อ)")

    # Table 3.3 is doc.tables[10] (was index 9 originally, now index 10 because a new table was inserted)
    # Page 39 fits header + 8 data rows. Split after row 8
    split_docx_table(doc, 10, 8, "ตารางที่ 3.3  หมวดของ API (ต่อ)")

    doc.save(dst)
    print(f"Saved prepared Chapter 3: {dst}")
    return dst

def prepare_ch4():
    src = os.path.join(ROOT, "output", "Chatshop_บทที่4_ผลการดำเนินงาน.docx")
    dst = os.path.join(TEMP_DIR, "ch4_prepared.docx")
    print(f"Preparing Chapter 4 from {src}...")
    doc = docx.Document(src)

    # Table 4.1 is doc.tables[0] (9 rows). Split after row 2 (header + Test Case ID + Test Scenario)
    # Part 1: rows 0..2 (3 rows). Part 2: header + rows 3..8 (7 rows)
    split_docx_table(doc, 0, 2, "ตารางที่ 4.1  ขอบเขตข้อมูลที่ใช้ในตารางกรณีทดสอบ (ต่อ)")

    # Test cases 4.1 to 4.17: Set page_break_before on Heading 2 so each test case starts on a new page cleanly
    for p in doc.paragraphs:
        txt = p.text.strip()
        if p.style.name == "Heading 2" and ("ผลการทดสอบ" in txt or "ข้อผิดพลาด" in txt or "แบบประเมิน" in txt):
            p.paragraph_format.page_break_before = True

    doc.save(dst)
    print(f"Saved prepared Chapter 4: {dst}")
    return dst

def prepare_ch5():
    src = os.path.join(ROOT, "output", "Chatshop_บทที่5_สรุปผลและข้อเสนอแนะ.docx")
    dst = os.path.join(TEMP_DIR, "ch5_prepared.docx")
    print(f"Preparing Chapter 5 from {src}...")
    doc = docx.Document(src)

    # Verify all runs are tagged Thai
    for p in doc.paragraphs:
        for r in p.runs:
            mark_thai(r)

    doc.save(dst)
    print(f"Saved prepared Chapter 5: {dst}")
    return dst

def main():
    print("=" * 60)
    print("Step 1: Preparing chapters with strict thesis rules...")
    print("=" * 60)
    
    # Ensure Chapter 1-2 is up to date
    ch1_2_docx = os.path.join(ROOT, "output", "Chatshop_บทที่1_2.docx")
    if not os.path.exists(ch1_2_docx):
        import build_ch1_ch2
        build_ch1_ch2.main()
        
    ch3_docx = prepare_ch3()
    ch4_docx = prepare_ch4()
    ch5_docx = prepare_ch5()

    print("\n" + "=" * 60)
    print("Step 2: Merging Chapters 1-5 via Word COM...")
    print("=" * 60)

    print(f"Base file: {ch1_2_docx} -> {OUT_DOCX}")
    shutil.copy2(ch1_2_docx, OUT_DOCX)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(OUT_DOCX))

        # Merge Chapter 3
        print("Appending Chapter 3...")
        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd
        rng.InsertBreak(2)  # wdSectionBreakNextPage
        rng.Collapse(0)
        rng.InsertFile(os.path.abspath(ch3_docx))

        # Merge Chapter 4
        print("Appending Chapter 4...")
        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd
        rng.InsertBreak(2)  # wdSectionBreakNextPage
        rng.Collapse(0)
        rng.InsertFile(os.path.abspath(ch4_docx))

        # Merge Chapter 5
        print("Appending Chapter 5...")
        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd
        rng.InsertBreak(2)  # wdSectionBreakNextPage
        rng.Collapse(0)
        rng.InsertFile(os.path.abspath(ch5_docx))

        print("Enforcing section margins on all sections...")
        for i in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(i)
            # RMUTI thesis margin: Top 3.81 cm (108 pt), Bottom 2.54 cm (72 pt), Left 3.81 cm (108 pt), Right 2.54 cm (72 pt)
            sec.PageSetup.TopMargin = 108
            sec.PageSetup.BottomMargin = 72
            sec.PageSetup.LeftMargin = 108
            sec.PageSetup.RightMargin = 72

        print("Updating fields and repaginating...")
        doc.Fields.Update()
        doc.Repaginate()
        doc.Save()

        print(f"Exporting PDF: {OUT_PDF}...")
        doc.ExportAsFixedFormat(os.path.abspath(OUT_PDF), 17)
        total_pages = doc.ComputeStatistics(2)
        print(f"Successfully generated Chatshop_บทที่1_5.docx and PDF (Total pages: {total_pages})")
        doc.Close(False)
    finally:
        word.Quit()

    print("\n" + "=" * 60)
    print("Step 3: Rendering key verification pages...")
    print("=" * 60)
    pdf = pymupdf.open(OUT_PDF)
    print(f"Total PDF pages: {len(pdf)}")

    # Search for specific pages
    p_fig2_1 = None
    p_tbl3_1_cont = None
    p_tbl4_1 = None
    p_tbl4_1_cont = None
    p_ch5_start = None

    for i, page in enumerate(pdf):
        txt = page.get_text()
        if not p_fig2_1 and "รูปที่ 2.1 " in txt:
            p_fig2_1 = i + 1
        if not p_tbl3_1_cont and "ตารางที่ 3.1" in txt and "(ต่อ)" in txt:
            p_tbl3_1_cont = i + 1
        if not p_tbl4_1 and "ตารางที่ 4.1" in txt and "(ต่อ)" not in txt:
            p_tbl4_1 = i + 1
        if not p_tbl4_1_cont and "ตารางที่ 4.1" in txt and "(ต่อ)" in txt:
            p_tbl4_1_cont = i + 1
        if not p_ch5_start and "บทที่ 5" in txt and "สรุปผล และข้อเสนอแนะ" in txt:
            p_ch5_start = i + 1

    print(f"Key pages found: Fig 2.1=p{p_fig2_1}, Tbl 3.1(ต่อ)=p{p_tbl3_1_cont}, Tbl 4.1=p{p_tbl4_1}, Tbl 4.1(ต่อ)=p{p_tbl4_1_cont}, Ch5 Start=p{p_ch5_start}")
    
    render_pages = [p for p in [1, p_fig2_1, p_tbl3_1_cont, p_tbl4_1, p_tbl4_1_cont, p_ch5_start, len(pdf)] if p]
    for pno in render_pages:
        if 1 <= pno <= len(pdf):
            img_path = os.path.join(TEMP_DIR, f"master_ch1_5_p{pno}.png")
            pdf[pno - 1].get_pixmap(dpi=130).save(img_path)
            print(f"Rendered page {pno} -> {img_path}")

    print("\nMaster document build complete!")

if __name__ == "__main__":
    main()
