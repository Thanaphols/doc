import os, sys
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import win32com.client
import pymupdf

sys.stdout.reconfigure(encoding='utf-8')

ROOT = os.path.abspath(os.path.dirname(__file__))
TEMPLATE = os.path.join(ROOT, "input", "reference", "rmuti_thesis_template.docx")
OUT_DIR = os.path.join(ROOT, "output")
TEMP_DIR = os.path.join(ROOT, "process", "temp")
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)

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
    s.top_margin = Cm(3.81)
    s.bottom_margin = Cm(2.54)
    s.left_margin = Cm(3.81)
    s.right_margin = Cm(2.54)
    return doc

def add_p(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, size=16, space_b=0, space_a=0, line_spacing=1.0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_b)
    p.paragraph_format.space_after = Pt(space_a)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        r = p.add_run(text)
        mark_thai(r, font_size=size, bold=bold, italic=italic)
    return p

def set_no_borders(table):
    tblPr = table._tbl.tblPr
    existing = tblPr.find(qn("w:tblBorders"))
    if existing is not None:
        tblPr.remove(existing)
    tblBorders = OxmlElement("w:tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "none")
        tblBorders.append(border)
    tblLook = tblPr.find(qn("w:tblLook"))
    if tblLook is not None:
        tblPr.insert(tblPr.index(tblLook), tblBorders)
    else:
        tblPr.append(tblBorders)

# -------------------------------------------------------------
# 1. หน้าปก (Cover Pages: ปกนอก, ปกในไทย, ปกในอังกฤษ, ใบรับรอง)
# -------------------------------------------------------------
def build_cover():
    print("Building หน้าปก (Cover Pages)...")
    doc = create_base_doc()
    
    # --- PAGE 1: ปกนอก ---
    # Top spacing
    add_p(doc, "", space_b=36)
    add_p(doc, "แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ (Chatshop)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "CHATSHOP: CENTRALIZED CHAT AND ORDER MANAGEMENT APPLICATION FOR ONLINE BUSINESS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=120)
    
    add_p(doc, "ธนพล  วุฒิธาดา", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=140)
    
    add_p(doc, "ปริญญานิพนธ์นี้เป็นส่วนหนึ่งของการศึกษาตามหลักสูตรปริญญาวิศวกรรมศาสตรบัณฑิต", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "สาขาวิชาวิศวกรรมคอมพิวเตอร์", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "คณะวิศวกรรมศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน วิทยาเขตขอนแก่น", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "พ.ศ. 2569", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "ลิขสิทธิ์ของคณะวิศวกรรมศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)

    # --- PAGE 2: ปกในไทย ---
    doc.add_page_break()
    add_p(doc, "", space_b=36)
    add_p(doc, "แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ (Chatshop)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=180)
    
    add_p(doc, "ธนพล  วุฒิธาดา", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=200)
    
    add_p(doc, "ปริญญานิพนธ์นี้เป็นส่วนหนึ่งของการศึกษาตามหลักสูตรปริญญาวิศวกรรมศาสตรบัณฑิต", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "สาขาวิชาวิศวกรรมคอมพิวเตอร์", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "คณะวิศวกรรมศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน วิทยาเขตขอนแก่น", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "พ.ศ. 2569", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "ลิขสิทธิ์ของคณะวิศวกรรมศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)

    # --- PAGE 3: ปกในอังกฤษ ---
    doc.add_page_break()
    add_p(doc, "", space_b=36)
    add_p(doc, "Chatshop: Centralized Chat and Order Management Application for Online Business", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=180)
    
    add_p(doc, "Thanaphol  Wuttitada", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=200)
    
    add_p(doc, "A Report Submitted in Partial Fulfillment of the Requirements for", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "the Degree of Bachelor of Engineering", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "Department of Computer Engineering, Faculty of Engineering", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "Rajamangala University of Technology Isan, Khon Kaen Campus", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "2026", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)
    add_p(doc, "© Faculty of Engineering, Rajamangala University of Technology Isan", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16)

    # --- PAGE 4: ใบรับรองปริญญานิพนธ์ ---
    doc.add_page_break()
    # Margins top 8.0 cm -> add space before
    add_p(doc, "", space_b=50)
    
    t_info = doc.add_table(rows=4, cols=2)
    tblBorders = OxmlElement("w:tblBorders")
    for b in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{b}")
        el.set(qn("w:val"), "none")
        tblBorders.append(el)
    t_info._tbl.tblPr.append(tblBorders)
    
    info_data = [
        ("หัวข้อปริญญานิพนธ์  : ", "แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ (Chatshop)"),
        ("จัดทำโดย   : ", "นายธนพล  วุฒิธาดา"),
        ("สาขาวิชา  : ", "วิศวกรรมคอมพิวเตอร์"),
        ("ประธานที่ปรึกษา   : ", "อาจารย์ ดร.ประภาส  ผ่องสนาม")
    ]
    for row_idx, (col1, col2) in enumerate(info_data):
        row = t_info.rows[row_idx]
        p1 = row.cells[0].paragraphs[0]
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(col1)
        mark_thai(r1, bold=True)
        
        p2 = row.cells[1].paragraphs[0]
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after = Pt(2)
        r2 = p2.add_run(col2)
        mark_thai(r2)

    add_p(doc, "", space_b=24)
    add_p(doc, "ได้รับอนุมัติให้เป็นส่วนหนึ่งของการศึกษาตามหลักสูตรปริญญาวิศวกรรมศาสตรบัณฑิต", align=WD_ALIGN_PARAGRAPH.CENTER, size=16)
    add_p(doc, "คณะวิศวกรรมศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน วิทยาเขตขอนแก่น", align=WD_ALIGN_PARAGRAPH.CENTER, size=16, space_a=36)

    add_p(doc, "………………………….……..………................................... คณบดีคณะวิศวกรรมศาสตร์", align=WD_ALIGN_PARAGRAPH.CENTER, size=16)
    add_p(doc, "(ผู้ช่วยศาสตราจารย์ ดร.ศุภฤกษ์  ชามงคลประดิษฐ์) วันที่......................................", align=WD_ALIGN_PARAGRAPH.CENTER, size=16, space_a=36)

    add_p(doc, "คณะกรรมการสอบปริญญานิพนธ์", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=16)
    
    # 2-col committee
    t_comm = doc.add_table(rows=4, cols=2)
    t_comm._tbl.tblPr.append(tblBorders)
    comm_data = [
        (".............................................ประธานกรรมการสอบ", "................................................ประธานที่ปรึกษา"),
        ("(อาจารย์ อรินธร  เจษฎาเมธาขจร)", "(อาจารย์ ดร.ประภาส  ผ่องสนาม)"),
        (".............................................กรรมการสอบ", ".............................................. กรรมการ"),
        ("(อาจารย์ ดร.ปิยะนุช  ตั้งกิตติพล)", "(อาจารย์ ดร.ประภาส  ผ่องสนาม)")
    ]
    for row_idx, (c1, c2) in enumerate(comm_data):
        row = t_comm.rows[row_idx]
        p1 = row.cells[0].paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_after = Pt(12 if row_idx in (1, 3) else 2)
        r1 = p1.add_run(c1)
        mark_thai(r1)
        
        p2 = row.cells[1].paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(12 if row_idx in (1, 3) else 2)
        r2 = p2.add_run(c2)
        mark_thai(r2)

    out_docx = os.path.join(OUT_DIR, "Chatshop_หน้าปก.docx")
    out_pdf = os.path.join(OUT_DIR, "Chatshop_หน้าปก.pdf")
    doc.save(out_docx)
    export_pdf(out_docx, out_pdf)
    print(f"Saved: {out_docx} and {out_pdf}")

# -------------------------------------------------------------
# 2. กิตติกรรมประกาศ (`output/Chatshop_กิตติกรรมประกาศ.docx` & `.pdf`)
# -------------------------------------------------------------
def build_acknowledgements():
    print("Building กิตติกรรมประกาศ (Acknowledgements)...")
    doc = create_base_doc()
    
    # Page number ฉ top-right
    p_num = add_p(doc, "ฉ", align=WD_ALIGN_PARAGRAPH.RIGHT, size=16)
    p_num.paragraph_format.space_after = Pt(18)

    # Heading
    h = add_p(doc, "กิตติกรรมประกาศ", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=18)
    
    # Content
    body_text_1 = (
        "\tปริญญานิพนธ์ฉบับนี้สำเร็จลุล่วงด้วยความกรุณาและความเอาใจใส่จากหลายภาคส่วน "
        "ผู้จัดทำขอกราบขอบพระคุณ อาจารย์ ดร.ประภาส ผ่องสนาม อาจารย์ที่ปรึกษาปริญญานิพนธ์ "
        "ที่ได้ให้คำแนะนำ ชี้แนะแนวทางการดำเนินงาน และถ่ายทอดองค์ความรู้ด้านการพัฒนาระบบอย่างใกล้ชิด "
        "ด้วยความเอื้ออาทรและอดทนตลอดระยะเวลาของการทำโครงงาน รวมถึงขอขอบพระคุณ คณะกรรมการสอบปริญญานิพนธ์ ทุกท่าน "
        "ที่ได้ให้ข้อเสนอแนะและข้อคิดเห็นอันทรงคุณค่า ซึ่งมีส่วนสำคัญในการพัฒนาโครงงานให้มีความสมบูรณ์ยิ่งขึ้น"
    )
    p1 = add_p(doc, body_text_1, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, size=16, space_a=6, line_spacing=1.0)

    body_text_2 = (
        "\tผู้จัดทำขอแสดงความขอบคุณอย่างยิ่งต่อ สาขาวิชาวิศวกรรมคอมพิวเตอร์ คณะวิศวกรรมศาสตร์ "
        "มหาวิทยาลัยเทคโนโลยีราชมงคลอีสาน วิทยาเขตขอนแก่น ที่ให้โอกาสทางการศึกษาและส่งเสริมการเรียนรู้ตลอดหลักสูตร "
        "ทั้งด้านวิชาการและภาคปฏิบัติ ซึ่งมีส่วนสำคัญในการพัฒนาทักษะและองค์ความรู้ที่สามารถนำมาประยุกต์ใช้ในการทำโครงงานครั้งนี้ได้อย่างมีประสิทธิภาพ"
    )
    p2 = add_p(doc, body_text_2, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, size=16, space_a=6, line_spacing=1.0)

    body_text_3 = (
        "\tประโยชน์และคุณค่าอันพึงมีจากปริญญานิพนธ์ฉบับนี้ ผู้จัดทำโครงงานขอมอบเป็นกตัญญุตาบูชาแด่ "
        "บิดามารดา ครูอาจารย์ ตลอดจนผู้มีพระคุณทุกท่าน ที่คอยอบรมสั่งสอน ให้กำลังใจ และให้การสนับสนุนเสมอมา"
    )
    p3 = add_p(doc, body_text_3, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, size=16, space_a=48, line_spacing=1.0)

    # Signature
    add_p(doc, "\t\t\t\t\t\t\t\t  ธนพล  วุฒิธาดา", align=WD_ALIGN_PARAGRAPH.LEFT, size=16)

    out_docx = os.path.join(OUT_DIR, "Chatshop_กิตติกรรมประกาศ.docx")
    out_pdf = os.path.join(OUT_DIR, "Chatshop_กิตติกรรมประกาศ.pdf")
    doc.save(out_docx)
    export_pdf(out_docx, out_pdf)
    print(f"Saved: {out_docx} and {out_pdf}")

# -------------------------------------------------------------
# 3. บทคัดย่อ (Abstract Thai & English)
# -------------------------------------------------------------
def build_abstract():
    print("Building บทคัดย่อ (Abstract)...")
    doc = create_base_doc()
    
    # Page ง - Thai Abstract
    add_p(doc, "ง", align=WD_ALIGN_PARAGRAPH.RIGHT, size=16, space_a=18)
    
    t_info = doc.add_table(rows=5, cols=2)
    set_no_borders(t_info)
    
    info_th = [
        ("หัวข้อปริญญานิพนธ์", "แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ (Chatshop)"),
        ("จัดทำโดย", "นายธนพล  วุฒิธาดา"),
        ("ปีที่ปริญญานิพนธ์สำเร็จ", "พ.ศ. 2569"),
        ("สาขาวิชา", "วิศวกรรมคอมพิวเตอร์"),
        ("ประธานที่ปรึกษา", "อาจารย์ ดร.ประภาส  ผ่องสนาม")
    ]
    for row_idx, (col1, col2) in enumerate(info_th):
        row = t_info.rows[row_idx]
        p1 = row.cells[0].paragraphs[0]
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(col1)
        mark_thai(r1, bold=True)
        
        p2 = row.cells[1].paragraphs[0]
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after = Pt(2)
        r2 = p2.add_run(f": {col2}")
        mark_thai(r2)

    add_p(doc, "บทคัดย่อ", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_b=12, space_a=12)
    
    abs_th_1 = (
        "\tโครงงานนี้นำเสนอการพัฒนาแอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ (Chatshop) "
        "เพื่อแก้ไขปัญหาความยุ่งยากในการบริหารจัดการข้อความของร้านค้าออนไลน์ที่ติดต่อสื่อสารกับลูกค้าผ่านหลากหลายแพลตฟอร์ม "
        "อาทิ LINE Official Account, Facebook Messenger, Instagram และ WebChat ซึ่งทำให้เกิดปัญหาการตอบกลับล่าช้า "
        "ข้อความตกหล่น และข้อมูลคำสั่งซื้อกระจัดกระจาย โดยระบบประกอบด้วยระบบบริการหลังบ้านที่พัฒนาด้วย Laravel 10 "
        "ร่วมกับฐานข้อมูล MySQL และ MongoDB สำหรับประมวลผล Webhook ขาเข้า และแอปพลิเคชันบนสมาร์ตโฟนที่พัฒนาด้วย Flutter "
        "ที่ช่วยให้เจ้าหน้าที่สามารถติดตามและตอบกลับข้อความจากทุกช่องทางได้ภายในหน้าต่างเดียวแบบเรียลไทม์"
    )
    p = add_p(doc, abs_th_1, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, size=16, space_a=6)

    abs_th_2 = (
        "\tนอกจากนี้ ระบบยังได้บูรณาการระบบแชทบอทอัจฉริยะที่รองรับการค้นคืนข้อมูลจากฐานความรู้ (Retrieval-Augmented Generation: RAG) "
        "เพื่อช่วยแนะนำคำตอบและตอบกลับลูกค้าอัตโนมัติ พร้อมทั้งฟังก์ชันการแจ้งเตือนแชทที่ค้างตอบเกินกำหนด "
        "และการเปิดคำสั่งซื้อสินค้าได้โดยตรงจากหน้าต่างสนทนา จากผลการทดสอบระบบและการประเมินความพึงพอใจจากกลุ่มผู้ใช้งานจริงจำนวน 15 คน "
        "พบว่าระบบได้รับความพึงพอใจในภาพรวมอยู่ในระดับดี โดยมีค่าเฉลี่ยรวมเท่ากับ 4.10 และค่าส่วนเบี่ยงเบนมาตรฐานเท่ากับ 0.37 "
        "ซึ่งแสดงให้เห็นว่าระบบสามารถช่วยเพิ่มประสิทธิภาพและลดภาระในการบริหารจัดการร้านค้าออนไลน์ได้อย่างเป็นรูปธรรม"
    )
    p = add_p(doc, abs_th_2, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, size=16, space_a=8)

    add_p(doc, "คำสำคัญ : แอปพลิเคชันสมาร์ตโฟน, รวมศูนย์ข้อความ, จัดการคำสั่งซื้อ, แชทบอท, Flutter, Laravel", bold=True, size=16)

    # Page จ - English Abstract
    doc.add_page_break()
    p_num_en = add_p(doc, "จ", align=WD_ALIGN_PARAGRAPH.RIGHT, size=16, space_a=12)
    
    t_info_en = doc.add_table(rows=5, cols=2)
    set_no_borders(t_info_en)
    info_en = [
        ("Project Title", "Chatshop: Centralized Chat and Order Management Application for Online Business"),
        ("Proposed by", "Mr. Thanaphol  Wuttitada"),
        ("Year", "2026"),
        ("Department", "Computer Engineering"),
        ("Advisor", "Prapas  Phongsanam, D.Eng.")
    ]
    for row_idx, (col1, col2) in enumerate(info_en):
        row = t_info_en.rows[row_idx]
        p1 = row.cells[0].paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(1)
        r1 = p1.add_run(col1)
        mark_thai(r1, bold=True)
        
        p2 = row.cells[1].paragraphs[0]
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(1)
        r2 = p2.add_run(f": {col2}")
        mark_thai(r2)

    add_p(doc, "Abstract", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_b=8, space_a=8)
    
    abs_en_1 = (
        "\tThis project presents the development of Chatshop, a centralized messaging and order management mobile application for online businesses. "
        "The system aims to address the operational challenges of managing omnichannel customer inquiries across platforms such as LINE Official Account, "
        "Facebook Messenger, Instagram, and WebChat, which typically lead to delayed responses and fragmented order data. "
        "The architecture comprises a robust backend service developed using Laravel 10 with MySQL and MongoDB databases for real-time webhook ingestion, "
        "and a cross-platform mobile application developed using Flutter, enabling store staff to manage all communications within a single unified inbox."
    )
    p = add_p(doc, abs_en_1, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, size=16, space_a=4)

    abs_en_2 = (
        "\tFurthermore, the system integrates an intelligent chatbot equipped with Retrieval-Augmented Generation (RAG) to provide automated responses "
        "and suggested replies, along with overdue response alert notifications and integrated chat-order creation features. "
        "The functional test results and user satisfaction evaluation conducted with 15 representative users indicated a high level of satisfaction, "
        "with an overall mean score of 4.10 and a standard deviation of 0.37. "
        "These results demonstrate that Chatshop significantly enhances workflow efficiency and streamlines communication for modern online merchants."
    )
    p = add_p(doc, abs_en_2, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, size=16, space_a=6)

    add_p(doc, "Keywords : Mobile Application, Centralized Messaging, Order Management, Chatbot, Flutter, Laravel", bold=True, size=16)

    out_docx = os.path.join(OUT_DIR, "Chatshop_บทคัดย่อ.docx")
    out_pdf = os.path.join(OUT_DIR, "Chatshop_บทคัดย่อ.pdf")
    doc.save(out_docx)
    export_pdf(out_docx, out_pdf)
    print(f"Saved: {out_docx} and {out_pdf}")

# -------------------------------------------------------------
# 4. สารบัญ (TOC, List of Tables, List of Figures)
# -------------------------------------------------------------
def add_toc_line(doc, title, page_str, level=0, is_header=False, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.0

    # Tab stop: right tab at 14.65 cm with dots leader
    leader = WD_TAB_LEADER.SPACES if is_header else WD_TAB_LEADER.DOTS
    p.paragraph_format.tab_stops.add_tab_stop(Cm(14.65), WD_TAB_ALIGNMENT.RIGHT, leader)

    prefix = "\t" * level if level > 0 else ""
    r_title = p.add_run(prefix + title)
    mark_thai(r_title, bold=bold)
    
    r_tab = p.add_run("\t")
    mark_thai(r_tab)
    
    r_page = p.add_run(str(page_str))
    mark_thai(r_page, bold=bold)
    return p

def build_toc():
    print("Building สารบัญ (Table of Contents)...")
    doc = create_base_doc()

    # --- SECTION 1: สารบัญเนื้อหา ---
    add_p(doc, "สารบัญ", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=12)
    add_toc_line(doc, "", "หน้า", is_header=True, bold=True)
    
    toc_items = [
        ("บทคัดย่อภาษาไทย", "ง", 0, False),
        ("Abstract", "จ", 0, False),
        ("กิตติกรรมประกาศ", "ฉ", 0, False),
        ("สารบัญตาราง", "ญ", 0, False),
        ("สารบัญรูป", "ฏ", 0, False),
        ("บทที่ 1 บทนำ", "1", 0, True),
        ("1.1 ความเป็นมาและความสำคัญของปัญหา", "1", 1, False),
        ("1.2 วัตถุประสงค์ของโครงงาน", "2", 1, False),
        ("1.3 ขอบเขตของโครงงาน", "2", 1, False),
        ("1.4 ประโยชน์ที่คาดว่าจะได้รับ", "3", 1, False),
        ("1.5 นิยามศัพท์เฉพาะ", "3", 1, False),
        ("บทที่ 2 ทฤษฎีและงานวิจัยที่เกี่ยวข้อง", "5", 0, True),
        ("2.1 Flutter", "5", 1, False),
        ("2.2 ภาษา Dart", "6", 1, False),
        ("2.3 Flutter Riverpod", "7", 1, False),
        ("2.4 go_router", "8", 1, False),
        ("2.5 Laravel Framework", "9", 1, False),
        ("2.6 สถาปัตยกรรม REST API", "10", 1, False),
        ("2.7 โพรโทคอล OAuth 2.0", "12", 1, False),
        ("2.8 LINE Messaging API", "13", 1, False),
        ("2.9 Meta Graph API", "14", 1, False),
        ("2.10 Webhook", "15", 1, False),
        ("2.11 คิวงานและการประมวลผลเบื้องหลัง", "16", 1, False),
        ("2.12 Laravel Reverb", "17", 1, False),
        ("2.13 Firebase Cloud Messaging", "18", 1, False),
        ("2.14 ฐานข้อมูล MySQL", "19", 1, False),
        ("2.15 ฐานข้อมูล MongoDB", "20", 1, False),
        ("2.16 สถาปัตยกรรมแบบหลายผู้เช่า", "21", 1, False),
        ("2.17 การควบคุมการเข้าถึงตามบทบาท", "22", 1, False),
        ("2.18 แชทบอท", "23", 1, False),
        ("2.19 แบบจำลองภาษาขนาดใหญ่", "24", 1, False),
        ("2.20 การค้นคืนข้อมูลร่วมกับการสร้างข้อความ", "25", 1, False),
        ("2.21 เวกเตอร์และเวกเตอร์ฝังตัว", "26", 1, False),
        ("2.22 งานวิจัยที่เกี่ยวข้อง", "27", 1, False),
        ("บทที่ 3 วิธีดำเนินงาน", "28", 0, True),
        ("3.1 การวิเคราะห์ความต้องการของระบบ", "28", 1, False),
        ("3.2 การออกแบบสถาปัตยกรรมระบบ", "29", 1, False),
        ("3.3 การออกแบบฐานข้อมูล", "32", 1, False),
        ("3.4 การประมวลผลข้อความขาเข้าผ่าน Webhook", "35", 1, False),
        ("3.5 กระบวนการตอบกลับของแชทบอท", "36", 1, False),
        ("3.6 การออกแบบ API", "39", 1, False),
        ("3.7 การพัฒนาแอปพลิเคชันบนสมาร์ตโฟน", "42", 1, False),
        ("3.8 การตั้งค่าแพลตฟอร์มภายนอก", "56", 1, False),
        ("บทที่ 4 ผลการดำเนินงาน", "65", 0, True),
        ("4.1 ผลการทดสอบการเข้าสู่ระบบด้วยอีเมลและรหัสผ่าน", "67", 1, False),
        ("4.2 ผลการทดสอบการเข้าสู่ระบบด้วยบัญชี Facebook และ Google", "68", 1, False),
        ("4.3 ผลการทดสอบการตรวจลายเซ็นของ Webhook จากแพลตฟอร์มภายนอก", "69", 1, False),
        ("4.4 ผลการทดสอบการส่งข้อความเกินกรอบเวลา 24 ชั่วโมงของ Meta", "70", 1, False),
        ("4.5 ผลการทดสอบการล็อกบทสนทนาเพื่อป้องกันการตอบซ้ำซ้อน", "71", 1, False),
        ("4.6 ผลการทดสอบการแจ้งเตือนแชทที่ค้างตอบเกินกำหนด", "72", 1, False),
        ("4.7 ผลการทดสอบการรวมข้อความที่ลูกค้าส่งติดกันก่อนแชทบอทตอบ", "73", 1, False),
        ("4.8 ผลการทดสอบการส่งต่อบทสนทนาเมื่อแชทบอทตอบไม่ได้", "74", 1, False),
        ("4.9 ผลการทดสอบการป้องกันการตอบซ้ำเมื่องานถูกประมวลผลซ้ำ", "75", 1, False),
        ("4.10 ผลการทดสอบการค้นคืนส่วนย่อยของเอกสารความรู้", "76", 1, False),
        ("4.11 ผลการทดสอบการเปิดและแก้ไขคำสั่งซื้อจากบทสนทนา", "77", 1, False),
        ("4.12 ผลการทดสอบการแยกข้อมูลระหว่างร้านค้า", "78", 1, False),
        ("4.13 ผลการทดสอบการควบคุมสิทธิ์ตามบทบาท", "79", 1, False),
        ("4.14 ผลการทดสอบการรับเหตุการณ์แบบเรียลไทม์บนแอปพลิเคชัน", "80", 1, False),
        ("4.15 ผลการทดสอบการตอบกลับโดยอ้างอิงข้อความและการส่งไฟล์สื่อ", "81", 1, False),
        ("4.16 ผลการทดสอบการจัดการไฟล์สื่อขาเข้า", "82", 1, False),
        ("4.17 ข้อผิดพลาดที่พบจากการทดสอบ", "83", 1, False),
        ("4.18 แบบประเมินความพึงพอใจในการใช้งานแอปพลิเคชัน", "84", 1, False),
        ("บทที่ 5 สรุปผล และข้อเสนอแนะ", "90", 0, True),
        ("5.1 สรุปผลการดำเนินงาน", "90", 1, False),
        ("5.2 อภิปรายผล", "91", 1, False),
        ("5.3 ปัญหาและอุปสรรค", "91", 1, False),
        ("5.4 ข้อเสนอแนะและแนวทางในการพัฒนาต่อ", "92", 1, False),
        ("บรรณานุกรม", "93", 0, True)
    ]
    for title, pno, lvl, bld in toc_items:
        add_toc_line(doc, title, pno, level=lvl, bold=bld)

    # --- SECTION 2: สารบัญตาราง ---
    doc.add_page_break()
    add_p(doc, "สารบัญตาราง", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=12)
    add_toc_line(doc, "ตารางที่", "หน้า", is_header=True, bold=True)
    
    tables_list = [
        ("3.1 ตารางในฐานข้อมูล MySQL", "32"),
        ("3.2 คอลเลกชันในฐานข้อมูล MongoDB", "33"),
        ("3.3 หมวดของ API", "39"),
        ("3.4 แอปของ Meta ที่ใช้ในระบบ", "57"),
        ("3.5 Valid OAuth Redirect URIs ของแอปเชื่อมต่อช่องทาง", "59"),
        ("3.6 สิทธิ์ที่ขอใน Configuration connect", "60"),
        ("4.1 ขอบเขตข้อมูลที่ใช้ในตารางกรณีทดสอบ", "65"),
        ("4.2 การทดสอบการเข้าสู่ระบบด้วยอีเมลและรหัสผ่าน", "67"),
        ("4.3 การทดสอบการเข้าสู่ระบบด้วยบัญชี Facebook และ Google", "68"),
        ("4.4 การทดสอบการตรวจลายเซ็นของ Webhook จากแพลตฟอร์มภายนอก", "69"),
        ("4.5 การทดสอบการส่งข้อความเกินกรอบเวลา 24 ชั่วโมงของ Meta", "70"),
        ("4.6 การทดสอบการล็อกบทสนทนาเพื่อป้องกันการตอบซ้ำซ้อน", "71"),
        ("4.7 การทดสอบการแจ้งเตือนแชทที่ค้างตอบเกินกำหนด", "72"),
        ("4.8 การทดสอบการรวมข้อความที่ลูกค้าส่งติดกันก่อนแชทบอทตอบ", "73"),
        ("4.9 การทดสอบการส่งต่อบทสนทนาเมื่อแชทบอทตอบไม่ได้", "74"),
        ("4.10 การทดสอบการป้องกันการตอบซ้ำเมื่องานถูกประมวลผลซ้ำ", "75"),
        ("4.11 การทดสอบการค้นคืนส่วนย่อยของเอกสารความรู้", "76"),
        ("4.12 การทดสอบการเปิดและแก้ไขคำสั่งซื้อจากบทสนทนา", "77"),
        ("4.13 การทดสอบการแยกข้อมูลระหว่างร้านค้า", "78"),
        ("4.14 การทดสอบการควบคุมสิทธิ์ตามบทบาท", "79"),
        ("4.15 การทดสอบการรับเหตุการณ์แบบเรียลไทม์บนแอปพลิเคชัน", "80"),
        ("4.16 การทดสอบการตอบกลับโดยอ้างอิงข้อความและการส่งไฟล์สื่อ", "81"),
        ("4.17 การทดสอบการจัดการไฟล์สื่อขาเข้า", "82"),
        ("4.18 ผลการประเมินความพึงพอใจด้านฟังก์ชันการทำงานและความสามารถของระบบ", "84"),
        ("4.19 ผลการประเมินความพึงพอใจด้านการออกแบบส่วนต่อประสานและความง่ายในการใช้งาน", "85"),
        ("4.20 ผลการประเมินความพึงพอใจด้านประสิทธิภาพและความเสถียรของระบบ", "86"),
        ("4.21 ผลการประเมินความพึงพอใจด้านประโยชน์และความพึงพอใจในภาพรวม", "87"),
        ("4.22 สรุปผลการประเมินความพึงพอใจในภาพรวมทุกด้าน", "88")
    ]
    for t_name, pno in tables_list:
        add_toc_line(doc, t_name, pno)

    # --- SECTION 3: สารบัญรูป ---
    doc.add_page_break()
    add_p(doc, "สารบัญรูป", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=12)
    add_toc_line(doc, "รูปที่", "หน้า", is_header=True, bold=True)
    
    figures_list = [
        ("2.1 โครงสร้างการทำงานของ Flutter", "5"),
        ("2.2 วงจรชีวิตของแอปพลิเคชัน Flutter", "6"),
        ("2.3 กลไกการจัดการสถานะด้วย Riverpod", "7"),
        ("2.4 การจัดการเส้นทางด้วย go_router", "8"),
        ("2.5 สถาปัตยกรรม MVC ของ Laravel Framework", "9"),
        ("2.6 การทำงานของ RESTful API", "11"),
        ("2.7 ลำดับขั้นตอนการยืนยันตัวตนด้วย OAuth 2.0 Authorization Code Flow", "12"),
        ("2.8 โครงสร้างการทำงานของ LINE Messaging API", "13"),
        ("2.9 การเชื่อมต่อกับ Meta Graph API", "14"),
        ("2.10 วงจรการทำงานของ Webhook", "15"),
        ("2.11 การทำงานของคิวงานใน Laravel", "16"),
        ("2.12 การสื่อสารแบบสองทางผ่าน WebSocket ด้วย Laravel Reverb", "17"),
        ("2.13 โครงสร้างการส่งการแจ้งเตือนผ่าน Firebase Cloud Messaging", "18"),
        ("2.14 สถาปัตยกรรมฐานข้อมูลเชิงสัมพันธ์ของ MySQL", "19"),
        ("2.15 โครงสร้างการจัดเก็บเอกสารใน MongoDB", "20"),
        ("2.16 สถาปัตยกรรมระบบแบบแยกฐานข้อมูลตามผู้เช่า", "21"),
        ("2.17 โมเดลการควบคุมสิทธิ์ตามบทบาท", "22"),
        ("2.18 สถาปัตยกรรมระบบแชทบอทบริการลูกค้า", "23"),
        ("2.19 กลไกความสนใจของ Transformer", "25"),
        ("3.1 Use Case Diagram ของระบบ Chatshop", "29"),
        ("3.2 สถาปัตยกรรมของระบบ Chatshop", "30"),
        ("3.3 ความสัมพันธ์ของข้อมูลใน MySQL และ MongoDB", "32"),
        ("3.4 ขั้นตอนการรับข้อความจาก Webhook", "35"),
        ("3.5 ขั้นตอนการสร้างคำตอบของแชทบอท", "37"),
        ("3.6 ขั้นตอนการรับเหตุการณ์และการแจ้งเตือนบนแอปพลิเคชัน", "38"),
        ("3.7 โครงสร้างโค้ดของแอปพลิเคชัน", "43"),
        ("3.8 หน้ากล่องข้อความ", "44"),
        ("3.9 หน้าค้นหาบทสนทนา", "45"),
        ("3.10 หน้าต่างสนทนา", "46"),
        ("3.11 เมนูเมื่อกดค้างที่ข้อความ", "47"),
        ("3.12 การเลือกโหมดแชทบอท", "48"),
        ("3.13 แผ่นวิเคราะห์บทสนทนา", "49"),
        ("3.14 แผ่นข้อมูลลูกค้า", "50"),
        ("3.15 แบบฟอร์มเปิดคำสั่งซื้อ", "51"),
        ("3.16 หน้าการแจ้งเตือน", "52"),
        ("3.17 Push Notification บนสมาร์ตโฟน", "52"),
        ("3.18 หน้าข้อมูลลูกค้า", "53"),
        ("3.19 หน้าสินค้า", "53"),
        ("3.20 หน้ารายการคำสั่งซื้อ", "54"),
        ("3.21 หน้าจัดการทีม", "54"),
        ("3.22 หน้ารายงาน", "55"),
        ("3.23 หน้าเชื่อมต่อช่องทาง", "55"),
        ("3.24 หน้าข้อมูลผู้ใช้", "56"),
        ("3.25 เมนูการตั้งค่า", "56"),
        ("3.26 หน้าตั้งค่าแชทบอท", "56"),
        ("3.27 Channel ID และ Channel Secret ใน LINE Developers Console", "57"),
        ("3.28 การเปิด Use webhook และปิดการตอบกลับอัตโนมัติ", "58"),
        ("3.29 การกรอก Channel ID และ Channel Secret ในแอปพลิเคชัน", "58"),
        ("3.30 การตั้งค่า Facebook Login ของแอปเข้าสู่ระบบ", "59"),
        ("3.31 การเพิ่มแพลตฟอร์ม Android ในแอปเข้าสู่ระบบ", "59"),
        ("3.32 การตั้งค่า Valid OAuth Redirect URIs ของแอปเชื่อมต่อช่องทาง", "60"),
        ("3.33 การตั้งค่า Webhook ของ Meta", "61"),
        ("3.34 การเลือกเพจที่ต้องการเชื่อมต่อ", "62"),
        ("3.35 โปรเจกต์ Firebase และแอปพลิเคชัน Android", "63"),
        ("3.36 โค้ดสำหรับติด WebChat บนเว็บไซต์", "64")
    ]
    for f_name, pno in figures_list:
        add_toc_line(doc, f_name, pno)

    out_docx = os.path.join(OUT_DIR, "Chatshop_สารบัญ.docx")
    out_pdf = os.path.join(OUT_DIR, "Chatshop_สารบัญ.pdf")
    doc.save(out_docx)
    export_pdf(out_docx, out_pdf)
    print(f"Saved: {out_docx} and {out_pdf}")

def export_pdf(docx_path, pdf_path):
    word = win32com.client.DispatchEx("Word.Application")  # own instance, never the user's open Word
    word.Visible = False
    word.DisplayAlerts = 0  # no hidden modal dialogs
    try:
        d = word.Documents.Open(os.path.abspath(docx_path))
        d.Fields.Update()
        d.Repaginate()
        d.Save()
        d.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)
        d.Close(False)
    finally:
        word.Quit()

def main():
    build_cover()
    build_acknowledgements()
    build_abstract()
    build_toc()
    print("\nAll separate front matter files created successfully!")

if __name__ == "__main__":
    main()
