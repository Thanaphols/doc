"""Build Complete Thesis Book v6 (output/Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v6.docx & .pdf)
Modeled directly after the approved reference thesis:
    input/reference/ระบบรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์.docx

Matches all 15 sections of the approved book:
  - Sec 0: ปกนอก (Co-op format, 2x2 author table)
  - Sec 1: ปกใน (Co-op workplace format: ปฏิบัติงาน ณ บริษัทไฮเฟนพลัสจำกัด)
  - Sec 2: บทคัดย่อ (หน้า ง, thaiLetters start=4)
  - Sec 3: Abstract (หน้า จ)
  - Sec 4: กิตติกรรมประกาศ (หน้า ฉ, Co-op format)
  - Sec 5: สารบัญ (DiffFirstPage=True, Primary header: สารบัญ (ต่อ), toc 1 / toc 2 styles)
  - Sec 6: สารบัญตาราง (DiffFirstPage=True, Primary header: สารบัญตาราง (ต่อ), table of figures style)
  - Sec 7: สารบัญรูป (DiffFirstPage=True, Primary header: สารบัญรูป (ต่อ), table of figures style)
  - Sec 8..12: บทที่ 1..5 (Arabic page numbering starting at 1, DiffFirstPage=True)
  - Sec 13: บรรณานุกรม (Heading 1 numId=0, automatic הבรรณานุกรม (ต่อ) on page breaks)
  - Sec 14: ภาคผนวก divider + ภาคผนวก ก divider + ภาคผนวก ก คู่มือการใช้งานแอปพลิเคชัน Chatshop (with real screenshots) + ประวัติผู้เขียน (with top-right photo frame)
"""
import copy
import io
import os
import re
import shutil
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
import pymupdf
import win32com.client

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, ".agent", "skills", "rmuti-thesis", "scripts"))
from thesis_build import _open, _save, export_pdf, mark_thai, normalize_figures  # noqa: E402
from build_ch1_ch5 import src  # noqa: E402

FRIEND_DOCX = os.path.join(ROOT, "input", "reference", "ระบบรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์.docx")
TMP_DIR = os.path.join(ROOT, "process", "temp")
OUT_DIR = os.path.join(ROOT, "output")
os.makedirs(TMP_DIR, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

BODY_CH1_5 = os.path.join(TMP_DIR, "Chatshop_บทที่1_5.docx")
APPENDIX_DIV_DOCX = os.path.join(TMP_DIR, "v6_appendix_div.docx")
APPENDIX_BODY_DOCX = os.path.join(TMP_DIR, "v6_appendix_body.docx")
BIB_DOCX = os.path.join(TMP_DIR, "v6_bib.docx")
V6_DOCX = os.path.join(OUT_DIR, "Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v6.docx")
V6_PDF = os.path.join(OUT_DIR, "Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v6.pdf")

THAI_LETTERS = "กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ"

TH_TITLE = "แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์"
EN_TITLE = "Centralized Chat and Order Management Application for Online Business"
AUTHOR_FIRST_TH = "ธนพล"
AUTHOR_LAST_TH = "วุฒิธาดา"
AUTHOR_FULL_TH = "ธนพล วุฒิธาดา"
AUTHOR_FULL_EN = "Thanaphol Wuttitada"


def word_app():
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    return w


def set_para_text(p, text):
    """Replace paragraph text while preserving first text run's rPr and <w:cs/>."""
    runs = [r for r in p.runs if r.text]
    if not runs:
        r = p.add_run(text)
        mark_thai(r)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""
    mark_thai(runs[0])


def disable_auto_numbering(p):
    """Set <w:numPr><w:ilvl w:val='0'/><w:numId w:val='0'/></w:numPr> on Heading 1/2
    when used in Bibliography or Appendix A so Word does not prepend 'บทที่ N'."""
    pPr = p._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:numPr")):
        pPr.remove(old)
    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    numId = OxmlElement("w:numId")
    numId.set(qn("w:val"), "0")
    numPr.append(ilvl)
    numPr.append(numId)
    pPr.append(numPr)


def extract_ch3_screenshots():
    """Extract real mobile app screenshots (blob + width_cm) from Chapter 3 v2."""
    ch3_path = src("Chatshop_บทที่3_วิธีดำเนินงาน_v2.docx")
    ch3_doc = docx.Document(ch3_path)
    screenshots = {}

    def grab(tbl_idx, pic_idx=0):
        t = ch3_doc.tables[tbl_idx]
        blips = t._tbl.xpath(".//a:blip/@r:embed")
        exts = t._tbl.xpath(".//wp:extent")
        rId = blips[pic_idx]
        blob = ch3_doc.part.related_parts[rId].blob
        w_cm = int(exts[pic_idx].get("cx")) / 360000.0
        return blob, w_cm

    screenshots["ก.1"] = grab(12, 0)   # หน้าเข้าสู่ระบบ
    screenshots["ก.2"] = grab(12, 1)   # หน้าสมัครใช้งาน
    screenshots["ก.3"] = grab(13, 0)   # หน้ากล่องข้อความ
    screenshots["ก.4"] = grab(14, 0)   # หน้าค้นหาบทสนทนา
    screenshots["ก.5"] = grab(15, 0)   # หน้าต่างสนทนา
    screenshots["ก.6"] = grab(16, 0)   # เมนูเมื่อกดค้างที่ข้อความ
    screenshots["ก.7"] = grab(17, 0)   # การเลือกโหมดแชทบอท
    screenshots["ก.8"] = grab(18, 0)   # แผ่นวิเคราะห์บทสนทนา
    screenshots["ก.9"] = grab(19, 0)   # แผ่นข้อมูลลูกค้า
    screenshots["ก.10"] = grab(20, 0)  # แบบฟอร์มเปิดคำสั่งซื้อ
    screenshots["ก.11"] = grab(21, 0)  # หน้าการแจ้งเตือน
    screenshots["ก.12"] = grab(22, 0)  # Push Notification บนสมาร์ตโฟน
    screenshots["ก.13"] = grab(23, 0)  # หน้าข้อมูลลูกค้า
    screenshots["ก.14"] = grab(24, 0)  # หน้าสินค้า
    screenshots["ก.15"] = grab(25, 0)  # หน้ารายการคำสั่งซื้อ
    screenshots["ก.16"] = grab(26, 0)  # หน้าจัดการทีม
    screenshots["ก.17"] = grab(27, 0)  # หน้ารายงาน
    screenshots["ก.18"] = grab(28, 0)  # หน้าเชื่อมต่อช่องทาง
    screenshots["ก.19"] = grab(34, 0)  # การกรอก Channel ID และ Channel Secret ในแอปพลิเคชัน
    screenshots["ก.20"] = grab(29, 0)  # หน้าข้อมูลผู้ใช้
    screenshots["ก.21"] = grab(30, 0)  # เมนูการตั้งค่า
    screenshots["ก.22"] = grab(31, 0)  # หน้าตั้งค่าแชทบอท
    return screenshots


def build_appendix_div_doc():
    """Build Section 14 divider page ('ภาคผนวก') into APPENDIX_DIV_DOCX."""
    print("Building Appendix Divider...")
    tmpl_path = os.path.join(ROOT, "input", "reference", "rmuti_thesis_template.docx")
    doc = docx.Document(tmpl_path)
    for el in list(doc.element.body):
        if el.tag != qn("w:sectPr"):
            doc.element.body.remove(el)

    s = doc.sections[0]
    s.top_margin = Cm(3.81)
    s.bottom_margin = Cm(2.54)
    s.left_margin = Cm(3.81)
    s.right_margin = Cm(2.54)

    for _ in range(10):
        doc.add_paragraph()
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_div.add_run("ภาคผนวก")
    r_div.bold = True
    mark_thai(r_div)

    _save(doc, APPENDIX_DIV_DOCX)
    print(f"Saved Appendix Divider: {APPENDIX_DIV_DOCX}")


def build_appendix_body_doc():
    """Build Section 15 (Appendix A manual with real screenshots + Author Bio)
    into APPENDIX_BODY_DOCX."""
    print("Building Appendix A (with real screenshots) + Author Bio...")
    friend = docx.Document(FRIEND_DOCX)
    screenshots = extract_ch3_screenshots()

    # Start from template so styles match
    tmpl_path = os.path.join(ROOT, "input", "reference", "rmuti_thesis_template.docx")
    doc = docx.Document(tmpl_path)
    for el in list(doc.element.body):
        if el.tag != qn("w:sectPr"):
            doc.element.body.remove(el)

    s = doc.sections[0]
    s.top_margin = Cm(3.81)
    s.bottom_margin = Cm(2.54)
    s.left_margin = Cm(3.81)
    s.right_margin = Cm(2.54)

    # Main Heading 1 of Appendix A
    h1 = doc.add_paragraph(style="Heading 1")
    disable_auto_numbering(h1)
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Pt(36)
    h1.paragraph_format.space_after = Pt(18)
    h1.paragraph_format.keep_with_next = True
    r_h1 = h1.add_run("ภาคผนวก ก\nคู่มือการใช้งานแอปพลิเคชัน Chatshop")
    mark_thai(r_h1)

    def add_apx_h2(title, space_before_pt=14):
        p = doc.add_paragraph(style="Heading 2")
        disable_auto_numbering(p)
        p.paragraph_format.space_before = Pt(space_before_pt)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title)
        mark_thai(r)
        return p

    def add_apx_fig(fig_key, title):
        blob, w_cm = screenshots[fig_key]
        seq_no = int(fig_key.split(".")[1])
        p_pre = doc.add_paragraph()
        p_pre.paragraph_format.space_before = Pt(0)
        p_pre.paragraph_format.space_after = Pt(0)

        box = doc.add_table(rows=1, cols=1)
        trPr = box.rows[0]._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
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
        ph.paragraph_format.keep_with_next = True
        ph.paragraph_format.space_before = Pt(0)
        ph.paragraph_format.space_after = Pt(0)
        run = ph.add_run()
        run.add_picture(io.BytesIO(blob), width=Cm(w_cm))

        p_mid = cell.add_paragraph()
        p_mid.paragraph_format.keep_with_next = True
        p_mid.paragraph_format.space_before = Pt(0)
        p_mid.paragraph_format.space_after = Pt(0)

        p_cap = cell.add_paragraph(style="Caption")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(0)
        p_cap.paragraph_format.space_after = Pt(0)
        p_cap.paragraph_format.keep_with_next = True
        p_cap.add_run(f"รูปที่ {fig_key}  {title}")

        p_post = doc.add_paragraph()
        p_post.paragraph_format.space_before = Pt(0)
        p_post.paragraph_format.space_after = Pt(0)

    # ก.1
    add_apx_h2("ก.1 ข้อกำหนดและผู้ใช้งานแอปพลิเคชัน", space_before_pt=0)
    doc.add_paragraph(
        "\tแอปพลิเคชัน Chatshop ทำงานบนสมาร์ตโฟนระบบปฏิบัติการ Android ที่เชื่อมต่ออินเทอร์เน็ต "
        "โดยเชื่อมต่อกับระบบบริการหลังบ้านผ่าน REST API และรับเหตุการณ์แบบเรียลไทม์ผ่าน WebSocket "
        "ผู้ใช้งานแบ่งตามบทบาทที่ร้านค้ากำหนด ได้แก่ Owner เข้าถึงและตั้งค่าได้ทุกส่วน Manager ตอบแชท มอบหมายแชท "
        "จัดการคำสั่งซื้อและดูรายงาน Admin ตอบแชทและจัดการข้อมูลร้าน Agent ตอบแชทและเปิดคำสั่งซื้อ และ Viewer ดูข้อมูลและรายงานได้อย่างเดียว"
    )
    doc.add_paragraph(
        "\tโครงสร้างหน้าจอหลักของแอปพลิเคชันแบ่งแถบนำทางด้านล่างออกเป็น 3 แท็บ ได้แก่ แท็บกล่องข้อความสำหรับรวมและตอบแชทจากทุกช่องทาง "
        "แท็บการแจ้งเตือนสำหรับติดตามเหตุการณ์สำคัญ และแท็บข้อมูลผู้ใช้สำหรับเข้าถึงเมนูจัดการลูกค้า สินค้า คำสั่งซื้อ ทีมงาน รายงาน ช่องทาง และการตั้งค่าระบบ"
    )

    # ก.2
    add_apx_h2("ก.2 การเข้าสู่ระบบและการสมัครใช้งาน")
    doc.add_paragraph("\tก.2.1 การเข้าสู่ระบบ มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) เปิดแอปพลิเคชัน Chatshop บนสมาร์ตโฟน กรอกอีเมลและรหัสผ่านในหน้าเข้าสู่ระบบ แล้วแตะปุ่มเข้าสู่ระบบ")
    doc.add_paragraph("\t\t2) หากต้องการเข้าสู่ระบบด้วยบัญชีสังคมออนไลน์ ให้แตะปุ่มเข้าสู่ระบบด้วย Facebook หรือ Google แล้วอนุญาตสิทธิ์บนหน้าจอของผู้ให้บริการ")
    doc.add_paragraph("\t\t3) เมื่อยืนยันตัวตนสำเร็จ แอปพลิเคชันจะบันทึกโทเคนความปลอดภัยไว้บนอุปกรณ์ ลงทะเบียนรับ Push Notification และนำเข้าสู่หน้ากล่องข้อความ หน้าเข้าสู่ระบบแสดงดังรูปที่ ก.1")
    add_apx_fig("ก.1", "หน้าเข้าสู่ระบบ")

    doc.add_paragraph("\tก.2.2 การสมัครใช้งาน มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) แตะลิงก์สมัครใช้งานจากหน้าเข้าสู่ระบบเพื่อเปิดแบบฟอร์มสร้างบัญชีใหม่")
    doc.add_paragraph("\t\t2) กรอกชื่อผู้ใช้ ชื่อร้านค้า อีเมล รหัสผ่าน และยืนยันรหัสผ่านให้ครบถ้วน")
    doc.add_paragraph("\t\t3) แตะปุ่มสมัครใช้งาน ระบบหลังบ้านจะสร้างข้อมูลร้านค้าใหม่พร้อมกำหนดสิทธิ์เจ้าของร้าน (Owner) และนำเข้าสู่หน้ากล่องข้อความทันที หน้าสมัครใช้งานแสดงดังรูปที่ ก.2")
    add_apx_fig("ก.2", "หน้าสมัครใช้งาน")

    # ก.3
    add_apx_h2("ก.3 กล่องข้อความและการค้นหาบทสนทนา")
    doc.add_paragraph("\tก.3.1 การใช้งานกล่องข้อความ มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) แตะแท็บกล่องข้อความ แอปพลิเคชันจะแสดงรายการบทสนทนาจาก LINE Official Account, Facebook Messenger, Instagram และ WebChat รวมในรายการเดียว โดยเรียงบทสนทนาที่ปักหมุดไว้ด้านบนสุด ตามด้วยเวลาข้อความล่าสุด")
    doc.add_paragraph("\t\t2) เลือกกรองรายการบทสนทนาตามช่องทาง หรือเลือกแสดงเฉพาะบทสนทนาที่ยังไม่อ่านจากแถบตัวกรองด้านบน")
    doc.add_paragraph("\t\t3) ปัดรายการบทสนทนาเพื่อจัดเก็บเข้าคลังหรือลบ หรือกดค้างที่รายการบทสนทนาเพื่อปักหมุดและมอบหมายบทสนทนาให้สมาชิกในทีม หน้ากล่องข้อความแสดงดังรูปที่ ก.3")
    add_apx_fig("ก.3", "หน้ากล่องข้อความ")

    doc.add_paragraph("\tก.3.2 การค้นหาบทสนทนา มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) แตะไอคอนค้นหาที่แถบด้านบนของหน้ากล่องข้อความ")
    doc.add_paragraph("\t\t2) พิมพ์คำค้น เช่น ชื่อลูกค้า ข้อความในแชท หรือชื่อแท็ก แอปพลิเคชันจะส่งคำค้นไปยังระบบหลังบ้านและแสดงรายการบทสนทนาที่ตรงกับเงื่อนไขให้แตะเปิดดูได้ทันที หน้าค้นหาบทสนทนาแสดงดังรูปที่ ก.4")
    add_apx_fig("ก.4", "หน้าค้นหาบทสนทนา")

    # ก.4
    add_apx_h2("ก.4 หน้าต่างสนทนาและการจัดการข้อความ")
    doc.add_paragraph("\tก.4.1 การตอบข้อความและแนบไฟล์สื่อ มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) แตะเลือกบทสนทนาจากกล่องข้อความเพื่อเปิดหน้าต่างสนทนา ซึ่งแสดงประวัติข้อความทั้งหมดและโหลดข้อความเก่าเพิ่มเมื่อเลื่อนขึ้น")
    doc.add_paragraph("\t\t2) พิมพ์ข้อความตอบกลับในช่องพิมพ์ด้านล่าง หรือแตะเลือกข้อความสำเร็จรูปที่ตั้งไว้ แล้วแตะปุ่มส่งข้อความ")
    doc.add_paragraph("\t\t3) หากต้องการส่งไฟล์สื่อ ให้แตะปุ่มแนบไฟล์เพื่อเลือกรูปภาพ วิดีโอ หรือเอกสารจากอุปกรณ์ หรือถ่ายภาพจากกล้อง แอปพลิเคชันจะแสดงข้อความในหน้าต่างสนทนาทันทีและส่งไปยังช่องทางของลูกค้า หน้าต่างสนทนาแสดงดังรูปที่ ก.5")
    add_apx_fig("ก.5", "หน้าต่างสนทนา")

    doc.add_paragraph("\tก.4.2 การจัดการข้อความในบทสนทนา มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) กดค้างที่ข้อความใดข้อความหนึ่งในหน้าต่างสนทนาเพื่อเปิดเมนูจัดการข้อความ")
    doc.add_paragraph("\t\t2) เลือกตอบกลับโดยอ้างอิงข้อความเดิม (Reply) ปักหมุดข้อความสำคัญ คัดลอกข้อความ หรือส่งอีโมจิแสดงความรู้สึกต่อข้อความ เมนูเมื่อกดค้างที่ข้อความแสดงดังรูปที่ ก.6")
    add_apx_fig("ก.6", "เมนูเมื่อกดค้างที่ข้อความ")

    # ก.5
    add_apx_h2("ก.5 การควบคุมแชทบอทและการวิเคราะห์บทสนทนา")
    doc.add_paragraph("\tก.5.1 การเลือกโหมดแชทบอท มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) แตะปุ่มสถานะแชทบอทบริเวณส่วนหัวของหน้าต่างสนทนา")
    doc.add_paragraph("\t\t2) เลือกโหมดการทำงานของแชทบอทสำหรับบทสนทนานั้น ได้แก่ Manual (เจ้าหน้าที่ตอบเองทั้งหมด), Suggest (แชทบอทแนะนำคำตอบให้เจ้าหน้าที่ตรวจและกดส่ง) หรือ Autopilot (แชทบอทตอบลูกค้าอัตโนมัติจากฐานความรู้ของร้าน) การเลือกโหมดแชทบอทแสดงดังรูปที่ ก.7")
    add_apx_fig("ก.7", "การเลือกโหมดแชทบอท")

    doc.add_paragraph("\tก.5.2 การดูแผ่นวิเคราะห์บทสนทนา มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) แตะปุ่มวิเคราะห์บทสนทนาบนส่วนหัวของหน้าต่างสนทนา")
    doc.add_paragraph("\t\t2) แอปพลิเคชันจะขอให้ระบบหลังบ้านส่งประวัติบทสนทนาให้แบบจำลองภาษาวิเคราะห์ และแสดงผลสรุปความรู้สึกของลูกค้า เจตนา หัวข้อที่สนใจ และสัญญาณเตือน เช่น ลูกค้าเร่งด่วนหรือพร้อมสั่งซื้อ แผ่นวิเคราะห์บทสนทนาแสดงดังรูปที่ ก.8")
    add_apx_fig("ก.8", "แผ่นวิเคราะห์บทสนทนา")

    # ก.6
    add_apx_h2("ก.6 ข้อมูลลูกค้าและการเปิดคำสั่งซื้อจากหน้าต่างสนทนา")
    doc.add_paragraph("\tก.6.1 การดูและแก้ไขแผ่นข้อมูลลูกค้า มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) แตะชื่อลูกค้าหรือปุ่มข้อมูลลูกค้าที่ส่วนหัวของหน้าต่างสนทนาเพื่อเปิดแผ่นข้อมูลลูกค้า")
    doc.add_paragraph("\t\t2) ตรวจสอบหรือแก้ไขเบอร์โทรศัพท์ ที่อยู่จัดส่ง บันทึกภายใน และแท็กของลูกค้า รวมถึงดูประวัติคำสั่งซื้อของลูกค้ารายนั้น แผ่นข้อมูลลูกค้าแสดงดังรูปที่ ก.9")
    add_apx_fig("ก.9", "แผ่นข้อมูลลูกค้า")

    doc.add_paragraph("\tก.6.2 การเปิดคำสั่งซื้อจากหน้าต่างสนทนา มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) แตะปุ่มสร้างคำสั่งซื้อจากหน้าต่างสนทนาหรือจากแผ่นข้อมูลลูกค้า")
    doc.add_paragraph("\t\t2) เลือกสินค้าจากแคตตาล็อกของร้านหรือกรอกชื่อสินค้าและราคาเอง ระบุจำนวนสินค้า และตรวจสอบที่อยู่จัดส่งซึ่งระบบดึงมาจากข้อมูลลูกค้าให้อัตโนมัติ")
    doc.add_paragraph("\t\t3) แตะปุ่มบันทึกคำสั่งซื้อ ระบบจะคำนวณยอดรวม สร้างเลขที่คำสั่งซื้อ และผูกคำสั่งซื้อเข้ากับบทสนทนาทันที แบบฟอร์มเปิดคำสั่งซื้อแสดงดังรูปที่ ก.10")
    add_apx_fig("ก.10", "แบบฟอร์มเปิดคำสั่งซื้อ")

    # ก.7
    add_apx_h2("ก.7 การแจ้งเตือนและ Push Notification")
    doc.add_paragraph("\tก.7.1 การตรวจสอบรายการแจ้งเตือน มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) แตะแท็บการแจ้งเตือนที่แถบนำทางด้านล่างเพื่อดูรายการแจ้งเตือนทั้ง 5 ประเภท ได้แก่ เร่งด่วน ข้อความใหม่ ชำระเงินแล้ว การกล่าวถึง และระบบ")
    doc.add_paragraph("\t\t2) แตะรายการแจ้งเตือนเพื่อเปิดไปยังบทสนทนาที่เกี่ยวข้องทันที หรือแตะทำเครื่องหมายว่าอ่านทั้งหมด หน้าการแจ้งเตือนแสดงดังรูปที่ ก.11")
    add_apx_fig("ก.11", "หน้าการแจ้งเตือน")

    doc.add_paragraph("\tก.7.2 การรับ Push Notification บนสมาร์ตโฟน")
    doc.add_paragraph("\t\tเมื่อมีข้อความใหม่จากลูกค้าหรือมีแชทค้างตอบเกินเวลาที่ร้านกำหนด ระบบหลังบ้านจะส่งแจ้งเตือนผ่าน Firebase Cloud Messaging มายังสมาร์ตโฟนของเจ้าหน้าที่แม้ไม่ได้เปิดแอปพลิเคชันอยู่ เมื่อแตะที่แจ้งเตือน แอปพลิเคชันจะเปิดหน้าต่างสนทนานั้นทันที ดังรูปที่ ก.12")
    add_apx_fig("ก.12", "Push Notification บนสมาร์ตโฟน")

    # ก.8
    add_apx_h2("ก.8 การจัดการลูกค้า สินค้า คำสั่งซื้อ และทีมงาน")
    doc.add_paragraph("\tก.8.1 การจัดการข้อมูลลูกค้า")
    doc.add_paragraph("\t\tเข้าที่แท็บข้อมูลผู้ใช้แล้วเลือกเมนูลูกค้า เพื่อดูรายชื่อลูกค้าจากทุกช่องทาง ค้นหาลูกค้า และแก้ไขเบอร์โทรศัพท์ ที่อยู่ และบันทึกภายใน หน้าข้อมูลลูกค้าแสดงดังรูปที่ ก.13")
    add_apx_fig("ก.13", "หน้าข้อมูลลูกค้า")

    doc.add_paragraph("\tก.8.2 การจัดการสินค้า")
    doc.add_paragraph("\t\tเลือกเมนูสินค้าจากแท็บข้อมูลผู้ใช้ เพื่อดูแคตตาล็อกสินค้า ค้นหา เพิ่มสินค้าใหม่ แก้ไขราคาและรายละเอียด หรือลบสินค้า หน้าสินค้าแสดงดังรูปที่ ก.14")
    add_apx_fig("ก.14", "หน้าสินค้า")

    doc.add_paragraph("\tก.8.3 การจัดการคำสั่งซื้อ")
    doc.add_paragraph("\t\tเลือกเมนูคำสั่งซื้อจากแท็บข้อมูลผู้ใช้ เพื่อดูรายการคำสั่งซื้อแยกแท็บตามสถานะ ได้แก่ รอดำเนินการ ชำระเงินแล้ว จัดส่งแล้ว และยกเลิก พร้อมค้นหาและปรับเปลี่ยนสถานะคำสั่งซื้อ หน้ารายการคำสั่งซื้อแสดงดังรูปที่ ก.15")
    add_apx_fig("ก.15", "หน้ารายการคำสั่งซื้อ")

    doc.add_paragraph("\tก.8.4 การจัดการทีมงาน")
    doc.add_paragraph("\t\tเจ้าของร้านเลือกเมนูทีมจากแท็บข้อมูลผู้ใช้ เพื่อเชิญสมาชิกใหม่ทางอีเมลพร้อมกำหนดบทบาท เปลี่ยนบทบาทของสมาชิก ยกเลิกคำเชิญ หรือโอนสิทธิ์ความเป็นเจ้าของร้าน หน้าจัดการทีมแสดงดังรูปที่ ก.16")
    add_apx_fig("ก.16", "หน้าจัดการทีม")

    # ก.9
    add_apx_h2("ก.9 รายงานสรุปผลและการเชื่อมต่อช่องทางแชท")
    doc.add_paragraph("\tก.9.1 การดูรายงานสรุปผล")
    doc.add_paragraph("\t\tเลือกเมนูรายงานจากแท็บข้อมูลผู้ใช้ แล้วเลือกช่วงเวลา ได้แก่ วันนี้ 7 วัน 30 วัน หรือรายไตรมาส เพื่อดูตัวชี้วัดการให้บริการ เช่น จำนวนแชททั้งหมด แชทค้างตอบ เวลาตอบครั้งแรก เวลาปิดแชท และสัดส่วนข้อความแยกตามช่องทาง หน้ารายงานแสดงดังรูปที่ ก.17")
    add_apx_fig("ก.17", "หน้ารายงาน")

    doc.add_paragraph("\tก.9.2 การเชื่อมต่อช่องทางแชท มีขั้นตอนดังนี้")
    doc.add_paragraph("\t\t1) เลือกเมนูช่องทางจากแท็บข้อมูลผู้ใช้ เพื่อดูสถานะช่องทางที่เชื่อมต่อแล้วและเพิ่มช่องทางใหม่ ได้แก่ Facebook Page, Instagram, LINE Official Account และ WebChat หน้าเชื่อมต่อช่องทางแสดงดังรูปที่ ก.18")
    add_apx_fig("ก.18", "หน้าเชื่อมต่อช่องทาง")
    doc.add_paragraph("\t\t2) สำหรับการเชื่อมต่อ LINE Official Account ให้คัดลอก Channel ID และ Channel Secret จาก LINE Developers Console มากรอกในแอปพลิเคชันแล้วแตะเชื่อมต่อ ดังรูปที่ ก.19")
    add_apx_fig("ก.19", "การกรอก Channel ID และ Channel Secret ในแอปพลิเคชัน")

    # ก.10
    add_apx_h2("ก.10 ข้อมูลผู้ใช้และการตั้งค่าระบบ")
    doc.add_paragraph("\tก.10.1 หน้าข้อมูลผู้ใช้และเมนูการตั้งค่า")
    doc.add_paragraph("\t\tเมื่อแตะแท็บข้อมูลผู้ใช้ แอปพลิเคชันจะแสดงสถิติการทำงานของผู้ใช้และเมนูลัดไปยังส่วนจัดการต่าง ๆ ดังรูปที่ ก.20 และเมื่อเลื่อนลงมาด้านล่างจะพบหมวดการตั้งค่าร้านค้า หมวดการแจ้งเตือน และหมวดบัญชีผู้ใช้ ดังรูปที่ ก.21")
    add_apx_fig("ก.20", "หน้าข้อมูลผู้ใช้")
    add_apx_fig("ก.21", "เมนูการตั้งค่า")

    doc.add_paragraph("\tก.10.2 การตั้งค่าแชทบอท")
    doc.add_paragraph("\t\tแตะเลือกเมนูแชทบอทในหมวดการตั้งค่าร้านค้า เพื่อเลือกผู้ให้บริการแบบจำลองภาษา กรอกคีย์ API เลือกชื่อแบบจำลอง และทดสอบการเชื่อมต่อ หน้าตั้งค่าแชทบอทแสดงดังรูปที่ ก.22")
    add_apx_fig("ก.22", "หน้าตั้งค่าแชทบอท")

    # Tag all runs in Appendix A
    for p in doc.paragraphs:
        for r in p.runs:
            mark_thai(r)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        mark_thai(r)

    # 4. Clone Author Bio (paragraphs 1412..1424) directly from friend's book!
    bio_replacements = {
        1415: f"ชื่อ-นามสกุล\t\tนาย{AUTHOR_FULL_TH}",
        1416: "วัน-เดือน-ปีเกิด\t\t4 สิงหาคม 2545",
        1417: "ที่อยู่\t\t\t73 หมู่ 1 ต.ปากตม อ.เชียงคาน จ.เลย 42110",
        1418: "อีเมล\t\t\tthanaphol.wu@rmuti.ac.th",
        1419: "เบอร์โทรศัพท์\t\t0968783234",
    }
    for idx in range(1412, 1425):
        cloned = copy.deepcopy(friend.paragraphs[idx]._p)
        if cloned.pPr is not None:
            for sp in cloned.pPr.findall(qn("w:sectPr")):
                cloned.pPr.remove(sp)
        doc.element.body.append(cloned)
        if idx == 1415:
            # Keep the first run containing the photo frame drawing (mc:AlternateContent), update text runs
            p_new = docx.text.paragraph.Paragraph(cloned, doc)
            for r in p_new.runs:
                if "วชิรวิทย์" in r.text:
                    r.text = f"นาย{AUTHOR_FULL_TH}"
                    mark_thai(r)
        elif idx in bio_replacements:
            p_new = docx.text.paragraph.Paragraph(cloned, doc)
            set_para_text(p_new, bio_replacements[idx])

    normalize_figures(doc)
    _save(doc, APPENDIX_BODY_DOCX)
    print(f"Saved Appendix A + Author Bio: {APPENDIX_BODY_DOCX}")


def prepare_bib_doc():
    """Prepare Bibliography with Heading 1 (numId=0) matching friend's Sec 13."""
    bib_src = src("Chatshop_บรรณานุกรม.docx")
    doc = docx.Document(bib_src)
    h = doc.paragraphs[0]
    h.style = doc.styles["Heading 1"]
    disable_auto_numbering(h)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.paragraph_format.space_before = Pt(36)
    h.paragraph_format.keep_with_next = True
    for r in h.runs:
        r.bold = None
        mark_thai(r)
    _save(doc, BIB_DOCX)
    return BIB_DOCX


def update_front_matter_metadata(doc):
    """Update Cover, Inner Cover, Thai Abstract, English Abstract, and Co-op Acknowledgements
    in a copy of the friend's passed book (Sections 0..4)."""
    ps = doc.paragraphs

    # Title split across 2 balanced lines for Cover & Inner Cover so it fits exactly on 1 page each
    EN_TITLE_COVER = "Centralized Chat and Order Management\nApplication for Online Business"

    # Sec 0: Outer Cover
    set_para_text(ps[3], TH_TITLE)
    set_para_text(ps[4], EN_TITLE_COVER)
    set_para_text(doc.tables[0].rows[0].cells[0].paragraphs[0], AUTHOR_FIRST_TH)
    set_para_text(doc.tables[0].rows[0].cells[1].paragraphs[0], AUTHOR_LAST_TH)

    # Change Sec 0 w:sectPr type from continuous to nextPage so Sec 0 never spills into Sec 1
    sectPr0 = ps[23]._p.xpath(".//w:sectPr")[0]
    for t in sectPr0.findall(qn("w:type")):
        sectPr0.remove(t)
    type_el = OxmlElement("w:type")
    type_el.set(qn("w:val"), "nextPage")
    sectPr0.append(type_el)

    # Sec 1: Inner Cover
    set_para_text(ps[27], TH_TITLE)
    set_para_text(ps[28], EN_TITLE_COVER)
    set_para_text(doc.tables[1].rows[0].cells[0].paragraphs[0], AUTHOR_FIRST_TH)
    set_para_text(doc.tables[1].rows[0].cells[1].paragraphs[0], AUTHOR_LAST_TH)

    # Sec 2: Thai Abstract
    set_para_text(ps[48], f"หัวข้อปริญญานิพนธ์\t{TH_TITLE}")
    set_para_text(ps[49], f"จัดทำโดย\t\t\t{AUTHOR_FULL_TH}")
    set_para_text(
        ps[57],
        "\tโครงงานนี้มีวัตถุประสงค์เพื่อพัฒนาแอปพลิเคชันบนสมาร์ตโฟนของ Chatshop ซึ่งเป็นแอปพลิเคชันรวมศูนย์ข้อความและ"
        "การจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ ร้านค้าที่ให้บริการลูกค้าหลายช่องทางต้องสลับเปิดหลายแอปพลิเคชันบนสมาร์ตโฟน "
        "ทำให้พลาดการแจ้งเตือน ตอบลูกค้าล่าช้า ตอบข้อความซ้ำซ้อน และข้อมูลคำสั่งซื้อกระจัดกระจาย แอปพลิเคชันนี้จึงรวบรวมแชทจาก "
        "LINE Official Account, Facebook Messenger, Instagram และ WebChat ไว้ในกล่องข้อความเดียว",
    )
    set_para_text(
        ps[58],
        "\tแอปพลิเคชันพัฒนาด้วย Flutter และภาษา Dart จัดการสถานะด้วย Riverpod เชื่อมต่อระบบบริการหลังบ้านที่พัฒนาด้วย Laravel "
        "ร่วมกับฐานข้อมูล MySQL และ MongoDB รับส่งเหตุการณ์แบบเรียลไทม์ผ่าน Laravel Reverb และแจ้งเตือน Push Notification "
        "ผ่าน Firebase Cloud Messaging รองรับการค้นหาและคัดกรองบทสนทนา การล็อกและมอบหมายบทสนทนา การควบคุมโหมดแชทบอทและ"
        "ดูผลวิเคราะห์บทสนทนา การจัดการข้อมูลสินค้า ลูกค้า และการเปิดคำสั่งซื้อจากหน้าต่างสนทนา รวมถึงการดูรายงานสรุปผลการให้บริการ",
    )
    set_para_text(
        ps[59],
        "\tผลการทดสอบแบบ Black Box ตามกรณีทดสอบ 16 กรณี พบว่าระบบทำงานได้ถูกต้องทุกกรณี และผลการประเมินความพึงพอใจของ"
        "ผู้ใช้งานจำนวน 15 คน มีค่าเฉลี่ยรวมเท่ากับ 4.10 และส่วนเบี่ยงเบนมาตรฐานเท่ากับ 0.37 อยู่ในระดับดี",
    )
    set_para_text(ps[60], "คำสำคัญ\t\t:  แอปพลิเคชันรวมศูนย์ข้อความ, ร้านค้าออนไลน์, แชทบอท, Flutter, Laravel")

    # Sec 3: English Abstract
    set_para_text(ps[62], f"Project Title\t{EN_TITLE}")
    set_para_text(ps[63], f"Proposed by\t{AUTHOR_FULL_EN}")
    set_para_text(
        ps[71],
        "\tThis project aims to develop the mobile application of Chatshop, a centralized chat and order management "
        "application for online businesses. Online shops serving customers across multiple channels must switch between "
        "separate smartphone applications, causing missed notifications, delayed replies, duplicate responses, and "
        "scattered order records. The application therefore consolidates chats from LINE Official Account, Facebook "
        "Messenger, Instagram, and WebChat into a single mobile inbox.",
    )
    set_para_text(
        ps[72],
        "\tThe mobile application is developed with Flutter and Dart using Riverpod for state management, connecting "
        "to a Laravel backend with MySQL and MongoDB databases. It receives real-time events via Laravel Reverb and "
        "mobile Push Notifications via Firebase Cloud Messaging, supporting conversation search and filtering, "
        "conversation locking and assignment, chatbot mode control and AI conversation analysis, product and customer "
        "management, order creation directly from the chat screen, and service summary reports.",
    )
    set_para_text(
        ps[73],
        "\tBlack-box testing across 16 test cases showed that the system operated correctly in all cases. The user "
        "satisfaction survey of 15 participants yielded an overall mean score of 4.10 with a standard deviation of "
        "0.37, which corresponds to the good satisfaction level.",
    )
    set_para_text(ps[75], "Keywords: Centralized Chat Application, Online Shop, Chatbot, Flutter, Laravel")

    # Sec 4: Co-op Acknowledgements
    # Replace triple tabs with single tab to prevent line wrap and text stretching
    set_para_text(ps[83], ps[83].text.replace("\t\t\t", "\t").rstrip())
    set_para_text(ps[89], f"\t\t\t\t\t\t\t\t  {AUTHOR_FIRST_TH}      {AUTHOR_LAST_TH}")

    # Remove 4 spacing paragraphs in Sec 0 (Cover) and Sec 1 (Inner Cover) so neither overflows
    for idx in sorted([9, 18, 35, 45], reverse=True):
        p = ps[idx]
        p._p.getparent().remove(p._p)


def populate_toc_sections(doc, toc_entries, lot_entries, lof_entries, front_letters):
    """Replace the entries inside Sections 5 (สารบัญ), 6 (สารบัญตาราง), and 7 (สารบัญรูป)
    of the friend's passed book using the exact 'toc 1', 'toc 2', and 'table of figures'
    paragraph templates from the friend's book."""
    friend_ref = docx.Document(FRIEND_DOCX)
    f_ps = friend_ref.paragraphs
    f_sects = [i for i, p in enumerate(f_ps) if p._p.xpath(".//w:sectPr")]
    toc1_tmpl = copy.deepcopy(next(p._p for p in f_ps[f_sects[4] + 1 : f_sects[5]] if p.style.name == "toc 1"))
    toc2_tmpl = copy.deepcopy(next(p._p for p in f_ps[f_sects[4] + 1 : f_sects[5]] if p.style.name == "toc 2"))
    tof_tmpl = copy.deepcopy(next(p._p for p in f_ps[f_sects[5] + 1 : f_sects[6]] if p.style.name == "table of figures"))

    def make_entry_el(tmpl_el, left_text, right_page):
        el = copy.deepcopy(tmpl_el)
        if el.pPr is not None:
            for sp in el.pPr.findall(qn("w:sectPr")):
                el.pPr.remove(sp)
        for child in list(el):
            if child.tag != qn("w:pPr"):
                el.remove(child)
        p_obj = docx.text.paragraph.Paragraph(el, doc)
        r1 = p_obj.add_run(left_text)
        mark_thai(r1)
        r_tab = p_obj.add_run("\t")
        mark_thai(r_tab)
        r2 = p_obj.add_run(str(right_page))
        mark_thai(r2)
        return el

    # --- Update Sec 5 (สารบัญ) ---
    ps = doc.paragraphs
    sect_p_indices = [i for i, p in enumerate(ps) if p._p.xpath(".//w:sectPr")]
    sec4_end, sec5_end = sect_p_indices[4], sect_p_indices[5]
    sec5_paras = ps[sec4_end + 1 : sec5_end]
    anchor_sec5 = sec5_paras[1]._p  # 'หน้า' paragraph
    for p in sec5_paras[2:]:
        p._p.getparent().remove(p._p)

    full_toc = [
        ("บทคัดย่อ", front_letters.get("บทคัดย่อ", "ง"), 1),
        ("Abstract", front_letters.get("Abstract", "จ"), 1),
        ("กิตติกรรมประกาศ", front_letters.get("กิตติกรรมประกาศ", "ฉ"), 1),
        ("สารบัญ", front_letters.get("สารบัญ", "ช"), 1),
        ("สารบัญตาราง", front_letters.get("สารบัญตาราง", "ญ"), 1),
        ("สารบัญรูป", front_letters.get("สารบัญรูป", "ฎ"), 1),
    ] + toc_entries

    cur = anchor_sec5
    for text, page, lvl in full_toc:
        el = make_entry_el(toc1_tmpl if lvl == 1 else toc2_tmpl, text, page)
        cur.addnext(el)
        cur = el

    # --- Update Sec 6 (สารบัญตาราง) ---
    ps = doc.paragraphs
    sect_p_indices = [i for i, p in enumerate(ps) if p._p.xpath(".//w:sectPr")]
    sec5_end, sec6_end = sect_p_indices[5], sect_p_indices[6]
    sec6_paras = ps[sec5_end + 1 : sec6_end]
    anchor_sec6 = sec6_paras[1]._p  # 'ตาราง\tหน้า'
    for p in sec6_paras[2:]:
        p._p.getparent().remove(p._p)

    cur = anchor_sec6
    for text, page in lot_entries:
        el = make_entry_el(tof_tmpl, text, page)
        cur.addnext(el)
        cur = el

    # --- Update Sec 7 (สารบัญรูป) ---
    ps = doc.paragraphs
    sect_p_indices = [i for i, p in enumerate(ps) if p._p.xpath(".//w:sectPr")]
    sec6_end, sec7_end = sect_p_indices[6], sect_p_indices[7]
    sec7_paras = ps[sec6_end + 1 : sec7_end]
    anchor_sec7 = sec7_paras[1]._p  # 'รูป\tหน้า'
    for p in sec7_paras[2:]:
        p._p.getparent().remove(p._p)

    cur = anchor_sec7
    for text, page in lof_entries:
        el = make_entry_el(tof_tmpl, text, page)
        cur.addnext(el)
        cur = el


def ensure_clean_body_sections():
    """Ensure BODY_CH1_5 has exactly 5 sections (one per chapter, no old page-1 break inside Ch3)."""
    from thesis_build import split_tables, unsplit_tables
    bdoc = _open(BODY_CH1_5)
    if len(bdoc.sections) > 5:
        print(f"Cleaning stray section break in {BODY_CH1_5} ({len(bdoc.sections)} sections -> 5)...")
        for p in bdoc.paragraphs:
            if p.text.strip().startswith("3.2.1.1 ") and p._p.pPr is not None:
                for sp in p._p.pPr.findall(qn("w:sectPr")):
                    p._p.pPr.remove(sp)
        unsplit_tables(bdoc)
        _save(bdoc, BODY_CH1_5)
        split_tables(BODY_CH1_5)


def assemble_v6():
    # Step 1: Ensure clean 5-section body, and prepare Appendix Divider, Appendix Body + Bio, and Bibliography
    ensure_clean_body_sections()
    build_appendix_div_doc()
    build_appendix_body_doc()
    prepare_bib_doc()

    # Step 2: Copy friend's passed book to V6_DOCX, update Front Matter metadata,
    # strip friend's body (after Sec 7), and insert our Body (Chapters 1-5 + Bib + Appendix Div + Appendix Body & Bio)
    print("Preparing base v6 document from friend's approved thesis...")
    shutil.copy2(FRIEND_DOCX, V6_DOCX)
    doc = _open(V6_DOCX)
    update_front_matter_metadata(doc)

    # Remove all body elements after Sec 7's sectPr paragraph ( keep body's final sectPr )
    ps = doc.paragraphs
    sec7_end_p = [p for p in ps if p._p.xpath(".//w:sectPr")][7]._p
    body = doc.element.body
    idx_sec7 = list(body).index(sec7_end_p)
    for el in list(body)[idx_sec7 + 1 :]:
        if el.tag != qn("w:sectPr"):
            body.remove(el)
    _save(doc, V6_DOCX)

    # Step 3: Insert Chapters 1-5, Bibliography, Appendix Divider, and Appendix Body + Bio via Word COM
    print("Inserting Chapters 1-5, Bibliography, Appendix Divider, and Appendix Body + Bio via Word COM...")
    word = word_app()
    try:
        d = word.Documents.Open(os.path.abspath(V6_DOCX))
        rng = d.Content
        rng.Collapse(0)
        rng.InsertFile(os.path.abspath(BODY_CH1_5))

        rng = d.Content
        rng.Collapse(0)
        rng.InsertBreak(2)  # wdSectionBreakNextPage -> Sec 14 in Word COM (Bibliography)
        rng.Collapse(0)
        rng.InsertFile(os.path.abspath(BIB_DOCX))

        rng = d.Content
        rng.Collapse(0)
        rng.InsertBreak(2)  # wdSectionBreakNextPage -> Sec 15 in Word COM (Appendix Divider)
        rng.Collapse(0)
        rng.InsertFile(os.path.abspath(APPENDIX_DIV_DOCX))

        rng = d.Content
        rng.Collapse(0)
        rng.InsertBreak(2)  # wdSectionBreakNextPage -> Sec 16 in Word COM (Appendix Body + Bio)
        rng.Collapse(0)
        rng.InsertFile(os.path.abspath(APPENDIX_BODY_DOCX))

        # Configure headers and page numbering on Sections 9..N (1-based in Word COM: Sec 9 is Chapter 1)
        # Sec 9..13: Chapters 1..5
        # Sec 14: Bibliography
        # Sec 15: Appendix Divider (NO page number on both first and primary headers)
        # Sec 16: Appendix Body + Bio (NO page number on first page, continues Arabic page numbering on page 2)
        total_secs = d.Sections.Count
        sec_div_idx = total_secs - 1
        sec_body_idx = total_secs

        for si in range(9, total_secs + 1):
            sec = d.Sections(si)
            sec.PageSetup.TopMargin = 108     # 3.81 cm
            sec.PageSetup.BottomMargin = 72   # 2.54 cm
            sec.PageSetup.LeftMargin = 108    # 3.81 cm
            sec.PageSetup.RightMargin = 72    # 2.54 cm
            sec.PageSetup.HeaderDistance = 72 # 2.54 cm
            sec.PageSetup.DifferentFirstPageHeaderFooter = True
            h_prim = sec.Headers(1)
            h_first = sec.Headers(2)

            if si == 9:
                h_prim.LinkToPrevious = False
                h_first.LinkToPrevious = False
                h_first.Range.Text = ""
                h_prim.Range.Text = ""
                h_prim.Range.Fields.Add(h_prim.Range, 33)  # wdFieldPage
                h_prim.Range.ParagraphFormat.Alignment = 2
                h_prim.Range.Font.Name = "TH SarabunPSK"
                h_prim.Range.Font.NameBi = "TH SarabunPSK"
                h_prim.Range.Font.Size = 16
                h_prim.Range.Font.SizeBi = 16
                pn = h_prim.PageNumbers
                pn.NumberStyle = 0  # Arabic
                pn.RestartNumberingAtSection = True
                pn.StartingNumber = 1
            elif si == sec_div_idx:
                # Appendix Divider: completely empty header
                h_prim.LinkToPrevious = False
                h_first.LinkToPrevious = False
                h_first.Range.Text = ""
                h_prim.Range.Text = ""
            elif si == sec_body_idx:
                # Appendix Body + Bio: First page header empty, primary header continues page numbering
                h_first.LinkToPrevious = False
                h_first.Range.Text = ""
                h_prim.LinkToPrevious = False
                h_prim.Range.Text = ""
                h_prim.Range.Fields.Add(h_prim.Range, 33)  # wdFieldPage
                h_prim.Range.ParagraphFormat.Alignment = 2
                h_prim.Range.Font.Name = "TH SarabunPSK"
                h_prim.Range.Font.NameBi = "TH SarabunPSK"
                h_prim.Range.Font.Size = 16
                h_prim.Range.Font.SizeBi = 16
                pn = h_prim.PageNumbers
                pn.NumberStyle = 0  # Arabic
                pn.RestartNumberingAtSection = False
            else:
                h_first.LinkToPrevious = False
                h_first.Range.Text = ""
                h_prim.LinkToPrevious = True
                h_prim.PageNumbers.RestartNumberingAtSection = False

        d.Fields.Update()
        d.Repaginate()
        d.Save()
        d.Close(False)
    finally:
        word.Quit()

    # Clean up any empty Heading 1 paragraph at junction after Sec 7
    doc = _open(V6_DOCX)
    for p in doc.paragraphs:
        if not p.text.strip() and p.style.name == "Heading 1":
            p.style = doc.styles["Normal"]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
    normalize_figures(doc)
    _save(doc, V6_DOCX)

    # Step 4: Paginate Bibliography ("บรรณานุกรม (ต่อ)" on page breaks) and collect exact page numbers
    print("Paginating Bibliography and collecting exact TOC/LOT/LOF page numbers...")
    insert_bib_continuations(V6_DOCX)

    # Iteratively update TOC/LOT/LOF until front-matter Thai letters and body page numbers settle
    front_letters = {}
    for round_no in range(1, 4):
        toc_entries, lot_entries, lof_entries, new_letters = collect_entries_from_word(V6_DOCX)
        print(
            f"  Round {round_no}: {len(toc_entries)} TOC, {len(lot_entries)} LOT, "
            f"{len(lof_entries)} LOF, letters={new_letters}"
        )
        doc = _open(V6_DOCX)
        populate_toc_sections(doc, toc_entries, lot_entries, lof_entries, new_letters)
        _save(doc, V6_DOCX)
        if new_letters == front_letters and round_no >= 2:
            break
        front_letters = new_letters

    # Final field update & PDF export
    print(f"Exporting final PDF to {V6_PDF}...")
    word = word_app()
    try:
        d = word.Documents.Open(os.path.abspath(V6_DOCX))
        d.Fields.Update()
        d.Repaginate()
        d.Save()
        d.ExportAsFixedFormat(os.path.abspath(V6_PDF), 17)
        total_pages = d.ComputeStatistics(2)
        print(f"Exported {V6_PDF} ({total_pages} pages)")
        d.Close(False)
    finally:
        word.Quit()


def insert_bib_continuations(docx_path):
    """When bibliography entries cross a page boundary, insert a centered bold
    'บรรณานุกรม (ต่อ)' + blank line at the top of the new page, matching friend's Sec 13."""
    for _ in range(5):
        word = word_app()
        split_entry_text = None
        try:
            d = word.Documents.Open(os.path.abspath(docx_path))
            d.Repaginate()
            bib_sec = None
            for si in range(9, d.Sections.Count + 1):
                if "บรรณานุกรม" in d.Sections(si).Range.Paragraphs(1).Range.Text:
                    bib_sec = d.Sections(si)
                    break
            if bib_sec is not None:
                paras = bib_sec.Range.Paragraphs
                cur_page = paras(1).Range.Information(3)
                for pi in range(2, paras.Count + 1):
                    p = paras(pi)
                    txt = p.Range.Text.strip()
                    if not txt:
                        continue
                    pg = p.Range.Information(3)
                    if pg > cur_page:
                        if txt != "บรรณานุกรม (ต่อ)":
                            split_entry_text = txt[:30]
                            break
                        cur_page = pg
            d.Close(False)
        finally:
            word.Quit()

        if not split_entry_text:
            break

        doc = _open(docx_path)
        target_p = next(
            p for p in doc.paragraphs if p.text.strip().startswith(split_entry_text[:20])
        )
        # Insert page-break 'บรรณานุกรม (ต่อ)' + blank paragraph before target_p
        p_cont = docx.text.paragraph.Paragraph(OxmlElement("w:p"), doc)
        p_cont.style = doc.styles["Normal"]
        p_cont.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cont.paragraph_format.page_break_before = True
        p_cont.paragraph_format.keep_with_next = True
        r_cont = p_cont.add_run("บรรณานุกรม (ต่อ)")
        r_cont.bold = True
        mark_thai(r_cont)

        p_blank = docx.text.paragraph.Paragraph(OxmlElement("w:p"), doc)
        p_blank.style = doc.styles["Normal"]
        p_blank.paragraph_format.keep_with_next = True

        target_p._p.addprevious(p_cont._p)
        target_p._p.addprevious(p_blank._p)
        _save(doc, docx_path)
        print(f"  Inserted 'บรรณานุกรม (ต่อ)' before {split_entry_text!r}")


def collect_entries_from_word(docx_path):
    """Read exact page numbers of front matter sections, chapters/headings (including
    Appendix A headings and Author Bio), tables, and figures via Word COM."""
    word = word_app()
    toc, tables, figures, letters = [], [], [], {}
    try:
        d = word.Documents.Open(os.path.abspath(docx_path))
        d.Fields.Update()
        d.Repaginate()

        # Front matter section start page letters (Sections 3..8 in 1-based Word COM)
        front_names = {
            3: "บทคัดย่อ",
            4: "Abstract",
            5: "กิตติกรรมประกาศ",
            6: "สารบัญ",
            7: "สารบัญตาราง",
            8: "สารบัญรูป",
        }
        for si, name in front_names.items():
            pg_num = d.Sections(si).Range.Paragraphs(1).Range.Information(1)  # wdActiveEndAdjustedPageNumber
            letters[name] = THAI_LETTERS[pg_num - 1]

        # Body sections (Sections 9..15 in 1-based Word COM)
        for si in range(9, d.Sections.Count + 1):
            sec_rng = d.Sections(si).Range
            for p in sec_rng.Paragraphs:
                raw_unstripped = p.Range.Text.replace("\x0b", " ").replace("\r", "").replace("\x07", "")
                raw = raw_unstripped.strip()
                if not raw:
                    continue
                lvl = p.OutlineLevel
                pg = str(p.Range.Information(1))  # adjusted page number (1-based Arabic)
                if lvl == 1:
                    num = p.Range.ListFormat.ListString.strip()
                    if raw.startswith("ภาคผนวก ก"):
                        toc.append(("ภาคผนวก ก คู่มือการใช้งานแอปพลิเคชัน Chatshop", pg, 1))
                    elif raw == "บรรณานุกรม":
                        toc.append(("บรรณานุกรม", pg, 1))
                    else:
                        toc.append((f"{num} {raw}".strip(), pg, 1))
                elif lvl == 2:
                    num = p.Range.ListFormat.ListString.strip()
                    toc.append((f"{num} {raw}".strip() if num else raw, pg, 2))
                elif raw == "ประวัติผู้เขียน" and p.Range.ParagraphFormat.Alignment == 1:
                    toc.append(("ประวัติผู้เขียน", pg, 1))
                elif not raw_unstripped.startswith("\t") and not raw.endswith("(ต่อ)"):
                    if re.match(r"^ตารางที่ \d+\.\d+\s+", raw):
                        tables.append((raw, pg))
                    elif re.match(r"^รูปที่ (?:\d+|ก)\.\d+\s+", raw):
                        figures.append((raw, pg))
        d.Close(False)
    finally:
        word.Quit()
    return toc, tables, figures, letters


if __name__ == "__main__":
    assemble_v6()
