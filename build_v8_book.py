"""
build_v8_book.py
ระบบสร้างเอกสาร Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v8.docx และ v8.pdf
ฟังก์ชันหลัก:
1. นำเข้าจาก Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v7.docx
2. รวมเนื้อหาบทที่ 2 จาก Chatshop_บทที่2_ทฤษฎีและงานวิจัยที่เกี่ยวข้อง.docx
   - กำหนด Style ID w:val="1" สำหรับ "ทฤษฎีและงานวิจัยที่เกี่ยวข้อง" เพื่อผูกกับตัวนับ บทที่ %1 ของ Word
   - กำหนด w:val="2" สำหรับ Heading 2 และ w:val="3" สำหรับ Heading 3
   - รักษาส่วนแบ่งตอน w:sectPr ครบถ้วน
3. ปลด outline / border ของรูปภาพทุกรูปในเอกสารทั้งเล่ม (ไม่มีเส้นขอบใดๆ ทั้งสิ้น)
4. ตั้งค่า SEQ \\r 1 สำหรับรูปแรกของบทที่ 2 (รูปที่ 2.1) และบทที่ 3 (รูปที่ 3.1)
5. ปรับปรุงภาคผนวก ก คู่มือการใช้งานแอปพลิเคชัน:
   - เพิ่ม 4 รูปและหัวข้อเกี่ยวกับการตอบแชทของ AI แต่ละโหมดใน ก.5:
     - ก.5.2 โหมดตอบกลับอัตโนมัติ (AI Autopilot) -> รูปที่ ก.8 (fig_ai_autopilot.jpg)
     - ก.5.3 โหมดแนะนำคำตอบ (AI Suggestion) -> รูปที่ ก.9 (fig_ai_suggestion.jpg) และรูปที่ ก.10 (fig_ai_suggestion_sent.jpg)
     - ก.5.4 เมื่อบอทตอบไม่ได้และส่งต่อ (Escalate) -> รูปที่ ก.11 (fig_ai_escalate.jpg)
     - เลื่อนแผ่นวิเคราะห์บทสนทนาเป็น ก.5.5 -> รูปที่ ก.12
   - เลื่อนลำดับรูปภาพเดิมตั้งแต่ ก.9..ก.22 เป็น ก.13..ก.26 พร้อมอัปเดตข้อความอ้างอิงในเนื้อหาภาคผนวก ก
   - เพิ่ม 9 รูปและหัวข้อการตั้งค่าระบบใน ก.10 เป็น รูปที่ ก.27..ก.35
   - รวมรูปภาพในภาคผนวก ก ทั้งหมด 35 รูป
6. วัดเลขหน้าจริง (Adjusted Page Number) ผ่าน Word COM แบบ Chapter-scoped matching
7. อัปเดตสารบัญเนื้อหา (TOC) แบบ In-place รักษาสไตล์และการเยื้องทั้งหมด 100%
8. สร้างสารบัญตาราง (LOT) โดยไม่ตัดคำตารางที่ 4.4 (บรรทัดเดียวชิดขวาไปหน้า 75)
9. สร้างสารบัญรูป (LOF) ครบ 93 รูป พร้อมแท็บจุดไข่ปลาชิดขวา
10. ส่งออกเป็น PDF ผ่าน Word COM และแปลงหน้าสารบัญเป็นภาพ PNG เพื่อตรวจสอบ
"""

import os
import sys
import shutil
import io
import copy
import re
import docx
from docx.shared import Pt, Inches, Cm
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn, nsdecls
import win32com.client
import pymupdf

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

ROOT = os.path.abspath(os.path.dirname(__file__))
V7_PATH = os.path.join(ROOT, "output", "Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v7.docx")
CH2_SRC = os.path.join(ROOT, "output", "Chatshop_บทที่2_ทฤษฎีและงานวิจัยที่เกี่ยวข้อง.docx")
V8_DOCX = os.path.join(ROOT, "output", "Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v8.docx")
V8_PDF  = os.path.join(ROOT, "output", "Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v8.pdf")
TEMP_MEASURE_DOCX = os.path.join(ROOT, "output", "temp_v8_measure.docx")
APP_FIGS_DIR = os.path.join(ROOT, "output", "appendix_figures")
VERIFY_DIR = os.path.join(ROOT, "process", "temp", "v8_verify")
os.makedirs(VERIFY_DIR, exist_ok=True)

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"

def xml_escape(text):
    if not text:
        return ""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;")
    )

def set_font_thai(r, font_name="TH SarabunPSK", size_pt=16, bold=False, italic=False):
    r.font.name = font_name
    r.font.size = Pt(size_pt)
    r.bold = bold
    r.italic = italic
    rFonts = r._r.get_or_add_rPr().find(qn('w:rFonts'))
    if rFonts is not None:
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)

def fix_styles(doc):
    for s_name in ['Normal', 'toc 1', 'toc 2', 'toc 3', 'Header', 'Footer', 'Caption']:
        try:
            s = doc.styles[s_name]
            s.font.name = "TH SarabunPSK"
            s.font.size = Pt(16)
        except Exception:
            pass

    styles_el = doc.styles.element
    for style in styles_el.findall(qn('w:style')):
        s_id = style.get(qn('w:styleId'), '')
        if s_id in ['1', '2', '3', 'a', 'af9', 'ab']:
            rPr = style.find(qn('w:rPr'))
            if rPr is not None:
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is not None:
                    rFonts.set(qn('w:ascii'), "TH SarabunPSK")
                    rFonts.set(qn('w:hAnsi'), "TH SarabunPSK")
                    rFonts.set(qn('w:cs'), "TH SarabunPSK")
                else:
                    rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'))
                
                sz = rPr.find(qn('w:sz'))
                if sz is not None:
                    sz.set(qn('w:val'), "32")
                else:
                    rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="32"/>'))
                    
                szCs = rPr.find(qn('w:szCs'))
                if szCs is not None:
                    szCs.set(qn('w:val'), "32")
                else:
                    rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="32"/>'))

def remove_all_figure_outlines(doc):
    count = 0
    for spPr in doc.element.body.iter(PIC + "spPr"):
        ln = spPr.find(A + "ln")
        if ln is not None:
            spPr.remove(ln)
            count += 1
    print(f"  Removed {count} figure outlines across document.")

def merge_chapter2_clean(doc):
    print("Merging Chapter 2 clean from Chatshop_บทที่2_ทฤษฎีและงานวิจัยที่เกี่ยวข้อง.docx...")
    d_src = docx.Document(CH2_SRC)

    # 1. Map all image relationships into target doc
    rId_map = {}
    for rId, rel in d_src.part.rels.items():
        if "image" in rel.reltype:
            new_part = doc.part.package.get_or_add_image_part(io.BytesIO(rel.target_part.blob))
            new_rel = doc.part.relate_to(new_part, rel.reltype)
            rId_map[rId] = new_rel

    # 2. Find Ch2 start and Ch3 start in doc body
    p_ch2 = None
    p_ch3 = None
    for p in doc.paragraphs:
        if p.text.strip() == "ทฤษฎีและงานวิจัยที่เกี่ยวข้อง":
            p_ch2 = p._p
        if p.text.strip() == "วิธีดำเนินงาน":
            p_ch3 = p._p
            break

    if p_ch2 is None or p_ch3 is None:
        raise RuntimeError("Could not find Chapter 2 or Chapter 3 boundary in document!")

    body = doc.element.body
    idx_ch2 = body.index(p_ch2)
    idx_ch3 = body.index(p_ch3)

    last_elem = body[idx_ch3 - 1]
    sectPr = None
    for sp in last_elem.iter(qn("w:sectPr")):
        sectPr = copy.deepcopy(sp)

    src_elements = []
    for child in d_src.element.body:
        if child.tag.endswith("sectPr"):
            continue
        el = copy.deepcopy(child)
        for blip in el.iter(qn("a:blip")):
            old_rId = blip.get(qn("r:embed"))
            if old_rId in rId_map:
                blip.set(qn("r:embed"), rId_map[old_rId])
                
        # Make sure no outline on pictures in src_elements
        for spPr in el.iter(PIC + "spPr"):
            ln = spPr.find(A + "ln")
            if ln is not None:
                spPr.remove(ln)

        # CRITICAL FIX: Ensure Style ID matches the thesis master template!
        if el.tag.endswith("p"):
            pPr = el.find(qn("w:pPr"))
            full_txt = "".join(el.itertext()).strip()
            if pPr is not None:
                pStyle = pPr.find(qn("w:pStyle"))
                if pStyle is not None:
                    val = pStyle.get(qn("w:val"))
                    if val == "Heading1" or "ทฤษฎีและงานวิจัยที่เกี่ยวข้อง" in full_txt:
                        pStyle.set(qn("w:val"), "1")
                        spacing = pPr.find(qn("w:spacing"))
                        if spacing is not None:
                            spacing.set(qn("w:before"), "720")
                    elif val == "Heading2":
                        pStyle.set(qn("w:val"), "2")
                    elif val == "Heading3":
                        pStyle.set(qn("w:val"), "3")
                else:
                    if full_txt == "ทฤษฎีและงานวิจัยที่เกี่ยวข้อง":
                        pPr.append(parse_xml(f'<w:pStyle {nsdecls("w")} w:val="1"/>'))
            else:
                if full_txt == "ทฤษฎีและงานวิจัยที่เกี่ยวข้อง":
                    pPr = el.get_or_add_pPr()
                    pPr.append(parse_xml(f'<w:pStyle {nsdecls("w")} w:val="1"/>'))


        src_elements.append(el)

    if sectPr is not None:
        last_src_p = src_elements[-1]
        pPr = last_src_p.get_or_add_pPr()
        pPr.append(sectPr)

    for _ in range(idx_ch3 - idx_ch2):
        elem_to_remove = body[idx_ch2]
        body.remove(elem_to_remove)

    insert_pos = idx_ch2
    for el in src_elements:
        body.insert(insert_pos, el)
        insert_pos += 1

    print(f"  Successfully inserted {len(src_elements)} elements for clean Chapter 2.")

def fix_abstract_indentation(doc):
    print("Fixing English Abstract Project Title hanging indentation (1620 twips)...")
    fixed = 0
    for p in doc.paragraphs[:120]:
        txt = p.text.strip()
        if txt.startswith("Project Title") and "Centralized Chat" in txt:
            pPr = p._p.get_or_add_pPr()
            old_tabs = pPr.find(qn("w:tabs"))
            if old_tabs is not None:
                pPr.remove(old_tabs)
            old_ind = pPr.find(qn("w:ind"))
            if old_ind is not None:
                pPr.remove(old_ind)
            pPr.append(parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="1620"/></w:tabs>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:left="1620" w:hanging="1620"/>'))
            fixed += 1
    print(f"  Fixed {fixed} Project Title paragraph(s) in Abstract.")

def fix_figure_captions_and_seq(doc):
    print("Fixing figure captions and SEQ fields across all chapters...")
    ch_counts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}
    seen_p = set()

    for t in doc.tables:
        for r in t.rows:
            for cell in r.cells:
                for p in cell.paragraphs:
                    if p._p in seen_p:
                        continue
                    seen_p.add(p._p)
                    m_ch = re.search(r'รูปที่\s+(\d+)\.', p.text)
                    if not m_ch:
                        continue
                    curr_ch = int(m_ch.group(1))
                    fld_simples = [node for node in p._p.iter() if node.tag == qn('w:fldSimple')]
                    for fld in fld_simples:
                        instr = fld.get(qn('w:instr'), '')
                        if 'SEQ รูปที่' in instr:
                            ch_counts[curr_ch] = ch_counts.get(curr_ch, 0) + 1
                            num = ch_counts[curr_ch]
                            if num == 1:
                                fld.set(qn('w:instr'), ' SEQ รูปที่ \\r 1 \\* ARABIC ')
                            else:
                                fld.set(qn('w:instr'), ' SEQ รูปที่ \\* ARABIC ')
                            for t_node in fld.iter(qn('w:t')):
                                t_node.text = str(num)

    print(f"  SEQ fields and figure numbers fixed per chapter: {ch_counts}")
    assert ch_counts[2] == 19, f"Expected 19 figures in Ch2, got {ch_counts[2]}"
    assert ch_counts[3] == 38, f"Expected 38 figures in Ch3, got {ch_counts[3]}"

def make_heading_p(num_str, title_str):
    return parse_xml(
        f'<w:p {nsdecls("w")}>'
        f'<w:pPr>'
        f'<w:spacing w:before="120" w:after="60" w:line="240" w:lineRule="auto"/>'
        f'<w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK" w:hint="cs"/>'
        f'<w:sz w:val="32"/><w:szCs w:val="32"/><w:cs/><w:lang w:bidi="th-TH"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK" w:hint="cs"/>'
        f'<w:sz w:val="32"/><w:szCs w:val="32"/><w:cs/><w:lang w:bidi="th-TH"/>'
        f'</w:rPr>'
        f'<w:tab/><w:t>{xml_escape(num_str)}  {xml_escape(title_str)}</w:t>'
        f'</w:r>'
        f'</w:p>'
    )


def make_appendix_body(text_str):
    return parse_xml(
        f'<w:p {nsdecls("w")}>'
        f'<w:pPr>'
        f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:ind w:left="0" w:firstLine="567"/>'
        f'<w:jc w:val="both"/>'
        f'<w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK" w:hint="cs"/>'
        f'<w:sz w:val="32"/><w:szCs w:val="32"/><w:cs/><w:lang w:bidi="th-TH"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK" w:hint="cs"/>'
        f'<w:sz w:val="32"/><w:szCs w:val="32"/><w:cs/><w:lang w:bidi="th-TH"/>'
        f'</w:rPr>'
        f'<w:t>{xml_escape(text_str.strip())}</w:t>'
        f'</w:r>'
        f'</w:p>'
    )

def make_empty_p():
    return parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>')

def create_figure_table(doc, img_filename, caption_text, width_cm=4.6):
    img_full_path = os.path.join(APP_FIGS_DIR, img_filename)
    temp_p = doc.add_paragraph()
    r_temp = temp_p.add_run()
    r_temp.add_picture(img_full_path, width=Cm(width_cm))
    drawings = [el for el in r_temp._r.iter() if el.tag == qn('w:drawing')]
    drawing_elem = drawings[0]
    temp_p._p.getparent().remove(temp_p._p)
    
    # Remove any ln in drawing_elem
    for spPr in drawing_elem.iter(PIC + "spPr"):
        ln = spPr.find(A + "ln")
        if ln is not None:
            spPr.remove(ln)

    tbl_xml = (
        f'<w:tbl {nsdecls("w")}>'
        f'<w:tblPr>'
        f'<w:tblStyle w:val="ad"/>'
        f'<w:tblW w:w="0" w:type="auto"/>'
        f'<w:tblBorders>'
        f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>'
        f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        f'</w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="8306"/></w:tblGrid>'
        f'<w:tr><w:trPr><w:cantSplit/></w:trPr>'
        f'<w:tc>'
        f'<w:tcPr><w:tcW w:w="8306" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr><w:keepNext/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr></w:r></w:p>'
        f'<w:p><w:pPr><w:keepNext/></w:pPr></w:p>'
        f'<w:p><w:pPr><w:pStyle w:val="ab"/><w:keepNext/><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr>'
        f'<w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK" w:hint="cs"/>'
        f'<w:sz w:val="32"/><w:szCs w:val="32"/><w:cs/><w:lang w:bidi="th-TH"/>'
        f'</w:rPr>'
        f'<w:t>{xml_escape(caption_text)}</w:t></w:r></w:p>'
        f'</w:tc></w:tr></w:tbl>'
    )
    tbl_elem = parse_xml(tbl_xml)
    first_r = [el for el in tbl_elem.iter() if el.tag == qn('w:r')][0]
    first_r.append(drawing_elem)
    return tbl_elem

def build_appendix_content(doc):
    print("Building complete Appendix A content (AI modes and Settings)...")
    body = doc.element.body

    # 1. Update text references in Appendix A body from รูปที่ ก.X (X >= 9) to ก.{X+4}
    # and ก.8 to ก.12 (in descending order)
    in_app = False
    for p in doc.paragraphs:
        txt = p.text.strip()
        if "ภาคผนวก ก" in txt:
            in_app = True
        if not in_app:
            continue
            
        for num in range(22, 7, -1):
            old_str = f"รูปที่ ก.{num}"
            new_num = num + 4
            new_str = f"รูปที่ ก.{new_num}"
            if old_str in p.text:
                for r in p.runs:
                    if old_str in r.text:
                        r.text = r.text.replace(old_str, new_str)
                # also check full p.text in case runs split
                if old_str in p.text:
                    p.text = p.text.replace(old_str, new_str)

    # 2. Shift existing Appendix A figure tables captions from ก.8..ก.22
    for t in doc.tables:
        txt = t.rows[0].cells[0].text.strip()
        if "รูปที่ ก." in txt:
            m = re.search(r'รูปที่ ก\.(\d+)', txt)
            if m:
                num = int(m.group(1))
                if num >= 8:
                    old_caption_num = f"รูปที่ ก.{num}"
                    new_caption_num = f"รูปที่ ก.{num + 4}"
                    for p in t.rows[0].cells[0].paragraphs:
                        if old_caption_num in p.text:
                            for r in p.runs:
                                if old_caption_num in r.text:
                                    r.text = r.text.replace(old_caption_num, new_caption_num)
                            if old_caption_num in p.text:
                                p.text = p.text.replace(old_caption_num, new_caption_num)

    # 3. Insert the 4 AI mode figures and subsections into ก.5
    # Find table of รูปที่ ก.7
    tbl_a7_idx = None
    for idx, el in enumerate(body):
        if el.tag.endswith("tbl"):
            t_txt = "".join(el.itertext()).strip()
            if "รูปที่ ก.7" in t_txt:
                tbl_a7_idx = idx
                break

    if tbl_a7_idx is None:
        raise RuntimeError("Could not find Table for รูปที่ ก.7 in Appendix A!")

    ai_sections = [
        {
            "num": "ก.5.2",
            "title": "การทำงานของแชทบอทในโหมดตอบกลับอัตโนมัติ (AI Autopilot)",
            "paragraphs": [
                "เมื่อเปิดใช้งานโหมด Autopilot สำหรับบทสนทนาใด แชทบอทจะทำหน้าที่ประมวลผลข้อความของลูกค้าและตอบกลับอัตโนมัติทันที โดยดึงข้อมูลจากแคตตาล็อกสินค้า นโยบายร้านค้า และฐานข้อมูลความรู้ที่เชื่อมโยงกับธุรกิจนั้น เช่น เมื่อลูกค้าสอบถามรายการสินค้า แชทบอทจะค้นหาและตอบกลับด้วยรายการสินค้าพร้อมราคาอย่างถูกต้อง พร้อมแสดงแท็กระบุการตอบกลับอัตโนมัติ (AUTO-REPLY) ในหน้าต่างสนทนา และปรากฏแถบสถานะ AI Autopilot is replying for you บริเวณส่วนบนของหน้าจอ ทั้งนี้เจ้าหน้าที่สามารถแตะปุ่ม Take over เพื่อปิดโหมดอัตโนมัติและเข้าดูแลบทสนทนาด้วยตนเองได้ทันที หน้าต่างสนทนาขณะแชทบอทตอบกลับอัตโนมัติแสดงดังรูปที่ ก.8"
            ],
            "fig_file": "fig_ai_autopilot.jpg",
            "fig_caption": "รูปที่ ก.8  การตอบกลับข้อความอัตโนมัติของแชทบอทในโหมด AI Autopilot"
        },
        {
            "num": "ก.5.3",
            "title": "การแนะนำคำตอบและการส่งข้อความในโหมดแนะนำ (AI Suggestion)",
            "paragraphs": [
                "ในโหมด AI Suggestion เมื่อลูกค้าส่งข้อความเข้ามา แชทบอทจะไม่ส่งข้อความตอบกลับไปยังลูกค้าโดยตรง แต่จะทำหน้าที่เป็นผู้ช่วยร่างคำตอบให้แก่เจ้าหน้าที่ โดยระบบจะประมวลผลและนำเสนอข้อความคำตอบที่เหมาะสมในกล่อง AI SUGGESTION บริเวณด้านล่างของหน้าจอ พร้อมแสดงตัวเลือกการส่งข้อความ เจ้าหน้าที่สามารถแตะปุ่ม Send as-is เพื่อส่งข้อความที่แชทบอทแนะนำไปยังลูกค้าได้ทันที หรือแตะปุ่ม Edit · แก้ไข เพื่อปรับแต่งข้อความก่อนส่ง หน้าต่างแนะนำคำตอบของแชทบอทแสดงดังรูปที่ ก.9"
            ],
            "fig_file": "fig_ai_suggestion.jpg",
            "fig_caption": "รูปที่ ก.9  แผ่นแนะนำข้อความคำตอบของแชทบอทในโหมด AI Suggestion"
        },
        {
            "num": "ก.5.4",
            "title": "การส่งข้อความคำตอบที่แชทบอทแนะนำไปยังบทสนทนา",
            "paragraphs": [
                "เมื่อเจ้าหน้าที่แตะปุ่ม Send as-is จากกล่องข้อความแนะนำ ข้อความดังกล่าวจะถูกส่งไปยังลูกค้าผ่านช่องทางที่ลูกค้าติดต่อเข้ามาทันที โดยแสดงเป็นข้อความสนทนาฝั่งร้านค้าอย่างต่อเนื่อง ช่วยลดภาระและประหยัดเวลาของเจ้าหน้าที่ในการพิมพ์ตอบคำถามที่พบบ่อยได้อย่างมีประสิทธิภาพ ผลลัพธ์หลังการส่งข้อความคำตอบที่แชทบอทแนะนำแสดงดังรูปที่ ก.10"
            ],
            "fig_file": "fig_ai_suggestion_sent.jpg",
            "fig_caption": "รูปที่ ก.10  การส่งข้อความคำตอบที่แชทบอทแนะนำไปยังบทสนทนา"
        },
        {
            "num": "ก.5.5",
            "title": "การส่งต่อบทสนทนาเมื่อแชทบอทไม่สามารถตอบได้ (Escalate)",
            "paragraphs": [
                "เมื่อลูกค้าส่งข้อความที่มีความซับซ้อน ข้อความที่แชทบอทไม่เข้าใจ หรืออยู่นอกเหนือขอบเขตข้อมูลความรู้ที่กำหนดไว้ในระบบ แชทบอทจะไม่สร้างคำตอบที่ไม่แน่นอน แต่จะส่งข้อความแจ้งลูกค้าว่า 'บอทตอบไม่ได้ — รอแอดมินช่วยตอบ' เพื่อให้ลูกค้ารับทราบสถานะ พร้อมกันนั้นระบบจะส่งคำสั่งส่งต่อบทสนทนา (<<ESCALATE>>) และเปลี่ยนแท็กของบทสนทนาเป็น 'รอแอดมิน' เพื่อแจ้งเตือนให้เจ้าหน้าที่ร้านค้าทราบว่ามีบทสนทนาที่ต้องการความช่วยเหลือจากมนุษย์ และเข้ามารับช่วงต่อในการให้บริการลูกค้าได้อย่างทันท่วงที หน้าต่างสนทนาเมื่อแชทบอทไม่สามารถตอบได้และส่งต่อแสดงดังรูปที่ ก.11"
            ],
            "fig_file": "fig_ai_escalate.jpg",
            "fig_caption": "รูปที่ ก.11  การส่งต่อบทสนทนาเมื่อแชทบอทไม่สามารถตอบได้"
        }
    ]

    insert_pos = tbl_a7_idx + 1
    # Check if next element is empty p
    if insert_pos < len(body) and body[insert_pos].tag.endswith("p") and not "".join(body[insert_pos].itertext()).strip():
        insert_pos += 1

    for s in ai_sections:
        h_elem = make_heading_p(s["num"], s["title"])
        body.insert(insert_pos, h_elem)
        insert_pos += 1
        
        for p_text in s["paragraphs"]:
            body.insert(insert_pos, make_appendix_body(p_text))
            insert_pos += 1
            
        body.insert(insert_pos, make_empty_p())
        insert_pos += 1
        
        tbl_element = create_figure_table(doc, s["fig_file"], s["fig_caption"], width_cm=4.6)
        body.insert(insert_pos, tbl_element)
        insert_pos += 1
        
        body.insert(insert_pos, make_empty_p())
        insert_pos += 1

    # Update former ก.5.2 to ก.5.6 (การดูแผ่นวิเคราะห์บทสนทนา)
    for p in doc.paragraphs:
        if "ก.5.2 การดูแผ่นวิเคราะห์บทสนทนา" in p.text:
            p.text = p.text.replace("ก.5.2 การดูแผ่นวิเคราะห์บทสนทนา", "ก.5.6 การดูแผ่นวิเคราะห์บทสนทนา")
            for r in p.runs:
                set_font_thai(r, "TH SarabunPSK", 16)
            break

    # 4. Under ก.10, insert the 9 Settings subsections and figure tables (รูปที่ ก.27..ก.35)
    target_idx = None
    for idx, p in enumerate(doc.paragraphs):
        if p.text.strip() == "ก.10 ข้อมูลผู้ใช้และการตั้งค่าระบบ":
            target_idx = idx
            break
            
    if target_idx is None:
        raise RuntimeError("Could not locate 'ก.10 ข้อมูลผู้ใช้และการตั้งค่าระบบ' in document!")

    p_a10_elem = doc.paragraphs[target_idx]._p
    insert_pos_a10 = body.index(p_a10_elem) + 1
    
    # Check if table or empty p after
    while insert_pos_a10 < len(body):
        child_txt = "".join(body[insert_pos_a10].itertext()).strip()
        if body[insert_pos_a10].tag.endswith("tbl") and "รูปที่ ก.26" in child_txt:
            insert_pos_a10 += 1
            if insert_pos_a10 < len(body) and body[insert_pos_a10].tag.endswith("p") and not "".join(body[insert_pos_a10].itertext()).strip():
                insert_pos_a10 += 1
            break
        insert_pos_a10 += 1

    settings_sections = [
        {
            "num": "ก.10.3",
            "title": "การตั้งค่าข้อมูลทั่วไปของร้านค้า",
            "paragraphs": [
                "เมนูร้านค้าในหมวดการตั้งค่าร้านค้า ใช้สำหรับจัดการข้อมูลอัตลักษณ์และการตั้งค่าพื้นฐานของร้านค้า ประกอบด้วย ชื่อร้านค้าสำหรับระบุชื่อพื้นที่ทำงานขององค์กร รหัสย่อร้านค้าซึ่งระบบสร้างขึ้นเป็นค่าเฉพาะสำหรับใช้อ้างอิงในระบบ ลิงก์รูปภาพสัญลักษณ์ร้านค้าสำหรับแสดงตราสัญลักษณ์ประจำร้าน การเลือกภาษาที่ใช้แสดงผลในแอปพลิเคชันซึ่งรองรับภาษาไทยและภาษาอังกฤษ และการแสดงประเภทบัญชีผู้ใช้ เมื่อปรับปรุงข้อมูลเรียบร้อยแล้วสามารถแตะปุ่มบันทึกเพื่อบันทึกข้อมูลเข้าสู่ระบบ หน้าตั้งค่าข้อมูลทั่วไปแสดงดังรูปที่ ก.27"
            ],
            "fig_file": "fig_a22_general.jpg",
            "fig_caption": "รูปที่ ก.27  หน้าตั้งค่าข้อมูลทั่วไปของร้านค้า"
        },
        {
            "num": "ก.10.4",
            "title": "การจัดการแท็กของร้านค้า",
            "paragraphs": [
                "เมนูแท็กในหมวดการตั้งค่าร้านค้า ใช้สำหรับบริหารจัดการแท็กหรือป้ายกำกับของพื้นที่ทำงาน เพื่อใช้จัดกลุ่มและจำแนกประเภทบทสนทนารวมถึงกลุ่มลูกค้า เช่น แท็กสำหรับลูกค้าประจำพร้อมแถบสีประจำแท็ก เจ้าหน้าที่สามารถแตะปุ่มเครื่องหมายบวกที่มุมขวาบนเพื่อสร้างแท็กใหม่ พร้อมทั้งกำหนดชื่อและเลือกสีประจำแท็ก และสามารถแตะไอคอนรูปถังขยะเพื่อลบแท็กที่ไม่ต้องการใช้งานออกจากระบบ ทั้งนี้เมื่อมีการลบแท็ก ระบบจะถอดแท็กดังกล่าวออกจากทุกบทสนทนาโดยอัตโนมัติ หน้าจัดการแท็กแสดงดังรูปที่ ก.28"
            ],
            "fig_file": "fig_a23_tags.jpg",
            "fig_caption": "รูปที่ ก.28  หน้าจัดการแท็กของร้านค้า"
        },
        {
            "num": "ก.10.5",
            "title": "การตั้งค่าระบบอัตโนมัติ",
            "paragraphs": [
                "เมนูระบบอัตโนมัติในหมวดการตั้งค่าร้านค้า ใช้สำหรับควบคุมพฤติกรรมการทำงานของระบบและแชทบอทในภาพรวม ประกอบด้วย การเปิดใช้งานสวิตช์ล็อกบทสนทนาอัตโนมัติ ซึ่งเมื่อเปิดใช้งาน ระบบจะทำการล็อกบทสนทนาให้อยู่ในความดูแลของเจ้าหน้าที่ทันทีที่เจ้าหน้าที่เริ่มพิมพ์ตอบข้อความลูกค้า เพื่อป้องกันไม่ให้เจ้าหน้าที่คนอื่นตอบซ้อน และการเลือกโหมดการทำงานเริ่มต้นของแชทบอท ได้แก่ ปิดการทำงาน (Manual), แนะนำคำตอบ (Suggest) หรือตอบกลับอัตโนมัติ (Autopilot) หน้าตั้งค่าระบบอัตโนมัติแสดงดังรูปที่ ก.29"
            ],
            "fig_file": "fig_a24_automation.jpg",
            "fig_caption": "รูปที่ ก.29  หน้าตั้งค่าระบบอัตโนมัติ"
        },
        {
            "num": "ก.10.6",
            "title": "การตั้งค่าแชทบอทและการเชื่อมต่อแบบจำลองภาษา",
            "paragraphs": [
                "เมนูแชทบอทในหมวดการตั้งค่าร้านค้า ใช้สำหรับกำหนดค่าและเชื่อมต่อแบบจำลองภาษาขนาดใหญ่ (LLM) ประกอบด้วย การเลือกผู้ให้บริการแชทบอท เช่น Ollama, Anthropic, Gemini หรือ OpenAI การระบุ API Key เพื่อยืนยันสิทธิ์การใช้งาน การระบุชื่อแบบจำลองที่ต้องการเรียกใช้ เช่น qwen3:14b และการแตะปุ่มทดสอบการเชื่อมต่อเพื่อตรวจสอบความพร้อมในการทำงานของแบบจำลอง เมื่อตั้งค่าสำเร็จแชทบอทจะสามารถประมวลผลข้อความและตอบกลับลูกค้าได้อย่างถูกต้อง หน้าตั้งค่าแชทบอทแสดงดังรูปที่ ก.30"
            ],
            "fig_file": "fig_a25_ai_chatbot.jpg",
            "fig_caption": "รูปที่ ก.30  หน้าตั้งค่าแชทบอท"
        },
        {
            "num": "ก.10.7",
            "title": "การจัดการธุรกิจ",
            "paragraphs": [
                "เมนูธุรกิจในหมวดการตั้งค่าร้านค้า ใช้สำหรับบริหารจัดการหน่วยธุรกิจย่อยภายในร้านค้า โดยแสดงรายการธุรกิจที่เปิดใช้งานอยู่ เช่น ร้านหลักและร้านสาขา พร้อมแสดงสถานะการใช้งานและจำนวนช่องทางที่ผูกกับธุรกิจนั้น เจ้าหน้าที่สามารถแตะปุ่มเครื่องหมายบวกเพื่อเพิ่มธุรกิจใหม่ หรือแตะรายการธุรกิจเพื่อเข้าไปแก้ไขรายละเอียด หน้าจัดการธุรกิจแสดงดังรูปที่ ก.31"
            ],
            "fig_file": "fig_a26_businesses.jpg",
            "fig_caption": "รูปที่ ก.31  หน้าจัดการธุรกิจ"
        },
        {
            "num": "ก.10.8",
            "title": "การแก้ไขและตั้งค่ารายละเอียดธุรกิจ",
            "paragraphs": [
                "เมื่อแตะเลือกหน่วยธุรกิจในหน้าจัดการธุรกิจ ระบบจะเปิดหน้าแก้ไขข้อมูลธุรกิจ ประกอบด้วย การแก้ไขชื่อธุรกิจ การกำหนดบุคลิกภาพหรือคำสั่งเฉพาะสำหรับแชทบอทประจำธุรกิจ ข้อมูลบริบทของร้านค้า เช่น นโยบายการจัดส่งสินค้า ช่องทางการชำระเงิน และบัญชีธนาคารสำหรับแจ้งลูกค้า เพื่อให้แชทบอทนำข้อมูลเหล่านี้ไปใช้ในการตอบคำถามได้อย่างถูกต้องแม่นยำ หน้าแก้ไขรายละเอียดธุรกิจแสดงดังรูปที่ ก.32"
            ],
            "fig_file": "fig_a27_edit_business.jpg",
            "fig_caption": "รูปที่ ก.32  หน้าแก้ไขรายละเอียดธุรกิจ"
        },
        {
            "num": "ก.10.9",
            "title": "การตั้งค่าการแจ้งเตือน",
            "paragraphs": [
                "เมนูการแจ้งเตือนในหมวดการแจ้งเตือนของแท็บผู้ใช้ ใช้สำหรับปรับแต่งการรับข้อความแจ้งเตือนผ่านโทรศัพท์สมาร์ตโฟน ประกอบด้วย สวิตช์เปิดปิดการแจ้งเตือนแบบ Push Notification สวิตช์เปิดปิดเสียงแจ้งเตือน ตัวเลือกการรับแจ้งเตือนเฉพาะบทสนทนาที่ได้รับมอบหมาย และการเลือกปิดการแจ้งเตือนสำหรับช่องทางที่ไม่ต้องการติดตาม ช่วยให้เจ้าหน้าที่เลือกรับเฉพาะการแจ้งเตือนที่เกี่ยวข้องกับหน้าที่ได้อย่างมีประสิทธิภาพ หน้าตั้งค่าการแจ้งเตือนแสดงดังรูปที่ ก.33"
            ],
            "fig_file": "fig_a28_notifications.jpg",
            "fig_caption": "รูปที่ ก.33  หน้าตั้งค่าการแจ้งเตือน"
        },
        {
            "num": "ก.10.10",
            "title": "การตั้งค่าการเตือนแชทค้างตอบ",
            "paragraphs": [
                "เมนูแชทค้างตอบในหมวดการแจ้งเตือน ใช้สำหรับควบคุมการแจ้งเตือนเมื่อมีลูกค้าส่งข้อความเข้ามาแล้วยังไม่ได้รับการตอบกลับ ประกอบด้วย สวิตช์เปิดปิดการแจ้งเตือนแชทค้างตอบ การเลือกเกณฑ์ระยะเวลาที่ยอมรับได้ เช่น 5 นาที, 10 นาที หรือ 15 นาที และการเลือกช่องทางที่ต้องการตรวจสอบ เพื่อให้เจ้าหน้าที่สามารถเข้าตอบข้อความลูกค้าได้ทันตามมาตรฐานการให้บริการที่กำหนด หน้าตั้งค่าการเตือนแชทค้างตอบแสดงดังรูปที่ ก.34"
            ],
            "fig_file": "fig_a29_sla_alerts.jpg",
            "fig_caption": "รูปที่ ก.34  หน้าตั้งค่าการเตือนแชทค้างตอบ"
        },
        {
            "num": "ก.10.11",
            "title": "หน้าข้อมูลบัญชีผู้ใช้",
            "paragraphs": [
                "เมนูโปรไฟล์ในหมวดบัญชีผู้ใช้ แสดงข้อมูลส่วนบุคคลของเจ้าหน้าที่ที่เข้าสู่ระบบ ประกอบด้วย รูปโปรไฟล์ ชื่อ นามสกุล อีเมลประจำบัญชี บทบาทหน้าที่ในพื้นที่ทำงาน และปุ่มสำหรับออกจากระบบ (Sign Out) เพื่อความปลอดภัยในการใช้งานเมื่อสิ้นสุดการปฏิบัติงาน หน้าข้อมูลบัญชีผู้ใช้แสดงดังรูปที่ ก.35"
            ],
            "fig_file": "fig_a30_profile.jpg",
            "fig_caption": "รูปที่ ก.35  หน้าข้อมูลบัญชีผู้ใช้"
        }
    ]

    for s in settings_sections:
        h_elem = make_heading_p(s["num"], s["title"])
        body.insert(insert_pos_a10, h_elem)
        insert_pos_a10 += 1
        
        for p_text in s["paragraphs"]:
            body.insert(insert_pos_a10, make_appendix_body(p_text))
            insert_pos_a10 += 1
            
        body.insert(insert_pos_a10, make_empty_p())
        insert_pos_a10 += 1
        
        tbl_element = create_figure_table(doc, s["fig_file"], s["fig_caption"], width_cm=5.42)
        body.insert(insert_pos_a10, tbl_element)
        insert_pos_a10 += 1
        
        body.insert(insert_pos_a10, make_empty_p())
        insert_pos_a10 += 1

    p_pb = parse_xml(f'<w:p {nsdecls("w")}><w:r><w:br w:type="page"/></w:r></w:p>')
    body.insert(insert_pos_a10, p_pb)
    print("  Appendix A building completed: 35 figures total.")

def _p_full_text(p_elem):
    return "".join(el.text or "" for el in p_elem.iter(qn("w:t"))).strip()

def _ch4_captioned_tables(doc):
    out = []
    for i, t in enumerate(doc.tables):
        prev = t._tbl.getprevious()
        if prev is None or prev.tag != qn("w:p"):
            continue
        text = _p_full_text(prev)
        m = re.match(r"^(ตารางที่ 4\.\d+)\s+(.*)$", text)
        if not m:
            continue
        code = m.group(1)
        title = m.group(2).strip()
        if title.endswith("(ต่อ)"):
            title = title[:-5].strip()
        out.append((i, f"{code}  {title}"))
    return out

def _split_ch4_table(doc, tbl_index, split_after_row, cont_caption_text):
    t = doc.tables[tbl_index]
    parent = t._tbl.getparent()
    idx = parent.index(t._tbl)
    tbl2 = copy.deepcopy(t._tbl)
    t2 = docx.table.Table(tbl2, doc)
    keep = split_after_row + 1  # header row 0 + data rows 1..split_after_row
    for _ in range(len(t.rows) - keep):
        t._tbl.remove(t.rows[keep]._tr)
    for _ in range(split_after_row):
        t2._tbl.remove(t2.rows[1]._tr)

    p_xml = (
        f'<w:p {nsdecls("w")}>'
        f'<w:pPr><w:pStyle w:val="ab"/><w:keepNext/><w:pageBreakBefore/><w:jc w:val="left"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:hint="cs"/><w:cs/></w:rPr>'
        f'<w:t xml:space="preserve">{xml_escape(cont_caption_text)}</w:t></w:r>'
        f'</w:p>'
    )
    p_elem = parse_xml(p_xml)
    parent.insert(idx + 1, p_elem)
    parent.insert(idx + 2, tbl2)

def prepare_chapter4_layout(doc):
    print("Preparing Chapter 4 continuous layout (removing forced page breaks on 4.1..4.18)...")
    # Remove bold from any existing Chapter 3 (ต่อ) captions so they match main captions
    for p in list(doc.paragraphs):
        t = _p_full_text(p._p)
        if t.startswith("ตารางที่ 3.") and t.endswith("(ต่อ)"):
            for r in p.runs:
                r.bold = False

    # 1. Unsplit any previously split Chapter 4 (ต่อ) tables before re-paginating
    unsplit_count = 0
    for p in list(doc.paragraphs):
        t = _p_full_text(p._p)
        if t.startswith("ตารางที่ 4.") and t.endswith("(ต่อ)"):
            prev, nxt = p._p.getprevious(), p._p.getnext()
            if prev is not None and nxt is not None and prev.tag == qn("w:tbl") and nxt.tag == qn("w:tbl"):
                for tr in nxt.findall(qn("w:tr"))[1:]:
                    prev.append(tr)
                nxt.getparent().remove(nxt)
                p._p.getparent().remove(p._p)
                unsplit_count += 1
    if unsplit_count:
        print(f"  Unsplit {unsplit_count} previously split Chapter 4 table(s).")

    # 2. Remove pageBreakBefore from Heading 2 (4.1..4.18) in Chapter 4 and ensure 1 blank line before each H2
    in_ch4 = False
    h2_fixed = 0
    for p in list(doc.paragraphs):
        t = _p_full_text(p._p)
        pPr = p._p.find(qn("w:pPr"))
        pStyle = (
            pPr.find(qn("w:pStyle")).get(qn("w:val"))
            if (pPr is not None and pPr.find(qn("w:pStyle")) is not None)
            else ""
        )
        if pStyle == "1" and "ผลการดำเนินงานและการทดสอบ" in t:
            in_ch4 = True
            continue
        if pStyle == "1" and in_ch4 and "สรุปผล" in t:
            in_ch4 = False
            break
        if not in_ch4:
            continue

        if pStyle == "2":
            p.paragraph_format.page_break_before = False
            p.paragraph_format.keep_with_next = True
            prev = p._p.getprevious()
            if prev is not None and prev.tag == qn("w:p"):
                prev_t = _p_full_text(prev)
                if prev_t != "":
                    empty_p = parse_xml(
                        f'<w:p {nsdecls("w")}><w:pPr><w:rPr>'
                        f'<w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                        f'<w:sz w:val="32"/><w:szCs w:val="32"/>'
                        f'</w:rPr></w:pPr></w:p>'
                    )
                    p._p.addprevious(empty_p)
            h2_fixed += 1
        elif re.match(r"^4\.17\.[1-5]\s", t):
            p.paragraph_format.page_break_before = False
            p.paragraph_format.keep_with_next = False
        elif t.startswith("ตารางที่ 4."):
            p.paragraph_format.page_break_before = False
            p.paragraph_format.keep_with_next = True

    # 3. Ensure cantSplit on all rows and keep_with_next only on header row (row 0) for Chapter 4 tables
    ch4_tbls = _ch4_captioned_tables(doc)
    for i, _ in ch4_tbls:
        t = doc.tables[i]
        for r_idx, row in enumerate(t.rows):
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))
            for cell in row.cells:
                for cp in cell.paragraphs:
                    cp.paragraph_format.keep_with_next = (r_idx == 0)
        docx.text.paragraph.Paragraph(t._tbl.getprevious(), doc).paragraph_format.keep_with_next = True

    print(f"  Prepared {h2_fixed} Chapter 4 headings and {len(ch4_tbls)} tables for continuous flow.")

def _save_docx_retry(doc, path, tries=10):
    import time
    for n in range(tries):
        try:
            return doc.save(path)
        except OSError:
            if n == tries - 1:
                raise
            time.sleep(0.8)

def measure_actual_page_numbers(doc_obj, docx_path):
    print("Splitting overflowing Chapter 4 tables & measuring exact page numbers via Word COM...")
    toc_titles = []
    current = False
    for p in doc_obj.paragraphs[:400]:
        t = p.text.strip()
        if t == "สารบัญ":
            current = True
        elif t == "สารบัญตาราง":
            break
        elif current and t and "\t" in p.text:
            title = p.text.split("\t")[0].strip()
            toc_titles.append(title)

    def get_search_phrase(item):
        if item.startswith("บทที่ "):
            parts = item.split(" ", 2)
            return parts[2] if len(parts) > 2 else parts[1]
        if item.startswith("ภาคผนวก "):
            return "คู่มือการใช้งานแอปพลิเคชัน Chatshop"
        if item.startswith("ก."):
            return item
        if item == "ประวัติผู้เขียน":
            return "ประวัติผู้เขียน"
        parts = item.split(" ", 1)
        if len(parts) > 1 and parts[0][0].isdigit():
            return parts[1].strip()
        return item

    searches = {item: get_search_phrase(item) for item in toc_titles}
    abs_path = os.path.abspath(docx_path)

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        splits = 0
        for round_idx in range(40):
            doc = word.Documents.Open(abs_path)
            doc.Repaginate()
            hit = None
            for i, caption in _ch4_captioned_tables(doc_obj):
                tbl = doc.Tables(i + 1)
                rows = tbl.Rows
                first = rows(1).Range.Information(3)  # wdActiveEndPageNumber
                for r in range(3, rows.Count + 1):
                    if rows(r).Range.Information(3) > first:
                        hit = (i, r - 2, caption)
                        break
                if hit:
                    break
            if not hit:
                print(f"  Chapter 4 table splitting completed ({splits} split(s) total).")
                break
            doc.Close(False)
            i, keep, caption = hit
            _split_ch4_table(doc_obj, i, keep, caption + " (ต่อ)")
            _save_docx_retry(doc_obj, abs_path)
            splits += 1
            print(f"  Split {caption} after data row {keep} -> added (ต่อ) + repeated header")
        
        toc_pages = {}
        lot_pages = {}
        lof_pages = {}
        
        # Preliminary front matter
        toc_pages["บทคัดย่อ"] = "ง"
        toc_pages["Abstract"] = "จ"
        toc_pages["กิตติกรรมประกาศ"] = "ฉ"
        toc_pages["สารบัญ"] = "ช"
        toc_pages["สารบัญตาราง"] = "ฎ"
        toc_pages["สารบัญรูป"] = "ฐ"
        
        body = []
        found_intro = False
        for p in doc.Paragraphs:
            txt = p.Range.Text.strip()
            if not found_intro:
                if txt == "บทนำ":
                    found_intro = True
                else:
                    continue
                    
            if not txt:
                continue
                
            pg = p.Range.Information(1)
            body.append((txt, pg))
            
            # Check LOT
            if txt.startswith("ตารางที่ ") and "(ต่อ)" not in txt:
                m_tbl = re.match(r'^(ตารางที่ \d+\.\d+)', txt)
                if m_tbl:
                    t_num = m_tbl.group(1)
                    if t_num not in lot_pages:
                        lot_pages[t_num] = str(pg)
                        
            # Check LOF
            if txt.startswith("รูปที่ ") and "(ต่อ)" not in txt:
                m_fig = re.match(r'^(รูปที่ (?:[^\s]+))', txt)
                if m_fig:
                    f_num = m_fig.group(1)
                    if f_num not in lof_pages:
                        lof_pages[f_num] = str(pg)

        # Split body paragraphs by chapter to avoid cross-chapter false matches
        ch_ranges = {1: [], 2: [], 3: [], 4: [], 5: [], "app": [], "back": []}
        curr_ch = 1
        for txt, pg in body:
            if txt == "ทฤษฎีและงานวิจัยที่เกี่ยวข้อง":
                curr_ch = 2
            elif txt == "วิธีดำเนินงาน":
                curr_ch = 3
            elif txt == "ผลการดำเนินงานและการทดสอบ":
                curr_ch = 4
            elif txt == "สรุปผลและข้อเสนอแนะ":
                curr_ch = 5
            elif txt == "บรรณานุกรม":
                curr_ch = "back"
            elif "ภาคผนวก" in txt:
                curr_ch = "app"
            elif txt == "ประวัติผู้เขียน":
                curr_ch = "back"
                
            ch_ranges[curr_ch].append((txt, pg))

        for item in toc_titles:
            if item in toc_pages:
                continue
                
            target_ch = None
            if item.startswith("บทที่ 1") or item.startswith("1."):
                target_ch = 1
            elif item.startswith("บทที่ 2") or item.startswith("2."):
                target_ch = 2
            elif item.startswith("บทที่ 3") or item.startswith("3."):
                target_ch = 3
            elif item.startswith("บทที่ 4") or item.startswith("4."):
                target_ch = 4
            elif item.startswith("บทที่ 5") or item.startswith("5."):
                target_ch = 5
            elif item.startswith("ภาคผนวก") or item.startswith("ก."):
                target_ch = "app"
            else:
                target_ch = "back"
                
            phrase = searches[item]
            found_pg = None
            
            paras_to_search = ch_ranges.get(target_ch, body)
            for txt, pg in paras_to_search:
                if txt == phrase or txt.startswith(phrase):
                    # Avoid introductory bullet lists in Chapter 2/4 (which start with '2.X' or '4.X')
                    if not item.startswith("ก.") and any(txt.startswith(f"{c}.") for c in range(1, 10)):
                        continue
                    found_pg = str(pg)
                    break
                    
            if found_pg:
                toc_pages[item] = found_pg
            else:
                print(f"  Warning: TOC item not found in body: {item} (phrase: {phrase})")

        doc.Close(False)
        print(f"  Word COM Measurement complete: {len(toc_pages)} TOC, {len(lot_pages)} LOT, {len(lof_pages)} LOF.")
        return toc_pages, lot_pages, lof_pages
    finally:
        word.Quit()

def update_table_of_contents_in_place(doc, toc_page_map):
    print("Updating Table of Contents (TOC) in-place...")
    toc_start = None
    toc_end = None
    for idx, p in enumerate(doc.paragraphs[:450]):
        txt = p.text.strip()
        if txt == "สารบัญ" and toc_start is None:
            toc_start = idx
        if toc_start is not None and idx > toc_start:
            if txt == "สารบัญตาราง":
                toc_end = idx
                break
                
    if toc_start is None or toc_end is None:
        print(f"Error: Could not locate TOC range! start={toc_start}, end={toc_end}")
        return
        
    col_header_idx = toc_start + 1
    if "หน้า" not in doc.paragraphs[col_header_idx].text:
        col_header_idx = toc_start
        
    p_body_start = col_header_idx + 1
    updated_count = 0
    for i in range(p_body_start, toc_end):
        p = doc.paragraphs[i]
        full_text = p.text.strip()
        if not full_text:
            continue
        parts = p.text.split('\t')
        title = parts[0].strip()
        
        matched_page = None
        if title in toc_page_map:
            matched_page = toc_page_map[title]
        else:
            m_code = re.match(r'^((?:\d+|ก)\.\d+)', title)
            if m_code and m_code.group(1) in toc_page_map:
                matched_page = toc_page_map[m_code.group(1)]
            else:
                for k, v in toc_page_map.items():
                    if title == k or title.startswith(k) or k.startswith(title):
                        matched_page = v
                        break
                        
        if matched_page is not None:
            if len(p.runs) >= 2:
                p.runs[-1].text = f"\t{matched_page}"
            elif len(p.runs) == 1:
                p.runs[0].text = f"{title}\t{matched_page}"
            else:
                p.text = f"{title}\t{matched_page}"
                
            for r in p.runs:
                set_font_thai(r, "TH SarabunPSK", 16)
            updated_count += 1
            
    print(f"  Successfully updated {updated_count} TOC entries in-place.")

SPLIT_RULES_EXACT_LOT = {
    # ตัดบรรทัดเฉพาะตารางที่ยาวจนล้นบรรทัด 14.65 ซม. จริงเท่านั้น
    # ตารางที่ 4.4 และ 4.18 ไม่ล้น ให้แสดงบรรทัดเดียวเสมอ
    "ตารางที่ 4.19": ("ตารางที่ 4.19  ผลการประเมินความพึงพอใจด้านการออกแบบส่วนต่อประสานและความง่ายใน", "การใช้งาน"),
}

def rebuild_list_of_tables(doc, lot_pages):
    print("Rebuilding List of Tables (LOT) in front matter...")
    lot_start = None
    lot_end = None
    for idx, p in enumerate(doc.paragraphs[:450]):
        txt = p.text.strip()
        if txt == "สารบัญตาราง" and lot_start is None:
            lot_start = idx
        if lot_start is not None and idx > lot_start:
            if txt == "สารบัญรูป":
                lot_end = idx
                break
                
    if lot_start is None or lot_end is None:
        print(f"Error: Could not locate LOT range! start={lot_start}, end={lot_end}")
        return
        
    col_header_idx = lot_start + 1
    if "ตาราง" not in doc.paragraphs[col_header_idx].text:
        col_header_idx = lot_start
        
    p_body_start = col_header_idx + 1
    
    p_to_remove = []
    sect_break_p = None
    for i in range(p_body_start, lot_end):
        p = doc.paragraphs[i]
        has_sect = len([el for el in p._p.iter() if el.tag == qn('w:sectPr')]) > 0
        if has_sect:
            sect_break_p = p._p
            break
        p_to_remove.append(p._p)
        
    for p_elem in p_to_remove:
        p_elem.getparent().remove(p_elem)
        
    if sect_break_p is not None:
        ref_element = sect_break_p
        insert_mode = "before"
    else:
        ref_element = doc.paragraphs[col_header_idx]._p
        insert_mode = "after"
        
    lot_titles = [
        # บทที่ 3
        ("ตารางที่ 3.1", "ตารางที่ 3.1  ตารางในฐานข้อมูล MySQL"),
        ("ตารางที่ 3.2", "ตารางที่ 3.2  คอลเลกชันในฐานข้อมูล MongoDB"),
        ("ตารางที่ 3.3", "ตารางที่ 3.3  หมวดของ API"),
        ("ตารางที่ 3.4", "ตารางที่ 3.4  แอปของ Meta ที่ใช้ในระบบ"),
        ("ตารางที่ 3.5", "ตารางที่ 3.5  Valid OAuth Redirect URIs ของแอปเชื่อมต่อช่องทาง"),
        ("ตารางที่ 3.6", "ตารางที่ 3.6  สิทธิ์ที่ขอใน Configuration connect"),
        
        # บทที่ 4
        ("ตารางที่ 4.1", "ตารางที่ 4.1  ขอบเขตข้อมูลที่ใช้ในตารางกรณีทดสอบ"),
        ("ตารางที่ 4.2", "ตารางที่ 4.2  การทดสอบการเข้าสู่ระบบด้วยอีเมลและรหัสผ่าน"),
        ("ตารางที่ 4.3", "ตารางที่ 4.3  การทดสอบการเข้าสู่ระบบด้วยบัญชี Facebook และ Google"),
        ("ตารางที่ 4.4", "ตารางที่ 4.4  การทดสอบการตรวจลายเซ็นของ Webhook จาก LINE และ Meta"),
        ("ตารางที่ 4.5", "ตารางที่ 4.5  การทดสอบการส่งข้อความเกินกรอบเวลา 24 ชั่วโมง"),
        ("ตารางที่ 4.6", "ตารางที่ 4.6  การทดสอบการล็อกบทสนทนาเพื่อป้องกันการตอบซ้ำซ้อน"),
        ("ตารางที่ 4.7", "ตารางที่ 4.7  การทดสอบการแจ้งเตือนแชทที่ค้างตอบเกินกำหนด"),
        ("ตารางที่ 4.8", "ตารางที่ 4.8  การทดสอบการรวมข้อความที่ลูกค้าส่งติดกัน"),
        ("ตารางที่ 4.9", "ตารางที่ 4.9  การทดสอบการส่งต่อบทสนทนาเมื่อแชทบอทตอบไม่ได้"),
        ("ตารางที่ 4.10", "ตารางที่ 4.10  การทดสอบการป้องกันการตอบซ้ำเมื่องานถูกประมวลผลซ้ำ"),
        ("ตารางที่ 4.11", "ตารางที่ 4.11  การทดสอบการค้นคืนส่วนย่อยของเอกสารความรู้"),
        ("ตารางที่ 4.12", "ตารางที่ 4.12  การทดสอบการเปิดและแก้ไขคำสั่งซื้อจากบทสนทนา"),
        ("ตารางที่ 4.13", "ตารางที่ 4.13  การทดสอบการแยกข้อมูลระหว่างร้านค้า"),
        ("ตารางที่ 4.14", "ตารางที่ 4.14  การทดสอบการควบคุมสิทธิ์ตามบทบาท"),
        ("ตารางที่ 4.15", "ตารางที่ 4.15  การทดสอบการรับเหตุการณ์แบบเรียลไทม์บนแอปพลิเคชัน"),
        ("ตารางที่ 4.16", "ตารางที่ 4.16  การทดสอบการตอบกลับโดยอ้างอิงข้อความและการส่งไฟล์"),
        ("ตารางที่ 4.17", "ตารางที่ 4.17  การทดสอบการจัดการไฟล์สื่อขาเข้า"),
        ("ตารางที่ 4.18", "ตารางที่ 4.18  ผลการประเมินความพึงพอใจด้านฟังก์ชันการทำงานและความสามารถของระบบ"),
        ("ตารางที่ 4.19", "ตารางที่ 4.19  ผลการประเมินความพึงพอใจด้านการออกแบบส่วนต่อประสานและความง่ายในการใช้งาน"),
        ("ตารางที่ 4.20", "ตารางที่ 4.20  ผลการประเมินความพึงพอใจด้านประสิทธิภาพและความเสถียรของระบบ"),
        ("ตารางที่ 4.21", "ตารางที่ 4.21  ผลการประเมินความพึงพอใจด้านประโยชน์และความพึงพอใจในภาพรวม"),
        ("ตารางที่ 4.22", "ตารางที่ 4.22  สรุปผลการประเมินความพึงพอใจในภาพรวมทุกด้าน"),
        ("ตารางที่ 4.23", "ตารางที่ 4.23  การเปรียบเทียบการทำงานก่อนและหลังใช้แอปพลิเคชัน Chatshop"),
    ]
    
    for t_code, full_title in lot_titles:
        page_num = lot_pages.get(t_code, "0")
        matched_key = None
        for key in sorted(SPLIT_RULES_EXACT_LOT.keys(), key=len, reverse=True):
            if full_title.startswith(key):
                matched_key = key
                break
                
        p_xml = (
            f'<w:p {nsdecls("w")}>'
            f'<w:pPr>'
            f'<w:pStyle w:val="af9"/>'
            f'<w:tabs><w:tab w:val="right" w:pos="8294"/></w:tabs>'
            f'<w:spacing w:before="0" w:after="40" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
            f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
            f'</w:pPr>'
        )
        
        if matched_key:
            line1_text, line2_text = SPLIT_RULES_EXACT_LOT[matched_key]
            p_xml += (
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:t>{xml_escape(line1_text)}</w:t></w:r>'
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:tab/><w:t>{xml_escape(page_num)}</w:t><w:br/></w:r>'
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:t>{xml_escape(line2_text)}</w:t></w:r>'
            )
        else:
            p_xml += (
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:t>{xml_escape(full_title)}</w:t></w:r>'
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:tab/><w:t>{xml_escape(page_num)}</w:t></w:r>'
            )
            
        p_xml += '</w:p>'
        new_p_elem = parse_xml(p_xml)
        if insert_mode == "before":
            ref_element.addprevious(new_p_elem)
        else:
            ref_element.addnext(new_p_elem)
            ref_element = new_p_elem
            
    print(f"  Inserted {len(lot_titles)} clean LOT entries successfully.")

# สารบัญรูปทุกรายการมีความยาวไม่เกิน 14.65 ซม. ให้แสดงบรรทัดเดียวทั้งหมด ห้ามตัดบรรทัดเทียม
SPLIT_RULES_EXACT_LOF = {}

def rebuild_list_of_figures(doc, lof_pages):
    print("Rebuilding List of Figures (LOF) in front matter...")
    lof_start = None
    lof_end = None
    for idx, p in enumerate(doc.paragraphs[:450]):
        txt = p.text.strip()
        if txt == "สารบัญรูป" and lof_start is None:
            lof_start = idx
        if lof_start is not None and idx > lof_start:
            if txt == "บทนำ" or txt == "บทที่ 1" or txt.startswith("1.1"):
                lof_end = idx
                break
                
    if lof_start is None or lof_end is None:
        print(f"Error: Could not locate LOF range! start={lof_start}, end={lof_end}")
        return
        
    lof_titles = [
        # บทที่ 1
        ("รูปที่ 1.1", "รูปที่ 1.1  ที่ตั้งบริษัท ไฮเฟน พลัส จำกัด"),
        
        # บทที่ 2 (19 รูป)
        ("รูปที่ 2.1", "รูปที่ 2.1  โครงสร้างการทำงานของ Flutter"),
        ("รูปที่ 2.2", "รูปที่ 2.2  ความสัมพันธ์ระหว่าง Provider และวิดเจ็ตใน Riverpod สำหรับการจัดการสถานะ"),
        ("รูปที่ 2.3", "รูปที่ 2.3  โครงสร้างเส้นทางหน้าจอของแอปพลิเคชันด้วย go_router และ StatefulShellRoute"),
        ("รูปที่ 2.4", "รูปที่ 2.4  การทำงานตามรูปแบบ MVC ของ Laravel"),
        ("รูปที่ 2.5", "รูปที่ 2.5  การรับส่งข้อมูลผ่าน REST API"),
        ("รูปที่ 2.6", "รูปที่ 2.6  ขั้นตอนการขอสิทธิ์และการออกโทเคนด้วย OAuth 2.0 แบบ Authorization Code"),
        ("รูปที่ 2.7", "รูปที่ 2.7  การรับส่งข้อความผ่าน LINE Messaging API และการส่งต่อ Webhook"),
        ("รูปที่ 2.8", "รูปที่ 2.8  การรับส่งข้อความผ่าน Messenger Platform และ Instagram Messaging API"),
        ("รูปที่ 2.9", "รูปที่ 2.9  การทำงานของ Webhook เทียบกับการสอบถามข้อมูลเป็นระยะ (Polling)"),
        ("รูปที่ 2.10", "รูปที่ 2.10  การทำงานของคิวงาน"),
        ("รูปที่ 2.11", "รูปที่ 2.11  การสื่อสารแบบสองทางผ่าน WebSocket เทียบกับวงรอบของ HTTP"),
        ("รูปที่ 2.12", "รูปที่ 2.12  การส่ง Push Notification ผ่าน Firebase Cloud Messaging มายังแอปพลิเคชัน"),
        ("รูปที่ 2.13", "รูปที่ 2.13  ตัวอย่างตารางที่เชื่อมกันด้วยคีย์หลักและคีย์นอก"),
        ("รูปที่ 2.14", "รูปที่ 2.14  โครงสร้างฐานข้อมูล คอลเลกชัน และเอกสาร JSON ของ MongoDB"),
        ("รูปที่ 2.15", "รูปที่ 2.15  การแยกข้อมูลของผู้เช่าด้วยคอลัมน์รหัสผู้เช่า (Tenant Column Separation)"),
        ("รูปที่ 2.16", "รูปที่ 2.16  ความสัมพันธ์ระหว่างผู้ใช้ บทบาท และสิทธิ์ ในการเข้าถึงข้อมูล (RBAC)"),
        ("รูปที่ 2.17", "รูปที่ 2.17  สถาปัตยกรรม Transformer (ที่มา: [20])"),
        ("รูปที่ 2.18", "รูปที่ 2.18  ขั้นตอนการทำงานของ RAG"),
        ("รูปที่ 2.19", "รูปที่ 2.19  ตัวอย่างตำแหน่งของข้อความในปริภูมิเวกเตอร์และการคำนวณความคล้ายโคไซน์"),
        
        # บทที่ 3 (38 รูป)
        ("รูปที่ 3.1", "รูปที่ 3.1  Use Case Diagram ของระบบ Chatshop"),
        ("รูปที่ 3.2", "รูปที่ 3.2  สถาปัตยกรรมของระบบ Chatshop"),
        ("รูปที่ 3.3", "รูปที่ 3.3  ความสัมพันธ์ของข้อมูลใน MySQL และ MongoDB"),
        ("รูปที่ 3.4", "รูปที่ 3.4  ขั้นตอนการรับข้อความจาก Webhook"),
        ("รูปที่ 3.5", "รูปที่ 3.5  ขั้นตอนการสร้างคำตอบของแชทบอท"),
        ("รูปที่ 3.6", "รูปที่ 3.6  ขั้นตอนการรับเหตุการณ์และการแจ้งเตือนบนแอปพลิเคชันสมาร์ตโฟน"),
        ("รูปที่ 3.7", "รูปที่ 3.7  โครงสร้างโค้ดของแอปพลิเคชัน"),
        ("รูปที่ 3.8", "รูปที่ 3.8  หน้าเข้าสู่ระบบ"),
        ("รูปที่ 3.9", "รูปที่ 3.9  หน้าสมัครใช้งาน"),
        ("รูปที่ 3.10", "รูปที่ 3.10  หน้ากล่องข้อความ"),
        ("รูปที่ 3.11", "รูปที่ 3.11  หน้าค้นหาบทสนทนา"),
        ("รูปที่ 3.12", "รูปที่ 3.12  หน้าต่างสนทนา"),
        ("รูปที่ 3.13", "รูปที่ 3.13  เมนูเมื่อกดค้างที่ข้อความ"),
        ("รูปที่ 3.14", "รูปที่ 3.14  การเลือกโหมดแชทบอท"),
        ("รูปที่ 3.15", "รูปที่ 3.15  แผ่นวิเคราะห์บทสนทนา"),
        ("รูปที่ 3.16", "รูปที่ 3.16  แผ่นข้อมูลลูกค้า"),
        ("รูปที่ 3.17", "รูปที่ 3.17  แบบฟอร์มเปิดคำสั่งซื้อ"),
        ("รูปที่ 3.18", "รูปที่ 3.18  หน้าการแจ้งเตือน"),
        ("รูปที่ 3.19", "รูปที่ 3.19  Push Notification บนสมาร์ตโฟน"),
        ("รูปที่ 3.20", "รูปที่ 3.20  หน้าข้อมูลลูกค้า"),
        ("รูปที่ 3.21", "รูปที่ 3.21  หน้าสินค้า"),
        ("รูปที่ 3.22", "รูปที่ 3.22  หน้ารายการคำสั่งซื้อ"),
        ("รูปที่ 3.23", "รูปที่ 3.23  หน้าจัดการทีม"),
        ("รูปที่ 3.24", "รูปที่ 3.24  หน้ารายงาน"),
        ("รูปที่ 3.25", "รูปที่ 3.25  หน้าเชื่อมต่อช่องทาง"),
        ("รูปที่ 3.26", "รูปที่ 3.26  หน้าข้อมูลผู้ใช้"),
        ("รูปที่ 3.27", "รูปที่ 3.27  เมนูการตั้งค่า"),
        ("รูปที่ 3.28", "รูปที่ 3.28  หน้าตั้งค่าแชทบอท"),
        ("รูปที่ 3.29", "รูปที่ 3.29  Channel ID และ Channel Secret ใน LINE Developers Console"),
        ("รูปที่ 3.30", "รูปที่ 3.30  การเปิด Use webhook และปิดการตอบกลับอัตโนมัติของ LINE"),
        ("รูปที่ 3.31", "รูปที่ 3.31  การกรอก Channel ID และ Channel Secret ในแอปพลิเคชัน Chatshop"),
        ("รูปที่ 3.32", "รูปที่ 3.32  การตั้งค่า Facebook Login ของแอปเข้าสู่ระบบ"),
        ("รูปที่ 3.33", "รูปที่ 3.33  การเพิ่มแพลตฟอร์ม Android ในแอปเข้าสู่ระบบ"),
        ("รูปที่ 3.34", "รูปที่ 3.34  การตั้งค่า Valid OAuth Redirect URIs ของแอปเชื่อมต่อช่องทาง"),
        ("รูปที่ 3.35", "รูปที่ 3.35  การตั้งค่า Webhook ของ Meta"),
        ("รูปที่ 3.36", "รูปที่ 3.36  การเลือกเพจที่ต้องการเชื่อมต่อ"),
        ("รูปที่ 3.37", "รูปที่ 3.37  โปรเจกต์ Firebase และแอปพลิเคชัน Android"),
        ("รูปที่ 3.38", "รูปที่ 3.38  โค้ดสำหรับติด WebChat บนเว็บไซต์"),
        
        # ภาคผนวก ก (35 รูป)
        ("รูปที่ ก.1", "รูปที่ ก.1  หน้าเข้าสู่ระบบ"),
        ("รูปที่ ก.2", "รูปที่ ก.2  หน้าสมัครใช้งาน"),
        ("รูปที่ ก.3", "รูปที่ ก.3  หน้ากล่องข้อความ"),
        ("รูปที่ ก.4", "รูปที่ ก.4  หน้าค้นหาบทสนทนา"),
        ("รูปที่ ก.5", "รูปที่ ก.5  หน้าต่างสนทนา"),
        ("รูปที่ ก.6", "รูปที่ ก.6  เมนูเมื่อกดค้างที่ข้อความ"),
        ("รูปที่ ก.7", "รูปที่ ก.7  การเลือกโหมดแชทบอท"),
        ("รูปที่ ก.8", "รูปที่ ก.8  การตอบกลับข้อความอัตโนมัติของแชทบอทในโหมด AI Autopilot"),
        ("รูปที่ ก.9", "รูปที่ ก.9  แผ่นแนะนำข้อความคำตอบของแชทบอทในโหมด AI Suggestion"),
        ("รูปที่ ก.10", "รูปที่ ก.10  การส่งข้อความคำตอบที่แชทบอทแนะนำไปยังบทสนทนา"),
        ("รูปที่ ก.11", "รูปที่ ก.11  การส่งต่อบทสนทนาเมื่อแชทบอทไม่สามารถตอบได้"),
        ("รูปที่ ก.12", "รูปที่ ก.12  แผ่นวิเคราะห์บทสนทนา"),
        ("รูปที่ ก.13", "รูปที่ ก.13  แผ่นข้อมูลลูกค้า"),
        ("รูปที่ ก.14", "รูปที่ ก.14  แบบฟอร์มเปิดคำสั่งซื้อสินค้า"),
        ("รูปที่ ก.15", "รูปที่ ก.15  หน้าการแจ้งเตือน"),
        ("รูปที่ ก.16", "รูปที่ ก.16  การแจ้งเตือน Push Notification บนสมาร์ตโฟน"),
        ("รูปที่ ก.17", "รูปที่ ก.17  หน้ารายชื่อลูกค้า"),
        ("รูปที่ ก.18", "รูปที่ ก.18  หน้าแคตตาล็อกสินค้า"),
        ("รูปที่ ก.19", "รูปที่ ก.19  หน้ารายการคำสั่งซื้อ"),
        ("รูปที่ ก.20", "รูปที่ ก.20  หน้าจัดการทีมงาน"),
        ("รูปที่ ก.21", "รูปที่ ก.21  หน้ารายงานภาพรวม"),
        ("รูปที่ ก.22", "รูปที่ ก.22  หน้าเชื่อมต่อช่องทางการสนทนา"),
        ("รูปที่ ก.23", "รูปที่ ก.23  การกรอก Channel ID และ Channel Secret ในแอปพลิเคชัน"),
        ("รูปที่ ก.24", "รูปที่ ก.24  หน้าข้อมูลผู้ใช้"),
        ("รูปที่ ก.25", "รูปที่ ก.25  เมนูการตั้งค่าระบบ"),
        ("รูปที่ ก.26", "รูปที่ ก.26  หน้าตั้งค่าแชทบอท"),
        ("รูปที่ ก.27", "รูปที่ ก.27  หน้าตั้งค่าข้อมูลทั่วไปของร้านค้า"),
        ("รูปที่ ก.28", "รูปที่ ก.28  หน้าจัดการแท็กของร้านค้า"),
        ("รูปที่ ก.29", "รูปที่ ก.29  หน้าตั้งค่าระบบอัตโนมัติ"),
        ("รูปที่ ก.30", "รูปที่ ก.30  หน้าตั้งค่าแชทบอท"),
        ("รูปที่ ก.31", "รูปที่ ก.31  หน้าจัดการธุรกิจ"),
        ("รูปที่ ก.32", "รูปที่ ก.32  หน้าแก้ไขรายละเอียดธุรกิจ"),
        ("รูปที่ ก.33", "รูปที่ ก.33  หน้าตั้งค่าการแจ้งเตือน"),
        ("รูปที่ ก.34", "รูปที่ ก.34  หน้าตั้งค่าการเตือนแชทค้างตอบ"),
        ("รูปที่ ก.35", "รูปที่ ก.35  หน้าข้อมูลบัญชีผู้ใช้"),
    ]
    
    col_header_idx = lof_start + 1
    if "รูป" not in doc.paragraphs[col_header_idx].text:
        col_header_idx = lof_start
        
    p_body_start = col_header_idx + 1
    
    p_to_remove = []
    sect_break_p = None
    for i in range(p_body_start, lof_end):
        p = doc.paragraphs[i]
        has_sect = len([el for el in p._p.iter() if el.tag == qn('w:sectPr')]) > 0
        if has_sect:
            sect_break_p = p._p
            break
        p_to_remove.append(p._p)
        
    for p_elem in p_to_remove:
        p_elem.getparent().remove(p_elem)
        
    if sect_break_p is not None:
        ref_element = sect_break_p
        insert_mode = "before"
    else:
        ref_element = doc.paragraphs[col_header_idx]._p
        insert_mode = "after"
        
    for f_code, full_title in lof_titles:
        page_num = lof_pages.get(f_code, "0")
        matched_key = None
        for key in sorted(SPLIT_RULES_EXACT_LOF.keys(), key=len, reverse=True):
            if full_title.startswith(key):
                matched_key = key
                break
                
        p_xml = (
            f'<w:p {nsdecls("w")}>'
            f'<w:pPr>'
            f'<w:pStyle w:val="af9"/>'
            f'<w:tabs><w:tab w:val="right" w:pos="8294"/></w:tabs>'
            f'<w:spacing w:before="0" w:after="40" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
            f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
            f'</w:pPr>'
        )
        
        if matched_key:
            line1_text, line2_text = SPLIT_RULES_EXACT_LOF[matched_key]
            p_xml += (
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:t>{xml_escape(line1_text)}</w:t></w:r>'
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:tab/><w:t>{xml_escape(page_num)}</w:t><w:br/></w:r>'
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:t>{xml_escape(line2_text)}</w:t></w:r>'
            )
        else:
            p_xml += (
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:t>{xml_escape(full_title)}</w:t></w:r>'
                f'<w:r><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>'
                f'<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
                f'<w:tab/><w:t>{xml_escape(page_num)}</w:t></w:r>'
            )
            
        p_xml += '</w:p>'
        new_p_elem = parse_xml(p_xml)
        if insert_mode == "before":
            ref_element.addprevious(new_p_elem)
        else:
            ref_element.addnext(new_p_elem)
            ref_element = new_p_elem
            
    print(f"  Inserted {len(lof_titles)} clean LOF entries successfully.")

def export_docx_to_pdf(docx_path, pdf_path):
    print(f"Exporting PDF from {docx_path} to {pdf_path}...")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.Repaginate()
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(pdf_path),
            ExportFormat=17, # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,   # wdExportOptimizeForPrint
            CreateBookmarks=1
        )
        doc.Close(False)
    finally:
        word.Quit()
    print("  PDF export complete!")

def render_preview_pages(pdf_path, out_dir):
    print("Rendering preview PNGs for preliminary and Chapter 4 pages...")
    doc = pymupdf.open(pdf_path)
    pages_to_render = list(range(min(16, len(doc)))) + list(range(85, min(107, len(doc))))
    for p_idx in pages_to_render:
        page = doc[p_idx]
        pix = page.get_pixmap(dpi=150)
        out_png = os.path.join(out_dir, f"page_{p_idx+1}.png")
        pix.save(out_png)
    print(f"  Saved preview PNGs to {out_dir}")

def main():
    print("==================================================")
    print("=== Step 1: Loading v7 master and fixing styles ==")
    print("==================================================")
    doc = docx.Document(V7_PATH)
    fix_styles(doc)
    fix_abstract_indentation(doc)

    print("\n==================================================")
    print("=== Step 2: Merging Clean Chapter 2 ==============")
    print("==================================================")
    merge_chapter2_clean(doc)

    print("\n==================================================")
    print("=== Step 3: Removing Outlines from All Pictures ==")
    print("==================================================")
    remove_all_figure_outlines(doc)

    print("\n==================================================")
    print("=== Step 4: Fixing SEQ Fields & Captions =========")
    print("==================================================")
    fix_figure_captions_and_seq(doc)

    print("\n==================================================")
    print("=== Step 5: Building Appendix A Content ==========")
    print("==================================================")
    build_appendix_content(doc)
    remove_all_figure_outlines(doc)

    print("\n==================================================")
    print("=== Step 5.5: Reflowing Chapter 4 Layout =========")
    print("==================================================")
    prepare_chapter4_layout(doc)

    print("\n==================================================")
    print("=== Step 6: Saving temp docx for page measure ====")
    print("==================================================")
    doc.save(TEMP_MEASURE_DOCX)
    print(f"Saved temp measurement docx to {TEMP_MEASURE_DOCX}")

    print("\n==================================================")
    print("=== Step 7: Splitting Ch4 Tables & Measuring =====")
    print("==================================================")
    toc_pages, lot_pages, lof_pages = measure_actual_page_numbers(doc, TEMP_MEASURE_DOCX)
    if os.path.exists(TEMP_MEASURE_DOCX):
        try:
            os.remove(TEMP_MEASURE_DOCX)
        except Exception:
            pass

    print("\n==================================================")
    print("=== Step 8: Updating TOC In-Place ================")
    print("==================================================")
    update_table_of_contents_in_place(doc, toc_pages)

    print("\n==================================================")
    print("=== Step 9: Rebuilding List of Tables (LOT) ======")
    print("==================================================")
    rebuild_list_of_tables(doc, lot_pages)

    print("\n==================================================")
    print("=== Step 10: Rebuilding List of Figures (LOF) ====")
    print("==================================================")
    rebuild_list_of_figures(doc, lof_pages)

    print(f"\nSaving final v8 document to {V8_DOCX}...")
    doc.save(V8_DOCX)
    print("Saved v8.docx successfully!")

    print("\n==================================================")
    print("=== Step 11: Exporting Final PDF via Word COM ====")
    print("==================================================")
    export_docx_to_pdf(V8_DOCX, V8_PDF)
    print(f"Generated {V8_PDF} ({os.path.getsize(V8_PDF)} bytes) successfully!")

    print("\n==================================================")
    print("=== Step 12: Rendering Verification PNGs =========")
    print("==================================================")
    render_preview_pages(V8_PDF, VERIFY_DIR)
    print("=== ALL DONE! ===")

if __name__ == "__main__":
    main()
