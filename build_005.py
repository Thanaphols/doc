# -*- coding: utf-8 -*-
"""
Script to build 005_หนังสือรับรองการนำไปใช้ประโยชน์_ธนพล.doc and .pdf
From template output/005_หนังสือรับรองการนำไปใช้ประโยชน์_วชิรวิทย์.doc
"""
import os
import sys
import shutil
import win32com.client
import pymupdf

DOC_IN = os.path.abspath("output/005_หนังสือรับรองการนำไปใช้ประโยชน์_วชิรวิทย์.doc")
DOC_OUT = os.path.abspath("output/005_หนังสือรับรองการนำไปใช้ประโยชน์_ธนพล.doc")
PDF_OUT = os.path.abspath("output/005_หนังสือรับรองการนำไปใช้ประโยชน์_ธนพล.pdf")

def replace_all(doc, find_text, replace_text, max_iters=20):
    count = 0
    while count < max_iters:
        rng = doc.Content
        f = rng.Find
        f.ClearFormatting()
        if f.Execute(FindText=find_text, Forward=True):
            rng.Text = replace_text
            count += 1
        else:
            break
    print(f"  Replaced '{find_text}' -> '{replace_text}' ({count} times)")
    sys.stdout.flush()
    return count

def main():
    sys.stdout.reconfigure(encoding="utf-8")
    orig_mtime = os.path.getmtime(DOC_IN)
    print(f"[1] Original template mtime: {orig_mtime}")
    sys.stdout.flush()

    # Step 1: Copy template to new file
    shutil.copy2(DOC_IN, DOC_OUT)
    print(f"[2] Copied template to {DOC_OUT}")
    sys.stdout.flush()

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Open(DOC_OUT)
        print("[3] Opened doc in Word COM")
        sys.stdout.flush()

        # --- A. GLOBAL WORD REPLACEMENTS ---
        # 1. Spelling corrections & Name changes
        replace_all(doc, "อาจาร์ย", "อาจารย์")
        replace_all(doc, "วิศกรรม", "วิศวกรรม")
        replace_all(doc, "วชิรวิทย์ ชื่นจิตร", "ธนพล วุฒิธาดา")
        replace_all(doc, "ปัจุบัน", "ปัจจุบัน")
        
        # 2. Title replacements
        replace_all(doc, "ระบบรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์",
                         "แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์")
        replace_all(doc, "ระบบรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจ       ออนไลน์",
                         "แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์")

        # --- B. COMMERCIAL PARAGRAPH (PAGE 2) ---
        new_commercial_text = (
            "บริษัทนำแอปพลิเคชันบนสมาร์ตโฟนและระบบหลังบ้านที่นักศึกษาพัฒนาไปเป็นส่วนหนึ่งของ Chatshop "
            "ซึ่งเป็นผลิตภัณฑ์ที่บริษัทจะให้บริการแก่ร้านค้าออนไลน์ แอปพลิเคชันพัฒนาด้วย Flutter "
            "รองรับระบบปฏิบัติการ Android รวมแชทจาก LINE Official Account, Facebook Messenger, Instagram "
            "และ WebChat ไว้ในกล่องข้อความเดียว แสดงข้อความใหม่แบบเรียลไทม์และแจ้งเตือนผ่าน Push Notification "
            "เจ้าหน้าที่สามารถตอบแชท มอบหมายบทสนทนาให้ทีมงาน ปรับโหมดการทำงานของแชทบอท และเปิดคำสั่งซื้อจากหน้าแชทได้จากสมาร์ตโฟน "
            "ช่วยให้ร้านค้าตอบลูกค้าได้เร็วขึ้นแม้ไม่ได้อยู่หน้าคอมพิวเตอร์ "
            "ปัจจุบันระบบผ่านการทดสอบกับบัญชีทดสอบของบริษัทแล้ว และอยู่ระหว่างการตรวจโค้ดและการขออนุมัติแอปจาก Meta ก่อนเปิดให้ร้านค้าใช้งานจริง\r"
        )
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            if "บริษัทนำระบบเว็บที่นักศึกษาพัฒนา" in p.Range.Text:
                p_next = doc.Paragraphs(i + 1)
                if "แอปจาก Meta ก่อนเปิดให้ร้านค้าใช้งานจริง" in p_next.Range.Text:
                    p_next.Range.Delete()
                    print("  [P2] Deleted old trailing remnant paragraph")
                p.Range.Text = new_commercial_text
                p.Format.FirstLineIndent = 36.0
                p.Format.LeftIndent = 0.0
                p.Format.RightIndent = 0.0
                p.Format.SpaceBefore = 0.0
                p.Format.SpaceAfter = 0.0
                p.Format.LineSpacingRule = 5
                p.Format.LineSpacing = 13.8
                p.Format.Alignment = 9 # Thai Justify
                p.Range.Font.Name = "TH SarabunPSK"
                p.Range.Font.Size = 16.0
                p.Range.Font.Bold = False
                p.Range.Font.Italic = False
                print("  [P2] Updated Commercial paragraph")
                sys.stdout.flush()
                break

        # --- C. PAGE 3 (รูปประกอบ) ---
        # Scale shapes to 180 pt height
        for i in range(1, doc.InlineShapes.Count + 1):
            s = doc.InlineShapes(i)
            orig_w = s.Width
            orig_h = s.Height
            target_h = 180.0
            s.Height = target_h
            s.Width = orig_w * (target_h / orig_h)
            print(f"  [P3] Scaled Shape {i}: {s.Width:.1f} x {s.Height:.1f}")
            sys.stdout.flush()

        # Center all tables containing pictures
        for i in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(i)
            tbl.Rows.Alignment = 1 # Center
            for r in range(1, tbl.Rows.Count + 1):
                for c in range(1, tbl.Columns.Count + 1):
                    tbl.Cell(r, c).Range.ParagraphFormat.Alignment = 1

        # Replace Captions
        replace_all(doc, "รูปที่ 1 หน้าแรกของเว็บไซต์", "รูปที่ 1 ภาพหน้าจอเข้าสู่ระบบของแอปพลิเคชัน")
        replace_all(doc, "รูปที่ 2 ภาพหน้าข้อมูลผู้ใช้ของแอปพลิเคชัน", "รูปที่ 2 ภาพหน้าจอข้อมูลบัญชีผู้ใช้และพื้นที่ทำงานบนแอปพลิเคชัน")
        replace_all(doc, "รูปที่ 2 ภาพเมนูการตั้งค่าของแอปพลิเคชัน", "รูปที่ 3 ภาพหน้าจอเมนูการตั้งค่าของแอปพลิเคชัน")
        
        # Format captions: Centered, TH SarabunPSK 16pt
        for cap_txt in [
            "รูปที่ 1 ภาพหน้าจอเข้าสู่ระบบของแอปพลิเคชัน",
            "รูปที่ 2 ภาพหน้าจอข้อมูลบัญชีผู้ใช้และพื้นที่ทำงานบนแอปพลิเคชัน",
            "รูปที่ 3 ภาพหน้าจอเมนูการตั้งค่าของแอปพลิเคชัน"
        ]:
            rng = doc.Content
            if rng.Find.Execute(FindText=cap_txt):
                p_cap = rng.Paragraphs(1)
                p_cap.Range.Font.Name = "TH SarabunPSK"
                p_cap.Range.Font.Size = 16.0
                p_cap.Range.Font.Bold = False
                p_cap.Range.Font.Italic = False
                p_cap.Format.Alignment = 1 # Center
                p_cap.Format.FirstLineIndent = 0.0
                p_cap.Format.SpaceBefore = 0.0
                p_cap.Format.SpaceAfter = 0.0

        # Clean empty paragraphs between tables before objectives
        obj_idx = None
        for i in range(1, doc.Paragraphs.Count + 1):
            if "การนำไปใช้ประโยชน์ตามวัตถุประสงค์" in doc.Paragraphs(i).Range.Text:
                obj_idx = i
                break

        empty_paras = []
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            t = p.Range.Text.strip()
            if t == "" and p.Range.InlineShapes.Count == 0 and p.Range.Tables.Count == 0:
                empty_paras.append(i)

        for idx in reversed(empty_paras):
            if obj_idx and 34 < idx < obj_idx:
                try:
                    doc.Paragraphs(idx).Range.Delete()
                except Exception:
                    pass
        print("  [P3] Cleaned empty paragraphs between tables")
        sys.stdout.flush()

        # --- D. PAGE 4 (วัตถุประสงค์, ช่วงเวลา, ผลการเปลี่ยนแปลง) ---
        # Set PageBreakBefore on Objectives Header
        rng = doc.Content
        if rng.Find.Execute(FindText="การนำไปใช้ประโยชน์ตามวัตถุประสงค์"):
            p_obj_hdr = rng.Paragraphs(1)
            p_obj_hdr.Format.PageBreakBefore = True
            print("  [P4] Set PageBreakBefore on Objectives Header")
            sys.stdout.flush()

        # Objectives 1-5
        new_objs = [
            "1. เพื่อให้เจ้าหน้าที่รับและตอบข้อความจาก LINE Official Account, Facebook Messenger, Instagram และ WebChat ได้ในหน้าจอเดียวบนสมาร์ตโฟนแบบเรียลไทม์",
            "2. เพื่อแจ้งเตือนเจ้าหน้าที่ผ่าน Push Notification เมื่อมีข้อความใหม่หรือมีแชทที่ค้างตอบ",
            "3. เพื่อให้เจ้าหน้าที่จัดการบทสนทนา มอบหมายผู้รับผิดชอบ ค้นหาบทสนทนา และควบคุมโหมดการทำงานของแชทบอทได้จากแอปพลิเคชัน",
            "4. เพื่อให้เจ้าหน้าที่จัดการสินค้า ข้อมูลลูกค้า และเปิดคำสั่งซื้อจากหน้าแชทบนแอปพลิเคชันได้",
            "5. เพื่อเชื่อมต่อแอปพลิเคชันกับระบบหลังบ้านผ่าน API และระบบยืนยันตัวตนที่ปลอดภัย"
        ]
        for idx, obj_txt in enumerate(new_objs, 1):
            rng = doc.Content
            if rng.Find.Execute(FindText=f"{idx}. เพื่อให้"):
                p_obj = rng.Paragraphs(1)
                p_obj.Range.Text = obj_txt + "\r"
                p_obj.Range.Font.Name = "TH SarabunPSK"
                p_obj.Range.Font.Size = 16.0
                p_obj.Range.Font.Bold = False
                p_obj.Range.Font.Italic = False
                p_obj.Format.Alignment = 0
                p_obj.Format.FirstLineIndent = 0.0
                p_obj.Format.LeftIndent = 0.0
                p_obj.Format.RightIndent = 0.0
                p_obj.Format.SpaceBefore = 0.0
                p_obj.Format.SpaceAfter = 0.0
                print(f"  [P4] Updated Objective {idx}")
                sys.stdout.flush()

        # Results ๑-๓
        new_results = [
            "\t๑. บริษัทมีแอปพลิเคชันบนสมาร์ตโฟนที่รวมแชทจาก 4 ช่องทาง พร้อมการแจ้งเตือน แชทบอท และระบบคำสั่งซื้อ สำหรับเป็นส่วนหนึ่งของผลิตภัณฑ์ Chatshop\r",
            "\t๒. ลดภาระงานของทีมพัฒนาของบริษัทในส่วนแอปพลิเคชันบนสมาร์ตโฟนและระบบหลังบ้าน ระหว่างที่นักศึกษาปฏิบัติงานสหกิจศึกษา\r",
            "\t๓. บริษัทมีแอปพลิเคชันที่ผ่านการทดสอบการทำงานกับบัญชีทดสอบแล้ว และพร้อมเปิดให้ร้านค้าใช้งานหลังตรวจโค้ดและได้รับอนุมัติแอปจาก Meta\r"
        ]
        
        # Replace ๑.
        rng = doc.Content
        if rng.Find.Execute(FindText="๑. บริษัทมีระบบ"):
            p_res1 = rng.Paragraphs(1)
            p_res1.Range.Text = new_results[0]
            p_res1.Range.Font.Name = "TH SarabunPSK"
            p_res1.Range.Font.Size = 16.0
            p_res1.Range.Font.Bold = False
            p_res1.Range.Font.Italic = False
            p_res1.Format.Alignment = 9
            p_res1.Format.FirstLineIndent = 0.0
            p_res1.Format.LeftIndent = 0.0
            p_res1.Format.SpaceBefore = 0.0
            p_res1.Format.SpaceAfter = 0.0
            print("  [P4] Updated Result ๑.")
            sys.stdout.flush()

        # Replace ๒.
        rng = doc.Content
        if rng.Find.Execute(FindText="๒. ลดภาระงานของทีมพัฒนา"):
            p_res2 = rng.Paragraphs(1)
            p_res2.Range.Text = new_results[1]
            p_res2.Range.Font.Name = "TH SarabunPSK"
            p_res2.Range.Font.Size = 16.0
            p_res2.Range.Font.Bold = False
            p_res2.Range.Font.Italic = False
            p_res2.Format.Alignment = 9
            p_res2.Format.FirstLineIndent = 0.0
            p_res2.Format.LeftIndent = 0.0
            p_res2.Format.SpaceBefore = 0.0
            p_res2.Format.SpaceAfter = 0.0
            print("  [P4] Updated Result ๒.")
            sys.stdout.flush()

        # Replace ๓.
        rng = doc.Content
        if rng.Find.Execute(FindText="๓. บริษัทมีระบบที่ผ่านการทดสอบ"):
            p_res3 = rng.Paragraphs(1)
            p_res3.Range.Text = new_results[2]
            p_res3.Range.Font.Name = "TH SarabunPSK"
            p_res3.Range.Font.Size = 16.0
            p_res3.Range.Font.Bold = False
            p_res3.Range.Font.Italic = False
            p_res3.Format.Alignment = 9
            p_res3.Format.FirstLineIndent = 0.0
            p_res3.Format.LeftIndent = 0.0
            p_res3.Format.SpaceBefore = 0.0
            p_res3.Format.SpaceAfter = 0.0
            print("  [P4] Updated Result ๓.")
            sys.stdout.flush()

        # Save As .doc (Word 97 format = 0)
        doc.SaveAs(DOC_OUT, 0)
        print(f"[4] Saved .doc to {DOC_OUT}")
        sys.stdout.flush()

        # Export As .pdf (17)
        doc.ExportAsFixedFormat(PDF_OUT, 17)
        print(f"[5] Exported .pdf to {PDF_OUT}")
        sys.stdout.flush()

        doc.Close(False)
    finally:
        word.Quit()

    # Step 2: Check template mtime
    post_mtime = os.path.getmtime(DOC_IN)
    print(f"[6] Template mtime after: {post_mtime}")
    assert orig_mtime == post_mtime, "ERROR: Template was modified!"
    print("  Template file integrity verified (NOT modified)")

    # Step 3: Verify with PyMuPDF
    pdf = pymupdf.open(PDF_OUT)
    print(f"[7] PDF page count: {len(pdf)}")
    os.makedirs("process/temp/thanaphon_005", exist_ok=True)
    full_text = ""
    for i, page in enumerate(pdf):
        txt = page.get_text()
        full_text += txt + "\n"
        print(f"  Page {i+1}: {len(txt.strip())} chars")
        pix = page.get_pixmap(dpi=150)
        pix.save(f"process/temp/thanaphon_005/page_{i+1}.png")
    print("  Rendered all pages to PNG in process/temp/thanaphon_005/")

    # Step 4: Check forbidden words
    forbidden = ["วชิรวิทย์", "ชื่นจิตร", "ระบบเว็บ", "หน้าแรกของเว็บไซต์", "อาจาร์ย", "วิศกรรม", "ปัจุบัน"]
    print("[8] Checking forbidden words in output PDF text:")
    found_any = False
    for f in forbidden:
        count = full_text.count(f)
        if count > 0:
            print(f"  FAILED: Found {count} occurrences of forbidden word '{f}'!")
            found_any = True
        else:
            print(f"  PASSED: '{f}' NOT found.")
    if not found_any:
        print("  ALL FORBIDDEN WORDS CHECKS PASSED!")
    sys.stdout.flush()

if __name__ == "__main__":
    main()
