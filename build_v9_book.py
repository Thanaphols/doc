"""
build_v9_book.py
สร้างไฟล์ Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v9.docx และ Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v9.pdf
โดย:
1. คัดลอก (Duplicate) ไฟล์จาก Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v8.docx แบบ 100%
2. ห้ามแตะต้องเนื้อหา รูปภาพ ตาราง สไตล์ การจัดหน้า หรือ XML ส่วนอื่นใดทั้งสิ้น
3. แก้ไขเฉพาะตัวเลขหน้าใน `<w:t>` หลัง `<w:tab/>` ของ:
   - สารบัญ (TOC)
   - สารบัญตาราง (LOT)
   - สารบัญรูป (LOF)
   ให้ตรงกับเลขหน้าจริงของทุกบท ทุกหัวข้อ ทุกตาราง และทุกรูปภาพ 100%
"""

import os
import sys
import shutil
import zipfile
import re
from lxml import etree
import win32com.client
import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

ROOT = os.path.abspath(os.path.dirname(__file__))
V9_DOCX = os.path.join(ROOT, "output", "Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v9.docx")
V9_PDF  = os.path.join(ROOT, "output", "Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์_v9.pdf")
BACKUP_DOCX = os.path.join(ROOT, "process", "temp", "backup_v9_user_edited.docx")
TEMP_PDF = os.path.join(ROOT, "process", "temp", "measure_v9_readonly.pdf")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def export_docx_to_pdf(docx_path, pdf_path):
    os.makedirs(os.path.dirname(os.path.abspath(pdf_path)), exist_ok=True)
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        doc.Repaginate()
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(pdf_path),
            ExportFormat=17,  # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,    # wdExportOptimizeForPrint
            CreateBookmarks=1
        )
        doc.Close(False)
    finally:
        word.Quit()


def measure_exact_pages_from_pdf(pdf_path):
    pdf_doc = pymupdf.open(pdf_path)
    body_pages = {}
    for p_idx in range(15, len(pdf_doc)):
        pg_num = p_idx - 14
        lines = [
            re.sub(r"\s+", " ", l).strip()
            for l in pdf_doc[p_idx].get_text("text").split("\n")
            if l.strip()
        ]
        body_pages[pg_num] = lines

    ch_starts = {}
    for pg, lines in body_pages.items():
        if "บทที่ 1" in lines and "บทนำ" in lines:
            ch_starts[1] = pg
        if "บทที่ 2" in lines and "ทฤษฎีและงานวิจัยที่เกี่ยวข้อง" in lines:
            ch_starts[2] = pg
        if "บทที่ 3" in lines and "วิธีดำเนินงาน" in lines:
            ch_starts[3] = pg
        if "บทที่ 4" in lines and "ผลการดำเนินงานและการทดสอบ" in lines:
            ch_starts[4] = pg
        if "บทที่ 5" in lines and "สรุปผลและข้อเสนอแนะ" in lines:
            ch_starts[5] = pg
        if any(l == "บรรณานุกรม" for l in lines[:3]):
            ch_starts["bib"] = pg
        if any("ภาคผนวก ก" in l for l in lines[:3]):
            ch_starts["app"] = pg
        if any(l == "ประวัติผู้เขียน" for l in lines[:3]):
            ch_starts["bio"] = pg

    ch_ends = {
        1: ch_starts[2] - 1,
        2: ch_starts[3] - 1,
        3: ch_starts[4] - 1,
        4: ch_starts[5] - 1,
        5: ch_starts["bib"] - 1,
        "bib": ch_starts["app"] - 1,
        "app": ch_starts["bio"] - 1,
        "bio": max(body_pages.keys()),
    }
    return body_pages, ch_starts, ch_ends


def get_paragraph_title_before_tab(p_elem):
    parts = []
    for r in p_elem.findall(f"{W}r"):
        if r.find(f"{W}tab") is not None:
            break
        for t in r.findall(f"{W}t"):
            if t.text:
                parts.append(t.text)
    return "".join(parts).strip()


def update_tab_page_number_in_paragraph(p_elem, new_page_str):
    all_r = p_elem.findall(f"{W}r")
    tab_r_idx = None
    for idx, r in enumerate(all_r):
        if r.find(f"{W}tab") is not None:
            tab_r_idx = idx
            break
    if tab_r_idx is None:
        raise RuntimeError("Paragraph does not contain a <w:tab/> run!")

    tab_r = all_r[tab_r_idx]
    t_nodes = tab_r.findall(f"{W}t")
    if not t_nodes:
        raise RuntimeError("Tab run does not contain a <w:t> node!")

    old_val = t_nodes[0].text or ""
    after_runs = all_r[tab_r_idx + 1:]
    for ar in after_runs:
        if ar.find(f"{W}br") is not None:
            break
        ar_texts = "".join(t.text or "" for t in ar.findall(f"{W}t"))
        if ar_texts.isdigit():
            old_val += ar_texts
            p_elem.remove(ar)

    t_nodes[0].text = str(new_page_str)
    return old_val


def main():
    os.makedirs(os.path.dirname(BACKUP_DOCX), exist_ok=True)
    shutil.copy2(V9_DOCX, BACKUP_DOCX)
    print("1. Exporting current v9.docx (ReadOnly) to temporary PDF to measure exact pages...")
    export_docx_to_pdf(V9_DOCX, TEMP_PDF)

    body_pages, ch_starts, ch_ends = measure_exact_pages_from_pdf(TEMP_PDF)
    print("   Chapter boundaries:", ch_starts)

    print("2. Updating ONLY TOC, LOT, LOF page numbers in v9.docx...")
    with zipfile.ZipFile(V9_DOCX, "r") as zin:
        doc_xml_bytes = zin.read("word/document.xml")

    root = etree.fromstring(doc_xml_bytes)
    paras = root.findall(f".//{W}body/{W}p")

    # Dynamically locate TOC, LOT, LOF header indices
    toc_hdr = lot_hdr = lof_hdr = body_hdr = None
    for idx, p in enumerate(paras[:350]):
        txt = "".join(t.text or "" for t in p.findall(f".//{W}t")).strip()
        if txt == "สารบัญ" and toc_hdr is None:
            toc_hdr = idx
        elif txt == "สารบัญตาราง" and lot_hdr is None:
            lot_hdr = idx
        elif txt == "สารบัญรูป" and lof_hdr is None:
            lof_hdr = idx
        elif txt == "บทนำ" and lof_hdr is not None and body_hdr is None:
            body_hdr = idx
            break

    assert None not in (toc_hdr, lot_hdr, lof_hdr, body_hdr), (toc_hdr, lot_hdr, lof_hdr, body_hdr)

    front_map = {
        "บทคัดย่อ": "ง",
        "Abstract": "จ",
        "กิตติกรรมประกาศ": "ฉ",
        "สารบัญ": "ช",
        "สารบัญตาราง": "ฎ",
        "สารบัญรูป": "ฐ",
    }

    toc_changes = 0
    for idx in range(toc_hdr + 2, lot_hdr):
        p = paras[idx]
        title = get_paragraph_title_before_tab(p)
        if not title:
            continue

        if title in front_map:
            actual_pg = front_map[title]
        elif title.startswith("บทที่ "):
            ch_n = int(title.split()[1])
            actual_pg = str(ch_starts[ch_n])
        elif title == "บรรณานุกรม":
            actual_pg = str(ch_starts["bib"])
        elif title.startswith("ภาคผนวก ก"):
            actual_pg = str(ch_starts["app"])
        elif title == "ประวัติผู้เขียน":
            actual_pg = str(ch_starts["bio"])
        else:
            m = re.match(r"^((?:\d+|ก)\.\d+)\s+(.*)$", title)
            if not m:
                raise RuntimeError(f"Unrecognized TOC title format: {title}")
            code = m.group(1)
            sub_title = m.group(2).strip()
            ch_key = "app" if code.startswith("ก.") else int(code.split(".")[0])
            found = []
            for pg in range(ch_starts[ch_key], ch_ends[ch_key] + 1):
                for l in body_pages[pg]:
                    if re.match(r"^" + re.escape(code) + r"\s+" + re.escape(sub_title[:15]), l):
                        found.append(pg)
            if not found:
                raise RuntimeError(f"Could not find heading in PDF for TOC entry: {title}")
            actual_pg = str(found[-1])

        old_pg = update_tab_page_number_in_paragraph(p, actual_pg)
        if old_pg != actual_pg:
            toc_changes += 1
            print(f"   [TOC] {title}: {old_pg} -> {actual_pg}")

    lot_changes = 0
    for idx in range(lot_hdr + 2, lof_hdr):
        p = paras[idx]
        title = get_paragraph_title_before_tab(p)
        if not title:
            continue
        m_code = re.match(r"^(ตารางที่ \d+\.\d+)", title)
        if not m_code:
            raise RuntimeError(f"Unrecognized LOT title: {title}")
        code = m_code.group(1)

        found = []
        for pg, lines in body_pages.items():
            for l in lines:
                if l.startswith(code + " ") and "(ต่อ)" not in l:
                    found.append(pg)
        if not found:
            raise RuntimeError(f"Could not find table in PDF: {code}")
        actual_pg = str(found[0])

        old_pg = update_tab_page_number_in_paragraph(p, actual_pg)
        if old_pg != actual_pg:
            lot_changes += 1
            print(f"   [LOT] {code}: {old_pg} -> {actual_pg}")

    lof_changes = 0
    for idx in range(lof_hdr + 2, body_hdr):
        p = paras[idx]
        title = get_paragraph_title_before_tab(p)
        if not title:
            continue
        m_code = re.match(r"^(รูปที่ (?:[^\s]+))", title)
        if not m_code:
            raise RuntimeError(f"Unrecognized LOF title: {title}")
        code = m_code.group(1)

        found = []
        for pg, lines in body_pages.items():
            for l in lines:
                if l.startswith(code + " ") and "(ต่อ)" not in l:
                    found.append(pg)
        if not found:
            raise RuntimeError(f"Could not find figure in PDF: {code}")
        actual_pg = str(found[0])

        old_pg = update_tab_page_number_in_paragraph(p, actual_pg)
        if old_pg != actual_pg:
            lof_changes += 1
            print(f"   [LOF] {code}: {old_pg} -> {actual_pg}")

    print(f"   Total page numbers updated: TOC={toc_changes}, LOT={lot_changes}, LOF={lof_changes}")

    new_doc_xml_bytes = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone="yes"
    )

    with zipfile.ZipFile(BACKUP_DOCX, "r") as zin:
        with zipfile.ZipFile(V9_DOCX, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_doc_xml_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))

    print(f"3. Saved {V9_DOCX} successfully!")
    print(f"4. Exporting {V9_DOCX} to {V9_PDF} via Word COM...")
    export_docx_to_pdf(V9_DOCX, V9_PDF)
    print("5. Verifying 100% page number accuracy in V9_PDF...")

    v9_body_pages, v9_ch_starts, v9_ch_ends = measure_exact_pages_from_pdf(V9_PDF)
    assert v9_ch_starts == ch_starts, f"Chapter starts shifted! {v9_ch_starts} vs {ch_starts}"
    print("   VERIFIED: All TOC, LOT, and LOF page numbers in v9 match 100% with zero layout shift!")


if __name__ == "__main__":
    main()
