import os, sys, re
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.stdout.reconfigure(encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__)))
TEMPLATE = os.path.join(ROOT, "input", "reference", "rmuti_thesis_template.docx")
NUM_RE = re.compile(r"^(\d+(?:\.\d+)+)\s+")
LIST_RE = re.compile(r"^\d+\)\s")
COL_CM = {3: [2.2, 5.6, 6.85], 4: [3.6, 2.9, 4.7, 3.45]}

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

def add_caption(container, label, chapter, text, align):
    p = container.add_paragraph(style="Caption")
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    seq_no = container._seq[label] = container._seq.get(label, 0) + 1
    
    # Text run: "รูปที่ 2.X  ชื่อรูป" (directly write number, plus SEQ field for Word TOC)
    p.add_run(f"{label} {chapter}.{seq_no}  {text}")
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

    # 1 blank line before figure box
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
    
    # Paragraph 0: Image (centered)
    ph = cell.paragraphs[0]
    ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ph.paragraph_format.keep_with_next = True
    ph.paragraph_format.space_before = Pt(0)
    ph.paragraph_format.space_after = Pt(0)
    if img_path:
        run = ph.add_run()
        run.add_picture(img_path, width=Cm(12.0))
    else:
        ph.add_run(f"[ใส่รูป: {title}]")
    
    # Paragraph 1: 1 blank line (16pt) between image and caption
    p_mid = cell.add_paragraph()
    p_mid.paragraph_format.keep_with_next = True
    p_mid.paragraph_format.space_before = Pt(0)
    p_mid.paragraph_format.space_after = Pt(0)

    # Paragraph 2: Caption (centered)
    cell._seq = doc._seq
    add_caption(cell, "รูปที่", chapter, title, WD_ALIGN_PARAGRAPH.CENTER)

    # 1 blank line after figure box
    p_post = doc.add_paragraph()
    p_post.paragraph_format.space_before = Pt(0)
    p_post.paragraph_format.space_after = Pt(0)

def build_chapter_content(doc, lines, is_first_chapter=False):
    lines = [l.rstrip() for l in lines if l.strip()]
    chapter, title, body = int(lines[0]), lines[1].strip(), lines[2:]
    
    # Heading 1: "บทที่ X" \n title (Centered, bold, space before 36pt)
    h1 = doc.add_paragraph(style="Heading 1")
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Pt(36)
    h1.paragraph_format.space_after = Pt(16)
    # Style Heading 1 has automatic 'บทที่ X' list numbering, so run only needs '\n' + title
    r1 = h1.add_run(f"\n{title}")
    r1.bold = True

    topics = [NUM_RE.sub("", l) for l in body if (m := NUM_RE.match(l)) and m.group(1).count(".") == 1]
    
    # Process body paragraphs
    merged, just_merged = [], False
    for line in (l.strip() for l in body):
        prev = merged[-1] if merged else ""
        pm = NUM_RE.match(prev)
        if (chapter > 1 and not just_merged and pm and pm.group(1).count(".") >= 2
                and not NUM_RE.match(line) and not line.startswith("@")
                and " | " not in line and not LIST_RE.match(line)
                and not prev.endswith("ดังนี้")):
            merged[-1] = f"{prev} {line}"
            just_merged = True
        else:
            merged.append(line)
            just_merged = False

    first_h2 = True
    under_item = False
    i = 0
    while i < len(merged):
        line = merged[i]
        i += 1
        if line.startswith("@fig "):
            add_figure(doc, chapter, line[5:].strip())
            continue
        if LIST_RE.match(line):
            doc.add_paragraph("\t\t\t" + line)
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
            # Blank line before X.X heading
            doc.add_paragraph()
            first_h2 = False
            under_item = False
            
            p = doc.add_paragraph(style="Heading 2")
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            clean_title = NUM_RE.sub("", line).strip()
            r = p.add_run(clean_title)
            r.bold = True
        elif level >= 3:
            # Sub-heading (1.2.1, 2.1.1, etc.) -> Normal text, indented, typed number
            # level 3 = 1 tab, level 4 = 2 tabs
            tabs = "\t" * (level - 2)
            doc.add_paragraph(tabs + line)
            under_item = True
        else:
            # Body paragraph
            tabs = "\t\t" if under_item else "\t"
            doc.add_paragraph(tabs + line)

def main():
    doc = docx.Document(TEMPLATE)
    # Keep only section settings
    for el in list(doc.element.body):
        if el.tag != qn("w:sectPr"):
            doc.element.body.remove(el)

    s1 = doc.sections[0]
    s1.top_margin = Cm(3.81)
    doc._seq = {}

    print("Building Chapter 1...")
    with open(os.path.join(ROOT, "input", "chapters", "ch1.txt"), encoding="utf-8") as f:
        build_chapter_content(doc, f.readlines(), is_first_chapter=True)

    print("Adding Section Break for Chapter 2...")
    s2 = doc.add_section(WD_SECTION.NEW_PAGE)
    s2.top_margin = Cm(3.81)
    # reset chapter sequence for figures
    doc._seq["รูปที่"] = 0

    print("Building Chapter 2...")
    with open(os.path.join(ROOT, "input", "chapters", "ch2.txt"), encoding="utf-8") as f:
        build_chapter_content(doc, f.readlines(), is_first_chapter=False)

    print("Tagging Thai runs...")
    for p in doc.paragraphs:
        for r in p.runs:
            mark_thai(r)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        mark_thai(r)

    out_docx = os.path.join(ROOT, "output", "Chatshop_บทที่1_2.docx")
    out_pdf = os.path.join(ROOT, "output", "Chatshop_บทที่1_2.pdf")
    doc.save(out_docx)
    print(f"Saved: {out_docx}")

    # Export to PDF via Word COM
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(os.path.abspath(out_docx))
        d.Fields.Update()
        d.Repaginate()
        d.Save()
        d.ExportAsFixedFormat(os.path.abspath(out_pdf), 17)
        print(f"Exported PDF: {out_pdf} (Total pages: {d.ComputeStatistics(2)})")
        d.Close(False)
    finally:
        word.Quit()

    # Render PNGs
    import pymupdf
    out_dir = os.path.join(ROOT, "process", "temp")
    os.makedirs(out_dir, exist_ok=True)
    pdf = pymupdf.open(out_pdf)
    for i, page in enumerate(pdf):
        page.get_pixmap(dpi=130).save(os.path.join(out_dir, f"ch1_2_p{i + 1}.png"))
    print(f"Rendered {len(pdf)} page images to {out_dir}")

if __name__ == "__main__":
    main()
