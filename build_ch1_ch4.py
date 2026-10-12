import os, sys, shutil
import win32com.client
import pymupdf

sys.stdout.reconfigure(encoding='utf-8')

ROOT = os.path.abspath(os.path.dirname(__file__))
SRC_1_3 = os.path.join(ROOT, "output", "Chatshop_บทที่1_3.docx")
SRC_4 = os.path.join(ROOT, "output", "Chatshop_บทที่4_ผลการดำเนินงาน.docx")
OUT_DOCX = os.path.join(ROOT, "output", "Chatshop_บทที่1_4.docx")
OUT_PDF = os.path.join(ROOT, "output", "Chatshop_บทที่1_4.pdf")

def main():
    print(f"Duplicating Chapter 1-3 baseline: {SRC_1_3} -> {OUT_DOCX}...")
    shutil.copy2(SRC_1_3, OUT_DOCX)
    
    print("Merging Chapter 4 via Word COM...")
    word = win32com.client.DispatchEx("Word.Application")  # own instance, never the user's open Word
    word.Visible = False
    word.DisplayAlerts = 0  # no hidden modal dialogs
    try:
        doc = word.Documents.Open(os.path.abspath(OUT_DOCX))
        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd = 0
        rng.InsertBreak(2)  # wdSectionBreakNextPage = 2
        rng.Collapse(0)
        rng.InsertFile(os.path.abspath(SRC_4))
        
        print("Updating fields and repaginating...")
        doc.Fields.Update()
        doc.Repaginate()
        doc.Save()
        
        print(f"Exporting PDF: {OUT_PDF}...")
        doc.ExportAsFixedFormat(os.path.abspath(OUT_PDF), 17)
        total_pages = doc.ComputeStatistics(2)
        print(f"Successfully generated Chatshop_บทที่1_4.docx and PDF (Total pages: {total_pages})")
        doc.Close(False)
    finally:
        word.Quit()

    # Render inspection pages to temp
    print("Rendering inspection pages...")
    out_dir = os.path.join(ROOT, "process", "temp")
    os.makedirs(out_dir, exist_ok=True)
    pdf = pymupdf.open(OUT_PDF)
    print(f"PDF page count: {len(pdf)}")
    # Chapter 3 ended around page 63, so let's render pages 62, 63, 64 (start of Ch4), 65, 66, and the end
    inspection_pages = [1, 28, 63, 64, 65, len(pdf)]
    for pno in inspection_pages:
        if 1 <= pno <= len(pdf):
            pdf[pno - 1].get_pixmap(dpi=130).save(os.path.join(out_dir, f"ch1_4_p{pno}.png"))
    print("Inspection pages rendered successfully.")

if __name__ == "__main__":
    main()
