import os, sys, shutil
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

ROOT = os.path.abspath(os.path.dirname(__file__))
SRC_1_2 = os.path.join(ROOT, "output", "Chatshop_บทที่1_2.docx")
SRC_3 = os.path.join(ROOT, "output", "Chatshop_บทที่3_วิธีดำเนินงาน.docx")
OUT_DOCX = os.path.join(ROOT, "output", "Chatshop_บทที่1_3.docx")
OUT_PDF = os.path.join(ROOT, "output", "Chatshop_บทที่1_3.pdf")

def main():
    print("Duplicating Chapter 1-2 as baseline for Chapter 1-3...")
    shutil.copy2(SRC_1_2, OUT_DOCX)
    
    print("Merging Chapter 3 via Word COM...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(OUT_DOCX))
        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd
        rng.InsertBreak(2)  # wdSectionBreakNextPage
        rng.Collapse(0)
        rng.InsertFile(os.path.abspath(SRC_3))
        
        print("Updating fields and repaginating...")
        doc.Fields.Update()
        doc.Repaginate()
        doc.Save()
        
        print(f"Exporting PDF: {OUT_PDF}...")
        doc.ExportAsFixedFormat(os.path.abspath(OUT_PDF), 17)
        total_pages = doc.ComputeStatistics(2)
        print(f"Successfully generated Chatshop_บทที่1_3.docx and PDF (Total pages: {total_pages})")
        doc.Close(False)
    finally:
        word.Quit()

if __name__ == "__main__":
    main()
