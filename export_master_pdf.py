import win32com.client, os, sys
sys.stdout.reconfigure(encoding='utf-8')

word = win32com.client.DispatchEx('Word.Application')  # own instance, never the user's open Word
word.Visible = False
word.DisplayAlerts = 0  # no hidden modal dialogs
try:
    docx_path = os.path.abspath('output/Chatshop_ฉบับรวม5บท.docx')
    pdf_path = os.path.abspath('output/Chatshop_ฉบับรวม5บท.pdf')
    d = word.Documents.Open(docx_path)
    d.Fields.Update()
    d.Repaginate()
    d.Save()
    d.SaveAs(pdf_path, FileFormat=17) # 17 = wdFormatPDF
    print(f'Exported {pdf_path} successfully. Total pages: {d.ComputeStatistics(2)}')
    d.Close(False)
finally:
    word.Quit()
