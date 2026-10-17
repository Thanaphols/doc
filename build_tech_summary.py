import os, sys
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(ROOT, "output")
os.makedirs(OUT_DIR, exist_ok=True)

DOCX_OUT = os.path.join(OUT_DIR, "Chatshop_สรุปเทคโนโลยีที่ใช้ในโครงงาน.docx")
PDF_OUT = os.path.join(OUT_DIR, "Chatshop_สรุปเทคโนโลยีที่ใช้ในโครงงาน.pdf")

def mark_thai(run, font_size=15, bold=False, italic=False, color_rgb=None):
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
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

def add_p(doc, text="", align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, size=15, space_b=0, space_a=3, line_spacing=1.1, color_rgb=None):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_b)
    pf.space_after = Pt(space_a)
    pf.line_spacing = line_spacing
    if text:
        if "**" in text:
            parts = text.split("**")
            for idx, part in enumerate(parts):
                if not part: continue
                is_b = bold or (idx % 2 == 1)
                r = p.add_run(part)
                mark_thai(r, font_size=size, bold=is_b, italic=italic, color_rgb=color_rgb)
        else:
            r = p.add_run(text)
            mark_thai(r, font_size=size, bold=bold, italic=italic, color_rgb=color_rgb)
    return p

def add_heading_1(doc, text):
    p = add_p(doc, text, bold=True, size=17, space_b=10, space_a=4, color_rgb=RGBColor(24, 76, 120))
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="184C78"/></w:pBdr>')
    pPr.append(pBdr)
    return p

def add_heading_2(doc, text):
    return add_p(doc, text, bold=True, size=15.5, space_b=6, space_a=2, color_rgb=RGBColor(40, 110, 160))

def style_table(tbl, col_widths, headers, data, header_bg="184C78"):
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = tbl._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="B0C4DE"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="184C78"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E6ED"/>'
        '<w:left w:val="none"/>'
        '<w:right w:val="none"/>'
        '<w:insideV w:val="none"/>'
        '</w:tblBorders>'
    )
    tblPr.append(borders)
    
    # Header row
    hdr_row = tbl.rows[0]
    trPr = hdr_row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    
    for i, title in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.width = Cm(col_widths[i])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{header_bg}"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(title)
        mark_thai(r, font_size=14.5, bold=True, color_rgb=RGBColor(255, 255, 255))
        
    for r_idx, row_data in enumerate(data):
        row = tbl.rows[r_idx + 1]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        bg_color = "F7F9FB" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = Cm(col_widths[c_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            if bg_color != "FFFFFF":
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{bg_color}"/>')
                cell._tc.get_or_add_tcPr().append(shd)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            
            # Formatted text support (lines separated by \n)
            lines = val.split("\n")
            for l_idx, line in enumerate(lines):
                if l_idx > 0:
                    p = cell.add_paragraph()
                    p.paragraph_format.space_before = Pt(1)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.15
                
                # Check for bold tags like **text**
                parts = line.split("**")
                for p_idx, part in enumerate(parts):
                    if not part: continue
                    is_b = (p_idx % 2 == 1)
                    r = p.add_run(part)
                    mark_thai(r, font_size=14, bold=is_b, color_rgb=RGBColor(30, 30, 30))

def build_tech_document():
    doc = docx.Document()
    
    # Page setup: A4 with 2.0 cm margins (clean document layout)
    s = doc.sections[0]
    s.page_width = Cm(21.0)
    s.page_height = Cm(29.7)
    s.top_margin = Cm(2.0)
    s.bottom_margin = Cm(2.0)
    s.left_margin = Cm(2.2)
    s.right_margin = Cm(2.2)
    
    # Header & Footer
    footer = s.footer
    f_p = footer.paragraphs[0]
    f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    f_run = f_p.add_run("เอกสารสรุปเทคโนโลยีประกอบการนำเสนอโครงงาน Chatshop | สาขาวิชาวิศวกรรมคอมพิวเตอร์ มทร.อีสาน ขอนแก่น")
    mark_thai(f_run, font_size=12, italic=True, color_rgb=RGBColor(128, 128, 128))
    
    # Title Banner
    p_title = add_p(doc, "สรุปภาพรวมเทคโนโลยีและสถาปัตยกรรมระบบ (Technology Stack)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=21, space_b=0, space_a=2, color_rgb=RGBColor(24, 76, 120))
    p_sub = add_p(doc, "แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ (Chatshop)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_a=4, color_rgb=RGBColor(70, 70, 70))
    p_auth = add_p(doc, "ผู้จัดทำ: นายธนพล วุฒิธาดา  •  อาจารย์ที่ปรึกษา: ผศ. ดร.อติราช สุขสวัสดิ์  •  สถานประกอบการ: บจก. ไฮเฟนพลัส", align=WD_ALIGN_PARAGRAPH.CENTER, size=13.5, space_a=14, color_rgb=RGBColor(100, 100, 100))
    
    # Intro Callout Box
    tbl_intro = doc.add_table(rows=1, cols=1)
    tbl_intro.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_intro = tbl_intro.rows[0].cells[0]
    c_intro.width = Cm(16.6)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="EEF4F9"/>')
    c_intro._tc.get_or_add_tcPr().append(shd)
    bdr = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="184C78"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
    c_intro._tc.get_or_add_tcPr().append(bdr)
    p_in = c_intro.paragraphs[0]
    p_in.paragraph_format.space_before = Pt(4)
    p_in.paragraph_format.space_after = Pt(4)
    r_in = p_in.add_run("💡 วัตถุประสงค์เอกสาร: ใช้สำหรับเตรียมตัวนำเสนอโครงงานและเตรียมตอบข้อซักถามของคณะกรรมการสอบปริญญานิพนธ์ โดยสรุปครอบคลุมทุกหมวดหมู่ของเทคโนโลยี เหตุผลทางวิศวกรรมในการคัดเลือกเครื่องมือ และแนวทางตอบคำถามยอดฮิต")
    mark_thai(r_in, font_size=14, italic=True, color_rgb=RGBColor(20, 50, 80))

    add_p(doc, "", size=6, space_a=4)

    # -------------------------------------------------------------
    # Section 1: Mobile App (Flutter)
    # -------------------------------------------------------------
    add_heading_1(doc, "1. ฝั่งแอปพลิเคชันบนสมาร์ตโฟน (Frontend / Mobile App)")
    add_p(doc, "• Framework: **Flutter (Dart SDK)** พัฒนาแบบ Cross-Platform (มุ่งเน้นการทดสอบและใช้งานบน Android)")
    add_p(doc, "• Architectural Pattern: **Feature-First / Layered Architecture** แบ่งโครงสร้างโค้ดเป็น 4 เลเยอร์หลักตามหน้าที่ ได้แก่ **core** (บริการส่วนกลาง), **data** (Data Models & Repositories), **features** (13 โมดูลฟังก์ชัน), และ **shared** (Reusable Widgets & Bottom Sheets)")
    
    flutter_headers = ["หมวดการทำงาน", "ไลบรารี / เครื่องมือ", "หน้าที่และเหตุผลการเลือกใช้ (Key Defense Arguments)"]
    flutter_data = [
        [
            "**State Management**",
            "**Riverpod**\n(flutter_riverpod)",
            "• จัดการ State ทั้งแอปพลิเคชันแบบ Compile-time Safe (ตรวจจับข้อผิดพลาดตั้งแต่ตอนเขียนโค้ด ไม่เกิด Runtime Error เรื่อง Provider Not Found)\n• แยก Business Logic ออกจาก UI อย่างเด็ดขาด\n• **จุดเด่นสำคัญ:** เมื่อผู้ใช้ออกจากระบบ (Logout) สามารถสั่ง Invalidate State ต้นทางเพื่อเคลียร์แคช, ยกเลิกการเชื่อมต่อ WebSocket และล้างโทเคนของผู้ใช้เดิมทิ้งได้ทันทีในคำสั่งเดียว"
        ],
        [
            "**Navigation & Routing**",
            "**go_router**",
            "• จัดการเส้นทางหน้าจอแบบ Declarative URL\n• **รองรับ Deep Linking:** เมื่อผู้ใช้แตะ Push Notification สามารถสั่งให้แอปเปิดตรงไปยังหน้าห้องแชท หรือหน้ารายการออเดอร์ของลูกค้ารายนั้นได้ทันที\n• **Auto-Redirect:** มีระบบตรวจสอบ Auth State หากระบบหลังบ้านตอบกลับรหัส 401 Unauthorized จะลบโทเคนและดีดกลับไปหน้า Login ทันที"
        ],
        [
            "**Secure Storage**",
            "**flutter_secure_storage**",
            "• จัดเก็บข้อมูลความลับ เช่น Bearer Token (JWT) และข้อมูลเซสชัน\n• **ความปลอดภัยสูง:** เข้ารหัสผ่านฮาร์ดแวร์โดยตรง (**Android KeyStore** และ **iOS Keychain**) ไม่ใช้ SharedPreferences ธรรมดา ป้องกันการถูกโจรกรรมข้อมูลเมื่ออุปกรณ์ถูก Root"
        ],
        [
            "**Network & API Client**",
            "**Custom HTTP Client**\n(REST API)",
            "• เรียกใช้งาน REST API กว่า 99 เส้นทางของระบบหลังบ้าน\n• มี Interceptor คอยแนบ Bearer Token ให้อัตโนมัติทุก Request\n• รองรับการส่งข้อมูลแบบ **Multipart Request** สำหรับการอัปโหลดไฟล์ รูปภาพ วิดีโอ และสลิปการโอนเงิน"
        ],
        [
            "**Real-time Client**",
            "**WebSocket Client**\n(Pusher Protocol)",
            "• เชื่อมต่อไปยัง Laravel Reverb ผ่านโปรโตคอล Pusher\n• รับเหตุการณ์แบบเรียลไทม์: ข้อความใหม่, อัปเดตสถานะการอ่าน, การล็อกบทสนทนา, ออเดอร์เปลี่ยนสถานะ และ Typing Indicator (กำลังพิมพ์...)"
        ],
        [
            "**Push Notification**",
            "**firebase_messaging**\n**flutter_local_notifications**",
            "• เชื่อมต่อ Firebase Cloud Messaging (FCM)\n• รับการแจ้งเตือนข้อความใหม่และแชทค้างตอบได้ทันทีแม้ปิดแอปพลิเคชันอยู่\n• แสดงผลเป็น Heads-up Banner พร้อมปุ่มเปิดห้องแชทเมื่อผู้ใช้กำลังเปิดแอปอยู่"
        ],
        [
            "**Social Authentication**",
            "**google_sign_in**\n**flutter_facebook_auth**",
            "• จัดการการเข้าสู่ระบบด้วยบัญชี Google และ Facebook จากหน้า Native ของสมาร์ตโฟน แล้วส่ง Token ไปตรวจสอบความถูกต้องกับระบบหลังบ้าน"
        ]
    ]
    t1 = doc.add_table(rows=len(flutter_data)+1, cols=3)
    style_table(t1, [3.2, 3.8, 9.6], flutter_headers, flutter_data)

    add_p(doc, "", size=6, space_a=6)

    # -------------------------------------------------------------
    # Section 2: Backend (Laravel)
    # -------------------------------------------------------------
    add_heading_1(doc, "2. ฝั่งระบบบริการหลังบ้าน (Backend Server Engine)")
    add_p(doc, "• Core Framework: **Laravel 10** บนสภาพแวดล้อม **PHP 8.2+ (PHP-FPM)**")
    add_p(doc, "• Web Server & Security: รันภายใต้ **Nginx Reverse Proxy** พร้อมใบรับรองความปลอดภัย **SSL/TLS (HTTPS)** ผ่าน Let's Encrypt ซึ่งจำเป็นอย่างยิ่งเพราะ LINE และ Meta บังคับให้ Webhook URL ต้องเป็น HTTPS เท่านั้น")

    backend_headers = ["องค์ประกอบ", "เทคโนโลยีที่ใช้", "รายละเอียดการทำงานเชิงเทคนิค"]
    backend_data = [
        [
            "**API & Authentication**",
            "**Laravel Sanctum**\n+ HMAC-SHA256",
            "• ออก Bearer API Token (Personal Access Token) ให้กับแอปพลิเคชันมือถือ\n• ตรวจสอบความถูกต้องของ Social Token จาก Google/Facebook ก่อนออก Token ของระบบ\n• **Webhook Security:** ตรวจสอบลายเซ็นดิจิทัลของแพลตฟอร์ม LINE และ Meta ด้วยอัลกอริทึม **HMAC-SHA256** ป้องกันการปลอมแปลง Request"
        ],
        [
            "**Real-time Engine**",
            "**Laravel Reverb**\n(WebSocket Server)",
            "• เซิร์ฟเวอร์ WebSocket ประสิทธิภาพสูงของ Laravel รันบน Event Loop\n• Broadcast เหตุการณ์สำคัญ เช่น `NewMessage`, `OrderUpdated`, `SettingsChanged` สู่แอปพลิเคชันระดับมิลลิวินาที โดยไม่ต้องเสียค่าบริการ Third-party Cloud"
        ],
        [
            "**Queue & Workers**",
            "**Laravel Queue**\n(Supervisor Control)",
            "• สถาปัตยกรรม Asynchronous Queue แบ่งเป็น 3 คิวตามความสำคัญ:\n  1. `high`: ประมวลผล Webhook ขาเข้าแบบด่วน\n  2. `default`: งานทั่วไป เช่น อัปเดตข้อมูลลูกค้าและออเดอร์\n  3. `typing`: จำลองสถานะการพิมพ์ของแชทบอท\n• ควบคุมกระบวนการทำงานด้วย **Supervisor** เพื่อสั่ง Auto-Restart ทันทีหากโปรเซสหยุดทำงาน"
        ],
        [
            "**Task Scheduler**",
            "**Laravel Scheduler**\n(Linux Cron 1 min)",
            "• รันเบื้องหลังทุก 1 นาที ตรวจหาบทสนทนาที่ค้างตอบเกิน SLA ที่ร้านตั้งไว้ (5, 10, 15, 20 นาที) และส่ง Push Notification ไปยังเจ้าหน้าที่ผู้รับผิดชอบ"
        ],
        [
            "**Media Processing**",
            "**PHP GD** & **FFmpeg**",
            "• ปรับขนาดและบีบอัดรูปภาพด้วย PHP GD Library\n• แปลงและลดขนาดไฟล์วิดีโอด้วย FFmpeg ก่อนจัดเก็บลงดิสก์ ช่วยลดปริมาณพื้นที่และการใช้แบนด์วิดท์"
        ]
    ]
    t2 = doc.add_table(rows=len(backend_data)+1, cols=3)
    style_table(t2, [3.2, 3.8, 9.6], backend_headers, backend_data)

    add_p(doc, "", size=6, space_a=6)

    # -------------------------------------------------------------
    # Section 3: Dual Database Architecture
    # -------------------------------------------------------------
    add_heading_1(doc, "3. สถาปัตยกรรมฐานข้อมูลแบบผสมผสาน (Hybrid / Dual Database Architecture)")
    add_p(doc, "ระบบออกแบบโครงสร้างจัดเก็บข้อมูลแบบ **Polyglot Persistence** โดยเลือกใช้ฐานข้อมูล 2 ประเภทควบคู่กันตามธรรมชาติของข้อมูล:")

    db_headers = ["ฐานข้อมูล", "ประเภท", "ข้อมูลที่จัดเก็บ", "เหตุผลทางวิศวกรรม (Engineering Rationale)"]
    db_data = [
        [
            "**MySQL**\n(InnoDB)",
            "Relational Database\n(RDBMS)",
            "• ร้านค้า (shops)\n• ผู้ใช้งาน (users)\n• สิทธิ์และบทบาท (roles/RBAC)\n• สมาชิกทีม (team)\n• แคตตาล็อกสินค้า (products)\n• รายการคำสั่งซื้อ (orders)",
            "• **ต้องการความถูกต้องแม่นยำสูง (ACID Compliance):** มีระบบ Transaction และ Foreign Keys แบบลบตามกัน (Cascade Delete) ป้องกันข้อมูลค้างเมื่อลบร้านค้า\n• เหมาะกับข้อมูลเชิงโครงสร้างที่มีความสัมพันธ์แน่นอน เช่น ข้อมูลการเงินและสต็อกสินค้า\n• ใช้ชุดอักขระ `utf8mb4` รองรับภาษาไทยและ Emoji สมบูรณ์"
        ],
        [
            "**MongoDB**\n(Document)",
            "NoSQL\n(Document Store)",
            "• ข้อความแชท (`messages`)\n• บทสนทนา (`conversations`)\n• เอกสารความรู้ (`knowledge_documents`)\n• เวกเตอร์ Embeddings (`knowledge_chunks`)\n• โทเคนอุปกรณ์ (`device_tokens`)",
            "• **โครงสร้างยืดหยุ่นสูง (Flexible Schema):** ข้อความแชทจาก 4 แพลตฟอร์มมี Payload ต่างกันมาก (ข้อความ, รูปภาพ, สติกเกอร์, พิกัด, ปุ่มกด)\n• **รองรับปริมาณมหาศาล (High Write Throughput):** ข้อความแชทเพิ่มขึ้นตลอดเวลา MongoDB อ่านและเขียนข้อมูลอาร์เรย์ในเอกสารเดียวได้รวดเร็วโดยไม่ต้อง Join ตาราง"
        ]
    ]
    t3 = doc.add_table(rows=len(db_data)+1, cols=4)
    style_table(t3, [2.5, 3.0, 4.5, 6.6], db_headers, db_data)

    add_p(doc, "", size=6, space_a=6)

    # -------------------------------------------------------------
    # Section 4: AI & RAG System
    # -------------------------------------------------------------
    add_heading_1(doc, "4. ปัญญาประดิษฐ์และระบบ RAG (AI, LLM Integration & RAG Architecture)")
    
    ai_headers = ["ส่วนประกอบ", "เทคโนโลยี / โมเดล", "รายละเอียดและกลไกการทำงาน"]
    ai_data = [
        [
            "**Multi-LLM Provider**",
            "• **Ollama** (Local/On-Prem)\n• **Google Gemini API**\n• **Anthropic Claude**\n• **OpenAI API**",
            "• ระบบหลังบ้านออกแบบเป็น Interface รองรับการสลับผู้ให้บริการโมเดลภาษาได้อิสระจากหน้าจอตั้งค่าของแอปพลิเคชัน\n• ร้านค้าเลือกเชื่อมต่อ Local LLM เพื่อความเป็นส่วนตัวของข้อมูล หรือใช้ Cloud LLM เพื่อความฉลาดในการตอบ"
        ],
        [
            "**RAG Knowledge Base**",
            "**Text Chunking**\n+ **Dense Embeddings**\n+ **Cosine Similarity**",
            "• **Chunking:** นำเอกสารคู่มือ/สินค้าของร้านมาตัดแบ่งเป็นท่อนย่อย\n• **Embedding:** แปลงเนื้อหาเป็น Dense Vector บันทึกลงในคอลเลกชัน `knowledge_chunks` ของ MongoDB\n• **Retrieval:** เมื่อลูกค้าทักถาม ระบบแปลงคำถามเป็นเวกเตอร์และคำนวณความคล้ายโคไซน์ (Cosine Similarity) เพื่อดึงเฉพาะเนื้อหาที่เกี่ยวข้องส่งไปเป็นบริบท (Context) ให้ AI ตอบ"
        ],
        [
            "**Chatbot Control Modes**",
            "1. **Manual** (คนตอบ 100%)\n2. **Suggest** (AI เสนอคำตอบ)\n3. **Autopilot** (AI ตอบอัตโนมัติ)",
            "• เจ้าหน้าที่สามารถปรับโหมดการทำงานของแชทบอทแยกตามแต่ละบทสนทนาได้ทันทีตามสถานการณ์จริง\n• ในโหมด Autopilot หากแชทบอทตอบไม่ได้ ระบบจะตัดเข้าโหมดส่งต่อเจ้าหน้าที่และแจ้งเตือนทันที"
        ],
        [
            "**Conversation Analytics**",
            "LLM Prompt Pipeline",
            "• วิเคราะห์สรุปบทสนทนา (Summary)\n• จำแนกอารมณ์ลูกค้า (Sentiment: Positive, Neutral, Negative)\n• วิเคราะห์เจตนา (Intent) และตรวจจับสัญญาณสำคัญ (เช่น ลูกค้าเร่งด่วน, พร้อมชำระเงิน)"
        ]
    ]
    t4 = doc.add_table(rows=len(ai_data)+1, cols=3)
    style_table(t4, [3.2, 3.8, 9.6], ai_headers, ai_data)

    add_p(doc, "", size=6, space_a=6)

    # -------------------------------------------------------------
    # Section 5: Omnichannel Integration & DevOps
    # -------------------------------------------------------------
    add_heading_1(doc, "5. การเชื่อมต่อแพลตฟอร์มภายนอกและโครงสร้างพื้นฐาน (Integration & DevOps)")
    
    infra_headers = ["หมวดหมู่", "แพลตฟอร์ม / เครื่องมือ", "แนวทางการประยุกต์ใช้งานจริง"]
    infra_data = [
        [
            "**LINE Integration**",
            "**LINE Messaging API**\n(Channel Access Token v2.1)",
            "• ร้านค้าเชื่อมต่อด้วย Channel ID และ Channel Secret\n• รับ Webhook ขาเข้า และส่งข้อความขาออกผ่าน LINE Push Message API\n• ปิดบอทอัตโนมัติของ LINE Official Account Manager เพื่อให้ระบบ Chatshop ควบคุมแต่เพียงผู้เดียว"
        ],
        [
            "**Meta Integration**",
            "**Meta Graph API (v19.0)**\n(Facebook & Instagram)",
            "• **แยกเป็น 2 Meta Apps ในระบบ:**\n  1. *Consumer App:* สำหรับ Facebook Login ของผู้ใช้ทั่วไป\n  2. *Business App:* สำหรับรับ Webhook และส่งข้อความผ่าน Facebook Page & Instagram Messaging\n• แก้ปัญหาการขอสิทธิ์ขั้นสูง (Advanced Permissions) ทำให้ผู้ใช้ที่ไม่มีเพจสามารถเข้าใช้งานแอปได้"
        ],
        [
            "**WebChat Integration**",
            "**JavaScript Embed Widget**",
            "• สคริปต์วิดเจ็ตขนาดเล็กสำหรับนำไปติดบนหน้าเว็บไซต์ของร้านค้าเพื่อเปิดกล่องแชทสดคุยเชื่อมโยงเข้ามายังแอป Chatshop"
        ],
        [
            "**CI/CD Pipeline**",
            "**Bitbucket Pipelines**\n(Automated Deployment)",
            "• เมื่อรวมโค้ดเข้าสาขา `release` ระบบจะรันชุดทดสอบ ติดตั้งไลบรารี และ Deploy ไฟล์ขึ้นเซิร์ฟเวอร์ผ่าน SSH ด้วย `rsync` อัตโนมัติ โดยไม่เขียนทับไฟล์ `.env` และพื้นที่เก็บรูปภาพ"
        ],
        [
            "**Architecture Design**",
            "**Multi-Tenant Architecture**\n& **RBAC Security**",
            "• **Multi-Tenant:** แยกข้อมูลของแต่ละร้านค้าอย่างเด็ดขาดด้วย `shop_id` ทุกคำสั่งค้นหา ป้องกันปัญหาข้อมูลรั่วไหลข้ามร้าน\n• **RBAC:** ควบคุมสิทธิ์การเข้าถึงเมนูและฟังก์ชันตามบทบาท (Owner, Admin, Member)"
        ]
    ]
    t5 = doc.add_table(rows=len(infra_data)+1, cols=3)
    style_table(t5, [3.2, 3.8, 9.6], infra_headers, infra_data)

    add_p(doc, "", size=6, space_a=6)

    # -------------------------------------------------------------
    # Section 6: Thesis Defense Q&A Cheat Sheet
    # -------------------------------------------------------------
    add_heading_1(doc, "6. แนวคำถาม-คำตอบเก็งข้อสอบสำหรับป้องกันโครงงาน (Defense Q&A Cheat Sheet)")
    
    qa_list = [
        (
            "Q1: ทำไมเลือกพัฒนาแอปด้วย Flutter แทนที่จะเป็น Native (Kotlin บน Android / Swift บน iOS)?",
            "A: เนื่องจากระบบ Chatshop ต้องการรองรับผู้ใช้งานทั้งบน Android และ iOS ในอนาคต การเลือกใช้ Flutter ช่วยให้สามารถพัฒนาบนฐานโค้ดชุดเดียวกัน (Single Codebase) แต่ได้ประสิทธิภาพการทำงานและการวาดผลหน้าจอในระดับเนทีฟ มีเอนจินกราฟิกของตัวเองทำให้ UI แสดงผลเหมือนกันทุกรุ่น ทุกยี่ห้อสมาร์ตโฟน และช่วยลดระยะเวลาในการพัฒนาลงได้อย่างมีนัยสำคัญครับ"
        ),
        (
            "Q2: ทำไมเลือกใช้ Riverpod ในการจัดการ State แทนที่จะเป็น Bloc หรือ GetX?",
            "A: Riverpod ได้รับการพัฒนาขึ้นเพื่อแก้ข้อจำกัดของ Provider เดิม โดยทำงานเป็นอิสระอยู่นอก Widget Tree ทำให้สามารถจัดการ Global Services เช่น การสั่งตัดการเชื่อมต่อ WebSocket หรือการล้างโทเคนเมื่อ Logout ได้อย่างสะดวกมาก และที่สำคัญที่สุดคือ Riverpod มีความปลอดภัยสูงมากเพราะตรวจจับข้อผิดพลาดเรื่อง Type และ Dependency ตั้งแต่ขั้นตอน Compile (Compile-time Safety) ป้องกันการเกิดข้อผิดพลาด Runtime Error ระหว่างผู้ใช้เปิดแอปครับ"
        ),
        (
            "Q3: ทำไมฐานข้อมูลต้องใช้ทั้ง MySQL และ MongoDB ร่วมกัน ทำไมไม่ใช้ตัวใดตัวหนึ่ง?",
            "A: เพราะลักษณะของข้อมูลในระบบมีความต้องการทางวิศวกรรมที่แตกต่างกันอย่างสิ้นเชิงครับ\n1) ข้อมูลร้านค้า สมาชิก สินค้า และคำสั่งซื้อ เป็นข้อมูลเชิงโครงสร้างที่ต้องการความถูกต้องแม่นยำและเสถียรภาพทางการเงินระดับสูงสุด (ACID Compliance) จึงใช้ MySQL ที่มี Foreign Keys และ Transaction คอยควบคุม\n2) ในขณะที่ข้อความแชทจาก 4 แพลตฟอร์มมีโครงสร้างไม่แน่นอน (ข้อความ, รูปภาพ, สติกเกอร์, พิกัด) และมีปริมาณไหลเข้าสู่ระบบอย่างต่อเนื่องมหาศาล MongoDB ซึ่งเป็น Document NoSQL ตอบโจทย์เรื่องความยืดหยุ่นของฟิลด์และมีประสิทธิภาพในการอ่านเขียนข้อมูลปริมาณมากได้เร็วกว่าการ Join ตารางใน RDBMS ครับ"
        ),
        (
            "Q4: หากมีลูกค้าทักแชทเข้ามาพร้อมกันจากหลายช่องทางจำนวนมาก เซิร์ฟเวอร์จะค้างหรือล่มหรือไม่?",
            "A: ไม่ค้างครับ เพราะระบบออกแบบเป็น Asynchronous Queue Architecture ทันทีที่ Webhook จาก LINE หรือ Meta วิ่งเข้ามา ระบบหลังบ้านจะทำเพียงแค่ตรวจสอบความถูกต้องของลายเซ็นดิจิทัล (HMAC) แล้วผลักงานเข้าคิว High ทันที พร้อมกับส่งการตอบกลับ HTTP Status 200 กลับไปหาแพลตฟอร์มต้นทางภายในเวลาไม่กี่มิลลิวินาที แพลตฟอร์มจึงไม่เกิดการส่งข้อความซ้ำ จากนั้น Worker ในเบื้องหลังจะทยอยดึงงานไปประมวลผลและส่งต่อผ่าน WebSocket และ FCM ไปยังแอปพลิเคชันอย่างราบรื่นครับ"
        ),
        (
            "Q5: ทำไมระบบแชทบอทต้องทำกระบวนการ RAG เอง ทำไมไม่ส่งเอกสารทั้งหมดไปใน Prompt เลย?",
            "A: เพราะข้อมูลแคตตาล็อกสินค้าและเงื่อนไขการให้บริการของร้านค้ามีขนาดยาวมาก หากส่งเนื้อหาทั้งหมดไปในทุกคำขอ จะทำให้สิ้นเปลืองค่าใช้จ่ายโทเคน (Token Cost) สูงมาก และโมเดลภาษาอาจเกิดอาการภาพหลอน (Hallucination) หรือหลงลืมเนื้อหาตรงกลาง การใช้ RAG ด้วยการทำ Chunking และคำนวณความคล้ายโคไซน์ (Cosine Similarity) ทำให้เราส่งเฉพาะส่วนย่อยของเนื้อหาที่ตรงกับคำถามของลูกค้าที่สุดไปเป็นบริบท ทำให้ AI ตอบได้อย่างแม่นยำ ตรงประเด็น และประหยัดค่าใช้จ่าย API ได้ถึง 80-90% ครับ"
        )
    ]
    
    for q, a in qa_list:
        p_q = add_p(doc, q, bold=True, size=15, space_b=4, space_a=1, color_rgb=RGBColor(24, 76, 120))
        p_a = add_p(doc, a, size=14, space_b=0, space_a=5, line_spacing=1.15, color_rgb=RGBColor(40, 40, 40))
        pPr = p_a._p.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:left="400"/>'))

    # Save DOCX
    doc.save(DOCX_OUT)
    print(f"Generated DOCX: {DOCX_OUT}")
    
    # Export to PDF via Word COM
    print("Exporting PDF via Word COM...")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        wdoc = word.Documents.Open(os.path.abspath(DOCX_OUT))
        wdoc.ExportAsFixedFormat(os.path.abspath(PDF_OUT), 17) # 17 = wdExportFormatPDF
        wdoc.Close(False)
        print(f"Generated PDF: {PDF_OUT}")
    finally:
        word.Quit()

if __name__ == "__main__":
    build_tech_document()
