# Chatshop — เอกสารส่งต่องานทำเล่มปริญญานิพนธ์

> อ่านไฟล์นี้ก่อนทำงานต่อ ไม่ว่าจะเป็นคนหรือ AI
> สรุปจาก `input/SCOPE.md`, `.agent/rules/chatshop.md` (ตรวจกับโค้ดเมื่อ 2026-09-30) และสิ่งที่ผู้จัดทำยืนยันระหว่างทำเล่ม (ถึง 2026-10-07)
> ถ้าข้อมูลในไฟล์นี้ขัดกับโค้ดจริง ให้เชื่อโค้ด และถ้าขัดกับ `input/SCOPE.md` ให้เชื่อ SCOPE.md

---

## 1. โปรเจกต์คืออะไร

| | |
|---|---|
| ชื่อไทย | แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ (Chatshop) |
| ชื่ออังกฤษ | Chatshop: Centralized Chat and Order Management Application for Online Business |
| มหาวิทยาลัย | มทร.อีสาน วิทยาเขตขอนแก่น (รูปแบบเล่มตามคู่มือ พ.ศ. 2566) |
| ผู้พัฒนา | ผู้จัดทำคนเดียวทั้งระบบ **ห้ามเขียนว่ามีทีมพัฒนาหรือผู้ร่วมพัฒนา** ("ทีม" ในเล่มหมายถึงเจ้าหน้าที่ของร้านที่ใช้ระบบ) |

**ปัญหาที่แก้:** ร้านค้าออนไลน์ขายผ่านแชทหลายช่องทาง (LINE, Facebook Messenger, Instagram, เว็บไซต์) ต้องสลับหลายแอป ตอบช้า ข้อมูลลูกค้ากับออเดอร์กระจัดกระจาย
**สิ่งที่ทำ:** รวมแชททุกช่องทางไว้ในแอปพลิเคชันบนสมาร์ตโฟนแอปเดียว ให้เจ้าหน้าที่ตอบแชท เปิดคำสั่งซื้อ และใช้แชทบอทช่วยตอบได้

**จุดเน้นของเล่ม:** เล่มนี้เน้น**แอปพลิเคชันบนสมาร์ตโฟน (Flutter)** เป็นหน้าจอหลักของผู้ใช้ ส่วนระบบหลังบ้าน (Laravel) อธิบายในฐานะส่วนที่รับข้อมูลจากแพลตฟอร์ม ประมวลผล และเปิด API ให้แอป ทั้งสองส่วนนำเสนอเป็นระบบเดียวกัน

---

## 2. ที่อยู่ไฟล์

| อะไร | ที่ไหน |
|---|---|
| โค้ดระบบหลังบ้าน (Laravel 10) | `C:\interm\chatshop` |
| โค้ดแอปพลิเคชัน (Flutter) | `C:\interm\chatshop_app` |
| ขอบเขตฟีเจอร์ (ผู้จัดทำยืนยันเอง) | `input/SCOPE.md` |
| ข้อเท็จจริงของระบบแบบละเอียด | `.agent/rules/chatshop.md` |
| กฎการจัดรูปเล่ม มทร.อีสาน | `.agent/skills/rmuti-thesis/SKILL.md` |
| เล่มรุ่นพี่ (ใช้ดู**รูปแบบ**เท่านั้น ห้ามเอาเนื้อหา) | `input/reference/ไฟล์เล่มรุ่นพี่ที่ตรวจผ่านแล้ว.docx` |
| ต้นฉบับข้อความแต่ละบท | `input/chapters/ch1.txt` … `ch5.txt`, `bib.txt` |
| รูปประกอบ | `input/figures/ch2/2.N.png`, `input/figures/ch3/3.N.png` |
| สคริปต์วาดแผนภาพบทที่ 3 (รูป 3.1–3.7) | `input/figures/ch3/src/gen.js` |
| สคริปต์สร้าง docx/pdf | `.agent/skills/rmuti-thesis/scripts/thesis_build.py` |
| ผลลัพธ์ | `output/Chatshop_บทที่N_<ชื่อบท>.docx/.pdf` |

---

## 3. Stack

| ส่วน | ใช้อะไร |
|---|---|
| แอปพลิเคชัน | Flutter 3.x, Dart, Flutter Riverpod, go_router (ShellRoute), laravel_reverb, firebase_messaging + local notifications, flutter_secure_storage, native Facebook/Google login, image/video/file picker |
| ระบบหลังบ้าน | Laravel 10, PHP 8.1+ |
| ฐานข้อมูล | **MySQL** = ข้อมูลเชิงสัมพันธ์ · **MongoDB** = `conversations`, `messages`, `knowledge_documents`, `knowledge_chunks`, `device_tokens` |
| เรียลไทม์ | Laravel Reverb (WebSocket, โพรโทคอล Pusher) ช่อง `private-tenant.{id}` |
| แจ้งเตือน | Firebase Cloud Messaging |
| ยืนยันตัวตน | Sanctum token (แอป), Session (เว็บ), Socialite (Facebook/Google) |
| คิว | `high` = webhook + นำเข้าประวัติ · `default` = แชทบอท, สรุปแชท, ทำดัชนีความรู้, broadcast · `typing` = สถานะกำลังพิมพ์ |
| ตัวจัดตารางเวลา | `sla:scan` และ `broadcasts:dispatch-due` ทุก 1 นาที, ลบคำเชิญหมดอายุทุกวัน |
| ส่งข้อความออก | LINE Push API, Meta Graph API v19.0 |
| สื่อ | PHP GD (ย่อรูปเป็น WebP ด้านยาว ≤ 2,000 px), ffmpeg (ย่อวิดีโอด้านยาว 1,280 px) |

Route ทั้งหมด 297 route ในนั้นเป็น API ของแอป (`api/mobile/*`) 99 route

---

## 4. ผู้ใช้งาน

| ผู้ใช้ | ทำอะไรได้ |
|---|---|
| เจ้าของร้าน | ทุกอย่างของเจ้าหน้าที่ + จัดการช่องทาง ทีม บทบาท การตั้งค่าร้าน แชทบอท โอนความเป็นเจ้าของ |
| เจ้าหน้าที่ | ตอบแชท จัดการบทสนทนา ลูกค้า สินค้า คำสั่งซื้อ ตามสิทธิ์ของบทบาท |
| ลูกค้า | ส่งข้อความผ่าน LINE / Messenger / Instagram / WebChat และได้คำตอบจากเจ้าหน้าที่หรือแชทบอท (ไม่ได้ใช้แอป) |

บทบาท: owner / manager / admin / agent / viewer + บทบาทที่ร้านสร้างเอง
สิทธิ์ 12 รายการ: reply, assign, tags, broadcast, orders, products, export, reports, channels, settings, team, billing

---

## 5. ฟีเจอร์ในขอบเขตเล่ม

### 5.1 ระบบหลังบ้าน

| ฟีเจอร์ | สรุป | ไฟล์หลัก |
|---|---|---|
| สมัคร/เข้าสู่ระบบ | email+password, Facebook/Google, สมัครพร้อมสร้างร้าน · แอปใช้ `mobile/login`, `mobile/register`, `mobile/auth/{facebook\|google}` ได้ Sanctum token | `AuthController`, `SocialAuthController`, `TenantController@register` |
| เชื่อมช่องทาง | LINE, Facebook Page, Instagram (ผ่านเพจที่ผูก), WebChat widget | `ChannelController`, `WebChatChannelController` |
| รับ Webhook | LINE `POST webhook/line/{credential}` ตรวจ `X-Line-Signature` · Meta `GET/POST webhook/meta` ตรวจ verify token และ `X-Hub-Signature-256` · ผ่านแล้วส่งเข้าคิว `high` ทันที | `LineWebhookController`, `MetaWebhookController` |
| ประมวลผลข้อความขาเข้า | บันทึกลูกค้า (MySQL) → บทสนทนา+ข้อความ (MongoDB) → คัดลอกสื่อ → ปลดล็อกบทสนทนา → เรียกแชทบอท → broadcast `MessageReceived` + แจ้งเตือน | `ProcessLineWebhook`, `ProcessMetaWebhook` |
| ส่งข้อความออก | LINE Push API (quick reply, loading indicator) · Graph API v19.0 · เกิน 24 ชม. เจ้าหน้าที่ส่งซ้ำด้วยแท็ก `HUMAN_AGENT` ส่วนอัตโนมัติไม่ส่ง · token หมดอายุ (error 190) ปิดช่องทาง | `LineService`, `FacebookService`, `InstagramService` |
| สื่อ | proxy รูป LINE ด้วย signed URL, คัดลอกสื่อขาเข้า ≤ 25 MB, ย่อรูป/วิดีโอ | `MediaMirror`, `ImageCompressor`, `VideoCompressor` |
| บรอดแคสต์ | กลุ่มเป้าหมาย 6 แบบ, ตั้งเวลา, ส่งชุดละ 100, ส่งเฉพาะข้อความ | `BroadcastService`, `ProcessBroadcastJob` |
| แท็ก | แท็กพร้อมสี `#RRGGBB`, ลบแล้วถอดออกทุกบทสนทนา | `TenantTagController` |
| นำเข้าประวัติแชท | เฉพาะ Facebook/Instagram ย้อนหลัง ≤ 90 วัน, ชน rate limit รอ 60 วิ ลองใหม่ ≤ 3 ครั้ง | `ImportController`, `ImportHistoryJob` |
| ทีมและบทบาท | เชิญทางอีเมล (หมดอายุ 7 วัน), บทบาทกำหนดเอง, โอนเจ้าของ | `TeamController`, `InvitationController` |
| ตอบอัตโนมัติด้วยคีย์เวิร์ด | กฎต่อธุรกิจ, จับคู่ `exact`/`contains`, ทำงานแยกจากแชทบอท | `AutomationRuleController`, `ProcessAutomationReply` |
| แชทบอท | ดูหัวข้อ 6 | `ChatbotService`, `ProcessChatbotReply` |
| ฐานความรู้ (RAG) | เอกสาร → แบ่งส่วน ≤ 500 ตัวอักษร → embedding (Gemini `text-embedding-004` หรือ Ollama `bge-m3`) → cosine ≥ 0.3 เอา 4 อันดับแรก **แบบ brute-force ใน PHP ไม่มี vector database** | `KnowledgeController`, `Chatbot/Knowledge/*` |
| สินค้า / ลูกค้า | CRUD สินค้า (ชื่อ ราคา คำอธิบาย), ลูกค้า (เบอร์ ที่อยู่ บันทึก ≤ 2,000 ตัวอักษร) | `ProductController`, `CustomerController` |
| คำสั่งซื้อ | เปิดจากแชท **สินค้า 1 ชนิด × จำนวน** (ไม่ใช่ตะกร้า), 4 สถานะ, ตั้งที่อยู่จัดส่งแล้วอัปเดตที่อยู่ลูกค้าด้วย | `ChatOrderController` |
| แชทค้างตอบ (SLA) | เกณฑ์ 5/10/15/20 นาที (ค่าเริ่มต้น 5), สแกนทุก 1 นาที, แจ้งผู้มีสิทธิ์ `reply` ผ่าน Reverb + FCM ครั้งเดียว | `ScanSlaBreaches` |
| ตั้งค่าร้าน/ธุรกิจ | หลายธุรกิจต่อร้าน, ผูกเพจกับธุรกิจ, ข้อความสำเร็จรูป, การแจ้งเตือนรายผู้ใช้ | `TenantSettingsController`, `BusinessController`, `SavedReplyController` |
| รายงาน | first response, resolution, SLA, ปริมาณ, สัดส่วนช่องทาง, heatmap, export CSV | `OverviewController` |
| เรียลไทม์ | event: MessageReceived, NotificationCreated, NotificationRead, ConversationUpdated, MessagePinned/Unpinned, ChatbotTyping, BroadcastSent, ChannelConnected, OrderUpdated, SavedReplyChanged, SettingsChanged | `app/Events/*` |
| Push | ส่ง FCM ไปเครื่องที่ลงทะเบียน (`mobile/devices`) | `PushService` |

### 5.2 แอปพลิเคชันบนสมาร์ตโฟน

โครงสร้างโค้ด `lib/`: **core** (เก็บโทเคน, ตัวเรียก API, เรียลไทม์, แจ้งเตือน, go_router) · **data** (แบบจำลองข้อมูล + Repository) · **features** (13 โมดูล) · **shared** (วิดเจ็ตใช้ร่วม)

13 โมดูล: การยืนยันตัวตน, กล่องข้อความ, หน้าต่างสนทนา, การค้นหา, การแจ้งเตือน, ลูกค้า, สินค้า, คำสั่งซื้อ, ทีม, รายงาน, ช่องทาง, ข้อมูลผู้ใช้, การตั้งค่า

**หน้าตาจริงของแอป (จากภาพหน้าจอของผู้จัดทำ, UI เป็นภาษาอังกฤษ):**
- แถบนำทางล่าง 3 แท็บ: **Inbox** (กล่องข้อความ) · **Activity** (การแจ้งเตือน) · **You** (ข้อมูลผู้ใช้)
- แท็บ **You** เรียงจากบนลงล่าง: สถิติ + เมนูจัดการ (… Orders, Products, Reporting) → หัวข้อ **SETTINGS**: General, Tags, Automation, AI Chatbot, Businesses → หัวข้อ **NOTIFICATIONS**: Push notifications, SLA alerts → หัวข้อ **ACCOUNT**: Profile, Sign out → `chatshop mobile · v1.0`
- **การตั้งค่าไม่ใช่หน้าแยก** เป็นส่วนหนึ่งของแท็บ You
- แผ่นข้อมูลลูกค้าในหน้าต่างสนทนา: ชื่อ, CHANNELS, ปุ่ม "+ สร้างออเดอร์ · New order", ปุ่มตะกร้า

| ชื่อในเล่ม | ชื่อในแอป |
|---|---|
| ร้านค้า | General |
| แท็ก | Tags |
| ระบบอัตโนมัติ (ล็อกบทสนทนา + โหมดแชทบอทเริ่มต้น) | Automation |
| แชทบอท (ผู้ให้บริการ, คีย์ API, แบบจำลอง, ทดสอบการเชื่อมต่อ) | AI Chatbot |
| ธุรกิจ | Businesses |
| การแจ้งเตือน (เสียง, push, เฉพาะที่ได้รับมอบหมาย, ปิดบางช่องทาง) | Push notifications |
| แชทค้างตอบ | SLA alerts |

ความสามารถหลักของแอป:
- กล่องข้อความรวมทุกช่องทาง เรียงปักหมุดก่อน กรองได้ ปัดเพื่อจัดเก็บ/ลบ กดค้างเพื่อปักหมุด/มอบหมาย อัปเดตเรียลไทม์
- หน้าต่างสนทนา: ตอบ/แนบรูป วิดีโอ ไฟล์/ข้อความสำเร็จรูป, ตอบอ้างอิง, ปักหมุดข้อความ, react, แท็ก, ล็อกบทสนทนา
- ควบคุมแชทบอท 3 โหมดต่อบทสนทนา (แสดงผลทันที ย้อนกลับถ้าเซิร์ฟเวอร์ตอบว่าไม่สำเร็จ)
- แผ่นวิเคราะห์บทสนทนา (`POST .../analysis` ใช้ LLM จริง): บทสรุป, ความรู้สึก, เจตนา, หัวข้อ, สัญญาณ (ลูกค้าประจำ/เร่งด่วน)
- เปิดคำสั่งซื้อจากแชท (เติมชื่อสินค้าจากแคตตาล็อก)
- การแจ้งเตือน 5 ประเภท: urgent, message, paid, mention, system
- คำสั่งซื้อแยกแท็บตามสถานะ + ค้นหา + โหลดเพิ่มเมื่อเลื่อน
- รายงานในแอปแสดงเฉพาะการ์ดสรุป + สัดส่วนช่องทาง (`GET mobile/overview?range=7d`)
- ได้รับ 401 → ลบโทเคนแล้วกลับหน้าเข้าสู่ระบบ

---

## 6. แชทบอท (รายละเอียดที่เล่มใช้)

- 3 โหมด: **Manual** (ไม่ตอบ) / **Suggest** (แนะนำคำตอบให้เจ้าหน้าที่) / **Autopilot** (ตอบลูกค้าเอง) ตั้งค่าเริ่มต้นระดับธุรกิจ และเปลี่ยนรายบทสนทนาได้
- รอ **6 วินาที** รวมข้อความที่ส่งติดกัน (≤ 20 ข้อความ ห่างกันไม่เกิน 120 วิ) · ลองใหม่ 3 ครั้ง · timeout 150 วิ
- ข้อความสั่งงานประกอบจาก: บุคลิกร้าน, คำสั่งเพิ่มเติม, ข้อมูลแบรนด์/การชำระเงิน, บทสรุปสะสม, ผลค้นคืนความรู้, ข้อความล่าสุด **20** ข้อความ
- ผู้ให้บริการ: Ollama (ค่าเริ่มต้น `qwen3:14b`), Anthropic Claude, Google Gemini, บริการที่รองรับรูปแบบ OpenAI API
- ส่งต่อเจ้าหน้าที่: แบบจำลองตอบ `<<ESCALATE>>` → ติดแท็ก "รอแอดมิน" + เปลี่ยนเป็น Manual (แท็กถอดเมื่อเจ้าหน้าที่ตอบ)
- รูป/เสียง/วิดีโอ/ไฟล์ส่งต่อเจ้าหน้าที่พร้อมข้อความรับทราบ · รูปภายใน 24 ชม. หลังพูดเรื่องชำระเงิน = สลิปโอนเงิน
- เขียนในเล่มว่า "แชทบอท" เฉยๆ **ห้ามใช้ "แชทบอทปัญญาประดิษฐ์"**

---

## 7. การตั้งค่าแพลตฟอร์มภายนอก (บทที่ 3 หัวข้อ 3.8)

### LINE Official Account
- ตั้งแต่ 4 ก.ย. 2567 ต้องสร้างช่องทาง Messaging API ผ่านบัญชี LINE OA ก่อน
- ร้านกรอก Channel ID + Channel Secret ในแอป → หลังบ้านขอ Channel Access Token, ดึงชื่อ/รูปบัญชี, ตั้ง Webhook URL `/webhook/line/{รหัสช่องทาง}` ให้อัตโนมัติ
- ใน LINE Developers: เปิด Use webhook, ปิดข้อความตอบกลับอัตโนมัติ/ทักทายของ LINE
- เล่มไม่ต้องอธิบายขั้นสมัครบัญชีละเอียด (ผู้จัดทำตัดรูปขั้นสมัครออกแล้ว)

### Meta — มี 2 แอปแยกกัน (สำคัญ เคยเขียนผิด)
| แอป | ประเภท | Valid OAuth Redirect URIs |
|---|---|---|
| **แอปเข้าสู่ระบบ** | Consumer | `/auth/facebook/callback` **อย่างเดียว** (ยืนยันจากภาพหน้าจอ) |
| **แอปเชื่อมต่อช่องทาง** | Business | `/channels/facebook/callback`, `/channels/instagram/callback`, `/mobile/oauth/facebook/callback` |

- เหตุที่แยก 2 แอป: แอป Business ไม่มี Facebook Login ปกติ ถ้าใช้แอปเดียวต้องขอ `business_management` ทำให้คนไม่มีเพจเข้าสู่ระบบไม่ได้
- แอปเข้าสู่ระบบ: สิทธิ์ `public_profile`, `email` · แพลตฟอร์ม Android: แพ็กเกจ `co.chatshop.chatshop`, คลาส `co.chatshop.chatshop.MainActivity`, Key Hash · ตั้งค่าที่ยืนยันแล้ว: Client OAuth login ใช่, Web OAuth login ใช่, บังคับ HTTPS ใช่, โหมดเคร่งครัด ใช่, เข้าสู่ระบบจากอุปกรณ์ ไม่ใช่, JS SDK ไม่ใช่
- **App Domains อยู่ที่ App settings → Basic ไม่ใช่หน้า Facebook Login Settings**
- แอปเชื่อมต่อช่องทาง: Facebook Login for Business + Messenger, Configuration ชื่อ `connect` (User access token), สิทธิ์ เช่น `pages_show_list`, `pages_messaging`, `pages_manage_metadata`, `pages_read_engagement` (ตรวจรายการเต็มกับโค้ดก่อนใช้)
- Webhook: Callback `/webhook/meta`, Verify Token ที่ผู้จัดทำกำหนด, ตรวจลายเซ็นด้วย App Secret
- Instagram ต้องเป็นบัญชีมืออาชีพที่ผูกกับเพจ เพราะเชื่อมผ่านเพจ

### Firebase / Google
- แอป Android `co.chatshop.chatshop` + SHA-1 → `google-services.json` ในโปรเจกต์ Flutter
- คีย์บัญชีบริการใช้ส่ง FCM จากหลังบ้าน, หน้าขอความยินยอม OAuth สำหรับ Google Sign-In

### WebChat
- ไม่ต้องตั้งค่าแพลตฟอร์มภายนอก สร้างจากแอป กำหนดชื่อ ข้อความทักทาย สี แล้วนำสคริปต์ไปติดเว็บ · รับข้อความใหม่ด้วย Server-Sent Events

### ข้อมูลที่ต้องปิดบังในภาพหน้าจอ
- **ปิด:** App Secret, Channel Secret, Channel Access Token, คีย์ API ของ LLM, คีย์การถอดรหัส Install Referrer, คีย์บัญชีบริการ
- **แสดงได้:** ชื่อแพ็กเกจ, ชื่อคลาส, Key Hash, โดเมนหลังบ้าน, App ID
- ใช้กล่องสีทึบ 100% (ปากกาไฮไลต์ยังมองทะลุได้) หรือครอปส่วนนั้นออก

---

## 8. ขอบเขต: ใส่ / ไม่ใส่ / ห้ามอ้าง

**ไม่ใส่ในเล่ม (มีในโค้ด):** แผงผู้ดูแลแพลตฟอร์ม (`admin/*`), CSAT แบบประเมินความพึงพอใจ, CF-Shops SSO, แพ็กเกจ/subscription/โควตา/ค่าใช้จ่าย LLM (หน้าแพ็กเกจในแอปก็ไม่ต้องพูดถึง)

**ห้ามอ้าง (ไม่มีจริง):** vector database, ตะกร้าหลายสินค้า, payment gateway, แอป iOS พร้อม release (มีแค่ Android), WhatsApp/TikTok/Shopee/Lazada, Horizon, Vue, dio

**คำที่ใช้:**
- "แอปพลิเคชันบนสมาร์ตโฟน" (สะกด **สมาร์ตโฟน**) ไม่ใช้ "อุปกรณ์เคลื่อนที่" ไม่ใช้ "มือถือ"
- "แชทบอท" ไม่ใช่ "แชทบอทปัญญาประดิษฐ์"
- SLA ไม่ใช่จุดเด่น เขียนว่า "แจ้งเตือนแชทที่ค้างตอบ" ไม่ตั้งเป็นวัตถุประสงค์ ไม่ใช้คำว่า "เกณฑ์ SLA" ในบทที่ 1–2
- "ผู้จัดทำพัฒนา..." ทั้งสองส่วน

---

## 9. สถานะของเล่ม (2026-10-07)

| บท | สถานะ | ไฟล์หลัก |
|---|---|---|
| 1 บทนำ | build จาก `ch1.txt` | `output/Chatshop_บทที่1_บทนำ.docx` |
| 2 ทฤษฎี | build จาก `ch2.txt`, รูปครบ 19 รูป | `output/Chatshop_บทที่2_...docx` |
| 3 วิธีดำเนินงาน | **ผู้จัดทำแก้ใน Word แล้ว ไฟล์ docx คือฉบับหลัก** | `output/Chatshop_บทที่3_วิธีดำเนินงาน_v2.docx` (2026-10-09) |
| 4 ผลการดำเนินงาน | build จาก `ch4.txt` · ฉบับตัดตาราง (ต่อ) แล้ว = `_v2` | `output/Chatshop_บทที่4_ผลการดำเนินงาน_v2.docx` |
| 5 สรุปผล | build จาก `ch5.txt` | `output/Chatshop_บทที่5_...docx` |
| บรรณานุกรม | build จาก `bib.txt` (APA 7) | `output/Chatshop_บรรณานุกรม.docx` |

### ⚠️ ข้อควรระวังบทที่ 3
- **ห้าม build บทที่ 3 จาก `ch3.txt` อีก** จะเขียนทับรูปที่ผู้จัดทำใส่เองใน Word ให้แก้ใน docx โดยตรง (python-docx บนสำเนา แล้วบันทึกเป็นไฟล์ใหม่ ห้ามทับไฟล์ที่ผู้จัดทำเปิดอยู่)
- รูป 3.1–3.7 เป็นแผนภาพที่สร้างจาก `gen.js` · รูปตั้งแต่ 3.8 เป็นภาพหน้าจอจริง
- รูปขั้นสมัคร LINE OA เดิม (3.29, 3.30) ถูกตัดแล้ว เลขรูปหลังจากนั้นเลื่อนลง 2 (ปัจจุบันมี 38 รูป)

### งานค้างบทที่ 3
1. 3.8.2 Meta: (เสร็จแล้วเมื่อ 2026-10-07) แก้ไขเส้นทาง URL ตัด `/mobile/...` ออกจากแอปเข้าสู่ระบบ, แยกหน้า URL (Valid OAuth Redirect URIs) กับสิทธิ์ (Permissions) และย้าย App Domains ไปหน้า Basic settings ทั้งใน docx และ ch3.txt เรียบร้อยแล้ว พร้อมข้อความแนะนำการใส่รูปในกรอบ
2. ~~3.7.13 การตั้งค่า~~ (เสร็จใน v2 เมื่อ 2026-10-09): แก้ให้บอกว่าเป็นส่วนล่างของแท็บ You และเรียงหน้าย่อยตามแอป (General, Tags, Automation, AI Chatbot, Businesses, Push notifications, SLA alerts)
3. รูป 3.16 แผ่นข้อมูลลูกค้า: ภาพเดิมข้อมูลว่าง ต้องถ่ายลูกค้าที่มีที่อยู่ บันทึก แท็ก และประวัติคำสั่งซื้อ
4. รูป 3.26 หน้าข้อมูลผู้ใช้: ต้องเห็นสถิติด้านบนของแท็บ You
5. ~~3.5.7 การติดตั้งระบบ~~ (เสร็จใน v2 เมื่อ 2026-10-09): เขียนภาพรวมจาก bitbucket-pipelines.yml ไม่ระบุโดเมน
6. **ยังไม่ได้ตรวจบทที่ 3 ทั้งบทกับโค้ด** (เคยเริ่มแล้วหยุด) ควรตรวจทีละหัวข้อ 3.1–3.3, 3.4–3.6, 3.7, 3.8 โดยเทียบโค้ดจริงทุกตัวเลขและชื่อ

---

## 10. วิธีทำงานกับเล่ม

**Build บท (ยกเว้นบทที่ 3):** ปิด Word ก่อน (ถ้าเปิดค้างจะได้ไฟล์ `_new.docx`)
```bash
py .agent/skills/rmuti-thesis/scripts/thesis_build.py 2
```

**ประกอบเล่มเต็ม (ปกนอก → บรรณานุกรม):** ปิด Word ก่อน แล้วรัน
```bash
py build_book.py --rebuild-body
```
- ส่วนหน้าใช้ไฟล์ตัวอย่าง `input/reference/01 ปกนอก ถึง กิตติกรรมประกาศ.docx` (มีตรามหาวิทยาลัย)
- สารบัญ, สารบัญตาราง และสารบัญรูปวัดเลขหน้าจริงจาก Word
- ผลลัพธ์ `output/Chatshop_ปริญญานิพนธ์ฉบับสมบูรณ์.docx/.pdf`
- ภาคผนวก ก ยังไม่มีเนื้อหา เลขหน้าในสารบัญเป็นค่าประมาณจาก `APPENDIX_A_PAGES` ใน `build_book.py`
- ถ้า Word ถามว่า "ครั้งสุดท้ายที่เปิด … ทำให้เกิดข้อผิดพลาดร้ายแรง" (เกิดหลังสั่งปิด Word ที่ค้าง) ให้ตอบ "ใช่"

**รูปแบบไฟล์บท `chN.txt`:** บรรทัด 1 = เลขบท, บรรทัด 2 = ชื่อบท, `2.1 ...` = หัวข้อ, `@fig ชื่อรูป` = รูป (ดึง `input/figures/chN/N.n.png` อัตโนมัติ), `@table ชื่อ` + แถว `a | b` = ตาราง, `1) ...` = รายการ

**ไฟล์เก่า:** ย้ายไป `C:/Users/NIER/Downloads/doc/New folder/<YYYY-MM-DD>/` ตามวันที่แก้ไฟล์

**สิ่งที่สคริปต์ทำให้อัตโนมัติ:**
- เลขรูป/ตารางเป็น field `SEQ` (ทำสารบัญรูป/ตารางใน Word ได้) — อ้างถึงในเนื้อหาให้พิมพ์เลขเอง เช่น "ดังรูปที่ 3.1"
- ตารางล้นหน้า → `split_tables()` ให้ Word วัดหน้าจริงแล้วตัดตาราง ขึ้นหน้าใหม่ด้วย "ตารางที่ N.n  (ต่อ)" พร้อมหัวตารางซ้ำ
- หน้าแรกของบทเว้นขอบบน 5.08 ซม. หน้าถัดไป 3.81 ซม.
- เว้น 1 บรรทัดเหนือรูปและใต้คำบรรยายรูป (แบบเล่มรุ่นพี่)

**แผนภาพบทที่ 3:** แก้ `input/figures/ch3/src/gen.js` แล้ว `node input/figures/ch3/src/gen.js 3.2` (ใส่เลขรูปเพื่อทำเฉพาะรูปนั้น)

**กฎเขียน:** ดู `.agent/skills/rmuti-thesis/SKILL.md` — TH SarabunPSK 16pt, ไม่ใช้ hanging indent (ยกเว้นบรรณานุกรม), ไม่เว้นวรรคเทียมกลางประโยค, ไม่ใส่วงเล็บอังกฤษซ้ำคำไทย, เว้น 2 เคาะหลังเลขรูป/ตาราง

**ก่อนเขียนข้อเท็จจริงใหม่ทุกครั้ง:** เปิดโค้ดยืนยันก่อน และเช็กกับ `input/SCOPE.md` ว่าอยู่ในขอบเขต
