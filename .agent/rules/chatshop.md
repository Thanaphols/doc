# Project Fact Sheet: Chatshop (เล่มขอบเขตแอปพลิเคชันบนสมาร์ตโฟน)

> กฎการจัดหน้า/พิมพ์ไทยทั้งหมดอยู่ที่ skill `.agent/skills/rmuti-thesis/SKILL.md` ที่เดียว ไฟล์นี้เก็บ **ข้อเท็จจริงของโปรเจกต์** เท่านั้น
> ทุกข้อด้านล่างตรวจกับโค้ดจริงแล้ว (2026-09-30) ถ้าจะเขียนอะไรนอกเหนือจากนี้ ต้องเปิดโค้ดยืนยันก่อน

## ชื่อโปรเจกต์ (ใช้ชื่อนี้ชื่อเดียวทั้งเล่ม)
- **TH**: แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ (Chatshop)
- **EN**: Chatshop: Centralized Chat and Order Management Application for Online Business
- เล่มรุ่นพี่ "บ้านฟาร์มขนม" ใน `input/reference/` ใช้ดูสไตล์เท่านั้น **ห้ามนำเนื้อหามาใช้**

## Source code (ใช้ยืนยันข้อเท็จจริง)
- ระบบหลังบ้าน: `C:\interm\chatshop` — Laravel 10 (MySQL, MongoDB, Queue, Reverb, Sanctum)
- แอปพลิเคชัน: `C:\interm\chatshop_app` — Flutter 3.x, Dart (sdk ^3.10.7), Flutter Riverpod (^3.3.2), GoRouter (^17.5.0), laravel_reverb, firebase_messaging, flutter_secure_storage

## ขอบเขตงาน: ดู `input/SCOPE.md` เป็นหลัก
`input/SCOPE.md` คือรายการฟีเจอร์ที่ผู้จัดทำยืนยันเองว่าใส่/ไม่ใส่ในเล่ม ถ้าขัดกับไฟล์นี้ ให้ยึด SCOPE.md

## ผู้พัฒนา: ผู้จัดทำคนเดียวทั้งระบบ
**ไม่มีทีมพัฒนา ไม่มีผู้ร่วมพัฒนา** ผู้จัดทำออกแบบและพัฒนาเองทั้งหมด 2 ส่วนที่ทำงานร่วมกัน:

| ส่วน | สิ่งที่ผู้จัดทำพัฒนา |
|---|---|
| **ระบบหลังบ้าน (Laravel)** | ฐานข้อมูล MySQL/MongoDB, REST API `mobile/*` + Sanctum, เชื่อมช่องทาง + webhook ขาเข้า, ส่งข้อความออกแพลตฟอร์ม, จัดเก็บ/ย่อ media, คิวและ scheduler, เซิร์ฟเวอร์ Reverb + event, ส่ง push ผ่าน FCM, ทีมและสิทธิ์, บรอดแคสต์, แท็ก, นำเข้าประวัติ, ตอบอัตโนมัติด้วยคีย์เวิร์ด, แชทบอท + RAG, หน้าเว็บจัดการ (Blade/React) |
| **แอปพลิเคชันบนสมาร์ตโฟน (Flutter)** | หน้าจอแอปทั้งหมด (13 โมดูล), Riverpod, go_router (ShellRoute), รับ push (FCM + local notifications), Realtime (laravel_reverb), เก็บโทเคน (flutter_secure_storage), native Facebook/Google login, เลือก/ส่งสื่อ (image/video/file picker), ควบคุมแชทบอท + analysis sheet |

วิธีเขียนในเล่ม:
- ใช้ "ผู้จัดทำพัฒนา..." กับทั้งสองส่วน ห้ามเขียนว่าส่วนใด "พัฒนาโดยผู้อื่น" / "ผู้ร่วมพัฒนา" / "ทีมพัฒนา"
- นำเสนอเป็นระบบเดียว: แอปเป็นหน้าจอหลักของผู้ใช้ หลังบ้านรับข้อมูลจากแพลตฟอร์ม ประมวลผล และเปิด API ให้แอป
- "ทีมงาน" ในเล่มหมายถึงเจ้าหน้าที่ของร้านค้าที่ใช้ระบบ (ฟีเจอร์ทีมและบทบาท) ไม่ใช่ทีมพัฒนา

## ข้อเท็จจริงของระบบ

### ภาพรวม
- Multi-tenant: ร้านค้า (tenant) หนึ่งแห่งมีได้หลายธุรกิจ (`businesses`) แต่ละธุรกิจมีแบรนด์ ข้อมูลการชำระเงิน และการตั้งค่าแชทบอทของตัวเอง
- ช่องทาง: LINE OA, Facebook Messenger, Instagram, WebChat (TikTok มีใน enum แต่ยังไม่มี service)
- ฐานข้อมูล: **MySQL** = ข้อมูลเชิงสัมพันธ์ · **MongoDB** (`mongodb/laravel-mongodb`) = `conversations`, `messages`, `knowledge_documents`, `knowledge_chunks`, `device_tokens`
- Queue: `high` = webhook jobs + import · `default` = `ProcessChatbotReply`, `SummarizeConversationJob`, `IndexKnowledgeDocumentJob`, broadcast · `typing` = `KeepChatbotTypingJob`
- Scheduler (`app/Console/Kernel.php`): `sla:scan` และ `broadcasts:dispatch-due` ทุก 1 นาที, `model:prune` คำเชิญที่หมดอายุเกิน 30 วัน ทุกวัน
- Realtime: Laravel Reverb (WebSocket, Pusher protocol) ช่อง `private-tenant.{id}` · Push: Firebase Cloud Messaging
- App เชื่อม backend ด้วย REST `mobile/*` + Bearer token (Sanctum)

### ฝั่งเว็บ: กล่องข้อความ ช่องทาง และทีม
- **Inbox** (`ConversationController`, React ใน `public/js/inbox/`):
  - รายการบทสนทนาสูงสุด 200 รายการ เรียงแบบปักหมุดก่อน แล้วตามด้วยข้อความล่าสุด
  - กรองตามเพจ/ช่องทาง/ข้อความในหน้าเว็บ และค้นหาฝั่ง server ตามชื่อ/ข้อความ/แท็ก/ช่องทาง ได้สูงสุด 50 รายการ
  - โหลดข้อความแบบ cursor ครั้งละ 30 ข้อความ
  - การทำงาน: ตอบ, อ่านแล้ว, react, archive, ลบ, แท็ก, มอบหมาย (สิทธิ์ `assign`), ปักหมุดบทสนทนา/ข้อความ, export (≤ 5,000), ขอคำตอบแนะนำจากแชทบอท
  - ตอบได้เป็น ข้อความ (≤ 2,000 ตัวอักษร), สติกเกอร์, ไฟล์ (รูป ≤ 10 MB, ไฟล์อื่น ≤ 25 MB) และตอบกลับอ้างอิงข้อความได้
  - ข้อความขาเข้ารองรับ text/image/video/audio/file/sticker และบันทึกการยกเลิกข้อความ
- **Webhook ขาเข้า**:
  - LINE `POST webhook/line/{credential}` ตรวจ `X-Line-Signature` (HMAC-SHA256 ด้วย channel secret)
  - Meta (Facebook + Instagram ใช้ endpoint เดียวกัน) มี `GET` สำหรับยืนยัน verify token และ `POST` ตรวจ `X-Hub-Signature-256`
  - ผ่านการตรวจแล้วส่งงานเข้าคิว `high` ทันที
  - `ProcessLineWebhook` / `ProcessMetaWebhook`: บันทึกลูกค้า (MySQL) → บทสนทนาและข้อความ (MongoDB) → mirror media → ปลดล็อกบทสนทนา → เรียกงานแชทบอท → broadcast `MessageReceived` + การแจ้งเตือน
  - WebChat: widget ฝังบนเว็บไซต์ร้าน รับส่งผ่าน `api/webchat/*` และรับข้อความใหม่ด้วย Server-Sent Events
- **ส่งข้อความออก**:
  - LINE ใช้ Push API (`v2/bot/message/push`) รองรับ quick reply และ loading indicator
  - Facebook/Instagram ใช้ Graph API v19.0 `/{page}/messages`
  - เกินกรอบเวลา 24 ชั่วโมงของ Meta: การตอบโดยเจ้าหน้าที่จะส่งซ้ำพร้อมแท็ก `HUMAN_AGENT` ส่วนการส่งอัตโนมัติ (แชทบอท/บรอดแคสต์) จะไม่ส่ง
  - token หมดอายุ (error 190) → ปิดการเชื่อมต่อช่องทางนั้น
- **Media**:
  - รูปของ LINE ส่งผ่าน server ด้วย signed URL เพราะ LINE ไม่มี URL สาธารณะ
  - `MediaMirror` คัดลอก media ขาเข้า (≤ 25 MB) ไปเก็บใน storage ของระบบ แปลง JPEG/PNG/WebP เป็น WebP
  - ย่อรูปด้วย PHP GD (ด้านยาวไม่เกิน 2,000 px)
  - ย่อวิดีโอด้วย ffmpeg (ด้านยาว 1,280 px) พร้อมภาพตัวอย่าง
- **บรอดแคสต์**:
  - เลือกช่องทางได้ และเลือกกลุ่มเป้าหมายได้ 6 แบบ: ทั้งหมด, ลูกค้าใหม่ 7 วัน, ยังไม่ตอบ, เงียบเกิน 30 วัน, ตามแท็ก, เลือกเอง
  - ตั้งเวลาส่งได้ ตัวจัดตารางเวลาตรวจทุก 1 นาที ถ้าเลยเวลาเกิน 1 วันจะถือว่าล้มเหลว
  - ส่งเป็นชุด ชุดละ 100 บทสนทนา ส่งเฉพาะข้อความ และไม่ส่งซ้ำผู้รับที่ส่งสำเร็จแล้ว
  - สถานะ: draft/scheduled/sending/sent/failed พร้อมจำนวนผู้รับ/ส่งแล้ว/อ่านแล้ว
- **แท็ก**: รายการแท็กของร้านพร้อมสี (`#RRGGBB`) ใช้สิทธิ์ `tags` และลบแท็กแล้วจะถอดออกจากทุกบทสนทนา
- **นำเข้าประวัติแชท**:
  - รองรับเฉพาะ Facebook/Instagram (LINE ไม่รองรับ)
  - เลือกช่วงวันที่ได้ไม่เกิน 90 วัน ทำงานในคิว `high`
  - ถ้าชน rate limit จะรอ 60 วินาทีแล้วลองใหม่ ไม่เกิน 3 ครั้ง
  - ติดตามความคืบหน้าได้จาก `import_runs`
- **ทีมและบทบาท**:
  - บทบาท owner/manager/admin/agent/viewer + custom role (owner เป็นผู้สร้าง)
  - สิทธิ์ 12 รายการ: reply, assign, tags, broadcast, orders, products, export, reports, channels, settings, team, billing
  - โอนความเป็นเจ้าของร้านได้
  - คำเชิญทางอีเมลหมดอายุใน 7 วัน
- **การยืนยันตัวตน**:
  - เว็บ: email/password และ Facebook/Google ผ่าน Socialite
  - แอป: `mobile/login`, `mobile/register`, `mobile/auth/{facebook|google}` คืน Sanctum personal access token และ route ของแอปป้องกันด้วย `auth:sanctum`
- **REST API สำหรับแอป**: ประมาณ 99 route ภายใต้ `mobile/*` แบ่งกลุ่ม auth, conversations, AI, customers, orders, products, saved replies, notifications, devices, tags, team, overview, settings, channels
- **Realtime (Reverb)**:
  - ช่อง `private-tenant.{tenantId}`
  - event: MessageReceived, NotificationCreated, NotificationRead, ConversationUpdated, MessagePinned/Unpinned, ChatbotTyping, BroadcastSent, ChannelConnected, OrderUpdated, SavedReplyChanged, SettingsChanged
- **ตารางเพิ่มเติม**:
  - `broadcasts`, `broadcast_recipients`
  - `import_runs`
  - `tenant_tags`
  - `platform_credentials` (มี access_token, channel_secret และค่าของ widget)
  - `conversation_assignments`
  - `personal_access_tokens`
  - `device_tokens` เก็บใน MongoDB

### ตอบอัตโนมัติด้วยคีย์เวิร์ด (Automation rules)
- ตั้งกฎต่อธุรกิจ: ชื่อกฎ, คีย์เวิร์ดหลายคำ, วิธีจับคู่ `exact` (ตรงทั้งข้อความ) / `contains` (มีคำนี้), ข้อความตอบกลับ, เปิด/ปิดรายกฎ และเปิด/ปิดทั้งหมด (`AutomationRuleController`, สิทธิ์ `settings`)
- ข้อความขาเข้าที่ตรงกฎ → `ProcessAutomationReply` (คิว `high`) ส่งข้อความตอบกลับผ่าน LINE / Facebook / Instagram และบันทึกข้อความพร้อมชื่อกฎ ทำงานแยกจากแชทบอท
- หน้าตั้งค่าแสดงจำนวนกฎที่เปิดและจำนวนครั้งที่กฎทำงานในสัปดาห์นี้

### แชทบอท
- 3 โหมด: `manual` / `suggest` / `autopilot` — ค่าเริ่มต้นระดับธุรกิจ, override รายบทสนทนาได้ (`ChatbotService::effectiveMode`)
- `ProcessChatbotReply`: รอ 6 วินาทีเพื่อรวมข้อความที่ส่งมาติดกัน (สูงสุด 20 ข้อความ ห่างกันไม่เกิน 120 วินาที), กันตอบซ้ำด้วย cache lock + `reply_to_incoming_id`, retry 3 ครั้ง, timeout 150 วินาที, ถ้าล้มเหลวจะแจ้งเตือนเจ้าหน้าที่
- Prompt ประกอบจาก: persona, directives, ข้อมูลแบรนด์/การชำระเงิน, คำสั่ง escalation, บทสรุป, ผล RAG, ข้อความล่าสุด 20 ข้อความ (รวมข้อความ OCR จากรูป)
- LLM engine เลือกได้ต่อธุรกิจ: Ollama (ค่าเริ่มต้น `qwen3:14b`), Anthropic Claude, Google Gemini, OpenAI-compatible (OpenRouter, OpenAI, DeepSeek, Groq, Typhoon ฯลฯ) — `ReplyGeneratorFactory`, `config/chatbot_models.php`
- Escalation: model ตอบ `<<ESCALATE>>` → ติดแท็ก "รอแอดมิน" + เปลี่ยนเป็น manual + broadcast; แท็กจะถูกถอดเมื่อเจ้าหน้าที่ตอบ
- รูป/เสียง/วิดีโอ/ไฟล์ → ส่งต่อเจ้าหน้าที่พร้อมข้อความรับทราบอัตโนมัติ; รูปที่ส่งมาภายใน 24 ชั่วโมงหลังมีคำเกี่ยวกับการชำระเงิน → ถือเป็นสลิปโอนเงิน
- `KeepChatbotTypingJob`: ส่ง `typing_on` ให้ Messenger/IG ทุก 10 วินาที ไม่เกิน 6 ครั้ง
- `ConversationLockService`: ถ้าเปิด auto-lock เจ้าหน้าที่คนแรกที่ตอบจะล็อกบทสนทนา (ผู้มีสิทธิ์ `assign` ข้ามได้)
- `SummarizeConversationJob`: เก็บบทสรุปแบบสะสม (`context_summary`, `summarized_through`) โดยเก็บข้อความล่าสุด 10 ข้อความไว้แบบเต็ม
- บันทึก token/ค่าใช้จ่ายทุกครั้งที่เรียกใช้ (`chatbot_usage_logs`); แพ็กเกจกำหนดสิทธิ์ใช้ Anthropic/Gemini และโควตาเป็น USD ต่อเดือน
- (แอป) badge/แถบสถานะ AI, sheet สำหรับเปลี่ยนโหมด (optimistic update + rollback), `POST .../ai/suggest`, `PATCH .../ai-mode`; analysis sheet `POST .../analysis` ที่ใช้ LLM จริง → summary, sentiment, intent, topics, signals (VIP/returning/urgent) · `ai_mock.dart` เป็นแค่ fallback เมื่อเรียก endpoint ไม่สำเร็จ

### RAG / Knowledge
- CRUD เอกสาร FAQ (`KnowledgeController`, ต้องมีสิทธิ์ `settings`) → `IndexKnowledgeDocumentJob`
- `Chunker`: แบ่งตามย่อหน้า ส่วนละไม่เกิน 500 ตัวอักษร → embedding (Gemini `text-embedding-004` หรือ Ollama `bge-m3`) → เก็บเป็น array ใน MongoDB `knowledge_chunks`
- `Retriever`: คำนวณ cosine similarity แบบไล่เทียบทุกส่วนใน PHP (brute-force) คัดเฉพาะที่มีค่า ≥ 0.3 แล้วเอา 4 อันดับแรก — **ไม่มี vector database**
- `ProductController::syncKnowledge` นำข้อมูลสินค้าเข้าฐานความรู้ได้; มี job re-index ทั้งธุรกิจ/ทั้ง tenant

### Orders / Products / Customers
- คำสั่งซื้อ 1 รายการมี **สินค้า 1 ชนิด × จำนวน**, `total_amount = price × quantity` (แก้ไขทับได้) — **ไม่ใช่ตะกร้าหลายสินค้า**
- สถานะ: `pending` รอดำเนินการ · `paid` ชำระแล้ว · `shipped` จัดส่งแล้ว · `cancelled` ยกเลิก
- `chat_orders`: tenant_id, business_id, customer_id, conversation_id, order_number, product_id, product_name, price, quantity, total_amount, status, description, shipping_address, external_order_id
- การตั้ง shipping address จะอัปเดตที่อยู่ของลูกค้าไปด้วย
- `products`: tenant_id, business_id, name, price, description
- `customers`: tenant_id, platform, platform_id, name, phone, address, note (≤ 2000 ตัวอักษร), avatar_url
- (แอป) หน้าออเดอร์แยกแท็บตามสถานะ + ค้นหา + infinite scroll; form ช่วยเติมชื่อสินค้าอัตโนมัติจาก catalog; profile sheet แสดงประวัติคำสั่งซื้อของลูกค้าในบทสนทนา

### Settings / SLA / Notifications / Reports
- SLA: `tenants.sla_minutes` เลือกได้ 5/10/15/20 นาที (ค่าเริ่มต้น 5) · ถือว่า breach เมื่อมีข้อความยังไม่อ่าน ยังไม่มีการตอบครั้งแรก และข้อความล่าสุดเก่ากว่าเกณฑ์ → broadcast `NotificationCreated` + FCM ไปยังผู้มีสิทธิ์ `reply`, ตั้ง `sla_alerted_at` เพื่อไม่ให้แจ้งซ้ำ
- (แอป) Settings ในแอปมี 8 จอ: ทั่วไป (tenant), ธุรกิจ, ระบบอัตโนมัติ (lock + chatbot mode), AI Chatbot (engine/API key/model/ทดสอบการเชื่อมต่อ), SLA, การแจ้งเตือน, แพ็กเกจ (ดูได้อย่างเดียว), แท็ก
- การแจ้งเตือนตั้งค่าได้รายผู้ใช้: sound, push, trigger (ทั้งหมด/เฉพาะที่ได้รับมอบหมาย), muted channels
- Saved reply: CRUD + นำเข้าจาก Facebook Graph API
- (แอป) Notifications feed: ประเภท urgent/message/paid/mention/system, badge จำนวนที่ยังไม่อ่าน, อ่านทั้งหมด, แตะแล้วเปิดแชท
- Reports (`OverviewController`): first response, resolution, SLA, volume, channel mix, heatmap, performance ต่อเพจ/ต่อเจ้าหน้าที่, export CSV (UTF-8 BOM) · ในแอปแสดงเฉพาะ summary cards + channel mix (`GET mobile/overview?range=7d`)


## คำที่ใช้ในเล่ม
- **SLA ไม่ใช่จุดเด่น**: ระบบแค่แจ้งเตือนเมื่อแชทค้างตอบนานเกินกำหนด ให้เขียนว่า "แจ้งเตือนแชทที่ค้างตอบ" เป็นส่วนย่อย ไม่ตั้งเป็นวัตถุประสงค์แยก ไม่ต้องมีนิยามศัพท์ SLA และไม่ใช้คำว่า "เกณฑ์ SLA" ในบทที่ 1–2 (รายละเอียด 5/10/15/20 นาทีเขียนได้ในบทที่ 3)
- ใช้ "แชทบอท" เฉยๆ **ห้ามใช้ "แชทบอทปัญญาประดิษฐ์"** เพราะอ่านแล้วเหมือน AI เขียน
- ใช้ "แอปพลิเคชันบนสมาร์ตโฟน" / "แอปพลิเคชันสมาร์ตโฟน" แทน "อุปกรณ์เคลื่อนที่"

## ตัดออกจากเล่ม (มีในโค้ดแต่ไม่ใส่)
- แผงผู้ดูแลแพลตฟอร์ม (`admin/*`: admin login + 2FA, จัดการ/ระงับร้าน) — ผู้จัดทำตัดออก
- CSAT แบบประเมินความพึงพอใจหลังปิดแชท — มีในโค้ดแต่ผู้จัดทำไม่ใส่ในเล่ม
- CF-Shops SSO / การเชื่อมร้านจาก CF-Shops — ไม่ใส่ในเล่ม
- แพ็กเกจ / subscription / ราคา / โควตาการใช้งาน และการบันทึก token หรือค่าใช้จ่าย LLM — ไม่ใส่ในเล่ม (ในแอปมีจอตั้งค่าแพ็กเกจก็ไม่ต้องกล่าวถึง)

## ห้ามอ้างในเล่ม (ไม่มีในโค้ด)
- vector database (Pinecone/Qdrant/pgvector ฯลฯ) — ระบบใช้ cosine แบบ brute-force บน MongoDB
- ตะกร้าหลายสินค้าต่อหนึ่งคำสั่งซื้อ
- ระบบชำระเงิน / payment gateway (subscription ดูได้อย่างเดียว)
- แอป iOS ที่พร้อม release (ตั้งค่า release ไว้แค่ Android)
- ช่องทาง WhatsApp, TikTok, Shopee, Lazada (มีแค่ใน mock UI ไม่มี endpoint เชื่อมต่อ)
- Horizon, Vue, dio
