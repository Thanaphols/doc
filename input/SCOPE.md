# SCOPE — ขอบเขตงาน Chatshop (ระบบหลังบ้าน + แอปพลิเคชัน)

> ใช้เป็นแหล่งข้อมูลหลักสำหรับเขียนเล่มปริญญานิพนธ์
> แอปพลิเคชันรวมศูนย์ข้อความและการจัดการคำสั่งซื้อสำหรับธุรกิจออนไลน์ (Chatshop)
> ระบบหลังบ้าน: `C:\interm\chatshop` (Laravel) · แอปพลิเคชัน: `C:\interm\chatshop_app` (Flutter)

## ผู้พัฒนา
ผู้จัดทำพัฒนาเองคนเดียวทั้งหมด ไม่มีทีมพัฒนา:
- **ระบบหลังบ้าน**: Laravel 10 — ฐานข้อมูล, REST API, Webhook, Queue, Realtime, Push, แชทบอท
- **แอปพลิเคชันบนสมาร์ตโฟน**: Flutter 3.x, Dart — หน้าจอ 13 โมดูล, Riverpod, Reverb Realtime, FCM Push Notification, Secure Storage, Native Auth

## Stack (ตรวจจากโค้ดแล้ว)
| ส่วน | ใช้อะไร | หลักฐาน |
|---|---|---|
| Backend | Laravel 10, PHP 8.1+ | `composer.json`, `CLAUDE.md` |
| ฐานข้อมูล | MySQL (ข้อมูลเชิงสัมพันธ์) + MongoDB (`conversations`, `messages`, `knowledge_*`, `device_tokens`) | `mongodb/laravel-mongodb` |
| หน้าเว็บทั่วไป | Blade | `resources/views/` |
| หน้ากล่องข้อความ | React 18 จาก CDN + `@babel/standalone` แปลง `.jsx` ในเบราว์เซอร์ (ไม่ผ่าน Vite) | `resources/views/inbox.blade.php:330-341`, `public/js/inbox/*.jsx` |
| Vite | bundle เฉพาะ `resources/js/app.js` (Laravel Echo + pusher-js) | `package.json` |
| Realtime | Laravel Reverb (WebSocket, โพรโทคอล Pusher) | `config/broadcasting.php` |
| Auth | Session (เว็บ), Sanctum token (แอป), Socialite (Facebook/Google) | `routes/api.php` |
| Queue | คิว `high`, `default`, `typing` + scheduler | `app/Console/Kernel.php` |

จำนวน route ทั้งหมด 297 (จาก `php artisan route:list`) ในจำนวนนี้เป็น API ของแอป (`api/mobile/*`) 99 route

## ฟีเจอร์

| # | ฟีเจอร์ | ทำอะไร (สรุปจากโค้ด) | ไฟล์หลัก | สถานะ | ในเล่ม |
|---|---|---|---|---|---|
| 1 | สมัคร / เข้าสู่ระบบ | email+password, Facebook/Google (Socialite), สมัครร้านใหม่ | `AuthController`, `SocialAuthController`, `TenantController@register` | ของเรา | ใส่ |
| 2 | เชื่อมช่องทาง | LINE, Facebook, Instagram, WebChat widget | `ChannelController`, `WebChatChannelController` | ของเรา | ใส่ |
| 3 | รับ Webhook | ตรวจลายเซ็น LINE / Meta (HMAC-SHA256) แล้วส่งเข้าคิว `high` | `LineWebhookController`, `MetaWebhookController`, `ProcessLineWebhook`, `ProcessMetaWebhook` | ของเรา | ใส่ |
| 4 | กล่องข้อความ | รายการแชท, ค้นหา, ตอบ (ข้อความ/สติกเกอร์/ไฟล์), react, archive, ปักหมุด, มอบหมาย, ล็อกแชท, export | `ConversationController`, `public/js/inbox/` | ของเรา | ใส่ |
| 5 | ส่งข้อความออก | LINE Push API, Graph API v19.0, จัดการกรอบ 24 ชม. (`HUMAN_AGENT`) | `LineService`, `FacebookService`, `InstagramService` | ของเรา | ใส่ |
| 6 | Media | proxy รูป LINE, mirror ลง storage, ย่อรูป (GD→WebP) / วิดีโอ (ffmpeg) | `LineMediaController`, `MediaMirror`, `ImageCompressor`, `VideoCompressor` | ของเรา | ใส่ |
| 7 | บรอดแคสต์ | 6 กลุ่มเป้าหมาย, ตั้งเวลา, ส่งเป็นชุดละ 100 | `BroadcastController`, `BroadcastService`, `ProcessBroadcastJob`, `SendBroadcastBatchJob` | ของเรา | ใส่ |
| 8 | แท็ก | แคตตาล็อกแท็กพร้อมสี, ติด/ถอดบนแชท | `TenantTagController` | ของเรา | ใส่ |
| 9 | นำเข้าประวัติแชท | Facebook/Instagram ย้อนหลังไม่เกิน 90 วัน | `ImportController`, `ImportHistoryJob` | ของเรา | ใส่ |
| 10 | ทีมและบทบาท | 5 บทบาท + custom role, 12 สิทธิ์, คำเชิญ 7 วัน, โอนเจ้าของ | `TeamController`, `TeamRoleController`, `InvitationController`, `TenantRole` | ของเรา | ใส่ |
| 11 | ตอบอัตโนมัติด้วยคีย์เวิร์ด | กฎ keyword (ตรงทั้งคำ / มีคำนี้) → ตอบข้อความที่กำหนด | `AutomationRuleController`, `ProcessAutomationReply`, `AutomationRuleMatcher` | ของเรา | ใส่ |
| 12 | แชทบอท | 3 โหมด, รวมข้อความ, persona/directives, wizard + playground, ส่งต่อเจ้าหน้าที่, สรุปแชท, อ่านรูป | `ChatbotService`, `ProcessChatbotReply`, `ChatbotWizardController`, `ChatbotDirectiveController` | ของเรา | ใส่ |
| 13 | ผู้ให้บริการ LLM | Ollama, Claude, Gemini, OpenAI-compatible (custom provider + ทดสอบการเชื่อมต่อ) | `ChatbotSettingsController`, `CustomProviderController`, `ChatbotKeyController` | ของเรา | ใส่ |
| 14 | ฐานความรู้ (RAG) | เอกสาร FAQ → chunk ≤ 500 ตัวอักษร → embedding → cosine (≥ 0.3, 4 อันดับแรก) | `KnowledgeController`, `Chatbot/Knowledge/*` | ของเรา | ใส่ |
| 15 | สินค้า / ลูกค้า | CRUD สินค้า, ข้อมูลลูกค้า (เบอร์, ที่อยู่, โน้ต) | `ProductController`, `CustomerController` | ของเรา | ใส่ |
| 16 | คำสั่งซื้อ | เปิดจากแชท, สินค้า 1 ชนิด × จำนวน, 4 สถานะ | `ChatOrderController`, `ChatOrder` | ของเรา | ใส่ |
| 17 | SLA | เกณฑ์ 5/10/15/20 นาที, สแกนทุก 1 นาที, แจ้งเตือน Reverb + FCM | `ScanSlaBreaches`, `Api/SettingsController@updateSla` | ของเรา | ใส่ |
| 18 | CSAT ความพึงพอใจ | ปิดแชทแล้วส่งแบบประเมิน 1–5 คะแนนให้ลูกค้า | `CsatService`, `CsatSettingsController` | ของเรา | ไม่ใส่ |
| 19 | ตั้งค่าร้าน / ธุรกิจ | ข้อมูลร้าน, หลายธุรกิจต่อร้าน, ผูกเพจกับธุรกิจ, ข้อความตอบกลับสำเร็จรูป, การแจ้งเตือน | `TenantSettingsController`, `BusinessController`, `SavedReplyController`, `NotificationSettingsController` | ของเรา | ใส่ |
| 20 | รายงาน | first response, resolution, SLA, ปริมาณ, สัดส่วนช่องทาง, heatmap, ต่อเพจ/ต่อคน, export CSV | `OverviewController` | ของเรา | ใส่ |
| 21 | Realtime | event 11+ ชนิดผ่านช่อง `private-tenant.{id}` | `app/Events/*` | ของเรา | ใส่ |
| 22 | REST API สำหรับแอป | 99 route ใต้ `api/mobile/*` + Sanctum | `routes/api.php`, `Api/*Controller` | ของเรา | ใส่ |
| 23 | ส่ง push (FCM) | ส่งแจ้งเตือนไปยังเครื่องที่ลงทะเบียน | `PushService`, `mobile/devices` | ของเรา | ใส่ |
| 24 | แผงผู้ดูแลแพลตฟอร์ม | admin login + 2FA, จัดการ/ระงับร้าน, dashboard, รีวิว/รายงานปัญหา | `Admin/*Controller` | ของเรา | ไม่ใส่ (เกี่ยวกับเงิน) |
| 25 | แพ็กเกจ / subscription | ดูแพ็กเกจ, admin กำหนดแพ็กเกจให้ร้าน, บันทึก token/ค่าใช้จ่าย LLM | `SubscriptionController`, `Admin\TenantController@assignPlan`, `ChatbotUsageController` | ของเรา | ไม่ใส่ |
| 26 | CF-Shops SSO | เข้าสู่ระบบ/เชื่อมร้านจาก CF-Shops | `CfShopsConnectController` | ของเรา | ไม่ใส่ |

## คำตอบของผู้จัดทำ (2026-09-29)
1. ข้อ 24 แผงผู้ดูแลแพลตฟอร์ม: ตัดออก ฟังก์ชันน้อยและส่วนใหญ่เกี่ยวกับเงิน
2. ข้อ 11 ตอบอัตโนมัติด้วยคีย์เวิร์ด: ใส่ในเล่ม / ข้อ 18 CSAT: ไม่ใส่
3. ทุกฟีเจอร์เป็นงานของผู้จัดทำทั้งหมด ไม่มีผู้ร่วมพัฒนา (ยืนยัน 2026-09-30)
4. งานที่ไม่อยู่ในตาราง: ไม่มี ครบแล้ว
5. ไฟล์ที่แก้ค้างใน branch `eark`: ยังไม่ได้ตอบ
