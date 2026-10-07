// Generates ch3 diagrams (3.1-3.7) as SVG -> PNG via headless Chrome.
// Run: node input/figures/ch3/src/gen.js
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const OUT = path.join(__dirname, "..");
const CHROME = [
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
].find((p) => fs.existsSync(p));

const T = {
  head: { fill: "#1e3a8a", stroke: "#1e3a8a", color: "#ffffff" },
  app: { fill: "#e0f2fe", stroke: "#0284c7", color: "#075985" },
  be: { fill: "#eef2ff", stroke: "#1e3a8a", color: "#1e3a8a" },
  db: { fill: "#d1fae5", stroke: "#059669", color: "#065f46" },
  ext: { fill: "#e2e8f0", stroke: "#475569", color: "#1e293b" },
  warn: { fill: "#fee2e2", stroke: "#dc2626", color: "#991b1b" },
  dec: { fill: "#fef3c7", stroke: "#d97706", color: "#92400e" },
  pill: { fill: "#1e3a8a", stroke: "#1e3a8a", color: "#ffffff" },
};
const LINE = "#475569";

function fig(W, H) {
  const els = [];
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const api = {
    raw: (s) => els.push(s),
    text(x, y, lines, { size = 24, weight = 400, color = "#1e293b", anchor = "middle", halo = false } = {}) {
      lines = [].concat(lines);
      const lh = size * 1.35;
      const y0 = y - ((lines.length - 1) * lh) / 2 + size * 0.35;
      const h = halo ? ` stroke="#fff" stroke-width="7" paint-order="stroke"` : "";
      lines.forEach((l, i) => {
        const w = typeof l === "object" ? l.weight : weight;
        const s = typeof l === "object" ? l.size || size : size;
        const t = typeof l === "object" ? l.t : l;
        els.push(`<text x="${x}" y="${y0 + i * lh}" font-size="${s}" font-weight="${w}" fill="${color}" text-anchor="${anchor}"${h}>${esc(t)}</text>`);
      });
    },
    // rounded box with centered text; returns anchor points
    box(x, y, w, h, lines, th = T.be, { r = 14, size = 24, weight = 600, sw = 2.5, dash } = {}) {
      const d = dash ? ` stroke-dasharray="10 7"` : "";
      els.push(`<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${r}" fill="${th.fill}" stroke="${th.stroke}" stroke-width="${sw}"${d}/>`);
      if (lines) api.text(x + w / 2, y + h / 2, lines, { size, weight, color: th.color });
      return { x, y, w, h, cx: x + w / 2, cy: y + h / 2, l: x, r: x + w, t: y, b: y + h };
    },
    // container with a header bar
    group(x, y, w, h, title, th, { size = 26 } = {}) {
      els.push(`<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="18" fill="#f8fafc" stroke="${th.stroke}" stroke-width="3"/>`);
      els.push(`<rect x="${x + 14}" y="${y + 14}" width="${w - 28}" height="56" rx="12" fill="${th.stroke}"/>`);
      api.text(x + w / 2, y + 42, title, { size, weight: 700, color: "#fff" });
    },
    diamond(cx, cy, w, h, lines, th = T.dec, size = 22) {
      els.push(`<polygon points="${cx},${cy - h / 2} ${cx + w / 2},${cy} ${cx},${cy + h / 2} ${cx - w / 2},${cy}" fill="${th.fill}" stroke="${th.stroke}" stroke-width="2.5"/>`);
      api.text(cx, cy, lines, { size, weight: 600, color: th.color });
      return { cx, cy, t: cy - h / 2, b: cy + h / 2, l: cx - w / 2, r: cx + w / 2 };
    },
    ellipse(cx, cy, rx, ry, label, th = T.app, size = 22) {
      els.push(`<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" fill="${th.fill}" stroke="${th.stroke}" stroke-width="2.5"/>`);
      api.text(cx, cy, label, { size, weight: 600, color: th.color });
      return { cx, cy, l: cx - rx, r: cx + rx, t: cy - ry, b: cy + ry };
    },
    actor(cx, cy, label) {
      const s = `stroke="#1e293b" stroke-width="3" fill="none"`;
      els.push(`<circle cx="${cx}" cy="${cy - 70}" r="22" ${s}/>`);
      els.push(`<line x1="${cx}" y1="${cy - 48}" x2="${cx}" y2="${cy}" ${s}/>`);
      els.push(`<line x1="${cx - 40}" y1="${cy - 32}" x2="${cx + 40}" y2="${cy - 32}" ${s}/>`);
      els.push(`<line x1="${cx}" y1="${cy}" x2="${cx - 32}" y2="${cy + 50}" ${s}/>`);
      els.push(`<line x1="${cx}" y1="${cy}" x2="${cx + 32}" y2="${cy + 50}" ${s}/>`);
      api.text(cx, cy + 82, label, { size: 24, weight: 700 });
    },
    // polyline arrow; opts: label, at (label point), both, dash, head
    arrow(pts, { label, at, both, dash, head = true, color = LINE, size = 20 } = {}) {
      const p = pts.map((q) => q.join(",")).join(" ");
      const d = dash ? ` stroke-dasharray="10 7"` : "";
      const m = head ? ` marker-end="url(#ah)"` : "";
      const ms = both ? ` marker-start="url(#ahs)"` : "";
      els.push(`<polyline points="${p}" fill="none" stroke="${color}" stroke-width="2.5"${d}${m}${ms}/>`);
      if (label) {
        const [lx, ly] = at || [(pts[0][0] + pts[1][0]) / 2, (pts[0][1] + pts[1][1]) / 2 - 16];
        api.text(lx, ly, label, { size, weight: 600, color: "#334155", halo: true });
      }
    },
    svg() {
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="Leelawadee UI, Tahoma, sans-serif">
<defs>
<marker id="ah" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="9" markerHeight="9" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="${LINE}"/></marker>
<marker id="ahs" viewBox="0 0 10 10" refX="1" refY="5" markerWidth="9" markerHeight="9" orient="auto"><path d="M10,0 L0,5 L10,10 z" fill="${LINE}"/></marker>
<marker id="gen" viewBox="0 0 20 20" refX="19" refY="10" markerWidth="16" markerHeight="16" orient="auto"><path d="M0,0 L20,10 L0,20 z" fill="#fff" stroke="#1e293b" stroke-width="2"/></marker>
</defs>
<rect width="${W}" height="${H}" fill="#fff"/>
${els.join("\n")}
</svg>`;
    },
    W, H,
  };
  return api;
}

const figs = {};

// 3.1 Use Case Diagram
figs["3.1"] = () => {
  const f = fig(1600, 1060);
  f.raw(`<rect x="420" y="30" width="800" height="1000" rx="20" fill="#f8fafc" stroke="#1e3a8a" stroke-width="3"/>`);
  f.text(820, 75, "ระบบ Chatshop", { size: 30, weight: 700, color: "#1e3a8a" });
  const owner = ["จัดการช่องทาง", "จัดการทีมและบทบาท", "ตั้งค่าร้านค้า", "ตั้งค่าแชทบอท"].map((t, i) => f.ellipse(660, 170 + i * 95, 180, 38, t));
  const staff = ["ตอบแชท", "จัดการบทสนทนา", "จัดการลูกค้า", "จัดการสินค้า", "จัดการคำสั่งซื้อ"].map((t, i) => f.ellipse(660, 580 + i * 95, 180, 38, t, T.be));
  f.ellipse(1020, 540, 170, 38, "ส่งข้อความ", T.db);
  f.ellipse(1020, 670, 170, 52, ["รับคำตอบจาก", "เจ้าหน้าที่หรือแชทบอท"], T.db, 21);
  f.actor(200, 300, "เจ้าของร้าน");
  f.actor(200, 760, "เจ้าหน้าที่");
  f.actor(1420, 600, "ลูกค้า");
  owner.forEach((e) => f.arrow([[245, 268], [e.l, e.cy]], { head: false }));
  staff.forEach((e) => f.arrow([[245, 728], [e.l, e.cy]], { head: false }));
  f.arrow([[1375, 568], [1190, 540]], { head: false });
  f.arrow([[1375, 568], [1190, 670]], { head: false });
  // generalization: owner inherits staff
  f.raw(`<line x1="200" y1="410" x2="200" y2="664" stroke="#1e293b" stroke-width="2.5" marker-end="url(#gen)"/>`);
  f.text(215, 540, ["ใช้ทุกฟังก์ชัน", "ของเจ้าหน้าที่ได้"], { size: 20, color: "#475569", anchor: "start" });
  return f;
};

// 3.2 System architecture
figs["3.2"] = () => {
  const f = fig(1720, 1000);
  // app
  f.group(40, 60, 320, 500, ["แอปพลิเคชัน", "บนสมาร์ตโฟน"], T.app, { size: 22 });
  f.box(70, 150, 260, 70, "ตัวเรียก REST API", T.app, { size: 22 });
  f.box(70, 290, 260, 70, "บริการเรียลไทม์", T.app, { size: 22 });
  f.box(70, 430, 260, 70, "บริการแจ้งเตือน", T.app, { size: 22 });
  f.text(200, 530, "Flutter", { size: 22, weight: 700, color: "#0284c7" });
  // backend
  f.group(440, 60, 640, 500, "ระบบบริการหลังบ้าน (Laravel 10)", T.head);
  const api = f.box(470, 130, 280, 110, ["REST API", { t: "Laravel Sanctum", weight: 400, size: 20 }]);
  const wh = f.box(780, 130, 280, 110, ["รับ Webhook / ส่งข้อความ", { t: "HMAC-SHA256", weight: 400, size: 20 }], T.be, { size: 22 });
  const rv = f.box(470, 270, 280, 110, ["Laravel Reverb", { t: "WebSocket", weight: 400, size: 20 }]);
  f.box(780, 270, 280, 110, ["คิวงานและ", "ตัวจัดตารางเวลา"], T.be, { size: 22 });
  f.box(470, 410, 280, 110, ["จัดการไฟล์สื่อ", { t: "PHP GD / ffmpeg", weight: 400, size: 20 }]);
  const bot = f.box(780, 410, 280, 110, "แชทบอท");
  // database
  f.group(440, 660, 640, 230, "ฐานข้อมูล", T.db);
  f.box(470, 750, 280, 110, ["MySQL", { t: "ข้อมูลเชิงสัมพันธ์", weight: 400, size: 20 }], T.db);
  f.box(780, 750, 280, 110, ["MongoDB", { t: "บทสนทนา ข้อความ ฐานความรู้", weight: 400, size: 19 }], T.db);
  // external
  f.group(1160, 60, 320, 830, "บริการภายนอก", T.ext);
  const line = f.box(1185, 130, 270, 90, "LINE Messaging API", T.ext, { size: 22 });
  const meta = f.box(1185, 250, 270, 110, ["Meta Graph API", { t: "Messenger / Instagram", weight: 400, size: 20 }], T.ext, { size: 22 });
  const llm = f.box(1185, 410, 270, 110, ["บริการแบบจำลองภาษา", { t: "Ollama / Claude / Gemini", weight: 400, size: 18 }], T.ext, { size: 22 });
  const fcm = f.box(1185, 750, 270, 110, ["Firebase Cloud", "Messaging"], T.ext, { size: 22 });
  f.actor(1620, 260, "ลูกค้า");
  // links
  f.arrow([[330, 185], [470, 185]], { both: true, label: "REST API", at: [400, 160] });
  f.arrow([[470, 325], [330, 325]], { label: "WebSocket", at: [400, 300] });
  f.arrow([[1185, 165], [1060, 165]], { both: true });
  f.arrow([[1185, 305], [1120, 305], [1120, 210], [1060, 210]], { both: true });
  f.text(1320, 236, "Webhook / ส่งข้อความ", { size: 18, weight: 600, color: "#334155", halo: true });
  f.arrow([[1060, 465], [1185, 465]], { both: true });
  f.arrow([[760, 560], [760, 660]], { both: true, label: "อ่าน / บันทึกข้อมูล", at: [870, 610] });
  f.arrow([[1080, 530], [1120, 530], [1120, 805], [1185, 805]]);
  f.arrow([[1320, 860], [1320, 950], [200, 950], [200, 560]], { label: "Push Notification", at: [760, 950] });
  f.arrow([[1575, 230], [1455, 175]], { both: true });
  f.arrow([[1575, 260], [1455, 305]], { both: true });
  return f;
};

// 3.3 Data relationships (MySQL + MongoDB)
figs["3.3"] = () => {
  const f = fig(1860, 1000);
  f.group(20, 20, 1210, 900, "MySQL", T.db);
  f.group(1260, 20, 580, 900, "MongoDB", T.app);
  const E = (x, y, name, th = T.db, w = 200) => f.box(x, y, w, 56, name, th, { size: 20, weight: 700, r: 8 });
  const card = (x, y, s, anchor = "middle") => f.text(x, y, s, { size: 20, weight: 700, color: "#b45309", anchor, halo: true });
  const tenants = E(500, 110, "tenants", T.db, 220);
  const row2 = [
    E(40, 270, "businesses"), E(290, 270, "tenant_users"), E(520, 270, "tenant_roles"),
    E(750, 270, "tenant_tags"), E(980, 270, "customers"),
  ];
  row2.forEach((b) => { f.arrow([[tenants.cx, tenants.b], [b.cx, b.t]], { head: false }); card(b.cx - 16, b.t - 12, "N"); });
  card(tenants.cx + 18, tenants.b + 20, "1");
  // businesses children (tree)
  const biz = row2[0];
  const kids = [
    ["business_chatbot_settings", "1"], ["platform_credentials", "N"],
    ["products", "N"], ["automation_rules", "N"],
  ].map(([n, c], i) => {
    const k = E(200, 440 + i * 110, n, T.db, 280);
    f.arrow([[140, biz.b], [140, k.cy], [k.l, k.cy]], { head: false });
    card(k.l - 14, k.cy - 12, c);
    return k;
  });
  card(158, biz.b + 20, "1");
  // customers -> chat_orders
  const cust = row2[4];
  const orders = E(980, 480, "chat_orders");
  f.arrow([[cust.cx, cust.b], [orders.cx, orders.t]], { head: false });
  card(cust.cx + 16, cust.b + 20, "1"); card(orders.cx + 16, orders.t - 12, "N");
  // assignments -> tenant_users
  const asg = E(600, 620, "conversation_assignments", T.db, 290);
  const tu = row2[1];
  f.arrow([[tu.cx, tu.b], [asg.cx, asg.t]], { head: false });
  card(tu.cx + 18, tu.b + 22, "1"); card(asg.cx - 20, asg.t - 14, "N");
  // MongoDB
  const conv = E(1300, 270, "conversations", T.app, 240);
  const msg = E(1300, 480, "messages", T.app, 240);
  const kd = E(1570, 270, "knowledge_documents", T.app, 240);
  const kc = E(1570, 480, "knowledge_chunks", T.app, 240);
  E(1570, 690, "device_tokens", T.app, 240);
  f.arrow([[conv.cx, conv.b], [msg.cx, msg.t]], { head: false });
  card(conv.cx + 16, conv.b + 20, "1"); card(msg.cx + 16, msg.t - 12, "N");
  f.arrow([[kd.cx, kd.b], [kc.cx, kc.t]], { head: false });
  card(kd.cx + 16, kd.b + 20, "1"); card(kc.cx + 16, kc.t - 12, "N");
  // cross-database references (dashed)
  const X = "#d97706";
  f.arrow([[cust.r, 285], [conv.l, 285]], { head: false, dash: true, color: X });
  card(cust.r + 14, 272, "1", "start"); card(conv.l - 14, 272, "N", "end");
  f.arrow([[orders.r, orders.cy], [1240, orders.cy], [1240, 303], [conv.l, 303]], { head: false, dash: true, color: X });
  f.arrow([[asg.r, asg.cy], [1250, asg.cy], [1250, 316], [conv.l, 316]], { head: false, dash: true, color: X });
  // legend
  f.raw(`<line x1="40" y1="960" x2="110" y2="960" stroke="${LINE}" stroke-width="2.5"/>`);
  f.text(125, 960, "ความสัมพันธ์ภายในฐานข้อมูลเดียวกัน", { size: 21, anchor: "start" });
  f.raw(`<line x1="640" y1="960" x2="710" y2="960" stroke="${X}" stroke-width="2.5" stroke-dasharray="10 7"/>`);
  f.text(725, 960, "อ้างอิงข้ามฐานข้อมูลด้วยรหัส (ระบบหลังบ้านตรวจความถูกต้องเอง)", { size: 21, anchor: "start" });
  f.text(1820, 960, "1 / N = จำนวนความสัมพันธ์", { size: 21, anchor: "end", color: "#b45309" });
  return f;
};

// 3.4 Webhook receive flow
figs["3.4"] = () => {
  const f = fig(1640, 900);
  f.group(30, 30, 760, 840, "ระบบหลังบ้าน (ตัวรับ Webhook)", T.head);
  f.group(850, 30, 760, 840, "งานในคิว high", T.be);
  const s = f.box(270, 120, 400, 70, "แพลตฟอร์มส่ง Webhook", T.pill, { r: 35 });
  const d = f.diamond(470, 300, 260, 140, ["ลายเซ็น", "ถูกต้อง?"]);
  const rej = f.box(50, 265, 190, 70, "ปฏิเสธคำขอ", T.warn, { size: 22 });
  const q = f.box(270, 440, 400, 80, "ส่งงานเข้าคิว high", T.be);
  const ok = f.box(270, 580, 400, 80, "ตอบกลับแพลตฟอร์มทันที", T.be);
  const e1 = f.box(370, 730, 200, 64, "สิ้นสุด", T.pill, { r: 32 });
  f.arrow([[470, s.b], [470, d.t]]);
  f.arrow([[d.l, 300], [rej.r, 300]], { label: "ไม่ใช่", at: [290, 278] });
  f.arrow([[470, d.b], [470, q.t]], { label: "ใช่", at: [500, 405] });
  f.arrow([[470, q.b], [470, ok.t]]);
  f.arrow([[470, ok.b], [470, e1.t]]);
  const steps = [
    "บันทึกข้อมูลลูกค้าลง MySQL",
    "บันทึกบทสนทนาและข้อความลง MongoDB",
    "คัดลอกไฟล์สื่อมาเก็บในระบบ",
    "ส่งเหตุการณ์ไปยังแอปพลิเคชันผ่าน Laravel Reverb",
    "ส่ง Push Notification ผ่าน Firebase Cloud Messaging",
    "ส่งต่อให้กฎตอบกลับอัตโนมัติและแชทบอท",
  ].map((t, i) => f.box(900, 120 + i * 105, 660, 70, t, T.be, { size: 22 }));
  steps.slice(1).forEach((b, i) => f.arrow([[1230, steps[i].b], [1230, b.t]]));
  const e2 = f.box(1130, 760, 200, 64, "สิ้นสุด", T.pill, { r: 32 });
  f.arrow([[1230, steps[5].b], [1230, e2.t]]);
  f.arrow([[q.r, 480], [820, 480], [820, 155], [900, 155]], { dash: true, label: "งานในคิว", at: [820, 330] });
  return f;
};

// 3.5 Chatbot reply flow
figs["3.5"] = () => {
  const f = fig(1420, 1100);
  const X = 520;
  const s = f.box(X - 330, 40, 660, 80, ["ข้อความใหม่จากลูกค้า", { t: "บทสนทนาอยู่ในโหมด Suggest หรือ Autopilot", weight: 400, size: 20 }], T.pill, { r: 40, size: 24 });
  const b1 = f.box(X - 330, 170, 660, 76, "รอ 6 วินาทีเพื่อรวมข้อความที่ลูกค้าส่งติดกัน");
  const b2 = f.box(X - 330, 296, 660, 76, "ค้นคืนส่วนย่อยของเอกสารความรู้ที่เกี่ยวข้อง");
  const b3 = f.box(X - 330, 422, 660, 110, ["ประกอบข้อความสั่งงาน", { t: "บุคลิกของร้าน · ข้อมูลแบรนด์ · บทสรุป · ข้อความล่าสุด 20 ข้อความ", weight: 400, size: 20 }]);
  const b4 = f.box(X - 330, 582, 660, 76, "ส่งให้แบบจำลองภาษาสร้างคำตอบ");
  const d = f.diamond(X, 790, 380, 160, ["แบบจำลองตอบว่า", "ต้องส่งต่อ?"]);
  const y = f.box(920, 735, 440, 110, ["ติดแท็กรอแอดมิน", "และเปลี่ยนเป็นโหมด Manual"], T.warn);
  const n = f.box(X - 330, 930, 660, 110, ["Autopilot: ส่งคำตอบถึงลูกค้า", "Suggest: แสดงคำตอบแนะนำให้เจ้าหน้าที่"], T.db);
  [[s, b1], [b1, b2], [b2, b3], [b3, b4]].forEach(([a, b]) => f.arrow([[X, a.b], [X, b.t]]));
  f.arrow([[X, b4.b], [X, d.t]]);
  f.arrow([[d.r, 790], [y.l, 790]], { label: "ใช่", at: [810, 768] });
  f.arrow([[X, d.b], [X, n.t]], { label: "ไม่ใช่", at: [X + 45, 895] });
  return f;
};

// 3.6 Realtime events and notifications
figs["3.6"] = () => {
  const f = fig(1700, 810);
  const row = ["ผู้ใช้เข้าสู่ระบบ", ["ขออนุญาตแสดง", "การแจ้งเตือน"], ["ลงทะเบียนโทเคนของ", "อุปกรณ์กับระบบหลังบ้าน"], ["เชื่อมต่อช่องส่วนตัวของ", "ร้านค้าผ่าน Laravel Reverb"]]
    .map((t, i) => f.box(40 + i * 415, 40, 370, 100, t, i ? T.be : T.pill, { size: 22, r: i ? 14 : 50 }));
  row.slice(1).forEach((b, i) => f.arrow([[row[i].r, 90], [b.l, 90]]));
  const d = f.diamond(850, 290, 340, 130, ["แอปพลิเคชัน", "เปิดอยู่?"]);
  f.arrow([[row[3].cx, row[3].b], [row[3].cx, 190], [850, 190], [850, d.t]]);
  // open -> websocket events
  const ws = f.box(330, 400, 460, 80, "รับเหตุการณ์ผ่าน WebSocket", T.app);
  f.arrow([[d.l, 290], [560, 290], [560, ws.t]], { label: "ใช่", at: [620, 270] });
  const ev = [
    ["MessageReceived", "เพิ่มข้อความในหน้าต่าง", "สนทนาและปรับกล่องข้อความ"],
    ["ConversationUpdated", "ปรับผู้รับผิดชอบ แท็ก", "และโหมดแชทบอท"],
    ["NotificationCreated", "ดึงรายการแจ้งเตือนใหม่"],
    ["OrderUpdated", "โหลดคำสั่งซื้อใหม่"],
    ["SettingsChanged", "โหลดการตั้งค่า", "ร้านค้าใหม่"],
  ].map(([h, ...ls], i) => f.box(30 + i * 222, 600, 210, 170, [{ t: h, weight: 700, size: 18 }, ...ls.map((t) => ({ t, weight: 400, size: 17 }))], T.app, { size: 18 }));
  f.raw(`<line x1="${ev[0].cx}" y1="545" x2="${ev[4].cx}" y2="545" stroke="${LINE}" stroke-width="2.5"/>`);
  f.arrow([[560, ws.b], [560, 545]], { head: false });
  ev.forEach((e) => f.arrow([[e.cx, 545], [e.cx, e.t]]));
  // closed -> push notification
  const p1 = f.box(1200, 400, 460, 80, ["รับ Push Notification จาก", "Firebase Cloud Messaging"], T.ext, { size: 21 });
  const p2 = f.box(1200, 540, 460, 70, "ผู้ใช้แตะการแจ้งเตือน", T.ext, { size: 22 });
  const p3 = f.box(1200, 670, 460, 100, ["เปิดหน้าต่างสนทนา หรือหน้ารายการ", "คำสั่งซื้อเมื่อเป็นการแจ้งชำระเงิน"], T.ext, { size: 21 });
  f.arrow([[d.r, 290], [1430, 290], [1430, p1.t]], { label: "ไม่ใช่", at: [1100, 270] });
  f.arrow([[1430, p1.b], [1430, p2.t]]);
  f.arrow([[1430, p2.b], [1430, p3.t]]);
  return f;
};

// 3.7 App code structure
figs["3.7"] = () => {
  const f = fig(1720, 1000);
  f.group(40, 30, 1540, 350, "features — หน้าจอและตรรกะของแต่ละโมดูล (13 โมดูล)", T.app);
  const mods = ["การยืนยันตัวตน", "กล่องข้อความ", "หน้าต่างสนทนา", "การค้นหา", "การแจ้งเตือน", "ลูกค้า", "สินค้า",
    "คำสั่งซื้อ", "ทีม", "รายงาน", "ช่องทาง", "ข้อมูลผู้ใช้", "การตั้งค่า"];
  mods.forEach((m, i) => {
    const r = i < 7 ? 0 : 1, c = i < 7 ? i : i - 7;
    const x0 = r ? 70 + 105 : 70;
    f.box(x0 + c * 210, 125 + r * 120, 190, 80, m, T.app, { size: 22 });
  });
  f.group(40, 420, 1000, 250, "data — แบบจำลองข้อมูลและ Repository", T.db);
  f.box(70, 515, 450, 125, ["แบบจำลองข้อมูล", { t: "บทสนทนา ข้อความ การแจ้งเตือน ฯลฯ", weight: 400, size: 20 }], T.db);
  f.box(560, 515, 450, 125, ["Repository", { t: "แปลง JSON จากระบบหลังบ้าน", weight: 400, size: 20 }, { t: "เป็นแบบจำลองข้อมูล", weight: 400, size: 20 }], T.db);
  f.group(1080, 420, 500, 250, "shared — วิดเจ็ตที่ใช้ร่วมกัน", T.be);
  f.box(1110, 515, 440, 55, "แถบนำทางด้านล่าง", T.be, { size: 21 });
  f.box(1110, 585, 440, 55, "แผ่นฟอร์มเลื่อนขึ้นจากด้านล่าง", T.be, { size: 21 });
  f.group(40, 710, 1540, 260, "core — บริการพื้นฐานที่ทุกหน้าจอใช้ร่วมกัน", T.ext);
  [
    ["ส่วนเก็บโทเคน", "flutter_secure_storage"],
    ["ตัวเรียก REST API", "แนบโทเคน / multipart"],
    ["บริการเรียลไทม์", "Laravel Reverb"],
    ["บริการแจ้งเตือน", "Firebase Cloud Messaging"],
    ["go_router", "กำหนดเส้นทางหน้าจอ"],
  ].forEach(([h, s], i) => f.box(70 + i * 300, 805, 280, 130, [h, { t: s, weight: 400, size: 19 }], T.ext, { size: 22 }));
  // Riverpod side bar
  f.raw(`<rect x="1610" y="30" width="80" height="940" rx="16" fill="#1e3a8a"/>`);
  f.raw(`<text x="1650" y="500" font-size="26" font-weight="700" fill="#fff" text-anchor="middle" transform="rotate(-90 1650 500)">Riverpod — จัดการสถานะของทุกส่วน</text>`);
  return f;
};

const only = process.argv.slice(2);
for (const [id, make] of Object.entries(figs)) {
  if (only.length && !only.includes(id)) continue;
  const f = make();
  const html = path.join(__dirname, `${id}.html`);
  fs.writeFileSync(path.join(__dirname, `${id}.svg`), f.svg());
  fs.writeFileSync(html, `<!doctype html><meta charset="utf-8"><style>html,body{margin:0;overflow:hidden}</style>${f.svg()}`);
  execFileSync(CHROME, [
    "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=2",
    `--window-size=${f.W},${f.H}`, `--screenshot=${path.join(OUT, `${id}.png`)}`, "file:///" + html.replace(/\\/g, "/"),
  ], { stdio: "ignore" });
  fs.unlinkSync(html);
  console.log("wrote", `${id}.png`);
}
