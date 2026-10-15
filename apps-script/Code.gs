/**
 * Balance Waste — ระบบบริหารจัดการของเสีย (Waste Management)
 * บริษัท เทพวัฒนา จำกัด
 * Backend for Google Apps Script (data in Google Sheets).
 *
 * Setup (once): run setupSystem() from the Apps Script editor. It creates the
 * "Balance Waste Database" spreadsheet (one sheet per table, header row on
 * top; a NEW database is seeded with the 40 permit lines and 8 receivers of
 * สก.2 2569-12560, see seedThepwatanaPermit) and stores the spreadsheet ID in Script Properties
 * (WASTE_SPREADSHEET_ID). Apps Script project name: "Balance Waste Thepwatana".
 * Then deploy as a web app (Execute as: me, Who has access: Anyone).
 *
 * Open site: anyone may read (dashboard, balance, history, export) and record
 * store / opening / ship without a password. The name typed in the top-right
 * of the page is sent as "actor" and recorded on every movement and audit_logs
 * row ("ผู้บันทึก"); writes need it.
 *
 * Password levels (the "ข้อมูลภายใน" menu and the admin tools). The passwords
 * are Script Properties only (never in the source):
 *   WASTE_VIEW_PASSWORD   viewer  "ดูได้อย่างเดียว": may see the in-house rows
 *                         (TR- list, details, CSV) but not record them
 *   WASTE_STAFF_PASSWORD  staff   "พนักงานกรอกข้อมูล": viewer + record "treat"
 *   WASTE_RESET_PASSWORD  admin   "ผู้ดูแลระบบ": everything (also edit / delete
 *                         any movement, waste types, audit log, reset_data)
 * An unset property means that level cannot unlock. login compares the password
 * with all three (SHA-256 digests, constant time, highest role wins) and issues
 * a random token kept in CacheService (SESSION_TTL = 12 h absolute; the cache
 * entry is renewed while in use because CacheService keeps a value <= 6 h).
 * Changing / removing a password ends the sessions of that level. Wrong
 * passwords: 1 s delay each; lockout per client (5 / 15 min) and global
 * (30 / 15 min). The server checks the level on EVERY action (ACTION_ROLE);
 * a public action ignores a missing / unknown token (it is simply anonymous).
 * reset_data also needs the admin password typed again (same lockout).
 *
 * Ledger integrity: every write that changes movements (ship, edit, delete,
 * also store/opening) is checked by replaying ALL movements in time order
 * (validateLedger_): the weight of every waste type and the container count of
 * every (waste type, container) pair must never go below zero at any moment.
 * Nothing is written unless the replay passes.
 *
 * Units: the sheet stores kilograms (weight_kg, 0.01 kg precision; existing
 * rows are kg). The page sends weight_t (tons, up to 5 decimals = 0.01 kg;
 * the page shows 2) and the server converts exactly: kg = round(t * 100000) / 100.
 * API numbers stay *_kg; every text the server writes (messages, audit log,
 * CSV) shows tons.
 *
 * In-house data (kind "treat", ข้อมูลภายใน; shown as "บำบัดเอง" in totals):
 * stock handled inside the plant (e.g. the plant's own wastewater system). It
 * leaves the in-plant balance exactly like a shipment (same replay / container
 * checks) but is NOT a shipment: the สก.2 quota (quota %, receiver quotas)
 * counts "ship" only. Aggregate totals (per-type treated_kg, dashboard KPIs and
 * the monthly series) are sent only to a request with a valid token of any
 * level (aggOn_; hideTreat does not matter); anonymous requests get none.
 * The individual TR- rows (date, method, note, recorder: history, CSV, the
 * time of the last movement) are sent only to a request with a valid session
 * of any level that does not ask to hide them (hideTreat). Recording "treat"
 * needs staff. Every row is still in the movements sheet and in audit_logs.
 *
 * Optional: run installKeepWarmTrigger() once from the editor; a 10-minute
 * trigger then runs keepWarm() (read-only) to keep the read cache hot.
 *
 * Transport (no CORS preflight): POST text/plain JSON
 *   {action, token?, actor, ...fields} -> {ok:true,data} | {ok:false,error,code}
 *   code AUTH_REQUIRED (no / expired session) or FORBIDDEN (level too low).
 *   GET ?action=<public read> also works but never uses a token (anonymous);
 *   anything carrying a token is POST only, so a token never sits in a URL.
 */

var SPREADSHEET_ID_PROPERTY = "WASTE_SPREADSHEET_ID";
var SPREADSHEET_NAME = "Balance Waste Database";

// Access levels. Passwords: Script Properties only (see the header comment).
var RESET_ACTION = "reset_data";
var RESET_PASSWORD_PROPERTY = "WASTE_RESET_PASSWORD";
var ROLE_PROPERTY = { viewer: "WASTE_VIEW_PASSWORD", staff: "WASTE_STAFF_PASSWORD", admin: RESET_PASSWORD_PROPERTY };
var ROLE_LEVEL = { viewer: 1, staff: 2, admin: 3 };
var ROLE_NAME = { viewer: "ดูได้อย่างเดียว", staff: "พนักงานกรอกข้อมูล", admin: "ผู้ดูแลระบบ" };
var ROLE_ORDER = ["admin", "staff", "viewer"]; // the same password set for two levels: the higher one wins
// The lowest level allowed to run each action (checked on every request).
// "public": no password (a token, if any, only widens what reads may show).
var PUBLIC = "public";
var ACTION_ROLE = {
  bootstrap: PUBLIC, dashboard: PUBLIC, history: PUBLIC, "export": PUBLIC, batch: PUBLIC,
  store: PUBLIC, opening: PUBLIC, ship: PUBLIC,
  treat: "staff",
  movement_update: "admin", movement_delete: "admin", type_save: "admin", receiver_save: "admin", receiver_delete: "admin", type_delete: "admin",
  reset_data: "admin", logs: "admin", logs_export: "admin"
};
var SESSION_PREFIX = "WASTE_SS_";
var SESSION_TTL = 43200;          // 12 h absolute session length
var CACHE_MAX_TTL = 21600;        // CacheService keeps a value at most 6 h
var SESSION_RENEW_MS = 3600000;   // re-put the session entry at most once an hour while in use
var FAIL_GLOBAL_KEY = "WASTE_LOGIN_FAILURES";
var FAIL_CLIENT_PREFIX = "WASTE_LF_";
var CLIENT_MAX_FAILURES = 5;      // wrong passwords per browser (client id) per 15 min
var GLOBAL_MAX_FAILURES = 30;     // wrong passwords from everyone per 15 min
var PASSWORD_LOCK_SECONDS = 900;
var RESET_TABLES = ["movements", "audit_logs"]; // waste_types are kept

// Asia/Bangkok has no daylight saving time: fixed +07:00.
var TZ_OFFSET_MS = 7 * 60 * 60 * 1000;

// Company identity. name is always shown; the others are OPTIONAL: the page shows
// the factory-registration / permit chips only when they are filled in.
var COMPANY = {
  name: "บริษัท เทพวัฒนา จำกัด",
  factoryReg: "820200300125423",   // ทะเบียนโรงงานเลขที่
  permitNo: "2569-12560",     // เลขที่หนังสืออนุญาต (สก.2)
  permitYear: 2569    // ปีของหนังสืออนุญาต
};

var CONTAINERS = { drum200: "ถัง 200 ลิตร", ibc: "IBC", bigbag: "Bigbag" };
var CONTAINER_KEYS = ["drum200", "ibc", "bigbag"];
var KIND_NAME = { store: "บันทึกจัดเก็บ", ship: "ส่งกำจัด / นำออก", opening: "ยอดยกมา", treat: "ข้อมูลภายใน" };
var KINDS = ["store", "ship", "opening", "treat"];
var OUT_KINDS = ["ship", "treat"]; // leave the in-plant balance
var PREFIX = { store: "ST", ship: "SH", opening: "OB", treat: "TR" };
var ACT_NAME = {
  store: "บันทึกจัดเก็บ", ship: "ส่งกำจัด", opening: "ยอดยกมา", treat: "บันทึกข้อมูลภายใน",
  movement_update: "แก้ไขรายการ", movement_delete: "ลบรายการ", type_save: "ตั้งค่าประเภทของเสีย",
  receiver_save: "ตั้งค่าผู้รับดำเนินการ", receiver_delete: "ลบผู้รับดำเนินการ", type_delete: "ลบประเภทของเสีย"
};
var LOG_CAT_NAME = { waste: "ของเสีย", edit: "แก้ไข / ลบ", setting: "ตั้งค่า", warn: "ถูกปฏิเสธ" };

var QUOTA_WARN_PCT = 80;        // >= 80 % of the permitted quantity: warning
var STORAGE_ALERT_DAYS = 90;    // oldest stored lot (FIFO) older than this: alert
var MAX_WEIGHT_KG = 1000000;    // one movement (1,000 t)
var MAX_PERMIT_TONS = 100000;
var FUTURE_SLACK_MS = 10 * 60 * 1000; // client clock may be a few minutes ahead

/* Column definitions. Types: int, num (decimal), str, dt ("YYYY-MM-DD HH:MM:SS").
 * A trailing "?" marks columns that are null when empty. New columns are only
 * ever APPENDED (older sheets get them added on first use, see table_). */
var SCHEMA = {
  waste_types: [
    ["id", "int"], ["code", "str"], ["name", "str"], ["permit_tons", "num"], ["mgmt_code", "str"],
    ["receiver_reg", "str"], ["receiver", "str"], ["icon", "str"], ["color", "str"], ["active", "int"],
    ["sort", "int"], ["updated_at", "dt?"], ["updated_by", "str"]
  ],
  // Receivers (ผู้รับดำเนินการ). A type keeps its receiver name + registration
  // (waste_types.receiver / receiver_reg) as before; this sheet only adds the
  // receivers that have no type yet, an optional default management code and the
  // active flag. Older sheets are created and seeded from the types (ensureReceivers_).
  receivers: [
    ["id", "int"], ["name", "str"], ["reg", "str"], ["mgmt_code", "str"], ["active", "int"],
    ["sort", "int"], ["updated_at", "dt?"], ["updated_by", "str"]
  ],
  movements: [
    ["id", "int"], ["doc_no", "str"], ["kind", "str"], ["type_id", "int"], ["weight_kg", "num"],
    ["container", "str"], ["containers", "int"], ["location", "str"], ["manifest_no", "str"], ["vehicle", "str"],
    ["note", "str"], ["moved_at", "dt"], ["created_at", "dt"], ["actor", "str"],
    ["updated_at", "dt?"], ["updated_by", "str"], ["method", "str"]
  ],
  audit_logs: [
    ["id", "int"], ["category", "str"], ["action", "str"], ["ref", "str"], ["detail", "str"],
    ["actor", "str"], ["created_at", "dt"]
  ]
};
var TABLE_ORDER = ["waste_types", "receivers", "movements", "audit_logs"];
var AUTO_CREATE_TABLES = ["receivers"];

/* Permit seed: the 40 lines of the สก.2 letter 2569-12560 (total 1,101.55 t), from the
 * workbook Waste_Summary_Report_Thepwatana_2569.xlsx, sheet "รายการที่ได้รับอนุญาตทั้งหมด".
 * Each line is its own waste type (a code can repeat with another receiver / management
 * code). setupSystem() seeds a NEW database with it; for an existing EMPTY database run
 * seedThepwatanaPermit() once. Admin can still edit everything in the settings page. */
var SEED_ON_SETUP = true;
var PERMIT_LINES = [
  // [code, name, permit tons, management code, receiver registration, receiver, icon, colour]
  ["150111","กระป๋องสเปรย์",1,"073","20190300225401","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-spray-can","#D92626"],
  ["160215","หลอดไฟ",1,"073","20190300225401","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-lightbulb","#20B64C"],
  ["080409","กากกาว",20,"042","10190000825494","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-fill-drip","#8E26D9"],
  ["150202","วัสดุปนเปื้อน",60,"042","10190000825494","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-biohazard","#B6A320"],
  ["160698","ถ่านไฟฉาย",0.05,"073","20190300225401","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-battery-half","#26BBD9"],
  ["191204","เศษสติ๊กเกอร์",40,"046","10190000825494","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-note-sticky","#B62071"],
  ["150203","อุปกรณ์ PPE",30,"071","20190300225401","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-user-shield","#53D926"],
  ["160213","อุปกรณ์อิเล็กทรอนิกส์",5,"073","20190300225401","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-microchip","#2720B6"],
  ["140603","โซลเว้นท์ใช้แล้ว",20,"042","10190000825494","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-flask","#D96226"],
  ["150110","ตลับหมึก ภาชนะปนเปื้อน",40,"073","20190300225401","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-print","#20B67E"],
  ["160601","แบตเตอรี่",5,"021","10190000825494","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-car-battery","#CA26D9"],
  ["150110","บรรจุภัณฑ์ปนเปื้อน",400,"039","10740101225521","บจก. อารยาค้าถัง","fa-box-archive","#97B620"],
  ["150102","บรรจุภัณฑ์พลาสติก พลาสติก",10,"046","10190000825494","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-bottle-water","#267FD9"],
  ["150102","บรรจุภัณฑ์พลาสติก พลาสติก",30,"011","10110020125675","บจก. มั่นคงยั่งยืน รีไซเคิล","fa-recycle","#B6203F"],
  ["150101","กระดาษ",15,"011","10110020125675","บจก. มั่นคงยั่งยืน รีไซเคิล","fa-newspaper","#26D935"],
  ["130208","น้ำมันใช้แล้ว",2,"075","82020000125442","บมจ. อัคคีปราการ","fa-oil-can","#5920B6"],
  ["070404","Xylene",5,"075","82020000125442","บมจ. อัคคีปราการ","fa-vial","#D99E26"],
  ["070411","ตะกอนน้ำเสีย",30,"075","82020000125442","บมจ. อัคคีปราการ","fa-water","#20B6B0"],
  ["070413","วัตถุดิบเสื่อมสภาพ ฝุ่นจากกระบวนการผลิต",20,"075","82020000125442","บมจ. อัคคีปราการ","fa-industry","#D926AC"],
  ["080409","กากกาว",10,"075","82020000125442","บมจ. อัคคีปราการ","fa-droplet","#64B620"],
  ["150202","วัสดุปนเปื้อน ทรายดูดซับ",30,"075","82020000125442","บมจ. อัคคีปราการ","fa-hill-rockslide","#2644D9"],
  ["070401","Waste water contaminated with chemical Cleaning Process สารเคมีเสื่อมสภาพ น้ำเสียจากการล้าง",20,"075","82020000125442","บมจ. อัคคีปราการ","fa-flask-vial","#B63320"],
  ["160506","Lab waste",20,"075","82020000125442","บมจ. อัคคีปราการ","fa-pump-soap","#26D971"],
  ["160213","อุปกรณ์อิเล็กทรอนิกส์",1,"075","82020000125442","บมจ. อัคคีปราการ","fa-microscope","#8B20B6"],
  ["150110","ภาชนะปนเปื้อน",30,"075","82020000125442","บมจ. อัคคีปราการ","fa-computer","#D8D926"],
  ["150111","กระป๋องสเปรย์",1,"075","82020000125442","บมจ. อัคคีปราการ","fa-box","#208AB6"],
  ["160305","Expire chemical Liquid Expire chemical Solid",40,"075","82020000125442","บมจ. อัคคีปราการ","fa-spray-can-sparkles","#D92670"],
  ["150104","แผ่นฟอยล์",5,"071","20190300225401","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-skull-crossbones","#32B620"],
  ["150202","วัสดุปนเปื้อน ทรายดูดซับ",30,"042","72070000125407","บมจ. บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม","fa-layer-group","#4526D9"],
  ["190814","ตะกอนน้ำเสีย",20,"071","10700001425472","บมจ. บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม","fa-battery-full","#B66520"],
  ["160709","Waste water contaminated with chemical",20,"065","10103300225380","บมจ. บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม","fa-burst","#26D9AD"],
  ["191201","ของเสียทั่วไปจากการผลิต",30,"046","72070000125407","บมจ. บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม","fa-tint","#B620AF"],
  ["1601001","น้ำล้างเครื่องจักร น้ำเสียผ่านระบบบำบัดแล้ว",50,"042","10190000825494","บมจ. เบตเตอร์ เวิลด์ กรีน","fa-hand-sparkles","#9CD926"],
  ["1601001","น้ำล้างเครื่องจักร น้ำเสียผ่านระบบบำบัดแล้ว",20,"065","10103300225380","บมจ. บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม","fa-gas-pump","#2058B6"],
  ["140603","ตัวทำละลาย",5,"051","10240002925477","บจก. สุขเจริญทรัพย์ วังเย็น","fa-faucet-drip","#D92634"],
  ["130208","น้ำมันเครื่องยนต์ น้ำมันเกียร์ น้ำมันหล่อลื่นที่ใช้แล้ว",5,"042","10240002925477","บจก. สุขเจริญทรัพย์ วังเย็น","fa-vials","#20B640"],
  ["191212","ของเสียทั่วไปจากกระบวนการผลิต",10,"011","10110100225577","บจก. กรีนเซฟรีไซเคิ้ล","fa-trash-can","#8026D9"],
  ["160215","Fluorescent Lamp",0.5,"049","10240004525614","บจก. ซี.อี.เอ คลีนโปรดักส์","fa-lightbulb","#B69820"],
  ["150110","Contaminated Container",10,"039","10110102325581","บจก. กรีนเซฟรีไซเคิ้ล","fa-box-open","#26C9D9"],
  ["150202","Contaminated Material Absorbent",10,"042","10110102325581","บจก. กรีนเซฟรีไซเคิ้ล","fa-radiation","#B6207D"]
];
// [receiver name, primary registration] - a receiver that has more registrations keeps the first
// one here (the registration of each line stays on its waste type).
var PERMIT_RECEIVERS = [
  ["บจก. อารยาค้าถัง", "10740101225521"],
  ["บมจ. เบตเตอร์ เวิลด์ กรีน", "10190000825494"],
  ["บมจ. อัคคีปราการ", "82020000125442"],
  ["บมจ. บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม", "72070000125407"],
  ["บจก. มั่นคงยั่งยืน รีไซเคิล", "10110020125675"],
  ["บจก. กรีนเซฟรีไซเคิ้ล", "10110102325581"],
  ["บจก. สุขเจริญทรัพย์ วังเย็น", "10240002925477"],
  ["บจก. ซี.อี.เอ คลีนโปรดักส์", "10240004525614"]
];

var READ_ACTIONS = ["bootstrap", "dashboard", "history", "export", "logs", "logs_export", "batch"];
// "batch": several reads in one request; writes accept the same optional
// "reads" list and return the post-write results as data.reads.
var BATCH_ACTIONS = ["bootstrap", "dashboard", "history"];
var BATCH_MAX = 6;

/* Read cache (CacheService), keyed by a per-sheet data version that every
 * write bumps right after its changes are flushed (see flush_). Writes always
 * read the sheets live under the script lock. Hand edits in the spreadsheet
 * show up after READ_CACHE_TTL seconds, or at once after clearReadCache(). */
var CACHED_TABLES = ["waste_types", "receivers", "movements", "audit_logs"];
var VERSION_PREFIX = "WASTE_V_";
var VERSION_TTL = 21600;
var RC_PREFIX = "WASTE_RC_";
var READ_CACHE_TTL = 600;
var KEEP_WARM_TTL = 900;
var KEEP_WARM_MINUTES = 10;
var RC_CHUNK = 30000;
var RC_MAX_CHUNKS = 30;
var WRITE_ACTIONS = ["store", "ship", "opening", "treat", RESET_ACTION, "movement_update", "movement_delete", "type_save", "receiver_save", "receiver_delete", "type_delete"];

var MAX_PERSON = 100;
var MAX_TEXT = 255;

/* ===================== web app entry points ===================== */

// GET: public reads only, always anonymous (a token in a URL is never used).
// Everything else (and every request carrying a token) is POST.
function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = String(params.action || "");
  try {
    if (!action) return jsonResponse_({ ok: true, data: { service: "Balance Waste API" } });
    if (READ_ACTIONS.indexOf(action) !== -1 && ACTION_ROLE[action] === PUBLIC) {
      resetRequest_("");
      REQ_.role = null;
      REQ_.treat = false;
      return jsonResponse_({ ok: true, data: runRead_(action, params) });
    }
    if (action === "login" || action === "logout" || Object.prototype.hasOwnProperty.call(ACTION_ROLE, action)) {
      throw new Error("คำสั่งนี้ต้องส่งแบบ POST");
    }
    fail_("Unknown action");
  } catch (error) {
    return errorResponse_(error);
  }
}

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  } catch (parseError) {
    return jsonResponse_({ ok: false, error: "ข้อมูลที่ส่งมาไม่ใช่ JSON ที่ถูกต้อง" });
  }
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) payload = {};
  var action = String(payload.action || "");

  var reads = null;
  var sess;
  try {
    if (action === "login") return jsonResponse_({ ok: true, data: actionLogin_(payload) });
    if (action === "logout") return jsonResponse_({ ok: true, data: actionLogout_(payload.token) });
    if (!Object.prototype.hasOwnProperty.call(ACTION_ROLE, action)) fail_("Unknown action");
    // Server-side access control on every action (never trusts the page).
    // Public actions: a missing / unknown / expired token is simply anonymous.
    sess = ACTION_ROLE[action] === PUBLIC
      ? (session_(payload.token) || { role: null })
      : requireRole_(payload.token, ACTION_ROLE[action]);
    // In-house rows: any valid session that does not ask to hide them.
    sess.treat = Boolean(sess.role) && !phpTruthy_(payload.hideTreat);
    if (READ_ACTIONS.indexOf(action) !== -1) {
      resetRequest_("");
      REQ_.role = sess.role;
      REQ_.treat = sess.treat;
      return jsonResponse_({ ok: true, data: runRead_(action, payload) });
    }
    if (payload.reads != null && payload.reads !== "") reads = parseReads_(payload.reads);
    // Every change must say who made it (name set in the page header).
    if (!actorIn_(payload)) fail_("กรุณาระบุชื่อผู้ใช้งาน (มุมขวาบน) ก่อนบันทึก", "NAME_REQUIRED");
    // Wiping all data: the admin password is typed again. Checked before taking
    // the lock so guessing (1 s delay) never blocks writers.
    if (action === RESET_ACTION) assertResetPassword_(payload.resetPassword, payload.client);
  } catch (error) {
    return errorResponse_(error);
  }
  return runWrite_(action, payload, reads, sess);
}

function actorIn_(input) {
  var v = input.actor;
  return mbSubstr_(phpTrim_(v !== null && typeof v === "object" ? "" : str_(v)), 0, MAX_PERSON);
}

function runWrite_(action, input, reads, sess) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockError) {
    return jsonResponse_({ ok: false, error: "ระบบกำลังบันทึกรายการอื่นอยู่ กรุณาลองใหม่อีกครั้ง" });
  }
  var actor = actorIn_(input);
  var result;
  var start = function () {
    resetRequest_(actor);
    REQ_.role = sess.role;
    REQ_.treat = sess.treat;
  };
  try {
    start();
    try {
      result = handleWrite_(action, input);
      flush_();
      if (reads) versions_();
    } catch (error) {
      // Nothing buffered by the failed action is written (acts as a rollback).
      var actName = ACT_NAME[action];
      if (actName && action !== RESET_ACTION) {
        try {
          start();
          audit_("warn", action, "ปฏิเสธ" + actName + ": " + errorMessage_(error), "");
          flush_();
        } catch (ignored) {}
      }
      return errorResponse_(error);
    }
  } finally {
    lock.releaseLock();
  }
  if (reads) result.reads = afterReads_(reads);
  return jsonResponse_({ ok: true, data: result });
}

/* ===================== reads: single, batch, after a write ===================== */

function runRead_(action, p) {
  REQ_.cacheReads = true;
  return handleRead_(action, p);
}

function parseReads_(raw) {
  var list = raw;
  if (typeof raw === "string") {
    try { list = JSON.parse(raw); } catch (e) { fail_("reads ไม่ถูกต้อง"); }
  }
  if (!Array.isArray(list) || !list.length || list.length > BATCH_MAX) fail_("reads ไม่ถูกต้อง");
  return list.map(function (r) {
    if (typeof r === "string") r = { action: r };
    if (!r || typeof r !== "object" || Array.isArray(r)) fail_("reads ไม่ถูกต้อง");
    if (BATCH_ACTIONS.indexOf(String(r.action || "")) === -1) fail_("Unknown action");
    return r;
  });
}

function actionBatch_(p) {
  var reads = parseReads_(p.reads);
  return { results: reads.map(function (r) { return handleRead_(r.action, r); }) };
}

function afterReads_(reads) {
  try {
    REQ_.cacheReads = true;
    return reads.map(function (r) { return handleRead_(r.action, r); });
  } catch (error) {
    return null;
  }
}

/* ===================== login / sessions / passwords ===================== */

function cache_() {
  return CacheService.getScriptCache();
}

// SHA-256 of a text as signed bytes (Apps Script digest format).
function digest_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
}

// Constant-time comparison of two digests (no early exit).
function sameBytes_(a, b) {
  var diff = a.length ^ b.length;
  for (var i = 0; i < a.length && i < b.length; i++) diff |= (a[i] ^ b[i]);
  return diff === 0;
}

function hex_(bytes) {
  return bytes.map(function (b) { return ("0" + (b & 255).toString(16)).slice(-2); }).join("");
}

// Ties a session to the password it was opened with: a changed / removed
// password ends the sessions of that level. Never the password itself.
function fingerprint_(password) {
  return hex_(digest_("balance-waste-session|" + password)).slice(0, 24);
}

function passwordError_(message) {
  var error = new Error(message);
  error.passwordError = true;
  return error;
}

// Optional random id of the browser (sent by the page) for the per-client lockout.
function clientId_(v) {
  return typeof v === "string" && /^[0-9a-f]{16,64}$/.test(v) ? v : "";
}

var LOCKED_MESSAGE = "ใส่รหัสผิดหลายครั้งเกินไป กรุณารอ 15 นาที / Too many wrong attempts. Try again in 15 minutes.";

function lockoutCheck_(client) {
  var cache = cache_();
  if (Number(cache.get(FAIL_GLOBAL_KEY) || 0) >= GLOBAL_MAX_FAILURES) throw passwordError_(LOCKED_MESSAGE);
  if (client && Number(cache.get(FAIL_CLIENT_PREFIX + client) || 0) >= CLIENT_MAX_FAILURES) throw passwordError_(LOCKED_MESSAGE);
}

function recordFailure_(client) {
  var cache = cache_();
  cache.put(FAIL_GLOBAL_KEY, String(Number(cache.get(FAIL_GLOBAL_KEY) || 0) + 1), PASSWORD_LOCK_SECONDS);
  if (client) {
    var key = FAIL_CLIENT_PREFIX + client;
    cache.put(key, String(Number(cache.get(key) || 0) + 1), PASSWORD_LOCK_SECONDS);
  }
  Utilities.sleep(1000);
}

function missingPasswordsMessage_() {
  return "ยังไม่ได้ตั้งรหัสผ่านใน Script Properties (" + ROLE_ORDER.map(function (r) { return ROLE_PROPERTY[r]; }).join(", ") +
    ") กรุณาให้ผู้ดูแลระบบตั้งค่าที่ Project Settings > Script properties ก่อน";
}

// Compares a typed password with one stored password (digests, constant time).
function passwordMatches_(typed, expected) {
  var ok = sameBytes_(digest_(typeof typed === "string" ? typed : ""), digest_(expected || ""));
  return ok && typeof typed === "string" && Boolean(expected);
}

// login {password, client?} -> {token, role, role_name, expires_at, ttl}
function actionLogin_(p) {
  var client = clientId_(p.client);
  var props = PropertiesService.getScriptProperties().getProperties() || {};
  var configured = ROLE_ORDER.filter(function (r) { return Boolean(props[ROLE_PROPERTY[r]]); });
  if (!configured.length) throw passwordError_(missingPasswordsMessage_());
  lockoutCheck_(client);
  var role = null;
  ROLE_ORDER.forEach(function (r) { // every level is compared (no early exit)
    var same = passwordMatches_(p.password, props[ROLE_PROPERTY[r]]);
    if (same && !role) role = r;
  });
  if (!role) {
    recordFailure_(client);
    throw passwordError_("รหัสไม่ถูกต้อง / Incorrect password.");
  }
  if (client) {
    try { cache_().remove(FAIL_CLIENT_PREFIX + client); } catch (ignored) {}
  }
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "").toLowerCase();
  var now = Date.now();
  var exp = now + SESSION_TTL * 1000;
  try {
    cache_().put(SESSION_PREFIX + token, JSON.stringify({ r: role, e: exp, p: now, f: fingerprint_(props[ROLE_PROPERTY[role]]) }), CACHE_MAX_TTL);
  } catch (error) {
    fail_("ระบบเข้าสู่ระบบขัดข้องชั่วคราว กรุณาลองใหม่อีกครั้ง");
  }
  return { token: token, role: role, role_name: ROLE_NAME[role], expires_at: exp, ttl: SESSION_TTL };
}

function actionLogout_(token) {
  if (typeof token === "string" && /^[0-9a-f]{64}$/.test(token)) {
    try { cache_().remove(SESSION_PREFIX + token); } catch (ignored) {}
  }
  return { loggedOut: true };
}

// The session of a token, or null (unknown, expired, password changed, cache error).
function session_(token) {
  if (typeof token !== "string" || !/^[0-9a-f]{64}$/.test(token)) return null;
  var key = SESSION_PREFIX + token;
  var rec;
  try { rec = JSON.parse(cache_().get(key) || "null"); } catch (error) { return null; }
  var now = Date.now();
  if (!rec || typeof rec !== "object" || !Object.prototype.hasOwnProperty.call(ROLE_LEVEL, rec.r) || !(Number(rec.e) > now)) return null;
  var pw = PropertiesService.getScriptProperties().getProperty(ROLE_PROPERTY[rec.r]) || "";
  if (!pw || fingerprint_(pw) !== rec.f) return null;
  if (!(now - Number(rec.p) < SESSION_RENEW_MS)) {
    // CacheService keeps an entry <= 6 h: renew it while the 12 h session is in use.
    rec.p = now;
    var ttl = Math.max(1, Math.min(CACHE_MAX_TTL, Math.ceil((rec.e - now) / 1000)));
    try { cache_().put(key, JSON.stringify(rec), ttl); } catch (ignored) {}
  }
  return { role: rec.r, expires_at: rec.e };
}

function requireRole_(token, need) {
  var s = session_(token);
  if (!s) {
    fail_("ต้องใส่รหัสผ่าน “" + ROLE_NAME[need] + "” ขึ้นไปก่อน (ยังไม่ได้ปลดล็อก หรือหมดเวลา)", "AUTH_REQUIRED");
  }
  if (ROLE_LEVEL[s.role] < ROLE_LEVEL[need]) {
    fail_("สิทธิ์ “" + ROLE_NAME[s.role] + "” ทำรายการนี้ไม่ได้ (ต้องเป็น “" + ROLE_NAME[need] + "”)", "FORBIDDEN");
  }
  return s;
}

// reset_data: the admin password typed again (same lockout as login).
function assertResetPassword_(password, client) {
  var expected = PropertiesService.getScriptProperties().getProperty(RESET_PASSWORD_PROPERTY) || "";
  if (!expected) {
    throw passwordError_("ยังไม่ได้ตั้งรหัสผ่านผู้ดูแลใน Script Properties (" + RESET_PASSWORD_PROPERTY +
      ") กรุณาให้ผู้ดูแลระบบตั้งค่าที่ Project Settings > Script properties ก่อน");
  }
  client = clientId_(client);
  lockoutCheck_(client);
  if (!passwordMatches_(password, expected)) {
    recordFailure_(client);
    throw passwordError_("รหัสไม่ถูกต้อง / Incorrect password.");
  }
}

// True while the current request may see the individual in-house (treat) rows:
// a valid session of any level that did not ask to hide them. Totals are public.
function treatOn_() {
  return Boolean(REQ_ && REQ_.treat);
}

// True while the current request carries a valid session of any level: only
// then are the AGGREGATE treatment totals (treated_kg, monthly treat series,
// total waste) sent. Independent of hideTreat (that only hides the TR- rows).
function aggOn_() {
  return Boolean(REQ_ && REQ_.role);
}

function isOut_(kind) {
  return OUT_KINDS.indexOf(kind) !== -1;
}

/* ===================== router ===================== */

function handleRead_(action, p) {
  switch (action) {
    case "bootstrap": return actionBootstrap_();
    case "dashboard": return actionDashboard_();
    case "history": return { items: queryHistory_(p) };
    case "export": return exportHistory_(queryHistory_(p));
    case "logs": return { items: queryLogs_(p) };
    case "logs_export": return exportLogs_(queryLogs_(p));
    case "batch": return actionBatch_(p);
  }
  fail_("Unknown action");
}

function handleWrite_(action, input) {
  switch (action) {
    case "store": return actionMove_("store", input);
    case "ship": return actionMove_("ship", input);
    case "opening": return actionMove_("opening", input);
    case "treat": return actionMove_("treat", input);
    case RESET_ACTION: return actionResetData_();
    case "movement_update": return actionMovementUpdate_(input);
    case "movement_delete": return actionMovementDelete_(input);
    case "type_save": return actionTypeSave_(input);
    case "receiver_save": return actionReceiverSave_(input);
    case "receiver_delete": return actionReceiverDelete_(input);
    case "type_delete": return actionTypeDelete_(input);
  }
  fail_("Unknown action");
}

/* ===================== balance math =====================
 * Weights are summed in integer hundredths of a kg ("cents") so 0.1 + 0.2
 * never drifts. in = store + opening; balance = in - shipped - treated.
 */

function cents_(kg) {
  return Math.round(Number(kg || 0) * 100);
}

function kgOf_(cents) {
  return Math.round(cents) / 100;
}

function sortedTypes_() {
  return sortBy_(table_("waste_types").rows.slice(), function (a, b) { return a.sort - b.sort || a.id - b.id; });
}

// Per waste type: totals, containers in storage, this year's shipments vs the
// permitted quantity and the date of the oldest lot still in storage (FIFO).
// treated_kg / treated_year_kg (in-house totals) go only to a session (see
// publicBalances_); balance_kg is always the true in-plant balance (see publicBalances_ for last_at).
function balances_() {
  var now = nowParts_();
  var year = now.date.slice(0, 4);
  var out = {};
  var fifo = {};
  table_("waste_types").rows.forEach(function (t) {
    out[t.id] = {
      type_id: t.id, stored_c: 0, opening_c: 0, shipped_c: 0, year_c: 0, treated_c: 0, treat_year_c: 0,
      containers: { drum200: 0, ibc: 0, bigbag: 0 }, last_at: null, last_vis_at: null
    };
    fifo[t.id] = [];
  });
  sortBy_(table_("movements").rows.slice(), byMovedAsc_).forEach(function (m) {
    var b = out[m.type_id];
    if (!b) return;
    var c = cents_(m.weight_kg);
    var sign = isOut_(m.kind) ? -1 : 1;
    if (m.kind === "store") b.stored_c += c;
    else if (m.kind === "opening") b.opening_c += c;
    else if (m.kind === "ship") {
      b.shipped_c += c;
      if (m.moved_at.slice(0, 4) === year) b.year_c += c;
    } else if (m.kind === "treat") {
      b.treated_c += c;
      if (m.moved_at.slice(0, 4) === year) b.treat_year_c += c;
    }
    if (b.containers[m.container] !== undefined) b.containers[m.container] += sign * (m.containers || 0);
    b.last_at = m.moved_at;
    if (m.kind !== "treat") b.last_vis_at = m.moved_at; // latest movement a hidden-mode request may see
    // FIFO lots: what is shipped / treated leaves the oldest lot first.
    var q = fifo[m.type_id];
    if (sign > 0) q.push({ at: m.moved_at, c: c });
    else {
      var left = c;
      while (left > 0 && q.length) {
        var take = Math.min(left, q[0].c);
        q[0].c -= take; left -= take;
        if (q[0].c <= 0) q.shift();
      }
    }
  });
  var nowMs = parseDt_(now.datetime);
  table_("waste_types").rows.forEach(function (t) {
    var b = out[t.id];
    var permitC = Math.round(Number(t.permit_tons || 0) * 100000);
    var bal = b.stored_c + b.opening_c - b.shipped_c - b.treated_c;
    var oldest = bal > 0 && fifo[t.id].length ? fifo[t.id][0].at : null;
    out[t.id] = {
      type_id: t.id,
      stored_kg: kgOf_(b.stored_c),
      opening_kg: kgOf_(b.opening_c),
      in_kg: kgOf_(b.stored_c + b.opening_c),
      shipped_kg: kgOf_(b.shipped_c),
      treated_kg: kgOf_(b.treated_c),
      treated_year_kg: kgOf_(b.treat_year_c),
      balance_kg: kgOf_(bal),
      containers: b.containers,
      shipped_year_kg: kgOf_(b.year_c),
      permit_kg: kgOf_(permitC),
      quota_pct: permitC > 0 ? Math.round(b.year_c / permitC * 1000) / 10 : null,
      quota_level: quotaLevel_(b.year_c, permitC),
      oldest_at: oldest,
      age_days: oldest ? Math.max(0, Math.floor((nowMs - parseDt_(oldest)) / 86400000)) : null,
      last_at: b.last_at,
      last_vis_at: b.last_vis_at
    };
  });
  return out;
}

// Balances as sent to a request. The totals are public; without access to the
// in-house rows last_at never shows the time of one (row-level detail).
function publicBalances_(bal) {
  var showTreat = treatOn_();
  var showAgg = aggOn_();
  var out = {};
  Object.keys(bal).forEach(function (id) {
    var o = plain_(bal[id]);
    if (!showAgg) { delete o.treated_kg; delete o.treated_year_kg; }
    if (!showTreat) o.last_at = o.last_vis_at;
    delete o.last_vis_at;
    out[id] = o;
  });
  return out;
}

function quotaLevel_(usedC, permitC) {
  if (!(permitC > 0)) return usedC > 0 ? "over" : "ok";
  if (usedC > permitC) return "over";
  if (usedC * 100 >= permitC * QUOTA_WARN_PCT) return "warn";
  return "ok";
}

/* Chronological replay of every movement (time order, then id): the weight of
 * each waste type and the container count of each (type, container) pair must
 * never go below zero. Returns null when fine, else the first problem. */
function replay_(moves) {
  var w = {};
  var n = {};
  var list = sortBy_(moves.slice(), byMovedAsc_);
  for (var i = 0; i < list.length; i++) {
    var m = list[i];
    var c = cents_(m.weight_kg);
    var k = m.type_id + "|" + m.container;
    var sign = isOut_(m.kind) ? -1 : 1;
    var beforeW = w[m.type_id] || 0;
    var beforeN = n[k] || 0;
    w[m.type_id] = beforeW + sign * c;
    n[k] = beforeN + sign * (m.containers || 0);
    if (w[m.type_id] < 0) return { move: m, what: "weight", have: beforeW, need: c };
    if (n[k] < 0) return { move: m, what: "containers", have: beforeN, need: m.containers };
  }
  return null;
}

// verb: what is being tried (e.g. "ส่งกำจัดไม่ได้"); newMove: the row being
// added / changed (its problem gets a direct message).
function validateLedger_(moves, verb, newMove) {
  var p = replay_(moves);
  if (!p) return;
  var t = typeLabel_(p.move.type_id);
  var when = dtTh_(p.move.moved_at);
  var detail = p.what === "weight"
    ? "น้ำหนักคงเหลือของ " + t + " ณ " + when + " มี " + fmtT_(kgOf_(p.have)) + " ตัน แต่ต้องนำออก " + fmtT_(kgOf_(p.need)) + " ตัน"
    : "จำนวน" + (CONTAINERS[p.move.container] || p.move.container) + "ของ " + t + " ณ " + when + " มี " + p.have +
      " ภาชนะ แต่ต้องนำออก " + p.need + " ภาชนะ";
  if (newMove && p.move === newMove) fail_(verb + " — " + detail, "NEGATIVE_BALANCE");
  if (p.move.kind === "treat" && !treatOn_()) {
    // Hidden mode: do not reveal the treatment row (number / kind) in the message.
    fail_(verb + " — ยอดคงเหลือจะติดลบในรายการภายหลัง ณ " + when, "NEGATIVE_BALANCE");
  }
  fail_(verb + " — ยอดคงเหลือจะติดลบที่รายการ " + (p.move.doc_no || "") + " (" + KIND_NAME[p.move.kind] + "): " + detail, "NEGATIVE_BALANCE");
}

function typeLabel_(id) {
  var t = typeMap_()[id];
  return t ? t.code + " " + t.name : "#" + id;
}

/* ===================== read actions ===================== */

function actionBootstrap_() {
  var now = nowParts_();
  var core = cachedDerived_("boot4_" + modeTag_() + now.date, ["waste_types", "receivers", "movements"], function () {
    var used = {};
    visibleMoves_().forEach(function (m) { used[m.type_id] = true; });
    var types = sortedTypes_().map(function (t) { var o = plain_(t); o.has_moves = used[t.id] === true; return o; });
    return { types: types, receivers: receiverStats_(types), balance: publicBalances_(balances_()) };
  });
  return {
    types: core.types,
    receivers: core.receivers,
    balance: core.balance,
    treat: treatOn_(),
    role: (REQ_ && REQ_.role) || null,
    role_name: REQ_ && REQ_.role ? ROLE_NAME[REQ_.role] : "",
    containers: CONTAINERS,
    company: COMPANY,
    limits: { quotaWarnPct: QUOTA_WARN_PCT, storageAlertDays: STORAGE_ALERT_DAYS },
    now: { date: now.date, time: now.time }
  };
}

// Cache-key part for derived reads that differ by treatment visibility.
function modeTag_() {
  return (aggOn_() ? "a" : "") + (treatOn_() ? "t_" : "_");
}

// Movements a request may see (in-house "treat" rows only with access, see treatOn_).
function visibleMoves_() {
  var rows = table_("movements").rows;
  return treatOn_() ? rows : rows.filter(function (m) { return m.kind !== "treat"; });
}

function actionDashboard_() {
  var now = nowParts_();
  var month = now.date.slice(0, 7);
  var year = now.date.slice(0, 4);
  // Treatment aggregates (kpi treated / total / month / year, monthly treat
  // series) only for a session (aggOn_); individual in-house rows never appear.
  // ("4" in the cache ids: a new shape, never reuse a cached answer of the previous version)
  var agg = aggOn_();
  var core = cachedDerived_("dash5_" + (agg ? "a" : "p") + now.date, ["waste_types", "receivers", "movements", "audit_logs"], function () {
    var types = sortedTypes_();
    var bal = balances_();
    var moves = table_("movements").rows;
    var publicMoves = moves.filter(function (m) { return m.kind !== "treat"; });

    // last 12 months (oldest first): kg stored / shipped
    var months = [];
    var monthIdx = {};
    var y = Number(year), mo = Number(now.date.slice(5, 7));
    for (var i = 11; i >= 0; i--) {
      var yy = y, mm = mo - i;
      while (mm < 1) { mm += 12; yy -= 1; }
      var key = yy + "-" + pad2_(mm);
      monthIdx[key] = months.length;
      months.push({ ym: key, in_c: 0, out_c: 0, treat_c: 0 });
    }
    var monthIn = 0, monthOut = 0, yearIn = 0, monthTreat = 0, yearTreat = 0;
    moves.forEach(function (m) {
      var c = cents_(m.weight_kg);
      var ym = m.moved_at.slice(0, 7);
      if (monthIdx[ym] !== undefined) {
        if (m.kind === "store") months[monthIdx[ym]].in_c += c;
        if (m.kind === "ship") months[monthIdx[ym]].out_c += c;
        if (m.kind === "treat") months[monthIdx[ym]].treat_c += c;
      }
      if (m.kind === "store" && ym === month) monthIn += c;
      if (m.kind === "ship" && ym === month) monthOut += c;
      if (m.kind === "store" && ym.slice(0, 4) === year) yearIn += c;
      if (m.kind === "treat" && ym === month) monthTreat += c;
      if (m.kind === "treat" && ym.slice(0, 4) === year) yearTreat += c;
    });

    var totalBal = 0, totalYear = 0, totalPermit = 0, totalIn = 0, totalShip = 0, totalTreat = 0;
    var containers = { drum200: 0, ibc: 0, bigbag: 0 };
    var alerts = [];
    types.forEach(function (t) {
      var b = bal[t.id];
      totalBal += cents_(b.balance_kg);
      totalYear += cents_(b.shipped_year_kg);
      totalPermit += cents_(b.permit_kg);
      totalIn += cents_(b.in_kg); totalShip += cents_(b.shipped_kg); totalTreat += cents_(b.treated_kg);
      CONTAINER_KEYS.forEach(function (k) { containers[k] += Math.max(0, b.containers[k]); });
      if (b.quota_level !== "ok") {
        alerts.push({ level: b.quota_level === "over" ? "bad" : "warn", kind: "quota", type_id: t.id, pct: b.quota_pct,
          used_kg: b.shipped_year_kg, permit_kg: b.permit_kg });
      }
      if (b.age_days !== null && b.age_days > STORAGE_ALERT_DAYS) {
        alerts.push({ level: "warn", kind: "age", type_id: t.id, days: b.age_days, since: b.oldest_at, balance_kg: b.balance_kg });
      }
    });
    alerts.sort(function (a, b) { return (a.level === "bad" ? 0 : 1) - (b.level === "bad" ? 0 : 1); });

    var recent = sortBy_(publicMoves.slice(), byMovedDesc_).slice(0, 10).map(plain_);
    var kpi = {
        balance_kg: kgOf_(totalBal),
        month_in_kg: kgOf_(monthIn),
        month_out_kg: kgOf_(monthOut),
        year_in_kg: kgOf_(yearIn),
        year_out_kg: kgOf_(totalYear),
        permit_kg: kgOf_(totalPermit),
        quota_pct: totalPermit > 0 ? Math.round(totalYear / totalPermit * 1000) / 10 : null,
        moves: publicMoves.length
    };
    if (agg) {
      // All-time totals that add up: in = shipped + treated + balance.
      kpi.total_in_kg = kgOf_(totalIn);
      kpi.shipped_kg = kgOf_(totalShip);
      kpi.treated_kg = kgOf_(totalTreat);
      kpi.month_treat_kg = kgOf_(monthTreat);
      kpi.year_treat_kg = kgOf_(yearTreat);
    }
    return {
      kpi: kpi,
      months: months.map(function (m) {
        var o = { ym: m.ym, in_kg: kgOf_(m.in_c), out_kg: kgOf_(m.out_c) };
        if (agg) o.treat_kg = kgOf_(m.treat_c);
        return o;
      }),
      containers: containers,
      receivers: receiverStats_(types.map(plain_), bal).filter(function (r) { return r.active === 1 || r.shipped_year_kg > 0 || r.balance_kg > 0; }).map(function (r) {
        var permitC = Math.round(r.permit_tons * 100000);
        var yearC = cents_(r.shipped_year_kg);
        return {
          id: r.id, receiver: r.name, types: r.active_types, permit_kg: kgOf_(permitC), shipped_year_kg: r.shipped_year_kg,
          balance_kg: r.balance_kg, quota_pct: permitC > 0 ? Math.round(yearC / permitC * 1000) / 10 : null,
          quota_level: quotaLevel_(yearC, permitC)
        };
      }),
      alerts: alerts,
      recent: recent,
      logsToday: table_("audit_logs").rows.filter(function (l) {
        return l.created_at.slice(0, 10) === now.date && !treatLog_(l);
      }).length
    };
  });
  return core;
}

// Audit rows about in-house data (left out of the public dashboard count).
function treatLog_(l) {
  return l.action === "treat" || l.action === "treat_unlock" || /^TR-/.test(String(l.ref || "")); // treat_unlock: older rows
}

function queryHistory_(p) {
  var typeMap = typeMap_();
  var from = phpEmpty_(p.from) ? null : str_(p.from) + " 00:00:00";
  var to = phpEmpty_(p.to) ? null : str_(p.to) + " 23:59:59";
  var type = phpEmpty_(p.type) ? null : phpInt_(p.type);
  var kind = phpEmpty_(p.kind) ? null : str_(p.kind);
  var receiver = phpEmpty_(p.receiver) ? null : str_(p.receiver);
  var q = phpEmpty_(p.q) ? null : likeMatcher_(str_(p.q));
  var rows = visibleMoves_().filter(function (m) {
    var t = typeMap[m.type_id];
    if (!t) return false;
    if (from !== null && !(m.moved_at >= from)) return false;
    if (to !== null && !(m.moved_at <= to)) return false;
    if (type !== null && m.type_id !== type) return false;
    if (kind !== null && m.kind !== kind) return false;
    if (receiver !== null && t.receiver !== receiver) return false;
    if (q && !(q(m.doc_no) || q(m.note) || q(m.actor) || q(m.location) || q(m.manifest_no) || q(m.vehicle) ||
      q(m.method) || q(t.code) || q(t.name))) return false;
    return true;
  });
  return sortBy_(rows, byMovedDesc_).slice(0, 2000).map(function (m) {
    var t = typeMap[m.type_id];
    var o = plain_(m);
    o.code = t.code; o.type_name = t.name; o.receiver = t.receiver; o.mgmt_code = t.mgmt_code;
    o.icon = t.icon; o.color = t.color;
    return o;
  });
}

// "รายการ" is the movement type column; weights in tons. With access (treatOn_) the export
// also has the in-house (treat) rows and a "ระบบ / วิธีการ" column; such a row
// has no receiver (handled inside the plant, not shipped).
function exportHistory_(items) {
  var showTreat = treatOn_();
  var head = ["เลขที่เอกสาร", "วันที่", "เวลา", "รายการ", "รหัสของเสีย", "ชื่อของเสีย", "รหัสการจัดการ",
    "ผู้รับดำเนินการ", "น้ำหนัก (ตัน)", "ภาชนะ", "จำนวนภาชนะ", "สถานที่จัดเก็บ", "เลขที่ใบกำกับการขนส่ง",
    "ทะเบียนรถ", "หมายเหตุ", "ผู้บันทึก"];
  if (showTreat) head.push("ระบบ / วิธีการ");
  var lines = [csvLine_(head)];
  items.forEach(function (r) {
    var treat = r.kind === "treat";
    // Codes keep their leading zeros in Excel (="080409"; validated digits only).
    var line = [r.doc_no, r.moved_at.slice(0, 10), r.moved_at.slice(11, 16), KIND_NAME[r.kind] || r.kind,
      excelText_(r.code), csvText_(r.type_name), excelText_(r.mgmt_code), treat ? "ภายในโรงงาน" : csvText_(r.receiver),
      tonsText_(r.weight_kg, false), CONTAINERS[r.container] || r.container, r.containers, csvText_(r.location),
      csvText_(r.manifest_no), csvText_(r.vehicle), csvText_(r.note), csvText_(r.actor)];
    if (showTreat) line.push(csvText_(r.method));
    lines.push(csvLine_(line));
  });
  return { filename: "waste_thepwatana_history_" + stamp_() + ".csv", csv: lines.join("") };
}

// User text in a CSV cell: never starts like a spreadsheet formula.
function csvText_(v) {
  var s = v == null ? "" : String(v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function excelText_(digits) {
  return /^\d+$/.test(String(digits)) ? '="' + digits + '"' : csvText_(digits);
}

function queryLogs_(p) {
  var from = phpEmpty_(p.from) ? null : str_(p.from) + " 00:00:00";
  var to = phpEmpty_(p.to) ? null : str_(p.to) + " 23:59:59";
  var cat = phpEmpty_(p.cat) ? null : str_(p.cat);
  var q = phpEmpty_(p.q) ? null : likeMatcher_(str_(p.q));
  return table_("audit_logs").rows.filter(function (l) {
    if (from !== null && !(l.created_at >= from)) return false;
    if (to !== null && !(l.created_at <= to)) return false;
    if (cat !== null && !eqCi_(l.category, cat)) return false;
    if (q && !(q(l.detail) || q(l.ref) || q(l.actor))) return false;
    return true;
  }).sort(function (a, b) { return b.id - a.id; }).slice(0, 3000).map(plain_);
}

function exportLogs_(items) {
  var lines = [csvLine_(["ลำดับ", "วันที่", "เวลา", "หมวด", "เลขที่อ้างอิง", "รายละเอียด", "ผู้ทำรายการ"])];
  items.forEach(function (r) {
    lines.push(csvLine_([r.id, r.created_at.slice(0, 10), r.created_at.slice(11, 19), LOG_CAT_NAME[r.category] || r.category,
      r.ref, csvText_(r.detail), csvText_(r.actor)]));
  });
  return { filename: "waste_thepwatana_log_" + stamp_() + ".csv", csv: lines.join("") };
}

/* ===================== write actions ===================== */

// store / ship / opening / treat. All take: type_id, weight_t (tons), container,
// containers, date, time, location, note; ship also manifest_no, vehicle;
// treat also method (ระบบ / วิธีการ). Ship and treat may use an inactive
// type (they only take out what is already in the plant).
function actionMove_(kind, input) {
  var f = moveFields_(kind, input, null);
  var t = getType_(f.type_id);
  if (!isOut_(kind) && t.active !== 1) fail_("ประเภทของเสียนี้ถูกปิดใช้งานแล้ว");
  var doc = nextNo_(PREFIX[kind], "movements", "doc_no");
  var row = insert_("movements", {
    doc_no: doc, kind: kind, type_id: f.type_id, weight_kg: f.weight_kg, container: f.container,
    containers: f.containers, location: f.location, manifest_no: f.manifest_no, vehicle: f.vehicle, note: f.note,
    moved_at: f.moved_at, created_at: nowParts_().datetime, actor: REQ_.actor, updated_at: null, updated_by: "",
    method: f.method
  });
  var verb = kind === "ship" ? "ส่งกำจัดเกินยอดคงเหลือ" : kind === "treat" ? "ข้อมูลภายในเกินยอดคงเหลือ" : "บันทึกไม่ได้";
  validateLedger_(table_("movements").rows, verb, row);
  audit_("waste", kind, describeMove_(row), doc);
  var b = balances_()[f.type_id];
  return {
    id: row.id, doc_no: doc, kind: kind, balance_kg: b.balance_kg,
    message: KIND_NAME[kind] + " " + t.name + " " + fmtT_(f.weight_kg) + " ตัน เรียบร้อย (" + doc + ")"
  };
}

// Validates the fields of a movement. base: the existing row (edit) or null.
function moveFields_(kind, input, base) {
  var pick = function (k) { return base && input[k] === undefined ? base[k] : input[k]; };
  var typeId = phpInt_(pick("type_id"));
  if (!typeMap_()[typeId]) fail_("กรุณาเลือกประเภทของเสีย");
  // weight_t (tons) from the page; an edit without weight_t keeps the stored kg.
  var weight = base && input.weight_t === undefined ? base.weight_kg : tonsToKg_(input.weight_t);
  if (!(weight > 0)) fail_("น้ำหนักต้องมากกว่า 0 ตัน");
  if (weight > MAX_WEIGHT_KG) fail_("น้ำหนักต่อรายการต้องไม่เกิน " + fmtT_(MAX_WEIGHT_KG) + " ตัน");
  var container = phpTrim_(safeStr_(pick("container")));
  if (CONTAINER_KEYS.indexOf(container) === -1) fail_("กรุณาเลือกภาชนะบรรจุ (ถัง 200 ลิตร / IBC / Bigbag)");
  var cRaw = phpTrim_(safeStr_(pick("containers")));
  if (!/^\d{1,6}$/.test(cRaw) || Number(cRaw) < 1) fail_("จำนวนภาชนะต้องเป็นจำนวนเต็มตั้งแต่ 1 ขึ้นไป");
  var at;
  if (base && input.date === undefined && input.time === undefined) at = base.moved_at;
  // edit: an empty / missing date or time keeps the stored one (never silently "now")
  else at = moment_({ date: (input.date == null || input.date === "") && base ? base.moved_at.slice(0, 10) : input.date,
    time: (input.time == null || input.time === "") && base ? base.moved_at.slice(11, 16) : input.time });
  if (parseDt_(at) > Date.now() + TZ_OFFSET_MS + FUTURE_SLACK_MS) fail_("วันที่/เวลาต้องไม่เกินเวลาปัจจุบัน");
  var text = function (k, max, label) {
    var v = pick(k);
    return textIn_({ v: v == null ? "" : v }, "v", max, label);
  };
  var ship = kind === "ship";
  return {
    type_id: typeId, weight_kg: weight, container: container, containers: Number(cRaw), moved_at: at,
    location: text("location", 100, "สถานที่จัดเก็บ"),
    manifest_no: ship ? text("manifest_no", 60, "เลขที่ใบกำกับการขนส่ง") : "",
    vehicle: ship ? text("vehicle", 40, "ทะเบียนรถ") : "",
    method: kind === "treat" ? text("method", 120, "ระบบ / วิธีการ") : "",
    note: text("note", MAX_TEXT, "หมายเหตุ")
  };
}

// Decimal input: number or text like "1,234.50"; at most `decimals` places.
function numIn_(v, label, decimals) {
  if (v === null || v === undefined || v === "" || typeof v === "boolean" || typeof v === "object") fail_("กรุณากรอก" + label);
  var s = typeof v === "number" ? String(v) : phpTrim_(String(v)).replace(/,/g, "");
  if (typeof v === "number" && !isFinite(v)) fail_(label + " ไม่ถูกต้อง");
  if (!/^\d+(\.\d+)?$/.test(s) && !(typeof v === "number" && /^\d+(\.\d+)?(e-?\d+)?$/.test(s))) fail_(label + " ต้องเป็นตัวเลข");
  var n = Number(s);
  if (!isFinite(n)) fail_(label + " ไม่ถูกต้อง");
  if (Math.abs(Math.round(n * Math.pow(10, decimals)) - n * Math.pow(10, decimals)) > 1e-6) {
    fail_(label + " ใส่ทศนิยมได้ไม่เกิน " + decimals + " ตำแหน่ง");
  }
  return Math.round(n * Math.pow(10, decimals)) / Math.pow(10, decimals);
}

function describeMove_(m) {
  var t = typeMap_()[m.type_id] || { code: "?", name: "?", receiver: "" };
  return KIND_NAME[m.kind] + " " + t.code + " " + t.name + " น้ำหนัก " + fmtT_(m.weight_kg) + " ตัน · " +
    (CONTAINERS[m.container] || m.container) + " " + m.containers + " ภาชนะ · เวลา " + dtTh_(m.moved_at) +
    (m.kind === "ship" ? " · ผู้รับ: " + t.receiver : "") +
    (m.location ? " · ที่เก็บ: " + m.location : "") +
    (m.manifest_no ? " · ใบกำกับ: " + m.manifest_no : "") +
    (m.vehicle ? " · ทะเบียนรถ: " + m.vehicle : "") +
    (m.method ? " · ระบบ / วิธีการ: " + m.method : "") +
    (m.note ? " · " + m.note : "");
}

// Wipes all data rows of movements + audit_logs (waste types are kept).
function actionResetData_() {
  var removed = {};
  RESET_TABLES.forEach(function (name) {
    var sheet = db_().getSheetByName(name);
    if (!sheet) throw new Error('ไม่พบชีต "' + name + '" กรุณารัน setupSystem() อีกครั้ง');
    var n = Math.max(0, sheet.getLastRow() - 1);
    if (n > 0) {
      var width = Math.max(sheet.getLastColumn(), SCHEMA[name].length);
      sheet.getRange(2, 1, n, width).clearContent();
    }
    removed[name] = n;
    delete REQ_.tables[name];
    delete REQ_.appends[name];
    REQ_.dirty[name] = true;
  });
  REQ_.updates = REQ_.updates.filter(function (u) { return RESET_TABLES.indexOf(u.name) === -1; });
  REQ_.deletes = REQ_.deletes.filter(function (d) { return RESET_TABLES.indexOf(d.name) === -1; });
  audit_("setting", "reset_data", "รีเซ็ตข้อมูล: ล้างชีต movements (" + removed.movements + " แถว), audit_logs (" +
    removed.audit_logs + " แถว) · คงไว้: waste_types");
  return {
    removed: removed,
    message: "รีเซ็ตข้อมูลแล้ว — ลบรายการเคลื่อนไหว " + removed.movements + " แถว, บันทึกประวัติ " + removed.audit_logs + " แถว"
  };
}

/* ---------- edit / delete (WASTE_RESET_PASSWORD) ---------- */

function findMovement_(input) {
  var id = phpInt_(input.id);
  var m = null;
  table_("movements").rows.forEach(function (r) { if (r.id === id) m = r; });
  if (!m) fail_("ไม่พบรายการ (อาจถูกลบไปแล้ว)");
  if (!phpEmpty_(input.doc_no) && str_(input.doc_no) !== m.doc_no) fail_("ข้อมูลรายการไม่ตรงกัน กรุณาโหลดหน้าใหม่");
  return m;
}

function actionMovementUpdate_(input) {
  var m = findMovement_(input);
  if (input.kind !== undefined && str_(input.kind) !== m.kind) fail_("เปลี่ยนชนิดรายการไม่ได้ — ให้ลบแล้วบันทึกใหม่");
  var f = moveFields_(m.kind, input, m);
  var fields = ["type_id", "weight_kg", "container", "containers", "moved_at", "location", "manifest_no", "vehicle", "method", "note"];
  var changes = [];
  fields.forEach(function (k) { if (f[k] !== m[k]) changes.push(k); });
  if (!changes.length) return { changed: false, message: "ไม่มีการเปลี่ยนแปลง" };
  if (f.type_id !== m.type_id && !isOut_(m.kind) && getType_(f.type_id).active !== 1) fail_("ประเภทของเสียนี้ถูกปิดใช้งานแล้ว");
  var copy = plain_(m);
  fields.forEach(function (k) { copy[k] = f[k]; });
  var moves = table_("movements").rows.map(function (r) { return r === m ? copy : r; });
  validateLedger_(moves, "แก้ไขไม่ได้", copy);
  var before = describeMove_(m);
  var set = { updated_at: nowParts_().datetime, updated_by: REQ_.actor };
  fields.forEach(function (k) { set[k] = f[k]; });
  update_("movements", m, set);
  audit_("edit", "movement_update", "แก้ไขรายการ " + m.doc_no + " — เดิม: " + before + " → ใหม่: " + describeMove_(m), m.doc_no);
  return { changed: true, id: m.id, doc_no: m.doc_no, message: "แก้ไขรายการ " + m.doc_no + " เรียบร้อย" };
}

function actionMovementDelete_(input) {
  var m = findMovement_(input);
  var moves = table_("movements").rows.filter(function (r) { return r !== m; });
  validateLedger_(moves, "ลบไม่ได้", null);
  var detail = "ลบรายการ " + m.doc_no + " — " + describeMove_(m);
  delete_("movements", m);
  audit_("edit", "movement_delete", detail, m.doc_no);
  return { id: m.id, doc_no: m.doc_no, message: "ลบรายการ " + m.doc_no + " แล้ว" };
}

/* ---------- receivers (ผู้รับดำเนินการ) ---------- */

function recvKey_(name) {
  return String(name || "").replace(/\s+/g, " ").trim().toLowerCase();
}

// Receivers: the receivers sheet plus any receiver named by a type that is not
// listed there yet (older data; id 0 until ensureReceivers_ stores it).
function receiverList_() {
  var rows = sortBy_(table_("receivers").rows.slice(), function (a, b) { return a.sort - b.sort || a.id - b.id; }).map(plain_);
  var seen = {};
  rows.forEach(function (r) { seen[recvKey_(r.name)] = true; });
  var next = rows.reduce(function (m, r) { return Math.max(m, r.sort); }, 0);
  sortedTypes_().forEach(function (t) {
    var k = recvKey_(t.receiver);
    if (!k || seen[k]) return;
    seen[k] = true;
    rows.push({ id: 0, name: t.receiver, reg: t.receiver_reg || "", mgmt_code: "", active: 1, sort: ++next });
  });
  return rows;
}

// Stores the derived receivers (inside a write, under the lock). Idempotent.
function ensureReceivers_() {
  var at = nowParts_().datetime;
  receiverList_().forEach(function (r) {
    if (r.id) return;
    insert_("receivers", { name: r.name, reg: r.reg, mgmt_code: "", active: 1, sort: r.sort, updated_at: at, updated_by: REQ_.actor });
  });
}

// Per receiver: types, permitted tons of its ACTIVE types, this year's shipments
// of those types and the in-plant balance. A receiver without types has 0 quota.
function receiverStats_(types, bal) {
  var list = receiverList_();
  var byKey = {};
  list.forEach(function (r) {
    r.types = 0; r.active_types = 0; r.permit_c = 0; r.year_c = 0; r.balance_c = 0;
    byKey[recvKey_(r.name)] = r;
  });
  types.forEach(function (t) {
    var r = byKey[recvKey_(t.receiver)];
    if (!r) return;
    var b = bal && bal[t.id];
    r.types += 1;
    if (b) r.balance_c += cents_(b.balance_kg);
    if (+t.active === 1) {
      r.active_types += 1;
      r.permit_c += Math.round(Number(t.permit_tons || 0) * 100000);
      if (b) r.year_c += cents_(b.shipped_year_kg);
    }
  });
  return list.map(function (r) {
    return {
      id: r.id, name: r.name, reg: r.reg, mgmt_code: r.mgmt_code, active: r.active, types: r.types, active_types: r.active_types,
      permit_tons: r.permit_c / 100000, shipped_year_kg: kgOf_(r.year_c), balance_kg: kgOf_(r.balance_c)
    };
  });
}

/* ---------- waste types (WASTE_RESET_PASSWORD) ---------- */

function actionTypeSave_(input) {
  var id = phpInt_(input.id);
  var existing = null;
  if (id) {
    existing = typeMap_()[id];
    if (!existing) fail_("ไม่พบประเภทของเสียนี้");
  }
  var code = phpTrim_(safeStr_(input.code));
  if (!/^\d{6}$/.test(code)) fail_("รหัสของเสียต้องเป็นตัวเลข 6 หลัก");
  var name = textIn_(input, "name", 120, "ชื่อของเสีย");
  if (!name) fail_("กรุณากรอกชื่อของเสีย");
  var tons = numIn_(input.permit_tons, "ปริมาณที่ได้รับอนุญาต (ตัน)", 2);
  if (tons > MAX_PERMIT_TONS) fail_("ปริมาณที่ได้รับอนุญาตสูงเกินไป");
  ensureReceivers_();
  // The receiver: receiver_id (a listed receiver) or, for older callers, a name
  // (an existing receiver of that name, otherwise a new one is added).
  var rcv = null, newRcv = null;
  var rid = phpInt_(input.receiver_id);
  var givenReg = phpTrim_(safeStr_(input.receiver_reg));
  if (rid) {
    rcv = table_("receivers").rows.filter(function (r) { return r.id === rid; })[0];
    if (!rcv) fail_("ไม่พบผู้รับดำเนินการนี้");
  } else {
    var rname = textIn_(input, "receiver", 120, "ผู้รับดำเนินการ").replace(/\s+/g, " ");
    if (!rname) fail_("กรุณาเลือกหรือกรอกชื่อผู้รับดำเนินการ");
    if (!/^\d{0,20}$/.test(givenReg)) fail_("เลขทะเบียนผู้รับต้องเป็นตัวเลข (ไม่เกิน 20 หลัก)");
    rcv = table_("receivers").rows.filter(function (r) { return recvKey_(r.name) === recvKey_(rname); })[0] || null;
    if (!rcv) newRcv = { name: rname, reg: givenReg };
  }
  var mgmt = phpTrim_(safeStr_(input.mgmt_code));
  if (!mgmt && rcv && rcv.mgmt_code) mgmt = rcv.mgmt_code;
  if (!/^\d{3}$/.test(mgmt)) fail_("รหัสวิธีการจัดการต้องเป็นตัวเลข 3 หลัก");
  var icon = phpTrim_(safeStr_(input.icon)) || (existing ? existing.icon : "fa-recycle");
  if (!/^fa-[a-z0-9-]{1,40}$/.test(icon)) fail_("ไอคอนไม่ถูกต้อง");
  var color = phpTrim_(safeStr_(input.color)) || (existing ? existing.color : "#10B981");
  if (!/^#[0-9a-fA-F]{6}$/.test(color)) fail_("สีไม่ถูกต้อง");
  var active = input.active === undefined ? (existing ? existing.active : 1) : (phpTruthy_(input.active) && str_(input.active) !== "false" ? 1 : 0);
  if (active === 1 && rcv && rcv.active !== 1) fail_("ผู้รับดำเนินการนี้ปิดใช้งานอยู่ — เปิดใช้งานผู้รับก่อน");
  var rName = rcv ? rcv.name : newRcv.name;
  var rReg = rcv ? rcv.reg : newRcv.reg;
  if (rcv && !rcv.reg && !rid && /^\d{1,20}$/.test(givenReg)) rReg = givenReg; // older caller names the registration
  var dup = table_("waste_types").rows.some(function (t) {
    return t !== existing && t.code === code && t.mgmt_code === mgmt && t.receiver_reg === rReg && t.name === name;
  });
  if (dup) fail_("มีประเภทของเสียนี้ (รหัส ชื่อ วิธีจัดการ และผู้รับเดียวกัน) อยู่แล้ว");
  var at = nowParts_().datetime;
  var createdReceiver = false;
  if (newRcv) {
    var maxRs = 0;
    table_("receivers").rows.forEach(function (r) { if (r.sort > maxRs) maxRs = r.sort; });
    rcv = insert_("receivers", { name: newRcv.name, reg: newRcv.reg, mgmt_code: "", active: 1, sort: maxRs + 1, updated_at: at, updated_by: REQ_.actor });
    audit_("setting", "receiver_save", "เพิ่มผู้รับดำเนินการ #" + rcv.id + ": " + rcv.name + (rcv.reg ? " (" + rcv.reg + ")" : ""), "R" + rcv.id);
    createdReceiver = true;
  } else if (rcv.reg !== rReg) {
    update_("receivers", rcv, { reg: rReg, updated_at: at, updated_by: REQ_.actor });
  }
  var fields = {
    code: code, name: name, permit_tons: tons, mgmt_code: mgmt, receiver_reg: rReg, receiver: rName,
    icon: icon, color: color.toUpperCase(), active: active, updated_at: at, updated_by: REQ_.actor
  };
  var desc = code + " " + name + " · อนุญาต " + tons + " ตัน · รหัสจัดการ " + mgmt + " · " + rName +
    (rReg ? " (" + rReg + ")" : "") + " · " + (active ? "ใช้งาน" : "ปิดใช้งาน");
  if (existing) {
    var before = existing.code + " " + existing.name + " · อนุญาต " + existing.permit_tons + " ตัน · รหัสจัดการ " +
      existing.mgmt_code + " · " + existing.receiver + (existing.receiver_reg ? " (" + existing.receiver_reg + ")" : "") +
      " · " + (existing.active ? "ใช้งาน" : "ปิดใช้งาน");
    update_("waste_types", existing, fields);
    audit_("setting", "type_save", "แก้ไขประเภทของเสีย #" + existing.id + " — เดิม: " + before + " → ใหม่: " + desc, "T" + existing.id);
    return { id: existing.id, created: false, receiver_id: rcv.id, receiver_created: createdReceiver, message: "บันทึกประเภทของเสีย " + name + " แล้ว" };
  }
  var maxSort = 0;
  table_("waste_types").rows.forEach(function (t) { if (t.sort > maxSort) maxSort = t.sort; });
  fields.sort = maxSort + 1;
  var row = insert_("waste_types", fields);
  audit_("setting", "type_save", "เพิ่มประเภทของเสีย #" + row.id + ": " + desc, "T" + row.id);
  return { id: row.id, created: true, receiver_id: rcv.id, receiver_created: createdReceiver, message: "เพิ่มประเภทของเสีย " + name + " แล้ว" };
}

// A type that has movements can only be deactivated (history keeps its name).
function actionTypeDelete_(input) {
  var t = typeMap_()[phpInt_(input.id)];
  if (!t) fail_("ไม่พบประเภทของเสียนี้");
  var used = table_("movements").rows.some(function (m) { return m.type_id === t.id; });
  if (used) fail_("ประเภทของเสียนี้มีรายการเคลื่อนไหวแล้ว ลบไม่ได้ — ให้ปิดใช้งานแทน");
  delete_("waste_types", t);
  audit_("setting", "type_delete", "ลบประเภทของเสีย #" + t.id + ": " + t.code + " " + t.name + " · " + t.receiver, "T" + t.id);
  return { id: t.id, message: "ลบประเภทของเสีย " + t.name + " แล้ว" };
}

function actionReceiverSave_(input) {
  ensureReceivers_();
  var rows = table_("receivers").rows;
  var existing = null;
  var id = phpInt_(input.id);
  var cur = phpEmpty_(input.current_name) ? "" : recvKey_(safeStr_(input.current_name));
  if (id) existing = rows.filter(function (r) { return r.id === id; })[0] || null;
  else if (cur) existing = rows.filter(function (r) { return recvKey_(r.name) === cur; })[0] || null;
  if ((id || cur) && !existing) fail_("ไม่พบผู้รับดำเนินการนี้");
  var name = textIn_(input, "name", 120, "ชื่อผู้รับดำเนินการ").replace(/\s+/g, " ");
  if (!name) fail_("กรุณากรอกชื่อผู้รับดำเนินการ");
  var reg = phpTrim_(safeStr_(input.reg !== undefined ? input.reg : input.receiver_reg));
  if (!/^\d{0,20}$/.test(reg)) fail_("เลขทะเบียนผู้รับต้องเป็นตัวเลข (ไม่เกิน 20 หลัก)");
  var mgmt = phpTrim_(safeStr_(input.mgmt_code));
  if (!/^(\d{3})?$/.test(mgmt)) fail_("รหัสวิธีการจัดการต้องเป็นตัวเลข 3 หลัก (หรือเว้นว่าง)");
  if (rows.some(function (r) { return r !== existing && recvKey_(r.name) === recvKey_(name); })) fail_("มีผู้รับดำเนินการชื่อนี้อยู่แล้ว");
  var active = input.active === undefined ? (existing ? existing.active : 1) : (phpTruthy_(input.active) && str_(input.active) !== "false" ? 1 : 0);
  var at = nowParts_().datetime;
  if (existing) {
    var typesOf = table_("waste_types").rows.filter(function (t) { return recvKey_(t.receiver) === recvKey_(existing.name); });
    if (active === 0 && typesOf.some(function (t) { return t.active === 1; })) {
      fail_("ผู้รับนี้ยังมีประเภทของเสียที่เปิดใช้งานอยู่ — ปิดใช้งานประเภทเหล่านั้นก่อน");
    }
    var before = existing.name + (existing.reg ? " (" + existing.reg + ")" : "") + " · " + (existing.active ? "ใช้งาน" : "ปิดใช้งาน");
    update_("receivers", existing, { name: name, reg: reg, mgmt_code: mgmt, active: active, updated_at: at, updated_by: REQ_.actor });
    typesOf.forEach(function (t) {
      if (t.receiver !== name || t.receiver_reg !== reg) update_("waste_types", t, { receiver: name, receiver_reg: reg, updated_at: at, updated_by: REQ_.actor });
    });
    audit_("setting", "receiver_save", "แก้ไขผู้รับดำเนินการ #" + existing.id + " — เดิม: " + before + " → ใหม่: " + name +
      (reg ? " (" + reg + ")" : "") + " · " + (active ? "ใช้งาน" : "ปิดใช้งาน") + " · " + typesOf.length + " ประเภท", "R" + existing.id);
    return { id: existing.id, created: false, message: "บันทึกผู้รับดำเนินการ " + name + " แล้ว" };
  }
  var maxSort = 0;
  rows.forEach(function (r) { if (r.sort > maxSort) maxSort = r.sort; });
  var row = insert_("receivers", { name: name, reg: reg, mgmt_code: mgmt, active: active, sort: maxSort + 1, updated_at: at, updated_by: REQ_.actor });
  audit_("setting", "receiver_save", "เพิ่มผู้รับดำเนินการ #" + row.id + ": " + name + (reg ? " (" + reg + ")" : ""), "R" + row.id);
  return { id: row.id, created: true, message: "เพิ่มผู้รับดำเนินการ " + name + " แล้ว" };
}

// Only a receiver without any waste type can be removed (a type that has
// movements stays, so its receiver is deactivated instead).
function actionReceiverDelete_(input) {
  ensureReceivers_();
  var id = phpInt_(input.id);
  var r = table_("receivers").rows.filter(function (x) { return x.id === id; })[0];
  if (!r) fail_("ไม่พบผู้รับดำเนินการนี้");
  if (table_("waste_types").rows.some(function (t) { return recvKey_(t.receiver) === recvKey_(r.name); })) {
    fail_("ผู้รับนี้ยังมีประเภทของเสียอยู่ ลบไม่ได้ — ให้ปิดใช้งานแทน");
  }
  delete_("receivers", r);
  audit_("setting", "receiver_delete", "ลบผู้รับดำเนินการ #" + r.id + ": " + r.name, "R" + r.id);
  return { id: r.id, message: "ลบผู้รับดำเนินการ " + r.name + " แล้ว" };
}

/* ===================== business helpers ===================== */

function getType_(id) {
  var t = typeMap_()[id];
  if (!t) fail_("ไม่พบประเภทของเสีย");
  return t;
}

function typeMap_() {
  var map = {};
  table_("waste_types").rows.forEach(function (t) { map[t.id] = t; });
  return map;
}

// date/time from the client (default: now), "YYYY-MM-DD" + "HH:MM".
function moment_(input) {
  var now = nowParts_();
  var d = input.date == null || input.date === "" ? now.date : safeStr_(input.date);
  var t = input.time == null || input.time === "" ? now.time : safeStr_(input.time);
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2}) (\d{1,2}):(\d{2})$/.exec(d + " " + t.slice(0, 5));
  if (!m) fail_("วันที่/เวลาไม่ถูกต้อง");
  var Y = Number(m[1]), M = Number(m[2]), D = Number(m[3]), h = Number(m[4]), mi = Number(m[5]);
  var dt = new Date(Date.UTC(Y, M - 1, D, h, mi, 0));
  if (Y < 2000 || dt.getUTCMonth() !== M - 1 || dt.getUTCDate() !== D || h > 23 || mi > 59) fail_("วันที่/เวลาไม่ถูกต้อง");
  return formatUtcParts_(dt).datetime.slice(0, 16) + ":00";
}

// Next running number for today's tag (audit refs scanned too, so a number
// freed by a delete is never handed out again).
function nextNo_(prefix, table, col) {
  var tag = prefix + "-" + nowParts_().ymd + "-";
  var max = 0;
  var scan = function (v) {
    var s = String(v || "").toUpperCase();
    if (s.indexOf(tag) !== 0) return;
    var n = parseInt(s.slice(tag.length), 10);
    if (n > max) max = n;
  };
  table_(table).rows.forEach(function (r) { scan(r[col]); });
  table_("audit_logs").rows.forEach(function (l) { scan(l.ref); });
  var next = max + 1;
  return tag + ("0000" + next).slice(-Math.max(4, String(next).length));
}

function audit_(category, action, detail, ref) {
  insert_("audit_logs", {
    category: category, action: action, ref: ref || "",
    detail: mbSubstr_(detail, 0, 500),
    actor: REQ_.actor || "",
    created_at: nowParts_().datetime
  });
}

function textIn_(input, key, max, label) {
  var value = phpTrim_(safeStr_(input[key]));
  if (mbLen_(value) > max) fail_(label + "ยาวเกิน " + max + " ตัวอักษร");
  return value;
}

/* Tons <-> kg. The sheet keeps kg with 0.01 kg precision, so a ton value has
 * at most 5 decimals; the conversion works on whole hundredths of a kg. */

// weight_t from the page ("0.125", "1,234.5", 2) -> kg (0.125 t -> 125 kg).
function tonsToKg_(v) {
  var t = numIn_(v, "น้ำหนัก (ตัน)", 5);
  return Math.round(t * 100000) / 100;
}

// kg -> tons text: thousands separators, 2 decimals (more only when the kg
// amount has a fraction, e.g. older rows: 1234.56 kg -> "1.23456").
function tonsText_(kg, sep) {
  var c = Math.round(Number(kg || 0) * 100);
  var neg = c < 0;
  c = Math.abs(c);
  var whole = String(Math.floor(c / 100000));
  var frac = String(c % 100000);
  while (frac.length < 5) frac = "0" + frac;
  frac = frac.replace(/0+$/, "");
  while (frac.length < 2) frac += "0";
  if (sep) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return (neg ? "-" : "") + whole + "." + frac;
}

function fmtT_(kg) {
  return tonsText_(kg, true);
}

function dtTh_(s) {
  return s.slice(8, 10) + "/" + s.slice(5, 7) + "/" + (Number(s.slice(0, 4)) + 543) + " " + s.slice(11, 16) + " น.";
}

/* ===================== sheet storage ===================== */

var REQ_ = null; // per-request state: spreadsheet, loaded tables, pending writes

// actor: the name typed in the page header ("" when none).
function resetRequest_(actor) {
  REQ_ = {
    ss: null, tables: {}, appends: {}, updates: [], deletes: [], actor: actor || "",
    dirty: {},          // sheets changed outside the buffers (reset_data) -> version bump on flush
    cacheReads: false,  // true only while answering read actions
    refreshCache: false, // keepWarm: read the sheets live and store them again (fresh TTL)
    vers: null          // data versions read for this request (see versions_)
  };
}

function db_() {
  if (!REQ_.ss) {
    var id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY);
    if (!id) throw new Error("ยังไม่ได้ตั้งค่าระบบ กรุณารัน setupSystem() ใน Apps Script ก่อน");
    REQ_.ss = SpreadsheetApp.openById(id);
  }
  return REQ_.ss;
}

// Reads the whole sheet once per request and converts rows to typed objects.
// Sheets in AUTO_CREATE_TABLES are created when missing, and columns added to
// SCHEMA after a sheet was created are appended to its header row (existing
// rows read them as empty), so an older live spreadsheet upgrades itself.
function table_(name) {
  if (REQ_.tables[name]) return REQ_.tables[name];
  var cacheable = REQ_.cacheReads && CACHED_TABLES.indexOf(name) !== -1;
  var ver = null;
  if (cacheable) {
    ver = versions_()[name] || null; // taken BEFORE the sheet is read
    var hit = ver && !REQ_.refreshCache ? cachedRows_(name, ver) : null;
    if (hit) return (REQ_.tables[name] = hit);
  }
  var sheet = db_().getSheetByName(name);
  if (!sheet && AUTO_CREATE_TABLES.indexOf(name) !== -1) sheet = prepareSheet_(db_(), name, null);
  if (!sheet) throw new Error('ไม่พบชีต "' + name + '" กรุณารัน setupSystem() อีกครั้ง');
  var values = sheet.getDataRange().getValues();
  var header = (values[0] || []).map(function (h) { return String(h).trim(); });
  var cols = SCHEMA[name];
  var idx = cols.map(function (c) { return header.indexOf(c[0]); });
  if (idx.indexOf(-1) !== -1) {
    if (!header.some(function (h) { return h !== ""; })) {
      throw new Error('หัวตารางของชีต "' + name + '" ไม่ครบ กรุณารัน setupSystem()');
    }
    header = addMissingColumns_(sheet, name, header);
    idx = cols.map(function (c) { return header.indexOf(c[0]); });
  }
  var rows = [];
  var maxId = 0;
  for (var r = 1; r < values.length; r++) {
    var raw = values[r];
    var empty = true;
    for (var k = 0; k < raw.length; k++) if (raw[k] !== "" && raw[k] !== null) { empty = false; break; }
    if (empty) continue;
    var obj = {};
    cols.forEach(function (col, i) { obj[col[0]] = fromCell_(col[1], raw[idx[i]]); });
    Object.defineProperty(obj, "_row", { value: r + 1, writable: true, enumerable: false });
    if (obj.id > maxId) maxId = obj.id;
    rows.push(obj);
  }
  // values: the raw sheet rows, so flush_ can update a row without reading it again.
  REQ_.tables[name] = { name: name, sheet: sheet, header: header, idx: idx, rows: rows, maxId: maxId, values: values };
  if (ver) storeRows_(name, ver, rows);
  return REQ_.tables[name];
}

/* ---------- read cache helpers (see CACHED_TABLES) ---------- */

function newVersion_() {
  return Date.now().toString(36) + "." + Utilities.getUuid().replace(/-/g, "").slice(0, 12);
}

// The current data version of every cached sheet (one CacheService call per
// request). A missing version (evicted / first use) gets a new random one, so
// nothing stored under an older version can match it. Cache errors disable the
// read cache for the request ({}).
function versions_() {
  if (REQ_.vers) return REQ_.vers;
  var vers = {};
  try {
    var keys = CACHED_TABLES.map(function (n) { return VERSION_PREFIX + n; });
    var got = cache_().getAll(keys) || {};
    var missing = {};
    CACHED_TABLES.forEach(function (n) {
      var v = got[VERSION_PREFIX + n];
      if (!v) { v = newVersion_(); missing[VERSION_PREFIX + n] = v; }
      vers[n] = v;
    });
    if (Object.keys(missing).length) cache_().putAll(missing, VERSION_TTL);
  } catch (error) {
    vers = {};
  }
  REQ_.vers = vers;
  return vers;
}

// Gives every named (cached) sheet a new data version. Called after the
// changes are flushed to the spreadsheet. If the cache cannot be updated the
// versions are removed instead (readers then start new ones).
function bumpVersions_(names) {
  var map = {};
  names.forEach(function (n) { if (CACHED_TABLES.indexOf(n) !== -1) map[VERSION_PREFIX + n] = newVersion_(); });
  var keys = Object.keys(map);
  if (!keys.length) return;
  try {
    cache_().putAll(map, VERSION_TTL);
  } catch (error) {
    try { cache_().removeAll(keys); } catch (ignored) {}
  }
  if (REQ_ && REQ_.vers) REQ_.vers = null; // re-read on the next cached read of this request
}

// Stores text under key in chunks of RC_CHUNK characters: key holds "<n>|<chunk 0>",
// key_1..key_<n-1> the rest. Too big (> RC_MAX_CHUNKS) or cache errors: not stored.
function rcPut_(key, text) {
  var ttl = REQ_ && REQ_.refreshCache ? KEEP_WARM_TTL : READ_CACHE_TTL;
  var n = Math.max(1, Math.ceil(text.length / RC_CHUNK));
  if (n > RC_MAX_CHUNKS) return;
  var map = {};
  for (var i = 0; i < n; i++) {
    var part = text.substr(i * RC_CHUNK, RC_CHUNK);
    map[i ? key + "_" + i : key] = i ? part : n + "|" + part;
  }
  try { cache_().putAll(map, ttl); } catch (ignored) {}
}

function rcGet_(key) {
  try {
    var head = cache_().get(key);
    if (!head) return null;
    var bar = head.indexOf("|");
    var n = parseInt(head.slice(0, bar), 10);
    if (!(n >= 1)) return null;
    var parts = [head.slice(bar + 1)];
    if (n > 1) {
      var keys = [];
      for (var i = 1; i < n; i++) keys.push(key + "_" + i);
      var got = cache_().getAll(keys) || {};
      for (var j = 0; j < keys.length; j++) {
        if (typeof got[keys[j]] !== "string") return null;
        parts.push(got[keys[j]]);
      }
    }
    return parts.join("");
  } catch (error) {
    return null;
  }
}

// Rows are stored as arrays in SCHEMA order, normalised exactly like a sheet
// round trip (toCell_ -> fromCell_), so cached and live reads are identical.
function storeRows_(name, ver, rows) {
  var cols = SCHEMA[name];
  var data = rows.map(function (r) {
    return cols.map(function (c) { return fromCell_(c[1], toCell_(c[1], r[c[0]])); });
  });
  rcPut_(RC_PREFIX + name + "_" + ver, JSON.stringify(data));
}

function cachedRows_(name, ver) {
  var text = rcGet_(RC_PREFIX + name + "_" + ver);
  if (!text) return null;
  var data;
  try { data = JSON.parse(text); } catch (error) { return null; }
  var cols = SCHEMA[name];
  var maxId = 0;
  var rows = data.map(function (a) {
    var obj = {};
    cols.forEach(function (c, i) { obj[c[0]] = a[i]; });
    Object.defineProperty(obj, "_row", { value: 0, writable: true, enumerable: false });
    if (obj.id > maxId) maxId = obj.id;
    return obj;
  });
  // No sheet: tables from the cache are read-only (writes always read live).
  return { name: name, sheet: null, header: null, idx: null, rows: rows, maxId: maxId, values: null, cached: true };
}

// Result of fn() cached under the versions of the sheets it depends on.
function cachedDerived_(id, deps, fn) {
  if (!REQ_.cacheReads) return fn();
  var vers = versions_();
  var parts = [];
  for (var i = 0; i < deps.length; i++) {
    if (!vers[deps[i]]) return fn();
    parts.push(vers[deps[i]]);
  }
  var key = RC_PREFIX + id + "_" + parts.join("_");
  var text = REQ_.refreshCache ? null : rcGet_(key);
  if (text) {
    try { return JSON.parse(text); } catch (ignored) {}
  }
  var value = fn();
  rcPut_(key, JSON.stringify(value));
  return value;
}

/**
 * Run from the Apps Script editor after editing the spreadsheet by hand: gives
 * every cached sheet a new data version, so the next reads load the sheets.
 */
function clearReadCache() {
  bumpVersions_(CACHED_TABLES);
  return { cleared: CACHED_TABLES };
}

/* ---------- keep-warm (time-driven trigger) ----------
 * keepWarm() runs the reads the pages open with (bootstrap, dashboard)
 * the way a web request would, but reads every sheet live and stores
 * it again in the read cache under the CURRENT data versions with a TTL longer
 * than the trigger period, so the next web request is answered from
 * CacheService. It only reads: no sheet row is written, no version is bumped,
 * no lock is taken (a missing version is created exactly as by any read).
 * Hand edits in the spreadsheet are picked up at each run.
 * Run installKeepWarmTrigger() once from the Apps Script editor.
 */
function keepWarm() {
  try {
    resetRequest_("");
    REQ_.cacheReads = true;
    REQ_.refreshCache = true;
    handleRead_("bootstrap", {});
    handleRead_("dashboard", {});
    return { warmed: true };
  } catch (error) {
    console.log("keepWarm: " + errorMessage_(error));
    return { warmed: false, error: errorMessage_(error) };
  } finally {
    REQ_ = null;
  }
}

/** Run once from the editor: (re)creates the one keepWarm trigger (every 10 minutes). */
function installKeepWarmTrigger() {
  removeKeepWarmTrigger();
  ScriptApp.newTrigger("keepWarm").timeBased().everyMinutes(KEEP_WARM_MINUTES).create();
  var warm = keepWarm();
  var result = { installed: true, everyMinutes: KEEP_WARM_MINUTES, warmed: warm.warmed };
  console.log(JSON.stringify(result));
  return result;
}

/** Deletes every keepWarm trigger of this project (other triggers are left alone). */
function removeKeepWarmTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "keepWarm") { ScriptApp.deleteTrigger(t); removed++; }
  });
  return { removed: removed };
}

function insert_(name, obj) {
  var t = table_(name);
  var row = {};
  SCHEMA[name].forEach(function (col) { row[col[0]] = obj[col[0]] === undefined ? null : obj[col[0]]; });
  t.maxId += 1;
  row.id = t.maxId;
  Object.defineProperty(row, "_row", { value: 0, writable: true, enumerable: false });
  t.rows.push(row);
  (REQ_.appends[name] = REQ_.appends[name] || []).push(row);
  return row;
}

function update_(name, row, fields) {
  Object.keys(fields).forEach(function (k) { row[k] = fields[k]; });
  if (row._row > 0) {
    for (var i = 0; i < REQ_.updates.length; i++) if (REQ_.updates[i].row === row) return;
    REQ_.updates.push({ name: name, row: row });
  } // rows inserted in this request are written with their latest values on flush
}

// Removes a row from the in-memory table and buffers the sheet row deletion.
function delete_(name, row) {
  var t = table_(name);
  var i = t.rows.indexOf(row);
  if (i !== -1) t.rows.splice(i, 1);
  REQ_.updates = REQ_.updates.filter(function (u) { return u.row !== row; });
  if (row._row > 0) {
    REQ_.deletes.push({ name: name, row: row });
  } else if (REQ_.appends[name]) {
    REQ_.appends[name] = REQ_.appends[name].filter(function (r) { return r !== row; });
  }
}

// Writes buffered changes in a safe order: updates (row numbers are still the
// ones read at the start of the request), then deletions (highest row first,
// each run of consecutive rows in one deleteRows call), then appends (one
// setNumberFormats + one setValues per table). Tables that lost rows are
// dropped from the per-request cache because their row numbers have shifted.
// Afterwards every changed sheet gets a new data version (read cache), also when
// a write fails half-way.
function flush_() {
  var changed = Object.keys(REQ_.dirty);
  REQ_.updates.forEach(function (u) { changed.push(u.name); });
  REQ_.deletes.forEach(function (d) { changed.push(d.name); });
  Object.keys(REQ_.appends).forEach(function (n) { if (REQ_.appends[n].length) changed.push(n); });
  REQ_.dirty = {};
  try {
    flushBuffers_();
  } finally {
    if (changed.length) {
      // Make the writes visible to other executions before announcing them.
      try { SpreadsheetApp.flush(); } catch (ignored) {}
      bumpVersions_(changed.filter(function (n, i, a) { return a.indexOf(n) === i; }));
    }
  }
}

function flushBuffers_() {
  REQ_.updates.forEach(function (u) {
    var t = REQ_.tables[u.name];
    var range = t.sheet.getRange(u.row._row, 1, 1, t.header.length);
    // The row as read at the start of this request (under the lock), else read it now.
    var current = t.values && t.values[u.row._row - 1] ? t.values[u.row._row - 1] : range.getValues()[0];
    var next = rowValues_(u.name, t, u.row, t.header.length, current);
    range.setValues([next]);
    if (t.values && t.values[u.row._row - 1]) t.values[u.row._row - 1] = next;
  });
  REQ_.updates = [];

  var byTable = {};
  REQ_.deletes.forEach(function (d) { (byTable[d.name] = byTable[d.name] || []).push(d.row._row); });
  Object.keys(byTable).forEach(function (name) {
    var sheet = REQ_.tables[name].sheet;
    var rows = byTable[name]
      .filter(function (r, i, a) { return r > 1 && a.indexOf(r) === i; })
      .sort(function (a, b) { return b - a; });
    // Sheets refuses to delete every non-frozen row: keep one spare row.
    if (sheet.getMaxRows() - rows.length < 2) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    var i = 0;
    while (i < rows.length) {
      var j = i;
      while (j + 1 < rows.length && rows[j + 1] === rows[j] - 1) j++;
      sheet.deleteRows(rows[j], j - i + 1); // rows[j] = lowest row number of the run
      i = j + 1;
    }
  });
  REQ_.deletes = [];

  Object.keys(REQ_.appends).forEach(function (name) {
    var list = REQ_.appends[name];
    if (!list.length) return;
    var t = REQ_.tables[name];
    var width = t.header.length;
    var start = Math.max(t.sheet.getLastRow(), 1) + 1;
    var needed = start + list.length - 1;
    var maxRows = t.sheet.getMaxRows();
    if (needed > maxRows) t.sheet.insertRowsAfter(maxRows, needed - maxRows);
    var formats = [];
    var values = list.map(function (row) {
      formats.push(rowFormats_(name, t));
      row._row = start + formats.length - 1;
      return rowValues_(name, t, row, width);
    });
    var range = t.sheet.getRange(start, 1, list.length, width);
    range.setNumberFormats(formats);
    range.setValues(values);
  });
  REQ_.appends = {};
  Object.keys(byTable).forEach(function (name) { delete REQ_.tables[name]; });
}

// Appends the SCHEMA columns missing from the header row (written at once, not
// buffered; plain-text / number format for the whole new columns). Returns the
// new header.
function addMissingColumns_(sheet, name, header) {
  while (header.length && header[header.length - 1] === "") header.pop();
  var missing = SCHEMA[name].filter(function (c) { return header.indexOf(c[0]) === -1; });
  if (!missing.length) return header;
  var start = header.length + 1;
  var needCols = start + missing.length - 1;
  var maxCols = sheet.getMaxColumns();
  if (needCols > maxCols) sheet.insertColumnsAfter(maxCols, needCols - maxCols);
  sheet.getRange(1, start, 1, missing.length).setValues([missing.map(function (c) { return c[0]; })]);
  var maxRows = sheet.getMaxRows();
  var rowFormat = missing.map(function (c) { return cellFormat_(c[1]); });
  var formats = [];
  for (var r = 0; r < maxRows; r++) formats.push(rowFormat);
  sheet.getRange(1, start, maxRows, missing.length).setNumberFormats(formats);
  sheet.getRange(1, start, 1, missing.length).setFontWeight("bold");
  bumpVersions_([name]);
  return header.concat(missing.map(function (c) { return c[0]; }));
}

function rowValues_(name, t, row, width, base) {
  var out = base ? base.slice() : [];
  while (out.length < width) out.push("");
  SCHEMA[name].forEach(function (col, i) { out[t.idx[i]] = toCell_(col[1], row[col[0]]); });
  return out;
}

function rowFormats_(name, t) {
  var out = [];
  for (var i = 0; i < t.header.length; i++) out.push("@");
  SCHEMA[name].forEach(function (col, i) { out[t.idx[i]] = cellFormat_(col[1]); });
  return out;
}

// Sheet number format of a column type: whole numbers, decimals or plain text.
function cellFormat_(type) {
  if (type.indexOf("int") === 0) return "0";
  if (type.indexOf("num") === 0) return "0.00#";
  return "@";
}

function toCell_(type, v) {
  if (v === null || v === undefined) return "";
  if (type.indexOf("int") === 0 || type.indexOf("num") === 0) return Number(v);
  var s = String(v);
  // Never let user text become a formula (cells are also plain-text formatted).
  if (s.charAt(0) === "=") s = "'" + s;
  return s;
}

function fromCell_(type, v) {
  var nullable = type.charAt(type.length - 1) === "?";
  var base = nullable ? type.slice(0, -1) : type;
  if (v === "" || v === null || v === undefined) return nullable ? null : (base === "int" || base === "num" ? 0 : "");
  if (base === "num") {
    var x = Number(v);
    return isFinite(x) ? Math.round(x * 1000) / 1000 : 0;
  }
  if (base === "int") {
    var n = Number(v);
    return isFinite(n) ? Math.trunc(n) : 0;
  }
  if (Object.prototype.toString.call(v) === "[object Date]") {
    // A cell Sheets converted to a Date (spreadsheet time zone is Asia/Bangkok).
    return formatUtcParts_(new Date(v.getTime() + TZ_OFFSET_MS)).datetime;
  }
  var s = String(v);
  if (s.charAt(0) === "'" && s.charAt(1) === "=") s = s.slice(1);
  return s;
}

/* ===================== setup ===================== */

// Inserts the permit receivers + waste types (inside a request, buffered; the caller flushes).
// Does nothing when the waste_types sheet already has rows. Never touches movements.
function seedPermit_() {
  if (table_("waste_types").rows.length) return { types: 0, receivers: 0, permit_tons: 0 };
  var at = nowParts_().datetime;
  var haveRcv = {};
  table_("receivers").rows.forEach(function (r) { haveRcv[recvKey_(r.name)] = true; });
  var maxRs = table_("receivers").rows.reduce(function (m, r) { return Math.max(m, r.sort); }, 0);
  var rcv = 0;
  PERMIT_RECEIVERS.forEach(function (r) {
    if (haveRcv[recvKey_(r[0])]) return;
    insert_("receivers", { name: r[0], reg: r[1], mgmt_code: "", active: 1, sort: ++maxRs, updated_at: at, updated_by: REQ_.actor });
    rcv++;
  });
  var cents = 0;
  PERMIT_LINES.forEach(function (t, i) {
    insert_("waste_types", {
      code: t[0], name: t[1], permit_tons: t[2], mgmt_code: t[3], receiver_reg: t[4], receiver: t[5],
      icon: t[6], color: t[7], active: 1, sort: i + 1, updated_at: at, updated_by: REQ_.actor
    });
    cents += Math.round(t[2] * 100);
  });
  return { types: PERMIT_LINES.length, receivers: rcv, permit_tons: cents / 100 };
}

/**
 * Run ONCE from the Apps Script editor on an existing EMPTY database: adds the 40
 * waste types and 8 receivers of สก.2 2569-12560 (1,101.55 t). Refuses (and changes
 * nothing) when the waste_types sheet already has rows; never touches movements.
 */
function seedThepwatanaPermit() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    resetRequest_("seedThepwatanaPermit");
    REQ_.actor = "seedThepwatanaPermit";
    var existing = table_("waste_types").rows.length;
    if (existing) {
      var refused = { seeded: false, existingTypes: existing, message: "ไม่ได้เพิ่มข้อมูล: ชีต waste_types มีข้อมูลอยู่แล้ว " + existing + " แถว (รันได้เฉพาะฐานข้อมูลที่ยังว่าง)" };
      console.log(refused.message);
      REQ_ = null;
      return refused;
    }
    var r = seedPermit_();
    flush_();
    REQ_ = null;
    var result = { seeded: true, types: r.types, receivers: r.receivers, permit_tons: r.permit_tons,
      message: "เพิ่มข้อมูลสำเร็จ: " + r.types + " ประเภท, " + r.receivers + " ผู้รับดำเนินการ, อนุญาตรวม " + r.permit_tons + " ตัน" };
    console.log(result.message);
    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * One-time setup. Run from the Apps Script editor and approve permissions.
 * If WASTE_SPREADSHEET_ID is already set, the existing spreadsheet is reused
 * (missing sheets/headers are added; existing data is never seeded or changed) instead of creating another.
 */
function setupSystem() {
  var properties = PropertiesService.getScriptProperties();
  var id = properties.getProperty(SPREADSHEET_ID_PROPERTY);
  var ss;
  var created = false;
  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (error) {
      throw new Error("เปิดสเปรดชีตเดิม (" + id + ") ไม่ได้: " + errorMessage_(error) +
        " — ถ้าต้องการสร้างใหม่ ให้ลบ Script Property " + SPREADSHEET_ID_PROPERTY + " แล้วรันอีกครั้ง");
    }
  } else {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    ss.setSpreadsheetTimeZone("Asia/Bangkok");
    created = true;
    properties.setProperty(SPREADSHEET_ID_PROPERTY, ss.getId());
  }

  var defaultSheets = created ? ss.getSheets() : [];
  TABLE_ORDER.forEach(function (name, i) {
    var sheet = ss.getSheetByName(name);
    if (!sheet && created && i === 0 && defaultSheets.length) {
      sheet = defaultSheets[0];
      sheet.setName(name);
    }
    prepareSheet_(ss, name, sheet);
  });

  resetRequest_("setupSystem");
  REQ_.ss = ss;
  TABLE_ORDER.forEach(function (name) { table_(name); REQ_.dirty[name] = true; });
  REQ_.actor = "setupSystem";
  var seeded = created && SEED_ON_SETUP ? seedPermit_().types : 0;
  ensureReceivers_();
  flush_();
  REQ_ = null;

  var passwords = {};
  var missing = [];
  ROLE_ORDER.forEach(function (r) {
    passwords[ROLE_PROPERTY[r]] = Boolean(properties.getProperty(ROLE_PROPERTY[r]));
    if (!passwords[ROLE_PROPERTY[r]]) missing.push(ROLE_PROPERTY[r] + " (" + ROLE_NAME[r] + ")");
  });
  var result = {
    created: created,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    seededTypes: seeded,
    resetPasswordConfigured: passwords[RESET_PASSWORD_PROPERTY],
    passwords: passwords,
    passwordMessage: missing.length
      ? "ยังไม่ได้ตั้ง Script Property: " + missing.join(", ") + " — ระดับนั้นจะใส่รหัสปลดล็อกไม่ได้จนกว่าจะตั้งค่า (Project Settings > Script properties)"
      : "ตั้งรหัสผ่านครบทั้ง 3 ระดับแล้ว"
  };
  console.log(JSON.stringify(result));
  console.log(result.passwordMessage);
  return result;
}

// Creates the sheet when missing, writes the header row on an empty sheet,
// bolds/freezes it and sets number / plain-text formats per column so Sheets
// never converts values (e.g. "080409" into 80409 or a date text into a Date).
function prepareSheet_(ss, name, sheet) {
  if (!sheet) sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  var cols = SCHEMA[name];
  var headers = cols.map(function (c) { return c[0]; });
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  var maxRows = sheet.getMaxRows();
  var formats = [];
  var rowFormat = cols.map(function (c) { return cellFormat_(c[1]); });
  for (var r = 0; r < maxRows; r++) formats.push(rowFormat);
  sheet.getRange(1, 1, maxRows, headers.length).setNumberFormats(formats);
  return sheet;
}

/* ===================== small utilities ===================== */

function fail_(message, code) {
  var error = new Error(message);
  if (code) error.code = code;
  throw error;
}

function errorMessage_(error) {
  return (error && error.message) ? String(error.message) : String(error);
}

function errorResponse_(error) {
  var body = { ok: false, error: errorMessage_(error) };
  if (error && error.code) body.code = error.code;
  if (error && error.passwordError) body.passwordError = true;
  return jsonResponse_(body);
}

function jsonResponse_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function plain_(obj) {
  var out = {};
  Object.keys(obj).forEach(function (k) { out[k] = obj[k]; });
  return out;
}

function sortBy_(arr, cmp) { // stable sort
  return arr.map(function (v, i) { return [v, i]; })
    .sort(function (a, b) { return cmp(a[0], b[0]) || a[1] - b[1]; })
    .map(function (p) { return p[0]; });
}

function byMovedAsc_(a, b) {
  if (a.moved_at !== b.moved_at) return a.moved_at < b.moved_at ? -1 : 1;
  return a.id - b.id;
}

function byMovedDesc_(a, b) {
  if (a.moved_at !== b.moved_at) return a.moved_at < b.moved_at ? 1 : -1;
  return b.id - a.id;
}

function str_(v) {
  if (v === null || v === undefined || v === false) return "";
  if (v === true) return "1";
  return String(v);
}

// Like str_ but rejects arrays/objects (PHP trim() on an array throws).
function safeStr_(v) {
  if (v !== null && typeof v === "object") fail_("ข้อมูลที่ส่งมาไม่ถูกต้อง");
  return str_(v);
}

function phpTrim_(s) {
  return String(s).replace(/^[ \t\n\r\0\x0B]+|[ \t\n\r\0\x0B]+$/g, "");
}

function phpInt_(v) {
  if (v === null || v === undefined || v === false) return 0;
  if (v === true) return 1;
  if (typeof v === "number") return isFinite(v) ? Math.trunc(v) : 0;
  if (typeof v === "object") return Array.isArray(v) && v.length ? 1 : 0;
  var m = /^[ \t\n\r\v\f]*([+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?)/.exec(String(v));
  if (!m) return 0;
  var n = Number(m[1]);
  return isFinite(n) ? Math.trunc(n) : 0;
}

function phpTruthy_(v) {
  if (v === null || v === undefined || v === false || v === 0 || v === "" || v === "0") return false;
  if (Array.isArray(v)) return v.length > 0;
  return true;
}

function phpEmpty_(v) {
  return !phpTruthy_(v);
}

function mbLen_(s) {
  return Array.from(String(s)).length;
}

function mbSubstr_(s, start, len) {
  return Array.from(String(s)).slice(start, start + len).join("");
}

function eqCi_(a, b) {
  if (a === null || a === undefined || b === null || b === undefined) return false;
  return String(a).toLowerCase() === String(b).toLowerCase();
}

// MySQL "col LIKE '%q%'" (case-insensitive, % and _ wildcards, \ escape).
function likeMatcher_(q) {
  var re = "";
  var chars = Array.from(q);
  for (var i = 0; i < chars.length; i++) {
    var ch = chars[i];
    if (ch === "\\" && i + 1 < chars.length) { re += escapeRe_(chars[++i]); continue; }
    if (ch === "%") re += "[\\s\\S]*";
    else if (ch === "_") re += "[\\s\\S]";
    else re += escapeRe_(ch);
  }
  var rx = new RegExp(re, "i");
  return function (v) { return v !== null && v !== undefined && rx.test(String(v)); };
}

function escapeRe_(ch) {
  return ch.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");
}

// CSV line: quoted when PHP fputcsv() would quote, "\n" line ending.
function csvLine_(fields) {
  return fields.map(function (f) {
    var s = f === null || f === undefined ? "" : (f === true ? "1" : f === false ? "" : String(f));
    if (!/[,"\\\n\r\t ]/.test(s)) return s;
    // Every '"' doubled (RFC 4180, how Excel reads it). PHP's backslash escape left
    // '\"' undoubled, which split the cell into two columns in Excel.
    return '"' + s.replace(/"/g, '""') + '"';
  }).join(",") + "\n";
}

function pad2_(n) {
  return (n < 10 ? "0" : "") + n;
}

function formatUtcParts_(d) {
  var y = String(d.getUTCFullYear());
  while (y.length < 4) y = "0" + y;
  var date = y + "-" + pad2_(d.getUTCMonth() + 1) + "-" + pad2_(d.getUTCDate());
  var time = pad2_(d.getUTCHours()) + ":" + pad2_(d.getUTCMinutes());
  return {
    date: date,
    time: time,
    datetime: date + " " + time + ":" + pad2_(d.getUTCSeconds()),
    ymd: y.slice(-2) + pad2_(d.getUTCMonth() + 1) + pad2_(d.getUTCDate())
  };
}

// Current wall-clock time in Asia/Bangkok.
function nowParts_() {
  return formatUtcParts_(new Date(Date.now() + TZ_OFFSET_MS));
}

function parseDt_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(String(s));
  if (!m) return NaN;
  return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
}

function shiftDate_(date, days) {
  return formatUtcParts_(new Date(parseDt_(date) + days * 86400000)).date;
}

function shiftDateTime_(datetime, days) {
  return formatUtcParts_(new Date(parseDt_(datetime) + days * 86400000)).datetime;
}

function stamp_() {
  var p = nowParts_();
  return p.date.replace(/-/g, "") + "_" + p.datetime.slice(11).replace(/:/g, "");
}
