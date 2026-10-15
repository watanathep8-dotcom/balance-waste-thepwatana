/*
 * In-memory mocks of the Apps Script services used by apps-script/Code.gs (Balance Waste),
 * plus an end-to-end test of every API action through doGet/doPost.
 *
 *   node test/gas-mock.js
 */
"use strict";
const vm = require("vm");
const fs = require("fs");
const path = require("path");
const assert = require("assert");
const crypto = require("crypto");
const { Buffer } = require("buffer"); // explicit: test/dev-server.js evaluates this file in a bare vm context

const ROOT = path.join(__dirname, "..");

/* ===================== mocks ===================== */
function createGas() {
  const state = { props: {}, cache: {}, cachePuts: [], cachePutAlls: [], cacheGetAlls: 0, flushes: 0, cacheDown: false, sheetReads: {}, rowReads: 0, events: [], sleeps: 0, locks: 0, unlocks: 0, lockFail: false, spreadsheets: {}, seq: 0, formatCalls: 0, created: 0, fetches: [], fetchStatus: 202, fetchError: null, logs: [] };

  class Range {
    constructor(sheet, row, col, nr, nc) {
      if (row < 1 || col < 1 || nr < 1 || nc < 1) throw new Error("Range out of bounds");
      if (row + nr - 1 > sheet.maxRows) throw new Error("Range exceeds sheet rows: " + (row + nr - 1) + " > " + sheet.maxRows);
      if (col + nc - 1 > sheet.maxCols) throw new Error("Range exceeds sheet columns: " + (col + nc - 1) + " > " + sheet.maxCols);
      Object.assign(this, { sheet, row, col, nr, nc });
    }
    getValues() {
      if (this.nr === 1 && this.row > 1) state.rowReads++;
      const out = [];
      for (let r = 0; r < this.nr; r++) {
        const src = this.sheet.data[this.row - 1 + r] || [];
        const line = [];
        for (let c = 0; c < this.nc; c++) {
          const v = src[this.col - 1 + c];
          line.push(v === undefined || v === null ? "" : v);
        }
        out.push(line);
      }
      return out;
    }
    setValues(values) {
      state.events.push("write:" + this.sheet.name);
      assert.strictEqual(values.length, this.nr, "setValues row count");
      values.forEach((line, r) => {
        assert.strictEqual(line.length, this.nc, "setValues column count");
        const idx = this.row - 1 + r;
        while (this.sheet.data.length <= idx) this.sheet.data.push([]);
        line.forEach((v, c) => {
          if (typeof v === "string" && v.charAt(0) === "'") v = v.slice(1); // Sheets consumes the apostrophe prefix (stored as text)
          else if (typeof v === "string" && v.charAt(0) === "=") throw new Error("Formula written to sheet: " + v);
          this.sheet.data[idx][this.col - 1 + c] = v;
        });
      });
      return this;
    }
    setNumberFormats(f) {
      assert.strictEqual(f.length, this.nr, "setNumberFormats rows");
      f.forEach(line => assert.strictEqual(line.length, this.nc, "setNumberFormats cols"));
      state.formatCalls++;
      return this;
    }
    setFontWeight() { return this; }
    clearContent() {
      state.clearCalls = (state.clearCalls || 0) + 1;
      for (let r = 0; r < this.nr; r++) {
        const line = this.sheet.data[this.row - 1 + r];
        if (line) for (let c = 0; c < this.nc; c++) if (this.col - 1 + c < line.length) line[this.col - 1 + c] = "";
      }
      return this;
    }
  }

  class Sheet {
    constructor(name) { this.name = name; this.data = []; this.maxRows = 1000; this.maxCols = 26; }
    getName() { return this.name; }
    setName(n) { this.name = n; return this; }
    getLastRow() {
      for (let r = this.data.length - 1; r >= 0; r--) {
        if ((this.data[r] || []).some(v => v !== "" && v !== undefined && v !== null)) return r + 1;
      }
      return 0;
    }
    getLastColumn() { return this.data.reduce((m, row) => Math.max(m, row.length), 0); }
    getMaxRows() { return this.maxRows; }
    insertRowsAfter(after, n) { this.maxRows += n; return this; }
    getMaxColumns() { return this.maxCols; }
    insertColumnsAfter(after, n) { this.maxCols += n; state.colInserts = (state.colInserts || 0) + 1; return this; }
    getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
    getDataRange() {
      state.sheetReads[this.name] = (state.sheetReads[this.name] || 0) + 1;
      return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()));
    }
    setFrozenRows(n) { this.frozen = n; return this; }
    getFrozenRows() { return this.frozen || 0; }
    deleteRows(pos, n) {
      if (pos < 1 || n < 1 || pos + n - 1 > this.maxRows) throw new Error("deleteRows out of bounds");
      if (this.maxRows - n <= (this.frozen || 0)) throw new Error("Sorry, it is not possible to delete all non-frozen rows.");
      state.deleteCalls = (state.deleteCalls || []).concat([[this.name, pos, n]]);
      this.data.splice(pos - 1, n);
      this.maxRows -= n;
      return this;
    }
    deleteRow(pos) { return this.deleteRows(pos, 1); }
  }

  class Spreadsheet {
    constructor(name) { this.id = "ss" + (++state.seq); this.name = name; this.sheets = [new Sheet("Sheet1")]; this.tz = null; }
    getId() { return this.id; }
    getUrl() { return "https://docs.google.com/spreadsheets/d/" + this.id + "/edit"; }
    getSheets() { return this.sheets.slice(); }
    getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
    insertSheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; }
    setSpreadsheetTimeZone(tz) { this.tz = tz; }
  }

  const gas = {
    SpreadsheetApp: {
      create(name) { const ss = new Spreadsheet(name); state.spreadsheets[ss.id] = ss; state.created++; return ss; },
      openById(id) { const ss = state.spreadsheets[id]; if (!ss) throw new Error("Spreadsheet not found: " + id); state.opens = (state.opens || 0) + 1; return ss; },
      flush() { state.flushes++; state.events.push("flush"); }
    },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty: k => (k in state.props ? state.props[k] : null),
          getProperties: () => Object.assign({}, state.props),
          setProperty: (k, v) => { state.props[k] = String(v); },
          setProperties: (o) => { Object.assign(state.props, o); },
          deleteProperty: k => { delete state.props[k]; }
        };
      }
    },
    CacheService: {
      getScriptCache() {
        return {
          get: k => { if (state.cacheDown) throw new Error("Cache down"); return k in state.cache ? state.cache[k] : null; },
          put: (k, v, ttl) => { state.cache[k] = String(v); state.cachePuts.push(ttl); },
          remove: k => { delete state.cache[k]; },
          // Batch calls (read cache / data versions); tracked apart from put().
          getAll: keys => { if (state.cacheDown || state.batchCacheDown) throw new Error("Cache down"); state.cacheGetAlls++; const o = {}; keys.forEach(k => { if (k in state.cache) o[k] = state.cache[k]; }); return o; },
          putAll: (map, ttl) => {
            if (state.cacheDown || state.batchCacheDown) throw new Error("Cache down");
            Object.keys(map).forEach(k => {
              if (/^WASTE_V_/.test(k)) state.events.push("version:" + k.slice(8));
              const v = String(map[k]);
              assert.ok(k.length <= 250, "cache key too long: " + k);
              assert.ok(Buffer.byteLength(v, "utf8") <= 100 * 1024, "cache value over 100 KB: " + k);
              state.cache[k] = v;
            });
            state.cachePutAlls.push(ttl);
          },
          removeAll: keys => { keys.forEach(k => { delete state.cache[k]; }); }
        };
      }
    },
    LockService: {
      getScriptLock() {
        return {
          waitLock() { if (state.lockFail) throw new Error("Lock timeout"); state.locks++; },
          tryLock() { state.locks++; return true; },
          releaseLock() { state.unlocks++; }
        };
      }
    },
    Utilities: {
      sleep(ms) { state.sleeps += ms; },
      DigestAlgorithm: { SHA_256: "SHA_256" },
      Charset: { UTF_8: "UTF_8" },
      getUuid() { return crypto.randomUUID(); },
      // Like Apps Script: string (charset) or byte array in, signed byte array out.
      computeDigest(alg, value) {
        assert.strictEqual(alg, "SHA_256");
        const buf = typeof value === "string" ? Buffer.from(value, "utf8") : Buffer.from(value.map(b => b & 255));
        return Array.from(crypto.createHash("sha256").update(buf).digest(), b => (b > 127 ? b - 256 : b));
      }
    },
    ContentService: {
      MimeType: { JSON: "application/json", JAVASCRIPT: "application/javascript", CSV: "text/csv", TEXT: "text/plain" },
      createTextOutput(text) {
        return { text, mime: null, setMimeType(m) { this.mime = m; return this; }, getContent() { return this.text; }, getMimeType() { return this.mime; } };
      }
    },
    Logger: { log() {} },
    // Everything Code.gs logs is kept in state.logs (checked for leaked secrets).
    console: {
      log(...a) { state.logs.push(["log", a.join(" ")]); },
      warn(...a) { state.logs.push(["warn", a.join(" ")]); },
      error(...a) { state.logs.push(["error", a.join(" ")]); console.error(...a); }
    },
    // Teams webhook: never touches the network. Records each call (and whether the
    // script lock was held at that moment); answers state.fetchStatus or throws
    // state.fetchError.
    UrlFetchApp: {
      fetch(url, opts) {
        state.fetches.push({ url, opts, lockHeld: state.locks !== state.unlocks });
        if (state.fetchError) throw new Error(state.fetchError);
        const code = state.fetchStatus;
        return { getResponseCode: () => code, getContentText: () => "" };
      }
    },
    // Project triggers (keep-warm). state.triggers: [{handler, everyMinutes, id}]
    ScriptApp: {
      getProjectTriggers() {
        return (state.triggers || []).map(t => ({ getHandlerFunction: () => t.handler, getUniqueId: () => t.id, _t: t }));
      },
      deleteTrigger(tr) { state.triggers = (state.triggers || []).filter(t => t !== tr._t); },
      newTrigger(handler) {
        const t = { handler, id: "trg" + (++state.seq) };
        const b = {
          timeBased() { t.timeBased = true; return b; },
          everyMinutes(n) { assert.ok([1, 5, 10, 15, 30].includes(n), "everyMinutes(" + n + ")"); t.everyMinutes = n; return b; },
          create() { assert.ok(t.timeBased, "time-based"); (state.triggers = state.triggers || []).push(t); return { getUniqueId: () => t.id }; }
        };
        return b;
      }
    }
  };
  return { gas, state };
}

function loadCode(gas) {
  const ctx = vm.createContext(Object.assign({}, gas));
  const src = fs.readFileSync(path.join(ROOT, "apps-script", "Code.gs"), "utf8");
  vm.runInContext(src, ctx, { filename: "Code.gs" });
  return ctx;
}


/* ===================== synthetic test fixtures =====================
 * The real install starts EMPTY (no seed). Tests that need waste types create these
 * 13 synthetic types (fake receivers / registrations) through the admin API
 * (type_save), exactly like the settings page does. Not real company data. */
const RCV_A = "บจก. ผู้รับ ทดสอบ เอ", RCV_B = "บมจ. ผู้รับ ทดสอบ บี", RCV_C = "บมจ. ผู้รับ ทดสอบ ซี";
// [code, name, permit tons, management code, receiver registration, receiver, icon, colour]
const FIXTURE_TYPES = [
  ["150101", "กระดาษลูกฟูก กระดาษลัง", 20, "011", "1000000000001", RCV_A, "fa-box-open", "#B7791F"],
  ["150102", "ถุงพลาสติก ถังพลาสติก", 30, "011", "1000000000001", RCV_A, "fa-bottle-water", "#0EA5E9"],
  ["170405", "เศษเหล็ก", 10, "011", "1000000000001", RCV_A, "fa-gears", "#64748B"],
  ["150111", "กระป๋องสเปรย์", 2, "073", "1000000000002", RCV_B, "fa-spray-can", "#EF4444"],
  ["080409", "กากกาว", 10, "042", "1000000000003", RCV_B, "fa-fill-drip", "#A855F7"],
  ["150202", "วัสดุปนเปื้อน", 30, "042", "1000000000003", RCV_B, "fa-biohazard", "#F97316"],
  ["160215", "หลอดไฟ", 0.5, "073", "1000000000002", RCV_B, "fa-lightbulb", "#EAB308"],
  ["160213", "อุปกรณ์อิเล็กทรอนิกส์", 1, "073", "1000000000002", RCV_B, "fa-microchip", "#6366F1"],
  ["150102", "บรรจุภัณฑ์พลาสติก", 20, "046", "1000000000003", RCV_B, "fa-jug-detergent", "#06B6D4"],
  ["191201", "เศษสติ๊กเกอร์", 20, "046", "1000000000003", RCV_B, "fa-note-sticky", "#EC4899"],
  ["190813", "ตะกอนจากระบบบำบัดน้ำเสีย", 10, "073", "1000000000004", RCV_C, "fa-water", "#8B5E3C"],
  ["160215", "หลอดไฟ", 0.5, "073", "1000000000004", RCV_C, "fa-lightbulb", "#F59E0B"],
  ["161001", "น้ำล้างเครื่องจักร (น้ำแรก)", 40, "065", "1000000000005", RCV_C, "fa-faucet-drip", "#14B8A6"]
];
// call(body): posts one request as admin with an actor (body gets action + fields)
function seedFixtureTypes(call) {
  FIXTURE_TYPES.forEach(t => {
    const r = call({ action: "type_save", code: t[0], name: t[1], permit_tons: t[2], mgmt_code: t[3], receiver_reg: t[4], receiver: t[5], icon: t[6], color: t[7], active: 1 });
    assert.strictEqual(r.ok, true, "fixture type_save: " + JSON.stringify(r));
  });
}

/* ===================== test runner ===================== */
let passed = 0, failed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log("  ok   " + name); }
  catch (e) { failed++; console.log("  FAIL " + name + "\n       " + String(e && e.stack || e).split("\n").slice(0, 4).join("\n       ")); }
}

const { gas, state } = createGas();
const ctx = loadCode(gas);
// Fake passwords of the in-memory mock (tests only; the real ones live in Script Properties).
const VPW = "view-pass", SPW = "staff-pass", RPW = "admin-pass";
const PROPS = { WASTE_VIEW_PASSWORD: VPW, WASTE_STAFF_PASSWORD: SPW, WASTE_RESET_PASSWORD: RPW };
const ACTOR = "Tester";
// public (no password), the in-house write (staff+), admin tools
const READS = ["bootstrap", "dashboard", "history", "export", "batch"];
const PUBLIC_WRITES = ["store", "ship", "opening"];
const STAFF_ONLY = ["treat"];
const ADMIN_ONLY = ["movement_update", "movement_delete", "type_save", "receiver_save", "receiver_delete", "type_delete", "reset_data", "logs", "logs_export"];
const PUBLIC_ACTIONS = READS.concat(PUBLIC_WRITES);
const ALL_ACTIONS = PUBLIC_ACTIONS.concat(STAFF_ONLY, ADMIN_ONLY);
const getRaw = (params = {}) => {
  const out = ctx.doGet({ parameter: params });
  assert.strictEqual(out.getMimeType(), "application/json");
  return JSON.parse(out.getContent());
};
const rawPost = payload => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(payload) } }).getContent());
const postRaw = (payload, who = ACTOR) => rawPost(Object.assign(who ? { actor: who } : {}, payload));
const loginRes = (password, client) => rawPost({ action: "login", password, client });
const login = (password, client) => okData(loginRes(password, client)).token;
// session tokens of the three levels (set after login)
let TV = "", TS = "", TA = "";
// the lowest token each action needs: none for the public ones
const tokFor = a => PUBLIC_ACTIONS.includes(a) ? undefined : STAFF_ONLY.includes(a) ? TS : TA;
const post = (action, body = {}, who = ACTOR) => postRaw(Object.assign({ token: tokFor(action) }, action === "reset_data" ? { resetPassword: RPW } : {}, body, { action }), who);
// a public read (anonymous POST, no token)
const get = (action, params = {}) => postRaw(Object.assign({ action }, params), null);
// the level a token gets on a public read (null: anonymous / unknown / expired)
const roleOf = tok => okData(postRaw({ action: "bootstrap", token: tok }, null)).role;
const okData = res => { assert.strictEqual(res.ok, true, "expected ok, got: " + JSON.stringify(res)); return res.data; };
const err = (res, re) => {
  assert.strictEqual(res.ok, false, "expected error, got: " + JSON.stringify(res).slice(0, 300));
  if (re) assert.ok(re.test(res.error), "error '" + res.error + "' !~ " + re);
  return res;
};
const clearFails = () => Object.keys(state.cache).forEach(k => { if (/FAILURES|^WASTE_LF_/.test(k)) delete state.cache[k]; });
// kg -> exact tons text (what the page sends as weight_t), e.g. 1234.5 -> "1.2345"
const T = kg => {
  const c = Math.round(Number(String(kg).replace(/,/g, "")) * 100);
  return String(Math.floor(c / 100000)) + "." + String(c % 100000).padStart(5, "0");
};
const p2 = n => String(n).padStart(2, "0");
const bkkNow = () => new Date(Date.now() + 7 * 3600 * 1000);
const fmtD = d => `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}`;
const TODAY = fmtD(bkkNow());
const YMD = TODAY.slice(2).replace(/-/g, "");
// a moment `back` days ago at hh:mm (Asia/Bangkok)
const at = (back, hh = 9, mm = 0) => ({ date: fmtD(new Date(bkkNow().getTime() - back * 86400000)), time: p2(hh) + ":" + p2(mm) });
const sheet = n => state.spreadsheets[state.props.WASTE_SPREADSHEET_ID].getSheetByName(n);
const dataRows = n => sheet(n).getLastRow() - 1;
const boot = () => okData(get("bootstrap"));
const bal = id => boot().balance[id];
const logs = (p = {}) => okData(post("logs", p)).items;
const lastLog = () => logs()[0];
// a movement; the weight is given in kg here and sent as weight_t (tons) like the page does
const mvT = (kind, type_id, weight_t, extra = {}) => post(kind, Object.assign({ type_id, weight_t, container: "drum200", containers: 1 }, at(5), extra));
const mv = (kind, type_id, kg, extra = {}) => mvT(kind, type_id, T(kg), extra);

console.log("Balance Waste GAS backend tests (today " + TODAY + " Asia/Bangkok)");

/* ---------- login / setup ---------- */
test("login: no password property set -> nobody can log in (clear Thai message naming the 3 properties)", () => {
  const r = err(loginRes("anything"), /ยังไม่ได้ตั้งรหัสผ่านใน Script Properties/);
  assert.ok(/WASTE_RESET_PASSWORD, WASTE_STAFF_PASSWORD, WASTE_VIEW_PASSWORD/.test(r.error), r.error);
  assert.strictEqual(r.passwordError, true);
  err(get("bootstrap"), /setupSystem/); // public read: no password needed (only the setup is missing)
});

test("API before setup: reads and writes report the setupSystem error (after a valid login)", () => {
  Object.assign(state.props, PROPS);
  TV = login(VPW); TS = login(SPW); TA = login(RPW);
  err(get("bootstrap"), /setupSystem/);
  err(mv("store", 1, 10), /setupSystem/);
  delete state.props.WASTE_VIEW_PASSWORD; delete state.props.WASTE_STAFF_PASSWORD; delete state.props.WASTE_RESET_PASSWORD;
});

test("setupSystem creates an EMPTY spreadsheet: 4 sheets with headers, no waste types, no receivers", () => {
  const r = ctx.setupSystem();
  assert.strictEqual(r.created, true);
  assert.strictEqual(r.seededTypes, 0);
  assert.strictEqual(r.resetPasswordConfigured, false);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r.passwords)), { WASTE_RESET_PASSWORD: false, WASTE_STAFF_PASSWORD: false, WASTE_VIEW_PASSWORD: false });
  assert.ok(/ยังไม่ได้ตั้ง Script Property: WASTE_RESET_PASSWORD \(ผู้ดูแลระบบ\), WASTE_STAFF_PASSWORD \(พนักงานกรอกข้อมูล\), WASTE_VIEW_PASSWORD/.test(r.passwordMessage), r.passwordMessage);
  const ss = state.spreadsheets[state.props.WASTE_SPREADSHEET_ID];
  assert.strictEqual(ss.tz, "Asia/Bangkok");
  assert.strictEqual(ss.name, "Balance Waste Database");
  assert.deepStrictEqual(ss.getSheets().map(s => s.getName()), ["waste_types", "receivers", "movements", "audit_logs"]);
  assert.deepStrictEqual(sheet("receivers").data[0], ["id", "name", "reg", "mgmt_code", "active", "sort", "updated_at", "updated_by"]);
  assert.deepStrictEqual(sheet("waste_types").data[0], ["id", "code", "name", "permit_tons", "mgmt_code", "receiver_reg", "receiver", "icon", "color", "active", "sort", "updated_at", "updated_by"]);
  assert.strictEqual(sheet("movements").data[0][0], "id");
  assert.strictEqual(dataRows("receivers"), 0);
  assert.strictEqual(dataRows("waste_types"), 0);
  assert.strictEqual(dataRows("movements"), 0);
  assert.strictEqual(dataRows("audit_logs"), 0);
  Object.assign(state.props, PROPS);
  TV = login(VPW); TS = login(SPW); TA = login(RPW);
});

test("company identity: Thepwatana name; factory registration / permit number / year are optional and empty; no original-company data in Code.gs", () => {
  const b = boot();
  assert.strictEqual(b.company.name, "บริษัท เทพวัฒนา จำกัด");
  assert.strictEqual(b.company.factoryReg, "");
  assert.strictEqual(b.company.permitNo, "");
  assert.strictEqual(b.company.permitYear, "");
  const code = fs.readFileSync(path.join(ROOT, "apps-script", "Code.gs"), "utf8");
  for (const bad of ["แน็ป", "82260007625693", "2569-8409", "SEED_TYPES", "Nutri", "NAP"]) assert.ok(!code.includes(bad), "Code.gs must not contain " + bad);
});

test("EMPTY install: every read works with zero types / receivers / movements (zeros, empty lists, no error)", () => {
  const b = boot();
  assert.deepStrictEqual([b.types, b.receivers, b.balance], [[], [], {}]);
  assert.strictEqual(b.role, null);
  const d = okData(get("dashboard"));
  assert.strictEqual(d.kpi.balance_kg, 0);
  assert.strictEqual(d.kpi.permit_kg, 0);
  assert.strictEqual(d.kpi.year_out_kg, 0);
  assert.ok(d.kpi.quota_pct == null || d.kpi.quota_pct === 0);
  assert.deepStrictEqual(d.receivers, []);
  assert.deepStrictEqual(d.recent, []);
  assert.deepStrictEqual(d.alerts, []);
  assert.strictEqual(d.months.length, 12);
  assert.deepStrictEqual(okData(get("history")).items, []);
  assert.ok(/^waste_thepwatana_history_\d{8}_\d{6}\.csv$/.test(okData(get("export")).filename));
  assert.deepStrictEqual(okData(post("logs")).items, []);
  assert.ok(okData(get("batch", { reads: ["bootstrap", "dashboard", "history"] })));
  // before any type exists a movement is refused with a clear message and writes nothing
  err(mv("store", 1, 10), /ประเภทของเสีย/);
  assert.strictEqual(dataRows("movements"), 0);
});

test("first run: admin adds the first receiver + first type (settings page), then the first movements work", () => {
  okData(post("receiver_save", { name: "บจก. ผู้รับรายแรก", reg: "1234567890123" }));
  assert.strictEqual(okData(get("bootstrap")).receivers.length, 1);
  assert.strictEqual(okData(get("bootstrap")).types.length, 0);
  okData(post("type_save", { code: "150101", name: "กระดาษทดสอบ", permit_tons: "5", mgmt_code: "011", receiver: "บจก. ผู้รับรายแรก", receiver_reg: "1234567890123", icon: "fa-box-open", color: "#B7791F", active: 1 }));
  const b0 = boot();
  assert.strictEqual(b0.types.length, 1);
  assert.strictEqual(b0.receivers.length, 1, "no duplicate receiver");
  const id = b0.types[0].id;
  assert.strictEqual(b0.balance[id].balance_kg, 0);
  assert.strictEqual(okData(get("dashboard")).kpi.permit_kg, 5000);
  assert.strictEqual(okData(mv("store", id, 1500)).balance_kg, 1500);
  assert.strictEqual(okData(mv("ship", id, 500)).balance_kg, 1000);
  const d = okData(get("dashboard"));
  assert.strictEqual(d.kpi.balance_kg, 1000);
  assert.strictEqual(d.kpi.year_out_kg, 500);
  assert.strictEqual(d.recent.length, 2);
  assert.strictEqual(dataRows("movements"), 2);
});

test("back to empty (test setup): reset, then the first-run type and receiver are deleted", () => {
  okData(post("reset_data"));
  const b = boot();
  okData(post("type_delete", { id: b.types[0].id }));
  okData(post("receiver_delete", { id: b.receivers[0].id }));
  assert.deepStrictEqual([boot().types.length, boot().receivers.length], [0, 0]);
  assert.strictEqual(dataRows("waste_types"), 0);
});

test("fixture: 13 synthetic types created through type_save match the permit table (codes keep leading zeros, total 194 t)", () => {
  seedFixtureTypes(b => postRaw(Object.assign({ token: TA }, b)));
  const b = boot();
  const t = b.types;
  assert.strictEqual(t.length, 13);
  assert.deepStrictEqual(t.map(x => x.code), ["150101", "150102", "170405", "150111", "080409", "150202", "160215", "160213", "150102", "191201", "190813", "160215", "161001"]);
  assert.deepStrictEqual(t.map(x => x.permit_tons), [20, 30, 10, 2, 10, 30, 0.5, 1, 20, 20, 10, 0.5, 40]);
  assert.strictEqual(t[4].code, "080409");
  assert.strictEqual(Math.round(t.reduce((a, x) => a + x.permit_tons, 0) * 1000) / 1000, 194);
  assert.strictEqual(b.receivers.length, 3);
  assert.ok(t.every(x => x.active === 1 && /^fa-/.test(x.icon) && /^#[0-9A-F]{6}$/.test(x.color)));
  assert.notStrictEqual(t[6].id, t[11].id);
  assert.deepStrictEqual(b.containers, { drum200: "ถัง 200 ลิตร", ibc: "IBC", bigbag: "Bigbag" });
  Object.values(b.balance).forEach(x => { assert.strictEqual(x.balance_kg, 0); assert.strictEqual(x.age_days, null); });
  assert.strictEqual(dataRows("movements"), 0);
});

test("setupSystem is idempotent (no second spreadsheet, never seeds, keeps existing types)", () => {
  const id = state.props.WASTE_SPREADSHEET_ID;
  const r = ctx.setupSystem();
  assert.strictEqual(r.created, false);
  assert.strictEqual(r.seededTypes, 0);
  assert.strictEqual(state.props.WASTE_SPREADSHEET_ID, id);
  assert.strictEqual(state.created, 1);
  assert.strictEqual(dataRows("waste_types"), 13);
  assert.strictEqual(r.resetPasswordConfigured, true);
  assert.ok(Object.values(r.passwords).every(Boolean));
  assert.strictEqual(r.passwordMessage, "ตั้งรหัสผ่านครบทั้ง 3 ระดับแล้ว");
  delete state.props.WASTE_VIEW_PASSWORD;
  assert.ok(/^ยังไม่ได้ตั้ง Script Property: WASTE_VIEW_PASSWORD \(ดูได้อย่างเดียว\) — /.test(ctx.setupSystem().passwordMessage));
  state.props.WASTE_VIEW_PASSWORD = VPW;
});

/* ---------- transport ---------- */
test("GET: public reads work anonymously (a token in the URL is ignored); every other action is POST only", () => {
  assert.strictEqual(okData(getRaw({})).service, "Balance Waste API");
  err(getRaw({ action: "nope" }), /Unknown action/);
  for (const a of READS) {
    const d = okData(getRaw({ action: a, token: TA, reads: JSON.stringify(["bootstrap"]) }));
    if (a === "bootstrap") assert.ok(d.role === null && d.treat === false, "a token in a URL is never used");
  }
  for (const a of PUBLIC_WRITES.concat(STAFF_ONLY, ADMIN_ONLY, ["login", "logout"])) err(getRaw({ action: a, token: TA }), /POST/);
  for (const a of ["treat_unlock", "treat_lock", "verifyResetPassword"]) err(getRaw({ action: a }), /Unknown action/); // removed
  const bad = JSON.parse(ctx.doPost({ postData: { contents: "{oops" } }).getContent());
  err(bad, /JSON/);
  err(postRaw({ action: "nope", token: TA }), /Unknown action/);
  for (const a of ["treat_unlock", "treat_lock", "verifyResetPassword"]) err(postRaw({ action: a, token: TA, resetPassword: RPW }), /Unknown action/);
});

test("actor name required for every write (empty / whitespace / object rejected, nothing written)", () => {
  const logsBefore = dataRows("audit_logs");
  for (const who of [null, "", "   ", { x: 1 }]) {
    for (const a of ["store", "ship", "opening", "treat", "movement_update", "movement_delete", "type_save", "reset_data"]) {
      const r = postRaw(Object.assign({ action: a, token: TA, type_id: 1, weight_t: "0.005", container: "drum200", containers: 1, resetPassword: RPW }, who && typeof who === "object" ? { actor: who } : {}), typeof who === "string" ? who : null);
      err(r, /ระบุชื่อผู้ใช้งาน/);
      assert.strictEqual(r.code, "NAME_REQUIRED");
    }
  }
  assert.strictEqual(dataRows("movements"), 0);
  assert.strictEqual(dataRows("audit_logs"), logsBefore);
});

/* ---------- store validations ---------- */
test("store: field validation (type, weight, decimals, container, count, date, text length)", () => {
  err(mv("store", 0, 10), /เลือกประเภทของเสีย/);
  err(mv("store", 99, 10), /เลือกประเภทของเสีย/);
  err(mvT("store", 1, 0), /มากกว่า 0 ตัน/);
  err(mvT("store", 1, "0.000"), /มากกว่า 0 ตัน/);
  err(mvT("store", 1, ""), /กรอกน้ำหนัก \(ตัน\)/);
  err(mvT("store", 1, null), /กรอกน้ำหนัก/);
  err(mvT("store", 1, undefined), /กรอกน้ำหนัก/);
  err(mvT("store", 1, -5), /ตัวเลข/);
  err(mvT("store", 1, "abc"), /ตัวเลข/);
  err(mvT("store", 1, "1e3"), /ตัวเลข/);
  err(mvT("store", 1, "0.000001"), /ทศนิยมได้ไม่เกิน 5/); // finer than 0.01 kg
  err(mvT("store", 1, 0.0000015), /ทศนิยมได้ไม่เกิน 5|ตัวเลข/);
  err(mvT("store", 1, [5]), /กรอกน้ำหนัก/);
  err(mvT("store", 1, "1000.001"), /ไม่เกิน 1,000.00 ตัน/);
  // the old kg field is not a weight any more (the page sends tons)
  err(post("store", Object.assign({ type_id: 1, weight_kg: 5, container: "drum200", containers: 1 }, at(5))), /กรอกน้ำหนัก/);
  err(mv("store", 1, 10, { container: "box" }), /ภาชนะ/);
  err(mv("store", 1, 10, { container: "" }), /ภาชนะ/);
  err(mv("store", 1, 10, { containers: 0 }), /จำนวนเต็มตั้งแต่ 1/);
  err(mv("store", 1, 10, { containers: "1.5" }), /จำนวนเต็มตั้งแต่ 1/);
  err(mv("store", 1, 10, { containers: -2 }), /จำนวนเต็มตั้งแต่ 1/);
  err(mv("store", 1, 10, { date: "2026-02-30" }), /วันที่\/เวลาไม่ถูกต้อง/);
  err(mv("store", 1, 10, { date: "15/10/2026" }), /วันที่\/เวลาไม่ถูกต้อง/);
  err(mv("store", 1, 10, { time: "25:00" }), /วันที่\/เวลาไม่ถูกต้อง/);
  err(mv("store", 1, 10, at(-2)), /ไม่เกินเวลาปัจจุบัน/);
  err(mv("store", 1, 10, { note: "x".repeat(256) }), /หมายเหตุยาวเกิน 255/);
  err(mv("store", 1, 10, { location: "x".repeat(101) }), /สถานที่จัดเก็บยาวเกิน 100/);
  err(mv("store", 1, 10, { type_id: { a: 1 } }), /เลือกประเภท/);
  assert.strictEqual(dataRows("movements"), 0);
  // rejected writes are audited as warnings
  const w = logs({ cat: "warn" });
  assert.ok(w.length >= 10 && w.every(l => /^ปฏิเสธ/.test(l.detail) && l.actor === ACTOR));
});

test("store: saved with doc number, actor, weight in tons stored as kg; comma input accepted; audit row", () => {
  const r = okData(mvT("store", 1, "1.2345", { container: "bigbag", containers: 3, location: "โรงเก็บ A", note: "ทดสอบ" }));
  assert.strictEqual(r.doc_no, "ST-" + YMD + "-0001");
  assert.strictEqual(r.kind, "store");
  assert.strictEqual(r.balance_kg, 1234.5);
  assert.ok(/บันทึกจัดเก็บ .* 1\.2345 ตัน เรียบร้อย/.test(r.message), r.message);
  const row = sheet("movements").data[1];
  assert.strictEqual(row[1], r.doc_no);
  assert.strictEqual(row[2], "store");
  assert.strictEqual(row[4], 1234.5);
  assert.strictEqual(row[5], "bigbag");
  assert.strictEqual(row[6], 3);
  assert.strictEqual(row[7], "โรงเก็บ A");
  assert.strictEqual(row[13], ACTOR);
  const l = lastLog();
  assert.strictEqual(l.category, "waste");
  assert.strictEqual(l.ref, r.doc_no);
  assert.ok(l.detail.includes("1.2345 ตัน") && l.detail.includes("Bigbag 3 ภาชนะ"), l.detail);
  const r2 = okData(mvT("store", 1, "0.0001", { container: "bigbag", containers: 1 })); // 0.1 kg
  assert.strictEqual(r2.doc_no, "ST-" + YMD + "-0002");
});

test("balance math is exact to 0.01 kg (no float drift) and counts containers per type", () => {
  // type 2: 0.1 + 0.2 ... x10 = 3.00 exactly
  for (let i = 0; i < 10; i++) okData(mv("store", 2, i % 2 ? "0.2" : "0.1", { container: "ibc", containers: 1 }));
  const b = bal(2);
  assert.strictEqual(b.balance_kg, 1.5);
  assert.strictEqual(b.stored_kg, 1.5);
  assert.deepStrictEqual(b.containers, { drum200: 0, ibc: 10, bigbag: 0 });
  okData(mv("ship", 2, "0.3", { container: "ibc", containers: 2, ...at(4) }));
  assert.strictEqual(bal(2).balance_kg, 1.2);
  assert.strictEqual(bal(2).containers.ibc, 8);
  assert.strictEqual(bal(1).balance_kg, 1234.6);
  assert.strictEqual(bal(1).containers.bigbag, 4);
});

/* ---------- units: tons on the page, kg in the sheet ---------- */
function freshEnv() {
  const g = createGas();
  const c = loadCode(g.gas);
  c.setupSystem();
  Object.assign(g.state.props, PROPS);
  const raw = body => JSON.parse(c.doPost({ postData: { contents: JSON.stringify(Object.assign({ actor: ACTOR }, body)) } }).getContent());
  const tok = okData(raw({ action: "login", password: RPW })).token;
  seedFixtureTypes(b => raw(Object.assign({ token: tok }, b)));
  const ss = g.state.spreadsheets[g.state.props.WASTE_SPREADSHEET_ID];
  return { c, ss, call: b => raw(Object.assign({ token: tok }, b)) };
}

test("units: weight_t (tons) -> kg in the sheet exactly (0.001 t = 1 kg, finest 0.01 kg); no float drift", () => {
  const E = freshEnv();
  const ms = E.ss.getSheetByName("movements");
  const cases = [["0.125", 125], ["1", 1000], ["0.001", 1], ["12.345", 12345], [0.3, 300], ["0.1234", 123.4], ["0.00001", 0.01],
    [2.675, 2675], ["1,000", 1000000], ["0.07", 70], [1.005, 1005], ["999.99999", 999999.99], [" 0.5 ", 500]];
  for (const [t, kg] of cases) {
    const r = okData(E.call({ action: "store", type_id: 1, weight_t: t, container: "bigbag", containers: 1, ...at(3) }));
    const row = ms.data.find(x => x[1] === r.doc_no);
    assert.strictEqual(row[4], kg, "weight_t " + JSON.stringify(t));
    assert.strictEqual(typeof row[4], "number");
  }
  // 0.1 t + 0.2 t ... never drifts (sums in hundredths of a kg)
  for (let i = 0; i < 10; i++) okData(E.call({ action: "store", type_id: 2, weight_t: i % 2 ? "0.2" : "0.1", container: "ibc", containers: 1, ...at(3) }));
  const b = okData(E.call({ action: "bootstrap" })).balance[2];
  assert.strictEqual(b.balance_kg, 1500);
  // the text forms: thousands separators, 3 decimals, more only for a fraction of a kg
  assert.strictEqual(E.c.fmtT_(1234567.89), "1,234.56789");
  assert.strictEqual(E.c.fmtT_(125), "0.125");
  assert.strictEqual(E.c.fmtT_(0), "0.00");
  assert.strictEqual(E.c.fmtT_(1000000), "1,000.00");
  assert.strictEqual(E.c.fmtT_(-5), "-0.005");
  assert.strictEqual(E.c.fmtT_(0.01), "0.00001");
  assert.strictEqual(E.c.tonsText_(1234567.8, false), "1234.5678");
  assert.strictEqual(E.c.tonsToKg_("0.125"), 125);
  for (const k of [0.01, 0.1, 1.23, 99.99, 1234.56, 50000.05]) {
    assert.strictEqual(E.c.tonsToKg_(E.c.tonsText_(k, false)), k, "round trip " + k); // every stored kg can be typed back exactly
  }
});

test("units: older kg rows (2 decimals) — shown exactly in tons, ship-all exact, edit without weight keeps the kg", () => {
  const E = freshEnv();
  const ms = E.ss.getSheetByName("movements");
  const st = okData(E.call({ action: "store", type_id: 13, weight_t: "1.23456", container: "ibc", containers: 2, ...at(9) }));
  // an older row typed in kg (as the previous version stored it)
  const row = ms.data.find(x => x[1] === st.doc_no);
  assert.strictEqual(row[4], 1234.56);
  assert.ok(/1\.23456 ตัน เรียบร้อย/.test(st.message), st.message);
  // edit the note only: the weight stays 1234.56 kg (the page sends weight_t only when it was changed)
  okData(E.call({ action: "movement_update", id: st.id, doc_no: st.doc_no, note: "แก้หมายเหตุ" }));
  assert.strictEqual(ms.data.find(x => x[1] === st.doc_no)[4], 1234.56);
  // more than the balance by 0.01 kg is refused with tons in the message
  const r = err(E.call({ action: "ship", type_id: 13, weight_t: "1.23457", container: "ibc", containers: 1, ...at(2) }), /มี 1\.23456 ตัน แต่ต้องนำออก 1\.23457 ตัน/);
  assert.strictEqual(r.code, "NEGATIVE_BALANCE");
  okData(E.call({ action: "ship", type_id: 13, weight_t: "1.23456", container: "ibc", containers: 2, ...at(2) }));
  assert.strictEqual(okData(E.call({ action: "bootstrap" })).balance[13].balance_kg, 0);
  const csv = okData(E.call({ action: "export", type: "13" })).csv;
  assert.ok(csv.split("\n")[0].includes('"น้ำหนัก (ตัน)"'));
  assert.ok(csv.includes(",1.23456,"), csv);
  // quota: permit in tons, shipped in kg -> % unchanged (type 7: 0.5 t)
  okData(E.call({ action: "store", type_id: 7, weight_t: "0.7", container: "drum200", containers: 1, ...at(3) }));
  okData(E.call({ action: "ship", type_id: 7, weight_t: "0.399", container: "drum200", containers: 1, ...at(2) }));
  let b = okData(E.call({ action: "bootstrap" })).balance[7];
  assert.strictEqual(b.permit_kg, 500);
  assert.strictEqual(b.quota_pct, 79.8);
  assert.strictEqual(b.quota_level, "ok");
  // the edit dialog sends tons too
  const sh = okData(E.call({ action: "history", type: "7", kind: "ship" })).items[0];
  okData(E.call({ action: "movement_update", id: sh.id, doc_no: sh.doc_no, weight_t: "0.4" }));
  b = okData(E.call({ action: "bootstrap" })).balance[7];
  assert.strictEqual(b.shipped_year_kg, 400);
  assert.strictEqual(b.quota_pct, 80);
  assert.strictEqual(b.quota_level, "warn");
  const d = okData(E.call({ action: "dashboard" }));
  assert.strictEqual(d.kpi.permit_kg, 194000);
});

/* ---------- ship ---------- */
test("ship: cannot exceed the current balance or containers; exact remaining amount allowed", () => {
  okData(mv("store", 3, 100, { container: "bigbag", containers: 2, ...at(10) }));
  const before = dataRows("movements");
  const r = err(mv("ship", 3, "100.01", { container: "bigbag", containers: 1, ...at(2) }), /ส่งกำจัดเกินยอดคงเหลือ/);
  assert.strictEqual(r.code, "NEGATIVE_BALANCE");
  assert.ok(/มี 0.10 ตัน แต่ต้องนำออก 0\.10001 ตัน/.test(r.error), r.error);
  err(mv("ship", 3, 50, { container: "bigbag", containers: 3, ...at(2) }), /Bigbag.*มี 2 ภาชนะ แต่ต้องนำออก 3/);
  err(mv("ship", 3, 50, { container: "drum200", containers: 1, ...at(2) }), /ถัง 200 ลิตร/);
  err(mv("ship", 4, 1, { ...at(2) }), /มี 0.00 ตัน/);
  assert.strictEqual(dataRows("movements"), before);
  const s = okData(mv("ship", 3, 100, { container: "bigbag", containers: 2, manifest_no: "MN-1", vehicle: "70-1234", ...at(2) }));
  assert.strictEqual(s.doc_no, "SH-" + YMD + "-0002"); // SH-0001 used by the balance-math test
  assert.strictEqual(s.balance_kg, 0);
  const b = bal(3);
  assert.strictEqual(b.balance_kg, 0);
  assert.strictEqual(b.containers.bigbag, 0);
  assert.strictEqual(b.age_days, null);
  const row = sheet("movements").data[dataRows("movements")];
  assert.strictEqual(row[8], "MN-1");
  assert.strictEqual(row[9], "70-1234");
  assert.ok(/ผู้รับ: บจก\. ผู้รับ ทดสอบ เอ/.test(lastLog().detail));
  err(mv("ship", 3, "0.01", { ...at(1) }), /ส่งกำจัดเกินยอดคงเหลือ/);
});

test("ship: a back-dated shipment before the stock existed is refused (chronological replay)", () => {
  okData(mv("store", 5, 50, { ...at(3) }));
  err(mv("ship", 5, 10, { ...at(6) }), /ส่งกำจัดเกินยอดคงเหลือ/);
  okData(mv("ship", 5, 10, { ...at(2) }));
  assert.strictEqual(bal(5).balance_kg, 40);
});

test("ship of a store-only field is ignored; store ignores manifest/vehicle", () => {
  const r = okData(mv("store", 6, 20, { manifest_no: "X", vehicle: "Y" }));
  const row = sheet("movements").data.find(x => x[1] === r.doc_no);
  assert.strictEqual(row[8], "");
  assert.strictEqual(row[9], "");
});

/* ---------- opening ---------- */
test("opening balance (ยอดยกมา) adds to balance but not to 'stored this month'", () => {
  const dash0 = okData(get("dashboard")).kpi;
  const r = okData(mv("opening", 13, 54000, { container: "ibc", containers: 54, ...at(1) }));
  assert.strictEqual(r.doc_no, "OB-" + YMD + "-0001");
  const b = bal(13);
  assert.strictEqual(b.balance_kg, 54000);
  assert.strictEqual(b.opening_kg, 54000);
  assert.strictEqual(b.stored_kg, 0);
  assert.strictEqual(b.in_kg, 54000);
  const dash = okData(get("dashboard")).kpi;
  assert.strictEqual(dash.month_in_kg, dash0.month_in_kg);
  assert.strictEqual(Math.round((dash.balance_kg - dash0.balance_kg) * 100) / 100, 54000);
});

/* ---------- quota ---------- */
test("quota %: shipped this year vs permit (tons -> kg), warn at >= 80 %, over at > 100 %", () => {
  // type 7: หลอดไฟ 0.5 t = 500 kg
  okData(mv("store", 7, 700, { containers: 5, ...at(3) }));
  okData(mv("ship", 7, 399, { ...at(2) }));
  let b = bal(7);
  assert.strictEqual(b.permit_kg, 500);
  assert.strictEqual(b.quota_pct, 79.8);
  assert.strictEqual(b.quota_level, "ok");
  okData(mv("ship", 7, 1, { ...at(2) }));
  b = bal(7);
  assert.strictEqual(b.quota_pct, 80);
  assert.strictEqual(b.quota_level, "warn");
  okData(mv("ship", 7, "100.01", { ...at(1) }));
  b = bal(7);
  assert.strictEqual(b.shipped_year_kg, 500.01);
  assert.strictEqual(b.quota_level, "over");
  const d = okData(get("dashboard"));
  const a = d.alerts.find(x => x.type_id === 7);
  assert.ok(a && a.kind === "quota" && a.level === "bad");
  assert.strictEqual(d.alerts[0].level, "bad"); // over-quota alerts first
  const bwg = d.receivers.find(r => r.receiver === "บมจ. ผู้รับ ทดสอบ บี");
  assert.strictEqual(bwg.permit_kg, 83500);
  assert.strictEqual(bwg.types, 7);
  assert.strictEqual(bwg.shipped_year_kg, 500.01 + 10); // + 10 kg of type 5
  assert.strictEqual(d.kpi.permit_kg, 194000);
  assert.deepStrictEqual(d.receivers.map(r => r.permit_kg), [60000, 83500, 50500]);
});

test("dashboard: 12 months oldest first, KPI totals, containers in storage, recent list", () => {
  const d = okData(get("dashboard"));
  assert.strictEqual(d.months.length, 12);
  assert.strictEqual(d.months[11].ym, TODAY.slice(0, 7));
  assert.ok(d.months[0].ym < d.months[11].ym);
  const b = boot().balance;
  const sum = Object.values(b).reduce((a, x) => a + Math.round(x.balance_kg * 100), 0) / 100;
  assert.strictEqual(d.kpi.balance_kg, sum);
  const yearOut = Object.values(b).reduce((a, x) => a + Math.round(x.shipped_year_kg * 100), 0) / 100;
  assert.strictEqual(d.kpi.year_out_kg, yearOut);
  assert.strictEqual(d.kpi.quota_pct, Math.round(yearOut / 194000 * 1000) / 10);
  assert.strictEqual(d.containers.ibc, Object.values(b).reduce((a, x) => a + x.containers.ibc, 0));
  assert.strictEqual(d.recent.length, 10);
  assert.ok(d.recent[0].moved_at >= d.recent[9].moved_at);
  assert.ok(d.logsToday > 0);
});

test("storage age (FIFO): oldest lot still in storage, alert after 90 days", () => {
  okData(mv("store", 8, 10, { ...at(120, 0) }));
  okData(mv("store", 8, 10, { ...at(30, 0) }));
  let b = bal(8);
  assert.strictEqual(b.age_days, 120);
  okData(mv("ship", 8, 10, { ...at(20) })); // ships the oldest lot first
  b = bal(8);
  assert.strictEqual(b.age_days, 30);
  okData(mv("store", 10, 5, { ...at(100, 0) }));
  const d = okData(get("dashboard"));
  const a = d.alerts.find(x => x.type_id === 10 && x.kind === "age");
  assert.ok(a && a.days === 100 && a.balance_kg === 5);
});

/* ---------- history / CSV ---------- */
test("history: newest first with type info; filters kind / type / receiver / dates / search", () => {
  const all = okData(get("history")).items;
  assert.strictEqual(all.length, dataRows("movements"));
  for (let i = 1; i < all.length; i++) assert.ok(all[i - 1].moved_at >= all[i].moved_at);
  assert.ok(all.every(m => m.code && m.type_name && m.receiver && m.icon && m.color));
  assert.ok(okData(get("history", { kind: "ship" })).items.every(m => m.kind === "ship"));
  assert.strictEqual(okData(get("history", { kind: "opening" })).items.length, 1);
  assert.ok(okData(get("history", { type: "7" })).items.every(m => m.type_id === 7));
  const g = okData(get("history", { receiver: "บมจ. ผู้รับ ทดสอบ ซี" })).items;
  assert.ok(g.length && g.every(m => [11, 12, 13].includes(m.type_id)));
  assert.strictEqual(okData(get("history", { q: "MN-1" })).items.length, 1);
  assert.ok(okData(get("history", { q: "หลอดไฟ" })).items.length >= 4);
  const d = at(30).date;
  assert.ok(okData(get("history", { from: d, to: d })).items.every(m => m.moved_at.slice(0, 10) === d));
  assert.strictEqual(okData(get("history", { from: "2001-01-01", to: "2001-01-02" })).items.length, 0);
});

test("CSV export: Thai headers, codes keep leading zeros, formula-like text neutralised", () => {
  okData(mv("store", 5, 1, { note: "=HYPERLINK(\"http://x\")", location: "+cmd", ...at(1) })); // also: never written as a formula
  const r = okData(get("export", { type: "5" }));
  assert.ok(/^waste_thepwatana_history_\d{8}_\d{6}\.csv$/.test(r.filename));
  const lines = r.csv.trim().split("\n");
  assert.ok(lines[0].startsWith('เลขที่เอกสาร,วันที่,เวลา,รายการ,รหัสของเสีย,ชื่อของเสีย,รหัสการจัดการ,ผู้รับดำเนินการ,"น้ำหนัก (ตัน)",ภาชนะ,จำนวนภาชนะ'), lines[0]);
  assert.ok(lines[0].endsWith("ผู้บันทึก"));
  assert.ok(r.csv.includes('"=""080409"""'), "code as Excel text");
  assert.ok(r.csv.includes('"=""042"""'));
  assert.ok(r.csv.includes("'=HYPERLINK") && r.csv.includes("'+cmd"));
  assert.ok(!/,=HYPERLINK/.test(r.csv));
  assert.ok(r.csv.includes(",0.001,")); // 1 kg in tons
  assert.strictEqual(lines.length - 1, okData(get("history", { type: "5" })).items.length);
  const row = sheet("movements").data.find(x => String(x[10]).includes("HYPERLINK"));
  assert.strictEqual(row[10], "=HYPERLINK(\"http://x\")"); // stored as text via apostrophe
});

/* ---------- in-house data (ข้อมูลภายใน, kind "treat"): rows need a password, totals are public ---------- */
// a read with a given session token
const tread = (tok, action, params = {}) => okData(postRaw(Object.assign({ action, token: tok }, params), null));
let TR1 = null;

test("treat: recording needs staff or admin (viewer FORBIDDEN, anonymous / bad token AUTH_REQUIRED, nothing written)", () => {
  const before = dataRows("movements");
  let r = err(mv("treat", 11, 5, { token: TV }), /ต้องเป็น “พนักงานกรอกข้อมูล”/);
  assert.strictEqual(r.code, "FORBIDDEN");
  for (const tok of [undefined, "", "a".repeat(64), TA + "0"]) {
    r = err(mv("treat", 11, 5, { token: tok }), /ต้องใส่รหัสผ่าน “พนักงานกรอกข้อมูล”/);
    assert.strictEqual(r.code, "AUTH_REQUIRED");
  }
  assert.strictEqual(dataRows("movements"), before);
  assert.strictEqual(boot().treat, false, "anonymous: no in-house rows");
  for (const tok of [TV, TS, TA]) assert.strictEqual(tread(tok, "bootstrap").treat, true, "every level sees the rows");
  assert.strictEqual(tread(TA, "bootstrap", { hideTreat: true }).treat, false, "a page may ask to hide them");
});

test("treat: lowers the in-plant balance like a shipment but is never counted in the สก.2 quota", () => {
  okData(mv("store", 12, 300, { container: "drum200", containers: 3, ...at(6) }));
  okData(mv("ship", 12, 100, { container: "drum200", containers: 1, ...at(4) }));
  const d0 = okData(get("dashboard"));
  const r = okData(mv("treat", 12, 150, { container: "drum200", containers: 1, method: "ระบบบำบัดน้ำเสียภายในโรงงาน", note: "ทดสอบ", ...at(3) }));
  TR1 = r;
  assert.strictEqual(r.doc_no, "TR-" + YMD + "-0001");
  assert.strictEqual(r.kind, "treat");
  assert.strictEqual(r.balance_kg, 50);
  assert.ok(/^ข้อมูลภายใน .* 0.15 ตัน เรียบร้อย/.test(r.message), r.message);
  const row = sheet("movements").data.find(x => x[1] === r.doc_no);
  assert.strictEqual(row[2], "treat");
  assert.strictEqual(row[13], ACTOR, "recorded under the name set on the page");
  assert.strictEqual(row[16], "ระบบบำบัดน้ำเสียภายในโรงงาน"); // method: appended column
  assert.strictEqual(row[8], ""); // no manifest / vehicle for a treatment
  // public (no password): true balance, quota counts the shipment only, in-house totals shown
  const b = bal(12);
  assert.strictEqual(b.balance_kg, 50);
  assert.strictEqual(b.containers.drum200, 1);
  assert.strictEqual(b.shipped_kg, 100);
  assert.strictEqual(b.shipped_year_kg, 100);
  assert.strictEqual(b.permit_kg, 500);
  assert.strictEqual(b.quota_pct, 20);
  assert.ok(!("treated_kg" in b) && !("treated_year_kg" in b), "public balance: no treatment total");
  assert.strictEqual(boot().treat, false);
  const bt = tread(TV, "bootstrap").balance[12];
  assert.strictEqual(bt.treated_kg, 150);
  assert.strictEqual(bt.treated_year_kg, 150);
  assert.strictEqual(Math.round((bt.in_kg - bt.shipped_kg - bt.treated_kg) * 100) / 100, bt.balance_kg);
  assert.strictEqual(bt.balance_kg, b.balance_kg, "balances identical with and without a token");
  assert.strictEqual(tread(TV, "bootstrap", { hideTreat: true }).balance[12].treated_kg, 150, "hideTreat hides rows only, not the totals");
  // public dashboard: the summary card (all = บำบัดเอง + ส่งกำจัด + คงเหลือ) and the monthly series
  const dPub = okData(get("dashboard"));
  for (const k of ["treated_kg", "month_treat_kg", "year_treat_kg", "total_in_kg", "shipped_kg"]) assert.ok(!(k in dPub.kpi), "public kpi: " + k);
  assert.ok(dPub.months.every(m => !("treat_kg" in m)), "public monthly series: no treat");
  assert.ok(!/treat/i.test(JSON.stringify(dPub)), "public dashboard never mentions treatment");
  const d = tread(TV, "dashboard");
  assert.strictEqual(dPub.kpi.balance_kg, d.kpi.balance_kg);
  assert.deepStrictEqual(dPub.months.map(m => [m.in_kg, m.out_kg]), d.months.map(m => [m.in_kg, m.out_kg]));
  assert.deepStrictEqual(dPub.alerts, d.alerts);
  assert.deepStrictEqual(dPub.receivers, d.receivers);
  assert.strictEqual(tread(TA, "dashboard", { hideTreat: true }).kpi.treated_kg, d.kpi.treated_kg, "hideTreat does not hide the totals");
  assert.strictEqual(d.kpi.year_out_kg, d0.kpi.year_out_kg, "quota use unchanged by a treatment");
  assert.strictEqual(d.kpi.quota_pct, d0.kpi.quota_pct);
  assert.deepStrictEqual(d.receivers.map(x => x.shipped_year_kg), d0.receivers.map(x => x.shipped_year_kg));
  assert.strictEqual(dPub.kpi.year_out_kg, d0.kpi.year_out_kg);
  assert.strictEqual(d.kpi.treated_kg >= 150, true);
  const k = d.kpi;
  assert.strictEqual(Math.round(k.total_in_kg * 100), Math.round(k.shipped_kg * 100) + Math.round(k.treated_kg * 100) + Math.round(k.balance_kg * 100));
  assert.ok(k.year_treat_kg >= 150 && k.month_treat_kg >= 0);
  assert.ok(d.months.reduce((a, m) => a + m.treat_kg, 0) >= 150);
  assert.ok(!d.alerts.some(a => a.type_id === 12 && a.kind === "quota"));
  const l = logs()[0];
  assert.strictEqual(l.action, "treat");
  assert.strictEqual(l.ref, r.doc_no);
  assert.ok(l.detail.startsWith("ข้อมูลภายใน ") && l.detail.includes("0.15 ตัน") && l.detail.includes("ระบบ / วิธีการ: ระบบบำบัดน้ำเสียภายในโรงงาน") && l.actor === ACTOR, l.detail);
});

test("treat: cannot exceed the balance or containers (chronological replay); without access the error never names the in-house row", () => {
  const before = dataRows("movements");
  let r = err(mv("treat", 12, "50.01", { ...at(2) }), /ข้อมูลภายในเกินยอดคงเหลือ/);
  assert.strictEqual(r.code, "NEGATIVE_BALANCE");
  err(mv("treat", 12, 10, { containers: 2, ...at(2) }), /ถัง 200 ลิตร.*มี 1 ภาชนะ แต่ต้องนำออก 2/);
  err(mv("treat", 12, 10, { container: "ibc", ...at(2) }), /IBC/);
  err(mv("treat", 4, 1, { ...at(2) }), /มี 0.00 ตัน/);
  err(mv("treat", 12, 10, { ...at(7) }), /ข้อมูลภายในเกินยอดคงเหลือ/); // before the stock existed
  err(mv("treat", 12, 10, { method: "x".repeat(121), ...at(2) }), /ระบบ \/ วิธีการยาวเกิน 120/);
  // a back-dated (public) shipment that would make the later in-house row negative
  r = err(mv("ship", 12, 60, { ...at(5) }), /ติดลบในรายการภายหลัง/); // anonymous: the in-house row is not named
  assert.ok(!/TR-|ภายใน\)/.test(r.error), r.error);
  for (const tok of [TV, TS, TA]) err(mv("ship", 12, 60, { token: tok, ...at(5) }), /ติดลบที่รายการ TR-.*\(ข้อมูลภายใน\)/);
  err(mv("ship", 12, 60, { token: TA, hideTreat: 1, ...at(5) }), /ติดลบในรายการภายหลัง/); // hidden on that page
  assert.strictEqual(dataRows("movements"), before);
  okData(mv("treat", 12, "0.5", { ...at(2) }));
  assert.strictEqual(bal(12).balance_kg, 49.5);
});

test("in-house rows: only with a password (any level); dashboard identical for everyone; totals public", () => {
  // anonymous (and a bad / expired token): no rows anywhere
  for (const tok of [undefined, "f".repeat(64)]) {
    const all = tread(tok, "history").items;
    assert.ok(all.length > 0 && all.every(m => m.kind !== "treat"));
    assert.strictEqual(all.length, dataRows("movements") - 2);
    assert.strictEqual(tread(tok, "history", { kind: "treat" }).items.length, 0);
    assert.strictEqual(tread(tok, "history", { q: "บำบัดน้ำเสีย" }).items.length, 0);
    assert.strictEqual(tread(tok, "history", { q: "TR-" }).items.length, 0);
    const bh = tread(tok, "batch", { reads: [{ action: "history", kind: "treat" }] });
    assert.strictEqual(bh.results[0].items.length, 0);
    const c0 = tread(tok, "export", { type: "12" }).csv;
    assert.ok(!/TR-|ภายในโรงงาน|ระบบ \/ วิธีการ/.test(c0), "CSV without in-house rows / column");
    assert.ok(!/TR-|ระบบบำบัดน้ำเสียภายในโรงงาน/.test(JSON.stringify(tread(tok, "bootstrap"))));
  }
  // the dashboard: totals only with a token (any level, same answer), rows never
  const dAnon = okData(get("dashboard"));
  const dTok = tread(TV, "dashboard");
  for (const tok of [TS, TA]) assert.deepStrictEqual(tread(tok, "dashboard"), dTok);
  assert.ok(dTok.kpi.treated_kg >= 150.5 && dTok.months.some(m => m.treat_kg > 0));
  assert.ok(!("treated_kg" in dAnon.kpi) && dAnon.months.every(m => !("treat_kg" in m)));
  assert.ok(!dTok.recent.some(m => m.kind === "treat") && !/TR-/.test(JSON.stringify(dTok)) && !/TR-/.test(JSON.stringify(dAnon)));
  // a bad / expired token is anonymous: no totals anywhere
  assert.ok(!("treated_kg" in tread("a".repeat(64), "dashboard").kpi));
  assert.ok(!("treated_kg" in tread("a".repeat(64), "bootstrap").balance[12]));
  for (const a of ["bootstrap", "history", "export", "dashboard", "batch"]) {
    const txt = JSON.stringify(a === "batch" ? okData(get("batch", { reads: JSON.stringify(["bootstrap", "dashboard", "history"]) })) : okData(get(a, a === "history" ? {} : a === "export" ? { type: "12" } : {})));
    assert.ok(!/treated|treat_|total_in_kg|month_treat|TR-/.test(txt), "public " + a);
  }
  // a client flag never widens access
  assert.strictEqual(tread(undefined, "bootstrap", { hideTreat: false, treat: true, role: "admin" }).treat, false);
  // view / staff / admin: the rows, the details and the CSV column
  for (const tok of [TV, TS, TA]) {
    const th = tread(tok, "history", { kind: "treat" }).items;
    assert.strictEqual(th.length, 2);
    assert.strictEqual(th[1].method, "ระบบบำบัดน้ำเสียภายในโรงงาน");
    assert.strictEqual(th[1].actor, ACTOR);
    assert.strictEqual(tread(tok, "history").items.length, dataRows("movements"));
    assert.strictEqual(tread(tok, "history", { q: "บำบัดน้ำเสีย" }).items.length, 1);
    const bt = tread(tok, "batch", { reads: [{ action: "bootstrap" }, { action: "history", kind: "treat" }] });
    assert.strictEqual(bt.results[0].treat, true);
    assert.strictEqual(bt.results[1].items.length, 2);
    const c1 = tread(tok, "export", { type: "12" }).csv;
    const lines = c1.trim().split("\n");
    assert.ok(lines[0].includes(",\"น้ำหนัก (ตัน)\",") && lines[0].endsWith(",ผู้บันทึก,\"ระบบ / วิธีการ\""), lines[0]);
    const tr = lines.find(x => x.startsWith(TR1.doc_no));
    assert.ok(tr && tr.includes(",ข้อมูลภายใน,") && tr.includes(",ภายในโรงงาน,0.15,") && tr.endsWith(",ระบบบำบัดน้ำเสียภายในโรงงาน"), tr);
    assert.strictEqual(lines.length - 1, tread(tok, "history", { type: "12" }).items.length);
    // hidden on request (the page does this outside the ข้อมูลภายใน page)
    assert.ok(tread(tok, "history", { hideTreat: true }).items.every(m => m.kind !== "treat"));
  }
  // a write carrying reads answers in the mode of its request
  const w0 = okData(mv("store", 12, 1, { ...at(1), reads: [{ action: "bootstrap" }, { action: "history", kind: "treat" }] }));
  assert.strictEqual(w0.reads[0].treat, false);
  assert.strictEqual(w0.reads[0].role, null);
  assert.ok(!("treated_kg" in w0.reads[0].balance[12]), "no totals without a token");
  assert.strictEqual(w0.reads[1].items.length, 0);
  const w1 = okData(mv("store", 12, 1, { ...at(1), token: TS, reads: [{ action: "bootstrap" }, { action: "history", kind: "treat" }] }));
  assert.strictEqual(w1.reads[0].treat, true);
  assert.strictEqual(w1.reads[0].role, "staff");
  assert.strictEqual(w1.reads[0].balance[12].treated_kg, 150.5);
  assert.strictEqual(w1.reads[1].items.length, 2);
  // the dashboard "logs today" count never counts in-house audit rows (same for everyone)
  assert.strictEqual(okData(get("dashboard")).logsToday, tread(TA, "dashboard").logsToday);
});

test("treat rows: edit / delete by admin like any row (replay checked, audited); view / staff cannot", () => {
  const id = TR1.id;
  assert.strictEqual(err(post("movement_update", { id, weight_t: "0.1", token: TS }), /ผู้ดูแลระบบ/).code, "FORBIDDEN");
  assert.strictEqual(err(post("movement_delete", { id, token: TV }), /ผู้ดูแลระบบ/).code, "FORBIDDEN");
  assert.strictEqual(err(post("movement_delete", { id, token: undefined }), /ผู้ดูแลระบบ/).code, "AUTH_REQUIRED");
  // stock of type 12: 300 + 1 + 1 in, 100 shipped, 150 + 0.5 treated -> 51.5 left
  err(post("movement_update", { id, doc_no: TR1.doc_no, weight_t: "0.20151" }), /แก้ไขไม่ได้/);
  err(post("movement_update", { id, kind: "ship" }), /เปลี่ยนชนิดรายการไม่ได้/);
  const ok = okData(post("movement_update", { id, doc_no: TR1.doc_no, weight_t: "0.120", method: "เตาเผาภายในโรงงาน" }));
  assert.strictEqual(ok.changed, true);
  assert.strictEqual(bal(12).balance_kg, 81.5);
  assert.strictEqual(bal(12).quota_pct, 20);
  const row = sheet("movements").data.find(x => x[1] === TR1.doc_no);
  assert.strictEqual(row[4], 120);
  assert.strictEqual(row[16], "เตาเผาภายในโรงงาน");
  const l = lastLog();
  assert.strictEqual(l.action, "movement_update");
  assert.ok(l.detail.includes("ระบบบำบัดน้ำเสียภายในโรงงาน") && l.detail.includes("เตาเผาภายในโรงงาน") && l.detail.includes("0.12 ตัน"));
  // a store the in-house row depends on cannot be deleted; the in-house row itself can
  const b6 = bal(6).balance_kg;
  const st = okData(mv("store", 6, 40, { container: "ibc", containers: 1, ...at(9) }));
  const tr = okData(mv("treat", 6, 40, { container: "ibc", containers: 1, ...at(8) }));
  err(post("movement_delete", { id: st.id, doc_no: st.doc_no, hideTreat: true }), /ลบไม่ได้.*ติดลบในรายการภายหลัง/);
  okData(post("movement_delete", { id: tr.id, doc_no: tr.doc_no }));
  assert.strictEqual(lastLog().action, "movement_delete");
  okData(post("movement_delete", { id: st.id, doc_no: st.doc_no }));
  assert.strictEqual(bal(6).balance_kg, b6);
});

test("migration: an older movements sheet without the method column keeps working (column appended on first use)", () => {
  const g = createGas();
  const c2 = loadCode(g.gas);
  c2.setupSystem();
  Object.assign(g.state.props, PROPS);
  const ss = g.state.spreadsheets[g.state.props.WASTE_SPREADSHEET_ID];
  const ms = ss.getSheetByName("movements");
  const p = body => JSON.parse(c2.doPost({ postData: { contents: JSON.stringify(Object.assign({ actor: ACTOR }, body)) } }).getContent());
  const tok = okData(p({ action: "login", password: RPW })).token;
  seedFixtureTypes(b => p(Object.assign({ token: tok }, b)));
  okData(p({ action: "store", token: tok, type_id: 13, weight_t: "0.5", container: "ibc", containers: 1, ...at(3) }));
  // simulate the sheet as created by the previous version: no "method" column
  ms.data.forEach(row => { row.length = 16; });
  assert.ok(!ms.data[0].includes("method"));
  c2.clearReadCache();
  assert.strictEqual(okData(p({ action: "bootstrap", token: tok })).balance[13].balance_kg, 500);
  assert.strictEqual(ms.data[0][16], "method", "column appended to the header");
  const r = okData(p({ action: "treat", token: tok, type_id: 13, weight_t: "0.2", container: "ibc", containers: 1, method: "ระบบบำบัดน้ำเสียภายในโรงงาน", ...at(2) }));
  assert.strictEqual(r.balance_kg, 300);
  const row = ms.data.find(x => x[1] === r.doc_no);
  assert.strictEqual(row[16], "ระบบบำบัดน้ำเสียภายในโรงงาน");
  assert.strictEqual(ms.data[1][16] || "", ""); // the old row reads the new column as empty
});

/* ---------- login, sessions, roles ---------- */
test("login: each password gives its level; token 64 hex in CacheService (12 h session, 6 h cache entry); no secrets stored", () => {
  clearFails();
  const puts = state.cachePuts.length;
  const now = Date.now();
  for (const [pw, role, name] of [[VPW, "viewer", "ดูได้อย่างเดียว"], [SPW, "staff", "พนักงานกรอกข้อมูล"], [RPW, "admin", "ผู้ดูแลระบบ"]]) {
    const d = okData(loginRes(pw));
    assert.ok(/^[0-9a-f]{64}$/.test(d.token));
    assert.strictEqual(d.role, role);
    assert.strictEqual(d.role_name, name);
    assert.strictEqual(d.ttl, 43200);
    assert.ok(Math.abs(d.expires_at - (now + 43200e3)) < 5000);
    const rec = JSON.parse(state.cache["WASTE_SS_" + d.token]);
    assert.strictEqual(rec.r, role);
    assert.ok(!JSON.stringify(rec).includes(pw), "password not in the session entry");
    assert.strictEqual(tread(d.token, "bootstrap").role, role);
    assert.strictEqual(tread(d.token, "bootstrap").role_name, name);
  }
  assert.ok(state.cachePuts.slice(puts).every(t => t === 21600), "session entries kept 6 h (CacheService maximum)");
  // wrong / odd passwords
  for (const bad of ["", "nope", " " + VPW, VPW.toUpperCase(), null, 123, { p: VPW }, [SPW]]) {
    const r = err(loginRes(bad), /รหัสไม่ถูกต้อง/);
    assert.strictEqual(r.passwordError, true);
  }
  clearFails();
  // the same password set for two levels: the higher level wins
  state.props.WASTE_STAFF_PASSWORD = VPW;
  assert.strictEqual(okData(loginRes(VPW)).role, "staff");
  state.props.WASTE_STAFF_PASSWORD = SPW;
});

test("roles: every action x every level, enforced on the server (open site; ข้อมูลภายใน write staff+; admin tools)", () => {
  const rows0 = dataRows("movements"), logs0 = dataRows("audit_logs");
  const allowed = {
    anonymous: PUBLIC_ACTIONS, viewer: PUBLIC_ACTIONS,
    staff: PUBLIC_ACTIONS.concat(STAFF_ONLY), admin: ALL_ACTIONS
  };
  for (const [lv, tok] of [["anonymous", undefined], ["viewer", TV], ["staff", TS], ["admin", TA]]) {
    for (const a of ALL_ACTIONS) {
      // writes without a name stop right after the access check (nothing written)
      const r = postRaw({ action: a, token: tok, reads: a === "batch" ? [{ action: "bootstrap" }] : undefined }, null);
      if (!allowed[lv].includes(a)) {
        assert.strictEqual(r.ok, false, lv + " " + a);
        assert.strictEqual(r.code, lv === "anonymous" ? "AUTH_REQUIRED" : "FORBIDDEN", lv + " " + a + ": " + r.error);
        assert.ok(lv === "anonymous" ? /ต้องใส่รหัสผ่าน “(พนักงานกรอกข้อมูล|ผู้ดูแลระบบ)”/.test(r.error)
          : /ต้องเป็น “(พนักงานกรอกข้อมูล|ผู้ดูแลระบบ)”/.test(r.error), r.error);
      } else if (READS.includes(a) || a === "logs" || a === "logs_export") {
        assert.strictEqual(r.ok, true, lv + " " + a + ": " + r.error);
      } else {
        assert.strictEqual(r.code, "NAME_REQUIRED", lv + " " + a + ": " + r.error);
      }
    }
  }
  // the minimum level of each action, as the server enforces it
  const minRole = a => PUBLIC_ACTIONS.includes(a) ? "public" : STAFF_ONLY.includes(a) ? "staff" : "admin";
  assert.deepStrictEqual(Object.keys(ctx.ACTION_ROLE).sort(), ALL_ACTIONS.slice().sort());
  ALL_ACTIONS.forEach(a => assert.strictEqual(ctx.ACTION_ROLE[a], minRole(a), a));
  // a public write really works without any password; staff cannot use the admin tools
  okData(mv("store", 1, 1, { token: undefined, container: "bigbag", containers: 1, ...at(1) }));
  err(post("type_save", { token: TS, code: "123456", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "r" }), /ผู้ดูแลระบบ/);
  err(post("reset_data", { token: TS }), /ผู้ดูแลระบบ/);
  err(post("logs", { token: TS }), /ผู้ดูแลระบบ/);
  err(post("movement_delete", { token: TS, id: 1 }), /ผู้ดูแลระบบ/);
  assert.strictEqual(dataRows("movements"), rows0 + 1);
  assert.strictEqual(dataRows("audit_logs"), logs0 + 1, "refused requests are not audited (no lock taken)");
});

test("sessions: no / bad / unknown / expired token -> AUTH_REQUIRED on password actions, anonymous on public ones; logout ends a session", () => {
  const before = dataRows("movements");
  for (const tok of [undefined, "", "x", "a".repeat(64), "A".repeat(64), { t: TA }, TA + "0"]) {
    for (const a of STAFF_ONLY.concat(ADMIN_ONLY)) {
      const r = postRaw({ action: a, token: tok, type_id: 1, weight_t: "1", container: "drum200", containers: 1, resetPassword: RPW });
      assert.strictEqual(r.code, "AUTH_REQUIRED", a + " " + JSON.stringify(tok) + ": " + JSON.stringify(r));
      assert.ok(/ต้องใส่รหัสผ่าน/.test(r.error));
    }
    assert.strictEqual(roleOf(tok), null, "public read: anonymous");
    assert.strictEqual(okData(postRaw({ action: "history", kind: "treat", token: tok }, null)).items.length, 0);
  }
  assert.strictEqual(dataRows("movements"), before);
  // expired (absolute 12 h)
  const t = login(SPW);
  const rec = JSON.parse(state.cache["WASTE_SS_" + t]);
  rec.e = Date.now() - 1;
  state.cache["WASTE_SS_" + t] = JSON.stringify(rec);
  assert.strictEqual(postRaw({ action: "treat", token: t }).code, "AUTH_REQUIRED");
  assert.strictEqual(roleOf(t), null);
  // evicted / corrupt entry
  const t2 = login(VPW);
  state.cache["WASTE_SS_" + t2] = "{bad json";
  assert.strictEqual(roleOf(t2), null);
  state.cache["WASTE_SS_" + t2] = JSON.stringify({ r: "root", e: Date.now() + 1e6, p: Date.now(), f: "x" });
  assert.strictEqual(roleOf(t2), null);
  // logout
  const t3 = login(RPW);
  assert.strictEqual(roleOf(t3), "admin");
  assert.strictEqual(okData(rawPost({ action: "logout", token: t3 })).loggedOut, true);
  assert.ok(!("WASTE_SS_" + t3 in state.cache));
  assert.strictEqual(postRaw({ action: "logs", token: t3 }).code, "AUTH_REQUIRED");
  assert.strictEqual(roleOf(t3), null);
  okData(rawPost({ action: "logout", token: "nope" })); // harmless
  okData(rawPost({ action: "logout" }));
  // cache outage: sessions cannot be checked -> no access (never "unlocked by default"), public reads still answer
  state.cacheDown = true;
  try {
    assert.strictEqual(postRaw({ action: "logs", token: TA }).code, "AUTH_REQUIRED");
    assert.strictEqual(roleOf(TA), null);
    err(loginRes(VPW));
  } finally { state.cacheDown = false; }
  assert.strictEqual(roleOf(TA), "admin"); // fine again afterwards
});

test("sessions: the cache entry is renewed while in use (12 h > 6 h cache limit), never past the 12 h end", () => {
  const t = login(VPW);
  const key = "WASTE_SS_" + t;
  let rec = JSON.parse(state.cache[key]);
  const p0 = rec.p;
  const puts = state.cachePuts.length;
  tread(t, "bootstrap");
  assert.strictEqual(JSON.parse(state.cache[key]).p, p0, "no renewal within the first hour");
  assert.strictEqual(state.cachePuts.length, puts);
  rec.p = Date.now() - 2 * 3600e3;
  rec.e = Date.now() + 3 * 3600e3; // 3 h left
  state.cache[key] = JSON.stringify(rec);
  tread(t, "dashboard");
  rec = JSON.parse(state.cache[key]);
  assert.ok(Date.now() - rec.p < 5000, "renewed");
  const ttl = state.cachePuts[state.cachePuts.length - 1];
  assert.ok(ttl > 10700 && ttl <= 10800, "renewed only up to the session end: " + ttl);
  assert.ok(rec.e < Date.now() + 3 * 3600e3 + 1000, "absolute end unchanged");
});

test("sessions: changing or removing a password ends the sessions of that level only; unset level cannot unlock", () => {
  const tv = login(VPW), ts = login(SPW);
  state.props.WASTE_STAFF_PASSWORD = "staff-pass-2";
  assert.strictEqual(roleOf(ts), null);
  assert.strictEqual(tread(tv, "bootstrap").role, "viewer");
  err(loginRes(SPW), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(okData(loginRes("staff-pass-2")).role, "staff");
  state.props.WASTE_STAFF_PASSWORD = SPW;
  clearFails();
  delete state.props.WASTE_VIEW_PASSWORD;
  assert.strictEqual(roleOf(tv), null);
  err(loginRes(VPW), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(okData(loginRes(SPW)).role, "staff"); // other levels still work
  state.props.WASTE_VIEW_PASSWORD = VPW;
  clearFails();
  assert.strictEqual(tread(TV, "bootstrap").role, "viewer");
});

test("login lockout: 5 wrong per browser (client id) and 30 wrong overall per 15 min lock even the right password; open sessions keep working", () => {
  clearFails();
  const A = "0123456789abcdef", B = "fedcba9876543210";
  const s0 = state.sleeps;
  for (let i = 0; i < 5; i++) err(loginRes("wrong" + i, A), /รหัสไม่ถูกต้อง/);
  assert.ok(state.sleeps - s0 >= 5000, "1 s per wrong attempt");
  let r = err(loginRes(RPW, A), /หลายครั้งเกินไป/);
  assert.strictEqual(r.passwordError, true);
  assert.strictEqual(state.cache["WASTE_LF_" + A], "5");
  assert.ok(state.cachePuts.includes(900));
  assert.strictEqual(okData(loginRes(VPW, B)).role, "viewer", "another browser is not locked");
  // a right password clears that browser's counter
  for (let i = 0; i < 3; i++) err(loginRes("wrong" + i, B), /รหัสไม่ถูกต้อง/);
  okData(loginRes(VPW, B));
  assert.ok(!("WASTE_LF_" + B in state.cache));
  // a client id that is not hex is ignored (global counter only)
  err(loginRes("wrong", "../evil"), /รหัสไม่ถูกต้อง/);
  assert.ok(!Object.keys(state.cache).some(k => k.includes("evil")));
  clearFails();
  for (let i = 0; i < 30; i++) err(loginRes("wrong" + i), /รหัสไม่ถูกต้อง/);
  r = err(loginRes(VPW, B), /หลายครั้งเกินไป/);
  err(loginRes(RPW), /หลายครั้งเกินไป/);
  assert.strictEqual(state.cache.WASTE_LOGIN_FAILURES, "30");
  assert.strictEqual(tread(TA, "bootstrap").role, "admin", "existing sessions are not affected");
  // reset_data shares the lockout
  err(post("reset_data", {}), /หลายครั้งเกินไป/);
  clearFails();
  assert.strictEqual(okData(loginRes(VPW)).role, "viewer");
});

test("reset_data: admin session + the admin password typed again (missing property, wrong password, lockout counter)", () => {
  const saved = state.props.WASTE_RESET_PASSWORD;
  delete state.props.WASTE_RESET_PASSWORD;
  assert.strictEqual(err(post("reset_data", {}), /ต้องใส่รหัสผ่าน “ผู้ดูแลระบบ”/).code, "AUTH_REQUIRED"); // the admin session ends with the password
  state.props.WASTE_RESET_PASSWORD = saved;
  clearFails();
  err(post("reset_data", { resetPassword: "bad", client: "00112233445566aa" }), /รหัสไม่ถูกต้อง/);
  err(post("reset_data", { resetPassword: undefined }), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(state.cache.WASTE_LOGIN_FAILURES, "2");
  assert.strictEqual(state.cache.WASTE_LF_00112233445566aa, "1");
  assert.ok(dataRows("movements") > 0, "nothing wiped");
  clearFails();
});

test("password never echoed in responses, logs or audit rows; no secrets in the source", () => {
  const all = JSON.stringify(state.logs) + JSON.stringify(sheet("audit_logs").data) + JSON.stringify(state.cache);
  for (const pw of [VPW, SPW, RPW]) assert.ok(!all.includes(pw), pw);
  assert.ok(!JSON.stringify(sheet("audit_logs").data).includes(TA), "tokens never in the audit log");
  const src = fs.readFileSync(path.join(ROOT, "apps-script", "Code.gs"), "utf8");
  assert.ok(!/_PASSWORD\s*=\s*["'][^"']+["']\s*;?\s*$/m.test(src.replace(/(RESET_PASSWORD_PROPERTY|[A-Z_]*PASSWORD[A-Z_]*) = "WASTE_[A-Z]+_PASSWORD"/g, "")));
  for (const s of ["dev-reset", "dev-staff", "dev-view", VPW, SPW, RPW]) assert.ok(!src.includes(s), s);
  assert.ok(!/AKfycb|https:\/\/hooks\./.test(src));
  for (const prop of ["WASTE_VIEW_PASSWORD", "WASTE_STAFF_PASSWORD", "WASTE_RESET_PASSWORD"]) assert.ok(src.includes('"' + prop + '"'), prop);
});

/* ---------- edit / delete integrity ---------- */
test("movement_update: edits with replay check; lowering a store that a later ship needs is refused", () => {
  const st = okData(mv("store", 9, 100, { container: "bigbag", containers: 2, ...at(8) }));
  const sh = okData(mv("ship", 9, 80, { container: "bigbag", containers: 1, ...at(4) }));
  let r = err(post("movement_update", { id: st.id, doc_no: st.doc_no, weight_t: "0.07999" }), /แก้ไขไม่ได้.*ติดลบที่รายการ SH-/);
  assert.strictEqual(r.code, "NEGATIVE_BALANCE");
  err(post("movement_update", { id: st.id, doc_no: st.doc_no, containers: 0 }), /จำนวนเต็มตั้งแต่ 1/);
  err(post("movement_update", { id: sh.id, doc_no: sh.doc_no, containers: 3 }), /แก้ไขไม่ได้/);
  err(post("movement_update", { id: st.id, ...at(3) }), /แก้ไขไม่ได้/); // moving the store after the ship
  err(post("movement_update", { id: st.id, type_id: 10 }), /แก้ไขไม่ได้/); // the ship of type 9 loses its stock
  err(post("movement_update", { id: st.id, doc_no: "ST-000000-9999", weight_t: "0.12" }), /ไม่ตรงกัน/);
  err(post("movement_update", { id: 99999, weight_t: "0.001" }), /ไม่พบรายการ/);
  err(post("movement_update", { id: st.id, kind: "ship" }), /เปลี่ยนชนิดรายการไม่ได้/);
  assert.strictEqual(okData(post("movement_update", { id: st.id, doc_no: st.doc_no })).changed, false);
  const ok = okData(post("movement_update", { id: st.id, doc_no: st.doc_no, weight_t: "0.1205", note: "แก้แล้ว", location: "B" }));
  assert.strictEqual(ok.changed, true);
  assert.strictEqual(bal(9).balance_kg, 40.5);
  const row = sheet("movements").data.find(x => x[1] === st.doc_no);
  assert.strictEqual(row[4], 120.5);
  assert.strictEqual(row[10], "แก้แล้ว");
  assert.strictEqual(row[15], ACTOR);
  assert.ok(row[14]);
  const l = lastLog();
  assert.strictEqual(l.category, "edit");
  assert.ok(l.detail.includes("เดิม:") && l.detail.includes("0.10 ตัน") && l.detail.includes("0.1205 ตัน"), l.detail);
  // a ship edited above the stock is refused, within it accepted
  err(post("movement_update", { id: sh.id, weight_t: "0.12051" }), /แก้ไขไม่ได้/);
  okData(post("movement_update", { id: sh.id, weight_t: 0.1205, manifest_no: "MN-77" }));
  assert.strictEqual(bal(9).balance_kg, 0);
});

test("movement_delete: a store that a ship depends on cannot be deleted; deleting the ship first works", () => {
  const st = okData(mv("store", 4, 3, { ...at(9) }));
  const sh = okData(mv("ship", 4, 3, { ...at(8) }));
  const before = dataRows("movements");
  err(post("movement_delete", { id: st.id, doc_no: st.doc_no }), /ลบไม่ได้.*ติดลบ/);
  assert.strictEqual(dataRows("movements"), before);
  const r = okData(post("movement_delete", { id: sh.id, doc_no: sh.doc_no }));
  assert.ok(/ลบรายการ/.test(r.message));
  assert.strictEqual(bal(4).balance_kg, 3);
  okData(post("movement_delete", { id: st.id, doc_no: st.doc_no }));
  assert.strictEqual(bal(4).balance_kg, 0);
  assert.strictEqual(dataRows("movements"), before - 2);
  const l = lastLog();
  assert.strictEqual(l.action, "movement_delete");
  assert.strictEqual(l.ref, st.doc_no);
  // a deleted doc number is never handed out again
  const again = okData(mv("ship", 1, 1, { container: "bigbag", containers: 1, ...at(1) }));
  assert.notStrictEqual(again.doc_no, sh.doc_no);
});

/* ---------- waste type settings ---------- */
test("type_save: validation, add, edit, duplicate check, deactivate (no new store, ship still allowed)", () => {
  err(post("type_save", { code: "12345", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "r" }), /6 หลัก/);
  err(post("type_save", { code: "123456", name: "", permit_tons: 1, mgmt_code: "011", receiver: "r" }), /ชื่อของเสีย/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: "abc", mgmt_code: "011", receiver: "r" }), /ตัวเลข/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "11", receiver: "r" }), /3 หลัก/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "" }), /ผู้รับดำเนินการ/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "r", receiver_reg: "12a" }), /ทะเบียนผู้รับ/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "r", icon: "evil\"><script>" }), /ไอคอน/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "r", color: "red" }), /สี/);
  err(post("type_save", { code: "150101", name: "กระดาษลูกฟูก กระดาษลัง", permit_tons: 5, mgmt_code: "011", receiver_reg: "1000000000001", receiver: "บจก. ผู้รับ ทดสอบ เอ" }), /อยู่แล้ว/);
  const a = okData(post("type_save", { code: "130205", name: "น้ำมันเครื่องใช้แล้ว", permit_tons: "1.25", mgmt_code: "061", receiver_reg: "123", receiver: "บจก. ทดสอบ", icon: "fa-oil-can", color: "#0ea5e9" }));
  assert.strictEqual(a.created, true);
  assert.strictEqual(a.id, 14);
  let t = boot().types.find(x => x.id === 14);
  assert.strictEqual(t.permit_tons, 1.25);
  assert.strictEqual(t.color, "#0EA5E9");
  assert.strictEqual(t.sort, 14);
  assert.strictEqual(t.updated_by, ACTOR);
  assert.strictEqual(bal(14).permit_kg, 1250);
  assert.strictEqual(lastLog().category, "setting");
  okData(mv("store", 14, 10, { ...at(2) }));
  const e = okData(post("type_save", { id: 14, code: "130205", name: "น้ำมันเครื่องใช้แล้ว", permit_tons: 2, mgmt_code: "061", receiver_reg: "123", receiver: "บจก. ทดสอบ", active: 0 }));
  assert.strictEqual(e.created, false);
  t = boot().types.find(x => x.id === 14);
  assert.strictEqual(t.active, 0);
  assert.strictEqual(t.permit_tons, 2);
  assert.strictEqual(t.icon, "fa-oil-can"); // kept when not sent
  err(mv("store", 14, 1), /ปิดใช้งาน/);
  okData(mv("ship", 14, 10, { ...at(1) }));
  assert.ok(/เดิม:.*1\.25 ตัน.*ใช้งาน → ใหม่:.*2 ตัน.*ปิดใช้งาน/.test(logs({ cat: "setting" })[0].detail));
  err(post("type_save", { id: 999, code: "130205", name: "x", permit_tons: 1, mgmt_code: "061", receiver: "r" }), /ไม่พบ/);
  assert.strictEqual(dataRows("waste_types"), 14);
});

/* ---------- settings: receivers (ผู้รับดำเนินการ) ---------- */
test("settings: permit tons accept 2 decimals only; a type with a NEW receiver creates it and shows in both dashboard quota sections", () => {
  err(post("type_save", { code: "130206", name: "สีเก่า", permit_tons: "1.255", mgmt_code: "061", receiver: "บจก. ใหม่" }), /ไม่เกิน 2 ตำแหน่ง/);
  const r0 = boot().receivers.length;
  const a = okData(post("type_save", { code: "130206", name: "สีเก่า", permit_tons: "12.5", mgmt_code: "061", receiver: "บจก. ผู้รับใหม่", receiver_reg: "99001" }));
  assert.strictEqual(a.receiver_created, true);
  const b = boot();
  assert.strictEqual(b.receivers.length, r0 + 1);
  const rc = b.receivers.find(x => x.name === "บจก. ผู้รับใหม่");
  assert.deepStrictEqual([rc.reg, rc.active, rc.types, rc.active_types, rc.permit_tons], ["99001", 1, 1, 1, 12.5]);
  const t = b.types.find(x => x.id === a.id);
  assert.strictEqual(t.receiver, "บจก. ผู้รับใหม่");
  assert.strictEqual(t.receiver_reg, "99001");
  assert.strictEqual(t.has_moves, false);
  // quota ring (per type) and receiver bar (dashboard) pick it up with no other change
  assert.strictEqual(b.balance[a.id].permit_kg, 12500);
  const d = okData(get("dashboard"));
  const bar = d.receivers.find(x => x.receiver === "บจก. ผู้รับใหม่");
  assert.deepStrictEqual([bar.types, bar.permit_kg, bar.shipped_year_kg, bar.quota_pct, bar.quota_level], [1, 12500, 0, 0, "ok"]);
  // store + ship: history filter by receiver, CSV, quota %
  okData(mv("store", a.id, 5000, { ...at(3) }));
  okData(mv("ship", a.id, 2500, { ...at(2) }));
  const bar2 = okData(get("dashboard")).receivers.find(x => x.receiver === "บจก. ผู้รับใหม่");
  assert.deepStrictEqual([bar2.shipped_year_kg, bar2.quota_pct, bar2.balance_kg], [2500, 20, 2500]);
  assert.strictEqual(okData(get("history", { receiver: "บจก. ผู้รับใหม่" })).items.length, 2);
  assert.ok(okData(get("export", { receiver: "บจก. ผู้รับใหม่" })).csv.includes("บจก. ผู้รับใหม่"));
  assert.strictEqual(boot().types.find(x => x.id === a.id).has_moves, true);
  // the second type of the same receiver by id adds to the receiver's quota
  const c = okData(post("type_save", { code: "130207", name: "ตัวทำละลาย", permit_tons: "7.25", mgmt_code: "061", receiver_id: rc.id }));
  assert.strictEqual(c.receiver_created, false);
  const rc2 = boot().receivers.find(x => x.id === rc.id);
  assert.deepStrictEqual([rc2.types, rc2.permit_tons], [2, 19.75]);
  assert.strictEqual(okData(get("dashboard")).receivers.find(x => x.receiver === "บจก. ผู้รับใหม่").permit_kg, 19750);
  err(post("type_save", { code: "130208", name: "x", permit_tons: 1, mgmt_code: "061", receiver_id: 9999 }), /ไม่พบผู้รับ/);
});

test("settings: receiver add / edit (name + registration cascade to types) / duplicates / validation", () => {
  err(post("receiver_save", { name: "", reg: "1" }), /ชื่อผู้รับ/);
  err(post("receiver_save", { name: "ก", reg: "12a" }), /ทะเบียน/);
  err(post("receiver_save", { name: "ก", reg: "1", mgmt_code: "12" }), /3 หลัก/);
  err(post("receiver_save", { name: " บจก.  ผู้รับใหม่ " }), /มีผู้รับดำเนินการชื่อนี้อยู่แล้ว/);
  const a = okData(post("receiver_save", { name: "หจก. ยังไม่มีประเภท", reg: "5550001", mgmt_code: "061" }));
  assert.strictEqual(a.created, true);
  let rc = boot().receivers.find(x => x.id === a.id);
  assert.deepStrictEqual([rc.types, rc.permit_tons, rc.mgmt_code], [0, 0, "061"]);
  // a receiver with no types: 0 quota bar on the dashboard
  const bar = okData(get("dashboard")).receivers.find(x => x.receiver === "หจก. ยังไม่มีประเภท");
  assert.deepStrictEqual([bar.types, bar.permit_kg, bar.quota_pct, bar.quota_level], [0, 0, null, "ok"]);
  // its default management code fills a type that leaves the field empty
  const t = okData(post("type_save", { code: "130209", name: "ทดสอบรหัสจัดการ", permit_tons: 3, mgmt_code: "", receiver_id: a.id }));
  assert.strictEqual(boot().types.find(x => x.id === t.id).mgmt_code, "061");
  // edit: rename + registration change reach the types of that receiver
  const e = okData(post("receiver_save", { id: a.id, name: "หจก. เปลี่ยนชื่อ", reg: "5550002", mgmt_code: "" }));
  assert.strictEqual(e.created, false);
  const b = boot();
  const ty = b.types.find(x => x.id === t.id);
  assert.deepStrictEqual([ty.receiver, ty.receiver_reg], ["หจก. เปลี่ยนชื่อ", "5550002"]);
  assert.strictEqual(b.receivers.find(x => x.id === a.id).name, "หจก. เปลี่ยนชื่อ");
  assert.ok(b.receivers.every(x => x.name !== "หจก. ยังไม่มีประเภท"));
  assert.ok(/เดิม: หจก\. ยังไม่มีประเภท \(5550001\).*→ ใหม่: หจก\. เปลี่ยนชื่อ \(5550002\)/.test(logs({ cat: "setting" })[0].detail));
  err(post("receiver_save", { id: a.id, name: "บจก. ผู้รับใหม่" }), /มีผู้รับดำเนินการชื่อนี้อยู่แล้ว/);
  err(post("receiver_save", { id: 9999, name: "x" }), /ไม่พบ/);
});

test("settings: deactivate / delete rules (receivers and types with movements are only deactivated)", () => {
  const b0 = boot();
  const used = b0.types.find(x => x.receiver === "บจก. ผู้รับใหม่" && x.has_moves);
  const rc = b0.receivers.find(x => x.name === "บจก. ผู้รับใหม่");
  // delete: refused while the type has movements / the receiver has types
  err(post("type_delete", { id: used.id }), /มีรายการเคลื่อนไหวแล้ว.*ปิดใช้งาน/);
  err(post("receiver_delete", { id: rc.id }), /ยังมีประเภทของเสียอยู่.*ปิดใช้งาน/);
  // deactivate the receiver: refused while it has active types
  err(post("receiver_save", { id: rc.id, name: rc.name, reg: rc.reg, active: 0 }), /ยังมีประเภทของเสียที่เปิดใช้งาน/);
  for (const t of b0.types.filter(x => x.receiver === rc.name)) {
    okData(post("type_save", { id: t.id, code: t.code, name: t.name, permit_tons: t.permit_tons, mgmt_code: t.mgmt_code, receiver_id: rc.id, active: 0 }));
  }
  // inactive types leave the receiver's quota (permit 0) but their stock and history stay
  const rc1 = boot().receivers.find(x => x.id === rc.id);
  assert.deepStrictEqual([rc1.types, rc1.active_types, rc1.permit_tons], [2, 0, 0]);
  okData(post("receiver_save", { id: rc.id, name: rc.name, reg: rc.reg, active: 0 }));
  assert.strictEqual(boot().receivers.find(x => x.id === rc.id).active, 0);
  const bar = okData(get("dashboard")).receivers.find(x => x.receiver === rc.name);
  assert.strictEqual(bar.balance_kg, 2500, "stock of a deactivated receiver is still shown");
  // an inactive receiver cannot get an active type; reactivating works
  err(post("type_save", { code: "130210", name: "x", permit_tons: 1, mgmt_code: "061", receiver_id: rc.id }), /ปิดใช้งานอยู่/);
  okData(post("receiver_save", { id: rc.id, name: rc.name, reg: rc.reg, active: 1 }));
  okData(post("type_save", { id: used.id, code: used.code, name: used.name, permit_tons: used.permit_tons, mgmt_code: used.mgmt_code, receiver_id: rc.id, active: 1 }));
  // delete works when nothing refers to it
  const fresh = okData(post("type_save", { code: "130211", name: "ลบได้", permit_tons: 1, mgmt_code: "061", receiver: "ผู้รับชั่วคราว" }));
  assert.strictEqual(okData(post("type_delete", { id: fresh.id })).id, fresh.id);
  const tmp = boot().receivers.find(x => x.name === "ผู้รับชั่วคราว");
  okData(post("receiver_delete", { id: tmp.id }));
  assert.ok(!boot().receivers.some(x => x.name === "ผู้รับชั่วคราว"));
  assert.ok(logs({ cat: "setting" }).some(l => l.action === "receiver_delete"));
});

test("settings: older sheets without a receivers sheet keep working (receivers derived from types, stored on first write)", () => {
  const ss = state.spreadsheets[state.props.WASTE_SPREADSHEET_ID];
  const idx = ss.sheets.findIndex(s => s.name === "receivers");
  const saved = ss.sheets.splice(idx, 1)[0];
  ctx.clearReadCache();
  try {
    const b = boot();
    const names = b.receivers.map(r => r.name);
    assert.ok(names.includes("บจก. ผู้รับ ทดสอบ เอ") && b.receivers.every(r => r.id === 0 || r.id > 0));
    assert.ok(b.receivers.length >= 3 && b.receivers.every(r => r.active === 1));
    const rc = b.receivers.find(r => r.name === "บจก. ผู้รับ ทดสอบ เอ");
    assert.deepStrictEqual([rc.types, rc.permit_tons], [3, 60]);
    assert.ok(okData(get("dashboard")).receivers.some(r => r.receiver === "บจก. ผู้รับ ทดสอบ เอ" && r.permit_kg === 60000));
    // first write: the sheet is created and seeded from the types, then editable by name
    const e = okData(post("receiver_save", { current_name: "บจก. ผู้รับ ทดสอบ เอ", name: "บจก. ผู้รับ ทดสอบ เอ", reg: "1000000000001", mgmt_code: "011" }));
    assert.ok(e.id > 0);
    const rows = ss.getSheetByName("receivers").data.length - 1;
    assert.ok(rows >= 3, "seeded rows: " + rows);
    assert.strictEqual(boot().receivers.find(r => r.id === e.id).mgmt_code, "011");
  } finally {
    if (!ss.getSheetByName("receivers")) ss.sheets.splice(idx, 0, saved);
    ctx.clearReadCache();
  }
});

test("settings: admin only (viewer / staff / anonymous refused, nothing written)", () => {
  const n0 = dataRows("receivers"), t0 = dataRows("waste_types");
  const body = { name: "ไม่ควรสร้าง", code: "130299", permit_tons: 1, mgmt_code: "061", receiver: "ไม่ควรสร้าง", id: 1 };
  for (const a of ["receiver_save", "receiver_delete", "type_save", "type_delete"]) {
    err(post(a, { ...body, token: TS }), /ผู้ดูแลระบบ/);
    err(post(a, { ...body, token: TV }), /ผู้ดูแลระบบ/);
    err(post(a, { ...body, token: undefined }), /ผู้ดูแลระบบ/);
  }
  assert.strictEqual(dataRows("receivers"), n0);
  assert.strictEqual(dataRows("waste_types"), t0);
});

test("two decimals everywhere: page helpers, inputs and README use 2-decimal tons", () => {
  const app = fs.readFileSync(path.join(__dirname, "../docs/assets/app.js"), "utf8");
  const readme = fs.readFileSync(path.join(__dirname, "../README.md"), "utf8");
  assert.ok(!/0\.000/.test(app), "no 3-decimal placeholder left in app.js");
  assert.ok(app.includes("fmtT = kg =>") && /minimumFractionDigits: 2, maximumFractionDigits: 2/.test(app));
  assert.ok(app.includes("permitStr") && app.includes("ทศนิยมไม่เกิน 2 ตำแหน่ง"));
  assert.ok(!/ทศนิยม 3 ตำแหน่ง/.test(readme), "README still says 3 decimals");
});

/* ---------- read cache / versions ---------- */
test("read cache: repeated reads served from CacheService; every write bumps the data version", () => {
  boot();
  const n0 = state.sheetReads.movements || 0;
  boot(); boot(); okData(get("history"));
  assert.strictEqual(state.sheetReads.movements || 0, n0, "cached reads must not read the sheet");
  const v0 = state.cache.WASTE_V_movements;
  assert.ok(v0);
  const b1 = bal(1).balance_kg;
  const ev0 = state.events.length;
  okData(mv("store", 1, 2, { container: "bigbag", containers: 1, ...at(1) }));
  const v1 = state.cache.WASTE_V_movements;
  assert.notStrictEqual(v1, v0);
  const evs = state.events.slice(ev0);
  assert.ok(evs.indexOf("flush") < evs.lastIndexOf("version:movements"), "flush before the version bump");
  const b = bal(1);
  assert.strictEqual(b.balance_kg, Math.round((b1 + 2) * 100) / 100); // fresh data after the write
  // waste_types version unchanged by a movement write
  const vt = state.cache.WASTE_V_waste_types;
  okData(mv("store", 1, 1, { container: "bigbag", containers: 1, ...at(1) }));
  assert.strictEqual(state.cache.WASTE_V_waste_types, vt);
  // a rejected write changes nothing but the audit_logs version
  const vm0 = state.cache.WASTE_V_movements;
  err(mv("ship", 4, 999));
  assert.strictEqual(state.cache.WASTE_V_movements, vm0);
});

test("read cache: clearReadCache() forces fresh reads (hand edits); a read-cache outage still answers", () => {
  boot();
  const n0 = state.sheetReads.waste_types || 0;
  sheet("waste_types").data[1][2] = "กระดาษลัง (แก้ด้วยมือ)";
  assert.notStrictEqual(boot().types[0].name, "กระดาษลัง (แก้ด้วยมือ)");
  ctx.clearReadCache();
  assert.strictEqual(boot().types[0].name, "กระดาษลัง (แก้ด้วยมือ)");
  assert.ok((state.sheetReads.waste_types || 0) > n0);
  sheet("waste_types").data[1][2] = "กระดาษลูกฟูก กระดาษลัง";
  ctx.clearReadCache();
  state.batchCacheDown = true; // the read cache / data versions fail; sessions (get) still answer
  try {
    assert.strictEqual(boot().types[0].name, "กระดาษลูกฟูก กระดาษลัง");
    okData(mv("store", 1, 1, { container: "bigbag", containers: 1, ...at(1) }));
  } finally { state.batchCacheDown = false; }
  ctx.clearReadCache();
});

test("cache values stay under 100 KB per key (chunked) with many rows", () => {
  for (let i = 0; i < 40; i++) okData(mv("store", 6, 1, { note: "ข้อความภาษาไทยยาว ๆ ".repeat(10), ...at(1) }));
  ctx.clearReadCache();
  okData(get("history"));
  okData(get("history")); // asserts inside the mock putAll: <= 100 KB, key <= 250 chars
});

test("batch reads and write responses carrying post-write reads", () => {
  const r = okData(get("batch", { reads: JSON.stringify([{ action: "bootstrap" }, { action: "dashboard" }, { action: "history", kind: "ship" }]) }));
  assert.strictEqual(r.results.length, 3);
  assert.ok(r.results[0].types && r.results[1].kpi && r.results[2].items.every(m => m.kind === "ship"));
  err(get("batch", { reads: JSON.stringify([{ action: "logs" }]) }), /Unknown action/);
  err(get("batch", { reads: "[]" }), /reads/);
  err(get("batch", { reads: JSON.stringify(new Array(7).fill("bootstrap")) }), /reads/);
  const w = okData(mv("store", 2, 5, { container: "ibc", containers: 1, ...at(1), reads: [{ action: "bootstrap" }] }));
  assert.strictEqual(w.reads.length, 1);
  assert.strictEqual(w.reads[0].balance[2].balance_kg, bal(2).balance_kg);
  err(mv("store", 2, 5, { reads: [{ action: "logs" }] }), /Unknown action/);
});

test("LockService: writes take and release the script lock; lock timeout answers politely", () => {
  assert.strictEqual(state.locks, state.unlocks);
  const n = state.locks;
  okData(mv("store", 1, 1, { container: "bigbag", containers: 1, ...at(1) }));
  assert.strictEqual(state.locks, n + 1);
  assert.strictEqual(state.unlocks, n + 1);
  state.lockFail = true;
  try { err(mv("store", 1, 1), /กำลังบันทึกรายการอื่น/); } finally { state.lockFail = false; }
  err(mv("ship", 4, 1000));
  assert.strictEqual(state.locks, state.unlocks, "released after a failed write");
});

test("keepWarm / installKeepWarmTrigger / removeKeepWarmTrigger", () => {
  const w = ctx.keepWarm();
  assert.strictEqual(w.warmed, true);
  const r = ctx.installKeepWarmTrigger();
  assert.strictEqual(r.installed, true);
  assert.strictEqual(r.everyMinutes, 10);
  ctx.installKeepWarmTrigger();
  assert.strictEqual(state.triggers.filter(t => t.handler === "keepWarm").length, 1);
  assert.ok(state.cachePutAlls.includes(900), "keepWarm stores with the longer TTL");
  assert.strictEqual(ctx.removeKeepWarmTrigger().removed, 1);
  assert.strictEqual(state.triggers.length, 0);
});

/* ---------- audit log ---------- */
test("audit log (admin session): filters by category / search / date, CSV; texts in tons", () => {
  const all = logs();
  assert.ok(all.length > 20);
  for (let i = 1; i < all.length; i++) assert.ok(all[i - 1].id > all[i].id);
  assert.ok(logs({ cat: "edit" }).every(l => l.category === "edit"));
  assert.ok(logs({ q: "MN-1" }).length >= 1);
  assert.strictEqual(logs({ from: "2001-01-01", to: "2001-01-01" }).length, 0);
  const c = okData(post("logs_export", { cat: "setting" }));
  assert.ok(/^waste_thepwatana_log_\d{8}_\d{6}\.csv$/.test(c.filename));
  assert.ok(c.csv.startsWith("ลำดับ,วันที่,เวลา,หมวด,เลขที่อ้างอิง,รายละเอียด,ผู้ทำรายการ\n"));
  assert.ok(c.csv.includes("ตั้งค่า"));
  assert.ok(logs({ cat: "waste" }).slice(0, 5).every(l => / ตัน · /.test(l.detail)));
});

/* ---------- reset data ---------- */
test("reset_data: admin session + password + name required; wipes movements and logs, keeps waste types; numbers restart", () => {
  assert.strictEqual(postRaw({ action: "reset_data", resetPassword: RPW }).code, "AUTH_REQUIRED");
  err(post("reset_data", { resetPassword: "bad" }), /รหัสไม่ถูกต้อง/);
  clearFails();
  assert.ok(dataRows("movements") > 0);
  const nTypes = dataRows("waste_types"), nRecv = dataRows("receivers");
  const r = okData(post("reset_data", { reads: [{ action: "bootstrap" }] }));
  assert.ok(r.removed.movements > 0 && r.removed.audit_logs > 0);
  assert.strictEqual(dataRows("movements"), 0);
  assert.strictEqual(dataRows("waste_types"), nTypes);
  assert.strictEqual(dataRows("receivers"), nRecv, "receivers are kept");
  const l = logs();
  assert.strictEqual(l.length, 1);
  assert.strictEqual(l[0].action, "reset_data");
  assert.ok(Object.values(r.reads[0].balance).every(b => b.balance_kg === 0));
  const s = okData(mv("store", 1, 5, { container: "bigbag", containers: 1, ...at(1) }));
  assert.strictEqual(s.doc_no, "ST-" + YMD + "-0001");
});

test("sheet storage: numbers / codes / dates keep their formats (no Date conversion, no formulas)", () => {
  const types = sheet("waste_types").data;
  assert.strictEqual(types[5][1], "080409");
  assert.strictEqual(typeof types[7][3], "number"); // 0.5 t
  const mvRow = sheet("movements").data[1];
  assert.ok(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:00$/.test(mvRow[11]));
  assert.strictEqual(typeof mvRow[4], "number");
});

/* ---------- regressions (bug hunt 2026-10) ---------- */
test("no password: balance last_at never reveals the time of an in-house row (the total is public)", () => {
  const s = okData(mv("store", 2, 40, { container: "ibc", containers: 2, ...at(10) }));
  okData(mv("treat", 2, 5, { container: "ibc", containers: 1, ...at(1) }));
  const atStore = okData(get("history", { type: "2" })).items.find(m => m.doc_no === s.doc_no).moved_at;
  const hidden = boot().balance[2];
  assert.strictEqual(hidden.last_at, atStore);
  assert.ok(!("last_vis_at" in hidden));
  const shown = tread(TA, "bootstrap").balance[2];
  assert.strictEqual(shown.last_at, at(1).date + " 09:00:00");
  assert.ok(!("last_vis_at" in shown));
  assert.strictEqual(hidden.balance_kg, shown.balance_kg); // the balance itself stays true
  assert.ok(!("treated_kg" in hidden) && typeof shown.treated_kg === "number"); // the total only with a token
  assert.strictEqual(okData(get("batch", { reads: JSON.stringify(["bootstrap"]) })).results[0].balance[2].last_at, atStore);
});

test("movement_update: an empty date / time keeps the stored moment (not silently moved to now)", () => {
  const s = okData(mv("store", 2, 3, { container: "ibc", containers: 1, ...at(8, 7, 15) }));
  const r = okData(post("movement_update", { id: s.id, doc_no: s.doc_no, weight_t: "0.004", date: "", time: "" }));
  assert.strictEqual(r.changed, true);
  const row = okData(get("history", { type: "2" })).items.find(m => m.id === s.id);
  assert.strictEqual(row.moved_at, at(8).date + " 07:15:00");
  assert.strictEqual(row.weight_kg, 4);
  // only the time cleared: the date given is used, the stored time kept
  okData(post("movement_update", { id: s.id, doc_no: s.doc_no, date: at(9).date, time: "" }));
  assert.strictEqual(okData(get("history", { type: "2" })).items.find(m => m.id === s.id).moved_at, at(9).date + " 07:15:00");
});

test("CSV: a note with backslash + quote stays in one Excel cell (RFC 4180 quote doubling)", () => {
  const s = okData(mv("store", 2, 1, { container: "ibc", containers: 1, note: 'ท่อ 3\\" ยาว "พิเศษ"', ...at(1) }));
  const csv = okData(get("export", { type: "2" })).csv;
  const line = csv.split("\n").find(x => x.startsWith(s.doc_no));
  // minimal RFC 4180 reader (what Excel does): "" inside quotes is one quote, backslash is plain text
  const cells = [];
  for (let i = 0, cur = "", q = false; i <= line.length; i++) {
    const ch = line[i];
    if (q) { if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
    else if (ch === '"') q = true;
    else if (ch === "," || ch === undefined) { cells.push(cur); cur = ""; }
    else cur += ch;
  }
  assert.strictEqual(cells.length, 16, line);
  assert.strictEqual(cells[14], 'ท่อ 3\\" ยาว "พิเศษ"');
});

/* ---------- frontend ---------- */
const readDoc = f => fs.readFileSync(path.join(ROOT, f), "utf8");
test("docs JS files parse; config shape; POST text/plain only; token in the body; actor sent; UTF-8 BOM CSV", () => {
  for (const f of ["docs/config.js", "docs/assets/app.js", "test/dev-server.js"]) new Function(readDoc(f).replace(/^#!.*\n/, ""));
  const cfg = {};
  new Function("window", readDoc("docs/config.js"))(cfg);
  assert.deepStrictEqual(Object.keys(cfg.WASTE_CONFIG), ["apiUrl"]);
  assert.ok(cfg.WASTE_CONFIG.apiUrl === "" || /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(cfg.WASTE_CONFIG.apiUrl));
  const app = readDoc("docs/assets/app.js");
  assert.ok(app.includes("window.WASTE_CONFIG"));
  assert.ok(app.includes("'Content-Type': 'text/plain;charset=UTF-8'"));
  assert.ok(!/method: 'GET'|searchParams\.set/.test(app), "the page only POSTs (a token never goes into a URL)");
  assert.ok(/actor: S\.user/.test(app) && app.includes("store('get', 'wthep_user')"));
  assert.ok(app.includes("'\\uFEFF' + r.csv"));
  for (const a of ["store", "ship", "opening", "treat", "movement_update", "movement_delete", "type_save", "reset_data"]) assert.ok(new RegExp("NAME_REQUIRED = \\[[^\\]]*'" + a + "'").test(app), a);
  for (const a of ["'movement_update'", "'movement_delete'", "'type_save'", "'logs'", "'logs_export'", "'batch'", "action: 'reset_data', resetPassword", "action: 'login', password", "action: 'logout', token"]) assert.ok(app.includes(a), a);
  for (const gone of ["'treat_unlock'", "'treat_lock'", "verifyResetPassword", "recordPassword", "treatToken", "showLogin", "$('#login')", "HT_KEY", "#treatLock"]) assert.ok(!app.includes(gone), gone);
  // open site: no login gate; the first thing the page does is draw the current page
  assert.ok(/\npaintRole\(\);\r?\nroute\(\);\s*$/.test(app), "no login screen before the dashboard");
  assert.ok(!/if \(!S\.sess\) \{ showLogin/.test(app));
  // session: token in sessionStorage + POST bodies only; passwords never stored
  assert.ok(app.includes("body.token = S.sess.t; // in the POST body only, never in a URL") && app.includes("ss('set', SESS_KEY, JSON.stringify(S.sess))"));
  assert.ok(!/localStorage\.setItem\([^)]*(SESS|sess|token)/i.test(app));
  assert.ok(!/(local|session)Storage\.setItem\([^)]*([Pp]ass|[Rr]eset)/.test(app));
  assert.ok(!/(store|ss)\('set', *'[^']*([Pp]ass|[Rr]eset)/.test(app));
  assert.ok(app.includes("if (e.code === 'AUTH_REQUIRED' && S.sess) sessionEnded("));
  // in-house rows only on the unlocked ข้อมูลภายใน page; every other page asks the server to hide them
  assert.ok(app.includes("const treatOn = () => can('viewer') && currentPage() === 'treat';"));
  assert.ok(app.includes("if (!treatOn()) body.hideTreat = true;"));
  // read cache: this tab only, never for the ข้อมูลภายใน page, never the log
  assert.ok(app.includes("if (treatOn()) return null;") && app.includes("if (!d || treatOn()) return;"));
  assert.ok(!/localStorage\.setItem\(rcKey/.test(app) && app.includes("sessionStorage.setItem(rcKey(r), s)"));
  assert.ok(!/RC_PREFIX[^\n]*logs/.test(app));
  // pages that need a password: ข้อมูลภายใน (view; the form staff), Log / settings (admin)
  assert.ok(app.includes("const PAGE_ROLE = { treat: 'viewer', logs: 'admin', settings: 'admin' };"));
  assert.ok(app.includes("const pageNeed = (p = currentPage()) => (p === 'treat' && treatTab() === 'new' ? 'staff' : PAGE_ROLE[p] || '');"));
  assert.ok(app.includes("const NAME_PAGES = ['store', 'ship'];"), "store / ship open to everyone (name only)");
  assert.ok(app.includes("lockedPage(p, need);") && app.includes("await askUnlock(need)"));
  assert.ok(app.includes("treat: ([h]) => treatPage(h)") && app.includes("historyPage(h, 'treat')") && app.includes("txForm('treat')"));
  assert.ok(app.includes("$('#tbLock').onclick = lockNow;") && app.includes("$('#logoutBtn').onclick = lockNow;"));
  // the server is the authority: a session it no longer knows ends here too
  assert.ok(app.includes("if (S.sess && j.role === null) {") && app.includes("if (!applyBoot(res[0])) return;"));
  // admin tools ask for the admin password on click (as before the login change)
  assert.ok(app.includes("needRole('admin', 'การแก้ไขรายการต้องใช้รหัสผ่านผู้ดูแลระบบ')") && app.includes("needRole('admin', 'การลบรายการต้องใช้รหัสผ่านผู้ดูแลระบบ')"));
  assert.ok(app.includes("const acts = m => `<div class=\"rec-acts\">"), "edit / delete buttons for everyone (password on click)");
  // a lower-level password never replaces a higher session; the extra token is logged out
  assert.ok(app.includes("if (!S.sess || ROLE_LV[r.role] >= ROLE_LV[S.sess.role]) {"));
  // dashboard: the summary card only when the server sent the totals (unlocked tab); unlock / lock redraws it
  assert.ok(app.includes("sessionEnded(msg, noRoute)") && app.includes("if (TOTALS_PAGES.includes(currentPage())) route();") && app.includes("else if (!noRoute && TOTALS_PAGES.includes(currentPage())) route();"));
  assert.ok(app.includes("${k.treated_kg !== undefined ? treatSummary(k) : ''}"));
  assert.ok(!/treatSummary[^\n]*(can\(|treatOn\(\))/.test(app));
  // regression: the level painter must never match <body> (a fresh login hid the whole page)
  assert.ok(app.includes("$$('.nav [data-lock]').forEach(") && !/body\.dataset\.role|body\.setAttribute\('data-role'/.test(app));
  // tons on the page: input / quick-add / exact conversion; kg only as a hint
  assert.ok(app.includes("/^\\d+(\\.\\d{1,5})?$/") && app.includes("weight_t: tonsStr(parseT(F.weight))"));
  assert.ok(app.includes("const QUICK_T = [0.1, 0.5, 1, 5, 10];"));
  assert.ok(!/weight_kg: /.test(app), "the page never sends kg");
  assert.ok(!/fmtKg\(|parseKg/.test(app));
  assert.ok(app.includes("if (wt !== w0) {"), "edit sends weight_t only when changed");
  const fx = new Function(app.match(/const fmt = [^\n]+/)[0] + "\n" + app.match(/const fmtT = [^\n]+/)[0] + "\n" +
    app.slice(app.indexOf("const tonsStr = "), app.indexOf("/* ---------- container illustrations")) + "\nreturn { fmtT, tonsStr, parseT };")();
  assert.strictEqual(fx.tonsStr(125), "0.125");
  assert.strictEqual(fx.tonsStr(1234.56), "1.23456");
  assert.strictEqual(fx.tonsStr(1000), "1.00");
  assert.strictEqual(fx.tonsStr(100), "0.10");
  assert.strictEqual(fx.tonsStr(0.01), "0.00001");
  assert.strictEqual(fx.parseT("0.125"), 125);
  assert.strictEqual(fx.parseT("1,234.5"), 1234500);
  assert.strictEqual(fx.parseT("0.3"), 300);
  assert.ok(isNaN(fx.parseT("0.000001")) && isNaN(fx.parseT("-1")) && isNaN(fx.parseT("1e3")));
  for (const k of [0.01, 1, 99.99, 125, 1234.56, 54000]) assert.strictEqual(fx.parseT(fx.tonsStr(k)), k, "round trip " + k);
  assert.ok(/^1,234\.56[0-9]*$/.test(fx.fmtT(1234560).replace(/[^\d.,]/g, "")) || fx.fmtT(1234560).includes("1,234.560"));
  // client-side checks mirror the server (ship <= balance)
  assert.ok(app.includes("เกินยอดคงเหลือ"));
  // regressions: no global function named like a window property (history() replaced window.history)
  assert.ok(!/^(async )?function (history|location|name|status|close|open|print|stop|focus|blur|event)\(/m.test(app));
  // a save in flight: redrawing the summary (quick-add, typing, fresh data) never re-enables the button
  assert.ok(app.includes("if (!go || go.disabled || busy) return;") && app.includes("if (busy) { $('#fGo').disabled = true;"));
  // Enter pressed again while a password is being checked: one request only
  assert.ok(/onsubmit = async e => \{\r?\n\s+e\.preventDefault\(\);\r?\n\s+if \(go\.disabled\) return;/.test(app));
});

test("names: the menu / page is “ข้อมูลภายใน”, the public dashboard totals say “บำบัดเอง” (internal keys unchanged)", () => {
  const app = readDoc("docs/assets/app.js"), html = readDoc("docs/index.html"), readme = readDoc("README.md");
  const src = readDoc("apps-script/Code.gs");
  const code = t => t.split(/\r?\n/).filter(l => !/^\s*(\*|\/\/|\/\*)/.test(l)).join("\n").replace(/\/\/[^\n"']*$/gm, "");
  // only data values may still contain the word (seed waste name, preset system names)
  const strip = t => t.replace(/ตะกอนจากระบบบำบัดน้ำเสีย|ระบบบำบัดน้ำเสียภายในโรงงาน|บำบัดน้ำเสีย/g, "");
  for (const [f, t] of [["index.html", html], ["Code.gs (code)", code(src)]]) {
    assert.ok(!/บำบัด/.test(strip(t)), f + ": " + (strip(t).match(/.{0,30}บำบัด.{0,30}/) || [""])[0]);
  }
  for (const [f, t] of [["app.js", app], ["index.html", html], ["Code.gs", src], ["README.md", readme]]) assert.ok(!/บำบัดเองภายใน/.test(t), f + ": old menu name");
  // app.js: “บำบัดเอง” only as the label of the public totals
  assert.ok(strip(app).match(/บำบัด.{0,4}/g).every(x => x.startsWith("บำบัดเอง")), "only the word บำบัดเอง");
  assert.ok(app.includes("['บำบัดเอง', k.treated_kg, 'fa-flask-vial'") && app.includes("label: 'บำบัดเอง (ตัน)'"));
  assert.ok(app.includes("treat:   { name: 'ข้อมูลภายใน', short: 'ข้อมูลภายใน'") && app.includes("treat: ['ข้อมูลภายใน', "));
  assert.ok(html.includes('<span>ข้อมูลภายใน</span>'));
  assert.ok(readme.includes("ข้อมูลภายใน") && readme.includes("บำบัดเอง"));
  assert.strictEqual(ctx.KIND_NAME.treat, "ข้อมูลภายใน");
  assert.strictEqual(ctx.PREFIX.treat, "TR"); // doc numbers unchanged
  assert.ok(ctx.KINDS.includes("treat") && ctx.ACTION_ROLE.treat === "staff"); // API action / kind unchanged
});

test("storage isolation: every browser storage key of the page is namespaced wthep_ (shares nothing with the other Balance Waste site on the same origin); no original-company data in docs / README", () => {
  const app = readDoc("docs/assets/app.js"), html = readDoc("docs/index.html"), cfg = readDoc("docs/config.js"), readme = readDoc("README.md");
  const keys = [...app.matchAll(/(?:store|ss)\('(?:get|set|del)',\s*(?:'([^']+)'|(\w+))/g)].map(m => m[1] || m[2]);
  assert.ok(keys.length >= 5, "storage calls found");
  const consts = {};
  for (const m of app.matchAll(/const (\w+) = '([^']+)';/g)) consts[m[1]] = m[2];
  for (const k of keys) { const v = consts[k] || k; assert.ok(/^wthep_/.test(v), "storage key not namespaced: " + v); }
  assert.ok(/RC_PREFIX = 'wthep_rc:'/.test(app));
  assert.ok(!/['"`]waste(User|Sess|Client|RC)/.test(app), "old un-namespaced key left");
  // every direct use of a storage API goes through the namespaced helpers / prefix
  assert.ok(!/\b(indexedDB|caches|document\.cookie)\b/.test(app), "no other storage APIs");
  for (const m of app.matchAll(/(?:localStorage|sessionStorage)\.(?:get|set|remove)Item\(([^,)]+)/g)) assert.ok(/rcKey|k\b|RC_PREFIX/.test(m[1]), "raw storage key: " + m[1]);
  assert.ok(html.includes("บริษัท เทพวัฒนา จำกัด"));
  for (const [n, txt] of [["app.js", app], ["index.html", html], ["config.js", cfg], ["README.md", readme]]) {
    for (const bad of ["แน็ป", "82260007625693", "2569-8409", "AKfycb", "balance-waste-nap", "Waste Nap"]) assert.ok(!txt.includes(bad), n + " must not contain " + bad);
  }
  assert.strictEqual(cfg.replace(/^\/\/.*$/gm, "").match(/apiUrl:\s*"([^"]*)"/)[1], "", "config.js apiUrl stays empty in the repo");
  assert.ok(readme.includes("https://watanathep8-dotcom.github.io/balance-waste-thepwatana/") && readme.includes("Balance Waste Thepwatana"));
});

test("empty-state: the page has a friendly no-types notice (dashboard / store / ship / balance) linking to the settings page; nothing indexes types[0]", () => {
  const app = readDoc("docs/assets/app.js");
  assert.ok(app.includes("ยังไม่มีประเภทของเสีย — ไปที่ <b>ตั้งค่าประเภทของเสีย (ผู้ดูแลระบบ)</b> เพื่อเพิ่มประเภทและผู้รับดำเนินการ"));
  assert.ok(/noTypesBox[\s\S]{0,600}href="#settings"/.test(app));
  assert.ok(/function txForm[\s\S]{0,900}noTypesBox/.test(app) && /function balance\(\) \{\s*if \(!S\.types\.length\)/.test(app) && /\$\{S\.types\.length \? '' : noTypesBox\(\)\}/.test(app));
  assert.ok(!/S\.types\[0\]/.test(app));
  // the optional identity chips are shown only when set
  assert.ok(/S\.company\.factoryReg \?/.test(app) && /S\.company\.permitNo \|\| S\.company\.permitYear \?/.test(app));
});

test("index.html / style.css: brand, rainbow credit (7 colours, 21 s), rotating icons, reduced motion, CDNs, cache busters", () => {
  const html = readDoc("docs/index.html"), css = readDoc("docs/assets/style.css");
  assert.ok(html.includes("<title>Balance Waste"));
  assert.strictEqual((html.match(/class="programming-credit">Programming by Duongruthai_Milk</g) || []).length, 2); // sidebar + footer
  assert.ok(css.includes(".programming-credit{font-size:11px;font-weight:700;letter-spacing:.02em;margin-top:2px;animation:rainbowText 21s ease-in-out infinite}"));
  assert.ok(css.includes("@keyframes rainbowText{0%,100%{color:#ef4444}14.28%{color:#f97316}28.56%{color:#eab308}42.84%{color:#22c55e}57.12%{color:#06b6d4}71.4%{color:#3b82f6}85.68%{color:#8b5cf6}}"));
  assert.ok(/prefers-reduced-motion:reduce[\s\S]*\.programming-credit\{animation:none/.test(css));
  assert.ok(css.includes("conic-gradient") && css.includes("@keyframes spin{to{transform:rotate(360deg)}}"));
  assert.ok(css.includes(".ccard.sel{transform:rotateY(180deg)}") && css.includes(".orbit .dot"));
  assert.ok(css.includes(".needle-gauge .needle"));
  // background animations: transform / opacity only (no animated blur)
  const kf = name => (css.match(new RegExp("@keyframes " + name + "\\{[^@]*?\\}\\}")) || [""])[0];
  for (const k of ["float", "drift", "spin", "pulse"]) { assert.ok(kf(k), k); assert.ok(!/filter|blur|width|height|top|left/.test(kf(k)), k + " animates layout/blur"); }
  assert.ok(html.includes("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css"));
  assert.ok(html.includes("https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"));
  assert.ok(html.includes("family=Kanit") && html.includes("Prompt"));
  assert.ok(html.includes('id="userChip"') && html.includes('data-page="logs"') && html.includes('data-page="settings"'));
  // open site: no login screen; every menu listed; ข้อมูลภายใน / Log / settings carry a lock until unlocked
  assert.ok(!html.includes('id="login"') && !/class="login"/.test(html) && !css.includes(".login{"));
  assert.ok(!/data-role=/.test(html), "no menu hidden by level");
  assert.ok(/data-page="treat" class="treat-nav locked" id="treatNav" data-lock="viewer">[^\n]*<span>ข้อมูลภายใน<\/span><b class="tn-lock"/.test(html));
  for (const pg of ["logs", "settings"]) assert.ok(new RegExp(`data-page="${pg}" class="locked" data-lock="admin">[^\\n]*tn-lock`).test(html), pg);
  for (const pg of ["store", "ship"]) assert.ok(html.includes(`<a href="#${pg}" data-page="${pg}"><i`), pg);
  assert.ok(!html.includes("treatLock"), "the old show / hide toggle is gone");
  assert.ok(html.includes('id="logoutBtn" type="button" hidden') && html.includes('id="roleTag" hidden') && html.includes('<div class="user-box" id="userBox">'));
  assert.ok(css.includes(".nav a.locked .tn-lock{display:inline-block}") && css.includes(".unlock-card{") && css.includes(".locked-page{") && css.includes(".treat-bar{") && css.includes(".role-tag.r-admin"));
  assert.ok(css.includes("--treat:#7C3AED") && css.includes("@keyframes brew") && css.includes(".btn-treat"));
  // cache busters bumped with this change
  assert.ok(html.includes("assets/app.js?v=1") && html.includes("assets/style.css?v=5"));
  assert.ok(html.includes("บริษัท เทพวัฒนา จำกัด") && !html.includes("82260007625693") && !html.includes("2569-8409"));
  assert.ok(!/<script[^>]+src="(?!https:\/\/cdnjs\.cloudflare\.com\/|config\.js|assets\/app\.js)/.test(html));
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
