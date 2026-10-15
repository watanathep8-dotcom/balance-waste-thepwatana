// Local-only preview of docs/ against the in-memory GAS mock (no Google, no network).
//   node test/dev-server.js [port] [--demo | --seed]
// Without --demo the database is EMPTY (like a fresh setupSystem): add types in the settings page (admin: dev-reset).
// --demo first creates 13 synthetic sample types (fake receivers), then fills the in-memory mock with generated SAMPLE movements (for screenshots
// only; nothing is stored anywhere and the data is not real).
"use strict";
const http = require("http");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const ROOT = path.join(__dirname, "..");
const mockSrc = fs.readFileSync(path.join(ROOT, "test/gas-mock.js"), "utf8");
const head = mockSrc.slice(0, mockSrc.indexOf("/* ===================== test runner"));
const sandbox = { require, console, __dirname: path.join(ROOT, "test"), module: {}, exports: {} };
vm.runInNewContext(head + "\n;this.createGas = createGas; this.loadCode = loadCode; this.seedFixtureTypes = seedFixtureTypes;", sandbox);
const { gas, state } = sandbox.createGas();
const ctx = sandbox.loadCode(gas);
// dev-only passwords of the in-memory mock (never used anywhere else)
Object.assign(state.props, { WASTE_VIEW_PASSWORD: "dev-view", WASTE_STAFF_PASSWORD: "dev-staff", WASTE_RESET_PASSWORD: "dev-reset" });
ctx.setupSystem();

const args = process.argv.slice(2);
const port = Number(args.find(a => /^\d+$/.test(a))) || 8777;
if (args.includes("--seed")) ctx.seedThepwatanaPermit(); // the real permit data (40 types / 8 receivers)
if (args.includes("--demo")) seedDemo();

function seedDemo() {
  let seed = 7;
  const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
  const call = body => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } }).getContent());
  const token = call({ action: "login", password: state.props.WASTE_RESET_PASSWORD }).data.token;
  sandbox.seedFixtureTypes(b => call(Object.assign({ actor: "ผู้ทดสอบ (ตัวอย่าง)", token }, b)));
  // weights: kg here, sent as tons (weight_t) like the page does
  const T = kg => (Math.round(Number(kg) * 100) / 100000).toFixed(5);
  const post = body => {
    const b = Object.assign({ actor: "ผู้ทดสอบ (ตัวอย่าง)", token }, body);
    if (b.weight_kg !== undefined) { b.weight_t = T(b.weight_kg); delete b.weight_kg; }
    return call(b);
  };
  const read = action => call({ action, token }).data;
  const p2 = n => String(n).padStart(2, "0");
  const now = new Date(Date.now() + 7 * 3600e3);
  const day = (back, h) => { const d = new Date(now.getTime() - back * 86400e3); return { date: `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}`, time: `${p2(h)}:${p2(Math.floor(rnd() * 59))}` }; };
  // [type id, typical kg per store, container]
  const plan = [[1, 180, "bigbag"], [2, 120, "bigbag"], [3, 90, "bigbag"], [4, 12, "drum200"], [5, 160, "drum200"], [6, 260, "drum200"],
    [7, 6, "drum200"], [8, 10, "bigbag"], [9, 140, "bigbag"], [10, 150, "bigbag"], [11, 420, "ibc"], [13, 900, "ibc"]];
  post({ action: "opening", type_id: 13, weight_kg: "54000", container: "ibc", containers: "54", ...day(330, 8), note: "ตัวอย่าง" });
  post({ action: "opening", type_id: 11, weight_kg: "25000", container: "ibc", containers: "25", ...day(330, 8), note: "ตัวอย่าง" });
  for (let back = 320; back >= 0; back -= 9) {
    plan.forEach(([id, kg, c]) => {
      if (rnd() < 0.45) {
        const w = Math.round(kg * (0.6 + rnd() * 0.8) * 100) / 100;
        const n = c === "ibc" ? Math.max(1, Math.round(w / 1000)) : c === "drum200" ? Math.max(1, Math.round(w / 150)) : Math.max(1, Math.round(w / 300));
        post({ action: "store", type_id: id, weight_kg: String(w), container: c, containers: String(n), location: "โรงเก็บของเสีย", ...day(back, 9 + Math.floor(rnd() * 6)) });
      }
    });
    if (back % 27 < 9) {
      const boot = read("bootstrap");
      plan.forEach(([id, , c]) => {
        const b = boot.balance[id];
        const n = b.containers[c];
        if (b.balance_kg > 0 && n > 0 && rnd() < (id === 6 || id === 4 ? 0.9 : 0.6)) {
          const take = Math.max(1, Math.floor(n * (0.6 + rnd() * 0.4)));
          const w = Math.round(b.balance_kg * take / n * 100) / 100;
          post({ action: "ship", type_id: id, weight_kg: String(w), container: c, containers: String(take), manifest_no: "MN-" + (1000 + back), vehicle: "70-" + (1000 + Math.floor(rnd() * 8999)), ...day(back, 15) });
        }
      });
    }
    if (back % 27 >= 18 && rnd() < 0.7) {
      // in-house data (ข้อมูลภายใน): part of the machine wash water goes to the plant's own system
      const b = read("bootstrap").balance[13];
      const take = Math.min(b.containers.ibc, 1 + Math.floor(rnd() * 2));
      if (take > 0) post({ action: "treat", type_id: 13, weight_kg: String(Math.min(take * 1000, b.balance_kg)), container: "ibc", containers: String(take), method: "ระบบบำบัดน้ำเสียภายในโรงงาน", ...day(back, 11) });
    }
  }
}

const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml" };
http.createServer((req, res) => {
  const url = new URL(req.url, "http://localhost");
  if (url.pathname === "/api") {
    let body = "";
    req.on("data", c => body += c);
    req.on("end", () => {
      const out = req.method === "POST" ? ctx.doPost({ postData: { contents: body } }) : ctx.doGet({ parameter: Object.fromEntries(url.searchParams) });
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(out.getContent());
    });
    return;
  }
  if (url.pathname === "/config.js") {
    res.writeHead(200, { "Content-Type": TYPES[".js"] });
    return res.end('window.WASTE_CONFIG = { apiUrl: "/api" };');
  }
  const docs = path.join(ROOT, "docs");
  const file = path.normalize(path.join(docs, url.pathname === "/" ? "index.html" : decodeURIComponent(url.pathname)));
  if (!file.startsWith(docs)) { res.writeHead(403); return res.end("forbidden"); }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); return res.end("not found"); }
    res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream" });
    res.end(data);
  });
}).listen(port, "127.0.0.1", () => console.log("ready http://127.0.0.1:" + port + "/"));
