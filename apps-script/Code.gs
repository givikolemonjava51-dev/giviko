/* ====== Giviko — სერვერი (Google Apps Script) ======
   მონაცემები ინახება Google Sheets ცხრილში, რომელიც პირველ გაშვებაზე თავად შეიქმნება.
   ახალ მოთხოვნაზე და ჯგუფში ჯავშანზე შეტყობინება მოდის მეილზე და WhatsApp-ში.

   ⚠ შეავსეთ ქვემოთ მხოლოდ Apps Script-ში ჩასმულ ასლში, GitHub-ზე ნუ ატვირთავთ. */
const ADMIN_PASSWORD = "";     // ტრენერის პანელის პაროლი (მინ. 6 სიმბოლო)
const NOTIFY_EMAIL = "";       // მეილი, სადაც მოთხოვნები მოვა
const WHATSAPP_PHONE = "";     // WhatsApp ნომერი 995-ით, მაგ. "995555123456"
const CALLMEBOT_APIKEY = "";   // CallMeBot-ის გასაღები

/* ---------- საიტის პარამეტრები (ემთხვევა index.html-ს) ---------- */
const TZ = "Asia/Tbilisi";
const CAPACITY = 4;
const PASS = 8;
const PRICES = { 1: 140, 2: 180, 3: 210, 4: 240 };
const EXPERIENCE = ["სრულიად დამწყები", "0-1 წელი", "1+ წელი"];
const SCHEDULE = [
  { key: "mw", label: "ორშაბათი-ოთხშაბათი", days: [1, 3], hours: ["19:00", "20:00", "21:00", "22:00"] },
  { key: "tt", label: "სამშაბათი-ხუთშაბათი", days: [2, 4], hours: ["19:00", "20:00", "21:00", "22:00"] },
  { key: "ss", label: "შაბათი-კვირა", days: [6, 0], hours: ["09:00", "10:00", "18:00"] },
];
const pad_ = n => String(n).padStart(2, "0");
const GROUPS = {};
SCHEDULE.forEach(day => day.hours.forEach(h => {
  const id = day.key + "-" + h.slice(0, 2), end = pad_((+h.slice(0, 2) + 1) % 24) + ":00";
  GROUPS[id] = { id, days: day.days, time: h, label: `${day.label} ${h}-${end}` };
}));

const SESSION_DAYS = 180, ADMIN_HOURS = 12, MAX_TRIES = 5;

/* ---------- ცხრილები ---------- */
const TABLES = {
  Students: ["phone", "name", "pinHash", "created"],
  Groups: ["id", "group", "name", "phone", "active", "pending", "purchased", "created"],
  Requests: ["id", "people", "price", "date", "time", "name", "phone", "experience", "status", "created"],
  Attendance: ["date", "kind", "label", "member", "phone", "status"],
  Reserved: ["group", "count"],
};

function book_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SHEET_ID"), ss = null, blank = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create("Giviko — მონაცემები");
    blank = ss.getSheets()[0];
    props.setProperty("SHEET_ID", ss.getId());
  }
  Object.keys(TABLES).forEach(name => {
    if (ss.getSheetByName(name)) return;
    const sh = ss.insertSheet(name);
    sh.getRange("A:Z").setNumberFormat("@"); // ნომრები და თარიღები ტექსტად
    sh.getRange(1, 1, 1, TABLES[name].length).setValues([TABLES[name]]).setFontWeight("bold");
    sh.setFrozenRows(1);
    if (name === "Reserved") sh.getRange(2, 1, Object.keys(GROUPS).length, 2).setValues(Object.keys(GROUPS).map(g => [g, "0"]));
  });
  if (blank) ss.deleteSheet(blank);
  return ss;
}

function Db_() {
  const ss = book_(), cache = {};
  const load = name => {
    if (cache[name]) return cache[name];
    const sh = ss.getSheetByName(name), cols = TABLES[name];
    const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getDisplayValues() : [];
    return (cache[name] = vals.map((r, i) => {
      const o = { _row: i + 2 };
      cols.forEach((c, j) => (o[c] = r[j]));
      return o;
    }));
  };
  const row = (name, o) => TABLES[name].map(c => (o[c] === undefined || o[c] === null ? "" : String(o[c])));
  return {
    all: load,
    add(name, o) {
      const sh = ss.getSheetByName(name);
      sh.appendRow(row(name, o));
      o._row = sh.getLastRow();
      load(name).push(o);
      return o;
    },
    save(name, o) { ss.getSheetByName(name).getRange(o._row, 1, 1, TABLES[name].length).setValues([row(name, o)]); },
    remove(name, o) {
      ss.getSheetByName(name).deleteRow(o._row);
      const list = load(name), i = list.indexOf(o);
      list.splice(i, 1);
      list.forEach(x => { if (x._row > o._row) x._row--; });
    },
    nextId(prefix) {
      const props = PropertiesService.getScriptProperties();
      const n = Number(props.getProperty("SEQ") || 100) + 1;
      props.setProperty("SEQ", String(n));
      return prefix + n;
    },
  };
}

/* ---------- დამხმარე ---------- */
const today_ = () => Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
const key_ = p => { const x = String(p || "").replace(/\D/g, ""); return x.length > 9 ? x.slice(-9) : x; };
const bool_ = v => v === true || v === "TRUE" || v === "true" || v === "1";
const fail_ = code => { throw { code }; };
const weekday_ = s => { const [y, m, d] = s.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); };

function secret_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty("TOKEN_SECRET");
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty("TOKEN_SECRET", s); }
  return s;
}
const b64_ = bytes => Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, "");
const sign_ = data => b64_(Utilities.computeHmacSha256Signature(data, secret_()));
const hashPin_ = (phone, pin) => b64_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret_() + "|" + phone + "|" + pin));

function makeToken_(kind, subject, hours) {
  const data = kind + "." + subject + "." + (Date.now() + hours * 3600e3);
  return data + "." + sign_(data);
}
function readToken_(token, kind) {
  const p = String(token || "").split(".");
  if (p.length !== 4 || p[0] !== kind) return null;
  const data = p.slice(0, 3).join(".");
  if (sign_(data) !== p[3] || Number(p[2]) < Date.now()) return null;
  return p[1];
}

function tries_(id) { return Number(CacheService.getScriptCache().get("lk:" + id) || 0); }
function checkLocked_(id) { if (tries_(id) >= MAX_TRIES) fail_("locked"); }
function badTry_(id) { CacheService.getScriptCache().put("lk:" + id, String(tries_(id) + 1), 900); }
function goodTry_(id) { CacheService.getScriptCache().remove("lk:" + id); }

/* ---------- HTTP ---------- */
function doGet() {
  return json_(handle_({ action: "counts" }));
}
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "bad_request" }); }
  return json_(handle_(d));
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

const WRITES = ["group", "individual", "register", "admin_mark", "admin_update_request", "admin_renew", "admin_confirm_payment", "admin_cancel_member"];

function handle_(d) {
  const lock = WRITES.includes(d.action) ? LockService.getScriptLock() : null;
  try {
    if (lock) lock.waitLock(20000);
    const res = route_(d || {}, Db_());
    return res;
  } catch (e) {
    if (e && e.code) return { ok: false, error: e.code };
    console.error(e && e.stack || e);
    return { ok: false, error: "server" };
  } finally {
    if (lock) { SpreadsheetApp.flush(); lock.releaseLock(); }
  }
}

function route_(d, db) {
  const students = () => db.all("Students"), groups = () => db.all("Groups");
  const inds = () => db.all("Requests"), att = () => db.all("Attendance");
  const reserved = () => { const r = {}; db.all("Reserved").forEach(x => (r[x.group] = Number(x.count) || 0)); return r; };
  const used = id => att().filter(a => a.member === id).length;
  const active = g => bool_(g.active);
  const cnt = () => {
    const c = {}, res = reserved();
    Object.keys(GROUPS).forEach(k => (c[k] = (res[k] || 0) + groups().filter(g => g.group === k && active(g)).length));
    return c;
  };
  const student = phone => students().find(x => x.phone === phone);
  const sToken = phone => makeToken_("s", phone, SESSION_DAYS * 24);
  const who = () => {
    if (d.token) {
      const phone = readToken_(d.token, "s"), s = phone && student(phone);
      if (!s) fail_("auth");
      return s;
    }
    const phone = key_(d.phone), name = String(d.name || "").trim().slice(0, 60);
    if (name.length < 2 || phone.length !== 9) fail_("invalid");
    if (!/^\d{4}$/.test(d.pin || "")) fail_("pin_format");
    let s = student(phone);
    if (s) {
      checkLocked_(phone);
      if (s.pinHash !== hashPin_(phone, d.pin)) { badTry_(phone); fail_("pin_mismatch"); }
      goodTry_(phone);
    } else {
      s = db.add("Students", { phone, name, pinHash: hashPin_(phone, d.pin), created: today_() });
    }
    return s;
  };
  const admin = () => { if (readToken_(d.token, "a") !== "admin") fail_("auth"); };
  const prof = phone => {
    const s = student(phone); if (!s) fail_("auth");
    const mine = att().filter(a => a.phone === phone), month = today_().slice(0, 7);
    return {
      name: s.name, phone,
      memberships: groups().filter(g => g.phone === phone && active(g) && GROUPS[g.group]).map(g => {
        const G = GROUPS[g.group], u = used(g.id), p = Number(g.purchased) || 0;
        return { id: g.id, group: g.group, label: G.label, days: G.days, time: G.time, since: g.created, pending: bool_(g.pending), purchased: p, used: u, remaining: Math.max(0, p - u) };
      }),
      individual: inds().filter(r => r.phone === phone && r.status !== "გაუქმებული").map(pubReq_),
      history: mine.slice().sort((a, b) => (a.date < b.date ? 1 : -1)).map(a => ({ date: a.date, kind: a.kind, label: a.label, status: a.status })),
      stats: {
        attended: mine.filter(a => a.status === "მოვიდა").length,
        missed: mine.filter(a => a.status === "გააცდინა").length,
        month: mine.filter(a => a.status === "მოვიდა" && a.date.slice(0, 7) === month).length,
      },
    };
  };
  const day = () => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date || "")) fail_("invalid");
    const wd = weekday_(d.date), mark = id => (att().find(a => a.date === d.date && a.member === id) || {}).status || "";
    const sessions = [];
    Object.values(GROUPS).forEach(g => {
      if (!g.days.includes(wd)) return;
      sessions.push({ kind: "group", ref: g.id, time: g.time, label: g.label,
        members: groups().filter(m => m.group === g.id && active(m) && !bool_(m.pending)).map(m => ({ id: m.id, name: m.name, phone: m.phone, remaining: Math.max(0, (Number(m.purchased) || 0) - used(m.id)), mark: mark(m.id) })) });
    });
    inds().filter(r => r.date === d.date && r.status !== "გაუქმებული").forEach(r => sessions.push({
      kind: "individual", ref: r.id, time: r.time, label: `ინდივიდუალური · ${r.people} ადამიანი · ${r.price} ₾`,
      members: [{ id: r.id, name: r.name, phone: r.phone, mark: mark(r.id) }] }));
    sessions.sort((a, b) => (a.time < b.time ? -1 : 1));
    return { ok: true, date: d.date, sessions };
  };
  const reqs = () => ({ ok: true, requests: inds()
    .filter(r => r.status === "ახალი" || (r.status === "დადასტურებული" && r.date >= today_()))
    .sort((a, b) => (a.date + a.time < b.date + b.time ? -1 : 1)).map(pubReq_) });
  const mems = () => ({ ok: true, reserved: reserved(), members: groups().filter(g => active(g) && GROUPS[g.group]).map(g => {
    const u = used(g.id), p = Number(g.purchased) || 0;
    return { id: g.id, group: g.group, label: GROUPS[g.group].label, name: g.name, phone: g.phone, pending: bool_(g.pending), purchased: p, used: u, remaining: Math.max(0, p - u) };
  }) });
  const member = () => { const g = groups().find(x => x.id === d.member); if (!g) fail_("invalid"); return g; };

  switch (d.action) {
    case "counts": return { ok: true, counts: cnt() };

    case "group": {
      const s = who(), g = GROUPS[d.group]; if (!g) fail_("invalid");
      const inG = groups().filter(x => x.group === g.id && active(x));
      if (inG.some(x => x.phone === s.phone)) fail_("already");
      if (inG.length + (reserved()[g.id] || 0) >= CAPACITY) fail_("full");
      db.add("Groups", { id: db.nextId("G"), group: g.id, name: s.name, phone: s.phone, active: "TRUE", pending: "TRUE", purchased: PASS, created: today_() });
      notify_("group", { name: s.name, phone: s.phone, group: g.label });
      return { ok: true, counts: cnt(), token: sToken(s.phone) };
    }

    case "individual": {
      if (!d.token) fail_("login_required");
      if (!EXPERIENCE.includes(d.experience)) fail_("experience");
      const s = who(), people = Number(d.people);
      if (!PRICES[people]) fail_("invalid");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date || "") || !/^\d{2}:00$/.test(d.time || "")) fail_("invalid");
      if (d.date < today_()) fail_("past");
      const r = db.add("Requests", { id: db.nextId("I"), people, price: PRICES[people], date: d.date, time: d.time, name: s.name, phone: s.phone, experience: d.experience, status: "ახალი", created: today_() });
      notify_("individual", r);
      return { ok: true, token: sToken(s.phone) };
    }

    case "register": {
      const phone = key_(d.phone), name = String(d.name || "").trim().slice(0, 60);
      if (name.length < 2 || phone.length !== 9) fail_("invalid");
      if (!/^\d{4}$/.test(d.pin || "")) fail_("pin_format");
      const s = student(phone);
      if (s) {
        checkLocked_(phone);
        if (s.pinHash !== hashPin_(phone, d.pin)) { badTry_(phone); fail_("exists"); }
        goodTry_(phone);
      } else {
        db.add("Students", { phone, name, pinHash: hashPin_(phone, d.pin), created: today_() });
      }
      return { ok: true, token: sToken(phone), profile: prof(phone) };
    }

    case "login": {
      const phone = key_(d.phone), s = student(phone);
      if (!s) fail_("no_account");
      checkLocked_(phone);
      if (s.pinHash !== hashPin_(phone, d.pin)) { badTry_(phone); fail_("pin_wrong"); }
      goodTry_(phone);
      return { ok: true, token: sToken(phone), profile: prof(phone) };
    }

    case "me": {
      const phone = readToken_(d.token, "s");
      if (!phone || !student(phone)) fail_("auth");
      return { ok: true, profile: prof(phone) };
    }

    case "admin_login": {
      if (String(ADMIN_PASSWORD).length < 6) fail_("set_password");
      checkLocked_("admin");
      if (String(d.password || "") !== ADMIN_PASSWORD) { badTry_("admin"); fail_("pin_wrong"); }
      goodTry_("admin");
      return { ok: true, token: makeToken_("a", "admin", ADMIN_HOURS) };
    }

    case "admin_day": admin(); return day();

    case "admin_mark": {
      admin();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date || "")) fail_("invalid");
      (d.marks || []).forEach(m => {
        const ex = att().find(a => a.date === d.date && a.member === m.member);
        const g = groups().find(x => x.id === m.member), r = inds().find(x => x.id === m.member), src = g || r;
        if (!src) return;
        if (r) {
          r.status = m.status === "მოვიდა" ? "ჩატარდა" : m.status === "გააცდინა" ? "გაცდენილი" : "დადასტურებული";
          db.save("Requests", r);
        }
        if (ex && !m.status) db.remove("Attendance", ex);
        else if (ex) { ex.status = m.status; db.save("Attendance", ex); }
        else if (m.status) db.add("Attendance", { date: d.date, kind: g ? "ჯგუფური" : "ინდივიდუალური",
          label: g ? (GROUPS[g.group] || {}).label : `ინდივიდუალური · ${r.people} ადამიანი`, member: m.member, phone: src.phone, status: m.status });
      });
      return day();
    }

    case "admin_requests": admin(); return reqs();

    case "admin_update_request": {
      admin();
      const r = inds().find(x => x.id === d.id); if (!r) fail_("invalid");
      if (d.date) r.date = d.date;
      if (d.time) r.time = d.time;
      if (d.status) r.status = d.status;
      db.save("Requests", r);
      return reqs();
    }

    case "admin_members": admin(); return mems();

    case "admin_students": {
      admin();
      return { ok: true, students: students().map(s => {
        const mine = groups().filter(g => g.phone === s.phone && active(g) && GROUPS[g.group]);
        const done = att().filter(a => a.phone === s.phone && a.status === "მოვიდა").map(a => a.date).sort();
        return { name: s.name, phone: s.phone, created: s.created, groups: mine.map(g => GROUPS[g.group].label),
          remaining: mine.length ? mine.reduce((x, g) => x + Math.max(0, (Number(g.purchased) || 0) - used(g.id)), 0) : null,
          attended: done.length, last: done[done.length - 1] || "",
          individual: inds().filter(r => r.phone === s.phone && r.status !== "გაუქმებული").length };
      }).sort((a, b) => (a.created < b.created ? 1 : -1)) };
    }

    case "admin_renew": { admin(); const g = member(); g.purchased = (Number(g.purchased) || 0) + (Number(d.add) || PASS); db.save("Groups", g); return mems(); }
    case "admin_confirm_payment": { admin(); const g = member(); g.pending = "FALSE"; db.save("Groups", g); return mems(); }
    case "admin_cancel_member": { admin(); const g = member(); g.active = "FALSE"; db.save("Groups", g); return mems(); }
  }
  fail_("bad_request");
}

function pubReq_(r) {
  return { id: r.id, people: Number(r.people), price: Number(r.price), date: r.date, time: r.time, name: r.name, phone: r.phone, experience: r.experience, status: r.status };
}

/* ---------- შეტყობინებები: მეილი + WhatsApp ---------- */
function notify_(kind, d) {
  const subject = kind === "group" ? "ახალი ჯავშანი ჯგუფში" : "ახალი ინდივიდუალური მოთხოვნა";
  const lines = ["სახელი: " + d.name, "ტელეფონი: " + d.phone];
  if (kind === "group") lines.push("ჯგუფი: " + d.group);
  else lines.push("თარიღი: " + d.date + " " + d.time, "ადამიანი: " + d.people + " · " + d.price + " ₾", "გამოცდილება: " + d.experience);
  const text = lines.join("\n");

  if (NOTIFY_EMAIL) {
    try { MailApp.sendEmail(NOTIFY_EMAIL, subject, text); }
    catch (e) { console.error("email notify failed: " + e); }
  }
  if (WHATSAPP_PHONE && CALLMEBOT_APIKEY) {
    try {
      const url = "https://api.callmebot.com/whatsapp.php?phone=" + encodeURIComponent(String(WHATSAPP_PHONE).replace(/\D/g, "")) +
        "&text=" + encodeURIComponent(subject + "\n\n" + text) + "&apikey=" + encodeURIComponent(CALLMEBOT_APIKEY);
      const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) console.error("whatsapp notify failed: " + r.getContentText());
    } catch (e) { console.error("whatsapp notify failed: " + e); }
  }
}

/* ---------- პირველი გაშვება ----------
   რედაქტორში აირჩიეთ setup და დააჭირეთ Run: შექმნის ცხრილს, მოითხოვს ნებართვებს
   და გამოგიგზავნით სატესტო შეტყობინებას. */
function setup() {
  const ss = book_();
  console.log("ცხრილი: " + ss.getUrl());
  notify_("individual", { name: "ტესტი", phone: "555123456", date: today_(), time: "19:00", people: 2, price: PRICES[2], experience: EXPERIENCE[1] });
  console.log("სატესტო შეტყობინება გაიგზავნა" + (NOTIFY_EMAIL ? " მეილზე" : "") + (WHATSAPP_PHONE && CALLMEBOT_APIKEY ? " და WhatsApp-ში" : ""));
}
