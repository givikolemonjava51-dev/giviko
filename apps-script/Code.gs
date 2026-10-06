/* ====== Giviko — სერვერი (Google Apps Script) ======
   მონაცემები ინახება Google Sheets ცხრილში, რომელიც პირველ გაშვებაზე თავად შეიქმნება.
   ახალ მოთხოვნაზე, ჯავშანზე, გაუქმებაზე და გადატანაზე შეტყობინება მოდის მეილზე და WhatsApp-ში.
   ვარჯიშამდე 24 და 2 საათით ადრე მოდის შეხსენება (ერთხელ გაუშვით installReminders).

   ⚠ შეავსეთ ქვემოთ მხოლოდ Apps Script-ში ჩასმულ ასლში, GitHub-ზე ნუ ატვირთავთ. */
const ADMIN_PASSWORD = "";     // ტრენერის პანელის პაროლი (მინ. 6 სიმბოლო)
const NOTIFY_EMAIL = "";       // მეილი, სადაც მოთხოვნები მოვა
const WHATSAPP_PHONE = "";     // WhatsApp ნომერი 995-ით, მაგ. "995555123456"
const CALLMEBOT_APIKEY = "";   // CallMeBot-ის გასაღები
const BANK_IBAN = "GE62TB7243645064300033"; // ანგარიში, რომელიც ინდივიდუალური ვარჯიშის მეილში წერია
const BANK_NAME = "გივი ლემონჯავა";

/* ---------- საიტის პარამეტრები (ემთხვევა index.html-ს) ---------- */
const TZ = "Asia/Tbilisi";
const CAPACITY = 4;
const PASS = 8;
// ინდივიდუალური ვარჯიშის 1 საათის ფასი ადამიანების მიხედვით:
// day = ორშ-პარ 18:00-მდე, eve = 18:00-დან და შაბ-კვ ნებისმიერ დროს (ითვლება დაწყების დროით)
const IND_PRICES = {
  day: { 1: 130, 2: 170, 3: 200, 4: 230 },
  eve: { 1: 150, 2: 190, 3: 220, 4: 250 },
};
const IND_EVE_FROM = "18:00";
const DURS = [60, 90, 120]; // ხანგრძლივობა წუთებში; 1.5 და 2 საათი პროპორციულია
// სტატისტიკა: ტრენერის სუფთა მოგება ერთ ვარჯიშზე
const PROFIT_IND_HOUR = 80;              // ინდივიდუალური, 1 საათზე (1.5 სთ = 120, 2 სთ = 160)
const PROFIT_GROUP = { 3: 75, 4: 120 };  // ჯგუფი ადამიანების მიხედვით; 2 ან ნაკლები = 0
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
const FREE_DAYS = 14; // ინდივიდუალურისთვის რამდენი დღის თავისუფალი დრო ჩანს საიტზე

/* ---------- ცხრილები ---------- */
const TABLES = {
  Students: ["phone", "name", "pinHash", "created", "email", "lang"],
  Groups: ["id", "group", "name", "phone", "active", "pending", "purchased", "created", "renew"],
  Requests: ["id", "people", "price", "date", "time", "name", "phone", "experience", "status", "created", "note", "duration"],
  Attendance: ["date", "kind", "label", "member", "phone", "status"],
  Reserved: ["group", "count"],
  Waitlist: ["id", "group", "name", "phone", "status", "created"],
  Changes: ["id", "kind", "ref", "name", "phone", "date", "time", "action", "newDate", "newTime", "status", "created"],
  Feedback: ["date", "member", "phone", "name", "label", "text", "updated"],
  Free: ["date", "from", "to", "source"],
};
// ცხრილების აღწერა: როცა იცვლება, ძველ ცხრილს ემატება ახალი ფურცლები და სვეტები
const SCHEMA = Object.keys(TABLES).map(k => k + ":" + TABLES[k].length).join(",");

function book_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SHEET_ID"), ss = null, blank = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create("Giviko — მონაცემები");
    blank = ss.getSheets()[0];
    props.setProperty("SHEET_ID", ss.getId());
  }
  if (!blank && props.getProperty("SCHEMA") === SCHEMA) return ss;
  Object.keys(TABLES).forEach(name => {
    const cols = TABLES[name], old = ss.getSheetByName(name);
    if (old) {
      const have = old.getLastColumn();
      if (have < cols.length) {
        old.getRange(1, have + 1, old.getMaxRows(), cols.length - have).setNumberFormat("@");
        old.getRange(1, have + 1, 1, cols.length - have).setValues([cols.slice(have)]).setFontWeight("bold");
      }
      return;
    }
    const sh = ss.insertSheet(name);
    sh.getRange("A:Z").setNumberFormat("@"); // ნომრები და თარიღები ტექსტად
    sh.getRange(1, 1, 1, TABLES[name].length).setValues([TABLES[name]]).setFontWeight("bold");
    sh.setFrozenRows(1);
    if (name === "Reserved") sh.getRange(2, 1, Object.keys(GROUPS).length, 2).setValues(Object.keys(GROUPS).map(g => [g, "0"]));
  });
  if (blank) ss.deleteSheet(blank);
  props.setProperty("SCHEMA", SCHEMA);
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
const isDate_ = s => /^\d{4}-\d{2}-\d{2}$/.test(s || "");
const isTime_ = s => /^\d{2}:\d{2}$/.test(s || "");
const mins_ = s => +s.slice(0, 2) * 60 + +s.slice(3);
const hhmm_ = m => pad_(Math.floor(m / 60)) + ":" + pad_(m % 60);
const dur_ = r => Number(r.duration) || 60; // ძველ ჩანაწერებში ხანგრძლივობა ცარიელია = 1 საათი
const durText_ = m => (m / 60) + " სთ";
const endTime_ = r => hhmm_(mins_(r.time) + dur_(r));
const nextDay_ = s => { const [y, m, d] = s.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d + 1)).toISOString().slice(0, 10); };
const groupProfit_ = n => PROFIT_GROUP[Math.min(n, 4)] || 0;
const hm_ = v => { const m = String(v || "").trim().match(/^(\d{1,2})[:.](\d{2})/); return m && +m[1] < 25 ? pad_(+m[1]) + ":" + m[2] : ""; };
const addDays_ = (s, n) => { const [y, m, d] = s.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10); };
const weekday_ = s => { const [y, m, d] = s.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); };
const indTier_ = (date, time) => { const w = weekday_(date); return w === 0 || w === 6 || time >= IND_EVE_FROM ? "eve" : "day"; };
const indPrice_ = (date, time, dur, people) => DURS.includes(dur) && IND_PRICES.day[people] ? Math.round(IND_PRICES[indTier_(date, time)][people] * dur / 60) : 0;
// მეილი არასავალდებულოა: ცარიელი = არ აქვს, არასწორი ფორმატი = შეცდომა
const email_ = v => {
  const x = String(v || "").trim().toLowerCase().slice(0, 100);
  if (x && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x)) fail_("email");
  return x;
};
const lang_ = v => (v === "en" ? "en" : "ka");

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

const WRITES = ["group", "individual", "register", "waitlist", "waitlist_leave", "renew", "cancel", "reschedule",
  "admin_mark", "admin_update_request", "admin_renew", "admin_confirm_payment", "admin_cancel_member", "admin_change", "admin_waitlist_remove", "admin_feedback", "set_email"];

// ხშირი წაკითხვები (განრიგი, კაბინეტი) ინახება ქეშში, რომ ცხრილი ყოველ ჯერზე არ გაიხსნას.
// ნებისმიერი ჩაწერა ქეშს აახლებს; ცხრილში ხელით შეცვლილი ციფრი საიტზე მაქსიმუმ CACHE_SEC წამში ჩანს.
const CACHE_SEC = 60;
const CACHED = ["counts", "me"];

function handle_(d) {
  d = d || {};
  const lock = WRITES.includes(d.action) ? LockService.getScriptLock() : null;
  const cache = CacheService.getScriptCache();
  let ck = null;
  if (CACHED.includes(d.action)) {
    const sub = d.action === "me" ? readToken_(d.token, "s") : "all";
    if (sub) {
      ck = "r:" + (cache.get("ver") || "0") + ":" + d.action + ":" + sub;
      const hit = cache.get(ck);
      if (hit) return JSON.parse(hit);
    }
  }
  try {
    if (lock) lock.waitLock(20000);
    const res = route_(d, Db_());
    if (ck && res && res.ok) { try { cache.put(ck, JSON.stringify(res), CACHE_SEC); } catch (e) {} }
    return res;
  } catch (e) {
    if (e && e.code) return { ok: false, error: e.code };
    console.error(e && e.stack || e);
    return { ok: false, error: "server" };
  } finally {
    if (lock) {
      SpreadsheetApp.flush();
      cache.put("ver", String(Date.now()), 21600); // ძველი ქეში აღარ გამოიყენება
      lock.releaseLock();
    }
  }
}

function route_(d, db) {
  const students = () => db.all("Students"), groups = () => db.all("Groups");
  const inds = () => db.all("Requests"), att = () => db.all("Attendance");
  const wl = () => db.all("Waitlist").filter(w => w.status === "active"), chg = () => db.all("Changes");
  const reserved = () => { const r = {}; db.all("Reserved").forEach(x => (r[x.group] = Number(x.count) || 0)); return r; };
  const fbs = () => db.all("Feedback");
  const used = id => att().filter(a => a.member === id).length;
  // ტრენერის თავისუფალი შუალედები (ფურცელი Free) წუთებში: { "2026-10-05": [[600, 780], [960, 1080]] }.
  // დაკავებულია: უკვე მოთხოვნილი ინდივიდუალური (მთელი ხანგრძლივობით), ჯგუფური საათები და დღევანდელი გასული დრო.
  // ფურცელი ცარიელია → null და საიტი ძველებურად ყველა საათს აჩვენებს.
  const free = () => {
    const rows = db.all("Free");
    if (!rows.length) return null;
    const t = today_(), last = addDays_(t, FREE_DAYS), now = mins_(Utilities.formatDate(new Date(), TZ, "HH:mm"));
    const open = {}, busy = {};
    const add = (m, date, a, b) => { if (b > a) (m[date] = m[date] || []).push([a, b]); };
    rows.forEach(r => {
      const date = String(r.date).trim(), from = hm_(r.from), to = hm_(r.to);
      if (isDate_(date) && from && to && date >= t && date <= last) add(open, date, mins_(from), mins_(to));
    });
    inds().filter(r => r.status === "ახალი" || r.status === "დადასტურებული")
      .forEach(r => isTime_(r.time) && add(busy, r.date, mins_(r.time), mins_(r.time) + dur_(r)));
    const out = {};
    Object.keys(open).forEach(date => {
      Object.values(GROUPS).forEach(g => g.days.includes(weekday_(date)) && add(busy, date, mins_(g.time), mins_(g.time) + 60));
      if (date === t) add(busy, date, 0, now);
      const w = minus_(merge_(open[date]), merge_(busy[date] || []));
      if (w.length) out[date] = w;
    });
    return out;
  };
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
      const email = email_(d.email);
      if (email && !s.email) { s.email = email; s.lang = lang_(d.lang); db.save("Students", s); }
    } else {
      s = db.add("Students", { phone, name, pinHash: hashPin_(phone, d.pin), created: today_(), email: email_(d.email), lang: lang_(d.lang) });
    }
    return s;
  };
  const admin = () => { if (readToken_(d.token, "a") !== "admin") fail_("auth"); };
  const prof = phone => {
    const s = student(phone); if (!s) fail_("auth");
    const mine = att().filter(a => a.phone === phone), month = today_().slice(0, 7);
    return {
      name: s.name, phone, email: s.email || "",
      memberships: groups().filter(g => g.phone === phone && active(g) && GROUPS[g.group]).map(g => {
        const G = GROUPS[g.group], u = used(g.id), p = Number(g.purchased) || 0;
        return { id: g.id, group: g.group, label: G.label, days: G.days, time: G.time, since: g.created, pending: bool_(g.pending), renew: bool_(g.renew), purchased: p, used: u, remaining: Math.max(0, p - u) };
      }),
      individual: inds().filter(r => r.phone === phone && r.status !== "გაუქმებული").map(pubReq_),
      waitlist: wl().filter(w => w.phone === phone && GROUPS[w.group]).map(w => ({ id: w.id, group: w.group, label: GROUPS[w.group].label })),
      changes: chg().filter(c => c.phone === phone).map(pubChg_),
      feedback: fbs().filter(f => f.phone === phone && f.text).sort((a, b) => (a.date < b.date ? 1 : -1)).map(f => ({ date: f.date, label: f.label, text: f.text })),
      history: mine.slice().sort((a, b) => (a.date < b.date ? 1 : -1)).map(a => ({ date: a.date, kind: a.kind, label: a.label, status: a.status })),
      stats: {
        attended: mine.filter(a => a.status === "მოვიდა").length,
        missed: mine.filter(a => a.status === "გააცდინა").length,
        month: mine.filter(a => a.status === "მოვიდა" && a.date.slice(0, 7) === month).length,
      },
    };
  };
  const day = () => {
    if (!isDate_(d.date)) fail_("invalid");
    const wd = weekday_(d.date), mark = id => (att().find(a => a.date === d.date && a.member === id) || {}).status || "";
    const fb = id => (fbs().find(f => f.date === d.date && f.member === id) || {}).text || "";
    const left = m => Math.max(0, (Number(m.purchased) || 0) - used(m.id));
    // მოსწავლემ გააფრთხილა, რომ ვერ მოვა, ან სხვა დღეზე გადაიტანა
    const away = id => chg().find(c => c.ref === id && c.date === d.date && (c.action === "cancel" || c.status !== "უარყოფილი"));
    const sessions = [];
    Object.values(GROUPS).forEach(g => {
      if (!g.days.includes(wd)) return;
      sessions.push({ kind: "group", ref: g.id, time: g.time, label: g.label,
        members: groups().filter(m => m.group === g.id && active(m) && !bool_(m.pending)).map(m => {
          const c = away(m.id);
          return { id: m.id, name: m.name, phone: m.phone, remaining: left(m), mark: mark(m.id), feedback: fb(m.id), away: !!c && (c.action === "cancel" || c.status === "დადასტურებული"),
            note: !c ? "" : c.action === "cancel" ? "გააფრთხილა: ვერ მოვა" : (c.status === "ახალი" ? "ითხოვს გადატანას: " : "გადატანილია: ") + c.newDate + " " + c.newTime };
        }) });
    });
    // ანაზღაურებითი ვარჯიშები: დადასტურებული გადატანა ამ დღეზე
    chg().filter(c => c.action === "reschedule" && c.status === "დადასტურებული" && c.newDate === d.date).forEach(c => {
      const m = groups().find(x => x.id === c.ref); if (!m) return;
      const row = { id: m.id, name: m.name, phone: m.phone, remaining: left(m), mark: mark(m.id), feedback: fb(m.id), note: "ანაზღაურება (" + c.date + "-ის ნაცვლად)" };
      const s = sessions.find(x => x.kind === "group" && x.time === c.newTime);
      if (s) s.members.push(row);
      else sessions.push({ kind: "makeup", ref: c.id, time: c.newTime, label: "ანაზღაურებითი ვარჯიში", members: [row] });
    });
    inds().filter(r => r.date === d.date && r.status !== "გაუქმებული").forEach(r => sessions.push({
      kind: "individual", ref: r.id, time: r.time, label: `ინდივიდუალური · ${r.time}-${endTime_(r)} · ${r.people} ადამიანი · ${r.price} ₾`,
      members: [{ id: r.id, name: r.name, phone: r.phone, mark: mark(r.id), feedback: fb(r.id) }] }));
    sessions.sort((a, b) => (a.time < b.time ? -1 : 1));
    return { ok: true, date: d.date, sessions };
  };
  const reqs = () => ({ ok: true, requests: inds()
    .filter(r => r.status === "ახალი" || (r.status === "დადასტურებული" && r.date >= today_()))
    .sort((a, b) => (a.date + a.time < b.date + b.time ? -1 : 1)).map(pubReq_),
    changes: chg().filter(c => c.action === "reschedule" && c.status === "ახალი").map(c => {
      const g = groups().find(x => x.id === c.ref);
      return Object.assign(pubChg_(c), { label: g && GROUPS[g.group] ? GROUPS[g.group].label : "" });
    }) });
  const mems = () => ({ ok: true, reserved: reserved(), members: groups().filter(g => active(g) && GROUPS[g.group]).map(g => {
    const u = used(g.id), p = Number(g.purchased) || 0;
    return { id: g.id, group: g.group, label: GROUPS[g.group].label, name: g.name, phone: g.phone, pending: bool_(g.pending), renew: bool_(g.renew), purchased: p, used: u, remaining: Math.max(0, p - u) };
  }), waitlist: wl().filter(w => GROUPS[w.group]).map(w => ({ id: w.id, group: w.group, name: w.name, phone: w.phone, created: w.created })) });
  const seatFreed = gid => {
    const next = wl().filter(w => w.group === gid);
    if (next.length) notify_("ადგილი გათავისუფლდა", ["ჯგუფი: " + GROUPS[gid].label, "მოლოდინის სიაში " + next.length + " ადამიანია. პირველი:", next[0].name + " · " + next[0].phone]);
  };
  // სტატისტიკა: ჩატარებული ვარჯიშები from-დან to-მდე (ორივე ჩათვლით), საათები და სუფთა მოგება
  const stats = () => {
    if (!isDate_(d.from) || !isDate_(d.to) || d.from > d.to) fail_("invalid");
    const now = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm"), res = reserved(), out = [];
    const over = (date, end) => date + " " + end <= now; // მხოლოდ დასრულებული ვარჯიშები
    for (let date = d.from, n = 0; date <= d.to && n < 400; date = nextDay_(date), n++) {
      const wd = weekday_(date);
      Object.values(GROUPS).forEach(g => {
        if (!g.days.includes(wd) || !over(date, hhmm_(mins_(g.time) + 60))) return;
        // ჯგუფის ზომა: საიტის გარეშე დაკავებული (Reserved) + გადახდილი წევრები, ვინც ამ დღისთვის უკვე ჩაწერილი იყო
        const people = (res[g.id] || 0) + groups().filter(m => m.group === g.id && active(m) && !bool_(m.pending) && m.created <= date).length;
        if (people) out.push({ date, time: g.time, kind: "group", label: g.label, people, minutes: 60, income: groupProfit_(people) });
      });
    }
    inds().filter(r => r.date >= d.from && r.date <= d.to && (r.status === "დადასტურებული" || r.status === "ჩატარდა") && over(r.date, endTime_(r)))
      .forEach(r => out.push({ date: r.date, time: r.time, kind: "individual", label: "ინდივიდუალური · " + r.name, people: Number(r.people) || 1, minutes: dur_(r), income: PROFIT_IND_HOUR * dur_(r) / 60 }));
    out.sort((a, b) => (a.date + a.time < b.date + b.time ? -1 : 1));
    return { ok: true, from: d.from, to: d.to, sessions: out };
  };
  const member = () => { const g = groups().find(x => x.id === d.member); if (!g) fail_("invalid"); return g; };

  switch (d.action) {
    case "counts": return { ok: true, counts: cnt(), free: free() };

    case "group": {
      const s = who(), g = GROUPS[d.group]; if (!g) fail_("invalid");
      const inG = groups().filter(x => x.group === g.id && active(x));
      if (inG.some(x => x.phone === s.phone)) fail_("already");
      if (inG.length + (reserved()[g.id] || 0) >= CAPACITY) fail_("full");
      db.add("Groups", { id: db.nextId("G"), group: g.id, name: s.name, phone: s.phone, active: "TRUE", pending: "TRUE", purchased: PASS, created: today_() });
      wl().filter(w => w.group === g.id && w.phone === s.phone).forEach(w => { w.status = "joined"; db.save("Waitlist", w); });
      notify_("ახალი ჯავშანი ჯგუფში", ["სახელი: " + s.name, "ტელეფონი: " + s.phone, "ჯგუფი: " + g.label]);
      return { ok: true, counts: cnt(), token: sToken(s.phone) };
    }

    case "individual": {
      if (!d.token) fail_("login_required");
      if (!EXPERIENCE.includes(d.experience)) fail_("experience");
      const s = who(), people = Number(d.people), duration = Number(d.duration) || 60;
      if (!DURS.includes(duration) || !IND_PRICES.day[people]) fail_("invalid");
      if (mins_(d.time || "00:00") + duration > 24 * 60) fail_("invalid");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date || "") || !/^\d{2}:00$/.test(d.time || "")) fail_("invalid");
      if (d.date < today_()) fail_("past");
      const f = free();
      if (f && !fits_(f[d.date], mins_(d.time), duration)) fail_("slot_taken");
      const r = db.add("Requests", { id: db.nextId("I"), people, duration, price: indPrice_(d.date, d.time, duration, people), date: d.date, time: d.time, name: s.name, phone: s.phone, experience: d.experience, status: "ახალი", created: today_() });
      notify_("ახალი ინდივიდუალური მოთხოვნა", reqLines_(r));
      return { ok: true, token: sToken(s.phone), free: free() };
    }

    case "waitlist": {
      const s = who(), g = GROUPS[d.group]; if (!g) fail_("invalid");
      const inG = groups().filter(x => x.group === g.id && active(x));
      if (inG.some(x => x.phone === s.phone)) fail_("already");
      if (inG.length + (reserved()[g.id] || 0) < CAPACITY) fail_("not_full");
      if (wl().some(w => w.group === g.id && w.phone === s.phone)) fail_("already_wait");
      db.add("Waitlist", { id: db.nextId("W"), group: g.id, name: s.name, phone: s.phone, status: "active", created: today_() });
      const pos = wl().filter(w => w.group === g.id).length;
      notify_("მოლოდინის სიაში ჩაეწერა", ["სახელი: " + s.name, "ტელეფონი: " + s.phone, "ჯგუფი: " + g.label, "რიგში: " + pos]);
      return { ok: true, position: pos, token: sToken(s.phone) };
    }

    case "waitlist_leave": {
      if (!d.token) fail_("auth");
      const s = who(), w = wl().find(x => x.id === d.id && x.phone === s.phone); if (!w) fail_("invalid");
      w.status = "left"; db.save("Waitlist", w);
      return { ok: true, profile: prof(s.phone) };
    }

    case "renew": {
      if (!d.token) fail_("auth");
      const s = who(), g = groups().find(x => x.id === d.member && x.phone === s.phone && active(x) && GROUPS[x.group]);
      if (!g) fail_("invalid");
      if (!bool_(g.renew)) {
        g.renew = "TRUE"; db.save("Groups", g);
        const u = used(g.id), p = Number(g.purchased) || 0;
        notify_("აბონემენტის გაგრძელება", ["სახელი: " + s.name, "ტელეფონი: " + s.phone, "ჯგუფი: " + GROUPS[g.group].label, "დარჩა: " + Math.max(0, p - u) + " / " + p, "ჩარიცხვის შემდეგ დააჭირე „+8 განახლება“."]);
      }
      return { ok: true, profile: prof(s.phone) };
    }

    case "cancel":
    case "reschedule": {
      if (!d.token) fail_("auth");
      const s = who(), moving = d.action === "reschedule";
      if (moving && (!isDate_(d.newDate) || !isTime_(d.newTime) || d.newDate < today_())) fail_("invalid");
      const who2 = ["სახელი: " + s.name, "ტელეფონი: " + s.phone];
      if (d.kind === "individual") {
        const r = inds().find(x => x.id === d.ref && x.phone === s.phone);
        if (!r || !["ახალი", "დადასტურებული"].includes(r.status) || r.date < today_()) fail_("invalid");
        const was = r.date + " " + r.time;
        if (moving) { r.note = "გადატანა " + was + "-დან"; r.date = d.newDate; r.time = d.newTime; r.status = "ახალი"; }
        else r.status = "გაუქმებული";
        db.save("Requests", r);
        notify_(moving ? "ინდივიდუალური ვარჯიშის გადატანის მოთხოვნა" : "ინდივიდუალური ვარჯიში გაუქმდა",
          who2.concat(["იყო: " + was], moving ? ["ახალი დრო: " + r.date + " " + r.time, "დაადასტურე ტრენერის პანელში."] : []));
      } else if (d.kind === "group") {
        const g = groups().find(x => x.id === d.ref && x.phone === s.phone && active(x)), G = g && GROUPS[g.group];
        if (!G || !isDate_(d.date) || d.date < today_() || !G.days.includes(weekday_(d.date))) fail_("invalid");
        if (chg().some(c => c.ref === g.id && c.date === d.date && c.status !== "უარყოფილი")) fail_("already_changed");
        db.add("Changes", { id: db.nextId("C"), kind: "group", ref: g.id, name: s.name, phone: s.phone, date: d.date, time: G.time,
          action: moving ? "reschedule" : "cancel", newDate: moving ? d.newDate : "", newTime: moving ? d.newTime : "", status: moving ? "ახალი" : "მიღებული", created: today_() });
        notify_(moving ? "ჯგუფური ვარჯიშის გადატანის მოთხოვნა" : "მოსწავლე ვერ მოვა ვარჯიშზე",
          who2.concat(["ჯგუფი: " + G.label, "თარიღი: " + d.date], moving ? ["სურს: " + d.newDate + " " + d.newTime, "დაადასტურე ტრენერის პანელში."] : []));
      } else fail_("invalid");
      return { ok: true, profile: prof(s.phone) };
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
        const email = email_(d.email);
        if (email) { s.email = email; s.lang = lang_(d.lang); db.save("Students", s); }
      } else {
        db.add("Students", { phone, name, pinHash: hashPin_(phone, d.pin), created: today_(), email: email_(d.email), lang: lang_(d.lang) });
      }
      return { ok: true, token: sToken(phone), profile: prof(phone) };
    }

    case "set_email": {
      const phone = readToken_(d.token, "s"), s = phone && student(phone);
      if (!s) fail_("auth");
      s.email = email_(d.email); s.lang = lang_(d.lang);
      db.save("Students", s);
      return { ok: true, profile: prof(phone) };
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
      if (!isDate_(d.date)) fail_("invalid");
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
          label: g ? (GROUPS[g.group] || {}).label : `ინდივიდუალური · ${durText_(dur_(r))} · ${r.people} ადამიანი`, member: m.member, phone: src.phone, status: m.status });
      });
      return day();
    }

    case "admin_feedback": {
      admin();
      if (!isDate_(d.date)) fail_("invalid");
      const g = groups().find(x => x.id === d.member), r = inds().find(x => x.id === d.member), src = g || r;
      if (!src) fail_("invalid");
      const text = String(d.text || "").trim().slice(0, 1000), ex = fbs().find(f => f.date === d.date && f.member === d.member);
      if (ex && !text) db.remove("Feedback", ex);
      else if (ex) { ex.text = text; ex.updated = today_(); db.save("Feedback", ex); }
      else if (text) db.add("Feedback", { date: d.date, member: d.member, phone: src.phone, name: src.name, text, updated: today_(),
        label: g ? (GROUPS[g.group] || {}).label : `ინდივიდუალური · ${durText_(dur_(r))}` });
      return { ok: true, member: d.member, feedback: text };
    }

    case "admin_stats": admin(); return stats();

    case "admin_requests": admin(); return reqs();

    case "admin_update_request": {
      admin();
      const r = inds().find(x => x.id === d.id); if (!r) fail_("invalid");
      const was = r.status + "|" + r.date + "|" + r.time;
      if (d.date) r.date = d.date;
      if (d.time) r.time = d.time;
      if (d.status) r.status = d.status;
      db.save("Requests", r);
      // დადასტურებისას (ან დადასტურებულის დროის შეცვლისას) მოსწავლეს მეილი მიდის
      const OK = "დადასტურებული", changed = was !== r.status + "|" + r.date + "|" + r.time;
      const emailed = r.status === OK && changed ? mailConfirm_(student(r.phone), "individual", r, was.split("|")[0] === OK) : false;
      return Object.assign(reqs(), { emailed });
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

    case "admin_renew": { admin(); const g = member(); g.purchased = (Number(g.purchased) || 0) + (Number(d.add) || PASS); g.renew = "FALSE"; db.save("Groups", g); return mems(); }
    case "admin_confirm_payment": {
      admin(); const g = member(), was = bool_(g.pending); g.pending = "FALSE"; db.save("Groups", g);
      const emailed = was && GROUPS[g.group] ? mailConfirm_(student(g.phone), "group", g, false) : false;
      return Object.assign(mems(), { emailed });
    }
    case "admin_cancel_member": {
      admin(); const g = member(), was = active(g); g.active = "FALSE"; db.save("Groups", g);
      if (was && GROUPS[g.group]) seatFreed(g.group);
      return mems();
    }

    case "admin_change": {
      admin();
      const c = chg().find(x => x.id === d.id);
      if (!c || !["დადასტურებული", "უარყოფილი"].includes(d.status)) fail_("invalid");
      c.status = d.status; db.save("Changes", c);
      return reqs();
    }

    case "admin_waitlist_remove": {
      admin();
      const w = wl().find(x => x.id === d.id); if (!w) fail_("invalid");
      w.status = "removed"; db.save("Waitlist", w);
      return mems();
    }
  }
  fail_("bad_request");
}

function pubReq_(r) {
  return { id: r.id, people: Number(r.people), duration: dur_(r), price: Number(r.price), date: r.date, time: r.time, name: r.name, phone: r.phone, experience: r.experience, status: r.status, note: r.note || "" };
}
function pubChg_(c) {
  return { id: c.id, kind: c.kind, ref: c.ref, name: c.name, phone: c.phone, date: c.date, time: c.time, action: c.action, newDate: c.newDate, newTime: c.newTime, status: c.status };
}
function reqLines_(r) {
  return ["სახელი: " + r.name, "ტელეფონი: " + r.phone, "თარიღი: " + r.date + " " + r.time + "-" + endTime_(r), "ხანგრძლივობა: " + durText_(dur_(r)), "ადამიანი: " + r.people + " · " + r.price + " ₾", "გამოცდილება: " + r.experience];
}

/* ---------- შეტყობინებები: მეილი + WhatsApp ---------- */
// html: მეილის ვერსია ბმულებით (არასავალდებულო). WhatsApp-ში მიდის მოკლე ტექსტი.
function notify_(subject, lines, html) {
  const text = lines.join("\n");
  if (NOTIFY_EMAIL) {
    try { MailApp.sendEmail(NOTIFY_EMAIL, subject, text, html ? { htmlBody: html } : {}); }
    catch (e) { console.error("email notify failed: " + e); }
  }
  if (WHATSAPP_PHONE && CALLMEBOT_APIKEY) {
    try {
      const base = "https://api.callmebot.com/whatsapp.php?phone=" + encodeURIComponent(String(WHATSAPP_PHONE).replace(/\D/g, "")) +
        "&apikey=" + encodeURIComponent(CALLMEBOT_APIKEY) + "&text=";
      // Apps Script-ში ბმული 2000 სიმბოლოზე გრძელი ვერ იქნება, ამიტომ გრძელ ტექსტს ვჭრით
      let msg = subject + "\n\n" + text;
      while (base.length + encodeURIComponent(msg).length > 2000) msg = msg.slice(0, Math.floor(msg.length * 0.9));
      const r = UrlFetchApp.fetch(base + encodeURIComponent(msg), { muteHttpExceptions: true });
      const body = r.getContentText().replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
      // CallMeBot returns 200 even for errors, so log its answer to see what happened
      console.log("WhatsApp (CallMeBot): " + r.getResponseCode() + " " + body.slice(0, 300));
    } catch (e) { console.error("whatsapp notify failed: " + e); }
  }
}

/* ---------- დასტურის მეილი მოსწავლეს ----------
   იგზავნება თქვენი Gmail-იდან (სკრიპტი თქვენი სახელით მუშაობს), პასუხი მოდის NOTIFY_EMAIL-ზე.
   მეილი მიდის მხოლოდ იმ მოსწავლესთან, ვინც რეგისტრაციისას ან კაბინეტში მეილი მიუთითა. */
const WD_KA_ = ["კვირა", "ორშაბათი", "სამშაბათი", "ოთხშაბათი", "ხუთშაბათი", "პარასკევი", "შაბათი"];
const WD_EN_ = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const MON_KA_ = ["იანვარი", "თებერვალი", "მარტი", "აპრილი", "მაისი", "ივნისი", "ივლისი", "აგვისტო", "სექტემბერი", "ოქტომბერი", "ნოემბერი", "დეკემბერი"];
const MON_EN_ = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
function niceDate_(s, en) {
  const [, m, d] = s.split("-").map(Number), w = weekday_(s);
  return en ? `${WD_EN_[w]}, ${MON_EN_[m - 1]} ${d}` : `${d} ${MON_KA_[m - 1]}, ${WD_KA_[w]}`;
}
const escHtml_ = v => String(v ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

const MON_KA_SHORT_ = ["იან", "თებ", "მარ", "აპრ", "მაი", "ივნ", "ივლ", "აგვ", "სექ", "ოქტ", "ნოე", "დეკ"];
const MON_EN_SHORT_ = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function sendStudentMail_(to, subject, lines) {
  const body = lines.join("\n");
  const html = `<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.5;color:#1d2a2a">${lines.map(l => l ? `<p style="margin:0">${escHtml_(l)}</p>` : "<br>").join("")}</div>`;
  try {
    const opts = { to, subject, body, htmlBody: html, name: "Giviko" };
    if (NOTIFY_EMAIL) opts.replyTo = NOTIFY_EMAIL;
    MailApp.sendEmail(opts);
    return true;
  } catch (e) { console.error("student email failed: " + e); return false; }
}

function mailConfirm_(s, kind, x, moved) {
  if (!s || !s.email) return false;
  const en = s.lang === "en", T = (ka, eng) => (en ? eng : ka), first = String(s.name || "").trim().split(" ")[0];
  if (kind === "individual") {
    // ტექსტი ზუსტად ემთხვევა პანელის „WhatsApp დასტური“ ღილაკის ტექსტს (okWaText, index.html)
    const [, m, d] = x.date.split("-").map(Number), w = weekday_(x.date);
    const when = en ? `${WD_EN_[w]}, ${MON_EN_SHORT_[m - 1]} ${d}` : `${d} ${MON_KA_SHORT_[m - 1]}, ${WD_KA_[w]}`;
    const subject = moved ? T("ვარჯიშის დრო შეიცვალა", "Your training time has changed") : T("ვარჯიშის ჯავშანი", "Your training booking");
    return sendStudentMail_(s.email, subject, en ? [
      `Hi, ${first}! Training: ${when}, ${x.time}-${endTime_(x)}, Padelbade Krtsanisi - ${x.price} ₾`, "",
      `To confirm, transfer ${x.price} ₾ to this account: #${BANK_IBAN} (${BANK_NAME}).`, "",
      "Your training is booked as soon as the payment arrives",
    ] : [
      `გამარჯობა, ${first}! ვარჯიში: ${when}, ${x.time}-${endTime_(x)}, Padelbade Krtsanisi - ${x.price} ₾`, "",
      `დასადასტურებლად გადმორიცხე ${x.price} ₾ ამ ანგარიშზე: #${BANK_IBAN} (${BANK_NAME}).`, "",
      "ვარჯიში დაიჯავშნება ჩარიცხვისთანავე",
    ]);
  }
  const G = GROUPS[x.group];
  return sendStudentMail_(s.email, T("ჯგუფში ჩარიცხვა დადასტურებულია ✓", "You're in the group ✓"), [
    T("გამარჯობა", "Hi") + ", " + first + "!", "",
    T("გადახდა მივიღე, ჯგუფში ჩარიცხული ხარ.", "I've received your payment, you're in the group."), "",
    T("ჯგუფი", "Group") + ": " + G.label,
    T("აბონემენტი", "Pass") + ": " + x.purchased + T(" გაკვეთილი", " lessons"),
    T("ადგილი", "Place") + ": " + PLACE, "",
    "Giviko",
  ]);
}

/* ---------- შეხსენებები: ვარჯიშამდე 24 და 2 საათით ადრე ----------
   CallMeBot მხოლოდ თქვენს ნომერზე წერს, ამიტომ შეხსენება მოდის თქვენთან: ვინ მოდის და
   მეილში თითოეულ მოსწავლესთან WhatsApp ბმული მზა ტექსტით (ერთი დაჭერით გაგზავნა).
   იგივე ღილაკები ტრენერის პანელშიც არის. ჩასართავად ერთხელ გაუშვით installReminders. */
const PLACE = "Padelbade Krtsanisi, გიორგი გურამიშვილის ქ. 7";

function remindText_(name, date, time) {
  return "გამარჯობა " + String(name || "").split(" ")[0] + "! შეგახსენებ, " + (date === today_() ? "დღეს" : "ხვალ") + " " + time +
    "-ზე პადელის ვარჯიში გვაქვს. " + PLACE + ". თუ ვერ მოდიხარ, გააუქმე კაბინეტიდან ან მომწერე. Giviko";
}

function reminders() {
  const db = Db_(), now = Date.now(), props = PropertiesService.getScriptProperties();
  let sent = [];
  try { sent = JSON.parse(props.getProperty("REMINDED") || "[]"); } catch (e) {}
  const tok = makeToken_("a", "admin", 1), due = [];
  [now, now + 864e5].map(t => Utilities.formatDate(new Date(t), TZ, "yyyy-MM-dd")).forEach(date => {
    route_({ action: "admin_day", token: tok, date }, db).sessions.forEach(s => {
      const h = (new Date(date + "T" + s.time + ":00+04:00").getTime() - now) / 36e5; // თბილისი UTC+4
      const kind = h > 23 && h <= 24 ? "24" : h > 1 && h <= 2 ? "2" : "";
      const people = s.members.filter(m => !m.away);
      const key = date + " " + s.time + " " + s.ref + " " + kind;
      if (!kind || !people.length || sent.includes(key)) return;
      sent.push(key);
      due.push({ date, s, kind, people });
    });
  });
  if (!due.length) return;
  props.setProperty("REMINDED", JSON.stringify(sent.slice(-300)));
  due.forEach(({ date, s, kind, people }) => {
    const when = kind === "2" ? "2 საათში" : "24 საათში";
    const wa = m => "https://wa.me/995" + m.phone + "?text=" + encodeURIComponent(remindText_(m.name, date, s.time));
    notify_("შეხსენება: ვარჯიში " + when + " · " + s.time,
      [date + " " + s.time + " · " + s.label].concat(people.map(m => m.name + " · " + m.phone), ["შეხსენების გაგზავნა: ტრენერის პანელი → დასწრება → WhatsApp"]),
      "<p><b>" + date + " " + s.time + "</b> · " + s.label + "</p><p>შეხსენების გასაგზავნად დააჭირე სახელს:</p><ul>" +
        people.map(m => '<li><a href="' + wa(m) + '">' + m.name + " · " + m.phone + "</a></li>").join("") + "</ul>");
  });
}

function installReminders() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "reminders").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("reminders").timeBased().everyHours(1).create();
  console.log("შეხსენებები ჩაირთო: ყოველ საათში შემოწმდება, ხომ არ არის ვარჯიში 24 ან 2 საათში.");
}

/* ---------- კლუბის (Book&Go) განრიგიდან თავისუფალი საათები ----------
   Book&Go-ს მხარდაჭერის მიერ მოცემული გასაღები ჩაწერეთ მხოლოდ Apps Script-ში:
   Project Settings → Script Properties → Add property → CLUB_API_KEY (აქ, კოდში ან საიტზე არასდროს).
   სურვილისამებრ: WORK_HOURS, მაგ. 09:00-23:00 — გამოიყენება მხოლოდ მაშინ, თუ კლუბის სამუშაო საათები ვერ წაიკითხა.
   შემდეგ ერთხელ გაუშვით installClubSync: ყოველ 15 წუთში ფურცელში Free ჩაიწერება
   სამუშაო საათებს გამოკლებული კლუბში დაჯავშნილი ვარჯიშები (source = club).
   ფურცელში ხელით ჩაწერილი სტრიქონები (source ცარიელი) უცვლელი რჩება. */
const CLUB_API = "https://docsapi.bookandgo.app/api/coach/";

function clubGet_(path) {
  const key = String(PropertiesService.getScriptProperties().getProperty("CLUB_API_KEY") || "").trim();
  if (!key) throw new Error("CLUB_API_KEY ცარიელია (Project Settings → Script Properties).");
  const res = UrlFetchApp.fetch(CLUB_API + path, { headers: { Authorization: "Bearer " + key }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error("Book&Go " + path.split("?")[0] + ": HTTP " + res.getResponseCode() + " " + res.getContentText().slice(0, 200));
  return JSON.parse(res.getContentText());
}

function syncClubFree() {
  const t = today_(), last = addDays_(t, FREE_DAYS);
  const sched = clubGet_("schedule?start_date=" + t + "&end_date=" + last);
  let wh = null;
  try { wh = clubGet_("working-hours"); } catch (e) { console.warn(e.message); }
  const hours = workingHours_(wh);
  if (!hours) console.warn("სამუშაო საათები ვერ წავიკითხე, ვიყენებ WORK_HOURS-ს. გაუშვით testClubApi და პასუხი გაუგზავნეთ დეველოპერს.");
  const fallback = String(PropertiesService.getScriptProperties().getProperty("WORK_HOURS") || "08:00-23:00").split("-").map(hm_);
  const busy = {};
  (sched.sessions || []).forEach(s => {
    const a = hm_(s.start_time_local), b = hm_(s.end_time_local);
    if (!isDate_(s.date) || !a || !b) return;
    const to = mins_(b) > mins_(a) ? mins_(b) : 1440; // შუაღამეს გადასული ვარჯიში
    (busy[s.date] = busy[s.date] || []).push([mins_(a), to]);
  });
  const rows = [];
  for (let date = t; date <= last; date = addDays_(date, 1)) {
    let open;
    if (hours) open = hours.special[date] || hours.weekly[weekday_(date)] || [];
    else open = fallback[0] && fallback[1] ? [[mins_(fallback[0]), mins_(fallback[1])]] : [];
    minus_(merge_(open), merge_(busy[date] || [])).forEach(w => rows.push([date, hhmm_(w[0]), hhmm_(w[1]), "club"]));
  }
  const sh = book_().getSheetByName("Free");
  const manual = Db_().all("Free").filter(r => r.source !== "club").map(r => TABLES.Free.map(c => r[c]));
  const all = manual.concat(rows);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TABLES.Free.length).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, TABLES.Free.length).setNumberFormat("@").setValues(all);
  CacheService.getScriptCache().put("ver", String(Date.now()), 21600); // საიტი მაშინვე ხედავს ახალ საათებს
  console.log("Book&Go: " + (sched.sessions || []).length + " ვარჯიში, ჩაიწერა " + rows.length + " თავისუფალი შუალედი.");
}

function installClubSync() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "syncClubFree").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("syncClubFree").timeBased().everyMinutes(15).create();
  syncClubFree();
  console.log("კლუბის განრიგის სინქრონიზაცია ჩაირთო: ყოველ 15 წუთში.");
}

// შემოწმება: აჩვენებს, რას აბრუნებს Book&Go (კლიენტების სახელები პასუხში არ არის).
function testClubApi() {
  console.log("working-hours:\n" + JSON.stringify(clubGet_("working-hours"), null, 1).slice(0, 6000));
  const s = clubGet_("schedule");
  console.log("schedule: " + (s.sessions || []).length + " ვარჯიში, პირველი: " + JSON.stringify((s.sessions || [])[0] || null));
  console.log("ასე წაიკითხა სამუშაო საათები (0 = კვირა): " + JSON.stringify(workingHours_(clubGet_("working-hours"))));
}

/* Book&Go-ს სამუშაო საათები → { weekly: { 0..6: [[from, to]] }, special: { "YYYY-MM-DD": [[from, to]] } } წუთებში.
   ორივე პროფილის (Day და Evening/Weekends) საათები ერთიანდება; შესვენებები აკლდება; unavailable = დასვენების დღე. */
const DAY_NAMES_ = { sun: 0, mon: 1, tue: 2, wed: 3, thu: 4, fri: 5, sat: 6 };
function workingHours_(json) {
  if (!json) return null;
  const weekly = {}, special = {}, specialOff = {};
  const add = (m, k, w) => (m[k] = m[k] || []).push(w);
  const pick = (o, keys) => { for (const k of keys) if (o[k] !== undefined && o[k] !== null) return o[k]; };
  const span = o => {
    const a = hm_(pick(o, ["start_time_local", "start_time", "start", "from", "open", "opens_at", "time_from"]));
    const b = hm_(pick(o, ["end_time_local", "end_time", "end", "to", "close", "closes_at", "time_to"]));
    return a && b ? [mins_(a), mins_(b) > mins_(a) ? mins_(b) : 1440] : null;
  };
  const dayOf = o => {
    const v = pick(o, ["weekday", "day_of_week", "dayOfWeek", "day", "dow", "week_day"]);
    if (typeof v === "string" && isNaN(+v)) return DAY_NAMES_[v.trim().slice(0, 3).toLowerCase()];
    if (v === undefined || v === "") return undefined;
    return +v === 7 ? 0 : +v; // 0 ან 7 = კვირა, 1 = ორშაბათი
  };
  const breaks = o => Object.keys(o).filter(k => /break/i.test(k) && Array.isArray(o[k])).reduce((a, k) => a.concat(o[k].map(span).filter(Boolean)), []);
  const walk = (o, dayHint) => {
    if (Array.isArray(o)) return o.forEach(x => walk(x, dayHint));
    if (!o || typeof o !== "object") return;
    const date = isDate_(String(o.date || "").slice(0, 10)) ? String(o.date).slice(0, 10) : "";
    const day = dayOf(o) !== undefined ? dayOf(o) : dayHint;
    const w = span(o);
    if (date) {
      if (o.unavailable === true || o.is_day_off === true || o.closed === true) specialOff[date] = true;
      else if (w) minus_([w], breaks(o)).forEach(x => add(special, date, x));
    } else if (w && day >= 0 && day <= 6 && o.unavailable !== true) {
      minus_([w], breaks(o)).forEach(x => add(weekly, day, x));
    }
    Object.keys(o).forEach(k => {
      if (/break/i.test(k)) return;
      const hint = DAY_NAMES_[k.slice(0, 3).toLowerCase()];
      if (o[k] && typeof o[k] === "object") walk(o[k], hint !== undefined && k.length >= 3 && /^[a-z]+$/i.test(k) ? hint : day);
    });
  };
  walk(json);
  Object.keys(specialOff).forEach(d => (special[d] = special[d] || []));
  if (!Object.keys(weekly).length) return null;
  return { weekly, special };
}

// შუალედების გაერთიანება და გამოკლება (წუთებში)
function merge_(list) {
  const out = [];
  list.slice().sort((a, b) => a[0] - b[0]).forEach(([a, b]) => {
    const l = out[out.length - 1];
    if (l && a <= l[1]) l[1] = Math.max(l[1], b); else out.push([a, b]);
  });
  return out;
}
function minus_(wins, busy) {
  let out = merge_(wins);
  merge_(busy).forEach(([a, b]) => {
    out = out.reduce((acc, [x, y]) => {
      if (b <= x || a >= y) acc.push([x, y]);
      else { if (a > x) acc.push([x, a]); if (b < y) acc.push([b, y]); }
      return acc;
    }, []);
  });
  return out;
}
// ეტევა თუ არა start-დან dur წუთი ერთ თავისუფალ შუალედში
const fits_ = (wins, start, dur) => (wins || []).some(([a, b]) => start >= a && start + dur <= b);

/* ---------- პირველი გაშვება ----------
   რედაქტორში აირჩიეთ setup და დააჭირეთ Run: შექმნის ცხრილს, მოითხოვს ნებართვებს
   და გამოგიგზავნით სატესტო შეტყობინებას. */
function setup() {
  const ss = book_();
  console.log("ცხრილი: " + ss.getUrl());
  notify_("ახალი ინდივიდუალური მოთხოვნა", reqLines_({ name: "ტესტი", phone: "555123456", date: today_(), time: "19:00", people: 2, duration: 90, price: indPrice_(today_(), "19:00", 90, 2), experience: EXPERIENCE[1] }));
  console.log("სატესტო შეტყობინება გაიგზავნა" + (NOTIFY_EMAIL ? " მეილზე" : "") + (WHATSAPP_PHONE && CALLMEBOT_APIKEY ? " და WhatsApp-ში" : ""));
}
