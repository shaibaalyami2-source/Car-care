/**
 * CRM · نظام إدارة خدمات النظافة — التخزين السحابي المبدئي على Google Sheet
 * ---------------------------------------------------------------------------
 * 1) من ملف Google Sheet: الإضافات (Extensions) ← Apps Script ← الصق هذا الملف بالكامل ← حفظ.
 * 2) اختر الدالة setup من الأعلى واضغط تشغيل (Run) ووافق على الأذونات. ستُنشأ الجداول ومفتاح الوصول في تبويب «إعداد».
 * 3) نشر (Deploy) ← نشر جديد ← تطبيق ويب (Web app) ← التنفيذ باسم: أنا ← من يمكنه الوصول: أي شخص (Anyone) ← نشر، وانسخ الرابط.
 * 4) في النظام: الإعدادات ← المزامنة السحابية ← الصق الرابط والمفتاح ← ربط.
 * عند تعديل هذا الكود لاحقًا: نشر ← إدارة عمليات النشر ← تعديل ← إصدار جديد (حتى يبقى الرابط نفسه).
 */
var TABLES = ['requests', 'clients', 'staff', 'contracts', 'tenders', 'projects', 'courses', 'enrollments', 'messages', 'docs', 'roles', 'payAdj', 'locations', 'chatRead', 'guide', 'settings'];
var TITLES = { requests: 'الطلبات', clients: 'العملاء', staff: 'الموظفات', contracts: 'العقود', tenders: 'المقاولات', projects: 'المشاريع', courses: 'الدورات', enrollments: 'التدريب', messages: 'المحادثات', docs: 'المستندات', roles: 'الأدوار والصلاحيات', payAdj: 'تعديلات الرواتب', locations: 'مواقع الميدان', chatRead: 'حالة القراءة', guide: 'تقدم الدروس', settings: 'الإعدادات' };
/* readable columns so the office can read the sheet; the full record lives in the json columns */
var VIEW = {
  requests: [['العميل', 'clientName'], ['الجوال', 'clientPhone'], ['الخدمة', 'service'], ['الحي', 'district'], ['الموعد', 'prefDate'], ['الحالة', 'status'], ['المشرفة', 'supervisorId'], ['العاملة', 'workerId']],
  clients: [['الاسم', 'name'], ['الجوال', 'phone'], ['النوع', 'type'], ['العنوان', 'address']],
  staff: [['الاسم', 'name'], ['الدور', 'role'], ['الجوال', 'phone'], ['اللغة', 'lang'], ['نشطة', 'active']],
  contracts: [['العميل', 'clientId'], ['العنوان', 'title'], ['من', 'start'], ['إلى', 'end']],
  tenders: [['العنوان', 'title'], ['الجهة', 'client'], ['الحالة', 'status']],
  projects: [['الاسم', 'name'], ['الحالة', 'status']],
  messages: [['المحادثة', 'conv'], ['من', 'from'], ['النص', 'text'], ['الوقت', 'at']],
  docs: [['العنوان', 'title'], ['النوع', 'type'], ['التاريخ', 'at']],
  locations: [['خط العرض', 'lat'], ['خط الطول', 'lng'], ['الدقة م', 'acc'], ['الطلب', 'task']]
};
var CHUNK = 45000, NJSON = 4, HEAD = ['id', 'u', 'x', 's'];

function doGet() { return out({ ok: true, app: 'CRM Clean sync', time: Date.now() }); }
function doPost(e) {
  var req; try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad-json' }); }
  var a = req.action;
  try {
    if (a === 'pubRequest') return out(pubRequest(req));
    if (a === 'pubTrack') return out(pubTrack(req));
    if (a === 'pubRate') return out(pubRate(req));
    var key = PropertiesService.getScriptProperties().getProperty('KEY');
    if (!key) return out({ ok: false, error: 'setup' });
    if (String(req.key || '') !== key) return out({ ok: false, error: 'auth' });
    if (a === 'ping') return out(ping());
    if (a === 'push') return out(withLock(function () { return push(req.rows || []); }));
    if (a === 'pull') return out(pull(Number(req.since) || 0));
    if (a === 'alloc') return out(withLock(function () { return alloc(req.kind, Number(req.size) || 20, Number(req.min) || 1); }));
    if (a === 'file') return out(saveFile(req.data, req.name));
    if (a === 'translate') return out({ ok: true, text: LanguageApp.translate(String(req.text || '').slice(0, 4500), req.from || '', req.to || 'ar') });
    return out({ ok: false, error: 'unknown-action' });
  } catch (err) { return out({ ok: false, error: String(err && err.message || err) }); }
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withLock(fn) { var l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }

/* ---------- setup ---------- */
function setup() {
  var ss = SpreadsheetApp.getActive(); var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KEY')) props.setProperty('KEY', randKey());
  TABLES.forEach(function (t) { sheetFor(t); });
  var info = ss.getSheetByName('إعداد') || ss.insertSheet('إعداد', 0);
  info.clear(); info.setRightToLeft(true);
  info.getRange(1, 1, 8, 2).setValues([
    ['نظام CRM لإدارة خدمات النظافة', ''], ['مفتاح الوصول (انسخيه إلى النظام)', props.getProperty('KEY')], ['الخطوة التالية', 'نشر ← نشر جديد ← تطبيق ويب ← أي شخص ← انسخي الرابط'],
    ['الجداول', TABLES.map(function (t) { return TITLES[t]; }).join('، ')], ['تنبيه', 'لا تعدّلي الأعمدة id و u و x و s و json يدويًا؛ النظام يحدثها تلقائيًا'],
    ['الصور الكبيرة', 'تُحفظ في مجلد «CRM Clean — الملفات» في Google Drive'], ['تاريخ الإعداد', new Date()], ['الإصدار', '3.0']]);
  info.getRange('A1').setFontSize(16).setFontWeight('bold'); info.getRange('B2').setFontWeight('bold').setFontColor('#b4235f');
  info.setColumnWidth(1, 260); info.setColumnWidth(2, 520);
  var s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1'); if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  Logger.log('Access key: ' + props.getProperty('KEY'));
}
function resetKey() { PropertiesService.getScriptProperties().setProperty('KEY', randKey()); setup(); }
function onOpen() { SpreadsheetApp.getUi().createMenu('CRM').addItem('إعداد / تحديث الجداول', 'setup').addItem('تغيير مفتاح الوصول', 'resetKey').addToUi(); }
function randKey() { var c = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789', k = ''; for (var i = 0; i < 24; i++) k += c.charAt(Math.floor(Math.random() * c.length)); return k; }
function headFor(t) { var h = HEAD.slice(); (VIEW[t] || []).forEach(function (v) { h.push(v[0]); }); for (var i = 1; i <= NJSON; i++) h.push('json' + (i > 1 ? i : '')); return h; }
function sheetFor(t) {
  var ss = SpreadsheetApp.getActive(); var sh = ss.getSheetByName(t);
  if (!sh) {
    sh = ss.insertSheet(t); var h = headFor(t);
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#fde7f0');
    sh.setFrozenRows(1); sh.setRightToLeft(true); sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@');
    sh.getRange(1, 1).setNote(TITLES[t] || t);
  }
  return sh;
}

/* ---------- read / write ---------- */
function readTable(t) {
  var sh = sheetFor(t); var n = sh.getLastRow(); var h = headFor(t);
  var vals = n > 1 ? sh.getRange(2, 1, n - 1, h.length).getValues() : [];
  return { sh: sh, h: h, vals: vals };
}
function recOf(row, h) { var j = ''; for (var i = h.length - NJSON; i < h.length; i++) j += String(row[i] || ''); try { return j ? JSON.parse(j) : null; } catch (e) { return null; } }
function rowOf(t, id, u, x, s, d) {
  var row = [id, String(u), x ? '1' : '', String(s)];
  (VIEW[t] || []).forEach(function (v) { var val = d && d[v[1]]; row.push(val == null ? '' : String(typeof val === 'object' ? JSON.stringify(val) : val).slice(0, 500)); });
  var j = x || !d ? '' : JSON.stringify(d); if (j.length > CHUNK * NJSON) throw new Error('record-too-large:' + t + ':' + id);
  for (var i = 0; i < NJSON; i++) row.push(j.substr(i * CHUNK, CHUNK));
  return row;
}
function writeTable(tb) { if (!tb.vals.length) return; tb.sh.getRange(2, 1, tb.vals.length, tb.h.length).setNumberFormat('@').setValues(tb.vals); }
function push(rows) {
  var now = Date.now(); var by = {};
  rows.forEach(function (r) { if (TABLES.indexOf(r.t) >= 0) (by[r.t] = by[r.t] || []).push(r); });
  Object.keys(by).forEach(function (t) {
    var tb = readTable(t); var idx = {}; tb.vals.forEach(function (v, i) { idx[String(v[0])] = i; });
    by[t].forEach(function (r) {
      var id = String(r.id); var i = idx[id];
      if (i != null && Number(tb.vals[i][1]) > Number(r.u)) return; // the sheet already has a newer edit
      var row = rowOf(t, id, r.u, r.x, now, r.d);
      if (i != null) tb.vals[i] = row; else { idx[id] = tb.vals.length; tb.vals.push(row); }
    });
    writeTable(tb);
  });
  return { ok: true, time: now };
}
function pull(since) {
  var now = Date.now(); var rows = [];
  TABLES.forEach(function (t) {
    var tb = readTable(t);
    tb.vals.forEach(function (v) {
      if (!v[0] || Number(v[3]) < since - 5000) return;
      if (v[2] === '1' || v[2] === 1) rows.push({ t: t, id: String(v[0]), u: Number(v[1]), x: 1 });
      else { var d = recOf(v, tb.h); if (d) rows.push({ t: t, id: String(v[0]), u: Number(v[1]), d: d }); }
    });
  });
  return { ok: true, rows: rows, time: now };
}
function ping() {
  var r = readTable('requests'), s = readTable('staff');
  return { ok: true, sheet: SpreadsheetApp.getActive().getName(), empty: !r.vals.length && !s.vals.length, time: Date.now() };
}
function alloc(kind, size, min) {
  var p = PropertiesService.getScriptProperties(); var k = 'ctr_' + kind;
  var start = Math.max(Number(p.getProperty(k)) || 0, min || 1); var end = start + Math.max(1, Math.min(size, 200)) - 1;
  p.setProperty(k, String(end + 1));
  return { ok: true, start: start, end: end };
}
function saveFile(data, name) {
  var m = String(data || '').match(/^data:([^;]+);base64,(.+)$/); if (!m) return { ok: false, error: 'bad-file' };
  var it = DriveApp.getFoldersByName('CRM Clean — الملفات'); var folder = it.hasNext() ? it.next() : DriveApp.createFolder('CRM Clean — الملفات');
  var f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (name || 'file') + '.' + (m[1].split('/')[1] || 'bin')));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* some Workspace domains forbid public links */ }
  return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w2000' };
}

/* ---------- the public client portal (no key: limited, rate-checked) ---------- */
function phone9(p) { var d = String(p || '').replace(/\D/g, ''); return d.slice(-9); }
function limited(k, max) { var c = CacheService.getScriptCache(); var n = Number(c.get(k) || 0); if (n >= max) return true; c.put(k, String(n + 1), 3600); return false; }
function findRow(tb, id) { for (var i = 0; i < tb.vals.length; i++) if (String(tb.vals[i][0]).toUpperCase() === String(id).toUpperCase()) return i; return -1; }
function pubRequest(req) {
  var d = req.d || {}; var ph = phone9(d.clientPhone);
  if (!d.clientName || ph.length < 9 || !d.district) return { ok: false, error: 'missing' };
  if (limited('pr_' + ph, 6)) return { ok: false, error: 'too-many' };
  return withLock(function () {
    var now = Date.now(); var iso = new Date(now).toISOString();
    var n = alloc('req', 1, 1001).start; var id = 'REQ-' + n;
    var cl = readTable('clients'); var clientId = null;
    cl.vals.forEach(function (v) { var c = recOf(v, cl.h); if (c && phone9(c.phone) === ph) clientId = c.id; });
    if (!clientId) { clientId = 'p' + now.toString(36); cl.vals.push(rowOf('clients', clientId, now, 0, now, { id: clientId, name: String(d.clientName).slice(0, 80), type: 'individual', phone: String(d.clientPhone).slice(0, 20), address: String(d.district).slice(0, 80) })); writeTable(cl); }
    var w = Math.max(1, Math.min(10, Number(d.workers) || 1)), h = Math.max(1, Math.min(12, Number(d.hoursEst) || 3));
    var r = { id: id, createdAt: iso, source: 'portal', clientId: clientId, clientName: String(d.clientName).slice(0, 80), clientPhone: String(d.clientPhone).slice(0, 20), service: String(d.service || 'home').slice(0, 20), district: String(d.district).slice(0, 80), floor: String(d.floor || '').slice(0, 20), apartment: String(d.apartment || '').slice(0, 40), gps: String(d.gps || '').slice(0, 300), prefDate: String(d.prefDate || iso.slice(0, 10)).slice(0, 10), period: String(d.period || 'morning').slice(0, 12), workers: w, hoursEst: h, notes: String(d.notes || '').slice(0, 800), priority: 'normal', contractId: null, status: 'new', supervisorId: null, workerId: null, events: [{ t: 'created', at: iso }], estimate: Math.round(w * h * (Number(d.rate) || 45) * 1.15) };
    var rq = readTable('requests'); rq.vals.push(rowOf('requests', id, now, 0, now, r)); writeTable(rq);
    return { ok: true, id: id };
  });
}
function pubTrack(req) {
  if (limited('pt_' + phone9(req.phone), 120)) return { ok: false, error: 'too-many' };
  var rq = readTable('requests'); var i = findRow(rq, req.id); if (i < 0) return { ok: false, error: 'not-found' };
  var r = recOf(rq.vals[i], rq.h); if (!r || phone9(r.clientPhone) !== phone9(req.phone)) return { ok: false, error: 'not-found' };
  var st = readTable('staff'); var sup = null, wk = null;
  st.vals.forEach(function (v) { var s = recOf(v, st.h); if (!s) return; if (s.id === r.supervisorId) sup = s; if (s.id === r.workerId) wk = s; });
  var out = { id: r.id, service: r.service, status: r.status, createdAt: r.createdAt, supAt: r.supAt, workerAt: r.workerAt, startedAt: r.startedAt, approvedAt: r.approvedAt, cancelledAt: r.cancelledAt, prefDate: r.prefDate, period: r.period, district: r.district, gps: r.gps, onwayAt: r.onwayAt, arrivedAt: r.arrivedAt, rating: r.rating || null, supName: sup ? String(sup.name).split(' ')[0] : '', workerId: r.workerId ? 'w' : null, photo: r.status === 'done' && r.photo && String(r.photo).indexOf('http') === 0 ? r.photo : null };
  if (r.status === 'worker' && r.onwayAt && !r.arrivedAt && wk) {
    var lc = readTable('locations'); var j = findRow(lc, wk.id);
    if (j >= 0) { var l = recOf(lc.vals[j], lc.h); if (l && Date.now() - l.at < 20 * 60000) out.live = { lat: Math.round(l.lat * 1000) / 1000, lng: Math.round(l.lng * 1000) / 1000, at: l.at, spd: l.spd || null }; }
  }
  return { ok: true, r: out };
}
function pubRate(req) {
  var stars = Math.max(1, Math.min(5, Number(req.stars) || 0)); if (!stars) return { ok: false, error: 'missing' };
  return withLock(function () {
    var rq = readTable('requests'); var i = findRow(rq, req.id); if (i < 0) return { ok: false, error: 'not-found' };
    var r = recOf(rq.vals[i], rq.h); if (!r || phone9(r.clientPhone) !== phone9(req.phone) || r.status !== 'done' || r.rating) return { ok: false, error: 'not-allowed' };
    var now = Date.now(); r.rating = { stars: stars, comment: String(req.comment || '').slice(0, 600), at: new Date(now).toISOString() };
    (r.events = r.events || []).push({ t: 'rate', n: stars, at: r.rating.at });
    rq.vals[i] = rowOf('requests', r.id, now, 0, now, r); writeTable(rq);
    return { ok: true };
  });
}
