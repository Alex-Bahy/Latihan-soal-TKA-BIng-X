/**
 * Backend Ujian TKA v2 - Google Apps Script.
 * Perbaikan: lock singkat, SID stabil (bisa lanjut dari perangkat lain), jalur GET/JSONP cadangan,
 * cache config, pencarian sesi cepat, menu pemulihan untuk guru.
 */
const SS_ID = '1Jzo_c5cjPouJ8DB9LDf8nTpu2dpe3ke1lwZqz7S__bc';
const IMG_INFO = 'https://iili.io/nXwyNiG.md.png';
const L = 'ABCDE', BS = ['Benar', 'Salah'], ST = ['Sesuai', 'Tidak Sesuai'];

const BANK = [
 {s:null,q:[
  [1,'A',"Rina says, \"Hello, my name is Rina. What's your name?\" What is the best response from Doni?",["I'm fine, thank you.","Hi, I'm Doni. Nice to meet you.","See you later!","Good night!"],'B'],
  [2,'A',"Which expression is used to introduce a friend to another person?",["Let me introduce myself.","This is my friend, Doni.","How old are you?","Where do you live?"],'B'],
  [3,'A',"Information: Doni is from Surabaya. Which question can be made from the information?",["Where is Doni from?","Who is Doni?","How old is Doni?","What is Doni's major?"],'A'],
  [4,'A',"Information: Sinta is sixteen years old. Which question is correct?",["Who is Sinta?","Where is Sinta from?","How old is Sinta?","Is Sinta a student?"],'C'],
  [5,'A',"Sinta ____ a student at SMK Negeri 1 Maluku Tengah.",["am","is","are","be"],'B'],
  [6,'A',"Rina and Doni ____ students of Accounting major.",["is","am","are","be"],'C'],
  [7,'A',"\"Her name is Sinta.\" The word \"Her\" refers to ....",["a male person","a female person","a group of people","the speaker"],'B'],
  [8,'A',"Which expression is used to say goodbye?",["Good morning!","Nice to meet you.","See you later!","What's your name?"],'C'],
  [9,'A',"Leni dreams of becoming a member of the Women Army Corps (Kowad). Which description fits her dream best?",["Caring with heart, helping with love.","Brave, strong, and dedicated to serve the nation.","Inspiring through knowledge and technology.","Creating opportunities, building the future."],'B',IMG_INFO],
  [10,'A',"Regina wants to be a lawyer. According to the infographic, what will she do?",["Educate students about science.","Care for patients in the hospital.","Uphold justice and defend the truth.","Serve God and guide others."],'C',IMG_INFO]]},
 {s:null,q:[
  [11,'B',"Which sentences are used to introduce yourself? (Choose 3 answers)",["My name is Rina.","This is my friend, Doni.","I am from Surabaya.","I live in Maluku Tengah.","Her name is Sinta."],['A','C','D']],
  [12,'B',"Which question words are matched with their correct use? (Choose 3 answers)",["WHO \u2013 asks about a person","WHAT \u2013 asks about a place","WHERE \u2013 asks about a place","HOW OLD \u2013 asks about age","IS/ARE \u2013 asks about a person's job only"],['A','C','D']],
  [13,'B',"Which sentences use the correct form of to be? (Choose 3 answers)",["I am a student.","She are a student.","We are students.","He is from Surabaya.","They is friends."],['A','C','D']],
  [14,'B',"Which dreams are connected with defending the nation? (Choose 3 answers)",["Women Army Corps (Kowad)","Soldier","Military Academy cadet","Entrepreneur","Informatics Teacher"],['A','B','C']],
  [15,'B',"Which words are possessive adjectives? (Choose 3 answers)",["My","They","Their","Our","We"],['A','C','D']]]},
 {s:"Rina: Hello, my name is Rina. What's your name?\nDoni: Hi, I'm Doni. Nice to meet you.\nRina: Nice to meet you, too. What is your major?\nDoni: My major is Accounting. How about you?\nRina: I'm an Accounting student, too.",q:[
  [16,'C',"Read the dialogue. Tentukan Benar atau Salah.",BS,[["Doni's major is Accounting.","Benar"],["Rina studies Computer and Network Engineering.","Salah"],["Rina and Doni are both Accounting students.","Benar"],["Doni asks Rina's name first.","Salah"]]]]},
 {s:"Hello, my name is Rina. I am a student of SMK Negeri 1 Maluku Tengah. This is my friend, Doni. He is from Surabaya. We are students of Accounting major. I also have another friend. Her name is Sinta and she is sixteen years old. She studies at the same school with me.",q:[
  [17,'C',"Read the text. Tentukan Benar atau Salah.",BS,[["Rina studies at SMK Negeri 1 Maluku Tengah.","Benar"],["Doni is from Maluku Tengah.","Salah"],["Sinta is sixteen years old.","Benar"],["Sinta studies at a different school from Rina.","Salah"]]]]},
 {s:"Question: Is Rina a student of SMK Negeri 1 Maluku Tengah?",q:[
  [18,'C',"Tentukan Sesuai atau Tidak Sesuai.",ST,[["The question is a yes/no question.","Sesuai"],["\"Yes, she is.\" is an appropriate answer.","Sesuai"],["\"He is from Surabaya.\" is an appropriate answer.","Tidak Sesuai"],["The pattern is Is/Are + Subject + Information?","Sesuai"]]]]},
 {s:null,q:[
  [19,'C',"Match the dream with the person based on the infographic Our Dreams, Our Future. Tentukan Sesuai atau Tidak Sesuai.",ST,[["Marvin dreams of becoming an Informatics Teacher.","Sesuai"],["Kelvin dreams of becoming a soldier.","Tidak Sesuai"],["Elleisye dreams of becoming a nurse / BMKG officer.","Sesuai"],["Sintia dreams of becoming a nurse.","Tidak Sesuai"]],IMG_INFO]]},
 {s:null,q:[
  [20,'C',"Decide whether the meanings of the pronouns and possessive adjectives are correct. Tentukan Benar atau Salah.",BS,[["He = dia (laki-laki)","Benar"],["They = kami/kita","Salah"],["Her = milik dia (perempuan)","Benar"],["Our = milik mereka","Salah"]]]]}
];
const TOTAL = 30;

// ---------- Menu & setup ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('TKA')
    .addItem('Setup Awal', 'setup')
    .addItem('Reset Sesi Siswa (izinkan ulang)', 'resetSesi')
    .addSeparator()
    .addItem('Pulihkan / Kumpulkan Paksa (siswa gagal kirim)', 'pulihkan')
    .addItem('Input Jawaban dari Kode Cadangan', 'inputManual')
    .addItem('Rekap Otomatis', 'rekap').addToUi();
}

function setup() {
  const ss = SS();
  const H = {
    Hasil: ['Mulai','Selesai','Durasi (menit)','Nama','NIS','Kelas','Skor A','Skor B','Skor C','Skor Total (maks 30)','Nilai','Pelanggaran','Status','User Agent'],
    Jawaban_Detail: ['Timestamp','NIS','Nama','No Soal Asli','Jawaban Siswa','Kunci','Poin'],
    Log_Pelanggaran: ['Timestamp','NIS','Nama','Jenis','Pelanggaran ke-'],
    Sesi: ['NIS','Nama','Kelas','Mulai (ms)','Status','Jawaban','Pelanggaran','SID','User Agent','Skor'],
    Config: ['KEY','VALUE']
  };
  Object.keys(H).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getLastRow() === 0) { s.getRange(1,1,1,H[n].length).setValues([H[n]]).setFontWeight('bold'); s.setFrozenRows(1); }
  });
  ss.getSheetByName('Sesi').getRange('A:A').setNumberFormat('@');
  ss.getSheetByName('Hasil').getRange('E:E').setNumberFormat('@');
  ss.getSheetByName('Jawaban_Detail').getRange('B:B').setNumberFormat('@');
  ss.getSheetByName('Log_Pelanggaran').getRange('B:B').setNumberFormat('@');
  const c = ss.getSheetByName('Config');
  if (c.getLastRow() < 2) c.getRange(2,1,6,2).setValues([['TOKEN','TKA2026'],['DURASI_MENIT',45],['MAKS_PELANGGARAN',3],['NAVIGASI_BEBAS',true],['TAMPILKAN_NILAI',true],['UJIAN_AKTIF',true]]);
  CacheService.getScriptCache().remove('cfg');
  rekap();
}

function rekap() {
  const s = SS().getSheetByName('Rekap') || SS().insertSheet('Rekap');
  s.clear();
  s.getRange('A1').setFormula('=IFERROR(QUERY(Hasil!A:N,"select F, count(D), avg(K), max(K), min(K) where F is not null group by F label F \'Kelas\', count(D) \'Jumlah\', avg(K) \'Rata-rata\', max(K) \'Tertinggi\', min(K) \'Terendah\'",1),"Belum ada data")');
}

function resetSesi() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Reset sesi', 'Masukkan NIS siswa:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const o = sesi(clean(r.getResponseText()));
  if (!o.v) return ui.alert('NIS tidak ditemukan di sheet Sesi.');
  o.s.deleteRow(o.r);
  ui.alert('Sesi direset. Siswa bisa mengulang (data lama di sheet Hasil tetap ada).');
}

// Kumpulkan paksa sesi AKTIF yang waktunya sudah habis (memakai jawaban autosave terakhir)
// dan lengkapi Hasil untuk sesi yang sudah SELESAI tetapi barisnya belum masuk Hasil.
function pulihkan() {
  const ui = SpreadsheetApp.getUi(), c = cfg(), s = sh('Sesi'), v = s.getDataRange().getValues();
  const ada = {}; sh('Hasil').getDataRange().getValues().slice(1).forEach(r => ada[nk(r[4])] = 1);
  const batas = Number(c.DURASI_MENIT) * 60000 + 60000, now = Date.now(); let a = 0, b = 0;
  for (let i = 1; i < v.length; i++) {
    const x = v[i]; if (!nk(x[0])) continue;
    const o = { r: i + 1, s, v: x }, j = safeJson(x[5]);
    if (x[4] === 'AKTIF') { if (now > Number(x[3]) + batas) { catat(o, j, 'DIKUMPULKAN GURU (waktu habis)'); a++; } }
    else if (!ada[nk(x[0])]) { catat(o, j, 'DIPULIHKAN', true); b++; }
  }
  ui.alert('Selesai.\nDikumpulkan paksa: ' + a + ' siswa\nHasil dipulihkan: ' + b + ' siswa');
}

// Tempel "kode cadangan" dari layar siswa yang gagal kirim.
function inputManual() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Input kode cadangan', 'Tempel kode cadangan siswa:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    const k = JSON.parse(Utilities.newBlob(Utilities.base64Decode(r.getResponseText().trim())).getDataAsString());
    const o = sesi(clean(k.n));
    if (!o.v) return ui.alert('NIS tidak ditemukan di sheet Sesi.');
    if (o.v[4] !== 'AKTIF') return ui.alert('Siswa ini sudah tercatat selesai.');
    catat(o, k.j || {}, 'DIINPUT GURU');
    ui.alert('Jawaban ' + o.v[1] + ' berhasil dicatat.');
  } catch (e) { ui.alert('Kode tidak valid.'); }
}

// ---------- Utilitas ----------
let _ss; const SS = () => _ss || (_ss = SpreadsheetApp.openById(SS_ID));
function sh(n) {
  let s = SS().getSheetByName(n);
  if (!s) { setup(); s = SS().getSheetByName(n); }
  return s;
}
const bool = v => String(v).toUpperCase() === 'TRUE';
const fail = m => ({ ok: false, msg: m });
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const nk = s => String(s == null ? '' : s).trim().replace(/^0+(?=.)/, '');
const safeJson = s => { try { return JSON.parse(s || '{}'); } catch (e) { return {}; } };
function clean(s) { s = String(s == null ? '' : s).trim().slice(0, 100); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function cfg() {
  const ca = CacheService.getScriptCache(), h = ca.get('cfg');
  if (h) return JSON.parse(h);
  const c = {}; sh('Config').getDataRange().getValues().slice(1).forEach(r => c[r[0]] = r[1]);
  ca.put('cfg', JSON.stringify(c), 30);
  return c;
}
function sesi(nis) { // baca hanya kolom NIS, lalu satu baris
  const s = sh('Sesi'), n = s.getLastRow(), k = nk(nis);
  if (n > 1) {
    const a = s.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < a.length; i++) if (nk(a[i][0]) === k) return { r: i + 2, s, v: s.getRange(i + 2, 1, 1, 10).getValues()[0] };
  }
  return { s };
}
function chk(d) {
  const o = sesi(clean(d.nis));
  if (!o.v || String(o.v[7]) !== String(d.sid)) return { err: fail('Sesi tidak valid. Muat ulang halaman lalu login lagi dengan data yang sama.') };
  if (o.v[4] !== 'AKTIF') return { err: { ok: false, selesai: true, msg: 'Ujian sudah dikumpulkan.' } };
  return o;
}
function soalKlien() {
  return BANK.map(g => ({ s: g.s, q: g.q.map(x => x[1] === 'C'
    ? { no: x[0], tipe: 'C', img: x[5], teks: x[2], label: x[3], p: x[4].map(p => p[0]) }
    : { no: x[0], tipe: x[1], img: x[5], teks: x[2], n: x[1] === 'B' ? x[4].length : undefined, opsi: Object.fromEntries(x[3].map((o, i) => [L[i], o])) }) }));
}
function nilai(j) {
  const d = []; let a = 0, b = 0, c = 0;
  BANK.forEach(g => g.q.forEach(x => {
    const no = x[0], t = x[1], u = j[no]; let p = 0, key, ans;
    if (t === 'A') { key = x[4]; ans = u || ''; p = ans === key ? 1 : 0; a += p; }
    else if (t === 'B') {
      const s = Array.isArray(u) ? u.filter((v, i, r) => r.indexOf(v) === i).sort() : [];
      key = x[4].join(','); ans = s.join(',');
      if (s.length && s.every(v => x[4].includes(v))) p = s.length === x[4].length ? 2 : 1;
      b += p;
    } else {
      const r = Array.isArray(u) ? u : [];
      key = x[4].map(q => q[1]).join(' | '); ans = x[4].map((q, i) => r[i] || '-').join(' | ');
      const k = x[4].filter((q, i) => r[i] === q[1]).length;
      p = k === x[4].length ? 2 : (k === x[4].length - 1 ? 1 : 0); c += p;
    }
    d.push([no, ans, key, p]);
  }));
  return { a, b, c, d };
}
const hasil = (tampil, skor) => ({ ok: true, tampil, skor, nilai: Math.round(skor / TOTAL * 10000) / 100, total: TOTAL });
function kunci(fn) { const k = LockService.getScriptLock(); k.waitLock(30000); try { return fn(); } finally { k.releaseLock(); } }

// ---------- API ----------
function route(d) {
  const F = { ping: () => ({ ok: true }), mulai, simpan, log: logP, submit };
  if (!F[d.action]) return fail('Aksi tidak dikenal.');
  try { return F[d.action](d); }
  catch (err) { return { ok: false, retry: true, msg: 'Server sedang sibuk (' + err.message + ')' }; }
}
function doPost(e) {
  let d; try { d = JSON.parse(e.postData.contents); } catch (x) { return out(fail('Data tidak valid.')); }
  return out(route(d));
}
// GET: status API, atau jalur cadangan JSONP (untuk browser/jaringan yang gagal memakai POST)
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) return out({ ok: true, status: 'API ujian aktif' });
  let d = {}; try { d = JSON.parse(p.d || '{}'); } catch (x) {}
  d.action = p.action;
  const t = JSON.stringify(route(d));
  if (!p.cb) return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.JSON);
  return ContentService.createTextOutput(String(p.cb).replace(/[^\w.$]/g, '') + '(' + t + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function mulai(d) {
  const c = cfg();
  if (!bool(c.UJIAN_AKTIF)) return fail('Ujian belum dibuka.');
  if (String(d.token).trim().toUpperCase() !== String(c.TOKEN).trim().toUpperCase()) return fail('Token salah.');
  const nis = clean(d.nis), nama = clean(d.nama);
  if (!nis || !nama) return fail('Nama dan NIS wajib diisi.');
  return kunci(() => {
    let o = sesi(nis), sid;
    if (!o.v) {
      sid = Utilities.getUuid();
      const r = o.s.getLastRow() + 1;
      o.s.getRange(r, 1).setNumberFormat('@');
      o.s.getRange(r, 1, 1, 10).setValues([[nis, nama, clean(d.kelas), Date.now(), 'AKTIF', '{}', 0, sid, clean(d.ua), '']]);
      o = sesi(nis);
    } else if (o.v[4] !== 'AKTIF') return { ok: false, selesai: true, msg: 'Anda sudah mengerjakan ujian ini.' };
    else { sid = String(o.v[7] || '') || Utilities.getUuid(); if (!o.v[7]) o.s.getRange(o.r, 8).setValue(sid); } // SID tetap sama -> bisa lanjut di perangkat lain
    return { ok: true, sid, now: Date.now(), mulai: Number(o.v[3]), durasi: Number(c.DURASI_MENIT), nav: bool(c.NAVIGASI_BEBAS),
      maks: Number(c.MAKS_PELANGGARAN), soal: soalKlien(), jawaban: safeJson(o.v[5]), pel: Number(o.v[6]) };
  });
}

function simpan(d) { // tanpa lock: ringan & sering dipanggil
  const o = chk(d); if (o.err) return o.err;
  o.s.getRange(o.r, 6).setValue(JSON.stringify(d.jawaban || {}));
  return { ok: true };
}

function logP(d) {
  return kunci(() => {
    const o = chk(d); if (o.err) return o.err;
    const n = Number(o.v[6]) + 1;
    o.s.getRange(o.r, 7).setValue(n);
    sh('Log_Pelanggaran').appendRow([new Date(), o.v[0], o.v[1], clean(d.jenis), n]);
    return { ok: true, pel: n };
  });
}

function submit(d) {
  const tampil = bool(cfg().TAMPILKAN_NILAI);
  return kunci(() => {
    const o = sesi(clean(d.nis));
    if (!o.v || String(o.v[7]) !== String(d.sid)) return fail('Sesi tidak valid. Muat ulang, login lagi, lalu kumpulkan. Jika tetap gagal, serahkan kode cadangan ke guru.');
    if (o.v[4] !== 'AKTIF') return hasil(tampil, Number(o.v[9]) || 0); // idempotent
    const lewat = Date.now() > Number(o.v[3]) + Number(cfg().DURASI_MENIT) * 60000 + 60000;
    const st = d.auto === 'pelanggaran' ? 'AUTO-SUBMIT PELANGGARAN' : (lewat ? 'TERLAMBAT' : 'SELESAI');
    return hasil(tampil, catat(o, d.jawaban || {}, st));
  });
}

// Tulis hasil. Status Sesi ditulis lebih dulu agar jawaban pasti aman; jika Hasil/Detail gagal, jalankan menu "Pulihkan".
function catat(o, j, st, sudah) {
  const r = nilai(j), skor = r.a + r.b + r.c, nl = Math.round(skor / TOTAL * 10000) / 100;
  const now = Date.now(), mulai = Number(o.v[3]);
  if (!sudah) o.s.getRange(o.r, 5, 1, 6).setValues([['SELESAI', JSON.stringify(j), Number(o.v[6]), o.v[7], o.v[8] || '', skor]]);
  try {
    sh('Hasil').appendRow([new Date(mulai), new Date(now), Math.round((now - mulai) / 600) / 100, o.v[1], o.v[0], o.v[2], r.a, r.b, r.c, skor, nl, Number(o.v[6]), st, o.v[8] || '']);
    const det = sh('Jawaban_Detail'), ts = new Date();
    const rows = r.d.map(x => [ts, o.v[0], o.v[1], x[0], x[1], x[2], x[3]]);
    det.getRange(det.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  } catch (e) {}
  return skor;
}
