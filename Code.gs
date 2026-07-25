/**
 * ═══════════════════════════════════════════════════════════════════════
 *  นักสืบไฮโดรคาร์บอน (Hydrocarbon Sleuth Lab) — Google Apps Script
 *  รับข้อมูลจากเกมแบบเป็นชุด (batch) แล้วแยกลงชีตตามประเภทเหตุการณ์
 *
 *  วิธีติดตั้ง : ดูไฟล์ "วิธีติดตั้ง Google Sheet.md"
 *  สรุปย่อ    : เปิด Google Sheet > ส่วนขยาย > Apps Script > วางไฟล์นี้ทั้งหมด
 *               > Deploy > New deployment > Web app
 *               > Execute as: Me | Who has access: Anyone > คัดลอก URL /exec
 *               > นำ URL ไปวางในบล็อก CFG บนหัวไฟล์เกม
 * ═══════════════════════════════════════════════════════════════════════
 */

/* ต้องตรงกับค่า TOKEN ในบล็อก CFG ของไฟล์เกม */
var TOKEN = 'PTK-CHEM-2569';

var SH_LOG = 'log';        // ทุกเหตุการณ์ รายละเอียดระดับข้อ
var SH_SUM = 'summary';    // สรุปรายคน (อัปเดตทับแถวเดิม)

var H_LOG = ['เวลา', 'ประเภท', 'ชื่อ', 'ห้อง', 'เลขที่', 'XP',
             'สาร', 'สูตร', 'การทดสอบ', 'O2', 'เกิดปฏิกิริยา', 'เขม่า%', 'วินาที',
             'โหมด', 'ตอบถูก', 'จำนวนทดสอบ', 'คะแนน/10', 'ภารกิจ', 'ดาว', 'เฉลี่ย/10'];

var H_SUM = ['อัปเดตล่าสุด', 'ห้อง', 'เลขที่', 'ชื่อ', 'XP', 'ยศ',
             'ภารกิจสำเร็จ', 'ดาวรวม', 'ข้อนักสืบที่ทำ', 'ตอบถูก', 'เฉลี่ย/10', 'สถิติดีที่สุด/10'];

/* ─────────────────────────── ตัวช่วย ─────────────────────────── */

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#d9e2f3');
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function val_(r, k) { return (r[k] === undefined || r[k] === null) ? '' : r[k]; }

/* ─────────────────────── รับข้อมูลจากเกม ─────────────────────── */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // กันการเขียนชนกันเมื่อนักเรียนหลายคนส่งพร้อมกัน
    lock.waitLock(20000);

    var body = JSON.parse(e.postData.contents);
    if (TOKEN && body.token !== TOKEN) {
      return json_({ status: 'error', message: 'invalid token' });
    }

    var rows = body.rows || [];
    if (!rows.length) return json_({ status: 'ok', written: 0 });

    // 1) เขียนลงชีต log ทีเดียวทั้งชุด (เร็วกว่าเขียนทีละแถวมาก)
    var shLog = sheet_(SH_LOG, H_LOG);
    var out = rows.map(function (r) {
      return [
        r.ts ? new Date(r.ts) : new Date(),
        val_(r, 'kind'), val_(r, 'name'), val_(r, 'room'), val_(r, 'no'), val_(r, 'xp'),
        val_(r, 'mol'), val_(r, 'formula'), val_(r, 'test'), val_(r, 'o2'),
        val_(r, 'reacts'), val_(r, 'soot'), val_(r, 'secs'),
        val_(r, 'mode'), val_(r, 'correct'), val_(r, 'tests'), val_(r, 'score10'),
        val_(r, 'mission'), val_(r, 'stars'), val_(r, 'avg10')
      ];
    });
    shLog.getRange(shLog.getLastRow() + 1, 1, out.length, H_LOG.length).setValues(out);

    // 2) อัปเดตชีต summary ให้เหลือคนละ 1 แถว (ใช้ ห้อง+เลขที่ เป็นกุญแจ)
    // เฉพาะเหตุการณ์ 'summary' เท่านั้น เพราะเป็นชนิดเดียวที่มีข้อมูลครบทุกช่อง
    // (ถ้าใช้ชนิดอื่นด้วย ช่องที่ไม่มีข้อมูลจะไปทับของเดิมให้กลายเป็นค่าว่าง)
    rows.forEach(function (r) {
      if (r.kind === 'summary') upsertSummary_(r);
    });

    return json_({ status: 'ok', written: out.length });
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function upsertSummary_(r) {
  var sh = sheet_(SH_SUM, H_SUM);
  var key = String(val_(r, 'room')) + '|' + String(val_(r, 'no'));
  var last = sh.getLastRow();
  var rowIdx = -1;

  if (last > 1) {
    var keys = sh.getRange(2, 2, last - 1, 2).getValues(); // คอลัมน์ ห้อง, เลขที่
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) + '|' + String(keys[i][1]) === key) { rowIdx = i + 2; break; }
    }
  }

  var line = [
    new Date(), val_(r, 'room'), val_(r, 'no'), val_(r, 'name'),
    val_(r, 'xp'), val_(r, 'rank'),
    val_(r, 'mission'), val_(r, 'stars'),
    val_(r, 'tests'), val_(r, 'correct'), val_(r, 'avg10'), val_(r, 'score10')
  ];

  if (rowIdx === -1) sh.appendRow(line);
  else sh.getRange(rowIdx, 1, 1, line.length).setValues([line]);
}

function doGet() {
  return json_({ status: 'ok', message: 'Hydrocarbon Sleuth Lab endpoint พร้อมใช้งาน' });
}

/* ───────────────── เมนูสำหรับครู (เปิดในหน้าชีต) ───────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 นักสืบไฮโดรคาร์บอน')
    .addItem('สรุปคะแนนรายคน', 'reportByStudent')
    .addItem('วิเคราะห์รายสาร (จุดที่นักเรียนพลาด)', 'reportByMolecule')
    .addSeparator()
    .addItem('ล้างข้อมูลทั้งหมด', 'clearAll')
    .showUi();
}

/** สรุปคะแนนรายคน เรียงตามห้องและเลขที่ */
function reportByStudent() {
  var rows = readLog_();
  var map = {};

  rows.forEach(function (r) {
    var key = r['ห้อง'] + '|' + r['เลขที่'];
    if (!map[key]) {
      map[key] = { room: r['ห้อง'], no: r['เลขที่'], name: r['ชื่อ'], xp: 0,
                   tests: 0, cases: 0, correct: 0, score: 0, missions: {} };
    }
    var m = map[key];
    m.name = r['ชื่อ'] || m.name;
    m.xp = Math.max(m.xp, Number(r['XP']) || 0);
    if (r['ประเภท'] === 'test') m.tests++;
    if (r['ประเภท'] === 'case') {
      m.cases++;
      if (Number(r['ตอบถูก']) === 1) m.correct++;
      m.score += Number(r['คะแนน/10']) || 0;
    }
    if (r['ประเภท'] === 'mission' && r['ภารกิจ']) m.missions[r['ภารกิจ']] = 1;
  });

  var out = [['ห้อง', 'เลขที่', 'ชื่อ', 'XP', 'ภารกิจสำเร็จ',
              'จำนวนการทดสอบ', 'ข้อนักสืบ', 'ตอบถูก', 'ร้อยละถูก', 'คะแนนเฉลี่ย/10']];

  Object.keys(map).forEach(function (k) {
    var m = map[k];
    out.push([m.room, m.no, m.name, m.xp, Object.keys(m.missions).length,
      m.tests, m.cases, m.correct,
      m.cases ? Math.round(m.correct / m.cases * 100) : '',
      m.cases ? Math.round(m.score / m.cases * 10) / 10 : '']);
  });

  var head = out.shift();
  out.sort(function (a, b) {
    if (String(a[0]) !== String(b[0])) return String(a[0]) < String(b[0]) ? -1 : 1;
    return (Number(a[1]) || 0) - (Number(b[1]) || 0);
  });
  out.unshift(head);
  writeSheet_('สรุปรายคน', out);
}

/**
 * วิเคราะห์รายสาร — เรียงจากสารที่นักเรียนระบุผิดบ่อยที่สุดขึ้นก่อน
 * เพื่อให้ครูเห็นทันทีว่าควรสอนซ่อมมโนทัศน์ตัวไหน
 */
function reportByMolecule() {
  var rows = readLog_();
  var map = {};

  rows.forEach(function (r) {
    if (r['ประเภท'] !== 'case') return;
    var id = r['สาร'];
    if (!id) return;
    if (!map[id]) map[id] = { id: id, formula: r['สูตร'], n: 0, ok: 0, tests: 0 };
    map[id].n++;
    if (Number(r['ตอบถูก']) === 1) map[id].ok++;
    map[id].tests += Number(r['จำนวนทดสอบ']) || 0;
  });

  var out = [['สาร', 'สูตร', 'จำนวนครั้งที่เจอ', 'ตอบถูก', 'ร้อยละถูก',
              'ทดสอบเฉลี่ย (ครั้ง/ข้อ)', 'ระดับความยาก']];

  Object.keys(map).forEach(function (k) {
    var m = map[k];
    var p = m.n ? m.ok / m.n : 0;
    var label = p >= 0.8 ? 'ง่าย'
              : p >= 0.5 ? 'ปานกลาง'
              : p >= 0.3 ? 'ยาก'
              : 'ยากมาก — ควรสอนซ่อม';
    out.push([m.id, m.formula, m.n, m.ok, Math.round(p * 100),
      m.n ? Math.round(m.tests / m.n * 10) / 10 : '', label]);
  });

  var head = out.shift();
  out.sort(function (a, b) { return a[4] - b[4]; }); // ยากที่สุดขึ้นก่อน
  out.unshift(head);
  writeSheet_('วิเคราะห์รายสาร', out);
}

function readLog_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var head = values[0];
  return values.slice(1).map(function (row) {
    var o = {};
    head.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
}

function writeSheet_(name, out) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  if (!out.length) return;
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(1, 1, 1, out[0].length).setFontWeight('bold').setBackground('#d9e2f3');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, out[0].length);
  ss.setActiveSheet(sh);
}

function clearAll() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('ยืนยันล้างข้อมูล',
      'จะลบข้อมูลทุกแถวในชีต log และ summary (เก็บหัวตารางไว้) ดำเนินการต่อหรือไม่',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  [[SH_LOG, H_LOG], [SH_SUM, H_SUM]].forEach(function (p) {
    var sh = sheet_(p[0], p[1]);
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });
  ui.alert('ล้างข้อมูลเรียบร้อย');
}
