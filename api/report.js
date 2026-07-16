// Vercel Serverless Function - 구글 시트를 읽고 쓰는 작은 서버
// 필요한 환경변수(Vercel Settings > Environment Variables):
//   GOOGLE_SERVICE_ACCOUNT_KEY  = 서비스 계정 JSON 키 파일 "전체 내용"
//   DB_SHEET_ID                 = 주간업무보고_DB 스프레드시트 ID
//   ASSET_SHEET_ID              = 외부 자산시트 ID (선택; 없으면 비품 불러오기만 비활성)
const { google } = require('googleapis');

const DB = process.env.DB_SHEET_ID;
const ASSET = process.env.ASSET_SHEET_ID || '';
const ASSET_GID = 2060678306;

const TAB = {
  team:      { name: '1_팀원업무', w: 8 },
  facility:  { name: '2_시설',     w: 7 },
  assets:    { name: '3_비품',     w: 7 },
  contracts: { name: '4_계약',     w: 5 },
  notes:     { name: '5_특이사항', w: 6 }
};
const WEEKS = '설정', WEEKS_W = 8, MEMBERS = 'M_담당자';

function sheetsClient() {
  var raw = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!raw) throw new Error('환경변수 GOOGLE_SERVICE_ACCOUNT_KEY 가 없습니다.');
  var key = JSON.parse(raw);
  var pk = key.private_key.indexOf('\\n') > -1 ? key.private_key.replace(/\\n/g, '\n') : key.private_key;
  var auth = new google.auth.JWT(key.client_email, null, pk, ['https://www.googleapis.com/auth/spreadsheets']);
  return google.sheets({ version: 'v4', auth });
}

module.exports = async (req, res) => {
  try {
    if (req.method !== 'POST') { res.status(200).json({ ok: false, error: 'POST 요청만 지원합니다.' }); return; }
    var body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
    var action = body.action, args = body.args || [];
    var sheets = sheetsClient();
    var result = await dispatch(sheets, action, args);
    res.status(200).json({ ok: true, result: result });
  } catch (e) {
    res.status(200).json({ ok: false, error: String(e && e.message ? e.message : e) });
  }
};

async function dispatch(sheets, action, args) {
  if (action === 'getReportData') return getReportData(sheets);
  if (action === 'saveWeek') return saveWeek(sheets, JSON.parse(args[0]));
  if (action === 'deleteWeek') return deleteWeek(sheets, args[0]);
  if (action === 'renameWeek') return renameWeek(sheets, args[0], args[1]);
  if (action === 'importAssets') return importAssets(sheets, args[0], args[1]);
  throw new Error('알 수 없는 action: ' + action);
}

/* ---------- 공통 ---------- */
async function vget(sheets, id, range) { var r = await sheets.spreadsheets.values.get({ spreadsheetId: id, range: range }); return r.data.values || []; }
function body_(vals) { return (vals || []).slice(1).map(function (r) { return r.map(function (c) { return c == null ? '' : String(c); }); }); }
function pad(r, w) { r = r.slice(); while (r.length < w) r.push(''); return r.slice(0, w); }
async function writeTab(sheets, name, rows, w) {
  await sheets.spreadsheets.values.clear({ spreadsheetId: DB, range: name + '!A2:Z' });
  if (rows && rows.length) {
    var padded = rows.map(function (r) { return pad(r, w); });
    await sheets.spreadsheets.values.update({ spreadsheetId: DB, range: name + '!A2', valueInputOption: 'RAW', requestBody: { values: padded } });
  }
}
function tabRows(vals, w) {
  return body_(vals).map(function (r) { return pad(r, w); }).filter(function (r) {
    var empty = true; for (var c = 1; c < w; c++) if (String(r[c]).trim()) empty = false; return !empty && String(r[0]).trim();
  });
}

/* ---------- 조회 ---------- */
async function getReportData(sheets) {
  var q = await Promise.all([
    vget(sheets, DB, WEEKS), vget(sheets, DB, MEMBERS),
    vget(sheets, DB, TAB.team.name), vget(sheets, DB, TAB.facility.name),
    vget(sheets, DB, TAB.assets.name), vget(sheets, DB, TAB.contracts.name), vget(sheets, DB, TAB.notes.name)
  ]);
  var wk = q[0], mem = q[1];
  return {
    dbReady: true,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + DB + '/edit',
    weeks: body_(wk).filter(function (r) { return String(r[0] || '').trim(); }).map(function (r) {
      return { week: r[0], start: r[1] || '', end: r[2] || '', dept: r[3] || '자산시설관리팀', status: (r[4] || 'warning'), statusLabel: r[5] || '', date: r[6] || '', updated: r[7] || '' };
    }),
    members: body_(mem).filter(function (r) { return String(r[0] || '').trim() && String(r[3] || '').toUpperCase() !== 'N'; }).map(function (r) {
      return { name: r[0], team: r[1] || '', title: r[2] || '' };
    }),
    team: tabRows(q[2], TAB.team.w), facility: tabRows(q[3], TAB.facility.w),
    assets: tabRows(q[4], TAB.assets.w), contracts: tabRows(q[5], TAB.contracts.w), notes: tabRows(q[6], TAB.notes.w)
  };
}

/* ---------- 저장 ---------- */
async function saveWeek(sheets, p) {
  var week = String(p.week || '').trim(); if (!week) throw new Error('주차 값이 비어 있습니다.');
  var m = p.meta || {}, now = new Date(), z = function (n) { return ('0' + n).slice(-2); };
  var nowS = now.getFullYear() + '.' + z(now.getMonth() + 1) + '.' + z(now.getDate()) + ' ' + z(now.getHours()) + ':' + z(now.getMinutes());
  var wv = body_(await vget(sheets, DB, WEEKS)).filter(function (r) { return String(r[0] || '').trim(); });
  var found = false;
  var rows = wv.map(function (r) {
    if (String(r[0]).trim() === week) { found = true; return [week, m.start || '', m.end || '', m.dept || '', m.status || 'warning', m.statusLabel || '', m.date || '', nowS]; }
    return pad(r, WEEKS_W);
  });
  if (!found) rows.push([week, m.start || '', m.end || '', m.dept || '', m.status || 'warning', m.statusLabel || '', m.date || '', nowS]);
  await writeTab(sheets, WEEKS, rows, WEEKS_W);
  for (var k in TAB) { await replaceWeekRows(sheets, TAB[k], week, p[k] || []); }
  return { ok: true, savedAt: z(now.getHours()) + ':' + z(now.getMinutes()) + ':' + z(now.getSeconds()) };
}
async function replaceWeekRows(sheets, t, week, rows) {
  var kept = body_(await vget(sheets, DB, t.name)).map(function (r) { return pad(r, t.w); })
    .filter(function (r) { return r.join('').trim() !== '' && String(r[0]).trim() !== week; });
  (rows || []).forEach(function (rr) { if (rr.join('').trim() !== '') kept.push([week].concat(rr)); });
  await writeTab(sheets, t.name, kept, t.w);
}
async function deleteWeek(sheets, week) {
  week = String(week || '').trim(); if (!week) return { ok: false };
  var rows = body_(await vget(sheets, DB, WEEKS)).filter(function (r) { return String(r[0] || '').trim() && String(r[0]).trim() !== week; }).map(function (r) { return pad(r, WEEKS_W); });
  await writeTab(sheets, WEEKS, rows, WEEKS_W);
  for (var k in TAB) { await replaceWeekRows(sheets, TAB[k], week, []); }
  return { ok: true };
}
async function renameWeek(sheets, oldW, newW) {
  oldW = String(oldW || '').trim(); newW = String(newW || '').trim(); if (!oldW || !newW) return { ok: false };
  var wr = body_(await vget(sheets, DB, WEEKS)).filter(function (r) { return String(r[0] || '').trim(); })
    .map(function (r) { r = pad(r, WEEKS_W); if (String(r[0]).trim() === oldW) r[0] = newW; return r; });
  await writeTab(sheets, WEEKS, wr, WEEKS_W);
  for (var k in TAB) {
    var t = TAB[k];
    var rr = body_(await vget(sheets, DB, t.name)).map(function (r) { r = pad(r, t.w); if (String(r[0]).trim() === oldW) r[0] = newW; return r; })
      .filter(function (r) { return r.join('').trim() !== '' && String(r[0]).trim(); });
    await writeTab(sheets, t.name, rr, t.w);
  }
  return { ok: true };
}

/* ---------- 외부 자산시트 → 비품 집계 ---------- */
async function importAssets(sheets, startISO, endISO) {
  if (!ASSET) throw new Error('환경변수 ASSET_SHEET_ID 가 없습니다.');
  var meta = await sheets.spreadsheets.get({ spreadsheetId: ASSET });
  var title = null;
  (meta.data.sheets || []).forEach(function (s) { if (s.properties.sheetId === ASSET_GID) title = s.properties.title; });
  var vals;
  if (title) vals = await vget(sheets, ASSET, title);
  else {
    for (var i = 0; i < (meta.data.sheets || []).length; i++) {
      var t = meta.data.sheets[i].properties.title; var v = await vget(sheets, ASSET, t);
      var hr = findHeader(v); if (hr > -1) { title = t; vals = v; break; }
    }
  }
  if (!vals) throw new Error('자산시트 탭을 찾지 못했습니다.');
  var hRow = findHeader(vals); if (hRow < 0) hRow = 0;
  var hdr = (vals[hRow] || []).map(function (x) { return String(x).trim(); });
  function ci(n) { return hdr.indexOf(n); }
  var iCat = ci('분류'), iCat2 = ci('카테고리'), iQty = ci('수량'), iStat = ci('상태'), iReg = ci('지역'),
    iGive = ci('지급일'), iRet = ci('회수일'), iArr = ci('도착일'), iMemo = ci('메모'), iBuy = ci('매입처'), iVen = ci('거래처'), iTo = ci('지급대상자');
  function iso(x) { var s = String(x == null ? '' : x).trim(); if (!s) return ''; s = s.replace(/\./g, '-').replace(/\s/g, ''); var p = s.split('-'); if (p.length < 3) return ''; return p[0] + '-' + ('0' + p[1]).slice(-2) + '-' + ('0' + p[2]).slice(-2); }
  function inWin(x) { var d = iso(x); return d && d >= startISO && d <= endISO; }
  function regionOf(row) { var reg = iReg > -1 ? String(row[iReg] || '').trim() : ''; if (reg) return reg;
    var txt = [iMemo, iBuy, iVen, iTo].map(function (i) { return i > -1 ? String(row[i] || '') : ''; }).join(' ');
    return /부산|김해|창원|동래|서면|영남|덕천|경남|울산|사상/.test(txt) ? '부산' : '서울'; }
  var agg = {}, order = [];
  for (var r = hRow + 1; r < vals.length; r++) {
    var row = vals[r];
    var qty = parseFloat(String(iQty > -1 ? row[iQty] : '').replace(/[^0-9.\-]/g, '')); if (isNaN(qty) || qty === 0) continue;
    var cat = iCat > -1 ? String(row[iCat] || '').trim() : ''; if (!cat && iCat2 > -1) cat = String(row[iCat2] || '').trim(); if (!cat) cat = '기타';
    var reg = regionOf(row);
    var key = reg + '|' + cat; if (!(key in agg)) { agg[key] = { reg: reg, cat: cat, out: 0, inn: 0, stock: 0, pend: 0 }; order.push(key); }
    var stat = iStat > -1 ? String(row[iStat] || '').trim() : '';
    if (stat === '재고') agg[key].stock += qty; else if (stat === '발주완료') agg[key].pend += qty;
    if (iArr > -1 && inWin(row[iArr])) agg[key].inn += qty;
    if (iRet > -1 && inWin(row[iRet])) agg[key].inn += qty;
    if (iGive > -1 && inWin(row[iGive])) agg[key].out += qty;
  }
  var out = [];
  order.forEach(function (k) { var a = agg[k];
    if (a.out || a.inn || a.stock) out.push(['이번주', a.reg, a.cat, String(a.out), String(a.inn), String(a.stock)]);
    if (a.pend) out.push(['다음주', a.reg, a.cat, '0', String(a.pend), '']);
  });
  return out;
}
function findHeader(vals) { for (var i = 0; i < Math.min(vals.length, 5); i++) { var h = (vals[i] || []).map(String); if (h.indexOf('수량') > -1 && h.indexOf('상태') > -1) return i; } return -1; }
