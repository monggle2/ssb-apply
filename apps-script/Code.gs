/**
 * 사회연대은행 신청·면접 시스템 (구글 앱스 스크립트)
 * - 신청 페이지(GitHub Pages)에서 오는 사업 설정 요청·신청서 제출을 처리
 * - 담당자 화면(설계·접수함·적/부 판정·면접 기록지)을 제공
 * 이 스크립트는 스프레드시트에 붙어 있어야 합니다(확장 프로그램 → Apps Script).
 */

const SH = { P: '사업설정', S: '담당자', A: '신청목록', I: '면접기록', O: '서류판독', N: '사업자상태' };
const PH = ['사업ID', '사업명', '모듈', '모집시작', '모집마감', '신청코드', '접수중', '시연용', '안내문구', '설정JSON', '수정일시', '수정자'];
const SH_STAFF = ['이메일', '이름', '메모'];
const AH = ['접수번호', '사업ID', '사업명', '제출일시', '상호', '대표자', '연락처', '이메일', '상태', '사유구분', '판정메모', '판정일시', '면접총점', '서류폴더', '원본JSON'];
const OH = ['접수번호', '항목', '칸', '파일명', '파일ID', '글자수', '판독글자', '판독일시', '판독자'];
const NH = ['접수번호', '사업자번호', '상태', '상태코드', '과세유형', '폐업일', '조회일시', '조회자'];
const IH = ['접수번호', '면접관', '면접일', '총점', '영역점수JSON', '점수JSON', '메모JSON', '저장일시', '저장자'];
const TZ = 'Asia/Seoul';
const MAX_FILE = 10 * 1024 * 1024;
const MAX_TOTAL = 25 * 1024 * 1024;
const OK_EXT = ['pdf', 'jpg', 'jpeg', 'png', 'gif', 'webp', 'hwp', 'hwpx', 'doc', 'docx', 'xls', 'xlsx'];

/* ---------- 메뉴·초기 설정 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('신청 시스템')
    .addItem('1. 초기 설정', 'setup')
    .addItem('2. 담당자에게 서류 폴더 공유', 'shareFolder')
    .addItem('관리 정보 보기', 'showInfo')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(SH.P, PH);
  const staff = ensureSheet_(SH.S, SH_STAFF);
  ensureSheet_(SH.A, AH);
  ensureSheet_(SH.I, IH);
  ensureSheet_(SH.O, OH);
  ensureSheet_(SH.N, NH);
  const me = Session.getActiveUser().getEmail();
  if (me && !staffList_().includes(me.toLowerCase())) staff.appendRow([me, '', '초기 설정한 사람']);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ROOT_FOLDER')) {
    const f = DriveApp.createFolder('신청서류_' + ss.getName());
    props.setProperty('ROOT_FOLDER', f.getId());
  }
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  SpreadsheetApp.getUi().alert('초기 설정을 마쳤습니다.\n\n다음 단계: 배포 → 새 배포로 웹 앱 주소 두 개(신청용·담당자용)를 만드세요. 자세한 방법은 README를 보세요.');
}

function ensureSheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#e1efec');
  }
  return sh;
}

function shareFolder() {
  const folder = rootFolder_();
  const owner = Session.getEffectiveUser().getEmail().toLowerCase();
  const list = staffList_().filter(e => e !== owner);
  list.forEach(e => { try { folder.addEditor(e); } catch (err) {} });
  SpreadsheetApp.getUi().alert('서류 폴더를 담당자 ' + list.length + '명에게 편집 권한으로 공유했습니다.\n시트도 같은 사람들에게 공유되어 있어야 담당자 화면이 동작합니다.');
}

function showInfo() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER');
  SpreadsheetApp.getUi().alert('서류 폴더: ' + (id ? 'https://drive.google.com/drive/folders/' + id : '아직 없음(초기 설정을 먼저 실행)') +
    '\n신청 페이지 주소: ' + (props.getProperty('FORM_URL') || '담당자 화면에서 입력'));
}

/* ---------- 공통 ---------- */
function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function rows_(name) {
  const sh = sheet_(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function staffList_() { return rows_(SH.S).map(r => String(r[0]).trim().toLowerCase()).filter(Boolean); }
function me_() { return (Session.getActiveUser().getEmail() || '').toLowerCase(); }
function isStaff_() { const m = me_(); return !!m && staffList_().includes(m); }
function requireStaff_() { if (!isStaff_()) throw new Error('담당자로 등록된 계정이 아닙니다.'); return me_(); }
function rootFolder_() {
  const id = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER');
  if (!id) throw new Error('초기 설정이 필요합니다(시트 메뉴 → 신청 시스템 → 초기 설정).');
  return DriveApp.getFolderById(id);
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ymd_(d) { return d ? Utilities.formatDate(new Date(d), TZ, 'yyyy-MM-dd') : ''; }
function str_(v) { return v instanceof Date ? ymd_(v) : String(v == null ? '' : v); }
function clip_(v, n) {
  if (typeof v === 'string') return v.slice(0, n);
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  if (Array.isArray(v)) return v.slice(0, 50).map(x => clip_(x, n));
  if (v && typeof v === 'object') { const o = {}; Object.keys(v).slice(0, 50).forEach(k => o[k] = clip_(v[k], n)); return o; }
  return '';
}
/* 시트에 넣을 때 =, +, -, @ 로 시작하는 글자가 수식으로 바뀌지 않게 막음 */
function cell_(v) {
  if (v == null) return '';
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}
function escHtml_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

/* ---------- 사업 설정 ---------- */
function readPrograms_() {
  return rows_(SH.P).filter(r => r[0]).map(r => {
    let cfg = {};
    try { cfg = JSON.parse(r[9] || '{}'); } catch (e) {}
    return {
      id: String(r[0]), name: str_(r[1]), module: str_(r[2]) === 'loan' ? 'loan' : 'online',
      start: str_(r[3]), end: str_(r[4]), code: str_(r[5]), open: r[6] === true || r[6] === 'Y' || r[6] === 'TRUE',
      demo: r[7] === true || r[7] === 'Y' || r[7] === 'TRUE', notice: str_(r[8]),
      items: cfg.items || {}, qs: cfg.qs || {}, rules: cfg.rules || {},
      updatedAt: r[10] instanceof Date ? r[10].getTime() : (Number(r[10]) || 0), updatedBy: str_(r[11])
    };
  });
}
function inPeriod_(p) {
  const today = ymd_(new Date());
  if (p.start && today < p.start) return false;
  if (p.end && today > p.end) return false;
  return true;
}

/* 동의 항목은 Consent.gs 문구로 고정 */
function withConsent_(secs) {
  secs.forEach(s => s.items.forEach(it => it.fields.forEach(f => {
    if (f.t === 'consent' && typeof CONSENT !== 'undefined') { f.opts = CONSENT.map(c => c.title); f.bodies = CONSENT.map(c => c.body); }
  })));
  return secs;
}

/* ---------- 신청 페이지용 API ---------- */
function doGet(e) {
  const q = (e && e.parameter) || {};
  if (q.api === 'programs') {
    const list = readPrograms_().filter(p => p.open && inPeriod_(p))
      .map(p => ({ id: p.id, name: p.name, module: p.module, start: p.start, end: p.end, demo: p.demo }));
    return out_({ ok: true, programs: list });
  }
  if (q.api === 'config') {
    const p = readPrograms_().find(x => x.id === q.pid);
    if (!p || !p.open) return out_({ ok: false, error: '접수 중인 사업이 아닙니다.' });
    if (!inPeriod_(p)) return out_({ ok: false, error: '모집 기간이 아닙니다(' + (p.start || '') + ' ~ ' + (p.end || '') + ').' });
    const sections = withConsent_(effItems(p)).map(s => ({ sec: s.sec, title: s.title, items: s.items.filter(i => i.on).map(i => ({ key: i.key, label: i.label, req: i.req, help: i.help, anyOf: !!i.anyOf, fields: i.fields })) })).filter(s => s.items.length);
    return out_({ ok: true, program: { id: p.id, name: p.name, module: p.module, start: p.start, end: p.end, notice: p.notice, demo: p.demo, hasCode: !!p.code, sections } });
  }
  if (!isStaff_()) {
    return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;padding:24px;line-height:1.7"><h3>담당자 화면입니다</h3>시트의 <b>담당자</b> 탭에 등록된 구글 계정으로 열어 주세요.<br>현재 계정: ' + escHtml_(me_() || '확인할 수 없음 (담당자용 주소로 열었는지 확인해 주세요)') + '</div>').setTitle('신청 시스템');
  }
  return HtmlService.createTemplateFromFile('Admin').evaluate()
    .setTitle('신청·면접 시스템 담당자 화면')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    if (b.website) return out_({ ok: true, no: '-' }); // 자동 입력 방지용 숨은 칸
    const p = readPrograms_().find(x => x.id === b.pid);
    if (!p || !p.open) return out_({ ok: false, error: '접수 중인 사업이 아닙니다.' });
    if (!inPeriod_(p)) return out_({ ok: false, error: '모집 기간이 아닙니다.' });
    if (p.code && String(b.code || '').trim().toUpperCase() !== p.code.trim().toUpperCase()) return out_({ ok: false, error: '신청 코드가 맞지 않습니다. 공고문의 신청 코드를 확인해 주세요.', field: 'code' });

    const secs = withConsent_(effItems(p));
    const raw = b.answers || {};
    // 켜진 항목만 받고 글자 수를 제한
    const answers = {};
    const fileField = {};
    secs.forEach(s => s.items.filter(i => i.on).forEach(it => {
      const src = raw[it.key] || {};
      const o = {};
      it.fields.forEach(f => {
        if (f.t === 'file') { o[f.k] = []; fileField[it.key + '.' + f.k] = true; return; }
        const v = src[f.k];
        if (v == null || v === '') return;
        o[f.k] = clip_(v, f.t === 'textarea' ? (f.max || 2000) : 300);
      });
      answers[it.key] = o;
    }));
    const files = Array.isArray(b.files) ? b.files : [];
    if (files.length > 30) return out_({ ok: false, error: '첨부 파일은 30개까지 올릴 수 있습니다.' });
    let total = 0;
    for (const f of files) {
      if (!fileField[f.ik + '.' + f.fk]) return out_({ ok: false, error: '첨부 항목을 확인할 수 없습니다.' });
      const ext = String(f.name || '').split('.').pop().toLowerCase();
      if (!OK_EXT.includes(ext)) return out_({ ok: false, error: '올릴 수 없는 파일 형식입니다: ' + f.name });
      const size = Math.floor(String(f.data || '').length * 3 / 4);
      if (size > MAX_FILE) return out_({ ok: false, error: '파일 하나는 10MB 이하여야 합니다: ' + f.name });
      total += size;
      answers[f.ik][f.fk].push({ name: String(f.name).slice(0, 120) });
    }
    if (total > MAX_TOTAL) return out_({ ok: false, error: '첨부 파일 합계는 25MB 이하여야 합니다.' });
    const miss = [];
    secs.forEach(s => s.items.filter(i => i.on && i.req).forEach(it => {
      const vals = it.fields.map(f => [f, (answers[it.key] || {})[f.k]]);
      if (it.anyOf) { if (!vals.some(([f, v]) => filled(f, v))) miss.push(it.label); return; }
      if (vals.some(([f, v]) => !f.opt && !filled(f, v))) miss.push(it.label);
    }));
    if (miss.length) return out_({ ok: false, error: '필수 항목이 비어 있습니다: ' + miss.join(', ') });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const no = 'A' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
      // 파일 저장: 서류폴더 / 사업명(사업ID) / 접수번호 상호
      const root = rootFolder_();
      const pname = (p.name || p.id) + ' (' + p.id + ')';
      const pit = root.getFoldersByName(pname);
      const pf = pit.hasNext() ? pit.next() : root.createFolder(pname);
      const af = pf.createFolder(no + ' ' + String(((answers.biz || {}).name) || '').replace(/[\\/:*?"<>|]/g, ' ').slice(0, 40));
      Object.keys(fileField).forEach(k => { const [ik, fk] = k.split('.'); answers[ik][fk] = []; });
      files.forEach(f => {
        const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.type || 'application/octet-stream', String(f.name).slice(0, 120));
        const file = af.createFile(blob);
        answers[f.ik][f.fk].push({ id: file.getId(), name: file.getName(), type: f.type || '', size: file.getSize(), url: file.getUrl() });
      });
      // 신청 내용 원본도 폴더에 함께 보관
      let json = JSON.stringify(answers);
      const jf = af.createFile(Utilities.newBlob(json, 'application/json', '신청서_' + no + '.json'));
      if (json.length > 45000) json = 'FILE:' + jf.getId();
      const biz = answers.biz || {}, ct = answers.contact || {};
      const now = new Date();
      sheet_(SH.A).appendRow([no, p.id, p.name, now, cell_(biz.name), cell_(biz.ceo), cell_(ct.phone), cell_(ct.email), '접수', '', '', '', '', af.getUrl(), json]);
      writeProgramSheet_(p, secs, no, now, answers);
      return out_({ ok: true, no });
    } finally { lock.releaseLock(); }
  } catch (err) {
    return out_({ ok: false, error: '제출 중 오류가 났습니다: ' + err.message });
  }
}

/* 사업별 탭: 사람이 읽기 쉬운 형태로 한 줄씩 */
function writeProgramSheet_(p, secs, no, when, answers) {
  const name = ('신청_' + (p.name || p.id)).replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 90);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  const cols = [];
  secs.forEach(s => s.items.filter(i => i.on).forEach(it => it.fields.forEach(f => cols.push({ h: it.label + (it.fields.length > 1 ? ' - ' + f.label : ''), it, f }))));
  let head = sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
  if (!head.length) { head = ['접수번호', '제출일시']; sh.appendRow(head); sh.setFrozenRows(1); }
  cols.forEach(c => { if (!head.includes(c.h)) { head.push(c.h); sh.getRange(1, head.length).setValue(c.h); } });
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e1efec');
  const row = head.map(h => {
    if (h === '접수번호') return no;
    if (h === '제출일시') return when;
    const c = cols.find(x => x.h === h);
    if (!c) return '';
    const v = (answers[c.it.key] || {})[c.f.k];
    if (c.f.t === 'file') return (v || []).map(x => x.name + ' ' + (x.url || '')).join('\n');
    return cell_(plain(c.f, v));
  });
  sh.appendRow(row);
}

/* ---------- 담당자 화면 API (google.script.run) ---------- */
function adminInit() {
  const me = requireStaff_();
  const props = PropertiesService.getScriptProperties();
  return {
    me, programs: readPrograms_(),
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    folderUrl: props.getProperty('ROOT_FOLDER') ? 'https://drive.google.com/drive/folders/' + props.getProperty('ROOT_FOLDER') : '',
    formUrl: props.getProperty('FORM_URL') || ''
  };
}

function adminSetFormUrl(url) {
  requireStaff_();
  PropertiesService.getScriptProperties().setProperty('FORM_URL', String(url || '').trim());
  return true;
}

function adminSaveProgram(p) {
  const me = requireStaff_();
  if (!p || !p.id || !String(p.name || '').trim()) throw new Error('사업 이름이 필요합니다.');
  const sh = sheet_(SH.P);
  const row = [p.id, p.name, p.module === 'loan' ? 'loan' : 'online', p.start || '', p.end || '', p.code || '', p.open ? 'Y' : 'N', p.demo ? 'Y' : 'N', p.notice || '',
    JSON.stringify({ items: p.items || {}, qs: p.qs || {}, rules: p.rules || {} }), new Date(), me];
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = ids.indexOf(p.id);
  sh.getRange(1, 4, sh.getMaxRows(), 2).setNumberFormat('@'); // 날짜를 글자로 보관
  if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  return readPrograms_();
}

function adminListApps(pid) {
  requireStaff_();
  const ocr = {};
  rows_(SH.O).forEach(r => {
    const no = String(r[0]);
    (ocr[no] = ocr[no] || []).push({ ik: str_(r[1]), fk: str_(r[2]), name: str_(r[3]), len: Number(r[5]) || 0, text: String(r[6] || '').slice(0, 8000), err: Number(r[5]) < 0 ? String(r[6] || '') : '', at: r[7] instanceof Date ? r[7].getTime() : 0 });
  });
  const iv = {};
  rows_(SH.I).forEach(r => {
    const o = { interviewer: str_(r[1]), date: str_(r[2]), scores: {}, notes: {} };
    try { o.scores = JSON.parse(r[5] || '{}'); } catch (e) {}
    try { o.notes = JSON.parse(r[6] || '{}'); } catch (e) {}
    iv[String(r[0])] = o;
  });
  const nts = {};
  rows_(SH.N).forEach(r => { const t = v => str_(v).replace(/^'/, '');
    nts[String(r[0])] = { bno: t(r[1]), stt: str_(r[2]), cd: t(r[3]), tax: str_(r[4]), end: t(r[5]), at: r[6] instanceof Date ? r[6].getTime() : 0 }; });
  return rows_(SH.A).filter(r => String(r[1]) === pid).map(r => {
    let answers = {};
    try {
      const raw = String(r[14] || '{}');
      answers = JSON.parse(raw.indexOf('FILE:') === 0 ? DriveApp.getFileById(raw.slice(5)).getBlob().getDataAsString('UTF-8') : raw);
    } catch (e) {}
    return {
      id: String(r[0]), pid: String(r[1]), progName: str_(r[2]), submittedAt: r[3] instanceof Date ? r[3].getTime() : 0,
      bizName: str_(r[4]), ceo: str_(r[5]), status: str_(r[8]) || '접수', elig: { kind: str_(r[9]), note: str_(r[10]) },
      folder: str_(r[13]), answers, ocr: ocr[String(r[0])] || null, nts: nts[String(r[0])] || null, interview: iv[String(r[0])] || { scores: {}, notes: {}, interviewer: '', date: '' }
    };
  });
}

function findRow_(name, no) {
  const sh = sheet_(name);
  if (sh.getLastRow() < 2) return -1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(no);
  return i < 0 ? -1 : i + 2;
}

function adminSaveElig(no, status, kind, note) {
  requireStaff_();
  const r = findRow_(SH.A, no);
  if (r < 0) throw new Error('접수번호를 찾을 수 없습니다: ' + no);
  sheet_(SH.A).getRange(r, 9, 1, 4).setValues([[status || '접수', kind || '', note || '', new Date()]]);
  return true;
}

function adminSaveInterview(no, iv, total, areas) {
  const me = requireStaff_();
  const sh = sheet_(SH.I);
  const row = [no, iv.interviewer || '', iv.date || '', total == null ? '' : total, JSON.stringify(areas || {}), JSON.stringify(iv.scores || {}), JSON.stringify(iv.notes || {}), new Date(), me];
  const r = findRow_(SH.I, no);
  if (r < 0) sh.appendRow(row); else sh.getRange(r, 1, 1, row.length).setValues([row]);
  const ra = findRow_(SH.A, no);
  if (ra > 0) sheet_(SH.A).getRange(ra, 13).setValue(total == null ? '' : total);
  return true;
}

/* ---------- 서류 글자 읽기 (구글 드라이브 OCR) ----------
 * Apps Script 왼쪽 '서비스' ＋ → Drive API 를 추가해야 동작함 */
const OCR_EXT = ['pdf', 'jpg', 'jpeg', 'png', 'gif', 'docx'];
const OCR_SKIP = ['m_prod.photo', 'proof.file']; // 상품·사업장 사진은 읽지 않음

function ocrFile_(fileId, name) {
  if (typeof Drive === 'undefined') throw new Error('Drive API 서비스가 꺼져 있습니다. Apps Script 왼쪽 서비스 ＋ → Drive API 를 추가해 주세요.');
  const blob = DriveApp.getFileById(fileId).getBlob();
  let doc;
  if (Drive.Files.create) doc = Drive.Files.create({ name: '_판독_' + name, mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'ko', fields: 'id' });
  else doc = Drive.Files.insert({ title: '_판독_' + name, mimeType: 'application/vnd.google-apps.document' }, blob, { ocr: true, ocrLanguage: 'ko' });
  try { return docText_(doc.id); }
  finally { try { DriveApp.getFileById(doc.id).setTrashed(true); } catch (e) {} }
}

/* 변환된 구글 문서에서 글자 꺼내기: 드라이브 고급 서비스 → 드라이브 API 직접 호출 → 문서 열기 순서로 시도 */
function docText_(id) {
  try {
    const r = Drive.Files.export(id, 'text/plain', { alt: 'media' });
    if (typeof r === 'string' && r) return r;
    if (r && r.getDataAsString) return r.getDataAsString('UTF-8');
  } catch (e) {}
  try {
    const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=text%2Fplain',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() === 200) return res.getContentText('UTF-8');
  } catch (e) {}
  return DocumentApp.openById(id).getBody().getText();
}

function adminOcr(no) {
  const me = requireStaff_();
  const r = findRow_(SH.A, no);
  if (r < 0) throw new Error('접수번호를 찾을 수 없습니다: ' + no);
  const raw = String(sheet_(SH.A).getRange(r, 15).getValue() || '{}');
  const answers = JSON.parse(raw.indexOf('FILE:') === 0 ? DriveApp.getFileById(raw.slice(5)).getBlob().getDataAsString('UTF-8') : raw);
  const out = [];
  Object.keys(answers).forEach(ik => Object.keys(answers[ik] || {}).forEach(fk => {
    const v = answers[ik][fk];
    if (!Array.isArray(v)) return;
    v.forEach(f => {
      if (!f || !f.id || !f.name) return;
      const ext = String(f.name).split('.').pop().toLowerCase();
      if (!OCR_EXT.includes(ext) || OCR_SKIP.includes(ik + '.' + fk)) return;
      try {
        const text = ocrFile_(f.id, f.name);
        out.push({ ik, fk, name: f.name, id: f.id, len: text.length, text });
      } catch (e) {
        if (/Drive API|권한|permission|authoriz/i.test(e.message)) throw e; // 권한 문제는 기록하지 않고 바로 알림
        out.push({ ik, fk, name: f.name, id: f.id, len: -1, text: '', err: e.message });
      }
    });
  }));
  const sh = sheet_(SH.O) || ensureSheet_(SH.O, OH);
  // 같은 접수번호의 예전 판독 결과 지우기
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  for (let i = ids.length - 1; i >= 0; i--) if (ids[i] === no) sh.deleteRow(i + 2);
  const now = new Date();
  out.forEach(o => sh.appendRow([no, o.ik, o.fk, o.name, o.id, o.len, o.err ? o.err : cell_(o.text.slice(0, 45000)), now, me]));
  return out.map(o => ({ ik: o.ik, fk: o.fk, name: o.name, len: o.len, text: o.text.slice(0, 8000), err: o.err || '', at: now.getTime() }));
}

/* ---------- 사업자 상태 조회 (국세청 공공데이터 API) ----------
 * 공공데이터포털 인증키(Decoding)를 스크립트 속성 NTS_KEY 에 넣어야 동작함 */
function adminBizStatus(nos) {
  const me = requireStaff_();
  const key = PropertiesService.getScriptProperties().getProperty('NTS_KEY');
  if (!key) throw new Error('인증키가 없습니다. Apps Script 프로젝트 설정 → 스크립트 속성에 NTS_KEY 를 넣어 주세요.');
  const want = {};
  (nos || []).forEach(n => want[String(n)] = true);
  const list = [];
  rows_(SH.A).forEach(r => {
    if (!want[String(r[0])]) return;
    let a = {};
    try { const raw = String(r[14] || '{}'); a = JSON.parse(raw.indexOf('FILE:') === 0 ? DriveApp.getFileById(raw.slice(5)).getBlob().getDataAsString('UTF-8') : raw); } catch (e) {}
    const b = String(((a.biz || {}).brn) || '').replace(/[^0-9]/g, '');
    list.push({ no: String(r[0]), b });
  });
  const ok = list.filter(x => x.b.length === 10);
  const res = {};
  for (let i = 0; i < ok.length; i += 100) {
    const chunk = ok.slice(i, i + 100);
    // 일반 인증키 Decoding·Encoding 어느 쪽을 넣어도 되게 두 형태로 시도
    const k = key.trim();
    const forms = /%[0-9A-Fa-f]{2}/.test(k) ? [k, encodeURIComponent(decodeURIComponent(k))] : [encodeURIComponent(k), k];
    let resp;
    for (const f of forms) {
      resp = UrlFetchApp.fetch('https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey=' + f, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ b_no: chunk.map(x => x.b) })
      });
      if (resp.getResponseCode() !== 401) break;
    }
    const code = resp.getResponseCode();
    if (code === 401) throw new Error('인증키가 아직 등록되지 않았습니다. 발급 직후에는 1~2시간 걸릴 수 있습니다. 스크립트 속성 NTS_KEY 값(일반 인증키)도 확인해 주세요.');
    if (code !== 200) throw new Error('국세청 조회 실패(' + code + '): ' + resp.getContentText().slice(0, 200));
    const j = JSON.parse(resp.getContentText());
    (j.data || []).forEach(d => { res[d.b_no] = d; });
  }
  const sh = sheet_(SH.N) || ensureSheet_(SH.N, NH);
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const now = new Date();
  const out = {};
  list.forEach(x => {
    const d = res[x.b];
    const o = x.b.length !== 10 ? { bno: x.b, stt: '', cd: '', tax: '사업자번호 자릿수 오류', end: '' }
      : d ? { bno: x.b, stt: d.b_stt || '', cd: d.b_stt_cd || '', tax: d.tax_type || '', end: d.end_dt || '' }
      : { bno: x.b, stt: '', cd: '', tax: '조회 결과 없음', end: '' };
    const tx = v => v ? "'" + v : '';
    const row = [x.no, tx(x.b), o.stt, tx(o.cd), o.tax, tx(o.end), now, me];
    const i = ids.indexOf(x.no);
    if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]); else { sh.appendRow(row); ids.push(x.no); }
    out[x.no] = Object.assign(o, { at: now.getTime() });
  });
  return out;
}
