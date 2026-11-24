/**
 * 사회연대은행 신청·면접 시스템 (구글 앱스 스크립트)
 * - 신청 페이지(GitHub Pages)에서 오는 사업 설정 요청·신청서 제출을 처리
 * - 담당자 화면(설계·접수함·적/부 판정·면접 기록지)을 제공
 * - 자동 처리(5분마다): 새 신청의 서류 글자 읽기 → 항목 정리 → 국세청 사업자 상태 → 자동 점검 결과 기록
 * 이 스크립트는 스프레드시트에 붙어 있어야 합니다(확장 프로그램 → Apps Script).
 */

const SH = { P: '사업설정', S: '담당자', A: '신청목록', I: '면접기록', O: '서류판독', N: '사업자상태', L: '변경이력', C: '동의문구이력' };
const PH = ['사업ID', '사업명', '모듈', '모집시작', '모집마감', '신청코드', '접수중', '시연용', '안내문구', '설정JSON', '수정일시', '수정자'];
const SH_STAFF = ['이메일', '이름', '메모'];
// 신청목록: 앞 15칸은 위치 고정, 그 뒤 칸은 이름으로 찾음
const AH = ['접수번호', '사업ID', '사업명', '제출일시', '상호', '대표자', '연락처', '이메일', '상태', '사유구분', '판정메모', '판정일시', '면접총점', '서류폴더', '원본JSON',
  '사업자번호', '처리상태', '처리시도', '처리일시', '처리메모', '자동판정', '탈락후보', '확인필요', '자동사유', '판정자', '제출ID', '동의버전', '동의일시',
  '판독필요', '보완토큰', '보완요청내용', '보완요청일시', '보완제출일시', '면접위원수', '최종결과', '최종결과일시', '최종결과자'];
const OH = ['접수번호', '항목', '칸', '파일명', '파일ID', '글자수', '판독글자', '판독일시', '판독자'];
const NH = ['접수번호', '사업자번호', '상태', '상태코드', '과세유형', '폐업일', '조회일시', '조회자'];
const IH = ['접수번호', '면접관', '면접일', '총점', '영역점수JSON', '점수JSON', '메모JSON', '저장일시', '저장자'];
const LH = ['일시', '사용자', '접수번호', '구분', '이전', '이후'];
const CH = ['동의버전', '처음사용일시', '문구JSON'];
const TZ = 'Asia/Seoul';
const MAX_FILE = 10 * 1024 * 1024;
const MAX_TOTAL = 25 * 1024 * 1024;
const OK_EXT = ['pdf', 'jpg', 'jpeg', 'png', 'gif', 'webp', 'hwp', 'hwpx', 'doc', 'docx', 'xls', 'xlsx'];
const AUTO_MAX_TRY = 3;

/* ---------- 메뉴·초기 설정 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('신청 시스템')
    .addItem('1. 초기 설정', 'setup')
    .addItem('2. 담당자에게 서류 폴더 공유', 'shareFolder')
    .addSeparator()
    .addItem('3. 자동 처리 켜기 (5분마다)', 'autoOn')
    .addItem('4. 자동 처리 끄기', 'autoOff')
    .addItem('자동 처리 지금 한 번 실행', 'autoRun')
    .addSeparator()
    .addItem('관리 정보 보기', 'showInfo')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(SH.P, PH);
  const staff = ensureSheet_(SH.S, SH_STAFF);
  cols_(SH.A, AH);
  ensureSheet_(SH.I, IH);
  ensureSheet_(SH.O, OH);
  ensureSheet_(SH.N, NH);
  ensureSheet_(SH.L, LH);
  const me = Session.getActiveUser().getEmail();
  if (me && !staffList_().includes(me.toLowerCase())) staff.appendRow([me, '', '초기 설정한 사람']);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ROOT_FOLDER')) {
    const f = DriveApp.createFolder('신청서류_' + ss.getName());
    props.setProperty('ROOT_FOLDER', f.getId());
  }
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  SpreadsheetApp.getUi().alert('초기 설정을 마쳤습니다.\n\n새 신청을 자동으로 읽게 하려면 메뉴 → 신청 시스템 → 3. 자동 처리 켜기 를 누르세요.');
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

/* 머리글에 없는 칸은 뒤에 붙이고, {칸이름: 열번호} 를 돌려줌 */
function cols_(name, header) {
  const sh = ensureSheet_(name, header);
  const lc = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lc).getValues()[0].map(String);
  let changed = false;
  header.forEach(h => { if (!head.includes(h)) { head.push(h); changed = true; } });
  if (changed) sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e1efec');
  const m = { _sh: sh, _n: head.length };
  head.forEach((h, i) => { if (h && !m[h]) m[h] = i + 1; });
  return m;
}
function setCells_(m, r, obj) { Object.keys(obj).forEach(k => { if (m[k]) m._sh.getRange(r, m[k]).setValue(obj[k]); }); }
function colVals_(sh, col) { return sh.getLastRow() > 1 ? sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []; }

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
    '\n신청 페이지 주소: ' + (props.getProperty('FORM_URL') || '담당자 화면에서 입력') +
    '\n자동 처리: ' + (props.getProperty('AUTO_ON') === 'Y' ? '켜짐' : '꺼짐') + (props.getProperty('AUTO_LAST') ? ' · 마지막 실행 ' + props.getProperty('AUTO_LAST') : ''));
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
function nowStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
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
function answersOf_(raw) {
  raw = String(raw || '{}');
  try { return JSON.parse(raw.indexOf('FILE:') === 0 ? DriveApp.getFileById(raw.slice(5)).getBlob().getDataAsString('UTF-8') : raw); } catch (e) { return {}; }
}
function log_(who, no, kind, before, after) {
  try {
    const sh = ensureSheet_(SH.L, LH);
    sh.appendRow([new Date(), who || '', no || '', kind, cell_(String(before == null ? '' : before).slice(0, 500)), cell_(String(after == null ? '' : after).slice(0, 500))]);
  } catch (e) {}
}

/* ---------- 사업 설정 ---------- */
function readPrograms_() {
  return rows_(SH.P).filter(r => r[0]).map(r => {
    let cfg = {};
    try { cfg = JSON.parse(r[9] || '{}'); } catch (e) {}
    return {
      id: String(r[0]), name: str_(r[1]), module: str_(r[2]) === 'loan' ? 'loan' : 'online',
      start: str_(r[3]), end: str_(r[4]), endTime: cfg.endTime || '', code: str_(r[5]), open: r[6] === true || r[6] === 'Y' || r[6] === 'TRUE',
      demo: r[7] === true || r[7] === 'Y' || r[7] === 'TRUE', notice: str_(r[8]),
      items: cfg.items || {}, qs: cfg.qs || {}, rules: cfg.rules || {}, custom: cfg.custom || [], sel: cfg.sel || { n: '', m: '' },
      updatedAt: r[10] instanceof Date ? r[10].getTime() : (Number(r[10]) || 0), updatedBy: str_(r[11])
    };
  });
}
/* 모집 기간: 시작일 00:00 ~ 마감일 마감시각(없으면 23:59) */
function inPeriod_(p) {
  const now = nowStr_();
  if (p.start && now < p.start + ' 00:00') return false;
  if (p.end && now > p.end + ' ' + (p.endTime || '23:59')) return false;
  return true;
}

/* 동의 항목은 Consent.gs 문구로 고정 */
function withConsent_(secs) {
  secs.forEach(s => s.items.forEach(it => it.fields.forEach(f => {
    if (f.t === 'consent' && typeof CONSENT !== 'undefined') { f.opts = CONSENT.map(c => c.title); f.bodies = CONSENT.map(c => c.body); }
  })));
  return secs;
}
/* 동의 문구 버전: 문구 내용으로 만든 8자리 값. 처음 쓰일 때 문구 전체를 이력 시트에 보관 */
function consentVer_() {
  if (typeof CONSENT === 'undefined') return '';
  const txt = JSON.stringify(CONSENT);
  const ver = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, txt, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').slice(0, 8);
  try {
    const sh = ensureSheet_(SH.C, CH);
    if (!colVals_(sh, 1).includes(ver)) sh.appendRow([ver, new Date(), txt.slice(0, 45000)]);
  } catch (e) {}
  return ver;
}

/* ---------- 신청 페이지용 API ---------- */
function doGet(e) {
  const q = (e && e.parameter) || {};
  if (q.api === 'programs') {
    const list = readPrograms_().filter(p => p.open && inPeriod_(p))
      .map(p => ({ id: p.id, name: p.name, module: p.module, start: p.start, end: p.end, endTime: p.endTime, demo: p.demo }));
    return out_({ ok: true, programs: list });
  }
  if (q.api === 'config') {
    const p = readPrograms_().find(x => x.id === q.pid);
    if (!p || !p.open) return out_({ ok: false, error: '접수 중인 사업이 아닙니다.' });
    if (!inPeriod_(p)) return out_({ ok: false, error: '모집 기간이 아닙니다(' + (p.start || '') + ' ~ ' + (p.end || '') + (p.endTime ? ' ' + p.endTime : '') + ').' });
    const sections = withConsent_(effItems(p)).map(s => ({ sec: s.sec, title: s.title, items: s.items.filter(i => i.on).map(i => ({ key: i.key, label: i.label, req: i.req, help: i.help, anyOf: !!i.anyOf, fields: i.fields })) })).filter(s => s.items.length);
    return out_({ ok: true, program: { id: p.id, name: p.name, module: p.module, start: p.start, end: p.end, endTime: p.endTime, notice: p.notice, demo: p.demo, hasCode: !!p.code, sections } });
  }
  if (q.api === 'bizcheck') {
    const p = readPrograms_().find(x => x.id === q.pid);
    if (!p || !p.open || !inPeriod_(p)) return out_({ ok: false, error: '접수 중인 사업이 아닙니다.' });
    const c = bizCheck_(p, q.b);
    delete c.raw;
    return out_(c);
  }
  if (q.api === 'fix') {
    const f = fixFind_(q.no, q.t);
    if (!f) return out_({ ok: false, error: '링크가 맞지 않거나 보완 제출 기간이 아닙니다. 담당자에게 문의해 주세요.' });
    return out_({ ok: true, fix: { no: f.no, prog: f.p.name, biz: f.biz, note: f.note, slots: f.slots, done: f.done } });
  }
  if (!isStaff_()) {
    return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;padding:24px;line-height:1.7"><h3>담당자 화면입니다</h3>시트의 <b>담당자</b> 탭에 등록된 구글 계정으로 열어 주세요.<br>현재 계정: ' + escHtml_(me_() || '확인할 수 없음 (담당자용 주소로 열었는지 확인해 주세요)') + '</div>').setTitle('신청 시스템');
  }
  return HtmlService.createTemplateFromFile('Admin').evaluate()
    .setTitle('신청·면접 시스템 담당자 화면')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  let af = null;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.website) return out_({ ok: true, no: '-' }); // 자동 입력 방지용 숨은 칸
    if (b.fix) return out_(fixSubmit_(b));
    const p = readPrograms_().find(x => x.id === b.pid);
    if (!p || !p.open) return out_({ ok: false, error: '접수 중인 사업이 아닙니다.' });
    if (!inPeriod_(p)) return out_({ ok: false, error: '모집 기간이 아닙니다.' });
    if (p.code && String(b.code || '').trim().toUpperCase() !== p.code.trim().toUpperCase()) return out_({ ok: false, error: '신청 코드가 맞지 않습니다. 공고문의 신청 코드를 확인해 주세요.', field: 'code' });
    const sid = String(b.sid || '').replace(/[^\w\-]/g, '').slice(0, 40);

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
    // 사업자번호 국세청 확인 (신청 화면과 같은 기준으로 다시 확인)
    let bc = null;
    if (answers.biz && answers.biz.brn) {
      bc = bizCheck_(p, answers.biz.brn);
      if (!bc.allowed) return out_({ ok: false, error: bc.msg, field: 'biz' });
      answers.__biz = { stt: bc.stt || '', cd: bc.cd || '', tax: bc.tax || '', end: bc.end || '', kind: bc.kind || '', tk: bc.tk || '', doc: bc.doc || '', unverified: !!bc.unverified, demo: !!bc.demo, at: Date.now() };
      if (bc.btype && !bc.unverified) answers.biz.btype = bc.btype;
    }

    const miss = [];
    secs.forEach(s => s.items.filter(i => i.on && i.req).forEach(it => {
      const vals = it.fields.map(f => [f, (answers[it.key] || {})[f.k]]);
      if (it.anyOf) { if (!vals.some(([f, v]) => filled(f, v))) miss.push(it.label); return; }
      if (vals.some(([f, v]) => !f.opt && !filled(f, v))) miss.push(it.label);
    }));
    if (miss.length) return out_({ ok: false, error: '필수 항목이 비어 있습니다: ' + miss.join(', ') });

    const m = cols_(SH.A, AH);
    const sh = m._sh;
    // 같은 제출(인터넷이 끊겨 다시 누른 경우)은 한 번만 저장
    const sidHit = () => { if (!sid) return ''; const v = colVals_(sh, m['제출ID']); const i = v.indexOf(sid); return i < 0 ? '' : colVals_(sh, 1)[i]; };
    const dup0 = sidHit();
    if (dup0) return out_({ ok: true, no: dup0, dup: true });

    // 접수번호: 날짜 + 임의 5자리, 이미 있으면 다시 뽑음
    const used = colVals_(sh, 1);
    let no;
    do { no = 'A' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 7).toUpperCase(); } while (used.includes(no));

    // 파일 저장은 잠금 밖에서 (마감 직전 신청이 몰려도 서로 기다리지 않게)
    const root = rootFolder_();
    const pname = (p.name || p.id) + ' (' + p.id + ')';
    const pit = root.getFoldersByName(pname);
    const pf = pit.hasNext() ? pit.next() : root.createFolder(pname);
    af = pf.createFolder(no + ' ' + String(((answers.biz || {}).name) || '').replace(/[\\/:*?"<>|]/g, ' ').slice(0, 40));
    Object.keys(fileField).forEach(k => { const [ik, fk] = k.split('.'); answers[ik][fk] = []; });
    files.forEach(f => {
      const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.type || 'application/octet-stream', String(f.name).slice(0, 120));
      const file = af.createFile(blob);
      answers[f.ik][f.fk].push({ id: file.getId(), name: file.getName(), type: f.type || '', size: file.getSize(), url: file.getUrl() });
    });
    let json = JSON.stringify(answers);
    const jf = af.createFile(Utilities.newBlob(json, 'application/json', '신청서_' + no + '.json'));
    if (json.length > 45000) json = 'FILE:' + jf.getId();
    const biz = answers.biz || {}, ct = answers.contact || {};
    const ci = secs.flatMap(s => s.items).find(i => i.on && i.key === 'consent');
    const agreed = ci ? ((answers.consent || {})[(ci.fields.find(f => f.t === 'consent') || {}).k] || []).length : 0;
    const cver = consentVer_();
    const now = new Date();

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const dup1 = sidHit();
      if (dup1) { try { af.setTrashed(true); } catch (x) {} return out_({ ok: true, no: dup1, dup: true }); }
      const rowObj = {
        '접수번호': no, '사업ID': p.id, '사업명': p.name, '제출일시': now, '상호': cell_(biz.name), '대표자': cell_(biz.ceo), '연락처': cell_(ct.phone), '이메일': cell_(ct.email),
        '상태': '접수', '서류폴더': af.getUrl(), '원본JSON': json, '사업자번호': String(biz.brn || '').replace(/[^0-9]/g, '') ? "'" + String(biz.brn || '').replace(/[^0-9]/g, '') : '',
        '처리상태': '대기', '처리시도': 0, '판독필요': 'Y', '제출ID': sid, '동의버전': agreed ? cver : '', '동의일시': agreed ? now : ''
      };
      const row = new Array(m._n).fill('');
      Object.keys(rowObj).forEach(k => { if (m[k]) row[m[k] - 1] = rowObj[k]; });
      sh.appendRow(row);
      af = null; // 여기부터는 저장 완료: 오류가 나도 파일을 지우지 않음
      if (bc && bc.raw && !bc.demo) { try { ntsSave_([{ no, b: String(answers.biz.brn).replace(/[^0-9]/g, '') }], bc.raw, '신청 시 확인'); } catch (x) {} }
      try { writeProgramSheet_(p, secs, no, now, answers); } catch (x) { log_('시스템', no, '사업별 탭 기록 실패', '', x.message); }
    } finally { lock.releaseLock(); }
    return out_({ ok: true, no });
  } catch (err) {
    if (af) { try { af.setTrashed(true); } catch (x) {} }
    return out_({ ok: false, error: '제출 중 오류가 났습니다. 잠시 후 다시 제출해 주세요. (' + err.message + ')' });
  }
}

/* 사업별 탭: 사람이 읽기 쉬운 형태로 한 줄씩 (사업 이름을 바꿔도 같은 탭을 쓰도록 사업ID로 찾음) */
function writeProgramSheet_(p, secs, no, when, answers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const key = 'TAB_' + p.id;
  let sh = props.getProperty(key) ? ss.getSheets().find(s => String(s.getSheetId()) === props.getProperty(key)) : null;
  if (!sh) {
    let name = ('신청_' + (p.name || p.id)).replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 90);
    if (ss.getSheetByName(name) && !props.getProperty(key)) sh = ss.getSheetByName(name);
    else { while (ss.getSheetByName(name)) name += '_'; sh = ss.insertSheet(name); }
    props.setProperty(key, String(sh.getSheetId()));
  }
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
  const counts = {};
  const sh = sheet_(SH.A);
  if (sh) colVals_(sh, 2).forEach(pid => counts[pid] = (counts[pid] || 0) + 1);
  return {
    me, programs: readPrograms_(), counts,
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    folderUrl: props.getProperty('ROOT_FOLDER') ? 'https://drive.google.com/drive/folders/' + props.getProperty('ROOT_FOLDER') : '',
    formUrl: props.getProperty('FORM_URL') || '',
    auto: { on: props.getProperty('AUTO_ON') === 'Y', last: props.getProperty('AUTO_LAST') || '' }
  };
}

function adminSetFormUrl(url) {
  const me = requireStaff_();
  PropertiesService.getScriptProperties().setProperty('FORM_URL', String(url || '').trim());
  log_(me, '', '신청 페이지 주소', '', url);
  return true;
}

function adminSaveProgram(p) {
  const me = requireStaff_();
  if (!p || !p.id || !String(p.name || '').trim()) throw new Error('사업 이름이 필요합니다.');
  const sh = sheet_(SH.P);
  const row = [p.id, p.name, p.module === 'loan' ? 'loan' : 'online', p.start || '', p.end || '', p.code || '', p.open ? 'Y' : 'N', p.demo ? 'Y' : 'N', p.notice || '',
    JSON.stringify({ items: p.items || {}, qs: p.qs || {}, rules: p.rules || {}, endTime: p.endTime || '', custom: cleanCustom_(p.custom), sel: p.sel || { n: '', m: '' } }), new Date(), me];
  const ids = colVals_(sh, 1);
  const i = ids.indexOf(p.id);
  sh.getRange(1, 4, sh.getMaxRows(), 2).setNumberFormat('@'); // 날짜를 글자로 보관
  const before = i >= 0 ? sh.getRange(i + 2, 1, 1, 9).getValues()[0].map(str_).join(' / ') : '';
  if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  log_(me, '', i >= 0 ? '사업 설정 수정' : '사업 만들기', before, row.slice(0, 9).map(str_).join(' / '));
  try { refreshVerdicts_(r => String(r[1]) === p.id); } catch (e) {}
  return readPrograms_();
}

/* 담당자가 만든 추가 항목: 형식·길이를 정리해서 저장 */
const C_TYPES_OK = ['text', 'textarea', 'number', 'date', 'select', 'radio', 'checks', 'file', 'url', 'email', 'tel'];
function cleanCustom_(arr) {
  if (!Array.isArray(arr)) return [];
  const seen = {};
  return arr.slice(0, 50).map(c => {
    const key = String((c && c.key) || '').replace(/[^\w]/g, '').slice(0, 30);
    if (!/^c_/.test(key) || seen[key]) return null;
    seen[key] = true;
    const fk = {};
    const fields = (Array.isArray(c.fields) ? c.fields : []).slice(0, 10).map((f, i) => {
      let k = String((f && f.k) || '').replace(/[^\w]/g, '').slice(0, 10) || ('f' + (i + 1));
      while (fk[k]) k += 'x';
      fk[k] = true;
      const t = C_TYPES_OK.includes(f.t) ? f.t : 'text';
      const o = { k, label: String(f.label || '').slice(0, 60) || ('칸 ' + (i + 1)), t };
      if (['select', 'radio', 'checks'].includes(t)) {
        o.opts = (Array.isArray(f.opts) ? f.opts : []).map(x => String(x).slice(0, 60)).filter(Boolean).slice(0, 30);
        if (!o.opts.length) o.opts = ['예', '아니오'];
      }
      if (t === 'textarea') o.max = Math.min(Math.max(parseInt(f.max, 10) || 500, 10), 5000);
      if (t === 'number' && f.unit) o.unit = String(f.unit).slice(0, 10);
      return o;
    });
    if (!fields.length) return null;
    return { key, sec: String(c.sec || 'X').slice(0, 4), label: String(c.label || '새 항목').slice(0, 80), help: String(c.help || '').slice(0, 200), req: !!c.req, fields };
  }).filter(Boolean);
}

function adminListApps(pid) {
  requireStaff_();
  const m = cols_(SH.A, AH);
  // 판독 결과는 글자 수만 (원문은 신청을 열 때 adminGetOcr 로)
  const ocr = {};
  const so = sheet_(SH.O);
  if (so && so.getLastRow() > 1) {
    const n = so.getLastRow() - 1;
    const a = so.getRange(2, 1, n, 6).getValues(), t = so.getRange(2, 8, n, 1).getValues();
    a.forEach((r, i) => {
      const no = String(r[0]);
      (ocr[no] = ocr[no] || []).push({ ik: str_(r[1]), fk: str_(r[2]), name: str_(r[3]), len: Number(r[5]) || 0, err: Number(r[5]) < 0 ? '읽기 오류' : '', at: t[i][0] instanceof Date ? t[i][0].getTime() : 0 });
    });
  }
  const ivs = ivMap_();
  const nts = {};
  rows_(SH.N).forEach(r => { const t = v => str_(v).replace(/^'/, '');
    nts[String(r[0])] = { bno: t(r[1]), stt: str_(r[2]), cd: t(r[3]), tax: str_(r[4]), end: t(r[5]), at: r[6] instanceof Date ? r[6].getTime() : 0 }; });
  const g = (r, k) => m[k] ? r[m[k] - 1] : '';
  return rows_(SH.A).filter(r => String(r[1]) === pid).map(r => {
    const no = String(r[0]);
    return {
      id: no, pid: String(r[1]), progName: str_(r[2]), submittedAt: r[3] instanceof Date ? r[3].getTime() : 0,
      bizName: str_(r[4]), ceo: str_(r[5]), status: str_(r[8]) || '접수', elig: { kind: str_(r[9]), note: str_(r[10]), by: str_(g(r, '판정자')), at: r[11] instanceof Date ? r[11].getTime() : 0 },
      folder: str_(r[13]), answers: answersOf_(r[14]), ocr: ocr[no] || null, nts: nts[no] || null,
      proc: { st: str_(g(r, '처리상태')), tries: Number(g(r, '처리시도')) || 0, at: g(r, '처리일시') instanceof Date ? g(r, '처리일시').getTime() : 0, memo: str_(g(r, '처리메모')) },
      auto: str_(g(r, '자동판정')) ? { verdict: str_(g(r, '자동판정')), f: Number(g(r, '탈락후보')) || 0, w: Number(g(r, '확인필요')) || 0, reason: str_(g(r, '자동사유')) } : null,
      consent: { ver: str_(g(r, '동의버전')), at: g(r, '동의일시') instanceof Date ? g(r, '동의일시').getTime() : 0 },
      ivs: ivs[no] || [],
      fix: { link: g(r, '보완토큰') ? fixLink_(no, str_(g(r, '보완토큰'))) : '', note: str_(g(r, '보완요청내용')), reqAt: g(r, '보완요청일시') instanceof Date ? g(r, '보완요청일시').getTime() : 0, subAt: g(r, '보완제출일시') instanceof Date ? g(r, '보완제출일시').getTime() : 0 },
      result: { v: str_(g(r, '최종결과')), by: str_(g(r, '최종결과자')), at: g(r, '최종결과일시') instanceof Date ? g(r, '최종결과일시').getTime() : 0 }
    };
  });
}

/* 한 신청의 판독 원문 */
function adminGetOcr(no) {
  requireStaff_();
  return ocrRows_(no);
}
function ocrRows_(no) {
  const out = [];
  rows_(SH.O).forEach(r => {
    if (String(r[0]) !== no) return;
    const len = Number(r[5]) || 0;
    out.push({ ik: str_(r[1]), fk: str_(r[2]), name: str_(r[3]), len, text: len < 0 ? '' : String(r[6] || '').slice(0, 20000), err: len < 0 ? String(r[6] || '') : '', at: r[7] instanceof Date ? r[7].getTime() : 0 });
  });
  return out;
}

function findRow_(name, no) {
  const sh = sheet_(name);
  if (!sh || sh.getLastRow() < 2) return -1;
  const i = colVals_(sh, 1).indexOf(no);
  return i < 0 ? -1 : i + 2;
}

function adminSaveElig(no, status, kind, note, fixNote) {
  const me = requireStaff_();
  const m = cols_(SH.A, AH);
  const r = findRow_(SH.A, no);
  if (r < 0) throw new Error('접수번호를 찾을 수 없습니다: ' + no);
  const old = m._sh.getRange(r, 9, 1, 3).getValues()[0].map(str_);
  const now = new Date();
  m._sh.getRange(r, 9, 1, 4).setValues([[status || '접수', kind || '', note || '', now]]);
  setCells_(m, r, { '판정자': me });
  let link = '';
  if (status === '보완 요청') {
    let tok = str_(m._sh.getRange(r, m['보완토큰']).getValue());
    if (!tok) tok = Utilities.getUuid().replace(/-/g, '');
    setCells_(m, r, { '보완토큰': tok, '보완요청내용': cell_(String(fixNote || '').slice(0, 1000)), '보완요청일시': now });
    link = fixLink_(no, tok);
  }
  log_(me, no, '적/부 판정', old.filter(Boolean).join(' / '), [status, kind, note].filter(Boolean).join(' / ') + (status === '보완 요청' && fixNote ? ' / 요청: ' + fixNote : ''));
  return { by: me, at: now.getTime(), link, fixNote: status === '보완 요청' ? String(fixNote || '') : undefined };
}

/* ---------- 사업자번호 확인 (신청 화면) ----------
 * 국세청 조회 → 폐업·미등록 차단, 휴업은 사업 설정(휴업자 허용)에 따라, 과세유형에 맞는 매출 서류 안내 */
const NTS_DAY_LIMIT = 5000;
function ntsQuota_() {
  const props = PropertiesService.getScriptProperties();
  const today = ymd_(new Date());
  const [d, c] = String(props.getProperty('NTS_DAY') || '').split('|');
  const n = d === today ? (Number(c) || 0) : 0;
  if (n >= NTS_DAY_LIMIT) return false;
  props.setProperty('NTS_DAY', today + '|' + (n + 1));
  return true;
}
function bizCheck_(p, brnRaw) {
  const b = String(brnRaw || '').replace(/[^0-9]/g, '');
  const R = rulesOf(p);
  if (b.length !== 10) return { ok: true, allowed: false, msg: '사업자등록번호 10자리를 확인해 주세요.' };
  let d = null, demo = false, unavailable = '';
  if (p.demo && b === '1234567890') { demo = true; d = { b_no: b, b_stt: '계속사업자', b_stt_cd: '01', tax_type: '부가가치세 일반과세자', end_dt: '' }; }
  else {
    const cache = CacheService.getScriptCache();
    const hit = cache.get('nts_' + b);
    if (hit) d = JSON.parse(hit);
    else {
      try {
        if (!PropertiesService.getScriptProperties().getProperty('NTS_KEY')) throw new Error('인증키 없음');
        if (!ntsQuota_()) throw new Error('오늘 조회 한도 초과');
        const res = ntsFetch_([b]);
        d = res[b] || { b_no: b, b_stt: '', b_stt_cd: '', tax_type: '국세청에 등록되지 않은 사업자등록번호입니다.', end_dt: '' };
        cache.put('nts_' + b, JSON.stringify(d), 21600);
      } catch (e) { unavailable = e.message; }
    }
  }
  const kind = bizKind(b);
  if (unavailable) return { ok: true, allowed: true, unverified: true, kind, msg: '지금은 국세청 확인을 할 수 없어 담당자가 나중에 확인합니다. 사업자 유형을 직접 고르고 계속 작성해 주세요.' };
  const cd = d.b_stt_cd || '', tk = taxKind(d.tax_type);
  const o = { ok: true, stt: d.b_stt || '', cd, tax: d.tax_type || '', end: d.end_dt || '', kind, tk, demo, raw: { [b]: d } };
  const endTxt = o.end ? ' (폐업일 ' + o.end.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3') + ')' : '';
  if (cd === '03') return Object.assign(o, { allowed: false, msg: '국세청 확인 결과 폐업한 사업자' + endTxt + '라 신청할 수 없습니다.' });
  if (cd === '02' && !R.allowSuspended) return Object.assign(o, { allowed: false, msg: '국세청 확인 결과 휴업 중인 사업자라 이 사업에는 신청할 수 없습니다.' });
  if (cd !== '01' && cd !== '02') return Object.assign(o, { allowed: false, msg: '국세청에 등록되지 않은 번호입니다. 사업자등록번호를 다시 확인해 주세요.' });
  const doc = docFor(kind, tk || '일반');
  return Object.assign(o, { allowed: true, doc, docLabel: DOC_NAME[doc], incomeOk: !!R.allowIncome && kind === '개인', btype: btypeOf(kind, tk || '일반'),
    msg: (cd === '02' ? '휴업 중인 사업자입니다. ' : '') + kind + ' ' + (o.tax || '') + ' → 매출 서류는 ' + DOC_NAME[doc] + (R.allowIncome && kind === '개인' ? ' (또는 소득금액증명)' : '') + '을 올려 주세요.' });
}

/* 보완 제출: 링크(접수번호+토큰) 확인 */
function fixFind_(no, tok) {
  no = String(no || ''); tok = String(tok || '');
  if (!no || tok.length < 20) return null;
  const m = cols_(SH.A, AH);
  const r = findRow_(SH.A, no);
  if (r < 0) return null;
  const row = m._sh.getRange(r, 1, 1, m._n).getValues()[0];
  const g = k => row[m[k] - 1];
  if (!str_(g('보완토큰')) || str_(g('보완토큰')) !== tok) return null;
  if (!['보완 요청', '보완 제출'].includes(str_(row[8]))) return null;
  const p = readPrograms_().find(x => x.id === String(row[1]));
  if (!p) return null;
  const answers = answersOf_(row[14]);
  const slots = [];
  withConsent_(effItems(p)).forEach(s => s.items.forEach(it => {
    if (!it.on && !answers[it.key]) return;
    it.fields.filter(f => f.t === 'file').forEach(f => slots.push({ ik: it.key, fk: f.k, label: it.label + (it.fields.length > 1 ? ' · ' + f.label : '') }));
  }));
  const done = (answers.__fix || []).map(x => ({ at: x.at, files: x.files || [], memo: x.memo || '' }));
  return { no, r, m, row, p, answers, slots, done, biz: str_(row[4]), note: str_(g('보완요청내용')) };
}
function fixSubmit_(b) {
  const f = fixFind_(b.fix, b.t);
  if (!f) return { ok: false, error: '링크가 맞지 않거나 보완 제출 기간이 아닙니다. 담당자에게 문의해 주세요.' };
  const files = Array.isArray(b.files) ? b.files : [];
  const memo = String(b.memo || '').slice(0, 1000);
  if (!files.length && !memo.trim()) return { ok: false, error: '올릴 서류나 메모를 넣어 주세요.' };
  if (files.length > 20) return { ok: false, error: '파일은 한 번에 20개까지 올릴 수 있습니다.' };
  const allow = {};
  f.slots.forEach(x => allow[x.ik + '.' + x.fk] = true);
  let total = 0;
  for (const x of files) {
    if (!allow[x.ik + '.' + x.fk]) return { ok: false, error: '서류 항목을 확인할 수 없습니다.' };
    const ext = String(x.name || '').split('.').pop().toLowerCase();
    if (!OK_EXT.includes(ext)) return { ok: false, error: '올릴 수 없는 파일 형식입니다: ' + x.name };
    const size = Math.floor(String(x.data || '').length * 3 / 4);
    if (size > MAX_FILE) return { ok: false, error: '파일 하나는 10MB 이하여야 합니다: ' + x.name };
    total += size;
  }
  if (total > MAX_TOTAL) return { ok: false, error: '파일 합계는 25MB 이하여야 합니다.' };
  // 같은 신청 폴더 안에 '보완_날짜' 폴더를 만들어 저장
  const fid = (/folders\/([\w-]+)/.exec(str_(f.row[13])) || [])[1];
  const base = fid ? DriveApp.getFolderById(fid) : rootFolder_();
  const now = new Date();
  const sub = base.createFolder('보완_' + Utilities.formatDate(now, TZ, 'yyMMdd-HHmm'));
  const saved = files.map(x => {
    const file = sub.createFile(Utilities.newBlob(Utilities.base64Decode(x.data), x.type || 'application/octet-stream', String(x.name).slice(0, 120)));
    return { ik: x.ik, fk: x.fk, o: { id: file.getId(), name: file.getName(), type: x.type || '', size: file.getSize(), url: file.getUrl(), fix: true, at: now.getTime() } };
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const raw = String(f.m._sh.getRange(f.r, 15).getValue() || '{}');
    const answers = answersOf_(raw);
    saved.forEach(x => { answers[x.ik] = answers[x.ik] || {}; if (!Array.isArray(answers[x.ik][x.fk])) answers[x.ik][x.fk] = []; answers[x.ik][x.fk].push(x.o); });
    answers.__fix = (answers.__fix || []).concat([{ at: now.getTime(), memo, files: saved.map(x => x.o.name) }]);
    const json = JSON.stringify(answers);
    if (raw.indexOf('FILE:') === 0) DriveApp.getFileById(raw.slice(5)).setContent(json);
    else if (json.length > 45000) f.m._sh.getRange(f.r, 15).setValue('FILE:' + sub.createFile(Utilities.newBlob(json, 'application/json', '신청서_' + f.no + '_보완.json')).getId());
    else f.m._sh.getRange(f.r, 15).setValue(json);
    f.m._sh.getRange(f.r, 9).setValue('보완 제출');
    setCells_(f.m, f.r, { '보완제출일시': now, '처리상태': '대기', '처리시도': 0, '판독필요': 'Y' });
  } finally { lock.releaseLock(); }
  log_('신청자', f.no, '보완 제출', '', saved.map(x => x.o.name).join(', ') + (memo ? ' / ' + memo : ''));
  return { ok: true, no: f.no, count: saved.length };
}

/* 보완 제출 전용 링크 */
function fixLink_(no, tok) {
  const u = PropertiesService.getScriptProperties().getProperty('FORM_URL') || '';
  if (!u || !tok) return '';
  return u.replace(/[?#].*$/, '').replace(/\/?$/, '/') + '?fix=' + encodeURIComponent(no) + '&t=' + tok;
}

/* 위원별 면접 기록: {접수번호: [{interviewer, date, scores, notes, savedAt, by}]} */
function ivMap_() {
  const out = {};
  rows_(SH.I).forEach(r => {
    const o = { interviewer: str_(r[1]), date: str_(r[2]), scores: {}, notes: {}, savedAt: r[7] instanceof Date ? r[7].getTime() : 0, by: str_(r[8]) };
    try { o.scores = JSON.parse(r[5] || '{}'); } catch (e) {}
    try { o.notes = JSON.parse(r[6] || '{}'); } catch (e) {}
    (out[String(r[0])] = out[String(r[0])] || []).push(o);
  });
  Object.keys(out).forEach(k => out[k].sort((a, b) => a.savedAt - b.savedAt));
  return out;
}
function ivRow_(no, who) {
  const sh = ensureSheet_(SH.I, IH);
  if (sh.getLastRow() < 2) return -1;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === no && String(v[i][1]) === who) return i + 2;
  return -1;
}
/* 위원 평균을 신청목록에 기록 */
function ivSummary_(no, avg, n) {
  const m = cols_(SH.A, AH);
  const ra = findRow_(SH.A, no);
  if (ra > 0) { m._sh.getRange(ra, 13).setValue(avg == null ? '' : avg); setCells_(m, ra, { '면접위원수': n || '' }); }
}

function adminSaveInterview(no, iv, total, areas, orig, avg, n) {
  const me = requireStaff_();
  const who = String((iv && iv.interviewer) || '').trim().slice(0, 40);
  if (!who) throw new Error('면접관(위원) 이름을 넣어 주세요.');
  const sh = ensureSheet_(SH.I, IH);
  const r = ivRow_(no, orig == null ? who : String(orig));
  const dup = ivRow_(no, who);
  if (dup > 0 && dup !== r) throw new Error('같은 이름의 위원 기록이 이미 있습니다: ' + who);
  const row = [no, who, iv.date || '', total == null ? '' : total, JSON.stringify(areas || {}), JSON.stringify(iv.scores || {}), JSON.stringify(iv.notes || {}), new Date(), me];
  const before = r > 0 ? str_(sh.getRange(r, 4).getValue()) : '';
  if (r < 0) sh.appendRow(row); else sh.getRange(r, 1, 1, row.length).setValues([row]);
  ivSummary_(no, avg, n);
  log_(me, no, '면접 기록(' + who + ')', before, total == null ? '' : total);
  return { ivs: ivMap_()[no] || [] };
}

function adminDeleteInterview(no, who, avg, n) {
  const me = requireStaff_();
  const r = ivRow_(no, String(who));
  if (r > 0) sheet_(SH.I).deleteRow(r);
  ivSummary_(no, avg, n);
  log_(me, no, '면접 기록 삭제', who, '');
  return { ivs: ivMap_()[no] || [] };
}

/* 최종 결과(선정·예비·탈락) 저장 */
function adminSaveResults(pid, list, sel) {
  const me = requireStaff_();
  const m = cols_(SH.A, AH);
  const ids = colVals_(m._sh, 1);
  const now = new Date();
  (list || []).forEach(x => {
    const i = ids.indexOf(String(x.no));
    if (i < 0) return;
    const v = ['선정', '예비', '탈락'].includes(x.v) ? x.v : '';
    const before = str_(m._sh.getRange(i + 2, m['최종결과']).getValue());
    if (before === v) return;
    setCells_(m, i + 2, { '최종결과': v, '최종결과일시': now, '최종결과자': me });
    log_(me, x.no, '최종 결과', before, v);
  });
  if (sel) {
    const sp = sheet_(SH.P);
    const pi = colVals_(sp, 1).indexOf(pid);
    if (pi >= 0) {
      let cfg = {};
      try { cfg = JSON.parse(sp.getRange(pi + 2, 10).getValue() || '{}'); } catch (e) {}
      cfg.sel = { n: String(sel.n || '').replace(/\D/g, ''), m: String(sel.m || '').replace(/\D/g, '') };
      sp.getRange(pi + 2, 10).setValue(JSON.stringify(cfg));
    }
  }
  return { at: now.getTime(), by: me };
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

/* 한 신청의 서류를 읽어 서류판독·판독_* 시트에 저장하고 결과를 돌려줌 */
function ocrApp_(no, answers, who) {
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
  const sh = ensureSheet_(SH.O, OH);
  const ids = colVals_(sh, 1);
  for (let i = ids.length - 1; i >= 0; i--) if (ids[i] === no) sh.deleteRow(i + 2);
  const now = new Date();
  out.forEach(o => sh.appendRow([no, o.ik, o.fk, o.name, o.id, o.len, o.err ? o.err : cell_(o.text.slice(0, 45000)), now, who]));
  // 서류 종류별 시트에 항목별로 정리해 보관
  Object.keys(PARSE_DOCS).forEach(ik => {
    const d = PARSE_DOCS[ik];
    const head = ['접수번호', '파일명'].concat(d.fields, ['판독일시']);
    const pm = cols_('판독_' + d.kind, head); // 칸이 늘어나도 이름으로 맞춰 씀
    const ps = pm._sh;
    const pids = colVals_(ps, 1);
    for (let i = pids.length - 1; i >= 0; i--) if (pids[i] === no) ps.deleteRow(i + 2);
    out.filter(o => o.ik === ik && !o.err && o.len > 0).forEach(o => {
      const v = parseDoc(ik, o.text);
      const row = new Array(pm._n).fill('');
      const put = (k, x) => { if (pm[k]) row[pm[k] - 1] = x; };
      put('접수번호', no); put('파일명', o.name); put('판독일시', now);
      d.fields.forEach(f => put(f, cell_(v[f] || '')));
      ps.appendRow(row);
    });
  });
  return out.map(o => ({ ik: o.ik, fk: o.fk, name: o.name, len: o.len, text: o.text.slice(0, 20000), err: o.err || '', at: now.getTime() }));
}

function adminOcr(no) {
  const me = requireStaff_();
  const r = findRow_(SH.A, no);
  if (r < 0) throw new Error('접수번호를 찾을 수 없습니다: ' + no);
  const out = ocrApp_(no, answersOf_(sheet_(SH.A).getRange(r, 15).getValue()), me);
  let auto = null;
  try { auto = refreshVerdicts_(x => String(x[0]) === no)[no] || null; } catch (e) {}
  return { ocr: out, auto };
}

/* ---------- 사업자 상태 조회 (국세청 공공데이터 API) ----------
 * 공공데이터포털 인증키(Decoding)를 스크립트 속성 NTS_KEY 에 넣어야 동작함 */
function ntsFetch_(bnos) {
  const key = PropertiesService.getScriptProperties().getProperty('NTS_KEY');
  if (!key) throw new Error('인증키가 없습니다. Apps Script 프로젝트 설정 → 스크립트 속성에 NTS_KEY 를 넣어 주세요.');
  const k = key.replace(/\s/g, '');
  const forms = /%[0-9A-Fa-f]{2}/.test(k) ? [k, encodeURIComponent(decodeURIComponent(k))] : [encodeURIComponent(k), k];
  const res = {};
  const uniq = bnos.filter((b, i) => b.length === 10 && bnos.indexOf(b) === i);
  for (let i = 0; i < uniq.length; i += 100) {
    const chunk = uniq.slice(i, i + 100);
    let resp;
    for (const f of forms) {
      resp = UrlFetchApp.fetch('https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey=' + f, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ b_no: chunk })
      });
      if (resp.getResponseCode() !== 401) break;
    }
    const code = resp.getResponseCode();
    if (code === 401) throw new Error('인증키가 아직 등록되지 않았습니다. 발급 직후에는 1~2시간 걸릴 수 있습니다. 스크립트 속성 NTS_KEY 값(일반 인증키)도 확인해 주세요.');
    if (code !== 200) throw new Error('국세청 조회 실패(' + code + '): ' + resp.getContentText().slice(0, 200));
    (JSON.parse(resp.getContentText()).data || []).forEach(d => { res[d.b_no] = d; });
  }
  return res;
}
function ntsSave_(list, res, who) {
  const sh = ensureSheet_(SH.N, NH);
  const ids = colVals_(sh, 1);
  const now = new Date();
  const out = {};
  list.forEach(x => {
    const d = res[x.b];
    const o = x.b.length !== 10 ? { bno: x.b, stt: '', cd: '', tax: '사업자번호 자릿수 오류', end: '' }
      : d ? { bno: x.b, stt: d.b_stt || '', cd: d.b_stt_cd || '', tax: d.tax_type || '', end: d.end_dt || '' }
      : { bno: x.b, stt: '', cd: '', tax: '조회 결과 없음', end: '' };
    const tx = v => v ? "'" + v : '';
    const row = [x.no, tx(x.b), o.stt, tx(o.cd), o.tax, tx(o.end), now, who];
    const i = ids.indexOf(x.no);
    if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]); else { sh.appendRow(row); ids.push(x.no); }
    out[x.no] = Object.assign(o, { at: now.getTime() });
  });
  return out;
}
function adminBizStatus(nos) {
  const me = requireStaff_();
  const want = {};
  (nos || []).forEach(n => want[String(n)] = true);
  const list = rows_(SH.A).filter(r => want[String(r[0])]).map(r => ({ no: String(r[0]), b: String(((answersOf_(r[14]).biz || {}).brn) || '').replace(/[^0-9]/g, '') }));
  const out = ntsSave_(list, ntsFetch_(list.map(x => x.b)), me);
  let auto = {};
  try { auto = refreshVerdicts_(r => want[String(r[0])]); } catch (e) {}
  return { nts: out, auto };
}

/* 자동 처리 다시 하기: 처리상태를 '대기'로 되돌림 */
function adminRequeue(nos) {
  const me = requireStaff_();
  const m = cols_(SH.A, AH);
  const ids = colVals_(m._sh, 1);
  (nos || []).forEach(no => { const i = ids.indexOf(String(no)); if (i >= 0) setCells_(m, i + 2, { '처리상태': '대기', '처리시도': 0, '처리메모': '', '판독필요': 'Y' }); });
  log_(me, (nos || []).join(','), '자동 처리 다시 요청', '', '');
  return true;
}

/* ---------- 자동 점검 결과 계산·기록 ---------- */
function verdictCtx_() {
  const m = cols_(SH.A, AH);
  const sh = m._sh;
  const n = sh.getLastRow() - 1;
  const vals = n > 0 ? sh.getRange(2, 1, n, m._n).getValues() : [];
  const ocr = {};
  rows_(SH.O).forEach(r => {
    const no = String(r[0]); const len = Number(r[5]) || 0;
    (ocr[no] = ocr[no] || []).push({ ik: str_(r[1]), fk: str_(r[2]), name: str_(r[3]), len, text: len < 0 ? '' : String(r[6] || ''), err: len < 0 ? String(r[6] || '') : '' });
  });
  const nts = {};
  rows_(SH.N).forEach(r => { const t = v => str_(v).replace(/^'/, ''); nts[String(r[0])] = { bno: t(r[1]), stt: str_(r[2]), cd: t(r[3]), tax: str_(r[4]), end: t(r[5]) }; });
  const progs = {};
  readPrograms_().forEach(p => progs[p.id] = p);
  const ans = {};
  const answersAt = i => (ans[i] = ans[i] || answersOf_(vals[i][14]));
  const brnAt = i => { const v = String(vals[i][m['사업자번호'] - 1] || '').replace(/[^0-9]/g, ''); return v || String(((answersAt(i).biz || {}).brn) || '').replace(/[^0-9]/g, ''); };
  return { m, sh, vals, ocr, nts, progs, answersAt, brnAt };
}
function verdictOf_(ctx, i) {
  const r = ctx.vals[i];
  const no = String(r[0]), pid = String(r[1]);
  const p = ctx.progs[pid];
  if (!p) return null;
  const all = [];
  ctx.vals.forEach((x, j) => { if (String(x[1]) === pid) all.push({ id: String(x[0]), answers: { biz: { brn: ctx.brnAt(j) } } }); });
  const a = { id: no, answers: ctx.answersAt(i), ocr: ctx.ocr[no] || null, nts: ctx.nts[no] || null };
  const res = autoCheck(a, p, all);
  const s = autoSum(res);
  const reason = res.filter(x => x.k === 'fail' || x.k === 'warn').map(x => (x.k === 'fail' ? '[탈락후보] ' : '[확인필요] ') + x.label + ': ' + x.detail).join('\n');
  return { verdict: s.verdict, f: s.f, w: s.w, reason };
}
function storeVerdict_(ctx, i, v) {
  setCells_(ctx.m, i + 2, { '자동판정': v.verdict, '탈락후보': v.f, '확인필요': v.w, '자동사유': cell_(v.reason.slice(0, 45000)) });
}
/* 조건에 맞는 신청의 자동 점검 결과를 다시 계산해 기록 */
function refreshVerdicts_(filterFn) {
  const ctx = verdictCtx_();
  const out = {};
  ctx.vals.forEach((r, i) => {
    if (!filterFn(r)) return;
    const v = verdictOf_(ctx, i);
    if (v) { storeVerdict_(ctx, i, v); out[String(r[0])] = v; }
  });
  return out;
}

/* ---------- 자동 처리 (시간 트리거) ---------- */
function autoOn() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'autoRun').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('autoRun').timeBased().everyMinutes(5).create();
  PropertiesService.getScriptProperties().setProperty('AUTO_ON', 'Y');
  try { SpreadsheetApp.getUi().alert('자동 처리를 켰습니다. 5분마다 새 신청의 서류를 읽고 사업자 상태를 조회합니다.'); } catch (e) {}
}
function autoOff() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'autoRun').forEach(t => ScriptApp.deleteTrigger(t));
  PropertiesService.getScriptProperties().setProperty('AUTO_ON', 'N');
  try { SpreadsheetApp.getUi().alert('자동 처리를 껐습니다.'); } catch (e) {}
}
/* 처리상태가 비었거나 '대기', 또는 '오류'인데 시도 횟수가 남은 신청을 순서대로 처리 (한 번에 4분 30초까지) */
function autoRun() {
  const lock = LockService.getDocumentLock() || LockService.getUserLock(); // 제출용 잠금과 따로 씀
  if (!lock.tryLock(1000)) return;
  const t0 = Date.now(), BUDGET = 270 * 1000;
  const props = PropertiesService.getScriptProperties();
  try {
    props.setProperty('AUTO_LAST', nowStr_());
    // 처리할 신청이 없으면 다른 시트는 읽지 않고 끝냄
    const m0 = cols_(SH.A, AH);
    if (m0._sh.getLastRow() < 2) return;
    const st0 = m0._sh.getRange(2, m0['처리상태'], m0._sh.getLastRow() - 1, 2).getValues();
    if (!st0.some(r => { const st = String(r[0] || ''); return st === '' || st === '대기' || (st === '오류' && (Number(r[1]) || 0) < AUTO_MAX_TRY); })) return;
    const ctx = verdictCtx_();
    const m = ctx.m;
    const pend = [];
    ctx.vals.forEach((r, i) => {
      const st = String(r[m['처리상태'] - 1] || ''), tries = Number(r[m['처리시도'] - 1]) || 0;
      if (st === '' || st === '대기' || (st === '오류' && tries < AUTO_MAX_TRY)) pend.push(i);
    });
    if (!pend.length) return;
    const notes = {};
    // 1) 국세청 사업자 상태: 한꺼번에 조회
    if (props.getProperty('NTS_KEY')) {
      try {
        const list = pend.filter(i => !ctx.nts[String(ctx.vals[i][0])] && !((ctx.answersAt(i).__biz || {}).demo)).map(i => ({ no: String(ctx.vals[i][0]), b: ctx.brnAt(i) }));
        if (list.length) {
          const saved = ntsSave_(list, ntsFetch_(list.map(x => x.b)), '자동');
          Object.keys(saved).forEach(no => ctx.nts[no] = saved[no]);
        }
      } catch (e) { pend.forEach(i => notes[i] = '사업자 상태 조회 실패: ' + e.message); }
    } else pend.forEach(i => notes[i] = '국세청 인증키(NTS_KEY) 없음');
    // 2) 서류 글자 읽기 + 자동 점검
    const driveOn = typeof Drive !== 'undefined';
    for (const i of pend) {
      if (Date.now() - t0 > BUDGET) break;
      const no = String(ctx.vals[i][0]);
      const tries = (Number(ctx.vals[i][m['처리시도'] - 1]) || 0) + 1;
      try {
        const need = String(ctx.vals[i][m['판독필요'] - 1] || '') === 'Y';
        if (driveOn && (need || !ctx.ocr[no])) {
          const o = ocrApp_(no, ctx.answersAt(i), '자동');
          ctx.ocr[no] = o;
        }
        const bn = ctx.brnAt(i);
        const v = verdictOf_(ctx, i);
        if (v) storeVerdict_(ctx, i, v);
        setCells_(m, i + 2, { '처리상태': '완료', '처리시도': tries, '처리일시': new Date(), '판독필요': driveOn ? '' : String(ctx.vals[i][m['판독필요'] - 1] || ''), '처리메모': [driveOn ? '' : 'Drive API 꺼짐(서류 읽기 안 함)', notes[i] || ''].filter(Boolean).join(' / '), '사업자번호': bn ? "'" + bn : '' });
      } catch (e) {
        setCells_(m, i + 2, { '처리상태': tries >= AUTO_MAX_TRY ? '오류' : '대기', '처리시도': tries, '처리일시': new Date(), '처리메모': cell_(String(e.message).slice(0, 300)) });
      }
    }
  } finally { lock.releaseLock(); }
}
