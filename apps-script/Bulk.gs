/* ---------- 사업자번호 일괄조회 ----------
 * '일괄조회' 시트 A열에 사업자번호 붙여넣기 → 메뉴 [사업자번호 일괄조회 실행]
 * 국세청 상태조회 API (100개씩) → 결과·분류를 같은 시트에, 건수 요약은 '일괄조회_요약' 시트에 */
const BULK_SH = '일괄조회', BULK_SUM = '일괄조회_요약';
const BULK_HEAD = ['사업자번호(입력)', '사업자번호', '사업자 구분', '상태', '과세유형', '분류', '폐업일',
  '과세유형 전환일', '직전 과세유형', '세금계산서 적용일', '단위과세 전환폐업', '비고', '조회일시'];

function bulkKind_(b) {
  const m = Number(b.slice(3, 5));
  if ([81, 86, 87, 88].indexOf(m) >= 0) return '법인(영리 본점)';
  if (m === 85) return '법인(영리 지점)';
  if (m === 82) return '법인(비영리)';
  if (m === 83) return '국가·지자체';
  if (m === 84) return '외국법인';
  if (m === 89) return '법인 아닌 단체';
  if (m === 80) return '개인(비영리 등)';
  if (m >= 90) return '개인(면세)';
  return '개인';
}
function bulkValid_(b) {
  if (!/^\d{10}$/.test(b)) return false;
  const w = [1, 3, 7, 1, 3, 7, 1, 3, 5], d = b.split('').map(Number);
  let s = 0;
  for (let i = 0; i < 9; i++) s += d[i] * w[i];
  s += Math.floor(d[8] * 5 / 10);
  return (10 - s % 10) % 10 === d[9];
}
function bulkDate_(v) { v = String(v || ''); return /^\d{8}$/.test(v) ? v.slice(0, 4) + '-' + v.slice(4, 6) + '-' + v.slice(6) : v; }
function bulkTax_(t) {
  t = String(t || '');
  if (/간이/.test(t)) return '간이과세자';
  if (/면세/.test(t)) return '면세사업자';
  if (/비영리|국가|고유번호/.test(t)) return '비영리·고유번호';
  if (/일반/.test(t)) return '일반과세자';
  return t ? '기타' : '';
}

function bulkRun() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(BULK_SH);
  if (!sh) {
    sh = ss.insertSheet(BULK_SH);
    sh.getRange(1, 1, 1, BULK_HEAD.length).setValues([BULK_HEAD]).setFontWeight('bold').setBackground('#eef0f5');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
    ui.alert("'일괄조회' 시트를 만들었습니다.\nA열 2행부터 사업자번호를 붙여넣고 메뉴를 다시 실행해 주세요.");
    return;
  }
  const last = sh.getLastRow();
  if (last < 2) { ui.alert('A열 2행부터 사업자번호를 붙여넣어 주세요.'); return; }
  const inputs = sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => String(r[0]).trim());

  const rows = inputs.map(v => ({ raw: v, b: v.replace(/[^0-9]/g, '') }));
  const first = {};
  rows.forEach((r, i) => {
    r.note = [];
    if (!r.raw) { r.skip = true; return; }
    if (r.b.length !== 10) { r.err = '자릿수 오류 (' + r.b.length + '자리)'; return; }
    if (!bulkValid_(r.b)) { r.err = '번호 형식 오류 (검증번호 불일치)'; return; }
    if (first[r.b] !== undefined) r.note.push('중복 (' + (first[r.b] + 2) + '행과 같음)'); else first[r.b] = i;
  });
  const targets = Object.keys(first);
  const calls = Math.ceil(targets.length / 100);
  if (ui.alert('사업자번호 ' + targets.length + '개를 조회합니다 (국세청 호출 ' + calls + '회).\n계속할까요?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const res = {};
  for (let i = 0; i < targets.length; i += 100) {
    if (!ntsQuota_()) { ui.alert('오늘 국세청 조회 한도를 넘었습니다. ' + i + '개까지만 조회했습니다.'); break; }
    Object.assign(res, ntsFetch_(targets.slice(i, i + 100)));
  }

  const now = new Date();
  const out = rows.map(r => {
    if (r.skip) return BULK_HEAD.map(() => '').slice(1);
    const kind = r.b.length === 10 ? bulkKind_(r.b) : '';
    if (r.err) return [r.b, kind, '조회 불가', '', '', '', '', '', '', '', r.err, now];
    const d = res[r.b];
    if (!d) return [r.b, kind, '미조회', '', '', '', '', '', '', '', r.note.concat('조회되지 않음 (한도 초과 등)').join(' / '), now];
    const unreg = !d.b_stt_cd;
    const stt = unreg ? '미등록' : ({ '01': '계속', '02': '휴업', '03': '폐업' })[d.b_stt_cd] || String(d.b_stt || '');
    const tax = unreg ? '' : bulkTax_(d.tax_type);
    const cls = unreg ? '미등록' : stt + ' · ' + tax;
    if (unreg) r.note.push('국세청에 등록되지 않은 번호');
    return [r.b, kind, stt, unreg ? '' : d.tax_type, cls, bulkDate_(d.end_dt), bulkDate_(d.tax_type_change_dt),
      d.rbf_tax_type || '', bulkDate_(d.invoice_apply_dt), d.utcc_yn === 'Y' ? '예' : '', r.note.join(' / '), now];
  });

  const n = out.length, w = BULK_HEAD.length - 1;
  sh.getRange(1, 1, 1, BULK_HEAD.length).setValues([BULK_HEAD]).setFontWeight('bold').setBackground('#eef0f5');
  sh.setFrozenRows(1);
  const lc = Math.max(sh.getLastColumn(), BULK_HEAD.length);
  const f0 = sh.getFilter(); if (f0) f0.remove();
  if (lc > BULK_HEAD.length) sh.getRange(1, BULK_HEAD.length + 1, 1, lc - BULK_HEAD.length).clearContent().setBackground(null);
  if (sh.getMaxRows() > 1) sh.getRange(2, 2, sh.getMaxRows() - 1, lc - 1).clearContent().setBackground(null);
  sh.getRange(2, 2, n, 1).setNumberFormat('@');
  sh.getRange(2, 7, n, 4).setNumberFormat('@');
  sh.getRange(2, 2, n, w).setValues(out);
  const color = { '계속': '#e3f6ef', '휴업': '#fdf0e1', '폐업': '#fbe6e6', '미등록': '#fbe6e6', '조회 불가': '#eef0f5', '미조회': '#eef0f5' };
  sh.getRange(2, 4, n, 1).setBackgrounds(out.map(r => [color[r[2]] || null]));
  sh.getRange(1, 1, n + 1, BULK_HEAD.length).createFilter();

  bulkSummary_(ss, out.filter((r, i) => !rows[i].skip), now);
  ui.alert('완료: ' + out.filter((r, i) => !rows[i].skip).length + '건 정리.\n요약은 \'' + BULK_SUM + '\' 시트에 있습니다.');
}

function bulkSummary_(ss, out, now) {
  let sh = ss.getSheetByName(BULK_SUM);
  if (!sh) sh = ss.insertSheet(BULK_SUM); else sh.clear();
  const cnt = (f) => out.filter(f).length;
  const N = out.length;
  const pct = (k, base) => (base ? (k / base * 100).toFixed(1) : '0.0') + '%';
  const blocks = [];
  blocks.push([['일괄조회 요약', Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') + ' 기준']]);
  const stts = ['계속', '휴업', '폐업', '미등록', '조회 불가', '미조회'];
  blocks.push([['상태', '건수', '비율'], ...stts.filter(k => k !== '미조회' || cnt(r => r[2] === k)).map(k => { const c = cnt(r => r[2] === k); return [k, c, pct(c, N)]; }), ['합계', N, '100%']]);
  const taxes = ['일반과세자', '간이과세자', '면세사업자', '비영리·고유번호', '기타'];
  const t = [['상태 \\ 과세유형', ...taxes, '계']];
  ['계속', '휴업', '폐업'].forEach(st => t.push([st, ...taxes.map(x => cnt(r => r[2] === st && bulkTax_(r[3]) === x)), cnt(r => r[2] === st)]));
  t.push(['계', ...taxes.map(x => cnt(r => ['계속', '휴업', '폐업'].indexOf(r[2]) >= 0 && bulkTax_(r[3]) === x)), cnt(r => ['계속', '휴업', '폐업'].indexOf(r[2]) >= 0)]);
  const M = t[t.length - 1][taxes.length + 1];
  t.push(['비율', ...taxes.map((x, i) => pct(t[t.length - 1][i + 1], M)), '100%']);
  blocks.push(t);
  const kinds = [];
  out.forEach(r => { if (r[1] && kinds.indexOf(r[1]) < 0) kinds.push(r[1]); });
  blocks.push([['사업자 구분', '건수', '비율', '그중 계속', '계속 비율'], ...kinds.map(k => { const a = cnt(r => r[1] === k), b = cnt(r => r[1] === k && r[2] === '계속'); return [k, a, pct(a, N), b, pct(b, a)]; })]);
  const notes = {};
  out.forEach(r => String(r[10] || '').split(' / ').forEach(x => { if (x) { const k = x.replace(/\(.*\)/, '').trim(); notes[k] = (notes[k] || 0) + 1; } }));
  if (Object.keys(notes).length) blocks.push([['확인할 것', '건수'], ...Object.keys(notes).map(k => [k, notes[k]])]);

  let row = 1;
  blocks.forEach((b, i) => {
    const wd = Math.max(...b.map(x => x.length));
    const v = b.map(x => x.concat(Array(wd - x.length).fill('')));
    sh.getRange(row, 1, v.length, wd).setValues(v);
    sh.getRange(row, 1, 1, wd).setFontWeight('bold');
    if (i > 0) sh.getRange(row, 1, 1, wd).setBackground('#eef0f5');
    row += v.length + 1;
  });
  sh.setColumnWidth(1, 160);
}
