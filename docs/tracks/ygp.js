/* 정산 모듈: YGP (와이지플러스) MD 로열티 · 응원봉 디자인권 사용료 (월, 매출)
 * 판별(detect) → 셸이 이 정산의 비공개 설정을 서버에서 받아 mount 호출 → 정산·품의·ERP 탭 렌더.
 * 회사 데이터(아티스트→ERP 코드 매핑·거래처·품의 문구)는 이 파일에 없다(DB "YGP_설정" 시트). */
(function () {
/* YGP 정산 엔진 — YGP_MD_응원봉_ERP업로드_자동화.xlsx 로직 이관
 * (26.7·8·9월 실제 AR 업로드 파일과 줄·금액·적요·코드 일치 검증) */
var YGP = (function () {
  var N = function (s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); };
  var K = function (s) { return N(s).replace(/\s+/g, ''); };
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  function num(v) { var n = typeof v === 'number' ? v : parseFloat(String(v == null ? '' : v).replace(/[,\s]/g, '')); return isFinite(n) ? n : null; }
  function fmt(n) { return Math.round(n).toLocaleString('ko-KR'); }
  function rowsOf(ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }); }
  function ymd(v) {
    if (typeof v === 'number') { var d = new Date(Math.round((v - 25569) * 864e5)); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
    var m = String(v || '').match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/); return m ? m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]) : '';
  }
  function monthEnd(ym) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7); return ym + '-' + pad(new Date(Date.UTC(y, m, 0)).getUTCDate()); }
  function addMonths(ym, n) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + n; y += Math.floor(m / 12); m = ((m % 12) + 12) % 12; return y + '-' + pad(m + 1); }
  var SHEET = '항목별정산내역_요약';

  // 1) YGP 정산 원본(Ent 정산 xlsb / 응원봉 디자인권 정산 xlsx)의 "항목별정산내역_요약" 시트 해석
  function parse(wb) {
    var ws = wb.Sheets[SHEET];
    if (!ws) throw new Error('"' + SHEET + '" 시트가 없습니다. YGP 정산 원본(Ent 정산 / 응원봉 디자인권 정산)인지 확인해 주세요.');
    var rows = rowsOf(ws), title = N((rows[0] || [])[0]);
    var kind = /응원봉/.test(title) ? 'cb' : /YG\s*Ent/i.test(title) ? 'md' : '';
    if (!kind) throw new Error('"' + SHEET + '" 첫 줄 제목에서 MD 로열티/응원봉 정산을 구분하지 못했습니다: ' + title);
    var tm = title.match(/(\d{4})\.(\d{1,2})/);
    if (!tm) throw new Error('제목에서 정산월(예: 2026.9)을 찾지 못했습니다: ' + title);
    var p = { kind: kind, title: title, ym: tm[1] + '-' + pad(+tm[2]), artists: [], actorTotal: 0 };
    var cell = function (r, i) { return N((rows[r] || [])[i]); };
    for (var i = 0; i < rows.length; i++) {
      if (cell(i, 0) === '작성일자' && !p.invoice) {
        var r = rows[i + 1] || [];
        p.invoice = { date: ymd(r[0]), memo: N(r[1]), supply: num(r[3]), vat: num(r[4]), total: num(r[5]) };
      }
      if (/②\s*아티스트별/.test(cell(i, 0)) && !p.artistAt) p.artistAt = i;
      if (/^-\s*배우/.test(cell(i, 0))) p.actorAt = i;
    }
    if (!p.invoice || p.invoice.supply == null) throw new Error('"① YG ENT 발행내역"(작성일자·공급가액·세액) 표를 찾지 못했습니다.');
    if (p.artistAt == null) throw new Error('"② 아티스트별" 표를 찾지 못했습니다.');
    // ② 아티스트별 - 가수: 머리글(아티스트) 다음 줄부터 "합계"까지 (B 아티스트, C 수량, D 공급가액, E 부가세, F 공급대가)
    var h = p.artistAt; while (h < rows.length && cell(h, 1) !== '아티스트') h++;
    for (var j = h + 1; j < rows.length; j++) {
      if (cell(j, 0) === '합계' || cell(j, 1) === '합계') { p.artistTotal = { qty: num(rows[j][2]), supply: num(rows[j][3]), vat: num(rows[j][4]), total: num(rows[j][5]) }; break; }
      var name = cell(j, 1); if (!name) continue;
      p.artists.push({ name: name, qty: num(rows[j][2]) || 0, supply: num(rows[j][3]) || 0, vat: num(rows[j][4]) || 0, total: num(rows[j][5]) || 0 });
    }
    if (p.actorAt != null) {
      for (var k = p.actorAt + 1; k < rows.length; k++) if (cell(k, 0) === '합계' || cell(k, 1) === '합계') { p.actorTotal = num(rows[k][3]) || 0; break; }
    }
    if (!p.artists.length) throw new Error('"② 아티스트별" 표에 아티스트 행이 없습니다.');
    return p;
  }

  // 2) 계산·검증. MD: 매핑표 순서·코드 / 응원봉: 원본 순서, 프로젝트·그룹 공통("_")
  function compute(p, cfg) {
    var map = (p.kind === 'md' ? cfg.MD_MAP : cfg.CB_MAP) || [], flags = [], byKey = {};
    map.forEach(function (m) { byKey[K(m.src)] = m; });
    var sum = function (k) { return p.artists.reduce(function (a, x) { return a + x[k]; }, 0); };
    var S = { qty: sum('qty'), supply: sum('supply'), vat: sum('vat'), total: sum('total') }, inv = p.invoice;
    if (S.supply === inv.supply && S.vat === inv.vat) flags.push({ level: 'ok', text: '② 아티스트별 ' + p.artists.length + '행 합계 = ① 발행내역: 공급가액 ' + fmt(inv.supply) + ' · 부가세 ' + fmt(inv.vat) + ' · 합계 ' + fmt(inv.total) + '원' });
    else flags.push({ level: 'bad', text: '② 아티스트별 합계(공급 ' + fmt(S.supply) + ' / 부가세 ' + fmt(S.vat) + ')와 ① 발행내역(공급 ' + fmt(inv.supply) + ' / 부가세 ' + fmt(inv.vat) + ')이 다릅니다 — 정산서 확인 필요' });
    if (p.artistTotal && (p.artistTotal.supply !== S.supply || p.artistTotal.vat !== S.vat)) flags.push({ level: 'bad', text: '② 아티스트별 "합계" 행(' + fmt(p.artistTotal.supply) + ')이 행 합계(' + fmt(S.supply) + ')와 다릅니다' });
    if (inv.supply + inv.vat !== inv.total) flags.push({ level: 'bad', text: '발행내역 공급가액 + 세액 ≠ 합계' });
    var v10 = Math.round(inv.supply * 0.1);
    flags.push({ level: 'ok', text: '세액 ' + fmt(inv.vat) + '원은 정산서 값 그대로 사용 (공급가액 × 10% = ' + fmt(v10) + '원' + (inv.vat === v10 ? ', 일치' : ', 아티스트별 반올림 합이라 ' + (inv.vat - v10 > 0 ? '+' : '') + (inv.vat - v10) + '원 차이 — 정상') + ')' });
    if (p.actorTotal) flags.push({ level: 'warn', text: '"-배우" 표 합계가 ' + fmt(p.actorTotal) + '원입니다 — 보통 0원이라 확인 필요(가수 표만 반영됨)' });
    var big = p.artists.filter(function (x) { return /빅뱅/.test(x.name) && x.supply; })[0];
    if (big) flags.push({ level: 'warn', text: '빅뱅 행(' + fmt(big.supply) + '원)이 있습니다 — 26.8월분부터 빅뱅 마스터 계약 체결 전이라 "빅뱅제외" 버전 정산서로 처리했습니다. 이 파일이 최종본인지 확인' });
    var lines = [], unmapped = [];
    if (p.kind === 'md') {
      var have = {}; p.artists.forEach(function (x) { have[K(x.name)] = x; });
      map.forEach(function (m) { var x = have[K(m.src)]; if (x && x.supply) lines.push({ src: x.name, label: m.label, amt: x.supply, proj: m.proj, grp: m.grp, art: m.art }); });
      p.artists.forEach(function (x) { if (!byKey[K(x.name)] && x.supply) unmapped.push(x); });
    } else {
      p.artists.forEach(function (x) { if (!x.supply) return; var m = byKey[K(x.name)]; if (!m) unmapped.push(x); lines.push({ src: x.name, label: m ? m.label : x.name, amt: x.supply }); });
    }
    if (unmapped.length) flags.push({ level: p.kind === 'md' ? 'bad' : 'warn', text: '매핑표에 없는 아티스트: ' + unmapped.map(function (x) { return x.name + ' ' + fmt(x.supply) + '원'; }).join(', ') +
      (p.kind === 'md' ? ' — ERP 프로젝트·그룹·아티스트 코드가 필요해 업로드에서 빠졌습니다. DB "YGP_설정" MD_MAP에 추가해 주세요' : ' — 원본 이름 그대로 적요에 넣었습니다. 표기 확인 필요') });
    var lineSum = lines.reduce(function (a, l) { return a + l.amt; }, 0);
    if (lineSum === inv.supply) flags.push({ level: 'ok', text: 'AR 업로드 ' + lines.length + '줄 대변 합계 ' + fmt(lineSum) + '원 = 공급가액 (0원 행 ' + p.artists.filter(function (x) { return !x.supply; }).length + '개 제외)' });
    else flags.push({ level: 'bad', text: 'AR 업로드 줄 합계 ' + fmt(lineSum) + '원 ≠ 공급가액 ' + fmt(inv.supply) + '원 (' + (lineSum - inv.supply > 0 ? '+' : '') + fmt(lineSum - inv.supply) + '원)' });
    return { S: S, lines: lines, flags: flags, unmapped: unmapped };
  }

  // 적요: MD "2026.09 YGP MD 로열티: 블랙핑크" / 응원봉 "2026.9 YGP 응원봉 디자인권 사용료: 트레저(10인)" (실제 AR 표기 그대로)
  function memo(p, l) {
    var y = p.ym.slice(0, 4), m = p.ym.slice(5, 7);
    return p.kind === 'md' ? y + '.' + m + ' YGP MD 로열티: ' + l.label : y + '.' + (+m) + ' YGP 응원봉 디자인권 사용료: ' + l.label;
  }
  function serial(d) { var t = d.split('-'); return Math.round((Date.UTC(+t[0], +t[1] - 1, +t[2]) - Date.UTC(1899, 11, 30)) / 86400000); }

  // ERP "AR 거래내역 일괄생성 업로드" 양식 ([md]/[응원]AR거래내역엑셀업로드_YGPMD 파일과 같은 열·값)
  function buildErpUpload(p, r, cfg) {
    var P = cfg.PROJECT || {}, start = p.ym + '-01', end = monthEnd(p.ym);
    var head = ['프로젝트', '회계부서', '계정코드', '예산코드', '배부코드', '통화', '환율', '차변금액', '대변금액', '거래처', '반제전표', '적요', '세금일자', '부가세코드', '공급가액', '부가세금액', '신고사업장', '세금계산서번호', '아티스트그룹', '아티스트', '관리항목1', '관리항목2', '관리항목3', '관리항목4', '관리항목5', '관리항목6', '관리항목7', '관리항목8'];
    var aoa = [['AR 거래내역 일괄생성 업로드 양식 // ERP자동전표_IP'], [], head];
    r.lines.forEach(function (l) {
      var md = p.kind === 'md';
      aoa.push([md ? +l.proj : '_', +P.dept, String(P.account), null, null, 'KRW', 1, null, l.amt, +P.vendor, null, memo(p, l), null, null, null, null, null, null,
        md ? +l.grp : '_', md ? +l.art : +P.artist,
        // 관리항목1·2: MD는 날짜(yyyy-mm-dd 서식), 응원봉은 8자리 숫자(20260901) — 각 실제 AR 파일 표기 그대로
        md ? serial(start) : +start.replace(/-/g, ''), md ? serial(end) : +end.replace(/-/g, ''), 'KR', null, null, null, null, null]);
    });
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    if (p.kind === 'md') for (var ri = 3; ri < aoa.length; ri++) [20, 21].forEach(function (ci) { var c = ws[XLSX.utils.encode_cell({ r: ri, c: ci })]; if (c && c.t === 'n') c.z = 'yyyy-mm-dd'; });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'AR거래내역엑셀업로드 List');
    return { wb: wb, rows: aoa.slice(3) };
  }

  return { parse: parse, compute: compute, memo: memo, buildErpUpload: buildErpUpload, fmt: fmt, monthEnd: monthEnd, addMonths: addMonths };
})();
if (typeof module !== 'undefined') module.exports = YGP;
if (typeof DESK === 'undefined') return;

// ---------------- 화면 ----------------
var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
var fmt = YGP.fmt;
function tabsHtml(tabs) {
  return '<div class="tabs-in" role="tablist" aria-label="작업 단계">' + tabs.map(function (t, i) {
    return '<button class="tab" role="tab" id="t-' + t[0] + '" data-tab="' + t[0] + '" aria-selected="' + (i ? 'false' : 'true') + '"><span>' + (i + 1) + '</span>' + t[1] + '</button>';
  }).join('') + '</div>';
}
function tableHtml(t) {
  var h = '<tr>' + t.head.map(function (c, i) { return '<th' + (t.numCols.indexOf(i) >= 0 ? ' class="n"' : '') + '>' + esc(c) + '</th>'; }).join('') + '</tr>';
  t.rows.forEach(function (r, ri) {
    var sum = t.sumLast && ri === t.rows.length - 1;
    h += '<tr' + (sum ? ' class="sum"' : '') + '>' + r.map(function (c, i) { return '<td' + (t.numCols.indexOf(i) >= 0 ? ' class="n"' : '') + '>' + esc(c) + '</td>'; }).join('') + '</tr>';
  });
  return h;
}
function kv(list, grp) {
  return list.map(function (x, i) { return '<dt>' + esc(x[0]) + '</dt><dd>' + esc(x[1]) + '</dd><button class="btn sm" data-copyval="' + grp + ':' + i + '">복사</button>'; }).join('');
}
function flagsHtml(fl) {
  return fl.map(function (f) { return '<div class="flag ' + f.level + '"><b>' + { ok: '확인', warn: '주의', bad: '조치' }[f.level] + '</b><span>' + esc(f.text) + '</span></div>'; }).join('');
}
function statusHtml(fl) {
  var nb = fl.filter(function (f) { return f.level === 'bad'; }).length, nw = fl.filter(function (f) { return f.level === 'warn'; }).length, lvl = nb ? 'bad' : nw ? 'warn' : 'ok';
  return { cls: 'status ' + lvl, html: '<b>' + (nb ? '조치 필요 ' + nb + '건' : nw ? '체크 ' + nw + '건' : '검증 이상 없음') + '</b><span>' + esc((fl.filter(function (f) { return f.level === lvl; })[0] || {}).text || '') + '</span>' };
}
// 위임전결기준표(261006) — 사업부문 매출 품의
function approval(supply) {
  if (supply < 5e6) return '500만원 미만 (합의종결): 기안 → 팀 리더(최종) · 합의: 재무관리팀 담당, 회계팀 담당';
  if (supply < 5e7) return '500만원 이상 5,000만원 미만 (합의종결): 기안 → 팀 리더 → 그룹 리더(최종) · 합의: 재무관리팀 담당, 회계팀 담당 → 회계팀 리더';
  if (supply < 3e8) return '5,000만원 이상 3억원 미만: 기안 → 팀 리더 → 그룹 리더 · 합의: 재무관리팀·회계팀 담당 → 회계팀 리더 → CFO → 부문 대표(최종)';
  return '3억원 이상: 기안 → 팀 리더 → 그룹 리더 · 합의: 재무관리팀·회계팀 담당 → 회계팀 리더 → CFO → 부문 대표 → 대표이사(최종)';
}

function mount(ctx) {
  var $ = function (id) { return document.getElementById(id); };
  var CFG = (ctx.data && ctx.data.cfg) || {}, P = CFG.PROJECT || {}, PUM = CFG.PUM || {};
  var p = YGP.parse(ctx.wb), R = YGP.compute(p, CFG), TXT = {}, TABLES = {};
  var md = p.kind === 'md', mNum = +p.ym.slice(5), inv = p.invoice;
  var key = (md ? 'MD ' : '응원봉 ') + p.ym, saved = (ctx.data.saved || {})[key];
  var title = '[매출품의] YGP MD 로열티 ' + mNum + '월 정산의 건' + (md ? '' : '_응원봉 디자인권 사용료');
  // 입금예정일: 거래일자(세금계산서 작성일)의 3개월 뒤 말일 — 26.8월분 품의(8/31 → 11/30) 기준
  var pay = inv.date ? YGP.monthEnd(YGP.addMonths(inv.date.slice(0, 7), PUM.payMonths || 3)) : '';

  ctx.tabsEl.innerHTML = tabsHtml([['calc', '정산'], ['pum', '품의 복붙'], ['erp', 'ERP 전표']]);
  ctx.root.innerHTML =
    '<div class="tabpanel" id="p-calc"><section class="panel dash"><div class="tiles" id="tiles"></div><div class="status" id="statusLine"></div><div class="meta" id="amtMeta"></div>' +
    '<details class="fold vfold" open><summary>정산서 숫자 검증</summary><div class="flags" id="flags"></div></details></section>' +
    '<section class="panel"><div class="ph"><h2>아티스트별 ' + (md ? 'MD 로열티' : '응원봉 사용료') + ' → ERP 줄</h2><p>원본 "항목별정산내역_요약" 시트 "② 아티스트별 - 가수" 표 기준입니다 ("제품매출" 시트 피벗은 특수채널 조정 전 숫자라 쓰지 않습니다).</p></div><div class="tablebox"><table id="lineTable"></table></div></section>' +
    '<details class="panel fold logic"><summary><b>적용 중인 정산 로직</b> <span class="small muted">YGP_MD_응원봉_ERP업로드_자동화.xlsx · 26.7~9월 실제 AR 기준</span></summary><ol>' +
    '<li><b>금액</b>: 정산서 "① YG ENT 발행내역"의 공급가액·세액·합계를 그대로 씁니다. 세액은 아티스트별 반올림 합이라 정확히 10%가 아닐 수 있습니다.</li>' +
    '<li><b>ERP 줄</b>: ' + (md ? '아티스트별 공급가액 = 대변금액. 매핑표 순서·코드(프로젝트/그룹/아티스트)로 변환. 원본 "BABYMONSTER" = ERP "BABYMONSTER_라미제외"(6인), "BABYMONSTER(라미포함)" = ERP "BABYMONSTER"(7인), "WINNER(위너)(민호제외)" = "WINNER(위너)_민호제외".' : '아티스트별 공급가액 = 대변금액. 원본 순서, 프로젝트·그룹 "_"(공통), 트레저 → "트레저(10인)".') + ' 0원 행 제외.</li>' +
    '<li><b>관리항목1·2</b>: 정산월 1일~말일 (' + (md ? '날짜 서식' : '8자리 숫자') + ').</li></ol></details></div>' +

    '<div class="tabpanel" id="p-pum" hidden><section class="panel"><div class="ph"><h2>매출품의 입력 내용</h2><p>직전 품의(26.8월분 ' + (md ? 'YGK-광고IP사업팀-26-0694' : 'YGK-광고IP사업팀-26-0691') + ') 양식·문구 그대로입니다. 첨부: 정산 원본 파일.</p></div>' +
    '<div class="allcopy"><button class="btn primary" data-copy="all">본문 전체 복사</button></div>' +
    '<div class="block"><header><h3>제목</h3><button class="btn sm" data-copy="title">복사</button></header><div class="body"><div id="pTitle"></div></div></div>' +
    '<div class="block"><header><h3>필수 입력 · 1. 기본 내역 · 2. 세금계산서 발행 정보</h3><button class="btn sm" data-copy="basic">전체 복사</button></header><div class="body"><dl class="kv" id="pBasic"></dl></div></div>' +
    '<div class="block"><header><h3>3. 상세내역</h3><button class="btn sm" data-copy="detail">문구 복사</button></header><div class="body"><pre class="fmt" id="pDetail"></pre>' +
    '<div class="row" style="justify-content:space-between"><b class="small">아티스트별 표</b><button class="btn sm" data-copy="artTable">표 복사</button></div><div class="tablebox"><table id="artTable"></table></div>' +
    '<pre class="fmt" id="pNote"></pre></div></div></section></div>' +

    '<div class="tabpanel" id="p-erp" hidden><section class="panel"><div class="ph"><h2>ERP AR 업로드</h2><p>품의 승인 후 회계 → 매출 추가 → 기안문서 연동 → 세금계산서 등록(부가세만 생성) → 거래내역 엑셀 업로드에 아래 파일을 올리고 처리 → 결재. 업로드 후 날짜 칸이 에러 나는지 한 번 확인합니다.</p></div>' +
    '<div class="row"><button class="btn primary" id="btnErp">AR 업로드 파일 내려받기</button><button class="btn" id="btnSave">이 달 확정 저장</button><span class="small muted" id="erpMsg"></span></div>' +
    '<div class="block"><header><h3>전표 헤더 · 세금계산서</h3><button class="btn sm" data-copy="head">전체 복사</button></header><div class="body"><dl class="kv" id="pHead"></dl></div></div>' +
    '<div class="block"><header><h3>AR 줄 (업로드 파일 내용)</h3><button class="btn sm" data-copy="arTable">표 복사</button></header><div class="body"><div class="tablebox"><table id="arTable"></table></div></div></div></section></div>';

  function render() {
    var flags = R.flags.slice();
    if (saved) flags.push({ level: 'ok', text: '이 달(' + key + ')은 ' + (saved.savedAt || '') + ' 확정 저장된 기록이 있습니다 (' + fmt(saved.supply) + '원)' });
    $('tiles').innerHTML = [['정산월', p.ym + (md ? ' · MD 로열티' : ' · 응원봉')], ['공급가액', fmt(inv.supply) + '원'], ['부가세', fmt(inv.vat) + '원'], ['총금액', fmt(inv.total) + '원'], ['ERP 줄', R.lines.length + '줄']]
      .map(function (t) { return '<div class="tile"><span>' + esc(t[0]) + '</span><b>' + esc(t[1]) + '</b></div>'; }).join('');
    var st = statusHtml(flags); $('statusLine').className = st.cls; $('statusLine').innerHTML = st.html;
    $('amtMeta').textContent = '세금계산서 작성일 ' + (inv.date || '-') + ' · ' + ctx.fileName;
    $('flags').innerHTML = flagsHtml(flags);
    TABLES.lines = { head: ['원본 표기', 'ERP 적요 표기', '공급가액(대변)'].concat(md ? ['프로젝트', '그룹', '아티스트'] : []), numCols: [2], sumLast: true,
      rows: R.lines.map(function (l) { return [l.src, l.label, fmt(l.amt)].concat(md ? [l.proj, l.grp, l.art] : []); }).concat([['합계', '', fmt(R.lines.reduce(function (a, l) { return a + l.amt; }, 0))].concat(md ? ['', '', ''] : [])]) };
    $('lineTable').innerHTML = tableHtml(TABLES.lines);

    // ---- 품의 ----
    TXT.title = title; $('pTitle').textContent = title;
    var basic = [['비거주자 여부', '거주자'], ['외국환 신고 필요여부', '1.경상거래 해당없음 / 2.자본거래 해당없음'], ['거래처명', (P.vendorName || '') + ' (' + (P.vendor || '') + ')'], ['거래일자', inv.date || '-'],
      ['금액', fmt(inv.supply)], ['VAT', fmt(inv.vat)], ['총금액', fmt(inv.total)], ['프로젝트', '프로젝트공통 (_)'], ['입금예정일 (이내)', pay || '-'],
      ['아티스트(그룹)', '그룹공통 (_)'], ['아티스트(개인)', (P.artistName || '아티스트공통') + ' (' + (P.artist || '') + ')'], ['매출 종류', '로열티'], ['계약서 작성 여부', '작성'], ['통화종류', 'KRW'],
      ['청구방법', '정산서 수령 후 일시'], ['발행정보', '세금계산서'], ['사업자등록번호', P.bizNo || ''], ['담당자 / 이메일', (P.contact || '') + ' / ' + (P.email || '')], ['관련문서', '[매출품의] YGP MD 로열티 ' + (+YGP.addMonths(p.ym, -1).slice(5)) + '월 정산의 건' + (md ? '' : '_응원봉 디자인권 사용료')],
      ['결재 라인 (위임전결)', approval(inv.supply)]];
    TXT.basic = basic.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
    $('pBasic').innerHTML = kv(basic, 'basic');
    var d1 = '1. ' + mNum + '월 ' + (md ? 'MD 로열티 내역' : '가수 MD_응원봉 디자인상표권 사용료') + '\n\n- 총공급가액 : ' + fmt(inv.supply) + '원 - VAT : ' + fmt(inv.vat) + '원 - 총금액 : ' + fmt(inv.total) + '원\n\n2. 아티스트별 ' + (md ? 'MD 로열티 내역' : '응원봉 내역');
    TABLES.artTable = { head: ['아티스트', '수량', '공급가액', '부가세', '공급대가'], numCols: [1, 2, 3, 4], sumLast: true,
      rows: p.artists.map(function (x) { return [x.name, fmt(x.qty), fmt(x.supply), fmt(x.vat), fmt(x.total)]; }).concat([['합계', fmt(R.S.qty), fmt(R.S.supply), fmt(R.S.vat), fmt(R.S.total)]]) };
    var note = '3. 참고사항 ' + ((md ? PUM.mdNote : PUM.cbNote) || '');
    TXT.detail = d1 + '\n\n' + [TABLES.artTable.head].concat(TABLES.artTable.rows).map(function (r) { return r.join('\t'); }).join('\n') + '\n\n' + note;
    $('pDetail').textContent = d1; $('artTable').innerHTML = tableHtml(TABLES.artTable); $('pNote').textContent = note;
    TXT.all = title + '\n\n' + TXT.basic + '\n\n3. 상세내역\n\n' + TXT.detail;

    // ---- ERP ----
    var head = [['회계일자', inv.date || '-'], ['증빙구분', '전자세금계산서'], ['청구처', (P.vendorName || '') + ' (' + (P.vendor || '') + ')'], ['통화 / 환율', 'KRW / 1'],
      ['제목', title.replace('[매출품의] ', '')], ['세금계산서 (부가세만 생성)', '공급가액 ' + fmt(inv.supply) + ' · 세액 ' + fmt(inv.vat) + ' · 합계 ' + fmt(inv.total) + ' (계산서일자 ' + (inv.date || '-') + ')'],
      ['입금예정일', pay || '-'], ['결재조건', '입금계좌'], ['거래내역 엑셀 업로드', R.lines.length + '줄 · 대변 합계 ' + fmt(R.lines.reduce(function (a, l) { return a + l.amt; }, 0)) + '원']];
    TXT.head = head.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
    $('pHead').innerHTML = kv(head, 'head');
    TABLES.arTable = { head: ['프로젝트', '그룹', '아티스트', '대변금액', '적요'], numCols: [3], sumLast: true,
      rows: R.lines.map(function (l) { return [md ? l.proj : '_', md ? l.grp : '_', md ? l.art : (P.artist || ''), fmt(l.amt), YGP.memo(p, l)]; })
        .concat([['합계', '', '', fmt(R.lines.reduce(function (a, l) { return a + l.amt; }, 0)), '']]) };
    $('arTable').innerHTML = tableHtml(TABLES.arTable);
    var blocked = R.flags.some(function (f) { return f.level === 'bad'; });
    $('btnErp').disabled = false;
    $('erpMsg').textContent = blocked ? '검증 "조치" 항목을 먼저 확인하세요 (파일은 내려받을 수 있습니다)' : '';
  }

  // 복사·탭
  function showFallback(text) { $('copyFallback').hidden = false; var ta = $('fbText'); ta.value = text; ta.focus(); ta.select(); ctx.toast('자동 복사가 막혀 있어 아래 칸에 표시했습니다'); }
  function copyText(text, label) {
    if (!navigator.clipboard || !navigator.clipboard.writeText) { showFallback(text); return; }
    navigator.clipboard.writeText(text).then(function () { ctx.toast((label || '') + ' 복사했습니다'); }, function () { showFallback(text); });
  }
  function copyTable(t) {
    var tsv = [t.head].concat(t.rows).map(function (r) { return r.join('\t'); }).join('\n');
    var html = '<table border="1" style="border-collapse:collapse">' + tableHtml(t).replace(/ class="[^"]*"/g, '') + '</table>';
    try {
      if (window.ClipboardItem && navigator.clipboard.write) {
        navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([tsv], { type: 'text/plain' }) })])
          .then(function () { ctx.toast('표를 복사했습니다 (표 서식 포함)'); }, function () { copyText(tsv, '표를'); });
        return;
      }
    } catch (e) { }
    copyText(tsv, '표를');
  }
  ctx.root.addEventListener('click', function (e) {
    var b = e.target.closest('[data-copy]');
    if (b) { var k = b.getAttribute('data-copy'); if (TABLES[k]) copyTable(TABLES[k]); else if (TXT[k] != null) copyText(TXT[k], '내용을'); return; }
    var v = e.target.closest('[data-copyval]');
    if (v) { var q = v.getAttribute('data-copyval').split(':'), line = (TXT[q[0]] || '').split('\n')[+q[1]] || ''; copyText(line.split('\t').slice(1).join('\t'), '값을'); }
  });
  var TABS = ['calc', 'pum', 'erp'];
  function showTab(t) {
    if (TABS.indexOf(t) < 0) t = TABS[0];
    TABS.forEach(function (k) { $('p-' + k).hidden = k !== t; $('t-' + k).setAttribute('aria-selected', k === t ? 'true' : 'false'); });
    try { localStorage.setItem('ygp.tab', t); } catch (e) { }
  }
  ctx.tabsEl.querySelectorAll('.tab').forEach(function (b) { b.addEventListener('click', function () { showTab(b.getAttribute('data-tab')); window.scrollTo(0, 0); }); });
  var st = null; try { st = localStorage.getItem('ygp.tab'); } catch (e) { }

  $('btnErp').addEventListener('click', function () {
    var out = YGP.buildErpUpload(p, R, CFG);
    ctx.saveFile((md ? '[md]' : '[응원]') + 'AR거래내역엑셀업로드_YGPMD' + (md ? '' : '_응원봉') + '_' + p.ym.replace('-', '.') + '.xlsx', XLSX.write(out.wb, { type: 'array', bookType: 'xlsx' }));
    ctx.toast('AR 업로드 파일을 내려받았습니다 (' + out.rows.length + '줄)');
  });
  $('btnSave').addEventListener('click', async function () {
    try {
      $('erpMsg').textContent = '저장 중…';
      await ctx.post('saveQuarter', { record: { quarter: key, supply: inv.supply, vat: inv.vat, gross: inv.total, settleMonth: p.ym, txDate: inv.date, payDate: pay, file: ctx.fileName,
        lines: R.lines.map(function (l) { return { label: l.label, amt: l.amt }; }), royaltyTable: [] } });
      $('erpMsg').textContent = ''; ctx.toast(key + ' 확정 저장했습니다');
    } catch (e) { $('erpMsg').textContent = '저장하지 못했습니다 (' + e.message + ')'; }
  });

  render(); showTab(st || 'calc');
  return { quarter: p.ym + (md ? ' MD' : ' 응원봉'), note: '항목별정산내역_요약 · 아티스트 ' + p.artists.length + '명 · 작성일 ' + (inv.date || '-') };
}

var kindOf = function (wb) { try { return YGP.parse(wb).kind; } catch (e) { return ''; } };
var why = function (wb) { if (!wb.Sheets['항목별정산내역_요약']) return ''; try { YGP.parse(wb); return ''; } catch (e) { return e.message; } };
DESK.register({ id: 'ygp-md', server: 'ygp', name: 'YGP · MD 로열티', desc: '와이지플러스 굿즈 판매 MD 로열티 · 월 매출',
  fileHint: 'YYYY.MM_Ent 정산_EUC_V*.xlsb (암호 제거본)', outputs: ['정산서 숫자 검증', '매출품의', 'ERP AR 업로드 파일'],
  detect: function (wb) { return kindOf(wb) === 'md'; }, why: why, mount: mount });
DESK.register({ id: 'ygp-cb', server: 'ygp', name: 'YGP · 응원봉 디자인권', desc: '와이지플러스 응원봉 디자인권·상표권 사용료 · 월 매출',
  fileHint: 'YYYY.MM_응원봉 디자인권 및 상표권 사용료 정산_V*.xlsx', outputs: ['정산서 숫자 검증', '매출품의', 'ERP AR 업로드 파일'],
  detect: function (wb) { return kindOf(wb) === 'cb'; }, why: why, mount: mount });
})();
