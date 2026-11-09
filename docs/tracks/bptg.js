/* 정산 모듈: 블랙핑크더게임(BPTG) 게임 매출정산 (테이크원 → YG, 분기 매출)
 * 판별(detect) → 셸이 비공개 설정(서버 트랙 thegirls: 같은 BPTG 사업·테이크원 거래처)을 받아 mount → 정산·품의·ERP 탭.
 * 회사 데이터(거래처·코드·누적 정산 이력)는 이 파일에 없다(DB "THEGIRLS_설정" BPTG · BPTG_HIST). */
(function () {
/* BPTG 엔진 — 26.2Q 매출품의(YGK-광고IP사업팀-26-0620)·ERP(AR26081800006) 금액과 1원 단위 일치 검증 */
var BPTG = (function () {
  var N = function (s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); };
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  function num(v) { var n = typeof v === 'number' ? v : parseFloat(String(v == null ? '' : v).replace(/[,\s]/g, '')); return isFinite(n) ? n : null; }
  function fmt(n) { return Math.round(n).toLocaleString('ko-KR'); }
  function rowsOf(ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }); }
  function monthEnd(ym) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7); return ym + '-' + pad(new Date(Date.UTC(y, m, 0)).getUTCDate()); }
  function addMonths(ym, n) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + n; y += Math.floor(m / 12); m = ((m % 12) + 12) % 12; return y + '-' + pad(m + 1); }

  // 테이크원 "BPTG_게임매출_YYYY.MM(또는 nQ)_정산서_YG.xlsx": 월별 "공문_YG_YY.MM" 시트 + "지급정산서" 시트
  // 분기 공문 = 분기 마지막 달 공문 시트(그 분기 3개월이 한 표에 있음). 테이크원이 이전 달 숫자를 수정해 다시 보내므로 항상 최신 파일 기준
  function parse(wb) {
    if (!wb.Sheets['지급정산서']) throw new Error('"지급정산서" 시트가 없습니다. 테이크원 BPTG 게임매출 정산서(엑셀)인지 확인해 주세요.');
    var quarters = [];
    wb.SheetNames.forEach(function (s) {
      var m = s.match(/^공문_YG_(\d{2})\.(\d{2}|\dQ)$/); if (!m) return;
      var y = 2000 + +m[1], q = /Q$/.test(m[2]) ? +m[2][0] : (+m[2] % 3 === 0 ? +m[2] / 3 : 0);
      if (!q) return;
      var months = officialRows(wb.Sheets[s]);
      var want = [1, 2, 3].map(function (k) { return y + '-' + pad((q - 1) * 3 + k); });
      if (!want.every(function (ym) { return months.some(function (r) { return r.ym === ym; }); })) return;
      quarters.push({ key: y + '-Q' + q, y: y, q: q, sheet: s, months: want.map(function (ym) { return months.filter(function (r) { return r.ym === ym; })[0]; }) });
    });
    if (!quarters.length) throw new Error('분기 공문 시트(공문_YG_YY.03/06/09/12)를 찾지 못했습니다.');
    quarters.sort(function (a, b) { return a.key < b.key ? -1 : 1; });
    return { quarters: quarters, comp: components(wb.Sheets['지급정산서']) };
  }
  // 공문 시트 "3) 정산 세부내역": 매출월 · 순매출액 · 수익배분율 · 정산금액 · 부가가치세 · 합계 (첫 "합계" 행까지)
  function officialRows(ws) {
    var rows = rowsOf(ws), out = [], on = false;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i].map(N), j = r.indexOf('매출월');
      if (j >= 0) { on = true; continue; }
      if (!on) continue;
      if (r.indexOf('합계') >= 0 || r.some(function (c) { return /^합계/.test(c); })) break;
      var mi = -1; r.forEach(function (c, k) { if (mi < 0 && /^\d{4}년 \d{1,2}월$/.test(c)) mi = k; });
      if (mi < 0) continue;
      var mm = r[mi].match(/(\d{4})년 (\d{1,2})월/), v = rows[i].slice(mi + 1).map(num).filter(function (x) { return x != null; });
      out.push({ ym: mm[1] + '-' + pad(+mm[2]), net: v[0] || 0, rate: v[1] || 0, share: v[2] || 0, vat: v[3] || 0, total: v[4] || 0 });
    }
    return out;
  }
  // 지급정산서 하단 월별 구성표: 게임 순매출액 / MD_1 / IAA / MD_2 / 음원 / YouTube / 수익분배 기준금액 (머리글 행 "구분 | 2026년 01월 …")
  function components(ws) {
    var rows = rowsOf(ws), out = {}, head = -1, cols = [];
    for (var i = rows.length - 1; i >= 0; i--) {
      var r = rows[i].map(N), k = r.indexOf('구분');
      if (k >= 0 && r.some(function (c) { return /^\d{4}년 \d{1,2}월$/.test(c); })) { head = i; cols = r; break; }
    }
    if (head < 0) return out;
    var lab = cols.indexOf('구분');
    for (var j = head + 1; j < Math.min(rows.length, head + 12); j++) {
      var name = N(rows[j][lab]); if (!name) continue;
      cols.forEach(function (c, ci) {
        var m = c.match(/^(\d{4})년 (\d{1,2})월$/); if (!m) return;
        var ym = m[1] + '-' + pad(+m[2]); out[ym] = out[ym] || {}; out[ym][name] = num(rows[j][ci]) || 0;
      });
    }
    return out;
  }

  // 분기 계산: 공급가액 = 월별 정산금액 합(원 단위 반올림), VAT = 월별 부가세 합(반올림) — 26.2Q 2,830,788 / 283,078 일치
  function compute(qt, comp, cfg) {
    var B = cfg.BPTG || {}, rate = B.rate || 0.23, flags = [];
    var sum = function (k) { return qt.months.reduce(function (a, r) { return a + r[k]; }, 0); };
    var supply = Math.round(sum('share')), vat = Math.round(sum('vat')), net = Math.round(sum('net'));
    var vat10 = Math.floor(supply * 0.1);
    qt.months.forEach(function (r) {
      var exp = Math.round(r.net * rate);
      if (Math.abs(exp - r.share) >= 1) flags.push({ level: 'warn', text: r.ym + ' 정산금액 ' + fmt(r.share) + '원이 순매출 × ' + Math.round(rate * 100) + '% (' + fmt(exp) + '원)와 다릅니다 — 공문 금액 그대로 사용' });
      var c = comp[r.ym];
      if (c && c['수익분배 기준금액'] != null && Math.abs(c['수익분배 기준금액'] - r.net) >= 1) flags.push({ level: 'bad', text: r.ym + ' 지급정산서 수익분배 기준금액 ' + fmt(c['수익분배 기준금액']) + '원 ≠ 공문 순매출액 ' + fmt(r.net) + '원' });
    });
    if (!flags.some(function (f) { return f.level === 'bad'; })) flags.push({ level: 'ok', text: '공문(' + qt.sheet + ') 월별 순매출액 = 지급정산서 수익분배 기준금액, 정산금액 = 순매출 × ' + Math.round(rate * 100) + '% 확인' });
    flags.push({ level: 'ok', text: '분기 공급가액 ' + fmt(supply) + '원 (월별 정산금액 합) · 부가세 ' + fmt(vat) + '원 (공문 월별 부가세 합) · 합계 ' + fmt(supply + vat) + '원' });
    if (vat !== vat10) flags.push({ level: 'warn', text: '공문 부가세 ' + fmt(vat) + '원 ≠ 공급가액 × 10% ' + fmt(vat10) + '원 (' + (vat - vat10) + '원) — 월별로 부가세를 끊어 계산해 생기는 차이. 세금계산서·ERP 부가세는 공문 기준 ' + fmt(vat) + '원으로 두었습니다. 회계팀 기준과 다르면 수정하세요' });
    var y2 = String(qt.y).slice(2), qEnd = qt.y + '-' + pad(qt.q * 3);
    var tx = addMonths(qEnd, 2) + '-01';                          // 세금계산서 발행일: 분기 마지막 달 + 2개월 1일 (2Q → 8/1, 3Q → 11/1)
    var pay = monthEnd(addMonths(tx.slice(0, 7), 1));             // 입금예정일: 거래일자 다음 달 말일 (26.2Q 품의 9/30)
    return { supply: supply, vat: vat, total: supply + vat, net: net, vat10: vat10, tx: tx, pay: pay, y2: y2, flags: flags,
      start: qt.months[0].ym + '-01', end: monthEnd(qt.months[2].ym) };
  }
  return { parse: parse, compute: compute, fmt: fmt, addMonths: addMonths, monthEnd: monthEnd, pad: pad };
})();
if (typeof module !== 'undefined') module.exports = BPTG;
if (typeof DESK === 'undefined') return;

// ---------------- 화면 ----------------
var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
var fmt = BPTG.fmt;
var tsv = function (t) { return [t.head].concat(t.rows).map(function (r) { return r.join('\t'); }).join('\n'); };
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
function kv(list, grp) { return list.map(function (x, i) { return '<dt>' + esc(x[0]) + '</dt><dd>' + esc(x[1]) + '</dd><button class="btn sm" data-copyval="' + grp + ':' + i + '">복사</button>'; }).join(''); }
function flagsHtml(fl) { return fl.map(function (f) { return '<div class="flag ' + f.level + '"><b>' + { ok: '확인', warn: '주의', bad: '조치' }[f.level] + '</b><span>' + esc(f.text) + '</span></div>'; }).join(''); }
function statusHtml(fl) {
  var nb = fl.filter(function (f) { return f.level === 'bad'; }).length, nw = fl.filter(function (f) { return f.level === 'warn'; }).length, lvl = nb ? 'bad' : nw ? 'warn' : 'ok';
  return { cls: 'status ' + lvl, html: '<b>' + (nb ? '조치 필요 ' + nb + '건' : nw ? '체크 ' + nw + '건' : '검증 이상 없음') + '</b><span>' + esc((fl.filter(function (f) { return f.level === lvl; })[0] || {}).text || '') + '</span>' };
}
function approval(supply) {
  if (supply < 5e6) return '500만원 미만 (합의종결): 기안 → 팀 리더(최종) · 합의: 재무관리팀 담당, 회계팀 담당';
  if (supply < 5e7) return '500만원 이상 5,000만원 미만 (합의종결): 기안 → 팀 리더 → 그룹 리더(최종) · 합의: 재무관리팀 담당, 회계팀 담당 → 회계팀 리더';
  if (supply < 3e8) return '5,000만원 이상 3억원 미만: 기안 → 팀 리더 → 그룹 리더 · 합의: 재무관리팀·회계팀 담당 → 회계팀 리더 → CFO → 부문 대표(최종)';
  return '3억원 이상: 기안 → 팀 리더 → 그룹 리더 · 합의: 재무관리팀·회계팀 담당 → 회계팀 리더 → CFO → 부문 대표 → 대표이사(최종)';
}
var ymKo = function (ym) { return ym.slice(0, 4) + '년 ' + (+ym.slice(5)) + '월'; };

function mount(ctx) {
  var $ = function (id) { return document.getElementById(id); };
  var CFG = (ctx.data && ctx.data.cfg) || {}, B = CFG.BPTG || {}, H = CFG.BPTG_HIST || { rows: [] };
  var P = BPTG.parse(ctx.wb), S = { key: P.quarters[P.quarters.length - 1].key }, TXT = {}, TABLES = {};

  ctx.tabsEl.innerHTML = tabsHtml([['calc', '정산'], ['pum', '품의 복붙'], ['erp', 'ERP 전표']]);
  ctx.root.innerHTML =
    '<div class="tabpanel" id="p-calc"><section class="panel"><div class="form"><label>정산 분기<select id="inQ">' + P.quarters.map(function (q) { return '<option value="' + q.key + '">' + q.y + '년 ' + q.q + '분기 (' + esc(q.sheet) + ')</option>'; }).join('') + '</select></label></div>' +
    '<p class="small muted" style="margin:0">테이크원이 지난달 숫자를 고쳐서 다시 보내므로 <b>가장 최근에 받은 정산서 엑셀</b>을 올려 주세요(26.2Q는 품의 도중 광고매출이 수정돼 2,830,720 → 2,830,788원으로 바뀜).</p></section>' +
    '<section class="panel dash"><div class="tiles" id="tiles"></div><div class="status" id="statusLine"></div><div class="meta" id="amtMeta"></div><details class="fold vfold" open><summary>정산서 숫자 검증</summary><div class="flags" id="flags"></div></details></section>' +
    '<section class="panel"><div class="ph"><h2>월별 수익배분</h2><p>지급정산서 하단 구성표와 분기 공문 시트 기준입니다.</p></div><div class="tablebox"><table id="monTable"></table></div></section>' +
    '<details class="panel fold logic"><summary><b>적용 중인 정산 로직</b> <span class="small muted">26.2Q 매출품의 26-0620 · ERP AR26081800006 · 인수인계 녹음 기준</span></summary><ol>' +
    '<li><b>수익분배 기준금액</b>: 게임 순매출 + MD_1 매출총이익(MD + OST MERCH) + IAA 매출총이익 + MD_2 + 음원 + YouTube.</li>' +
    '<li><b>YG 정산금액</b>: 월별 기준금액 × 23%. 분기 공급가액 = 3개월 합(원 단위), 부가세 = 공문 월별 부가세 합.</li>' +
    '<li><b>MG</b>: 20억 원은 2024년 10월에 모두 차감 완료 — 이후 정산금액 = 실입금.</li>' +
    '<li><b>날짜</b>: 거래일자(세금계산서 발행일) = 분기 마지막 달 + 2개월 1일, 입금예정일 = 그 다음 달 말일.</li></ol></details></div>' +

    '<div class="tabpanel" id="p-pum" hidden><section class="panel"><div class="ph"><h2>매출품의 입력 내용</h2><p>직전 품의(26.2Q YGK-광고IP사업팀-26-0620) 양식·문구 그대로입니다.</p></div>' +
    '<div class="allcopy"><button class="btn primary" data-copy="all">본문 전체 복사</button></div>' +
    '<div class="block"><header><h3>제목</h3><button class="btn sm" data-copy="title">복사</button></header><div class="body"><div id="pTitle"></div></div></div>' +
    '<div class="block"><header><h3>필수 입력 · 1. 기본 내역 · 2. 세금계산서 발행 정보</h3><button class="btn sm" data-copy="basic">전체 복사</button></header><div class="body"><dl class="kv" id="pBasic"></dl></div></div>' +
    '<div class="block fixed"><header><h3>3. 상세내역 — 1. 상세 내역 <span class="stamp">지정양식</span></h3><button class="btn sm" data-copy="fixed">텍스트 복사</button></header><div class="body"><pre class="fmt" id="pFixed"></pre>' +
    '<div class="row" style="justify-content:space-between"><b class="small">판매월 표</b><button class="btn sm" data-copy="fixTable">표 복사</button></div><div class="tablebox"><table id="fixTable"></table></div></div></div>' +
    '<div class="block"><header><h3>3. 상세내역 — 2. 상세 정산 내역 · 3. 첨부 파일</h3><button class="btn sm" data-copy="detail">문구 복사</button></header><div class="body"><pre class="fmt" id="pDetail"></pre>' +
    '<div class="row" style="justify-content:space-between"><b class="small">누적 월별 정산 세부내역 &amp; MG 차감 후 실입금 내역</b><button class="btn sm" data-copy="histTable">표 복사</button></div><div class="tablebox"><table id="histTable"></table></div>' +
    '<pre class="fmt" id="pNote"></pre></div></div></section></div>' +

    '<div class="tabpanel" id="p-erp" hidden><section class="panel"><div class="ph"><h2>ERP AR 매출 전표</h2><p>품의 승인 후 회계 → 매출 추가에서 직접 입력합니다(공급가 1줄). 기안문서 연동 → 세금계산서 등록(부가세만 생성) → 거래내역 입력·<b>내역 저장</b> → 결재 생성. "결제 취소"는 누르지 마세요.</p></div>' +
    '<div class="row"><button class="btn" id="btnErp">AR 업로드 파일(1줄)도 받기</button><button class="btn" id="btnSave">이 분기 확정 저장</button><span class="small muted" id="erpMsg"></span></div>' +
    '<div class="block"><header><h3>전표 헤더 · 세금계산서</h3><button class="btn sm" data-copy="head">전체 복사</button></header><div class="body"><dl class="kv" id="pHead"></dl></div></div>' +
    '<div class="block"><header><h3>거래내역 (공급가 1줄)</h3><button class="btn sm" data-copy="line">전체 복사</button></header><div class="body"><dl class="kv" id="pLine"></dl></div></div></section></div>';

  var Q, R;
  function render() {
    Q = P.quarters.filter(function (q) { return q.key === S.key; })[0]; R = BPTG.compute(Q, P.comp, CFG);
    var qLab = R.y2 + '년 ' + Q.q + '분기', qShort = Q.y + '.' + Q.q + 'Q';
    var saved = (ctx.data.saved || {})['GAME ' + qShort], flags = R.flags.slice();
    if (saved) flags.push({ level: 'ok', text: qShort + '은 ' + (saved.savedAt || '') + ' 확정 저장된 기록이 있습니다 (' + fmt(saved.supply) + '원)' });
    $('tiles').innerHTML = [['정산 분기', qShort], ['순매출 합계', fmt(R.net) + '원'], ['공급가액 (YG 23%)', fmt(R.supply) + '원'], ['부가세', fmt(R.vat) + '원'], ['총금액', fmt(R.total) + '원']]
      .map(function (t) { return '<div class="tile"><span>' + esc(t[0]) + '</span><b>' + esc(t[1]) + '</b></div>'; }).join('');
    var st = statusHtml(flags); $('statusLine').className = st.cls; $('statusLine').innerHTML = st.html;
    $('amtMeta').textContent = '거래일자(세금계산서) ' + R.tx + ' · 입금예정일 ' + R.pay + ' · ' + ctx.fileName;
    $('flags').innerHTML = flagsHtml(flags);
    var keys = ['게임 순매출액', 'MD_1 매출총이익', 'IAA 매출총이익', 'MD_2 매출총이익', '음원 매출', 'YouTube 매출'];
    TABLES.mon = { head: ['매출월'].concat(keys, ['수익분배 기준금액', 'YG 정산금액', '부가세']), numCols: [1, 2, 3, 4, 5, 6, 7, 8, 9], sumLast: true,
      rows: Q.months.map(function (r) { var c = P.comp[r.ym] || {}; return [r.ym].concat(keys.map(function (k) { return fmt(c[k] || 0); }), [fmt(r.net), fmt(r.share), fmt(r.vat)]); })
        .concat([['합계'].concat(keys.map(function (k) { return fmt(Q.months.reduce(function (a, r) { return a + ((P.comp[r.ym] || {})[k] || 0); }, 0)); }), [fmt(R.net), fmt(R.supply), fmt(R.vat)])]) };
    $('monTable').innerHTML = tableHtml(TABLES.mon);

    // ---- 품의 ----
    var title = '[매출품의] 블랙핑크더게임(BPTG) ' + qLab + ' 매출 정산 건';
    TXT.title = title; $('pTitle').textContent = title;
    var prevQ = Q.q === 1 ? (Q.y - 1) + '-Q4' : Q.y + '-Q' + (Q.q - 1), pq = prevQ.split('-Q');
    var basic = [['비거주자 여부', '거주자'], ['외국환 신고 필요여부', '1.경상거래 해당없음 / 2.자본거래 해당없음'], ['거래처명', (B.vendor || '') + ' (' + (B.vendorCode || '') + ')'], ['거래일자', R.tx],
      ['금액', fmt(R.supply)], ['VAT', fmt(R.vat)], ['총금액', fmt(R.total)], ['프로젝트', (B.projName || '') + ' (' + (B.proj || '') + ')'], ['입금예정일 (이내)', R.pay],
      ['아티스트(그룹)', (B.groupName || '') + ' (' + (B.group || '') + ')'], ['아티스트(개인)', (B.artistName || '') + ' (' + (B.artist || '') + ')'], ['매출 종류', '로열티'],
      ['계약서 작성 여부', '작성'], ['전체계약기간', B.contract || ''], ['통화종류', 'KRW'], ['청구방법', '정산서 수령 후'], ['에이전시 유/무', '없음'],
      ['발행정보', '세금계산서'], ['사업자등록번호', B.bizNo || ''], ['담당자 / 이메일', (B.contact || '') + ' / ' + (B.email || '')],
      ['관련문서', '[매출품의] 블랙핑크더게임(BPTG) ' + String(pq[0]).slice(2) + '년 ' + pq[1] + '분기 매출 정산 건' + (B.relatedDoc ? ' / ' + B.relatedDoc : '')],
      ['첨부파일', '월별 정산서 엑셀·공문 PDF·MD 관련 수불부·YouTube/AdMob/Unity 정산서 (' + Q.months.map(function (r) { return +r.ym.slice(5) + '월'; }).join('·') + ')'],
      ['결재 라인 (위임전결)', approval(R.supply)]];
    TXT.basic = basic.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
    $('pBasic').innerHTML = kv(basic, 'basic');
    var fixed = '1. 상세 내역\n- 정산월 : ' + R.tx.slice(0, 7) + '\n- 프로젝트 : ' + (B.bisProject || 'BPTG') + '\n- 사업구조 : ' + (B.struct || 'MG_RS') + '\n- MG : -\n- 통화 : KRW\n- 환율 : 1';
    TXT.fixed = fixed; $('pFixed').textContent = fixed;
    TABLES.fixTable = { head: ['판매월', '아티스트', '권역', '판매수량', '총 매출 (vat-)', '로열티 (vat-)'], numCols: [4, 5], rows: Q.months.map(function (r) { return [r.ym, B.artistLabel || '블랙핑크', B.region || '글로벌', '-', fmt(r.net), fmt(r.share)]; }) };
    $('fixTable').innerHTML = tableHtml(TABLES.fixTable);
    // 누적 표: DB 이력(26.2Q 품의 표) + 이번 분기. 합계는 직전 품의 합계에 이번 분기를 더한다(직전 품의 합계와 이어지게)
    var rows = (H.rows || []).filter(function (r) { return !Q.months.some(function (m) { return m.ym === r[0]; }); }).concat(Q.months.map(function (m) { return [m.ym, Math.round(m.net), Math.round(m.share), Math.round(m.share)]; }));
    rows.sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    var base = H.totals && H.totals.asOf < Q.months[0].ym ? H.totals : null;
    var added = rows.filter(function (r) { return !base || r[0] > base.asOf; });
    var tot = base ? { net: base.net, share: base.share, paid: base.paid } : { net: 0, share: 0, paid: 0 };
    added.forEach(function (r) { tot.net += r[1]; tot.share += r[2]; tot.paid += r[3] || 0; });
    TABLES.histTable = { head: ['매출월', '순매출액(원)', '수익배분율', '정산금액(원)', '실입금액(원)'], numCols: [1, 3, 4], sumLast: true,
      rows: rows.map(function (r) { return [ymKo(r[0]), fmt(r[1]), '23%', fmt(r[2]), r[4] || (r[3] == null ? '-' : fmt(r[3]))]; }).concat([['합계', fmt(tot.net), '23%', fmt(tot.share), fmt(tot.paid)]]) };
    $('histTable').innerHTML = tableHtml(TABLES.histTable);
    var d2 = '2. 상세 정산 내역\n- ' + qLab + ' 총매출 : ' + fmt(R.net) + '원\n- ' + qLab + ' YG 정산금액 : ' + fmt(R.supply) + '원 (VAT 제외)\n- 누적 월별 정산 세부내역 & MG 차감 후 실입금 내역';
    var hasMusic = Q.months.some(function (m) { return ((P.comp[m.ym] || {})['음원 매출'] || 0) !== 0; });
    var notes = (B.notes || ['※ 2024년 10월, 잔여 MG 차감 및 4분기 매출부터 실입금 발생', '※ 정산 대상: 앱마켓 매출 + MD 매출 + IAA 매출 + 유튜브 매출']).map(function (n) { return hasMusic && /정산 대상/.test(n) && !/음원/.test(n) ? n + ' + 음원 매출' : n; });
    if (B.endNote) notes.push(B.endNote);
    var note = notes.join('\n') + '\n3. 첨부 파일\n - ' + R.y2 + '.' + Q.q + '분기 항목별 월 정산 내역서 각 1부.\n - ' + R.y2 + '.' + Q.q + '분기 정산 공문 1부. 끝.';
    $('pDetail').textContent = d2; $('pNote').textContent = note;
    TXT.detail = d2 + '\n' + tsv(TABLES.histTable) + '\n' + note;
    TXT.all = title + '\n\n' + TXT.basic + '\n\n3. 상세내역\n' + fixed + '\n' + tsv(TABLES.fixTable) + '\n' + TXT.detail;

    // ---- ERP ----
    var erpTitle = '블랙핑크더게임(BPTG) ' + qLab + ' 매출 정산', memo = qShort + ' 블랙핑크더게임(BPTG) 매출 로열티 정산';
    var head = [['회계일자', R.tx], ['증빙구분', '전자세금계산서 (정발행)'], ['AR 유형', '국내 매출'], ['청구처', (B.vendor || '') + ' (' + (B.vendorCode || '') + ')'], ['통화 / 환율', 'KRW / 1'],
      ['결제조건 / 입금예정일', '입금계좌 / ' + R.pay], ['프로젝트', (B.projName || '') + ' (' + (B.proj || '') + ')'], ['제목', erpTitle], ['기안문서', '위 매출품의 연동'],
      ['세금계산서 (부가세만 생성)', '공급가액 ' + fmt(R.supply) + ' · 세액 ' + fmt(R.vat) + ' · 계산서일자 ' + R.tx + ' · 신고사업장 (주)YG엔터테인먼트'],
      ['결재선', '재무관리팀 담당 합의 → 팀 리더 결재 (공급가 2억 원 미만)']];
    TXT.head = head.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n'); $('pHead').innerHTML = kv(head, 'head');
    var line = [['계정코드', (B.accountName || '') + ' (' + (B.account || '') + ')'], ['회계부서', (B.deptName || '') + ' (' + (B.dept || '') + ')'], ['프로젝트', (B.projName || '') + ' (' + (B.proj || '') + ')'],
      ['대변 금액 (공급가)', fmt(R.supply)], ['적요', memo], ['아티스트그룹 / 아티스트', (B.groupName || '') + ' (' + (B.group || '') + ') / ' + (B.artistName || '') + ' (' + (B.artist || '') + ')'],
      ['정산시작일 / 종료일', R.start + ' / ' + R.end], ['국가구분', 'KR']];
    TXT.line = line.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n'); $('pLine').innerHTML = kv(line, 'line');
    R.memo = memo; R.qShort = qShort;
  }

  function showFallback(text) { $('copyFallback').hidden = false; var ta = $('fbText'); ta.value = text; ta.focus(); ta.select(); ctx.toast('자동 복사가 막혀 있어 아래 칸에 표시했습니다'); }
  function copyText(text, label) {
    if (!navigator.clipboard || !navigator.clipboard.writeText) { showFallback(text); return; }
    navigator.clipboard.writeText(text).then(function () { ctx.toast((label || '') + ' 복사했습니다'); }, function () { showFallback(text); });
  }
  function copyTable(t) {
    var html = '<table border="1" style="border-collapse:collapse">' + tableHtml(t).replace(/ class="[^"]*"/g, '') + '</table>';
    try {
      if (window.ClipboardItem && navigator.clipboard.write) {
        navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([tsv(t)], { type: 'text/plain' }) })])
          .then(function () { ctx.toast('표를 복사했습니다 (표 서식 포함)'); }, function () { copyText(tsv(t), '표를'); });
        return;
      }
    } catch (e) { }
    copyText(tsv(t), '표를');
  }
  ctx.root.addEventListener('click', function (e) {
    var b = e.target.closest('[data-copy]');
    if (b) { var k = b.getAttribute('data-copy'); if (TABLES[k]) copyTable(TABLES[k]); else if (TXT[k] != null) copyText(TXT[k], '내용을'); return; }
    var v = e.target.closest('[data-copyval]');
    if (v) { var q = v.getAttribute('data-copyval').split(':'), line = (TXT[q[0]] || '').split('\n')[+q[1]] || ''; copyText(line.split('\t').slice(1).join('\t'), '값을'); }
  });
  $('inQ').value = S.key;
  $('inQ').addEventListener('change', function () { S.key = this.value; render(); });
  $('btnErp').addEventListener('click', function () {
    var serial = function (d) { var t = d.split('-'); return Math.round((Date.UTC(+t[0], +t[1] - 1, +t[2]) - Date.UTC(1899, 11, 30)) / 86400000); };
    var head = ['프로젝트', '회계부서', '계정코드', '예산코드', '배부코드', '통화', '환율', '차변금액', '대변금액', '거래처', '반제전표', '적요', '세금일자', '부가세코드', '공급가액', '부가세금액', '신고사업장', '세금계산서번호', '아티스트그룹', '아티스트', '관리항목1', '관리항목2', '관리항목3', '관리항목4', '관리항목5', '관리항목6', '관리항목7', '관리항목8'];
    var aoa = [['AR 거래내역 일괄생성 업로드 양식 // ERP자동전표_IP'], [], head, [+B.proj, +B.dept, String(B.account), null, null, 'KRW', 1, null, R.supply, +B.vendorCode, null, R.memo, null, null, null, null, null, null, +B.group, +B.artist, serial(R.start), serial(R.end), 'KR', null, null, null, null, null]];
    var ws = XLSX.utils.aoa_to_sheet(aoa); [20, 21].forEach(function (ci) { var c = ws[XLSX.utils.encode_cell({ r: 3, c: ci })]; if (c) c.z = 'yyyy-mm-dd'; });
    var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'AR거래내역엑셀업로드 List');
    ctx.saveFile('AR거래내역엑셀업로드_BPTG_' + R.qShort + '.xlsx', XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
    ctx.toast('AR 업로드 파일을 내려받았습니다');
  });
  $('btnSave').addEventListener('click', async function () {
    try {
      $('erpMsg').textContent = '저장 중…';
      await ctx.post('saveQuarter', { record: { quarter: 'GAME ' + R.qShort, supply: R.supply, vat: R.vat, gross: R.total, settleMonth: Q.months[2].ym, txDate: R.tx, payDate: R.pay, file: ctx.fileName,
        lines: Q.months.map(function (m) { return { ym: m.ym, net: m.net, share: m.share, vat: m.vat }; }), royaltyTable: [] } });
      $('erpMsg').textContent = ''; ctx.toast(R.qShort + ' 확정 저장했습니다');
    } catch (er) { $('erpMsg').textContent = '저장하지 못했습니다 (' + er.message + ')'; }
  });
  var TABS = ['calc', 'pum', 'erp'];
  function showTab(t) {
    if (TABS.indexOf(t) < 0) t = TABS[0];
    TABS.forEach(function (k) { $('p-' + k).hidden = k !== t; $('t-' + k).setAttribute('aria-selected', k === t ? 'true' : 'false'); });
    try { localStorage.setItem('bptg.tab', t); } catch (e) { }
  }
  ctx.tabsEl.querySelectorAll('.tab').forEach(function (b) { b.addEventListener('click', function () { showTab(b.getAttribute('data-tab')); window.scrollTo(0, 0); }); });
  var st = null; try { st = localStorage.getItem('bptg.tab'); } catch (e) { }
  render(); showTab(st || 'calc');
  return { quarter: Q.y + '.' + Q.q + 'Q', note: Q.sheet + ' · 분기 ' + P.quarters.length + '개 인식' };
}

DESK.register({ id: 'bptg-game', server: 'thegirls', name: '블랙핑크더게임 · 게임매출', desc: 'BPTG 게임 매출 로열티 23% · 분기 매출 (테이크원)',
  fileHint: 'BPTG_게임매출_YYYY.MM(또는 nQ)_정산서_YG.xlsx — 가장 최근에 받은 파일', outputs: ['정산서 숫자 검증', '매출품의', 'ERP 매출 전표 입력값'],
  detect: function (wb) { try { BPTG.parse(wb); return true; } catch (e) { return false; } },
  why: function (wb) { if (!wb.Sheets['지급정산서']) return ''; try { BPTG.parse(wb); return ''; } catch (e) { return e.message; } },
  mount: mount });
})();
