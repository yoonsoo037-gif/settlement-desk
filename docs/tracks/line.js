/* 정산 모듈: LINE 프리미엄 (LY Corporation 라인 프리미엄 스티커 로열티, 월)
 * 판별(detect) → 셸이 이 정산의 비공개 설정을 서버에서 받아 mount 호출 → 정산·품의·ERP 탭 렌더.
 * 회사 데이터(상품→코드 매핑·배분 비율·ERP 코드·품의 문구)는 이 파일에 없다(DB "LINE_설정" 시트). */
(function () {
/* LINE 프리미엄 정산 엔진 — 라인프리미엄.xlsx("라인스티커 신규" 시트) 로직 역산
 * (26.7월·26.8월 실제 AR 업로드 금액과 1엔 단위 일치 검증) */
var LP = (function () {
  var N = function (s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); };
  var MON = { January: 1, February: 2, March: 3, April: 4, May: 5, June: 6, July: 7, August: 8, September: 9, October: 10, November: 11, December: 12 };
  var MON3 = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  function num(v) { var n = typeof v === 'number' ? v : parseFloat(String(v == null ? '' : v).replace(/[,\s]/g, '')); return isFinite(n) ? n : null; }
  function fmt(n) { return Math.round(n).toLocaleString('ko-KR'); }
  function fmt2(n) { return n.toLocaleString('ko-KR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function rowsOf(ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }); }
  // 한 행에서 라벨 뒤에 오는 첫 숫자
  function after(row, i) { for (var j = i + 1; j < row.length; j++) { var n = num(row[j]); if (n != null && String(row[j]).trim() !== '') return n; } return null; }

  // 1) LY "Premium Note of Payment" 엑셀 해석 (Note of Payment + Breakdown 시트)
  function parse(wb) {
    var noteName = wb.SheetNames.filter(function (s) { return /note\s*of\s*payment/i.test(s); })[0];
    var bdName = wb.SheetNames.filter(function (s) { return /breakdown/i.test(s); })[0];
    if (!noteName || !bdName) throw new Error('"Note of Payment"·"Breakdown" 시트가 없습니다. LINE 프리미엄 Note_of_Payment 엑셀인지 확인해 주세요.');
    var nr = rowsOf(wb.Sheets[noteName]);
    if (!nr.some(function (r) { return r.some(function (c) { return /premium\s+note\s+of\s+payment/i.test(N(c)); }); }))
      throw new Error('Premium Note of Payment 문서가 아닙니다(테마 등 다른 LINE 정산서로 보입니다).');
    var p = { noteSheet: noteName };
    nr.forEach(function (r) {
      r.forEach(function (c, i) {
        var t = N(c);
        if (/^Date:?$/i.test(t)) { for (var j = i + 1; j < r.length; j++) if (r[j] != null && N(r[j])) { p.dateRaw = r[j]; break; } }
        if (/^Royalty amount/i.test(t)) p.total = after(r, i);
        if (/^Amount Paid/i.test(t)) p.paid = after(r, i);
        if (/^Deposit Amount/i.test(t)) p.net = after(r, i);
        if (/^Deposit Date/i.test(t)) { for (var m = i + 1; m < r.length; m++) if (r[m] != null && /\d/.test(String(r[m]))) { p.depRaw = r[m]; break; } }
        if (/^Premium sold in/i.test(t)) { for (var q = i + 1; q < r.length; q++) if (r[q] != null && N(r[q])) { p.periodRaw = r[q]; break; } }
      });
    });
    p.invoice = toYmd(p.dateRaw); p.deposit = toYmd(p.depRaw);
    var pm = String(p.periodRaw || '').match(/([A-Za-z]+)\s+(\d{4})/);
    if (pm && MON[pm[1]]) p.ym = pm[2] + '-' + pad(MON[pm[1]]);
    // Breakdown: 상품별 Total (소수점)
    var br = rowsOf(wb.Sheets[bdName]), items = [], tot = null;
    br.forEach(function (r) {
      var name = r.map(N).filter(function (c) { return /_(GL|JP)$/i.test(c); })[0];
      var nums = r.map(num).filter(function (n, i) { return n != null && String(r[i]).trim() !== ''; });
      if (name && nums.length) items.push({ key: name, count: nums.length > 1 ? nums[nums.length - 2] : null, raw: nums[nums.length - 1], side: /_JP$/i.test(name) ? 'JP' : 'GL' });
      else if (r.some(function (c) { return /^Total$/i.test(N(c)); }) && nums.length) tot = { count: nums.length > 1 ? nums[nums.length - 2] : null, amt: nums[nums.length - 1] };
      if (!p.ym) { var sm = r.map(N).join(' ').match(/Sales for (\d{4})\.(\d{2})/); if (sm) p.ym = sm[1] + '-' + sm[2]; }
    });
    if (!items.length) throw new Error('Breakdown 시트에서 상품(_GL/_JP) 행을 찾지 못했습니다.');
    if (p.total == null) throw new Error('Note of Payment 시트에서 "Royalty amount"를 찾지 못했습니다.');
    p.items = items; p.bdTotal = tot;
    return p;
  }
  function toYmd(v) {
    if (v == null) return '';
    if (typeof v === 'number') { var d = new Date(Math.round((v - 25569) * 864e5)); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
    var s = N(v), m;
    if ((m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/)) && MON3[m[2]]) return m[3] + '-' + pad(MON3[m[2]]) + '-' + pad(+m[1]);
    if ((m = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/)) && MON[m[2]]) return m[3] + '-' + pad(MON[m[2]]) + '-' + pad(+m[1]);
    if ((m = s.match(/^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/))) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
    return '';
  }

  // 2) 금액 계산
  //  ① 상품별 Total 반올림 → 합계가 정산서 총액과 다르면 "일본(_JP)" 금액만 조정 (26.7월: _JP 467.40 → 468)
  //  ② _JP 금액을 YGK/AVEX/YGJ 비율로 나눠 각각 반올림
  //  ③ 배분 합계 ≠ _JP 금액이면 같은 상품 묶음의 _GL 줄이 차이를 흡수 (라인프리미엄.xlsx 수식 · 26.7월 GL 105→106, 26.8월 90→89)
  function compute(p, cfg) {
    var PR = cfg.PRODUCTS || [], SP = cfg.JP_SPLIT || [], flags = [];
    var info = {}; PR.forEach(function (x) { info[N(x.key).toUpperCase()] = x; });
    var items = p.items.map(function (it) { return { key: it.key, side: it.side, raw: it.raw, count: it.count, amt: Math.round(it.raw), cfg: info[N(it.key).toUpperCase()] || null }; });
    items.filter(function (it) { return !it.cfg; }).forEach(function (it) { flags.push({ level: 'bad', text: '처음 보는 상품 "' + it.key + '" (' + it.raw + '엔) — DB "LINE_설정" PRODUCTS에 프로젝트 코드를 추가해야 ERP 줄이 만들어집니다' }); });
    var rawSum = items.reduce(function (a, it) { return a + it.raw; }, 0), rSum = items.reduce(function (a, it) { return a + it.amt; }, 0);
    // 정산서 자체 검증
    if (p.paid != null && p.paid !== p.total) flags.push({ level: 'bad', text: '정산서 Royalty amount ' + fmt(p.total) + '엔 ≠ Amount Paid ' + fmt(p.paid) + '엔' });
    if (Math.round(rawSum) !== p.total) flags.push({ level: 'bad', text: 'Breakdown 상품 Total 합계 ' + rawSum.toFixed(2) + '엔이 총액 ' + fmt(p.total) + '엔과 맞지 않습니다' });
    else flags.push({ level: 'ok', text: 'Breakdown 상품 ' + items.length + '개 Total 합계 ' + rawSum.toFixed(2) + '엔 → 반올림 ' + fmt(p.total) + '엔 = 정산서 총액' });
    if (p.bdTotal && p.bdTotal.count != null) {
      var cnt = items.reduce(function (a, it) { return a + (it.count || 0); }, 0);
      if (cnt !== p.bdTotal.count) flags.push({ level: 'warn', text: 'Breakdown 판매수량 합계 ' + cnt + ' ≠ Total 행 ' + p.bdTotal.count });
    }
    var wht = p.net != null ? p.total - p.net : null, whtCalc = Math.floor(p.total * 0.1);
    if (p.net == null) flags.push({ level: 'bad', text: '정산서에서 실입금액(Deposit Amount)을 찾지 못했습니다' });
    else if (wht !== whtCalc) flags.push({ level: 'warn', text: '원천세 ' + fmt(wht) + '엔(총액 − 실입금)이 총액 × 10% 절사 ' + fmt(whtCalc) + '엔과 다릅니다 — 정산서 금액 기준으로 진행' });
    else flags.push({ level: 'ok', text: '원천세 10% ' + fmt(wht) + '엔 (총액 ' + fmt(p.total) + ' × 10% 절사) · 실입금 ' + fmt(p.net) + '엔 = 정산서 Deposit Amount' });
    // ① 총액 맞춤: 일본(_JP)만 조정
    var adj = p.total - rSum, adjNote = '';
    if (adj) {
      var jp = items.filter(function (it) { return it.side === 'JP'; }).sort(function (a, b) { return b.raw - a.raw; })[0];
      var tgt = jp || items.slice().sort(function (a, b) { return adj > 0 ? (b.raw - b.amt) - (a.raw - a.amt) : (a.raw - a.amt) - (b.raw - b.amt); })[0];
      tgt.amt += adj; tgt.adj = adj;
      adjNote = tgt.key + ' ' + (adj > 0 ? '+' : '') + adj + '엔';
      flags.push({ level: jp ? 'ok' : 'warn', text: '반올림 합계 ' + fmt(rSum) + '엔 ≠ 총액 ' + fmt(p.total) + '엔 → ' + (jp ? '일본(_JP) 금액만 조정(회계팀 지침): ' : '일본(_JP) 상품이 없어 소수점이 가장 큰 상품을 조정: ') + adjNote });
    }
    // ②·③ 일본 배분 + 같은 묶음 _GL 줄 흡수
    var blocks = [], byBlock = {};
    items.forEach(function (it) {
      var b = it.cfg ? it.cfg.block : '?' + it.key;
      if (!byBlock[b]) { byBlock[b] = { key: b, gl: null, jp: null, label: it.cfg ? it.cfg.label : it.key, cfg: it.cfg }; blocks.push(byBlock[b]); }
      byBlock[b][it.side === 'JP' ? 'jp' : 'gl'] = it;
    });
    // ERP 줄은 설정(PRODUCTS) 순서 — 라인프리미엄.xlsx·26.7월 AR 순서
    var order = {}; PR.forEach(function (x, i) { if (!(x.block in order)) order[x.block] = i; });
    blocks.sort(function (a, b) { return (a.key in order ? order[a.key] : 999) - (b.key in order ? order[b.key] : 999); });
    var lines = [], absorbed = [];
    blocks.forEach(function (b) {
      var glAmt = b.gl ? b.gl.amt : 0, split = [];
      if (b.jp) {
        split = SP.map(function (s) { return { tag: s.tag, rate: s.rate, amt: Math.round(b.jp.amt * s.rate) }; });
        var diff = b.jp.amt - split.reduce(function (a, s) { return a + s.amt; }, 0);
        if (diff) {
          if (b.gl) { glAmt += diff; absorbed.push(b.label + ' 일본외 ' + fmt(b.gl.amt) + '→' + fmt(glAmt) + '엔 (일본 ' + fmt(b.jp.amt) + '엔 배분 ' + split.map(function (s) { return s.amt; }).join('/') + ')'); }
          else { split[0].amt += diff; absorbed.push(b.label + ' ' + split[0].tag + ' ' + (diff > 0 ? '+' : '') + diff + '엔 (같은 묶음 일본외 줄 없음)'); }
        }
      }
      if (b.gl || b.jp) lines.push({ block: b.key, label: b.label, tag: '일본외', amt: glAmt, cfg: b.cfg, src: b.gl });
      split.forEach(function (s) { lines.push({ block: b.key, label: b.label, tag: s.tag, amt: s.amt, cfg: b.cfg, src: b.jp, rate: s.rate }); });
    });
    if (absorbed.length) flags.push({ level: 'ok', text: '일본 3사 배분 반올림 차이 처리(라인프리미엄.xlsx 방식): ' + absorbed.join(', ') });
    var lineSum = lines.reduce(function (a, l) { return a + l.amt; }, 0);
    if (lineSum !== p.total) flags.push({ level: 'bad', text: 'AR 줄 합계 ' + fmt(lineSum) + '엔 ≠ 총액 ' + fmt(p.total) + '엔' });
    else flags.push({ level: 'ok', text: 'AR 줄 합계 ' + fmt(lineSum) + '엔 = 정산서 총액 (0엔 줄 ' + lines.filter(function (l) { return !l.amt; }).length + '개는 업로드에서 제외)' });
    return { items: items, lines: lines, upload: lines.filter(function (l) { return l.amt && l.cfg; }), total: p.total, net: p.net, wht: wht, flags: flags, adjNote: adjNote };
  }

  // 적요: 라인프리미엄.xlsx 수식과 같은 표기 ("…_위너/ JPY1,429 (일본외)")
  function memo(ym, l) { return ym.slice(0, 4) + '.' + (+ym.slice(5)) + ' 라인 프리미엄 로열티 매출_' + l.label + '/ JPY' + fmt(l.amt) + ' (' + l.tag + ')'; }
  function serial(ymd) { var t = ymd.split('-'); return Math.round((Date.UTC(+t[0], +t[1] - 1, +t[2]) - Date.UTC(1899, 11, 30)) / 86400000); }
  function monthEnd(ym) { var y = +ym.slice(0, 4), m = +ym.slice(5); return ym + '-' + pad(new Date(Date.UTC(y, m, 0)).getUTCDate()); }

  // ERP "AR 거래내역 일괄생성 업로드" 양식 (26.7월 AR_라인프리미엄 업로드 파일과 같은 열·값 구성)
  function buildErpUpload(p, r, cfg, rate100) {
    var P = cfg.PROJECT || {}, fx = Math.round(rate100 * 100) / 1e4;
    var head = ['프로젝트', '회계부서', '계정코드', '예산코드', '배부코드', '통화', '환율', '차변금액', '대변금액', '거래처', '반제전표', '적요', '세금일자', '부가세코드', '공급가액', '부가세금액', '신고사업장', '세금계산서번호', '아티스트그룹', '아티스트', '관리항목1', '관리항목2', '관리항목3', '관리항목4', '관리항목5', '관리항목6', '관리항목7', '관리항목8'];
    var aoa = [['AR 거래내역 일괄생성 업로드 양식'], [], head];
    r.upload.forEach(function (l) {
      aoa.push([+l.cfg.proj, +P.dept, String(P.account), null, null, 'JPY', fx, null, l.amt, String(P.vendor), null, memo(p.ym, l), null, null, null, null, null, null,
        +l.cfg.grp, +P.artist, serial(p.ym + '-01'), serial(monthEnd(p.ym)), P.country || 'JP', null, null, null, null, null]);
    });
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    // 관리항목1·2(정산 시작·종료일)는 기존 AR 원본처럼 날짜 서식 — 일반 숫자면 ERP가 46204처럼 숫자로 읽는다
    for (var ri = 3; ri < aoa.length; ri++) [20, 21].forEach(function (ci) { var c = ws[XLSX.utils.encode_cell({ r: ri, c: ci })]; if (c && c.t === 'n') c.z = 'm/d/yy'; });
    XLSX.utils.book_append_sheet(wb, ws, 'AR거래내역엑셀업로드 List');
    return { wb: wb, rows: aoa.slice(3), fx: fx };
  }

  return { parse: parse, compute: compute, memo: memo, buildErpUpload: buildErpUpload, toYmd: toYmd, monthEnd: monthEnd, fmt: fmt, fmt2: fmt2, pad: pad };
})();
if (typeof module !== 'undefined') module.exports = LP;
if (typeof DESK === 'undefined') return;

// ---------------- 화면 ----------------
var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
var fmt = LP.fmt, fmt2 = LP.fmt2;
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
  return { cls: 'status ' + lvl, html: '<b>' + (nb ? '조치 필요 ' + nb + '건' : nw ? '체크 ' + nw + '건' : '정산서 숫자 검증 이상 없음') + '</b><span>' + esc((fl.filter(function (f) { return f.level === lvl; })[0] || {}).text || '') + '</span>' };
}

function mount(ctx) {
  var $ = function (id) { return document.getElementById(id); };
  var CFG = (ctx.data && ctx.data.cfg) || {}, P = CFG.PROJECT || {}, PUM = CFG.PUM || {};
  var p = LP.parse(ctx.wb), R = LP.compute(p, CFG), TXT = {}, TABLES = {};
  var saved = (ctx.data.saved || {})['PREMIUM ' + p.ym];
  var S = { rate: '' };
  var mNum = +p.ym.slice(5), yNum = p.ym.slice(0, 4);

  ctx.tabsEl.innerHTML = tabsHtml([['calc', '정산'], ['pum', '품의 복붙'], ['erp', 'ERP 전표']]);
  ctx.root.innerHTML =
    '<div class="tabpanel" id="p-calc"><section class="panel">' +
    '<div class="form"><label>환율 (인보이스 발행일 ' + esc(p.invoice || '-') + ', 100엔당 원)<input id="inFx" inputmode="decimal" placeholder="예: 903.72"></label></div>' +
    '<p class="small muted" style="margin:0">환율: <a href="http://www.smbs.biz/ExRate/StdExRate.jsp" target="_blank" rel="noopener">서울외국환중개 매매기준율</a>에서 인보이스 발행일(정산서 Date) JPY(100) 값을 넣어 주세요. ERP 환율은 이 값 ÷ 100 입니다 (26.7월: 903.72 → 9.0372).</p></section>' +
    '<section class="panel dash"><div class="tiles" id="tiles"></div><div class="status" id="statusLine"></div><div class="meta" id="amtMeta"></div>' +
    '<details class="fold vfold" open><summary>정산서 숫자 검증</summary><div class="flags" id="flags"></div></details></section>' +
    '<section class="panel"><div class="ph"><h2>상품별 금액 → ERP 줄</h2><p>정산서 Breakdown 소수점 금액을 반올림하고, 일본(_JP)은 YGK·AVEX·YGJ 3사로 나눕니다.</p></div><div class="tablebox"><table id="itemTable"></table></div><div class="tablebox"><table id="lineTable"></table></div></section>' +
    '<details class="panel fold logic"><summary><b>적용 중인 정산 로직</b> <span class="small muted">라인프리미엄.xlsx "라인스티커 신규" 시트 · 26.7월/8월 AR 기준</span></summary><ol>' +
    '<li><b>상품별 금액</b>: Breakdown Total(소수점)을 엔 단위 반올림. 합계가 정산서 총액과 다르면 <b>일본(_JP) 금액만</b> 조정(회계팀 지침).</li>' +
    '<li><b>일본 3사 배분</b>: _JP 금액 × ' + esc((CFG.JP_SPLIT || []).map(function (s) { return s.tag.replace('일본수익 중 ', '').replace(' 몫', '') + ' ' + Math.round(s.rate * 100) + '%'; }).join(' / ')) + ' 각각 반올림.</li>' +
    '<li><b>배분 차이</b>: 3사 합계가 _JP 금액과 다르면 같은 상품의 일본외(_GL) 줄이 차이를 흡수(라인프리미엄.xlsx 수식, 26.7월 105→106 · 26.8월 90→89).</li>' +
    '<li><b>원천세</b>: 총액 − 실입금(정산서 기준, 총액 10% 절사와 대조). <b>AR</b>: 금액 0인 줄 제외, 관리항목 = 판매월 1일~말일 · JP.</li></ol></details></div>' +

    '<div class="tabpanel" id="p-pum" hidden><section class="panel"><div class="ph"><h2>매출품의 입력 내용</h2><p>직전 품의(YGK-광고IP사업팀-26-0622, 26.7월분) 양식·문구 그대로입니다. 첨부: 정산서 PDF.</p></div>' +
    '<div class="allcopy"><button class="btn primary" data-copy="all">본문 전체 복사</button></div>' +
    '<div class="block"><header><h3>제목</h3><button class="btn sm" data-copy="title">복사</button></header><div class="body"><div id="pTitle"></div></div></div>' +
    '<div class="block"><header><h3>필수 입력 · 1. 기본 내역 · 2. 발행 정보</h3><button class="btn sm" data-copy="basic">전체 복사</button></header><div class="body"><dl class="kv" id="pBasic"></dl></div></div>' +
    '<div class="block"><header><h3>3. 상세내역</h3><button class="btn sm" data-copy="detail">문구 복사</button></header><div class="body"><pre class="fmt" id="pDetail"></pre>' +
    '<div class="row" style="justify-content:space-between"><b class="small">로열티 상세 내역 표</b><button class="btn sm" data-copy="royTable">표 복사</button></div><div class="tablebox"><table id="royTable"></table></div>' +
    '<div class="row" style="justify-content:space-between"><b class="small">매출금액 표</b><button class="btn sm" data-copy="paidTable">표 복사</button></div><div class="tablebox"><table id="paidTable"></table></div>' +
    '<pre class="fmt" id="pDetail2"></pre></div></div></section></div>' +

    '<div class="tabpanel" id="p-erp" hidden><section class="panel"><div class="ph"><h2>ERP AR 업로드</h2><p>내려받은 파일을 ERP "거래내역 엑셀 업로드"로 올리면 전표가 생성됩니다. 회계일자는 인보이스 발행일입니다.</p></div>' +
    '<div class="row"><button class="btn primary" id="btnErp">AR 업로드 파일 내려받기</button><button class="btn" id="btnSave">이 달 확정 저장</button><span class="small muted" id="erpMsg"></span></div>' +
    '<div class="block"><header><h3>전표 헤더</h3><button class="btn sm" data-copy="head">전체 복사</button></header><div class="body"><dl class="kv" id="pHead"></dl></div></div>' +
    '<div class="block"><header><h3>AR 줄 (업로드 파일 내용)</h3><button class="btn sm" data-copy="arTable">표 복사</button></header><div class="body"><div class="tablebox"><table id="arTable"></table></div></div></div></section></div>';

  function rate100() { var v = parseFloat(String(S.rate).replace(/,/g, '')); return isFinite(v) && v > 0 ? v : null; }
  function render() {
    var fx = rate100(), fxE = fx ? Math.round(fx * 100) / 1e4 : null, flags = R.flags.slice();
    if (!fx) flags.unshift({ level: 'warn', text: '환율을 입력하면 AR 파일·원화 환산이 완성됩니다 (인보이스 발행일 ' + (p.invoice || '-') + ' 서울외국환중개 JPY100 매매기준율)' });
    if (saved) flags.push({ level: 'ok', text: '이 판매월(' + p.ym + ')은 ' + (saved.savedAt || '') + ' 확정 저장된 기록이 있습니다 (' + fmt(saved.supply) + '엔)' });
    var krw = fxE ? R.upload.reduce(function (a, l) { return a + Math.round(l.amt * fxE); }, 0) : null;
    $('tiles').innerHTML = [['판매월', p.ym], ['총액 (원천세 전)', fmt(R.total) + '엔'], ['원천세 10%', R.wht == null ? '-' : fmt(R.wht) + '엔'], ['실입금', R.net == null ? '-' : fmt(R.net) + '엔'], ['원화 환산', krw == null ? '환율 입력' : fmt(krw) + '원']]
      .map(function (t) { return '<div class="tile"><span>' + esc(t[0]) + '</span><b>' + esc(t[1]) + '</b></div>'; }).join('');
    var st = statusHtml(flags); $('statusLine').className = st.cls; $('statusLine').innerHTML = st.html;
    $('amtMeta').textContent = '인보이스 ' + (p.invoice || '-') + ' · 입금예정 ' + (p.deposit || '-') + ' · ' + ctx.fileName;
    $('flags').innerHTML = flagsHtml(flags);

    TABLES.items = { head: ['상품', '구분', '판매수량', '정산서 Total', '반올림', '조정 후'], numCols: [2, 3, 4, 5], sumLast: true,
      rows: R.items.map(function (it) { return [it.key, it.side === 'JP' ? '일본' : '일본외', it.count == null ? '-' : fmt(it.count), fmt2(it.raw), fmt(Math.round(it.raw)), fmt(it.amt) + (it.adj ? ' (' + (it.adj > 0 ? '+' : '') + it.adj + ')' : '')]; })
        .concat([['합계', '', fmt(R.items.reduce(function (a, it) { return a + (it.count || 0); }, 0)), fmt2(R.items.reduce(function (a, it) { return a + it.raw; }, 0)), fmt(R.items.reduce(function (a, it) { return a + Math.round(it.raw); }, 0)), fmt(R.items.reduce(function (a, it) { return a + it.amt; }, 0))]]) };
    $('itemTable').innerHTML = tableHtml(TABLES.items);
    TABLES.lines = { head: ['ERP 줄', '구분', '금액(JPY)', '원화'], numCols: [2, 3], sumLast: true,
      rows: R.lines.map(function (l) { return [l.label, l.tag + (l.rate ? ' ' + Math.round(l.rate * 100) + '%' : '') + (l.amt ? '' : ' · 업로드 제외'), fmt(l.amt), fxE ? fmt(Math.round(l.amt * fxE)) : '-']; })
        .concat([['합계', '', fmt(R.lines.reduce(function (a, l) { return a + l.amt; }, 0)), krw == null ? '-' : fmt(krw)]]) };
    $('lineTable').innerHTML = tableHtml(TABLES.lines);

    // ---- 품의 ----
    var title = '[매출품의] ' + yNum + '년 라인프리미엄 YG 아티스트 아이템 ' + mNum + '월 매출 정산';
    var residual = PUM.endYm && p.ym > PUM.endYm;
    TXT.title = title; $('pTitle').textContent = title;
    var basic = [['거래처명', (P.vendorName || '') + ' (' + (P.vendor || '') + ')'], ['거래일자', p.invoice || '-'], ['금액 (원천세 전)', fmt(R.total)], ['총금액 (원천세 차감)', R.net == null ? '-' : fmt(R.net)],
      ['프로젝트', '프로젝트공통 (_)'], ['입금예정일', p.deposit || '-'], ['아티스트(그룹)', '그룹공통 (_)'], ['아티스트(개인)', (P.artistName || '') + ' (' + (P.artist || '') + ')'],
      ['매출 종류', '로열티'], ['통화종류', 'JPY'], ['비거주자 여부', '비거주자'], ['발행정보', 'INVOICE'], ['별도요청사항', PUM.request || '']];
    TXT.basic = basic.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
    $('pBasic').innerHTML = kv(basic, 'basic');
    var invMonth = p.invoice ? p.invoice.slice(0, 7) : '';
    var d1 = '1. 로열티 상세 내역 - 정산월 : ' + invMonth + '\n- 프로젝트 : ' + (PUM.proj || '') + '\n- 사업구조 : ' + (PUM.struct || 'RS') + '\n- MG : -\n- 통화 : JPY\n- 환율 : ' + (fxE ? String(fxE) : '(환율 입력)');
    TABLES.royTable = { head: ['판매월', '아티스트', '권역', '판매수량', '총 매출 (vat-)', '로열티 (vat-)'], numCols: [5], sumLast: false,
      rows: R.items.map(function (it) { return [p.ym, it.key, it.side === 'JP' ? '일본' : '글로벌', '-', '-', fmt2(it.raw)]; }) };
    TABLES.paidTable = { head: ['Artist', 'Amount Paid (tax incl.)', '스티커 출시일'], numCols: [1], sumLast: true,
      rows: R.items.map(function (it) { return [it.key, fmt2(it.raw), it.cfg ? it.cfg.release || '' : '']; }).concat([['TOTAL', fmt(R.total), '']]) };
    var won = fxE ? Math.round(R.total * fxE) : null;
    var wonTxt = won == null ? '(환율 입력)' : won >= 1e4 ? '약 ' + Math.floor(won / 1e4) + '만원' : '약 ' + fmt(won) + '원';
    var invDot = p.invoice ? p.invoice.slice(2).replace(/-0?/g, '.') : '';
    var d2 = '2. 매출금액 : (tax incl.) : ' + fmt(R.total) + ' 円 (tax 제외) : ' + (R.net == null ? '-' : fmt(R.net)) + ' 円\n\n' +
      '*상기 금액은 엔화 표기이며, 원화 환산시 ' + wonTxt + '입니다. (JPY 100 = ' + (fx ? fmt2(fx) : '(환율)') + '원/' + invDot + ' 인보이스 발행일 환율 적용)\n\n' +
      '3. 첨부파일 : ' + mNum + '월 라인프리미엄 매출리포트\n\n' +
      '4. 기타 참고사항 : - ' + (PUM.note || '') +
      (residual ? '\n- ' + (+PUM.endYm.slice(2, 4)) + '.' + (+PUM.endYm.slice(5)) + '월말 판매 종료 이후 라인 서비스 중단이 지연되어 발생한 잔여 매출(' + yNum.slice(2) + '.' + mNum + '월분) 정산 건.' : '') + '\n\n이상입니다.';
    TXT.detail = d1 + '\n\n판매월\t아티스트\t권역\t판매수량\t총 매출 (vat-)\t로열티 (vat-)\n' + TABLES.royTable.rows.map(function (r) { return r.join('\t'); }).join('\n') + '\n\n' + d2;
    $('pDetail').textContent = d1; $('pDetail2').textContent = d2;
    $('royTable').innerHTML = tableHtml(TABLES.royTable); $('paidTable').innerHTML = tableHtml(TABLES.paidTable);
    TXT.all = title + '\n\n' + TXT.basic + '\n\n' + TXT.detail;

    // ---- ERP ----
    var head = [['회계일자', p.invoice || '-'], ['증빙구분', '기타'], ['청구처', P.vendorName || ''], ['통화 / 환율', 'JPY / ' + (fxE ? fxE : '(환율 입력)')],
      ['제목', yNum + ' 라인프리미엄 YG 아티스트 아이템 ' + mNum + '월 매출'], ['외화 합계', fmt(R.total) + ' JPY (채권 ' + (R.net == null ? '-' : fmt(R.net)) + ' / 원천세 ' + (R.wht == null ? '-' : fmt(R.wht)) + ')'],
      ['원화 공급가액', krw == null ? '-' : fmt(krw) + '원 (줄별 원 단위 반올림 합)'], ['부가세', '0 (영세·비거주자)']];
    TXT.head = head.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
    $('pHead').innerHTML = kv(head, 'head');
    var P2 = P;
    TABLES.arTable = { head: ['프로젝트', '그룹', '대변금액(JPY)', '적요', '관리항목1', '관리항목2'], numCols: [2], sumLast: true,
      rows: R.upload.map(function (l) { return [l.cfg.proj + ' ' + l.cfg.projName, l.cfg.grp, fmt(l.amt), LP.memo(p.ym, l), p.ym + '-01', LP.monthEnd(p.ym)]; })
        .concat([['합계', '', fmt(R.upload.reduce(function (a, l) { return a + l.amt; }, 0)), P2.vendorName + ' · ' + (P2.country || 'JP'), '', '']]) };
    $('arTable').innerHTML = tableHtml(TABLES.arTable);
    $('btnErp').disabled = !fx || R.upload.length !== R.lines.filter(function (l) { return l.amt; }).length;
    $('btnSave').disabled = !fx;
    $('erpMsg').textContent = fx ? '' : '환율을 입력하면 내려받을 수 있습니다';
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
    try { localStorage.setItem('line.prem.tab', t); } catch (e) { }
  }
  ctx.tabsEl.querySelectorAll('.tab').forEach(function (b) { b.addEventListener('click', function () { showTab(b.getAttribute('data-tab')); window.scrollTo(0, 0); }); });
  var st = null; try { st = localStorage.getItem('line.prem.tab'); } catch (e) { }

  $('inFx').addEventListener('input', function () { S.rate = this.value.trim(); clearTimeout(render._t); render._t = setTimeout(render, 200); });
  $('btnErp').addEventListener('click', function () {
    var out = LP.buildErpUpload(p, R, CFG, rate100());
    ctx.saveFile('AR거래내역엑셀업로드_라인프리미엄(' + mNum + '월).xlsx', XLSX.write(out.wb, { type: 'array', bookType: 'xlsx' }));
    ctx.toast('AR 업로드 파일을 내려받았습니다 (' + out.rows.length + '줄)');
  });
  $('btnSave').addEventListener('click', async function () {
    try {
      $('erpMsg').textContent = '저장 중…';
      await ctx.post('saveQuarter', { record: { quarter: 'PREMIUM ' + p.ym, supply: R.total, vat: R.wht, gross: R.net, settleMonth: p.ym, txDate: p.invoice, payDate: p.deposit, file: ctx.fileName,
        roundNote: R.adjNote, lines: R.upload.map(function (l) { return { label: l.label, tag: l.tag, amt: l.amt }; }), royaltyTable: [{ rate100: rate100() }] } });
      $('erpMsg').textContent = ''; ctx.toast(p.ym + ' 라인프리미엄을 확정 저장했습니다');
    } catch (e) { $('erpMsg').textContent = '저장하지 못했습니다 (' + e.message + ')'; }
  });

  render(); showTab(st || 'calc');
  return { quarter: p.ym + ' 판매분', note: p.noteSheet + ' · 상품 ' + p.items.length + '개 · 인보이스 ' + (p.invoice || '-') };
}

DESK.register({ id: 'line-premium', server: 'line', name: 'LINE 프리미엄', desc: 'LY Corporation 라인 프리미엄 스티커 로열티 · 월 매출 (JPY)',
  fileHint: 'LINE スタンプ（月額） Note_of_Payment_<월>_<연도>.xls', outputs: ['정산서 숫자 검증', '매출품의', 'ERP AR 업로드 파일'],
  detect: function (wb) { try { LP.parse(wb); return true; } catch (e) { return false; } },
  why: function (wb) { if (!wb.SheetNames.some(function (s) { return /breakdown/i.test(s); })) return ''; try { LP.parse(wb); return ''; } catch (e) { return e.message; } },
  mount: mount });
})();
