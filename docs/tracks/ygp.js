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
    p.gross = grossOf(wb);
    return p;
  }
  // 품의 판매월 표 "총 매출 (vat-)": "제품매출" 시트 상세 행 중 정산대상(YGK(정산대상)) 매출 "금액"을 아티스트별 합산.
  // 아티스트 키 = "추가구분"(예: BABYMONSTER(라미포함)·WINNER(위너)(민호제외))이 있으면 그것, 없으면 "품목군" — ② 아티스트별 표기와 같음 (26.10 품의 34행 일치)
  function grossOf(wb) {
    var ws = wb.Sheets['제품매출']; if (!ws) return null;
    var rows = rowsOf(ws), out = {}, h = -1, c;
    for (var i = 0; i < Math.min(rows.length, 300); i++) {
      var r = (rows[i] || []).map(N);
      if (r.indexOf('매출일자') >= 0 && r.indexOf('정산대상여부') >= 0 && r.indexOf('금액') >= 0 && r.indexOf('품목군') >= 0) { h = i; c = { t: r.indexOf('정산대상여부'), g: r.indexOf('품목군'), x: r.indexOf('추가구분'), a: r.indexOf('금액') }; break; }
    }
    if (h < 0) return null;
    for (var j = h + 1; j < rows.length; j++) {
      var x = rows[j] || []; if (N(x[c.t]) !== 'YGK(정산대상)') continue;
      var key = (c.x >= 0 && N(x[c.x])) || N(x[c.g]); if (!key) continue;
      out[K(key)] = (out[K(key)] || 0) + (num(x[c.a]) || 0);
    }
    return out;
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

  // 빅뱅: YGP-빅뱅 마스터 계약 날인 전까지 모든 품의·ERP는 빅뱅 제외 (26.8월분부터, YGP가 원본(빅뱅 포함)·빅뱅 제외 정산서 2종 전달)
  var isBig = function (name) { return /빅뱅/.test(name); };
  function bigOf(p) { var r = p.artists.filter(function (x) { return isBig(x.name); }); return r.length ? r.reduce(function (a, x) { return { name: '빅뱅', qty: a.qty + x.qty, supply: a.supply + x.supply, vat: a.vat + x.vat, total: a.total + x.total }; }, { qty: 0, supply: 0, vat: 0, total: 0 }) : null; }
  // 원본(빅뱅 포함)에서 빅뱅 행을 빼 빅뱅 제외 기준을 만든다 (26.8·9월: 결과가 YGP 빅뱅 제외 정산서와 1원까지 같음)
  function exclude(p) {
    var b = bigOf(p) || { qty: 0, supply: 0, vat: 0, total: 0 }, inv = p.invoice;
    var q = Object.assign({}, p, { artists: p.artists.filter(function (x) { return !isBig(x.name); }), invoice: Object.assign({}, inv, { supply: inv.supply - b.supply, vat: inv.vat - b.vat, total: inv.total - b.total }) });
    if (p.artistTotal) q.artistTotal = { qty: p.artistTotal.qty - b.qty, supply: p.artistTotal.supply - b.supply, vat: p.artistTotal.vat - b.vat, total: p.artistTotal.total - b.total };
    return q;
  }
  // 원본 − 빅뱅 = 빅뱅 제외본 대조 (발행내역 + 아티스트별 수량·공급가액·부가세·공급대가)
  function compare(orig, excl) {
    var d = exclude(orig), out = [], bm = {};
    if (orig.ym !== excl.ym) out.push('정산월이 다릅니다(원본 ' + orig.ym + ' / 제외본 ' + excl.ym + ')');
    if (orig.invoice.date !== excl.invoice.date) out.push('세금계산서 작성일이 다릅니다(원본 ' + orig.invoice.date + ' / 제외본 ' + excl.invoice.date + ')');
    ['supply', 'vat', 'total'].forEach(function (k) { if (d.invoice[k] !== excl.invoice[k]) out.push('발행내역 ' + { supply: '공급가액', vat: '세액', total: '합계' }[k] + ': 원본−빅뱅 ' + fmt(d.invoice[k]) + ' ≠ 제외본 ' + fmt(excl.invoice[k])); });
    excl.artists.forEach(function (x) { bm[K(x.name)] = x; });
    d.artists.forEach(function (x) {
      var y = bm[K(x.name)]; delete bm[K(x.name)];
      if (!y) { if (x.supply) out.push(x.name + ': 원본에만 있음(' + fmt(x.supply) + '원)'); return; }
      ['qty', 'supply', 'vat', 'total'].forEach(function (k) { if (x[k] !== y[k]) out.push(x.name + ' ' + { qty: '수량', supply: '공급가액', vat: '부가세', total: '공급대가' }[k] + ': 원본 ' + fmt(x[k]) + ' ≠ 제외본 ' + fmt(y[k])); });
    });
    Object.keys(bm).forEach(function (k) { if (bm[k].supply) out.push(bm[k].name + ': 제외본에만 있음(' + fmt(bm[k].supply) + '원)'); });
    return { diffs: out, rows: d.artists.length };
  }

  return { parse: parse, K: K, compute: compute, memo: memo, buildErpUpload: buildErpUpload, bigOf: bigOf, exclude: exclude, compare: compare, fmt: fmt, monthEnd: monthEnd, addMonths: addMonths };
})();
if (typeof module !== 'undefined') module.exports = YGP;
if (typeof DESK === 'undefined') return;

// ---------------- 화면 ----------------
var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
var fmt = YGP.fmt;
var KIND = { md: { name: 'MD 로열티', short: 'MD', mail: 'MD 제품 로열티' }, cb: { name: '응원봉 디자인권', short: '응원봉', mail: '응원봉 디자인권 및 상표권 사용료' } };
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
var tsv = function (t) { return [t.head].concat(t.rows).map(function (r) { return r.join('\t'); }).join('\n'); };
// 위임전결기준표(261006) — 사업부문 매출 품의
// 위임전결기준표(261006) 1.1.2 매출 품의 · 사업 부문. 광고IP사업팀은 팀 리더·그룹 리더 결재를 그룹리더 1명이 받음(26.10 YGP 품의 26-0852·0853 실제)
function approval(supply) {
  var NOTE = ' — 광고IP사업팀은 팀 리더·그룹 리더 결재를 그룹리더가 한 번에 받음';
  if (supply < 5e6) return '500만원 미만 (합의종결): 결재 팀 리더(최종) · 합의: 재무관리팀 담당 & 회계팀 담당' + NOTE;
  if (supply < 5e7) return '500만원 이상 5,000만원 미만 (합의종결): 결재 팀 리더 → 그룹 리더(최종) · 합의: 재무관리팀 담당 & 회계팀 담당 → 담당 팀 리더(재무관리그룹 리더)' + NOTE;
  if (supply < 3e8) return '5,000만원 이상 3억원 미만: 결재 팀 리더 → 그룹 리더 · 합의: 재무관리팀 담당 & 회계팀 담당 → 담당 팀 리더(재무관리그룹 리더) → CFO(경영지원부문 대표) → 결재 부문 대표(최종)' + NOTE;
  return '3억원 이상: 결재 팀 리더 → 그룹 리더 · 합의: 재무관리팀 담당 & 회계팀 담당 → 담당 팀 리더(재무관리그룹 리더) → CFO(경영지원부문 대표) → 부문 대표 → 대표이사(최종)' + NOTE;
}

// 올린 파일(1~4개)을 MD/응원봉 × 원본(빅뱅 포함)/빅뱅 제외본으로 나누고, 빅뱅 제외 기준(base)을 정한다
function group(files, cfg) {
  var G = {}, flags = [];
  files.forEach(function (f) {
    var p = YGP.parse(f.wb), big = YGP.bigOf(p), role = big && (big.supply || big.qty) ? 'orig' : 'excl';
    var g = G[p.kind] = G[p.kind] || { kind: p.kind, flags: [] };
    if (g[role]) g.flags.push({ level: 'bad', text: (role === 'orig' ? '원본(빅뱅 포함)' : '빅뱅 제외본') + '이 2개 올라왔습니다: ' + g[role].file + ', ' + f.name + ' — 하나만 올려 주세요' });
    g[role] = { p: p, file: f.name, big: big };
  });
  var yms = Object.keys(G).map(function (k) { var g = G[k]; return (g.excl || g.orig).p.ym; });
  if (yms.some(function (y) { return y !== yms[0]; })) flags.push({ level: 'bad', text: 'MD·응원봉 파일의 정산월이 다릅니다: ' + yms.join(', ') });
  ['md', 'cb'].forEach(function (k) {
    var g = G[k]; if (!g) return;
    var nm = KIND[k].name;
    if (g.orig && g.excl) {
      var c = YGP.compare(g.orig.p, g.excl.p);
      if (c.diffs.length) g.flags.push({ level: 'bad', text: nm + ' 원본 − 빅뱅 ≠ 빅뱅 제외본 (' + c.diffs.length + '건): ' + c.diffs.slice(0, 6).join(' · ') + (c.diffs.length > 6 ? ' 외' : '') });
      else g.flags.push({ level: 'ok', text: nm + ' 원본 − 빅뱅(' + fmt(g.orig.big.supply) + '원) = 빅뱅 제외본: 발행내역(공급가액·세액·합계)과 아티스트 ' + c.rows + '명 수량·공급가액·부가세·공급대가 모두 일치' });
      g.base = g.excl.p;
    } else if (g.excl) {
      g.base = g.excl.p;
      g.flags.push({ level: 'warn', text: nm + ' 원본(빅뱅 포함) 정산서가 없습니다 — 빅뱅 제외본으로 산출물을 만들었고, 원본 대조와 회계팀 빅뱅 메일은 원본을 함께 올려야 완성됩니다' });
    } else {
      g.base = YGP.exclude(g.orig.p);
      g.flags.push({ level: 'warn', text: nm + ' 빅뱅 제외본이 없습니다 — 원본에서 빅뱅 행(' + fmt(g.orig.big.supply) + '원)을 빼 산출물을 만들었습니다. YGP 빅뱅 제외 정산서를 받으면 함께 올려 대조해 주세요' });
    }
    g.R = YGP.compute(g.base, cfg);
  });
  return { G: G, flags: flags, ym: yms[0] };
}

function mount(ctx) {
  var $ = function (id) { return document.getElementById(id); };
  var CFG = (ctx.data && ctx.data.cfg) || {}, P = CFG.PROJECT || {}, PUM = CFG.PUM || {}, MAIL = CFG.MAIL || {};
  var files = ctx.files && ctx.files.length ? ctx.files : [{ name: ctx.fileName, wb: ctx.wb }];
  var X = group(files, CFG), G = X.G, kinds = ['md', 'cb'].filter(function (k) { return G[k]; });
  var ym = X.ym, mNum = +ym.slice(5), TXT = {}, TABLES = {}, S = { note: '' };
  var saved = ctx.data.saved || {};

  ctx.tabsEl.innerHTML = tabsHtml([['calc', '정산·검증'], ['pum', '품의 복붙'], ['erp', 'ERP 전표'], ['mail', '회계팀 메일']]);
  var sec = function (k, body) { return '<section class="panel"><div class="ph"><h2>' + KIND[k].name + ' · ' + ym + '</h2></div>' + body + '</section>'; };
  ctx.root.innerHTML =
    '<div class="tabpanel" id="p-calc"><section class="panel dash"><div class="tiles" id="tiles"></div><div class="status" id="statusLine"></div><div class="meta" id="amtMeta"></div>' +
    '<details class="fold vfold" open><summary>원본 · 빅뱅 제외본 대조 및 정산서 숫자 검증</summary><div class="flags" id="flags"></div></details></section>' +
    kinds.map(function (k) { return sec(k, '<p class="small muted" style="margin:0" id="files-' + k + '"></p><div class="tablebox"><table id="line-' + k + '"></table></div>'); }).join('') +
    '<details class="panel fold logic"><summary><b>적용 중인 정산 로직</b> <span class="small muted">YGP_MD_응원봉_ERP업로드_자동화.xlsx · 26.7~9월 실제 AR · 빅뱅 별도 관리(26.9.10 회계팀 메일)</span></summary><ol>' +
    '<li><b>빅뱅 제외</b>: YGP-빅뱅 마스터 계약 날인 전까지 모든 품의·ERP는 빅뱅 제외 기준. 원본(빅뱅 포함)과 빅뱅 제외본을 함께 올리면 "원본 − 빅뱅 행 = 제외본"을 발행내역·아티스트별로 대조합니다. 원본만 있으면 빅뱅 행을 빼서 계산합니다.</li>' +
    '<li><b>금액</b>: 정산서 "① YG ENT 발행내역"의 공급가액·세액·합계(빅뱅 제외). 세액은 아티스트별 반올림 합이라 정확히 10%가 아닐 수 있습니다.</li>' +
    '<li><b>ERP 줄</b>: "② 아티스트별 - 가수" 공급가액 = 대변금액. MD는 매핑표 순서·코드(원본 "BABYMONSTER" = ERP "BABYMONSTER_라미제외" 6인, "BABYMONSTER(라미포함)" = ERP "BABYMONSTER" 7인), 응원봉은 원본 순서·프로젝트 "_"·트레저 → "트레저(10인)". 0원 행 제외.</li>' +
    '<li><b>관리항목1·2</b>: 정산월 1일~말일 (MD 날짜 서식 / 응원봉 8자리 숫자).</li></ol></details></div>' +
    '<div class="tabpanel" id="p-pum" hidden>' + kinds.map(function (k) {
      return sec(k, '<div class="block chk"><header><h3>결재 전 체크 — 결재 라인 · 첨부파일 · 관련문서</h3><button class="btn sm" data-copy="chk-' + k + '">전체 복사</button></header><div class="body"><dl class="kv" id="pChk-' + k + '"></dl></div></div>' +
        '<div class="allcopy"><button class="btn primary" data-copy="all-' + k + '">본문 전체 복사</button></div>' +
        '<div class="block"><header><h3>제목</h3><button class="btn sm" data-copy="title-' + k + '">복사</button></header><div class="body"><div id="pTitle-' + k + '"></div></div></div>' +
        '<div class="block"><header><h3>필수 입력 · 1. 기본 내역 · 2. 세금계산서 발행 정보</h3><button class="btn sm" data-copy="basic-' + k + '">전체 복사</button></header><div class="body"><dl class="kv" id="pBasic-' + k + '"></dl></div></div>' +
        '<div class="block fixed"><header><h3>3. 상세내역 — 1. 로열티 상세 내역 <span class="stamp">지정양식</span></h3><button class="btn sm" data-copy="fixed-' + k + '">복사</button></header><div class="body"><pre class="fmt" id="pFixed-' + k + '"></pre>' +
        '<div class="row" style="justify-content:space-between"><b class="small">판매월 표 (빅뱅 제외 · 로열티 = 공급가액)</b><button class="btn sm" data-copy="art-' + k + '">표 복사</button></div><div class="tablebox"><table id="art-' + k + '"></table></div></div></div>' +
        '<div class="block"><header><h3>3. 상세내역 — 2. 참고사항</h3><button class="btn sm" data-copy="note-' + k + '">복사</button></header><div class="body"><pre class="fmt" id="pNote-' + k + '"></pre></div></div>');
    }).join('') + '</div>' +
    '<div class="tabpanel" id="p-erp" hidden>' + kinds.map(function (k) {
      return sec(k, '<p class="small muted" style="margin:0">품의 승인 후 회계 → 매출 추가 → 기안문서 연동 → 세금계산서 등록(부가세만 생성) → 거래내역 엑셀 업로드에 아래 파일을 올리고 처리 → 결재. 업로드 후 날짜 칸이 에러 나는지 한 번 확인합니다.</p>' +
        '<div class="row"><button class="btn primary" data-erp="' + k + '">AR 업로드 파일 내려받기</button><button class="btn" data-save="' + k + '">이 달 확정 저장</button><span class="small muted" id="erpMsg-' + k + '"></span></div>' +
        '<div class="block"><header><h3>전표 헤더 · 세금계산서</h3><button class="btn sm" data-copy="head-' + k + '">전체 복사</button></header><div class="body"><dl class="kv" id="pHead-' + k + '"></dl></div></div>' +
        '<div class="block"><header><h3>AR 줄 (업로드 파일 내용)</h3><button class="btn sm" data-copy="ar-' + k + '">표 복사</button></header><div class="body"><div class="tablebox"><table id="ar-' + k + '"></table></div></div></div>');
    }).join('') + '</div>' +
    '<div class="tabpanel" id="p-mail" hidden><section class="panel"><div class="ph"><h2>회계팀 빅뱅 매출 공유 메일</h2><p>빅뱅 계약 날인 전까지 매월 빅뱅 매출을 원본 정산서(빅뱅 포함)와 함께 공유합니다 (26.9.10 · 10.13 메일 양식).</p></div>' +
    '<div class="form"><label style="grid-column:1/-1">YGP측 전달 특이사항 (없으면 비워 두세요 — 비우면 문단 생략)<textarea id="inNote" rows="3" style="width:100%;font:inherit;padding:8px;border-radius:10px;border:1.5px solid var(--line)"></textarea></label></div>' +
    '<div class="block"><header><h3>받는 사람 · 참조 · 제목</h3><button class="btn sm" data-copy="mailHead">전체 복사</button></header><div class="body"><dl class="kv" id="mHead"></dl></div></div>' +
    '<div class="block"><header><h3>본문</h3><button class="btn sm primary" data-copy="mailBody">본문 복사</button></header><div class="body"><pre class="fmt" id="mBody"></pre>' +
    '<div id="mTables"></div></div></div></section></div>';

  function renderKind(k) {
    var g = G[k], p = g.base, R = g.R, inv = p.invoice, md = k === 'md';
    var title = '[매출품의] YGP MD 로열티 ' + mNum + '월 정산의 건' + (md ? '' : '_응원봉 디자인권 사용료');
    var pay = inv.date ? YGP.monthEnd(YGP.addMonths(inv.date.slice(0, 7), PUM.payMonths || 3)) : '';
    var lineSum = R.lines.reduce(function (a, l) { return a + l.amt; }, 0);
    $('files-' + k).textContent = '원본(빅뱅 포함): ' + (g.orig ? g.orig.file : '미첨부') + ' · 빅뱅 제외본: ' + (g.excl ? g.excl.file : '미첨부(원본에서 빅뱅 행 차감)') + (g.orig ? ' · 빅뱅 ' + fmt(g.orig.big.supply) + '원 별도 관리' : '');
    TABLES['line-' + k] = { head: ['원본 표기', 'ERP 적요 표기', '공급가액(대변)'].concat(md ? ['프로젝트', '그룹', '아티스트'] : []), numCols: [2], sumLast: true,
      rows: R.lines.map(function (l) { return [l.src, l.label, fmt(l.amt)].concat(md ? [l.proj, l.grp, l.art] : []); }).concat([['합계 (빅뱅 제외)', '', fmt(lineSum)].concat(md ? ['', '', ''] : [])]) };
    $('line-' + k).innerHTML = tableHtml(TABLES['line-' + k]);

    // 품의
    TXT['title-' + k] = title; $('pTitle-' + k).textContent = title;
    var basic = [['비거주자 여부', '거주자'], ['외국환 신고 필요여부', '1.경상거래 해당없음 / 2.자본거래 해당없음'], ['거래처명', (P.vendorName || '') + ' (' + (P.vendor || '') + ')'], ['거래일자', inv.date || '-'],
      ['금액', fmt(inv.supply)], ['VAT', fmt(inv.vat)], ['총금액', fmt(inv.total)], ['프로젝트', '프로젝트공통 (_)'], ['입금예정일 (이내)', pay || '-'],
      ['아티스트(그룹)', '그룹공통 (_)'], ['아티스트(개인)', (P.artistName || '아티스트공통') + ' (' + (P.artist || '') + ')'], ['매출 종류', '로열티'], ['계약서 작성 여부', '작성'], ['통화종류', 'KRW'],
      ['청구방법', '정산서 수령 후'], ['에이전시 유/무', '없음'], ['발행정보', '세금계산서'], ['사업자등록번호', P.bizNo || ''], ['담당자 / 이메일', (P.contact || '') + ' / ' + (P.email || '')],

    ];
    // 결재 전 체크: 26.10 품의 26-0852·0853 기준 — 첨부는 빅뱅 제외본 1개, 전월 품의는 관련문서로 연결
    var chk = [['결재 라인 (위임전결)', approval(inv.supply)], ['첨부파일 1', g.excl ? g.excl.file : '빅뱅 제외본(YGP 전달, 암호 제거본) — 아직 올리지 않음. 원본(빅뱅 포함)은 첨부하지 않습니다'],
      ['관련문서 1', '[매출품의] YGP MD 로열티 ' + (+YGP.addMonths(ym, -1).slice(5)) + '월 정산의 건' + (md ? '' : '_응원봉 디자인권 사용료') + ' (전월 결재 문서)']];
    TXT['chk-' + k] = chk.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n'); $('pChk-' + k).innerHTML = kv(chk, 'chk-' + k);
    TXT['basic-' + k] = basic.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
    $('pBasic-' + k).innerHTML = kv(basic, 'basic-' + k);
    // 1. 로열티 상세 내역 — 지정양식. 정산월 = 판매월 = 정산서 월(26.10 품의: 2026.10 정산서 → 2026-10), 프로젝트 = BIS 프로젝트명(MD)
    var fixed = '1. 로열티 상세 내역\n- 정산월 : ' + ym + '\n- 프로젝트 : ' + ((md ? PUM.mdProject : PUM.cbProject) || 'MD') + '\n- 사업구조 : ' + (PUM.struct || 'RS') + '\n- MG : -\n- 통화 : KRW\n- 환율 : 1';
    $('pFixed-' + k).textContent = fixed;
    var gross = p.gross || {}, noGross = p.artists.filter(function (x) { return gross[YGP.K(x.name)] == null; }).map(function (x) { return x.name; });
    TABLES['art-' + k] = { head: ['판매월', '아티스트', '권역', '판매수량', '총 매출 (vat-)', '로열티 (vat-)'], numCols: [3, 4, 5], sumLast: true,
      rows: p.artists.map(function (x) { var gv = gross[YGP.K(x.name)]; return [ym, x.name, '글로벌', fmt(x.qty), gv == null ? '-' : fmt(gv), fmt(x.supply)]; })
        .concat([['합계', '', '', fmt(R.S.qty), fmt(p.artists.reduce(function (a, x) { return a + (gross[YGP.K(x.name)] || 0); }, 0)), fmt(R.S.supply)]]) };
    if (noGross.length) g.flags.push({ level: 'warn', text: '판매월 표 총 매출: "제품매출" 시트에서 ' + noGross.join(', ') + ' 매출을 찾지 못해 "-"로 두었습니다' });
    var note = '2. 참고사항\n' + ((md ? PUM.mdNote : PUM.cbNote) || '');
    TXT['note-' + k] = note;
    $('art-' + k).innerHTML = tableHtml(TABLES['art-' + k]); $('pNote-' + k).textContent = note;
    TXT['fixed-' + k] = fixed + '\n' + tsv(TABLES['art-' + k]);
    TXT['all-' + k] = title + '\n\n' + TXT['basic-' + k] + '\n\n3. 상세내역\n' + TXT['fixed-' + k] + '\n\n' + note;

    // ERP
    // ERP AR마스터·계산서등록 화면 순서 (26.10 AR26111000001·AR26111000002 화면 기준)
    var head = [['인보이스여부', '아니오'], ['증빙서류', '전자세금계산서'], ['계산서발행구분', '정발행'], ['AR유형', '국내매출'], ['회계일자', inv.date || '-'],
      ['청구처', (P.vendor || '') + ' ' + (P.vendorName || '')], ['결제조건', '입금계좌'], ['입금예정일', pay || '-'], ['사업체 / 사업부', '1000 YG엔터테인먼트 / 1 YG엔터테인먼트'],
      ['프로젝트', '_ (프로젝트공통)'], ['회계부서', (P.dept || '') + ' ' + (P.deptName || '')], ['통화 / 환율', 'KRW / 1'], ['제목', title.replace('[매출품의] ', '')], ['기안문서', '위 매출품의 선택'],
      ['계산서등록 · 부가세일자', inv.date || '-'], ['계산서등록 · 부가세코드', '과세(세금계산서)'], ['계산서등록 · 사업자번호 / 판매처', (P.bizNo || '') + ' / ' + (P.vendor || '')],
      ['계산서등록 · 공급가액 / 세액', fmt(inv.supply) + ' / ' + fmt(inv.vat) + ' (빅뱅 제외)'],
      ['거래내역 엑셀 업로드', R.lines.length + '줄 · 대변 합계 ' + fmt(lineSum) + '원'],
      ['ERP 결재선 (위임전결 2.1)', inv.supply < 2e8 ? '재무관리팀 담당 합의 → 팀 리더 결재 → 회계 담당 검토·승인 (2억원 미만)' : '재무관리팀 담당·리더 합의 → 팀 리더 결재 → 회계 담당 → 회계 리더 승인 (2억원 이상)']];
    TXT['head-' + k] = head.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
    $('pHead-' + k).innerHTML = kv(head, 'head-' + k);
    TABLES['ar-' + k] = { head: ['프로젝트', '그룹', '아티스트', '대변금액', '적요'], numCols: [3], sumLast: true,
      rows: R.lines.map(function (l) { return [md ? l.proj : '_', md ? l.grp : '_', md ? l.art : (P.artist || ''), fmt(l.amt), YGP.memo(p, l)]; }).concat([['합계', '', '', fmt(lineSum), '']]) };
    $('ar-' + k).innerHTML = tableHtml(TABLES['ar-' + k]);
    var blocked = g.flags.concat(R.flags, X.flags).some(function (f) { return f.level === 'bad'; });
    $('erpMsg-' + k).textContent = blocked ? '검증 "조치" 항목을 먼저 확인하세요' : '';
    return { title: title, pay: pay, inv: inv };
  }

  function renderMail() {
    var mk = kinds.filter(function (k) { return G[k].orig; }), miss = kinds.filter(function (k) { return !G[k].orig; });
    var subj = '[공유] 빅뱅관련 데이터 업데이트_YGP MD 로열티 ' + mNum + '월 정산의 건';
    var to = (MAIL.to || []).join(', '), cc = (MAIL.cc || []).join(', ');
    var head = [['받는 사람', to || '(DB YGP_설정 MAIL.to)'], ['참조', cc || '-'], ['제목', subj], ['첨부 (원본 정산서·빅뱅 포함)', mk.map(function (k) { return G[k].orig.file; }).join(', ') || '없음 — 원본을 함께 올려 주세요']];
    TXT.mailHead = head.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
    $('mHead').innerHTML = kv(head, 'mailHead');
    var tbl = function (k) { var b = G[k].orig.big; return { head: ['아티스트', '수량', '공급가액', '부가세', '공급대가'], rows: [['빅뱅', fmt(b.qty), fmt(b.supply), fmt(b.vat), fmt(b.total)]], numCols: [1, 2, 3, 4] }; };
    var order = ['md', 'cb'].filter(function (k) { return mk.indexOf(k) >= 0; });
    var note = S.note.trim();
    var body = (MAIL.greeting || '안녕하세요, 광고 IP 사업팀 김윤수입니다.') + '\n\n' +
      '금일 YGP로부터 빅뱅 관련매출 전달받아\n' + mNum + '월 정산분 빅뱅 매출 아래와 같이 공유드리며, 원본정산서(빅뱅 포함) 파일 첨부드립니다.\n\n' +
      (note ? '[YGP측 전달_특이사항]\n' + note + '\n\n' : '') +
      order.map(function (k) { TABLES['mail-' + k] = tbl(k); return KIND[k].mail + '_' + mNum + '월 정산분\n' + tsv(TABLES['mail-' + k]); }).join('\n\n') +
      '\n\n관련하여 문의사항은 언제든 말씀 부탁드립니다.\n\n감사합니다.\n\n' + (MAIL.sign || '김윤수 드림');
    TXT.mailBody = order.length ? body : '';
    $('mBody').textContent = order.length ? body : '원본 정산서(빅뱅 포함)를 올리면 빅뱅 매출 공유 메일이 만들어집니다. 지금은 빅뱅 제외본만 있습니다.';
    $('mTables').innerHTML = order.map(function (k) { return '<div class="row" style="justify-content:space-between"><b class="small">' + esc(KIND[k].mail + '_' + mNum + '월 정산분') + '</b><button class="btn sm" data-copy="mail-' + k + '">표 복사</button></div><div class="tablebox"><table>' + tableHtml(TABLES['mail-' + k]) + '</table></div>'; }).join('') +
      (miss.length && order.length ? '<p class="small" style="color:var(--bad)">' + esc(miss.map(function (k) { return KIND[k].name; }).join('·')) + ' 원본(빅뱅 포함)이 없어 메일에서 빠졌습니다.</p>' : '');
  }

  function render() {
    var flags = X.flags.slice(), info = {};
    kinds.forEach(function (k) {
      info[k] = renderKind(k);
      flags = flags.concat(G[k].flags, G[k].R.flags.map(function (f) { return { level: f.level, text: KIND[k].short + ' · ' + f.text }; }));
      var sv = saved[(k === 'md' ? 'MD ' : '응원봉 ') + ym];
      if (sv) flags.push({ level: 'ok', text: KIND[k].short + ' ' + ym + '은 ' + (sv.savedAt || '') + ' 확정 저장된 기록이 있습니다 (' + fmt(sv.supply) + '원)' });
    });
    if (!G.md || !G.cb) flags.push({ level: 'warn', text: (G.md ? '응원봉' : 'MD 로열티') + ' 정산서가 없습니다 — 같은 달 파일을 함께 올리면 품의·ERP·메일이 한 화면에 모두 만들어집니다' });
    var tiles = [['정산월', ym + ' (빅뱅 제외)']];
    kinds.forEach(function (k) { tiles.push([KIND[k].short + ' 공급가액', fmt(info[k].inv.supply) + '원']); tiles.push([KIND[k].short + ' 총금액', fmt(info[k].inv.total) + '원']); });
    kinds.forEach(function (k) { if (G[k].orig) tiles.push([KIND[k].short + ' 빅뱅 (별도)', fmt(G[k].orig.big.supply) + '원']); });
    $('tiles').innerHTML = tiles.map(function (t) { return '<div class="tile"><span>' + esc(t[0]) + '</span><b>' + esc(t[1]) + '</b></div>'; }).join('');
    var st = statusHtml(flags); $('statusLine').className = st.cls; $('statusLine').innerHTML = st.html;
    $('amtMeta').textContent = files.length + '개 파일 · ' + files.map(function (f) { return f.name; }).join(' · ');
    $('flags').innerHTML = flagsHtml(flags);
    renderMail();
    return info;
  }

  // 복사·탭
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
  var info = render();
  ctx.root.addEventListener('click', async function (e) {
    var b = e.target.closest('[data-copy]');
    if (b) { var k = b.getAttribute('data-copy'); if (TABLES[k]) copyTable(TABLES[k]); else if (TXT[k] != null) copyText(TXT[k], '내용을'); return; }
    var v = e.target.closest('[data-copyval]');
    if (v) { var q = v.getAttribute('data-copyval').split(':'), line = (TXT[q[0]] || '').split('\n')[+q[1]] || ''; copyText(line.split('\t').slice(1).join('\t'), '값을'); return; }
    var d = e.target.closest('[data-erp]');
    if (d) {
      var kk = d.getAttribute('data-erp'), out = YGP.buildErpUpload(G[kk].base, G[kk].R, CFG);
      ctx.saveFile((kk === 'md' ? '[md]' : '[응원]') + 'AR거래내역엑셀업로드_YGPMD' + (kk === 'md' ? '' : '_응원봉') + '_' + ym.replace('-', '.') + '(빅뱅제외).xlsx', XLSX.write(out.wb, { type: 'array', bookType: 'xlsx' }));
      ctx.toast(KIND[kk].short + ' AR 업로드 파일을 내려받았습니다 (' + out.rows.length + '줄)'); return;
    }
    var sv = e.target.closest('[data-save]');
    if (sv) {
      var ks = sv.getAttribute('data-save'), g = G[ks], iv = info[ks].inv, key = (ks === 'md' ? 'MD ' : '응원봉 ') + ym;
      try {
        $('erpMsg-' + ks).textContent = '저장 중…';
        await ctx.post('saveQuarter', { record: { quarter: key, supply: iv.supply, vat: iv.vat, gross: iv.total, settleMonth: ym, txDate: iv.date, payDate: info[ks].pay,
          file: [g.orig ? g.orig.file : '', g.excl ? g.excl.file : ''].filter(Boolean).join(' + '), roundNote: g.orig ? '빅뱅 별도 ' + g.orig.big.supply : '',
          lines: g.R.lines.map(function (l) { return { label: l.label, amt: l.amt }; }), royaltyTable: g.orig ? [g.orig.big] : [] } });
        $('erpMsg-' + ks).textContent = ''; ctx.toast(key + ' 확정 저장했습니다');
      } catch (er) { $('erpMsg-' + ks).textContent = '저장하지 못했습니다 (' + er.message + ')'; }
    }
  });
  $('inNote').addEventListener('input', function () { S.note = this.value; renderMail(); });
  var TABS = ['calc', 'pum', 'erp', 'mail'];
  function showTab(t) {
    if (TABS.indexOf(t) < 0) t = TABS[0];
    TABS.forEach(function (k) { $('p-' + k).hidden = k !== t; $('t-' + k).setAttribute('aria-selected', k === t ? 'true' : 'false'); });
    try { localStorage.setItem('ygp.tab', t); } catch (e) { }
  }
  ctx.tabsEl.querySelectorAll('.tab').forEach(function (b) { b.addEventListener('click', function () { showTab(b.getAttribute('data-tab')); window.scrollTo(0, 0); }); });
  var st = null; try { st = localStorage.getItem('ygp.tab'); } catch (e) { }
  showTab(st || 'calc');
  return { quarter: ym + ' YGP', note: files.length + '개 파일 · ' + kinds.map(function (k) { return KIND[k].short + (G[k].orig ? ' 원본' : '') + (G[k].excl ? ' 제외본' : ''); }).join(' / ') };
}

DESK.register({ id: 'ygp', server: 'ygp', multi: true, name: 'YGP · MD 로열티 · 응원봉', desc: '와이지플러스 MD 로열티 · 응원봉 디자인권 사용료 · 월 매출 (빅뱅 제외)',
  fileHint: 'Ent 정산(.xlsb) · 응원봉 디자인권 정산(.xlsx) — 원본(빅뱅 포함)·빅뱅 제외본 함께, 암호 제거본', outputs: ['원본·제외본 대조 검증', '매출품의 2건', 'ERP AR 업로드 파일 2개', '회계팀 빅뱅 매출 메일'],
  detect: function (wb) { try { YGP.parse(wb); return true; } catch (e) { return false; } },
  why: function (wb) { if (!wb.Sheets['항목별정산내역_요약']) return ''; try { YGP.parse(wb); return ''; } catch (e) { return e.message; } },
  mount: mount });
})();
