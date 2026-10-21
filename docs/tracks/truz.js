/* 정산 모듈: TRUZ (IPX × 트레저 캐릭터 상품 로열티, 분기)
 * 판별(detect) → 셸이 이 정산의 비공개 설정을 서버에서 받아 mount 호출 → 정산·품의·ERP 탭 렌더.
 * 회사 데이터(매핑·이력·코드·거래처·검증 기준)는 이 파일에 없다. */
(function () {
/* TRUZ 정산 엔진 — YG검토본 로직 역산 (실제 품의·ERP 금액과 1원 단위 일치 검증) */
var TRUZ = (function () {
  var N = function (s) { return String(s == null ? '' : s).replace(/\s+/g, ''); };
  var MON = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };

  // IPX "by Biz Cat." → 정산 카테고리 (품의 표기 / ERP 적요 표기)
  var CATS = [
    { key: 'OEM', label: 'OEM', erp: null },
    { key: 'B2B', label: 'B2B', erp: 'B2B' },
    { key: 'STICKER', label: 'LINE STICKER', erp: 'STICKER' },
    { key: 'LICENSE', label: '라이선스', erp: '라이선시제품' },
    { key: 'YT', label: '유튜브 정규', erp: '유튜브' },
    { key: 'YT_ETC', label: '정규 외', erp: '정규에피소드외 유튜브 수익' }
  ];
  // 21.2Q~ 누적 피벗(정산 피벗 참고.png) 색 구분 기준: Partnership·STICKER(고요율), GAME, LICENSE(Product)는 라이선스(빨강)
  function catOf(biz, pct) {
    var b = N(biz).toUpperCase();
    if (b === 'OFFLINE' || b === 'ONLINE') return 'OEM';
    if (b === 'B2B') return 'B2B';
    if (b === 'PARTNERSHIP' || b === 'STICKER') return pct >= 0.3 ? 'LICENSE' : 'STICKER';
    if (b === 'LICENSE' || b.indexOf('LICENSE(') === 0 || b === 'GAME') return 'LICENSE';
    if (b === 'CONTENTS') return pct >= 0.5 ? 'YT' : 'YT_ETC';
    return null;
  }

  // 멤버 → ERP 프로젝트·아티스트 코드: 비공개 설정(TRUZ 정산 DB "설정" 시트 MEMBERS)에서 로그인 후 주입
  var MEMBERS = [], MEMBER_NAMES = [];
  function setMembers(list) {
    MEMBERS.length = 0; MEMBER_NAMES.length = 0;
    (list || []).forEach(function (x) { MEMBERS.push(x); MEMBER_NAMES.push(x.m); });
  }
  // 품의 표기 라벨(요율 포함)과 ERP 공통 코드도 비공개 설정에서 주입
  var ERP = { dept: null, account: null, vendorCode: null, commonArtist: null };
  function setConfig(cfg) {
    cfg = cfg || {};
    if (cfg.MEMBERS) setMembers(cfg.MEMBERS);
    if (cfg.CAT_LABELS) CATS.forEach(function (c) { if (cfg.CAT_LABELS[c.key]) c.label = cfg.CAT_LABELS[c.key]; });
    if (cfg.PROJECT) { ERP.dept = cfg.PROJECT.dept; ERP.account = cfg.PROJECT.account; ERP.commonArtist = Number((String(cfg.PROJECT.artist || '').match(/\((\d+)\)/) || [])[1]) || null; }
    if (cfg.VENDOR) ERP.vendorCode = cfg.VENDOR.code;
  }

  function rowsOf(ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }); }
  function num(v) { var n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, '')); return isFinite(n) ? n : 0; }

  // 1) IPX 원본 파일 해석
  function parse(wb) {
    var dataName = wb.SheetNames.filter(function (s) { return /^Data/.test(s); })[0];
    if (!dataName) throw new Error('"Data (26Qn)" 시트를 찾지 못했습니다. IPX 원본 파일(TRUZ_Royalty_Summary_…)인지 확인해 주세요.');
    var rows = rowsOf(wb.Sheets[dataName]);
    var hi = rows.findIndex(function (r) { return r.some(function (c) { return N(c) === 'Year'; }); });
    if (hi < 0) throw new Error('Data 시트에서 "Year" 머리글을 찾지 못했습니다.');
    var h = rows[hi].map(N);
    var col = function (name) { var i = h.indexOf(name); if (i < 0) throw new Error('Data 시트에 "' + name + '" 열이 없습니다.'); return i; };
    var iY = col('Year'), iM = col('Month'), iB = col('byBizCat.'), iCh = col('byChannel'), iP = col('Royalty%'),
      iS = col('SalesTotalAmt.(SubjecttoRoyalty)'), iR = col('RoyaltyTotalAmt.');
    var iC = h.indexOf('Currency'), iX = h.indexOf('ExchangeRate');
    var data = [], fx = {};
    rows.slice(hi + 1).forEach(function (r) {
      if (r[iY] == null || r[iR] == null || r[iR] === '') return;
      var mo = MON[String(r[iM]).slice(0, 3)];
      if (!mo) return;
      var ym = parseInt(r[iY], 10) + '-' + pad(mo);
      data.push({ ym: ym, biz: N(r[iB]), channel: String(r[iCh] || '').trim(), pct: num(r[iP]), sales: num(r[iS]), roy: num(r[iR]) });
      if (iC >= 0 && iX >= 0 && r[iC] && r[iX] != null) {
        var k = ym + '|' + String(r[iC]).trim();
        fx[k] = fx[k] || {}; fx[k][r[iX]] = (fx[k][r[iX]] || 0) + 1;
      }
    });
    if (!data.length) throw new Error('Data 시트에 정산 행이 없습니다.');
    var months = Array.from(new Set(data.map(function (d) { return d.ym; }))).sort();
    var last = months[months.length - 1], y = +last.slice(0, 4), q = Math.ceil(+last.slice(5) / 3);

    var detName = wb.SheetNames.filter(function (s) { return ['Detail', '품목 상세', '자재'].indexOf(s.trim()) >= 0; })[0];
    if (!detName) detName = wb.SheetNames.filter(function (s) { var r0 = rowsOf(wb.Sheets[s])[0] || []; var hh = r0.map(N); return hh.indexOf('제품명') >= 0 && hh.indexOf('수수료') >= 0; })[0];
    if (!detName) throw new Error('품목 상세(Detail) 시트를 찾지 못했습니다.');
    var dr = rowsOf(wb.Sheets[detName]); var dh = dr[0].map(N);
    var dc = function (alts) { for (var i = 0; i < alts.length; i++) { var j = dh.indexOf(alts[i]); if (j >= 0) return j; } throw new Error('Detail 시트에 "' + alts[0] + '" 열이 없습니다.'); };
    var jSet = dc(['정산월']), jCty = dc(['국가']), jCh = dc(['유통경로', '채널']), jChN = dc(['채널명']), jMk = dc(['제작방식']),
      jNm = dc(['제품명']), jQty = dc(['주문수량']), jCur = dc(['거래통화']), jFee = dc(['수수료']);
    var detail = [];
    dr.slice(1).forEach(function (r) {
      if (r[jSet] == null || r[jSet] === '') return;
      detail.push({ ym: String(r[jSet]).trim(), country: String(r[jCty] || '').trim(), channel: String(r[jCh] || '').trim(),
        channelName: String(r[jChN] || '').trim(), make: String(r[jMk] || '').trim(), name: String(r[jNm] || '').trim(),
        qty: num(r[jQty]), cur: String(r[jCur] || '').trim().toUpperCase(), fee: num(r[jFee]), raw: r });
    });
    // 환율 기본값: Data 시트 Exchange Rate (월·통화별 최빈값)
    var rates = {};
    Object.keys(fx).forEach(function (k) {
      var best = null, bc = -1; Object.keys(fx[k]).forEach(function (v) { if (fx[k][v] > bc) { bc = fx[k][v]; best = +v; } });
      rates[k] = best;
    });
    detail.forEach(function (d) { var k = d.ym + '|' + d.cur; if (d.cur === 'KRW') rates[k] = 1; else if (!(k in rates)) rates[k] = null; });
    return { wb: wb, quarter: { y: y, q: q, label: y + '.' + q + 'Q', key: y + 'Q' + q, months: months }, data: data, detail: detail,
      detailHeader: dr[0], rates: rates, dataSheet: dataName, detailSheet: detName };
  }

  // 2) 계산
  function compute(p, rates, map, opt) {
    opt = opt || {};
    var flags = [];
    var cat = {}, catMonth = {}, unknownBiz = {}, total = 0;
    CATS.forEach(function (c) { cat[c.key] = 0; catMonth[c.key] = {}; });
    p.data.forEach(function (d) {
      total += d.roy;
      var k = catOf(d.biz, d.pct);
      if (!k) { unknownBiz[d.biz] = (unknownBiz[d.biz] || 0) + d.roy; return; }
      cat[k] += d.roy; catMonth[k][d.ym] = (catMonth[k][d.ym] || 0) + d.roy;
    });
    Object.keys(unknownBiz).forEach(function (b) { flags.push({ level: 'bad', text: '처음 보는 사업 구분 "' + b + '" (' + fmt(unknownBiz[b]) + '원) — 카테고리 분류 규칙 확인 필요' }); });

    var mem = {}, memCty = {}, unmapped = {}, missingRate = {}, countries = {}, g12 = 0;
    MEMBER_NAMES.forEach(function (m) { mem[m] = 0; memCty[m] = {}; });
    var rows = p.detail.map(function (d) {
      var isOem = d.channel === '온라인' || d.channel === '오프라인';
      var rk = d.ym + '|' + d.cur, rate = d.cur === 'KRW' ? 1 : rates[rk];
      var m = map[N(d.name)] || null;
      var krw = isOem ? d.fee * (rate || 0) : 0;
      if (isOem) {
        if (rate == null || rate === '' || isNaN(rate)) missingRate[rk] = true;
        var bucket = m;
        if (m === '트레저 12인') { bucket = '트레저 10인'; g12 += krw; }
        if (!m) { unmapped[d.name] = (unmapped[d.name] || 0) + krw; bucket = '(미매핑)'; }
        if (!(bucket in mem)) { mem[bucket] = 0; memCty[bucket] = {}; }
        mem[bucket] += krw; memCty[bucket][d.country] = (memCty[bucket][d.country] || 0) + krw; countries[d.country] = 1;
      }
      return { d: d, rate: d.cur === 'KRW' ? 1 : rate, krw: krw, member: m };
    });
    Object.keys(missingRate).forEach(function (k) { flags.push({ level: 'bad', text: k.replace('|', ' ') + ' 환율이 비어 있습니다 — 환율 표에 입력해 주세요' }); });
    var unN = Object.keys(unmapped).length;
    if (unN) flags.push({ level: 'bad', text: '멤버 매핑이 없는 제품 ' + unN + '개 (' + fmt(Object.values(unmapped).reduce(function (a, b) { return a + b; }, 0)) + '원) — 아래 "새 제품 매핑"에서 지정' });
    if (g12) flags.push({ level: 'warn', text: '"트레저 12인" 제품 ' + fmt(g12) + '원을 트레저(10인) 프로젝트로 합산했습니다 — 회계팀 처리 기준 확인 권장' });

    var memSum = Object.keys(mem).reduce(function (a, k) { return a + mem[k]; }, 0);
    var oemDiff = memSum - cat.OEM;
    flags.push(Math.abs(oemDiff) < 1
      ? { level: 'ok', text: 'OEM 합계 대사 일치 — Data 시트 ' + fmt(cat.OEM) + '원 = 멤버별 합계 ' + fmt(memSum) + '원 (차액 ' + oemDiff.toFixed(2) + '원, 0원 처리)' }
      : { level: Math.abs(oemDiff) < 1000 ? 'warn' : 'bad', text: 'OEM 합계 차액 ' + fmt(oemDiff) + '원 — Data 시트 ' + fmt(cat.OEM) + '원 vs 멤버별 ' + fmt(memSum) + '원. 환율·매핑 확인 필요' });

    // 3) 원 단위 처리 (인수인계 녹음 · 26.1Q/2Q 품의·ERP 대조로 확정)
    //    공급가액 = 원본 총액 반올림(고정) / 카테고리 = 월별 반올림 후 합산 / 멤버 = 멤버별 반올림
    //    1원 차이는 수정하지 않고 체크만 표시 → 회계팀이 조정, 품의에 주석으로 명시
    var supply = Math.round(total);
    var catR = {}, catMonthR = {}, catDiff = {};
    CATS.forEach(function (c) {
      catMonthR[c.key] = {}; var s = 0;
      Object.keys(catMonth[c.key]).forEach(function (m) { var v = Math.round(catMonth[c.key][m]); catMonthR[c.key][m] = v; s += v; });
      catR[c.key] = s; catDiff[c.key] = Math.round(cat[c.key]) - s;
    });
    var tableSum = CATS.reduce(function (a, c) { return a + catR[c.key]; }, 0);
    var lines = [];
    CATS.forEach(function (c) { if (c.erp) lines.push({ kind: 'cat', key: c.key, label: c.erp, member: '트레저 10인', raw: cat[c.key], amt: catR[c.key] }); });
    MEMBERS.forEach(function (x) { lines.push({ kind: 'oem', key: 'OEM:' + x.m, label: 'OEM제품 ' + x.erp, member: x.m, raw: mem[x.m] || 0, amt: Math.round(mem[x.m] || 0) }); });
    if (mem['(미매핑)']) lines.push({ kind: 'oem', key: 'OEM:(미매핑)', label: 'OEM제품 (미매핑)', member: '(미매핑)', raw: mem['(미매핑)'], amt: Math.round(mem['(미매핑)']) });
    var lineSum = lines.reduce(function (a, l) { return a + l.amt; }, 0);
    var memRSum = lines.filter(function (l) { return l.kind === 'oem'; }).reduce(function (a, l) { return a + l.amt; }, 0);
    var tableDiff = supply - tableSum, erpDiff = supply - lineSum;
    var diffCats = CATS.filter(function (c) { return catDiff[c.key]; }).map(function (c) { return c.label.replace(/[()]/g, '').replace(/\s+/g, ' ').trim(); });
    var roundNote = tableDiff ? '※ ' + (diffCats.length ? diffCats.join(', ') + ' 매출 관련, ' : '') + '소수점 반올림 건으로 인해 ' + Math.abs(tableDiff) + '원 차이 발생.' : '';
    if (!tableDiff && !erpDiff) flags.push({ level: 'ok', text: '원 단위 반올림 일치 — 공급가액 ' + fmt(supply) + '원 = 카테고리 표 합계 = ERP 줄 합계' });
    if (tableDiff) flags.push({ level: 'warn', text: '체크: 카테고리 표 합계 ' + fmt(tableSum) + '원 vs 공급가액 ' + fmt(supply) + '원 (' + (tableDiff > 0 ? '+' : '') + tableDiff + '원, ' + (diffCats.join('·') || '반올림') + '). 수정하지 않고 품의에 주석으로 명시 — 회계팀 조정' });
    if (erpDiff) flags.push({ level: 'warn', text: '체크: ERP 줄 합계 ' + fmt(lineSum) + '원 vs 공급가액 ' + fmt(supply) + '원 (' + (erpDiff > 0 ? '+' : '') + erpDiff + '원' + (memRSum !== catR.OEM ? ', 멤버 반올림 합계 ' + fmt(memRSum) + '원 vs OEM ' + fmt(catR.OEM) + '원' : '') + '). 금액란은 공급가액 그대로, 회계팀 조정' });
    // 부가세: 원 미만 절사 (26.1Q 품의·BPTG 26.2Q 전표로 확인)
    var vat = Math.floor(supply * 0.1);

    // 품의 "1. 로열티 상세 내역" 표: TREASURE = 총액 − 멤버 개인 OEM
    var memberRows = MEMBERS.filter(function (x) { return x.m !== '트레저 10인'; }).map(function (x) {
      return { artist: x.m, amt: lines.filter(function (l) { return l.key === 'OEM:' + x.m; })[0].amt };
    });
    var treasure = supply - memberRows.reduce(function (a, r) { return a + r.amt; }, 0);

    return { total: total, supply: supply, vat: vat, gross: supply + vat, cat: cat, catMonth: catMonth, catMonthR: catMonthR, catR: catR,
      tableSum: tableSum, tableDiff: tableDiff, lineSum: lineSum, erpDiff: erpDiff, roundNote: roundNote, mem: mem, memCty: memCty,
      countries: Object.keys(countries).sort(), lines: lines, flags: flags, unmapped: unmapped, rows: rows,
      royaltyTable: [{ artist: 'TREASURE', amt: treasure }].concat(memberRows), oemDiff: oemDiff };
  }

  function fmt(n) { return Math.round(n).toLocaleString('ko-KR'); }

  // 위임전결기준표(261006) — 사업부문 매출 품의
  function approval(supply) {
    if (supply < 5e6) return { band: '500만원 미만 (합의종결)', line: '기안 → 팀 리더(최종) · 합의: 재무관리팀 담당, 회계팀 담당' };
    if (supply < 5e7) return { band: '500만원 이상 5,000만원 미만 (합의종결)', line: '기안 → 팀 리더 → 그룹 리더(최종) · 합의: 재무관리팀 담당, 회계팀 담당 → 회계팀 리더' };
    if (supply < 3e8) return { band: '5,000만원 이상 3억원 미만', line: '기안 → 팀 리더 → 그룹 리더 · 합의: 재무관리팀·회계팀 담당 → 회계팀 리더 → CFO → 부문 대표(최종)' };
    return { band: '3억원 이상', line: '기안 → 팀 리더 → 그룹 리더 · 합의: 재무관리팀·회계팀 담당 → 회계팀 리더 → CFO → 부문 대표 → 대표이사(최종)' };
  }

  // YG검토본 (26.1Q·2Q 검토본 구조): IPX 원본 시트 유지 + Detail 자리에 자재 + PIVOT / PIVOT 멤버 / 멤버 (+ ERP입력·품의입력)
  function buildReview(p, r, opt) {
    opt = opt || {};
    var src = p.wb, wb = XLSX.utils.book_new(), A = XLSX.utils.aoa_to_sheet, add = XLSX.utils.sheet_add_aoa;
    var qm = p.quarter.months, monLab = qm.map(function (m) { return (+m.slice(5)) + '월'; });
    // 자재
    var det = [p.detailHeader.concat(['환율', 'KRW', '멤버'])];
    r.rows.forEach(function (x) {
      var oem = x.d.channel === '온라인' || x.d.channel === '오프라인';
      det.push(x.d.raw.concat([x.rate == null ? '' : x.rate, x.krw, oem ? (x.member || '#N/A') : (x.member || '')]));
    });
    src.SheetNames.forEach(function (s) {
      if (s === p.detailSheet) XLSX.utils.book_append_sheet(wb, A(det), '자재');
      else XLSX.utils.book_append_sheet(wb, src.Sheets[s], s);
    });
    // PIVOT — A열: 전체 데이터 피벗(사업구분 × 요율 × 월), H열: 품의용 카테고리 표 + 체크 메모, 누적
    var agg = {}, keys = [];
    p.data.forEach(function (d) { var k = d.biz + '|' + d.pct; if (!agg[k]) { agg[k] = {}; keys.push(k); } agg[k][d.ym] = (agg[k][d.ym] || 0) + d.roy; });
    keys.sort();
    var pa = [['합계 : Royalty Total Amt.'], ['사업구분', 'Royalty%'].concat(monLab, ['총합계'])];
    keys.forEach(function (k) { var v = qm.map(function (m) { return agg[k][m] || 0; }); pa.push([k.split('|')[0], +k.split('|')[1]].concat(v, [v.reduce(function (a, b) { return a + b; }, 0)])); });
    var colT = qm.map(function (m) { return p.data.filter(function (d) { return d.ym === m; }).reduce(function (a, d) { return a + d.roy; }, 0); });
    pa.push(['총합계', ''].concat(colT, [r.total]));
    var ws = A(pa);
    var ct = [['구분'].concat(monLab, ['TTL', '체크 (수정 안 함 · 회계팀 조정)'])];
    CATS.forEach(function (c) {
      var v = qm.map(function (m) { return r.catMonthR[c.key][m] || 0; }), d = Math.round(r.cat[c.key]) - r.catR[c.key];
      ct.push([c.label].concat(v, [r.catR[c.key], d ? '분기 반올림 ' + Math.round(r.cat[c.key]) + ' (' + (d > 0 ? '+' : '') + d + '원)' : '']));
    });
    var mt = qm.map(function (m) { return CATS.reduce(function (a, c) { return a + (r.catMonthR[c.key][m] || 0); }, 0); });
    ct.push(['TTL'].concat(mt, [r.tableSum, r.tableDiff ? '공급가액 ' + r.supply + ' / 차이 ' + (r.tableDiff > 0 ? '+' : '') + r.tableDiff + '원' : '합은 맞음']));
    if (r.roundNote) ct.push([r.roundNote]);
    ct.push([]); ct.push(['구분', '로열티']);
    (opt.hist || []).forEach(function (h) { ct.push([h[0], h[1]]); });
    ct.push(['총합계', (opt.hist || []).reduce(function (a, h) { return a + h[1]; }, 0)]);
    add(ws, ct, { origin: 'H1' });
    XLSX.utils.book_append_sheet(wb, ws, 'PIVOT');
    // PIVOT 멤버 — 멤버 × 국가 (KRW)
    var ctys = r.countries, pm = [['합계 : KRW'], ['멤버'].concat(ctys, ['총합계'])];
    Object.keys(r.memCty).filter(function (m) { return r.mem[m]; }).forEach(function (m) { pm.push([m].concat(ctys.map(function (c) { return r.memCty[m][c] || 0; }), [r.mem[m]])); });
    pm.push(['총합계'].concat(ctys.map(function (c) { return Object.keys(r.memCty).reduce(function (a, m) { return a + (r.memCty[m][c] || 0); }, 0); }), [Object.keys(r.mem).reduce(function (a, m) { return a + r.mem[m]; }, 0)]));
    XLSX.utils.book_append_sheet(wb, A(pm), 'PIVOT 멤버');
    // 멤버 — 이번 분기 OEM 제품 매핑
    var seen = {}, mm = [['자재명', '멤버']];
    r.rows.forEach(function (x) { if ((x.d.channel === '온라인' || x.d.channel === '오프라인') && !seen[x.d.name]) { seen[x.d.name] = 1; mm.push([x.d.name, x.member || '#N/A']); } });
    XLSX.utils.book_append_sheet(wb, A(mm), '멤버');
    if (opt.erp) XLSX.utils.book_append_sheet(wb, A(opt.erp), 'ERP입력');
    if (opt.pum) XLSX.utils.book_append_sheet(wb, A(opt.pum), '품의입력');
    return wb;
  }


  // ERP "AR 거래내역 일괄생성 업로드" 양식 (YGP MD 업로드 파일과 동일한 시트·열 구성)
  function buildErpUpload(p, r, qi) {
    var serial = function (ymd) { var t = ymd.split('-'); return Math.round((Date.UTC(+t[0], +t[1] - 1, +t[2]) - Date.UTC(1899, 11, 30)) / 86400000); };
    var head = ['프로젝트', '회계부서', '계정코드', '예산코드', '배부코드', '통화', '환율', '차변금액', '대변금액', '거래처', '반제전표', '적요', '세금일자', '부가세코드', '공급가액', '부가세금액', '신고사업장', '세금계산서번호', '아티스트그룹', '아티스트', '관리항목1', '관리항목2', '관리항목3', '관리항목4', '관리항목5', '관리항목6', '관리항목7', '관리항목8'];
    var aoa = [['AR 거래내역 일괄생성 업로드 양식 // ERP자동전표_IP'], [], head];
    var info = {}; MEMBERS.forEach(function (x) { info[x.m] = x; });
    var pre = qi.y + '.' + qi.q + 'Q 트레저 TRUZ 콜라보 제품 판매 로열티 : ';
    r.lines.forEach(function (l) {
      var mi = info[l.member];
      if (!mi) return; // (미매핑) 줄은 업로드 대상에서 제외 — 검증 단계에서 매핑 먼저
      var grp = l.kind === 'cat' ? info['트레저 10인'] : mi;
      aoa.push([+grp.code, ERP.dept, String(ERP.account), null, null, 'KRW', 1, null, l.amt, ERP.vendorCode, null, pre + l.label, null, null, null, null, null, null,
        grp.grp, l.kind === 'cat' ? ERP.commonArtist : grp.art, serial(qi.start), serial(qi.end), 'KR', null, null, null, null, null]);
    });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'AR거래내역엑셀업로드 List');
    return { wb: wb, rows: aoa.slice(3) };
  }

  return { setConfig: setConfig, setMembers: setMembers, buildErpUpload: buildErpUpload, buildReview: buildReview, N: N, CATS: CATS, MEMBERS: MEMBERS, MEMBER_NAMES: MEMBER_NAMES, parse: parse, compute: compute, fmt: fmt, approval: approval, pad: pad };
})();


var TABS_HTML = `  <div class="tabs-in" role="tablist" aria-label="작업 단계">
    <button class="tab" role="tab" id="t-calc" data-tab="calc" aria-selected="true"><span>1</span>정산</button>
    <button class="tab" role="tab" id="t-pum" data-tab="pum" aria-selected="false"><span>2</span>품의 복붙</button>
    <button class="tab" role="tab" id="t-erp" data-tab="erp" aria-selected="false"><span>3</span>ERP 전표</button>
  </div>`;
var VIEW = `    <!-- 탭 1: 업로드 → 대시보드 → YG검토본 -->
    <div class="tabpanel" id="p-calc" role="tabpanel" aria-labelledby="t-calc">
      <section class="panel dash" id="s-amt">
        <div class="tiles">
          <div class="tile main"><span>공급가액 (VAT 제외)</span><b id="tSupply">0</b><em>품의 "금액" · 세금계산서 공급가액</em></div>
          <div class="tile"><span>부가세액 (10%, 원 미만 절사)</span><b id="tVat">0</b><em>품의 "VAT"</em></div>
          <div class="tile"><span>총액</span><b id="tGross">0</b><em>품의 "총금액" · 외상매출금</em></div>
        </div>
        <button class="status" id="statusLine" type="button"></button>
        <div class="meta" id="amtMeta"></div>
        <div class="row dl">
          <button class="btn primary big" id="btnXlsx">YG검토본 엑셀 내려받기</button>
          <button class="btn" id="btnSave">이 분기 확정 저장</button>
          <span class="small muted" id="outMsg"></span>
        </div>
        <p class="small muted" style="margin:0">YG검토본 = IPX 원본 시트 + 자재(환율·KRW·멤버 열 추가) + PIVOT·PIVOT 멤버·멤버 시트. 품의 첨부로 그대로 씁니다.</p>
      </section>

      <details class="panel fold" id="dCheck">
        <summary><b>검증 상세</b> <span class="small muted" id="checkSum"></span></summary>
        <div class="flags" id="flags"></div>
        <div id="mapFix" hidden></div>
      </details>
      <details class="panel fold" id="dFx">
        <summary><b>환율</b> <span class="small muted">IPX 원본 Data 시트 기준 자동 기입 · 멤버별 OEM 배분에만 사용 · 수정 시 즉시 재계산</span></summary>
        <div class="tablebox"><table id="fxTable"></table></div>
      </details>
      <details class="panel fold logic">
        <summary><b>적용 중인 정산 로직</b> <span class="small muted">26.1Q·26.2Q 품의·ERP·YG검토본, 인수인계 녹음 기준</span></summary>
        <ol>
          <li><b>총 로열티(공급가액)</b>: Data 시트 Royalty Total Amt. 전체 합계를 원 단위로 반올림합니다.</li>
          <li><b>카테고리 분류</b>: Offline·Online → OEM / B2B / Partnership·STICKER(요율 기준) → LINE STICKER 또는 라이선스 / License·GAME → 라이선스 / CONTENTS(요율 기준) → 유튜브 정규 또는 정규 외. 요율은 서버 설정에서 관리합니다.</li>
          <li><b>멤버 배분</b>: Detail 시트에서 유통경로가 온라인·오프라인인 행만 대상(B2B 제외)으로, 수수료 × 월별 환율 = KRW를 계산한 뒤 제품명 → 멤버 매핑(1,406개 제품)으로 합산합니다.</li>
          <li><b>대사</b>: 멤버별 합계와 Data 시트 OEM 합계의 차액은 0원으로 처리합니다(IPX 확인 방식).</li>
          <li><b>원 단위 처리</b>: 공급가액은 원본 총액 반올림으로 고정합니다. 카테고리 금액은 <b>월별 반올림 후 합산</b>, 멤버 OEM은 멤버별 반올림입니다.</li>
          <li><b>1원 차이</b>: 표·ERP 합계가 공급가액과 1~2원 다르면 <b>수정하지 않고 체크만</b> 해 두고, 품의에 "※ ○○ 매출 관련, 소수점 반올림 건으로 인해 N원 차이 발생." 주석을 답니다. 조정은 회계팀이 합니다(인수인계 녹음 1:21:26, 26.1Q 품의).</li>
          <li><b>부가세</b>: 공급가액 × 10%, <b>원 미만 절사</b>(26.1Q 품의·BPTG 26.2Q 전표로 확인). 총액 = 공급가액 + 부가세.</li>
          <li><b>분류 예외</b>: Partnership·STICKER(고요율), GAME, LICENSE(Product)는 라이선스로 봅니다(누적 피벗 색 구분 기준).</li>
          <li><b>품의 로열티 상세 표</b>: TREASURE = 총액 − 멤버 10명 개인 OEM. 판매월은 분기 마지막 달, 정산월은 품의 상신월입니다.</li>
        </ol>
      </details>
    </div>

    <!-- 탭 2: 품의 복붙 -->
    <div class="tabpanel" id="p-pum" role="tabpanel" aria-labelledby="t-pum" hidden>
      <section class="panel" id="s-pum">
        <div class="ph"><h2>품의 입력 내용</h2>
          <p>직전 분기 품의 양식·문구 그대로입니다. 그룹웨어 입력 순서대로 위에서부터 복사하세요.</p></div>
        <div class="form">
          <label for="inSettle">정산월 (품의 상신월)<input type="month" id="inSettle"></label>
          <label for="inTx">거래일자 (세금계산서 작성일)<input type="date" id="inTx"></label>
          <label for="inPay">입금예정일 (이내)<input type="date" id="inPay"></label>
          <div class="allcopy"><button class="btn primary" data-copy="all">본문 전체 복사</button></div>
        </div>
        <div class="block"><header><h3>제목</h3><button class="btn sm" data-copy="title">복사</button></header>
          <div class="body"><div id="pTitle"></div></div></div>
        <div class="block"><header><h3>매출 품의 필수 입력 항목</h3><button class="btn sm" data-copy="req">전체 복사</button></header>
          <div class="body"><dl class="kv" id="pReq"></dl></div></div>
        <div class="block"><header><h3>1. 기본 내역 · 2. 세금계산서 발행 정보</h3><button class="btn sm" data-copy="basic">전체 복사</button></header>
          <div class="body"><dl class="kv" id="pBasic"></dl></div></div>
        <div class="block fixed"><header><h3>3. 상세내역 — 1. 로열티 상세 내역 <span class="stamp">양식 고정 · 유관부서 자동 이관</span></h3>
            <div class="row"><button class="btn sm" data-copy="fixed">양식 복사</button><button class="btn sm" data-copy="fixedTable">표 복사</button></div></header>
          <div class="body"><pre class="fmt" id="pFixed"></pre><div class="tablebox"><table id="pFixedTable"></table></div>
            <p class="small muted" style="margin:0">줄 구성·순서·띄어쓰기를 바꾸지 마세요. 값만 분기에 맞춰 바뀝니다. 직전 품의 기준: 사업구조 RS, MG "-", 통화 KRW, 환율 1.</p></div></div>
        <div class="block"><header><h3>2. 매출 요약</h3><button class="btn sm" data-copy="summary">문구 복사</button></header>
          <div class="body"><div id="pSummary"></div>
            <div class="cols2">
              <div><div class="row" style="justify-content:space-between;margin-bottom:6px"><b class="small" id="catTitle">상품 카테고리별 로열티 분류</b><button class="btn sm" data-copy="catTable">표 복사</button></div><div class="tablebox"><table id="catTable"></table></div></div>
              <div><div class="row" style="justify-content:space-between;margin-bottom:6px"><b class="small" id="memTitle">멤버별 국가별 OEM 로열티 분류</b><button class="btn sm" data-copy="memTable">표 복사</button></div><div class="tablebox"><table id="memTable"></table></div></div>
            </div></div></div>
        <div class="block"><header><h3>3. 누적 로열티 · 4. 첨부파일 · 5. 참고사항</h3><button class="btn sm" data-copy="tail">문구 복사</button></header>
          <div class="body"><pre class="fmt" id="pTail"></pre><div class="hist" id="hist"></div></div></div>
        <div class="block"><header><h3>결재선 (위임전결기준표 261006 · 사업부문 매출 품의)</h3></header>
          <div class="body"><div id="pApproval"></div></div></div>
      </section>
    </div>

    <!-- 탭 3: ERP 전표 -->
    <div class="tabpanel" id="p-erp" role="tabpanel" aria-labelledby="t-erp" hidden>
      <section class="panel" id="s-erp">
        <div class="ph"><h2>ERP 매출 전표</h2>
          <p id="erpDesc">품의 승인 후 입력합니다.</p></div>
        <div class="block"><header><h3>ERP 업로드 파일 (거래내역 엑셀업로드)</h3><button class="btn primary" id="btnErp">ERP 업로드 파일 내려받기</button></header>
          <div class="body small">
            <ol style="margin:0;padding-left:18px;display:grid;gap:3px">
              <li>ERP 매출 신청 → 기안문서 연동 → 세금계산서 등록("부가세만 생성" 체크)까지 평소대로 진행</li>
              <li>거래내역 화면에서 <b>"거래내역 엑셀업로드"</b> → 내려받은 파일 선택 → 매출 16줄이 한 번에 들어감 (부가세·외상매출금 줄은 ERP가 자동 생성)</li>
              <li>날짜 오류만 확인 후 내역 저장 → 결재 생성</li>
            </ol>
            <p class="muted" style="margin:0">YGP MD 업로드 양식과 같은 시트·열 구성입니다. 업로드 화면이 .xls만 받으면 엑셀에서 "다른 이름으로 저장 → Excel 97-2003 통합 문서(.xls)"로 바꿔 올려 주세요. 첫 업로드 때 거래처 코드 인식 여부를 확인해 주세요.</p>
            <span id="erpMsg" class="muted"></span>
          </div></div>
        <div class="row"><button class="btn sm" data-copy="erpTable">ERP 표 복사 (직접 입력용)</button></div>
        <div class="tablebox"><table id="erpTable"></table></div>
      </section>
    </div>
`;

function mount(ctx) {
  ctx.root.innerHTML = VIEW; ctx.tabsEl.innerHTML = TABS_HTML;
  var CFG = {}, BASE_MAP = {}, HIST = [], KEY_2Q = null, KEY_1Q = null;
  var post = ctx.post, saveFile = ctx.saveFile;
  var $ = function (id) { return document.getElementById(id); };
  var fmt = TRUZ.fmt;
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var S = { mode: 'live', parsed: null, rates: {}, quarter: null, result: null,
    map: {}, extraMap: {}, saved: {}, fileName: '', inputs: { settle: '', tx: '', pay: '' } };
  var online = true;

  function toast(t) { var el = $('toast'); el.textContent = t; el.hidden = false; clearTimeout(toast._t); toast._t = setTimeout(function () { el.hidden = true; }, 2200); }
  function addMonths(ym, n) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + n; y += Math.floor(m / 12); m = ((m % 12) + 12) % 12; return y + '-' + TRUZ.pad(m + 1); }
  function monthEnd(ym) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7); return ym + '-' + TRUZ.pad(new Date(y, m, 0).getDate()); }
  function qInfo() {
    var q = S.quarter, startM = (q.q - 1) * 3 + 1, endM = q.q * 3;
    return { y: q.y, q: q.q, yy: String(q.y).slice(2), start: q.y + '-' + TRUZ.pad(startM) + '-01', end: monthEnd(q.y + '-' + TRUZ.pad(endM)),
      lastMonth: q.y + '-' + TRUZ.pad(endM), months: [0, 1, 2].map(function (i) { return q.y + '-' + TRUZ.pad(startM + i); }), key: q.y + 'Q' + q.q };
  }
  function defaults() {
    var qi = qInfo();
    if (!S.inputs.settle) S.inputs.settle = addMonths(qi.lastMonth, 2);
    if (!S.inputs.pay && S.inputs.tx) S.inputs.pay = monthEnd(addMonths(S.inputs.tx.slice(0, 7), 1));
  }

  // ---------- 계산 ----------
  function recompute() {
    if (S.mode === 'live') S.result = TRUZ.compute(S.parsed, S.rates, S.map, {});
    render();
  }

  // ---------- 렌더 ----------
  function render() {
    defaults();
    var r = S.result, qi = qInfo();
    $('chipQ').textContent = S.quarter.label;
    var cm = $('chipMode');
    if (S.mode === 'sample') { cm.className = 'chip sample'; cm.textContent = '예시: 26.2Q 확정 실적'; }
    else { cm.className = 'chip live'; cm.textContent = '작업 중: ' + S.fileName; }
    $('inSettle').value = S.inputs.settle; $('inTx').value = S.inputs.tx; $('inPay').value = S.inputs.pay;
    $('inTx').classList.toggle('need', !S.inputs.tx);
    renderFx(); renderFlags(); renderAmt(); renderPum(); renderErp();
    $('btnXlsx').disabled = S.mode !== 'live';
    $('btnSave').disabled = S.mode !== 'live' || !online;
    $('outMsg').textContent = S.mode !== 'live' ? '원본을 올리면 사용할 수 있습니다.' : (online ? '' : '서버에 연결되지 않아 확정 저장이 꺼져 있습니다.');
  }

  function renderFx() {
    var keys = Object.keys(S.rates), months = Array.from(new Set(keys.map(function (k) { return k.split('|')[0]; }))).sort();
    var curs = Array.from(new Set(keys.map(function (k) { return k.split('|')[1]; }))).filter(function (c) { return c !== 'KRW'; }).sort();
    var h = '<tr><th>통화</th>' + months.map(function (m) { return '<th class="n">' + m + '</th>'; }).join('') + '</tr>';
    curs.forEach(function (c) {
      h += '<tr><td><b>' + c + '</b> <span class="muted small">KRW / 1' + c + '</span></td>' + months.map(function (m) {
        var k = m + '|' + c, v = S.rates[k];
        if (!(k in S.rates)) return '<td class="n muted">—</td>';
        return '<td class="n"><input class="rate' + (v == null || v === '' ? ' empty' : '') + '" id="fx-' + k.replace('|', '-') + '" data-k="' + k + '" inputmode="decimal" value="' + (v == null ? '' : v) + '"' + (S.mode === 'sample' ? ' disabled' : '') + '></td>';
      }).join('') + '</tr>';
    });
    h += '<tr><td><b>KRW</b> <span class="muted small">국내(LS)</span></td>' + months.map(function () { return '<td class="n num">1</td>'; }).join('') + '</tr>';
    $('fxTable').innerHTML = h;
  }

  function renderFlags() {
    var r = S.result, fl = r.flags.slice();
    if (S.mode === 'sample') fl.unshift({ level: 'warn', text: '지금 화면은 26.2Q 확정 실적 예시입니다. 3분기 원본을 올리면 그 파일로 다시 계산합니다.' });
    if (S.mode === 'live' && KEY_2Q && S.quarter.key === '2026Q2') {
      var same = r.supply === KEY_2Q.supply && r.vat === KEY_2Q.vat && r.lines.slice(0, 16).every(function (l, i) { return l.amt === KEY_2Q.lines[i]; });
      fl.unshift(same ? { level: 'ok', text: '26.2Q 실제 품의·ERP 전표와 공급가액·부가세·16개 라인 모두 1원 단위 일치' }
        : { level: 'bad', text: '26.2Q 실제 ERP 전표와 불일치 — 환율 또는 매핑을 확인해 주세요' });
    }
    if (S.mode === 'live' && KEY_1Q && S.quarter.key === '2026Q1') {
      var ok1 = r.supply === KEY_1Q.supply && r.vat === KEY_1Q.vat;
      fl.unshift(ok1 ? { level: 'ok', text: '26.1Q 실제 품의와 공급가액 ' + fmt(KEY_1Q.supply) + '원·부가세 ' + fmt(KEY_1Q.vat) + '원 일치' }
        : { level: 'bad', text: '26.1Q 실제 품의 금액과 불일치 — 확인 필요' });
    }
    var prev = prevQuarterAmt();
    if (prev) { var ch = (r.supply - prev[1]) / prev[1] * 100; fl.push({ level: Math.abs(ch) > 50 ? 'warn' : 'ok', text: '직전 분기(' + prev[0].replace('Q', '.') + 'Q) ' + fmt(prev[1]) + '원 대비 ' + (ch >= 0 ? '+' : '') + ch.toFixed(1) + '%' + (Math.abs(ch) > 50 ? ' — 변동 폭이 커서 품의 참고사항에 사유 기재 권장' : '') }); }
    var nb = fl.filter(function (f) { return f.level === 'bad'; }).length, nw = fl.filter(function (f) { return f.level === 'warn'; }).length;
    var st = $('statusLine'), lvl = nb ? 'bad' : nw ? 'warn' : 'ok';
    st.className = 'status ' + lvl;
    st.innerHTML = '<b>' + (nb ? '조치 필요 ' + nb + '건' : nw ? '체크 ' + nw + '건' : '검증 이상 없음') + '</b><span>' + esc((fl.filter(function (f) { return f.level === lvl; })[0] || {}).text || '') + '</span><span class="small">상세 보기 ›</span>';
    $('checkSum').textContent = '조치 ' + nb + ' · 체크 ' + nw + ' · 확인 ' + (fl.length - nb - nw);
    if (S.mode === 'live' && (nb || nw)) $('dCheck').open = true;
    $('flags').innerHTML = fl.map(function (f) { return '<div class="flag ' + f.level + '"><b>' + { ok: '확인', warn: '주의', bad: '조치' }[f.level] + '</b><span>' + esc(f.text) + '</span></div>'; }).join('');
    renderMapFix();
  }

  function renderMapFix() {
    var box = $('mapFix'), un = S.result.unmapped || {}, names = Object.keys(un);
    if (!names.length) { box.hidden = true; box.innerHTML = ''; return; }
    box.hidden = false;
    var opts = TRUZ.MEMBER_NAMES.concat(['트레저 12인']).map(function (m) { return '<option>' + esc(m) + '</option>'; }).join('');
    box.innerHTML = '<div class="block"><header><h3>새 제품 매핑 (' + names.length + '개)</h3><button class="btn sm primary" id="btnMap">매핑 적용' + (online ? '·저장' : '') + '</button></header><div class="body"><p class="small muted" style="margin:0">제품명의 캐릭터(루루=하루토, 본본=준규, 칠리=최현석, 포동=소정환, 라우=윤재혁, 로미=지훈 등)를 보고 멤버를 고르세요. 저장하면 다음 분기부터 자동 적용됩니다.</p><div class="tablebox"><table><tr><th>제품명</th><th class="n">KRW</th><th>멤버</th></tr>' +
      names.map(function (n, i) { return '<tr><td>' + esc(n) + '</td><td class="n num">' + fmt(un[n]) + '</td><td><select id="map-' + i + '" data-name="' + esc(n) + '"><option value="">선택</option>' + opts + '</select></td></tr>'; }).join('') + '</table></div></div></div>';
    $('btnMap').onclick = applyMappings;
  }

  function prevQuarterAmt() {
    var all = allHist(), k = S.quarter.key, idx = all.findIndex(function (h) { return h[0] === k; });
    if (idx > 0) return all[idx - 1];
    if (idx < 0 && all.length) return all[all.length - 1];
    return null;
  }
  function allHist() {
    var m = {}; HIST.forEach(function (h) { m[h[0]] = h[1]; });
    Object.keys(S.saved).forEach(function (k) { m[k] = S.saved[k].supply; });
    return Object.keys(m).sort().map(function (k) { return [k, m[k]]; });
  }
  function cumulativeThrough() {
    var all = allHist().filter(function (h) { return h[0] !== S.quarter.key && h[0] < S.quarter.key; });
    return all.reduce(function (a, h) { return a + h[1]; }, 0) + S.result.supply;
  }

  function renderAmt() {
    var r = S.result, a = TRUZ.approval(r.supply), qi = qInfo();
    $('tSupply').textContent = fmt(r.supply); $('tVat').textContent = fmt(r.vat); $('tGross').textContent = fmt(r.gross);
    $('amtMeta').innerHTML = [
      ['정산 기간', qi.start + ' ~ ' + qi.end],
      ['위임전결 구간', a.band],
      ['ERP 결재', r.supply < 2e8 ? '2억원 미만 — 재무관리 담당(합) → 팀 리더 → 회계 담당 검토·승인' : '2억원 이상 — 재무관리 담당·리더(합) → 팀 리더 → 회계 담당 검토 → 리더 승인'],
      ['TRUZ 누적 로열티', fmt(cumulativeThrough()) + '원']
    ].map(function (x) { return '<div><span>' + x[0] + '</span>' + esc(x[1]) + '</div>'; }).join('');
  }

  // ---------- 품의 텍스트 ----------
  var TXT = {};
  function renderPum() {
    var r = S.result, qi = qInfo(), I = S.inputs, a = TRUZ.approval(r.supply), V = CFG.VENDOR || {}, P = CFG.PROJECT || {};
    var qName = qi.y + '년 ' + qi.q + '분기';
    TXT.title = "[매출품의] " + qName + " 'IPX-트레저' TRUZ 캐릭터 상품 판매 로열티 정산의 건";
    $('pTitle').textContent = TXT.title;

    var req = [
      ['거래처명', V.name + ' (' + V.code + ')'], ['거래일자', I.tx || '입력 필요 — 세금계산서 작성일'],
      ['금액 (VAT 제외)', fmt(r.supply)], ['VAT', fmt(r.vat)], ['총금액', fmt(r.gross)],
      ['프로젝트', P.proj], ['입금예정일 (이내)', I.pay || '입력 필요'],
      ['아티스트(그룹)', P.group], ['아티스트(개인)', P.artist],
      ['매출 종류', '로열티'], ['비거주자 여부', '거주자'], ['외국환 신고 필요여부', '경상거래·자본거래 모두 해당없음']
    ];
    $('pReq').innerHTML = kv(req, 'req');
    TXT.req = req.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');

    var basic = [
      ['계약서 작성 여부', '작성'], ['전체계약기간', V.contract], ['통화종류', 'KRW'],
      ['청구방법', '정산서 수령 후 일시'], ['에이전시 유/무', '없음'], ['정산기간', qi.start.replace(/-/g, '.') + ' ~ ' + qi.end.replace(/-/g, '.')],
      ['발행정보', '세금계산서'], ['사업자등록번호', V.bizNo], ['담당자', V.taxContact], ['이메일', V.taxEmail],
      ['별도요청사항', '세금계산서 수신자 추가: ' + V.ccEmail]
    ];
    $('pBasic').innerHTML = kv(basic, 'basic');
    TXT.basic = basic.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');

    TXT.fixed = '1. 로열티 상세 내역\n\n- 정산월 : ' + I.settle + '\n\n- 프로젝트 : TRUZ\n\n- 사업구조 : RS\n\n- MG : -\n\n- 통화 : KRW\n\n- 환율 : 1';
    $('pFixed').textContent = TXT.fixed;
    var fh = ['판매월', '아티스트', '권역', '판매수량', '총 매출 (vat-)', '로열티 (vat-)'];
    var fr = r.royaltyTable.map(function (x) { return [qi.lastMonth, x.artist, '글로벌', '-', '-', fmt(x.amt)]; });
    TABLES.fixedTable = { head: fh, rows: fr, numCols: [5] };
    $('pFixedTable').innerHTML = tableHtml(TABLES.fixedTable);

    TXT.summary = '2. [매출 요약]\n- TRUZ ' + qi.yy + '년 ' + qi.q + '분기 총 로열티 : ' + fmt(r.supply) + ' 원' + (r.roundNote ? '\n' + r.roundNote : '');
    $('pSummary').innerHTML = esc('2. [매출 요약] - TRUZ ' + qi.yy + '년 ' + qi.q + '분기 총 로열티 : ' + fmt(r.supply) + ' 원') + (r.roundNote ? '<br><b style="color:var(--warn)">' + esc(r.roundNote) + '</b> <span class="small muted">(카테고리 표 아래에 붙여 넣기)</span>' : '');
    $('catTitle').textContent = '[매출 상세 표 - 상품 카테고리별 로열티 분류_' + qi.yy + '.' + qi.q + 'Q] (단위 : 원)';
    $('memTitle').textContent = '[매출 상세 표 - 멤버별 국가별 OEM 로열티 분류_' + qi.yy + '.' + qi.q + 'Q] (단위 : 원)';
    var mlab = qi.months.map(function (m) { return (+m.slice(5)) + '월'; });
    var crow = TRUZ.CATS.map(function (c) { var vals = qi.months.map(function (m) { return (r.catMonthR || {})[c.key] ? (r.catMonthR[c.key][m] || 0) : Math.round(r.catMonth[c.key][m] || 0); }); return [c.label].concat(vals.map(fmt), [fmt(vals.reduce(function (a, b) { return a + b; }, 0))]); });
    var colT = qi.months.map(function (m) { return TRUZ.CATS.reduce(function (a, c) { return a + Math.round(r.catMonth[c.key][m] || 0); }, 0); });
    crow.push(['TTL'].concat(colT.map(fmt), [fmt(colT.reduce(function (a, b) { return a + b; }, 0))]));
    TABLES.catTable = { head: ['구분'].concat(mlab, ['TTL']), rows: crow, numCols: [1, 2, 3, 4], sumLast: true };
    $('catTable').innerHTML = tableHtml(TABLES.catTable);
    var ctys = r.countries.length ? r.countries : ['CN', 'JP', 'LS', 'US'];
    var mrows = Object.keys(r.memCty).filter(function (m) { return r.mem[m]; }).map(function (m) { return [m].concat(ctys.map(function (c) { return fmt(r.memCty[m][c] || 0); }), [fmt(r.mem[m])]); });
    mrows.push(['총합계'].concat(ctys.map(function (c) { return fmt(Object.keys(r.memCty).reduce(function (a, m) { return a + (r.memCty[m][c] || 0); }, 0)); }), [fmt(Object.keys(r.mem).reduce(function (a, m) { return a + r.mem[m]; }, 0))]));
    TABLES.memTable = { head: ['멤버'].concat(ctys, ['총합계']), rows: mrows, numCols: ctys.map(function (_, i) { return i + 1; }).concat([ctys.length + 1]), sumLast: true };
    $('memTable').innerHTML = tableHtml(TABLES.memTable);

    var cum = cumulativeThrough();
    TXT.tail = '3. [매출 상세 표 - TRUZ 누적 로열티: 21년 1분기 ~ ' + qi.yy + '년 ' + qi.q + '분기] (단위 : 원)\n- TRUZ 총 누적 로열티 : ' + fmt(cum) + '원\n\n4. 첨부파일\n1) TRUZ_Royalty_Summary_' + qi.y + 'Q' + qi.q + '_매출\n\n5. 참고사항\n- 멤버별 OEM 로열티 데이터의 경우, 첨부 파일내 \'TRUZ_자재\'탭에서 확인 가능\n- 원계약 4월 28일 종료이나, 연장합의서 2차 협의로 인해 12월 31일까지 연장.\n\n이상입니다. 감사합니다.';
    $('pTail').textContent = TXT.tail;
    var hl = allHist(); if (!hl.some(function (h) { return h[0] === S.quarter.key; })) hl.push([S.quarter.key, r.supply]);
    $('hist').innerHTML = hl.map(function (h) { return '<div class="' + (h[0] === S.quarter.key ? 'cur' : '') + '"><span>' + h[0].slice(2).replace('Q', '.') + 'Q</span><span class="num">' + fmt(h[1]) + '</span></div>'; }).join('');
    var tsv = function (t) { return [t.head].concat(t.rows).map(function (r) { return r.join('\t'); }).join('\n'); };
    TXT.all = [TXT.title, '', TXT.req, '', '1. 기본 내역 / 2. 세금계산서 발행 정보', TXT.basic, '', '3. 상세내역', TXT.fixed, '', tsv(TABLES.fixedTable), '', TXT.summary, '',
      $('catTitle').textContent, tsv(TABLES.catTable), '', $('memTitle').textContent, tsv(TABLES.memTable), '', TXT.tail].join('\n');
    $('pApproval').innerHTML = '<b>' + esc(a.band) + '</b><br><span class="muted">' + esc(a.line) + '</span>' + (CFG.APPROVAL_NOTE ? '<br><span class="small muted">' + esc(CFG.APPROVAL_NOTE) + '</span>' : '');
  }

  function kv(list, grp) {
    return list.map(function (x, i) { return '<dt>' + esc(x[0]) + '</dt><dd>' + esc(x[1]) + '</dd><button class="btn sm" data-copyval="' + grp + ':' + i + '">복사</button>'; }).join('');
  }

  // ---------- ERP ----------
  function renderErp() {
    var r = S.result, qi = qInfo(), pre = qi.y + '.' + qi.q + 'Q 트레저 TRUZ 콜라보 제품 판매 로열티';
    var memInfo = {}; TRUZ.MEMBERS.forEach(function (x) { memInfo[x.m] = x; });
    var t10 = memInfo['트레저 10인'] || {}, rows = [['1', '부가세예수금-2010502', t10.proj + '-' + t10.code, pre, '대변', fmt(r.vat)]];
    r.lines.forEach(function (l, i) {
      var mi = memInfo[l.member] || memInfo['트레저 10인'];
      rows.push([String(i + 2), '로열티매출_초상-' + (CFG.PROJECT || {}).account, mi.proj + '-' + mi.code, pre + ' : ' + l.label, '대변', fmt(l.amt)]);
    });
    rows.push([String(r.lines.length + 2), '한화외상매출금-101030101', '프로젝트공통', qi.y + '년 ' + qi.q + "분기 'IPX-트레저' TRUZ 캐릭터 상품 판매 로열티 정산의 건", '차변', fmt(r.gross)]);
    TABLES.erpTable = { head: ['순번', '계정과목', '프로젝트', '적요', '차/대', '금액'], rows: rows, numCols: [5] };
    $('erpTable').innerHTML = tableHtml(TABLES.erpTable) ;
  }

  var TABLES = {};
  function tableHtml(t) {
    var h = '<tr>' + t.head.map(function (c, i) { return '<th' + (t.numCols.indexOf(i) >= 0 ? ' class="n"' : '') + '>' + esc(c) + '</th>'; }).join('') + '</tr>';
    t.rows.forEach(function (r, ri) {
      var sum = t.sumLast && ri === t.rows.length - 1;
      h += '<tr' + (sum ? ' class="sum"' : '') + '>' + r.map(function (c, i) { return '<td' + (t.numCols.indexOf(i) >= 0 ? ' class="n"' : '') + '>' + esc(c) + '</td>'; }).join('') + '</tr>';
    });
    return h;
  }

  // ---------- 복사 ----------
  function showFallback(text) { $('copyFallback').hidden = false; var ta = $('fbText'); ta.value = text; ta.focus(); ta.select(); toast('자동 복사가 막혀 있어 아래 칸에 표시했습니다'); }
  function copyText(text, label) {
    if (!navigator.clipboard || !navigator.clipboard.writeText) { showFallback(text); return; }
    navigator.clipboard.writeText(text).then(function () { toast((label || '') + ' 복사했습니다'); }, function () { showFallback(text); });
  }
  function copyTable(t, label) {
    var tsv = [t.head].concat(t.rows).map(function (r) { return r.join('\t'); }).join('\n');
    var html = '<table border="1" style="border-collapse:collapse">' + tableHtml(t).replace(/ class="[^"]*"/g, '') + '</table>';
    try {
      if (window.ClipboardItem && navigator.clipboard.write) {
        navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([tsv], { type: 'text/plain' }) })])
          .then(function () { toast(label + ' 복사했습니다 (표 서식 포함)'); }, function () { copyText(tsv, label); });
        return;
      }
    } catch (e) { }
    copyText(tsv, label);
  }
  ctx.root.addEventListener('click', function (e) {
    var b = e.target.closest('[data-copy]');
    if (b) {
      var k = b.getAttribute('data-copy');
      if (TABLES[k]) copyTable(TABLES[k], '표를');
      else if (TXT[k] != null) copyText(TXT[k], '내용을');
      return;
    }
    var v = e.target.closest('[data-copyval]');
    if (v) {
      var p = v.getAttribute('data-copyval').split(':'), lines = TXT[p[0]].split('\n'), line = lines[+p[1]] || '';
      copyText(line.split('\t').slice(1).join('\t'), '값을');
    }
  });

  // ---------- 입력 ----------
  ctx.root.addEventListener('input', function (e) {
    var t = e.target;
    if (t.classList.contains('rate')) {
      var k = t.getAttribute('data-k'), v = t.value.trim();
      S.rates[k] = v === '' ? null : parseFloat(v.replace(/,/g, ''));
      t.classList.toggle('empty', v === '');
      clearTimeout(recompute._t); recompute._t = setTimeout(function () { S.result = TRUZ.compute(S.parsed, S.rates, S.map, {}); renderFlags(); renderAmt(); renderPum(); renderErp(); }, 250);
    }
  });
  $('inSettle').addEventListener('change', function () { S.inputs.settle = this.value; renderPum(); });
  $('inTx').addEventListener('change', function () { S.inputs.tx = this.value; S.inputs.pay = this.value ? monthEnd(addMonths(this.value.slice(0, 7), 1)) : ''; render(); });
  $('inPay').addEventListener('change', function () { S.inputs.pay = this.value; renderPum(); });

  // ---------- 매핑 ----------
  function hashId(s) { var h = 2166136261; for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return 'p' + (h >>> 0).toString(16) + s.length; }
  function applyMappings() {
    var sels = document.querySelectorAll('#mapFix select'), adds = [];
    sels.forEach(function (s) { if (s.value) adds.push([s.getAttribute('data-name'), s.value]); });
    if (!adds.length) { toast('멤버를 하나 이상 선택해 주세요'); return; }
    adds.forEach(function (a) { S.map[TRUZ.N(a[0])] = a[1]; });
    recompute();
    if (online) {
      post('saveMapping', { items: adds.map(function (a) { return { name: a[0], member: a[1] }; }) })
        .then(function (res) { toast('매핑 ' + adds.length + '건을 적용하고 DB에 저장했습니다 (신규 ' + res.added + ' · 수정 ' + res.updated + ')'); },
          function (e) { toast('매핑은 적용했지만 저장하지 못했습니다: ' + e.message); });
    } else toast('매핑 ' + adds.length + '건을 이번 작업에 적용했습니다');
  }

  // ---------- 내보내기 ----------
  $('btnXlsx').addEventListener('click', async function () {
    if (S.mode !== 'live') return;
    var r = S.result, qi = qInfo();
    var numify = function (t) { return [t.head].concat(t.rows.map(function (row) { return row.map(function (c, i) { return t.numCols.indexOf(i) >= 0 && /^-?[\d,]+$/.test(c) ? +c.replace(/,/g, '') : c; }); })); };
    var pum = [['항목', '값'], ['제목', TXT.title], ['정산기간', qi.start + ' ~ ' + qi.end], ['공급가액', r.supply], ['부가세 (원 미만 절사)', r.vat], ['총액', r.gross],
      ['거래일자', S.inputs.tx], ['입금예정일', S.inputs.pay], ['정산월', S.inputs.settle], ['반올림 주석', r.roundNote || '없음'], ['원본 파일', S.fileName], [], ['1. 로열티 상세 내역 양식'], [TXT.fixed]].concat([[]], numify(TABLES.fixedTable));
    var hl = allHist().filter(function (h) { return h[0] !== S.quarter.key && h[0] < S.quarter.key; }).concat([[S.quarter.key, r.supply]]);
    var wb = TRUZ.buildReview(S.parsed, r, { hist: hl.map(function (h) { return [h[0].slice(0, 4) + '.Q' + h[0].slice(5) + '.', h[1]]; }), erp: numify(TABLES.erpTable), pum: pum });
    var out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    var name = S.fileName.replace(/\.xlsx?$/i, '') + '_YG검토.xlsx';
    saveFile(name, out); toast('YG검토본을 내려받았습니다');
  });

  $('btnErp').addEventListener('click', async function () {
    if (S.mode !== 'live') { $('erpMsg').textContent = '정산 탭에서 원본을 먼저 올려 주세요.'; return; }
    var qi = qInfo(), o = TRUZ.buildErpUpload(S.parsed, S.result, qi);
    if (S.result.mem['(미매핑)']) { $('erpMsg').textContent = '미매핑 제품이 있어 업로드 파일을 만들지 않았습니다. 정산 탭 검증에서 매핑을 먼저 지정해 주세요.'; return; }
    var out = XLSX.write(o.wb, { type: 'array', bookType: 'xlsx' });
    saveFile('AR거래내역엑셀업로드_TRUZ_' + qi.y + 'Q' + qi.q + '.xlsx', out); toast('ERP 업로드 파일을 내려받았습니다 (' + o.rows.length + '줄)');
  });

  $('btnSave').addEventListener('click', async function () {
    if (!online || S.mode !== 'live') return;
    var r = S.result, bad = r.flags.some(function (f) { return f.level === 'bad'; });
    if (bad) { $('outMsg').textContent = '"조치" 항목이 남아 있어 확정 저장을 막았습니다. 검증 단계를 먼저 정리해 주세요.'; return; }
    try {
      await post('saveQuarter', { record: { quarter: S.quarter.key, supply: r.supply, vat: r.vat, gross: r.gross, settleMonth: S.inputs.settle,
        txDate: S.inputs.tx, payDate: S.inputs.pay, file: S.fileName, roundNote: r.roundNote, lines: r.lines.map(function (l) { return { label: l.label, amt: l.amt }; }),
        royaltyTable: r.royaltyTable } });
      S.saved[S.quarter.key] = { supply: r.supply }; render(); toast(S.quarter.label + ' 확정 저장했습니다');
    } catch (e) { $('outMsg').textContent = '저장하지 못했습니다 (' + e.message + ')'; }
  });

  // ---------- 탭 ----------
  var TABS = ['calc', 'pum', 'erp'];
  function showTab(t) {
    if (TABS.indexOf(t) < 0) t = 'calc';
    TABS.forEach(function (k) { $('p-' + k).hidden = k !== t; $('t-' + k).setAttribute('aria-selected', k === t ? 'true' : 'false'); });
    try { localStorage.setItem('truz.tab', t); } catch (e) { }
  }
  ctx.tabsEl.querySelectorAll('.tab').forEach(function (b) { b.addEventListener('click', function () { showTab(b.getAttribute('data-tab')); window.scrollTo(0, 0); }); });
  $('statusLine').addEventListener('click', function () { var d = $('dCheck'); d.open = true; d.scrollIntoView({ behavior: 'smooth', block: 'start' }); });
  (function () { var h = (location.hash || '').slice(1), saved = null; try { saved = localStorage.getItem('truz.tab'); } catch (e) { } showTab(TABS.indexOf(h) >= 0 ? h : saved || 'calc'); })();

  // ---------- 시작: 이 정산의 비공개 설정을 받은 뒤 업로드 파일 계산 ----------
  var j = ctx.data || {};
  CFG = j.cfg || {}; TRUZ.setConfig(CFG);
  BASE_MAP = j.map || {}; HIST = j.hist || []; KEY_2Q = CFG.KEY_2Q || null; KEY_1Q = CFG.KEY_1Q || null;
  S.map = Object.assign({}, BASE_MAP); S.saved = j.saved || {};
  var P = CFG.PROJECT || {}, V = CFG.VENDOR || {};
  $('erpDesc').textContent = '품의 승인 후 입력합니다. 계정 로열티매출_초상(' + P.account + ') · 회계부서 광고IP사업팀(' + P.dept + ') · 거래처 ' + V.name + '(' + V.code + ') · 국가구분 KR. 순서는 직전 분기 ERP 전표와 같습니다. 세금계산서 등록 시 "부가세만 생성"을 체크하고, 거래내역 저장 후 결재를 생성합니다.';
  var p = TRUZ.parse(ctx.wb);
  S.parsed = p; S.rates = Object.assign({}, p.rates); S.quarter = p.quarter; S.fileName = ctx.fileName;
  S.result = TRUZ.compute(p, S.rates, S.map, {});
  render();
  return { quarter: p.quarter.label, note: p.quarter.label + ' · Data ' + p.data.length + '행 · 품목 상세 ' + p.detail.length + '행' };
}

DESK.register({
  id: 'truz', name: 'TRUZ', desc: 'IPX × 트레저 캐릭터 상품 로열티 · 분기 정산',
  fileHint: 'IPX 원본 TRUZ_Royalty_Summary_20xxQn_….xlsx',
  outputs: ['YG검토본', '품의 입력 내용', 'ERP 전표 업로드 파일'],
  detect: function (wb) { try { TRUZ.parse(wb); return true; } catch (e) { return false; } },
  mount: mount
});
})();
