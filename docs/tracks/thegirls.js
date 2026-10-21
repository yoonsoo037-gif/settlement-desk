/* 정산 모듈: BPTG OST 'THE GIRLS' 매입 정산 2종
 *  - tg-takeone : 테이크원컴퍼니 음원 로열티 수익 배분 (분기, KRW, 세금계산서)
 *  - tg-writer  : 해외 작가(PATRIOT·LINDGREN) 음원·OST MERCH 로열티 배분 (반기, USD, 원천세)
 * 입력은 회계팀 "THE GIRLS 음원 매출 데이터" 엑셀("매출 정리" 시트). 두 모듈이 같은 파일을 받으므로 셸에서 정산 종류를 고른다.
 * 서버 트랙은 둘 다 'thegirls' 하나를 쓴다: _매핑 탭 = 누적 원장(행 키 → 행 JSON), _이력/_확정 = 회차 기록.
 * 회사 데이터(코드·거래처·계좌·연락처·검증 기준)는 이 파일에 없고 로그인 후 서버 설정에서 받는다. */
(function () {
  var TG = (function () {
    var N = function (s) { return String(s == null ? '' : s).replace(/[\s\\]+/g, ''); };
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    var fmt = function (n) { return n == null || isNaN(n) ? '-' : Math.round(n).toLocaleString('en-US'); };
    var fmt2 = function (n) { return (Math.round(n * 100) / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
    var r2 = function (n) { return Math.round(n * 100) / 100; };
    var floor10 = function (n) { return Math.floor(n / 10 + 1e-9) * 10; };

    // 회계일자: '2026-04-20' / '2023/10/20' / '7/20/2024' / 엑셀 일련번호 / Date → 'YYYY-MM-DD'
    function normDate(v) {
      if (v == null || v === '') return null;
      if (v instanceof Date && !isNaN(v)) return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
      if (typeof v === 'number' && v > 20000 && v < 80000) { var d = new Date(Math.round((v - 25569) * 864e5)); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
      var s = String(v).trim(), m;
      if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/))) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
      if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return m[3] + '-' + pad(+m[1]) + '-' + pad(+m[2]);
      return null;
    }
    function num(v) {
      if (typeof v === 'number') return v;
      if (v == null) return 0;
      var s = String(v).replace(/[,\s₩]/g, ''), neg = /^\(.*\)$/.test(s);
      var n = parseFloat(s.replace(/[()]/g, ''));
      return isNaN(n) ? 0 : (neg ? -n : n);
    }
    var isAdsense = function (r) { return /애드센스|유튜브/.test(r.desc) || /^Google/i.test(r.vendor || ''); };

    // 적요 앞머리의 판매 기간: '2026.05' / '26.1Q' / '2025.01-03' / '23.07-23.10' → {start:'YYYY-MM', end:'YYYY-MM'}
    function period(desc) {
      var s = String(desc || '').replace(/^\[[^\]]*\]+\s*/, '').replace(/^\[[^\]]*\]/, '');
      var yy = function (y) { y = +y; return y < 100 ? 2000 + y : y; };
      var m = s.match(/(\d{2,4})\.(\d)Q(?:\s*-\s*(\d)Q)?/);
      var m2 = s.match(/(\d{2,4})\.(\d{1,2})(?:\.\d{1,2})?\s*-\s*(?:(\d{2,4})\.)?(\d{1,2})(?!\d)/);
      var m3 = s.match(/(\d{2,4})\.(\d{1,2})(?!\d|Q)/);
      var idxQ = m ? s.indexOf(m[0]) : 1e9, idx2 = m2 ? s.indexOf(m2[0]) : 1e9, idx3 = m3 ? s.indexOf(m3[0]) : 1e9;
      if (m && idxQ <= idx2 && idxQ <= idx3) {
        var y = yy(m[1]), q1 = +m[2], q2 = m[3] ? +m[3] : q1;
        return { start: y + '-' + pad(q1 * 3 - 2), end: y + '-' + pad(q2 * 3) };
      }
      if (m2 && idx2 <= idx3 && +m2[2] <= 12 && +m2[4] <= 12) {
        var ys = yy(m2[1]), ye = m2[3] ? yy(m2[3]) : ys;
        return { start: ys + '-' + pad(+m2[2]), end: ye + '-' + pad(+m2[4]) };
      }
      if (m3 && +m3[2] >= 1 && +m3[2] <= 12) { var ym = yy(m3[1]) + '-' + pad(+m3[2]); return { start: ym, end: ym }; }
      return null;
    }

    // "매출 정리"(없으면 머리글이 회계일자·적요·매출인 시트)에서 행 추출
    function findSheet(wb) {
      var names = wb.SheetNames.slice().sort(function (a, b) { return (N(b) === '매출정리') - (N(a) === '매출정리'); });
      for (var i = 0; i < names.length; i++) {
        var rows = XLSX.utils.sheet_to_json(wb.Sheets[names[i]], { header: 1, raw: true, defval: null });
        for (var h = 0; h < Math.min(rows.length, 10); h++) {
          var hd = (rows[h] || []).map(N);
          if (hd.indexOf('회계일자') >= 0 && hd.indexOf('적요') >= 0 && hd.indexOf('매출') >= 0 && hd.indexOf('원천세') >= 0) return { name: names[i], rows: rows, h: h, hd: hd };
        }
      }
      return null;
    }
    function parse(wb) {
      var f = findSheet(wb);
      if (!f) throw new Error('"매출 정리" 시트(회계일자·적요·매출·거래처·원천세)를 찾지 못했습니다.');
      var c = { d: f.hd.indexOf('회계일자'), s: f.hd.indexOf('적요'), a: f.hd.indexOf('매출'), v: f.hd.indexOf('거래처'), w: f.hd.indexOf('원천세') };
      var out = [], seen = {};
      f.rows.slice(f.h + 1).forEach(function (r) {
        var d = normDate(r[c.d]), desc = r[c.s];
        if (!d || desc == null || String(desc).trim() === '' || r[c.a] == null || r[c.a] === '') return;
        var row = { d: d, s: String(desc).trim(), a: num(r[c.a]), v: String(r[c.v] == null ? '' : r[c.v]).trim(), w: num(r[c.w]) };
        var k = rowKey(row); seen[k] = (seen[k] || 0) + 1; if (seen[k] > 1) k += '#' + seen[k];
        row.key = k; out.push(row);
      });
      if (!out.length) throw new Error('매출 정리 시트에 정산 행이 없습니다.');
      if (!out.some(function (r) { return /THE\s*GIRLS/i.test(r.s); })) throw new Error('THE GIRLS 음원 매출 데이터가 아닙니다.');
      var ledger = wb.SheetNames.filter(function (n) { return N(n) === '원장'; })[0];
      return { sheet: f.name, rows: out, ledgerSheet: ledger ? wb.Sheets[ledger] : null };
    }
    function rowKey(r) { return r.d + '|' + N(r.s) + '|' + Math.round(r.a); }
    var withDesc = function (r) { r.desc = r.s; r.vendor = r.v; return r; };

    var q = function (d) { return { y: +d.slice(0, 4), q: Math.ceil(+d.slice(5, 7) / 3) }; };
    var qLabel = function (x) { return String(x.y).slice(2) + '.' + x.q + 'Q'; };
    var qStart = function (x) { return x.y + '-' + pad(x.q * 3 - 2) + '-01'; };
    function monthEnd(ym) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7); return ym + '-' + pad(new Date(y, m, 0).getDate()); }
    var qEnd = function (x) { return monthEnd(x.y + '-' + pad(x.q * 3)); };
    function addMonths(ym, n) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + n; y += Math.floor(m / 12); m = ((m % 12) + 12) % 12; return y + '-' + pad(m + 1); }

    // 테이크원: 행 상태 판정. settled = 이미 정산에 쓴 행 키 집합, cutoff = 'YYYY-MM'(계약 종료 기준 판매월)
    function classify(rows, settled, cutoff) {
      return rows.map(function (r0) {
        var r = withDesc(Object.assign({}, r0)), p = period(r.s);
        r.p = p;
        if (isAdsense(r)) { r.st = 'adsense'; r.sel = false; }
        else if (settled[r.key]) { r.st = 'settled'; r.sel = false; r.tk = settled[r.key]; }
        else if (p && p.start > cutoff) { r.st = 'after'; r.sel = false; }
        else if (p && p.end > cutoff) { r.st = 'straddle'; r.sel = true; }
        else if (!p) { r.st = 'unknown'; r.sel = true; }
        else { r.st = 'ok'; r.sel = true; }
        return r;
      });
    }
    function takeone(rows, rate) {
      var sel = rows.filter(function (r) { return r.sel; });
      var sales = sel.reduce(function (a, r) { return a + r.a; }, 0), wht = sel.reduce(function (a, r) { return a + r.w; }, 0);
      var net = sales - wht, supply = Math.round(net * rate), vat = Math.floor(supply * 0.1 + 1e-9);
      var dates = sel.map(function (r) { return r.d; }).sort();
      var qa = dates.length ? q(dates[0]) : null, qb = dates.length ? q(dates[dates.length - 1]) : null;
      var label = qa ? (qLabel(qa) === qLabel(qb) ? qLabel(qa) : qLabel(qa) + '-' + qLabel(qb)) : '';
      return { sel: sel, sales: sales, wht: wht, net: Math.round(net), netRaw: net, supply: supply, vat: vat, gross: supply + vat,
        label: label, rangeText: label.replace('-', ' ~ '), from: qa ? qStart(qa) : '', to: qb ? qEnd(qb) : '' };
    }

    // 정산서 "정산 기간" 문구 (이전 정산서 형식): 한 분기 → "2026년 3분기 내용", 여러 분기 → "2025년 4분기 미정산 분 및 2026년 2분기 내용",
    // 판매월이 계약 종료 기준 월에서 끊기면 끝에 "(26.5월)"
    function salesEnd(to, cutoff, from) { var ce = cutoff ? monthEnd(cutoff) : ''; return ce && ce < to && !(from && ce < from) ? ce : to; }
    function periodText(from, to, cutoff) {
      if (!from || !to) return '';
      var qk = function (d) { return d.slice(0, 4) + '년 ' + (Math.floor((+d.slice(5, 7) - 1) / 3) + 1) + '분기'; };
      var end = salesEnd(to, cutoff, from), a = qk(from), b = qk(end);
      var t = a === b ? a + ' 내용' : a + ' 미정산 분 및 ' + b + ' 내용';
      if (end !== to) t += '(' + end.slice(2, 4) + '.' + (+end.slice(5, 7)) + '월)';
      return t;
    }

    // 작가: 회계일자가 반기 안인 행(애드센스 제외) → 음원 순매출 × 요율, OST MERCH(제조원가 제외 매출 × 23%) × 요율
    function half(y, h) { return { y: y, h: h, from: y + (h === 1 ? '-01-01' : '-07-01'), to: y + (h === 1 ? '-06-30' : '-12-31'), label: String(y).slice(2) + '.' + h + 'H', kor: String(y).slice(2) + '년 ' + (h === 1 ? '상반기' : '하반기'), eng: h + 'H' + String(y).slice(2) }; }
    function writer(rows, hf, merch, fx, writers, merchShare) {
      var sel = rows.map(withDesc).filter(function (r) { return !isAdsense(r) && r.d >= hf.from && r.d <= hf.to; })
        .sort(function (a, b) { return a.d < b.d ? -1 : a.d > b.d ? 1 : 0; });
      var sales = sel.reduce(function (a, r) { return a + r.a; }, 0), wht = sel.reduce(function (a, r) { return a + r.w; }, 0), net = sales - wht;
      var merchYG = Math.round((merch.net || 0) * merchShare);
      var ws = writers.map(function (w) {
        var ostK = Math.round(net * w.rate), merchK = Math.round(merchYG * w.rate), totalK = ostK + merchK;
        var ostU = fx ? r2(ostK / fx) : 0, merchU = fx ? r2(merchK / fx) : 0, totalU = r2(ostU + merchU);
        var corp = floor10(totalK * 0.1), local = floor10(corp * 0.1), pay = totalK - corp - local;
        return Object.assign({}, w, { ostK: ostK, merchK: merchK, totalK: totalK, ostU: ostU, merchU: merchU, totalU: totalU,
          checkU: fx ? r2(totalK / fx) : 0, corp: corp, local: local, wht: corp + local, pay: pay, payU: fx ? r2(pay / fx) : 0 });
      });
      return { sel: sel, sales: Math.round(sales), wht: Math.round(wht), net: Math.round(net), merchYG: merchYG, merch: merch, fx: fx, hf: hf, writers: ws,
        totalK: ws.reduce(function (a, w) { return a + w.totalK; }, 0), totalU: r2(ws.reduce(function (a, w) { return a + w.totalU; }, 0)),
        salesU: fx ? r2((Math.round(net) + merchYG) / fx) : 0, rateSum: ws.reduce(function (a, w) { return a + w.rate; }, 0) };
    }

    // MD 관련 수불부 "FY YYYY_MD 관련 Summary" 시트: 'OST 패키지'(매출) · 'OST 패키지 매출총이익' 행의 월별 값 → 반기 합계
    function merchFromLedger(wb, hf) {
      var y = hf.y, name = wb.SheetNames.filter(function (n) { return new RegExp('FY\\s*' + y).test(n) && /Summary/i.test(n); })[0];
      if (!name) throw new Error('"FY ' + y + '_MD 관련 Summary" 시트를 찾지 못했습니다.');
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
      var hi = rows.findIndex(function (r) { return r.some(function (c) { return N(c) === y + '년1월'; }); });
      if (hi < 0) throw new Error('수불부 Summary 시트에서 월 머리글(' + y + '년 1월)을 찾지 못했습니다.');
      var col = {}; rows[hi].forEach(function (c, i) { var m = N(c).match(/^(\d{4})년(\d{1,2})월$/); if (m && +m[1] === y) col[+m[2]] = i; });
      var months = hf.h === 1 ? [1, 2, 3, 4, 5, 6] : [7, 8, 9, 10, 11, 12];
      var find = function (label) { for (var i = hi + 1; i < rows.length; i++) if (rows[i].some(function (c) { return N(c) === label; })) return rows[i]; return null; };
      var rs = find('OST패키지'), rn = find('OST패키지매출총이익');
      if (!rs || !rn) throw new Error('수불부에서 "OST 패키지" 매출·매출총이익 행을 찾지 못했습니다.');
      var sum = function (r) { return months.reduce(function (a, m) { return a + (col[m] != null ? num(r[col[m]]) : 0); }, 0); };
      var byMonth = months.map(function (m) { return { m: m, sales: col[m] != null ? num(rs[col[m]]) : 0, net: col[m] != null ? num(rn[col[m]]) : 0 }; });
      // OST MERCH = OST 패키지만. 24.2H 작가 매입품의(리유저블백·하트쿠션 총이익 +1,631,898원 발생 반기)도 OST 패키지 합계 58,404,455원만 반영
      // → 두 품목은 합산하지 않고 확인용으로만 표시
      var extras = ['리유저블백', '하트쿠션'].map(function (k) { var a = find(k), b = find(k + '매출총이익'); return { item: k, sales: a ? Math.round(sum(a)) : 0, net: b ? Math.round(sum(b)) : 0 }; }).filter(function (x) { return x.sales || x.net; });
      return { sales: Math.round(sum(rs)), net: Math.round(sum(rn)), byMonth: byMonth, sheet: name, extras: extras };
    }

    return { merchFromLedger: merchFromLedger, N: N, fmt: fmt, fmt2: fmt2, pad: pad, normDate: normDate, period: period, parse: parse, classify: classify, takeone: takeone,
      half: half, writer: writer, isAdsense: isAdsense, monthEnd: monthEnd, addMonths: addMonths, salesEnd: salesEnd, periodText: periodText, q: q, qLabel: qLabel, rowKey: rowKey };
  })();
  if (typeof module !== 'undefined') module.exports = TG;
  if (typeof DESK === 'undefined') return;

  // ---------------- 공통 화면 도우미 ----------------
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var fmt = TG.fmt, fmt2 = TG.fmt2;
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
  // 복사·탭·토스트: 모듈 루트에 위임
  function wire(ctx, TXT, TABLES, TABS, storeKey) {
    var $ = function (id) { return document.getElementById(id); };
    function showFallback(text) { $('copyFallback').hidden = false; var ta = $('fbText'); ta.value = text; ta.focus(); ta.select(); ctx.toast('자동 복사가 막혀 있어 아래 칸에 표시했습니다'); }
    function copyText(text, label) {
      if (!navigator.clipboard || !navigator.clipboard.writeText) { showFallback(text); return; }
      navigator.clipboard.writeText(text).then(function () { ctx.toast((label || '') + ' 복사했습니다'); }, function () { showFallback(text); });
    }
    function copyTable(t, label) {
      var tsv = [t.head].concat(t.rows).map(function (r) { return r.join('\t'); }).join('\n');
      var html = '<table border="1" style="border-collapse:collapse">' + tableHtml(t).replace(/ class="[^"]*"/g, '') + '</table>';
      try {
        if (window.ClipboardItem && navigator.clipboard.write) {
          navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([tsv], { type: 'text/plain' }) })])
            .then(function () { ctx.toast(label + ' 복사했습니다 (표 서식 포함)'); }, function () { copyText(tsv, label); });
          return;
        }
      } catch (e) { }
      copyText(tsv, label);
    }
    ctx.root.addEventListener('click', function (e) {
      var b = e.target.closest('[data-copy]');
      if (b) { var k = b.getAttribute('data-copy'); if (TABLES[k]) copyTable(TABLES[k], '표를'); else if (TXT[k] != null) copyText(TXT[k], '내용을'); return; }
      var v = e.target.closest('[data-copyval]');
      if (v) { var p = v.getAttribute('data-copyval').split(':'), line = (TXT[p[0]] || '').split('\n')[+p[1]] || ''; copyText(line.split('\t').slice(1).join('\t'), '값을'); }
    });
    function showTab(t) {
      if (TABS.indexOf(t) < 0) t = TABS[0];
      TABS.forEach(function (k) { $('p-' + k).hidden = k !== t; $('t-' + k).setAttribute('aria-selected', k === t ? 'true' : 'false'); });
      try { localStorage.setItem(storeKey, t); } catch (e) { }
    }
    ctx.tabsEl.querySelectorAll('.tab').forEach(function (b) { b.addEventListener('click', function () { showTab(b.getAttribute('data-tab')); window.scrollTo(0, 0); }); });
    var saved = null; try { saved = localStorage.getItem(storeKey); } catch (e) { }
    showTab(saved || TABS[0]);
    return { copyText: copyText, showTab: showTab };
  }
  function flagsHtml(fl) {
    return fl.map(function (f) { return '<div class="flag ' + f.level + '"><b>' + { ok: '확인', warn: '주의', bad: '조치' }[f.level] + '</b><span>' + esc(f.text) + '</span></div>'; }).join('');
  }
  function statusHtml(fl) {
    var nb = fl.filter(function (f) { return f.level === 'bad'; }).length, nw = fl.filter(function (f) { return f.level === 'warn'; }).length, lvl = nb ? 'bad' : nw ? 'warn' : 'ok';
    return { cls: 'status ' + lvl, html: '<b>' + (nb ? '조치 필요 ' + nb + '건' : nw ? '체크 ' + nw + '건' : '검증 이상 없음') + '</b><span>' + esc((fl.filter(function (f) { return f.level === lvl; })[0] || {}).text || '') + '</span>' };
  }
  function cfgOf(data) {
    var CFG = data.cfg || {}, ledger = {};
    Object.keys(data.map || {}).forEach(function (k) { try { var r = JSON.parse(data.map[k]); r.key = r.key || k; ledger[r.key] = r; } catch (e) { } });
    return { CFG: CFG, ledger: ledger, hist: data.hist || [], saved: data.saved || {} };
  }
  var PRJ = function (P) { return [['회계부서', P.deptName + ' (' + P.dept + ')'], ['프로젝트', P.proj + ' (' + P.projCode + ')'], ['계정', P.account + ' (' + P.accountCode + ')'], ['아티스트(그룹)', P.group + ' (' + P.groupCode + ')'], ['아티스트(개인)', P.artist + ' (' + P.artistCode + ')'], ['제작구분', P.prod]]; };
  var today = function () { var d = new Date(); return d.getFullYear() + '-' + TG.pad(d.getMonth() + 1) + '-' + TG.pad(d.getDate()); };
  var dotDate = function (d) { return d ? d.slice(2).replace(/-/g, '.') : ''; };

  // ================= ① 테이크원 (분기) =================
  function mountTakeone(ctx) {
    var $ = function (id) { return document.getElementById(id); };
    var D = cfgOf(ctx.data), CFG = D.CFG, P = CFG.PROJECT || {}, T = CFG.TAKEONE || {};
    var parsed = TG.parse(ctx.wb), settled = {};
    Object.keys(D.ledger).forEach(function (k) { if (D.ledger[k].tk) settled[k] = D.ledger[k].tk; });
    var S = { rows: TG.classify(parsed.rows, settled, T.cutoff || '2026-05'), inputs: { tx: '', label: '', period: '', sales: '', note: '', endNote: false, reshare: null } };
    var TXT = {}, TABLES = {}, R = null;

    ctx.tabsEl.innerHTML = tabsHtml([['calc', '정산'], ['out', '정산서·메일'], ['pum', '품의 복붙'], ['erp', 'ERP 전표']]);
    ctx.root.innerHTML =
      '<div class="tabpanel" id="p-calc"><section class="panel dash">' +
      '<div class="tiles"><div class="tile main"><span>매입금액 (VAT 제외)</span><b id="tSupply">0</b><em>실매출 × ' + Math.round((T.rate || 0) * 100) + '% · 품의 "매입금액"</em></div>' +
      '<div class="tile"><span>부가세 (10%, 원 미만 절사)</span><b id="tVat">0</b><em>세금계산서 세액</em></div>' +
      '<div class="tile"><span>총 정산액</span><b id="tGross">0</b><em>실지급액</em></div></div>' +
      '<div class="status" id="statusLine"></div><div class="meta" id="amtMeta"></div>' +
      '<div class="flags" id="flags"></div></section>' +
      '<section class="panel"><div class="ph"><h2>정산 대상 행</h2><p>회계팀 "매출 정리" 시트 기준입니다. 애드센스·유튜브는 쉐어 대상이 아니고, 이미 정산한 행과 판매월이 계약 종료 기준(' + esc(T.cutoff || '') + ') 이후인 행은 자동으로 뺐습니다. 체크를 바꾸면 즉시 다시 계산합니다.</p></div>' +
      '<div class="tablebox"><table id="rowTable"></table></div></section>' +
      '<details class="panel fold logic"><summary><b>적용 중인 정산 로직</b> <span class="small muted">25.4Q-26.2Q 매입품의·세금계산서, 인수인계 기준</span></summary><ol>' +
      '<li><b>대상</b>: 회계팀 음원 매출 중 애드센스·유튜브(Google) 제외, 판매월이 계약 종료 기준 월까지인 행만. 기간이 걸치는 분기 매출은 체크 후 확인합니다.</li>' +
      '<li><b>실매출</b>: 매출 합계 − 원천세 합계. <b>매입금액</b> = 실매출 × 부속합의서 쉐어율(원 단위 반올림).</li>' +
      '<li><b>부가세</b>: 매입금액 × 10% 원 미만 절사. 총 정산액 = 매입금액 + 부가세(세금계산서 금액).</li>' +
      '<li><b>중복 방지</b>: 확정 저장 시 이번에 정산한 행을 DB 원장에 기록하고, 다음 회차에는 "기정산"으로 표시합니다.</li></ol></details></div>' +

      '<div class="tabpanel" id="p-out" hidden><section class="panel">' +
      '<div class="ph"><h2>정산서 엑셀 · 세금계산서 요청</h2><p>정산서를 테이크원에 보내 세금계산서를 요청하고, 같은 파일을 품의에 첨부합니다.</p></div>' +
      '<div class="form"><label>정산 기간 라벨<input id="inLabel"></label><label>정산 기간 문구 (정산서 2번)<input id="inPeriod"></label><label>매출 기간 (정산서 표기)<input id="inSales"></label><label>참고사항 (선택)<input id="inNote" placeholder="예: 오정산 정정분 포함"></label></div>' +
      '<div class="row"><button class="btn primary big" id="btnXlsx">정산서 엑셀 내려받기</button><button class="btn" id="btnSave">이 회차 확정 저장</button><span class="small muted" id="outMsg"></span></div>' +
      '<div class="block"><header><h3>정산서 요약 (엑셀 첫 시트)</h3><button class="btn sm" data-copy="sheet">표 복사</button></header><div class="body"><div class="tablebox"><table id="sheetTable"></table></div></div></div>' +
      '<div class="block"><header><h3>세금계산서 요청 메일</h3><button class="btn sm" data-copy="mail">본문 복사</button></header><div class="body"><pre class="fmt" id="pMail"></pre></div></div>' +
      '</section></div>' +

      '<div class="tabpanel" id="p-pum" hidden><section class="panel"><div class="ph"><h2>매입품의 입력 내용</h2><p>직전 매입품의(YGK-광고IP사업팀-26-0596) 양식·문구 그대로입니다. 세금계산서를 받은 뒤 작성일자를 넣으세요.</p></div>' +
      '<div class="form"><label>거래일자 (세금계산서 작성일)<input type="date" id="inTx"></label>' +
      '<label class="chk"><span>추후 23% 재쉐어 문구</span><select id="inReshare"><option value="1">포함</option><option value="0">제외</option></select></label>' +
      '<label class="chk"><span>정산 종료 비고</span><select id="inEnd"><option value="0">넣지 않음</option><option value="1">넣음 (마지막 회차)</option></select></label>' +
      '<div class="allcopy"><button class="btn primary" data-copy="all">본문 전체 복사</button></div></div>' +
      '<div class="block"><header><h3>제목</h3><button class="btn sm" data-copy="title">복사</button></header><div class="body"><div id="pTitle"></div></div></div>' +
      '<div class="block"><header><h3>필수 입력 · 1. 기본 내역</h3><button class="btn sm" data-copy="basic">전체 복사</button></header><div class="body"><dl class="kv" id="pBasic"></dl></div></div>' +
      '<div class="block"><header><h3>2. 기타(상세내역)</h3><button class="btn sm" data-copy="detail">문구 복사</button></header><div class="body"><pre class="fmt" id="pDetail"></pre><div class="tablebox"><table id="payTable"></table></div></div></div>' +
      '<div class="block"><header><h3>잔여 예산 확인 · 매입 내역</h3><button class="btn sm" data-copy="buy">전체 복사</button></header><div class="body"><dl class="kv" id="pBuy"></dl></div></div>' +
      '</section></div>' +

      '<div class="tabpanel" id="p-erp" hidden><section class="panel"><div class="ph"><h2>ERP 매입 전표 (AP)</h2><p>품의 승인 후 입력합니다. 세금계산서 등록에서 테이크원이 발행한 승인번호를 연계하고 "부가세만 생성"으로 처리한 뒤, 거래내역에 아래 한 줄을 입력합니다.</p></div>' +
      '<div class="block"><header><h3>입력값</h3><button class="btn sm" data-copy="erpTable">표 복사</button></header><div class="body"><div class="tablebox"><table id="erpTable"></table></div></div></div></section></div>';

    function recalc() { R = TG.takeone(S.rows, T.rate || 0.15); render(); }
    function render() {
      var I = S.inputs;
      var label = I.label || R.label, salesText = I.sales || (R.from ? dotDate(R.from) + ' ~ ' + dotDate(TG.salesEnd(R.to, T.cutoff, R.from)) : '');
      var periodText = I.period || TG.periodText(R.from, R.to, T.cutoff);
      $('inLabel').placeholder = R.label; $('inSales').placeholder = salesText; $('inPeriod').placeholder = periodText;
      $('tSupply').textContent = fmt(R.supply); $('tVat').textContent = fmt(R.vat); $('tGross').textContent = fmt(R.gross);
      $('amtMeta').innerHTML = [['정산 기간', label + ' (' + R.from + ' ~ ' + R.to + ')'], ['총 매출', fmt(R.sales) + '원'], ['원천세', fmt(R.wht) + '원'], ['실매출 (원천세 제외)', fmt(R.net) + '원'], ['대상 행', R.sel.length + '건 / 전체 ' + S.rows.length + '건']]
        .map(function (x) { return '<div><span>' + x[0] + '</span>' + esc(x[1]) + '</div>'; }).join('');
      var fl = [], cnt = function (st) { return S.rows.filter(function (r) { return r.st === st; }).length; };
      var K = (CFG.KEYS || {}).TAKEONE;
      if (K && R.label === K.label) fl.push(R.supply === K.supply && R.vat === K.vat ? { level: 'ok', text: K.label + ' 실제 매입품의·세금계산서와 매입금액 ' + fmt(K.supply) + '원·부가세 ' + fmt(K.vat) + '원 일치' } : { level: 'bad', text: K.label + ' 실제 품의 금액(' + fmt(K.supply) + '원)과 불일치 — 대상 행을 확인해 주세요' });
      if (cnt('settled')) fl.push({ level: 'ok', text: '이전 회차에 정산한 행 ' + cnt('settled') + '건은 제외했습니다' });
      if (cnt('adsense')) fl.push({ level: 'ok', text: '애드센스·유튜브 행 ' + cnt('adsense') + '건은 쉐어 대상이 아니라 제외했습니다' });
      if (cnt('after')) fl.push({ level: 'warn', text: '판매월이 ' + (T.cutoff || '') + ' 이후인 행 ' + cnt('after') + '건을 제외했습니다 (계약 종료 후 판매분)' });
      var st = S.rows.filter(function (r) { return r.sel && (r.st === 'straddle' || r.st === 'unknown'); });
      if (st.length) fl.push({ level: 'warn', text: '판매 기간이 걸치거나 판단이 어려운 행 ' + st.length + '건을 포함했습니다 — 회계팀 확인 후 체크를 조정하세요' });
      if (R.net < 0) fl.push({ level: 'bad', text: '실매출이 음수입니다 — 정정(마이너스) 행이 포함됐는지 확인하세요' });
      if (!R.sel.length) fl.push(cnt('settled') && cnt('settled') + cnt('adsense') === S.rows.length ? { level: 'warn', text: '이 파일의 행은 모두 이전 회차에 정산했습니다 — 새로 정산할 행이 없습니다 (회계팀 최신 데이터인지 확인)' } : { level: 'bad', text: '정산 대상 행이 없습니다' });
      var sh = statusHtml(fl); $('statusLine').className = sh.cls; $('statusLine').innerHTML = sh.html; $('flags').innerHTML = flagsHtml(fl);
      var stLabel = { ok: '대상', straddle: '기간 걸침', unknown: '기간 불명', settled: '기정산', adsense: '애드센스', after: '종료 후' };
      $('rowTable').innerHTML = '<tr><th></th><th>회계일자</th><th>적요</th><th>거래처</th><th class="n">매출</th><th class="n">원천세</th><th>판매기간</th><th>구분</th></tr>' +
        S.rows.map(function (r, i) {
          return '<tr' + (r.sel ? '' : ' class="muted"') + '><td><input type="checkbox" data-row="' + i + '"' + (r.sel ? ' checked' : '') + (r.st === 'adsense' ? ' disabled' : '') + '></td><td>' + r.d + '</td><td>' + esc(r.s) + '</td><td>' + esc(r.v) + '</td><td class="n">' + fmt(r.a) + '</td><td class="n">' + fmt(r.w) + '</td><td>' + (r.p ? r.p.start + (r.p.end !== r.p.start ? '~' + r.p.end : '') : '-') + '</td><td>' + stLabel[r.st] + (r.tk ? ' (' + esc(r.tk) + ')' : '') + '</td></tr>';
        }).join('') + '<tr class="sum"><td></td><td colspan="3">선택 합계</td><td class="n">' + fmt(R.sales) + '</td><td class="n">' + fmt(R.wht) + '</td><td colspan="2">실매출 ' + fmt(R.net) + '</td></tr>';

      // 정산서 요약
      var rate = Math.round((T.rate || 0) * 100) + '%';
      TABLES.sheet = { head: ['매출 기간', '총 매출액', '수익배분율', '정산금액', '부가가치세', '합계'], rows: [[salesText, fmt(R.net), rate, fmt(R.supply), fmt(R.vat), fmt(R.gross)]], numCols: [1, 3, 4, 5] };
      $('sheetTable').innerHTML = tableHtml(TABLES.sheet);
      TXT.mail = (T.contact ? T.contact + ' 님, 안녕하세요.\n' : '안녕하세요.\n') + 'YG엔터테인먼트 광고IP사업팀입니다.\n\n블랙핑크더게임 OST "THE GIRLS" ' + label + ' 음원 로열티 수익 배분 정산서를 첨부드립니다.\n' +
        '- 매출 기간 : ' + salesText + '\n- 총 매출액(원천세 제외) : ' + fmt(R.net) + '원\n- 수익배분율 : ' + rate + '\n- 정산금액 : ' + fmt(R.supply) + '원 (VAT 별도 ' + fmt(R.vat) + '원 / 합계 ' + fmt(R.gross) + '원)\n\n' +
        '첨부 정산서 기준으로 세금계산서 발행 부탁드립니다.\n(품목: BPTG \'THE GIRLS\' (' + label + ') 음원 매출, 공급가액 ' + fmt(R.supply) + '원, 세액 ' + fmt(R.vat) + '원)\n\n감사합니다.';
      $('pMail').textContent = TXT.mail;

      // 품의
      // 거래일자 기본값: 대상 분기 종료 다음 달 말일(26.2Q → 7/31, 직전 세금계산서 작성일 기준). 세금계산서 작성일이 다르면 수정
      var tx = I.tx || (R.to ? TG.monthEnd(TG.addMonths(R.to.slice(0, 7), 1)) : ''), pay = tx ? TG.monthEnd(TG.addMonths(tx.slice(0, 7), 1)) : '';
      if ($('inTx').value !== tx) $('inTx').value = tx;
      // 23% 재쉐어 문구 기본값: 정산 기간이 게임 서비스 종료일 이전이면 포함, 이후 기간이 섞이면 제외
      var reshare = I.reshare != null ? I.reshare : (!T.serviceEnd || R.to <= T.serviceEnd);
      $('inReshare').value = reshare ? '1' : '0';
      TXT.title = "[매입품의] " + label + " 블랙핑크더게임 OST 'THE GIRLS' 음원 로열티 수익 배분 (테이크원컴퍼니)";
      $('pTitle').textContent = TXT.title;
      var basic = [['비거주자 여부', '거주자'], ['외국환 신고 필요여부', '1.경상거래 해당없음 / 2.자본거래 해당없음'], ['계약서 작성 여부', '작성'], ['지급방법', '송금'],
        ['계약기간', T.contract || ''], ['총금액 (VAT제외)', fmt(R.supply)], ['기간', R.from + ' ~ ' + R.to], ['통화 종류', 'KRW']];
      $('pBasic').innerHTML = kv(basic, 'basic'); TXT.basic = basic.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
      var re = Math.round(R.gross * (T.reshareRate || 0.23));
      TXT.detail = '1. 내용 : 블랙핑크더게임 OST "THE GIRLS" ' + R.rangeText.replace(' ~ ', '~') + ' 음원 로열티 정산(' + (T.detailName || T.name || '') + ')\n\n2. 기간 : ' + R.rangeText + '\n' +
        '※ BPTG 계약기간 종료로 인해 해당 정산 이후 판매월 기준 ' + (T.cutoffText || '') + ' 내 데이터만 추가 정산 예정\n\n' +
        '3. 매입금액(VAT제외, 원천징수전) : ' + fmt(R.supply) + ' 원\n\n4. 지급예정일(카드번호) : 품의 결재 후 자금집행일\n\n5. 상세내역 :\n' +
        '- 총 매출 : ' + fmt(R.net) + '원(원천세 제외)\n- 총 정산액 : ' + fmt(R.gross) + '원 (VAT 포함)\n- YG X 테이크원 체결 부속합의서 근거 음원 순매출 ' + rate + ' 쉐어' +
        (reshare ? '\n- 해당 음원 매출 지급분은 게임 "총 매출"에 포함 예정으로, 추후 테이크원 → YG 분기별 정산 시 재차 ' + Math.round((T.reshareRate || 0.23) * 100) + '% 매출 YG에 쉐어 예정 * 본 정산액 ' + fmt(R.gross) + '원의 ' + Math.round((T.reshareRate || 0.23) * 100) + '%인 약 ' + (re >= 10000 ? Math.round(re / 10000) + '만원' : fmt(re) + '원') + ' 예상' : '') +
        (I.endNote ? '\n※ 본 정산을 끝으로 테이크원 측에 배분되는 THE GIRLS 음원 로열티 수익은 없습니다.' : '') +
        (I.note ? '\n※ ' + I.note : '') +
        '\n\n6. 첨부파일 :\n- 음원 매출 데이터 1부.\n- 세금계산서 1부.\n\n7. 지급처 정보(원화)';
      $('pDetail').textContent = TXT.detail;
      TABLES.pay = { head: ['지급처', '실지급액', '발행일자', '은행', '계좌번호', '예금주'], rows: [[T.name || '', fmt(R.gross), dotDate(tx) || '입력 필요', T.bank || '', T.account || '', T.holder || '']], numCols: [1] };
      $('payTable').innerHTML = tableHtml(TABLES.pay);
      var buy = PRJ(P).slice(0, 3).concat([['회계일자', tx || '입력 필요'], ['총금액 (VAT제외)', fmt(R.supply)], ['거래처명', (T.name || '') + ' (' + (T.code || '') + ')'], ['거래일자', tx || '입력 필요 — 세금계산서 작성일'],
        ['금액', fmt(R.supply)], ['VAT', fmt(R.vat)], ['총 금액', fmt(R.gross)], ['지급예정일(이내)', pay || '거래일자 입력 시 자동'], ['아티스트(그룹)', P.group + ' (' + P.groupCode + ')'], ['아티스트(개인)', P.artist + ' (' + P.artistCode + ')']]);
      $('pBuy').innerHTML = kv(buy, 'buy'); TXT.buy = buy.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
      var tsv = function (t) { return [t.head].concat(t.rows).map(function (r) { return r.join('\t'); }).join('\n'); };
      TXT.all = [TXT.title, '', TXT.basic, '', '2. 기타(상세내역)', TXT.detail, tsv(TABLES.pay), '', TXT.buy].join('\n');

      // ERP
      var erp = [['AP 유형', '미지급금 일반'], ['증빙', '전자세금계산서 (승인번호 연계 · 부가세만 생성)'], ['거래처', (T.name || '') + ' (' + (T.code || '') + ')'], ['회계일자', tx || '입력 필요'], ['지급예정일', pay || '-'], ['통화/환율', 'KRW / 1'],
        ['계정코드', P.account + ' (' + P.accountCode + ')'], ['차변 금액', fmt(R.supply)], ['부가세대급금', fmt(R.vat) + ' (세금계산서 연계로 자동)'], ['미지급금 합계', fmt(R.gross)]].concat(PRJ(P).filter(function (x) { return x[0] !== '계정'; }))
        .concat([['적요', label + " 블랙핑크더게임 OST 'THE GIRLS' 음원 로열티 수익 배분 (테이크원컴퍼니)"], ['기안문서', '위 매입품의 연동']]);
      TABLES.erpTable = { head: ['항목', '값'], rows: erp, numCols: [] };
      $('erpTable').innerHTML = tableHtml(TABLES.erpTable);
      $('btnSave').disabled = !R.sel.length;
    }

    ctx.root.addEventListener('change', function (e) {
      var t = e.target;
      if (t.hasAttribute('data-row')) { S.rows[+t.getAttribute('data-row')].sel = t.checked; recalc(); }
      if (t.id === 'inTx') { S.inputs.tx = t.value; render(); }
      if (t.id === 'inReshare') { S.inputs.reshare = t.value === '1'; render(); }
      if (t.id === 'inEnd') { S.inputs.endNote = t.value === '1'; render(); }
    });
    ctx.root.addEventListener('input', function (e) {
      var id = e.target.id;
      if (id === 'inLabel') S.inputs.label = e.target.value.trim();
      else if (id === 'inSales') S.inputs.sales = e.target.value.trim();
      else if (id === 'inPeriod') S.inputs.period = e.target.value.trim();
      else if (id === 'inNote') S.inputs.note = e.target.value.trim();
      else return;
      clearTimeout(render._t); render._t = setTimeout(render, 200);
    });

    $('btnXlsx').addEventListener('click', function () {
      var I = S.inputs, label = I.label || R.label, salesText = I.sales || $('inSales').placeholder, rate = (T.rate || 0.15);
      var periodText = I.period || $('inPeriod').placeholder, written = $('inTx').value || today();
      var aoa = [[], ['블랙핑크더게임 OST "THE GIRLS" 음원 매출 수익 배분 건'], [], ['1) 정 산 대 상 : ', null, T.name || ''], ['2) 정 산 기 간 : ', null, periodText],
        ['3) 정 산 금 액 : ', null, fmt(R.gross) + '원 (VAT 포함)'], ['4) 정산 세부내역'], [], ['매출 기간', null, '총 매출액', '수익배분율', '정산금액', '부가가치세', '합계'],
        [salesText, null, R.net, rate, R.supply, R.vat, R.gross], ['합계', null, R.net, null, R.supply, R.vat, R.gross], [],
        ['5) 정산금액 지급일 :', null, '품의 결재 후 자금집행일'], ['6) 참고사항', null, I.note || ''], [], [null, null, '작 성 일 자', written], [null, null, '작 성 자', CFG.AUTHOR || '광고IP사업팀']];
      var wb = XLSX.utils.book_new();
      var s1 = XLSX.utils.aoa_to_sheet(aoa); s1['!cols'] = [{ wch: 22 }, { wch: 2 }, { wch: 16 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 14 }];
      XLSX.utils.book_append_sheet(wb, s1, 'Sheet1');
      var rows = [['회계일자', '적요', '매출', '거래처', '원천세']].concat(R.sel.map(function (r) { return [r.d, r.s, r.a, r.v, r.w || null]; }));
      rows.push([], [null, null, R.sales, '< 매출', R.wht], [null, null, R.wht, '< 원천세'], [null, null, R.net, '< 실매출'], [], [null, null, '테이크원 음원 매출 쉐어(' + Math.round(rate * 100) + '%)'], [null, null, R.supply]);
      var s2 = XLSX.utils.aoa_to_sheet(rows); s2['!cols'] = [{ wch: 12 }, { wch: 70 }, { wch: 14 }, { wch: 34 }, { wch: 12 }];
      XLSX.utils.book_append_sheet(wb, s2, '매출 정리');
      if (parsed.ledgerSheet) XLSX.utils.book_append_sheet(wb, parsed.ledgerSheet, '원장');
      ctx.saveFile(label + ' 블랙핑크더게임 THE GIRLS 음원 매출 데이터_FINAL_' + written.slice(2).replace(/-/g, '') + '.xlsx', XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
      ctx.toast('정산서 엑셀을 내려받았습니다');
    });

    $('btnSave').addEventListener('click', async function () {
      var label = S.inputs.label || R.label;
      if (R.net < 0 || !R.sel.length) { $('outMsg').textContent = '조치 항목이 남아 있어 저장하지 않았습니다.'; return; }
      var items = [];
      S.rows.forEach(function (r) {
        if (r.st === 'adsense' || r.st === 'settled') return;
        items.push({ name: r.key, member: JSON.stringify({ key: r.key, d: r.d, s: r.s, a: r.a, v: r.v, w: r.w, tk: r.sel ? label : '' }) });
      });
      try {
        $('outMsg').textContent = '저장 중…';
        await ctx.post('saveMapping', { items: items });
        await ctx.post('saveQuarter', { record: { quarter: 'TK ' + label, supply: R.supply, vat: R.vat, gross: R.gross, txDate: S.inputs.tx, file: ctx.fileName,
          lines: R.sel.map(function (r) { return { key: r.key, amt: r.a, wht: r.w }; }), royaltyTable: [{ net: R.net, rate: T.rate }] } });
        S.rows.forEach(function (r) { if (r.sel) { r.st = 'settled'; r.tk = label; r.sel = false; } });
        $('outMsg').textContent = ''; ctx.toast(label + ' 확정 저장했습니다 (대상 행은 다음 회차에서 기정산으로 표시)'); recalc();
      } catch (e) { $('outMsg').textContent = '저장하지 못했습니다 (' + e.message + ')'; }
    });

    wire(ctx, TXT, TABLES, ['calc', 'out', 'pum', 'erp'], 'tg.takeone.tab');
    recalc();
    return { quarter: R.label, note: parsed.sheet + ' · ' + parsed.rows.length + '행' };
  }

  // ================= ② 해외 작가 (반기) =================
  function mountWriter(ctx) {
    var $ = function (id) { return document.getElementById(id); };
    var D = cfgOf(ctx.data), CFG = D.CFG, P = CFG.PROJECT || {}, W = CFG.WRITERS || [], WC = CFG.WRITER || {};
    var parsed = TG.parse(ctx.wb);
    // DB 누적 원장 + 이번 파일의 행(같은 키는 이번 파일 값)
    var all = {}; Object.keys(D.ledger).forEach(function (k) { all[k] = D.ledger[k]; }); parsed.rows.forEach(function (r) { all[r.key] = r; });
    var rows = Object.keys(all).map(function (k) { return all[k]; });
    var maxD = parsed.rows.map(function (r) { return r.d; }).sort().pop();
    var S = { y: +maxD.slice(0, 4), h: +maxD.slice(5, 7) <= 6 ? 1 : 2, fx: '', fxNote: '', merchSales: '', merchNet: '', merchNote: '', tx: '' };
    function merchDefault() {
      var key = S.y + 'H' + S.h, m = (CFG.MERCH || {})[key], f = (CFG.FX || {})[key];
      S.merchSales = m ? m.sales : ''; S.merchNet = m ? m.net : ''; S.merchNote = m ? 'DB 기록값 (' + key + ')' : '';
      S.fx = f || ''; S.fxNote = f ? 'DB 확정값' : '직접 입력 필요 (서울외국환중개 반기 말일 매매기준율)';
      S.tx = S.h === 1 ? S.y + '-08-01' : (S.y + 1) + '-01-31';   // 직전 회차 인보이스·ERP 회계일자 기준(1H → 8/1, 2H → 익년 1/31)
    }
    var TXT = {}, TABLES = {}, R = null;

    ctx.tabsEl.innerHTML = tabsHtml([['calc', '정산'], ['out', '리포트·메일'], ['pum', '품의 복붙'], ['erp', 'ERP 사업소득']]);
    ctx.root.innerHTML =
      '<div class="tabpanel" id="p-calc"><section class="panel">' +
      '<div class="form"><label>정산 반기<select id="inHalf"></select></label>' +
      '<label>환율 (반기 말일, KRW/USD)<input id="inFx" inputmode="decimal" placeholder="예: 1,400.00"></label>' +
      '<label>OST MERCH 총매출 (KRW)<input id="inMs" inputmode="numeric"></label>' +
      '<label>OST MERCH 제조원가 제외 매출 (KRW)<input id="inMn" inputmode="numeric"></label></div>' +
      '<div class="row"><label class="btn sm" for="mdFile">MD 관련 수불부 올리기 (OST MERCH 자동 입력)</label><input type="file" id="mdFile" accept=".xlsx,.xls" hidden><span class="small muted" id="fxNote"></span><span class="small muted" id="mdNote"></span></div>' +
      '<p class="small muted" style="margin:0">환율: <a href="http://www.smbs.biz/ExRate/StdExRate.jsp" target="_blank" rel="noopener">서울외국환중개 매매기준율</a> 반기 말일(6/30·12/31) 기준. OST MERCH: 테이크원 MD 관련 수불부 "OST 패키지" 월별 금액의 반기 합계(매출추이 시트 OST MERCH 총이익 행과 동일).</p></section>' +
      '<section class="panel dash"><div class="tiles" id="tiles"></div><div class="status" id="statusLine"></div><div class="meta" id="amtMeta"></div><div class="flags" id="flags"></div></section>' +
      '<section class="panel"><div class="ph"><h2>음원 매출 대상 행</h2><p>회계일자가 정산 반기 안인 행(애드센스·유튜브 제외)입니다. DB 누적 원장과 이번 파일을 합쳐서 봅니다.</p></div><div class="tablebox"><table id="rowTable"></table></div></section>' +
      '<details class="panel fold logic"><summary><b>적용 중인 정산 로직</b> <span class="small muted">26년 상반기 작가 배분 매입품의·인보이스·ERP 화면 기준</span></summary><ol>' +
      '<li><b>음원</b>: 반기 회계일자 행의 (매출 − 원천세) × 작가별 요율(원 단위 반올림).</li>' +
      '<li><b>OST MERCH</b>: 제조원가 제외 매출 × YG 수익배분율(원 단위 반올림) × 작가별 요율.</li>' +
      '<li><b>USD</b>: 반기 말일 매매기준율로 구분별 환산(센트 반올림) 후 합산.</li>' +
      '<li><b>원천세</b>: 법인세 = 지급액 × 10% (10원 미만 절사), 주민세 = 법인세 × 10% (10원 미만 절사). 실지급액 = 지급액 − 원천세.</li></ol></details></div>' +

      '<div class="tabpanel" id="p-out" hidden><section class="panel"><div class="ph"><h2>작가별 Royalty Report · 인보이스 요청</h2><p>리포트를 보내 인보이스를 요청하고, 받은 인보이스와 함께 품의에 첨부합니다.</p></div>' +
      '<div class="row" id="dlRow"></div><div id="mails"></div>' +
      '<div class="row"><button class="btn" id="btnSave">이 반기 확정 저장</button><span class="small muted" id="outMsg"></span></div></section></div>' +

      '<div class="tabpanel" id="p-pum" hidden><section class="panel"><div class="ph"><h2>매입품의 입력 내용</h2><p>직전 매입품의(YGK-광고IP사업팀-26-0600) 양식·문구 그대로입니다. 거래일자는 인보이스 작성일입니다.</p></div>' +
      '<div class="form"><label>거래일자 (인보이스 작성일)<input type="date" id="inTx"></label><div class="allcopy"><button class="btn primary" data-copy="all">본문 전체 복사</button></div></div>' +
      '<div class="block"><header><h3>제목</h3><button class="btn sm" data-copy="title">복사</button></header><div class="body"><div id="pTitle"></div></div></div>' +
      '<div class="block"><header><h3>필수 입력 · 1. 기본 내역</h3><button class="btn sm" data-copy="basic">전체 복사</button></header><div class="body"><dl class="kv" id="pBasic"></dl></div></div>' +
      '<div class="block"><header><h3>2. 기타(상세내역)</h3><button class="btn sm" data-copy="detail">문구 복사</button></header><div class="body"><pre class="fmt" id="pDetail"></pre><div id="wTables"></div></div></div>' +
      '<div class="block"><header><h3>잔여 예산 확인 · 매입 내역 (작가별)</h3><button class="btn sm" data-copy="buy">전체 복사</button></header><div class="body"><pre class="fmt" id="pBuy"></pre></div></div>' +
      '</section></div>' +

      '<div class="tabpanel" id="p-erp" hidden><section class="panel"><div class="ph"><h2>ERP 사업기타소득 · 사업소득관리</h2><p>매입 처리가 아니라 사업소득관리 화면에서 작가별로 1건씩 입력합니다(기안문서 선택 → 저장 → 전표발행 → 결재 생성).</p></div>' +
      '<div class="block"><header><h3>작가별 입력값</h3><button class="btn sm" data-copy="erpTable">표 복사</button></header><div class="body"><div class="tablebox"><table id="erpTable"></table></div></div></div></section></div>';

    var halves = []; for (var y = S.y + 1; y >= 2024; y--) { halves.push([y, 2], [y, 1]); }
    $('inHalf').innerHTML = halves.map(function (x) { var h = TG.half(x[0], x[1]); return '<option value="' + x[0] + '-' + x[1] + '"' + (x[0] === S.y && x[1] === S.h ? ' selected' : '') + '>' + h.kor + ' (' + h.from + ' ~ ' + h.to + ')</option>'; }).join('');
    function syncInputs() {
      $('inFx').value = S.fx; $('inFx').classList.toggle('need', !S.fx); $('inMs').value = S.merchSales; $('inMn').value = S.merchNet; $('inTx').value = S.tx;
      $('fxNote').textContent = S.fxNote ? '환율: ' + S.fxNote : ''; $('mdNote').textContent = S.merchNote ? ' · OST MERCH: ' + S.merchNote : ' · OST MERCH: 테이크원 MD 수불부(반기 마지막 달 파일)를 올리면 자동 입력';
    }
    merchDefault(); syncInputs();

    function recalc() {
      var fx = parseFloat(String(S.fx).replace(/,/g, '')) || 0, hf = TG.half(S.y, S.h);
      R = TG.writer(rows, hf, { sales: +String(S.merchSales).replace(/,/g, '') || 0, net: +String(S.merchNet).replace(/,/g, '') || 0 }, fx, W, CFG.MERCH_SHARE || 0.23);
      render();
    }
    function render() {
      var hf = R.hf, fx = R.fx;
      $('tiles').innerHTML = R.writers.map(function (w) { return '<div class="tile"><span>' + esc(w.key) + ' (' + Math.round(w.rate * 100) + '%)</span><b>' + (fx ? 'USD ' + fmt2(w.totalU) : fmt(w.totalK) + '원') + '</b><em>' + fmt(w.totalK) + '원 · 원천세 ' + fmt(w.wht) + ' · 실지급 ' + fmt(w.pay) + '</em></div>'; }).join('') +
        '<div class="tile main"><span>총 정산액 (원천세 제외 전)</span><b>' + (fx ? 'USD ' + fmt2(R.totalU) : fmt(R.totalK) + '원') + '</b><em>' + fmt(R.totalK) + '원 · 품의 "매입금액"</em></div>';
      $('amtMeta').innerHTML = [['정산 기간', hf.from + ' ~ ' + hf.to + ' (' + hf.label + ')'], ['음원 총매출 / 원천세', fmt(R.sales) + ' / ' + fmt(R.wht) + '원'], ['YG 음원 순매출', fmt(R.net) + '원'], ['OST MERCH YG 순매출', fmt(R.merchYG) + '원 (' + fmt(R.merch.net) + ' × ' + Math.round((CFG.MERCH_SHARE || 0.23) * 100) + '%)'], ['총 매출 (USD)', fx ? fmt2(R.salesU) : '환율 입력 필요']]
        .map(function (x) { return '<div><span>' + x[0] + '</span>' + esc(x[1]) + '</div>'; }).join('');
      var fl = [], K = ((CFG.KEYS || {}).WRITER || {})[hf.y + 'H' + hf.h];
      if (K) { var ok = R.writers.every(function (w) { var k = K[w.key]; return k && k.krw === w.totalK && Math.abs(k.usd - w.totalU) < 0.005 && (k.wht == null || k.wht === w.wht); });
        fl.push(ok ? { level: 'ok', text: hf.kor + ' 실제 품의·인보이스와 작가별 KRW·USD·원천세 모두 일치' } : { level: 'bad', text: hf.kor + ' 실제 품의 금액과 불일치 — 환율·OST MERCH·대상 행을 확인해 주세요' }); }
      if (!fx) fl.push({ level: 'bad', text: '반기 말일 환율을 입력해 주세요 — 서울외국환중개 매매기준율 (' + TG.half(S.y, S.h).to + ' 기준)' });
      else if (S.fxNote) fl.push({ level: 'ok', text: '환율 USD ' + fmt2(fx) + ' — ' + S.fxNote });
      var ex = S.merchExtras || ((CFG.MERCH || {})[hf.y + 'H' + hf.h] || {}).extras || [];
      if (ex.length) fl.push({ level: 'ok', text: '리유저블백·하트쿠션은 작가 OST MERCH에서 제외 (' + ex.map(function (x) { return x.item + ' 총이익 ' + fmt(x.net); }).join(', ') + ') — 24.2H 매입품의 기준: 두 품목 판매가 있었어도 OST 패키지만 반영' });
      if (R.merch.net && S.merchNote) fl.push({ level: 'ok', text: 'OST MERCH 제조원가 제외 매출 ' + fmt(R.merch.net) + '원 — ' + S.merchNote });
      if (!R.merch.net) fl.push({ level: 'warn', text: 'OST MERCH 금액이 비어 있습니다 — 정산 탭에서 테이크원 MD 관련 수불부(반기 마지막 달)를 올리면 자동으로 채워집니다' });
      if (!R.sel.length) fl.push({ level: 'bad', text: '반기 안의 음원 매출 행이 없습니다 — 회계팀 최신 데이터를 올렸는지 확인하세요' });
      var lastD = R.sel.length ? R.sel[R.sel.length - 1].d : '';
      if (R.sel.length && lastD < TG.addMonths(hf.to.slice(0, 7), 0) + '-01') fl.push({ level: 'warn', text: '반기 마지막 달 회계 데이터가 아직 없습니다 (최종 회계일자 ' + lastD + ') — 반기 결산 데이터 수령 후 확정하세요' });
      fl.push({ level: 'ok', text: '작가 쉐어 합계 ' + Math.round(R.rateSum * 100) + '% (' + R.writers.map(function (w) { return w.key + ' ' + Math.round(w.rate * 100) + '%'; }).join(' · ') + ')' });
      var sh = statusHtml(fl); $('statusLine').className = sh.cls; $('statusLine').innerHTML = sh.html; $('flags').innerHTML = flagsHtml(fl);
      $('rowTable').innerHTML = '<tr><th>회계일자</th><th>적요</th><th>거래처</th><th class="n">매출</th><th class="n">원천세</th></tr>' +
        R.sel.map(function (r) { return '<tr><td>' + r.d + '</td><td>' + esc(r.s) + '</td><td>' + esc(r.v) + '</td><td class="n">' + fmt(r.a) + '</td><td class="n">' + fmt(r.w) + '</td></tr>'; }).join('') +
        '<tr class="sum"><td colspan="3">합계 (YG 순매출 ' + fmt(R.net) + ')</td><td class="n">' + fmt(R.sales) + '</td><td class="n">' + fmt(R.wht) + '</td></tr>';

      // 리포트·메일
      $('dlRow').innerHTML = R.writers.map(function (w, i) { return '<button class="btn primary" data-dl="' + i + '">' + esc(w.key) + ' Royalty Report 내려받기</button>'; }).join('');
      var fxDate = hf.to, fxTxt = fx ? fmt2(fx) : '(환율)';
      $('mails').innerHTML = R.writers.map(function (w, i) {
        TXT['mail' + i] = 'Dear ' + (w.contact || w.name) + ',\n\nI hope this email finds you well. This is ' + (CFG.AUTHOR_EN || 'YG Entertainment, Advertising IP Business Team') + '.\n\n' +
          'Please find attached the ' + hf.eng + ' royalty report for BLACKPINK THE GAME OST "THE GIRLS" (' + hf.from + ' ~ ' + hf.to + ').\n' +
          '- THE GIRLS OST (DSP) : KRW ' + fmt(w.ostK) + '\n- OST MERCH : KRW ' + fmt(w.merchK) + '\n- TOTAL : KRW ' + fmt(w.totalK) + ' → USD ' + fmt2(w.totalU) + ' (KRW/USD ' + fxTxt + ', as of ' + fxDate + ')\n\n' +
          'Could you kindly issue an invoice for USD ' + fmt2(w.totalU) + ' addressed to YG Entertainment Inc.?\nPlease note that Korean withholding tax will be deducted at payment in accordance with the applicable tax treaty.\n\nThank you.\nBest regards,';
        return '<div class="block"><header><h3>' + esc(w.key) + ' 인보이스 요청 메일 (영문)</h3><button class="btn sm" data-copy="mail' + i + '">본문 복사</button></header><div class="body"><pre class="fmt">' + esc(TXT['mail' + i]) + '</pre></div></div>';
      }).join('');

      // 품의
      var tx = S.tx, pay = tx ? TG.monthEnd(TG.addMonths(tx.slice(0, 7), 1)) : '';
      TXT.title = "[매입품의] 블랙핑크더게임 OST 'THE GIRLS' " + hf.kor + " 음원 로열티 수익 작가 배분 건";
      $('pTitle').textContent = TXT.title;
      var basic = [['비거주자 여부', '비거주자'], ['외국환 신고 필요여부', '직전 품의 선택값과 동일'], ['계약서 작성 여부', '작성'], ['지급방법', '송금'], ['전체계약금액 (VAT제외)', fx ? fmt2(R.totalU) : '-'], ['계약기간', WC.contract || ''], ['통화 종류', 'USD']];
      $('pBasic').innerHTML = kv(basic, 'basic'); TXT.basic = basic.map(function (x) { return x[0] + '\t' + x[1]; }).join('\n');
      var order = R.writers.slice().sort(function (a, b) { return b.rate - a.rate; });
      var names = order.map(function (w) { return w.detailName || w.name; }).join(', ');
      TXT.detail = '아래와 같이 품의하오니 재가하여 주시기 바랍니다.\n\n- 아 래 -\n\n1. 내용 : 블랙핑크더게임 OST "THE GIRLS" 발매 후 ' + hf.kor + ' 음원 로열티 정산(' + names + ')\n\n' +
        '2. 기간 : ' + hf.from.slice(2).replace(/-/g, '/') + ' ~ ' + hf.to.slice(2).replace(/-/g, '/') + ' (' + hf.label + ')\n\n' +
        '3. 매입금액(VAT제외, 원천징수전) : ' + fmt2(R.totalU) + ' USD\n\n4. 상세내역 :\n\n- 총 매출 : ' + fmt2(R.salesU) + ' USD (원천세 제외)\n\n- 총 정산액 : ' + fmt2(R.totalU) + ' USD (원천세 제외)\n\n' +
        R.writers.map(function (w) { return w.short + ' (' + Math.round(w.rate * 100) + '%) : ' + fmt2(w.totalU) + ' USD (원천세 제외)'; }).join(' ') + '\n\n' +
        '※ ' + hf.to.replace(/-/g, '.') + ' 기준 환율 USD ' + fxTxt + '(서울외국환중개소) 적용 ※ 정산액 환율 적용 기준 매 반기 마지막 일자 (6월 말일/12월 말일) ※ YG음원 매출 중 ' + Math.round(R.rateSum * 100) + '% 작가진 쉐어 : ' + R.writers.map(function (w) { return w.key + ' ' + Math.round(w.rate * 100) + '%'; }).join(', ') + ' (반기별 지급)\n\n' +
        '5. 첨부파일 :\n\n- 음원 및 OST MERCH 매출 데이터 ' + R.writers.length + '부.\n- 인보이스 ' + R.writers.length + '부.\n\n6. 지급처 정보(USD)';
      $('pDetail').textContent = TXT.detail;
      var wt = '';
      order.forEach(function (w, i) {
        var pct = Math.round(w.rate * 100) + '%';
        TABLES['w' + i] = { head: ['구분', 'YG 매출 (KRW)', '정산비율', w.key + ' (KRW)', w.key + ' (USD)'], rows: [['OST MERCH', fmt(R.merchYG), pct, fmt(w.merchK), fmt2(w.merchU)], ['OST 음원', fmt(R.net), pct, fmt(w.ostK), fmt2(w.ostU)], ['TOTAL (KRW)', fmt(R.merchYG + R.net), pct, fmt(w.totalK), fmt2(w.totalU)]], numCols: [1, 3, 4], sumLast: true };
        TABLES['p' + i] = { head: ['지급처', '실지급액', '은행', '계좌번호', 'ROUTING#', 'SWIFT CODE'], rows: [[w.name, 'USD ' + fmt2(w.totalU), w.bank || '', w.account || '', w.routing || '', w.swift || '']], numCols: [] };
        wt += '<div class="row" style="justify-content:space-between"><b class="small">' + esc(w.key) + ' 배분표</b><button class="btn sm" data-copy="w' + i + '">표 복사</button></div><div class="tablebox"><table>' + tableHtml(TABLES['w' + i]) + '</table></div>' +
          '<div class="row" style="justify-content:space-between"><b class="small">' + esc(w.key) + ' 지급처</b><button class="btn sm" data-copy="p' + i + '">표 복사</button></div><div class="tablebox"><table>' + tableHtml(TABLES['p' + i]) + '</table></div>';
      });
      $('wTables').innerHTML = wt;
      TXT.buy = order.map(function (w, i) {
        return '■ 잔여 예산 확인 ' + (i + 1) + '-1\n1) 회계부서 ' + P.deptName + '(' + P.dept + ')  2) 프로젝트 ' + P.proj + '(' + P.projCode + ')  3) 계정 ' + P.account + '(' + P.accountCode + ')  4) 회계일자 ' + (tx || '입력 필요') + '\n총금액(VAT제외) ' + fmt(w.totalK) + '\n' +
          '■ 매입 내역 ' + (i + 1) + '\n거래처명\t' + w.erpName + ' (' + w.code + ')\n거래일자\t' + (tx || '입력 필요 — 인보이스 작성일') + '\n금액\t' + fmt(w.totalK) + '\n원천세\t' + fmt(w.wht) + '\n총 금액\t' + fmt(w.pay) +
          '\n프로젝트\t' + P.proj + ' (' + P.projCode + ')\n지급예정일(이내)\t' + (pay || '거래일자 입력 시 자동') + '\n아티스트(그룹)\t' + P.group + ' (' + P.groupCode + ')\n아티스트(개인)\t' + P.artist + ' (' + P.artistCode + ')';
      }).join('\n\n');
      $('pBuy').textContent = TXT.buy;
      var tsv = function (t) { return [t.head].concat(t.rows).map(function (r) { return r.join('\t'); }).join('\n'); };
      TXT.all = [TXT.title, '', TXT.basic, '', '2. 기타(상세내역)', TXT.detail].concat(order.map(function (w, i) { return tsv(TABLES['p' + i]); })).concat(order.map(function (w, i) { return tsv(TABLES['w' + i]); })).concat(['', TXT.buy]).join('\n');

      // ERP 사업소득관리
      var head = ['항목'].concat(order.map(function (w) { return w.key; }));
      var rowsE = [['매입처 (소득자)', function (w) { return w.erpName + ' (' + w.code + ')'; }], ['비거주자 국가', function (w) { return w.country || 'US'; }], ['은행 / SWIFT', function (w) { return '999 기타 / ' + (w.swift || ''); }], ['계좌번호', function (w) { return w.account || ''; }],
        ['지급예정일', function () { return pay || '-'; }], ['기안문서', function () { return '위 매입품의 선택'; }], ['소득구분', function () { return 'F61 사용료소득'; }], ['세율', function () { return '0.100'; }], ['회사부담', function () { return '아니오'; }],
        ['회계일자', function () { return tx || '입력 필요'; }], ['통화 / 환율', function () { return 'USD / ' + (fx ? fmt2(fx) : '-'); }], ['지급액 = 소득금액 (KRW)', function (w) { return fmt(w.totalK); }], ['법인세', function (w) { return fmt(w.corp); }], ['주민세', function (w) { return fmt(w.local); }],
        ['실지급액 (KRW)', function (w) { return fmt(w.pay); }], ['외화AP (USD)', function (w) { return fx ? fmt2(w.payU) : '-'; }], ['예술인고용보험', function () { return '미적용'; }], ['세금처리', function () { return '지급처리시 원천세공제 · 예수금만 생성'; }],
        ['회계부서', function () { return P.dept + ' ' + P.deptName; }], ['프로젝트', function () { return P.projCode + ' ' + P.proj; }], ['계정코드', function () { return P.accountCode + ' ' + P.account; }], ['아티스트그룹 / 아티스트', function () { return P.groupCode + ' ' + P.group + ' / ' + P.artistCode + ' ' + P.artist; }],
        ['제작구분', function () { return P.prod; }], ['적요', function () { return P.proj + ' 매출 해외작가 정산금 (' + hf.label + ')'; }]];
      TABLES.erpTable = { head: head, rows: rowsE.map(function (r) { return [r[0]].concat(order.map(function (w) { return r[1](w); })); }), numCols: [] };
      $('erpTable').innerHTML = tableHtml(TABLES.erpTable);
      $('btnSave').disabled = !fx || !R.sel.length;
    }

    function reportWb(w) {
      var hf = R.hf, fx = R.fx, pct = Math.round(w.rate * 100) + '%';
      var wb = XLSX.utils.book_new();
      var a = [[null, 'BPTG OST(THE GIRLS) MERCH ROYALTY SHARE REVENUE', null, null, null, null, null, 'exchange rate date', hf.to.slice(2).replace(/-/g, '.')],
        [null, null, null, null, null, null, null, 'KRW > USD', fx], [], [null, 'CATEGORY', 'THE GIRLS OST', 'OST MERCH', 'TOTAL (KRW)', 'TOTAL (USD)'],
        [null, 'ROYALTY (' + pct + ')', w.ostK, w.merchK, w.totalK, w.totalU], [], [null, 'PERIOD: ' + hf.eng + ' (' + hf.from + ' ~ ' + hf.to + ')'], [null, 'ARRANGEMENT: OST sales + OST merch sales']];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(a), 'Sales Review');
      var b = [['DSP Sales Report', null, 'Revenue Period: ' + hf.from + ' ~ ' + hf.to], [], ['Total Sales (KRW)', R.sales], ['Withholding Tax (KRW)', R.wht], ['YG Net Receipt (KRW)', R.net], ['R/S(' + pct + ') KRW', w.ostK], [],
        ['Accounting Date', 'SUMMARY', 'Revenue', 'Distributor', 'Withholding tax']].concat(R.sel.map(function (r) { return [r.d, r.s, r.a, r.v, r.w || null]; }))
        .concat([[], [null, null, R.sales, '< Total Sales', R.wht], [null, null, R.wht, '< Withholding Tax'], [null, null, R.net, '< YG Receipts'], [null, null, 'DSP royalty Share(' + pct + ')'], [null, null, w.ostK]]);
      var sb = XLSX.utils.aoa_to_sheet(b); sb['!cols'] = [{ wch: 20 }, { wch: 70 }, { wch: 14 }, { wch: 34 }, { wch: 14 }];
      XLSX.utils.book_append_sheet(wb, sb, 'DSP Sales Report');
      var c = [[null, 'BPTG OST (THE GIRLS) MERCH ROYALTY SHARE REVENUE'], [null, 'CATEGORY', 'TOTAL SALES REVENUE', 'REVENUE EXCLUDING MANUFACTURING COSTS', 'NET SALES FOR YG', 'ROYALTY SHARE(' + pct + ')'],
        [null, 'OST MERCH', R.merch.sales, R.merch.net, R.merchYG, w.merchK]];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(c), 'Merchandise Sales Report');
      return wb;
    }

    ctx.root.addEventListener('click', function (e) {
      var b = e.target.closest('[data-dl]'); if (!b) return;
      var w = R.writers[+b.getAttribute('data-dl')];
      if (!R.fx) { ctx.toast('환율을 먼저 입력해 주세요'); return; }
      ctx.saveFile('[' + w.key.charAt(0) + w.key.slice(1).toLowerCase() + '] ' + R.hf.label + ' BLACKPINK THE GAME_THE GIRLS Song Royalty Report.xlsx', XLSX.write(reportWb(w), { type: 'array', bookType: 'xlsx' }));
      ctx.toast(w.key + ' 리포트를 내려받았습니다');
    });
    ctx.root.addEventListener('change', function (e) {
      var t = e.target;
      if (t.id === 'inHalf') { var p = t.value.split('-'); S.y = +p[0]; S.h = +p[1]; merchDefault(); syncInputs(); recalc(); }
      if (t.id === 'inTx') { S.tx = t.value; render(); }
      if (t.id === 'mdFile' && t.files[0]) {
        var f = t.files[0]; t.value = ''; $('mdNote').textContent = ' · 수불부 읽는 중… (용량이 크면 수십 초 걸립니다)';
        var rd = new FileReader();
        rd.onload = function () {
          try {
            var data = new Uint8Array(rd.result), names = XLSX.read(data, { type: 'array', bookSheets: true }).SheetNames;
            var want = names.filter(function (n) { return new RegExp('FY\\s*' + S.y).test(n) && /Summary/i.test(n); });
            var wb = XLSX.read(data, { type: 'array', sheets: want.length ? want : undefined });
            var m = TG.merchFromLedger(wb, TG.half(S.y, S.h));
            S.merchSales = m.sales; S.merchNet = m.net; S.merchExtras = m.extras;
            S.merchNote = f.name + ' · ' + m.byMonth.filter(function (x) { return x.sales || x.net; }).map(function (x) { return x.m + '월 ' + fmt(x.net); }).join(', ');
            syncInputs(); recalc(); ctx.toast('수불부에서 OST MERCH 금액을 채웠습니다');
          } catch (err) { $('mdNote').textContent = ' · 수불부를 읽지 못했습니다: ' + (err && err.message ? err.message : err); }
        };
        rd.readAsArrayBuffer(f);
      }
    });
    ctx.root.addEventListener('input', function (e) {
      var id = e.target.id, v = e.target.value.trim();
      if (id === 'inFx') { S.fx = v; S.fxNote = '직접 입력'; } else if (id === 'inMs') { S.merchSales = v; S.merchNote = '직접 입력'; } else if (id === 'inMn') { S.merchNet = v; S.merchNote = '직접 입력'; } else return;
      clearTimeout(recalc._t); recalc._t = setTimeout(recalc, 250);
    });
    $('btnSave').addEventListener('click', async function () {
      try {
        $('outMsg').textContent = '저장 중…';
        var items = parsed.rows.filter(function (r) { return !D.ledger[r.key]; }).map(function (r) { return { name: r.key, member: JSON.stringify({ key: r.key, d: r.d, s: r.s, a: r.a, v: r.v, w: r.w, tk: '' }) }; });
        if (items.length) await ctx.post('saveMapping', { items: items });
        await ctx.post('saveQuarter', { record: { quarter: 'WR ' + R.hf.label, supply: R.totalK, vat: 0, gross: R.totalK, txDate: S.tx, file: ctx.fileName,
          lines: R.writers.map(function (w) { return { label: w.key, amt: w.totalK, usd: w.totalU, wht: w.wht }; }), royaltyTable: [{ fx: R.fx, merch: R.merch, net: R.net }] } });
        $('outMsg').textContent = ''; ctx.toast(R.hf.kor + ' 작가 배분을 확정 저장했습니다');
      } catch (e) { $('outMsg').textContent = '저장하지 못했습니다 (' + e.message + ')'; }
    });

    wire(ctx, TXT, TABLES, ['calc', 'out', 'pum', 'erp'], 'tg.writer.tab');
    recalc();
    return { quarter: TG.half(S.y, S.h).label, note: parsed.sheet + ' · ' + parsed.rows.length + '행 + DB 원장 ' + Object.keys(D.ledger).length + '행' };
  }

  var detect = function (wb) { try { TG.parse(wb); return true; } catch (e) { return false; } };
  DESK.register({ id: 'tg-takeone', server: 'thegirls', name: 'THE GIRLS · 테이크원', desc: "BPTG OST 'THE GIRLS' 음원 로열티 수익 배분 · 분기 매입",
    fileHint: '회계팀 THE GIRLS 음원 매출 데이터 엑셀', outputs: ['정산서 엑셀', '세금계산서 요청 메일', '매입품의', 'ERP 매입 입력값'], detect: detect, mount: mountTakeone });
  DESK.register({ id: 'tg-writer', server: 'thegirls', name: 'THE GIRLS · 해외 작가', desc: 'PATRIOT·LINDGREN 음원·OST MERCH 로열티 배분 · 반기 매입 (USD)',
    fileHint: '회계팀 THE GIRLS 음원 매출 데이터 엑셀 (같은 파일)', outputs: ['작가별 Royalty Report', '인보이스 요청 메일', '매입품의', 'ERP 사업소득 입력값'], detect: detect, mount: mountWriter });
})();
