/* 정산 캘린더: 매월 며칠에 무엇을 해야 하는지(절차별 마감일)를 달력 + 목록으로 보여 준다.
 * 규칙은 아래 RULES 한 곳에 모은다. 마감일이 토·일이면 직전 금요일로 당긴다(공휴일은 반영하지 않음). */
(function () {
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  var key = function (d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var DOW = ['일', '월', '화', '수', '목', '금', '토'];
  function day(y, m, d) { return new Date(y, m - 1, d); }                       // m: 1~12 (넘치면 다음 해로 자동 이월)
  function prevBiz(dt) { var d = new Date(dt); while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() - 1); return d; }   // 토·일 → 직전 금요일
  function nextBiz(dt) { var d = new Date(dt); while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1); return d; }   // 월초 첫 영업일
  var qOf = function (m) { return Math.ceil(m / 3); };

  var TRACKS = {
    ygp: { name: 'YGP', color: 'var(--mint-2)', bg: 'var(--mint-soft)' },
    truz: { name: 'TRUZ', color: 'var(--sky-2)', bg: 'var(--sky-soft)' },
    bptg: { name: 'BPTG 테이크원', color: '#7a6fd0', bg: 'var(--lilac-soft)' }
  };
  // 한 달(y, m)에 걸리는 일정. 분기 정산은 "분기 마지막 달(E)" 기준: E+1월 15일 정산서·데이터, E+2월 초 품의
  function RULES(y, m) {
    var ev = [], pm = m === 1 ? 12 : m - 1, py = m === 1 ? y - 1 : y;
    // YGP — 매월
    var r9 = day(y, m, 9), r9b = prevBiz(r9);
    ev.push({ t: 'ygp', d: r9b, title: 'YGP 정산서 수령 (' + pm + '월분 MD·응원봉)', note: '9일 수령(늦으면 10일)' + (r9b.getDate() !== 9 ? ' · 9일이 주말이라 ' + (r9b.getMonth() + 1) + '/' + r9b.getDate() + '(금) 요청·수령' : '') + ' · 원본·빅뱅 제외본 받자마자 바로 처리(품의·ERP·회계팀 빅뱅 메일)' });
    ev.push({ t: 'ygp', d: day(y, m, 3), end: day(y, m, 5), title: 'YGP 내부 품의 마감 (3~5일)', note: '더 늦어지지 않도록 수령 당일 상신 원칙' });
    // 분기: m 이 E+1 또는 E+2 인 경우
    var e1 = m - 1 === 0 ? 12 : m - 1;                     // m 이 E+1 이면 E = m-1
    if (e1 % 3 === 0) {
      var ey = m === 1 ? y - 1 : y, q = qOf(e1), ql = String(ey).slice(2) + '년 ' + q + 'Q';
      ev.push({ t: 'truz', d: day(y, m, 1), title: 'TRUZ ' + ql + ' 정산서 요청 (IPX)', note: '15일까지 보내 달라고 요청 (계약상 23일)' });
      var t15 = prevBiz(day(y, m, 15));
      ev.push({ t: 'truz', d: t15, title: 'TRUZ ' + ql + ' 정산서 수령 마감', note: '받자마자 바로 처리하면 이슈 없음 (YG검토본 → 품의 → ERP)' });
      ev.push({ t: 'bptg', d: t15, title: 'BPTG 테이크원 ' + ql + ' 회계팀에 음원 매출 데이터 요청', note: '중순 요청 → 데이터 받으면 테이크원에 정산서 보내고 세금계산서 발행 요청', memo: true });
    }
    var e2 = m - 2 <= 0 ? m + 10 : m - 2;                  // m 이 E+2 이면 E = m-2
    if (e2 % 3 === 0) {
      var ey2 = m <= 2 ? y - 1 : y, q2 = qOf(e2), ql2 = String(ey2).slice(2) + '년 ' + q2 + 'Q', b1 = nextBiz(day(y, m, 1));
      ev.push({ t: 'truz', d: b1, title: 'TRUZ ' + ql2 + ' 매출품의 상신 (늦어도 월초)', note: '정산서 받은 즉시 처리 권장 — 넉넉하게 미리 상신' });
      ev.push({ t: 'bptg', d: b1, title: 'BPTG 테이크원 ' + ql2 + ' 매입품의 (월초)', note: '테이크원 세금계산서 수령 후 상신', memo: true });
    }
    return ev;
  }
  var MEMO = '※ 테이크원 계약 종료(정산 대상 판매월 26년 5월까지, 25.4Q~26.2Q 품의로 정산 완료) — 잔여 매출 발생 시에만 진행. 일정은 참고용으로 표시';

  function eventsOf(y, m) {
    return RULES(y, m).sort(function (a, b) { return a.d - b.d || (a.t < b.t ? -1 : 1); });
  }

  function render(root, state) {
    var y = state.y, m = state.m, today = new Date(), tk = key(today);
    var ev = eventsOf(y, m), byDay = {};
    ev.forEach(function (e) {
      var d = new Date(e.d), end = e.end || e.d;
      while (d <= end) { (byDay[key(d)] = byDay[key(d)] || []).push(e); d.setDate(d.getDate() + 1); }
    });
    var first = day(y, m, 1), start = new Date(first); start.setDate(1 - first.getDay());
    var last = day(y, m + 1, 0), cells = '';
    for (var d = new Date(start); d <= last || d.getDay() !== 0; d.setDate(d.getDate() + 1)) {
      var k = key(d), inM = d.getMonth() === first.getMonth(), list = byDay[k] || [];
      cells += '<div class="cal-cell' + (inM ? '' : ' out') + (k === tk ? ' today' : '') + (d.getDay() === 0 ? ' sun' : d.getDay() === 6 ? ' sat' : '') + '">' +
        '<span class="cal-d">' + d.getDate() + '</span>' +
        (inM ? list.map(function (e) { var T = TRACKS[e.t]; return '<span class="cal-ev" style="background:' + T.bg + ';color:' + T.color + '" title="' + esc(e.title + ' — ' + e.note) + '">' + esc(e.title) + '</span>'; }).join('') : '') + '</div>';
    }
    var rows = ev.map(function (e) {
      var T = TRACKS[e.t], dd = Math.round((day(e.d.getFullYear(), e.d.getMonth() + 1, e.d.getDate()) - day(today.getFullYear(), today.getMonth() + 1, today.getDate())) / 864e5);
      var past = (e.end || e.d) < day(today.getFullYear(), today.getMonth() + 1, today.getDate());
      var dl = dd === 0 ? 'D-day' : dd > 0 ? 'D-' + dd : 'D+' + (-dd);
      var when = (e.d.getMonth() + 1) + '/' + e.d.getDate() + '(' + DOW[e.d.getDay()] + ')' + (e.end ? ' ~ ' + (e.end.getMonth() + 1) + '/' + e.end.getDate() + '(' + DOW[e.end.getDay()] + ')' : '');
      return '<tr class="' + (past ? 'past' : '') + '"><td><b>' + when + '</b></td><td><span class="cal-tag" style="background:' + T.bg + ';color:' + T.color + '">' + esc(T.name) + '</span></td><td>' + esc(e.title) +
        '<div class="small muted">' + esc(e.note) + (e.memo ? '<br>' + esc(MEMO) : '') + '</div></td><td class="n">' + (past ? '지남' : dl) + '</td></tr>';
    }).join('');
    root.innerHTML =
      '<div class="ph"><h2>정산 캘린더</h2><p>매월 며칠에 무엇을 해야 하는지, 절차별 마감일입니다. 마감일이 토·일이면 직전 금요일로 당겨 표시합니다(공휴일 미반영).</p></div>' +
      '<div class="row cal-nav"><button class="btn sm" data-cal="-1" type="button">◀ 이전 달</button><b class="cal-title">' + y + '년 ' + m + '월</b><button class="btn sm" data-cal="1" type="button">다음 달 ▶</button><button class="btn sm" data-cal="0" type="button">이번 달</button>' +
      '<span class="cal-legend">' + Object.keys(TRACKS).map(function (k) { return '<span class="cal-tag" style="background:' + TRACKS[k].bg + ';color:' + TRACKS[k].color + '">' + TRACKS[k].name + '</span>'; }).join('') + '</span></div>' +
      '<div class="cal-grid">' + DOW.map(function (w, i) { return '<div class="cal-h' + (i === 0 ? ' sun' : i === 6 ? ' sat' : '') + '">' + w + '</div>'; }).join('') + cells + '</div>' +
      '<div class="tablebox"><table class="cal-list"><tr><th>날짜</th><th>정산</th><th>할 일 · 비고</th><th class="n">D-day</th></tr>' + rows + '</table></div>' +
      '<details class="fold"><summary><b>일정 규칙</b></summary><ul class="small" style="margin:8px 0 0 18px;line-height:1.8">' +
      '<li><b>YGP (매월)</b>: 9일 정산서 수령(늦으면 10일, 주말이면 직전 금요일) → 받자마자 바로 처리. 내부 품의 마감 매월 3~5일.</li>' +
      '<li><b>TRUZ (분기)</b>: 분기 마지막 달 + 1개월 15일까지 IPX 정산서 수령(계약상 23일 → 15일로 요청) → 받자마자 처리, 늦어도 + 2개월 초(2·5·8·11월) 품의.</li>' +
      '<li><b>BPTG 테이크원 매입 (분기)</b>: 분기 마지막 달 + 1개월 중순 회계팀에 음원 매출 데이터 요청 → 테이크원 세금계산서 발행 요청 → + 2개월 초 매입품의. ' + esc(MEMO) + '</li></ul></details>';
    root.querySelectorAll('[data-cal]').forEach(function (b) {
      b.addEventListener('click', function () {
        var v = +b.getAttribute('data-cal');
        if (v === 0) { state.y = today.getFullYear(); state.m = today.getMonth() + 1; }
        else { var t = day(state.y, state.m + v, 1); state.y = t.getFullYear(); state.m = t.getMonth() + 1; }
        render(root, state);
      });
    });
  }
  var now = new Date();
  window.DESK_CAL = { state: { y: now.getFullYear(), m: now.getMonth() + 1 }, render: function (root) { render(root, this.state); }, rules: RULES };
})();
