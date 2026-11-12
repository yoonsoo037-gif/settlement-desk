/* 정산 캘린더: 정산별 카드에 이번 회차 절차를 타임라인(체크리스트)으로 보여 준다 — 날짜 · D-day · 지남/다음 할 일.
 * 규칙은 아래 CYCLES 한 곳에 모은다. 마감일이 토·일이면 직전 금요일로 당긴다(공휴일은 반영하지 않음). */
(function () {
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var DOW = ['일', '월', '화', '수', '목', '금', '토'];
  function day(y, m, d) { return new Date(y, m - 1, d); }                       // m 이 13 이상이면 다음 해로 자동 이월
  function prevBiz(dt) { var d = new Date(dt); while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() - 1); return d; }   // 토·일 → 직전 금요일
  function nextBiz(dt) { var d = new Date(dt); while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1); return d; }   // 월초 첫 영업일
  var md = function (d) { return (d.getMonth() + 1) + '/' + d.getDate() + '(' + DOW[d.getDay()] + ')'; };
  var MEMO_BPTG = '테이크원 계약 종료(정산 대상 판매월 26년 5월까지, 25.4Q~26.2Q 품의로 정산 완료) — 잔여 매출이 생길 때만 진행. 일정은 참고용';

  // 정산별 회차 생성: n = 기준 회차(월 또는 분기 마지막 달)에서 몇 번째 앞/뒤인지
  var CYCLES = [
    { id: 'yg', name: 'YG 품의 마감', sub: '공통 · 매월', color: 'var(--bad)', bg: 'var(--bad-bg)',
      make: function (y, m) { return { label: y + '년 ' + m + '월', steps: [
        { d: prevBiz(day(y, m, 3)), end: prevBiz(day(y, m, 5)), t: 'YG 품의 마감일 (3~5일)', n: '이 기간 안에 그 달 처리할 품의를 올려야 함 — 주말이면 직전 금요일까지' }] }; } },
    { id: 'ygp', name: 'YGP', sub: 'MD 로열티 · 응원봉 · 매월', color: 'var(--mint-2)', bg: 'var(--mint-soft)',
      make: function (y, m) {
        var pm = m === 1 ? 12 : m - 1, r = prevBiz(day(y, m, 9));
        return { label: pm + '월분 (' + m + '월 정산)', steps: [
          { d: r, t: '정산서 수령 (원본·빅뱅 제외본)', n: '9일 수령, 늦으면 10일' + (r.getDate() !== 9 ? ' — 9일이 주말이라 ' + md(r) + '에 요청·수령' : '') },
          { d: r, t: '받자마자 처리: 품의 2건 · ERP · 회계팀 빅뱅 메일', n: '수령 당일 상신 원칙' }] };
      } },
    { id: 'truz', name: 'TRUZ', sub: 'IPX × 트레저 · 분기', color: 'var(--sky-2)', bg: 'var(--sky-soft)', quarterly: true,
      make: function (y, q) {
        var e = q * 3, a = day(y, e + 1, 1), b = day(y, e + 2, 1);
        return { label: String(y).slice(2) + '년 ' + q + 'Q', steps: [
          { d: a, t: 'IPX에 정산서 요청', n: '15일까지 보내 달라고 요청 (계약상 23일)' },
          { d: prevBiz(day(a.getFullYear(), a.getMonth() + 1, 15)), t: '정산서 수령 마감 (15일)', n: '받자마자 YG검토본 → 품의 → ERP 바로 처리' },
          { d: nextBiz(b), t: '매출품의 상신 (늦어도 월초)', n: '넉넉하게 미리 상신' }] };
      } },
    { id: 'bptg', name: 'BPTG 테이크원', sub: 'THE GIRLS 음원 매입 · 분기', color: '#7a6fd0', bg: 'var(--lilac-soft)', quarterly: true, memo: MEMO_BPTG,
      make: function (y, q) {
        var e = q * 3, a = day(y, e + 1, 15), b = day(y, e + 2, 1);
        return { label: String(y).slice(2) + '년 ' + q + 'Q', steps: [
          { d: prevBiz(a), t: '회계팀에 음원 매출 데이터 요청 (중순)', n: '' },
          { d: null, after: true, t: '데이터 받으면 정산서 보내고 세금계산서 발행 요청', n: '데이터 수령 직후' },
          { d: nextBiz(b), t: '매입품의 상신 (월초)', n: '테이크원 세금계산서 받은 뒤' }] };
      } }
  ];

  // 날짜 직접 수정: DB 시트 "CAL_일정" 탭에 저장(서버 calendar / saveCalendar). 키 = 정산id|회차|단계번호
  // 브라우저에는 화면을 바로 그리기 위한 사본만 둔다. 서버에 저장 기능이 아직 배포되지 않았으면 이 브라우저에만 저장한다
  var OVK = 'desk.cal.ov', OV = null, SRV = null, API = null, ROOT = null, MSG = '';
  function loadOv() { if (OV) return OV; try { OV = JSON.parse(localStorage.getItem(OVK) || '{}') || {}; } catch (e) { OV = {}; } return OV; }
  function cacheOv(o) { OV = o || {}; try { localStorage.setItem(OVK, JSON.stringify(OV)); } catch (e) { } }
  function sync() {
    if (!API || !API.get) { SRV = false; return; }
    API.get({ section: 'calendar' }).then(function (j) {
      var local = loadOv(), keys = Object.keys(local).filter(function (k) { return !j.ov[k]; });
      SRV = true; cacheOv(j.ov); if (ROOT) render(ROOT);
      // 서버 배포 전에 이 브라우저에만 저장했던 수정분은 한 번 서버로 옮긴다
      if (keys.length) API.post('saveCalendar', { items: keys.map(function (k) { return { key: k, d: local[k].d, end: local[k].end }; }) }).then(function (r) { cacheOv(r.ov); if (ROOT) render(ROOT); }, function () { });
    }, function () { SRV = false; if (ROOT) render(ROOT); });
  }
  function persist(items, remove) {
    var o = Object.assign({}, loadOv());
    remove.forEach(function (p) { Object.keys(o).forEach(function (k) { if (k.indexOf(p) === 0) delete o[k]; }); });
    items.forEach(function (it) { o[it.key] = { d: it.d }; if (it.end) o[it.key].end = it.end; });
    cacheOv(o);
    if (!SRV) { MSG = '서버 저장 기능이 아직 배포되지 않아 이 브라우저에만 저장했습니다 (Code.gs 재배포 후 자동으로 DB 시트로 옮겨집니다)'; return Promise.resolve(); }
    MSG = 'DB 시트에 저장 중…';
    return API.post('saveCalendar', { items: items, remove: remove }).then(function (r) { cacheOv(r.ov); MSG = 'DB 시트(CAL_일정)에 저장했습니다'; }, function (e) { MSG = 'DB 시트 저장 실패 — 이 브라우저에만 저장됨 (' + e.message + ')'; });
  }
  var iso = function (d) { return d.getFullYear() + '-' + (d.getMonth() < 9 ? '0' : '') + (d.getMonth() + 1) + '-' + (d.getDate() < 10 ? '0' : '') + d.getDate(); };
  var fromIso = function (v) { var t = String(v || '').split('-'); return t.length === 3 ? day(+t[0], +t[1], +t[2]) : null; };
  function applyOv(c, cy, ov) {
    cy.steps.forEach(function (s, i) {
      var o = ov[c.id + '|' + cy.label + '|' + i]; if (!o || !s.d) return;
      s.d0 = s.d; s.end0 = s.end;
      if (o.d) s.d = fromIso(o.d); if (o.end && s.end) s.end = fromIso(o.end); s.ov = true;
    });
    return cy;
  }
  var today0 = function () { var t = new Date(); return day(t.getFullYear(), t.getMonth() + 1, t.getDate()); };
  // 이번 회차 = 마지막 단계가 오늘 이후인 가장 이른 회차
  function cyclesFor(c, T) {
    var list = [], y = T.getFullYear(), m = T.getMonth() + 1, ov = loadOv(), mk = function (a, b) { return applyOv(c, c.make(a, b), ov); };
    if (c.quarterly) { var q = Math.ceil(m / 3); for (var k = -3; k <= 3; k++) { var qq = q + k, yy = y + Math.floor((qq - 1) / 4); qq = ((qq - 1) % 4 + 4) % 4 + 1; list.push(mk(yy, qq)); } }
    else for (var j = -2; j <= 3; j++) { var dd = day(y, m + j, 1); list.push(mk(dd.getFullYear(), dd.getMonth() + 1)); }
    var lastOf = function (cy) { var s = cy.steps.filter(function (x) { return x.d; }); var l = s[s.length - 1]; return l.end || l.d; };
    list.sort(function (a, b) { return a.steps[0].d - b.steps[0].d; });
    var i = 0; while (i < list.length - 1 && lastOf(list[i]) < T) i++;
    return { cur: list[i], next: list[i + 1] };
  }
  function dlab(d, T) { var n = Math.round((d - T) / 864e5); return n === 0 ? 'D-day' : n > 0 ? 'D-' + n : '지남'; }

  var EDIT = null;   // 편집 중인 정산 id
  function render(root) {
    ROOT = root;
    var T = today0();
    var cards = CYCLES.map(function (c) {
      var r = cyclesFor(c, T), cy = r.cur, nextMarked = false, prevD = null, ed = EDIT === c.id;
      var steps = cy.steps.map(function (s, i) {
        var d = s.d || prevD, end = s.end || d, done = s.d ? end < T : false;
        prevD = d;
        var isNext = !done && !nextMarked; if (isNext) nextMarked = true;
        var when = s.d ? md(s.d) + (s.end ? ' ~ ' + md(s.end) : '') : '이어서';
        var dateHtml = ed && s.d
          ? '<span class="tl-edit"><input type="date" data-i="' + i + '" data-k="d" data-def="' + iso(s.d0 || s.d) + '" value="' + iso(s.d) + '">' + (s.end ? ' ~ <input type="date" data-i="' + i + '" data-k="end" data-def="' + iso(s.end0 || s.end) + '" value="' + iso(s.end) + '">' : '') + '</span>'
          : '<b class="tl-date">' + esc(when) + '</b>' + (s.ov ? '<span class="tl-ov" title="기본 ' + esc(md(s.d0) + (s.end0 ? ' ~ ' + md(s.end0) : '')) + '">수정됨</span>' : '');
        return '<li class="tl-step' + (done ? ' done' : '') + (isNext ? ' next' : '') + '"><span class="tl-dot"></span><div class="tl-body">' +
          '<div class="tl-top">' + dateHtml + (s.d && !ed ? '<span class="tl-dd">' + (done ? '지남' : dlab(s.d <= T && s.end ? T : s.d, T)) + '</span>' : '') + '</div>' +
          '<div class="tl-t">' + esc(s.t) + '</div>' + (s.n ? '<div class="small muted">' + esc(s.n) + '</div>' : '') + '</div></li>';
      }).join('');
      var nx = r.next ? r.next.steps[0] : null;
      var tools = ed
        ? '<div class="row tl-tools"><button class="btn sm primary" data-save="' + c.id + '" data-cyc="' + esc(cy.label) + '" type="button">저장</button><button class="btn sm" data-reset="' + c.id + '" data-cyc="' + esc(cy.label) + '" type="button">기본값으로</button><button class="btn sm" data-cancel="1" type="button">취소</button></div>'
        : '';
      return '<article class="tl-card" data-card="' + c.id + '" style="--c:' + c.color + ';--cb:' + c.bg + '"><header><span class="cal-tag">' + esc(c.name) + '</span><span class="small muted">' + esc(c.sub) + '</span><b class="tl-cyc">' + esc(cy.label) + '</b>' +
        (ed ? '' : '<button class="btn sm tl-editbtn" data-edit="' + c.id + '" type="button">날짜 수정</button>') + '</header>' +
        '<ol class="tl">' + steps + '</ol>' + tools +
        (c.memo ? '<p class="tl-memo small">※ ' + esc(c.memo) + '</p>' : '') +
        (nx ? '<p class="small muted tl-nx">다음 회차 ' + esc(r.next.label) + ': ' + esc(md(nx.d)) + ' ' + esc(nx.t) + '</p>' : '') + '</article>';
    }).join('');
    root.innerHTML =
      '<div class="ph"><h2>정산 캘린더</h2><p>정산별 이번 회차 절차와 마감일입니다. 색으로 강조된 단계가 다음에 할 일입니다. 마감일이 토·일이면 직전 금요일로 당겨 표시합니다(공휴일 미반영). "날짜 수정"으로 이번 회차 날짜를 직접 바꿀 수 있고, 수정한 날짜는 DB 시트(CAL_일정)에 저장되어 어느 기기에서나 같게 보입니다. 기준일 ' + esc(md(T)) + '</p>' +
        (MSG ? '<p class="small" style="margin:6px 0 0;color:var(--warn)">' + esc(MSG) + '</p>' : SRV === false ? '<p class="small" style="margin:6px 0 0;color:var(--warn)">서버 저장 기능 미배포 — 수정한 날짜는 이 브라우저에만 저장됩니다 (Code.gs 재배포 필요)</p>' : '') + '</div>' +
      '<div class="tl-grid">' + cards + '</div>';
    root.querySelectorAll('[data-edit]').forEach(function (b) { b.addEventListener('click', function () { EDIT = b.getAttribute('data-edit'); render(root); }); });
    root.querySelectorAll('[data-cancel]').forEach(function (b) { b.addEventListener('click', function () { EDIT = null; render(root); }); });
    root.querySelectorAll('[data-save]').forEach(function (b) {
      b.addEventListener('click', function () {
        var id = b.getAttribute('data-save'), cyc = b.getAttribute('data-cyc'), card = b.closest('.tl-card'), bad = false, by = {};
        card.querySelectorAll('input[type=date]').forEach(function (inp) {
          if (!inp.value) { bad = true; return; }
          var k = id + '|' + cyc + '|' + inp.getAttribute('data-i'); by[k] = by[k] || { key: k, changed: false }; by[k][inp.getAttribute('data-k')] = inp.value;
          if (inp.value !== inp.getAttribute('data-def')) by[k].changed = true;
        });
        if (bad) { alert('날짜를 모두 입력해 주세요'); return; }
        // 기본 날짜와 같은 단계는 저장하지 않고(있으면 지움), 바꾼 단계만 저장한다
        var keys = Object.keys(by), items = keys.filter(function (k) { return by[k].changed; }).map(function (k) { return { key: k, d: by[k].d, end: by[k].end }; });
        var rm = keys.filter(function (k) { return !by[k].changed && loadOv()[k]; });
        EDIT = null; if (!items.length && !rm.length) { render(root); return; }
        persist(items, rm).then(function () { render(root); }); render(root);
      });
    });
    root.querySelectorAll('[data-reset]').forEach(function (b) {
      b.addEventListener('click', function () {
        var pre = b.getAttribute('data-reset') + '|' + b.getAttribute('data-cyc') + '|';
        EDIT = null; persist([], [pre]).then(function () { render(root); }); render(root);
      });
    });
  }
  window.DESK_CAL = { render: function (root, api) { if (api && !API) { API = api; sync(); } else if (api && SRV === false) sync(); render(root); }, cycles: CYCLES, cyclesFor: cyclesFor };
})();
