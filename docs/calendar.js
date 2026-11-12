/* 정산 캘린더 — 월별 할 일: 선택한 달에 정산 프로젝트별로 "언제 · 무엇을 · 어떤 순서로" 해야 하는지 보여 준다.
 * - 맨 위: 이 달 YG 품의 마감일(공통) + 이 달 전체 할 일 순서(날짜순)
 * - 아래: 프로젝트별 카드(YGP · TRUZ · BPTG 테이크원 · LINE 프리미엄 · THE GIRLS 작가)에 할 일 ①②③ 순서, 날짜, D-day, 완료 체크
 * 규칙은 TASKS 한 곳에 모은다. 마감일이 토·일·빨간날(HOLI 목록)이면 직전 영업일로 당긴다.
 * 완료 체크·날짜 수정·담당자 이름은 DB 시트 "CAL_일정" 탭에 저장한다(키: done|… / 프로젝트|YYYY-MM|번호 / cfg|acct). */
(function () {
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var DOW = ['일', '월', '화', '수', '목', '금', '토'];
  function day(y, m, d) { return new Date(y, m - 1, d); }                       // m 이 13 이상·0 이하면 연도 자동 이월
  // 관공서 공휴일(대체공휴일·선거일 포함). 2026년부터 노동절(5/1)·제헌절(7/17) 공휴일. 근거: 우주항공청 월력요항·공개 달력 자료(2026.6~ 확인)
  // 회사가 공식적으로 쉬는 달력 빨간날(법정 공휴일·대체공휴일·선거일)만 넣는다. 임시공휴일은 정부 지정 시 추가
  var HOLI = {
    '2026-01-01': '신정', '2026-02-16': '설날 연휴', '2026-02-17': '설날', '2026-02-18': '설날 연휴', '2026-03-02': '대체공휴일(삼일절)', '2026-05-01': '노동절',
    '2026-05-05': '어린이날', '2026-05-25': '대체공휴일(부처님오신날)', '2026-06-03': '지방선거', '2026-07-17': '제헌절', '2026-08-17': '대체공휴일(광복절)',
    '2026-09-24': '추석 연휴', '2026-09-25': '추석', '2026-10-05': '대체공휴일(개천절)', '2026-10-09': '한글날', '2026-12-25': '성탄절',
    '2027-01-01': '신정', '2027-02-08': '설날 연휴', '2027-02-09': '대체공휴일(설날)', '2027-03-01': '삼일절', '2027-05-03': '대체공휴일(노동절)', '2027-05-05': '어린이날',
    '2027-05-13': '부처님오신날', '2027-07-19': '대체공휴일(제헌절)', '2027-08-16': '대체공휴일(광복절)', '2027-09-14': '추석 연휴', '2027-09-15': '추석', '2027-09-16': '추석 연휴',
    '2027-10-04': '대체공휴일(개천절)', '2027-10-11': '대체공휴일(한글날)', '2027-12-27': '대체공휴일(성탄절)',
    '2028-01-26': '설날 연휴', '2028-01-27': '설날', '2028-01-28': '설날 연휴', '2028-03-01': '삼일절', '2028-04-12': '국회의원선거', '2028-05-01': '노동절', '2028-05-02': '부처님오신날',
    '2028-05-05': '어린이날', '2028-06-06': '현충일', '2028-07-17': '제헌절', '2028-08-15': '광복절', '2028-10-02': '추석 연휴', '2028-10-03': '추석·개천절', '2028-10-04': '추석 연휴',
    '2028-10-05': '대체공휴일(개천절)', '2028-10-09': '한글날', '2028-12-25': '성탄절'
  };
  var key0 = function (d) { return d.getFullYear() + '-' + (d.getMonth() < 9 ? '0' : '') + (d.getMonth() + 1) + '-' + (d.getDate() < 10 ? '0' : '') + d.getDate(); };
  function holiday(d) { var k = key0(d); return HOLI[k] || ''; }
  function off(d) { return d.getDay() === 0 || d.getDay() === 6 || !!holiday(d); }
  function prevBiz(dt) { var d = new Date(dt); while (off(d)) d.setDate(d.getDate() - 1); return d; }   // 토·일·공휴일 → 직전 영업일
  function nextBiz(dt) { var d = new Date(dt); while (off(d)) d.setDate(d.getDate() + 1); return d; }   // 토·일·공휴일 → 다음 영업일
  function shiftNote(base, d) { if (key0(base) === key0(d)) return ''; var w = holiday(base) || (base.getDay() === 0 ? '일요일' : '토요일'); return (base.getMonth() + 1) + '/' + base.getDate() + '이 ' + w + '이라 ' + (d.getMonth() + 1) + '/' + d.getDate() + '(' + DOW[d.getDay()] + ')로 당김'; }
  function addDays(dt, n) { var d = new Date(dt); d.setDate(d.getDate() + n); return d; }
  var md = function (d) { return (d.getMonth() + 1) + '/' + d.getDate() + '(' + DOW[d.getDay()] + ')'; };
  var ym = function (y, m) { var d = day(y, m, 1); return d.getFullYear() + '-' + (d.getMonth() < 9 ? '0' : '') + (d.getMonth() + 1); };
  var iso = function (d) { return d.getFullYear() + '-' + (d.getMonth() < 9 ? '0' : '') + (d.getMonth() + 1) + '-' + (d.getDate() < 10 ? '0' : '') + d.getDate(); };
  var fromIso = function (v) { var t = String(v || '').split('-'); return t.length === 3 ? day(+t[0], +t[1], +t[2]) : null; };
  var today0 = function () { var t = new Date(); return day(t.getFullYear(), t.getMonth() + 1, t.getDate()); };
  var MEMO_LINE = 'LINE 프리미엄 판매 종료(26년 8월말) — 이후에는 잔여 매출 Note of Payment가 온 달만 진행';
  var MEMO_BPTG = '테이크원 계약 종료(정산 대상 판매월 26년 5월까지, 25.4Q~26.2Q 품의로 정산 완료) — 잔여 매출이 생길 때만 진행. 일정은 참고용';

  // YG 품의 마감일: 매월 3~5일(주말·공휴일 제외). before = 마감 시작 전 영업일(지연 공유·품의 상신 기한)
  // 3~5일 중 영업일만 마감 기간으로 본다. 3~5일이 모두 휴일이면 5일 직전 영업일 하루
  function ygDeadline(y, m) {
    var biz = [3, 4, 5].map(function (n) { return day(y, m, n); }).filter(function (d) { return !off(d); });
    var s = biz.length ? biz[0] : prevBiz(day(y, m, 5)), e = biz.length ? biz[biz.length - 1] : s;
    return { start: s, end: e, before: prevBiz(addDays(s, -1)), text: md(s) + (iso(s) !== iso(e) ? ' ~ ' + md(e) : '') };
  }
  var qLabel = function (y, m) { var e = m - 1 <= 0 ? m + 11 : m - 1, ey = m - 1 <= 0 ? y - 1 : y; return String(ey).slice(2) + '년 ' + Math.ceil(e / 3) + 'Q'; };   // m = 분기 다음 달 → 직전 분기
  var phase = function (m) { return ((m - 1) % 3) + 1; };   // 1: 분기 다음 달(1·4·7·10월) 2: 다다음 달(2·5·8·11월) 3: 분기 마지막 달

  // 프로젝트별 그 달 할 일 (순서대로). d 가 없는 할 일은 "앞 단계 직후"
  function TASKS(y, m, acct) {
    var Y = day(y, m, 1); y = Y.getFullYear(); m = Y.getMonth() + 1;
    var dl = ygDeadline(y, m), pm = m === 1 ? 12 : m - 1, ph = phase(m), first = nextBiz(day(y, m, 1)), r9 = prevBiz(day(y, m, 9));
    var who = acct || '회계팀 담당';
    var out = [];
    // YGP — 매월: 9일 수령인데 YG 품의 마감(3~5일)이 먼저 오므로, 마감 전에 지연을 미리 공유한다
    var chk = first < dl.before ? first : prevBiz(addDays(dl.before, -1)), pre = first <= dl.before ? first : dl.before;   // 확인 단계는 마감 전날보다 늦지 않게
    out.push({ id: 'ygp', name: 'YGP', sub: 'MD 로열티 · 응원봉 · 매월', color: 'var(--mint-2)', bg: 'var(--mint-soft)', label: pm + '월분 정산', tasks: [
      { d: chk, t: 'YGP에 ' + pm + '월분 정산서 수령 예정일 확인', n: '9일 전후 수령인지, 더 늦어지는지 확인' },
      { d: dl.before, t: who + '께 YGP 정산 지연 여부 전달', n: 'YG 품의 마감일(' + dl.text + ') 이후에 정산서가 오므로 마감 전에 미리 공유' },
      { d: r9, t: '정산서 수령 (원본 · 빅뱅 제외본)', n: '9일 부근까지 수령, 늦으면 10일' + (shiftNote(day(y, m, 9), r9) ? ' — ' + shiftNote(day(y, m, 9), r9) + ' (그날까지 보내 달라고 YGP에 요청)' : '') },
      { d: r9, t: '받은 당일 바로 처리: 매출품의 2건(MD · 응원봉) → ERP → 회계팀 빅뱅 메일', n: '정산 계산기에서 원본·빅뱅 제외본을 함께 올리면 한 번에 생성' }] });
    // TRUZ — 분기
    var t = { id: 'truz', name: 'TRUZ', sub: 'IPX × 트레저 · 분기', color: 'var(--sky-2)', bg: 'var(--sky-soft)' };
    if (ph === 1) { t.label = qLabel(y, m) + ' 정산'; t.tasks = [
      { d: first, t: 'IPX에 ' + qLabel(y, m) + ' 정산서 요청 — 15일까지', n: '계약상 23일이지만 앞으로 15일까지 받기로' },
      { d: prevBiz(day(y, m, 15)), t: '정산서 수령 마감', n: '못 받았으면 IPX에 바로 재요청' },
      { t: '받는 즉시 처리: YG검토본 → 매출품의 → ERP', n: '바로 처리하면 다음 달 YG 품의 마감 전에 여유 있게 끝남' }]; }
    else if (ph === 2) { var ql = qLabel(y, m - 1); t.label = ql + ' 정산'; t.tasks = [
      { d: pre, t: ql + ' 매출품의 상신 여부 확인', n: '지난달에 끝냈으면 완료 체크' },
      { d: dl.end, t: '(아직이면) 매출품의 상신 — YG 품의 마감(' + dl.text + ') 내', n: '26.1Q 5/4 · 26.2Q 8/3 기안처럼 마감 기간 안에 상신 · 늦어지면 마감 전에 ' + who + '께 공유' }]; }
    else { t.tasks = []; t.nextHint = md(nextBiz(day(y, m + 1, 1))) + ' IPX에 ' + qLabel(y, m + 1) + ' 정산서 요청'; }
    out.push(t);
    // BPTG 테이크원 매입 — 분기
    var b = { id: 'bptg', name: 'BPTG 테이크원', sub: 'THE GIRLS 음원 매입 · 분기', color: '#7a6fd0', bg: 'var(--lilac-soft)', memo: MEMO_BPTG };
    if (ph === 1) { b.label = qLabel(y, m) + ' 정산'; b.tasks = [
      { d: prevBiz(day(y, m, 15)), t: '회계팀에 ' + qLabel(y, m) + ' THE GIRLS 음원 매출 데이터 요청 (중순)', n: '' },
      { t: '데이터 받으면 정산서 만들어 테이크원에 보내고 세금계산서 발행 요청', n: '정산 계산기 THE GIRLS · 테이크원 메뉴' }]; }
    else if (ph === 2) { var bl = qLabel(y, m - 1); b.label = bl + ' 정산'; b.tasks = [
      { d: pre, t: '테이크원 세금계산서 수령 확인', n: '못 받았으면 재요청' },
      { d: dl.end, t: '매입품의 상신 — YG 품의 마감(' + dl.text + ') 내', n: '세금계산서 작성일을 거래일자로 입력 (25.4Q~26.2Q: 7/31 세금계산서 → 8/4 기안)' }]; }
    else { b.tasks = []; b.nextHint = md(prevBiz(day(y, m + 1, 15))) + ' 회계팀에 ' + qLabel(y, m + 1) + ' 데이터 요청'; }
    out.push(b);
    // LINE 프리미엄 — 매월: 중순 LINE 자동 메일(Note of Payment)로 전월분 정산서 수령 → 인보이스일 환율 → 품의·ERP (인수인계: "라인은 매월 중순에 자동 메일로 옴")
    var ndl = ygDeadline(day(y, m + 1, 1).getFullYear(), day(y, m + 1, 1).getMonth() + 1);
    out.push({ id: 'line', name: 'LINE 프리미엄', sub: 'LY Corporation · JPY · 매월', color: '#1f9d55', bg: '#e3f6ea', memo: MEMO_LINE, label: pm + '월분 정산', tasks: [
      { d: prevBiz(day(y, m, 15)), t: 'LINE Note of Payment(' + pm + '월분) 자동 메일 수령 확인', n: '매월 중순 자동 메일로 옴 — 안 오면 LINE 측에 확인' },
      { t: '인보이스 발행일 JPY100 환율 확인 → 정산 계산기 LINE 메뉴에 입력', n: '서울외국환중개 매매기준율' },
      { t: '매출품의 상신 → ERP AR 업로드', n: '받자마자 처리 권장 · 늦어도 다음 달 YG 품의 마감(' + ndl.text + ') 전' }] });
    // THE GIRLS 해외 작가 — 반기 (26.1H 품의 26-0600: 인보이스 8/1, 지급 9/30)
    var w = { id: 'tgw', name: 'THE GIRLS 작가', sub: 'PATRIOT · LINDGREN · USD · 반기', color: '#c0569a', bg: '#fbe8f3' };
    var hf = function (yy, mm) { return String(yy).slice(2) + '년 ' + (mm <= 6 ? '상반기' : '하반기'); };
    // 반기 마지막 달 + 1(1·7월) 중순: 데이터 요청 → 리포트 준비 / + 2(2·8월) 초: 리포트 보내며 인보이스 요청 → 받자마자 품의(이 달 YG 품의 마감 맞추기)
    if (m === 1 || m === 7) { var hl = m === 7 ? hf(y, 6) : hf(y - 1, 12); w.label = hl + ' 정산'; w.tasks = [
      { d: prevBiz(day(y, m, 15)), t: '회계팀에 ' + hl + ' THE GIRLS 음원 매출 데이터 요청 (중순)', n: '테이크원 분기 정산과 같은 데이터' },
      { t: '데이터 받으면 반기 말일 환율로 작가별 Royalty Report 미리 작성', n: (m === 7 ? '6/30' : '12/31') + ' 서울외국환중개 USD 매매기준율 · 정산 계산기 THE GIRLS · 해외 작가 메뉴 · 발송은 다음 달 초' }]; }
    else if (m === 2 || m === 8) { var hl2 = m === 8 ? hf(y, 6) : hf(y - 1, 12); w.label = hl2 + ' 정산'; w.tasks = [
      { d: first, t: '작가 2곳에 Royalty Report 보내면서 인보이스 발행 요청 (월초)', n: '정산 계산기의 인보이스 요청 메일 사용 · YG 품의 마감(' + dl.text + ')에 맞추도록 최대한 빨리' },
      { t: '인보이스 받자마자 매입품의 상신 → ERP 사업소득(원천세) 입력', n: '이 달 YG 품의 마감(' + dl.text + ')에 맞추도록 노력 · 넘기면 회계팀에 지연 공유 · 지급예정일은 둘째 주 목요일 기준' }]; }
    else { w.tasks = []; var nm2 = m < 7 ? 7 : 13; w.nextHint = md(prevBiz(day(y, nm2, 15))) + ' 회계팀에 ' + (m < 7 ? hf(y, 6) : hf(y, 12)) + ' 음원 매출 데이터 요청'; }
    out.push(w);
    return { dl: dl, projects: out };
  }

  // ---------- 저장(DB 시트 CAL_일정) ----------
  var OVK = 'desk.cal.ov', OV = null, SRV = null, API = null, ROOT = null, MSG = '', STATE = null, EDIT = null;
  function loadOv() { if (OV) return OV; try { OV = JSON.parse(localStorage.getItem(OVK) || '{}') || {}; } catch (e) { OV = {}; } return OV; }
  function cacheOv(o) { OV = o || {}; try { localStorage.setItem(OVK, JSON.stringify(OV)); } catch (e) { } }
  function sync() {
    if (!API || !API.get) { SRV = false; return; }
    API.get({ section: 'calendar' }).then(function (j) {
      var local = loadOv(), keys = Object.keys(local).filter(function (k) { return !j.ov[k] && !/^cfg\|/.test(k); });
      SRV = true; cacheOv(j.ov); if (ROOT) render(ROOT);
      if (keys.length) API.post('saveCalendar', { items: keys.map(function (k) { return { key: k, d: local[k].d, end: local[k].end }; }) }).then(function (r) { cacheOv(r.ov); if (ROOT) render(ROOT); }, function () { });
    }, function () { SRV = false; if (ROOT) render(ROOT); });
  }
  function persist(items, remove) {
    var o = Object.assign({}, loadOv());
    remove.forEach(function (p) { Object.keys(o).forEach(function (k) { if (k.indexOf(p) === 0) delete o[k]; }); });
    items.forEach(function (it) { o[it.key] = { d: it.d }; if (it.end) o[it.key].end = it.end; });
    cacheOv(o);
    if (!SRV) { MSG = '서버에 연결되지 않아 이 브라우저에만 저장했습니다 (연결되면 DB 시트로 자동으로 옮겨집니다)'; return Promise.resolve(); }
    MSG = '';
    return API.post('saveCalendar', { items: items, remove: remove }).then(function (r) { cacheOv(r.ov); }, function (e) { MSG = 'DB 시트 저장 실패 — 이 브라우저에만 저장됨 (' + e.message + ')'; });
  }
  function monthData(y, m) {
    var ov = loadOv(), acct = (ov['cfg|acct'] || {}).d || '';
    var R = TASKS(y, m, acct), key = ym(y, m);
    R.projects.forEach(function (p) {
      var prevD = null, nextMarked = false;
      p.tasks.forEach(function (t, i) {
        var o = ov[p.id + '|' + key + '|' + i];
        t.k = p.id + '|' + key + '|' + i; t.d0 = t.d;
        if (o && o.d && t.d) { t.d = fromIso(o.d); t.ov = true; }
        t.when = t.d || prevD; prevD = t.when;
        t.done = !!ov['done|' + t.k]; t.doneAt = t.done ? ov['done|' + t.k].d : '';
        if (!t.done && !nextMarked) { t.next = true; nextMarked = true; }
      });
    });
    return R;
  }

  function dlab(d, T) { if (!d) return ''; var n = Math.round((d - T) / 864e5); return n === 0 ? 'D-day' : n > 0 ? 'D-' + n : 'D+' + (-n); }
  function render(root) {
    ROOT = root;
    var T = today0(); STATE = STATE || { y: T.getFullYear(), m: T.getMonth() + 1 };
    var y = STATE.y, m = STATE.m, R = monthData(y, m), dl = R.dl;
    var hol = []; for (var hd = day(y, m, 1); hd.getMonth() === day(y, m, 1).getMonth(); hd = addDays(hd, 1)) if (holiday(hd)) hol.push(md(hd) + ' ' + holiday(hd));
    var all = [];
    R.projects.forEach(function (p) { p.tasks.forEach(function (t, i) { all.push({ p: p, t: t, i: i }); }); });
    // 다음 달 할 일 중 날짜가 이 달에 오는 것(예: 1월 YG 품의 마감 대비 12/30·12/31)도 이 달 순서에 함께 보여 준다
    var nm = day(y, m + 1, 1), R2 = monthData(nm.getFullYear(), nm.getMonth() + 1);
    R2.projects.forEach(function (p) { p.tasks.forEach(function (t, i) { if (t.d && t.d.getMonth() === m - 1 && t.d.getFullYear() === y) all.push({ p: p, t: t, i: i, ahead: nm.getMonth() + 1 }); }); });
    R = monthData(y, m);
    all.sort(function (a, b) { return (a.t.when - b.t.when) || (a.p.id === b.p.id ? a.i - b.i : 0); });
    var order = all.map(function (x) {
      var late = !x.t.done && x.t.when && x.t.when < T;
      return '<li class="mo-row' + (x.t.done ? ' done' : '') + (late ? ' late' : '') + '"><b class="mo-d">' + (x.t.d ? esc(md(x.t.d)) : '이어서') + '</b><span class="cal-tag" style="--c:' + x.p.color + ';--cb:' + x.p.bg + '">' + esc(x.p.name) + '</span><span class="mo-t">' + esc(x.t.t) + (x.ahead ? ' <small class="muted">(' + x.ahead + '월 마감 대비 · ' + x.ahead + '월 카드에서 체크)</small>' : '') + '</span>' +
        '<span class="mo-dd">' + (x.t.done ? '완료' : late ? '지연 ' + esc(dlab(x.t.when, T)) : x.t.d ? esc(dlab(x.t.d, T)) : '') + '</span></li>';
    }).join('');
    var cards = R.projects.map(function (p) {
      var ed = EDIT === p.id;
      var body = p.tasks.length ? '<ol class="tl">' + p.tasks.map(function (t, i) {
        var late = !t.done && t.when && t.when < T;
        var dateHtml = ed && t.d ? '<input type="date" data-i="' + i + '" data-def="' + iso(t.d0) + '" value="' + iso(t.d) + '">'
          : '<b class="tl-date">' + (t.d ? esc(md(t.d)) : '이어서') + '</b>' + (t.ov ? '<span class="tl-ov" title="기본 ' + esc(md(t.d0)) + '">수정됨</span>' : '');
        return '<li class="tl-step' + (t.done ? ' done' : '') + (t.next ? ' next' : '') + (late ? ' late' : '') + '">' +
          '<label class="tl-chk" title="완료 체크"><input type="checkbox" data-done="' + esc(t.k) + '"' + (t.done ? ' checked' : '') + '><span>' + (i + 1) + '</span></label><div class="tl-body">' +
          '<div class="tl-top">' + dateHtml + (ed ? '' : '<span class="tl-dd">' + (t.done ? '완료 ' + esc(String(t.doneAt).slice(5).replace('-', '/')) : late ? '지연' : t.d ? esc(dlab(t.d, T)) : '') + '</span>') + '</div>' +
          '<div class="tl-t">' + esc(t.t) + '</div>' + (t.n ? '<div class="small muted">' + esc(t.n) + '</div>' : '') + '</div></li>';
      }).join('') + '</ol>'
        : '<p class="small muted tl-none">이번 달 할 일 없음 · 다음 일정 ' + esc(p.nextHint || '') + '</p>';
      var tools = ed ? '<div class="row tl-tools"><button class="btn sm primary" data-save="' + p.id + '" type="button">저장</button><button class="btn sm" data-reset="' + p.id + '" type="button">기본값으로</button><button class="btn sm" data-cancel="1" type="button">취소</button></div>' : '';
      var cnt = p.tasks.length ? p.tasks.filter(function (t) { return t.done; }).length + '/' + p.tasks.length : '';
      return '<article class="tl-card' + (p.tasks.length ? '' : ' empty') + '" data-card="' + p.id + '" style="--c:' + p.color + ';--cb:' + p.bg + '"><header><span class="cal-tag">' + esc(p.name) + '</span><span class="small muted">' + esc(p.sub) + '</span>' +
        (p.label ? '<b class="tl-cyc">' + esc(p.label) + (cnt ? ' · ' + cnt : '') + '</b>' : '') +
        (p.tasks.length && !ed ? '<button class="btn sm tl-editbtn" data-edit="' + p.id + '" type="button">날짜 수정</button>' : '') + '</header>' + body + tools +
        (p.memo && p.tasks.length ? '<p class="tl-memo small">※ ' + esc(p.memo) + '</p>' : '') + '</article>';
    }).join('');
    root.innerHTML =
      '<div class="ph"><h2>정산 캘린더 · 월별 할 일</h2><p>선택한 달에 무엇을 어떤 순서로 해야 하는지 정산별로 보여 줍니다. 끝낸 일은 번호를 눌러 완료 체크하세요(DB 시트에 저장). 마감일이 주말·공휴일이면 직전 영업일로 당깁니다(2026~2028년 빨간날 반영).</p>' +
      (MSG ? '<p class="small" style="margin:6px 0 0;color:var(--warn)">' + esc(MSG) + '</p>' : SRV === false ? '<p class="small" style="margin:6px 0 0;color:var(--warn)">서버에 연결되지 않아 완료 체크·날짜 수정이 이 브라우저에만 저장됩니다</p>' : '') + '</div>' +
      '<div class="row cal-nav"><button class="btn sm" data-mon="-1" type="button">◀ 이전 달</button><b class="cal-title">' + y + '년 ' + m + '월</b><button class="btn sm" data-mon="1" type="button">다음 달 ▶</button><button class="btn sm" data-mon="0" type="button">이번 달</button></div>' +
      '<div class="mo-yg"><b>YG 품의 마감일 ' + esc(dl.text) + '</b><span>전월 거래일자(세금계산서·인보이스 작성일) 건은 이 기간 안에 상신 · 늦어지는 정산은 마감 전 <b>' + esc(md(dl.before)) + '</b>까지 회계팀에 미리 공유</span>' +
        (hol.length ? '<span class="mo-hol">이 달 공휴일: ' + esc(hol.join(' · ')) + '</span>' : '') + '</div>' +
      '<section class="mo-order"><h3>' + m + '월 전체 할 일 순서</h3><ol>' + order + '</ol></section>' +
      '<div class="tl-grid">' + cards + '</div>';
    root.querySelectorAll('[data-mon]').forEach(function (b) {
      b.addEventListener('click', function () {
        var v = +b.getAttribute('data-mon'); EDIT = null;
        if (v === 0) STATE = { y: T.getFullYear(), m: T.getMonth() + 1 }; else { var d = day(STATE.y, STATE.m + v, 1); STATE = { y: d.getFullYear(), m: d.getMonth() + 1 }; }
        render(root);
      });
    });
    root.querySelectorAll('[data-done]').forEach(function (c) {
      c.addEventListener('change', function () {
        var k = 'done|' + c.getAttribute('data-done');
        (c.checked ? persist([{ key: k, d: iso(today0()) }], []) : persist([], [k])).then(function () { render(root); }); render(root);
      });
    });
    root.querySelectorAll('[data-edit]').forEach(function (b) { b.addEventListener('click', function () { EDIT = b.getAttribute('data-edit'); render(root); }); });
    root.querySelectorAll('[data-cancel]').forEach(function (b) { b.addEventListener('click', function () { EDIT = null; render(root); }); });
    root.querySelectorAll('[data-save]').forEach(function (b) {
      b.addEventListener('click', function () {
        var id = b.getAttribute('data-save'), pre = id + '|' + ym(STATE.y, STATE.m) + '|', items = [], rm = [], bad = false;
        b.closest('.tl-card').querySelectorAll('input[type=date]').forEach(function (inp) {
          if (!inp.value) { bad = true; return; }
          var k = pre + inp.getAttribute('data-i');
          if (inp.value !== inp.getAttribute('data-def')) items.push({ key: k, d: inp.value }); else if (loadOv()[k]) rm.push(k);
        });
        if (bad) { alert('날짜를 모두 입력해 주세요'); return; }
        EDIT = null; if (!items.length && !rm.length) { render(root); return; }
        persist(items, rm).then(function () { render(root); }); render(root);
      });
    });
    root.querySelectorAll('[data-reset]').forEach(function (b) {
      b.addEventListener('click', function () {
        // 날짜 수정만 지운다(완료 체크는 유지): 키 "프로젝트|YYYY-MM|" 로 시작하는 것만
        EDIT = null; persist([], [b.getAttribute('data-reset') + '|' + ym(STATE.y, STATE.m) + '|']).then(function () { render(root); }); render(root);
      });
    });
  }
  window.DESK_CAL = { render: function (root, api) { if (api && !API) { API = api; sync(); } else if (api && SRV === false) sync(); render(root); }, tasks: TASKS, ygDeadline: ygDeadline };
})();
