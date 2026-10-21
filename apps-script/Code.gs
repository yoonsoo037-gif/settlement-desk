/**
 * 정산 데스크 — 백엔드 (Google Apps Script Web App)
 *
 * 역할: 비밀번호를 확인한 뒤에만 "정산 데스크 DB" 구글시트에서 트랙별(TRUZ 등) 설정·매핑·이력을 JSON으로
 * 내려주고, 새 제품 매핑과 분기 확정 저장을 같은 시트에 기록한다. 업로드한 정산 원본 엑셀은 브라우저 안에서만
 * 처리하므로 이 스크립트로 보내지지 않는다. 공개 저장소에는 회사 데이터가 없고, 데이터는 전부 이 스크립트가
 * 읽는 비공개 시트에만 있다.
 *
 * DB 시트 탭 규칙: <트랙코드>_매핑 / <트랙코드>_이력 / <트랙코드>_확정 / <트랙코드>_설정 (예: TRUZ_매핑)
 * 새 트랙을 추가할 때는 TRACKS 에 한 줄 추가하고 같은 이름 규칙으로 탭 4개를 만든다.
 *
 * 배포 방법(회사 계정으로 1회):
 * 1. script.google.com → 새 프로젝트 → 이름 "정산 데스크 백엔드" → 이 파일 내용을 통째로 붙여넣기 → 저장
 * 2. 프로젝트 설정(⚙) → 스크립트 속성 → 속성 추가 2개
 *    - DESK_PASSWORD : 사이트 비밀번호(원하는 값)
 *    - DB_SPREADSHEET_ID : "정산 데스크 DB" 구글시트 ID (주소창 /d/ 와 /edit 사이 문자열)
 * 3. 상단 함수 선택에서 manualAuthTest 선택 → ▶ 실행 → 권한 검토 → 회사 계정 → 허용
 * 4. 배포 → 새 배포 → 유형(⚙) 웹 앱 / 실행 계정: 나 / 액세스 권한: 모든 사용자 → 배포
 * 5. 나온 웹 앱 URL(.../exec)을 클로드에게 전달 → docs/config.js 의 appsScriptUrl 에 반영
 * 코드 수정 후 재배포: 배포 → 배포 관리 → 연필(수정) → 새 버전 (URL 유지)
 */

var TRACKS = { truz: 'TRUZ' };

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function db_() { return SpreadsheetApp.openById(prop_('DB_SPREADSHEET_ID')); }
function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function norm_(s) { return String(s == null ? '' : s).replace(/\s+/g, ''); }
function authed_(pw) { var p = prop_('DESK_PASSWORD'); return !!p && pw === p; }
function tab_(track, name) {
  var code = TRACKS[String(track || '').toLowerCase()];
  if (!code) throw new Error('unknown_track');
  var sh = db_().getSheetByName(code + '_' + name);
  if (!sh) throw new Error('missing_sheet: ' + code + '_' + name);
  return sh;
}

// 읽기: ?section=ping|config&track=truz&password=...
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!authed_(p.password)) return out_({ ok: false, error: 'unauthorized' });
  try {
    if (p.section === 'ping') return out_({ ok: true, tracks: Object.keys(TRACKS) });
    if (p.section === 'config') return out_(getConfig_(p.track));
    return out_({ ok: false, error: 'unknown_section' });
  } catch (err) {
    return out_({ ok: false, error: 'server_error', message: String(err) });
  }
}

// 쓰기: POST body(JSON 텍스트) {password, track, section, ...}. 브라우저는 text/plain으로 보내 사전요청(CORS preflight)을 피한다.
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { return out_({ ok: false, error: 'bad_json' }); }
  if (!authed_(body.password)) return out_({ ok: false, error: 'unauthorized' });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.section === 'saveMapping') return out_(saveMapping_(body.track, body.items || []));
    if (body.section === 'saveQuarter') return out_(saveQuarter_(body.track, body.record || {}));
    return out_({ ok: false, error: 'unknown_section' });
  } catch (err) {
    return out_({ ok: false, error: 'server_error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getConfig_(track) {
  var map = {};
  var mv = tab_(track, '매핑').getDataRange().getValues();
  for (var i = 1; i < mv.length; i++) { if (mv[i][0] && mv[i][1]) map[norm_(mv[i][0])] = String(mv[i][1]).trim(); }
  var hist = [];
  var hv = tab_(track, '이력').getDataRange().getValues();
  for (var j = 1; j < hv.length; j++) { if (hv[j][0]) hist.push([String(hv[j][0]), Number(hv[j][1])]); }
  var saved = {};
  var sv = tab_(track, '확정').getDataRange().getValues();
  for (var k = 1; k < sv.length; k++) { if (sv[k][0]) saved[String(sv[k][0])] = { supply: Number(sv[k][1]), savedAt: String(sv[k][8] || '') }; }
  var cfg = {};
  var cv = tab_(track, '설정').getDataRange().getValues();
  for (var c = 1; c < cv.length; c++) {
    var key = String(cv[c][0] || '');
    if (!/^[A-Z_0-9]+$/.test(key)) continue;
    try { cfg[key] = JSON.parse(cv[c][1]); } catch (x) { cfg[key] = cv[c][1]; }
  }
  return { ok: true, track: track, map: map, hist: hist, saved: saved, cfg: cfg };
}

function saveMapping_(track, items) {
  var sh = tab_(track, '매핑');
  var vals = sh.getDataRange().getValues(), idx = {};
  for (var i = 1; i < vals.length; i++) idx[norm_(vals[i][0])] = i + 1;
  var added = 0, updated = 0;
  items.forEach(function (it) {
    var key = norm_(it.name), member = String(it.member || '').trim();
    if (!key || !member) return;
    if (idx[key]) { sh.getRange(idx[key], 2).setValue(member); updated++; }
    else { sh.appendRow([key, member]); idx[key] = sh.getLastRow(); added++; }
  });
  return { ok: true, added: added, updated: updated };
}

function saveQuarter_(track, r) {
  if (!r.quarter) return { ok: false, error: 'missing_quarter' };
  var sh = tab_(track, '확정');
  var row = [r.quarter, r.supply, r.vat, r.gross, r.settleMonth || '', r.txDate || '', r.payDate || '', r.file || '',
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'), r.roundNote || '',
    JSON.stringify(r.lines || []), JSON.stringify(r.royaltyTable || [])];
  var vals = sh.getDataRange().getValues(), at = -1;
  for (var i = 1; i < vals.length; i++) if (String(vals[i][0]) === String(r.quarter)) at = i + 1;
  if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  // 확정한 분기는 이력에도 반영(이미 있으면 금액 갱신)
  var hs = tab_(track, '이력'), hv = hs.getDataRange().getValues(), hat = -1;
  for (var j = 1; j < hv.length; j++) if (String(hv[j][0]) === String(r.quarter)) hat = j + 1;
  if (hat > 0) hs.getRange(hat, 2, 1, 2).setValues([[r.supply, '사이트 확정 저장']]);
  else hs.appendRow([r.quarter, r.supply, '사이트 확정 저장']);
  return { ok: true };
}

// 권한 승인·연결 확인용 — 스크립트 편집기에서 ▶ 실행(웹 요청에서는 호출되지 않음)
function manualAuthTest() {
  Object.keys(TRACKS).forEach(function (t) {
    var c = getConfig_(t);
    Logger.log(TRACKS[t] + ': 매핑 ' + Object.keys(c.map).length + '건, 이력 ' + c.hist.length + '분기, 설정 키: ' + Object.keys(c.cfg).join(','));
  });
}
