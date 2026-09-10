// ================================================================
// 경북 상주 국내아웃리치 (2026.09.12) — 출석/방문 체크 백엔드
// ★ 사용법: script.google.com 에서 새 프로젝트 생성 → 이 파일 내용 전체 붙여넣기
//    → 아래 seedRoster() 함수를 한 번 실행하여 초대자 명단 시트 생성
//    → 웹앱으로 배포 (자세한 내용은 배포안내.txt 참고)
// ================================================================

var SHEET_ROSTER = '초대자명단';

// 방문 체크가 필요한 9개 팀 (안내팀=출석, 나머지 8팀=방문이력)
var TEAMS = ['안내팀', '오병이어팀', '꽃단장팀', '발지압팀', '추나요법팀', '소망카페팀', '소망놀이팀', '촬영팀', '소망택배팀'];

var _cb = null; // JSONP callback 이름

// ── GET/POST 공통 처리 ────────────────────────────────────────
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || '';
  _cb = params.callback || null;

  if (action === 'getRoster')   return handleGetRoster();
  if (action === 'setCheck')    return handleSetCheck(params);
  if (action === 'addPerson')   return handleAddPerson(params);

  return jsonRes({ success: false, error: 'unknown action' });
}

function doPost(e) {
  var params = {};
  try { params = JSON.parse(e.postData.contents); } catch (err) { params = e.parameter || {}; }
  return doGet({ parameter: params });
}

// ── 시트 준비 (없으면 헤더 생성) ──────────────────────────────
function ensureSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ROSTER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ROSTER);
    var header = ['순번', '이름', '나이', '비고'];
    TEAMS.forEach(function (t) { header.push(t + '_체크'); header.push(t + '_시각'); });
    sheet.appendRow(header);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function teamColIndex_(team) {
  var idx = TEAMS.indexOf(team);
  if (idx < 0) return null;
  // 1~4열: 순번,이름,나이,비고 / 이후 팀당 2열(체크,시각)
  return { checkCol: 5 + idx * 2, timeCol: 6 + idx * 2 };
}

// ── 명단 + 체크 현황 불러오기 (공개) ──────────────────────────
function handleGetRoster() {
  var sheet = ensureSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonRes({ success: true, teams: TEAMS, people: [] });

  var lastCol = 4 + TEAMS.length * 2;
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var people = data.map(function (row, i) {
    var checks = {};
    TEAMS.forEach(function (t, idx) {
      var c = row[4 + idx * 2];
      var tm = row[5 + idx * 2];
      checks[t] = { checked: c === true || c === 'TRUE' || c === '체크', time: tm ? String(tm) : '' };
    });
    return {
      row: i + 2,
      no: row[0],
      name: String(row[1] || ''),
      age: row[2],
      note: String(row[3] || ''),
      checks: checks
    };
  }).filter(function (p) { return p.name; });

  return jsonRes({ success: true, teams: TEAMS, people: people });
}

// ── 체크 토글 (공개, 팀별 출석/방문 체크) ─────────────────────
function handleSetCheck(params) {
  var row = parseInt(params.row || '0', 10);
  var team = String(params.team || '').trim();
  var checked = String(params.checked) === 'true';
  if (!row || row < 2) return jsonRes({ success: false, error: '잘못된 행 번호' });
  var col = teamColIndex_(team);
  if (!col) return jsonRes({ success: false, error: '알 수 없는 팀' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = ensureSheet_();
    sheet.getRange(row, col.checkCol).setValue(checked);
    sheet.getRange(row, col.timeCol).setValue(checked ? new Date().toLocaleString('ko-KR') : '');
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true });
}

// ── 현장 등록(명단에 없는 방문객 추가) ────────────────────────
function handleAddPerson(params) {
  var name = String(params.name || '').trim();
  var age = String(params.age || '').trim();
  var note = String(params.note || '현장등록').trim();
  if (!name) return jsonRes({ success: false, error: '이름 필요' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = ensureSheet_();
    var newNo = sheet.getLastRow(); // 헤더 제외 인원수+1
    var row = [newNo, name, age, note];
    TEAMS.forEach(function () { row.push(false); row.push(''); });
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true });
}

// ── 초기 명단 시딩 (Apps Script 편집기에서 최초 1회 직접 실행) ─
// 9.초대 페이지(56명, 마을주민) 원본 명단 기준. [확인필요] 표시는
// 인쇄 원본이 흐려 정확히 판독하지 못한 항목이니 현장 명단과 대조해 주세요.
function seedRoster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName(SHEET_ROSTER);
  if (old) ss.deleteSheet(old); // 재시딩 시 초기화
  var sheet = ensureSheet_();

  var raw = [
    [1, '반재복', 65, '목회자'], [2, '김종심', 64, '목회자(사모)'], [3, '강정안', 82, '목회자'],
    [4, '김병연', 71, '목회자'], [5, '김종욱', 78, '목회자'], [6, '안희철', 78, '목회자'],
    [7, '안석현', 64, '현재 이장'], [8, '박미신', 64, '(사모)'], [9, '권욱련', 82, ''],
    [10, '김경숙', 80, ''], [11, '김경자', 68, ''], [12, '김기환', 55, ''],
    [13, '김두이', 83, ''], [14, '김숙열', 89, ''], [15, '김순옥', 87, ''],
    [16, '김월일', 92, ''], [17, '배옥선', 85, ''], [18, '손동원', 73, ''],
    [19, '[확인필요]', '', '원본 인쇄 흐림 - 대조 필요'],
    [20, '손수림', 85, ''], [21, '신동식', 83, ''], [22, '심영규', 60, ''],
    [23, '안규태', 80, ''], [24, '안달오', 85, ''], [25, '안상식', 55, ''],
    [26, '안선기', 82, ''], [27, '안인태', 85, ''], [28, '안종국', 62, ''],
    [29, '안종노', 90, ''], [30, '안종준', 76, ''], [31, '안한주', 55, ''],
    [32, '안호준', 82, ''], [33, '엄금자', 92, ''], [34, '엄용순', 90, ''],
    [35, '윤수전', 90, ''], [36, '이국희', 74, ''], [37, '이덕희', 69, ''],
    [38, '이명희', 52, ''], [39, '이봉자', 83, ''], [40, '이상순', 72, ''],
    [41, '이상애', 80, ''], [42, '이상철', 85, ''], [43, '이영희', 75, ''],
    [44, '이완주', 64, ''], [45, '이윤정', 55, ''], [46, '이정상', 90, ''],
    [47, '이준옥', 85, ''], [48, '이학선', 68, ''], [49, '이차자', 86, ''],
    [50, '이춘옥', 67, ''], [51, '임성진', 73, ''], [52, '장귀자', 70, ''],
    [53, '전상희', 64, ''], [54, '정정숙', 70, ''], [55, '최명숙', 64, ''],
    [56, '최미경', 58, '']
  ];

  var rows = raw.map(function (r) {
    var row = r.slice();
    TEAMS.forEach(function () { row.push(false); row.push(''); });
    return row;
  });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

// ── 공통: JSON 응답 (JSONP 지원) ──────────────────────────────
function jsonRes(obj) {
  var json = JSON.stringify(obj);
  if (_cb) {
    return ContentService.createTextOutput(_cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
