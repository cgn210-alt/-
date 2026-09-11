// ================================================================
// 경북 상주 국내아웃리치 (2026.09.12) — 출석/방문 체크 백엔드
// ★ 사용법: script.google.com 에서 새 프로젝트 생성 → 이 파일 내용 전체 붙여넣기
//    → 아래 seedRoster() 함수를 한 번 실행하여 초대자 명단 시트 생성
//    → 웹앱으로 배포 (자세한 내용은 배포안내.txt 참고)
// ================================================================

var SHEET_ROSTER = '초대자명단';

// 방문 체크가 필요한 10개 팀 (안내팀=출석, 나머지 9팀=팀별 체험 방문이력)
var TEAMS = ['안내팀', '오병이어팀', '꽃단장팀', '발지압팀', '추나요법팀', '소망카페팀', '소망놀이팀', '보수팀', '촬영팀', '소망택배팀'];

// 인원배치(버스/팀별) 관리자 비밀번호 — 필요시 이 값을 직접 원하는 비밀번호로 바꾸세요.
var ADMIN_PW = '상주2026';
var SHEET_BUSINFO = '버스정보';
var SHEET_BUSMEMBER = '버스인원';
var SHEET_TEAMMEMBER = '팀별인원';
var SHEET_TEAMROLE = '팀별역할';

// 활동체크(출석·방문) 팀장 전용 비밀번호 — 관리자 비밀번호와는 별개의 가벼운 비밀번호입니다.
// 이 비밀번호를 아는 사람만 아래 CHECK_RESTRICTED_TEAMS 팀의 체크를 할 수 있습니다.
var CHECK_PW = '팀장2026';
var CHECK_RESTRICTED_TEAMS = ['안내팀', '오병이어팀', '꽃단장팀', '발지압팀', '추나요법팀', '소망카페팀', '소망놀이팀', '보수팀', '촬영팀', '소망택배팀'];

var _cb = null; // JSONP callback 이름

// ── GET/POST 공통 처리 ────────────────────────────────────────
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || '';
  _cb = params.callback || null;

  if (action === 'getRoster')   return handleGetRoster();
  if (action === 'setCheck')    return handleSetCheck(params);
  if (action === 'addPerson')   return handleAddPerson(params);
  if (action === 'checkPwVerify') return jsonRes({ success: checkTeamPw_(params.pw) });
  if (action === 'adminEditPerson')   return handleAdminEditPerson(params);
  if (action === 'adminDeletePerson') return handleAdminDeletePerson(params);

  if (action === 'getDeployment')     return handleGetDeployment();
  if (action === 'adminCheckPw')      return jsonRes({ success: checkAdmin_(params.pw) });
  if (action === 'adminSaveBusInfo')  return handleAdminSaveBusInfo(params);
  if (action === 'adminAddMember')    return handleAdminAddMember(params);
  if (action === 'adminRemoveMember') return handleAdminRemoveMember(params);
  if (action === 'adminSetColor')     return handleAdminSetColor(params);
  if (action === 'adminSetCompanion') return handleAdminSetCompanion(params);
  if (action === 'getRoles')          return handleGetRoles();
  if (action === 'adminEditRole')     return handleAdminEditRole(params);

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
// 활동체크 팀장 전용 비밀번호(또는 관리자 비밀번호) 확인
function checkTeamPw_(pw) {
  return String(pw) === CHECK_PW || checkAdmin_(pw);
}

function handleSetCheck(params) {
  var row = parseInt(params.row || '0', 10);
  var team = String(params.team || '').trim();
  var checked = String(params.checked) === 'true';
  if (!row || row < 2) return jsonRes({ success: false, error: '잘못된 행 번호' });
  var col = teamColIndex_(team);
  if (!col) return jsonRes({ success: false, error: '알 수 없는 팀' });
  if (CHECK_RESTRICTED_TEAMS.indexOf(team) >= 0 && !checkTeamPw_(params.checkPw)) {
    return jsonRes({ success: false, error: '이 팀은 팀장만 체크할 수 있습니다. 팀장 체크 권한을 먼저 확인해 주세요.' });
  }

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

// ── 관리자: 초대자명단 오타 수정(이름·나이·비고) ──────────────
function handleAdminEditPerson(params) {
  if (!checkAdmin_(params.pw)) return jsonRes({ success: false, error: '비밀번호가 올바르지 않습니다.' });
  var row = parseInt(params.row || '0', 10);
  if (!row || row < 2) return jsonRes({ success: false, error: '잘못된 행 번호' });
  var name = String(params.name || '').trim();
  if (!name) return jsonRes({ success: false, error: '이름 필요' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = ensureSheet_();
    sheet.getRange(row, 2, 1, 3).setValues([[name, params.age || '', params.note || '']]);
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true });
}

// ── 관리자: 초대자명단에서 삭제 ────────────────────────────────
function handleAdminDeletePerson(params) {
  if (!checkAdmin_(params.pw)) return jsonRes({ success: false, error: '비밀번호가 올바르지 않습니다.' });
  var row = parseInt(params.row || '0', 10);
  if (!row || row < 2) return jsonRes({ success: false, error: '잘못된 행 번호' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = ensureSheet_();
    sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true });
}

// ── 초기 명단 시딩 (Apps Script 편집기에서 최초 1회 직접 실행) ─
// 9.초대 페이지(56명, 마을주민) 원본 명단 기준. [확인필요] 표시는
// 인쇄 원본이 흐려 정확히 판독하지 못한 항목이니 현장 명단과 대조해 주세요.
// ★ 보수팀 체크 항목이 새로 추가되어 TEAMS 배열의 팀 구성이 바뀌었습니다.
//   이미 "초대자명단" 시트를 만든 적이 있다면(과거에 seedRoster를 실행한 적
//   있다면) 이 함수를 다시 한 번 실행해야 열(컬럼) 배치가 새 TEAMS와 맞습니다.
//   ⚠ 재실행하면 기존 체크 기록은 모두 초기화됩니다 — 실제 출석/방문 체크를
//     이미 시작하셨다면 재실행 전에 꼭 확인해 주세요.
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

// ================================================================
// 인원배치(버스배치 / 팀별 인원배치) — 관리자 추가·삭제·색상 변경
// ================================================================

function checkAdmin_(pw) { return String(pw) === ADMIN_PW; }

function ensureBusInfoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_BUSINFO);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_BUSINFO);
    sheet.appendRow(['호차', '총인원', '기사명', '기사연락처', '차량번호', '차장', '수령물품']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function ensureBusMemberSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_BUSMEMBER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_BUSMEMBER);
    sheet.appendRow(['호차', '팀명', '이름', '팀장여부', '색상']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function ensureTeamMemberSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TEAMMEMBER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TEAMMEMBER);
    sheet.appendRow(['팀명', '이름', '팀장여부', '색상', '동행지기여부']);
    sheet.setFrozenRows(1);
  } else if (String(sheet.getRange(1, 5).getValue()) !== '동행지기여부') {
    // 기존에 이미 만들어진 시트라면 5번째 열(동행지기여부) 헤더만 보충
    sheet.getRange(1, 5).setValue('동행지기여부');
  }
  return sheet;
}

// 시트 2행부터의 값들을 {row, v:[...]} 형태 배열로 반환 (첫 칸이 빈 행은 제외)
function sheetRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] !== '' && values[i][0] !== null) out.push({ row: i + 2, v: values[i] });
  }
  return out;
}

// ── 공개: 버스배치 + 팀별 인원배치 전체 불러오기 ─────────────
function handleGetDeployment() {
  var busInfo = sheetRows_(ensureBusInfoSheet_()).map(function (r) {
    return { row: r.row, bus: r.v[0], total: r.v[1], driverName: r.v[2], driverPhone: r.v[3], plate: r.v[4], conductor: r.v[5], items: r.v[6] };
  });
  var busMembers = sheetRows_(ensureBusMemberSheet_()).map(function (r) {
    return { row: r.row, bus: r.v[0], team: r.v[1], name: r.v[2], leader: r.v[3] === true || r.v[3] === 'TRUE', color: r.v[4] || '' };
  });
  var teamMembers = sheetRows_(ensureTeamMemberSheet_()).map(function (r) {
    return { row: r.row, team: r.v[0], name: r.v[1], leader: r.v[2] === true || r.v[2] === 'TRUE', color: r.v[3] || '', companion: r.v[4] === true || r.v[4] === 'TRUE' };
  });
  return jsonRes({ success: true, busInfo: busInfo, busMembers: busMembers, teamMembers: teamMembers });
}

// ── 관리자: 버스 정보 저장(있으면 갱신, 없으면 추가) ─────────
function handleAdminSaveBusInfo(p) {
  if (!checkAdmin_(p.pw)) return jsonRes({ success: false, error: '비밀번호가 올바르지 않습니다.' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = ensureBusInfoSheet_();
    var bus = String(p.bus || '');
    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === bus) { targetRow = i + 1; break; } }
    var row = [bus, p.total || '', p.driverName || '', p.driverPhone || '', p.plate || '', p.conductor || '', p.items || ''];
    if (targetRow > 0) sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true });
}

// ── 관리자: 인원 추가 (target: 'bus' | 'team') ───────────────
// 새로 추가된 행 번호를 함께 반환합니다 — 클라이언트가 전체 명단을
// 다시 불러오지 않고 그 자리에서 바로 반영할 수 있어 응답이 빠릅니다.
function handleAdminAddMember(p) {
  if (!checkAdmin_(p.pw)) return jsonRes({ success: false, error: '비밀번호가 올바르지 않습니다.' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var newRow;
  try {
    if (p.target === 'bus') {
      var sheet = ensureBusMemberSheet_();
      sheet.appendRow([p.bus || '', p.team || '', p.name || '', p.leader === 'true', p.color || '']);
      newRow = sheet.getLastRow();
    } else {
      var sheet2 = ensureTeamMemberSheet_();
      sheet2.appendRow([p.team || '', p.name || '', p.leader === 'true', p.color || '']);
      newRow = sheet2.getLastRow();
    }
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true, row: newRow });
}

// ── 관리자: 인원 삭제 ────────────────────────────────────────
function handleAdminRemoveMember(p) {
  if (!checkAdmin_(p.pw)) return jsonRes({ success: false, error: '비밀번호가 올바르지 않습니다.' });
  var row = parseInt(p.row || '0', 10);
  if (!row || row < 2) return jsonRes({ success: false, error: '잘못된 행 번호' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = (p.target === 'bus') ? ensureBusMemberSheet_() : ensureTeamMemberSheet_();
    sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true });
}

// ── 관리자: 이름 폰트 색상 변경 ──────────────────────────────
function handleAdminSetColor(p) {
  if (!checkAdmin_(p.pw)) return jsonRes({ success: false, error: '비밀번호가 올바르지 않습니다.' });
  var row = parseInt(p.row || '0', 10);
  if (!row || row < 2) return jsonRes({ success: false, error: '잘못된 행 번호' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = (p.target === 'bus') ? ensureBusMemberSheet_() : ensureTeamMemberSheet_();
    var colIndex = (p.target === 'bus') ? 5 : 4;
    sheet.getRange(row, colIndex).setValue(p.color || '');
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true });
}

// ── 관리자: 동행지기 지정/해제 (팀별인원 전용 — 같은 팀 안에서
//   팀원들의 사역시간·간식제공 등을 챙기는 담당자를 표시) ────────
function handleAdminSetCompanion(p) {
  if (!checkAdmin_(p.pw)) return jsonRes({ success: false, error: '비밀번호가 올바르지 않습니다.' });
  var row = parseInt(p.row || '0', 10);
  if (!row || row < 2) return jsonRes({ success: false, error: '잘못된 행 번호' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = ensureTeamMemberSheet_();
    sheet.getRange(row, 5).setValue(String(p.value) === 'true');
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true });
}

// ================================================================
// 팀별역할("팀별역할" 탭) — 관리자 수정 가능, 처음 열 때 기본 내용으로 자동 생성
// ================================================================
function ensureTeamRoleSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TEAMROLE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TEAMROLE);
    sheet.appendRow(['팀명', '역할', '세부내용', '도형', '색상']);
    sheet.setFrozenRows(1);
    var defaults = [
      ['본부팀', '전체 운영, 배치, 음향, 서포팅 스텝', '구급약품, 지원물품, 각 Zone 명패제작, 안전위원, 블루투스 스피커 2개', '', ''],
      ['안내팀', '안내, 등록, 이름표', '어르신 맞이, Zone 소개, 초청 목회자 전담 케어, 이름표(라벨)', '', ''],
      ['오병이어팀', '바베큐, 식사', '야외 천막에서 식사, 숯불구이, 짝꿍 식사', 'circle', '#f2c230'],
      ['꽃단장팀', '염색, 컷트, 머리감기기, 네일아트', '온수, 물뿌리개, 소요시간 1인(염색 1시간 / 머리행굼 10분 / 컷트 30분 / 네일아트 30분)', 'heart', '#e0405c'],
      ['발지압팀', '세족 및 지압', '침대 2개, 양말, 수건, 세수대야, 소요시간 1인(세족 및 지압 10분)', 'circle', '#f48fb1'],
      ['추나요법팀', '교정, 통증 완화', '장소: 교회(어르신들 교회 체험), 소요시간 1인(20분)/1명의 치료사가 10명 교정 가능', 'circle', '#2e9e5b'],
      ['소망카페팀', '음료, 차, 다과, 사진', '맞춤 음료, 다과, 아이스크림 제공, 카페 데코, 신발정리 집게', '', ''],
      ['소망놀이팀', '어르신 맞춤', '스크래치 카드, 명랑게임', 'star', '#2f80ed'],
      ['보수팀', '가정 시설 수리', '시설수리 3건', '', ''],
      ['촬영팀', '동영상 및 사진', '전체 사역에 대한 영상 및 스냅 사진 촬영', 'circle', '#2f80ed'],
      ['소망택배팀', '개별 복음 제시', '맞춤 전도', 'circle', '#e03131']
    ];
    sheet.getRange(2, 1, defaults.length, 5).setValues(defaults);
  }
  return sheet;
}

// ── 공개: 팀별역할 불러오기 ────────────────────────────────────
function handleGetRoles() {
  var roles = sheetRows_(ensureTeamRoleSheet_()).map(function (r) {
    return { row: r.row, team: r.v[0], role: r.v[1] || '', detail: r.v[2] || '', shape: r.v[3] || '', color: r.v[4] || '' };
  });
  return jsonRes({ success: true, roles: roles });
}

// ── 관리자: 팀별역할 수정 ──────────────────────────────────────
function handleAdminEditRole(p) {
  if (!checkAdmin_(p.pw)) return jsonRes({ success: false, error: '비밀번호가 올바르지 않습니다.' });
  var row = parseInt(p.row || '0', 10);
  if (!row || row < 2) return jsonRes({ success: false, error: '잘못된 행 번호' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = ensureTeamRoleSheet_();
    sheet.getRange(row, 2, 1, 4).setValues([[p.role || '', p.detail || '', p.shape || '', p.color || '']]);
  } finally {
    lock.releaseLock();
  }
  return jsonRes({ success: true });
}

// ── 초기 인원배치 시딩 (Apps Script 편집기에서 최초 1회 직접 실행) ─
// 제공해주신 버스배치표·팀별 인원배치표 이미지를 기준으로 입력했습니다.
// 색상(오렌지·녹색 등)은 이미지에서 정확히 판독하지 못해 우선 기본색(빈 값)으로
// 두었으니, 배포 후 관리자 페이지에서 직접 지정해 주세요.
// [확인필요] 표시 항목은 두 이미지(버스별/팀별) 간 인원수가 서로 맞지 않아
// 원본과 대조가 필요합니다.
function seedDeployment() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['버스정보', '버스인원', '팀별인원'].forEach(function (name) {
    var old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old);
  });
  var busInfoSheet = ensureBusInfoSheet_();
  var busMemberSheet = ensureBusMemberSheet_();
  var teamMemberSheet = ensureTeamMemberSheet_();

  // ── 버스 정보 ──
  busInfoSheet.getRange(2, 1, 2, 7).setValues([
    // ⚠ 기사님 연락처는 개인정보 보호를 위해 이 코드(공개 저장소)에는 넣지 않았습니다.
    //   배포 후 앱의 "인원배치 → 관리자 모드 → 버스 정보 수정"에서 직접 입력해 주세요.
    ['1호차', 37, '이상문', '', '서울72바6384', '주웅현', '김밥, 과일, 생수, 이름표 라벨'],
    ['2호차', 42, '최국호', '', '서울72바6382', '정두식', '김밥, 과일, 생수, 이름표 라벨'],
  ]);

  // ── 버스별 탑승 인원(팀 단위) ── [leader, name, team, bus]
  var busRaw = [
    // 1호차 (37명)
    ['1호차', '리더십', '노치형 목사', true], ['1호차', '리더십', '김의래 장로', false],
    ['1호차', '안내팀', '김문경', true], ['1호차', '안내팀', '박성애', false], ['1호차', '안내팀', '정형지', false], ['1호차', '안내팀', '정세정', false],
    ['1호차', '오병이어팀', '이보윤', true],
    ['1호차', '오병이어팀', '구자영', false], ['1호차', '오병이어팀', '김성희', false], ['1호차', '오병이어팀', '김종중', false], ['1호차', '오병이어팀', '김희정(B)', false],
    ['1호차', '오병이어팀', '이상문', false], ['1호차', '오병이어팀', '이성우', false], ['1호차', '오병이어팀', '이창권', false], ['1호차', '오병이어팀', '유덕영', false],
    ['1호차', '오병이어팀', '장은자', false], ['1호차', '오병이어팀', '장혜영', false], ['1호차', '오병이어팀', '정기순', false], ['1호차', '오병이어팀', '정순영', false],
    ['1호차', '오병이어팀', '조재관', false], ['1호차', '오병이어팀', '지형근', false], ['1호차', '오병이어팀', '최재권', false], ['1호차', '오병이어팀', '홍순전', false],
    ['1호차', '소망카페팀', '김남정', true],
    ['1호차', '소망카페팀', '김영옥', false], ['1호차', '소망카페팀', '박선영', false], ['1호차', '소망카페팀', '박은규', false], ['1호차', '소망카페팀', '서정욱', false],
    ['1호차', '소망카페팀', '이용경', false], ['1호차', '소망카페팀', '이일순', false], ['1호차', '소망카페팀', '이희자', false], ['1호차', '소망카페팀', '최황', false],
    ['1호차', '소망놀이팀', '주웅현', true], ['1호차', '소망놀이팀', '정중균', false],
    ['1호차', '촬영팀', '주찬혁', true], ['1호차', '촬영팀', '김진국', false], ['1호차', '촬영팀', '이지연', false],
    // 2호차 (42명)
    ['2호차', '리더십', '오명 장로', false],
    ['2호차', '꽃단장팀', '황한나', true],
    ['2호차', '꽃단장팀', '강미란', false], ['2호차', '꽃단장팀', '김지은', false], ['2호차', '꽃단장팀', '김희정(A)', false], ['2호차', '꽃단장팀', '박애우', false],
    ['2호차', '꽃단장팀', '박주연', false], ['2호차', '꽃단장팀', '배은하', false], ['2호차', '꽃단장팀', '복유선', false], ['2호차', '꽃단장팀', '변지숙', false],
    ['2호차', '꽃단장팀', '신유순', false], ['2호차', '꽃단장팀', '안희철', false], ['2호차', '꽃단장팀', '오영옥', false], ['2호차', '꽃단장팀', '이언옥', false],
    ['2호차', '꽃단장팀', '이연주', false], ['2호차', '꽃단장팀', '이영원(본)', false], ['2호차', '꽃단장팀', '이창배', false], ['2호차', '꽃단장팀', '임동욱', false],
    ['2호차', '꽃단장팀', '장윤지', false], ['2호차', '꽃단장팀', '장은자(오)', false], ['2호차', '꽃단장팀', '정경희', false], ['2호차', '꽃단장팀', '진선희', false],
    ['2호차', '꽃단장팀', '최인경 [확인필요-인원수 21명 대비 명단 확인]', false],
    ['2호차', '발지압팀', '안승립', true],
    ['2호차', '발지압팀', '김현식', false], ['2호차', '발지압팀', '남관우', false], ['2호차', '발지압팀', '박지수', false], ['2호차', '발지압팀', '임소미', false],
    ['2호차', '추나요법팀', '엄재준', true], ['2호차', '추나요법팀', '김민종', false], ['2호차', '추나요법팀', '서병이 (자차이용)', false],
  ];
  // 컬럼 순서: [호차, 팀명, 이름, 팀장여부, 색상]
  var busRows = busRaw.map(function (r) { return [r[0], r[1], r[2], r[3], '']; });
  busMemberSheet.getRange(2, 1, busRows.length, 5).setValues(busRows);

  // ── 팀별 인원배치(전체, 구역명 없이 팀명 기준) ──
  // ※ 촬영팀/보수팀/소망택배팀 등 일부는 버스별 목록에 없어 원본 구역별표 기준으로만 입력했습니다.
  var teamRaw = [
    ['본부팀', '정두식', true], ['본부팀', '남관우', false], ['본부팀', '서정욱', false], ['본부팀', '이영미', false], ['본부팀', '주웅현', false],
    ['안내팀', '김문경', true], ['안내팀', '박성애', false], ['안내팀', '정형지', false], ['안내팀', '정세정', false], ['안내팀', '[확인필요-인원 5명 대비 1명 부족]', false],
    ['오병이어팀', '이보윤', true], ['오병이어팀', '구자영', false], ['오병이어팀', '김성희', false], ['오병이어팀', '김종중', false], ['오병이어팀', '김희정(B)', false],
    ['오병이어팀', '이상문', false], ['오병이어팀', '이성우', false], ['오병이어팀', '이창권', false], ['오병이어팀', '유덕영', false], ['오병이어팀', '장은자', false],
    ['오병이어팀', '장혜영', false], ['오병이어팀', '정기순', false], ['오병이어팀', '정순영', false], ['오병이어팀', '조재관', false], ['오병이어팀', '지형근', false],
    ['오병이어팀', '최재권', false], ['오병이어팀', '홍순전', false], ['오병이어팀', '[확인필요-인원 18명 대비 1명 부족]', false],
    ['꽃단장팀', '황한나', true], ['꽃단장팀', '강미란', false], ['꽃단장팀', '김지은', false], ['꽃단장팀', '김희정(A)', false], ['꽃단장팀', '박애우', false],
    ['꽃단장팀', '박주연', false], ['꽃단장팀', '배은하', false], ['꽃단장팀', '복유선', false], ['꽃단장팀', '변지숙', false], ['꽃단장팀', '신유순', false],
    ['꽃단장팀', '안희철', false], ['꽃단장팀', '오영옥', false], ['꽃단장팀', '이언옥', false], ['꽃단장팀', '이연주', false], ['꽃단장팀', '이영원(본)', false],
    ['꽃단장팀', '이창배', false], ['꽃단장팀', '임동욱', false], ['꽃단장팀', '장윤지', false], ['꽃단장팀', '장은자(오)', false], ['꽃단장팀', '정경희', false],
    ['꽃단장팀', '진선희', false], ['꽃단장팀', '최인경', false],
    ['발지압팀', '안승립', true], ['발지압팀', '김현식', false], ['발지압팀', '남관우', false], ['발지압팀', '박지수', false], ['발지압팀', '임소미', false],
    ['추나요법팀', '엄재준', true], ['추나요법팀', '김민종', false], ['추나요법팀', '서병이 (자차이용)', false],
    ['소망카페팀', '김남정', true], ['소망카페팀', '김영옥', false], ['소망카페팀', '박선영', false], ['소망카페팀', '박은규', false], ['소망카페팀', '서정욱', false],
    ['소망카페팀', '이용경', false], ['소망카페팀', '이일순', false], ['소망카페팀', '이희자', false], ['소망카페팀', '최황', false],
    ['소망놀이팀', '주웅현', true], ['소망놀이팀', '정중균', false],
    ['보수팀', '김윤미', true], ['보수팀', '박제하', false], ['보수팀', '[확인필요-원본 대조 필요]', false], ['보수팀', '[확인필요-원본 대조 필요]', false], ['보수팀', '[확인필요-원본 대조 필요]', false],
    ['촬영팀', '주찬혁', true], ['촬영팀', '김진국', false], ['촬영팀', '이지연', false],
    ['소망택배팀', '이지영', true], ['소망택배팀', '[확인필요-원본 대조 필요]', false], ['소망택배팀', '[확인필요-원본 대조 필요]', false],
    ['소망택배팀', '[확인필요-원본 대조 필요]', false], ['소망택배팀', '[확인필요-원본 대조 필요]', false], ['소망택배팀', '[확인필요-원본 대조 필요]', false],
  ];
  var teamRows = teamRaw.map(function (r) { return [r[0], r[1], r[2], '']; });
  teamMemberSheet.getRange(2, 1, teamRows.length, 4).setValues(teamRows);
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
