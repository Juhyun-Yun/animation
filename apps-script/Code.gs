/************************************************************
 *  우리 반 애니메이션 스튜디오  -  서버 코드 (Google Apps Script)
 *
 *  이 파일은 학생들의 요청을 받아 다음을 처리합니다.
 *   - 학생 명단 불러오기 / 저장
 *   - 애니메이션 작품 저장 (사진 컷들은 구글 드라이브에, 정보는 구글 시트에)
 *   - 하트 / 댓글 / 동료평가 저장
 *
 *  데이터 저장소(스프레드시트 1개 + 드라이브 폴더 1개)는
 *  처음 실행될 때 자동으로 만들어집니다. 선생님은 따로 만들 필요가 없어요.
 ************************************************************/

// 선생님 설정 화면에 들어갈 때 필요한 비밀번호(PIN). 원하시면 숫자를 바꾸세요.
var DEFAULT_TEACHER_PIN = '1234';

/* ---------- 웹앱 진입점: 학생이 링크를 열면 이 함수가 화면을 보여줍니다 ---------- */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('우리 반 애니메이션 스튜디오 🎬')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 다른 HTML 파일(Style, App 등)을 Index 안으로 끼워 넣어주는 도우미
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================================================
 *  저장소 준비 (스프레드시트 + 드라이브 폴더 자동 생성)
 * ========================================================= */
function getStore_() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SS_ID');
  var ss;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('애니메이션 스튜디오 데이터');
    props.setProperty('SS_ID', ss.getId());
    setupSheets_(ss);
  }

  var folderId = props.getProperty('FOLDER_ID');
  var folder;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder('애니메이션 스튜디오 작품');
    props.setProperty('FOLDER_ID', folder.getId());
  }

  if (!props.getProperty('TEACHER_PIN')) {
    props.setProperty('TEACHER_PIN', DEFAULT_TEACHER_PIN);
  }
  return { ss: ss, folder: folder, props: props };
}

function setupSheets_(ss) {
  var first = ss.getSheets()[0];
  first.setName('roster');
  first.getRange(1, 1, 1, 3).setValues([['id', 'number', 'name']]);

  var anim = ss.insertSheet('animations');
  anim.getRange(1, 1, 1, 11).setValues([[
    'id', 'studentId', 'studentName', 'title', 'plan',
    'fps', 'frameCount', 'thumbnail', 'driveFileId', 'selfEval', 'createdAt'
  ]]);

  var cmt = ss.insertSheet('comments');
  cmt.getRange(1, 1, 1, 6).setValues([['id', 'animId', 'studentId', 'studentName', 'text', 'createdAt']]);

  var hrt = ss.insertSheet('hearts');
  hrt.getRange(1, 1, 1, 3).setValues([['animId', 'studentId', 'createdAt']]);

  var pe = ss.insertSheet('peerEvals');
  pe.getRange(1, 1, 1, 5).setValues([['animId', 'studentId', 'studentName', 'ratings', 'createdAt']]);
}

function sheet_(name) {
  return getStore_().ss.getSheetByName(name);
}

// 시트 한 장을 [{컬럼:값}, ...] 형태로 읽기
function readRows_(name) {
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var header = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < header.length; c++) obj[header[c]] = values[r][c];
    out.push(obj);
  }
  return out;
}

/* =========================================================
 *  부팅 정보
 * ========================================================= */
function apiGetConfig() {
  var store = getStore_();
  var roster = readRows_('roster');
  return {
    rosterLoaded: roster.length > 0,
    rosterCount: roster.length,
    dataUrl: store.ss.getUrl()
  };
}

/* =========================================================
 *  학생 명단
 * ========================================================= */
function apiGetRoster() {
  var rows = readRows_('roster');
  return rows.map(function (r) {
    return { id: String(r.id), number: r.number, name: String(r.name) };
  });
}

// 선생님: 명단을 텍스트로 저장 (한 줄에 한 명, "번호,이름" 또는 "이름" 형식)
function apiSetRoster(pin, namesText) {
  requireTeacher_(pin);
  var lines = String(namesText || '').split(/\r?\n/);
  var rows = [];
  var n = 0;
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    n++;
    var num = n, name = line;
    var m = line.split(/[,\t]/);
    if (m.length >= 2 && m[0].trim() !== '') {
      var maybeNum = parseInt(m[0].trim(), 10);
      if (!isNaN(maybeNum)) { num = maybeNum; name = m.slice(1).join(',').trim(); }
    }
    rows.push([Utilities.getUuid(), num, name]);
  }
  if (rows.length === 0) throw new Error('명단이 비어 있어요.');
  writeRoster_(rows);
  return apiGetRoster();
}

// 선생님: 다른 구글 시트 URL에서 명단 가져오기 (B열 또는 첫 글자열에서 이름 읽기)
function apiImportRosterFromSheet(pin, url) {
  requireTeacher_(pin);
  var src;
  try { src = SpreadsheetApp.openByUrl(url); }
  catch (e) { throw new Error('시트를 열 수 없어요. 선생님 계정으로 접근 가능한 시트인지 확인하세요.'); }
  var data = src.getSheets()[0].getDataRange().getValues();
  var rows = [];
  var n = 0;
  for (var i = 0; i < data.length; i++) {
    var cells = data[i];
    // 숫자처럼 보이는 헤더/빈 줄 건너뛰기
    var name = '', num = null;
    for (var c = 0; c < cells.length; c++) {
      var v = String(cells[c]).trim();
      if (v === '') continue;
      if (num === null && /^\d+$/.test(v)) { num = parseInt(v, 10); continue; }
      if (name === '') { name = v; }
    }
    if (!name) continue;
    if (/번호|이름|성명|name/i.test(name) && i === 0) continue; // 헤더 줄 제외
    n++;
    rows.push([Utilities.getUuid(), num === null ? n : num, name]);
  }
  if (rows.length === 0) throw new Error('이름을 찾지 못했어요. 시트 첫 장에 이름이 있는지 확인하세요.');
  writeRoster_(rows);
  return apiGetRoster();
}

function writeRoster_(rows) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_('roster');
    sh.clearContents();
    sh.getRange(1, 1, 1, 3).setValues([['id', 'number', 'name']]);
    if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
 *  선생님 인증
 * ========================================================= */
function apiVerifyPin(pin) {
  var stored = getStore_().props.getProperty('TEACHER_PIN') || DEFAULT_TEACHER_PIN;
  return String(pin) === String(stored);
}
function apiChangePin(oldPin, newPin) {
  requireTeacher_(oldPin);
  if (!newPin || String(newPin).length < 4) throw new Error('새 비밀번호는 4자리 이상이어야 해요.');
  getStore_().props.setProperty('TEACHER_PIN', String(newPin));
  return true;
}
function requireTeacher_(pin) {
  if (!apiVerifyPin(pin)) throw new Error('선생님 비밀번호가 맞지 않아요.');
}

/* =========================================================
 *  애니메이션 저장 / 목록 / 상세
 * ========================================================= */
function apiSaveAnimation(payload) {
  if (!payload || !payload.frames || payload.frames.length < 2) {
    throw new Error('컷이 2장 이상 있어야 영상을 만들 수 있어요.');
  }
  var store = getStore_();
  var id = Utilities.getUuid();

  // 사진 컷들은 드라이브에 JSON 파일로 저장 (용량이 크기 때문)
  var fileBody = JSON.stringify({ fps: payload.fps, frames: payload.frames });
  var file = store.folder.createFile(id + '.json', fileBody, 'application/json');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sheet_('animations').appendRow([
      id,
      String(payload.studentId || ''),
      String(payload.studentName || ''),
      String(payload.title || '제목 없음'),
      String(payload.plan || ''),
      Number(payload.fps) || 6,
      payload.frames.length,
      String(payload.thumbnail || ''),
      file.getId(),
      JSON.stringify(payload.selfEval || {}),
      new Date().getTime()
    ]);
  } finally {
    lock.releaseLock();
  }
  return { id: id };
}

// 보드 목록 (하트 많은 순으로 정렬해서 반환 — 사진 컷은 제외, 썸네일만)
function apiListAnimations() {
  var anims = readRows_('animations');
  var heartCounts = countHearts_();
  var list = anims.map(function (a) {
    return {
      id: String(a.id),
      studentId: String(a.studentId),
      studentName: String(a.studentName),
      title: String(a.title),
      fps: Number(a.fps) || 6,
      frameCount: Number(a.frameCount) || 0,
      thumbnail: String(a.thumbnail || ''),
      hearts: heartCounts[a.id] || 0,
      createdAt: Number(a.createdAt) || 0
    };
  });
  // 하트 많은 순 → 같으면 최신순
  list.sort(function (x, y) {
    if (y.hearts !== x.hearts) return y.hearts - x.hearts;
    return y.createdAt - x.createdAt;
  });
  return list;
}

// 한 작품의 상세 정보 (사진 컷 포함) + 댓글 + 동료평가 + 내 하트 여부
function apiGetAnimationDetail(animId, viewerId) {
  var anims = readRows_('animations');
  var meta = null;
  for (var i = 0; i < anims.length; i++) {
    if (String(anims[i].id) === String(animId)) { meta = anims[i]; break; }
  }
  if (!meta) throw new Error('작품을 찾을 수 없어요.');

  var frames = [], fps = Number(meta.fps) || 6;
  try {
    var file = DriveApp.getFileById(meta.driveFileId);
    var data = JSON.parse(file.getBlob().getDataAsString());
    frames = data.frames || [];
    fps = data.fps || fps;
  } catch (e) { /* 파일이 없으면 빈 프레임 */ }

  var comments = readRows_('comments')
    .filter(function (c) { return String(c.animId) === String(animId); })
    .map(function (c) {
      return { id: String(c.id), studentName: String(c.studentName), text: String(c.text), createdAt: Number(c.createdAt) };
    })
    .sort(function (a, b) { return a.createdAt - b.createdAt; });

  var hearts = readRows_('hearts').filter(function (h) { return String(h.animId) === String(animId); });
  var hearted = false;
  for (var j = 0; j < hearts.length; j++) {
    if (String(hearts[j].studentId) === String(viewerId)) { hearted = true; break; }
  }

  var peer = readRows_('peerEvals').filter(function (p) { return String(p.animId) === String(animId); });
  var peerSummary = summarizePeer_(peer);
  var myPeerEval = null;
  for (var k = 0; k < peer.length; k++) {
    if (String(peer[k].studentId) === String(viewerId)) {
      try { myPeerEval = JSON.parse(peer[k].ratings); } catch (e) {}
      break;
    }
  }

  return {
    meta: {
      id: String(meta.id),
      studentName: String(meta.studentName),
      title: String(meta.title),
      plan: String(meta.plan),
      selfEval: safeParse_(meta.selfEval)
    },
    fps: fps,
    frames: frames,
    hearts: hearts.length,
    hearted: hearted,
    comments: comments,
    peerSummary: peerSummary,
    myPeerEval: myPeerEval
  };
}

/* =========================================================
 *  하트 / 댓글 / 동료평가
 * ========================================================= */
function apiToggleHeart(animId, studentId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_('hearts');
    var values = sh.getDataRange().getValues();
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][0]) === String(animId) && String(values[r][1]) === String(studentId)) {
        sh.deleteRow(r + 1); // 이미 눌렀으면 취소
        return { hearted: false, hearts: countHeartsFor_(animId) };
      }
    }
    sh.appendRow([String(animId), String(studentId), new Date().getTime()]);
    return { hearted: true, hearts: countHeartsFor_(animId) };
  } finally {
    lock.releaseLock();
  }
}

function apiAddComment(animId, studentId, studentName, text) {
  text = String(text || '').trim();
  if (!text) throw new Error('댓글 내용을 적어주세요.');
  if (text.length > 300) text = text.substring(0, 300);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var id = Utilities.getUuid();
    var now = new Date().getTime();
    sheet_('comments').appendRow([id, String(animId), String(studentId), String(studentName), text, now]);
    return { id: id, studentName: String(studentName), text: text, createdAt: now };
  } finally {
    lock.releaseLock();
  }
}

function apiAddPeerEval(animId, studentId, studentName, ratings) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_('peerEvals');
    var values = sh.getDataRange().getValues();
    var ratingsStr = JSON.stringify(ratings || {});
    var now = new Date().getTime();
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][0]) === String(animId) && String(values[r][1]) === String(studentId)) {
        sh.getRange(r + 1, 4).setValue(ratingsStr); // 이미 평가했으면 갱신
        sh.getRange(r + 1, 5).setValue(now);
        return { updated: true };
      }
    }
    sh.appendRow([String(animId), String(studentId), String(studentName), ratingsStr, now]);
    return { updated: false };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 내부 도우미 ---------- */
function countHearts_() {
  var hearts = readRows_('hearts');
  var counts = {};
  for (var i = 0; i < hearts.length; i++) {
    var k = hearts[i].animId;
    counts[k] = (counts[k] || 0) + 1;
  }
  return counts;
}
function countHeartsFor_(animId) {
  return readRows_('hearts').filter(function (h) { return String(h.animId) === String(animId); }).length;
}
function summarizePeer_(peerRows) {
  var keys = ['story', 'motion', 'creative'];
  var sums = { story: 0, motion: 0, creative: 0 }, count = 0;
  for (var i = 0; i < peerRows.length; i++) {
    var r;
    try { r = JSON.parse(peerRows[i].ratings); } catch (e) { continue; }
    count++;
    keys.forEach(function (k) { sums[k] += Number(r[k]) || 0; });
  }
  var avg = {};
  keys.forEach(function (k) { avg[k] = count ? Math.round((sums[k] / count) * 10) / 10 : 0; });
  return { count: count, avg: avg };
}
function safeParse_(s) { try { return JSON.parse(s); } catch (e) { return {}; } }
