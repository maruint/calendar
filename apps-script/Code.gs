/**
 * 2027 캘린더 공모전 접수 수신 (Google Apps Script 웹앱)
 *
 * entry.html 이 보내는 요청 (Content-Type: text/plain, 본문은 JSON)
 *   1) { action:'upload', folder, filename, mimeType, data(base64) }  → 파일 1개를 드라이브에 저장
 *   2) { action:'submit', id, folder, data:{...}, image, process:[] } → 시트에 한 행 기록
 *
 * 설정: 아래 두 ID를 채우십시오.
 *   SHEET_ID  : 구글 시트 URL  https://docs.google.com/spreadsheets/d/【여기】/edit
 *   FOLDER_ID : 드라이브 폴더 URL https://drive.google.com/drive/folders/【여기】
 */
var SHEET_ID   = '158Ek8HL_aqtSQgwt2SCJSoNGhtZcYpygVgy6O-CX-FE';
var SHEET_NAME = '접수';
var FOLDER_ID  = '1sfKi7BI9uKvzoLh7wWHilqZCTVf5anCR';

// 접수 확인 메일 — 접수자가 답장하면 REPLY_TO로 가도록 설정. CC는 비워두면 발송 안 함 (예: 'hsbae@maru.co.kr')
var MAIL_SENDER_NAME = '2027 캘린더 공모전 · 마루인터내셔널';
var MAIL_REPLY_TO    = 'hsbae@maru.co.kr';
var MAIL_CC          = '';

var FIELDS = ['이름','활동명','이메일','휴대전화','소속구분','학교·회사','응모부문','작품제목',
  '사용소프트웨어','작업설명','이미지파일','제작과정자료','외부소스','외부소스상세','시상식참석',
  '전달사항','과정자료활용동의','소식수신동의','접수시각'];

function doGet() {
  return ContentService.createTextOutput('2027 캘린더 공모전 접수 수신 중');
}

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.action === 'upload') return json_(upload_(p));
    if (p.action === 'submit') return json_(submit_(p));
    throw new Error('unknown action');
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function upload_(p) {
  var folder = entryFolder_(p.folder);
  var blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mimeType || 'image/jpeg', p.filename);
  var file = folder.createFile(blob);
  return { ok: true, id: file.getId(), url: file.getUrl() };
}

function submit_(p) {
  var folderUrl = entryFolder_(p.folder).getUrl();
  var d = p.data || {};
  var header = ['접수번호'].concat(FIELDS, ['작품이미지 링크', '과정자료 링크', '폴더 링크', '확인메일']);
  var sh, rowNum;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    // 첫 행 제목 (기존 시트에 '확인메일' 열이 없으면 추가)
    if (sh.getLastRow() === 0 || sh.getLastColumn() < header.length) {
      sh.getRange(1, 1, 1, header.length).setValues([header]);
    }

    var row = [p.id].concat(FIELDS.map(function (k) { return d[k] == null ? '' : d[k]; }), [
      p.image || '',
      (p.process || []).join('\n'),
      folderUrl,
      '발송 대기'
    ]);
    // 전화번호 등이 숫자로 바뀌지 않도록 텍스트로 기록
    sh.appendRow(row.map(function (v) { return typeof v === 'string' && /^[=+\-0-9]/.test(v) ? "'" + v : v; }));
    rowNum = sh.getLastRow();
  } finally {
    lock.releaseLock();
  }

  // 메일 실패가 접수 실패가 되지 않도록 분리 — 결과는 시트 '확인메일' 열에 기록
  var status;
  try {
    sendConfirmation_(p.id, d);
    status = '발송 ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM-dd HH:mm');
  } catch (err) {
    console.error('confirmation mail failed', err);
    status = '실패: ' + String(err && err.message || err);
  }
  sh.getRange(rowNum, header.length).setValue(status);

  return { ok: true, id: p.id };
}

function sendConfirmation_(id, d) {
  var to = String(d['이메일'] || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) throw new Error('이메일 주소 없음');

  var rows = [
    ['접수번호', id],
    ['이름 (활동명)', d['이름'] + ' (' + d['활동명'] + ')'],
    ['응모부문', d['응모부문']],
    ['작품제목', d['작품제목']],
    ['작품 이미지', d['이미지파일']],
    ['제작 과정 자료', d['제작과정자료']],
    ['시상식 참석', d['시상식참석']],
    ['접수시각', d['접수시각']]
  ];
  var subject = '[2027 캘린더 공모전] 접수가 완료되었습니다 (접수번호 ' + id + ')';

  var text = d['이름'] + '님, 2027 캘린더 공모전에 응모해 주셔서 감사합니다.\n아래 내용으로 접수되었습니다.\n\n'
    + rows.map(function (r) { return r[0] + ' : ' + r[1]; }).join('\n')
    + '\n\n자격 검수는 9월 28일부터 29일까지 진행하며, 확인이 필요하면 등록하신 연락처로 연락드립니다.'
    + '\n1인 1작품이므로 추가 접수는 받지 않습니다. 제출 내용 수정·교체는 마감 전까지 아래로 연락 주십시오.'
    + '\n\n문의 · 담당자 hsbae@maru.co.kr · 02-6245-1628\n마루인터내셔널';

  var html = '<div style="font-family:sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:560px">'
    + '<p><b>' + esc_(d['이름']) + '</b>님, 2027 캘린더 공모전에 응모해 주셔서 감사합니다.<br>아래 내용으로 접수되었습니다.</p>'
    + '<table style="border-collapse:collapse;width:100%;margin:12px 0">'
    + rows.map(function (r) {
        return '<tr><th style="text-align:left;padding:6px 10px;border:1px solid #ddd;background:#f6f6f6;white-space:nowrap;font-weight:600">'
          + esc_(r[0]) + '</th><td style="padding:6px 10px;border:1px solid #ddd">' + esc_(r[1]) + '</td></tr>';
      }).join('')
    + '</table>'
    + '<p>자격 검수는 9월 28일부터 29일까지 진행하며, 확인이 필요하면 등록하신 연락처로 연락드립니다.<br>'
    + '1인 1작품이므로 추가 접수는 받지 않습니다. 제출 내용 수정·교체는 마감 전까지 아래로 연락 주십시오.</p>'
    + '<p style="color:#666;font-size:13px">문의 · 담당자 hsbae@maru.co.kr · 02-6245-1628<br>마루인터내셔널</p></div>';

  var opt = { name: MAIL_SENDER_NAME, htmlBody: html };
  if (MAIL_REPLY_TO) opt.replyTo = MAIL_REPLY_TO;
  if (MAIL_CC) opt.cc = MAIL_CC;
  MailApp.sendEmail(to, subject, text, opt);
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

/** 접수 건별 하위 폴더 (동시 업로드 시 중복 생성 방지) */
function entryFolder_(name) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var root = DriveApp.getFolderById(FOLDER_ID);
    var it = root.getFoldersByName(name);
    return it.hasNext() ? it.next() : root.createFolder(name);
  } finally {
    lock.releaseLock();
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 편집기에서 한 번 실행해 드라이브·시트 권한을 승인하십시오. */
function authorize() {
  DriveApp.getFolderById(FOLDER_ID).getName();
  SpreadsheetApp.openById(SHEET_ID).getName();
  Logger.log('오늘 남은 메일 발송 가능 수: ' + MailApp.getRemainingDailyQuota());
}
