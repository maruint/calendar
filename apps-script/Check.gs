/**
 * 중복 접수 점검 (편집기에서 실행, 로그로만 결과를 보여 줌 — 시트·파일은 고치지 않음)
 * checkDuplicates() : 이름이 같은 접수 건들의 작품 이미지가 같은 파일인지 비교
 */
var DUP_NAME = '김서진';   // 비워 두면 이름이 2건 이상인 사람 모두

function checkDuplicates() {
  var sh = typeof entrySheet_ === 'function' ? entrySheet_()
    : SpreadsheetApp.openById('1_Iph2kBcSI_LFgeh2bFwk5oXUPaRMSAFSOEPQN36nng').getSheets().filter(function (t) { return t.getSheetId() === 1556491049; })[0];
  var v = sh.getDataRange().getValues(), rich = sh.getDataRange().getRichTextValues();
  var h = v[0].map(function (x) { return String(x).trim(); });
  var c = function (n) { return h.indexOf(n); };
  var cName = c('이름'), cId = c('접수번호'), cTime = c('접수시각'), cTrack = c('응모부문'), cTitle = c('작품제목'),
      cImg = c('작품이미지 링크'), cFolder = c('폴더 링크'), cPub = c('공개');

  var groups = {};
  for (var i = 1; i < v.length; i++) {
    var name = String(v[i][cName]).trim();
    if (!name || (DUP_NAME && name !== DUP_NAME)) continue;
    (groups[name] = groups[name] || []).push(i);
  }
  var auth = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  var meta = function (id) {
    var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id +
      '?supportsAllDrives=true&fields=name,size,md5Checksum,imageMediaMetadata(width,height),createdTime', auth);
    return r.getResponseCode() === 200 ? JSON.parse(r.getContentText()) : null;
  };

  Object.keys(groups).forEach(function (name) {
    var rows = groups[name];
    if (rows.length < 2 && !DUP_NAME) return;
    Logger.log('━━ ' + name + ' · 접수 ' + rows.length + '건 ━━');
    var byHash = {};
    rows.forEach(function (i) {
      var r = v[i], fileId = null, extra = 0;
      var fids = cFolder >= 0 ? driveIds_(r[cFolder], rich[i][cFolder], 'folder') : [];
      if (fids.length) {
        try {
          var fi = folderImages_(fids[0]); fileId = fi.work; extra = fi.proc.length;
        } catch (e) {}
      }
      if (!fileId && cImg >= 0) fileId = driveIds_(r[cImg], rich[i][cImg], 'file')[0];
      var m = fileId ? meta(fileId) : null;
      var key = m ? (m.md5Checksum || (m.size + '_' + (m.imageMediaMetadata ? m.imageMediaMetadata.width + 'x' + m.imageMediaMetadata.height : ''))) : '없음';
      (byHash[key] = byHash[key] || []).push(String(r[cId]));
      Logger.log((i + 1) + '행 · ' + r[cId] + ' · ' + r[cTime] + ' · ' + r[cTrack] + ' · 「' + r[cTitle] + '」' +
        ' · 공개:' + (cPub >= 0 ? (r[cPub] || '-') : '?') +
        ' · 파일: ' + (m ? m.name + ' (' + Math.round(m.size / 1024) + 'KB, ' +
          (m.imageMediaMetadata ? m.imageMediaMetadata.width + '×' + m.imageMediaMetadata.height : '?') + ')' : '못 찾음') +
        ' · making ' + extra + '장');
    });
    var keys = Object.keys(byHash);
    Logger.log('▶ 작품 이미지 종류: ' + keys.length + '가지');
    keys.forEach(function (k, n) { Logger.log('   ' + (n + 1) + '번 그림 = ' + byHash[k].join(', ')); });
  });
}
