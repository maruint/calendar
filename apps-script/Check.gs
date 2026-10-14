/**
 * 중복 접수 점검 (편집기에서 실행, 로그로만 결과를 보여 줌 — 시트·파일은 고치지 않음)
 * checkDuplicates() : 드라이브에서 이름이 들어간 접수 폴더를 모두 찾아,
 *                     폴더별 작품 이미지가 같은 파일인지(해시·크기·해상도) 비교하고, 시트에 접수 기록이 있는지 표시
 */
var DUP_NAME = '김서진';

function checkDuplicates() {
  // 시트에 남은 접수번호
  var sh = typeof entrySheet_ === 'function' ? entrySheet_()
    : SpreadsheetApp.openById('1_Iph2kBcSI_LFgeh2bFwk5oXUPaRMSAFSOEPQN36nng').getSheets().filter(function (t) { return t.getSheetId() === 1556491049; })[0];
  var v = sh.getDataRange().getValues();
  var h = v[0].map(function (x) { return String(x).trim(); });
  var cId = h.indexOf('접수번호'), cName = h.indexOf('이름'), cTitle = h.indexOf('작품제목'), cPub = h.indexOf('공개');
  var inSheet = {};
  for (var i = 1; i < v.length; i++) inSheet[String(v[i][cId])] = { row: i + 1, name: v[i][cName], title: v[i][cTitle], pub: cPub >= 0 ? v[i][cPub] : '' };

  var auth = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  var meta = function (id) {
    var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id +
      '?supportsAllDrives=true&fields=name,size,md5Checksum,imageMediaMetadata(width,height)', auth);
    return r.getResponseCode() === 200 ? JSON.parse(r.getContentText()) : null;
  };

  // 드라이브에서 이름이 들어간 폴더 찾기 (접수 폴더 이름 = 접수번호_이름)
  var folders = [], it = DriveApp.searchFolders("title contains '" + DUP_NAME + "' and trashed = false");
  while (it.hasNext()) folders.push(it.next());
  folders.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
  Logger.log('━━ "' + DUP_NAME + '" 이 들어간 폴더 ' + folders.length + '개 ━━');

  var byHash = {};
  folders.forEach(function (f) {
    var sid = f.getName().split('_').slice(0, 2).join('_');
    var rec = inSheet[sid];
    var imgs = [], fit = f.getFiles();
    while (fit.hasNext()) { var x = fit.next(); if (/^image\//.test(x.getMimeType())) imgs.push(x); }
    var mk = 0, sub = f.getFolders();
    while (sub.hasNext()) { var s = sub.next(); var sf = s.getFiles(); while (sf.hasNext()) { if (/^image\//.test(sf.next().getMimeType())) mk++; } }
    imgs.sort(function (a, b) { return b.getSize() - a.getSize(); });
    var parts = imgs.map(function (x) {
      var m = meta(x.getId()) || {};
      var key = m.md5Checksum || (x.getSize() + '');
      (byHash[key] = byHash[key] || []).push(f.getName());
      return x.getName() + ' (' + Math.round(x.getSize() / 1024) + 'KB, ' +
        (m.imageMediaMetadata ? m.imageMediaMetadata.width + '×' + m.imageMediaMetadata.height : '?') + ', #' + key.slice(0, 6) + ')';
    });
    Logger.log(f.getName() + ' · 시트: ' + (rec ? rec.row + '행 「' + rec.title + '」 공개:' + (rec.pub || '-') : '기록 없음(제출 미완료)') +
      ' · 이미지 ' + imgs.length + '장: ' + (parts.join(' / ') || '없음') + ' · 하위폴더 이미지 ' + mk + '장');
  });

  var keys = Object.keys(byHash);
  Logger.log('▶ 서로 다른 이미지 파일: ' + keys.length + '가지 (같은 # 은 완전히 같은 파일)');
  keys.forEach(function (k, n) { Logger.log('   #' + k.slice(0, 6) + ' → ' + byHash[k].length + '개 폴더: ' + byHash[k].join(', ')); });
}
