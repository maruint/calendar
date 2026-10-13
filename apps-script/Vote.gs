/**
 * 2027 캘린더 공모전 인기상 온라인 투표 (Code.gs 와 같은 프로젝트에 넣는 파일)
 *
 * 요청은 Vercel /api/vote 를 거쳐 들어온다 (ip 를 붙여 줌). Code.gs 의 doGet/doPost 가 여기로 넘긴다.
 *   GET  ?action=list            → 공개 작품 목록 (개인정보 제외)
 *   GET  ?action=img&id=&w=      → 작품 이미지 (축소본, base64)
 *   POST {action:'code', email, preview, ip}               → 인증번호 메일 발송
 *   POST {action:'vote', email, code, id, device, preview, ip} → 투표 기록
 *
 * 운영자가 할 일
 *   1) 「접수」 시트에 「공개」 열을 만들고, 투표에 공개할 작품 행에 O 를 적는다.
 *   2) 편집기에서 setupVote() 를 한 번 실행한다 → 「투표」·「집계」 탭과 개인정보 삭제 예약이 만들어진다.
 *   3) 배포 → 배포 관리 → 편집(연필) → 버전: 새 버전 → 배포. (주소는 그대로)
 */
var VOTE_OPEN  = new Date('2026-10-14T00:00:00+09:00');
var VOTE_CLOSE = new Date('2026-10-22T23:59:59+09:00');
var VOTE_PURGE = new Date('2026-11-27T09:00:00+09:00');   // 결과 발표(10/27) 후 1개월 — 투표자 이메일·IP 삭제
var VOTE_PAUSED = false;                                  // 비상시 true 로 바꾸고 새 버전 배포 → 투표 일시 중지

var VOTE_SHEET = '투표';
var TALLY_SHEET = '집계';
var VOTE_HEADER = ['시각', '구분', '부문', '접수번호', '작품제목', '활동명', '이메일', '중복확인키', 'IP', '기기'];

var DISPOSABLE = ['mailinator.com', 'guerrillamail.com', 'guerrillamail.net', 'sharklasers.com', '10minutemail.com',
  '10minutemail.net', 'temp-mail.org', 'tempmail.com', 'tempmail.net', 'yopmail.com', 'yopmail.net', 'getnada.com',
  'dispostable.com', 'trashmail.com', 'maildrop.cc', 'throwawaymail.com', 'fakeinbox.com', 'mohmal.com',
  'emailondeck.com', 'mintemail.com', 'spamgourmet.com', 'tempinbox.com', 'mail.tm', 'tmpmail.org', 'tmails.net'];

function voteGet_(e) {
  var p = e.parameter || {};
  if (p.action === 'list') return json_(voteList_());
  if (p.action === 'img') return json_(voteImg_(p.id, p.w));
  return null;
}

function votePost_(p) {
  if (p.action === 'code') return json_(voteCode_(p));
  if (p.action === 'vote') return json_(voteCast_(p));
  return null;
}

/* ── 공개 작품 ───────────────────────────────────────── */

function publicEntries_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('entries');
  if (hit) return JSON.parse(hit);

  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var v = sh.getDataRange().getValues();
  var h = v[0].map(function (x) { return String(x).trim(); });
  var c = function (name) { return h.indexOf(name); };
  var cPub = c('공개'), cId = c('접수번호'), cTrack = c('응모부문'), cTitle = c('작품제목'),
      cArtist = c('활동명'), cSw = c('사용소프트웨어'), cDesc = c('작업설명'), cImg = c('작품이미지 링크');
  if (cPub < 0) return [];

  var list = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    if (!/^(o|○|◯|y|yes|공개|v|✓|true)$/i.test(String(r[cPub]).trim())) continue;
    var m = String(r[cImg]).match(/\/d\/([\w-]{20,})/) || String(r[cImg]).match(/[?&]id=([\w-]{20,})/);
    if (!m) continue;
    list.push({
      id: String(r[cId]),
      track: String(r[cTrack]).indexOf('ZBrush') >= 0 ? 'ZBrush' : 'Cinema 4D',
      title: String(r[cTitle] || '무제'),
      artist: String(r[cArtist] || '응모자'),
      sw: String(r[cSw] || ''),
      desc: String(r[cDesc] || ''),
      file: m[1]
    });
  }
  cache.put('entries', JSON.stringify(list), 300);
  return list;
}

function voteList_() {
  var now = new Date();
  return {
    ok: true,
    open: VOTE_OPEN.toISOString(), close: VOTE_CLOSE.toISOString(),
    paused: VOTE_PAUSED, now: now.toISOString(),
    entries: publicEntries_().map(function (x) {
      return { id: x.id, track: x.track, title: x.title, artist: x.artist, sw: x.sw, desc: x.desc };
    })
  };
}

function voteImg_(id, w) {
  var x = publicEntries_().filter(function (e) { return e.id === id; })[0];
  if (!x) return { ok: false, error: 'not found' };
  var size = Math.min(Math.max(parseInt(w, 10) || 800, 200), 2000);
  var token = ScriptApp.getOAuthToken();
  var auth = { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
  var blob = null;
  try {
    var meta = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + x.file + '?fields=thumbnailLink&supportsAllDrives=true', auth);
    var link = JSON.parse(meta.getContentText()).thumbnailLink;
    if (link) {
      var res = UrlFetchApp.fetch(link.replace(/=s\d+$/, '') + '=s' + size, auth);
      if (res.getResponseCode() === 200) blob = res.getBlob();
    }
  } catch (err) { console.error('thumb', err); }
  if (!blob) blob = DriveApp.getFileById(x.file).getBlob();   // 축소본을 못 받으면 원본
  if (blob.getBytes().length > 3200000) return { ok: false, error: 'image too large' };
  return { ok: true, type: blob.getContentType() || 'image/jpeg', data: Utilities.base64Encode(blob.getBytes()) };
}

/* ── 인증번호 ───────────────────────────────────────── */

function normEmail_(s) {
  var e = String(s || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e)) return null;
  var at = e.lastIndexOf('@'), local = e.slice(0, at), dom = e.slice(at + 1);
  local = local.split('+')[0];
  if (dom === 'gmail.com' || dom === 'googlemail.com') { local = local.replace(/\./g, ''); dom = 'gmail.com'; }
  return { addr: e, key: local + '@' + dom, domain: dom };
}

function phase_(preview) {
  var now = new Date();
  if (VOTE_PAUSED) return { ok: false, error: '투표가 잠시 중지되었습니다. 잠시 후 다시 시도해 주십시오.' };
  if (now > VOTE_CLOSE) return { ok: false, error: '투표가 마감되었습니다. 결과는 10월 27일 발표합니다.' };
  if (now < VOTE_OPEN) return preview ? { ok: true, mode: '시험' } : { ok: false, error: '투표는 10월 14일에 시작합니다.' };
  return { ok: true, mode: '정식' };
}

/** 같은 키로 window 초 안에 max 회 넘게 오면 true */
function limited_(key, max, windowSec) {
  var cache = CacheService.getScriptCache();
  var n = parseInt(cache.get(key) || '0', 10) + 1;
  cache.put(key, String(n), windowSec);
  return n > max;
}

function voteCode_(p) {
  var ph = phase_(p.preview);
  if (!ph.ok) return ph;
  var em = normEmail_(p.email);
  if (!em) return { ok: false, error: '이메일 주소를 확인해 주십시오.' };
  if (DISPOSABLE.indexOf(em.domain) >= 0) return { ok: false, error: '일회용 메일 주소로는 투표할 수 없습니다.' };

  var ip = String(p.ip || '').slice(0, 64);
  if (ip && limited_('ipc:' + ip, 15, 600)) return { ok: false, error: '요청이 너무 많습니다. 10분 뒤에 다시 시도해 주십시오.' };
  if (limited_('emc:' + em.key, 3, 600)) return { ok: false, error: '인증번호는 10분에 3번까지 받을 수 있습니다. 받은 메일함·스팸함을 확인해 주십시오.' };

  var voted = votedTracks_(em.key, ph.mode);
  if (voted.length >= 2) return { ok: false, error: '이 이메일은 두 부문 모두 이미 투표했습니다.' };

  var code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('code:' + em.key, JSON.stringify({ code: code, tries: 0 }), 1800);

  var subject = '[2027 캘린더 공모전] 인기상 투표 인증번호 ' + code;
  var text = '인증번호: ' + code + '\n\n투표 화면에 위 6자리 번호를 입력해 주십시오. 30분 동안 유효합니다.'
    + '\n이 번호로 Cinema 4D·ZBrush 부문에 각 1표씩 투표할 수 있습니다.'
    + '\n직접 요청하지 않았다면 이 메일은 무시하셔도 됩니다.\n\n마루인터내셔널 · 2027 캘린더 공모전';
  var html = '<div style="font-family:sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:520px">'
    + '<p>2027 캘린더 공모전 인기상 투표 인증번호입니다.</p>'
    + '<p style="font-size:30px;font-weight:700;letter-spacing:6px;margin:16px 0">' + code + '</p>'
    + '<p>투표 화면에 위 6자리 번호를 입력해 주십시오. 30분 동안 유효합니다.<br>'
    + '이 번호로 Cinema 4D·ZBrush 부문에 각 1표씩 투표할 수 있습니다.</p>'
    + '<p style="color:#666;font-size:13px">직접 요청하지 않았다면 이 메일은 무시하셔도 됩니다.<br>마루인터내셔널 · 2027 캘린더 공모전</p></div>';
  try {
    MailApp.sendEmail(em.addr, subject, text, { name: MAIL_SENDER_NAME, htmlBody: html });
  } catch (err) {
    console.error('code mail failed', err);
    return { ok: false, error: '메일을 보내지 못했습니다. 잠시 후 다시 시도해 주십시오.' };
  }
  return { ok: true, voted: voted, mode: ph.mode };
}

/* ── 투표 ───────────────────────────────────────────── */

function voteSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(VOTE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VOTE_SHEET);
    sh.getRange(1, 1, 1, VOTE_HEADER.length).setValues([VOTE_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 이 이메일이 이미 투표한 부문 목록 */
function votedTracks_(key, mode) {
  var sh = voteSheet_();
  var found = sh.getRange('H:H').createTextFinder(key + '|').matchEntireCell(false).findAll();
  var out = [];
  found.forEach(function (r) {
    var parts = String(r.getValue()).split('|');   // key|부문|구분
    if (parts[0] === key && parts[2] === mode && out.indexOf(parts[1]) < 0) out.push(parts[1]);
  });
  return out;
}

function voteCast_(p) {
  var ph = phase_(p.preview);
  if (!ph.ok) return ph;
  var em = normEmail_(p.email);
  if (!em) return { ok: false, error: '이메일 주소를 확인해 주십시오.' };

  var cache = CacheService.getScriptCache();
  var raw = cache.get('code:' + em.key);
  if (!raw) return { ok: false, error: '인증번호가 만료되었습니다. 다시 받아 주십시오.', expired: true };
  var c = JSON.parse(raw);
  if (c.tries >= 5) return { ok: false, error: '인증번호를 5번 틀렸습니다. 다시 받아 주십시오.', expired: true };
  if (String(p.code || '').trim() !== c.code) {
    c.tries++;
    cache.put('code:' + em.key, JSON.stringify(c), 1800);
    return { ok: false, error: '인증번호가 맞지 않습니다. (' + c.tries + '/5)' };
  }

  var x = publicEntries_().filter(function (e) { return e.id === String(p.id); })[0];
  if (!x) return { ok: false, error: '작품을 찾을 수 없습니다. 화면을 새로고침해 주십시오.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (votedTracks_(em.key, ph.mode).indexOf(x.track) >= 0)
      return { ok: false, error: '이 이메일은 ' + x.track + ' 부문에 이미 투표했습니다.', already: x.track };
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    var row = [now, ph.mode, x.track, x.id, x.title, x.artist, em.addr, em.key + '|' + x.track + '|' + ph.mode,
      String(p.ip || '').slice(0, 64), String(p.device || '').slice(0, 40)];
    voteSheet_().appendRow(row.map(function (v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; }));
  } finally {
    lock.releaseLock();
  }
  return { ok: true, track: x.track, voted: votedTracks_(em.key, ph.mode), mode: ph.mode };
}

/* ── 운영자용 (편집기에서 실행) ─────────────────────────── */

/** 처음 한 번 실행: 「투표」·「집계」 탭, 개인정보 삭제 예약 */
function setupVote() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  voteSheet_();
  var t = ss.getSheetByName(TALLY_SHEET) || ss.insertSheet(TALLY_SHEET);
  t.clear();
  t.getRange('A1').setValue('정식 투표만 집계 (시험 투표 제외). 마감까지 외부 비공개.');
  t.getRange('A3').setFormula("=IFERROR(QUERY('" + VOTE_SHEET + "'!A:F,\"select C, D, E, F, count(D) where B = '정식' group by C, D, E, F order by C, count(D) desc label count(D) '득표'\",1),\"아직 투표 없음\")");
  t.getRange('H3').setFormula("=IFERROR(QUERY('" + VOTE_SHEET + "'!A:C,\"select C, count(C) where B = '정식' group by C label C '부문', count(C) '총 투표수'\",1),\"\")");

  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'purgeVoterInfo') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('purgeVoterInfo').timeBased().at(VOTE_PURGE).create();

  // UrlFetch·메일 권한 확인
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  Logger.log('공개 작품 ' + publicEntries_().length + '점 · 오늘 남은 메일 ' + MailApp.getRemainingDailyQuota() + '통 · 개인정보 삭제 예약 ' + VOTE_PURGE);
}

/** 「공개」 표시를 바꾼 뒤 바로 반영하고 싶을 때 (안 해도 5분 안에 반영) */
function refreshVoteList() {
  CacheService.getScriptCache().remove('entries');
  Logger.log('공개 작품 ' + publicEntries_().length + '점');
}

/** 결과 발표 1개월 뒤 자동 실행: 투표자 이메일·중복확인키·IP·기기 칸을 지움 (득표 기록은 남음) */
function purgeVoterInfo() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(VOTE_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  sh.getRange(2, 7, sh.getLastRow() - 1, 4).clearContent();
}
