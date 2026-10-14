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
var VOTE_OPEN  = new Date('2026-10-14T17:00:00+09:00');
var VOTE_CLOSE = new Date('2026-10-22T23:59:59+09:00');
var VOTE_PURGE = new Date('2026-11-27T09:00:00+09:00');   // 결과 발표(10/27) 후 1개월 — 투표자 이메일·IP 삭제
var VOTE_PAUSED = false;                                  // 비상시 true 로 바꾸고 새 버전 배포 → 투표 일시 중지

var VOTE_SHEET = '투표';
var TALLY_SHEET = '집계';
var VOTE_HEADER = ['시각', '구분', '부문', '접수번호', '작품제목', '활동명', '이메일/휴대폰', '중복확인키', 'IP', '기기'];

var DISPOSABLE = ['mailinator.com', 'guerrillamail.com', 'guerrillamail.net', 'sharklasers.com', '10minutemail.com',
  '10minutemail.net', 'temp-mail.org', 'tempmail.com', 'tempmail.net', 'yopmail.com', 'yopmail.net', 'getnada.com',
  'dispostable.com', 'trashmail.com', 'maildrop.cc', 'throwawaymail.com', 'fakeinbox.com', 'mohmal.com',
  'emailondeck.com', 'mintemail.com', 'spamgourmet.com', 'tempinbox.com', 'mail.tm', 'tmpmail.org', 'tmails.net'];

function voteGet_(e) {
  var p = e.parameter || {};
  if (p.action === 'list') return json_(voteList_());
  if (p.action === 'img') return json_(voteImg_(p.id, p.w, p.n));
  if (p.action === 'diag') {   // 공개 작품 점검 (개인정보 없음)
    CacheService.getScriptCache().remove('entries');
    var d = scanEntries_();
    return json_({ ok: true, sheet: ss_().getName(), count: d.list.length, skipped: d.skipped || [], note: d.note || '', headers: d.headers || [],
      entries: d.list.map(function (e) { return { id: e.id, title: e.title, proc: e.proc.length }; }) });
  }
  if (p.action === 'pimg') return json_(judgeImg_(p));   // 심사위원 전용 과정 자료 (Judge.gs)
  return null;
}

function votePost_(p) {
  if (p.action === 'code') return json_(voteCode_(p));
  if (p.action === 'vote') return json_(voteCast_(p));
  if (/^j/.test(p.action || '')) return judgePost_(p);   // 심사위원 (Judge.gs)
  return null;
}

/** 투표·심사에 쓰는 스프레드시트 (접수 시트) — 2026-10-14 팀장님 확인 */
var VOTE_SS_ID = '1_Iph2kBcSI_LFgeh2bFwk5oXUPaRMSAFSOEPQN36nng';
var VOTE_TAB_GID = 1556491049;   // 접수 탭

function ss_() {
  return SpreadsheetApp.openById(VOTE_SS_ID);
}

/** 접수 탭: 이름이 「접수」인 탭 → 주소의 gid 탭 → 「작품이미지 링크」 제목이 있는 첫 탭 */
function entrySheet_() {
  var ss = ss_(), sh = ss.getSheetByName(SHEET_NAME);
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) if (all[i].getSheetId() === VOTE_TAB_GID) return all[i];
  for (var j = 0; j < all.length; j++) {
    var w = all[j].getLastColumn();
    if (w && all[j].getRange(1, 1, 1, w).getValues()[0].map(String).indexOf('작품이미지 링크') >= 0) return all[j];
  }
  return null;
}

/* ── 공개 작품 ───────────────────────────────────────── */

function publicEntries_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('entries');
  if (hit) return JSON.parse(hit);

  var list = scanEntries_().list;
  cache.put('entries', JSON.stringify(list), 300);
  return list;
}

/** 「공개」 칸 값: 비어 있지 않고 X·N·미공개 같은 부정이 아니면 공개 (O·o·ㅇ·0·○·V 등 모두 허용) */
function isPub_(v) {
  var s = String(v == null ? '' : v).trim();
  return !!s && !/^(x|×|n|no|false|미공개|비공개|제외|-)$/i.test(s);
}

/** 셀 글자 또는 셀에 걸린 링크에서 드라이브 파일/폴더 ID 뽑기 */
function driveIds_(text, rich, kind) {
  var urls = [String(text || '')];
  if (rich) rich.getRuns().forEach(function (run) { if (run.getLinkUrl()) urls.push(run.getLinkUrl()); });
  var re = kind === 'folder' ? /\/folders\/([\w-]{20,})/g : /(?:\/d\/|[?&]id=)([\w-]{20,})/g;
  var out = [];
  urls.join('\n').replace(re, function (_, id) { if (out.indexOf(id) < 0) out.push(id); return _; });
  return out;
}

function isImage_(f) { return /^image\//.test(f.getMimeType()); }

/** 접수 폴더 구조: making 폴더 안 이미지 = 과정 자료. making 폴더가 없으면 작품을 뺀 나머지 이미지를 과정 자료로 본다.
 *  작품 = 시트의 작품이미지 링크 파일 → 이름이 「작품」으로 시작하는 파일 → 가장 큰 파일 순 */
function folderImages_(folderId, preferId) {
  var folder = DriveApp.getFolderById(folderId), root = [], making = [], hasMaking = false;
  var it = folder.getFiles();
  while (it.hasNext()) { var f = it.next(); if (isImage_(f)) root.push(f); }
  var mk = folder.getFolders();
  while (mk.hasNext()) {
    var sub = mk.next();
    if (!/^making$/i.test(sub.getName().trim())) continue;
    hasMaking = true;
    var it2 = sub.getFiles();
    while (it2.hasNext()) { var g = it2.next(); if (isImage_(g)) making.push(g); }
  }
  var byName = function (a, b) { return a.getName() < b.getName() ? -1 : 1; };
  var work = preferId || null;
  if (!work) {
    var named = root.filter(function (x) { return /^작품/.test(x.getName()); });
    var pool = (named.length ? named : root).slice().sort(function (a, b) { return b.getSize() - a.getSize(); });
    work = pool[0] ? pool[0].getId() : null;
  }
  var proc = (hasMaking ? making : root.filter(function (x) { return x.getId() !== work; })).sort(byName);
  return { work: work, proc: proc.map(function (x) { return x.getId(); }) };
}

function scanEntries_() {
  var sh = entrySheet_();
  if (!sh) return { list: [], note: '접수 탭을 찾지 못함' };
  var range = sh.getDataRange(), v = range.getValues(), rich = range.getRichTextValues();
  var h = v[0].map(function (x) { return String(x).trim(); });
  var c = function (name) { return h.indexOf(name); };
  var cPub = c('공개'), cId = c('접수번호'), cTrack = c('응모부문'), cTitle = c('작품제목'),
      cArtist = c('활동명'), cSw = c('사용소프트웨어'), cDesc = c('작업설명'), cImg = c('작품이미지 링크'),
      cProc = c('과정자료 링크'), cFolder = c('폴더 링크'), cExt = c('외부소스'), cExtD = c('외부소스상세'),
      cAgree = c('과정자료활용동의');
  if (cPub < 0) return { list: [], note: '「공개」 제목 칸 없음', headers: h };

  var list = [], skipped = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    if (!isPub_(r[cPub])) continue;
    // 작품: 시트의 작품이미지 링크가 우선. 과정 자료: 접수 폴더의 making 폴더 (없으면 폴더 안 나머지 이미지)
    var file = cImg >= 0 ? (driveIds_(r[cImg], rich[i][cImg], 'file')[0] || null) : null, proc = [];
    var fids = cFolder >= 0 ? driveIds_(r[cFolder], rich[i][cFolder], 'folder') : [];
    if (fids.length) {
      try { var fi = folderImages_(fids[0], file); file = fi.work; proc = fi.proc; }
      catch (err) { console.error('folder ' + r[cId], err); }
    }
    if (!proc.length && cProc >= 0) proc = driveIds_(r[cProc], rich[i][cProc], 'file');
    if (!file) { skipped.push((i + 1) + '행'); continue; }
    list.push({
      id: String(r[cId] || 'row' + (i + 1)),
      track: String(r[cTrack]).indexOf('ZBrush') >= 0 ? 'ZBrush' : 'Cinema 4D',
      title: String(r[cTitle] || '무제'),
      artist: String(r[cArtist] || '응모자'),
      sw: String(r[cSw] || ''),
      desc: String(r[cDesc] || ''),
      file: file,
      proc: proc.filter(function (id) { return id !== file; }),
      ext: [cExt >= 0 ? r[cExt] : '', cExtD >= 0 ? r[cExtD] : ''].filter(String).join(' · '),
      procPublic: cAgree < 0 || String(r[cAgree]).trim() === '동의'   // 과정 자료 활용 동의한 사람만 일반 공개
    });
  }
  return { list: list, skipped: skipped, headers: h };
}

/** 편집기에서 실행: 공개 작품이 왜 안 나오는지 로그로 보여 줌 */
function checkVoteList() {
  CacheService.getScriptCache().remove('entries');
  var s = scanEntries_();
  Logger.log('읽은 시트: ' + ss_().getName() + ' · 탭 ' + (entrySheet_() ? entrySheet_().getName() : '없음'));
  Logger.log('공개 작품 ' + s.list.length + '점' + (s.note ? ' · ' + s.note : ''));
  if (s.skipped && s.skipped.length) Logger.log('O 표시는 있지만 이미지를 못 찾은 행: ' + s.skipped.join(', '));
  s.list.forEach(function (e) { Logger.log(e.id + ' · ' + e.title + ' · 과정 자료 ' + e.proc.length + '장'); });
  Logger.log('제목 줄: ' + (s.headers || []).join(' | '));
}

function voteList_() {
  var now = new Date();
  return {
    ok: true,
    open: VOTE_OPEN.toISOString(), close: VOTE_CLOSE.toISOString(),
    paused: VOTE_PAUSED, now: now.toISOString(), auth: authMode_(null), ipLimit: ipLimit_(),
    entries: publicEntries_().map(function (x) {
      return { id: x.id, track: x.track, title: x.title, artist: x.artist, sw: x.sw, desc: x.desc,
        proc: x.procPublic ? x.proc.length : 0 };
    })
  };
}

function voteImg_(id, w, n) {
  var x = publicEntries_().filter(function (e) { return e.id === id; })[0];
  if (!x) return { ok: false, error: 'not found' };
  if (n != null && n !== '') {   // making(제작 과정) 이미지
    var k = parseInt(n, 10);
    if (!x.procPublic || !(k >= 0 && k < x.proc.length)) return { ok: false, error: 'not found' };
    return driveImg_(x.proc[k], w);
  }
  return driveImg_(x.file, w);
}

function driveImg_(fileId, w) {
  var size = Math.min(Math.max(parseInt(w, 10) || 800, 200), 2000);
  var token = ScriptApp.getOAuthToken();
  var auth = { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
  var blob = null;
  try {
    var meta = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '?fields=thumbnailLink&supportsAllDrives=true', auth);
    var link = JSON.parse(meta.getContentText()).thumbnailLink;
    if (link) {
      var res = UrlFetchApp.fetch(link.replace(/=s\d+$/, '') + '=s' + size, auth);
      if (res.getResponseCode() === 200) blob = res.getBlob();
    }
  } catch (err) { console.error('thumb', err); }
  if (!blob) blob = DriveApp.getFileById(fileId).getBlob();   // 축소본을 못 받으면 원본
  if (blob.getBytes().length > 3200000) return { ok: false, error: 'image too large' };
  return { ok: true, type: blob.getContentType() || 'image/jpeg', data: Utilities.base64Encode(blob.getBytes()) };
}

/* ── 인증 (이메일 / 휴대폰) ─────────────────────────────── */
/*
 * 인증 방식 전환: 편집기에서 usePhoneAuth() 또는 useEmailAuth() 를 ▶ 실행 (재배포 필요 없음)
 * 휴대폰 문자(솔라피) 키: 프로젝트 설정(톱니바퀴) → 스크립트 속성에
 *   SOLAPI_KEY, SOLAPI_SECRET, SOLAPI_FROM(등록된 발신번호, 숫자만) 을 넣는다.
 * IP 제한: 스크립트 속성 IP_LIMIT (기본 1 = 같은 인터넷 주소에서 1표, 0 = 제한 없음)
 */
function prop_(k, d) { var v = PropertiesService.getScriptProperties().getProperty(k); return v == null || v === '' ? d : v; }

function authMode_(p) {
  var m = prop_('AUTH_MODE', 'email');
  if (p && p.preview && (p.auth === 'phone' || p.auth === 'email')) m = p.auth;   // 시험 모드에서만 방식 지정 가능
  return m === 'phone' ? 'phone' : 'email';
}
function ipLimit_() { return Math.max(0, parseInt(prop_('IP_LIMIT', '1'), 10) || 0); }

function usePhoneAuth() {
  var P = PropertiesService.getScriptProperties();
  if (!P.getProperty('SOLAPI_KEY') || !P.getProperty('SOLAPI_SECRET') || !P.getProperty('SOLAPI_FROM'))
    throw new Error('스크립트 속성에 SOLAPI_KEY · SOLAPI_SECRET · SOLAPI_FROM 을 먼저 넣어 주십시오.');
  P.setProperty('AUTH_MODE', 'phone');
  Logger.log('휴대폰 인증으로 전환했습니다. (1분 안에 투표 화면에 반영)');
}
function useEmailAuth() {
  PropertiesService.getScriptProperties().setProperty('AUTH_MODE', 'email');
  Logger.log('이메일 인증으로 전환했습니다.');
}
/** 문자 발송 시험: 아래 번호를 팀장님 번호로 바꿔 실행 */
function testSms() {
  var to = '01000000000';
  sendSms_(to, '[MAXON 캘린더 공모전] 문자 발송 시험입니다.');
  Logger.log('보냄: ' + to);
}

function normEmail_(s) {
  var e = String(s || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e)) return null;
  var at = e.lastIndexOf('@'), local = e.slice(0, at), dom = e.slice(at + 1);
  local = local.split('+')[0];
  if (dom === 'gmail.com' || dom === 'googlemail.com') { local = local.replace(/\./g, ''); dom = 'gmail.com'; }
  return { addr: e, key: local + '@' + dom, domain: dom, kind: 'email' };
}
function normPhone_(s) {
  var d = String(s || '').replace(/\D/g, '');
  if (/^82/.test(d)) d = '0' + d.slice(2);
  if (!/^01[016789]\d{7,8}$/.test(d)) return null;
  return { addr: d, key: 'tel:' + d, kind: 'phone' };
}
/** 요청에서 투표자 확인값 꺼내기 */
function ident_(p, mode) {
  if (mode === 'phone') {
    var t = normPhone_(p.phone || p.email);
    return t ? t : { error: '휴대폰 번호를 확인해 주십시오. (예: 010-1234-5678)' };
  }
  var e = normEmail_(p.email);
  if (!e) return { error: '이메일 주소를 확인해 주십시오.' };
  if (DISPOSABLE.indexOf(e.domain) >= 0) return { error: '일회용 메일 주소로는 투표할 수 없습니다.' };
  return e;
}
function who_(id) { return id.kind === 'phone' ? '이 휴대폰 번호는' : '이 이메일은'; }

function sendSms_(to, text) {
  var key = prop_('SOLAPI_KEY', ''), secret = prop_('SOLAPI_SECRET', ''), from = String(prop_('SOLAPI_FROM', '')).replace(/\D/g, '');
  if (!key || !secret || !from) throw new Error('문자 발송 설정이 없습니다.');
  var date = new Date().toISOString(), salt = Utilities.getUuid().replace(/-/g, '');
  var sig = Utilities.computeHmacSha256Signature(date + salt, secret).map(function (b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');
  var res = UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'HMAC-SHA256 apiKey=' + key + ', date=' + date + ', salt=' + salt + ', signature=' + sig },
    payload: JSON.stringify({ message: { to: to, from: from, text: text } })
  });
  if (res.getResponseCode() !== 200) {
    console.error('solapi', res.getResponseCode(), res.getContentText());
    throw new Error('문자 발송 실패 (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 200));
  }
}

function phase_(preview) {
  var now = new Date();
  if (VOTE_PAUSED) return { ok: false, error: '투표가 잠시 중지되었습니다. 잠시 후 다시 시도해 주십시오.' };
  if (now > VOTE_CLOSE) return { ok: false, error: '투표가 마감되었습니다. 결과는 10월 27일 발표합니다.' };
  if (now < VOTE_OPEN) return preview ? { ok: true, mode: '시험' } : { ok: false, error: '투표는 10월 14일(수) 오후 5시에 시작합니다.' };
  return { ok: true, mode: '정식' };
}

/** 같은 키로 window 초 안에 max 회 넘게 오면 true */
function limited_(key, max, windowSec) {
  var cache = CacheService.getScriptCache();
  var n = parseInt(cache.get(key) || '0', 10) + 1;
  cache.put(key, String(n), windowSec);
  return n > max;
}

var IP_MSG = '같은 인터넷(IP)에서 이미 투표가 이루어졌습니다. 학교·회사처럼 여러 사람이 같은 인터넷을 쓰는 곳이라면 휴대폰 데이터(LTE·5G)로 접속해 다시 시도해 주십시오.';
var NO_IP_MSG = '접속 정보를 확인할 수 없습니다. 투표 페이지에서 다시 시도해 주십시오.';

function voteCode_(p) {
  var ph = phase_(p.preview);
  if (!ph.ok) return ph;
  var mode = authMode_(p), id = ident_(p, mode);
  if (id.error) return { ok: false, error: id.error };

  var ip = String(p.ip || '').slice(0, 64);
  if (!ip) return { ok: false, error: NO_IP_MSG };
  if (limited_('ipc:' + ip, 15, 600)) return { ok: false, error: '요청이 너무 많습니다. 10분 뒤에 다시 시도해 주십시오.' };
  if (limited_('idc:' + id.key, 3, 600)) return { ok: false, error: '인증번호는 10분에 3번까지 받을 수 있습니다.' + (mode === 'email' ? ' 받은 메일함·스팸함을 확인해 주십시오.' : '') };

  var voted = votedTracks_(id.key, ph.mode);
  if (voted.length >= 1) return { ok: false, error: who_(id) + ' 이미 투표했습니다. (1인 1표)', already: true };
  if (ipLimit_() && ipVotes_(ip, ph.mode) >= ipLimit_()) return { ok: false, error: IP_MSG, ipBlocked: true };

  var code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('code:' + id.key, JSON.stringify({ code: code, tries: 0 }), 1800);

  try {
    if (mode === 'phone') {
      sendSms_(id.addr, '[MAXON 캘린더 공모전] 인기상 투표 인증번호 [' + code + '] 30분 안에 입력해 주세요.');
    } else {
      var subject = '[MAXON 2027 캘린더 공모전] 인기상 투표 인증번호 ' + code;
      var text = '인증번호: ' + code + '\n\n투표 화면에 위 6자리 번호를 입력해 주십시오. 30분 동안 유효합니다.'
        + '\n전체 작품 중 1작품에만 투표할 수 있습니다. (1인 1표)'
        + '\n직접 요청하지 않았다면 이 메일은 무시하셔도 됩니다.\n\n마루인터내셔널 · MAXON 2027 캘린더 공모전';
      var html = '<div style="font-family:sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:520px">'
        + '<p>MAXON 2027 캘린더 공모전 인기상 투표 인증번호입니다.</p>'
        + '<p style="font-size:30px;font-weight:700;letter-spacing:6px;margin:16px 0">' + code + '</p>'
        + '<p>투표 화면에 위 6자리 번호를 입력해 주십시오. 30분 동안 유효합니다.<br>'
        + '전체 작품 중 1작품에만 투표할 수 있습니다. (1인 1표)</p>'
        + '<p style="color:#666;font-size:13px">직접 요청하지 않았다면 이 메일은 무시하셔도 됩니다.<br>마루인터내셔널 · MAXON 2027 캘린더 공모전</p></div>';
      MailApp.sendEmail(id.addr, subject, text, { name: MAIL_SENDER_NAME, htmlBody: html });
    }
  } catch (err) {
    console.error('code send failed', err);
    return { ok: false, error: (mode === 'phone' ? '문자를' : '메일을') + ' 보내지 못했습니다. 잠시 후 다시 시도해 주십시오.' };
  }
  return { ok: true, voted: voted, mode: ph.mode, auth: mode };
}

/* ── 투표 ───────────────────────────────────────────── */

function voteSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(VOTE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VOTE_SHEET);
    sh.getRange(1, 1, 1, VOTE_HEADER.length).setValues([VOTE_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 이 투표자가 이미 투표한 부문 목록 */
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

/** 이 IP 에서 이미 들어온 표 수 */
function ipVotes_(ip, mode) {
  var sh = voteSheet_();
  var found = sh.getRange('I:I').createTextFinder(ip).matchEntireCell(true).findAll();
  var n = 0;
  found.forEach(function (r) { if (String(sh.getRange(r.getRow(), 2).getValue()) === mode) n++; });
  return n;
}

function voteCast_(p) {
  var ph = phase_(p.preview);
  if (!ph.ok) return ph;
  var mode = authMode_(p), id = ident_(p, mode);
  if (id.error) return { ok: false, error: id.error };
  var ip = String(p.ip || '').slice(0, 64);
  if (!ip) return { ok: false, error: NO_IP_MSG };

  var cache = CacheService.getScriptCache();
  var raw = cache.get('code:' + id.key);
  if (!raw) return { ok: false, error: '인증번호가 만료되었습니다. 다시 받아 주십시오.', expired: true };
  var c = JSON.parse(raw);
  if (c.tries >= 5) return { ok: false, error: '인증번호를 5번 틀렸습니다. 다시 받아 주십시오.', expired: true };
  if (String(p.code || '').trim() !== c.code) {
    c.tries++;
    cache.put('code:' + id.key, JSON.stringify(c), 1800);
    return { ok: false, error: '인증번호가 맞지 않습니다. (' + c.tries + '/5)' };
  }

  var x = publicEntries_().filter(function (e) { return e.id === String(p.id); })[0];
  if (!x) return { ok: false, error: '작품을 찾을 수 없습니다. 화면을 새로고침해 주십시오.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (votedTracks_(id.key, ph.mode).length)   // 1인 1표: 부문과 관계없이 한 번만
      return { ok: false, error: who_(id) + ' 이미 투표했습니다. (1인 1표)', already: true };
    if (ipLimit_() && ipVotes_(ip, ph.mode) >= ipLimit_()) return { ok: false, error: IP_MSG, ipBlocked: true };
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    var row = [now, ph.mode, x.track, x.id, x.title, x.artist, id.addr, id.key + '|' + x.track + '|' + ph.mode,
      ip, String(p.device || '').slice(0, 40)];
    voteSheet_().appendRow(row.map(function (v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; }));
    cache.remove('code:' + id.key);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, track: x.track, voted: [x.track], mode: ph.mode };
}

/* ── 운영자용 (편집기에서 실행) ─────────────────────────── */

/** 처음 한 번 실행: 「투표」·「집계」 탭, 개인정보 삭제 예약 */
function setupVote() {
  var ss = ss_();
  voteSheet_();
  setupJudge_();   // 심사위원 계정·심사 탭 (Judge.gs)
  var t = ss.getSheetByName(TALLY_SHEET) || ss.insertSheet(TALLY_SHEET);
  t.clear();
  t.getRange('A1').setValue('정식 투표만 집계 (시험 투표 제외). 1인 1표 · 부문별 최다 득표작이 인기상. 마감까지 외부 비공개.');
  t.getRange('A3').setFormula("=IFERROR(QUERY('" + VOTE_SHEET + "'!A:F,\"select D, E, F, C, count(D) where B = '정식' group by D, E, F, C order by count(D) desc label count(D) '득표'\",1),\"아직 투표 없음\")");
  t.getRange('H3').setFormula("=IFERROR(QUERY('" + VOTE_SHEET + "'!A:C,\"select count(B) where B = '정식' label count(B) '총 투표수'\",1),\"\")");

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
  var sh = ss_().getSheetByName(VOTE_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  sh.getRange(2, 7, sh.getLastRow() - 1, 4).clearContent();
}
