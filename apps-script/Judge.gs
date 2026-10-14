/**
 * 2027 캘린더 공모전 심사위원 심사 (Code.gs·Vote.gs 와 같은 프로젝트에 넣는 파일)
 *
 * 투표 페이지(vote.html) 맨 아래 「심사위원 로그인」 → 아이디·비밀번호로 들어와 작품별 점수를 매긴다.
 *   POST {action:'jlogin', id, pw}                         → 토큰
 *   POST {action:'jdata',  token}                          → 작품(과정 자료 포함) + 내가 매긴 점수
 *   POST {action:'jscore', token, entry, s1, s2, s3, pick, note} → 점수 저장 (마감 전까지 고칠 수 있음)
 *   GET  ?action=pimg&id=접수번호&n=0&t=토큰&w=           → 과정 자료 이미지
 *
 * 계정: 「심사위원」 탭 (아이디 · 비밀번호 · 이름). setupVote() 가 처음 한 번 만들어 둔다.
 * 결과: 「심사」 탭 (심사위원·작품당 한 줄, 다시 저장하면 그 줄을 고침), 「심사집계」 탭 (작품별 평균).
 */
var JUDGE_CLOSE = new Date('2026-10-25T23:59:59+09:00');   // 10/26 합평 전날까지 저장 가능
var JUDGE_SHEET = '심사위원';
var SCORE_SHEET = '심사';
var SCORE_TALLY = '심사집계';
var SCORE_HEADER = ['저장시각', '심사위원', '접수번호', '부문', '작품제목', '활동명', '완성도', '독창성', '캘린더 적합성', '합계', '수록 추천', '의견', '키'];
var CRITERIA = ['완성도', '독창성', '캘린더 적합성'];   // 각 10점

function judgePost_(p) {
  if (p.action === 'jlogin') return json_(judgeLogin_(p));
  var j = judgeAuth_(p.token);
  if (!j) return json_({ ok: false, error: '로그인이 만료되었습니다. 다시 로그인해 주십시오.', relogin: true });
  if (p.action === 'jdata') return json_(judgeData_(j));
  if (p.action === 'jscore') return json_(judgeScore_(j, p));
  return json_({ ok: false, error: 'unknown action' });
}

function judgeSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('JUDGE_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('JUDGE_SECRET', s); }
  return s;
}

function sign_(text) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text, judgeSecret_())).replace(/=+$/, '');
}

function judges_() {
  var sh = ss_().getSheetByName(JUDGE_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return String(r[0]).trim() && String(r[1]).trim(); })
    .map(function (r) { return { id: String(r[0]).trim(), pw: String(r[1]).trim(), name: String(r[2] || r[0]).trim() }; });
}

function judgeLogin_(p) {
  var id = String(p.id || '').trim();
  if (limited_('jl:' + id.toLowerCase(), 10, 600) || (p.ip && limited_('jlip:' + p.ip, 30, 600)))
    return { ok: false, error: '로그인 시도가 너무 많습니다. 10분 뒤에 다시 시도해 주십시오.' };
  var j = judges_().filter(function (x) { return x.id.toLowerCase() === id.toLowerCase() && x.pw === String(p.pw || ''); })[0];
  if (!j) return { ok: false, error: '아이디 또는 비밀번호가 맞지 않습니다.' };
  var exp = Date.now() + 12 * 3600 * 1000;   // 12시간
  var body = Utilities.base64EncodeWebSafe(j.id + '|' + exp).replace(/=+$/, '');
  return { ok: true, token: body + '.' + sign_(body), name: j.name };
}

function judgeAuth_(token) {
  var t = String(token || '').split('.');
  if (t.length !== 2 || sign_(t[0]) !== t[1]) return null;
  var pad = t[0] + '==='.slice((t[0].length + 3) % 4);
  var parts = Utilities.newBlob(Utilities.base64DecodeWebSafe(pad)).getDataAsString().split('|');
  if (Number(parts[1]) < Date.now()) return null;
  return judges_().filter(function (x) { return x.id === parts[0]; })[0] || null;   // 계정을 지우면 바로 막힘
}

function scoreSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SCORE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SCORE_SHEET);
    sh.getRange(1, 1, 1, SCORE_HEADER.length).setValues([SCORE_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function myScores_(j) {
  var sh = scoreSheet_(), out = {};
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, SCORE_HEADER.length).getValues().forEach(function (r) {
    if (String(r[12]).indexOf(j.id + '|') !== 0) return;
    out[String(r[2])] = { s1: r[6], s2: r[7], s3: r[8], pick: r[10] === 'O', note: String(r[11] || ''), at: String(r[0]) };
  });
  return out;
}

function judgeData_(j) {
  return {
    ok: true, name: j.name, criteria: CRITERIA, close: JUDGE_CLOSE.toISOString(),
    entries: publicEntries_().map(function (x) {
      return { id: x.id, track: x.track, title: x.title, artist: x.artist, sw: x.sw, desc: x.desc, ext: x.ext, proc: x.proc.length };
    }),
    scores: myScores_(j)
  };
}

function judgeScore_(j, p) {
  if (new Date() > JUDGE_CLOSE) return { ok: false, error: '심사 저장 기간이 끝났습니다.' };
  var x = publicEntries_().filter(function (e) { return e.id === String(p.entry); })[0];
  if (!x) return { ok: false, error: '작품을 찾을 수 없습니다.' };
  var s = [p.s1, p.s2, p.s3].map(function (v) { return parseInt(v, 10); });
  if (s.some(function (v) { return !(v >= 1 && v <= 10); })) return { ok: false, error: '세 항목 모두 1~10점으로 매겨 주십시오.' };
  var key = j.id + '|' + x.id;
  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  var row = [now, j.name, x.id, x.track, x.title, x.artist, s[0], s[1], s[2], s[0] + s[1] + s[2],
    p.pick ? 'O' : '', String(p.note || '').slice(0, 2000), key];
  row = row.map(function (v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = scoreSheet_();
    var hit = sh.getRange('M:M').createTextFinder(key).matchEntireCell(true).findNext();
    if (hit) sh.getRange(hit.getRow(), 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, at: now };
}

function judgeImg_(p) {
  var j = judgeAuth_(p.t);
  if (!j) return { ok: false, error: 'auth' };
  var x = publicEntries_().filter(function (e) { return e.id === String(p.id); })[0];
  var n = parseInt(p.n, 10);
  if (!x || !(n >= 0 && n < x.proc.length)) return { ok: false, error: 'not found' };
  return driveImg_(x.proc[n], p.w);
}

/** setupVote() 에서 부름: 심사위원 계정 탭과 심사집계 탭 */
function setupJudge_() {
  var ss = ss_();
  if (!ss.getSheetByName(JUDGE_SHEET)) {
    var sh = ss.insertSheet(JUDGE_SHEET);
    var pw = function () { return Utilities.getUuid().replace(/-/g, '').slice(0, 8); };
    sh.getRange(1, 1, 4, 3).setValues([
      ['아이디', '비밀번호', '이름 (심사 결과에 표시)'],
      ['judge1', pw(), '문영우'],
      ['judge2', pw(), '이동후'],
      ['judge3', pw(), '시포디유저그룹 운영진']
    ]);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
    sh.getRange(6, 1).setValue('줄을 더하면 계정이 늘고, 줄을 지우면 그 계정은 바로 로그인할 수 없습니다. 비밀번호는 심사위원께 따로 전달하십시오.');
  }
  scoreSheet_();
  var t = ss.getSheetByName(SCORE_TALLY) || ss.insertSheet(SCORE_TALLY);
  t.clear();
  t.getRange('A1').setValue('작품별 심사 점수 평균 (30점 만점) · 수록 추천 수 · 심사한 위원 수');
  t.getRange('A3').setFormula("=IFERROR(QUERY('" + SCORE_SHEET + "'!A:L,\"select D, C, E, F, avg(G), avg(H), avg(I), avg(J), count(C) where C is not null group by D, C, E, F order by D, avg(J) desc label avg(G) '완성도', avg(H) '독창성', avg(I) '캘린더 적합성', avg(J) '평균 합계', count(C) '심사 수'\",1),\"아직 심사 없음\")");
  t.getRange('L3').setFormula("=IFERROR(QUERY('" + SCORE_SHEET + "'!A:L,\"select C, count(K) where K = 'O' group by C label C '접수번호', count(K) '수록 추천'\",1),\"\")");
}
