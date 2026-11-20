/**
 * 셀러팩트 접수 큐 — Google Apps Script 웹앱 (무료·슬립 없음)
 *
 * 역할
 *  1) 고객 페이지(docs/index.html)가 POST로 보낸 CSV를 받아 큐(스프레드시트)에 적재
 *  2) 유료 주문은 Gumroad 라이선스 키를 실시간 검증(1키=1회)
 *  3) GitHub Actions(10분 스케줄)가 GET으로 큐를 소비하고 결과를 ack
 *  4) ack된 건은 구매자 이메일로 리포트 링크 발송
 *
 * 배포: 웹앱 → 실행 계정 "나", 액세스 "모든 사용자"
 * 스크립트 속성
 *   QUEUE_KEY        (필수) 큐 공유 키 — GitHub 저장소 시크릿 QUEUE_KEY와 동일
 *   GUMROAD_TOKEN    (유료 판매 시) Gumroad 액세스 토큰
 *   GUMROAD_PRODUCT_ID (유료 판매 시) 셀러팩트 상품 ID
 *   REPORT_BASE      (선택) 기본 https://neo970720.github.io/sellerfact
 *   NOTIFY_EMAIL     (선택) 운영 알림 받을 메일
 */

var MAX_SKUS = 300;
var MAX_B64 = 60000;           // base64 기준(원문 약 44KB)
var DEMO_DAILY_CAP = 20;       // 무료 체험 일일 상한(남용 방지)

function props() { return PropertiesService.getScriptProperties(); }
function reportBase() { return props().getProperty('REPORT_BASE') || 'https://neo970720.github.io/sellerfact'; }

/* ------------------------------------------------------------------ 저장소 */
function db() {
  var id = props().getProperty('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (err) { /* 재생성 */ }
  }
  var ss = SpreadsheetApp.create('셀러팩트 접수 큐');
  var sh = ss.getSheets()[0];
  sh.setName('jobs');
  sh.appendRow(['job', 'created_at', 'client', 'email', 'mode', 'license_hash', 'skus',
                'status', 'message', 'url', 'finished_at', 'notified']);
  sh.setFrozenRows(1);
  var payload = ss.insertSheet('payloads');
  payload.appendRow(['job', 'csv_b64']);
  payload.hideSheet();
  props().setProperty('SHEET_ID', ss.getId());
  return ss;
}

function jobsSheet() { return db().getSheetByName('jobs'); }
function payloadSheet() { return db().getSheetByName('payloads'); }

function rows() {
  var sh = jobsSheet();
  var v = sh.getDataRange().getValues();
  var head = v.shift();
  return v.map(function (r, i) {
    var o = {row: i + 2};
    head.forEach(function (h, j) { o[h] = r[j]; });
    return o;
  });
}

function findRow(job) {
  var found = null;
  rows().forEach(function (o) { if (o.job === job) found = o; });
  return found;
}

/* ------------------------------------------------------------------ 유틸 */
function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function clean(v, max) {
  var s = String(v == null ? '' : v).trim();
  return s.length > max ? s.slice(0, max) : s;
}
function sanitizeJob(v) {
  return clean(v, 40).replace(/[^A-Za-z0-9_-]/g, '');
}
function hash(s) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('').slice(0, 32);
}
function nowIso() { return Utilities.formatDate(new Date(), 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm:ssXXX"); }
function keyOk(k) {
  var want = props().getProperty('QUEUE_KEY');
  return !!want && String(k || '') === want;
}

/* ------------------------------------------------------------------ POST: 접수 */
function doPost(e) {
  try {
    var p = (e && e.parameter) || {};
    var job = sanitizeJob(p.job);
    if (!job) return out({ok: false, state: 'error', message: '작업 ID가 없습니다.'});
    if (findRow(job)) return out({ok: true, job: job, state: 'duplicate'});

    var client = clean(p.client, 40) || '셀러';
    var email = clean(p.email, 120);
    var b64 = String(p.csv_b64 || '');
    if (!b64) return out({ok: false, state: 'error', message: '상품 목록이 비어 있습니다.'});
    if (b64.length > MAX_B64) return out({ok: false, state: 'error', message: '입력이 너무 큽니다(약 44KB 초과). SKU를 나눠 요청해 주세요.'});

    var csv = '';
    try { csv = Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString('UTF-8'); }
    catch (err) { return out({ok: false, state: 'error', message: 'CSV를 읽지 못했습니다(인코딩 확인 필요).'}); }

    var lines = csv.split(/\r?\n/).filter(function (l) { return l.trim(); });
    if (lines.length < 2) return out({ok: false, state: 'error', message: '데이터 행이 없습니다.'});
    var header = lines[0].toLowerCase();
    if (header.indexOf('sku_id') < 0 || header.indexOf('product_name') < 0) {
      return out({ok: false, state: 'error', message: '첫 줄이 헤더여야 합니다(sku_id, product_name, …).'});
    }
    var skuCount = lines.length - 1;
    if (skuCount > MAX_SKUS) return out({ok: false, state: 'error', message: '한 번에 ' + MAX_SKUS + '개까지 점검합니다.'});

    var license = clean(p.license, 120);
    var mode, licHash = '';
    if (license) {
      var v = verifyLicense(license);
      if (!v.ok) return out({ok: false, state: 'error', message: v.message, hint: v.hint || ''});
      licHash = hash(license);
      var used = rows().some(function (o) { return o.license_hash === licHash && o.status !== 'error'; });
      if (used) {
        return out({ok: false, state: 'error',
          message: '이미 사용된 라이선스 키입니다.', hint: '키 1개 = 점검 1회입니다. 추가 점검은 새 키를 구매해 주세요.'});
      }
      mode = 'paid';
      if (v.email && !email) email = v.email;
    } else {
      if (skuCount > 2) {
        return out({ok: false, state: 'error', message: '무료 체험은 2 SKU까지입니다.',
          hint: '전체 점검은 라이선스 키를 입력해 주세요.'});
      }
      if (!demoQuotaLeft()) {
        return out({ok: false, state: 'error', message: '오늘의 무료 체험 한도가 모두 찼습니다.', hint: '내일 다시 시도해 주세요.'});
      }
      mode = 'demo';
    }

    jobsSheet().appendRow([job, nowIso(), client, email, mode, licHash, skuCount, 'queued', '', '', '', '']);
    payloadSheet().appendRow([job, b64]);
    if (mode === 'demo') bumpDemoCounter();
    notify('새 접수', job + ' / ' + mode + ' / ' + client + ' / ' + skuCount + ' SKU');
    dispatchRunner(job);   // 즉시 실행 요청(실패해도 10분 스케줄이 처리)
    return out({ok: true, job: job, state: 'queued'});
  } catch (err) {
    notify('접수 오류', String(err));
    return out({ok: false, state: 'error', message: '접수 중 오류가 발생했습니다.'});
  }
}

/**
 * GitHub Actions 실행을 즉시 요청한다(스크립트 속성 GITHUB_TOKEN: actions:write 전용 토큰).
 * 실패해도 큐에 남아 있으므로 10분 스케줄이 이어서 처리한다.
 */
function dispatchRunner(job) {
  var token = props().getProperty('GITHUB_TOKEN');
  if (!token) return false;
  var repo = props().getProperty('GITHUB_REPO') || 'Neo970720/sellerfact';
  var url = 'https://api.github.com/repos/' + repo + '/actions/workflows/audit.yml/dispatches';
  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28'
      },
      payload: JSON.stringify({ref: 'main', inputs: {mode: 'queue'}}),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code >= 300) notify('실행 요청 실패', job + ' / HTTP ' + code + ' / ' + res.getContentText().slice(0, 200));
    return code < 300;
  } catch (err) {
    notify('실행 요청 오류', job + ' / ' + String(err));
    return false;
  }
}

function verifyLicense(key) {
  var token = props().getProperty('GUMROAD_TOKEN');
  var productId = props().getProperty('GUMROAD_PRODUCT_ID');
  if (!token || !productId) {
    return {ok: false, message: '유료 결제 연동이 아직 준비되지 않았습니다.', hint: '무료 체험은 라이선스 키 없이 2 SKU까지 가능합니다.'};
  }
  var url = 'https://api.gumroad.com/v2/licenses/verify?product_id=' + encodeURIComponent(productId) +
            '&license_key=' + encodeURIComponent(key) + '&increment_uses_count=false';
  try {
    var res = UrlFetchApp.fetch(url, {method: 'post', muteHttpExceptions: true});
    var body = JSON.parse(res.getContentText());
    if (!body.success) return {ok: false, message: '라이선스 키를 확인할 수 없습니다.', hint: '구매 메일의 키를 그대로 붙여넣었는지 확인해 주세요.'};
    var purchase = body.purchase || {};
    if (purchase.refunded) return {ok: false, message: '환불된 주문의 키입니다.'};
    if (purchase.chargebacked) return {ok: false, message: '지불 거절된 주문의 키입니다.'};
    return {ok: true, email: purchase.email || ''};
  } catch (err) {
    return {ok: false, message: '라이선스 확인 서버에 연결하지 못했습니다.', hint: '잠시 후 다시 시도해 주세요.'};
  }
}

function demoQuotaLeft() {
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('demo:' + today) || 0);
  return n < DEMO_DAILY_CAP;
}
function bumpDemoCounter() {
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('demo:' + today) || 0) + 1;
  cache.put('demo:' + today, String(n), 60 * 60 * 26);
}

/* ------------------------------------------------------------------ GET: 큐 소비·상태 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'status';

  if (action === 'status') {
    var job = sanitizeJob(p.job);
    var r = job ? findRow(job) : null;
    if (!r) return out({state: 'unknown'});
    var state = r.status === 'done' ? 'done' : (r.status === 'error' ? 'error' : 'queued');
    return out({state: state, message: r.message || ''});
  }

  if (!keyOk(p.key)) return out({ok: false, error: 'unauthorized'});

  if (action === 'queue') {
    var pending = rows().filter(function (o) { return o.status === 'queued'; }).slice(0, 20);
    return out({ok: true, jobs: pending.map(function (o) {
      return {job: o.job, client: o.client, skus: o.skus, mode: o.mode};
    })});
  }

  if (action === 'payload') {
    var j = sanitizeJob(p.job);
    var d = payloadSheet().getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === j) {
        var r2 = findRow(j);
        return out({ok: true, job: j, client: r2 ? r2.client : '', csv_b64: d[i][1]});
      }
    }
    return out({ok: false, error: 'not_found'});
  }

  if (action === 'ack') {
    var jj = sanitizeJob(p.job);
    var row = findRow(jj);
    if (!row) return out({ok: false, error: 'not_found'});
    var sheet = jobsSheet();
    var status = p.status === 'done' ? 'done' : 'error';
    sheet.getRange(row.row, 8).setValue(status);
    sheet.getRange(row.row, 9).setValue(clean(p.message, 300));
    sheet.getRange(row.row, 10).setValue(clean(p.url, 300) || (reportBase() + '/view.html?job=' + jj));
    sheet.getRange(row.row, 11).setValue(nowIso());
    if (status === 'done') mailReport(row);
    return out({ok: true});
  }

  if (action === 'stats') {
    var all = rows();
    var by = {};
    all.forEach(function (o) { by[o.status] = (by[o.status] || 0) + 1; });
    return out({ok: true, total: all.length, by: by, sheet: db().getUrl()});
  }

  return out({ok: false, error: 'unknown_action'});
}

/* ------------------------------------------------------------------ 메일 */
function mailReport(row) {
  if (!row.email || row.notified === 'y') return;
  var url = row.url || (reportBase() + '/view.html?job=' + row.job);
  var subject = '[셀러팩트] ' + row.client + ' 상품 점검 리포트가 준비되었습니다';
  var body = [
    '<div style="font-family:-apple-system,\'Malgun Gothic\',sans-serif;color:#1C2434;line-height:1.6">',
    '<h2 style="color:#0F1B2D;margin:0 0 8px">점검이 끝났습니다</h2>',
    '<p>' + row.client + ' · SKU ' + row.skus + '개 상품의 KC 인증·표시광고·고시정보 사전 점검이 완료되었습니다.</p>',
    '<p><a href="' + url + '" style="display:inline-block;background:#0F1B2D;color:#fff;padding:12px 20px;border-radius:8px;text-decoration:none">리포트 열기</a></p>',
    '<p style="font-size:13px;color:#5A6474">PDF·엑셀 파일은 위 페이지에서 바로 내려받을 수 있습니다. '
      + '이 링크를 저장해 두시면 나중에 다시 열 수 있습니다.</p>',
    '<p style="font-size:12px;color:#5A6474">본 리포트는 공개 법령·고시·공공 인증정보 기반 1차 스크리닝이며, '
      + '최종 판단은 지정 시험·인증기관의 확인이 필요합니다.</p>',
    '</div>'
  ].join('');
  try {
    MailApp.sendEmail({to: row.email, subject: subject, htmlBody: body, name: '셀러팩트'});
    jobsSheet().getRange(row.row, 12).setValue('y');
  } catch (err) {
    notify('메일 발송 실패', row.job + ' / ' + String(err));
  }
}

function notify(subject, text) {
  var to = props().getProperty('NOTIFY_EMAIL');
  if (!to) return;
  try { MailApp.sendEmail(to, '[셀러팩트 운영] ' + subject, text); } catch (err) { /* 무시 */ }
}

/* ------------------------------------------------------------------ 정리 트리거 */
/** 10분마다 실행: ack 누락·미발송 메일 정리. 트리거는 installTriggers()로 1회 설치. */
function sweep() {
  var all = rows();
  all.forEach(function (o) {
    if (o.status === 'done' && o.notified !== 'y') mailReport(o);
  });
  var cutoff = Date.now() - 1000 * 60 * 60 * 6;
  all.forEach(function (o) {
    if (o.status === 'queued' && new Date(o.created_at).getTime() < cutoff) {
      jobsSheet().getRange(o.row, 8).setValue('error');
      jobsSheet().getRange(o.row, 9).setValue('6시간 내 처리되지 않아 만료되었습니다.');
    }
  });
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sweep') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sweep').timeBased().everyMinutes(10).create();
}

/** 최초 1회 실행: 시트 생성 + 큐 키 확인 + 트리거 설치. 결과를 반환값으로 돌려준다. */
function setup() {
  db();
  if (!props().getProperty('QUEUE_KEY')) {
    props().setProperty('QUEUE_KEY', Utilities.getUuid().replace(/-/g, ''));
  }
  installTriggers();
  var out = {
    ok: true,
    sheet: db().getUrl(),
    queue_key: props().getProperty('QUEUE_KEY'),
    endpoint: ScriptApp.getService().getUrl()
  };
  Logger.log(JSON.stringify(out));
  return JSON.stringify(out);
}

/** 운영자 설정 주입: setProps(["KEY=VALUE", ...]) — 코드에 비밀값을 남기지 않는다. */
function setProps(list) {
  (list || []).forEach(function (pair) {
    var i = String(pair).indexOf('=');
    if (i > 0) props().setProperty(String(pair).slice(0, i), String(pair).slice(i + 1));
  });
  return JSON.stringify({ok: true, keys: Object.keys(props().getProperties())});
}
