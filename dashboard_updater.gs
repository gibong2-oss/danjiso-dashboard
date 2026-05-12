// ============================================================
// 단지서비스팀 대시보드 자동 업데이트 스크립트
// Google Apps Script - 매일 자동 실행
// ============================================================

// ── 설정 ─────────────────────────────────────────────────────
const SPREADSHEET_ID = "1Aq5wxM4J8eW2zo9_euCI5odRlb4FFJMQMbL0vWnwM4I";

const REDASH = {
  vote:   { url: "http://redash.aptner.com/api/queries/171/results.json",  key: "S3OhKfRpfhWvpdg6jtZNQN1wIapQaBjbgwPY4zKF" },
  notify: { url: "http://redash.aptner.com/api/queries/210/results.json",  key: "nBujW8OYVJIzRAkheXgWjOm7IKq15kax4NdNR9rW" },
  survey: { url: "http://redash.aptner.com/api/queries/239/results.json",  key: "QEzIiGV92lPRVl2IFfukW5FeSuNXWR45rkGE5lna" }
};

const VAT      = 1.1;
const MIN_BILL = 1000;  // 1,000원 미만 발행제외
const NOTIFY_PRICE = { "카카오알림톡": 20, "SMS": 20, "LMS": 40, "MMS": 80 };

// ── 참조 탭 로드 ──────────────────────────────────────────────
function loadReferenceTabs(ss) {
  // 무료단지: { "YYYY-MM": Set([code, ...]) }
  var FREE = {};
  var freeSheet = ss.getSheetByName("무료단지");
  if (freeSheet) {
    var fd = freeSheet.getDataRange().getValues();
    for (var i = 1; i < fd.length; i++) {
      if (!fd[i][0] || !fd[i][1]) continue;
      var mo = String(fd[i][0]).substring(0, 7);
      var code = String(fd[i][1]);
      if (!FREE[mo]) FREE[mo] = {};
      FREE[mo][code] = true;
    }
  }

  // 직방전환단지: { code: [{sd: Date, ed: Date|null}] }
  var ZIGBANG = {};
  var zbSheet = ss.getSheetByName("직방전환단지");
  if (zbSheet) {
    var zd = zbSheet.getDataRange().getValues();
    for (var i = 1; i < zd.length; i++) {
      var code = zd[i][1] ? String(zd[i][1]) : null;
      var sd   = zd[i][3] ? new Date(zd[i][3]) : null;
      var ed   = zd[i][4] ? new Date(zd[i][4]) : null;
      if (!code || !sd) continue;
      if (!ZIGBANG[code]) ZIGBANG[code] = [];
      ZIGBANG[code].push({ sd: sd, ed: ed });
    }
  }

  // 프로모션단지: { "YYYY-MM": { code: quota } }
  var PROMO = {};
  var promoSheet = ss.getSheetByName("프로모션단지");
  if (promoSheet) {
    var pd = promoSheet.getDataRange().getValues();
    for (var i = 1; i < pd.length; i++) {
      if (!pd[i][0] || !pd[i][1]) continue;
      var mo    = String(pd[i][0]).substring(0, 7);
      var code  = String(pd[i][1]);
      var quota = pd[i][3] ? parseInt(pd[i][3]) : 0;
      if (!PROMO[mo]) PROMO[mo] = {};
      PROMO[mo][code] = quota;
    }
  }

  return { FREE: FREE, ZIGBANG: ZIGBANG, PROMO: PROMO };
}

// ── 헬퍼 함수 ─────────────────────────────────────────────────
function isZigbang(code, d, ZIGBANG) {
  var entries = ZIGBANG[code] || [];
  for (var i = 0; i < entries.length; i++) {
    var sd = entries[i].sd, ed = entries[i].ed;
    if (d >= sd && (ed === null || d <= ed)) return true;
  }
  return false;
}

function isFree(code, mo, FREE) {
  return FREE[mo] && FREE[mo][code] === true;
}

// 프로모션 비율 계산: 월별 총 발송 대비 초과분만 청구
function buildPromoRatios(parsedRows, PROMO) {
  // parsedRows: [{mo, code, sends}]
  var moSends = {};
  for (var i = 0; i < parsedRows.length; i++) {
    var mo = parsedRows[i].mo, code = parsedRows[i].code, sends = parsedRows[i].sends;
    if (PROMO[mo] && PROMO[mo][code] !== undefined) {
      if (!moSends[mo]) moSends[mo] = {};
      moSends[mo][code] = (moSends[mo][code] || 0) + sends;
    }
  }
  var ratios = {};
  for (var mo in PROMO) {
    for (var code in PROMO[mo]) {
      var quota = PROMO[mo][code];
      var total = (moSends[mo] && moSends[mo][code]) ? moSends[mo][code] : 0;
      if (!ratios[mo]) ratios[mo] = {};
      ratios[mo][code] = (total <= 0 || total <= quota) ? 0.0 : (total - quota) / total;
    }
  }
  return ratios;
}

// 월별 집계 결과 생성
function makeMonthly(parsedRows, promoRatios) {
  var moAmt  = {};  // mo → 원 합계
  var moCnt  = {};  // mo → 건수 합계
  var moCplx = {};  // mo → Set(code)

  for (var i = 0; i < parsedRows.length; i++) {
    var p = parsedRows[i];
    var mo = p.mo, code = p.code, sends = p.sends, unitPrice = p.unitPrice;
    var ratio = (promoRatios[mo] && promoRatios[mo][code] !== undefined)
                ? promoRatios[mo][code] : 1.0;
    var amt = sends * unitPrice * ratio * VAT;

    if (!moAmt[mo])  { moAmt[mo] = 0; moCnt[mo] = 0; moCplx[mo] = {}; }
    moAmt[mo]  += amt;
    moCnt[mo]  += sends;
    moCplx[mo][code] = true;
  }
  return { moAmt: moAmt, moCnt: moCnt, moCplx: moCplx };
}

// ── Redash API 호출 ───────────────────────────────────────────
function fetchRedash(cfg) {
  var res = UrlFetchApp.fetch(cfg.url + "?api_key=" + cfg.key, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error("Redash 호출 실패: " + cfg.url + " → " + res.getResponseCode());
  }
  return JSON.parse(res.getContentText()).query_result.data.rows;
}

// ── 서비스별 처리 ─────────────────────────────────────────────

// 전자투표: {send_date, kapt_code, sms_send_count, ...}
function processVote(rows, refs) {
  var parsed = [];
  for (var i = 0; i < rows.length; i++) {
    var r    = rows[i];
    var mo   = String(r.send_date).substring(0, 7);
    var d    = new Date(r.send_date);
    var code = String(r.kapt_code || "");
    var sends = parseFloat(r.sms_send_count) || 0;
    if (!code || sends <= 0) continue;
    if (isFree(code, mo, refs.FREE)) continue;
    var unitPrice = isZigbang(code, d, refs.ZIGBANG) ? 20 : 40;
    parsed.push({ mo: mo, code: code, sends: sends, unitPrice: unitPrice });
  }
  var promoRatios = buildPromoRatios(parsed, refs.PROMO);
  return makeMonthly(parsed, promoRatios);
}

// 알림서비스: {reg_datetime, kapt_code, send_type, send_cnt, ...}
function processNotify(rows, refs) {
  var parsed = [];
  for (var i = 0; i < rows.length; i++) {
    var r     = rows[i];
    var dt    = r.reg_datetime ? String(r.reg_datetime) : "";
    var mo    = dt.substring(0, 7);
    var code  = String(r.kapt_code || "");
    var means = String(r.send_type || "LMS");
    var sends = parseFloat(r.send_cnt) || 0;
    if (!code || !mo || sends <= 0) continue;
    if (isFree(code, mo, refs.FREE)) continue;
    var unitPrice = NOTIFY_PRICE[means] !== undefined ? NOTIFY_PRICE[means] : 40;
    parsed.push({ mo: mo, code: code, sends: sends, unitPrice: unitPrice });
  }
  var promoRatios = buildPromoRatios(parsed, refs.PROMO);
  return makeMonthly(parsed, promoRatios);
}

// 설문조사: {send_kakaotalk_date, kapt_code, send_kakaotalk_count, ...}
function processSurvey(rows, refs) {
  var parsed = [];
  for (var i = 0; i < rows.length; i++) {
    var r    = rows[i];
    var dt   = r.send_kakaotalk_date ? String(r.send_kakaotalk_date) : "";
    var mo   = dt.substring(0, 7);
    var code = String(r.kapt_code || "");
    var sends = parseFloat(r.send_kakaotalk_count) || 0;
    if (!code || !mo || sends <= 0) continue;
    if (isFree(code, mo, refs.FREE)) continue;
    parsed.push({ mo: mo, code: code, sends: sends, unitPrice: 40 });
  }
  var promoRatios = buildPromoRatios(parsed, refs.PROMO);
  return makeMonthly(parsed, promoRatios);
}

// ── 메인 함수 (매일 자동 실행) ────────────────────────────────
function updateDashboard() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var refs = loadReferenceTabs(ss);

  Logger.log("참조 로드 완료 - 무료:" + Object.keys(refs.FREE).length + "월, 직방:" + Object.keys(refs.ZIGBANG).length + "단지, 프로모션:" + Object.keys(refs.PROMO).length + "월");

  // Redash 데이터 수신
  var voteRows   = fetchRedash(REDASH.vote);
  var notifyRows = fetchRedash(REDASH.notify);
  var surveyRows = fetchRedash(REDASH.survey);

  Logger.log("Redash 수신 - 전자투표:" + voteRows.length + "행, 알림:" + notifyRows.length + "행, 설문:" + surveyRows.length + "행");

  // 서비스별 처리
  var vote   = processVote(voteRows, refs);
  var notify = processNotify(notifyRows, refs);
  var survey = processSurvey(surveyRows, refs);

  // 전체 월 목록
  var allMoSet = {};
  [vote, notify, survey].forEach(function(svc) {
    Object.keys(svc.moAmt).forEach(function(mo) { allMoSet[mo] = true; });
  });
  var allMonths = Object.keys(allMoSet).sort();

  // 출력 데이터 구성 (Looker Studio용 long format)
  var now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm");
  var output = [["월", "서비스", "건수", "이용단지수", "추정금액(원)", "추정금액(만원)", "업데이트일시"]];

  var services = [
    { name: "전자투표",  data: vote   },
    { name: "알림서비스", data: notify },
    { name: "설문조사",  data: survey }
  ];

  for (var mi = 0; mi < allMonths.length; mi++) {
    var mo = allMonths[mi];
    for (var si = 0; si < services.length; si++) {
      var svc = services[si];
      if (!svc.data.moAmt[mo]) continue;
      var amt    = Math.round(svc.data.moAmt[mo]);
      var cnt    = Math.round(svc.data.moCnt[mo]);
      var cplxCnt = Object.keys(svc.data.moCplx[mo] || {}).length;
      var amt만원 = Math.round(amt / 100) / 100;  // 소수점 2자리 만원
      output.push([mo, svc.name, cnt, cplxCnt, amt, amt만원, now]);
    }
  }

  // 대시보드 탭 작성
  var dashSheet = ss.getSheetByName("대시보드");
  if (!dashSheet) dashSheet = ss.insertSheet("대시보드");
  dashSheet.clearContents();
  dashSheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  // 헤더 스타일
  var headerRange = dashSheet.getRange(1, 1, 1, output[0].length);
  headerRange.setBackground("#4A90D9").setFontColor("white").setFontWeight("bold");
  dashSheet.setFrozenRows(1);

  Logger.log("완료: " + (output.length - 1) + "행 작성 → 대시보드 탭");
}

// ── 트리거 설정 (최초 1회만 실행) ────────────────────────────
function setDailyTrigger() {
  // 기존 트리거 삭제
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "updateDashboard") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 매일 오전 7시 실행 등록
  ScriptApp.newTrigger("updateDashboard")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  Logger.log("트리거 설정 완료: 매일 오전 7시 자동 실행");
}
