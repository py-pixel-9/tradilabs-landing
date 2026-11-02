/**
 * 랜딩 문의 수집 → Google Sheet (v2, 2026-11-02)
 *  - 접수일/접수시각 분리, 유입 채널 자동 분류(인스타그램·유튜브·메타광고·네이버·구글·직접)
 *  - 상태 관리 열(드롭다운) + 행 색상, "요약" 탭(채널별·날짜별·광고별 집계)
 *
 * 설치/갱신:
 *  1. 스프레드시트 > 확장 프로그램 > Apps Script → Code.gs 전체를 이 파일로 교체 → 저장
 *  2. 함수 선택 "setupSheet" → 실행 (최초 1회 권한 승인) → leads/요약 탭 생성·서식 적용
 *  3. 배포 > 배포 관리 > 연필 > 버전 "새 버전" → 배포  (처음이면 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스: 모든 사용자)
 *  4. 웹 앱 URL 을 index.html 의 CONFIG.ENDPOINT 에
 */

var BRAND = { name: "Tradilabs", service: "tradilabs-leads", notifyTo: "tradilabs@gmail.com" };

var SHEET_NAME = "leads";
var SUMMARY_NAME = "요약";
var STATUSES = ["신규", "연락완료", "체험중", "구매", "보류", "무효"];
var HEADERS = [
  "접수일", "접수시각", "채널", "광고(ad)", "이름", "연락처", "SNS아이디/메모", "상태", "담당 메모",
  "사이트", "캠페인", "광고세트", "utm_source", "utm_medium", "utm_content",
  "유입경로(referrer)", "첫 랜딩 URL", "제출 페이지", "기기(UA)", "접수(원본)"
];
var COL = {}; HEADERS.forEach(function (h, i) { COL[h] = i + 1; });

function doGet() {
  return json({ ok: true, service: BRAND.service, version: 2 });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var p = {};
    try { p = JSON.parse(e.postData.contents || "{}"); } catch (err) { p = e.parameter || {}; }
    if (!p.name || !p.phone) return json({ ok: false, error: "name/phone required" });

    var phone = String(p.phone).replace(/[^0-9]/g, "").replace(/^(01[016789])(\d{3,4})(\d{4})$/, "$1-$2-$3");
    if (!/^010-\d{4}-\d{4}$/.test(phone)) return json({ ok: false, error: "invalid phone" });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ensureLeads_(ss, false);
    var now = new Date();
    var tz = "Asia/Seoul";
    var memo = [p.memo || "", p.interest || ""].filter(Boolean).join(" · ");
    var row = [
      Utilities.formatDate(now, tz, "yyyy-MM-dd"), Utilities.formatDate(now, tz, "HH:mm"),
      channelOf_(p), p.ad || "", p.name, phone, memo, "신규", "",
      p.site || "", p.campaign || "", p.adset || "", p.utm_source || "", p.utm_medium || "", p.utm_content || "",
      p.referrer || "", p.landing || "", p.page || "", p.ua || "", now
    ];
    sh.appendRow(row);
    var r = sh.getLastRow();
    sh.getRange(r, COL["연락처"]).setNumberFormat("@");
    sh.getRange(r, COL["접수(원본)"]).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    notify_(p, phone, row[2]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

/** 유입 채널 분류 */
function channelOf_(p) {
  var s = [p.utm_source, p.utm_medium, p.ad, p.campaign, p.referrer, p.landing, p.ua].join(" ").toLowerCase();
  if (/instagram|insta|ig_|igshid/.test(s)) return "인스타그램";
  if (/youtube|youtu\.be|yt_/.test(s)) return "유튜브";
  if (/facebook|fb\.|meta|fbclid/.test(s)) return "메타광고";
  if (/naver|blog\.naver|cafe\.naver/.test(s)) return "네이버";
  if (/google|gclid/.test(s)) return "구글";
  if (/telegram|t\.me/.test(s)) return "텔레그램";
  if (/kakao/.test(s)) return "카카오";
  return "직접/기타";
}

/** 최초 1회 또는 갱신 시 수동 실행: 시트 생성·서식·요약 탭 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureLeads_(ss, true);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#1F3A5F").setFontColor("#FFFFFF")
    .setVerticalAlignment("middle").setWrap(false);
  sh.setFrozenRows(1); sh.setFrozenColumns(2); sh.setRowHeight(1, 32);
  var widths = { "접수일": 95, "접수시각": 70, "채널": 90, "광고(ad)": 120, "이름": 90, "연락처": 120, "SNS아이디/메모": 180, "상태": 90, "담당 메모": 220,
    "사이트": 110, "캠페인": 120, "광고세트": 120, "utm_source": 100, "utm_medium": 90, "utm_content": 110, "유입경로(referrer)": 180, "첫 랜딩 URL": 220, "제출 페이지": 220, "기기(UA)": 160, "접수(원본)": 140 };
  Object.keys(widths).forEach(function (h) { sh.setColumnWidth(COL[h], widths[h]); });
  var maxR = Math.max(sh.getMaxRows(), 1000);
  if (sh.getMaxRows() < maxR) sh.insertRowsAfter(sh.getMaxRows(), maxR - sh.getMaxRows());
  sh.getRange(2, COL["연락처"], maxR - 1).setNumberFormat("@");
  sh.getRange(2, COL["접수(원본)"], maxR - 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  sh.getRange(2, 1, maxR - 1, HEADERS.length).setVerticalAlignment("middle");
  sh.getRange(2, COL["접수일"], maxR - 1, 2).setHorizontalAlignment("center");
  sh.getRange(2, COL["채널"], maxR - 1).setHorizontalAlignment("center");
  sh.getRange(2, COL["상태"], maxR - 1).setHorizontalAlignment("center");
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, COL["상태"], maxR - 1).setDataValidation(rule);
  var rng = sh.getRange(2, 1, maxR - 1, HEADERS.length);
  var sL = colLetter_(COL["상태"]);
  var cf = [];
  [["신규", "#FFF4D6"], ["연락완료", "#E8F0FE"], ["체험중", "#E6F4EA"], ["구매", "#CDE8D5"], ["보류", "#F1F3F4"], ["무효", "#FCE8E6"]].forEach(function (pair) {
    cf.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + sL + '2="' + pair[0] + '"').setBackground(pair[1]).setRanges([rng]).build());
  });
  var chRng = sh.getRange(2, COL["채널"], maxR - 1);
  [["인스타그램", "#F8D7E8"], ["유튜브", "#F9D4D4"], ["메타광고", "#D6E4FF"], ["네이버", "#D5F5DD"], ["구글", "#FFF0C2"], ["직접/기타", "#EEEEEE"]].forEach(function (pair) {
    cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(pair[0]).setBackground(pair[1]).setBold(true).setRanges([chRng]).build());
  });
  sh.setConditionalFormatRules(cf);
  if (!sh.getFilter()) sh.getRange(1, 1, maxR, HEADERS.length).createFilter();
  ["utm_medium", "utm_content", "기기(UA)"].forEach(function (h) { sh.hideColumns(COL[h]); });
  buildSummary_(ss);
  ss.setActiveSheet(sh);
  var s1 = ss.getSheetByName("시트1") || ss.getSheetByName("Sheet1");
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

function ensureLeads_(ss, migrate) {
  var sh = ss.getSheetByName(SHEET_NAME);
  if (sh && sh.getLastRow() > 0) {
    var first = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    var isOld = first[0] !== HEADERS[0];
    if (isOld && migrate) { sh.setName(SHEET_NAME + "_v1_" + Utilities.formatDate(new Date(), "Asia/Seoul", "MMdd")); sh = null; }
  }
  if (!sh) { sh = ss.insertSheet(SHEET_NAME, 0); }
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); }
  return sh;
}

function buildSummary_(ss) {
  var s = ss.getSheetByName(SUMMARY_NAME);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(SUMMARY_NAME, 1);
  var L = SHEET_NAME;
  var d = colLetter_(COL["접수일"]), ch = colLetter_(COL["채널"]), ad = colLetter_(COL["광고(ad)"]), st = colLetter_(COL["상태"]);
  s.getRange("A1").setValue(BRAND.name + " 문의 요약").setFontSize(16).setFontWeight("bold");
  s.getRange("A2").setValue("자동 집계 · leads 탭 기준 · 이 탭은 수정하지 마세요").setFontColor("#888888");
  s.getRange("A4:D4").setValues([["전체 문의", "오늘", "최근 7일", "구매 전환"]]).setFontWeight("bold").setBackground("#1F3A5F").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  s.getRange("A5").setFormula("=COUNTA(" + L + "!E2:E)");
  s.getRange("B5").setFormula("=COUNTIF(" + L + "!" + d + "2:" + d + ",TEXT(TODAY(),\"yyyy-mm-dd\"))");
  s.getRange("C5").setFormula("=COUNTIFS(" + L + "!" + d + "2:" + d + ",\">=\"&TEXT(TODAY()-6,\"yyyy-mm-dd\"))");
  s.getRange("D5").setFormula("=COUNTIF(" + L + "!" + st + "2:" + st + ",\"구매\")");
  s.getRange("A5:D5").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");
  s.getRange("A7:C7").setValues([["채널", "문의 수", "구매"]]).setFontWeight("bold").setBackground("#E8EDF3");
  s.getRange("A8").setFormula("=IFERROR(SORT(UNIQUE(FILTER(" + L + "!" + ch + "2:" + ch + "," + L + "!" + ch + "2:" + ch + "<>\"\"))),\"\")");
  s.getRange("B8").setFormula("=ARRAYFORMULA(IF(A8:A20=\"\",\"\",COUNTIF(" + L + "!" + ch + "2:" + ch + ",A8:A20)))");
  s.getRange("C8").setFormula("=ARRAYFORMULA(IF(A8:A20=\"\",\"\",COUNTIFS(" + L + "!" + ch + "2:" + ch + ",A8:A20," + L + "!" + st + "2:" + st + ",\"구매\")))");
  s.getRange("E7:F7").setValues([["광고(ad)", "문의 수"]]).setFontWeight("bold").setBackground("#E8EDF3");
  s.getRange("E8").setFormula("=IFERROR(SORT(UNIQUE(FILTER(" + L + "!" + ad + "2:" + ad + "," + L + "!" + ad + "2:" + ad + "<>\"\"))),\"\")");
  s.getRange("F8").setFormula("=ARRAYFORMULA(IF(E8:E40=\"\",\"\",COUNTIF(" + L + "!" + ad + "2:" + ad + ",E8:E40)))");
  s.getRange("H7:J7").setValues([["날짜", "문의 수", "요일"]]).setFontWeight("bold").setBackground("#E8EDF3");
  s.getRange("H8").setFormula("=ARRAYFORMULA(TEXT(TODAY()-SEQUENCE(30,1,0,1),\"yyyy-mm-dd\"))");
  s.getRange("I8").setFormula("=ARRAYFORMULA(IF(H8:H37=\"\",\"\",COUNTIF(" + L + "!" + d + "2:" + d + ",H8:H37)))");
  s.getRange("J8").setFormula("=ARRAYFORMULA(IF(H8:H37=\"\",\"\",TEXT(DATEVALUE(H8:H37),\"ddd\")))");
  s.setColumnWidths(1, 10, 110); s.setColumnWidth(1, 130); s.setColumnWidth(5, 160); s.setColumnWidth(8, 120);
  s.getRange("B8:C20").setHorizontalAlignment("center"); s.getRange("F8:F40").setHorizontalAlignment("center"); s.getRange("I8:J37").setHorizontalAlignment("center");
  var chart = s.newChart().setChartType(Charts.ChartType.COLUMN).addRange(s.getRange("H8:I37"))
    .setPosition(7, 12, 0, 0).setOption("title", "최근 30일 일별 문의").setOption("legend", { position: "none" }).setOption("width", 520).setOption("height", 260).build();
  s.insertChart(chart);
}

function colLetter_(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

function notify_(p, phone, channel) {
  if (!BRAND.notifyTo) return;
  try {
    MailApp.sendEmail({
      to: BRAND.notifyTo,
      subject: "[" + BRAND.name + " 문의] " + p.name + " / " + channel + " / " + (p.ad || "-"),
      body: "이름: " + p.name + "\n연락처: " + phone + "\n채널: " + channel + "\n메모: " + (p.memo || "-") +
        "\n광고: " + (p.ad || "-") + " / 캠페인: " + (p.campaign || "-") + " / 세트: " + (p.adset || "-") +
        "\n유입: " + (p.utm_source || "-") + " / " + (p.referrer || "-") + "\n랜딩: " + (p.landing || "-") + "\n"
    });
  } catch (err) {}
}

function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
