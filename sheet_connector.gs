// =====================================================
// 마케팅해줘 상담 신청 자동 저장 스크립트
// 1단계: 이 코드 전체 붙여넣기
// 2단계: setupSheet() 실행 → 구글시트 자동 생성
// 3단계: doPost 함수로 웹앱 배포 → URL 복사
// =====================================================

// ✅ 1단계: 시트 자동 생성 (처음 한 번만 실행)
function setupSheet() {
  // 새 스프레드시트 생성
  var ss = SpreadsheetApp.create("마케팅해줘 상담 신청 관리");
  var sheet = ss.getActiveSheet();
  sheet.setName("상담신청");

  // 헤더 설정
  var headers = ["신청일시", "이름", "연락처", "업종", "현재마케팅방식", "문의내용", "상태"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 헤더 스타일
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#E8202A");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(12);

  // 열 너비 자동 조정
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 250);
  sheet.setColumnWidth(7, 100);

  // 시트 ID 저장
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  Logger.log("✅ 구글시트 생성 완료!");
  Logger.log("📋 시트 URL: " + ss.getUrl());
  Logger.log("🔑 시트 ID: " + ss.getId());
  Logger.log("");
  Logger.log("👉 다음 단계: 배포 > 웹 앱으로 배포 클릭!");
}

// =====================================================
// ✅ 2단계: 이 함수가 실제 폼 데이터를 받아서 저장
// (웹 앱으로 배포한 뒤 자동으로 실행됨 - 직접 실행 X)
// =====================================================
function doPost(e) {
  try {
    var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName("상담신청");

    // 데이터 파싱
    var data = JSON.parse(e.postData.contents);

    // 시트에 행 추가
    sheet.appendRow([
      data.date    || new Date().toLocaleString('ko-KR'),
      data.name    || "",
      data.contact || "",
      data.type    || "",
      data.current || "",
      data.message || "",
      "신규"  // 기본 상태값
    ]);

    // 새 행 스타일
    var lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, 7).setBackground("#FFF5F5");
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================
// ✅ 선택: 테스트용 - 실제 데이터가 잘 들어가는지 확인
// =====================================================
function testSave() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("상담신청");

  sheet.appendRow([
    new Date().toLocaleString('ko-KR'),
    "테스트 홍길동",
    "010-1234-5678",
    "세무사 / 회계사",
    "블로그 대행사 이용 중",
    "마케팅 상담 문의드립니다.",
    "신규"
  ]);

  Logger.log("✅ 테스트 데이터 저장 완료! 시트 확인해보세요.");
}
