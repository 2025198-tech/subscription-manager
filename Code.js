/**
 * 정기 구독 관리 매니저 - 백엔드 스크립트 (GAS)
 */

// 1. 웹앱 접속 시 메인 HTML 페이지 렌더링
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('정기 구독 관리 매니저')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// 2. 외부 HTML/CSS/JS 파일 로드용 헬퍼 함수 (모듈화 지원)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 3. 구글 시트 데이터베이스 연동 (없으면 자동 생성)
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Subscriptions');
  if (!sheet) {
    sheet = ss.insertSheet('Subscriptions');
    // 헤더 행 생성 및 스타일 적용
    sheet.appendRow(['ID', '서비스명', '카테고리', '금액', '주기', '결제일', '파티원수']);
    sheet.getRange('A1:G1').setFontWeight('bold').setBackground('#f3f3f3');
  }
  return sheet;
}

// 4. 구독 목록 전체 불러오기
function getSubscriptions() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // 데이터가 헤더밖에 없으면 빈 배열 반환
  
  var subscriptions = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    subscriptions.push({
      id: row[0],
      name: row[1],
      category: row[2],
      price: Number(row[3]),
      period: row[4],  // 'monthly' 또는 'yearly'
      payDate: row[5], // 결제일 (1~31)
      members: Number(row[6]) || 1
    });
  }
  return subscriptions;
}

// 5. 새 구독 서비스 추가
function addSubscription(form) {
  var sheet = getSheet();
  var id = new Date().getTime().toString(); // 고유 타임스탬프 ID
  
  // 날짜 형식(YYYY-MM-DD)에서 '일자(Day)' 부분만 추출
  var dayOnly = form.payDate;
  if (typeof form.payDate === 'string' && form.payDate.indexOf('-') !== -1) {
    dayOnly = parseInt(form.payDate.split('-')[2], 10);
  }

  sheet.appendRow([
    id,
    form.name,
    form.category,
    Number(form.price),
    form.period,
    dayOnly,
    Number(form.members) || 1
  ]);
  return { success: true };
}

// 6. 구독 서비스 삭제
function deleteSubscription(id) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: '항목을 찾을 수 없습니다.' };
}
