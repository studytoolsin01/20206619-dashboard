/**
 * Google Apps Script - 영업 및 재고 대시보드 백엔드 (Code.gs)
 * 
 * 이 파일은 웹앱의 진입점(doGet)과 스프레드시트 데이터를 읽어와 
 * 프론트엔드로 전달하는 API 역할을 수행합니다.
 */

/**
 * 웹앱 진입점 함수
 * 웹 브라우저에서 웹앱 URL로 접속 시 Index.html을 렌더링하여 반환합니다.
 */
function doGet(e) {
  // 1. 외부 API 데이터 요청 대응 (Netlify 등 외부 fetch 통신용 JSON 반환)
  if (e && e.parameter && e.parameter.action === 'getData') {
    var data = getDashboardData();
    return ContentService.createTextOutput(JSON.stringify(data))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // 2. 기존 HTML 서비스 응답 (웹앱으로 다이렉트 접속 시)
  var output = HtmlService.createHtmlOutputFromFile('Index');
  
  // 보안 및 iframe 렌더링 설정
  output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  try {
    // 이전 API 하위 호환을 위해 추가 (현재는 기본값으로 작동하지만 요구사항 준수)
    output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  } catch (err) {
    Logger.log("SandboxMode 설정 에러 (최신 GAS 런타임에서는 기본 무시될 수 있음): " + err.message);
  }
  
  // 브라우저 탭 타이틀 설정
  output.setTitle('📊 AppSheet 영업·재고 통합 대시보드');
  
  return output;
}

/**
 * 대시보드에 필요한 모든 데이터를 하나의 JSON 객체로 가공하여 반환합니다.
 * 프론트엔드에서 비동기 호출(google.script.run)로 데이터를 가져갑니다.
 */
function getDashboardData() {
  try {
    // -------------------------------------------------------------
    // 방법 1: 현재 바인딩된 컨테이너 스프레드시트 사용 (기본값)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 방법 2: 스프레드시트 ID로 별도 오픈하는 경우 (웹앱 배포 환경에 따라 아래 주석 해제)
    // var ss = SpreadsheetApp.openById('스프레드시트_ID_입력');
    // -------------------------------------------------------------
    
    if (!ss) {
      throw new Error("스프레드시트를 열 수 없습니다. 활성화된 시트가 존재하지 않거나 권한이 없습니다.");
    }
    
    // 각 시트별 데이터 가공 함수 호출
    var kpis = getKpiData(ss);
    var sales = getSalesChartData(ss);
    var inventory = getInventoryData(ss);
    
    // 최종 가공된 데이터를 묶어서 반환
    return {
      success: true,
      kpis: kpis,
      sales: sales,
      inventory: inventory,
      // 데이터 최종 로드 일시 추가
      updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
    };
    
  } catch (error) {
    Logger.log("데이터 추출 중 에러 발생: " + error.toString());
    return {
      success: false,
      error: "데이터를 읽는 중 오류가 발생했습니다: " + error.message
    };
  }
}

/**
 * KPI 시트에서 주요 요약 지표를 동적으로 추출합니다.
 * A열의 지표명을 기준으로 값을 찾아 매핑합니다 (행 번호 하드코딩 금지).
 */
function getKpiData(ss) {
  var sheet = ss.getSheetByName("KPI");
  if (!sheet) {
    throw new Error("'KPI' 시트를 찾을 수 없습니다. 시트명을 확인해 주세요.");
  }
  
  var rangeValues = sheet.getDataRange().getValues();
  var kpiData = {};
  
  // 1번째 행은 헤더이므로 index = 1부터 시작
  for (var i = 1; i < rangeValues.length; i++) {
    var row = rangeValues[i];
    
    // 지표명 (A열)이 없는 빈 행 스킵
    var metricName = row[0] ? row[0].toString().trim() : "";
    if (metricName === "") continue;
    
    var value = row[1]; // 수식/실제값 (B열)
    var unit = row[2] ? row[2].toString().trim() : ""; // 단위 (C열)
    var desc = row[3] ? row[3].toString().trim() : ""; // 비고 (D열)
    
    // 값의 데이터 타입 정리 (텍스트로 된 백분율 또는 콤마 등을 수치화)
    var parsedValue = value;
    if (typeof value === 'string') {
      if (value.indexOf('%') !== -1) {
        parsedValue = parseFloat(value.replace(/[^0-9.-]/g, ''));
      } else {
        var cleanStr = value.replace(/[^0-9.-]/g, '');
        if (!isNaN(cleanStr) && cleanStr !== '') {
          parsedValue = parseFloat(cleanStr);
        }
      }
    } else if (typeof value === 'number') {
      parsedValue = value;
    }
    
    kpiData[metricName] = {
      value: parsedValue,
      originalValue: value,
      unit: unit,
      desc: desc
    };
  }
  
  return kpiData;
}

/**
 * Sales 시트의 판매 정보를 정제하고 가공합니다.
 * 월별 매출 트렌드, 카테고리별 매출, 채널별 매출, 지역별 매출(상위5) 집계
 */
function getSalesChartData(ss) {
  var sheet = ss.getSheetByName("Sales");
  if (!sheet) {
    throw new Error("'Sales' 시트를 찾을 수 없습니다. 시트명을 확인해 주세요.");
  }
  
  var rangeValues = sheet.getDataRange().getValues();
  
  var monthlySales = {};
  var categorySales = {};
  var channelSales = {};
  var regionSales = {};
  
  // 헤더 매핑 정보 정의 (에러 예방 및 유지보수용)
  // 주문번호(0), 주문일자(1), 제품코드(2), 제품명(3), 카테고리(4), 수량(5), 단가(6), 할인율(7), 판매금액(8), 고객명(9), 지역(10), 판매채널(11), 담당자(12)
  for (var i = 1; i < rangeValues.length; i++) {
    var row = rangeValues[i];
    
    // 주문번호(A열)이 없는 빈 행 스킵
    if (!row[0] || row[0].toString().trim() === "") continue;
    
    // 1. 주문일자 파싱 (월별 트렌드용)
    var dateVal = row[1];
    var yearMonth = "기타";
    if (dateVal) {
      if (dateVal instanceof Date) {
        yearMonth = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM");
      } else {
        var dateStr = dateVal.toString().trim();
        if (dateStr.length >= 7) {
          yearMonth = dateStr.substring(0, 7); // "YYYY-MM" 추출
        }
      }
    }
    
    // 2. 항목별 데이터 매핑 및 숫자 변환
    var category = row[4] ? row[4].toString().trim() : "미분류";
    var qty = Number(row[5]) || 0;
    var price = Number(row[6]) || 0;
    var discount = Number(row[7]) || 0;
    var salesAmount = Number(row[8]) || 0; // 판매금액
    var region = row[10] ? row[10].toString().trim() : "기타";
    var channel = row[11] ? row[11].toString().trim() : "기타";
    
    // 3. 월별 판매금액 누적
    monthlySales[yearMonth] = (monthlySales[yearMonth] || 0) + salesAmount;
    
    // 4. 카테고리별 판매금액 누적
    categorySales[category] = (categorySales[category] || 0) + salesAmount;
    
    // 5. 판매채널별 판매금액 누적
    channelSales[channel] = (channelSales[channel] || 0) + salesAmount;
    
    // 6. 지역별 판매금액 누적
    regionSales[region] = (regionSales[region] || 0) + salesAmount;
  }
  
  // 7. 지역별 판매 매출액 기준 내림차순 정렬 및 상위 5개 추출
  var sortedRegions = Object.keys(regionSales).map(function(key) {
    return { name: key, value: regionSales[key] };
  }).sort(function(a, b) {
    return b.value - a.value;
  });
  
  var top5Regions = sortedRegions.slice(0, 5);
  
  return {
    monthly: monthlySales,
    category: categorySales,
    channel: channelSales,
    region: top5Regions
  };
}

/**
 * Inventory 시트의 재고 정보를 정제하고 가공합니다.
 * 재고상태별 분포(도넛 차트), 현재재고 vs 안전재고 비교(그룹 바 차트), 재고 경고 목록 가공
 */
function getInventoryData(ss) {
  var sheet = ss.getSheetByName("Inventory");
  if (!sheet) {
    throw new Error("'Inventory' 시트를 찾을 수 없습니다. 시트명을 확인해 주세요.");
  }
  
  var rangeValues = sheet.getDataRange().getValues();
  
  var statusCounts = { "정상": 0, "주의": 0, "부족": 0 };
  var warnings = [];
  var chartData = [];
  
  // 제품코드(0), 제품명(1), 카테고리(2), 총입고(3), 총출고(4), 현재재고(5), 안전재고(6), 재고상태(7)
  for (var i = 1; i < rangeValues.length; i++) {
    var row = rangeValues[i];
    
    // 제품코드(A열)가 비어 있는 경우 스킵
    if (!row[0] || row[0].toString().trim() === "") continue;
    
    var code = row[0].toString().trim();
    var name = row[1].toString().trim();
    var category = row[2].toString().trim();
    var currentStock = Number(row[5]) || 0;
    var safetyStock = Number(row[6]) || 0;
    var status = row[7] ? row[7].toString().trim() : "정상";
    
    // 재고 상태 카운트 집계
    if (statusCounts.hasOwnProperty(status)) {
      statusCounts[status]++;
    } else {
      // 정의되지 않은 상태가 들어올 경우 동적 생성 또는 스킵
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    }
    
    // 차트 5번을 위해 전 제품의 현재재고와 안전재고 데이터 저장
    chartData.push({
      name: name,
      current: currentStock,
      safety: safetyStock
    });
    
    // 재고 경고 기준: 현재재고 < 안전재고 * 1.5인 품목
    if (currentStock < safetyStock * 1.5) {
      // 재고 여유율 계산: (현재재고 / 안전재고 - 1) * 100
      var safetyMargin = 0;
      if (safetyStock > 0) {
        safetyMargin = ((currentStock / safetyStock) - 1) * 100;
      }
      
      warnings.push({
        code: code,
        name: name,
        category: category,
        current: currentStock,
        safety: safetyStock,
        status: status,
        margin: safetyMargin.toFixed(1) // 소수점 1자리 문자열 변환
      });
    }
  }
  
  // 재고 여유율이 낮은 품목(가장 시급한 품목) 순으로 경고 테이블 정렬
  warnings.sort(function(a, b) {
    return parseFloat(a.margin) - parseFloat(b.margin);
  });
  
  return {
    statusCounts: statusCounts,
    warnings: warnings,
    chartData: chartData
  };
}
