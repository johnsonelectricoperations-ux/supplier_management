/**
 * Utils.gs - 공통 유틸리티 함수
 */

/**
 * 날짜 포맷 변환 (yyyy-MM-dd)
 */
function formatDate(date) {
  if (!date) return '';
  
  if (typeof date === 'string') {
    return date;
  }
  
  return Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM-dd');
}

/**
 * 날짜시간 포맷 변환 (yyyy-MM-dd HH:mm:ss)
 */
function formatDateTime(datetime) {
  if (!datetime) return '';
  
  if (typeof datetime === 'string') {
    return datetime;
  }
  
  return Utilities.formatDate(datetime, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

/**
 * 시간 
 */
function getTimeOptions() {
  return ['오전', '오후'];
}

/**
 * 입력값 검증 - 필수값 체크
 */
function validateRequired(value, fieldName) {
  if (!value || value.toString().trim() === '') {
    return {
      valid: false,
      message: `${fieldName}은(는) 필수 입력 항목입니다.`
    };
  }
  return { valid: true };
}

/**
 * 입력값 검증 - 날짜 형식 체크
 */
function validateDate(dateString) {
  const datePattern = /^\d{4}-\d{2}-\d{2}$/;
  
  if (!datePattern.test(dateString)) {
    return {
      valid: false,
      message: '날짜 형식이 올바르지 않습니다. (yyyy-MM-dd)'
    };
  }
  
  const date = new Date(dateString);
  if (isNaN(date.getTime())) {
    return {
      valid: false,
      message: '유효하지 않은 날짜입니다.'
    };
  }
  
  return { valid: true };
}

/**
 * 입력값 검증 - 숫자 체크
 */
function validateNumber(value, fieldName) {
  if (isNaN(value) || value === '') {
    return {
      valid: false,
      message: `${fieldName}은(는) 숫자만 입력 가능합니다.`
    };
  }
  
  if (Number(value) < 0) {
    return {
      valid: false,
      message: `${fieldName}은(는) 0 이상이어야 합니다.`
    };
  }
  
  return { valid: true };
}

/**
 * 데이터 입력 전체 검증
 */
function validateDataInput(dataObj) {
  let errors = [];
  
  // 날짜 검증
  let dateCheck = validateRequired(dataObj.date, '입고날짜');
  if (!dateCheck.valid) {
    errors.push(dateCheck.message);
  } else {
    dateCheck = validateDate(dataObj.date);
    if (!dateCheck.valid) {
      errors.push(dateCheck.message);
    }
  }
  
  // 시간 검증
  let timeCheck = validateRequired(dataObj.time, '입고시간');
  if (!timeCheck.valid) {
    errors.push(timeCheck.message);
  }
  
  // TM-NO 검증
  let tmNoCheck = validateRequired(dataObj.tmNo, 'TM-NO');
  if (!tmNoCheck.valid) {
    errors.push(tmNoCheck.message);
  }
  
  // 제품명 검증
  let productCheck = validateRequired(dataObj.productName, '제품명');
  if (!productCheck.valid) {
    errors.push(productCheck.message);
  }
  
  // 수량 검증
  let quantityCheck = validateRequired(dataObj.quantity, '수량');
  if (!quantityCheck.valid) {
    errors.push(quantityCheck.message);
  } else {
    quantityCheck = validateNumber(dataObj.quantity, '수량');
    if (!quantityCheck.valid) {
      errors.push(quantityCheck.message);
    }
  }
  
  if (errors.length > 0) {
    return {
      valid: false,
      errors: errors
    };
  }
  
  return { valid: true };
}

/**
 * HTML 특수문자 이스케이프
 */
function escapeHtml(text) {
  if (!text) return '';
  
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  
  return text.toString().replace(/[&<>"']/g, function(m) { return map[m]; });
}

/**
 * 에러 로깅
 */
function logError(functionName, error) {
  const timestamp = new Date();
  const errorMessage = `[${timestamp}] ${functionName}: ${error.toString()}`;
  Logger.log(errorMessage);
  
  // 필요시 에러 로그 시트에 기록
  try {
    const ss = getSpreadsheet();
    let errorSheet = ss.getSheetByName('ErrorLog');
    
    if (!errorSheet) {
      errorSheet = ss.insertSheet('ErrorLog');
      errorSheet.getRange('A1:C1').setValues([['날짜시간', '함수명', '에러메시지']]);
      errorSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#ea4335').setFontColor('#ffffff');
    }
    
    errorSheet.appendRow([timestamp, functionName, error.toString()]);
  } catch (e) {
    Logger.log('에러 로그 기록 실패: ' + e.toString());
  }
}

/**
 * 성공 응답 생성
 */
function createSuccessResponse(message, data = null) {
  const response = {
    success: true,
    message: message
  };
  
  if (data !== null) {
    response.data = data;
  }
  
  return response;
}

/**
 * 에러 응답 생성
 */
function createErrorResponse(message, errors = null) {
  const response = {
    success: false,
    message: message
  };
  
  if (errors !== null) {
    response.errors = errors;
  }
  
  return response;
}

/**
 * 배열을 페이지로 나누기
 */
function paginate(array, page, pageSize) {
  const startIndex = (page - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  
  return {
    data: array.slice(startIndex, endIndex),
    page: page,
    pageSize: pageSize,
    total: array.length,
    totalPages: Math.ceil(array.length / pageSize)
  };
}

/**
 * 문자열을 안전한 파일명으로 변환
 */
function sanitizeFileName(fileName) {
  if (!fileName) return 'file';
  
  // 특수문자 제거 및 공백을 언더스코어로 변경
  return fileName
    .replace(/[^a-zA-Z0-9가-힣._-]/g, '_')
    .replace(/\s+/g, '_')
    .substring(0, 100); // 길이 제한
}

/**
 * 업체별 시트 생성
 * @param {string} companyName - 업체명
 * @param {string} companyCode - 업체코드 (필수)
 */
function createCompanySheets(companyName, companyCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 업체코드가 없으면 업체명으로 조회
    if (!companyCode) {
      companyCode = findCompanyCodeByName(companyName);
    }

    // 업체코드가 없으면 에러
    if (!companyCode) {
      throw new Error('업체코드를 찾을 수 없습니다.');
    }

    // 마지막 시트 위치 가져오기 (맨 오른쪽에 생성하기 위함)
    const allSheets = ss.getSheets();
    const lastPosition = allSheets.length;

    // 1. Data 시트 생성 (성적서 업로드용)
    const dataSheetName = `Data_${companyCode}`;
    let dataSheet = ss.getSheetByName(dataSheetName);
    if (!dataSheet) {
      dataSheet = ss.insertSheet(dataSheetName, lastPosition);
      dataSheet.getRange('A1:L1').setValues([[
        '업체CODE', 'ID', '업체명', '입고날짜', '입고시간', 'TM-NO', '제품명', '수량', 'PDF_URL', '등록일시', '등록자', '수정일시'
      ]]);
      dataSheet.getRange('A1:L1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');

      // TM-NO 열(F열, 6번째 컬럼)을 텍스트 형식으로 미리 설정
      dataSheet.getRange('F:F').setNumberFormat('@STRING@');

      Logger.log(`시트 생성 완료: ${dataSheetName}`);
    }

    // 2. List 시트 생성 (ItemList → List)
    const listSheetName = `List_${companyCode}`;
    let listSheet = ss.getSheetByName(listSheetName);
    if (!listSheet) {
      listSheet = ss.insertSheet(listSheetName, lastPosition + 1);
      listSheet.getRange('A1:F1').setValues([[
        '업체CODE', 'TM-NO', '제품명', '업체명', '검사형태', '검사기준서'
      ]]);
      listSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#6aa84f').setFontColor('#ffffff');

      // TM-NO 열(B열, 2번째 컬럼)을 텍스트 형식으로 미리 설정
      listSheet.getRange('B:B').setNumberFormat('@STRING@');
      // 검사기준서 URL 열(F열, 6번째 컬럼)을 텍스트 형식으로 미리 설정
      listSheet.getRange('F:F').setNumberFormat('@STRING@');

      Logger.log(`시트 생성 완료: ${listSheetName}`);
    }

    // 3. Spec 시트 생성 (InspectionSpec → Spec)
    const specSheetName = `Spec_${companyCode}`;
    let specSheet = ss.getSheetByName(specSheetName);
    if (!specSheet) {
      specSheet = ss.insertSheet(specSheetName, lastPosition + 2);
      specSheet.getRange('A1:J1').setValues([[
        '업체CODE', 'TM-NO', '제품명', '업체명', '검사항목', '검사유형', '측정방법', '규격하한', '규격상한', '시료수'
      ]]);
      specSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#e69138').setFontColor('#ffffff');

      // TM-NO 열(B열, 2번째 컬럼)을 텍스트 형식으로 미리 설정
      specSheet.getRange('B:B').setNumberFormat('@STRING@');

      Logger.log(`시트 생성 완료: ${specSheetName}`);
    }

    // 4. Result 시트 생성 (수입검사결과)
    const resultSheetName = `Result_${companyCode}`;
    let resultSheet = ss.getSheetByName(resultSheetName);
    if (!resultSheet) {
      resultSheet = ss.insertSheet(resultSheetName, lastPosition + 3);
      const headers = ['업체CODE', '입고ID', '날짜', '업체명', 'TM-NO', '품명', '검사항목', '검사유형', '검사방법', '규격하한', '규격상한'];
      // 최대 10개 시료까지 지원
      for (let i = 1; i <= 10; i++) {
        headers.push('시료' + i);
      }
      headers.push('합부결과', '등록일시', '등록자');
      resultSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      resultSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');

      // TM-NO 열(E열, 5번째 컬럼)을 텍스트 형식으로 미리 설정
      resultSheet.getRange('E:E').setNumberFormat('@STRING@');

      Logger.log(`시트 생성 완료: ${resultSheetName}`);
    }

    return {
      success: true,
      message: `업체 "${companyName}" (${companyCode}) 시트 생성 완료`,
      sheets: {
        data: dataSheetName,
        list: listSheetName,
        spec: specSheetName,
        result: resultSheetName
      }
    };

  } catch (error) {
    logError('createCompanySheets', error);
    return {
      success: false,
      message: '시트 생성 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 시트 이름 정리 (Google Sheets 제한사항 고려)
 * @param {string} name - 원본 이름
 * @returns {string} 정리된 이름
 */
function sanitizeSheetName(name) {
  if (!name) return 'Company';

  // 특수문자 제거 (Google Sheets에서 금지된 문자: : / \ ? * [ ])
  let sanitized = name
    .replace(/[:\\/\?\*\[\]]/g, '_')
    .trim();

  // 최대 길이 30자로 제한 (Google Sheets 시트명 제한)
  if (sanitized.length > 30) {
    sanitized = sanitized.substring(0, 30);
  }

  return sanitized;
}

/**
 * 다음 업체 코드 생성
 * @returns {string} 새로운 업체 코드 (예: C01, C02, ...)
 */
function generateNextCompanyCode() {
  try {
    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return 'C01'; // Users 시트가 없으면 첫 번째 코드 반환
    }

    const data = userSheet.getDataRange().getDisplayValues();

    // 헤더만 있는 경우
    if (data.length <= 1) {
      return 'C01';
    }

    let maxNumber = 0;

    // 모든 업체 코드에서 최대 번호 찾기
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const code = String(row[0] || '').trim(); // A컬럼 = 업체CODE

      // C01, C02 형식에서 숫자 부분 추출
      const match = code.match(/^C(\d+)$/);
      if (match) {
        const number = parseInt(match[1], 10);
        if (number > maxNumber) {
          maxNumber = number;
        }
      }
    }

    // 다음 번호 생성 (2자리 0 패딩)
    const nextNumber = maxNumber + 1;
    const nextCode = 'C' + String(nextNumber).padStart(2, '0');

    Logger.log(`다음 업체 코드 생성: ${nextCode}`);
    return nextCode;

  } catch (error) {
    logError('generateNextCompanyCode', error);
    return 'C01'; // 에러 시 기본값
  }
}

/**
 * 업체명으로 업체 코드 찾기
 * @param {string} companyName - 업체명
 * @returns {string|null} 업체 코드 또는 null
 */
function findCompanyCodeByName(companyName) {
  try {
    if (!companyName) return null;

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) return null;

    const data = userSheet.getDataRange().getDisplayValues();

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const existingCompanyName = String(row[1] || '').trim(); // B컬럼 = 업체명

      if (existingCompanyName === companyName.trim()) {
        return String(row[0] || '').trim(); // A컬럼 = 업체CODE 반환
      }
    }

    return null; // 찾지 못함

  } catch (error) {
    logError('findCompanyCodeByName', error);
    return null;
  }
}

/**
 * 업체별 Data 시트 이름 생성 (성적서 업로드용)
 * @param {string} companyName - 업체명
 * @returns {string} 시트 이름
 */
function getDataSheetName(companyName) {
  const companyCode = findCompanyCodeByName(companyName);
  if (!companyCode) {
    Logger.log(`경고: ${companyName}의 업체코드를 찾을 수 없습니다.`);
    return `Data_${sanitizeSheetName(companyName)}`;
  }
  return `Data_${companyCode}`;
}

/**
 * 업체별 List 시트 이름 생성 (구 ItemList)
 * @param {string} companyName - 업체명
 * @returns {string} 시트 이름
 */
function getItemListSheetName(companyName) {
  const companyCode = findCompanyCodeByName(companyName);
  if (!companyCode) {
    Logger.log(`경고: ${companyName}의 업체코드를 찾을 수 없습니다.`);
    return `List_${sanitizeSheetName(companyName)}`;
  }
  return `List_${companyCode}`;
}

/**
 * 업체별 Spec 시트 이름 생성 (구 InspectionSpec)
 * @param {string} companyName - 업체명
 * @returns {string} 시트 이름
 */
function getInspectionSpecSheetName(companyName) {
  const companyCode = findCompanyCodeByName(companyName);
  if (!companyCode) {
    Logger.log(`경고: ${companyName}의 업체코드를 찾을 수 없습니다.`);
    return `Spec_${sanitizeSheetName(companyName)}`;
  }
  return `Spec_${companyCode}`;
}

/**
 * 업체별 Result 시트 이름 생성 (수입검사결과)
 * @param {string} companyName - 업체명
 * @returns {string} 시트 이름
 */
function getResultSheetName(companyName) {
  const companyCode = findCompanyCodeByName(companyName);
  if (!companyCode) {
    Logger.log(`경고: ${companyName}의 업체코드를 찾을 수 없습니다.`);
    return `Result_${sanitizeSheetName(companyName)}`;
  }
  return `Result_${companyCode}`;
}

/**
 * 모든 업체명 목록 가져오기 (Users 시트에서)
 * @returns {Array<string>} 업체명 목록
 */
function getAllCompanyNames() {
  try {
    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return [];
    }

    const data = userSheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
      return [];
    }

    const companySet = new Set();

    // 헤더 제외하고 업체명 수집
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const companyName = String(row[1] || '').trim(); // B컬럼 = 업체명

      if (companyName) {
        companySet.add(companyName);
      }
    }

    return Array.from(companySet).sort();

  } catch (error) {
    logError('getAllCompanyNames', error);
    return [];
  }
}

/**
 * 업체 시트명 변경 (업체명이 변경되었을 때)
 * 주의: 현재 시스템은 업체코드 기반 시트명을 사용하므로 (Data_C01, List_C01 등)
 * 업체명이 변경되어도 시트명은 변경하지 않음
 *
 * @param {string} oldCompanyName - 기존 업체명
 * @param {string} newCompanyName - 새 업체명
 * @returns {Object} {success: boolean, message: string}
 */
function renameCompanySheets(oldCompanyName, newCompanyName) {
  try {
    // 현재 시스템은 업체코드 기반 시트명을 사용하므로
    // 업체명이 변경되어도 시트명은 변경하지 않음
    Logger.log(`업체명 변경: ${oldCompanyName} → ${newCompanyName} (시트명은 업체코드 기반이므로 변경 불필요)`);

    return {
      success: true,
      message: '시트명은 업체코드 기반이므로 변경이 불필요합니다.'
    };

  } catch (error) {
    logError('renameCompanySheets', error);
    return {
      success: false,
      message: '시트명 변경 중 오류가 발생했습니다.'
    };
  }
}

/**
 * Result 시트들의 헤더 마이그레이션 (C열부터 헤더를 왼쪽으로 한 칸씩 이동)
 * 기존 데이터는 그대로 유지하고 헤더만 수정
 * @returns {Object} {success: boolean, message: string, details: Array}
 */
function migrateResultSheetHeaders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    const results = [];
    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;

    Logger.log('=== Result 시트 헤더 마이그레이션 시작 ===');

    // 모든 시트 중에서 Result로 시작하는 시트 찾기
    for (let i = 0; i < allSheets.length; i++) {
      const sheet = allSheets[i];
      const sheetName = sheet.getName();

      // Result로 시작하는 시트만 처리
      if (sheetName.startsWith('Result_')) {
        try {
          Logger.log(`처리 중: ${sheetName}`);

          // 헤더 행(1행) 읽기
          const lastCol = sheet.getLastColumn();
          const headerRange = sheet.getRange(1, 1, 1, lastCol);
          const headers = headerRange.getValues()[0];

          Logger.log(`  현재 헤더 (처음 10개): ${headers.slice(0, 10).join(', ')}`);

          // C열(인덱스 2)이 'ID'인지 확인하여 마이그레이션이 필요한지 판단
          if (headers[2] === 'ID') {
            Logger.log(`  마이그레이션 필요: C열이 'ID'임`);

            // C열부터 끝까지 헤더를 한 칸씩 왼쪽으로 이동
            // C열 = D열 값, D열 = E열 값, ...
            for (let col = 3; col <= lastCol; col++) {
              if (col < lastCol) {
                // 다음 열의 헤더를 현재 열로 복사
                const nextHeader = headers[col]; // col은 1-based, headers는 0-based이므로 col이 다음 열
                sheet.getRange(1, col).setValue(nextHeader);
                Logger.log(`    ${col}열: "${headers[col-1]}" → "${nextHeader}"`);
              } else {
                // 마지막 열은 빈 값으로 설정하거나 삭제
                sheet.getRange(1, col).setValue('');
                Logger.log(`    ${col}열: "${headers[col-1]}" → (빈값)`);
              }
            }

            successCount++;
            results.push({
              sheet: sheetName,
              status: 'success',
              message: 'C열부터 헤더를 왼쪽으로 한 칸씩 이동 완료'
            });
            Logger.log(`  ✓ 수정 완료`);

          } else if (headers[2] === '날짜') {
            // 이미 수정되어 있음 (C열이 '날짜'면 정상)
            skipCount++;
            results.push({
              sheet: sheetName,
              status: 'skipped',
              message: '이미 수정되어 있음 (C열=날짜)'
            });
            Logger.log(`  - 이미 수정되어 있음`);
          } else {
            // 예상과 다른 헤더
            skipCount++;
            results.push({
              sheet: sheetName,
              status: 'skipped',
              message: `예상과 다른 C열 헤더: ${headers[2]}`
            });
            Logger.log(`  ? 예상과 다른 헤더: C열="${headers[2]}"`);
          }

        } catch (sheetError) {
          errorCount++;
          results.push({
            sheet: sheetName,
            status: 'error',
            message: sheetError.toString()
          });
          Logger.log(`  ✗ 에러: ${sheetError.message}`);
        }
      }
    }

    const summary = `마이그레이션 완료 - 성공: ${successCount}, 스킵: ${skipCount}, 에러: ${errorCount}`;
    Logger.log('=== ' + summary + ' ===');

    return {
      success: true,
      message: summary,
      details: results,
      stats: {
        success: successCount,
        skipped: skipCount,
        error: errorCount,
        total: results.length
      }
    };

  } catch (error) {
    logError('migrateResultSheetHeaders', error);
    return {
      success: false,
      message: '마이그레이션 중 오류가 발생했습니다: ' + error.message,
      details: []
    };
  }
}
/**
 * Spec 및 Result 시트에 검사유형 열 추가 마이그레이션
 * F열(검사항목 다음)에 검사유형 열을 삽입하고 기본값 '정량' 설정
 * @returns {Object} {success, message, details}
 */
function migrateAddInspectionTypeColumn() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    const results = [];
    let specSuccess = 0;
    let specSkip = 0;
    let resultSuccess = 0;
    let resultSkip = 0;
    let errorCount = 0;

    Logger.log('=== 검사유형 열 추가 마이그레이션 시작 ===');

    for (let i = 0; i < allSheets.length; i++) {
      const sheet = allSheets[i];
      const sheetName = sheet.getName();

      try {
        // Spec 시트 처리
        if (sheetName.startsWith('Spec_')) {
          Logger.log(`처리 중: ${sheetName}`);

          const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
          const headers = headerRange.getValues()[0];

          // F열(인덱스 5)이 '측정방법'이면 마이그레이션 필요
          if (headers[5] === '측정방법') {
            Logger.log(`  마이그레이션 필요: F열이 '측정방법'임`);

            // F열에 새 열 삽입
            sheet.insertColumnAfter(5);

            // F열 헤더를 '검사유형'으로 설정
            sheet.getRange(1, 6).setValue('검사유형');
            sheet.getRange(1, 6).setFontWeight('bold').setBackground('#e69138').setFontColor('#ffffff');

            // 기존 데이터가 있으면 모든 행에 '정량' 설정
            const lastRow = sheet.getLastRow();
            if (lastRow > 1) {
              sheet.getRange(2, 6, lastRow - 1, 1).setValue('정량');
            }

            specSuccess++;
            results.push({
              sheet: sheetName,
              type: 'Spec',
              status: 'success',
              message: 'F열에 검사유형 추가 완료'
            });
            Logger.log(`  ✓ 수정 완료`);

          } else if (headers[5] === '검사유형') {
            specSkip++;
            results.push({
              sheet: sheetName,
              type: 'Spec',
              status: 'skipped',
              message: '이미 검사유형 열이 있음'
            });
            Logger.log(`  - 이미 수정되어 있음`);

          } else {
            specSkip++;
            results.push({
              sheet: sheetName,
              type: 'Spec',
              status: 'skipped',
              message: `예상과 다른 F열 헤더: ${headers[5]}`
            });
            Logger.log(`  ? 예상과 다른 헤더: F열="${headers[5]}"`);
          }
        }

        // Result 시트 처리
        else if (sheetName.startsWith('Result_')) {
          Logger.log(`처리 중: ${sheetName}`);

          const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
          const headers = headerRange.getValues()[0];

          // H열(인덱스 7)이 '검사방법'이면 마이그레이션 필요
          if (headers[7] === '검사방법') {
            Logger.log(`  마이그레이션 필요: H열이 '검사방법'임`);

            // H열에 새 열 삽입
            sheet.insertColumnAfter(7);

            // H열 헤더를 '검사유형'으로 설정
            sheet.getRange(1, 8).setValue('검사유형');
            sheet.getRange(1, 8).setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');

            // 기존 데이터가 있으면 모든 행에 '정량' 설정
            const lastRow = sheet.getLastRow();
            if (lastRow > 1) {
              sheet.getRange(2, 8, lastRow - 1, 1).setValue('정량');
            }

            resultSuccess++;
            results.push({
              sheet: sheetName,
              type: 'Result',
              status: 'success',
              message: 'H열에 검사유형 추가 완료'
            });
            Logger.log(`  ✓ 수정 완료`);

          } else if (headers[7] === '검사유형') {
            resultSkip++;
            results.push({
              sheet: sheetName,
              type: 'Result',
              status: 'skipped',
              message: '이미 검사유형 열이 있음'
            });
            Logger.log(`  - 이미 수정되어 있음`);

          } else {
            resultSkip++;
            results.push({
              sheet: sheetName,
              type: 'Result',
              status: 'skipped',
              message: `예상과 다른 H열 헤더: ${headers[7]}`
            });
            Logger.log(`  ? 예상과 다른 헤더: H열="${headers[7]}"`);
          }
        }

      } catch (sheetError) {
        errorCount++;
        results.push({
          sheet: sheetName,
          type: sheetName.startsWith('Spec_') ? 'Spec' : 'Result',
          status: 'error',
          message: sheetError.toString()
        });
        Logger.log(`  ✗ 에러: ${sheetError.message}`);
      }
    }

    const summary = `마이그레이션 완료 - Spec(성공:${specSuccess}, 스킵:${specSkip}), Result(성공:${resultSuccess}, 스킵:${resultSkip}), 에러:${errorCount}`;
    Logger.log('=== ' + summary + ' ===');

    return {
      success: true,
      message: summary,
      details: results,
      stats: {
        spec: { success: specSuccess, skipped: specSkip },
        result: { success: resultSuccess, skipped: resultSkip },
        error: errorCount,
        total: results.length
      }
    };

  } catch (error) {
    logError('migrateAddInspectionTypeColumn', error);
    return {
      success: false,
      message: '마이그레이션 중 오류가 발생했습니다: ' + error.message,
      details: []
    };
  }
}
