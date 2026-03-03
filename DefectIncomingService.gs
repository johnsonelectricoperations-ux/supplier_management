/**
 * 소재불량반입 서비스
 * DefectIncomingService.gs
 */

const DEFECT_TYPES_SHEET_NAME = 'DefectTypes';
const SINTER_DEFECT_SHEET_NAME = 'sinterdefect';

/**
 * 업체의 불량유형 그룹 가져오기
 * @param {string} companyName - 업체명
 * @returns {string} 불량유형 그룹 (A, B 등)
 */
function getDefectTypeGroup(companyName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);

    if (!userSheet) {
      throw new Error('Users 시트를 찾을 수 없습니다.');
    }

    const data = userSheet.getDataRange().getValues();

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCompanyName = row[1]; // B열: 업체명
      const defectGroup = row[8]; // I열: 불량유형그룹

      if (rowCompanyName === companyName) {
        // 빈 문자열, null, undefined, 공백만 있는 경우 모두 'B'로 처리
        if (defectGroup && defectGroup.toString().trim()) {
          return defectGroup.toString().trim();
        } else {
          return 'B'; // 기본값 B
        }
      }
    }

    return 'B'; // 기본값
  } catch (error) {
    Logger.log('불량유형 그룹 조회 오류: ' + error.toString());
    return 'B';
  }
}

/**
 * 업체의 불량유형 목록 가져오기
 * @param {string} token - 세션 토큰
 * @returns {object} {success, defectTypes: [{code, name, order}]}
 */
function getDefectTypesForCompany(token) {
  try {
    // 세션 확인
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '세션이 만료되었습니다.' };
    }

    const companyName = session.companyName;
    const defectGroup = getDefectTypeGroup(companyName);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const defectTypesSheet = ss.getSheetByName(DEFECT_TYPES_SHEET_NAME);

    if (!defectTypesSheet) {
      return { success: false, message: 'DefectTypes 시트를 찾을 수 없습니다.' };
    }

    const data = defectTypesSheet.getDataRange().getValues();
    const defectTypes = [];

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const group = row[0]; // A열: 그룹
      const code = row[1]; // B열: 불량유형코드
      const name = row[2]; // C열: 불량유형명
      const order = row[3]; // D열: 순서
      const active = row[4]; // E열: 활성화여부

      // 해당 그룹이고 활성화된 불량유형만 가져오기
      if (group === defectGroup && active === 'Y') {
        defectTypes.push({
          code: code,
          name: name,
          order: order
        });
      }
    }

    // 순서대로 정렬
    defectTypes.sort((a, b) => a.order - b.order);

    return {
      success: true,
      defectTypes: defectTypes,
      defectGroup: defectGroup
    };

  } catch (error) {
    Logger.log('불량유형 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '불량유형 조회 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 모든 불량유형 코드 가져오기 (시트 헤더용)
 * @returns {Array} 불량유형 코드 배열
 */
function getAllDefectTypeCodes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const defectTypesSheet = ss.getSheetByName(DEFECT_TYPES_SHEET_NAME);

    if (!defectTypesSheet) {
      return [];
    }

    const data = defectTypesSheet.getDataRange().getValues();
    const codes = [];

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const code = row[1]; // B열: 불량유형코드
      const active = row[4]; // E열: 활성화여부

      if (active === 'Y' && !codes.includes(code)) {
        codes.push(code);
      }
    }

    return codes;

  } catch (error) {
    Logger.log('전체 불량유형 코드 조회 오류: ' + error.toString());
    return [];
  }
}

/**
 * 모든 불량유형 가져오기 (출력용 - 코드와 이름 매핑)
 * @returns {object} {success, defectTypes: [{code, name}]}
 */
function getAllDefectTypes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const defectTypesSheet = ss.getSheetByName(DEFECT_TYPES_SHEET_NAME);

    if (!defectTypesSheet) {
      return { success: false, message: 'DefectTypes 시트를 찾을 수 없습니다.' };
    }

    const data = defectTypesSheet.getDataRange().getValues();
    const defectTypes = [];
    const addedCodes = new Set(); // 중복 방지

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const code = row[1]; // B열: 불량유형코드
      const name = row[2]; // C열: 불량유형명
      const active = row[4]; // E열: 활성화여부

      // 활성화되어 있고 아직 추가하지 않은 코드만
      if (active === 'Y' && !addedCodes.has(code)) {
        defectTypes.push({
          code: code,
          name: name
        });
        addedCodes.add(code);
      }
    }

    return {
      success: true,
      defectTypes: defectTypes
    };

  } catch (error) {
    Logger.log('전체 불량유형 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '불량유형 조회 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 업체의 Item 목록 가져오기
 * @param {string} token - 세션 토큰
 * @returns {object} {success, items: [{tmNo, itemName}]}
 */
function getItemListForDefect(token) {
  try {
    // 세션 확인
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '세션이 만료되었습니다.' };
    }

    const companyName = session.companyName;
    const listSheetName = getItemListSheetName(companyName);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listSheet = ss.getSheetByName(listSheetName);

    if (!listSheet) {
      return { success: false, message: 'Item List 시트를 찾을 수 없습니다. (' + listSheetName + ')' };
    }

    const data = listSheet.getDataRange().getValues();
    const items = [];

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tmNo = row[1]; // B열: TM-NO
      const itemName = row[2]; // C열: 제품명

      if (tmNo && itemName) {
        items.push({
          tmNo: tmNo.toString(),
          itemName: itemName
        });
      }
    }

    return {
      success: true,
      items: items
    };

  } catch (error) {
    Logger.log('Item 목록 조회 오류: ' + error.toString());
    return {
      success: false,
      message: 'Item 목록 조회 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 소재불량반입 데이터 저장
 * @param {string} token - 세션 토큰
 * @param {object} data - {incomingDate, rows: [{tmNo, itemName, category, defects: {code: qty}}]}
 * @returns {object} {success, message}
 */
function saveDefectIncoming(token, data) {
  try {
    // 세션 확인
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '세션이 만료되었습니다.' };
    }

    const companyName = session.companyName;
    const userId = session.userId;
    const incomingDate = data.incomingDate;
    const rows = data.rows;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sinterDefectSheet = ss.getSheetByName(SINTER_DEFECT_SHEET_NAME);

    if (!sinterDefectSheet) {
      return { success: false, message: 'sinterdefect 시트를 찾을 수 없습니다.' };
    }

    // D열(TM-NO)을 텍스트 형식으로 미리 설정
    sinterDefectSheet.getRange('D:D').setNumberFormat('@STRING@');

    // 헤더 가져오기 (불량유형 컬럼 매핑용)
    const headers = sinterDefectSheet.getRange(1, 1, 1, sinterDefectSheet.getLastColumn()).getValues()[0];

    // 각 행 저장
    let savedCount = 0;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];

      // 식별ID 생성 (타임스탬프 + 랜덤)
      const id = 'DFCT' + new Date().getTime() + Math.floor(Math.random() * 1000);

      // 불량수량 합계 계산
      let totalQty = 0;
      const defects = row.defects || {};
      for (let code in defects) {
        totalQty += parseInt(defects[code]) || 0;
      }

      // 행 데이터 준비
      const newRow = [];

      // 고정 컬럼 (A~G)
      newRow[0] = id; // A: 식별ID
      newRow[1] = incomingDate; // B: 반입예정일자
      newRow[2] = companyName; // C: 업체명
      newRow[3] = row.tmNo; // D: TM-NO
      newRow[4] = row.itemName; // E: 품명
      newRow[5] = row.category; // F: 구분
      newRow[6] = totalQty; // G: 합계

      // 불량유형별 수량 (H열부터)
      for (let j = 7; j < headers.length; j++) {
        const defectCode = headers[j];
        newRow[j] = defects[defectCode] || 0;
      }

      // 시트에 추가
      const nextRow = sinterDefectSheet.getLastRow() + 1;
      sinterDefectSheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);
      savedCount++;
    }

    return {
      success: true,
      message: savedCount + '건의 소재불량반입 데이터가 저장되었습니다.',
      savedCount: savedCount
    };

  } catch (error) {
    Logger.log('소재불량반입 저장 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 저장 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 업체의 소재불량반입 리스트 조회
 * @param {string} token - 세션 토큰
 * @returns {object} {success, data: [...]}
 */
function getDefectIncomingList(token) {
  try {
    Logger.log('getDefectIncomingList 시작');

    // 세션 확인
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      Logger.log('세션 확인 실패');
      return { success: false, message: '세션이 만료되었습니다.' };
    }

    const companyName = session.companyName;
    Logger.log('업체명: ' + companyName);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sinterDefectSheet = ss.getSheetByName(SINTER_DEFECT_SHEET_NAME);

    if (!sinterDefectSheet) {
      Logger.log('sinterdefect 시트를 찾을 수 없음');
      return { success: false, message: 'sinterdefect 시트를 찾을 수 없습니다.' };
    }

    const lastRow = sinterDefectSheet.getLastRow();
    Logger.log('마지막 행: ' + lastRow);

    // 헤더만 있거나 데이터가 없는 경우
    if (lastRow <= 1) {
      Logger.log('데이터 없음, 빈 리스트 반환');
      const headers = sinterDefectSheet.getRange(1, 1, 1, sinterDefectSheet.getLastColumn()).getValues()[0];
      const safeHeaders = (headers || []).map(h => {
        if (h === null || h === undefined) return '';
        return h.toString();
      });
      return {
        success: true,
        data: [],
        headers: safeHeaders
      };
    }

    const allData = sinterDefectSheet.getDataRange().getValues();
    const headers = allData[0];
    Logger.log('헤더 개수: ' + headers.length);
    Logger.log('헤더: ' + JSON.stringify(headers));

    const result = [];

    // 헤더 제외하고 해당 업체 데이터만 필터링
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const rowCompanyName = row[2]; // C열: 업체명

      if (rowCompanyName === companyName) {
        // 행 데이터를 객체로 변환
        const rowObj = {};
        for (let j = 0; j < headers.length; j++) {
          const headerName = headers[j];

          // 빈 헤더는 건너뛰기
          if (!headerName || headerName.toString().trim() === '') {
            continue;
          }

          // 값을 안전하게 변환 (Date 객체 등을 문자열로)
          let value = row[j];
          if (value instanceof Date) {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          } else if (value !== null && value !== undefined) {
            value = value.toString();
          } else {
            value = '';
          }

          rowObj[headerName.toString()] = value;
        }
        result.push(rowObj);
      }
    }

    Logger.log('필터링된 데이터 개수: ' + result.length);

    // 반환 전 직렬화 테스트
    try {
      const testJson = JSON.stringify(result);
      Logger.log('데이터 직렬화 성공, 크기: ' + testJson.length);
    } catch (e) {
      Logger.log('데이터 직렬화 실패: ' + e.toString());
    }

    // 최신순 정렬 (식별ID 기준)
    if (result.length > 0) {
      result.sort((a, b) => {
        const idA = (a['식별ID'] || '').toString();
        const idB = (b['식별ID'] || '').toString();
        return idB.localeCompare(idA);
      });
    }

    // headers 배열도 안전하게 변환
    const safeHeaders = headers.map(h => {
      if (h === null || h === undefined) return '';
      return h.toString();
    });

    Logger.log('getDefectIncomingList 완료');
    return {
      success: true,
      data: result,
      headers: safeHeaders
    };

  } catch (error) {
    Logger.log('소재불량반입 리스트 조회 오류: ' + error.toString());
    Logger.log('에러 스택: ' + error.stack);
    return {
      success: false,
      message: '리스트 조회 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 특정 ID의 소재불량반입 데이터 상세 조회
 * @param {string} token - 세션 토큰
 * @param {string} id - 식별ID
 * @returns {object} {success, data}
 */
function getDefectIncomingById(token, id) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '세션이 만료되었습니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sinterDefectSheet = ss.getSheetByName(SINTER_DEFECT_SHEET_NAME);

    if (!sinterDefectSheet) {
      return { success: false, message: 'sinterdefect 시트를 찾을 수 없습니다.' };
    }

    const data = sinterDefectSheet.getDataRange().getValues();
    const headers = data[0];

    // ID로 행 찾기
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        const rowData = {};
        for (let j = 0; j < headers.length; j++) {
          let value = data[i][j];
          if (value instanceof Date) {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }
          rowData[headers[j]] = value;
        }
        rowData.rowIndex = i + 1; // 시트 행 번호 (1-based)

        return {
          success: true,
          data: rowData
        };
      }
    }

    return { success: false, message: '데이터를 찾을 수 없습니다.' };

  } catch (error) {
    Logger.log('소재불량반입 상세 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 조회 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 소재불량반입 데이터 수정
 * @param {string} token - 세션 토큰
 * @param {string} id - 식별ID
 * @param {object} data - {incomingDate, tmNo, itemName, category, defects}
 * @returns {object} {success, message}
 */
function updateDefectIncoming(token, id, data) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '세션이 만료되었습니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sinterDefectSheet = ss.getSheetByName(SINTER_DEFECT_SHEET_NAME);

    if (!sinterDefectSheet) {
      return { success: false, message: 'sinterdefect 시트를 찾을 수 없습니다.' };
    }

    const sheetData = sinterDefectSheet.getDataRange().getValues();
    const headers = sheetData[0];

    // ID로 행 찾기
    let targetRow = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0] === id) {
        targetRow = i + 1; // 1-based index
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: '수정할 데이터를 찾을 수 없습니다.' };
    }

    // 불량수량 합계 계산
    let totalQty = 0;
    const defects = data.defects || {};
    for (let code in defects) {
      totalQty += parseInt(defects[code]) || 0;
    }

    // 행 데이터 준비
    const newRow = [];
    newRow[0] = id; // A: 식별ID (유지)
    newRow[1] = data.incomingDate; // B: 반입예정일자
    newRow[2] = session.companyName; // C: 업체명 (유지)
    newRow[3] = data.tmNo; // D: TM-NO
    newRow[4] = data.itemName; // E: 품명
    newRow[5] = data.category; // F: 구분
    newRow[6] = totalQty; // G: 합계

    // 불량유형별 수량 (H열부터)
    for (let j = 7; j < headers.length; j++) {
      const defectCode = headers[j];
      newRow[j] = defects[defectCode] || 0;
    }

    // 행 업데이트
    sinterDefectSheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);

    Logger.log(`소재불량반입 수정 완료: ${id} by ${session.name}`);

    return {
      success: true,
      message: '데이터가 수정되었습니다.'
    };

  } catch (error) {
    Logger.log('소재불량반입 수정 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 수정 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 소재불량반입 데이터 삭제
 * @param {string} token - 세션 토큰
 * @param {string} id - 식별ID
 * @returns {object} {success, message}
 */
function deleteDefectIncoming(token, id) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '세션이 만료되었습니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sinterDefectSheet = ss.getSheetByName(SINTER_DEFECT_SHEET_NAME);

    if (!sinterDefectSheet) {
      return { success: false, message: 'sinterdefect 시트를 찾을 수 없습니다.' };
    }

    const data = sinterDefectSheet.getDataRange().getValues();

    // ID로 행 찾기
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        targetRow = i + 1; // 1-based index
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: '삭제할 데이터를 찾을 수 없습니다.' };
    }

    // 행 삭제
    sinterDefectSheet.deleteRow(targetRow);

    Logger.log(`소재불량반입 삭제 완료: ${id} by ${session.name}`);

    return {
      success: true,
      message: '데이터가 삭제되었습니다.'
    };

  } catch (error) {
    Logger.log('소재불량반입 삭제 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 삭제 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}
