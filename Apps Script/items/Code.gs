/**
 * ============================================================================
 * 동호회 협찬물품 관리 웹앱 - 백엔드 (Code.gs)
 * ============================================================================
 */

// [필수 설정] 2단계에서 복사한 구글 드라이브 폴더 ID를 아래 따옴표 안에 넣으세요.
const FOLDER_ID = '1p3fQDreMkfl7fnIofZx2YelOqH-9wp2P';
const SHEET_NAME = 'Items'; // 시트 탭 이름

/**
 * 웹 앱 접속 시 HTML 화면(Index.html)을 렌더링하는 기본 함수
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('동호회 협찬물품 관리 시스템')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    // 모바일 화면에서 반응형(Viewport)이 정상 작동하도록 메타 태그 추가
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * 1. [조회] 구글 시트의 모든 물품 데이터를 가져오는 함수
 */
function getItems() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`'${SHEET_NAME}' 시트를 찾을 수 없습니다. 시트 이름을 확인해주세요.`);
    }

    const lastRow = sheet.getLastRow();
    // 헤더(1행)만 있고 데이터가 없는 경우 빈 배열 반환
    if (lastRow < 2) {
      return { success: true, data: [] };
    }

    // 2행부터 마지막 행까지 9개 컬럼(A~I) 데이터 가져오기
    const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    
    const items = values.map(row => {
      // 등록일이 Date 객체로 저장된 경우 YYYY-MM-DD 문자열로 변환
      let formattedDate = row[5];
      if (row[5] instanceof Date) {
        formattedDate = Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      return {
        id: String(row[0]),
        name: String(row[1]),
        category: String(row[2]),
        quantity: Number(row[3]),
        sponsor: String(row[4]),
        regDate: String(formattedDate),
        imageUrl: String(row[6]),
        status: String(row[7]),
        memo: row[8] !== undefined && row[8] !== null ? String(row[8]) : ''
      };
    });

    // 최신 등록 물품이 위로 오도록 역순 정렬 후 반환
    return { success: true, data: items.reverse() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 2. [추가] 새로운 협찬 물품 등록 및 이미지 구글 드라이브 업로드
 */
function addItem(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    let imageUrl = '';
    // 첨부된 이미지 데이터가 있는 경우 구글 드라이브에 업로드
    if (formData.imageBase64 && formData.imageName) {
      imageUrl = uploadImageToDrive(formData.imageBase64, formData.imageName, formData.imageType);
    }

    // 고유 ID 생성 (타임스탬프 기반)
    const id = 'ITEM-' + new Date().getTime();
    // 오늘 날짜 (YYYY-MM-DD)
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // 시트 맨 아래에 새 행 추가 (A~H 컬럼 순서 일치)
    sheet.appendRow([
      id,
      formData.name,
      formData.category,
      Number(formData.quantity),
      formData.sponsor,
      today,
      imageUrl,
      formData.status,
      formData.memo || '' // I열: 메모 저장
    ]);

    return { success: true, message: '물품이 성공적으로 등록되었습니다.' };
  } catch (error) {
    return { success: false, message: '등록 중 오류 발생: ' + error.toString() };
  }
}

/**
 * 3. [수정] 기존 물품 정보 업데이트 (새 이미지 업로드 시 교체 포함)
 */
function updateItem(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) throw new Error('수정할 데이터가 없습니다.');

    // ID 컬럼(A열)에서 일치하는 행 찾기
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let targetRowIndex = -1;

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(formData.id)) {
        targetRowIndex = i + 2; // 배열 인덱스는 0부터 시작, 시트는 2행부터 시작하므로 +2
        break;
      }
    }

    if (targetRowIndex === -1) {
      throw new Error('해당 ID의 물품을 찾을 수 없습니다.');
    }

    // 기존 이미지 URL 유지, 새 이미지가 첨부되었다면 새로 업로드 후 URL 교체
    let imageUrl = formData.existingImageUrl || '';
    if (formData.imageBase64 && formData.imageName) {
      imageUrl = uploadImageToDrive(formData.imageBase64, formData.imageName, formData.imageType);
    }

    // 시트 데이터 업데이트 (B열:물품명, C열:카테고리, D열:수량, E열:협찬자, G열:이미지URL, H열:상태, I열:메모)
    sheet.getRange(targetRowIndex, 2).setValue(formData.name);
    sheet.getRange(targetRowIndex, 3).setValue(formData.category);
    sheet.getRange(targetRowIndex, 4).setValue(Number(formData.quantity));
    sheet.getRange(targetRowIndex, 5).setValue(formData.sponsor);
    sheet.getRange(targetRowIndex, 7).setValue(imageUrl);
    sheet.getRange(targetRowIndex, 8).setValue(formData.status);
    sheet.getRange(targetRowIndex, 9).setValue(formData.memo || ''); // I열: 메모 수정

    return { success: true, message: '물품 정보가 수정되었습니다.' };
  } catch (error) {
    return { success: false, message: '수정 중 오류 발생: ' + error.toString() };
  }
}

/**
 * 4. [삭제] ID로 물품 행 삭제
 */
function deleteItem(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) throw new Error('삭제할 데이터가 없습니다.');

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        return { success: true, message: '물품이 삭제되었습니다.' };
      }
    }

    throw new Error('삭제할 물품을 찾을 수 없습니다.');
  } catch (error) {
    return { success: false, message: '삭제 중 오류 발생: ' + error.toString() };
  }
}

/**
 * [내부 유틸리티] Base64 이미지를 구글 드라이브 폴더에 저장하고 웹 출력용 URL 반환
 */
function uploadImageToDrive(base64Data, fileName, mimeType) {
  if (!FOLDER_ID || FOLDER_ID === '여기에_구글_드라이브_폴더_ID를_입력하세요') {
    throw new Error('Code.gs 파일 상단에 구글 드라이브 FOLDER_ID를 설정해주세요.');
  }

  // Base64 문자열 디코딩 후 Blob(파일 객체) 생성
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);

  // 지정한 폴더에 파일 생성
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const file = folder.createFile(blob);

  // 외부 웹 브라우저에서 이미지가 바로 보이도록 링크 공유 권한 설정
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Google Drive 이미지 직접 렌더링을 위한 최신 CDN URL 형식 반환
  const fileId = file.getId();
  return `https://lh3.googleusercontent.com/d/${fileId}`;
}

/**
 * [선택] 외부 API(POST 요청)로도 데이터를 받을 수 있도록 구현한 doPost
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    if (data.action === 'add') result = addItem(data);
    else if (data.action === 'update') result = updateItem(data);
    else if (data.action === 'delete') result = deleteItem(data.id);
    else result = { success: false, message: '지원하지 않는 요청입니다.' };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
