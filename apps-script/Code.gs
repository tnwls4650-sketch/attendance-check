/*
  출석체크 체크인 저장소 (Google Apps Script)

  사용법
  1. 구글 스프레드시트를 새로 만든다.
  2. 메뉴 [확장 프로그램] → [Apps Script]를 열고, 기본 코드를 지운 뒤 이 파일 내용을 붙여넣는다.
  3. [배포] → [새 배포] → 유형 "웹 앱"
     · 실행 사용자: 나
     · 액세스 권한: 모든 사용자
  4. 배포 후 나오는 "웹 앱 URL"을 data/roster.js의 checkinUrl에 넣는다.
  5. 이 코드를 수정하면 [배포] → [배포 관리] → 연필 아이콘 → [새 버전]으로 다시 배포해야 반영된다.

  시트 "출석" 구성: 시간 | 이름 | 조 | 좌석  (이름당 한 줄, 중복은 저장하지 않음)
*/

const SHEET_NAME = '출석';

// 체크인 기록 저장 (이미 기록된 이름이면 건너뜀)
function doPost(e){
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const name = String(data.name || '').trim();
    if(!name) return out_({ ok: false, error: 'name required' });

    const sheet = getSheet_();
    const rows = sheet.getDataRange().getValues();
    const exists = rows.some((r, i) => i > 0 && String(r[1]).trim() === name);
    if(!exists){
      sheet.appendRow([new Date(), name, data.team || '', data.seat || '']);
    }
    return out_({ ok: true, duplicate: exists });
  } catch(err){
    return out_({ ok: false, error: String(err) });
  }
}

// 관리자 화면: 체크인 목록 조회
function doGet(){
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues().slice(1).filter(r => String(r[1]).trim() !== '');
  const checkins = rows.map(r => ({
    time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    name: String(r[1]).trim(),
    team: r[2],
    seat: r[3]
  }));
  return out_({ ok: true, count: checkins.length, checkins: checkins });
}

function getSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['시간', '이름', '조', '좌석']);
  }
  return sheet;
}

function out_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
