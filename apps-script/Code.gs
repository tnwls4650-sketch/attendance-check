/*
  출석체크 체크인 저장소 (Google Apps Script)

  동작
  - 참가자가 이름을 확인하면, 구글 시트의 명단 탭에서 그 사람 행을 찾아 "참석여부" 칸에 O를 표시합니다.
  - 새 행을 쌓지 않습니다. 명단 탭의 "성명"과 "참석여부" 열을 이름으로 찾습니다.
  - 관리자 화면은 참석여부가 O인 이름 목록을 받아 보여 줍니다.

  사용법
  1. 명단이 들어 있는 구글 스프레드시트를 엽니다. (탭 이름은 아래 ROSTER_SHEET와 같아야 합니다)
  2. 메뉴 [확장 프로그램] → [Apps Script]를 열고, 기본 코드를 지운 뒤 이 파일 내용을 붙여넣습니다.
  3. [배포] → [새 배포] → 유형 "웹 앱"
     · 실행 사용자: 나
     · 액세스 권한: 모든 사용자
  4. 웹 앱 URL을 data/roster.js의 checkinUrl에 넣습니다.
  5. 코드를 고치면 [배포] → [배포 관리] → 연필 → [새 버전] → 배포 (URL은 그대로 유지됩니다)
*/

const SHEET_ID = '1VfigZ4E6KF06ybVEyIvihmz6gcyMzyZiwTRVzVIVhJs'; // 출석체크_명단(11월) 시트
const ROSTER_SHEET = '명단(11월)';   // 명단 탭 이름
const NAME_HEADER = '성명';           // 이름이 있는 열의 제목
const CHECK_HEADER = '참석여부';      // O를 표시할 열의 제목

// 참가자가 이름을 확인했을 때: 참석여부에 O 표시
function doPost(e){
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const name = normalize_(data.name);
    if(!name) return out_({ ok: false, error: 'name required' });

    const sheet = getRosterSheet_();
    const cols = findColumns_(sheet);
    const rowIndex = findRow_(sheet, cols.name, name);
    if(rowIndex < 0) return out_({ ok: false, error: 'not in roster' });

    const cell = sheet.getRange(rowIndex, cols.check);
    const already = String(cell.getValue()).trim().toUpperCase() === 'O';
    if(!already) cell.setValue('O');
    return out_({ ok: true, name: name, alreadyChecked: already });
  } catch(err){
    return out_({ ok: false, error: String(err) });
  }
}

// 관리자 화면: 참석여부가 O인 이름 목록
function doGet(){
  const sheet = getRosterSheet_();
  const cols = findColumns_(sheet);
  const values = sheet.getDataRange().getValues();
  const checked = [];
  let total = 0;
  for(let r = 1; r < values.length; r++){
    const name = normalize_(values[r][cols.name - 1]);
    if(!name) continue;
    total++;
    if(String(values[r][cols.check - 1]).trim().toUpperCase() === 'O') checked.push(name);
  }
  return out_({ ok: true, total: total, checked: checked });
}

function getRosterSheet_(){
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ROSTER_SHEET);
  if(!sheet) throw new Error('탭을 찾을 수 없습니다: ' + ROSTER_SHEET);
  return sheet;
}

// 1행의 제목에서 "성명", "참석여부" 열 번호(1부터) 찾기
function findColumns_(sheet){
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const name = headers.indexOf(NAME_HEADER) + 1;
  const check = headers.indexOf(CHECK_HEADER) + 1;
  if(!name || !check) throw new Error('1행에 "' + NAME_HEADER + '" 또는 "' + CHECK_HEADER + '" 제목이 없습니다.');
  return { name: name, check: check };
}

// 이름이 있는 행 번호(시트 기준, 1부터) 찾기. 없으면 -1
function findRow_(sheet, nameCol, name){
  const values = sheet.getRange(2, nameCol, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  for(let i = 0; i < values.length; i++){
    if(normalize_(values[i][0]) === name) return i + 2;
  }
  return -1;
}

function normalize_(v){
  return String(v == null ? '' : v).replace(/\s+/g, '');
}

function out_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
