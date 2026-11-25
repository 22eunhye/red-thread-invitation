/**
 * 붉은 실 초대장 — 닉네임 등록용 Google Apps Script
 *
 * 1) 구글 스프레드시트를 새로 만들고  확장 프로그램 → Apps Script 를 엽니다.
 * 2) 이 코드를 통째로 붙여넣고 저장합니다.
 * 3) 배포 → 새 배포 → 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4) 나오는 웹 앱 URL을 index.html 의 CONFIG.scriptUrl 에 넣습니다.
 *
 * 등록 내용은 이 스프레드시트의 '닉네임' 시트에 쌓이고,
 * 초대장 화면에는 닉네임만 보입니다(이름은 시트에서 주최자만 볼 수 있어요).
 */

const SHEET_NAME = '닉네임';
const NOTIFY_EMAIL = '';   // 등록될 때마다 메일로 받고 싶으면 이메일 주소를 적으세요. 비우면 알림 없음.

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['제출 시각', '이름', '닉네임']);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 초대장이 닉네임 목록을 불러갈 때
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  return json_({ nicknames: rows.map(r => String(r[2])).filter(Boolean) });
}

// 닉네임을 등록할 때 (같은 이름으로 다시 내면 닉네임이 바뀝니다)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    const name = String(data.name || '').trim().slice(0, 20);
    const nickname = String(data.nickname || '').trim().slice(0, 12);
    if (!name || !nickname) return json_({ ok: false, error: '이름과 닉네임을 모두 적어주세요.' });

    const sh = sheet_();
    const rows = sh.getDataRange().getValues().slice(1);
    const mine = rows.findIndex(r => String(r[1]) === name);
    const taken = rows.some((r, i) => i !== mine && String(r[2]).toLowerCase() === nickname.toLowerCase());
    if (taken) return json_({ ok: false, error: '이미 누가 쓰고 있는 닉네임이에요.' });

    if (mine >= 0) sh.getRange(mine + 2, 1, 1, 3).setValues([[new Date(), name, nickname]]);
    else sh.appendRow([new Date(), name, nickname]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, '[붉은 실] 닉네임 등록: ' + nickname, name + ' → ' + nickname);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: '등록하지 못했어요. 다시 시도해 주세요.' });
  } finally {
    lock.releaseLock();
  }
}
