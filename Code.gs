/**
 * 별의 밝기 탐구실 — 모둠 자료 저장용 Apps Script
 * 구글 스프레드시트에서 [확장 프로그램 > Apps Script]를 열고
 * 이 파일 내용을 통째로 붙여 넣은 뒤 배포하면 됩니다.
 * 자세한 순서는 설정안내.md를 참고하세요.
 */

var SHEET_NAME = '측정기록';
var HEADERS = ['키','반','모둠','모둠원','거리설정','측정자료(JSON)','답변(JSON)','지수n','제출시각'];

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return out({ ok: false, error: '다른 모둠이 저장 중입니다. 잠시 후 다시 눌러 주세요.' });
  }
  try {
    var data = JSON.parse(e.postData.contents);
    if (!data.key) throw new Error('키가 없습니다');

    var sh = getSheet();
    var row = [
      data.key,
      data.cls,
      data.grp,
      data.members || '',
      data.preset || '',
      JSON.stringify(data.rows || []),
      JSON.stringify(data.answers || []),
      data.expo === '' ? '' : data.expo,
      new Date()
    ];

    var vals = sh.getDataRange().getValues();
    var found = 0;
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(data.key)) { found = i + 1; break; }
    }
    if (found > 0) sh.getRange(found, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var sh = getSheet();
    var vals = sh.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < vals.length; i++) {
      var v = vals[i];
      if (!v[0]) continue;
      list.push({
        key: String(v[0]),
        cls: v[1],
        grp: v[2],
        members: String(v[3] || ''),
        preset: String(v[4] || 'A'),
        rows: safeParse(v[5]),
        answers: safeParse(v[6]),
        expo: v[7],
        updated: toIso(v[8])
      });
    }
    return out({ ok: true, list: list });
  } catch (err) {
    return out({ ok: false, error: String(err.message || err), list: [] });
  }
}

function safeParse(s) {
  try { return JSON.parse(s); } catch (err) { return []; }
}

function toIso(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm");
  return String(v || '');
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 학기 끝나고 자료를 비울 때 이 함수를 한 번 실행하세요. */
function 자료전체지우기() {
  var sh = getSheet();
  var last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
}
