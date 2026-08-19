/**
 * 스낵앤가든 크루 보이스 — 백엔드 (Google Apps Script 웹앱)
 * ──────────────────────────────────────────────────────────────
 *  · 접수 글  →  Google Sheet 에 한 행씩 저장
 *  · 사진      →  Google Drive 폴더에 저장하고, 시트에 링크 기록
 *  · 관리자    →  admin.html 에서 doGet(action=list) 으로 목록 조회 / 상태 변경
 *
 *  시트와 폴더는 이 스크립트가 "처음 실행될 때 자동으로 생성"됩니다.
 *  (스크립트를 배포한 구글 계정의 내 드라이브에 생성됨)
 * ──────────────────────────────────────────────────────────────
 */

var CFG = {
  SHEET_NAME : '스낵앤가든_크루보이스_접수',
  FOLDER_NAME: '크루보이스_사진',
  // 관리자 페이지 접근 키. admin.html 잠금화면에 입력하는 값. 원하는 값으로 바꾸세요.
  ADMIN_KEY  : '1234',
  HEADERS    : ['접수시각','접수번호','분류','위치','내용','사진','상태','처리메모']
};

/* ===== 저장(접수) ===== */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || '{}');
    var sheet = getSheet_();
    var folder = getFolder_();

    // 사진 저장
    var links = [];
    (data.photos || []).forEach(function (dataUrl, i) {
      var m = /^data:(image\/\w+);base64,(.+)$/.exec(dataUrl || '');
      if (!m) return;
      var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1],
        (data.ticket || 'photo') + '_' + (i + 1) + '.' + (m[1].split('/')[1] || 'jpg'));
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      links.push(file.getUrl());
    });

    var now = new Date();
    sheet.appendRow([
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      data.ticket || '', data.category || '', data.location || '',
      data.message || '', links.join('\n'), '접수', ''
    ]);

    return json_({ ok: true, ticket: data.ticket, photos: links });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ===== 조회 / 상태변경 (관리자) ===== */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.key !== CFG.ADMIN_KEY) return json_({ ok: false, error: 'unauthorized' });

  var sheet = getSheet_();

  if (p.action === 'status' && p.row) {
    var r = parseInt(p.row, 10);
    if (r >= 2) {
      sheet.getRange(r, 7).setValue(p.status || '접수');      // 상태
      if (p.memo != null) sheet.getRange(r, 8).setValue(p.memo); // 처리메모
    }
    return json_({ ok: true });
  }

  // 기본: 목록 반환(최신순)
  var last = sheet.getLastRow();
  var items = [];
  if (last >= 2) {
    var rows = sheet.getRange(2, 1, last - 1, 8).getValues();
    rows.forEach(function (row, idx) {
      items.push({
        row: idx + 2,
        time: row[0], ticket: row[1], category: row[2], location: row[3],
        message: row[4],
        photos: String(row[5] || '').split('\n').filter(String),
        status: row[6] || '접수', memo: row[7] || ''
      });
    });
    items.reverse();
  }
  return json_({ ok: true, items: items });
}

/* ===== 헬퍼 ===== */
function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SS_ID');
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(CFG.SHEET_NAME);
    props.setProperty('SS_ID', ss.getId());
    var sh = ss.getSheets()[0];
    sh.setName('접수');
    sh.appendRow(CFG.HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CFG.HEADERS.length).setFontWeight('bold');
  }
  return ss.getSheets()[0];
}

function getFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var folder = DriveApp.createFolder(CFG.FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
