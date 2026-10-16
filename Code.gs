// Secret Santa poll: writes each friend's ticks into the "Poll" tab of the Google Sheet.
// Setup: Extensions > Apps Script, paste this in, then Deploy > New deployment > Web app,
// "Execute as: Me", "Who has access: Anyone". Copy the web app URL into index.html.

const SLOT_IDS = ["s1", "s2", "s3", "s4", "s5", "s6", "s7", "s8"]; // sheet columns B to I
const FIRST_ROW = 5;   // Anni
const ROW_COUNT = 7;   // seven friends

function rows_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Poll");
  return { sheet, values: sheet.getRange(FIRST_ROW, 1, ROW_COUNT, 9).getValues() };
}

function nameOf_(cell) {
  return String(cell).split(" (")[0];
}

function doGet() {
  const answers = {};
  rows_().values.forEach(function (row) {
    answers[nameOf_(row[0])] = SLOT_IDS.filter(function (id, i) { return row[i + 1] === true; });
  });
  return ContentService.createTextOutput(JSON.stringify({ answers: answers }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const picked = (body.slots || []).filter(function (id) { return SLOT_IDS.indexOf(id) !== -1; });
    const r = rows_();
    for (let i = 0; i < r.values.length; i++) {
      if (nameOf_(r.values[i][0]) === body.name) {
        const row = SLOT_IDS.map(function (id) { return picked.indexOf(id) !== -1; });
        r.sheet.getRange(FIRST_ROW + i, 2, 1, SLOT_IDS.length).setValues([row]);
        break;
      }
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("ok");
}
