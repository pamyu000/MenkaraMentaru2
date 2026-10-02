const SHEET_NAME = "responses";

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["timestamp","date","q1","q2","q3","q4","q5","extra","resultType","subType"]);
  }
  sheet.appendRow([
    new Date(),
    data.createdAt || "",
    data.q1 || "",
    data.q2 || "",
    data.q3 || "",
    data.q4 || "",
    data.q5 || "",
    data.extra || "",
    data.resultType || "",
    data.subType || ""
  ]);
  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.action !== "stats") {
    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return ContentService.createTextOutput(JSON.stringify({total:0,results:{}}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const values = sheet.getDataRange().getValues().slice(1);
  const results = {};
  values.forEach(r => {
    const code = String(r[8] || "");
    if (code) results[code] = (results[code] || 0) + 1;
  });
  return ContentService.createTextOutput(JSON.stringify({
    total: values.length,
    results: results
  })).setMimeType(ContentService.MimeType.JSON);
}