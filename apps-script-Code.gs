const SPREADSHEET_ID = ""; // Standalone script로 쓸 경우 시트 ID를 넣으세요. 시트에 묶인 스크립트면 비워둡니다.
const SHEET_NAME = "작업지시서";

const HEADERS = [
  "docNo",
  "updatedAt",
  "orderDate",
  "dueDate",
  "docQuarter",
  "docSeq",
  "manager",
  "managerPhone",
  "brand",
  "model",
  "config",
  "armrest",
  "fabric",
  "fabricManual",
  "color",
  "colorManual",
  "qty",
  "customer",
  "phone",
  "address",
  "memo",
  "option2Text",
  "productImageUrl"
];

function getSpreadsheet_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsHeaders = HEADERS.some((header, index) => currentHeaders[index] !== header);
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowToObject_(headers, row) {
  return headers.reduce((record, header, index) => {
    record[header] = row[index] || "";
    return record;
  }, {});
}

function json_(payload, callback) {
  const text = JSON.stringify(payload);
  const output = callback ? `${callback}(${text});` : text;
  return ContentService
    .createTextOutput(output)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function doGet(e) {
  const params = e.parameter || {};
  const callback = params.callback || "";
  const action = params.action || "load";
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || HEADERS;
  const docNoIndex = headers.indexOf("docNo");

  if (action === "list") {
    const records = values.slice(1).filter(row => row[docNoIndex]).map(row => rowToObject_(headers, row));
    return json_({ ok: true, records }, callback);
  }

  if (action === "load") {
    const docNo = params.docNo || "";
    const row = values.slice(1).find(item => item[docNoIndex] === docNo);
    return json_({ ok: Boolean(row), record: row ? rowToObject_(headers, row) : null }, callback);
  }

  return json_({ ok: false, error: "Unknown action" }, callback);
}

function doPost(e) {
  const params = e.parameter || {};
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || HEADERS;
  const docNo = params.docNo || "";
  const docNoIndex = headers.indexOf("docNo");

  if (!docNo) {
    return json_({ ok: false, error: "docNo is required" });
  }

  const nextRow = headers.map(header => {
    if (header === "updatedAt") return new Date();
    return params[header] || "";
  });

  const existingOffset = values.slice(1).findIndex(row => row[docNoIndex] === docNo);
  if (existingOffset >= 0) {
    sheet.getRange(existingOffset + 2, 1, 1, headers.length).setValues([nextRow]);
  } else {
    sheet.appendRow(nextRow);
  }

  return json_({ ok: true, docNo });
}
