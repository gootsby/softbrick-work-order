const SPREADSHEET_ID = ""; // Standalone script로 쓸 경우 시트 ID를 넣으세요. 시트에 묶인 스크립트면 비워둡니다.
const SHEET_NAME = "작업지시서";
const IMAGE_FOLDER_NAME = "작업지시서_소프트브릭_이미지";

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
  "configManual",
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
  "productImageUrl",
  "fabricImageUrl",
  "option1ImageUrl",
  "option2ImageUrl"
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

function getImageFolder_() {
  const folders = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(IMAGE_FOLDER_NAME);
}

function imageExtension_(mimeType) {
  const map = {
    "image/jpeg": "jpg",
    "image/png": "png",
    "image/webp": "webp",
    "image/gif": "gif"
  };
  return map[mimeType] || "png";
}

function saveImageData_(dataUrl, fileNameBase) {
  const match = String(dataUrl || "").match(/^data:([^;]+);base64,(.+)$/);
  if (!match) return "";

  const mimeType = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const extension = imageExtension_(mimeType);
  const blob = Utilities.newBlob(bytes, mimeType, `${fileNameBase}.${extension}`);
  const file = getImageFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1600`;
}

function normalizeParams_(params) {
  const next = Object.assign({}, params);
  const docNo = String(next.docNo || "work-order").replace(/[^\w가-힣-]+/g, "_");
  const imageFields = [
    ["productImageData", "productImageUrl", "product"],
    ["fabricImageData", "fabricImageUrl", "fabric"],
    ["option1ImageData", "option1ImageUrl", "label"],
    ["option2ImageData", "option2ImageUrl", "etc"]
  ];

  imageFields.forEach(([dataKey, urlKey, label]) => {
    if (next[dataKey]) {
      next[urlKey] = saveImageData_(next[dataKey], `${docNo}_${label}_${Date.now()}`);
    }
    delete next[dataKey];
  });

  return next;
}

function saveRecord_(params) {
  params = normalizeParams_(params);
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || HEADERS;
  const docNo = params.docNo || "";
  const docNoIndex = headers.indexOf("docNo");

  if (!docNo) {
    return { ok: false, error: "docNo is required" };
  }

  const nextRow = headers.map(header => {
    if (header === "updatedAt") return new Date();
    return params[header] || "";
  });

  const existingOffset = values.slice(1).findIndex(row => row[docNoIndex] === docNo);
  if (existingOffset >= 0) {
    sheet.getRange(existingOffset + 2, 1, 1, headers.length).setValues([nextRow]);
    return { ok: true, docNo, mode: "updated" };
  }

  sheet.appendRow(nextRow);
  return { ok: true, docNo, mode: "created" };
}

function json_(payload, callback) {
  const text = JSON.stringify(payload);
  const output = callback ? `${callback}(${text});` : text;
  return ContentService
    .createTextOutput(output)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function postMessage_(payload) {
  const text = JSON.stringify(payload).replace(/</g, "\\u003c");
  return HtmlService.createHtmlOutput(
    `<script>
      (function () {
        var payload = ${text};
        function send() {
          try { window.parent.postMessage(payload, "*"); } catch (e) {}
          try { window.top.postMessage(payload, "*"); } catch (e) {}
          try { if (window.opener) window.opener.postMessage(payload, "*"); } catch (e) {}
        }
        send();
        setTimeout(send, 300);
        setTimeout(send, 1000);
      })();
    </script>`
  );
}

function authorizeOnce() {
  getSheet_();
  getImageFolder_();
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

  if (action === "save") {
    return json_(saveRecord_(params), callback);
  }

  return json_({ ok: false, error: "Unknown action" }, callback);
}

function doPost(e) {
  const params = e.parameter || {};
  try {
    const result = saveRecord_(params);
    result.postToken = params.postToken || "";
    return postMessage_(result);
  } catch (error) {
    return postMessage_({
      ok: false,
      postToken: params.postToken || "",
      error: error && error.message ? error.message : String(error)
    });
  }
}
