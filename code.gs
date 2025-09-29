const SHEET_NAME = 'datauser'; // tên sheet chứa data - phải copy đúng tên sheet
const HEADER_ROW = 1;

function generateTokensSafe() {
  const sh = (SHEET_NAME && SHEET_NAME.trim()) ?
               SpreadsheetApp.getActive().getSheetByName(SHEET_NAME.trim()) :
               SpreadsheetApp.getActiveSheet();
  if (!sh) { SpreadsheetApp.getUi().alert('Không tìm thấy sheet. Kiểm tra SHEET_NAME.'); return; }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= HEADER_ROW) { SpreadsheetApp.getUi().alert('Sheet chỉ có header hoặc rỗng.'); return; }

  // đọc header và chuẩn hóa
  let header = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0].map(h => (h || '').toString().trim());
  // tìm cột token (case-insensitive)
  let tokenIdx = header.findIndex(h => h.toLowerCase() === 'token');

  // nếu không thấy, tạo cột Token mới ở cuối
  if (tokenIdx === -1) {
    sh.getRange(HEADER_ROW, lastCol + 1).setValue('Token');
    tokenIdx = lastCol; // index 0-based: new column là lastCol (0-based)
    header = sh.getRange(HEADER_ROW, 1, 1, lastCol + 1).getValues()[0].map(h => (h || '').toString().trim());
  }

  const tokenCol = tokenIdx + 1; // 1-based
  const numDataRows = lastRow - HEADER_ROW;
  if (numDataRows <= 0) { SpreadsheetApp.getUi().alert('Không có hàng dữ liệu để tạo token.'); return; }

  const rng = sh.getRange(HEADER_ROW + 1, tokenCol, numDataRows, 1);
  const vals = rng.getValues();

  for (let i = 0; i < vals.length; i++) {
    if (!vals[i][0] || String(vals[i][0]).trim() === '') {
      // mã ngắn gọn + số ngẫu nhiên để giảm trùng
      vals[i][0] = Utilities.getUuid().split('-')[0] + Math.floor(Math.random() * 9000 + 1000);
    }
  }
  rng.setValues(vals);
  SpreadsheetApp.getUi().alert('Hoàn tất tạo token cho ' + numDataRows + ' hàng (cột ' + tokenCol + ').');
}

// Xuất ra Docs để QR kèm label tên KH
function exportQRCodesWithLabels() {
  const SHEET_NAME = "datauser";
   const DOC_ID = "1jp2r*********************Emii47qQ"; // id file docx
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  const header = data[0];
  
  const nameIndex  = header.indexOf("Họ tên");
  const phoneIndex = header.indexOf("Số điện thoại");
  const tokenIndex = header.indexOf("Token");
  const groupIndex = header.indexOf("Nhóm");
  let statusIndex   = header.indexOf("ExportStatus");

  if (nameIndex === -1 || phoneIndex === -1 || tokenIndex === -1 || groupIndex === -1) {
    throw new Error("Sheet bị thiếu các cột cần thiết. Hãy liên hệ Dev để fix");
  }

  const baseUrl = "https://script.google.com/macros/s/AKfycbzcDh******************************************WS/exec"; // link url app triển khai
  // const doc = DocumentApp.create("Danh sách QR Code khách mời");
  const doc = DocumentApp.openById(DOC_ID);
  const body = doc.getBody();

  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameIndex];
    const phone = data[i][phoneIndex];
    const token = data[i][tokenIndex];
    const status = data[i][statusIndex];
    const group = data[i][groupIndex];

    if (!token || status === "Đã xuất") continue;

    const checkinUrl = baseUrl + "?k=" + encodeURIComponent(token);
    const qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(checkinUrl);

    const body = doc.getBody();
    body.appendParagraph(name + " (" + phone + ") - Nhóm: " + group).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendImage(UrlFetchApp.fetch(qrUrl).getBlob()).setWidth(300);
    Utilities.sleep(2500);
    body.appendPageBreak();
  }

    // Đánh dấu đã xuất trong Google Sheet
   // sh.getRange(i + 1, statusIndex + 1).setValue("Đã xuất");
  // Logger.log("File tạo xong: " + doc.getUrl());
}

// Tạo Web App check-in
function doGet(e) {
  const token = (e.parameter.k || "").trim();
  if (!token) return HtmlService.createHtmlOutput("Thiếu mã");

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  let header = data[0].map(h => (h || "").toString().trim());

  const findIndex = (arr, name) => arr.findIndex(h => h.toLowerCase() === name.toLowerCase());

  let tokenIndex  = findIndex(header, "Token");
  let statusIndex = findIndex(header, "Đã tham dự");
  let nameIndex   = findIndex(header, "Họ tên");
  let timeIndex   = findIndex(header, "CheckinAt");
  let dateIndex   = findIndex(header, "Ngày");
  let hourIndex   = findIndex(header, "Giờ");

  // Nếu thiếu thì thêm cột mới
  if (statusIndex === -1) {
    sh.getRange(1, header.length + 1).setValue("Đã tham dự");
    statusIndex = header.length;
    header.push("Đã tham dự");
  }
  if (timeIndex === -1) {
    sh.getRange(1, header.length + 1).setValue("CheckinAt");
    timeIndex = header.length;
    header.push("CheckinAt");
  }
  if (dateIndex === -1) {
    sh.getRange(1, header.length + 1).setValue("Ngày");
    dateIndex = header.length;
    header.push("Ngày");
  }
  if (hourIndex === -1) {
    sh.getRange(1, header.length + 1).setValue("Giờ");
    hourIndex = header.length;
    header.push("Giờ");
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][tokenIndex]).trim() === token) {
      const row = i + 1;
      const name = data[i][nameIndex] || "Khách";
      const status = sh.getRange(row, statusIndex + 1).getValue();

      if (status !== "Có") {
        const now = new Date();
        const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
        const hourStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");

        sh.getRange(row, statusIndex + 1).setValue("Có");
        sh.getRange(row, timeIndex + 1).setValue(now);       // full datetime
        sh.getRange(row, dateIndex + 1).setValue(dateStr);   // chỉ ngày
        sh.getRange(row, hourIndex + 1).setValue(hourStr);   // chỉ giờ

        // return HtmlService.createHtmlOutput(name + " đã check-in thành công lúc " + hourStr + " ngày " + dateStr);
        return HtmlService.createHtmlOutput(`
          <html>
            <head>
              <meta name="viewport" content="width=device-width, initial-scale=1.0">
              <style>
                body {
                  font-family: Arial, sans-serif;
                  text-align: center;
                  padding: 30px;
                  background: #f4f9ff;
                  color: #333;
                }
                h1 {
                  color: #2e7d32;
                  font-size: 26px;
                  margin-bottom: 15px;
                }
                p {
                  font-size: 18px;
                  margin-bottom: 20px;
                }
                img {
                  width: 120px;
                  margin-bottom: 20px;
                }
                button {
                  background: #1976d2;
                  color: white;
                  border: none;
                  padding: 12px 20px;
                  font-size: 16px;
                  border-radius: 8px;
                  cursor: pointer;
                }
                button:hover {
                  background: #0d47a1;
                }
              </style>
            </head>
            <body>
              <img src="https://cdn-icons-png.flaticon.com/512/190/190411.png" alt="success" />
              <h1>Check-in thành công</h1>
              <p>Xin chào, ${name} đã check-in lúc <b>${hourStr}</b> ngày <b>${dateStr}</b></p>
              <button onclick="window.close()">Hãy tận hưởng tại Learning Chain</button>
            </body>
          </html>
        `);
      } else {
        return HtmlService.createHtmlOutput("Uh oh! " + name + " đã check-in trước đó");
      }
    }
  }
  return HtmlService.createHtmlOutput("Mã không hợp lệ");
}

