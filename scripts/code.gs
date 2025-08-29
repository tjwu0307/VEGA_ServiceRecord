const FOLDER_ID = '1YY9e1mTTWAnXhgCoq9oUzyJe27MUwIlX';
const SHEET_ID = '1I_vz6Psk2g_5yOj0sBUlVTquv1RLEJ50xJfTp6wKAWA';

function doGet(e) {
  // 回傳主要頁面，index.html需新增在Script專案裡
  return HtmlService.createHtmlOutputFromFile('index');
}

function doPost(e) {
  Logger.log('postData: ' + (e.postData ? e.postData.contents : 'no data'));
  // 輸入防錯與debug
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput("No payload").setMimeType(ContentService.MimeType.TEXT);
  }
  try {
    var payload = JSON.parse(e.postData.contents);

    // 儲存PDF到雲端
    var pdfUrl = '';
    if (payload.pdfBase64 && payload.filename) {
      var base64 = payload.pdfBase64.split(',')[1];
      var pdfBlob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', payload.filename);
      var pdfFile = DriveApp.getFolderById(FOLDER_ID).createFile(pdfBlob);
      pdfUrl = pdfFile.getUrl();
    }

    // 儲存簽名
    function saveSignature(dataUrl, name) {
      if (!dataUrl) return '';
      var bytes = Utilities.base64Decode(dataUrl.split(',')[1]);
      var blob = Utilities.newBlob(bytes, 'image/png', `${name}.png`);
      var file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
      return file.getUrl();
    }
    var engineerUrl = (payload.signatures && payload.signatures.engineer) ? saveSignature(payload.signatures.engineer, 'sign_engineer') : '';
    var customerUrl = (payload.signatures && payload.signatures.customer) ? saveSignature(payload.signatures.customer, 'sign_customer') : '';

    // 寫入表單到指定試算表
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("VEGA服務記錄表_回傳");
    var serial = sheet.getLastRow() + 1;
    var form = payload.form || {};
    var rocDate = `民國${form.rocY}年${form.rocM}月${form.rocD}日`;

    sheet.appendRow([
      form.timestamp || new Date(),
      form.projectCode || '',
      form.ticketNo || '',
      form.customerName || '',
      form.address || '',
      form.contact || '',
      form.phone || '',
      rocDate,
      form.arriveTxt || '',
      form.finishTxt || '',
      form.jobNo || '',
      form.products || '',
      form.serviceContent || '',
      form.customerFeedback || '',
      engineerUrl,
      customerUrl,
      pdfUrl
    ]);

function getTemplate() {
  return HtmlService.createHtmlOutputFromFile('pdf_template').getContent();
}

function appendData(formData) {
  const sheet = SpreadsheetApp.openById("1I_vz6Psk2g_5yOj0sBUlVTquv1RLEJ50xJfTp6wKAWA").getSheetByName("回傳");
  sheet.appendRow([
    new Date(),
    formData.serialNo,
    formData.projectId,
    formData.ticketNo,
    formData.customer,
    formData.address,
    formData.contact,
    formData.phone,
    formData.serviceDate,
    formData.arrivalTime,
    formData.finishTime,
    formData.serviceMethod.join(', '),
    formData.serviceProduct.join(', '),
    formData.serviceContent,
    formData.customerFeedback,
    formData.engineer
  ]);
}

function saveSignature(dataUrl, filename) {
  const folder = DriveApp.getFolderById("1YY9e1mTTWAnXhgCoq9oUzyJe27MUwIlX");  // 修改為你的資料夾 ID
  const blob = Utilities.base64Decode(dataUrl.split(',')[1]);
  folder.createFile(blob).setName(filename);
}

function savePDFtoDrive(base64, filename) {
  const folder = DriveApp.getFolderById("1YY9e1mTTWAnXhgCoq9oUzyJe27MUwIlX");
  const blob = Utilities.base64Decode(base64.split(',')[1]);
  folder.createFile(blob).setName(filename);
}

function sendEmailWithPDF(base64, filename, data) {
  const recipient = "your_notify_email@t-tech.com.tw";  // ← 收件者改你要的
  const blob = Utilities.base64Decode(base64.split(',')[1]);
  const pdfBlob = Utilities.newBlob(blob, MimeType.PDF, filename);
  MailApp.sendEmail({
    to: recipient,
    subject: `[服務記錄] ${data.customer} ${data.serviceDate}`,
    body: `您好，請查收服務記錄 PDF 附件。\n\n客戶：${data.customer}\n工程師：${data.engineer}`,
    attachments: [pdfBlob]
  });
}

    // Email功能
if (payload.to) {
  GmailApp.sendEmail(payload.to, '服務記錄表 PDF', '附件為 PDF', {attachments: pdfBlob ? [pdfBlob] : []});
}
    return ContentService.createTextOutput(JSON.stringify({result:"ok"})).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService.createTextOutput("Error: " + e.message).setMimeType(ContentService.MimeType.TEXT);
  }

}
