const FOLDER_ID = '1YY9e1mTTWAnXhgCoq9oUzyJe27MUwIlX';
const SHEET_ID = '1I_vz6Psk2g_5yOj0sBUlVTquv1RLEJ50xJfTp6wKAWA';

function doGet(e) {
  // 回傳主要頁面，index.html需新增在Script專案裡
  return HtmlService.createHtmlOutputFromFile('index');
}

function doPost(e) {
  Logger.log('postData: ' + (e.postData ? e.postData.contents : 'no data'));

  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput("No payload").setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    var payload = JSON.parse(e.postData.contents);

    // 儲存 PDF
    var pdfUrl = '';
    if (payload.pdfBase64 && payload.filename) {
      var base64 = payload.pdfBase64.split(',')[1];
      var pdfBlob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', payload.filename);
      var pdfFile = DriveApp.getFolderById(FOLDER_ID).createFile(pdfBlob);
      pdfUrl = pdfFile.getUrl();
    }

    // 儲存簽名
    var engineerUrl = payload.signatures?.engineer ? saveSignature(payload.signatures.engineer, 'sign_engineer') : '';
    var customerUrl = payload.signatures?.customer ? saveSignature(payload.signatures.customer, 'sign_customer') : '';
    var managerUrl  = payload.signatures?.manager  ? saveSignature(payload.signatures.manager,  'sign_manager')  : '';

    // 寫入 Google Sheet
    var form = payload.form || {};
    appendData({
      serialNo: form.serialNo,
      projectId: form.projectId,
      ticketNo: form.ticketNo,
      customer: form.customer,
      address: form.address,
      contact: form.contact,
      phone: form.phone,
      serviceDate: form.serviceDate,
      arrivalTime: form.arrivalTime,
      finishTime: form.finishTime,
      serviceMethod: form.serviceMethod || [],
      serviceProduct: form.serviceProduct || [],
      serviceContent: form.serviceContent,
      customerFeedback: form.customerFeedback,
      engineer: form.engineer
    });

    // 寄送 Email
    if (payload.filename && pdfBlob) {
      sendEmailWithPDF(payload.pdfBase64, payload.filename, form);
    }

    return ContentService.createTextOutput(JSON.stringify({ result: "ok", pdfUrl })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    Logger.log("doPost error: " + e.message);
    return ContentService.createTextOutput("Error: " + e.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

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
