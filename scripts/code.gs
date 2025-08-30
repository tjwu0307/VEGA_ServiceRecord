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
    var form = payload.form || {};

    // 欄位防呆檢查
    var requiredFields = ['projectCode', 'ticketNo', 'customerName', 'address', 'contact', 'phone', 'rocY', 'rocM', 'rocD'];
    var missing = [];
    requiredFields.forEach(function(field) {
      if (!form[field] || form[field].toString().trim() == '') missing.push(field);
    });
    if (missing.length > 0) {
      Logger.log('缺少欄位: ' + missing.join(', '));
      return ContentService.createTextOutput("缺少欄位：" + missing.join(',')).setMimeType(ContentService.MimeType.TEXT);
    }

    // 儲存PDF到雲端
    var pdfUrl = '';
    var pdfBlob = null;
    if (payload.pdfBase64 && payload.filename) {
      var base64 = payload.pdfBase64;
      if (base64.indexOf(',') > -1) base64 = base64.split(',')[1];

      Logger.log('接收的 base64 字串長度: ' + base64.length); // base64資料長度檢查

      // PDF magic number驗證
      var bins = Utilities.base64Decode(base64);
      var head = '';
      for (var i = 0; i < 4; i++) head += String.fromCharCode(bins[i]);
      Logger.log('PDF 前導字串（應為%PDF）：' + head);
      if (head !== '%PDF') {
        Logger.log('警告：Base64內容未正常包含PDF檔頭');
        return ContentService.createTextOutput("錯誤：PDF base64內容異常").setMimeType(ContentService.MimeType.TEXT);
      }

      try {
        pdfBlob = Utilities.newBlob(bins, 'application/pdf', payload.filename);
        var pdfFile = DriveApp.getFolderById(FOLDER_ID).createFile(pdfBlob);
        pdfUrl = pdfFile.getUrl();
      } catch(err) {
        Logger.log("PDF存檔失敗：" + err.message);
        return ContentService.createTextOutput("PDF存檔失敗：" + err.message).setMimeType(ContentService.MimeType.TEXT);
      }
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

    // 寫入表單到指定試算表（強化防呆）
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("VEGA服務記錄表_回傳");
    var rocDate = `民國${form.rocY}年${form.rocM}月${form.rocD}日`;
    try {
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
        engineerUrl || '',
        customerUrl || '',
        pdfUrl || ''
      ]);
    } catch(rowErr) {
      Logger.log('寫入VEGA服務記錄表_回傳失敗：' + rowErr.message);
      return ContentService.createTextOutput("寫入試算表失敗：" + rowErr.message).setMimeType(ContentService.MimeType.TEXT);
    }

    // Email功能（加上null防呆）
    if (payload.to) {
      GmailApp.sendEmail(payload.to, '服務記錄表 PDF', '附件為 PDF', {attachments: pdfBlob ? [pdfBlob] : []});
    }
    return ContentService.createTextOutput(JSON.stringify({result:"ok"})).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("主流程錯誤：" + err.message);
    return ContentService.createTextOutput("Error: " + err.message).setMimeType(ContentService.MimeType.TEXT);
  }

}
