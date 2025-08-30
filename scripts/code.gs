const FOLDER_ID = '1YY9e1mTTWAnXhgCoq9oUzyJe27MUwIlX';
const SHEET_ID = '1I_vz6Psk2g_5yOj0sBUlVTquv1RLEJ50xJfTp6wKAWA';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}

// 供前端 base64 logo 呼叫
function getLogoBase64() {
  return "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAA..."; // 請換成你的 logo
}

// 完整可用 doPost：同時支援 text/plain 與 FormData 兩種送法
function doPost(e) {
  try {
    // ---- 1) 解析 payload（支援兩種前端送法）----
    var raw = (e && e.postData && e.postData.contents) || (e && e.parameter && e.parameter.data);
    if (!raw) return _json({ ok: false, error: 'No payload' });

    var payload;
    try {
      payload = JSON.parse(raw);
    } catch (err) {
      return _json({ ok: false, error: 'Invalid JSON: ' + err.message, rawPreview: String(raw).slice(0, 200) });
    }

    // form 可能在 payload.form 或 payload.data
    var form = payload.form || payload.data || {};

    // ---- 2) 儲存 PDF 到雲端 ----
    var pdfUrl = '';
    var pdfBlob = null;
    if (payload.pdfBase64 && payload.filename) {
      var base64 = String(payload.pdfBase64).split(',')[1] || '';
      if (base64) {
        pdfBlob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', payload.filename);
        var pdfFile = DriveApp.getFolderById(FOLDER_ID).createFile(pdfBlob);
        pdfUrl = pdfFile.getUrl();
      }
    }

    // ---- 3) 儲存簽名到雲端（支援 engineer / customer / manager 三種；有才存）----
    var folder = DriveApp.getFolderById(FOLDER_ID);
    function saveSignature(dataUrl, namePrefix) {
      if (!dataUrl) return '';
      var b64 = String(dataUrl).split(',')[1] || '';
      if (!b64) return '';
      var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', namePrefix + '_' + Date.now() + '.png');
      var f = folder.createFile(blob);
      return f.getUrl();
    }
    var signatures = payload.signatures || {};
    var engineerUrl = signatures.engineer ? saveSignature(signatures.engineer, 'sign_engineer') : '';
    var customerUrl = signatures.customer ? saveSignature(signatures.customer, 'sign_customer') : '';
    var managerUrl  = signatures.manager  ? saveSignature(signatures.manager,  'sign_manager')  : '';

    // ---- 4) 寫回試算表 ----
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('VEGA服務記錄表_回傳');
    if (!sheet) throw new Error('找不到工作表：VEGA服務記錄表_回傳');

    var rocDate = '民國' + _safe(form.rocY) + '年' + _safe(form.rocM) + '月' + _safe(form.rocD) + '日';

    // 兼容不同鍵名 & 陣列/字串
    var methods = (form.serviceMethods != null ? form.serviceMethods : form.serviceMethod);
    if (Array.isArray(methods)) methods = methods.join(', ');
    methods = _safe(methods);

    var products = (form.products != null ? form.products : form.serviceProduct);
    if (Array.isArray(products)) products = products.join(', ');
    products = _safe(products);

    // 你的原表頭順序（依你貼的 appendRow）
    sheet.appendRow([
      form.timestamp || new Date(),            // 時戳
      _safe(form.projectCode),
      _safe(form.ticketNo),
      _safe(form.customerName),
      _safe(form.address),
      _safe(form.contact),
      _safe(form.phone),
      rocDate,
      _safe(form.arriveTxt),
      _safe(form.finishTxt),
      _safe(form.jobNo),                       // 若未使用可留空
      products,
      _safe(form.serviceContent),
      _safe(form.customerFeedback),
      engineerUrl,
      customerUrl,
      pdfUrl
      // 若你的試算表有「manager 簽名網址」欄位，可把 managerUrl 加在這裡
    ]);

    // ---- 5) 視需求寄信（有填 sendTo/to 或 sendEmail=true）----
    var recipient = payload.sendTo || payload.to || '';
    var doSend = !!recipient || payload.sendEmail === true;
    var mailResult = '';

    if (doSend) {
      var target = recipient || Session.getActiveUser().getEmail(); // 沒指定就寄給自己
      var subject = '服務記錄表 PDF';
      var body = '附件為 PDF。\n客戶：' + _safe(form.customerName) + '\nPDF：' + pdfUrl;
      var opt = {};
      if (pdfBlob) opt.attachments = [pdfBlob];
      GmailApp.sendEmail(target, subject, body, opt);
      mailResult = 'Email sent to ' + target;
    }

    // ---- 6) 回傳 JSON ----
    return _json({
      ok: true,
      pdfUrl: pdfUrl,
      signUrls: { engineer: engineerUrl, customer: customerUrl, manager: managerUrl },
      mailResult: mailResult
    });

  } catch (err) {
    return _json({ ok: false, error: err.message || String(err) });
  }
}

/*** 小工具 ***/
function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function _safe(v) { return (v == null ? '' : String(v)); }

