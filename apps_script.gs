// Google Apps Script — endpoint для форми waitlist.
// Деплой: extensions → Apps Script → New project → вставити цей код →
//   Deploy → New deployment → Type: Web app → Execute as: Me →
//   Who has access: Anyone → Deploy. Скопіювати URL → вставити в index.html (ENDPOINT).

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Waitlist');
    if (!sheet) {
      sheet = ss.insertSheet('Waitlist');
      sheet.appendRow([
        'Дата', 'Email', 'UTM source', 'UTM medium', 'UTM campaign',
        'Source', 'User Agent', 'IP hash'
      ]);
      sheet.setFrozenRows(1);
    }

    var data = JSON.parse(e.postData.contents);

    if (data.hp && data.hp.length > 0) {
      return ok();
    }

    if (!data.email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(data.email)) {
      return fail('invalid email');
    }

    var existing = sheet.getRange('B:B').getValues().flat();
    if (existing.indexOf(data.email) !== -1) {
      return ok();
    }

    sheet.appendRow([
      new Date(),
      data.email,
      data.utm_source || '',
      data.utm_medium || '',
      data.utm_campaign || '',
      data.source || '',
      data.ua || '',
      ''
    ]);

    return ok();
  } catch (err) {
    return fail(err.message);
  }
}

function ok() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function fail(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'NMT-Bot waitlist endpoint' }))
    .setMimeType(ContentService.MimeType.JSON);
}
