// Google Apps Script — вставь в script.google.com, привяжи к своей таблице
// Деплой: Развернуть → Новое развёртывание → Тип: Веб-приложение
//   Выполнять от имени: Меня
//   Доступ: Все
// Скопируй URL развёртывания и вставь в index.html → SCRIPT_URL

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data  = JSON.parse(e.postData.contents);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Дата', 'Имя', 'Контакт', 'Тема', 'Описание', 'Статус']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#D4AF6A');
    }

    sheet.appendRow([
      new Date(),
      data.name,
      data.contact,
      data.topic,
      data.description,
      'Новая',
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('OK');
}
