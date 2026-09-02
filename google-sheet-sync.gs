function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const sheetName = '學生作答紀錄';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    const headers = [
      '時間',
      '學生姓名',
      '班級',
      '模式',
      '級別',
      '總分',
      '答對數',
      '總答題數',
      '正確率',
      '最高連擊',
      '作答紀錄(JSON)'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }

    const row = [
      payload.timestamp || new Date().toISOString(),
      payload.studentName || '未填姓名',
      payload.studentClass || '未填班級',
      payload.mode || '',
      payload.selectedLevel || '',
      Number(payload.score || 0),
      Number(payload.correctCount || 0),
      Number(payload.answeredCount || 0),
      Number(payload.accuracyRate || 0),
      Number(payload.maxCombo || 0),
      JSON.stringify(payload.answers || [])
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(
      JSON.stringify({
        ok: true,
        message: 'Google Sheets 同步成功',
        rowCount: sheet.getLastRow()
      })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        ok: false,
        message: error.message || '同步失敗'
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, message: 'Google Apps Script 已啟動' })
  ).setMimeType(ContentService.MimeType.JSON);
}
