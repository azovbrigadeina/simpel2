/**
 * Responden Module
 */

function getPertanyaan() { 
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Pertanyaan").getDataRange().getValues().slice(1); 
}

function simpanSemuaJawaban(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jawaban");
  const ts = new Date();
  const rows = payload.jawaban.map(item => [ts, payload.opd, item.id, item.skala, item.link]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return "Berhasil";
}
