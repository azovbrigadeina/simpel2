/**
 * FILE INI SUDAH DIMODULARISASI.
 * Silakan lihat file-file berikut untuk logika backend:
 * - Main.gs (Entry point & helper)
 * - Auth.gs (Login & User management)
 * - Responden.gs (Logika Responden)
 * - Evaluator.gs (Logika Evaluator & Admin)
 */

function debugData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ds_ck = ss.getSheetByName("Catatan_Komponen") && ss.getSheetByName("Catatan_Komponen").getLastRow() > 1 ? ss.getSheetByName("Catatan_Komponen").getDataRange().getValues().slice(1) : [];
  return JSON.stringify(ds_ck);
}