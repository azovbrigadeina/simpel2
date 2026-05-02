/**
 * Evaluator & Admin Module
 */

function simpanVerifikasi(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Verifikasi");
  const data = sheet.getDataRange().getValues();
  const ts = new Date();
  payload.items.forEach(item => {
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === payload.opd && data[i][2].toString() === item.id_soal.toString()) { rowIdx = i + 1; break; }
    }
    const val = [ts, payload.opd, item.id_soal, item.skala_responden, item.skala_evaluator, item.catatan];
    if (rowIdx !== -1) sheet.getRange(rowIdx, 1, 1, 6).setValues([val]);
    else sheet.appendRow(val);
  });
  return "Sukses";
}

function getStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resps = ss.getSheetByName("Users").getDataRange().getValues().filter(r => r[2] === "Responden").length;
  const sudah = ss.getSheetByName("Jawaban").getLastRow() > 1 ? [...new Set(ss.getSheetByName("Jawaban").getDataRange().getValues().slice(1).map(r => r[1]))].length : 0;
  return { total: resps, sudah: sudah };
}

function getOPDSudahKirim() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jawaban");
  if (sheet.getLastRow() < 2) return {};
  const data = sheet.getDataRange().getValues().slice(1);
  const uniqueOPD = {};
  data.forEach(r => { uniqueOPD[r[1]] = "Ada"; });
  return uniqueOPD;
}

function getJawabanByOPD(namaOPD) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dj = ss.getSheetByName("Jawaban").getDataRange().getValues().slice(1);
  const ds = ss.getSheetByName("Master_Pertanyaan").getDataRange().getValues().slice(1);
  const dv = ss.getSheetByName("Verifikasi").getLastRow() > 1 ? ss.getSheetByName("Verifikasi").getDataRange().getValues().slice(1) : [];
  return dj.filter(r => r[1] === namaOPD).map(j => {
    const soal = ds.find(s => s[0].toString() === j[2].toString());
    const verif = dv.find(v => v[1] === namaOPD && v[2].toString() === j[2].toString());
    return {
      id_soal: j[2], pertanyaan: soal ? soal[2] : "N/A", skala_responden: j[3], link: j[4],
      skala_evaluator: verif ? verif[4] : "", catatan: verif ? verif[5] : ""
    };
  });
}

function getRekapVerifikasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataV = ss.getSheetByName("Verifikasi").getDataRange().getValues().slice(1);
  const dataS = ss.getSheetByName("Master_Pertanyaan").getDataRange().getValues().slice(1);
  let soalMap = {};
  dataS.forEach(s => { soalMap[s[0].toString()] = { aspek: s[1], pertanyaan: s[2], bobot: Number(s[4] || 0) }; });
  const rekap = {};
  dataV.forEach(r => {
    const opd = r[1]; const idSoal = r[2].toString(); const skorEval = Number(r[4] || 0); const info = soalMap[idSoal];
    if (!info) return;
    if (!rekap[opd]) rekap[opd] = { opd: opd, rincian: [] };
    rekap[opd].rincian.push({ aspek: info.aspek, soal: info.pertanyaan, bobot: info.bobot, skor: skorEval, nilai: ((info.bobot / 5) * skorEval).toFixed(3) });
  });
  return Object.values(rekap);
}

function hapusJawabanOPD(namaOPD) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ["Jawaban", "Verifikasi"].forEach(shName => {
    let sh = ss.getSheetByName(shName);
    let d = sh.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) { if (d[i][1] === namaOPD) sh.deleteRow(i + 1); }
  });
  return "Sukses";
}

function getListOPD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master_OPD");
  if (!sheet) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(row => row[0]).filter(name => name !== "");
}
