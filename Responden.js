/**
 * Responden Module
 */

function getPertanyaan() {
  if (SETTINGS.USE_FIREBASE) {
    const pert = Firebase.getCachedMasterPertanyaan();
    if (!pert) return [];
    return Object.values(pert).map(p => [
      Firebase.unescapeKey(p.id_soal),
      p.aspek || "",
      p.pertanyaan || "",
      p.kriteria || "",
      p.bobot || 0
    ]);
  }
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Pertanyaan").getDataRange().getValues().slice(1); 
}

function simpanSemuaJawaban(payload) {
  const ts = new Date().toISOString();
  
  if (SETTINGS.USE_FIREBASE) {
    const opd = Firebase.escapeKey(payload.opd);
    const answers = {};
    payload.jawaban.forEach(item => {
      answers[Firebase.escapeKey(item.id)] = {
        timestamp: ts,
        skala: item.skala,
        link: item.link
      };
    });
    Firebase.patch(`jawaban/${opd}`, answers);
    return "Berhasil";
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jawaban");
  const ts_date = new Date();
  const rows = payload.jawaban.map(item => [ts_date, payload.opd, item.id, item.skala, item.link]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return "Berhasil";
}

/**
 * Cek apakah periode pengisian sedang terbuka untuk OPD tertentu.
 * Urutan prioritas: Per-OPD → Global → Tidak ada pengaturan (terbuka)
 * @param {string} namaOPD - Nama OPD yang akan dicek
 * @returns {{ isOpen: boolean, pesan: string, buka: string, tutup: string }}
 */
function cekStatusPeriode(namaOPD) {
  const now = new Date();
  const fmt = (d) => Utilities.formatDate(new Date(d), "GMT+7", "dd MMM yyyy HH:mm");

  let settingOPD = null;
  let settingGlobal = null;

  if (SETTINGS.USE_FIREBASE) {
    const peng = Firebase.get("pengaturan") || {};
    const keyOPD = Firebase.escapeKey(namaOPD.trim().toUpperCase());
    settingOPD = peng[keyOPD];
    settingGlobal = peng["GLOBAL"];
    
    let setting = settingOPD || settingGlobal;
    
    if (!setting || !setting.tglBuka || !setting.tglTutup) {
      return { isOpen: true, pesan: "Tidak ada pengaturan periode aktif.", buka: "", tutup: "" };
    }
    
    const tglBuka = new Date(setting.tglBuka);
    const tglTutup = new Date(setting.tglTutup);
    
    if (now < tglBuka) {
      return { isOpen: false, pesan: `Periode pengisian belum dibuka. Dibuka pada ${fmt(tglBuka)}.`, buka: fmt(tglBuka), tutup: fmt(tglTutup) };
    }
    if (now > tglTutup) {
      return { isOpen: false, pesan: `Periode pengisian telah ditutup pada ${fmt(tglTutup)}.`, buka: fmt(tglBuka), tutup: fmt(tglTutup) };
    }
    return { isOpen: true, pesan: `Pengisian terbuka hingga ${fmt(tglTutup)}.`, buka: fmt(tglBuka), tutup: fmt(tglTutup) };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pengaturan");

  if (!sheet || sheet.getLastRow() < 2) {
    return { isOpen: true, pesan: "Tidak ada pengaturan periode aktif.", buka: "", tutup: "" };
  }

  const rows = sheet.getDataRange().getValues().slice(1); // skip header

  // Cari pengaturan per-OPD dulu
  let setting = rows.find(r => r[0].toString().trim().toUpperCase() === namaOPD.toString().trim().toUpperCase());

  // Fallback ke pengaturan GLOBAL
  if (!setting) {
    setting = rows.find(r => r[0].toString().trim().toUpperCase() === "GLOBAL");
  }

  // Jika tidak ada pengaturan sama sekali, buka
  if (!setting || !setting[1] || !setting[2]) {
    return { isOpen: true, pesan: "Tidak ada pengaturan periode aktif.", buka: "", tutup: "" };
  }

  const tglBuka = new Date(setting[1]);
  const tglTutup = new Date(setting[2]);

  if (now < tglBuka) {
    return {
      isOpen: false,
      pesan: `Periode pengisian belum dibuka. Dibuka pada ${fmt(tglBuka)}.`,
      buka: fmt(tglBuka),
      tutup: fmt(tglTutup)
    };
  }

  if (now > tglTutup) {
    return {
      isOpen: false,
      pesan: `Periode pengisian telah ditutup pada ${fmt(tglTutup)}.`,
      buka: fmt(tglBuka),
      tutup: fmt(tglTutup)
    };
  }

  return {
    isOpen: true,
    pesan: `Pengisian terbuka hingga ${fmt(tglTutup)}.`,
    buka: fmt(tglBuka),
    tutup: fmt(tglTutup)
  };
}
