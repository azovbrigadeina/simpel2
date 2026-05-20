/**
 * Skrip Migrasi Data Satu-Kali-Jalan (One-Time Migration)
 * Menjalankan fungsi ini akan menyalin seluruh data dari Google Sheets saat ini ke Firebase Realtime Database.
 */
function migrateDataToFirebase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Migrasi Pengguna (Users)
  Logger.log("Migrasi Users...");
  const shUsers = ss.getSheetByName("Users");
  if (shUsers && shUsers.getLastRow() > 1) {
    const dataUsers = shUsers.getDataRange().getValues().slice(1);
    const usersPayload = {};
    dataUsers.forEach(r => {
      const u = Firebase.escapeKey(r[0].toString().trim());
      if (u) {
        usersPayload[u] = {
          password: r[1].toString().trim(),
          role: r[2].toString().trim(),
          nama_opd: r[3].toString().trim()
        };
      }
    });
    Firebase.put("users", usersPayload);
  }

  // 2. Migrasi Master Pertanyaan
  Logger.log("Migrasi Master Pertanyaan...");
  const shPertanyaan = ss.getSheetByName("Master_Pertanyaan");
  if (shPertanyaan && shPertanyaan.getLastRow() > 1) {
    const dataPertanyaan = shPertanyaan.getDataRange().getValues().slice(1);
    const pertPayload = {};
    dataPertanyaan.forEach(r => {
      const idSoal = Firebase.escapeKey(r[0].toString().trim());
      if (idSoal) {
        pertPayload[idSoal] = {
          id_soal: idSoal,
          aspek: r[1].toString().trim(),
          pertanyaan: r[2].toString().trim(),
          kriteria: r[3].toString().trim(),
          bobot: Number(r[4] || 0)
        };
      }
    });
    Firebase.put("master_pertanyaan", pertPayload);
  }

  // 3. Migrasi Master OPD
  Logger.log("Migrasi Master OPD...");
  const shOPD = ss.getSheetByName("Master_OPD");
  if (shOPD && shOPD.getLastRow() > 1) {
    const dataOPD = shOPD.getRange(2, 1, shOPD.getLastRow() - 1, 1).getValues();
    const opdPayload = {};
    dataOPD.forEach(r => {
      const nama = r[0].toString().trim();
      const opdKey = Firebase.escapeKey(nama);
      if (opdKey) {
        opdPayload[opdKey] = { nama: nama };
      }
    });
    Firebase.put("master_opd", opdPayload);
  }

  // 4. Migrasi Pengaturan Periode
  Logger.log("Migrasi Pengaturan Periode...");
  const shPengaturan = ss.getSheetByName("Pengaturan");
  if (shPengaturan && shPengaturan.getLastRow() > 1) {
    const dataPengaturan = shPengaturan.getDataRange().getValues().slice(1);
    const pengaturanPayload = {};
    dataPengaturan.forEach(r => {
      const tipeKey = Firebase.escapeKey(r[0].toString().trim().toUpperCase());
      if (tipeKey) {
        pengaturanPayload[tipeKey] = {
          tipe: r[0].toString().trim().toUpperCase(),
          tglBuka: r[1] ? new Date(r[1]).toISOString() : "",
          tglTutup: r[2] ? new Date(r[2]).toISOString() : ""
        };
      }
    });
    Firebase.put("pengaturan", pengaturanPayload);
  }

  // 5. Migrasi Jawaban
  Logger.log("Migrasi Jawaban...");
  const shJawaban = ss.getSheetByName("Jawaban");
  if (shJawaban && shJawaban.getLastRow() > 1) {
    const dataJawaban = shJawaban.getDataRange().getValues().slice(1);
    const jawabanPayload = {};
    dataJawaban.forEach(r => {
      const opd = Firebase.escapeKey(r[1].toString().trim());
      const idSoal = Firebase.escapeKey(r[2].toString().trim());
      if (opd && idSoal) {
        if (!jawabanPayload[opd]) jawabanPayload[opd] = {};
        jawabanPayload[opd][idSoal] = {
          timestamp: r[0] ? new Date(r[0]).toISOString() : new Date().toISOString(),
          skala: r[3].toString(),
          link: r[4] ? r[4].toString() : ""
        };
      }
    });
    Firebase.put("jawaban", jawabanPayload);
  }

  // 6. Migrasi Verifikasi
  Logger.log("Migrasi Verifikasi...");
  const shVerif = ss.getSheetByName("Verifikasi");
  if (shVerif && shVerif.getLastRow() > 1) {
    const dataVerif = shVerif.getDataRange().getValues().slice(1);
    const verifPayload = {};
    dataVerif.forEach(r => {
      const opd = Firebase.escapeKey(r[1].toString().trim());
      const idSoal = Firebase.escapeKey(r[2].toString().trim());
      if (opd && idSoal) {
        if (!verifPayload[opd]) verifPayload[opd] = {};
        verifPayload[opd][idSoal] = {
          timestamp: r[0] ? new Date(r[0]).toISOString() : new Date().toISOString(),
          skala_responden: r[3].toString(),
          skala_evaluator: r[4].toString(),
          catatan: (r[5] || "").toString()
        };
      }
    });
    Firebase.put("verifikasi", verifPayload);
  }

  // 7. Migrasi SKM
  Logger.log("Migrasi SKM...");
  const shSKM = ss.getSheetByName("SKM");
  if (shSKM && shSKM.getLastRow() > 1) {
    const dataSKM = shSKM.getDataRange().getValues().slice(1);
    const skmPayload = {};
    dataSKM.forEach(r => {
      const opd = Firebase.escapeKey(r[1].toString().trim());
      if (opd) {
        skmPayload[opd] = {
          timestamp: r[0] ? new Date(r[0]).toISOString() : new Date().toISOString(),
          nilaiSKM: Number(r[2] || 0)
        };
      }
    });
    Firebase.put("skm", skmPayload);
  }

  // 8. Migrasi Catatan Komponen
  Logger.log("Migrasi Catatan Komponen...");
  const shCK = ss.getSheetByName("Catatan_Komponen");
  if (shCK && shCK.getLastRow() > 1) {
    const dataCK = shCK.getDataRange().getValues().slice(1);
    const ckPayload = {};
    dataCK.forEach(r => {
      const opd = Firebase.escapeKey(r[1].toString().trim());
      if (opd) {
        ckPayload[opd] = {
          timestamp: r[0] ? new Date(r[0]).toISOString() : new Date().toISOString(),
          pt_p: (r[2] || "").toString(),
          pt_r: (r[3] || "").toString(),
          ai_p: (r[4] || "").toString(),
          ai_r: (r[5] || "").toString(),
          pm_p: (r[6] || "").toString(),
          pm_r: (r[7] || "").toString(),
          ep_p: (r[8] || "").toString(),
          ep_r: (r[9] || "").toString()
        };
      }
    });
    Firebase.put("catatan_komponen", ckPayload);
  }

  Logger.log("Migrasi Selesai!");
}
