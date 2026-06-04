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

/**
 * Fungsi Wrapper untuk Push Data secara Interaktif dengan Umpan Balik UI
 */
function syncSheetsToFirebaseInteractive() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Konfirmasi Push Data",
    "Apakah Anda yakin ingin MENGIRIM seluruh data dari Sheets saat ini untuk menimpa database Firebase?\n\n(Tindakan ini akan menimpa data di Firebase dengan data dari Sheets ini)",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    migrateDataToFirebase();
    ui.alert("Sukses", "Seluruh data Sheets berhasil dikirim dan disinkronkan ke Firebase!", ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Gagal", "Terjadi kesalahan saat sinkronisasi: " + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Menarik seluruh data dari Firebase RTDB kembali ke Google Sheets untuk SImpel2
 */
function pullFirebaseToSheetsInteractive() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Konfirmasi Tarik Data (Pull)",
    "Apakah Anda yakin ingin MENARIK seluruh data dari Firebase?\n\n(Tindakan ini akan MEMBERSIHKAN dan MENIMPA data pada Sheets saat ini dengan data terbaru dari Firebase)",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 1. Pull Users
    const shUsers = ss.getSheetByName("Users");
    if (shUsers) {
      const usersData = Firebase.get("users") || {};
      const rows = [["Username", "Password", "Role", "Nama_OPD"]];
      Object.keys(usersData).forEach(escapedKey => {
        const u = usersData[escapedKey];
        rows.push([
          Firebase.unescapeKey(escapedKey),
          u.password || "",
          u.role || "",
          u.nama_opd || ""
        ]);
      });
      shUsers.clearContents();
      shUsers.getRange(1, 1, rows.length, 4).setValues(rows);
      Logger.log("Users pulled successfully.");
    }

    // 2. Pull Master OPD
    const shOPD = ss.getSheetByName("Master_OPD");
    if (shOPD) {
      const opdData = Firebase.get("master_opd") || {};
      const rows = [["Nama OPD"]];
      Object.keys(opdData).forEach(escapedKey => {
        rows.push([opdData[escapedKey].nama || Firebase.unescapeKey(escapedKey)]);
      });
      shOPD.clearContents();
      shOPD.getRange(1, 1, rows.length, 1).setValues(rows);
      Logger.log("Master OPD pulled successfully.");
    }

    // 3. Pull Master Pertanyaan
    const shPert = ss.getSheetByName("Master_Pertanyaan");
    if (shPert) {
      const pertData = Firebase.get("master_pertanyaan") || {};
      const rows = [["id_soal", "aspek", "pertanyaan", "kriteria", "bobot"]];
      Object.keys(pertData).forEach(k => {
        const p = pertData[k];
        rows.push([
          Firebase.unescapeKey(p.id_soal || k),
          p.aspek || "",
          p.pertanyaan || "",
          p.kriteria || "",
          p.bobot || 0
        ]);
      });
      shPert.clearContents();
      shPert.getRange(1, 1, rows.length, 5).setValues(rows);
      Logger.log("Master Pertanyaan pulled successfully.");
    }

    // 4. Pull Jawaban
    const shJawaban = ss.getSheetByName("Jawaban");
    if (shJawaban) {
      const jawData = Firebase.get("jawaban") || {};
      const rows = [["Timestamp", "Nama OPD", "ID Soal", "Skala Pilihan / Jawaban", "Link Bukti Dukung"]];
      Object.keys(jawData).forEach(opdEscaped => {
        const opdName = Firebase.unescapeKey(opdEscaped);
        const opdAnswers = jawData[opdEscaped] || {};
        Object.keys(opdAnswers).forEach(idSoalEscaped => {
          const ans = opdAnswers[idSoalEscaped] || {};
          rows.push([
            ans.timestamp || new Date().toISOString(),
            opdName,
            Firebase.unescapeKey(idSoalEscaped),
            ans.skala || "",
            ans.link || ""
          ]);
        });
      });
      shJawaban.clearContents();
      if (rows.length > 1) {
        shJawaban.getRange(1, 1, rows.length, 5).setValues(rows);
      } else {
        shJawaban.getRange(1, 1, 1, 5).setValues(rows);
      }
      Logger.log("Jawaban pulled successfully.");
    }

    // 5. Pull Verifikasi
    const shVerif = ss.getSheetByName("Verifikasi");
    if (shVerif) {
      const verifData = Firebase.get("verifikasi") || {};
      const rows = [["Timestamp", "Nama OPD", "ID Soal", "Skala Responden", "Skala Evaluator", "Catatan Evaluator"]];
      Object.keys(verifData).forEach(opdEscaped => {
        const opdName = Firebase.unescapeKey(opdEscaped);
        const opdVerif = verifData[opdEscaped] || {};
        Object.keys(opdVerif).forEach(idSoalEscaped => {
          const v = opdVerif[idSoalEscaped] || {};
          rows.push([
            v.timestamp || new Date().toISOString(),
            opdName,
            Firebase.unescapeKey(idSoalEscaped),
            v.skala_responden || "",
            v.skala_evaluator || "",
            v.catatan || ""
          ]);
        });
      });
      shVerif.clearContents();
      if (rows.length > 1) {
        shVerif.getRange(1, 1, rows.length, 6).setValues(rows);
      } else {
        shVerif.getRange(1, 1, 1, 6).setValues(rows);
      }
      Logger.log("Verifikasi pulled successfully.");
    }

    // 6. Pull Pengaturan
    const shPeng = ss.getSheetByName("Pengaturan");
    if (shPeng) {
      const pengData = Firebase.get("pengaturan") || {};
      const rows = [["Scope", "Tanggal Buka", "Tanggal Tutup"]];
      Object.keys(pengData).forEach(scopeEscaped => {
        const p = pengData[scopeEscaped] || {};
        rows.push([
          p.tipe || Firebase.unescapeKey(scopeEscaped),
          p.tglBuka || "",
          p.tglTutup || ""
        ]);
      });
      shPeng.clearContents();
      shPeng.getRange(1, 1, rows.length, 3).setValues(rows);
      Logger.log("Pengaturan pulled successfully.");
    }

    // 7. Pull SKM
    const shSKM = ss.getSheetByName("SKM");
    if (shSKM) {
      const skmData = Firebase.get("skm") || {};
      const rows = [["Timestamp", "Nama OPD", "Nilai SKM"]];
      Object.keys(skmData).forEach(opdEscaped => {
        const opdName = Firebase.unescapeKey(opdEscaped);
        const skm = skmData[opdEscaped] || {};
        rows.push([
          skm.timestamp || new Date().toISOString(),
          opdName,
          skm.nilaiSKM || 0
        ]);
      });
      shSKM.clearContents();
      if (rows.length > 1) {
        shSKM.getRange(1, 1, rows.length, 3).setValues(rows);
      } else {
        shSKM.getRange(1, 1, 1, 3).setValues(rows);
      }
      Logger.log("SKM pulled successfully.");
    }

    // 8. Pull Catatan Komponen
    const shCK = ss.getSheetByName("Catatan_Komponen");
    if (shCK) {
      const ckData = Firebase.get("catatan_komponen") || {};
      const rows = [["Timestamp", "Nama OPD", "pt_p", "pt_r", "ai_p", "ai_r", "pm_p", "pm_r", "ep_p", "ep_r"]];
      Object.keys(ckData).forEach(opdEscaped => {
        const opdName = Firebase.unescapeKey(opdEscaped);
        const ck = ckData[opdEscaped] || {};
        rows.push([
          ck.timestamp || new Date().toISOString(),
          opdName,
          ck.pt_p || "",
          ck.pt_r || "",
          ck.ai_p || "",
          ck.ai_r || "",
          ck.pm_p || "",
          ck.pm_r || "",
          ck.ep_p || "",
          ck.ep_r || ""
        ]);
      });
      shCK.clearContents();
      if (rows.length > 1) {
        shCK.getRange(1, 1, rows.length, 10).setValues(rows);
      } else {
        shCK.getRange(1, 1, 1, 10).setValues(rows);
      }
      Logger.log("Catatan Komponen pulled successfully.");
    }

    ui.alert("Sukses", "Seluruh data berhasil ditarik dari Firebase ke Google Sheets!", ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Gagal", "Terjadi kesalahan saat menarik data: " + e.message, ui.ButtonSet.OK);
  }
}

