/**
 * Evaluator & Admin Module
 */

function simpanVerifikasi(payload) {
  if (SETTINGS.USE_FIREBASE) {
    const ts = new Date().toISOString();
    const opd = Firebase.escapeKey(payload.opd);
    
    // Ambil verifikasi saat ini
    const currentVerif = Firebase.get(`verifikasi/${opd}`) || {};
    
    // 1. Update Verifikasi
    payload.items.forEach(item => {
      const idSoal = Firebase.escapeKey(item.id_soal);
      currentVerif[idSoal] = {
        timestamp: ts,
        skala_responden: item.skala_responden,
        skala_evaluator: item.skala_evaluator,
        catatan: item.catatan
      };
    });
    
    // Simpan ke Firebase
    Firebase.put(`verifikasi/${opd}`, currentVerif);
    
    // 2. Update SKM
    if (payload.nilaiSKM !== undefined) {
      Firebase.put(`skm/${opd}`, {
        timestamp: ts,
        nilaiSKM: payload.nilaiSKM
      });
    }
    
    // 3. Update Catatan Komponen
    if (payload.catatanKomponen) {
      const c = payload.catatanKomponen;
      Firebase.put(`catatan_komponen/${opd}`, {
        timestamp: ts,
        pt_p: c.pt_p, pt_r: c.pt_r,
        ai_p: c.ai_p, ai_r: c.ai_r,
        pm_p: c.pm_p, pm_r: c.pm_r,
        ep_p: c.ep_p, ep_r: c.ep_r
      });
    }

    // Invalidasi cache
    try { CacheService.getScriptCache().remove('rekap_verifikasi'); } catch(e) {}
    try { CacheService.getScriptCache().remove('dashboard_stats'); } catch(e) {}
    
    return "Sukses";
  }

  // Fallback Google Sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ts = new Date();
  
  // 1. Batch Update Verifikasi
  const sheetV = ss.getSheetByName("Verifikasi");
  let dataV = sheetV.getDataRange().getValues();
  const headerV = dataV[0];
  let rowsV = dataV.slice(1);

  payload.items.forEach(item => {
    let found = false;
    for (let i = 0; i < rowsV.length; i++) {
      if (rowsV[i][1] === payload.opd && rowsV[i][2].toString() === item.id_soal.toString()) {
        rowsV[i] = [ts, payload.opd, item.id_soal, item.skala_responden, item.skala_evaluator, item.catatan];
        found = true;
        break;
      }
    }
    if (!found) rowsV.push([ts, payload.opd, item.id_soal, item.skala_responden, item.skala_evaluator, item.catatan]);
  });
  
  sheetV.clearContents();
  sheetV.getRange(1, 1, rowsV.length + 1, 6).setValues([headerV, ...rowsV]);
  
  // 2. Batch Update SKM jika ada
  if (payload.nilaiSKM !== undefined) {
    const sheetSKM = ss.getSheetByName("SKM") || ss.insertSheet("SKM");
    if (sheetSKM.getLastRow() === 0) sheetSKM.appendRow(["Timestamp", "OPD", "NilaiSKM"]);
    let dataSKM = sheetSKM.getDataRange().getValues();
    const headerSKM = dataSKM[0];
    let rowsSKM = dataSKM.slice(1);
    
    let foundSKM = false;
    for (let i = 0; i < rowsSKM.length; i++) {
      if (rowsSKM[i][1] === payload.opd) {
        rowsSKM[i] = [ts, payload.opd, payload.nilaiSKM];
        foundSKM = true;
        break;
      }
    }
    if (!foundSKM) rowsSKM.push([ts, payload.opd, payload.nilaiSKM]);
    
    sheetSKM.clearContents();
    sheetSKM.getRange(1, 1, rowsSKM.length + 1, 3).setValues([headerSKM, ...rowsSKM]);
  }

  // 3. Batch Update Catatan Komponen jika ada
  if (payload.catatanKomponen) {
    const sheetCK = ss.getSheetByName("Catatan_Komponen") || ss.insertSheet("Catatan_Komponen");
    if (sheetCK.getLastRow() === 0) sheetCK.appendRow(["Timestamp", "OPD", "PT_Pengamatan", "PT_Rekomendasi", "AI_Pengamatan", "AI_Rekomendasi", "PM_Pengamatan", "PM_Rekomendasi", "EP_Pengamatan", "EP_Rekomendasi"]);
    let dataCK = sheetCK.getDataRange().getValues();
    const headerCK = dataCK[0];
    let rowsCK = dataCK.slice(1);
    
    let foundCK = false;
    const c = payload.catatanKomponen;
    for (let i = 0; i < rowsCK.length; i++) {
      if (rowsCK[i][1] === payload.opd) {
        rowsCK[i] = [ts, payload.opd, c.pt_p, c.pt_r, c.ai_p, c.ai_r, c.pm_p, c.pm_r, c.ep_p, c.ep_r];
        foundCK = true;
        break;
      }
    }
    if (!foundCK) rowsCK.push([ts, payload.opd, c.pt_p, c.pt_r, c.ai_p, c.ai_r, c.pm_p, c.pm_r, c.ep_p, c.ep_r]);
    
    sheetCK.clearContents();
    sheetCK.getRange(1, 1, rowsCK.length + 1, 10).setValues([headerCK, ...rowsCK]);
  }

  // Invalidasi cache rekap agar data terbaru langsung tersedia
  try { CacheService.getScriptCache().remove('rekap_verifikasi'); } catch(e) {}
  try { CacheService.getScriptCache().remove('dashboard_stats'); } catch(e) {}

  return "Sukses";
}

function getStats() {
  if (SETTINGS.USE_FIREBASE) {
    const cache = CacheService.getScriptCache();
    const cachedStats = cache.get("dashboard_stats");
    if (cachedStats) {
      try { return JSON.parse(cachedStats); } catch(e) {}
    }

    const users = Firebase.get("users") || {};
    let totalResp = 0;
    Object.values(users).forEach(u => {
      if (u.role === "Responden") totalResp++;
    });

    const jawaban = Firebase.get("jawaban") || {};
    const sudah = Object.keys(jawaban).length;

    const stats = { total: totalResp, sudah: sudah };
    try { cache.put("dashboard_stats", JSON.stringify(stats), 1800); } catch(e) {}
    return stats;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resps = ss.getSheetByName("Users").getDataRange().getValues().filter(r => r[2] === "Responden").length;
  const sudah = ss.getSheetByName("Jawaban").getLastRow() > 1 ? [...new Set(ss.getSheetByName("Jawaban").getDataRange().getValues().slice(1).map(r => r[1]))].length : 0;
  return { total: resps, sudah: sudah };
}

function getOPDSudahKirim() {
  if (SETTINGS.USE_FIREBASE) {
    const jawaban = Firebase.get("jawaban") || {};
    const uniqueOPD = {};
    Object.keys(jawaban).forEach(opdKey => {
      uniqueOPD[Firebase.unescapeKey(opdKey)] = "Ada";
    });
    return uniqueOPD;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jawaban");
  if (sheet.getLastRow() < 2) return {};
  const data = sheet.getDataRange().getValues().slice(1);
  const uniqueOPD = {};
  data.forEach(r => { uniqueOPD[r[1]] = "Ada"; });
  return uniqueOPD;
}

function getJawabanByOPD(namaOPD) {
  if (SETTINGS.USE_FIREBASE) {
    const opdKey = Firebase.escapeKey(namaOPD);
    
    // Ambil data pararel
    const dj = Firebase.get(`jawaban/${opdKey}`) || {};
    const ds = Firebase.getCachedMasterPertanyaan() || {};
    const dv = Firebase.get(`verifikasi/${opdKey}`) || {};
    const skmObj = Firebase.get(`skm/${opdKey}`) || {};
    const ckObj = Firebase.get(`catatan_komponen/${opdKey}`) || {};

    const items = Object.keys(dj).map(idSoalKey => {
      const j = dj[idSoalKey];
      const soal = ds[idSoalKey] || {};
      const verif = dv[idSoalKey] || {};
      
      return {
        id_soal: Firebase.unescapeKey(idSoalKey),
        pertanyaan: soal.pertanyaan || "N/A",
        skala_responden: j.skala || "",
        link: j.link || "",
        skala_evaluator: verif.skala_evaluator || "",
        catatan: verif.catatan || ""
      };
    });

    return {
      items: items,
      nilaiSKM: skmObj.nilaiSKM || 0,
      catatanKomponen: ckObj.pt_p !== undefined ? ckObj : null
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Baca semua sheet yang diperlukan
  const dj = ss.getSheetByName("Jawaban").getDataRange().getValues().slice(1);
  const ds = ss.getSheetByName("Master_Pertanyaan").getDataRange().getValues().slice(1);
  const dv = ss.getSheetByName("Verifikasi").getLastRow() > 1 ? ss.getSheetByName("Verifikasi").getDataRange().getValues().slice(1) : [];
  const ds_skm = ss.getSheetByName("SKM") && ss.getSheetByName("SKM").getLastRow() > 1 ? ss.getSheetByName("SKM").getDataRange().getValues().slice(1) : [];
  const ds_ck = ss.getSheetByName("Catatan_Komponen") && ss.getSheetByName("Catatan_Komponen").getLastRow() > 1 ? ss.getSheetByName("Catatan_Komponen").getDataRange().getValues().slice(1) : [];

  // Gunakan Map/Object untuk lookup O(1)
  const soalMap = {};
  ds.forEach(s => { soalMap[s[0].toString()] = s; });

  const verifMap = {};
  dv.filter(v => v[1] === namaOPD).forEach(v => { verifMap[v[2].toString()] = v; });

  // Cari SKM dan Catatan Komponen langsung (satu pass)
  const skm = ds_skm.find(s => s[1] === namaOPD);
  const ck  = ds_ck.find(s => s[1] === namaOPD);

  const catatanKomponen = ck ? {
    pt_p: ck[2] || "", pt_r: ck[3] || "",
    ai_p: ck[4] || "", ai_r: ck[5] || "",
    pm_p: ck[6] || "", pm_r: ck[7] || "",
    ep_p: ck[8] || "", ep_r: ck[9] || ""
  } : null;

  return {
    items: dj.filter(r => r[1] === namaOPD).map(j => {
      const soal  = soalMap[j[2].toString()];
      const verif = verifMap[j[2].toString()];
      return {
        id_soal: j[2],
        pertanyaan: soal ? soal[2] : "N/A",
        skala_responden: j[3],
        link: j[4],
        skala_evaluator: verif ? verif[4] : "",
        catatan: verif ? verif[5] : ""
      };
    }),
    nilaiSKM: skm ? skm[2] : 0,
    catatanKomponen: catatanKomponen
  };
}

function getRekapVerifikasi() {
  // Cek cache dulu — jika ada dan belum kedaluwarsa, kembalikan langsung
  const cache = CacheService.getScriptCache();
  const cached = cache.get('rekap_verifikasi');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  let rekap = {};
  let result = [];

  if (SETTINGS.USE_FIREBASE) {
    const dv = Firebase.get("verifikasi") || {};
    const ds = Firebase.getCachedMasterPertanyaan() || {};
    const dSKM = Firebase.get("skm") || {};

    Object.keys(dv).forEach(opdKey => {
      const opd = Firebase.unescapeKey(opdKey);
      const verifOpd = dv[opdKey];
      
      if (!rekap[opd]) rekap[opd] = { opd: opd, rincian: [], nilaiTataKelola: 0, nilaiSKM: 0 };
      
      Object.keys(verifOpd).forEach(idSoalKey => {
        const item = verifOpd[idSoalKey];
        const soal = ds[idSoalKey];
        if (!soal) return;
        
        const skorEval = Number(item.skala_evaluator || 0);
        const bobot = Number(soal.bobot || 0);
        const n = (bobot / 5) * skorEval;
        
        rekap[opd].rincian.push({
          aspek: soal.aspek,
          soal: soal.pertanyaan,
          bobot: bobot,
          skor: skorEval,
          nilai: n.toFixed(3)
        });
        rekap[opd].nilaiTataKelola += n;
      });

      if (dSKM[opdKey]) {
        rekap[opd].nilaiSKM = Number(dSKM[opdKey].nilaiSKM || 0);
      }
    });

    result = Object.values(rekap).map(h => {
      h.nilaiIPPOP = (h.nilaiTataKelola * 0.5) + (h.nilaiSKM * 0.5);
      return h;
    });

  } else {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataV = ss.getSheetByName("Verifikasi").getDataRange().getValues().slice(1);
    const dataS = ss.getSheetByName("Master_Pertanyaan").getDataRange().getValues().slice(1);
    const dataSKM = ss.getSheetByName("SKM") && ss.getSheetByName("SKM").getLastRow() > 1 ? ss.getSheetByName("SKM").getDataRange().getValues().slice(1) : [];

    const soalMap = {};
    dataS.forEach(s => { soalMap[s[0].toString()] = { aspek: s[1], pertanyaan: s[2], bobot: Number(s[4] || 0) }; });

    dataV.forEach(r => {
      const opd = r[1]; const idSoal = r[2].toString(); const skorEval = Number(r[4] || 0); const info = soalMap[idSoal];
      if (!info) return;
      if (!rekap[opd]) rekap[opd] = { opd: opd, rincian: [], nilaiTataKelola: 0, nilaiSKM: 0 };
      const n = (info.bobot / 5) * skorEval;
      rekap[opd].rincian.push({ aspek: info.aspek, soal: info.pertanyaan, bobot: info.bobot, skor: skorEval, nilai: n.toFixed(3) });
      rekap[opd].nilaiTataKelola += n;
    });

    dataSKM.forEach(s => {
      const opd = s[1];
      if (rekap[opd]) rekap[opd].nilaiSKM = Number(s[2] || 0);
    });

    result = Object.values(rekap).map(h => {
      h.nilaiIPPOP = (h.nilaiTataKelola * 0.5) + (h.nilaiSKM * 0.5);
      return h;
    });
  }

  // Simpan ke cache selama 60 detik (1 menit)
  try { cache.put('rekap_verifikasi', JSON.stringify(result), 60); } catch(e) {}

  return result;
}

function hapusJawabanOPD(namaOPD) {
  if (SETTINGS.USE_FIREBASE) {
    const opdKey = Firebase.escapeKey(namaOPD);
    try { Firebase.remove(`jawaban/${opdKey}`); } catch(e) {}
    try { Firebase.remove(`verifikasi/${opdKey}`); } catch(e) {}
    try { Firebase.remove(`skm/${opdKey}`); } catch(e) {}
    try { CacheService.getScriptCache().remove('dashboard_stats'); } catch(e) {}
    return "Sukses";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ["Jawaban", "Verifikasi", "SKM"].forEach(shName => {
    let sh = ss.getSheetByName(shName);
    if (!sh) return;
    let d = sh.getDataRange().getValues();
    if (d.length < 2) return;
    
    const header = d[0];
    const filtered = d.slice(1).filter(row => row[1] !== namaOPD);
    
    sh.clearContents();
    if (filtered.length > 0) {
      sh.getRange(1, 1, filtered.length + 1, header.length).setValues([header, ...filtered]);
    } else {
      sh.getRange(1, 1, 1, header.length).setValues([header]);
    }
  });
  return "Sukses";
}

function getListOPD() {
  if (SETTINGS.USE_FIREBASE) {
    const opds = Firebase.getCachedMasterOPD() || {};
    return Object.keys(opds).map(k => Firebase.unescapeKey(k));
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master_OPD");
  if (!sheet) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(row => row[0]).filter(name => name !== "");
}

/**
 * Generate Berita Acara based on Google Docs Template (Mail Merge)
 */
function generateBeritaAcara(namaOPD) {
  try {
    // 1. Ambil Data Rekap
    const rekapData = getRekapVerifikasi();
    const data = rekapData.find(h => h.opd === namaOPD);
    if (!data) throw new Error("Data rekap untuk OPD ini tidak ditemukan.");
    
    // 2. Ambil Data Catatan Komponen
    let ck = { pt_p:'-', pt_r:'-', ai_p:'-', ai_r:'-', pm_p:'-', pm_r:'-', ep_p:'-', ep_r:'-' };
    
    if (SETTINGS.USE_FIREBASE) {
      const ckObj = Firebase.get(`catatan_komponen/${Firebase.escapeKey(namaOPD)}`);
      if (ckObj) {
        const getVal = (val) => (val !== undefined && val !== null && val.toString().trim() !== "") ? val.toString().trim() : "-";
        ck = {
          pt_p: getVal(ckObj.pt_p), pt_r: getVal(ckObj.pt_r),
          ai_p: getVal(ckObj.ai_p), ai_r: getVal(ckObj.ai_r),
          pm_p: getVal(ckObj.pm_p), pm_r: getVal(ckObj.pm_r),
          ep_p: getVal(ckObj.ep_p), ep_r: getVal(ckObj.ep_r)
        };
      }
    } else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const ds_ck = ss.getSheetByName("Catatan_Komponen") && ss.getSheetByName("Catatan_Komponen").getLastRow() > 1 ? ss.getSheetByName("Catatan_Komponen").getDataRange().getValues().slice(1) : [];
      const ckRow = ds_ck.find(s => s[1].toString().trim() === namaOPD.toString().trim());
      
      const getVal = (val) => (val !== undefined && val !== null && val.toString().trim() !== "") ? val.toString().trim() : "-";
      
      if (ckRow) {
        ck = {
          pt_p: getVal(ckRow[2]), pt_r: getVal(ckRow[3]),
          ai_p: getVal(ckRow[4]), ai_r: getVal(ckRow[5]),
          pm_p: getVal(ckRow[6]), pm_r: getVal(ckRow[7]),
          ep_p: getVal(ckRow[8]), ep_r: getVal(ckRow[9])
        };
      }
    }

    // 3. Construct HTML
    const getPredikatServer = (v) => {
      const n = parseFloat(v);
      if (n >= 4.51) return "PELAYANAN PRIMA (A)";
      if (n >= 4.01) return "SANGAT BAIK (A-)";
      if (n >= 3.51) return "BAIK (B)";
      if (n >= 3.01) return "BAIK DENGAN CATATAN (B-)";
      if (n >= 2.51) return "CUKUP (C)";
      if (n >= 2.01) return "CUKUP DENGAN CATATAN (C-)";
      if (n >= 1.51) return "BURUK (D)";
      if (n >= 1.01) return "SANGAT BURUK (E)";
      return "GAGAL (F)";
    };

    const ippop = parseFloat(data.nilaiIPPOP || 0).toFixed(3);
    const predikat = getPredikatServer(ippop);
    const tanggalCetak = Utilities.formatDate(new Date(), "GMT+7", "dd MMMM yyyy");
    const opdBersih = data.opd.toString().trim();
    const fileName = `Berita_Acara_PEKPPP_${opdBersih.replace(/\s+/g, '_')}_${new Date().getTime()}`;

    // Ambil logo sebagai Base64 agar MS Word tidak memblokir gambar eksternal
    let logoSrc = "https://upload.wikimedia.org/wikipedia/commons/4/47/Lambang_Kabupaten_Muaro_Jambi.png";
    try {
      const imgFetch = UrlFetchApp.fetch(logoSrc, {
        headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }
      });
      logoSrc = "data:image/png;base64," + Utilities.base64Encode(imgFetch.getBlob().getBytes());
    } catch(e) {}

    const htmlContent = `
      <table style="width: 100%; border: none; margin-bottom: 5px;">
        <tr>
          <td style="width: 15%; text-align: center; border: none; vertical-align: middle;">
            <img src="${logoSrc}" width="80" height="100" />
          </td>
          <td style="width: 85%; text-align: center; border: none; vertical-align: middle;">
            <p style="margin: 0; font-weight: bold; font-size: 14pt; line-height: 1.2;">
              PEMERINTAH KABUPATEN MUARO JAMBI<br>SEKRETARIAT DAERAH
            </p>
            <p style="margin: 0; font-size: 10pt; line-height: 1.1; font-weight: normal;">
              Kompleks Perkantoran Bukit Cinto Kenang,<br>
              Jalan Lintas Timur, Sengeti 36381<br>
              Telepon (0741) 590022, 590023; Faksimile. (0741) 590028<br>
              Laman <span style="color: blue;">www.muarojambikab.go.id</span>; pos-el <span style="color: blue;">pemkab@muarojambikab.go.id</span>
            </p>
          </td>
        </tr>
      </table>
      
      <div style="border-bottom: 2pt solid black; width: 100%; margin-top: 5px; margin-bottom: 20px; font-size: 1px; line-height: 1px;">&nbsp;</div>
      
      <div style="text-align: center; margin-bottom: 30px;">
        <p style="margin: 0; font-weight: normal; font-size: 12pt; line-height: 1.1; letter-spacing: 0.5px;">
          BERITA ACARA HASIL PEMANTAUAN DAN EVALUASI PENYELENGGARAAN PELAYANAN PUBLIK (PEKPPP) MANDIRI<br>
          PADA ${opdBersih.toUpperCase()}<br>
          KABUPATEN MUARO JAMBI TAHUN 2026
        </p>
      </div>

      <p style="font-family: 'Times New Roman', Times, serif; text-align: justify; line-height: 1.5;">
        Pada Tanggal ${tanggalCetak} Pemantauan dan Evaluasi Kinerja Penyelenggaraan Pelayanan Publik, berdasarkan hasil pengamatan di lapangan menyatakan sebagai berikut :
      </p>

      <table style="width: 100%; border-collapse: collapse; font-family: 'Times New Roman', Times, serif; margin-bottom: 20px;" border="1" cellpadding="5">
        <thead>
          <tr style="background-color: #f2f2f2;">
            <th style="width: 5%; text-align: center;">No</th>
            <th style="width: 25%; text-align: center;">Komponen</th>
            <th style="width: 35%; text-align: center;">Pengamatan Lapangan</th>
            <th style="width: 35%; text-align: center;">Rekomendasi</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="text-align: center;">1</td>
            <td>PONDASI TEKNIS</td>
            <td style="white-space: pre-wrap;">${ck.pt_p}</td>
            <td style="white-space: pre-wrap;">${ck.pt_r}</td>
          </tr>
          <tr>
            <td style="text-align: center;">2</td>
            <td>AKSESIBILITAS DAN INKLUSIF</td>
            <td style="white-space: pre-wrap;">${ck.ai_p}</td>
            <td style="white-space: pre-wrap;">${ck.ai_r}</td>
          </tr>
          <tr>
            <td style="text-align: center;">3</td>
            <td>PELIBATAN MASYARAKAT</td>
            <td style="white-space: pre-wrap;">${ck.pm_p}</td>
            <td style="white-space: pre-wrap;">${ck.pm_r}</td>
          </tr>
          <tr>
            <td style="text-align: center;">4</td>
            <td>EFEKTIVITAS PEMERINTAHAN</td>
            <td style="white-space: pre-wrap;">${ck.ep_p}</td>
            <td style="white-space: pre-wrap;">${ck.ep_r}</td>
          </tr>
        </tbody>
      </table>

      <p style="font-family: 'Times New Roman', Times, serif; text-align: justify; margin-bottom: 30px;">
        Demikian Berita Acara ini dibuat sebagaimana mestinya.
      </p>

      <!-- Page Break untuk Tanda Tangan -->
      <br clear="all" style="page-break-before: always" />

      <div style="font-family: 'Times New Roman', Times, serif; width: 100%; margin-top: 30px;">
        <p style="text-align: center; font-weight: bold; margin-bottom: 30px;">Tim Evaluator</p>
        
        <table style="width: 100%; border: none; text-align: center; margin-bottom: 50px;">
          <tr>
            <td style="width: 50%; border: none;">
              Evaluator 1<br><br><br><br><br>
              <span style="text-decoration: underline;">(...............................................)</span>
            </td>
            <td style="width: 50%; border: none;">
              Evaluator 2<br><br><br><br><br>
              <span style="text-decoration: underline;">(...............................................)</span>
            </td>
          </tr>
        </table>

        <p style="text-align: center; font-weight: bold; margin-bottom: 30px;">Mengetahui,</p>
        
        <table style="width: 100%; border: none; text-align: center;">
          <tr>
            <td style="width: 50%; border: none;">
              Ketua Evaluator<br>
              Kepala Bagian Organisasi<br><br><br><br><br>
              <span style="text-decoration: underline;">(...............................................)</span><br>
              NIP. ........................................
            </td>
            <td style="width: 50%; border: none;">
              Kepala ........................................<br><br><br><br><br><br>
              <span style="text-decoration: underline;">(...............................................)</span><br>
              NIP. ........................................
            </td>
          </tr>
        </table>
      </div>
    `;

    return {
      status: "success",
      html: htmlContent,
      fileName: fileName + ".doc"
    };

  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * Simpan atau update pengaturan periode pengisian.
 * @param {{ tipe: 'GLOBAL'|string, tglBuka: string, tglTutup: string }} payload
 */
function simpanPeriodePengisian(payload) {
  const tipe = payload.tipe.trim().toUpperCase() === "GLOBAL" ? "GLOBAL" : payload.tipe.trim();
  const tglBuka = new Date(payload.tglBuka);
  const tglTutup = new Date(payload.tglTutup);

  if (tglTutup <= tglBuka) throw new Error("Tanggal tutup harus lebih besar dari tanggal buka.");

  if (SETTINGS.USE_FIREBASE) {
    const tipeKey = Firebase.escapeKey(tipe.toUpperCase());
    Firebase.put(`pengaturan/${tipeKey}`, {
      tipe: tipe.toUpperCase(),
      tglBuka: tglBuka.toISOString(),
      tglTutup: tglTutup.toISOString()
    });
    return { status: "success", pesan: `Pengaturan periode untuk "${tipe}" berhasil disimpan di Firebase.` };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Pengaturan");

  if (!sheet) {
    sheet = ss.insertSheet("Pengaturan");
    sheet.getRange(1, 1, 1, 3).setValues([["OPD/Scope", "Tanggal Buka", "Tanggal Tutup"]]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  }

  const data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues() : [];
  let found = false;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === tipe) {
      sheet.getRange(i + 2, 1, 1, 3).setValues([[tipe, tglBuka, tglTutup]]);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([tipe, tglBuka, tglTutup]);
  }

  return { status: "success", pesan: `Pengaturan periode untuk "${tipe}" berhasil disimpan.` };
}

/**
 * Ambil semua pengaturan periode yang ada di sheet Pengaturan.
 */
function getPengaturanPeriode() {
  const fmt = (d) => {
    if (!d || d === "") return "";
    try {
      return Utilities.formatDate(new Date(d), "GMT+7", "yyyy-MM-dd'T'HH:mm");
    } catch(e) { return ""; }
  };

  if (SETTINGS.USE_FIREBASE) {
    const p = Firebase.get("pengaturan") || {};
    return Object.values(p).map(r => ({
      tipe: r.tipe,
      tglBuka: fmt(r.tglBuka),
      tglTutup: fmt(r.tglTutup)
    }));
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pengaturan");
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter(r => r[0] !== "")
    .map(r => ({
      tipe: r[0].toString().trim(),
      tglBuka: fmt(r[1]),
      tglTutup: fmt(r[2])
    }));
}

/**
 * Hapus satu baris pengaturan periode berdasarkan tipe/nama OPD.
 */
function hapusPeriodePengisian(tipe) {
  if (SETTINGS.USE_FIREBASE) {
    const tipeKey = Firebase.escapeKey(tipe.trim().toUpperCase());
    try {
      Firebase.remove(`pengaturan/${tipeKey}`);
      return { status: "success" };
    } catch(e) {
      return { status: "error", message: e.message };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pengaturan");
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === tipe.toString().trim().toUpperCase()) {
      sheet.deleteRow(i + 2);
      return { status: "success" };
    }
  }
  return { status: "notfound" };
}

/**
 * Optimasi: Gabungkan dua panggilan server menjadi satu.
 * Mengembalikan daftar OPD dari Master_OPD + daftar periode pengaturan sekaligus.
 */
function getPengaturanDanOPD() {
  return {
    opdList: getListOPD(),
    periodeList: getPengaturanPeriode()
  };
}
