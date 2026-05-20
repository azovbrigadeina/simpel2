/**
 * Auth & User Management Module
 */

function prosesLogin(username, password) {
  const u = username.trim(); const p = password.trim();

  if (SETTINGS.USE_FIREBASE) {
    try {
      const escapedU = Firebase.escapeKey(u);
      const userData = Firebase.get(`users/${escapedU}`);
      if (userData && userData.password === p) {
        let sudahIsi = false;
        if (userData.role === "Responden") {
          const jawabanData = Firebase.get(`jawaban/${Firebase.escapeKey(userData.nama_opd)}`);
          sudahIsi = (jawabanData !== null);
        }
        return { 
          status: "success", 
          role: userData.role, 
          nama_opd: userData.nama_opd, 
          username: u,
          sudahIsi: sudahIsi
        };
      }
      return { status: "error", message: "Username atau Password Salah!" };
    } catch(e) {
      return { status: "error", message: "Gagal login melalui Firebase: " + e.message };
    }
  }

  // Fallback Google Sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const data = userSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === u && data[i][1].toString() === p) {
      const role = data[i][2];
      const nama_opd = data[i][3];
      
      let sudahIsi = false;
      if (role === "Responden") {
        const jawabanSheet = ss.getSheetByName("Jawaban");
        const dataJawaban = jawabanSheet.getDataRange().getValues();
        sudahIsi = dataJawaban.some(row => row[1] === nama_opd);
      }

      return { 
        status: "success", 
        role: role, 
        nama_opd: nama_opd, 
        username: data[i][0],
        sudahIsi: sudahIsi
      };
    }
  }
  return { status: "error", message: "Username atau Password Salah!" };
}

function getAllUsers() {
  if (SETTINGS.USE_FIREBASE) {
    const users = Firebase.get("users") || {};
    return Object.keys(users).map(k => [
      Firebase.unescapeKey(k),
      users[k].password,
      users[k].role,
      users[k].nama_opd
    ]);
  }

  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").getDataRange().getValues().slice(1);
}

/**
 * Mengembalikan daftar user sebagai array objek terstruktur.
 * Digunakan oleh tab Pengaturan untuk mengisi dropdown pilihan OPD.
 */
function getDaftarUser() {
  if (SETTINGS.USE_FIREBASE) {
    const users = Firebase.get("users") || {};
    return Object.keys(users).map(k => ({
      username: Firebase.unescapeKey(k).trim(),
      role: (users[k].role || "").trim().toLowerCase(),
      opd: (users[k].nama_opd || "").trim()
    }));
  }

  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").getDataRange().getValues().slice(1);
  return rows.map(r => ({
    username: r[0].toString().trim(),
    role: r[2].toString().trim().toLowerCase(),
    opd: r[3].toString().trim()
  }));
}

function simpanUserBaru(payload) {
  if (SETTINGS.USE_FIREBASE) {
    const escapedU = Firebase.escapeKey(payload.username);
    const existing = Firebase.get(`users/${escapedU}`);
    if (existing) {
      throw new Error("Username sudah terdaftar!");
    }
    Firebase.put(`users/${escapedU}`, {
      password: payload.password,
      role: payload.role,
      nama_opd: payload.nama_opd
    });
    return "User Berhasil Ditambahkan";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const usernameExist = data.some(row => row[0].toString() === payload.username.toString());
  
  if (usernameExist) {
    throw new Error("Username sudah terdaftar!");
  }
  
  sheet.appendRow([
    payload.username, 
    payload.password, 
    payload.role, 
    payload.nama_opd
  ]);
  
  return "User Berhasil Ditambahkan";
}

function hapusUser(username) {
  if (SETTINGS.USE_FIREBASE) {
    const escapedU = Firebase.escapeKey(username);
    const existing = Firebase.get(`users/${escapedU}`);
    if (!existing) {
      throw new Error("User tidak ditemukan");
    }
    Firebase.remove(`users/${escapedU}`);
    return "User berhasil dihapus";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === username.toString()) {
      sheet.deleteRow(i + 1);
      return "User berhasil dihapus";
    }
  }
  throw new Error("User tidak ditemukan");
}
