/**
 * Firebase Realtime Database REST API Client for Google Apps Script
 * Dilengkapi dengan sistem escaping kunci dan in-memory caching untuk optimasi performa.
 */
const Firebase = {
  /**
   * Mengamankan kunci database dari karakter terlarang Firebase (. # $ [ ] /)
   */
  escapeKey: function(key) {
    if (!key) return "";
    return key.toString()
              .replace(/%/g, '%25')
              .replace(/\./g, '%2E')
              .replace(/#/g, '%23')
              .replace(/\$/g, '%24')
              .replace(/\[/g, '%5B')
              .replace(/\]/g, '%5D')
              .replace(/\//g, '%2F');
  },

  /**
   * Mengembalikan kunci yang sudah diamankan kembali ke bentuk aslinya
   */
  unescapeKey: function(key) {
    if (!key) return "";
    return key.toString()
              .replace(/%2E/g, '.')
              .replace(/%23/g, '#')
              .replace(/%24/g, '$')
              .replace(/%5B/g, '[')
              .replace(/%5D/g, ']')
              .replace(/%2F/g, '/')
              .replace(/%25/g, '%');
  },

  getDbUrl: function() {
    const prop = SETTINGS.FIREBASE_DB_URL;
    if (prop) return prop.replace(/\/$/, ""); 
    return "";
  },

  getSecret: function() {
    return SETTINGS.FIREBASE_SECRET || "";
  },

  buildUrl: function(path) {
    const baseUrl = this.getDbUrl();
    const secret = this.getSecret();
    if (!baseUrl || !secret) {
      throw new Error("Firebase DB URL atau Secret belum diatur di GlobalSettings.");
    }
    const cleanPath = path.startsWith("/") ? path : "/" + path;
    return `${baseUrl}${cleanPath}.json?auth=${secret}`;
  },

  /**
   * Mengambil data (GET)
   */
  get: function(path) {
    const url = this.buildUrl(path);
    const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      throw new Error(`Firebase GET Error [${response.getResponseCode()}]: ${response.getContentText()}`);
    }
    const data = JSON.parse(response.getContentText());
    return data;
  },

  /**
   * Menimpa data secara keseluruhan (PUT)
   */
  put: function(path, data) {
    const url = this.buildUrl(path);
    const options = {
      method: "put",
      contentType: "application/json",
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Firebase PUT Error [${response.getResponseCode()}]: ${response.getContentText()}`);
    }
    return JSON.parse(response.getContentText());
  },

  /**
   * Memperbarui sebagian data (PATCH)
   */
  patch: function(path, data) {
    const url = this.buildUrl(path);
    const options = {
      method: "patch",
      contentType: "application/json",
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Firebase PATCH Error [${response.getResponseCode()}]: ${response.getContentText()}`);
    }
    return JSON.parse(response.getContentText());
  },

  /**
   * Menambahkan data baru dengan ID unik buatan Firebase (POST)
   */
  post: function(path, data) {
    const url = this.buildUrl(path);
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Firebase POST Error [${response.getResponseCode()}]: ${response.getContentText()}`);
    }
    return JSON.parse(response.getContentText()); // Return objek { "name": "-Nid_generated" }
  },

  /**
   * Menghapus data (DELETE)
   */
  remove: function(path) {
    const url = this.buildUrl(path);
    const options = {
      method: "delete",
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Firebase DELETE Error [${response.getResponseCode()}]: ${response.getContentText()}`);
    }
    return true;
  },

  /**
   * Helper Caching: Mengambil daftar pertanyaan dari Cache atau Firebase
   */
  getCachedMasterPertanyaan: function() {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get("master_pertanyaan");
    if (cachedData) {
      try {
        return JSON.parse(cachedData);
      } catch(e) {}
    }
    const pert = this.get("master_pertanyaan");
    if (pert) {
      cache.put("master_pertanyaan", JSON.stringify(pert), 1800); // 30 menit
    }
    return pert;
  },

  /**
   * Helper Caching: Mengambil daftar master OPD dari Cache atau Firebase
   */
  getCachedMasterOPD: function() {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get("master_opd");
    if (cachedData) {
      try {
        return JSON.parse(cachedData);
      } catch(e) {}
    }
    const opds = this.get("master_opd");
    if (opds) {
      cache.put("master_opd", JSON.stringify(opds), 1800); // 30 menit
    }
    return opds;
  }
};
