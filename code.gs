/**
 * OWNER COMMAND CENTER - BPR KS
 * Versi: 48.0 (Ultra Performance Engine - OPTIMIZED)
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Executive Dashboard - BPR KS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- OPTIMASI 1: CACHING IDENTITY ---
function getUserData() {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  var cache = CacheService.getUserCache();
  var cached = cache.get("user_profile");
  
  if (cached) return JSON.parse(cached);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('User_ID') || ss.insertSheet('User_ID');
  if (sheet.getLastRow() === 0) sheet.appendRow(["Email", "Nama", "Tanggal", "Role", "Cabang"]);

  var data = sheet.getDataRange().getValues();
  var user = { email: userEmail, nama: "GUEST", role: "Admin", cabang: "ALL", ok: true };

  // Hard-lock Owner
  if (userEmail.includes("yulius") || userEmail === "yulius.puji.laksono@gmail.com") {
    user = { email: userEmail, nama: "YULIUS PUJI LAKSONO", role: "Admin", cabang: "ALL", ok: true };
  } else {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toLowerCase() === userEmail) {
        user = { email: data[i][0], nama: data[i][1], role: data[i][3], cabang: data[i][4], ok: true };
        break;
      }
    }
  }

  // Simpan di cache selama 20 menit
  cache.put("user_profile", JSON.stringify(user), 1200);
  return user;
}

// --- OPTIMASI 2: CACHING & DATA COMPRESSION ---
function getInitialData() {
  var user = getUserData();
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get("global_records_v48");
  
  var rawArray;
  if (cachedData) {
    rawArray = JSON.parse(cachedData);
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    rawArray = [];
    
    // Batch Reading 3 Kamar sekaligus
    ["Raw_Achv_CS", "Raw_Achv_SPV", "Raw_Achv_Tele"].forEach(shName => {
      var sh = ss.getSheetByName(shName);
      if (!sh || sh.getLastRow() < 2) return;
      var data = sh.getDataRange().getValues();
      var key = shName.split('_')[2].toLowerCase();

      for (var i = 1; i < data.length; i++) {
        var row = data[i]; if (!row[0]) continue;
        var tgl = (row[0] instanceof Date) ? Utilities.formatDate(row[0], "GMT+7", "yyyy-MM-dd") : row[0].toString().substring(0, 10);
        
        // Simpan dalam format Array murni (Hemat Karakter JSON)
        // [Tanggal, Cabang, Nama, Jabatan, Target, Realisasi]
        rawArray.push([tgl, row[2], row[3], key, Number(row[4])||0, Number(row[6])||0]);
      }
    });
    
    // Simpan di Script Cache (Global) selama 10 menit
    cache.put("global_records_v48", JSON.stringify(rawArray), 600);
  }

  // Filter Keamanan dilakukan setelah data ditarik (Server-side filtering cepat)
  var filtered = rawArray.filter(r => {
    if (user.role === "Staff") return r[2] === user.nama;
    if (user.role === "BM") return r[1] === user.cabang;
    if (user.role === "RH") return user.cabang.split(',').indexOf(r[1]) !== -1;
    return true; // Admin/Owner
  });

  return { u: user, d: filtered };
}

// --- OPTIMASI 3: LOCK SERVICE SAAT UPLOAD ---
function processExcelData(rows, tgl, tipe) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Tunggu 30 detik jika ada proses lain
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mapping = { "1":"Raw_Achv_CS", "2":"Raw_Achv_SPV", "3":"Raw_Achv_Tele" };
    var sheet = ss.getSheetByName(mapping[tipe]) || ss.insertSheet(mapping[tipe]);
    
    var finalData = [];
    var clean = function(v) { 
      if (!v) return 0;
      var n = Number(v.toString().replace(/[^0-9.-]+/g, ""));
      return isNaN(n) ? 0 : n;
    };

    if (tipe === "1" || tipe === "2") {
      var it = (tipe === "1") ? 5 : 4; var ir = (tipe === "1") ? 8 : 7;
      for (var i = 2; i < rows.length; i++) {
        var r = rows[i]; if (!r[1] || r[1].toString().includes("TOTAL")) continue;
        finalData.push([tgl, r[1], r[2], r[3], clean(r[it]), clean(r[it+1]), clean(r[ir]), clean(r[ir+1])]);
      }
    } else if (tipe === "3") {
      for (var i = 3; i < rows.length; i++) {
        var r = rows[i]; if (!r[1] || r[1].toString().includes("TOTAL") || r[1] === "") continue;
        finalData.push([tgl, "000", "TELEMARKETING", r[1], clean(r[2]), clean(r[3]), clean(r[5]), clean(r[6])]);
      }
    }

    if (finalData.length > 0) {
      // Pembersihan data harian yang ada
      var old = sheet.getDataRange().getValues();
      for(var j=old.length-1; j>=1; j--) {
        var dStr = (old[j][0] instanceof Date) ? Utilities.formatDate(old[j][0], "GMT+7", "yyyy-MM-dd") : old[j][0].toString().substring(0,10);
        if(dStr === tgl) sheet.deleteRow(j+1);
      }
      sheet.getRange(sheet.getLastRow()+1, 1, finalData.length, finalData[0].length).setValues(finalData);
      
      // OPTIMASI: Clear Cache setelah upload agar data baru muncul
      CacheService.getScriptCache().remove("global_records_v48");
      return "✅ Sukses.";
    }
    return "Gagal.";
  } finally {
    lock.releaseLock();
  }
}

function getMenu() {
  return [{id:'dashboard', label:'Dashboard', icon:'bi-speedometer2'}, {id:'upload', label:'Admin', icon:'bi-shield-lock-fill'}];
}
