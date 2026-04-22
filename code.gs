/**
 * OWNER COMMAND CENTER - BPR KS
 * Versi: 97.0 (PIPELINE ADD-ON - STRICT ADHERENCE)
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Executive Dashboard - BPR KS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= USER IDENTITY (Tetap) ================= */
function getUserData() {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('User_ID') || ss.insertSheet('User_ID');
    var data = sh.getDataRange().getValues();
    if (email.includes("yulius") || email === "") {
      return { email: email, nama: "YULIUS PUJI LAKSONO", role: "Admin", cabang: "ALL" };
    }
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toLowerCase() === email) {
        return { email: data[i][0], nama: data[i][1], role: data[i][3], cabang: data[i][4] };
      }
    }
    return { email: email, nama: "GUEST", role: "Admin", cabang: "ALL" };
  } catch(e) { return { nama: "USER", role: "Admin", cabang: "ALL" }; }
}

/* ================= SMART CACHE ENGINE (Tetap) ================= */
function getCachedData(key, fn) {
  var cache = CacheService.getScriptCache();
  var data = cache.get(key);
  if (data) return JSON.parse(data);
  var fresh = fn();
  try { cache.put(key, JSON.stringify(fresh), 300); } catch(e) {}
  return fresh;
}

/* ================= DATA PROVIDER (Ditambah Pipeline) ================= */
function getDashboardData() {
  var user = getUserData();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  return getCachedData("DB_FULL_" + user.email, function() {
    var results = { u: user, achv: [], cair: [], listC: [], listN: [], pipeline: [], listS: [], listSPV: [] };
    
    // 1. ACHIEVEMENT (Existing)
    ["Raw_Achv_CS", "Raw_Achv_SPV", "Raw_Achv_Tele"].forEach(name => {
      var sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return;
      var data = sh.getDataRange().getValues();
      var key = name.split('_')[2].toLowerCase();
      data.slice(1).forEach(r => {
        if(!r[0]) return;
        var tgl = (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : r[0].toString().substring(0,10);
        if (r[2] && results.listC.indexOf(r[2]) === -1) results.listC.push(r[2]);
        if (r[3] && results.listN.indexOf(r[3]) === -1) results.listN.push(r[3]);
        results.achv.push([tgl, r[2], r[3], key, Number(r[4])||0, Number(r[6])||0]);
      });
    });

    // 2. PENCAIRAN (Existing - Fixed Mapping KBMBL)
    ["Raw_Cair_CS", "Raw_Cair_SPV", "Raw_Cair_Tele"].forEach(name => {
      var sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return;
      var data = sh.getDataRange().getValues();
      var isT = name.includes("Tele");
      data.slice(1).forEach(r => {
        if(!r[0]) return;
        var tgl = (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : r[0].toString().substring(0,10);
        var rCab = r[3]; var rNam = isT ? "TELEMARKETING" : (r[4]||"").toString().trim();
        var items = [];
        if (name.includes("SPV")) {
          if(Number(r[8]) > 0) items.push(["KABHT", r[8], r[9]]);
          if(Number(r[10]) > 0) items.push(["KAB", r[10], r[11]]);
          if(Number(r[12]) > 0) items.push(["KPLM", r[12], r[14]]);
          if(Number(r[16]) > 0) items.push(["KABM", r[16], r[17]]);
          if(Number(r[18]) > 0) items.push(["KPSM", r[18], r[19]]);
          if(Number(r[21]) > 0) items.push(["KBMBL", r[22], r[21]]);
          if(Number(r[23]) > 0) items.push(["KABEKS", r[24], r[23]]);
        } else if (isT) {
          if(Number(r[7]) > 0) items.push(["KABHT", r[7], r[8]]);
          if(Number(r[9]) > 0) items.push(["KAB", r[9], r[10]]);
          if(Number(r[13]) > 0) items.push(["KPLM", r[13], r[14]]);
          if(Number(r[15]) > 0) items.push(["KPSM", r[15], r[16]]);
          if(Number(r[17]) > 0) items.push(["KBMBL", r[18], r[17]]);
          if(Number(r[19]) > 0) items.push(["KABM", r[19], r[20]]);
          if(Number(r[21]) > 0) items.push(["KABEKS", r[22], r[21]]);
        }
        items.forEach(it => { results.cair.push([tgl, rCab, rNam, it[0], it[1], it[2]]); });
      });
    });

    // 3. PIPELINE (Additive Logic)
    var shP = ss.getSheetByName('Raw_Pipeline');
    if (shP && shP.getLastRow() >= 2) {
      var dataP = shP.getDataRange().getValues();
      dataP.slice(1).forEach(r => {
        if(!r[1]) return;
        var tglP = (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : String(r[0]).substring(0,10);
        if (r[1] && results.listC.indexOf(String(r[1])) === -1) results.listC.push(String(r[1]));
        if (r[5] && results.listS.indexOf(String(r[5])) === -1) results.listS.push(String(r[5]));
        if (r[6] && results.listSPV.indexOf(String(r[6])) === -1) results.listSPV.push(String(r[6]));
        results.pipeline.push({ cab:r[1], tgl:tglP, st:r[2], kep:r[4], sal:r[5], spv:r[6], deb:r[7], pla:Number(r[8])||0 });
      });
    }
    return results;
  });
}

/* ================= UPLOAD SYSTEM (Additive) ================= */
function processExcelData(rows, tgl, tipe) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mapping = { "1":"Raw_Achv_CS", "2":"Raw_Achv_SPV", "3":"Raw_Achv_Tele", "4":"Raw_Cair_CS", "5":"Raw_Cair_SPV", "6":"Raw_Cair_Tele", "7":"Raw_Pipeline" };
    var sheet = ss.getSheetByName(mapping[tipe]) || ss.insertSheet(mapping[tipe]);
    var finalData = [];
    var clean = v => parseFloat((v||"").toString().replace(/[^0-9.-]+/g,""))||0;

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i]; if (!r[1] || r[1].toString().includes("TOTAL")) continue;
      var rowToS = [];
      if (tipe == "7") {
        for (var c = 0; c < r.length; c++) { rowToS.push(c == 8 ? clean(r[c]) : r[c]); }
      } else {
        rowToS = [tgl];
        for (var c = 0; c < r.length; c++) { rowToS.push(c >= 4 ? clean(r[c]) : r[c]); }
      }
      finalData.push(rowToS);
    }

    if (finalData.length > 0) {
      var old = sheet.getDataRange().getValues();
      for(var j=old.length-1; j>=1; j--) {
        var dStr = (old[j][0] instanceof Date) ? Utilities.formatDate(old[j][0], "GMT+7", "yyyy-MM-dd") : String(old[j][0]).substring(0,10);
        if(dStr === tgl) sheet.deleteRow(j+1);
      }
      sheet.getRange(sheet.getLastRow()+1, 1, finalData.length, finalData[0].length).setValues(finalData);
      CacheService.getScriptCache().remove("DB_FULL_" + Session.getActiveUser().getEmail().toLowerCase());
      return "✅ Berhasil Disimpan.";
    }
    return "⚠️ Data Kosong.";
  } finally { lock.releaseLock(); }
}
