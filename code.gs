/**
 * OWNER COMMAND CENTER - BPR KS
 * Versi: 70.0 (High Visibility & Detail UI - FINAL)
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Executive Dashboard - BPR KS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserData() {
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  var cache = CacheService.getUserCache();
  var cached = cache.get("user_profile");
  if (cached) return JSON.parse(cached);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('User_ID') || ss.insertSheet('User_ID');
  if (sheet.getLastRow() === 0) sheet.appendRow(["Email", "Nama", "Tanggal", "Role", "Cabang"]);

  var data = sheet.getDataRange().getValues();
  var user = { email: userEmail, nama: "YULIUS PUJI LAKSONO", role: "Admin", cabang: "ALL", ok: true };

  if (!userEmail.includes("yulius") && userEmail !== "") {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toLowerCase() === userEmail) {
        user = { email: data[i][0], nama: data[i][1], role: data[i][3], cabang: data[i][4], ok: true };
        break;
      }
    }
  }
  cache.put("user_profile", JSON.stringify(user), 1200);
  return user;
}

function getInitialData() {
  var user = getUserData();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = { u: user, achv: [], cair: [] };

  try {
    // 1. DATA ACHIEVEMENT
    ["Raw_Achv_CS", "Raw_Achv_SPV", "Raw_Achv_Tele"].forEach(shName => {
      var sh = ss.getSheetByName(shName);
      if (!sh || sh.getLastRow() < 2) return;
      var data = sh.getDataRange().getValues();
      var key = shName.split('_')[2].toLowerCase();

      for (var i = 1; i < data.length; i++) {
        var row = data[i]; if (!row[0]) continue;
        var tgl = (row[0] instanceof Date) ? Utilities.formatDate(row[0], "GMT+7", "yyyy-MM-dd") : row[0].toString().substring(0, 10);
        if (user.role === "Staff" && row[3] !== user.nama) continue;
        if (user.role === "BM" && row[2] !== user.cabang) continue;
        results.achv.push([tgl, row[2], row[3], key, Number(row[4])||0, Number(row[6])||0]);
      }
    });

    // 2. DATA PENCAIRAN
    ["Raw_Cair_CS", "Raw_Cair_SPV", "Raw_Cair_Tele"].forEach(shName => {
      var sh = ss.getSheetByName(shName);
      if (!sh || sh.getLastRow() < 2) return;
      var data = sh.getDataRange().getValues();
      var isT = shName.includes("Tele");

      for (var i = 1; i < data.length; i++) {
        var r = data[i]; if (!r[0]) continue;
        var tgl = (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd") : r[0].toString().substring(0, 10);
        var rCab = r[3]; var rNam = isT ? "TELEMARKETING" : (r[4]||"").toString().trim();
        
        if (user.role === "Staff" && rNam !== user.nama && rNam !== "TELEMARKETING") continue;
        if (user.role === "BM" && rCab !== user.cabang) continue;

        var items = [];
        if (shName.includes("SPV")) {
          if(Number(r[10]) > 0) items.push(["KAB", r[10], r[11]]);
          if(Number(r[12]) > 0) items.push(["KPLM", r[12], r[14]]);
          if(Number(r[16]) > 0) items.push(["KABM", r[16], r[17]]);
          if(Number(r[18]) > 0) items.push(["KPSM", r[18], r[19]]);
        } else if (isT) {
          if(Number(r[9]) > 0) items.push(["KAB", r[9], r[10]]);
          if(Number(r[13]) > 0) items.push(["KPLM", r[13], r[14]]);
          if(Number(r[15]) > 0) items.push(["KPSM", r[15], r[16]]);
          if(Number(r[19]) > 0) items.push(["KABM", r[19], r[20]]);
        }
        items.forEach(it => { results.cair.push([tgl, rCab, rNam, it[0], Number(it[1]), Number(it[2])]); });
      }
    });
  } catch(e) { results.error = e.message; }
  return results;
}

function processExcelData(rows, tgl, tipe) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mapping = { "1":"Raw_Achv_CS", "2":"Raw_Achv_SPV", "3":"Raw_Achv_Tele", "4":"Raw_Cair_CS", "5":"Raw_Cair_SPV", "6":"Raw_Cair_Tele" };
    var sheet = ss.getSheetByName(mapping[tipe]) || ss.insertSheet(mapping[tipe]);
    var finalData = [];
    var clean = v => { 
      if(!v || v==="-") return 0; 
      var n = v.toString().replace(/,/g, "").replace(/"/g, "").replace(/[^0-9.-]+/g, "");
      return parseFloat(n) || 0; 
    };
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i]; if (!r[1] || r[1].toString().includes("TOTAL")) continue;
      var rowToS = [tgl];
      for (var c = 0; c < r.length; c++) { rowToS.push(c >= 4 ? clean(r[c]) : r[c]); }
      finalData.push(rowToS);
    }
    if (finalData.length > 0) {
      var old = sheet.getDataRange().getValues();
      for(var j=old.length-1; j>=1; j--) {
        var dStr = (old[j][0] instanceof Date) ? Utilities.formatDate(old[j][0], "GMT+7", "yyyy-MM-dd") : old[j][0].toString().substring(0,10);
        if(dStr === tgl) sheet.deleteRow(j+1);
      }
      sheet.getRange(sheet.getLastRow()+1, 1, finalData.length, finalData[0].length).setValues(finalData);
      return "✅ Sukses.";
    }
    return "Gagal.";
  } finally { lock.releaseLock(); }
}
