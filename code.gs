/**
 * COACHING MONITORING SYSTEM v8.0 + PENCAPAIAN MODULE
 * Full Package - Backend Google Apps Script
 * 
 * Fitur:
 * - Coaching Module (existing)
 * - Pencapaian/Approve BLM Cair Module (NEW)
 * - Mobile-first, PermataBank-inspired UI
 */

//==================== KONSTANTA ====================
const SHEET_KARYAWAN = 'DB_KARYAWAN';
const SHEET_USERS = 'DB_USERS';
const SHEET_COACHING_HEADER = 'COACHING_HEADER';
const SHEET_COACHING_DETAIL = 'COACHING_DETAIL';
const SHEET_REMINDER_LOG = 'REMINDER_LOG';
const SHEET_APPROVE_BLM_CAIR = 'DB_APPROVE_BLM_CAIR';

const ALLOWED_UPLOAD_NPK = '2510285';
const ALLOWED_UPLOAD_JABATAN = ['Konsultan', 'HR', 'PORTFOLIO PERFORMANCE AND DATA ANALYTICS OFFICER'];

//==================== HELPER: SETUP SHEETS ====================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheets = [
    { name: SHEET_KARYAWAN, headers: ['NPK', 'Nama', 'Jabatan', 'CL', 'Cabang', 'Region', 'Atasan_NPK', 'Email', 'No_HP'] },
    { name: SHEET_USERS, headers: ['NPK', 'Nama', 'Role', 'Status', 'Created_At', 'Updated_At', 'Created_By'] },
    { name: SHEET_COACHING_HEADER, headers: ['coaching_id', 'coach_npk', 'coachee_npk', 'cabang', 'root_cause', 'topic', 'status', 'created_date', 'target_date'] },
    { name: SHEET_COACHING_DETAIL, headers: ['detail_id', 'coaching_id', 'week', 'action', 'how', 'target_date', 'result', 'feedback', 'update_date', 'target_nominal', 'target_satuan'] },
    { name: SHEET_REMINDER_LOG, headers: ['reminder_id', 'coaching_id', 'coach_npk', 'coachee_npk', 'reminder_date', 'channel', 'status', 'sent_timestamp', 'message_content', 'created_at'] }
  ];
  
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.getRange(1, 1, 1, s.headers.length).setValues([s.headers]);
      sheet.getRange(1, 1, 1, s.headers.length)
        .setFontWeight('bold')
        .setBackground('#1e3a5f')
        .setFontColor('#ffffff');
    }
  }
  
  setupApproveBlmCairSheet();
  return true;
}

function setupApproveBlmCairSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_APPROVE_BLM_CAIR);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_APPROVE_BLM_CAIR);
      var headers = [
        'ID', 'TGL', 'NAMA', 'CABANG', 'SUMBER', 'STATUS_APLIKASI', 
        'JENIS_KREDIT', 'KEPUTUSAN', 'NIK', 'NAMA_SALES', 'NAMA_SPV',
        'NAMA_SALES_QA', 'NAMA_SALES_QA_BMR', 'NAMA_SPV_QA', 'NAMA_SPV_QA_BMR',
        'STATUS_NASABAH', 'FPK_AWAL', 'FPK_CABANG', 'NAMA_SCORING',
        'PLAFOND', 'NOMINAL_ACC', 'BUNGA', 'KATEGORI', 'KATEGORI_BUNGA2',
        'MERK', 'TYPE', 'TAHUN', 'JARAK', 'AN_JAMINAN', 'NO_HP',
        'NO_LOAN', 'TGL_HANDOFF', 'KETERANGAN_SALES', 'UPLOAD_DATE', 'UPLOAD_BY'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1e3a5f')
        .setFontColor('#ffffff');
    }
    return true;
  } catch (e) {
    console.error('Error setupApproveBlmCairSheet:', e);
    return false;
  }
}

//==================== HELPER: CEK UPLOAD PERMISSION ====================
function canUploadData(userNpk, userJabatan) {
  if (userNpk === ALLOWED_UPLOAD_NPK) return true;
  if (ALLOWED_UPLOAD_JABATAN.includes(userJabatan)) return true;
  return false;
}

//==================== HELPER: GET CABANG LIST ====================
function getCabangList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var cabangSet = new Set();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][3]) cabangSet.add(data[i][3].toString().trim());
    }
    
    return Array.from(cabangSet).sort();
  } catch (e) {
    console.error('Error getCabangList:', e);
    return [];
  }
}

//==================== HELPER: PARSE DATE ====================
function parseDate(value) {
  if (!value) return '';
  if (value instanceof Date) return value;
  
  if (typeof value === 'number' && value > 30000) {
    return new Date((value - 25569) * 86400 * 1000);
  }
  
  return value;
}

//==================== HELPER: PARSE NUMBER ====================
function parseNumber(value) {
  if (!value) return 0;
  if (typeof value === 'number') return value;
  
  var cleaned = value.toString().replace(/[^\d.-]/g, '');
  var parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

//==================== MAIN ENTRY POINT ====================
function doGet(e) {
  setupSheets();
  
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Coaching Monitoring System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return htmlOutput;
}

//==================== API: LOGIN ====================
function login(npk) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    if (!sheet) return { success: false, message: 'Database tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var user = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === npk.toString()) {
        user = {
          npk: data[i][0],
          nama: data[i][1],
          jabatan: data[i][2],
          cl: parseInt(data[i][3]) || 99,
          cabang: data[i][4],
          region: data[i][5],
          atasan_npk: data[i][6],
          email: data[i][7],
          no_hp: data[i][8]
        };
        break;
      }
    }
    
    if (!user) return { success: false, message: 'NPK tidak ditemukan' };
    
    // Create session
    var sessionId = Utilities.getUuid();
    var expiry = new Date();
    expiry.setHours(expiry.getHours() + 24);
    
    // Store in PropertiesService
    var props = PropertiesService.getScriptProperties();
    props.setProperty('SESSION_' + sessionId, JSON.stringify({
      user: user,
      expiry: expiry.getTime()
    }));
    
    return {
      success: true,
      sessionId: sessionId,
      user: user,
      canUpload: canUploadData(user.npk, user.jabatan),
      landingPage: 'pencapaian' // NEW: Landing page = Dashboard Pencapaian
    };
    
  } catch (e) {
    console.error('Error login:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: CHECK SESSION ====================
function checkSession(sessionId) {
  try {
    if (!sessionId) return null;
    
    var props = PropertiesService.getScriptProperties();
    var sessionData = props.getProperty('SESSION_' + sessionId);
    
    if (!sessionData) return null;
    
    var session = JSON.parse(sessionData);
    if (new Date().getTime() > session.expiry) {
      props.deleteProperty('SESSION_' + sessionId);
      return null;
    }
    
    return session.user;
    
  } catch (e) {
    console.error('Error checkSession:', e);
    return null;
  }
}

//==================== API: GET CURRENT USER ====================
function getCurrentUser() {
  try {
    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    
    for (var key in allProps) {
      if (key.startsWith('SESSION_')) {
        var session = JSON.parse(allProps[key]);
        if (new Date().getTime() <= session.expiry) {
          return session.user;
        }
      }
    }
    return null;
  } catch (e) {
    console.error('Error getCurrentUser:', e);
    return null;
  }
}

//==================== API: LOGOUT ====================
function logout(sessionId) {
  try {
    if (sessionId) {
      var props = PropertiesService.getScriptProperties();
      props.deleteProperty('SESSION_' + sessionId);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

//==================== API: GET DASHBOARD DATA (COACHING - EXISTING) ====================
function getDashboardData(sessionId) {
  return getInitialData(sessionId);
}

function getInitialData(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var totalCoaching = 0;
    var belumCoaching = 0;
    var sudahCoachingBelumUpdate = 0;
    var sudahCoachingSudahUpdate = 0;
    
    for (var i = 1; i < data.length; i++) {
      var status = data[i][6] ? data[i][6].toString().trim() : '';
      
      if (status === 'OPEN') belumCoaching++;
      else if (status === 'ON PROGRESS') sudahCoachingBelumUpdate++;
      else if (status === 'DONE') sudahCoachingSudahUpdate++;
      
      totalCoaching++;
    }
    
    return {
      success: true,
      data: {
        totalCoaching: totalCoaching,
        belumCoaching: belumCoaching,
        sudahCoachingBelumUpdate: sudahCoachingBelumUpdate,
        sudahCoachingSudahUpdate: sudahCoachingSudahUpdate
      },
      user: user,
      canUpload: canUploadData(user.npk, user.jabatan)
    };
    
  } catch (e) {
    console.error('Error getInitialData:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET COACHING LIST ====================
function getCoachingList(sessionId, filters) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      result.push({
        coaching_id: row[0],
        coach_npk: row[1],
        coachee_npk: row[2],
        cabang: row[3],
        root_cause: row[4],
        topic: row[5],
        status: row[6],
        created_date: row[7] instanceof Date ? row[7].toISOString().split('T')[0] : '',
        target_date: row[8] instanceof Date ? row[8].toISOString().split('T')[0] : ''
      });
    }
    
    return { success: true, data: result, user: user };
    
  } catch (e) {
    console.error('Error getCoachingList:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: SAVE COACHING ====================
function saveCoachingHeader(sessionId, coachingData) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    
    var coachingId = 'C' + new Date().getTime();
    var createdDate = new Date();
    var targetDate = coachingData.target_date ? new Date(coachingData.target_date) : new Date(createdDate.getTime() + 7 * 24 * 60 * 60 * 1000);
    
    sheet.appendRow([
      coachingId,
      coachingData.coach_npk || user.npk,
      coachingData.coachee_npk,
      coachingData.cabang || user.cabang,
      coachingData.root_cause,
      coachingData.topic,
      'OPEN',
      createdDate,
      targetDate
    ]);
    
    return { success: true, coaching_id: coachingId };
    
  } catch (e) {
    console.error('Error saveCoachingHeader:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: UPDATE COACHING ====================
function updateCoaching(sessionId, coachingId, updateData) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === coachingId) {
        // Update status
        if (updateData.status) {
          sheet.getRange(i + 1, 7).setValue(updateData.status);
        }
        break;
      }
    }
    
    // Save detail
    if (updateData.detail) {
      var detailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_DETAIL);
      if (detailSheet) {
        var detailId = 'D' + new Date().getTime();
        detailSheet.appendRow([
          detailId,
          coachingId,
          updateData.detail.week || 1,
          updateData.detail.action,
          updateData.detail.how,
          updateData.detail.target_date ? new Date(updateData.detail.target_date) : new Date(),
          updateData.detail.result,
          updateData.detail.feedback,
          new Date(),
          updateData.detail.target_nominal || 0,
          updateData.detail.target_satuan || ''
        ]);
      }
    }
    
    return { success: true };
    
  } catch (e) {
    console.error('Error updateCoaching:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: COMPLETE COACHING ====================
function completeCoaching(sessionId, coachingId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === coachingId) {
        sheet.getRange(i + 1, 7).setValue('DONE');
        return { success: true };
      }
    }
    
    return { success: false, message: 'Coaching tidak ditemukan' };
    
  } catch (e) {
    console.error('Error completeCoaching:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: FORCE CLOSE COACHING ====================
function forceCloseCoaching(sessionId, coachingId, reason) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === coachingId) {
        sheet.getRange(i + 1, 7).setValue('FORCE_CLOSED');
        // Save reason to detail
        var detailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_DETAIL);
        if (detailSheet) {
          detailSheet.appendRow([
            'D_FORCE_' + new Date().getTime(),
            coachingId,
            0,
            'FORCE_CLOSE',
            reason,
            new Date(),
            'FORCE_CLOSED',
            'Force closed by ' + user.nama,
            new Date(),
            0,
            ''
          ]);
        }
        return { success: true };
      }
    }
    
    return { success: false, message: 'Coaching tidak ditemukan' };
    
  } catch (e) {
    console.error('Error forceCloseCoaching:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET COACHING DETAIL ====================
function getCoachingDetail(sessionId, coachingId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var headerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    var detailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_DETAIL);
    
    if (!headerSheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var headerData = headerSheet.getDataRange().getValues();
    var coaching = null;
    
    for (var i = 1; i < headerData.length; i++) {
      if (headerData[i][0] === coachingId) {
        coaching = {
          coaching_id: headerData[i][0],
          coach_npk: headerData[i][1],
          coachee_npk: headerData[i][2],
          cabang: headerData[i][3],
          root_cause: headerData[i][4],
          topic: headerData[i][5],
          status: headerData[i][6],
          created_date: headerData[i][7] instanceof Date ? headerData[i][7].toISOString().split('T')[0] : '',
          target_date: headerData[i][8] instanceof Date ? headerData[i][8].toISOString().split('T')[0] : ''
        };
        break;
      }
    }
    
    if (!coaching) return { success: false, message: 'Coaching tidak ditemukan' };
    
    var details = [];
    if (detailSheet) {
      var detailData = detailSheet.getDataRange().getValues();
      for (var i = 1; i < detailData.length; i++) {
        if (detailData[i][1] === coachingId) {
          details.push({
            detail_id: detailData[i][0],
            week: detailData[i][2],
            action: detailData[i][3],
            how: detailData[i][4],
            target_date: detailData[i][5] instanceof Date ? detailData[i][5].toISOString().split('T')[0] : '',
            result: detailData[i][6],
            feedback: detailData[i][7],
            update_date: detailData[i][8] instanceof Date ? detailData[i][8].toISOString().split('T')[0] : '',
            target_nominal: detailData[i][9],
            target_satuan: detailData[i][10]
          });
        }
      }
    }
    
    return { success: true, coaching: coaching, details: details, user: user };
    
  } catch (e) {
    console.error('Error getCoachingDetail:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET REMINDER DATA ====================
function getReminderData(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    var upcoming = [];
    var overdue = [];
    
    for (var i = 1; i < data.length; i++) {
      var status = data[i][6] ? data[i][6].toString().trim() : '';
      var targetDate = data[i][8] instanceof Date ? data[i][8] : new Date();
      
      if (status === 'OPEN' || status === 'ON PROGRESS') {
        var diffDays = Math.ceil((targetDate - now) / (1000 * 60 * 60 * 24));
        
        if (diffDays < 0) {
          overdue.push({
            coaching_id: data[i][0],
            coach_npk: data[i][1],
            coachee_npk: data[i][2],
            cabang: data[i][3],
            topic: data[i][5],
            status: status,
            target_date: targetDate.toISOString().split('T')[0],
            days_overdue: Math.abs(diffDays)
          });
        } else if (diffDays <= 7) {
          upcoming.push({
            coaching_id: data[i][0],
            coach_npk: data[i][1],
            coachee_npk: data[i][2],
            cabang: data[i][3],
            topic: data[i][5],
            status: status,
            target_date: targetDate.toISOString().split('T')[0],
            days_remaining: diffDays
          });
        }
      }
    }
    
    return {
      success: true,
      upcoming: upcoming,
      overdue: overdue,
      total_upcoming: upcoming.length,
      total_overdue: overdue.length,
      user: user
    };
    
  } catch (e) {
    console.error('Error getReminderData:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET KARYAWAN LIST ====================
function getKaryawanList(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      result.push({
        npk: data[i][0],
        nama: data[i][1],
        jabatan: data[i][2],
        cl: data[i][3],
        cabang: data[i][4],
        region: data[i][5],
        atasan_npk: data[i][6],
        email: data[i][7],
        no_hp: data[i][8]
      });
    }
    
    return { success: true, data: result, user: user };
    
  } catch (e) {
    console.error('Error getKaryawanList:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: SAVE KARYAWAN ====================
function saveKaryawan(sessionId, karyawanData) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    
    // Check if NPK exists
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === karyawanData.npk.toString()) {
        // Update existing
        sheet.getRange(i + 1, 2).setValue(karyawanData.nama);
        sheet.getRange(i + 1, 3).setValue(karyawanData.jabatan);
        sheet.getRange(i + 1, 4).setValue(karyawanData.cl);
        sheet.getRange(i + 1, 5).setValue(karyawanData.cabang);
        sheet.getRange(i + 1, 6).setValue(karyawanData.region);
        sheet.getRange(i + 1, 7).setValue(karyawanData.atasan_npk);
        sheet.getRange(i + 1, 8).setValue(karyawanData.email);
        sheet.getRange(i + 1, 9).setValue(karyawanData.no_hp);
        return { success: true, message: 'Data karyawan diperbarui' };
      }
    }
    
    // Insert new
    sheet.appendRow([
      karyawanData.npk,
      karyawanData.nama,
      karyawanData.jabatan,
      karyawanData.cl,
      karyawanData.cabang,
      karyawanData.region,
      karyawanData.atasan_npk,
      karyawanData.email,
      karyawanData.no_hp
    ]);
    
    return { success: true, message: 'Karyawan baru ditambahkan' };
    
  } catch (e) {
    console.error('Error saveKaryawan:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: DELETE KARYAWAN ====================
function deleteKaryawan(sessionId, npk) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === npk.toString()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Karyawan dihapus' };
      }
    }
    
    return { success: false, message: 'Karyawan tidak ditemukan' };
    
  } catch (e) {
    console.error('Error deleteKaryawan:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET USER LIST ====================
function getUserList(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      result.push({
        npk: data[i][0],
        nama: data[i][1],
        role: data[i][2],
        status: data[i][3],
        created_at: data[i][4] instanceof Date ? data[i][4].toISOString().split('T')[0] : '',
        updated_at: data[i][5] instanceof Date ? data[i][5].toISOString().split('T')[0] : '',
        created_by: data[i][6]
      });
    }
    
    return { success: true, data: result, user: user };
    
  } catch (e) {
    console.error('Error getUserList:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: SAVE USER ====================
function saveUser(sessionId, userData) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    
    var now = new Date();
    
    // Check if NPK exists
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === userData.npk.toString()) {
        // Update existing
        sheet.getRange(i + 1, 2).setValue(userData.nama);
        sheet.getRange(i + 1, 3).setValue(userData.role);
        sheet.getRange(i + 1, 4).setValue(userData.status);
        sheet.getRange(i + 1, 5).setValue(now);
        return { success: true, message: 'User diperbarui' };
      }
    }
    
    // Insert new
    sheet.appendRow([
      userData.npk,
      userData.nama,
      userData.role,
      userData.status || 'ACTIVE',
      now,
      now,
      user.nama
    ]);
    
    return { success: true, message: 'User baru ditambahkan' };
    
  } catch (e) {
    console.error('Error saveUser:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: DELETE USER ====================
function deleteUser(sessionId, npk) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === npk.toString()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'User dihapus' };
      }
    }
    
    return { success: false, message: 'User tidak ditemukan' };
    
  } catch (e) {
    console.error('Error deleteUser:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: EXPORT COACHING TO EXCEL ====================
function exportCoachingToExcel(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var csv = '';
    
    // Headers
    csv += 'Coaching ID,Coach NPK,Coachee NPK,Cabang,Root Cause,Topic,Status,Created Date,Target Date\n';
    
    for (var i = 1; i < data.length; i++) {
      csv += [
        data[i][0], data[i][1], data[i][2], data[i][3],
        '"' + (data[i][4] || '').toString().replace(/"/g, '""') + '"',
        '"' + (data[i][5] || '').toString().replace(/"/g, '""') + '"',
        data[i][6],
        data[i][7] instanceof Date ? data[i][7].toISOString().split('T')[0] : '',
        data[i][8] instanceof Date ? data[i][8].toISOString().split('T')[0] : ''
      ].join(',') + '\n';
    }
    
    return {
      success: true,
      data: Utilities.base64Encode(Utilities.newBlob(csv, 'text/csv').getBytes()),
      filename: 'Coaching_Data_' + new Date().toISOString().split('T')[0] + '.csv'
    };
    
  } catch (e) {
    console.error('Error exportCoachingToExcel:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: EXPORT KARYAWAN TO EXCEL ====================
function exportKaryawanToExcel(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var csv = '';
    
    csv += 'NPK,Nama,Jabatan,CL,Cabang,Region,Atasan NPK,Email,No HP\n';
    
    for (var i = 1; i < data.length; i++) {
      csv += [
        data[i][0], data[i][1], data[i][2], data[i][3], data[i][4],
        data[i][5], data[i][6], data[i][7], data[i][8]
      ].join(',') + '\n';
    }
    
    return {
      success: true,
      data: Utilities.base64Encode(Utilities.newBlob(csv, 'text/csv').getBytes()),
      filename: 'Karyawan_Data_' + new Date().toISOString().split('T')[0] + '.csv'
    };
    
  } catch (e) {
    console.error('Error exportKaryawanToExcel:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}
//==================== API: UPLOAD APPROVE BLM CAIR DATA ====================
function uploadApproveBlmCairData(fileData, fileName) {
  try {
    var user = getCurrentUser();
    if (!user) return { success: false, message: 'Session expired' };
    
    if (!canUploadData(user.npk, user.jabatan)) {
      return { success: false, message: 'Anda tidak memiliki akses upload data' };
    }
    
    setupApproveBlmCairSheet();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPROVE_BLM_CAIR);
    if (!sheet) throw new Error('Sheet tidak ditemukan');
    
    var decoded = Utilities.base64Decode(fileData.split(',')[1] || fileData);
    var blob = Utilities.newBlob(decoded, MimeType.MICROSOFT_EXCEL, fileName);
    
    var parsedData = parseExcelData(blob);
    if (!parsedData || parsedData.length === 0) {
      return { success: false, message: 'Data kosong atau format tidak valid' };
    }
    
    var existingData = sheet.getDataRange().getValues();
    var existingIds = new Set();
    for (var i = 1; i < existingData.length; i++) {
      if (existingData[i][0]) existingIds.add(existingData[i][0].toString());
    }
    
    var newRows = [];
    var duplicateCount = 0;
    var uploadDate = new Date();
    
    for (var i = 0; i < parsedData.length; i++) {
      var row = parsedData[i];
      var id = row.FPK_AWAL || row.FPK_CABANG || ('AUTO_' + Utilities.getUuid());
      
      if (existingIds.has(id.toString())) {
        duplicateCount++;
        continue;
      }
      
      var newRow = [
        id,
        parseDate(row.TGL),
        row.NAMA || '',
        row.CABANG || '',
        row.SUMBER2 || '',
        row.STATUS_APLIKASI || '',
        row.JENIS_KREDIT || '',
        row.KEPUTUSAN || '',
        row.NIK || '',
        row.NAMA_SALES || '',
        row.NAMA_SPV || '',
        row.NAMA_SALES_QA || '',
        row.NAMA_SALES_QA_BMR || '',
        row.NAMA_SPV_QA || '',
        row.NAMA_SPV_QA_BMR || '',
        row.STATUS_NASABAH || '',
        row.FPK_AWAL || '',
        row.FPK_CABANG || '',
        row.NAMA_SCORING || '',
        parseNumber(row.PLAFOND),
        parseNumber(row.NOMINAL_ACC),
        parseNumber(row.BUNGA),
        row.KATEGORI || '',
        row.KATEGORI_BUNGA2 || '',
        row.MERK || '',
        row.TYPE || '',
        row.TAHUN || '',
        row.JARAK || '',
        row.AN_JAMINAN || '',
        row.NO_HP || '',
        row.NO_LOAN || '',
        parseDate(row.TGL_HANDOFF),
        row.KETERANGAN_SALES || '',
        uploadDate,
        user.npk + ' - ' + user.nama
      ];
      
      newRows.push(newRow);
      existingIds.add(id.toString());
    }
    
    if (newRows.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
    
    return {
      success: true,
      message: 'Upload berhasil!',
      totalProcessed: parsedData.length,
      newInserted: newRows.length,
      duplicateSkipped: duplicateCount
    };
    
  } catch (e) {
    console.error('Error uploadApproveBlmCairData:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== HELPER: PARSE EXCEL DATA ====================
function parseExcelData(blob) {
  try {
    var file = Drive.Files.insert({
      title: 'temp_' + new Date().getTime(),
      mimeType: MimeType.GOOGLE_SHEETS
    }, blob);
    
    var tempSheet = SpreadsheetApp.openById(file.id);
    var sheet = tempSheet.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    
    Drive.Files.remove(file.id);
    
    if (data.length < 2) return [];
    
    var headers = data[0];
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j].toString()
          .replace(/\s+/g, '_')
          .replace(/<br>/g, '')
          .toUpperCase();
        row[header] = data[i][j];
      }
      result.push(row);
    }
    
    return result;
  } catch (e) {
    console.error('Error parseExcelData:', e);
    return [];
  }
}

//==================== API: GET DASHBOARD PENCAPAIAN DATA ====================
function getDashboardPencapaianData(filters) {
  try {
    var user = getCurrentUser();
    if (!user) return { success: false, message: 'Session expired' };
    
    setupApproveBlmCairSheet();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPROVE_BLM_CAIR);
    if (!sheet) {
      return {
        success: true,
        data: getEmptyDashboardData(),
        canUpload: canUploadData(user.npk, user.jabatan)
      };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        success: true,
        data: getEmptyDashboardData(),
        canUpload: canUploadData(user.npk, user.jabatan)
      };
    }
    
    var periodeStart = filters && filters.periodeStart ? new Date(filters.periodeStart) : null;
    var periodeEnd = filters && filters.periodeEnd ? new Date(filters.periodeEnd) : null;
    var selectedCabang = filters && filters.cabang ? filters.cabang : [];
    var selectedStatus = filters && filters.status ? filters.status : [];
    var selectedKeputusan = filters && filters.keputusan ? filters.keputusan : [];
    
    var filteredData = [];
    var cabangList = new Set();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var tgl = row[1] instanceof Date ? row[1] : parseDate(row[1]);
      var cabang = row[3] ? row[3].toString() : '';
      var statusAplikasi = row[5] ? row[5].toString().trim() : '';
      var keputusan = row[7] ? row[7].toString().trim() : '';
      
      if (cabang) cabangList.add(cabang);
      
      if (periodeStart && tgl && tgl < periodeStart) continue;
      if (periodeEnd && tgl && tgl > periodeEnd) continue;
      
      if (selectedCabang && selectedCabang.length > 0 && !selectedCabang.includes(cabang)) continue;
      if (selectedStatus && selectedStatus.length > 0 && !selectedStatus.includes(statusAplikasi)) continue;
      if (selectedKeputusan && selectedKeputusan.length > 0 && !selectedKeputusan.includes(keputusan)) continue;
      
      filteredData.push(row);
    }
    
    var metrics = calculateMetrics(filteredData);
    var perCabang = calculatePerCabang(filteredData);
    var recentData = getRecentData(filteredData, 10);
    
    return {
      success: true,
      data: {
        metrics: metrics,
        perCabang: perCabang,
        recentData: recentData,
        cabangList: Array.from(cabangList).sort()
      },
      canUpload: canUploadData(user.npk, user.jabatan),
      userCl: user.cl
    };
    
  } catch (e) {
    console.error('Error getDashboardPencapaianData:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== HELPER: CALCULATE METRICS ====================
function calculateMetrics(data) {
  var total = data.length;
  var cair = 0, blmCair = 0, apkNaik = 0;
  var approve = 0, reject = 0, pending = 0;
  var totalPlafond = 0, totalNominalAcc = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var statusAplikasi = row[5] ? row[5].toString().trim() : '';
    var keputusan = row[7] ? row[7].toString().trim() : '';
    var plafond = parseFloat(row[19]) || 0;
    var nominalAcc = parseFloat(row[20]) || 0;
    
    if (statusAplikasi === 'CAIR') cair++;
    else if (statusAplikasi === 'APP BLM CAIR') blmCair++;
    else if (statusAplikasi === 'APK NAIK') apkNaik++;
    
    if (keputusan === 'APPROVE') approve++;
    else if (keputusan === 'REJECT') reject++;
    else if (keputusan === 'PENDING') pending++;
    
    totalPlafond += plafond;
    totalNominalAcc += nominalAcc;
  }
  
  var achievementRate = totalPlafond > 0 ? (totalNominalAcc / totalPlafond * 100).toFixed(1) : 0;
  var approvalRate = total > 0 ? (approve / total * 100).toFixed(1) : 0;
  var gap = totalPlafond - totalNominalAcc;
  
  return {
    totalMasuk: total,
    cair: cair,
    blmCair: blmCair,
    apkNaik: apkNaik,
    approve: approve,
    reject: reject,
    pending: pending,
    approvalRate: approvalRate,
    totalPlafond: totalPlafond,
    totalNominalAcc: totalNominalAcc,
    achievementRate: achievementRate,
    gap: gap
  };
}

//==================== HELPER: CALCULATE PER CABANG ====================
function calculatePerCabang(data) {
  var cabangMap = {};
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var cabang = row[3] ? row[3].toString() : 'Unknown';
    var plafond = parseFloat(row[19]) || 0;
    var nominalAcc = parseFloat(row[20]) || 0;
    var statusAplikasi = row[5] ? row[5].toString().trim() : '';
    var keputusan = row[7] ? row[7].toString().trim() : '';
    
    if (!cabangMap[cabang]) {
      cabangMap[cabang] = {
        cabang: cabang,
        total: 0,
        plafond: 0,
        nominalAcc: 0,
        cair: 0,
        blmCair: 0,
        apkNaik: 0,
        approve: 0,
        reject: 0,
        pending: 0
      };
    }
    
    var c = cabangMap[cabang];
    c.total++;
    c.plafond += plafond;
    c.nominalAcc += nominalAcc;
    
    if (statusAplikasi === 'CAIR') c.cair++;
    else if (statusAplikasi === 'APP BLM CAIR') c.blmCair++;
    else if (statusAplikasi === 'APK NAIK') c.apkNaik++;
    
    if (keputusan === 'APPROVE') c.approve++;
    else if (keputusan === 'REJECT') c.reject++;
    else if (keputusan === 'PENDING') c.pending++;
  }
  
  return Object.values(cabangMap).sort(function(a, b) {
    return b.plafond - a.plafond;
  });
}

//==================== HELPER: GET RECENT DATA ====================
function getRecentData(data, limit) {
  var sorted = data.sort(function(a, b) {
    var dateA = a[1] instanceof Date ? a[1] : new Date(0);
    var dateB = b[1] instanceof Date ? b[1] : new Date(0);
    return dateB - dateA;
  });
  
  var result = [];
  for (var i = 0; i < Math.min(limit, sorted.length); i++) {
    var row = sorted[i];
    result.push({
      id: row[0],
      tgl: row[1] instanceof Date ? row[1].toISOString().split('T')[0] : '',
      nama: row[2],
      cabang: row[3],
      statusAplikasi: row[5],
      keputusan: row[7],
      plafond: parseFloat(row[19]) || 0,
      nominalAcc: parseFloat(row[20]) || 0,
      merk: row[24],
      type: row[25],
      tahun: row[26]
    });
  }
  
  return result;
}

//==================== HELPER: EMPTY DASHBOARD ====================
function getEmptyDashboardData() {
  return {
    metrics: {
      totalMasuk: 0, cair: 0, blmCair: 0, apkNaik: 0,
      approve: 0, reject: 0, pending: 0, approvalRate: 0,
      totalPlafond: 0, totalNominalAcc: 0, achievementRate: 0, gap: 0
    },
    perCabang: [],
    recentData: [],
    cabangList: []
  };
}

//==================== API: GET TABLE PENCAPAIAN DATA (PAGINATION) ====================
function getTablePencapaianData(params) {
  try {
    var user = getCurrentUser();
    if (!user) return { success: false, message: 'Session expired' };
    
    var page = params && params.page ? parseInt(params.page) : 1;
    var limit = params && params.limit ? parseInt(params.limit) : 20;
    var search = params && params.search ? params.search.toString().toLowerCase() : '';
    var sortField = params && params.sortField ? params.sortField : 'TGL';
    var sortOrder = params && params.sortOrder ? params.sortOrder : 'DESC';
    var filters = params && params.filters ? params.filters : {};
    
    setupApproveBlmCairSheet();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPROVE_BLM_CAIR);
    if (!sheet) {
      return { success: true, data: [], total: 0, page: page, totalPages: 0 };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: [], total: 0, page: page, totalPages: 0 };
    }
    
    var filteredData = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      if (filters.periodeStart) {
        var tgl = row[1] instanceof Date ? row[1] : parseDate(row[1]);
        if (tgl && tgl < new Date(filters.periodeStart)) continue;
      }
      if (filters.periodeEnd) {
        var tgl = row[1] instanceof Date ? row[1] : parseDate(row[1]);
        if (tgl && tgl > new Date(filters.periodeEnd)) continue;
      }
      if (filters.cabang && filters.cabang.length > 0) {
        if (!filters.cabang.includes(row[3])) continue;
      }
      if (filters.status && filters.status.length > 0) {
        if (!filters.status.includes(row[5])) continue;
      }
      if (filters.keputusan && filters.keputusan.length > 0) {
        if (!filters.keputusan.includes(row[7])) continue;
      }
      
      if (search) {
        var found = false;
        for (var j = 0; j < row.length; j++) {
          if (row[j] && row[j].toString().toLowerCase().includes(search)) {
            found = true;
            break;
          }
        }
        if (!found) continue;
      }
      
      filteredData.push(row);
    }
    
    var sortIndex = 0;
    var headerMap = {
      'ID': 0, 'TGL': 1, 'NAMA': 2, 'CABANG': 3, 'STATUS_APLIKASI': 5,
      'KEPUTUSAN': 7, 'PLAFOND': 19, 'NOMINAL_ACC': 20
    };
    sortIndex = headerMap[sortField] || 1;
    
    filteredData.sort(function(a, b) {
      var valA = a[sortIndex] || '';
      var valB = b[sortIndex] || '';
      
      if (sortField === 'TGL') {
        valA = valA instanceof Date ? valA : new Date(0);
        valB = valB instanceof Date ? valB : new Date(0);
      }
      
      if (sortOrder === 'ASC') {
        return valA > valB ? 1 : -1;
      } else {
        return valA < valB ? 1 : -1;
      }
    });
    
    var total = filteredData.length;
    var totalPages = Math.ceil(total / limit);
    var startIndex = (page - 1) * limit;
    var paginatedData = filteredData.slice(startIndex, startIndex + limit);
    
    var formattedData = [];
    for (var i = 0; i < paginatedData.length; i++) {
      var row = paginatedData[i];
      formattedData.push({
        id: row[0],
        tgl: row[1] instanceof Date ? row[1].toISOString().split('T')[0] : '',
        nama: row[2],
        cabang: row[3],
        sumber: row[4],
        statusAplikasi: row[5],
        jenisKredit: row[6],
        keputusan: row[7],
        nik: row[8],
        namaSales: row[9],
        namaSpv: row[10],
        statusNasabah: row[15],
        fpkAwal: row[16],
        fpkCabang: row[17],
        namaScoring: row[18],
        plafond: parseFloat(row[19]) || 0,
        nominalAcc: parseFloat(row[20]) || 0,
        bunga: parseFloat(row[21]) || 0,
        kategori: row[22],
        merk: row[24],
        type: row[25],
        tahun: row[26],
        jarak: row[27],
        anJaminan: row[28],
        noHp: row[29],
        noLoan: row[30],
        tglHandoff: row[31] instanceof Date ? row[31].toISOString().split('T')[0] : '',
        keteranganSales: row[32]
      });
    }
    
    return {
      success: true,
      data: formattedData,
      total: total,
      page: page,
      totalPages: totalPages,
      limit: limit
    };
    
  } catch (e) {
    console.error('Error getTablePencapaianData:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET CABANG LIST FOR FILTER ====================
function getCabangListForFilter() {
  try {
    var user = getCurrentUser();
    if (!user) return { success: false, message: 'Session expired' };
    
    var allCabang = getCabangList();
    
    // Filter berdasarkan CL level
    if (user.cl >= 3) {
      // User biasa hanya lihat cabang sendiri
      return {
        success: true,
        data: [user.cabang],
        canSelectMultiple: false
      };
    } else if (user.cl === 2) {
      // Admin/Region lihat cabang dalam region
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        var regionCabang = new Set();
        for (var i = 1; i < data.length; i++) {
          if (data[i][5] === user.region) {
            regionCabang.add(data[i][3].toString());
          }
        }
        return {
          success: true,
          data: Array.from(regionCabang).sort(),
          canSelectMultiple: true
        };
      }
    }
    
    // Owner (CL 1) lihat semua
    return {
      success: true,
      data: allCabang,
      canSelectMultiple: true
    };
    
  } catch (e) {
    console.error('Error getCabangListForFilter:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: DELETE PENCAPAIAN DATA ====================
function deletePencapaianData(sessionId, ids) {
  try {
    var user = getCurrentUser();
    if (!user) return { success: false, message: 'Session expired' };
    
    if (!canUploadData(user.npk, user.jabatan)) {
      return { success: false, message: 'Anda tidak memiliki akses' };
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_APPROVE_BLM_CAIR);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var deletedCount = 0;
    
    // Delete dari bawah ke atas agar index tidak bergeser
    for (var i = data.length - 1; i >= 1; i--) {
      if (ids.includes(data[i][0].toString())) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    
    return {
      success: true,
      message: deletedCount + ' data berhasil dihapus'
    };
    
  } catch (e) {
    console.error('Error deletePencapaianData:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: EXPORT PENCAPAIAN TO EXCEL ====================
function exportPencapaianToExcel(params) {
  try {
    var user = getCurrentUser();
    if (!user) return { success: false, message: 'Session expired' };
    
    var result = getTablePencapaianData({
      page: 1,
      limit: 999999,
      search: params.search || '',
      sortField: params.sortField || 'TGL',
      sortOrder: params.sortOrder || 'DESC',
      filters: params.filters || {}
    });
    
    if (!result.success) return result;
    
    var csv = 'ID,TGL,NAMA,CABANG,SUMBER,STATUS_APLIKASI,JENIS_KREDIT,KEPUTUSAN,NIK,NAMA_SALES,NAMA_SPV,STATUS_NASABAH,FPK_AWAL,FPK_CABANG,NAMA_SCORING,PLAFOND,NOMINAL_ACC,BUNGA,KATEGORI,MERK,TYPE,TAHUN,JARAK,AN_JAMINAN,NO_HP,NO_LOAN,TGL_HANDOFF,KETERANGAN_SALES\n';
    
    for (var i = 0; i < result.data.length; i++) {
      var d = result.data[i];
      csv += [
        d.id, d.tgl, d.nama, d.cabang, d.sumber, d.statusAplikasi,
        d.jenisKredit, d.keputusan, d.nik, d.namaSales, d.namaSpv,
        d.statusNasabah, d.fpkAwal, d.fpkCabang, d.namaScoring,
        d.plafond, d.nominalAcc, d.bunga, d.kategori,
        d.merk, d.type, d.tahun, d.jarak, d.anJaminan,
        d.noHp, d.noLoan, d.tglHandoff, d.keteranganSales
      ].join(',') + '\n';
    }
    
    return {
      success: true,
      data: Utilities.base64Encode(Utilities.newBlob(csv, 'text/csv').getBytes()),
      filename: 'Pencapaian_Data_' + new Date().toISOString().split('T')[0] + '.csv'
    };
    
  } catch (e) {
    console.error('Error exportPencapaianToExcel:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}
//==================== API: GET KARYAWAN BY NPK ====================
function getKaryawanByNpk(sessionId, npk) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === npk.toString()) {
        return {
          success: true,
          data: {
            npk: data[i][0],
            nama: data[i][1],
            jabatan: data[i][2],
            cl: data[i][3],
            cabang: data[i][4],
            region: data[i][5],
            atasan_npk: data[i][6],
            email: data[i][7],
            no_hp: data[i][8]
          }
        };
      }
    }
    
    return { success: false, message: 'Karyawan tidak ditemukan' };
    
  } catch (e) {
    console.error('Error getKaryawanByNpk:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET ATASAN LIST ====================
function getAtasanList(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] && parseInt(data[i][3]) <= 2) {
        result.push({
          npk: data[i][0],
          nama: data[i][1],
          jabatan: data[i][2],
          cl: data[i][3]
        });
      }
    }
    
    return { success: true, data: result };
    
  } catch (e) {
    console.error('Error getAtasanList:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: VALIDATE COACHING ACTIVE ====================
function validateCoachingActive(sessionId, coacheeNpk) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: true, isActive: false };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][2].toString() === coacheeNpk.toString()) {
        var status = data[i][6] ? data[i][6].toString().trim() : '';
        if (status === 'OPEN' || status === 'ON PROGRESS') {
          return {
            success: true,
            isActive: true,
            coachingId: data[i][0],
            status: status
          };
        }
      }
    }
    
    return { success: true, isActive: false };
    
  } catch (e) {
    console.error('Error validateCoachingActive:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET COACHING BY COACHEE ====================
function getCoachingByCoachee(sessionId, coacheeNpk) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_HEADER);
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    var data = sheet.getDataRange().getValues();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][2].toString() === coacheeNpk.toString()) {
        result.push({
          coaching_id: data[i][0],
          coach_npk: data[i][1],
          coachee_npk: data[i][2],
          cabang: data[i][3],
          root_cause: data[i][4],
          topic: data[i][5],
          status: data[i][6],
          created_date: data[i][7] instanceof Date ? data[i][7].toISOString().split('T')[0] : '',
          target_date: data[i][8] instanceof Date ? data[i][8].toISOString().split('T')[0] : ''
        });
      }
    }
    
    return { success: true, data: result };
    
  } catch (e) {
    console.error('Error getCoachingByCoachee:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET WEEKLY PROGRESS ====================
function getWeeklyProgress(sessionId, coachingId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var detailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COACHING_DETAIL);
    if (!detailSheet) return { success: true, data: [] };
    
    var data = detailSheet.getDataRange().getValues();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === coachingId) {
        result.push({
          detail_id: data[i][0],
          week: data[i][2],
          action: data[i][3],
          how: data[i][4],
          target_date: data[i][5] instanceof Date ? data[i][5].toISOString().split('T')[0] : '',
          result: data[i][6],
          feedback: data[i][7],
          update_date: data[i][8] instanceof Date ? data[i][8].toISOString().split('T')[0] : '',
          target_nominal: data[i][9],
          target_satuan: data[i][10]
        });
      }
    }
    
    return { success: true, data: result };
    
  } catch (e) {
    console.error('Error getWeeklyProgress:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: SAVE REMINDER LOG ====================
function saveReminderLog(sessionId, reminderData) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REMINDER_LOG);
    if (!sheet) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.insertSheet(SHEET_REMINDER_LOG);
      sheet.getRange(1, 1, 1, 10).setValues([['reminder_id', 'coaching_id', 'coach_npk', 'coachee_npk', 'reminder_date', 'channel', 'status', 'sent_timestamp', 'message_content', 'created_at']]);
    }
    
    var reminderId = 'R' + new Date().getTime();
    
    sheet.appendRow([
      reminderId,
      reminderData.coaching_id,
      reminderData.coach_npk,
      reminderData.coachee_npk,
      new Date(reminderData.reminder_date),
      reminderData.channel || 'APP',
      reminderData.status || 'SENT',
      new Date(),
      reminderData.message_content || '',
      new Date()
    ]);
    
    return { success: true, reminder_id: reminderId };
    
  } catch (e) {
    console.error('Error saveReminderLog:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET REMINDER HISTORY ====================
function getReminderHistory(sessionId, coachingId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REMINDER_LOG);
    if (!sheet) return { success: true, data: [] };
    
    var data = sheet.getDataRange().getValues();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === coachingId) {
        result.push({
          reminder_id: data[i][0],
          coaching_id: data[i][1],
          coach_npk: data[i][2],
          coachee_npk: data[i][3],
          reminder_date: data[i][4] instanceof Date ? data[i][4].toISOString().split('T')[0] : '',
          channel: data[i][5],
          status: data[i][6],
          sent_timestamp: data[i][7] instanceof Date ? data[i][7].toISOString().split('T')[0] + ' ' + data[i][7].toTimeString().split(' ')[0] : '',
          message_content: data[i][8]
        });
      }
    }
    
    return { success: true, data: result };
    
  } catch (e) {
    console.error('Error getReminderHistory:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET DASHBOARD SUMMARY ====================
function getDashboardSummary(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var coachingResult = getInitialData(sessionId);
    var pencapaianResult = getDashboardPencapaianData({});
    
    return {
      success: true,
      coaching: coachingResult.success ? coachingResult.data : null,
      pencapaian: pencapaianResult.success ? pencapaianResult.data.metrics : null,
      user: user,
      canUpload: canUploadData(user.npk, user.jabatan)
    };
    
  } catch (e) {
    console.error('Error getDashboardSummary:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET NOTIFICATIONS ====================
function getNotifications(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    var notifications = [];
    
    // Coaching reminders
    var reminderData = getReminderData(sessionId);
    if (reminderData.success) {
      if (reminderData.total_overdue > 0) {
        notifications.push({
          type: 'warning',
          title: 'Coaching Overdue',
          message: reminderData.total_overdue + ' coaching melewati target date',
          link: 'reminder'
        });
      }
    }
    
    // Pencapaian alerts
    var pencapaianData = getDashboardPencapaianData({});
    if (pencapaianData.success && pencapaianData.data.metrics) {
      var m = pencapaianData.data.metrics;
      if (m.blmCair > 0) {
        notifications.push({
          type: 'info',
          title: 'Approve Belum Cair',
          message: m.blmCair + ' aplikasi menunggu pencairan',
          link: 'pencapaian'
        });
      }
    }
    
    return {
      success: true,
      notifications: notifications,
      totalUnread: notifications.length
    };
    
  } catch (e) {
    console.error('Error getNotifications:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: MARK NOTIFICATION READ ====================
function markNotificationRead(sessionId, notificationId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    // Implementation would track read status in a separate sheet
    return { success: true };
    
  } catch (e) {
    console.error('Error markNotificationRead:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== HELPER: FORMAT CURRENCY ====================
function formatCurrency(value) {
  if (!value || isNaN(value)) return 'Rp 0';
  
  var suffix = '';
  var num = parseFloat(value);
  
  if (num >= 1000000000) {
    num = num / 1000000000;
    suffix = 'M';
  } else if (num >= 1000000) {
    num = num / 1000000;
    suffix = 'jt';
  } else if (num >= 1000) {
    num = num / 1000;
    suffix = 'rb';
  }
  
  return 'Rp ' + num.toFixed(1).replace('.0', '') + suffix;
}

//==================== HELPER: FORMAT NUMBER ====================
function formatNumber(value) {
  if (!value || isNaN(value)) return '0';
  return parseFloat(value).toLocaleString('id-ID');
}

//==================== HELPER: GET MONTH NAME ====================
function getMonthName(monthIndex) {
  var months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return months[monthIndex] || '';
}

//==================== HELPER: GET PERIODE OPTIONS ====================
function getPeriodeOptions() {
  var options = [];
  var now = new Date();
  var currentYear = now.getFullYear();
  
  for (var year = currentYear; year >= currentYear - 2; year--) {
    for (var month = 11; month >= 0; month--) {
      var label = getMonthName(month) + ' ' + year;
      var value = year + '-' + String(month + 1).padStart(2, '0');
      options.push({ label: label, value: value });
    }
  }
  
  return options;
}

//==================== API: GET INITIAL DATA FOR LANDING ====================
function getLandingData(sessionId) {
  try {
    var user = checkSession(sessionId);
    if (!user) return { success: false, message: 'Session expired' };
    
    // Default landing page = Pencapaian Dashboard
    var pencapaianData = getDashboardPencapaianData({});
    
    return {
      success: true,
      landingPage: 'pencapaian',
      user: user,
      canUpload: canUploadData(user.npk, user.jabatan),
      pencapaian: pencapaianData.success ? pencapaianData.data : getEmptyDashboardData()
    };
    
  } catch (e) {
    console.error('Error getLandingData:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: CLEAR ALL SESSIONS ====================
function clearAllSessions() {
  try {
    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    
    for (var key in allProps) {
      if (key.startsWith('SESSION_')) {
        props.deleteProperty(key);
      }
    }
    
    return { success: true, message: 'All sessions cleared' };
    
  } catch (e) {
    console.error('Error clearAllSessions:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== API: GET SYSTEM STATUS ====================
function getSystemStatus() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var sheetNames = [];
    
    for (var i = 0; i < sheets.length; i++) {
      sheetNames.push(sheets[i].getName());
    }
    
    return {
      success: true,
      sheets: sheetNames,
      totalSheets: sheets.length,
      url: ss.getUrl(),
      timestamp: new Date().toISOString()
    };
    
  } catch (e) {
    console.error('Error getSystemStatus:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== TRIGGER: ON OPEN ====================
function onOpen() {
  setupSheets();
  
  SpreadsheetApp.getUi()
    .createMenu('🎯 Coaching System')
    .addItem('Buka Aplikasi', 'showWebApp')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Clear Sessions', 'clearAllSessions')
    .addToUi();
}

//==================== HELPER: SHOW WEB APP ====================
function showWebApp() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('Buka URL ini di browser:\n\n' + url);
}

//==================== API: TEST CONNECTION ====================
function testConnection() {
  return {
    success: true,
    message: 'Connected to Coaching Monitoring System v8.0 + Pencapaian Module',
    timestamp: new Date().toISOString(),
    version: '8.1.0'
  };
}
