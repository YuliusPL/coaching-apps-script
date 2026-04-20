//==================== MANAGEMENT USER CRUD ====================

function getAllUsers() {
  try {
    const user = getUserSession();
    if (!user || getUserLevel(user.cl) < 4) return { success: false, message: 'Akses ditolak' };

    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_USERS);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_USERS);
      sheet.appendRow(['NPK', 'Nama', 'Role', 'Status', 'Created_At', 'Updated_At', 'Created_By']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
      return { success: true, data: [] };
    }

    const data = sheet.getDataRange().getValues();
    const users = [];

    for (let i = 1; i < data.length; i++) {
      users.push({
        npk: String(data[i][0]), nama: data[i][1] || '', role: data[i][2] || 'User',
        status: data[i][3] || 'Active', created_at: data[i][4] || '',
        updated_at: data[i][5] || '', created_by: data[i][6] || ''
      });
    }
    return { success: true, data: users };
  } catch (e) { 
    console.error('Error in getAllUsers:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

function saveUser(userData) {
  try {
    const user = getUserSession();
    if (!user || getUserLevel(user.cl) < 4) return { success: false, message: 'Akses ditolak' };

    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_USERS);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_USERS);
      sheet.appendRow(['NPK', 'Nama', 'Role', 'Status', 'Created_At', 'Updated_At', 'Created_By']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    }

    const data = sheet.getDataRange().getValues();
    const timestamp = getCurrentTimestamp();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userData.npk)) {
        sheet.getRange(i + 1, 2, 1, 5).setValues([[
          userData.nama, userData.role, userData.status, data[i][4], timestamp
        ]]);
        return { success: true, message: 'User diperbarui' };
      }
    }

    sheet.appendRow([
      userData.npk, userData.nama, userData.role || 'User',
      userData.status || 'Active', timestamp, timestamp, user.nama
    ]);
    return { success: true, message: 'User ditambahkan' };
  } catch (e) { 
    console.error('Error in saveUser:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

function deleteUser(npk) {
  try {
    const user = getUserSession();
    if (!user || getUserLevel(user.cl) < 4) return { success: false, message: 'Akses ditolak' };

    const sheet = getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(npk)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'User dihapus' };
      }
    }
    return { success: false, message: 'User tidak ditemukan' };
  } catch (e) { 
    console.error('Error in deleteUser:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

//==================== REMINDER & REPORT (PLACEHOLDER) ====================

function sendReminderEmail(data) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const ss = getSpreadsheet();
    let logSheet = ss.getSheetByName(SHEET_REMINDER_LOG);
    
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_REMINDER_LOG);
      logSheet.appendRow(['reminder_id', 'coaching_id', 'coach_npk', 'coachee_npk', 'reminder_date', 'channel', 'status', 'sent_timestamp', 'message_content', 'created_at']);
      logSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    }

    const reminderId = generateUUID();
    const timestamp = getCurrentTimestamp();

    logSheet.appendRow([
      reminderId, data.coaching_id || '', data.coach_npk || '',
      data.coachee_npk || '', data.reminder_date || '', 'Email',
      'Sent', timestamp, data.message || 'Reminder email dummy', timestamp
    ]);

    return {
      success: true,
      message: 'Email reminder logged (dummy)',
      reminder_id: reminderId,
      log: {
        coaching_id: data.coaching_id,
        coach_npk: data.coach_npk,
        coachee_npk: data.coachee_npk,
        channel: 'Email',
        status: 'Sent',
        timestamp: timestamp
      }
    };
  } catch (e) { 
    console.error('Error in sendReminderEmail:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

function sendReminderWA(data) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const ss = getSpreadsheet();
    let logSheet = ss.getSheetByName(SHEET_REMINDER_LOG);
    
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_REMINDER_LOG);
      logSheet.appendRow(['reminder_id', 'coaching_id', 'coach_npk', 'coachee_npk', 'reminder_date', 'channel', 'status', 'sent_timestamp', 'message_content', 'created_at']);
      logSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    }

    const reminderId = generateUUID();
    const timestamp = getCurrentTimestamp();

    logSheet.appendRow([
      reminderId, data.coaching_id || '', data.coach_npk || '',
      data.coachee_npk || '', data.reminder_date || '', 'WhatsApp',
      'Sent', timestamp, data.message || 'Reminder WA dummy', timestamp
    ]);

    return {
      success: true,
      message: 'WhatsApp reminder logged (dummy)',
      reminder_id: reminderId,
      log: {
        coaching_id: data.coaching_id,
        coach_npk: data.coach_npk,
        coachee_npk: data.coachee_npk,
        channel: 'WhatsApp',
        status: 'Sent',
        timestamp: timestamp
      }
    };
  } catch (e) { 
    console.error('Error in sendReminderWA:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

//==================== DOWNLOAD / EXPORT (PLACEHOLDER) ====================

function exportCoachingToExcel(filters) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const headerSheet = getSheetByName(SHEET_COACHING_HEADER);
    const detailSheet = getSheetByName(SHEET_COACHING_DETAIL);
    const karyawanSheet = getSheetByName(SHEET_KARYAWAN);

    const headerData = headerSheet.getDataRange().getValues();
    const karyawanData = karyawanSheet.getDataRange().getValues();

    const karyawanMap = {};
    for (let i = 1; i < karyawanData.length; i++) {
      karyawanMap[String(karyawanData[i][0])] = karyawanData[i][1] || '';
    }

    const exportData = [];
    exportData.push([
      'Coaching ID', 'Coach NPK', 'Coach Nama', 'Coachee NPK', 'Coachee Nama',
      'Cabang', 'Root Cause', 'Topic', 'Status', 'Created Date', 'Target Date'
    ]);

    for (let i = 1; i < headerData.length; i++) {
      const coachNpk = String(headerData[i][1]);
      const coacheeNpk = String(headerData[i][2]);
      
      if (filters && filters.cabang && headerData[i][3] !== filters.cabang) continue;
      if (filters && filters.status && headerData[i][6] !== filters.status) continue;

      exportData.push([
        headerData[i][0], coachNpk, karyawanMap[coachNpk] || '',
        coacheeNpk, karyawanMap[coacheeNpk] || '',
        headerData[i][3], headerData[i][4], headerData[i][5],
        headerData[i][6], headerData[i][7], headerData[i][8] || ''
      ]);
    }

    const newSs = SpreadsheetApp.create('Export_Coaching_' + getCurrentTimestamp().replace(/[: ]/g, '_'));
    const newSheet = newSs.getActiveSheet();
    newSheet.getRange(1, 1, exportData.length, exportData[0].length).setValues(exportData);
    newSheet.getRange(1, 1, 1, exportData[0].length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');

    return {
      success: true,
      message: 'Export coaching berhasil',
      download_url: newSs.getUrl(),
      record_count: exportData.length - 1
    };
  } catch (e) { 
    console.error('Error in exportCoachingToExcel:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

function exportKaryawanToExcel() {
  try {
    const user = getUserSession();
    if (!user || getUserLevel(user.cl) < 4) return { success: false, message: 'Akses ditolak' };

    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();

    const newSs = SpreadsheetApp.create('Export_Karyawan_' + getCurrentTimestamp().replace(/[: ]/g, '_'));
    const newSheet = newSs.getActiveSheet();
    newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    newSheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');

    return {
      success: true,
      message: 'Export karyawan berhasil',
      download_url: newSs.getUrl(),
      record_count: data.length - 1
    };
  } catch (e) { 
    console.error('Error in exportKaryawanToExcel:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

//==================== UTILITY API ====================

function getAvailableCoachees(cabang) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    const coachees = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === cabang) {
        coachees.push({ npk: String(data[i][0]), nama: data[i][1] || '', jabatan: data[i][2] || '' });
      }
    }
    return { success: true, data: coachees };
  } catch (e) { 
    console.error('Error in getAvailableCoachees:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

function getWeeksList() {
  try {
    const weeks = [];
    const now = new Date();
    const currentYear = now.getFullYear();
    for (let i = 1; i <= 52; i++) {
      weeks.push(currentYear + '-W' + i.toString().padStart(2, '0'));
    }
    return { success: true, data: weeks };
  } catch (e) { 
    console.error('Error in getWeeksList:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

function getAtasanList() {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    const atasanList = [];

    for (let i = 1; i < data.length; i++) {
      const nama = data[i][1];
      const npk = String(data[i][0]);
      if (nama && npk) atasanList.push({ npk: npk, nama: nama });
    }
    return { success: true, data: atasanList };
  } catch (e) { 
    console.error('Error in getAtasanList:', e);
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; 
  }
}

//==================== SETUP ====================

function setupSpreadsheet() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(function(s) { return s.getName(); });
    let messages = [];

    if (sheetNames.indexOf(SHEET_KARYAWAN) === -1) {
      const sheet = ss.insertSheet(SHEET_KARYAWAN);
      sheet.appendRow(['NPK', 'Nama', 'Jabatan', 'CL', 'Cabang', 'Region', 'Atasan_NPK', 'Email', 'No_HP']);
      messages.push('Sheet "DB_KARYAWAN" dibuat');
    }
    if (sheetNames.indexOf(SHEET_USERS) === -1) {
      const sheet = ss.insertSheet(SHEET_USERS);
      sheet.appendRow(['NPK', 'Nama', 'Role', 'Status', 'Created_At', 'Updated_At', 'Created_By']);
      messages.push('Sheet "DB_USERS" dibuat');
    }
    if (sheetNames.indexOf(SHEET_COACHING_HEADER) === -1) {
      const sheet = ss.insertSheet(SHEET_COACHING_HEADER);
      sheet.appendRow(['coaching_id', 'coach_npk', 'coachee_npk', 'cabang', 'root_cause', 'topic', 'status', 'created_date', 'target_date']);
      messages.push('Sheet "COACHING_HEADER" dibuat');
    }
    if (sheetNames.indexOf(SHEET_COACHING_DETAIL) === -1) {
      const sheet = ss.insertSheet(SHEET_COACHING_DETAIL);
      sheet.appendRow(['detail_id', 'coaching_id', 'week', 'action', 'how', 'target_date', 'result', 'feedback', 'update_date', 'target_nominal', 'target_satuan']);
      messages.push('Sheet "COACHING_DETAIL" dibuat');
    }
    if (sheetNames.indexOf(SHEET_REMINDER_LOG) === -1) {
      const sheet = ss.insertSheet(SHEET_REMINDER_LOG);
      sheet.appendRow(['reminder_id', 'coaching_id', 'coach_npk', 'coachee_npk', 'reminder_date', 'channel', 'status', 'sent_timestamp', 'message_content', 'created_at']);
      messages.push('Sheet "REMINDER_LOG" dibuat');
    }

    return { success: true, message: messages.join(', ') || 'Semua sheet sudah ada' };
  } catch (e) { 
    console.error('Error in setupSpreadsheet:', e);
    return { success: false, message: 'Setup failed: ' + e.toString() }; 
  }
}

//==================== OWNER DETECTION ====================

function isOwner(npk) {
  return String(npk).trim() === '2510285';
}

//==================== END OF FILE ====================
