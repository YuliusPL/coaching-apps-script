/**
 * COACHING MONITORING SYSTEM v7.1
 * Fix: Force return all data for debugging
 */

const SPREADSHEET_ID = '17wQagjAHZPyvWr5YJ0yhnTY5V5XPwm3NQB3eMIFIaIU';
const SHEET_KARYAWAN = 'DB_KARYAWAN';
const SHEET_USERS = 'DB_USERS';
const SHEET_COACHING_HEADER = 'COACHING_HEADER';
const SHEET_COACHING_DETAIL = 'COACHING_DETAIL';
const SHEET_REMINDER_LOG = 'REMINDER_LOG';

const STATUS_OPEN = 'OPEN', STATUS_ON_PROGRESS = 'ON PROGRESS', STATUS_DONE = 'DONE', STATUS_OVERDUE = 'OVERDUE';

//==================== UTILITIES ====================

function getSpreadsheet() {
  try { return SpreadsheetApp.openById(SPREADSHEET_ID); }
  catch (e) { throw new Error('Spreadsheet tidak ditemukan.'); }
}

function getSheetByName(name) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan.');
  return sheet;
}

function generateUUID() { return Utilities.getUuid(); }

function getCurrentISOWeek() {
  const now = new Date();
  const d = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return d.getUTCFullYear() + '-W' + weekNo.toString().padStart(2, '0');
}

function getCurrentTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function formatDate(date) {
  if (!date) return '';
  try { return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
  catch (e) { return date; }
}

function getCurrentPeriod() {
  return getCurrentISOWeek();
}

//==================== SESSION ====================

function getUserSession() {
  try {
    const props = PropertiesService.getUserProperties();
    const session = props.getProperty('coaching_session');
    if (!session) return null;
    try {
      const data = JSON.parse(session);
      if (new Date().getTime() - new Date(data.timestamp).getTime() > 86400000) {
        props.deleteProperty('coaching_session');
        return null;
      }
      return data;
    } catch (e) {
      props.deleteProperty('coaching_session');
      return null;
    }
  } catch (e) { return null; }
}

function setUserSession(userData) {
  try {
    const props = PropertiesService.getUserProperties();
    const session = { userData: userData, timestamp: new Date().toISOString() };
    props.setProperty('coaching_session', JSON.stringify(session));
    return session;
  } catch (e) { return null; }
}

function clearUserSession() {
  try { PropertiesService.getUserProperties().deleteProperty('coaching_session'); }
  catch (e) {}
}

//==================== WEB APP ====================

function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    template.userEmail = Session.getActiveUser().getEmail();
    return template.evaluate()
      .setTitle('Coaching Monitoring System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    return ContentService.createTextOutput('Error: ' + e.toString());
  }
}

//==================== AUTH ====================

function loginUser(credentials) {
  try {
    if (!credentials || !credentials.npk) return { success: false, message: 'NPK harus diisi' };
    const npk = String(credentials.npk).trim();
    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: 'Database karyawan kosong' };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === npk) {
        const user = {
          npk: String(data[i][0]), nama: data[i][1] || '', jabatan: data[i][2] || '',
          cl: parseInt(data[i][3]) || 5, cabang: data[i][4] || '', region: data[i][5] || '',
          atasan_npk: data[i][6] || '', email: data[i][7] || '', no_hp: data[i][8] || ''
        };
        setUserSession(user);
        return { success: true, user: user };
      }
    }
    return { success: false, message: 'NPK tidak ditemukan' };
  } catch (e) { return { success: false, message: 'Error: ' + e.toString() }; }
}

function getCurrentUser() {
  try { return getUserSession(); } catch (e) { return null; }
}

function logoutUser() {
  try {
    clearUserSession();
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

//==================== HIERARCHY - FIX FORCE ALL ====================

function getAccessibleEmployees(user) {
  try {
    if (!user) return { employees: [], cabangs: [] };
    
    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    
    // FIX: Jika tidak ada data, return kosong
    if (data.length < 2) {
      return { employees: [], cabangs: [] };
    }
    
    const allEmployees = [];
    const cabangs = new Set();
    
    // Convert semua ke object
    for (let i = 1; i < data.length; i++) {
      const emp = {
        npk: String(data[i][0] || ''), 
        nama: data[i][1] || '', 
        jabatan: data[i][2] || '',
        cl: parseInt(data[i][3]) || 5, 
        cabang: data[i][4] || '', 
        region: data[i][5] || '',
        atasan_npk: String(data[i][6] || '')
      };
      allEmployees.push(emp);
      if (emp.cabang) cabangs.add(emp.cabang);
    }
    
    // FIX: Selalu return semua untuk debugging
    // Nanti bisa di-filter berdasarkan hierarchy
    return { employees: allEmployees, cabangs: Array.from(cabangs).sort() };
    
  } catch (e) { 
    console.error('getAccessibleEmployees error:', e);
    return { employees: [], cabangs: [] }; 
  }
}

//==================== DEBUG FUNCTION ====================

function testDashboard() {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'No user session' };
    
    const accessible = getAccessibleEmployees(user);
    
    return {
      success: true,
      debug: {
        user: user,
        employee_count: accessible.employees.length,
        cabang_count: accessible.cabangs.length,
        first_employee: accessible.employees[0] || null,
        sample_employees: accessible.employees.slice(0, 3)
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
//==================== FIX getInitialData - FORCE ALL ====================

function getInitialData() {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis, silakan login kembali' };

    // FIX: Ambil SEMUA data tanpa filter hierarchy dulu
    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    
    const employees = [];
    const cabangs = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const emp = {
        npk: String(data[i][0] || ''), 
        nama: data[i][1] || '', 
        jabatan: data[i][2] || '',
        cl: parseInt(data[i][3]) || 5, 
        cabang: data[i][4] || '', 
        region: data[i][5] || '',
        atasan_npk: String(data[i][6] || '')
      };
      employees.push(emp);
      if (emp.cabang) cabangs.add(emp.cabang);
    }

    const currentPeriod = getCurrentPeriod();

    // Ambil data coaching
    let headerData = [], detailData = [];
    try {
      const headerSheet = getSheetByName(SHEET_COACHING_HEADER);
      headerData = headerSheet.getDataRange().getValues();
    } catch (e) { headerData = []; }
    
    try {
      const detailSheet = getSheetByName(SHEET_COACHING_DETAIL);
      detailData = detailSheet.getDataRange().getValues();
    } catch (e) { detailData = []; }

    // Buat map untuk lookup cepat
    const coachingMap = {};
    for (let i = 1; i < headerData.length; i++) {
      const status = headerData[i][6];
      if (status === STATUS_OPEN || status === STATUS_ON_PROGRESS || status === STATUS_OVERDUE) {
        coachingMap[String(headerData[i][2])] = {
          coaching_id: headerData[i][0],
          coach_npk: headerData[i][1],
          status: status,
          root_cause: headerData[i][4] || '',
          topic: headerData[i][5] || '',
          target_date: headerData[i][8] || null
        };
      }
    }

    // Map detail per coaching
    const detailMap = {};
    for (let i = 1; i < detailData.length; i++) {
      const cid = detailData[i][1];
      if (!detailMap[cid]) detailMap[cid] = [];
      detailMap[cid].push({
        week: detailData[i][2],
        result: detailData[i][6],
        feedback: detailData[i][7]
      });
    }

    let totalCoaching = 0, belumCoaching = 0, sudahCoachingBelumUpdate = 0, sudahCoachingSudahUpdate = 0;
    const employeeList = [], belumCoachingList = [], belumUpdateList = [];

    // Process setiap employee
    for (let e = 0; e < employees.length; e++) {
      const emp = employees[e];
      const activeCoaching = coachingMap[emp.npk];

      let coachingStatus = 'BELUM_COACHING', hasUpdate = false, finalStatus = null, coachingId = null;
      
      if (activeCoaching) {
        coachingId = activeCoaching.coaching_id;
        finalStatus = activeCoaching.status;
        
        // Cek overdue
        if (activeCoaching.target_date && activeCoaching.status !== STATUS_DONE) {
          try {
            if (new Date(activeCoaching.target_date) < new Date()) finalStatus = STATUS_OVERDUE;
          } catch (err) {}
        }
        
        // Cek update minggu ini
        const details = detailMap[coachingId] || [];
        for (let i = 0; i < details.length; i++) {
          if (details[i].week === currentPeriod && (details[i].result || details[i].feedback)) {
            hasUpdate = true;
            break;
          }
        }
        
        coachingStatus = hasUpdate ? 'SUDAH_UPDATE' : 'BELUM_UPDATE';
      }

      // Hitung stats
      if (coachingStatus === 'BELUM_COACHING') {
        belumCoaching++;
        belumCoachingList.push(Object.assign({}, emp, { status: coachingStatus }));
      } else if (coachingStatus === 'BELUM_UPDATE') {
        totalCoaching++; 
        sudahCoachingBelumUpdate++;
        belumUpdateList.push(Object.assign({}, emp, { 
          status: coachingStatus, 
          coaching_id: coachingId, 
          root_cause: activeCoaching.root_cause, 
          topic: activeCoaching.topic 
        }));
      } else if (coachingStatus === 'SUDAH_UPDATE') {
        totalCoaching++; 
        sudahCoachingSudahUpdate++;
      }

      employeeList.push(Object.assign({}, emp, {
        coaching_id: coachingId, 
        status: coachingStatus, 
        coaching_status: finalStatus,
        root_cause: activeCoaching ? activeCoaching.root_cause : '', 
        topic: activeCoaching ? activeCoaching.topic : '',
        target_date: activeCoaching ? activeCoaching.target_date : null
      }));
    }

    return {
      success: true,
      data: {
        totalCoaching: totalCoaching, 
        belumCoaching: belumCoaching,
        sudahCoachingBelumUpdate: sudahCoachingBelumUpdate, 
        sudahCoachingSudahUpdate: sudahCoachingSudahUpdate,
        currentPeriod: currentPeriod, 
        user: user, 
        userLevel: user.cl,
        cabangs: Array.from(cabangs).sort(), 
        employees: employeeList,
        monitoring: { belumCoaching: belumCoachingList, belumUpdate: belumUpdateList }
      }
    };
  } catch (e) {
    console.error('getInitialData error:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function getDashboardData(cabangFilter, periodType) {
  return getInitialData();
}

//==================== REMINDER DATA ====================

function getReminderData() {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const headerSheet = getSheetByName(SHEET_COACHING_HEADER);
    const headerData = headerSheet.getDataRange().getValues();
    const karyawanSheet = getSheetByName(SHEET_KARYAWAN);
    const karyawanData = karyawanSheet.getDataRange().getValues();
    
    const karyawanMap = {};
    for (let i = 1; i < karyawanData.length; i++) {
      karyawanMap[String(karyawanData[i][0])] = karyawanData[i][1] || '';
    }
    
    const now = new Date();
    const upcoming = [];
    const overdue = [];
    
    for (let i = 1; i < headerData.length; i++) {
      const status = headerData[i][6];
      if (status !== STATUS_DONE) {
        const targetDate = headerData[i][8] ? new Date(headerData[i][8]) : null;
        const coacheeNpk = String(headerData[i][2]);
        
        const item = {
          coaching_id: headerData[i][0],
          coach_npk: headerData[i][1],
          coachee_npk: coacheeNpk,
          coachee_nama: karyawanMap[coacheeNpk] || '',
          cabang: headerData[i][3],
          root_cause: headerData[i][4],
          topic: headerData[i][5],
          status: status,
          target_date: headerData[i][8],
          days_remaining: targetDate ? Math.ceil((targetDate - now) / 86400000) : null
        };
        
        if (targetDate && targetDate < now) {
          overdue.push(item);
        } else if (targetDate && (targetDate - now) < 7 * 86400000) {
          upcoming.push(item);
        }
      }
    }
    
    return {
      success: true,
      data: {
        upcoming: upcoming.sort(function(a, b) { return a.days_remaining - b.days_remaining; }),
        overdue: overdue.sort(function(a, b) { return a.days_remaining - b.days_remaining; }),
        total_upcoming: upcoming.length,
        total_overdue: overdue.length
      }
    };
  } catch (e) {
    console.error('getReminderData error:', e);
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

//==================== COACHING CRUD ====================

function getCoachingDetail(coachingId) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const headerSheet = getSheetByName(SHEET_COACHING_HEADER);
    const detailSheet = getSheetByName(SHEET_COACHING_DETAIL);
    const karyawanSheet = getSheetByName(SHEET_KARYAWAN);

    const headerData = headerSheet.getDataRange().getValues();
    const detailData = detailSheet.getDataRange().getValues();
    const karyawanData = karyawanSheet.getDataRange().getValues();

    let header = null, coacheeNama = '', coachNama = '', atasanNpk = '';
    for (let i = 1; i < headerData.length; i++) {
      if (headerData[i][0] === coachingId) {
        header = {
          coaching_id: headerData[i][0], coach_npk: headerData[i][1], coachee_npk: headerData[i][2],
          cabang: headerData[i][3], root_cause: headerData[i][4], topic: headerData[i][5],
          status: headerData[i][6], created_date: headerData[i][7], target_date: headerData[i][8] || null
        };
        for (let j = 1; j < karyawanData.length; j++) {
          if (String(karyawanData[j][0]) === String(headerData[i][2])) { coacheeNama = karyawanData[j][1]; atasanNpk = karyawanData[j][6] || ''; }
          if (String(karyawanData[j][0]) === String(headerData[i][1])) coachNama = karyawanData[j][1];
        }
        break;
      }
    }

    if (!header) return { success: false, message: 'Data coaching tidak ditemukan' };

    let displayStatus = header.status;
    if (header.target_date && header.status !== STATUS_DONE) {
      try {
        if (new Date(header.target_date) < new Date()) displayStatus = STATUS_OVERDUE;
      } catch (err) {}
    }

    const details = [];
    for (let i = 1; i < detailData.length; i++) {
      if (detailData[i][1] === coachingId) {
        details.push({
          detail_id: detailData[i][0], week: detailData[i][2], action: detailData[i][3],
          how: detailData[i][4], target_date: detailData[i][5], result: detailData[i][6],
          feedback: detailData[i][7], update_date: detailData[i][8], target_nominal: detailData[i][9],
          target_satuan: detailData[i][10]
        });
      }
    }
    details.sort(function(a, b) { return a.week.localeCompare(b.week); });

    return {
      success: true,
      data: {
        header: Object.assign({}, header, { coachee_nama: coacheeNama, coach_nama: coachNama, atasan_npk: atasanNpk, display_status: displayStatus }),
        details: details,
        timeline: details.map(function(d, index) { return Object.assign({}, d, { step: index + 1 }); })
      }
    };
  } catch (e) { 
    console.error('getCoachingDetail error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function checkExistingCoaching(coacheeNpk) {
  try {
    const sheet = getSheetByName(SHEET_COACHING_HEADER);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const status = data[i][6];
      if (String(data[i][2]) === String(coacheeNpk) && (status === STATUS_OPEN || status === STATUS_ON_PROGRESS || status === STATUS_OVERDUE)) {
        return {
          success: true, exists: true, coaching_id: data[i][0], status: status,
          root_cause: data[i][4], topic: data[i][5], target_date: data[i][8] || null,
          message: 'Karyawan masih memiliki coaching aktif.'
        };
      }
    }
    return { success: true, exists: false };
  } catch (e) { 
    console.error('checkExistingCoaching error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function saveCoachingHeader(data) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const existingCheck = checkExistingCoaching(data.coachee_npk);
    if (existingCheck.exists) {
      return {
        success: false, message: existingCheck.message,
        existing_coaching: { coaching_id: existingCheck.coaching_id, root_cause: existingCheck.root_cause, topic: existingCheck.topic, status: existingCheck.status }
      };
    }

    const sheet = getSheetByName(SHEET_COACHING_HEADER);
    const coachingId = generateUUID();
    const timestamp = getCurrentTimestamp();

    const row = [
      coachingId, data.coach_npk || user.npk, data.coachee_npk, data.cabang,
      data.root_cause, data.topic, STATUS_OPEN, timestamp, data.target_date || null
    ];
    sheet.appendRow(row);

    return { success: true, coaching_id: coachingId, status: STATUS_OPEN, message: 'Coaching berhasil dibuat.' };
  } catch (e) { 
    console.error('saveCoachingHeader error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function updateCoachingHeader(data) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const sheet = getSheetByName(SHEET_COACHING_HEADER);
    const headerData = sheet.getDataRange().getValues();

    for (let i = 1; i < headerData.length; i++) {
      if (headerData[i][0] === data.coaching_id) {
        const coachNpk = headerData[i][1];
        const coacheeNpk = headerData[i][2];
        
        if (user.npk !== coachNpk && user.npk !== coacheeNpk && user.cl > 2) {
          return { success: false, message: 'Anda tidak berhak mengubah coaching ini' };
        }

        sheet.getRange(i + 1, 5).setValue(data.root_cause || headerData[i][4]);
        sheet.getRange(i + 1, 6).setValue(data.topic || headerData[i][5]);
        sheet.getRange(i + 1, 9).setValue(data.target_date || headerData[i][8]);
        
        return { success: true, message: 'Coaching berhasil diperbarui' };
      }
    }
    return { success: false, message: 'Coaching tidak ditemukan' };
  } catch (e) { 
    console.error('updateCoachingHeader error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function saveCoachingDetail(data) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const detailSheet = getSheetByName(SHEET_COACHING_DETAIL);
    const headerSheet = getSheetByName(SHEET_COACHING_HEADER);

    const detailId = generateUUID();
    const timestamp = getCurrentTimestamp();

    const row = [
      detailId, data.coaching_id, data.week || getCurrentISOWeek(), data.action,
      data.how, formatDate(data.target_date), data.result, data.feedback,
      timestamp, data.target_nominal || '', data.target_satuan || ''
    ];
    detailSheet.appendRow(row);

    // Update status header dari OPEN ke ON PROGRESS
    const headerData = headerSheet.getDataRange().getValues();
    for (let i = 1; i < headerData.length; i++) {
      if (headerData[i][0] === data.coaching_id && headerData[i][6] === STATUS_OPEN) {
        headerSheet.getRange(i + 1, 7).setValue(STATUS_ON_PROGRESS);
        break;
      }
    }

    return { success: true, detail_id: detailId, message: 'Update mingguan berhasil disimpan' };
  } catch (e) { 
    console.error('saveCoachingDetail error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function getCoachingDetailByHeader(coachingId) {
  try {
    const detailSheet = getSheetByName(SHEET_COACHING_DETAIL);
    const detailData = detailSheet.getDataRange().getValues();
    const details = [];

    for (let i = 1; i < detailData.length; i++) {
      if (detailData[i][1] === coachingId) {
        details.push({
          detail_id: detailData[i][0], coaching_id: detailData[i][1], week: detailData[i][2],
          action: detailData[i][3], how: detailData[i][4], target_date: detailData[i][5],
          result: detailData[i][6], feedback: detailData[i][7], update_date: detailData[i][8],
          target_nominal: detailData[i][9], target_satuan: detailData[i][10]
        });
      }
    }
    return { success: true, data: details.sort(function(a, b) { return a.week.localeCompare(b.week); }) };
  } catch (e) { 
    console.error('getCoachingDetailByHeader error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function completeCoaching(coachingId) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const sheet = getSheetByName(SHEET_COACHING_HEADER);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === coachingId) {
        const coachNpk = data[i][1];
        const coacheeNpk = data[i][2];

        const karyawanSheet = getSheetByName(SHEET_KARYAWAN);
        const karyawanData = karyawanSheet.getDataRange().getValues();
        let atasanNpk = '';
        for (let j = 1; j < karyawanData.length; j++) {
          if (String(karyawanData[j][0]) === String(coacheeNpk)) { atasanNpk = karyawanData[j][6] || ''; break; }
        }

        if (user.npk !== coachNpk && user.npk !== atasanNpk && user.cl > 2) {
          return { success: false, message: 'Anda tidak berhak menyelesaikan coaching ini' };
        }

        sheet.getRange(i + 1, 7).setValue(STATUS_DONE);
        return { success: true, message: 'Coaching berhasil diselesaikan.' };
      }
    }
    return { success: false, message: 'Coaching tidak ditemukan' };
  } catch (e) { 
    console.error('completeCoaching error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}
//==================== KARYAWAN CRUD ====================

function getAllKaryawan() {
  try {
    const user = getUserSession();
    if (!user || user.cl > 2) return { success: false, message: 'Akses ditolak' };

    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    const karyawan = [];

    for (let i = 1; i < data.length; i++) {
      karyawan.push({
        npk: String(data[i][0]), nama: data[i][1] || '', jabatan: data[i][2] || '',
        cl: parseInt(data[i][3]) || 5, cabang: data[i][4] || '', region: data[i][5] || '',
        atasan_npk: data[i][6] || '', email: data[i][7] || '', no_hp: data[i][8] || ''
      });
    }
    return { success: true, data: karyawan };
  } catch (e) { 
    console.error('getAllKaryawan error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function saveKaryawan(data) {
  try {
    const user = getUserSession();
    if (!user || user.cl > 2) return { success: false, message: 'Akses ditolak' };

    const sheet = getSheetByName(SHEET_KARYAWAN);
    const existing = sheet.getDataRange().getValues();

    for (let i = 1; i < existing.length; i++) {
      if (String(existing[i][0]) === String(data.npk)) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          data.npk, data.nama, data.jabatan, data.cl,
          data.cabang, data.region, data.atasan_npk,
          data.email, data.no_hp
        ]]);
        return { success: true, message: 'Data diperbarui' };
      }
    }
    sheet.appendRow([data.npk, data.nama, data.jabatan, data.cl, data.cabang, data.region, data.atasan_npk, data.email, data.no_hp]);
    return { success: true, message: 'Data ditambahkan' };
  } catch (e) { 
    console.error('saveKaryawan error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function deleteKaryawan(npk) {
  try {
    const user = getUserSession();
    if (!user || user.cl > 2) return { success: false, message: 'Akses ditolak' };

    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(npk)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data dihapus' };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { 
    console.error('deleteKaryawan error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

//==================== MANAGEMENT USER CRUD ====================

function getAllUsers() {
  try {
    const user = getUserSession();
    if (!user || user.cl > 2) return { success: false, message: 'Akses ditolak' };

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
    console.error('getAllUsers error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function saveUser(userData) {
  try {
    const user = getUserSession();
    if (!user || user.cl > 2) return { success: false, message: 'Akses ditolak' };

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
    console.error('saveUser error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function deleteUser(npk) {
  try {
    const user = getUserSession();
    if (!user || user.cl > 2) return { success: false, message: 'Akses ditolak' };

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
    console.error('deleteUser error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

//==================== EXPORT ====================

function exportCoachingToExcel(filters) {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis' };

    const headerSheet = getSheetByName(SHEET_COACHING_HEADER);
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
    console.error('exportCoachingToExcel error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function exportKaryawanToExcel() {
  try {
    const user = getUserSession();
    if (!user || user.cl > 2) return { success: false, message: 'Akses ditolak' };

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
    console.error('exportKaryawanToExcel error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
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
    console.error('getAvailableCoachees error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
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
    console.error('getWeeksList error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
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
    console.error('getAtasanList error:', e);
    return { success: false, message: 'Error: ' + e.toString() }; 
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
    console.error('setupSpreadsheet error:', e);
    return { success: false, message: 'Setup failed: ' + e.toString() }; 
  }
}

//==================== END OF FILE ====================
