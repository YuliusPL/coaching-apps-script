/**
 * COACHING MONITORING SYSTEM v5.0
 * Full Package dengan Admin CRUD + Coaching Dashboard
 */

const SPREADSHEET_ID = '17wQagjAHZPyvWr5YJ0yhnTY5V5XPwm3NQB3eMIFIaIU';
const SHEET_KARYAWAN = 'DB_KARYAWAN';
const SHEET_COACHING_HEADER = 'COACHING_HEADER';
const SHEET_COACHING_DETAIL = 'COACHING_DETAIL';
const SHEET_REMINDER_LOG = 'REMINDER_LOG';

const CL_DIREKSI = 0, CL_HRD = 1, CL_KONSULTAN = 2, CL_LEVEL_3 = 3, CL_LEVEL_2 = 4, CL_LEVEL_1 = 5;
const STATUS_OPEN = 'OPEN', STATUS_ON_PROGRESS = 'ON PROGRESS', STATUS_DONE = 'DONE', STATUS_OVERDUE = 'OVERDUE';

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

function getCurrentPeriod(periodType) {
  const now = new Date();
  if (periodType === 'monthly') return now.getFullYear() + '-' + (now.getMonth() + 1).toString().padStart(2, '0');
  return getCurrentISOWeek();
}

function getUserLevel(cl) {
  if (cl <= CL_KONSULTAN) return 4;
  if (cl === CL_LEVEL_3) return 3;
  if (cl === CL_LEVEL_2) return 2;
  return 1;
}

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

function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    template.userEmail = Session.getActiveUser().getEmail();
    return template.evaluate()
      .setTitle('Coaching Monitoring System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    return ContentService.createTextOutput('Error loading application: ' + e.toString());
  }
}
//==================== AUTHENTICATION ====================

function loginUser(credentials) {
  try {
    if (!credentials || !credentials.npk) return { success: false, message: 'NPK harus diisi' };
    const npk = credentials.npk;
    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: 'Database karyawan kosong' };
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(npk).trim()) {
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
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
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

//==================== ACCESSIBLE EMPLOYEES ====================

function getAccessibleEmployees(user) {
  try {
    if (!user) return { employees: [], cabangs: [] };
    const level = getUserLevel(user.cl);
    const karyawanSheet = getSheetByName(SHEET_KARYAWAN);
    const allKaryawan = karyawanSheet.getDataRange().getValues();
    let accessibleEmployees = [];
    let accessibleCabangs = new Set();

    const makeEmp = function(row) {
      return {
        npk: String(row[0]), nama: row[1] || '', jabatan: row[2] || '',
        cl: parseInt(row[3]) || 5, cabang: row[4] || '', region: row[5] || '',
        atasan_npk: row[6] || ''
      };
    };

    if (level === 4) {
      for (let i = 1; i < allKaryawan.length; i++) {
        const emp = makeEmp(allKaryawan[i]);
        accessibleEmployees.push(emp);
        if (emp.cabang) accessibleCabangs.add(emp.cabang);
      }
    } else if (level === 3) {
      for (let i = 1; i < allKaryawan.length; i++) {
        const emp = makeEmp(allKaryawan[i]);
        if (emp.region === user.region) {
          accessibleEmployees.push(emp);
          if (emp.cabang) accessibleCabangs.add(emp.cabang);
        }
      }
    } else if (level === 2) {
      const subordinates = getSubordinates(user.npk, true);
      subordinates.forEach(function(sub) {
        accessibleEmployees.push(sub);
        if (sub.cabang) accessibleCabangs.add(sub.cabang);
      });
      for (let i = 1; i < allKaryawan.length; i++) {
        const emp = makeEmp(allKaryawan[i]);
        if (emp.cabang === user.cabang && !accessibleEmployees.find(function(e) { return e.npk === emp.npk; })) {
          accessibleEmployees.push(emp);
          accessibleCabangs.add(emp.cabang);
        }
      }
    } else {
      accessibleEmployees.push({
        npk: user.npk, nama: user.nama, jabatan: user.jabatan,
        cl: user.cl, cabang: user.cabang, region: user.region, atasan_npk: user.atasan_npk
      });
      if (user.cabang) accessibleCabangs.add(user.cabang);
    }

    return { employees: accessibleEmployees, cabangs: Array.from(accessibleCabangs).sort() };
  } catch (e) { return { employees: [], cabangs: [] }; }
}

function getSubordinates(npk, includeSelf) {
  try {
    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    const subordinates = [];
    
    function findSubordinates(atasanNPK) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][6]) === String(atasanNPK)) {
          subordinates.push({
            npk: String(data[i][0]), nama: data[i][1] || '', jabatan: data[i][2] || '',
            cl: parseInt(data[i][3]) || 5, cabang: data[i][4] || '', region: data[i][5] || '',
            atasan_npk: data[i][6] || ''
          });
          findSubordinates(data[i][0]);
        }
      }
    }
    
    if (includeSelf) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(npk)) {
          subordinates.push({
            npk: String(data[i][0]), nama: data[i][1] || '', jabatan: data[i][2] || '',
            cl: parseInt(data[i][3]) || 5, cabang: data[i][4] || '', region: data[i][5] || '',
            atasan_npk: data[i][6] || ''
          });
          break;
        }
      }
    }
    findSubordinates(npk);
    return subordinates;
  } catch (e) { return []; }
}
//==================== COACHING STATUS ====================

function getCoachingStatusForPeriod(coacheeNpk, period, periodType) {
  try {
    const headerSheet = getSheetByName(SHEET_COACHING_HEADER);
    const detailSheet = getSheetByName(SHEET_COACHING_DETAIL);
    const headerData = headerSheet.getDataRange().getValues();
    const detailData = detailSheet.getDataRange().getValues();

    let activeCoaching = null;
    for (let i = 1; i < headerData.length; i++) {
      if (String(headerData[i][2]) === String(coacheeNpk)) {
        const status = headerData[i][6];
        if (status === STATUS_OPEN || status === STATUS_ON_PROGRESS || status === STATUS_OVERDUE) {
          activeCoaching = {
            coaching_id: headerData[i][0], coach_npk: headerData[i][1], coachee_npk: headerData[i][2],
            cabang: headerData[i][3], root_cause: headerData[i][4], topic: headerData[i][5],
            status: status, created_date: headerData[i][7], target_date: headerData[i][8] || null
          };
          break;
        }
      }
    }

    if (!activeCoaching) {
      return { status: 'BELUM_COACHING', coaching_id: null, hasUpdateInPeriod: false, details: [] };
    }

    let hasUpdateInPeriod = false;
    const details = [];
    for (let i = 1; i < detailData.length; i++) {
      if (detailData[i][1] === activeCoaching.coaching_id) {
        const detailWeek = detailData[i][2];
        const result = detailData[i][6];
        const feedback = detailData[i][7];
        const detailPeriod = periodType === 'monthly' ? detailWeek.substring(0, 7) : detailWeek;
        if (detailPeriod === period && (result || feedback)) hasUpdateInPeriod = true;
        details.push({
          detail_id: detailData[i][0], week: detailWeek, action: detailData[i][3],
          how: detailData[i][4], target_date: detailData[i][5], result: result,
          feedback: feedback, update_date: detailData[i][8], target_nominal: detailData[i][9],
          target_satuan: detailData[i][10]
        });
      }
    }

    let finalStatus = activeCoaching.status;
    if (activeCoaching.target_date && activeCoaching.status !== STATUS_DONE) {
      if (new Date(activeCoaching.target_date) < new Date()) finalStatus = STATUS_OVERDUE;
    }

    return {
      status: hasUpdateInPeriod ? 'SUDAH_UPDATE' : 'BELUM_UPDATE',
      coaching_id: activeCoaching.coaching_id, coaching_status: finalStatus,
      hasUpdateInPeriod: hasUpdateInPeriod, root_cause: activeCoaching.root_cause,
      topic: activeCoaching.topic, target_date: activeCoaching.target_date,
      details: details.sort(function(a, b) { return a.week.localeCompare(b.week); })
    };
  } catch (e) {
    return { status: 'BELUM_COACHING', coaching_id: null, hasUpdateInPeriod: false, details: [] };
  }
}

//==================== INITIAL DATA (OPTIMIZED) ====================

function getInitialData() {
  try {
    const user = getUserSession();
    if (!user) return { success: false, message: 'Sesi habis, silakan login kembali' };

    const accessible = getAccessibleEmployees(user);
    const employees = accessible.employees;
    const cabangs = accessible.cabangs;
    const currentPeriod = getCurrentPeriod('weekly');
    
    if (!employees || employees.length === 0) {
      return {
        success: true,
        data: {
          totalCoaching: 0, belumCoaching: 0, sudahCoachingBelumUpdate: 0, sudahCoachingSudahUpdate: 0,
          currentPeriod: currentPeriod, user: user, userLevel: getUserLevel(user.cl),
          cabangs: cabangs || [], employees: [], monitoring: { belumCoaching: [], belumUpdate: [] }
        }
      };
    }

    let totalCoaching = 0, belumCoaching = 0, sudahCoachingBelumUpdate = 0, sudahCoachingSudahUpdate = 0;
    const employeeList = [], belumCoachingList = [], belumUpdateList = [];

    let headerData = [], detailData = [];
    try {
      const headerSheet = getSheetByName(SHEET_COACHING_HEADER);
      headerData = headerSheet.getDataRange().getValues();
    } catch (e) { headerData = []; }
    
    try {
      const detailSheet = getSheetByName(SHEET_COACHING_DETAIL);
      detailData = detailSheet.getDataRange().getValues();
    } catch (e) { detailData = []; }

    for (let e = 0; e < employees.length; e++) {
      const emp = employees[e];
      let activeCoaching = null;
      for (let i = 1; i < headerData.length; i++) {
        if (String(headerData[i][2]) === String(emp.npk)) {
          const status = headerData[i][6];
          if (status === STATUS_OPEN || status === STATUS_ON_PROGRESS || status === STATUS_OVERDUE) {
            activeCoaching = {
              coaching_id: headerData[i][0], coach_npk: headerData[i][1], status: status,
              root_cause: headerData[i][4], topic: headerData[i][5], target_date: headerData[i][8] || null
            };
            break;
          }
        }
      }

      let coachingStatus = 'BELUM_COACHING', hasUpdate = false, finalStatus = null, coachingId = null;
      if (activeCoaching) {
        coachingId = activeCoaching.coaching_id;
        finalStatus = activeCoaching.status;
        if (activeCoaching.target_date) {
          if (new Date(activeCoaching.target_date) < new Date() && activeCoaching.status !== STATUS_DONE) {
            finalStatus = STATUS_OVERDUE;
          }
        }
        for (let i = 1; i < detailData.length; i++) {
          if (detailData[i][1] === activeCoaching.coaching_id) {
            const detailWeek = detailData[i][2];
            const result = detailData[i][6];
            const feedback = detailData[i][7];
            if (detailWeek === currentPeriod && (result || feedback)) { hasUpdate = true; break; }
          }
        }
        coachingStatus = hasUpdate ? 'SUDAH_UPDATE' : 'BELUM_UPDATE';
      }

      if (coachingStatus === 'BELUM_COACHING') {
        belumCoaching++;
        belumCoachingList.push(Object.assign({}, emp, { status: coachingStatus }));
      } else if (coachingStatus === 'BELUM_UPDATE') {
        totalCoaching++; sudahCoachingBelumUpdate++;
        belumUpdateList.push(Object.assign({}, emp, { status: coachingStatus, coaching_id: coachingId, root_cause: activeCoaching.root_cause, topic: activeCoaching.topic }));
      } else if (coachingStatus === 'SUDAH_UPDATE') {
        totalCoaching++; sudahCoachingSudahUpdate++;
      }

      employeeList.push(Object.assign({}, emp, {
        coaching_id: coachingId, status: coachingStatus, coaching_status: finalStatus,
        root_cause: activeCoaching ? activeCoaching.root_cause : '', topic: activeCoaching ? activeCoaching.topic : '',
        target_date: activeCoaching ? activeCoaching.target_date : null
      }));
    }

    return {
      success: true,
      data: {
        totalCoaching: totalCoaching, belumCoaching: belumCoaching,
        sudahCoachingBelumUpdate: sudahCoachingBelumUpdate, sudahCoachingSudahUpdate: sudahCoachingSudahUpdate,
        currentPeriod: currentPeriod, user: user, userLevel: getUserLevel(user.cl),
        cabangs: cabangs, employees: employeeList,
        monitoring: { belumCoaching: belumCoachingList, belumUpdate: belumUpdateList }
      }
    };
  } catch (e) {
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() };
  }
}

function getDashboardData(cabangFilter, periodType) {
  return getInitialData();
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
      if (new Date(header.target_date) < new Date()) displayStatus = STATUS_OVERDUE;
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
        details: details, timeline: details.map(function(d, index) { return Object.assign({}, d, { step: index + 1 }); })
      }
    };
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
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
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
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
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
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

    const headerData = headerSheet.getDataRange().getValues();
    for (let i = 1; i < headerData.length; i++) {
      if (headerData[i][0] === data.coaching_id && headerData[i][6] === STATUS_OPEN) {
        headerSheet.getRange(i + 1, 7).setValue(STATUS_ON_PROGRESS);
        break;
      }
    }

    return { success: true, detail_id: detailId, message: 'Update mingguan berhasil disimpan' };
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
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

        if (user.npk !== coachNpk && user.npk !== atasanNpk && getUserLevel(user.cl) < 4) {
          return { success: false, message: 'Anda tidak berhak menyelesaikan coaching ini' };
        }

        sheet.getRange(i + 1, 7).setValue(STATUS_DONE);
        return { success: true, message: 'Coaching berhasil diselesaikan.' };
      }
    }
    return { success: false, message: 'Coaching tidak ditemukan' };
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
}

//==================== ADMIN KARYAWAN CRUD ====================

function getAllKaryawan() {
  try {
    const user = getUserSession();
    if (!user || getUserLevel(user.cl) < 4) return { success: false, message: 'Akses ditolak' };

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
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
}

function saveKaryawan(data) {
  try {
    const user = getUserSession();
    if (!user || getUserLevel(user.cl) < 4) return { success: false, message: 'Akses ditolak' };

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
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
}

function deleteKaryawan(npk) {
  try {
    const user = getUserSession();
    if (!user || getUserLevel(user.cl) < 4) return { success: false, message: 'Akses ditolak' };

    const sheet = getSheetByName(SHEET_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(npk)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data dihapus' };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
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
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
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
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
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
  } catch (e) { return { success: false, message: 'Terjadi kesalahan: ' + e.toString() }; }
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
  } catch (e) { return { success: false, message: 'Setup failed: ' + e.toString() }; }
}
