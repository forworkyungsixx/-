// ===============================================================================================
// Code.gs - COMPLETE FILE Rev14 - ระบบจองห้องประชุม YUNGS!XX
// Parts 1-5: Core, Auto, Check-in/out, Booking, Get Bookings & Delete
// ===============================================================================================

const ADMIN_EMAIL = 'forworkyungsixx@gmail.com';
const ADMIN_PASSWORD = '1234';
const TEST_USER_EMAIL = 'test@kkumail.com';

// ===============================
// PART 1: CORE FUNCTIONS & SETUP
// ===============================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบจองห้องประชุม YUNGS!XX')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function validateLogin(email, password) {
  try {
    const emailLower = email.toLowerCase().trim();
    
    if (emailLower === ADMIN_EMAIL.toLowerCase()) {
      if (password === ADMIN_PASSWORD) {
        return { success: true, userType: 'admin', message: 'ยินดีต้อนรับ Admin' };
      } else {
        return { success: false, message: 'รหัสผ่านไม่ถูกต้อง' };
      }
    }
    
    if (emailLower === TEST_USER_EMAIL.toLowerCase()) {
      return { success: true, userType: 'test', message: 'ยินดีต้อนรับ Test User' };
    }
    
    if (emailLower.endsWith('@kkumail.com') || emailLower.endsWith('@kku.ac.th')) {
      return { success: true, userType: 'user', message: 'ยินดีต้อนรับเข้าสู่ระบบ' };
    } else {
      return { success: false, message: 'กรุณาใช้อีเมล @kkumail.com หรือ @kku.ac.th เท่านั้น' };
    }
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function setupTimeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  ScriptApp.newTrigger('checkUpcomingEndTimes')
    .timeBased()
    .everyMinutes(1)
    .create();
  
  ScriptApp.newTrigger('autoCheckOutExpired')
    .timeBased()
    .everyMinutes(1)
    .create();
  
  ScriptApp.newTrigger('autoCancelNoCheckin')
    .timeBased()
    .everyMinutes(1)
    .create();
  
  return { success: true, message: 'ตั้งค่า Trigger สำเร็จ' };
}

function setupCheckinSystem() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet Reservations' };
    
    const lastCol = sheet.getLastColumn();
    
    if (lastCol < 14) {
      const newHeaders = ['Check-in Code', 'Check-in Time', 'Check-out Time', 'Edit Count'];
      const startCol = lastCol + 1;
      for (let i = 0; i < newHeaders.length; i++) {
        if (startCol + i <= 14) {
          sheet.getRange(1, startCol + i).setValue(newHeaders[i]);
        }
      }
      sheet.getRange(1, 1, 1, 14).setFontWeight('bold');
    }
    
    setupTimeTriggers();
    
    return { success: true, message: 'ตั้งค่าระบบ Check-in สำเร็จ' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function timeToMinutes(time) {
  if (time instanceof Date) return time.getHours() * 60 + time.getMinutes();
  if (typeof time !== 'string') return 0;
  if (time.indexOf(':') > -1) {
    const parts = time.split(':');
    if (parts.length === 2) {
      const hours = parseInt(parts[0], 10);
      const minutes = parseInt(parts[1], 10);
      if (!isNaN(hours) && !isNaN(minutes)) return hours * 60 + minutes;
    }
  }
  return 0;
}

function formatDate(date) {
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  return `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getDate().toString().padStart(2, '0')}`;
}

function formatTime(time) {
  if (time instanceof Date) {
    return Utilities.formatDate(time, Session.getScriptTimeZone(), 'HH:mm');
  } else if (typeof time === 'number') {
    return Utilities.formatDate(new Date(time), Session.getScriptTimeZone(), 'HH:mm');
  } else {
    let timeStr = String(time).trim();
    if (timeStr.indexOf(':') === -1) timeStr = timeStr + ':00';
    return timeStr;
  }
}

function saveLoginLog(loginEmail, name, room, date, startTime, action) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('LoginLog');
    if (!logSheet) {
      logSheet = ss.insertSheet('LoginLog');
      logSheet.getRange(1, 1, 1, 7).setValues([['วันที่บันทึก', 'เวลาบันทึก', 'อีเมลผู้ใช้', 'ชื่อผู้จอง', 'ห้อง', 'วันที่จอง', 'การกระทำ']]);
      logSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }
    const now = new Date();
    const recordDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const recordTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 7).setValues([[recordDate, recordTime, loginEmail, name, room, date + ' ' + startTime, action]]);
    return true;
  } catch (error) {
    console.error('Error saving login log:', error);
    return false;
  }
}

function generateCheckinCode(bookingId) {
  return Math.random().toString(36).substring(2, 8).toUpperCase();
}

// ===============================
// PART 2: AUTO FUNCTIONS
// ===============================

function checkUpcomingEndTimes() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const now = new Date();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[11] || row[12]) continue;
      
      const bookingDate = new Date(row[0]);
      const endTimeParts = formatTime(row[3]).split(':');
      const endDateTime = new Date(bookingDate);
      endDateTime.setHours(parseInt(endTimeParts[0]), parseInt(endTimeParts[1]), 0, 0);
      
      const timeDiff = endDateTime - now;
      const minutesLeft = Math.floor(timeDiff / 60000);
      
      if (minutesLeft === 5) {
        sendEndingSoonEmail(row[8], row[4], row[1], formatDate(bookingDate), formatTime(row[2]), formatTime(row[3]));
      }
    }
  } catch (error) {
    console.error('Error checking upcoming end times:', error);
  }
}

function autoCheckOutExpired() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const now = new Date();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[11] || row[12]) continue;
      
      const bookingDate = new Date(row[0]);
      const endTimeParts = formatTime(row[3]).split(':');
      const endDateTime = new Date(bookingDate);
      endDateTime.setHours(parseInt(endTimeParts[0]), parseInt(endTimeParts[1]), 0, 0);
      
      if (now >= endDateTime) {
        const rowIndex = i + 2;
        sheet.getRange(rowIndex, 13).setValue(now);
        
        saveLoginLog(row[8], row[4], row[1], formatDate(bookingDate), formatTime(row[2]), 'Auto Check-out');
        sendAutoCheckoutEmail(row[8], row[4], row[1], formatDate(bookingDate), formatTime(row[2]), formatTime(row[3]));
      }
    }
  } catch (error) {
    console.error('Error auto check-out:', error);
  }
}

function autoCancelNoCheckin() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const now = new Date();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || row[11] || row[12]) continue;
      
      const userEmail = String(row[8]).toLowerCase().trim();
      if (userEmail === TEST_USER_EMAIL.toLowerCase()) continue;
      
      const bookingDate = new Date(row[0]);
      const startTimeParts = formatTime(row[2]).split(':');
      const startDateTime = new Date(bookingDate);
      startDateTime.setHours(parseInt(startTimeParts[0]), parseInt(startTimeParts[1]), 0, 0);
      
      const timeDiff = now - startDateTime;
      const minutesPassed = Math.floor(timeDiff / 60000);
      
      if (minutesPassed > 15) {
        const rowIndex = i + 2;
        sheet.getRange(rowIndex, 13).setValue('AUTO_CANCELLED');
        
        saveLoginLog(row[8], row[4], row[1], formatDate(bookingDate), formatTime(row[2]), 'Auto Cancel - No Check-in');
        sendAutoCancelEmail(row[8], row[4], row[1], formatDate(bookingDate), formatTime(row[2]), formatTime(row[3]));
      }
    }
  } catch (error) {
    console.error('Error auto cancel:', error);
  }
}

// ===============================
// PART 3: CHECK-IN/CHECK-OUT
// ===============================

function checkIn(checkinCode, userEmail) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Reservations' };
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    let found = false;
    let rowIndex = -1;
    let rowData = null;
    const isTestUser = userEmail.toLowerCase().trim() === TEST_USER_EMAIL.toLowerCase();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[10]) continue;
      
      const savedCode = String(row[10]).toUpperCase();
      const bookingEmail = String(row[8]).toLowerCase().trim();
      
      if (savedCode === checkinCode.toUpperCase()) {
        if (bookingEmail !== userEmail.toLowerCase().trim()) {
          return { success: false, message: 'รหัสนี้ไม่ใช่ของคุณ' };
        }
        
        if (row[12] === 'AUTO_CANCELLED') {
          return { success: false, message: 'การจองนี้ถูกยกเลิกแล้ว' };
        }
        
        if (row[11]) {
          return { success: false, message: 'Check-in ไปแล้วเมื่อ ' + formatTime(row[11]) };
        }
        
        if (!isTestUser) {
          const now = new Date();
          const bookingDate = new Date(row[0]);
          const startTimeParts = formatTime(row[2]).split(':');
          const startDateTime = new Date(bookingDate);
          startDateTime.setHours(parseInt(startTimeParts[0]), parseInt(startTimeParts[1]), 0, 0);
          
          const timeDiff = now - startDateTime;
          const minutesDiff = Math.floor(timeDiff / 60000);
          
          if (minutesDiff < 0) {
            return { success: false, message: 'ยังไม่ถึงเวลาเริ่มจอง ไม่สามารถ Check-in ได้' };
          }
          
          if (minutesDiff > 15) {
            sheet.getRange(i + 2, 13).setValue('AUTO_CANCELLED');
            return { success: false, message: 'เกินเวลา Check-in แล้ว การจองถูกยกเลิกอัตโนมัติ' };
          }
        }
        
        found = true;
        rowIndex = i + 2;
        rowData = row;
        break;
      }
    }
    
    if (!found) {
      return { success: false, message: 'ไม่พบรหัส Check-in นี้ หรือรหัสไม่ถูกต้อง' };
    }
    
    const now = new Date();
    sheet.getRange(rowIndex, 12).setValue(now);
    
    saveLoginLog(userEmail, rowData[4], rowData[1], formatDate(new Date(rowData[0])), formatTime(rowData[2]), 'Check-in');
    
    return { 
      success: true, 
      message: 'Check-in สำเร็จ',
      checkinTime: Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm')
    };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function checkOut(checkinCode, userEmail) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Reservations' };
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    let found = false;
    let rowIndex = -1;
    let rowData = null;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[10]) continue;
      
      const savedCode = String(row[10]).toUpperCase();
      const bookingEmail = String(row[8]).toLowerCase().trim();
      
      if (savedCode === checkinCode.toUpperCase()) {
        if (bookingEmail !== userEmail.toLowerCase().trim() && userEmail.toLowerCase() !== ADMIN_EMAIL.toLowerCase()) {
          return { success: false, message: 'รหัสนี้ไม่ใช่ของคุณ' };
        }
        
        if (row[12] === 'AUTO_CANCELLED') {
          return { success: false, message: 'การจองนี้ถูกยกเลิกแล้ว' };
        }
        
        if (!row[11]) {
          return { success: false, message: 'กรุณา Check-in ก่อน' };
        }
        
        if (row[12] && row[12] !== 'AUTO_CANCELLED') {
          return { success: false, message: 'Check-out ไปแล้วเมื่อ ' + formatTime(row[12]) };
        }
        
        found = true;
        rowIndex = i + 2;
        rowData = row;
        break;
      }
    }
    
    if (!found) {
      return { success: false, message: 'ไม่พบรหัส Check-in นี้' };
    }
    
    const now = new Date();
    sheet.getRange(rowIndex, 13).setValue(now);
    
    saveLoginLog(userEmail, rowData[4], rowData[1], formatDate(new Date(rowData[0])), formatTime(rowData[2]), 'Check-out');
    
    return { 
      success: true, 
      message: 'Check-out สำเร็จ',
      checkoutTime: Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm')
    };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function checkInReservation(code) {
  return checkIn(code, Session.getActiveUser().getEmail());
}

function checkOutReservation(code) {
  return checkOut(code, Session.getActiveUser().getEmail());
}

// ===============================
// PART 4: BOOKING FUNCTIONS
// ===============================

function checkUserDailyBooking(email, date) {
  try {
    if (email.toLowerCase().trim() === TEST_USER_EMAIL.toLowerCase()) {
      return { hasBooking: false, bookingCount: 0 };
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return { hasBooking: false, bookingCount: 0 };
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    let bookingCount = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[8]) continue;
      
      if (row[12] === 'AUTO_CANCELLED' || row[12] === 'USER_CANCELLED') continue;
      
      const bookingDate = formatDate(new Date(row[0]));
      const bookingEmail = String(row[8]).toLowerCase().trim();
      
      if (bookingDate === date && bookingEmail === email.toLowerCase().trim()) {
        bookingCount++;
      }
    }
    
    return { hasBooking: bookingCount > 0, bookingCount: bookingCount };
  } catch (error) {
    console.error('Error checking daily booking:', error);
    return { hasBooking: false, bookingCount: 0 };
  }
}

function saveReservation(data) {
  const lock = LockService.getScriptLock();
  
  try {
    const hasLock = lock.tryLock(30000);
    
    if (!hasLock) {
      return { 
        success: false, 
        message: 'ระบบกำลังประมวลผลการจองอื่นอยู่ กรุณารอสักครู่แล้วลองใหม่อีกครั้ง' 
      };
    }
    
    if (!data.room || !data.date || !data.startTime || !data.endTime) {
      return { success: false, message: 'ข้อมูลไม่ครบถ้วน กรุณากรอกข้อมูลให้ครบ' };
    }
    
    const isTestUser = data.loginEmail.toLowerCase().trim() === TEST_USER_EMAIL.toLowerCase();
    
    if (!isTestUser) {
      const dailyCheck = checkUserDailyBooking(data.loginEmail, data.date);
      if (dailyCheck.hasBooking) {
        return { success: false, message: 'คุณได้จองห้องในวันนี้ไปแล้ว (จำกัด 1 ครั้งต่อวัน)' };
      }
      
      const startMinutes = timeToMinutes(data.startTime);
      const endMinutes = timeToMinutes(data.endTime);
      const durationMinutes = endMinutes - startMinutes;
      const maxDurationMinutes = 3 * 60;
      
      if (durationMinutes > maxDurationMinutes) {
        return { success: false, message: 'สามารถจองได้สูงสุด 3 ชั่วโมงเท่านั้น' };
      }
    }
    
    const isDuplicate = checkDuplicateBooking(data.date, data.room, data.startTime, data.endTime);
    if (isDuplicate) {
      return { 
        success: false, 
        message: 'เวลาที่เลือกมีการจองแล้ว กรุณารีเฟรชหน้าและเลือกเวลาอื่น' 
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Reservations');
    
    if (!sheet) {
      sheet = ss.insertSheet('Reservations');
      sheet.getRange(1, 1, 1, 14).setValues([
        ['วันที่', 'ชื่อห้องที่ต้องการจอง', 'เวลาเริ่มจอง', 'เวลาจบ', 'ชื่อผู้จอง', 
         'สาขาวิชา', 'รหัสนักศึกษา', 'เบอร์ติดต่อกลับ', 'อีเมลผู้จอง', 'วันที่บันทึก', 
         'Check-in Code', 'Check-in Time', 'Check-out Time', 'Edit Count']
      ]);
      sheet.getRange(1, 1, 1, 14).setFontWeight('bold');
    }
    
    const now = new Date();
    const recordDateTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const checkinCode = generateCheckinCode(sheet.getLastRow() + 1);
    
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 11).setValues([[
      data.date, data.room, data.startTime, data.endTime, data.name, 
      data.department, data.studentId, data.phone, data.loginEmail, 
      recordDateTime, checkinCode
    ]]);
    
    sheet.getRange(sheet.getLastRow(), 14).setValue(0);
    
    if (data.loginEmail) {
      saveLoginLog(data.loginEmail, data.name, data.room, data.date, data.startTime, 'จอง');
      sendBookingConfirmationEmail(data, checkinCode);
    }
    
    return { 
      success: true, 
      message: 'บันทึกการจองเรียบร้อยแล้ว และส่งอีเมลยืนยันไปที่ ' + data.loginEmail 
    };
    
  } catch (error) {
    return { 
      success: false, 
      message: 'เกิดข้อผิดพลาด: ' + error.toString() 
    };
  } finally {
    lock.releaseLock();
  }
}

function checkDuplicateBooking(date, room, startTime, endTime) {
  try {
    const reservations = getReservationsByDateAndRoom(date, room);
    const newStartMinutes = timeToMinutes(startTime);
    const newEndMinutes = timeToMinutes(endTime);
    
    if (newStartMinutes >= newEndMinutes) {
      return true;
    }
    
    for (const reservation of reservations) {
      const existingStartMinutes = timeToMinutes(reservation.startTime);
      const existingEndMinutes = timeToMinutes(reservation.endTime);
      
      if (newStartMinutes >= existingStartMinutes && newStartMinutes < existingEndMinutes) {
        return true;
      }
      
      if (newEndMinutes > existingStartMinutes && newEndMinutes <= existingEndMinutes) {
        return true;
      }
      
      if (newStartMinutes <= existingStartMinutes && newEndMinutes >= existingEndMinutes) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    console.error('Error checking duplicate:', error);
    return true;
  }
}

function getUserReservations(email) {
  return getMyBookingsWithCheckin(email);
}

function loginUser(email, password) {
  return validateLogin(email, password);
}

// ===============================
// PART 5: GET BOOKINGS & DELETE
// ===============================

function getMyBookingsWithCheckin(email) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const reservations = [];
    const now = new Date();
    const isTestUser = email.toLowerCase().trim() === TEST_USER_EMAIL.toLowerCase();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1]) continue;
      
      const bookingEmail = String(row[8]).toLowerCase().trim();
      
      if (row[12] === 'AUTO_CANCELLED' || row[12] === 'USER_CANCELLED') continue;
      
      if (bookingEmail === email.toLowerCase().trim()) {
        let shouldBeCancelled = false;
        
        if (!isTestUser && !row[11] && !row[12]) {
          const bookingDate = new Date(row[0]);
          const startTimeParts = formatTime(row[2]).split(':');
          const startDateTime = new Date(bookingDate);
          startDateTime.setHours(parseInt(startTimeParts[0]), parseInt(startTimeParts[1]), 0, 0);
          
          const timeDiff = now - startDateTime;
          const minutesPassed = Math.floor(timeDiff / 60000);
          
          if (minutesPassed > 15) {
            shouldBeCancelled = true;
            sheet.getRange(i + 2, 13).setValue('AUTO_CANCELLED');
            saveLoginLog(email, row[4], row[1], formatDate(bookingDate), formatTime(row[2]), 'Auto Cancel - No Check-in (Real-time)');
            sendAutoCancelEmail(email, row[4], row[1], formatDate(bookingDate), formatTime(row[2]), formatTime(row[3]));
          }
        }
        
        if (shouldBeCancelled) continue;
        
        reservations.push({
          rowIndex: i + 2,
          date: formatDate(new Date(row[0])),
          room: String(row[1]).trim(),
          startTime: formatTime(row[2]),
          endTime: formatTime(row[3]),
          name: row[4],
          department: row[5],
          studentId: row[6],
          phone: row[7],
          email: row[8],
          checkinCode: row[10] || '',
          checkinTime: row[11] ? formatTime(row[11]) : null,
          checkoutTime: row[12] && row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED' ? formatTime(row[12]) : null,
          isCheckedIn: row[11] ? true : false,
          isCheckedOut: row[12] && row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED' ? true : false,
          isCanceled: false,
          editCount: row[13] || 0
        });
      }
    }
    
    reservations.sort((a, b) => {
      const dateA = new Date(a.date + ' ' + a.startTime);
      const dateB = new Date(b.date + ' ' + b.startTime);
      return dateB - dateA;
    });
    
    return reservations;
  } catch (error) {
    console.error('Error getting bookings with checkin:', error);
    return [];
  }
}

function getAllReservations() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const reservations = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1]) continue;
      
      let canceledStatus = null;
      if (row[12] === 'AUTO_CANCELLED') {
        canceledStatus = 'AUTO';
      } else if (row[12] === 'USER_CANCELLED') {
        canceledStatus = 'USER';
      }
      
      reservations.push({
        rowIndex: i + 2,
        date: formatDate(new Date(row[0])),
        room: String(row[1]).trim(),
        startTime: formatTime(row[2]),
        endTime: formatTime(row[3]),
        name: row[4],
        department: row[5],
        studentId: row[6],
        phone: row[7],
        email: row[8],
        checkinCode: row[10] || '',
        isCheckedIn: row[11] ? true : false,
        isCheckedOut: row[12] && row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED' ? true : false,
        isCanceled: canceledStatus !== null,
        canceledType: canceledStatus
      });
    }
    
    return reservations;
  } catch (error) {
    console.error('Error getting all reservations:', error);
    return [];
  }
}

function getTodayCheckins() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const today = formatDate(new Date());
    const checkins = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1]) continue;
      
      const bookingDate = formatDate(new Date(row[0]));
      
      if (bookingDate === today) {
        let status = 'รอ Check-in';
        if (row[12] === 'AUTO_CANCELLED') {
          status = 'ยกเลิกอัตโนมัติ';
        } else if (row[12] === 'USER_CANCELLED') {
          status = 'ยกเลิกโดยผู้ใช้';
        } else if (row[12]) {
          status = 'เสร็จสิ้น';
        } else if (row[11]) {
          status = 'กำลังใช้งาน';
        }
        
        checkins.push({
          room: String(row[1]).trim(),
          startTime: formatTime(row[2]),
          endTime: formatTime(row[3]),
          name: row[4],
          department: row[5],
          email: row[8],
          checkinCode: row[10] || 'N/A',
          checkinTime: row[11] ? formatTime(row[11]) : 'ยังไม่ Check-in',
          checkoutTime: row[12] && row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED' ? formatTime(row[12]) : 'ยังไม่ Check-out',
          isCheckedIn: row[11] ? true : false,
          isCheckedOut: row[12] && row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED' ? true : false,
          status: status
        });
      }
    }
    
    return checkins;
  } catch (error) {
    console.error('Error getting today checkins:', error);
    return [];
  }
}

function getReservationsByDateAndRoom(date, room) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const reservations = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1] || row[2] === '' || row[3] === '') continue;
      
      if (row[12] === 'AUTO_CANCELLED' || row[12] === 'USER_CANCELLED') continue;
      
      const rowDate = formatDate(new Date(row[0]));
      const roomName = String(row[1]).trim();
      
      if (rowDate === date && roomName === room) {
        reservations.push({
          date: rowDate,
          room: roomName,
          startTime: formatTime(row[2]),
          endTime: formatTime(row[3]),
          name: row[4],
          isCheckedIn: row[11] ? true : false,
          isCheckedOut: row[12] && row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED' ? true : false
        });
      }
    }
    
    return reservations;
  } catch (error) {
    console.error('Error getting reservations by date and room:', error);
    return [];
  }
}

function deleteReservation(rowIndex, loginEmail, reservationEmail) {
  try {
    if (loginEmail.toLowerCase() !== ADMIN_EMAIL.toLowerCase()) {
      if (loginEmail.toLowerCase() !== reservationEmail.toLowerCase()) {
        return { success: false, message: 'คุณสามารถลบได้เฉพาะการจองของตัวเองเท่านั้น' };
      }
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Reservations' };
    
    const rowData = sheet.getRange(rowIndex, 1, 1, 14).getValues()[0];
    
    if (rowData[12] === 'USER_CANCELLED' || rowData[12] === 'AUTO_CANCELLED') {
      return { success: false, message: 'การจองนี้ถูกยกเลิกไปแล้ว' };
    }
    
    const bookingData = {
      name: rowData[4],
      room: rowData[1],
      date: formatDate(new Date(rowData[0])),
      startTime: formatTime(rowData[2]),
      endTime: formatTime(rowData[3])
    };
    
    sheet.getRange(rowIndex, 13).setValue('USER_CANCELLED');
    
    saveLoginLog(loginEmail, rowData[4], rowData[1], formatDate(new Date(rowData[0])), formatTime(rowData[2]), 'ยกเลิกการจอง');
    sendCancellationEmail(bookingData, reservationEmail);
    
    return { success: true, message: 'ยกเลิกการจองเรียบร้อยแล้ว และส่งอีเมลแจ้งเตือนแล้ว' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}// ===============================================================================================
// Code.gs - COMPLETE FILE Rev14 - Part 2
// Parts 6-10: Edit Booking, Schedule, Room Lock, Analytics, Export & Email Functions
// ===============================================================================================

// ===============================
// PART 6: EDIT BOOKING FUNCTIONS
// ===============================

function editBooking(rowIndex, newRoom, newStartTime, newEndTime, loginEmail, reservationEmail) {
  try {
    if (loginEmail.toLowerCase() !== ADMIN_EMAIL.toLowerCase()) {
      if (loginEmail.toLowerCase() !== reservationEmail.toLowerCase()) {
        return { success: false, message: 'คุณสามารถแก้ไขได้เฉพาะการจองของตัวเองเท่านั้น' };
      }
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Reservations' };
    
    const rowData = sheet.getRange(rowIndex, 1, 1, 14).getValues()[0];
    
    if (rowData[12] === 'USER_CANCELLED' || rowData[12] === 'AUTO_CANCELLED') {
      return { success: false, message: 'ไม่สามารถแก้ไขการจองที่ถูกยกเลิกแล้ว' };
    }
    
    if (rowData[11]) {
      return { success: false, message: 'ไม่สามารถแก้ไขการจองหลังจาก Check-in แล้ว' };
    }
    
    const currentEditCount = rowData[13] || 0;
    if (currentEditCount >= 2) {
      return { success: false, message: 'คุณแก้ไขการจองครบ 2 ครั้งแล้ว ไม่สามารถแก้ไขเพิ่มได้' };
    }
    
    const date = formatDate(new Date(rowData[0]));
    const oldRoom = rowData[1];
    const oldStartTime = formatTime(rowData[2]);
    const oldEndTime = formatTime(rowData[3]);
    
    const startMinutes = timeToMinutes(newStartTime);
    const endMinutes = timeToMinutes(newEndTime);
    const durationMinutes = endMinutes - startMinutes;
    
    if (durationMinutes <= 0) {
      return { success: false, message: 'เวลาสิ้นสุดต้องมากกว่าเวลาเริ่มต้น' };
    }
    
    const isTestUser = reservationEmail.toLowerCase().trim() === TEST_USER_EMAIL.toLowerCase();
    
    if (!isTestUser && durationMinutes > 180) {
      return { success: false, message: 'สามารถจองได้สูงสุด 3 ชั่วโมงเท่านั้น' };
    }
    
    const isDuplicate = checkDuplicateBookingExcludingCurrent(date, newRoom, newStartTime, newEndTime, rowIndex);
    if (isDuplicate) {
      return { 
        success: false, 
        message: 'ห้อง "' + newRoom + '" ในช่วงเวลา ' + newStartTime + ' - ' + newEndTime + ' ถูกจองแล้ว กรุณาเลือกเวลาอื่น' 
      };
    }
    
    sheet.getRange(rowIndex, 2).setValue(newRoom);
    sheet.getRange(rowIndex, 3).setValue(newStartTime);
    sheet.getRange(rowIndex, 4).setValue(newEndTime);
    
    const newEditCount = currentEditCount + 1;
    sheet.getRange(rowIndex, 14).setValue(newEditCount);
    
    let changes = [];
    if (oldRoom !== newRoom) changes.push('ห้อง: ' + oldRoom + ' → ' + newRoom);
    if (oldStartTime !== newStartTime || oldEndTime !== newEndTime) {
      changes.push('เวลา: ' + oldStartTime + '-' + oldEndTime + ' → ' + newStartTime + '-' + newEndTime);
    }
    const changeText = changes.join(', ');
    
    saveLoginLog(
      loginEmail, 
      rowData[4], 
      changeText, 
      date, 
      oldStartTime, 
      'แก้ไขการจอง (ครั้งที่ ' + newEditCount + ')'
    );
    
    sendBookingEditEmail(rowData, oldRoom, newRoom, oldStartTime, oldEndTime, newStartTime, newEndTime, reservationEmail, newEditCount);
    
    const remainingEdits = 2 - newEditCount;
    let message = 'แก้ไขการจองสำเร็จ';
    if (remainingEdits > 0) {
      message += ' (เหลือสิทธิ์แก้ไขอีก ' + remainingEdits + ' ครั้ง)';
    } else {
      message += ' (ใช้สิทธิ์แก้ไขครบแล้ว)';
    }
    
    return { 
      success: true, 
      message: message,
      remainingEdits: remainingEdits
    };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function checkDuplicateBookingExcludingCurrent(date, room, startTime, endTime, excludeRowIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return false;
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const newStartMinutes = timeToMinutes(startTime);
    const newEndMinutes = timeToMinutes(endTime);
    
    if (newStartMinutes >= newEndMinutes) {
      return true;
    }
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const currentRowIndex = i + 2;
      
      if (currentRowIndex === excludeRowIndex) continue;
      
      if (!row[0] || !row[1] || row[2] === '' || row[3] === '') continue;
      
      if (row[12] === 'AUTO_CANCELLED' || row[12] === 'USER_CANCELLED') continue;
      
      const rowDate = formatDate(new Date(row[0]));
      const roomName = String(row[1]).trim();
      
      if (rowDate === date && roomName === room) {
        const existingStartMinutes = timeToMinutes(formatTime(row[2]));
        const existingEndMinutes = timeToMinutes(formatTime(row[3]));
        
        if (newStartMinutes >= existingStartMinutes && newStartMinutes < existingEndMinutes) {
          return true;
        }
        
        if (newEndMinutes > existingStartMinutes && newEndMinutes <= existingEndMinutes) {
          return true;
        }
        
        if (newStartMinutes <= existingStartMinutes && newEndMinutes >= existingEndMinutes) {
          return true;
        }
      }
    }
    
    return false;
  } catch (error) {
    console.error('Error checking duplicate:', error);
    return true;
  }
}

function adminEditBooking(rowIndex, newRoom, newDate, newStartTime, newEndTime, adminEmail) {
  try {
    if (adminEmail.toLowerCase() !== ADMIN_EMAIL.toLowerCase()) {
      return { success: false, message: 'เฉพาะ Admin เท่านั้นที่สามารถแก้ไขได้' };
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Reservations' };
    
    const rowData = sheet.getRange(rowIndex, 1, 1, 14).getValues()[0];
    
    if (rowData[12] === 'USER_CANCELLED' || rowData[12] === 'AUTO_CANCELLED') {
      return { success: false, message: 'ไม่สามารถแก้ไขการจองที่ถูกยกเลิกแล้ว' };
    }
    
    const oldRoom = rowData[1];
    const oldDate = formatDate(new Date(rowData[0]));
    const oldStartTime = formatTime(rowData[2]);
    const oldEndTime = formatTime(rowData[3]);
    const userName = rowData[4];
    const userEmail = rowData[8];
    
    const startMinutes = timeToMinutes(newStartTime);
    const endMinutes = timeToMinutes(newEndTime);
    const durationMinutes = endMinutes - startMinutes;
    
    if (durationMinutes <= 0) {
      return { success: false, message: 'เวลาสิ้นสุดต้องมากกว่าเวลาเริ่มต้น' };
    }
    
    const isDuplicate = checkDuplicateBookingExcludingCurrent(newDate, newRoom, newStartTime, newEndTime, rowIndex);
    if (isDuplicate) {
      return { 
        success: false, 
        message: 'ห้อง "' + newRoom + '" ในวันที่ ' + newDate + ' เวลา ' + newStartTime + ' - ' + newEndTime + ' ถูกจองแล้ว' 
      };
    }
    
    sheet.getRange(rowIndex, 1).setValue(newDate);
    sheet.getRange(rowIndex, 2).setValue(newRoom);
    sheet.getRange(rowIndex, 3).setValue(newStartTime);
    sheet.getRange(rowIndex, 4).setValue(newEndTime);
    
    let changes = [];
    if (oldDate !== newDate) changes.push('วันที่: ' + oldDate + ' → ' + newDate);
    if (oldRoom !== newRoom) changes.push('ห้อง: ' + oldRoom + ' → ' + newRoom);
    if (oldStartTime !== newStartTime || oldEndTime !== newEndTime) {
      changes.push('เวลา: ' + oldStartTime + '-' + oldEndTime + ' → ' + newStartTime + '-' + newEndTime);
    }
    const changeText = changes.join(', ');
    
    saveLoginLog(adminEmail, userName, changeText, newDate, newStartTime, 'Admin Edit Booking');
    
    sendAdminEditNotificationEmail(rowData, oldDate, newDate, oldRoom, newRoom, oldStartTime, oldEndTime, newStartTime, newEndTime, userEmail);
    
    return { 
      success: true, 
      message: 'แก้ไขการจองสำเร็จ\n' + changeText
    };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

// ===============================
// PART 7: SCHEDULE FUNCTIONS
// ===============================

function getSchedule(date) {
  try {
    const allRooms = [
      'ห้อง 1 (ชั้น 1)', 'ห้อง 2 (ชั้น 1)', 'ห้อง 3 (ชั้น 1)',
      'ห้อง 4 (ชั้น 1)', 'ห้อง 5 (ชั้น 1)', 'ห้อง 6 (ชั้น 1)',
      'ห้อง 1 (ชั้นใต้ดิน)', 'ห้อง 2 (ชั้นใต้ดิน)', 'ห้อง 3 (ชั้นใต้ดิน)'
    ];
    
    const timeSlots = ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', 
                       '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30'];
    
    const reservations = getAllReservationsByDate(date);
    const locks = getRoomLocks();
    
    const schedule = [];
    const headerRow = ['เวลา', ...allRooms];
    schedule.push(headerRow);
    
    const now = new Date();
    const selectedDate = new Date(date);
    const isToday = formatDate(now) === formatDate(selectedDate);
    
    timeSlots.forEach(function(time) {
      const row = [time];
      
      const currentMinutes = isToday ? now.getHours() * 60 + now.getMinutes() : -1;
      const slotMinutes = timeToMinutes(time);
      const isPast = isToday && slotMinutes < currentMinutes;
      
      allRooms.forEach(function(room) {
        const lockStatus = isRoomLocked(room, date);
        
        if (lockStatus.isLocked) {
          row.push({
            status: 'locked',
            info: {
              reason: lockStatus.reason,
              startDate: lockStatus.startDate,
              endDate: lockStatus.endDate
            }
          });
        } else if (isPast) {
          row.push({ status: 'past' });
        } else {
          const booking = reservations.find(function(r) {
            return r.room === room && 
                   timeToMinutes(r.startTime) <= slotMinutes && 
                   timeToMinutes(r.endTime) > slotMinutes;
          });
          
          if (booking) {
            if (booking.isCheckedOut) {
              row.push({ status: 'completed' });
            } else if (booking.isCheckedIn) {
              row.push({
                status: 'using',
                info: {
                  name: booking.name,
                  time: booking.startTime + ' - ' + booking.endTime
                }
              });
            } else {
              row.push({
                status: 'booked',
                info: {
                  name: booking.name,
                  time: booking.startTime + ' - ' + booking.endTime
                }
              });
            }
          } else {
            row.push({ status: 'available' });
          }
        }
      });
      
      schedule.push(row);
    });
    
    return schedule;
  } catch (error) {
    console.error('Error getting schedule:', error);
    return [];
  }
}

function getAllReservationsByDate(date) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const reservations = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1] || row[2] === '' || row[3] === '') continue;
      
      if (row[12] === 'AUTO_CANCELLED' || row[12] === 'USER_CANCELLED') continue;
      
      if (formatDate(new Date(row[0])) === date) {
        reservations.push({
          room: String(row[1]).trim(),
          startTime: formatTime(row[2]),
          endTime: formatTime(row[3]),
          name: row[4],
          isCheckedIn: row[11] ? true : false,
          isCheckedOut: row[12] && row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED' ? true : false
        });
      }
    }
    
    return reservations;
  } catch (error) {
    console.error('Error getting reservations by date:', error);
    return [];
  }
}

function getBookedSlots(date, room) {
  try {
    const reservations = getReservationsByDateAndRoom(date, room);
    const bookedSlots = [];
    
    const now = new Date();
    const selectedDate = new Date(date);
    const isToday = formatDate(now) === formatDate(selectedDate);
    
    if (isToday) {
      const currentMinutes = now.getHours() * 60 + now.getMinutes();
      const allTimes = ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', 
                        '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30'];
      
      allTimes.forEach(function(time) {
        if (timeToMinutes(time) <= currentMinutes) {
          bookedSlots.push(time);
        }
      });
    }
    
    reservations.forEach(function(reservation) {
      const startMinutes = timeToMinutes(reservation.startTime);
      const endMinutes = timeToMinutes(reservation.endTime);
      
      const allTimes = ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', 
                        '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30'];
      
      allTimes.forEach(function(time) {
        const timeMinutes = timeToMinutes(time);
        if (timeMinutes >= startMinutes && timeMinutes < endMinutes) {
          if (!bookedSlots.includes(time)) {
            bookedSlots.push(time);
          }
        }
      });
    });
    
    return bookedSlots;
  } catch (error) {
    console.error('Error getting booked slots:', error);
    return [];
  }
}

function getBookedSlotsExcludingCurrent(date, room, excludeRowIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const bookedSlots = [];
    
    const now = new Date();
    const selectedDate = new Date(date);
    const isToday = formatDate(now) === formatDate(selectedDate);
    
    if (isToday) {
      const currentMinutes = now.getHours() * 60 + now.getMinutes();
      const allTimes = ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', 
                        '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30'];
      
      allTimes.forEach(function(time) {
        if (timeToMinutes(time) <= currentMinutes) {
          bookedSlots.push(time);
        }
      });
    }
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const currentRowIndex = i + 2;
      
      if (currentRowIndex === excludeRowIndex) continue;
      
      if (!row[0] || !row[1] || row[2] === '' || row[3] === '') continue;
      
      if (row[12] === 'AUTO_CANCELLED' || row[12] === 'USER_CANCELLED') continue;
      
      const rowDate = formatDate(new Date(row[0]));
      const roomName = String(row[1]).trim();
      
      if (rowDate === date && roomName === room) {
        const startMinutes = timeToMinutes(formatTime(row[2]));
        const endMinutes = timeToMinutes(formatTime(row[3]));
        
        const allTimes = ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', 
                          '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30'];
        
        allTimes.forEach(function(time) {
          const timeMinutes = timeToMinutes(time);
          if (timeMinutes >= startMinutes && timeMinutes < endMinutes) {
            if (!bookedSlots.includes(time)) {
              bookedSlots.push(time);
            }
          }
        });
      }
    }
    
    return bookedSlots;
  } catch (error) {
    console.error('Error getting booked slots excluding current:', error);
    return [];
  }
}

function getScheduleData(date) {
  try {
    const reservations = getAllReservationsByDate(date);
    const locks = getRoomLocks();
    return { success: true, reservations: reservations, locks: locks };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getAvailableTimeSlots(date, room) {
  try {
    const bookedSlots = getBookedSlots(date, room);
    const availableSlots = [];
    const allTimes = ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', 
                      '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30'];
    
    for (let i = 0; i < allTimes.length; i++) {
      if (!bookedSlots.includes(allTimes[i])) {
        availableSlots.push(allTimes[i]);
      }
    }
    
    return availableSlots;
  } catch (error) {
    console.error('Error getting available time slots:', error);
    return [];
  }
}

function getAvailableEndTimes(date, room, startTime) {
  try {
    const bookedSlots = getBookedSlots(date, room);
    const allTimes = ['08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30', 
                      '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30', '16:00'];
    const startIndex = allTimes.indexOf(startTime);
    const availableEndTimes = [];
    
    if (startIndex === -1) return [];
    
    const maxDuration = 6;
    let consecutiveAvailable = 0;
    
    for (let i = startIndex + 1; i < allTimes.length && consecutiveAvailable < maxDuration; i++) {
      if (bookedSlots.includes(allTimes[i])) {
        break;
      }
      consecutiveAvailable++;
      availableEndTimes.push(allTimes[i]);
    }
    
    if (!availableEndTimes.includes('16:00') && consecutiveAvailable >= maxDuration) {
      availableEndTimes.push('16:00');
    }
    
    return availableEndTimes;
  } catch (error) {
    console.error('Error getting available end times:', error);
    return [];
  }
}

// ===============================
// PART 8: ROOM LOCK FUNCTIONS
// ===============================

function setupRoomLockSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let lockSheet = ss.getSheetByName('RoomLocks');
    
    if (!lockSheet) {
      lockSheet = ss.insertSheet('RoomLocks');
      lockSheet.getRange(1, 1, 1, 8).setValues([
        ['Lock ID', 'ห้อง', 'วันที่เริ่ม', 'วันที่สิ้นสุด', 'เวลาเริ่ม', 'เวลาสิ้นสุด', 'เหตุผล', 'Admin Email']
      ]);
      lockSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }
    
    return { success: true, message: 'ตั้งค่าระบบล็อคห้องสำเร็จ' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function lockRoom(roomName, startDate, endDate, startTime, endTime, reason, adminEmail) {
  try {
    if (adminEmail.toLowerCase() !== ADMIN_EMAIL.toLowerCase()) {
      return { success: false, message: 'เฉพาะ Admin เท่านั้นที่สามารถล็อคห้องได้' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let lockSheet = ss.getSheetByName('RoomLocks');
    
    if (!lockSheet) {
      setupRoomLockSheet();
      lockSheet = ss.getSheetByName('RoomLocks');
    }
    
    const existingLocks = getRoomLocks();
    for (let lock of existingLocks) {
      if (lock.room === roomName) {
        if (startDate <= lock.endDate && endDate >= lock.startDate) {
          return { 
            success: false, 
            message: 'ห้องนี้ถูกล็อคไว้แล้วในช่วงวันที่ที่เลือก' 
          };
        }
      }
    }
    
    const lockId = 'LOCK_' + new Date().getTime();
    
    lockSheet.getRange(lockSheet.getLastRow() + 1, 1, 1, 8).setValues([[
      lockId,
      roomName,
      startDate,
      endDate,
      startTime || '08:00',
      endTime || '16:00',
      reason || 'ไม่ระบุ',
      adminEmail
    ]]);
    
    saveLoginLog(adminEmail, 'SYSTEM', roomName, startDate, startTime, 'Lock Room: ' + reason);
    
    return { 
      success: true, 
      message: 'ล็อคห้อง "' + roomName + '" สำเร็จ\nวันที่: ' + startDate + ' ถึง ' + endDate 
    };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function unlockRoom(lockId, adminEmail) {
  try {
    if (adminEmail.toLowerCase() !== ADMIN_EMAIL.toLowerCase()) {
      return { success: false, message: 'เฉพาะ Admin เท่านั้นที่สามารถปลดล็อคห้องได้' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lockSheet = ss.getSheetByName('RoomLocks');
    
    if (!lockSheet) {
      return { success: false, message: 'ไม่พบระบบล็อคห้อง' };
    }
    
    const data = lockSheet.getRange(2, 1, lockSheet.getLastRow() - 1, 8).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === lockId) {
        lockSheet.deleteRow(i + 2);
        
        saveLoginLog(adminEmail, 'SYSTEM', data[i][1], data[i][2], 'Unlock Room');
        
        return { success: true, message: 'ปลดล็อคห้องสำเร็จ' };
      }
    }
    
    return { success: false, message: 'ไม่พบข้อมูลการล็อคนี้' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function getRoomLocks() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lockSheet = ss.getSheetByName('RoomLocks');
    
    if (!lockSheet || lockSheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = lockSheet.getRange(2, 1, lockSheet.getLastRow() - 1, 8).getValues();
    const locks = [];
    
    for (let i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      locks.push({
        lockId: data[i][0],
        room: data[i][1],
        startDate: formatDate(new Date(data[i][2])),
        endDate: formatDate(new Date(data[i][3])),
        startTime: formatTime(data[i][4]),
        endTime: formatTime(data[i][5]),
        reason: data[i][6],
        adminEmail: data[i][7]
      });
    }
    
    return locks;
  } catch (error) {
    console.error('Error getting room locks:', error);
    return [];
  }
}

function isRoomLocked(roomName, date) {
  try {
    const locks = getRoomLocks();
    
    for (let lock of locks) {
      if (lock.room === roomName) {
        if (date >= lock.startDate && date <= lock.endDate) {
          return { 
            isLocked: true, 
            reason: lock.reason,
            startDate: lock.startDate,
            endDate: lock.endDate
          };
        }
      }
    }
    
    return { isLocked: false };
  } catch (error) {
    console.error('Error checking room lock:', error);
    return { isLocked: false };
  }
}

// ===============================
// PART 9: ANALYTICS FUNCTIONS
// ===============================

function getDashboardAnalytics() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        totalBookings: 0,
        activeBookings: 0,
        completedBookings: 0,
        cancelledBookings: 0,
        todayBookings: 0,
        weekBookings: 0,
        monthBookings: 0,
        roomUsage: {},
        peakHours: {},
        userStats: {
          totalUsers: 0,
          activeUsers: 0
        }
      };
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const now = new Date();
    const today = formatDate(now);
    
    const startOfWeek = new Date(now);
    startOfWeek.setDate(now.getDate() - now.getDay());
    const weekStart = formatDate(startOfWeek);
    
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const monthStart = formatDate(startOfMonth);
    
    let totalBookings = 0;
    let activeBookings = 0;
    let completedBookings = 0;
    let cancelledBookings = 0;
    let todayBookings = 0;
    let weekBookings = 0;
    let monthBookings = 0;
    
    const roomUsage = {};
    const peakHours = {};
    const uniqueUsers = new Set();
    const activeUsers = new Set();
    
    const allRooms = [
      'ห้อง 1 (ชั้น 1)', 'ห้อง 2 (ชั้น 1)', 'ห้อง 3 (ชั้น 1)',
      'ห้อง 4 (ชั้น 1)', 'ห้อง 5 (ชั้น 1)', 'ห้อง 6 (ชั้น 1)',
      'ห้อง 1 (ชั้นใต้ดิน)', 'ห้อง 2 (ชั้นใต้ดิน)', 'ห้อง 3 (ชั้นใต้ดิน)'
    ];
    
    allRooms.forEach(room => {
      roomUsage[room] = 0;
    });
    
    for (let hour = 8; hour <= 15; hour++) {
      const timeSlot = hour.toString().padStart(2, '0') + ':00';
      peakHours[timeSlot] = 0;
    }
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1]) continue;
      
      const bookingDate = formatDate(new Date(row[0]));
      const room = String(row[1]).trim();
      const userEmail = String(row[8]).toLowerCase().trim();
      const isCheckedIn = row[11] ? true : false;
      const checkoutStatus = row[12];
      
      uniqueUsers.add(userEmail);
      
      if (checkoutStatus !== 'AUTO_CANCELLED' && checkoutStatus !== 'USER_CANCELLED') {
        totalBookings++;
        
        if (roomUsage.hasOwnProperty(room)) {
          roomUsage[room]++;
        }
        
        const startHour = formatTime(row[2]).split(':')[0];
        const timeSlot = startHour + ':00';
        if (peakHours.hasOwnProperty(timeSlot)) {
          peakHours[timeSlot]++;
        }
        
        if (checkoutStatus && checkoutStatus !== 'AUTO_CANCELLED' && checkoutStatus !== 'USER_CANCELLED') {
          completedBookings++;
        } else if (isCheckedIn) {
          activeBookings++;
          activeUsers.add(userEmail);
        }
        
        if (bookingDate === today) {
          todayBookings++;
        }
        if (bookingDate >= weekStart) {
          weekBookings++;
        }
        if (bookingDate >= monthStart) {
          monthBookings++;
        }
      } else {
        cancelledBookings++;
      }
    }
    
    return {
      totalBookings: totalBookings,
      activeBookings: activeBookings,
      completedBookings: completedBookings,
      cancelledBookings: cancelledBookings,
      todayBookings: todayBookings,
      weekBookings: weekBookings,
      monthBookings: monthBookings,
      roomUsage: roomUsage,
      peakHours: peakHours,
      userStats: {
        totalUsers: uniqueUsers.size,
        activeUsers: activeUsers.size
      }
    };
  } catch (error) {
    console.error('Error getting dashboard analytics:', error);
    return {
      totalBookings: 0,
      activeBookings: 0,
      completedBookings: 0,
      cancelledBookings: 0,
      todayBookings: 0,
      weekBookings: 0,
      monthBookings: 0,
      roomUsage: {},
      peakHours: {},
      userStats: {
        totalUsers: 0,
        activeUsers: 0
      }
    };
  }
}

function getRecentActivity(limit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('LoginLog');
    
    if (!logSheet || logSheet.getLastRow() <= 1) {
      return [];
    }
    
    const lastRow = logSheet.getLastRow();
    const startRow = Math.max(2, lastRow - (limit || 10) + 1);
    const numRows = lastRow - startRow + 1;
    
    const data = logSheet.getRange(startRow, 1, numRows, 7).getValues();
    const activities = [];
    
    for (let i = data.length - 1; i >= 0; i--) {
      if (!data[i][0]) continue;
      
      activities.push({
        date: data[i][0],
        time: data[i][1],
        email: data[i][2],
        name: data[i][3],
        room: data[i][4],
        bookingDate: data[i][5],
        action: data[i][6]
      });
    }
    
    return activities;
  } catch (error) {
    console.error('Error getting recent activity:', error);
    return [];
  }
}

function getBookingStats() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { totalBookings: 0, todayBookings: 0, upcomingBookings: 0 };
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const now = new Date();
    const today = formatDate(now);
    
    let totalBookings = 0;
    let todayBookings = 0;
    let upcomingBookings = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      if (row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED') {
        totalBookings++;
        
        const bookingDate = formatDate(new Date(row[0]));
        if (bookingDate === today) {
          todayBookings++;
        }
        if (bookingDate >= today) {
          upcomingBookings++;
        }
      }
    }
    
    return {
      totalBookings: totalBookings,
      todayBookings: todayBookings,
      upcomingBookings: upcomingBookings
    };
  } catch (error) {
    console.error('Error getting booking stats:', error);
    return { totalBookings: 0, todayBookings: 0, upcomingBookings: 0 };
  }
}

function getQuickStats() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        totalToday: 0,
        waitingCheckin: 0,
        checkedInToday: 0,
        completedToday: 0,
        cancelledToday: 0
      };
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const today = formatDate(new Date());
    
    let totalToday = 0;
    let waitingCheckin = 0;
    let checkedInToday = 0;
    let completedToday = 0;
    let cancelledToday = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      const bookingDate = formatDate(new Date(row[0]));
      if (bookingDate === today) {
        totalToday++;
        
        if (row[12] === 'AUTO_CANCELLED' || row[12] === 'USER_CANCELLED') {
          cancelledToday++;
        } else if (row[12]) {
          completedToday++;
        } else if (row[11]) {
          checkedInToday++;
        } else {
          waitingCheckin++;
        }
      }
    }
    
    return {
      totalToday: totalToday,
      waitingCheckin: waitingCheckin,
      checkedInToday: checkedInToday,
      completedToday: completedToday,
      cancelledToday: cancelledToday
    };
  } catch (error) {
    console.error('Error getting quick stats:', error);
    return {
      totalToday: 0,
      waitingCheckin: 0,
      checkedInToday: 0,
      completedToday: 0,
      cancelledToday: 0
    };
  }
}

function searchBookings(searchTerm, searchType) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const results = [];
    const searchLower = searchTerm.toLowerCase().trim();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1]) continue;
      
      let match = false;
      
      if (searchType === 'all') {
        match = String(row[4]).toLowerCase().includes(searchLower) ||
                String(row[8]).toLowerCase().includes(searchLower) ||
                String(row[1]).toLowerCase().includes(searchLower) ||
                String(row[6]).toLowerCase().includes(searchLower) ||
                String(row[10]).toLowerCase().includes(searchLower);
      } else if (searchType === 'name') {
        match = String(row[4]).toLowerCase().includes(searchLower);
      } else if (searchType === 'email') {
        match = String(row[8]).toLowerCase().includes(searchLower);
      } else if (searchType === 'room') {
        match = String(row[1]).toLowerCase().includes(searchLower);
      } else if (searchType === 'studentId') {
        match = String(row[6]).toLowerCase().includes(searchLower);
      } else if (searchType === 'checkinCode') {
        match = String(row[10]).toLowerCase().includes(searchLower);
      }
      
      if (match) {
        let canceledStatus = null;
        if (row[12] === 'AUTO_CANCELLED') {
          canceledStatus = 'AUTO';
        } else if (row[12] === 'USER_CANCELLED') {
          canceledStatus = 'USER';
        }
        
        results.push({
          rowIndex: i + 2,
          date: formatDate(new Date(row[0])),
          room: String(row[1]).trim(),
          startTime: formatTime(row[2]),
          endTime: formatTime(row[3]),
          name: row[4],
          department: row[5],
          studentId: row[6],
          phone: row[7],
          email: row[8],
          checkinCode: row[10] || '',
          isCheckedIn: row[11] ? true : false,
          isCheckedOut: row[12] && row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED' ? true : false,
          isCanceled: canceledStatus !== null,
          canceledType: canceledStatus
        });
      }
    }
    
    return results;
  } catch (error) {
    console.error('Error searching bookings:', error);
    return [];
  }
}

// ===============================
// PART 10: EXPORT FUNCTIONS
// ===============================

function exportBookingsToExcel(exportType, startDate, endDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: false, message: 'ไม่มีข้อมูลให้ Export' };
    }
    
    let filterStartDate = null;
    let filterEndDate = null;
    let filterStatus = null;
    
    const today = formatDate(new Date());
    
    if (exportType === 'today') {
      filterStartDate = today;
      filterEndDate = today;
    } else if (exportType === 'week') {
      const now = new Date();
      const startOfWeek = new Date(now);
      startOfWeek.setDate(now.getDate() - now.getDay());
      filterStartDate = formatDate(startOfWeek);
      filterEndDate = today;
    } else if (exportType === 'month') {
      const now = new Date();
      const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
      filterStartDate = formatDate(startOfMonth);
      filterEndDate = today;
    } else if (exportType === 'daterange') {
      filterStartDate = startDate;
      filterEndDate = endDate;
    } else if (exportType === 'active') {
      filterStatus = 'active';
    } else if (exportType === 'completed') {
      filterStatus = 'completed';
    } else if (exportType === 'cancelled') {
      filterStatus = 'cancelled';
    }
    
    const result = exportToExcel(filterStatus, filterStartDate, filterEndDate);
    
    if (result.success && result.url) {
      const spreadsheetId = result.url.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
      return {
        success: true,
        message: result.message,
        url: result.url,
        spreadsheetId: spreadsheetId,
        sheetId: 0,
        sheetName: 'รายงานการจอง',
        count: result.message.match(/\d+/)[0]
      };
    }
    
    return result;
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function exportToExcel(filterType, startDate, endDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: false, message: 'ไม่มีข้อมูลให้ Export' };
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    const exportData = [];
    
    exportData.push([
      'วันที่จอง', 
      'ห้อง', 
      'เวลาเริ่ม', 
      'เวลาสิ้นสุด', 
      'ชื่อผู้จอง', 
      'สาขาวิชา', 
      'รหัสนักศึกษา', 
      'เบอร์โทร', 
      'อีเมล', 
      'วันที่บันทึก',
      'Check-in Time',
      'Check-out Time',
      'สถานะ'
    ]);
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1]) continue;
      
      const bookingDate = formatDate(new Date(row[0]));
      
      if (startDate && endDate) {
        if (bookingDate < startDate || bookingDate > endDate) {
          continue;
        }
      }
      
      let status = 'รอ Check-in';
      if (row[12] === 'AUTO_CANCELLED') {
        status = 'ยกเลิกอัตโนมัติ';
        if (filterType === 'active') continue;
      } else if (row[12] === 'USER_CANCELLED') {
        status = 'ยกเลิกโดยผู้ใช้';
        if (filterType === 'active') continue;
      } else if (row[12]) {
        status = 'เสร็จสิ้น';
        if (filterType === 'active' || filterType === 'pending') continue;
      } else if (row[11]) {
        status = 'กำลังใช้งาน';
        if (filterType === 'completed' || filterType === 'cancelled') continue;
      } else {
        if (filterType === 'completed' || filterType === 'cancelled') continue;
      }
      
      exportData.push([
        bookingDate,
        String(row[1]).trim(),
        formatTime(row[2]),
        formatTime(row[3]),
        row[4],
        row[5],
        row[6],
        row[7],
        row[8],
        row[9],
        row[11] ? formatTime(row[11]) : '-',
        row[12] && row[12] !== 'AUTO_CANCELLED' && row[12] !== 'USER_CANCELLED' ? formatTime(row[12]) : '-',
        status
      ]);
    }
    
    if (exportData.length <= 1) {
      return { success: false, message: 'ไม่มีข้อมูลตามเงื่อนไขที่เลือก' };
    }
    
    const newSS = SpreadsheetApp.create('รายงานการจอง_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss'));
    const newSheet = newSS.getActiveSheet();
    
    newSheet.getRange(1, 1, exportData.length, exportData[0].length).setValues(exportData);
    
    newSheet.getRange(1, 1, 1, exportData[0].length)
      .setFontWeight('bold')
      .setBackground('#667eea')
      .setFontColor('#ffffff');
    
    for (let i = 1; i <= exportData[0].length; i++) {
      newSheet.autoResizeColumn(i);
    }
    
    newSheet.setFrozenRows(1);
    
    const url = newSS.getUrl();
    
    return { 
      success: true, 
      message: 'Export สำเร็จ! จำนวน ' + (exportData.length - 1) + ' รายการ',
      url: url,
      spreadsheetId: newSS.getId()
    };
    
  } catch (error) {
    console.error('Error exporting to Excel:', error);
    return { 
      success: false, 
      message: 'เกิดข้อผิดพลาด: ' + error.toString() 
    };
  }
}

// ===============================================================================================
// EMAIL FUNCTIONS - REAL IMPLEMENTATION
// ===============================================================================================

function sendBookingConfirmationEmail(data, checkinCode) {
  try {
    const subject = '✅ ยืนยันการจองห้องประชุม - ' + data.room;
    
    const htmlBody = `
      <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background: #f8f9fa; border-radius: 10px;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px 10px 0 0; text-align: center;">
          <h1 style="margin: 0; font-size: 28px;">🎉 การจองสำเร็จ!</h1>
        </div>
        
        <div style="background: white; padding: 30px; border-radius: 0 0 10px 10px;">
          <h2 style="color: #667eea; margin-top: 0;">รายละเอียดการจอง</h2>
          
          <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
            <tr>
              <td style="padding: 12px; background: #f8f9fa; font-weight: bold; width: 40%;">🏢 ห้อง:</td>
              <td style="padding: 12px;">${data.room}</td>
            </tr>
            <tr>
              <td style="padding: 12px; background: #f8f9fa; font-weight: bold;">📅 วันที่:</td>
              <td style="padding: 12px;">${data.date}</td>
            </tr>
            <tr>
              <td style="padding: 12px; background: #f8f9fa; font-weight: bold;">⏰ เวลา:</td>
              <td style="padding: 12px;">${data.startTime} - ${data.endTime}</td>
            </tr>
            <tr>
              <td style="padding: 12px; background: #f8f9fa; font-weight: bold;">👤 ชื่อผู้จอง:</td>
              <td style="padding: 12px;">${data.name}</td>
            </tr>
            <tr>
              <td style="padding: 12px; background: #f8f9fa; font-weight: bold;">🎓 สาขา:</td>
              <td style="padding: 12px;">${data.department}</td>
            </tr>
          </table>
          
          <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 25px; border-radius: 10px; text-align: center; margin: 30px 0;">
            <p style="margin: 0 0 10px 0; font-size: 16px;">รหัส Check-in ของคุณ</p>
            <div style="font-size: 42px; font-weight: bold; letter-spacing: 8px; font-family: 'Courier New', monospace;">${checkinCode}</div>
            <p style="margin: 10px 0 0 0; font-size: 14px; opacity: 0.9;">กรุณาเก็บรหัสนี้ไว้</p>
          </div>
          
          <div style="background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; border-radius: 5px;">
            <h3 style="color: #856404; margin-top: 0; font-size: 18px;">⚠️ สิ่งสำคัญที่ต้องจำ!</h3>
            <ul style="color: #856404; line-height: 1.8; margin: 10px 0;">
              <li><strong>Check-in ภายใน 15 นาที</strong> หลังเวลาเริ่มจอง</li>
              <li>ไม่ Check-in ภายใน 15 นาที = <strong style="color: #dc2743;">ยกเลิกอัตโนมัติ</strong></li>
              <li>Check-out เมื่อใช้เสร็จ</li>
              <li>แก้ไขได้สูงสุด 2 ครั้ง (ก่อน Check-in)</li>
            </ul>
          </div>
          
          <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 2px solid #e9ecef;">
            <p style="color: #666; font-size: 14px; margin: 5px 0;">ระบบจองห้องประชุม YUNGS!XX</p>
            <p style="color: #999; font-size: 12px; margin: 5px 0;">อีเมลนี้ส่งอัตโนมัติ กรุณาอย่าตอบกลับ</p>
          </div>
        </div>
      </div>
    `;
    
    GmailApp.sendEmail(
      data.loginEmail,
      subject,
      'กรุณาเปิดอีเมลในรูปแบบ HTML เพื่อดูรายละเอียดการจอง',
      {
        htmlBody: htmlBody,
        name: 'ระบบจองห้องประชุม YUNGS!XX'
      }
    );
    
    return true;
  } catch (error) {
    console.error('Error sending booking confirmation email:', error);
    return false;
  }
}

function sendEndingSoonEmail(email, name, room, date, startTime, endTime) {
  try {
    const subject = '⏰ แจ้งเตือน: อีก 5 นาทีจะหมดเวลาใช้ห้อง - ' + room;
    
    const htmlBody = `
      <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background: #f8f9fa;">
        <div style="background: linear-gradient(135deg, #f2994a 0%, #f2c94c 100%); color: white; padding: 25px; border-radius: 10px; text-align: center;">
          <h1 style="margin: 0; font-size: 32px;">⏰</h1>
          <h2 style="margin: 10px 0 0 0;">อีก 5 นาทีจะหมดเวลา!</h2>
        </div>
        
        <div style="background: white; padding: 25px; border-radius: 10px; margin-top: 20px;">
          <p style="font-size: 16px;">สวัสดี ${name},</p>
          <p style="font-size: 16px; line-height: 1.8;">
            การจองห้อง <strong>${room}</strong> ของคุณจะสิ้นสุดในอีก <strong style="color: #f57c00;">5 นาที</strong>
          </p>
          
          <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">📅 วันที่:</td>
              <td style="padding: 10px;">${date}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">⏰ เวลา:</td>
              <td style="padding: 10px;">${startTime} - ${endTime}</td>
            </tr>
          </table>
          
          <div style="background: #fff3cd; padding: 15px; border-radius: 5px; border-left: 4px solid #ffc107;">
            <p style="margin: 0; color: #856404;">
              💡 กรุณา <strong>Check-out</strong> เมื่อใช้เสร็จ<br>
              หากไม่ Check-out ระบบจะทำอัตโนมัติเมื่อหมดเวลา
            </p>
          </div>
        </div>
      </div>
    `;
    
    GmailApp.sendEmail(email, subject, 'กรุณาเปิดอีเมลในรูปแบบ HTML', { htmlBody: htmlBody, name: 'ระบบจองห้องประชุม YUNGS!XX' });
  } catch (error) {
    console.error('Error sending ending soon email:', error);
  }
}

function sendAutoCheckoutEmail(email, name, room, date, startTime, endTime) {
  try {
    const subject = '✅ Auto Check-out สำเร็จ - ' + room;
    
    const htmlBody = `
      <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background: #f8f9fa;">
        <div style="background: linear-gradient(135deg, #56ab2f 0%, #a8e063 100%); color: white; padding: 25px; border-radius: 10px; text-align: center;">
          <h1 style="margin: 0; font-size: 32px;">✅</h1>
          <h2 style="margin: 10px 0 0 0;">Auto Check-out สำเร็จ</h2>
        </div>
        
        <div style="background: white; padding: 25px; border-radius: 10px; margin-top: 20px;">
          <p style="font-size: 16px;">สวัสดี ${name},</p>
          <p style="font-size: 16px;">ระบบได้ Check-out ห้อง <strong>${room}</strong> ให้คุณอัตโนมัติแล้ว เนื่องจากหมดเวลาการจอง</p>
          
          <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">📅 วันที่:</td>
              <td style="padding: 10px;">${date}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">⏰ เวลา:</td>
              <td style="padding: 10px;">${startTime} - ${endTime}</td>
            </tr>
          </table>
          
          <p style="text-align: center; color: #2e7d32; font-size: 18px; font-weight: bold;">ขอบคุณที่ใช้บริการ! 🙏</p>
        </div>
      </div>
    `;
    
    GmailApp.sendEmail(email, subject, 'กรุณาเปิดอีเมลในรูปแบบ HTML', { htmlBody: htmlBody, name: 'ระบบจองห้องประชุม YUNGS!XX' });
  } catch (error) {
    console.error('Error sending auto checkout email:', error);
  }
}

function sendAutoCancelEmail(email, name, room, date, startTime, endTime) {
  try {
    const subject = '❌ การจองถูกยกเลิกอัตโนมัติ - ' + room;
    
    const htmlBody = `
      <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background: #f8f9fa;">
        <div style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color: white; padding: 25px; border-radius: 10px; text-align: center;">
          <h1 style="margin: 0; font-size: 32px;">❌</h1>
          <h2 style="margin: 10px 0 0 0;">การจองถูกยกเลิกอัตโนมัติ</h2>
        </div>
        
        <div style="background: white; padding: 25px; border-radius: 10px; margin-top: 20px;">
          <p style="font-size: 16px;">สวัสดี ${name},</p>
          <p style="font-size: 16px;">การจองห้อง <strong>${room}</strong> ของคุณถูกยกเลิกอัตโนมัติ</p>
          
          <div style="background: #ffebee; border-left: 4px solid #f5576c; padding: 15px; margin: 20px 0; border-radius: 5px;">
            <p style="margin: 0; color: #c62828; font-weight: bold;">
              ⚠️ เหตุผล: ไม่ได้ Check-in ภายใน 15 นาทีหลังเวลาเริ่มจอง
            </p>
          </div>
          
          <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">🏢 ห้อง:</td>
              <td style="padding: 10px;">${room}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">📅 วันที่:</td>
              <td style="padding: 10px;">${date}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">⏰ เวลา:</td>
              <td style="padding: 10px;">${startTime} - ${endTime}</td>
            </tr>
          </table>
          
          <p style="color: #666;">หากต้องการจองอีกครั้ง กรุณาเข้าสู่ระบบและทำการจองใหม่</p>
          <p style="color: #666;"><strong>สำคัญ:</strong> กรุณา Check-in ภายในเวลาที่กำหนดในการจองครั้งถัดไป</p>
        </div>
      </div>
    `;
    
    GmailApp.sendEmail(email, subject, 'กรุณาเปิดอีเมลในรูปแบบ HTML', { htmlBody: htmlBody, name: 'ระบบจองห้องประชุม YUNGS!XX' });
  } catch (error) {
    console.error('Error sending auto cancel email:', error);
  }
}

function sendCancellationEmail(bookingData, email) {
  try {
    const subject = '🗑️ ยืนยันการยกเลิกการจอง - ' + bookingData.room;
    
    const htmlBody = `
      <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background: #f8f9fa;">
        <div style="background: linear-gradient(135deg, #9e9e9e 0%, #757575 100%); color: white; padding: 25px; border-radius: 10px; text-align: center;">
          <h1 style="margin: 0; font-size: 32px;">🗑️</h1>
          <h2 style="margin: 10px 0 0 0;">การจองถูกยกเลิกแล้ว</h2>
        </div>
        
        <div style="background: white; padding: 25px; border-radius: 10px; margin-top: 20px;">
          <p style="font-size: 16px;">สวัสดี ${bookingData.name},</p>
          <p style="font-size: 16px;">การจองห้องของคุณถูกยกเลิกเรียบร้อยแล้ว</p>
          
          <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">🏢 ห้อง:</td>
              <td style="padding: 10px;">${bookingData.room}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">📅 วันที่:</td>
              <td style="padding: 10px;">${bookingData.date}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">⏰ เวลา:</td>
              <td style="padding: 10px;">${bookingData.startTime} - ${bookingData.endTime}</td>
            </tr>
          </table>
          
          <p style="text-align: center; color: #666;">หากต้องการจองอีกครั้ง กรุณาเข้าสู่ระบบและทำการจองใหม่</p>
        </div>
      </div>
    `;
    
    GmailApp.sendEmail(email, subject, 'กรุณาเปิดอีเมลในรูปแบบ HTML', { htmlBody: htmlBody, name: 'ระบบจองห้องประชุม YUNGS!XX' });
  } catch (error) {
    console.error('Error sending cancellation email:', error);
  }
}

function sendBookingEditEmail(rowData, oldRoom, newRoom, oldStartTime, oldEndTime, newStartTime, newEndTime, email, editCount) {
  try {
    const subject = '✏️ การจองของคุณถูกแก้ไข - ' + newRoom;
    
    let changes = [];
    if (oldRoom !== newRoom) changes.push('ห้อง: ' + oldRoom + ' → ' + newRoom);
    if (oldStartTime !== newStartTime || oldEndTime !== newEndTime) {
      changes.push('เวลา: ' + oldStartTime + '-' + oldEndTime + ' → ' + newStartTime + '-' + newEndTime);
    }
    
    const htmlBody = `
      <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background: #f8f9fa;">
        <div style="background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%); color: white; padding: 25px; border-radius: 10px; text-align: center;">
          <h1 style="margin: 0; font-size: 32px;">✏️</h1>
          <h2 style="margin: 10px 0 0 0;">การจองถูกแก้ไขแล้ว</h2>
        </div>
        
        <div style="background: white; padding: 25px; border-radius: 10px; margin-top: 20px;">
          <p style="font-size: 16px;">สวัสดี ${rowData[4]},</p>
          <p style="font-size: 16px;">การจองของคุณถูกแก้ไขเรียบร้อยแล้ว (ครั้งที่ ${editCount}/2)</p>
          
          <div style="background: #e3f2fd; padding: 15px; margin: 20px 0; border-radius: 5px; border-left: 4px solid #2196f3;">
            <h3 style="margin-top: 0; color: #1565c0;">การเปลี่ยนแปลง:</h3>
            <ul style="color: #1976d2; line-height: 1.8;">
              ${changes.map(c => '<li>' + c + '</li>').join('')}
            </ul>
          </div>
          
          <h3 style="color: #667eea;">รายละเอียดการจองใหม่:</h3>
          <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">🏢 ห้อง:</td>
              <td style="padding: 10px;">${newRoom}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">📅 วันที่:</td>
              <td style="padding: 10px;">${formatDate(new Date(rowData[0]))}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">⏰ เวลา:</td>
              <td style="padding: 10px;">${newStartTime} - ${newEndTime}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">🔑 รหัส Check-in:</td>
              <td style="padding: 10px; font-weight: bold; color: #667eea;">${rowData[10]}</td>
            </tr>
          </table>
          
          ${editCount < 2 ? 
            '<p style="color: #666;">💡 คุณสามารถแก้ไขได้อีก ' + (2 - editCount) + ' ครั้ง</p>' :
            '<div style="background: #ffebee; padding: 15px; border-radius: 5px; border-left: 4px solid #f5576c;"><p style="margin: 0; color: #c62828;">⚠️ คุณได้ใช้สิทธิ์แก้ไขครบ 2 ครั้งแล้ว ไม่สามารถแก้ไขเพิ่มได้</p></div>'
          }
        </div>
      </div>
    `;
    
    GmailApp.sendEmail(email, subject, 'กรุณาเปิดอีเมลในรูปแบบ HTML', { htmlBody: htmlBody, name: 'ระบบจองห้องประชุม YUNGS!XX' });
  } catch (error) {
    console.error('Error sending booking edit email:', error);
  }
}

function sendAdminEditNotificationEmail(rowData, oldDate, newDate, oldRoom, newRoom, oldStartTime, oldEndTime, newStartTime, newEndTime, email) {
  try {
    const subject = '🔧 Admin แก้ไขการจองของคุณ - ' + newRoom;
    
    let changes = [];
    if (oldDate !== newDate) changes.push('วันที่: ' + oldDate + ' → ' + newDate);
    if (oldRoom !== newRoom) changes.push('ห้อง: ' + oldRoom + ' → ' + newRoom);
    if (oldStartTime !== newStartTime || oldEndTime !== newEndTime) {
      changes.push('เวลา: ' + oldStartTime + '-' + oldEndTime + ' → ' + newStartTime + '-' + newEndTime);
    }
    
    const htmlBody = `
      <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background: #f8f9fa;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 25px; border-radius: 10px; text-align: center;">
          <h1 style="margin: 0; font-size: 32px;">🔧</h1>
          <h2 style="margin: 10px 0 0 0;">Admin แก้ไขการจองของคุณ</h2>
        </div>
        
        <div style="background: white; padding: 25px; border-radius: 10px; margin-top: 20px;">
          <p style="font-size: 16px;">สวัสดี ${rowData[4]},</p>
          <p style="font-size: 16px;">Admin ได้ทำการแก้ไขการจองของคุณ</p>
          
          <div style="background: #e3f2fd; padding: 15px; margin: 20px 0; border-radius: 5px; border-left: 4px solid #2196f3;">
            <h3 style="margin-top: 0; color: #1565c0;">การเปลี่ยนแปลง:</h3>
            <ul style="color: #1976d2; line-height: 1.8;">
              ${changes.map(c => '<li>' + c + '</li>').join('')}
            </ul>
          </div>
          
          <h3 style="color: #667eea;">รายละเอียดการจองใหม่:</h3>
          <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">🏢 ห้อง:</td>
              <td style="padding: 10px;">${newRoom}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">📅 วันที่:</td>
              <td style="padding: 10px;">${newDate}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">⏰ เวลา:</td>
              <td style="padding: 10px;">${newStartTime} - ${newEndTime}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background: #f8f9fa; font-weight: bold;">🔑 รหัส Check-in:</td>
              <td style="padding: 10px; font-weight: bold; color: #667eea;">${rowData[10]}</td>
            </tr>
          </table>
          
          <div style="background: #fff3cd; padding: 15px; border-radius: 5px; border-left: 4px solid #ffc107;">
            <p style="margin: 0; color: #856404;">
              💡 <strong>หมายเหตุ:</strong> รหัส Check-in เดิมยังใช้ได้ตามปกติ
            </p>
          </div>
        </div>
      </div>
    `;
    
    GmailApp.sendEmail(email, subject, 'กรุณาเปิดอีเมลในรูปแบบ HTML', { htmlBody: htmlBody, name: 'ระบบจองห้องประชุม YUNGS!XX' });
  } catch (error) {
    console.error('Error sending admin edit notification email:', error);
  }
}