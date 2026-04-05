// =========================================================================
// System: reAgentics - Laboratory Reagent Management System
// Version: 0.8.8 (AI Core)
// Developer: P. PURICUMPEE & AI Assistant
// Description: Backend Google Apps Script (Server-side logic)
// Update: Added Company, Price, Transport Temp/Speed, Delivery Note PDF
// =========================================================================

// -------------------------------------------------------------------------
// 1. DATABASE CONFIGURATION (SaaS Architecture)
// -------------------------------------------------------------------------
const DEFAULT_DB = {
  MAIN: 'ใส่_ID_ไฟล์_reAgentics_DB_ที่นี่',       // Items, Stock_Balance
  UNIT: 'ใส่_ID_ไฟล์_reAgentics_Units_ที่นี่',     // Units, ReagUnits, Analyzers, storageLocation, Company
  USER: 'ใส่_ID_ไฟล์_reAgentics_User_ที่นี่',      // User
  CONFIG: 'ใส่_ID_ไฟล์_reAgentics_Config_ที่นี่',  // Sticker_Config, App_Logo, Year_Config
  LOG: 'ใส่_ID_ไฟล์_reAgentics_Log_ที่นี่',        // System_Logs
  FOLDER_PROFILE: '', // Folder ID สำหรับโปรไฟล์
  FOLDER_LOGO: ''     // Folder ID สำหรับโลโก้
};

function getDbConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    MAIN: props.getProperty('DB_MAIN') || DEFAULT_DB.MAIN,
    UNIT: props.getProperty('DB_UNIT') || DEFAULT_DB.UNIT,
    USER: props.getProperty('DB_USER') || DEFAULT_DB.USER,
    CONFIG: props.getProperty('DB_CONFIG') || DEFAULT_DB.CONFIG,
    LOG: props.getProperty('DB_LOG') || DEFAULT_DB.LOG,
    FOLDER_PROFILE: props.getProperty('FOLDER_PROFILE') || DEFAULT_DB.FOLDER_PROFILE,
    FOLDER_LOGO: props.getProperty('FOLDER_LOGO') || DEFAULT_DB.FOLDER_LOGO
  };
}

function checkDatabaseSetup() {
  const config = getDbConfig();
  if (!config.MAIN || config.MAIN.includes('ใส่_ID_ไฟล์') || !config.USER || config.USER.includes('ใส่_ID_ไฟล์')) {
    throw new Error("คุณยังไม่ได้ตั้งค่า Google Sheet ID ครับ กรุณานำ ID มาใส่ในหน้า 'ตั้งค่าฐานข้อมูล' ให้ครบถ้วน");
  }
}

// Helper: ลบไฟล์เดิมใน Google Drive หากมี URL เก่า
function deleteOldDriveFile(oldUrl) {
  if (oldUrl && oldUrl.includes("drive.google.com")) {
    try {
      let fileId = "";
      const idMatch = oldUrl.match(/id=([^&]+)/);
      const dMatch = oldUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      
      if (idMatch && idMatch[1]) {
        fileId = idMatch[1];
      } else if (dMatch && dMatch[1]) {
        fileId = dMatch[1];
      }
      
      if (fileId) {
        DriveApp.getFileById(fileId).setTrashed(true);
      }
    } catch (err) {
      console.log("Delete old file error: " + err);
    }
  }
}

// -------------------------------------------------------------------------
// 2. CORE WEB APP FUNCTIONS
// -------------------------------------------------------------------------
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('reAgentics | Lab Inventory System (v0.8.8 AI Core)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function safeString(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (val === null || val === undefined) return "";
  return String(val).trim();
}

// -------------------------------------------------------------------------
// 3. SYSTEM LOGGING (reAgentics_Log) - PDPA Compliance
// -------------------------------------------------------------------------
function logSystem(action, detail, userId) {
  try {
    checkDatabaseSetup(); 
    const config = getDbConfig();
    if (config.LOG && !config.LOG.includes('ใส่_ID_ไฟล์')) {
      const logSS = SpreadsheetApp.openById(config.LOG);
      let sheet = logSS.getSheetByName('System_Logs');
      
      if (!sheet) {
        sheet = logSS.insertSheet('System_Logs');
        sheet.appendRow(['Timestamp', 'UserID', 'Action', 'Detail']);
        sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#e2e8f0");
        sheet.setFrozenRows(1);
      }
      sheet.appendRow([new Date(), userId, action, detail]);
    }
  } catch(e) { console.error("Log Sys Error: " + e); }
}

// -------------------------------------------------------------------------
// 4. AUTHENTICATION & SECURITY
// -------------------------------------------------------------------------
function verifyLogin(userId, password) {
  try {
    checkDatabaseSetup(); 
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName("User");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId) && String(data[i][2]) === String(password)) {
        
        // เช็ค Status การระงับใช้งาน (Col I / index 8)
        let status = String(data[i][8] || "ปกติ").trim();
        if (status === "ระงับการใช้งาน") {
          logSystem("Login Blocked", "Suspended user tried to login", userId);
          return { success: false, message: `บัญชีผู้ใช้ ${userId} ถูกระงับ โปรดติดต่อผู้ดูแลระบบ` };
        }

        let email = data[i][3];
        let name = data[i][0];
        // เช็คสถานะ OTP Bypass (Col J / index 9) ถ้าว่างให้ถือว่า ON
        let otpStatus = String(data[i][9] || "ON").trim().toUpperCase();
        
        if (!email && otpStatus === "ON") {
          logSystem("Login Failed", "Account missing email", userId);
          return { success: false, message: "บัญชีนี้ยังไม่ได้ตั้งค่า Email กรุณาติดต่อ Admin" };
        }

        // กรณีตั้งค่า OTP Bypass = OFF (ข้ามการส่ง OTP)
        if (otpStatus === "OFF") {
          let userProfile = getUserProfileById(userId);
          if (userProfile) {
            logSystem("Login Success", "User authenticated (OTP Bypassed)", userId);
            let availableYears = getAvailableYears();
            return { 
              success: true, 
              bypassed: true, // แจ้ง Frontend ให้ข้ามหน้า OTP
              message: "เข้าสู่ระบบสำเร็จ (OTP Bypassed)", 
              user: userProfile, 
              years: availableYears 
            };
          } else { 
            return { success: false, message: "ไม่พบข้อมูลโปรไฟล์ผู้ใช้งาน" }; 
          }
        }

        // กรณี OTP ปกติ (ON)
        let cache = CacheService.getScriptCache();
        let existingOtp = cache.get("OTP_" + userId);
        
        // หากมี OTP ตัวเดิมในระบบ (ยังไม่ครบ 5 นาที) จะไม่สร้างใหม่และไม่ส่งเมลซ้ำ
        if (existingOtp) {
          logSystem("Login Info", "User logged in with active OTP session", userId);
          return { 
            success: true, 
            message: "ระบบได้ส่ง OTP ไปก่อนหน้านี้แล้ว กรุณาใช้รหัสเดิม (อายุรหัส 5 นาที)", 
            email: email, 
            userId: userId, 
            otpExists: true 
          };
        } else {
          // หากไม่มี OTP ในระบบ ให้สร้างใหม่และส่งเมล
          let otpResult = generateAndSendOTP(userId, email, name);
          if (otpResult.success) {
            logSystem("OTP Requested", "New OTP sent to email", userId);
            return { success: true, message: "กรุณาตรวจสอบ OTP ที่อีเมลของคุณ", email: email, userId: userId };
          } else {
            logSystem("OTP Error", otpResult.error, userId);
            return { success: false, message: "ไม่สามารถส่งอีเมล OTP ได้: " + otpResult.error };
          }
        }
      }
    }
    logSystem("Login Failed", "Invalid credentials", userId);
    return { success: false, message: "UserID หรือ Password ไม่ถูกต้อง!" };
  } catch (error) {
    return { success: false, message: "ระบบฐานข้อมูลขัดข้อง: " + error.message };
  }
}

function apiRequestNewOTP(userId) {
  try {
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName("User");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId)) {
        let email = data[i][3];
        let name = data[i][0];
        
        if (!email) return { success: false, message: "บัญชีนี้ยังไม่ได้ตั้งค่า Email" };
        
        // ล้าง Cache เดิมแล้วสร้างใหม่
        CacheService.getScriptCache().remove("OTP_" + userId);
        let otpResult = generateAndSendOTP(userId, email, name);
        
        if (otpResult.success) {
          logSystem("OTP Resent", "User explicitly requested a new OTP", userId);
          return { success: true, message: "ส่ง OTP รหัสใหม่เรียบร้อยแล้ว" };
        } else {
          return { success: false, message: "เกิดข้อผิดพลาด: " + otpResult.error };
        }
      }
    }
    return { success: false, message: "ไม่พบข้อมูลผู้ใช้งาน" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function generateAndSendOTP(userId, email, name) {
  try {
    let otp = Math.floor(100000 + Math.random() * 900000).toString(); 
    CacheService.getScriptCache().put("OTP_" + userId, otp, 300); // 300 seconds = 5 minutes
    
    let logoUrl = "https://cdn-icons-png.flaticon.com/512/3003/3003251.png"; 
    try {
      const config = getDbConfig();
      if(config.CONFIG && !config.CONFIG.includes('ใส่_ID_ไฟล์')) {
        const configSS = SpreadsheetApp.openById(config.CONFIG);
        const logoSheet = configSS.getSheetByName('App_Logo');
        if (logoSheet && logoSheet.getLastRow() > 1) {
            logoUrl = logoSheet.getRange(2, 2).getValue();
        }
      }
    } catch(e) {}

    const htmlTemplate = `
        <div style="font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif; color: #1e293b; max-width: 600px; margin: 0 auto; padding: 20px;">
            <div style="text-align: center; margin-bottom: 24px;">
                <img src="${logoUrl}" alt="reAgentics Logo" style="width: 56px; height: 56px; border-radius: 12px; object-fit: contain; vertical-align: middle;">
                <span style="font-size: 28px; font-weight: 700; color: #0ea5e9; vertical-align: middle; margin-left: 12px; letter-spacing: -0.5px; display: inline-block;">reAgentics</span>
            </div>
            <h2 style="font-size: 22px; font-weight: 500; text-align: center; margin-bottom: 24px; color: #334155;">
                กรุณายืนยันตัวตนของคุณ, <strong>${name}</strong>
            </h2>
            <div style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 24px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
                <p style="margin-top: 0; margin-bottom: 16px; font-size: 15px; text-align: center;">นี่คือรหัส OTP สำหรับเข้าสู่ระบบบริหารจัดการน้ำยา:</p>
                <div style="text-align: center; font-size: 36px; font-family: ui-monospace, SFMono-Regular, Consolas, monospace; font-weight: 700; letter-spacing: 10px; color: #0f172a; margin: 28px 0; background-color: #f8fafc; padding: 16px; border-radius: 8px;">
                    ${otp}
                </div>
                <p style="font-size: 14px; margin-bottom: 16px; text-align: center; color: #475569;">
                    รหัสนี้มีอายุการใช้งาน <strong>5 นาที</strong> และใช้ได้เพียงครั้งเดียว
                </p>
                <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 24px 0;">
                <p style="font-size: 13px; margin-bottom: 12px; color: #64748b;">
                    <strong style="color: #ef4444;">ข้อควรระวัง (PDPA):</strong> โปรดอย่าแชร์รหัสนี้กับบุคคลอื่น ทีมงาน reAgentics จะไม่ขอรหัสผ่านหรือ OTP ของคุณผ่านช่องทางใดๆ โดยเด็ดขาด
                </p>
            </div>
        </div>
    `;

    MailApp.sendEmail({ to: email, subject: "รหัส OTP สำหรับเข้าสู่ระบบ reAgentics", htmlBody: htmlTemplate, name: "reAgentics LIS" });
    return { success: true };
  } catch (error) { return { success: false, error: error.message }; }
}

function verifyOTP(userId, inputOtp) {
  try {
    checkDatabaseSetup();
    let cache = CacheService.getScriptCache();
    let cachedOtp = cache.get("OTP_" + userId);
    
    if (!cachedOtp) {
      logSystem("Login Failed", "Expired or missing OTP", userId);
      return { success: false, message: "OTP หมดอายุหรือไม่ถูกต้อง กรุณาเข้าสู่ระบบใหม่" };
    }
    
    if (cachedOtp === inputOtp.toString()) {
      cache.remove("OTP_" + userId);
      let userProfile = getUserProfileById(userId);
      if(userProfile) {
        logSystem("Login Success", "User successfully authenticated", userId);
        let availableYears = getAvailableYears();
        return { success: true, message: "เข้าสู่ระบบสำเร็จ", user: userProfile, years: availableYears };
      } else { return { success: false, message: "ไม่พบข้อมูลโปรไฟล์ผู้ใช้งาน" }; }
    } else {
      logSystem("Login Failed", "Invalid OTP entered", userId);
      return { success: false, message: "รหัส OTP ไม่ถูกต้อง" };
    }
  } catch (e) { return { success: false, message: "Verify Error: " + e.message }; }
}

function verifyPasswordOnly(userId, password) {
  try {
    checkDatabaseSetup();
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName('User');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId) && String(data[i][2]) === String(password)) {
        
        let status = String(data[i][8] || "ปกติ").trim();
        if (status === "ระงับการใช้งาน") {
          logSystem("Unlock Blocked", "Suspended user tried to unlock screen", userId);
          return { success: false, message: `บัญชีผู้ใช้ ${userId} ถูกระงับ โปรดติดต่อผู้ดูแลระบบ` };
        }

        logSystem("Unlock Screen", "Successfully unlocked screen", userId);
        return { success: true };
      }
    }
    logSystem("Unlock Failed", "Invalid password during screen unlock", userId);
    return { success: false, message: 'รหัสผ่านไม่ถูกต้อง' };
  } catch (e) { return { success: false, message: 'System Error: ' + e.message }; }
}

function getUserProfileById(userId) {
  const config = getDbConfig();
  const ss = SpreadsheetApp.openById(config.USER);
  const sheet = ss.getSheetByName("User");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(userId)) { 
      let profile = { 
        name: data[i][0], 
        userId: data[i][1], 
        group: data[i][4], 
        role: data[i][5], 
        unitIdRaw: data[i][6], 
        image: data[i][7] || "",
        allowedUnits: []       
      };

      try {
        if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
          const unitSS = SpreadsheetApp.openById(config.UNIT);
          const unitSheet = unitSS.getSheetByName("Units");
          if (unitSheet) {
            const unitData = unitSheet.getDataRange().getValues();
            const role = String(profile.role).toUpperCase();
            const isAdmin = role === 'ADMIN';

            for (let r = 1; r < unitData.length; r++) {
              const uGroup = String(unitData[r][0]).trim();
              const uName = String(unitData[r][1]).trim(); 
              
              if (isAdmin) {
                if (uName && !profile.allowedUnits.includes(uName)) profile.allowedUnits.push(uName);
              } else {
                if (uGroup === String(profile.group).trim()) {
                  if (uName && !profile.allowedUnits.includes(uName)) profile.allowedUnits.push(uName);
                }
              }
            }
          }
        }
      } catch (e) {
        console.error("Failed to load Units mapping: " + e.message);
        if (profile.unitIdRaw) profile.allowedUnits = String(profile.unitIdRaw).split(',').map(s => s.trim());
      }
      return profile;
    }
  }
  return null;
}

function getAvailableYears() {
  try {
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) return [];
    
    const dbSS = SpreadsheetApp.openById(config.CONFIG); 
    let sheet = dbSS.getSheetByName('Year_Config');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let years = [];
    for (let i = 1; i < data.length; i++) {
      if(data[i][0] && (data[i][2] === 'Connected' || !data[i][2])) years.push(String(data[i][0]));
    }
    return years.length > 0 ? years : [];
  } catch (e) { return []; }
}

// -------------------------------------------------------------------------
// 4.5 USER MANAGEMENT API (ADMIN SYSTEM)
// -------------------------------------------------------------------------
function apiGetUsersList(actionByUserId) {
  try {
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName("User");
    const data = sheet.getDataRange().getValues();
    
    let usersList = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) { // ถ้ามี UserID
        usersList.push({
          originalUserId: data[i][1], // ใช้เป็น Key อ้างอิงเวลาเซฟ
          name: data[i][0],
          userId: data[i][1],
          email: data[i][3],
          group: data[i][4],
          role: data[i][5],
          unitIdRaw: data[i][6] || '',
          status: data[i][8] || 'ปกติ',   // Col I (index 8)
          otpStatus: data[i][9] || 'ON' // Col J (index 9) - สถานะการใช้งาน OTP
        });
      }
    }
    return { success: true, data: usersList };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function apiSaveUserAdmin(payload, actionByUserId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.USER);
    const sheet = ss.getSheetByName("User");
    const data = sheet.getDataRange().getValues();
    
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(payload.originalUserId)) {
        rowIdx = i + 1;
        break;
      }
    }
    
    if (rowIdx === -1) throw new Error("ไม่พบข้อมูลผู้ใช้งานที่ต้องการแก้ไข หรืออาจถูกลบไปแล้ว");
    
    let role = String(payload.role).trim().toUpperCase();
    let unitIdRaw = String(payload.unitIdRaw).trim();
    
    if (role === 'ADMIN') unitIdRaw = 'ALL';
    else if (role === 'USER') unitIdRaw = '';
    
    sheet.getRange(rowIdx, 1).setValue(payload.name);
    sheet.getRange(rowIdx, 2).setValue(payload.userId);
    sheet.getRange(rowIdx, 4).setValue(payload.email);
    sheet.getRange(rowIdx, 5).setValue(payload.group);
    sheet.getRange(rowIdx, 6).setValue(role);
    sheet.getRange(rowIdx, 7).setValue(unitIdRaw);
    sheet.getRange(rowIdx, 9).setValue(payload.status);
    sheet.getRange(rowIdx, 10).setValue(payload.otpStatus || 'ON'); 
    
    SpreadsheetApp.flush();
    logSystem("Admin Action", `Updated user details for UserID: ${payload.userId}`, actionByUserId);
    
    return { success: true, message: "บันทึกข้อมูลผู้ใช้งานเรียบร้อยแล้ว" };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------------
// 5. DATABASE CONFIG MANAGER
// -------------------------------------------------------------------------
function apiGetDbConfig() {
  try {
    const config = getDbConfig();
    let years = [];
    let unitFolders = [];
    let deliveryNoteFolders = [];
    
    // ดึงข้อมูลปี Config
    try {
      if(config.CONFIG && !config.CONFIG.includes('ใส่_ID_ไฟล์')) {
        const configSS = SpreadsheetApp.openById(config.CONFIG); 
        let sheet = configSS.getSheetByName('Year_Config');
        if (sheet) {
          const data = sheet.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) {
            if(data[i][0]) {
              let status = data[i][2] || 'Connected'; 
              years.push({ year: String(data[i][0]), fileId: String(data[i][1]), status: status });
            }
          }
        }
      }
    } catch(e) { console.log("Year_Config fetch error:", e); }

    // ดึงข้อมูล Unit Folders & Delivery Note Folders
    try {
      if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
         const unitSS = SpreadsheetApp.openById(config.UNIT);
         let unitSheet = unitSS.getSheetByName('Units');
         if(unitSheet) {
             const data = unitSheet.getDataRange().getValues();
             for(let i = 1; i < data.length; i++) {
                 if(data[i][1]) { 
                     // Image Folders (Col D)
                     unitFolders.push({
                         name: String(data[i][1]).trim(),
                         folderId: String(data[i][3] || '').trim()
                     });
                     // Delivery Note Folders (Col E)
                     deliveryNoteFolders.push({
                         name: String(data[i][1]).trim(),
                         folderId: String(data[i][4] || '').trim()
                     });
                 }
             }
         }
      }
    } catch(e) { console.log("Unit_Folder fetch error:", e); }

    return { 
      success: true, 
      config: { 
        mainId: config.MAIN, 
        unitId: config.UNIT, 
        userId: config.USER, 
        configId: config.CONFIG, 
        logId: config.LOG, 
        folderProfile: config.FOLDER_PROFILE,
        folderLogo: config.FOLDER_LOGO,
        years: years,
        unitFolders: unitFolders,
        deliveryNoteFolders: deliveryNoteFolders
      } 
    };
  } catch (e) { return { success: false, message: e.message }; }
}

function apiSaveCoreDbConfig(payload, userId) {
  try {
    const props = PropertiesService.getScriptProperties();
    // 1. Core DBs
    if (payload.mainId) props.setProperty('DB_MAIN', payload.mainId.trim());
    if (payload.unitId) props.setProperty('DB_UNIT', payload.unitId.trim());
    if (payload.userId) props.setProperty('DB_USER', payload.userId.trim());
    if (payload.configId) props.setProperty('DB_CONFIG', payload.configId.trim());
    if (payload.logId) props.setProperty('DB_LOG', payload.logId.trim());
    
    // 2. Profile/Logo Folders
    if (payload.folderProfile !== undefined) props.setProperty('FOLDER_PROFILE', payload.folderProfile.trim());
    if (payload.folderLogo !== undefined) props.setProperty('FOLDER_LOGO', payload.folderLogo.trim());

    // 3. Unit Folders (Images) & Delivery Note Folders (PDF)
    const config = getDbConfig();
    if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
        const unitSS = SpreadsheetApp.openById(config.UNIT);
        let unitSheet = unitSS.getSheetByName('Units');
        if(unitSheet) {
            const data = unitSheet.getDataRange().getValues();
            for(let i = 1; i < data.length; i++) {
                let uName = String(data[i][1]).trim();
                
                // Save Image Folders -> Col D (Index 3 -> Range Column 4)
                if (payload.unitFolders && payload.unitFolders.length > 0) {
                    let matchImage = payload.unitFolders.find(u => u.name === uName);
                    if(matchImage) unitSheet.getRange(i + 1, 4).setValue(matchImage.folderId); 
                }

                // Save Delivery Note Folders -> Col E (Index 4 -> Range Column 5)
                if (payload.deliveryNoteFolders && payload.deliveryNoteFolders.length > 0) {
                    let matchPDF = payload.deliveryNoteFolders.find(u => u.name === uName);
                    if(matchPDF) unitSheet.getRange(i + 1, 5).setValue(matchPDF.folderId); 
                }
            }
        }
    }

    logSystem("Update DB Config", "Admin updated core database & folder configurations", userId);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function apiCreateYearSheet(year, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    checkDatabaseSetup();
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("กรุณาระบุ Sheet ID สำหรับไฟล์ Config ก่อนสร้างปีงบประมาณ");
    
    const configSS = SpreadsheetApp.openById(config.CONFIG); 
    let yearSheet = configSS.getSheetByName('Year_Config');
    
    if (!yearSheet) {
      yearSheet = configSS.insertSheet('Year_Config');
      yearSheet.appendRow(['Year', 'FileID', 'Status']);
      yearSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e2e8f0");
      yearSheet.setFrozenRows(1);
    }
    
    const data = yearSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(year)) {
        return { success: false, message: `มีการตั้งค่าไฟล์ของปี ${year} อยู่แล้วในระบบครับ` };
      }
    }
    
    // Create new Google Sheet for Transactions
    const fileName = `reAgentics_Transactions_${year}`;
    const newSS = SpreadsheetApp.create(fileName);
    const fileId = newSS.getId();
    
    let sheet = newSS.getSheets()[0];
    sheet.setName(String(year));
    // Updated Headers with Temp, Speed, PDF URL
    const headers = ['transactionID', 'timestamp', 'type', 'itemID', 'lot', 'expiry_Date', 'quantity', 'actionBy_UserID', 'Transport_Temp', 'Transport_Speed', 'Delivery_Note_URL'];
    sheet.appendRow(headers);
    sheet.getRange("A1:K1").setFontWeight("bold").setBackground("#f8fafc");
    sheet.setFrozenRows(1);
    
    yearSheet.appendRow([year, fileId, 'Connected']);
    
    logSystem("Create Year Sheet", `Created new transaction file for year ${year} (ID: ${fileId})`, userId);
    return { success: true, fileId: fileId, year: year, status: 'Connected' };
    
  } catch (e) { return { success: false, message: e.message }; } 
  finally { lock.releaseLock(); }
}

function apiManualAddYearSheet(year, fileId, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    checkDatabaseSetup();
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("Missing Config DB ID");
    
    const configSS = SpreadsheetApp.openById(config.CONFIG); 
    let yearSheet = configSS.getSheetByName('Year_Config');
    
    if (!yearSheet) {
      yearSheet = configSS.insertSheet('Year_Config');
      yearSheet.appendRow(['Year', 'FileID', 'Status']);
      yearSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e2e8f0");
      yearSheet.setFrozenRows(1);
    }
    
    const data = yearSheet.getDataRange().getValues();
    let isExist = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(year)) {
        yearSheet.getRange(i + 1, 2).setValue(fileId);
        yearSheet.getRange(i + 1, 3).setValue('Connected');
        isExist = true;
        break;
      }
    }
    
    if (!isExist) {
       yearSheet.appendRow([year, fileId, 'Connected']);
    }

    try {
      SpreadsheetApp.openById(fileId);
    } catch(err) {
      throw new Error("ไม่สามารถเข้าถึงไฟล์ Sheet ID ที่ระบุได้ กรุณาตรวจสอบสิทธิ์การเข้าถึง");
    }
    
    logSystem("Manual Connect Year", `Manually connected file for year ${year} (ID: ${fileId})`, userId);
    return { success: true, fileId: fileId, year: year, status: 'Connected' };
    
  } catch (e) { return { success: false, message: e.message }; } 
  finally { lock.releaseLock(); }
}

function apiDisconnectYearSheet(year, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("Missing Config DB ID");

    const configSS = SpreadsheetApp.openById(config.CONFIG); 
    let yearSheet = configSS.getSheetByName('Year_Config');

    if (!yearSheet) return { success: false, message: 'ไม่พบตารางตั้งค่าปี' };

    const data = yearSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(year)) {
        yearSheet.getRange(i + 1, 3).setValue('Disconnected'); 
        logSystem("Disconnect Year", `Disconnected transaction file for year ${year}`, userId);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลปีที่ต้องการระงับการเชื่อมต่อ' };
  } catch (e) { return { success: false, message: e.message }; } 
  finally { lock.releaseLock(); }
}

function apiDeleteYearSheet(year, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("Missing Config DB ID");

    const configSS = SpreadsheetApp.openById(config.CONFIG); 
    let yearSheet = configSS.getSheetByName('Year_Config');

    if (!yearSheet) return { success: false, message: 'ไม่พบตารางตั้งค่าปี' };

    const data = yearSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(year)) {
        yearSheet.deleteRow(i + 1); 
        logSystem("Delete Year Link", `Removed year ${year} from config database`, userId);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลปีที่ต้องการลบ' };
  } catch (e) { return { success: false, message: e.message }; } 
  finally { lock.releaseLock(); }
}

// -------------------------------------------------------------------------
// 6. IMAGE, PROFILE & PDF UPLOAD APIs
// -------------------------------------------------------------------------
function apiChangePassword(userId, newPassword) {
  try {
    const config = getDbConfig();
    const userSS = SpreadsheetApp.openById(config.USER); 
    const sheet = userSS.getSheetByName('User'); 
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { 
      if (String(data[i][1]) === String(userId)) { 
        sheet.getRange(i + 1, 3).setValue(newPassword); 
        logSystem("Change Password", "Updated password", userId); 
        return { status: 'success' }; 
      } 
    }
    return { status: 'error', message: 'User not found' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiSaveProfileImage(userId, base64Data) {
  try {
    const config = getDbConfig();
    const userSS = SpreadsheetApp.openById(config.USER); 
    const sheet = userSS.getSheetByName('User'); 
    const data = sheet.getDataRange().getDisplayValues();
    
    let rowIndex = -1; let oldFileUrl = "";
    for (let i = 1; i < data.length; i++) { 
      if (String(data[i][1]) === String(userId)) { rowIndex = i + 1; oldFileUrl = data[i][7]; break; } 
    }
    if (rowIndex === -1) return { status: 'error', message: 'User not found' };
    
    deleteOldDriveFile(oldFileUrl);
    
    let folder;
    if (config.FOLDER_PROFILE) {
        try { folder = DriveApp.getFolderById(config.FOLDER_PROFILE); } catch(e) {}
    }
    if (!folder) {
        const folderName = "reAgentics_Profiles";
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) folder = folders.next();
        else folder = DriveApp.createFolder(folderName);
    }
    
    const contentType = base64Data.substring(5, base64Data.indexOf(';')); 
    let ext = "png"; if (contentType.includes("gif")) ext = "gif"; else if (contentType.includes("jpeg") || contentType.includes("jpg")) ext = "jpg";
    
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
    const blob = Utilities.newBlob(bytes, contentType, `profile_${userId}_${Date.now()}.${ext}`); 
    const file = folder.createFile(blob); 
    
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) { console.log("Share err:", e); }
    
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s800`; 
    sheet.getRange(rowIndex, 8).setValue(fileUrl); 
    
    logSystem("Change Profile Pic", "Updated profile image", userId); 
    return { status: 'success', url: fileUrl };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiSaveSystemLogo(base64Data, userId) {
  try {
    const config = getDbConfig();
    if(!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("กรุณาตั้งค่า Config ID ในหน้าตั้งค่าฐานข้อมูลก่อน");
    const sysSS = SpreadsheetApp.openById(config.CONFIG); 
    let sheet = sysSS.getSheetByName('App_Logo');
    if (!sheet) { sheet = sysSS.insertSheet('App_Logo'); }
    
    let oldFileUrl = "";
    try {
        oldFileUrl = sheet.getRange("B2").getValue();
        if (!oldFileUrl) {
            oldFileUrl = sheet.getRange("B1").getValue(); 
        }
    } catch(e) {}

    deleteOldDriveFile(oldFileUrl);

    let folder;
    if (config.FOLDER_LOGO) {
        try { folder = DriveApp.getFolderById(config.FOLDER_LOGO); } catch(e) {}
    }
    if (!folder) {
        const folderName = "reAgentics_Logos";
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) folder = folders.next();
        else folder = DriveApp.createFolder(folderName);
    }
    
    const contentType = base64Data.substring(5, base64Data.indexOf(';')); 
    let ext = "png"; if (contentType.includes("gif")) ext = "gif"; else if (contentType.includes("jpeg") || contentType.includes("jpg")) ext = "jpg";
    
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7)); 
    const blob = Utilities.newBlob(bytes, contentType, `app_logo_${Date.now()}.${ext}`);
    const file = folder.createFile(blob); 
    
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) { console.log("Share err:", e); }
    
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s800`;
    
    sheet.clear();
    sheet.getRange("A1").setValue("Name").setFontWeight("bold");
    sheet.getRange("B1").setValue("Url").setFontWeight("bold");
    sheet.getRange("A2").setValue("MainLogo");
    sheet.getRange("B2").setValue(fileUrl);
    
    SpreadsheetApp.flush();
    
    logSystem("Change Logo", "Updated system logo", userId); 
    return { status: 'success', url: fileUrl };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiUploadReagentImage(base64Data, unitName, oldUrl, userId) {
    try {
        const config = getDbConfig();
        let targetFolderId = "";
        
        if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
             const unitSS = SpreadsheetApp.openById(config.UNIT);
             let unitSheet = unitSS.getSheetByName('Units');
             if(unitSheet) {
                 const data = unitSheet.getDataRange().getValues();
                 for(let i = 1; i < data.length; i++) {
                     if(String(data[i][1]).trim() === String(unitName).trim()) {
                         targetFolderId = String(data[i][3] || '').trim(); // Col D (Images)
                         break;
                     }
                 }
             }
        }
        
        deleteOldDriveFile(oldUrl);

        let folder;
        if (targetFolderId) {
            try { folder = DriveApp.getFolderById(targetFolderId); } catch(e) { console.log("Invalid Unit Folder ID"); }
        }
        
        if (!folder) {
            const folderName = "reAgentics_Items_" + unitName;
            const folders = DriveApp.getFoldersByName(folderName);
            if (folders.hasNext()) folder = folders.next();
            else folder = DriveApp.createFolder(folderName);
        }

        const contentType = base64Data.substring(5, base64Data.indexOf(';')); 
        let ext = "png"; if (contentType.includes("gif")) ext = "gif"; else if (contentType.includes("jpeg") || contentType.includes("jpg")) ext = "jpg";
        
        const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
        const blob = Utilities.newBlob(bytes, contentType, `item_${Date.now()}.${ext}`); 
        const file = folder.createFile(blob); 
        
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) { console.log("Share err:", e); }
        
        const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s800`; 
        
        logSystem("Upload Reagent Image", `Uploaded image to folder for unit: ${unitName}`, userId);
        return { success: true, url: fileUrl };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

function apiUploadDeliveryNote(base64Data, unitName, userId) {
    try {
        const config = getDbConfig();
        let targetFolderId = "";
        let prefix = unitName; // Default to name if prefix not found
        
        // ค้นหา Folder ID และ Prefix จาก reAgentics_Units
        if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
             const unitSS = SpreadsheetApp.openById(config.UNIT);
             let unitSheet = unitSS.getSheetByName('Units');
             if(unitSheet) {
                 const data = unitSheet.getDataRange().getValues();
                 for(let i = 1; i < data.length; i++) {
                     if(String(data[i][1]).trim() === String(unitName).trim()) {
                         prefix = String(data[i][2]).trim() || unitName; // Col C (Prefix)
                         targetFolderId = String(data[i][4] || '').trim(); // Col E (Delivery Notes)
                         break;
                     }
                 }
             }
        }
        
        let folder;
        if (targetFolderId) {
            try { folder = DriveApp.getFolderById(targetFolderId); } catch(e) { console.log("Invalid Delivery Note Folder ID"); }
        }
        
        if (!folder) {
            const folderName = "reAgentics_DeliveryNotes_" + unitName;
            const folders = DriveApp.getFoldersByName(folderName);
            if (folders.hasNext()) folder = folders.next();
            else folder = DriveApp.createFolder(folderName);
        }

        // ตั้งชื่อไฟล์รูปแบบ {Prefix}_{YYYY-MM-DD}
        const today = new Date();
        const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
        // สุ่มเลขต่อท้ายเล็กน้อยป้องกันการอัพโหลดไฟล์ซ้ำในวันเดียวกันแล้วชื่อชนกัน
        const randomStr = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
        const fileName = `${prefix}_${dateStr}_${randomStr}.pdf`;

        // ตัด header data:application/pdf;base64, ออก
        const contentType = 'application/pdf';
        const base64Clean = base64Data.split(',')[1];
        
        const bytes = Utilities.base64Decode(base64Clean);
        const blob = Utilities.newBlob(bytes, contentType, fileName); 
        const file = folder.createFile(blob); 
        
        // แชร์เพื่อให้ดูได้ (ใส่ try-catch ป้องกัน error กรณีโฟลเดอร์ข้ามบัญชี)
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) { console.log("Share err:", e); }
        
        const fileUrl = file.getUrl(); // ใช้ getUrl ตรงๆ สำหรับ PDF
        
        logSystem("Upload Delivery Note", `Uploaded PDF to folder for unit: ${unitName}`, userId);
        return { success: true, url: fileUrl };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}


// -------------------------------------------------------------------------
// 7. STICKER CONFIG API (reAgentics_Config)
// -------------------------------------------------------------------------
function apiGetStickerConfig() {
  try {
    const dbConfig = getDbConfig();
    if(!dbConfig.CONFIG || dbConfig.CONFIG.includes('ใส่_ID_ไฟล์')) {
      return { status: 'success', config: getDefaultStickerConfig() };
    }

    const sysSS = SpreadsheetApp.openById(dbConfig.CONFIG);
    let sheet = sysSS.getSheetByName('Sticker_Config');
    let config = getDefaultStickerConfig();

    if (!sheet) {
      sheet = sysSS.insertSheet('Sticker_Config');
      sheet.appendRow(['Key', 'Value', 'Description']);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f1f5f9");
      
      const descriptions = {
        width: "ความกว้างของสติ๊กเกอร์ (mm)", height: "ความสูงของสติ๊กเกอร์ (mm)",
        autoPrintCount: "จำนวนแผ่นที่จะพิมพ์อัตโนมัติเมื่อรับเข้าเสร็จ", manualPrintCount: "จำนวนแผ่นที่จะพิมพ์เมื่อกดปุ่มพิมพ์จากหน้าจอ",
        barcodeHeight: "ความสูงของเส้นบาร์โค้ด (px)", barcodeWidth: "สเกลความกว้างเส้นบาร์โค้ด", layoutJSON: "พิกัด X/Y, ขนาด, การหมุน ของแต่ละองค์ประกอบ (ห้ามแก้ไขด้วยมือ)"
      };

      for (let key in config) { sheet.appendRow([key, config[key], descriptions[key]]); }
      sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 250); sheet.setColumnWidth(3, 300);
    } else {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (config.hasOwnProperty(data[i][0])) {
          let val = data[i][1];
          if (val === 'true' || val === true) config[data[i][0]] = true;
          else if (val === 'false' || val === false) config[data[i][0]] = false;
          else if (data[i][0] === 'layoutJSON') config[data[i][0]] = String(val);
          else config[data[i][0]] = Number(val) || val;
        }
      }
    }
    return { status: 'success', config: config };
  } catch (e) { return { status: 'error', message: 'Get Sticker Config Error: ' + e.message }; }
}

function getDefaultStickerConfig() {
  return {
    width: 50, height: 30, autoPrintCount: 2, manualPrintCount: 1, barcodeHeight: 35, barcodeWidth: 1.5,   
    layoutJSON: JSON.stringify({
      cyto: { x: 25, y: 4, size: 11, rot: 0, visible: true, bold: true, font: 'Montserrat' }, 
      name: { x: 25, y: 9, size: 9, rot: 0, visible: false, bold: false, font: 'Montserrat' }, 
      age:  { x: 10, y: 26, size: 10, rot: 0, visible: true, bold: true, font: 'Roboto Mono' }, 
      spec: { x: 40, y: 26, size: 10, rot: 0, visible: true, bold: true, font: 'Roboto Mono' }, 
      unit: { x: 25, y: 28, size: 8, rot: 0, visible: false, bold: false, font: 'Montserrat' }, 
      bar:  { x: 25, y: 14, rot: 0, visible: true, width: 1.5 } 
    })
  };
}

function apiSaveStickerConfig(newConfig, userId) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const dbConfig = getDbConfig();
    if(!dbConfig.CONFIG || dbConfig.CONFIG.includes('ใส่_ID_ไฟล์')) throw new Error("กรุณาตั้งค่า Config ID ในหน้าตั้งค่าฐานข้อมูลก่อน");

    const sysSS = SpreadsheetApp.openById(dbConfig.CONFIG);
    let sheet = sysSS.getSheetByName('Sticker_Config');
    if (!sheet) return { status: 'error', message: 'Sticker_Config sheet not found' };

    const data = sheet.getDataRange().getValues();
    
    for (let key in newConfig) {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(newConfig[key]);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow([key, newConfig[key], "Auto-generated field"]);
    }

    logSystem("Update Config", "Admin updated Sticker Configuration Layout", userId);
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: 'Save Sticker Config Error: ' + e.message };
  } finally { lock.releaseLock(); }
}

// -------------------------------------------------------------------------
// 8. DATA FETCHING (DROPDOWNS & AUTO-IDs) 
// -------------------------------------------------------------------------
function apiGetFormOptions() {
  try {
    const config = getDbConfig();
    if(!config.UNIT || config.UNIT.includes('ใส่_ID_ไฟล์')) return { success: false, message: 'ไม่ได้ตั้งค่า Unit DB' };
    
    const unitSS = SpreadsheetApp.openById(config.UNIT);
    const options = { units: [], reagUnits: [], analyzers: [], storageLocations: [], companies: [] };
    
    // 1. Units - เพิ่ม Group มาด้วย
    let sheet = unitSS.getSheetByName('Units');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) {
        if(data[i][1]) {
          options.units.push({ 
            group: String(data[i][0]).trim(),
            name: String(data[i][1]).trim(), 
            prefix: String(data[i][2]).trim() 
          });
        }
      }
    }
    
    // 2. ReagUnits
    sheet = unitSS.getSheetByName('ReagUnits');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) {
        if(data[i][0]) options.reagUnits.push(data[i][0]);
      }
    }
    
    // 3. Analyzers
    sheet = unitSS.getSheetByName('Analyzers');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) {
        if(data[i][0] && data[i][1]) options.analyzers.push({ unit: data[i][0], name: data[i][1] });
      }
    }
    
    // 4. storageLocation
    sheet = unitSS.getSheetByName('storageLocation');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) {
        if(data[i][0]) options.storageLocations.push(data[i][0]);
      }
    }

    // 5. Companies (นำเข้า/จำหน่าย)
    let companySheet = unitSS.getSheetByName('Company');
    if (companySheet) {
      const data = companySheet.getRange("A2:A" + companySheet.getLastRow()).getValues();
      const uniqueCompanies = [...new Set(data.map(r => String(r[0]).trim()).filter(String))];
      options.companies = uniqueCompanies;
    }
    
    return { success: true, data: options };
  } catch(e) { return { success: false, message: e.message }; }
}

function apiGetNextItemID(unitName) {
  try {
    const config = getDbConfig();
    const unitSS = SpreadsheetApp.openById(config.UNIT);
    const unitSheet = unitSS.getSheetByName('Units');
    if (!unitSheet) throw new Error("ไม่พบแท็บ Units ในฐานข้อมูลหน่วยงาน");
    
    const unitData = unitSheet.getDataRange().getValues();
    let prefix = "";
    for(let i=1; i<unitData.length; i++) {
      if(String(unitData[i][1]).trim() === String(unitName).trim()) { 
        prefix = String(unitData[i][2]).trim(); 
        break; 
      }
    }
    
    if(!prefix) return { success: false, message: "ไม่พบรหัส Prefix สำหรับหน่วยงานนี้" };
    
    // หายอดรหัสล่าสุดจาก Main DB
    const mainSS = SpreadsheetApp.openById(config.MAIN);
    const itemSheet = mainSS.getSheetByName('Items');
    let maxNum = 0;
    
    if (itemSheet) {
      const itemData = itemSheet.getDataRange().getValues();
      for(let i=1; i<itemData.length; i++) {
        const id = String(itemData[i][0]).trim();
        if(id.startsWith(prefix + "-")) {
          const numPart = parseInt(id.replace(prefix + "-", ""), 10);
          if(!isNaN(numPart) && numPart > maxNum) maxNum = numPart;
        }
      }
    }
    
    const nextId = prefix + "-" + String(maxNum + 1).padStart(3, '0');
    return { success: true, nextId: nextId };
  } catch(e) { return { success: false, message: e.message }; }
}

function apiGetActiveLots(itemID) {
  try {
    const config = getDbConfig();
    const mainSS = SpreadsheetApp.openById(config.MAIN);
    const stockSheet = mainSS.getSheetByName('Stock_Balance');
    if(!stockSheet) return { success: true, lots: [] };
    
    const data = stockSheet.getDataRange().getValues();
    const lots = [];
    // ค้นหาเฉพาะที่มี Quantity (Col E - index 4) >= 1
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0]).trim() === String(itemID).trim() && Number(data[i][4]) >= 1) {
        lots.push({
          lot: String(data[i][2]),
          exp: safeString(data[i][3]),
          qty: Number(data[i][4]),
          unit: String(data[i][5])
        });
      }
    }
    return { success: true, lots: lots };
  } catch(e) { return { success: false, message: e.message }; }
}


// -------------------------------------------------------------------------
// 9. INVENTORY & TRANSACTION ENGINE (v0.8.8 AI Core)
// -------------------------------------------------------------------------
function getItemsData() {
  try {
    checkDatabaseSetup();
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.MAIN);
    
    // Fetch unit mapping to get Groups
    let unitGroupMap = {};
    try {
      if(config.UNIT && !config.UNIT.includes('ใส่_ID_ไฟล์')) {
        const unitSS = SpreadsheetApp.openById(config.UNIT);
        const unitSheet = unitSS.getSheetByName('Units');
        if (unitSheet) {
          const unitData = unitSheet.getDataRange().getValues();
          for (let i=1; i<unitData.length; i++) {
             let group = String(unitData[i][0]).trim();
             let uName = String(unitData[i][1]).trim();
             if(uName) unitGroupMap[uName] = group;
          }
        }
      }
    } catch(e) { console.error("Error fetching unit groups:", e); }

    // 1. ดึงข้อมูลจากฐานข้อมูล Items (ข้อมูล Master Data)
    const itemSheet = ss.getSheetByName("Items");
    if (!itemSheet) throw new Error("ไม่พบชีต 'Items' ในฐานข้อมูลหลัก");
    const itemData = itemSheet.getDataRange().getDisplayValues();
    
    // 2. ดึงข้อมูลจากฐานข้อมูล Stock_Balance (ยอดคงคลังทั้งหมด)
    let stockSheet = ss.getSheetByName("Stock_Balance");
    let stockData = [];
    if (stockSheet) {
      stockData = stockSheet.getDataRange().getValues();
    }
    
    // 3. รวมยอดคงเหลือ (Aggregate) แยกตาม ItemID
    const balanceMap = {};
    if (stockData.length > 1) {
      for (let r = 1; r < stockData.length; r++) {
        let itemId = String(stockData[r][0]).trim();
        let qty = Number(stockData[r][4]) || 0;
        
        if (!balanceMap[itemId]) balanceMap[itemId] = 0;
        balanceMap[itemId] += qty;
      }
    }
    
    // 4. Map ข้อมูลส่งให้ Frontend
    const resultData = [];
    for (let i = 1; i < itemData.length; i++) {
      let row = itemData[i];
      let itemId = String(row[0]).trim();
      if (!itemId) continue; 
      
      let currentBalance = balanceMap[itemId] || 0;
      let status = String(row[8]).trim(); 
      let imageUrl = String(row[9] || '').trim(); // คอลัมน์ J (index 9)
      
      let company = String(row[10] || '').trim(); // คอลัมน์ K (index 10)
      let price = Number(row[11]) || 0;           // คอลัมน์ L (index 11)

      let uName = String(row[4]).trim();
      let group = unitGroupMap[uName] || 'ไม่ระบุ';
      
      resultData.push({
        itemID: itemId, 
        itemName: row[1], 
        minLevel: row[2], 
        unit: row[3], 
        unitID: uName, 
        group: group,
        analyzer: row[5], 
        storageTemp: row[6], 
        storageLocation: row[7], 
        status: status,
        image: imageUrl,
        company: company,
        price: price,
        balance: currentBalance 
      });
    }
    
    return { success: true, data: resultData };
  } catch (error) { 
    return { success: false, message: error.message }; 
  }
}

function apiRegisterNewItem(payload, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.MAIN);
    const sheet = ss.getSheetByName("Items");
    
    const existingData = sheet.getRange("A:A").getValues().flat();
    if(existingData.includes(payload.itemID)) {
       throw new Error(`รหัสน้ำยา ${payload.itemID} มีอยู่ในระบบแล้ว กรุณาลองใหม่อีกครั้ง`);
    }
    
    // Col A: itemID, B: itemName, C: minLevel, D: unit, E: unitID, F: analyzer, G: storageTemp, H: storageLocation, I: status, J: image, K: company, L: price
    sheet.appendRow([
      payload.itemID, 
      payload.itemName, 
      payload.minLevel, 
      payload.unit, 
      payload.unitID, 
      payload.analyzer, 
      payload.storageTemp, 
      payload.storageLocation, 
      'Active',
      payload.image || '',
      payload.company || '',
      payload.price || 0
    ]);
    
    SpreadsheetApp.flush(); 
    logSystem("Register Item", `Registered new item: ${payload.itemID}`, userId);
    return { success: true, message: 'ลงทะเบียนน้ำยาใหม่เรียบร้อยแล้ว' };
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function apiUpdateItem(payload, userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getDbConfig();
    const ss = SpreadsheetApp.openById(config.MAIN);
    const sheet = ss.getSheetByName("Items");
    const data = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    for(let i=1; i<data.length; i++) {
       if(String(data[i][0]).trim() === String(payload.itemID).trim()) {
           targetRow = i + 1;
           break;
       }
    }
    
    if(targetRow === -1) throw new Error("ไม่พบรายการที่ต้องการแก้ไข");
    
    sheet.getRange(targetRow, 2).setValue(payload.itemName);
    sheet.getRange(targetRow, 3).setValue(payload.minLevel);
    sheet.getRange(targetRow, 4).setValue(payload.unit);
    sheet.getRange(targetRow, 5).setValue(payload.unitID);
    sheet.getRange(targetRow, 6).setValue(payload.analyzer);
    sheet.getRange(targetRow, 7).setValue(payload.storageTemp);
    sheet.getRange(targetRow, 8).setValue(payload.storageLocation);
    sheet.getRange(targetRow, 9).setValue(payload.status);
    
    if (payload.image !== undefined) {
        sheet.getRange(targetRow, 10).setValue(payload.image); 
    }

    sheet.getRange(targetRow, 11).setValue(payload.company || '');
    sheet.getRange(targetRow, 12).setValue(payload.price || 0);
    
    SpreadsheetApp.flush();
    logSystem("Update Item", `Updated details for item: ${payload.itemID} | Status: ${payload.status}`, userId);
    return { success: true, message: 'อัปเดตข้อมูลน้ำยาเรียบร้อยแล้ว' };
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function processTransaction(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const config = getDbConfig();
    const { type, yearSheetId, userId, items, transportSpeed, deliveryNoteUrl } = payload;
    const dbSS = SpreadsheetApp.openById(config.MAIN);
    
    let stockSheet = dbSS.getSheetByName('Stock_Balance');
    if (!stockSheet) {
      stockSheet = dbSS.insertSheet('Stock_Balance');
      stockSheet.appendRow(['ItemID', 'ItemName', 'Lot', 'EXP', 'Qty', 'Unit', 'LastUpdate']);
      stockSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#f8fafc");
      stockSheet.setFrozenRows(1);
    }
    
    let stockData = stockSheet.getDataRange().getValues();
    
    for (let i = 0; i < items.length; i++) {
      let item = items[i];
      let reqItemId = String(item.itemID).trim();
      let reqLot = String(item.lot).trim().toUpperCase(); 
      let reqQty = Number(item.qty);

      let rowToUpdate = -1; 
      let currentQty = 0;
      
      // ค้นหาแถวที่มี ItemID และ Lot ตรงกัน
      for (let r = 1; r < stockData.length; r++) {
        let sheetItemId = String(stockData[r][0]).trim();
        let sheetLot = String(stockData[r][2]).trim().toUpperCase();

        if (sheetItemId === reqItemId && sheetLot === reqLot) {
          rowToUpdate = r + 1; 
          currentQty = Number(stockData[r][4]) || 0; 
          break;
        }
      }
      
      let newQty = currentQty;

      if (type === 'RECEIVE') {
        newQty = currentQty + reqQty;
        if (rowToUpdate === -1) {
          stockSheet.appendRow([item.itemID, item.itemName, reqLot, item.exp, newQty, item.unit, new Date()]);
        } else {
          stockSheet.getRange(rowToUpdate, 5).setValue(newQty); 
          stockSheet.getRange(rowToUpdate, 7).setValue(new Date()); 
        }

      } else if (type === 'DISPENSE') {
        if (rowToUpdate === -1) {
          throw new Error(`ไม่พบข้อมูล Lot: ${reqLot} ของน้ำยารหัส ${reqItemId} ในคลัง กรุณาตรวจสอบให้แน่ใจ`);
        }
        if (currentQty < reqQty) {
          throw new Error(`ยอดคงเหลือของ ${reqItemId} (Lot: ${reqLot}) ไม่เพียงพอ (มีอยู่ ${currentQty} แต่ต้องการเบิก ${reqQty})`);
        }

        newQty = currentQty - reqQty;
        stockSheet.getRange(rowToUpdate, 5).setValue(newQty);
        stockSheet.getRange(rowToUpdate, 7).setValue(new Date());
      }

      // ส่ง payload เพิ่มเติมไปใน logRealTransaction
      logRealTransaction(yearSheetId, type, item, userId, transportSpeed, deliveryNoteUrl);
    }
    
    SpreadsheetApp.flush(); 
    
    logSystem("Transaction Success", `Processed ${type} for ${items.length} items`, userId);
    return { success: true, message: `บันทึกรายการ ${type === 'RECEIVE' ? 'รับเข้า' : 'เบิกใช้'} จำนวน ${items.length} รายการ สำเร็จ` };
  } catch (e) {
    logSystem("Transaction Failed", e.message, payload.userId || "Unknown");
    return { success: false, message: e.message };
  } finally { lock.releaseLock(); }
}

function logRealTransaction(year, type, item, userId, transportSpeed, deliveryNoteUrl) {
  const config = getDbConfig();
  if (!config.CONFIG || config.CONFIG.includes('ใส่_ID_ไฟล์')) return;
  
  const configSS = SpreadsheetApp.openById(config.CONFIG);
  let yearSheet = configSS.getSheetByName('Year_Config');
  if (!yearSheet) return;
  
  let transFileId = "";
  let data = yearSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(year)) { 
      if(data[i][2] === 'Disconnected') throw new Error("ไฟล์ปี " + year + " ถูกระงับการเชื่อมต่อ กรุณาเชื่อมต่อก่อนทำรายการ");
      transFileId = data[i][1]; 
      break; 
    }
  }
  if (!transFileId) throw new Error("ไม่พบไฟล์ Transactions สำหรับปี " + year);
  
  const transSS = SpreadsheetApp.openById(transFileId);
  let sheet = transSS.getSheetByName(String(year));
  if (!sheet) {
    sheet = transSS.insertSheet(String(year));
    // Header updated with Transport_Temp, Transport_Speed, Delivery_Note_URL
    sheet.appendRow(['transactionID', 'timestamp', 'type', 'itemID', 'lot', 'expiry_Date', 'quantity', 'actionBy_UserID', 'Transport_Temp', 'Transport_Speed', 'Delivery_Note_URL']);
    sheet.getRange("A1:K1").setFontWeight("bold").setBackground("#f8fafc");
    sheet.setFrozenRows(1);
  }
  
  let transId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMddHHmmss") + Math.floor(Math.random() * 1000);
  
  // เตรียมข้อมูลสำหรับแถวใหม่
  const rowData = [
    transId, 
    new Date(), 
    type, 
    item.itemID, 
    item.lot.trim().toUpperCase(), 
    item.exp || "-", 
    item.qty, 
    userId,
    item.transportTemp || "",     // Col I
    transportSpeed || "",         // Col J
    deliveryNoteUrl || ""         // Col K
  ];
  
  sheet.appendRow(rowData);
}
