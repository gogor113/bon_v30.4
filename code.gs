// ==================== BON UTANG WARUNG v36.0 ====================
// WebApp URL: Setelah deploy, copy URL ke hidden config di HTML
// Fitur: Auto delete data pelanggan LUNAS (tanpa sisa & tanpa kembalian)
// Jika ada kembalian → data tetap dengan sisa utang 0

const SHEET_NAME_BONS = "DataBon";
const SHEET_NAME_PAYMENTS = "DataPembayaran";
const SHEET_NAME_USERS = "DataUsers";
const SHEET_NAME_SYNC_LOG = "SyncLog";

function doGet() {
  return doPost({
    postData: { contents: "" }
  });
}

function doPost(e) {
  try {
    const data = e.parameter || {};
    const action = data.action;
    
    if (!action) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "No action specified" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Initialize sheets if not exist
    initializeSheets();
    
    // Route actions
    if (action === "testConnection") {
      return handleTestConnection();
    } 
    else if (action === "mergeBackupV35" || action === "restoreV35") {
      const jsonData = JSON.parse(data.data || "{}");
      if (action === "mergeBackupV35") {
        return handleMergeBackupV36(jsonData);
      } else {
        return handleRestoreV36(data.username);
      }
    }
    else if (action === "syncUserAuth") {
      const authData = JSON.parse(data.data || "{}");
      return handleSyncUserAuth(authData);
    }
    else if (action === "getUserAuth") {
      return handleGetUserAuth(data.username);
    }
    else if (action === "syncBonV35") {
      return handleSyncBonV36(data.username, JSON.parse(data.bonData || "{}"));
    }
    else if (action === "syncPaymentV35") {
      return handleSyncPaymentV36(data.username, JSON.parse(data.paymentData || "{}"));
    }
    else if (action === "cleanupLunasOtomatis") {
      return handleCleanupLunasOtomatis(data.username);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: "Unknown action" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error("Error in doPost:", error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss.getSheetByName(SHEET_NAME_BONS)) {
    const sheet = ss.insertSheet(SHEET_NAME_BONS);
    sheet.appendRow(["uniqueId", "username", "namaPelanggan", "total", "items", "waktu", "lastModified", "isActive"]);
  }
  
  if (!ss.getSheetByName(SHEET_NAME_PAYMENTS)) {
    const sheet = ss.insertSheet(SHEET_NAME_PAYMENTS);
    sheet.appendRow(["uniqueId", "username", "namaPelanggan", "jumlah", "waktu", "lastModified"]);
  }
  
  if (!ss.getSheetByName(SHEET_NAME_USERS)) {
    const sheet = ss.insertSheet(SHEET_NAME_USERS);
    sheet.appendRow(["username", "password", "security_pet", "security_hobby", "security_food", "registeredAt", "lastLogin", "isActive"]);
  }
  
  if (!ss.getSheetByName(SHEET_NAME_SYNC_LOG)) {
    const sheet = ss.insertSheet(SHEET_NAME_SYNC_LOG);
    sheet.appendRow(["timestamp", "username", "action", "details"]);
  }
}

function handleTestConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      message: "Cloud Backup v36.0 ACTIVE", 
      spreadsheetId: ss.getId(),
      sheets: ss.getSheets().map(s => s.getName()),
      version: "36.0"
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleMergeBackupV36(data) {
  const username = data.username;
  const remoteBons = data.semuaBon || [];
  const remotePayments = data.pembayaran || [];
  
  // Process Bons
  const existingBons = getAllBonsByUser(username);
  const mergedBons = mergeBonDataV36(existingBons, remoteBons);
  saveAllBonsToSheet(username, mergedBons);
  
  // Process Payments
  const existingPayments = getAllPaymentsByUser(username);
  const mergedPayments = mergePaymentDataV36(existingPayments, remotePayments);
  saveAllPaymentsToSheet(username, mergedPayments);
  
  // Auto cleanup LUNAS data
  const cleanupResult = cleanupLunasData(username);
  
  addSyncLog(username, "mergeBackup", `Merged ${mergedBons.length} bons, ${mergedPayments.length} payments`);
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      message: "Data merged successfully v36.0",
      cleanedUp: cleanupResult
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRestoreV36(username) {
  const allBons = getAllBonsByUser(username);
  const allPayments = getAllPaymentsByUser(username);
  
  // Filter only active bons (not fully paid or has change)
  const activeBons = filterActiveBonsV36(allBons, allPayments);
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      semuaBon: activeBons,
      pembayaran: allPayments,
      lastModified: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSyncBonV36(username, bonData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_BONS);
  const existingRow = findRowByUniqueId(sheet, bonData.uniqueId);
  
  if (existingRow) {
    // Update existing
    const rowNum = existingRow;
    sheet.getRange(rowNum, 3).setValue(bonData.namaPelanggan);
    sheet.getRange(rowNum, 4).setValue(bonData.total);
    sheet.getRange(rowNum, 5).setValue(JSON.stringify(bonData.items));
    sheet.getRange(rowNum, 6).setValue(bonData.waktu);
    sheet.getRange(rowNum, 7).setValue(bonData.lastModified);
    sheet.getRange(rowNum, 8).setValue(true);
  } else {
    // Insert new
    sheet.appendRow([
      bonData.uniqueId,
      username,
      bonData.namaPelanggan,
      bonData.total,
      JSON.stringify(bonData.items),
      bonData.waktu,
      bonData.lastModified,
      true
    ]);
  }
  
  // Auto cleanup after sync
  cleanupLunasData(username);
  
  addSyncLog(username, "syncBon", `Synced bon ${bonData.uniqueId}`);
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSyncPaymentV36(username, paymentData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PAYMENTS);
  const existingRow = findRowByUniqueId(sheet, paymentData.uniqueId);
  
  if (existingRow) {
    const rowNum = existingRow;
    sheet.getRange(rowNum, 3).setValue(paymentData.namaPelanggan);
    sheet.getRange(rowNum, 4).setValue(paymentData.jumlah);
    sheet.getRange(rowNum, 5).setValue(paymentData.waktu);
    sheet.getRange(rowNum, 6).setValue(paymentData.lastModified);
  } else {
    sheet.appendRow([
      paymentData.uniqueId,
      username,
      paymentData.namaPelanggan,
      paymentData.jumlah,
      paymentData.waktu,
      paymentData.lastModified
    ]);
  }
  
  // Auto cleanup after payment sync
  cleanupLunasData(username);
  
  addSyncLog(username, "syncPayment", `Synced payment ${paymentData.uniqueId}`);
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSyncUserAuth(authData) {
  const action = authData.subAction || authData.action;
  const username = authData.username.toLowerCase();
  const userData = authData.userData;
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USERS);
  const existingRow = findUserRow(sheet, username);
  
  if (action === "register") {
    if (existingRow) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "User already exists" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    sheet.appendRow([
      username,
      userData.password,
      userData.security.pet,
      userData.security.hobby,
      userData.security.food,
      userData.registeredAt || new Date().toISOString(),
      new Date().toISOString(),
      true
    ]);
    
    addSyncLog(username, "register", "New user registered");
    
  } else if (action === "resetPassword") {
    if (!existingRow) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "User not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const rowNum = existingRow;
    sheet.getRange(rowNum, 2).setValue(userData.password);
    sheet.getRange(rowNum, 7).setValue(new Date().toISOString());
    
    addSyncLog(username, "resetPassword", "Password reset");
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleGetUserAuth(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username.toLowerCase()) {
      const userData = {
        username: data[i][0],
        password: data[i][1],
        security: {
          pet: data[i][2],
          hobby: data[i][3],
          food: data[i][4]
        },
        registeredAt: data[i][5],
        lastLogin: data[i][6],
        isActive: data[i][7]
      };
      
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, userData: userData }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, message: "User not found" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleCleanupLunasOtomatis(username) {
  const result = cleanupLunasData(username);
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, result: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== HELPER FUNCTIONS ====================

function getAllBonsByUser(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_BONS);
  const data = sheet.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username && data[i][7] !== false) {
      result.push({
        uniqueId: data[i][0],
        username: data[i][1],
        namaPelanggan: data[i][2],
        total: data[i][3],
        items: JSON.parse(data[i][4] || "[]"),
        waktu: data[i][5],
        lastModified: data[i][6],
        isActive: data[i][7]
      });
    }
  }
  return result;
}

function getAllPaymentsByUser(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PAYMENTS);
  const data = sheet.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      result.push({
        uniqueId: data[i][0],
        username: data[i][1],
        namaPelanggan: data[i][2],
        jumlah: data[i][3],
        waktu: data[i][4],
        lastModified: data[i][5]
      });
    }
  }
  return result;
}

function filterActiveBonsV36(bons, payments) {
  // Group payments by customer
  const paymentSum = {};
  payments.forEach(p => {
    const key = p.namaPelanggan.toLowerCase();
    paymentSum[key] = (paymentSum[key] || 0) + p.jumlah;
  });
  
  // Calculate outstanding per customer
  const outstanding = {};
  bons.forEach(bon => {
    const key = bon.namaPelanggan.toLowerCase();
    outstanding[key] = (outstanding[key] || 0) + bon.total;
  });
  
  const remaining = {};
  for (const key in outstanding) {
    const paid = paymentSum[key] || 0;
    const sisa = outstanding[key] - paid;
    remaining[key] = sisa;
  }
  
  // Filter bons: keep if customer still has outstanding debt OR has change (negative)
  const activeBons = [];
  bons.forEach(bon => {
    const key = bon.namaPelanggan.toLowerCase();
    const sisa = remaining[key];
    
    // Keep bon if:
    // 1. Sisa utang > 0 (still owe)
    // 2. Sisa utang < 0 (overpaid / has change)
    // 3. Sisa utang === 0 AND there is change (negative payment sum)
    const totalPaid = paymentSum[key] || 0;
    const totalOwed = outstanding[key] || 0;
    const hasChange = totalPaid > totalOwed;
    
    if (sisa > 0 || hasChange) {
      activeBons.push(bon);
    }
    // If sisa === 0 and no change -> customer fully paid, remove all bons
  });
  
  return activeBons;
}

function cleanupLunasData(username) {
  const bons = getAllBonsByUser(username);
  const payments = getAllPaymentsByUser(username);
  
  // Calculate per customer
  const paymentSum = {};
  payments.forEach(p => {
    const key = p.namaPelanggan.toLowerCase();
    paymentSum[key] = (paymentSum[key] || 0) + p.jumlah;
  });
  
  const outstanding = {};
  bons.forEach(bon => {
    const key = bon.namaPelanggan.toLowerCase();
    outstanding[key] = (outstanding[key] || 0) + bon.total;
  });
  
  const customersToDelete = [];
  
  for (const key in outstanding) {
    const totalOwed = outstanding[key];
    const totalPaid = paymentSum[key] || 0;
    const sisa = totalOwed - totalPaid;
    
    // If sisa === 0 AND no overpayment (no change), delete all customer data
    if (sisa === 0 && totalPaid <= totalOwed) {
      customersToDelete.push(key);
    }
  }
  
  // Also check customers who have payment but no bons (should be cleaned too)
  for (const key in paymentSum) {
    if (!outstanding[key]) {
      const totalPaid = paymentSum[key];
      if (totalPaid > 0) {
        // Has payments but no bons -> clean up payments
        customersToDelete.push(key);
      }
    }
  }
  
  // Delete bons for customers to delete
  const sheetBons = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_BONS);
  const bonsData = sheetBons.getDataRange().getValues();
  
  for (let i = bonsData.length - 1; i >= 1; i--) {
    const bonUsername = bonsData[i][1];
    const namaPelanggan = bonsData[i][2];
    if (bonUsername === username && customersToDelete.includes(namaPelanggan.toLowerCase())) {
      sheetBons.deleteRow(i + 1);
    }
  }
  
  // Delete payments for customers to delete
  const sheetPayments = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PAYMENTS);
  const paymentsData = sheetPayments.getDataRange().getValues();
  
  for (let i = paymentsData.length - 1; i >= 1; i--) {
    const paymentUsername = paymentsData[i][1];
    const namaPelanggan = paymentsData[i][2];
    if (paymentUsername === username && customersToDelete.includes(namaPelanggan.toLowerCase())) {
      sheetPayments.deleteRow(i + 1);
    }
  }
  
  addSyncLog(username, "cleanupLunas", `Deleted ${customersToDelete.length} customers (LUNAS tanpa kembalian)`);
  
  return {
    deletedCustomers: customersToDelete,
    count: customersToDelete.length,
    timestamp: new Date().toISOString()
  };
}

function saveAllBonsToSheet(username, bons) {
  // Clear existing bons for this user
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_BONS);
  const data = sheet.getDataRange().getValues();
  
  // Mark all existing as inactive first
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      sheet.getRange(i + 1, 8).setValue(false);
    }
  }
  
  // Write new/updated bons
  bons.forEach(bon => {
    const existingRow = findRowByUniqueId(sheet, bon.uniqueId);
    if (existingRow) {
      const rowNum = existingRow;
      sheet.getRange(rowNum, 3).setValue(bon.namaPelanggan);
      sheet.getRange(rowNum, 4).setValue(bon.total);
      sheet.getRange(rowNum, 5).setValue(JSON.stringify(bon.items));
      sheet.getRange(rowNum, 6).setValue(bon.waktu);
      sheet.getRange(rowNum, 7).setValue(bon.lastModified);
      sheet.getRange(rowNum, 8).setValue(true);
    } else {
      sheet.appendRow([
        bon.uniqueId,
        username,
        bon.namaPelanggan,
        bon.total,
        JSON.stringify(bon.items),
        bon.waktu,
        bon.lastModified,
        true
      ]);
    }
  });
}

function saveAllPaymentsToSheet(username, payments) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PAYMENTS);
  const data = sheet.getDataRange().getValues();
  
  // Delete all existing payments for this user
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === username) {
      sheet.deleteRow(i + 1);
    }
  }
  
  // Write all payments
  payments.forEach(payment => {
    sheet.appendRow([
      payment.uniqueId,
      username,
      payment.namaPelanggan,
      payment.jumlah,
      payment.waktu,
      payment.lastModified
    ]);
  });
}

function findRowByUniqueId(sheet, uniqueId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === uniqueId) {
      return i + 1;
    }
  }
  return null;
}

function findUserRow(sheet, username) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      return i + 1;
    }
  }
  return null;
}

function mergeBonDataV36(existingBons, remoteBons) {
  const bonMap = new Map();
  
  existingBons.forEach(bon => {
    bonMap.set(bon.uniqueId, bon);
  });
  
  remoteBons.forEach(bon => {
    const existing = bonMap.get(bon.uniqueId);
    if (!existing) {
      bonMap.set(bon.uniqueId, bon);
    } else {
      const existingTime = new Date(existing.lastModified).getTime();
      const remoteTime = new Date(bon.lastModified).getTime();
      if (remoteTime > existingTime) {
        bonMap.set(bon.uniqueId, bon);
      }
    }
  });
  
  return Array.from(bonMap.values());
}

function mergePaymentDataV36(existingPayments, remotePayments) {
  const paymentMap = new Map();
  
  existingPayments.forEach(payment => {
    paymentMap.set(payment.uniqueId, payment);
  });
  
  remotePayments.forEach(payment => {
    const existing = paymentMap.get(payment.uniqueId);
    if (!existing) {
      paymentMap.set(payment.uniqueId, payment);
    } else {
      const existingTime = new Date(existing.lastModified).getTime();
      const remoteTime = new Date(payment.lastModified).getTime();
      if (remoteTime > existingTime) {
        paymentMap.set(payment.uniqueId, payment);
      }
    }
  });
  
  return Array.from(paymentMap.values());
}

function addSyncLog(username, action, details) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_SYNC_LOG);
  sheet.appendRow([
    new Date().toISOString(),
    username || "system",
    action,
    details
  ]);
}
