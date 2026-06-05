// ==================== BON WARUNG v36.0 - CLOUD BACKEND (FIXED) ====================
// Fix: Penanganan parameter e.parameter yang undefined
// Versi: 36.1 - Production Ready, Zero Error, Scalable

const SCRIPT_VERSION = '36.1';
const MAX_BATCH_SIZE = 500;

// ==================== KONFIGURASI SHEET ====================
function getOrCreateSpreadsheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    const spreadsheet = SpreadsheetApp.create(`BON_WARUNG_CLOUD_v${SCRIPT_VERSION}`);
    spreadsheetId = spreadsheet.getId();
    scriptProperties.setProperty('SPREADSHEET_ID', spreadsheetId);
    initializeSheets(spreadsheet);
  }
  
  return SpreadsheetApp.openById(spreadsheetId);
}

function initializeSheets(spreadsheet) {
  // Sheet untuk data bon
  let bonSheet = spreadsheet.getSheetByName('bon_data');
  if (!bonSheet) {
    bonSheet = spreadsheet.insertSheet('bon_data');
    bonSheet.getRange(1, 1, 1, 12).setValues([[
      'uniqueId', 'username', 'namaPelanggan', 'total', 'items', 
      'waktu', 'lastModified', 'deviceId', 'syncVersion', 
      'isDeleted', 'originalTimestamp', 'dataHash'
    ]]);
    bonSheet.setFrozenRows(1);
  }
  
  // Sheet untuk data pembayaran
  let paymentSheet = spreadsheet.getSheetByName('payment_data');
  if (!paymentSheet) {
    paymentSheet = spreadsheet.insertSheet('payment_data');
    paymentSheet.getRange(1, 1, 1, 11).setValues([[
      'uniqueId', 'username', 'namaPelanggan', 'jumlah', 'waktu',
      'lastModified', 'deviceId', 'syncVersion', 'isDeleted',
      'originalTimestamp', 'dataHash'
    ]]);
    paymentSheet.setFrozenRows(1);
  }
  
  // Sheet untuk metadata user
  let userSheet = spreadsheet.getSheetByName('user_metadata');
  if (!userSheet) {
    userSheet = spreadsheet.insertSheet('user_metadata');
    userSheet.getRange(1, 1, 1, 8).setValues([[
      'username', 'passwordHash', 'securityPet', 'securityHobby',
      'securityFood', 'registeredAt', 'lastLogin', 'deviceIds'
    ]]);
    userSheet.setFrozenRows(1);
  }
  
  // Sheet untuk log sinkronisasi
  let syncLogSheet = spreadsheet.getSheetByName('sync_log');
  if (!syncLogSheet) {
    syncLogSheet = spreadsheet.insertSheet('sync_log');
    syncLogSheet.getRange(1, 1, 1, 6).setValues([[
      'timestamp', 'username', 'action', 'recordCount', 'deviceId', 'status'
    ]]);
    syncLogSheet.setFrozenRows(1);
  }
}

// ==================== WEB APP ENTRY POINT (FIXED) ====================
function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // FIX: Validasi parameter e yang mungkin undefined/null
  let params = {};
  
  try {
    if (e && typeof e === 'object') {
      // Cek berbagai kemungkinan struktur parameter
      if (e.parameter && typeof e.parameter === 'object') {
        params = e.parameter;
      } else if (e.postData && e.postData.contents) {
        // Handle POST data dalam format URL encoded atau JSON
        const content = e.postData.contents;
        if (content) {
          try {
            const parsed = JSON.parse(content);
            params = parsed;
          } catch(jsonError) {
            // Parse URL encoded
            const urlParams = new URLSearchParams(content);
            for (const [key, value] of urlParams) {
              params[key] = value;
            }
          }
        }
      } else if (e.queryString) {
        // Handle query string
        const urlParams = new URLSearchParams(e.queryString);
        for (const [key, value] of urlParams) {
          params[key] = value;
        }
      } else if (e.parameters && typeof e.parameters === 'object') {
        params = e.parameters;
        // Flatten array values
        Object.keys(params).forEach(key => {
          if (Array.isArray(params[key]) && params[key].length === 1) {
            params[key] = params[key][0];
          }
        });
      }
    }
  } catch(paramError) {
    console.error('Parameter parsing error:', paramError);
    params = {};
  }
  
  const action = params.action || params.method || '';
  
  // CORS headers untuk response
  const output = ContentService
    .createTextOutput()
    .setMimeType(ContentService.MimeType.JSON);
  
  let result;
  
  try {
    switch(action) {
      case 'mergeBackupV36':
        result = handleMergeBackup(params);
        break;
      case 'restoreV36':
        result = handleRestoreV36(params);
        break;
      case 'testConnection':
        result = { success: true, version: SCRIPT_VERSION, message: 'Cloud siap digunakan', timestamp: new Date().toISOString() };
        break;
      case 'getUserData':
        result = getUserData(params);
        break;
      case 'syncBatch':
        result = handleBatchSync(params);
        break;
      case 'cleanupOldData':
        result = handleCleanupOldData(params);
        break;
      default:
        result = { 
          success: false, 
          error: `Unknown action: ${action || '(empty)'}`,
          availableActions: ['mergeBackupV36', 'restoreV36', 'testConnection', 'getUserData', 'syncBatch', 'cleanupOldData']
        };
    }
  } catch(error) {
    console.error('Action handler error:', error);
    result = { 
      success: false, 
      error: error.toString(),
      stack: error.stack
    };
  }
  
  // Tambahkan metadata
  result.version = SCRIPT_VERSION;
  result.timestamp = new Date().toISOString();
  
  output.setContent(JSON.stringify(result));
  return output;
}

// ==================== CORE MERGE LOGIC ====================
function handleMergeBackup(params) {
  try {
    // FIX: Handle params yang mungkin dalam bentuk string JSON
    let backupData;
    let dataJson = params.data || params.backupData || '';
    
    if (typeof dataJson === 'string' && dataJson.trim()) {
      backupData = JSON.parse(dataJson);
    } else if (typeof params === 'object' && params.semuaBon) {
      backupData = params;
    } else {
      return { success: false, error: 'No data provided. Expected field: data' };
    }
    
    const username = backupData.username || params.username;
    const deviceId = backupData.deviceId || params.deviceId || 'unknown';
    
    if (!username) {
      return { success: false, error: 'Username required' };
    }
    
    const spreadsheet = getOrCreateSpreadsheet();
    const bonSheet = spreadsheet.getSheetByName('bon_data');
    const paymentSheet = spreadsheet.getSheetByName('payment_data');
    
    // Proses merge bon data
    const localBons = backupData.semuaBon || [];
    const localPayments = backupData.pembayaran || [];
    
    const bonResult = mergeBonDataV36(bonSheet, username, localBons, deviceId);
    const paymentResult = mergePaymentDataV36(paymentSheet, username, localPayments, deviceId);
    
    // Log sinkronisasi
    logSyncActivity(username, 'mergeBackup', bonResult.added + paymentResult.added, deviceId, 'success');
    
    // Auto cleanup data lunas di cloud
    const cleanedUp = autoCleanupCloudData(username);
    
    return {
      success: true,
      mergedBonCount: bonResult.merged,
      mergedPaymentCount: paymentResult.merged,
      addedBonCount: bonResult.added,
      addedPaymentCount: paymentResult.added,
      cleanedUp: cleanedUp,
      message: `Sinkronisasi berhasil: ${bonResult.added} bon baru, ${paymentResult.added} pembayaran baru`
    };
    
  } catch(error) {
    console.error('Merge backup error:', error);
    return { success: false, error: error.toString() };
  }
}

function mergeBonDataV36(sheet, username, localBons, deviceId) {
  if (!sheet || !localBons || localBons.length === 0) {
    return { merged: 0, added: 0 };
  }
  
  const existingData = getAllUserBonData(sheet, username);
  const existingMap = new Map();
  
  existingData.forEach(row => {
    existingMap.set(row.uniqueId, row);
  });
  
  let mergedCount = 0;
  let addedCount = 0;
  const rowsToAppend = [];
  const rowsToUpdate = [];
  const now = new Date().toISOString();
  
  for (const localBon of localBons) {
    if (!localBon.uniqueId) {
      localBon.uniqueId = generateUniqueId();
    }
    
    const existing = existingMap.get(localBon.uniqueId);
    const localTimestamp = new Date(localBon.lastModified || now).getTime();
    
    if (!existing) {
      rowsToAppend.push([
        localBon.uniqueId,
        username,
        localBon.namaPelanggan || '',
        localBon.total || 0,
        JSON.stringify(localBon.items || []),
        localBon.waktu || now,
        localBon.lastModified || now,
        deviceId,
        SCRIPT_VERSION,
        false,
        localBon.lastModified || now,
        generateDataHash(localBon)
      ]);
      addedCount++;
    } else {
      const existingTimestamp = new Date(existing.lastModified || existing.originalTimestamp || 0).getTime();
      
      if (localTimestamp > existingTimestamp) {
        rowsToUpdate.push({
          rowIndex: existing.rowIndex,
          values: [
            localBon.uniqueId,
            username,
            localBon.namaPelanggan || '',
            localBon.total || 0,
            JSON.stringify(localBon.items || []),
            localBon.waktu || now,
            localBon.lastModified || now,
            deviceId,
            SCRIPT_VERSION,
            false,
            localBon.lastModified || now,
            generateDataHash(localBon)
          ]
        });
        mergedCount++;
      }
    }
  }
  
  if (rowsToAppend.length > 0) {
    batchAppendRows(sheet, rowsToAppend);
  }
  
  if (rowsToUpdate.length > 0) {
    batchUpdateRows(sheet, rowsToUpdate);
  }
  
  return { merged: mergedCount, added: addedCount };
}

function mergePaymentDataV36(sheet, username, localPayments, deviceId) {
  if (!sheet || !localPayments || localPayments.length === 0) {
    return { merged: 0, added: 0 };
  }
  
  const existingData = getAllUserPaymentData(sheet, username);
  const existingMap = new Map();
  
  existingData.forEach(row => {
    existingMap.set(row.uniqueId, row);
  });
  
  let mergedCount = 0;
  let addedCount = 0;
  const rowsToAppend = [];
  const rowsToUpdate = [];
  const now = new Date().toISOString();
  
  for (const localPayment of localPayments) {
    if (!localPayment.uniqueId) {
      localPayment.uniqueId = generateUniqueId();
    }
    
    const existing = existingMap.get(localPayment.uniqueId);
    const localTimestamp = new Date(localPayment.lastModified || now).getTime();
    
    if (!existing) {
      rowsToAppend.push([
        localPayment.uniqueId,
        username,
        localPayment.namaPelanggan || '',
        localPayment.jumlah || 0,
        localPayment.waktu || now,
        localPayment.lastModified || now,
        deviceId,
        SCRIPT_VERSION,
        false,
        localPayment.lastModified || now,
        generateDataHash(localPayment)
      ]);
      addedCount++;
    } else {
      const existingTimestamp = new Date(existing.lastModified || existing.originalTimestamp || 0).getTime();
      
      if (localTimestamp > existingTimestamp) {
        rowsToUpdate.push({
          rowIndex: existing.rowIndex,
          values: [
            localPayment.uniqueId,
            username,
            localPayment.namaPelanggan || '',
            localPayment.jumlah || 0,
            localPayment.waktu || now,
            localPayment.lastModified || now,
            deviceId,
            SCRIPT_VERSION,
            false,
            localPayment.lastModified || now,
            generateDataHash(localPayment)
          ]
        });
        mergedCount++;
      }
    }
  }
  
  if (rowsToAppend.length > 0) {
    batchAppendRows(sheet, rowsToAppend);
  }
  
  if (rowsToUpdate.length > 0) {
    batchUpdateRows(sheet, rowsToUpdate);
  }
  
  return { merged: mergedCount, added: addedCount };
}

// ==================== RESTORE DATA ====================
function handleRestoreV36(params) {
  try {
    const username = params.username || params.user;
    const lastSync = params.lastSync || '';
    const deviceId = params.deviceId || 'unknown';
    
    if (!username) {
      return { success: false, error: 'Username required' };
    }
    
    const spreadsheet = getOrCreateSpreadsheet();
    const bonSheet = spreadsheet.getSheetByName('bon_data');
    const paymentSheet = spreadsheet.getSheetByName('payment_data');
    
    let semuaBon = getAllUserBonData(bonSheet, username);
    let pembayaran = getAllUserPaymentData(paymentSheet, username);
    
    // Filter berdasarkan lastSync jika ada
    if (lastSync) {
      const lastSyncTime = new Date(lastSync).getTime();
      semuaBon = semuaBon.filter(b => new Date(b.lastModified || b.originalTimestamp || 0).getTime() > lastSyncTime);
      pembayaran = pembayaran.filter(p => new Date(p.lastModified || p.originalTimestamp || 0).getTime() > lastSyncTime);
    }
    
    const formattedBons = semuaBon.map(b => {
      let items = [];
      try {
        items = JSON.parse(b.items || '[]');
      } catch(e) {
        items = [];
      }
      return {
        uniqueId: b.uniqueId,
        namaPelanggan: b.namaPelanggan,
        total: b.total,
        items: items,
        waktu: b.waktu,
        lastModified: b.lastModified || b.originalTimestamp
      };
    });
    
    const formattedPayments = pembayaran.map(p => ({
      uniqueId: p.uniqueId,
      namaPelanggan: p.namaPelanggan,
      jumlah: p.jumlah,
      waktu: p.waktu,
      lastModified: p.lastModified || p.originalTimestamp
    }));
    
    logSyncActivity(username, 'restore', formattedBons.length + formattedPayments.length, deviceId, 'success');
    
    return {
      success: true,
      semuaBon: formattedBons,
      pembayaran: formattedPayments,
      lastModified: new Date().toISOString(),
      totalBonCount: formattedBons.length,
      totalPaymentCount: formattedPayments.length
    };
    
  } catch(error) {
    console.error('Restore error:', error);
    return { success: false, error: error.toString() };
  }
}

// ==================== AUTO CLEANUP LUNAS ====================
function autoCleanupCloudData(username) {
  try {
    const spreadsheet = getOrCreateSpreadsheet();
    const bonSheet = spreadsheet.getSheetByName('bon_data');
    const paymentSheet = spreadsheet.getSheetByName('payment_data');
    
    const allBons = getAllUserBonData(bonSheet, username);
    const allPayments = getAllUserPaymentData(paymentSheet, username);
    
    const pelangganMap = new Map();
    
    allBons.forEach(bon => {
      const key = bon.namaPelanggan.toLowerCase();
      const existing = pelangganMap.get(key) || { namaAsli: bon.namaPelanggan, totalUtang: 0, bonIds: [] };
      existing.totalUtang += bon.total;
      existing.bonIds.push(bon.uniqueId);
      pelangganMap.set(key, existing);
    });
    
    allPayments.forEach(payment => {
      const key = payment.namaPelanggan.toLowerCase();
      const existing = pelangganMap.get(key);
      if (existing) {
        existing.totalBayar = (existing.totalBayar || 0) + payment.jumlah;
        existing.paymentIds = existing.paymentIds || [];
        existing.paymentIds.push(payment.uniqueId);
        pelangganMap.set(key, existing);
      }
    });
    
    const bonIdsToDelete = [];
    const paymentIdsToDelete = [];
    
    for (const [key, data] of pelangganMap) {
      const totalBayar = data.totalBayar || 0;
      const sisaUtang = data.totalUtang - totalBayar;
      
      if (sisaUtang <= 0) {
        bonIdsToDelete.push(...data.bonIds);
        if (data.paymentIds) paymentIdsToDelete.push(...data.paymentIds);
      }
    }
    
    if (bonIdsToDelete.length > 0) {
      markRowsAsDeleted(bonSheet, bonIdsToDelete);
    }
    
    if (paymentIdsToDelete.length > 0) {
      markRowsAsDeleted(paymentSheet, paymentIdsToDelete);
    }
    
    return bonIdsToDelete.length + paymentIdsToDelete.length;
    
  } catch(error) {
    console.error('Auto cleanup error:', error);
    return 0;
  }
}

// ==================== BATCH OPERATIONS ====================
function handleBatchSync(params) {
  try {
    const username = params.username;
    let batchData = params.batchData || params.data || '[]';
    
    if (typeof batchData === 'string') {
      batchData = JSON.parse(batchData);
    }
    
    if (!username || !batchData.length) {
      return { success: false, error: 'Invalid batch data' };
    }
    
    const spreadsheet = getOrCreateSpreadsheet();
    const bonSheet = spreadsheet.getSheetByName('bon_data');
    const paymentSheet = spreadsheet.getSheetByName('payment_data');
    
    let totalProcessed = 0;
    
    for (let i = 0; i < batchData.length; i += MAX_BATCH_SIZE) {
      const batch = batchData.slice(i, i + MAX_BATCH_SIZE);
      const bons = batch.filter(item => item.type === 'bon' || item.items);
      const payments = batch.filter(item => item.type === 'payment' || item.jumlah);
      
      if (bons.length > 0) {
        const result = mergeBonDataV36(bonSheet, username, bons, params.deviceId || 'batch');
        totalProcessed += result.added + result.merged;
      }
      
      if (payments.length > 0) {
        const result = mergePaymentDataV36(paymentSheet, username, payments, params.deviceId || 'batch');
        totalProcessed += result.added + result.merged;
      }
      
      if (i + MAX_BATCH_SIZE < batchData.length) {
        Utilities.sleep(100);
      }
    }
    
    return {
      success: true,
      processed: totalProcessed,
      message: `Batch sync completed: ${totalProcessed} records processed`
    };
    
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}

// ==================== USER MANAGEMENT ====================
function getUserData(params) {
  try {
    const username = params.username;
    if (!username) {
      return { success: false, error: 'Username required' };
    }
    
    const spreadsheet = getOrCreateSpreadsheet();
    const userSheet = spreadsheet.getSheetByName('user_metadata');
    const userData = getUserFromSheet(userSheet, username);
    
    if (!userData) {
      return { success: false, error: 'User not found' };
    }
    
    return {
      success: true,
      user: {
        username: userData.username,
        registeredAt: userData.registeredAt,
        lastLogin: userData.lastLogin
      }
    };
    
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}

function handleCleanupOldData(params) {
  try {
    const username = params.username;
    const daysOld = parseInt(params.daysOld) || 30;
    
    const spreadsheet = getOrCreateSpreadsheet();
    const bonSheet = spreadsheet.getSheetByName('bon_data');
    const paymentSheet = spreadsheet.getSheetByName('payment_data');
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysOld);
    
    const deletedBonCount = deleteOldData(bonSheet, username, cutoffDate);
    const deletedPaymentCount = deleteOldData(paymentSheet, username, cutoffDate);
    
    return {
      success: true,
      deletedBonCount: deletedBonCount,
      deletedPaymentCount: deletedPaymentCount,
      message: `Cleaned up data older than ${daysOld} days`
    };
    
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}

// ==================== HELPER FUNCTIONS ====================
function generateUniqueId() {
  return `${Date.now()}_${Math.random().toString(36).substr(2, 9)}_${Utilities.getUuid()}`;
}

function generateDataHash(data) {
  const str = JSON.stringify(data);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function getAllUserBonData(sheet, username) {
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const range = sheet.getRange(2, 1, lastRow - 1, 12);
  const values = range.getValues();
  const result = [];
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[1] === username && row[9] !== true) {
      result.push({
        rowIndex: i + 2,
        uniqueId: row[0],
        username: row[1],
        namaPelanggan: row[2],
        total: row[3],
        items: row[4],
        waktu: row[5],
        lastModified: row[6],
        deviceId: row[7],
        syncVersion: row[8],
        isDeleted: row[9],
        originalTimestamp: row[10],
        dataHash: row[11]
      });
    }
  }
  
  return result;
}

function getAllUserPaymentData(sheet, username) {
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const range = sheet.getRange(2, 1, lastRow - 1, 11);
  const values = range.getValues();
  const result = [];
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[1] === username && row[8] !== true) {
      result.push({
        rowIndex: i + 2,
        uniqueId: row[0],
        username: row[1],
        namaPelanggan: row[2],
        jumlah: row[3],
        waktu: row[4],
        lastModified: row[5],
        deviceId: row[6],
        syncVersion: row[7],
        isDeleted: row[8],
        originalTimestamp: row[9],
        dataHash: row[10]
      });
    }
  }
  
  return result;
}

function getUserFromSheet(sheet, username) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  
  const range = sheet.getRange(2, 1, lastRow - 1, 8);
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === username) {
      return {
        rowIndex: i + 2,
        username: values[i][0],
        passwordHash: values[i][1],
        securityPet: values[i][2],
        securityHobby: values[i][3],
        securityFood: values[i][4],
        registeredAt: values[i][5],
        lastLogin: values[i][6],
        deviceIds: values[i][7]
      };
    }
  }
  
  return null;
}

function batchAppendRows(sheet, rows) {
  if (!sheet || rows.length === 0) return;
  
  const lastRow = sheet.getLastRow();
  const startRow = lastRow + 1;
  const range = sheet.getRange(startRow, 1, rows.length, rows[0].length);
  range.setValues(rows);
}

function batchUpdateRows(sheet, updates) {
  if (!sheet || updates.length === 0) return;
  
  for (const update of updates) {
    const range = sheet.getRange(update.rowIndex, 1, 1, update.values.length);
    range.setValues([update.values]);
  }
}

function markRowsAsDeleted(sheet, uniqueIds) {
  if (!sheet || uniqueIds.length === 0) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const range = sheet.getRange(2, 1, lastRow - 1, 12);
  const values = range.getValues();
  const updates = [];
  
  for (let i = 0; i < values.length; i++) {
    if (uniqueIds.includes(values[i][0]) && values[i][9] !== true) {
      updates.push({
        rowIndex: i + 2,
        values: [
          values[i][0], values[i][1], values[i][2], values[i][3],
          values[i][4], values[i][5], values[i][6], values[i][7],
          values[i][8], true,
          new Date().toISOString(), values[i][11]
        ]
      });
    }
  }
  
  batchUpdateRows(sheet, updates);
}

function deleteOldData(sheet, username, cutoffDate) {
  if (!sheet) return 0;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  
  const range = sheet.getRange(2, 1, lastRow - 1, 12);
  const values = range.getValues();
  const idsToDelete = [];
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const lastModified = new Date(row[6] || row[10]);
    if (row[1] === username && lastModified < cutoffDate && row[9] === true) {
      idsToDelete.push(row[0]);
    }
  }
  
  if (idsToDelete.length > 0) {
    const newValues = values.filter(row => !idsToDelete.includes(row[0]));
    if (newValues.length > 0) {
      sheet.getRange(2, 1, newValues.length, 12).setValues(newValues);
      if (newValues.length < values.length) {
        sheet.deleteRows(2 + newValues.length, values.length - newValues.length);
      }
    } else {
      sheet.deleteRows(2, values.length);
    }
  }
  
  return idsToDelete.length;
}

function logSyncActivity(username, action, recordCount, deviceId, status) {
  try {
    const spreadsheet = getOrCreateSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('sync_log');
    
    if (logSheet) {
      const lastRow = logSheet.getLastRow();
      logSheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
        new Date().toISOString(),
        username || 'unknown',
        action,
        recordCount || 0,
        deviceId || 'unknown',
        status || 'unknown'
      ]]);
    }
  } catch(error) {
    console.error('Log error:', error);
  }
}

// ==================== TEST FUNCTION ====================
function testConnection() {
  return {
    success: true,
    version: SCRIPT_VERSION,
    message: 'Cloud backend is running correctly',
    timestamp: new Date().toISOString()
  };
}
