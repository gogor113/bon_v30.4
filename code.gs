// ==================== BON WARUNG v35.0 - COMPLETE CLOUD BACKEND (GS) ====================
// Versi: 2.0.0 (Compatible with BON_V35.html)
// Nama File: Code.gs
// Fungsi: CRUD lengkap untuk data utang warung - CREATE, READ, UPDATE, DELETE, MERGE, DRAW
// Fitur Lengkap:
// 1. Multi-user authentication & data isolation
// 2. CRUD Operations untuk Bon (Create, Read, Update, Delete, Restore)
// 3. CRUD Operations untuk Payment (Create, Read, Update, Delete, Restore)
// 4. Conflict resolution berdasarkan timestamp terbaru
// 5. Merge logic untuk cross-device sync
// 6. Search & Filter data
// 7. Backup & Restore
// 8. Data integrity dengan unique ID per record
// 9. Real-time drawing data support
// 10. Audit log untuk semua perubahan
// ==============================================================================

// ==================== KONFIGURASI SPREADSHEET ====================
const SPREADSHEET_ID = '1CUUbyDsEgFFUwHmuSTL73gOJYXXtN2cADDYygpc4U5g'; // <-- HANYA ID, tanpa parameter tambahan!

// Nama-nama sheet yang akan dibuat secara otomatis
const SHEET_USERS = 'Users';
const SHEET_BONS = 'Bons';
const SHEET_PAYMENTS = 'Payments';
const SHEET_DRAWINGS = 'Drawings';
const SHEET_AUDIT_LOG = 'AuditLog';
const SHEET_BACKUP = 'Backup';

// ==================== FUNGSI UTAMA DO POST ====================
function doPost(e) {
  const response = {
    success: false,
    message: '',
    data: null,
    timestamp: new Date().toISOString()
  };
  
  try {
    if (!e || !e.parameter) {
      response.message = 'No parameters provided';
      return sendJsonResponse(response, 400);
    }
    
    const action = e.parameter.action;
    const dataParam = e.parameter.data;
    
    Logger.log(`[${new Date().toISOString()}] Action: ${action}`);
    
    switch(action) {
      // ========== TEST CONNECTION ==========
      case 'testConnection':
        response.success = true;
        response.message = 'Cloud Backup v35.0 siap dan berfungsi!';
        response.data = { 
          serverTime: new Date().toISOString(),
          version: '35.0',
          features: ['CRUD', 'MERGE', 'DELETE', 'RESTORE', 'DRAW', 'SEARCH']
        };
        break;
      
      // ========== USER AUTHENTICATION ==========
      case 'getUserAuth':
        const username = e.parameter.username;
        if (!username) {
          response.message = 'Username required';
          return sendJsonResponse(response, 400);
        }
        const userData = getUserFromSheet(username);
        if (userData) {
          response.success = true;
          response.userData = userData;
        } else {
          response.success = false;
          response.message = 'User not found';
        }
        break;
        
      case 'syncUserAuth':
        const syncUserData = JSON.parse(dataParam || '{}');
        const targetUsername = syncUserData.username;
        const userInfo = syncUserData.userData;
        
        if (!targetUsername || !userInfo) {
          response.message = 'Username and userData required';
          return sendJsonResponse(response, 400);
        }
        
        const saveResult = saveUserToSheet(targetUsername, userInfo);
        if (saveResult) {
          response.success = true;
          response.message = 'User synced successfully';
          auditLog(targetUsername, 'USER_SYNC', 'User authentication synced');
        } else {
          response.success = false;
          response.message = 'Failed to save user data';
        }
        break;
      
      // ========== BON CRUD OPERATIONS ==========
      case 'createBon':
        const createBonUsername = e.parameter.username;
        const createBonData = JSON.parse(dataParam || '{}');
        
        if (!createBonUsername || !createBonData) {
          response.message = 'Username and bonData required';
          return sendJsonResponse(response, 400);
        }
        
        const newBon = {
          ...createBonData,
          uniqueId: generateUniqueId(),
          createdAt: new Date().toISOString(),
          lastModified: new Date().toISOString(),
          isDeleted: false,
          deletedAt: null
        };
        
        const createdBon = saveSingleBonToSheet(createBonUsername, newBon);
        if (createdBon) {
          response.success = true;
          response.message = 'Bon created successfully';
          response.data = newBon;
          auditLog(createBonUsername, 'BON_CREATE', `Created bon: ${newBon.uniqueId}`);
        } else {
          response.success = false;
          response.message = 'Failed to create bon';
        }
        break;
        
      case 'readBons':
        const readBonUsername = e.parameter.username;
        const includeDeleted = e.parameter.includeDeleted === 'true';
        
        if (!readBonUsername) {
          response.message = 'Username required';
          return sendJsonResponse(response, 400);
        }
        
        let bons = getBonsFromSheet(readBonUsername);
        if (!includeDeleted) {
          bons = bons.filter(bon => !bon.isDeleted);
        }
        
        response.success = true;
        response.data = bons;
        response.count = bons.length;
        break;
        
      case 'updateBon':
        const updateBonUsername = e.parameter.username;
        const updateBonData = JSON.parse(dataParam || '{}');
        const bonId = updateBonData.uniqueId;
        
        if (!updateBonUsername || !bonId) {
          response.message = 'Username and bon uniqueId required';
          return sendJsonResponse(response, 400);
        }
        
        const updatedBon = updateBonInSheet(updateBonUsername, bonId, updateBonData);
        if (updatedBon) {
          response.success = true;
          response.message = 'Bon updated successfully';
          response.data = updatedBon;
          auditLog(updateBonUsername, 'BON_UPDATE', `Updated bon: ${bonId}`);
        } else {
          response.success = false;
          response.message = 'Failed to update bon or bon not found';
        }
        break;
        
      case 'deleteBon':
        const deleteBonUsername = e.parameter.username;
        const deleteBonId = e.parameter.bonId;
        const permanentDelete = e.parameter.permanent === 'true';
        
        if (!deleteBonUsername || !deleteBonId) {
          response.message = 'Username and bonId required';
          return sendJsonResponse(response, 400);
        }
        
        let deleteResult;
        if (permanentDelete) {
          deleteResult = permanentlyDeleteBon(deleteBonUsername, deleteBonId);
          response.message = 'Bon permanently deleted';
        } else {
          deleteResult = softDeleteBon(deleteBonUsername, deleteBonId);
          response.message = 'Bon soft deleted (moved to trash)';
        }
        
        if (deleteResult) {
          response.success = true;
          auditLog(deleteBonUsername, 'BON_DELETE', `Deleted bon: ${deleteBonId}, permanent: ${permanentDelete}`);
        } else {
          response.success = false;
          response.message = 'Failed to delete bon';
        }
        break;
        
      case 'restoreBon':
        const restoreBonUsername = e.parameter.username;
        const restoreBonId = e.parameter.bonId;
        
        if (!restoreBonUsername || !restoreBonId) {
          response.message = 'Username and bonId required';
          return sendJsonResponse(response, 400);
        }
        
        const restoredBon = restoreBonFromSheet(restoreBonUsername, restoreBonId);
        if (restoredBon) {
          response.success = true;
          response.message = 'Bon restored successfully';
          response.data = restoredBon;
          auditLog(restoreBonUsername, 'BON_RESTORE', `Restored bon: ${restoreBonId}`);
        } else {
          response.success = false;
          response.message = 'Failed to restore bon';
        }
        break;
      
      // ========== PAYMENT CRUD OPERATIONS ==========
      case 'createPayment':
        const createPaymentUsername = e.parameter.username;
        const createPaymentData = JSON.parse(dataParam || '{}');
        
        if (!createPaymentUsername || !createPaymentData) {
          response.message = 'Username and paymentData required';
          return sendJsonResponse(response, 400);
        }
        
        const newPayment = {
          ...createPaymentData,
          uniqueId: generateUniqueId(),
          createdAt: new Date().toISOString(),
          lastModified: new Date().toISOString(),
          isDeleted: false,
          deletedAt: null
        };
        
        const createdPayment = saveSinglePaymentToSheet(createPaymentUsername, newPayment);
        if (createdPayment) {
          response.success = true;
          response.message = 'Payment created successfully';
          response.data = newPayment;
          auditLog(createPaymentUsername, 'PAYMENT_CREATE', `Created payment: ${newPayment.uniqueId}`);
        } else {
          response.success = false;
          response.message = 'Failed to create payment';
        }
        break;
        
      case 'readPayments':
        const readPaymentUsername = e.parameter.username;
        const includeDeletedPayments = e.parameter.includeDeleted === 'true';
        
        if (!readPaymentUsername) {
          response.message = 'Username required';
          return sendJsonResponse(response, 400);
        }
        
        let payments = getPaymentsFromSheet(readPaymentUsername);
        if (!includeDeletedPayments) {
          payments = payments.filter(payment => !payment.isDeleted);
        }
        
        response.success = true;
        response.data = payments;
        response.count = payments.length;
        break;
        
      case 'updatePayment':
        const updatePaymentUsername = e.parameter.username;
        const updatePaymentData = JSON.parse(dataParam || '{}');
        const paymentId = updatePaymentData.uniqueId;
        
        if (!updatePaymentUsername || !paymentId) {
          response.message = 'Username and payment uniqueId required';
          return sendJsonResponse(response, 400);
        }
        
        const updatedPayment = updatePaymentInSheet(updatePaymentUsername, paymentId, updatePaymentData);
        if (updatedPayment) {
          response.success = true;
          response.message = 'Payment updated successfully';
          response.data = updatedPayment;
          auditLog(updatePaymentUsername, 'PAYMENT_UPDATE', `Updated payment: ${paymentId}`);
        } else {
          response.success = false;
          response.message = 'Failed to update payment';
        }
        break;
        
      case 'deletePayment':
        const deletePaymentUsername = e.parameter.username;
        const deletePaymentId = e.parameter.paymentId;
        const permanentDeletePayment = e.parameter.permanent === 'true';
        
        if (!deletePaymentUsername || !deletePaymentId) {
          response.message = 'Username and paymentId required';
          return sendJsonResponse(response, 400);
        }
        
        let deletePaymentResult;
        if (permanentDeletePayment) {
          deletePaymentResult = permanentlyDeletePayment(deletePaymentUsername, deletePaymentId);
          response.message = 'Payment permanently deleted';
        } else {
          deletePaymentResult = softDeletePayment(deletePaymentUsername, deletePaymentId);
          response.message = 'Payment soft deleted (moved to trash)';
        }
        
        if (deletePaymentResult) {
          response.success = true;
          auditLog(deletePaymentUsername, 'PAYMENT_DELETE', `Deleted payment: ${deletePaymentId}`);
        } else {
          response.success = false;
          response.message = 'Failed to delete payment';
        }
        break;
        
      case 'restorePayment':
        const restorePaymentUsername = e.parameter.username;
        const restorePaymentId = e.parameter.paymentId;
        
        if (!restorePaymentUsername || !restorePaymentId) {
          response.message = 'Username and paymentId required';
          return sendJsonResponse(response, 400);
        }
        
        const restoredPayment = restorePaymentFromSheet(restorePaymentUsername, restorePaymentId);
        if (restoredPayment) {
          response.success = true;
          response.message = 'Payment restored successfully';
          response.data = restoredPayment;
          auditLog(restorePaymentUsername, 'PAYMENT_RESTORE', `Restored payment: ${restorePaymentId}`);
        } else {
          response.success = false;
          response.message = 'Failed to restore payment';
        }
        break;
      
      // ========== DRAWING DATA OPERATIONS ==========
      case 'saveDrawing':
        const drawingUsername = e.parameter.username;
        const drawingData = e.parameter.drawingData;
        const drawingId = e.parameter.drawingId || generateUniqueId();
        
        if (!drawingUsername || !drawingData) {
          response.message = 'Username and drawingData required';
          return sendJsonResponse(response, 400);
        }
        
        const savedDrawing = saveDrawingToSheet(drawingUsername, drawingId, drawingData);
        if (savedDrawing) {
          response.success = true;
          response.message = 'Drawing saved successfully';
          response.data = { drawingId: drawingId };
          auditLog(drawingUsername, 'DRAWING_SAVE', `Saved drawing: ${drawingId}`);
        } else {
          response.success = false;
          response.message = 'Failed to save drawing';
        }
        break;
        
      case 'getDrawing':
        const getDrawingUsername = e.parameter.username;
        const getDrawingId = e.parameter.drawingId;
        
        if (!getDrawingUsername || !getDrawingId) {
          response.message = 'Username and drawingId required';
          return sendJsonResponse(response, 400);
        }
        
        const drawing = getDrawingFromSheet(getDrawingUsername, getDrawingId);
        if (drawing) {
          response.success = true;
          response.data = drawing;
        } else {
          response.success = false;
          response.message = 'Drawing not found';
        }
        break;
        
      case 'getAllDrawings':
        const allDrawingsUsername = e.parameter.username;
        
        if (!allDrawingsUsername) {
          response.message = 'Username required';
          return sendJsonResponse(response, 400);
        }
        
        const drawings = getAllDrawingsFromSheet(allDrawingsUsername);
        response.success = true;
        response.data = drawings;
        response.count = drawings.length;
        break;
        
      case 'deleteDrawing':
        const deleteDrawingUsername = e.parameter.username;
        const deleteDrawingId = e.parameter.drawingId;
        
        if (!deleteDrawingUsername || !deleteDrawingId) {
          response.message = 'Username and drawingId required';
          return sendJsonResponse(response, 400);
        }
        
        const drawingDeleted = deleteDrawingFromSheet(deleteDrawingUsername, deleteDrawingId);
        if (drawingDeleted) {
          response.success = true;
          response.message = 'Drawing deleted successfully';
          auditLog(deleteDrawingUsername, 'DRAWING_DELETE', `Deleted drawing: ${deleteDrawingId}`);
        } else {
          response.success = false;
          response.message = 'Failed to delete drawing';
        }
        break;
      
      // ========== MERGE & SYNC OPERATIONS ==========
      case 'restoreV35':
        const restoreUsername = e.parameter.username;
        
        if (!restoreUsername) {
          response.message = 'Username required';
          return sendJsonResponse(response, 400);
        }
        
        const userBons = getBonsFromSheet(restoreUsername);
        const userPayments = getPaymentsFromSheet(restoreUsername);
        const userDrawings = getAllDrawingsFromSheet(restoreUsername);
        
        response.success = true;
        response.semuaBon = userBons.filter(bon => !bon.isDeleted);
        response.pembayaran = userPayments.filter(payment => !payment.isDeleted);
        response.gambar = userDrawings;
        response.deletedBons = userBons.filter(bon => bon.isDeleted);
        response.deletedPayments = userPayments.filter(payment => payment.isDeleted);
        response.lastModified = new Date().toISOString();
        response.serverTimestamp = new Date().toISOString();
        break;
        
      case 'mergeBackupV35':
        const backupData = JSON.parse(dataParam || '{}');
        const backupUsername = backupData.username;
        
        if (!backupUsername) {
          response.message = 'Username required in backup data';
          return sendJsonResponse(response, 400);
        }
        
        const existingBons = getBonsFromSheet(backupUsername);
        const existingPayments = getPaymentsFromSheet(backupUsername);
        
        const clientBons = backupData.semuaBon || [];
        const clientPayments = backupData.pembayaran || [];
        const clientDrawings = backupData.gambar || [];
        
        const mergedBons = mergeDataWithTimestamp(existingBons, clientBons, 'uniqueId');
        const mergedPayments = mergeDataWithTimestamp(existingPayments, clientPayments, 'uniqueId');
        
        const bonsSaved = saveBonsToSheet(backupUsername, mergedBons);
        const paymentsSaved = savePaymentsToSheet(backupUsername, mergedPayments);
        
        for (const drawing of clientDrawings) {
          saveDrawingToSheet(backupUsername, drawing.drawingId || generateUniqueId(), drawing.drawingData);
        }
        
        if (bonsSaved && paymentsSaved) {
          response.success = true;
          response.message = 'Merge backup successful';
          response.bonsCount = mergedBons.length;
          response.paymentsCount = mergedPayments.length;
          response.drawingsCount = clientDrawings.length;
          auditLog(backupUsername, 'MERGE_BACKUP', `Merged ${mergedBons.length} bons, ${mergedPayments.length} payments`);
        } else {
          response.success = false;
          response.message = 'Failed to save merged data';
        }
        break;
        
      case 'syncBonV35':
        const syncBonUsername = e.parameter.username;
        const syncBonData = JSON.parse(dataParam || '{}');
        
        if (!syncBonUsername || !syncBonData) {
          response.message = 'Username and bonData required';
          return sendJsonResponse(response, 400);
        }
        
        const bonSavedResult = saveSingleBonToSheet(syncBonUsername, syncBonData);
        if (bonSavedResult) {
          response.success = true;
          response.message = 'Bon synced successfully';
        } else {
          response.success = false;
          response.message = 'Failed to sync bon';
        }
        break;
        
      case 'syncPaymentV35':
        const syncPaymentUsername = e.parameter.username;
        const syncPaymentData = JSON.parse(dataParam || '{}');
        
        if (!syncPaymentUsername || !syncPaymentData) {
          response.message = 'Username and paymentData required';
          return sendJsonResponse(response, 400);
        }
        
        const paymentSavedResult = saveSinglePaymentToSheet(syncPaymentUsername, syncPaymentData);
        if (paymentSavedResult) {
          response.success = true;
          response.message = 'Payment synced successfully';
        } else {
          response.success = false;
          response.message = 'Failed to sync payment';
        }
        break;
      
      // ========== SEARCH OPERATIONS ==========
      case 'searchBons':
        const searchBonUsername = e.parameter.username;
        const searchKeyword = e.parameter.keyword || '';
        const searchField = e.parameter.field || 'all';
        
        if (!searchBonUsername) {
          response.message = 'Username required';
          return sendJsonResponse(response, 400);
        }
        
        const allUserBons = getBonsFromSheet(searchBonUsername).filter(bon => !bon.isDeleted);
        const searchResults = searchBonsData(allUserBons, searchKeyword, searchField);
        
        response.success = true;
        response.data = searchResults;
        response.count = searchResults.length;
        response.keyword = searchKeyword;
        break;
        
      case 'getStats':
        const statsUsername = e.parameter.username;
        
        if (!statsUsername) {
          response.message = 'Username required';
          return sendJsonResponse(response, 400);
        }
        
        const statsBons = getBonsFromSheet(statsUsername).filter(bon => !bon.isDeleted);
        const statsPayments = getPaymentsFromSheet(statsUsername).filter(payment => !payment.isDeleted);
        
        const totalUtang = statsBons.reduce((sum, bon) => sum + (parseFloat(bon.nominal) || 0), 0);
        const totalBayar = statsPayments.reduce((sum, payment) => sum + (parseFloat(payment.nominal) || 0), 0);
        
        response.success = true;
        response.data = {
          totalBon: statsBons.length,
          totalPayment: statsPayments.length,
          totalUtang: totalUtang,
          totalBayar: totalBayar,
          saldo: totalUtang - totalBayar,
          lastUpdated: new Date().toISOString()
        };
        break;
      
      // ========== BACKUP & RESTORE ==========
      case 'createBackup':
        const backupUserUsername = e.parameter.username;
        
        if (!backupUserUsername) {
          response.message = 'Username required';
          return sendJsonResponse(response, 400);
        }
        
        const backupId = createFullBackup(backupUserUsername);
        if (backupId) {
          response.success = true;
          response.message = 'Backup created successfully';
          response.data = { backupId: backupId };
          auditLog(backupUserUsername, 'BACKUP_CREATE', `Created backup: ${backupId}`);
        } else {
          response.success = false;
          response.message = 'Failed to create backup';
        }
        break;
        
      case 'restoreFromBackup':
        const restoreBackupUsername = e.parameter.username;
        const restoreBackupId = e.parameter.backupId;
        
        if (!restoreBackupUsername || !restoreBackupId) {
          response.message = 'Username and backupId required';
          return sendJsonResponse(response, 400);
        }
        
        const restoredFromBackup = restoreFromBackup(restoreBackupUsername, restoreBackupId);
        if (restoredFromBackup) {
          response.success = true;
          response.message = 'Data restored from backup successfully';
          auditLog(restoreBackupUsername, 'RESTORE_BACKUP', `Restored from backup: ${restoreBackupId}`);
        } else {
          response.success = false;
          response.message = 'Failed to restore from backup';
        }
        break;
        
      default:
        response.message = `Unknown action: ${action}`;
        return sendJsonResponse(response, 400);
    }
    
    return sendJsonResponse(response);
    
  } catch(error) {
    Logger.log(`Error in doPost: ${error.toString()}`);
    response.success = false;
    response.message = `Server error: ${error.toString()}`;
    return sendJsonResponse(response, 500);
  }
}

// ==================== DO GET ====================
function doGet() {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Bon Warung Cloud Backup v35.0</title>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; }
        .container { max-width: 1200px; margin: auto; background: white; padding: 30px; border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); }
        h1 { color: #667eea; margin-bottom: 10px; }
        .status { color: #38a169; font-weight: bold; background: #c6f6d5; padding: 5px 10px; border-radius: 20px; display: inline-block; }
        .info { background: #ebf8ff; padding: 15px; border-radius: 10px; margin: 20px 0; }
        .endpoint-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 15px; margin: 20px 0; }
        .endpoint-card { background: #f7fafc; padding: 15px; border-radius: 10px; border-left: 4px solid #667eea; }
        .endpoint-card h4 { color: #2d3748; margin-bottom: 5px; }
        .endpoint-card p { color: #718096; font-size: 12px; font-family: monospace; }
        .badge { background: #e2e8f0; padding: 2px 8px; border-radius: 10px; font-size: 11px; margin-left: 10px; }
        .feature-list { display: flex; flex-wrap: wrap; gap: 10px; margin: 20px 0; }
        .feature { background: #e6fffa; padding: 5px 12px; border-radius: 15px; font-size: 12px; color: #234e52; }
        hr { margin: 20px 0; border-color: #e2e8f0; }
        .footer { text-align: center; margin-top: 20px; color: #718096; font-size: 12px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>☁️ Bon Warung Cloud Backup v35.0</h1>
        <p><span class="status">✅ ONLINE & SIAP</span> <span class="badge">Production Ready</span></p>
        <p>Server Time: ${new Date().toLocaleString('id-ID')}</p>
        
        <div class="info">
          <strong>📊 Sistem CRUD Lengkap untuk Data Warung</strong><br>
          Support: Create | Read | Update | Delete (Soft & Permanent) | Restore | Merge | Drawing | Search | Backup
        </div>
        
        <h3>✨ Fitur Lengkap:</h3>
        <div class="feature-list">
          <span class="feature">✅ Create Bon</span>
          <span class="feature">✅ Read Bons</span>
          <span class="feature">✅ Update Bon</span>
          <span class="feature">✅ Delete Bon (Soft)</span>
          <span class="feature">✅ Delete Bon (Permanent)</span>
          <span class="feature">✅ Restore Bon</span>
          <span class="feature">✅ Create Payment</span>
          <span class="feature">✅ Read Payments</span>
          <span class="feature">✅ Update Payment</span>
          <span class="feature">✅ Delete Payment</span>
          <span class="feature">✅ Restore Payment</span>
          <span class="feature">✅ Save Drawing</span>
          <span class="feature">✅ Get Drawing</span>
          <span class="feature">✅ Search Data</span>
          <span class="feature">✅ Get Statistics</span>
          <span class="feature">✅ Create Backup</span>
          <span class="feature">✅ Restore Backup</span>
          <span class="feature">✅ Merge Sync</span>
          <span class="feature">✅ Audit Log</span>
        </div>
        
        <hr>
        
        <h3>📡 Endpoint Actions:</h3>
        <div class="endpoint-grid">
          <div class="endpoint-card"><h4>🔐 User</h4><p>getUserAuth, syncUserAuth</p></div>
          <div class="endpoint-card"><h4>📝 Bon CRUD</h4><p>createBon, readBons, updateBon, deleteBon, restoreBon</p></div>
          <div class="endpoint-card"><h4>💰 Payment CRUD</h4><p>createPayment, readPayments, updatePayment, deletePayment, restorePayment</p></div>
          <div class="endpoint-card"><h4>🎨 Drawing</h4><p>saveDrawing, getDrawing, getAllDrawings, deleteDrawing</p></div>
          <div class="endpoint-card"><h4>🔄 Sync & Merge</h4><p>restoreV35, mergeBackupV35, syncBonV35, syncPaymentV35</p></div>
          <div class="endpoint-card"><h4>🔍 Search & Stats</h4><p>searchBons, getStats</p></div>
          <div class="endpoint-card"><h4>💾 Backup</h4><p>createBackup, restoreFromBackup</p></div>
          <div class="endpoint-card"><h4>🔧 Utility</h4><p>testConnection</p></div>
        </div>
        
        <hr>
        
        <h3>📖 Cara Penggunaan:</h3>
        <div class="info">
          <strong>Deploy Web App:</strong><br>
          1. Buka editor Apps Script<br>
          2. Klik Deploy → New deployment<br>
          3. Pilih "Web app"<br>
          4. Execute as: "Me", Who has access: "Anyone"<br>
          5. Copy URL yang dihasilkan untuk digunakan di aplikasi<br><br>
          <strong>⚠️ Penting:</strong> Ganti SPREADSHEET_ID dengan ID spreadsheet Anda sendiri!
        </div>
        
        <div class="footer">
          Bon Warung v35.0 | Complete CRUD Cloud Backend | Multi-device Sync with Merge Logic
        </div>
      </div>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html);
}

// ==================== SPREADSHEET UTILITIES ====================
function getOrCreateSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    return sheet;
  } catch(error) {
    Logger.log(`Error getting/creating sheet ${sheetName}: ${error.toString()}`);
    throw error;
  }
}

function ensureSheetHasHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }
  
  const range = sheet.getRange(1, 1, 1, headers.length);
  const currentHeaders = range.getValues()[0];
  
  let needsUpdate = false;
  for (let i = 0; i < headers.length; i++) {
    if (currentHeaders[i] !== headers[i]) {
      needsUpdate = true;
      break;
    }
  }
  
  if (needsUpdate) {
    range.setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

// ==================== USER MANAGEMENT ====================
function getUserFromSheet(username) {
  try {
    const sheet = getOrCreateSheet(SHEET_USERS);
    const headers = ['username', 'userData', 'lastUpdated', 'deviceId', 'lastLogin'];
    ensureSheetHasHeaders(sheet, headers);
    
    const data = sheet.getDataRange().getValues();
    const usernameLower = username.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower) {
        const userDataStr = data[i][1];
        if (userDataStr) {
          return JSON.parse(userDataStr);
        }
        return null;
      }
    }
    return null;
  } catch(error) {
    Logger.log(`getUserFromSheet error: ${error.toString()}`);
    return null;
  }
}

function saveUserToSheet(username, userData) {
  try {
    const sheet = getOrCreateSheet(SHEET_USERS);
    const headers = ['username', 'userData', 'lastUpdated', 'deviceId', 'lastLogin'];
    ensureSheetHasHeaders(sheet, headers);
    
    const data = sheet.getDataRange().getValues();
    const usernameLower = username.toLowerCase();
    let rowToUpdate = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower) {
        rowToUpdate = i + 1;
        break;
      }
    }
    
    const now = new Date().toISOString();
    const userDataStr = JSON.stringify(userData);
    const deviceId = userData.deviceId || '';
    const lastLogin = userData.lastLogin || now;
    
    if (rowToUpdate !== -1) {
      sheet.getRange(rowToUpdate, 2, 1, 4).setValues([[userDataStr, now, deviceId, lastLogin]]);
    } else {
      sheet.appendRow([usernameLower, userDataStr, now, deviceId, lastLogin]);
    }
    
    return true;
  } catch(error) {
    Logger.log(`saveUserToSheet error: ${error.toString()}`);
    return false;
  }
}

// ==================== BONS MANAGEMENT (Full CRUD) ====================
function getBonsFromSheet(username) {
  try {
    const sheet = getOrCreateSheet(SHEET_BONS);
    const headers = ['username', 'uniqueId', 'bonData', 'lastModified', 'deviceId', 'isDeleted', 'deletedAt'];
    ensureSheetHasHeaders(sheet, headers);
    
    const data = sheet.getDataRange().getValues();
    const usernameLower = username.toLowerCase();
    const bons = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower) {
        const bonDataStr = data[i][2];
        if (bonDataStr) {
          try {
            const bon = JSON.parse(bonDataStr);
            bon.isDeleted = data[i][5] === true;
            bon.deletedAt = data[i][6];
            if (data[i][3]) {
              bon.lastModified = data[i][3];
            }
            bons.push(bon);
          } catch(e) {
            Logger.log(`Error parsing bon data: ${e.toString()}`);
          }
        }
      }
    }
    
    return bons;
  } catch(error) {
    Logger.log(`getBonsFromSheet error: ${error.toString()}`);
    return [];
  }
}

function saveBonsToSheet(username, bons) {
  try {
    const sheet = getOrCreateSheet(SHEET_BONS);
    const headers = ['username', 'uniqueId', 'bonData', 'lastModified', 'deviceId', 'isDeleted', 'deletedAt'];
    ensureSheetHasHeaders(sheet, headers);
    
    const usernameLower = username.toLowerCase();
    
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower) {
        rowsToDelete.push(i + 1);
      }
    }
    
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    
    for (const bon of bons) {
      const uniqueId = bon.uniqueId || generateUniqueId();
      const lastModified = bon.lastModified || new Date().toISOString();
      const deviceId = bon.deviceId || '';
      const isDeleted = bon.isDeleted || false;
      const deletedAt = bon.deletedAt || '';
      const bonToSave = { ...bon, uniqueId, lastModified };
      const bonDataStr = JSON.stringify(bonToSave);
      
      sheet.appendRow([usernameLower, uniqueId, bonDataStr, lastModified, deviceId, isDeleted, deletedAt]);
    }
    
    return true;
  } catch(error) {
    Logger.log(`saveBonsToSheet error: ${error.toString()}`);
    return false;
  }
}

function saveSingleBonToSheet(username, bon) {
  try {
    const sheet = getOrCreateSheet(SHEET_BONS);
    const headers = ['username', 'uniqueId', 'bonData', 'lastModified', 'deviceId', 'isDeleted', 'deletedAt'];
    ensureSheetHasHeaders(sheet, headers);
    
    const usernameLower = username.toLowerCase();
    const uniqueId = bon.uniqueId || generateUniqueId();
    const lastModified = bon.lastModified || new Date().toISOString();
    const deviceId = bon.deviceId || '';
    const isDeleted = bon.isDeleted || false;
    const deletedAt = bon.deletedAt || '';
    
    const data = sheet.getDataRange().getValues();
    let existingRow = -1;
    let existingLastModified = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === uniqueId) {
        existingRow = i + 1;
        existingLastModified = data[i][3];
        break;
      }
    }
    
    if (existingRow !== -1 && existingLastModified) {
      const existingTime = new Date(existingLastModified).getTime();
      const newTime = new Date(lastModified).getTime();
      
      if (newTime < existingTime) {
        return true;
      }
    }
    
    const bonToSave = { ...bon, uniqueId, lastModified };
    const bonDataStr = JSON.stringify(bonToSave);
    
    if (existingRow !== -1) {
      sheet.getRange(existingRow, 3, 1, 5).setValues([[bonDataStr, lastModified, deviceId, isDeleted, deletedAt]]);
    } else {
      sheet.appendRow([usernameLower, uniqueId, bonDataStr, lastModified, deviceId, isDeleted, deletedAt]);
    }
    
    return true;
  } catch(error) {
    Logger.log(`saveSingleBonToSheet error: ${error.toString()}`);
    return false;
  }
}

function updateBonInSheet(username, bonId, updatedData) {
  try {
    const sheet = getOrCreateSheet(SHEET_BONS);
    const headers = ['username', 'uniqueId', 'bonData', 'lastModified', 'deviceId', 'isDeleted', 'deletedAt'];
    ensureSheetHasHeaders(sheet, headers);
    
    const usernameLower = username.toLowerCase();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === bonId) {
        
        const existingBonData = JSON.parse(data[i][2]);
        const updatedBon = {
          ...existingBonData,
          ...updatedData,
          uniqueId: bonId,
          lastModified: new Date().toISOString()
        };
        
        const bonDataStr = JSON.stringify(updatedBon);
        const newLastModified = new Date().toISOString();
        const deviceId = updatedBon.deviceId || data[i][4] || '';
        const isDeleted = updatedBon.isDeleted !== undefined ? updatedBon.isDeleted : (data[i][5] || false);
        const deletedAt = updatedBon.deletedAt || data[i][6] || '';
        
        sheet.getRange(i + 1, 3, 1, 5).setValues([[bonDataStr, newLastModified, deviceId, isDeleted, deletedAt]]);
        
        return updatedBon;
      }
    }
    
    return null;
  } catch(error) {
    Logger.log(`updateBonInSheet error: ${error.toString()}`);
    return null;
  }
}

function softDeleteBon(username, bonId) {
  try {
    const sheet = getOrCreateSheet(SHEET_BONS);
    const usernameLower = username.toLowerCase();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === bonId) {
        
        sheet.getRange(i + 1, 6, 1, 2).setValues([[true, new Date().toISOString()]]);
        
        const bonData = JSON.parse(data[i][2]);
        bonData.isDeleted = true;
        bonData.deletedAt = new Date().toISOString();
        sheet.getRange(i + 1, 3).setValue(JSON.stringify(bonData));
        
        return true;
      }
    }
    
    return false;
  } catch(error) {
    Logger.log(`softDeleteBon error: ${error.toString()}`);
    return false;
  }
}

function permanentlyDeleteBon(username, bonId) {
  try {
    const sheet = getOrCreateSheet(SHEET_BONS);
    const usernameLower = username.toLowerCase();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === bonId) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    
    return false;
  } catch(error) {
    Logger.log(`permanentlyDeleteBon error: ${error.toString()}`);
    return false;
  }
}

function restoreBonFromSheet(username, bonId) {
  try {
    const sheet = getOrCreateSheet(SHEET_BONS);
    const usernameLower = username.toLowerCase();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === bonId) {
        
        sheet.getRange(i + 1, 6, 1, 2).setValues([[false, '']]);
        
        const bonData = JSON.parse(data[i][2]);
        bonData.isDeleted = false;
        bonData.deletedAt = null;
        bonData.lastModified = new Date().toISOString();
        sheet.getRange(i + 1, 3).setValue(JSON.stringify(bonData));
        
        return bonData;
      }
    }
    
    return null;
  } catch(error) {
    Logger.log(`restoreBonFromSheet error: ${error.toString()}`);
    return null;
  }
}

// ==================== PAYMENTS MANAGEMENT (Full CRUD) ====================
function getPaymentsFromSheet(username) {
  try {
    const sheet = getOrCreateSheet(SHEET_PAYMENTS);
    const headers = ['username', 'uniqueId', 'paymentData', 'lastModified', 'deviceId', 'isDeleted', 'deletedAt'];
    ensureSheetHasHeaders(sheet, headers);
    
    const data = sheet.getDataRange().getValues();
    const usernameLower = username.toLowerCase();
    const payments = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower) {
        const paymentDataStr = data[i][2];
        if (paymentDataStr) {
          try {
            const payment = JSON.parse(paymentDataStr);
            payment.isDeleted = data[i][5] === true;
            payment.deletedAt = data[i][6];
            if (data[i][3]) {
              payment.lastModified = data[i][3];
            }
            payments.push(payment);
          } catch(e) {
            Logger.log(`Error parsing payment data: ${e.toString()}`);
          }
        }
      }
    }
    
    return payments;
  } catch(error) {
    Logger.log(`getPaymentsFromSheet error: ${error.toString()}`);
    return [];
  }
}

function savePaymentsToSheet(username, payments) {
  try {
    const sheet = getOrCreateSheet(SHEET_PAYMENTS);
    const headers = ['username', 'uniqueId', 'paymentData', 'lastModified', 'deviceId', 'isDeleted', 'deletedAt'];
    ensureSheetHasHeaders(sheet, headers);
    
    const usernameLower = username.toLowerCase();
    
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower) {
        rowsToDelete.push(i + 1);
      }
    }
    
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    
    for (const payment of payments) {
      const uniqueId = payment.uniqueId || generateUniqueId();
      const lastModified = payment.lastModified || new Date().toISOString();
      const deviceId = payment.deviceId || '';
      const isDeleted = payment.isDeleted || false;
      const deletedAt = payment.deletedAt || '';
      const paymentToSave = { ...payment, uniqueId, lastModified };
      const paymentDataStr = JSON.stringify(paymentToSave);
      
      sheet.appendRow([usernameLower, uniqueId, paymentDataStr, lastModified, deviceId, isDeleted, deletedAt]);
    }
    
    return true;
  } catch(error) {
    Logger.log(`savePaymentsToSheet error: ${error.toString()}`);
    return false;
  }
}

function saveSinglePaymentToSheet(username, payment) {
  try {
    const sheet = getOrCreateSheet(SHEET_PAYMENTS);
    const headers = ['username', 'uniqueId', 'paymentData', 'lastModified', 'deviceId', 'isDeleted', 'deletedAt'];
    ensureSheetHasHeaders(sheet, headers);
    
    const usernameLower = username.toLowerCase();
    const uniqueId = payment.uniqueId || generateUniqueId();
    const lastModified = payment.lastModified || new Date().toISOString();
    const deviceId = payment.deviceId || '';
    const isDeleted = payment.isDeleted || false;
    const deletedAt = payment.deletedAt || '';
    
    const data = sheet.getDataRange().getValues();
    let existingRow = -1;
    let existingLastModified = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === uniqueId) {
        existingRow = i + 1;
        existingLastModified = data[i][3];
        break;
      }
    }
    
    if (existingRow !== -1 && existingLastModified) {
      const existingTime = new Date(existingLastModified).getTime();
      const newTime = new Date(lastModified).getTime();
      
      if (newTime < existingTime) {
        return true;
      }
    }
    
    const paymentToSave = { ...payment, uniqueId, lastModified };
    const paymentDataStr = JSON.stringify(paymentToSave);
    
    if (existingRow !== -1) {
      sheet.getRange(existingRow, 3, 1, 5).setValues([[paymentDataStr, lastModified, deviceId, isDeleted, deletedAt]]);
    } else {
      sheet.appendRow([usernameLower, uniqueId, paymentDataStr, lastModified, deviceId, isDeleted, deletedAt]);
    }
    
    return true;
  } catch(error) {
    Logger.log(`saveSinglePaymentToSheet error: ${error.toString()}`);
    return false;
  }
}

function updatePaymentInSheet(username, paymentId, updatedData) {
  try {
    const sheet = getOrCreateSheet(SHEET_PAYMENTS);
    const headers = ['username', 'uniqueId', 'paymentData', 'lastModified', 'deviceId', 'isDeleted', 'deletedAt'];
    ensureSheetHasHeaders(sheet, headers);
    
    const usernameLower = username.toLowerCase();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === paymentId) {
        
        const existingPaymentData = JSON.parse(data[i][2]);
        const updatedPayment = {
          ...existingPaymentData,
          ...updatedData,
          uniqueId: paymentId,
          lastModified: new Date().toISOString()
        };
        
        const paymentDataStr = JSON.stringify(updatedPayment);
        const newLastModified = new Date().toISOString();
        const deviceId = updatedPayment.deviceId || data[i][4] || '';
        const isDeleted = updatedPayment.isDeleted !== undefined ? updatedPayment.isDeleted : (data[i][5] || false);
        const deletedAt = updatedPayment.deletedAt || data[i][6] || '';
        
        sheet.getRange(i + 1, 3, 1, 5).setValues([[paymentDataStr, newLastModified, deviceId, isDeleted, deletedAt]]);
        
        return updatedPayment;
      }
    }
    
    return null;
  } catch(error) {
    Logger.log(`updatePaymentInSheet error: ${error.toString()}`);
    return null;
  }
}

function softDeletePayment(username, paymentId) {
  try {
    const sheet = getOrCreateSheet(SHEET_PAYMENTS);
    const usernameLower = username.toLowerCase();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === paymentId) {
        
        sheet.getRange(i + 1, 6, 1, 2).setValues([[true, new Date().toISOString()]]);
        
        const paymentData = JSON.parse(data[i][2]);
        paymentData.isDeleted = true;
        paymentData.deletedAt = new Date().toISOString();
        sheet.getRange(i + 1, 3).setValue(JSON.stringify(paymentData));
        
        return true;
      }
    }
    
    return false;
  } catch(error) {
    Logger.log(`softDeletePayment error: ${error.toString()}`);
    return false;
  }
}

function permanentlyDeletePayment(username, paymentId) {
  try {
    const sheet = getOrCreateSheet(SHEET_PAYMENTS);
    const usernameLower = username.toLowerCase();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === paymentId) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    
    return false;
  } catch(error) {
    Logger.log(`permanentlyDeletePayment error: ${error.toString()}`);
    return false;
  }
}

function restorePaymentFromSheet(username, paymentId) {
  try {
    const sheet = getOrCreateSheet(SHEET_PAYMENTS);
    const usernameLower = username.toLowerCase();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === paymentId) {
        
        sheet.getRange(i + 1, 6, 1, 2).setValues([[false, '']]);
        
        const paymentData = JSON.parse(data[i][2]);
        paymentData.isDeleted = false;
        paymentData.deletedAt = null;
        paymentData.lastModified = new Date().toISOString();
        sheet.getRange(i + 1, 3).setValue(JSON.stringify(paymentData));
        
        return paymentData;
      }
    }
    
    return null;
  } catch(error) {
    Logger.log(`restorePaymentFromSheet error: ${error.toString()}`);
    return null;
  }
}

// ==================== DRAWING MANAGEMENT ====================
function saveDrawingToSheet(username, drawingId, drawingData) {
  try {
    const sheet = getOrCreateSheet(SHEET_DRAWINGS);
    const headers = ['username', 'drawingId', 'drawingData', 'lastModified', 'deviceId'];
    ensureSheetHasHeaders(sheet, headers);
    
    const usernameLower = username.toLowerCase();
    const lastModified = new Date().toISOString();
    const data = sheet.getDataRange().getValues();
    let existingRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === drawingId) {
        existingRow = i + 1;
        break;
      }
    }
    
    if (existingRow !== -1) {
      sheet.getRange(existingRow, 3, 1, 3).setValues([[drawingData, lastModified, '']]);
    } else {
      sheet.appendRow([usernameLower, drawingId, drawingData, lastModified, '']);
    }
    
    return true;
  } catch(error) {
    Logger.log(`saveDrawingToSheet error: ${error.toString()}`);
    return false;
  }
}

function getDrawingFromSheet(username, drawingId) {
  try {
    const sheet = getOrCreateSheet(SHEET_DRAWINGS);
    const data = sheet.getDataRange().getValues();
    const usernameLower = username.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === drawingId) {
        return {
          drawingId: drawingId,
          drawingData: data[i][2],
          lastModified: data[i][3]
        };
      }
    }
    
    return null;
  } catch(error) {
    Logger.log(`getDrawingFromSheet error: ${error.toString()}`);
    return null;
  }
}

function getAllDrawingsFromSheet(username) {
  try {
    const sheet = getOrCreateSheet(SHEET_DRAWINGS);
    const data = sheet.getDataRange().getValues();
    const usernameLower = username.toLowerCase();
    const drawings = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower) {
        drawings.push({
          drawingId: data[i][1],
          drawingData: data[i][2],
          lastModified: data[i][3]
        });
      }
    }
    
    return drawings;
  } catch(error) {
    Logger.log(`getAllDrawingsFromSheet error: ${error.toString()}`);
    return [];
  }
}

function deleteDrawingFromSheet(username, drawingId) {
  try {
    const sheet = getOrCreateSheet(SHEET_DRAWINGS);
    const data = sheet.getDataRange().getValues();
    const usernameLower = username.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === usernameLower &&
          data[i][1] && data[i][1].toString() === drawingId) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    
    return false;
  } catch(error) {
    Logger.log(`deleteDrawingFromSheet error: ${error.toString()}`);
    return false;
  }
}

// ==================== SEARCH FUNCTION ====================
function searchBonsData(bons, keyword, field) {
  if (!keyword || keyword.trim() === '') {
    return bons;
  }
  
  const lowerKeyword = keyword.toLowerCase();
  
  return bons.filter(bon => {
    if (field === 'all') {
      return (
        (bon.namaPeminjam && bon.namaPeminjam.toLowerCase().includes(lowerKeyword)) ||
        (bon.nominal && bon.nominal.toString().includes(lowerKeyword)) ||
        (bon.keterangan && bon.keterangan.toLowerCase().includes(lowerKeyword)) ||
        (bon.tanggal && bon.tanggal.includes(lowerKeyword)) ||
        (bon.uniqueId && bon.uniqueId.includes(lowerKeyword))
      );
    } else if (field === 'nama') {
      return bon.namaPeminjam && bon.namaPeminjam.toLowerCase().includes(lowerKeyword);
    } else if (field === 'nominal') {
      return bon.nominal && bon.nominal.toString().includes(lowerKeyword);
    } else if (field === 'keterangan') {
      return bon.keterangan && bon.keterangan.toLowerCase().includes(lowerKeyword);
    }
    return false;
  });
}

// ==================== BACKUP MANAGEMENT ====================
function createFullBackup(username) {
  try {
    const backupId = generateUniqueId();
    const sheet = getOrCreateSheet(SHEET_BACKUP);
    const headers = ['backupId', 'username', 'backupData', 'createdAt'];
    ensureSheetHasHeaders(sheet, headers);
    
    const bons = getBonsFromSheet(username);
    const payments = getPaymentsFromSheet(username);
    const drawings = getAllDrawingsFromSheet(username);
    
    const backupData = {
      username: username,
      bons: bons,
      payments: payments,
      drawings: drawings,
      createdAt: new Date().toISOString(),
      backupId: backupId
    };
    
    sheet.appendRow([backupId, username.toLowerCase(), JSON.stringify(backupData), new Date().toISOString()]);
    
    return backupId;
  } catch(error) {
    Logger.log(`createFullBackup error: ${error.toString()}`);
    return null;
  }
}

function restoreFromBackup(username, backupId) {
  try {
    const sheet = getOrCreateSheet(SHEET_BACKUP);
    const data = sheet.getDataRange().getValues();
    const usernameLower = username.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === backupId &&
          data[i][1] && data[i][1].toString().toLowerCase() === usernameLower) {
        
        const backupData = JSON.parse(data[i][2]);
        
        saveBonsToSheet(username, backupData.bons);
        savePaymentsToSheet(username, backupData.payments);
        
        for (const drawing of backupData.drawings) {
          saveDrawingToSheet(username, drawing.drawingId, drawing.drawingData);
        }
        
        return true;
      }
    }
    
    return false;
  } catch(error) {
    Logger.log(`restoreFromBackup error: ${error.toString()}`);
    return false;
  }
}

// ==================== AUDIT LOG ====================
function auditLog(username, action, details) {
  try {
    const sheet = getOrCreateSheet(SHEET_AUDIT_LOG);
    const headers = ['timestamp', 'username', 'action', 'details', 'ipAddress'];
    ensureSheetHasHeaders(sheet, headers);
    
    sheet.appendRow([new Date().toISOString(), username, action, details, '']);
  } catch(error) {
    Logger.log(`auditLog error: ${error.toString()}`);
  }
}

// ==================== MERGE & UTILITY FUNCTIONS ====================
function mergeDataWithTimestamp(existingData, newData, idField) {
  const mergedMap = new Map();
  
  for (const item of existingData) {
    const id = item[idField];
    if (id) {
      const timestamp = new Date(item.lastModified || 0).getTime();
      mergedMap.set(id, {
        data: item,
        timestamp: timestamp
      });
    }
  }
  
  for (const item of newData) {
    const id = item[idField];
    if (id) {
      const newTimestamp = new Date(item.lastModified || 0).getTime();
      const existing = mergedMap.get(id);
      
      if (!existing || newTimestamp > existing.timestamp) {
        mergedMap.set(id, {
          data: item,
          timestamp: newTimestamp
        });
      }
    }
  }
  
  return Array.from(mergedMap.values()).map(entry => entry.data);
}

function generateUniqueId() {
  return `${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

function sendJsonResponse(data, statusCode = 200) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  
  if (statusCode !== 200) {
    return output;
  }
  return output;
}
