// ===============================================
// ULTRA Stock - Google Apps Script Database
// ===============================================

// Sheet Names
const SHEETS = {
  ADMINS: 'Admins',           // เก็บ users ทุก role (owner, super_admin, admin, reseller, customer)
  MAIN_EMAILS: 'MainEmails',
  SUB_EMAILS: 'SubEmails',
  ORDERS: 'Orders',
  TRANSACTIONS: 'Transactions',
  COMMISSION_LOG: 'CommissionLog',
  PACKAGES: 'Packages',
  SETTINGS: 'Settings',
  PAYMENT_LOGS: 'PaymentLogs',       // บันทึกการชำระเงิน + SlipRef ป้องกันสลิปซ้ำ
  PENDING_ORDERS: 'PendingOrders',   // รอชำระ (สำหรับ Customer)
  COMMISSION_PAYMENTS: 'CommissionPayments'  // บันทึกการจ่ายค่าคอมให้ Admin
};

// Get Spreadsheet
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Get or Create Sheet
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initializeSheet(sheet, sheetName);
  }

  return sheet;
}

// Initialize Sheet with Headers
function initializeSheet(sheet, sheetName) {
  const headers = {
    [SHEETS.ADMINS]: ['Id', 'Username', 'Password', 'Role', 'Permissions', 'Commissions', 'Balance', 'CreatedAt', 'Name', 'Phone', 'LineId'],
    [SHEETS.MAIN_EMAILS]: ['Id', 'Email', 'Password', 'Capacity', 'CreatedAt', 'CreatedBy'],
    [SHEETS.SUB_EMAILS]: ['Id', 'MainEmailId', 'Email', 'Password', 'Status', 'CreatedAt', 'CreatedBy'],
    [SHEETS.ORDERS]: ['Id', 'SubEmailId', 'SubEmail', 'AdminId', 'AdminName', 'CustomerName', 'PackageDays', 'SoldAt', 'ExpiresAt', 'SaleType', 'CommissionAmount', 'Status', 'Remark', 'BatchOrderId'],
    [SHEETS.TRANSACTIONS]: ['Id', 'UserId', 'UserName', 'OrderId', 'SubEmail', 'PackageDays', 'Amount', 'Status', 'CreatedAt', 'PaidAt', 'SlipRef', 'Type'],
    [SHEETS.COMMISSION_LOG]: ['Id', 'AdminId', 'AdminName', 'OrderId', 'SubEmail', 'PackageDays', 'Amount', 'CreatedAt', 'Type'],
    [SHEETS.PACKAGES]: ['Id', 'Name', 'Days', 'ResellerPrice', 'CustomerPrice', 'Active'],
    [SHEETS.SETTINGS]: ['Key', 'Value'],
    [SHEETS.PAYMENT_LOGS]: ['Id', 'UserId', 'UserName', 'UserRole', 'Amount', 'SlipRef', 'Status', 'VerifiedAt', 'TransactionIds'],
    [SHEETS.PENDING_ORDERS]: ['Id', 'UserId', 'UserName', 'SubEmailIds', 'PackageDays', 'Amount', 'Status', 'CreatedAt', 'ExpiresAt', 'PaidAt', 'SlipRef'],
    [SHEETS.COMMISSION_PAYMENTS]: ['Id', 'AdminId', 'AdminName', 'Amount', 'Note', 'PaidBy', 'PaidAt', 'Month']
  };

  if (headers[sheetName]) {
    sheet.getRange(1, 1, 1, headers[sheetName].length).setValues([headers[sheetName]]);
    sheet.getRange(1, 1, 1, headers[sheetName].length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Initialize default owner if Admins sheet
  if (sheetName === SHEETS.ADMINS) {
    const defaultOwner = [
      generateId(),
      'owner',
      'owner123', // Change this!
      'owner',
      JSON.stringify({ all: true }),
      JSON.stringify({}),
      0,
      new Date().toISOString()
    ];
    sheet.appendRow(defaultOwner);
  }

  // Initialize default packages (ResellerPrice, CustomerPrice)
  if (sheetName === SHEETS.PACKAGES) {
    const defaultPackages = [
      [generateId(), '7 วัน', 7, 500, 799, true],
      [generateId(), '20 วัน', 20, 500, 799, true],
      [generateId(), '30 วัน', 30, 500, 799, true]
    ];
    defaultPackages.forEach(pkg => sheet.appendRow(pkg));
  }

  // Initialize default settings
  if (sheetName === SHEETS.SETTINGS) {
    const defaultSettings = [
      ['telegramBotToken', ''],
      ['telegramChatId', ''],
      ['notifyDaysBefore', '3'],
      ['bankName', 'ธ.ไทยพาณิชย์'],
      ['bankAccountNumber', '7122556905'],
      ['bankAccountName', 'จิณณวัตร ไทยพุทรา']
    ];
    defaultSettings.forEach(s => sheet.appendRow(s));
  }
}

// Generate unique ID
function generateId() {
  return 'id_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

// Web App Entry Point
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    let result;

    switch (action) {
      // Auth
      case 'login': result = login(data.username, data.password); break;
      case 'register': result = registerUser(data); break;
      case 'getUser': result = getUser(data.userId); break;
      case 'resetUserPassword': result = resetUserPassword(data.userId, data.newPassword); break;

      // Dashboard
      case 'getDashboardStats': result = getDashboardStats(data.userId, data.role); break;

      // Main Emails
      case 'getMainEmails': result = getMainEmails(data.role); break;
      case 'addMainEmail': result = addMainEmail(data.email, data.password, data.createdBy); break;
      case 'updateMainEmail': result = updateMainEmail(data.id, data.email, data.password); break;
      case 'deleteMainEmail': result = deleteMainEmail(data.id); break;

      // Sub Emails
      case 'getSubEmails': result = getSubEmails(data.role, data.status); break;
      case 'addSubEmails': result = addSubEmails(data.mainEmailId, data.emails, data.createdBy); break;
      case 'deleteSubEmail': result = deleteSubEmail(data.id); break;
      case 'searchSubEmail': result = searchSubEmail(data.email, data.role); break;

      // Orders
      case 'createOrder': result = createOrder(data); break;
      case 'getOrders': result = getOrders(data.userId, data.role, data.filter, data.saleType); break;
      case 'extendOrder': result = extendOrder(data.orderId, data.days); break;
      case 'markOrderExpired': result = markOrderExpired(data.orderId); break;

      // Expiring
      case 'getExpiringEmails': result = getExpiringEmails(data.days); break;

      // Admins
      case 'getAdmins': result = getAdmins(); break;
      case 'addAdmin': result = addAdmin(data.username, data.password, data.role, data.permissions); break;
      case 'updateAdmin': result = updateAdmin(data.id, data.username, data.password, data.role, data.permissions); break;
      case 'deleteAdmin': result = deleteAdmin(data.id); break;

      // Packages
      case 'getPackages': result = getPackages(); break;
      case 'savePackages': result = savePackages(data.packages); break;

      // Commissions
      case 'getCommissionSettings': result = getCommissionSettings(); break;
      case 'updateCommission': result = updateCommission(data.adminId, data.commissions); break;
      case 'getCommissionLog': result = getCommissionLog(data.userId, data.role, data.month); break;

      // Reseller
      case 'getResellerBalance': result = getResellerBalance(data.userId); break;
      case 'getResellerTransactions': result = getResellerTransactions(data.userId, data.role); break;
      case 'markTransactionPaid': result = markTransactionPaid(data.transactionId); break;

      // Settings
      case 'getSettings': result = getSettings(); break;
      case 'updateSettings': result = updateSettings(data); break;

      // Bank Account
      case 'getBankAccount': result = getBankAccount(); break;
      case 'updateBankAccount': result = updateBankAccount(data); break;

      // Payment (EasySlip)
      case 'checkDuplicateSlip': result = checkDuplicateSlip(data.slipRef); break;
      case 'savePaymentLog': result = savePaymentLog(data); break;
      case 'getResellerUnpaidAmount': result = getResellerUnpaidAmount(data.userId); break;
      case 'markResellerTransactionsPaid': result = markResellerTransactionsPaid(data); break;

      // Customer Pending Orders
      case 'createPendingOrder': result = createPendingOrder(data); break;
      case 'getPendingOrders': result = getPendingOrders(data.userId); break;
      case 'completePendingOrder': result = completePendingOrder(data); break;
      case 'cancelPendingOrder': result = cancelPendingOrder(data.orderId); break;

      // Renewal
      case 'renewOrder': result = renewOrder(data); break;
      case 'renewOrderByEmail': result = renewOrderByEmail(data); break;

      // Enhanced Dashboard Stats
      case 'getDashboardStatsEnhanced': result = getDashboardStatsEnhanced(data); break;

      // Commission Payment System
      case 'getAdminCommissionStats': result = getAdminCommissionStats(data); break;
      case 'getAllAdminsCommissionStats': result = getAllAdminsCommissionStats(data); break;
      case 'getCommissionLogAll': result = getCommissionLogAll(data); break;
      case 'payCommissionToAdmin': result = payCommissionToAdmin(data); break;
      case 'getCommissionPayments': result = getCommissionPayments(data); break;

      // User Password Reset
      case 'resetUserPassword': result = resetUserPassword(data.userId, data.newPassword); break;

      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =============== AUTH FUNCTIONS ===============

function login(username, password) {
  const sheet = getSheet(SHEETS.ADMINS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username && String(data[i][2]) === String(password)) {
      return {
        success: true,
        user: {
          id: data[i][0],
          username: data[i][1],
          role: data[i][3],
          permissions: tryParseJSON(data[i][4], {}),
          commissions: tryParseJSON(data[i][5], {}),
          name: data[i][8] || '',
          phone: data[i][9] || '',
          lineId: data[i][10] || ''
        }
      };
    }
  }

  return { success: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function resetUserPassword(userId, newPassword) {
  const sheet = getSheet(SHEETS.ADMINS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      sheet.getRange(i + 1, 3).setValue(newPassword);
      return { success: true, message: 'รีเซ็ตรหัสผ่านสำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบผู้ใช้' };
}

function getUser(userId) {
  const sheet = getSheet(SHEETS.ADMINS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      return {
        success: true,
        user: {
          id: data[i][0],
          username: data[i][1],
          role: data[i][3],
          permissions: tryParseJSON(data[i][4], {}),
          commissions: tryParseJSON(data[i][5], {})
        }
      };
    }
  }

  return { success: false, error: 'User not found' };
}

// =============== DASHBOARD FUNCTIONS ===============

function getDashboardStats(userId, role) {
  const ordersSheet = getSheet(SHEETS.ORDERS);
  const ordersData = ordersSheet.getDataRange().getValues();

  const subEmailsSheet = getSheet(SHEETS.SUB_EMAILS);
  const subEmailsData = subEmailsSheet.getDataRange().getValues();

  const mainEmailsSheet = getSheet(SHEETS.MAIN_EMAILS);
  const mainEmailsData = mainEmailsSheet.getDataRange().getValues();

  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  let totalSales = 0;
  let todaySales = 0;
  let monthSales = 0;
  let mySales = 0;
  let myTodaySales = 0;
  let myMonthSales = 0;
  let totalCommission = 0;
  let myCommission = 0;

  // Count orders
  for (let i = 1; i < ordersData.length; i++) {
    const orderDate = new Date(ordersData[i][7]); // SoldAt
    const adminId = ordersData[i][3];
    const commission = ordersData[i][10] || 0;

    totalSales++;
    totalCommission += commission;

    if (orderDate >= startOfToday) {
      todaySales++;
    }
    if (orderDate >= startOfMonth) {
      monthSales++;
    }

    if (adminId === userId) {
      mySales++;
      myCommission += commission;
      if (orderDate >= startOfToday) {
        myTodaySales++;
      }
      if (orderDate >= startOfMonth) {
        myMonthSales++;
      }
    }
  }

  // Count stock
  let stockCount = 0;
  for (let i = 1; i < subEmailsData.length; i++) {
    if (subEmailsData[i][4] === 'stock') {
      stockCount++;
    }
  }

  // Count expiring (next 7 days)
  let expiringCount = 0;
  const sevenDaysLater = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  for (let i = 1; i < ordersData.length; i++) {
    const expiresAt = new Date(ordersData[i][8]);
    const status = ordersData[i][11];
    if (status !== 'expired' && expiresAt <= sevenDaysLater && expiresAt >= now) {
      expiringCount++;
    }
  }

  // Count main email capacity (for owner/super_admin)
  let mainEmailStats = null;
  if (role === 'owner' || role === 'super_admin') {
    // Count sub emails per main email (stock vs sold)
    const subStatsByMain = {};
    for (let i = 1; i < subEmailsData.length; i++) {
      const mainId = subEmailsData[i][1];
      const status = subEmailsData[i][4];
      if (!subStatsByMain[mainId]) {
        subStatsByMain[mainId] = { stock: 0, sold: 0, total: 0 };
      }
      subStatsByMain[mainId].total++;
      if (status === 'stock') {
        subStatsByMain[mainId].stock++;
      } else {
        subStatsByMain[mainId].sold++;
      }
    }

    const mainEmails = [];
    let totalMainEmails = 0;
    let fullMainEmails = 0;
    let totalAvailableSlots = 0;
    let totalStock = 0;
    let totalSold = 0;

    for (let i = 1; i < mainEmailsData.length; i++) {
      const id = mainEmailsData[i][0];
      const email = mainEmailsData[i][1];
      const capacity = mainEmailsData[i][3] || 50;
      const stats = subStatsByMain[id] || { stock: 0, sold: 0, total: 0 };
      const used = stats.total;
      const available = capacity - used;

      totalMainEmails++;
      totalAvailableSlots += available;
      totalStock += stats.stock;
      totalSold += stats.sold;

      if (available === 0) {
        fullMainEmails++;
      }

      mainEmails.push({
        id,
        email,
        capacity,
        used,
        available,
        stock: stats.stock,
        sold: stats.sold,
        isFull: available === 0
      });
    }

    // Sort: full ones first, then by available slots ascending
    mainEmails.sort((a, b) => {
      if (a.isFull && !b.isFull) return -1;
      if (!a.isFull && b.isFull) return 1;
      return a.available - b.available;
    });

    mainEmailStats = {
      totalMainEmails,
      fullMainEmails,
      availableMainEmails: totalMainEmails - fullMainEmails,
      totalAvailableSlots,
      totalStock,
      totalSold,
      mainEmails
    };
  }

  const stats = {
    totalSales,
    todaySales,
    monthSales,
    stockCount,
    expiringCount
  };

  // Add role-specific stats
  if (role === 'owner' || role === 'super_admin') {
    stats.totalCommission = totalCommission;
    stats.mainEmailStats = mainEmailStats;
  }

  if (role !== 'owner') {
    stats.mySales = mySales;
    stats.myTodaySales = myTodaySales;
    stats.myMonthSales = myMonthSales;
    stats.myCommission = myCommission;
  }

  return { success: true, stats };
}

// =============== MAIN EMAIL FUNCTIONS ===============

function getMainEmails(role) {
  const sheet = getSheet(SHEETS.MAIN_EMAILS);
  const data = sheet.getDataRange().getValues();

  const subSheet = getSheet(SHEETS.SUB_EMAILS);
  const subData = subSheet.getDataRange().getValues();

  const emails = [];

  for (let i = 1; i < data.length; i++) {
    // Count used slots
    let usedSlots = 0;
    for (let j = 1; j < subData.length; j++) {
      if (subData[j][1] === data[i][0]) {
        usedSlots++;
      }
    }

    emails.push({
      id: data[i][0],
      email: data[i][1],
      password: (role === 'owner' || role === 'super_admin') ? data[i][2] : '********',
      capacity: data[i][3] || 50,
      usedSlots: usedSlots,
      createdAt: data[i][4]
    });
  }

  return { success: true, emails };
}

function addMainEmail(email, password, createdBy) {
  const sheet = getSheet(SHEETS.MAIN_EMAILS);

  // Check duplicate
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      return { success: false, error: 'หัวเมลนี้มีอยู่แล้ว' };
    }
  }

  const id = generateId();
  sheet.appendRow([id, email, password, 50, new Date().toISOString(), createdBy]);

  return { success: true, id, message: 'เพิ่มหัวเมลสำเร็จ' };
}

function updateMainEmail(id, email, password) {
  const sheet = getSheet(SHEETS.MAIN_EMAILS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      if (email) sheet.getRange(i + 1, 2).setValue(email);
      if (password) sheet.getRange(i + 1, 3).setValue(password);
      return { success: true, message: 'อัพเดทหัวเมลสำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบหัวเมล' };
}

function deleteMainEmail(id) {
  const sheet = getSheet(SHEETS.MAIN_EMAILS);
  const data = sheet.getDataRange().getValues();

  // Check if has sub emails
  const subSheet = getSheet(SHEETS.SUB_EMAILS);
  const subData = subSheet.getDataRange().getValues();

  for (let i = 1; i < subData.length; i++) {
    if (subData[i][1] === id) {
      return { success: false, error: 'ไม่สามารถลบได้ เพราะยังมีซับเมลอยู่' };
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบหัวเมลสำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบหัวเมล' };
}

// =============== SUB EMAIL FUNCTIONS ===============

function getSubEmails(role, status) {
  const sheet = getSheet(SHEETS.SUB_EMAILS);
  const data = sheet.getDataRange().getValues();

  const mainSheet = getSheet(SHEETS.MAIN_EMAILS);
  const mainData = mainSheet.getDataRange().getValues();

  // Create main email lookup
  const mainLookup = {};
  for (let i = 1; i < mainData.length; i++) {
    mainLookup[mainData[i][0]] = {
      email: mainData[i][1],
      password: mainData[i][2]
    };
  }

  const emails = [];

  for (let i = 1; i < data.length; i++) {
    if (status && data[i][4] !== status) continue;

    const mainEmail = mainLookup[data[i][1]] || {};

    emails.push({
      id: data[i][0],
      mainEmailId: data[i][1],
      mainEmail: mainEmail.email || '-',
      mainPassword: (role === 'owner' || role === 'super_admin') ? mainEmail.password : null,
      email: data[i][2],
      password: data[i][3],
      status: data[i][4],
      createdAt: data[i][5]
    });
  }

  return { success: true, emails };
}

function addSubEmails(mainEmailId, emails, createdBy) {
  const sheet = getSheet(SHEETS.SUB_EMAILS);

  // Verify main email exists
  const mainSheet = getSheet(SHEETS.MAIN_EMAILS);
  const mainData = mainSheet.getDataRange().getValues();
  let mainExists = false;
  let capacity = 50;

  for (let i = 1; i < mainData.length; i++) {
    if (mainData[i][0] === mainEmailId) {
      mainExists = true;
      capacity = mainData[i][3] || 50;
      break;
    }
  }

  if (!mainExists) {
    return { success: false, error: 'ไม่พบหัวเมล' };
  }

  // Check capacity
  const subData = sheet.getDataRange().getValues();
  let currentCount = 0;
  for (let i = 1; i < subData.length; i++) {
    if (subData[i][1] === mainEmailId) {
      currentCount++;
    }
  }

  if (currentCount + emails.length > capacity) {
    return { success: false, error: `เกินจำนวนที่รองรับ (เหลือ ${capacity - currentCount} slot)` };
  }

  // Add emails
  let added = 0;
  for (const emailData of emails) {
    const id = generateId();
    sheet.appendRow([
      id,
      mainEmailId,
      emailData.email,
      emailData.password,
      'stock',
      new Date().toISOString(),
      createdBy
    ]);
    added++;
  }

  return { success: true, added, message: `เพิ่ม ${added} ซับเมลสำเร็จ` };
}

function deleteSubEmail(id) {
  const sheet = getSheet(SHEETS.SUB_EMAILS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      // Check if sold
      if (data[i][4] === 'sold') {
        return { success: false, error: 'ไม่สามารถลบซับเมลที่ขายแล้วได้' };
      }
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบซับเมลสำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบซับเมล' };
}

function searchSubEmail(email, role) {
  const sheet = getSheet(SHEETS.SUB_EMAILS);
  const data = sheet.getDataRange().getValues();

  const mainSheet = getSheet(SHEETS.MAIN_EMAILS);
  const mainData = mainSheet.getDataRange().getValues();

  const ordersSheet = getSheet(SHEETS.ORDERS);
  const ordersData = ordersSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toLowerCase().includes(email.toLowerCase())) {
      // Find main email
      let mainEmail = null;
      for (let j = 1; j < mainData.length; j++) {
        if (mainData[j][0] === data[i][1]) {
          mainEmail = {
            email: mainData[j][1],
            password: (role === 'owner' || role === 'super_admin') ? mainData[j][2] : null
          };
          break;
        }
      }

      // Find order info
      let orderInfo = null;
      for (let j = 1; j < ordersData.length; j++) {
        if (ordersData[j][1] === data[i][0]) {
          orderInfo = {
            customerName: ordersData[j][5],
            packageDays: ordersData[j][6],
            soldAt: ordersData[j][7],
            expiresAt: ordersData[j][8],
            soldBy: ordersData[j][4],
            status: ordersData[j][11]
          };
        }
      }

      return {
        success: true,
        result: {
          id: data[i][0],
          email: data[i][2],
          password: data[i][3],
          status: data[i][4],
          mainEmail: mainEmail,
          order: orderInfo
        }
      };
    }
  }

  return { success: false, error: 'ไม่พบซับเมลที่ค้นหา' };
}

// =============== ORDER FUNCTIONS ===============

function createOrder(data) {
  const { subEmailIds, packageDays, customerName, saleType, remark, adminId, adminRole } = data;

  const subSheet = getSheet(SHEETS.SUB_EMAILS);
  const subData = subSheet.getDataRange().getValues();

  const ordersSheet = getSheet(SHEETS.ORDERS);
  const adminsSheet = getSheet(SHEETS.ADMINS);
  const adminsData = adminsSheet.getDataRange().getValues();

  // Get admin info
  let adminName = '';
  let adminCommissions = {};
  for (let i = 1; i < adminsData.length; i++) {
    if (adminsData[i][0] === adminId) {
      adminName = adminsData[i][1];
      adminCommissions = tryParseJSON(adminsData[i][5], {});
      break;
    }
  }

  // Get package info for reseller pricing
  const packagesSheet = getSheet(SHEETS.PACKAGES);
  const packagesData = packagesSheet.getDataRange().getValues();
  let packagePrice = 0;
  for (let i = 1; i < packagesData.length; i++) {
    if (packagesData[i][2] === packageDays && packagesData[i][4]) {
      packagePrice = packagesData[i][3];
      break;
    }
  }

  const now = new Date();
  const expiresAt = new Date(now.getTime() + packageDays * 24 * 60 * 60 * 1000);

  // Generate batch ID for grouping all emails in this purchase
  const batchOrderId = generateId();
  const results = [];

  for (const subEmailId of subEmailIds) {
    // Find and update sub email
    for (let i = 1; i < subData.length; i++) {
      if (subData[i][0] === subEmailId && subData[i][4] === 'stock') {
        // Update status to sold
        subSheet.getRange(i + 1, 5).setValue('sold');

        // Calculate commission
        let commissionAmount = 0;
        if (saleType === 'direct') {
          commissionAmount = adminCommissions[packageDays] || 0;
        }

        // Create order - each item gets unique ID but shares batchOrderId
        const orderId = generateId();
        ordersSheet.appendRow([
          orderId,
          subEmailId,
          subData[i][2], // SubEmail
          adminId,
          adminName,
          customerName,
          packageDays,
          now.toISOString(),
          expiresAt.toISOString(),
          saleType,
          commissionAmount,
          'active',
          remark || '',
          batchOrderId // Column 14: Group multiple items together
        ]);

        // Log commission if direct sale
        if (saleType === 'direct' && commissionAmount > 0) {
          const commissionSheet = getSheet(SHEETS.COMMISSION_LOG);
          commissionSheet.appendRow([
            generateId(),
            adminId,
            adminName,
            orderId,
            subData[i][2],
            packageDays,
            commissionAmount,
            now.toISOString(),
            'sale' // Type column
          ]);
        }

        // Create transaction for reseller
        if (adminRole === 'reseller') {
          const txnSheet = getSheet(SHEETS.TRANSACTIONS);
          txnSheet.appendRow([
            generateId(),
            adminId,
            adminName,
            orderId,
            subData[i][2],
            packageDays,
            packagePrice,
            'pending',
            now.toISOString(),
            ''
          ]);
        }

        results.push({
          orderId,
          email: subData[i][2],
          password: subData[i][3],
          expiresAt: expiresAt.toISOString()
        });

        break;
      }
    }
  }

  if (results.length === 0) {
    return { success: false, error: 'ไม่พบซับเมลที่พร้อมขาย' };
  }

  // Send Telegram notification if configured
  sendExpiryNotification();

  return { success: true, orders: results, message: `สั่งซื้อ ${results.length} รายการสำเร็จ` };
}

function getOrders(userId, role, filter, saleType) {
  const sheet = getSheet(SHEETS.ORDERS);
  const data = sheet.getDataRange().getValues();

  // Also get SubEmails sheet for password lookup
  // Columns: [0]=Id, [1]=MainEmailId, [2]=Email, [3]=Password, [4]=Status
  const subSheet = getSheet(SHEETS.SUB_EMAILS);
  const subData = subSheet.getDataRange().getValues();
  const subEmailMap = {};
  for (let i = 1; i < subData.length; i++) {
    subEmailMap[subData[i][0]] = {
      email: subData[i][2],    // Email column
      password: subData[i][3]  // Password column
    };
  }

  const now = new Date();
  const orders = [];

  for (let i = 1; i < data.length; i++) {
    // Filter by role
    if (role !== 'owner' && role !== 'super_admin' && data[i][3] !== userId) {
      continue;
    }

    // Filter by type
    if (filter === 'expiring') {
      const expiresAt = new Date(data[i][8]);
      const daysUntilExpiry = Math.ceil((expiresAt - now) / (24 * 60 * 60 * 1000));
      if (daysUntilExpiry > 7 || data[i][11] === 'expired') continue;
    }

    // Filter by saleType
    if (saleType && data[i][9] !== saleType) {
      continue;
    }

    // Get password from SubEmails
    const subEmailInfo = subEmailMap[data[i][1]] || {};

    orders.push({
      id: data[i][0],
      subEmailId: data[i][1],
      subEmail: data[i][2],
      password: subEmailInfo.password || '',
      adminId: data[i][3],
      adminName: data[i][4],
      soldBy: data[i][4],
      customerName: data[i][5],
      packageDays: data[i][6],
      soldAt: data[i][7],
      expiresAt: data[i][8],
      saleType: data[i][9],
      commissionAmount: data[i][10],
      status: data[i][11],
      remark: data[i][12],
      orderId: data[i][13] || (data[i][7] + '_' + data[i][4]) // batchOrderId or timestamp+admin combo for old orders
    });
  }

  // Sort by soldAt descending (newest first)
  orders.sort((a, b) => new Date(b.soldAt) - new Date(a.soldAt));

  return { success: true, orders };
}

function extendOrder(orderId, days) {
  const sheet = getSheet(SHEETS.ORDERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      const currentExpiry = new Date(data[i][8]);
      const newExpiry = new Date(currentExpiry.getTime() + days * 24 * 60 * 60 * 1000);

      sheet.getRange(i + 1, 9).setValue(newExpiry.toISOString());
      sheet.getRange(i + 1, 12).setValue('active');

      return { success: true, newExpiresAt: newExpiry.toISOString(), message: `ต่ออายุ ${days} วันสำเร็จ` };
    }
  }

  return { success: false, error: 'ไม่พบ Order' };
}

function markOrderExpired(orderId) {
  const ordersSheet = getSheet(SHEETS.ORDERS);
  const ordersData = ordersSheet.getDataRange().getValues();

  for (let i = 1; i < ordersData.length; i++) {
    if (ordersData[i][0] === orderId) {
      // Mark order as expired
      ordersSheet.getRange(i + 1, 12).setValue('expired');

      // Update sub email status
      const subEmailId = ordersData[i][1];
      const subSheet = getSheet(SHEETS.SUB_EMAILS);
      const subData = subSheet.getDataRange().getValues();

      for (let j = 1; j < subData.length; j++) {
        if (subData[j][0] === subEmailId) {
          // Delete the sub email (free up slot)
          subSheet.deleteRow(j + 1);
          break;
        }
      }

      return { success: true, message: 'ลบเมลและ mark เป็นหมดอายุสำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบ Order' };
}

// =============== EXPIRING FUNCTIONS ===============

function getExpiringEmails(days) {
  const sheet = getSheet(SHEETS.ORDERS);
  const data = sheet.getDataRange().getValues();

  const mainSheet = getSheet(SHEETS.MAIN_EMAILS);
  const mainData = mainSheet.getDataRange().getValues();

  const subSheet = getSheet(SHEETS.SUB_EMAILS);
  const subData = subSheet.getDataRange().getValues();

  const now = new Date();
  const targetDate = new Date(now.getTime() + days * 24 * 60 * 60 * 1000);

  const expiring = {
    today: [],
    soon: [],    // 1-3 days
    upcoming: [] // 4-7 days
  };

  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === 'expired') continue;

    const expiresAt = new Date(data[i][8]);
    const daysUntil = Math.ceil((expiresAt - now) / (24 * 60 * 60 * 1000));

    if (daysUntil > days) continue;

    // Find main email info
    let mainEmail = null;
    const subEmailId = data[i][1];

    for (let j = 1; j < subData.length; j++) {
      if (subData[j][0] === subEmailId) {
        const mainId = subData[j][1];
        for (let k = 1; k < mainData.length; k++) {
          if (mainData[k][0] === mainId) {
            mainEmail = {
              email: mainData[k][1],
              password: mainData[k][2]
            };
            break;
          }
        }
        break;
      }
    }

    const item = {
      id: data[i][0],
      subEmail: data[i][2],
      customerName: data[i][5],
      packageDays: data[i][6],
      expiresAt: data[i][8],
      daysUntil: daysUntil,
      soldBy: data[i][4],
      mainEmail: mainEmail
    };

    if (daysUntil <= 0) {
      expiring.today.push(item);
    } else if (daysUntil <= 3) {
      expiring.soon.push(item);
    } else {
      expiring.upcoming.push(item);
    }
  }

  return { success: true, expiring };
}

// =============== ADMIN FUNCTIONS ===============

function getAdmins() {
  const sheet = getSheet(SHEETS.ADMINS);
  const data = sheet.getDataRange().getValues();

  const admins = [];

  for (let i = 1; i < data.length; i++) {
    admins.push({
      id: data[i][0],
      username: data[i][1],
      role: data[i][3],
      permissions: tryParseJSON(data[i][4], {}),
      commissions: tryParseJSON(data[i][5], {}),
      balance: data[i][6] || 0,
      createdAt: data[i][7]
    });
  }

  return { success: true, admins };
}

function addAdmin(username, password, role, permissions) {
  const sheet = getSheet(SHEETS.ADMINS);

  // Check duplicate
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      return { success: false, error: 'ชื่อผู้ใช้นี้มีอยู่แล้ว' };
    }
  }

  const id = generateId();
  sheet.appendRow([
    id,
    username,
    password,
    role,
    JSON.stringify(permissions || {}),
    JSON.stringify({}),
    0,
    new Date().toISOString()
  ]);

  return { success: true, id, message: 'เพิ่มผู้ใช้สำเร็จ' };
}

function updateAdmin(id, username, password, role, permissions) {
  const sheet = getSheet(SHEETS.ADMINS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      if (username) sheet.getRange(i + 1, 2).setValue(username);
      if (password) sheet.getRange(i + 1, 3).setValue(password);
      if (role) sheet.getRange(i + 1, 4).setValue(role);
      if (permissions !== undefined) sheet.getRange(i + 1, 5).setValue(JSON.stringify(permissions));
      return { success: true, message: 'อัพเดทผู้ใช้สำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบผู้ใช้' };
}

function deleteAdmin(id) {
  const sheet = getSheet(SHEETS.ADMINS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      if (data[i][3] === 'owner') {
        return { success: false, error: 'ไม่สามารถลบ Owner ได้' };
      }
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบผู้ใช้สำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบผู้ใช้' };
}

// =============== PACKAGE FUNCTIONS ===============

function getPackages() {
  const sheet = getSheet(SHEETS.PACKAGES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find column indices dynamically (in case old sheet structure)
  const idCol = headers.indexOf('Id');
  const nameCol = headers.indexOf('Name');
  const daysCol = headers.indexOf('Days');
  const activeCol = headers.indexOf('Active');

  // Support both old (Price) and new (ResellerPrice, CustomerPrice) columns
  let resellerPriceCol = headers.indexOf('ResellerPrice');
  let customerPriceCol = headers.indexOf('CustomerPrice');
  const oldPriceCol = headers.indexOf('Price');

  // If old structure, use Price for both
  if (resellerPriceCol === -1) resellerPriceCol = oldPriceCol;
  if (customerPriceCol === -1) customerPriceCol = oldPriceCol;

  const packages = [];

  for (let i = 1; i < data.length; i++) {
    const resellerPrice = resellerPriceCol >= 0 ? (parseFloat(data[i][resellerPriceCol]) || 0) : 0;
    const customerPrice = customerPriceCol >= 0 ? (parseFloat(data[i][customerPriceCol]) || 0) : 0;

    packages.push({
      id: idCol >= 0 ? data[i][idCol] : data[i][0],
      name: nameCol >= 0 ? data[i][nameCol] : data[i][1],
      days: daysCol >= 0 ? data[i][daysCol] : data[i][2],
      resellerPrice: resellerPrice,
      customerPrice: customerPrice || resellerPrice, // fallback to reseller price
      price: resellerPrice, // for backwards compatibility
      active: activeCol >= 0 ? data[i][activeCol] : true
    });
  }

  return { success: true, packages };
}

function savePackages(packages) {
  const sheet = getSheet(SHEETS.PACKAGES);

  // Clear existing
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  // Add new packages
  for (const pkg of packages) {
    sheet.appendRow([
      pkg.id || generateId(),
      pkg.name,
      pkg.days,
      pkg.resellerPrice || pkg.price || 0,
      pkg.customerPrice || pkg.price || 0,
      pkg.active !== false
    ]);
  }

  return { success: true, message: 'บันทึกแพ็กเกจสำเร็จ' };
}

// =============== COMMISSION FUNCTIONS ===============

function getCommissionSettings() {
  const adminsSheet = getSheet(SHEETS.ADMINS);
  const adminsData = adminsSheet.getDataRange().getValues();

  const packagesSheet = getSheet(SHEETS.PACKAGES);
  const packagesData = packagesSheet.getDataRange().getValues();

  const packages = [];
  for (let i = 1; i < packagesData.length; i++) {
    if (packagesData[i][4]) {
      packages.push({
        days: packagesData[i][2],
        name: packagesData[i][1]
      });
    }
  }

  const admins = [];
  for (let i = 1; i < adminsData.length; i++) {
    if (adminsData[i][3] !== 'owner' && adminsData[i][3] !== 'reseller') {
      admins.push({
        id: adminsData[i][0],
        username: adminsData[i][1],
        role: adminsData[i][3],
        commissions: tryParseJSON(adminsData[i][5], {})
      });
    }
  }

  return { success: true, packages, admins };
}

function updateCommission(adminId, commissions) {
  const sheet = getSheet(SHEETS.ADMINS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === adminId) {
      sheet.getRange(i + 1, 6).setValue(JSON.stringify(commissions));
      return { success: true, message: 'อัพเดทค่าคอมสำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบผู้ใช้' };
}

function getCommissionLog(userId, role, month) {
  const sheet = getSheet(SHEETS.COMMISSION_LOG);
  const data = sheet.getDataRange().getValues();

  const logs = [];
  let total = 0;

  for (let i = 1; i < data.length; i++) {
    // Filter by user if not owner/super_admin
    if (role !== 'owner' && role !== 'super_admin' && data[i][1] !== userId) {
      continue;
    }

    // Filter by month
    if (month) {
      const logMonth = data[i][7].substring(0, 7);
      if (logMonth !== month) continue;
    }

    const amount = data[i][6] || 0;
    total += amount;

    logs.push({
      id: data[i][0],
      adminId: data[i][1],
      adminName: data[i][2],
      orderId: data[i][3],
      subEmail: data[i][4],
      packageDays: data[i][5],
      amount: amount,
      createdAt: data[i][7]
    });
  }

  return { success: true, logs, total };
}

// =============== RESELLER FUNCTIONS ===============

function getResellerBalance(userId) {
  const sheet = getSheet(SHEETS.TRANSACTIONS);
  const data = sheet.getDataRange().getValues();

  let pending = 0;
  let paid = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId) {
      if (data[i][7] === 'pending') {
        pending += data[i][6] || 0;
      } else {
        paid += data[i][6] || 0;
      }
    }
  }

  return { success: true, pending, paid, total: pending + paid };
}

function getResellerTransactions(userId, role) {
  const sheet = getSheet(SHEETS.TRANSACTIONS);
  const data = sheet.getDataRange().getValues();

  const transactions = [];

  for (let i = 1; i < data.length; i++) {
    // Filter by user if not owner/super_admin
    if (role !== 'owner' && role !== 'super_admin' && data[i][1] !== userId) {
      continue;
    }

    transactions.push({
      id: data[i][0],
      resellerId: data[i][1],
      resellerName: data[i][2],
      orderId: data[i][3],
      subEmail: data[i][4],
      packageDays: data[i][5],
      amount: data[i][6],
      status: data[i][7],
      createdAt: data[i][8],
      paidAt: data[i][9]
    });
  }

  return { success: true, transactions };
}

function markTransactionPaid(transactionId) {
  const sheet = getSheet(SHEETS.TRANSACTIONS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === transactionId) {
      sheet.getRange(i + 1, 8).setValue('paid');
      sheet.getRange(i + 1, 10).setValue(new Date().toISOString());
      return { success: true, message: 'บันทึกการชำระเงินสำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบรายการ' };
}

// =============== SETTINGS FUNCTIONS ===============

function getSettings() {
  const sheet = getSheet(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();

  const settings = {};

  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }

  return { success: true, settings };
}

function updateSettings(newSettings) {
  const sheet = getSheet(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();

  for (const key in newSettings) {
    if (key === 'action') continue;

    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(newSettings[key]);
        found = true;
        break;
      }
    }

    if (!found) {
      sheet.appendRow([key, newSettings[key]]);
    }
  }

  return { success: true, message: 'บันทึกการตั้งค่าสำเร็จ' };
}

// =============== TELEGRAM NOTIFICATION ===============

function sendExpiryNotification() {
  const settings = getSettings().settings;

  if (!settings.telegramBotToken || !settings.telegramChatId) {
    return;
  }

  const expiring = getExpiringEmails(parseInt(settings.notifyDaysBefore) || 3).expiring;

  if (expiring.today.length === 0 && expiring.soon.length === 0) {
    return;
  }

  let message = '🔔 <b>ULTRA Stock - แจ้งเตือนเมลใกล้หมดอายุ</b>\n\n';

  if (expiring.today.length > 0) {
    message += `🔴 <b>หมดวันนี้ (${expiring.today.length})</b>\n`;
    expiring.today.forEach(item => {
      message += `• ${item.subEmail}\n`;
    });
    message += '\n';
  }

  if (expiring.soon.length > 0) {
    message += `🟠 <b>หมดใน 1-3 วัน (${expiring.soon.length})</b>\n`;
    expiring.soon.forEach(item => {
      message += `• ${item.subEmail} (${item.daysUntil} วัน)\n`;
    });
  }

  const url = `https://api.telegram.org/bot${settings.telegramBotToken}/sendMessage`;

  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: settings.telegramChatId,
      text: message,
      parse_mode: 'HTML'
    })
  });
}

// Daily trigger for expiry notification
function dailyExpiryCheck() {
  sendExpiryNotification();
}

// Hourly stock report - ตั้ง trigger ให้ทำงานทุกชั่วโมง
function hourlyStockReport() {
  const settings = getSettings().settings;

  if (!settings.telegramBotToken || !settings.telegramChatId) {
    return;
  }

  // Get dashboard stats
  const stats = getDashboardStats('system', 'owner').stats;
  const mes = stats.mainEmailStats;

  if (!mes) return;

  const now = new Date();
  const thaiTime = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');

  let message = `📊 <b>ULTRA Stock Report</b>\n`;
  message += `🕐 ${thaiTime}\n\n`;

  message += `📈 <b>ยอดขาย</b>\n`;
  message += `• ทั้งหมด: ${stats.totalSales || 0}\n`;
  message += `• เดือนนี้: ${stats.monthSales || 0}\n\n`;

  message += `📦 <b>สต็อก</b>\n`;
  message += `• รอขาย: ${mes.totalStock || 0}\n`;
  message += `• ขายแล้ว: ${mes.totalSold || 0}\n`;
  message += `• Slots ว่าง: ${mes.totalAvailableSlots || 0}\n\n`;

  message += `📧 <b>หัวเมล (${mes.totalMainEmails})</b>\n`;
  if (mes.fullMainEmails > 0) {
    message += `🔴 เต็ม: ${mes.fullMainEmails}\n`;
  }

  mes.mainEmails.forEach(m => {
    const status = m.isFull ? '🔴' : m.available <= 10 ? '🟠' : '🟢';
    const emailShort = m.email.split('@')[0];
    message += `${status} ${emailShort}@... : ${m.used}/${m.capacity} (สต็อก ${m.stock}, ขาย ${m.sold}, ว่าง ${m.available})\n`;
  });

  if (stats.expiringCount > 0) {
    message += `\n⏰ <b>ใกล้หมดอายุ:</b> ${stats.expiringCount} รายการ`;
  }

  const url = `https://api.telegram.org/bot${settings.telegramBotToken}/sendMessage`;

  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: settings.telegramChatId,
      text: message,
      parse_mode: 'HTML'
    })
  });
}

// =============== HELPER FUNCTIONS ===============

function tryParseJSON(str, defaultValue) {
  try {
    return JSON.parse(str) || defaultValue;
  } catch (e) {
    return defaultValue;
  }
}

// =============== REGISTRATION ===============

function registerUser(data) {
  const { username, password, name, phone, lineId } = data;

  if (!username || !password) {
    return { success: false, error: 'กรุณากรอก username และ password' };
  }

  const sheet = getSheet(SHEETS.ADMINS);
  const allData = sheet.getDataRange().getValues();

  // Check duplicate username
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][1] === username) {
      return { success: false, error: 'Username นี้ถูกใช้งานแล้ว' };
    }
  }

  const userId = generateId();
  const now = new Date().toISOString();

  // New user gets 'customer' role by default
  sheet.appendRow([
    userId,
    username,
    password,
    'customer',  // Default role
    JSON.stringify({}),  // Permissions
    JSON.stringify({}),  // Commissions
    0,  // Balance
    now,
    name || '',
    phone || '',
    lineId || ''
  ]);

  return {
    success: true,
    message: 'สมัครสมาชิกสำเร็จ',
    user: {
      id: userId,
      username,
      role: 'customer',
      name: name || ''
    }
  };
}

// =============== BANK ACCOUNT ===============

function getBankAccount() {
  const sheet = getSheet(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();

  let bankName = 'ธ.ไทยพาณิชย์';
  let accountNumber = '7122556905';
  let accountName = 'จิณณวัตร ไทยพุทรา';

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'bankName') bankName = data[i][1] || bankName;
    if (data[i][0] === 'bankAccountNumber') accountNumber = data[i][1] || accountNumber;
    if (data[i][0] === 'bankAccountName') accountName = data[i][1] || accountName;
  }

  return {
    success: true,
    bankAccount: { bankName, accountNumber, accountName }
  };
}

function updateBankAccount(data) {
  const { bankName, accountNumber, accountName } = data;
  const sheet = getSheet(SHEETS.SETTINGS);
  const allData = sheet.getDataRange().getValues();

  const updates = {
    'bankName': bankName,
    'bankAccountNumber': accountNumber,
    'bankAccountName': accountName
  };

  for (const [key, value] of Object.entries(updates)) {
    let found = false;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, value]);
    }
  }

  return { success: true, message: 'อัพเดทข้อมูลธนาคารสำเร็จ' };
}

// =============== PAYMENT LOGS (ป้องกันสลิปซ้ำ) ===============

function checkDuplicateSlip(slipRef) {
  if (!slipRef) {
    return { success: true, isDuplicate: false };
  }

  const sheet = getSheet(SHEETS.PAYMENT_LOGS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === slipRef) {  // SlipRef column
      return { success: true, isDuplicate: true };
    }
  }

  return { success: true, isDuplicate: false };
}

function savePaymentLog(data) {
  const { userId, userName, userRole, amount, slipRef, status, transactionIds } = data;

  const sheet = getSheet(SHEETS.PAYMENT_LOGS);
  const logId = generateId();
  const now = new Date().toISOString();

  sheet.appendRow([
    logId,
    userId,
    userName,
    userRole,
    amount,
    slipRef,
    status,
    now,
    JSON.stringify(transactionIds || [])
  ]);

  return { success: true, logId };
}

// =============== RESELLER PAYMENT ===============

function getResellerUnpaidAmount(userId) {
  const sheet = getSheet(SHEETS.TRANSACTIONS);
  const data = sheet.getDataRange().getValues();

  let unpaidAmount = 0;
  const unpaidTransactions = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId && data[i][7] !== 'paid') {
      unpaidAmount += Number(data[i][6]) || 0;
      unpaidTransactions.push({
        id: data[i][0],
        orderSubEmail: data[i][4],
        packageDays: data[i][5],
        amount: data[i][6],
        createdAt: data[i][8]
      });
    }
  }

  return {
    success: true,
    unpaidAmount,
    unpaidTransactions
  };
}

function markResellerTransactionsPaid(data) {
  const { transactionIds, slipRef, paidAt } = data;

  const sheet = getSheet(SHEETS.TRANSACTIONS);
  const allData = sheet.getDataRange().getValues();

  let updatedCount = 0;

  for (let i = 1; i < allData.length; i++) {
    if (transactionIds.includes(allData[i][0])) {
      sheet.getRange(i + 1, 8).setValue('paid');  // Status
      sheet.getRange(i + 1, 10).setValue(paidAt); // PaidAt
      sheet.getRange(i + 1, 11).setValue(slipRef); // SlipRef
      updatedCount++;
    }
  }

  return { success: true, updatedCount };
}

// =============== CUSTOMER PENDING ORDERS ===============

function createPendingOrder(data) {
  const { userId, userName, subEmailIds, packageDays, amount } = data;

  const sheet = getSheet(SHEETS.PENDING_ORDERS);
  const orderId = generateId();
  const now = new Date();
  const expiresAt = new Date(now.getTime() + 30 * 60 * 1000); // 30 minutes to pay

  sheet.appendRow([
    orderId,
    userId,
    userName,
    JSON.stringify(subEmailIds),
    packageDays,
    amount,
    'pending',
    now.toISOString(),
    expiresAt.toISOString(),
    '',
    ''
  ]);

  return {
    success: true,
    orderId,
    amount,
    expiresAt: expiresAt.toISOString()
  };
}

function getPendingOrders(userId) {
  const sheet = getSheet(SHEETS.PENDING_ORDERS);
  const data = sheet.getDataRange().getValues();

  const orders = [];
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId && data[i][6] === 'pending') {
      const expiresAt = new Date(data[i][8]);
      if (expiresAt > now) {
        orders.push({
          id: data[i][0],
          subEmailIds: tryParseJSON(data[i][3], []),
          packageDays: data[i][4],
          amount: data[i][5],
          createdAt: data[i][7],
          expiresAt: data[i][8]
        });
      }
    }
  }

  return { success: true, orders };
}

function completePendingOrder(data) {
  const { orderId, slipRef } = data;

  const pendingSheet = getSheet(SHEETS.PENDING_ORDERS);
  const pendingData = pendingSheet.getDataRange().getValues();

  let pendingOrder = null;
  let pendingRow = -1;

  for (let i = 1; i < pendingData.length; i++) {
    if (pendingData[i][0] === orderId && pendingData[i][6] === 'pending') {
      pendingOrder = {
        userId: pendingData[i][1],
        userName: pendingData[i][2],
        subEmailIds: tryParseJSON(pendingData[i][3], []),
        packageDays: pendingData[i][4],
        amount: pendingData[i][5]
      };
      pendingRow = i + 1;
      break;
    }
  }

  if (!pendingOrder) {
    return { success: false, error: 'ไม่พบคำสั่งซื้อหรือหมดอายุแล้ว' };
  }

  // Create actual orders
  const orderResult = createOrder({
    subEmailIds: pendingOrder.subEmailIds,
    packageDays: pendingOrder.packageDays,
    customerName: pendingOrder.userName,
    saleType: 'customer',
    adminId: pendingOrder.userId,
    adminRole: 'customer'
  });

  if (!orderResult.success) {
    return orderResult;
  }

  // Mark pending order as completed
  pendingSheet.getRange(pendingRow, 7).setValue('completed');
  pendingSheet.getRange(pendingRow, 10).setValue(new Date().toISOString());
  pendingSheet.getRange(pendingRow, 11).setValue(slipRef);

  return {
    success: true,
    message: 'ชำระเงินสำเร็จ',
    orders: orderResult.orders
  };
}

function cancelPendingOrder(orderId) {
  const sheet = getSheet(SHEETS.PENDING_ORDERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      sheet.getRange(i + 1, 7).setValue('cancelled');
      return { success: true };
    }
  }

  return { success: false, error: 'ไม่พบคำสั่งซื้อ' };
}

// =============== RENEWAL ===============

function renewOrder(data) {
  const { orderId, packageDays, userId, userRole, slipRef } = data;

  const ordersSheet = getSheet(SHEETS.ORDERS);
  const ordersData = ordersSheet.getDataRange().getValues();

  let orderRow = -1;
  let orderInfo = null;

  for (let i = 1; i < ordersData.length; i++) {
    if (ordersData[i][0] === orderId) {
      orderRow = i + 1;
      orderInfo = {
        subEmailId: ordersData[i][1],
        subEmail: ordersData[i][2],
        adminId: ordersData[i][3],
        adminName: ordersData[i][4],
        currentExpiry: new Date(ordersData[i][8])
      };
      break;
    }
  }

  if (!orderInfo) {
    return { success: false, error: 'ไม่พบ Order' };
  }

  // Get package price
  const packagesSheet = getSheet(SHEETS.PACKAGES);
  const packagesData = packagesSheet.getDataRange().getValues();
  let resellerPrice = 0;
  let customerPrice = 0;

  for (let i = 1; i < packagesData.length; i++) {
    if (packagesData[i][2] === packageDays && packagesData[i][5]) {
      resellerPrice = packagesData[i][3];
      customerPrice = packagesData[i][4];
      break;
    }
  }

  // Calculate new expiry
  const now = new Date();
  const baseDate = orderInfo.currentExpiry > now ? orderInfo.currentExpiry : now;
  const newExpiry = new Date(baseDate.getTime() + packageDays * 24 * 60 * 60 * 1000);

  // Update order expiry
  ordersSheet.getRange(orderRow, 9).setValue(newExpiry.toISOString());
  ordersSheet.getRange(orderRow, 12).setValue('active');

  // Handle based on role
  if (userRole === 'admin' || userRole === 'super_admin' || userRole === 'owner') {
    // Admin renewal - 50% commission
    const adminsSheet = getSheet(SHEETS.ADMINS);
    const adminsData = adminsSheet.getDataRange().getValues();

    let adminCommissions = {};
    for (let i = 1; i < adminsData.length; i++) {
      if (adminsData[i][0] === userId) {
        adminCommissions = tryParseJSON(adminsData[i][5], {});
        break;
      }
    }

    const baseCommission = adminCommissions[packageDays] || 0;
    const renewalCommission = Math.floor(baseCommission * 0.5); // 50%

    if (renewalCommission > 0) {
      const commissionSheet = getSheet(SHEETS.COMMISSION_LOG);
      commissionSheet.appendRow([
        generateId(),
        userId,
        orderInfo.adminName,
        orderId,
        orderInfo.subEmail,
        packageDays,
        renewalCommission,
        now.toISOString(),
        'renewal'
      ]);
    }

    return {
      success: true,
      message: `ต่ออายุสำเร็จ ${packageDays} วัน (ค่าคอม ฿${renewalCommission})`,
      newExpiresAt: newExpiry.toISOString(),
      commission: renewalCommission
    };

  } else if (userRole === 'reseller') {
    // Reseller - add to balance
    const txnSheet = getSheet(SHEETS.TRANSACTIONS);
    txnSheet.appendRow([
      generateId(),
      userId,
      orderInfo.adminName,
      orderId,
      orderInfo.subEmail,
      packageDays,
      resellerPrice,
      'pending',
      now.toISOString(),
      '',
      '',
      'renewal'
    ]);

    return {
      success: true,
      message: `ต่ออายุสำเร็จ ${packageDays} วัน (ยอดค้าง +฿${resellerPrice})`,
      newExpiresAt: newExpiry.toISOString(),
      amount: resellerPrice
    };

  } else {
    // Customer - should have paid already (slipRef provided)
    return {
      success: true,
      message: `ต่ออายุสำเร็จ ${packageDays} วัน`,
      newExpiresAt: newExpiry.toISOString(),
      amount: customerPrice
    };
  }
}

// =============== ENHANCED DASHBOARD STATS ===============

function getDashboardStatsEnhanced(data) {
  const { userId, role, startDate, endDate } = data;

  const ordersSheet = getSheet(SHEETS.ORDERS);
  const ordersData = ordersSheet.getDataRange().getValues();

  const adminsSheet = getSheet(SHEETS.ADMINS);
  const adminsData = adminsSheet.getDataRange().getValues();

  const subSheet = getSheet(SHEETS.SUB_EMAILS);
  const subData = subSheet.getDataRange().getValues();

  const txnSheet = getSheet(SHEETS.TRANSACTIONS);
  const txnData = txnSheet.getDataRange().getValues();

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const thisMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  // Parse date filters
  const filterStart = startDate ? new Date(startDate) : null;
  const filterEnd = endDate ? new Date(endDate) : null;

  // Build admin role map
  const adminRoles = {};
  for (let i = 1; i < adminsData.length; i++) {
    adminRoles[adminsData[i][0]] = adminsData[i][3]; // userId -> role
  }

  // Count stats
  let totalSales = 0;
  let todaySales = 0;
  let monthSales = 0;
  let adminSales = 0;
  let resellerSales = 0;
  let customerSales = 0;
  let filteredSales = 0;
  let expiringCount = 0;
  let stockCount = 0;

  // Revenue stats
  let adminRevenue = 0;
  let resellerRevenue = 0;
  let customerRevenue = 0;

  // Get packages for price lookup
  const packagesSheet = getSheet(SHEETS.PACKAGES);
  const packagesData = packagesSheet.getDataRange().getValues();
  const packagePrices = {};
  for (let i = 1; i < packagesData.length; i++) {
    packagePrices[packagesData[i][2]] = {
      reseller: packagesData[i][3],
      customer: packagesData[i][4]
    };
  }

  for (let i = 1; i < ordersData.length; i++) {
    const soldAt = new Date(ordersData[i][7]);
    const expiresAt = new Date(ordersData[i][8]);
    const adminId = ordersData[i][3];
    const saleType = ordersData[i][9];
    const status = ordersData[i][11];
    const packageDays = ordersData[i][6];

    totalSales++;

    if (soldAt >= today) todaySales++;
    if (soldAt >= thisMonth) monthSales++;

    // Filter by date range
    if (filterStart && filterEnd) {
      if (soldAt >= filterStart && soldAt <= filterEnd) {
        filteredSales++;
      }
    }

    // Count by seller type
    const sellerRole = adminRoles[adminId] || 'unknown';
    const prices = packagePrices[packageDays] || { reseller: 0, customer: 0 };

    if (sellerRole === 'customer') {
      customerSales++;
      customerRevenue += prices.customer;
    } else if (sellerRole === 'reseller') {
      resellerSales++;
      resellerRevenue += prices.reseller;
    } else if (saleType === 'direct') {
      adminSales++;
      // Admin sales don't have direct revenue (commission-based)
    }

    // Expiring count
    if (status === 'active') {
      const daysUntil = Math.ceil((expiresAt - now) / (24 * 60 * 60 * 1000));
      if (daysUntil <= 7 && daysUntil > 0) expiringCount++;
    }
  }

  // Stock count
  for (let i = 1; i < subData.length; i++) {
    if (subData[i][4] === 'stock') stockCount++;
  }

  // Unpaid amounts
  let resellerUnpaid = 0;
  let resellerUnpaidCount = 0;
  const unpaidByUser = {};

  for (let i = 1; i < txnData.length; i++) {
    if (txnData[i][7] !== 'paid') {
      const amount = Number(txnData[i][6]) || 0;
      resellerUnpaid += amount;
      resellerUnpaidCount++;

      const uid = txnData[i][1];
      if (!unpaidByUser[uid]) unpaidByUser[uid] = 0;
      unpaidByUser[uid] += amount;
    }
  }

  // User counts
  let adminCount = 0;
  let resellerCount = 0;
  let customerCount = 0;
  let newUsersThisMonth = 0;

  for (let i = 1; i < adminsData.length; i++) {
    const r = adminsData[i][3];
    const createdAt = new Date(adminsData[i][7]);

    if (r === 'admin' || r === 'super_admin') adminCount++;
    else if (r === 'reseller') resellerCount++;
    else if (r === 'customer') customerCount++;

    if (createdAt >= thisMonth) newUsersThisMonth++;
  }

  return {
    success: true,
    stats: {
      // Overview
      totalSales,
      todaySales,
      monthSales,
      filteredSales,
      stockCount,
      expiringCount,

      // By seller type
      adminSales,
      resellerSales,
      customerSales,

      // Revenue
      resellerRevenue,
      customerRevenue,
      totalRevenue: resellerRevenue + customerRevenue,

      // Unpaid
      resellerUnpaid,
      resellerUnpaidCount,
      unpaidUserCount: Object.keys(unpaidByUser).length,

      // Users
      adminCount,
      resellerCount,
      customerCount,
      totalUsers: adminCount + resellerCount + customerCount,
      newUsersThisMonth
    }
  };
}

// ===============================================
// RENEWAL BY EMAIL
// ===============================================

function renewOrderByEmail(data) {
  const { email, packageDays, userId, userRole, slipRef, paidAmount } = data;

  const ordersSheet = getSheet(SHEETS.ORDERS);
  const ordersData = ordersSheet.getDataRange().getValues();
  const ordersHeaders = ordersData[0];

  // Find column indices
  const subEmailCol = ordersHeaders.indexOf('SubEmail');
  const expiresAtCol = ordersHeaders.indexOf('ExpiresAt');
  const statusCol = ordersHeaders.indexOf('Status');

  // Find the order for this email
  let orderRow = -1;
  for (let i = 1; i < ordersData.length; i++) {
    if (ordersData[i][subEmailCol] === email) {
      orderRow = i + 1;
      break;
    }
  }

  if (orderRow === -1) {
    return { success: false, error: 'ไม่พบอีเมลนี้ในระบบ' };
  }

  // Get current expiry
  const currentExpiry = new Date(ordersData[orderRow - 1][expiresAtCol]);
  const now = new Date();

  // Calculate new expiry (from current expiry if still valid, or from now if expired)
  const baseDate = currentExpiry > now ? currentExpiry : now;
  const newExpiry = new Date(baseDate);
  newExpiry.setDate(newExpiry.getDate() + packageDays);

  // Update the order
  ordersSheet.getRange(orderRow, expiresAtCol + 1).setValue(newExpiry.toISOString());
  ordersSheet.getRange(orderRow, statusCol + 1).setValue('active');

  // Get package price for commission calculation
  const packagesSheet = getSheet(SHEETS.PACKAGES);
  const packagesData = packagesSheet.getDataRange().getValues();
  let packagePrice = 0;

  for (let i = 1; i < packagesData.length; i++) {
    if (packagesData[i][2] === packageDays) { // Days column
      packagePrice = userRole === 'customer' ? packagesData[i][4] : packagesData[i][3]; // CustomerPrice or ResellerPrice
      break;
    }
  }

  // Handle commission for renewal (50% for admin)
  if (userRole === 'admin' || userRole === 'super_admin') {
    const user = findUserById(userId);
    if (user) {
      const commissions = JSON.parse(user.commissions || '{}');
      const commissionRate = (commissions.rate || 0) / 100;
      const commissionAmount = Math.round(packagePrice * commissionRate * 0.5); // 50% for renewal

      if (commissionAmount > 0) {
        const commissionSheet = getSheet(SHEETS.COMMISSION_LOG);
        commissionSheet.appendRow([
          generateId(),
          userId,
          user.username,
          '', // OrderId
          email,
          packageDays,
          commissionAmount,
          new Date().toISOString(),
          'renewal'
        ]);
      }
    }
  }

  // Handle Reseller - add to balance
  if (userRole === 'reseller') {
    const transactionsSheet = getSheet(SHEETS.TRANSACTIONS);
    transactionsSheet.appendRow([
      generateId(),
      userId,
      findUserById(userId)?.username || 'Unknown',
      '', // OrderId
      email,
      packageDays,
      packagePrice,
      'unpaid',
      new Date().toISOString(),
      '',
      slipRef || '',
      'renewal'
    ]);
  }

  return {
    success: true,
    message: `ต่ออายุสำเร็จ! หมดอายุใหม่: ${formatDateThai(newExpiry)}`,
    newExpiry: newExpiry.toISOString()
  };
}

// ===============================================
// COMMISSION PAYMENT SYSTEM
// ===============================================

// Get Admin Commission Stats (includes paid/unpaid/bonus)
function getAdminCommissionStats(data) {
  const { userId, month } = data;

  const user = findUserById(userId);
  if (!user) {
    return { success: false, error: 'ไม่พบผู้ใช้' };
  }

  // Get all commission logs for this admin
  const commissionSheet = getSheet(SHEETS.COMMISSION_LOG);
  const commissionData = commissionSheet.getDataRange().getValues();
  const commissionHeaders = commissionData[0];

  const adminIdCol = commissionHeaders.indexOf('AdminId');
  const amountCol = commissionHeaders.indexOf('Amount');
  const createdAtCol = commissionHeaders.indexOf('CreatedAt');
  const typeCol = commissionHeaders.indexOf('Type');

  // Filter for this admin and month
  const targetMonth = month || new Date().toISOString().substring(0, 7);
  let totalEarned = 0;
  let salesCommission = 0;
  let renewalCommission = 0;

  for (let i = 1; i < commissionData.length; i++) {
    if (commissionData[i][adminIdCol] === userId) {
      const createdAt = commissionData[i][createdAtCol];
      if (createdAt && createdAt.substring(0, 7) === targetMonth) {
        const amount = parseFloat(commissionData[i][amountCol]) || 0;
        totalEarned += amount;

        const type = commissionData[i][typeCol];
        if (type === 'renewal') {
          renewalCommission += amount;
        } else {
          salesCommission += amount;
        }
      }
    }
  }

  // Get payments for this admin
  const paymentsSheet = getSheet(SHEETS.COMMISSION_PAYMENTS);
  const paymentsData = paymentsSheet.getDataRange().getValues();
  const paymentHeaders = paymentsData[0];

  const paymentAdminIdCol = paymentHeaders.indexOf('AdminId');
  const paymentAmountCol = paymentHeaders.indexOf('Amount');
  const paymentMonthCol = paymentHeaders.indexOf('Month');
  const paymentNoteCol = paymentHeaders.indexOf('Note');
  const paymentPaidAtCol = paymentHeaders.indexOf('PaidAt');

  let totalPaid = 0;
  let payments = [];

  for (let i = 1; i < paymentsData.length; i++) {
    if (paymentsData[i][paymentAdminIdCol] === userId) {
      const paymentMonth = paymentsData[i][paymentMonthCol];
      if (paymentMonth === targetMonth) {
        const amount = parseFloat(paymentsData[i][paymentAmountCol]) || 0;
        totalPaid += amount;
        payments.push({
          amount,
          note: paymentsData[i][paymentNoteCol],
          paidAt: paymentsData[i][paymentPaidAtCol]
        });
      }
    }
  }

  // Calculate
  const unpaid = Math.max(0, totalEarned - totalPaid);
  const bonus = Math.max(0, totalPaid - totalEarned);

  return {
    success: true,
    stats: {
      adminId: userId,
      adminName: user.username,
      month: targetMonth,
      totalEarned,
      salesCommission,
      renewalCommission,
      totalPaid,
      unpaid,
      bonus,
      payments
    }
  };
}

// Get All Admins Commission Stats (Owner only)
function getAllAdminsCommissionStats(data) {
  const { month } = data;
  const targetMonth = month || new Date().toISOString().substring(0, 7);

  const adminsSheet = getSheet(SHEETS.ADMINS);
  const adminsData = adminsSheet.getDataRange().getValues();

  const admins = [];

  for (let i = 1; i < adminsData.length; i++) {
    const role = adminsData[i][3]; // Role column
    if (role === 'admin' || role === 'super_admin' || role === 'owner') {
      const userId = adminsData[i][0];
      const stats = getAdminCommissionStats({ userId, month: targetMonth });
      if (stats.success) {
        admins.push(stats.stats);
      }
    }
  }

  // Calculate totals
  const totals = {
    totalEarned: admins.reduce((sum, a) => sum + a.totalEarned, 0),
    totalPaid: admins.reduce((sum, a) => sum + a.totalPaid, 0),
    totalUnpaid: admins.reduce((sum, a) => sum + a.unpaid, 0),
    totalBonus: admins.reduce((sum, a) => sum + a.bonus, 0)
  };

  return {
    success: true,
    month: targetMonth,
    admins,
    totals
  };
}

// Get All Commission Logs (for Owner to see raw data)
function getCommissionLogAll(data) {
  const { month } = data;
  const targetMonth = month || new Date().toISOString().substring(0, 7);

  const sheet = getSheet(SHEETS.COMMISSION_LOG);
  const sheetData = sheet.getDataRange().getValues();
  const headers = sheetData[0];

  const idCol = headers.indexOf('Id');
  const adminIdCol = headers.indexOf('AdminId');
  const adminNameCol = headers.indexOf('AdminName');
  const orderIdCol = headers.indexOf('OrderId');
  const subEmailCol = headers.indexOf('SubEmail');
  const packageDaysCol = headers.indexOf('PackageDays');
  const amountCol = headers.indexOf('Amount');
  const createdAtCol = headers.indexOf('CreatedAt');
  const typeCol = headers.indexOf('Type');

  const logs = [];

  for (let i = 1; i < sheetData.length; i++) {
    const createdAt = sheetData[i][createdAtCol];
    if (!createdAt) continue;

    // Filter by month
    const logMonth = createdAt.substring(0, 7);
    if (logMonth !== targetMonth) continue;

    logs.push({
      id: sheetData[i][idCol],
      adminId: sheetData[i][adminIdCol],
      adminName: sheetData[i][adminNameCol],
      orderId: sheetData[i][orderIdCol],
      subEmail: sheetData[i][subEmailCol],
      packageDays: sheetData[i][packageDaysCol],
      amount: parseFloat(sheetData[i][amountCol]) || 0,
      createdAt: createdAt,
      type: sheetData[i][typeCol] || 'sale'
    });
  }

  // Sort by date descending
  logs.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  return {
    success: true,
    month: targetMonth,
    logs,
    total: logs.reduce((sum, l) => sum + l.amount, 0)
  };
}

// Pay Commission to Admin
function payCommissionToAdmin(data) {
  const { adminId, amount, note, paidBy, paidAt } = data;

  const user = findUserById(adminId);
  if (!user) {
    return { success: false, error: 'ไม่พบผู้ใช้' };
  }

  // Get current month
  const currentMonth = new Date().toISOString().substring(0, 7);

  const paymentsSheet = getSheet(SHEETS.COMMISSION_PAYMENTS);
  paymentsSheet.appendRow([
    generateId(),
    adminId,
    user.username,
    amount,
    note || '',
    paidBy,
    paidAt || new Date().toISOString(),
    currentMonth
  ]);

  return {
    success: true,
    message: `จ่ายค่าคอม ฿${amount.toLocaleString()} ให้ ${user.username} สำเร็จ!`
  };
}

// Get Commission Payments History
function getCommissionPayments(data) {
  const { userId } = data;

  const paymentsSheet = getSheet(SHEETS.COMMISSION_PAYMENTS);
  const paymentsData = paymentsSheet.getDataRange().getValues();
  const headers = paymentsData[0];

  const adminIdCol = headers.indexOf('AdminId');
  const amountCol = headers.indexOf('Amount');
  const noteCol = headers.indexOf('Note');
  const paidAtCol = headers.indexOf('PaidAt');
  const monthCol = headers.indexOf('Month');

  const payments = [];

  for (let i = 1; i < paymentsData.length; i++) {
    if (paymentsData[i][adminIdCol] === userId) {
      payments.push({
        amount: parseFloat(paymentsData[i][amountCol]) || 0,
        note: paymentsData[i][noteCol],
        paidAt: paymentsData[i][paidAtCol],
        month: paymentsData[i][monthCol]
      });
    }
  }

  // Sort by date descending
  payments.sort((a, b) => new Date(b.paidAt) - new Date(a.paidAt));

  return {
    success: true,
    payments
  };
}

// Helper function to format date in Thai
function formatDateThai(date) {
  const d = new Date(date);
  return d.toLocaleDateString('th-TH', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}
