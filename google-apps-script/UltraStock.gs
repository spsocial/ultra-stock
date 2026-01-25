// ===============================================
// ULTRA Stock - Google Apps Script Database
// ===============================================

// Sheet Names
const SHEETS = {
  ADMINS: 'Admins',
  MAIN_EMAILS: 'MainEmails',
  SUB_EMAILS: 'SubEmails',
  ORDERS: 'Orders',
  TRANSACTIONS: 'Transactions',
  COMMISSION_LOG: 'CommissionLog',
  PACKAGES: 'Packages',
  SETTINGS: 'Settings'
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
    [SHEETS.ADMINS]: ['Id', 'Username', 'Password', 'Role', 'Permissions', 'Commissions', 'Balance', 'CreatedAt'],
    [SHEETS.MAIN_EMAILS]: ['Id', 'Email', 'Password', 'Capacity', 'CreatedAt', 'CreatedBy'],
    [SHEETS.SUB_EMAILS]: ['Id', 'MainEmailId', 'Email', 'Password', 'Status', 'CreatedAt', 'CreatedBy'],
    [SHEETS.ORDERS]: ['Id', 'SubEmailId', 'SubEmail', 'AdminId', 'AdminName', 'CustomerName', 'PackageDays', 'SoldAt', 'ExpiresAt', 'SaleType', 'CommissionAmount', 'Status', 'Remark'],
    [SHEETS.TRANSACTIONS]: ['Id', 'ResellerId', 'ResellerName', 'OrderId', 'SubEmail', 'PackageDays', 'Amount', 'Status', 'CreatedAt', 'PaidAt'],
    [SHEETS.COMMISSION_LOG]: ['Id', 'AdminId', 'AdminName', 'OrderId', 'SubEmail', 'PackageDays', 'Amount', 'CreatedAt'],
    [SHEETS.PACKAGES]: ['Id', 'Name', 'Days', 'Price', 'Active'],
    [SHEETS.SETTINGS]: ['Key', 'Value']
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

  // Initialize default packages
  if (sheetName === SHEETS.PACKAGES) {
    const defaultPackages = [
      [generateId(), '7 วัน', 7, 500, true],
      [generateId(), '20 วัน', 20, 500, true],
      [generateId(), '30 วัน', 30, 500, true]
    ];
    defaultPackages.forEach(pkg => sheet.appendRow(pkg));
  }

  // Initialize default settings
  if (sheetName === SHEETS.SETTINGS) {
    const defaultSettings = [
      ['telegramBotToken', ''],
      ['telegramChatId', ''],
      ['notifyDaysBefore', '3']
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
      case 'getUser': result = getUser(data.userId); break;

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
      case 'getOrders': result = getOrders(data.userId, data.role, data.filter); break;
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
          commissions: tryParseJSON(data[i][5], {})
        }
      };
    }
  }

  return { success: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
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

  let totalSales = 0;
  let monthSales = 0;
  let mySales = 0;
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

    if (orderDate >= startOfMonth) {
      monthSales++;
    }

    if (adminId === userId) {
      mySales++;
      myCommission += commission;
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

        // Create order
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
          remark || ''
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
            now.toISOString()
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

function getOrders(userId, role, filter) {
  const sheet = getSheet(SHEETS.ORDERS);
  const data = sheet.getDataRange().getValues();

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

    orders.push({
      id: data[i][0],
      subEmailId: data[i][1],
      subEmail: data[i][2],
      adminId: data[i][3],
      adminName: data[i][4],
      customerName: data[i][5],
      packageDays: data[i][6],
      soldAt: data[i][7],
      expiresAt: data[i][8],
      saleType: data[i][9],
      commissionAmount: data[i][10],
      status: data[i][11],
      remark: data[i][12]
    });
  }

  // Sort by expiresAt
  orders.sort((a, b) => new Date(a.expiresAt) - new Date(b.expiresAt));

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

  const packages = [];

  for (let i = 1; i < data.length; i++) {
    packages.push({
      id: data[i][0],
      name: data[i][1],
      days: data[i][2],
      price: data[i][3],
      active: data[i][4]
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
      pkg.price,
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
