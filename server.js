const express = require('express');
const jwt = require('jsonwebtoken');
const cors = require('cors');
const fetch = require('node-fetch');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Configuration
const JWT_SECRET = process.env.JWT_SECRET || 'ultra-stock-secret-key-change-in-production';
const GOOGLE_SCRIPT_URL = process.env.GOOGLE_SCRIPT_URL || 'YOUR_GOOGLE_SCRIPT_URL_HERE';
const EASYSLIP_API_KEY = process.env.EASYSLIP_API_KEY || 'bf4c6851-0df7-4020-8488-cfe5a7f4f276';
const PORT = process.env.PORT || 3000;

// Helper: Call Google Apps Script
async function callGoogleScript(action, data = {}) {
  try {
    const response = await fetch(GOOGLE_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action, ...data })
    });
    return await response.json();
  } catch (error) {
    console.error('Google Script Error:', error);
    return { success: false, error: error.message };
  }
}

// Helper: Verify Slip with EasySlip API
async function verifySlipWithEasySlip(base64Image) {
  try {
    const response = await fetch('https://developer.easyslip.com/api/v1/verify', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${EASYSLIP_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ image: base64Image })
    });

    const data = await response.json();

    if (data.status === 200 && data.data) {
      return { success: true, data: data.data };
    }

    return { success: false, error: data.message || 'ไม่สามารถตรวจสอบสลิปได้' };
  } catch (error) {
    console.error('EasySlip API error:', error);
    return { success: false, error: error.message || 'เกิดข้อผิดพลาดในการเชื่อมต่อ EasySlip' };
  }
}

// Middleware: Verify JWT Token
function authenticateToken(req, res, next) {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ success: false, error: 'Token required' });
  }

  jwt.verify(token, JWT_SECRET, (err, user) => {
    if (err) {
      return res.status(403).json({ success: false, error: 'Invalid token' });
    }
    req.user = user;
    next();
  });
}

// Middleware: Check Role
function requireRole(...roles) {
  return (req, res, next) => {
    if (!roles.includes(req.user.role)) {
      return res.status(403).json({ success: false, error: 'Permission denied' });
    }
    next();
  };
}

// ============ AUTH ROUTES ============

// Login
app.post('/api/login', async (req, res) => {
  const { username, password } = req.body;

  if (!username || !password) {
    return res.status(400).json({ success: false, error: 'Username and password required' });
  }

  const result = await callGoogleScript('login', { username, password });

  if (result.success) {
    const token = jwt.sign(
      {
        id: result.user.id,
        username: result.user.username,
        role: result.user.role,
        permissions: result.user.permissions || {}
      },
      JWT_SECRET,
      { expiresIn: '7d' }
    );

    res.json({
      success: true,
      token,
      user: {
        id: result.user.id,
        username: result.user.username,
        role: result.user.role,
        permissions: result.user.permissions || {}
      }
    });
  } else {
    res.status(401).json(result);
  }
});

// Register (สมัครสมาชิก - ได้ role customer อัตโนมัติ)
app.post('/api/register', async (req, res) => {
  const { username, password, name, phone, lineId } = req.body;

  if (!username || !password) {
    return res.status(400).json({ success: false, error: 'กรุณากรอก Username และ Password' });
  }

  const result = await callGoogleScript('register', { username, password, name, phone, lineId });

  if (result.success) {
    // Auto login after register
    const token = jwt.sign(
      {
        id: result.user.id,
        username: result.user.username,
        role: result.user.role
      },
      JWT_SECRET,
      { expiresIn: '7d' }
    );

    res.json({
      success: true,
      message: result.message,
      token,
      user: result.user
    });
  } else {
    res.status(400).json(result);
  }
});

// Check Session
app.get('/api/check-session', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getUser', { userId: req.user.id });
  if (result.success) {
    res.json({
      success: true,
      user: {
        id: result.user.id,
        username: result.user.username,
        role: result.user.role,
        permissions: result.user.permissions || {}
      }
    });
  } else {
    res.status(401).json({ success: false, error: 'Session expired' });
  }
});

// ============ DASHBOARD ROUTES ============

// Get Dashboard Stats
app.get('/api/dashboard', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getDashboardStats', {
    userId: req.user.id,
    role: req.user.role
  });
  res.json(result);
});

// ============ MAIN EMAIL ROUTES (Owner & Super Admin) ============

// Get Main Emails
app.get('/api/main-emails', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const result = await callGoogleScript('getMainEmails', { role: req.user.role });
  res.json(result);
});

// Add Main Email
app.post('/api/main-emails', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const { email, password } = req.body;
  const result = await callGoogleScript('addMainEmail', {
    email,
    password,
    createdBy: req.user.id
  });
  res.json(result);
});

// Update Main Email
app.put('/api/main-emails/:id', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const { email, password } = req.body;
  const result = await callGoogleScript('updateMainEmail', {
    id: req.params.id,
    email,
    password
  });
  res.json(result);
});

// Delete Main Email
app.delete('/api/main-emails/:id', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const result = await callGoogleScript('deleteMainEmail', { id: req.params.id });
  res.json(result);
});

// ============ SUB EMAIL ROUTES ============

// Get Sub Emails (Stock)
app.get('/api/sub-emails', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getSubEmails', {
    role: req.user.role,
    status: req.query.status || 'stock'
  });
  res.json(result);
});

// Add Sub Emails (Stock) - Owner & Super Admin
app.post('/api/sub-emails', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const { mainEmailId, emails } = req.body; // emails = [{email, password}, ...]
  const result = await callGoogleScript('addSubEmails', {
    mainEmailId,
    emails,
    createdBy: req.user.id
  });
  res.json(result);
});

// Delete Sub Email
app.delete('/api/sub-emails/:id', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const result = await callGoogleScript('deleteSubEmail', { id: req.params.id });
  res.json(result);
});

// Search Sub Email
app.get('/api/sub-emails/search', authenticateToken, async (req, res) => {
  const { email } = req.query;
  const result = await callGoogleScript('searchSubEmail', {
    email,
    role: req.user.role
  });
  res.json(result);
});

// ============ ORDER ROUTES ============

// Buy Sub Email
app.post('/api/orders', authenticateToken, async (req, res) => {
  const { subEmailIds, packageDays, customerName, saleType, remark } = req.body;
  // saleType: 'direct' (นับคอม) or 'stock' (ไม่นับคอม)

  // Check permission for stock type
  if (saleType === 'stock') {
    const canBuyStock = req.user.role === 'owner' || req.user.role === 'super_admin' ||
      (req.user.permissions && req.user.permissions.canBuyWithoutCommission);
    if (!canBuyStock) {
      return res.status(403).json({ success: false, error: 'ไม่มีสิทธิ์ซื้อแบบไม่นับคอม' });
    }
  }

  const result = await callGoogleScript('createOrder', {
    subEmailIds,
    packageDays,
    customerName,
    saleType,
    remark,
    adminId: req.user.id,
    adminRole: req.user.role
  });
  res.json(result);
});

// Get Orders
app.get('/api/orders', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getOrders', {
    userId: req.user.id,
    role: req.user.role,
    filter: req.query.filter, // 'all', 'mine', 'expiring'
    saleType: req.query.saleType // 'direct', 'stock', or undefined for all
  });
  res.json(result);
});

// Extend Order (ต่ออายุ)
app.put('/api/orders/:id/extend', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const { days } = req.body;
  const result = await callGoogleScript('extendOrder', {
    orderId: req.params.id,
    days
  });
  res.json(result);
});

// Mark Order as Expired (ลบเมลแล้ว)
app.put('/api/orders/:id/expire', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const result = await callGoogleScript('markOrderExpired', { orderId: req.params.id });
  res.json(result);
});

// ============ EXPIRING EMAILS ROUTES ============

// Get Expiring Emails
app.get('/api/expiring', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const result = await callGoogleScript('getExpiringEmails', {
    days: parseInt(req.query.days) || 7
  });
  res.json(result);
});

// ============ ADMIN MANAGEMENT ROUTES (Owner Only) ============

// Get All Admins
app.get('/api/admins', authenticateToken, requireRole('owner'), async (req, res) => {
  const result = await callGoogleScript('getAdmins');
  res.json(result);
});

// Add Admin
app.post('/api/admins', authenticateToken, requireRole('owner'), async (req, res) => {
  const { username, password, role, permissions } = req.body;
  const result = await callGoogleScript('addAdmin', {
    username,
    password,
    role,
    permissions
  });
  res.json(result);
});

// Update Admin
app.put('/api/admins/:id', authenticateToken, requireRole('owner'), async (req, res) => {
  const { username, password, role, permissions } = req.body;
  const result = await callGoogleScript('updateAdmin', {
    id: req.params.id,
    username,
    password,
    role,
    permissions
  });
  res.json(result);
});

// Delete Admin
app.delete('/api/admins/:id', authenticateToken, requireRole('owner'), async (req, res) => {
  const result = await callGoogleScript('deleteAdmin', { id: req.params.id });
  res.json(result);
});

// ============ PACKAGE SETTINGS (Owner Only) ============

// Get Packages
app.get('/api/packages', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getPackages');
  res.json(result);
});

// Save Packages
app.post('/api/packages', authenticateToken, requireRole('owner'), async (req, res) => {
  const { packages } = req.body; // [{days, price, name}, ...]
  const result = await callGoogleScript('savePackages', { packages });
  res.json(result);
});

// ============ COMMISSION SETTINGS (Owner Only) ============

// Get Commission Settings
app.get('/api/commissions', authenticateToken, requireRole('owner'), async (req, res) => {
  const result = await callGoogleScript('getCommissionSettings');
  res.json(result);
});

// Update Commission for Admin
app.put('/api/commissions/:adminId', authenticateToken, requireRole('owner'), async (req, res) => {
  const { commissions } = req.body; // {packageDays: amount, ...}
  const result = await callGoogleScript('updateCommission', {
    adminId: req.params.adminId,
    commissions
  });
  res.json(result);
});

// ============ COMMISSION LOG ============

// Get Commission Log
app.get('/api/commission-log', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getCommissionLog', {
    userId: req.user.id,
    role: req.user.role,
    month: req.query.month // optional: YYYY-MM
  });
  res.json(result);
});

// ============ RESELLER TRANSACTIONS (Buy Now Pay Later) ============

// Get Reseller Balance
app.get('/api/reseller/balance', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getResellerBalance', {
    userId: req.user.id
  });
  res.json(result);
});

// Get Reseller Transactions
app.get('/api/reseller/transactions', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getResellerTransactions', {
    userId: req.user.id,
    role: req.user.role
  });
  res.json(result);
});

// Mark Transaction as Paid (Owner/Super Admin)
app.put('/api/reseller/transactions/:id/paid', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const result = await callGoogleScript('markTransactionPaid', {
    transactionId: req.params.id
  });
  res.json(result);
});

// ============ SETTINGS ============

// Get Settings
app.get('/api/settings', authenticateToken, requireRole('owner'), async (req, res) => {
  const result = await callGoogleScript('getSettings');
  res.json(result);
});

// Update Settings
app.put('/api/settings', authenticateToken, requireRole('owner'), async (req, res) => {
  const result = await callGoogleScript('updateSettings', req.body);
  res.json(result);
});

// ============ TELEGRAM NOTIFICATION ============

// Test Telegram
app.post('/api/telegram/test', authenticateToken, requireRole('owner'), async (req, res) => {
  const { botToken, chatId } = req.body;
  try {
    const response = await fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: chatId,
        text: '🔔 ทดสอบการแจ้งเตือนจาก ULTRA Stock\n\nหากเห็นข้อความนี้ แสดงว่าตั้งค่าสำเร็จ!',
        parse_mode: 'HTML'
      })
    });
    const data = await response.json();
    if (data.ok) {
      res.json({ success: true, message: 'ส่งข้อความทดสอบสำเร็จ' });
    } else {
      res.json({ success: false, error: data.description });
    }
  } catch (error) {
    res.json({ success: false, error: error.message });
  }
});

// Send Stock Report to Telegram
app.post('/api/send-stock-report', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const { mainEmailStats, dashboardStats } = req.body;

  // Get settings from Google Script
  const settingsResult = await callGoogleScript('getSettings');
  if (!settingsResult.success) {
    return res.json({ success: false, error: 'ไม่สามารถโหลดการตั้งค่าได้' });
  }

  const { telegramBotToken, telegramChatId } = settingsResult.settings;
  if (!telegramBotToken || !telegramChatId) {
    return res.json({ success: false, error: 'ยังไม่ได้ตั้งค่า Telegram' });
  }

  // Build message
  const mes = mainEmailStats;
  const now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });

  let message = `📊 <b>ULTRA Stock Report</b>\n`;
  message += `🕐 ${now}\n\n`;

  message += `📈 <b>ยอดขาย</b>\n`;
  message += `• วันนี้: ${dashboardStats.todaySales || 0}\n`;
  message += `• เดือนนี้: ${dashboardStats.monthSales || 0}\n`;
  message += `• ทั้งหมด: ${dashboardStats.totalSales || 0}\n\n`;

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
    message += `${status} ${m.email.split('@')[0]}@... : ${m.used}/${m.capacity} (ว่าง ${m.available})\n`;
  });

  if (dashboardStats.expiringCount > 0) {
    message += `\n⏰ <b>ใกล้หมดอายุ:</b> ${dashboardStats.expiringCount} รายการ`;
  }

  try {
    const response = await fetch(`https://api.telegram.org/bot${telegramBotToken}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: telegramChatId,
        text: message,
        parse_mode: 'HTML'
      })
    });
    const data = await response.json();
    if (data.ok) {
      res.json({ success: true, message: 'ส่งรายงานสำเร็จ' });
    } else {
      res.json({ success: false, error: data.description });
    }
  } catch (error) {
    res.json({ success: false, error: error.message });
  }
});

// ============ BANK ACCOUNT ROUTES ============

// Get Bank Account (Public - ใช้ตอนแสดงข้อมูลโอนเงิน)
app.get('/api/bank-account', async (req, res) => {
  const result = await callGoogleScript('getBankAccount');
  res.json(result);
});

// Update Bank Account (Owner only)
app.put('/api/bank-account', authenticateToken, requireRole('owner'), async (req, res) => {
  const { bankName, accountNumber, accountName } = req.body;
  const result = await callGoogleScript('updateBankAccount', { bankName, accountNumber, accountName });
  res.json(result);
});

// ============ PAYMENT ROUTES ============

// Reseller: Get Unpaid Amount
app.get('/api/reseller/unpaid', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getResellerUnpaidAmount', { userId: req.user.id });
  res.json(result);
});

// Reseller: Submit Payment (แนบสลิป)
app.post('/api/reseller/submit-payment', authenticateToken, async (req, res) => {
  const { slipImage } = req.body;

  if (!slipImage) {
    return res.status(400).json({ success: false, error: 'กรุณาแนบสลิปโอนเงิน' });
  }

  // 1. Get unpaid amount
  const unpaidResult = await callGoogleScript('getResellerUnpaidAmount', { userId: req.user.id });
  if (!unpaidResult.success || unpaidResult.unpaidAmount === 0) {
    return res.status(400).json({ success: false, error: 'ไม่มียอดค้างชำระ' });
  }

  const unpaidAmount = unpaidResult.unpaidAmount;

  // 2. Verify slip with EasySlip
  const easySlipResult = await verifySlipWithEasySlip(slipImage);
  if (!easySlipResult.success) {
    return res.status(400).json({ success: false, error: easySlipResult.error || 'ไม่สามารถอ่านสลิปได้' });
  }

  const slipData = easySlipResult.data;
  const slipRef = slipData.transRef || '';
  const slipAmount = slipData.amount?.amount || 0;

  // 3. Check duplicate slip
  if (slipRef) {
    const dupCheck = await callGoogleScript('checkDuplicateSlip', { slipRef });
    if (dupCheck.isDuplicate) {
      return res.status(400).json({ success: false, error: 'สลิปนี้เคยใช้แล้ว! กรุณาใช้สลิปใหม่' });
    }
  }

  // 4. Check receiver account
  const bankResult = await callGoogleScript('getBankAccount');
  const ourAccount = bankResult.bankAccount?.accountNumber?.replace(/-/g, '') || '';
  if (ourAccount && slipData.receiver?.account?.value) {
    const receiverAccount = slipData.receiver.account.value.replace(/-/g, '');
    if (!receiverAccount.includes(ourAccount) && !ourAccount.includes(receiverAccount)) {
      return res.status(400).json({ success: false, error: 'บัญชีปลายทางไม่ตรง!' });
    }
  }

  // 5. Check slip date (< 24 hours)
  if (slipData.transTimestamp) {
    const slipDate = new Date(slipData.transTimestamp);
    const hoursDiff = (Date.now() - slipDate) / (1000 * 60 * 60);
    if (hoursDiff > 24) {
      return res.status(400).json({ success: false, error: 'สลิปนี้เก่าเกิน 24 ชั่วโมง!' });
    }
  }

  // 6. Check amount
  if (slipAmount !== unpaidAmount) {
    return res.status(400).json({
      success: false,
      error: `ยอดไม่ตรง! ค้างชำระ ฿${unpaidAmount.toLocaleString()} แต่สลิป ฿${slipAmount.toLocaleString()}`
    });
  }

  // 7. Mark transactions as paid
  const transactionIds = unpaidResult.unpaidTransactions.map(t => t.id);
  await callGoogleScript('markResellerTransactionsPaid', {
    transactionIds,
    slipRef,
    paidAt: new Date().toISOString()
  });

  // 8. Save payment log
  await callGoogleScript('savePaymentLog', {
    userId: req.user.id,
    userName: req.user.username,
    userRole: req.user.role,
    amount: slipAmount,
    slipRef,
    status: 'approved',
    transactionIds
  });

  res.json({
    success: true,
    message: `ชำระเงินสำเร็จ! ยอด ฿${slipAmount.toLocaleString()}`,
    paidAmount: slipAmount
  });
});

// Customer: Create Pending Order
app.post('/api/customer/pending-order', authenticateToken, async (req, res) => {
  const { subEmailIds, packageDays, amount } = req.body;

  const result = await callGoogleScript('createPendingOrder', {
    userId: req.user.id,
    userName: req.user.username,
    subEmailIds,
    packageDays,
    amount
  });

  res.json(result);
});

// Customer: Get Pending Orders
app.get('/api/customer/pending-orders', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('getPendingOrders', { userId: req.user.id });
  res.json(result);
});

// Customer: Pay Pending Order (แนบสลิป)
app.post('/api/customer/pay-order', authenticateToken, async (req, res) => {
  const { orderId, slipImage } = req.body;

  if (!slipImage) {
    return res.status(400).json({ success: false, error: 'กรุณาแนบสลิปโอนเงิน' });
  }

  // Get pending order to verify amount
  const pendingResult = await callGoogleScript('getPendingOrders', { userId: req.user.id });
  const order = pendingResult.orders?.find(o => o.id === orderId);

  if (!order) {
    return res.status(400).json({ success: false, error: 'ไม่พบคำสั่งซื้อหรือหมดอายุแล้ว' });
  }

  // Verify slip
  const easySlipResult = await verifySlipWithEasySlip(slipImage);
  if (!easySlipResult.success) {
    return res.status(400).json({ success: false, error: easySlipResult.error });
  }

  const slipData = easySlipResult.data;
  const slipRef = slipData.transRef || '';
  const slipAmount = slipData.amount?.amount || 0;

  // Check duplicate
  if (slipRef) {
    const dupCheck = await callGoogleScript('checkDuplicateSlip', { slipRef });
    if (dupCheck.isDuplicate) {
      return res.status(400).json({ success: false, error: 'สลิปนี้เคยใช้แล้ว!' });
    }
  }

  // Check receiver account
  const bankResult = await callGoogleScript('getBankAccount');
  const ourAccount = bankResult.bankAccount?.accountNumber?.replace(/-/g, '') || '';
  if (ourAccount && slipData.receiver?.account?.value) {
    const receiverAccount = slipData.receiver.account.value.replace(/-/g, '');
    if (!receiverAccount.includes(ourAccount) && !ourAccount.includes(receiverAccount)) {
      return res.status(400).json({ success: false, error: 'บัญชีปลายทางไม่ตรง!' });
    }
  }

  // Check amount
  if (slipAmount !== order.amount) {
    return res.status(400).json({
      success: false,
      error: `ยอดไม่ตรง! ต้องโอน ฿${order.amount.toLocaleString()} แต่สลิป ฿${slipAmount.toLocaleString()}`
    });
  }

  // Complete order
  const result = await callGoogleScript('completePendingOrder', { orderId, slipRef });

  if (result.success) {
    // Save payment log
    await callGoogleScript('savePaymentLog', {
      userId: req.user.id,
      userName: req.user.username,
      userRole: 'customer',
      amount: slipAmount,
      slipRef,
      status: 'approved',
      transactionIds: [orderId]
    });
  }

  res.json(result);
});

// Cancel Pending Order
app.delete('/api/customer/pending-order/:id', authenticateToken, async (req, res) => {
  const result = await callGoogleScript('cancelPendingOrder', { orderId: req.params.id });
  res.json(result);
});

// ============ RENEWAL ROUTES ============

// Renew Order by ID (old endpoint)
app.post('/api/orders/:id/renew', authenticateToken, async (req, res) => {
  const { packageDays, slipRef } = req.body;

  const result = await callGoogleScript('renewOrder', {
    orderId: req.params.id,
    packageDays,
    userId: req.user.id,
    userRole: req.user.role,
    slipRef
  });

  res.json(result);
});

// Renew Order by Email (for Admin/Reseller)
app.post('/api/orders/renew', authenticateToken, async (req, res) => {
  const { email, packageDays } = req.body;

  const result = await callGoogleScript('renewOrderByEmail', {
    email,
    packageDays,
    userId: req.user.id,
    userRole: req.user.role
  });

  res.json(result);
});

// Renew Order with Payment (for Customer - EasySlip verification)
app.post('/api/orders/renew-with-payment', authenticateToken, async (req, res) => {
  const { email, packageDays, amount, slipImage } = req.body;

  if (!slipImage) {
    return res.status(400).json({ success: false, error: 'กรุณาแนบรูปสลิป' });
  }

  // Verify slip with EasySlip
  const slipResult = await verifySlipWithEasySlip(slipImage);
  if (!slipResult.success) {
    return res.json(slipResult);
  }

  // Verify amount
  if (slipResult.amount < amount) {
    return res.json({
      success: false,
      error: `ยอดเงินไม่ตรง: โอน ฿${slipResult.amount.toLocaleString()} แต่ต้องจ่าย ฿${amount.toLocaleString()}`
    });
  }

  // Process renewal
  const result = await callGoogleScript('renewOrderByEmail', {
    email,
    packageDays,
    userId: req.user.id,
    userRole: req.user.role,
    slipRef: slipResult.transRef,
    paidAmount: slipResult.amount
  });

  // Log payment
  if (result.success) {
    await callGoogleScript('savePaymentLog', {
      userId: req.user.id,
      type: 'renewal',
      amount: slipResult.amount,
      slipRef: slipResult.transRef,
      description: `ต่ออายุ ${email} - ${packageDays} วัน`
    });
  }

  res.json(result);
});

// ============ ENHANCED DASHBOARD ROUTES ============

// Get Enhanced Stats (with date filter)
app.get('/api/dashboard-enhanced', authenticateToken, requireRole('owner', 'super_admin'), async (req, res) => {
  const { startDate, endDate } = req.query;

  const result = await callGoogleScript('getDashboardStatsEnhanced', {
    userId: req.user.id,
    role: req.user.role,
    startDate,
    endDate
  });

  res.json(result);
});

// ============ ADMIN: RESET PASSWORD ============

app.post('/api/users/:id/reset-password', authenticateToken, requireRole('owner', 'super_admin', 'admin'), async (req, res) => {
  const { newPassword } = req.body;

  if (!newPassword) {
    return res.status(400).json({ success: false, error: 'กรุณากรอกรหัสผ่านใหม่' });
  }

  const result = await callGoogleScript('resetUserPassword', {
    userId: req.params.id,
    newPassword
  });

  res.json(result);
});

// ============ COMMISSION PAYMENT SYSTEM ============

// Get Admin Commission Stats (for Owner to view all admins, or admin to view self)
app.get('/api/admin-commission/:userId', authenticateToken, async (req, res) => {
  const targetUserId = req.params.userId;

  // Only owner can view other users' commission, others can only view their own
  if (req.user.role !== 'owner' && req.user.id !== targetUserId) {
    return res.status(403).json({ success: false, error: 'ไม่มีสิทธิ์ดูข้อมูลนี้' });
  }

  const result = await callGoogleScript('getAdminCommissionStats', {
    userId: targetUserId,
    month: req.query.month
  });

  res.json(result);
});

// Get All Admins Commission Summary (Owner only)
app.get('/api/all-admins-commission', authenticateToken, requireRole('owner'), async (req, res) => {
  const result = await callGoogleScript('getAllAdminsCommissionStats', {
    month: req.query.month
  });

  res.json(result);
});

// Get Commission Log (Owner only)
app.get('/api/commission-log', authenticateToken, requireRole('owner'), async (req, res) => {
  const result = await callGoogleScript('getCommissionLogAll', {
    month: req.query.month
  });

  res.json(result);
});

// Pay Commission to Admin (Owner only)
app.post('/api/pay-commission', authenticateToken, requireRole('owner'), async (req, res) => {
  const { adminId, amount, note } = req.body;

  if (!adminId || !amount) {
    return res.status(400).json({ success: false, error: 'กรุณากรอกข้อมูลให้ครบ' });
  }

  const result = await callGoogleScript('payCommissionToAdmin', {
    adminId,
    amount: parseFloat(amount),
    note: note || '',
    paidBy: req.user.id,
    paidAt: new Date().toISOString()
  });

  res.json(result);
});

// Get Commission Payment History
app.get('/api/commission-payments/:userId', authenticateToken, async (req, res) => {
  const targetUserId = req.params.userId;

  // Only owner can view other users' payments, others can only view their own
  if (req.user.role !== 'owner' && req.user.id !== targetUserId) {
    return res.status(403).json({ success: false, error: 'ไม่มีสิทธิ์ดูข้อมูลนี้' });
  }

  const result = await callGoogleScript('getCommissionPayments', {
    userId: targetUserId
  });

  res.json(result);
});

// ============ STATIC PAGES ============

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

// Start Server
app.listen(PORT, () => {
  console.log(`🚀 ULTRA Stock Server running on port ${PORT}`);
  console.log(`📁 Open: http://localhost:${PORT}`);
});
