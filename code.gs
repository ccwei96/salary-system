/***********************
 * 薪资管理系统 - 后端代码
 * 修复版本 v2.0
 ***********************/

/***********************
 * 配置
 ***********************/
const SHEET_USERS = 'Users';
const SHEET_STAFF = 'Staff';       // 员工表（书记管理）
const SHEET_MANAGERS = 'Managers'; // 主管表（Admin管理）
const SHEET_SALARY = 'SalaryRecords';
const SHEET_PAYMENTS = 'PaymentLog';

// 银行变更提醒天数（超过此天数后提醒自动消失）
const BANK_CHANGE_ALERT_DAYS = 14;

// 状态：DRAFT(草稿) -> SUBMITTED(已提交) -> PAID(已发放)

/***********************
 * 数据验证工具
 ***********************/

/**
 * 验证员工/主管数据
 */
function validateStaffData_(staff) {
  const errors = [];
  
  if (!staff.staffName || String(staff.staffName).trim() === '') {
    errors.push('姓名不能为空');
  } else if (String(staff.staffName).trim().length > 50) {
    errors.push('姓名不能超过50个字符');
  }
  
  if (staff.salary !== undefined && staff.salary !== '') {
    const salary = Number(staff.salary);
    if (isNaN(salary)) {
      errors.push('工资必须是数字');
    } else if (salary < 0) {
      errors.push('工资不能为负数');
    } else if (salary > 1000000) {
      errors.push('工资数额异常，请检查');
    }
  }
  
  if (staff.bankAccount && String(staff.bankAccount).trim() !== '') {
    const account = String(staff.bankAccount).trim();
    if (!/^[0-9\-\s]{5,30}$/.test(account.replace(/\s/g, ''))) {
      errors.push('银行账号格式不正确（应为5-30位数字）');
    }
  }
  
  if (staff.paymentMethod && !['BANK', 'CASH'].includes(staff.paymentMethod)) {
    errors.push('付款方式无效，只能是 BANK 或 CASH');
  }
  
  if (staff.joinDate && !isValidDate_(staff.joinDate)) {
    errors.push('入职日期格式不正确');
  }
  if (staff.leaveDate && !isValidDate_(staff.leaveDate)) {
    errors.push('离职日期格式不正确');
  }
  
  if (staff.totalDebt !== undefined && staff.totalDebt !== '') {
    const debt = Number(staff.totalDebt);
    if (isNaN(debt) || debt < 0) {
      errors.push('欠款金额必须是非负数');
    }
  }
  if (staff.monthlyDeduction !== undefined && staff.monthlyDeduction !== '') {
    const ded = Number(staff.monthlyDeduction);
    if (isNaN(ded) || ded < 0) {
      errors.push('月扣款金额必须是非负数');
    }
  }
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * 验证工资记录数据
 */
function validateSalaryRecord_(record) {
  const errors = [];
  
  if (!record.month || !record.month.match(/^\d{4}-\d{2}$/)) {
    errors.push('月份格式不正确，应为 YYYY-MM（如 2025-01）');
  } else {
    const [year, mon] = record.month.split('-').map(Number);
    if (year < 2020 || year > 2100 || mon < 1 || mon > 12) {
      errors.push('月份数值不合理');
    }
  }
  
  if (!record.date) {
    errors.push('日期不能为空');
  } else if (!isValidDate_(record.date)) {
    errors.push('日期格式不正确');
  }
  
  if (!record.staffName || String(record.staffName).trim() === '') {
    errors.push('请选择员工/主管');
  }
  
  if (record.basicSalary === undefined || record.basicSalary === '') {
    errors.push('基本工资不能为空');
  } else {
    const basic = Number(record.basicSalary);
    if (isNaN(basic)) {
      errors.push('基本工资必须是数字');
    } else if (basic < 0) {
      errors.push('基本工资不能为负数');
    } else if (basic > 1000000) {
      errors.push('基本工资数额异常，请检查');
    }
  }
  
  if (record.deduction !== undefined && record.deduction !== '') {
    const ded = Number(record.deduction);
    if (isNaN(ded)) {
      errors.push('扣款必须是数字');
    } else if (ded < 0) {
      errors.push('扣款不能为负数');
    } else if (ded > Number(record.basicSalary || 0)) {
      errors.push('扣款不能超过基本工资');
    }
  }
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * 验证用户账号数据
 */
function validateUserData_(user) {
  const errors = [];
  
  if (!user.username || String(user.username).trim() === '') {
    errors.push('用户名不能为空');
  } else {
    const username = String(user.username).trim();
    if (username.length < 3) {
      errors.push('用户名至少3个字符');
    } else if (username.length > 20) {
      errors.push('用户名不能超过20个字符');
    } else if (!/^[a-zA-Z0-9_]+$/.test(username)) {
      errors.push('用户名只能包含字母、数字和下划线');
    }
  }
  
  if (!user.password || String(user.password).trim() === '') {
    errors.push('密码不能为空');
  } else if (String(user.password).length < 6) {
    errors.push('密码至少6个字符');
  }
  
  if (!user.role || !['ADMIN', 'SECRETARY', 'ACCOUNTANT'].includes(user.role)) {
    errors.push('角色无效');
  }
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * 验证日期格式
 */
function isValidDate_(dateStr) {
  if (!dateStr) return false;
  const date = new Date(dateStr);
  return date instanceof Date && !isNaN(date);
}

/**
 * 验证月份格式并返回解析结果
 */
function parseMonth_(monthStr) {
  if (!monthStr || !monthStr.match(/^\d{4}-\d{2}$/)) {
    return { valid: false };
  }
  const [year, mon] = monthStr.split('-').map(Number);
  if (year < 2020 || year > 2100 || mon < 1 || mon > 12) {
    return { valid: false };
  }
  return { valid: true, year: year, month: mon };
}

/**
 * 获取上个月份
 */
function getPreviousMonth_(monthStr) {
  const parsed = parseMonth_(monthStr);
  if (!parsed.valid) return '';
  
  const year = parsed.year;
  const mon = parsed.month;
  
  if (mon === 1) {
    return (year - 1) + '-12';
  } else {
    return year + '-' + String(mon - 1).padStart(2, '0');
  }
}

/**
 * 清理和标准化输入数据
 */
function sanitizeInput_(value, type) {
  if (value === undefined || value === null) return '';
  
  const str = String(value).trim();
  
  switch (type) {
    case 'string':
      return str.replace(/<[^>]*>/g, '').replace(/[<>'"&]/g, '').substring(0, 200);
    case 'number':
      const num = Number(str);
      return isNaN(num) ? 0 : num;
    case 'date':
      return isValidDate_(str) ? str : '';
    case 'name':
      return str.replace(/<[^>]*>/g, '').replace(/[<>'"&;]/g, '').substring(0, 50);
    default:
      return str;
  }
}

/**
 * 将 Date 对象或字符串转换为 YYYY-MM 格式
 */
function formatMonthValue_(value) {
  if (!value) return '';
  
  if (value instanceof Date) {
    return value.getFullYear() + '-' + String(value.getMonth() + 1).padStart(2, '0');
  }
  
  const str = String(value).trim();
  
  if (/^\d{4}-\d{2}$/.test(str)) {
    return str;
  }
  
  const date = new Date(str);
  if (!isNaN(date.getTime())) {
    return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
  }
  
  return str;
}

/**
 * 辅助函数：安全格式化日期
 */
function formatDateSafe_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

/**
 * 辅助函数：获取当前月份
 */
function getCurrentMonth_() {
  const now = new Date();
  return now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
}

/***********************
 * 人员管理统一模块
 ***********************/

const PERSON_CONFIG = {
  STAFF: {
    sheetName: SHEET_STAFF,
    label: '员工',
    hasIsManagerCol: true,
    allowedRoles: ['SECRETARY', 'ADMIN'],
    managerOnly: false
  },
  MANAGER: {
    sheetName: SHEET_MANAGERS,
    label: '主管',
    hasIsManagerCol: false,
    allowedRoles: ['ADMIN'],
    managerOnly: true
  }
};

/**
 * 获取人员列表（统一方法）
 */
function getPersonList_(type, options) {
  options = options || {};
  const config = PERSON_CONFIG[type];
  if (!config) return [];
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const now = new Date();
  const alertCutoffDate = new Date(now.getTime() - BANK_CHANGE_ALERT_DAYS * 24 * 60 * 60 * 1000);
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[cols['StaffName']];
    if (!name) continue;
    
    const status = row[cols['Status']] || 'ACTIVE';
    
    if (!options.includeLeft && status === 'LEFT') continue;
    
    let bankChanged = false;
    let changeNote = '';
    if (cols['LastBankUpdate'] !== undefined && row[cols['LastBankUpdate']]) {
      const updateDate = new Date(row[cols['LastBankUpdate']]);
      if (updateDate > alertCutoffDate) {
        bankChanged = true;
        changeNote = row[cols['BankChangeNote']] || '';
      }
    }
    
    const totalDebt = Number(row[cols['TotalDebt']]) || 0;
    const debtPaid = Number(row[cols['DebtPaid']]) || 0;
    
    function safeDate(val) {
      if (!val) return '';
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return String(val);
    }
    
    result.push({
      rowIndex: i + 1,
      type: type,
      staffName: String(name),
      salary: Number(row[cols['Salary']]) || 0,
      companyName: String(row[cols['CompanyName']] || ''),
      bankHolder: String(row[cols['BankHolder']] || ''),
      bankType: String(row[cols['BankType']] || ''),
      bankAccount: String(row[cols['BankAccount']] || ''),
      paymentMethod: String(row[cols['PaymentMethod']] || 'BANK'),
      isManager: type === 'MANAGER' || (cols['IsManager'] !== undefined && row[cols['IsManager']] === 'YES'),
      joinDate: safeDate(row[cols['JoinDate']]),
      leaveDate: safeDate(row[cols['LeaveDate']]),
      totalDebt: totalDebt,
      monthlyDeduction: Number(row[cols['MonthlyDeduction']]) || 0,
      debtPaid: debtPaid,
      debtRemaining: totalDebt - debtPaid,
      debtReason: String(row[cols['DebtReason']] || ''),
      bankChanged: bankChanged,
      bankChangeNote: String(changeNote),
      status: String(status)
    });
  }
  
  return result;
}

/**
 * 根据姓名查找人员
 */
function findPerson_(name, type) {
  const cleanName = sanitizeInput_(name, 'name');
  if (!cleanName) return null;
  
  if (type === 'ALL') {
    let person = findPersonInSheet_(cleanName, 'STAFF');
    if (!person) person = findPersonInSheet_(cleanName, 'MANAGER');
    return person;
  }
  return findPersonInSheet_(cleanName, type);
}

/**
 * 在指定表中查找人员
 */
function findPersonInSheet_(name, type) {
  const config = PERSON_CONFIG[type];
  if (!config) return null;
  
  const cleanName = sanitizeInput_(name, 'name');
  if (!cleanName) return null;
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][cols['StaffName']] === cleanName) {
      const row = data[i];
      const totalDebt = Number(row[cols['TotalDebt']]) || 0;
      const debtPaid = Number(row[cols['DebtPaid']]) || 0;
      
      return {
        rowIndex: i + 1,
        sheet: sheet,
        sheetName: config.sheetName,
        type: type,
        cols: cols,
        data: row,
        staffName: cleanName,
        salary: Number(row[cols['Salary']]) || 0,
        companyName: row[cols['CompanyName']] || '',
        bankHolder: row[cols['BankHolder']] || '',
        bankType: row[cols['BankType']] || '',
        bankAccount: row[cols['BankAccount']] || '',
        paymentMethod: row[cols['PaymentMethod']] || 'BANK',
        isManager: type === 'MANAGER',
        joinDate: row[cols['JoinDate']] || '',
        leaveDate: row[cols['LeaveDate']] || '',
        totalDebt: totalDebt,
        monthlyDeduction: Number(row[cols['MonthlyDeduction']]) || 0,
        debtPaid: debtPaid,
        debtRemaining: totalDebt - debtPaid,
        debtReason: row[cols['DebtReason']] || '',
        status: row[cols['Status']] || 'ACTIVE'
      };
    }
  }
  
  return null;
}

/**
 * 添加人员（统一方法）
 */
function addPerson_(currentUser, personData, type) {
  const config = PERSON_CONFIG[type];
  if (!config) return { success: false, message: '无效的人员类型' };
  
  if (!currentUser || !config.allowedRoles.includes(currentUser.role)) {
    logOperation_('ADD_' + type + '_DENIED', {
      staffName: personData.staffName,
      reason: '权限不足'
    }, currentUser ? currentUser.username : 'UNKNOWN');
    return { success: false, message: '无权限添加' + config.label };
  }
  
  const validation = validateStaffData_(personData);
  if (!validation.valid) {
    logOperation_('ADD_' + type + '_FAILED', {
      staffName: personData.staffName,
      reason: '验证失败: ' + validation.errors.join('; ')
    }, currentUser.username);
    return { success: false, message: validation.errors.join('；') };
  }
  
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(config.sheetName);
  
  // 完整表头（25列 STAFF / 24列 MANAGERS）
  const headers = type === 'STAFF' 
    ? ['StaffName', 'Salary', 'CompanyName', 'BankHolder', 'BankType', 'BankAccount', 
       'PaymentMethod', 'IsManager', 'JoinDate', 'LeaveDate', 'TotalDebt', 'MonthlyDeduction', 
       'DebtPaid', 'DebtReason', 'LastBankUpdate', 'BankChangeNote', 
       'OldBankHolder', 'OldBankType', 'OldBankAccount',
       'OldSalary', 'SalaryChangeDate', 'SalaryChangeNote',
       'Status', 'CreatedBy', 'CreatedAt']
    : ['StaffName', 'Salary', 'CompanyName', 'BankHolder', 'BankType', 'BankAccount', 
       'PaymentMethod', 'JoinDate', 'LeaveDate', 'TotalDebt', 'MonthlyDeduction', 
       'DebtPaid', 'DebtReason', 'LastBankUpdate', 'BankChangeNote', 
       'OldBankHolder', 'OldBankType', 'OldBankAccount',
       'OldSalary', 'SalaryChangeDate', 'SalaryChangeNote',
       'Status', 'CreatedBy', 'CreatedAt'];
  
  if (!sheet) {
    sheet = ss.insertSheet(config.sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  const cleanData = {
    staffName: sanitizeInput_(personData.staffName, 'name'),
    salary: sanitizeInput_(personData.salary, 'number'),
    companyName: sanitizeInput_(personData.companyName, 'string'),
    bankHolder: sanitizeInput_(personData.bankHolder, 'string'),
    bankType: sanitizeInput_(personData.bankType, 'string'),
    bankAccount: sanitizeInput_(personData.bankAccount, 'string'),
    paymentMethod: personData.paymentMethod || 'BANK',
    joinDate: personData.joinDate || '',
    totalDebt: sanitizeInput_(personData.totalDebt, 'number'),
    monthlyDeduction: sanitizeInput_(personData.monthlyDeduction, 'number'),
    debtReason: sanitizeInput_(personData.debtReason, 'string')
  };
  
  const existing = findPersonInSheet_(cleanData.staffName, type);
  if (existing) {
    logOperation_('ADD_' + type + '_FAILED', {
      staffName: cleanData.staffName,
      reason: '姓名已存在'
    }, currentUser.username);
    return { success: false, message: config.label + '姓名已存在' };
  }
  
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  // ✅ 完整数据行（与表头对应）
  const rowData = type === 'STAFF'
      ? [cleanData.staffName, cleanData.salary, cleanData.companyName,
         cleanData.bankHolder, cleanData.bankType, cleanData.bankAccount,
         cleanData.paymentMethod, '',  // IsManager
         cleanData.joinDate, '',  // LeaveDate
         cleanData.totalDebt, cleanData.monthlyDeduction, 0,  // DebtPaid
         cleanData.debtReason,
         '', '',  // LastBankUpdate, BankChangeNote
         '', '', '',  // OldBankHolder, OldBankType, OldBankAccount
         '', '', '',  // OldSalary, SalaryChangeDate, SalaryChangeNote
         'ACTIVE', currentUser.username, now]
      : [cleanData.staffName, cleanData.salary, cleanData.companyName,
         cleanData.bankHolder, cleanData.bankType, cleanData.bankAccount,
         cleanData.paymentMethod,
         cleanData.joinDate, '',  // LeaveDate
         cleanData.totalDebt, cleanData.monthlyDeduction, 0,  // DebtPaid
         cleanData.debtReason,
         '', '',  // LastBankUpdate, BankChangeNote
         '', '', '',  // OldBankHolder, OldBankType, OldBankAccount
         '', '', '',  // OldSalary, SalaryChangeDate, SalaryChangeNote
         'ACTIVE', currentUser.username, now];
  
  sheet.appendRow(rowData);
  
  logOperation_('ADD_' + type, {
    staffName: cleanData.staffName,
    salary: cleanData.salary
  }, currentUser.username);
  
  return { success: true, type: type };
}

/**
 * 更新人员信息（统一方法）
 */
function updatePerson_(currentUser, staffName, updates, type) {
  const config = PERSON_CONFIG[type];
  if (!config) return { success: false, message: '无效的人员类型' };
  
  if (!currentUser || !config.allowedRoles.includes(currentUser.role)) {
    return { success: false, message: '无权限修改' + config.label };
  }
  
  const cleanName = sanitizeInput_(staffName, 'name');
  const person = findPersonInSheet_(cleanName, type);
  if (!person) {
    return { success: false, message: '找不到' + config.label };
  }
  
  const sheet = person.sheet;
  const rowIndex = person.rowIndex;
  const cols = person.cols;
  const rowData = person.data.slice();
  
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  let changeNote = '';
  const oldBankHolder = person.bankHolder;
  const oldBankType = person.bankType;
  const oldBank = person.bankAccount;
  const oldMethod = person.paymentMethod;
  
  const cleanUpdates = {};
  if (updates.salary !== undefined) cleanUpdates.salary = sanitizeInput_(updates.salary, 'number');
  if (updates.companyName !== undefined) cleanUpdates.companyName = sanitizeInput_(updates.companyName, 'string');
  if (updates.bankHolder !== undefined) cleanUpdates.bankHolder = sanitizeInput_(updates.bankHolder, 'string');
  if (updates.bankType !== undefined) cleanUpdates.bankType = sanitizeInput_(updates.bankType, 'string');
  if (updates.bankAccount !== undefined) cleanUpdates.bankAccount = sanitizeInput_(updates.bankAccount, 'string');
  if (updates.paymentMethod !== undefined) cleanUpdates.paymentMethod = updates.paymentMethod;
  if (updates.leaveDate !== undefined) cleanUpdates.leaveDate = updates.leaveDate;
  if (updates.totalDebt !== undefined) cleanUpdates.totalDebt = sanitizeInput_(updates.totalDebt, 'number');
  if (updates.monthlyDeduction !== undefined) cleanUpdates.monthlyDeduction = sanitizeInput_(updates.monthlyDeduction, 'number');
  if (updates.debtReason !== undefined) cleanUpdates.debtReason = sanitizeInput_(updates.debtReason, 'string');
  if (updates.status !== undefined) cleanUpdates.status = updates.status;
  
  // 检测银行信息是否变更
  let bankInfoChanged = false;
  if (cleanUpdates.bankAccount && cleanUpdates.bankAccount !== oldBank) {
    changeNote += '银行账号: ' + oldBank + ' → ' + cleanUpdates.bankAccount + '; ';
    bankInfoChanged = true;
  }
  if (cleanUpdates.bankHolder && cleanUpdates.bankHolder !== oldBankHolder) {
    changeNote += '户名: ' + oldBankHolder + ' → ' + cleanUpdates.bankHolder + '; ';
    bankInfoChanged = true;
  }
  if (cleanUpdates.bankType && cleanUpdates.bankType !== oldBankType) {
    changeNote += '银行: ' + oldBankType + ' → ' + cleanUpdates.bankType + '; ';
    bankInfoChanged = true;
  }
  if (cleanUpdates.paymentMethod && cleanUpdates.paymentMethod !== oldMethod) {
    changeNote += '付款方式: ' + oldMethod + ' → ' + cleanUpdates.paymentMethod + '; ';
  }
  
  const fieldMap = {
    salary: 'Salary',
    companyName: 'CompanyName',
    bankHolder: 'BankHolder',
    bankType: 'BankType',
    bankAccount: 'BankAccount',
    paymentMethod: 'PaymentMethod',
    leaveDate: 'LeaveDate',
    totalDebt: 'TotalDebt',
    monthlyDeduction: 'MonthlyDeduction',
    debtReason: 'DebtReason',
    status: 'Status'
  };
  
  for (const [field, colName] of Object.entries(fieldMap)) {
    if (cleanUpdates[field] !== undefined && cols[colName] !== undefined) {
      rowData[cols[colName]] = cleanUpdates[field];
    }
  }
  
  // 如果银行信息变更，保存旧值到 OldBank* 列
  if (bankInfoChanged) {
    if (cols['OldBankHolder'] !== undefined) rowData[cols['OldBankHolder']] = oldBankHolder;
    if (cols['OldBankType'] !== undefined) rowData[cols['OldBankType']] = oldBankType;
    if (cols['OldBankAccount'] !== undefined) rowData[cols['OldBankAccount']] = oldBank;
  }
  
  if (changeNote) {
    if (cols['LastBankUpdate'] !== undefined) rowData[cols['LastBankUpdate']] = now;
    if (cols['BankChangeNote'] !== undefined) rowData[cols['BankChangeNote']] = changeNote;
  }
  
  sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  
  logOperation_('UPDATE_' + type, {
    staffName: cleanName,
    changes: cleanUpdates,
    changeNote: changeNote
  }, currentUser.username);
  
  return { success: true, changeNote: changeNote };
}

/**
 * 设置人员离职（统一方法）
 */
function setPersonLeave_(currentUser, staffName, leaveDate, type) {
  const config = PERSON_CONFIG[type];
  if (!config) return { success: false, message: '无效的人员类型' };
  
  if (!currentUser || !config.allowedRoles.includes(currentUser.role)) {
    return { success: false, message: '无权限处理' + config.label + '离职' };
  }
  
  const result = updatePerson_(currentUser, staffName, {
    leaveDate: leaveDate,
    status: 'LEFT'
  }, type);
  
  if (result.success) {
    logOperation_('LEAVE_' + type, {
      staffName: staffName,
      leaveDate: leaveDate
    }, currentUser.username);
  }
  
  return result;
}

/***********************
 * 密码安全
 ***********************/

/**
 * 密码哈希（带盐值）
 */
function hashPassword_(plain, salt) {
  if (!salt) {
    salt = Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  }
  
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + plain);
  const hash = bytes.map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join('');
  
  return salt + ':' + hash;
}

/**
 * 验证密码
 */
function verifyPassword_(inputPassword, storedPassword) {
  if (!storedPassword || !inputPassword) return false;
  
  if (storedPassword.includes(':')) {
    const [salt, hash] = storedPassword.split(':');
    const inputHash = hashPassword_(inputPassword, salt);
    return inputHash === storedPassword;
  } else {
    const oldHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, inputPassword);
    const oldHashStr = oldHash.map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join('');
    return storedPassword === oldHashStr;
  }
}

/**
 * 检查密码是否需要升级
 */
function needsPasswordUpgrade_(storedPassword) {
  return storedPassword && !storedPassword.includes(':');
}

/***********************
 * 操作日志系统
 ***********************/

const LOG_ACTIONS = {
  ADD_STAFF: '添加员工',
  ADD_MANAGER: '添加主管',
  ADD_STAFF_FAILED: '添加员工失败',
  ADD_MANAGER_FAILED: '添加主管失败',
  ADD_STAFF_DENIED: '添加员工拒绝',
  ADD_MANAGER_DENIED: '添加主管拒绝',
  UPDATE_STAFF: '修改员工',
  UPDATE_MANAGER: '修改主管',
  LEAVE_STAFF: '员工离职',
  LEAVE_MANAGER: '主管离职',
  ADD_SALARY: '录入工资',
  ADD_SALARY_FAILED: '录入工资失败',
  SUBMIT_SALARY: '提交工资',
  DELETE_SALARY: '删除工资',
  MARK_PAID: '标记已发放',
  BATCH_PAID: '批量发放',
  ADD_USER: '创建账号',
  ADD_USER_FAILED: '创建账号失败',
  UPDATE_USER: '修改账号',
  LOGIN: '用户登录',
  LOGIN_FAILED: '登录失败',
  LOGOUT: '用户登出',
  CHANGE_PASSWORD: '修改密码',
  CHANGE_PASSWORD_FAILED: '修改密码失败'
};

/**
 * 记录操作日志
 */
function logOperation_(action, details, operator) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName(SHEET_PAYMENTS);
    
    const headers = ['LogID', 'Timestamp', 'Action', 'ActionName', 'Operator', 
                     'TargetType', 'TargetName', 'Details', 'IPInfo'];
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_PAYMENTS);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#2D3748')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    const tz = Session.getScriptTimeZone();
    const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    const logId = 'LOG-' + new Date().getTime();
    
    const targetType = details.type || (action && action.includes('MANAGER') ? 'MANAGER' : 
                       action && action.includes('STAFF') ? 'STAFF' : 
                       action && action.includes('USER') ? 'USER' : 
                       action && action.includes('SALARY') ? 'SALARY' : 'OTHER');
    const targetName = details.staffName || details.username || details.month || '';
    
    let detailStr = '';
    if (typeof details === 'object') {
      const parts = [];
      for (const [key, value] of Object.entries(details)) {
        if (key !== 'type' && key !== 'staffName' && key !== 'username' && value) {
          if (typeof value === 'object') {
            parts.push(key + ': ' + JSON.stringify(value));
          } else {
            parts.push(key + ': ' + value);
          }
        }
      }
      detailStr = parts.join('; ');
    } else {
      detailStr = String(details);
    }
    
    const actionName = LOG_ACTIONS[action] || action;
    
    sheet.appendRow([
      logId,
      now,
      action,
      actionName,
      operator || 'SYSTEM',
      targetType,
      targetName,
      detailStr.substring(0, 500),
      ''
    ]);
    
    return true;
  } catch (e) {
    Logger.log('日志记录失败: ' + e.message);
    return false;
  }
}

/**
 * 获取操作日志
 */
function getOperationLogs(currentUser, filter) {
  if (!currentUser || currentUser.role !== 'ADMIN') {
    return { success: false, message: '只有管理员可以查看日志' };
  }
  
  filter = filter || {};
  const limit = filter.limit || 100;
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_PAYMENTS);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, logs: [] };
  }
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  
  // 构建列索引映射（忽略大小写）
  const cols = {};
  header.forEach((h, i) => {
    const key = String(h).trim();
    cols[key] = i;
    cols[key.toLowerCase()] = i;
  });
  
  // 获取列索引的辅助函数
  function getCol(name) {
    return cols[name] !== undefined ? cols[name] : 
           cols[name.toLowerCase()] !== undefined ? cols[name.toLowerCase()] : -1;
  }
  
  const colLogID = getCol('LogID');
  const colTimestamp = getCol('Timestamp');
  const colAction = getCol('Action');
  const colActionName = getCol('ActionName');
  const colOperator = getCol('Operator');
  const colTargetType = getCol('TargetType');
  const colTargetName = getCol('TargetName');
  const colDetails = getCol('Details');
  
  const logs = [];
  const cleanTargetName = filter.targetName ? sanitizeInput_(filter.targetName, 'string') : '';
  
  for (let i = data.length - 1; i >= 1 && logs.length < limit; i--) {
    const row = data[i];
    
    // 如果行为空，跳过
    if (!row[0] && !row[1]) continue;
    
    const rowAction = colAction >= 0 ? String(row[colAction] || '') : '';
    const rowOperator = colOperator >= 0 ? String(row[colOperator] || '') : '';
    const rowTargetName = colTargetName >= 0 ? String(row[colTargetName] || '') : '';
    
    if (filter.action && rowAction !== filter.action) continue;
    if (filter.operator && rowOperator !== filter.operator) continue;
    if (cleanTargetName && !rowTargetName.toLowerCase().includes(cleanTargetName.toLowerCase())) continue;
    
    if (filter.startDate && colTimestamp >= 0) {
      const logDate = new Date(row[colTimestamp]);
      if (logDate < new Date(filter.startDate)) continue;
    }
    if (filter.endDate && colTimestamp >= 0) {
      const logDate = new Date(row[colTimestamp]);
      if (logDate > new Date(filter.endDate)) continue;
    }
    
    logs.push({
      logId: colLogID >= 0 ? String(row[colLogID] || '') : '',
      timestamp: colTimestamp >= 0 ? String(row[colTimestamp] || '') : '',
      action: rowAction,
      actionName: colActionName >= 0 ? String(row[colActionName] || rowAction) : rowAction,
      operator: rowOperator,
      targetType: colTargetType >= 0 ? String(row[colTargetType] || '') : '',
      target: rowTargetName,
      targetName: rowTargetName,
      details: colDetails >= 0 ? String(row[colDetails] || '') : ''
    });
  }
  
  return { success: true, logs: logs };
}

/***********************
 * 主要功能函数
 ***********************/

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('薪资管理系统')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 登录（完整版 - 包含首次登录检测）
 */
function login(username, password) {
  const cleanUsername = sanitizeInput_(username, 'string');
  const cleanPassword = String(password || '').trim();
  
  if (!cleanUsername || !cleanPassword) {
    return { success: false, message: '请输入用户名和密码' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, message: 'Users 表不存在' };

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const uName = row[cols['Username']] || row[1];
    const pwd = row[cols['Password']] || row[2];
    const role = row[cols['Role']] || row[3];
    const status = row[cols['Status']] || row[4];
    const displayName = row[cols['DisplayName']] || row[5];
    
    if (!uName) continue;
    
    const pwdCell = String(pwd).trim();
    
    if (String(uName).trim() === cleanUsername) {
      if (verifyPassword_(cleanPassword, pwdCell)) {
        if (status === 'INACTIVE' || status === 'DISABLED') {
          logOperation_('LOGIN_FAILED', {
            username: cleanUsername,
            reason: '账号已停用'
          }, cleanUsername);
          return { success: false, message: '账号已停用，请联系管理员' };
        }
        
        if (needsPasswordUpgrade_(pwdCell)) {
          const newHash = hashPassword_(cleanPassword);
          const pwdColIdx = cols['Password'] !== undefined ? cols['Password'] : 2;
          sheet.getRange(i + 1, pwdColIdx + 1).setValue(newHash);
        }
        
        let mustChangePassword = false;
        if (cols['MustChangePassword'] !== undefined) {
          const mustChange = row[cols['MustChangePassword']];
          mustChangePassword = mustChange === true || mustChange === 'YES' || mustChange === 1 || mustChange === 'TRUE';
        }
        
        logOperation_('LOGIN', { role: role || 'SECRETARY' }, uName);
        
        return {
          success: true,
          user: { 
            username: uName, 
            role: role || 'SECRETARY', 
            displayName: displayName || uName,
            mustChangePassword: mustChangePassword
          }
        };
      }
    }
  }
  
  logOperation_('LOGIN_FAILED', {
    username: cleanUsername,
    reason: '用户名或密码错误'
  }, cleanUsername);
  
  return { success: false, message: '用户名或密码错误' };
}

/**
 * 更换密码
 */
function changePassword(currentUser, oldPassword, newPassword, isFirstLogin) {
  if (!currentUser || !currentUser.username) {
    return { success: false, message: '未登录' };
  }
  
  const cleanNewPassword = String(newPassword || '').trim();
  if (!cleanNewPassword || cleanNewPassword.length < 6) {
    return { success: false, message: '新密码至少需要6个字符' };
  }
  
  if (cleanNewPassword.length > 50) {
    return { success: false, message: '密码不能超过50个字符' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, message: 'Users 表不存在' };
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const usernameCol = cols['Username'] !== undefined ? cols['Username'] : 1;
  const passwordCol = cols['Password'] !== undefined ? cols['Password'] : 2;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][usernameCol]).trim() === currentUser.username) {
      const storedPassword = String(data[i][passwordCol]).trim();
      
      if (!isFirstLogin) {
        const cleanOldPassword = String(oldPassword || '').trim();
        if (!cleanOldPassword) {
          return { success: false, message: '请输入当前密码' };
        }
        if (!verifyPassword_(cleanOldPassword, storedPassword)) {
          logOperation_('CHANGE_PASSWORD_FAILED', {
            username: currentUser.username,
            reason: '当前密码错误'
          }, currentUser.username);
          return { success: false, message: '当前密码错误' };
        }
      }
      
      if (verifyPassword_(cleanNewPassword, storedPassword)) {
        return { success: false, message: '新密码不能与当前密码相同' };
      }
      
      const newHash = hashPassword_(cleanNewPassword);
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, passwordCol + 1).setValue(newHash);
      
      if (cols['MustChangePassword'] !== undefined) {
        sheet.getRange(rowIndex, cols['MustChangePassword'] + 1).setValue('');
      }
      
      logOperation_('CHANGE_PASSWORD', {
        username: currentUser.username,
        isFirstLogin: isFirstLogin
      }, currentUser.username);
      
      return { success: true, message: '密码修改成功' };
    }
  }
  
  return { success: false, message: '找不到用户' };
}

/**
 * 添加用户账号（完整版）
 */
function addUserAccount(currentUser, newUser) {
  if (!currentUser || currentUser.role !== 'ADMIN') {
    return { success: false, message: '无权限' };
  }
  
  const validation = validateUserData_(newUser);
  if (!validation.valid) {
    logOperation_('ADD_USER_FAILED', {
      username: newUser.username,
      reason: '验证失败: ' + validation.errors.join('; ')
    }, currentUser.username);
    return { success: false, message: validation.errors.join('；') };
  }
  
  newUser.username = sanitizeInput_(newUser.username, 'string');
  newUser.displayName = sanitizeInput_(newUser.displayName, 'string');
  
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    sheet.getRange(1, 1, 1, 7).setValues([['UserID', 'Username', 'Password', 'Role', 'Status', 'DisplayName', 'MustChangePassword']]);
  }
  
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let mustChangePwdCol = header.indexOf('MustChangePassword');
  if (mustChangePwdCol === -1) {
    mustChangePwdCol = sheet.getLastColumn();
    sheet.getRange(1, mustChangePwdCol + 1).setValue('MustChangePassword');
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === newUser.username) {
      return { success: false, message: '用户名已存在' };
    }
  }

  const newId = data.length;
  const hashedPwd = hashPassword_(newUser.password);
  const requireFirstChange = newUser.requirePasswordChange !== false;
  
  sheet.appendRow([
    newId, 
    newUser.username, 
    hashedPwd, 
    newUser.role || 'SECRETARY', 
    'ACTIVE', 
    newUser.displayName || newUser.username,
    requireFirstChange ? 'YES' : ''
  ]);
  
  logOperation_('ADD_USER', {
    username: newUser.username,
    role: newUser.role,
    displayName: newUser.displayName
  }, currentUser.username);
  
  return { success: true };
}

/***********************
 * 员工管理
 ***********************/

/**
 * 获取员工列表（书记只能看到自己创建的员工）
 */
function getStaffList(currentUser) {
  if (!currentUser) return [];
  
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('Staff');
    
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var cols = {};
    for (var h = 0; h < headers.length; h++) {
      cols[String(headers[h]).trim()] = h;
    }
    
    var result = [];
    var username = String(currentUser.username || '').toLowerCase().trim();
    var now = new Date();
    var alertCutoffDate = new Date(now.getTime() - BANK_CHANGE_ALERT_DAYS * 24 * 60 * 60 * 1000);
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var staffName = row[cols['StaffName']];
      if (!staffName) continue;
      
      var createdBy = String(row[cols['CreatedBy']] || '').toLowerCase().trim();
      
      if (currentUser.role === 'SECRETARY') {
        if (createdBy !== username) continue;
      }
      
      var status = row[cols['Status']] || 'ACTIVE';
      var totalDebt = Number(row[cols['TotalDebt']]) || 0;
      var debtPaid = Number(row[cols['DebtPaid']]) || 0;
      var monthlyDeduction = Number(row[cols['MonthlyDeduction']]) || 0;
      var debtRemaining = totalDebt - debtPaid;
      
      var bankChanged = false;
      if (cols['LastBankUpdate'] !== undefined && row[cols['LastBankUpdate']]) {
        var updateDate = new Date(row[cols['LastBankUpdate']]);
        if (updateDate > alertCutoffDate) {
          bankChanged = true;
        }
      }
      
      var obj = {
        staffName: String(staffName || ''),
        companyName: String(row[cols['CompanyName']] || ''),
        salary: Number(row[cols['Salary']]) || 0,
        status: String(status),
        Status: String(status),
        paymentMethod: String(row[cols['PaymentMethod']] || 'BANK'),
        bankHolder: String(row[cols['BankHolder']] || ''),
        bankType: String(row[cols['BankType']] || ''),
        bankAccount: String(row[cols['BankAccount']] || ''),
        joinDate: '',
        leaveDate: '',
        bankChanged: bankChanged,
        monthlyDeduction: monthlyDeduction,
        totalDebt: totalDebt,
        debtPaid: debtPaid,
        debtRemaining: debtRemaining,
        debtReason: String(row[cols['DebtReason']] || ''),
        createdBy: String(row[cols['CreatedBy']] || '')
      };
      
      result.push(obj);
    }
    
    return result;
    
  } catch (e) {
    console.log('错误: ' + e.message);
    return [];
  }
}

/**
 * 获取主管列表
 */
function getManagerList(currentUser) {
  if (!currentUser) return [];
  
  try {
    const result = getPersonList_('MANAGER', { includeLeft: false });
    return result || [];
  } catch (e) {
    Logger.log('[getManagerList] 错误: ' + e.message);
    return [];
  }
}

/**
 * 添加员工
 */
function addStaff(currentUser, staff) {
  if (!currentUser || currentUser.role === 'ACCOUNTANT') {
    return { success: false, message: '会计无权限添加员工' };
  }
  
  const isManager = staff.isManager === true || staff.isManager === 'YES';
  
  if (currentUser.role === 'SECRETARY' && isManager) {
    return { success: false, message: '书记不能添加主管，请联系管理员' };
  }
  if (currentUser.role === 'ADMIN' && !isManager) {
    return { success: false, message: '管理员请添加主管。员工由书记添加' };
  }
  
  const type = isManager ? 'MANAGER' : 'STAFF';
  return addPerson_(currentUser, staff, type);
}

/**
 * 添加主管
 */
function addManager(currentUser, manager) {
  return addPerson_(currentUser, manager, 'MANAGER');
}

/**
 * 更新员工信息
 */
function updateStaffInfo(currentUser, staffName, updates) {
  if (!currentUser || currentUser.role === 'ACCOUNTANT') {
    return { success: false, message: '会计无权限修改' };
  }
  return updatePerson_(currentUser, staffName, updates, 'STAFF');
}

/**
 * 更新主管信息
 */
function updateManagerInfo(currentUser, staffName, updates) {
  return updatePerson_(currentUser, staffName, updates, 'MANAGER');
}

/**
 * 设置员工离职
 */
function setStaffLeave(currentUser, staffName, leaveDate) {
  if (!currentUser) return { success: false, message: '未登录' };
  
  const cleanName = sanitizeInput_(staffName, 'name');
  
  const staffPerson = findPersonInSheet_(cleanName, 'STAFF');
  const managerPerson = findPersonInSheet_(cleanName, 'MANAGER');
  
  if (staffPerson) {
    if (currentUser.role === 'ADMIN') {
      return { success: false, message: '管理员请处理主管离职，员工离职由书记处理' };
    }
    return setPersonLeave_(currentUser, cleanName, leaveDate, 'STAFF');
  } else if (managerPerson) {
    if (currentUser.role === 'SECRETARY') {
      return { success: false, message: '书记不能处理主管离职' };
    }
    return setPersonLeave_(currentUser, cleanName, leaveDate, 'MANAGER');
  } else {
    return { success: false, message: '找不到该人员' };
  }
}

/**
 * 处理员工离职（包含创建最后工资草稿）
 */
function processStaffLeave(currentUser, leaveData) {
  if (!currentUser) return { success: false, message: '未登录' };
  
  const {
    staffName,
    staffType,
    leaveDate,
    leaveReason,
    salaryMonth,
    paymentMethod,
    basicSalary,
    deduction,
    remark
  } = leaveData;
  
  const cleanName = sanitizeInput_(staffName, 'name');
  const cleanMonth = sanitizeInput_(salaryMonth, 'string');
  const cleanReason = sanitizeInput_(leaveReason, 'string');
  const cleanRemark = sanitizeInput_(remark, 'string');
  
  if (!cleanName) return { success: false, message: '员工姓名不能为空' };
  if (!leaveDate) return { success: false, message: '请选择离职日期' };
  if (!cleanMonth) return { success: false, message: '请选择工资月份' };
  
  const type = staffType === 'MANAGER' ? 'MANAGER' : 'STAFF';
  
  // 权限检查
  if (type === 'STAFF' && currentUser.role === 'ADMIN') {
    return { success: false, message: '管理员请处理主管离职，员工离职由书记处理' };
  }
  if (type === 'MANAGER' && currentUser.role === 'SECRETARY') {
    return { success: false, message: '书记不能处理主管离职' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  // 1. 设置离职状态
  const leaveResult = setPersonLeave_(currentUser, cleanName, leaveDate, type);
  if (!leaveResult.success) {
    return leaveResult;
  }
  
  // 2. 更新离职原因（如果有）
  if (cleanReason) {
    const sheetName = type === 'MANAGER' ? SHEET_MANAGERS : SHEET_STAFF;
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      const header = data[0];
      const cols = {};
      header.forEach((h, i) => cols[String(h).trim()] = i);
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][cols['StaffName']] || '').toLowerCase() === cleanName.toLowerCase()) {
          // 如果有LeaveReason列则更新
          if (cols['LeaveReason'] !== undefined) {
            sheet.getRange(i + 1, cols['LeaveReason'] + 1).setValue(cleanReason);
          }
          break;
        }
      }
    }
  }
  
  // 3. 创建最后工资草稿
  const salarySheet = ss.getSheetByName(SHEET_SALARY);
  if (!salarySheet) {
    return { success: true, message: '离职已设置，但工资表不存在，请手动创建工资记录' };
  }
  
  const salaryHeader = salarySheet.getRange(1, 1, 1, salarySheet.getLastColumn()).getValues()[0];
  const salaryCols = {};
  salaryHeader.forEach((h, i) => salaryCols[String(h).trim()] = i);
  
  // 检查是否已存在该月份的记录
  const existingData = salarySheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    const rowName = String(existingData[i][salaryCols['StaffName']] || '').toLowerCase();
    const rowMonth = String(existingData[i][salaryCols['Month']] || '');
    if (rowName === cleanName.toLowerCase() && rowMonth === cleanMonth) {
      // 已存在记录，更新它
      const rowIndex = i + 1;
      if (salaryCols['BasicSalary'] !== undefined) {
        salarySheet.getRange(rowIndex, salaryCols['BasicSalary'] + 1).setValue(basicSalary);
      }
      if (salaryCols['Deduction'] !== undefined) {
        salarySheet.getRange(rowIndex, salaryCols['Deduction'] + 1).setValue(deduction || 0);
      }
      if (salaryCols['DeductionReason'] !== undefined) {
        salarySheet.getRange(rowIndex, salaryCols['DeductionReason'] + 1).setValue(cleanRemark || '离职最后工资');
      }
      if (salaryCols['UpdatedAt'] !== undefined) {
        salarySheet.getRange(rowIndex, salaryCols['UpdatedAt'] + 1).setValue(now);
      }
      
      logOperation_('LEAVE_WITH_SALARY', {
        staffName: cleanName,
        type: type,
        leaveDate: leaveDate,
        leaveReason: cleanReason,
        month: cleanMonth,
        basicSalary: basicSalary,
        deduction: deduction,
        paymentMethod: paymentMethod,
        action: 'updated_existing'
      }, currentUser.username);
      
      return { success: true, message: '离职已设置，已更新现有工资记录' };
    }
  }
  
  // 创建新的工资记录
  const newRow = new Array(salaryHeader.length).fill('');
  
  if (salaryCols['StaffName'] !== undefined) newRow[salaryCols['StaffName']] = cleanName;
  if (salaryCols['Month'] !== undefined) newRow[salaryCols['Month']] = cleanMonth;
  if (salaryCols['BasicSalary'] !== undefined) newRow[salaryCols['BasicSalary']] = basicSalary;
  if (salaryCols['Deduction'] !== undefined) newRow[salaryCols['Deduction']] = deduction || 0;
  if (salaryCols['DeductionReason'] !== undefined) newRow[salaryCols['DeductionReason']] = cleanRemark || '离职最后工资';
  if (salaryCols['PaymentMethod'] !== undefined) newRow[salaryCols['PaymentMethod']] = paymentMethod || 'BANK';
  if (salaryCols['SubmitStatus'] !== undefined) newRow[salaryCols['SubmitStatus']] = 'DRAFT';
  if (salaryCols['PaymentStatus'] !== undefined) newRow[salaryCols['PaymentStatus']] = 'PENDING';
  if (salaryCols['CreatedBy'] !== undefined) newRow[salaryCols['CreatedBy']] = currentUser.username;
  if (salaryCols['CreatedAt'] !== undefined) newRow[salaryCols['CreatedAt']] = now;
  if (salaryCols['UpdatedAt'] !== undefined) newRow[salaryCols['UpdatedAt']] = now;
  if (salaryCols['IsManager'] !== undefined) newRow[salaryCols['IsManager']] = type === 'MANAGER' ? 'YES' : 'NO';
  
  salarySheet.appendRow(newRow);
  
  logOperation_('LEAVE_WITH_SALARY', {
    staffName: cleanName,
    type: type,
    leaveDate: leaveDate,
    leaveReason: cleanReason,
    month: cleanMonth,
    basicSalary: basicSalary,
    deduction: deduction,
    paymentMethod: paymentMethod,
    action: 'created_new'
  }, currentUser.username);
  
  return { success: true, message: '离职已设置，最后工资草稿已创建' };
}

/**
 * 快速创建工资草稿
 */
function createQuickSalary(currentUser, data) {
  if (!currentUser) return { success: false, message: '未登录' };
  
  const {
    staffName,
    staffType,
    month,
    paymentMethod,
    basicSalary,
    bonus,
    remark,
    startDate
  } = data;
  
  const cleanName = sanitizeInput_(staffName, 'name');
  const cleanMonth = sanitizeInput_(month, 'string');
  const cleanRemark = sanitizeInput_(remark, 'string');
  
  if (!cleanName) return { success: false, message: '人员姓名不能为空' };
  if (!cleanMonth) return { success: false, message: '请选择工资月份' };
  
  const type = staffType === 'MANAGER' ? 'MANAGER' : 'STAFF';
  
  // 权限检查
  if (type === 'STAFF' && currentUser.role === 'ADMIN') {
    return { success: false, message: '管理员请为主管起薪，员工起薪由书记处理' };
  }
  if (type === 'MANAGER' && currentUser.role === 'SECRETARY') {
    return { success: false, message: '书记不能为主管起薪' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  const salarySheet = ss.getSheetByName(SHEET_SALARY);
  if (!salarySheet) {
    return { success: false, message: '工资表不存在' };
  }
  
  const salaryHeader = salarySheet.getRange(1, 1, 1, salarySheet.getLastColumn()).getValues()[0];
  const salaryCols = {};
  salaryHeader.forEach((h, i) => salaryCols[String(h).trim()] = i);
  
  // 检查是否已存在该月份的记录
  const existingData = salarySheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    const rowName = String(existingData[i][salaryCols['StaffName']] || '').toLowerCase();
    const rowMonth = String(existingData[i][salaryCols['Month']] || '');
    if (rowName === cleanName.toLowerCase() && rowMonth === cleanMonth) {
      return { success: false, message: `${cleanMonth} 月份的工资记录已存在` };
    }
  }
  
  // 计算实际金额
  const actualBasic = Number(basicSalary) || 0;
  const actualBonus = Number(bonus) || 0;
  const netSalary = actualBasic + actualBonus;
  
  // 构建备注
  let finalRemark = cleanRemark || '';
  if (startDate) {
    finalRemark = '起薪日期: ' + startDate + (finalRemark ? ' | ' + finalRemark : '');
  }
  if (actualBonus > 0) {
    finalRemark += (finalRemark ? ' | ' : '') + '加款: RM ' + actualBonus;
  }
  
  // 创建新的工资记录
  const newRow = new Array(salaryHeader.length).fill('');
  
  if (salaryCols['StaffName'] !== undefined) newRow[salaryCols['StaffName']] = cleanName;
  if (salaryCols['Month'] !== undefined) newRow[salaryCols['Month']] = cleanMonth;
  if (salaryCols['BasicSalary'] !== undefined) newRow[salaryCols['BasicSalary']] = netSalary; // 基本工资+奖金
  if (salaryCols['Deduction'] !== undefined) newRow[salaryCols['Deduction']] = 0; // 起薪不扣款
  if (salaryCols['DeductionReason'] !== undefined) newRow[salaryCols['DeductionReason']] = finalRemark;
  if (salaryCols['PaymentMethod'] !== undefined) newRow[salaryCols['PaymentMethod']] = paymentMethod || 'BANK';
  if (salaryCols['SubmitStatus'] !== undefined) newRow[salaryCols['SubmitStatus']] = 'DRAFT';
  if (salaryCols['PaymentStatus'] !== undefined) newRow[salaryCols['PaymentStatus']] = 'PENDING';
  if (salaryCols['CreatedBy'] !== undefined) newRow[salaryCols['CreatedBy']] = currentUser.username;
  if (salaryCols['CreatedAt'] !== undefined) newRow[salaryCols['CreatedAt']] = now;
  if (salaryCols['UpdatedAt'] !== undefined) newRow[salaryCols['UpdatedAt']] = now;
  if (salaryCols['IsManager'] !== undefined) newRow[salaryCols['IsManager']] = type === 'MANAGER' ? 'YES' : 'NO';
  
  salarySheet.appendRow(newRow);
  
  logOperation_('QUICK_SALARY', {
    staffName: cleanName,
    type: type,
    month: cleanMonth,
    basicSalary: actualBasic,
    bonus: actualBonus,
    netSalary: netSalary,
    startDate: startDate,
    paymentMethod: paymentMethod
  }, currentUser.username);
  
  return { success: true, message: '工资草稿已创建' };
}

/**
 * 调整薪资
 */
function adjustSalary(currentUser, data) {
  if (!currentUser) return { success: false, message: '未登录' };
  
  const {
    staffName,
    staffType,
    newSalary,
    effectiveDate,
    reason
  } = data;
  
  const cleanName = sanitizeInput_(staffName, 'name');
  const cleanReason = sanitizeInput_(reason, 'string');
  
  if (!cleanName) return { success: false, message: '人员姓名不能为空' };
  if (!newSalary || newSalary <= 0) return { success: false, message: '新月薪必须大于0' };
  if (!effectiveDate) return { success: false, message: '请选择生效日期' };
  if (!cleanReason) return { success: false, message: '请填写调薪原因' };
  
  const type = staffType === 'MANAGER' ? 'MANAGER' : 'STAFF';
  
  // 权限检查
  if (type === 'STAFF' && currentUser.role === 'ADMIN') {
    return { success: false, message: '管理员不能为员工调薪，请联系书记处理' };
  }
  if (type === 'MANAGER' && currentUser.role === 'SECRETARY') {
    return { success: false, message: '书记不能为主管调薪' };
  }
  if (currentUser.role === 'ACCOUNTANT') {
    return { success: false, message: '会计无权限调薪' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  // 选择正确的表
  const sheetName = type === 'MANAGER' ? SHEET_MANAGERS : SHEET_STAFF;
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return { success: false, message: '数据表不存在' };
  }
  
  const data2 = sheet.getDataRange().getValues();
  const headers = data2[0];
  const cols = {};
  headers.forEach((h, i) => cols[String(h).trim()] = i);
  
  // 查找人员
  let rowIndex = -1;
  let oldSalary = 0;
  
  for (let i = 1; i < data2.length; i++) {
    if (String(data2[i][cols['StaffName']] || '').toLowerCase() === cleanName.toLowerCase()) {
      rowIndex = i + 1;
      oldSalary = Number(data2[i][cols['Salary']]) || 0;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, message: '找不到人员信息' };
  }
  
  // 检查薪资是否有变化
  if (oldSalary === newSalary) {
    return { success: false, message: '新月薪与当前月薪相同，无需调整' };
  }
  
  // 更新薪资
  if (cols['Salary'] !== undefined) {
    sheet.getRange(rowIndex, cols['Salary'] + 1).setValue(newSalary);
  }
  
  // 保存旧薪资信息
  if (cols['OldSalary'] !== undefined) {
    sheet.getRange(rowIndex, cols['OldSalary'] + 1).setValue(oldSalary);
  }
  if (cols['SalaryChangeDate'] !== undefined) {
    sheet.getRange(rowIndex, cols['SalaryChangeDate'] + 1).setValue(effectiveDate);
  }
  if (cols['SalaryChangeNote'] !== undefined) {
    sheet.getRange(rowIndex, cols['SalaryChangeNote'] + 1).setValue(cleanReason);
  }
  
  // 记录日志
  const diff = newSalary - oldSalary;
  const percent = ((diff / oldSalary) * 100).toFixed(1);
  
  logOperation_('ADJUST_SALARY', {
    staffName: cleanName,
    type: type,
    oldSalary: oldSalary,
    newSalary: newSalary,
    diff: diff,
    percent: percent + '%',
    effectiveDate: effectiveDate,
    reason: cleanReason
  }, currentUser.username);
  
  // 同步更新 Salary 表中当月的工资记录（如果存在）
  const salarySheet = ss.getSheetByName(SHEET_SALARY);
  if (salarySheet && salarySheet.getLastRow() > 1) {
    const currentMonth = getCurrentMonth_();
    const salaryData = salarySheet.getDataRange().getValues();
    const salaryHeaders = salaryData[0];
    const salCols = {};
    salaryHeaders.forEach((h, i) => salCols[String(h).trim()] = i);
    
    // 查找当月该人员的工资记录
    for (let i = 1; i < salaryData.length; i++) {
      const rowName = String(salaryData[i][salCols['StaffName']] || '').toLowerCase();
      const rowMonth = String(salaryData[i][salCols['Month']] || '');
      const paymentStatus = String(salaryData[i][salCols['PaymentStatus']] || '').toUpperCase();
      
      // 只更新当月且未发放的记录
      if (rowName === cleanName.toLowerCase() && rowMonth === currentMonth && paymentStatus !== 'PAID') {
        const salaryRowIndex = i + 1;
        
        // 更新 BasicSalary
        if (salCols['BasicSalary'] !== undefined) {
          salarySheet.getRange(salaryRowIndex, salCols['BasicSalary'] + 1).setValue(newSalary);
        }
        
        // 重新计算实发工资
        const deduction = Number(salaryData[i][salCols['Deduction']]) || 0;
        const addition = Number(salaryData[i][salCols['Addition']]) || 0;
        const netSalary = newSalary - deduction + addition;
        
        if (salCols['NetSalary'] !== undefined) {
          salarySheet.getRange(salaryRowIndex, salCols['NetSalary'] + 1).setValue(netSalary);
        }
        
        // 更新修改时间
        if (salCols['UpdatedAt'] !== undefined) {
          salarySheet.getRange(salaryRowIndex, salCols['UpdatedAt'] + 1).setValue(now);
        }
        
        break; // 找到并更新后退出
      }
    }
  }
  
  return { 
    success: true, 
    message: `调薪成功！${cleanName} 月薪从 RM ${oldSalary.toLocaleString()} 调整为 RM ${newSalary.toLocaleString()}` 
  };
}

/***********************
 * 工资记录
 ***********************/

/**
 * 添加工资记录
 */
function addSalaryRecord(currentUser, record) {
  if (!currentUser) return { success: false, message: '未登录' };
  if (currentUser.role === 'ACCOUNTANT') {
    return { success: false, message: '会计无权限添加工资记录' };
  }
  
  const validation = validateSalaryRecord_(record);
  if (!validation.valid) {
    return { success: false, message: validation.errors.join('；') };
  }
  
  record.month = sanitizeInput_(record.month, 'string');
  record.staffName = sanitizeInput_(record.staffName, 'name');
  record.basicSalary = sanitizeInput_(record.basicSalary, 'number');
  record.deduction = sanitizeInput_(record.deduction, 'number');
  record.remark = sanitizeInput_(record.remark, 'string');
  
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET_SALARY);
  
  const headers = ['Month', 'Date', 'StaffName', 'IsManagerRecord', 'BasicSalary', 'ManualDeduction', 'AutoDeduction', 
                 'BankFee', 'Deduction', 'NetSalary', 'Remark', 'CreatedBy', 'CreatedAt', 'SubmitStatus', 'SubmittedAt',
                 'PaymentStatus', 'PaymentMethod', 'PaidAt', 'PaidBy'];
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SALARY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  const isAdminAddingManager = (currentUser.role === 'ADMIN');
  const targetSheetName = isAdminAddingManager ? SHEET_MANAGERS : SHEET_STAFF;
  const personSheet = ss.getSheetByName(targetSheetName);
  
  let isManager = isAdminAddingManager;
  let autoDeduction = 0;
  let proRataInfo = '';
  let basic = Number(record.basicSalary) || 0;
  let paymentMethod = 'BANK';
  let personRowIndex = -1;
  
  if (personSheet && personSheet.getLastRow() > 1) {
    const personData = personSheet.getDataRange().getValues();
    const personHeader = personData[0];
    const cols = {};
    personHeader.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < personData.length; i++) {
      if (personData[i][cols['StaffName']] === record.staffName) {
        personRowIndex = i + 1;
        
        if (!isAdminAddingManager && cols['IsManager'] !== undefined) {
          isManager = personData[i][cols['IsManager']] === 'YES';
        }
        
        const personSalary = Number(personData[i][cols['Salary']]) || 0;
        const joinDate = personData[i][cols['JoinDate']];
        const leaveDate = personData[i][cols['LeaveDate']];
        const totalDebt = Number(personData[i][cols['TotalDebt']]) || 0;
        const monthlyDed = Number(personData[i][cols['MonthlyDeduction']]) || 0;
        const debtPaid = Number(personData[i][cols['DebtPaid']]) || 0;
        paymentMethod = personData[i][cols['PaymentMethod']] || 'BANK';
        
        if (basic === 0) basic = personSalary;
        
        const parsed = parseMonth_(record.month);
        if (parsed.valid) {
          const year = parsed.year;
          const mon = parsed.month;
          const daysInMonth = new Date(year, mon, 0).getDate();
          let workDays = daysInMonth;
          
          if (joinDate) {
            const jd = new Date(joinDate);
            if (jd.getFullYear() === year && jd.getMonth() + 1 === mon) {
              workDays = daysInMonth - jd.getDate() + 1;
              proRataInfo = '入职' + jd.getDate() + '号,工作' + workDays + '/' + daysInMonth + '天';
            }
          }
          
          if (leaveDate) {
            const ld = new Date(leaveDate);
            if (ld.getFullYear() === year && ld.getMonth() + 1 === mon) {
              if (joinDate) {
                const jd = new Date(joinDate);
                if (jd.getFullYear() === year && jd.getMonth() + 1 === mon) {
                  workDays = ld.getDate() - jd.getDate() + 1;
                } else {
                  workDays = ld.getDate();
                }
              } else {
                workDays = ld.getDate();
              }
              proRataInfo = '离职' + ld.getDate() + '号,工作' + workDays + '/' + daysInMonth + '天';
            }
          }
          
          if (workDays < daysInMonth && workDays > 0) {
            basic = Math.round((personSalary / daysInMonth) * workDays * 100) / 100;
          }
        }
        
        const debtRemaining = totalDebt - debtPaid;
        if (monthlyDed > 0 && debtRemaining > 0) {
          autoDeduction = Math.min(monthlyDed, debtRemaining);
        }
        break;
      }
    }
  }
  
  if (currentUser.role === 'SECRETARY' && isManager) {
    return { success: false, message: '书记不能添加主管工资' };
  }
  if (currentUser.role === 'ADMIN' && !isManager) {
    return { success: false, message: '管理员请添加主管工资，员工工资由书记添加' };
  }
  
  const manualDeduction = Number(record.deduction) || 0;
  
  let bankFee = 0;
  if (paymentMethod === 'BANK' && !isManager) {
    bankFee = 15;
  }
  
  const totalDeduction = manualDeduction + autoDeduction + bankFee;
  const net = basic - totalDeduction;
  
  let remark = record.remark || '';
  if (proRataInfo) remark = (remark ? remark + '; ' : '') + '[按比例] ' + proRataInfo;
  if (autoDeduction > 0) remark = (remark ? remark + '; ' : '') + '[分期扣款] RM' + autoDeduction.toFixed(2);
  if (bankFee > 0) remark = (remark ? remark + '; ' : '') + '[汇款手续费] RM' + bankFee.toFixed(2);
  
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  sheet.appendRow([
    record.month || '', record.date || '', record.staffName || '', 
    isManager ? 'YES' : '',
    basic, manualDeduction, autoDeduction, bankFee, totalDeduction, net, remark,
    currentUser.username, now, 'DRAFT', '',
    'PENDING', paymentMethod, '', ''
  ]);
  
  if (autoDeduction > 0 && personRowIndex > 0 && personSheet) {
    const header = personSheet.getRange(1, 1, 1, personSheet.getLastColumn()).getValues()[0];
    const colDebtPaid = header.indexOf('DebtPaid');
    if (colDebtPaid >= 0) {
      const currentPaid = Number(personSheet.getRange(personRowIndex, colDebtPaid + 1).getValue()) || 0;
      personSheet.getRange(personRowIndex, colDebtPaid + 1).setValue(currentPaid + autoDeduction);
    }
  }
  
  logOperation_('ADD_SALARY', {
    staffName: record.staffName,
    month: record.month,
    basic: basic,
    net: net,
    isManager: isManager
  }, currentUser.username);
  
  return { 
    success: true,
    calculated: {
      basic: basic,
      manualDeduction: manualDeduction,
      autoDeduction: autoDeduction,
      totalDeduction: totalDeduction,
      net: net,
      proRataInfo: proRataInfo,
      isManager: isManager
    }
  };
}

/**
 * 获取草稿记录
 */
function getDraftRecords(currentUser, month) {
  if (!currentUser) return { success: false, message: '未登录', rows: [] };
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_SALARY);
  if (!sheet || sheet.getLastRow() < 2) return { success: true, rows: [] };
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const cleanMonth = sanitizeInput_(month, 'string');
  
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const submitStatus = cols['SubmitStatus'] !== undefined ? String(row[cols['SubmitStatus']]).trim() : '';
    if (submitStatus !== 'DRAFT') continue;
    
    const rowMonth = formatMonthValue_(row[cols['Month']]);
    if (cleanMonth && rowMonth !== cleanMonth) continue;
    
    const isManagerRecord = cols['IsManagerRecord'] !== undefined && String(row[cols['IsManagerRecord']] || '').trim().toUpperCase() === 'YES';
    
    if (currentUser.role === 'SECRETARY') {
      if (isManagerRecord) continue;
      if (cols['CreatedBy'] !== undefined && row[cols['CreatedBy']] !== currentUser.username) continue;
    }
    if (currentUser.role === 'ADMIN') {
      if (!isManagerRecord) continue;
    }
    
    rows.push({
      rowIndex: i + 1,
      month: rowMonth,
      staffName: row[cols['StaffName']] || '',
      basicSalary: Number(row[cols['BasicSalary']]) || 0,
      deduction: Number(row[cols['Deduction']]) || 0,
      autoDeduction: Number(row[cols['AutoDeduction']]) || 0,
      netSalary: Number(row[cols['NetSalary']]) || 0,
      remark: row[cols['Remark']] || '',
      isManager: isManagerRecord
    });
  }
  
  return { success: true, rows: rows, drafts: rows };
}

/**
 * 更新草稿记录
 */
function updateDraftRecord(currentUser, rowIndex, updates) {
  if (!currentUser) return { success: false, message: '未登录' };
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_SALARY);
  if (!sheet) return { success: false, message: '工资表不存在' };
  
  if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
    return { success: false, message: '无效的记录' };
  }
  
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const submitStatus = String(rowData[cols['SubmitStatus']] || '').trim();
  if (submitStatus !== 'DRAFT') {
    return { success: false, message: '只能编辑草稿状态的记录' };
  }
  
  const isManagerRecord = String(rowData[cols['IsManagerRecord']] || '').trim().toUpperCase() === 'YES';
  const createdBy = rowData[cols['CreatedBy']];
  
  if (currentUser.role === 'SECRETARY') {
    if (isManagerRecord) return { success: false, message: '书记不能编辑主管记录' };
    if (createdBy !== currentUser.username) return { success: false, message: '只能编辑自己创建的记录' };
  }
  if (currentUser.role === 'ADMIN') {
    if (!isManagerRecord) return { success: false, message: 'Admin只能编辑主管记录' };
  }
  if (currentUser.role === 'ACCOUNTANT') {
    return { success: false, message: '会计不能编辑记录' };
  }
  
  const basicSalary = sanitizeInput_(updates.basicSalary, 'number');
  const deduction = sanitizeInput_(updates.deduction, 'number');
  const remark = sanitizeInput_(updates.remark, 'string');
  
  if (basicSalary < 0 || basicSalary > 1000000) {
    return { success: false, message: '基本工资数值不合理' };
  }
  if (deduction < 0 || deduction > basicSalary) {
    return { success: false, message: '扣款数值不合理' };
  }
  
  const autoDeduction = Number(rowData[cols['AutoDeduction']]) || 0;
  const bankFee = Number(rowData[cols['BankFee']]) || 0;
  const totalDeduction = deduction + autoDeduction + bankFee;
  const netSalary = basicSalary - totalDeduction;
  
  if (cols['BasicSalary'] !== undefined) sheet.getRange(rowIndex, cols['BasicSalary'] + 1).setValue(basicSalary);
  if (cols['ManualDeduction'] !== undefined) sheet.getRange(rowIndex, cols['ManualDeduction'] + 1).setValue(deduction);
  if (cols['Deduction'] !== undefined) sheet.getRange(rowIndex, cols['Deduction'] + 1).setValue(totalDeduction);
  if (cols['NetSalary'] !== undefined) sheet.getRange(rowIndex, cols['NetSalary'] + 1).setValue(netSalary);
  if (cols['Remark'] !== undefined) sheet.getRange(rowIndex, cols['Remark'] + 1).setValue(remark);
  
  const staffName = rowData[cols['StaffName']];
  logOperation_('UPDATE_DRAFT', {
    staffName: staffName,
    basicSalary: basicSalary,
    deduction: deduction,
    netSalary: netSalary
  }, currentUser.username);
  
  return { 
    success: true, 
    message: '更新成功',
    updated: {
      basicSalary: basicSalary,
      deduction: deduction,
      netSalary: netSalary,
      remark: remark
    }
  };
}

/**
 * 删除草稿记录
 */
function deleteDraftRecord(currentUser, rowIndex) {
  if (!currentUser || currentUser.role === 'ACCOUNTANT') {
    return { success: false, message: '无权限' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_SALARY);
  if (!sheet) return { success: false, message: '工资表不存在' };
  
  if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
    return { success: false, message: '无效的记录' };
  }
  
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const submitStatus = String(rowData[cols['SubmitStatus']] || '').trim();
  if (submitStatus !== 'DRAFT') {
    return { success: false, message: '只能删除草稿状态的记录' };
  }
  
  const isManagerRecord = String(rowData[cols['IsManagerRecord']] || '').trim().toUpperCase() === 'YES';
  const createdBy = rowData[cols['CreatedBy']];
  const staffName = rowData[cols['StaffName']];
  const month = formatMonthValue_(rowData[cols['Month']]);
  
  if (currentUser.role === 'SECRETARY') {
    if (isManagerRecord) return { success: false, message: '书记不能删除主管记录' };
    if (createdBy !== currentUser.username) return { success: false, message: '只能删除自己创建的记录' };
  }
  if (currentUser.role === 'ADMIN') {
    if (!isManagerRecord) return { success: false, message: 'Admin只能删除主管记录' };
  }
  
  sheet.deleteRow(rowIndex);
  
  logOperation_('DELETE_SALARY', {
    staffName: staffName,
    month: month
  }, currentUser.username);
  
  return { success: true, message: '删除成功', staffName: staffName, month: month };
}

/**
 * 提交工资记录
 */
function submitSalaryRecords(currentUser, month) {
  if (!currentUser) return { success: false, message: '未登录' };
  if (currentUser.role !== 'ADMIN' && currentUser.role !== 'SECRETARY') {
    return { success: false, message: '无权限提交' };
  }
  
  const cleanMonth = sanitizeInput_(month, 'string');
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_SALARY);
  if (!sheet || sheet.getLastRow() < 2) return { success: true, count: 0 };
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const idxMonth = cols['Month'];
  const idxSubmitStatus = cols['SubmitStatus'];
  const idxSubmittedAt = cols['SubmittedAt'];
  const idxIsManager = cols['IsManagerRecord'];
  const idxCreatedBy = cols['CreatedBy'];
  
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 1;
    
    let rowMonth = row[idxMonth];
    if (rowMonth instanceof Date) {
      rowMonth = rowMonth.getFullYear() + '-' + String(rowMonth.getMonth() + 1).padStart(2, '0');
    } else {
      rowMonth = String(rowMonth || '').trim();
    }
    
    if (cleanMonth && rowMonth !== cleanMonth) continue;
    
    const submitStatus = String(row[idxSubmitStatus] || '').trim();
    if (submitStatus !== 'DRAFT') continue;
    
    const isManagerRecord = idxIsManager !== undefined && String(row[idxIsManager] || '').trim().toUpperCase() === 'YES';
    const createdBy = idxCreatedBy !== undefined ? String(row[idxCreatedBy] || '').trim() : '';
    
    if (currentUser.role === 'SECRETARY') {
      if (isManagerRecord) continue;
      if (createdBy.toLowerCase() !== currentUser.username.toLowerCase()) continue;
    }
    if (currentUser.role === 'ADMIN') {
      if (!isManagerRecord) continue;
    }
    
    try {
      var statusCell = sheet.getRange(rowIndex, idxSubmitStatus + 1);
      statusCell.clearDataValidations();
      statusCell.setValue('SUBMITTED');
      
      if (idxSubmittedAt !== undefined) {
        var timeCell = sheet.getRange(rowIndex, idxSubmittedAt + 1);
        timeCell.clearDataValidations();
        timeCell.setValue(now);
      }
      count++;
    } catch (e) {
      console.log('提交失败:', e.message);
    }
  }
  
  if (count > 0) {
    logOperation_('SUBMIT_SALARY', {
      month: cleanMonth,
      count: count
    }, currentUser.username);
  }
  
  return { success: true, count: count };
}

/***********************
 * 支付管理
 ***********************/

/**
 * 获取待发工资列表
 */
function getPendingPayments(currentUser, month) {
  if (!currentUser) return { success: false, message: '未登录' };
  if (currentUser.role !== 'ADMIN' && currentUser.role !== 'ACCOUNTANT') {
    return { success: false, message: '无权限' };
  }

  const cleanMonth = sanitizeInput_(month, 'string');

  const ss = SpreadsheetApp.getActive();
  const salarySheet = ss.getSheetByName(SHEET_SALARY);
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  const managerSheet = ss.getSheetByName(SHEET_MANAGERS);

  if (!salarySheet) return { success: false, message: '工资表不存在' };

  // 统计人数
  let totalEmployeeCount = 0;
  let totalManagerCount = 0;
  const staffBySecretary = {};
  
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const staffCountData = staffSheet.getDataRange().getValues();
    const staffCountHeader = staffCountData[0];
    const empStatusIdx = staffCountHeader.indexOf('Status');
    const empNameIdx = staffCountHeader.indexOf('StaffName');
    const createdByIdx = staffCountHeader.indexOf('CreatedBy');
    
    for (let i = 1; i < staffCountData.length; i++) {
      const name = staffCountData[i][empNameIdx];
      const status = staffCountData[i][empStatusIdx];
      if (name && status !== 'LEFT') {
        totalEmployeeCount++;
        if (createdByIdx !== -1) {
          const createdBy = staffCountData[i][createdByIdx] || '未知';
          if (!staffBySecretary[createdBy]) staffBySecretary[createdBy] = 0;
          staffBySecretary[createdBy]++;
        }
      }
    }
  }
  
  if (managerSheet && managerSheet.getLastRow() > 1) {
    const mgrCountData = managerSheet.getDataRange().getValues();
    const mgrCountHeader = mgrCountData[0];
    const mgrStatusIdx = mgrCountHeader.indexOf('Status');
    const mgrNameIdx = mgrCountHeader.indexOf('StaffName');
    
    for (let i = 1; i < mgrCountData.length; i++) {
      const name = mgrCountData[i][mgrNameIdx];
      const status = mgrCountData[i][mgrStatusIdx];
      if (name && status !== 'LEFT') totalManagerCount++;
    }
  }

  // 获取人员银行信息
  const staffMap = {};
  const now = new Date();
  const alertCutoffDate = new Date(now.getTime() - BANK_CHANGE_ALERT_DAYS * 24 * 60 * 60 * 1000);
  
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const staffData = staffSheet.getDataRange().getValues();
    const staffHeader = staffData[0];
    const sCols = {};
    staffHeader.forEach((h, i) => sCols[String(h).trim()] = i);

    for (let i = 1; i < staffData.length; i++) {
      const row = staffData[i];
      const name = row[sCols['StaffName']];
      if (!name) continue;

      let bankChanged = false;
      if (sCols['LastBankUpdate'] !== undefined && row[sCols['LastBankUpdate']]) {
        const updateDate = new Date(row[sCols['LastBankUpdate']]);
        if (updateDate > alertCutoffDate) bankChanged = true;
      }

      staffMap[name] = {
        bankHolder: row[sCols['BankHolder']] || '',
        bankType: row[sCols['BankType']] || '',
        bankAccount: row[sCols['BankAccount']] || '',
        paymentMethod: row[sCols['PaymentMethod']] || 'BANK',
        bankChanged: bankChanged,
        bankChangeNote: row[sCols['BankChangeNote']] || '',
        isManager: false,
        secretary: row[sCols['CreatedBy']] || '',
        // 欠款信息
        totalDebt: Number(row[sCols['TotalDebt']]) || 0,
        debtPaid: Number(row[sCols['DebtPaid']]) || 0,
        monthlyDeduction: Number(row[sCols['MonthlyDeduction']]) || 0,
        debtReason: row[sCols['DebtReason']] || ''
      };
    }
  }

  if (managerSheet && managerSheet.getLastRow() > 1) {
    const managerData = managerSheet.getDataRange().getValues();
    const managerHeader = managerData[0];
    const mCols = {};
    managerHeader.forEach((h, i) => mCols[String(h).trim()] = i);

    for (let i = 1; i < managerData.length; i++) {
      const row = managerData[i];
      const name = row[mCols['StaffName']];
      if (!name) continue;

      let bankChanged = false;
      if (mCols['LastBankUpdate'] !== undefined && row[mCols['LastBankUpdate']]) {
        const updateDate = new Date(row[mCols['LastBankUpdate']]);
        if (updateDate > alertCutoffDate) bankChanged = true;
      }

      staffMap[name] = {
        bankHolder: row[mCols['BankHolder']] || '',
        bankType: row[mCols['BankType']] || '',
        bankAccount: row[mCols['BankAccount']] || '',
        paymentMethod: row[mCols['PaymentMethod']] || 'BANK',
        bankChanged: bankChanged,
        bankChangeNote: row[mCols['BankChangeNote']] || '',
        isManager: true,
        // 欠款信息
        totalDebt: Number(row[mCols['TotalDebt']]) || 0,
        debtPaid: Number(row[mCols['DebtPaid']]) || 0,
        monthlyDeduction: Number(row[mCols['MonthlyDeduction']]) || 0,
        debtReason: row[mCols['DebtReason']] || ''
      };
    }
  }

  // 获取工资记录
  const salaryData = salarySheet.getDataRange().getValues();
  const salaryHeader = salaryData[0];
  const cols = {};
  salaryHeader.forEach((h, i) => cols[String(h).trim()] = i);

  const prevMonth = cleanMonth ? getPreviousMonth_(cleanMonth) : '';

  // 获取上月数据用于对比
  const lastMonthMap = {};
  for (let i = 1; i < salaryData.length; i++) {
    const rowMonth = formatMonthValue_(salaryData[i][cols['Month']]);
    if (rowMonth === prevMonth) {
      const staffName = salaryData[i][cols['StaffName']];
      lastMonthMap[staffName] = {
        basic: Number(salaryData[i][cols['BasicSalary']]) || 0,
        deduction: Number(salaryData[i][cols['Deduction']]) || 0,
        net: Number(salaryData[i][cols['NetSalary']]) || 0
      };
    }
  }

  const employeePending = [], employeePaid = [], managerPending = [], managerPaid = [];

  for (let i = 1; i < salaryData.length; i++) {
    const row = salaryData[i];
    const rowMonth = formatMonthValue_(row[cols['Month']]);
    if (cleanMonth && rowMonth !== cleanMonth) continue;
    
    if (currentUser.role === 'ACCOUNTANT') {
      const submitStatus = cols['SubmitStatus'] !== undefined ? String(row[cols['SubmitStatus']] || '').trim() : 'SUBMITTED';
      if (submitStatus !== 'SUBMITTED') continue;
    }

    const staffName = row[cols['StaffName']];
    const staffInfo = staffMap[staffName] || {};
    const status = cols['PaymentStatus'] !== undefined ? String(row[cols['PaymentStatus']] || '').trim() : 'PENDING';
    if (!status) continue;
    
    const isManagerRecord = cols['IsManagerRecord'] !== undefined && String(row[cols['IsManagerRecord']] || '').trim().toUpperCase() === 'YES';

    const basic = Number(row[cols['BasicSalary']]) || 0;
    const deduction = Number(row[cols['Deduction']]) || 0;
    const net = Number(row[cols['NetSalary']]) || 0;

    const lastMonth = lastMonthMap[staffName];
    const comparison = {
      isNew: !lastMonth,
      sameAsLast: lastMonth ? (basic === lastMonth.basic && deduction === lastMonth.deduction) : false,
      basicChanged: lastMonth ? (basic !== lastMonth.basic) : false,
      deductionChanged: lastMonth ? (deduction !== lastMonth.deduction) : false,
      basicDiff: lastMonth ? basic - lastMonth.basic : 0,
      deductionDiff: lastMonth ? deduction - lastMonth.deduction : 0,
      lastBasic: lastMonth ? lastMonth.basic : 0,
      lastDeduction: lastMonth ? lastMonth.deduction : 0,
      lastMonthNet: lastMonth ? lastMonth.net : 0
    };

    const autoDeduction = cols['AutoDeduction'] !== undefined ? Number(row[cols['AutoDeduction']]) || 0 : 0;
    
    const record = {
      rowIndex: i + 1,
      month: rowMonth,
      staffName: staffName,
      basicSalary: basic,
      deduction: deduction,
      autoDeduction: autoDeduction,
      netSalary: net,
      remark: row[cols['Remark']] || '',
      deductionReason: row[cols['Remark']] || '',  // 手动扣款原因
      createdBy: row[cols['CreatedBy']] || '',
      status: status,
      bankHolder: staffInfo.bankHolder || '',
      bankType: staffInfo.bankType || '',
      bankAccount: staffInfo.bankAccount || '',
      paymentMethod: staffInfo.paymentMethod || 'BANK',
      bankChanged: staffInfo.bankChanged || false,
      bankChangeNote: staffInfo.bankChangeNote || '',
      hasDeduction: deduction > 0 || autoDeduction > 0,
      comparison: comparison,
      isManager: isManagerRecord,
      secretary: staffInfo.secretary || '',
      // 欠款信息
      debtReason: staffInfo.debtReason || '',
      totalDebt: staffInfo.totalDebt || 0,
      monthlyDeduction: staffInfo.monthlyDeduction || 0
    };

    if (isManagerRecord) {
      if (status === 'PAID') managerPaid.push(record);
      else managerPending.push(record);
    } else {
      if (status === 'PAID') employeePaid.push(record);
      else employeePending.push(record);
    }
  }

  const allPending = [...employeePending, ...managerPending];
  const bankPending = allPending.filter(p => p.paymentMethod === 'BANK');
  const cashPending = allPending.filter(p => p.paymentMethod === 'CASH');
  
  const stats = {
    employeeCount: totalEmployeeCount,
    managerCount: totalManagerCount,
    bankTotal: bankPending.reduce((sum, r) => sum + r.netSalary, 0),
    cashTotal: cashPending.reduce((sum, r) => sum + r.netSalary, 0),
    employeePendingTotal: employeePending.reduce((sum, r) => sum + r.netSalary, 0),
    managerPendingTotal: managerPending.reduce((sum, r) => sum + r.netSalary, 0),
    staffBySecretary: staffBySecretary
  };

  const alerts = {
    bankChanges: allPending.filter(p => p.bankChanged).length,
    hasDeductions: allPending.filter(p => p.hasDeduction).length,
    salaryChanges: allPending.filter(p => !p.comparison.isNew && !p.comparison.sameAsLast).length,
    newStaff: allPending.filter(p => p.comparison.isNew).length
  };

  return {
    success: true,
    employeePending, employeePaid, managerPending, managerPaid,
    stats, alerts, prevMonth,
    pending: allPending,
    paid: [...employeePaid, ...managerPaid],
    totalPending: stats.employeePendingTotal + stats.managerPendingTotal
  };
}

/**
 * 标记已发放
 */
function markAsPaid(currentUser, rowIndexes, paymentMethod) {
  if (!currentUser || currentUser.role === 'SECRETARY') {
    return { success: false, message: '书记无权限发放工资' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_SALARY);
  if (!sheet) return { success: false, message: '工资表不存在' };
  
  if (!Array.isArray(rowIndexes)) {
    rowIndexes = [rowIndexes];
  }
  
  const validRowIndexes = rowIndexes.filter(idx => typeof idx === 'number' && idx > 1);
  if (validRowIndexes.length === 0) {
    return { success: false, message: '无有效记录' };
  }
  
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  let count = 0;
  const paidNames = [];
  
  for (const rowIndex of validRowIndexes) {
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) continue;
    
    const rowData = sheet.getRange(rowIndex, 1, 1, header.length).getValues()[0];
    
    const submitStatus = cols['SubmitStatus'] !== undefined ? String(rowData[cols['SubmitStatus']] || '').trim() : 'SUBMITTED';
    const payStatus = cols['PaymentStatus'] !== undefined ? String(rowData[cols['PaymentStatus']] || '').trim() : 'PENDING';
    const staffName = rowData[cols['StaffName']] || '';
    
    if ((submitStatus === 'SUBMITTED' || cols['SubmitStatus'] === undefined) && payStatus !== 'PAID') {
      try {
        if (cols['PaymentStatus'] !== undefined) {
          sheet.getRange(rowIndex, cols['PaymentStatus'] + 1).setValue('PAID');
        }
        if (cols['PaidAt'] !== undefined) {
          sheet.getRange(rowIndex, cols['PaidAt'] + 1).setValue(now);
        }
        if (cols['PaidBy'] !== undefined) {
          sheet.getRange(rowIndex, cols['PaidBy'] + 1).setValue(currentUser.username);
        }
        
        paidNames.push(staffName);
        count++;
      } catch (e) {
        console.log('标记失败:', staffName, '-', e.message);
      }
    }
  }
  
  if (count > 0) {
    logOperation_(count > 1 ? 'BATCH_PAID' : 'MARK_PAID', {
      count: count,
      names: paidNames.join(', '),
      paymentMethod: paymentMethod
    }, currentUser.username);
  }
  
  return { success: true, count };
}

/**
 * 获取仪表板统计
 */
function getDashboardStats(currentUser, month) {
  if (!currentUser) return { success: false };
  
  const ss = SpreadsheetApp.getActive();
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  const managerSheet = ss.getSheetByName(SHEET_MANAGERS);
  
  let staffCount = 0, managerCount = 0;
  
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const staffData = staffSheet.getDataRange().getValues();
    const staffHeader = staffData[0];
    const staffCols = {};
    staffHeader.forEach((h, i) => staffCols[String(h).trim()] = i);
    
    for (let i = 1; i < staffData.length; i++) {
      const row = staffData[i];
      if (!row[staffCols['StaffName']]) continue;
      
      const status = row[staffCols['Status']] || 'ACTIVE';
      if (status === 'LEFT') continue;
      
      if (currentUser.role === 'SECRETARY') {
        const createdBy = String(row[staffCols['CreatedBy']] || '').toLowerCase().trim();
        if (createdBy !== currentUser.username.toLowerCase()) continue;
      }
      
      staffCount++;
    }
  }
  
  if (managerSheet && managerSheet.getLastRow() > 1) {
    const mgrData = managerSheet.getDataRange().getValues();
    const mgrHeader = mgrData[0];
    const mgrCols = {};
    mgrHeader.forEach((h, i) => mgrCols[String(h).trim()] = i);
    
    for (let i = 1; i < mgrData.length; i++) {
      if (!mgrData[i][mgrCols['StaffName']]) continue;
      const status = mgrData[i][mgrCols['Status']] || 'ACTIVE';
      if (status !== 'LEFT') managerCount++;
    }
  }
  
  const salarySheet = ss.getSheetByName(SHEET_SALARY);
  let draftCount = 0, submittedCount = 0, paidCount = 0;
  let staffBankTotal = 0, staffCashTotal = 0, managerBankTotal = 0, managerCashTotal = 0;
  let bankChanges = [];
  
  let cleanMonth = month ? sanitizeInput_(month, 'string') : '';
  if (!cleanMonth) {
    cleanMonth = getCurrentMonth_();
  }
  
  if (salarySheet && salarySheet.getLastRow() > 1) {
    const data = salarySheet.getDataRange().getValues();
    const header = data[0];
    const cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowMonth = formatMonthValue_(row[cols['Month']]);
      
      if (cleanMonth && rowMonth !== cleanMonth) continue;
      
      const submitStatus = String(row[cols['SubmitStatus']] || '').trim();
      const payStatus = String(row[cols['PaymentStatus']] || '').trim();
      const createdBy = String(row[cols['CreatedBy']] || '').trim();
      
      if (currentUser.role === 'SECRETARY') {
        if (createdBy.toLowerCase() !== currentUser.username.toLowerCase()) continue;
      }
      
      const isManagerRecord = cols['IsManagerRecord'] !== undefined && 
                              String(row[cols['IsManagerRecord']] || '').trim().toUpperCase() === 'YES';
      
      if (currentUser.role === 'ADMIN' && !isManagerRecord) continue;
      
      const net = Number(row[cols['NetSalary']]) || 0;
      const pm = String(row[cols['PaymentMethod']] || 'BANK').trim();
      
      if (submitStatus === 'DRAFT') {
        draftCount++;
      } else if (submitStatus === 'SUBMITTED' && payStatus !== 'PAID') {
        submittedCount++;
        if (isManagerRecord) {
          if (pm === 'BANK') managerBankTotal += net; else managerCashTotal += net;
        } else {
          if (pm === 'BANK') staffBankTotal += net; else staffCashTotal += net;
        }
      } else if (payStatus === 'PAID') {
        paidCount++;
      }
    }
  }
  
  const staffChanges = getRecentChanges_(staffSheet);
  const managerChanges = getRecentChanges_(managerSheet);
  Object.keys(staffChanges).forEach(name => { bankChanges.push({ name, type: 'STAFF', note: staffChanges[name] }); });
  Object.keys(managerChanges).forEach(name => { bankChanges.push({ name, type: 'MANAGER', note: managerChanges[name] }); });
  
  return {
    success: true,
    stats: {
      staffCount, managerCount, draftCount, submittedCount, paidCount,
      staffBankTotal, staffCashTotal, managerBankTotal, managerCashTotal,
      totalBankAmount: staffBankTotal + managerBankTotal,
      totalCashAmount: staffCashTotal + managerCashTotal
    },
    bankChanges,
    currentMonth: cleanMonth
  };
}

/**
 * 获取近期银行变更
 */
function getRecentChanges_(sheet) {
  const changes = {};
  if (!sheet) return changes;
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const now = new Date();
  const alertCutoffDate = new Date(now.getTime() - BANK_CHANGE_ALERT_DAYS * 24 * 60 * 60 * 1000);
  
  for (let i = 1; i < data.length; i++) {
    const lastUpdate = data[i][cols['LastBankUpdate']];
    if (lastUpdate) {
      const updateDate = new Date(lastUpdate);
      if (updateDate > alertCutoffDate) {
        changes[data[i][cols['StaffName']]] = data[i][cols['BankChangeNote']] || '资料已变更';
      }
    }
  }
  return changes;
}

/***********************
 * 批量操作
 ***********************/

/**
 * 批量录入工资草稿
 */
function batchAddSalaryRecords(currentUser, month, date) {
  if (!currentUser) return { success: false, message: '未登录' };
  if (currentUser.role === 'ACCOUNTANT') return { success: false, message: '会计无权限' };
  
  month = String(month || '').trim();
  date = String(date || '').trim();
  
  if (!month || !/^\d{4}-\d{2}$/.test(month)) {
    return { success: false, message: '月份格式不正确' };
  }
  if (!date) {
    date = new Date().toISOString().split('T')[0];
  }
  
  var ss = SpreadsheetApp.getActive();
  var salarySheet = ss.getSheetByName('SalaryRecords');
  
  if (!salarySheet) {
    var headers = ['Month', 'Date', 'StaffName', 'IsManagerRecord', 'BasicSalary', 'ManualDeduction', 'AutoDeduction', 
                   'BankFee', 'Deduction', 'NetSalary', 'Remark', 'CreatedBy', 'CreatedAt', 'SubmitStatus', 'SubmittedAt',
                   'PaymentStatus', 'PaymentMethod', 'PaidAt', 'PaidBy'];
    salarySheet = ss.insertSheet('SalaryRecords');
    salarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  var staffList = [];
  var isManager = false;
  
  if (currentUser.role === 'ADMIN') {
    staffList = getManagerList(currentUser);
    isManager = true;
  } else {
    staffList = getStaffList(currentUser);
  }
  
  if (!staffList || staffList.length === 0) {
    return { success: false, message: '没有可录入的' + (isManager ? '主管' : '员工') };
  }
  
  // 检查已有记录
  var existingRecords = {};
  if (salarySheet.getLastRow() > 1) {
    var salaryData = salarySheet.getDataRange().getValues();
    var salaryHeader = salaryData[0];
    var sCols = {};
    salaryHeader.forEach(function(h, i) { sCols[String(h).trim()] = i; });
    
    for (var i = 1; i < salaryData.length; i++) {
      var rowMonth = salaryData[i][sCols['Month']];
      if (rowMonth instanceof Date) {
        rowMonth = rowMonth.getFullYear() + '-' + String(rowMonth.getMonth() + 1).padStart(2, '0');
      } else {
        rowMonth = String(rowMonth || '').trim();
      }
      var staffName = salaryData[i][sCols['StaffName']];
      if (String(rowMonth) === month && staffName) {
        existingRecords[staffName] = true;
      }
    }
  }
  
  var tz = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  var addedCount = 0;
  var skippedCount = 0;
  var skippedNames = [];
  
  for (var j = 0; j < staffList.length; j++) {
    var staff = staffList[j];
    var staffName = staff.staffName;
    
    if (existingRecords[staffName]) {
      skippedCount++;
      skippedNames.push(staffName);
      continue;
    }
    
    var status = staff.status || staff.Status || 'ACTIVE';
    if (status === 'LEFT') continue;
    
    var basicSalary = Number(staff.salary) || 0;
    if (basicSalary <= 0) {
      skippedCount++;
      skippedNames.push(staffName + '(无月薪)');
      continue;
    }
    
    var paymentMethod = staff.paymentMethod || 'BANK';
    
    var bankFee = 0;
    if (paymentMethod === 'BANK' && !isManager) {
      bankFee = 15;
    }
    
    var manualDeduction = 0;
    var autoDeduction = 0;
    var totalDeduction = manualDeduction + autoDeduction + bankFee;
    var netSalary = basicSalary - totalDeduction;
    
    var remark = '批量录入';
    if (bankFee > 0) {
      remark += '; [汇款手续费] RM' + bankFee.toFixed(2);
    }
    
    try {
      salarySheet.appendRow([
        month,
        date,
        staffName,
        isManager ? 'YES' : '',
        basicSalary,
        manualDeduction,
        autoDeduction,
        bankFee,
        totalDeduction,
        netSalary,
        remark,
        currentUser.username,
        now,
        'DRAFT',
        '',
        'PENDING',
        paymentMethod,
        '',
        ''
      ]);
      addedCount++;
    } catch (e) {
      skippedCount++;
      skippedNames.push(staffName + '(写入失败)');
    }
  }
  
  var message = '成功录入 ' + addedCount + ' 条';
  if (skippedCount > 0) {
    message += '，跳过 ' + skippedCount + ' 条';
    if (skippedNames.length <= 5) {
      message += ' (' + skippedNames.join(', ') + ')';
    }
  }
  
  logOperation_('BATCH_ADD_SALARY', {
    month: month,
    count: addedCount,
    skipped: skippedCount,
    isManager: isManager
  }, currentUser.username);
  
  return { 
    success: true, 
    message: message,
    added: addedCount,
    skipped: skippedCount,
    count: addedCount
  };
}

/**
 * 获取银行变更提醒
 */
function getBankChangeAlerts(user) {
  try {
    if (!user || (user.role !== 'ADMIN' && user.role !== 'ACCOUNTANT')) {
      return { success: false, message: '无权限' };
    }
    
    const ss = SpreadsheetApp.getActive();
    const alerts = [];
    const now = new Date();
    const alertCutoffDate = new Date(now.getTime() - BANK_CHANGE_ALERT_DAYS * 24 * 60 * 60 * 1000);
    
    // 处理 Staff 表
    const staffSheet = ss.getSheetByName('Staff');
    if (staffSheet && staffSheet.getLastRow() > 1) {
      const data = staffSheet.getDataRange().getValues();
      const headers = data[0];
      const cols = {};
      headers.forEach(function(h, i) { cols[String(h).trim()] = i; });
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        if (cols['Status'] !== undefined && row[cols['Status']] === 'LEFT') continue;
        
        const changeNote = cols['BankChangeNote'] !== undefined ? String(row[cols['BankChangeNote']] || '').trim() : '';
        if (!changeNote || changeNote.length === 0) continue;
        
        const colLastBankUpdate = cols['LastBankUpdate'] !== undefined ? cols['LastBankUpdate'] : cols['BankChangedAt'];
        if (colLastBankUpdate === undefined) continue;
        
        const lastUpdate = row[colLastBankUpdate];
        if (!lastUpdate) continue;
        
        const updateDate = new Date(lastUpdate);
        if (isNaN(updateDate.getTime()) || updateDate < alertCutoffDate) continue;
        
        alerts.push({
          type: 'STAFF',
          name: row[cols['StaffName']] || '',
          bankHolder: String(row[cols['BankHolder']] || ''),
          bankType: String(row[cols['BankType']] || ''),
          bankAccount: String(row[cols['BankAccount']] || ''),
          oldBankHolder: cols['OldBankHolder'] !== undefined ? String(row[cols['OldBankHolder']] || '') : '',
          oldBankType: cols['OldBankType'] !== undefined ? String(row[cols['OldBankType']] || '') : '',
          oldBankAccount: cols['OldBankAccount'] !== undefined ? String(row[cols['OldBankAccount']] || '') : '',
          changeNote: changeNote,
          changedAt: Utilities.formatDate(updateDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
          secretary: cols['CreatedBy'] !== undefined ? String(row[cols['CreatedBy']] || '') : ''
        });
      }
    }
    
    // 处理 Managers 表
    const mgrSheet = ss.getSheetByName('Managers');
    if (mgrSheet && mgrSheet.getLastRow() > 1) {
      const data = mgrSheet.getDataRange().getValues();
      const headers = data[0];
      const cols = {};
      headers.forEach(function(h, i) { cols[String(h).trim()] = i; });
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        if (cols['Status'] !== undefined && row[cols['Status']] === 'LEFT') continue;
        
        const changeNote = cols['BankChangeNote'] !== undefined ? String(row[cols['BankChangeNote']] || '').trim() : '';
        if (!changeNote || changeNote.length === 0) continue;
        
        const colLastBankUpdate = cols['LastBankUpdate'] !== undefined ? cols['LastBankUpdate'] : cols['BankChangedAt'];
        if (colLastBankUpdate === undefined) continue;
        
        const lastUpdate = row[colLastBankUpdate];
        if (!lastUpdate) continue;
        
        const updateDate = new Date(lastUpdate);
        if (isNaN(updateDate.getTime()) || updateDate < alertCutoffDate) continue;
        
        alerts.push({
          type: 'MANAGER',
          name: row[cols['StaffName']] || '',
          bankHolder: String(row[cols['BankHolder']] || ''),
          bankType: String(row[cols['BankType']] || ''),
          bankAccount: String(row[cols['BankAccount']] || ''),
          oldBankHolder: cols['OldBankHolder'] !== undefined ? String(row[cols['OldBankHolder']] || '') : '',
          oldBankType: cols['OldBankType'] !== undefined ? String(row[cols['OldBankType']] || '') : '',
          oldBankAccount: cols['OldBankAccount'] !== undefined ? String(row[cols['OldBankAccount']] || '') : '',
          changeNote: changeNote,
          changedAt: Utilities.formatDate(updateDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
        });
      }
    }
    
    return { success: true, alerts: alerts, count: alerts.length };
    
  } catch (e) {
    Logger.log('getBankChangeAlerts error: ' + e.stack);
    return { success: false, message: e.message };
  }
}

/**
 * 获取欠款提醒
 */
function getDebtAlerts(currentUser) {
  if (!currentUser) return { success: false };
  
  const ss = SpreadsheetApp.getActive();
  const alerts = [];
  
  if (currentUser.role !== 'ADMIN') {
    const staffSheet = ss.getSheetByName(SHEET_STAFF);
    if (staffSheet && staffSheet.getLastRow() > 1) {
      const data = staffSheet.getDataRange().getValues();
      const header = data[0];
      const cols = {};
      header.forEach((h, i) => cols[String(h).trim()] = i);
      
      for (let i = 1; i < data.length; i++) {
        if (currentUser.role === 'SECRETARY') {
          const createdBy = String(data[i][cols['CreatedBy']] || '').toLowerCase();
          if (createdBy !== currentUser.username.toLowerCase()) continue;
        }
        
        const totalDebt = Number(data[i][cols['TotalDebt']]) || 0;
        const debtPaid = Number(data[i][cols['DebtPaid']]) || 0;
        const remaining = totalDebt - debtPaid;
        
        if (remaining > 0) {
          alerts.push({
            type: 'STAFF',
            name: data[i][cols['StaffName']],
            totalDebt: totalDebt,
            debtPaid: debtPaid,
            remaining: remaining,
            monthlyDeduction: Number(data[i][cols['MonthlyDeduction']]) || 0,
            reason: data[i][cols['DebtReason']] || ''
          });
        }
      }
    }
  }
  
  if (currentUser.role === 'ADMIN') {
    const mgrSheet = ss.getSheetByName(SHEET_MANAGERS);
    if (mgrSheet && mgrSheet.getLastRow() > 1) {
      const data = mgrSheet.getDataRange().getValues();
      const header = data[0];
      const cols = {};
      header.forEach((h, i) => cols[String(h).trim()] = i);
      
      for (let i = 1; i < data.length; i++) {
        const totalDebt = Number(data[i][cols['TotalDebt']]) || 0;
        const debtPaid = Number(data[i][cols['DebtPaid']]) || 0;
        const remaining = totalDebt - debtPaid;
        
        if (remaining > 0) {
          alerts.push({
            type: 'MANAGER',
            name: data[i][cols['StaffName']],
            totalDebt: totalDebt,
            debtPaid: debtPaid,
            remaining: remaining,
            monthlyDeduction: Number(data[i][cols['MonthlyDeduction']]) || 0,
            reason: data[i][cols['DebtReason']] || ''
          });
        }
      }
    }
  }
  
  const totalRemaining = alerts.reduce((sum, a) => sum + a.remaining, 0);
  
  return { 
    success: true, 
    alerts: alerts, 
    count: alerts.length,
    totalRemaining: totalRemaining
  };
}

/**
 * 获取新员工提醒（近30天入职）
 */
function getNewStaffAlerts(currentUser) {
  if (!currentUser) return { success: false };
  if (currentUser.role !== 'ADMIN' && currentUser.role !== 'ACCOUNTANT') {
    return { success: false, message: '无权限' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const alerts = [];
  const now = new Date();
  const thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const tz = Session.getScriptTimeZone();
  
  // 获取员工新入职
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const data = staffSheet.getDataRange().getValues();
    const header = data[0];
    const cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][cols['Status']] || '').toUpperCase();
      if (status === 'LEFT') continue;
      
      const createdAt = data[i][cols['CreatedAt']];
      if (!createdAt) continue;
      
      let createdDate;
      try {
        createdDate = new Date(createdAt);
        if (isNaN(createdDate.getTime())) continue;
      } catch (e) {
        continue;
      }
      
      if (createdDate >= thirtyDaysAgo) {
        const daysAgo = Math.floor((now - createdDate) / (24 * 60 * 60 * 1000));
        alerts.push({
          type: 'STAFF',
          name: data[i][cols['StaffName']] || '',
          company: data[i][cols['CompanyName']] || '',
          salary: Number(data[i][cols['Salary']]) || 0,
          bankHolder: data[i][cols['BankHolder']] || '',
          bankType: data[i][cols['BankType']] || '',
          bankAccount: data[i][cols['BankAccount']] || '',
          paymentMethod: data[i][cols['PaymentMethod']] || 'BANK',
          createdBy: data[i][cols['CreatedBy']] || '',
          createdAt: Utilities.formatDate(createdDate, tz, 'yyyy-MM-dd'),
          daysAgo: daysAgo
        });
      }
    }
  }
  
  // 获取主管新入职
  const mgrSheet = ss.getSheetByName(SHEET_MANAGERS);
  if (mgrSheet && mgrSheet.getLastRow() > 1) {
    const data = mgrSheet.getDataRange().getValues();
    const header = data[0];
    const cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][cols['Status']] || '').toUpperCase();
      if (status === 'LEFT') continue;
      
      const createdAt = data[i][cols['CreatedAt']];
      if (!createdAt) continue;
      
      let createdDate;
      try {
        createdDate = new Date(createdAt);
        if (isNaN(createdDate.getTime())) continue;
      } catch (e) {
        continue;
      }
      
      if (createdDate >= thirtyDaysAgo) {
        const daysAgo = Math.floor((now - createdDate) / (24 * 60 * 60 * 1000));
        alerts.push({
          type: 'MANAGER',
          name: data[i][cols['StaffName']] || '',
          company: data[i][cols['CompanyName']] || '',
          salary: Number(data[i][cols['Salary']]) || 0,
          bankType: data[i][cols['BankType']] || '',
          bankAccount: data[i][cols['BankAccount']] || '',
          paymentMethod: data[i][cols['PaymentMethod']] || 'BANK',
          createdAt: Utilities.formatDate(createdDate, tz, 'yyyy-MM-dd'),
          daysAgo: daysAgo
        });
      }
    }
  }
  
  // 按入职日期排序（最新的在前）
  alerts.sort((a, b) => a.daysAgo - b.daysAgo);
  
  return { 
    success: true, 
    alerts: alerts, 
    count: alerts.length
  };
}

/**
 * 获取离职员工提醒（近30天离职）
 */
function getLeftStaffAlerts(currentUser) {
  if (!currentUser) return { success: false };
  if (currentUser.role !== 'ADMIN' && currentUser.role !== 'ACCOUNTANT') {
    return { success: false, message: '无权限' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const alerts = [];
  const now = new Date();
  const thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const tz = Session.getScriptTimeZone();
  
  // 获取离职员工
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const data = staffSheet.getDataRange().getValues();
    const header = data[0];
    const cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][cols['Status']] || '').toUpperCase();
      if (status !== 'LEFT') continue;
      
      const leaveDate = data[i][cols['LeaveDate']] || data[i][cols['UpdatedAt']];
      if (!leaveDate) continue;
      
      let leftDate;
      try {
        leftDate = new Date(leaveDate);
        if (isNaN(leftDate.getTime())) continue;
      } catch (e) {
        continue;
      }
      
      if (leftDate >= thirtyDaysAgo) {
        const daysAgo = Math.floor((now - leftDate) / (24 * 60 * 60 * 1000));
        alerts.push({
          type: 'STAFF',
          name: data[i][cols['StaffName']] || '',
          company: data[i][cols['CompanyName']] || '',
          salary: Number(data[i][cols['Salary']]) || 0,
          bankHolder: data[i][cols['BankHolder']] || '', 
          bankType: data[i][cols['BankType']] || '',
          bankAccount: data[i][cols['BankAccount']] || '',
          paymentMethod: data[i][cols['PaymentMethod']] || 'BANK',
          createdBy: data[i][cols['CreatedBy']] || '',
          leftAt: Utilities.formatDate(leftDate, tz, 'yyyy-MM-dd'),
          daysAgo: daysAgo
        });
      }
    }
  }
  
  // 获取离职主管
  const mgrSheet = ss.getSheetByName(SHEET_MANAGERS);
  if (mgrSheet && mgrSheet.getLastRow() > 1) {
    const data = mgrSheet.getDataRange().getValues();
    const header = data[0];
    const cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][cols['Status']] || '').toUpperCase();
      if (status !== 'LEFT') continue;
      
      const leaveDate = data[i][cols['LeaveDate']] || data[i][cols['UpdatedAt']];
      if (!leaveDate) continue;
      
      let leftDate;
      try {
        leftDate = new Date(leaveDate);
        if (isNaN(leftDate.getTime())) continue;
      } catch (e) {
        continue;
      }
      
      if (leftDate >= thirtyDaysAgo) {
        const daysAgo = Math.floor((now - leftDate) / (24 * 60 * 60 * 1000));
        alerts.push({
          type: 'MANAGER',
          name: data[i][cols['StaffName']] || '',
          company: data[i][cols['CompanyName']] || '',
          salary: Number(data[i][cols['Salary']]) || 0,
          bankHolder: data[i][cols['BankHolder']] || '',
          bankType: data[i][cols['BankType']] || '',
          bankAccount: data[i][cols['BankAccount']] || '',
          paymentMethod: data[i][cols['PaymentMethod']] || 'BANK',
          leftAt: Utilities.formatDate(leftDate, tz, 'yyyy-MM-dd'),
          daysAgo: daysAgo
        });
      }
    }
  }
  
  // 按离职日期排序（最新的在前）
  alerts.sort((a, b) => a.daysAgo - b.daysAgo);
  
  return { 
    success: true, 
    alerts: alerts, 
    count: alerts.length
  };
}

/**
 * 获取起薪通知（本月已创建的工资记录，通知会计有人被起薪）
 */
function getSalaryNotifyAlerts(currentUser, month) {
  if (!currentUser) return { success: false };
  if (currentUser.role !== 'ADMIN' && currentUser.role !== 'ACCOUNTANT') {
    return { success: false, message: '无权限' };
  }
  
  const cleanMonth = sanitizeInput_(month, 'string') || getCurrentMonth_();
  
  const ss = SpreadsheetApp.getActive();
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  const mgrSheet = ss.getSheetByName(SHEET_MANAGERS);
  
  const now = new Date();
  const thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const tz = Session.getScriptTimeZone();
  const list = [];
  
  // 检查员工调薪记录
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const data = staffSheet.getDataRange().getValues();
    const header = data[0];
    const cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const changeDate = data[i][cols['SalaryChangeDate']];
      if (!changeDate) continue;
      
      let changeDateObj;
      try {
        changeDateObj = new Date(changeDate);
        if (isNaN(changeDateObj.getTime())) continue;
      } catch (e) {
        continue;
      }
      
      // 只显示30天内的调薪记录
      if (changeDateObj < thirtyDaysAgo) continue;
      
      const daysAgo = Math.floor((now - changeDateObj) / (24 * 60 * 60 * 1000));
      const staffName = data[i][cols['StaffName']] || '';
      const oldSalary = Number(data[i][cols['OldSalary']]) || 0;
      const newSalary = Number(data[i][cols['Salary']]) || 0;
      const changeNote = data[i][cols['SalaryChangeNote']] || '';
      
      if (!staffName || oldSalary === 0) continue;
      
      const diff = newSalary - oldSalary;
      const percent = oldSalary > 0 ? ((diff / oldSalary) * 100).toFixed(1) : 0;
      
      list.push({
        staffName: staffName,
        isManager: false,
        oldSalary: oldSalary,
        newSalary: newSalary,
        diff: diff,
        percent: percent,
        changeNote: changeNote,
        changeDate: Utilities.formatDate(changeDateObj, tz, 'yyyy-MM-dd'),
        daysAgo: daysAgo
      });
    }
  }
  
  // 检查主管调薪记录
  if (mgrSheet && mgrSheet.getLastRow() > 1) {
    const data = mgrSheet.getDataRange().getValues();
    const header = data[0];
    const cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const changeDate = data[i][cols['SalaryChangeDate']];
      if (!changeDate) continue;
      
      let changeDateObj;
      try {
        changeDateObj = new Date(changeDate);
        if (isNaN(changeDateObj.getTime())) continue;
      } catch (e) {
        continue;
      }
      
      // 只显示30天内的调薪记录
      if (changeDateObj < thirtyDaysAgo) continue;
      
      const daysAgo = Math.floor((now - changeDateObj) / (24 * 60 * 60 * 1000));
      const staffName = data[i][cols['StaffName']] || '';
      const oldSalary = Number(data[i][cols['OldSalary']]) || 0;
      const newSalary = Number(data[i][cols['Salary']]) || 0;
      const changeNote = data[i][cols['SalaryChangeNote']] || '';
      
      if (!staffName || oldSalary === 0) continue;
      
      const diff = newSalary - oldSalary;
      const percent = oldSalary > 0 ? ((diff / oldSalary) * 100).toFixed(1) : 0;
      
      list.push({
        staffName: staffName,
        isManager: true,
        oldSalary: oldSalary,
        newSalary: newSalary,
        diff: diff,
        percent: percent,
        changeNote: changeNote,
        changeDate: Utilities.formatDate(changeDateObj, tz, 'yyyy-MM-dd'),
        daysAgo: daysAgo
      });
    }
  }
  
  // 按调薪时间排序（最新的在前）
  list.sort((a, b) => a.daysAgo - b.daysAgo);
  
  return { 
    success: true, 
    list: list, 
    count: list.length,
    month: cleanMonth
  };
}

/**
 * 获取银行汇款汇总
 */
function getBankTransferSummary(currentUser, month) {
  if (!currentUser) return { success: false };
  if (currentUser.role === 'SECRETARY') return { success: false, message: '无权限' };
  
  const cleanMonth = sanitizeInput_(month, 'string');
  
  const ss = SpreadsheetApp.getActive();
  const salarySheet = ss.getSheetByName(SHEET_SALARY);
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  const mgrSheet = ss.getSheetByName(SHEET_MANAGERS);
  
  if (!salarySheet) return { success: false, message: '工资表不存在' };
  
  const bankInfo = {};
  
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const data = staffSheet.getDataRange().getValues();
    const header = data[0];
    const cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const name = data[i][cols['StaffName']];
      if (name) {
        bankInfo[name] = {
          bankHolder: data[i][cols['BankHolder']] || '',
          bankType: data[i][cols['BankType']] || '',
          bankAccount: data[i][cols['BankAccount']] || '',
          isManager: false
        };
      }
    }
  }
  
  if (mgrSheet && mgrSheet.getLastRow() > 1) {
    const data = mgrSheet.getDataRange().getValues();
    const header = data[0];
    const cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const name = data[i][cols['StaffName']];
      if (name) {
        bankInfo[name] = {
          bankHolder: data[i][cols['BankHolder']] || '',
          bankType: data[i][cols['BankType']] || '',
          bankAccount: data[i][cols['BankAccount']] || '',
          isManager: true
        };
      }
    }
  }
  
  const salaryData = salarySheet.getDataRange().getValues();
  const salaryHeader = salaryData[0];
  const cols = {};
  salaryHeader.forEach((h, i) => cols[String(h).trim()] = i);
  
  const transfers = [];
  
  for (let i = 1; i < salaryData.length; i++) {
    const row = salaryData[i];
    const submitStatus = String(row[cols['SubmitStatus']] || '').trim();
    const payStatus = String(row[cols['PaymentStatus']] || '').trim();
    const method = String(row[cols['PaymentMethod']] || '').trim();
    
    if (submitStatus !== 'SUBMITTED' || payStatus === 'PAID' || method !== 'BANK') continue;
    
    const rowMonth = formatMonthValue_(row[cols['Month']]);
    if (cleanMonth && rowMonth !== cleanMonth) continue;
    
    const name = row[cols['StaffName']];
    const info = bankInfo[name] || {};
    const isManagerRecord = cols['IsManagerRecord'] !== undefined && 
                            String(row[cols['IsManagerRecord']] || '').trim().toUpperCase() === 'YES';
    
    transfers.push({
      staffName: name,
      bankHolder: info.bankHolder || name,
      bankType: info.bankType || '',
      bankAccount: info.bankAccount || '',
      amount: Number(row[cols['NetSalary']]) || 0,
      isManager: isManagerRecord
    });
  }
  
  const byBank = {};
  transfers.forEach(t => {
    const bank = t.bankType || '未知银行';
    if (!byBank[bank]) byBank[bank] = [];
    byBank[bank].push(t);
  });
  
  return {
    success: true,
    transfers: transfers,
    byBank: byBank,
    totalAmount: transfers.reduce((sum, t) => sum + t.amount, 0),
    totalCount: transfers.length
  };
}

/**
 * 更新欠款设置
 */
function updateDebtSettings(currentUser, staffName, type, debtInfo) {
  if (!currentUser) return { success: false, message: '未登录' };
  
  const cleanName = sanitizeInput_(staffName, 'name');
  
  if (type === 'MANAGER' && currentUser.role !== 'ADMIN') {
    return { success: false, message: '只有管理员可以修改主管欠款' };
  }
  if (type === 'STAFF' && currentUser.role === 'ACCOUNTANT') {
    return { success: false, message: '会计无权限修改' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(type === 'MANAGER' ? SHEET_MANAGERS : SHEET_STAFF);
  if (!sheet) return { success: false, message: '表格不存在' };
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][cols['StaffName']] === cleanName) {
      const rowIndex = i + 1;
      if (debtInfo.totalDebt !== undefined) sheet.getRange(rowIndex, cols['TotalDebt'] + 1).setValue(Number(debtInfo.totalDebt) || 0);
      if (debtInfo.monthlyDeduction !== undefined) sheet.getRange(rowIndex, cols['MonthlyDeduction'] + 1).setValue(Number(debtInfo.monthlyDeduction) || 0);
      if (debtInfo.resetDebtPaid) sheet.getRange(rowIndex, cols['DebtPaid'] + 1).setValue(0);
      if (debtInfo.debtReason !== undefined) sheet.getRange(rowIndex, cols['DebtReason'] + 1).setValue(sanitizeInput_(debtInfo.debtReason, 'string'));
      
      logOperation_('UPDATE_DEBT', {
        staffName: cleanName,
        type: type,
        debtInfo: debtInfo
      }, currentUser.username);
      
      return { success: true };
    }
  }
  
  return { success: false, message: '找不到人员' };
}

/**
 * 工资预览计算
 */
function previewSalaryCalculation(staffName, month, inputBasic) {
  const cleanName = sanitizeInput_(staffName, 'name');
  const cleanMonth = sanitizeInput_(month, 'string');
  
  if (!cleanName) return { success: false, message: '请选择员工' };
  
  const ss = SpreadsheetApp.getActive();
  
  let sheet = ss.getSheetByName(SHEET_STAFF);
  let found = false;
  let isStaff = true;
  let data, header, cols;
  
  if (sheet && sheet.getLastRow() > 1) {
    data = sheet.getDataRange().getValues();
    header = data[0];
    cols = {};
    header.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][cols['StaffName']] === cleanName) {
        found = true;
        break;
      }
    }
  }
  
  if (!found) {
    sheet = ss.getSheetByName(SHEET_MANAGERS);
    isStaff = false;
    if (sheet && sheet.getLastRow() > 1) {
      data = sheet.getDataRange().getValues();
      header = data[0];
      cols = {};
      header.forEach((h, i) => cols[String(h).trim()] = i);
    }
  }
  
  if (!sheet) return { success: false };
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][cols['StaffName']] !== cleanName) continue;
    
    const staffSalary = Number(data[i][cols['Salary']]) || 0;
    const joinDate = data[i][cols['JoinDate']];
    const leaveDate = data[i][cols['LeaveDate']];
    const totalDebt = Number(data[i][cols['TotalDebt']]) || 0;
    const monthlyDed = Number(data[i][cols['MonthlyDeduction']]) || 0;
    const debtPaid = Number(data[i][cols['DebtPaid']]) || 0;
    const debtReason = data[i][cols['DebtReason']] || '';
    const paymentMethod = data[i][cols['PaymentMethod']] || 'BANK';
    
    let calculatedBasic = inputBasic > 0 ? inputBasic : staffSalary;
    let proRataInfo = '';
    let autoDeduction = 0;
    
    const parsed = parseMonth_(cleanMonth);
    if (parsed.valid) {
      const year = parsed.year;
      const mon = parsed.month;
      const daysInMonth = new Date(year, mon, 0).getDate();
      let workDays = daysInMonth;
      
      if (joinDate) {
        const jd = new Date(joinDate);
        if (jd.getFullYear() === year && jd.getMonth() + 1 === mon) {
          const joinDay = jd.getDate();
          workDays = daysInMonth - joinDay + 1;
          proRataInfo = '入职' + mon + '月' + joinDay + '日, 工作' + workDays + '/' + daysInMonth + '天';
        }
      }
      
      if (leaveDate) {
        const ld = new Date(leaveDate);
        if (ld.getFullYear() === year && ld.getMonth() + 1 === mon) {
          const leaveDay = ld.getDate();
          if (joinDate) {
            const jd = new Date(joinDate);
            if (jd.getFullYear() === year && jd.getMonth() + 1 === mon) {
              workDays = leaveDay - jd.getDate() + 1;
            } else {
              workDays = leaveDay;
            }
          } else {
            workDays = leaveDay;
          }
          proRataInfo = '离职' + mon + '月' + leaveDay + '日, 工作' + workDays + '/' + daysInMonth + '天';
        }
      }
      
      if (workDays < daysInMonth && workDays > 0) {
        calculatedBasic = Math.round((staffSalary / daysInMonth) * workDays * 100) / 100;
      }
    }
    
    const debtRemaining = totalDebt - debtPaid;
    let debtInfo = '';
    if (monthlyDed > 0 && debtRemaining > 0) {
      autoDeduction = Math.min(monthlyDed, debtRemaining);
      debtInfo = (debtReason || '分期扣款') + ': 本月扣 RM' + autoDeduction.toFixed(2) + ', 剩余 RM' + (debtRemaining - autoDeduction).toFixed(2);
    }
    
    var bankFee = 0;
    if (paymentMethod === 'BANK' && isStaff) {
      bankFee = 15;
    }

    return {
      success: true,
      staffSalary: staffSalary,
      calculatedBasic: calculatedBasic,
      proRataInfo: proRataInfo,
      autoDeduction: autoDeduction,
      debtRemaining: debtRemaining,
      debtReason: debtReason,
      debtInfo: debtInfo,
      bankFee: bankFee,
      bankFeeInfo: bankFee > 0 ? '银行汇款手续费 RM' + bankFee.toFixed(2) : ''
    };
  }
  
  return { success: false, message: '找不到该员工' };
}

/***********************
 * 前端兼容别名函数
 * 这些函数用于匹配前端调用的函数名
 ***********************/

// ========== 草稿管理 ==========

/**
 * 获取草稿列表（前端调用名）
 */
function getDrafts(currentUser, month, isManager) {
  const result = getDraftRecords(currentUser, month);
  return result;
}

/**
 * 更新草稿（前端调用名）
 */
function updateDraft(currentUser, rowIndex, basic, deduction, remark, isManager) {
  return updateDraftRecord(currentUser, rowIndex, {
    basicSalary: basic,
    deduction: deduction,
    remark: remark
  });
}

/**
 * 删除草稿（前端调用名）
 */
function deleteDraft(currentUser, rowIndex, isManager) {
  return deleteDraftRecord(currentUser, rowIndex);
}

/**
 * 批量删除草稿（前端调用名）
 */
function deleteMultipleDrafts(currentUser, rowIndexes, isManager) {
  if (!currentUser) return { success: false, message: '未登录' };
  if (!Array.isArray(rowIndexes) || rowIndexes.length === 0) {
    return { success: false, message: '请选择要删除的草稿' };
  }
  
  // 从大到小排序，避免删除后行号变化
  rowIndexes.sort((a, b) => b - a);
  
  let count = 0;
  let errors = [];
  
  for (const idx of rowIndexes) {
    const result = deleteDraftRecord(currentUser, idx);
    if (result.success) {
      count++;
    } else {
      errors.push(result.message);
    }
  }
  
  return { 
    success: count > 0, 
    count: count,
    message: count > 0 ? '成功删除 ' + count + ' 条' : errors[0] || '删除失败'
  };
}

/**
 * 提交所有草稿（前端调用名）
 */
function submitAllDrafts(currentUser, month, isManager) {
  return submitSalaryRecords(currentUser, month);
}

// ========== 出粮管理 ==========

/**
 * 获取员工待发工资（前端调用名）
 */
function getEmployeePending(currentUser, month) {
  const result = getPendingPayments(currentUser, month);
  if (!result.success) return result;
  
  // 添加 lastMonthNet 字段用于前端对比
  const pending = (result.employeePending || []).map(r => ({
    ...r,
    lastMonthNet: r.comparison ? r.comparison.lastMonthNet : 0
  }));
  
  return { 
    success: true, 
    pending: pending,
    paid: result.employeePaid || [],
    stats: result.stats
  };
}

/**
 * 获取主管待发工资（前端调用名）
 */
function getManagerPending(currentUser, month) {
  const result = getPendingPayments(currentUser, month);
  if (!result.success) return result;
  
  const pending = (result.managerPending || []).map(r => ({
    ...r,
    lastMonthNet: r.comparison ? r.comparison.lastMonthNet : 0
  }));
  
  return { 
    success: true, 
    pending: pending,
    paid: result.managerPaid || [],
    stats: result.stats
  };
}

/**
 * 标记单条发放（前端调用名）
 */
function markPaid(currentUser, type, rowIndex, method) {
  return markAsPaid(currentUser, [rowIndex], method);
}

/**
 * 批量标记发放（前端调用名）
 */
function markMultiplePaid(currentUser, type, rowIndexes, method) {
  return markAsPaid(currentUser, rowIndexes, method);
}

// ========== 用户管理 ==========

/**
 * 获取用户列表（前端调用名）
 */
function getUsers(currentUser) {
  if (!currentUser || currentUser.role !== 'ADMIN') {
    return { success: false, message: '无权限' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, users: [] };
  }
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    users.push({
      username: row[cols['Username']] || row[1] || '',
      role: row[cols['Role']] || row[3] || 'SECRETARY',
      status: row[cols['Status']] || row[4] || 'ACTIVE',
      displayName: row[cols['DisplayName']] || row[5] || row[1] || ''
    });
  }
  
  return { success: true, users: users };
}

/**
 * 创建用户（前端调用名）
 */
function createUser(currentUser, userData) {
  return addUserAccount(currentUser, userData);
}

/**
 * 管理员重置用户密码
 */
function adminResetPassword(currentUser, targetUsername, newPassword) {
  if (!currentUser || currentUser.role !== 'ADMIN') {
    return { success: false, message: '只有管理员可以重置密码' };
  }
  
  const cleanUsername = sanitizeInput_(targetUsername, 'username');
  const cleanPassword = sanitizeInput_(newPassword, 'string');
  
  if (!cleanUsername) {
    return { success: false, message: '用户名无效' };
  }
  
  if (!cleanPassword || cleanPassword.length < 6) {
    return { success: false, message: '密码至少需要6位' };
  }
  
  // 不允许重置自己的密码（应该用修改密码功能）
  if (cleanUsername.toLowerCase() === currentUser.username.toLowerCase()) {
    return { success: false, message: '请使用修改密码功能来更改自己的密码' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, message: '用户表不存在' };
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cols['Username']] || '').toLowerCase() === cleanUsername.toLowerCase()) {
      // 简单哈希（生产环境应使用更强的加密）
      const hashedPassword = Utilities.base64Encode(
        Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cleanPassword + cleanUsername.toLowerCase())
      );
      
      sheet.getRange(i + 1, cols['Password'] + 1).setValue(hashedPassword);
      
      logOperation_('RESET_PASSWORD', {
        targetUsername: cleanUsername,
        resetBy: currentUser.username
      }, currentUser.username);
      
      return { success: true, message: '密码已重置' };
    }
  }
  
  return { success: false, message: '找不到该用户' };
}

/**
 * 停用用户账号
 */
function disableUserAccount(currentUser, targetUsername) {
  if (!currentUser || currentUser.role !== 'ADMIN') {
    return { success: false, message: '只有管理员可以停用账号' };
  }
  
  const cleanUsername = sanitizeInput_(targetUsername, 'username');
  
  if (!cleanUsername) {
    return { success: false, message: '用户名无效' };
  }
  
  // 不允许停用自己
  if (cleanUsername.toLowerCase() === currentUser.username.toLowerCase()) {
    return { success: false, message: '不能停用自己的账号' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, message: '用户表不存在' };
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  // 检查是否有Status列，没有则添加
  if (cols['Status'] === undefined) {
    const lastCol = header.length;
    sheet.getRange(1, lastCol + 1).setValue('Status');
    cols['Status'] = lastCol;
  }
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cols['Username']] || '').toLowerCase() === cleanUsername.toLowerCase()) {
      sheet.getRange(i + 1, cols['Status'] + 1).setValue('DISABLED');
      
      logOperation_('DISABLE_USER', {
        targetUsername: cleanUsername,
        disabledBy: currentUser.username
      }, currentUser.username);
      
      return { success: true, message: '账号已停用' };
    }
  }
  
  return { success: false, message: '找不到该用户' };
}

/**
 * 启用用户账号
 */
function enableUserAccount(currentUser, targetUsername) {
  if (!currentUser || currentUser.role !== 'ADMIN') {
    return { success: false, message: '只有管理员可以启用账号' };
  }
  
  const cleanUsername = sanitizeInput_(targetUsername, 'username');
  
  if (!cleanUsername) {
    return { success: false, message: '用户名无效' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, message: '用户表不存在' };
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  if (cols['Status'] === undefined) {
    return { success: false, message: '该账号状态正常，无需启用' };
  }
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cols['Username']] || '').toLowerCase() === cleanUsername.toLowerCase()) {
      sheet.getRange(i + 1, cols['Status'] + 1).setValue('ACTIVE');
      
      logOperation_('ENABLE_USER', {
        targetUsername: cleanUsername,
        enabledBy: currentUser.username
      }, currentUser.username);
      
      return { success: true, message: '账号已启用' };
    }
  }
  
  return { success: false, message: '找不到该用户' };
}

// ========== 日志管理 ==========

/**
 * 获取日志（前端调用名）
 */
function getLogs(currentUser, action, search) {
  return getOperationLogs(currentUser, { 
    action: action || '', 
    targetName: search || '',
    limit: 100 
  });
}

/**
 * 获取当前用户的历史记录（所有角色可用）
 */
function getMyHistory(currentUser, action, search) {
  if (!currentUser) {
    return { success: false, message: '未登录' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_PAYMENTS);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, logs: [], message: '表格为空或不存在' };
  }
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  
  // 构建列索引映射（忽略大小写和空格）
  const cols = {};
  header.forEach((h, i) => {
    const key = String(h).trim();
    cols[key] = i;
    cols[key.toLowerCase()] = i;
  });
  
  // 获取列索引的辅助函数
  function getColIndex(name) {
    return cols[name] !== undefined ? cols[name] : 
           cols[name.toLowerCase()] !== undefined ? cols[name.toLowerCase()] : -1;
  }
  
  const colLogID = getColIndex('LogID');
  const colTimestamp = getColIndex('Timestamp');
  const colAction = getColIndex('Action');
  const colActionName = getColIndex('ActionName');
  const colOperator = getColIndex('Operator');
  const colTargetName = getColIndex('TargetName');
  const colDetails = getColIndex('Details');
  
  const logs = [];
  const limit = 100;
  const username = String(currentUser.username || '').toLowerCase();
  const cleanSearch = search ? sanitizeInput_(search, 'string').toLowerCase() : '';
  
  for (let i = data.length - 1; i >= 1 && logs.length < limit; i--) {
    const row = data[i];
    
    // 如果第一列为空，跳过
    if (!row[0] && !row[1]) continue;
    
    const operator = colOperator >= 0 ? String(row[colOperator] || '').toLowerCase() : '';
    const targetName = colTargetName >= 0 ? String(row[colTargetName] || '') : '';
    const details = colDetails >= 0 ? String(row[colDetails] || '') : '';
    const rowAction = colAction >= 0 ? String(row[colAction] || '') : '';
    const timestamp = colTimestamp >= 0 ? String(row[colTimestamp] || '') : '';
    
    // 根据角色过滤
    let canSee = false;
    if (currentUser.role === 'ADMIN') {
      canSee = true;
    } else if (currentUser.role === 'ACCOUNTANT') {
      canSee = (rowAction === 'MARK_PAID' || operator === username);
    } else {
      canSee = (operator === username);
    }
    
    if (!canSee) continue;
    
    // 操作类型过滤
    if (action && rowAction !== action) continue;
    
    // 搜索过滤
    if (cleanSearch) {
      const searchText = (targetName + ' ' + details + ' ' + operator + ' ' + rowAction).toLowerCase();
      if (!searchText.includes(cleanSearch)) continue;
    }
    
    logs.push({
      logId: colLogID >= 0 ? String(row[colLogID] || '') : '',
      timestamp: timestamp,
      action: rowAction,
      actionName: colActionName >= 0 ? String(row[colActionName] || rowAction) : rowAction,
      operator: colOperator >= 0 ? String(row[colOperator] || '') : '',
      target: targetName,
      details: details
    });
  }
  
  return { success: true, logs: logs, count: logs.length };
}

// ========== 银行汇总 ==========

/**
 * 获取银行汇款汇总（前端调用名）
 */
function getBankSummary(currentUser, month) {
  const result = getBankTransferSummary(currentUser, month);
  if (!result.success) return result;
  
  // 转换为前端期望的格式
  const groups = [];
  for (const [bankType, items] of Object.entries(result.byBank || {})) {
    groups.push({
      bankType: bankType,
      items: items.map(t => ({
        name: t.staffName,
        bankHolder: t.bankHolder,
        bankAccount: t.bankAccount,
        amount: t.amount,
        isManager: t.isManager || false
      })),
      total: items.reduce((sum, t) => sum + t.amount, 0)
    });
  }
  
  return {
    success: true,
    groups: groups,
    grandTotal: result.totalAmount || 0
  };
}

// ========== 批量录入 ==========

/**
 * 批量创建工资草稿（前端调用名）
 */
function batchCreateSalaryDrafts(currentUser, month, isManager) {
  const today = new Date().toISOString().split('T')[0];
  return batchAddSalaryRecords(currentUser, month, today);
}

// ========== 其他辅助函数 ==========

/**
 * 获取完整员工列表（含详细信息）
 */
function getStaffListFull(currentUser) {
  if (!currentUser) return { success: false, message: '未登录' };
  
  const list = getStaffList(currentUser);
  
  return { 
    success: true, 
    list: list,
    total: list.length,
    activeCount: list.filter(s => s.status === 'ACTIVE').length,
    leftCount: list.filter(s => s.status === 'LEFT').length,
    bankChangeCount: list.filter(s => s.bankChanged).length,
    debtCount: list.filter(s => s.debtRemaining > 0).length
  };
}

/**
 * 获取完整主管列表（含详细信息）
 */
function getManagerListFull(currentUser) {
  if (!currentUser || currentUser.role === 'SECRETARY') {
    return { success: false, message: '无权限' };
  }
  
  const list = getManagerList(currentUser);
  
  return { 
    success: true, 
    list: list,
    total: list.length,
    activeCount: list.filter(s => s.status === 'ACTIVE').length
  };
}

/**
 * 获取工资录入进度
 */
function getSalaryEntryProgress(currentUser, month) {
  if (!currentUser) return { success: false };
  
  const cleanMonth = sanitizeInput_(month, 'string') || getCurrentMonth_();
  
  const ss = SpreadsheetApp.getActive();
  const salarySheet = ss.getSheetByName(SHEET_SALARY);
  
  let totalPeople = 0;
  let enteredNames = new Set();
  
  if (currentUser.role === 'ADMIN') {
    const mgrSheet = ss.getSheetByName(SHEET_MANAGERS);
    if (mgrSheet && mgrSheet.getLastRow() > 1) {
      const mgrData = mgrSheet.getDataRange().getValues();
      const mgrHeader = mgrData[0];
      const mgrCols = {};
      mgrHeader.forEach((h, i) => mgrCols[String(h).trim()] = i);
      
      for (let i = 1; i < mgrData.length; i++) {
        if (mgrData[i][mgrCols['StaffName']] && mgrData[i][mgrCols['Status']] !== 'LEFT') {
          totalPeople++;
        }
      }
    }
  } else if (currentUser.role === 'SECRETARY') {
    const staffSheet = ss.getSheetByName(SHEET_STAFF);
    if (staffSheet && staffSheet.getLastRow() > 1) {
      const staffData = staffSheet.getDataRange().getValues();
      const staffHeader = staffData[0];
      const staffCols = {};
      staffHeader.forEach((h, i) => staffCols[String(h).trim()] = i);
      
      for (let i = 1; i < staffData.length; i++) {
        const createdBy = String(staffData[i][staffCols['CreatedBy']] || '').toLowerCase();
        const status = staffData[i][staffCols['Status']];
        if (createdBy === currentUser.username.toLowerCase() && status !== 'LEFT') {
          totalPeople++;
        }
      }
    }
  }
  
  if (salarySheet && salarySheet.getLastRow() > 1) {
    const salaryData = salarySheet.getDataRange().getValues();
    const salaryHeader = salaryData[0];
    const cols = {};
    salaryHeader.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < salaryData.length; i++) {
      const rowMonth = formatMonthValue_(salaryData[i][cols['Month']]);
      if (rowMonth !== cleanMonth) continue;
      
      const createdBy = String(salaryData[i][cols['CreatedBy']] || '').toLowerCase();
      const isManager = String(salaryData[i][cols['IsManagerRecord']] || '').toUpperCase() === 'YES';
      
      if (currentUser.role === 'ADMIN' && isManager) {
        enteredNames.add(salaryData[i][cols['StaffName']]);
      } else if (currentUser.role === 'SECRETARY' && !isManager && createdBy === currentUser.username.toLowerCase()) {
        enteredNames.add(salaryData[i][cols['StaffName']]);
      }
    }
  }
  
  const enteredCount = enteredNames.size;
  const notEnteredCount = Math.max(0, totalPeople - enteredCount);
  const progressPercent = totalPeople > 0 ? Math.round((enteredCount / totalPeople) * 100) : 0;
  
  return {
    success: true,
    month: cleanMonth,
    totalPeople: totalPeople,
    enteredCount: enteredCount,
    notEnteredCount: notEnteredCount,
    progressPercent: progressPercent
  };
}

/**
 * 获取已发放记录（历史）
 */
function getPaidHistory(currentUser, month) {
  if (!currentUser) return { success: false };
  if (currentUser.role === 'SECRETARY') return { success: false, message: '无权限' };
  
  const cleanMonth = sanitizeInput_(month, 'string');
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_SALARY);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, records: [], stats: {} };
  }
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const cols = {};
  header.forEach((h, i) => cols[String(h).trim()] = i);
  
  const records = [];
  let bankTotal = 0, cashTotal = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const payStatus = String(row[cols['PaymentStatus']] || '').trim();
    if (payStatus !== 'PAID') continue;
    
    const rowMonth = formatMonthValue_(row[cols['Month']]);
    if (cleanMonth && rowMonth !== cleanMonth) continue;
    
    const net = Number(row[cols['NetSalary']]) || 0;
    const method = row[cols['PaymentMethod']] || 'BANK';
    const isManager = String(row[cols['IsManagerRecord']] || '').toUpperCase() === 'YES';
    
    if (method === 'BANK') bankTotal += net;
    else cashTotal += net;
    
    records.push({
      month: rowMonth,
      staffName: row[cols['StaffName']],
      isManager: isManager,
      basicSalary: Number(row[cols['BasicSalary']]) || 0,
      deduction: Number(row[cols['Deduction']]) || 0,
      netSalary: net,
      paymentMethod: method,
      paidAt: row[cols['PaidAt']] || '',
      paidBy: row[cols['PaidBy']] || ''
    });
  }
  
  return {
    success: true,
    records: records,
    stats: {
      totalCount: records.length,
      bankTotal: bankTotal,
      cashTotal: cashTotal,
      grandTotal: bankTotal + cashTotal
    }
  };
}

/**
 * 更新员工/主管银行信息
 */
function updateStaffBankInfo(user, staffName, type, bankInfo) {
  try {
    if (type === 'STAFF' && user.role !== 'SECRETARY' && user.role !== 'ADMIN') {
      return { success: false, message: '无权限修改员工银行信息' };
    }
    if (type === 'MANAGER' && user.role !== 'ADMIN') {
      return { success: false, message: '无权限修改主管银行信息' };
    }
    
    const ss = SpreadsheetApp.getActive();
    const sheetName = type === 'MANAGER' ? 'Managers' : 'Staff';
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, message: '找不到' + sheetName + '表' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const colIndex = {};
    headers.forEach(function(header, index) {
      colIndex[header] = index;
    });
    
    const colStaffName = colIndex['StaffName'];
    const colBankHolder = colIndex['BankHolder'];
    const colBankType = colIndex['BankType'];
    const colBankAccount = colIndex['BankAccount'];
    const colLastBankUpdate = colIndex['LastBankUpdate'] !== undefined ? colIndex['LastBankUpdate'] : colIndex['BankChangedAt'];
    const colBankChangeNote = colIndex['BankChangeNote'];
    const colOldBankHolder = colIndex['OldBankHolder'];
    const colOldBankType = colIndex['OldBankType'];
    const colOldBankAccount = colIndex['OldBankAccount'];
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][colStaffName] === staffName) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: '找不到' + staffName };
    }
    
    const currentRow = data[rowIndex - 1];
    const oldBankHolder = currentRow[colBankHolder] || '';
    const oldBankType = currentRow[colBankType] || '';
    const oldBankAccount = currentRow[colBankAccount] || '';
    
    if (colOldBankHolder !== undefined) {
      sheet.getRange(rowIndex, colOldBankHolder + 1).setValue(oldBankHolder);
    }
    if (colOldBankType !== undefined) {
      sheet.getRange(rowIndex, colOldBankType + 1).setValue(oldBankType);
    }
    if (colOldBankAccount !== undefined) {
      sheet.getRange(rowIndex, colOldBankAccount + 1).setValue(oldBankAccount);
    }
    
    if (colBankHolder !== undefined) {
      sheet.getRange(rowIndex, colBankHolder + 1).setValue(bankInfo.bankHolder || '');
    }
    if (colBankType !== undefined) {
      sheet.getRange(rowIndex, colBankType + 1).setValue(bankInfo.bankType || '');
    }
    if (colBankAccount !== undefined) {
      sheet.getRange(rowIndex, colBankAccount + 1).setValue(bankInfo.bankAccount || '');
    }
    
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    if (colLastBankUpdate !== undefined) {
      sheet.getRange(rowIndex, colLastBankUpdate + 1).setValue(timestamp);
    }
    
    const changeNote = '银行: ' + (oldBankType || '无') + ' → ' + (bankInfo.bankType || '无');
    if (colBankChangeNote !== undefined) {
      sheet.getRange(rowIndex, colBankChangeNote + 1).setValue(changeNote + (bankInfo.changeNote ? ' (' + bankInfo.changeNote + ')' : ''));
    }
    
    logOperation_('UPDATE_BANK_INFO', {
      staffName: staffName,
      type: type,
      changeNote: changeNote
    }, user.username);
    
    return { success: true, message: '银行信息已更新' };
    
  } catch (e) {
    Logger.log('updateStaffBankInfo error: ' + e.stack);
    return { success: false, message: e.message };
  }
}

// ========== 调试函数 ==========

function testLogin() {
  const result = login('admin', '123456');
  Logger.log('登录测试: ' + JSON.stringify(result));
  return result;
}

function testGetStaffList() {
  const user = { username: 'secretary1', role: 'SECRETARY' };
  const result = getStaffList(user);
  Logger.log('员工列表测试: ' + JSON.stringify(result));
  return result;
}

function testGetManagerList() {
  const user = { username: 'admin', role: 'ADMIN' };
  const result = getManagerList(user);
  Logger.log('主管列表测试: ' + JSON.stringify(result));
  return result;
}

/**
 * 获取指定书记的详细统计数据
 */
function getSecretaryStats(currentUser, secretaryName) {
  if (!currentUser || (currentUser.role !== 'ADMIN' && currentUser.role !== 'ACCOUNTANT')) {
    return { success: false, message: '无权限' };
  }
  
  const cleanName = sanitizeInput_(secretaryName, 'name');
  if (!cleanName) {
    return { success: false, message: '书记名称无效' };
  }
  
  const ss = SpreadsheetApp.getActive();
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  const salarySheet = ss.getSheetByName(SHEET_SALARY);
  
  const staffList = [];
  const currentMonth = getCurrentMonth_();
  
  // 获取该书记的所有员工
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const data = staffSheet.getDataRange().getValues();
    const headers = data[0];
    const cols = {};
    headers.forEach((h, i) => cols[String(h).trim()] = i);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const createdBy = String(row[cols['CreatedBy']] || '').trim();
      
      if (createdBy.toLowerCase() === cleanName.toLowerCase()) {
        const totalDebt = Number(row[cols['TotalDebt']]) || 0;
        const debtPaid = Number(row[cols['DebtPaid']]) || 0;
        
        staffList.push({
          staffName: String(row[cols['StaffName']] || ''),
          companyName: String(row[cols['CompanyName']] || ''),
          salary: Number(row[cols['Salary']]) || 0,
          paymentMethod: String(row[cols['PaymentMethod']] || 'BANK'),
          bankType: String(row[cols['BankType']] || ''),
          status: String(row[cols['Status']] || 'ACTIVE'),
          totalDebt: totalDebt,
          debtPaid: debtPaid,
          debtRemaining: totalDebt - debtPaid,
          monthlyDeduction: Number(row[cols['MonthlyDeduction']]) || 0
        });
      }
    }
  }
  
  // 统计本月工资录入情况
  let enteredCount = 0;
  const enteredNames = new Set();
  
  if (salarySheet && salarySheet.getLastRow() > 1) {
    const salaryData = salarySheet.getDataRange().getValues();
    const salaryHeaders = salaryData[0];
    const sCols = {};
    salaryHeaders.forEach((h, i) => sCols[String(h).trim()] = i);
    
    for (let i = 1; i < salaryData.length; i++) {
      const row = salaryData[i];
      const rowMonth = formatMonthValue_(row[sCols['Month']]);
      const createdBy = String(row[sCols['CreatedBy']] || '').toLowerCase();
      const isManager = String(row[sCols['IsManagerRecord']] || '').toUpperCase() === 'YES';
      const staffName = row[sCols['StaffName']];
      
      if (rowMonth === currentMonth && !isManager && createdBy === cleanName.toLowerCase()) {
        enteredNames.add(staffName);
      }
    }
    enteredCount = enteredNames.size;
  }
  
  // 按状态排序：在职优先
  staffList.sort((a, b) => {
    if (a.status === 'LEFT' && b.status !== 'LEFT') return 1;
    if (a.status !== 'LEFT' && b.status === 'LEFT') return -1;
    return a.staffName.localeCompare(b.staffName);
  });
  
  return {
    success: true,
    secretaryName: cleanName,
    staffList: staffList,
    currentMonth: currentMonth,
    stats: {
      totalCount: staffList.length,
      activeCount: staffList.filter(s => s.status !== 'LEFT').length,
      leftCount: staffList.filter(s => s.status === 'LEFT').length,
      enteredCount: enteredCount,
      totalSalary: staffList.filter(s => s.status !== 'LEFT').reduce((sum, s) => sum + s.salary, 0),
      withDebtCount: staffList.filter(s => s.status !== 'LEFT' && s.debtRemaining > 0).length
    }
  };
}
