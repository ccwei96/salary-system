/**
 * ============================================
 * 📁 文件名: setup.gs
 * 📝 描述: 薪资管理系统 - 高级初始化与管理脚本
 * 🔖 版本: 2.0
 * ============================================
 * 
 * 功能特性：
 * ✅ 一键初始化系统
 * ✅ 数据验证与条件格式
 * ✅ 系统健康检查与自动修复
 * ✅ 数据备份与恢复
 * ✅ 批量用户管理
 * ✅ 密码安全升级
 * ✅ 数据统计报告
 * ✅ 表结构自动迁移
 * 
 * 使用方法：
 * 1. 在 Apps Script 编辑器中点击 "+" 新建文件
 * 2. 命名为 "setup"
 * 3. 粘贴此代码
 * 4. 保存后刷新 Google Sheets
 * 5. 点击菜单 "🛠️ 薪资系统" > "🚀 一键初始化"
 */

// ==================== 系统配置 ====================
const SETUP_CONFIG = {
  // 系统版本
  version: '2.0',
  
  // 工作表配置
  sheets: {
    Users: {
      headers: ['UserID', 'Username', 'Password', 'Role', 'Status', 'DisplayName', 'MustChangePassword', 'CreatedAt', 'LastLoginAt'],
      columnWidths: [60, 120, 200, 100, 80, 150, 120, 150, 150],
      color: '#4285F4',
      description: '用户账号表',
      validation: {
        Role: ['ADMIN', 'SECRETARY', 'ACCOUNTANT'],
        Status: ['ACTIVE', 'INACTIVE'],
        MustChangePassword: ['YES', '']
      }
    },
    Staff: {
      headers: [
        'StaffName', 'Salary', 'CompanyName', 'BankHolder', 'BankType', 'BankAccount', 
        'PaymentMethod', 'IsManager', 'JoinDate', 'LeaveDate', 'TotalDebt', 'MonthlyDeduction', 
        'DebtPaid', 'DebtReason', 'LastBankUpdate', 'BankChangeNote', 
        'OldBankHolder', 'OldBankType', 'OldBankAccount',
        'OldSalary', 'SalaryChangeDate', 'SalaryChangeNote',  // ← 新增3列
        'Status', 'CreatedBy', 'CreatedAt'
      ],
      columnWidths: [120, 100, 150, 150, 100, 150, 100, 80, 100, 100, 100, 100, 100, 150, 150, 200, 150, 100, 150, 100, 150, 200, 80, 100, 150],
      color: '#34A853',
      description: '员工信息表',
      validation: {
        PaymentMethod: ['BANK', 'CASH'],
        IsManager: ['YES', ''],
        Status: ['ACTIVE', 'LEFT']
      }
    },
    Managers: {
      headers: [
        'StaffName', 'Salary', 'CompanyName', 'BankHolder', 'BankType', 'BankAccount', 
        'PaymentMethod', 'JoinDate', 'LeaveDate', 'TotalDebt', 'MonthlyDeduction', 
        'DebtPaid', 'DebtReason', 'LastBankUpdate', 'BankChangeNote', 
        'OldBankHolder', 'OldBankType', 'OldBankAccount',
        'OldSalary', 'SalaryChangeDate', 'SalaryChangeNote',  // ← 新增3列
        'Status', 'CreatedBy', 'CreatedAt'
      ],
      columnWidths: [120, 100, 150, 150, 100, 150, 100, 100, 100, 100, 100, 100, 150, 150, 200, 150, 100, 150, 100, 150, 200, 80, 100, 150],
      color: '#FBBC04',
      description: '主管信息表',
      validation: {
        PaymentMethod: ['BANK', 'CASH'],
        Status: ['ACTIVE', 'LEFT']
      }
    },
    SalaryRecords: {
      headers: [
        'Month', 'Date', 'StaffName', 'IsManagerRecord', 'BasicSalary', 'ManualDeduction', 
        'AutoDeduction', 'BankFee', 'Deduction', 'NetSalary', 'Remark', 'CreatedBy', 'CreatedAt', 
        'SubmitStatus', 'SubmittedAt', 'PaymentStatus', 'PaymentMethod', 'PaidAt', 'PaidBy'
      ],
      columnWidths: [80, 100, 120, 100, 100, 100, 100, 80, 100, 100, 200, 100, 150, 100, 150, 100, 100, 150, 100],
      color: '#EA4335',
      description: '工资记录表',
      validation: {
        SubmitStatus: ['DRAFT', 'SUBMITTED'],
        PaymentStatus: ['PENDING', 'PAID'],
        IsManagerRecord: ['YES', ''],
        PaymentMethod: ['BANK', 'CASH']
      }
    },
    PaymentLog: {
      headers: ['LogID', 'Timestamp', 'Action', 'ActionName', 'Operator', 'TargetType', 'TargetName', 'Details', 'IPInfo'],
      columnWidths: [150, 150, 150, 150, 100, 100, 120, 400, 100],
      color: '#9C27B0',
      description: '操作日志表',
      validation: {}
    }
  },
  
  // 默认管理员
  defaultAdmin: {
    username: 'admin',
    password: 'admin123',
    role: 'ADMIN',
    displayName: '系统管理员'
  },
  
  // 样式配置
  styles: {
    headerBgColor: '#2D3748',
    headerFontColor: '#FFFFFF',
    headerFontSize: 11,
    dataFontSize: 10,
    alternateRowColor: '#F7FAFC'
  },
  
  // 备份配置
  backup: {
    prefix: 'BACKUP_',
    maxBackups: 5
  }
};

// ==================== 菜单系统 ====================

/**
 * 添加自定义菜单
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🛠️ 薪资系统')
    // 初始化
    .addItem('🚀 一键初始化', 'initializeSystem')
    .addItem('🔄 升级表结构', 'upgradeTableStructure')
    .addSeparator()
    
    // 系统管理
    .addSubMenu(ui.createMenu('📊 系统管理')
      .addItem('📋 检查系统状态', 'checkSystemStatus')
      .addItem('🔧 自动修复问题', 'autoFixProblems')
      .addItem('📈 生成统计报告', 'generateStatisticsReport')
      .addItem('🔄 重置表格格式', 'resetAllFormatting'))
    
    // 用户管理
    .addSubMenu(ui.createMenu('👥 用户管理')
      .addItem('👤 添加新用户', 'showAddUserDialog')
      .addItem('🔑 重置用户密码', 'showResetPasswordDialog')
      .addItem('🔐 升级密码格式', 'upgradeAllPasswords')
      .addItem('📋 查看用户列表', 'showUserList')
      .addItem('🚫 停用用户账号', 'disableUserAccount'))
    
    // 数据管理
    .addSubMenu(ui.createMenu('💾 数据管理')
      .addItem('📦 备份所有数据', 'backupAllData')
      .addItem('📥 恢复数据备份', 'restoreFromBackup')
      .addItem('📤 导出为CSV', 'exportToCSV')
      .addItem('🗑️ 清理旧日志', 'cleanupOldLogs'))
    
    // 危险操作
    .addSeparator()
    .addSubMenu(ui.createMenu('⚠️ 危险操作')
      .addItem('🧹 清空所有数据', 'clearAllData')
      .addItem('💥 重置整个系统', 'resetEntireSystem'))
    
    .addToUi();
}

/**
 * 安装触发器
 */
function installTriggers() {
  // 删除旧触发器
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // 安装 onOpen 触发器
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}

// ==================== 核心初始化函数 ====================

/**
 * 🚀 一键初始化系统
 */
function initializeSystem() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🚀 薪资管理系统初始化',
    '此操作将创建以下工作表：\n\n' +
    '• Users - 用户账号表\n' +
    '• Staff - 员工信息表\n' +
    '• Managers - 主管信息表\n' +
    '• SalaryRecords - 工资记录表\n' +
    '• PaymentLog - 操作日志表\n\n' +
    '同时会创建默认管理员账户。\n' +
    '已存在的工作表将保留数据，仅更新格式。\n\n' +
    '是否继续？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('❌ 操作已取消');
    return;
  }
  
  const startTime = new Date();
  
  try {
    const ss = SpreadsheetApp.getActive();
    let createdCount = 0;
    let updatedCount = 0;
    const results = [];
    
    // 显示进度
    ss.toast('正在初始化...', '🚀 系统初始化', 30);
    
    // 创建所有工作表
    for (const [sheetName, config] of Object.entries(SETUP_CONFIG.sheets)) {
      const result = createOrUpdateSheet_(ss, sheetName, config);
      if (result.created) {
        createdCount++;
        results.push(`✅ ${sheetName} - 新建成功`);
      } else {
        updatedCount++;
        results.push(`🔄 ${sheetName} - 已更新格式`);
      }
    }
    
    // 创建默认管理员
    const adminResult = createDefaultAdmin_(ss);
    results.push(adminResult);
    
    // 清理默认工作表
    cleanupDefaultSheet_(ss);
    
    // 创建说明表
    createInstructionSheet_(ss);
    results.push('📖 使用说明 - 已创建');
    
    // 安装触发器
    try {
      installTriggers();
      results.push('⚙️ 触发器 - 已安装');
    } catch (e) {
      results.push('⚠️ 触发器 - 安装失败（需手动授权）');
    }
    
    const elapsed = ((new Date() - startTime) / 1000).toFixed(1);
    
    // 显示结果
    ui.alert(
      '✨ 初始化完成！',
      `⏱️ 耗时: ${elapsed} 秒\n` +
      `📄 新建: ${createdCount} 个工作表\n` +
      `🔄 更新: ${updatedCount} 个工作表\n\n` +
      '═══════════════════════════════\n' +
      results.join('\n') + '\n\n' +
      '═══════════════════════════════\n' +
      '🔐 默认管理员账户：\n' +
      `   用户名: ${SETUP_CONFIG.defaultAdmin.username}\n` +
      `   密码: ${SETUP_CONFIG.defaultAdmin.password}\n\n` +
      '⚠️ 请首次登录后立即修改密码！',
      ui.ButtonSet.OK
    );
    
    // 记录日志
    logSetupOperation_('SYSTEM_INIT', '系统初始化完成', {
      created: createdCount,
      updated: updatedCount,
      elapsed: elapsed
    });
    
  } catch (error) {
    ui.alert('❌ 初始化失败', '错误信息：' + error.message + '\n\n请查看执行日志获取详细信息。', ui.ButtonSet.OK);
    Logger.log('初始化错误: ' + error.stack);
  }
}

/**
 * 升级表结构（添加新列）
 */
function upgradeTableStructure() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🔄 升级表结构',
    '此操作将检查并添加缺失的列，不会删除现有数据。\n\n是否继续？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActive();
  const results = [];
  
  try {
    for (const [sheetName, config] of Object.entries(SETUP_CONFIG.sheets)) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        results.push(`⏭️ ${sheetName} - 表不存在，跳过`);
        continue;
      }
      
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const targetHeaders = config.headers;
      let addedCols = [];
      
      // 检查缺失的列
      for (const header of targetHeaders) {
        if (!currentHeaders.includes(header)) {
          // 在末尾添加新列
          const newColIndex = sheet.getLastColumn() + 1;
          sheet.getRange(1, newColIndex).setValue(header);
          addedCols.push(header);
        }
      }
      
      if (addedCols.length > 0) {
        results.push(`✅ ${sheetName} - 添加了 ${addedCols.length} 列: ${addedCols.join(', ')}`);
        
        // 重新应用格式
        applySheetFormatting_(sheet, config);
      } else {
        results.push(`✓ ${sheetName} - 结构完整`);
      }
    }
    
    ui.alert('🔄 升级完成', results.join('\n'), ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ 升级失败', error.message, ui.ButtonSet.OK);
    Logger.log('升级错误: ' + error.stack);
  }
}

// ==================== 工作表创建与格式化 ====================

/**
 * 创建或更新工作表
 */
function createOrUpdateSheet_(ss, sheetName, config) {
  let sheet = ss.getSheetByName(sheetName);
  let created = false;
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    created = true;
  }
  
  const headers = config.headers;
  const numCols = headers.length;
  
  // 写入表头
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setValues([headers]);
  
  // 应用格式
  applySheetFormatting_(sheet, config);
  
  // 添加数据验证
  addDataValidation_(sheet, config);
  
  // 添加条件格式
  addConditionalFormatting_(sheet, sheetName, headers);
  
  return { created, sheet };
}

/**
 * 应用工作表格式
 */
function applySheetFormatting_(sheet, config) {
  const headers = config.headers;
  const numCols = headers.length;
  
  // 格式化表头
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground(SETUP_CONFIG.styles.headerBgColor)
    .setFontColor(SETUP_CONFIG.styles.headerFontColor)
    .setFontWeight('bold')
    .setFontSize(SETUP_CONFIG.styles.headerFontSize)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#1A202C', SpreadsheetApp.BorderStyle.SOLID);
  
  sheet.setRowHeight(1, 40);
  
  // 设置列宽
  if (config.columnWidths) {
    for (let i = 0; i < Math.min(config.columnWidths.length, numCols); i++) {
      sheet.setColumnWidth(i + 1, config.columnWidths[i]);
    }
  }
  
  // 冻结表头
  sheet.setFrozenRows(1);
  
  // 数据区域格式
  if (sheet.getLastRow() > 1 || sheet.getMaxRows() > 1) {
    const maxRows = Math.max(sheet.getLastRow(), 100);
    const dataRange = sheet.getRange(2, 1, maxRows - 1, numCols);
    dataRange
      .setFontSize(SETUP_CONFIG.styles.dataFontSize)
      .setVerticalAlignment('middle');
  }
  
  // 标签颜色
  sheet.setTabColor(config.color);
}

/**
 * 添加数据验证
 */
function addDataValidation_(sheet, config) {
  if (!config.validation) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  for (const [colName, options] of Object.entries(config.validation)) {
    const colIndex = headers.indexOf(colName) + 1;
    if (colIndex > 0) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options, true)
        .setAllowInvalid(false)
        .build();
      
      // 应用到数据区域（第2行到第500行）
      sheet.getRange(2, colIndex, 499, 1).setDataValidation(rule);
    }
  }
}

/**
 * 添加条件格式
 */
function addConditionalFormatting_(sheet, sheetName, headers) {
  // 清除现有条件格式
  sheet.clearConditionalFormatRules();
  
  const rules = [];
  const dataRange = sheet.getRange(2, 1, 500, headers.length);
  
  // 通用状态格式
  const statusCol = headers.indexOf('Status') + 1;
  if (statusCol > 0) {
    // INACTIVE / LEFT - 灰色背景
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('INACTIVE')
      .setBackground('#E2E8F0')
      .setFontColor('#718096')
      .setStrikethrough(true)
      .setRanges([dataRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('LEFT')
      .setBackground('#E2E8F0')
      .setFontColor('#718096')
      .setStrikethrough(true)
      .setRanges([dataRange])
      .build());
  }
  
  // 工资记录特定格式
  if (sheetName === 'SalaryRecords') {
    const submitCol = headers.indexOf('SubmitStatus') + 1;
    const payCol = headers.indexOf('PaymentStatus') + 1;
    
    if (submitCol > 0) {
      const submitRange = sheet.getRange(2, submitCol, 500, 1);
      
      // DRAFT - 黄色
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('DRAFT')
        .setBackground('#FEFCBF')
        .setFontColor('#975A16')
        .setRanges([submitRange])
        .build());
      
      // SUBMITTED - 绿色
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('SUBMITTED')
        .setBackground('#C6F6D5')
        .setFontColor('#276749')
        .setRanges([submitRange])
        .build());
    }
    
    if (payCol > 0) {
      const payRange = sheet.getRange(2, payCol, 500, 1);
      
      // PAID - 绿色
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('PAID')
        .setBackground('#C6F6D5')
        .setFontColor('#276749')
        .setRanges([payRange])
        .build());
      
      // PENDING - 红色
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('PENDING')
        .setBackground('#FED7D7')
        .setFontColor('#C53030')
        .setRanges([payRange])
        .build());
    }
  }
  
  // 日志特定格式
  if (sheetName === 'PaymentLog') {
    const actionCol = headers.indexOf('Action') + 1;
    if (actionCol > 0) {
      // 失败操作 - 红色
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('FAILED')
        .setBackground('#FED7D7')
        .setFontColor('#C53030')
        .setRanges([dataRange])
        .build());
      
      // 拒绝操作 - 橙色
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('DENIED')
        .setBackground('#FEEBC8')
        .setFontColor('#C05621')
        .setRanges([dataRange])
        .build());
      
      // 登录操作 - 蓝色
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('LOGIN')
        .setBackground('#BEE3F8')
        .setFontColor('#2B6CB0')
        .setRanges([dataRange])
        .build());
    }
  }
  
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
  }
}

// ==================== 密码安全函数 ====================

/**
 * 密码哈希（带盐值）
 */
function hashPasswordLocal_(plain, salt) {
  if (!salt) {
    salt = Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  }
  
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + plain);
  const hash = bytes.map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join('');
  
  return salt + ':' + hash;
}

/**
 * 获取哈希函数（优先使用 Code.gs 中的）
 */
function getHashFunction_() {
  try {
    if (typeof hashPassword_ === 'function') {
      return hashPassword_;
    }
  } catch (e) {}
  return hashPasswordLocal_;
}

/**
 * 验证密码强度
 */
function validatePasswordStrength_(password) {
  const errors = [];
  
  if (!password || password.length < 6) {
    errors.push('密码至少6个字符');
  }
  if (password.length > 50) {
    errors.push('密码不能超过50个字符');
  }
  
  // 可选的强度检查
  // if (!/[A-Z]/.test(password)) errors.push('建议包含大写字母');
  // if (!/[0-9]/.test(password)) errors.push('建议包含数字');
  
  return {
    valid: errors.length === 0,
    errors: errors,
    strength: password.length >= 12 ? 'strong' : password.length >= 8 ? 'medium' : 'weak'
  };
}

// ==================== 用户管理函数 ====================

/**
 * 创建默认管理员
 */
function createDefaultAdmin_(ss) {
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return '❌ Users 表不存在';
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('Username');
  
  // 检查是否已存在
  for (let i = 1; i < data.length; i++) {
    if (data[i][usernameCol] === SETUP_CONFIG.defaultAdmin.username) {
      return '⏭️ 管理员账户已存在';
    }
  }
  
  const hashFunc = getHashFunction_();
  const hashedPwd = hashFunc(SETUP_CONFIG.defaultAdmin.password);
  const now = new Date().toISOString();
  
  // 构建行数据
  const rowData = [];
  for (const header of headers) {
    switch (header) {
      case 'UserID': rowData.push(data.length); break;
      case 'Username': rowData.push(SETUP_CONFIG.defaultAdmin.username); break;
      case 'Password': rowData.push(hashedPwd); break;
      case 'Role': rowData.push(SETUP_CONFIG.defaultAdmin.role); break;
      case 'Status': rowData.push('ACTIVE'); break;
      case 'DisplayName': rowData.push(SETUP_CONFIG.defaultAdmin.displayName); break;
      case 'MustChangePassword': rowData.push('YES'); break;
      case 'CreatedAt': rowData.push(now); break;
      default: rowData.push('');
    }
  }
  
  sheet.appendRow(rowData);
  return '✅ 管理员账户已创建';
}

/**
 * 显示添加用户对话框
 */
function showAddUserDialog() {
  const ui = SpreadsheetApp.getUi();
  
  // 用户名
  const usernameResponse = ui.prompt('👤 添加新用户', '用户名（3-20位字母数字下划线）：', ui.ButtonSet.OK_CANCEL);
  if (usernameResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const username = usernameResponse.getResponseText().trim();
  if (!username) return;
  
  if (username.length < 3 || username.length > 20 || !/^[a-zA-Z0-9_]+$/.test(username)) {
    ui.alert('❌ 错误', '用户名格式不正确（3-20位字母数字下划线）', ui.ButtonSet.OK);
    return;
  }
  
  // 密码
  const passwordResponse = ui.prompt('🔑 设置密码', '密码（至少6位）：', ui.ButtonSet.OK_CANCEL);
  if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const password = passwordResponse.getResponseText().trim();
  if (!password) return;
  
  const pwdValidation = validatePasswordStrength_(password);
  if (!pwdValidation.valid) {
    ui.alert('❌ 密码不符合要求', pwdValidation.errors.join('\n'), ui.ButtonSet.OK);
    return;
  }
  
  // 角色
  const roleResponse = ui.prompt('👔 选择角色', '角色 (ADMIN / SECRETARY / ACCOUNTANT)：', ui.ButtonSet.OK_CANCEL);
  if (roleResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const role = roleResponse.getResponseText().trim().toUpperCase();
  if (!['ADMIN', 'SECRETARY', 'ACCOUNTANT'].includes(role)) {
    ui.alert('❌ 无效角色', '请输入 ADMIN、SECRETARY 或 ACCOUNTANT', ui.ButtonSet.OK);
    return;
  }
  
  // 显示名称
  const displayResponse = ui.prompt('📛 显示名称', '显示名称（可选，直接确定使用用户名）：', ui.ButtonSet.OK_CANCEL);
  if (displayResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const displayName = displayResponse.getResponseText().trim() || username;
  
  // 创建用户
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) {
    ui.alert('❌ 错误', 'Users 表不存在，请先初始化系统', ui.ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('Username');
  
  // 检查用户名是否已存在
  for (let i = 1; i < data.length; i++) {
    if (data[i][usernameCol] === username) {
      ui.alert('❌ 错误', '用户名已存在', ui.ButtonSet.OK);
      return;
    }
  }
  
  const hashFunc = getHashFunction_();
  const hashedPwd = hashFunc(password);
  const now = new Date().toISOString();
  
  // 构建行数据
  const rowData = [];
  for (const header of headers) {
    switch (header) {
      case 'UserID': rowData.push(data.length); break;
      case 'Username': rowData.push(username); break;
      case 'Password': rowData.push(hashedPwd); break;
      case 'Role': rowData.push(role); break;
      case 'Status': rowData.push('ACTIVE'); break;
      case 'DisplayName': rowData.push(displayName); break;
      case 'MustChangePassword': rowData.push('YES'); break;
      case 'CreatedAt': rowData.push(now); break;
      default: rowData.push('');
    }
  }
  
  sheet.appendRow(rowData);
  
  logSetupOperation_('ADD_USER', '添加用户', { username, role, displayName });
  
  ui.alert('✅ 成功', `用户 "${username}" 创建成功！\n\n角色: ${role}\n首次登录需修改密码`, ui.ButtonSet.OK);
}

/**
 * 显示重置密码对话框
 */
function showResetPasswordDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const usernameResponse = ui.prompt('🔑 重置密码', '要重置密码的用户名：', ui.ButtonSet.OK_CANCEL);
  if (usernameResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const username = usernameResponse.getResponseText().trim();
  if (!username) return;
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) {
    ui.alert('❌ 错误', 'Users 表不存在', ui.ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('Username');
  const passwordCol = headers.indexOf('Password');
  const mustChangeCol = headers.indexOf('MustChangePassword');
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][usernameCol] === username) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    ui.alert('❌ 错误', '找不到用户: ' + username, ui.ButtonSet.OK);
    return;
  }
  
  const passwordResponse = ui.prompt('🔑 新密码', '新密码（至少6位）：', ui.ButtonSet.OK_CANCEL);
  if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const newPassword = passwordResponse.getResponseText().trim();
  if (!newPassword) return;
  
  const pwdValidation = validatePasswordStrength_(newPassword);
  if (!pwdValidation.valid) {
    ui.alert('❌ 密码不符合要求', pwdValidation.errors.join('\n'), ui.ButtonSet.OK);
    return;
  }
  
  const hashFunc = getHashFunction_();
  const hashedPwd = hashFunc(newPassword);
  
  sheet.getRange(rowIndex, passwordCol + 1).setValue(hashedPwd);
  
  // 设置需要首次修改密码
  if (mustChangeCol >= 0) {
    sheet.getRange(rowIndex, mustChangeCol + 1).setValue('YES');
  }
  
  logSetupOperation_('RESET_PASSWORD', '重置密码', { username });
  
  ui.alert('✅ 成功', `用户 "${username}" 的密码已重置！\n\n用户下次登录需要修改密码。`, ui.ButtonSet.OK);
}

/**
 * 升级所有密码格式
 */
function upgradeAllPasswords() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🔐 升级密码格式',
    '此操作将检查所有旧格式密码并提示升级。\n\n' +
    '旧格式密码（无盐值）的用户将被重置为临时密码。\n\n' +
    '是否继续？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) {
    ui.alert('❌ 错误', 'Users 表不存在', ui.ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('Username');
  const passwordCol = headers.indexOf('Password');
  const mustChangeCol = headers.indexOf('MustChangePassword');
  
  let upgradedCount = 0;
  const upgradedUsers = [];
  const hashFunc = getHashFunction_();
  
  for (let i = 1; i < data.length; i++) {
    const pwd = String(data[i][passwordCol] || '');
    const username = data[i][usernameCol];
    
    // 检查是否是旧格式（不包含冒号且长度为64的纯哈希）
    if (pwd && !pwd.includes(':') && pwd.length === 64) {
      // 生成临时密码
      const tempPassword = 'Temp' + Math.random().toString(36).substring(2, 8).toUpperCase();
      const hashedPwd = hashFunc(tempPassword);
      
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, passwordCol + 1).setValue(hashedPwd);
      
      // 设置需要首次修改密码
      if (mustChangeCol >= 0) {
        sheet.getRange(rowIndex, mustChangeCol + 1).setValue('YES');
      }
      
      upgradedUsers.push(`${username}: ${tempPassword}`);
      upgradedCount++;
    }
  }
  
  if (upgradedCount > 0) {
    logSetupOperation_('UPGRADE_PASSWORDS', '升级密码格式', { count: upgradedCount });
    
    ui.alert(
      '✅ 升级完成',
      `已升级 ${upgradedCount} 个账户的密码格式。\n\n` +
      '临时密码列表（请妥善保管）：\n' +
      upgradedUsers.join('\n') + '\n\n' +
      '⚠️ 这些用户下次登录时需要修改密码。',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert('ℹ️ 提示', '没有需要升级的旧格式密码。', ui.ButtonSet.OK);
  }
}

/**
 * 显示用户列表
 */
function showUserList() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Users');
  
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('ℹ️ 提示', '暂无用户数据', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('Username');
  const roleCol = headers.indexOf('Role');
  const statusCol = headers.indexOf('Status');
  const displayNameCol = headers.indexOf('DisplayName');
  
  let userList = '👥 用户列表\n\n';
  userList += '序号 | 用户名 | 显示名 | 角色 | 状态\n';
  userList += '─'.repeat(50) + '\n';
  
  for (let i = 1; i < data.length; i++) {
    const username = data[i][usernameCol] || '';
    const role = data[i][roleCol] || '';
    const status = data[i][statusCol] || '';
    const displayName = data[i][displayNameCol] || username;
    
    const statusIcon = status === 'ACTIVE' ? '✅' : '❌';
    userList += `${i} | ${username} | ${displayName} | ${role} | ${statusIcon} ${status}\n`;
  }
  
  SpreadsheetApp.getUi().alert('用户列表', userList, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 停用用户账号
 */
function disableUserAccount() {
  const ui = SpreadsheetApp.getUi();
  
  const usernameResponse = ui.prompt('🚫 停用账号', '要停用的用户名：', ui.ButtonSet.OK_CANCEL);
  if (usernameResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const username = usernameResponse.getResponseText().trim();
  if (!username) return;
  
  if (username === 'admin') {
    ui.alert('❌ 错误', '不能停用默认管理员账号', ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) {
    ui.alert('❌ 错误', 'Users 表不存在', ui.ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('Username');
  const statusCol = headers.indexOf('Status');
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][usernameCol] === username) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    ui.alert('❌ 错误', '找不到用户: ' + username, ui.ButtonSet.OK);
    return;
  }
  
  const confirm = ui.alert('⚠️ 确认', `确定要停用用户 "${username}" 吗？`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  
  sheet.getRange(rowIndex, statusCol + 1).setValue('INACTIVE');
  
  logSetupOperation_('DISABLE_USER', '停用用户', { username });
  
  ui.alert('✅ 成功', `用户 "${username}" 已停用`, ui.ButtonSet.OK);
}

// ==================== 系统管理函数 ====================

/**
 * 检查系统状态
 */
function checkSystemStatus() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  let status = '╔══════════════════════════════════════╗\n';
  status += '║       📊 系统状态报告                ║\n';
  status += '╠══════════════════════════════════════╣\n\n';
  
  let hasProblems = false;
  const problems = [];
  
  // 检查工作表
  status += '【工作表状态】\n';
  for (const [sheetName, config] of Object.entries(SETUP_CONFIG.sheets)) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const rowCount = Math.max(0, sheet.getLastRow() - 1);
      const colCount = sheet.getLastColumn();
      const expectedCols = config.headers.length;
      
      if (colCount < expectedCols) {
        status += `⚠️ ${sheetName}: ${rowCount} 条 (缺少 ${expectedCols - colCount} 列)\n`;
        problems.push(`${sheetName} 缺少列`);
        hasProblems = true;
      } else {
        status += `✅ ${sheetName}: ${rowCount} 条记录\n`;
      }
    } else {
      status += `❌ ${sheetName}: 不存在\n`;
      problems.push(`${sheetName} 表不存在`);
      hasProblems = true;
    }
  }
  
  // 检查密码格式
  status += '\n【安全检查】\n';
  const usersSheet = ss.getSheetByName('Users');
  if (usersSheet && usersSheet.getLastRow() > 1) {
    const userData = usersSheet.getDataRange().getValues();
    const headers = userData[0];
    const passwordCol = headers.indexOf('Password');
    
    let oldFormatCount = 0;
    let emptyPasswordCount = 0;
    
    for (let i = 1; i < userData.length; i++) {
      const pwd = String(userData[i][passwordCol] || '');
      if (!pwd) {
        emptyPasswordCount++;
      } else if (!pwd.includes(':')) {
        oldFormatCount++;
      }
    }
    
    if (oldFormatCount > 0) {
      status += `⚠️ 发现 ${oldFormatCount} 个旧格式密码\n`;
      problems.push('存在旧格式密码');
      hasProblems = true;
    } else {
      status += `✅ 所有密码格式正确\n`;
    }
    
    if (emptyPasswordCount > 0) {
      status += `❌ 发现 ${emptyPasswordCount} 个空密码\n`;
      problems.push('存在空密码');
      hasProblems = true;
    }
  }
  
  // 检查数据完整性
  status += '\n【数据完整性】\n';
  const salarySheet = ss.getSheetByName('SalaryRecords');
  if (salarySheet && salarySheet.getLastRow() > 1) {
    const salaryData = salarySheet.getDataRange().getValues();
    const headers = salaryData[0];
    const staffNameCol = headers.indexOf('StaffName');
    const basicCol = headers.indexOf('BasicSalary');
    
    let emptyNameCount = 0;
    let invalidSalaryCount = 0;
    
    for (let i = 1; i < salaryData.length; i++) {
      if (!salaryData[i][staffNameCol]) emptyNameCount++;
      const basic = Number(salaryData[i][basicCol]);
      if (isNaN(basic) || basic < 0) invalidSalaryCount++;
    }
    
    if (emptyNameCount > 0 || invalidSalaryCount > 0) {
      status += `⚠️ 工资记录: ${emptyNameCount} 条无姓名, ${invalidSalaryCount} 条无效金额\n`;
      hasProblems = true;
    } else {
      status += `✅ 工资记录数据完整\n`;
    }
  }
  
  // 总结
  status += '\n╠══════════════════════════════════════╣\n';
  if (hasProblems) {
    status += '║  ⚠️ 发现 ' + problems.length + ' 个问题需要处理        ║\n';
    status += '║  建议执行「自动修复问题」          ║\n';
  } else {
    status += '║  ✅ 系统运行正常                    ║\n';
  }
  status += '╚══════════════════════════════════════╝';
  
  ui.alert('系统状态', status, ui.ButtonSet.OK);
}

/**
 * 自动修复问题
 */
function autoFixProblems() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🔧 自动修复',
    '此操作将尝试自动修复以下问题：\n\n' +
    '• 缺失的工作表\n' +
    '• 缺失的列\n' +
    '• 格式问题\n' +
    '• 数据验证规则\n\n' +
    '是否继续？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActive();
  const fixes = [];
  
  try {
    ss.toast('正在修复...', '🔧 自动修复', 30);
    
    // 修复缺失的工作表和列
    for (const [sheetName, config] of Object.entries(SETUP_CONFIG.sheets)) {
      let sheet = ss.getSheetByName(sheetName);
      
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
        fixes.push(`✅ 创建了 ${sheetName} 表`);
      } else {
        // 检查并添加缺失的列
        const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        let addedCols = [];
        
        for (const header of config.headers) {
          if (!currentHeaders.includes(header)) {
            const newColIndex = sheet.getLastColumn() + 1;
            sheet.getRange(1, newColIndex).setValue(header);
            addedCols.push(header);
          }
        }
        
        if (addedCols.length > 0) {
          fixes.push(`✅ ${sheetName}: 添加了列 ${addedCols.join(', ')}`);
        }
      }
      
      // 重新应用格式
      applySheetFormatting_(sheet, config);
      addDataValidation_(sheet, config);
      addConditionalFormatting_(sheet, sheetName, config.headers);
    }
    
    if (fixes.length === 0) {
      fixes.push('✓ 未发现需要修复的问题');
    }
    
    logSetupOperation_('AUTO_FIX', '自动修复', { fixes: fixes.length });
    
    ui.alert('🔧 修复完成', fixes.join('\n'), ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ 修复失败', error.message, ui.ButtonSet.OK);
    Logger.log('修复错误: ' + error.stack);
  }
}

/**
 * 生成统计报告
 */
function generateStatisticsReport() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  let report = '📈 薪资系统统计报告\n';
  report += '生成时间: ' + new Date().toLocaleString('zh-CN') + '\n\n';
  report += '═'.repeat(40) + '\n\n';
  
  // 用户统计
  const usersSheet = ss.getSheetByName('Users');
  if (usersSheet && usersSheet.getLastRow() > 1) {
    const userData = usersSheet.getDataRange().getValues();
    const headers = userData[0];
    const roleCol = headers.indexOf('Role');
    const statusCol = headers.indexOf('Status');
    
    const roleCount = { ADMIN: 0, SECRETARY: 0, ACCOUNTANT: 0 };
    let activeCount = 0;
    
    for (let i = 1; i < userData.length; i++) {
      const role = userData[i][roleCol];
      const status = userData[i][statusCol];
      if (roleCount[role] !== undefined) roleCount[role]++;
      if (status === 'ACTIVE') activeCount++;
    }
    
    report += '【用户统计】\n';
    report += `总用户数: ${userData.length - 1}\n`;
    report += `活跃用户: ${activeCount}\n`;
    report += `管理员: ${roleCount.ADMIN}, 主管: ${roleCount.SECRETARY}, 会计: ${roleCount.ACCOUNTANT}\n\n`;
  }
  
  // 员工统计
  const staffSheet = ss.getSheetByName('Staff');
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const staffData = staffSheet.getDataRange().getValues();
    const headers = staffData[0];
    const statusCol = headers.indexOf('Status');
    const salaryCol = headers.indexOf('Salary');
    const debtCol = headers.indexOf('TotalDebt');
    const debtPaidCol = headers.indexOf('DebtPaid');
    
    let activeStaff = 0, leftStaff = 0;
    let totalSalary = 0, totalDebt = 0, totalDebtPaid = 0;
    
    for (let i = 1; i < staffData.length; i++) {
      const status = staffData[i][statusCol];
      if (status === 'ACTIVE') {
        activeStaff++;
        totalSalary += Number(staffData[i][salaryCol]) || 0;
      } else if (status === 'LEFT') {
        leftStaff++;
      }
      totalDebt += Number(staffData[i][debtCol]) || 0;
      totalDebtPaid += Number(staffData[i][debtPaidCol]) || 0;
    }
    
    report += '【员工统计】\n';
    report += `在职员工: ${activeStaff}\n`;
    report += `已离职: ${leftStaff}\n`;
    report += `月薪总额: RM ${totalSalary.toLocaleString()}\n`;
    report += `欠款总额: RM ${totalDebt.toLocaleString()}\n`;
    report += `已还欠款: RM ${totalDebtPaid.toLocaleString()}\n`;
    report += `未还欠款: RM ${(totalDebt - totalDebtPaid).toLocaleString()}\n\n`;
  }
  
  // 主管统计
  const mgrSheet = ss.getSheetByName('Managers');
  if (mgrSheet && mgrSheet.getLastRow() > 1) {
    const mgrData = mgrSheet.getDataRange().getValues();
    const headers = mgrData[0];
    const statusCol = headers.indexOf('Status');
    const salaryCol = headers.indexOf('Salary');
    
    let activeMgr = 0;
    let totalMgrSalary = 0;
    
    for (let i = 1; i < mgrData.length; i++) {
      const status = mgrData[i][statusCol];
      if (status === 'ACTIVE' || !status) {
        activeMgr++;
        totalMgrSalary += Number(mgrData[i][salaryCol]) || 0;
      }
    }
    
    report += '【主管统计】\n';
    report += `在职主管: ${activeMgr}\n`;
    report += `月薪总额: RM ${totalMgrSalary.toLocaleString()}\n\n`;
  }
  
  // 工资记录统计
  const salarySheet = ss.getSheetByName('SalaryRecords');
  if (salarySheet && salarySheet.getLastRow() > 1) {
    const salaryData = salarySheet.getDataRange().getValues();
    const headers = salaryData[0];
    const submitCol = headers.indexOf('SubmitStatus');
    const payCol = headers.indexOf('PaymentStatus');
    const netCol = headers.indexOf('NetSalary');
    
    let draftCount = 0, submittedCount = 0, paidCount = 0;
    let pendingAmount = 0, paidAmount = 0;
    
    for (let i = 1; i < salaryData.length; i++) {
      const submit = salaryData[i][submitCol];
      const pay = salaryData[i][payCol];
      const net = Number(salaryData[i][netCol]) || 0;
      
      if (submit === 'DRAFT') {
        draftCount++;
      } else if (submit === 'SUBMITTED') {
        if (pay === 'PAID') {
          paidCount++;
          paidAmount += net;
        } else {
          submittedCount++;
          pendingAmount += net;
        }
      }
    }
    
    report += '【工资记录统计】\n';
    report += `草稿: ${draftCount} 条\n`;
    report += `待发放: ${submittedCount} 条 (RM ${pendingAmount.toLocaleString()})\n`;
    report += `已发放: ${paidCount} 条 (RM ${paidAmount.toLocaleString()})\n\n`;
  }
  
  // 日志统计
  const logSheet = ss.getSheetByName('PaymentLog');
  if (logSheet && logSheet.getLastRow() > 1) {
    report += '【操作日志】\n';
    report += `总记录: ${logSheet.getLastRow() - 1} 条\n`;
  }
  
  report += '═'.repeat(40);
  
  ui.alert('统计报告', report, ui.ButtonSet.OK);
}

/**
 * 重置表格格式
 */
function resetAllFormatting() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert('🔄 重置格式', '重新应用所有表格的格式？\n\n这不会影响数据。', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActive();
  
  try {
    ss.toast('正在重置格式...', '🔄 格式重置', 30);
    
    for (const [sheetName, config] of Object.entries(SETUP_CONFIG.sheets)) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        applySheetFormatting_(sheet, config);
        addDataValidation_(sheet, config);
        addConditionalFormatting_(sheet, sheetName, config.headers);
      }
    }
    
    ui.alert('✅ 完成', '所有表格格式已重置！', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ 错误', error.message, ui.ButtonSet.OK);
  }
}

// ==================== 数据管理函数 ====================

/**
 * 备份所有数据
 */
function backupAllData() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '📦 备份数据',
    '此操作将为所有工作表创建备份副本。\n\n是否继续？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActive();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const backedUp = [];
  
  try {
    ss.toast('正在备份...', '📦 数据备份', 30);
    
    for (const sheetName of Object.keys(SETUP_CONFIG.sheets)) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const backupName = `${SETUP_CONFIG.backup.prefix}${sheetName}_${timestamp}`;
        sheet.copyTo(ss).setName(backupName);
        backedUp.push(backupName);
      }
    }
    
    logSetupOperation_('BACKUP', '数据备份', { sheets: backedUp.length, timestamp });
    
    ui.alert('✅ 备份完成', `已创建 ${backedUp.length} 个备份：\n\n${backedUp.join('\n')}`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ 备份失败', error.message, ui.ButtonSet.OK);
  }
}

/**
 * 恢复数据备份
 */
function restoreFromBackup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  // 查找所有备份
  const backups = [];
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name.startsWith(SETUP_CONFIG.backup.prefix)) {
      backups.push(name);
    }
  }
  
  if (backups.length === 0) {
    ui.alert('ℹ️ 提示', '没有找到任何备份', ui.ButtonSet.OK);
    return;
  }
  
  // 显示备份列表
  const backupList = backups.map((b, i) => `${i + 1}. ${b}`).join('\n');
  const response = ui.prompt(
    '📥 恢复备份',
    `可用备份：\n${backupList}\n\n请输入要恢复的备份编号（1-${backups.length}）：`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const index = parseInt(response.getResponseText()) - 1;
  if (isNaN(index) || index < 0 || index >= backups.length) {
    ui.alert('❌ 错误', '无效的编号', ui.ButtonSet.OK);
    return;
  }
  
  const backupName = backups[index];
  
  // 解析原表名
  const match = backupName.match(new RegExp(`^${SETUP_CONFIG.backup.prefix}(\\w+)_\\d+_\\d+$`));
  if (!match) {
    ui.alert('❌ 错误', '无法解析备份名称', ui.ButtonSet.OK);
    return;
  }
  
  const originalName = match[1];
  
  const confirm = ui.alert(
    '⚠️ 确认恢复',
    `将用 "${backupName}" 的数据覆盖 "${originalName}" 表。\n\n` +
    '当前数据将被替换，此操作不可撤销！\n\n是否继续？',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  try {
    const backupSheet = ss.getSheetByName(backupName);
    let targetSheet = ss.getSheetByName(originalName);
    
    if (targetSheet) {
      // 清空目标表（保留表头）
      if (targetSheet.getLastRow() > 1) {
        targetSheet.deleteRows(2, targetSheet.getLastRow() - 1);
      }
    } else {
      targetSheet = ss.insertSheet(originalName);
    }
    
    // 复制数据
    const data = backupSheet.getDataRange().getValues();
    if (data.length > 0) {
      targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }
    
    // 重新应用格式
    const config = SETUP_CONFIG.sheets[originalName];
    if (config) {
      applySheetFormatting_(targetSheet, config);
    }
    
    logSetupOperation_('RESTORE', '恢复备份', { backup: backupName, target: originalName });
    
    ui.alert('✅ 恢复完成', `已从 "${backupName}" 恢复数据到 "${originalName}"`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ 恢复失败', error.message, ui.ButtonSet.OK);
  }
}

/**
 * 导出为CSV
 */
function exportToCSV() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  // 选择要导出的表
  const sheetNames = Object.keys(SETUP_CONFIG.sheets).filter(name => {
    const sheet = ss.getSheetByName(name);
    return sheet && sheet.getLastRow() > 1;
  });
  
  if (sheetNames.length === 0) {
    ui.alert('ℹ️ 提示', '没有可导出的数据', ui.ButtonSet.OK);
    return;
  }
  
  const sheetList = sheetNames.map((n, i) => `${i + 1}. ${n}`).join('\n');
  const response = ui.prompt(
    '📤 导出CSV',
    `可导出的表：\n${sheetList}\n\n请输入要导出的编号（用逗号分隔多个，如 1,2,3）：`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const indices = response.getResponseText().split(',').map(s => parseInt(s.trim()) - 1);
  const toExport = indices.filter(i => !isNaN(i) && i >= 0 && i < sheetNames.length).map(i => sheetNames[i]);
  
  if (toExport.length === 0) {
    ui.alert('❌ 错误', '无效的选择', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const folder = DriveApp.getRootFolder();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const exportedFiles = [];
    
    for (const sheetName of toExport) {
      const sheet = ss.getSheetByName(sheetName);
      const data = sheet.getDataRange().getValues();
      
      // 转换为 CSV
      const csv = data.map(row => row.map(cell => {
        const str = String(cell);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',')).join('\n');
      
      const fileName = `${sheetName}_${timestamp}.csv`;
      const file = folder.createFile(fileName, csv, MimeType.CSV);
      exportedFiles.push(fileName);
    }
    
    ui.alert(
      '✅ 导出完成',
      `已导出 ${exportedFiles.length} 个文件到 Google Drive 根目录：\n\n${exportedFiles.join('\n')}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ 导出失败', error.message, ui.ButtonSet.OK);
  }
}

/**
 * 清理旧日志
 */
function cleanupOldLogs() {
  const ui = SpreadsheetApp.getUi();
  
  const daysResponse = ui.prompt(
    '🗑️ 清理旧日志',
    '保留最近多少天的日志？（默认30天）：',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (daysResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const days = parseInt(daysResponse.getResponseText()) || 30;
  
  const confirm = ui.alert(
    '⚠️ 确认',
    `将删除 ${days} 天前的所有日志。\n\n是否继续？`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PaymentLog');
  
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('ℹ️ 提示', '没有日志需要清理', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const timestampCol = headers.indexOf('Timestamp');
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    let deletedCount = 0;
    
    // 从后往前删除，避免行号错位
    for (let i = data.length - 1; i >= 1; i--) {
      const timestamp = new Date(data[i][timestampCol]);
      if (timestamp < cutoffDate) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    
    logSetupOperation_('CLEANUP_LOGS', '清理日志', { deleted: deletedCount, days });
    
    ui.alert('✅ 清理完成', `已删除 ${deletedCount} 条旧日志`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ 清理失败', error.message, ui.ButtonSet.OK);
  }
}

// ==================== 危险操作 ====================

/**
 * 清空所有数据
 */
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  
  const response1 = ui.alert(
    '⚠️ 危险操作',
    '将清空所有工作表的数据！\n\n此操作不可撤销！\n\n是否继续？',
    ui.ButtonSet.YES_NO
  );
  
  if (response1 !== ui.Button.YES) return;
  
  const response2 = ui.prompt(
    '⚠️ 二次确认',
    '请输入 "DELETE ALL" 确认清空所有数据：',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response2.getSelectedButton() !== ui.Button.OK || response2.getResponseText() !== 'DELETE ALL') {
    ui.alert('❌ 操作已取消', '确认文本不匹配', ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActive();
  
  try {
    for (const sheetName of Object.keys(SETUP_CONFIG.sheets)) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
    }
    
    logSetupOperation_('CLEAR_ALL', '清空所有数据', {});
    
    ui.alert('✅ 完成', '所有数据已清空！', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ 错误', error.message, ui.ButtonSet.OK);
  }
}

/**
 * 重置整个系统
 */
function resetEntireSystem() {
  const ui = SpreadsheetApp.getUi();
  
  const response1 = ui.alert(
    '💥 危险操作',
    '将删除所有工作表并重新初始化系统！\n\n' +
    '所有数据将永久丢失！\n\n' +
    '是否继续？',
    ui.ButtonSet.YES_NO
  );
  
  if (response1 !== ui.Button.YES) return;
  
  const response2 = ui.prompt(
    '💥 最终确认',
    '请输入 "RESET SYSTEM" 确认重置：',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response2.getSelectedButton() !== ui.Button.OK || response2.getResponseText() !== 'RESET SYSTEM') {
    ui.alert('❌ 操作已取消', '确认文本不匹配', ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.getActive();
  
  try {
    // 删除所有系统工作表
    for (const sheetName of Object.keys(SETUP_CONFIG.sheets)) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        ss.deleteSheet(sheet);
      }
    }
    
    // 删除说明表
    const instrSheet = ss.getSheetByName('📖 使用说明');
    if (instrSheet) {
      ss.deleteSheet(instrSheet);
    }
    
    // 重新初始化
    initializeSystem();
    
  } catch (error) {
    ui.alert('❌ 错误', error.message, ui.ButtonSet.OK);
  }
}

// ==================== 辅助函数 ====================

/**
 * 清理默认工作表
 */
function cleanupDefaultSheet_(ss) {
  const defaultNames = ['Sheet1', '工作表1', 'シート1', 'Hoja 1', 'Feuille 1'];
  for (const name of defaultNames) {
    const sheet = ss.getSheetByName(name);
    if (sheet && ss.getSheets().length > 1 && sheet.getLastRow() <= 1) {
      try {
        ss.deleteSheet(sheet);
      } catch (e) {
        // 忽略错误
      }
    }
  }
}

/**
 * 创建使用说明表
 */
function createInstructionSheet_(ss) {
  const sheetName = '📖 使用说明';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 500);
  
  const now = new Date().toLocaleString('zh-CN');
  
  const data = [
    ['薪资管理系统 v' + SETUP_CONFIG.version, ''],
    ['', ''],
    ['【系统角色】', ''],
    ['ADMIN (管理员)', '• 创建用户账号\n• 管理主管信息\n• 录入主管工资\n• 查看所有统计和日志'],
    ['SECRETARY (主管)', '• 管理员工信息\n• 录入员工工资\n• 处理员工离职\n• 查看自己录入的数据'],
    ['ACCOUNTANT (会计)', '• 查看已提交工资\n• 处理工资发放\n• 查看银行变更提醒\n• 导出银行汇款列表'],
    ['', ''],
    ['【工作表说明】', ''],
    ['Users', '系统用户账号（用户名、密码、角色）'],
    ['Staff', '员工信息（姓名、工资、银行账户、欠款）'],
    ['Managers', '主管信息（与员工类似）'],
    ['SalaryRecords', '月度工资记录（草稿→提交→发放）'],
    ['PaymentLog', '操作日志（所有关键操作都会记录）'],
    ['', ''],
    ['【工作流程】', ''],
    ['第一步', '主管添加员工 / 管理员添加主管'],
    ['第二步', '每月录入工资，保存为草稿'],
    ['第三步', '检查无误后提交工资记录'],
    ['第四步', '会计审核并标记已发放'],
    ['', ''],
    ['【安全特性】', ''],
    ['密码安全', '使用盐值+SHA256哈希，不存储明文密码'],
    ['输入验证', '所有用户输入都经过清理和验证'],
    ['权限控制', '每个操作都检查用户角色权限'],
    ['操作日志', '所有关键操作都会记录详细日志'],
    ['首次登录', '新用户首次登录必须修改密码'],
    ['', ''],
    ['【默认账户】', ''],
    ['用户名', SETUP_CONFIG.defaultAdmin.username],
    ['初始密码', SETUP_CONFIG.defaultAdmin.password],
    ['', '⚠️ 请首次登录后立即修改密码！'],
    ['', ''],
    ['【技术支持】', ''],
    ['菜单位置', '🛠️ 薪资系统（在菜单栏）'],
    ['系统检查', '使用「检查系统状态」功能'],
    ['问题修复', '使用「自动修复问题」功能'],
    ['数据备份', '定期使用「备份所有数据」功能'],
    ['', ''],
    ['创建时间', now],
    ['系统版本', SETUP_CONFIG.version]
  ];
  
  sheet.getRange(1, 1, data.length, 2).setValues(data);
  
  // 格式化标题
  sheet.getRange(1, 1, 1, 2).merge()
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#2D3748')
    .setFontColor('#FFFFFF');
  sheet.setRowHeight(1, 50);
  
  // 格式化小节标题
  const sectionRows = [3, 8, 15, 21, 28, 34];
  sectionRows.forEach(row => {
    if (row <= data.length) {
      sheet.getRange(row, 1)
        .setFontWeight('bold')
        .setFontColor('#4285F4')
        .setFontSize(12);
    }
  });
  
  // 设置换行
  sheet.getRange(1, 1, data.length, 2).setWrap(true);
  
  sheet.setTabColor('#607D8B');
  
  // 移到最后
  const sheets = ss.getSheets();
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(sheets.length);
}

/**
 * 记录设置操作日志
 */
function logSetupOperation_(action, description, details) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('PaymentLog');
    if (!sheet) return;
    
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const logId = 'SETUP-' + new Date().getTime();
    const detailStr = typeof details === 'object' ? JSON.stringify(details) : String(details);
    
    sheet.appendRow([
      logId,
      now,
      action,
      description,
      Session.getActiveUser().getEmail() || 'SYSTEM',
      'SYSTEM',
      'SETUP',
      detailStr,
      ''
    ]);
  } catch (e) {
    Logger.log('日志记录失败: ' + e.message);
  }
}
