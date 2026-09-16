// ============================================
// COMPLIANCE ACTIVITY TRACKER - BACKEND
// ============================================

var SPREADSHEET_ID = '1n7Ov-sVWSpACa2nTkKgu-MdvZddyKzl90iTEQkNR2-s';

// ============================================
// WEB APP ENTRY POINT
// ============================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Compliance Activity')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// AUTHENTICATION
// ============================================

function loginUser(identifier, password) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PIC Setup');
    if (!sheet) return { success: false, message: 'Database not found' };
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: 'No users found' };
    
    var headers = data[0];
    var userCol = headers.indexOf('Username');
    var emailCol = headers.indexOf('Email');
    var passCol = headers.indexOf('Password');
    var statusCol = headers.indexOf('Status');
    var nameCol = headers.indexOf('Nama Lengkap');
    var roleCol = headers.indexOf('Role');
    var jabatanCol = headers.indexOf('Jabatan');
    
    var cleanInput = identifier.toString().trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      var rowUsername = (userCol != -1 && data[i][userCol]) ? data[i][userCol].toString().trim().toLowerCase() : '';
      var rowEmail = (emailCol != -1 && data[i][emailCol]) ? data[i][emailCol].toString().trim().toLowerCase() : '';
      
      if (rowUsername === cleanInput || rowEmail === cleanInput) {
        var status = (statusCol != -1 && data[i][statusCol]) ? data[i][statusCol].toString() : '';
        if (status !== 'Aktif' && status !== 'Active') {
          return { success: false, message: 'inactive' };
        }
        
        var storedPass = (passCol != -1 && data[i][passCol]) ? data[i][passCol].toString() : '';
        if (storedPass !== password) {
          return { success: false, message: 'password_wrong' };
        }
        
        var user = {
          name: (nameCol != -1 && data[i][nameCol]) ? data[i][nameCol] : 'User',
          username: rowUsername,
          email: rowEmail,
          role: (roleCol != -1 && data[i][roleCol]) ? data[i][roleCol] : 'User',
          jabatan: (jabatanCol != -1 && data[i][jabatanCol]) ? data[i][jabatanCol] : '',
          rowIndex: i + 1
        };
        
        PropertiesService.getUserProperties().setProperty('currentUser', JSON.stringify(user));
        return { success: true, user: user };
      }
    }
    return { success: false, message: 'user_not_found' };
  } catch (e) {
    return { success: false, message: 'system_error: ' + e.toString() };
  }
}

function logoutUser() {
  PropertiesService.getUserProperties().deleteProperty('currentUser');
  return { success: true };
}

// ============================================
// MENU BY ROLE (DYNAMIC)
// ============================================

function getMenusByRole(roleName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Roles');
    if (!sheet) return getDefaultMenus(roleName);
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var roleCol = headers.indexOf('Nama Role');
    var permCol = headers.indexOf('Menu Access');
    var statusCol = headers.indexOf('Status');
    
    if (roleCol === -1 || permCol === -1) return getDefaultMenus(roleName);
    
    for (var i = 1; i < data.length; i++) {
      var rowRole = data[i][roleCol] ? data[i][roleCol].toString() : '';
      var rowStatus = data[i][statusCol] ? data[i][statusCol].toString() : '';
      
      if (rowRole === roleName && (rowStatus === 'Aktif' || rowStatus === 'Active')) {
        try {
          var permString = data[i][permCol];
          var permissions = typeof permString === 'string' ? JSON.parse(permString) : permString;
          return buildMenusFromPermissions(permissions);
        } catch(e) {
          return getDefaultMenus(roleName);
        }
      }
    }
    return getDefaultMenus(roleName);
  } catch (e) {
    return getDefaultMenus(roleName);
  }
}

function buildMenusFromPermissions(permissions) {
  var menus = { main: [], task: [], library: [], reporting: [], configuration: [] };
  
  if (permissions['Dashboard'] && permissions['Dashboard'].view) menus.main.push({ id: 'dashboard', name: 'Dashboard', icon: 'dashboard' });
  if (permissions['Task Management'] && permissions['Task Management'].view) menus.task.push({ id: 'tasks', name: 'Task Management', icon: 'assignment' });
  
  if (permissions['Repository'] && permissions['Repository'].view) {
    menus.library.push({ id: 'repo-policy', name: 'Policy & Procedure', icon: 'policy' });
    menus.library.push({ id: 'repo-bi', name: 'BI Regulation', icon: 'account_balance' });
    menus.library.push({ id: 'repo-directory', name: 'Directory', icon: 'folder' });
  }
  
  if (permissions['Reporting OKR'] && permissions['Reporting OKR'].view) {
    menus.reporting.push({ id: 'team-okr', name: 'Reporting OKR', icon: 'trending_up' });
  }
  
  var hasUserMgmt = permissions['User Management'] && permissions['User Management'].view;
  var hasRoleSetup = permissions['Role Setup'] && permissions['Role Setup'].view;
  var hasMenuSettings = permissions['Menu Settings'] && permissions['Menu Settings'].view;
  
  if (hasUserMgmt || hasRoleSetup || hasMenuSettings) {
    if (hasUserMgmt) menus.configuration.push({ id: 'config-users', name: 'User Management', icon: 'group' });
    if (hasRoleSetup) menus.configuration.push({ id: 'setup-role', name: 'Role Setup', icon: 'shield' });
    if (hasMenuSettings) menus.configuration.push({ id: 'config-menu', name: 'Menu Settings', icon: 'menu_open' });
  }
  
  return menus;
}

function getDefaultMenus(roleName) {
  return {
    main: [{ id: 'dashboard', name: 'Dashboard', icon: 'dashboard' }],
    task: [{ id: 'tasks', name: 'Task Management', icon: 'assignment' }],
    library: [],
    reporting: [{ id: 'team-okr', name: 'Reporting OKR', icon: 'trending_up' }],
    configuration: roleName === 'Superadmin' ? [
      { id: 'config-users', name: 'User Management', icon: 'group' },
      { id: 'setup-role', name: 'Role Setup', icon: 'shield' },
      { id: 'config-menu', name: 'Menu Settings', icon: 'menu_open' }
    ] : []
  };
}

// ============================================
// USER MANAGEMENT
// ============================================

function getAllUsers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('PIC Setup');
    if (!sheet) {
      sheet = ss.insertSheet('PIC Setup');
      sheet.appendRow(['Nama Lengkap', 'Username', 'Email', 'Password', 'No Telepon', 'Telegram ID', 'Role', 'Jabatan', 'Status', 'Foto Profil', 'Tanggal Buat']);
    }
    
    var headers = sheet.getDataRange().getValues()[0];
    if (headers.indexOf('Jabatan') === -1) {
      sheet.getRange(1, headers.length + 1).setValue('Jabatan');
      headers.push('Jabatan');
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, users: [] };
    
    var users = [];
    for (var i = 1; i < data.length; i++) {
      var user = { rowIndex: i + 1 };
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        user[headers[j]] = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy') : val;
      }
      users.push(user);
    }
    return { success: true, users: users };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function createUser(userData) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (currentUser.role !== 'Superadmin') return { success: false, message: 'Access denied' };
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PIC Setup');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var userCol = headers.indexOf('Username');
    var emailCol = headers.indexOf('Email');
    for (var i = 1; i < data.length; i++) {
      if (data[i][userCol] && data[i][userCol].toString().toLowerCase() === userData.Username.toLowerCase()) {
        return { success: false, message: 'Username is already in use!' };
      }
      if (data[i][emailCol] && data[i][emailCol].toString().toLowerCase() === userData.Email.toLowerCase()) {
        return { success: false, message: 'Email is already registered!' };
      }
    }
    
    var password = userData.TempPassword || generatePassword(10);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    
    var newRow = [];
    headers.forEach(function(h) {
      if (h === 'Nama Lengkap') newRow.push(userData['Nama Lengkap']);
      else if (h === 'Username') newRow.push(userData.Username);
      else if (h === 'Email') newRow.push(userData.Email);
      else if (h === 'Password') newRow.push(password);
      else if (h === 'No Telepon') newRow.push(userData['No Telepon'] || '');
      else if (h === 'Telegram ID') newRow.push(userData['Telegram ID'] || '');
      else if (h === 'Role') newRow.push(userData.Role);
      else if (h === 'Jabatan') newRow.push(userData.Jabatan || '');
      else if (h === 'Status') newRow.push('Aktif');
      else if (h === 'Tanggal Buat') newRow.push(today);
      else newRow.push('');
    });
    
    sheet.appendRow(newRow);
    return { success: true, message: 'User added successfully. Password: ' + password };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function updateUser(rowIndex, userData) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (currentUser.role !== 'Superadmin') return { success: false, message: 'Access denied' };
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PIC Setup');
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return { success: false, message: 'Invalid row' };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    if (headers.indexOf('Jabatan') === -1) {
      sheet.getRange(1, headers.length + 1).setValue('Jabatan');
      headers.push('Jabatan');
    }

    var updateMap = {
      'Nama Lengkap': userData['Nama Lengkap'], 
      'Username': userData.Username,
      'Email': userData.Email, 
      'No Telepon': userData['No Telepon'],
      'Telegram ID': userData['Telegram ID'], 
      'Role': userData.Role, 
      'Jabatan': userData.Jabatan || '',
      'Status': userData.Status
    };
    
    for (var key in updateMap) {
      var colIndex = headers.indexOf(key);
      if (colIndex !== -1) sheet.getRange(rowIndex, colIndex + 1).setValue(updateMap[key]);
    }
    return { success: true, message: 'User updated successfully' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function resetUserPassword(rowIndex) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PIC Setup');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var passCol = headers.indexOf('Password');
    var emailCol = headers.indexOf('Email');
    
    var newPassword = generatePassword(10);
    sheet.getRange(rowIndex, passCol + 1).setValue(newPassword);
    
    return { success: true, message: 'Password reset successfully. New password: ' + newPassword };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ============================================
// ROLE MANAGEMENT & PERMISSIONS
// ============================================

function getAllRoles() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Roles');
    if (!sheet) {
      sheet = ss.insertSheet('Roles');
      sheet.appendRow(['Nama Role', 'Deskripsi', 'Menu Access', 'Status', 'Tanggal Buat']);
      sheet.appendRow(['Superadmin', 'Full access', '{"Dashboard":{"view":true,"create":true,"edit":true,"delete":true,"review":true,"approve":true},"Task Management":{"view":true,"create":true,"edit":true,"delete":true,"review":true,"approve":true},"Repository":{"view":true,"create":true,"edit":true,"delete":true,"review":true,"approve":true},"Reporting OKR":{"view":true,"create":true,"edit":true,"delete":true,"review":true,"approve":true},"User Management":{"view":true,"create":true,"edit":true,"delete":true,"review":true,"approve":true},"Role Setup":{"view":true,"create":true,"edit":true,"delete":true,"review":true,"approve":true},"Menu Settings":{"view":true,"create":true,"edit":true,"delete":true,"review":true,"approve":true}}', 'Aktif', new Date()]);
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, roles: [] };
    
    var headers = data[0];
    var roles = [];
    for (var i = 1; i < data.length; i++) {
      var role = { rowIndex: i + 1 };
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        role[headers[j]] = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy') : val;
      }
      roles.push(role);
    }
    return { success: true, roles: roles };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function createRole(roleData) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (currentUser.role !== 'Superadmin') return { success: false, message: 'Access denied' };
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Roles');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var nameCol = headers.indexOf('Nama Role');
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][nameCol] && data[i][nameCol].toString().toLowerCase() === roleData['Nama Role'].toLowerCase()) {
        return { success: false, message: 'Role already exists!' };
      }
    }
    
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    sheet.appendRow([roleData['Nama Role'], roleData.Deskripsi || '', roleData['Menu Access'] || '{}', roleData.Status || 'Aktif', today]);
    return { success: true, message: 'Role added successfully' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function updateRole(rowIndex, roleData) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (currentUser.role !== 'Superadmin') return { success: false, message: 'Access denied' };
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Roles');
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return { success: false, message: 'Invalid row' };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var updateMap = {
      'Nama Role': roleData['Nama Role'], 
      'Deskripsi': roleData.Deskripsi,
      'Menu Access': roleData['Menu Access'], 
      'Status': roleData.Status
    };
    
    for (var key in updateMap) {
      var colIndex = headers.indexOf(key);
      if (colIndex !== -1) sheet.getRange(rowIndex, colIndex + 1).setValue(updateMap[key]);
    }
    return { success: true, message: 'Role updated successfully' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function getUserPermissions() {
  try {
    var userProp = PropertiesService.getUserProperties().getProperty('currentUser');
    if (!userProp) return null;
    
    var currentUser = JSON.parse(userProp);
    if (!currentUser || !currentUser.role) return null;
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Roles');
    if (!sheet) return null;
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var roleCol = headers.indexOf('Nama Role');
    var permCol = headers.indexOf('Menu Access');
    
    if (roleCol === -1 || permCol === -1) return null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][roleCol] === currentUser.role) {
        try {
          return JSON.parse(data[i][permCol]);
        } catch(e) {
          return null;
        }
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

// ============================================
// OKR MANAGEMENT
// ============================================

function getOKRData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('OKR');
    if (!sheet) {
      sheet = ss.insertSheet('OKR');
      sheet.appendRow(['Objective', 'Objective1', 'Objective2', 'Objective3', 'Objective4', 'Objective5', 'KeyResults', 'KeyResult1', 'KeyResult2', 'KeyResult3', 'KeyResult4', 'KeyResult5', 'Initiative1', 'Initiative2', 'Initiative3', 'Initiative4', 'Initiative5', 'Name', 'Position', 'Department', 'Quarter', 'Progress', 'Status', 'ApprovalStatus', 'ReviewNotes', 'CreatedAt', 'CreatedBy', 'CreatorRole', 'ManagerName', 'ManagerTitle', 'DirectorName', 'DirectorTitle', 'ApprovedDate']);
      sheet.getRange(1, 1, 1, 33).setFontWeight('bold').setBackground('#1e293b');
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, okrs: [] };
    var headers = data[0];
    var okrs = [];
    for (var i = 1; i < data.length; i++) {
      var okr = { rowIndex: i + 1 };
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        okr[headers[j]] = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy') : val;
      }
      if (!okr.ApprovalStatus) okr.ApprovalStatus = 'Pending Manager';
      okrs.push(okr);
    }
    return { success: true, okrs: okrs };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function saveOKR(rowIndex, okrData) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('OKR');
    if (!sheet) {
      sheet = ss.insertSheet('OKR');
      sheet.appendRow(['Objective','Objective1','Objective2','Objective3','Objective4','Objective5','KeyResults','KeyResult1','KeyResult2','KeyResult3','KeyResult4','KeyResult5','Initiative1','Initiative2','Initiative3','Initiative4','Initiative5','Name','Position','Department','Quarter','Progress','Status','ApprovalStatus','ReviewNotes','CreatedAt','CreatedBy','CreatorRole','ManagerName','ManagerTitle','DirectorName','DirectorTitle','ApprovedDate']);
      sheet.getRange(1,1,1,33).setFontWeight('bold').setBackground('#1e293b');
    }
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    var headers = sheet.getDataRange().getValues()[0];
    if (rowIndex) {
      var updateMap = { 
        'Objective': okrData.Objective, 'Objective1': okrData.Objective1, 'Objective2': okrData.Objective2, 
        'Objective3': okrData.Objective3, 'Objective4': okrData.Objective4, 'Objective5': okrData.Objective5, 
        'KeyResults': okrData.KeyResults, 'KeyResult1': okrData.KeyResult1, 'KeyResult2': okrData.KeyResult2, 
        'KeyResult3': okrData.KeyResult3, 'KeyResult4': okrData.KeyResult4, 'KeyResult5': okrData.KeyResult5, 
        'Initiative1': okrData.Initiative1, 'Initiative2': okrData.Initiative2, 'Initiative3': okrData.Initiative3, 
        'Initiative4': okrData.Initiative4, 'Initiative5': okrData.Initiative5, 'Name': okrData.Name, 
        'Position': okrData.Position, 'Department': okrData.Department, 'Quarter': okrData.Quarter, 
        'Progress': okrData.Progress, 'Status': okrData.Status 
      };
      for (var key in updateMap) { 
        var colIndex = headers.indexOf(key); 
        if (colIndex !== -1) sheet.getRange(rowIndex, colIndex + 1).setValue(updateMap[key]); 
      }
      return { success: true, message: 'OKR updated' };
    } else {
      var newRow = [];
      headers.forEach(function(header) {
        if (header === 'ApprovalStatus') newRow.push('Pending Manager');
        else if (header === 'CreatorRole') newRow.push(currentUser.role);
        else if (header === 'CreatedBy') newRow.push(currentUser.name);
        else if (header === 'CreatedAt') newRow.push(today);
        else if (okrData[header] !== undefined) newRow.push(okrData[header]);
        else newRow.push('');
      });
      sheet.appendRow(newRow);
      return { success: true, message: 'OKR created' };
    }
  } catch (e) { 
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function approveOKR(rowIndex, action, notes) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('OKR');
    if (!sheet) return { success: false, message: 'Sheet OKR tidak ditemukan' };

    var headers = sheet.getDataRange().getValues()[0];
    var approvalCol = headers.indexOf('ApprovalStatus');
    var notesCol = headers.indexOf('ReviewNotes');
    var approvedDateCol = headers.indexOf('ApprovedDate');
    var mgrNameCol = headers.indexOf('ManagerName');
    var mgrTitleCol = headers.indexOf('ManagerTitle');
    var mgrDateCol = headers.indexOf('ManagerDate');
    var dirNameCol = headers.indexOf('DirectorName');
    var dirTitleCol = headers.indexOf('DirectorTitle');
    var dirDateCol = headers.indexOf('DirectorDate');

    var newStatus = action;
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    var timestamp = new Date().toLocaleString();

    // Update Status
    if (approvalCol !== -1) {
      sheet.getRange(rowIndex, approvalCol + 1).setValue(newStatus);
    }

    // Update Manager Info
    if (newStatus === 'Pending Director') {
      var userJabatanMgr = currentUser.jabatan || currentUser.Role || '';
      if (mgrNameCol !== -1) sheet.getRange(rowIndex, mgrNameCol + 1).setValue(currentUser.name);
      if (mgrTitleCol !== -1) sheet.getRange(rowIndex, mgrTitleCol + 1).setValue(userJabatanMgr);
      if (mgrDateCol !== -1) sheet.getRange(rowIndex, mgrDateCol + 1).setValue(today);
    }

    // Update Director Info
    if (newStatus === 'Approved') {
      var userJabatan = currentUser.jabatan || currentUser.Role || '';
      if (dirNameCol !== -1) sheet.getRange(rowIndex, dirNameCol + 1).setValue(currentUser.name);
      if (dirTitleCol !== -1) sheet.getRange(rowIndex, dirTitleCol + 1).setValue(userJabatan);
      if (dirDateCol !== -1) sheet.getRange(rowIndex, dirDateCol + 1).setValue(today);
      if (approvedDateCol !== -1) sheet.getRange(rowIndex, approvedDateCol + 1).setValue(today);
    }

    // Update Notes
    if (notesCol !== -1 && notes) {
      var existingNotes = sheet.getRange(rowIndex, notesCol + 1).getValue() || '';
      var newNoteEntry = '[' + timestamp + '] ' + currentUser.name + ' (' + (currentUser.jabatan || currentUser.Role) + '): ' + notes + '\n';
      sheet.getRange(rowIndex, notesCol + 1).setValue(newNoteEntry + existingNotes);
    }

    return { success: true, message: 'Status berhasil diupdate ke ' + newStatus };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function closeOKR(rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('OKR');
    if (!sheet) return { success: false, message: 'OKR sheet not found' };
    var headers = sheet.getDataRange().getValues()[0];
    var approvalCol = headers.indexOf('ApprovalStatus');
    if (approvalCol !== -1) sheet.getRange(rowIndex, approvalCol + 1).setValue('Closed');
    return { success: true, message: 'OKR closed' };
  } catch (e) { 
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

function deleteOKR(rowIndex) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (currentUser.role !== 'Superadmin') return { success: false, message: 'Access denied' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('OKR');
    if (!sheet) return { success: false, message: 'OKR sheet not found' };
    rowIndex = parseInt(rowIndex);
    if (isNaN(rowIndex) || rowIndex < 2) return { success: false, message: 'Invalid row' };
    sheet.deleteRow(rowIndex);
    return { success: true, message: 'OKR deleted' };
  } catch (e) { 
    return { success: false, message: 'Error: ' + e.toString() }; 
  }
}

// ============================================
// APPROVAL CONFIGURATION (DINAMIS)
// ============================================

function getApprovalConfig(requesterRole) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Approval_Config');
    if (!sheet) {
      // Default config jika sheet belum ada
      return getDefaultApprovalConfig(requesterRole);
    }
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var roleCol = headers.indexOf('Requester Role');
    for (var i = 1; i < data.length; i++) {
      if (data[i][roleCol] && data[i][roleCol].toString().toLowerCase() === requesterRole.toLowerCase()) {
        return {
          step1: data[i][1] || '',
          step2: data[i][2] || '',
          step3: data[i][3] || ''
        };
      }
    }
    return getDefaultApprovalConfig(requesterRole);
  } catch (e) {
    return getDefaultApprovalConfig(requesterRole);
  }
}

function getDefaultApprovalConfig(requesterRole) {
  var role = requesterRole.toLowerCase();
  // Manager & Direktur: skip Manager Review, langsung ke Director
  if (role === 'manager' || role === 'direktur' || role === 'director') {
    return { step1: 'Director', step2: '', step3: '' };
  }
  // Staff & Head: butuh Manager Review dulu, lalu Director
  return { step1: 'Manager', step2: 'Director', step3: '' };
}

// ============================================
// APPROVE OKR (DIPERBAIKI - NO ERROR)
// ============================================

function approveOKR(rowIndex, action, notes) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('OKR');
    if (!sheet) return { success: false, message: 'Sheet OKR tidak ditemukan' };

    var headers = sheet.getDataRange().getValues()[0];
    
    // Helper function untuk update kolom dengan aman
    function updateColumn(colName, value) {
      var colIndex = headers.indexOf(colName);
      if (colIndex !== -1 && colIndex >= 0) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(value);
        return true;
      }
      return false;
    }

    // 1. Update Status
    updateColumn('ApprovalStatus', action);

    // 2. Update Approved Date & Nama Director jika Approved
    if (action === 'Approved') {
      var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
      updateColumn('ApprovedDate', today);
      
      var userJabatan = currentUser.jabatan || currentUser.Role || '';
      updateColumn('DirectorName', currentUser.name);
      updateColumn('DirectorTitle', userJabatan);
    } 
    // 3. Update Nama Manager jika forward ke Director
    else if (action === 'Pending Director') {
      var userJabatanMgr = currentUser.jabatan || currentUser.Role || '';
      updateColumn('ManagerName', currentUser.name);
      updateColumn('ManagerTitle', userJabatanMgr);
    }

    // 4. Update Notes (tambah di paling atas)
    if (notes) {
      var notesCol = headers.indexOf('ReviewNotes');
      if (notesCol !== -1 && notesCol >= 0) {
        var timestamp = new Date().toLocaleString();
        var existingNotes = sheet.getRange(rowIndex, notesCol + 1).getValue() || '';
        var newNoteEntry = '[' + timestamp + '] ' + currentUser.name + ': ' + notes + '\n';
        sheet.getRange(rowIndex, notesCol + 1).setValue(newNoteEntry + existingNotes);
      }
    }

    return { success: true, message: 'Status berhasil diupdate ke ' + action };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ============================================
// SAVE OKR (DIPERBAIKI - AUTO STATUS BERDASARKAN ROLE)
// ============================================

function saveOKR(rowIndex, okrData) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('OKR');
    if (!sheet) {
      sheet = ss.insertSheet('OKR');
      sheet.appendRow(['Objective','Objective1','Objective2','Objective3','Objective4','Objective5','KeyResults','KeyResult1','KeyResult2','KeyResult3','KeyResult4','KeyResult5','Initiative1','Initiative2','Initiative3','Initiative4','Initiative5','Name','Position','Department','Quarter','Progress','Status','ApprovalStatus','ReviewNotes','CreatedAt','CreatedBy','CreatorRole','ManagerName','ManagerTitle','DirectorName','DirectorTitle','ApprovedDate']);
      sheet.getRange(1,1,1,33).setFontWeight('bold').setBackground('#1e293b');
    }
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    var headers = sheet.getDataRange().getValues()[0];
    
    // Helper function
    function updateColumn(colName, value) {
      var colIndex = headers.indexOf(colName);
      if (colIndex !== -1 && colIndex >= 0) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(value);
      }
    }
    
    if (rowIndex) {
      var updateMap = {
        'Objective': okrData.Objective,
        'Objective1': okrData.Objective1,
        'Objective2': okrData.Objective2,
        'Objective3': okrData.Objective3,
        'Objective4': okrData.Objective4,
        'Objective5': okrData.Objective5,
        'KeyResults': okrData.KeyResults,
        'KeyResult1': okrData.KeyResult1,
        'KeyResult2': okrData.KeyResult2,
        'KeyResult3': okrData.KeyResult3,
        'KeyResult4': okrData.KeyResult4,
        'KeyResult5': okrData.KeyResult5,
        'Initiative1': okrData.Initiative1,
        'Initiative2': okrData.Initiative2,
        'Initiative3': okrData.Initiative3,
        'Initiative4': okrData.Initiative4,
        'Initiative5': okrData.Initiative5,
        'Name': okrData.Name,
        'Position': okrData.Position,
        'Department': okrData.Department,
        'Quarter': okrData.Quarter,
        'Progress': okrData.Progress,
        'Status': okrData.Status
      };
      for (var key in updateMap) {
        updateColumn(key, updateMap[key]);
      }
      return { success: true, message: 'OKR updated' };
    } else {
      // Tentukan status approval berdasarkan role pembuat
      var creatorRole = currentUser.role || 'Staff';
      var initialApprovalStatus = 'Pending Review'; // Default
      
      // Manager & Direktur: skip Manager Review, langsung ke Director
      if (creatorRole.toLowerCase() === 'manager' || 
          creatorRole.toLowerCase() === 'direktur' || 
          creatorRole.toLowerCase() === 'director') {
        initialApprovalStatus = 'Pending Director';
      }
      
      var newRow = [];
      headers.forEach(function(header) {
        if (header === 'ApprovalStatus') newRow.push(initialApprovalStatus);
        else if (header === 'CreatorRole') newRow.push(creatorRole);
        else if (header === 'CreatedBy') newRow.push(currentUser.name);
        else if (header === 'CreatedAt') newRow.push(today);
        else if (header === 'ApprovedDate') newRow.push('');
        else if (header === 'ManagerName') newRow.push('');
        else if (header === 'ManagerTitle') newRow.push('');
        else if (header === 'DirectorName') newRow.push('');
        else if (header === 'DirectorTitle') newRow.push('');
        else if (header === 'ReviewNotes') newRow.push('');
        else if (okrData[header] !== undefined) newRow.push(okrData[header]);
        else newRow.push('');
      });
      sheet.appendRow(newRow);
      return { success: true, message: 'OKR created with status: ' + initialApprovalStatus };
    }
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ============================================
// GET OKR DATA (DIPERBAIKI)
// ============================================

function getOKRData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('OKR');
    if (!sheet) {
      sheet = ss.insertSheet('OKR');
      sheet.appendRow(['Objective','Objective1','Objective2','Objective3','Objective4','Objective5','KeyResults','KeyResult1','KeyResult2','KeyResult3','KeyResult4','KeyResult5','Initiative1','Initiative2','Initiative3','Initiative4','Initiative5','Name','Position','Department','Quarter','Progress','Status','ApprovalStatus','ReviewNotes','CreatedAt','CreatedBy','CreatorRole','ManagerName','ManagerTitle','DirectorName','DirectorTitle','ApprovedDate']);
      sheet.getRange(1,1,1,33).setFontWeight('bold').setBackground('#1e293b');
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, okrs: [] };
    var headers = data[0];
    var okrs = [];
    for (var i = 1; i < data.length; i++) {
      var okr = { rowIndex: i + 1 };
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        okr[headers[j]] = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy') : val;
      }
      if (!okr.ApprovalStatus) okr.ApprovalStatus = 'Pending Review';
      okrs.push(okr);
    }
    return { success: true, okrs: okrs };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ============================================
// DELETE OKR
// ============================================

function deleteOKR(rowIndex) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (currentUser.role !== 'Superadmin') return { success: false, message: 'Access denied' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('OKR');
    if (!sheet) return { success: false, message: 'OKR sheet not found' };
    rowIndex = parseInt(rowIndex);
    if (isNaN(rowIndex) || rowIndex < 2) return { success: false, message: 'Invalid row' };
    sheet.deleteRow(rowIndex);
    return { success: true, message: 'OKR deleted' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ============================================
// REPOSITORY MANAGEMENT
// ============================================

function getRepoData(pageId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = getRepoSheetName(pageId);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      createRepoSheetHeaders(sheet, pageId);
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, documents: [] };
    var headers = data[0];
    var documents = [];
    for (var i = 1; i < data.length; i++) {
      var doc = { rowIndex: i + 1 };
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        doc[headers[j]] = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy') : val;
      }
      documents.push(doc);
    }
    return { success: true, documents: documents };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function getRepoSheetName(pageId) {
  if (pageId === 'repo-policy') return 'Repo_Policy_Procedure';
  if (pageId === 'repo-bi') return 'Repo_BI_Regulation';
  if (pageId === 'repo-directory') return 'Repo_Directory';
  return 'Repo_Unknown';
}

function createRepoSheetHeaders(sheet, pageId) {
  if (pageId === 'repo-policy') {
    sheet.appendRow(['No', 'DocType', 'DocumentName', 'DocNumber', 'DriveLink', 'Department', 'Owner', 'Status', 'EffectiveDate', 'CreatedBy', 'CreatedDate']);
  } else if (pageId === 'repo-bi') {
    sheet.appendRow(['No', 'NumberDoc', 'NameDocument', 'Status', 'Description', 'DriveLink', 'CreatedBy', 'CreatedDate']);
  } else if (pageId === 'repo-directory') {
    sheet.appendRow(['No', 'NumberDoc', 'NameDocument', 'Status', 'Description', 'DriveLink', 'CreatedBy', 'CreatedDate']);
  }
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#1e293b');
}

function saveRepoDocument(pageId, rowIndex, docData) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (currentUser.role !== 'Superadmin') return { success: false, message: 'Access denied' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = getRepoSheetName(pageId);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      createRepoSheetHeaders(sheet, pageId);
    }
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    var headers = sheet.getDataRange().getValues()[0];
    if (rowIndex) {
      for (var key in docData) {
        var colIndex = headers.indexOf(key);
        if (colIndex !== -1) sheet.getRange(rowIndex, colIndex + 1).setValue(docData[key]);
      }
      return { success: true, message: 'Document updated successfully' };
    } else {
      var newRow = [];
      headers.forEach(function(header) {
        if (header === 'DriveLink') newRow.push(docData.DriveLink || '');
        else if (header === 'CreatedBy') newRow.push(currentUser.name);
        else if (header === 'CreatedDate') newRow.push(today);
        else if (header === 'No') newRow.push(sheet.getLastRow());
        else newRow.push(docData[header] || '');
      });
      sheet.appendRow(newRow);
      return { success: true, message: 'Document added successfully' };
    }
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function deleteRepoDocument(type, rowIndex) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (currentUser.role !== 'Superadmin') return { success: false, message: 'Access denied' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var pageId = 'repo-' + type;
    var sheetName = getRepoSheetName(pageId);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet not found' };
    rowIndex = parseInt(rowIndex);
    if (isNaN(rowIndex) || rowIndex < 2) return { success: false, message: 'Invalid row' };
    sheet.deleteRow(rowIndex);
    return { success: true, message: 'Document deleted successfully' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function getRepoFolder() {
  try {
    var folders = DriveApp.getFoldersByName('Repository Documents');
    return folders.hasNext() ? folders.next() : DriveApp.createFolder('Repository Documents');
  } catch (e) {
    return DriveApp.getRootFolder();
  }
}

// ============================================
// TASK MANAGEMENT
// ============================================

var TASK_SHEET_NAME = 'Tasks';

function getTaskData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TASK_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(TASK_SHEET_NAME);
      var headers = ['SubTaskID', 'ParentMainTask', 'CreatedDate', 'Month', 'Year', 'DueDate', 'TargetDuration', 'ActualFinish', 'DeadlineStatus', 'Priority', 'AssignedBy', 'DelegationTo', 'Category', 'ActivityType', 'BusinessModel', 'MainTask', 'DetailTask', 'StatusTask', 'ReviewerNotes', 'ActionLog', 'Progress', 'StatusCode', 'ReviewedBy', 'ApprovalBy', 'LastUpdate', 'LastUpdatedBy', 'NotesLog'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1e293b');
      sheet.setFrozenRows(1);
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, tasks: [] };
    var headers = data[0];
    var tasks = [];
    for (var i = 1; i < data.length; i++) {
      var task = { rowIndex: i + 1 };
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        if (headers[j] === 'NotesLog') {
          try { task[headers[j]] = JSON.parse(val); } catch(e) { task[headers[j]] = []; }
        } else {
          task[headers[j]] = (val instanceof Date) ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy') : val;
        }
      }
      tasks.push(task);
    }
    return { success: true, tasks: tasks };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function saveTask(rowIndex, taskData) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (!currentUser || !currentUser.name) return { success: false, message: 'User not authenticated' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TASK_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Tasks sheet not found' };
    var headers = sheet.getDataRange().getValues()[0];
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    var now = new Date();
    var month = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM');
    var year = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy');
    var deadlineStatus = calculateDeadlineStatus(taskData.DueDate, taskData.ActualFinish, taskData.StatusTask);
    var statusCode = generateStatusCode(taskData.StatusTask);
    
    if (rowIndex) {
      var updateMap = {
        'MainTask': taskData.MainTask, 'DetailTask': taskData.DetailTask, 'Category': taskData.Category,
        'ActivityType': taskData.ActivityType, 'AssignedBy': taskData.AssignedBy, 'DelegationTo': taskData.DelegationTo,
        'Priority': taskData.Priority, 'DueDate': taskData.DueDate, 'TargetDuration': taskData.TargetDuration,
        'BusinessModel': taskData.BusinessModel, 'StatusTask': taskData.StatusTask, 'Progress': taskData.Progress,
        'ActualFinish': taskData.ActualFinish, 'ReviewerNotes': taskData.ReviewerNotes, 'ActionLog': taskData.ActionLog,
        'ReviewedBy': taskData.ReviewedBy, 'ApprovalBy': taskData.ApprovalBy, 'DeadlineStatus': deadlineStatus,
        'StatusCode': statusCode, 'LastUpdate': today, 'LastUpdatedBy': currentUser.name
      };
      for (var key in updateMap) {
        var colIndex = headers.indexOf(key);
        if (colIndex !== -1) sheet.getRange(rowIndex, colIndex + 1).setValue(updateMap[key]);
      }
      return { success: true, message: 'Task updated successfully' };
    } else {
      var subTaskId = generateSubTaskId(sheet);
      var newRow = [];
      headers.forEach(function(header) {
        if (header === 'SubTaskID') newRow.push(subTaskId);
        else if (header === 'CreatedDate') newRow.push(today);
        else if (header === 'Month') newRow.push(month);
        else if (header === 'Year') newRow.push(year);
        else if (header === 'LastUpdate') newRow.push(today);
        else if (header === 'LastUpdatedBy') newRow.push(currentUser.name);
        else if (header === 'DeadlineStatus') newRow.push(deadlineStatus);
        else if (header === 'StatusCode') newRow.push(statusCode);
        else if (header === 'NotesLog') newRow.push('[]');
        else if (taskData[header] !== undefined) newRow.push(taskData[header]);
        else newRow.push('');
      });
      sheet.appendRow(newRow);
      return { success: true, message: 'Task created successfully' };
    }
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function deleteTask(rowIndex) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    if (!currentUser) return { success: false, message: 'Access denied' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TASK_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Tasks sheet not found' };
    rowIndex = parseInt(rowIndex);
    if (isNaN(rowIndex) || rowIndex < 2) return { success: false, message: 'Invalid row' };
    sheet.deleteRow(rowIndex);
    return { success: true, message: 'Task deleted successfully' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function calculateDeadlineStatus(dueDate, actualFinish, statusTask) {
  if (statusTask === 'Completed' && actualFinish) return 'On Time';
  if (!dueDate) return 'Not Set';
  try {
    var due = new Date(dueDate);
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var diffTime = due - today;
    var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    if (diffDays < 0) return 'Overdue';
    if (diffDays <= 3) return 'Due Soon';
    return 'On Track';
  } catch (e) { return 'Unknown'; }
}

function generateStatusCode(statusTask) {
  var codes = { 'Not Started': 'NS', 'In Progress': 'IP', 'Completed': 'CP', 'On Hold': 'OH' };
  return codes[statusTask] || 'NS';
}

function generateSubTaskId(sheet) {
  var lastRow = sheet.getLastRow();
  var year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yy');
  return 'TSK-' + year + '-' + String(lastRow).padStart(4, '0');
}

function updateTaskNotes(rowIndex, notesJsonString) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Tasks');
    if (!sheet) return { success: false, message: 'Tasks sheet not found' };
    var headers = sheet.getDataRange().getValues()[0];
    var notesCol = headers.indexOf('NotesLog');
    if (notesCol === -1) {
      notesCol = headers.length;
      sheet.getRange(1, notesCol + 1).setValue('NotesLog').setFontWeight('bold').setBackground('#1e293b');
    }
    sheet.getRange(rowIndex, notesCol + 1).setValue(notesJsonString);
    return { success: true, message: 'Notes updated' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function reassignTask(rowIndex, newAssignee, reason) {
  try {
    var currentUser = JSON.parse(PropertiesService.getUserProperties().getProperty('currentUser'));
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Tasks');
    if (!sheet) return { success: false, message: 'Tasks sheet not found' };
    var headers = sheet.getDataRange().getValues()[0];
    var delegationCol = headers.indexOf('DelegationTo');
    var notesCol = headers.indexOf('NotesLog');
    var lastUpdateCol = headers.indexOf('LastUpdate');
    var lastUpdatedByCol = headers.indexOf('LastUpdatedBy');
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    
    if (delegationCol !== -1) sheet.getRange(rowIndex, delegationCol + 1).setValue(newAssignee);
    if (lastUpdateCol !== -1) sheet.getRange(rowIndex, lastUpdateCol + 1).setValue(today);
    if (lastUpdatedByCol !== -1) sheet.getRange(rowIndex, lastUpdatedByCol + 1).setValue(currentUser.name);
    
    if (notesCol !== -1) {
      var existingNotes = sheet.getRange(rowIndex, notesCol + 1).getValue();
      var notes = [];
      try { notes = JSON.parse(existingNotes); } catch(e) {}
      notes.push({ id: Date.now(), timestamp: new Date().toISOString(), author: currentUser.name, text: 'Task reassigned to ' + newAssignee + (reason ? '. Reason: ' + reason : '') });
      sheet.getRange(rowIndex, notesCol + 1).setValue(JSON.stringify(notes));
    }
    return { success: true, message: 'Task reassigned to ' + newAssignee };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ============================================
// HELPERS
// ============================================

function generatePassword(length) {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$';
  var password = '';
  for (var i = 0; i < length; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return password;
}

// ============================================
// SETUP: TAMBAH KOLOM YANG HILANG DI SHEET OKR
// (Jalankan fungsi ini 1 kali saja dari toolbar Apps Script)
// ============================================
function addMissingOKRColumns() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('OKR');
  if (!sheet) {
    Logger.log('Sheet OKR tidak ditemukan!');
    return;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var requiredHeaders = ['Position', 'CreatorRole', 'ManagerName', 'ManagerTitle', 'DirectorName', 'DirectorTitle', 'ReviewNotes', 'ApprovedDate'];
  var addedCount = 0;
  
  requiredHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      var nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(header).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
      Logger.log('✅ Kolom "' + header + '" berhasil ditambahkan di kolom ke-' + nextCol);
      addedCount++;
    } else {
      Logger.log('⏭️ Kolom "' + header + '" sudah ada, dilewati.');
    }
  });
  
  Logger.log('--- SELESAI --- Total ' + addedCount + ' kolom baru ditambahkan.');
}

ini code GS saya 
