// ============================================================
//  PRISM — Print Roll Inventory & Substrate Management
//  BACKEND MODULE (PRISMCode.gs)  v2.0
//  Standalone WebApp — PRISM Database spreadsheet
// ============================================================

// ============================================================
//  SHEET NAMES
// ============================================================
const PRISM_SHEETS = {
  JOB_ORDERS: 'JobOrders',
  LFP_MATERIALS: 'LFP_Materials',
  LFP_ROLLS: 'LFP_Rolls',
  LFP_USAGE: 'LFP_Usage',
  PLOTTING_LOG: 'Plotting_Log',
  ROLES: 'Role and Permission',
  AUDIT: 'PRISM_AuditLog',
  SETTINGS: 'PRISM_Settings'
};

// ============================================================
//  COLUMN MAPS  (0-indexed)
// ============================================================

// JobOrders  A-M
const JO_COL = {
  JO_NUMBER: 0, CUSTOMER: 1, JOB_DESCRIPTION: 2, CATEGORY: 3,
  WIDTH: 4, HEIGHT: 5, QUANTITY: 6, UNIT: 7,
  PLOTTING_LINK: 8, STATUS: 9, ROLL_ID: 10, CREATED_BY: 11, DATE_CREATED: 12
};

// LFP_Materials  A-F
const MAT_COL = {
  MATERIAL_CODE: 0, MATERIAL_NAME: 1, WIDTH: 2,
  STANDARD_LENGTH: 3, SUPPLIER: 4, COST_PER_ROLL: 5
};

// LFP_Rolls  A-I
const ROLL_COL = {
  ROLL_ID: 0, MATERIAL_CODE: 1, WIDTH: 2, ORIGINAL_LENGTH: 3,
  REMAINING_LENGTH: 4, STATUS: 5, DATE_RECEIVED: 6, DATE_OPENED: 7, OPENED_BY: 8
};

// LFP_Usage  A-H
const USAGE_COL = {
  USAGE_ID: 0, JO_NUMBER: 1, ROLL_ID: 2, WIDTH_USED: 3,
  LENGTH_USED: 4, OPERATOR: 5, PLOTTING_LINK: 6, DATE_USED: 7
};

// Plotting_Log  A-F
const PLOT_COL = {
  PLOT_ID:     0,
  TYPE:        1,
  ROLL_ID:     2,
  JO_NUMBERS:  3,
  STATUS:      4,
  START_FT:    5,
  END_FT:      6,
  LENGTH_FT:   7,
  IS_VOID:     8,
  IS_REPRINT:  9,
  PNG_URL:     10,
  OPERATOR:    11,
  DATE_PLOTTED:12,
  REMARKS:     13
};

const PLOT_TYPE = {
  PLOT:       'PLOT',
  DAMAGE:     'DAMAGE',
  ALLOWANCE: 'ALLOWANCE',
  TEST_PRINT: 'TEST_PRINT',
  REPRINT:    'REPRINT'
};
 
const PLOT_STATUS = {
  PLANNED:  'PLANNED',
  PRINTING: 'PRINTING',
  PRINTED:  'PRINTED',
  VOIDED:   'VOIDED',
  DAMAGED:  'DAMAGED'
};

// ============================================================
//  STATUS CONSTANTS
// ============================================================
const JO_STATUS = {
  FOR_PLOTTING: 'FOR_PLOTTING',
  READY_TO_PRINT: 'READY_TO_PRINT',
  PRINTING: 'PRINTING',
  COMPLETED: 'COMPLETED'
};
const ROLL_STATUS = { UNOPENED: 'UNOPENED', OPEN: 'OPEN', CONSUMED: 'CONSUMED' };
const TEST_PRINT_MARKER = 'TEST-PRINT';
const PRISM_PLOTTING_DRIVE_FOLDER_ID = '1IFPphBZ3IjcbkBTMSazeqbge_ZJN4S12';

// ============================================================
//  UNIT CONVERSION UTILITY
// ============================================================
function prism_toFt_(value, unit) {
  const v = parseFloat(value) || 0;
  if (v === 0) return 0;
  switch ((unit || 'ft').toString().trim().toLowerCase()) {
    case 'ft': return Math.round(v * 10000) / 10000;
    case 'in': return Math.round((v / 12) * 10000) / 10000;
    case 'cm': return Math.round((v / 30.48) * 10000) / 10000;
    case 'm': return Math.round((v * 3.28084) * 10000) / 10000;
    default: return Math.round(v * 10000) / 10000;
  }
}

// ============================================================
//  WEB APP ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index_PRISM')
    .evaluate()
    .setTitle('PRISM — Roll Inventory')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
//  BOOTSTRAP — creates all sheets if missing
// ============================================================
function prism_bootstrap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = [
    { name: PRISM_SHEETS.JOB_ORDERS, headers: ['JO_Number', 'Customer', 'JobDescription', 'Category', 'Width', 'Height', 'Quantity', 'Unit', 'PlottingLink', 'Status', 'RollID', 'CreatedBy', 'DateCreated'] },
    { name: PRISM_SHEETS.LFP_MATERIALS, headers: ['MaterialCode', 'MaterialName', 'Width', 'StandardLength', 'Supplier', 'CostPerRoll'] },
    { name: PRISM_SHEETS.LFP_ROLLS, headers: ['RollID', 'MaterialCode', 'Width', 'OriginalLength', 'RemainingLength', 'Status', 'DateReceived', 'DateOpened', 'OpenedBy'] },
    { name: PRISM_SHEETS.LFP_USAGE, headers: ['UsageID', 'JO_Number', 'RollID', 'WidthUsed', 'LengthUsed', 'Operator', 'PlottingLink', 'DateUsed'] },
    { name: PRISM_SHEETS.PLOTTING_LOG, headers: ['PlotID','Type','RollID','JONumbers','Status','StartFt','EndFt','LengthFt','IsVoid','IsReprint','PngUrl','Operator','DatePlotted','Remarks'] },{ name: PRISM_SHEETS.AUDIT, headers: ['DateTime', 'Action', 'User', 'Role', 'PayloadJSON'] },
    { name: PRISM_SHEETS.SETTINGS, headers: ['SettingKey', 'SettingValue'] }
  ];
  schemas.forEach(s => {
    if (!ss.getSheetByName(s.name)) {
      const sh = ss.insertSheet(s.name);
      sh.getRange(1, 1, 1, s.headers.length).setValues([s.headers]);
      sh.getRange(1, 1, 1, s.headers.length).setFontWeight('bold').setBackground('#f8fafc');
      sh.setFrozenRows(1);
    }
  });
  // Seed settings
  const set = ss.getSheetByName(PRISM_SHEETS.SETTINGS);
  if (set && set.getLastRow() < 2) {
    set.getRange(2, 1, 2, 2).setValues([['near_empty_threshold_ft', '30'], ['auto_consume_on_zero', 'true']]);
  }
  return { success: true, message: 'PRISM sheets initialized.' };
}

// ============================================================
//  SHEET GETTER — auto-bootstraps if sheet missing
// ============================================================
function prism_sh_(name) {
  let sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) { prism_bootstrap(); sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
  return sh;
}

// ============================================================
//  DATE HELPERS
// ============================================================
function prism_fmtDate_(val) {
  if (!val) return '';
  try {
    const d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return '';
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  } catch (e) { return ''; }
}
function prism_fmtShort_(val) {
  if (!val) return '';
  try {
    const d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return '';
    const mo = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return mo[d.getMonth()] + ' ' + String(d.getDate()).padStart(2, '0') + ', ' + d.getFullYear();
  } catch (e) { return ''; }
}

function prism_writePlotRow_(plotSh, {
  plotId, type, rollId, joNumbers = [], status,
  startFt = 0, endFt = 0, lengthFt = 0,
  isVoid = false, isReprint = false,
  pngUrl = '', operator, date, remarks = {}
}) {
  plotSh.getRange(plotSh.getLastRow() + 1, 1, 1, 14).setValues([[
    plotId,
    type,
    rollId,
    joNumbers.join(', '),
    status,
    startFt,
    endFt,
    lengthFt,
    isVoid,
    isReprint,
    pngUrl,
    operator,
    date,
    JSON.stringify(remarks)
  ]]);
}


function prism_readPlotRows_(plotSh) {
  const lr = plotSh.getLastRow();
  if (lr < 2) return [];
  return plotSh.getRange(2, 1, lr - 1, 14).getValues()
    .filter(r => String(r[PLOT_COL.PLOT_ID] || '').trim())
    .map((r, i) => {
      let remarks = {};
      try { remarks = JSON.parse(String(r[PLOT_COL.REMARKS] || '{}')); } catch(e) {}
      return {
        _rowIdx:    i + 2,
        plotId:     String(r[PLOT_COL.PLOT_ID]     || '').trim(),
        type:       String(r[PLOT_COL.TYPE]         || '').trim(),
        rollId:     String(r[PLOT_COL.ROLL_ID]      || '').trim(),
        joNumbers:  String(r[PLOT_COL.JO_NUMBERS]   || '').split(',').map(x=>x.trim()).filter(Boolean),
        status:     String(r[PLOT_COL.STATUS]        || '').trim(),
        startFt:    parseFloat(r[PLOT_COL.START_FT]) || 0,joNumbers:  String(r[PLOT_COL.JO_NUMBERS]   || '').split(',').map(x=>x.trim()).filter(function(x) {
            return x.length > 0 && x.indexOf(' ') === -1 && x.indexOf(':') === -1;
          }),
        endFt:      parseFloat(r[PLOT_COL.END_FT])   || 0,
        lengthFt:   parseFloat(r[PLOT_COL.LENGTH_FT])|| 0,
        isVoid:     r[PLOT_COL.IS_VOID] === true || String(r[PLOT_COL.IS_VOID]).toLowerCase() === 'true',
        isReprint:  r[PLOT_COL.IS_REPRINT] === true || String(r[PLOT_COL.IS_REPRINT]).toLowerCase() === 'true',
        pngUrl:     String(r[PLOT_COL.PNG_URL]       || '').trim(),
        operator:   String(r[PLOT_COL.OPERATOR]      || '').trim(),
        date:       r[PLOT_COL.DATE_PLOTTED],
        remarks
      };
    });
}
// ============================================================
//  ROLE & PERMISSION
// ============================================================
function prism_getUserInfo_() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase();
    const sh = prism_sh_(PRISM_SHEETS.ROLES);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const role = String(data[i][0] || '').trim();
      const emails = String(data[i][1] || '').replace(/"/g, '').toLowerCase().split(',').map(e => e.trim()).filter(Boolean);
      const abilities = String(data[i][2] || '').replace(/"/g, '').toLowerCase().split(',').map(a => a.trim()).filter(Boolean);
      if (emails.includes(email)) return { email, role, abilities };
    }
    return { email, role: 'No Role', abilities: [] };
  } catch (e) { return { email: Session.getActiveUser().getEmail(), role: 'No Role', abilities: [] }; }
}
function prism_getUserInfoPublic() { return prism_getUserInfo_(); }
function prism_isAdmin_(r) { return r.toLowerCase().includes('admin'); }
function prism_isATL_(r) { return r.toLowerCase().includes('admin team leader') || r.toLowerCase().includes('admin tl'); }
function prism_isSTL_(r) { return r.toLowerCase().includes('senior team leader') || r.toLowerCase().includes('stl'); }
function prism_isOperator_(r) { return r.toLowerCase().includes('digital operator') || r.toLowerCase().includes('operator'); }

// ============================================================
//  AUDIT LOG
// ============================================================
function prism_audit_(action, payload) {
  try {
    const sh = prism_sh_(PRISM_SHEETS.AUDIT);
    const user = prism_getUserInfo_();
    sh.insertRowBefore(2);
    sh.getRange(2, 1, 1, 5).setValues([[new Date(), action, user.email, user.role, JSON.stringify(payload)]]);
  } catch (e) { Logger.log('audit error: ' + e.message); }
}

// ============================================================
//  PATCHED SECTION — Add/Replace in PRISMCode.gs
//
//  PRISM_Settings sheet structure (2 columns):
//    Col A = key   (e.g. "near_empty_threshold_ft")
//    Col B = value (e.g. "30")
//
//  Keys used:
//    near_empty_threshold_ft
//    low_stock_threshold_ft
//    auto_consume_on_zero
// ============================================================

// ── Private helper: read all settings as a plain object ──────
function prism_getSettings_() {
  try {
    const sh = prism_sh_(PRISM_SHEETS.SETTINGS);
    const lr = sh.getLastRow();
    const out = {};
    if (lr < 1) return out;

    const data = sh.getRange(1, 1, lr, 2).getValues();
    data.forEach(function (row) {
      const key = String(row[0]).trim();
      const val = String(row[1]).trim();
      if (key) out[key] = val;
    });
    return out;
  } catch (e) {
    return {};
  }
}

// ── Private helper: upsert a single key in PRISM_Settings ──────
function prism_setSetting_(key, value) {
  const sh = prism_sh_(PRISM_SHEETS.SETTINGS);
  const lr = sh.getLastRow();

  if (lr >= 1) {
    const data = sh.getRange(1, 1, lr, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) {
        sh.getRange(i + 1, 2).setValue(String(value)); // update existing row
        return;
      }
    }
  }
  // Key not found — append new row
  sh.getRange(lr + 1, 1, 1, 2).setValues([[key, String(value)]]);
}


// ── Public: called by PRISMSettings.html cfg_save() ──────────
function prism_updateSettings(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role))
      return { success: false, message: 'Admin access required.' };

    const nearEmpty = parseFloat(payload.nearEmptyThreshold);
    const lowStock = parseFloat(payload.lowStockThreshold);
    const autoConsume = payload.autoConsumeOnZero;

    // Validate
    if (!nearEmpty || nearEmpty < 1)
      return { success: false, message: 'Near Empty Threshold must be ≥ 1 ft.' };
    if (!lowStock || lowStock < 1)
      return { success: false, message: 'Low Stock Threshold must be ≥ 1 ft.' };
    if (lowStock <= nearEmpty)
      return { success: false, message: 'Low Stock must be greater than Near Empty.' };

    // Write all three keys
    prism_setSetting_('near_empty_threshold_ft', nearEmpty);
    prism_setSetting_('low_stock_threshold_ft', lowStock);
    prism_setSetting_('auto_consume_on_zero', autoConsume === true || autoConsume === 'true' ? 'true' : 'false');

    prism_audit_('PRISM_UPDATE_SETTINGS', {
      nearEmptyThreshold: nearEmpty,
      lowStockThreshold: lowStock,
      autoConsumeOnZero: autoConsume,
      by: user.email
    });

    return {
      success: true,
      message: `Settings saved. Near Empty: ${nearEmpty} ft, Low Stock: ${lowStock} ft.`
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}


// ============================================================
//  ROLL ID GENERATOR  →  MOD001-R01 format, auto-incremented
// ============================================================
function prism_generateRollIds_(materialCode, qty, existingRows) {
  // Find highest R number already used for this exact materialCode
  let maxRoll = 0;
  existingRows.forEach(r => {
    const id = String(r[ROLL_COL.ROLL_ID] || '');
    const prefix = materialCode + '-R';
    if (id.startsWith(prefix)) {
      const m = id.match(/-R(\d+)$/);
      if (m) { const n = parseInt(m[1]); if (n > maxRoll) maxRoll = n; }
    }
  });
  const ids = [];
  for (let i = 1; i <= qty; i++) {
    ids.push(materialCode + '-R' + String(maxRoll + i).padStart(2, '0'));
  }
  return ids;
}

// ============================================================
//  COMBINED INIT DATA
// ============================================================
function prism_getInitData() {
  try {
    return {
      success: true,
      user: prism_getUserInfo_(),
      rolls: prism_getAllRolls_(),
      materials: prism_getAllMaterials_(),
      settings: prism_getSettings_()
    };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
//  LFP_MATERIALS
// ============================================================
function prism_getAllMaterials_() {
  const sh = prism_sh_(PRISM_SHEETS.LFP_MATERIALS);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, 6).getValues()
    .filter(r => r[MAT_COL.MATERIAL_CODE] && String(r[MAT_COL.MATERIAL_CODE]).trim())
    .map((r, i) => ({
      rowIndex: i + 2,
      materialCode: String(r[MAT_COL.MATERIAL_CODE]).trim(),
      materialName: String(r[MAT_COL.MATERIAL_NAME]).trim(),
      width: parseFloat(r[MAT_COL.WIDTH]) || 0,
      standardLength: parseFloat(r[MAT_COL.STANDARD_LENGTH]) || 0,
      supplier: String(r[MAT_COL.SUPPLIER] || '').trim(),
      costPerRoll: parseFloat(r[MAT_COL.COST_PER_ROLL]) || 0
    }));
}
function prism_getAllMaterialsPublic() {
  try { return { success: true, data: prism_getAllMaterials_() }; }
  catch (e) { return { success: false, message: e.message }; }
}
function prism_addMaterial(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isATL_(user.role))
      return { success: false, message: 'Admin or Admin Team Leader only.' };
    if (!payload.materialCode || !payload.materialName)
      return { success: false, message: 'Material Code and Name are required.' };
    if (!payload.width || payload.width <= 0) return { success: false, message: 'Width required.' };
    if (!payload.standardLength || payload.standardLength <= 0) return { success: false, message: 'Standard Length required.' };
    const existing = prism_getAllMaterials_();
    if (existing.some(m => m.materialCode.toLowerCase() === payload.materialCode.trim().toLowerCase()))
      return { success: false, message: 'Material Code "' + payload.materialCode + '" already exists.' };
    const sh = prism_sh_(PRISM_SHEETS.LFP_MATERIALS);
    sh.getRange(sh.getLastRow() + 1, 1, 1, 6).setValues([[
      payload.materialCode.trim().toUpperCase(), payload.materialName.trim(),
      parseFloat(payload.width), parseFloat(payload.standardLength),
      payload.supplier || '', parseFloat(payload.costPerRoll) || 0
    ]]);
    prism_audit_('PRISM_ADD_MATERIAL', { materialCode: payload.materialCode, by: user.email });
    return { success: true, message: 'Material "' + payload.materialCode + '" added.' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
//  LFP_ROLLS
// ============================================================
function prism_getAllRolls_() {
  const sh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  const matMap = {};
  try { prism_getAllMaterials_().forEach(m => { matMap[m.materialCode] = m.materialName || ''; }); } catch (_) { }
  return sh.getRange(2, 1, lr - 1, 9).getValues()
    .filter(r => r[ROLL_COL.ROLL_ID] && String(r[ROLL_COL.ROLL_ID]).trim())
    .map((r, i) => {
      const matCode = String(r[ROLL_COL.MATERIAL_CODE]).trim();
      return {
        rowIndex: i + 2,
        rollId: String(r[ROLL_COL.ROLL_ID]).trim(),
        materialCode: matCode,
        materialName: matMap[matCode] || '',
        width: parseFloat(r[ROLL_COL.WIDTH]) || 0,
        originalLength: parseFloat(r[ROLL_COL.ORIGINAL_LENGTH]) || 0,
        remainingLength: parseFloat(r[ROLL_COL.REMAINING_LENGTH]) || 0,
        status: String(r[ROLL_COL.STATUS] || 'UNOPENED').trim().toUpperCase(),
        dateReceived: prism_fmtShort_(r[ROLL_COL.DATE_RECEIVED]),
        dateOpened: prism_fmtShort_(r[ROLL_COL.DATE_OPENED]),
        openedBy: String(r[ROLL_COL.OPENED_BY] || '').trim()
      };
    });
}
function prism_getAllRollsPublic() {
  try { return { success: true, rolls: prism_getAllRolls_(), settings: prism_getSettings_() }; }
  catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
//  STOCK IN — Admin Team Leader
// ============================================================
function prism_stockIn(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isATL_(user.role))
      return { success: false, message: 'Admin Team Leader only.' };
    if (!payload.materialCode) return { success: false, message: 'Material Code required.' };
    if (!payload.qty || payload.qty < 1) return { success: false, message: 'Qty must be ≥ 1.' };
    if (!payload.width || payload.width <= 0) return { success: false, message: 'Width required.' };
    if (!payload.originalLength || payload.originalLength <= 0) return { success: false, message: 'Length required.' };

    const sh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const lr = sh.getLastRow();
    const existing = lr >= 2 ? sh.getRange(2, 1, lr - 1, 9).getValues() : [];
    const rollIds = prism_generateRollIds_(payload.materialCode, parseInt(payload.qty), existing);
    const today = new Date();
    const rows = rollIds.map(id => [
      id, payload.materialCode.trim().toUpperCase(),
      parseFloat(payload.width), parseFloat(payload.originalLength), parseFloat(payload.originalLength),
      ROLL_STATUS.UNOPENED, today, '', ''
    ]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
    prism_audit_('PRISM_STOCK_IN', { materialCode: payload.materialCode, qty: payload.qty, rollIds, by: user.email });
    return { success: true, message: rollIds.length + ' roll(s) stocked in: ' + rollIds.join(', '), rollIds };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
//  OPEN ROLL — Senior Team Leader
// ============================================================
function prism_openRoll(rollId) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isSTL_(user.role))
      return { success: false, message: 'Senior Team Leader only.' };
    const sh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const lr = sh.getLastRow();
    if (lr < 2) return { success: false, message: 'No rolls found.' };
    const data = sh.getRange(2, 1, lr - 1, 9).getValues();
    let rowIdx = -1;
    data.forEach((r, i) => { if (String(r[ROLL_COL.ROLL_ID]).trim() === rollId) rowIdx = i + 2; });
    if (rowIdx === -1) return { success: false, message: 'Roll "' + rollId + '" not found.' };
    const row = sh.getRange(rowIdx, 1, 1, 9).getValues()[0];
    const status = String(row[ROLL_COL.STATUS]).trim().toUpperCase();
    if (status === ROLL_STATUS.OPEN) return { success: false, message: 'Roll is already OPEN.' };
    if (status === ROLL_STATUS.CONSUMED) return { success: false, message: 'Roll is CONSUMED.' };
    const today = new Date();
    sh.getRange(rowIdx, ROLL_COL.STATUS + 1).setValue(ROLL_STATUS.OPEN);
    sh.getRange(rowIdx, ROLL_COL.DATE_OPENED + 1).setValue(today);
    sh.getRange(rowIdx, ROLL_COL.OPENED_BY + 1).setValue(user.email);
    prism_audit_('PRISM_OPEN_ROLL', { rollId, by: user.email });
    return { success: true, message: 'Roll ' + rollId + ' is now OPEN.' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
//  PATCHED SECTION — Replace prism_recordUsage() in PRISMCode.gs
//  Change: RollID column now stores comma-separated list
//          (appends new rollId, deduplicates, never overwrites)
// ============================================================

function prism_recordUsage(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role))
      return { success: false, message: 'Digital Operators only.' };

    // ── Validate inputs ──
    if (!payload.rollId)
      return { success: false, message: 'Roll ID required.' };
    if (!payload.joNumber || !payload.joNumber.trim())
      return { success: false, message: 'JO Number required.' };
    if (!payload.lengthUsed || payload.lengthUsed <= 0)
      return { success: false, message: 'Length used must be > 0.' };
    if (!payload.plottingLink)
      return { success: false, message: 'Plotting link required.' };

    // ── Fetch roll row ──
    const rollSh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollLr = rollSh.getLastRow();
    const rollData = rollSh.getRange(2, 1, rollLr - 1, 9).getValues();
    let rollIdx = -1, rollRow = null;
    rollData.forEach((r, i) => {
      if (String(r[ROLL_COL.ROLL_ID]).trim() === payload.rollId) {
        rollIdx = i + 2; rollRow = r;
      }
    });
    if (rollIdx === -1)
      return { success: false, message: `Roll "${payload.rollId}" not found.` };
    if (String(rollRow[ROLL_COL.STATUS]).trim().toUpperCase() !== ROLL_STATUS.OPEN)
      return { success: false, message: 'Roll must be OPEN.' };

    // ── Width validation ──
    const rollWidth = parseFloat(rollRow[ROLL_COL.WIDTH]) || 0;
    const jobWidth = parseFloat(payload.widthUsed) || 0;
    if (jobWidth > 0 && jobWidth > rollWidth)
      return { success: false, message: `Job width (${jobWidth} ft) exceeds roll width (${rollWidth} ft).` };

    // ── Length validation ──
    const remaining = parseFloat(rollRow[ROLL_COL.REMAINING_LENGTH]) || 0;
    const lengthUsed = parseFloat(payload.lengthUsed);
    if (lengthUsed > remaining)
      return { success: false, message: `Length used (${lengthUsed} ft) exceeds remaining (${remaining} ft).` };

    // ── Deduct from roll ──
    const settings = prism_getSettings_();
    const newRemaining = Math.max(0, remaining - lengthUsed);
    const newRollStatus = (newRemaining === 0 && settings.auto_consume_on_zero === 'true')
      ? ROLL_STATUS.CONSUMED : ROLL_STATUS.OPEN;

    rollSh.getRange(rollIdx, ROLL_COL.REMAINING_LENGTH + 1).setValue(newRemaining);
    rollSh.getRange(rollIdx, ROLL_COL.STATUS + 1).setValue(newRollStatus);

    // ── Write usage entry ──
    const today = new Date();
    const usageSh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    usageSh.getRange(usageSh.getLastRow() + 1, 1, 1, 8).setValues([[
      'USE-' + today.getTime(),
      payload.joNumber.trim().toUpperCase(),
      payload.rollId,
      jobWidth,
      lengthUsed,
      user.email,
      payload.plottingLink,
      today
    ]]);

    // ── Update JobOrders: append RollID (multi-roll support) ──
    const joSh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const joLr = joSh.getLastRow();
    if (joLr >= 2) {
      const joData = joSh.getRange(2, 1, joLr - 1, 13).getValues();
      joData.forEach((r, i) => {
        if (String(r[JO_COL.JO_NUMBER]).trim().toUpperCase() ===
          payload.joNumber.trim().toUpperCase()) {

          // ── Multi-roll: append new rollId to existing list ──
          const existingRolls = String(r[JO_COL.ROLL_ID] || '')
            .split(',')
            .map(x => x.trim())
            .filter(Boolean);

          if (!existingRolls.includes(payload.rollId)) {
            existingRolls.push(payload.rollId);
          }

          joSh.getRange(i + 2, JO_COL.ROLL_ID + 1)
            .setValue(existingRolls.join(', '));

          // ── Advance JO status to READY_TO_PRINT ──
          joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.READY_TO_PRINT);

          // ── Keep PlottingLink updated (latest wins) ──
          if (payload.plottingLink) {
            joSh.getRange(i + 2, JO_COL.PLOTTING_LINK + 1).setValue(payload.plottingLink);
          }
        }
      });
    }

    prism_audit_('PRISM_RECORD_USAGE', {
      rollId: payload.rollId,
      joNumber: payload.joNumber,
      lengthUsed,
      newRemaining,
      newRollStatus,
      by: user.email
    });

    return {
      success: true,
      message: `Usage recorded. ${payload.rollId}: ${newRemaining} ft remaining.`
        + (newRollStatus === ROLL_STATUS.CONSUMED ? ' Roll CONSUMED.' : ''),
      newRemaining,
      newRollStatus
    };

  } catch (e) { return { success: false, message: e.message }; }
}


function prism_recordTestPrint(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role))
      return { success: false, message: 'Digital Operators only.' };

    if (!payload.rollId)
      return { success: false, message: 'Roll ID required.' };
    if (!payload.lengthUsed || payload.lengthUsed <= 0)
      return { success: false, message: 'Length used must be > 0.' };

    // ── Fetch roll row ──
    const rollSh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollLr = rollSh.getLastRow();
    const rollData = rollSh.getRange(2, 1, rollLr - 1, 9).getValues();
    let rollIdx = -1, rollRow = null;
    rollData.forEach((r, i) => {
      if (String(r[ROLL_COL.ROLL_ID]).trim() === payload.rollId) {
        rollIdx = i + 2; rollRow = r;
      }
    });
    if (rollIdx === -1)
      return { success: false, message: `Roll "${payload.rollId}" not found.` };
    if (String(rollRow[ROLL_COL.STATUS]).trim().toUpperCase() !== ROLL_STATUS.OPEN)
      return { success: false, message: 'Roll must be OPEN.' };

    // ── Length validation ──
    const remaining = parseFloat(rollRow[ROLL_COL.REMAINING_LENGTH]) || 0;
    const lengthUsed = parseFloat(payload.lengthUsed);
    if (lengthUsed > remaining)
      return { success: false, message: `Length used (${lengthUsed} ft) exceeds remaining (${remaining} ft).` };

    // ── Deduct from roll ──
    const settings = prism_getSettings_();
    const newRemaining = Math.max(0, remaining - lengthUsed);
    const newRollStatus = (newRemaining === 0 && settings.auto_consume_on_zero === 'true')
      ? ROLL_STATUS.CONSUMED : ROLL_STATUS.OPEN;
    const rollWidth = parseFloat(rollRow[ROLL_COL.WIDTH]) || 0;

    rollSh.getRange(rollIdx, ROLL_COL.REMAINING_LENGTH + 1).setValue(newRemaining);
    rollSh.getRange(rollIdx, ROLL_COL.STATUS + 1).setValue(newRollStatus);

    const today = new Date();

    // ── Write to LFP_Usage with TEST-PRINT marker ──
    const usageSh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    var printType = (payload.type === 'ALLOWANCE') ? 'ALLOWANCE' : 'TEST_PRINT';
    usageSh.getRange(usageSh.getLastRow() + 1, 1, 1, 8).setValues([[
      (printType === 'ALLOWANCE' ? 'ALLOW-' : 'TEST-') + today.getTime(),
      TEST_PRINT_MARKER,
      payload.rollId,
      rollWidth,
      lengthUsed,
      user.email,
      payload.notes || '',
      today
    ]]);

    // ── Write to Plotting_Log so test print appears on the roll map ──
    const plotSh2   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows2 = prism_readPlotRows_(plotSh2);
    let maxEnd2 = 0;
    plotRows2.forEach(r => {
      if (r.rollId === payload.rollId && !r.isVoid && r.endFt > maxEnd2) maxEnd2 = r.endFt;
    });
    prism_writePlotRow_(plotSh2, {
      plotId:    'TEST-PLT-' + today.getTime(),
      type:      (payload.type === 'ALLOWANCE') ? PLOT_TYPE.ALLOWANCE : PLOT_TYPE.TEST_PRINT,
      rollId:    payload.rollId,
      joNumbers: [],
      status:    PLOT_STATUS.PRINTED,
      startFt:   maxEnd2,
      endFt:     maxEnd2 + lengthUsed,
      lengthFt:  lengthUsed,
      isVoid:    false,
      isReprint: false,
      pngUrl:    '',
      operator:  user.email,
      date:      today,
      remarks:   { notes: payload.notes || '', rollWidth: rollWidth }
    });

    prism_audit_('PRISM_TEST_PRINT', {
      rollId: payload.rollId,
      lengthUsed,
      newRemaining,
      newRollStatus,
      notes: payload.notes || '',
      by: user.email
    });

    return {
      success: true,
      message: `Test print recorded. ${payload.rollId}: ${newRemaining} ft remaining.`
        + (newRollStatus === ROLL_STATUS.CONSUMED ? ' Roll CONSUMED.' : ''),
      newRemaining,
      newRollStatus
    };

  } catch (e) { return { success: false, message: e.message }; }
}

function prism_getTestPrintLog() {
  try {
    const sh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    const lr = sh.getLastRow();
    if (lr < 2) return { success: true, data: [] };

    // Notes are stored in the PLOTTING_LINK column for test prints
    const data = sh.getRange(2, 1, lr - 1, 8).getValues()
      .filter(r => String(r[USAGE_COL.JO_NUMBER]).trim() === TEST_PRINT_MARKER)
      .map(r => {
        const d = r[USAGE_COL.DATE_USED];
        return {
          rollId: String(r[USAGE_COL.ROLL_ID] || '').trim(),
          lengthUsed: parseFloat(r[USAGE_COL.LENGTH_USED]) || 0,
          operator: String(r[USAGE_COL.OPERATOR] || '').trim(),
          dateUsed: prism_fmtShort_(d),
          notes: String(r[USAGE_COL.PLOTTING_LINK] || '').trim(),
          _dateRaw: d ? new Date(d).getTime() : 0
        };
      })
      .sort((a, b) => b._dateRaw - a._dateRaw);

    return { success: true, data: data };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
//  DAMAGE DECLARATION
// ============================================================
function prism_declareDamage(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role))
      return { success: false, message: 'Digital Operators only.' };

    if (!payload.rollId)
      return { success: false, message: 'Roll ID required.' };
    if (!payload.lengthUsed || payload.lengthUsed <= 0)
      return { success: false, message: 'Damage length must be > 0.' };

    // ── Fetch roll row ──
    const rollSh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollLr = rollSh.getLastRow();
    const rollData = rollSh.getRange(2, 1, rollLr - 1, 9).getValues();
    let rollIdx = -1, rollRow = null;
    rollData.forEach((r, i) => {
      if (String(r[ROLL_COL.ROLL_ID]).trim() === payload.rollId) {
        rollIdx = i + 2; rollRow = r;
      }
    });
    if (rollIdx === -1)
      return { success: false, message: `Roll "${payload.rollId}" not found.` };
    if (String(rollRow[ROLL_COL.STATUS]).trim().toUpperCase() !== ROLL_STATUS.OPEN)
      return { success: false, message: 'Roll must be OPEN.' };

    const remaining = parseFloat(rollRow[ROLL_COL.REMAINING_LENGTH]) || 0;
    const lengthUsed = parseFloat(payload.lengthUsed);

    const rollWidth = parseFloat(rollRow[ROLL_COL.WIDTH]) || 0;
    const originalLength = parseFloat(rollRow[ROLL_COL.ORIGINAL_LENGTH]) || 0;
    const today = new Date();

    // ── Calculate Auto-Refund from recent Usage ──
    let refundLength = 0;
    const targets = (payload.joNumbers || []).map(jn => String(jn).trim().toUpperCase());
    const usageSh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    const usageLr = usageSh.getLastRow();
    let latestUsageByJo = {};

    if (targets.length > 0 && usageLr >= 2) {
      const usageData = usageSh.getRange(2, 1, usageLr - 1, 8).getValues();
      usageData.forEach(r => {
        const uJo = String(r[USAGE_COL.JO_NUMBER]).trim().toUpperCase();
        const uRoll = String(r[USAGE_COL.ROLL_ID]).trim();
        // Since rows are appended chronologically, the last one we see is the latest
        if (uRoll === payload.rollId && targets.includes(uJo)) {
          latestUsageByJo[uJo] = parseFloat(r[USAGE_COL.LENGTH_USED]) || 0;
        }
      });
      Object.values(latestUsageByJo).forEach(len => { refundLength += len; });
    }

    if (lengthUsed > remaining + refundLength) {
      return { success: false, message: `Damage length (${lengthUsed} ft) exceeds available remainder (${remaining + refundLength} ft).` };
    }

    const settings = prism_getSettings_();
    // The refundLength brings back what was previously deducted for those JOs.
    // We only subtract the declared damage amount from whatever space those JOs occupied.
    // Net effect: roll gains back (refundLength - lengthUsed) feet.
    const newRemaining = Math.max(0, remaining + refundLength - lengthUsed);
    const newRollStatus = (newRemaining === 0 && settings.auto_consume_on_zero === 'true')
      ? ROLL_STATUS.CONSUMED : ROLL_STATUS.OPEN;

    rollSh.getRange(rollIdx, ROLL_COL.REMAINING_LENGTH + 1).setValue(newRemaining);
    // Always write the status — if the roll was previously CONSUMED due to the plot, restore it to OPEN
    rollSh.getRange(rollIdx, ROLL_COL.STATUS + 1).setValue(ROLL_STATUS.OPEN);
    // Then only override to CONSUMED if legitimately zero
    if (newRemaining === 0 && settings.auto_consume_on_zero === 'true') {
      rollSh.getRange(rollIdx, ROLL_COL.STATUS + 1).setValue(ROLL_STATUS.CONSUMED);
    }

    // ── Write Auto-Refund Negatives to LFP_Usage ──
    const refundRows = [];
    Object.keys(latestUsageByJo).forEach(uJo => {
      const len = latestUsageByJo[uJo];
      if (len > 0) {
        refundRows.push([
          'REF-' + today.getTime() + '-' + uJo,
          uJo,
          payload.rollId,
          rollWidth,
          -len, // Negative usage to offset the failed plot
          user.email,
          'Auto-Refund from Damage Declaration',
          today
        ]);
      }
    });

    if (refundRows.length > 0) {
      usageSh.getRange(usageLr + 1, 1, refundRows.length, 8).setValues(refundRows);
    }

    // ── Write Damage to LFP_Usage ──
    usageSh.getRange(usageSh.getLastRow() + 1, 1, 1, 8).setValues([[
      'DAM-' + today.getTime(),
      'DAMAGE',                    // JO_Number = 'DAMAGE'
      payload.rollId,
      rollWidth,                   // widthUsed = full roll width
      lengthUsed,
      user.email,
      payload.remarks || '',
      today
    ]]);

// ── Find and void previous PLANNED plot in Plotting_Log ──
    const plotSh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
    for (let i = plotRows.length - 1; i >= 0; i--) {
      const row = plotRows[i];
      if (row.rollId === payload.rollId &&
          !row.isVoid &&
          row.type === PLOT_TYPE.PLOT &&
          row.joNumbers.some(jo => targets.includes(jo.toUpperCase()))) {
        plotSh.getRange(row._rowIdx, PLOT_COL.IS_VOID + 1).setValue(true);
        plotSh.getRange(row._rowIdx, PLOT_COL.STATUS  + 1).setValue(PLOT_STATUS.VOIDED);
        break;
      }
    }
 
    // ── Write DAMAGE block to Plotting_Log ──
    const startAtFt = Math.max(0, originalLength - remaining);
    const endAtFt   = startAtFt + lengthUsed;
 
    prism_writePlotRow_(plotSh, {
      plotId:    'PLT-DAM-' + today.getTime(),
      type:      PLOT_TYPE.DAMAGE,
      rollId:    payload.rollId,
      joNumbers: payload.joNumbers || [],
      status:    PLOT_STATUS.PRINTED,
      startFt:   startAtFt,
      endFt:     endAtFt,
      lengthFt:  lengthUsed,
      isVoid:    false,
      isReprint: false,
      pngUrl:    '',
      operator:  user.email,
      date:      today,
      remarks: {
        rollWidth:      rollWidth,
        originalLength: originalLength,
        damageReason:   payload.remarks || ''
      }
    });

    // ── Revert affected JOs back to FOR_PLOTTING ──
    let affectedCount = 0;
    if (payload.joNumbers && payload.joNumbers.length > 0) {
      const joSh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
      const joLr = joSh.getLastRow();
      if (joLr >= 2) {
        const joData = joSh.getRange(2, 1, joLr - 1, 13).getValues();
        const targets = payload.joNumbers.map(jn => String(jn).trim().toUpperCase());
        joData.forEach((r, i) => {
          const joNum = String(r[JO_COL.JO_NUMBER]).trim().toUpperCase();
          if (targets.includes(joNum)) {
            joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.FOR_PLOTTING);
            affectedCount++;
          }
        });
      }
    }

    prism_audit_('PRISM_DECLARE_DAMAGE', {
      rollId: payload.rollId,
      lengthUsed,
      newRemaining,
      newRollStatus,
      affectedJOs: payload.joNumbers || [],
      remarks: payload.remarks || '',
      by: user.email
    });

    return {
      success: true,
      message: `Declared ${lengthUsed}ft damage. ` + (affectedCount > 0 ? `${affectedCount} JO(s) reverted to FOR_PLOTTING.` : ''),
      newRemaining,
      newRollStatus
    };

  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
//  DECLARE DAMAGE FROM PRINT QUEUE (by plotId)
//  Voids the specific PRINTED plot entry, logs a damage block,
//  reverts JOs to FOR_PLOTTING, writes a REPRINT entry.
// ============================================================
function prism_declareDamageForPlot(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role))
      return { success: false, message: 'Access denied.' };
 
    const plotSh   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
    const target   = plotRows.find(r => r.plotId === payload.plotId);
 
    if (!target || target.status !== PLOT_STATUS.PRINTING || target.isVoid)
      return { success: false, message: 'PRINTING entry not found or already voided.' };
 
    const rollId       = target.rollId;
    const joNumbers    = target.joNumbers;
    const today        = new Date();
    const printLength  = target.lengthFt;
    const printStart   = target.startFt;
    const damageLength = parseFloat(payload.lengthUsed) || 0;
 
    if (damageLength <= 0)
      return { success: false, message: 'Damage length must be > 0.' };
 
    // 1. Void the PRINTING entry
    plotSh.getRange(target._rowIdx, PLOT_COL.STATUS  + 1).setValue(PLOT_STATUS.VOIDED);
    plotSh.getRange(target._rowIdx, PLOT_COL.IS_VOID + 1).setValue(true);
 
    // 2. Write DAMAGE block
    const damageId  = 'DMG-' + today.getTime();
    prism_writePlotRow_(plotSh, {
      plotId:    damageId,
      type:      PLOT_TYPE.DAMAGE,
      rollId,
      joNumbers,
      status:    PLOT_STATUS.PRINTED,
      startFt:   printStart,
      endFt:     printStart + damageLength,
      lengthFt:  damageLength,
      isVoid:    false,
      isReprint: false,
      pngUrl:    '',
      operator:  user.email,
      date:      today,
      remarks: { damageReason: payload.remarks || '' }
    });
 
    // 3. Deduct damage length from roll
    const rollSh   = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollData = rollSh.getRange(2, 1, rollSh.getLastRow() - 1, 9).getValues();
    rollData.forEach((r, i) => {
      if (String(r[ROLL_COL.ROLL_ID]).trim() !== rollId) return;
      const cur  = parseFloat(r[ROLL_COL.REMAINING_LENGTH]) || 0;
      const next = Math.max(0, cur - damageLength);
      rollSh.getRange(i + 2, ROLL_COL.REMAINING_LENGTH + 1).setValue(next);
      rollSh.getRange(i + 2, ROLL_COL.STATUS + 1).setValue(next <= 0 ? ROLL_STATUS.CONSUMED : ROLL_STATUS.OPEN);
    });
 
    // 4. Write REPRINT entry (continues after damage block)
    const reprintId    = 'PLT-RPT-' + today.getTime();
    const reprintCount = (parseInt(target.remarks.reprintCount) || 0) + 1;
    prism_writePlotRow_(plotSh, {
      plotId:    reprintId,
      type:      PLOT_TYPE.REPRINT,
      rollId,
      joNumbers,
      status:    PLOT_STATUS.PLANNED,
      startFt:   printStart + damageLength,
      endFt:     printStart + damageLength + printLength,
      lengthFt:  printLength,
      isVoid:    false,
      isReprint: true,
      pngUrl:    target.pngUrl || '',
      operator:  user.email,
      date:      today,
      remarks: Object.assign({}, target.remarks, {
        reprintOf:       payload.plotId,
        reprintCount:    reprintCount,
        reprintAnchorFt: printStart + damageLength
      })
    });
 
    // 5. Revert JOs to READY_TO_PRINT
    if (joNumbers.length) {
      const joSh   = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
      const joData = joSh.getRange(2, 1, joSh.getLastRow() - 1, 13).getValues();
      joData.forEach((r, i) => {
        if (joNumbers.includes(String(r[JO_COL.JO_NUMBER]).trim().toUpperCase()))
          joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.READY_TO_PRINT);
      });
    }
 
    prism_audit_('PRISM_DECLARE_DAMAGE_FOR_PLOT', {
      plotId: payload.plotId, rollId, joNumbers, damageLength, by: user.email
    });
    return { success: true, message: `Damage declared. ${damageLength}ft recorded. Reprint added to queue.`, reprintId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
// ============================================================
//  JOB ORDERS
// ============================================================
function prism_getAllJobOrders_() {
  const latestPlotAssets = prism_getLatestPlotAssetsByJO_();
  const sh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, 13).getValues()
    .filter(r => r[JO_COL.JO_NUMBER] && String(r[JO_COL.JO_NUMBER]).trim())
    .map((r, i) => {
      const joNumber = String(r[JO_COL.JO_NUMBER]).trim();
      const plotAsset = latestPlotAssets[String(joNumber || '').toUpperCase()] || {};
      return {
        rowIndex: i + 2,
        joNumber: joNumber,
        customer: String(r[JO_COL.CUSTOMER] || '').trim(),
        jobDescription: String(r[JO_COL.JOB_DESCRIPTION] || '').trim(),
        category: String(r[JO_COL.CATEGORY] || '').trim(),
        width: parseFloat(r[JO_COL.WIDTH]) || 0,
        height: parseFloat(r[JO_COL.HEIGHT]) || 0,
        quantity: parseInt(r[JO_COL.QUANTITY]) || 0,
        unit: (String(r[JO_COL.UNIT] || 'ft').trim().toLowerCase()) || 'ft',
        widthFt: prism_toFt_(parseFloat(r[JO_COL.WIDTH]) || 0, String(r[JO_COL.UNIT] || 'ft').trim()),
        heightFt: prism_toFt_(parseFloat(r[JO_COL.HEIGHT]) || 0, String(r[JO_COL.UNIT] || 'ft').trim()),
        plottingLink: String(r[JO_COL.PLOTTING_LINK] || '').trim(),
        plottingImageUrl: String(plotAsset.pngUrl || r[JO_COL.PLOTTING_LINK] || '').trim(),
        plottingFolderUrl: String(plotAsset.folderUrl || '').trim(),
        plotDateMs: parseInt(plotAsset.dateMs) || 0,
        status: String(r[JO_COL.STATUS] || JO_STATUS.FOR_PLOTTING).trim(),
        rollId: String(r[JO_COL.ROLL_ID] || '').trim(),
        createdBy: String(r[JO_COL.CREATED_BY] || '').trim(),
        dateCreated: prism_fmtShort_(r[JO_COL.DATE_CREATED])
      }
    })
    .sort((a, b) => new Date(b.dateCreated) - new Date(a.dateCreated));
}
 
function prism_getLatestPlotAssetsByJO_() {
  const out = {};
  try {
    const plotSh   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
 
    plotRows.forEach(row => {
      if (row.type !== PLOT_TYPE.PLOT || row.isVoid) return;
      const dateMs = row.date ? new Date(row.date).getTime() : 0;
      row.joNumbers.forEach(jo => {
        const key = String(jo||'').trim().toUpperCase();
        if (!key) return;
        if (!out[key] || dateMs >= (out[key].dateMs || 0)) {
          out[key] = {
            dateMs,
            pngUrl:    row.pngUrl,
            folderUrl: row.remarks.folderUrl || '',
            rollId:    row.rollId
          };
        }
      });
    });
  } catch(e) {}
  return out;
}
 
function prism_getJobOrdersPublic() {
  try { return { success: true, data: prism_getAllJobOrders_() }; }
  catch (e) { return { success: false, message: e.message }; }
}

function prism_getUsageSummaryByJO_() {
  const sh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
  const lr = sh.getLastRow();
  const map = {};
  if (lr < 2) return map;

  const rows = sh.getRange(2, 1, lr - 1, 8).getValues();
  rows.forEach(r => {
    const jo = String(r[USAGE_COL.JO_NUMBER] || '').trim().toUpperCase();
    if (!jo) return;

    if (!map[jo]) {
      map[jo] = {
        count: 0,
        totalLength: 0,
        lastDate: null,
        lastDateLabel: '',
        lastRollId: '',
        rollIds: {}
      };
    }

    const rollId = String(r[USAGE_COL.ROLL_ID] || '').trim();
    const len = parseFloat(r[USAGE_COL.LENGTH_USED]) || 0;
    const dtRaw = r[USAGE_COL.DATE_USED];
    const dt = dtRaw ? new Date(dtRaw) : null;
    const validDate = dt && !isNaN(dt.getTime()) ? dt : null;

    map[jo].count += 1;
    map[jo].totalLength += len;
    if (rollId) map[jo].rollIds[rollId] = true;

    if (!map[jo].lastDate || (validDate && validDate.getTime() > map[jo].lastDate.getTime())) {
      map[jo].lastDate = validDate || map[jo].lastDate;
      map[jo].lastDateLabel = validDate ? prism_fmtShort_(validDate) : map[jo].lastDateLabel;
      map[jo].lastRollId = rollId || map[jo].lastRollId;
    }
  });

  return map;
}

function prism_getForPlottingJOs() {
  try {
    const sh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const lr = sh.getLastRow();
    if (lr < 2) return { success: true, data: [] };

    const usageSh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    const usageLr = usageSh.getLastRow();
    const usageMap = {};
    if (usageLr >= 2) {
      usageSh.getRange(2, 1, usageLr - 1, 8).getValues().forEach(u => {
        const jo = String(u[USAGE_COL.JO_NUMBER] || '').trim().toUpperCase();
        if (!jo) return;
        if (!usageMap[jo]) usageMap[jo] = { count: 0, rollIds: [], lastRollId: '', lastDate: '' };
        usageMap[jo].count++;
        const rid = String(u[USAGE_COL.ROLL_ID] || '').trim();
        if (rid && !usageMap[jo].rollIds.includes(rid)) usageMap[jo].rollIds.push(rid);
        usageMap[jo].lastRollId = rid || usageMap[jo].lastRollId;
        const d = u[USAGE_COL.DATE_USED];
        if (d) usageMap[jo].lastDate = prism_fmtShort_(d);
      });
    }

    const rows = sh.getRange(2, 1, lr - 1, 13).getValues();
    const data = [];

    rows.forEach((r, i) => {
      const joNumber = String(r[JO_COL.JO_NUMBER] || '').trim();
      if (!joNumber) return;
      const status = String(r[JO_COL.STATUS] || JO_STATUS.FOR_PLOTTING).trim();
      if (status !== JO_STATUS.FOR_PLOTTING) return;

      const unit = String(r[JO_COL.UNIT] || 'ft').trim();
      const usage = usageMap[joNumber.toUpperCase()] || {};
      const widthRaw  = parseFloat(r[JO_COL.WIDTH])  || 0;
      const heightRaw = parseFloat(r[JO_COL.HEIGHT]) || 0;

      data.push({
        rowIndex:       i + 2,
        joNumber:       joNumber,
        customer:       String(r[JO_COL.CUSTOMER]       || '').trim(),
        jobDescription: String(r[JO_COL.JOB_DESCRIPTION]|| '').trim(),
        category:       String(r[JO_COL.CATEGORY]       || '').trim(),
        width:          widthRaw,
        height:         heightRaw,
        quantity:       parseInt(r[JO_COL.QUANTITY])    || 1,
        unit:           unit,
        widthFt:        prism_toFt_(widthRaw,  unit),
        heightFt:       prism_toFt_(heightRaw, unit),
        plottingLink:   String(r[JO_COL.PLOTTING_LINK]  || '').trim(),
        status:         status,
        rollId:         String(r[JO_COL.ROLL_ID]        || '').trim(),
        dateCreated:    prism_fmtShort_(r[JO_COL.DATE_CREATED]),
        plottedBefore:  (usage.count || 0) > 0,
        usageCount:     usage.count     || 0,
        usageRollIds:   usage.rollIds   || [],
        usageLastRollId: usage.lastRollId || '',
        usageLastDate:   usage.lastDate   || ''
      });
    });

    data.sort((a, b) => {
      const dd = new Date(b.dateCreated) - new Date(a.dateCreated);
      return dd !== 0 ? dd : a.rowIndex - b.rowIndex;
    });

    return { success: true, data };
  } catch (e) {
    Logger.log('prism_getForPlottingJOs ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

function prism_submitJobOrder(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!payload.joNumber || !payload.joNumber.trim()) return { success: false, message: 'JO Number required.' };
    if (!payload.customer || !payload.customer.trim()) return { success: false, message: 'Customer required.' };
    if (!payload.category) return { success: false, message: 'Category required.' };
    const existing = prism_getAllJobOrders_();
    if (existing.some(j => j.joNumber.toLowerCase() === payload.joNumber.trim().toLowerCase()))
      return { success: false, message: 'JO "' + payload.joNumber + '" already exists.' };

    const sh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const today = new Date();

    const LFP_CATEGORIES = ['banner', 'sticker', 'signage', 'canvas', 'tarpaulin', 'standees/display'];
    const isLFP = LFP_CATEGORIES.includes((payload.category || '').toLowerCase());
    const unit = (payload.unit || 'ft').toString().trim().toLowerCase();
    const status = isLFP ? 'FOR_PLOTTING' : (payload.status || JO_STATUS.FOR_PLOTTING);

    sh.getRange(sh.getLastRow() + 1, 1, 1, 13).setValues([[
      payload.joNumber.trim().toUpperCase(), payload.customer.trim(),
      payload.jobDescription || '', payload.category,
      parseFloat(payload.width) || 0, parseFloat(payload.height) || 0, parseInt(payload.quantity) || 1,
      unit,                        // column H = Unit (replaced ProductionType)
      payload.plottingLink || '',
      status,
      '', user.email, today
    ]]);

    prism_audit_('PRISM_SUBMIT_JO', { joNumber: payload.joNumber, by: user.email });
    return { success: true, message: 'JO "' + payload.joNumber + '" submitted.' };
  } catch (e) { return { success: false, message: e.message }; }
}
function prism_updateJOStatus(joNumber, newStatus) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isSTL_(user.role))
      return { success: false, message: 'Admin or Senior Team Leader only.' };
    const sh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const lr = sh.getLastRow();
    if (lr < 2) return { success: false, message: 'No JOs found.' };
    const data = sh.getRange(2, 1, lr - 1, 13).getValues();
    let updated = 0;
    data.forEach((r, i) => {
      if (String(r[JO_COL.JO_NUMBER]).trim().toUpperCase() === joNumber.trim().toUpperCase()) {
        sh.getRange(i + 2, JO_COL.STATUS + 1).setValue(newStatus);
        updated++;
      }
    });
    if (!updated) return { success: false, message: 'JO "' + joNumber + '" not found.' };
    prism_audit_('PRISM_UPDATE_JO_STATUS', { joNumber, newStatus, by: user.email });
    return { success: true, message: 'JO "' + joNumber + '" → ' + newStatus };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
//  PLOTTING LOG
// ============================================================
function prism_submitPlottingLog(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role))
      return { success: false, message: 'Access denied.' };
    if (!payload.joNumber)
      return { success: false, message: 'JO Number required.' };
    if (!payload.plottingLink)
      return { success: false, message: 'Plotting link required.' };
 
    const today  = new Date();
    const plotId = 'PLT-' + today.getTime();
    const plotSh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
 
    prism_writePlotRow_(plotSh, {
      plotId,
      type:      PLOT_TYPE.PLOT,
      rollId:    payload.rollId || '',
      joNumbers: [payload.joNumber.trim().toUpperCase()],
      status:    PLOT_STATUS.PLANNED,
      startFt:   0,
      endFt:     0,
      lengthFt:  0,
      isVoid:    false,
      isReprint: false,
      pngUrl:    payload.plottingLink || '',
      operator:  user.email,
      date:      today,
      remarks:   { notes: payload.remarks || '' }
    });
 
    // Update JO plotting link
    const joSh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const lr   = joSh.getLastRow();
    if (lr >= 2) {
      const data = joSh.getRange(2, 1, lr - 1, 13).getValues();
      data.forEach((r, i) => {
        if (String(r[JO_COL.JO_NUMBER]).trim().toUpperCase() === payload.joNumber.trim().toUpperCase()) {
          joSh.getRange(i + 2, JO_COL.PLOTTING_LINK + 1).setValue(payload.plottingLink);
          if (String(r[JO_COL.STATUS]).trim() === JO_STATUS.FOR_PLOTTING)
            joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.READY_TO_PRINT);
        }
      });
    }
 
    prism_audit_('PRISM_SUBMIT_PLOT', { joNumber: payload.joNumber, by: user.email });
    return { success: true, message: 'Plotting log saved for ' + payload.joNumber, plotId };
  } catch(e) { return { success: false, message: e.message }; }
}


function prism_getPlottingLog() {
  try {
    const plotSh   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
 
    return {
      success: true,
      data: plotRows.map(r => ({
        plotId:      r.plotId,
        type:        r.type,
        rollId:      r.rollId,
        joNumbers:   r.joNumbers.join(', '),
        status:      r.status,
        startFt:     r.startFt,
        endFt:       r.endFt,
        lengthFt:    r.lengthFt,
        isVoid:      r.isVoid,
        isReprint:   r.isReprint,
        pngUrl:      r.pngUrl,
        operator:    r.operator,
        datePlotted: r.date ? prism_fmtShort_(new Date(r.date)) : '',
        remarks:     JSON.stringify(r.remarks)
      }))
    };
  } catch(e) { return { success: false, message: e.message }; }
}

function prism_compactRollRows_(rows) {
  return (Array.isArray(rows) ? rows : []).map(r => ({
    rowH: parseFloat(r.rowH) || 0,
    usedW: parseFloat(r.usedW) || 0,
    wasteW: parseFloat(r.wasteW) || 0,
    pieces: (Array.isArray(r.pieces) ? r.pieces : []).map(p => ({
      jo: String(p.joNumber || '').trim().toUpperCase(),
      w: parseFloat(p.w) || 0,
      h: parseFloat(p.h) || 0,
      dx: p.dx !== undefined ? (parseFloat(p.dx) || 0) : undefined,
      dy: p.dy !== undefined ? (parseFloat(p.dy) || 0) : undefined,
      r: !!p.rotated
    }))
  }));
}

function prism_getRollPlotHistoryMap(rollIds) {
  try {
    const wanted = (Array.isArray(rollIds) ? rollIds : [])
      .map(x => String(x||'').trim()).filter(Boolean);
    if (!wanted.length) return { success: true, data: {} };
 
    const wantedSet = {};
    wanted.forEach(id => { wantedSet[id] = true; });
    const out = {};
    wanted.forEach(id => { out[id] = []; });
 
    const plotSh   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
 
    plotRows.forEach(row => {
      if (!wantedSet[row.rollId]) return;
      if (row.isVoid) return;
      if (row.status !== PLOT_STATUS.PRINTED &&
          row.type   !== PLOT_TYPE.DAMAGE    &&
          row.type   !== PLOT_TYPE.TEST_PRINT) return;
 
      out[row.rollId].push({
        plotId:      row.plotId,
        rollId:      row.rollId,
        rollWidth:   parseFloat(row.remarks.rollWidth) || 0,
        originalLength: parseFloat(row.remarks.originalLength) || 0,
        startAtFt:   row.startFt,
        endAtFt:     row.endFt,
        lengthUsed:  row.lengthFt,
        joNumbers:   row.joNumbers,
        isDamage:    row.type === PLOT_TYPE.DAMAGE,
        isTestPrint: row.type === PLOT_TYPE.TEST_PRINT,
        rows:        row.remarks.rows || [],
        pngUrl:      row.pngUrl,
        dateMs:      row.date ? new Date(row.date).getTime() : 0,
        source:      'PLOT_LOG'
      });
    });
 
    // Sort each roll by startAtFt
    Object.keys(out).forEach(id => {
      out[id].sort((a,b) => a.startAtFt - b.startAtFt);
    });
 
    return { success: true, data: out };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
// ============================================================
//  USAGE LOG
// ============================================================
function prism_getUsageLog() {
  try {
    const sh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    const lr = sh.getLastRow();
    if (lr < 2) return { success: true, data: [] };
    return {
      success: true, data: sh.getRange(2, 1, lr - 1, 8).getValues()
        .filter(r => r[USAGE_COL.USAGE_ID])
        .map(r => ({
          usageId: String(r[USAGE_COL.USAGE_ID]).trim(),
          joNumber: String(r[USAGE_COL.JO_NUMBER]).trim(),
          rollId: String(r[USAGE_COL.ROLL_ID]).trim(),
          widthUsed: parseFloat(r[USAGE_COL.WIDTH_USED]) || 0,
          lengthUsed: parseFloat(r[USAGE_COL.LENGTH_USED]) || 0,
          operator: String(r[USAGE_COL.OPERATOR]).trim(),
          plottingLink: String(r[USAGE_COL.PLOTTING_LINK] || '').trim(),
          dateUsed: prism_fmtShort_(r[USAGE_COL.DATE_USED])
        }))
    };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
//  REMOVE FROM QUEUE (void a PLANNED layout, revert JOs)
// ============================================================
function prism_removeFromQueue(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role))
      return { success: false, message: 'Access denied.' };

    const plotId = String(payload.plotId || '').trim();
    if (!plotId) return { success: false, message: 'No plot ID provided.' };

    const plotSh   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
    const target   = plotRows.find(r => r.plotId === plotId);

    if (!target)
      return { success: false, message: 'Plot layout not found.' };
    if (target.status !== PLOT_STATUS.PLANNED || target.isVoid)
      return { success: false, message: 'Only PLANNED (waiting) layouts can be removed.' };

    // Void the plot entry
    plotSh.getRange(target._rowIdx, PLOT_COL.STATUS  + 1).setValue(PLOT_STATUS.VOIDED);
    plotSh.getRange(target._rowIdx, PLOT_COL.IS_VOID + 1).setValue(true);

    // Revert JOs back to FOR_PLOTTING
    const joNumbers = target.joNumbers;
    let reverted = 0;
    if (joNumbers.length) {
      const joSh   = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
      const joData = joSh.getRange(2, 1, joSh.getLastRow() - 1, 13).getValues();
      joData.forEach((r, i) => {
        const jo = String(r[JO_COL.JO_NUMBER]).trim().toUpperCase();
        if (joNumbers.includes(jo) && String(r[JO_COL.STATUS]).trim() === JO_STATUS.READY_TO_PRINT) {
          joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.FOR_PLOTTING);
          reverted++;
        }
      });
    }

    prism_audit_('PRISM_REMOVE_FROM_QUEUE', { plotId, joNumbers, reverted, by: user.email });
    return {
      success: true,
      message: `Layout removed. ${reverted} JO(s) returned to FOR_PLOTTING.`
    };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
//  ADMIN: Manual roll status override
// ============================================================
function prism_setRollStatus(rollId, newStatus) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role)) return { success: false, message: 'Admins only.' };
    const valid = Object.values(ROLL_STATUS);
    if (!valid.includes(newStatus.toUpperCase())) return { success: false, message: 'Invalid status.' };
    const sh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const lr = sh.getLastRow();
    const data = sh.getRange(2, 1, lr - 1, 1).getValues();
    let rowIdx = -1;
    data.forEach((r, i) => { if (String(r[0]).trim() === rollId) rowIdx = i + 2; });
    if (rowIdx === -1) return { success: false, message: 'Roll not found.' };
    sh.getRange(rowIdx, ROLL_COL.STATUS + 1).setValue(newStatus.toUpperCase());
    prism_audit_('PRISM_SET_ROLL_STATUS', { rollId, newStatus, by: user.email });
    return { success: true, message: rollId + ' → ' + newStatus };
  } catch (e) { return { success: false, message: e.message }; }
}

function prism_getStockMaterialsList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const linkSheet = ss.getSheetByName('DatabaseLink');
    if (!linkSheet) throw new Error('DatabaseLink sheet not found');

    const rows = linkSheet.getRange(2, 1, linkSheet.getLastRow() - 1, 2).getValues();
    const match = rows.find(r => r[0].toString().trim() === 'StockDatabase');
    if (!match) throw new Error('StockDatabase not found in DatabaseLink');

    const idMatch = match[1].toString().match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (!idMatch) throw new Error('Invalid StockDatabase URL');

    const sheet = SpreadsheetApp.openById(idMatch[1]).getSheetByName('AllItems');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
      .filter(r => r[0] && r[1])
      .map(r => ({
        itemCode: String(r[0]).trim(),
        itemDesc: String(r[1]).trim(),
        stockOnHand: r[6] !== '' ? Number(r[6]) : 0,
        unitCost: r[7] !== '' ? Number(r[7]) : 0
      }));

    return { success: true, data };
  } catch (e) {
    Logger.log('prism_getStockMaterialsList ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ============================================================
//  PLOTTING PLANNER — Confirm layout
//  Add this function to Code.js
// ============================================================
function prism_confirmPlotLayout(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role))
      return { success: false, message: 'Admin access required.' };

    if (payload.pendingDamages && payload.pendingDamages.length) {
      for (let i = 0; i < payload.pendingDamages.length; i++) {
        const d = payload.pendingDamages[i];
        const res = prism_declareDamage({
          rollId: d.rollId,
          lengthUsed: parseFloat(d.lengthUsed),
          joNumbers: d.joNumbers || [],
          remarks: 'Auto-Declared and Replotted from Planner'
        });
        if (!res.success) {
          throw new Error('Failed to process damage for roll ' + d.rollId + ': ' + res.message);
        }
      }
      SpreadsheetApp.flush();
    }

    if (!payload.joNumbers || !payload.joNumbers.length)
      return { success: false, message: 'No JOs provided.' };

    let rollPlans = [];
    if (Array.isArray(payload.rollPlans) && payload.rollPlans.length) {
      rollPlans = payload.rollPlans
        .map(p => ({
          rollId: String(p.rollId || '').trim(),
          lengthUsed: parseFloat(p.lengthUsed) || 0,
          rows: Array.isArray(p.rows) ? p.rows : []
        }))
        .filter(p => p.rollId && p.lengthUsed > 0);
    } else {
      if (!payload.rollId)
        return { success: false, message: 'Roll ID required.' };
      if (!payload.lengthUsed || payload.lengthUsed <= 0)
        return { success: false, message: 'Length used must be > 0.' };
      rollPlans = [{
        rollId: String(payload.rollId).trim(),
        lengthUsed: parseFloat(payload.lengthUsed) || 0,
        rows: Array.isArray(payload.rows) ? payload.rows : []
      }];
    }
    if (!rollPlans.length)
      return { success: false, message: 'No valid roll plan found.' };

    const rollJoNumbersById = {};
    rollPlans.forEach(plan => {
      const joSet = {};
      (Array.isArray(plan.rows) ? plan.rows : []).forEach(row => {
        (Array.isArray(row.pieces) ? row.pieces : []).forEach(piece => {
          const jo = String(piece.joNumber || piece.jo || '').trim().toUpperCase();
          if (jo) joSet[jo] = true;
        });
      });
      rollJoNumbersById[plan.rollId] = Object.keys(joSet);
    });

    // rowIndexes: new per-row targeting (sent by updated frontend)
    const rowIndexes = Array.isArray(payload.rowIndexes)
      ? payload.rowIndexes.map(Number).filter(Boolean)
      : [];

    // partialQtyMap + rowIndexToJO: handle partial quantity printing
    const partialQtyMap = (payload.partialQtyMap && typeof payload.partialQtyMap === 'object')
      ? payload.partialQtyMap : {};
    const rowIndexToJO = (payload.rowIndexToJO && typeof payload.rowIndexToJO === 'object')
      ? payload.rowIndexToJO : {};

    // Build joNumber → { printedQty, fullQty }
    const joPartialMap = {};
    Object.keys(rowIndexToJO).forEach(ri => {
      const joNum = String(rowIndexToJO[ri] || '').trim().toUpperCase();
      const pq = partialQtyMap[ri];
      if (joNum && pq) joPartialMap[joNum] = pq;
    });

    const plotImages = Array.isArray(payload.plotImages) ? payload.plotImages : [];
    const savedPlotsByRollId = {};
    let effectivePlottingLink = String(payload.plottingLink || '').trim();
    if (plotImages.length) {
      plotImages.forEach(img => {
        const rollId = String((img && img.rollId) || '').trim();
        const imageDataUrl = String((img && img.imageDataUrl) || '').trim();
        if (!rollId || !imageDataUrl) return;
        savedPlotsByRollId[rollId] = prism_saveRollMapToDrive_({
          rollId: rollId,
          imageDataUrl: imageDataUrl,
          userEmail: user.email,
          joNumbers: rollJoNumbersById[rollId] || []
        });
      });
      const firstSaved = savedPlotsByRollId[rollPlans[0].rollId] || savedPlotsByRollId[Object.keys(savedPlotsByRollId)[0]];
      if (firstSaved) effectivePlottingLink = String(firstSaved.pngUrl || '').trim();
    } else if (payload.imageDataUrl) {
      const fallbackSavedPlot = prism_saveRollMapToDrive_({
        rollId: String(payload.plotRollId || rollPlans[0].rollId || '').trim(),
        imageDataUrl: payload.imageDataUrl,
        userEmail: user.email,
        joNumbers: rollJoNumbersById[String(payload.plotRollId || rollPlans[0].rollId || '').trim()] || []
      });
      savedPlotsByRollId[String(payload.plotRollId || rollPlans[0].rollId || '').trim()] = fallbackSavedPlot;
      effectivePlottingLink = String(fallbackSavedPlot.pngUrl || '').trim();
    }

    // ── Validate + deduct from each roll ──
    const rollSh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollLr = rollSh.getLastRow();
    const rollData = rollSh.getRange(2, 1, rollLr - 1, 9).getValues();
    const rollMap = {};
    rollData.forEach((r, i) => {
      rollMap[String(r[ROLL_COL.ROLL_ID] || '').trim()] = { rowIdx: i + 2, row: r };
    });

    const rollUsageById = {};
    rollPlans.forEach(p => { rollUsageById[p.rollId] = (rollUsageById[p.rollId] || 0) + p.lengthUsed; });

    Object.keys(rollUsageById).forEach(rollId => {
      const entry = rollMap[rollId];
      if (!entry) throw new Error('Roll "' + rollId + '" not found.');
      if (String(entry.row[ROLL_COL.STATUS]).trim().toUpperCase() !== ROLL_STATUS.OPEN)
        throw new Error('Roll "' + rollId + '" must be OPEN.');
      const remaining = parseFloat(entry.row[ROLL_COL.REMAINING_LENGTH]) || 0;
      if (rollUsageById[rollId] > remaining + 0.0001)
        throw new Error('Roll "' + rollId + '" does not have enough remaining length.');
    });

    const rollSnapshots = {};
    Object.keys(rollUsageById).forEach(rollId => {
      const entry = rollMap[rollId];
      const remaining = parseFloat(entry.row[ROLL_COL.REMAINING_LENGTH]) || 0;
      const rollWidth = parseFloat(entry.row[ROLL_COL.WIDTH]) || 0;
      const originalLength = parseFloat(entry.row[ROLL_COL.ORIGINAL_LENGTH]) || 0;
      // No setValue calls — roll sheet is NOT touched during plotting
      rollSnapshots[rollId] = {
        width: rollWidth,
        originalLength: originalLength,
        startAtFt: Math.max(0, originalLength - remaining),
        endAtFt: Math.max(0, originalLength - remaining) + rollUsageById[rollId],
        newRemaining: Math.max(0, remaining - rollUsageById[rollId]),
        newStatus: ROLL_STATUS.OPEN
      };
    });

    const today = new Date();
    const joSh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const joLr = joSh.getLastRow();
    const joData = joLr >= 2 ? joSh.getRange(2, 1, joLr - 1, 13).getValues() : [];
    const totalLengthUsed = rollPlans.reduce((s, p) => s + (parseFloat(p.lengthUsed) || 0), 0);

    // Calculate per-roll JO membership (still needed to assign Roll IDs to JOs)
    const joRollLengths = {};
    rollPlans.forEach(plan => {
      (Array.isArray(plan.rows) ? plan.rows : []).forEach(row => {
        (Array.isArray(row.pieces) ? row.pieces : []).forEach(p => {
          const jo = String(p.joNumber || '').trim().toUpperCase();
          if (!jo) return;
          if (!joRollLengths[jo]) joRollLengths[jo] = {};
          joRollLengths[jo][plan.rollId] = true;
        });
      });
    });

    if (rowIndexes.length > 0) {
      // NEW PATH: only update the exact sheet rows that were confirmed
      joData.forEach((r, i) => {
        const sheetRow  = i + 2;
        const joNumber  = String(r[JO_COL.JO_NUMBER]).trim().toUpperCase();
        const rowStatus = String(r[JO_COL.STATUS]).trim();
        if (rowStatus !== JO_STATUS.FOR_PLOTTING) return;
        if (!rowIndexes.includes(sheetRow)) return;

        const existingRolls = String(r[JO_COL.ROLL_ID] || '')
          .split(',').map(x => x.trim()).filter(Boolean);
        Object.keys(joRollLengths[joNumber] || {}).forEach(rollId => {
          if (rollId && !existingRolls.includes(rollId)) existingRolls.push(rollId);
        });
        joSh.getRange(sheetRow, JO_COL.ROLL_ID + 1).setValue(existingRolls.join(', '));

        const pq = joPartialMap[joNumber];
        const isPartial = pq && pq.printedQty < pq.fullQty;
        if (isPartial) {
          const remaining = pq.fullQty - pq.printedQty;
          joSh.getRange(sheetRow, JO_COL.QUANTITY + 1).setValue(remaining);
          joSh.getRange(sheetRow, JO_COL.STATUS + 1).setValue(JO_STATUS.FOR_PLOTTING);
          prism_audit_('PRISM_PARTIAL_PRINT', { joNumber, printedQty: pq.printedQty, remaining, by: user.email });
        } else {
          joSh.getRange(sheetRow, JO_COL.STATUS + 1).setValue(JO_STATUS.READY_TO_PRINT);
        }
        if (effectivePlottingLink)
          joSh.getRange(sheetRow, JO_COL.PLOTTING_LINK + 1).setValue(effectivePlottingLink);
      });
    } else {
      // LEGACY PATH: update all rows matching joNumbers (old behavior)
      payload.joNumbers.forEach(joRaw => {
        const joNumber = String(joRaw || '').trim().toUpperCase();
        if (!joNumber) return;
        joData.forEach((r, i) => {
          if (String(r[JO_COL.JO_NUMBER]).trim().toUpperCase() !== joNumber) return;
          const existingRolls = String(r[JO_COL.ROLL_ID] || '')
            .split(',').map(x => x.trim()).filter(Boolean);
          Object.keys(joRollLengths[joNumber] || {}).forEach(rollId => {
            if (rollId && !existingRolls.includes(rollId)) existingRolls.push(rollId);
          });
          joSh.getRange(i + 2, JO_COL.ROLL_ID + 1).setValue(existingRolls.join(', '));
          joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.READY_TO_PRINT);
          if (effectivePlottingLink)
            joSh.getRange(i + 2, JO_COL.PLOTTING_LINK + 1).setValue(effectivePlottingLink);
        });
      });
    }

    // Persist roll layout snapshots
    const plotSh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    rollPlans.forEach(plan => {
      const rollId = plan.rollId;
      const snap   = rollSnapshots[rollId] || {};
      const joSet  = {};
      (Array.isArray(plan.rows) ? plan.rows : []).forEach(r => {
        (Array.isArray(r.pieces) ? r.pieces : []).forEach(p => {
          const jo = String(p.joNumber || '').trim().toUpperCase();
          if (jo) joSet[jo] = true;
        });
      });
      const joList   = Object.keys(joSet);
      const plotPngUrl = (savedPlotsByRollId[rollId] && savedPlotsByRollId[rollId].pngUrl) || effectivePlottingLink || '';
      const plotId   = 'PLN-' + today.getTime() + '-' + rollId;
 
      prism_writePlotRow_(plotSh, {
        plotId,
        type:      PLOT_TYPE.PLOT,
        rollId,
        joNumbers: joList,
        status:    PLOT_STATUS.PLANNED,
        startFt:   0,
        endFt:     parseFloat(plan.lengthUsed) || 0,
        lengthFt:  parseFloat(plan.lengthUsed) || 0,
        isVoid:    false,
        isReprint: false,
        pngUrl:    plotPngUrl,
        operator:  user.email,
        date:      today,
        remarks: {
          rollWidth:      snap.width || 0,
          originalLength: snap.originalLength || 0,
          folderUrl:      (savedPlotsByRollId[rollId] && savedPlotsByRollId[rollId].folderUrl) || '',
          rows:           prism_compactRollRows_(plan.rows)
        }
      });
    });

    prism_audit_('PRISM_CONFIRM_PLOT_LAYOUT', {
      rollPlans: rollPlans.map(p => ({ rollId: p.rollId, lengthUsed: p.lengthUsed })),
      totalLengthUsed: Number(totalLengthUsed.toFixed(3)),
      joNumbers: payload.joNumbers,
      by: user.email
    });

    return {
      success: true,
      message: 'Layout confirmed! ' + payload.joNumbers.length
        + ' JO(s) set to READY_TO_PRINT. '
        + totalLengthUsed.toFixed(1) + 'ft planned across '
        + rollPlans.length + ' roll(s). Roll deducted when printing is marked done.',
      plottingLink: effectivePlottingLink || '',
      savedPlots: Object.keys(savedPlotsByRollId).map(rollId => ({
        rollId: rollId,
        fileBaseName: savedPlotsByRollId[rollId].fileBaseName,
        pngUrl: savedPlotsByRollId[rollId].pngUrl,
        folderUrl: savedPlotsByRollId[rollId].folderUrl
      }))
    };

  } catch (e) { return { success: false, message: e.message }; }
}

function prism_exportPlottingSheet(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role)) {
      return { success: false, message: 'Admin or Digital Operator access required.' };
    }

    const rollPlans = Array.isArray(payload.rollPlans) ? payload.rollPlans : [];
    if (!rollPlans.length) return { success: false, message: 'No roll plan to export.' };

    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
    const ss = SpreadsheetApp.create('PRISM Plot Plan ' + stamp);
    const sh = ss.getSheets()[0];
    sh.setName('PlotPlan');

    let row = 1;
    sh.getRange(row, 1, 1, 8).merge();
    sh.getRange(row, 1).setValue('PRISM Smart Plotting Sheet');
    sh.getRange(row, 1).setFontSize(14).setFontWeight('bold');
    row++;
    sh.getRange(row, 1, 1, 8).merge();
    sh.getRange(row, 1).setValue('Generated ' + prism_fmtShort_(new Date()) + ' by ' + user.email);
    sh.getRange(row, 1).setFontColor('#4b5563');
    row += 2;

    rollPlans.forEach((plan, idx) => {
      const rollId = String(plan.rollId || '').trim();
      const rollWidth = parseFloat(plan.rollWidth) || 0;
      const rollAvail = parseFloat(plan.rollAvail) || 0;
      const lengthUsed = parseFloat(plan.lengthUsed) || 0;
      const rows = Array.isArray(plan.rows) ? plan.rows : [];

      sh.getRange(row, 1, 1, 8).merge();
      sh.getRange(row, 1).setValue(
        'Roll ' + rollId + '  |  ' + rollWidth + 'ft x ' + rollAvail + 'ft  |  Est. Used: ' + lengthUsed.toFixed(1) + 'ft'
      );
      sh.getRange(row, 1).setFontWeight('bold').setBackground('#e6f4f1');
      row++;

      sh.getRange(row, 1, 1, 8).setValues([[
        'Row', 'Pieces', 'Row Height (ft)', 'Used Width (ft)', 'Waste Width (ft)',
        'Cumulative Length (ft)', 'Roll', 'Notes'
      ]]);
      sh.getRange(row, 1, 1, 8).setFontWeight('bold').setBackground('#f3f4f6');
      row++;

      let cumLen = 0;
      rows.forEach((r, ri) => {
        cumLen += parseFloat(r.rowH) || 0;
        const pcs = (r.pieces || []).map(p => {
          return String(p.joNumber || '') + ' (' + (parseFloat(p.w) || 0) + 'x' + (parseFloat(p.h) || 0) + 'ft' + (p.rotated ? ' rot' : '') + ')';
        }).join(' | ');

        sh.getRange(row, 1, 1, 8).setValues([[
          'R' + (ri + 1),
          pcs,
          parseFloat(r.rowH) || 0,
          parseFloat(r.usedW) || 0,
          parseFloat(r.wasteW) || 0,
          Number(cumLen.toFixed(2)),
          rollId,
          idx === 0 && ri === 0 ? 'primary roll' : ''
        ]]);
        row++;
      });

      if (!rows.length) {
        sh.getRange(row, 1, 1, 8).setValues([['-', 'No rows allocated', '', '', '', '', rollId, '']]);
        row++;
      }

      row++;
    });

    sh.autoResizeColumns(1, 8);
    prism_audit_('PRISM_EXPORT_PLOT_SHEET', {
      spreadsheetId: ss.getId(),
      rollPlanCount: rollPlans.length,
      by: user.email
    });

    return {
      success: true,
      url: ss.getUrl(),
      spreadsheetId: ss.getId(),
      message: 'Plot sheet exported.'
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function prism_saveRollMapToDrive_(payload) {
  const rollId = String((payload && payload.rollId) || '').trim();
  const imageDataUrl = String((payload && payload.imageDataUrl) || '').trim();
  const userEmail = String((payload && payload.userEmail) || '').trim();
  const joNumbers = Array.isArray(payload && payload.joNumbers) ? payload.joNumbers : [];
  if (!rollId) throw new Error('Roll ID is required.');
  if (!imageDataUrl) throw new Error('Image data is required.');

  const m = imageDataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!m) throw new Error('Invalid image format.');

  const contentType = m[1];
  const b64 = m[2];
  const bytes = Utilities.base64Decode(b64);
  const safeRollId = rollId.replace(/[\\/:*?"<>|]/g, '_').trim() || 'ROLL';
  const rollMeta = prism_getRollDriveMeta_(rollId);
  const sizeFolderName = prism_buildRollSizeFolderName_(rollMeta);
  const safeJoNames = joNumbers
    .map(jo => String(jo || '').trim().toUpperCase())
    .filter(Boolean)
    .filter((jo, idx, arr) => arr.indexOf(jo) === idx)
    .map(jo => jo.replace(/[^A-Z0-9._-]+/g, '_'));
  const joSuffix = safeJoNames.length
    ? ' - ' + safeJoNames.slice(0, 3).join('_') + (safeJoNames.length > 3 ? '_and-' + (safeJoNames.length - 3) + '-more' : '')
    : '';
  const fileBaseName = (safeRollId + joSuffix).slice(0, 180);

  const rootFolder = DriveApp.getFolderById(PRISM_PLOTTING_DRIVE_FOLDER_ID);
  const sizeFolder = prism_getOrCreateDriveChildFolder_(rootFolder, sizeFolderName);
  const folder = prism_getOrCreateDriveChildFolder_(sizeFolder, safeRollId);
  const pngBlob = Utilities.newBlob(bytes, contentType, fileBaseName + '.png');
  const pngFile = folder.createFile(pngBlob);

  prism_audit_('PRISM_SAVE_ROLL_MAP_DRIVE', {
    rollId: rollId,
    sizeFolderName: sizeFolderName,
    sizeFolderId: sizeFolder.getId(),
    rollFolderId: folder.getId(),
    pngFileId: pngFile.getId(),
    folderId: PRISM_PLOTTING_DRIVE_FOLDER_ID,
    by: userEmail
  });

  return {
    fileBaseName: fileBaseName,
    pngUrl: pngFile.getUrl(),
    folderUrl: folder.getUrl()
  };
}

function prism_getOrCreateDriveChildFolder_(parentFolder, folderName) {
  const safeName = String(folderName || '').trim() || 'Uncategorized';
  const existing = parentFolder.getFoldersByName(safeName);
  return existing.hasNext() ? existing.next() : parentFolder.createFolder(safeName);
}

function prism_getRollDriveMeta_(rollId) {
  const sh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
  const lr = sh.getLastRow();
  if (lr < 2) return { materialCode: '', materialName: '', width: 0, originalLength: 0 };

  const rows = sh.getRange(2, 1, lr - 1, 4).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][ROLL_COL.ROLL_ID] || '').trim() !== rollId) continue;
    const materialCode = String(rows[i][ROLL_COL.MATERIAL_CODE] || '').trim();
    const materialMeta = prism_getMaterialDriveMeta_(materialCode);
    return {
      materialCode: materialCode,
      materialName: materialMeta.materialName || '',
      width: parseFloat(rows[i][ROLL_COL.WIDTH]) || 0,
      originalLength: parseFloat(rows[i][ROLL_COL.ORIGINAL_LENGTH]) || 0
    };
  }
  return { materialCode: '', materialName: '', width: 0, originalLength: 0 };
}

function prism_buildRollSizeFolderName_(rollMeta) {
  const materialCode = String(rollMeta && rollMeta.materialCode || '').trim();
  const materialName = String(rollMeta && rollMeta.materialName || '').trim();
  const width = parseFloat(rollMeta && rollMeta.width) || 0;
  const originalLength = parseFloat(rollMeta && rollMeta.originalLength) || 0;
  if (materialCode && materialName) return materialCode + ' - ' + materialName;
  if (materialName) return materialName;
  if (width > 0 && originalLength > 0) return width + 'ft x ' + originalLength + 'ft';
  if (width > 0) return width + 'ft wide';
  return 'Unknown Size';
}

function prism_getMaterialDriveMeta_(materialCode) {
  const code = String(materialCode || '').trim();
  if (!code) return { materialName: '' };

  const sh = prism_sh_(PRISM_SHEETS.LFP_MATERIALS);
  const lr = sh.getLastRow();
  if (lr < 2) return { materialName: '' };

  const rows = sh.getRange(2, 1, lr - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][MAT_COL.MATERIAL_CODE] || '').trim() !== code) continue;
    return {
      materialName: String(rows[i][MAT_COL.MATERIAL_NAME] || '').trim()
    };
  }
  return { materialName: '' };
}

function prism_saveRollMapToDrive(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role)) {
      return { success: false, message: 'Admin or Digital Operator access required.' };
    }

    const saved = prism_saveRollMapToDrive_({
      rollId: String((payload && payload.rollId) || '').trim(),
      imageDataUrl: String((payload && payload.imageDataUrl) || '').trim(),
      userEmail: user.email,
      joNumbers: Array.isArray(payload && payload.joNumbers) ? payload.joNumbers : []
    });

    return {
      success: true,
      message: 'Roll map saved to Drive.',
      fileBaseName: saved.fileBaseName,
      pngUrl: saved.pngUrl,
      folderUrl: saved.folderUrl
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}


function prism_startPrintingLayout(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role))
      return { success: false, message: 'Access denied.' };
 
    const plotId = String(payload.plotId || '').trim();
    if (!plotId) return { success: false, message: 'No plot ID provided.' };
 
    const plotSh   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
 
    // Find target row
    const target = plotRows.find(r => r.plotId === plotId);
    if (!target) return { success: false, message: 'Plot layout not found.' };
    if (target.status !== PLOT_STATUS.PLANNED || target.isVoid)
      return { success: false, message: 'Layout is not PLANNED or has been voided.' };
 
    const targetRollId = target.rollId;
    const targetLength = target.lengthFt;
 
    // Guard: only one PRINTING at a time per roll
    const alreadyPrinting = plotRows.find(r =>
      r.rollId === targetRollId && !r.isVoid && r.status === PLOT_STATUS.PRINTING
    );
    if (alreadyPrinting)
      return { success: false, message: 'Another layout is already printing on roll ' + targetRollId + '.' };
 
    // Find furthest PRINTED or PRINTING end on this roll
    let maxEnd = 0;
    plotRows.forEach(r => {
      if (r.rollId === targetRollId && !r.isVoid &&
         (r.status === PLOT_STATUS.PRINTED || r.status === PLOT_STATUS.PRINTING)) {
        if (r.endFt > maxEnd) maxEnd = r.endFt;
      }
    });
 
    // Update status to PRINTING + set startFt/endFt
    plotSh.getRange(target._rowIdx, PLOT_COL.STATUS   + 1).setValue(PLOT_STATUS.PRINTING);
    plotSh.getRange(target._rowIdx, PLOT_COL.START_FT + 1).setValue(maxEnd);
    plotSh.getRange(target._rowIdx, PLOT_COL.END_FT   + 1).setValue(maxEnd + targetLength);
 
    // Mark JOs as PRINTING
    const joNumbers = target.joNumbers;
    if (joNumbers.length) {
      const joSh   = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
      const joData = joSh.getRange(2, 1, joSh.getLastRow() - 1, 13).getValues();
      joData.forEach((r, i) => {
        const jo = String(r[JO_COL.JO_NUMBER]).trim().toUpperCase();
        if (joNumbers.includes(jo) && String(r[JO_COL.STATUS]).trim() === JO_STATUS.READY_TO_PRINT)
          joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.PRINTING);
      });
    }
 
    prism_audit_('PRISM_START_PRINTING_LAYOUT', { plotId, rollId: targetRollId, by: user.email });
    return { success: true, message: 'Layout sent to printer!', startAtFt: maxEnd, endAtFt: maxEnd + targetLength };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


function prism_markPrintingComplete(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isOperator_(user.role))
      return { success: false, message: 'Access denied.' };
 
    const plotId = String(payload.plotId || '').trim();
    if (!plotId) return { success: false, message: 'No plot ID provided.' };
 
    const plotSh   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
    const target   = plotRows.find(r => r.plotId === plotId);
 
    if (!target) return { success: false, message: 'Plot layout not found.' };
    if (target.status !== PLOT_STATUS.PRINTING)
      return { success: false, message: 'Layout is not currently PRINTING.' };
 
    // Flip to PRINTED
    plotSh.getRange(target._rowIdx, PLOT_COL.STATUS + 1).setValue(PLOT_STATUS.PRINTED);
 
    const targetLength = target.lengthFt;
    const targetRollId = target.rollId;
 
    const rollSh   = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollData = rollSh.getRange(2, 1, rollSh.getLastRow() - 1, 9).getValues();
    rollData.forEach((r, i) => {
      if (String(r[ROLL_COL.ROLL_ID]).trim() !== targetRollId) return;
      const cur  = parseFloat(r[ROLL_COL.REMAINING_LENGTH]) || 0;
      const next = Math.max(0, cur - targetLength);
      rollSh.getRange(i + 2, ROLL_COL.REMAINING_LENGTH + 1).setValue(next);
      rollSh.getRange(i + 2, ROLL_COL.STATUS + 1).setValue(next <= 0 ? ROLL_STATUS.CONSUMED : ROLL_STATUS.OPEN);
    });
 
    // Write usage log
    const today    = new Date();
    const usageSh  = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    const planRows = Array.isArray(target.remarks.rows) ? target.remarks.rows : [];
    const joLengths = {};
    planRows.forEach(row => {
      const rowJOs = {};
      (row.pieces || []).forEach(p => { rowJOs[String(p.jo || p.joNumber || '').trim().toUpperCase()] = true; });
      Object.keys(rowJOs).forEach(jo => {
        if (!jo) return;
        joLengths[jo] = (joLengths[jo] || 0) + (parseFloat(row.rowH) || 0);
      });
    });

    if (!Object.keys(joLengths).length && target.joNumbers.length) {
      const perJO = targetLength / target.joNumbers.length;
      target.joNumbers.forEach(jo => { joLengths[jo] = perJO; });
    }
    
    Object.keys(joLengths).forEach(jo => {
      const len = joLengths[jo];
      if (len <= 0) return;
      usageSh.getRange(usageSh.getLastRow() + 1, 1, 1, 8).setValues([[
        'USE-' + today.getTime() + '-' + jo + '-' + targetRollId,
        jo, targetRollId,
        target.remarks.rollWidth || 0,
        len, user.email,
        target.pngUrl || '',
        today
      ]]);
    });
 
    // Mark JOs COMPLETED
    if (target.joNumbers.length) {
      const joSh   = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
      const joData = joSh.getRange(2, 1, joSh.getLastRow() - 1, 13).getValues();
      joData.forEach((r, i) => {
        const jo = String(r[JO_COL.JO_NUMBER]).trim().toUpperCase();
        if (target.joNumbers.includes(jo) && String(r[JO_COL.STATUS]).trim() === JO_STATUS.PRINTING)
          joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.COMPLETED);
      });
    }
 
    prism_audit_('PRISM_MARK_PRINTING_COMPLETE', { plotId, rollId: targetRollId, by: user.email });
    return { success: true, message: 'Job marked as COMPLETED. Roll updated.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function prism_getPrintQueueData() {
  try {
    const plotSh   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
 
    // JO customer map
    const joSh   = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const joLr   = joSh.getLastRow();
    const joStatusMap   = {};
    const joCustomerMap = {};
    if (joLr >= 2) {
      joSh.getRange(2, 1, joLr - 1, 13).getValues().forEach(r => {
        const jo   = String(r[JO_COL.JO_NUMBER] || '').trim().toUpperCase();
        const st   = String(r[JO_COL.STATUS]    || '').trim();
        const cust = String(r[JO_COL.CUSTOMER]  || '').trim();
        if (jo) { joStatusMap[jo] = st; joCustomerMap[jo] = cust; }
      });
    }
 
    const planned      = [];
    const reprints     = [];
    let   printing     = null;
    const rollsHistory = {};
 
    plotRows.forEach(row => {
      if (row.isVoid) return;
 
      // Legacy: PLANNED PLOT whose JOs are all done → treat as PRINTED
      let effectiveStatus = row.status;
      if (effectiveStatus === PLOT_STATUS.PLANNED && row.type === PLOT_TYPE.PLOT && row.joNumbers.length) {
        const allDone = row.joNumbers.every(jo => {
          const st = joStatusMap[jo.toUpperCase()] || '';
          return st === JO_STATUS.PRINTING || st === JO_STATUS.COMPLETED;
        });
        if (allDone) effectiveStatus = PLOT_STATUS.PRINTED;
      }
 
      const base = {
        plotId:     row.plotId,
        joNumbers:  row.joNumbers,
        rollId:     row.rollId,
        rollWidth:  parseFloat(row.remarks.rollWidth) || 0,
        lengthUsed: row.lengthFt,
        date:       row.date ? prism_fmtShort_(new Date(row.date)) : '',
        pngUrl:     row.pngUrl,
        rows:       row.remarks.rows || [],
        customers:  row.joNumbers.map(jo => joCustomerMap[jo.toUpperCase()] || jo)
      };
 
      if (effectiveStatus === PLOT_STATUS.PRINTING) {
        const item = Object.assign({}, base, {
          startAtFt: row.startFt,
          endAtFt:   row.endFt,
          isReprint: row.isReprint
        });
        if (!printing || item.startAtFt > printing.startAtFt) printing = item;
 
      } else if (effectiveStatus === PLOT_STATUS.PRINTED ||
                 row.type === PLOT_TYPE.DAMAGE ||
                 row.type === PLOT_TYPE.TEST_PRINT) {
        if (!rollsHistory[row.rollId]) rollsHistory[row.rollId] = [];
        rollsHistory[row.rollId].push({
          plotId:      row.plotId,
          rollId:      row.rollId,
          rollWidth:   parseFloat(row.remarks.rollWidth) || 0,
          startAtFt:   row.startFt,
          endAtFt:     row.endFt,
          lengthUsed:  row.lengthFt,
          joNumbers:   row.joNumbers,
          isDamage:    row.type === PLOT_TYPE.DAMAGE,
          isTestPrint: row.type === PLOT_TYPE.TEST_PRINT,
          isVoid:      false,
          rows:        row.remarks.rows || [],
          pngUrl:      row.pngUrl,
          dateMs:      row.date ? new Date(row.date).getTime() : 0
        });
 
      } else if (effectiveStatus === PLOT_STATUS.PLANNED) {
        const item = Object.assign({}, base, {
          isReprint:    row.isReprint,
          reprintCount: parseInt(row.remarks.reprintCount) || 0,
          reprintOf:    row.remarks.reprintOf || null
        });
        if (row.isReprint) reprints.push(item);
        else               planned.push(item);
      }
    });
 
    // Sort history by startAtFt
    Object.keys(rollsHistory).forEach(id => {
      rollsHistory[id].sort((a,b) => a.startAtFt - b.startAtFt);
    });
 
    return { success: true, planned, reprints, printing, rollsHistory, joCustomerMap };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function prism_getJobOrderReports() {
  try {
    const plotSh   = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotRows = prism_readPlotRows_(plotSh);
 
    // Customer lookup
    const joSh   = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const joLr   = joSh.getLastRow();
    const custMap = {};
    if (joLr >= 2) {
      joSh.getRange(2, 1, joLr - 1, 13).getValues().forEach(r => {
        const jo   = String(r[JO_COL.JO_NUMBER] || '').trim().toUpperCase();
        const cust = String(r[JO_COL.CUSTOMER]  || '').trim();
        if (jo) custMap[jo] = cust;
      });
    }
 
    const damages = plotRows
      .filter(r => r.type === PLOT_TYPE.DAMAGE && !r.isVoid)
      .map(r => ({
        plotId:     r.plotId,
        rollId:     r.rollId,
        joNumbers:  r.joNumbers,
        customer:   [...new Set(r.joNumbers.map(jo => custMap[jo.toUpperCase()] || jo))].join(', '),
        lengthUsed: r.lengthFt,
        operator:   r.operator,
        reason:     r.remarks.damageReason || '',
        date:       r.date ? prism_fmtShort_(new Date(r.date)) : ''
      }))
      .reverse();
 
    const testPrints = plotRows
      .filter(r => r.type === PLOT_TYPE.TEST_PRINT && !r.isVoid)
      .map(r => ({
        rollId:      r.rollId,
        lengthUsed:  r.lengthFt,
        operator:    r.operator,
        plottingLink: r.remarks.notes || '',
        dateUsed:    r.date ? prism_fmtShort_(new Date(r.date)) : ''
      }))
      .reverse();
 
    return { success: true, damages, testPrints };
  } catch(e) {
    return { success: false, message: e.message };
  }
}