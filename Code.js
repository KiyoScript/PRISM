// ============================================================
//  PRISM — Print Roll Inventory & Substrate Management
//  BACKEND MODULE (PRISMCode.gs)  v2.0
//  Standalone WebApp — PRISM Database spreadsheet
// ============================================================

// ============================================================
//  SHEET NAMES
// ============================================================
const PRISM_SHEETS = {
  JOB_ORDERS:    'JobOrders',
  LFP_MATERIALS: 'LFP_Materials',
  LFP_ROLLS:     'LFP_Rolls',
  LFP_USAGE:     'LFP_Usage',
  PLOTTING_LOG:  'Plotting_Log',
  ROLES:         'Role and Permission',
  AUDIT:         'PRISM_AuditLog',
  SETTINGS:      'PRISM_Settings'
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
  PLOT_ID: 0, JO_NUMBER: 1, PLOTTING_LINK: 2,
  OPERATOR: 3, DATE_PLOTTED: 4, REMARKS: 5
};

// ============================================================
//  STATUS CONSTANTS
// ============================================================
const JO_STATUS = {
  FOR_PLOTTING:   'FOR_PLOTTING',
  READY_TO_PRINT: 'READY_TO_PRINT',
  PRINTING:       'PRINTING',
  COMPLETED:      'COMPLETED'
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
    case 'm':  return Math.round((v * 3.28084) * 10000) / 10000;
    default:   return Math.round(v * 10000) / 10000;
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
    { name: PRISM_SHEETS.JOB_ORDERS,    headers: ['JO_Number','Customer','JobDescription','Category','Width','Height','Quantity','Unit','PlottingLink','Status','RollID','CreatedBy','DateCreated'] },
    { name: PRISM_SHEETS.LFP_MATERIALS, headers: ['MaterialCode','MaterialName','Width','StandardLength','Supplier','CostPerRoll'] },
    { name: PRISM_SHEETS.LFP_ROLLS,     headers: ['RollID','MaterialCode','Width','OriginalLength','RemainingLength','Status','DateReceived','DateOpened','OpenedBy'] },
    { name: PRISM_SHEETS.LFP_USAGE,     headers: ['UsageID','JO_Number','RollID','WidthUsed','LengthUsed','Operator','PlottingLink','DateUsed'] },
    { name: PRISM_SHEETS.PLOTTING_LOG,  headers: ['PlotID','JO_Number','PlottingSheetLink','Operator','DatePlotted','Remarks'] },
    { name: PRISM_SHEETS.AUDIT,         headers: ['DateTime','Action','User','Role','PayloadJSON'] },
    { name: PRISM_SHEETS.SETTINGS,      headers: ['SettingKey','SettingValue'] }
  ];
  schemas.forEach(s => {
    if (!ss.getSheetByName(s.name)) {
      const sh = ss.insertSheet(s.name);
      sh.getRange(1,1,1,s.headers.length).setValues([s.headers]);
      sh.getRange(1,1,1,s.headers.length).setFontWeight('bold').setBackground('#f8fafc');
      sh.setFrozenRows(1);
    }
  });
  // Seed settings
  const set = ss.getSheetByName(PRISM_SHEETS.SETTINGS);
  if (set && set.getLastRow() < 2) {
    set.getRange(2,1,2,2).setValues([['near_empty_threshold_ft','30'],['auto_consume_on_zero','true']]);
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
    return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
  } catch(e) { return ''; }
}
function prism_fmtShort_(val) {
  if (!val) return '';
  try {
    const d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return '';
    const mo = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return mo[d.getMonth()] + ' ' + String(d.getDate()).padStart(2,'0') + ', ' + d.getFullYear();
  } catch(e) { return ''; }
}

// ============================================================
//  ROLE & PERMISSION
// ============================================================
function prism_getUserInfo_() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase();
    const sh    = prism_sh_(PRISM_SHEETS.ROLES);
    const data  = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const role      = String(data[i][0] || '').trim();
      const emails    = String(data[i][1] || '').replace(/"/g,'').toLowerCase().split(',').map(e=>e.trim()).filter(Boolean);
      const abilities = String(data[i][2] || '').replace(/"/g,'').toLowerCase().split(',').map(a=>a.trim()).filter(Boolean);
          const latestPlotAssets = prism_getLatestPlotAssetsByJO_();
      if (emails.includes(email)) return { email, role, abilities };
    }
    return { email, role: 'No Role', abilities: [] };
  } catch(e) { return { email: Session.getActiveUser().getEmail(), role: 'No Role', abilities: [] }; }
}
function prism_getUserInfoPublic()  { return prism_getUserInfo_(); }
function prism_isAdmin_(r)          { return r.toLowerCase().includes('admin'); }
function prism_isATL_(r)            { return r.toLowerCase().includes('admin team leader') || r.toLowerCase().includes('admin tl'); }
function prism_isSTL_(r)            { return r.toLowerCase().includes('senior team leader') || r.toLowerCase().includes('stl'); }
function prism_isOperator_(r)       { return r.toLowerCase().includes('digital operator') || r.toLowerCase().includes('operator'); }

// ============================================================
//  AUDIT LOG
// ============================================================
function prism_audit_(action, payload) {
  try {
    const sh   = prism_sh_(PRISM_SHEETS.AUDIT);
    const user = prism_getUserInfo_();
    sh.insertRowBefore(2);
    sh.getRange(2,1,1,5).setValues([[new Date(), action, user.email, user.role, JSON.stringify(payload)]]);
  } catch(e) { Logger.log('audit error: ' + e.message); }
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
    const sh   = prism_sh_(PRISM_SHEETS.SETTINGS);
    const lr   = sh.getLastRow();
    const out  = {};
    if (lr < 1) return out;
 
    const data = sh.getRange(1, 1, lr, 2).getValues();
    data.forEach(function(row) {
      const key = String(row[0]).trim();
      const val = String(row[1]).trim();
      if (key) out[key] = val;
    });
    return out;
  } catch(e) {
    return {};
  }
}

// ── Private helper: upsert a single key in PRISM_Settings ──────
function prism_setSetting_(key, value) {
  const sh  = prism_sh_(PRISM_SHEETS.SETTINGS);
  const lr  = sh.getLastRow();
 
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
 
    const nearEmpty  = parseFloat(payload.nearEmptyThreshold);
    const lowStock   = parseFloat(payload.lowStockThreshold);
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
    prism_setSetting_('low_stock_threshold_ft',  lowStock);
    prism_setSetting_('auto_consume_on_zero',     autoConsume === true || autoConsume === 'true' ? 'true' : 'false');
 
    prism_audit_('PRISM_UPDATE_SETTINGS', {
      nearEmptyThreshold: nearEmpty,
      lowStockThreshold:  lowStock,
      autoConsumeOnZero:  autoConsume,
      by: user.email
    });
 
    return {
      success: true,
      message: `Settings saved. Near Empty: ${nearEmpty} ft, Low Stock: ${lowStock} ft.`
    };
 
  } catch(e) {
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
      success:   true,
      user:      prism_getUserInfo_(),
      rolls:     prism_getAllRolls_(),
      materials: prism_getAllMaterials_(),
      settings:  prism_getSettings_()
    };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
//  LFP_MATERIALS
// ============================================================
function prism_getAllMaterials_() {
  const sh = prism_sh_(PRISM_SHEETS.LFP_MATERIALS);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2,1,lr-1,6).getValues()
    .filter(r => r[MAT_COL.MATERIAL_CODE] && String(r[MAT_COL.MATERIAL_CODE]).trim())
    .map((r,i) => ({
      rowIndex:       i+2,
      materialCode:   String(r[MAT_COL.MATERIAL_CODE]).trim(),
      materialName:   String(r[MAT_COL.MATERIAL_NAME]).trim(),
      width:          parseFloat(r[MAT_COL.WIDTH])||0,
      standardLength: parseFloat(r[MAT_COL.STANDARD_LENGTH])||0,
      supplier:       String(r[MAT_COL.SUPPLIER]||'').trim(),
      costPerRoll:    parseFloat(r[MAT_COL.COST_PER_ROLL])||0
    }));
}
function prism_getAllMaterialsPublic() {
  try { return { success:true, data: prism_getAllMaterials_() }; }
  catch(e) { return { success:false, message:e.message }; }
}
function prism_addMaterial(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isATL_(user.role))
      return { success:false, message:'Admin or Admin Team Leader only.' };
    if (!payload.materialCode || !payload.materialName)
      return { success:false, message:'Material Code and Name are required.' };
    if (!payload.width || payload.width<=0) return { success:false, message:'Width required.' };
    if (!payload.standardLength || payload.standardLength<=0) return { success:false, message:'Standard Length required.' };
    const existing = prism_getAllMaterials_();
    if (existing.some(m=>m.materialCode.toLowerCase()===payload.materialCode.trim().toLowerCase()))
      return { success:false, message:'Material Code "'+payload.materialCode+'" already exists.' };
    const sh = prism_sh_(PRISM_SHEETS.LFP_MATERIALS);
    sh.getRange(sh.getLastRow()+1,1,1,6).setValues([[
      payload.materialCode.trim().toUpperCase(), payload.materialName.trim(),
      parseFloat(payload.width), parseFloat(payload.standardLength),
      payload.supplier||'', parseFloat(payload.costPerRoll)||0
    ]]);
    prism_audit_('PRISM_ADD_MATERIAL',{materialCode:payload.materialCode,by:user.email});
    return { success:true, message:'Material "'+payload.materialCode+'" added.' };
  } catch(e) { return { success:false, message:e.message }; }
}

// ============================================================
//  LFP_ROLLS
// ============================================================
function prism_getAllRolls_() {
  const sh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  const matMap = {};
  try { prism_getAllMaterials_().forEach(m => { matMap[m.materialCode] = m.materialName || ''; }); } catch(_) {}
  return sh.getRange(2,1,lr-1,9).getValues()
    .filter(r => r[ROLL_COL.ROLL_ID] && String(r[ROLL_COL.ROLL_ID]).trim())
    .map((r,i) => {
      const matCode = String(r[ROLL_COL.MATERIAL_CODE]).trim();
      return {
        rowIndex:        i+2,
        rollId:          String(r[ROLL_COL.ROLL_ID]).trim(),
        materialCode:    matCode,
        materialName:    matMap[matCode] || '',
        width:           parseFloat(r[ROLL_COL.WIDTH])||0,
        originalLength:  parseFloat(r[ROLL_COL.ORIGINAL_LENGTH])||0,
        remainingLength: parseFloat(r[ROLL_COL.REMAINING_LENGTH])||0,
        status:          String(r[ROLL_COL.STATUS]||'UNOPENED').trim().toUpperCase(),
        dateReceived:    prism_fmtShort_(r[ROLL_COL.DATE_RECEIVED]),
        dateOpened:      prism_fmtShort_(r[ROLL_COL.DATE_OPENED]),
        openedBy:        String(r[ROLL_COL.OPENED_BY]||'').trim()
      };
    });
}
function prism_getAllRollsPublic() {
  try { return { success:true, rolls:prism_getAllRolls_(), settings:prism_getSettings_() }; }
  catch(e) { return { success:false, message:e.message }; }
}

// ============================================================
//  STOCK IN — Admin Team Leader
// ============================================================
function prism_stockIn(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isATL_(user.role))
      return { success:false, message:'Admin Team Leader only.' };
    if (!payload.materialCode)            return { success:false, message:'Material Code required.' };
    if (!payload.qty || payload.qty<1)    return { success:false, message:'Qty must be ≥ 1.' };
    if (!payload.width || payload.width<=0) return { success:false, message:'Width required.' };
    if (!payload.originalLength || payload.originalLength<=0) return { success:false, message:'Length required.' };

    const sh       = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const lr       = sh.getLastRow();
    const existing = lr>=2 ? sh.getRange(2,1,lr-1,9).getValues() : [];
    const rollIds  = prism_generateRollIds_(payload.materialCode, parseInt(payload.qty), existing);
    const today    = new Date();
    const rows     = rollIds.map(id => [
      id, payload.materialCode.trim().toUpperCase(),
      parseFloat(payload.width), parseFloat(payload.originalLength), parseFloat(payload.originalLength),
      ROLL_STATUS.UNOPENED, today, '', ''
    ]);
    sh.getRange(sh.getLastRow()+1,1,rows.length,9).setValues(rows);
    prism_audit_('PRISM_STOCK_IN',{materialCode:payload.materialCode,qty:payload.qty,rollIds,by:user.email});
    return { success:true, message:rollIds.length+' roll(s) stocked in: '+rollIds.join(', '), rollIds };
  } catch(e) { return { success:false, message:e.message }; }
}

// ============================================================
//  OPEN ROLL — Senior Team Leader
// ============================================================
function prism_openRoll(rollId) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isSTL_(user.role))
      return { success:false, message:'Senior Team Leader only.' };
    const sh  = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const lr  = sh.getLastRow();
    if (lr<2) return { success:false, message:'No rolls found.' };
    const data = sh.getRange(2,1,lr-1,9).getValues();
    let rowIdx=-1;
    data.forEach((r,i)=>{ if(String(r[ROLL_COL.ROLL_ID]).trim()===rollId) rowIdx=i+2; });
    if (rowIdx===-1) return { success:false, message:'Roll "'+rollId+'" not found.' };
    const row    = sh.getRange(rowIdx,1,1,9).getValues()[0];
    const status = String(row[ROLL_COL.STATUS]).trim().toUpperCase();
    if (status===ROLL_STATUS.OPEN)     return { success:false, message:'Roll is already OPEN.' };
    if (status===ROLL_STATUS.CONSUMED) return { success:false, message:'Roll is CONSUMED.' };
    const today = new Date();
    sh.getRange(rowIdx, ROLL_COL.STATUS+1).setValue(ROLL_STATUS.OPEN);
    sh.getRange(rowIdx, ROLL_COL.DATE_OPENED+1).setValue(today);
    sh.getRange(rowIdx, ROLL_COL.OPENED_BY+1).setValue(user.email);
    prism_audit_('PRISM_OPEN_ROLL',{rollId,by:user.email});
    return { success:true, message:'Roll '+rollId+' is now OPEN.' };
  } catch(e) { return { success:false, message:e.message }; }
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
    const rollSh   = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollLr   = rollSh.getLastRow();
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
    const jobWidth  = parseFloat(payload.widthUsed) || 0;
    if (jobWidth > 0 && jobWidth > rollWidth)
      return { success: false, message: `Job width (${jobWidth} ft) exceeds roll width (${rollWidth} ft).` };

    // ── Length validation ──
    const remaining   = parseFloat(rollRow[ROLL_COL.REMAINING_LENGTH]) || 0;
    const lengthUsed  = parseFloat(payload.lengthUsed);
    if (lengthUsed > remaining)
      return { success: false, message: `Length used (${lengthUsed} ft) exceeds remaining (${remaining} ft).` };

    // ── Deduct from roll ──
    const settings     = prism_getSettings_();
    const newRemaining = Math.max(0, remaining - lengthUsed);
    const newRollStatus = (newRemaining === 0 && settings.auto_consume_on_zero === 'true')
      ? ROLL_STATUS.CONSUMED : ROLL_STATUS.OPEN;

    rollSh.getRange(rollIdx, ROLL_COL.REMAINING_LENGTH + 1).setValue(newRemaining);
    rollSh.getRange(rollIdx, ROLL_COL.STATUS + 1).setValue(newRollStatus);

    // ── Write usage entry ──
    const today   = new Date();
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
      rollId:       payload.rollId,
      joNumber:     payload.joNumber,
      lengthUsed,
      newRemaining,
      newRollStatus,
      by:           user.email
    });

    return {
      success:      true,
      message:      `Usage recorded. ${payload.rollId}: ${newRemaining} ft remaining.`
                    + (newRollStatus === ROLL_STATUS.CONSUMED ? ' Roll CONSUMED.' : ''),
      newRemaining,
      newRollStatus
    };

  } catch(e) { return { success: false, message: e.message }; }
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
    const rollSh   = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollLr   = rollSh.getLastRow();
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
    const remaining  = parseFloat(rollRow[ROLL_COL.REMAINING_LENGTH]) || 0;
    const lengthUsed = parseFloat(payload.lengthUsed);
    if (lengthUsed > remaining)
      return { success: false, message: `Length used (${lengthUsed} ft) exceeds remaining (${remaining} ft).` };
 
    // ── Deduct from roll ──
    const settings     = prism_getSettings_();
    const newRemaining = Math.max(0, remaining - lengthUsed);
    const newRollStatus = (newRemaining === 0 && settings.auto_consume_on_zero === 'true')
      ? ROLL_STATUS.CONSUMED : ROLL_STATUS.OPEN;
 
    rollSh.getRange(rollIdx, ROLL_COL.REMAINING_LENGTH + 1).setValue(newRemaining);
    rollSh.getRange(rollIdx, ROLL_COL.STATUS + 1).setValue(newRollStatus);
 
    // ── Write to LFP_Usage with TEST-PRINT marker ──
    // notes stored in the PlottingLink column (col G / index 6)
    const today   = new Date();
    const usageSh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    const rollWidth = parseFloat(rollRow[ROLL_COL.WIDTH]) || 0;
    usageSh.getRange(usageSh.getLastRow() + 1, 1, 1, 8).setValues([[
      'TEST-' + today.getTime(),
      TEST_PRINT_MARKER,           // JO_Number = 'TEST-PRINT'
      payload.rollId,
      rollWidth,                   // widthUsed = full roll width
      lengthUsed,
      user.email,
      payload.notes || '',         // notes in PlottingLink column
      today
    ]]);
 
    prism_audit_('PRISM_TEST_PRINT', {
      rollId:       payload.rollId,
      lengthUsed,
      newRemaining,
      newRollStatus,
      notes:        payload.notes || '',
      by:           user.email
    });
 
    return {
      success:      true,
      message:      `Test print recorded. ${payload.rollId}: ${newRemaining} ft remaining.`
                    + (newRollStatus === ROLL_STATUS.CONSUMED ? ' Roll CONSUMED.' : ''),
      newRemaining,
      newRollStatus
    };
 
  } catch(e) { return { success: false, message: e.message }; }
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
  } catch(e) {
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
    const rollSh   = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollLr   = rollSh.getLastRow();
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
      
    const remaining  = parseFloat(rollRow[ROLL_COL.REMAINING_LENGTH]) || 0;
    const lengthUsed = parseFloat(payload.lengthUsed);
      
    const rollWidth = parseFloat(rollRow[ROLL_COL.WIDTH]) || 0;
    const originalLength = parseFloat(rollRow[ROLL_COL.ORIGINAL_LENGTH]) || 0;
    const today   = new Date();

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
    
    const settings     = prism_getSettings_();
    const newRemaining = Math.max(0, remaining + refundLength - lengthUsed);
    const newRollStatus = (newRemaining === 0 && settings.auto_consume_on_zero === 'true')
      ? ROLL_STATUS.CONSUMED : ROLL_STATUS.OPEN;
      
    rollSh.getRange(rollIdx, ROLL_COL.REMAINING_LENGTH + 1).setValue(newRemaining);
    rollSh.getRange(rollIdx, ROLL_COL.STATUS + 1).setValue(newRollStatus);
    
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
    
    // ── Find and Void previous aborted plot in Plotting_Log ──
    const plotSh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotLr = plotSh.getLastRow();
    if (plotLr >= 2) {
      const plotData = plotSh.getRange(2, 6, plotLr - 1, 1).getValues();
      for (let i = plotData.length - 1; i >= 0; i--) {
        const jStr = String(plotData[i][0]).trim();
        if (jStr.startsWith('{')) {
          try {
            const obj = JSON.parse(jStr);
            if (obj.rollId === payload.rollId && !obj.isDamage && !obj.isVoid && obj.type === 'ROLL_PLAN') {
              const objJos = obj.joNumbers || [];
              if (targets.some(t => objJos.includes(t))) {
                obj.isVoid = true;
                obj.voidReason = 'Replaced by Damage ' + lengthUsed + 'ft';
                plotSh.getRange(i + 2, 6).setValue(JSON.stringify(obj));
                break;
              }
            }
          } catch(e) {}
        }
      }
    }
    
    // ── Write to Plotting_Log for visual map ──
    const plotId = 'PLT-DAM-' + today.getTime();
    
    const startAtFt = Math.max(0, originalLength - remaining);
    const endAtFt = startAtFt + lengthUsed;
    
    const remarksObj = {
      type: "ROLL_PLAN",
      isDamage: true,
      rollId: payload.rollId,
      rollWidth: rollWidth,
      originalLength: originalLength,
      startAtFt: startAtFt,
      endAtFt: endAtFt,
      lengthUsed: lengthUsed,
      joNumbers: payload.joNumbers || [],
      createdBy: user.email,
      damageReason: payload.remarks || ''
    };
    
    plotSh.getRange(plotSh.getLastRow() + 1, 1, 1, 6).setValues([[
      plotId, 
      'DAMAGE', 
      '', // No link for damage
      user.email, 
      today, 
      JSON.stringify(remarksObj)
    ]]);
    
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
      rollId:       payload.rollId,
      lengthUsed,
      newRemaining,
      newRollStatus,
      affectedJOs:  payload.joNumbers || [],
      remarks:      payload.remarks || '',
      by:           user.email
    });
    
    return {
      success: true,
      message: `Declared ${lengthUsed}ft damage. ` + (affectedCount > 0 ? `${affectedCount} JO(s) reverted to FOR_PLOTTING.` : ''),
      newRemaining,
      newRollStatus
    };
    
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
//  JOB ORDERS
// ============================================================
function prism_getAllJobOrders_() {
  const latestPlotAssets = prism_getLatestPlotAssetsByJO_();
  const sh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
  const lr = sh.getLastRow();
  if (lr<2) return [];
  return sh.getRange(2,1,lr-1,13).getValues()
    .filter(r=>r[JO_COL.JO_NUMBER]&&String(r[JO_COL.JO_NUMBER]).trim())
    .map((r,i)=>{
      const joNumber = String(r[JO_COL.JO_NUMBER]).trim();
      const plotAsset = latestPlotAssets[String(joNumber || '').toUpperCase()] || {};
      return {
      rowIndex: i+2,
      joNumber:       joNumber,
      customer:       String(r[JO_COL.CUSTOMER]||'').trim(),
      jobDescription: String(r[JO_COL.JOB_DESCRIPTION]||'').trim(),
      category:       String(r[JO_COL.CATEGORY]||'').trim(),
      width:          parseFloat(r[JO_COL.WIDTH])||0,
      height:         parseFloat(r[JO_COL.HEIGHT])||0,
      quantity:       parseInt(r[JO_COL.QUANTITY])||0,
      unit:           (String(r[JO_COL.UNIT]||'ft').trim().toLowerCase()) || 'ft',
      widthFt:        prism_toFt_(parseFloat(r[JO_COL.WIDTH])||0, String(r[JO_COL.UNIT]||'ft').trim()),
      heightFt:       prism_toFt_(parseFloat(r[JO_COL.HEIGHT])||0, String(r[JO_COL.UNIT]||'ft').trim()),
      plottingLink:   String(r[JO_COL.PLOTTING_LINK]||'').trim(),
      plottingImageUrl: String(plotAsset.pngUrl || r[JO_COL.PLOTTING_LINK] || '').trim(),
      plottingFolderUrl: String(plotAsset.folderUrl || '').trim(),
      plotDateMs:     parseInt(plotAsset.dateMs) || 0,
      status:         String(r[JO_COL.STATUS]||JO_STATUS.FOR_PLOTTING).trim(),
      rollId:         String(r[JO_COL.ROLL_ID]||'').trim(),
      createdBy:      String(r[JO_COL.CREATED_BY]||'').trim(),
      dateCreated:    prism_fmtShort_(r[JO_COL.DATE_CREATED])
    }})
    .sort((a, b) => new Date(b.dateCreated) - new Date(a.dateCreated));
}

  function prism_getLatestPlotAssetsByJO_() {
    const sh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const lr = sh.getLastRow();
    const out = {};
    if (lr < 2) return out;

    const rows = sh.getRange(2, 1, lr - 1, 6).getValues();
    rows.forEach(r => {
      const rawRemarks = String(r[PLOT_COL.REMARKS] || '').trim();
      if (!rawRemarks || rawRemarks.charAt(0) !== '{') return;

      let parsed;
      try { parsed = JSON.parse(rawRemarks); } catch (e) { return; }
      if (!parsed || parsed.type !== 'ROLL_PLAN') return;

      const joNumbers = Array.isArray(parsed.joNumbers) ? parsed.joNumbers : [];
      if (!joNumbers.length) return;

      const dateRaw = r[PLOT_COL.DATE_PLOTTED];
      const dateMs = dateRaw ? new Date(dateRaw).getTime() : 0;
      joNumbers.forEach(jo => {
        const key = String(jo || '').trim().toUpperCase();
        if (!key) return;
        if (!out[key] || dateMs >= (out[key].dateMs || 0)) {
          out[key] = {
            dateMs: dateMs,
            pngUrl: String(parsed.pngUrl || parsed.plottingLink || '').trim(),
            folderUrl: String(parsed.folderUrl || '').trim(),
            rollId: String(parsed.rollId || '').trim()
          };
        }
      });
    });
    return out;
  }
function prism_getJobOrdersPublic() {
  try { return { success:true, data:prism_getAllJobOrders_() }; }
  catch(e) { return { success:false, message:e.message }; }
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
    const all = prism_getAllJobOrders_();
    const usageMap = prism_getUsageSummaryByJO_();
    const data = all
      .filter(j => j.status === JO_STATUS.FOR_PLOTTING || j.status === 'PRINTING')
      .map(j => {
        const u = usageMap[String(j.joNumber || '').trim().toUpperCase()];
        if (!u) return Object.assign({}, j, {
          plottedBefore: false,
          usageCount: 0,
          usageTotalLength: 0,
          usageLastDate: '',
          usageLastRollId: '',
          usageRollIds: []
        });
        return Object.assign({}, j, {
          plottedBefore: u.count > 0,
          usageCount: u.count,
          usageTotalLength: Number((u.totalLength || 0).toFixed(2)),
          usageLastDate: u.lastDateLabel || '',
          usageLastRollId: u.lastRollId || '',
          usageRollIds: Object.keys(u.rollIds)
        });
      });
    return { success: true, data: data };
  } catch(e) { return { success: false, message: e.message }; }
}
function prism_submitJobOrder(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!payload.joNumber||!payload.joNumber.trim()) return { success:false, message:'JO Number required.' };
    if (!payload.customer||!payload.customer.trim())  return { success:false, message:'Customer required.' };
    if (!payload.category)                            return { success:false, message:'Category required.' };
    const existing = prism_getAllJobOrders_();
    if (existing.some(j=>j.joNumber.toLowerCase()===payload.joNumber.trim().toLowerCase()))
      return { success:false, message:'JO "'+payload.joNumber+'" already exists.' };

    const sh    = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const today = new Date();

    const LFP_CATEGORIES = ['banner', 'sticker', 'signage', 'canvas', 'tarpaulin', 'standees/display'];
    const isLFP = LFP_CATEGORIES.includes((payload.category || '').toLowerCase());
    const unit = (payload.unit || 'ft').toString().trim().toLowerCase();
    const status = isLFP ? 'FOR_PLOTTING' : (payload.status || JO_STATUS.FOR_PLOTTING);
 
    sh.getRange(sh.getLastRow()+1,1,1,13).setValues([[
      payload.joNumber.trim().toUpperCase(), payload.customer.trim(),
      payload.jobDescription||'', payload.category,
      parseFloat(payload.width)||0, parseFloat(payload.height)||0, parseInt(payload.quantity)||1,
      unit,                        // column H = Unit (replaced ProductionType)
      payload.plottingLink||'',
      status,
      '', user.email, today
    ]]);

    prism_audit_('PRISM_SUBMIT_JO',{joNumber:payload.joNumber,by:user.email});
    return { success:true, message:'JO "'+payload.joNumber+'" submitted.' };
  } catch(e) { return { success:false, message:e.message }; }
}
function prism_updateJOStatus(joNumber, newStatus) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role) && !prism_isSTL_(user.role))
      return { success: false, message: 'Admin or Senior Team Leader only.' };
    const sh   = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const lr   = sh.getLastRow();
    if (lr<2) return { success:false, message:'No JOs found.' };
    const data = sh.getRange(2,1,lr-1,13).getValues();
    let updated=0;
    data.forEach((r,i)=>{
      if (String(r[JO_COL.JO_NUMBER]).trim().toUpperCase()===joNumber.trim().toUpperCase()) {
        sh.getRange(i+2, JO_COL.STATUS+1).setValue(newStatus);
        updated++;
      }
    });
    if (!updated) return { success:false, message:'JO "'+joNumber+'" not found.' };
    prism_audit_('PRISM_UPDATE_JO_STATUS',{joNumber,newStatus,by:user.email});
    return { success:true, message:'JO "'+joNumber+'" → '+newStatus };
  } catch(e) { return { success:false, message:e.message }; }
}

// ============================================================
//  PLOTTING LOG
// ============================================================
function prism_submitPlottingLog(payload) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role)&&!prism_isOperator_(user.role))
      return { success:false, message:'Access denied.' };
    if (!payload.joNumber)     return { success:false, message:'JO Number required.' };
    if (!payload.plottingLink) return { success:false, message:'Plotting link required.' };
    const today  = new Date();
    const plotId = 'PLT-'+today.getTime();
    const sh     = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    sh.getRange(sh.getLastRow()+1,1,1,6).setValues([[
      plotId, payload.joNumber.trim().toUpperCase(), payload.plottingLink,
      user.email, today, payload.remarks||''
    ]]);
    // Update JO → PLOTTED + save link
    const joSh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const lr   = joSh.getLastRow();
    if (lr>=2) {
      const data = joSh.getRange(2,1,lr-1,13).getValues();
      data.forEach((r,i)=>{
        if (String(r[JO_COL.JO_NUMBER]).trim().toUpperCase()===payload.joNumber.trim().toUpperCase()) {
          joSh.getRange(i+2, JO_COL.PLOTTING_LINK+1).setValue(payload.plottingLink);
          if (String(r[JO_COL.STATUS]).trim()===JO_STATUS.FOR_PLOTTING)
            joSh.getRange(i+2, JO_COL.STATUS+1).setValue(JO_STATUS.PLOTTED);
        }
      });
    }
    prism_audit_('PRISM_SUBMIT_PLOT',{joNumber:payload.joNumber,by:user.email});
    return { success:true, message:'Plotting log saved for '+payload.joNumber, plotId };
  } catch(e) { return { success:false, message:e.message }; }
}
function prism_getPlottingLog() {
  try {
    const sh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const lr = sh.getLastRow();
    if (lr<2) return { success:true, data:[] };
    return { success:true, data: sh.getRange(2,1,lr-1,6).getValues()
      .filter(r=>r[PLOT_COL.PLOT_ID])
      .map(r=>({
        plotId:       String(r[PLOT_COL.PLOT_ID]).trim(),
        joNumber:     String(r[PLOT_COL.JO_NUMBER]).trim(),
        plottingLink: String(r[PLOT_COL.PLOTTING_LINK]).trim(),
        operator:     String(r[PLOT_COL.OPERATOR]).trim(),
        datePlotted:  prism_fmtShort_(r[PLOT_COL.DATE_PLOTTED]),
        remarks:      String(r[PLOT_COL.REMARKS]||'').trim()
      }))
    };
  } catch(e) { return { success:false, message:e.message }; }
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
      .map(x => String(x || '').trim())
      .filter(Boolean);
    const wantedSet = {};
    wanted.forEach(id => { wantedSet[id] = true; });

    const out = {};
    wanted.forEach(id => { out[id] = []; });

    const rollMeta = {};
    const rollSh = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollLr = rollSh.getLastRow();
    if (rollLr >= 2) {
      rollSh.getRange(2, 1, rollLr - 1, 9).getValues().forEach(r => {
        const rollId = String(r[ROLL_COL.ROLL_ID] || '').trim();
        if (!rollId || !wantedSet[rollId]) return;
        rollMeta[rollId] = {
          width: parseFloat(r[ROLL_COL.WIDTH]) || 0,
          originalLength: parseFloat(r[ROLL_COL.ORIGINAL_LENGTH]) || 0,
          remainingLength: parseFloat(r[ROLL_COL.REMAINING_LENGTH]) || 0
        };
      });
    }

    const sh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const lr = sh.getLastRow();
    if (lr >= 2 && wanted.length) {
      const rows = sh.getRange(2, 1, lr - 1, 6).getValues();
      rows.forEach(r => {
        const remarks = String(r[PLOT_COL.REMARKS] || '').trim();
        if (!remarks || remarks.charAt(0) !== '{') return;

        let parsed = null;
        try { parsed = JSON.parse(remarks); } catch (_) { parsed = null; }
        if (!parsed || parsed.type !== 'ROLL_PLAN' || parsed.isVoid) return;

        const rollId = String(parsed.rollId || '').trim();
        if (!rollId || !wantedSet[rollId]) return;

        out[rollId].push({
          plotId: String(r[PLOT_COL.PLOT_ID] || '').trim(),
          rollId: rollId,
          rollWidth: parseFloat(parsed.rollWidth) || 0,
          originalLength: parseFloat(parsed.originalLength) || 0,
          startAtFt: parseFloat(parsed.startAtFt) || 0,
          endAtFt: parseFloat(parsed.endAtFt) || 0,
          lengthUsed: parseFloat(parsed.lengthUsed) || 0,
          joNumbers: Array.isArray(parsed.joNumbers) ? parsed.joNumbers : [],
          createdBy: String(parsed.createdBy || r[PLOT_COL.OPERATOR] || '').trim(),
          datePlotted: prism_fmtShort_(r[PLOT_COL.DATE_PLOTTED]),
          rows: Array.isArray(parsed.rows) ? parsed.rows : [],
          isDamage: !!parsed.isDamage,
          source: 'ROLL_PLAN'
        });
      });
    }

    // Remove the `firstStart > 0.05` constraint. Always parse USAGE_FALLBACK to fill gaps throughout the roll.
    const fallbackNeeded = wanted;

    if (fallbackNeeded.length) {
      const fallbackSet = {};
      fallbackNeeded.forEach(id => { fallbackSet[id] = true; });
      const usageSh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
      const usageLr = usageSh.getLastRow();
      if (usageLr >= 2) {
        const grouped = {};
        usageSh.getRange(2, 1, usageLr - 1, 8).getValues().forEach(r => {
          const rollId = String(r[USAGE_COL.ROLL_ID] || '').trim();
          if (!rollId || !fallbackSet[rollId]) return;
          const joNumber = String(r[USAGE_COL.JO_NUMBER] || '').trim().toUpperCase();
          const lengthUsed = parseFloat(r[USAGE_COL.LENGTH_USED]) || 0;
          if (!joNumber || lengthUsed <= 0) return;
          if (!grouped[rollId]) grouped[rollId] = [];
          grouped[rollId].push({
            usageId: String(r[USAGE_COL.USAGE_ID] || '').trim(),
            joNumber: joNumber,
            widthUsed: parseFloat(r[USAGE_COL.WIDTH_USED]) || 0,
            lengthUsed: lengthUsed,
            operator: String(r[USAGE_COL.OPERATOR] || '').trim(),
            dateRaw: r[USAGE_COL.DATE_USED],
            dateLabel: prism_fmtShort_(r[USAGE_COL.DATE_USED])
          });
        });

        Object.keys(grouped).forEach(rollId => {
          grouped[rollId].sort((a, b) => {
            const ad = a.dateRaw ? new Date(a.dateRaw).getTime() : 0;
            const bd = b.dateRaw ? new Date(b.dateRaw).getTime() : 0;
            if (ad !== bd) return ad - bd;
            return String(a.usageId).localeCompare(String(b.usageId));
          });

          // Sort out[rollId] so we can properly check for gaps between known plotted segments.
          out[rollId].sort((a, b) => (a.startAtFt || 0) - (b.startAtFt || 0));

          let cursorFt = 0;
          grouped[rollId].forEach(item => {
            const nextEnd = cursorFt + item.lengthUsed;

            // Check if this usage overlaps with an already-known visual snapshot in out[rollId]
            let isCovered = false;
            for (let s = 0; s < out[rollId].length; s++) {
              const seg = out[rollId][s];
              // If this usage fits entirely inside or mostly overlaps with a Plotting_Log snapshot
              if (Math.abs(cursorFt - seg.startAtFt) < 0.1 || (cursorFt >= seg.startAtFt && nextEnd <= seg.endAtFt + 0.1)) {
                isCovered = true;
                break;
              }
            }

            if (!isCovered) {
              out[rollId].push({
                plotId: item.usageId,
                rollId: rollId,
                rollWidth: (rollMeta[rollId] && rollMeta[rollId].width) || 0,
                originalLength: (rollMeta[rollId] && rollMeta[rollId].originalLength) || 0,
                startAtFt: cursorFt,
                endAtFt: nextEnd,
                lengthUsed: item.lengthUsed,
                joNumbers: [item.joNumber],
                createdBy: item.operator,
                datePlotted: item.dateLabel,
                notes: item.notes || '',
                isTestPrint: item.joNumber === TEST_PRINT_MARKER,
                rows: [],
                source: 'USAGE_FALLBACK'
              });
            }
            cursorFt = nextEnd;
          });
        });
      }
    }

    Object.keys(out).forEach(rollId => {
      out[rollId].sort((a, b) => (a.startAtFt || 0) - (b.startAtFt || 0));
    });

    return { success: true, data: out };
  } catch (e) {
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
    if (lr<2) return { success:true, data:[] };
    return { success:true, data: sh.getRange(2,1,lr-1,8).getValues()
      .filter(r=>r[USAGE_COL.USAGE_ID])
      .map(r=>({
        usageId:      String(r[USAGE_COL.USAGE_ID]).trim(),
        joNumber:     String(r[USAGE_COL.JO_NUMBER]).trim(),
        rollId:       String(r[USAGE_COL.ROLL_ID]).trim(),
        widthUsed:    parseFloat(r[USAGE_COL.WIDTH_USED])||0,
        lengthUsed:   parseFloat(r[USAGE_COL.LENGTH_USED])||0,
        operator:     String(r[USAGE_COL.OPERATOR]).trim(),
        plottingLink: String(r[USAGE_COL.PLOTTING_LINK]||'').trim(),
        dateUsed:     prism_fmtShort_(r[USAGE_COL.DATE_USED])
      }))
    };
  } catch(e) { return { success:false, message:e.message }; }
}

// ============================================================
//  ADMIN: Manual roll status override
// ============================================================
function prism_setRollStatus(rollId, newStatus) {
  try {
    const user = prism_getUserInfo_();
    if (!prism_isAdmin_(user.role)) return { success:false, message:'Admins only.' };
    const valid = Object.values(ROLL_STATUS);
    if (!valid.includes(newStatus.toUpperCase())) return { success:false, message:'Invalid status.' };
    const sh  = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const lr  = sh.getLastRow();
    const data = sh.getRange(2,1,lr-1,1).getValues();
    let rowIdx=-1;
    data.forEach((r,i)=>{ if(String(r[0]).trim()===rollId) rowIdx=i+2; });
    if (rowIdx===-1) return { success:false, message:'Roll not found.' };
    sh.getRange(rowIdx, ROLL_COL.STATUS+1).setValue(newStatus.toUpperCase());
    prism_audit_('PRISM_SET_ROLL_STATUS',{rollId,newStatus,by:user.email});
    return { success:true, message:rollId+' → '+newStatus };
  } catch(e) { return { success:false, message:e.message }; }
}

function prism_getStockMaterialsList() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const linkSheet = ss.getSheetByName('DatabaseLink');
    if (!linkSheet) throw new Error('DatabaseLink sheet not found');

    const rows  = linkSheet.getRange(2, 1, linkSheet.getLastRow() - 1, 2).getValues();
    const match = rows.find(r => r[0].toString().trim() === 'StockDatabase');
    if (!match) throw new Error('StockDatabase not found in DatabaseLink');

    const idMatch = match[1].toString().match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (!idMatch) throw new Error('Invalid StockDatabase URL');

    const sheet = SpreadsheetApp.openById(idMatch[1]).getSheetByName('AllItems');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
      .filter(r => r[0] && r[1])
      .map(r => ({
        itemCode:    String(r[0]).trim(),
        itemDesc:    String(r[1]).trim(),
        stockOnHand: r[6] !== '' ? Number(r[6]) : 0,
        unitCost:    r[7] !== '' ? Number(r[7]) : 0
      }));

    return { success: true, data };
  } catch(e) {
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
    const rollSh   = prism_sh_(PRISM_SHEETS.LFP_ROLLS);
    const rollLr   = rollSh.getLastRow();
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

    const settings    = prism_getSettings_();
    const rollSnapshots = {};
    Object.keys(rollUsageById).forEach(rollId => {
      const entry = rollMap[rollId];
      const remaining = parseFloat(entry.row[ROLL_COL.REMAINING_LENGTH]) || 0;
      const rollWidth = parseFloat(entry.row[ROLL_COL.WIDTH]) || 0;
      const originalLength = parseFloat(entry.row[ROLL_COL.ORIGINAL_LENGTH]) || 0;
      const newRemaining = Math.max(0, remaining - rollUsageById[rollId]);
      const newRollStatus = (newRemaining === 0 && settings.auto_consume_on_zero === 'true')
        ? ROLL_STATUS.CONSUMED : ROLL_STATUS.OPEN;

      rollSh.getRange(entry.rowIdx, ROLL_COL.REMAINING_LENGTH + 1).setValue(newRemaining);
      rollSh.getRange(entry.rowIdx, ROLL_COL.STATUS + 1).setValue(newRollStatus);

      rollSnapshots[rollId] = {
        width: rollWidth,
        originalLength: originalLength,
        startAtFt: Math.max(0, originalLength - remaining),
        endAtFt: Math.max(0, originalLength - newRemaining),
        newRemaining: newRemaining,
        newStatus: newRollStatus
      };
    });

    // ── Write usage by JO and by roll ──
    const today   = new Date();
    const usageSh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    const joSh    = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
    const joLr    = joSh.getLastRow();
    const joData  = joLr >= 2 ? joSh.getRange(2, 1, joLr - 1, 13).getValues() : [];

    // Calculate per-JO length used from all layout rows and each roll plan.
    const joLengths = {};
    const joRollLengths = {};

    rollPlans.forEach(plan => {
      const planRows = Array.isArray(plan.rows) ? plan.rows : [];
      planRows.forEach(row => {
        const rowJOs = {};
        (row.pieces || []).forEach(p => { rowJOs[String(p.joNumber || '').trim().toUpperCase()] = true; });
        Object.keys(rowJOs).forEach(jo => {
          if (!jo) return;
          const rowH = parseFloat(row.rowH || 0) || 0;
          joLengths[jo] = (joLengths[jo] || 0) + rowH;
          if (!joRollLengths[jo]) joRollLengths[jo] = {};
          joRollLengths[jo][plan.rollId] = (joRollLengths[jo][plan.rollId] || 0) + rowH;
        });
      });
    });

    const totalLengthUsed = rollPlans.reduce((s, p) => s + (parseFloat(p.lengthUsed) || 0), 0);
    if (!Object.keys(joLengths).length) {
      const perJO = payload.joNumbers.length ? (totalLengthUsed / payload.joNumbers.length) : 0;
      payload.joNumbers.forEach(jo => {
        const key = String(jo || '').trim().toUpperCase();
        if (!key) return;
        joLengths[key] = perJO;
        joRollLengths[key] = joRollLengths[key] || {};
        joRollLengths[key][rollPlans[0].rollId] = perJO;
      });
    }

    payload.joNumbers.forEach(joRaw => {
      const joNumber = String(joRaw || '').trim().toUpperCase();
      if (!joNumber) return;

      // Write one usage row per JO per roll where this JO consumed length.
      // Update JO: append all rollIds + set READY_TO_PRINT.
      joData.forEach((r, i) => {
        if (String(r[JO_COL.JO_NUMBER]).trim().toUpperCase() === joNumber) {
          const existingRolls = String(r[JO_COL.ROLL_ID] || '')
            .split(',').map(x => x.trim()).filter(Boolean);
          Object.keys(joRollLengths[joNumber] || {}).forEach(rollId => {
            if (rollId && !existingRolls.includes(rollId)) existingRolls.push(rollId);
          });
          joSh.getRange(i + 2, JO_COL.ROLL_ID + 1).setValue(existingRolls.join(', '));
          joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.READY_TO_PRINT);
          if (effectivePlottingLink) {
            joSh.getRange(i + 2, JO_COL.PLOTTING_LINK + 1).setValue(effectivePlottingLink);
          }
        }
      });
    });

    // Persist roll layout snapshots so future plotting on the same roll can show previous map.
    const plotSh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    rollPlans.forEach(plan => {
      const rollId = plan.rollId;
      const snap = rollSnapshots[rollId] || {};
      const joSet = {};
      (Array.isArray(plan.rows) ? plan.rows : []).forEach(r => {
        (Array.isArray(r.pieces) ? r.pieces : []).forEach(p => {
          const jo = String(p.joNumber || '').trim().toUpperCase();
          if (jo) joSet[jo] = true;
        });
      });

      const remarksObj = {
        type: 'ROLL_PLAN',
        status: 'PLANNED',
        rollId: rollId,
        rollWidth: snap.width || 0,
        originalLength: snap.originalLength || 0,
        startAtFt: 0,
        endAtFt: parseFloat(plan.lengthUsed) || 0,
        lengthUsed: parseFloat(plan.lengthUsed) || 0,
        joNumbers: Object.keys(joSet),
        plottingLink: (savedPlotsByRollId[rollId] && savedPlotsByRollId[rollId].pngUrl) || effectivePlottingLink || '',
        pngUrl: (savedPlotsByRollId[rollId] && savedPlotsByRollId[rollId].pngUrl) || '',
        folderUrl: (savedPlotsByRollId[rollId] && savedPlotsByRollId[rollId].folderUrl) || '',
        createdBy: user.email,
        rows: prism_compactRollRows_(plan.rows)
      };

      plotSh.getRange(plotSh.getLastRow() + 1, 1, 1, 6).setValues([[
        'PLN-' + today.getTime() + '-' + rollId,
        Object.keys(joSet).slice(0, 8).join(', '),
        (savedPlotsByRollId[rollId] && savedPlotsByRollId[rollId].pngUrl) || effectivePlottingLink || '',
        user.email,
        today,
        JSON.stringify(remarksObj)
      ]]);
    });

    prism_audit_('PRISM_CONFIRM_PLOT_LAYOUT', {
      rollPlans: rollPlans.map(p => ({ rollId: p.rollId, lengthUsed: p.lengthUsed })),
      totalLengthUsed: Number(totalLengthUsed.toFixed(3)),
      joNumbers: payload.joNumbers,
      by: user.email
    });

    const consumedRolls = Object.keys(rollSnapshots).filter(id => {
      return rollSnapshots[id].newStatus === ROLL_STATUS.CONSUMED;
    });

    return {
      success: true,
      message: 'Layout confirmed! ' + payload.joNumbers.length
        + ' JO(s) -> READY_TO_PRINT. Total roll used: ' + totalLengthUsed.toFixed(1) + 'ft across '
        + rollPlans.length + ' roll plan(s).'
        + (consumedRolls.length ? (' Consumed: ' + consumedRolls.join(', ') + '.') : ''),
      plottingLink: effectivePlottingLink || '',
      savedPlots: Object.keys(savedPlotsByRollId).map(rollId => ({
        rollId: rollId,
        fileBaseName: savedPlotsByRollId[rollId].fileBaseName,
        pngUrl: savedPlotsByRollId[rollId].pngUrl,
        folderUrl: savedPlotsByRollId[rollId].folderUrl
      }))
    };

  } catch(e) { return { success: false, message: e.message }; }
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

    const plotSh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotLr = plotSh.getLastRow();
    if (plotLr < 2) return { success: false, message: 'Plot log empty.' };
    const plotData = plotSh.getRange(2, 1, plotLr - 1, 6).getValues();

    let targetRowIdx = -1;
    let targetJson = null;
    let targetRollId = '';
    let targetLength = 0;

    for (let i = 0; i < plotData.length; i++) {
      if (String(plotData[i][0]).trim() === plotId) {
        targetRowIdx = i + 2;
        try { targetJson = JSON.parse(String(plotData[i][5] || '{}')); } catch(e){}
        if (!targetJson || targetJson.status !== 'PLANNED' || targetJson.isVoid) {
          return { success: false, message: 'Layout is not currently PLANNED or has been voided.' };
        }
        targetRollId = targetJson.rollId;
        targetLength = parseFloat(targetJson.lengthUsed) || 0;
        break;
      }
    }
    if (targetRowIdx === -1) return { success: false, message: 'Plot layout not found.' };

    let maxEnd = 0;
    plotData.forEach(r => {
      try {
        const j = JSON.parse(String(r[5] || '{}'));
        if (j.rollId === targetRollId && (j.status === 'PRINTED' || j.isDamage)) {
          const e = parseFloat(j.endAtFt) || 0;
          if (e > maxEnd) maxEnd = e;
        }
      } catch(e){}
    });

    targetJson.status = 'PRINTED';
    targetJson.startAtFt = maxEnd;
    targetJson.endAtFt = maxEnd + targetLength;

    plotSh.getRange(targetRowIdx, 6).setValue(JSON.stringify(targetJson));

    const joNumbers = Array.isArray(targetJson.joNumbers) ? targetJson.joNumbers : [];
    if (joNumbers.length) {
      const joSh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
      const joData = joSh.getRange(2, 1, joSh.getLastRow() - 1, 13).getValues();
      joData.forEach((r, i) => {
        const jo = String(r[JO_COL.JO_NUMBER]).trim().toUpperCase();
        if (joNumbers.includes(jo) && String(r[JO_COL.STATUS]).trim() === JO_STATUS.READY_TO_PRINT) {
          joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.PRINTING);
        }
      });
    }

    const today = new Date();
    const usageSh = prism_sh_(PRISM_SHEETS.LFP_USAGE);
    const planRows = Array.isArray(targetJson.rows) ? targetJson.rows : [];
    const joLengths = {};
    planRows.forEach(row => {
      const rowJOs = {};
      (row.pieces || []).forEach(p => { rowJOs[String(p.joNumber || '').trim().toUpperCase()] = true; });
      Object.keys(rowJOs).forEach(jo => {
        if (!jo) return;
        const rowH = parseFloat(row.rowH || 0) || 0;
        joLengths[jo] = (joLengths[jo] || 0) + rowH;
      });
    });

    if (!Object.keys(joLengths).length && joNumbers.length) {
      const perJO = targetLength / joNumbers.length;
      joNumbers.forEach(jo => { joLengths[jo] = perJO; });
    }

    Object.keys(joLengths).forEach(jo => {
      const len = joLengths[jo];
      if (len <= 0) return;
      usageSh.getRange(usageSh.getLastRow() + 1, 1, 1, 8).setValues([[
        'USE-' + today.getTime() + '-' + jo + '-' + targetRollId,
        jo,
        targetRollId,
        targetJson.rollWidth || 0,
        len,
        user.email,
        targetJson.plottingLink || '',
        today
      ]]);
    });

    prism_audit_('PRISM_START_PRINTING_LAYOUT', { plotId: plotId, rollId: targetRollId, by: user.email });

    return { success: true, message: 'Layout sent to printer!', startAtFt: maxEnd, endAtFt: maxEnd + targetLength };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function prism_getPrintQueueData() {
  try {
    const plotSh = prism_sh_(PRISM_SHEETS.PLOTTING_LOG);
    const plotLr = plotSh.getLastRow();
    const plotData = plotLr >= 2 ? plotSh.getRange(2, 1, plotLr - 1, 6).getValues() : [];

    const planned = [];
    const rollsHistory = {};

    plotData.forEach(r => {
      const id = String(r[0]).trim();
      let j = {};
      try { j = JSON.parse(String(r[5] || '{}')); } catch(e){}
      if (!j.rollId || j.isVoid) return;

      if (j.status === 'PLANNED') {
        planned.push({
          plotId: id,
          joNumbers: j.joNumbers || [],
          rollId: j.rollId,
          lengthUsed: j.lengthUsed || 0,
          date: r[4] ? prism_fmtShort_(new Date(r[4])) : '',
          pngUrl: r[2] || j.plottingLink || j.pngUrl || '',
          rows: j.rows || []
        });
      } else if (j.status === 'PRINTED' || j.isDamage) {
        if (!rollsHistory[j.rollId]) rollsHistory[j.rollId] = [];
        rollsHistory[j.rollId].push({
          plotId: id,
          rollId: j.rollId,
          startAtFt: j.startAtFt || 0,
          endAtFt: j.endAtFt || 0,
          lengthUsed: j.lengthUsed || 0,
          joNumbers: j.joNumbers || [],
          isDamage: !!j.isDamage,
          pngUrl: r[2] || j.plottingLink || j.pngUrl || ''
        });
      }
    });

    return { success: true, planned: planned, rollsHistory: rollsHistory };
  } catch(e) {
    return { success: false, message: e.message };
  }
}