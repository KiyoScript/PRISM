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
  WIDTH: 4, HEIGHT: 5, QUANTITY: 6, PRODUCTION_TYPE: 7,
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
  FOR_PLOTTING: 'FOR_PLOTTING', PLOTTED: 'PLOTTED',
  FOR_PRINTING: 'FOR_PRINTING', PRINTING: 'PRINTING', COMPLETED: 'COMPLETED'
};
const ROLL_STATUS = { UNOPENED: 'UNOPENED', OPEN: 'OPEN', CONSUMED: 'CONSUMED' };

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
    { name: PRISM_SHEETS.JOB_ORDERS,    headers: ['JO_Number','Customer','JobDescription','Category','Width','Height','Quantity','ProductionType','PlottingLink','Status','RollID','CreatedBy','DateCreated'] },
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
  return sh.getRange(2,1,lr-1,9).getValues()
    .filter(r => r[ROLL_COL.ROLL_ID] && String(r[ROLL_COL.ROLL_ID]).trim())
    .map((r,i) => ({
      rowIndex:        i+2,
      rollId:          String(r[ROLL_COL.ROLL_ID]).trim(),
      materialCode:    String(r[ROLL_COL.MATERIAL_CODE]).trim(),
      width:           parseFloat(r[ROLL_COL.WIDTH])||0,
      originalLength:  parseFloat(r[ROLL_COL.ORIGINAL_LENGTH])||0,
      remainingLength: parseFloat(r[ROLL_COL.REMAINING_LENGTH])||0,
      status:          String(r[ROLL_COL.STATUS]||'UNOPENED').trim().toUpperCase(),
      dateReceived:    prism_fmtShort_(r[ROLL_COL.DATE_RECEIVED]),
      dateOpened:      prism_fmtShort_(r[ROLL_COL.DATE_OPENED]),
      openedBy:        String(r[ROLL_COL.OPENED_BY]||'').trim()
    }));
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

          // ── Advance JO status to PRINTING ──
          const cs = String(r[JO_COL.STATUS]).trim();
          if (cs === JO_STATUS.FOR_PRINTING || cs === JO_STATUS.PLOTTED) {
            joSh.getRange(i + 2, JO_COL.STATUS + 1).setValue(JO_STATUS.PRINTING);
          }

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

// ============================================================
//  JOB ORDERS
// ============================================================
function prism_getAllJobOrders_() {
  const sh = prism_sh_(PRISM_SHEETS.JOB_ORDERS);
  const lr = sh.getLastRow();
  if (lr<2) return [];
  return sh.getRange(2,1,lr-1,13).getValues()
    .filter(r=>r[JO_COL.JO_NUMBER]&&String(r[JO_COL.JO_NUMBER]).trim())
    .map((r,i)=>({
      rowIndex: i+2,
      joNumber:       String(r[JO_COL.JO_NUMBER]).trim(),
      customer:       String(r[JO_COL.CUSTOMER]||'').trim(),
      jobDescription: String(r[JO_COL.JOB_DESCRIPTION]||'').trim(),
      category:       String(r[JO_COL.CATEGORY]||'').trim(),
      width:          parseFloat(r[JO_COL.WIDTH])||0,
      height:         parseFloat(r[JO_COL.HEIGHT])||0,
      quantity:       parseInt(r[JO_COL.QUANTITY])||0,
      productionType: String(r[JO_COL.PRODUCTION_TYPE]||'').trim(),
      plottingLink:   String(r[JO_COL.PLOTTING_LINK]||'').trim(),
      status:         String(r[JO_COL.STATUS]||JO_STATUS.FOR_PLOTTING).trim(),
      rollId:         String(r[JO_COL.ROLL_ID]||'').trim(),
      createdBy:      String(r[JO_COL.CREATED_BY]||'').trim(),
      dateCreated:    prism_fmtShort_(r[JO_COL.DATE_CREATED])
    }))
    .sort((a, b) => new Date(b.dateCreated) - new Date(a.dateCreated));
}
function prism_getJobOrdersPublic() {
  try { return { success:true, data:prism_getAllJobOrders_() }; }
  catch(e) { return { success:false, message:e.message }; }
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

    const LFP_CATEGORIES = ['banner', 'sticker', 'signage', 'canvas'];
    const isLFP = LFP_CATEGORIES.includes((payload.category || '').toLowerCase());
    const productionType = isLFP ? 'LFP' : (payload.productionType || '');
    const status = isLFP ? 'FOR_PLOTTING' : (payload.status || JO_STATUS.FOR_PLOTTING);

    sh.getRange(sh.getLastRow()+1,1,1,13).setValues([[
      payload.joNumber.trim().toUpperCase(), payload.customer.trim(),
      payload.jobDescription||'', payload.category,
      parseFloat(payload.width)||0, parseFloat(payload.height)||0, parseInt(payload.quantity)||1,
      productionType,              // ← computed, not payload.productionType
      payload.plottingLink||'',
      status,                      // ← computed, not payload.status
      '', user.email, today
    ]]);

    prism_audit_('PRISM_SUBMIT_JO',{joNumber:payload.joNumber,by:user.email});
    return { success:true, message:'JO "'+payload.joNumber+'" submitted.' };
  } catch(e) { return { success:false, message:e.message }; }
}
function prism_updateJOStatus(joNumber, newStatus) {
  try {
    const user = prism_getUserInfo_();
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