// ============================================================
// DEPOT DAILY - FIXED Code.gs  (v2 — all bugs corrected)
// Fixes applied:
//  [HIGH-1] updateTcRowsWithConductor: now saves cash/pass/students/dhaba/epkm
//  [HIGH-2] saveEntry DIESEL: sanKm now looked up from TC row, not optKm
//  [HIGH-3] saveEntry SERVICE: same sanKm fix
//  [HIGH-4] buildFuelRows: uses stored row[17] kmpl instead of recalculating
//  [MED-1]  detectColumns: now actually reads sheet headers to detect layout
//  [MED-2]  _COL global now properly used by getCol()
//  [MED-3]  Badge filter regex removed (was filtering valid badges)
//  [MED-4]  saveMultiSchFuel: re-reads rows after deletions for accuracy
//  [LOW-1]  getTrafficReport: nightDays only counted when ALL trips on date are night
//  [LOW-2]  buildTopEpkm: robust cache key separator
//  [LOW-3]  getSchNos: CC/EXT filter documented with comment
// ============================================================

var SHEET_NAME   = 'COMBINE DATA';
var MAIN_DATA    = 'MAIN DATA';
var VEHICLE_SH   = 'VEHICLE DATA';
var STAFF_SH     = 'STAFF DATA';
var MISC_FUEL_SH = 'MISC FUEL DATA';
var ONLY_CASH_SH = 'ONLY CASH';

// ── SERVE HTML ──────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Depot Daily')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════════════════════════════════════════════════════
// VEHICLE ALLOTMENT FUNCTIONS
// ═══════════════════════════════════════════════════════════════
var ALLOT_SH = 'ALLOTMENT DATA';
var ALLOT_HEADERS = ['Date','Schedule ID','SCH No','Vehicle No','Vehicle Type','Source','From Depot','Auth By','Remark','Time'];

function ensureAllotHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ALLOT_SH);
  if (!sh) {
    sh = ss.insertSheet(ALLOT_SH);
    sh.getRange(1,1,1,ALLOT_HEADERS.length).setValues([ALLOT_HEADERS]);
    sh.getRange(1,1,1,ALLOT_HEADERS.length).setFontWeight('bold').setBackground('#E8EAED');
    sh.setFrozenRows(1);
    return sh;
  }
  var first = sh.getLastColumn() > 0 ? sh.getRange(1,1).getValue() : '';
  if (first !== 'Date') {
    sh.getRange(1,1,1,ALLOT_HEADERS.length).setValues([ALLOT_HEADERS]);
    sh.getRange(1,1,1,ALLOT_HEADERS.length).setFontWeight('bold').setBackground('#E8EAED');
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveAllotment(date, schId, schNo, vehicleNo, vehicleType, source, depot, auth, remark) {
  try {
    var sh = ensureAllotHeaders();
    var rows = sh.getDataRange().getValues();
    var now = new Date();
    var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    source = (source || 'OWN').toString().toUpperCase();

    for (var i = rows.length - 1; i >= 1; i--) {
      var rd = formatDateVal(rows[i][0]);
      if (rd === date && String(rows[i][1]) === String(schId)) {
        sh.deleteRow(i + 1);
      }
    }
    sh.appendRow([
      date,
      String(schId),
      String(schNo || ''),
      String(vehicleNo),
      String(vehicleType || ''),
      source,
      String(depot  || ''),
      String(auth   || ''),
      String(remark || ''),
      timeStr
    ]);
    return { ok: true, message: 'Vehicle ' + vehicleNo + ' allotted to ' + schId };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getAllotment(date, schId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ALLOT_SH);
    if (!sh) return null;
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var rd = formatDateVal(rows[i][0]);
      if (rd === date && String(rows[i][1]) === String(schId)) {
        return {
          vehicleNo:   String(rows[i][3]),
          vehicleType: String(rows[i][4] || ''),
          source:      String(rows[i][5] || 'OWN'),
          depot:       String(rows[i][6] || ''),
          auth:        String(rows[i][7] || ''),
          remark:      String(rows[i][8] || '')
        };
      }
    }
    return null;
  } catch(e) { return null; }
}

function getDailyAllotments(date) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ALLOT_SH);
    if (!sh) return [];
    var rows = sh.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      var rd = formatDateVal(rows[i][0]);
      if (rd === date) {
        result.push({
          schId:       String(rows[i][1]),
          schNo:       String(rows[i][2] || ''),
          vehicleNo:   String(rows[i][3]),
          vehicleType: String(rows[i][4] || ''),
          source:      String(rows[i][5] || 'OWN'),
          depot:       String(rows[i][6] || ''),
          auth:        String(rows[i][7] || ''),
          remark:      String(rows[i][8] || ''),
          time:        String(rows[i][9] || '')
        });
      }
    }
    return result;
  } catch(e) { return []; }
}

// ═══════════════════════════════════════════════════════════════
// MISC FUEL
// ═══════════════════════════════════════════════════════════════
var MISC_FUEL_HEADERS = [
  'Date','Vehicle No','Vehicle Type','Driver Badge',
  'From Location','To Location','Distance KM',
  'Diesel L','Uria L','KMPL',
  'Movement Type','Work Order','Reason','Expected Return','Entry Time'
];

function saveMiscFuel(data) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MISC_FUEL_SH);

    if (!sheet) {
      sheet = ss.insertSheet(MISC_FUEL_SH);
      sheet.getRange(1,1,1,MISC_FUEL_HEADERS.length).setValues([MISC_FUEL_HEADERS]);
      sheet.getRange(1,1,1,MISC_FUEL_HEADERS.length)
           .setFontWeight('bold').setBackground('#E8EAED');
      sheet.setFrozenRows(1);
    }

    var distKm  = parseFloat(data.distanceKm) || 0;
    var dieselL = parseFloat(data.dieselL)    || 0;
    var kmpl    = (dieselL > 0 && distKm > 0) ? parseFloat((distKm / dieselL).toFixed(2)) : 0;

    var entryTime = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss'
    );

    sheet.appendRow([
      data.date          || '',
      data.vehicleNo     || '',
      data.vehicleType   || '',
      data.driverBadge   || '',
      data.fromLocation  || '',
      data.toLocation    || '',
      distKm,
      dieselL,
      parseFloat(data.uriaL) || 0,
      kmpl,
      (data.movementType || 'OTHER').toString().toUpperCase(),
      data.workOrder     || '',
      data.reason        || '',
      data.expectedReturn|| '',
      entryTime
    ]);

    return { success: true, message: 'Misc fuel entry saved successfully.' };
  } catch(e) {
    Logger.log('saveMiscFuel error: ' + e.message);
    return { success: false, message: 'Error saving misc fuel: ' + e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// MARK VEHICLE AWAY / AVAILABLE
// ═══════════════════════════════════════════════════════════════
function markVehicleAway(vehicleNo, status, fromDate, returnDate, reason) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(VEHICLE_SH);
    if (!sheet) return { success: false, message: 'VEHICLE DATA sheet not found.' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0];

    var vehicleNoCol  = -1;
    var statusCol     = -1;
    var fromDateCol   = -1;
    var returnDateCol = -1;
    var remarkCol     = -1;

    headers.forEach(function(h, i) {
      var hl = h.toString().toLowerCase().trim();
      if (hl === 'vehicle no' || hl === 'vehicleno' || hl === 'vehicle_no') vehicleNoCol  = i;
      if (hl === 'status')                                                    statusCol     = i;
      if (hl === 'away from' || hl === 'from date' || hl === 'away date')    fromDateCol   = i;
      if (hl === 'expected return' || hl === 'return date')                   returnDateCol = i;
      if (hl === 'remark' || hl === 'reason' || hl === 'remarks')            remarkCol     = i;
    });

    if (vehicleNoCol === -1) {
      return { success: false, message: 'Vehicle No column not found in VEHICLE DATA sheet.' };
    }

    var found = false;
    for (var r = 1; r < data.length; r++) {
      if (data[r][vehicleNoCol].toString().trim().toUpperCase() ===
          vehicleNo.toString().trim().toUpperCase()) {
        var rowNum = r + 1;
        if (statusCol     !== -1) sheet.getRange(rowNum, statusCol     + 1).setValue(status.toUpperCase());
        if (fromDateCol   !== -1) sheet.getRange(rowNum, fromDateCol   + 1).setValue(fromDate   || '');
        if (returnDateCol !== -1) sheet.getRange(rowNum, returnDateCol + 1).setValue(returnDate || '');
        if (remarkCol     !== -1) sheet.getRange(rowNum, remarkCol     + 1).setValue(reason     || '');
        found = true;
        break;
      }
    }

    if (!found) return { success: false, message: 'Vehicle ' + vehicleNo + ' not found in VEHICLE DATA sheet.' };
    return { success: true, message: 'Vehicle ' + vehicleNo + ' marked as ' + status + '.' };
  } catch(e) {
    Logger.log('markVehicleAway error: ' + e.message);
    return { success: false, message: 'Error updating vehicle status: ' + e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// TRAFFIC REPORT
// ═══════════════════════════════════════════════════════════════
function getTrafficReport(params) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var comSh = ss.getSheetByName(SHEET_NAME);
    var stfSh = ss.getSheetByName(STAFF_SH);
    var mainSh= ss.getSheetByName(MAIN_DATA);

    if (!comSh) return { error: 'COMBINE DATA sheet not found' };
    if (!stfSh) return { error: 'STAFF DATA sheet not found' };

    var badge    = String(params.badge||'').trim();
    var from     = String(params.from||'');
    var to       = String(params.to||'');
    var ptype    = String(params.ptype||'driver').toLowerCase();
    var lineRate = parseFloat(params.lineRate) || 0;
    var nightRate= parseFloat(params.nightRate)|| 0;
    var phCount  = parseInt(params.phCount)    || 0;

    if (!badge) return { error: 'Badge number is required' };

    var staffRows = stfSh.getDataRange().getValues();
    var staffInfo = null;
    for (var i = 1; i < staffRows.length; i++) {
      var sr = staffRows[i];
      var design = String(sr[1]||'').toUpperCase();
      var sbadge = String(sr[2]||'').replace(/\.0$/,'').trim();
      if (sbadge === badge) {
        staffInfo = {
          badge:   sbadge,
          design:  design,
          name:    String(sr[3]||''),
          offDay:  String(sr[4]||''),
          status:  String(sr[5]||''),
          mobile:  String(sr[6]||'').replace(/\.0$/,'')
        };
        break;
      }
    }
    if (!staffInfo) return { error: 'Badge ' + badge + ' not found in STAFF DATA' };

    detectColumns(comSh);
    var rows = comSh.getDataRange().getValues();
    var trips = [];

    var schNameMap = {};
    if (mainSh) {
      var mRows = mainSh.getDataRange().getValues();
      var lastId = '';
      for (var mi = 1; mi < mRows.length; mi++) {
        if (mRows[mi][0] && String(mRows[mi][0]).trim()) lastId = String(mRows[mi][0]).trim();
        if (lastId && !schNameMap[lastId]) schNameMap[lastId] = String(mRows[mi][1]||'');
      }
    }

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var rDate = formatDateVal(r[0]);
      if (from && rDate < from) continue;
      if (to   && rDate > to)   continue;
      var mod = String(getCol(r,'mod'));
      if (mod !== 'TC') continue;

      var dBadge = String(getCol(r,'driverBadge')||'').replace(/\.0$/,'').trim();
      var cBadge = String(getCol(r,'condBadge')||'').replace(/\.0$/,'').trim();

      var isMatch = (ptype === 'driver' && dBadge === badge) ||
                    (ptype === 'conductor' && cBadge === badge);
      if (!isMatch) continue;

      var schId   = String(getCol(r,'schId'));
      var schNo   = String(getCol(r,'schNo')||'').replace(/\.0$/,'');
      var route   = String(getCol(r,'route')||'');
      var optKm   = parseFloat(getCol(r,'optKm'))||0;
      var status  = String(getCol(r,'status'));
      var tripType= String(getCol(r,'tripType')||'');
      var isNight = (tripType === 'SLP' || tripType === 'NIGHT');

      trips.push({
        date:     rDate,
        schId:    schId,
        schNo:    schNo,
        schName:  schNameMap[schId]||'',
        route:    route,
        km:       optKm,
        status:   status,
        tripType: tripType,
        isNight:  isNight,
        vehicle:  String(getCol(r,'vehicle'))
      });
    }

    if (!trips.length) return { found: false, staffInfo: staffInfo, trips: [] };

    var datesSet   = {};
    var nightSet   = {};   // dates where ALL trips for this badge are night
    var totalKm    = 0;
    var totalTrips = 0;
    var yesTrips   = 0;

    // FIX [LOW-1]: Build per-date trip lists to properly determine night days
    var dateTripsMap = {};
    trips.forEach(function(t) {
      if (!dateTripsMap[t.date]) dateTripsMap[t.date] = [];
      dateTripsMap[t.date].push(t);
      datesSet[t.date] = true;
      totalKm    += t.km;
      totalTrips++;
      if (t.status === 'YES') yesTrips++;
    });

    // A date is a "night day" only if every trip on that date is a night trip
    Object.keys(dateTripsMap).forEach(function(dt) {
      var dayTrips = dateTripsMap[dt];
      var allNight = dayTrips.every(function(t) { return t.isNight; });
      if (allNight) nightSet[dt] = true;
    });

    var dutyDays  = Object.keys(datesSet).length;
    var nightDays = Object.keys(nightSet).length;
    var lineAllow = lineRate  * dutyDays;
    var nightAllow= nightRate * nightDays;
    var totalPay  = lineAllow + nightAllow;

    var summary = {
      dutyDays:   dutyDays,
      nightDays:  nightDays,
      phCount:    phCount,
      totalKm:    parseFloat(totalKm.toFixed(2)),
      totalTrips: totalTrips,
      yesTrips:   yesTrips,
      lineRate:   lineRate,
      nightRate:  nightRate,
      lineAllow:  lineAllow,
      nightAllow: nightAllow,
      totalPay:   totalPay
    };

    trips.sort(function(a,b){ return a.date < b.date ? -1 : 1; });
    return { found: true, staffInfo: staffInfo, trips: trips, summary: summary };
  } catch(e) {
    return { error: e.message };
  }
}

// ── DIAGNOSTIC ───────────────────────────────────────────────
function diagnoseSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = ss.getSheets().map(function(s) { return s.getName(); });
  Logger.log('=== All Sheet Names ===');
  names.forEach(function(n) { Logger.log(n); });
  Logger.log('Expected: MAIN DATA, COMBINE DATA, VEHICLE DATA, STAFF DATA');
  return names;
}

// ── COMBINE DATA HEADERS (30 columns) ───────────────────────
var HEADERS = [
  'Date',            // 0
  'Module',          // 1
  'Schedule ID',     // 2
  'SCH No',          // 3
  'Route',           // 4
  'Trip Type',       // 5
  'San KM',          // 6
  'Opt KM',          // 7
  'Cancel KM',       // 8
  'Status',          // 9
  'Vehicle No',      // 10
  'Vehicle Type',    // 11
  'Driver Badge',    // 12
  'Driver Name',     // 13
  'Driver Mobile',   // 14
  'Diesel L',        // 15
  'Uria L',          // 16
  'KMPL',            // 17
  'Dead KM',         // 18
  'Diversion KM',    // 19
  'Conductor Badge', // 20
  'Conductor Name',  // 21
  'Conductor Mobile',// 22
  'Bag No',          // 23
  'Cash Total',      // 24
  'All Pass',        // 25
  'Student Pass',    // 26
  'Dhaba',           // 27
  'EPKM',            // 28
  'Entry Time'       // 29
];

// FIX [MED-1 + MED-2]: _COL is now properly set by detectColumns()
// and used by getCol(). detectColumns() reads actual sheet headers.
var _COL = null;

var _COL_NEW = {
  date:0, mod:1, schId:2, schNo:3, route:4, tripType:5,
  sanKm:6, optKm:7, cancelKm:8, status:9,
  vehicle:10, vehType:11, driverBadge:12, driverName:13, driverMobile:14,
  diesel:15, uria:16, kmpl:17, deadKm:18, divKm:19,
  condBadge:20, condName:21, condMobile:22, bagNo:23,
  cash:24, allPass:25, students:26, dhaba:27, epkm:28, entryTime:29
};

var _COL_OLD = {
  date:0, mod:1, schId:2, schNo:3, route:4, tripType:-1,
  sanKm:5, optKm:6, cancelKm:7, status:8,
  vehicle:9, vehType:-1, driverBadge:10, driverName:-1, driverMobile:11,
  diesel:12, uria:13, kmpl:14, deadKm:15, divKm:16,
  condBadge:-1, condName:-1, condMobile:-1, bagNo:-1,
  cash:17, allPass:18, students:-1, dhaba:-1, epkm:-1, entryTime:-1
};

// FIX [MED-1]: Actually detect column layout from sheet headers
function detectColumns(sh) {
  _COL = null;
  if (!sh) {
    _COL = _COL_NEW;
    return _COL;
  }
  var lastCol = sh.getLastColumn();
  if (lastCol < 2) {
    _COL = _COL_NEW;
    return _COL;
  }
  // Check actual header count: new layout has 30 cols, old has ~19
  if (lastCol >= 27) {
    _COL = _COL_NEW;
  } else {
    _COL = _COL_OLD;
  }
  return _COL;
}

// FIX [MED-2]: getCol now uses the _COL global set by detectColumns()
// Row-length fallback retained as safety net
function getCol(row, colName) {
  var map = _COL ? _COL : ((row && row.length >= 27) ? _COL_NEW : _COL_OLD);
  var idx = map[colName];
  if (idx === undefined || idx < 0 || idx >= row.length) return '';
  var v = row[idx];
  return (v === null || v === undefined) ? '' : v;
}

function ensureHeaders() {
  _COL = null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#E8EAED');
    sh.setFrozenRows(1);
    detectColumns(sh);
    return sh;
  }
  var lastCol = sh.getLastColumn();
  var firstHeader = lastCol > 0 ? sh.getRange(1,1).getValue() : '';
  if (firstHeader !== 'Date' || lastCol < 20) {
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#E8EAED');
    sh.setFrozenRows(1);
  }
  detectColumns(sh);
  return sh;
}

// ── Helper: look up sanKm from TC row for a given date+schId ─
function getTcSanKm(sh, dateStr, schId) {
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (formatDateVal(rows[i][0]) === dateStr &&
        String(rows[i][1]) === 'TC' &&
        String(rows[i][2]) === String(schId)) {
      return parseFloat(getCol(rows[i], 'sanKm')) || 0;
    }
  }
  return 0;
}

function saveEntry(data, module) {
  try {
    var sh = ensureHeaders();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSh = ss.getSheetByName(MAIN_DATA);
    var now = new Date();
    var dateStr;
    if (module === 'TC') {
      dateStr = (data && data[0] && data[0].date)
                  ? String(data[0].date).substring(0, 10)
                  : Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateStr = (data && data.date)
                  ? String(data.date).substring(0, 10)
                  : Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    if (module === 'TC') {
      var schId = data[0] ? String(data[0].masterId) : '';
      deleteDuplicateRows(sh, dateStr, 'TC', schId);

      var rows = [];
      data.forEach(function(t) {
        var tripType = getTripType(mainSh, t.masterId, t.schNo, t.fallbackType || '');
        var row = new Array(HEADERS.length).fill('');
        row[0]  = dateStr;
        row[1]  = 'TC';
        row[2]  = t.masterId;
        row[3]  = t.schNo;
        row[4]  = t.route;
        row[5]  = tripType;
        row[6]  = parseFloat(t.km)       || 0;   // sanKm from schedule
        row[7]  = parseFloat(t.optKm)    || 0;
        row[8]  = parseFloat(t.cancelKm) || 0;
        row[9]  = t.status;
        row[10] = t.vehicle || '';    // Vehicle No — save immediately, also updated on fuel entry
        row[29] = now;
        rows.push(row);
      });
      if (rows.length > 0) {
        sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
      }
      return 'TC saved: ' + rows.length + ' trips for ' + schId + ' on ' + dateStr;
    }

    if (module === 'DIESEL') {
      deleteDuplicateRows(sh, dateStr, 'DIESEL', data.id);

      // FIX [HIGH-2]: Look up actual sanKm from TC row instead of using optKm
      var tcSanKm = getTcSanKm(sh, dateStr, data.id);

      var _dTripType = '';
      var _allRows = sh.getDataRange().getValues();
      for (var _di = 1; _di < _allRows.length; _di++) {
        if (formatDateVal(_allRows[_di][0]) === dateStr &&
            String(_allRows[_di][2]) === String(data.id) &&
            String(_allRows[_di][1]) === 'TC') {
          _dTripType = getCol(_allRows[_di], 'tripType') || '';
          break;
        }
      }

      var row = new Array(HEADERS.length).fill('');
      row[0]  = dateStr;
      row[1]  = 'DIESEL';
      row[2]  = data.id;
      row[3]  = data.schNo || '';
      row[4]  = 'Fuel Entry';
      row[5]  = _dTripType;
      row[6]  = tcSanKm;                          // FIX: actual sanKm
      row[7]  = parseFloat(data.optKm)    || 0;
      row[8]  = 0;
      row[9]  = 'DIESEL';
      row[10] = data.vehicle     || '';
      row[11] = data.vehicleType || '';
      row[12] = data.driverBadge || '';
      row[13] = data.driverName  || '';
      row[14] = data.driverMobile || '';
      row[15] = parseFloat(data.diesel)      || 0;
      row[16] = parseFloat(data.uria)        || 0;
      row[17] = parseFloat(data.kmpl)        || 0;
      row[18] = parseFloat(data.deadKm)      || 0;
      row[19] = parseFloat(data.diversionKm) || 0;
      row[29] = now;
      sh.getRange(sh.getLastRow() + 1, 1, 1, HEADERS.length).setValues([row]);

      updateTcRowsWithDriver(sh, dateStr, data.id, data.vehicle, data.vehicleType,
                             data.driverBadge, data.driverName, data.driverMobile);

      return 'Fuel saved for ' + data.id + ' on ' + dateStr;
    }

    if (module === 'SERVICE') {
      deleteDuplicateRows(sh, dateStr, 'SERVICE', data.id);

      // FIX [HIGH-3]: Look up actual sanKm from TC row instead of using optKm
      var tcSanKmSvc = getTcSanKm(sh, dateStr, data.id);

      var _sTripType = '';
      var _sAllRows = sh.getDataRange().getValues();
      for (var _si = 1; _si < _sAllRows.length; _si++) {
        if (formatDateVal(_sAllRows[_si][0]) === dateStr &&
            String(_sAllRows[_si][2]) === String(data.id) &&
            String(_sAllRows[_si][1]) === 'TC') {
          _sTripType = getCol(_sAllRows[_si], 'tripType') || '';
          break;
        }
      }

      var row = new Array(HEADERS.length).fill('');
      row[0]  = dateStr;
      row[1]  = 'SERVICE';
      row[2]  = data.id;
      row[3]  = data.schNo || '';
      row[4]  = 'Cash Entry';
      row[5]  = _sTripType;
      row[6]  = tcSanKmSvc;                          // FIX: actual sanKm
      row[7]  = parseFloat(data.optKm)    || 0;
      row[8]  = 0;
      row[9]  = 'SERVICE';
      row[20] = data.conductorBadge   || '';
      row[21] = data.conductorName    || '';
      row[22] = data.conductorMobile  || '';
      row[23] = data.bagNo            || '';
      row[24] = parseFloat(data.cashTotal)   || 0;
      row[25] = parseFloat(data.passCount)   || 0;
      row[26] = parseFloat(data.studentPass) || 0;
      row[27] = parseFloat(data.dhaba)       || 0;
      row[28] = parseFloat(data.epkm)        || 0;
      row[29] = now;
      sh.getRange(sh.getLastRow() + 1, 1, 1, HEADERS.length).setValues([row]);

      updateTcRowsWithConductor(sh, dateStr, data.id,
                                data.conductorBadge, data.conductorName, data.conductorMobile,
                                data.bagNo, data.cashTotal, data.passCount,
                                data.studentPass, data.dhaba, data.epkm);

      return 'Cash saved for ' + data.id + ' on ' + dateStr;
    }

    return 'Unknown module: ' + module;

  } catch(e) {
    return 'ERROR: ' + e.message + '\n' + e.stack;
  }
}

function getTripType(mainSh, schId, schNo, fallback) {
  if (mainSh) {
    var data = mainSh.getDataRange().getValues();
    var lastId = '';
    for (var i = 1; i < data.length; i++) {
      var rowId = data[i][0];
      if (rowId !== null && rowId !== undefined && String(rowId).trim() !== '') {
        lastId = String(rowId).trim();
      }
      if (lastId === String(schId).trim() && String(data[i][3]).replace(/\.0$/,'') === String(schNo)) {
        var found = data[i][5] || '';
        if (found) return found;
        break;
      }
    }
  }
  return fallback || '';
}

function deleteDuplicateRows(sh, dateStr, module, schId) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var data = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  var toDelete = [];
  for (var i = 0; i < data.length; i++) {
    var rDate = formatDateVal(data[i][0]);
    var rMod  = String(data[i][1]);
    var rSch  = String(data[i][2]);
    if (rDate === dateStr && rMod === String(module) && rSch === String(schId)) {
      toDelete.push(i + 2);
    }
  }
  for (var j = toDelete.length - 1; j >= 0; j--) {
    sh.deleteRow(toDelete[j]);
  }
  return toDelete.length;
}

function updateTcRowsWithDriver(sh, dateStr, schId, vehicle, vType, badge, name, mobile) {
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (formatDateVal(data[i][0]) === dateStr &&
        data[i][1] === 'TC' &&
        data[i][2] === schId) {
      sh.getRange(i+1, 11).setValue(vehicle || '');
      sh.getRange(i+1, 12).setValue(vType  || '');
      sh.getRange(i+1, 13).setValue(badge  || '');
      sh.getRange(i+1, 14).setValue(name   || '');
      sh.getRange(i+1, 15).setValue(mobile || '');
    }
  }
}

// FIX [HIGH-1]: Now also writes cash, allPass, students, dhaba, epkm to TC rows
function updateTcRowsWithConductor(sh, dateStr, schId, badge, name, mobile,
                                   bagNo, cash, pass, students, dhaba, epkm) {
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (formatDateVal(data[i][0]) === dateStr &&
        data[i][1] === 'TC' &&
        data[i][2] === schId) {
      sh.getRange(i+1, 21).setValue(badge    || '');
      sh.getRange(i+1, 22).setValue(name     || '');
      sh.getRange(i+1, 23).setValue(mobile   || '');
      sh.getRange(i+1, 24).setValue(bagNo    || '');
      sh.getRange(i+1, 25).setValue(parseFloat(cash)     || 0);  // FIX: cash
      sh.getRange(i+1, 26).setValue(parseFloat(pass)     || 0);  // FIX: allPass
      sh.getRange(i+1, 27).setValue(parseFloat(students) || 0);  // FIX: students
      sh.getRange(i+1, 28).setValue(parseFloat(dhaba)    || 0);  // FIX: dhaba
      sh.getRange(i+1, 29).setValue(parseFloat(epkm)     || 0);  // FIX: epkm
    }
  }
}

function getSchNameForId(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MAIN_DATA);
  if (!sh) return id;
  var data = sh.getDataRange().getValues();
  var lastId = '';
  for (var i = 1; i < data.length; i++) {
    var rowId = data[i][0];
    if (rowId !== null && rowId !== undefined && String(rowId).trim() !== '') {
      lastId = String(rowId).trim();
    }
    if (lastId === String(id).trim() && data[i][1]) return String(data[i][1]);
  }
  return id;
}

function getFullSchedule(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MAIN_DATA);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var trips = [];
  var lastId = '';

  for (var i = 1; i < data.length; i++) {
    var rowId = data[i][0];
    if (rowId !== null && rowId !== undefined && String(rowId).trim() !== '') {
      lastId = String(rowId).trim();
    }

    if (lastId !== String(id).trim()) continue;

    var timeStr = '';
    var rawTime = data[i][9];
    if (rawTime) {
      try {
        if (rawTime instanceof Date) {
          var h = rawTime.getHours();
          var m = rawTime.getMinutes();
          timeStr = (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
        } else {
          timeStr = String(rawTime).substring(0, 5);
        }
      } catch(e) { timeStr = ''; }
    }

    trips.push({
      schNo : data[i][3] !== null ? String(data[i][3]).replace(/\.0$/, '') : '',
      route : (data[i][6] || '') + ' → ' + (data[i][7] || ''),
      km    : parseFloat(data[i][8]) || 0,
      time  : timeStr,
      type  : data[i][5] || ''
    });
  }
  return trips;
}

function getSchNos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(MAIN_DATA);
    if (!sh) {
      var allNames = ss.getSheets().map(function(s){ return s.getName(); }).join(', ');
      return { error: 'Sheet "MAIN DATA" not found. Available: ' + allNames };
    }
    var data = sh.getDataRange().getValues();
    var ids = [];
    var seen = {};
    var lastId = '';
    for (var i = 1; i < data.length; i++) {
      var rowId = data[i][0];
      if (rowId !== null && rowId !== undefined && String(rowId).trim() !== '') {
        lastId = String(rowId).trim();
      }
      if (!lastId) continue;
      // CC-xx and EXT-xx are handled in a separate getCCExtSchNos() call
      var _upper = lastId.toUpperCase();
      if (_upper === 'CC' || _upper === 'EXT' || _upper.indexOf('CC-') === 0 || _upper.indexOf('EXT-') === 0) continue;
      if (!seen[lastId]) { seen[lastId] = true; ids.push(lastId); }
    }
    return ids.sort();
  } catch(e) {
    return { error: 'getSchNos error: ' + e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// CC / EXT SCHEDULE IDs
// Returns all schedule IDs from MAIN DATA whose col A starts with
// "CC-" or "EXT-" (case-insensitive). These are City Contract /
// External schedules — shown in a separate tab in TC entry.
// ═══════════════════════════════════════════════════════════════
function getCCExtSchNos() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName(MAIN_DATA);
    if (!sh) return ['CC','EXT']; // safe fallback — always show + CC and + EXT tiles
    var data = sh.getDataRange().getValues();
    var types = {};
    for (var i = 1; i < data.length; i++) {
      var t = String(data[i][5] || '').trim().toUpperCase(); // col F = TYPE
      if (t === 'CC' || t === 'EXT') types[t] = true;
    }
    var result = Object.keys(types).sort();
    return result.length ? result : ['CC','EXT']; // always return at least CC and EXT
  } catch(e) {
    return ['CC','EXT'];
  }
}

// ═══════════════════════════════════════════════════════════════
// CC / EXT TC ENTRY
// Saves ad-hoc CC or EXT trips to COMBINE DATA.
// Auto-assigns Sch No as CC-001, CC-002… or EXT-001, EXT-002…
// Each unique Sch No is a separate "schedule" (masterId).
// Trip No 1 on a new row = new schedule; Trip No 2+ = same schedule.
//
// saveCCExtEntry(type, date, rows)
//   type  : 'CC' or 'EXT'
//   date  : 'YYYY-MM-DD'
//   rows  : [ { tripNo, time, from, to, km, status } ]
//            tripNo = 1 → new schedule; tripNo > 1 → continue prev
//
// getExistingCCExtEntry(date, type)
//   Returns rows already saved today for this type.
// ═══════════════════════════════════════════════════════════════
function saveCCExtEntry(type, date, rows) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss.getSheetByName(SHEET_NAME);
    if (!sh) return 'ERROR: COMBINE DATA sheet not found';
    detectColumns(sh);

    var typeU  = String(type).toUpperCase().trim();   // 'CC' or 'EXT'
    var prefix = typeU + '-';
    var dateStr = String(date).substring(0, 10);
    var now     = new Date();

    // ── Delete all existing CC/EXT rows for this date+type ──
    var existing = sh.getDataRange().getValues();
    var toDelete = [];
    for (var i = 1; i < existing.length; i++) {
      if (formatDateVal(existing[i][0]) !== dateStr) continue;
      if (String(existing[i][1]) !== 'TC') continue;
      var sid = String(existing[i][2] || '').toUpperCase();
      if (sid.indexOf(prefix) === 0) toDelete.push(i + 1);
    }
    for (var d = toDelete.length - 1; d >= 0; d--) sh.deleteRow(toDelete[d]);

    // ── Assign Sch No based on trip No logic ─────────────────
    // Trip No 1 = new schedule → increment counter
    // Trip No 2+ = same schedule as previous
    var schCounter = 0;
    var currentSchId = '';
    var saveRows = [];

    rows.forEach(function(r) {
      var tripNo = parseInt(r.tripNo) || 1;
      if (tripNo === 1 || !currentSchId) {
        schCounter++;
        currentSchId = prefix + (schCounter < 10 ? '00' : schCounter < 100 ? '0' : '') + schCounter;
      }
      var km = parseFloat(r.km) || 0;
      var isCancelled = (r.status === 'NO' || r.status === 'BD');
      var route = (r.from || '') + ' → ' + (r.to || '');
      var row = new Array(HEADERS.length).fill('');
      row[0]  = dateStr;
      row[1]  = 'TC';
      row[2]  = currentSchId;
      row[3]  = tripNo;
      row[4]  = route;
      row[5]  = typeU;
      row[6]  = km;
      row[7]  = isCancelled ? 0 : km;   // FIX: optKm = 0 for cancelled trips
      row[8]  = isCancelled ? km : 0;   // FIX: cancelKm = km for cancelled trips
      row[9]  = r.status || 'YES';
      row[10] = r.vehicle     || '';
      row[11] = r.vehicleType || '';
      row[29] = now;
      saveRows.push(row);
    });

    if (saveRows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, saveRows.length, HEADERS.length).setValues(saveRows);
    }
    return 'OK:' + saveRows.length;
  } catch(e) {
    return 'ERROR: ' + e.message;
  }
}

function getExistingCCExtEntry(date, type) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss.getSheetByName(SHEET_NAME);
    if (!sh) return [];
    detectColumns(sh);

    var typeU   = String(type).toUpperCase().trim();
    var prefix  = typeU + '-';
    var dateStr = String(date).substring(0, 10);
    var data    = sh.getDataRange().getValues();
    var result  = [];

    for (var i = 1; i < data.length; i++) {
      if (formatDateVal(data[i][0]) !== dateStr) continue;
      if (String(data[i][1]) !== 'TC') continue;
      var sid = String(data[i][2] || '').toUpperCase();
      if (sid.indexOf(prefix) !== 0) continue;

      var route = String(getCol(data[i], 'route') || '');
      var parts  = route.split(' → ');
      result.push({
        schId       : String(data[i][2]),
        tripNo      : getCol(data[i], 'schNo') || 1,
        time        : '',
        from        : parts[0] || '',
        to          : parts[1] || '',
        km          : parseFloat(getCol(data[i], 'sanKm')) || 0,
        status      : getCol(data[i], 'status') || 'YES',
        vehicle     : String(getCol(data[i], 'vehicle')  || ''),
        vehicleType : String(getCol(data[i], 'vehType')  || '')
      });
    }
    return result;
  } catch(e) {
    return [];
  }
}

// Returns today's CC/EXT schIds (CC-001, CC-002, EXT-001…).
// SOURCE OF TRUTH: ALLOTMENT DATA — a CC/EXT schedule exists as soon as
// a vehicle is allotted to it. TC/fuel/cash flags come from COMBINE DATA.
function getCCExtDailySchIds(dateStr) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var allotSh = ss.getSheetByName(ALLOT_SH);
    var combSh  = ss.getSheetByName(SHEET_NAME);
    var dStr    = String(dateStr).substring(0, 10);
    var map     = {};

    // Step 1 — seed from ALLOTMENT DATA (CC-001, CC-002, EXT-001…)
    if (allotSh) {
      var aData = allotSh.getDataRange().getValues();
      for (var j = 1; j < aData.length; j++) {
        if (formatDateVal(aData[j][0]) !== dStr) continue;
        var aId = String(aData[j][1] || '').trim();
        var upper = aId.toUpperCase();
        if (!upper.startsWith('CC-') && !upper.startsWith('EXT-')) continue;
        if (!map[aId]) map[aId] = {schId:aId, tc:false, diesel:false, service:false, vehicle:''};
        if (aData[j][3]) map[aId].vehicle = String(aData[j][3]);
      }
    }

    // Step 2 — check COMBINE DATA for TC/fuel/cash done flags
    if (combSh) {
      detectColumns(combSh);
      var cData = combSh.getDataRange().getValues();
      for (var i = 1; i < cData.length; i++) {
        if (formatDateVal(cData[i][0]) !== dStr) continue;
        var schId = String(getCol(cData[i], 'schId') || '').trim();
        var upper2 = schId.toUpperCase();
        if (!upper2.startsWith('CC-') && !upper2.startsWith('EXT-')) continue;
        if (!map[schId]) map[schId] = {schId:schId, tc:false, diesel:false, service:false, vehicle:''};
        var mod = String(getCol(cData[i], 'mod') || '');
        if (mod === 'TC')      map[schId].tc = true;
        if (mod === 'DIESEL')  map[schId].diesel  = true;
        if (mod === 'SERVICE') map[schId].service = true;
      }
    }

    return Object.values(map).sort(function(a,b){ return a.schId < b.schId ? -1 : 1; });
  } catch(e) { return []; }
}

// Returns the next available CC/EXT numbered schId for today.
// e.g. if CC-001, CC-002 exist → returns CC-003
function getNextCCExtId(type, dateStr) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var allotSh = ss.getSheetByName(ALLOT_SH);
    var prefix  = String(type).toUpperCase().trim() + '-';
    var dStr    = String(dateStr).substring(0, 10);
    var max     = 0;
    if (allotSh) {
      var rows = allotSh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (formatDateVal(rows[i][0]) !== dStr) continue;
        var id = String(rows[i][1] || '').trim().toUpperCase();
        if (id.indexOf(prefix) !== 0) continue;
        var num = parseInt(id.substring(prefix.length)) || 0;
        if (num > max) max = num;
      }
    }
    var next = max + 1;
    return prefix + (next < 10 ? '00' : next < 100 ? '0' : '') + next;
  } catch(e) { return type + '-001'; }
}

function getVehicles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(VEHICLE_SH);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      list.push({
        vehicle : String(data[i][1]),
        type    : data[i][2] || '',
        model   : data[i][3] || '',
        company : data[i][4] || '',
        regNo   : data[i][5] || ''
      });
    }
  }
  return list;
}

function getDrivers()    { return _getStaff('DRIVER'); }
function getConductors() { return _getStaff('CONDUCTOR'); }
function _getStaff(design) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STAFF_SH);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === design && data[i][2]) {
      list.push({
        badge  : String(data[i][2]),
        name   : String(data[i][3] || ''),
        offDay : String(data[i][4] || ''),
        status : String(data[i][5] || ''),
        mobile : String(data[i][6] || '')
      });
    }
  }
  return list;
}

// ═══════════════════════════════════════════════════════════════
// DASHBOARD DATA
// ═══════════════════════════════════════════════════════════════
function getDashboardData(baseDate, period) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return emptyDash();

  var rows = sh.getDataRange().getValues();
  var base = new Date(baseDate);
  var from, to;
  if (period === 'day') {
    from = to = baseDate;
  } else if (period === 'week') {
    var d = new Date(base);
    d.setDate(d.getDate() - d.getDay());
    from = fmtDate(d);
    to   = baseDate;
  } else {
    from = baseDate.substring(0,7) + '-01';
    to   = baseDate;
  }

  var tcRows     = filterRows(rows, from, to, 'TC');
  var svcRows    = filterRows(rows, from, to, 'SERVICE');
  var dieselRows = filterRows(rows, from, to, 'DIESEL');
  var ocRows     = filterRows(rows, from, to, 'ONLY_CASH');
  return buildDashStats(tcRows, from, to, svcRows, dieselRows, ocRows);
}

function buildDashStats(tcRows, from, to, svcRows, dieselRows, ocRows) {
  // Sanctioned KM/Sch/Trips come from MAIN DATA
  var sanData = buildSanctionedMap();
  var sanBySchId = sanData.bySchId;
  var totalSanKmFromMain = 0;
  Object.keys(sanData.byType).forEach(function(t) { totalSanKmFromMain += sanData.byType[t].sanKm; });

  var svcMap = {};
  (svcRows||[]).forEach(function(r){ svcMap[String(getCol(r,'schId')||'')] = r; });

  var totalTrips=0, yesT=0, noT=0, bdT=0, accT=0, otherT=0;
  var totalOpt=0, totalCan=0;
  var totalDiesel=0, totalUria=0, totalKmplSum=0, kmplCount=0;
  var trendMap = {};
  var schMap = {};
  var driverMap = {};
  var conductorMap = {};

  tcRows.forEach(function(r) {
    var dt     = formatDateVal(r[0]);
    var status = getCol(r,'status') || '';
    var sanKm  = parseFloat(getCol(r,'sanKm'))  || 0;
    var optKm  = parseFloat(getCol(r,'optKm'))  || 0;
    var canKm  = parseFloat(getCol(r,'cancelKm')) || 0;
    var schId  = String(getCol(r,'schId') || '');
    var drvBdg = String(getCol(r,'driverBadge') || '').trim();
    var drvNm  = getCol(r,'driverName') || '';
    var drvMob = getCol(r,'driverMobile') || '';
    var cndBdg = String(getCol(r,'condBadge') || '').trim();
    var cndNm  = getCol(r,'condName') || '';
    var cndMob = getCol(r,'condMobile') || '';

    totalTrips++;
    if(status==='YES'){yesT++;totalOpt+=optKm;}
    else if(status==='NO'){noT++;totalCan+=canKm;}
    else if(status==='BD'){bdT++;}
    else if(status==='ACCIDENT'){accT++;}
    else{otherT++;}

    if(!trendMap[dt]) trendMap[dt]={label:dt.substring(5),sanKm:0,optKm:0,cancelKm:0,diesel:0,kmpl:0,kmplC:0,cash:0,epkm:0};
    // FIX [BUG-DASH-SANKM]: += not = so multiple schedules per day accumulate
    var _sanKmForTrend = (sanBySchId[schId] && sanBySchId[schId].sanKm) ? sanBySchId[schId].sanKm : sanKm;
    if (!trendMap[dt]._sanAdded) trendMap[dt]._sanAdded = {};
    if (!trendMap[dt]._sanAdded[schId]) { trendMap[dt].sanKm += _sanKmForTrend; trendMap[dt]._sanAdded[schId]=true; }
    if(status==='YES'){trendMap[dt].optKm+=optKm;}else{trendMap[dt].cancelKm+=canKm;}

    if(schId){
      if(!schMap[schId]) schMap[schId]={id:schId,totalTrips:0,yesTrips:0,
        sanKm: (sanBySchId[schId] && sanBySchId[schId].sanKm) ? sanBySchId[schId].sanKm : sanKm,
        optKm:0,cancelKm:0,diesel:0,kmpl:0,kmplC:0,cash:0,epkm:0};
      schMap[schId].totalTrips++;
      if(status==='YES'){schMap[schId].yesTrips++;schMap[schId].optKm+=optKm;}
      else{schMap[schId].cancelKm+=canKm;}
    }

    // FIX [MED-3]: Removed badge filter regex that was dropping valid badges
    if(drvBdg && drvBdg !== 'undefined'){
      if(!driverMap[drvBdg]) driverMap[drvBdg]={badge:drvBdg,name:drvNm,mobile:drvMob,schId:schId,optKm:0,diesel:0,kmpl:0,kmplC:0,deadKm:0};
      if(status==='YES') driverMap[drvBdg].optKm+=optKm;
    }
    if(cndBdg && cndBdg !== 'undefined'){
      if(!conductorMap[cndBdg]) conductorMap[cndBdg]={badge:cndBdg,name:cndNm,mobile:cndMob,schId:schId,cash:0,pass:0,students:0,dhaba:0,optKm:0};
      if(status==='YES') conductorMap[cndBdg].optKm+=optKm;
    }
  });

  (dieselRows||[]).forEach(function(r) {
    var dt     = formatDateVal(r[0]);
    var schId  = String(getCol(r,'schId') || '');
    var diesel = parseFloat(getCol(r,'diesel')) || 0;
    var uria   = parseFloat(getCol(r,'uria'))   || 0;
    var kmpl   = parseFloat(getCol(r,'kmpl'))   || 0;
    var deadKm = parseFloat(getCol(r,'deadKm')) || 0;
    var drvBdg = String(getCol(r,'driverBadge') || '').trim();
    var drvNm  = getCol(r,'driverName')  || '';
    var drvMob = getCol(r,'driverMobile')|| '';

    if(diesel>0){totalDiesel+=diesel;}
    if(uria>0){totalUria+=uria;}
    if(kmpl>0){totalKmplSum+=kmpl; kmplCount++;}
    if(trendMap[dt]){
      if(diesel>0) trendMap[dt].diesel+=diesel;
      if(kmpl>0){trendMap[dt].kmpl+=kmpl; trendMap[dt].kmplC++;}
    }
    if(schId && schMap[schId]){
      if(diesel>0) schMap[schId].diesel+=diesel;
      if(kmpl>0){schMap[schId].kmpl+=kmpl; schMap[schId].kmplC++;}
    }
    // FIX [MED-3]: Removed badge filter regex
    if(drvBdg && drvBdg !== 'undefined'){
      if(!driverMap[drvBdg]) driverMap[drvBdg]={badge:drvBdg,name:drvNm,mobile:drvMob,schId:schId,optKm:0,diesel:0,kmpl:0,kmplC:0,deadKm:0};
      if(diesel>0) driverMap[drvBdg].diesel+=diesel;
      if(kmpl>0){driverMap[drvBdg].kmpl+=kmpl; driverMap[drvBdg].kmplC++;}
      driverMap[drvBdg].deadKm+=deadKm;
    }
  });

  var totalCash=0, totalPass=0, totalStudents=0, totalDhaba=0;
  (svcRows||[]).forEach(function(sr) {
    var schId = String(getCol(sr,'schId')||'');
    var cash  = parseFloat(getCol(sr,'cash'))||0;
    var pass  = parseFloat(getCol(sr,'allPass'))||0;
    var stud  = parseFloat(getCol(sr,'students'))||0;
    var dhaba = parseFloat(getCol(sr,'dhaba'))||0;
    var dt    = formatDateVal(sr[0]);
    var cndBdg= String(getCol(sr,'condBadge')||'').trim();

    totalCash+=cash; totalPass+=pass; totalStudents+=stud; totalDhaba+=dhaba;
    if(schMap[schId]){ schMap[schId].cash+=cash; }
    if(trendMap[dt]){ trendMap[dt].cash+=cash; }
    // FIX [MED-3]: Removed badge filter regex
    if(cndBdg && cndBdg !== 'undefined'){
      if(!conductorMap[cndBdg]) conductorMap[cndBdg]={badge:cndBdg,name:String(getCol(sr,'condName')||''),mobile:String(getCol(sr,'condMobile')||''),schId:schId,cash:0,pass:0,students:0,dhaba:0,optKm:0};
      conductorMap[cndBdg].cash+=cash;
      conductorMap[cndBdg].pass+=pass;
      conductorMap[cndBdg].students+=stud;
      conductorMap[cndBdg].dhaba+=dhaba;
    }
  });

  // ── ONLY_CASH: add to totalCash + trendMap ──────────────
  (ocRows||[]).forEach(function(oc) {
    var dt   = formatDateVal(oc[0]);
    var cash = parseFloat(getCol(oc,'cash')) || 0;
    totalCash += cash;
    if (trendMap[dt]) trendMap[dt].cash += cash;
  });

  Object.values(schMap).forEach(function(s){
    s.epkm = s.optKm>0 ? parseFloat((s.cash/s.optKm).toFixed(2)) : 0;
  });
  Object.values(conductorMap).forEach(function(c){
    c.epkm = c.optKm>0 ? parseFloat((c.cash/c.optKm).toFixed(2)) : 0;
  });
  Object.values(trendMap).forEach(function(t){
    t.epkm = t.optKm>0 ? parseFloat((t.cash/t.optKm).toFixed(2)) : 0;
  });

  var overallEpkm = totalOpt>0 ? parseFloat((totalCash/totalOpt).toFixed(2)) : 0;

  var trend = Object.keys(trendMap).sort().map(function(dt){
    var t = trendMap[dt];
    return {label:t.label,sanKm:t.sanKm,optKm:t.optKm,cancelKm:t.cancelKm,
            diesel:t.diesel,kmpl:t.kmplC>0?parseFloat((t.kmpl/t.kmplC).toFixed(2)):0,
            cash:t.cash,epkm:t.epkm};
  });
  var schedules = Object.values(schMap).map(function(s){
    return {id:s.id,name:s.id,totalTrips:s.totalTrips,yesTrips:s.yesTrips,
            sanKm:s.sanKm,optKm:s.optKm,cancelKm:s.cancelKm,
            diesel:s.diesel,kmpl:s.kmplC>0?parseFloat((s.kmpl/s.kmplC).toFixed(2)):0,
            cash:s.cash,epkm:s.epkm};
  });
  var drivers = Object.values(driverMap).map(function(d){
    return {badge:d.badge,name:d.name,mobile:d.mobile,schId:d.schId,
            totalOptKm:d.optKm,totalDiesel:d.diesel,totalDeadKm:d.deadKm,
            avgKmpl:d.kmplC>0?parseFloat((d.kmpl/d.kmplC).toFixed(2)):0};
  });
  var conductors = Object.values(conductorMap).map(function(c){
    return {badge:c.badge,name:c.name,mobile:c.mobile,schId:c.schId,
            totalCash:c.cash,totalPass:c.pass,totalStudents:c.students,
            totalDhaba:c.dhaba,avgEpkm:c.epkm};
  });
  var topDrv = drivers.slice().sort(function(a,b){return b.avgKmpl-a.avgKmpl;}).slice(0,5);
  var lowDrv = drivers.slice().sort(function(a,b){return a.avgKmpl-b.avgKmpl;}).filter(function(d){return d.avgKmpl>0;}).slice(0,5);
  var topCnd = conductors.slice().sort(function(a,b){return b.avgEpkm-a.avgEpkm;}).slice(0,5);
  var lowCnd = conductors.slice().sort(function(a,b){return a.avgEpkm-b.avgEpkm;}).filter(function(c){return c.avgEpkm>0;}).slice(0,5);

  return {
    totalTrips:totalTrips, yesTrips:yesT, noTrips:noT, bdTrips:bdT,
    accidentTrips:accT, otherTrips:otherT,
    totalSanKm:totalSanKmFromMain, totalOptKm:totalOpt, totalCancelKm:totalCan,
    totalDiesel:totalDiesel, totalUria:totalUria,
    avgKmpl:kmplCount>0?parseFloat((totalKmplSum/kmplCount).toFixed(2)):0,
    totalCash:totalCash, avgEpkm:overallEpkm,
    totalPass:totalPass, totalStudents:totalStudents, totalDhaba:totalDhaba,
    trend:trend, schedules:schedules,
    topDrivers:topDrv, lowestDrivers:lowDrv,
    topConductors:topCnd, lowestConductors:lowCnd
  };
}

function makeSvcMap(svcRows) {
  var m = {};
  (svcRows||[]).forEach(function(sr){
    var key = formatDateVal(sr[0]) + '||' + String(getCol(sr,'schId')||'');
    m[key] = sr;
  });
  return m;
}
function svcVal(svcMap, date, schId, col) {
  var sr = svcMap[date + '||' + schId];
  return sr ? getCol(sr, col) : '';
}

// ═══════════════════════════════════════════════════════════════
// ALL REPORTS DATA
// ═══════════════════════════════════════════════════════════════
function safe(name, fn) {
  try { return fn(); }
  catch(e) { Logger.log('REPORT ERROR ['+name+']: '+e.message); return []; }
}

function getReportData(from, to) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return {_error:'Sheet not found: '+SHEET_NAME};
  detectColumns(sh);
  var rows = sh.getDataRange().getValues();
  var tcRows     = filterRows(rows, from, to, 'TC');
  var svcRows    = filterRows(rows, from, to, 'SERVICE');
  var dieselRows = filterRows(rows, from, to, 'DIESEL');
  var ocRows     = filterRows(rows, from, to, 'ONLY_CASH'); // Only Cash entries
  var dieselMap  = buildDieselMap(dieselRows);

  // Merge ONLY_CASH rows into svcRows for all cash-related reports
  // so that Only Cash amounts appear in cash totals, daily cash, DOR etc.
  var svcRowsWithOC = svcRows.concat(ocRows);

  Logger.log('getReportData: tcRows='+tcRows.length+' svcRows='+svcRows.length+' ocRows='+ocRows.length+' dieselRows='+dieselRows.length);

  return {
    _debug            : 'tcRows='+tcRows.length+' svcRows='+svcRows.length+' ocRows='+ocRows.length+' dieselRows='+dieselRows.length,
    from              : from,
    to                : to,
    dailyOpRows       : safe('buildDOR',                function(){ return buildDOR(tcRows, svcRowsWithOC, dieselMap, ocRows); }),
    tcRows            : safe('buildTCRows',             function(){ return buildTCRows(tcRows); }),
    tcDaily           : safe('buildTCDaily',            function(){ return buildTCDaily(tcRows); }),
    tc                : safe('buildTCSummary',          function(){ return buildTCSummary(tcRows); }),
    fuel              : safe('buildFuelSummary',        function(){ return buildFuelSummary(dieselRows); }),
    fuelRows          : safe('buildFuelRows',           function(){ return buildFuelRows(dieselRows); }),
    fuelDaily         : safe('buildFuelDaily',          function(){ return buildFuelDaily(dieselRows); }),
    cash              : safe('buildCashSummary',        function(){ return buildCashSummary(svcRowsWithOC); }),
    cashRows          : safe('buildCashRows',           function(){ return buildCashRows(svcRows); }),      // schedule-level rows — exclude OC (no schedule)
    cashDaily         : safe('buildCashDaily',          function(){ return buildCashDaily(svcRowsWithOC); }),
    onlyCashTotal     : safe('onlyCashTotal',           function(){ return ocRows.reduce(function(s,r){ return s+(parseFloat(getCol(r,'cash'))||0); },0); }),
    driverStats       : safe('buildDriverStats',        function(){ return buildDriverStats(tcRows, dieselMap); }),
    conductorStats    : safe('buildConductorStats',     function(){ return buildConductorStats(tcRows, svcRows); }),
    vehicleStats      : safe('buildVehicleStats',       function(){ return buildVehicleStats(tcRows, dieselMap); }),
    typeStats         : safe('buildTypeStats',          function(){ return buildTypeStats(tcRows, svcRows); }),
    cashTypeStats     : safe('buildCashTypeStats',      function(){ return buildCashTypeStats(tcRows, svcRows); }),
    dailyTypeStats    : safe('buildDailyTypeStats',     function(){ return buildDailyTypeStats(tcRows, svcRows); }),
    monthlyDriver     : safe('buildMonthlyDriver',      function(){ return buildMonthlyDriver(tcRows, dieselMap); }),
    monthlyConductor  : safe('buildMonthlyConductor',   function(){ return buildMonthlyConductor(tcRows, svcRows); }),
    monthlyVehicle    : safe('buildMonthlyVehicle',     function(){ return buildMonthlyVehicle(tcRows, dieselMap); }),
    weeklyStats       : safe('buildWeeklyStats',        function(){ return buildWeeklyStats(tcRows, svcRows, dieselMap, ocRows); }),
    cancellationRows  : safe('buildCancellationRows',   function(){ return buildCancellationRows(tcRows); }),
    ccRows            : safe('buildServiceTypeRows CC', function(){ return buildServiceTypeRows(tcRows, 'CC', svcRows); }),
    extRows           : safe('buildServiceTypeRows EXT',function(){ return buildServiceTypeRows(tcRows, 'EXT', svcRows); }),
    dhabaRows         : safe('buildDhabaRows',          function(){ return buildDhabaRows(svcRows); }),
    dailyServiceSummary: safe('buildDailyServiceSummary',function(){ return buildDailyServiceSummary(tcRows); }),
    dateWiseDetail    : safe('buildDateWiseDetail',     function(){ return buildDateWiseDetail(tcRows, svcRowsWithOC); }),
    top15Epkm         : safe('buildTopEpkm top',        function(){ return buildTopEpkm(tcRows, svcRows, 15, 'top'); }),
    low15Epkm         : safe('buildTopEpkm low',        function(){ return buildTopEpkm(tcRows, svcRows, 15, 'low'); }),
    weeklyDriver      : safe('buildWeeklyDriver',       function(){ return buildWeeklyDriver(tcRows, dieselMap); }),
    weeklyConductor   : safe('buildWeeklyConductor',    function(){ return buildWeeklyConductor(tcRows, svcRows); }),
    weeklyVehicle     : safe('buildWeeklyVehicle',      function(){ return buildWeeklyVehicle(tcRows, dieselMap); }),
    mergedDOR         : safe('buildMergedDOR',          function(){ return buildMergedDOR(tcRows, svcRowsWithOC); }),
    employeePerfDate  : safe('buildEmployeePerfDate',   function(){ return buildEmployeePerfDate(tcRows, svcRows, dieselMap); }),
    schWiseDay        : safe('buildSchWiseDayReport',   function(){ return buildSchWiseDayReport(tcRows, svcRows, dieselMap); }),
    utilisationReport : safe('buildUtilisationReport',  function(){ return buildUtilisationReport(tcRows); })
  };
}

function buildDieselMap(dieselRows) {
  var m = {};
  (dieselRows||[]).forEach(function(r){
    var key = formatDateVal(r[0]) + '|' + String(getCol(r,'schId')||r[2]||'');
    if (!m[key]) {
      m[key] = {
        diesel  : parseFloat(getCol(r,'diesel'))  || 0,
        uria    : parseFloat(getCol(r,'uria'))    || 0,
        kmpl    : parseFloat(getCol(r,'kmpl'))    || 0,  // stored KMPL — used directly
        deadKm  : parseFloat(getCol(r,'deadKm')) || 0,
        divKm   : parseFloat(getCol(r,'divKm'))  || 0,
        vehicle : getCol(r,'vehicle')      || '',
        vehType : getCol(r,'vehType')      || '',
        driverBadge  : getCol(r,'driverBadge')  || '',
        driverName   : getCol(r,'driverName')   || '',
        driverMobile : getCol(r,'driverMobile') || ''
      };
    }
  });
  return m;
}

function buildDOR(tcRows, svcRows, dieselMap, ocRows) {
  dieselMap = dieselMap || {};
  var sm = makeSvcMap(svcRows);
  var schCashed = {};
  var dayMap = {};
  tcRows.forEach(function(r) {
    var dt    = formatDateVal(r[0]);
    var schId = String(r[2] || '');
    var status= getCol(r,'status')     || '';
    var sanKm = parseFloat(getCol(r,'sanKm'))    || 0;
    var optKm = parseFloat(getCol(r,'optKm'))    || 0;
    var canKm = parseFloat(getCol(r,'cancelKm')) || 0;
    var _dmKey = dt + '|' + schId;
    var _dm    = dieselMap[_dmKey] || {};
    // FIX [BUG-DOR-DIESEL]: Only count diesel/uria/kmpl/deadKm ONCE per schedule per day.
    // Without this guard, a 2-trip schedule would add diesel twice (once per TC row).
    var _dieselAlreadyCounted = schCashed[_dmKey + '__d'];
    var diesel = (!_dieselAlreadyCounted && _dm.diesel) ? _dm.diesel : 0;
    var uria   = (!_dieselAlreadyCounted && _dm.uria)   ? _dm.uria   : 0;
    var kmpl   = (!_dieselAlreadyCounted && _dm.kmpl)   ? _dm.kmpl   : 0;
    var deadKm = (!_dieselAlreadyCounted && _dm.deadKm) ? _dm.deadKm : 0;
    if (!_dieselAlreadyCounted) schCashed[_dmKey + '__d'] = true;
    var type   = getCol(r,'tripType') || '';
    var cashKey = dt + '||' + schId;
    var cash=0, pass=0, stud=0, dhaba=0, epkm=0;
    if (!schCashed[cashKey]) {
      cash  = parseFloat(svcVal(sm,dt,schId,'cash'))     || 0;
      pass  = parseFloat(svcVal(sm,dt,schId,'allPass'))  || 0;
      stud  = parseFloat(svcVal(sm,dt,schId,'students')) || 0;
      dhaba = parseFloat(svcVal(sm,dt,schId,'dhaba'))    || 0;
      epkm  = parseFloat(svcVal(sm,dt,schId,'epkm'))     || 0;
      schCashed[cashKey] = true;
    }

    if (!dayMap[dt]) dayMap[dt] = {
      date:dt, schSet:new Set(), totalTrips:0, yesTrips:0, noTrips:0, bdTrips:0,
      sanKm:0, optKm:0, cancelKm:0, diesel:0, uria:0,
      kmplSum:0, kmplC:0, deadKm:0, cash:0, pass:0, students:0, dhaba:0,
      epkmSum:0, epkmC:0, typeMap:{}
    };
    var d = dayMap[dt];
    d.schSet.add(schId);
    d.totalTrips++; d.sanKm+=sanKm; d.optKm+=optKm; d.cancelKm+=canKm;
    if(status==='YES')d.yesTrips++;
    else if(status==='NO')d.noTrips++;
    else if(status==='BD')d.bdTrips++;
    if(diesel>0)d.diesel+=diesel;
    if(uria>0)d.uria+=uria;
    if(kmpl>0){d.kmplSum+=kmpl;d.kmplC++;}
    d.deadKm+=deadKm; d.cash+=cash; d.pass+=pass; d.students+=stud; d.dhaba+=dhaba;
    if(epkm>0){d.epkmSum+=epkm;d.epkmC++;}
    if(type) d.typeMap[type]=(d.typeMap[type]||0)+1;
  });

  // ── Add ONLY_CASH daily cash to day totals ───────────────
  (ocRows||[]).forEach(function(oc) {
    var dt   = formatDateVal(oc[0]);
    var cash = parseFloat(getCol(oc,'cash')) || 0;
    var pass = parseFloat(getCol(oc,'allPass')) || 0;
    if (!dayMap[dt]) return;  // only add to existing day entries
    dayMap[dt].cash += cash;
    dayMap[dt].pass += pass;
  });

  return Object.keys(dayMap).sort().map(function(dt){
    var d = dayMap[dt];
    var cr = d.totalTrips>0?((d.yesTrips/d.totalTrips)*100).toFixed(1):'0.0';
    var typeSummary = Object.keys(d.typeMap).map(function(t){return t+':'+d.typeMap[t];}).join(', ');
    return {
      date:dt, schCount:d.schSet.size, totalTrips:d.totalTrips,
      yesTrips:d.yesTrips, noTrips:d.noTrips, bdTrips:d.bdTrips,
      compRate:cr, sanKm:d.sanKm.toFixed(1), optKm:d.optKm.toFixed(1),
      cancelKm:d.cancelKm.toFixed(1), diesel:d.diesel.toFixed(1),
      uria:d.uria.toFixed(1), avgKmpl:d.kmplC>0?(d.kmplSum/d.kmplC).toFixed(2):'—',
      deadKm:d.deadKm.toFixed(1), cash:d.cash, pass:d.pass, students:d.students,
      dhaba:d.dhaba, avgEpkm:d.epkmC>0?(d.epkmSum/d.epkmC).toFixed(2):'—',
      typeBreakdown:typeSummary
    };
  });
}

function buildTypeStats(tcRows, svcRows) {
  // Sanctioned figures come from MAIN DATA via buildSanctionedMap()
  var sanData = buildSanctionedMap();
  var sanByType = sanData.byType;
  var sanBySchId = sanData.bySchId;

  // Cash from SERVICE rows
  var cashMap = {};
  (svcRows||[]).forEach(function(sr){
    var sid = String(sr[2]||'');
    cashMap[sid] = (cashMap[sid]||0) + (parseFloat(getCol(sr,'cash'))||0);
  });

  // Seed typeMap from MAIN DATA
  var typeMap = {};
  Object.keys(sanByType).forEach(function(type) {
    var s = sanByType[type];
    typeMap[type] = {
      type: type,
      sanSch: s.sanSch, sanTrips: s.sanTrips, sanKm: s.sanKm,
      optTrips: 0, optKm: 0, canKm: 0,
      optSchIds: new Set(), cash: 0
    };
  });

  // Accumulate operated/cancelled from TC rows
  tcRows.forEach(function(r) {
    var schId  = String(r[2] || '');
    var status = getCol(r,'status') || '';
    var optKm  = parseFloat(getCol(r,'optKm')) || 0;
    var canKm  = parseFloat(getCol(r,'cancelKm')) || 0;
    // Type from MAIN DATA is authoritative; fallback to TC row
    var type = (sanBySchId[schId] && sanBySchId[schId].type) ? sanBySchId[schId].type : (getCol(r,'tripType') || 'UNKNOWN');
    if (!typeMap[type]) typeMap[type] = {type:type,sanSch:0,sanTrips:0,sanKm:0,optTrips:0,optKm:0,canKm:0,optSchIds:new Set(),cash:0};
    if (status === 'YES') {
      typeMap[type].optTrips++;
      typeMap[type].optKm += optKm;
      typeMap[type].optSchIds.add(schId);
    } else {
      typeMap[type].canKm += canKm;
    }
  });

  // Add cash
  Object.keys(cashMap).forEach(function(schId){
    var type = (sanBySchId[schId] && sanBySchId[schId].type) ? sanBySchId[schId].type : 'UNKNOWN';
    if (typeMap[type]) typeMap[type].cash += cashMap[schId];
  });

  return Object.values(typeMap).sort(function(a,b){return b.sanTrips-a.sanTrips;}).map(function(t){
    return {
      type: t.type,
      sanSch: t.sanSch || 0,
      sanTrips: t.sanTrips || 0,
      sanKm: parseFloat((t.sanKm||0).toFixed(2)),
      optSch: t.optSchIds.size,
      optTrips: t.optTrips,
      optKm: parseFloat(t.optKm.toFixed(2)),
      canKm: parseFloat(t.canKm.toFixed(2)),
      cash: t.cash || 0
    };
  });
}

function buildCashTypeStats(tcRows, svcRows) {
  var sanBySchId = buildSanctionedMap().bySchId;
  var optKmBySchId = {};
  tcRows.forEach(function(r){
    if(getCol(r,'status')==='YES'){
      var sid = String(r[2]||'');
      optKmBySchId[sid] = (optKmBySchId[sid]||0) + (parseFloat(getCol(r,'optKm'))||0);
    }
  });
  var typeMap = {};
  (svcRows||[]).forEach(function(sr) {
    var schId = String(sr[2]||'');
    // Use MAIN DATA type as authoritative
    var type  = (sanBySchId[schId] && sanBySchId[schId].type) ? sanBySchId[schId].type : 'UNKNOWN';
    var cash  = parseFloat(getCol(sr,'cash'))||0;
    var allP  = parseFloat(getCol(sr,'allPass'))||0;
    var stud  = parseFloat(getCol(sr,'students'))||0;
    var dhaba = parseFloat(getCol(sr,'dhaba'))||0;
    var optKm = optKmBySchId[schId]||0;
    if(!typeMap[type]) typeMap[type]={type:type,income:0,cash:0,noPass:0,allPass:0,dhaba:0,km:0};
    typeMap[type].income+=cash; typeMap[type].cash+=cash;
    typeMap[type].noPass+=Math.max(0,allP-stud); typeMap[type].allPass+=allP;
    typeMap[type].dhaba+=dhaba; typeMap[type].km+=optKm;
  });
  return Object.values(typeMap).sort(function(a,b){return b.income-a.income;}).map(function(t){
    var epkm = t.km>0 ? parseFloat((t.cash/t.km).toFixed(2)) : 0;
    return {type:t.type,income:t.income,cash:t.cash,noPass:t.noPass,allPass:t.allPass,
            dhaba:t.dhaba,km:parseFloat(t.km.toFixed(2)),epkm:epkm};
  });
}

function buildDailyTypeStats(tcRows, svcRows) {
  // FIX [BUG-DAILY-TYPE]: Build sanData ONCE here, pass to each per-date call.
  // Previous code called buildTypeStats/buildCashTypeStats per date,
  // each of which called buildSanctionedMap() — for 30 days = 60 MAIN DATA reads → timeout.
  var dayMap = {};
  var dayMapSvc = {};
  tcRows.forEach(function(r) {
    var dt = formatDateVal(r[0]);
    if (!dayMap[dt]) dayMap[dt] = [];
    dayMap[dt].push(r);
  });
  (svcRows||[]).forEach(function(r) {
    var dt = formatDateVal(r[0]);
    if (!dayMapSvc[dt]) dayMapSvc[dt] = [];
    dayMapSvc[dt].push(r);
  });
  return Object.keys(dayMap).sort().map(function(dt) {
    return {date:dt, tcTypes:buildTypeStats(dayMap[dt], dayMapSvc[dt]||[]), cashTypes:buildCashTypeStats(dayMap[dt], dayMapSvc[dt]||[])};
  });
}

function buildTCRows(rows) {
  return rows.map(function(r){
    return {date:formatDateVal(r[0]),schId:r[2],route:r[4],
            type:getCol(r,'tripType')||'',sanKm:parseFloat(getCol(r,'sanKm'))||0,optKm:parseFloat(getCol(r,'optKm'))||0,
            cancelKm:parseFloat(getCol(r,'cancelKm'))||0,status:getCol(r,'status')||''};
  });
}
function buildTCDaily(rows) {
  var m={};
  rows.forEach(function(r){
    var dt=formatDateVal(r[0]);
    if(!m[dt])m[dt]={total:0,yes:0};
    m[dt].total++; if(getCol(r,'status')==='YES')m[dt].yes++;
  });
  return Object.keys(m).sort().map(function(dt){
    return {date:dt.substring(5),rate:m[dt].total>0?(m[dt].yes/m[dt].total*100).toFixed(1):0};
  });
}
function buildTCSummary(rows) {
  var t=0,yes=0,no=0,bd=0,acc=0,oth=0,opt=0,can=0;
  rows.forEach(function(r){
    t++; var s=getCol(r,'status');
    if(s==='YES')yes++; else if(s==='NO')no++; else if(s==='BD')bd++;
    else if(s==='ACCIDENT')acc++; else oth++;
    opt+=parseFloat(getCol(r,'optKm'))||0; can+=parseFloat(getCol(r,'cancelKm'))||0;
  });
  return {total:t,yes:yes,no:no,bd:bd,accident:acc,other:oth,optKm:opt,cancelKm:can};
}

// FIX [HIGH-4]: buildFuelRows now uses stored row[17] (kmpl) directly
// instead of recalculating, which was wrong for multi-schedule fuel entries.
function buildFuelRows(rows) {
  var seen={};
  var result=[];
  rows.forEach(function(r){
    var key=formatDateVal(r[0])+'|'+r[2];
    if(!seen[key]){
      seen[key]=true;
      var optKm   = parseFloat(getCol(r,'optKm'))  || 0;
      var deadKm  = parseFloat(getCol(r,'deadKm')) || 0;
      var divKm   = parseFloat(getCol(r,'divKm'))  || 0;
      var diesel  = parseFloat(getCol(r,'diesel')) || 0;
      var uria    = parseFloat(getCol(r,'uria'))   || 0;
      // FIX: Use stored KMPL from row[17] — do not recalculate.
      // Recalculating was wrong for multi-sch entries where secondary rows
      // have schOpt (not totalOptKm), giving an inflated/deflated KMPL.
      var kmpl    = parseFloat(getCol(r,'kmpl'))   || 0;
      var totalKm = optKm + deadKm + divKm;
      result.push({
        date    : formatDateVal(r[0]),
        schId   : r[2],
        schNo   : getCol(r,'schNo') || r[3] || '',
        vehicle : getCol(r,'vehicle')      || '',
        driver  : getCol(r,'driverName')   || '',
        mobile  : getCol(r,'driverMobile') || '',
        optKm   : optKm,
        deadKm  : deadKm,
        divKm   : divKm,
        totalKm : parseFloat(totalKm.toFixed(1)),
        diesel  : diesel,
        uria    : uria,
        kmpl    : kmpl   // stored value, not recalculated
      });
    }
  });
  return result.sort(function(a,b){
    return a.date.localeCompare(b.date) || String(a.schId).localeCompare(String(b.schId));
  });
}

function buildFuelSummary(rows) {
  var td=0,tu=0,to=0,tDead=0,tDiv=0,tKmpl=0,tKmplC=0;
  var seen={};
  rows.forEach(function(r){
    var key=formatDateVal(r[0])+'|'+r[2];
    if(!seen[key]){
      seen[key]=true;
      td    += parseFloat(getCol(r,'diesel'))  || 0;
      tu    += parseFloat(getCol(r,'uria'))    || 0;
      to    += parseFloat(getCol(r,'optKm'))   || 0;
      tDead += parseFloat(getCol(r,'deadKm'))  || 0;
      tDiv  += parseFloat(getCol(r,'divKm'))   || 0;
      // Use stored kmpl for average
      var k = parseFloat(getCol(r,'kmpl')) || 0;
      if (k > 0) { tKmpl += k; tKmplC++; }
    }
  });
  var totalKm = to + tDead + tDiv;
  // Prefer stored per-entry KMPL average; fallback to bulk calculation
  var avgKmpl = tKmplC > 0
    ? parseFloat((tKmpl / tKmplC).toFixed(2))
    : (td > 0 && totalKm > 0 ? parseFloat((totalKm / td).toFixed(2)) : 0);
  return {
    totalDiesel : parseFloat(td.toFixed(1)),
    totalUria   : parseFloat(tu.toFixed(1)),
    avgKmpl     : avgKmpl,
    totalOptKm  : parseFloat(to.toFixed(1)),
    totalDeadKm : parseFloat(tDead.toFixed(1)),
    totalDivKm  : parseFloat(tDiv.toFixed(1)),
    totalKm     : parseFloat(totalKm.toFixed(1))
  };
}

function buildFuelDaily(rows) {
  var m={};
  rows.forEach(function(r){
    var dt=formatDateVal(r[0]);
    if(!m[dt])m[dt]={diesel:0,uria:0,optKm:0,deadKm:0,divKm:0,kmpl:0,kmplC:0};
    m[dt].diesel += parseFloat(getCol(r,'diesel')) || 0;
    m[dt].uria   += parseFloat(getCol(r,'uria'))   || 0;
    m[dt].optKm  += parseFloat(getCol(r,'optKm'))  || 0;
    m[dt].deadKm += parseFloat(getCol(r,'deadKm')) || 0;
    m[dt].divKm  += parseFloat(getCol(r,'divKm'))  || 0;
    var k = parseFloat(getCol(r,'kmpl')) || 0;
    if (k > 0) { m[dt].kmpl += k; m[dt].kmplC++; }
  });
  return Object.keys(m).sort().map(function(dt){
    var d = m[dt];
    var totalKm = d.optKm + d.deadKm + d.divKm;
    // Use stored KMPL average for the day
    var kmpl = d.kmplC > 0 ? parseFloat((d.kmpl / d.kmplC).toFixed(2)) : 0;
    return {date:dt.substring(5),diesel:d.diesel,uria:d.uria,
            optKm:d.optKm,deadKm:d.deadKm,divKm:d.divKm,totalKm:totalKm,kmpl:kmpl};
  });
}

function buildCashRows(rows) {
  var seen={};var result=[];
  rows.forEach(function(r){
    var key=formatDateVal(r[0])+'|'+r[2];
    if(!seen[key] && getCol(r,'condBadge')){
      seen[key]=true;
      result.push({date:formatDateVal(r[0]),schId:r[2],
        conductor:getCol(r,'condName')||'',mobile:getCol(r,'condMobile')||'',bagNo:getCol(r,'bagNo')||'',
        cash:parseFloat(getCol(r,'cash'))||0,pass:parseFloat(getCol(r,'allPass'))||0,
        students:parseFloat(getCol(r,'students'))||0,dhaba:parseFloat(getCol(r,'dhaba'))||0,
        epkm:getCol(r,'epkm')||'—'});
    }
  });
  return result;
}
function buildCashSummary(rows) {
  var tc=0,tp=0,ts=0,td=0,te=0,tec=0;
  var seen={};
  rows.forEach(function(r){
    var key=formatDateVal(r[0])+'|'+r[2];
    if(!seen[key]){
      seen[key]=true;
      tc+=parseFloat(getCol(r,'cash'))||0; tp+=parseFloat(getCol(r,'allPass'))||0;
      ts+=parseFloat(getCol(r,'students'))||0; td+=parseFloat(getCol(r,'dhaba'))||0;
      var e=parseFloat(getCol(r,'epkm'))||0; if(e>0){te+=e;tec++;}
    }
  });
  return {totalCash:tc,totalPass:tp,totalStudents:ts,totalDhaba:td,avgEpkm:tec>0?te/tec:0};
}
function buildCashDaily(rows) {
  var m={};
  rows.forEach(function(r){
    var dt=formatDateVal(r[0]);
    if(!m[dt])m[dt]={cash:0,epkm:0,epkmC:0};
    m[dt].cash+=parseFloat(getCol(r,'cash'))||0;
    var e=parseFloat(getCol(r,'epkm'))||0; if(e>0){m[dt].epkm+=e;m[dt].epkmC++;}
  });
  return Object.keys(m).sort().map(function(dt){
    return {date:dt.substring(5),cash:m[dt].cash,epkm:m[dt].epkmC>0?parseFloat((m[dt].epkm/m[dt].epkmC).toFixed(2)):0};
  });
}

// FIX [MED-3]: Removed /^\d{4}-/ badge filter throughout driver/conductor stats
function buildDriverStats(rows, dieselMap) {
  var m={};
  rows.forEach(function(r){
    var badge=String(getCol(r,'driverBadge')||'').trim();
    // Only skip truly empty or obviously invalid badges
    if(!badge || badge==='' || badge==='undefined' || badge.length>30) return;
    if(!m[badge])m[badge]={badge:badge,name:getCol(r,'driverName')||'',mobile:getCol(r,'driverMobile')||'',
      schIds:new Set(),days:new Set(),optKm:0,diesel:0,deadKm:0,
      kmpl:0,kmplC:0,vehSet:new Set()};
    var _dKey=formatDateVal(r[0])+'|'+String(r[2]||''); var _dm=(dieselMap||{})[_dKey]||{};
    m[badge].schIds.add(r[2]); m[badge].days.add(formatDateVal(r[0]));
    m[badge].optKm+=parseFloat(getCol(r,'optKm'))||0;
    // FIX [BUG-DRV-DIESEL]: de-dup diesel per date+schId — multi-trip schedules must not double-count
    if(!m[badge]._dSeen) m[badge]._dSeen={};
    if(!m[badge]._dSeen[_dKey]){
      m[badge]._dSeen[_dKey]=1;
      m[badge].diesel+=_dm.diesel||0;
      m[badge].deadKm+=_dm.deadKm||0;
      if(_dm.vehicle) m[badge].vehSet.add(_dm.vehicle); else if(getCol(r,'vehicle')) m[badge].vehSet.add(getCol(r,'vehicle'));
      var k=_dm.kmpl||0; if(k>0){m[badge].kmpl+=k;m[badge].kmplC++;}
    }
  });
  return Object.values(m).map(function(d){
    return {badge:d.badge,name:d.name,mobile:d.mobile,
            schIds:[...d.schIds].join(', '),days:d.days.size,
            vehicles:[...d.vehSet].join(', '),
            totalOptKm:parseFloat(d.optKm.toFixed(1)),
            totalDiesel:parseFloat(d.diesel.toFixed(1)),
            totalDeadKm:parseFloat(d.deadKm.toFixed(1)),
            avgKmpl:d.kmplC>0?parseFloat((d.kmpl/d.kmplC).toFixed(2)):0};
  }).sort(function(a,b){return b.avgKmpl-a.avgKmpl;});
}

function buildConductorStats(rows, svcRows) {
  var sm = makeSvcMap(svcRows);
  var counted = {};
  var m={};
  rows.forEach(function(r){
    var badge=String(getCol(r,'condBadge')||'').trim();
    // FIX [MED-3]: Only skip truly invalid badges
    if(!badge || badge==='' || badge==='undefined' || badge.length>30) return;
    var dt=formatDateVal(r[0]); var schId=String(r[2]||'');
    if(!m[badge])m[badge]={badge:badge,name:getCol(r,'condName')||'',mobile:getCol(r,'condMobile')||'',
      schIds:new Set(),days:new Set(),cash:0,pass:0,students:0,dhaba:0,epkm:0,epkmC:0};
    m[badge].schIds.add(schId); m[badge].days.add(dt);
    var ck=dt+'||'+schId;
    if(!counted[ck]){
      m[badge].cash+=parseFloat(svcVal(sm,dt,schId,'cash'))||0;
      m[badge].pass+=parseFloat(svcVal(sm,dt,schId,'allPass'))||0;
      m[badge].students+=parseFloat(svcVal(sm,dt,schId,'students'))||0;
      m[badge].dhaba+=parseFloat(svcVal(sm,dt,schId,'dhaba'))||0;
      var e=parseFloat(svcVal(sm,dt,schId,'epkm'))||0; if(e>0){m[badge].epkm+=e;m[badge].epkmC++;}
      counted[ck]=true;
    }
  });
  return Object.values(m).map(function(c){
    return {badge:c.badge,name:c.name,mobile:c.mobile,
            schIds:[...c.schIds].join(', '),days:c.days.size,
            totalCash:c.cash,totalPass:c.pass,totalStudents:c.students,
            totalDhaba:c.dhaba,avgEpkm:c.epkmC>0?parseFloat((c.epkm/c.epkmC).toFixed(2)):0};
  }).sort(function(a,b){return b.avgEpkm-a.avgEpkm;});
}

function buildVehicleStats(rows, dieselMap) {
  var m={}; var _sd={};
  rows.forEach(function(r){
    var veh=String(getCol(r,'vehicle')||''); if(!veh||veh===''||veh==='DEPOT')return;
    if(!m[veh])m[veh]={vehicle:veh,type:getCol(r,'vehType')||'',driverSet:new Set(),
      days:new Set(),optKm:0,diesel:0,deadKm:0,kmpl:0,kmplC:0};
    var _dKey=formatDateVal(r[0])+'|'+String(r[2]||''); var _dm=(dieselMap||{})[_dKey]||{};
    m[veh].days.add(formatDateVal(r[0]));
    if(getCol(r,'driverBadge')) m[veh].driverSet.add(String(getCol(r,'driverBadge')));
    m[veh].optKm+=parseFloat(getCol(r,'optKm'))||0;
    if(!_sd[veh+'|'+_dKey]){_sd[veh+'|'+_dKey]=1;
      m[veh].diesel+=_dm.diesel||0; m[veh].deadKm+=_dm.deadKm||0;
      var k=_dm.kmpl||0; if(k>0){m[veh].kmpl+=k;m[veh].kmplC++;}
    }
  });
  return Object.values(m).map(function(v){
    return {vehicle:v.vehicle,type:v.type,days:v.days.size,
            drivers:[...v.driverSet].join(', '),
            totalOptKm:parseFloat(v.optKm.toFixed(1)),
            totalDiesel:parseFloat(v.diesel.toFixed(1)),
            totalDeadKm:parseFloat(v.deadKm.toFixed(1)),
            avgKmpl:v.kmplC>0?parseFloat((v.kmpl/v.kmplC).toFixed(2)):0};
  }).sort(function(a,b){return b.avgKmpl-a.avgKmpl;});
}

function buildMonthlyDriver(rows, dieselMap) {
  var m={};
  rows.forEach(function(r){
    var badge=String(getCol(r,'driverBadge')||''); if(!badge||badge==='undefined')return;
    var month=formatMonthVal(r[0]);
    var key=month+'|'+badge;
    if(!m[key])m[key]={month:month,badge:badge,name:getCol(r,'driverName')||'',
      optKm:0,diesel:0,kmpl:0,kmplC:0,deadKm:0,trips:0,yes:0};
    var _dKey=formatDateVal(r[0])+'|'+String(r[2]||''); var _dm=(dieselMap||{})[_dKey]||{};
    m[key].trips++; m[key].optKm+=parseFloat(getCol(r,'optKm'))||0;
    // FIX [BUG-MDRV-DIESEL]: de-dup diesel per date+schId
    if(!m[key]._dSeen) m[key]._dSeen={};
    if(!m[key]._dSeen[_dKey]){
      m[key]._dSeen[_dKey]=1;
      m[key].diesel+=_dm.diesel||0; m[key].deadKm+=_dm.deadKm||0;
      var k=_dm.kmpl||0; if(k>0){m[key].kmpl+=k;m[key].kmplC++;}
    }
    if(getCol(r,'status')==='YES')m[key].yes++;
  });
  return Object.values(m).map(function(x){
    return {month:x.month,badge:x.badge,name:x.name,trips:x.trips,
            yesTrips:x.yes,optKm:parseFloat(x.optKm.toFixed(1)),diesel:parseFloat(x.diesel.toFixed(1)),
            deadKm:parseFloat(x.deadKm.toFixed(1)),avgKmpl:x.kmplC>0?(x.kmpl/x.kmplC).toFixed(2):'—'};
  }).sort(function(a,b){return (a.month+a.badge).localeCompare(b.month+b.badge);});
}

function buildMonthlyConductor(rows, svcRows) {
  var _sm=makeSvcMap(svcRows); var _cc={};
  var m={};
  rows.forEach(function(r){
    var badge=String(getCol(r,'condBadge')||''); if(!badge||badge==='undefined')return;
    var month=formatMonthVal(r[0]);
    var key=month+'|'+badge;
    if(!m[key])m[key]={month:month,badge:badge,name:getCol(r,'condName')||'',
      cash:0,pass:0,students:0,dhaba:0,epkm:0,epkmC:0,days:new Set()};
    m[key].days.add(formatDateVal(r[0]));
    var _dt=formatDateVal(r[0]),_sid=String(r[2]||''),_ck=_dt+'||'+_sid;
    if(!_cc[_ck]){m[key].cash+=parseFloat(svcVal(_sm,_dt,_sid,'cash'))||0; m[key].pass+=parseFloat(svcVal(_sm,_dt,_sid,'allPass'))||0;
    m[key].students+=parseFloat(svcVal(_sm,_dt,_sid,'students'))||0; m[key].dhaba+=parseFloat(svcVal(_sm,_dt,_sid,'dhaba'))||0; _cc[_ck]=1;}
    var e=parseFloat(svcVal(_sm,_dt,_sid,'epkm'))||0; if(e>0){m[key].epkm+=e;m[key].epkmC++;}
  });
  return Object.values(m).map(function(x){
    return {month:x.month,badge:x.badge,name:x.name,days:x.days.size,
            cash:x.cash,pass:x.pass,students:x.students,dhaba:x.dhaba,
            avgEpkm:x.epkmC>0?(x.epkm/x.epkmC).toFixed(2):'—'};
  }).sort(function(a,b){return (a.month+a.badge).localeCompare(b.month+b.badge);});
}

function buildMonthlyVehicle(rows, dieselMap) {
  var m={};var _sd={};
  rows.forEach(function(r){
    var veh=String(getCol(r,'vehicle')||''); if(!veh||veh==='DEPOT')return;
    var month=formatMonthVal(r[0]);
    var key=month+'|'+veh;
    if(!m[key])m[key]={month:month,vehicle:veh,type:getCol(r,'vehType')||'',
      optKm:0,diesel:0,deadKm:0,kmpl:0,kmplC:0,days:new Set()};
    var _dKey=formatDateVal(r[0])+'|'+String(r[2]||''); var _dm=(dieselMap||{})[_dKey]||{};
    m[key].days.add(formatDateVal(r[0]));
    m[key].optKm+=parseFloat(getCol(r,'optKm'))||0;
    if(!_sd[key+'|'+_dKey]){_sd[key+'|'+_dKey]=1;
      m[key].diesel+=_dm.diesel||0; m[key].deadKm+=_dm.deadKm||0;
      var k=_dm.kmpl||0; if(k>0){m[key].kmpl+=k;m[key].kmplC++;}
    }
  });
  return Object.values(m).map(function(x){
    return {month:x.month,vehicle:x.vehicle,type:x.type,days:x.days.size,
            optKm:parseFloat(x.optKm.toFixed(1)),diesel:parseFloat(x.diesel.toFixed(1)),
            deadKm:parseFloat(x.deadKm.toFixed(1)),avgKmpl:x.kmplC>0?(x.kmpl/x.kmplC).toFixed(2):'—'};
  }).sort(function(a,b){return (a.month+a.vehicle).localeCompare(b.month+b.vehicle);});
}

function buildWeeklyStats(rows, svcRows, dieselMap, ocRows) {
  var _sm=makeSvcMap(svcRows); var _cc={}; var m={}; var _sd={};
  rows.forEach(function(r){
    var dt=new Date(formatDateVal(r[0]));
    var day=dt.getDay(); var diff=dt.getDate()-(day===0?6:day-1);
    var mon=new Date(dt.setDate(diff));
    var wk=fmtDate(mon);
    if(!m[wk])m[wk]={week:wk,trips:0,yes:0,no:0,optKm:0,cancelKm:0,
      diesel:0,kmpl:0,kmplC:0,cash:0,epkm:0,epkmC:0,pass:0};
    var _dKey=formatDateVal(r[0])+'|'+String(r[2]||''); var _dm=(dieselMap||{})[_dKey]||{};
    m[wk].trips++; m[wk].optKm+=parseFloat(getCol(r,'optKm'))||0; m[wk].cancelKm+=parseFloat(getCol(r,'cancelKm'))||0;
    if(getCol(r,'status')==='YES')m[wk].yes++; else m[wk].no++;
    if(!_sd[wk+'|'+_dKey]){_sd[wk+'|'+_dKey]=1;
      if(_dm.diesel>0) m[wk].diesel+=_dm.diesel||0;
      var k=_dm.kmpl||0; if(k>0){m[wk].kmpl+=k;m[wk].kmplC++;}
    }
    var _dt=formatDateVal(r[0]),_sid=String(r[2]||''),_ck=_dt+'||'+_sid;
    if(!_cc[_ck]){m[wk].cash+=parseFloat(svcVal(_sm,_dt,_sid,'cash'))||0; m[wk].pass+=parseFloat(svcVal(_sm,_dt,_sid,'allPass'))||0; _cc[_ck]=1;}
    var e=parseFloat(svcVal(_sm,formatDateVal(r[0]),String(r[2]||''),'epkm'))||0; if(e>0){m[wk].epkm+=e;m[wk].epkmC++;}
  });
  // FIX [BUG-WEEKLY-OC]: Add Only Cash entries to the correct week
  (ocRows||[]).forEach(function(oc){
    var dt=new Date(formatDateVal(oc[0]));
    var day=dt.getDay(); var diff=dt.getDate()-(day===0?6:day-1);
    var mon=new Date(dt.setDate(diff));
    var wk=fmtDate(mon);
    if(m[wk]){ m[wk].cash+=parseFloat(getCol(oc,'cash'))||0; }
  });
  return Object.values(m).sort(function(a,b){return a.week.localeCompare(b.week);}).map(function(w){
    return {week:w.week,trips:w.trips,yesTrips:w.yes,noTrips:w.no,
            optKm:parseFloat(w.optKm.toFixed(1)),cancelKm:parseFloat(w.cancelKm.toFixed(1)),
            diesel:parseFloat(w.diesel.toFixed(1)),avgKmpl:w.kmplC>0?(w.kmpl/w.kmplC).toFixed(2):'—',
            cash:w.cash,avgEpkm:w.epkmC>0?(w.epkm/w.epkmC).toFixed(2):'—',pass:w.pass,
            compRate:w.trips>0?(w.yes/w.trips*100).toFixed(1):'0.0'};
  });
}

// ── DATE HELPERS ─────────────────────────────────────────────
function formatDateVal(v) {
  if (!v) return '';
  if (v instanceof Date) {
    var y = v.getFullYear();
    var m = v.getMonth() + 1;
    var d = v.getDate();
    return y + '-' + (m < 10 ? '0' : '') + m + '-' + (d < 10 ? '0' : '') + d;
  }
  return String(v).substring(0, 10);
}
function formatMonthVal(v) {
  return formatDateVal(v).substring(0, 7);
}

// ═══════════════════════════════════════════════════════════════
// INDIVIDUAL PERFORMANCE REPORT
// ═══════════════════════════════════════════════════════════════
function getIndividualReport(params) {
  try {
    var badge    = String(params.badge    || '').trim().toUpperCase();
    var vehicle  = String(params.vehicle  || '').trim().toUpperCase();
    var from     = String(params.from     || '');
    var to       = String(params.to       || '');
    var type     = String(params.type     || 'driver');

    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { error: 'COMBINE DATA sheet not found' };
    var rows = sh.getDataRange().getValues();
    detectColumns(sh);

    // Build svcMap: date||schId -> SERVICE row
    var _indivSvcMap = {};
    for (var si = 1; si < rows.length; si++) {
      if (String(rows[si][1]) === 'SERVICE') {
        var _sk = formatDateVal(rows[si][0]) + '||' + String(rows[si][2]||'');
        _indivSvcMap[_sk] = rows[si];
      }
    }

    var trips = [];
    // Track which date+schId cash already counted (one cash entry per schedule per date)
    var _cashSeen = {};
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0] || r[1] !== 'TC') continue;
      var dt = formatDateVal(r[0]);
      if (from && dt < from) continue;
      if (to   && dt > to)   continue;

      var dBadge = String(getCol(r,'driverBadge') || '').trim().toUpperCase();
      var cBadge = String(getCol(r,'condBadge') || '').trim().toUpperCase();
      var vNo    = String(getCol(r,'vehicle') || '').trim().toUpperCase();

      var match = false;
      if (type === 'driver'    && badge   && dBadge === badge)   match = true;
      if (type === 'conductor' && badge   && cBadge === badge)   match = true;
      if (type === 'vehicle'   && vehicle && vNo    === vehicle) match = true;

      if (!match) continue;

      trips.push({
        date        : dt,
        schId       : String(r[2]  || ''),
        schNo       : String(r[3]  || ''),
        route       : String(r[4]  || ''),
        tripType    : String(getCol(r,'tripType')  || ''),
        sanKm       : parseFloat(getCol(r,'sanKm'))  || 0,
        optKm       : parseFloat(getCol(r,'optKm'))  || 0,
        cancelKm    : parseFloat(getCol(r,'cancelKm'))  || 0,
        status      : String(getCol(r,'status')  || ''),
        vehicle     : String(getCol(r,'vehicle') || ''),
        vehicleType : String(getCol(r,'vehType') || ''),
        driverBadge : String(getCol(r,'driverBadge') || ''),
        driverName  : String(getCol(r,'driverName') || ''),
        driverMobile: String(getCol(r,'driverMobile') || ''),
        diesel      : parseFloat(getCol(r,'diesel')) || 0,
        kmpl        : parseFloat(getCol(r,'kmpl')) || 0,
        deadKm      : parseFloat(getCol(r,'deadKm')) || 0,
        condBadge   : String(getCol(r,'condBadge') || ''),
        condName    : String(getCol(r,'condName') || ''),
        condMobile  : String(getCol(r,'condMobile') || ''),
        cash        : (function(){ var _ck=dt+'||'+String(r[2]||''); if(!_cashSeen[_ck]){_cashSeen[_ck]=true; var sr=_indivSvcMap[_ck]; return sr?parseFloat(getCol(sr,'cash'))||0:0;} return 0; })(),
        allPass     : (function(){ var _ck=dt+'||'+String(r[2]||''); var sr=_indivSvcMap[_ck]; return sr?parseFloat(getCol(sr,'allPass'))||0:0; })(),
        students    : (function(){ var _ck=dt+'||'+String(r[2]||''); var sr=_indivSvcMap[_ck]; return sr?parseFloat(getCol(sr,'students'))||0:0; })(),
        dhaba       : (function(){ var _ck=dt+'||'+String(r[2]||''); var sr=_indivSvcMap[_ck]; return sr?parseFloat(getCol(sr,'dhaba'))||0:0; })(),
        epkm        : (function(){ var _ck=dt+'||'+String(r[2]||''); var sr=_indivSvcMap[_ck]; return sr?parseFloat(getCol(sr,'epkm'))||0:0; })()
      });
    }

    if (!trips.length) return { found: false, trips: [], summary: {}, daily: [] };

    trips.sort(function(a,b){ return a.date.localeCompare(b.date) || a.schNo.localeCompare(b.schNo); });
    var summary = buildIndivSummary(trips, type);
    var dailyMap = {};
    trips.forEach(function(t) {
      if (!dailyMap[t.date]) dailyMap[t.date] = [];
      dailyMap[t.date].push(t);
    });
    var daily = Object.keys(dailyMap).sort().map(function(dt) {
      return { date: dt, summary: buildIndivSummary(dailyMap[dt], type), trips: dailyMap[dt] };
    });
    var first = trips[0];
    var identity = {
      badge  : type === 'vehicle' ? vehicle : badge,
      name   : type === 'driver' ? first.driverName : type === 'conductor' ? first.condName : '',
      mobile : type === 'driver' ? first.driverMobile : type === 'conductor' ? first.condMobile : '',
      vehicle: type === 'vehicle' ? vehicle : first.vehicle,
      type   : type
    };

    return { found: true, identity: identity, summary: summary, daily: daily, trips: trips };
  } catch(e) {
    return { error: e.message };
  }
}

function buildIndivSummary(trips, type) {
  var s = {
    totalTrips:0, yesTrips:0, noTrips:0, bdTrips:0,
    sanKm:0, optKm:0, cancelKm:0, deadKm:0,
    cash:0, allPass:0, students:0, dhaba:0,
    diesel:0, kmpl:0, kmplC:0, epkm:0, epkmC:0,
    uniqueDates: new Set(), uniqueSchs: new Set(), vehicles: new Set()
  };
  trips.forEach(function(t) {
    s.totalTrips++;
    if (t.status==='YES')  s.yesTrips++;
    if (t.status==='NO')   s.noTrips++;
    if (t.status==='BD')   s.bdTrips++;
    s.sanKm    += t.sanKm;
    s.optKm    += t.optKm;
    s.cancelKm += t.cancelKm;
    s.deadKm   += t.deadKm;
    s.cash     += t.cash;
    s.allPass  += t.allPass;
    s.students += t.students;
    s.dhaba    += t.dhaba;
    s.diesel   += t.diesel;
    if (t.kmpl>0) { s.kmpl+=t.kmpl; s.kmplC++; }
    if (t.epkm>0) { s.epkm+=t.epkm; s.epkmC++; }
    s.uniqueDates.add(t.date);
    s.uniqueSchs.add(t.schNo || t.schId);
    if (t.vehicle) s.vehicles.add(t.vehicle);
  });
  var avgKmpl = s.kmplC>0 ? parseFloat((s.kmpl/s.kmplC).toFixed(2)) : 0;
  var avgEpkm = s.epkmC>0 ? parseFloat((s.epkm/s.epkmC).toFixed(2)) :
                (s.optKm>0 ? parseFloat((s.cash/s.optKm).toFixed(2)) : 0);
  var compRate = s.totalTrips>0 ? parseFloat((s.yesTrips/s.totalTrips*100).toFixed(1)) : 0;
  return {
    totalTrips : s.totalTrips,
    yesTrips   : s.yesTrips,
    noTrips    : s.noTrips,
    bdTrips    : s.bdTrips,
    compRate   : compRate,
    days       : s.uniqueDates.size,
    schedules  : s.uniqueSchs.size,
    vehicles   : [...s.vehicles].join(', '),
    sanKm      : parseFloat(s.sanKm.toFixed(1)),
    optKm      : parseFloat(s.optKm.toFixed(1)),
    cancelKm   : parseFloat(s.cancelKm.toFixed(1)),
    deadKm     : parseFloat(s.deadKm.toFixed(1)),
    cash       : Math.round(s.cash),
    allPass    : Math.round(s.allPass),
    students   : Math.round(s.students),
    dhaba      : Math.round(s.dhaba),
    diesel     : parseFloat(s.diesel.toFixed(1)),
    avgKmpl    : avgKmpl,
    avgEpkm    : avgEpkm
  };
}

// ═══════════════════════════════════════════════════════════════
// RANGE-WISE REPORT
// ═══════════════════════════════════════════════════════════════
function getRangeReport(params) {
  try {
    var from    = String(params.from    || '');
    var to      = String(params.to      || '');
    var minKmpl = parseFloat(params.minKmpl)  || 0;
    var maxKmpl = parseFloat(params.maxKmpl)  || 999;
    var minEpkm = parseFloat(params.minEpkm)  || 0;
    var maxEpkm = parseFloat(params.maxEpkm)  || 999;
    var mode    = String(params.mode || 'both');

    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { error: 'COMBINE DATA not found' };
    var rows = sh.getDataRange().getValues();
    detectColumns(sh);

    // Build svcMap once: date||schId -> SERVICE row (cash read only from SERVICE rows)
    var _svcMap = {};
    for (var si = 1; si < rows.length; si++) {
      if (String(rows[si][1]) === 'SERVICE') {
        var _sk = formatDateVal(rows[si][0]) + '||' + String(rows[si][2]||'');
        _svcMap[_sk] = rows[si];
      }
    }
    function _svcGet(dt, schId, col) {
      var sr = _svcMap[dt+'||'+schId];
      return sr ? getCol(sr, col) : '';
    }

    var driverMap = {};
    var condMap   = {};
    var _condCashCounted = {}; // prevent double-counting cash per date+schId

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0] || r[1] !== 'TC') continue;
      var dt = formatDateVal(r[0]);
      if (from && dt < from) continue;
      if (to   && dt > to)   continue;

      var status = String(getCol(r,'status')||'');
      var schId  = String(r[2]||'');

      var dBadge = String(getCol(r,'driverBadge')||'').trim();
      if (dBadge && dBadge !== 'undefined' && (mode==='kmpl'||mode==='both')) {
        if (!driverMap[dBadge]) driverMap[dBadge] = {
          badge:dBadge, name:String(getCol(r,'driverName')||''), mobile:String(getCol(r,'driverMobile')||''),
          trips:0, yesTrips:0, noTrips:0, bdTrips:0,
          sanKm:0, optKm:0, cancelKm:0, deadKm:0,
          diesel:0, dieselC:0, kmpl:0, kmplC:0,
          dates:new Set(), vehicles:new Set(), schedules:new Set()
        };
        var d = driverMap[dBadge];
        d.trips++;
        if (status==='YES') { d.yesTrips++; d.optKm+=parseFloat(getCol(r,'optKm'))||0; }
        if (status==='NO')  d.noTrips++;
        if (status==='BD')  d.bdTrips++;
        d.sanKm    += parseFloat(getCol(r,'sanKm'))||0;
        d.cancelKm += parseFloat(getCol(r,'cancelKm'))||0;
        d.deadKm   += parseFloat(getCol(r,'deadKm'))||0;
        var dl=parseFloat(getCol(r,'diesel'))||0; if(dl>0){d.diesel+=dl;d.dieselC++;}
        var kp=parseFloat(getCol(r,'kmpl'))||0; if(kp>0){d.kmpl+=kp;d.kmplC++;}
        d.dates.add(dt);
        if(getCol(r,'vehicle')) d.vehicles.add(String(getCol(r,'vehicle')));
        if(r[3]) d.schedules.add(String(r[3]));
        if(getCol(r,'driverName')&&!d.name) d.name=String(getCol(r,'driverName'));
        if(getCol(r,'driverMobile')&&!d.mobile) d.mobile=String(getCol(r,'driverMobile'));
      }

      var cBadge = String(getCol(r,'condBadge')||'').trim();
      if (cBadge && cBadge !== 'undefined' && (mode==='epkm'||mode==='both')) {
        if (!condMap[cBadge]) condMap[cBadge] = {
          badge:cBadge, name:String(getCol(r,'condName')||''), mobile:String(getCol(r,'condMobile')||''),
          trips:0, yesTrips:0, noTrips:0, bdTrips:0,
          optKm:0, cancelKm:0,
          cash:0, allPass:0, students:0, dhaba:0,
          epkm:0, epkmC:0,
          dates:new Set(), schedules:new Set()
        };
        var c = condMap[cBadge];
        c.trips++;
        if(status==='YES'){c.yesTrips++;c.optKm+=parseFloat(getCol(r,'optKm'))||0;}
        if(status==='NO') c.noTrips++;
        if(status==='BD') c.bdTrips++;
        c.cancelKm += parseFloat(getCol(r,'cancelKm'))||0;
        // FIX: read cash/pass only once per date+schId from SERVICE rows, not TC rows
        var _ck = dt+'||'+schId;
        if (!_condCashCounted[_ck]) {
          _condCashCounted[_ck] = true;
          c.cash    += parseFloat(_svcGet(dt,schId,'cash'))     || 0;
          c.allPass += parseFloat(_svcGet(dt,schId,'allPass'))  || 0;
          c.students+= parseFloat(_svcGet(dt,schId,'students')) || 0;
          c.dhaba   += parseFloat(_svcGet(dt,schId,'dhaba'))    || 0;
          var ep=parseFloat(_svcGet(dt,schId,'epkm'))||0; if(ep>0){c.epkm+=ep;c.epkmC++;}
        }
        c.dates.add(dt);
        if(r[3]) c.schedules.add(String(r[3]));
        if(getCol(r,'condName')&&!c.name)    c.name=String(getCol(r,'condName'));
        if(getCol(r,'condMobile')&&!c.mobile) c.mobile=String(getCol(r,'condMobile'));
      }
    }

    var drivers = [];
    if (mode==='kmpl'||mode==='both') {
      Object.values(driverMap).forEach(function(d) {
        var avgKmpl = d.kmplC>0 ? parseFloat((d.kmpl/d.kmplC).toFixed(2)) : 0;
        if (avgKmpl < minKmpl || avgKmpl > maxKmpl) return;
        var compRate = d.trips>0 ? parseFloat((d.yesTrips/d.trips*100).toFixed(1)) : 0;
        drivers.push({
          badge:d.badge, name:d.name, mobile:d.mobile,
          days:d.dates.size, schedules:d.schedules.size,
          vehicles:[...d.vehicles].join(', '),
          trips:d.trips, yesTrips:d.yesTrips, noTrips:d.noTrips, bdTrips:d.bdTrips,
          compRate:compRate,
          sanKm:parseFloat(d.sanKm.toFixed(1)),
          optKm:parseFloat(d.optKm.toFixed(1)),
          cancelKm:parseFloat(d.cancelKm.toFixed(1)),
          deadKm:parseFloat(d.deadKm.toFixed(1)),
          diesel:parseFloat(d.diesel.toFixed(1)),
          avgKmpl:avgKmpl
        });
      });
      drivers.sort(function(a,b){return b.avgKmpl-a.avgKmpl;});
      drivers.forEach(function(d,i){d.rank=i+1;});
    }

    var conductors = [];
    if (mode==='epkm'||mode==='both') {
      Object.values(condMap).forEach(function(c) {
        var avgEpkm = c.epkmC>0 ? parseFloat((c.epkm/c.epkmC).toFixed(2)) :
                      c.optKm>0 ? parseFloat((c.cash/c.optKm).toFixed(2)) : 0;
        if (avgEpkm < minEpkm || avgEpkm > maxEpkm) return;
        var compRate = c.trips>0 ? parseFloat((c.yesTrips/c.trips*100).toFixed(1)) : 0;
        conductors.push({
          badge:c.badge, name:c.name, mobile:c.mobile,
          days:c.dates.size, schedules:c.schedules.size,
          trips:c.trips, yesTrips:c.yesTrips, noTrips:c.noTrips,
          compRate:compRate,
          optKm:parseFloat(c.optKm.toFixed(1)),
          cancelKm:parseFloat(c.cancelKm.toFixed(1)),
          cash:Math.round(c.cash),
          allPass:Math.round(c.allPass),
          students:Math.round(c.students),
          dhaba:Math.round(c.dhaba),
          avgEpkm:avgEpkm
        });
      });
      conductors.sort(function(a,b){return b.avgEpkm-a.avgEpkm;});
      conductors.forEach(function(c,i){c.rank=i+1;});
    }

    var dStats = drivers.length ? {
      count:drivers.length,
      avgKmpl:parseFloat((drivers.reduce(function(s,d){return s+d.avgKmpl;},0)/drivers.length).toFixed(2)),
      minKmpl:drivers[drivers.length-1].avgKmpl,
      maxKmpl:drivers[0].avgKmpl,
      totalDiesel:parseFloat(drivers.reduce(function(s,d){return s+d.diesel;},0).toFixed(1)),
      totalOptKm:parseFloat(drivers.reduce(function(s,d){return s+d.optKm;},0).toFixed(1))
    } : null;
    var cStats = conductors.length ? {
      count:conductors.length,
      avgEpkm:parseFloat((conductors.reduce(function(s,c){return s+c.avgEpkm;},0)/conductors.length).toFixed(2)),
      minEpkm:conductors[conductors.length-1].avgEpkm,
      maxEpkm:conductors[0].avgEpkm,
      totalCash:Math.round(conductors.reduce(function(s,c){return s+c.cash;},0)),
      totalPass:Math.round(conductors.reduce(function(s,c){return s+c.allPass;},0))
    } : null;

    return {
      found:true, from:from, to:to,
      minKmpl:minKmpl, maxKmpl:maxKmpl,
      minEpkm:minEpkm, maxEpkm:maxEpkm,
      mode:mode,
      drivers:drivers, dStats:dStats,
      conductors:conductors, cStats:cStats
    };
  } catch(e) {
    return { error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// UTILISATION REPORT
// ═══════════════════════════════════════════════════════════════
function buildUtilisationReport(tcRows) {
  var byDate = {};

  tcRows.forEach(function(r) {
    var dt     = formatDateVal(r[0]);
    var schNo  = String(r[3] || r[2] || '');
    var status = String(getCol(r,'status') || '');
    var optKm  = parseFloat(getCol(r,'optKm')) || 0;
    var canKm  = parseFloat(getCol(r,'cancelKm')) || 0;
    var driver = String(getCol(r,'driverBadge') || '').trim();
    var cond   = String(getCol(r,'condBadge') || '').trim();
    var veh    = String(getCol(r,'vehicle') || '').trim();

    if (!byDate[dt]) byDate[dt] = {
      date      : dt,
      optSchNos : new Set(),
      optTrips  : 0,
      optKm     : 0,
      canKm     : 0,
      drivers   : new Set(),
      conductors: new Set(),
      vehicles  : new Set()
    };

    var d = byDate[dt];

    if (status === 'YES') {
      d.optSchNos.add(schNo);
      d.optTrips++;
      d.optKm  += optKm;
    }
    d.canKm += canKm;
    if (driver) d.drivers.add(driver);
    if (cond)   d.conductors.add(cond);
    if (veh)    d.vehicles.add(veh);
  });

  return Object.keys(byDate).sort().map(function(dt) {
    var d       = byDate[dt];
    var optSch  = d.optSchNos.size;
    var optTrip = d.optTrips;
    var optKm   = parseFloat(d.optKm.toFixed(2));
    var canKm   = parseFloat(d.canKm.toFixed(2));
    var nDriver = d.drivers.size;
    var nCond   = d.conductors.size;
    var nVeh    = d.vehicles.size;
    var vehUtil  = nVeh > 0 ? parseFloat((optKm / nVeh).toFixed(2)) : 0;
    var crewCount = (nDriver + nCond) / 2;
    var crewUtil  = crewCount > 0 ? parseFloat((optKm / crewCount).toFixed(2)) : 0;

    return {
      date       : dt,
      optSch     : optSch,
      optTrips   : optTrip,
      optKm      : optKm,
      canKm      : canKm,
      drivers    : nDriver,
      conductors : nCond,
      vehicles   : nVeh,
      vehUtil    : vehUtil,
      crewUtil   : crewUtil
    };
  });
}

// ── SCH-WISE DAY REPORT ──────────────────────────────────────
function buildSchWiseDayReport(tcRows, svcRows, dieselMap) {
  var _sm=makeSvcMap(svcRows);
  var byDate = {};
  tcRows.forEach(function(r) {
    var dt = formatDateVal(r[0]);
    if (!byDate[dt]) byDate[dt] = {};
    var sid = String(r[2] || '');
    if (!byDate[dt][sid]) byDate[dt][sid] = {
      schId:sid, schNo:r[3]||'', route:r[4]||'',
      trips:0, yesTrips:0, optKm:0, canKm:0, cash:0,
      diesel:0, dieselC:0, kmpl:0, kmplC:0,
      epkm:0, epkmC:0,
      driverBadge:getCol(r,'driverBadge')||'', driverName:getCol(r,'driverName')||'',
      condBadge:getCol(r,'condBadge')||'', condName:getCol(r,'condName')||''
    };
    var s = byDate[dt][sid];
    s.trips++;
    if (getCol(r,'status') === 'YES') { s.yesTrips++; s.optKm += parseFloat(getCol(r,'optKm'))||0; }
    s.canKm += parseFloat(getCol(r,'cancelKm'))||0;
    var _dt2=formatDateVal(r[0]), _sid2=String(r[2]||'');
    s.cash  += parseFloat(svcVal(_sm,_dt2,_sid2,'cash'))||0;
    var _dKey2=formatDateVal(r[0])+'|'+String(r[2]||''); var _dm2=(dieselMap||{})[_dKey2]||{};
    var dl = _dm2.diesel||0; if(dl>0){s.diesel=dl;}
    var kp = _dm2.kmpl||0;   if(kp>0){s.kmpl=kp; s.kmplC=1;}
    var ep = parseFloat(getCol(r,'epkm'))||0; if(ep>0){s.epkm+=ep;s.epkmC++;}
    if(getCol(r,'driverBadge')) s.driverBadge=getCol(r,'driverBadge');
    if(getCol(r,'driverName'))  s.driverName=getCol(r,'driverName');
    if(getCol(r,'condBadge'))   s.condBadge=getCol(r,'condBadge');
    if(getCol(r,'condName'))    s.condName=getCol(r,'condName');
  });

  var result = {};
  Object.keys(byDate).sort().forEach(function(dt) {
    var rows2 = Object.values(byDate[dt]).map(function(s) {
      var avgKmpl = s.kmplC>0 ? parseFloat((s.kmpl/s.kmplC).toFixed(2)) : 0;
      var avgEpkm = s.epkmC>0 ? parseFloat((s.epkm/s.epkmC).toFixed(2)) :
                    (s.optKm>0 ? parseFloat((s.cash/s.optKm).toFixed(2)) : 0);
      return {
        schId:s.schId, schNo:s.schNo, route:s.route,
        trips:s.trips, yesTrips:s.yesTrips,
        optKm:parseFloat(s.optKm.toFixed(2)), canKm:parseFloat(s.canKm.toFixed(2)),
        cash:Math.round(s.cash), diesel:parseFloat(s.diesel.toFixed(1)),
        driverBadge:s.driverBadge, driverName:s.driverName,
        condBadge:s.condBadge, condName:s.condName,
        epkm:avgEpkm, kmpl:avgKmpl
      };
    });
    rows2.sort(function(a,b){return b.epkm-a.epkm;});
    rows2.forEach(function(r,i){r.rank=i+1;});
    result[dt] = rows2;
  });
  return result;
}

function buildMergedDOR(tcRows, svcRows) {
  var svcMap = {};
  (svcRows||[]).forEach(function(r){
    var k=formatDateVal(r[0])+'|'+String(getCol(r,'schId')||r[2]||'');
    svcMap[k] = r;
  });

  var m = {};
  tcRows.forEach(function(r) {
    var type  = getCol(r,'tripType')||'UNKNOWN';
    var schId = String(r[2]||'');
    var status= getCol(r,'status')||'';
    var sanKm = parseFloat(getCol(r,'sanKm'))||0;
    var optKm = parseFloat(getCol(r,'optKm'))||0;
    var canKm = parseFloat(getCol(r,'cancelKm'))||0;
    if(!m[type]) m[type]={type:type,sanKm:0,sanSch:new Set(),sanTrip:0,optSch:new Set(),optTrip:0,optKm:0,canKm:0,cash:0,noPass:0,allPass:0,dhaba:0};
    m[type].sanKm+=sanKm; m[type].sanTrip++; m[type].sanSch.add(schId);
    if(status==='YES'){m[type].optTrip++;m[type].optKm+=optKm;m[type].optSch.add(schId);}
    else{m[type].canKm+=canKm;}
  });
  Object.keys(svcMap).forEach(function(key){
    var sr   = svcMap[key];
    var cash = parseFloat(getCol(sr,'cash'))    || 0;
    var allP = parseFloat(getCol(sr,'allPass')) || 0;
    var stud = parseFloat(getCol(sr,'students'))|| 0;
    var dhab = parseFloat(getCol(sr,'dhaba'))   || 0;
    var schId = String(getCol(sr,'schId')||sr[2]||'');
    var matchRow = tcRows.find(function(r){ return String(r[2]||'')=== schId; });
    var type = matchRow ? (getCol(matchRow,'tripType')||'UNKNOWN') : 'UNKNOWN';
    if(!m[type]) return;
    m[type].cash+=cash; m[type].noPass+=Math.max(0,allP-stud); m[type].allPass+=allP; m[type].dhaba+=dhab;
  });

  var rows = Object.values(m).sort(function(a,b){return a.type.localeCompare(b.type);}).map(function(x){
    var epkm = x.optKm>0 ? parseFloat((x.cash/x.optKm).toFixed(2)) : 0;
    return{type:x.type,sanKm:parseFloat(x.sanKm.toFixed(2)),sanSch:x.sanSch.size,sanTrip:x.sanTrip,
           optSch:x.optSch.size,optTrip:x.optTrip,optKm:parseFloat(x.optKm.toFixed(2)),
           canKm:parseFloat(x.canKm.toFixed(2)),cash:x.cash,noPass:x.noPass,allPass:x.allPass,dhaba:x.dhaba,epkm:epkm};
  });
  var gt={type:'GRAND TOTAL',sanKm:0,sanSch:0,sanTrip:0,optSch:0,optTrip:0,optKm:0,canKm:0,cash:0,noPass:0,allPass:0,dhaba:0,epkm:0};
  rows.forEach(function(r){gt.sanKm+=r.sanKm;gt.sanSch+=r.sanSch;gt.sanTrip+=r.sanTrip;gt.optSch+=r.optSch;gt.optTrip+=r.optTrip;gt.optKm+=r.optKm;gt.canKm+=r.canKm;gt.cash+=r.cash;gt.noPass+=r.noPass;gt.allPass+=r.allPass;gt.dhaba+=r.dhaba;});
  gt.sanKm=parseFloat(gt.sanKm.toFixed(2)); gt.optKm=parseFloat(gt.optKm.toFixed(2)); gt.canKm=parseFloat(gt.canKm.toFixed(2));
  gt.epkm = gt.optKm>0 ? parseFloat((gt.cash/gt.optKm).toFixed(2)):0;
  return {rows:rows, grandTotal:gt};
}

function buildEmployeePerfDate(tcRows, svcRows, dieselMap) {
  var _sm=makeSvcMap(svcRows); var _cc={};
  var byDate={};
  tcRows.forEach(function(r){
    var dt=formatDateVal(r[0]);
    if(!byDate[dt])byDate[dt]={drivers:{},conductors:{}};
    var dbadge=String(getCol(r,'driverBadge')||'');
    if(dbadge && dbadge !== 'undefined'){
      var d=byDate[dt].drivers;
      if(!d[dbadge])d[dbadge]={badge:dbadge,name:getCol(r,'driverName')||'--',mobile:getCol(r,'driverMobile')||'--',trips:0,optKm:0,diesel:0,deadKm:0,kmpl:0,kmplC:0};
      var _dKey3=formatDateVal(r[0])+'|'+String(r[2]||''); var _dm3=(dieselMap||{})[_dKey3]||{};
      d[dbadge].trips++;d[dbadge].optKm+=parseFloat(getCol(r,'optKm'))||0;d[dbadge].diesel+=_dm3.diesel||0;d[dbadge].deadKm+=_dm3.deadKm||0;
      var k=_dm3.kmpl||0;if(k>0){d[dbadge].kmpl+=k;d[dbadge].kmplC++;}
    }
    var cbadge=String(getCol(r,'condBadge')||'');
    if(cbadge && cbadge !== 'undefined'){
      var cd=byDate[dt].conductors;
      if(!cd[cbadge])cd[cbadge]={badge:cbadge,name:getCol(r,'condName')||'--',mobile:getCol(r,'condMobile')||'--',trips:0,cash:0,pass:0,students:0,dhaba:0,km:0,epkm:0,epkmC:0};
      cd[cbadge].trips++; var _dt3=formatDateVal(r[0]),_sid3=String(r[2]||''),_ck3=cbadge+'|||'+_dt3+'|||'+_sid3;
      // FIX [LOW-2]: Use robust separator '|||' to avoid key collisions
      if(!_cc[_ck3]){cd[cbadge].cash+=parseFloat(svcVal(_sm,_dt3,_sid3,'cash'))||0;cd[cbadge].pass+=parseFloat(svcVal(_sm,_dt3,_sid3,'allPass'))||0;
      cd[cbadge].students+=parseFloat(svcVal(_sm,_dt3,_sid3,'students'))||0;cd[cbadge].dhaba+=parseFloat(svcVal(_sm,_dt3,_sid3,'dhaba'))||0; _cc[_ck3]=1;}
      cd[cbadge].km+=parseFloat(getCol(r,'optKm'))||0;
      var e=parseFloat(svcVal(_sm,_dt3,_sid3,'epkm'))||0;if(e>0){cd[cbadge].epkm+=e;cd[cbadge].epkmC++;}
    }
  });
  var result=[];
  Object.keys(byDate).sort().forEach(function(dt){
    var drvs=Object.values(byDate[dt].drivers).map(function(x){
      return{badge:x.badge,name:x.name,mobile:x.mobile,trips:x.trips,
             optKm:parseFloat(x.optKm.toFixed(1)),diesel:parseFloat(x.diesel.toFixed(1)),
             deadKm:parseFloat(x.deadKm.toFixed(1)),avgKmpl:x.kmplC>0?parseFloat((x.kmpl/x.kmplC).toFixed(2)):0};
    }).sort(function(a,b){return b.avgKmpl-a.avgKmpl;});
    var cnds=Object.values(byDate[dt].conductors).map(function(x){
      return{badge:x.badge,name:x.name,mobile:x.mobile,trips:x.trips,
             cash:x.cash,pass:x.pass,students:x.students,dhaba:x.dhaba,
             km:parseFloat(x.km.toFixed(1)),avgEpkm:x.epkmC>0?parseFloat((x.epkm/x.epkmC).toFixed(2)):0};
    }).sort(function(a,b){return b.avgEpkm-a.avgEpkm;});
    result.push({date:dt,drivers:drvs,conductors:cnds});
  });
  return result;
}

function buildCancellationRows(rows) {
  return rows.filter(function(r){ return getCol(r,'status')==="NO"||getCol(r,'status')==="BD"||getCol(r,'status')==="ACCIDENT"||getCol(r,'status')==="OTHER"; })
    .map(function(r){
      return {date:formatDateVal(r[0]),schId:r[2],route:r[4],type:getCol(r,'tripType')||"",
              sanKm:parseFloat(getCol(r,'sanKm'))||0,cancelKm:parseFloat(getCol(r,'cancelKm'))||0,
              status:getCol(r,'status')||"",driver:getCol(r,'driverName')||"",vehicle:getCol(r,'vehicle')||""};
    }).sort(function(a,b){return a.date.localeCompare(b.date);});
}
function buildServiceTypeRows(rows, typeName, svcRows) {
  var _sm=makeSvcMap(svcRows);
  return rows.filter(function(r){ return String(getCol(r,'tripType')||"").toUpperCase()===typeName.toUpperCase(); })
    .map(function(r){
      var _dt=formatDateVal(r[0]); var _sid=String(r[2]||'');
      return {date:_dt,schId:r[2],route:r[4],
              sanKm:parseFloat(getCol(r,'sanKm'))||0,optKm:parseFloat(getCol(r,'optKm'))||0,
              cancelKm:parseFloat(getCol(r,'cancelKm'))||0,status:getCol(r,'status')||"",
              vehicle:getCol(r,'vehicle')||"",driver:getCol(r,'driverName')||"",
              cash:parseFloat(svcVal(_sm,_dt,_sid,'cash'))||0,
              pass:parseFloat(svcVal(_sm,_dt,_sid,'allPass'))||0,
              epkm:svcVal(_sm,_dt,_sid,'epkm')||0};
    }).sort(function(a,b){return a.date.localeCompare(b.date);});
}
function buildDhabaRows(rows) {
  var m={};
  rows.forEach(function(r){
    var badge=String(getCol(r,'condBadge')||""); if(!badge)return;
    var dhaba=parseFloat(getCol(r,'dhaba'))||0; if(dhaba<=0)return;
    var dt=formatDateVal(r[0]);
    var key=dt+"|"+badge;
    if(!m[key])m[key]={date:dt,badge:badge,name:getCol(r,'condName')||"",schId:r[2],dhaba:0,cash:0};
    m[key].dhaba+=dhaba; m[key].cash+=parseFloat(getCol(r,'cash'))||0;
  });
  return Object.values(m).sort(function(a,b){return a.date.localeCompare(b.date)||b.dhaba-a.dhaba;});
}

// ── Only Cash item-wise breakdown ─────────────────────────
// Returns rows: [{date, itemId, itemName, cash, passengers}]
// grouped by date + item, sorted by date then itemName
function buildOnlyCashItemReport(ocRows) {
  var m = {};
  ocRows.forEach(function(r) {
    var dt       = formatDateVal(r[0]);
    var itemId   = String(r[2] || '').trim();       // col 2 = Schedule ID = itemId
    var itemName = String(r[4] || itemId).trim();   // col 4 = Route = itemName
    var cash     = parseFloat(getCol(r, 'cash')) || 0;
    var pass     = parseFloat(getCol(r, 'allPass')) || 0;
    var key      = dt + '|' + itemId;
    if (!m[key]) m[key] = { date: dt, itemId: itemId, itemName: itemName, cash: 0, passengers: 0 };
    m[key].cash       += cash;
    m[key].passengers += pass;
  });
  return Object.values(m).sort(function(a, b) {
    return a.date.localeCompare(b.date) || a.itemName.localeCompare(b.itemName);
  });
}
function buildDailyServiceSummary(rows) {
  var m={};
  rows.forEach(function(r){
    var dt=formatDateVal(r[0]);
    var type=String(getCol(r,'tripType')||"UNKNOWN");
    var key=dt+"|"+type;
    if(!m[key])m[key]={date:dt,type:type,trips:0,yes:0,no:0,bd:0,optKm:0,sanKm:0,cancelKm:0};
    m[key].trips++; m[key].sanKm+=parseFloat(getCol(r,'sanKm'))||0; m[key].optKm+=parseFloat(getCol(r,'optKm'))||0;
    m[key].cancelKm+=parseFloat(getCol(r,'cancelKm'))||0;
    if(getCol(r,'status')==="YES")m[key].yes++;
    else if(getCol(r,'status')==="BD")m[key].bd++;
    else m[key].no++;
  });
  return Object.values(m).sort(function(a,b){return a.date.localeCompare(b.date)||a.type.localeCompare(b.type);})
    .map(function(x){return {date:x.date,type:x.type,trips:x.trips,yes:x.yes,no:x.no,bd:x.bd,
      sanKm:x.sanKm.toFixed(2),optKm:x.optKm.toFixed(2),cancelKm:x.cancelKm.toFixed(2),
      compRate:x.trips>0?(x.yes/x.trips*100).toFixed(1):"0.0"};});
}
function buildDateWiseDetail(rows, svcRows) {
  var _sm=makeSvcMap(svcRows);
  return rows.map(function(r){
    var _dt=formatDateVal(r[0]); var _sid=String(r[2]||'');
    return {date:_dt,schId:r[2],schNo:r[3],route:r[4],type:getCol(r,'tripType')||"",
            sanKm:parseFloat(getCol(r,'sanKm'))||0,optKm:parseFloat(getCol(r,'optKm'))||0,
            cancelKm:parseFloat(getCol(r,'cancelKm'))||0,status:getCol(r,'status')||"",
            vehicle:getCol(r,'vehicle')||"",driver:getCol(r,'driverName')||"",
            conductor:getCol(r,'condName')||"",
            cash:parseFloat(svcVal(_sm,_dt,_sid,'cash'))||0,
            pass:parseFloat(svcVal(_sm,_dt,_sid,'allPass'))||0,
            epkm:svcVal(_sm,_dt,_sid,'epkm')||"0"};
  }).sort(function(a,b){return a.date.localeCompare(b.date)||a.schId.localeCompare(b.schId);});
}

// FIX [LOW-2]: Use robust '|||' separator in cache keys to avoid collisions
function buildTopEpkm(rows, svcRows, n, dir) {
  var _sm=makeSvcMap(svcRows); var _cc={};
  var m={};
  rows.forEach(function(r){
    var badge=String(getCol(r,'condBadge')||""); if(!badge)return;
    var km=parseFloat(getCol(r,'optKm'))||0; if(km<=0)return;
    var _dt=formatDateVal(r[0]),_sid=String(r[2]||'');
    var _ck = _dt + '|||' + _sid;           // FIX: robust separator
    var _ckp= _dt + '|||' + _sid + '|||p';  // FIX: robust pass key
    var cash=_cc[_ck]?0:(parseFloat(svcVal(_sm,_dt,_sid,'cash'))||0);
    if(!_cc[_ck]) _cc[_ck]=1;
    if(!m[badge])m[badge]={badge:badge,name:getCol(r,'condName')||"",mobile:getCol(r,'condMobile')||"",cash:0,km:0,pass:0,days:new Set()};
    var _pass=_cc[_ckp]?0:(parseFloat(svcVal(_sm,_dt,_sid,'allPass'))||0);
    if(!_cc[_ckp])_cc[_ckp]=1;
    m[badge].cash+=cash; m[badge].km+=km; m[badge].pass+=_pass;
    m[badge].days.add(formatDateVal(r[0]));
  });
  return Object.values(m).map(function(c){
    var epkm=c.km>0?parseFloat((c.cash/c.km).toFixed(2)):0;
    return {badge:c.badge,name:c.name,mobile:c.mobile,days:c.days.size,
            totalCash:c.cash,totalKm:c.km.toFixed(1),pass:c.pass,epkm:epkm};
  }).sort(function(a,b){return dir==="top"?b.epkm-a.epkm:a.epkm-b.epkm;}).slice(0,n);
}

function buildWeeklyDriver(rows, dieselMap) {
  var m={};var _sd={};
  rows.forEach(function(r){
    var badge=String(getCol(r,'driverBadge')||""); if(!badge||badge==='undefined')return;
    var dt=new Date(formatDateVal(r[0]));
    var day=dt.getDay(); var diff=dt.getDate()-(day===0?6:day-1);
    var mon=new Date(new Date(dt).setDate(diff)); var wk=fmtDate(mon);
    var key=wk+"|"+badge;
    if(!m[key])m[key]={week:wk,badge:badge,name:getCol(r,'driverName')||"",optKm:0,diesel:0,deadKm:0,kmpl:0,kmplC:0,trips:0,yes:0};
    var _dKey=formatDateVal(r[0])+'|'+String(r[2]||''); var _dm=(dieselMap||{})[_dKey]||{};
    m[key].trips++; m[key].optKm+=parseFloat(getCol(r,'optKm'))||0;
    if(getCol(r,'status')==="YES")m[key].yes++;
    if(!_sd[key+'|'+_dKey]){_sd[key+'|'+_dKey]=1;
      m[key].diesel+=_dm.diesel||0; m[key].deadKm+=_dm.deadKm||0;
      var k=_dm.kmpl||0; if(k>0){m[key].kmpl+=k;m[key].kmplC++;}
    }
  });
  return Object.values(m).map(function(x){
    return {week:x.week,badge:x.badge,name:x.name,trips:x.trips,yesTrips:x.yes,
            optKm:parseFloat(x.optKm.toFixed(1)),diesel:parseFloat(x.diesel.toFixed(1)),deadKm:parseFloat(x.deadKm.toFixed(1)),
            avgKmpl:x.kmplC>0?(x.kmpl/x.kmplC).toFixed(2):"--",
            compRate:x.trips>0?(x.yes/x.trips*100).toFixed(1):"0.0"};
  }).sort(function(a,b){return (a.week+a.badge).localeCompare(b.week+b.badge);});
}
function buildWeeklyConductor(rows, svcRows) {
  var _sm=makeSvcMap(svcRows); var _cc={};
  var m={};
  rows.forEach(function(r){
    var badge=String(getCol(r,'condBadge')||""); if(!badge||badge==='undefined')return;
    var dt=new Date(formatDateVal(r[0]));
    var day=dt.getDay(); var diff=dt.getDate()-(day===0?6:day-1);
    var mon=new Date(new Date(dt).setDate(diff)); var wk=fmtDate(mon);
    var key=wk+"|"+badge;
    if(!m[key])m[key]={week:wk,badge:badge,name:getCol(r,'condName')||"",cash:0,pass:0,students:0,dhaba:0,epkm:0,epkmC:0,days:new Set()};
    m[key].days.add(formatDateVal(r[0]));
    var _dt=formatDateVal(r[0]),_sid=String(r[2]||''),_ck=_dt+'||'+_sid;
    if(!_cc[_ck]){m[key].cash+=parseFloat(svcVal(_sm,_dt,_sid,'cash'))||0; m[key].pass+=parseFloat(svcVal(_sm,_dt,_sid,'allPass'))||0;
    m[key].students+=parseFloat(svcVal(_sm,_dt,_sid,'students'))||0; m[key].dhaba+=parseFloat(svcVal(_sm,_dt,_sid,'dhaba'))||0; _cc[_ck]=1;}
    var e=parseFloat(svcVal(_sm,_dt,_sid,'epkm'))||0; if(e>0){m[key].epkm+=e;m[key].epkmC++;}
  });
  return Object.values(m).map(function(x){
    return {week:x.week,badge:x.badge,name:x.name,days:x.days.size,
            cash:x.cash,pass:x.pass,students:x.students,dhaba:x.dhaba,
            avgEpkm:x.epkmC>0?(x.epkm/x.epkmC).toFixed(2):"--"};
  }).sort(function(a,b){return (a.week+a.badge).localeCompare(b.week+b.badge);});
}
function buildWeeklyVehicle(rows, dieselMap) {
  var m={};var _sd={};
  rows.forEach(function(r){
    var veh=String(getCol(r,'vehicle')||""); if(!veh||veh==="DEPOT")return;
    var dt=new Date(formatDateVal(r[0]));
    var day=dt.getDay(); var diff=dt.getDate()-(day===0?6:day-1);
    var mon=new Date(new Date(dt).setDate(diff)); var wk=fmtDate(mon);
    var key=wk+"|"+veh;
    if(!m[key])m[key]={week:wk,vehicle:veh,type:getCol(r,'vehType')||"",optKm:0,diesel:0,deadKm:0,kmpl:0,kmplC:0,days:new Set()};
    var _dKey=formatDateVal(r[0])+'|'+String(r[2]||''); var _dm=(dieselMap||{})[_dKey]||{};
    m[key].days.add(formatDateVal(r[0]));
    m[key].optKm+=parseFloat(getCol(r,'optKm'))||0;
    if(!_sd[key+'|'+_dKey]){_sd[key+'|'+_dKey]=1;
      m[key].diesel+=_dm.diesel||0; m[key].deadKm+=_dm.deadKm||0;
      var k=_dm.kmpl||0; if(k>0){m[key].kmpl+=k;m[key].kmplC++;}
    }
  });
  return Object.values(m).map(function(x){
    return {week:x.week,vehicle:x.vehicle,type:x.type,days:x.days.size,
            optKm:parseFloat(x.optKm.toFixed(1)),diesel:parseFloat(x.diesel.toFixed(1)),deadKm:parseFloat(x.deadKm.toFixed(1)),
            avgKmpl:x.kmplC>0?(x.kmpl/x.kmplC).toFixed(2):"--"};
  }).sort(function(a,b){return (a.week+a.vehicle).localeCompare(b.week+b.vehicle);});
}

function filterRows(rows, from, to, module) {
  return rows.slice(1).filter(function(r) {
    if (!r[0]) return false;
    var dt = formatDateVal(r[0]);
    if (module && String(getCol(r,'mod')) !== module) return false;
    if (from && dt < from) return false;
    if (to   && dt > to)   return false;
    return true;
  });
}
function fmtDate(d) {
  return d.getFullYear() + '-' +
    String(d.getMonth()+1).padStart(2,'0') + '-' +
    String(d.getDate()).padStart(2,'0');
}
function emptyDash() {
  return {totalTrips:0,yesTrips:0,noTrips:0,bdTrips:0,accidentTrips:0,otherTrips:0,
    totalSanKm:0,totalOptKm:0,totalCancelKm:0,totalDiesel:0,totalUria:0,avgKmpl:0,
    totalCash:0,avgEpkm:0,totalPass:0,totalStudents:0,totalDhaba:0,
    trend:[],schedules:[],topDrivers:[],lowestDrivers:[],topConductors:[],lowestConductors:[]};
}

// ═══════════════════════════════════════════════════════════════
// ENTRY STATUS FUNCTIONS
// ═══════════════════════════════════════════════════════════════
function getVehicleSchIds(date, vehicleNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allotSh = ss.getSheetByName(ALLOT_SH);
    var combSh  = ss.getSheetByName(SHEET_NAME);
    if (!combSh) return {error:'No data sheet', schedules:[]};

    detectColumns(combSh);
    var combRows = combSh.getDataRange().getValues();

    var allotSchIds = [];
    if (allotSh) {
      var allotRows = allotSh.getDataRange().getValues();
      for (var i = 1; i < allotRows.length; i++) {
        if (!allotRows[i][0]) continue;
        var rd = formatDateVal(allotRows[i][0]);
        if (rd === date && String(allotRows[i][3]) === String(vehicleNo)) {
          allotSchIds.push({
            schId: String(allotRows[i][1]),
            schNo: String(allotRows[i][2]||'')
          });
        }
      }
    }

    var tcVehSchIds = {};
    for (var j = 1; j < combRows.length; j++) {
      var r = combRows[j];
      if (!r[0]) continue;
      if (formatDateVal(r[0]) !== date) continue;
      if (String(getCol(r,'mod')) !== 'TC') continue;
      if (String(getCol(r,'vehicle')) === String(vehicleNo)) {
        tcVehSchIds[String(r[2])] = true;
      }
    }

    var schIdSet = {};
    allotSchIds.forEach(function(a){ schIdSet[a.schId] = a.schNo; });
    Object.keys(tcVehSchIds).forEach(function(sid){
      if (!schIdSet[sid]) schIdSet[sid] = sid;
    });

    if (Object.keys(schIdSet).length === 0) return {schedules:[]};

    var result = [];
    Object.keys(schIdSet).forEach(function(schId) {
      var tcTrips = [], dieselDone = false;
      var optKm = 0, trips = 0, yesTrips = 0;
      var driverBadge = '', driverName = '', driverMobile = '';

      for (var k = 1; k < combRows.length; k++) {
        var cr = combRows[k];
        if (!cr[0]) continue;
        if (formatDateVal(cr[0]) !== date) continue;
        if (String(cr[2]) !== schId) continue;
        var mod = String(getCol(cr,'mod'));
        if (mod === 'TC') {
          trips++;
          var status = String(getCol(cr,'status'));
          if (status === 'YES') { yesTrips++; optKm += parseFloat(getCol(cr,'optKm'))||0; }
          if (!driverBadge && getCol(cr,'driverBadge')) {
            driverBadge  = String(getCol(cr,'driverBadge'));
            driverName   = String(getCol(cr,'driverName')||'');
            driverMobile = String(getCol(cr,'driverMobile')||'');
          }
          tcTrips.push({
            schNo:    getCol(cr,'schNo')||cr[3]||'',
            route:    getCol(cr,'route')||cr[4]||'',
            status:   status,
            optKm:    parseFloat(getCol(cr,'optKm'))||0,
            cancelKm: parseFloat(getCol(cr,'cancelKm'))||0
          });
        }
        if (mod === 'DIESEL') dieselDone = true;
      }

      result.push({
        schId       : schId,
        schNo       : schIdSet[schId] || schId,
        tcDone      : trips > 0,
        trips       : trips,
        yesTrips    : yesTrips,
        optKm       : parseFloat(optKm.toFixed(2)),
        dieselDone  : dieselDone,
        driverBadge : driverBadge,
        driverName  : driverName,
        driverMobile: driverMobile,
        tcTrips     : tcTrips
      });
    });

    result.sort(function(a,b){
      if (a.tcDone !== b.tcDone) return a.tcDone ? -1 : 1;
      return String(a.schId).localeCompare(String(b.schId));
    });

    return {schedules: result};
  } catch(e) {
    return {error: e.message, schedules: []};
  }
}

// FIX [MED-4]: Re-read rows after deletions to avoid stale snapshot for schOptKm
function saveMultiSchFuel(date, vehicleNo, vehicleType,
                           totalDiesel, totalUria, deadKm, divKm,
                           selectedSchIds, schDrivers) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss.getSheetByName(SHEET_NAME);
    if (!sh) return {ok:false, error:'No data sheet'};
    detectColumns(sh);
    var now  = new Date();

    var drvMap = {};
    (schDrivers||[]).forEach(function(d){ drvMap[d.schId] = d; });

    // Delete duplicate rows first for all selected schedules
    selectedSchIds.forEach(function(sid) {
      deleteDuplicateRows(sh, date, 'DIESEL', sid);
    });

    // FIX [MED-4]: Re-read rows AFTER all deletions to get fresh data
    var rows = sh.getDataRange().getValues();

    var schOptKm = {};
    selectedSchIds.forEach(function(sid){ schOptKm[sid] = 0; });
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      if (formatDateVal(r[0]) !== date) continue;
      if (String(getCol(r,'mod')) !== 'TC') continue;
      var sid = String(r[2]||'');
      if (schOptKm.hasOwnProperty(sid) && String(getCol(r,'status'))==='YES') {
        schOptKm[sid] = (schOptKm[sid]||0) + (parseFloat(getCol(r,'optKm'))||0);
      }
    }

    var totalOptKm   = selectedSchIds.reduce(function(s,sid){ return s+(schOptKm[sid]||0); }, 0);
    var totalKmKmpl  = totalOptKm + (parseFloat(deadKm)||0) + (parseFloat(divKm)||0);
    var kmpl1st      = (totalDiesel>0 && totalKmKmpl>0)
                        ? parseFloat((totalKmKmpl/totalDiesel).toFixed(2)) : 0;

    var saved = [];

    selectedSchIds.forEach(function(sid, idx) {
      var isFirst = (idx === 0);
      var drv = drvMap[sid] || {};
      var drvBadge  = drv.driverBadge  || '';
      var drvName   = drv.driverName   || '';
      var drvMobile = drv.driverMobile || '';

      var tripType = '', schNo = '';
      for (var i = 1; i < rows.length; i++) {
        if (formatDateVal(rows[i][0])!==date||String(rows[i][2])!==sid) continue;
        if (!schNo) schNo = getCol(rows[i],'schNo')||rows[i][3]||'';
        if (String(rows[i][1])==='TC'&&!tripType) tripType=getCol(rows[i],'tripType')||'';
        if (schNo&&tripType) break;
      }

      // Look up actual sanKm for this schedule from TC row
      var schSanKm = 0;
      for (var j = 1; j < rows.length; j++) {
        if (formatDateVal(rows[j][0])===date && String(rows[j][2])===sid && String(rows[j][1])==='TC') {
          schSanKm = parseFloat(getCol(rows[j],'sanKm')) || 0;
          break;
        }
      }

      var schOpt = schOptKm[sid] || 0;
      var row = new Array(HEADERS.length).fill('');
      row[0]  = date;
      row[1]  = 'DIESEL';
      row[2]  = sid;
      row[3]  = schNo;
      row[4]  = 'Fuel Entry';
      row[5]  = tripType;
      row[6]  = isFirst ? schSanKm : schSanKm;  // Use actual sanKm per schedule
      row[7]  = isFirst ? totalOptKm : schOpt;
      row[8]  = 0;
      row[9]  = 'DIESEL';
      row[10] = vehicleNo;
      row[11] = vehicleType;
      row[12] = drvBadge;
      row[13] = drvName;
      row[14] = drvMobile;
      row[15] = isFirst ? (parseFloat(totalDiesel)||0) : 0;
      row[16] = isFirst ? (parseFloat(totalUria)||0)   : 0;
      row[17] = isFirst ? kmpl1st : 0;  // Store KMPL only on primary row
      row[18] = isFirst ? (parseFloat(deadKm)||0) : 0;
      row[19] = isFirst ? (parseFloat(divKm)||0)  : 0;
      row[29] = now;
      sh.getRange(sh.getLastRow()+1, 1, 1, HEADERS.length).setValues([row]);

      updateTcRowsWithDriver(sh, date, sid, vehicleNo, vehicleType, drvBadge, drvName, drvMobile);
      saved.push(sid);
    });

    return {
      ok: true, saved: saved,
      primarySch: selectedSchIds[0],
      totalOptKm: parseFloat(totalOptKm.toFixed(1)),
      kmpl: kmpl1st,
      message: 'Fuel saved. Primary: '+selectedSchIds[0]+' (KMPL '+kmpl1st+'). Linked: '+(saved.length-1)+'.'
    };
  } catch(e) {
    return {ok:false, error:e.message};
  }
}

function getExistingEntry(dateStr, schId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return {tc:[], diesel:null, service:null};
    detectColumns(sh);
    var rows = sh.getDataRange().getValues();
    var tc=[], diesel=null, service=null;
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      if (formatDateVal(r[0]) !== dateStr) continue;
      if (String(getCol(r,'schId')||r[2]) !== String(schId)) continue;
      var mod = String(getCol(r,'mod')||r[1]);
      if (mod === 'TC') {
        tc.push({
          schNo    : getCol(r,'schNo')     || r[3] || '',
          route    : getCol(r,'route')     || r[4] || '',
          sanKm    : parseFloat(getCol(r,'sanKm'))   || 0,
          optKm    : parseFloat(getCol(r,'optKm'))   || 0,
          cancelKm : parseFloat(getCol(r,'cancelKm'))|| 0,
          status   : getCol(r,'status')   || 'YES',
          tripType : getCol(r,'tripType') || ''
        });
      } else if (mod === 'DIESEL') {
        diesel = {
          vehicle      : getCol(r,'vehicle')      || '',
          vehicleType  : getCol(r,'vehType')      || '',
          driverBadge  : getCol(r,'driverBadge')  || '',
          driverName   : getCol(r,'driverName')   || '',
          driverMobile : getCol(r,'driverMobile') || '',
          diesel       : parseFloat(getCol(r,'diesel'))  || 0,
          uria         : parseFloat(getCol(r,'uria'))    || 0,
          kmpl         : parseFloat(getCol(r,'kmpl'))    || 0,
          deadKm       : parseFloat(getCol(r,'deadKm'))  || 0,
          divKm        : parseFloat(getCol(r,'divKm'))   || 0,
          optKm        : parseFloat(getCol(r,'optKm'))   || 0
        };
      } else if (mod === 'SERVICE') {
        service = {
          condBadge  : getCol(r,'condBadge')  || '',
          condName   : getCol(r,'condName')   || '',
          condMobile : getCol(r,'condMobile') || '',
          bagNo      : getCol(r,'bagNo')      || '',
          cash       : parseFloat(getCol(r,'cash'))     || 0,
          allPass    : parseFloat(getCol(r,'allPass'))  || 0,
          students   : parseFloat(getCol(r,'students')) || 0,
          dhaba      : parseFloat(getCol(r,'dhaba'))    || 0,
          epkm       : parseFloat(getCol(r,'epkm'))     || 0
        };
      }
    }
    return {tc:tc, diesel:diesel, service:service};
  } catch(e) {
    return {error: e.message, tc:[], diesel:null, service:null};
  }
}

function getEntryStatus(dateStr, schId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return {error:'No COMBINE DATA sheet',tc:{done:false},diesel:{done:false},service:{done:false},canEnterDiesel:false,canEnterService:false};
    detectColumns(sh);
    var rows = sh.getDataRange().getValues();
    var tcRows=[],dieselRows=[],serviceRows=[];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var rDate = formatDateVal(r[0]);
      var rMod  = String(getCol(r,'mod'));
      var rSch  = String(getCol(r,'schId'));
      if (rDate !== dateStr || rSch !== String(schId)) continue;
      if (rMod === 'TC')      tcRows.push(r);
      if (rMod === 'DIESEL')  dieselRows.push(r);
      if (rMod === 'SERVICE') serviceRows.push(r);
    }
    var tcDone  = tcRows.length > 0;
    var tcOptKm = tcRows.reduce(function(s,r){return s+(parseFloat(getCol(r,'optKm'))||0);},0);
    var tcYes   = tcRows.filter(function(r){return String(getCol(r,'status'))==='YES';}).length;
    var tcNo    = tcRows.filter(function(r){return String(getCol(r,'status'))==='NO';}).length;
    var tcBD    = tcRows.filter(function(r){return String(getCol(r,'status'))==='BD';}).length;
    var dieselDone   = dieselRows.length > 0;
    var dieselVeh    = dieselRows.length ? String(getCol(dieselRows[0],'vehicle')) : '';
    var dieselDriver = dieselRows.length ? String(getCol(dieselRows[0],'driverName')||getCol(dieselRows[0],'driverBadge')) : '';
    var dieselLitres = dieselRows.reduce(function(s,r){return s+(parseFloat(getCol(r,'diesel'))||0);},0);
    var dkmpl = 0;
    dieselRows.forEach(function(r){dkmpl += parseFloat(getCol(r,'kmpl'))||0;});
    if (dieselRows.length) dkmpl = dkmpl/dieselRows.length;
    var serviceDone = serviceRows.length > 0;
    var serviceCond = serviceRows.length ? String(getCol(serviceRows[0],'condName')||getCol(serviceRows[0],'condBadge')) : '';
    var serviceCash = serviceRows.reduce(function(s,r){return s+(parseFloat(getCol(r,'cash'))||0);},0);
    var sepkm = 0;
    serviceRows.forEach(function(r){sepkm += parseFloat(getCol(r,'epkm'))||0;});
    if (serviceRows.length) sepkm = sepkm/serviceRows.length;
    var tcAllCancelled = tcDone && tcYes === 0 && (tcNo + tcBD) > 0;
    return {
      date:dateStr, schId:schId,
      tc:{done:tcDone,trips:tcRows.length,optKm:parseFloat(tcOptKm.toFixed(2)),yes:tcYes,no:tcNo,bd:tcBD},
      diesel:{done:dieselDone,vehicle:dieselVeh,driver:dieselDriver,litres:parseFloat(dieselLitres.toFixed(2)),kmpl:parseFloat(dkmpl.toFixed(2))},
      service:{done:serviceDone,conductor:serviceCond,cash:serviceCash,epkm:parseFloat(sepkm.toFixed(2))},
      tcAllCancelled: tcAllCancelled,
      canEnterDiesel: tcDone && !tcAllCancelled,
      canEnterService: tcDone && !tcAllCancelled,
      pendingModules:[!tcDone?'TC':null,(!dieselDone&&!tcAllCancelled)?'DIESEL':null,(!serviceDone&&!tcAllCancelled)?'SERVICE':null].filter(Boolean),
      allComplete: tcDone && (tcAllCancelled || (dieselDone && serviceDone))
    };
  } catch(e) {
    return {error:e.message,tc:{done:false},diesel:{done:false},service:{done:false},canEnterDiesel:false,canEnterService:false};
  }
}

function getDailyEntryStatus(dateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return {date:dateStr,schedules:[],total:0,complete:0,pending:0,fresh:true};
    detectColumns(sh);
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return {date:dateStr,schedules:[],total:0,complete:0,pending:0,fresh:true};
    var schMap = {};
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var rDate = formatDateVal(r[0]);
      if (rDate !== dateStr) continue;
      var schId = String(getCol(r,'schId'));
      var mod   = String(getCol(r,'mod'));
      if (!schId) continue;
      if (!schMap[schId]) schMap[schId] = {schId:schId,tc:false,diesel:false,service:false};
      if (mod === 'TC')      schMap[schId].tc      = true;
      if (mod === 'DIESEL')  schMap[schId].diesel  = true;
      if (mod === 'SERVICE') schMap[schId].service = true;
    }
    var result = [];
    Object.keys(schMap).sort().forEach(function(id) {
      var s = schMap[id];
      result.push({schId:s.schId,tc:s.tc,diesel:s.diesel,service:s.service,
        complete:s.tc&&s.diesel&&s.service,
        status:s.tc&&s.diesel&&s.service?'COMPLETE':s.tc?'TC_DONE':'PENDING'});
    });
    return {date:dateStr,schedules:result,total:result.length,
      complete:result.filter(function(x){return x.complete;}).length,
      pending:result.filter(function(x){return !x.complete;}).length};
  } catch(e) {
    return {error:e.message,schedules:[],total:0,complete:0,pending:0};
  }
}

// ═══════════════════════════════════════════════════════════════
// SMART FEATURES — ANOMALY DETECTION & BASELINES  (v3)
// ═══════════════════════════════════════════════════════════════

// ── getSmartBaseline ──────────────────────────────────────────
// Returns historical avg/stddev for KMPL (per schId) and EPKM
// over the last `days` days (default 30), used for form validation.
//
// params: { schId, days }
// Returns: { kmpl:{avg,std,min,max,n}, epkm:{avg,std,min,max,n} }
// ─────────────────────────────────────────────────────────────
function getSmartBaseline(params) {
  try {
    var schId  = String(params.schId || '').trim();
    var days   = parseInt(params.days) || 30;
    var toDate = new Date();
    var fromDate = new Date();
    fromDate.setDate(fromDate.getDate() - days);
    var from = fmtDate(fromDate);
    var to   = fmtDate(toDate);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { error: 'Sheet not found' };
    detectColumns(sh);
    var rows = sh.getDataRange().getValues();

    var kmplVals = [], epkmVals = [];

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var dt  = formatDateVal(r[0]);
      if (dt < from || dt > to) continue;
      var mod = String(getCol(r,'mod'));
      var sid = String(getCol(r,'schId') || '');
      if (schId && sid !== schId) continue;

      if (mod === 'DIESEL') {
        var k = parseFloat(getCol(r,'kmpl')) || 0;
        if (k > 0) kmplVals.push(k);
      }
      if (mod === 'SERVICE') {
        var e = parseFloat(getCol(r,'epkm')) || 0;
        if (e > 0) epkmVals.push(e);
      }
    }

    function stats(arr) {
      if (!arr.length) return { avg:0, std:0, min:0, max:0, n:0 };
      var n = arr.length;
      var avg = arr.reduce(function(s,v){return s+v;},0) / n;
      var variance = arr.reduce(function(s,v){return s+Math.pow(v-avg,2);},0) / n;
      var std = Math.sqrt(variance);
      var sorted = arr.slice().sort(function(a,b){return a-b;});
      return {
        avg: parseFloat(avg.toFixed(2)),
        std: parseFloat(std.toFixed(2)),
        min: parseFloat(sorted[0].toFixed(2)),
        max: parseFloat(sorted[sorted.length-1].toFixed(2)),
        n:   n
      };
    }

    return { kmpl: stats(kmplVals), epkm: stats(epkmVals), schId: schId, days: days };
  } catch(e) {
    return { error: e.message };
  }
}

// ── getAnomalies ──────────────────────────────────────────────
// Scans COMBINE DATA for anomalies over the last `days` days:
//  - KMPL outliers (per schedule baseline, z > 2.5)
//  - EPKM outliers (per conductor baseline, z > 2.5)
//  - Schedules with TC but missing DIESEL or SERVICE entry
//  - Schedules with unusually high cancellation rate (> depotAvg + 1 std)
//  - KMPL drop vs prior period (>15% drop)
//
// params: { days, lookbackDays }
// days        — recent window to scan for anomalies (default 3)
// lookbackDays — historical window for baselines  (default 30)
// ─────────────────────────────────────────────────────────────
function getAnomalies(params) {
  try {
    var days        = parseInt((params||{}).days)        || 3;
    var lookback    = parseInt((params||{}).lookbackDays) || 30;
    var now         = new Date();
    var scanTo      = fmtDate(now);
    var scanFromDt  = new Date(); scanFromDt.setDate(now.getDate() - days + 1);
    var scanFrom    = fmtDate(scanFromDt);
    var baseFromDt  = new Date(); baseFromDt.setDate(now.getDate() - lookback);
    var baseFrom    = fmtDate(baseFromDt);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { alerts:[], summary:{ total:0 } };
    detectColumns(sh);
    var allRows = sh.getDataRange().getValues();

    // ── 1. Build baselines from lookback window ──────────────
    // Per-schedule KMPL baseline
    var schKmplBase = {};   // schId -> [kmpl]
    // Per-schedule EPKM baseline
    var schEpkmBase = {};   // schId -> [epkm]
    // Depot-wide cancellation rate baseline
    var baseCanRates = []; // per-day: yesTrips/total

    var baseDayMap = {};
    for (var i = 1; i < allRows.length; i++) {
      var r = allRows[i];
      if (!r[0]) continue;
      var dt  = formatDateVal(r[0]);
      if (dt < baseFrom || dt > scanTo) continue;
      var mod = String(getCol(r,'mod'));
      var sid = String(getCol(r,'schId') || r[2] || '');

      if (mod === 'DIESEL') {
        var k = parseFloat(getCol(r,'kmpl')) || 0;
        if (k > 0) {
          if (!schKmplBase[sid]) schKmplBase[sid] = [];
          schKmplBase[sid].push(k);
        }
      }
      if (mod === 'SERVICE') {
        var e = parseFloat(getCol(r,'epkm')) || 0;
        if (e > 0) {
          if (!schEpkmBase[sid]) schEpkmBase[sid] = [];
          schEpkmBase[sid].push(e);
        }
      }
      if (mod === 'TC') {
        if (!baseDayMap[dt]) baseDayMap[dt] = { total:0, yes:0 };
        baseDayMap[dt].total++;
        if (String(getCol(r,'status')) === 'YES') baseDayMap[dt].yes++;
      }
    }

    Object.keys(baseDayMap).forEach(function(dt) {
      var d = baseDayMap[dt];
      if (d.total > 0) baseCanRates.push(1 - d.yes/d.total);
    });

    function meanStd(arr) {
      if (!arr.length) return { mean:0, std:0 };
      var n = arr.length;
      var mean = arr.reduce(function(s,v){return s+v;},0) / n;
      var std = Math.sqrt(arr.reduce(function(s,v){return s+Math.pow(v-mean,2);},0)/n);
      return { mean: mean, std: std };
    }

    var canMs = meanStd(baseCanRates);

    // Compute per-sch KMPL baseline stats (mean, std)
    var schKmplStats = {};
    Object.keys(schKmplBase).forEach(function(sid) {
      schKmplStats[sid] = meanStd(schKmplBase[sid]);
    });
    var schEpkmStats = {};
    Object.keys(schEpkmBase).forEach(function(sid) {
      schEpkmStats[sid] = meanStd(schEpkmBase[sid]);
    });

    // ── 2. Scan recent window ────────────────────────────────
    var alerts = [];
    var scanRows = allRows.slice(1).filter(function(r) {
      if (!r[0]) return false;
      var dt = formatDateVal(r[0]);
      return dt >= scanFrom && dt <= scanTo;
    });

    // Group TC rows by date+schId to find missing DIESEL/SERVICE
    var scanMod = {};
    scanRows.forEach(function(r) {
      var dt = formatDateVal(r[0]);
      var sid = String(getCol(r,'schId')||r[2]||'');
      var mod = String(getCol(r,'mod'));
      var key = dt + '|' + sid;
      if (!scanMod[key]) scanMod[key] = { sid:sid, dt:dt, tc:false, diesel:false, service:false, tcYes:0, tcTotal:0, optKm:0 };
      if (mod==='TC') {
        scanMod[key].tc = true;
        scanMod[key].tcTotal++;
        if (String(getCol(r,'status'))==='YES') {
          scanMod[key].tcYes++;
          scanMod[key].optKm += parseFloat(getCol(r,'optKm'))||0;
        }
      }
      if (mod==='DIESEL') scanMod[key].diesel = true;
      if (mod==='SERVICE') scanMod[key].service = true;
    });

    // Check missing entries
    Object.keys(scanMod).forEach(function(key) {
      var e = scanMod[key];
      if (!e.tc) return;
      var allCancelled = (e.tcYes === 0 && e.tcTotal > 0);
      if (allCancelled) return; // all cancelled — no diesel/service needed

      if (!e.diesel) {
        alerts.push({
          type: 'missing_entry',
          severity: 'warn',
          title: 'Fuel Entry Missing',
          desc: 'SCH ' + e.sid + ' on ' + e.dt + ' — TC done but no fuel entered',
          schId: e.sid, date: e.dt, metric: 'DIESEL'
        });
      }
      if (!e.service) {
        alerts.push({
          type: 'missing_entry',
          severity: 'warn',
          title: 'Cash Entry Missing',
          desc: 'SCH ' + e.sid + ' on ' + e.dt + ' — TC done but no cash entered',
          schId: e.sid, date: e.dt, metric: 'SERVICE'
        });
      }
    });

    // KMPL and EPKM outliers
    scanRows.forEach(function(r) {
      var dt  = formatDateVal(r[0]);
      var mod = String(getCol(r,'mod'));
      var sid = String(getCol(r,'schId')||r[2]||'');

      if (mod === 'DIESEL') {
        var kmpl = parseFloat(getCol(r,'kmpl')) || 0;
        if (kmpl <= 0) return;
        var base = schKmplStats[sid];
        if (base && base.std > 0.05) {
          var z = Math.abs(kmpl - base.mean) / base.std;
          if (z > 2.5) {
            var dir = kmpl < base.mean ? 'Low' : 'High';
            alerts.push({
              type: 'kmpl_outlier',
              severity: kmpl < base.mean ? 'crit' : 'warn',
              title: dir + ' KMPL — SCH ' + sid,
              desc: 'KMPL: ' + kmpl.toFixed(2) + ' vs avg ' + base.mean.toFixed(2) + ' (z=' + z.toFixed(1) + '). Check fuel entry for ' + dt,
              schId: sid, date: dt, metric: 'KMPL', value: kmpl, avg: base.mean
            });
          }
        } else if (!base || base.std <= 0.05) {
          // No history — check absolute range
          if (kmpl > 0 && kmpl < 2.0) {
            alerts.push({
              type: 'kmpl_outlier',
              severity: 'crit',
              title: 'Very Low KMPL — SCH ' + sid,
              desc: 'KMPL: ' + kmpl.toFixed(2) + ' on ' + dt + ' — below safe minimum (2.0)',
              schId: sid, date: dt, metric: 'KMPL', value: kmpl, avg: 0
            });
          } else if (kmpl > 8) {
            alerts.push({
              type: 'kmpl_outlier',
              severity: 'warn',
              title: 'Suspiciously High KMPL — SCH ' + sid,
              desc: 'KMPL: ' + kmpl.toFixed(2) + ' on ' + dt + ' — exceeds expected range. Recheck entry.',
              schId: sid, date: dt, metric: 'KMPL', value: kmpl, avg: 0
            });
          }
        }
      }

      if (mod === 'SERVICE') {
        var epkm = parseFloat(getCol(r,'epkm')) || 0;
        if (epkm <= 0) return;
        var ebase = schEpkmStats[sid];
        if (ebase && ebase.std > 1) {
          var ze = Math.abs(epkm - ebase.mean) / ebase.std;
          if (ze > 2.5) {
            var dir2 = epkm < ebase.mean ? 'Low' : 'High';
            alerts.push({
              type: 'epkm_outlier',
              severity: epkm < ebase.mean ? 'crit' : 'warn',
              title: dir2 + ' EPKM — SCH ' + sid,
              desc: 'EPKM: ₹' + epkm.toFixed(2) + ' vs avg ₹' + ebase.mean.toFixed(2) + ' on ' + dt,
              schId: sid, date: dt, metric: 'EPKM', value: epkm, avg: ebase.mean
            });
          }
        } else if (epkm < 5) {
          alerts.push({
            type: 'epkm_outlier',
            severity: 'crit',
            title: 'Very Low EPKM — SCH ' + sid,
            desc: 'EPKM: ₹' + epkm.toFixed(2) + ' on ' + dt + ' — check cash entry',
            schId: sid, date: dt, metric: 'EPKM', value: epkm, avg: 0
          });
        }
      }
    });

    // Cancellation rate spike
    var recentDayMap = {};
    scanRows.forEach(function(r) {
      var dt  = formatDateVal(r[0]);
      if (String(getCol(r,'mod')) !== 'TC') return;
      if (!recentDayMap[dt]) recentDayMap[dt] = { total:0, yes:0 };
      recentDayMap[dt].total++;
      if (String(getCol(r,'status'))==='YES') recentDayMap[dt].yes++;
    });
    Object.keys(recentDayMap).forEach(function(dt) {
      var d = recentDayMap[dt];
      if (d.total < 3) return;
      var canRate = 1 - d.yes/d.total;
      if (canRate > 0 && canMs.mean >= 0 && canMs.std > 0) {
        var zc = (canRate - canMs.mean) / canMs.std;
        if (zc > 2.0) {
          alerts.push({
            type: 'cancellation_spike',
            severity: 'crit',
            title: 'High Cancellations on ' + dt,
            desc: Math.round(canRate*100) + '% trips cancelled — depot avg is ' + Math.round(canMs.mean*100) + '%. Investigate.',
            date: dt, metric: 'CANCEL_RATE', value: canRate, avg: canMs.mean
          });
        }
      }
    });

    // De-duplicate and sort by severity
    var seen = {};
    alerts = alerts.filter(function(a) {
      var k = a.type + '|' + (a.schId||'') + '|' + a.date + '|' + a.metric;
      if (seen[k]) return false;
      seen[k] = true;
      return true;
    });

    var sev = { crit:0, warn:1, info:2 };
    alerts.sort(function(a,b) {
      return (sev[a.severity]||2) - (sev[b.severity]||2);
    });

    // Summary counts
    var crits  = alerts.filter(function(a){return a.severity==='crit';}).length;
    var warns  = alerts.filter(function(a){return a.severity==='warn';}).length;

    return {
      alerts: alerts,
      summary: {
        total:  alerts.length,
        crits:  crits,
        warns:  warns,
        scanFrom: scanFrom,
        scanTo:   scanTo
      }
    };
  } catch(e) {
    Logger.log('getAnomalies error: ' + e.message);
    return { alerts:[], summary:{ total:0 }, error: e.message };
  }
}

// ── getTrendInsights ──────────────────────────────────────────
// Returns plain-text trend insights comparing this week vs last week:
//  KMPL trend, EPKM trend, cancellation rate, operated KM.
// Used by the dashboard Smart Insights panel.
// ─────────────────────────────────────────────────────────────
function getTrendInsights() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { insights:[] };
    detectColumns(sh);

    var now  = new Date();
    var today = fmtDate(now);

    // This week: Mon–today
    var thisMon = new Date(now);
    var dow = now.getDay(); // 0=Sun
    thisMon.setDate(now.getDate() - (dow===0?6:dow-1));
    var thisFrom = fmtDate(thisMon);

    // Last week: prev Mon – prev Sun
    var lastMon = new Date(thisMon); lastMon.setDate(thisMon.getDate()-7);
    var lastSun = new Date(thisMon); lastSun.setDate(thisMon.getDate()-1);
    var lastFrom = fmtDate(lastMon);
    var lastTo   = fmtDate(lastSun);

    var rows = sh.getDataRange().getValues();

    function weekStats(from, to) {
      var tcTotal=0, tcYes=0, optKm=0, dsum=0, dc=0, esum=0, ec=0;
      for (var i=1;i<rows.length;i++) {
        var r=rows[i]; if(!r[0])continue;
        var dt=formatDateVal(r[0]);
        if(dt<from||dt>to)continue;
        var mod=String(getCol(r,'mod'));
        if(mod==='TC'){
          tcTotal++;
          if(String(getCol(r,'status'))==='YES'){tcYes++;optKm+=parseFloat(getCol(r,'optKm'))||0;}
        }
        if(mod==='DIESEL'){var k=parseFloat(getCol(r,'kmpl'))||0;if(k>0){dsum+=k;dc++;}}
        if(mod==='SERVICE'){var e=parseFloat(getCol(r,'epkm'))||0;if(e>0){esum+=e;ec++;}}
      }
      return {
        compRate: tcTotal>0?tcYes/tcTotal:null,
        optKm:    optKm,
        avgKmpl:  dc>0?dsum/dc:null,
        avgEpkm:  ec>0?esum/ec:null
      };
    }

    var tw = weekStats(thisFrom, today);
    var lw = weekStats(lastFrom, lastTo);

    function pctDiff(curr, prev) {
      if (curr===null||prev===null||prev===0) return null;
      return ((curr - prev) / prev) * 100;
    }
    function arrowStr(pct, higherIsBetter) {
      if (pct===null) return null;
      var good = higherIsBetter ? pct >= 0 : pct <= 0;
      var abs  = Math.abs(pct).toFixed(1);
      return { pct: parseFloat(pct.toFixed(1)), good: good, abs: abs };
    }

    var insights = [];

    var kmplDiff = arrowStr(pctDiff(tw.avgKmpl, lw.avgKmpl), true);
    if (kmplDiff && Math.abs(kmplDiff.pct) >= 2) {
      insights.push({
        metric: 'KMPL',
        title: 'Fleet Efficiency ' + (kmplDiff.good?'▲ Up':'▼ Down') + ' ' + kmplDiff.abs + '% vs last week',
        detail: 'This week: ' + (tw.avgKmpl?tw.avgKmpl.toFixed(2):'—') + '  |  Last week: ' + (lw.avgKmpl?lw.avgKmpl.toFixed(2):'—'),
        good: kmplDiff.good, pct: kmplDiff.pct
      });
    }

    var epkmDiff = arrowStr(pctDiff(tw.avgEpkm, lw.avgEpkm), true);
    if (epkmDiff && Math.abs(epkmDiff.pct) >= 2) {
      insights.push({
        metric: 'EPKM',
        title: 'Revenue/KM ' + (epkmDiff.good?'▲ Up':'▼ Down') + ' ' + epkmDiff.abs + '% vs last week',
        detail: 'This week: ₹' + (tw.avgEpkm?tw.avgEpkm.toFixed(2):'—') + '  |  Last week: ₹' + (lw.avgEpkm?lw.avgEpkm.toFixed(2):'—'),
        good: epkmDiff.good, pct: epkmDiff.pct
      });
    }

    var compDiff = arrowStr(pctDiff(tw.compRate, lw.compRate), true);
    if (compDiff && Math.abs(compDiff.pct) >= 2) {
      insights.push({
        metric: 'COMPLETION',
        title: 'Trip Completion ' + (compDiff.good?'▲ Up':'▼ Down') + ' ' + compDiff.abs + '% vs last week',
        detail: 'This week: ' + (tw.compRate!==null?Math.round(tw.compRate*100):'—') + '%  |  Last week: ' + (lw.compRate!==null?Math.round(lw.compRate*100):'—') + '%',
        good: compDiff.good, pct: compDiff.pct
      });
    }

    var kmDiff = arrowStr(pctDiff(tw.optKm, lw.optKm), true);
    if (kmDiff && Math.abs(kmDiff.pct) >= 3) {
      insights.push({
        metric: 'OPT_KM',
        title: 'Operated KM ' + (kmDiff.good?'▲ Up':'▼ Down') + ' ' + kmDiff.abs + '% vs last week',
        detail: 'This week: ' + Math.round(tw.optKm) + ' km  |  Last week: ' + Math.round(lw.optKm) + ' km',
        good: kmDiff.good, pct: kmDiff.pct
      });
    }

    if (!insights.length) {
      insights.push({
        metric: 'ALL',
        title: 'Performance stable vs last week',
        detail: 'No significant changes detected in KMPL, EPKM, or completion rate.',
        good: true, pct: 0
      });
    }

    return {
      insights: insights,
      thisWeek: { from:thisFrom, to:today, ...tw },
      lastWeek: { from:lastFrom, to:lastTo, ...lw }
    };
  } catch(e) {
    Logger.log('getTrendInsights error: ' + e.message);
    return { insights:[], error: e.message };
  }
}


// ═══════════════════════════════════════════════════════════════
// ONLY CASH MODULE
// ═══════════════════════════════════════════════════════════════

/**
 * Returns all items from the ONLY CASH master sheet.
 * Col 1 = ID, Col 2 = Name
 */
function getOnlyCashItems() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ONLY_CASH_SH);
    if (!sh) return { error: 'ONLY CASH sheet not found' };
    var data = sh.getDataRange().getValues();
    var items = [];
    for (var i = 0; i < data.length; i++) {
      var id   = String(data[i][0] || '').trim();
      var name = String(data[i][1] || '').trim();
      if (id) items.push({ id: id, name: name });
    }
    return items;
  } catch(e) {
    return { error: e.message };
  }
}

/**
 * Save an Only Cash entry to COMBINE DATA with type ONLY_CASH.
 * data: { date, itemId, itemName, cash, passengers }
 */
function saveOnlyCashEntry(data) {
  try {
    var sh = ensureHeaders();
    var now = new Date();
    var dateStr = data.date
      ? String(data.date).substring(0, 10)
      : Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Delete existing entry for same date + item to avoid duplicates
    var all = sh.getDataRange().getValues();
    var toDelete = [];
    for (var i = all.length - 1; i >= 1; i--) {
      if (formatDateVal(all[i][0]) === dateStr &&
          String(all[i][1]) === 'ONLY_CASH' &&
          String(all[i][2]) === String(data.itemId)) {
        toDelete.push(i + 1);
      }
    }
    for (var _d = toDelete.length - 1; _d >= 0; _d--) { sh.deleteRow(toDelete[_d]); }

    var cashVal = parseFloat(data.cash);
    var passVal = parseFloat(data.passengers);
    if (isNaN(cashVal)) cashVal = 0;
    if (isNaN(passVal)) passVal = 0;

    var row = new Array(HEADERS.length).fill('');
    row[0]  = dateStr;
    row[1]  = 'ONLY_CASH';
    row[2]  = String(data.itemId   || '');
    row[3]  = String(data.itemId   || '');   // schNo = itemId
    row[4]  = String(data.itemName || '');   // route = item name
    row[5]  = 'ONLY_CASH';
    row[6]  = 0;
    row[7]  = 0;
    row[8]  = 0;
    row[9]  = 'ONLY_CASH';
    row[24] = cashVal;
    row[25] = passVal;
    row[29] = now;
    sh.getRange(sh.getLastRow() + 1, 1, 1, HEADERS.length).setValues([row]);

    return { success: true, msg: 'Only Cash saved: ' + data.itemId + ' on ' + dateStr };
  } catch(e) {
    return { error: e.message };
  }
}

/**
 * Get existing Only Cash entry for a given date + itemId (for edit mode).
 */
function getOnlyCashEntry(date, itemId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return null;
    var all = sh.getDataRange().getValues();
    var dateStr = String(date).substring(0, 10);
    for (var i = 1; i < all.length; i++) {
      if (formatDateVal(all[i][0]) === dateStr &&
          String(all[i][1]) === 'ONLY_CASH' &&
          String(all[i][2]) === String(itemId)) {
        return {
          itemId:     String(all[i][2]),
          itemName:   String(all[i][4]),
          cash:       parseFloat(all[i][24]) || 0,
          passengers: parseFloat(all[i][25]) || 0
        };
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}


/**
 * getOnlyCashReport(from, to)
 * Returns item-wise Only Cash breakdown for a date range.
 * Result: { rows: [{date, itemId, itemName, cash, passengers}], 
 *           byItem: [{itemId, itemName, totalCash, totalPass, days}],
 *           grandTotal: number }
 */
function getOnlyCashReport(from, to) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { error: 'COMBINE DATA sheet not found' };
    detectColumns(sh);
    var all  = sh.getDataRange().getValues();
    var fromStr = String(from).substring(0, 10);
    var toStr   = String(to).substring(0, 10);

    var rows   = [];   // every ONLY_CASH row in range
    var byItem = {};   // itemId → { itemName, totalCash, totalPass, days: Set }

    for (var i = 1; i < all.length; i++) {
      var dt = formatDateVal(all[i][0]);
      if (!dt) continue;
      if (dt < fromStr || dt > toStr) continue;
      if (String(all[i][1]).trim() !== 'ONLY_CASH') continue;

      var itemId   = String(all[i][2] || '').trim();
      var itemName = String(all[i][4] || itemId).trim();
      var cash     = parseFloat(all[i][24]) || 0;
      var pass     = parseFloat(all[i][25]) || 0;

      rows.push({ date: dt, itemId: itemId, itemName: itemName, cash: cash, passengers: pass });

      if (!byItem[itemId]) byItem[itemId] = { itemId: itemId, itemName: itemName, totalCash: 0, totalPass: 0, days: [] };
      byItem[itemId].totalCash += cash;
      byItem[itemId].totalPass += pass;
      if (byItem[itemId].days.indexOf(dt) === -1) byItem[itemId].days.push(dt);
    }

    // Sort rows by date
    rows.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

    // Convert byItem to array sorted by totalCash desc
    var itemArr = Object.keys(byItem).map(function(k) {
      var x = byItem[k];
      return { itemId: x.itemId, itemName: x.itemName, totalCash: x.totalCash, totalPass: x.totalPass, dayCount: x.days.length };
    }).sort(function(a, b) { return b.totalCash - a.totalCash; });

    var grandTotal = itemArr.reduce(function(s, x) { return s + x.totalCash; }, 0);

    return { rows: rows, byItem: itemArr, grandTotal: grandTotal, from: fromStr, to: toStr };
  } catch(e) {
    Logger.log('getOnlyCashReport error: ' + e.message);
    return { error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// SANCTIONED DATA HELPER
// Reads directly from MAIN DATA:
//   Col C (index 2) = Sch No
//   Col E (index 4) = Service/Trip Type
//   Col H (index 7) = KM per trip row
// Returns:
//   bySchId[schId]  = { schNo, type, sanTrips, sanKm }
//   byType[type]    = { sanSch (unique schNos), sanTrips, sanKm }
// ═══════════════════════════════════════════════════════════════
function buildSanctionedMap() {
  var result = { bySchId: {}, byType: {} };
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var mainSh = ss.getSheetByName(MAIN_DATA);
    if (!mainSh) return result;

    var rows   = mainSh.getDataRange().getValues();
    var lastId = '';   // col A master schId (spans multiple rows)

    for (var i = 1; i < rows.length; i++) {
      // Col A (index 0): master schedule ID — only written on first row of a group
      var colA = rows[i][0];
      if (colA !== null && colA !== undefined && String(colA).trim() !== '') {
        lastId = String(colA).trim();
      }
      if (!lastId) continue;

      var schNo = String(rows[i][3] || '').replace(/\.0$/, '').trim(); // Col D = SCH NO
      var type  = String(rows[i][5] || '').toUpperCase().trim();       // Col F = TYPE
      var km    = parseFloat(rows[i][8]) || 0;                         // Col I = KM

      if (!schNo) continue;  // skip blank rows

      // Per-schId accumulator
      if (!result.bySchId[lastId]) {
        result.bySchId[lastId] = { schNo: schNo, type: type || 'GNR', sanTrips: 0, sanKm: 0, schNos: new Set() };
      }
      result.bySchId[lastId].sanTrips += 1;
      result.bySchId[lastId].sanKm    += km;
      result.bySchId[lastId].schNos.add(schNo);
      if (!result.bySchId[lastId].type && type) result.bySchId[lastId].type = type;

      // Per-type accumulator
      var typeKey = type || 'GNR';
      if (!result.byType[typeKey]) {
        result.byType[typeKey] = { sanSchNos: new Set(), sanTrips: 0, sanKm: 0 };
      }
      result.byType[typeKey].sanSchNos.add(schNo);
      result.byType[typeKey].sanTrips += 1;
      result.byType[typeKey].sanKm    += km;
    }

    // Finalise Sets → counts
    Object.keys(result.byType).forEach(function(t) {
      result.byType[t].sanSch = result.byType[t].sanSchNos.size;
    });
    Object.keys(result.bySchId).forEach(function(id) {
      result.bySchId[id].sanSch = result.bySchId[id].schNos.size;
    });

  } catch(e) {
    Logger.log('buildSanctionedMap error: ' + e.message);
  }
  return result;
}

// ═══════════════════════════════════════════════════════════════
// SERVICE WISE REPORT
// Groups all TC + DIESEL + SERVICE rows by service/trip type
// (EXP, GNR, SLP, MINI, MIDI, DELUXE, LOCAL, etc.)
// Returns sanctioned, operated, cancelled metrics + cash/KMPL/EPKM
// ═══════════════════════════════════════════════════════════════

function getServiceWiseReport(from, to) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { error: 'COMBINE DATA sheet not found' };
    detectColumns(sh);

    var fromStr = String(from).substring(0, 10);
    var toStr   = String(to).substring(0, 10);
    var all     = sh.getDataRange().getValues();

    // ── Sanctioned baseline directly from MAIN DATA ──────────
    var sanData = buildSanctionedMap();
    var sanBySchId = sanData.bySchId;   // schId → { type, sanTrips, sanKm, sanSch }
    var sanByType  = sanData.byType;    // type  → { sanSch, sanTrips, sanKm }

    // ── Per-schedule accumulators (operated/cancelled/cash) ──
    var schAcc = {};

    function getAcc(schId, typeHint) {
      if (!schAcc[schId]) {
        var san = sanBySchId[schId] || { type: typeHint || 'GNR', sanKm: 0, sanTrips: 0 };
        schAcc[schId] = {
          schId    : schId,
          type     : String(san.type || typeHint || 'GNR').toUpperCase(),
          // TC accumulated
          optSchDates : {},
          optTrips    : 0,
          optKm       : 0,
          canTrips    : 0,
          canKm       : 0,
          // DIESEL
          diesel      : 0,
          kmpl        : 0,  kmplN: 0,
          deadKm      : 0,
          // SERVICE
          cash        : 0,
          pass        : 0,
          students    : 0,
          dhaba       : 0,
          epkm        : 0,  epkmN: 0
        };
      }
      return schAcc[schId];
    }

    var ocCashAccum = 0; // used only for getReportData, kept for safety

    for (var i = 1; i < all.length; i++) {
      var dateVal = formatDateVal(all[i][0]);
      if (!dateVal) continue;
      if (dateVal < fromStr || dateVal > toStr) continue;

      var mod    = String(all[i][1] || '').trim();
      var schId  = String(all[i][2] || '').trim();
      var tType  = String(getCol(all[i], 'tripType') || '').toUpperCase();
      if (!schId) continue;

      var acc = getAcc(schId, tType);
      // Carry forward type from data if not set from MAIN DATA
      if (!acc.type && tType) acc.type = tType;

      if (mod === 'TC') {
        var status  = String(all[i][9] || '').toUpperCase();
        var optKm   = parseFloat(getCol(all[i], 'optKm'))    || 0;
        var canKm   = parseFloat(getCol(all[i], 'cancelKm')) || 0;
        if (status === 'YES') {
          acc.optTrips += 1;
          acc.optKm    += optKm;
          acc.optSchDates[dateVal] = true;
        } else if (status === 'NO' || status === 'BD') {
          acc.canTrips += 1;
          acc.canKm    += canKm;
        }
      } else if (mod === 'DIESEL') {
        acc.diesel += parseFloat(getCol(all[i], 'diesel')) || 0;
        var kv = parseFloat(getCol(all[i], 'kmpl')) || 0;
        if (kv > 0) { acc.kmpl += kv; acc.kmplN++; }
        acc.deadKm += parseFloat(getCol(all[i], 'deadKm')) || 0;
      } else if (mod === 'SERVICE') {
        acc.cash     += parseFloat(getCol(all[i], 'cash'))     || 0;
        acc.pass     += parseFloat(getCol(all[i], 'allPass'))  || 0;
        acc.students += parseFloat(getCol(all[i], 'students')) || 0;
        acc.dhaba    += parseFloat(getCol(all[i], 'dhaba'))    || 0;
        var ev = parseFloat(getCol(all[i], 'epkm')) || 0;
        if (ev > 0) { acc.epkm += ev; acc.epkmN++; }
      } else if (mod === 'ONLY_CASH') {
        // Only Cash entries — store cash in acc so it flows into typeMap['ONLY_CASH'] row
        acc.cash += parseFloat(getCol(all[i], 'cash')) || 0;
        acc.pass += parseFloat(getCol(all[i], 'allPass')) || 0;
        ocCashAccum += parseFloat(getCol(all[i], 'cash')) || 0;
      }
    }

    // ── Aggregate by service type ─────────────────────────────
    // Start typeMap from MAIN DATA sanctioned baseline
    var typeMap = {};
    var SERVICE_ORDER = ['EXP','SLP','DELUXE','MIDI','MINI','LOCAL','GNR','UNKNOWN'];

    // Seed every type that exists in MAIN DATA
    Object.keys(sanByType).forEach(function(key) {
      var s = sanByType[key];
      typeMap[key] = {
        serviceType : key,
        sanSch      : s.sanSch,
        sanTrips    : s.sanTrips,
        sanKm       : Math.round(s.sanKm * 10) / 10,
        optSch      : 0,
        optTrips    : 0,
        optKm       : 0,
        canSch      : 0,
        canTrips    : 0,
        canKm       : 0,
        cash        : 0,
        epkmSum     : 0, epkmN: 0,
        kmplSum     : 0, kmplN: 0,
        pass        : 0,
        students    : 0,
        dhaba       : 0,
        cashSum     : 0,  // for correct EPKM = totalCash / totalOptKm
        optKmSum    : 0
      };
    });

    // Merge operated/cancelled/cash from COMBINE DATA accumulators
    Object.keys(schAcc).forEach(function(schId) {
      var a   = schAcc[schId];
      // Determine type: prefer MAIN DATA type, fallback to TC row type
      var key = (sanBySchId[schId] && sanBySchId[schId].type) ? sanBySchId[schId].type : (a.type || 'GNR');
      if (!key) key = 'GNR';

      if (!typeMap[key]) {
        typeMap[key] = {
          serviceType : key,
          sanSch      : 0, sanTrips: 0, sanKm: 0,
          optSch      : 0, optTrips: 0, optKm: 0,
          canSch      : 0, canTrips: 0, canKm: 0,
          cash: 0, epkmSum: 0, epkmN: 0,
          kmplSum: 0, kmplN: 0,
          pass: 0, students: 0, dhaba: 0,
          cashSum: 0, optKmSum: 0
        };
      }
      var t = typeMap[key];

      var optSchCount = Object.keys(a.optSchDates).length;
      t.optSch    += optSchCount > 0 ? 1 : 0;
      t.optTrips  += a.optTrips;
      t.optKm     += a.optKm;
      t.canSch    += (a.canTrips > 0 && a.optTrips === 0) ? 1 : 0;  // only fully cancelled sch (0 trips operated)
      t.canTrips  += a.canTrips;
      t.canKm     += a.canKm;
      t.cash      += a.cash;
      t.pass      += a.pass;
      t.students  += a.students;
      t.dhaba     += a.dhaba;
      t.cashSum   += a.cash;
      t.optKmSum  += a.optKm;
      if (a.kmplN > 0) { t.kmplSum += (a.kmpl / a.kmplN); t.kmplN++; }
      if (a.epkmN > 0) { t.epkmSum += (a.epkm / a.epkmN); t.epkmN++; }
    });

    // ── Build result rows ─────────────────────────────────────
    var rows = [];
    var totals = {
      serviceType:'TOTAL', sanSch:0, sanTrips:0, sanKm:0,
      optSch:0, optTrips:0, optKm:0,
      canSch:0, canTrips:0, canKm:0,
      cash:0, epkm:0, kmpl:0, pass:0, students:0, dhaba:0
    };

    // Sort by defined order, then alphabetically for others
    var keys = Object.keys(typeMap).sort(function(a, b) {
      var ai = SERVICE_ORDER.indexOf(a), bi = SERVICE_ORDER.indexOf(b);
      if (ai === -1) ai = 99; if (bi === -1) bi = 99;
      return ai !== bi ? ai - bi : a.localeCompare(b);
    });

    keys.forEach(function(key) {
      var t = typeMap[key];
      var avgKmpl = t.kmplN > 0 ? Math.round((t.kmplSum / t.kmplN) * 100) / 100 : 0;
      // EPKM = total cash / total operated KM (correct revenue-per-km metric)
      var avgEpkm = (t.optKmSum > 0) ? Math.round((t.cashSum / t.optKmSum) * 100) / 100 : 0;
      var row = {
        serviceType : key,
        sanSch      : t.sanSch,
        sanTrips    : t.sanTrips,
        sanKm       : Math.round(t.sanKm * 10) / 10,
        optSch      : t.optSch,
        optTrips    : t.optTrips,
        optKm       : Math.round(t.optKm * 10) / 10,
        canSch      : t.canSch,
        canTrips    : t.canTrips,
        canKm       : Math.round(t.canKm * 10) / 10,
        cash        : Math.round(t.cash),
        epkm        : avgEpkm,
        kmpl        : avgKmpl,
        pass        : t.pass,
        students    : t.students,
        dhaba       : Math.round(t.dhaba)
      };
      rows.push(row);

      totals.sanSch    += t.sanSch;
      totals.sanTrips  += t.sanTrips;
      totals.sanKm     += t.sanKm;
      totals.optSch    += t.optSch;
      totals.optTrips  += t.optTrips;
      totals.optKm     += t.optKm;
      totals.canSch    += t.canSch;
      totals.canTrips  += t.canTrips;
      totals.canKm     += t.canKm;
      totals.cash      += t.cash;
      totals.pass      += t.pass;
      totals.students  += t.students;
      totals.dhaba     += t.dhaba;
      if (t.kmplN > 0) { totals.kmpl += t.kmplSum / t.kmplN; }
      if (t.epkmN > 0) { totals.epkm += t.epkmSum / t.epkmN; }
    });

    // Finalise totals averages
    totals.sanKm  = Math.round(totals.sanKm * 10) / 10;
    totals.optKm  = Math.round(totals.optKm * 10) / 10;
    totals.canKm  = Math.round(totals.canKm * 10) / 10;
    totals.cash   = Math.round(totals.cash);
    totals.dhaba  = Math.round(totals.dhaba);
    totals.kmpl   = rows.length > 0 ? Math.round((totals.kmpl / rows.length) * 100) / 100 : 0;
    // Totals EPKM = grand total cash / grand total operated KM
    var _totCashSum  = rows.reduce(function(s,r){ return s + (r.cash||0); }, 0);
    var _totOptKmSum = rows.reduce(function(s,r){ return s + (r.optKm||0); }, 0);
    totals.epkm   = _totOptKmSum > 0 ? Math.round((_totCashSum / _totOptKmSum) * 100) / 100 : 0;

    return { rows: rows, totals: totals, from: fromStr, to: toStr };

  } catch(e) {
    Logger.log('getServiceWiseReport error: ' + e.message + '\n' + e.stack);
    return { error: e.message };
  }
}


// ═══════════════════════════════════════════════════════════════
// RAW COMBINE DATA EXPORT — for Trip Report panel
// Returns rows in REPORT.HTML-compatible JSON format
// ═══════════════════════════════════════════════════════════════
function getCombineRawRows(from, to) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { combineData: [], schedules: [] };

    detectColumns(sh);
    var rows = sh.getDataRange().getValues();
    // filterRows with no module filter — get all modules
    var filtered = rows.slice(1).filter(function(r) {
      if (!r[0]) return false;
      var dt = formatDateVal(r[0]);
      if (from && dt < from) return false;
      if (to   && dt > to)   return false;
      return true;
    });

    var combineData = filtered.map(function(r) {
      return {
        date:          formatDateVal(r[0]),
        module:        String(getCol(r,'mod')      || ''),
        sch_id:        String(getCol(r,'schId')     || ''),
        sch_no:        String(getCol(r,'schNo')     || ''),
        route:         String(getCol(r,'route')     || ''),
        trip_type:     String(getCol(r,'tripType')  || ''),
        san_km:        parseFloat(getCol(r,'sanKm'))   || 0,
        opt_km:        parseFloat(getCol(r,'optKm'))   || 0,
        cancel_km:     parseFloat(getCol(r,'cancelKm'))|| 0,
        status:        String(getCol(r,'status')    || ''),
        vehicle_no:    String(getCol(r,'vehicle')   || ''),
        vehicle_type:  String(getCol(r,'vehType')   || ''),
        driver_badge:  String(getCol(r,'driverBadge')|| ''),
        driver_name:   String(getCol(r,'driverName') || ''),
        driver_mobile: String(getCol(r,'driverMobile')|| ''),
        diesel_l:      parseFloat(getCol(r,'diesel'))  || 0,
        uria_l:        parseFloat(getCol(r,'uria'))    || 0,
        kmpl:          parseFloat(getCol(r,'kmpl'))    || 0,
        dead_km:       parseFloat(getCol(r,'deadKm'))  || 0,
        diversion_km:  parseFloat(getCol(r,'divKm'))   || 0,
        conductor_badge: String(getCol(r,'condBadge') || ''),
        conductor_name:  String(getCol(r,'condName')  || ''),
        bag_no:        String(getCol(r,'bagNo')     || ''),
        cash_total:    parseFloat(getCol(r,'cash'))    || 0,
        all_pass:      parseInt(getCol(r,'allPass'))   || 0,
        student_pass:  parseInt(getCol(r,'students'))  || 0,
        dhaba:         parseFloat(getCol(r,'dhaba'))   || 0,
        epkm:          parseFloat(getCol(r,'epkm'))    || 0,
        entry_time:    String(getCol(r,'entryTime')  || '')
      };
    });

    // Build SCHEDULES from MAIN DATA
    // MAIN DATA cols: A(0)=schId, B(1)=SCH NAME, C(2)=SCH TYPE, D(3)=SCH NO,
    //                 E(4)=TRIP NO, F(5)=TYPE, G(6)=FROM PLACE, H(7)=END PLACE,
    //                 I(8)=KM, J(9)=TIME
    var schedules = [];
    var mainSh = ss.getSheetByName(MAIN_DATA);
    if (mainSh) {
      var mRows = mainSh.getDataRange().getValues();
      var lastId = '', tripNo = 0;
      for (var i = 1; i < mRows.length; i++) {
        var colA = mRows[i][0];
        if (colA !== null && colA !== undefined && String(colA).trim() !== '') {
          lastId = String(colA).trim();
          tripNo = 0;
        }
        if (!lastId) continue;
        var schType = String(mRows[i][2] || '').trim();                    // col C = SCH TYPE (overall)
        var schNo   = String(mRows[i][3] || '').replace(/\.0$/, '').trim();// col D = SCH NO
        var type    = String(mRows[i][5] || '').toUpperCase().trim();      // col F = TYPE (per-trip)
        var km      = parseFloat(mRows[i][8]) || 0;                        // col I = KM
        var name    = String(mRows[i][1] || '').trim();                    // col B = SCH NAME
        var fromPl  = String(mRows[i][6] || '').trim();                    // col G = FROM PLACE
        var toPl    = String(mRows[i][7] || '').trim();                    // col H = END PLACE
        var rawTime = mRows[i][9];                                         // col J = TIME
        if (!schNo && !name) continue;
        tripNo++;
        var timeStr = '';
        if (rawTime) {
          try {
            if (rawTime instanceof Date || (typeof rawTime === 'object' && rawTime.getHours)) {
              var th = rawTime.getHours ? rawTime.getHours() : 0;
              var tm = rawTime.getMinutes ? rawTime.getMinutes() : 0;
              timeStr = (th<10?'0':'')+th+':'+(tm<10?'0':'')+tm;
            } else { timeStr = String(rawTime).substring(0,5); }
          } catch(e) { timeStr = ''; }
        }
        schedules.push({
          id:       lastId,
          name:     name,
          sch_type: schType || type,
          sch_no:   parseInt(schNo) || tripNo,
          trip_no:  tripNo,
          type:     type,
          from:     fromPl,
          to:       toPl,
          km:       km,
          time:     timeStr
        });
      }
    }

    return { combineData: combineData, schedules: schedules };
  } catch(e) {
    Logger.log('getCombineRawRows error: ' + e.message);
    return { combineData: [], schedules: [], error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// VEHICLE-CENTRIC FUEL ENTRY — NEW FUNCTIONS
// Added for vehicle tile selector in Fuel Entry
// ═══════════════════════════════════════════════════════════════

/**
 * getVehicleTCData(date, vehicleNo)
 *
 * Returns all TC trips saved for a specific vehicle on a given date.
 * A vehicle may have run trips across multiple schedules.
 * Called by Fuel Entry (renderFuelForVehicle) in the HTML app.
 */
function getVehicleTCData(date, vehicleNo) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { error: 'COMBINE DATA sheet not found', schIds: [] };
    detectColumns(sh);

    var rows   = sh.getDataRange().getValues();
    var dStr   = String(date).substring(0, 10);
    var vehStr = String(vehicleNo).trim();

    var schIds       = [];
    var schIdSet     = {};
    var tripsBySchId = {};
    var totalOptKm   = 0;
    var totalTrips   = 0;
    var yesTrips     = 0;
    var driverBadge  = '';
    var driverName   = '';
    var driverMobile = '';

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      if (formatDateVal(r[0]) !== dStr) continue;
      if (String(getCol(r, 'mod')) !== 'TC') continue;

      var rowVeh = String(getCol(r, 'vehicle') || '').trim();
      if (rowVeh !== vehStr) continue;

      var schId  = String(getCol(r, 'schId')    || r[2] || '').trim();
      var schNo  = String(getCol(r, 'schNo')    || r[3] || '').trim();
      var route  = String(getCol(r, 'route')    || r[4] || '');
      var status = String(getCol(r, 'status')   || '');
      var optKm  = parseFloat(getCol(r, 'optKm')) || 0;
      var type   = String(getCol(r, 'tripType') || '');

      if (!driverBadge && getCol(r, 'driverBadge')) {
        driverBadge  = String(getCol(r, 'driverBadge')  || '');
        driverName   = String(getCol(r, 'driverName')   || '');
        driverMobile = String(getCol(r, 'driverMobile') || '');
      }

      if (!schIdSet[schId]) {
        schIdSet[schId]     = true;
        schIds.push(schId);
        tripsBySchId[schId] = [];
      }

      tripsBySchId[schId].push({
        schNo    : schNo,
        route    : route,
        status   : status,
        optKm    : optKm,
        tripType : type
      });

      totalTrips++;
      if (status === 'YES') {
        yesTrips++;
        totalOptKm += optKm;
      }
    }

    // Check for existing DIESEL entry for this vehicle
    var existingDiesel = null;
    for (var j = 1; j < rows.length; j++) {
      var dr = rows[j];
      if (!dr[0]) continue;
      if (formatDateVal(dr[0]) !== dStr) continue;
      if (String(getCol(dr, 'mod')) !== 'DIESEL') continue;
      if (String(getCol(dr, 'vehicle') || '').trim() !== vehStr) continue;

      existingDiesel = {
        vehicle      : vehStr,
        vehicleType  : String(getCol(dr, 'vehType')      || ''),
        driverBadge  : String(getCol(dr, 'driverBadge')  || ''),
        driverName   : String(getCol(dr, 'driverName')   || ''),
        driverMobile : String(getCol(dr, 'driverMobile') || ''),
        diesel       : parseFloat(getCol(dr, 'diesel'))  || 0,
        uria         : parseFloat(getCol(dr, 'uria'))    || 0,
        kmpl         : parseFloat(getCol(dr, 'kmpl'))    || 0,
        deadKm       : parseFloat(getCol(dr, 'deadKm'))  || 0,
        divKm        : parseFloat(getCol(dr, 'divKm'))   || 0,
        optKm        : parseFloat(getCol(dr, 'optKm'))   || 0
      };
      break;
    }

    return {
      vehicleNo    : vehStr,
      schIds       : schIds,
      tripsBySchId : tripsBySchId,
      totalOptKm   : parseFloat(totalOptKm.toFixed(2)),
      totalTrips   : totalTrips,
      yesTrips     : yesTrips,
      driverBadge  : driverBadge,
      driverName   : driverName,
      driverMobile : driverMobile,
      existing     : { diesel: existingDiesel }
    };

  } catch (e) {
    Logger.log('getVehicleTCData error: ' + e.message);
    return { error: e.message, schIds: [] };
  }
}


/**
 * getTCVehicleMap(date)
 *
 * Returns { schId: vehicleNo } for all TC rows on a given date.
 * Used by Fuel Entry vehicle tile selector — reads vehicle from
 * saved TC rows (col 10, saved by saveEntry TC section).
 */
function getTCVehicleMap(date) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName(SHEET_NAME);
    if (!sh) return {};
    detectColumns(sh);

    var rows = sh.getDataRange().getValues();
    var map  = {};
    var dStr = String(date).substring(0, 10);

    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (formatDateVal(rows[i][0]) !== dStr) continue;
      if (String(getCol(rows[i], 'mod')) !== 'TC') continue;

      var schId   = String(getCol(rows[i], 'schId')   || rows[i][2] || '').trim();
      var vehicle = String(getCol(rows[i], 'vehicle') || '').trim();

      if (schId && vehicle && !map[schId]) {
        map[schId] = vehicle;
      }
    }

    return map;

  } catch (e) {
    Logger.log('getTCVehicleMap error: ' + e.message);
    return {};
  }
}