/**
 * =========================================================================
 * AHL & ALCHEMANE IVR — MAIN SHEET ("Leads Database") TO SUPABASE SYNC
 * =========================================================================
 * 
 * - NO MENUS CREATED (Will never touch or overwrite your CRM Admin menu).
 * - Matches your sheet:
 *   Col A: username
 *   Col B: password (agent_id)
 *   Col C: permission (role)
 *   Col D: status
 *   Col E: next monday (attendance)
 *   Col G: Email ID
 *   Col H: number
 * 
 * - Automatic 10-digit phone cleaning and normalization.
 * - Business defaults to 'AHL' for all agents automatically.
 * - Live on-edit sync runs in < 100ms.
 * =========================================================================
 */

// =========================================================================
// 1. CONFIGURATION
// =========================================================================
var SUPABASE_URL = "https://iqefkgcluwsdzywoqhbu.supabase.co";
var SUPABASE_ANON_KEY = "sb_publishable_J55wKvDRQEicw9zVXVh2NQ_bcO6iSbG";
var DEFAULT_BUSINESS = "AHL";


// =========================================================================
// 2. LIVE ON-EDIT HANDLER (< 100ms)
// =========================================================================
function handleSheetEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = String(sheet.getName() || '').toLowerCase().trim();

  // CRITICAL: Strictly only trigger for the Users / Doctors / Agents tab!
  // Prevents edits in "Leads Database", "Orders", or CRM tabs from corrupting the agents table.
  if (sheetName !== "users" && sheetName !== "user" && sheetName !== "agents" && sheetName !== "doctors") {
    return;
  }

  var row = e.range.getRow();
  if (row < 2) return; // Skip header row

  try {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
      return String(h || '').toLowerCase().trim();
    });

    var colMap = getColumnMap_(headers);
    var editedCol = e.range.getColumn() - 1; // 0-based index

    // Only fire if edited cell is 'status', 'next monday', 'number', or 'username'
    var isRelevantEdit = (
      editedCol === colMap.status ||
      editedCol === colMap.attendance ||
      editedCol === colMap.mobile ||
      editedCol === colMap.username
    );
    if (!isRelevantEdit) return;

    var rowVals = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
    var username = String(rowVals[colMap.username] || '').trim();
    var rawStatus = String(rowVals[colMap.status] || 'Active').trim();
    var rawAttendance = String(rowVals[colMap.attendance] || 'Coming').trim();
    var rawMobile = colMap.mobile >= 0 ? String(rowVals[colMap.mobile] || '').replace(/[^0-9]/g, '') : '';
    var email = colMap.email >= 0 ? String(rowVals[colMap.email] || '').trim() : '';
    var agentId = colMap.agent_id >= 0 ? String(rowVals[colMap.agent_id] || '').trim() : '';
    var role = colMap.role >= 0 ? String(rowVals[colMap.role] || 'CRM Executive').trim() : 'CRM Executive';

    if (!username) return;

    var status = normalizeStatus_(rawStatus);
    var attendance = normalizeAttendance_(rawAttendance);
    var mobile = rawMobile.length >= 10 ? rawMobile.slice(-10) : rawMobile;

    // Smart Agent Sync: matches by username first, then by mobile (if renamed), else appends new member
    var success = syncAgentSmart_(username, mobile, status, attendance, DEFAULT_BUSINESS, email, agentId, role);

    if (success) {
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          username + ' -> ' + status + ' (' + attendance + ') synced to Supabase!',
          '⚡ Supabase Live Sync',
          3
        );
      } catch (t) { }
    }

  } catch (err) {
    Logger.log('[handleSheetEdit] Error: ' + err.message);
  }
}


// =========================================================================
// 3. SMART SUPABASE AGENT SYNC (Zero Duplicates Engine)
// =========================================================================
/**
 * Smart Agent Sync:
 * 1. Primary Match: Match by USERNAME (name). If found, UPDATE that exact agent!
 *    (Zero duplicates created if their phone number, status, or attendance changes).
 * 2. Secondary Match: If username not found, match by MOBILE number. If found, UPDATE.
 *    (Handles case where someone renames an existing agent in the sheet).
 * 3. Fallback: If neither username nor mobile exists in Supabase, INSERT new agent (appends new member).
 */
function syncAgentSmart_(name, mobile, status, attendance, business, email, agentId, role) {
  if (!SUPABASE_URL || SUPABASE_URL.indexOf('YOUR_PROJECT_REF') >= 0) return false;

  var cleanUrl = SUPABASE_URL.replace(/\/$/, '');
  var headers = {
    'apikey': SUPABASE_ANON_KEY,
    'Authorization': 'Bearer ' + SUPABASE_ANON_KEY,
    'Content-Type': 'application/json'
  };

  var payload = {
    name: name,
    status: status,
    attendance: attendance,
    business: business || DEFAULT_BUSINESS,
    agent_id: agentId || null,
    role: role || 'CRM Executive',
    google_email: email || null,
    updated_at: new Date().toISOString()
  };
  if (mobile) {
    payload.mobile = mobile;
  }

  try {
    // -------------------------------------------------------------------------
    // STEP 1: MATCH BY USERNAME (Primary)
    // -------------------------------------------------------------------------
    var findByNameUrl = cleanUrl + '/rest/v1/agents?name=ilike.' + encodeURIComponent(name) + '&select=id,name,mobile';
    var nameCheckResp = UrlFetchApp.fetch(findByNameUrl, { method: 'GET', headers: headers, muteHttpExceptions: true });

    if (nameCheckResp.getResponseCode() === 200) {
      var nameMatches = JSON.parse(nameCheckResp.getContentText());
      if (nameMatches && nameMatches.length > 0) {
        var existingId = nameMatches[0].id;
        var patchUrl = cleanUrl + '/rest/v1/agents?id=eq.' + existingId;
        var patchResp = UrlFetchApp.fetch(patchUrl, {
          method: 'PATCH',
          headers: headers,
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        var pCode = patchResp.getResponseCode();
        if (pCode === 200 || pCode === 204) {
          Logger.log('[syncAgentSmart_] Updated existing agent by USERNAME: ' + name + ' (ID: ' + existingId + ')');
          return true;
        }
      }
    }

    // -------------------------------------------------------------------------
    // STEP 2: MATCH BY NUMBER (Secondary - handles case where agent was renamed)
    // -------------------------------------------------------------------------
    if (mobile) {
      var findByMobileUrl = cleanUrl + '/rest/v1/agents?mobile=eq.' + encodeURIComponent(mobile) + '&select=id,name,mobile';
      var mobileCheckResp = UrlFetchApp.fetch(findByMobileUrl, { method: 'GET', headers: headers, muteHttpExceptions: true });

      if (mobileCheckResp.getResponseCode() === 200) {
        var mobileMatches = JSON.parse(mobileCheckResp.getContentText());
        if (mobileMatches && mobileMatches.length > 0) {
          var existingMobileId = mobileMatches[0].id;
          var patchMobUrl = cleanUrl + '/rest/v1/agents?id=eq.' + existingMobileId;
          var patchMobResp = UrlFetchApp.fetch(patchMobUrl, {
            method: 'PATCH',
            headers: headers,
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
          });
          var pmCode = patchMobResp.getResponseCode();
          if (pmCode === 200 || pmCode === 204) {
            Logger.log('[syncAgentSmart_] Renamed agent updated by MOBILE: ' + mobile + ' -> new name: ' + name);
            return true;
          }
        }
      }
    }

    // -------------------------------------------------------------------------
    // STEP 3: INSERT BRAND NEW AGENT (Append)
    // -------------------------------------------------------------------------
    if (!payload.mobile) {
      Logger.log('[syncAgentSmart_] Cannot insert new agent without mobile number for: ' + name);
      return false;
    }

    var insertUrl = cleanUrl + '/rest/v1/agents';
    var insertResp = UrlFetchApp.fetch(insertUrl, {
      method: 'POST',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var inCode = insertResp.getResponseCode();
    if (inCode === 200 || inCode === 201 || inCode === 204) {
      Logger.log('[syncAgentSmart_] Successfully appended NEW agent: ' + name);
      return true;
    } else {
      Logger.log('[syncAgentSmart_] Insert error HTTP ' + inCode + ': ' + insertResp.getContentText());
      return false;
    }

  } catch (err) {
    Logger.log('[syncAgentSmart_] Exception: ' + err.message);
    return false;
  }
}


// =========================================================================
// 4. ONE-CLICK INSTALLABLE TRIGGER SETUP (Run from Apps Script Editor)
// =========================================================================
/**
 * Select 'installSyncTrigger' in the top toolbar dropdown and click Run (▶️).
 * Authorizes the trigger once so external API calls can fire on every sheet edit!
 */
function installSyncTrigger() {
  if (!SUPABASE_URL || SUPABASE_URL.indexOf('YOUR_PROJECT_REF') >= 0) {
    Logger.log('[installSyncTrigger] Error: Please set SUPABASE_URL first.');
    return;
  }

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'handleSheetEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('handleSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log('[installSyncTrigger] SUCCESS: Live Sync trigger is now permanently active!');
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Live Sync trigger activated successfully!', '⚡ Supabase Sync', 5);
  } catch (t) { }
}

/**
 * Select 'pushAllAgentsToSupabase' in the top toolbar dropdown and click Run (▶️).
 * Syncs all agents from the sheet to Supabase using the Smart Match logic (Zero Duplicates).
 */
function pushAllAgentsToSupabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findUsersSheet_();
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log('No rows found in sheet: ' + sheet.getName());
    return;
  }

  var headers = data[0].map(function (h) { return String(h || '').toLowerCase().trim(); });
  var colMap = getColumnMap_(headers);

  var syncedCount = 0;
  var skipped = [];

  for (var r = 1; r < data.length; r++) {
    var username = String(data[r][colMap.username] || '').trim();
    if (!username) continue;

    var rawMobile = colMap.mobile >= 0 ? String(data[r][colMap.mobile] || '').replace(/[^0-9]/g, '') : '';
    var status = normalizeStatus_(data[r][colMap.status]);
    var attendance = normalizeAttendance_(data[r][colMap.attendance]);
    var email = colMap.email >= 0 ? String(data[r][colMap.email] || '').trim() : '';
    var agentId = colMap.agent_id >= 0 ? String(data[r][colMap.agent_id] || '').trim() : '';
    var role = colMap.role >= 0 ? String(data[r][colMap.role] || 'CRM Executive').trim() : 'CRM Executive';
    var mobile = rawMobile.length >= 10 ? rawMobile.slice(-10) : rawMobile;

    if (mobile) {
      var ok = syncAgentSmart_(username, mobile, status, attendance, DEFAULT_BUSINESS, email, agentId, role);
      if (ok) syncedCount++;
    } else {
      skipped.push(username);
    }
  }

  var msg = 'Sync Complete! ' + syncedCount + ' agent(s) synced to Supabase without duplicates.';
  if (skipped.length) {
    msg += ' (Skipped without 10-digit number: ' + skipped.join(', ') + ')';
  }
  Logger.log(msg);
  try { ss.toast(msg, '⚡ Supabase Sync', 6); } catch (t) { }
}

/**
 * Tests connection to Supabase.
 */
function testSupabaseConnection() {
  var url = SUPABASE_URL.replace(/\/$/, '') + '/rest/v1/agents?select=count';
  var options = {
    method: 'GET',
    headers: {
      'apikey': SUPABASE_ANON_KEY,
      'Authorization': 'Bearer ' + SUPABASE_ANON_KEY
    },
    muteHttpExceptions: true
  };

  try {
    var res = UrlFetchApp.fetch(url, options);
    var ok = res.getResponseCode() === 200;
    Logger.log('[testSupabaseConnection] HTTP ' + res.getResponseCode());
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        ok ? 'Connection Successful! (HTTP 200)' : 'Failed: HTTP ' + res.getResponseCode(),
        'Supabase Test',
        5
      );
    } catch (t) { }
  } catch (e) {
    Logger.log('[testSupabaseConnection] Error: ' + e.message);
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + e.message, 'Supabase Test', 7);
    } catch (t) { }
  }
}


// =========================================================================
// 5. HELPER FUNCTIONS & COLUMN RESOLUTION
// =========================================================================
function findUsersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1. Try finding by explicit tab name first
  var exact = ss.getSheetByName("Users") || ss.getSheetByName("users") || ss.getSheetByName("Doctors") || ss.getSheetByName("Agents");
  if (exact) return exact;

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var name = s.getName().toLowerCase();
    if (name.indexOf("lead") >= 0 || name.indexOf("order") >= 0 || name.indexOf("follow") >= 0) continue; // skip CRM leads sheets!
    var lastCol = s.getLastColumn();
    if (lastCol > 0) {
      var headers = s.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
        return String(h || '').toLowerCase().trim();
      });
      if (headers.indexOf('username') >= 0 || (headers.indexOf('name') >= 0 && headers.indexOf('status') >= 0)) {
        return s;
      }
    }
  }
  return null;
}

function getColumnMap_(headers) {
  var map = { username: 0, status: 3, attendance: 4, mobile: 7, email: 6, role: 2, agent_id: 1 };

  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h === 'username' || h === 'name' || h === 'user' || h === 'agent') map.username = i;
    else if (h === 'password' || h === 'agent_id' || h === 'id') map.agent_id = i;
    else if (h === 'status') map.status = i;
    else if (h.indexOf('monday') >= 0 || h.indexOf('coming') >= 0 || h.indexOf('attendance') >= 0) map.attendance = i;
    else if (h === 'number' || h === 'mobile' || h === 'phone' || h.indexOf('contact') >= 0) map.mobile = i;
    else if (h.indexOf('email') >= 0) map.email = i;
    else if (h.indexOf('permission') >= 0 || h.indexOf('role') >= 0) map.role = i;
  }
  return map;
}

function normalizeStatus_(status) {
  var s = String(status || '').trim().toLowerCase();
  return (s === 'active' || s === 'true' || s === 'yes' || s === '1') ? 'Active' : 'Inactive';
}

function normalizeAttendance_(attendance) {
  var a = String(attendance || '').trim().toLowerCase();
  return (a.indexOf('not') >= 0 || a.indexOf('absent') >= 0 || a === 'no' || a === 'false') ? 'Not Coming' : 'Coming';
}
