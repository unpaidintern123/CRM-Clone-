/**
 * ============================================================
 *  DAILY LEAD-TO-CONTACT REPORT  (Google Apps Script)
 *  - Reads "Lead Database" sheet
 *  - Col A  = Lead Date (lead entered CRM)
 *  - Col K  = Assigned To (CRM person)
 *  - Col AP = First Contact Timestamp (when person contacted)
 *
 *  SLA RULES (working hours 11:00 - 20:00):
 *   - Lead arrives DURING working hours  -> contact within 30 min
 *   - Lead arrives OUTSIDE working hours  -> deadline = 12:00 PM next day
 *
 *  PDF -> saved to a specific Drive folder, made PUBLIC (anyone w/ link),
 *         verified public before the WhatsApp link is sent.
 *
 *  Run manually: generateAndSendLeadReport()
 * ============================================================
 */

// ---------------- CONFIG ----------------
const CONFIG = {
  SHEET_NAME: 'Leads Data',

  // Column indexes (1-based, matching spreadsheet letters)
  COL_LEAD_DATE: 1,      // A
  COL_ASSIGNED_TO: 11,   // K
  COL_FIRST_CONTACT: 42, // AP

  HEADER_ROW: 1,         // first data row is HEADER_ROW + 1

  // Working hours (24h)
  WORK_START_HOUR: 11,   // 11:00 AM
  WORK_END_HOUR: 20,     // 8:00 PM

  // SLA windows
  SLA_IN_HOURS_MIN: 30,         // within working hours -> 30 min
  SLA_AFTER_HOURS_DEADLINE: 12, // outside working hours -> by 12:00 PM next day

  // Report scope: 'yesterday' or 'today'
  REPORT_DAY: 'yesterday',

  // Drive
  DRIVE_FOLDER_ID: '1r1W_qsn5zjnJKgxXBvzrNWe4QKWib75l',

  // WhatsApp
  WA_MOBILE_NO: '9920975602', // <-- recipient number (with country code)
  WA_SESSION_ID: 'processC'
};

/**
 * MAIN — run this manually.
 */
function generateAndSendLeadReport() {
  const data = readSheetData_();
  const target = getReportDate_(CONFIG.REPORT_DAY);
  const stats = buildStats_(data, target);
  const pdfBlob = buildPdf_(stats, target);

  // Save to folder + make public + verify
  const publicUrl = savePublicPdf_(pdfBlob);

  // Send WhatsApp
  const message = buildWhatsAppMessage_(stats, target, publicUrl);
  sendReportOnWhatsApp_(CONFIG.WA_MOBILE_NO, message, CONFIG.WA_SESSION_ID);

  Logger.log('Report done. Public PDF: ' + publicUrl);
}

/* ---------------- DRIVE / PUBLIC LINK LAYER ---------------- */

/**
 * Saves the PDF into the configured folder, makes it public
 * (anyone with the link can view), verifies the access level,
 * and returns a public link.
 * Throws if the file cannot be confirmed public.
 */
function savePublicPdf_(pdfBlob) {
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const file = folder.createFile(pdfBlob);

  // Make public: anyone with the link can view
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log('setSharing failed, retrying with ANYONE: ' + e);
    file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  }

  // Verify the sharing access is actually public
  const access = file.getSharingAccess();
  const isPublic = (access === DriveApp.Access.ANYONE_WITH_LINK ||
                    access === DriveApp.Access.ANYONE);

  if (!isPublic) {
    throw new Error('PDF is NOT publicly accessible. Current access: ' + access +
                    '. Check folder/Drive sharing restrictions (e.g. Workspace admin policy).');
  }

  Logger.log('PDF verified public. Access: ' + access);

  // Direct-view link that works for anyone with the link
  return 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';
}

/* ---------------- DATA LAYER ---------------- */

function readSheetData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME);

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), CONFIG.COL_FIRST_CONTACT);
  if (lastRow <= CONFIG.HEADER_ROW) return [];

  const range = sheet.getRange(
    CONFIG.HEADER_ROW + 1, 1,
    lastRow - CONFIG.HEADER_ROW, lastCol
  );
  return range.getValues();
}

function getReportDate_(mode) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const d = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd') + ' 00:00:00');
  if (mode === 'yesterday') d.setDate(d.getDate() - 1);
  return d;
}

/**
 * Parse a cell that may be a Date object OR a "dd/MM/yyyy HH:mm:ss" string.
 */
function parseDateCell_(val) {
  if (val === '' || val === null || val === undefined) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  const s = String(val).trim();
  if (!s) return null;

  const m = s.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (m) {
    const day = +m[1], mon = +m[2] - 1, yr = +m[3];
    const hh = +(m[4] || 0), mm = +(m[5] || 0), ss = +(m[6] || 0);
    const d = new Date(yr, mon, day, hh, mm, ss);
    return isNaN(d.getTime()) ? null : d;
  }

  const fallback = new Date(s);
  return isNaN(fallback.getTime()) ? null : fallback;
}

function isSameDay_(a, b) {
  return a.getFullYear() === b.getFullYear() &&
         a.getMonth() === b.getMonth() &&
         a.getDate() === b.getDate();
}

/**
 * Compute the SLA deadline Date for a given lead-arrival Date.
 */
function computeDeadline_(leadDate) {
  const hour = leadDate.getHours();
  const startH = CONFIG.WORK_START_HOUR;
  const endH = CONFIG.WORK_END_HOUR;

  // Within working hours -> 30 minute window
  if (hour >= startH && hour < endH) {
    return {
      deadline: new Date(leadDate.getTime() + CONFIG.SLA_IN_HOURS_MIN * 60000),
      window: 'in_hours'
    };
  }

  // Before opening (same day) -> by 12:00 PM today
  if (hour < startH) {
    const d = new Date(leadDate.getFullYear(), leadDate.getMonth(),
                       leadDate.getDate(), CONFIG.SLA_AFTER_HOURS_DEADLINE, 0, 0);
    return { deadline: d, window: 'after_hours' };
  }

  // At/after closing (>= 20:00) -> by 12:00 PM next day
  const d = new Date(leadDate.getFullYear(), leadDate.getMonth(),
                     leadDate.getDate() + 1, CONFIG.SLA_AFTER_HOURS_DEADLINE, 0, 0);
  return { deadline: d, window: 'after_hours' };
}

/**
 * Build per-person statistics for the target day.
 */
function buildStats_(rows, targetDate) {
  const perPerson = {};

  rows.forEach(function (row) {
    const leadDate = parseDateCell_(row[CONFIG.COL_LEAD_DATE - 1]);
    if (!leadDate || !isSameDay_(leadDate, targetDate)) return;

    let name = String(row[CONFIG.COL_ASSIGNED_TO - 1] || '').trim();
    if (!name) name = 'Unassigned';

    const contactDate = parseDateCell_(row[CONFIG.COL_FIRST_CONTACT - 1]);
    const sla = computeDeadline_(leadDate);

    if (!perPerson[name]) {
      perPerson[name] = {
        assigned: 0, onTime: 0, delayed: 0, notContacted: 0,
        inHours: 0, afterHours: 0
      };
    }
    const p = perPerson[name];
    p.assigned++;
    if (sla.window === 'in_hours') p.inHours++; else p.afterHours++;

    if (!contactDate) {
      p.notContacted++;
    } else if (contactDate.getTime() <= sla.deadline.getTime() &&
               contactDate.getTime() >= leadDate.getTime()) {
      p.onTime++;
    } else {
      p.delayed++;
    }
  });

  const totals = {
    assigned: 0, onTime: 0, delayed: 0, notContacted: 0,
    inHours: 0, afterHours: 0
  };
  const people = Object.keys(perPerson).sort().map(function (name) {
    const p = perPerson[name];
    totals.assigned += p.assigned;
    totals.onTime += p.onTime;
    totals.delayed += p.delayed;
    totals.notContacted += p.notContacted;
    totals.inHours += p.inHours;
    totals.afterHours += p.afterHours;
    return {
      name: name,
      assigned: p.assigned,
      onTime: p.onTime,
      delayed: p.delayed,
      notContacted: p.notContacted,
      inHours: p.inHours,
      afterHours: p.afterHours,
      onTimePct: p.assigned ? Math.round((p.onTime / p.assigned) * 100) : 0
    };
  });

  return { people: people, totals: totals };
}

/* ---------------- PDF LAYER ---------------- */

function buildPdf_(stats, targetDate) {
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(targetDate, tz, 'dd MMM yyyy');
  const genStr = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy, HH:mm');
  const t = stats.totals;
  const overallPct = t.assigned ? Math.round((t.onTime / t.assigned) * 100) : 0;

  let rowsHtml = '';
  stats.people.forEach(function (p) {
    rowsHtml +=
      '<tr>' +
        '<td class="name">' + escapeHtml_(p.name) + '</td>' +
        '<td>' + p.assigned + '</td>' +
        '<td>' + p.inHours + '</td>' +
        '<td>' + p.afterHours + '</td>' +
        '<td class="ok">' + p.onTime + '</td>' +
        '<td class="bad">' + p.delayed + '</td>' +
        '<td class="warn">' + p.notContacted + '</td>' +
        '<td>' + p.onTimePct + '%</td>' +
      '</tr>';
  });
  if (!rowsHtml) {
    rowsHtml = '<tr><td colspan="8" style="text-align:center;padding:20px;">' +
               'No leads found for this day.</td></tr>';
  }

  const html =
    '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#222;margin:32px;}' +
    'h1{font-size:22px;margin:0 0 4px;}' +
    '.sub{color:#666;font-size:12px;margin-bottom:20px;}' +
    '.cards{display:flex;gap:12px;margin-bottom:24px;}' +
    '.card{flex:1;border:1px solid #e0e0e0;border-radius:8px;padding:12px;text-align:center;}' +
    '.card .num{font-size:24px;font-weight:bold;}' +
    '.card .lbl{font-size:11px;color:#666;text-transform:uppercase;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}' +
    'th,td{border:1px solid #e0e0e0;padding:7px 8px;text-align:center;}' +
    'th{background:#1a73e8;color:#fff;font-size:11px;}' +
    'td.name{text-align:left;font-weight:bold;}' +
    '.ok{color:#137333;font-weight:bold;}' +
    '.bad{color:#c5221f;font-weight:bold;}' +
    '.warn{color:#b06000;font-weight:bold;}' +
    'tr:nth-child(even){background:#f7f9fc;}' +
    '.footer{margin-top:24px;font-size:11px;color:#999;}' +
    '</style></head><body>' +

    '<h1>Lead-to-Contact Report</h1>' +
    '<div class="sub">Leads for <b>' + dateStr + '</b> &nbsp;|&nbsp; ' +
      'Working hours: ' + CONFIG.WORK_START_HOUR + ':00-' + CONFIG.WORK_END_HOUR + ':00 &nbsp;|&nbsp; ' +
      'In-hours SLA: ' + CONFIG.SLA_IN_HOURS_MIN + ' min &nbsp;|&nbsp; ' +
      'After-hours SLA: by ' + CONFIG.SLA_AFTER_HOURS_DEADLINE + ':00 next day &nbsp;|&nbsp; ' +
      'Generated: ' + genStr + '</div>' +

    '<div class="cards">' +
      '<div class="card"><div class="num">' + t.assigned + '</div><div class="lbl">Total Leads</div></div>' +
      '<div class="card"><div class="num ok">' + t.onTime + '</div><div class="lbl">On Time</div></div>' +
      '<div class="card"><div class="num bad">' + t.delayed + '</div><div class="lbl">Delayed</div></div>' +
      '<div class="card"><div class="num warn">' + t.notContacted + '</div><div class="lbl">Not Contacted</div></div>' +
      '<div class="card"><div class="num">' + overallPct + '%</div><div class="lbl">On-Time Rate</div></div>' +
    '</div>' +

    '<table>' +
      '<thead><tr>' +
        '<th>Assigned To</th><th>Leads Received</th>' +
        '<th>In-Hours</th><th>After-Hours</th>' +
        '<th>On Time</th><th>Delayed</th><th>Not Contacted</th><th>On-Time %</th>' +
      '</tr></thead>' +
      '<tbody>' + rowsHtml + '</tbody>' +
    '</table>' +

    '<div class="footer">Auto-generated from "' + CONFIG.SHEET_NAME + '". ' +
    'In-hours leads (11:00-20:00) must be contacted within ' + CONFIG.SLA_IN_HOURS_MIN +
    ' minutes. After-hours leads must be contacted by ' + CONFIG.SLA_AFTER_HOURS_DEADLINE +
    ':00 PM the next working day.</div>' +
    '</body></html>';

  const pdf = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf');
  pdf.setName('Lead_Report_' + Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd') + '.pdf');
  return pdf;
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/* ---------------- WHATSAPP LAYER ---------------- */

function buildWhatsAppMessage_(stats, targetDate, pdfUrl) {
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(targetDate, tz, 'dd MMM yyyy');
  const t = stats.totals;
  const overallPct = t.assigned ? Math.round((t.onTime / t.assigned) * 100) : 0;

  let msg = '*Lead-to-Contact Report*\n';
  msg += 'Date: ' + dateStr + '\n';
  msg += 'Working Hours: ' + CONFIG.WORK_START_HOUR + ':00-' + CONFIG.WORK_END_HOUR + ':00\n';
  msg += 'In-hours SLA: ' + CONFIG.SLA_IN_HOURS_MIN + ' min | After-hours: by ' +
         CONFIG.SLA_AFTER_HOURS_DEADLINE + ':00 next day\n\n';
  msg += 'Total Leads: ' + t.assigned + '\n';
  msg += 'On Time: ' + t.onTime + '\n';
  msg += 'Delayed: ' + t.delayed + '\n';
  msg += 'Not Contacted: ' + t.notContacted + '\n';
  msg += 'On-Time Rate: ' + overallPct + '%\n\n';

  stats.people.forEach(function (p) {
    msg += '• ' + p.name + ': ' + p.assigned + ' leads (' +
           p.inHours + ' in-hrs, ' + p.afterHours + ' after-hrs) | ' +
           p.onTime + ' on time, ' + p.delayed + ' delayed';
    if (p.notContacted) msg += ', ' + p.notContacted + ' not contacted';
    msg += '\n';
  });

  msg += '\nPDF: ' + pdfUrl;
  return msg;
}

function sendReportOnWhatsApp_(mobileNo, message, sessionID) {
  whatsappApi.sendWhatsAppMessage(mobileNo, message, sessionID);
}