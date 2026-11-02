function doGet(e) {
  let temp = "Login"
  if ('page' in e.parameters) {
    temp = e.parameters['page'][0];
  }
  try {
    const html = HtmlService.createTemplateFromFile(temp);
    const props = PropertiesService.getScriptProperties();
    html.CRM_USE_API_PROXY = props.getProperty("CRM_API_BASE_URL") ? "true" : "";
    html.CRM_WEB_APP_BASE = (ScriptApp.getService().getUrl() || "").replace(/\/dev$/, "/exec");
    html.data = {
      // userData : retriveLeadData(),
      e: e
    }
    return html.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  } catch (e) {
    return HtmlService.createHtmlOutput("<h1>Error : </h1>" + e)
  }
}

function setApiConfig(baseUrl, token) {
  if (!baseUrl || !token) {
    throw new Error("Both baseUrl and token are required.");
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty("CRM_API_BASE_URL", String(baseUrl).trim());
  props.setProperty("CRM_API_TOKEN", String(token).trim());
  return "API config saved in Script Properties.";
}

/**
 * Custom menu at the top of the Google Sheet (next to Help).
 * Reload the sheet after clasp push if the menu does not appear.
 */
function onOpen() {
  ensureLeadsDataHeaders();
  SpreadsheetApp.getUi()
    .createMenu("CRM Admin")
    .addItem("➕ Add New Lead", "menuAddNewLeadDialog")
    .addSeparator()
    .addItem("⚡ Sync Selected Row(s) to Firebase", "menuSyncSelectedRowsToFirebase")
    .addSeparator()
    .addItem("Run Consultation Visited Sync", "menuRunConsultationVisitedSync")
    .addToUi();
}

/**
 * Opens the modern Add New Lead dialog modal in Google Sheets.
 */
function menuAddNewLeadDialog() {
  const html = HtmlService.createTemplateFromFile("addNewLeadDialog")
    .evaluate()
    .setWidth(540)
    .setHeight(660);
  SpreadsheetApp.getUi().showModalDialog(html, "➕ Add New Lead");
}

/**
 * Retrieves active users from the Users sheet tab for dropdown selection.
 */
function getActiveUsersList() {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) {
      sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Users");
    }
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const allUsers = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const username = String(row[0] || "").trim();
      if (username && !allUsers.includes(username)) {
        allUsers.push(username);
      }
    }
    return allUsers;
  } catch (err) {
    Logger.log("getActiveUsersList error: " + err);
    return [];
  }
}

/**
 * Handles adding a new lead from the Google Sheets UI dialog directly into Firebase Firestore.
 * Firebase automatically assigns Lead ID, handles round-robin if unassigned, and appends the row in Google Sheets.
 */
function addNewLeadFromUi(payload) {
  if (!payload || !payload.clientName || !payload.mobileNo) {
    throw new Error("Client Name and Mobile Number are required.");
  }

  let actorEmail = "sheet_manual_entry";
  try {
    const user = Session.getActiveUser().getEmail();
    if (user) actorEmail = user;
  } catch (e) {
    // ignore
  }

  const cleanPayload = {
    clientName: String(payload.clientName).trim(),
    mobileNo: String(payload.mobileNo).trim(),
    city: payload.city ? String(payload.city).trim() : "",
    platform: payload.platform ? String(payload.platform).trim() : "",
    callType: payload.callType ? String(payload.callType).trim() : "Outgoing",
    leadStatus: payload.leadStatus ? String(payload.leadStatus).trim() : "New Lead",
    assignedTo: payload.assignedTo ? String(payload.assignedTo).trim() : "",
    consutlationBookedDate: payload.consutlationBookedDate || "",
    consutlationVisitedDate: payload.consutlationVisitedDate || "",
    followUpDate: payload.followUpDate || "",
    consultationType: payload.consultationType || "",
    initialRemark: payload.initialRemark || payload.remark || "",
    detailComment: payload.detailComment || payload.initialRemark || payload.remark || "",
    actor: actorEmail
  };

  const response = crmApiProxy("/gas/newLeadUpload", "POST", JSON.stringify(cleanPayload));
  writeTrackingFields_(response, cleanPayload.callType === "Incoming" ? { au: "Phone", aw: "Incoming Call" } : { aw: "Manual Entry" });
  return response;
}

/**
 * Automatically ensures 'Email' and 'Query' headers exist in Row 1 of Leads Data sheet.
 */
function ensureLeadsDataHeaders() {
  try {
    const sheet = getLeadsDataSheet();
    if (!sheet) return;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, Math.max(lastCol, 45)).getValues()[0] || [];
    const headerLower = headers.map(h => String(h || "").trim().toLowerCase());

    let emailIdx = headerLower.indexOf("email");
    let queryIdx = headerLower.indexOf("query");

    if (emailIdx === -1) {
      const emailCol = (lastCol >= 43 ? 44 : lastCol + 1); // Col AR (44)
      sheet.getRange(1, emailCol).setValue("Email");
      emailIdx = emailCol - 1;
    }

    if (queryIdx === -1) {
      const queryCol = (emailIdx !== -1 ? emailIdx + 2 : sheet.getLastColumn() + 1); // Col AS (45)
      sheet.getRange(1, queryCol).setValue("Query");
    }
  } catch (err) {
    Logger.log("ensureLeadsDataHeaders note: " + err);
  }
}

/**
 * Syncs the currently selected row(s) from the active sheet to Firebase Firestore.
 * Can be triggered via the "🔥 Update FB" menu or assigned to a clickable Drawing/Button in Google Sheets.
 */
function menuSyncSelectedRowsToFirebase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const range = sheet.getActiveRange();
  if (!range) {
    ui.alert("⚠️ Please select one or more rows to update in Firebase.");
    return;
  }

  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const endRow = startRow + numRows - 1;

  if (endRow < 2) {
    ui.alert("⚠️ Header row selected. Please select lead data rows (Row 2 onwards).");
    return;
  }

  const actualStartRow = Math.max(2, startRow);
  const totalToProcess = endRow - actualStartRow + 1;

  const headerMap = getLeadsHeaderMap(sheet);
  const lastCol = sheet.getLastColumn();
  const fullDataRange = sheet.getRange(actualStartRow, 1, totalToProcess, lastCol).getValues();

  let successCount = 0;
  const updatedDetails = [];
  const errors = [];

  function formatVal(val) {
    if (val === null || val === undefined) return "";
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    }
    return String(val).trim();
  }

  function getColVal(rowVals, colName) {
    const colIdx = headerMap[colName.toLowerCase()];
    if (!colIdx || colIdx > rowVals.length) return "";
    return formatVal(rowVals[colIdx - 1]);
  }

  for (let i = 0; i < fullDataRange.length; i++) {
    const rowVals = fullDataRange[i];
    const currentRowNum = actualStartRow + i;
    const leadUniqueId = getColVal(rowVals, "lead unique id") || (rowVals[1] ? formatVal(rowVals[1]) : "");

    if (!leadUniqueId) {
      Logger.log("Row " + currentRowNum + " skipped: No Lead Unique ID found.");
      continue;
    }

    const assignedToVal = getColVal(rowVals, "assigned to") || (rowVals[10] ? formatVal(rowVals[10]) : "");
    const clientNameVal = getColVal(rowVals, "client name") || (rowVals[2] ? formatVal(rowVals[2]) : "");

    const payload = {
      leadUniqueId: leadUniqueId,
      leadDate: getColVal(rowVals, "lead date"),
      clientName: clientNameVal,
      mobileNo: getColVal(rowVals, "mobile no") || getColVal(rowVals, "mobile") || getColVal(rowVals, "phone"),
      city: getColVal(rowVals, "city"),
      town: getColVal(rowVals, "town"),
      callType: getColVal(rowVals, "call type"),
      platform: getColVal(rowVals, "platform"),
      leadStatus: getColVal(rowVals, "lead status"),
      rating: getColVal(rowVals, "rating"),
      assignedTo: assignedToVal,
      interestedIn: getColVal(rowVals, "interested in"),
      followUpDate: getColVal(rowVals, "follow up date"),
      noOfFollowups: getColVal(rowVals, "no of follow ups") || getColVal(rowVals, "no of followups"),
      consutlationBookedDate: getColVal(rowVals, "consultation booked date"),
      consutlationVisitedDate: getColVal(rowVals, "consultation visited date"),
      consultationType: getColVal(rowVals, "consultation type"),
      consultationStatus: getColVal(rowVals, "consultation status"),
      initialRemark: getColVal(rowVals, "initial remark"),
      detailComment: getColVal(rowVals, "detail comment"),
      Q1: getColVal(rowVals, "how soon are you looking for solution?"),
      Q2: getColVal(rowVals, "do you have a budget in mind?"),
      Q3: getColVal(rowVals, "willing to visit to the center?"),
      hearAboutUs: getColVal(rowVals, "how did you hear about us?"),
      orderStatus: getColVal(rowVals, "order status"),
      totalSalesAmount: getColVal(rowVals, "total sales amount"),
      notBookedFeedback: getColVal(rowVals, "not booked feedback"),
      orderDate: getColVal(rowVals, "order date"),
      actor: "sheet_selection_update_fb"
    };

    try {
      crmApiProxy("/gas/updateLeadDetails", "POST", JSON.stringify(payload));
      successCount++;
      updatedDetails.push(leadUniqueId + " ➔ " + (assignedToVal || "Unassigned"));
    } catch (err) {
      errors.push("Row " + currentRowNum + " (" + leadUniqueId + "): " + err.message);
    }
  }

  if (successCount > 0) {
    ss.toast(
      "✅ " + successCount + " lead(s) updated in Firebase!\n" + updatedDetails.slice(0, 5).join(", "),
      "CRM Admin",
      6
    );
  }

  if (errors.length > 0) {
    ui.alert("⚠️ Some rows could not be updated in Firebase:\n\n" + errors.join("\n"));
  } else if (successCount === 0) {
    ui.alert("No valid leads found in the selected rows to update.");
  }
}

/**
 * Helper to get the Leads Data sheet object reliably.
 */
function getLeadsDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss ? ss.getSheetByName("Leads Data") : null;
  if (!sheet) {
    try {
      sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
    } catch (e) {
      Logger.log("Could not open spreadsheet by ID: " + e);
    }
  }
  return sheet;
}

/**
 * Helper to get column index map for header names in Leads Data sheet.
 */
function getLeadsHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  const map = {};
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] || "").trim().toLowerCase();
    if (h) {
      map[h] = c + 1; // 1-indexed column number
    }
  }
  return map;
}

/**
 * Returns an array of lead summaries [{ id, label, name, mobile, visitedDate }] for the UI datalist / dropdown.
 */
function getAllLeadSummariesForDropdown() {
  const sheet = getLeadsDataSheet();
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return [];

  const headerMap = getLeadsHeaderMap(sheet);
  const idCol = headerMap["lead unique id"] || 2;
  const nameCol = headerMap["client name"] || 3;
  const mobileCol = headerMap["mobile no"] || 4;
  const statusCol = headerMap["lead status"] || 9;
  const consStatusCol = headerMap["consultation status"] || 36;
  const visitedDateCol = headerMap["consultation visited date"] || 16;

  const results = [];
  // Include leads that have Consultation Visited status
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    const id = String(row[idCol - 1] || "").trim();
    if (!id) continue;

    const status = String(row[statusCol - 1] || "").trim().toLowerCase();
    const consStatus = String(row[consStatusCol - 1] || "").trim().toLowerCase();
    const isVisited = status === "consultation visited" || status === "visited" ||
      consStatus === "consultation visited" || consStatus === "visited" ||
      status.includes("visited") || consStatus.includes("visited");

    if (!isVisited) continue;

    const name = String(row[nameCol - 1] || "").trim();
    const mobile = String(row[mobileCol - 1] || "").trim();
    const visitedDate = row[visitedDateCol - 1];

    let visitedFormatted = "";
    if (visitedDate) {
      if (visitedDate instanceof Date) {
        visitedFormatted = Utilities.formatDate(visitedDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      } else {
        visitedFormatted = safeFormatDate(String(visitedDate), "dd/MM/yyyy HH:mm:ss");
      }
    }

    results.push({
      id: id,
      name: name,
      mobile: mobile,
      visitedDate: visitedFormatted,
      label: id + (name ? " — " + name : "") + (mobile ? " (" + mobile + ")" : "")
    });
  }
  return results;
}

/**
 * Fetches current lead information and date fields for the Update Lead Date dialog.
 * @param {string} leadIdInput - e.g. "Lead ID - 24966" or "24966"
 */
function getLeadDateInfo(leadIdInput) {
  const cleanInput = String(leadIdInput || "").trim();
  if (!cleanInput) {
    return { found: false, message: "Please provide a valid Lead ID." };
  }

  const sheet = getLeadsDataSheet();
  if (!sheet) {
    return { found: false, message: "Could not locate 'Leads Data' sheet." };
  }

  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return { found: false, message: "Leads Data sheet is empty." };
  }

  const headerMap = getLeadsHeaderMap(sheet);
  const idCol = headerMap["lead unique id"] || 2;
  const nameCol = headerMap["client name"] || 3;
  const mobileCol = headerMap["mobile no"] || 4;
  const cityCol = headerMap["city"] || 5;
  const statusCol = headerMap["lead status"] || 9;
  const assignedCol = headerMap["assigned to"] || 11;
  const consStatusCol = headerMap["consultation status"] || 36;
  const folDateCol = headerMap["follow up date"] || 13;
  const consBookedDateCol = headerMap["consultation booked date"] || 15;
  const consVisitedDateCol = headerMap["consultation visited date"] || 16;
  const leadDateCol = headerMap["lead date"] || 1;
  const orderDateCol = headerMap["order date"] || 38;

  const targetLower = cleanInput.toLowerCase();
  const digitsOnly = cleanInput.replace(/\D/g, "");

  let foundRow = null;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const rowId = String(row[idCol - 1] || "").trim();
    const rowIdLower = rowId.toLowerCase();
    const rowDigits = rowId.replace(/\D/g, "");

    if (
      (rowIdLower && rowIdLower === targetLower) ||
      (digitsOnly && rowDigits && rowDigits === digitsOnly) ||
      (rowIdLower.includes(targetLower))
    ) {
      foundRow = row;
      break;
    }
  }

  if (!foundRow) {
    return { found: false, message: "Lead ID '" + cleanInput + "' not found in Leads Data sheet." };
  }

  function formatDisplayDate(val) {
    if (!val) return "";
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    }
    return safeFormatDate(String(val), "dd/MM/yyyy HH:mm:ss");
  }

  function formatRawIso(val) {
    if (!val) return "";
    let d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) {
      const m = String(val).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{1,2}))?/);
      if (m) {
        d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0));
      }
    }
    if (isNaN(d.getTime())) return "";
    const pad = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }

  return {
    found: true,
    leadUniqueId: String(foundRow[idCol - 1] || "").trim(),
    clientName: String(foundRow[nameCol - 1] || "").trim(),
    mobileNo: String(foundRow[mobileCol - 1] || "").trim(),
    city: String(foundRow[cityCol - 1] || "").trim(),
    assignedTo: String(foundRow[assignedCol - 1] || "").trim(),
    leadStatus: String(foundRow[statusCol - 1] || "").trim(),
    consultationStatus: String(foundRow[consStatusCol - 1] || "").trim(),
    dates: {
      consutlationBookedDate: formatDisplayDate(foundRow[consBookedDateCol - 1]),
      followUpDate: formatDisplayDate(foundRow[folDateCol - 1]),
      consutlationVisitedDate: formatDisplayDate(foundRow[consVisitedDateCol - 1]),
      leadDate: formatDisplayDate(foundRow[leadDateCol - 1]),
      orderDate: formatDisplayDate(foundRow[orderDateCol - 1])
    },
    datesRaw: {
      consutlationBookedDate: formatRawIso(foundRow[consBookedDateCol - 1]),
      followUpDate: formatRawIso(foundRow[folDateCol - 1]),
      consutlationVisitedDate: formatRawIso(foundRow[consVisitedDateCol - 1]),
      leadDate: formatRawIso(foundRow[leadDateCol - 1]),
      orderDate: formatRawIso(foundRow[orderDateCol - 1])
    }
  };
}

/**
 * Converts datetime string to ISO 8601 string for Firestore.
 */
function convertToIsoStringForFirebase(value) {
  if (!value) return "";
  const str = String(value).trim();
  if (!str) return "";

  if (str.endsWith("Z") || (str.includes("T") && (str.includes("+") || str.lastIndexOf("-") > 10))) {
    return str;
  }

  if (str.includes("T")) {
    const parts = str.split("T");
    const [y, m, d] = parts[0].split("-");
    const timeParts = parts[1].split(":");
    const hr = String(timeParts[0] || "00").padStart(2, "0");
    const min = String(timeParts[1] || "00").padStart(2, "0");
    const sec = String(timeParts[2] || "00").padStart(2, "0");
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}T${hr}:${min}:${sec}+05:30`;
  }

  if (str.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return `${str}T00:00:00+05:30`;
  }

  const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (m) {
    const d = String(m[1]).padStart(2, "0");
    const mo = String(m[2]).padStart(2, "0");
    const y = m[3];
    const hr = String(m[4] || "00").padStart(2, "0");
    const min = String(m[5] || "00").padStart(2, "0");
    const sec = String(m[6] || "00").padStart(2, "0");
    return `${y}-${mo}-${d}T${hr}:${min}:${sec}+05:30`;
  }

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return parsed.toISOString();
  }
  return str;
}

/**
 * Updates a specific date field for a lead in BOTH Google Sheet and Firebase Firestore.
 * @param {Object} payload - { leadUniqueId: string, dateField: string, newDateValue: string }
 */
function updateLeadDateInFirebaseAndSheet(payload) {
  if (!payload || !payload.leadUniqueId || !payload.dateField) {
    throw new Error("Missing required parameters: leadUniqueId and dateField.");
  }

  let leadUniqueId = String(payload.leadUniqueId).trim();
  const dateField = String(payload.dateField).trim();
  const newDateValue = payload.newDateValue ? String(payload.newDateValue).trim() : "";

  const sheet = getLeadsDataSheet();
  if (!sheet) {
    throw new Error("Could not access 'Leads Data' sheet.");
  }

  const data = sheet.getDataRange().getValues();
  const headerMap = getLeadsHeaderMap(sheet);
  const idCol = headerMap["lead unique id"] || 2;

  const fieldColMap = {
    consutlationBookedDate: headerMap["consultation booked date"] || 15,
    followUpDate: headerMap["follow up date"] || 13,
    consutlationVisitedDate: headerMap["consultation visited date"] || 16,
    leadDate: headerMap["lead date"] || 1,
    orderDate: headerMap["order date"] || 38
  };

  const targetCol = fieldColMap[dateField];
  if (!targetCol) {
    throw new Error("Unknown date field: " + dateField);
  }

  let targetRowIndex = -1;
  const targetLower = leadUniqueId.toLowerCase();
  const digitsOnly = leadUniqueId.replace(/\D/g, "");
  for (let r = 1; r < data.length; r++) {
    const rawId = String(data[r][idCol - 1] || "").trim();
    const rowIdLower = rawId.toLowerCase();
    const rowDigits = rawId.replace(/\D/g, "");
    if (
      (rowIdLower && rowIdLower === targetLower) ||
      (digitsOnly && rowDigits && rowDigits === digitsOnly) ||
      (targetLower && rowIdLower.includes(targetLower))
    ) {
      targetRowIndex = r + 1; // 1-indexed row number
      leadUniqueId = rawId; // Normalize to full Lead ID string
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error("Lead '" + leadUniqueId + "' not found in Leads Data sheet.");
  }

  // Read existing cell and row data to preserve time
  const currentRowData = data[targetRowIndex - 1];
  const existingCellVal = currentRowData[targetCol - 1];

  let formattedSheetVal = "";
  let isoDate = "";

  if (newDateValue) {
    // Helper to extract non-zero time string (HH:mm:ss) from Date or string ONLY if time was present
    function extractExistingTime(val) {
      if (!val) return "";
      if (val instanceof Date && !isNaN(val.getTime())) {
        const t = Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm:ss");
        return t !== "00:00:00" ? t : "";
      }
      const str = String(val).trim();
      if (str.includes("T")) {
        const timePart = str.split("T")[1];
        const match = timePart.match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
        if (match) {
          const h = String(match[1]).padStart(2, "0");
          const m = String(match[2]).padStart(2, "0");
          const s = String(match[3] || "00").padStart(2, "0");
          if (h !== "00" || m !== "00" || s !== "00") return `${h}:${m}:${s}`;
        }
      }
      if (str.includes(" ")) {
        const timePart = str.split(" ")[1];
        const match = timePart.match(/(\d{1,2})[:.](\d{1,2})(?:[:.](\d{1,2}))?/);
        if (match) {
          const h = String(match[1]).padStart(2, "0");
          const m = String(match[2]).padStart(2, "0");
          const s = String(match[3] || "00").padStart(2, "0");
          if (h !== "00" || m !== "00" || s !== "00") return `${h}:${m}:${s}`;
        }
      }
      return "";
    }

    // Only preserve time if the existing target cell explicitly had one
    const preservedTime = extractExistingTime(existingCellVal);

    // Parse the new date component
    let y = "", mo = "", d = "";
    const cleanDateStr = String(newDateValue).trim();
    if (cleanDateStr.match(/^\d{4}-\d{2}-\d{2}/)) {
      const datePart = cleanDateStr.split("T")[0].split(" ")[0];
      [y, mo, d] = datePart.split("-");
    } else if (cleanDateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)) {
      const m = cleanDateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      d = String(m[1]).padStart(2, "0");
      mo = String(m[2]).padStart(2, "0");
      y = m[3];
    } else {
      const parsedD = new Date(cleanDateStr);
      if (!isNaN(parsedD.getTime())) {
        d = Utilities.formatDate(parsedD, Session.getScriptTimeZone(), "dd");
        mo = Utilities.formatDate(parsedD, Session.getScriptTimeZone(), "MM");
        y = Utilities.formatDate(parsedD, Session.getScriptTimeZone(), "yyyy");
      }
    }

    if (y && mo && d) {
      if (preservedTime) {
        // Had time -> preserve time
        formattedSheetVal = `${d}/${mo}/${y} ${preservedTime}`;
        isoDate = `${y}-${mo}-${d}T${preservedTime}+05:30`;
      } else {
        // Date only -> keep strictly date only
        formattedSheetVal = `${d}/${mo}/${y}`;
        isoDate = `${y}-${mo}-${d}T00:00:00+05:30`;
      }
    } else {
      formattedSheetVal = safeFormatDate(cleanDateStr, preservedTime ? "dd/MM/yyyy HH:mm:ss" : "dd/MM/yyyy") || cleanDateStr;
      isoDate = convertToIsoStringForFirebase(cleanDateStr);
    }
  }

  // 1. Write formatted date & preserved time to Google Sheet
  sheet.getRange(targetRowIndex, targetCol).setValue(formattedSheetVal);
  SpreadsheetApp.flush();
  Logger.log("Updated Sheet cell [" + targetRowIndex + ", " + targetCol + "] to '" + formattedSheetVal + "'");

  // 2. Update in Firebase / Firestore
  let firebaseSuccess = false;
  let firebaseErrorMsg = "";
  const props = PropertiesService.getScriptProperties();
  const baseUrl = props.getProperty("CRM_API_BASE_URL");
  const token = props.getProperty("CRM_API_TOKEN");

  if (baseUrl && token) {
    try {
      const patchBody = {
        leadUniqueId: leadUniqueId,
        [dateField]: isoDate
      };

      let result;
      try {
        result = crmApiProxy(
          "/gas/updateConsultationVisitedDate",
          "POST",
          JSON.stringify({
            leadUniqueId: leadUniqueId,
            dateField: dateField,
            newDateValue: isoDate,
            consutlationVisitedDate: isoDate
          })
        );
      } catch (err1) {
        result = crmApiProxy(
          "/leads/" + encodeURIComponent(leadUniqueId),
          "PATCH",
          JSON.stringify(patchBody)
        );
      }
      firebaseSuccess = true;
      Logger.log("Firebase patch succeeded for " + leadUniqueId + ": " + JSON.stringify(result));
    } catch (fbErr) {
      firebaseErrorMsg = String(fbErr && fbErr.message ? fbErr.message : fbErr);
      Logger.log("Firebase update failed for " + leadUniqueId + ": " + firebaseErrorMsg);
    }
  } else {
    Logger.log("Firebase API not configured in Script Properties; updated Sheet only.");
  }

  return {
    success: true,
    leadUniqueId: leadUniqueId,
    dateField: dateField,
    formattedValue: formattedSheetVal,
    firebaseUpdated: firebaseSuccess,
    firebaseWarning: firebaseErrorMsg || null,
    message: firebaseSuccess
      ? "Successfully updated date in Google Sheet and Firebase Firestore."
      : (firebaseErrorMsg ? "Updated in Sheet, but Firebase warning: " + firebaseErrorMsg : "Updated in Sheet.")
  };
}

/**
 * Manual Firebase consultation-visited sync (same job as Cloud Scheduler).
 * Requires Script Properties: CRM_API_BASE_URL, CRM_API_TOKEN.
 */
function menuRunConsultationVisitedSync() {
  const ui = SpreadsheetApp.getUi();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Running Firebase consultation visited sync…",
      "CRM Admin",
      10
    );

    const result = crmApiProxy(
      "/sync/consultation-visited",
      "POST",
      JSON.stringify({
        dryRun: false,
        triggeredBy: "apps_script_menu",
      })
    );

    const planned = result && result.updatesPlanned != null ? result.updatesPlanned : "?";
    const applied = result && result.updatesApplied != null ? result.updatesApplied : "?";
    const scanned =
      result && result.scannedBookedLeads != null ? result.scannedBookedLeads : "?";
    const errors =
      result && Array.isArray(result.errors) && result.errors.length
        ? "\nWarnings: " + result.errors.slice(0, 3).join("; ")
        : "";

    ui.alert(
      "Consultation Visited Sync — Done",
      "Booked leads scanned: " +
      scanned +
      "\nUpdates planned: " +
      planned +
      "\nUpdates applied: " +
      applied +
      errors,
      ui.ButtonSet.OK
    );
    Logger.log("menuRunConsultationVisitedSync: " + JSON.stringify(result));
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    Logger.log("menuRunConsultationVisitedSync ERROR: " + msg);
    ui.alert(
      "Consultation Visited Sync — Failed",
      msg +
      "\n\nCheck Script Properties CRM_API_BASE_URL and CRM_API_TOKEN, and that the Firebase API is deployed.",
      ui.ButtonSet.OK
    );
  }
}

/**
 * Server-side proxy for Firebase API calls. Keeps API_SHARED_TOKEN out of the browser.
 * @param {string} path - e.g. "/gas/retriveLeadData" or "/leads/Lead%20ID%20-%201"
 * @param {string} method - GET, POST, or PATCH
 * @param {string|null} payloadJson - JSON body for POST/PATCH
 * @returns {*} Parsed response data field from the API envelope
 */
function crmApiProxy(path, method, payloadJson) {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = String(props.getProperty("CRM_API_BASE_URL") || "").trim();
  const token = String(props.getProperty("CRM_API_TOKEN") || "").trim();
  if (!baseUrl || !token) {
    throw new Error("CRM_API_BASE_URL and CRM_API_TOKEN must be configured in Script Properties.");
  }

  const httpMethod = String(method || "GET").toUpperCase();
  const options = {
    method: "get",
    contentType: "application/json",
    headers: {
      "x-api-token": token,
    },
    muteHttpExceptions: true,
  };

  if (httpMethod === "POST") {
    options.method = "post";
    if (payloadJson) options.payload = payloadJson;
  } else if (httpMethod === "PATCH") {
    options.method = "patch";
    if (payloadJson) options.payload = payloadJson;
  } else if (httpMethod !== "GET") {
    throw new Error("Unsupported HTTP method: " + httpMethod);
  }

  const url = baseUrl.replace(/\/$/, "") + String(path || "");
  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  const raw = response.getContentText();
  let parsed = {};
  try {
    parsed = JSON.parse(raw || "{}");
  } catch (err) {
    parsed = { success: false, message: raw };
  }

  if (status < 200 || status >= 300 || parsed.success === false) {
    const msg =
      (parsed.error && parsed.error.message) ||
      parsed.message ||
      "API failed (" + status + "): " + raw;
    throw new Error(msg);
  }

  if (Object.prototype.hasOwnProperty.call(parsed, "data")) {
    return parsed.data;
  }
  return parsed;
}



function retriveLeadData(filters) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leads Data");
  const data = sheet.getDataRange().getValues();
  const headings = data[0];
  const leadData = data.splice(1);
  const leadDataBase = [];
  const assignedCol = headings.map(h => String(h || "").trim().toLowerCase()).indexOf("assigned to");
  const targetUser = filters && typeof filters === "object" ? String(filters.assignedTo || "").trim() : "";

  leadData.forEach((lead) => {
    if (targetUser && assignedCol !== -1 && String(lead[assignedCol] || "").trim() !== targetUser) {
      return;
    }
    let obj = {};
    headings.forEach((heading, index) => {
      obj[heading.toLowerCase()] = lead[index];
    });
    leadDataBase.push(obj);
  });
  return JSON.stringify(leadDataBase);
}

function updateLeadPriority(leadUniqueIdOrObj, leadPriority) {
  let targetId = leadUniqueIdOrObj;
  let targetPriority = leadPriority;
  if (typeof leadUniqueIdOrObj === "object" && leadUniqueIdOrObj !== null) {
    targetId = leadUniqueIdOrObj.leadUniqueId;
    targetPriority = leadUniqueIdOrObj.selectedPriority || leadUniqueIdOrObj.leadPriority || leadUniqueIdOrObj["lead label"];
  }

  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(header => header.toString().trim().toLowerCase());
  const leadIdColumn = headers.indexOf("lead unique id");
  const leadLabelColumn = headers.indexOf("lead label");

  if (leadIdColumn === -1 || leadLabelColumn === -1) {
    throw new Error('Required "Lead Unique ID" or "Lead Label" column not found.');
  }

  const targetClean = String(targetId || "").replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const rowIdClean = String(data[rowIndex][leadIdColumn] || "").replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
    if (rowIdClean && rowIdClean === targetClean) {
      sheet.getRange(rowIndex + 1, leadLabelColumn + 1).setValue(targetPriority);
      return true;
    }
  }

  throw new Error("Lead not found.");
}

function safeFormatDate(value, pattern) {
  try {
    if (!value) return "";
    const parts = value.split("T");
    if (parts.length === 2) {
      // Manually construct local time
      const [year, month, day] = parts[0].split("-");
      const [hour, minute] = parts[1].split(":");
      const localDate = new Date(
        parseInt(year),
        parseInt(month) - 1,
        parseInt(day),
        parseInt(hour),
        parseInt(minute)
      );
      return Utilities.formatDate(localDate, Session.getScriptTimeZone(), pattern);
    }
    const d = new Date(value);
    if (isNaN(d)) return "";
    return Utilities.formatDate(d, Session.getScriptTimeZone(), pattern);
  } catch (x) {
    return "";
  }
}

function updateLeadDetails(lead) {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
  const data = sheet.getDataRange().getValues();
  const leadData = data.splice(1);
  let folDateLost = "";
  if (lead.leadStatus === "Lost" || lead.leadStatus === "Dead" || lead.status === "Not Visited Lost") {
    folDateLost = new Date();
  } else {
    folDateLost = "";
  }

  for (let i = 0; i < leadData.length; i++) {
    const sheetLeadId = String(leadData[i][1] || "").trim().toLowerCase();
    const targetLeadId = String(lead.leadUniqueId || "").trim().toLowerCase();
    if (sheetLeadId && sheetLeadId === targetLeadId) {
      rowNo = i + 2;
      sheet.getRange(rowNo, 1).setValue(safeFormatDate(lead.leadDate, "dd/MM/yyyy HH:mm:ss")); // Lead Date
      sheet.getRange(rowNo, 2).setValue(lead.leadUniqueId); // Lead Unqiue ID
      sheet.getRange(rowNo, 3).setValue(lead.clientName); // Client Name
      sheet.getRange(rowNo, 4).setValue(lead.mobileNo); // Mobile No
      sheet.getRange(rowNo, 5).setValue(lead.city); // City
      sheet.getRange(rowNo, 6).setValue(lead.town); // Town
      sheet.getRange(rowNo, 7).setValue(lead.callType); // Call Type
      sheet.getRange(rowNo, 8).setValue(lead.platform); // Platform
      sheet.getRange(rowNo, 9).setValue(lead.leadStatus); // Lead Status
      sheet.getRange(rowNo, 10).setValue(lead.rating); // Rating
      sheet.getRange(rowNo, 11).setValue(lead.assignedTo); // Assigned To
      sheet.getRange(rowNo, 12).setValue(lead.interestedIn); // Interested In
      sheet.getRange(rowNo, 13).setValue(folDateLost === "" ? safeFormatDate(lead.followUpDate, "dd/MM/yyyy HH:mm:ss") : Utilities.formatDate(folDateLost, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")); // Follow Up Date

      sheet.getRange(rowNo, 15).setValue(safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss")); // Consultation Booked Date
      sheet.getRange(rowNo, 16).setValue(safeFormatDate(lead.consutlationVisitedDate, "dd/MM/yyyy HH:mm:ss")); // Consultation Visited Date
      sheet.getRange(rowNo, 17).setValue(lead.consultationType); // Consultation Type
      sheet.getRange(rowNo, 18).setValue(lead.initialRemark); // Initial Remark
      sheet.getRange(rowNo, 19).setValue(lead.detailComment); // Detail Comment
      sheet.getRange(rowNo, 20).setValue(lead["Q1"]); // How soon are you looking for solution?
      sheet.getRange(rowNo, 21).setValue(lead["Q2"]); // Do you have a budget in mind?
      sheet.getRange(rowNo, 22).setValue(lead["Q3"]); // Willing to visit to the center?
      // sheet.getRange(rowNo,23).setValue(lead["Q4"]); // Hair fall since?
      // sheet.getRange(rowNo,24).setValue(lead["Q5"]); // Done hair transplant before?
      // sheet.getRange(rowNo,25).setValue(lead["Q6"]); // Gotten a hair system before?
      sheet.getRange(rowNo, 26).setValue(lead["hearAboutUs"]); // How did you hear about us?
      sheet.getRange(rowNo, 36).setValue(lead["consultationStatus"]); // Consultation Status
      sheet.getRange(rowNo, 32).setValue(lead["orderStatus"]); // Order Status
      sheet.getRange(rowNo, 33).setValue(lead["totalSalesAmount"]); // total Sales Amount
      sheet.getRange(rowNo, 37).setValue(lead["notBookedFeedback"]); // Not Booked Feedback
      sheet.getRange(rowNo, 38).setValue(safeFormatDate(lead.orderDate, "dd/MM/yyyy")); // Not Booked Feedback

      const firstContactValue = sheet.getRange(rowNo, 42).getValue()
      Logger.log(firstContactValue)

      if (!firstContactValue) { sheet.getRange(rowNo, 42).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")) }


      if ((lead.followUpDate !== "" || folDateLost !== "") && (Utilities.formatDate(leadData[i][12] === "" ? new Date() : leadData[i][12], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") !== safeFormatDate(lead.followUpDate, "dd/MM/yyyy HH:mm:ss"))) {
        followUpLog(lead);
      }
      let followUpCount = countFollowUps(lead.leadUniqueId);
      sheet.getRange(rowNo, 14).setValue(followUpCount); // No of Follow Ups

      if (safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss") === "" ? "" : Utilities.formatDate(leadData[i][14] === "" ? new Date() : leadData[i][14], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") !== safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss") && (lead["consultationStatus"] === "New" || lead["consultationStatus"] === "Consultation Re-Booked")) {
        consultationLog(lead);

      }
    }
  }

  SpreadsheetApp.flush();

  const updatedLeadData = retriveLeadData();
  return updatedLeadData;
}

function followUpLog(lead) {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Follow-up Logs");
  const dataToadd = [
    new Date(),
    lead.leadDate,
    lead.leadUniqueId,
    lead.clientName,
    lead.mobileNo,
    lead.city,
    lead.town,
    lead.callType,
    lead.platform,
    lead.leadStatus,
    lead.rating,
    lead.assignedTo,
    lead.interestedIn,
    safeFormatDate(lead.followUpDate, "dd/MM/yyyy HH:mm:ss"),
    lead.initialRemark,
    lead.detailComment
  ]

  // if(lead.leadStatus === "Ringing / Busy" && lead.assignedTo === "Nagma"){
  //   const mobileNo = lead.mobileNo;
  //   const message = `Hi ${lead.clientName} Sir,\n\nThank you for reaching out about our hair solutions! 🙌We tried giving you a quick call but couldn’t connect 📞❌\n\nNo worries — whenever you’re free, just give us a ring or drop a message. 💬We’ll be happy to guide you and answer all your questions. ✨`

  //   whatsappApi.sendWhatsAppMessage(mobileNo,message,'nagmaE');

  // }

  sheet.appendRow(dataToadd);

}

function countFollowUps(leadUniqueId) {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Follow-up Logs");
  const data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues(); // Assuming header in row 1
  let count = 0;

  data.forEach(row => {
    if (row[0] === leadUniqueId) {
      count++;
    }
  });
  return count;
}

function consultationLog(lead) {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Consultation Booked Logs");
  const consData = [
    new Date(),
    lead.leadDate,
    lead.leadUniqueId,
    lead.clientName,
    lead.mobileNo,
    lead.city,
    lead.town,
    lead.callType,
    lead.platform,
    lead.leadStatus,
    lead.rating,
    lead.assignedTo,
    lead.interestedIn,
    safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss"),
    lead.initialRemark,
    lead.detailComment,
    lead.consultationType
  ]
  sheet.appendRow(consData);
}

function syncAssignmentUsersToFirestoreConfig() {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Users");
  if (!sheet) throw new Error("Users sheet not found.");

  const userData = sheet.getDataRange().getValues();
  const activeAssignees = [];
  const mondayEligibleAssignees = [];
  const allUsersStatus = [];

  for (let i = 1; i < userData.length; i++) {
    const row = userData[i];
    const username = String(row[0] || "").trim();
    const status = String(row[3] || "").trim();
    const nextMonday = String(row[4] || "").trim();

    if (username) {
      allUsersStatus.push({ username, status });
      if (status === "Active") {
        activeAssignees.push(username);
        if (nextMonday === "Coming") {
          mondayEligibleAssignees.push(username);
        }
      }
    }
  }

  const payload = {
    activeAssignees: activeAssignees,
    mondayEligibleAssignees: mondayEligibleAssignees,
    allUsersStatus: allUsersStatus
  };

  return crmApiProxy("/admin/sync-assignees", "POST", JSON.stringify(payload));
}

/**
 * Automatically triggers in real-time whenever someone edits the "Users" sheet.
 * Instantly syncs active/inactive users to Firestore!
 */
function installedOnEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheetName = e.range.getSheet().getName();
    if (sheetName === "Leads Data") {
      syncFbAdsSourceType_(e.range.getSheet(), e.range.getRow(), e.range.getNumRows());
    }
    if (sheetName === "Users") {
      Logger.log("User sheet edited -> automatically syncing assignees to Firestore...");
      syncAssignmentUsersToFirestoreConfig();
    }
  } catch (err) {
    Logger.log("installedOnEdit sync error: " + err);
  }
}

function onEdit(e) {
  installedOnEdit(e);
}

function installOnEditTrigger() {
  const fn = "installedOnEdit";
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(fn).forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  Logger.log("Installed edit trigger for user sync");
}


function checkLogin(loginDetail) {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const headings = data[0];
  const loginData = data.splice(1);

  for (let i = 0; i < loginData.length; i++) {
    if (loginData[i][0] === loginDetail.username && loginData[i][1] === loginDetail.password) {
      const personObj = {
        username: loginData[i][0],
        password: loginData[i][1],
        permission: loginData[i][2]
      }

      const loginSucess = {
        link: (ScriptApp.getService().getUrl() || "").replace(/\/dev$/, "/exec") + "?page=Home",
        message: "Login Successful",
        user: personObj
      }

      return JSON.stringify(loginSucess)
    }
  }

  return JSON.stringify({
    message: "Username or password is incorrect"
  })
}

function newLeadUpload(lead) {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
  const sheet1 = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Consultation Booked Logs")
  const sheet2 = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Follow-up Logs");
  const unqiueID = getNextLeadId();
  const dataToBeAdded = [
    new Date(),
    unqiueID,
    lead.clientName,
    lead.mobileNo,
    lead.city,
    lead.town,
    lead.callType,
    lead.platform,
    lead.leadStatus,
    lead.rating,
    lead.assignedTo,
    lead.interestedIn,
    safeFormatDate(lead.followUpDate, "dd/MM/yyyy HH:mm:ss"),
    lead.noOfFollowups,
    safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss"),
    safeFormatDate(lead.consutlationVisitedDate, "dd/MM/yyyy HH:mm:ss"),
    lead.consultationType,
    lead.initialRemark,
    lead.detailComment,
    lead["Q1"],
    lead["Q2"],
    lead["Q3"],
    "",
    "",
    "",
    lead.hearAboutUs,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    lead["consultationStatus"],
    "",
    "",
    "",
    "",
    "",
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")
  ]
  sheet.appendRow(dataToBeAdded);

  if (safeFormatDate(lead.followUpDate, "dd/MM/yyyy HH:mm:ss") !== "") {
    const follData = [new Date(), lead.leadDate, unqiueID, lead.clientName, lead.mobileNo, lead.city, lead.town, lead.callType, lead.platform, lead.leadStatus, lead.rating, lead.assignedTo, lead.interestedIn, safeFormatDate(lead.followUpDate, "dd/MM/yyyy HH:mm:ss"), lead.initialRemark, lead.detailComment]
    sheet2.appendRow(follData);
  }

  if (safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss") !== "") {
    const consData = [new Date(), lead.leadDate, unqiueID, lead.clientName, lead.mobileNo, lead.city, lead.town, lead.callType, lead.platform, lead.leadStatus, lead.rating, lead.assignedTo, lead.interestedIn, safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss"), lead.initialRemark, lead.detailComment, lead.consultationType]
    sheet1.appendRow(consData);
  }

  return "<strong>Success!</strong> Lead has been uploaded successfully."
}

function getNextLeadId() {
  const sheetName = "Leads Data";
  const columnToCheck = 2; // Column B
  const prefix = "Lead ID - ";

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet named "${sheetName}" not found.`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    return prefix + "1";
  }

  const idValues = sheet.getRange(1, columnToCheck, lastRow).getValues().flat().filter(String);

  if (idValues.length === 0) {
    return prefix + "1";
  }

  // Search backwards for the last valid "Lead ID - <number>" entry
  for (let i = idValues.length - 1; i >= 0; i--) {
    const val = String(idValues[i] || "").trim();
    const match = val.match(/^Lead ID\s*-\s*(\d+)$/i);
    if (match) {
      const newIdNumber = parseInt(match[1], 10) + 1;
      return `${prefix}${newIdNumber}`;
    }
  }

  return prefix + "1";
}

/**
 * assign in firebase now
 */
// function assignLeads() {
//   const leadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leads Data");
//   const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
//   if (!leadSheet || !userSheet) throw new Error("One or both sheets not found.");

//   const today = new Date();
//   const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

//   const leadData = leadSheet.getDataRange().getValues();
//   const assignCol = 10; // Column K
//   const userData = userSheet.getDataRange().getValues();

//   // Filter users based on day logic
//   const activeUsers = userData.filter(row => {
//     const isActive = row[3] === "Active";
//     const isComing = row[4] === "Coming"; // Column E
//     if (dayOfWeek === 1) { // Monday
//       return row[0] && isActive && isComing;
//     } else {
//       return row[0] && isActive;
//     }
//   }).map(row => row[0]);

//   if (activeUsers.length === 0) throw new Error("No eligible users found for assignment.");

//   const props = PropertiesService.getScriptProperties();
//   const lastAssigned = props.getProperty("lastAssignedUser") || "";
//   let nextIndex = activeUsers.indexOf(lastAssigned) + 1;
//   if (nextIndex >= activeUsers.length || nextIndex === 0) nextIndex = 0;

//   let updated = false;

//   for (let i = 1; i < leadData.length; i++) {
//     const row = leadData[i];
//     const isEmptyRow = row.every(cell => cell === "");
//     const isUnassigned = !row[assignCol];

//     if (!isEmptyRow && isUnassigned) {
//       leadSheet.getRange(i + 1, assignCol + 1).setValue(activeUsers[nextIndex]);
//       props.setProperty("lastAssignedUser", activeUsers[nextIndex]);
//       nextIndex = (nextIndex + 1) % activeUsers.length;
//       updated = true;
//     }
//   }

//   if (!updated) Logger.log("No unassigned leads were found or all rows are empty.");
// }


function updateActualTimes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FMS");
  const data = sheet.getDataRange().getValues();
  const currentTime = new Date();

  for (let i = 6; i < data.length; i++) { // start from row 7 (index 6)
    const status = data[i][10]; // Column I (index 8)
    const consultVisitStatus = data[i][16]; // Column Q (index 16)
    const orderStatus = data[i][21];


    // Check for "Consultation Booked", "Lost", or "Dead" in Column K
    if (["Consultation Booked", "Lost", "Dead", "Consultation Visited", "Not Visited Lost"].includes(status) && data[i][9] === "") {
      sheet.getRange(i + 1, 10).setValue(currentTime); // Column J
    }

    // Check for "Consultation Visited" in Column Q
    if (["Lost", "Consultation Visited", "Dead", "Not Visited Lost"].includes(consultVisitStatus) && data[i][15] === "") {
      sheet.getRange(i + 1, 16).setValue(currentTime); // Column P
    }

    // Check For "Order Status" in Column V
    if (["Booked", "Not Booked"].includes(orderStatus) && data[i][20] === "") {
      sheet.getRange(i + 1, 21).setValue(currentTime); // Column U
    }

  }
}

function checkAndMatchOrder() {
  //LeadDataSheet
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
  const data = sheet.getDataRange().getValues()
  const leadsData = data.splice(1);

  //Order Form 
  const sheet2 = SpreadsheetApp.openById("1BWQb_mxcwH6XCRpNYWCM_J3fYgZZ3Uq5hshqM0_d0l0").getSheetByName("Order Forms");
  const data2 = sheet2.getRange("A1:W").getValues();
  const orderData = data2.splice(1);
  let rowNo = 0;

  for (let i = 0; i < leadsData.length; i++) {
    rowNo = i + 2;
    orderData.forEach((row) => {

      // Get the mobile number from leadsData
      let mobile = leadsData[i][3];

      mobile = String(mobile).trim();

      // Check if the mobile starts with '91' and has 12 digits
      if (mobile && mobile.length === 12 && mobile.startsWith('91')) {
        mobile = mobile.substring(2); // Remove the first two characters ('91')
      }

      if (String(row[2]) === mobile && leadsData[i][8] === "Consultation Visited") {
        sheet.getRange(rowNo, 32).setValue('Booked');
        sheet.getRange(rowNo, 33).setValue(row[22]);
        sheet.getRange(rowNo, 38).setValue(Utilities.formatDate(new Date(row[6]), Session.getScriptTimeZone(), "dd/MM/yyyy"));

      }
    })
  }

  SpreadsheetApp.flush();

  Logger.log("Order Booked Status Updated Sucessfully")
}

function testingWhatsapp() {
  const phoneNumber = '9920975602'
  const message = 'Hi this is a test through Whatsapp libary'
  const sessionID = 'processC'

  whatsappApi.sendWhatsAppMessage(phoneNumber, message, sessionID);
  // whatsappApi.sendWhatsAppButtonMessage(phoneNumber,sessionID);
}

function calculateLeadScore() {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
  const startRow = 20;
  const data = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < data.length; i++) {
    let rowNumber = startRow + i;
    let leadScore = 0;

    // City Question Score
    if (data[i][4] === "Mumbai") {
      leadScore += 10;
    } else if (data[i][4] === "Delhi" || data[i][4] === "Bangalore") {
      leadScore += 8;
    } else if (data[i][4] === "") {
      leadScore += 0;
    } else {
      leadScore += 2;
    }

    // How soon are you looking for solution? Score
    if (data[i][19] === "As soon as possible") {
      leadScore += 5;
    } else if (data[i][19] === "In the next two weeks") {
      leadScore += 4;
    } else if (data[i][19] === "Within a month") {
      leadScore += 2;
    } else if (data[i][19] === "Not urgent") {
      leadScore += 1;
    } else if (data[i][19] === "Not given") {
      leadScore += 0;
    } else {
      leadScore += 0;
    }

    // Do you have a budget in mind? Score
    if (data[i][20] === "Below 20000") {
      leadScore += 1;
    } else if (data[i][20] === "20000 - 35000") {
      leadScore += 16;
    } else if (data[i][20] === "35000 - 45000") {
      leadScore += 18;
    } else if (data[i][20] === "Need the best! cost is not constraint.") {
      leadScore += 20;
    } else if (data[i][19] === "Not given") {
      leadScore += 0;
    } else {
      leadScore += 0;
    }

    // Willing to visit the center? score
    if (data[i][21] === "Yes") {
      leadScore += 10;
    } else if (data[i][21] === "No") {
      leadScore += 2;
    } else if (data[i][19] === "Not given") {
      leadScore += 0;
    } else {
      leadScore += 0;
    }

    // Logger.log(`${data[i][1]} - ${leadScore}`);
    sheet.getRange(rowNumber, 34).setValue(leadScore);

    if (leadScore > 15 || data[i][8] === "Consultation Booked" || data[i][8] === "Consultation Visited") {
      sheet.getRange(rowNumber, 35).setValue("Q");
    } else {
      sheet.getRange(rowNumber, 35).setValue("NQ");
    }
  }
}

/*================================================
    Function To add All Lead To Lead DataBase 
==================================================*/

/**
 * Canonical landing-page slug from a source sheet Path value.
 *
 * Replaces the old `path.slice(1)`, which blindly dropped the first character:
 *   "/stick-on-hair-system"     -> "stick-on-hair-system"    (correct by luck)
 *   "stick-on-hair-system"      -> "tick-on-hair-system"     (WRONG - ate the "s")
 *   "/lp/stick-on-hair-system"  -> "lp/stick-on-hair-system"  (WRONG - kept "lp/")
 *
 * Handles: "slug", "/slug", "lp/slug", "/lp/slug", trailing slashes, blanks.
 * Never removes a real first character. Returns "" when there is no path.
 */
function canonicalFunnelSlug_(raw) {
  var s = String(raw === null || raw === undefined ? "" : raw).trim();
  if (s === "") return "";
  s = s.replace(/^\/+/, "");      // leading slashes
  s = s.replace(/^lp\/+/i, "");   // "lp/" routing prefix
  s = s.replace(/\/+$/, "");      // trailing slashes
  return s.trim();
}

/**
 * Platform string for a Flexi Funnel lead.
 * Falls back to a bare "Flexi Funnel" when no page evidence exists, rather than
 * emitting the meaningless "Flexi Funnel ()".
 */
function flexiFunnelPlatform_(rawPath) {
  var slug = canonicalFunnelSlug_(rawPath);
  return slug === "" ? "Flexi Funnel" : "Flexi Funnel (" + slug + ")";
}

/**
 * Landing Page (AX) label for a Flexi Funnel source path. The two named pages get
 * their reporting labels; any other genuine page keeps its own slug.
 */
function trackingLandingPage_(rawPath) {
  var slug = canonicalFunnelSlug_(rawPath);
  if (slug === "hair-loss-solution-for-corporate-men") return "Corporate Men";
  if (slug === "stick-on-hair-system") return "Stick-On";
  return slug;
}

/** True only for a real campaign name (not blank, a placeholder or an importer default). */
function trackingIsRealCampaign_(campaign) {
  var c = String(campaign || "").trim();
  return c !== "" && !/campaign (name|source)$|^Website Contact Us Form$|^\{\{|placeholder|^unknown$|^null$|^aman$|test|^AHL$/i.test(c);
}

/** Ad platform from a source row's campaign source; blank unless the campaign is real. */
function trackingAdChannel_(campaign, source) {
  if (!trackingIsRealCampaign_(campaign)) return "";
  var s = String(source || "").trim().toLowerCase();
  if (s === "ig" || s === "instagram") return "Instagram";
  if (s === "fb" || s === "facebook") return "Facebook";
  if (/adwords|google/.test(s)) return "Google";
  if (/youtube/.test(s)) return "YouTube";
  return "";
}

/**
 * Writes Source Type, Channel, Funnel, Conversion Method and Landing Page (AT:AX)
 * on the row the backend just created, found by exact Lead Unique ID in column B.
 * Fills only blank cells, only when exactly one row matches, never throws.
 * fields: { at, au, av, aw, ax } - omitted or empty values are left untouched.
 */
function writeTrackingFields_(apiResult, fields) {
  try {
    var data = typeof apiResult === "string" ? JSON.parse(apiResult) : apiResult;
    var leadId = String((data && (data.leadUniqueId || (data.lead && data.lead.leadUniqueId))) || "").trim();
    if (!leadId) return;
    var sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
    var hits = sheet.getRange("B:B").createTextFinder(leadId).matchEntireCell(true).matchCase(true).findAll();
    if (hits.length !== 1) {
      Logger.log("Tracking not written for " + leadId + ": " + hits.length + " matching rows");
      return;
    }
    var cells = sheet.getRange(hits[0].getRow(), 46, 1, 5); // AT:AX
    var current = cells.getValues()[0];
    var wanted = [fields.at, fields.au, fields.av, fields.aw, fields.ax];
    var changed = false;
    for (var i = 0; i < 5; i++) {
      if (String(current[i]).trim() === "" && wanted[i]) { current[i] = wanted[i]; changed = true; }
    }
    if (changed) cells.setValues([current]);
    syncFbAdsSourceType_(sheet, hits[0].getRow(), 1);
  } catch (err) {
    Logger.log("Tracking write failed: " + err);
  }
}

function addNewLeadDataFormSheets() {
  ensureLeadsDataHeaders();
  syncFbAdsSourceType_();

  let leadData = {}

  //Gallabox add Lead Code
  const gallaboxSheet = SpreadsheetApp.openById("1mWcRRbKXwpdSFtFSFJdSr6elOvsmFne1auyn8KagsIo").getSheetByName("addToImport");
  const gallaboxLeadData = gallaboxSheet.getDataRange().getValues().splice(1);


  gallaboxLeadData.forEach((lead, index) => {
    let rowNumber = index + 2;
    if (lead[12] !== "Sent" && lead[3] !== "") {
      leadData["client name"] = lead[1]
      leadData["contact number"] = lead[13]
      leadData["city"] = lead[5]
      leadData["platform"] = "Gallabox"
      leadData["CampaignName"] = lead[9]
      leadData["AdSetName"] = lead[8]
      leadData["AdName"] = lead[10]
      leadData["CampaignSource"] = lead[7]
      leadData["Placement"] = lead[11]
      leadData["Consultation Mode"] = ""
      leadData["secondary contact number"] = ""
      leadData["preffered time to connect"] = ""

      Logger.log(leadData);
      writeTrackingFields_(sendLeadToFirebase(leadData), { at: trackingIsRealCampaign_(leadData["CampaignName"]) ? "Paid" : "", au: trackingAdChannel_(leadData["CampaignName"], leadData["CampaignSource"]) || "WhatsApp", av: "Gallabox", aw: "WhatsApp" });
      gallaboxSheet.getRange(rowNumber, 13).setValue("Sent")
    }
  })

  // Website add Lead code
  const websiteSheet = SpreadsheetApp.openById("1IlQYQC3NQ2PyttM2STk17bQ3EsKpMcusHJH-wN7Dc2U").getSheetByName("Data");
  const websiteLeadData = websiteSheet.getDataRange().getValues().splice(1);

  websiteLeadData.forEach((lead, index) => {
    let rowNumber = index + 2;
    if (lead[11] !== "Sent" && lead[2] !== "") {
      leadData["client name"] = lead[1]
      leadData["contact number"] = lead[2]
      leadData["city"] = lead[3]
      leadData["platform"] = "Website"
      leadData["CampaignName"] = lead[4]
      leadData["AdSetName"] = lead[5]
      leadData["AdName"] = lead[6]
      leadData["CampaignSource"] = lead[7]
      leadData["Placement"] = lead[8]
      leadData["Consultation Mode"] = ""
      leadData["secondary contact number"] = ""
      leadData["preffered time to connect"] = ""

      writeTrackingFields_(sendLeadToFirebase(leadData), { at: trackingIsRealCampaign_(leadData["CampaignName"]) ? "Paid" : "", au: trackingAdChannel_(leadData["CampaignName"], leadData["CampaignSource"]) || "Website", av: "Website Form", aw: "Website Form", ax: "Main Website" });
      websiteSheet.getRange(rowNumber, 12).setValue("Sent");
    }
  })

  // Website Contact Us Detail add Lead code
  const contactUsSheet = SpreadsheetApp.openById("1IlQYQC3NQ2PyttM2STk17bQ3EsKpMcusHJH-wN7Dc2U").getSheetByName("Contact Us Detail");
  const contactUsLeadData = contactUsSheet.getDataRange().getValues().splice(1);

  contactUsLeadData.forEach((lead, index) => {
    let rowNumber = index + 2;
    const sentStatus = String(lead[5] || "").trim();
    const contactNo = String(lead[2] || "").trim();
    if (sentStatus !== "Sent" && contactNo !== "") {
      const email = String(lead[3] || "").trim();
      const query = String(lead[4] || "").trim();

      const contactUsData = {
        "client name": String(lead[1] || "").trim(),
        "contact number": contactNo,
        "email": email,
        "query": query,
        "city": "",
        "platform": "Website",
        "initialRemark": "",
        "CampaignName": "Website Contact Us Form",
        "AdSetName": "Website Contact Us",
        "AdName": "Contact Us Page",
        "CampaignSource": "organic",
        "Placement": "Website",
        "Consultation Mode": "",
        "secondary contact number": "",
        "preffered time to connect": ""
      };

      writeTrackingFields_(sendLeadToFirebase(contactUsData), { au: "Website", av: "Contact Us Form", aw: "Website Form", ax: "Main Website" });
      contactUsSheet.getRange(rowNumber, 6).setValue("Sent");
    }
  });

  // Whatsapp ads add Lead Code
  const whatsappAdsSheet = SpreadsheetApp.openById("14sORB4JN349CTJBOzxByUzXKx4p2P-b8XqyFTq9dPfo").getSheetByName("DataTobeImported");
  const whatsappAdsLeadData = whatsappAdsSheet.getDataRange().getValues().splice(1);

  whatsappAdsLeadData.forEach((lead, index) => {
    let rowNumber = index + 2;
    if (lead[5] !== "Sent" && lead[1] !== "") {

      leadData["client name"] = lead[2]
      leadData["contact number"] = lead[4]
      leadData["city"] = ""
      leadData["platform"] = "Whatsapp Ads"
      leadData["CampaignName"] = "whatsapp ads campaign name"
      leadData["AdSetName"] = "whatsapp ads adset name"
      leadData["AdName"] = "whatsapp ads ad name"
      leadData["CampaignSource"] = "whatsapp ads campaign source"
      leadData["Placement"] = "whatsapp ads placement"
      leadData["Consultation Mode"] = ""
      leadData["secondary contact number"] = ""
      leadData["preffered time to connect"] = ""

      writeTrackingFields_(sendLeadToFirebase(leadData), { au: "WhatsApp", av: "WhatsApp Ads", aw: "WhatsApp" });
      whatsappAdsSheet.getRange(rowNumber, 6).setValue("Sent");

    }
  })

  // Ebook add Lead Code
  const ebookSheet = SpreadsheetApp.openById("12hDMIHru0UIvx4gTw3doWX4ZcExM3laXd-amZTWY6Xk").getSheetByName("Form responses 1");
  const ebookLeadData = ebookSheet.getDataRange().getValues().splice(1);

  ebookLeadData.forEach((lead, index) => {
    let rowNumber = index + 2;
    if (lead[4] !== "Sent" && lead[2] !== "") {

      leadData["client name"] = lead[1]
      leadData["contact number"] = lead[2]
      leadData["city"] = ""
      leadData["platform"] = "Ebook"
      leadData["CampaignName"] = "Ebook campaign name"
      leadData["AdSetName"] = "Ebook adset name"
      leadData["AdName"] = "Ebook ad name"
      leadData["CampaignSource"] = "Ebook campaign source"
      leadData["Placement"] = "Ebook placement"
      leadData["Consultation Mode"] = ""
      leadData["secondary contact number"] = ""
      leadData["preffered time to connect"] = ""

      writeTrackingFields_(sendLeadToFirebase(leadData), { at: "Organic", au: "Website", av: "Ebook", aw: "Ebook Form", ax: "Main Website" });
      ebookSheet.getRange(rowNumber, 5).setValue("Sent");

    }
  })

  /**
   * No need since Onchange trigger in New Lead Fms sheet sends lead to firebase directly 
   */
  // // Social Media Add Lead
  // const socialMediaSheet = SpreadsheetApp.openById("1cuN3zJKtoliHj_Dw0JfVHRRy_tUD2axZGQVcqBg34qA").getSheetByName("Social_Media_Leads");
  // const socialMediaLeadData = socialMediaSheet.getDataRange().getValues().splice(1);

  // socialMediaLeadData.forEach((lead, index) => {
  //   let rowNumber = index + 2;
  //   if (lead[9] !== "Sent" && lead[2] !== "") {

  //     leadData["client name"] = lead[1]
  //     leadData["contact number"] = lead[2]
  //     leadData["city"] = lead[3]
  //     leadData["platform"] = lead[6]
  //     leadData["CampaignName"] = "Social Media campaign name"
  //     leadData["AdSetName"] = "Social Media adset name"
  //     leadData["AdName"] = "Social Media ad name"
  //     leadData["CampaignSource"] = "Social Media campaign source"
  //     leadData["Placement"] = "Social Media placement"
  //     leadData["Consultation Mode"] = ""
  //     leadData["secondary contact number"] = ""
  //     leadData["preffered time to connect"] = ""

  //     sendLeadToFirebase(leadData);
  //     socialMediaSheet.getRange(rowNumber, 10).setValue("Sent");
  //   }
  // })


  // Flexi Funnel Pop Up form add Lead Code
  const flexiFunnelPopUpSheet = SpreadsheetApp.openById("178whhdEBgxDm1ar0SftrwSKDyO5bWeEk-o2eUCr1p74").getSheetByName("POP UP Data");
  const flexiFunnelPopUpLeadData = flexiFunnelPopUpSheet.getDataRange().getValues().splice(1);

  flexiFunnelPopUpLeadData.forEach((lead, index) => {
    let rowNumber = index + 2;
    if (lead[17] !== "Sent" && lead[2] !== "") {

      leadData["client name"] = lead[1]
      leadData["contact number"] = lead[16]
      leadData["city"] = lead[4]
      leadData["platform"] = flexiFunnelPlatform_(lead[13])   // POP UP Data col N = Path
      // leadData["platform"] = `Flexi Funnel`
      leadData["CampaignName"] = lead[6]
      leadData["AdSetName"] = lead[7]
      leadData["AdName"] = lead[8]
      leadData["CampaignSource"] = lead[9]
      leadData["Placement"] = lead[10]
      leadData["Consultation Mode"] = lead[15]
      leadData["secondary contact number"] = lead[3]
      leadData["preffered time to connect"] = lead[14]

      writeTrackingFields_(sendLeadToFirebase(leadData), { at: trackingIsRealCampaign_(leadData["CampaignName"]) ? "Paid" : "", au: trackingAdChannel_(leadData["CampaignName"], leadData["CampaignSource"]), av: canonicalFunnelSlug_(lead[13]), aw: "POP UP", ax: trackingLandingPage_(lead[13]) });
      flexiFunnelPopUpSheet.getRange(rowNumber, 18).setValue("Sent");

    }
  })



  // Flexi Funnel Whatsapp add Lead Code
  const flexiFunnelWhatsappSheet = SpreadsheetApp.openById("178whhdEBgxDm1ar0SftrwSKDyO5bWeEk-o2eUCr1p74").getSheetByName("Whatsapp Import Data");
  const flexifunnelWhatasppLeadData = flexiFunnelWhatsappSheet.getDataRange().getValues().splice(1);

  flexifunnelWhatasppLeadData.forEach((lead, index) => {
    let rowNumber = index + 2;
    if (lead[13] !== "Sent" && lead[2] !== "") {

      leadData["client name"] = lead[1]
      leadData["contact number"] = lead[12]
      leadData["city"] = lead[4]
      leadData["platform"] = flexiFunnelPlatform_(lead[11])   // Whatsapp Import Data col L = Path (header still says "Tracking ID")
      leadData["CampaignName"] = lead[5]
      leadData["AdSetName"] = lead[6]
      leadData["AdName"] = lead[7]
      leadData["CampaignSource"] = lead[8]
      leadData["Placement"] = lead[9]
      leadData["Consultation Mode"] = ""
      leadData["secondary contact number"] = ""
      leadData["preffered time to connect"] = ""

      writeTrackingFields_(sendLeadToFirebase(leadData), { at: trackingIsRealCampaign_(leadData["CampaignName"]) ? "Paid" : "", au: trackingAdChannel_(leadData["CampaignName"], leadData["CampaignSource"]), av: canonicalFunnelSlug_(lead[11]), aw: "WhatsApp", ax: trackingLandingPage_(lead[11]) });
      flexiFunnelWhatsappSheet.getRange(rowNumber, 14).setValue("Sent");

    }
  })

  // Reel To DM System
  const reelToDMSheet = SpreadsheetApp.openById("1AQF_9FrywXnBLnyOd0wt-__80H6qSvsakg23P3aYXew").getSheetByName("Leads");
  const reelToDMLeadData = reelToDMSheet.getDataRange().getValues().splice(1);

  reelToDMLeadData.forEach((lead, index) => {
    let rowNumber = index + 2;
    if (lead[13] !== "Sent" && lead[3] !== "") {

      // Source sheet "Leads" columns:
      //   E(4) Utm Source   F(5) Utm Medium   G(6) Utm Campaign
      //   H(7) Utm Content  I(8) Path
      // Only the two fields with an unambiguous destination are mapped. The old
      // hardcoded "Instagram Reels ..." strings are gone - a missing value is now
      // left blank rather than filled with a synthetic campaign name.
      //
      // Deliberately NOT mapped, pending an approved decision:
      //   Utm Medium  ("DM")        - a medium, not an ad set. No AdSetName source exists.
      //   Utm Content ("40LP01")    - always "<Campaign>01", carries no information
      //                               beyond Utm Campaign. Not a distinct ad name.
      //   Path        ("/40-lp")    - belongs in Landing Page (AX), not in AA-AE.
      //   Placement                 - no source field exists.
      leadData["client name"] = lead[1]
      leadData["contact number"] = lead[12]
      leadData["city"] = ''
      leadData["platform"] = `Instagram (Reels)`
      leadData["CampaignName"] = String(lead[6] || "").trim()    // Utm Campaign
      leadData["AdSetName"] = ""
      leadData["AdName"] = ""
      leadData["CampaignSource"] = String(lead[4] || "").trim()  // Utm Source
      leadData["Placement"] = ""
      leadData["Consultation Mode"] = ""
      leadData["secondary contact number"] = ""
      leadData["preffered time to connect"] = ""

      writeTrackingFields_(sendLeadToFirebase(leadData), { au: "Instagram", av: "Reel to DM", aw: "Instagram DM" });
      reelToDMSheet.getRange(rowNumber, 14).setValue("Sent");

    }
  })

  // assignLeads(); // Auto-assignment is now handled directly by Firebase upon lead creation!

  syncIncomingNumberSourceType_();
  syncFbAdsSourceType_();
}

/** Enforce fb_ads -> Paid for existing rows and each newly imported/edited row. */
function syncFbAdsSourceType_(sheet, firstRow, rowCount) {
  sheet = sheet || SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
  if (!sheet || sheet.getLastRow() < 2) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (value) {
    return String(value).trim().toLowerCase();
  });
  var adsetColumn = headers.indexOf("adsetname") + 1;
  var sourceColumn = headers.indexOf("source type") + 1;
  if (!adsetColumn || !sourceColumn) throw new Error("Leads Data requires AdSetName and Source Type headers.");
  var start = Math.max(2, firstRow || 2);
  var end = rowCount === undefined ? sheet.getLastRow() : Math.min(sheet.getLastRow(), firstRow + rowCount - 1);
  if (end < start) return;
  var count = end - start + 1;
  var adsets = sheet.getRange(start, adsetColumn, count, 1).getValues();
  var sources = sheet.getRange(start, sourceColumn, count, 1).getValues();
  var targets = [];
  for (var i = 0; i < count; i++) {
    if (String(adsets[i][0]).trim().toLowerCase() === "fb_ads" && sources[i][0] !== "Paid") {
      targets.push(sheet.getRange(start + i, sourceColumn).getA1Notation());
    }
  }
  for (var offset = 0; offset < targets.length; offset += 500) {
    sheet.getRangeList(targets.slice(offset, offset + 500)).setValue("Paid");
  }
}

/**
 * Phone-number source rule (the agent selects the number in Z; it is not auto-detected):
 * Z "Incoming through main phone" -> AT Organic, Z "Incoming Through Ads" -> AT Paid.
 * Runs with every import so leads created or edited in the CRM web app are kept in line.
 * Touches only AT on rows with H "Incoming" and one of these two Z values; never throws.
 */
function syncIncomingNumberSourceType_() {
  try {
    var sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
    var rows = sheet.getLastRow() - 1;
    if (rows < 1) return;
    var h = sheet.getRange(2, 8, rows, 1).getValues();
    var z = sheet.getRange(2, 26, rows, 1).getValues();
    var aa = sheet.getRange(2, 27, rows, 1).getValues();
    var at = sheet.getRange(2, 46, rows, 1).getValues();
    for (var i = 0; i < rows; i++) {
      var source = String(z[i][0]).trim();
      var want = source === "Incoming through main phone" ? "Organic" : source === "Incoming Through Ads" ? "Paid" : "";
      if (!want || String(h[i][0]).trim() !== "Incoming" || String(at[i][0]).trim() === want) continue;
      if (want === "Organic" && trackingIsRealCampaign_(aa[i][0])) continue; // real campaign evidence wins
      sheet.getRange(i + 2, 46).setValue(want);
    }
  } catch (err) {
    Logger.log("Incoming number source sync failed: " + err);
  }
}

/**
 * Sends incoming external lead payload directly to Firebase via API Proxy.
 * Auto-assignment round-robin runs on Firebase backend seamlessly.
 */
function sendLeadToFirebase(lead) {
  let vslVideoRemark = "";
  if (lead["platform"] === 'Flexi Funnel (hair-replacement-vsl-video)') {
    vslVideoRemark = "Auto Generated: Please directly pitch this lead for consultation.";
  }

  const rawQuery = lead["query"] || lead["Query"] || lead["customer remark"] || lead["Customer Remark"] || lead["remarks"] || lead["Remarks"] || lead["inquiry"] || lead.query || lead.inquiry || "";
  const queryVal = (rawQuery !== null && rawQuery !== undefined) ? String(rawQuery).trim() : "";

  const payload = {
    clientName: lead["client name"] || "",
    mobileNo: lead["contact number"] || "",
    email: lead["email"] || "",
    query: queryVal,
    city: lead["city"] || "",
    platform: lead["platform"] || "",
    callType: "Outgoing",
    leadStatus: "New Lead",
    initialRemark: lead["initialRemark"] || vslVideoRemark || "",
    campaignname: lead["CampaignName"] || "",
    adsetname: lead["AdSetName"] || "",
    adname: lead["AdName"] || "",
    campaignsource: lead["CampaignSource"] || "",
    placement: lead["Placement"] || "",
    consultationMode: lead["Consultation Mode"] || "",
    secondaryMobileNo: lead["secondary contact number"] || "",
    preferredTimeToCall: lead["preffered time to connect"] || ""
  };

  return crmApiProxy("/gas/newLeadUpload", "POST", JSON.stringify(payload));
}

/*
// COMMENTED OUT ORIGINAL SHEET APPEND FUNCTION FOR REFERENCE / FALLBACK
function addToCRMApp(lead) {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
  const lastRow = sheet.getLastRow();
  let vslVideoRemark = null;

  if (lead["platform"] === 'Flexi Funnel (hair-replacement-vsl-video)') {
    vslVideoRemark = "Auto Generated: Please directly pitch this lead for consultation.";
  } else {
    vslVideoRemark = ""
  }

  const dataToAdd = [
    new Date,
    getNextLeadId(),
    lead["client name"],
    lead["contact number"],
    lead["city"],
    "",
    "Outgoing",
    lead["platform"],
    "New Lead",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    vslVideoRemark,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    lead["CampaignName"],
    lead["AdSetName"],
    lead["AdName"],
    lead["CampaignSource"],
    lead["Placement"],
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    lead["Consultation Mode"],
    lead["secondary contact number"],
    lead["preffered time to connect"]
  ]
  sheet.appendRow(dataToAdd);

}
*/


function checkConsultationVisited() {

  const startRow = 5;
  //Lead Data
  const leadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leads Data");
  const leadSheetData = leadSheet.getRange(`A${startRow}:I`).getValues();

  //Consultation In-person
  const inpersonConsultationSheet = SpreadsheetApp.openById("1QG47SeCSSEGSUrtw1OorA9RkzEqoxr4o2IxIM6h9e6Y").getSheetByName("Consultation FMS");
  const inpersonConsultationData = inpersonConsultationSheet.getRange("A7:D").getValues();

  //After Consultation Data
  const afterConsultationSheet = SpreadsheetApp.openById("1QG47SeCSSEGSUrtw1OorA9RkzEqoxr4o2IxIM6h9e6Y").getSheetByName("After Consultation");
  const afterConsultationData = afterConsultationSheet.getDataRange().getValues().splice(1);

  for (let i = 0; i < leadSheetData.length; i++) {
    let rowNumber = i + startRow;


    // First will Check in Inperson Sheet
    inpersonConsultationData.forEach((row, index) => {

      // Get the mobile number from leadsData
      let mobile = leadSheetData[i][3];

      mobile = String(mobile).trim();

      // Check if the mobile starts with '91' and has 12 digits
      if (mobile && mobile.length === 12 && mobile.startsWith('91')) {
        mobile = mobile.substring(2); // Remove the first two characters ('91')
      }

      if (mobile && String(row[2]) === mobile && leadSheetData[i][8] === "Consultation Booked") {
        leadSheet.getRange(rowNumber, 9).setValue("Consultation Visited");
        leadSheet.getRange(rowNumber, 16).setValue(row[0]);
      }
    })

    //Will check in after Consultation Form
    afterConsultationData.forEach((row, index) => {

      // Get the mobile number from leadsData
      let mobile = leadSheetData[i][3];

      mobile = String(mobile).trim();

      // Check if the mobile starts with '91' and has 12 digits
      if (mobile && mobile.length === 12 && mobile.startsWith('91')) {
        mobile = mobile.substring(2); // Remove the first two characters ('91')
      }

      if (mobile && String(row[2]) === mobile && leadSheetData[i][8] === "Consultation Booked" && (row[4] !== "In-person" && row[5] !== "No Show")) {
        leadSheet.getRange(rowNumber, 9).setValue("Consultation Visited");
        leadSheet.getRange(rowNumber, 16).setValue(row[32]);
      }

    })

  }

}

function addOnlineBookingPaymentDone() {
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");
  const onlineConsultationPaymentSheet = SpreadsheetApp.openById("1EzW040ThQerX4CDkJOqhS9NANl-uWE26dSwkQRoo88Y").getSheetByName("Data")
  const onlineConsultationData = onlineConsultationPaymentSheet.getDataRange().getValues().splice(1);

  onlineConsultationData.forEach((consult, index) => {

    const rowNo = index + 2

    if (consult[21] !== "Added" && consult[9] !== "" && consult[20] !== "") {

      const date = consult[6]; // consult[6] is a Date object (date)
      const time = consult[7]; // consult[7] is a Date object (time)

      // Set the time (hours and minutes) from consult[7] into consult[6]
      date.setHours(time.getHours(), time.getMinutes());

      // Format the merged date and time
      const formattedDateTime = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      // Logger.log("Formatted Date and Time:" + formattedDateTime);

      const payload = {
        clientName: consult[2] || "",
        mobileNo: consult[20] || "",
        city: consult[22] || "",
        callType: "Outgoing",
        platform: "Online Paid Consultation (hair-loss-solutions-for-men)",
        leadStatus: "Consultation Booked",
        assignedTo: "Agent1",
        consutlationBookedDate: formattedDateTime,
        campaignname: consult[23] || "",
        adsetname: consult[24] || "",
        adname: consult[25] || "",
        campaignsource: consult[26] || "",
        placement: consult[27] || ""
      };

      writeTrackingFields_(crmApiProxy("/gas/newLeadUpload", "POST", JSON.stringify(payload)), { at: trackingIsRealCampaign_(payload.campaignname) ? "Paid" : "", au: trackingAdChannel_(payload.campaignname, payload.campaignsource) });
      // sheet.appendRow(dataToAdd);
      onlineConsultationPaymentSheet.getRange(rowNo, 22).setValue("Added");

    }

  })

}

function sendConsultationWhatsAppReminders() {
  const SHEET_NAME = "Leads Data";

  const COL_CLIENT_NAME = 3;    // Column C
  const COL_MOBILE = 4;         // Column D
  const COL_STATUS = 9;         // Column I
  const COL_CONSULTATION = 15;  // Column O

  const REQUIRED_STATUS = "Consultation Booked";

  // Replace these with your exact approved Gallabox template names
  const TEMPLATE_ONE_DAY_BEFORE = "consultation_booked_reminder_1day_before";
  const TEMPLATE_SAME_DAY = "consultation_booked_reminder_same_day";

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet not found: ${SHEET_NAME}`);
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No Consultation Reminder Sent");
    return;
  }

  const lastCol = Math.max(
    COL_CLIENT_NAME,
    COL_MOBILE,
    COL_STATUS,
    COL_CONSULTATION
  );

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();

  const today = startOfDay_(new Date());
  const tomorrow = addDays_(today, 1);

  let sentCount = 0;

  values.forEach((row, index) => {
    const rowNumber = index + 2;

    const clientName = String(displayValues[index][COL_CLIENT_NAME - 1]).trim() || "Sir/Ma’am";
    const mobileNo = String(displayValues[index][COL_MOBILE - 1]).trim();
    const leadStatus = String(displayValues[index][COL_STATUS - 1]).trim();

    const rawConsultationValue = row[COL_CONSULTATION - 1];
    const displayConsultationValue = displayValues[index][COL_CONSULTATION - 1];

    if (leadStatus !== REQUIRED_STATUS) {
      return;
    }

    if (!mobileNo) {
      Logger.log(`Row ${rowNumber}: Mobile number is missing`);
      return;
    }

    const consultationDate = parseConsultationDate_(
      rawConsultationValue,
      displayConsultationValue
    );

    if (!consultationDate) {
      Logger.log(`Row ${rowNumber}: Invalid consultation date`);
      return;
    }

    const consultationDay = startOfDay_(consultationDate);

    // Date and time stored in a separate variable
    const dateAndTime = formatConsultationDateTime_(consultationDate);

    let templateName = "";

    if (isSameDate_(consultationDay, tomorrow)) {
      templateName = TEMPLATE_ONE_DAY_BEFORE;
    } else if (isSameDate_(consultationDay, today)) {
      templateName = TEMPLATE_SAME_DAY;
    } else {
      return;
    }

    try {
      whatsappGallabox.sendConsultationBookedReminder(
        clientName,
        mobileNo,
        templateName,
        dateAndTime
      );

      sentCount++;

      Logger.log(`Reminder sent to ${clientName} - ${mobileNo} for row ${rowNumber}`);
    } catch (error) {
      Logger.log(`Failed to send reminder for row ${rowNumber}: ${error.message}`);
    }
  });

  if (sentCount === 0) {
    Logger.log("No Consultation Reminder Sent");
  } else {
    Logger.log(`Total Consultation Reminders Sent: ${sentCount}`);
  }
}


/**
 * Parses consultation date from either:
 * 1. Actual Google Sheets Date object
 * 2. Text format: dd/MM/yyyy HH:mm:ss
 */
function parseConsultationDate_(rawValue, displayValue) {
  if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
    return rawValue;
  }

  const dateText = String(displayValue).trim();

  if (!dateText) {
    return null;
  }

  // Expected format: dd/MM/yyyy HH:mm:ss
  const match = dateText.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/
  );

  if (!match) {
    return null;
  }

  const day = Number(match[1]);
  const month = Number(match[2]) - 1;
  const year = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);

  const parsedDate = new Date(year, month, day, hour, minute, second);

  if (isNaN(parsedDate.getTime())) {
    return null;
  }

  return parsedDate;
}


/**
 * Formats date like:
 * 23rd April at 2.00pm
 */
function formatConsultationDateTime_(dateObj) {
  const day = dateObj.getDate();

  const monthName = Utilities.formatDate(
    dateObj,
    Session.getScriptTimeZone(),
    "MMMM"
  );

  let hours = dateObj.getHours();
  const minutes = dateObj.getMinutes();

  const ampm = hours >= 12 ? "pm" : "am";

  hours = hours % 12;
  hours = hours ? hours : 12;

  const formattedMinutes = String(minutes).padStart(2, "0");

  return `${day}${getOrdinalSuffix_(day)} ${monthName} at ${hours}.${formattedMinutes}${ampm}`;
}


/**
 * Returns ordinal suffix for dates:
 * 1st, 2nd, 3rd, 4th, etc.
 */
function getOrdinalSuffix_(day) {
  if (day >= 11 && day <= 13) {
    return "th";
  }

  switch (day % 10) {
    case 1:
      return "st";
    case 2:
      return "nd";
    case 3:
      return "rd";
    default:
      return "th";
  }
}


/**
 * Converts date to start of day.
 */
function startOfDay_(dateObj) {
  return new Date(
    dateObj.getFullYear(),
    dateObj.getMonth(),
    dateObj.getDate()
  );
}


/**
 * Adds days to a date.
 */
function addDays_(dateObj, days) {
  const newDate = new Date(dateObj);
  newDate.setDate(newDate.getDate() + days);
  return startOfDay_(newDate);
}


/**
 * Compares two dates without checking time.
 */
function isSameDate_(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
    date1.getMonth() === date2.getMonth() &&
    date1.getDate() === date2.getDate();
}
