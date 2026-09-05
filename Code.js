function doGet(e){
  let temp = "Login"
  if('page' in e.parameters){
    temp = e.parameters['page'][0];
  }
  try {
    const html = HtmlService.createTemplateFromFile(temp);
    const props = PropertiesService.getScriptProperties();
    html.CRM_USE_API_PROXY = props.getProperty("CRM_API_BASE_URL") ? "true" : "";
    html.CRM_WEB_APP_BASE = (ScriptApp.getService().getUrl() || "").replace(/\/dev$/, "/exec");
    html.data = {
      // userData : retriveLeadData(),
      e : e
    }
    return html.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  }catch(e){
    return HtmlService.createHtmlOutput("<h1>Error : </h1>"+ e)
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


function retriveLeadData(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leads Data");
  const data = sheet.getDataRange().getValues();
  const headings = data[0]
  const leadData = data.splice(1);
  const leadDataBase = []
  leadData.forEach((lead)=>{
    let obj = {}
    headings.forEach((heading,index)=>{
      obj[heading.toLowerCase()] = lead[index];
    })
    leadDataBase.push(obj)
  })
  return JSON.stringify(leadDataBase);
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
  const sheet = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Leads Data");
  const data = sheet.getDataRange().getValues();
  const leadData = data.splice(1);
  let folDateLost = "";
  if(lead.leadStatus === "Lost" || lead.leadStatus === "Dead" || lead.status === "Not Visited Lost"){
    folDateLost = new Date();
  }else{
    folDateLost = "";
  }
  
  for(let i=0;i<leadData.length;i++){
    if(leadData[i][1] === lead.leadUniqueId){
      rowNo = i+2;
      sheet.getRange(rowNo, 1).setValue(safeFormatDate(lead.leadDate,"dd/MM/yyyy HH:mm:ss")); // Lead Date
      sheet.getRange(rowNo,2).setValue(lead.leadUniqueId); // Lead Unqiue ID
      sheet.getRange(rowNo,3).setValue(lead.clientName); // Client Name
      sheet.getRange(rowNo,4).setValue(lead.mobileNo); // Mobile No
      sheet.getRange(rowNo,5).setValue(lead.city); // City
      sheet.getRange(rowNo,6).setValue(lead.town); // Town
      sheet.getRange(rowNo,7).setValue(lead.callType); // Call Type
      sheet.getRange(rowNo,8).setValue(lead.platform); // Platform
      sheet.getRange(rowNo,9).setValue(lead.leadStatus); // Lead Status
      sheet.getRange(rowNo,10).setValue(lead.rating); // Rating
      sheet.getRange(rowNo,11).setValue(lead.assignedTo); // Assigned To
      sheet.getRange(rowNo,12).setValue(lead.interestedIn); // Interested In
      sheet.getRange(rowNo, 13).setValue(folDateLost === "" ? safeFormatDate(lead.followUpDate, "dd/MM/yyyy HH:mm:ss") : Utilities.formatDate(folDateLost,Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm:ss")); // Follow Up Date
      
      sheet.getRange(rowNo, 15).setValue(safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss")); // Consultation Booked Date
      sheet.getRange(rowNo, 16).setValue(safeFormatDate(lead.consutlationVisitedDate, "dd/MM/yyyy")); // Consultation Visited Date
      sheet.getRange(rowNo,17).setValue(lead.consultationType); // Consultation Type
      sheet.getRange(rowNo,18).setValue(lead.initialRemark); // Initial Remark
      sheet.getRange(rowNo,19).setValue(lead.detailComment); // Detail Comment
      sheet.getRange(rowNo,20).setValue(lead["Q1"]); // How soon are you looking for solution?
      sheet.getRange(rowNo,21).setValue(lead["Q2"]); // Do you have a budget in mind?
      sheet.getRange(rowNo,22).setValue(lead["Q3"]); // Willing to visit to the center?
      // sheet.getRange(rowNo,23).setValue(lead["Q4"]); // Hair fall since?
      // sheet.getRange(rowNo,24).setValue(lead["Q5"]); // Done hair transplant before?
      // sheet.getRange(rowNo,25).setValue(lead["Q6"]); // Gotten a hair system before?
      sheet.getRange(rowNo,26).setValue(lead["hearAboutUs"]); // How did you hear about us?
      sheet.getRange(rowNo,36).setValue(lead["consultationStatus"]); // Consultation Status
      sheet.getRange(rowNo,32).setValue(lead["orderStatus"]); // Order Status
      sheet.getRange(rowNo,33).setValue(lead["totalSalesAmount"]); // total Sales Amount
      sheet.getRange(rowNo,37).setValue(lead["notBookedFeedback"]); // Not Booked Feedback
      sheet.getRange(rowNo,38).setValue(safeFormatDate(lead.orderDate, "dd/MM/yyyy")); // Not Booked Feedback


    if((lead.followUpDate !== "" || folDateLost !== "") && (Utilities.formatDate(leadData[i][12] === "" ? new Date() : leadData[i][12],Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") !== safeFormatDate(lead.followUpDate, "dd/MM/yyyy HH:mm:ss"))){
      followUpLog(lead);
    }
    let followUpCount = countFollowUps(lead.leadUniqueId);
    sheet.getRange(rowNo,14).setValue(followUpCount); // No of Follow Ups

    if(safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss") === "" ? "" : Utilities.formatDate(leadData[i][14] === "" ? new Date() : leadData[i][14],Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm:ss") !== safeFormatDate(lead.consutlationBookedDate, "dd/MM/yyyy HH:mm:ss") && (lead["consultationStatus"] === "New" || lead["consultationStatus"] === "Consultation Re-Booked")){
      consultationLog(lead);

    }
  }
  }

  SpreadsheetApp.flush();

  const updatedLeadData = retriveLeadData();
  return updatedLeadData;
}

function followUpLog(lead){
  const sheet = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Follow-up Logs");
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
  const sheet = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Follow-up Logs");
  const data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues(); // Assuming header in row 1
  let count = 0;
  
  data.forEach(row => {
    if (row[0] === leadUniqueId) {
      count++;
    }
  });
  return count;
}

function consultationLog(lead){
  const sheet = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Consultation Booked Logs");
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


function checkLogin(loginDetail){
  const sheet = getCrmSpreadsheetForSync().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const headings = data[0];
  const loginData = data.splice(1);

  for(let i=0; i < loginData.length; i++){
    if(loginData[i][0] === loginDetail.username && loginData[i][1] === loginDetail.password){
      const status = String(loginData[i][3] || "").trim();
      if (status !== "Active") {
        return JSON.stringify({
          message: "Your account is inactive. Please contact admin."
        });
      }

      const personObj = {
        username : loginData[i][0],
        password : loginData[i][1],
        permission : loginData[i][2]
      }

      const loginSucess = {
        link : (ScriptApp.getService().getUrl() || "").replace(/\/dev$/, "/exec") + "?page=Home",
        message : "Login Successful",
        user: personObj 
      }

      return JSON.stringify(loginSucess)
    }
  }

  return JSON.stringify({
    message: "Username or password is incorrect"
  }) 
}

function newLeadUpload(lead){
  const sheet = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Leads Data");
  const sheet1 = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Consultation Booked Logs")
  const sheet2 = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Follow-up Logs");
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
    safeFormatDate(lead.consutlationBookedDate,"dd/MM/yyyy HH:mm:ss"),
    safeFormatDate(lead.consutlationVisitedDate,"dd/MM/yyyy HH:mm:ss"),
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
    lead["consultationStatus"]
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

  const lastId = idValues[idValues.length - 1];
  const match = lastId.match(/(\d+)$/);

  if (!match) {
    throw new Error(`The last ID "${lastId}" does not end with a number.`);
  }

  const newIdNumber = parseInt(match[1], 10) + 1;
  // Logger.log(`${prefix}${newIdNumber}`)
  return `${prefix}${newIdNumber}`;
}


function getCrmSpreadsheetForSync() {
  const props = PropertiesService.getScriptProperties();
  const explicitId = props.getProperty("CRM_SYNC_SHEET_ID") || "1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw";
  if (explicitId) {
    return SpreadsheetApp.openById(explicitId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Lead flow (column F) may be a numeric quota/weight band (10–100), or legacy text flags (Yes/No).
 * For safety we treat: number > 0 OR Yes/TRUE/Active/Enabled → eligible for assignment pool.
 * 0 / blank / No → not eligible (weighted distribution is not implemented; numbers are eligibility + future weight hints).
 */
function isLeadFlowEnabled(value) {
  if (value === null || value === undefined || value === "") return false;
  if (value === false) return false;
  if (typeof value === "boolean") return value === true;
  if (typeof value === "number") {
    if (isNaN(value)) return false;
    return value > 0;
  }
  const normalized = String(value).trim().toLowerCase();
  if (!normalized) return false;
  if (["yes", "true", "active", "enabled", "on", "y"].includes(normalized)) return true;
  if (["no", "false", "disabled", "off", "n"].includes(normalized)) return false;
  const numeric = Number(normalized.replace(/,/g, ""));
  if (!isNaN(numeric)) {
    return numeric > 0;
  }
  return false;
}

function getAssignmentEligibleUsers() {
  const spreadsheet = getCrmSpreadsheetForSync();
  const userSheet = spreadsheet.getSheetByName("Users");
  if (!userSheet) throw new Error("Users sheet not found.");

  const userData = userSheet.getDataRange().getValues();
  const today = new Date();
  const dayOfWeek = today.getDay();
  const isMonday = dayOfWeek === 1;
  const users = [];

  for (let i = 1; i < userData.length; i++) {
    const row = userData[i];
    const username = String(row[0] || "").trim();
    const permission = String(row[2] || "").trim();
    const status = String(row[3] || "").trim();
    const nextMonday = String(row[4] || "").trim();
    const leadFlowRaw = row[5];
    const email = String(row[6] || "").trim();

    const leadFlow =
      typeof leadFlowRaw === "number" || typeof leadFlowRaw === "boolean"
        ? leadFlowRaw
        : String(leadFlowRaw || "").trim();

    const reasons = [];
    if (!username) {
      reasons.push("missing_username");
    } else {
      if (status !== "Active") reasons.push("status_not_active");
      if (status === "Active" && !isLeadFlowEnabled(leadFlowRaw)) reasons.push("lead_flow_disabled_or_zero");
      if (
        status === "Active" &&
        isLeadFlowEnabled(leadFlowRaw) &&
        isMonday &&
        nextMonday !== "Coming"
      ) {
        reasons.push("not_coming_monday");
      }
    }

    const eligibleGeneral =
      username !== "" && status === "Active" && isLeadFlowEnabled(leadFlowRaw);
    /** Eligible when assignLeads() runs today: all general-eligible weekdays; Mondays also require Coming. */
    const eligibleToday = eligibleGeneral && (!isMonday || nextMonday === "Coming");
    /** Users who qualify for Monday round-robin (Active + lead flow + Coming), stored for audit/sync. */
    const eligibleMondayPool =
      eligibleGeneral && nextMonday === "Coming";

    users.push({
      username,
      permission,
      status,
      nextMonday,
      leadFlow,
      email,
      eligibleToday,
      eligibleGeneral,
      eligibleMondayPool,
      reason: reasons.length ? reasons.join(",") : "",
    });
  }

  return {
    generatedAt: new Date().toISOString(),
    dayOfWeek,
    users,
  };
}

function assignLeads() {
  const spreadsheet = getCrmSpreadsheetForSync();
  const leadSheet = spreadsheet.getSheetByName("Leads Data");
  if (!leadSheet) {
    throw new Error("Leads Data sheet not found.");
  }

  const eligibility = getAssignmentEligibleUsers();
  const leadData = leadSheet.getDataRange().getValues();
  const assignCol = 10; // Column K
  const activeUsers = eligibility.users
    .filter((user) => user.eligibleToday)
    .map((user) => user.username);

  if (activeUsers.length === 0) {
    Logger.log("assignLeads: No eligible users found for assignment.");
    return {
      success: false,
      assignedRows: 0,
      activeUsers: 0,
      message: "No eligible users found for assignment.",
    };
  }

  const props = PropertiesService.getScriptProperties();
  const lastAssigned = props.getProperty("lastAssignedUser") || "";
  let nextIndex = activeUsers.indexOf(lastAssigned) + 1;
  if (nextIndex >= activeUsers.length || nextIndex === 0) nextIndex = 0;

  let assignedRows = 0;

  for (let i = 1; i < leadData.length; i++) {
    const row = leadData[i];
    const isEmptyRow = row.every(cell => cell === "");
    const isUnassigned = !row[assignCol];

    if (!isEmptyRow && isUnassigned) {
      leadSheet.getRange(i + 1, assignCol + 1).setValue(activeUsers[nextIndex]);
      props.setProperty("lastAssignedUser", activeUsers[nextIndex]);
      nextIndex = (nextIndex + 1) % activeUsers.length;
      assignedRows += 1;
    }
  }

  if (!assignedRows) {
    Logger.log("assignLeads: No unassigned leads were found or all rows are empty.");
  } else {
    Logger.log(`assignLeads: Assigned ${assignedRows} lead(s) using ${activeUsers.length} eligible users.`);
  }

  return {
    success: true,
    assignedRows,
    activeUsers: activeUsers.length,
    dayOfWeek: eligibility.dayOfWeek,
  };
}

function runLeadAssignmentForSync() {
  try {
    const result = assignLeads();
    return {
      success: true,
      ...result,
    };
  } catch (err) {
    Logger.log(`runLeadAssignmentForSync failed: ${err && err.message ? err.message : err}`);
    return {
      success: false,
      assignedRows: 0,
      error: String(err && err.message ? err.message : err),
    };
  }
}

function syncAssignmentUsersToFirestoreConfig() {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = String(props.getProperty("CRM_API_BASE_URL") || "").trim();
  const token = String(props.getProperty("CRM_API_TOKEN") || "").trim();
  if (!baseUrl || !token) {
    throw new Error("CRM_API_BASE_URL and CRM_API_TOKEN must be configured in Script Properties.");
  }

  const eligibility = getAssignmentEligibleUsers();
  const activeAssignees = eligibility.users
    .filter((user) => user.eligibleGeneral)
    .map((user) => user.username);
  /** Matches Monday pool + audit: Active + lead flow + next monday = Coming (independent of today's weekday). */
  const mondayEligibleAssignees = eligibility.users
    .filter((user) => user.eligibleMondayPool)
    .map((user) => user.username);

  const payload = {
    activeAssignees,
    mondayEligibleAssignees,
    source: "google_sheet_users",
    syncedBy: "apps_script",
    usersSnapshot: eligibility.users,
  };

  const endpoint = `${baseUrl.replace(/\/$/, "")}/admin/sync-assignees`;
  const response = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-token": token,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const raw = response.getContentText();
  let parsed = {};
  try {
    parsed = JSON.parse(raw || "{}");
  } catch (err) {
    parsed = { success: false, raw };
  }

  if (status < 200 || status >= 300 || parsed.success === false) {
    throw new Error(`syncAssignmentUsersToFirestoreConfig failed (${status}): ${raw}`);
  }

  return {
    success: true,
    endpoint,
    status,
    activeAssigneesCount: activeAssignees.length,
    mondayEligibleAssigneesCount: mondayEligibleAssignees.length,
    usersScanned: eligibility.users.length,
    backend: parsed.data || parsed,
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CRM Admin")
    .addItem("Sync Assignment Users to Firestore", "menuSyncAssignmentUsersToFirestore")
    .addItem("Show Eligible Assignment Users", "menuShowEligibleAssignmentUsers")
    .addItem("Run Lead Assignment", "menuRunLeadAssignment")
    .addSeparator()
    .addItem("Install hourly assignment sync trigger", "installAssignmentSyncHourlyTrigger")
    .addItem("Install 15-minute assignment sync trigger", "installAssignmentSync15MinTrigger")
    .addItem("Remove assignment sync triggers", "removeAssignmentSyncTriggers")
    .addToUi();
}

function menuSyncAssignmentUsersToFirestore() {
  try {
    const result = syncAssignmentUsersToFirestoreConfig();
    const summary = [
      "Sync completed successfully.",
      "Active assignees (Firestore fallback): " + result.activeAssigneesCount,
      "Monday pool (next monday = Coming): " + result.mondayEligibleAssigneesCount,
      "Users rows scanned: " + result.usersScanned,
    ].join("\n");
    Logger.log("menuSyncAssignmentUsersToFirestore: " + JSON.stringify(result));
    SpreadsheetApp.getUi().alert(summary);
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    Logger.log("menuSyncAssignmentUsersToFirestore ERROR: " + msg);
    SpreadsheetApp.getUi().alert("Sync failed:\n" + msg);
  }
}

function menuShowEligibleAssignmentUsers() {
  try {
    const eligibility = getAssignmentEligibleUsers();
    const general = eligibility.users.filter((u) => u.eligibleGeneral).map((u) => u.username);
    const today = eligibility.users.filter((u) => u.eligibleToday).map((u) => u.username);
    const mondayPool = eligibility.users.filter((u) => u.eligibleMondayPool).map((u) => u.username);
    const lines = [
      "Day of week (0=Sun): " + eligibility.dayOfWeek,
      "Eligible for Firestore activeAssignees (" + general.length + "): " + general.join(", "),
      "Eligible for assignLeads today (" + today.length + "): " + today.join(", "),
      "Monday pool / mondayEligibleAssignees (" + mondayPool.length + "): " + mondayPool.join(", "),
    ];
    Logger.log("menuShowEligibleAssignmentUsers: " + lines.join(" | "));
    SpreadsheetApp.getUi().alert(lines.join("\n\n"));
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    Logger.log("menuShowEligibleAssignmentUsers ERROR: " + msg);
    SpreadsheetApp.getUi().alert("Could not read Users:\n" + msg);
  }
}

function menuRunLeadAssignment() {
  try {
    const result = assignLeads();
    const summary = result.success
      ? [
          "Lead assignment finished.",
          "Rows assigned: " + result.assignedRows,
          "Eligible users today: " + result.activeUsers,
          "Day of week: " + result.dayOfWeek,
        ].join("\n")
      : String(result.message || "No assignment performed.");
    Logger.log("menuRunLeadAssignment: " + JSON.stringify(result));
    SpreadsheetApp.getUi().alert(summary);
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    Logger.log("menuRunLeadAssignment ERROR: " + msg);
    SpreadsheetApp.getUi().alert("Assignment failed:\n" + msg);
  }
}

/**
 * Called by time-driven triggers; logs compact success line plus full error on failure.
 */
function timedSyncAssignmentUsersToFirestore() {
  try {
    const result = syncAssignmentUsersToFirestoreConfig();
    Logger.log(
      "timedSyncAssignmentUsersToFirestore OK activeAssignees=" +
        result.activeAssigneesCount +
        " mondayPool=" +
        result.mondayEligibleAssigneesCount
    );
  } catch (err) {
    Logger.log(
      "timedSyncAssignmentUsersToFirestore ERROR: " + String(err && err.message ? err.message : err)
    );
  }
}

function removeAssignmentSyncTriggersInternal() {
  const fn = "timedSyncAssignmentUsersToFirestore";
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === fn) {
      ScriptApp.deleteTrigger(t);
      removed += 1;
    }
  });
  Logger.log("removeAssignmentSyncTriggers: removed count=" + removed);
  return removed;
}

function removeAssignmentSyncTriggers() {
  const removed = removeAssignmentSyncTriggersInternal();
  SpreadsheetApp.getUi().alert("Removed " + removed + " assignment sync trigger(s).");
}

function installAssignmentSyncHourlyTrigger() {
  try {
    removeAssignmentSyncTriggersInternal();
    ScriptApp.newTrigger("timedSyncAssignmentUsersToFirestore").timeBased().everyHours(1).create();
    SpreadsheetApp.getUi().alert(
      "Installed: hourly sync trigger → timedSyncAssignmentUsersToFirestore()\n(Extensions → Apps Script → Triggers)"
    );
    Logger.log("installAssignmentSyncHourlyTrigger: created hourly trigger");
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    Logger.log("installAssignmentSyncHourlyTrigger ERROR: " + msg);
    SpreadsheetApp.getUi().alert("Could not install trigger:\n" + msg);
  }
}

function installAssignmentSync15MinTrigger() {
  try {
    removeAssignmentSyncTriggersInternal();
    ScriptApp.newTrigger("timedSyncAssignmentUsersToFirestore").timeBased().everyMinutes(15).create();
    SpreadsheetApp.getUi().alert(
      "Installed: 15-minute sync trigger → timedSyncAssignmentUsersToFirestore()\n(Extensions → Apps Script → Triggers)"
    );
    Logger.log("installAssignmentSync15MinTrigger: created 15-min trigger");
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    Logger.log("installAssignmentSync15MinTrigger ERROR: " + msg);
    SpreadsheetApp.getUi().alert("Could not install trigger:\n" + msg);
  }
}

function updateActualTimes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FMS");
  const data = sheet.getDataRange().getValues();
  const currentTime = new Date();

  for (let i = 6; i < data.length; i++) { // start from row 7 (index 6)
    const status = data[i][10]; // Column I (index 8)
    const consultVisitStatus = data[i][16]; // Column Q (index 16)
    const orderStatus = data[i][21];


    // Check for "Consultation Booked", "Lost", or "Dead" in Column I
    if (["Consultation Booked", "Lost", "Dead","Consultation Visited","Not Visited Lost"].includes(status) && data[i][9] === "") {
      sheet.getRange(i + 1, 10).setValue(currentTime); // Column J
    }

    // Check for "Consultation Visited" in Column Q
    if (["Lost","Consultation Visited","Dead","Not Visited Lost"].includes(consultVisitStatus) && data[i][15] === "") {
      sheet.getRange(i + 1, 16).setValue(currentTime); // Column P
    }

    // Check For "Order Status" in Column V
    if(["Booked","Not Booked"].includes(orderStatus) && data[i][20] === ""){
      sheet.getRange(i + 1, 21).setValue(currentTime); // Column U
    }

  }
}

function checkAndMatchOrder(){
  //LeadDataSheet
  const sheet = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Leads Data");
  const data = sheet.getDataRange().getValues()
  const leadsData = data.splice(1);

  //Order Form 
  const sheet2 = SpreadsheetApp.openById("1BWQb_mxcwH6XCRpNYWCM_J3fYgZZ3Uq5hshqM0_d0l0").getSheetByName("Order Forms");
  const data2 = sheet2.getRange("A1:W").getValues();
  const orderData = data2.splice(1);
  let rowNo = 0;
  
  for(let i=0; i < leadsData.length; i++) {
    rowNo = i+2;
    orderData.forEach((row)=>{

      // Get the mobile number from leadsData
      let mobile = leadsData[i][3];

      mobile = String(mobile).trim();

      // Check if the mobile starts with '91' and has 12 digits
      if (mobile && mobile.length === 12 && mobile.startsWith('91')) {
        mobile = mobile.substring(2); // Remove the first two characters ('91')
      }    
      
      if(String(row[2]) === mobile && leadsData[i][8] === "Consultation Visited"){
        sheet.getRange(rowNo,32).setValue('Booked');
        sheet.getRange(rowNo,33).setValue(row[22]);
        sheet.getRange(rowNo,38).setValue(Utilities.formatDate(new Date(row[6]),Session.getScriptTimeZone(),"dd/MM/yyyy"));
        
      }
    }) 
  }

  SpreadsheetApp.flush();

  Logger.log("Order Booked Status Updated Sucessfully") 
}

function testingWhatsapp(){
  const phoneNumber = '9920975602'
  const message = 'Hi this is a test through Whatsapp libary'
  const sessionID = 'processC'

  whatsappApi.sendWhatsAppMessage(phoneNumber,message,sessionID);
  // whatsappApi.sendWhatsAppButtonMessage(phoneNumber,sessionID);
}

function calculateLeadScore() {
  const sheet = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Leads Data");
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
    } else if(data[i][19] === "Not given"){
      leadScore += 0;
    }else {
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
    } else if(data[i][19] === "Not given"){
      leadScore += 0;
    } else {
      leadScore += 0;
    }

    // Willing to visit the center? score
    if (data[i][21] === "Yes") {
      leadScore += 10;
    } else if (data[i][21] === "No") {
      leadScore += 2;
    } else if(data[i][19] === "Not given"){
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

function addNewLeadDataFormSheets(){

  let leadData = {}

  //Gallabox add Lead Code
  const gallaboxSheet = SpreadsheetApp.openById("1mWcRRbKXwpdSFtFSFJdSr6elOvsmFne1auyn8KagsIo").getSheetByName("addToImport");
  const gallaboxLeadData = gallaboxSheet.getDataRange().getValues().splice(1);
  
  
  gallaboxLeadData.forEach((lead,index)=>{
    let rowNumber = index+2;
    if(lead[12] !== "Sent" && lead[3] !== ""){
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
      addToCRMApp(leadData);
      gallaboxSheet.getRange(rowNumber,13).setValue("Sent")
    }   
  })
  
  // Website add Lead code
  const websiteSheet = SpreadsheetApp.openById("1IlQYQC3NQ2PyttM2STk17bQ3EsKpMcusHJH-wN7Dc2U").getSheetByName("Data");
  const websiteLeadData = websiteSheet.getDataRange().getValues().splice(1);
  
  websiteLeadData.forEach((lead,index)=>{
    let rowNumber = index+2;
    if(lead[11] !== "Sent" && lead[2] !== ""){
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

    addToCRMApp(leadData);
    websiteSheet.getRange(rowNumber,12).setValue("Sent");
    }
  })

  // Whatsapp ads add Lead Code
  const whatsappAdsSheet = SpreadsheetApp.openById("14sORB4JN349CTJBOzxByUzXKx4p2P-b8XqyFTq9dPfo").getSheetByName("DataTobeImported");
  const whatsappAdsLeadData = whatsappAdsSheet.getDataRange().getValues().splice(1);

  whatsappAdsLeadData.forEach((lead,index)=>{
    let rowNumber = index + 2;
    if(lead[5] !== "Sent" && lead[1] !== ""){

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

    addToCRMApp(leadData);
    whatsappAdsSheet.getRange(rowNumber,6).setValue("Sent");

    }
  })

  // Ebook add Lead Code
  const ebookSheet = SpreadsheetApp.openById("12hDMIHru0UIvx4gTw3doWX4ZcExM3laXd-amZTWY6Xk").getSheetByName("Form responses 1");
  const ebookLeadData = ebookSheet.getDataRange().getValues().splice(1);

  ebookLeadData.forEach((lead,index)=>{
    let rowNumber = index + 2;
    if(lead[4] !== "Sent" && lead[2] !== ""){

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

     addToCRMApp(leadData);
     ebookSheet.getRange(rowNumber,5).setValue("Sent");

    }
  })


  // Social Media Add Lead
  const socialMediaSheet = SpreadsheetApp.openById("1cuN3zJKtoliHj_Dw0JfVHRRy_tUD2axZGQVcqBg34qA").getSheetByName("Social_Media_Leads");
  const socialMediaLeadData = socialMediaSheet.getDataRange().getValues().splice(1);

  socialMediaLeadData.forEach((lead,index)=>{
    let rowNumber = index + 2;
    if(lead[9] !== "Sent" && lead[2] !== ""){

      leadData["client name"] = lead[1]
      leadData["contact number"] = lead[2]
      leadData["city"] = lead[3]
      leadData["platform"] = lead[6]
      leadData["CampaignName"] = "Social Media campaign name"
      leadData["AdSetName"] = "Social Media adset name"
      leadData["AdName"] = "Social Media ad name"
      leadData["CampaignSource"] = "Social Media campaign source"
      leadData["Placement"] = "Social Media placement"
      leadData["Consultation Mode"] = ""
      leadData["secondary contact number"] = ""
      leadData["preffered time to connect"] = ""

      addToCRMApp(leadData);    
      socialMediaSheet.getRange(rowNumber,10).setValue("Sent");
    }
  })

  
  // Flexi Funnel Pop Up form add Lead Code
  const flexiFunnelPopUpSheet = SpreadsheetApp.openById("178whhdEBgxDm1ar0SftrwSKDyO5bWeEk-o2eUCr1p74").getSheetByName("POP UP Data");
  const flexiFunnelPopUpLeadData = flexiFunnelPopUpSheet.getDataRange().getValues().splice(1);

  flexiFunnelPopUpLeadData.forEach((lead,index)=>{
    let rowNumber = index + 2;
    if(lead[17] !== "Sent" && lead[2] !== ""){
      
      leadData["client name"] = lead[1]
      leadData["contact number"] = lead[16]
      leadData["city"] = lead[4]
      leadData["platform"] = `Flexi Funnel (${lead[13].slice(1)})`
      // leadData["platform"] = `Flexi Funnel`
      leadData["CampaignName"] = lead[6]
      leadData["AdSetName"] = lead[7]
      leadData["AdName"] = lead[8]
      leadData["CampaignSource"] = lead[9]
      leadData["Placement"] = lead[10]
      leadData["Consultation Mode"] = lead[15]
      leadData["secondary contact number"] = lead[3]
      leadData["preffered time to connect"] = lead[14]

    addToCRMApp(leadData);
    flexiFunnelPopUpSheet.getRange(rowNumber,18).setValue("Sent");

    }
  })



  // Flexi Funnel Whatsapp add Lead Code
  const flexiFunnelWhatsappSheet = SpreadsheetApp.openById("178whhdEBgxDm1ar0SftrwSKDyO5bWeEk-o2eUCr1p74").getSheetByName("Whatsapp Import Data");
  const flexifunnelWhatasppLeadData = flexiFunnelWhatsappSheet.getDataRange().getValues().splice(1);
  
  flexifunnelWhatasppLeadData.forEach((lead,index)=>{
    let rowNumber = index + 2;
    if(lead[13] !== "Sent" && lead[2] !== ""){
      
      leadData["client name"] = lead[1]
      leadData["contact number"] = lead[12]
      leadData["city"] = lead[4]
      leadData["platform"] = `Flexi Funnel (${lead[11].slice(1)})`
      leadData["CampaignName"] = lead[5]
      leadData["AdSetName"] = lead[6]
      leadData["AdName"] = lead[7]
      leadData["CampaignSource"] = lead[8]
      leadData["Placement"] = lead[9]
      leadData["Consultation Mode"] = ""
      leadData["secondary contact number"] = ""
      leadData["preffered time to connect"] = ""

    addToCRMApp(leadData);
    flexiFunnelWhatsappSheet.getRange(rowNumber,14).setValue("Sent");

    }
  })

  assignLeads();

}

function addToCRMApp(lead){
  const sheet = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Leads Data");
  const lastRow = sheet.getLastRow();
  let vslVideoRemark = null;

  if(lead["platform"] === 'Flexi Funnel (hair-replacement-vsl-video)'){
    vslVideoRemark = "Auto Generated: Please directly pitch this lead for consultation.";
  }else {
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


function checkConsultationVisited(){

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

  for(let i=0; i<leadSheetData.length; i++){
    let rowNumber = i + startRow;
    

    // First will Check in Inperson Sheet
    inpersonConsultationData.forEach((row,index)=>{

      // Get the mobile number from leadsData
      let mobile = leadSheetData[i][3];

      mobile = String(mobile).trim();

      // Check if the mobile starts with '91' and has 12 digits
      if (mobile && mobile.length === 12 && mobile.startsWith('91')) {
        mobile = mobile.substring(2); // Remove the first two characters ('91')
      }    
      
      if(mobile && String(row[2]) === mobile && leadSheetData[i][8] === "Consultation Booked"){
        leadSheet.getRange(rowNumber,9).setValue("Consultation Visited");
        leadSheet.getRange(rowNumber,16).setValue(row[0]);
      } 
    })

    //Will check in after Consultation Form
    afterConsultationData.forEach((row,index)=>{

    // Get the mobile number from leadsData
      let mobile = leadSheetData[i][3];

      mobile = String(mobile).trim();

    // Check if the mobile starts with '91' and has 12 digits
      if (mobile && mobile.length === 12 && mobile.startsWith('91')) {
        mobile = mobile.substring(2); // Remove the first two characters ('91')
      }        

    if(mobile && String(row[2]) === mobile && leadSheetData[i][8] === "Consultation Booked" && (row[4] !== "In-person" && row[5] !== "No Show")){
      leadSheet.getRange(rowNumber,9).setValue("Consultation Visited");
      leadSheet.getRange(rowNumber,16).setValue(row[32]);
    }

    })

  }

}

function addOnlineBookingPaymentDone() {
  const sheet = SpreadsheetApp.openById("1awaq6gNQspr-sjp-yuaH-Nu5JJmhyplVRevoaj47JZw").getSheetByName("Leads Data"); 
  const onlineConsultationPaymentSheet = SpreadsheetApp.openById("1EzW040ThQerX4CDkJOqhS9NANl-uWE26dSwkQRoo88Y").getSheetByName("Data")
  const onlineConsultationData = onlineConsultationPaymentSheet.getDataRange().getValues().splice(1);

  onlineConsultationData.forEach((consult, index)=>{

    const rowNo = index + 2
    
    if(consult[21] !== "Added" && consult[9] !== "" && consult[20] !== "") {
    
      const date = consult[6]; // consult[6] is a Date object (date)
      const time = consult[7]; // consult[7] is a Date object (time)

      // Set the time (hours and minutes) from consult[7] into consult[6]
      date.setHours(time.getHours(), time.getMinutes());

      // Format the merged date and time
      const formattedDateTime = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      // Logger.log("Formatted Date and Time:" + formattedDateTime);
    
    const dataToAdd = [
    new Date,
    getNextLeadId(),
    consult[2],
    consult[20],
    consult[22],
    "",
    "Outgoing",
    "Online Paid Consultation (hair-loss-solutions-for-men)",
    "Consultation Booked",
    "",
    "Agent1",
    "",
    "",
    "",
    formattedDateTime,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    consult[23],
    consult[24],
    consult[25],
    consult[26],
    consult[27]
  ]

  sheet.appendRow(dataToAdd);
  onlineConsultationPaymentSheet.getRange(rowNo,22).setValue("Added");

  }

})

}


function testWhatsappGallabox(){
  whatsappGallabox.processContactAndSendMessage("KUNAL DEMO", 9372356905,"customerservice1@americanhairline.com");
}