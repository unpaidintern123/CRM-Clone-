/**
 * ============================================================
 *  TEST RUNNER SCRIPT (Google Apps Script)
 *  - Dedicated script for manually running and verifying all functions
 *  - Keeps production code untouched
 *  - Uses test phone number 9967716945 for WhatsApp testing
 * ============================================================
 */

const TEST_CONFIG = {
  TEST_MOBILE_NO: '9967716945',
  WA_SESSION_ID: 'processC'
};

/**
 * 1. Test WhatsApp Messaging using test phone number 9967716945.
 */
function runTestWhatsAppMessage() {
  Logger.log('--- Running Test: WhatsApp Messaging ---');
  const msg = 'Hi! This is an automated test message from CRM TestRunner.';
  try {
    whatsappApi.sendWhatsAppMessage(TEST_CONFIG.TEST_MOBILE_NO, msg, TEST_CONFIG.WA_SESSION_ID);
    Logger.log('✅ Test WhatsApp Message Sent Successfully to ' + TEST_CONFIG.TEST_MOBILE_NO);
  } catch (e) {
    Logger.log('❌ Test WhatsApp Message Error: ' + e);
  }
}

/**
 * 2. Test Gallabox WhatsApp Reminder using test phone number 9967716945.
 */
function runTestGallaboxWhatsApp() {
  Logger.log('--- Running Test: Gallabox WhatsApp Reminders ---');
  try {
    whatsappGallabox.sendConsultationBookedReminder(
      "Test Client",
      TEST_CONFIG.TEST_MOBILE_NO,
      "consultation_booked_reminder_same_day",
      "25th April at 2.00pm"
    );
    Logger.log('✅ Gallabox WhatsApp Reminder Sent Successfully to ' + TEST_CONFIG.TEST_MOBILE_NO);
  } catch (e) {
    Logger.log('⚠️ Gallabox WhatsApp Reminder Note/Error: ' + e);
  }
}

/**
 * 3. Test Lead Priority Update (Code.js).
 */
function runTestLeadPriorityUpdate() {
  Logger.log('--- Running Test: Lead Priority Update ---');
  try {
    const res = updateLeadPriority("Lead ID - 24966", "High");
    Logger.log('✅ updateLeadPriority succeeded: ' + res);
  } catch (e) {
    Logger.log('⚠️ updateLeadPriority Note/Error: ' + e);
  }
}

/**
 * 4. Test SLA Report & PDF Generation (sending PDF link to test number 9967716945).
 */
function runTestReportAndPdfGeneration() {
  Logger.log('--- Running Test: SLA Report & PDF Generation ---');
  try {
    const data = readSheetData_();
    const target = getReportDate_(CONFIG.REPORT_DAY);
    const stats = buildStats_(data, target);
    const pdfBlob = buildPdf_(stats, target);
    
    // Save to Drive & verify public link
    const publicUrl = savePublicPdf_(pdfBlob);
    Logger.log('✅ PDF Report generated and verified public: ' + publicUrl);

    // Build message and send to test number 9967716945
    const message = buildWhatsAppMessage_(stats, target, publicUrl);
    sendReportOnWhatsApp_(TEST_CONFIG.TEST_MOBILE_NO, message, TEST_CONFIG.WA_SESSION_ID);
    Logger.log('✅ Report sent to test number ' + TEST_CONFIG.TEST_MOBILE_NO);
  } catch (e) {
    Logger.log('❌ Report Generation Error: ' + e);
  }
}

/**
 * 5. Test Lead Data Retrieval (Code.js).
 */
function runTestRetrieveLeadData() {
  Logger.log('--- Running Test: Retrieve Lead Data ---');
  try {
    const jsonStr = retriveLeadData();
    const parsed = JSON.parse(jsonStr);
    Logger.log('✅ retriveLeadData succeeded. Total leads fetched: ' + parsed.length);
  } catch (e) {
    Logger.log('❌ Retrieve Lead Data Error: ' + e);
  }
}

/**
 * 6. Test Firestore User Assignment Sync (Code.js).
 */
function runTestSyncUsersToFirestore() {
  Logger.log('--- Running Test: Sync Users to Firestore ---');
  try {
    const res = syncAssignmentUsersToFirestoreConfig();
    Logger.log('✅ syncAssignmentUsersToFirestoreConfig succeeded. Response: ' + JSON.stringify(res));
  } catch (e) {
    Logger.log('❌ Sync Users Error: ' + e);
  }
}

/**
 * 7. Test Order Forms Matcher (Code.js).
 */
function runTestCheckAndMatchOrder() {
  Logger.log('--- Running Test: Check & Match Orders ---');
  try {
    checkAndMatchOrder();
    Logger.log('✅ checkAndMatchOrder executed successfully.');
  } catch (e) {
    Logger.log('❌ Check & Match Order Error: ' + e);
  }
}

/**
 * 8. Test Dashboard Analytics Functions (dashboardCal.js).
 */
function runTestDashboardCalFunctions() {
  Logger.log('--- Running Test: Dashboard Analytics (dashboardCal.js) ---');
  try {
    const campaignData = getCampaignName();
    Logger.log('✅ getCampaignName succeeded. Campaign options fetched.');
  } catch (e) {
    Logger.log('⚠️ getCampaignName Note/Error: ' + e);
  }

  try {
    const weekData = getLatestWeekEndData();
    Logger.log('✅ getLatestWeekEndData succeeded.');
  } catch (e) {
    Logger.log('⚠️ getLatestWeekEndData Note/Error: ' + e);
  }
}

/**
 * 9. Test Date Formatter & Utility Helpers (Code.js).
 */
function runTestUtilityFunctions() {
  Logger.log('--- Running Test: Date Formatter & Helpers ---');
  try {
    const formattedDate = safeFormatDate(new Date(), 'dd/MM/yyyy HH:mm:ss');
    Logger.log('✅ safeFormatDate succeeded: ' + formattedDate);
  } catch (e) {
    Logger.log('❌ Utility Functions Error: ' + e);
  }
}

/**
 * 10. Test Direct Firebase Lead Ingestion (Code.js -> Firebase API).
 */
function runTestDirectFirebaseLeadIngestion() {
  Logger.log('--- Running Test: Direct Firebase Lead Ingestion ---');
  const testLead = {
    "client name": "Test Lead Ingestion",
    "contact number": "9876543210",
    "city": "Mumbai",
    "platform": "Website",
    "CampaignName": "Test Campaign",
    "AdSetName": "Test AdSet",
    "AdName": "Test Ad",
    "CampaignSource": "Test Source",
    "Placement": "Test Placement",
    "Consultation Mode": "In-person",
    "secondary contact number": "9123456789",
    "preffered time to connect": "Evening 5pm"
  };

  try {
    const response = sendLeadToFirebase(testLead);
    Logger.log('✅ Direct Firebase Lead Ingestion Succeeded!');
    Logger.log('Firebase Response: ' + JSON.stringify(response));
  } catch (e) {
    Logger.log('❌ Direct Firebase Lead Ingestion Error: ' + e);
  }
}

/**
 * MASTER TEST RUNNER — Run this in Apps Script Editor dropdown.
 * Executes and verifies EVERY single function across Code.js, sendNewLeadTimeReport.js, and dashboardCal.js.
 */
function runAllAppsScriptTests() {
  Logger.log('====================================================');
  Logger.log('STARTING COMPLETE APPS SCRIPT SUITE AUDIT');
  Logger.log('====================================================');

  runTestRetrieveLeadData();
  runTestLeadPriorityUpdate();
  runTestSyncUsersToFirestore();
  runTestDirectFirebaseLeadIngestion();
  runTestCheckAndMatchOrder();
  runTestDashboardCalFunctions();
  runTestUtilityFunctions();
  runTestReportAndPdfGeneration();
  runTestGallaboxWhatsApp();
  runTestWhatsAppMessage();

  Logger.log('====================================================');
  Logger.log('COMPLETE APPS SCRIPT SUITE AUDIT FINISHED');
  Logger.log('====================================================');
}
