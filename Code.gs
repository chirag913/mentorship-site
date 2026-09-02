function doPost(e) {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Sheet1");

  const data = JSON.parse(e.postData.contents);

  // Basic bot trap: real visitors never fill this hidden field.
  if (data.website) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  sheet.appendRow([
    new Date(),
    data.name,
    data.email,
    data.whatsapp,
    data.income,
    data.budget,
    data.goal,
    data.struggle,
    data.market,
    data.experience,
    data.interested_in
  ]);

  return ContentService
    .createTextOutput(
      JSON.stringify({
        success: true
      })
    )
    .setMimeType(ContentService.MimeType.JSON);
}

/* ====================================================
   MENTORSHIP LEAD EMAIL NOTIFIER
   ========================================================= */

// ---------- CONFIG ----------
const LEAD_SHEET_NAME = "Sheet1";
const NOTIFY_EMAIL = "sharma.chirag913@gmail.com";
const LAST_ROW_PROPERTY_KEY = "LEAD_TRACKER_LAST_ROW";
const SHEET_URL = "https://docs.google.com/spreadsheets/d/1YeSGHoSOd_7Yhhaun-Oswk6NVQE0Bo3AZ-ks2504wdw/edit";
const NUM_COLUMNS = 11; // A:Timestamp B:Name C:Email D:WhatsApp E:Income F:Budget G:Goal H:Struggle I:Market J:Experience K:Interested In

function getLeadSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEAD_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + LEAD_SHEET_NAME + '" was not found in this spreadsheet.');
  }
  return sheet;
}

function initializeLeadTracker() {
  const sheet = getLeadSheet_();
  const lastRow = sheet.getLastRow();
  PropertiesService.getScriptProperties().setProperty(LAST_ROW_PROPERTY_KEY, String(lastRow));
  Logger.log(
    "Lead tracker initialized. Last row recorded as " + lastRow +
    ". Only rows added after this will trigger emails."
  );
}

function checkForNewLeads() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(LAST_ROW_PROPERTY_KEY);

  if (stored === null) {
    initializeLeadTracker();
    return;
  }

  const lastProcessedRow = Number(stored);
  const sheet = getLeadSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow <= lastProcessedRow) {
    return;
  }

  const numNewRows = lastRow - lastProcessedRow;
  const range = sheet.getRange(lastProcessedRow + 1, 1, numNewRows, NUM_COLUMNS);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    const rowNumber = lastProcessedRow + 1 + i;
    const row = values[i];

    const isEmptyRow = row.every(function (cell) {
      return cell === "" || cell === null || cell === undefined;
    });

    if (!isEmptyRow) {
      const lead = {
        timestamp: row[0],
        name: row[1],
        email: row[2],
        whatsapp: row[3],
        income: row[4],
        budget: row[5],
        goal: row[6],
        struggle: row[7],
        market: row[8],
        experience: row[9],
        interestedIn: row[10]
      };
      sendLeadEmail_(lead);
    }

    props.setProperty(LAST_ROW_PROPERTY_KEY, String(rowNumber));
  }
}

function sendLeadEmail_(lead) {
  const name = lead.name || "Unknown";
  const subject = "\uD83D\uDD25 New " + (lead.interestedIn || "Lead") + " — " + name;

  const body =
    "\uD83D\uDD25 NEW " + (lead.interestedIn || "LEAD").toUpperCase() + "\n\n" +
    "Name: " + lead.name + "\n" +
    "Email: " + lead.email + "\n" +
    "WhatsApp: " + lead.whatsapp + "\n" +
    "Interested in: " + lead.interestedIn + "\n" +
    "Market: " + lead.market + "\n" +
    "Experience: " + lead.experience + "\n" +
    "Income: " + lead.income + "\n" +
    "Budget: " + lead.budget + "\n" +
    "Goal: " + lead.goal + "\n" +
    "Struggle: " + lead.struggle + "\n\n" +
    "Timestamp: " + formatTimestamp_(lead.timestamp) + "\n\n" +
    "View in Google Sheet: " + SHEET_URL;

  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

function formatTimestamp_(ts) {
  if (Object.prototype.toString.call(ts) === "[object Date]") {
    return Utilities.formatDate(ts, Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm a");
  }
  return String(ts);
}
