function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Email Tools')
    .addItem('Start Extraction', 'showSidebar')
    .addItem('Stop Script', 'stopScript')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Email Extractor');
  SpreadsheetApp.getUi().showSidebar(html);
}

function extractEmailsToSheet(startDate, endDate, sortOrder, dateTimeFormat, scope, label, whitelist, blacklist) {
  const start = new Date(startDate);
  const end = new Date(endDate);

  let baseQuery = `after:${formatDateForQuery(start)} before:${formatDateForQuery(end)}`;
  
  if (scope === "inbox") {
    baseQuery += " in:inbox";
  } else if (scope === "anywhere") {
    baseQuery += " in:anywhere -in:spam -in:trash";
  } else if (scope === "custom" && label) {
    baseQuery += ` label:${label}`;
  }

  const whitelistArray = whitelist ? whitelist.split(',').map(e => e.trim().toLowerCase()) : [];
  const blacklistArray = blacklist ? blacklist.split(',').map(e => e.trim().toLowerCase()) : [];

  const threads = GmailApp.search(baseQuery);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();

  // Status bar setup
  const labelCell = sheet.getRange("A1");
  labelCell.setValue("Status:")
    .setFontWeight("bold")
    .setBackground("#e0e0e0")
    .setFontColor("black");

  const statusCell = sheet.getRange("B1");
  statusCell.setValue("RUNNING")
    .setFontWeight("bold")
    .setFontColor("white")
    .setBackground("#4285F4");

  sheet.appendRow(["From", "Date", "Subject"]);
  const headerRange = sheet.getRange(sheet.getLastRow(), 1, 1, 3);
  headerRange
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#f1f3f4");

  const emailData = [];

  for (const thread of threads) {
  if (sheet.getRange("B1").getValue() === "STOP") {
    Logger.log("Process stopped by user.");
    sheet.appendRow(["--- Script stopped manually ---", "", ""]);
    return;
  }

  const messages = thread.getMessages().filter(msg => {
    const msgDate = msg.getDate();
    return msgDate >= start && msgDate <= end;
  });

  for (const message of messages) {
    if (sheet.getRange("B1").getValue() === "STOP") {
      Logger.log("Process stopped by user.");
      sheet.appendRow(["--- Script stopped manually ---", "", ""]);
      return;
    }

    const fromEmail = extractEmail(message.getFrom()).toLowerCase();
    const subject = message.getSubject();
    const formattedDate = formatDateForDisplay(message.getDate(), dateTimeFormat);

    if (whitelistArray.length > 0 && !whitelistArray.includes(fromEmail)) continue;
    if (blacklistArray.includes(fromEmail)) continue;

    emailData.push([fromEmail, formattedDate, subject]);
  }}

  emailData.sort((a, b) => {
    const dateA = new Date(a[1]);
    const dateB = new Date(b[1]);
    return sortOrder === 'asc' ? dateA - dateB : dateB - dateA;
  });

  for (const row of emailData) {
    sheet.appendRow(row);
  }

  // Mark status as DONE
  sheet.getRange("B1")
    .setValue("DONE")
    .setFontWeight("bold")
    .setFontColor("white")
    .setBackground("#34A853");
}

function extractEmail(fromString) {
  const match = fromString.match(/<([^>]+)>/);  // Match content inside angle brackets
  return match ? match[1] : fromString; 
}

function formatDateForDisplay(date, format) {
  if (!format || format === '' || format.length == 0) {
    format = 'ddd, YYYY-MM-DD HH:mm A'
  }
  var now = dayjs(date);
  Logger.log(now.format(format));
  return now.format(format)

}

function stopScript() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("B1")
    .setValue("STOP")
    .setFontWeight("bold")
    .setFontColor("white")
    .setBackground("#EA4335");  // Google red
}

function formatDateForQuery(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}
