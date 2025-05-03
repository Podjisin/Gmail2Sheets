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

function extractEmailsToSheet(startDate, endDate, sortOrder) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const query = `after:${formatDateForQuery(start)} before:${formatDateForQuery(end)}`;
  const threads = GmailApp.search(query);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.getRange("A1").setValue("Status:")
  sheet.getRange("B1").setValue("RUNNING");
  sheet.appendRow(["From", "Date", "Subject"]);

  const emailData = [];

  for (const thread of threads) {
    if (sheet.getRange("B1").getValue() === "STOP") {
      Logger.log("Process stopped by user.");
      sheet.appendRow(["--- Script stopped manually ---", "", ""]);
      break;
    }

    const messages = thread.getMessages();
    for (const message of messages) {
      if (sheet.getRange("B1").getValue() === "STOP") {
        Logger.log("Process stopped by user.");
        sheet.appendRow(["--- Script stopped manually ---", "", ""]);
        return;
      }

      // Extract only the email address from "From"
      const fromEmail = extractEmail(message.getFrom());

      // Format the message date
      const formattedDate = formatDateForDisplay(message.getDate());

      emailData.push([
        fromEmail,  // Just the email address
        formattedDate,
        message.getSubject()
      ]);
    }
  }

  // Sort emails by date (column index 1 is Date)
  emailData.sort((a, b) => {
    const dateA = new Date(a[1]);
    const dateB = new Date(b[1]);
    return sortOrder === 'asc' ? dateA - dateB : dateB - dateA;
  });

  // Write sorted data to sheet
  for (const row of emailData) {
    sheet.appendRow(row);
  }

  sheet.getRange("B1").setValue("DONE");
}

function extractEmail(fromString) {
  const match = fromString.match(/<([^>]+)>/);  // Match content inside angle brackets
  return match ? match[1] : fromString;  // Return the email inside the brackets, or the original string if no match
}

function formatDateForDisplay(date) {
  const options = { 
    weekday: 'short', year: 'numeric', month: 'short', day: 'numeric', 
    hour: 'numeric', minute: 'numeric', hour12: true 
  };
    // TODO: Turn this into a parameter.
  return date.toLocaleDateString('en-US', options);
}

function stopScript() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("B1").setValue("STOP");
}

function formatDateForQuery(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}
