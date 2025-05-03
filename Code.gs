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

function extractEmailsToSheet(startDate, endDate, sortOrder, dateTimeFormat) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const query = `after:${formatDateForQuery(start)} before:${formatDateForQuery(end)}`;
  // Uncomment this query if you want to filter promotions, social, and forums.
  // const query = `after:${formatDateForQuery(start)} before:${formatDateForQuery(end)} -category:promotions -category:social -category:forums is:important`;
  const threads = GmailApp.search(query);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();

  // Style and set status label
  const labelCell = sheet.getRange("A1");
  labelCell.setValue("Status:")
    .setFontWeight("bold")
    .setBackground("#e0e0e0")  // light gray
    .setFontColor("black");

  // Style and set status RUNNING
  const statusCell = sheet.getRange("B1");
  statusCell.setValue("RUNNING")
    .setFontWeight("bold")
    .setFontColor("white")
    .setBackground("#4285F4");  // Google blue

  sheet.appendRow(["From", "Date", "Subject"]);

  const lastRow = sheet.getLastRow();
  const headerRange = sheet.getRange(lastRow, 1, 1, 3);
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

    const messages = thread.getMessages();
    for (const message of messages) {
      if (sheet.getRange("B1").getValue() === "STOP") {
        Logger.log("Process stopped by user.");
        sheet.appendRow(["--- Script stopped manually ---", "", ""]);
        return;
      }

      const fromEmail = extractEmail(message.getFrom());
      const formattedDate = formatDateForDisplay(message.getDate(), dateTimeFormat);

      // Attempt to skip promotional-looking or scammy messages
      // Uncomment and tweak as you want.
      if (/noreply|no-reply|offers|marketing|promo/i.test(fromEmail)) continue;
      if (/free money|claim now|urgent action required/i.test(subject)) continue;

      emailData.push([fromEmail, formattedDate, message.getSubject()]);
    }
  }

  emailData.sort((a, b) => {
    const dateA = new Date(a[1]);
    const dateB = new Date(b[1]);
    return sortOrder === 'asc' ? dateA - dateB : dateB - dateA;
  });

  for (const row of emailData) {
    sheet.appendRow(row);
  }

  // Style and set status DONE
  sheet.getRange("B1")
    .setValue("DONE")
    .setFontWeight("bold")
    .setFontColor("white")
    .setBackground("#34A853");  // Google green
}


function extractEmail(fromString) {
  const match = fromString.match(/<([^>]+)>/);  // Match content inside angle brackets
  return match ? match[1] : fromString;  // Return the email inside the brackets, or the original string if no match
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
