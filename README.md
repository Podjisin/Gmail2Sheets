# Email Extractor Google Sheets Script

This Google Sheets Script allows you to extract email information from Gmail based on a specific date range and sort the results. The extracted data includes the sender's email, the date the email was sent, and the subject. The emails are sorted according to your preferred date order (ascending or descending).

## Features

* Extract email information from Gmail based on a date range.
* Sort emails by date (ascending or descending).
* Display extracted data in a Google Sheet.
* Stop the script manually if needed. (sometimes don't work lol)

## Prerequisites

* A Google account.
* Access to Google Sheets.
* Gmail account connected to your Google account.
* Tinie tiny technical knowledge in google environment. Don't worry! This is very easy!

## Installation

### Step 1: Open Google Sheets

1. Open the Google Sheet where you want to use the Email Extractor.

### Step 2: Open the Script Editor

1. In Google Sheets, go to the **Extensions** menu.
2. Select **Apps Script** from the dropdown.

### Step 3: Add the Script

1. In the Apps Script editor, replace the default code in the `Code.gs` file with the **Code.gs** content from this repository.
2. Click on `File > New > HTML` and name the file `Sidebar`.
3. Copy and paste the **Sidebar.html** content into the newly created `Sidebar.html` file.

### Step 4: Save and Close

1. After pasting the code, click on the **Save** button (disk icon) in the upper left corner of the Apps Script editor.
2. You can name the project (e.g., `Email Extractor`).
3. Don't forget to deploy it as Web App! `Deploy > New Deployment > Select Type (Gear Icon)`.
4. You can now close the Apps Script editor or not.

### Step 5: Set Permissions

1. Go back to the Google Sheets window.
2. The script will not work until the proper permissions are granted.
3. Go to **Email Tools > Open GUI** to trigger the script.
4. If a permissions dialog appear. Review the permissions and click **Allow** to grant access to your Gmail account.

### Step 6: Using the Script

1. After allowing the necessary permissions, you will see a sidebar on the right side of your Google Sheet with options to:

   * Choose a start and end date.
   * Select a sorting order for the extracted emails (ascending or descending).
   * Click the **Start** button to begin extracting emails.
   * Click the **Stop** button to manually stop the script at any time.

### Step 7: Results

1. The extracted email information (sender email, date, and subject) will appear in the Google Sheet, sorted according to your selected preferences.
2. Once the extraction is complete, you will see the status as **DONE** in cell `A1` of your Google Sheet.

## Troubleshooting

* **Script Not Running:** Make sure you have granted all necessary permissions. If you see an error about permissions, try going back to **Email > Open GUi** and re-allowing access.
* **Emails Not Extracting:** Double-check the date range to ensure it is correct. The script only pulls emails within the specified range.
* **Sorting Not Working:** If sorting is not working, ensure that the `Sort By Date` dropdown in the sidebar is correctly set to either ascending or descending.

* **If none works**: Submit an issue! I'll try to help ASAP.

## How It Works

* **onOpen()**: This function adds a custom menu (`Email Tools`) to Google Sheets when the spreadsheet is opened. The menu allows you to trigger the extraction and stop the script.
* **showSidebar()**: This function opens the sidebar UI where you can input dates, select the sort order, and control the extraction process.
* **extractEmailsToSheet()**: This function fetches emails from Gmail based on the date range you specify and sorts them based on the selected order.
* **stopScript()**: This function allows you to manually stop the script if needed.

## Motivation

Email logging can be a painful and time-consuming task. Manually sifting through countless emails, copying and pasting key information like sender addresses, dates, and subjects into a spreadsheet is not only inefficient but also tedious. Whether you're tracking client communications, managing project emails, or simply trying to stay organized, the manual process is often overwhelming.

The Email Extractor script was created to automate this process, saving you hours of effort and frustration. With just a few clicks, this tool extracts email data from your Gmail account and logs it directly into a Google Sheet. No more copy-pasting, no more wasting time on repetitive tasks—just quick, organized, and accurate results at your fingertips.

## License

WTFPL License. See [LICENSE](LICENSE) for more details.