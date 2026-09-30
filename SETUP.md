# Setup

## Requirements

- Google account
- Google Spreadsheet
- Google Drive folder
- Google Apps Script

## 1. Create a spreadsheet

Create a new Google Spreadsheet and copy the spreadsheet ID from:

`https://docs.google.com/spreadsheets/d/SHEET_ID/edit`

## 2. Create a Drive folder

Create a folder for saved PDFs and copy the folder ID from:

`https://drive.google.com/drive/folders/FOLDER_ID`

## 3. Create an Apps Script project

Create a standalone Apps Script project and add:

- `Code.gs`
- `appsscript.json`

## 4. Configure Script Properties

In **Project Settings → Script Properties**, add:

- `SHEET_ID`
- `FOLDER_ID`

Optional:

- `PROCESSED_LABEL` (default: `portfolio/processed`)
- `MAX_THREADS` (default: `50`)

Do not hard-code private IDs into source code.

## 5. Initialize

Run:

`setupSheets()`

It creates:

- `Records`
- `Preview`
- `Errors`

## 6. Test with synthetic data

Send yourself a test email.

Subject:

`Invoice - Demo Vendor`

Body:

```text
Demo invoice for portfolio testing.

Grand Total: USD 123.45

This message contains synthetic test data only.
```

Attach a harmless PDF file.

## 7. Preview

Run:

`previewInvoiceEmails()`

Confirm the message appears in the `Preview` sheet.

## 8. Process

Run:

`processInvoiceEmails()`

Confirm:

- a row is added to `Records`
- the PDF is saved to Drive
- the Gmail thread receives the processed label

Run `processInvoiceEmails()` again and confirm that the same message is not duplicated.
