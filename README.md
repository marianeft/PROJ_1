# Memorandum Order Automation (Google Sheets + Apps Script)

This project automates the management of **Memorandum Orders (MOs)** using Google Sheets and Google Apps Script.  
It provides tools for searching, numbering, PDF generation, and email notifications — streamlining the workflow for records offices and administrative staff.

---

## ✨ Features

- **Custom Menu in Google Sheets**
  - Search MO by Control Number
  - Search MO by Subject Keyword

- **Automated Control Numbering**
  - Generates sequential MO numbers per year and memo type
  - Supports multiple memo types with abbreviations:
    - Regional Administrative Order → `RAO`
    - Regional Office Order → `ROO`
    - Regional Special Order → `RSO`

- **Form Submission Handler**
  - Automatically assigns MO Number and timestamp on form submission
  - Generates a PDF copy of the MO with details
  - Stores the PDF in Google Drive
  - Sends an email notification with the PDF attached

- **PDF Output**
  - Includes MO details in a structured format
  - Saved in a designated Drive folder

---

## 📊 Spreadsheet Structure

The script expects a Google Sheet with the following columns:

| Column | Field                |
|--------|----------------------|
| A      | Timestamp (Form submission) |
| B      | MO Number            |
| C      | Assignment Timestamp |
| D      | Email Address        |
| E      | MO Subject           |
| F      | Date Issued          |
| G      | Department/Office    |
| H      | Notes                |
| I      | Memo Type            |

---

## ⚙️ Setup Instructions

1. Open your Google Sheet (with form responses).
2. Go to **Extensions → Apps Script**.
3. Copy the contents of `Code.gs` into the editor.
4. Update the **Google Drive Folder ID** in the script:
   ```javascript
   var folder = DriveApp.getFolderById("YOUR_FOLDER_ID_HERE");
   ```
5. Save the script and refresh the sheet.
6. A new menu **“MO Tools”** will appear with search options.
7. Link your Google Form to the sheet for automatic submissions.

---

## 📧 Email Notification

When a form is submitted:
- The system generates a control number.
- A PDF copy of the MO is created and stored in Drive.
- An email is sent to the provided address with the PDF attached.

---

## 🛠️ Example Control Numbers

- `MO-2026-RAO-1`
- `MO-2026-ROO-2`
- `MO-2026-RSO-3`

---

## 🚀 Future Enhancements

- Reindex tool to fix past MO numbering gaps
- Dashboard for quick MO tracking
- Error handling for missing memo types
- Improved PDF styling with headers/footers

---

## 📜 License

This project is licensed under the MIT License — feel free to use, modify, and distribute with attribution.

---

DEVELOPED FOR 💙 by MFT