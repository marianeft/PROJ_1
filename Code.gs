// MO GSHEETS TOOL
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("MO Tools")
    .addItem("Search MO by Number", "searchMO")
    .addItem("Search MO by Subject", "searchMOSubject")
    .addToUi();
}

//SEARCH FUNCTIONS
// Search by Control Number
function searchMO() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt("Enter MO Control Number (e.g., MO-2026-ADMIN-2):");
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
    var data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();

    for (var i=0; i<data.length; i++) {
      if (data[i][1] == response.getResponseText()) { // Column B = MO Number
        ui.alert(
          "MO Found:\n\n" +
          "Control Number: " + data[i][1] + "\n" +
          "Assignment Timestamp: " + data[i][2] + "\n" +
          "Subject: " + data[i][4] + "\n" +
          "Date: " + data[i][5] + "\n" +
          "Department: " + data[i][6] + "\n" +
          "Signatory: " + data[i][7] + "\n" +
          "Notes: " + data[i][8]
        );
        return;
      }
    }
    ui.alert("MO not found.");
  } catch (err) {
    Logger.log("UI not available in this context: " + err.message);
  }
}

// Search by Subject keyword
function searchMOSubject() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Enter MO Subject keyword:");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  var results = [];

  for (var i=0; i<data.length; i++) {
    if (data[i][4] && data[i][4].toLowerCase().includes(response.getResponseText().toLowerCase())) {
      results.push("MO Number: " + data[i][1] + " | Subject: " + data[i][4] + " | Department: " + data[i][6]);
    }
  }

  if (results.length > 0) {
    ui.alert("Results:\n\n" + results.join("\n"));
  } else {
    ui.alert("No MO found with that subject.");
  }
}

// FORM SUBMISSION AND EMAIL NOTIFICATION

// Generate Next MO Number supporting multiple memo types
function getNextMONumber(memoType) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var year = Utilities.formatDate(new Date(), "Asia/Manila", "yyyy")
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues(); // Columns B to J
  var usedSeqs = [];

  // Collect all used sequence numbers for the current year
  for (var i = 0; i < data.length; i++) {
    var val = data[i][1]; // Column B = MO Number
    var type = data[i][8]; // Column I = Memo Type
    if (val && abbrevMemoType(type) === memoType) {
      var match = val.toString().match(/^MO-(\d{4})-(\w+)-(\d+)$/);
      if (match && match[1] === year && match[2] === memoType) {
        usedSeqs.push(parseInt(match[3], 10));
      }
    }
  }

  // Sort and find the lowest missing sequence
  usedSeqs.sort(function(a, b) { return a - b; });
  var nextSeq = 1;
  for (var j = 0; j < usedSeqs.length; j++) {
    if (usedSeqs[j] === nextSeq) {
      nextSeq++;
    } else {
      break;
    }
  }

  return "MO-" + year + "-" + memoType + "-" + nextSeq;
}

  // Abbreviate Memo Types
  function abbrevMemoType(fullType){
    var map = {
      "Regional Administrative Order": "RAO",
      "Regional Office Order": "ROO",
      "Regional Special Order": "RSO",
    };
    return map[fullType] || fullType.replace(/\s+/g,"").toUpperCase(); //fallback
  }

// FORM SUBMISSION HANDLER
function onFormSubmit(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
    var lastRow = sheet.getLastRow();

    // Read Memo Type from form Memo Type (from Column J)
    var fullType = sheet.getRange(lastRow, 9).getValue(); // Column J
    if (!fullType || fullType.trim() === "") {
      throw new Error("Memo Type is missing. Cannot generate control number.");
    }
    var memoType = abbrevMemoType(fullType);
    var moNumber = getNextMONumber(memoType);
    sheet.getRange(lastRow, 2).setValue(moNumber); // Column B

    // Add Assignment Timestamp
    var timestamp = Utilities.formatDate(new Date(), "Asia/Manila", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(lastRow, 3).setValue(timestamp); // Column C

    // Collect other details
    var email = sheet.getRange(lastRow, 4).getValue(); // Email Address (Column D)
    var subjectText = sheet.getRange(lastRow, 5).getValue(); // Subject (Column E)
    var dateIssued = sheet.getRange(lastRow, 6).getValue(); // Date Issue (Column F)
    var department = sheet.getRange(lastRow, 7).getValue(); // Department (Column G)
    var notes = sheet.getRange(lastRow, 8).getValue(); // Notes (Column H)


  // HTML Email Body
   var htmlBody = `
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">

        <h2 style="color:#2E86C1;">
          Memorandum Order Notification
        </h2>

        <p>Good day,</p>

        <p>
          Please be informed that a <strong>Memorandum Order</strong> has been officially issued.
          The details of the memorandum are provided below for your reference:
        </p>
        
        <table style="border-collapse: collapse; width: 80%; margin-top: 10px;">
          <tr style="background-color:#f2f2f2;">
            <td style="border:1px solid #ddd; padding:8px;"><strong>Control Number</strong></td>
            <td style="border:1px solid #ddd; padding:8px;">${moNumber}</td>
          </tr>
          <tr>
            <td style="border:1px solid #ddd; padding:8px;"><strong>Assignment Timestamp</strong></td>
            <td style="border:1px solid #ddd; padding:8px;">${timestamp}</td>
          </tr>
          <tr style="background-color:#f9f9f9;">
            <td style="border:1px solid #ddd; padding:8px;"><strong>Subject</strong></td>
            <td style="border:1px solid #ddd; padding:8px;">${subjectText}</td>
          </tr>
          <tr style="background-color:#f9f9f9;">
            <td style="border:1px solid #ddd; padding:8px;"><strong>MO Type</strong></td>
            <td style="border:1px solid #ddd; padding:8px;">${fullType}</td>
          </tr>
          <tr>
            <td style="border:1px solid #ddd; padding:8px;"><strong>Date Issued</strong></td>
            <td style="border:1px solid #ddd; padding:8px;">${dateIssued}</td>
          </tr>
          <tr style="background-color:#f9f9f9;">
            <td style="border:1px solid #ddd; padding:8px;"><strong>Department</strong></td>
            <td style="border:1px solid #ddd; padding:8px;">${department}</td>
          </tr>
          <tr style="background-color:#f9f9f9;">
            <td style="border:1px solid #ddd; padding:8px;"><strong>Notes</strong></td>
            <td style="border:1px solid #ddd; padding:8px;">${notes}</td>
          </tr>
        </table>
        
        <hr style="margin: 25px 0; border: none; border-top:1px solid #ccc">

        <h3 style="color:#1F4E79; margin-bottom:10px;">
          Important Notice
        </h3>

        <p style="margin-top:5px;">
          Please be advised of the following policies regarding Memorandum Orders:
        </p>

        <ul style="margin-left: 20px;">
          <li> Once a Memorandum Order ahs been submitted and is issued it <strong> cannot be cannot be cancelled, modified, or resubmitted </strong> through the system.
          </li>
          <li> If an error is identified or a  revision is requiered, the sender must <strong>formally submit a request</strong> to the system administrator or the authorized office for evaluation and approval.
          </li>
        </ul>
          
        <p><i>Unauthorized resubmission of duplication of Memorandum Order is not permitted and may result in administrative review.
        </i></p>

        <p style="margin-top:20px;">
          An official PDF copy of the Memorandum Order is attached to this email.
          This message serves as formal notification.
        </p>

        <p>
          For your guidance and appropriate action.
        </p>


        <p style="margin-top:25px;">
          Respectfully,<br>
          <strong>Memorandum Automation System</strong>
        </p>

      </body>
    </html>
    `;

    // Create PDF content
    var doc = DocumentApp.create("Memorandum Order " + moNumber);
    var body = doc.getBody();
    body.appendParagraph("Memorandum Order").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Control Number: " + moNumber);
    body.appendParagraph("Date Issued: " + dateIssued);
    body.appendParagraph("Department/Office: " + department);
    body.appendParagraph("Subject: " + subjectText);
    body.appendParagraph("Memorandum Type: " + fullType);
    body.appendParagraph("Notes: " + notes);
    body.appendParagraph("Assignment Timestamp: " + timestamp);

    doc.saveAndClose();

    // Save as PDF in Drive
    var pdfFile = DriveApp.getFileById(doc.getId()).getAs("application/pdf");
    var folder = DriveApp.getFolderById("17IPOUCC6tnEJNkKPzLSk30AmlOHyCg_k");
    var savedFile = folder.createFile(pdfFile).setName("MO_" + moNumber + ".pdf");

    // Send Email
    MailApp.sendEmail({
      to: email,
      subject: "Memorandum Order " + moNumber,
      attachments: [savedFile],
      htmlBody: htmlBody
    });

  } catch (err) {
    Logger.log(err.message);
  }
}