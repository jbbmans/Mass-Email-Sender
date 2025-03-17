function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Mass Email Tool")
    .addItem("Open Dashboard", "showEmailDashboard")
    .addToUi();
}

function showEmailDashboard() {
  var html = HtmlService.createTemplateFromFile("Dashboard").evaluate()
    .setWidth(1100)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, "Mass Email Tool");
}

function getAvailableTokens() {
  return [
    { token: "{OFFICE}", description: "Office" },
    { token: "{CITY}", description: "City" },
    { token: "{FIRST_NAME}", description: "First Name" },
    { token: "{LAST_NAME}", description: "Last Name" },
    { token: "{EMAIL}", description: "Email Address" },
    { token: "{PHONE_NUMBER}", description: "Phone Number" },
    { token: "{ADDRESS}", description: "Address" }
  ];
}

function getUploadUrls(fileCount) {
  var arr = [];
  for (var i = 0; i < fileCount; i++){
    arr.push("FAKE_FILE_ID_" + i);
  }
  return arr;
}

function processDriveLinks(links) {
  var result = { fileIds: [], errors: [] };
  for (var i = 0; i < links.length; i++) {
    var link = links[i];
    if (link.indexOf("drive.google.com") !== -1) {
      result.fileIds.push({ id: "DRIVE_FILE_ID_" + i, name: "DriveFile_" + i });
    } else {
      result.errors.push("Invalid link: " + link);
    }
  }
  return result;
}

function getRecipients() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Contacts");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var recipients = [];
  var processedEmails = {};
  // Loop starting at row 1 (assuming row 0 is header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // Skip completely blank rows.
    if (row.join("").trim() === "") continue;
    var email = row[5].toString().trim();
    // If the email cell equals the header "EMAIL ADDRESS", skip this row.
    if (email.toUpperCase() === "EMAIL ADDRESS") continue;
    if (row[7] === true && email !== "") {
      if (processedEmails[email]) continue;
      processedEmails[email] = true;
      recipients.push(email);
    }
  }
  return recipients;
}

function sendMassEmails(emailConfig) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var contactsSheet = ss.getSheetByName("Contacts");
    if (!contactsSheet) throw new Error("Contacts sheet not found.");
    var data = contactsSheet.getDataRange().getValues();
    if (data.length < 2) throw new Error("No data in Contacts sheet.");
    var sentSheet = ss.getSheetByName("Emails Sent");
    if (!sentSheet) throw new Error("Emails Sent sheet not found.");
    
    var subject   = emailConfig.subject;
    var bodyHtml  = emailConfig.body;
    var signature = emailConfig.signature;
    var ccList    = emailConfig.ccAddresses || [];
    var bccList   = emailConfig.bccAddresses || [];
    var sentCount = 0;
    var processedEmails = {};
    
    // Fixed column order:
    // 0: OFFICE, 1: CITY, 2: FIRST NAME, 3: LAST NAME, 4: PHONE NUMBER,
    // 5: EMAIL ADDRESS, 6: ADDRESS, 7: SEND EMAIL
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row.join("").trim() === "") continue;
      
      var email = row[5].toString().trim();
      // Skip rows where the email cell is the header.
      if (email.toUpperCase() === "EMAIL ADDRESS") continue;
      
      if (row[7] === true && email !== "") {
        if (processedEmails[email]) continue;
        processedEmails[email] = true;
        
        var office    = row[0] || "";
        var city      = row[1] || "";
        var firstName = row[2] || "";
        var lastName  = row[3] || "";
        var phone     = row[4] || "";
        var address   = row[6] || "";
        
        var finalSubject = subject
          .replace("{OFFICE}", office)
          .replace("{CITY}", city)
          .replace("{FIRST_NAME}", firstName)
          .replace("{LAST_NAME}", lastName)
          .replace("{EMAIL}", email)
          .replace("{PHONE_NUMBER}", phone)
          .replace("{ADDRESS}", address);
          
        var finalBody = bodyHtml
          .replace("{OFFICE}", office)
          .replace("{CITY}", city)
          .replace("{FIRST_NAME}", firstName)
          .replace("{LAST_NAME}", lastName)
          .replace("{EMAIL}", email)
          .replace("{PHONE_NUMBER}", phone)
          .replace("{ADDRESS}", address);
          
        var finalSignature = signature
          .replace("{OFFICE}", office)
          .replace("{CITY}", city)
          .replace("{FIRST_NAME}", firstName)
          .replace("{LAST_NAME}", lastName)
          .replace("{EMAIL}", email)
          .replace("{PHONE_NUMBER}", phone)
          .replace("{ADDRESS}", address);
        
        var combined = finalBody + "<br><br>" + finalSignature;
        
        var mailOpts = { htmlBody: combined };
        if (ccList.length) mailOpts.cc = ccList.join(",");
        if (bccList.length) mailOpts.bcc = bccList.join(",");
        
        MailApp.sendEmail({
          to: email,
          subject: finalSubject,
          htmlBody: combined,
          cc: mailOpts.cc,
          bcc: mailOpts.bcc
        });
        
        var now = new Date();
        sentSheet.appendRow([
          Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss"),
          firstName + " " + lastName,
          email,
          "Sent"
        ]);
        sentCount++;
      }
    }
    return { success: true, message: "Successfully sent " + sentCount + " emails." };
  } catch (err) {
    throw new Error("An error occurred while sending emails.\n" + err.message);
  }
}
