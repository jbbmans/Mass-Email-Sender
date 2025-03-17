/**
 * Mass Email Tool for Google Apps Script
 * 
 * This script provides server-side functionality for a mass email tool
 * that integrates with Google Sheets to send personalized emails to multiple recipients.
 * 
 * @author Claude
 * @version 1.0.0
 */

/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mass Email Tool')
    .addItem('Open Dashboard', 'showDashboard')
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addToUi();
}

/**
 * Displays the email dashboard as a web app within Google Sheets.
 */
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Mass Email Tool')
    .setWidth(900)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Mass Email Tool');
}

/**
 * Displays the settings panel for configuring the email tool.
 */
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setTitle('Email Tool Settings')
    .setWidth(600)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

// ========================== API ENDPOINTS ==========================

/**
 * Sends mass emails to recipients with personalized content.
 * 
 * @param {Object} options - Email options and content
 * @param {string} options.subject - Email subject
 * @param {string} options.body - Email body (HTML)
 * @param {Array<string>} options.to - Array of primary recipient emails
 * @param {Array<string>} options.cc - Array of CC recipient emails
 * @param {Array<string>} options.bcc - Array of BCC recipient emails
 * @param {Array<Object>} options.attachments - Array of attachment objects
 * @param {Object} options.tokenData - Mapping of tokens to recipient data
 * @returns {Object} Status and results of the email sending operation
 */
function sendMassEmails(options) {
  try {
    Logger.log("Starting mass email operation");
    Logger.log("Recipients count: " + options.to.length);
    
    const emailService = new EmailService();
    const results = emailService.sendBulkEmails(options);
    
    return {
      success: true,
      sent: results.sent,
      failed: results.failed,
      errors: results.errors
    };
  } catch (error) {
    Logger.log("Error in sendMassEmails: " + error.toString());
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

/**
 * Uploads a file to Google Drive and returns its URL for embedding.
 * 
 * @param {Object} file - File object with base64 content
 * @param {string} file.name - File name
 * @param {string} file.mimeType - File MIME type
 * @param {string} file.content - Base64 encoded file content
 * @returns {Object} File information including URL
 */
function uploadFileToDrive(file) {
  try {
    const driveService = new DriveService();
    const uploadedFile = driveService.uploadFile(file);
    
    return {
      success: true,
      fileId: uploadedFile.fileId,
      url: uploadedFile.url,
      name: uploadedFile.name
    };
  } catch (error) {
    Logger.log("Error uploading file: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Retrieves recipient data from the specified Google Sheet.
 * 
 * @param {Object} options - Options for retrieving data
 * @param {string} options.sheetName - Name of the sheet to read from
 * @param {Array<string>} options.columns - Column names to retrieve
 * @returns {Object} Recipient data
 */
function getRecipientData(options) {
  try {
    const sheetService = new SheetService();
    const data = sheetService.getSheetData(options.sheetName, options.columns);
    
    return {
      success: true,
      data: data
    };
  } catch (error) {
    Logger.log("Error retrieving recipient data: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Gets the user's email signature from their settings.
 * 
 * @returns {Object} User's email signature
 */
function getUserSignature() {
  try {
    const userService = new UserService();
    const signature = userService.getEmailSignature();
    
    return {
      success: true,
      signature: signature
    };
  } catch (error) {
    Logger.log("Error retrieving user signature: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Saves user settings to Properties Service.
 * 
 * @param {Object} settings - User settings to save
 * @returns {Object} Status of the save operation
 */
function saveUserSettings(settings) {
  try {
    const userService = new UserService();
    userService.saveSettings(settings);
    
    return {
      success: true
    };
  } catch (error) {
    Logger.log("Error saving user settings: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ========================== SERVICE CLASSES ==========================

/**
 * Service for handling email operations.
 */
class EmailService {
  /**
   * Sends emails to multiple recipients with personalized content.
   * 
   * @param {Object} options - Email sending options
   * @returns {Object} Results of the send operation
   */
  sendBulkEmails(options) {
    const results = {
      sent: [],
      failed: [],
      errors: {}
    };
    
    // Process each recipient
    for (const recipient of options.to) {
      try {
        // Replace tokens in subject and body for this recipient
        const personalizedSubject = this._replaceTokens(options.subject, options.tokenData, recipient);
        const personalizedBody = this._replaceTokens(options.body, options.tokenData, recipient);
        
        // Send the email
        this._sendEmail({
          to: recipient,
          cc: options.cc || [],
          bcc: options.bcc || [],
          subject: personalizedSubject,
          body: personalizedBody,
          attachments: options.attachments || []
        });
        
        results.sent.push(recipient);
      } catch (error) {
        results.failed.push(recipient);
        results.errors[recipient] = error.toString();
        Logger.log("Error sending to " + recipient + ": " + error.toString());
      }
    }
    
    return results;
  }
  
  /**
   * Sends a single email with the provided options.
   * 
   * @param {Object} emailOptions - Options for this email
   * @private
   */
  _sendEmail(emailOptions) {
    // Prepare GmailApp options
    const options = {
      cc: emailOptions.cc.join(','),
      bcc: emailOptions.bcc.join(','),
      htmlBody: emailOptions.body,
      attachments: this._prepareAttachments(emailOptions.attachments)
    };
    
    // Send the email
    GmailApp.sendEmail(
      emailOptions.to,
      emailOptions.subject,
      "Your email client doesn't support HTML. Please use a modern email client to view this message.",
      options
    );
  }
  
  /**
   * Replaces tokens in text with recipient-specific data.
   * 
   * @param {string} text - Text containing tokens
   * @param {Object} tokenData - Mapping of tokens to values
   * @param {string} recipient - Recipient email
   * @returns {string} Text with tokens replaced
   * @private
   */
  _replaceTokens(text, tokenData, recipient) {
    if (!text || !tokenData || !tokenData[recipient]) {
      return text;
    }
    
    let result = text;
    const recipientData = tokenData[recipient];
    
    // Replace each token with its value
    for (const [token, value] of Object.entries(recipientData)) {
      // Create a regex that preserves formatting by looking for the token regardless of HTML tags
      const tokenRegex = new RegExp(`(\\{${token}\\})`, 'gi');
      result = result.replace(tokenRegex, value);
    }
    
    return result;
  }
  
  /**
   * Prepares attachment objects for GmailApp.
   * 
   * @param {Array<Object>} attachments - Array of attachment objects
   * @returns {Array<Object>} Prepared attachments for GmailApp
   * @private
   */
  _prepareAttachments(attachments) {
    if (!attachments || !attachments.length) {
      return [];
    }
    
    return attachments.map(attachment => {
      if (attachment.blob) {
        return attachment.blob;
      }
      
      if (attachment.fileId) {
        return DriveApp.getFileById(attachment.fileId).getBlob();
      }
      
      throw new Error("Invalid attachment format");
    });
  }
}

/**
 * Service for handling Google Drive operations.
 */
class DriveService {
  /**
   * Uploads a file to Google Drive.
   * 
   * @param {Object} file - File object with base64 content
   * @returns {Object} Uploaded file information
   */
  uploadFile(file) {
    // Decode base64 content
    const blob = this._base64ToBlob(file.content, file.mimeType, file.name);
    
    // Create file in Drive
    const folder = this._getOrCreateUploadFolder();
    const driveFile = folder.createFile(blob);
    
    // Set sharing permissions
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      fileId: driveFile.getId(),
      url: driveFile.getUrl(),
      name: driveFile.getName(),
      mimeType: driveFile.getMimeType()
    };
  }
  
  /**
   * Gets or creates a folder for uploading files.
   * 
   * @returns {Object} Drive folder
   * @private
   */
  _getOrCreateUploadFolder() {
    const folderName = "Mass Email Tool Uploads";
    
    // Try to find existing folder
    const folderIterator = DriveApp.getFoldersByName(folderName);
    
    if (folderIterator.hasNext()) {
      return folderIterator.next();
    }
    
    // Create new folder if not found
    return DriveApp.createFolder(folderName);
  }
  
  /**
   * Converts base64 string to Blob.
   * 
   * @param {string} base64 - Base64 encoded string
   * @param {string} mimeType - MIME type of the file
   * @param {string} fileName - Name of the file
   * @returns {Blob} File blob
   * @private
   */
  _base64ToBlob(base64, mimeType, fileName) {
    const decoded = Utilities.base64Decode(base64);
    return Utilities.newBlob(decoded, mimeType, fileName);
  }
}

/**
 * Service for handling Google Sheets operations.
 */
class SheetService {
  /**
   * Gets data from a specific sheet.
   * 
   * @param {string} sheetName - Name of the sheet
   * @param {Array<string>} columns - Column names to retrieve
   * @returns {Object} Sheet data
   */
  getSheetData(sheetName, columns) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    
    // Get all data
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      throw new Error("Sheet has no data or headers");
    }
    
    // First row is headers
    const headers = data[0];
    
    // Map column names to indices
    const columnIndices = columns.map(column => {
      const index = headers.indexOf(column);
      if (index === -1) {
        throw new Error(`Column "${column}" not found`);
      }
      return index;
    });
    
    // Extract data for requested columns
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowData = {};
      
      for (let j = 0; j < columns.length; j++) {
        rowData[columns[j]] = row[columnIndices[j]];
      }
      
      result.push(rowData);
    }
    
    return result;
  }
}

/**
 * Service for handling user settings and preferences.
 */
class UserService {
  /**
   * Gets the user's email signature.
   * 
   * @returns {string} HTML email signature
   */
  getEmailSignature() {
    const userProperties = PropertiesService.getUserProperties();
    const signature = userProperties.getProperty('emailSignature');
    
    if (!signature) {
      // Return default signature
      const user = Session.getActiveUser().getEmail();
      return `<p>Best regards,</p><p>${user}</p>`;
    }
    
    return signature;
  }
  
  /**
   * Saves user settings.
   * 
   * @param {Object} settings - Settings to save
   */
  saveSettings(settings) {
    const userProperties = PropertiesService.getUserProperties();
    
    for (const [key, value] of Object.entries(settings)) {
      userProperties.setProperty(key, value.toString());
    }
  }
}

// ========================== GLOBAL ERROR HANDLER ==========================

/**
 * Global error handler for client-side errors.
 * 
 * @param {Error} error - Error object
 * @returns {Object} Error information
 */
function reportError(error) {
  Logger.log("Client error reported: " + error.toString());
  
  return {
    message: error.message || error.toString(),
    stack: error.stack,
    timestamp: new Date().toISOString()
  };
}
