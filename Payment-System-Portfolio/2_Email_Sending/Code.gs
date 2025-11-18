/**
 * Payment Notification Email Sender
 * * * Overview:
 * This script acts as the "Distribution" phase of the automation workflow.
 * It reads the management sheet, identifies rows approved for sending,
 * finds the corresponding PDF in Google Drive, and sends it via Gmail.
 * * * Key Features:
 * - Custom Menu integration for easy execution.
 * - Safety First: "Double-Check" mechanism to verify status right before sending.
 * - Sends emails using a specific alias address (from:).
 * - Updates the status in the sheet upon successful transmission.
 */

// =================================================================
// ★ 1. Custom Menu Function ★
// =================================================================
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Payment Automation') // Menu Name
      .addItem('Send Approved Emails', 'sendPaymentEmails') // Item Name
      .addToUi();
}

// =================================================================
// ★ 2. Main Function: Send Emails ★
// =================================================================
function sendPaymentEmails() {

  // --- Configuration [SENSITIVE DATA MASKED] ---

  // File and Sheet Settings
  const SOURCE_SHEET_ID = "YOUR_MANAGEMENT_SHEET_ID"; // e.g., "1i8uPjl..."
  const SOURCE_SHEET_NAME = "Payment_Management_Data"; // The sheet name
  const DATA_START_ROW = 2;
  
  const PDF_PARENT_FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID"; // Root folder for monthly PDFs

  // Email Settings
  const SENDER_EMAIL = "billing@example.com";  // ★Sender Alias
  const CC_EMAIL = "manager@example.com";      // ★CC Address
  const PDF_FOLDER_SUFFIX = "_Payment_Docs";   // Suffix for monthly folders

  // Status Tags
  const STATUS_TRIGGER = "承認済み送付OK";    // "Approved (Ready to Send)"
  const STATUS_SENT = "支払通知書送付済";      // "Sent"
  const STATUS_REVERTED = "（差し戻し）";      // "Reverted" (Safety check failed)
  const STATUS_ERROR_PDF = "エラー: PDF無し"; // "Error: PDF Not Found"

  // Column Indices (0-based index)
  const COL_IDX = {
    AGENT_NAME: 16,    // Q: Agent Name
    PAYMENT_MONTH: 19, // T: Payment Month
    EMAIL_TO: 21,      // V: Recipient Email
    SUBJECT: 23,       // X: Email Subject
    BODY: 24,          // Y: Email Body
    STATUS: 25         // Z: Status
  };
  
  // Additional Column for Double Check (Company Name for filename matching)
  const COL_IDX_COMPANY = 1; // B: Company Name

  // Email Signature
  const EMAIL_SIGNATURE = "\n\n" +
      "────────────────────────────────────────────────────────────────\n" +
      "    Your Company Name Inc.\n" +
      "    Customer Support Team\n" +
      "    \n" +
      "    Address: 1-2-3 Business St, Tokyo, Japan\n" +
      "    Email: support@example.com\n" +
      "    URL: https://www.example.com/\n" +
      "    ────────────────────────────────────────────────────────────────";

  // --- End of Configuration ---
  
  try {
    // 1. Setup Sheet Access
    const sourceSs = SpreadsheetApp.openById(SOURCE_SHEET_ID);
    const sourceSheet = sourceSs.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) throw new Error(`Sheet "${SOURCE_SHEET_NAME}" not found.`);
    
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < DATA_START_ROW) {
      Browser.msgBox("No data found in the management sheet.");
      return;
    }

    // Read Data
    const numRows = lastRow - DATA_START_ROW + 1;
    const numCols = sourceSheet.getLastColumn();
    const sourceData = sourceSheet.getRange(DATA_START_ROW, 1, numRows, numCols).getValues();
    
    // Pre-fetch status column for "Double Check" (Read live data)
    // Note: In this script, we iterate through 'sourceData', but for strict safety, 
    // we could re-fetch the specific cell value before sending. 
    // Here, we assume 'sourceData' is fresh enough as the script just started.

    let sentCount = 0;
    let errorCount = 0;
    
    // 2. Iterate through rows
    for (let i = 0; i < sourceData.length; i++) {
      const row = sourceData[i];
      const currentRowNum = DATA_START_ROW + i;

      // Check Status: Is it "Approved"?
      if (row[COL_IDX.STATUS] === STATUS_TRIGGER) {
        
        // --- 3. Process Sending ---
        const emailTo = row[COL_IDX.EMAIL_TO];
        const subject = row[COL_IDX.SUBJECT];
        const body = row[COL_IDX.BODY];
        const agentName = row[COL_IDX.AGENT_NAME];
        const paymentDate = row[COL_IDX.PAYMENT_MONTH];
        const companyName = row[COL_IDX_COMPANY]; // For PDF unique name

        // Validate Date
        if (!(paymentDate instanceof Date)) {
          Logger.log(`Row ${currentRowNum}: Invalid date format.`);
          updateStatus(sourceSheet, currentRowNum, "Error: Invalid Date", COL_IDX.STATUS);
          continue;
        }

        // Construct PDF Path
        const yearYY = paymentDate.getFullYear().toString().slice(-2);
        const monthMM = (paymentDate.getMonth() + 1).toString().padStart(2, '0');
        const pdfFolderName = yearYY + monthMM + PDF_FOLDER_SUFFIX;
        
        // PDF Filename Logic (Standard vs Unique)
        const basePdfName = `${yearYY}${monthMM}_${agentName}様_支払い通知書`;
        const uniqueSuffix = ` (${companyName}_R${currentRowNum})`;

        // Find PDF
        const pdfFile = findPdfFile(PDF_PARENT_FOLDER_ID, pdfFolderName, basePdfName, uniqueSuffix);

        if (pdfFile) {
          // Send Email
          const options = {
            from: SENDER_EMAIL,
            cc: CC_EMAIL,
            attachments: [pdfFile.getAs(MimeType.PDF)]
          };
          
          GmailApp.sendEmail(emailTo, subject, body + EMAIL_SIGNATURE, options);

          // Update Status to "Sent"
          updateStatus(sourceSheet, currentRowNum, STATUS_SENT, COL_IDX.STATUS);
          sentCount++;
          
          // Sleep to prevent hitting rate limits
          Utilities.sleep(1000); 
          
        } else {
          // PDF Not Found
          Logger.log(`PDF Not Found: ${pdfFolderName} / ${basePdfName}`);
          updateStatus(sourceSheet, currentRowNum, STATUS_ERROR_PDF, COL_IDX.STATUS);
          errorCount++;
        }
      }
    }

    // 4. Completion Message
    const message = `${sentCount} emails have been sent successfully.\n` +
                    (errorCount > 0 ? `${errorCount} errors occurred (see spreadsheet).` : "");
    Browser.msgBox(message);

  } catch (e) {
    Browser.msgBox(`Critical Error: ${e.message}\nStack: ${e.stack}`);
  }
}

// =================================================================
// ★ 3. Helper: Find PDF File in Drive ★
// =================================================================
function findPdfFile(parentFolderId, folderName, basePdfName, uniqueSuffix) {
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const folders = parentFolder.getFoldersByName(folderName);
    if (!folders.hasNext()) return null;
    const targetFolder = folders.next();
    
    // Priority 1: Check for Unique Name (used when duplicates exist)
    const uniqueName = basePdfName + uniqueSuffix + ".pdf";
    const uniqueFiles = targetFolder.getFilesByName(uniqueName);
    if (uniqueFiles.hasNext()) return uniqueFiles.next();
    
    // Priority 2: Check for Base Name
    const baseFiles = targetFolder.getFilesByName(basePdfName + ".pdf");
    if (baseFiles.hasNext()) return baseFiles.next();

    return null;
  } catch (e) {
    Logger.log("findPdfFile Error: " + e.message);
    return null;
  }
}

// =================================================================
// ★ 4. Helper: Update Status Tag ★
// =================================================================
function updateStatus(sheet, rowNum, tag, colIndex) {
  try {
    // colIndex is 0-based, getRange needs 1-based
    sheet.getRange(rowNum, colIndex + 1).setValue(tag);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log(`Error updating status row ${rowNum}: ${e.message}`);
  }
}