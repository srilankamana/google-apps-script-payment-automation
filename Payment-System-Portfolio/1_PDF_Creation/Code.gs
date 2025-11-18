/**
 * Payment Notification PDF Generator
 * * Overview:
 * This script reads payment data from a Google Sheet, generates individual PDF files 
 * based on a template, saves them to a specific Google Drive folder, and updates the status.
 * * Key Features:
 * - Filters data by current month and empty status.
 * - Logic to skip rows where the payment amount is 0 or empty.
 * - Generates one PDF per row (1 transaction = 1 PDF).
 * - Handles duplicate file names by appending a suffix.
 * - Automatically adjusts font sizes to fit cell width.
 * - Updates the source sheet status to "Pending Approval" after creation.
 */

// =================================================================
// ★ Main Function: Create PDFs and Update Status ★
// =================================================================
function createPaymentPdfs() {

  // --- Configuration [SENSITIVE DATA MASKED] ---
  
  // File and Sheet IDs
  const SOURCE_SHEET_ID = "YOUR_SOURCE_SPREADSHEET_ID";     // e.g., "1i8uPjl..."
  const SOURCE_SHEET_NAME = "Payment_Management_Data";      // Sheet name containing the data
  const TEMPLATE_FILE_ID = "YOUR_TEMPLATE_SPREADSHEET_ID";  // e.g., "1Cj-SgU..."
  const TEMPLATE_SHEET_NAME = "Template_FMT";               // Sheet name of the invoice template
  const BASE_DRIVE_FOLDER_ID = "YOUR_DESTINATION_FOLDER_ID";// e.g., "10us6dx..."

  // Column Indices (0-based index)
  const COL_IDX = {
    COMPANY_NAME: 1,    // B: Recipient Company Name
    CHECK_AMOUNT: 12,   // M: Check column (must be 0 or empty to process)
    AGENT_NAME: 16,     // Q: Agent Name
    PAYMENT_MONTH: 19,  // T: Payment Month (Date format)
    PAYMENT_AMOUNT: 20, // U: Payment Amount
    BANK_ACCOUNT: 22,   // W: Bank Account Info
    STATUS: 25          // Z: Status Column
  };

  // Template Layout Settings
  const ROW_DETAIL_WRITE = 14; // Row number to write detail data (H14, X14)
  
  // Output Settings
  const PDF_NAME_SUFFIX = "_Payment_Notification"; // Suffix for PDF filename
  const FOLDER_NAME_SUFFIX = "_Payment_Notifications"; // Suffix for monthly folder

  // Status Tags
  const TAG_COMPLETE = "承認待ち";          // Status after PDF creation
  const TAG_WARNING = "債権残額要確認";      // Status if amount check fails

  // Font Size Adjustment Settings
  const FONT_H = { NORMAL: 10, SMALL: 8, THRESHOLD: 20 }; // For Company Name
  const FONT_F9 = { NORMAL: 8, SMALL: 6, THRESHOLD_LINES: 6 }; // For Bank Account

  // --- End of Configuration ---


  try {
    // --- 1. Setup Date and Folder Names ---
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth(); // 0-indexed
    const yearYY = year.toString().slice(-2);
    const monthMM = (month + 1).toString().padStart(2, '0');
    
    const sheetNamePrefix = yearYY + monthMM + "_";
    const targetFolderName = yearYY + monthMM + FOLDER_NAME_SUFFIX;

    // --- 2. Load Source Data ---
    const sourceSs = SpreadsheetApp.openById(SOURCE_SHEET_ID);
    const sourceSheet = sourceSs.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) throw new Error(`Sheet "${SOURCE_SHEET_NAME}" not found.`);

    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) {
      Browser.msgBox("No data found in the source sheet.");
      return;
    }

    // Read all data
    const sourceData = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
    
    // --- 3. Filter Data ---
    let dataToProcess = [];
    let rowsToUpdateWarning = [];

    for (let i = 0; i < sourceData.length; i++) {
      const row = sourceData[i];
      const originalRowNum = i + 2; // 1-based row index

      const paymentDate = row[COL_IDX.PAYMENT_MONTH];
      const currentStatus = row[COL_IDX.STATUS];
      const checkValue = row[COL_IDX.CHECK_AMOUNT]; // Column M

      // Filter: Current Month AND Empty Status
      let isTargetMonth = false;
      if (paymentDate instanceof Date) {
        if (paymentDate.getFullYear() === year && paymentDate.getMonth() === month) {
          isTargetMonth = true;
        }
      }

      if (isTargetMonth && currentStatus === "") {
        // Check Logic: M column must be "¥0", 0, or empty
        const isValidAmount = (checkValue === "¥0" || checkValue === 0 || checkValue === "");
        
        if (isValidAmount) {
          dataToProcess.push({ rowData: row, rowNum: originalRowNum });
        } else {
          rowsToUpdateWarning.push(originalRowNum);
        }
      }
    }

    // Update warnings if any invalid amounts found
    if (rowsToUpdateWarning.length > 0) {
      updateTags(sourceSheet, rowsToUpdateWarning, TAG_WARNING);
    }

    if (dataToProcess.length === 0) {
      Browser.msgBox("No valid data found to process.");
      return;
    }

    // --- 4. Prepare Template and Drive Folder ---
    const templateSs = SpreadsheetApp.openById(TEMPLATE_FILE_ID);
    const templateSheet = templateSs.getSheetByName(TEMPLATE_SHEET_NAME);
    if (!templateSheet) throw new Error(`Template "${TEMPLATE_SHEET_NAME}" not found.`);

    const baseFolder = DriveApp.getFolderById(BASE_DRIVE_FOLDER_ID);
    const folders = baseFolder.getFoldersByName(targetFolderName);
    const targetFolder = folders.hasNext() ? folders.next() : baseFolder.createFolder(targetFolderName);

    // --- 5. Generate PDFs (Loop) ---
    let processedCount = 0;

    for (const item of dataToProcess) {
      const row = item.rowData;
      const rowNum = item.rowNum;

      const companyName = row[COL_IDX.COMPANY_NAME];
      const agentName = row[COL_IDX.AGENT_NAME];
      const paymentAmount = row[COL_IDX.PAYMENT_AMOUNT];
      const bankInfo = row[COL_IDX.BANK_ACCOUNT];

      processedCount++;

      // 5-1. Create a temporary sheet from template
      // Naming: 2511_AgentName_CompanyName_RowNumber (Unique)
      const tempSheetName = `${sheetNamePrefix}${agentName}_${companyName}_R${rowNum}`;
      const tempSheet = templateSheet.copyTo(templateSs).setName(tempSheetName);
      tempSheet.showSheet();
      
      // Move to the first position (optional UI preference)
      tempSheet.activate();
      templateSs.moveActiveSheet(1);

      // 5-2. Write Data to the Sheet
      // Header Info
      tempSheet.getRange("B4").setValue(agentName);
      
      // Bank Info (F9) with Font Sizing
      const f9Cell = tempSheet.getRange("F9");
      f9Cell.setValue(bankInfo);
      const lineCount = bankInfo ? (bankInfo.toString().match(/\n/g) || []).length + 1 : 1;
      f9Cell.setFontSize(lineCount >= FONT_F9.THRESHOLD_LINES ? FONT_F9.SMALL : FONT_F9.NORMAL);
      f9Cell.setWrap(true).setVerticalAlignment("middle");

      // Totals (Z15-Z17)
      const subTotal = paymentAmount;
      const tax = Math.floor(subTotal * 0.1);
      tempSheet.getRange("Z15").setValue(subTotal);
      tempSheet.getRange("Z16").setValue(tax);
      tempSheet.getRange("Z17").setValue(subTotal + tax);

      // Detail Info (H14, X14) - 1 row only
      const hCell = tempSheet.getRange("H" + ROW_DETAIL_WRITE);
      const hText = `紹介手数料（${companyName}）`;
      hCell.setValue(hText);
      hCell.setFontSize(hText.length > FONT_H.THRESHOLD ? FONT_H.SMALL : FONT_H.NORMAL);
      hCell.setWrap(false).setVerticalAlignment("middle");

      tempSheet.getRange("X" + ROW_DETAIL_WRITE).setValue(paymentAmount);

      // 5-3. Export to PDF
      const basePdfName = `${sheetNamePrefix}${agentName}${PDF_NAME_SUFFIX}`;
      const uniqueSuffix = ` (${companyName}_R${rowNum})`; // Suffix for duplicates
      
      const isCreated = createPdf(templateSs, tempSheet, targetFolder, basePdfName, uniqueSuffix);

      // 5-4. Update Status
      if (isCreated) {
        updateTags(sourceSheet, [rowNum], TAG_COMPLETE);
      }
      
      // Cleanup: Optionally delete the temp sheet here if you don't want to keep it
      // templateSs.deleteSheet(tempSheet); 
    }

    // --- 6. Completion Message ---
    let msg = `Successfully created ${processedCount} PDFs in folder "${targetFolderName}".\n`;
    if (rowsToUpdateWarning.length > 0) {
      msg += `(Skipped ${rowsToUpdateWarning.length} rows due to M-column check)`;
    }
    Browser.msgBox(msg);

  } catch (e) {
    Browser.msgBox(`Error: ${e.message}\nStack: ${e.stack}`);
  }
}

// =================================================================
// ★ Helper: Create PDF with Duplicate Name Handling ★
// =================================================================
function createPdf(spreadsheet, sheet, folder, basePdfName, uniqueSuffix) {
  try {
    SpreadsheetApp.flush();
    const sheetId = sheet.getSheetId();
    const ssId = spreadsheet.getId();

    // Check for duplicate file names
    let pdfName = basePdfName + ".pdf";
    if (folder.getFilesByName(pdfName).hasNext()) {
      pdfName = basePdfName + uniqueSuffix + ".pdf";
    }

    // Construct Export URL
    const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
                `format=pdf&gid=${sheetId}&size=A4&portrait=false&fitw=true` +
                `&sheetnames=false&printtitle=false&gridlines=false` +
                `&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4`;

    const params = {
      method: "GET",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, params);
    if (response.getResponseCode() === 200) {
      folder.createFile(response.getBlob().setName(pdfName));
      return true;
    } else {
      Logger.log(`PDF Error: ${response.getContentText()}`);
      return false;
    }
  } catch (e) {
    Logger.log(`CreatePDF Error: ${e.message}`);
    return false;
  }
}

// =================================================================
// ★ Helper: Update Status Tags ★
// =================================================================
function updateTags(sheet, rowNumbers, tag) {
  if (!rowNumbers || rowNumbers.length === 0) return;
  const ranges = rowNumbers.map(r => "Z" + r);
  sheet.getRangeList(ranges).setValue(tag);
  SpreadsheetApp.flush();
}

// =================================================================
// ★ Helper: Add Custom Menu ★
// =================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Payment Automation')
    .addItem('Generate PDFs', 'createPaymentPdfs')
    .addToUi();
}