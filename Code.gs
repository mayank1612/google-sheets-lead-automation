function getTodayDateString() {
  const today = new Date();
  const day = today.getDate();
  const monthNames = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];
  const month = monthNames[today.getMonth()];
  
  return `${day} ${month}`;
}

function getYesterdayDateString() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  
  const day = yesterday.getDate();
  const monthNames = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];
  const month = monthNames[yesterday.getMonth()];
  
  return `${day} ${month}`;
}

// Function to set up daily trigger
function setupDailyTrigger() {
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyLeadManagement') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new daily trigger at 8:00 AM
  ScriptApp.newTrigger('dailyLeadManagement')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
    
  Logger.log("Daily trigger set up for 8:00 AM");
}

// Function to test the script manually
function testScript() {
  Logger.log("Starting manual test...");
  dailyLeadManagement();
  Logger.log("Manual test completed. Check the logs for details.");
}

// MAIN FUNCTION - Now properly separated
function dailyLeadManagement() {
  try {
    // Get the spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Step 1: Process Yesterday's Sheet
    const yesterdaySheetName = getYesterdayDateString();
    const yesterdaySheet = spreadsheet.getSheetByName(yesterdaySheetName);
    
    if (!yesterdaySheet) {
      throw new Error(`Yesterday's sheet "${yesterdaySheetName}" not found. Cannot proceed with daily lead management.`);
    }
    
    processYesterdaySheet(yesterdaySheet);
    
    // Step 2: Create Today's Sheet
    const todaySheetName = getTodayDateString();
    const todaySheet = createTodaySheet(spreadsheet, yesterdaySheet, todaySheetName);
    
    // Step 3: Clean Today's Sheet
    cleanTodaySheet(todaySheet);
    
    // Step 4: Fill from Pending Leads
    fillFromPendingLeads(spreadsheet, todaySheet);
    Logger.log("Completed filling from pending leads");
    
    // Step 5: Clear calling response column in today's sheet
    Logger.log("About to clear calling response column...");
    clearCallingResponseColumn(todaySheet);
    
    // Step 6: Move Pending Leads sheet to the end
    Logger.log("About to move pending sheet to end...");
    movePendingSheetToEnd(spreadsheet);
    
    Logger.log("Daily lead management completed successfully!");
    
    // Send success email
    sendSuccessEmail(spreadsheet.getName(), todaySheetName);
    
  } catch (error) {
    Logger.log("Error in dailyLeadManagement: " + error.toString());
    
    // Send error email
    sendErrorEmail(error.toString());
    
    throw error;
  }
}

function processYesterdaySheet(sheet) {
  const data = sheet.getDataRange().getValues();
  const callingResponseCol = 8; // Column I (0-indexed)
  const callerCol = 7; // Column H (0-indexed)
  
  // Skip header row, process in groups of 4
  for (let i = 1; i < data.length; i += 4) {
    const callerName = data[i][callerCol];
    if (!callerName) continue;
    
    // Check if this caller (4 rows) has any calling response
    let hasResponse = false;
    for (let j = 0; j < 4 && (i + j) < data.length; j++) {
      if (data[i + j][callingResponseCol] && data[i + j][callingResponseCol].toString().trim() !== '') {
        hasResponse = true;
        break;
      }
    }
    
    // If no response found, mark as "Reassigned"
    if (!hasResponse) {
      for (let j = 0; j < 4 && (i + j) < data.length; j++) {
        if (!data[i + j][callingResponseCol] || data[i + j][callingResponseCol].toString().trim() === '') {
          sheet.getRange(i + j + 1, callingResponseCol + 1).setValue("Reassigned");
        }
      }
    }
  }
  
  Logger.log("Yesterday's sheet processed for reassignments");
}

function createTodaySheet(spreadsheet, yesterdaySheet, todaySheetName) {
  // Check if today's sheet already exists
  let todaySheet = spreadsheet.getSheetByName(todaySheetName);
  if (todaySheet) {
    spreadsheet.deleteSheet(todaySheet);
  }
  
  // Duplicate yesterday's sheet
  todaySheet = yesterdaySheet.copyTo(spreadsheet);
  todaySheet.setName(todaySheetName);
  
  Logger.log("Today's sheet created: " + todaySheetName);
  return todaySheet;
}

function cleanTodaySheet(sheet) {
  const data = sheet.getDataRange().getValues();
  const callingResponseCol = 8; // Column I (0-indexed)
  const callerCol = 7; // Column H (0-indexed)
  
  let reassignedRows = [];
  let normalRows = [];
  
  // Separate reassigned and normal rows (skip header)
  for (let i = 1; i < data.length; i++) {
    if (data[i][callingResponseCol] === "Reassigned") {
      reassignedRows.push(data[i]);
    } else {
      // Keep only caller name, clear other data
      let cleanRow = new Array(data[i].length).fill('');
      cleanRow[callerCol] = data[i][callerCol]; // Keep caller name
      normalRows.push(cleanRow);
    }
  }
  
  // Clear all data except headers
  if (data.length > 1) {
    sheet.getRange(2, 1, data.length - 1, data[0].length).clearContent();
  }
  
  // Write reassigned rows first, then normal rows
  let currentRow = 2;
  
  // Write reassigned rows
  if (reassignedRows.length > 0) {
    sheet.getRange(currentRow, 1, reassignedRows.length, reassignedRows[0].length).setValues(reassignedRows);
    currentRow += reassignedRows.length;
  }
  
  // Write normal rows (with only caller names)
  if (normalRows.length > 0) {
    sheet.getRange(currentRow, 1, normalRows.length, normalRows[0].length).setValues(normalRows);
  }
  
  Logger.log("Today's sheet cleaned - reassigned rows moved to top");
}

function fillFromPendingLeads(spreadsheet, todaySheet) {
  const pendingSheet = spreadsheet.getSheetByName("Pending Leads");
  if (!pendingSheet) {
    Logger.log("Pending Leads sheet not found");
    return;
  }
  
  const pendingData = pendingSheet.getDataRange().getValues();
  if (pendingData.length <= 1) {
    Logger.log("No data in Pending Leads sheet");
    return;
  }
  
  const todayData = todaySheet.getDataRange().getValues();
  const callerCol = 7; // Column H (0-indexed)
  
  // Find empty rows that need to be filled (have caller name but no other data)
  let emptyRowsToFill = [];
  for (let i = 1; i < todayData.length; i++) {
    if (todayData[i][callerCol] && todayData[i][callerCol].toString().trim() !== '') {
      // Check if this row is essentially empty (except for caller name)
      let isEmpty = true;
      for (let j = 0; j < todayData[i].length; j++) {
        if (j !== callerCol && todayData[i][j] && todayData[i][j].toString().trim() !== '') {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) {
        emptyRowsToFill.push(i);
      }
    }
  }
  
  if (emptyRowsToFill.length === 0) {
    Logger.log("No empty rows to fill");
    return;
  }
  
  // Determine how many rows to take from pending leads
  const rowsToTake = Math.min(emptyRowsToFill.length, pendingData.length - 1);
  
  if (rowsToTake === 0) {
    Logger.log("No data available in Pending Leads");
    return;
  }
  
  // Get the data to move (excluding headers)
  const dataToMove = pendingData.slice(1, rowsToTake + 1);
  
  // Fill the empty rows in today's sheet
  for (let i = 0; i < dataToMove.length; i++) {
    const targetRow = emptyRowsToFill[i] + 1; // Convert to 1-indexed
    // Preserve the caller name, fill other columns
    const newRowData = [...dataToMove[i]];
    newRowData[callerCol] = todayData[emptyRowsToFill[i]][callerCol]; // Keep original caller
    
    todaySheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
  }
  
  // Delete the moved rows from Pending Leads sheet
  if (rowsToTake > 0) {
    pendingSheet.deleteRows(2, rowsToTake); // Delete from row 2, rowsToTake number of rows
  }
  
  Logger.log(`Moved ${rowsToTake} rows from Pending Leads to today's sheet`);
}

function clearCallingResponseColumn(sheet) {
  Logger.log("Starting to clear calling response column...");
  const data = sheet.getDataRange().getValues();
  const callingResponseCol = 9; // Column I (1-indexed for getRange)
  
  Logger.log(`Found ${data.length} rows in today's sheet`);
  
  if (data.length > 1) {
    // Clear calling response column for all data rows (excluding header)
    sheet.getRange(2, callingResponseCol, data.length - 1, 1).clearContent();
    Logger.log("Successfully cleared calling response column in today's sheet");
  } else {
    Logger.log("No data rows to clear in calling response column");
  }
}

function movePendingSheetToEnd(spreadsheet) {
  const pendingSheet = spreadsheet.getSheetByName("Pending Leads");
  if (!pendingSheet) {
    Logger.log("Pending Leads sheet not found");
    return;
  }
  
  const allSheets = spreadsheet.getSheets();
  const totalSheets = allSheets.length;
  
  // Activate the pending sheet first, then move it
  pendingSheet.activate();
  spreadsheet.moveActiveSheet(totalSheets);
  
  Logger.log("Moved Pending Leads sheet to the end");
}

// Email configuration - UPDATE THESE VALUES
const EMAIL_CONFIG = {
  recipients: "mayankbhagyawani101@gmail.com", // Change to your email address(es)
  // For multiple recipients: "email1@example.com,email2@example.com"
  senderName: "Lead Management System"
};

function sendSuccessEmail(spreadsheetName, todaySheetName) {
  try {
    const subject = `✅ Daily Lead Management Completed Successfully - ${todaySheetName}`;
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background-color: #28a745; color: white; padding: 20px; text-align: center;">
          <h2>✅ Daily Lead Management Successful</h2>
        </div>
        
        <div style="padding: 30px; background-color: #f8f9fa;">
          <h3 style="color: #28a745;">Process Completed Successfully!</h3>
          
          <div style="background-color: white; padding: 20px; border-radius: 5px; margin: 20px 0;">
            <p><strong>📊 Spreadsheet:</strong> ${spreadsheetName}</p>
            <p><strong>📅 Today's Sheet:</strong> ${todaySheetName}</p>
            <p><strong>⏰ Completed At:</strong> ${new Date().toLocaleString()}</p>
          </div>
          
          <h4 style="color: #333;">Tasks Completed:</h4>
          <ul style="color: #666;">
            <li>✓ Processed yesterday's sheet for reassignments</li>
            <li>✓ Created today's new sheet</li>
            <li>✓ Cleaned and organized today's data</li>
            <li>✓ Filled empty rows from pending leads</li>
            <li>✓ Cleared calling response column</li>
            <li>✓ Moved pending leads sheet to end</li>
          </ul>
          
          <div style="background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 5px; padding: 15px; margin: 20px 0;">
            <p style="margin: 0; color: #155724;">
              <strong>📋 Next Steps:</strong> Your team can now start making calls using the ${todaySheetName} sheet.
            </p>
          </div>
        </div>
        
        <div style="background-color: #6c757d; color: white; padding: 15px; text-align: center; font-size: 12px;">
          This is an automated message from your Lead Management System
        </div>
      </div>
    `;
    
    const plainTextBody = `
Daily Lead Management Completed Successfully!

Spreadsheet: ${spreadsheetName}
Today's Sheet: ${todaySheetName}
Completed At: ${new Date().toLocaleString()}

Tasks Completed:
- Processed yesterday's sheet for reassignments
- Created today's new sheet  
- Cleaned and organized today's data
- Filled empty rows from pending leads
- Cleared calling response column
- Moved pending leads sheet to end

Your team can now start making calls using the ${todaySheetName} sheet.

---
This is an automated message from your Lead Management System
    `;
    
    MailApp.sendEmail({
      to: EMAIL_CONFIG.recipients,
      subject: subject,
      htmlBody: htmlBody,
      body: plainTextBody,
      name: EMAIL_CONFIG.senderName
    });
    
    Logger.log(`Success email sent to: ${EMAIL_CONFIG.recipients}`);
    
  } catch (emailError) {
    Logger.log(`Failed to send success email: ${emailError.toString()}`);
  }
}

function sendErrorEmail(errorMessage) {
  try {
    const subject = `🚨 URGENT: Daily Lead Management Failed - ${new Date().toLocaleDateString()}`;
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background-color: #dc3545; color: white; padding: 20px; text-align: center;">
          <h2>🚨 Daily Lead Management Error</h2>
        </div>
        
        <div style="padding: 30px; background-color: #f8f9fa;">
          <div style="background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 5px; padding: 20px; margin: 20px 0;">
            <h3 style="color: #721c24; margin-top: 0;">⚠️ ATTENTION REQUIRED</h3>
            <p style="color: #721c24; margin-bottom: 0;">
              The daily lead management process has failed and requires immediate attention.
            </p>
          </div>
          
          <div style="background-color: white; padding: 20px; border-radius: 5px; margin: 20px 0;">
            <p><strong>⏰ Failed At:</strong> ${new Date().toLocaleString()}</p>
            <p><strong>📄 Expected Today's Sheet:</strong> ${getTodayDateString()}</p>
            <p><strong>📄 Expected Yesterday's Sheet:</strong> ${getYesterdayDateString()}</p>
          </div>
          
          <h4 style="color: #721c24;">Error Details:</h4>
          <div style="background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 5px; padding: 15px; margin: 15px 0;">
            <pre style="color: #856404; white-space: pre-wrap; font-family: monospace; margin: 0;">${errorMessage}</pre>
          </div>
          
          <h4 style="color: #333;">Immediate Actions Required:</h4>
          <ol style="color: #666;">
            <li>Check if yesterday's sheet exists with the correct name</li>
            <li>Verify the spreadsheet structure hasn't changed</li>
            <li>Ensure "Pending Leads" sheet exists</li>
            <li>Check Google Apps Script permissions</li>
            <li>Review the execution transcript in Google Apps Script</li>
            <li>Manually run the process if needed</li>
          </ol>
          
          <div style="background-color: #d1ecf1; border: 1px solid #bee5eb; border-radius: 5px; padding: 15px; margin: 20px 0;">
            <p style="margin: 0; color: #0c5460;">
              <strong>💡 Quick Fix:</strong> Try running the <code>testScript()</code> function manually from Google Apps Script to diagnose the issue.
            </p>
          </div>
        </div>
        
        <div style="background-color: #6c757d; color: white; padding: 15px; text-align: center; font-size: 12px;">
          This is an automated error notification from your Lead Management System
        </div>
      </div>
    `;
    
    const plainTextBody = `
🚨 URGENT: Daily Lead Management Failed

Failed At: ${new Date().toLocaleString()}
Expected Today's Sheet: ${getTodayDateString()}
Expected Yesterday's Sheet: ${getYesterdayDateString()}

ERROR DETAILS:
${errorMessage}

IMMEDIATE ACTIONS REQUIRED:
1. Check if yesterday's sheet exists with the correct name
2. Verify the spreadsheet structure hasn't changed  
3. Ensure "Pending Leads" sheet exists
4. Check Google Apps Script permissions
5. Review the execution transcript in Google Apps Script
6. Manually run the process if needed

Try running the testScript() function manually from Google Apps Script to diagnose the issue.

---
This is an automated error notification from your Lead Management System
    `;
    
    MailApp.sendEmail({
      to: EMAIL_CONFIG.recipients,
      subject: subject,
      htmlBody: htmlBody,
      body: plainTextBody,
      name: EMAIL_CONFIG.senderName
    });
    
    Logger.log(`Error email sent to: ${EMAIL_CONFIG.recipients}`);
    
  } catch (emailError) {
    Logger.log(`Failed to send error email: ${emailError.toString()}`);
    // Try sending a simple fallback email
    try {
      MailApp.sendEmail(
        EMAIL_CONFIG.recipients,
        "🚨 Lead Management System Error (Fallback)",
        `The daily lead management failed with error: ${errorMessage}\n\nAdditionally, the detailed error email failed to send: ${emailError.toString()}\n\nPlease check Google Apps Script immediately.`
      );
    } catch (fallbackError) {
      Logger.log(`Even fallback email failed: ${fallbackError.toString()}`);
    }
  }
}

/**
 * Web app entry point - can be called from Make.com
 */
function doPost(e) {
  try {
    const response = {
      success: true,
      message: "Daily lead management started",
      timestamp: new Date().toISOString()
    };
    
    // Run the main function
    dailyLeadManagement();
    
    response.message = "Daily lead management completed successfully";
    response.status = "completed";
    
    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    const errorResponse = {
      success: false,
      error: error.toString(),
      timestamp: new Date().toISOString()
    };
    
    return ContentService
      .createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
