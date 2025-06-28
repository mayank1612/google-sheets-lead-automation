# google-sheets-lead-automation
Automates daily lead management by processing caller performance, creating new daily sheets, and distributing fresh leads from a pending queue - all scheduled to run automatically at 8 AM.


# 📞 Daily Lead Management Automation Script

## 🎯 Overview
This Google Apps Script automates daily lead management for sales teams by processing call assignments, tracking caller performance, and preparing new daily sheets with fresh leads. The script runs automatically every day at 8:00 AM, ensuring your team always has organized lead data to work with.

## 🔄 Complete Workflow with Final Precision

### **Step 1: Process Yesterday's Sheet (Daily at 8 AM)**
1. Find sheet named with yesterday's date (e.g., "28 June")
2. Check "Calling Response" column (Column I) for empty cells
3. For each caller group (4 consecutive rows per caller):
   * **If at least 1 row has data in "Calling Response":** Caller completed work (do nothing)
   * **If all 4 rows are empty in "Calling Response":** Write "Reassigned" in those 4 empty cells

### **Step 2: Create Today's Sheet**
1. Duplicate yesterday's sheet
2. Rename to today's date (e.g., "29 June")

### **Step 3: Clean Today's Sheet**
1. Keep "Reassigned" rows intact with all their data
2. For non-reassigned rows: Clear all data EXCEPT "Caller" column (Column H)
3. Move all "Reassigned" rows to the top
4. Below reassigned rows: Empty rows with only caller names in Column H

### **Step 4: Fill Empty Rows from Pending Leads**
1. Count how many empty rows need to be filled (rows that have caller name but no other data)
2. Take that many rows from "Pending Leads" sheet **starting from Row 2** (Row 1 = headers)
3. **CUT** those rows from "Pending Leads" and paste into empty rows of today's sheet
4. **DELETE** the now-empty rows from "Pending Leads" sheet (so remaining data moves up)

**Example:**
* If 8 caller slots need filling
* Take rows 2-9 from Pending Leads
* Cut and paste to today's sheet
* Delete rows 2-9 from Pending Leads
* Next time, data will start fresh from row 2 again

### **Step 5: Clear Calling Response Column**
1. Clear all data in "Calling Response" column (Column I) for today's sheet
2. Ensures fresh start for today's calling activities

### **Step 6: Organize Sheets**
1. Move "Pending Leads" sheet to the end of all sheets
2. Keeps daily sheets organized chronologically

## 🔑 Key Points
* **Always start from row 2** (after headers) in Pending Leads
* **After cutting data, delete the empty rows** to maintain clean sheet
* **Only fill as many rows as caller slots available**
* **Pending Leads maintains clean structure** for next day's operation
* **Automatic email notifications** for success and failure cases

## 📊 Sheet Structure Requirements

### Daily Sheets (e.g., "28 June", "29 June")
| Column | Description |
|--------|-------------|
| A-G | Lead data (name, phone, etc.) |
| **H** | **Caller** (person assigned to call) |
| **I** | **Calling Response** (results of calls) |

### Pending Leads Sheet
| Column | Description |
|--------|-------------|
| A-G | Lead data ready to be assigned |
| H | Can be empty (will be filled with caller names) |
| I | Should be empty initially |

## 🚀 How to Set Up the Script

### **1. Open Google Apps Script:**
* Go to your Google Sheets
* Extensions → Apps Script
* Delete the default code and paste the script
* Save the project with a meaningful name

### **2. Configure Email Notifications:**
```javascript
const EMAIL_CONFIG = {
  recipients: "your-email@example.com", // ⚠️ CHANGE THIS
  senderName: "Lead Management System"
};
```
* **Single recipient:** `"john@company.com"`
* **Multiple recipients:** `"john@company.com,jane@company.com"`

### **3. Set Up the Trigger:**
* **Option A (Recommended):** Run the `setupDailyTrigger()` function once
  - Click the function dropdown → Select `setupDailyTrigger`
  - Click the ▶️ Run button
  - This automatically schedules daily execution at 8:00 AM
  
* **Option B (Manual):** Create trigger in Apps Script editor
  - Click "Triggers" in left sidebar (⏰ icon)
  - Add Trigger → Choose function: `dailyLeadManagement`
  - Event source: Time-driven → Day timer → Every day at 8:00 AM

### **4. Test the Script:**
* Use the `testScript()` function to test manually before scheduled runs
* Check execution logs for any errors
* Verify email notifications are working

## 📧 Email Notifications

### ✅ Success Email
* **Sent when:** Process completes successfully
* **Contains:** Process summary, timestamp, completed tasks checklist
* **Action needed:** None, informational only

### 🚨 Error Email  
* **Sent when:** Process fails for any reason
* **Contains:** Error details, troubleshooting steps, immediate actions needed
* **Action needed:** Check spreadsheet structure, fix issues, potentially run manually

## 🛠️ Troubleshooting

### Common Issues:

**"Yesterday's sheet not found"**
* Ensure sheets are named with exact date format: "28 June" (number + space + month)
* Check if yesterday's sheet exists

**"Pending Leads sheet not found"**
* Create a sheet named exactly "Pending Leads"
* Ensure it has the same column structure as daily sheets

**Script not running automatically**
* Verify trigger is set up correctly
* Check Google Apps Script quotas and permissions
* Review execution transcript for errors

**Email not sending**
* Verify email addresses in EMAIL_CONFIG
* Check if MailApp has necessary permissions
* Review execution logs for email errors

### Manual Recovery:
If the automatic process fails:
1. Run `testScript()` function manually
2. Check logs to identify the issue
3. Fix the underlying problem
4. Re-run if necessary

## 📅 Maintenance

### Regular Tasks:
* **Weekly:** Review execution logs for any recurring issues
* **Monthly:** Verify email notifications are reaching the right people
* **Quarterly:** Update email recipients if team changes

### Data Management:
* **Archive old daily sheets** periodically to maintain performance
* **Monitor Pending Leads sheet** to ensure adequate lead supply
* **Backup spreadsheet** regularly

## 🔒 Security & Permissions

### Required Permissions:
* **Google Sheets:** Read/write access to spreadsheet
* **Gmail:** Send emails as notifications
* **Time-based triggers:** Schedule automatic execution

### Best Practices:
* Limit script editor access to authorized personnel only
* Regularly review and update email recipient lists
* Monitor script execution logs for unusual activity

## 📈 Benefits

* **🕐 Time Savings:** Eliminates 15-30 minutes of daily manual work
* **📊 Consistency:** Ensures uniform data processing every day
* **🎯 Accountability:** Tracks caller performance automatically
* **📧 Monitoring:** Immediate notifications for issues
* **🔄 Efficiency:** Seamless lead pipeline management

## 🆘 Support

For technical issues:
1. Check the execution transcript in Google Apps Script
2. Review error emails for specific guidance
3. Test manually using `testScript()` function
4. Verify spreadsheet structure matches requirements

---

**⚠️ Important:** Always test the script thoroughly with sample data before deploying to production spreadsheets.
