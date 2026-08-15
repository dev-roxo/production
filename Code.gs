/**
 * Sanitaryware Production Dashboard with User Login
 * Main Apps Script Code - GUID Implementation
 */

// Session management using PropertiesService
function setUserSession(userId, userName) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("userId", userId);
  userProperties.setProperty("userName", userName);
  userProperties.setProperty("loginTime", new Date().toISOString());
}

function getUserSession() {
  const userProperties = PropertiesService.getUserProperties();
  return {
    userId: userProperties.getProperty("userId"),
    userName: userProperties.getProperty("userName"),
    loginTime: userProperties.getProperty("loginTime"),
  };
}

/**
 * Generate a unique GUID
 * @customfunction
 */
function GUID() {
  return Utilities.getUuid();
}

function clearUserSession() {
  PropertiesService.getUserProperties().deleteAllProperties();
}

function isUserLoggedIn() {
  const session = getUserSession();
  return session.userId !== null && session.userId !== "";
}

/**
 * Authenticate user
 */
function authenticateUser(userId, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName("Users");

    if (!usersSheet) {
      return {
        success: false,
        message: "Users sheet not found! Please create Users sheet first.",
      };
    }

    const data = usersSheet.getDataRange().getValues();

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const sheetUserId = String(data[i][0]).trim();
      const sheetPassword = String(data[i][1]).trim();
      const fullName = String(data[i][2]).trim();
      const role = String(data[i][3]).trim();
      const active = data[i][4];

      if (sheetUserId === userId) {
        if (!active) {
          return {
            success: false,
            message: "User account is disabled. Please contact administrator.",
          };
        }

        if (sheetPassword === password) {
          // Set session
          setUserSession(userId, fullName);

          // Log login activity
          logActivity(
            userId,
            fullName,
            "Login",
            "",
            "User logged in successfully",
          );

          return {
            success: true,
            message: "Login successful!",
            userId: userId,
            userName: fullName,
            role: role,
          };
        } else {
          return { success: false, message: "Invalid password!" };
        }
      }
    }

    return { success: false, message: "User ID not found!" };
  } catch (error) {
    Logger.log("Error in authenticateUser: " + error.message);
    return {
      success: false,
      message: "Authentication error: " + error.message,
    };
  }
}

/**
 * Logout user
 */
function logoutUser() {
  try {
    const session = getUserSession();

    if (session.userId) {
      logActivity(
        session.userId,
        session.userName,
        "Logout",
        "",
        "User logged out",
      );
    }

    clearUserSession();
    return { success: true, message: "Logged out successfully!" };
  } catch (error) {
    Logger.log("Error in logoutUser: " + error.message);
    return { success: false, message: "Logout error: " + error.message };
  }
}

/**
 * Log user activity
 */
function logActivity(userId, userName, action, itemType, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName("ActivityLog");

    // Create ActivityLog sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet("ActivityLog");
      logSheet.appendRow([
        "Date & Time",
        "User ID",
        "User Name",
        "Action",
        "Item Type",
        "Details",
      ]);

      // Format header
      const headerRange = logSheet.getRange(1, 1, 1, 6);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#667eea");
      headerRange.setFontColor("#ffffff");
    }

    // Format timestamp using spreadsheet timezone
      const timestamp = Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      );

    logSheet.appendRow([
      timestamp,
      userId,
      userName,
      action,
      itemType,
      details,
    ]);
  } catch (error) {
    Logger.log("Error in logActivity: " + error.message);
  }
}

/**
 * Check if user is logged in (called before any operation)
 */
function checkAuth() {
  const session = getUserSession();
  if (!session.userId) {
    return { success: false, message: "Session expired. Please login again." };
  }
  return { success: true, session: session };
}

function parseDDMMYYYY(dateStr) {
  const [dd, mm, yyyy] = dateStr.split("/").map(Number);

  const tz = Session.getScriptTimeZone();
  const now = new Date();

  const h = Utilities.formatDate(now, tz, "HH");
  const m = Utilities.formatDate(now, tz, "mm");
  const s = Utilities.formatDate(now, tz, "ss");

  // Create date at noon first (prevents day shifting)
  const safeDate = new Date(`${yyyy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}T12:00:00`);

  // Then overwrite time with current script-timezone time
  safeDate.setHours(Number(h), Number(m), Number(s), 0);

  return safeDate;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Dashboard")
    .setTitle("Sorting Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Production")
    .addItem("Open Dashboard", "openDashboard")
    .addItem("Logout", "performLogout")
    .addToUi();
}

function performLogout() {
  logoutUser();
  Browser.msgBox("Logged out successfully!");
}

function openDashboard() {
  const html = HtmlService.createHtmlOutputFromFile("Dashboard")
    .setWidth(1400)
    .setHeight(800)
    .setTitle("Sorting Dashboard");

  SpreadsheetApp.getUi().showModalDialog(html, "Sorting Dashboard");
}

/**
 * Get all unique item types from Items sheet with sorted items
 */
function getItemTypes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemsSheet = ss.getSheetByName("Items");

    if (!itemsSheet) {
      Logger.log("Items sheet not found");
      return {};
    }

    const data = itemsSheet.getDataRange().getValues();
    const itemTypes = {};

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const itemName = String(data[i][0]).trim();
      const itemType = String(data[i][1]).trim();

      // Skip empty rows
      if (!itemName || !itemType) continue;

      if (!itemTypes[itemType]) {
        itemTypes[itemType] = [];
      }
      itemTypes[itemType].push(itemName);
    }

    // Sort items within each type alphabetically (case-insensitive)
    for (let itemType in itemTypes) {
      itemTypes[itemType].sort((a, b) =>
        a.toLowerCase().localeCompare(b.toLowerCase()),
      );
    }

    return itemTypes;
  } catch (error) {
    Logger.log("Error in getItemTypes: " + error.message);
    return {};
  }
}

function toDateObj(cellValue) {
  const tz = Session.getScriptTimeZone();
  if (cellValue instanceof Date && !isNaN(cellValue)) {
    return cellValue;
  }
  if (typeof cellValue === "string" && cellValue.trim() !== "") {
    try {
      const parsed = parseDDMMYYYY(cellValue.trim());
      if (parsed && !isNaN(parsed)) return parsed;
    } catch (e) {}
  }
  return null;
}

function getProductionData(fromDate, toDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SortingData");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const tz = Session.getScriptTimeZone();

  const fromBase = parseDDMMYYYY(fromDate);
  const toBase   = parseDDMMYYYY(toDate);

  const from = new Date(
    Utilities.formatDate(fromBase, tz, "yyyy-MM-dd'T'00:00:00")
  );

  const to = new Date(
    Utilities.formatDate(toBase, tz, "yyyy-MM-dd'T'23:59:59")
  );

  const result = [];

  for (let i = 1; i < data.length; i++) {
    const rowDate = toDateObj(data[i][1]); // Column B (Date)

    if (!rowDate) continue;

    if (rowDate >= from && rowDate <= to) {
      result.push({
        guid: String(data[i][0]).trim(), // Column A (GUID)
        date: Utilities.formatDate(
          rowDate,
          Session.getScriptTimeZone(),
          "dd/MM/yyyy",
        ),
        itemName: String(data[i][2]).trim(), // Column C
        itemType: String(data[i][3]).trim(), // Column D
        first: Number(data[i][4]) || 0, // Column E
        second: Number(data[i][5]) || 0, // Column F
        reject: Number(data[i][6]) || 0, // Column G
        bhangar: Number(data[i][7]) || 0, // Column H
        total: Number(data[i][8]) || 0, // Column I
      });
    }
  }

  Logger.log(`Found ${result.length} records`);
  return result;
}

/**
 * Initialize SortingData sheet with proper headers including GUID
 */
function initializeSortingDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sortingSheet = ss.getSheetByName("SortingData");
  
  if (!sortingSheet) {
    sortingSheet = ss.insertSheet("SortingData");
  }
  
  // Check if headers exist
  const lastCol = sortingSheet.getLastColumn();
  if (lastCol === 0 || sortingSheet.getRange(1, 1).getValue() === "") {
    // Set headers
    const headers = [
      "GUID", 
      "Date", 
      "Item Name", 
      "Item Type", 
      "First", 
      "Second", 
      "Reject", 
      "Bhangar", 
      "Total",
      "Entered By",
      "Entry Time"
    ];
    
    sortingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header
    const headerRange = sortingSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#667eea");
    headerRange.setFontColor("#ffffff");
  }
}

/**
 * Save sorting data for a specific date - GUID-based implementation
 *
 * BUG FIX: The original code used `if (!(sheetDate instanceof Date)) continue;`
 * which skipped every row whose Date cell was stored as a string (e.g. "31/12/2025").
 * This happened because the script writes `formattedDate` (a plain string) to the
 * sheet, so Google Sheets keeps it as a string on subsequent reads.
 * As a result the match loop never found an existing row and always appended a
 * new duplicate instead of updating.
 *
 * FIX: Parse both Date objects AND "dd/MM/yyyy" strings uniformly before comparing.
 */
function saveProductionData(date, itemType, productionData) {
  try {
    const session = getUserSession();

    if (!session.userId) {
      return { success: false, message: "Please login first" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const entriesSheet = ss.getSheetByName("SortingData");

    if (!entriesSheet) {
      return { success: false, message: "SortingData sheet not found!" };
    }

    // Initialize sheet if needed
    initializeSortingDataSheet();

    const formattedDate = Utilities.formatDate(
      parseDDMMYYYY(date),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );
    const formattedDateObj = parseDDMMYYYY(formattedDate);

    const itemTypeStr = String(itemType).trim();
    const tz = Session.getScriptTimeZone();

    const timestamp = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
    const formattedDateKey = Utilities.formatDate(formattedDateObj, tz, "yyyy-MM-dd");

    // ── Build lookup map: key → (0-based row index, guid) ─────────────────
    const allData = entriesSheet.getDataRange().getValues();
    const rowMap = new Map();

    function toDateKey(cellValue) {
      if (cellValue instanceof Date && !isNaN(cellValue)) {
        return Utilities.formatDate(cellValue, tz, "yyyy-MM-dd");
      }
      if (typeof cellValue === "string" && cellValue.trim() !== "") {
        try {
          const parsed = parseDDMMYYYY(cellValue.trim());
          if (parsed && !isNaN(parsed)) {
            return Utilities.formatDate(parsed, tz, "yyyy-MM-dd");
          }
        } catch (e) {}
      }
      return null;
    }

    // Skip header row (i=0)
    for (let i = 1; i < allData.length; i++) {
      const dateKey = toDateKey(allData[i][1]);
      if (!dateKey) continue;
      const key = dateKey + "|" + String(allData[i][2]).trim() + "|" + String(allData[i][3]).trim();
      rowMap.set(key, { rowIndex: i, guid: String(allData[i][0]).trim() });
    }

    // ── Categorize items ─────────────────────────────────────────────────
    const updates = [];    // { sheetRow: 1-based, rowData: [...] }
    const inserts = [];    // rowData arrays
    const deletions = [];  // 1-based row indices (collected, then deleted bottom-up)

    let savedCount = 0;
    let skippedCount = 0;
    let deletedCount = 0;
    let itemDetails = [];

    productionData.forEach((item) => {
      const first   = Number(item.first)   || 0;
      const second  = Number(item.second)  || 0;
      const reject  = Number(item.reject)  || 0;
      const bhangar = Number(item.bhangar) || 0;
      const total   = Number(item.total)   || 0;

      const itemNameStr = String(item.itemName).trim();
      const hasData = first > 0 || second > 0 || reject > 0 || bhangar > 0;

      const lookupKey = formattedDateKey + "|" + itemNameStr + "|" + itemTypeStr;
      const existing = rowMap.get(lookupKey);

      if (hasData) {
        const rowData = [
          existing ? existing.guid : Utilities.getUuid(),
          formattedDate,
          itemNameStr,
          itemTypeStr,
          first,
          second,
          reject,
          bhangar,
          total,
          session.userName,
          timestamp,
        ];

        if (existing) {
          updates.push({ sheetRow: existing.rowIndex + 1, rowData: rowData });
          savedCount++;
          itemDetails.push(itemNameStr + " (updated)");
        } else {
          inserts.push(rowData);
          savedCount++;
          itemDetails.push(itemNameStr + " (new)");
        }
      } else {
        if (existing) {
          deletions.push(existing.rowIndex + 1);
          deletedCount++;
          itemDetails.push(itemNameStr + " (deleted)");
        } else {
          skippedCount++;
        }
      }
    });

    // ── Execute batch writes ─────────────────────────────────────────────
    // Updates: write each updated row (individual ranges, but no interleaved reads)
    if (updates.length > 0) {
      updates.forEach(function(u) {
        entriesSheet.getRange(u.sheetRow, 1, 1, 11).setValues([u.rowData]);
      });
    }

    // Inserts: append all at once in a single batch
    if (inserts.length > 0) {
      entriesSheet.getRange(allData.length + 1, 1, inserts.length, 11).setValues(inserts);
    }

    // Deletions: delete from bottom to top so indices stay valid
    if (deletions.length > 0) {
      deletions.sort(function(a, b) { return b - a; });
      deletions.forEach(function(row) { entriesSheet.deleteRow(row); });
    }

    // Log the operation
    logActivity(
      session.userId,
      session.userName,
      "Save Data",
      itemTypeStr,
      `Saved ${savedCount} items, Deleted ${deletedCount} items, Skipped ${skippedCount} items. Items: ${itemDetails.join(", ")}`,
    );

    return {
      success: true,
      message: `Saved ${savedCount} items, Deleted ${deletedCount} items, Skipped ${skippedCount} items`,
      details: itemDetails,
    };
  } catch (error) {
    Logger.log("Error in saveProductionData: " + error.toString());
    return { success: false, message: "Error: " + error.toString() };
  }
}

/**
 * Get dashboard summary data
 */
function getDashboardSummary(fromDate, toDate) {
  try {
    const data = getProductionData(fromDate, toDate);

    let totalFirst = 0;
    let totalSecond = 0;
    let totalReject = 0;
    let totalBhangar = 0;
    let grandTotal = 0;

    data.forEach((row) => {
      totalFirst += Number(row.first) || 0;
      totalSecond += Number(row.second) || 0;
      totalReject += Number(row.reject) || 0;
      totalBhangar += Number(row.bhangar) || 0;
      grandTotal += Number(row.total) || 0;
    });

    // Calculate percentages
    const firstPct =
      grandTotal > 0 ? ((totalFirst / grandTotal) * 100).toFixed(2) : "0.00";
    const secondPct =
      grandTotal > 0 ? ((totalSecond / grandTotal) * 100).toFixed(2) : "0.00";
    const rejectPct =
      grandTotal > 0 ? ((totalReject / grandTotal) * 100).toFixed(2) : "0.00";
    const bhangarPct =
      grandTotal > 0 ? ((totalBhangar / grandTotal) * 100).toFixed(2) : "0.00";

    return {
      totalFirst,
      totalSecond,
      totalReject,
      totalBhangar,
      grandTotal,
      firstPct,
      secondPct,
      rejectPct,
      bhangarPct,
    };
  } catch (error) {
    Logger.log("Error in getDashboardSummary: " + error.message);
    return {
      totalFirst: 0,
      totalSecond: 0,
      totalReject: 0,
      totalBhangar: 0,
      grandTotal: 0,
      firstPct: "0.00",
      secondPct: "0.00",
      rejectPct: "0.00",
      bhangarPct: "0.00",
    };
  }
}

/**
 * Process data for print
 */
function processDataForPrint(dataArray) {
  const processed = {};

  dataArray.forEach((row) => {
    const itemType = String(row.itemType).trim();
    const itemName = String(row.itemName).trim();

    if (!processed[itemType]) {
      processed[itemType] = {};
    }
    if (!processed[itemType][itemName]) {
      processed[itemType][itemName] = {
        first: 0,
        second: 0,
        reject: 0,
        bhangar: 0,
        total: 0,
      };
    }

    processed[itemType][itemName].first += Number(row.first) || 0;
    processed[itemType][itemName].second += Number(row.second) || 0;
    processed[itemType][itemName].reject += Number(row.reject) || 0;
    processed[itemType][itemName].bhangar += Number(row.bhangar) || 0;
    processed[itemType][itemName].total += Number(row.total) || 0;
  });

  return processed;
}

/**
 * Generate HTML for printing - keeping your existing function with user name added
 */
function generatePrintHTML(fromDate, toDate) {
  try {
    const session = getUserSession();
    const userName = session.userName || "Unknown User";

    if (session.userId) {
      logActivity(
        session.userId,
        userName,
        "Print Report",
        "",
        `Generated print report for ${fromDate} to ${toDate}`,
      );
    }

    const itemTypes = getItemTypes();
    const data = getProductionData(fromDate, toDate);
    const processedData = processDataForPrint(data);
    const summary = getDashboardSummary(fromDate, toDate);

    let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Sorting Dashboard - Print View</title>
  <style>
    @page {
      size: landscape;
      margin: 1cm;
    }
    
    @media print {
      body {
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      
      .page-break {
        page-break-after: always;
      }
      
      .no-print {
        display: none;
      }
    }
    
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      padding: 20px;
      background: white;
    }
    
    .print-header {
      text-align: center;
      margin-bottom: 20px;
      border-bottom: 3px solid #667eea;
      padding-bottom: 15px;
    }
    
    .print-header h1 {
      color: #667eea;
      font-size: 28px;
      margin-bottom: 5px;
    }
    
    .print-header .subtitle {
      color: #666;
      font-size: 14px;
    }
    
    .print-header .date-range {
      color: #333;
      font-weight: 600;
      margin-top: 5px;
      font-size: 16px;
    }
    
    .summary-section {
      display: flex;
      justify-content: space-around;
      margin: 20px 0;
      gap: 15px;
    }
    
    .summary-card {
      flex: 1;
      text-align: center;
      padding: 15px;
      border-radius: 8px;
      border: 2px solid #ddd;
    }
    
    .summary-card h3 {
      font-size: 12px;
      color: #666;
      margin-bottom: 8px;
    }
    
    .summary-card .value {
      font-size: 24px;
      font-weight: bold;
      margin-bottom: 5px;
    }
    
    .summary-card .percentage {
      font-size: 14px;
      color: #666;
    }
    
    .summary-card.total {
      background: #f0f0f0;
      border-color: #667eea;
    }
    
    .summary-card.first {
      background: #e8f5e9;
      border-color: #4caf50;
    }
    
    .summary-card.second {
      background: #e3f2fd;
      border-color: #2196f3;
    }
    
    .summary-card.reject {
      background: #fff3e0;
      border-color: #ff9800;
    }
    
    .summary-card.bhangar {
      background: #f3e5f5;
      border-color: #9c27b0;
    }
    
    .item-section {
      margin-bottom: 30px;
    }
    
    .section-title {
      background: linear-gradient(135deg, #C85A28 0%, #E67E22 100%);
      color: white;
      padding: 12px;
      text-align: center;
      font-size: 18px;
      font-weight: bold;
      border-radius: 8px 8px 0 0;
    }
    
    .data-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 10px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .data-table thead {
      background: #C85A28;
      color: white;
    }
    
    .data-table th {
      padding: 12px;
      text-align: center;
      font-weight: 600;
      font-size: 13px;
      border: 1px solid #ddd;
    }
    
    .data-table td {
      padding: 10px;
      text-align: center;
      border: 1px solid #ddd;
      font-size: 12px;
    }
    
    .data-table td:first-child {
      text-align: left;
      font-weight: 600;
      color: #28a745;
    }
    
    .data-table .total-row {
      background: #fff3cd;
      font-weight: bold;
    }
    
    .data-table .percentage-row {
      background: #d1ecf1;
      font-weight: bold;
      color: #0c5460;
    }
    
    .print-footer {
      text-align: center;
      margin-top: 20px;
      padding-top: 10px;
      border-top: 2px solid #ddd;
      font-size: 11px;
      color: #666;
    }
    
    .print-button {
      position: fixed;
      top: 20px;
      right: 20px;
      padding: 12px 24px;
      background: #667eea;
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      box-shadow: 0 4px 8px rgba(0,0,0,0.2);
      z-index: 1000;
    }
    
    .print-button:hover {
      background: #5568d3;
    }
  </style>
</head>
<body>
  <button class="print-button no-print" onclick="window.print()">🖨️ Print</button>
  
  <div class="print-header">
    <h1>Sorting Reporting</h1>
    <p class="subtitle">Real-time Sorting Tracking & Quality Analysis</p>
    <p class="date-range">Date Range: ${fromDate} to ${toDate}</p>
    <p class="subtitle" style="margin-top: 5px;">Generated by: ${userName}</p>
  </div>
  
  <div class="summary-section">
    <div class="summary-card total">
      <h3>TOTAL ITEMS</h3>
      <div class="value">${summary.grandTotal}</div>
      <div class="percentage">100%</div>
    </div>
    <div class="summary-card first">
      <h3>FIRST QUALITY</h3>
      <div class="value">${summary.totalFirst}</div>
      <div class="percentage">${summary.firstPct}%</div>
    </div>
    <div class="summary-card second">
      <h3>SECOND QUALITY</h3>
      <div class="value">${summary.totalSecond}</div>
      <div class="percentage">${summary.secondPct}%</div>
    </div>
    <div class="summary-card reject">
      <h3>REJECT</h3>
      <div class="value">${summary.totalReject}</div>
      <div class="percentage">${summary.rejectPct}%</div>
    </div>
    <div class="summary-card bhangar">
      <h3>BHANGAR</h3>
      <div class="value">${summary.totalBhangar}</div>
      <div class="percentage">${summary.bhangarPct}%</div>
    </div>
  </div>
`;

    // Sort item types alphabetically
    const itemTypeKeys = Object.keys(itemTypes).sort();

    // Generate tables for each item type
    let itemIndex = 0;

    for (let i = 0; i < itemTypeKeys.length; i++) {
      const itemType = itemTypeKeys[i];
      const items = itemTypes[itemType];
      const typeData = processedData[itemType] || {};

      let totals = { first: 0, second: 0, reject: 0, bhangar: 0, total: 0 };
      let tableRows = "";

      // Sort items alphabetically too
      const sortedItems = items
        .slice()
        .sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

      sortedItems.forEach((itemName) => {
        const data = typeData[itemName] || {
          first: 0,
          second: 0,
          reject: 0,
          bhangar: 0,
          total: 0,
        };
        totals.first += data.first;
        totals.second += data.second;
        totals.reject += data.reject;
        totals.bhangar += data.bhangar;
        totals.total += data.total;

        // Escape HTML in item names
        const escapedItemName = String(itemName)
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#039;");

        tableRows += `
          <tr>
            <td>${escapedItemName}</td>
            <td>${data.first}</td>
            <td>${data.second}</td>
            <td>${data.reject}</td>
            <td>${data.bhangar}</td>
            <td>${data.total}</td>
          </tr>
        `;
      });

      // Calculate percentages
      const firstPct =
        totals.total > 0
          ? ((totals.first / totals.total) * 100).toFixed(2)
          : "0.00";
      const secondPct =
        totals.total > 0
          ? ((totals.second / totals.total) * 100).toFixed(2)
          : "0.00";
      const rejectPct =
        totals.total > 0
          ? ((totals.reject / totals.total) * 100).toFixed(2)
          : "0.00";
      const bhangarPct =
        totals.total > 0
          ? ((totals.bhangar / totals.total) * 100).toFixed(2)
          : "0.00";

      // Add page break after every item type except the last one
      const pageBreak = itemIndex < itemTypeKeys.length - 1 ? "page-break" : "";

      // Escape HTML in item type name
      const escapedItemType = String(itemType)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");

      html += `
    <div class="item-section ${pageBreak}">
      <div class="section-title">${escapedItemType}</div>
      <table class="data-table">
        <thead>
          <tr>
            <th>${escapedItemType}</th>
            <th>FIRST</th>
            <th>SECOND</th>
            <th>REJECT</th>
            <th>BHANGAR</th>
            <th>TOTAL</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
          <tr class="total-row">
            <td>TOTAL</td>
            <td>${totals.first}</td>
            <td>${totals.second}</td>
            <td>${totals.reject}</td>
            <td>${totals.bhangar}</td>
            <td>${totals.total}</td>
          </tr>
          <tr class="percentage-row">
            <td>PERCENTAGE</td>
            <td>${firstPct}%</td>
            <td>${secondPct}%</td>
            <td>${rejectPct}%</td>
            <td>${bhangarPct}%</td>
            <td>100%</td>
          </tr>
        </tbody>
      </table>
    </div>
      `;

      itemIndex++;
    }

    const now = new Date();
    html += `
    <div class="print-footer">
      <p>Generated on ${now.toLocaleDateString()} at ${now.toLocaleTimeString()}</p>
      <p>Roxo Sanitaryware - Sorting Dashboard</p>
    </div>
  </body>
  </html>
    `;

    return html;
  } catch (error) {
    Logger.log("Error generating print HTML: " + error.message);
    throw new Error("Error generating print HTML: " + error.message);
  }
}

/**
 * Add new item to Items sheet
 */
function addNewItem(itemName, itemType) {
  try {
    const session = getUserSession();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemsSheet = ss.getSheetByName("Items");

    if (!itemsSheet) {
      return { success: false, message: "Items sheet not found!" };
    }

    const itemNameStr = String(itemName).trim();
    const itemTypeStr = String(itemType).trim();

    if (!itemNameStr || !itemTypeStr) {
      return { success: false, message: "Item name and type cannot be empty!" };
    }

    // Get all existing items
    const lastRow = itemsSheet.getLastRow();
    if (lastRow > 1) {
      const existingItems = itemsSheet
        .getRange(2, 1, lastRow - 1, 2)
        .getValues();

      // Check if item already exists (case-insensitive)
      for (let i = 0; i < existingItems.length; i++) {
        const existingName = String(existingItems[i][0]).trim();
        const existingType = String(existingItems[i][1]).trim();

        if (existingName === itemNameStr && existingType === itemTypeStr) {
          return { success: false, message: "Item already exists!" };
        }
      }
    }

    // Add new item
    itemsSheet.appendRow([itemNameStr, itemTypeStr]);

    // Log activity if user is logged in
    if (session.userId) {
      logActivity(
        session.userId,
        session.userName,
        "Add Item",
        itemTypeStr,
        `Added new item: ${itemNameStr}`,
      );
    }

    return {
      success: true,
      message: `Item "${itemNameStr}" added successfully to ${itemTypeStr}!`,
    };
  } catch (error) {
    Logger.log("Error adding item: " + error.message);
    return {
      success: false,
      message: "Error adding item: " + error.message,
    };
  }
}

/**
 * Get unique item types
 */
function getUniqueItemTypes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemsSheet = ss.getSheetByName("Items");

    if (!itemsSheet) {
      return [];
    }

    const lastRow = itemsSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    const itemTypes = itemsSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const uniqueTypes = [
      ...new Set(
        itemTypes
          .flat()
          .map((type) => String(type).trim())
          .filter((type) => type !== ""),
      ),
    ];

    return uniqueTypes.sort();
  } catch (error) {
    Logger.log("Error getting item types: " + error.message);
    return [];
  }
}

/**
 * Get daily report for specific item - GUID aware
 */
function getDailyReport(itemName, fromDate, toDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const productionSheet = ss.getSheetByName("SortingData");

    if (!productionSheet) {
      throw new Error("Production Data sheet not found!");
    }

    const lastRow = productionSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    const lastCol = productionSheet.getLastColumn();
    const data = productionSheet
      .getRange(2, 1, lastRow - 1, Math.max(lastCol, 11))
      .getValues();
    const dailyData = [];

    // Convert input dates
    const fromDateObj = parseDDMMYYYY(fromDate);
    fromDateObj.setHours(0, 0, 0, 0);

    const toDateObj = parseDDMMYYYY(toDate);
    toDateObj.setHours(23, 59, 59, 999);

    const targetItemName = String(itemName).trim();

    data.forEach((row) => {
      const rowDate = toDateObj(row[1]); // Column B (Date)
      const rowItemName = String(row[2]).trim(); // Column C

      if (
        rowDate &&
        rowItemName === targetItemName &&
        rowDate >= fromDateObj &&
        rowDate <= toDateObj
      ) {
        const first = Number(row[4]) || 0;
        const second = Number(row[5]) || 0;
        const reject = Number(row[6]) || 0;
        const bhangar = Number(row[7]) || 0;

        dailyData.push({
          guid: String(row[0]).trim(),
          date: Utilities.formatDate(
            rowDate,
            Session.getScriptTimeZone(),
            "d EEE",
          ),
          itemName: rowItemName,
          itemType: String(row[3]).trim(),
          first: first,
          second: second,
          reject: reject,
          bhangar: bhangar,
          total: first + second + reject + bhangar,
          enteredBy: row[9] ? String(row[9]).trim() : "",
          entryTime: row[10] ? row[10] : "",
        });
      }
    });

    // Sort by date
    dailyData.sort((a, b) => {
      const dateA = parseDDMMYYYY(a.date);
      const dateB = parseDDMMYYYY(b.date);
      return dateA - dateB;
    });

    return dailyData;
  } catch (error) {
    Logger.log("Error getting daily report: " + error.message);
    throw new Error("Error getting daily report: " + error.message);
  }
}

/**
 * Get all item names
 */
function getAllItemNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemsSheet = ss.getSheetByName("Items");

    if (!itemsSheet) {
      return [];
    }

    const lastRow = itemsSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    const items = itemsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const itemNames = items
      .flat()
      .map((name) => String(name).trim())
      .filter((name) => name !== "");

    // Remove duplicates and sort
    return [...new Set(itemNames)].sort();
  } catch (error) {
    Logger.log("Error getting item names: " + error.message);
    return [];
  }
}

/**
 * Generate daily report HTML - GUID aware
 */
function generateDailyReportHTML(itemName, fromDate, toDate) {
  try {
    const session = getUserSession();
    const userName = session.userName || "Unknown User";

    if (session.userId) {
      logActivity(
        session.userId,
        userName,
        "Daily Report",
        "",
        `Generated daily report for ${itemName} from ${fromDate} to ${toDate}`,
      );
    }

    const dailyData = getDailyReport(itemName, fromDate, toDate);

    if (dailyData.length === 0) {
      return `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="UTF-8">
          <title>Daily Report - No Data</title>
          <style>
            body {
              font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
              padding: 40px;
              text-align: center;
            }
            .no-data {
              font-size: 18px;
              color: #666;
            }
          </style>
        </head>
        <body>
          <div class="no-data">
            <h2>No Data Found</h2>
            <p>No sorting data found for "${String(itemName).replace(/</g, "&lt;").replace(/>/g, "&gt;")}" between ${fromDate} and ${toDate}</p>
          </div>
        </body>
        </html>
      `;
    }

    // Calculate totals
    let totals = {
      first: 0,
      second: 0,
      reject: 0,
      bhangar: 0,
      total: 0,
    };

    dailyData.forEach((day) => {
      totals.first += day.first;
      totals.second += day.second;
      totals.reject += day.reject;
      totals.bhangar += day.bhangar;
      totals.total += day.total;
    });

    // Calculate percentages
    const firstPct =
      totals.total > 0
        ? ((totals.first / totals.total) * 100).toFixed(2)
        : "0.00";
    const secondPct =
      totals.total > 0
        ? ((totals.second / totals.total) * 100).toFixed(2)
        : "0.00";
    const rejectPct =
      totals.total > 0
        ? ((totals.reject / totals.total) * 100).toFixed(2)
        : "0.00";
    const bhangarPct =
      totals.total > 0
        ? ((totals.bhangar / totals.total) * 100).toFixed(2)
        : "0.00";

    // Generate table rows
    let tableRows = "";
    dailyData.forEach((day) => {
      tableRows += `
        <tr>
          <td>${day.date}</td>
          <td>${day.first}</td>
          <td>${day.second}</td>
          <td>${day.reject}</td>
          <td>${day.bhangar}</td>
          <td class="total-cell">${day.total}</td>
        </tr>
      `;
    });

    const itemType = dailyData[0].itemType;
    const escapedItemName = String(itemName)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    const escapedItemType = String(itemType)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");

    const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Daily Sorting Report - ${escapedItemName}</title>
  <style>
    @page {
      size: portrait;
      margin: 1cm;
    }
    
    @media print {
      body {
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      
      .no-print {
        display: none;
      }
    }
    
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      padding: 20px;
      background: white;
    }
    
    .report-header {
      text-align: center;
      margin-bottom: 30px;
      border-bottom: 3px solid #667eea;
      padding-bottom: 20px;
    }
    
    .report-header h1 {
      color: #667eea;
      font-size: 32px;
      margin-bottom: 10px;
    }
    
    .report-header .item-info {
      font-size: 20px;
      color: #333;
      font-weight: 600;
      margin: 10px 0;
    }
    
    .report-header .date-range {
      color: #666;
      font-size: 16px;
      margin-top: 5px;
    }
    
    .report-header .item-type {
      display: inline-block;
      background: linear-gradient(135deg, #C85A28 0%, #E67E22 100%);
      color: white;
      padding: 8px 20px;
      border-radius: 20px;
      font-size: 14px;
      margin-top: 10px;
    }
    
    .summary-section {
      display: flex;
      justify-content: space-around;
      margin: 30px 0;
      gap: 15px;
    }
    
    .summary-card {
      flex: 1;
      text-align: center;
      padding: 20px;
      border-radius: 12px;
      border: 2px solid #ddd;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .summary-card h3 {
      font-size: 13px;
      color: #666;
      margin-bottom: 10px;
      text-transform: uppercase;
    }
    
    .summary-card .value {
      font-size: 28px;
      font-weight: bold;
      margin-bottom: 5px;
    }
    
    .summary-card .percentage {
      font-size: 16px;
      color: #666;
    }
    
    .summary-card.total {
      background: #f8f9fa;
      border-color: #667eea;
    }
    
    .summary-card.total .value {
      color: #667eea;
    }
    
    .summary-card.first {
      background: #e8f5e9;
      border-color: #4caf50;
    }
    
    .summary-card.first .value {
      color: #2e7d32;
    }
    
    .summary-card.second {
      background: #e3f2fd;
      border-color: #2196f3;
    }
    
    .summary-card.second .value {
      color: #1565c0;
    }
    
    .summary-card.reject {
      background: #fff3e0;
      border-color: #ff9800;
    }
    
    .summary-card.reject .value {
      color: #e65100;
    }
    
    .summary-card.bhangar {
      background: #f3e5f5;
      border-color: #9c27b0;
    }
    
    .summary-card.bhangar .value {
      color: #6a1b9a;
    }
    
    .daily-table-section {
      margin: 30px 0;
    }
    
    .section-title {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 15px;
      text-align: center;
      font-size: 20px;
      font-weight: bold;
      border-radius: 12px 12px 0 0;
    }
    
    .data-table {
      width: 100%;
      border-collapse: collapse;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      background: white;
    }
    
    .data-table thead {
      background: #667eea;
      color: white;
    }
    
    .data-table th {
      padding: 15px 10px;
      text-align: center;
      font-weight: 600;
      font-size: 13px;
      border: 1px solid #ddd;
    }
    
    .data-table td {
      padding: 12px 10px;
      text-align: center;
      border: 1px solid #ddd;
      font-size: 13px;
    }
    
    .data-table td:first-child {
      font-weight: 600;
      color: #667eea;
      text-align: center;
    }
    
    .data-table tbody tr:nth-child(even) {
      background: #f8f9fa;
    }
    
    .data-table tbody tr:hover {
      background: #e9ecef;
    }
    
    .data-table .total-cell {
      font-weight: bold;
      color: #667eea;
    }
    
    .data-table .total-row {
      background: #fff3cd !important;
      font-weight: bold;
      font-size: 14px;
    }
    
    .data-table .percentage-row {
      background: #d1ecf1 !important;
      font-weight: bold;
      color: #0c5460;
    }
    
    .print-footer {
      text-align: center;
      margin-top: 30px;
      padding-top: 15px;
      border-top: 2px solid #ddd;
      font-size: 12px;
      color: #666;
    }
    
    .print-button {
      position: fixed;
      top: 20px;
      right: 20px;
      padding: 12px 24px;
      background: #667eea;
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      box-shadow: 0 4px 8px rgba(0,0,0,0.2);
      z-index: 1000;
    }
    
    .print-button:hover {
      background: #5568d3;
    }
    
    .stats-info {
      display: flex;
      justify-content: space-around;
      margin: 20px 0;
      padding: 15px;
      background: #f8f9fa;
      border-radius: 8px;
    }
    
    .stats-info .stat {
      text-align: center;
    }
    
    .stats-info .stat .label {
      font-size: 12px;
      color: #666;
      margin-bottom: 5px;
    }
    
    .stats-info .stat .value {
      font-size: 18px;
      font-weight: bold;
      color: #333;
    }
  </style>
</head>
<body>
  <button class="print-button no-print" onclick="window.print()">🖨️ Print Report</button>
  
  <div class="report-header">
    <h1>Daily Sorting Report</h1>
    <div class="item-info">${escapedItemName}</div>
    <div class="item-type">${escapedItemType}</div>
    <div class="date-range">Period: ${fromDate} to ${toDate}</div>
    <p class="date-range" style="margin-top: 5px;">Generated by: ${userName}</p>
  </div>
  
  <div class="stats-info">
    <div class="stat">
      <div class="label">Total Days</div>
      <div class="value">${dailyData.length}</div>
    </div>
    <div class="stat">
      <div class="label">Average Per Day</div>
      <div class="value">${(totals.total / dailyData.length).toFixed(0)}</div>
    </div>
    <div class="stat">
      <div class="label">Best Day</div>
      <div class="value">${Math.max(...dailyData.map((d) => d.total))}</div>
    </div>
  </div>
  
  <div class="summary-section">
    <div class="summary-card total">
      <h3>Total Sorting</h3>
      <div class="value">${totals.total}</div>
      <div class="percentage">100%</div>
    </div>
    <div class="summary-card first">
      <h3>First Quality</h3>
      <div class="value">${totals.first}</div>
      <div class="percentage">${firstPct}%</div>
    </div>
    <div class="summary-card second">
      <h3>Second Quality</h3>
      <div class="value">${totals.second}</div>
      <div class="percentage">${secondPct}%</div>
    </div>
    <div class="summary-card reject">
      <h3>Reject</h3>
      <div class="value">${totals.reject}</div>
      <div class="percentage">${rejectPct}%</div>
    </div>
    <div class="summary-card bhangar">
      <h3>Bhangar</h3>
      <div class="value">${totals.bhangar}</div>
      <div class="percentage">${bhangarPct}%</div>
    </div>
  </div>

  <div class="daily-table-section">
    <div class="section-title">Daily Sorting Details</div>
    <table class="data-table">
      <thead>
        <tr>
          <th>Date</th>
          <th>First</th>
          <th>Second</th>
          <th>Reject</th>
          <th>Bhangar</th>
          <th>Total</th>
        </tr>
      </thead>
      <tbody>
        ${tableRows}
        <tr class="total-row">
          <td>TOTAL</td>
          <td>${totals.first}</td>
          <td>${totals.second}</td>
          <td>${totals.reject}</td>
          <td>${totals.bhangar}</td>
          <td>${totals.total}</td>
        </tr>
        <tr class="percentage-row">
          <td>OVERALL %</td>
          <td>${firstPct}%</td>
          <td>${secondPct}%</td>
          <td>${rejectPct}%</td>
          <td>${bhangarPct}%</td>
          <td>100%</td>
        </tr>
      </tbody>
    </table>
  </div>
  
  <div class="print-footer">
    <p>Generated on ${new Date().toLocaleDateString()} at ${new Date().toLocaleTimeString()}</p>
    <p>XYZ Sanitaryware - Daily Sorting Report</p>
  </div>
</body>
</html>
    `;

    return html;
  } catch (error) {
    Logger.log("Error generating daily report HTML: " + error.message);
    throw new Error("Error generating daily report HTML: " + error.message);
  }
}
