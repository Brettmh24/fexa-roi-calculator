// ============================================================
// Fexa ROI Calculator — Google Apps Script
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Create a new Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Paste this entire file into the script editor
// 4. Click Run > doSetup (this creates the header row)
// 5. Click Deploy > New deployment
//    - Type: Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Copy the Web App URL
// 7. Paste it into the calculator's "Google Sheets Config" field
// ============================================================

function doSetup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = [
    'Date',
    'Client Name',
    'Total Spend',
    'Median Invoice',
    'Total Invoices',
    'Total WOs',
    'Proposals',
    'Locations',
    'Vendors',
    'Call Center Volume',
    'Total Assets',
    'FM Team Size',
    'Location Growth %',
    'Cost Avoidance',
    'Time Back (Hours)',
    'Time Back ($)',
    'Budget Module',
    'Broker Savings',
    'Grand Total'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#00313D')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Auto-size columns
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Format currency columns
  var currencyCols = [3, 4, 14, 16, 17, 18, 19]; // 1-indexed
  currencyCols.forEach(function(col) {
    sheet.getRange(2, col, 500).setNumberFormat('$#,##0');
  });

  SpreadsheetApp.getUi().alert('Setup complete! Headers and formatting applied.');
}

function doPost(e) {
  try {
    // Support both form POST (payload field) and raw JSON
    var raw = (e.parameter && e.parameter.payload) ? e.parameter.payload : e.postData.contents;
    var data = JSON.parse(raw);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    var row = [
      data.date || new Date().toISOString().split('T')[0],
      data.clientName || '',
      data.totalSpend || 0,
      data.medianInvoice || 0,
      data.totalInvoices || 0,
      data.totalWOs || 0,
      data.proposals || 0,
      data.locations || 0,
      data.vendors || 0,
      data.callCenter || 0,
      data.totalAssets || 0,
      data.fmTeamSize || 0,
      data.locationGrowth || 0,
      data.costAvoidance || 0,
      data.timeBackHours || 0,
      data.timeBackDollars || 0,
      data.budgetModule || 0,
      data.brokerSavings || 0,
      data.grandTotal || 0,
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
