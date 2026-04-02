// ============================================================
// Fexa ROI Calculator — Google Apps Script
// ============================================================
// SETUP: Deploy > New deployment > Web app > Execute as: Me > Access: Anyone
// ============================================================

function doSetup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = [
    'Date','Client Name','Total Spend','Median Invoice','Total Invoices',
    'Total WOs','Proposals','Locations','Vendors','Call Center Volume',
    'Total Assets','FM Team Size','Location Growth %','Cost Avoidance',
    'Time Back (Hours)','Time Back ($)','Budget Module','Broker Savings','Grand Total'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#00313D').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    var p = e.parameter || {};

    // If it has an 'action' param or 'clientName', treat as save
    if (p.action === 'save' || p.clientName) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var row = [
        p.date || new Date().toISOString().split('T')[0],
        p.clientName || '',
        Number(p.totalSpend) || 0,
        Number(p.medianInvoice) || 0,
        Number(p.totalInvoices) || 0,
        Number(p.totalWOs) || 0,
        Number(p.proposals) || 0,
        Number(p.locations) || 0,
        Number(p.vendors) || 0,
        Number(p.callCenter) || 0,
        Number(p.totalAssets) || 0,
        Number(p.fmTeamSize) || 0,
        Number(p.locationGrowth) || 0,
        Number(p.costAvoidance) || 0,
        Number(p.timeBackHours) || 0,
        Number(p.timeBackDollars) || 0,
        Number(p.budgetModule) || 0,
        Number(p.brokerSavings) || 0,
        Number(p.grandTotal) || 0,
      ];
      sheet.appendRow(row);
      return ContentService.createTextOutput('OK');
    }

    return ContentService.createTextOutput('No action');
  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err.toString());
  }
}
