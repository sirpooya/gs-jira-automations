function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📜 Invoice')
    .addItem('Show Invoice', 'showInvoice')
    .addToUi();
}

function showInvoice() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "Report" not found!');
    return;
  }
  
  // Get all data
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find column indices
  var fullCol = headers.indexOf('Full');
  var devCol = headers.indexOf('Dev');
  var collabCol = headers.indexOf('Collab');
  var teamCol = headers.indexOf('Team');
  var costCol = headers.indexOf('Cost');
  
  if (fullCol === -1 || devCol === -1 || collabCol === -1 || teamCol === -1 || costCol === -1) {
    SpreadsheetApp.getUi().alert('Required columns not found!');
    return;
  }
  
  // Get prices from row 2 (index 1)
  var price_full = data[1][fullCol];
  var price_dev = data[1][devCol];
  var price_collab = data[1][collabCol];
  
  // Find "Subtotal" row
  var subtotalRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][teamCol] === 'Subtotal') {
      subtotalRow = i;
      break;
    }
  }
  
  if (subtotalRow === -1) {
    SpreadsheetApp.getUi().alert('Row with "Subtotal" not found!');
    return;
  }
  
  // Calculate team_count (row number minus 3, but we need index-based calculation)
  // subtotalRow is 0-indexed, so if it's row 10, subtotalRow = 9
  // team_count = (subtotalRow + 1) - 3 = subtotalRow - 2
  var team_count = subtotalRow - 2;
  
  // Create arrays for teams (i=0 to team_count-1, rows i+3 which is index i+2)
  var team_name = [];
  var count_full = [];
  var count_dev = [];
  var count_collab = [];
  
  for (var i = 0; i < team_count; i++) {
    var rowIndex = i + 2; // row i+3 is index i+2 (since row 1 is index 0)
    if (rowIndex < data.length) {
      team_name[i] = data[rowIndex][teamCol];
      count_full[i] = data[rowIndex][fullCol];
      count_dev[i] = data[rowIndex][devCol];
      count_collab[i] = data[rowIndex][collabCol];
    }
  }
  
  // Get totals from Subtotal row
  var totalcount_full = data[subtotalRow][fullCol];
  var totalcount_dev = data[subtotalRow][devCol];
  var totalcount_collab = data[subtotalRow][collabCol];
  var cost_subtotal = data[subtotalRow][costCol];
  
  // Find "Prorated Costs" row
  var proratedRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][teamCol] === 'Prorated Costs') {
      proratedRow = i;
      break;
    }
  }
  
  var cost_prorated = 0;
  if (proratedRow !== -1) {
    cost_prorated = data[proratedRow][costCol];
  }
  
  // Find "Total" row
  var totalRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][teamCol] === 'Total') {
      totalRow = i;
      break;
    }
  }
  
  var cost_total = 0;
  if (totalRow !== -1) {
    cost_total = data[totalRow][costCol];
  }
  
  // Build invoice text
  var invoiceText = 'قیمت هر سیت: فول (' + price_full + ') — دولوپر (' + price_dev + ') — کلب (' + price_collab + ')\n\n';
  
  // Add team information
  for (var i = 0; i < team_name.length; i++) {
    invoiceText += 'تیم ' + count_full[i] + ' : ' + team_name[i] + ' فول - ' + count_dev[i] + ' دولوپر - ' + count_collab[i] + ' کلب\n';
  }
  
  invoiceText += '\nکل سیت‌ها: ' + totalcount_full + ' فول - ' + totalcount_dev + ' دولوپر - ' + totalcount_collab + ' کلب\n\n';
  invoiceText += 'مجموعا ' + cost_subtotal + ' بعلاوه ' + cost_prorated + ' هزینه سرشکن ماه قبل، مبلغ نهایی ' + cost_total;
  
  // Show modal dialog
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<div style="font-family: Arial, sans-serif; padding: 20px; white-space: pre-wrap; direction: rtl; text-align: right;">' + invoiceText + '</div>')
      .setWidth(600)
      .setHeight(400),
    'Invoice'
  );
}

