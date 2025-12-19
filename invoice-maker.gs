function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📜 Invoice')
    .addItem('Show Invoice', 'showInvoice')
    .addToUi();
}

function showInvoice() {
  // Month data template
  var month_data_template = [
    {month: 'January', period_start: '30 January', period_end: '29 February'},
    {month: 'February', period_start: '29 February', period_end: '30 March'},
    {month: 'March', period_start: '30 March', period_end: '30 April'},
    {month: 'April', period_start: '30 April', period_end: '30 May'},
    {month: 'May', period_start: '30 May', period_end: '30 June'},
    {month: 'June', period_start: '30 June', period_end: '30 July'},
    {month: 'July', period_start: '30 July', period_end: '30 August'},
    {month: 'August', period_start: '30 August', period_end: '30 September'},
    {month: 'September', period_start: '30 September', period_end: '30 October'},
    {month: 'October', period_start: '30 October', period_end: '30 November'},
    {month: 'November', period_start: '30 November', period_end: '30 December'},
    {month: 'December', period_start: '30 December', period_end: '30 January'}
  ];
  
  // Get month data from Payments sheet
  var paymentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
  var month = '';
  var periodStart = '';
  var periodEnd = '';
  
  if (paymentsSheet) {
    var paymentsData = paymentsSheet.getDataRange().getValues();
    var paymentsHeaders = paymentsData[0];
    var monthCol = paymentsHeaders.indexOf('Month');
    
    if (monthCol !== -1) {
      // Find last row with data
      var lastRow = -1;
      for (var i = paymentsData.length - 1; i > 0; i--) {
        if (paymentsData[i][monthCol] && paymentsData[i][monthCol] !== '') {
          lastRow = i;
          break;
        }
      }
      
      if (lastRow !== -1) {
        var lastMonth = paymentsData[lastRow][monthCol];
        
        // Find current month in template and get next month
        for (var j = 0; j < month_data_template.length; j++) {
          if (month_data_template[j].month === lastMonth) {
            var nextIndex = (j + 1) % month_data_template.length; // Wrap around for December -> January
            month = month_data_template[nextIndex].month;
            periodStart = month_data_template[nextIndex].period_start;
            periodEnd = month_data_template[nextIndex].period_end;
            break;
          }
        }
      }
    }
  }
  
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
  var nameCol = headers.indexOf('Name');
  var costCol = headers.indexOf('Cost');
  
  if (fullCol === -1 || devCol === -1 || collabCol === -1 || teamCol === -1 || nameCol === -1 || costCol === -1) {
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
      team_name[i] = data[rowIndex][nameCol];
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
  var invoiceText = '';
  
  // Add subscription details at the beginning if month data is available
  if (month !== '') {
    invoiceText += 'جزئیات اشتراک فیگما فاکتور ماه ' + month + ' (از ' + periodStart + ' تا ' + periodEnd + ') به شرح زیر است:\n\n';
  }
  
  invoiceText += 'قیمت هر سیت: فول (<span dir="ltr">$' + price_full + '</span>) — دولوپر (<span dir="ltr">$' + price_dev + '</span>) — کلب (<span dir="ltr">$' + price_collab + '</span>)\n\n';
  
  // Add team information (skip if all counts are zero, otherwise print only non-zero counts)
  for (var i = 0; i < team_name.length; i++) {
    // Skip if all three counts are zero
    if (count_full[i] == 0 && count_dev[i] == 0 && count_collab[i] == 0) {
      continue;
    }
    
    // Build team line with only non-zero counts
    var teamLine = team_name[i] + ': ';
    var parts = [];
    
    if (count_full[i] > 0) {
      parts.push(count_full[i] + ' فول');
    }
    if (count_dev[i] > 0) {
      parts.push(count_dev[i] + ' دولوپر');
    }
    if (count_collab[i] > 0) {
      parts.push(count_collab[i] + ' کلب');
    }
    
    teamLine += parts.join(' - ');
    invoiceText += teamLine + '\n';
  }
  
  invoiceText += '\nکل سیت‌ها: ' + totalcount_full + ' فول - ' + totalcount_dev + ' دولوپر - ' + totalcount_collab + ' کلب\n\n';
  
  // Conditional cost printing
  if (cost_prorated == 0) {
    invoiceText += 'مجموعا مبلغ نهایی <span dir="ltr">$' + cost_total + '</span>';
  } else if (cost_prorated > 0) {
    invoiceText += 'مجموعا <span dir="ltr">$' + cost_subtotal + '</span> بعلاوه <span dir="ltr">$' + cost_prorated + '</span> هزینه سرشکن ماه قبل، مبلغ نهایی <span dir="ltr">$' + cost_total + '</span>';
  }
  
  // Show modal dialog
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<div style="font-family: Vazirmatn, sans-serif; padding: 20px; white-space: pre-wrap; direction: rtl; text-align: right;">' + invoiceText + '</div>')
      .setWidth(600)
      .setHeight(400),
    'Invoice'
  );
}

