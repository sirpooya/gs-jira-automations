function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📜 Invoice')
    .addItem('Generate Invoice', 'showInvoice')
    .addItem('Send ticket reminder', 'sendTicketReminder')
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
  
  // Get month data from Invoices sheet
  var paymentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoices');
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
        
        // Write next month data to Invoices sheet
        var periodCol = paymentsHeaders.indexOf('Period');
        var statusCol = paymentsHeaders.indexOf('Status');
        
        if (month !== '') {
          var nextRowIndex = lastRow + 1; // Next row (0-indexed)
          var nextRowNumber = nextRowIndex + 1; // Convert to 1-indexed for setValue
          
          // Set Month column
          paymentsSheet.getRange(nextRowNumber, monthCol + 1).setValue(month);
          
          // Set Period column if it exists
          if (periodCol !== -1) {
            paymentsSheet.getRange(nextRowNumber, periodCol + 1).setValue(periodStart + ' - ' + periodEnd);
          }
          
          // Set Status column if it exists
          if (statusCol !== -1) {
            paymentsSheet.getRange(nextRowNumber, statusCol + 1).setValue('🔵 Upcoming');
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
    invoiceText += 'جزئیات اشتراک فیگما فاکتور ماه ' + month + ' (از ' + periodStart + ' تا ' + periodEnd + ') به شرح زیر است:\n';
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
    invoiceText += 'مبلغ نهایی <span dir="ltr">$' + cost_total + '</span>';
  } else if (cost_prorated > 0) {
    invoiceText += 'مجموعا <span dir="ltr">$' + cost_subtotal + '</span> بعلاوه <span dir="ltr">$' + cost_prorated + '</span> هزینه سرشکن ماه قبل، مبلغ نهایی <span dir="ltr">$' + cost_total + '</span>';
  }
  
  // Show modal dialog with download button
  var htmlContent = '<div style="font-family: Vazirmatn, sans-serif; padding: 20px; white-space: pre-wrap; direction: rtl; text-align: right;">' + 
    invoiceText + 
    '</div>' +
    '<div style="padding: 20px; text-align: center;">' +
    '<button onclick="downloadReport(\'' + month.replace(/'/g, "\\'") + '\')" style="padding: 10px 20px; font-size: 14px; cursor: pointer; background-color: #4285f4; color: white; border: none; border-radius: 4px; margin-right: 10px;">Download Report</button>' +
    '<button onclick="sendEmail(\'' + month.replace(/'/g, "\\'") + '\')" style="padding: 10px 20px; font-size: 14px; cursor: pointer; background-color: #34a853; color: white; border: none; border-radius: 4px;">Send Email</button>' +
    '</div>' +
    '<script>' +
    'function downloadReport(month) {' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    if (result && result.data) {' +
    '      var byteCharacters = atob(result.data);' +
    '      var byteNumbers = new Array(byteCharacters.length);' +
    '      for (var i = 0; i < byteCharacters.length; i++) {' +
    '        byteNumbers[i] = byteCharacters.charCodeAt(i);' +
    '      }' +
    '      var byteArray = new Uint8Array(byteNumbers);' +
    '      var blob = new Blob([byteArray], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});' +
    '      var url = URL.createObjectURL(blob);' +
    '      var a = document.createElement("a");' +
    '      a.href = url;' +
    '      a.download = result.filename;' +
    '      document.body.appendChild(a);' +
    '      a.click();' +
    '      document.body.removeChild(a);' +
    '      URL.revokeObjectURL(url);' +
    '    }' +
    '  }).downloadReportFile(month);' +
    '}' +
    'function sendEmail(month) {' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    if (result) { alert("Email sent successfully!"); }' +
    '  }).sendInvoiceEmail(month);' +
    '}' +
    '</script>';
  
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(htmlContent)
      .setWidth(600)
      .setHeight(450),
    'Invoice'
  );
}

function downloadReportFile(month) {
  try {
    // Get prices from Report sheet
    var reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report');
    var price_full = 0;
    var price_dev = 0;
    var price_collab = 0;
    
    if (reportSheet) {
      var reportData = reportSheet.getDataRange().getValues();
      var reportHeaders = reportData[0];
      var fullCol = reportHeaders.indexOf('Full');
      var devCol = reportHeaders.indexOf('Dev');
      var collabCol = reportHeaders.indexOf('Collab');
      
      if (fullCol !== -1 && devCol !== -1 && collabCol !== -1) {
        price_full = reportData[1][fullCol] || 0;
        price_dev = reportData[1][devCol] || 0;
        price_collab = reportData[1][collabCol] || 0;
      }
    }
    
    var paidUsersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Paid Users');
    
    if (!paidUsersSheet) {
      SpreadsheetApp.getUi().alert('Sheet "Paid Users" not found!');
      return null;
    }
    
    var data = paidUsersSheet.getDataRange().getValues();
    if (data.length === 0) {
      SpreadsheetApp.getUi().alert('No data found in "Paid Users" sheet!');
      return null;
    }
    
    var headers = data[0];
    var columnsToInclude = ['Name', 'Email', 'DK Email', 'Team', 'Department', 'Seat'];
    var columnIndices = [];
    
    // Find column indices
    for (var i = 0; i < columnsToInclude.length; i++) {
      var colIndex = headers.indexOf(columnsToInclude[i]);
      if (colIndex === -1) {
        SpreadsheetApp.getUi().alert('Column "' + columnsToInclude[i] + '" not found!');
        return null;
      }
      columnIndices.push(colIndex);
    }
    
    // Find Seat column index for filtering and manipulation
    var seatColIndex = headers.indexOf('Seat');
    if (seatColIndex === -1) {
      SpreadsheetApp.getUi().alert('Column "Seat" not found!');
      return null;
    }
    
    // Find Seat column index in filtered columns (for manipulation)
    var seatColIndexInFiltered = columnsToInclude.indexOf('Seat');
    
    // Filter data: only rows where Seat is not empty, and only selected columns
    var filteredData = [];
    filteredData.push(columnsToInclude); // Header row
    
    for (var i = 1; i < data.length; i++) {
      // Check if Seat column is not empty
      var seatValue = data[i][seatColIndex];
      if (seatValue !== null && seatValue !== undefined && seatValue !== '') {
        // Add row with only selected columns
        var row = [];
        for (var j = 0; j < columnIndices.length; j++) {
          var cellValue = data[i][columnIndices[j]];
          
          // Manipulate Seat column value using regex
          if (j === seatColIndexInFiltered) {
            var seatStr = String(cellValue);
            
            // Match and replace patterns using regex (allowing optional spaces)
            // Full 🔵🟢🟣🟠 → Full $price_full
            if (/Full\s*🔵\s*🟢\s*🟣\s*🟠/.test(seatStr) || /Full.*🔵.*🟢.*🟣.*🟠/.test(seatStr)) {
              cellValue = 'Full $' + price_full;
            } 
            // Dev 🟢🟣🟠 → Dev $price_dev
            else if (/Dev\s*🟢\s*🟣\s*🟠/.test(seatStr) || /Dev.*🟢.*🟣.*🟠/.test(seatStr)) {
              cellValue = 'Dev $' + price_dev;
            } 
            // Collab 🟣🟠 → Collab $price_collab
            else if (/Collab\s*🟣\s*🟠/.test(seatStr) || /Collab.*🟣.*🟠/.test(seatStr)) {
              cellValue = 'Collab $' + price_collab;
            }
          }
          
          row.push(cellValue);
        }
        filteredData.push(row);
      }
    }
    
    if (filteredData.length === 1) {
      SpreadsheetApp.getUi().alert('No rows found with non-empty Seat values!');
      return null;
    }
    
    // Create temporary spreadsheet
    var tempSpreadsheet = SpreadsheetApp.create('Temp_' + new Date().getTime());
    var tempSheet = tempSpreadsheet.getActiveSheet();
    var tempSpreadsheetId = tempSpreadsheet.getId();
    var tempSheetId = tempSheet.getSheetId();
    
    // Write filtered data
    tempSheet.getRange(1, 1, filteredData.length, columnsToInclude.length).setValues(filteredData);
    
    // Force all pending changes to be applied before exporting
    SpreadsheetApp.flush();
    
    // Verify data was written (read back to confirm)
    var verifyData = tempSheet.getDataRange().getValues();
    if (verifyData.length === 0 || verifyData[0].length === 0) {
      DriveApp.getFileById(tempSpreadsheetId).setTrashed(true);
      throw new Error('Data was not written to temporary spreadsheet');
    }
    
    // Small delay to ensure spreadsheet is ready for export
    Utilities.sleep(1000);
    
    // Export as Excel using export URL
    var fileName = (month !== '' ? month : 'Report') + '_supernova_invoice.xlsx';
    // Try exporting the entire spreadsheet first (more reliable)
    var exportUrl = 'https://docs.google.com/spreadsheets/d/' + tempSpreadsheetId + '/export?format=xlsx';
    
    // Fetch the exported file with proper authentication
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });
    
    // Check if export was successful
    if (response.getResponseCode() !== 200) {
      // Try with gid parameter as fallback
      exportUrl = 'https://docs.google.com/spreadsheets/d/' + tempSpreadsheetId + '/export?format=xlsx&gid=' + tempSheetId;
      response = UrlFetchApp.fetch(exportUrl, {
        muteHttpExceptions: true,
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + token
        }
      });
      
      if (response.getResponseCode() !== 200) {
        DriveApp.getFileById(tempSpreadsheetId).setTrashed(true);
        throw new Error('Failed to export spreadsheet. Response code: ' + response.getResponseCode() + ', Content: ' + response.getContentText().substring(0, 200));
      }
    }
    
    // Get the blob from response
    var blob = response.getBlob();
    blob.setName(fileName);
    
    // Convert blob to base64
    var base64Data = Utilities.base64Encode(blob.getBytes());
    
    // Delete temporary spreadsheet
    DriveApp.getFileById(tempSpreadsheetId).setTrashed(true);
    
    // Return base64 data and filename for client-side download
    return {
      data: base64Data,
      filename: fileName
    };
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error creating report: ' + error.toString());
    return null;
  }
}

function sendInvoiceEmail(month) {
  try {
    // Get invoice text (reuse logic from showInvoice but create plain text version)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report');
    
    if (!sheet) {
      return 'Error: Sheet "Report" not found!';
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var fullCol = headers.indexOf('Full');
    var devCol = headers.indexOf('Dev');
    var collabCol = headers.indexOf('Collab');
    var teamCol = headers.indexOf('Team');
    var nameCol = headers.indexOf('Name');
    var costCol = headers.indexOf('Cost');
    
    if (fullCol === -1 || devCol === -1 || collabCol === -1 || teamCol === -1 || nameCol === -1 || costCol === -1) {
      return 'Error: Required columns not found!';
    }
    
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
      return 'Error: Row with "Subtotal" not found!';
    }
    
    var team_count = subtotalRow - 2;
    var team_name = [];
    var count_full = [];
    var count_dev = [];
    var count_collab = [];
    
    for (var i = 0; i < team_count; i++) {
      var rowIndex = i + 2;
      if (rowIndex < data.length) {
        team_name[i] = data[rowIndex][nameCol];
        count_full[i] = data[rowIndex][fullCol];
        count_dev[i] = data[rowIndex][devCol];
        count_collab[i] = data[rowIndex][collabCol];
      }
    }
    
    var totalcount_full = data[subtotalRow][fullCol];
    var totalcount_dev = data[subtotalRow][devCol];
    var totalcount_collab = data[subtotalRow][collabCol];
    var cost_subtotal = data[subtotalRow][costCol];
    
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
    
    // Get month data
    var paymentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoices');
    var periodStart = '';
    var periodEnd = '';
    
    if (paymentsSheet && month) {
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
      
      for (var j = 0; j < month_data_template.length; j++) {
        if (month_data_template[j].month === month) {
          periodStart = month_data_template[j].period_start;
          periodEnd = month_data_template[j].period_end;
          break;
        }
      }
    }
    
    // Build plain text invoice
    var invoiceText = '';
    
    if (month !== '') {
      invoiceText += 'جزئیات اشتراک فیگما فاکتور ماه ' + month + ' (از ' + periodStart + ' تا ' + periodEnd + ') به شرح زیر است:\n';
    }
    
    invoiceText += 'قیمت هر سیت: فول ($' + price_full + ') — دولوپر ($' + price_dev + ') — کلب ($' + price_collab + ')\n\n';
    
    for (var i = 0; i < team_name.length; i++) {
      if (count_full[i] == 0 && count_dev[i] == 0 && count_collab[i] == 0) {
        continue;
      }
      
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
    
    if (cost_prorated == 0) {
      invoiceText += 'مبلغ نهایی $' + cost_total;
    } else if (cost_prorated > 0) {
      invoiceText += 'مجموعا $' + cost_subtotal + ' بعلاوه $' + cost_prorated + ' هزینه سرشکن ماه قبل، مبلغ نهایی $' + cost_total;
    }
    
    // Generate Excel file
    var excelResult = downloadReportFile(month);
    if (!excelResult || !excelResult.data) {
      return 'Error: Could not generate Excel file';
    }
    
    // Convert base64 back to blob
    var base64Data = excelResult.data;
    var byteCharacters = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(byteCharacters, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', excelResult.filename);
    
    // Send email with RTL HTML formatting
    var emailSubject = (month !== '' ? month : 'Report') + ' Supernova Invoice';
    var recipient = 'p.kamel@digikala.com';
    
    // Wrap invoice text in HTML with RTL direction
    var htmlBody = '<div style="font-family: Vazirmatn, Arial, sans-serif; direction: rtl; text-align: right; white-space: pre-wrap;">' + 
                   invoiceText.replace(/\n/g, '<br>') + 
                   '</div>';
    
    MailApp.sendEmail({
      to: recipient,
      subject: emailSubject,
      htmlBody: htmlBody,
      body: invoiceText, // Plain text fallback
      attachments: [blob]
    });
    
    return 'Email sent successfully to ' + recipient;
    
  } catch (error) {
    return 'Error sending email: ' + error.toString();
  }
}

