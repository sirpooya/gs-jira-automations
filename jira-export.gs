/**
 * CSV to Jira Story Converter - Google Apps Script Version (Best Practice)
 * 
 * This script processes selected rows in Google Sheets and converts them
 * to Jira import format based on the configuration below.
 * 
 * USAGE:
 * 1. Select the rows you want to process in your Google Sheet (excluding header row)
 * 2. Run this script
 * 3. CSV file will be downloaded automatically
 */

// ============================================================================
// CONFIGURATION - Edit these settings as needed
// ============================================================================

const CONFIG = {
  // Column mapping: old column name -> new column name
  rename: {
    'Component': 'Summary',
    'Pages': 'Labels', 
    'Impl Est': 'Story Points',
    'Design': 'Description',
    'Prototype': 'Description',
    'Category': 'Description',
    'Phase': 'Epic Link'
  },

  // Columns to append to description
  appendToDescription: [
    'Design',
    'Prototype', 
    'Category'
  ],

  // Multiline text to append to description
  appendMultiline: `----
Checklist:
⬜︎ Logical — prop config and values, slot exposition
⬜︎ UI — token-mapping, textstyle, spacings, colors & shadow, motion & transition
⬜︎ Interaction — desktop hover, mobile active, user input, keypress & select
⬜︎ Edges — layout responsiveness, empty cases, text-overflow, other browsers
`,

  // Static fields to add to each row
  add: {
    'Issue Type': 'Story',
    'Status': 'Planning Web'
  },

  // Phase number to Epic Link mapping
  phaseEpicLinks: {
    '1': 'DDS-1',
    '2': 'DDS-276'
    // Add more phases as needed: '3': 'DDS-XXX'
  },

  // Default epic link for phases not in the mapping
  defaultEpicLink: 'DDS-1'
};

// ============================================================================
// MAIN FUNCTION - Run this to process selected rows
// ============================================================================

function processSelectedRows() {
  try {
    // Get the active sheet and selected ranges (handles non-contiguous selections)
    const sheet = SpreadsheetApp.getActiveSheet();
    const rangeList = sheet.getActiveRangeList();
    
    if (!rangeList) {
      throw new Error('Please select some rows to process (excluding header row)');
    }

    // Always use row 1 as headers, regardless of selection
    const headers = getHeadersFromRow1(sheet);
    
    // Get data from all selected ranges (handles non-contiguous selections)
    const data = [];
    const ranges = rangeList.getRanges();
    
    for (const range of ranges) {
      const rangeData = range.getValues();
      data.push(...rangeData);
    }
    
    console.log(`Selected ${ranges.length} range(s): ${ranges.map(r => r.getA1Notation()).join(', ')}`);
    console.log(`Processing ${data.length} selected rows...`);
    console.log(`Headers found: ${headers.length}`);
    
    // Process the data
    const processedData = processData(data, headers);
    
    if (processedData.length === 0) {
      throw new Error('No data to process');
    }
    
    console.log(`Original rows: ${data.length}`);
    console.log(`Processed rows (including web/mobile): ${processedData.length}`);
    
    // Download CSV automatically
    downloadAsCSV(processedData);
    
    // Show toast notification
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Processed ${processedData.length} rows. Download starting...`,
      'Jira Export Complete',
      3
    );
    
    console.log(`✅ Successfully processed ${processedData.length} rows`);
    console.log(`📥 CSV download initiated`);
    
  } catch (error) {
    console.error('❌ Error:', error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error: ${error.message}`,
      'Export Failed',
      5
    );
    throw error;
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get headers from row 1 (always)
 */
function getHeadersFromRow1(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  return headerRange.getValues()[0];
}

/**
 * Process the data according to configuration
 */
function processData(data, headers) {
  const processed = [];
  
  console.log(`Processing ${data.length} rows with ${headers.length} headers`);
  console.log(`Headers: ${headers.join(', ')}`);
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowData = {};
    
    // Create a map of column name to value
    for (let j = 0; j < headers.length; j++) {
      rowData[headers[j]] = row[j];
    }
    
    console.log(`Row ${i + 1}: ${Object.keys(rowData).length} columns`);
    
    // Apply column renaming
    const newRow = {};
    for (const [oldCol, newCol] of Object.entries(CONFIG.rename)) {
      if (newCol === 'Description') continue; // Handle description separately
      newRow[newCol] = rowData[oldCol] || '';
    }
    
    // Add static fields
    for (const [col, val] of Object.entries(CONFIG.add)) {
      newRow[col] = val;
    }
    
    // Convert Phase to Epic Link
    const phaseNumber = (rowData['Phase'] || '').toString().trim();
    const epicLink = CONFIG.phaseEpicLinks[phaseNumber] || CONFIG.defaultEpicLink;
    newRow['Epic Link'] = epicLink;
    
    // Compose Description
    const descriptionParts = [];
    for (const col of CONFIG.appendToDescription) {
      if (rowData[col]) {
        descriptionParts.push(`${col}: ${rowData[col]}`);
      }
    }
    descriptionParts.push(CONFIG.appendMultiline);
    newRow['Description'] = descriptionParts.join('\n');
    
    processed.push(newRow);
  }
  
  // Create web and mobile versions (web first, then mobile for each row)
  const finalData = [];
  
  for (const row of processed) {
    const originalSummary = row['Summary'] || '';
    const isDesktopOnly = originalSummary.includes('🖥️');
    
    // Web version first (with 🌐)
    const webRow = { ...row };
    webRow['Summary'] = `${webRow['Summary']} 🌐`;
    finalData.push(webRow);
    
    // Mobile version second (with 📱) - skip if title contains 🖥️
    if (!isDesktopOnly) {
      const mobileRow = { ...row };
      mobileRow['Summary'] = `${mobileRow['Summary']} 📱`;
      mobileRow['Status'] = 'Planning App';
      finalData.push(mobileRow);
    }
  }
  
  return finalData;
}

/**
 * Download processed data as CSV (auto-triggers download)
 */
function downloadAsCSV(data) {
  if (data.length === 0) {
    throw new Error('No data to download');
  }
  
  // Get headers from first row
  const headers = Object.keys(data[0]);
  
  // Create CSV content
  let csvContent = headers.join(',') + '\n';
  
  // Add data rows
  for (const row of data) {
    const values = headers.map(header => {
      const value = row[header] || '';
      // Escape quotes and wrap in quotes if contains comma, quote, or newline
      const escapedValue = value.toString().replace(/"/g, '""');
      if (escapedValue.includes(',') || escapedValue.includes('"') || escapedValue.includes('\n')) {
        return `"${escapedValue}"`;
      }
      return escapedValue;
    });
    csvContent += values.join(',') + '\n';
  }
  
  // Create timestamp for filename
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const filename = `jira-import-${timestamp}.csv`;
  
  // Create HTML that auto-triggers download and closes
  const htmlOutput = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"></head><body>' +
    '<script>' +
    '(function() {' +
    '  var csvContent = ' + JSON.stringify(csvContent) + ';' +
    '  var filename = ' + JSON.stringify(filename) + ';' +
    '  var link = document.createElement("a");' +
    '  link.href = "data:text/csv;charset=utf-8," + encodeURIComponent(csvContent);' +
    '  link.download = filename;' +
    '  link.style.display = "none";' +
    '  document.body.appendChild(link);' +
    '  link.click();' +
    '  document.body.removeChild(link);' +
    '  setTimeout(function() { google.script.host.close(); }, 300);' +
    '})();' +
    '</script></body></html>'
  )
    .setWidth(1)
    .setHeight(1);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Downloading...');
}

// ============================================================================
// CUSTOM MENU - Creates menu in Google Sheets toolbar
// ============================================================================

/**
 * Creates custom menu when sheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Jira Exporter')
    .addItem('Process Selected Rows', 'processSelectedRows')
    .addToUi();
}

/**
 * Alternative: More integrated menu name
 */
function onOpenIntegrated() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Jira Exporter')
    .addItem('Export Selected Rows', 'processSelectedRows')
    .addToUi();
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Test function to validate configuration
 */
function testConfiguration() {
  console.log('Testing configuration...');
  console.log('Rename mappings:', Object.keys(CONFIG.rename).length);
  console.log('Phase epic links:', Object.keys(CONFIG.phaseEpicLinks).length);
  console.log('Static fields:', Object.keys(CONFIG.add).length);
  console.log('✅ Configuration looks good!');
}

/**
 * Show current configuration
 */
function showConfiguration() {
  console.log('Current Configuration:');
  console.log(JSON.stringify(CONFIG, null, 2));
} 