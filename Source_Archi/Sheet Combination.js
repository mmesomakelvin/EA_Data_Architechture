// Google Apps Script to combine and MERGE your EduBridge Academy sheets
function combineSheets() {
  // CONFIGURATION - Your actual sheet details
  const SHEET1_ID = '1TBLTPiDlpjGV5LPb2ahZ0wWUvoLauhHOW9-_cn-ZUyI'; // Data Analyst program sheet
  const SHEET1_TAB = 'Form responses 1';
  
  const SHEET2_ID = '1fQ0lce6cSEvV_WSG9Vygu6FDonkBptc6ou1R-doQhNo'; // Second sheet
  const SHEET2_TAB = 'Form responses 1';
  
  const SHEET3_ID = '1LX2S5it81y7Eg8d1lqnx59kNhzgFFOcRPrJrLkz00Js'; // Third sheet
  const SHEET3_TAB = 'Came List';
  
  const SHEET4_ID = '10QNmKpFnAkDP3GrHV4BnzOOzOzI8vSfc3QiyhgJKcMY'; // Fourth sheet - Payment List
  const SHEET4_TAB = 'Payment_List(UU)';
  
  const DESTINATION_SHEET_ID = '1G2b3SHGf883wtuGnA88iO0BQY4xtMS5-T8hwXvgcpps';
  const DESTINATION_TAB = 'Sheet1';
  
  try {
    // Open the sheets with error checking
    console.log('Opening Sheet 1...');
    const sheet1 = SpreadsheetApp.openById(SHEET1_ID).getSheetByName(SHEET1_TAB);
    if (!sheet1) throw new Error('Sheet 1 not found');
    
    console.log('Opening Sheet 2...');
    const sheet2 = SpreadsheetApp.openById(SHEET2_ID).getSheetByName(SHEET2_TAB);
    if (!sheet2) throw new Error('Sheet 2 not found');
    
    console.log('Opening Sheet 3...');
    const sheet3 = SpreadsheetApp.openById(SHEET3_ID).getSheetByName(SHEET3_TAB);
    if (!sheet3) throw new Error('Sheet 3 not found');
    
    console.log('Opening Sheet 4...');
    const sheet4 = SpreadsheetApp.openById(SHEET4_ID).getSheetByName(SHEET4_TAB);
    if (!sheet4) throw new Error('Sheet 4 not found');
    
    console.log('Opening Destination Sheet...');
    const destSheet = SpreadsheetApp.openById(DESTINATION_SHEET_ID).getSheetByName(DESTINATION_TAB);
    if (!destSheet) throw new Error('Destination sheet not found');
    
    // Get data from all sheets
    const data1 = getSheetData(sheet1);
    const data2 = getSheetData(sheet2);
    const data3 = getSheetData(sheet3);
    const data4 = getSheetData(sheet4);
    
    if (data1.length === 0 && data2.length === 0 && data3.length === 0 && data4.length === 0) {
      console.log('No data found in source sheets');
      return;
    }
    
    // Get existing data from destination
    const existingData = getSheetData(destSheet);
    
    // Process the data with smart merging
    const processedData = smartMergeSheets(data1, data2, data3, data4, existingData);
    
    // Clear and write the merged data
    destSheet.clear();
    if (processedData.combinedData.length > 0) {
      destSheet.getRange(1, 1, processedData.combinedData.length, processedData.combinedData[0].length)
        .setValues(processedData.combinedData);
      
      // Format headers
      destSheet.getRange(1, 1, 1, processedData.combinedData[0].length)
        .setFontWeight('bold')
        .setBackground('#f0f0f0');
      
      // Color code rows that came from different sheets
      if (processedData.sheet2Rows.length > 0) {
        processedData.sheet2Rows.forEach(rowIndex => {
          destSheet.getRange(rowIndex, 1, 1, processedData.combinedData[0].length)
            .setBackground('#E1D5F0'); // Light purple for sheet 2
        });
      }
      
      if (processedData.sheet3Rows.length > 0) {
        processedData.sheet3Rows.forEach(rowIndex => {
          destSheet.getRange(rowIndex, 1, 1, processedData.combinedData[0].length)
            .setBackground('#D5F0E1'); // Light green for sheet 3
        });
      }
      
      if (processedData.sheet4Rows.length > 0) {
        processedData.sheet4Rows.forEach(rowIndex => {
          destSheet.getRange(rowIndex, 1, 1, processedData.combinedData[0].length)
            .setBackground('#FFE4B5'); // Light orange for sheet 4
        });
      }
    }
    
    console.log('EduBridge Academy sheets merged successfully!');
    console.log(`Total unique records: ${processedData.combinedData.length - 1}`);
    console.log(`Records merged: ${processedData.mergedCount}`);
    console.log(`New records added: ${processedData.newRecords}`);
    console.log(`Attended records: ${processedData.attendedCount}`);
    
  } catch (error) {
    console.error('Error combining sheets:', error);
  }
}

function smartMergeSheets(data1, data2, data3, data4, existingData) {
  const result = {
    combinedData: [],
    sheet2Rows: [],
    sheet3Rows: [],
    sheet4Rows: [],
    mergedCount: 0,
    newRecords: 0,
    attendedCount: 0
  };
  
  // Create sets of email addresses from Sheet 3 and Sheet 4 for attendance tracking
  const sheet3Emails = new Set();
  const sheet4Emails = new Set();
  
  // Extract emails from Sheet 3
  const records3 = dataToRecords(data3);
  records3.forEach(record => {
    if (record['Email Address']) {
      const email = record['Email Address'].toString().trim().toLowerCase();
      sheet3Emails.add(email);
    }
  });
  
  // Extract emails from Sheet 4
  const records4 = dataToRecords(data4);
  records4.forEach(record => {
    if (record['Email Address']) {
      const email = record['Email Address'].toString().trim().toLowerCase();
      sheet4Emails.add(email);
    }
  });
  
  // Create master headers from all sheets and add "Attended" column
  const allHeaders = getMasterHeaders(data1, data2, data3, data4, existingData);
  
  // Add "Attended" column if not already present
  if (!allHeaders.includes('Attended')) {
    allHeaders.push('Attended');
  }
  
  result.combinedData.push(allHeaders);
  
  // Convert all data to record objects for easier processing
  const records1 = dataToRecords(data1);
  const records2 = dataToRecords(data2);
  const existingRecords = dataToRecords(existingData);
  
  // Create a map to track all records by email
  const recordMap = new Map();
  
  // Add existing records first
  existingRecords.forEach(record => {
    if (record['Email Address']) {
      const email = record['Email Address'].toString().trim().toLowerCase();
      recordMap.set(email, { ...record, source: 'existing' });
    }
  });
  
  // Process Sheet 1 records
  records1.forEach(record => {
    if (record['Email Address']) {
      const email = record['Email Address'].toString().trim().toLowerCase();
      if (recordMap.has(email)) {
        // Merge with existing record
        const existingRecord = recordMap.get(email);
        const mergedRecord = mergeRecords(existingRecord, record);
        recordMap.set(email, { ...mergedRecord, source: existingRecord.source });
        result.mergedCount++;
      } else {
        // New record
        recordMap.set(email, { ...record, source: 'sheet1' });
        result.newRecords++;
      }
    }
  });
  
  // Process Sheet 2 records
  records2.forEach(record => {
    if (record['Email Address']) {
      const email = record['Email Address'].toString().trim().toLowerCase();
      if (recordMap.has(email)) {
        // Merge with existing record
        const existingRecord = recordMap.get(email);
        const mergedRecord = mergeRecords(existingRecord, record);
        recordMap.set(email, { ...mergedRecord, source: existingRecord.source === 'existing' ? 'existing' : 'merged' });
        result.mergedCount++;
      } else {
        // New record from sheet 2
        recordMap.set(email, { ...record, source: 'sheet2' });
        result.newRecords++;
      }
    }
  });
  
  // Process Sheet 3 records
  records3.forEach(record => {
    if (record['Email Address']) {
      const email = record['Email Address'].toString().trim().toLowerCase();
      if (recordMap.has(email)) {
        // Merge with existing record
        const existingRecord = recordMap.get(email);
        const mergedRecord = mergeRecords(existingRecord, record);
        recordMap.set(email, { ...mergedRecord, source: existingRecord.source === 'existing' ? 'existing' : 'merged' });
        result.mergedCount++;
      } else {
        // New record from sheet 3
        recordMap.set(email, { ...record, source: 'sheet3' });
        result.newRecords++;
      }
    }
  });
  
  // Process Sheet 4 records
  records4.forEach(record => {
    if (record['Email Address']) {
      const email = record['Email Address'].toString().trim().toLowerCase();
      if (recordMap.has(email)) {
        // Merge with existing record
        const existingRecord = recordMap.get(email);
        const mergedRecord = mergeRecords(existingRecord, record);
        recordMap.set(email, { ...mergedRecord, source: existingRecord.source === 'existing' ? 'existing' : 'merged' });
        result.mergedCount++;
      } else {
        // New record from sheet 4
        recordMap.set(email, { ...record, source: 'sheet4' });
        result.newRecords++;
      }
    }
  });
  
  // Convert back to array format and track sheet rows
  let rowIndex = 2; // Start from row 2 (after header)
  recordMap.forEach(record => {
    const email = record['Email Address'] ? record['Email Address'].toString().trim().toLowerCase() : '';
    
    // Determine attendance status - 1 if found in Sheet 3 or Sheet 4, 0 otherwise
    const attended = (sheet3Emails.has(email) || sheet4Emails.has(email)) ? 1 : 0;
    record['Attended'] = attended;
    
    if (attended === 1) {
      result.attendedCount++;
    }
    
    const recordArray = allHeaders.map(header => {
      const value = record[header] || '';
      // Special handling for financial columns from third sheet
      return getDefaultValue(value, header);
    });
    result.combinedData.push(recordArray);
    
    // Track which rows came from which sheets
    if (record.source === 'sheet2') {
      result.sheet2Rows.push(rowIndex);
    } else if (record.source === 'sheet3') {
      result.sheet3Rows.push(rowIndex);
    } else if (record.source === 'sheet4') {
      result.sheet4Rows.push(rowIndex);
    }
    rowIndex++;
  });
  
  return result;
}

function getMasterHeaders(data1, data2, data3, data4, existingData) {
  const allHeaders = new Set();
  
  // Add headers from all sources
  [data1, data2, data3, data4, existingData].forEach(data => {
    if (data.length > 0) {
      data[0].forEach(header => {
        if (header) allHeaders.add(header);
      });
    }
  });
  
  return Array.from(allHeaders);
}

function dataToRecords(data) {
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const records = [];
  
  for (let i = 1; i < data.length; i++) {
    const record = {};
    headers.forEach((header, index) => {
      const value = data[i][index];
      // Use the new getDefaultValue function for consistent handling
      record[header] = getDefaultValue(value, header);
    });
    records.push(record);
  }
  
  return records;
}

function getDefaultValue(value, header) {
  // Check if value is empty, null, or undefined
  const isEmpty = (value === '' || value === null || value === undefined);
  
  // Financial columns from third sheet should default to 0
  const financialColumns = ['Payment', 'Expected', 'Amount owed'];
  const isFinancialColumn = financialColumns.some(col => 
    header && header.toLowerCase().includes(col.toLowerCase())
  );
  
  // Discount column should be handled specially - keep as "Not Recorded" when empty
  const isDiscountColumn = header && header.toLowerCase().includes('discount');
  
  // Attended column should default to 0 when empty
  const isAttendedColumn = header && header.toLowerCase() === 'attended';
  
  if (isEmpty) {
    if (isDiscountColumn) {
      return 'Not Recorded';
    } else if (isFinancialColumn || isAttendedColumn) {
      return 0;
    } else {
      return 'Not Recorded';
    }
  }
  
  // If discount column has a value, ensure it's formatted as percentage
  if (isDiscountColumn && !isEmpty) {
    const valueStr = value.toString().trim();
    if (valueStr) {
      // If it's already a percentage (contains %), keep it as is
      if (valueStr.includes('%')) {
        return valueStr;
      }
      // If it's a decimal (like 0.1), convert to percentage
      const numValue = parseFloat(valueStr);
      if (!isNaN(numValue)) {
        // If the number is between 0 and 1, treat it as a decimal to convert to percentage
        if (numValue >= 0 && numValue <= 1) {
          return (numValue * 100) + '%';
        }
        // If the number is greater than 1, assume it's already in percentage format, just add %
        else {
          return numValue + '%';
        }
      }
    }
  }
  
  return value;
}

function mergeRecords(record1, record2) {
  const merged = { ...record1 };
  
  // Merge fields, preferring non-empty values
  Object.keys(record2).forEach(key => {
    if (key !== 'source') {
      const defaultValue1 = getDefaultValue(merged[key], key);
      const defaultValue2 = getDefaultValue(record2[key], key);
      
      // Special handling for "Industry you work in" column - Sheet 2 always overwrites Sheet 1
      const isIndustryColumn = key && (
        key.toLowerCase().includes('industry you work in') || 
        key.toLowerCase().includes('industry') && key.toLowerCase().includes('work')
      );
      
      if (isIndustryColumn && record2[key] && record2[key] !== 'Not Recorded' && record2[key].toString().trim() !== '') {
        merged[key] = record2[key];
      }
      // If record1 has default value but record2 has real data, use record2's value
      else if ((merged[key] === 'Not Recorded' || merged[key] === 0 || !merged[key] || merged[key].toString().trim() === '') && 
          record2[key] && record2[key] !== 'Not Recorded' && record2[key] !== 0 && record2[key].toString().trim() !== '') {
        merged[key] = record2[key];
      }
      // If both have values and they're different, combine them (but not if one is a default value)
      else if (!isIndustryColumn && merged[key] && record2[key] && 
               merged[key] !== 'Not Recorded' && record2[key] !== 'Not Recorded' &&
               merged[key] !== 0 && record2[key] !== 0 &&
               merged[key].toString().trim() !== record2[key].toString().trim() &&
               merged[key].toString().trim() !== '' &&
               record2[key].toString().trim() !== '') {
        // For certain fields, we might want to combine rather than overwrite
        if (key.includes('Why') || key.includes('skills') || key.includes('project')) {
          merged[key] = merged[key] + ' | ' + record2[key];
        }
      }
    }
  });
  
  return merged;
}

function getSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow === 0 || lastCol === 0) {
    return [];
  }
  
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

// Function to set up automatic updates (run this once)
function setupAutomaticUpdates() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'combineSheets') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger to run every hour
  ScriptApp.newTrigger('combineSheets')
    .timeBased()
    .everyHours(1)
    .create();
    
  console.log('Automatic updates set up successfully!');
}