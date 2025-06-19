// Google Apps Script to combine and MERGE your EduBridge Academy sheets
function combineSheets() {
  // CONFIGURATION - Your actual sheet details
  const SHEET1_ID = '1TBLTPiDlpjGV5LPb2ahZ0wWUvoLauhHOW9-_cn-ZUyI'; // Data Analyst program sheet
  const SHEET1_TAB = 'Form responses 1';
  
  const SHEET2_ID = '1fQ0lce6cSEvV_WSG9Vygu6FDonkBptc6ou1R-doQhNo'; // Second sheet
  const SHEET2_TAB = 'Form responses 1';
  
  const SHEET3_ID = '1LX2S5it81y7Eg8d1lqnx59kNhzgFFOcRPrJrLkz00Js'; // Third sheet
  const SHEET3_TAB = 'Came List';
  
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
    
    console.log('Opening Destination Sheet...');
    const destSheet = SpreadsheetApp.openById(DESTINATION_SHEET_ID).getSheetByName(DESTINATION_TAB);
    if (!destSheet) throw new Error('Destination sheet not found');
    
    // Get data from all sheets
    const data1 = getSheetData(sheet1);
    const data2 = getSheetData(sheet2);
    const data3 = getSheetData(sheet3);
    
    if (data1.length === 0 && data2.length === 0 && data3.length === 0) {
      console.log('No data found in source sheets');
      return;
    }
    
    // Get existing data from destination
    const existingData = getSheetData(destSheet);
    
    // Process the data with smart merging
    const processedData = smartMergeSheets(data1, data2, data3, existingData);
    
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
    }
    
    console.log('EduBridge Academy sheets merged successfully!');
    console.log(`Total unique records: ${processedData.combinedData.length - 1}`);
    console.log(`Records merged: ${processedData.mergedCount}`);
    console.log(`New records added: ${processedData.newRecords}`);
    
  } catch (error) {
    console.error('Error combining sheets:', error);
  }
}

function smartMergeSheets(data1, data2, data3, existingData) {
  const result = {
    combinedData: [],
    sheet2Rows: [],
    sheet3Rows: [],
    mergedCount: 0,
    newRecords: 0
  };
  
  // Create master headers from all sheets
  const allHeaders = getMasterHeaders(data1, data2, data3, existingData);
  result.combinedData.push(allHeaders);
  
  // Convert all data to record objects for easier processing
  const records1 = dataToRecords(data1);
  const records2 = dataToRecords(data2);
  const records3 = dataToRecords(data3);
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
  
  // Convert back to array format and track sheet rows
  let rowIndex = 2; // Start from row 2 (after header)
  recordMap.forEach(record => {
    const recordArray = allHeaders.map(header => {
      const value = record[header] || '';
      // Replace blank values with "Not Recorded"
      return (value === '' || value === null || value === undefined) ? 'Not Recorded' : value;
    });
    result.combinedData.push(recordArray);
    
    // Track which rows came from which sheets
    if (record.source === 'sheet2') {
      result.sheet2Rows.push(rowIndex);
    } else if (record.source === 'sheet3') {
      result.sheet3Rows.push(rowIndex);
    }
    rowIndex++;
  });
  
  return result;
}

function getMasterHeaders(data1, data2, data3, existingData) {
  const allHeaders = new Set();
  
  // Add headers from all sources
  [data1, data2, data3, existingData].forEach(data => {
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
      // Replace blank values with "Not Recorded" when creating records
      record[header] = (value === '' || value === null || value === undefined) ? 'Not Recorded' : value;
    });
    
    // Intelligent column mapping and combination
    record = smartColumnMapping(record);
    
    records.push(record);
  }
  
  return records;
}

function smartColumnMapping(record) {
  const mappedRecord = { ...record };
  
  // Combine Phone Number columns intelligently
  if (record['Whatsapp Phone Number'] || record['Phone Number']) {
    const whatsappPhone = record['Whatsapp Phone Number'] || 'Not Recorded';
    const regularPhone = record['Phone Number'] || 'Not Recorded';
    
    // If both exist and are different, combine them
    if (whatsappPhone !== 'Not Recorded' && regularPhone !== 'Not Recorded' && whatsappPhone !== regularPhone) {
      mappedRecord['Phone Number'] = `${regularPhone} | WhatsApp: ${whatsappPhone}`;
    }
    // If only WhatsApp exists, use it as main phone
    else if (whatsappPhone !== 'Not Recorded' && regularPhone === 'Not Recorded') {
      mappedRecord['Phone Number'] = whatsappPhone;
    }
    // If only regular phone exists, keep it
    else if (regularPhone !== 'Not Recorded') {
      mappedRecord['Phone Number'] = regularPhone;
    }
    
    // Remove the separate WhatsApp column since we've combined it
    delete mappedRecord['Whatsapp Phone Number'];
  }
  
  // Combine Role columns intelligently
  if (record['Current Role (If unemployed, put NIL)'] || record['Role']) {
    const currentRole = record['Current Role (If unemployed, put NIL)'] || 'Not Recorded';
    const role = record['Role'] || 'Not Recorded';
    
    // If both exist and are different, prefer the more detailed one
    if (currentRole !== 'Not Recorded' && role !== 'Not Recorded') {
      // If current role is NIL or unemployed, use the other role
      if (currentRole.toLowerCase().includes('nil') || currentRole.toLowerCase().includes('unemployed')) {
        mappedRecord['Current Role (If unemployed, put NIL)'] = role;
      }
      // If roles are different and both meaningful, combine them
      else if (currentRole.toLowerCase() !== role.toLowerCase()) {
        mappedRecord['Current Role (If unemployed, put NIL)'] = `${currentRole} | ${role}`;
      }
      // If they're similar, keep the more detailed one
      else {
        mappedRecord['Current Role (If unemployed, put NIL)'] = currentRole.length > role.length ? currentRole : role;
      }
    }
    // If only one exists, use it
    else if (currentRole !== 'Not Recorded') {
      mappedRecord['Current Role (If unemployed, put NIL)'] = currentRole;
    } else if (role !== 'Not Recorded') {
      mappedRecord['Current Role (If unemployed, put NIL)'] = role;
    }
    
    // Remove the separate Role column since we've combined it
    delete mappedRecord['Role'];
  }
  
  return mappedRecord;
}

function mergeRecords(record1, record2) {
  const merged = { ...record1 };
  
  // Merge fields, preferring non-empty values
  Object.keys(record2).forEach(key => {
    if (key !== 'source') {
      // If record1 has "Not Recorded" but record2 has real data, use record2's value
      if ((merged[key] === 'Not Recorded' || !merged[key] || merged[key].toString().trim() === '') && 
          record2[key] && record2[key] !== 'Not Recorded' && record2[key].toString().trim() !== '') {
        merged[key] = record2[key];
      }
      // If both have values and they're different, combine them (but not if one is "Not Recorded")
      else if (merged[key] && record2[key] && 
               merged[key] !== 'Not Recorded' && record2[key] !== 'Not Recorded' &&
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