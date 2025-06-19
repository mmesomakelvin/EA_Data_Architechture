// Google Apps Script to combine your two EduBridge Academy sheets (INCREMENTAL UPDATES)
function combineSheets() {
  // CONFIGURATION - Your actual sheet details
  const SHEET1_ID = '1TBLTPiDlpjGV5LPb2ahZ0wWUvoLauhHOW9-_cn-ZUyI'; // Data Analyst program sheet
  const SHEET1_TAB = 'Form responses 1'; // Common Google Forms tab name
  
  const SHEET2_ID = '1fQ0lce6cSEvV_WSG9Vygu6FDonkBptc6ou1R-doQhNo'; // Second sheet
  const SHEET2_TAB = 'Form responses 1'; // Common Google Forms tab name
  
  // Your destination sheet ID
  const DESTINATION_SHEET_ID = '1G2b3SHGf883wtuGnA88iO0BQY4xtMS5-T8hwXvgcpps';
  const DESTINATION_TAB = 'Sheet1'; // Default tab name
  
  try {
    // Open the sheets
    const sheet1 = SpreadsheetApp.openById(SHEET1_ID).getSheetByName(SHEET1_TAB);
    const sheet2 = SpreadsheetApp.openById(SHEET2_ID).getSheetByName(SHEET2_TAB);
    const destSheet = SpreadsheetApp.openById(DESTINATION_SHEET_ID).getSheetByName(DESTINATION_TAB);
    
    // Get existing data from destination to avoid duplicates
    const existingData = getSheetData(destSheet);
    const existingEmails = getExistingEmails(existingData);
    
    // Get data from both sheets
    const data1 = getSheetData(sheet1);
    const data2 = getSheetData(sheet2);
    
    // Only process if we have data
    if (data1.length === 0 && data2.length === 0) {
      console.log('No data found in source sheets');
      return;
    }
    
    // If destination is empty, do initial setup
    if (existingData.length === 0) {
      console.log('Initial setup - combining all data');
      const combinedData = combineAllData(data1, data2);
      writeDataWithColorCoding(destSheet, combinedData, data1.length - 1);
    } else {
      console.log('Incremental update - adding new records only');
      addNewRecordsOnly(destSheet, data1, data2, existingEmails, existingData);
    }
    
  } catch (error) {
    console.error('Error combining sheets:', error);
  }
}

function getExistingEmails(existingData) {
  if (existingData.length <= 1) return new Set();
  
  const headers = existingData[0];
  const emailIndex = headers.indexOf('Email Address');
  
  if (emailIndex === -1) return new Set();
  
  const emails = new Set();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][emailIndex]) {
      emails.add(existingData[i][emailIndex].toString().trim().toLowerCase());
    }
  }
  return emails;
}

function addNewRecordsOnly(destSheet, data1, data2, existingEmails, existingData) {
  const headers = existingData[0];
  const newRecords = [];
  
  // Check data1 for new records
  const newFromSheet1 = getNewRecords(data1, existingEmails, headers);
  const newFromSheet2 = getNewRecords(data2, existingEmails, headers);
  
  if (newFromSheet1.length === 0 && newFromSheet2.length === 0) {
    console.log('No new records found');
    return;
  }
  
  // Add new records
  const lastRow = destSheet.getLastRow();
  
  // Add records from sheet 1
  if (newFromSheet1.length > 0) {
    const startRow = lastRow + 1;
    destSheet.getRange(startRow, 1, newFromSheet1.length, newFromSheet1[0].length).setValues(newFromSheet1);
    console.log(`Added ${newFromSheet1.length} new records from Data Analyst program`);
  }
  
  // Add records from sheet 2 with color coding
  if (newFromSheet2.length > 0) {
    const startRow = destSheet.getLastRow() + 1;
    destSheet.getRange(startRow, 1, newFromSheet2.length, newFromSheet2[0].length).setValues(newFromSheet2);
    
    // Color code the first row of sheet 2 data light purple
    destSheet.getRange(startRow, 1, 1, newFromSheet2[0].length).setBackground('#E1D5F0');
    
    console.log(`Added ${newFromSheet2.length} new records from second sheet`);
  }
}

function getNewRecords(sourceData, existingEmails, masterHeaders) {
  if (sourceData.length <= 1) return [];
  
  const sourceHeaders = sourceData[0];
  const emailIndex = sourceHeaders.indexOf('Email Address');
  
  if (emailIndex === -1) return [];
  
  const newRecords = [];
  
  for (let i = 1; i < sourceData.length; i++) {
    const email = sourceData[i][emailIndex];
    if (email && !existingEmails.has(email.toString().trim().toLowerCase())) {
      // Map this record to master headers format
      const mappedRecord = mapRecordToHeaders(sourceData[i], sourceHeaders, masterHeaders);
      newRecords.push(mappedRecord);
    }
  }
  
  return newRecords;
}

function mapRecordToHeaders(record, sourceHeaders, masterHeaders) {
  const mappedRecord = new Array(masterHeaders.length).fill('');
  
  for (let i = 0; i < sourceHeaders.length; i++) {
    const headerName = sourceHeaders[i];
    const masterIndex = masterHeaders.indexOf(headerName);
    if (masterIndex !== -1) {
      mappedRecord[masterIndex] = record[i] || '';
    }
  }
  
  return mappedRecord;
}

function combineAllData(data1, data2) {
  if (data1.length === 0 && data2.length === 0) {
    return [];
  }
  
  if (data1.length === 0) {
    return data2;
  }
  
  if (data2.length === 0) {
    return data1;
  }
  
  // Get headers from both sheets
  const headers1 = data1[0];
  const headers2 = data2[0];
  
  // Create master header list (all unique headers)
  const allHeaders = [...new Set([...headers1, ...headers2])];
  
  // Convert data to objects for easier manipulation
  const objects1 = dataToObjects(data1);
  const objects2 = dataToObjects(data2);
  
  // Combine all objects
  const allObjects = [...objects1, ...objects2];
  
  // Convert back to array format with consistent columns
  return objectsToData(allObjects, allHeaders);
}

function writeDataWithColorCoding(destSheet, combinedData, sheet1RecordCount) {
  if (combinedData.length === 0) return;
  
  // Write all data
  destSheet.getRange(1, 1, combinedData.length, combinedData[0].length).setValues(combinedData);
  
  // Format headers
  destSheet.getRange(1, 1, 1, combinedData[0].length)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  
  // Color code where sheet 2 starts (light purple)
  const sheet2StartRow = sheet1RecordCount + 2; // +1 for header, +1 for next row
  if (sheet2StartRow <= combinedData.length) {
    destSheet.getRange(sheet2StartRow, 1, 1, combinedData[0].length)
      .setBackground('#E1D5F0');
  }
  
  console.log(`Initial setup complete:`);
  console.log(`- ${sheet1RecordCount} records from Data Analyst program`);
  console.log(`- ${combinedData.length - 1 - sheet1RecordCount} records from second sheet`);
  console.log(`- Sheet 2 starts at row ${sheet2StartRow} (light purple)`);
}

function getSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow === 0 || lastCol === 0) {
    return [];
  }
  
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

function combineData(data1, data2) {
  // This function is kept for compatibility but not used in incremental updates
  return combineAllData(data1, data2);
}

function dataToObjects(data) {
  if (data.length <= 1) {
    return [];
  }
  
  const headers = data[0];
  const objects = [];
  
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j] || '';
    }
    objects.push(obj);
  }
  
  return objects;
}

function objectsToData(objects, headers) {
  if (objects.length === 0) {
    return [headers];
  }
  
  const result = [headers];
  
  objects.forEach(obj => {
    const row = headers.map(header => obj[header] || '');
    result.push(row);
  });
  
  return result;
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

// Alternative: Manual trigger when source sheets change
function setupOnEditTrigger() {
  // This would need to be set up on each source sheet individually
  // You would add this as an installable trigger on each source sheet
  ScriptApp.newTrigger('combineSheets')
    .onEdit()
    .create();
}