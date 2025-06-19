// Google Apps Script to combine your two EduBridge Academy sheets
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
    
    // Clear destination sheet
    destSheet.clear();
    
    // Get data from both sheets
    const data1 = getSheetData(sheet1);
    const data2 = getSheetData(sheet2);
    
    // Combine the data
    const combinedData = combineData(data1, data2);
    
    // Write to destination sheet
    if (combinedData.length > 0) {
      destSheet.getRange(1, 1, combinedData.length, combinedData[0].length).setValues(combinedData);
      
      // Format headers
      destSheet.getRange(1, 1, 1, combinedData[0].length)
        .setFontWeight('bold')
        .setBackground('#f0f0f0');
    }
    
    console.log('EduBridge Academy sheets combined successfully!');
    console.log(`Combined ${data1.length - 1} records from Data Analyst program`);
    console.log(`Combined ${data2.length - 1} records from second sheet`);
    console.log(`Total records: ${combinedData.length - 1}`);
    
  } catch (error) {
    console.error('Error combining sheets:', error);
  }
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
