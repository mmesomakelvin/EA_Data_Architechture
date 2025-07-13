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
  
  const SHEET5_ID = '1Y4_QD6MbkR7cBYQFKcaP_iefU2akAnXX_ybYdcKKO8I'; // Fifth sheet - Coming Merge
  const SHEET5_TAB = 'Coming Merge';
  
  const SHEET6_ID = '1yGVCOK5-PzHc4HRNj8YrfbJ6AVqtQEt7EhejePRm0_w'; // Sixth sheet - Attend Use
  const SHEET6_TAB = 'Attend Use';
  
  const SHEET7_ID = '1l4Oa_2vNMMrX9ery8oorVwr5yRqb4MsDCnvdjTILl0U'; // Seventh sheet - Attend used
  const SHEET7_TAB = 'Attend used';

  const SHEET8_ID = '1Gpo01XcmuVfTr452socf2DmCNAuqIy_F3pbBzdkwKQ4'; // Eighth sheet
  const SHEET8_TAB = 'Sheet 9 paid';

  const SHEET9_ID = '1G2kp4gYXw8fQFjnnXkVS-OinoWU_itvtYIH6z5ZfskE'; // Ninth sheet
  const SHEET9_TAB = 'Used Sheet';
  
  const DESTINATION_SHEET_ID = '1G2b3SHGf883wtuGnA88iO0BQY4xtMS5-T8hwXvgcpps';
  const DESTINATION_TAB = 'Sheet1';
  
  try {
    // Open all sheets
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
    
    console.log('Opening Sheet 5...');
    const sheet5 = SpreadsheetApp.openById(SHEET5_ID).getSheetByName(SHEET5_TAB);
    if (!sheet5) throw new Error('Sheet 5 not found');
    
    console.log('Opening Sheet 6...');
    const sheet6 = SpreadsheetApp.openById(SHEET6_ID).getSheetByName(SHEET6_TAB);
    if (!sheet6) throw new Error('Sheet 6 not found');
    
    console.log('Opening Sheet 7...');
    const sheet7 = SpreadsheetApp.openById(SHEET7_ID).getSheetByName(SHEET7_TAB);
    if (!sheet7) throw new Error('Sheet 7 not found');

    console.log('Opening Sheet 8...');
    const sheet8 = SpreadsheetApp.openById(SHEET8_ID).getSheetByName(SHEET8_TAB);
    if (!sheet8) throw new Error('Sheet 8 not found');

    console.log('Opening Sheet 9...');
    const sheet9 = SpreadsheetApp.openById(SHEET9_ID).getSheetByName(SHEET9_TAB);
    if (!sheet9) throw new Error('Sheet 9 not found');
    
    console.log('Opening Destination Sheet...');
    const destSheet = SpreadsheetApp.openById(DESTINATION_SHEET_ID).getSheetByName(DESTINATION_TAB);
    if (!destSheet) throw new Error('Destination sheet not found');
    
    // Get data from all sheets
    const data1 = getSheetData(sheet1);
    const data2 = getSheetData(sheet2);
    const data3 = getSheetData(sheet3);
    const data4 = getSheetData(sheet4);
    const data5 = getSheetData(sheet5);
    const data6 = getSheetData(sheet6);
    const data7 = getSheetData(sheet7);
    const data8 = getSheetData(sheet8);
    const data9 = getSheetData(sheet9);
    
    if (
      data1.length === 0 && data2.length === 0 && data3.length === 0 && 
      data4.length === 0 && data5.length === 0 && data6.length === 0 && 
      data7.length === 0 && data8.length === 0 && data9.length === 0
    ) {
      console.log('No data found in source sheets');
      return;
    }
    
    // Get existing data from destination
    const existingData = getSheetData(destSheet);
    
    // Process the data with smart merging
    const processedData = smartMergeSheets(
      data1, data2, data3, data4, data5, data6, data7, data8, data9, existingData
    );
    
    // Clear and write the merged data
    destSheet.clear();
    if (processedData.combinedData.length > 0) {
      destSheet.getRange(1, 1, processedData.combinedData.length, processedData.combinedData[0].length)
        .setValues(processedData.combinedData);
      
      // Format headers
      destSheet.getRange(1, 1, 1, processedData.combinedData[0].length)
        .setFontWeight('bold')
        .setBackground('#f0f0f0');
      
      // Color code rows based on source
      if (processedData.coloredRows.length > 0) {
        processedData.coloredRows.forEach(rowInfo => {
          destSheet.getRange(rowInfo.rowIndex, 1, 1, processedData.combinedData[0].length)
            .setBackground(rowInfo.color);
        });
      }
    }
    
    console.log('EduBridge Academy sheets merged successfully!');
    console.log(`Total unique records: ${processedData.combinedData.length - 1}`);
    console.log(`Records merged: ${processedData.mergedCount}`);
    console.log(`New records added: ${processedData.newRecords}`);
    console.log(`WhatsApp-only records: ${processedData.whatsappOnlyRecords}`);
    console.log(`Attended records: ${processedData.attendedCount}`);
    console.log(`Employed records: ${processedData.employedCount}`);
    console.log(`Phone numbers cross-referenced: ${processedData.phonesCrossReferenced}`);
    console.log(`Multiple attendees: ${processedData.multipleAttendees}`);
    console.log(`Total attendance instances: ${processedData.totalAttendanceInstances}`);
    console.log(`Empty records skipped: ${processedData.emptyRecordsSkipped}`);
    
  } catch (error) {
    console.error('Error combining sheets:', error);
  }
}

function smartMergeSheets(data1, data2, data3, data4, data5, data6, data7, data8, data9, existingData) {
  const result = {
    combinedData: [],
    coloredRows: [],
    mergedCount: 0,
    newRecords: 0,
    whatsappOnlyRecords: 0,
    attendedCount: 0,
    employedCount: 0,
    phonesCrossReferenced: 0,
    multipleAttendees: 0,
    totalAttendanceInstances: 0,
    emptyRecordsSkipped: 0
  };
  
  // Email to exclude from results
  const EXCLUDED_EMAIL = 'mmesomakelvin@gmail.com';
  
  // Create master headers from all sheets and add new columns
  const allHeaders = getMasterHeaders(
    data1, data2, data3, data4, data5, data6, data7, data8, data9, existingData
  );
  
  if (!allHeaders.includes('Attended')) allHeaders.push('Attended');
  if (!allHeaders.includes('Employed')) allHeaders.push('Employed');
  if (!allHeaders.includes('Date Attended')) allHeaders.push('Date Attended');
  if (!allHeaders.includes('Attendance Count')) allHeaders.push('Attendance Count');
  
  result.combinedData.push(allHeaders);
  
  // Build email-phone mappings from ALL sheets
  const emailToPhoneMap = new Map();
  const phoneToEmailMap = new Map();
  
  function buildPhoneMappings(data, sheetName) {
    const records = dataToRecords(data);
    records.forEach(record => {
      const email = record['Email Address'] ? record['Email Address'].toString().trim().toLowerCase() : null;
      const phone = getRecordPhone(record);
      if (email && email !== EXCLUDED_EMAIL && phone) {
        emailToPhoneMap.set(email, phone);
        phoneToEmailMap.set(phone, email);
      }
    });
  }
  
  buildPhoneMappings(data1, 'Sheet1');
  buildPhoneMappings(data2, 'Sheet2');
  buildPhoneMappings(data3, 'Sheet3');
  buildPhoneMappings(data4, 'Sheet4');
  buildPhoneMappings(data5, 'Sheet5');
  buildPhoneMappings(data6, 'Sheet6');
  buildPhoneMappings(data7, 'Sheet7');
  buildPhoneMappings(data8, 'Sheet8');
  buildPhoneMappings(data9, 'Sheet9');
  buildPhoneMappings(existingData, 'Existing');
  
  // Convert all data to record objects
  const records1 = dataToRecords(data1);
  const records2 = dataToRecords(data2);
  const records3 = dataToRecords(data3);
  const records4 = dataToRecords(data4);
  const records5 = dataToRecords(data5);
  const records6 = dataToRecords(data6);
  const records7 = dataToRecords(data7);
  const records8 = dataToRecords(data8);
  const records9 = dataToRecords(data9);
  const existingRecords = dataToRecords(existingData);
  
  // Registration records (main base): sheets 1, 2, existing
  const baseRecordMap = new Map();
  const phoneToEmailMap2 = new Map();
  const attendanceRecords = [];
  
  function getRecordKey(record) {
    const email = record['Email Address'] ? record['Email Address'].toString().trim().toLowerCase() : null;
    const phone = getRecordPhone(record);
    if (email && email !== EXCLUDED_EMAIL) {
      return `email:${email}`;
    } else if (phone) {
      return `phone:${phone}`;
    }
    return null;
  }
  
  function crossReferenceRecord(record) {
    const email = record['Email Address'] ? record['Email Address'].toString().trim().toLowerCase() : null;
    let currentPhone = getRecordPhone(record);
    if (email && email !== EXCLUDED_EMAIL && !currentPhone && emailToPhoneMap.has(email)) {
      const mappedPhone = emailToPhoneMap.get(email);
      record['Whatsapp Phone Number'] = mappedPhone;
      result.phonesCrossReferenced++;
      currentPhone = mappedPhone;
    }
    if (currentPhone && !email && phoneToEmailMap.has(currentPhone)) {
      const mappedEmail = phoneToEmailMap.get(currentPhone);
      record['Email Address'] = mappedEmail;
      result.phonesCrossReferenced++;
    }
    return record;
  }
  
  function hasAnyMeaningfulData(record) {
    return Object.keys(record).some(key => {
      if (key === 'source' || key === 'color') return false;
      const value = record[key];
      return value && value !== 'Not Recorded' && value !== 0 && value.toString().trim() !== '';
    });
  }
  
  function addOrMergeBaseRecord(record, source) {
    record = crossReferenceRecord(record);
    if (!hasAnyMeaningfulData(record)) {
      result.emptyRecordsSkipped++;
      return;
    }
    const recordKey = getRecordKey(record);
    if (!recordKey) return;
    const email = record['Email Address'] ? record['Email Address'].toString().trim().toLowerCase() : null;
    if (email === EXCLUDED_EMAIL) return;
    const phone = getRecordPhone(record);
    let existingKey = recordKey;
    let foundExisting = false;
    if (phone && phoneToEmailMap2.has(phone)) {
      const linkedEmail = phoneToEmailMap2.get(phone);
      existingKey = `email:${linkedEmail}`;
      foundExisting = baseRecordMap.has(existingKey);
    } else if (baseRecordMap.has(recordKey)) {
      foundExisting = true;
    }
    if (foundExisting && baseRecordMap.has(existingKey)) {
      const existingRecord = baseRecordMap.get(existingKey);
      const mergedRecord = mergeRecords(existingRecord, record);
      mergedRecord.source = existingRecord.source;
      baseRecordMap.set(existingKey, mergedRecord);
      result.mergedCount++;
    } else {
      record.source = source;
      baseRecordMap.set(recordKey, record);
      if (phone && email && email !== EXCLUDED_EMAIL) {
        phoneToEmailMap2.set(phone, email);
      }
      result.newRecords++;
    }
  }
  
  existingRecords.forEach(record => addOrMergeBaseRecord(record, 'existing'));
  records1.forEach(record => addOrMergeBaseRecord(record, 'sheet1'));
  records2.forEach(record => addOrMergeBaseRecord(record, 'sheet2'));
  
  // Attendance sheets (3-9)
  function processAttendanceRecords(records, sheetSource, dateAttended, sheetColor) {
    records.forEach(record => {
      record = crossReferenceRecord(record);
      const email = record['Email Address'] ? record['Email Address'].toString().trim().toLowerCase() : null;
      if (email === EXCLUDED_EMAIL) return;
      const phone = getRecordPhone(record);
      const recordKey = getRecordKey(record);
      if (!recordKey) return;
      record['Date Attended'] = dateAttended;
      record['Attended'] = 1;
      record.source = sheetSource;
      record.color = sheetColor;
      let baseRecord = null;
      if (email && baseRecordMap.has(`email:${email}`)) {
        baseRecord = baseRecordMap.get(`email:${email}`);
      } else if (phone && phoneToEmailMap2.has(phone)) {
        const linkedEmail = phoneToEmailMap2.get(phone);
        if (baseRecordMap.has(`email:${linkedEmail}`)) {
          baseRecord = baseRecordMap.get(`email:${linkedEmail}`);
        }
      }
      let attendanceRecord;
      if (baseRecord) {
        attendanceRecord = mergeRecords(baseRecord, record);
        attendanceRecord.source = sheetSource;
        attendanceRecord.color = sheetColor;
        attendanceRecord['Date Attended'] = dateAttended;
        attendanceRecord['Attended'] = 1;
      } else {
        attendanceRecord = { ...record };
        result.whatsappOnlyRecords++;
      }
      attendanceRecords.push(attendanceRecord);
      result.totalAttendanceInstances++;
    });
  }
  processAttendanceRecords(records3, 'sheet3', '1 Jan 2025', '#D5F0E1');
  processAttendanceRecords(records4, 'sheet4', '1 May 2025', '#FFE4B5');
  processAttendanceRecords(records5, 'sheet5', '1 Nov 2024', '#FFD700');
  processAttendanceRecords(records6, 'sheet6', '1 Aug 2024', '#E6E6FA');
  processAttendanceRecords(records7, 'sheet7', '1 Apr 2024', '#F0E68C');
  processAttendanceRecords(records8, 'sheet8', '1 Jan 2024', '#B0E0E6');
  processAttendanceRecords(records9, 'sheet9', '1 Aug 2023', '#D0F0F6'); // Sheet 9 uses light teal

  const attendanceCountMap = new Map();
  attendanceRecords.forEach(record => {
    const email = record['Email Address'] ? record['Email Address'].toString().trim().toLowerCase() : null;
    const phone = getRecordPhone(record);
    const key = email && email !== EXCLUDED_EMAIL ? email : phone;
    if (key) {
      attendanceCountMap.set(key, (attendanceCountMap.get(key) || 0) + 1);
    }
  });
  attendanceCountMap.forEach((count, key) => {
    if (count > 1) result.multipleAttendees++;
  });
  let rowIndex = 2;
  baseRecordMap.forEach((record, key) => {
    const email = record['Email Address'] ? record['Email Address'].toString().trim().toLowerCase() : null;
    const phone = getRecordPhone(record);
    if (!hasAnyMeaningfulData(record)) {
      result.emptyRecordsSkipped++;
      return;
    }
    const attendanceKey = email && email !== EXCLUDED_EMAIL ? email : phone;
    const hasAttendance = attendanceKey && attendanceCountMap.has(attendanceKey);
    if (!hasAttendance) {
      record['Attended'] = 0;
      record['Date Attended'] = '';
      record['Attendance Count'] = 0;
      const companyField = record['Current Company (If unemployed, put NIL)'] || '';
      const companyStr = companyField.toString().trim().toUpperCase();
      const employed = (companyStr && 
                       companyStr !== 'NIL' && 
                       companyStr !== 'NOT RECORDED' && 
                       companyStr !== '') ? 1 : 0;
      record['Employed'] = employed;
      if (employed === 1) result.employedCount++;
      const recordArray = allHeaders.map(header => {
        const value = record[header] || '';
        return getDefaultValue(value, header);
      });
      result.combinedData.push(recordArray);
      let color = null;
      if (record.source === 'sheet2') color = '#E1D5F0'; // Light purple
      if (color) {
        result.coloredRows.push({
          rowIndex: rowIndex,
          color: color
        });
      }
      rowIndex++;
    }
  });
  attendanceRecords.forEach(record => {
    const email = record['Email Address'] ? record['Email Address'].toString().trim().toLowerCase() : null;
    const phone = getRecordPhone(record);
    const attendanceKey = email && email !== EXCLUDED_EMAIL ? email : phone;
    const attendanceCount = attendanceKey ? (attendanceCountMap.get(attendanceKey) || 1) : 1;
    record['Attendance Count'] = attendanceCount;
    const companyField = record['Current Company (If unemployed, put NIL)'] || '';
    const companyStr = companyField.toString().trim().toUpperCase();
    const employed = (companyStr && 
                     companyStr !== 'NIL' && 
                     companyStr !== 'NOT RECORDED' && 
                     companyStr !== '') ? 1 : 0;
    record['Employed'] = employed;
    if (employed === 1) result.employedCount++;
    result.attendedCount++;
    const recordArray = allHeaders.map(header => {
      const value = record[header] || '';
      return getDefaultValue(value, header);
    });
    result.combinedData.push(recordArray);
    let color = record.color;
    if (attendanceCount > 1) color = '#FF6B6B'; // Red for multiple attendances
    result.coloredRows.push({
      rowIndex: rowIndex,
      color: color
    });
    rowIndex++;
  });
  return result;
}

function getRecordPhone(record) {
  if (record['Whatsapp Phone Number']) {
    const phone = normalizePhoneNumber(record['Whatsapp Phone Number'].toString());
    if (phone) return phone;
  }
  return null;
}

function normalizePhoneNumber(phone) {
  if (!phone) return null;
  const cleaned = phone.toString().replace(/\D/g, '');
  if (cleaned.length < 10) return null;
  if (cleaned.startsWith('234')) {
    return cleaned;
  } else if (cleaned.startsWith('0') && cleaned.length === 11) {
    return '234' + cleaned.substring(1);
  } else if (cleaned.length === 10) {
    return '234' + cleaned;
  }
  return cleaned;
}

function getMasterHeaders(data1, data2, data3, data4, data5, data6, data7, data8, data9, existingData) {
  const allHeaders = new Set();
  [data1, data2, data3, data4, data5, data6, data7, data8, data9, existingData].forEach(data => {
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
      record[header] = getDefaultValue(value, header);
    });
    records.push(record);
  }
  return records;
}

function getDefaultValue(value, header) {
  const isEmpty = (value === '' || value === null || value === undefined);
  const financialColumns = ['Payment', 'Expected', 'Amount owed'];
  const isFinancialColumn = financialColumns.some(col => 
    header && header.toLowerCase().includes(col.toLowerCase())
  );
  const isDiscountColumn = header && header.toLowerCase().includes('discount');
  const isAttendedColumn = header && header.toLowerCase() === 'attended';
  const isEmployedColumn = header && header.toLowerCase() === 'employed';
  const isAttendanceCountColumn = header && header.toLowerCase() === 'attendance count';
  if (isEmpty) {
    if (isDiscountColumn) {
      return 'Not Recorded';
    } else if (isFinancialColumn || isAttendedColumn || isEmployedColumn || isAttendanceCountColumn) {
      return 0;
    } else {
      return 'Not Recorded';
    }
  }
  if (isDiscountColumn && !isEmpty) {
    const valueStr = value.toString().trim();
    if (valueStr) {
      if (valueStr.includes('%')) {
        return valueStr;
      }
      const numValue = parseFloat(valueStr);
      if (!isNaN(numValue)) {
        if (numValue >= 0 && numValue <= 1) {
          return (numValue * 100) + '%';
        } else {
          return numValue + '%';
        }
      }
    }
  }
  return value;
}

function mergeRecords(record1, record2) {
  const merged = { ...record1 };
  Object.keys(record2).forEach(key => {
    if (key !== 'source' && key !== 'color') {
      const isIndustryColumn = key && (
        key.toLowerCase().includes('industry you work in') || 
        key.toLowerCase().includes('industry') && key.toLowerCase().includes('work')
      );
      if (isIndustryColumn && record2[key] && record2[key] !== 'Not Recorded' && record2[key].toString().trim() !== '') {
        merged[key] = record2[key];
      }
      else if (key === 'Date Attended' && record2[key] && record2[key] !== 'Not Recorded' && record2[key].toString().trim() !== '') {
        merged[key] = record2[key];
      }
      else if ((merged[key] === 'Not Recorded' || merged[key] === 0 || !merged[key] || merged[key].toString().trim() === '') && 
          record2[key] && record2[key] !== 'Not Recorded' && record2[key] !== 0 && record2[key].toString().trim() !== '') {
        merged[key] = record2[key];
      }
      else if (!isIndustryColumn && merged[key] && record2[key] && 
               merged[key] !== 'Not Recorded' && record2[key] !== 'Not Recorded' &&
               merged[key] !== 0 && record2[key] !== 0 &&
               merged[key].toString().trim() !== record2[key].toString().trim() &&
               merged[key].toString().trim() !== '' &&
               record2[key].toString().trim() !== '') {
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
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'combineSheets') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('combineSheets')
    .timeBased()
    .everyHours(1)
    .create();
  console.log('Automatic updates set up successfully!');
}
