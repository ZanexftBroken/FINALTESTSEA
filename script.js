// Seaman Agency Database - New Version
// Created with your new Web URL

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let result;
  
  try {
    if (e.postData) {
      // POST request
      const data = JSON.parse(e.postData.contents);
      const action = data.action;
      
      if (action === 'addProfile') {
        result = addProfileToSheet(data.data);
      } else if (action === 'syncData') {
        result = syncDataToSheet(data.data);
      } else if (action === 'exportData') {
        result = exportDataToSheet(data.data);
      } else {
        throw new Error('Invalid action: ' + action);
      }
    } else {
      // GET request
      const action = e.parameter.action;
      
      if (action === 'getData') {
        result = getDataFromSheet();
      } else if (action === 'test') {
        result = { success: true, message: 'API is working!' };
      } else {
        throw new Error('Invalid action: ' + action);
      }
    }
    
    return createResponse(result);
      
  } catch (error) {
    return createResponse({
      success: false,
      error: error.message
    });
  }
}

function createResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  output.setHeaders({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  });
  return output;
}

function getDataFromSheet() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    
    // Skip header row
    const profiles = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Check if row has data
      if (row[0] && row[0].toString().trim() !== '') {
        profiles.push({
          id: row[0] || '',
          name: row[1] || '',
          rank: row[2] || '',
          nrc: row[3] || '',
          fatherName: row[4] || '',
          motherName: row[5] || '',
          phone: row[6] || '',
          birthdate: row[7] || '',
          address: row[8] || '',
          distinguishingMark: row[9] || '',
          fees: row[10] || '',
          feesPercentage: row[11] || '',
          timestamp: row[12] || ''
        });
      }
    }
    
    return {
      success: true,
      data: profiles,
      count: profiles.length
    };
  } catch (error) {
    return {
      success: false,
      error: 'Get data error: ' + error.message
    };
  }
}

function addProfileToSheet(profile) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    // Add headers if sheet is empty
    if (lastRow === 0) {
      sheet.getRange(1, 1, 1, 13).setValues([[
        'ID', 'Name', 'Rank', 'NRC', 'Father Name', 'Mother Name', 
        'Phone', 'Birthdate', 'Address', 'Distinguishing Mark', 
        'Fees', 'Fees Percentage', 'Timestamp'
      ]]);
    }
    
    // Add new profile
    sheet.getRange(lastRow + 1, 1, 1, 13).setValues([[
      profile.id || Date.now(),
      profile.name || '',
      profile.rank || '',
      profile.nrc || '',
      profile.fatherName || '',
      profile.motherName || '',
      profile.phone || '',
      profile.birthdate || '',
      profile.address || '',
      profile.distinguishingMark || '',
      profile.fees || '0',
      profile.feesPercentage || '0',
      profile.timestamp || new Date().toISOString()
    ]]);
    
    return {
      success: true,
      message: 'Profile added successfully',
      id: profile.id
    };
  } catch (error) {
    return {
      success: false,
      error: 'Add profile error: ' + error.message
    };
  }
}

function syncDataToSheet(profiles) {
  try {
    const sheet = getSheet();
    
    // Clear existing data and add headers
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 13).setValues([[
      'ID', 'Name', 'Rank', 'NRC', 'Father Name', 'Mother Name', 
      'Phone', 'Birthdate', 'Address', 'Distinguishing Mark', 
      'Fees', 'Fees Percentage', 'Timestamp'
    ]]);
    
    // Add all profiles
    const data = profiles.map(profile => [
      profile.id || Date.now(),
      profile.name || '',
      profile.rank || '',
      profile.nrc || '',
      profile.fatherName || '',
      profile.motherName || '',
      profile.phone || '',
      profile.birthdate || '',
      profile.address || '',
      profile.distinguishingMark || '',
      profile.fees || '0',
      profile.feesPercentage || '0',
      profile.timestamp || new Date().toISOString()
    ]);
    
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, 13).setValues(data);
    }
    
    return {
      success: true,
      message: `Synced ${profiles.length} profiles successfully`
    };
  } catch (error) {
    return {
      success: false,
      error: 'Sync error: ' + error.message
    };
  }
}

function exportDataToSheet(data) {
  try {
    const sheet = getSheet();
    
    // Clear and export all data
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 13).setValues([[
      'ID', 'Name', 'Rank', 'NRC', 'Father Name', 'Mother Name', 
      'Phone', 'Birthdate', 'Address', 'Distinguishing Mark', 
      'Fees', 'Fees Percentage', 'Timestamp'
    ]]);
    
    const exportData = data.map(profile => [
      profile.id || Date.now(),
      profile.name || '',
      profile.rank || '',
      profile.nrc || '',
      profile.fatherName || '',
      profile.motherName || '',
      profile.phone || '',
      profile.birthdate || '',
      profile.address || '',
      profile.distinguishingMark || '',
      profile.fees || '0',
      profile.feesPercentage || '0',
      profile.timestamp || new Date().toISOString()
    ]);
    
    if (exportData.length > 0) {
      sheet.getRange(2, 1, exportData.length, 13).setValues(exportData);
    }
    
    return {
      success: true,
      message: `Exported ${data.length} profiles successfully`
    };
  } catch (error) {
    return {
      success: false,
      error: 'Export error: ' + error.message
    };
  }
}

function getSheet() {
  // Your Spreadsheet ID - Same as before
  const spreadsheetId = '1DT2SFMHsQauE1I-Gp7bi9-hj-OVJYGpbMEx98Zz5O5g';
  
  const sheetName = 'SeamanProfiles';
  
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    // Add headers to new sheet
    sheet.getRange(1, 1, 1, 13).setValues([[
      'ID', 'Name', 'Rank', 'NRC', 'Father Name', 'Mother Name', 
      'Phone', 'Birthdate', 'Address', 'Distinguishing Mark', 
      'Fees', 'Fees Percentage', 'Timestamp'
    ]]);
  }
  
  return sheet;
}

// Test function to verify setup
function testConnection() {
  return {
    success: true,
    message: 'Seaman Agency Database API is working!',
    timestamp: new Date().toISOString(),
    sheet: getSheet().getName()
  };
}
