// Google Apps Script code for Seaman Agency Database
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getData') {
    return getDataFromSheet();
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    error: 'Invalid action'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  
  try {
    if (action === 'addProfile') {
      return addProfileToSheet(data.data);
    } else if (action === 'syncData') {
      return syncDataToSheet(data.data);
    } else if (action === 'exportData') {
      return exportDataToSheet(data.data);
    } else {
      throw new Error('Invalid action');
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getDataFromSheet() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    
    // Skip header row
    const profiles = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      profiles.push({
        id: row[0],
        name: row[1],
        rank: row[2],
        nrc: row[3],
        fatherName: row[4],
        motherName: row[5],
        phone: row[6],
        birthdate: row[7],
        address: row[8],
        distinguishingMark: row[9],
        fees: row[10],
        feesPercentage: row[11],
        timestamp: row[12]
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: profiles
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
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
      profile.id,
      profile.name,
      profile.rank,
      profile.nrc,
      profile.fatherName,
      profile.motherName,
      profile.phone,
      profile.birthdate,
      profile.address,
      profile.distinguishingMark,
      profile.fees,
      profile.feesPercentage,
      profile.timestamp || new Date().toISOString()
    ]]);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Profile added successfully'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function syncDataToSheet(profiles) {
  try {
    const sheet = getSheet();
    
    // Clear existing data (optional - depends on your needs)
    // sheet.clearContents();
    
    // Add headers
    sheet.getRange(1, 1, 1, 13).setValues([[
      'ID', 'Name', 'Rank', 'NRC', 'Father Name', 'Mother Name', 
      'Phone', 'Birthdate', 'Address', 'Distinguishing Mark', 
      'Fees', 'Fees Percentage', 'Timestamp'
    ]]);
    
    // Add all profiles
    const data = profiles.map(profile => [
      profile.id,
      profile.name,
      profile.rank,
      profile.nrc,
      profile.fatherName,
      profile.motherName,
      profile.phone,
      profile.birthdate,
      profile.address,
      profile.distinguishingMark,
      profile.fees,
      profile.feesPercentage,
      profile.timestamp || new Date().toISOString()
    ]);
    
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, 13).setValues(data);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: `Synced ${profiles.length} profiles successfully`
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function exportDataToSheet(data) {
  // This function can be customized based on your export needs
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'Export completed successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const spreadsheetId = 'https://docs.google.com/spreadsheets/d/1DT2SFMHsQauE1I-Gp7bi9-hj-OVJYGpbMEx98Zz5O5g/edit?usp=sharing'; // Replace with your Google Sheet ID
  const sheetName = 'SeamanProfiles'; // Your sheet name
  
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  return sheet;
}