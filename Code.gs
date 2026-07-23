// ============================================================
//  RKD PO OCR EXTRACTION SYSTEM — Code.gs
//  Google Apps Script Backend
//  Drive Folder : 1do_kkIqVM9BUnIR9hLlT92HelH1vjnnF
//  Sheet ID     : 1uCyRLko_G9OEOBuim8RI8uZqm5kNQrOUiDSTYlyFLXU
//  Sheet Name   : Data
// ============================================================

function getGeminiApiKey() {
  const propKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (propKey && propKey.trim() !== '') return propKey.trim();
  try {
    const encodedDefault = 'QVEuQWI4Uk42TGlwSElQV3Z5RzQ4TWt0SjhCSXQ2UFZUZWQyNXlFYkh6akR1ZHRKTEZIOVE=';
    return Utilities.newBlob(Utilities.base64Decode(encodedDefault)).getDataAsString();
  } catch (e) {
    return '';
  }
}

const CONFIG = {
  GEMINI_API_KEY: getGeminiApiKey(),
  FOLDER_ID : '1do_kkIqVM9BUnIR9hLlT92HelH1vjnnF',
  SHEET_ID  : '1uCyRLko_G9OEOBuim8RI8uZqm5kNQrOUiDSTYlyFLXU',
  SHEET_NAME: 'Data',
  DEFAULT_MAIN_FIELDS: [
    'PO Number','Order Date','Buyer / Company Name','Vendor Name','Ex-Factory Date',
    'Cancel Date','Total Order Amount','Total Cases','Ship Via','Payment Terms',
    'FOB / Port of Departure','Port of Entry / Destination','CBM','Material / Fabric','Remarks'
  ],
  DEFAULT_ITEM_FIELDS: [
    'SKU / Item #','Description','Color','Size','UPC','Quantity','Unit Price','Case Pack','Line Amount'
  ],
  MAYTAPI_PRODUCT_ID: '0d0df307-0553-4dfd-8597-e3c2fd5300eb',
  MAYTAPI_TOKEN: '54f10e32-bdf4-49cd-a464-33dc87c7c001',
  MAYTAPI_PHONE_ID: '34244'
};

// ─────────────────────────────────────────────
//  ENTRY POINT
// ─────────────────────────────────────────────
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var response = { success: false, error: "Invalid Action" };
    try {
      if (action === 'getAppConfig') {
        response = { success: true, config: getAppConfig() };
      } else if (action === 'getSavedData') {
        response = { success: true, data: getSavedData() };
      } else if (action === 'getSheetUrl') {
        response = { success: true, url: getSheetUrl() };
      } else if (action === 'getDraft') {
        response = getDraft(e.parameter.draftId);
      }
    } catch(err) {
      response = { success: false, error: err.toString() };
    }
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.draftId = (e.parameters && e.parameters.draftId && e.parameters.draftId.length > 0) ? e.parameters.draftId[0] : '';
  template.scriptUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle('RKD PO OCR System')
    .addMetaTag('viewport','width=device-width,initial-scale=1,maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var response = { success: false, error: "Invalid Action" };
  try {
    var contents = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    var action = contents.action;
    var args = contents.args || [];

    if (action === 'processMultipleFiles') {
      response = processMultipleFiles(args[0]);
    } else if (action === 'getAppConfig') {
      response = { success: true, config: getAppConfig() };
    } else if (action === 'updateFieldVisibility') {
      response = updateFieldVisibility(args[0]);
    } else if (action === 'addCustomColumn') {
      response = addCustomColumn(args[0], args[1], args[2]);
    } else if (action === 'getPasscode') {
      response = { success: true, passcode: getPasscode() };
    } else if (action === 'saveToSheet') {
      response = saveToSheet(args[0]);
    } else if (action === 'getSheetUrl') {
      response = { success: true, url: getSheetUrl() };
    } else if (action === 'saveDraft') {
      response = saveDraft(args[0]);
    } else if (action === 'getDraft') {
      response = getDraft(args[0]);
    } else if (action === 'sendWhatsAppNotification') {
      response = sendWhatsAppNotification(args[0], args[1], args[2]);
    } else if (action === 'getSavedData') {
      response = { success: true, data: getSavedData() };
    } else if (action === 'updateSingleCell') {
      response = updateSingleCell(args[0], args[1], args[2]);
    } else if (action === 'updateEntireRow') {
      response = updateEntireRow(args[0], args[1]);
    }
  } catch (err) {
    response = { success: false, error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
//  MAIN PROCESS FUNCTION (Multi-File Supported)
// ─────────────────────────────────────────────
function processMultipleFiles(filesArray) {
  try {
    const uploadedFilesData = [];
    const partsArray = [];

    // 1 · Upload all to Drive and prepare Gemini parts
    for (let i = 0; i < filesArray.length; i++) {
      const fileData = filesArray[i];
      const decodedBytes = Utilities.base64Decode(fileData.base64);
      const blob         = Utilities.newBlob(decodedBytes, fileData.type, fileData.name);
      const folder       = DriveApp.getFolderById(CONFIG.FOLDER_ID);
      const driveFile    = folder.createFile(blob);
      
      uploadedFilesData.push({
        id: driveFile.getId(),
        url: driveFile.getUrl(),
        name: fileData.name
      });

      const isExcel = fileData.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                   || fileData.type === 'application/vnd.ms-excel'
                   || fileData.name.toLowerCase().endsWith('.xlsx')
                   || fileData.name.toLowerCase().endsWith('.xls');

      if (isExcel) {
        const rawText = processExcelFile(driveFile.getId(), fileData.name);
        partsArray.push({ text: `\n\n--- DOCUMENT: ${fileData.name} ---\n${rawText}` });
      } else {
        partsArray.push({
          inlineData: {
            mimeType: fileData.type,
            data: fileData.base64
          }
        });
      }
    }

    // 2 · Extract data using Gemini API
    const extractedData = callGeminiAPI(partsArray);

    // 3 · Attach metadata (combining names and using first URL as primary link)
    extractedData.fileNames = uploadedFilesData.map(f => f.name).join(', ');
    extractedData.mainFileUrl = uploadedFilesData.length > 0 ? uploadedFilesData[0].url : '';

    return { success: true, extractedData: extractedData };

  } catch (err) {
    Logger.log('processMultipleFiles error: ' + err);
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────────
//  APP CONFIGURATION & DYNAMIC FIELDS
// ─────────────────────────────────────────────
function getAppConfig() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName('AppConfig');
  if (!sheet) {
    sheet = ss.insertSheet('AppConfig');
    sheet.hideSheet();
    
    const defaultConfig = {
      mainFields: CONFIG.DEFAULT_MAIN_FIELDS.map(f => ({ name: f, visible: true })),
      itemFields: CONFIG.DEFAULT_ITEM_FIELDS.map(f => ({ name: f, visible: true }))
    };
    
    sheet.getRange(1, 1).setValue('ConfigJSON');
    sheet.getRange(2, 1).setValue(JSON.stringify(defaultConfig));
    return defaultConfig;
  }
  
  const jsonStr = sheet.getRange(2, 1).getValue();
  if (!jsonStr) {
    const defaultConfig = {
      mainFields: CONFIG.DEFAULT_MAIN_FIELDS.map(f => ({ name: f, visible: true })),
      itemFields: CONFIG.DEFAULT_ITEM_FIELDS.map(f => ({ name: f, visible: true }))
    };
    sheet.getRange(2, 1).setValue(JSON.stringify(defaultConfig));
    return defaultConfig;
  }
  
  return JSON.parse(jsonStr);
}

function updateFieldVisibility(updatedConfig) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName('AppConfig');
    if (!sheet) return { success: false, error: 'AppConfig not found' };
    
    sheet.getRange(2, 1).setValue(JSON.stringify(updatedConfig));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getPasscode() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('Passcode');
    if (!sheet) return '';
    const values = sheet.getRange('A:A').getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim().toLowerCase() === 'passcode' && i + 1 < values.length) {
        if (values[i+1][0]) return String(values[i+1][0]).trim();
      }
    }
    // Fallback to A2
    return String(sheet.getRange('A2').getValue() || '').trim();
  } catch (err) {
    return '';
  }
}

function addCustomColumn(fieldName, section, fieldType) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    
    // 1. Update Data sheet headers
    let dataSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!dataSheet) {
      dataSheet = ss.insertSheet(CONFIG.SHEET_NAME);
    }
    
    if (dataSheet.getLastRow() === 0) {
      // If sheet is entirely empty, initialize with standard headers first
      const defaultHeaders = ['Timestamp', ...CONFIG.DEFAULT_MAIN_FIELDS, ...CONFIG.DEFAULT_ITEM_FIELDS, 'Source File'];
      dataSheet.appendRow(defaultHeaders);
      styleHeaders(dataSheet);
    }
    
    const headersRange = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn());
    const headers = headersRange.getValues()[0];
    
    if (!headers.includes(fieldName)) {
      // Append to the end, but before "Source File" if it exists at the end
      let targetCol = dataSheet.getLastColumn() + 1;
      if (headers[headers.length - 1] === 'Source File') {
        dataSheet.insertColumnBefore(targetCol - 1);
        dataSheet.getRange(1, targetCol - 1).setValue(fieldName);
        dataSheet.getRange(1, targetCol).setValue('Source File'); // Ensure Source File stays at the end
      } else {
        dataSheet.getRange(1, targetCol).setValue(fieldName);
      }
      styleHeaders(dataSheet);
    }

    // 2. Update AppConfig
    let configSheet = ss.getSheetByName('AppConfig');
    if (!configSheet) getAppConfig(); // initialize if not exists
    
    const config = getAppConfig();
    const isMain = section === 'Main';
    const targetArray = isMain ? config.mainFields : config.itemFields;
    
    if (!targetArray.find(f => f.name === fieldName)) {
      targetArray.push({ name: fieldName, visible: true, type: fieldType || 'text' });
      configSheet.getRange(2, 1).setValue(JSON.stringify(config));
    }
    
    return { success: true, newConfig: config };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────────
//  GEMINI API EXTRACTION (Replaces old Regex)
// ─────────────────────────────────────────────
function callGeminiAPI(documentParts) {
  const config = getAppConfig();
  
  // Construct dynamic JSON schema with Data Type instructions
  const getTypeInstruction = (type) => {
    if (type === 'number') return 'string (Extract ONLY the numerical value/digits. Remove any currency symbols like $ or commas)';
    if (type === 'date') return 'string (Extract and format as Short Date: dd-MMM-yyyy, e.g., 23-Jul-2026)';
    return 'string (Extract as text)';
  };

  let mainSchema = {};
  config.mainFields.filter(f => f.visible).forEach(f => {
    mainSchema[f.name] = getTypeInstruction(f.type);
  });
  
  let itemSchema = {};
  config.itemFields.filter(f => f.visible).forEach(f => {
    itemSchema[f.name] = getTypeInstruction(f.type);
  });
  
  const expectedJsonStructure = {
    ...mainSchema,
    "items": [ itemSchema ]
  };

  const apiKey = getGeminiApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;
  
  const promptText = `You are a highly skilled data extraction AI. You will be provided with one or MULTIPLE Purchase Order (PO) documents (images/pdfs/text). 
Your task is to merge the data from ALL provided documents into a single JSON object.

CRITICAL RULES FOR UNIVERSAL PO EXTRACTION:
1. FORMAT MEMORY & SPEED: If you recognize a vendor's PO format from your vast training data, immediately apply that known schema to extract data instantly and accurately without confusion.
2. UNIVERSAL ADAPTABILITY: You will process Purchase Orders from many different vendors with entirely different layouts. Do not assume any fixed format. Intelligently scan the document to find the correct data points regardless of where they are placed or how they are labeled.
3. EXHAUSTIVE ITEM EXTRACTION: Identify the core line-items table. Extract EVERY SINGLE distinct item. 
   - If the layout is complex/nested (e.g., a "Master Style" with multiple "child" sizes/colors beneath it), extract EACH child row as a SEPARATE item, inheriting prices/quantities from the master if necessary. Never skip items.
4. DATA MERGING: If multiple documents/pages are provided, merge ALL line items into a SINGLE continuous "items" array. If PO numbers differ across docs, comma-separate them.
5. IF A FIELD IS MISSING OR YOU CANNOT FIND IT, return an empty string "". Do not make up data.
6. DATE FORMATTING: Convert and format ALL extracted dates into Short Date format dd-MMM-yyyy (e.g., 23-Jul-2026).
7. Return ONLY valid JSON, exactly matching the structure below. Do not add markdown code blocks like \`\`\`json.

Required JSON Structure:
${JSON.stringify(expectedJsonStructure, null, 2)}`;

  const payload = {
    contents: [ { parts: [ { text: promptText }, ...documentParts ] } ],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const jsonResponse = JSON.parse(response.getContentText());
  
  if (jsonResponse.error) {
    throw new Error("Gemini API Error: " + jsonResponse.error.message);
  }

  try {
    let contentText = jsonResponse.candidates[0].content.parts[0].text;
    // Strip markdown JSON wrapping if present
    contentText = contentText.replace(/^```json/mi, '').replace(/```$/m, '').trim();
    let parsedData = JSON.parse(contentText);
    return formatExtractedObjectDates(parsedData, config);
  } catch (e) {
    throw new Error("Failed to parse Gemini response: " + e.message);
  }
}

// ─────────────────────────────────────────────
//  EXCEL / XLSX Processing
// ─────────────────────────────────────────────
function processExcelFile(fileId, fileName) {
  let tempSheetId = null;
  try {
    const resource = {
      title   : 'TEMP_' + fileName,
      mimeType: 'application/vnd.google-apps.spreadsheet',
      parents : [{ id: CONFIG.FOLDER_ID }]
    };
    const converted = Drive.Files.copy(resource, fileId, { convert: true });
    tempSheetId     = converted.id;

    Utilities.sleep(2000);

    const ss      = SpreadsheetApp.openById(tempSheetId);
    const sheets  = ss.getSheets();

    let allText = '';
    sheets.forEach(sheet => {
      const vals = sheet.getDataRange().getValues();
      vals.forEach(row => {
        allText += row.map(c => String(c)).join('  ') + '\n';
      });
    });

    return allText;

  } catch (err) {
    Logger.log('Excel error: ' + err);
    throw new Error('Excel processing failed. Make sure Drive API is enabled in Services. Details: ' + err.message);
  } finally {
    if (tempSheetId) {
      try { DriveApp.getFileById(tempSheetId).setTrashed(true); } catch(e) {}
    }
  }
}

// ─────────────────────────────────────────────
//  SAVE TO GOOGLE SHEET
// ─────────────────────────────────────────────
function saveToSheet(rowsData) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let   sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);

    // ── Headers ────────────────────────────────
    if (sheet.getLastRow() === 0) {
      const defaultHeaders = ['Timestamp', ...CONFIG.DEFAULT_MAIN_FIELDS, ...CONFIG.DEFAULT_ITEM_FIELDS, 'Source File'];
      sheet.appendRow(defaultHeaders);
      styleHeaders(sheet);
    }
    
    const headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const sheetHeaders = headersRange.getValues()[0];
    const appConfig = getAppConfig();

    // ── Duplicate Check ────────────────────────
    const poHeader = 'PO Number';
    const buyerHeader = 'Buyer / Company Name';
    
    const poIndex = sheetHeaders.indexOf(poHeader);
    const buyerIndex = sheetHeaders.indexOf(buyerHeader);
    
    if (poIndex !== -1 && buyerIndex !== -1 && rowsData.length > 0) {
      const incomingPO = String(rowsData[0][poHeader] || '').trim().toLowerCase();
      const incomingBuyer = String(rowsData[0][buyerHeader] || '').trim().toLowerCase();
      
      if (incomingPO !== '' || incomingBuyer !== '') {
        const existingData = sheet.getDataRange().getValues();
        for (let i = 1; i < existingData.length; i++) {
          const row = existingData[i];
          const existingPO = String(row[poIndex] || '').trim().toLowerCase();
          const existingBuyer = String(row[buyerIndex] || '').trim().toLowerCase();
          
          if (existingPO === incomingPO && existingBuyer === incomingBuyer) {
            return {
              success: false,
              error: `Duplicate Entry: PO Number "${rowsData[0][poHeader]}" for Buyer "${rowsData[0][buyerHeader]}" already exists in the database.`
            };
          }
        }
      }
    }

    // ── Append rows ────────────────────────────
    const timestampFormatted = formatLongDateTime(new Date());
    let rowsAdded   = 0;

    rowsData.forEach(data => {
      const row = [];
      sheetHeaders.forEach(header => {
        if (header === 'Timestamp') {
          row.push(timestampFormatted);
        } else if (header === 'Source File') {
          row.push(data.fileUrl ? `=HYPERLINK("${data.fileUrl}", "${data.sourceFile || ''}")` : (data.sourceFile || ''));
        } else {
          let val = data[header] || '';
          if (val && isDateKey(header, appConfig)) {
            val = formatShortDate(val);
          }
          row.push(val);
        }
      });
      sheet.appendRow(row);
      rowsAdded++;
    });

    // ── Auto-resize columns ────────────────────
    try { sheet.autoResizeColumns(1, sheet.getLastColumn()); } catch(e) {}

    return {
      success  : true,
      rowsAdded: rowsAdded,
      sheetUrl : ss.getUrl()
    };

  } catch (err) {
    Logger.log('saveToSheet error: ' + err);
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────────
//  STYLE HEADERS
// ─────────────────────────────────────────────
function styleHeaders(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if(lastCol === 0) return;
    const range = sheet.getRange(1, 1, 1, lastCol);
    range.setBackground('#1a237e');
    range.setFontColor('#ffffff');
    range.setFontWeight('bold');
    range.setFontSize(10);
    range.setHorizontalAlignment('center');
    range.setVerticalAlignment('middle');
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,1).setBackground('#0d47a1'); // Timestamp col accent
  } catch(e) {}
}

// ─────────────────────────────────────────────
//  UTILITY: Get Sheet URL
// ─────────────────────────────────────────────
function getSheetUrl() {
  try {
    return SpreadsheetApp.openById(CONFIG.SHEET_ID).getUrl();
  } catch(e) {
    return '';
  }
}

// ─────────────────────────────────────────────
//  DRAFT & WHATSAPP AUTOMATION
// ─────────────────────────────────────────────
function saveDraft(payload) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName('TempData');
    if (!sheet) {
      sheet = ss.insertSheet('TempData');
      sheet.hideSheet();
    }
    const draftId = 'DRAFT-' + new Date().getTime();
    const jsonStr = JSON.stringify(payload);
    sheet.appendRow([draftId, jsonStr, new Date()]);
    return { success: true, draftId: draftId };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getDraft(draftId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('TempData');
    if (!sheet) return { success: false, error: 'No drafts found.' };
    
    const data = sheet.getDataRange().getValues();
    for(let i = 0; i < data.length; i++){
      if(data[i][0] === draftId) {
        return { success: true, payload: JSON.parse(data[i][1]) };
      }
    }
    return { success: false, error: 'Draft not found or expired.' };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function sendWhatsAppNotification(draftId, missingFieldsArray, scriptUrl) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('WhatsappNumber');
    if(!sheet) throw new Error("WhatsappNumber tab not found in the sheet.");
    
    const data = sheet.getDataRange().getValues();
    const numbers = [];
    
    // Skip header row
    for(let i = 1; i < data.length; i++) {
      if(data[i][1]) numbers.push({ name: data[i][0] || 'Team Member', phone: data[i][1] });
    }
    
    if(numbers.length === 0) throw new Error("No phone numbers found in WhatsappNumber tab.");
    
    let draftLink = scriptUrl + (scriptUrl.includes('?') ? '&' : '?') + "draftId=" + draftId;
    
    // Shorten URL for lifetime free and fast access using is.gd
    try {
      const shortUrlApi = `https://is.gd/create.php?format=simple&url=${encodeURIComponent(draftLink)}`;
      const shortResponse = UrlFetchApp.fetch(shortUrlApi).getContentText();
      if(shortResponse && shortResponse.includes('is.gd')) {
        draftLink = shortResponse;
      }
    } catch(err) {
      Logger.log("Failed to shorten URL: " + err);
    }
    
    const url = `https://api.maytapi.com/api/${CONFIG.MAYTAPI_PRODUCT_ID}/${CONFIG.MAYTAPI_PHONE_ID}/sendMessage`;
    
    let successCount = 0;
    
    for(let i = 0; i < numbers.length; i++) {
      const p = String(numbers[i].phone).replace(/\D/g,'');
      if(p.length < 10) continue; // Skip invalid numbers
      
      const message = `🚨 *URGENT: PO Missing Data*\n\nHello ${numbers[i].name},\nA new Purchase Order was uploaded but it's missing the following critical fields:\n- ${missingFieldsArray.join('\n- ')}\n\n*Action Required:* Please click the secure link below to review the PO and fill in the missing details.\n\n🔗 ${draftLink}\n\n_RKD OCR Bot_`;
      
      const payload = {
        "to_number": (p.startsWith("91") ? p : "91" + p),
        "type": "text",
        "message": message
      };
      
      const options = {
        "method": "post",
        "headers": {
          "Content-Type": "application/json",
          "x-maytapi-key": CONFIG.MAYTAPI_TOKEN
        },
        "payload": JSON.stringify(payload)
      };
      
      try {
        UrlFetchApp.fetch(url, options);
        successCount++;
      } catch(e) {
        Logger.log("WA Send Error for " + numbers[i].name + ": " + e);
      }
    }
    
    return { success: true, sentCount: successCount };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────
//  FETCH SAVED DATA FOR DATA TABLE
// ─────────────────────────────────────────────
function getSavedData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // Only headers or empty
    
    const headers = data[0];
    const rows = [];
    const config = getAppConfig();
    
    for (let i = 1; i < data.length; i++) {
      let obj = { _id: i + 1 };
      for (let j = 0; j < headers.length; j++) {
        let val = data[i][j];
        const headerName = headers[j];
        if (headerName === 'Timestamp') {
          if (val) val = formatLongDateTime(val);
        } else if (isDateKey(headerName, config)) {
          if (val) val = formatShortDate(val);
        } else if (val instanceof Date) {
          val = formatShortDate(val);
        }
        obj[headerName] = val;
      }
      rows.push(obj);
    }
    
    return rows.reverse(); // Newest first
  } catch (e) {
    Logger.log("getSavedData error: " + e);
    return [];
  }
}

// ─────────────────────────────────────────────
//  UPDATE SINGLE CELL FROM DATA TABLE
// ─────────────────────────────────────────────
function updateSingleCell(rowNumber, headerName, newValue) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return { success: false, error: "Sheet not found" };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf(headerName);
    
    if (colIndex === -1) return { success: false, error: "Column not found: " + headerName };
    
    const config = getAppConfig();
    if (headerName === 'Timestamp' && newValue) {
      newValue = formatLongDateTime(newValue);
    } else if (isDateKey(headerName, config) && newValue) {
      newValue = formatShortDate(newValue);
    }

    sheet.getRange(rowNumber, colIndex + 1).setValue(newValue);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────────
//  UPDATE ENTIRE ROW FROM EDIT MODAL
// ─────────────────────────────────────────────
function updateEntireRow(rowNumber, updatedObj) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return { success: false, error: "Sheet not found" };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const config = getAppConfig();
    
    headers.forEach((header, colIdx) => {
      if (header && updatedObj.hasOwnProperty(header)) {
        let val = updatedObj[header];
        if (header === 'Timestamp' && val) {
          val = formatLongDateTime(val);
        } else if (isDateKey(header, config) && val) {
          val = formatShortDate(val);
        }
        sheet.getRange(rowNumber, colIdx + 1).setValue(val);
      }
    });
    
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────────
//  DATE EXTRACTION & FORMATTING UTILITIES
// ─────────────────────────────────────────────
function parseAnyDate(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  let str = String(val).trim();
  if (!str) return null;

  // Handle Excel date serial numbers (e.g. 45130 -> 2023-07-23)
  if (/^\d{5}(\.\d+)?$/.test(str)) {
    const serial = parseFloat(str);
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    const dateObj = new Date(utcValue * 1000);
    if (!isNaN(dateObj.getTime())) return dateObj;
  }

  const monthMap = {
    jan: 0, january: 0,
    feb: 1, february: 1,
    mar: 2, march: 2,
    apr: 3, april: 3,
    may: 4,
    jun: 5, june: 5,
    jul: 6, july: 6,
    aug: 7, august: 7,
    sep: 8, sept: 8, september: 8,
    oct: 9, october: 9,
    nov: 10, november: 10,
    dec: 11, december: 11
  };

  // 1. DD-MMM-YYYY or DD MMM YYYY or DD/MMM/YYYY with optional time
  let match = str.match(/^(\d{1,2})[\s\/\-\.]?([a-zA-Z]{3,9})[\s\/\-\.]?(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (match) {
    let day = parseInt(match[1], 10);
    let mStr = match[2].toLowerCase();
    let year = parseInt(match[3], 10);
    if (year < 100) year += 2000;
    let month = monthMap[mStr];
    if (month !== undefined) {
      let hh = match[4] ? parseInt(match[4], 10) : 0;
      let mm = match[5] ? parseInt(match[5], 10) : 0;
      let ss = match[6] ? parseInt(match[6], 10) : 0;
      return new Date(year, month, day, hh, mm, ss);
    }
  }

  // 2. MMM DD, YYYY or MMM DD YYYY with optional time
  match = str.match(/^([a-zA-Z]{3,9})[\s\/\-\.]?(\d{1,2}),?[\s\/\-\.]?(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (match) {
    let mStr = match[1].toLowerCase();
    let day = parseInt(match[2], 10);
    let year = parseInt(match[3], 10);
    if (year < 100) year += 2000;
    let month = monthMap[mStr];
    if (month !== undefined) {
      let hh = match[4] ? parseInt(match[4], 10) : 0;
      let mm = match[5] ? parseInt(match[5], 10) : 0;
      let ss = match[6] ? parseInt(match[6], 10) : 0;
      return new Date(year, month, day, hh, mm, ss);
    }
  }

  // 3. YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD with optional time
  match = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})(?:[\sT](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (match) {
    let year = parseInt(match[1], 10);
    let month = parseInt(match[2], 10) - 1;
    let day = parseInt(match[3], 10);
    let hh = match[4] ? parseInt(match[4], 10) : 0;
    let mm = match[5] ? parseInt(match[5], 10) : 0;
    let ss = match[6] ? parseInt(match[6], 10) : 0;
    return new Date(year, month, day, hh, mm, ss);
  }

  // 4. DD-MM-YYYY or DD/MM/YYYY or MM/DD/YYYY or DD.MM.YYYY
  match = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})(?:[\sT](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (match) {
    let p1 = parseInt(match[1], 10);
    let p2 = parseInt(match[2], 10);
    let year = parseInt(match[3], 10);
    if (year < 100) year += 2000;
    let hh = match[4] ? parseInt(match[4], 10) : 0;
    let mm = match[5] ? parseInt(match[5], 10) : 0;
    let ss = match[6] ? parseInt(match[6], 10) : 0;

    let day, month;
    if (p1 > 12) {
      day = p1;
      month = p2 - 1;
    } else if (p2 > 12) {
      day = p2;
      month = p1 - 1;
    } else {
      day = p1;
      month = p2 - 1;
    }
    if (month >= 0 && month < 12 && day >= 1 && day <= 31) {
      return new Date(year, month, day, hh, mm, ss);
    }
  }

  let d = new Date(str);
  if (!isNaN(d.getTime())) return d;

  return null;
}

function formatShortDate(val) {
  if (val === null || val === undefined || val === '') return '';
  const d = parseAnyDate(val);
  if (!d) return String(val);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  const month = months[d.getMonth()];
  const year = d.getFullYear();
  return `${day}-${month}-${year}`;
}

function formatLongDateTime(val) {
  if (val === null || val === undefined || val === '') return '';
  const d = parseAnyDate(val);
  if (!d) return String(val);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  const month = months[d.getMonth()];
  const year = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, '0');
  const mm = String(d.getMinutes()).padStart(2, '0');
  const ss = String(d.getSeconds()).padStart(2, '0');
  return `${day}-${month}-${year} ${hh}:${mm}:${ss}`;
}

function isDateKey(keyName, config) {
  if (!keyName) return false;
  const k = String(keyName).toLowerCase().trim();
  if (k === 'timestamp') return true;
  if (k.includes('date') || k.includes('dob')) return true;
  if (config) {
    const mainF = (config.mainFields || []).find(f => String(f.name).toLowerCase() === k);
    if (mainF && mainF.type === 'date') return true;
    const itemF = (config.itemFields || []).find(f => String(f.name).toLowerCase() === k);
    if (itemF && itemF.type === 'date') return true;
  }
  return false;
}

function formatExtractedObjectDates(data, config) {
  if (!data) return data;

  Object.keys(data).forEach(key => {
    if (key === 'items' && Array.isArray(data.items)) {
      data.items.forEach(item => {
        if (item && typeof item === 'object') {
          Object.keys(item).forEach(itemKey => {
            if (isDateKey(itemKey, config) || parseAnyDate(item[itemKey])) {
              if (item[itemKey]) {
                const parsed = formatShortDate(item[itemKey]);
                if (parsed) item[itemKey] = parsed;
              }
            }
          });
        }
      });
    } else if (key === 'Timestamp') {
      if (data[key]) {
        const parsed = formatLongDateTime(data[key]);
        if (parsed) data[key] = parsed;
      }
    } else if (isDateKey(key, config) || parseAnyDate(data[key])) {
      if (data[key]) {
        const parsed = formatShortDate(data[key]);
        if (parsed) data[key] = parsed;
      }
    }
  });

  return data;
}
