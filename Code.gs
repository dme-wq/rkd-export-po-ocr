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
    'PO Number','Order Date','Buyer / Company Name','Customer Address','Vendor Name','Ex-Factory Date',
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

  var rawHtml = UrlFetchApp.fetch('https://raw.githubusercontent.com/dme-wq/rkd-export-po-ocr/main/index.html').getContentText();
  var template = HtmlService.createTemplate(rawHtml);
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
    } else if (action === 'generatePI') {
      response = generatePI(args[0]);
    } else if (action === 'removeColumn') {
      response = removeColumn(args[0], args[1]);
    } else if (action === 'sendPIWhatsAppNotification') {
      response = sendPIWhatsAppNotification(args[0], args[1], args[2], args[3]);
    } else if (action === 'bulkUpdatePO') {
      response = bulkUpdatePO(args[0], args[1], args[2], args[3]);
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

function removeColumn(fieldName, section) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const configSheet = ss.getSheetByName('AppConfig');
    if (!configSheet) return { success: false, error: 'AppConfig not found' };

    const config = getAppConfig();
    const isMain = section === 'Main';
    const targetArray = isMain ? config.mainFields : config.itemFields;

    // Protect system default fields from deletion
    const protectedMain = CONFIG.DEFAULT_MAIN_FIELDS;
    const protectedItem = CONFIG.DEFAULT_ITEM_FIELDS;
    const protectedFields = isMain ? protectedMain : protectedItem;

    if (protectedFields.includes(fieldName)) {
      return { success: false, error: `"${fieldName}" is a system default field and cannot be deleted. You can only hide it using the toggle.` };
    }

    const newArray = targetArray.filter(f => f.name !== fieldName);
    if (isMain) {
      config.mainFields = newArray;
    } else {
      config.itemFields = newArray;
    }

    configSheet.getRange(2, 1).setValue(JSON.stringify(config));
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
  const getTypeInstruction = (type, fieldName) => {
    if (type === 'number') return 'string (Extract ONLY the numerical value/digits. Remove any currency symbols like $ or commas)';
    if (type === 'date') return 'string (Extract and format as Short Date: dd-MMM-yyyy, e.g., 23-Jul-2026)';
    if (fieldName && (fieldName.toLowerCase().includes('description') || fieldName.toLowerCase().includes('item name') || fieldName.toLowerCase().includes('product'))) {
      return 'string (CRITICAL: Extract the EXACT, FULL, ORIGINAL product description text as-it-is from the PO document without removing, cutting, or truncating any words, color, or size information! e.g., keep "Nudie Rudie Bath Mat Mini - Dahlia" exactly as written)';
    }
    if (fieldName && (fieldName.toLowerCase().includes('color') || fieldName.toLowerCase().includes('colour'))) {
      return 'string (CRITICAL: Extract color name. If Color is not in a separate table column, you MUST INTELLIGENTLY ANALYZE and extract the color from Item Description or SKU! e.g., in "Nudie Rudie Bath Mat Mini - Dahlia", extract "Dahlia"; in "Tula Nudie Bath Mat - Cornflower", extract "Cornflower"; in "Hayfolk Bath Mat - Terra", extract "Terra")';
    }
    if (fieldName && fieldName.toLowerCase().includes('size')) {
      return 'string (CRITICAL: Extract size/dimensions. If Size is not in a separate table column, you MUST INTELLIGENTLY ANALYZE and extract the size from Item Description or SKU! e.g., in "Nudie Rudie Bath Mat Mini - Dahlia", extract "Mini"; in "Bath Mat - XL", extract "XL")';
    }
    return 'string (Extract clean text. If embedded in combined fields, parse and extract accurately without altering original source text)';
  };

  let mainSchema = {};
  config.mainFields.filter(f => f.visible).forEach(f => {
    mainSchema[f.name] = getTypeInstruction(f.type, f.name);
  });
  
  let itemSchema = {};
  config.itemFields.filter(f => f.visible).forEach(f => {
    itemSchema[f.name] = getTypeInstruction(f.type, f.name);
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
8. INTELLIGENT ATTRIBUTE EXTRACTION & DECOMPOSITION (Color, Size, Material, etc.):
   - Vendors often combine attributes into a single "Product Description", "Item Description", or "SKU" column (e.g., "Nudie Rudie Bath Mat Mini - Dahlia", "Tula Nudie Bath Mat - Cornflower", "Tula Nudie Bath Mat - Terra").
   - CRITICAL PRESERVATION RULE: For the main "Description" / "Product Description" field, you MUST extract and return the EXACT, FULL, ORIGINAL string as it appears on the vendor PO (e.g., keep "Nudie Rudie Bath Mat Mini - Dahlia" or "Tula Nudie Bath Mat - Cornflower" completely untouched as-it-is without removing any words from it!).
   - CRITICAL ATTRIBUTE MINING RULE: At the same time, you MUST act as an intelligent analyst to populate separate attribute fields (like Color, Size, Material, or any newly configured fields). Deeply examine the full description and SKU text. If a color (e.g., Dahlia, Cornflower, Terra, Blue, Red, Natural, etc.) or size (e.g., Mini, Small, XL, 50x70cm, etc.) is mentioned—whether separated by hyphens (-), slashes (/), commas, parentheses, or embedded directly—you MUST automatically extract and copy those specific values into their respective Color/Size/attribute fields without altering or deleting anything from the main description!

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
    parsedData = enrichExtractedItems(parsedData);
    return formatExtractedObjectDates(parsedData, config);
  } catch (e) {
    throw new Error("Failed to parse Gemini response: " + e.message);
  }
}

function enrichExtractedItems(data) {
  if (!data || !Array.isArray(data.items)) return data;
  data.items.forEach(item => {
    if (!item || typeof item !== 'object') return;
    const descKey = Object.keys(item).find(k => k.toLowerCase().includes('description') || k.toLowerCase().includes('item name') || k.toLowerCase().includes('product'));
    const colorKey = Object.keys(item).find(k => k.toLowerCase().includes('color') || k.toLowerCase().includes('colour'));
    const sizeKey = Object.keys(item).find(k => k.toLowerCase().includes('size') || k.toLowerCase().includes('dimen'));

    if (descKey && item[descKey]) {
      const descVal = String(item[descKey]).trim();
      if (colorKey && (!item[colorKey] || String(item[colorKey]).trim() === '')) {
        const parts = descVal.split(/\s*[-/]\s*/);
        if (parts.length > 1) {
          const possibleColor = parts[parts.length - 1].trim();
          if (possibleColor && possibleColor.length <= 30 && !/^\d+$/.test(possibleColor)) {
            item[colorKey] = possibleColor;
          }
        }
      }
      if (sizeKey && (!item[sizeKey] || String(item[sizeKey]).trim() === '')) {
        const sizeMatch = descVal.match(/\b(Mini|Small|Medium|Large|XL|XXL|XS|S|M|L|\d+\s*x\s*\d+\s*(?:cm|in|mm)?)\b/i);
        if (sizeMatch && sizeMatch[1]) {
          item[sizeKey] = sizeMatch[1];
        }
      }
    }
  });
  return data;
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
    SpreadsheetApp.flush();

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
//  WHATSAPP PI NOTIFICATION (to WhatsappUser sheet)
//  Always Indian (+91) numbers
// ─────────────────────────────────────────────
function sendPIWhatsAppNotification(piNumber, piDate, piFileUrl, poData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('WhatsappUser');
    if (!sheet) throw new Error('WhatsappUser tab not found in the sheet.');

    const data = sheet.getDataRange().getValues();
    const recipients = [];

    // Skip header row (row 1 = Name, row 2+ = data)
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][0] || '').trim();
      const phone = String(data[i][1] || '').trim();
      if (name && phone) {
        recipients.push({ name: name, phone: phone });
      }
    }

    if (recipients.length === 0) throw new Error('No recipients found in WhatsappUser tab.');

    // Determine greeting based on IST time (UTC+5:30)
    const now = new Date();
    const istHour = (now.getUTCHours() + 5 + Math.floor((now.getUTCMinutes() + 30) / 60)) % 24;
    let greeting = 'Good Day';
    if (istHour >= 5 && istHour < 12) {
      greeting = 'Good Morning';
    } else if (istHour >= 12 && istHour < 17) {
      greeting = 'Good Afternoon';
    } else if (istHour >= 17 && istHour < 22) {
      greeting = 'Good Evening';
    }

    // Build poData summary (buyer + PO number)
    const buyerName = (poData && poData['Buyer / Company Name']) ? poData['Buyer / Company Name'] : '';
    const poNumber = (poData && poData['PO Number']) ? poData['PO Number'] : '';

    const apiUrl = `https://api.maytapi.com/api/${CONFIG.MAYTAPI_PRODUCT_ID}/${CONFIG.MAYTAPI_PHONE_ID}/sendMessage`;
    const headers = {
      'Content-Type': 'application/json',
      'x-maytapi-key': CONFIG.MAYTAPI_TOKEN
    };

    // Convert Drive view link to Base64 (so Maytapi sends it as a proper PDF)
    const fileIdMatch = piFileUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    let mediaPayloadString = piFileUrl;
    
    if (fileIdMatch && fileIdMatch[1]) {
      try {
        const file = DriveApp.getFileById(fileIdMatch[1]);
        const b64 = Utilities.base64Encode(file.getBlob().getBytes());
        // Use strict base64 standard. The MIME type application/pdf tells Maytapi it's a PDF.
        mediaPayloadString = `data:application/pdf;base64,${b64}`;
      } catch (err) {
        Logger.log('WA Base64 fallback error: ' + err);
        mediaPayloadString = `https://drive.google.com/uc?export=download&id=${fileIdMatch[1]}`;
      }
    }

    let successCount = 0;

    for (let i = 0; i < recipients.length; i++) {
      const r = recipients[i];

      // Always Indian number — clean digits and prefix 91
      let rawPhone = String(r.phone).replace(/\D/g, '');
      if (rawPhone.length === 10) {
        rawPhone = '91' + rawPhone;
      } else if (rawPhone.startsWith('0')) {
        rawPhone = '91' + rawPhone.substring(1);
      } else if (!rawPhone.startsWith('91')) {
        rawPhone = '91' + rawPhone;
      }

      if (rawPhone.length < 12) continue; // Skip invalid

      const caption = `${greeting} *${r.name} Ji* 🙏\n\n` +
        `Your *Proforma Invoice* has been generated successfully!\n\n` +
        `📋 *PI Details:*\n` +
        `• PI Number: *${piNumber}*\n` +
        `• PI Date: *${piDate}*\n` +
        (poNumber ? `• PO Number: *${poNumber}*\n` : '') +
        (buyerName ? `• Buyer: *${buyerName}*` : '');

      // Send PDF as media message with caption
      const payload = {
        to_number: rawPhone,
        type: 'media',
        message: mediaPayloadString,
        text: caption
      };

      try {
        UrlFetchApp.fetch(apiUrl, {
          method: 'post',
          headers: headers,
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        successCount++;
      } catch (sendErr) {
        Logger.log('PI WA Send Error for ' + r.name + ': ' + sendErr);
      }
    }

    return { success: true, sentCount: successCount };
  } catch (e) {
    Logger.log('sendPIWhatsAppNotification error: ' + e);
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
    
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const formulas = dataRange.getFormulas();
    if (data.length <= 1) return []; // Only headers or empty
    
    const headers = data[0];
    const rows = [];
    const config = getAppConfig();
    
    for (let i = 1; i < data.length; i++) {
      let obj = { _id: i + 1 };
      for (let j = 0; j < headers.length; j++) {
        let val = data[i][j];
        const headerName = headers[j];
        const formula = formulas[i][j];

        if (formula && formula.toUpperCase().startsWith('=HYPERLINK(')) {
          const match = formula.match(/=HYPERLINK\(\s*"([^"]+)"\s*(?:,\s*"([^"]*)")?\)/i);
          if (match && match[1]) {
            const url = match[1];
            const text = match[2] || val || 'View Link';
            val = `<a href="${url}" target="_blank" style="color:#2563eb; font-weight:600; text-decoration:none;"><i class="fa-solid fa-link"></i> ${text}</a>`;
          }
        } else if (headerName === 'Timestamp') {
          if (val) val = formatLongDateTime(val);
        } else if (isDateKey(headerName, config) || (headerName && String(headerName).toLowerCase().includes('date'))) {
          if (val) val = formatShortDate(val);
        } else if (val instanceof Date) {
          if (val) val = formatShortDate(val);
        } else if (val && typeof val === 'string' && val.trim() !== '') {
          const parsed = parseAnyDate(val);
          if (parsed && (String(headerName).toLowerCase().includes('date') || String(headerName).toLowerCase().includes('time'))) {
            val = formatShortDate(val);
          }
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
// ─────────────────────────────────────────────
//  BULK UPDATE PO (Replaces old PO lines)
// ─────────────────────────────────────────────
function bulkUpdatePO(oldPoNumber, oldBuyerName, rowsData, providedPasscode) {
  try {
    // 1. Authenticate Passcode
    if (providedPasscode !== getPasscode()) {
      return { success: false, error: "Invalid Passcode!" };
    }

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return { success: false, error: "Sheet not found" };

    const existingData = sheet.getDataRange().getValues();
    const headers = existingData[0];
    const poIndex = headers.indexOf('PO Number');
    const buyerIndex = headers.indexOf('Buyer / Company Name');
    const config = getAppConfig();

    if (poIndex === -1 || buyerIndex === -1) {
      return { success: false, error: "Required columns missing in database" };
    }

    // 2. Find matching rows (1-indexed for Sheet APIs)
    let matchIndices = [];
    const targetPO = String(oldPoNumber).trim().toLowerCase();
    const targetBuyer = String(oldBuyerName).trim().toLowerCase();

    for (let i = 1; i < existingData.length; i++) {
      const rowPO = String(existingData[i][poIndex] || '').trim().toLowerCase();
      const rowBuyer = String(existingData[i][buyerIndex] || '').trim().toLowerCase();
      
      if (rowPO === targetPO && rowBuyer === targetBuyer) {
        matchIndices.push(i + 1); // +1 because sheet rows are 1-indexed
      }
    }

    // 3. Prepare new row data array matching the headers
    const timestampFormatted = formatLongDateTime(new Date());
    const preparedRows = rowsData.map(data => {
      const row = [];
      headers.forEach(header => {
        if (header === 'Timestamp') {
          row.push(timestampFormatted); // Update timestamp
        } else if (header === 'Source File') {
          row.push(data.fileUrl ? `=HYPERLINK("${data.fileUrl}", "${data.sourceFile || ''}")` : (data.sourceFile || ''));
        } else {
          let val = data[header] || '';
          if (val && isDateKey(header, config)) {
            val = formatShortDate(val);
          }
          row.push(val);
        }
      });
      return row;
    });

    // 4. Overwrite existing matching rows
    const overwriteCount = Math.min(matchIndices.length, preparedRows.length);
    for (let i = 0; i < overwriteCount; i++) {
      const rowIndex = matchIndices[i];
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([preparedRows[i]]);
    }

    // 5. Append surplus new rows if any
    for (let i = overwriteCount; i < preparedRows.length; i++) {
      sheet.appendRow(preparedRows[i]);
    }

    // 6. Delete surplus old rows if any (bottom-up to avoid shifting indices)
    if (preparedRows.length < matchIndices.length) {
      for (let i = matchIndices.length - 1; i >= preparedRows.length; i--) {
        sheet.deleteRow(matchIndices[i]);
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────────
//  GENERATE PROFORMA INVOICE
// ─────────────────────────────────────────────
function generatePI(payload) {
  try {
    let finalPiNum = payload.piNumber || '';
    if (!finalPiNum) {
      finalPiNum = getNextPINumber();
    }
  
    const rawHtml = UrlFetchApp.fetch('https://raw.githubusercontent.com/dme-wq/rkd-export-po-ocr/main/pi_template.html').getContentText();
    const template = HtmlService.createTemplate(rawHtml);
    template.poData = payload.poData || {};
    template.items = payload.selectedItems || [];
    template.selectedColumns = payload.selectedColumns || null;
    template.selectedHeaderFields = payload.selectedHeaderFields || null;
    template.piNumber = finalPiNum;
    template.piDate = payload.piDate || '';
    
    const htmlContent = template.evaluate().getContent();
    const blob = Utilities.newBlob(htmlContent, 'text/html', 'PI_' + finalPiNum + '.html').getAs('application/pdf');
    blob.setName('PI_' + finalPiNum + '.pdf');
    
    const folderId = '1M7x1SpPD94nxPiTk4PVeLbFL7M8lfZO9'; // Provided specific folder
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();
    
    // Update the Sheet
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const piNumIdx = headers.indexOf('PI Number');
    const piDateIdx = headers.indexOf('PI Date');
    const piLinkIdx = headers.indexOf('PI PDF Link');
    
    if (payload.selectedItems && payload.selectedItems.length > 0) {
      payload.selectedItems.forEach(item => {
        const r = item._id; // _id maps directly to row number in sheet
        if (r && r > 1) {
          if (piNumIdx > -1) sheet.getRange(r, piNumIdx + 1).setValue(finalPiNum);
          if (piDateIdx > -1) sheet.getRange(r, piDateIdx + 1).setValue(payload.piDate);
          if (piLinkIdx > -1) sheet.getRange(r, piLinkIdx + 1).setValue(`=HYPERLINK("${fileUrl}", "View PI")`);
        }
      });
    }
    
    return { success: true, url: fileUrl };
  } catch (err) {
    Logger.log('generatePI Error: ' + err);
    return { success: false, error: err.toString() };
  }
}

function getNextPINumber() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return `RKD-EX-PI-${new Date().getFullYear()}/1`;
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return `RKD-EX-PI-${new Date().getFullYear()}/1`;
  
  const headers = data[0];
  const piNumIdx = headers.indexOf('PI Number');
  if (piNumIdx === -1) return `RKD-EX-PI-${new Date().getFullYear()}/1`;
  
  const currentYear = new Date().getFullYear();
  let maxSeq = 0;
  
  for (let i = 1; i < data.length; i++) {
    const pi = data[i][piNumIdx];
    if (pi && String(pi).includes(`RKD-EX-PI-${currentYear}/`)) {
      const parts = String(pi).split('/');
      if (parts.length === 2) {
        const seq = parseInt(parts[1], 10);
        if (!isNaN(seq) && seq > maxSeq) {
          maxSeq = seq;
        }
      }
    }
  }
  return `RKD-EX-PI-${currentYear}/${maxSeq + 1}`;
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

  // 1. YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD with optional time
  let match = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})(?:[\sT](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    let year = parseInt(match[1], 10);
    let month = parseInt(match[2], 10) - 1;
    let day = parseInt(match[3], 10);
    let hh = match[4] ? parseInt(match[4], 10) : 0;
    let mm = match[5] ? parseInt(match[5], 10) : 0;
    let ss = match[6] ? parseInt(match[6], 10) : 0;
    if (month >= 0 && month < 12 && day >= 1 && day <= 31) {
      return new Date(year, month, day, hh, mm, ss);
    }
  }

  // 2. DD-MMM-YYYY or DD MMM YYYY or DD/MMM/YYYY with optional time
  match = str.match(/^(\d{1,2})[\s\/\-\.]?([a-zA-Z]{3,9})[\s\/\-\.]?(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    let day = parseInt(match[1], 10);
    let mStr = match[2].toLowerCase();
    let year = parseInt(match[3], 10);
    if (year < 100) year += 2000;
    let month = monthMap[mStr];
    if (month !== undefined && day >= 1 && day <= 31) {
      let hh = match[4] ? parseInt(match[4], 10) : 0;
      let mm = match[5] ? parseInt(match[5], 10) : 0;
      let ss = match[6] ? parseInt(match[6], 10) : 0;
      return new Date(year, month, day, hh, mm, ss);
    }
  }

  // 3. MMM DD, YYYY or MMM DD YYYY with optional time
  match = str.match(/^([a-zA-Z]{3,9})[\s\/\-\.]?(\d{1,2}),?[\s\/\-\.]?(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    let mStr = match[1].toLowerCase();
    let day = parseInt(match[2], 10);
    let year = parseInt(match[3], 10);
    if (year < 100) year += 2000;
    let month = monthMap[mStr];
    if (month !== undefined && day >= 1 && day <= 31) {
      let hh = match[4] ? parseInt(match[4], 10) : 0;
      let mm = match[5] ? parseInt(match[5], 10) : 0;
      let ss = match[6] ? parseInt(match[6], 10) : 0;
      return new Date(year, month, day, hh, mm, ss);
    }
  }

  // 4. DD-MM-YYYY or DD/MM/YYYY or MM/DD/YYYY or DD.MM.YYYY with 4-digit year at end
  match = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})(?:[\sT](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    let p1 = parseInt(match[1], 10);
    let p2 = parseInt(match[2], 10);
    let year = parseInt(match[3], 10);
    let hh = match[4] ? parseInt(match[4], 10) : 0;
    let mm = match[5] ? parseInt(match[5], 10) : 0;
    let ss = match[6] ? parseInt(match[6], 10) : 0;

    let day, month;
    if (p1 > 12 && p2 <= 12) {
      day = p1;
      month = p2 - 1;
    } else if (p2 > 12 && p1 <= 12) {
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

  // 5. DD-MM-YY or MM/DD/YY with 2-digit year at end
  match = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2})(?:[\sT](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    let p1 = parseInt(match[1], 10);
    let p2 = parseInt(match[2], 10);
    let year = parseInt(match[3], 10) + 2000;
    let hh = match[4] ? parseInt(match[4], 10) : 0;
    let mm = match[5] ? parseInt(match[5], 10) : 0;
    let ss = match[6] ? parseInt(match[6], 10) : 0;

    let day, month;
    if (p1 > 12 && p2 <= 12) {
      day = p1;
      month = p2 - 1;
    } else if (p2 > 12 && p1 <= 12) {
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
  const str = String(val).trim();
  if (/^\d{2}-[a-zA-Z]{3}-\d{4}$/.test(str)) return str;

  const d = parseAnyDate(val);
  if (!d) return str;
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  const month = months[d.getMonth()];
  const year = d.getFullYear();
  return `${day}-${month}-${year}`;
}

function formatLongDateTime(val) {
  if (val === null || val === undefined || val === '') return '';
  const str = String(val).trim();
  if (/^\d{2}-[a-zA-Z]{3}-\d{4} \d{2}:\d{2}:\d{2}$/.test(str)) return str;

  const d = parseAnyDate(val);
  if (!d) return str;
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
            if (isDateKey(itemKey, config)) {
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
    } else if (isDateKey(key, config)) {
      if (data[key]) {
        const parsed = formatShortDate(data[key]);
        if (parsed) data[key] = parsed;
      }
    }
  });

  return data;
}
