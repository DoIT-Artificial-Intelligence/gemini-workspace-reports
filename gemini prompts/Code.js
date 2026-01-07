// --- CONFIGURATION ---
const properties = PropertiesService.getScriptProperties();
const CONFIG = {
  const properties = PropertiesService.getScriptProperties();
  MATTER_ID: properties.getProperty('MATTER_ID'),
  TARGET_USER: properties.getProperty('TARGET_USER'),
  
  // Specific Subfolder IDs
  XML_FOLDER_ID: properties.getProperty('XML_FOLDER_ID'),
  SHEETS_FOLDER_ID: properties.getProperty('SHEETS_FOLDER_ID'),
  
  SHEET_CELL_LIMIT: 49000 // Truncate text strictly to this length
};

function exportGeminiAndSave() {
  const TOKEN = ScriptApp.getOAuthToken();

  try {
    Logger.log(`Starting GEMINI export for User: ${CONFIG.TARGET_USER} and Matter ID: ${CONFIG.MATTER_ID}`);

    // 1. Create the Export Request
    const exportUrl = `https://vault.googleapis.com/v1/matters/${CONFIG.MATTER_ID}/exports`;

    const payload = {
      name: "Gemini Export - " + new Date().toISOString(),
      query: {
        corpus: "GEMINI",
        dataScope: "ALL_DATA",
        searchMethod: "ACCOUNT",
        accountInfo: {
          emails: [CONFIG.TARGET_USER]
        }
      },
      exportOptions: {
        geminiOptions: {
          exportFormat: "XML"
        }
      }
    };

    const requestOptions = {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: `Bearer ${TOKEN}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const createResponse = UrlFetchApp.fetch(exportUrl, requestOptions);
    if (createResponse.getResponseCode() !== 200) {
      throw new Error(`Failed to create export. Code: ${createResponse.getResponseCode()} | Response: ${createResponse.getContentText()}`);
    }
    
    const exportData = JSON.parse(createResponse.getContentText());
    const exportId = exportData.id;
    Logger.log(`Export initiated successfully. Export ID: ${exportId}`);

    // 2. Poll for Completion (30 checks x 2 mins = ~60 mins max)
    let status = "IN_PROGRESS";
    let completedExportData;

    for (let i = 0; i < 30; i++) {
      Logger.log(`Waiting 2 minutes before check #${i + 1}...`);
      Utilities.sleep(120000); 
      
      const statusUrl = `https://vault.googleapis.com/v1/matters/${CONFIG.MATTER_ID}/exports/${exportId}`;
      const statusResponse = UrlFetchApp.fetch(statusUrl, {
        method: "get",
        headers: { Authorization: `Bearer ${TOKEN}` },
        muteHttpExceptions: true
      });
      
      completedExportData = JSON.parse(statusResponse.getContentText());
      status = completedExportData.status;
      Logger.log(`Current Status: ${status}`);
      
      if (status === "COMPLETED") break;
      if (status === "FAILED") throw new Error("Export failed on Google's side.");
    }
    
    if (status !== "COMPLETED") {
      Logger.log("Timed out waiting for export.");
      return;
    }

    // 3. Download and Extract Files
    if (completedExportData.cloudStorageSink && completedExportData.cloudStorageSink.files) {
      const files = completedExportData.cloudStorageSink.files;
      
      // Get the specific folders
      const xmlFolder = DriveApp.getFolderById(CONFIG.XML_FOLDER_ID);
      const sheetsFolder = DriveApp.getFolderById(CONFIG.SHEETS_FOLDER_ID);

      Logger.log(`Export complete. Found ${files.length} file(s). Looking for ZIP...`);

      files.forEach(file => {
        const bucket = file.bucketName;
        const objectName = file.objectName;
        const fileName = objectName.split('/').pop();

        // Only process ZIP files
        if (!fileName.toLowerCase().endsWith('.zip')) {
           Logger.log(`Skipping auxiliary file: ${fileName}`);
           return; 
        }

        Logger.log(`Processing ZIP file: ${fileName}`);
        const downloadUrl = `https://storage.googleapis.com/storage/v1/b/${bucket}/o/${encodeURIComponent(objectName)}?alt=media`;
        
        const downloadResp = UrlFetchApp.fetch(downloadUrl, {
          method: "get",
          headers: { Authorization: `Bearer ${TOKEN}` },
          muteHttpExceptions: true
        });
        
        if (downloadResp.getResponseCode() === 200) {
          const zipBlob = downloadResp.getBlob();
          
          try {
            const unzippedBlobs = Utilities.unzip(zipBlob);
            let xmlSaved = false;

            unzippedBlobs.forEach(innerBlob => {
                const innerName = innerBlob.getName();
                
                if (innerName.toLowerCase().endsWith(".xml")) {
                    // --- SAVE XML LOGIC (OVERWRITE) ---
                    const targetXmlName = `${CONFIG.TARGET_USER}.xml`;
                    
                    // Check for existing file and delete it
                    const existingFiles = xmlFolder.getFilesByName(targetXmlName);
                    while (existingFiles.hasNext()) {
                      existingFiles.next().setTrashed(true);
                      Logger.log(`Deleted existing old file: ${targetXmlName}`);
                    }

                    // Save new file and rename it explicitly
                    const savedFile = xmlFolder.createFile(innerBlob);
                    savedFile.setName(targetXmlName);
                    
                    Logger.log(`SUCCESS: Saved "${targetXmlName}" to XML Folder.`);
                    
                    // --- CONVERT TO SHEET LOGIC ---
                    Logger.log(`Starting conversion to Google Sheets...`);
                    // We pass the blob and the specific sheets folder
                    convertXmlBlobToSheet(innerBlob, sheetsFolder);
                    xmlSaved = true;
                }
            });

            if (!xmlSaved) Logger.log("WARNING: Zip downloaded, but no XML file found inside.");

          } catch (zipErr) {
             Logger.log(`ERROR: Could not unzip ${fileName}. Details: ${zipErr.message}`);
          }

        } else {
          Logger.log(`ERROR downloading ${fileName}: ${downloadResp.getContentText()}`);
        }
      });
    } else {
      Logger.log("Export completed but contained no files.");
    }

  } catch (e) {
    Logger.log("CRITICAL ERROR: " + e.message);
  }
}

/**
 * Parses XML Blob and writes to a new Google Sheet.
 * Applies strict truncation to ALL cells to prevent API errors.
 * Overwrites existing sheet if present.
 */
function convertXmlBlobToSheet(xmlBlob, targetFolder) {
  try {
    const xmlContent = xmlBlob.getDataAsString();
    
    // Helper regex to extract content between tags
    const extract = (text, tag) => {
      const regex = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\/${tag}>`, 'i');
      const match = text.match(regex);
      return match ? decodeXmlEntities(match[1]) : "";
    };

    // Helper to safely truncate any value (string or otherwise)
    const safeTruncate = (val) => {
      if (typeof val !== 'string') return val;
      if (val.length <= CONFIG.SHEET_CELL_LIMIT) return val;
      return val.substring(0, CONFIG.SHEET_CELL_LIMIT) + "\n...[TRUNCATED]";
    };

    // 1. Parse User Email
    const userMatch = xmlContent.match(/<User>\s*<Email>(.*?)<\/Email>\s*<\/User>/i);
    const userEmail = userMatch ? userMatch[1] : "Unknown";

    // 2. Prepare Data Structure
    const outputData = [];
    const headers = [
      "User", "Conversation ID", "Conversation Topic", "Turn No.", 
      "Request ID", "Model Version", "Timestamp", "Prompt", "Response ID", "Response"
    ];
    outputData.push(headers);

    // 3. Process Conversations
    const conversationBlocks = xmlContent.split('<Conversation>');
    
    for (let i = 1; i < conversationBlocks.length; i++) {
      const convBlock = conversationBlocks[i];
      if (!convBlock.includes('</Conversation>')) continue;

      const convId = extract(convBlock, "ConversationId");
      const convTopic = extract(convBlock, "ConversationTopic").trim();

      const turnBlocks = convBlock.split('<ConversationTurn>');
      
      for (let j = 1; j < turnBlocks.length; j++) {
        const turnBlock = turnBlocks[j];
        if (!turnBlock.includes('</ConversationTurn>')) continue;

        // Extract metadata
        const turnNumber = j;
        const requestId = extract(turnBlock, "RequestId");
        const modelVersion = extract(turnBlock, "ModelVersion");
        const timestamp = extract(turnBlock, "Timestamp");

        // Extract Prompt
        let promptText = "";
        const promptBlockMatch = turnBlock.match(/<Prompt>([\s\S]*?)<\/Prompt>/i);
        if (promptBlockMatch) {
            promptText = extract(promptBlockMatch[0], "Text");
        }

        // Extract Response
        let respId = "";
        let respText = "";
        const respBlockMatch = turnBlock.match(/<PrimaryResponse>([\s\S]*?)<\/PrimaryResponse>/i);
        if (respBlockMatch) {
            const respBlock = respBlockMatch[0];
            respId = extract(respBlock, "ResponseId");
            respText = extract(respBlock, "Text");
        }

        // 4. Construct Row and Truncate EVERYTHING
        const row = [
          userEmail, convId, convTopic, turnNumber, 
          requestId, modelVersion, timestamp, promptText, respId, respText
        ].map(safeTruncate);

        outputData.push(row);
      }
    }

    // 5. Write to Sheet
    if (outputData.length > 1) { 
      // Force the sheet name to be the User Email (as requested)
      const sheetName = CONFIG.TARGET_USER;
      
      // --- OVERWRITE LOGIC ---
      // Check for duplicate in the specific folder and delete it
      const existingSheets = targetFolder.getFilesByName(sheetName);
      while (existingSheets.hasNext()) {
        existingSheets.next().setTrashed(true);
        Logger.log(`Deleted existing old sheet: ${sheetName}`);
      }

      // Create new Sheet (defaults to root folder)
      const ss = SpreadsheetApp.create(sheetName);
      const sheet = ss.getActiveSheet();
      
      // Batch write data
      sheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
      
      // Formatting
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, outputData[0].length).setFontWeight("bold");
      
      // Move to target sheets folder
      DriveApp.getFileById(ss.getId()).moveTo(targetFolder);
      
      Logger.log(`SUCCESS: Created Google Sheet "${sheetName}" with ${outputData.length - 1} rows in Sheets folder.`);
    } else {
      Logger.log("Parsed XML but found no conversation data rows.");
    }

  } catch (e) {
    Logger.log("ERROR converting XML to Sheet: " + e.message + " | Stack: " + e.stack);
  }
}

/**
 * Decodes XML entities to plain text
 */
function decodeXmlEntities(str) {
  if (!str) return "";
  return str.replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'") 
            .replace(/&#39;/g, "'");
}
