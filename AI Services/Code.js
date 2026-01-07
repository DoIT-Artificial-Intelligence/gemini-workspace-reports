function exportAiSettings() {
  // Configuration
  const properties = PropertiesService.getScriptProperties();
  const SPREADSHEET_ID = properties.getProperty('SPREADSHEET_ID');
  const FILTER = "setting.type.matches('gemini_app|notebooklm|ai_studio')";
  
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
    sheet.clear();
    
    // 1. Get Customer ID (Required for API calls)
    const customerId = AdminDirectory.Customers.get("my_customer").id;
    
    // 2. Fetch Policies
    const baseUrl = "https://cloudidentity.googleapis.com/v1/policies";
    let policies = [];
    let pageToken = null;
    
    do {
      const queryParams = [
        `filter=${encodeURIComponent(FILTER)}`,
        `pageSize=100`
      ];
      if (pageToken) queryParams.push(`pageToken=${pageToken}`);
      
      const url = `${baseUrl}?${queryParams.join('&')}`;
      const params = {
        method: "get",
        headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, params);
      
      if (response.getResponseCode() !== 200) {
        throw new Error(`API Error (${response.getResponseCode()}): ${response.getContentText()}`);
      }
      
      const result = JSON.parse(response.getContentText());
      if (result.policies) {
        policies = policies.concat(result.policies);
      }
      pageToken = result.nextPageToken;
      
    } while (pageToken);

    if (policies.length === 0) {
      console.log("No matching policies found.");
      sheet.appendRow(["No matching policies found"]);
      return;
    }

    // 3. Helper Functions for Resolving Names
    const ouCache = {};
    const groupCache = {};
    
    const resolveOrgUnit = (ouResource) => {
      if (!ouResource) return "/"; 
      if (ouCache[ouResource]) return ouCache[ouResource];
      
      const ouId = ouResource.split('/')[1];
      if (!ouId) return ouResource;

      try {
        const ou = AdminDirectory.Orgunits.get(customerId, `id:${ouId}`);
        const path = ou.orgUnitPath || "/";
        ouCache[ouResource] = path;
        return path;
      } catch (e) {
        console.warn(`Could not resolve OU ${ouId}: ${e.message}`);
        ouCache[ouResource] = ouResource;
        return ouResource;
      }
    };

    const resolveGroup = (groupResource) => {
      if (!groupResource) return "";
      if (groupCache[groupResource]) return groupCache[groupResource];

      const groupId = groupResource.split('/')[1];
      if (!groupId) return groupResource;

      try {
        const group = AdminDirectory.Groups.get(groupId);
        groupCache[groupResource] = group.email;
        return group.email;
      } catch (e) {
        console.warn(`Could not resolve Group ${groupId}: ${e.message}`);
        groupCache[groupResource] = groupResource;
        return groupResource;
      }
    };

    // 4. Transform Data
    const headers = [
      "name", 
      "policyQuery.orgUnit", 
      "policyQuery.orgUnitPath", // Index 2
      "policyQuery.sortOrder", 
      "setting.type",            // Index 4
      "setting.value.serviceState", 
      "type", 
      "policyQuery.group", 
      "policyQuery.groupEmail"   // Index 8
    ];
    
    const rows = policies.map(p => {
      const pq = p.policyQuery || {};
      const ouId = pq.orgUnit || "";
      const groupId = pq.group || "";
      
      // Resolve paths
      const ouPath = ouId ? resolveOrgUnit(ouId) : "";
      const groupEmail = groupId ? resolveGroup(groupId) : "";
      
      const serviceState = p.setting?.value?.serviceState || "";
      
      return [
        p.name,                      
        ouId,                        
        ouPath,                      
        pq.sortOrder || "",          
        p.setting?.type || "",       
        serviceState,                
        p.type || "ADMIN",           
        groupId,                     
        groupEmail                   
      ];
    });
    
    rows.sort((a, b) => {
      // 1. Sort by policyQuery.orgUnitPath (Index 2)
      let comparison = a[2].localeCompare(b[2]);
      if (comparison !== 0) return comparison;

      // 2. Sort by setting.type (Index 4)
      comparison = a[4].localeCompare(b[4]);
      if (comparison !== 0) return comparison;

      // 3. Sort by policyQuery.groupEmail (Index 8)
      return a[8].localeCompare(b[8]);
    });

    // 5. Write to Sheet
    sheet.appendRow(headers);
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    console.log(`Successfully wrote ${rows.length} rows.`);
    
  } catch (e) {
    console.error("Script failed: " + e.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: " + e.message);
  }
}
