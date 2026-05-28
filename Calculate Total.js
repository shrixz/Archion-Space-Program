function calculateTotalGSM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let totalGSM = 0;
  const searchText = "Total Departmental GSM";
  
  // Array to hold the breakdown data
  let breakdown = []; 

  // Add any non-department sheets here that should be ignored
  const sheetsToSkip = ["Settings", "Summary", "Summary Template", "Cover Template"];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    
    // SKIP empty sheets AND any sheet listed in the sheetsToSkip array
    if (lastRow === 0 || sheetsToSkip.includes(sheetName)) return; 

    try {
      const data = sheet.getRange(1, 1, lastRow, 7).getValues(); 
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === searchText) {
          const val = data[i][6]; // Column G
          if (val !== "" && !isNaN(parseFloat(val))) {
            const numVal = parseFloat(val);
            totalGSM += numVal;
            
            // Push the sheet name and its specific GSM into our array
            breakdown.push({
              name: sheetName,
              value: numVal.toFixed(2)
            });
          }
          break; 
        }
      }
    } catch (e) {
      console.log("Error reading sheet: " + sheetName);
    }
  });

  // Return an object containing both the total and the breakdown list
  return {
    total: totalGSM.toFixed(2),
    breakdown: breakdown
  };
}