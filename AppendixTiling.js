/**
 * Consolidates sheets listed in Settings Col H into an Appendix sheet.
 */
function buildAppendixConsolidation() {
  const ss = SpreadsheetApp.getActive();
  const settings = ss.getSheetByName("Settings");
  const template = ss.getSheetByName("Header_Footer");
  
  if (!settings || !template) {
    SpreadsheetApp.getUi().alert("Settings or Header_Footer sheet missing!");
    return;
  }

  // 1. Setup Appendix Sheet
  let appendix = ss.getSheetByName("Appendix");
  if (appendix) ss.deleteSheet(appendix);
  appendix = ss.insertSheet("Appendix");

  // --- SETTINGS ---
  const PAGE_HEIGHT = 38;          
  const DATA_START_REL = 2;        
  const DATA_END_REL = 35;         
  const ROWS_PER_PAGE = (DATA_END_REL - DATA_START_REL) + 1;

  // Sync Column Widths and ensure enough columns exist
  const maxCols = template.getLastColumn();
  if (appendix.getMaxColumns() < maxCols) {
    appendix.insertColumnsAfter(appendix.getMaxColumns(), maxCols - appendix.getMaxColumns());
  }

  for (let col = 1; col <= maxCols; col++) {
    appendix.setColumnWidth(col, template.getColumnWidth(col));
  }

  const templateRowHeights = [];
  for (let i = 1; i <= PAGE_HEIGHT; i++) {
    templateRowHeights.push(template.getRowHeight(i));
  }

  const settingsRowStart = 3;
  const sheetList = settings.getRange(settingsRowStart, 8, Math.max(settings.getLastRow() - (settingsRowStart - 1), 1), 1)
    .getValues()
    .map(r => r[0])
    .filter(name => name && ss.getSheetByName(name));

  let currentPage = 0;
  let currentRowInPage = 0;

  function checkAndTile() {
    const startOfBlock = (currentPage * PAGE_HEIGHT) + 1;
    
    const neededRows = startOfBlock + PAGE_HEIGHT - 1;
    if (appendix.getMaxRows() < neededRows) {
      appendix.insertRowsAfter(appendix.getMaxRows(), PAGE_HEIGHT);
    }
    
    const sourceRange = template.getRange(1, 1, PAGE_HEIGHT, maxCols);
    const targetRange = appendix.getRange(startOfBlock, 1, PAGE_HEIGHT, maxCols);
    
    sourceRange.copyTo(targetRange); 
    
    for (let h = 0; h < templateRowHeights.length; h++) {
      appendix.setRowHeight(startOfBlock + h, templateRowHeights[h]);
    }
    
    // CRITICAL: Unmerge only the data area to prevent the "Unexpected Error" 
    // when copying source rows into the tile.
    const dataArea = appendix.getRange(startOfBlock + DATA_START_REL - 1, 1, ROWS_PER_PAGE, maxCols);
    dataArea.breakApart(); 
    dataArea.clearContent();
  }

  // Process Sheets
  checkAndTile();

  sheetList.forEach(sheetName => {
    const sourceSheet = ss.getSheetByName(sheetName);
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 1) return;

    const sourceCols = Math.min(sourceSheet.getLastColumn(), maxCols);

    for (let r = 1; r <= lastRow; r++) {
      if (currentRowInPage >= ROWS_PER_PAGE) {
        currentPage++;
        currentRowInPage = 0;
        checkAndTile();
      }

      const targetRow = (currentPage * PAGE_HEIGHT) + DATA_START_REL + currentRowInPage;
      sourceSheet.getRange(r, 1, 1, sourceCols).copyTo(appendix.getRange(targetRow, 1));
      currentRowInPage++;
    }
    
    if (currentRowInPage < ROWS_PER_PAGE) {
      currentRowInPage++;
    }
  });

  // Hiding gridlines for a clean report look
  appendix.setHiddenGridlines(true); 
  appendix.activate();
  
  SpreadsheetApp.getUi().alert("Appendix Generated Successfully!");
}