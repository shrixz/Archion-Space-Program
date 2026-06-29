/**
 * Appendix Builder - 25pt Margin Max Size
 * Logic: 25pt margins all around. 290pt width for overflow. 540pt max height.
 */
function buildAppendixFromTemplate() {
  const ss = SpreadsheetApp.getActive();
  const settings = ss.getSheetByName("Settings");
  
  const lastRowSettings = settings.getLastRow();
  if (lastRowSettings < 5) return;
  const targetSheetNames = settings.getRange(5, 8, lastRowSettings - 4, 1).getValues().flat().filter(String);
  
  const uniqueRoomNames = new Set();
  targetSheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName.trim());
    if (sheet) {
      const lastRowSheet = sheet.getLastRow();
      if (lastRowSheet >= 1) { 
        const rooms = sheet.getRange(1, 1, lastRowSheet, 1).getValues().flat();
        rooms.forEach(room => {
          if (room) uniqueRoomNames.add(room.toString().trim());
        });
      }
    }
  });

  const IMAGE_FOLDER_ID = '1igI5KnZdENWddJtKwnkFvnDFFH1w8NRI';
  const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
  const driveFilesMap = {};
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const rawName = file.getName().replace(/\.[^/.]+$/, "");
    const cleanKey = rawName.toLowerCase().replace(/\s+/g, '');
    driveFilesMap[cleanKey] = file.getBlob();
  }

  const validImageBlobs = [];
  uniqueRoomNames.forEach(roomName => {
    const cleanRoomKey = roomName.toLowerCase().replace(/\s+/g, '');
    if (driveFilesMap[cleanRoomKey]) {
      validImageBlobs.push({ name: roomName, blob: driveFilesMap[cleanRoomKey] });
    }
  });

  if (validImageBlobs.length === 0) {
    SpreadsheetApp.getUi().alert("No matching images found in Drive.");
    return;
  }

  const TEMPLATE_DOC_ID = '1mtmddvKADZIVnZl5EQ-okGQDqyk-5wnDHjWviGJ7vsY';
  const newDocFile = DriveApp.getFileById(TEMPLATE_DOC_ID).makeCopy("Generated Appendix - " + new Date().toLocaleDateString());
  const doc = DocumentApp.openById(newDocFile.getId());
  const body = doc.getBody();

  // --- 1. NEW 25PT MARGINS ---
  body.setMarginLeft(25).setMarginRight(25).setMarginTop(25).setMarginBottom(25);

  // --- 2. MAXIMIZED DIMENSIONS ---
  const OVERFLOW_WIDTH = 290;   
  const MAX_HEIGHT = 540;       
  const COL_TECH_WIDTH = 247; // 742 / 3

  for (let i = 0; i < validImageBlobs.length; i += 3) {
    const chunk = validImageBlobs.slice(i, i + 3);
    let table = body.appendTable();
    table.setBorderWidth(0);
    table.setAttributes({
      [DocumentApp.Attribute.BORDER_WIDTH]: 0,
      [DocumentApp.Attribute.MARGIN_TOP]: 0,
      [DocumentApp.Attribute.MARGIN_BOTTOM]: 0
    });
    
    let currentRow = table.appendTableRow();

    chunk.forEach((imgData) => {
      const cell = currentRow.appendTableCell();
      cell.setWidth(COL_TECH_WIDTH); 
      cell.setPaddingLeft(0).setPaddingRight(0).setPaddingTop(0).setPaddingBottom(0);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

      const image = cell.appendImage(imgData.blob);
      const para = image.getParent().asParagraph();
      para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      para.setSpacingAfter(0).setSpacingBefore(0);
      para.setLineSpacing(0.06); 

      // --- PROPORTIONAL SCALING ---
      let w = image.getWidth();
      let h = image.getHeight();
      
      let ratio = OVERFLOW_WIDTH / w;
      let targetWidth = OVERFLOW_WIDTH;
      let targetHeight = h * ratio;

      if (targetHeight > MAX_HEIGHT) {
        let heightRatio = MAX_HEIGHT / targetHeight;
        targetWidth = targetWidth * heightRatio;
        targetHeight = MAX_HEIGHT;
      }
      
      image.setWidth(targetWidth).setHeight(targetHeight);
    });

    while (currentRow.getNumCells() < 3) {
      currentRow.appendTableCell().setWidth(COL_TECH_WIDTH).setPaddingLeft(0).setPaddingRight(0);
    }

    if (i + 3 < validImageBlobs.length) {
      body.appendPageBreak();
    }
  }

  doc.saveAndClose();
  
  const url = doc.getUrl();
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif; text-align:center; padding: 20px;">' +
    '<p style="font-size: 16px; font-weight: bold;">Appendix Generated Successfully.</p>' +
    '<p style="font-size: 14px; color: #333;">Total Unique Images: <b>' + validImageBlobs.length + '</b></p>' +
    '<br><a href="' + url + '" target="_blank" style="display:inline-block; padding:10px 20px; background:#0F9D58; color:white; text-decoration:none; border-radius:4px; font-weight:bold;">OPEN DOCUMENT</a>' +
    '</div>'
  ).setWidth(350).setHeight(170);
  
  SpreadsheetApp.getUi().showModalDialog(html, "Process Complete");
}