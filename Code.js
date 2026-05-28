/**
 * SHIM: Prevents "setTimeout is not defined" error from pdf-lib library
 */
var setTimeout = function(cb, ms) { Utilities.sleep(ms); cb(); };
var clearTimeout = function(id) {};

/**
 * REFRESH SHEET NAMES
 */
function refreshSheetNames() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Settings") || ss.insertSheet("Settings");
  const names = ss.getSheets().map(s => [s.getName()]);
  sheet.getRange(1, 1, 1000, 1).clearContent();
  sheet.getRange(5, 1, names.length, 1).setValues(names);
}

/**
 * MENU SETUP
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Report Automation')
    .addItem('Create Summary', 'buildSummaryFromSettings')
    .addSeparator()
    .addItem('Build Appendix', 'buildAppendixFromTemplate')
    .addItem('Refresh Sheet Names', 'refreshSheetNames')
    .addItem('Show Total GSM', 'showSidebar')
    .addSeparator()
    .addItem('Setup Auto-Refresh', 'setupOnChangeTrigger')
    .addItem('Grant Permissions', 'requestPermissions')
    .addToUi();
}

function requestPermissions() {
  const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  const status = authInfo.getAuthorizationStatus();

  if (status === ScriptApp.AuthorizationStatus.REQUIRED) {
    const authUrl = authInfo.getAuthorizationUrl();
    const html = HtmlService.createHtmlOutput(
      '<p style="font-family:Arial,sans-serif;font-size:13px;padding:10px">' +
      'New permissions are required.<br><br>' +
      '<a href="' + authUrl + '" target="_blank" ' +
      'style="color:#1a73e8;font-weight:bold;font-size:14px">Click here to grant permissions</a><br><br>' +
      'After approving, close this dialog and click <b>Setup Auto-Refresh</b> again.</p>'
    ).setWidth(400).setHeight(130);
    SpreadsheetApp.getUi().showModalDialog(html, 'Permission Required');
  } else {
    SpreadsheetApp.getUi().alert('All permissions are already granted! Now click Setup Auto-Refresh.');
  }
}

/**
 * RATE-LIMITED FETCH (Prevents 429 Errors)
 */
function fetchWithRetry(url, token) {
  var params = {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  };
  
  var maxRetries = 5;
  var waitTime = 2000; // Start with 2 seconds
  
  for (var i = 0; i < maxRetries; i++) {
    var response = UrlFetchApp.fetch(url, params);
    var code = response.getResponseCode();
    
    if (code === 200) {
      return response.getBlob();
    } else if (code === 429) {
      Utilities.sleep(waitTime);
      waitTime *= 2; // Double the wait time for next try
    } else {
      throw new Error("Export failed with code: " + code);
    }
  }
  throw new Error("Max retries reached. Google is still rate-limiting the PDF export.");
}

/**
 * MAIN FUNCTION: BUILD SUMMARY & MERGE PDF
 */
async function buildSummaryFromSettings() {
  const ss = SpreadsheetApp.getActive();
  const template = ss.getSheetByName("Summary Template");
  const coverTemplate = ss.getSheetByName("Cover Template");
  
  if (!template || !coverTemplate) throw new Error("Templates not found.");
  
  const versionText = coverTemplate.getRange(10, 9).getDisplayValue(); 
  const dateText = coverTemplate.getRange(15, 9).getDisplayValue();
  const fullHeaderValue = versionText + "\n" + dateText;
  const hospitalName = template.getRange("F30").getDisplayValue();
  
  const existingSummary = ss.getSheetByName("Summary");
  if (existingSummary) ss.deleteSheet(existingSummary);
  
  const target = coverTemplate.copyTo(ss).setName("Summary");
  
  const PAGE_HEIGHT = 31;          
  const DATA_START_REL = 11;    
  const DATA_END_REL = 27;      
  const ROWS_PER_PAGE = (DATA_END_REL - DATA_START_REL) + 1; 
  
  const templateRowHeights = [];
  for (let i = 1; i <= PAGE_HEIGHT; i++) {
    templateRowHeights.push(template.getRowHeight(i));
  }

  let currentPage = 1; 
  let currentRowInPage = 0; 

  function checkAndTile() {
    if (currentRowInPage >= ROWS_PER_PAGE || (currentPage === 1 && currentRowInPage === 0)) {
      if (currentRowInPage >= ROWS_PER_PAGE) {
        currentPage++;
        currentRowInPage = 0;
      }
      const start = (currentPage * PAGE_HEIGHT) + 1;
      target.insertRowsAfter(target.getMaxRows(), PAGE_HEIGHT);
      template.getRange(1, 1, PAGE_HEIGHT, 9).copyTo(target.getRange(start, 1));
      for (let h = 0; h < templateRowHeights.length; h++) {
        target.setRowHeight(start + h, templateRowHeights[h]);
      }
      const headerCell = target.getRange(start, 6);
      headerCell.setValue(fullHeaderValue);
      headerCell.setWrap(true).setHorizontalAlignment("right").setVerticalAlignment("top");
      target.getRange(start + 29, 6).setValue(hospitalName);
      target.getRange(start + DATA_START_REL - 1, 1, ROWS_PER_PAGE, 5).clearContent();
    }
  }

  checkAndTile();
  
  const settings = ss.getSheetByName("Settings");

  // --- AUTO-CLEAN SETTINGS: Remove rows in Col H whose sheet no longer exists ---
  const _preCleanRow = settings.getLastRow();
  if (_preCleanRow >= 5) {
    const _colH = settings.getRange(5, 8, _preCleanRow - 4, 1).getValues();
    for (let i = _colH.length - 1; i >= 0; i--) {
      const _ref = String(_colH[i][0]).trim();
      if (_ref !== "" && !ss.getSheetByName(_ref)) {
        settings.deleteRow(5 + i);
      }
    }
  }

  const lastRow = settings.getLastRow();

  // --- NEW: Fetch and Format Header Data for Appendix ---
  const rawDate = settings.getRange("C4").getValue();
  let headerDate = "";
  if (rawDate) {
    try {
      headerDate = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "MM/dd/yyyy");
    } catch(e) {
      headerDate = String(rawDate); // Fallback if cell doesn't contain a valid date
    }
  }
  const headerC3 = settings.getRange("C3").getDisplayValue();
  
  // --- UPDATED: Fetch Gross Multiplier from D2 ---
  const grossMultiplier = Number(settings.getRange("D2").getValue()) || 1.25; 
  // --------------------------------------------------------

  // Update: Read 8 columns starting from B (Col 2) up to I (Col 9). 
  // index 0 = B (Category), index 1 = C (Checkbox), index 5 = G (Section Full), index 6 = H (Sheet), index 7 = I (Display Name)
  let settingsData = (lastRow < 5) ? [] : settings.getRange(5, 2, lastRow - 4, 8).getValues().filter(row => row[6]); 
  const appendixList = (lastRow < 5) ? [] : settings.getRange(5, 8, lastRow - 4, 1).getValues().flat().filter(name => name !== "" && ss.getSheetByName(name));

  // --- NEW: Filter out deleted/ghost sheets so they don't produce empty gap rows! ---
  settingsData = settingsData.filter(row => ss.getSheetByName(String(row[6]).trim()));

  // --- NEW: Group identical categories together to prevent duplicate headers ---
  settingsData.forEach((row, idx) => row.push(idx)); // Store original index to maintain stable sort
  settingsData.sort((a, b) => {
    const catA = String(a[5] || "").trim();
    const catB = String(b[5] || "").trim();
    if (catA === catB) {
      return a[a.length - 1] - b[b.length - 1]; // Preserve original order within the same category
    }
    if (catA === "Unassigned") return 1;
    if (catB === "Unassigned") return -1;
    return catA.localeCompare(catB, undefined, {numeric: true, sensitivity: 'base'});
  });
  // -----------------------------------------------------------------------------

  let lastSection = "";
  let grandTotalC = 0, grandTotalE = 0;
  
  // Subtotal Trackers
  let sectionTotalC = 0, sectionTotalE = 0;
  let lastSectionNeedsSubtotal = false;

  settingsData.forEach((row) => {
    // Determine checkbox state from Column C (index 1)
    const needsSubtotal = (row[1] === true || String(row[1]).toLowerCase() === 'true');
    const sectionFull = String(row[5]).trim(); // <-- Updated: Trimmed to prevent accidental duplicate sections
    const sheetName = String(row[6]).trim();   // <-- Updated: Trimmed
    const displayRoomName = row[7] ? String(row[7]).trim() : sheetName; // Col I

    if (sectionFull !== lastSection) {
      if (lastSection !== "") {
        if (lastSectionNeedsSubtotal) {
          // 1 Blank row BEFORE subtotal
          currentRowInPage++;
          checkAndTile();
          
          // Print subtotal
          const subRow = (currentPage * PAGE_HEIGHT) + DATA_START_REL + currentRowInPage;
          target.getRange(subRow, 2).setValue("Sub Total").setFontWeight("normal").setHorizontalAlignment("right");
          target.getRange(subRow, 3).setValue(sectionTotalC).setFontWeight("bold");
          target.getRange(subRow, 5).setValue(sectionTotalE).setFontWeight("bold");
          currentRowInPage++;
          checkAndTile();
          
          // 1 Blank row AFTER subtotal (before the next entry)
          currentRowInPage++;
        } else {
          // Default 1 blank row if no subtotal
          currentRowInPage++; 
        }
      }
      
      checkAndTile();
      const headerRow = (currentPage * PAGE_HEIGHT) + DATA_START_REL + currentRowInPage;
      const match = sectionFull.match(/^([\d.]+)\s+(.*)$/);
      target.getRange(headerRow, 1).setValue(match ? match[1] : "");
      target.getRange(headerRow, 2).setValue(match ? match[2] : sectionFull);
      target.getRange(headerRow, 1, 1, 5).setFontWeight("bold");
      currentRowInPage++;
      
      lastSection = sectionFull;
      lastSectionNeedsSubtotal = needsSubtotal;
      sectionTotalC = 0;
      sectionTotalE = 0;
    }
    
    checkAndTile();
    const targetRow = (currentPage * PAGE_HEIGHT) + DATA_START_REL + currentRowInPage;
    const source = ss.getSheetByName(sheetName);
    
    if (source) {
      const data = source.getRange(1, 1, source.getLastRow() || 1, 7).getValues();
      let sqm = 0;
      let gsm = ""; // Variable for CIRC Factor
      
      for (let i = 0; i < data.length; i++) {
        const colA = String(data[i][0]).trim();
        if (colA.includes("Total Departmental GSM")) sqm = Number(data[i][6]) || 0;
        // RE-IMPLEMENTED CIRC FACTOR LOGIC
        if (colA.includes("Circulation-net/gross factor")) gsm = data[i][6];
      }
      
      sectionTotalC += sqm;
      sectionTotalE += (sqm * grossMultiplier); // <-- UPDATED
      
      grandTotalC += sqm;
      grandTotalE += (sqm * grossMultiplier); // <-- UPDATED
      
      target.getRange(targetRow, 2).setValue(displayRoomName); 
      target.getRange(targetRow, 3).setValue(sqm); 
      // OUTPUT CIRC FACTOR TO COLUMN D (Col 4)
      target.getRange(targetRow, 4).setValue(gsm);      
      target.getRange(targetRow, 5).setValue(sqm * grossMultiplier);  // <-- UPDATED
    }
    currentRowInPage++;
  });

  // Evaluate final section subtotal after loop ends
  if (lastSection !== "" && lastSectionNeedsSubtotal) {
    // 1 Blank row BEFORE subtotal
    currentRowInPage++;
    checkAndTile();
    
    // Print subtotal
    const subRow = (currentPage * PAGE_HEIGHT) + DATA_START_REL + currentRowInPage;
    target.getRange(subRow, 2).setValue("Sub Total").setFontWeight("normal").setHorizontalAlignment("right");
    target.getRange(subRow, 3).setValue(sectionTotalC).setFontWeight("bold");
    target.getRange(subRow, 5).setValue(sectionTotalE).setFontWeight("bold");
    currentRowInPage++;
    checkAndTile();
    
    // 1 Blank row AFTER subtotal (before Grand Total block)
    currentRowInPage++;
  }

  const totalRow = (currentPage * PAGE_HEIGHT) + DATA_START_REL + currentRowInPage;
  target.getRange(totalRow, 1, 1, 6).setBackground("#969696").setFontWeight("bold").setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  target.getRange(totalRow, 2).setValue("GRAND TOTAL");
  target.getRange(totalRow, 3).setValue(grandTotalC);
  target.getRange(totalRow, 5).setValue(grandTotalE);

  const footerStartRow = (currentPage * PAGE_HEIGHT) + 30;
  if (footerStartRow > totalRow + 1) {
    target.getRange(totalRow + 1, 1, footerStartRow - (totalRow + 1), 6).setBorder(false, false, false, false, false, false).clearContent();
  }

  for (let p = 1; p <= currentPage; p++) {
    target.getRange((p * PAGE_HEIGHT) + PAGE_HEIGHT, 3).setValue("Page " + p + " of " + currentPage).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(9);
  }

  SpreadsheetApp.flush();
  const token = ScriptApp.getOAuthToken();
  const allBlobs = [];

  const exportLastRow = (currentPage * PAGE_HEIGHT) + PAGE_HEIGHT;
  const summaryUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${target.getSheetId()}&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&top_margin=0.25&bottom_margin=0.1&left_margin=0.5&right_margin=0.5&r1=0&r2=${exportLastRow}&c1=0&c2=8`;
  const summaryBlob = fetchWithRetry(summaryUrl, token);
  allBlobs.push(summaryBlob.setName("Summary.pdf"));

  const docTemplateId = "1OmwT0K6ODvevqy1qJUObwpa5hY0b9M3J4LFPeN7jr9k";
  const docBlob = DriveApp.getFileById(docTemplateId).getAs('application/pdf');

  // APPENDIX LOOP
  for (const name of appendixList) {
    const appSheet = ss.getSheetByName(name);
    if (appSheet) {
      const appLastRow = appSheet.getLastRow();
      
      // LOGIC CHANGE: Use 'name' (the Sheet Name) instead of Row 4
      const pageTitle = name; 
      
      // EXPORT 1: Full Page 1 (Rows 1-3 visible)
      appSheet.showRows(1, 3);
      appSheet.setFrozenRows(0);
      const urlFull = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${appSheet.getSheetId()}&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&title=false&top_margin=0.1&bottom_margin=1.2&left_margin=0.5&right_margin=0.5&r1=0&r2=${appLastRow}&c1=0&c2=8`;
      const blobFull = fetchWithRetry(urlFull, token);

      // EXPORT 2: Next Pages (Rows 1-3 hidden, Rows 4-5 frozen + spacer)
      appSheet.hideRows(1, 3);
      appSheet.insertRowAfter(5);
      appSheet.setRowHeight(6, 4);
      appSheet.setFrozenRows(6);
      const urlNext = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${appSheet.getSheetId()}&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&title=false&top_margin=0.1&bottom_margin=1.2&left_margin=0.5&right_margin=0.5&r1=0&r2=${appLastRow+1}&c1=0&c2=8`;
      const blobNext = fetchWithRetry(urlNext, token);

      // CLEANUP
      appSheet.setFrozenRows(0);
      appSheet.deleteRow(6);
      appSheet.showRows(1, 3);
      
      // --- NEW: Passing header variables to overlay function ---
      const overlaidBlob = await overlayAppendix(docBlob, blobFull, blobNext, pageTitle, headerDate, headerC3);
      // ---------------------------------------------------------
      allBlobs.push(overlaidBlob.setName(name + ".pdf"));
    }
  }

  const finalPdfBlob = await mergePDFs(allBlobs);
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: "Automated Summary Report",
    htmlBody: `<p>Report generated successfully.</p>`,
    attachments: [finalPdfBlob.setName("Full_Report.pdf")]
  });
}

/**
 * SPECIALIZED APPENDIX OVERLAY
 * --- NEW: Added headerDate and headerC3 to function parameters ---
 */
async function overlayAppendix(templateBlob, blobFull, blobNext, pageTitle, headerDate, headerC3) {
  const libUrl = "https://unpkg.com/pdf-lib/dist/pdf-lib.min.js";
  const response = UrlFetchApp.fetch(libUrl);
  eval(response.getContentText()); 

  const mainDoc = await PDFLib.PDFDocument.load(new Uint8Array(templateBlob.getBytes()));
  const docFull = await PDFLib.PDFDocument.load(new Uint8Array(blobFull.getBytes()));
  const docNext = await PDFLib.PDFDocument.load(new Uint8Array(blobNext.getBytes()));
  const mergedPdf = await PDFLib.PDFDocument.create();
  
  const helveticaBold = await mergedPdf.embedFont(PDFLib.StandardFonts.HelveticaBold);
  const helvetica = await mergedPdf.embedFont(PDFLib.StandardFonts.Helvetica); // <-- NEW: Regular font for the headers
  
  const [templatePage] = await mergedPdf.copyPages(mainDoc, [0]);
  const pagesFull = await mergedPdf.copyPages(docFull, [0]); 
  const pagesNext = await mergedPdf.copyPages(docNext, docNext.getPageIndices()); 
  
  const totalPages = pagesNext.length;

  for (let i = 0; i < totalPages; i++) {
    const newPage = mergedPdf.addPage([templatePage.getWidth(), templatePage.getHeight()]);
    const embeddedTemplate = await mergedPdf.embedPage(templatePage);
    newPage.drawPage(embeddedTemplate);

    const sPage = (i === 0) ? pagesFull[0] : pagesNext[i];
    const embeddedSheet = await mergedPdf.embedPage(sPage);
    
    const scale = Math.min(templatePage.getWidth() / sPage.getWidth(), 1);
    const xPos = (templatePage.getWidth() - (sPage.getWidth() * scale)) / 2;
    const yPos = templatePage.getHeight() - (sPage.getHeight() * scale) - 55;

    newPage.drawPage(embeddedSheet, {
      x: xPos,
      y: yPos,
      width: sPage.getWidth() * scale,
      height: sPage.getHeight() * scale,
    });

    const pageLabel = `Page ${i + 1} of ${totalPages}`;
    newPage.drawText(pageLabel, {
      x: templatePage.getWidth() / 2 - 35,
      y: 28,
      size: 8,
      font: helveticaBold,
      color: PDFLib.rgb(0, 0, 0),
    });

    // Drawing the Sheet Name Title on the rightmost footer
    if (pageTitle) {
      const fontSize = 8;
      const textWidth = helveticaBold.widthOfTextAtSize(pageTitle, fontSize);
      newPage.drawText(pageTitle, {
        x: templatePage.getWidth() - textWidth - 50, 
        y: 28,
        size: fontSize,
        font: helveticaBold,
        color: PDFLib.rgb(0, 0, 0),
      });
    }

    // --- NEW: Stamping the Header Data on Top Right (Using normal Helvetica font) ---
    const headerFontSize = 10;
    
    // Line 1: Date from C4 (MM/DD/YYYY format)
    if (headerDate) {
      const dateWidth = helvetica.widthOfTextAtSize(headerDate, headerFontSize);
      newPage.drawText(headerDate, {
        x: templatePage.getWidth() - dateWidth - 50, 
        y: templatePage.getHeight() - 30, // 30 points down from top
        size: headerFontSize,
        font: helvetica, // <-- Modified to use regular font
        color: PDFLib.rgb(0, 0, 0),
      });
    }

    // Line 2: Text from C3
    if (headerC3) {
      const c3Width = helvetica.widthOfTextAtSize(headerC3, headerFontSize);
      newPage.drawText(headerC3, {
        x: templatePage.getWidth() - c3Width - 50, 
        y: templatePage.getHeight() - 42, // 12 points below the date
        size: headerFontSize,
        font: helvetica, // <-- Modified to use regular font
        color: PDFLib.rgb(0, 0, 0),
      });
    }
    // ---------------------------------------------------
  }
  return Utilities.newBlob(await mergedPdf.save(), "application/pdf");
}

async function mergePDFs(blobs) {
  const response = UrlFetchApp.fetch("https://unpkg.com/pdf-lib/dist/pdf-lib.min.js");
  eval(response.getContentText()); 
  const mergedPdf = await PDFLib.PDFDocument.create();
  for (const blob of blobs) {
    const pdfDoc = await PDFLib.PDFDocument.load(new Uint8Array(blob.getBytes()));
    const pages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
    pages.forEach(p => mergedPdf.addPage(p));
  }
  return Utilities.newBlob(await mergedPdf.save(), "application/pdf");
}

function showSidebar() { var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Total Departmental GSM'); SpreadsheetApp.getUi().showSidebar(html); }