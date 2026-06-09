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
    .addItem('Refresh Rooms from Master List', 'refreshAllDepartmentsFromMasterList')
    .addItem('Show Total GSM', 'showSidebar')
    .addSeparator()
    .addItem('Setup Auto-Refresh', 'setupOnChangeTrigger')
    .addItem('Grant Permissions', 'requestPermissions')
    .addItem('Debug Sub-Totals', 'debugSubTotals')
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
  const settings = ss.getSheetByName("Settings");
  if (!settings) throw new Error("'Settings' sheet not found.");

  // --- Header data for the Appendix PDF overlay (read BEFORE Summary build) ---
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

  // --- BUILD THE SUMMARY ---
  // Delegate to the single shared Summary builder in DepartmentGenerator.js so
  // this menu-triggered flow and the auto-triggered onEdit/onChange/generation
  // flow produce IDENTICAL output (Sub-Total rows from Settings col C,
  // category sorting, Grand Total, page layout, page numbers). There is now
  // only one place to edit if Summary logic needs to change.
  autoUpdateSummarySheet();

  const target = ss.getSheetByName("Summary");
  if (!target) throw new Error("Summary build failed: 'Summary' sheet not created.");

  // After the builder finishes, the last row of the Summary is the per-page
  // footer of the final page — exactly what the original PDF export used.
  const exportLastRow = target.getLastRow();

  // Appendix list is read AFTER autoUpdateSummarySheet so we use the cleaned
  // Settings (rows whose sheets no longer exist are removed by the builder).
  const lastRow = settings.getLastRow();
  const appendixList = (lastRow < 5)
    ? []
    : settings.getRange(5, 8, lastRow - 4, 1).getValues().flat()
        .filter(name => name !== "" && ss.getSheetByName(name));

  const token = ScriptApp.getOAuthToken();
  const allBlobs = [];

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