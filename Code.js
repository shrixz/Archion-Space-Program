/**
 * SHIM: Prevents "setTimeout is not defined" error from pdf-lib library
 */
var setTimeout = function(cb, ms) { Utilities.sleep(ms); cb(); };
var clearTimeout = function(id) {};

/**
 * RESET FILE — Clean Start
 *
 * Wipes derived/output state so the file is reusable for a new project while
 * preserving the template sheets and configuration the system depends on.
 *
 * DELETES:
 *   - Every generated department sheet (H1 tag === "GENERATED_DEPT")
 *   - "Summary", "Phasing Summary", "Appendix"
 *   - Any orphaned "Copy of Cover Template..." sheets left by a failed rebuild
 *
 * CLEARS:
 *   - Non-Standard Report rows 2+ (header row preserved)
 *   - Settings cols G + H from row 5 down (per-row category and sheet refs)
 *   - The GEN_SHEET_NAME_SNAPSHOT_V1 document property so rename detection
 *     starts fresh
 *
 * KEEPS:
 *   - Templates: Cover Template, Summary Template, Department Template,
 *     Standard Room Size, Header_Footer, Non-Standard Report (cleared),
 *     Settings (legend + headers kept; only rows 5+ cols G/H cleared)
 *
 * Finishes by calling refreshSheetNames() so Settings col A reflects the
 * surviving sheets immediately.
 */
function resetEverything() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const confirm = ui.alert(
    "Reset File — Clean Start",
    "This will permanently delete:\n" +
    "  • All generated department sheets\n" +
    "  • Summary, Phasing Summary, Appendix sheets\n" +
    "  • Non-Standard Report data (header kept)\n" +
    "  • Settings col G and col H (rows 5+)\n\n" +
    "Templates (Cover, Summary, Department, Standard Room Size, etc.) are kept.\n\n" +
    "Continue?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // 1. Delete generated / derived sheets.
  const derivedNames = new Set(["Summary", "Phasing Summary", "Appendix"]);

  ss.getSheets().forEach(sh => {
    const name = sh.getName();

    let isGen = false;
    try { if (sh.getRange("H1").getValue() === "GENERATED_DEPT") isGen = true; } catch (e) {}

    const isOrphanCopy =
      name === "Copy of Cover Template" || /^Copy of Cover Template \d+$/.test(name);

    if (isGen || derivedNames.has(name) || isOrphanCopy) {
      try { ss.deleteSheet(sh); } catch (e) {}
    }
  });

  // 2. Clear Non-Standard Report data (preserve header row 1).
  const report = ss.getSheetByName("Non-Standard Report");
  if (report) {
    const lr = report.getLastRow();
    if (lr >= 2) {
      report.getRange(2, 1, lr - 1, report.getMaxColumns()).clearContent();
    }
  }

  // 3. Clear Settings per-row dept assignments (cols G + H, rows 5+).
  //    Legend in col B and config in rows 1-4 are preserved.
  const settings = ss.getSheetByName("Settings");
  if (settings) {
    const lr = settings.getLastRow();
    if (lr >= 5) {
      settings.getRange(5, 7, lr - 4, 2).clearContent();
    }
  }

  // 4. Clear stored rename snapshot so the next snapshot rebuilds clean.
  try {
    PropertiesService.getDocumentProperties().deleteProperty("GEN_SHEET_NAME_SNAPSHOT_V1");
  } catch (e) {}

  // 5. Refresh Settings col A with current sheet names.
  try { refreshSheetNames(); } catch (e) {}

  ui.alert("Reset complete. The file is now clean and ready to use.");
}

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
    .addItem('Wrap Remarks (All Sheets)', 'wrapAllGeneratedRemarks')
    .addItem('Show Total GSM', 'showSidebar')
    .addSeparator()
    .addItem('Setup Auto-Refresh', 'setupOnChangeTrigger')
    .addItem('Grant Permissions', 'requestPermissions')
    .addSeparator()
    .addItem('Reset File (Clean Start)', 'resetEverything')
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

  _purgeCoverTemplateCopies(ss);

  // Defensive: ensure both templates are exactly REQUIRED_PAGE_HEIGHT rows.
  // Pad shorter templates so F33 reads don't fall off the end; trim trailing
  // blank, unmerged rows on longer templates (typically leftover padding from
  // a previous auto-pad cycle when PAGE_HEIGHT was higher). If real content
  // still sits beyond REQUIRED_PAGE_HEIGHT, throw a clear error — letting the
  // tile math run into it produces "you can't perform a paste that partially
  // intersects a merge" further down.
  const REQUIRED_PAGE_HEIGHT = 34;
  _padOrTrimTemplateRows(template, REQUIRED_PAGE_HEIGHT, "Summary Template");
  _padOrTrimTemplateRows(coverTemplate, REQUIRED_PAGE_HEIGHT, "Cover Template");

  const versionText = coverTemplate.getRange(10, 9).getDisplayValue();
  const dateText = coverTemplate.getRange(15, 9).getDisplayValue();
  const fullHeaderValue = versionText + "\n" + dateText;
  const hospitalName = template.getRange("F33").getDisplayValue();
  
  const existingSummary = ss.getSheetByName("Summary");

  // Snapshot col G notes keyed by (colA, colB) so they survive the rebuild.
  // Col G is user-typed notes; included in the PDF export and must persist.
  const noteSnapshot = {};
  if (existingSummary) {
    try {
      const sLastRow = existingSummary.getLastRow();
      const sMaxCols = existingSummary.getMaxColumns();
      if (sLastRow > 0 && sMaxCols >= 7) {
        const keyData = existingSummary.getRange(1, 1, sLastRow, 2).getValues();
        const noteData = existingSummary.getRange(1, 7, sLastRow, 1).getRichTextValues();
        for (let i = 0; i < keyData.length; i++) {
          const rt = noteData[i][0];
          const noteText = rt ? rt.getText() : "";
          if (noteText !== "") {
            const key = String(keyData[i][0]).trim() + "" + String(keyData[i][1]).trim();
            if (key !== "") noteSnapshot[key] = rt;
          }
        }
      }
    } catch (snapErr) { /* notes snapshot failed; continue with empty snapshot */ }
    ss.deleteSheet(existingSummary);
  }

  // Defensive: if setName fails (e.g. a concurrent rebuild already created
  // "Summary"), delete the orphan copy immediately so it does not linger as
  // "Copy of Cover Template N".
  const _tmpCopy = coverTemplate.copyTo(ss);
  let target;
  try {
    target = _tmpCopy.setName("Summary");
  } catch (nameErr) {
    try { ss.deleteSheet(_tmpCopy); } catch (e) {}
    throw nameErr;
  }

  // The notes column lives in col G but is owned by the user — not the Cover
  // Template — so the fresh copy may come back with fewer than 7 columns.
  // Ensure col G exists before restore, otherwise the snapshot has nowhere
  // to land and notes silently disappear on every rebuild.
  if (target.getMaxColumns() < 7) {
    target.insertColumnsAfter(target.getMaxColumns(), 7 - target.getMaxColumns());
  }

  const PAGE_HEIGHT = 34;
  const DATA_START_REL = 11;
  const DATA_END_REL = 30;
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

      // Break any merges in the destination block BEFORE the copy. Without
      // this, a merge inherited from the Cover Template (which was the
      // initial coverTemplate.copyTo source for `target`) that straddles the
      // copy boundary will throw "You can't perform a paste that partially
      // intersects a merge" on the copyTo line below.
      target.getRange(start, 1, PAGE_HEIGHT, 9).breakApart();

      template.getRange(1, 1, PAGE_HEIGHT, 9).copyTo(target.getRange(start, 1));
      for (let h = 0; h < templateRowHeights.length; h++) {
        target.setRowHeight(start + h, templateRowHeights[h]);
      }
      const headerCell = target.getRange(start, 6);
      headerCell.setValue(fullHeaderValue);
      headerCell.setWrap(true).setHorizontalAlignment("right").setVerticalAlignment("top");
      target.getRange(start + 32, 6).setValue(hospitalName);

      // breakApart again before clearContent so a Summary-Template merge that
      // partially overlaps the 5-column data window doesn't trip the same
      // partial-intersect rule.
      const dataArea = target.getRange(start + DATA_START_REL - 1, 1, ROWS_PER_PAGE, 5);
      dataArea.breakApart();
      dataArea.clearContent();
    }
  }

  checkAndTile();

  const settings = ss.getSheetByName("Settings");

  // --- AUTO-CLEAN SETTINGS: Clear cols G/H for rows whose sheet no longer exists ---
  // Only the per-row category (G) and sheet name (H) cells are cleared — the
  // full-row delete that used to live here also took out col B (category
  // legend) and col C (subtotal flag), which the user wants to keep intact.
  const _preCleanRow = settings.getLastRow();
  if (_preCleanRow >= 5) {
    const _colH = settings.getRange(5, 8, _preCleanRow - 4, 1).getValues();
    for (let i = _colH.length - 1; i >= 0; i--) {
      const _ref = String(_colH[i][0]).trim();
      if (_ref !== "" && !ss.getSheetByName(_ref)) {
        settings.getRange(5 + i, 7, 1, 2).clearContent();
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
  const headerC1 = settings.getRange("C1").getDisplayValue();

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

   try {
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
   } catch (deptErr) {
    throw new Error(
      "Summary build failed on Settings col H entry '" + sheetName +
      "' (section: '" + sectionFull + "'). Original: " +
      (deptErr && deptErr.message ? deptErr.message : deptErr)
    );
   }
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

  const footerStartRow = (currentPage * PAGE_HEIGHT) + 33;
  if (footerStartRow > totalRow + 1) {
    target.getRange(totalRow + 1, 1, footerStartRow - (totalRow + 1), 6).setBorder(false, false, false, false, false, false).clearContent();
  }

  // NOTE: cell-drawn "Page X of Y" footer is intentionally NOT written here —
  // the doc-template overlay (overlaySummary) draws the page label on top of
  // the template footer instead, matching the Appendix flow. Writing it to a
  // cell as well produced duplicate "Page 1 of 1" stamps on each PDF page.

  // Restore col G notes onto matching rows of the rebuilt Summary.
  if (Object.keys(noteSnapshot).length > 0) {
    try {
      const tLastRow = target.getLastRow();
      const tMaxCols = target.getMaxColumns();
      if (tLastRow > 0 && tMaxCols >= 7) {
        const tKeyData = target.getRange(1, 1, tLastRow, 2).getValues();
        for (let i = 0; i < tKeyData.length; i++) {
          const key = String(tKeyData[i][0]).trim() + "" + String(tKeyData[i][1]).trim();
          if (key !== "" && noteSnapshot.hasOwnProperty(key)) {
            target.getRange(i + 1, 7).setRichTextValue(noteSnapshot[key]);
          }
        }
      }
    } catch (restErr) { /* notes restore failed; sheet rebuilt cleanly without notes */ }
  }

  SpreadsheetApp.flush();
  const token = ScriptApp.getOAuthToken();
  const allBlobs = [];

  const exportLastRow = (currentPage * PAGE_HEIGHT) + PAGE_HEIGHT;
  const summaryUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${target.getSheetId()}&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&top_margin=0.25&bottom_margin=0.1&left_margin=0.5&right_margin=0.5&r1=0&r2=${exportLastRow}&c1=0&c2=7`;
  const summaryBlob = fetchWithRetry(summaryUrl, token);

  const docTemplateId = "1OmwT0K6ODvevqy1qJUObwpa5hY0b9M3J4LFPeN7jr9k";
  const docBlob = DriveApp.getFileById(docTemplateId).getAs('application/pdf');

  // Apply doc-template overlay to the Summary PDF starting on page 2
  // (page 1 is the cover and is passed through untouched).
  const overlaidSummaryBlob = await overlaySummary(docBlob, summaryBlob, "Summary", headerDate, headerC3, headerC1);
  allBlobs.push(overlaidSummaryBlob.setName("Summary.pdf"));

  // APPENDIX LOOP
  for (const name of appendixList) {
    const appSheet = ss.getSheetByName(name);
    if (appSheet) {
      const appLastRow = appSheet.getLastRow();

      // LOGIC CHANGE: Use 'name' (the Sheet Name) instead of Row 4.
      // String() guards against numeric sheet names (e.g. "1.1", "123") which
      // getValues() returns as Number and crashes PDFLib's drawText.
      const pageTitle = String(name);

      // EXPORT 1: Full Page 1 (Rows 1-3 visible)
      appSheet.showRows(1, 3);
      appSheet.setFrozenRows(0);
      const urlFull = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${appSheet.getSheetId()}&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&title=false&top_margin=0.1&bottom_margin=1.2&left_margin=0.5&right_margin=0.5&r1=0&r2=${appLastRow}&c1=0&c2=9`;
      const blobFull = fetchWithRetry(urlFull, token);

      // EXPORT 2: Next Pages (Rows 1-3 hidden, Rows 4-5 frozen + spacer)
      appSheet.hideRows(1, 3);
      appSheet.insertRowAfter(5);
      appSheet.setRowHeight(6, 4);
      appSheet.setFrozenRows(6);
      const urlNext = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${appSheet.getSheetId()}&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&title=false&top_margin=0.1&bottom_margin=1.2&left_margin=0.5&right_margin=0.5&r1=0&r2=${appLastRow+1}&c1=0&c2=9`;
      const blobNext = fetchWithRetry(urlNext, token);

      // CLEANUP
      appSheet.setFrozenRows(0);
      appSheet.deleteRow(6);
      appSheet.showRows(1, 3);

      // --- NEW: Passing header variables to overlay function ---
      const overlaidBlob = await overlayAppendix(docBlob, blobFull, blobNext, pageTitle, headerDate, headerC3, headerC1);
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
 * --- NEW: Added headerC1 — stamped directly above the sheet-name footer ---
 */
async function overlayAppendix(templateBlob, blobFull, blobNext, pageTitle, headerDate, headerC3, headerC1) {
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

    const sPage = pagesNext[i];
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
      const fontSize = 8.5;
      const textWidth = helveticaBold.widthOfTextAtSize(pageTitle, fontSize);
      newPage.drawText(pageTitle, {
        x: templatePage.getWidth() - textWidth - 50,
        y: 28,
        size: fontSize,
        font: helveticaBold,
        color: PDFLib.rgb(0, 0, 0),
      });
    }

    // NEW: Stamping Settings!C1 directly above the sheet-name in the footer.
    // Same x-anchor (right margin) so it sits visually above pageTitle.
    if (headerC1) {
      const c1FontSize = 8.5;
      const c1Width = helveticaBold.widthOfTextAtSize(headerC1, c1FontSize);
      newPage.drawText(headerC1, {
        x: templatePage.getWidth() - c1Width - 50,
        y: 40, // 12 points above the sheet-name line (y=28)
        size: c1FontSize,
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

/**
 * SUMMARY OVERLAY
 * Stamps the Google Doc footer template onto the Summary PDF starting from
 * page 2. Page 1 (the cover) is passed through untouched. Same stamping
 * style as overlayAppendix: pageTitle + Settings!C1 in the footer (lower
 * right, C1 directly above the tab name), plus C3 / date in the upper right.
 */
async function overlaySummary(templateBlob, summaryBlob, pageTitle, headerDate, headerC3, headerC1) {
  const libUrl = "https://unpkg.com/pdf-lib/dist/pdf-lib.min.js";
  const response = UrlFetchApp.fetch(libUrl);
  eval(response.getContentText());

  const mainDoc = await PDFLib.PDFDocument.load(new Uint8Array(templateBlob.getBytes()));
  const docSummary = await PDFLib.PDFDocument.load(new Uint8Array(summaryBlob.getBytes()));
  const mergedPdf = await PDFLib.PDFDocument.create();

  const helveticaBold = await mergedPdf.embedFont(PDFLib.StandardFonts.HelveticaBold);
  const helvetica = await mergedPdf.embedFont(PDFLib.StandardFonts.Helvetica);

  const [templatePage] = await mergedPdf.copyPages(mainDoc, [0]);
  const summaryPages = await mergedPdf.copyPages(docSummary, docSummary.getPageIndices());

  const totalPages = summaryPages.length;
  // Stamped pages = pages 2..N (cover excluded). Used for the "Page X of Y"
  // label so it counts only the stamped pages.
  const stampedTotal = Math.max(totalPages - 1, 0);

  for (let i = 0; i < totalPages; i++) {
    if (i === 0) {
      // Cover page — pass through untouched.
      mergedPdf.addPage(summaryPages[i]);
      continue;
    }

    const newPage = mergedPdf.addPage([templatePage.getWidth(), templatePage.getHeight()]);
    const embeddedTemplate = await mergedPdf.embedPage(templatePage);
    newPage.drawPage(embeddedTemplate);

    const sPage = summaryPages[i];
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

    const pageLabel = `Page ${i} of ${stampedTotal}`;
    newPage.drawText(pageLabel, {
      x: templatePage.getWidth() / 2 - 35,
      y: 28,
      size: 8,
      font: helveticaBold,
      color: PDFLib.rgb(0, 0, 0),
    });

    // Sheet-name footer (lower right)
    if (pageTitle) {
      const fontSize = 8.5;
      const textWidth = helveticaBold.widthOfTextAtSize(pageTitle, fontSize);
      newPage.drawText(pageTitle, {
        x: templatePage.getWidth() - textWidth - 50,
        y: 28,
        size: fontSize,
        font: helveticaBold,
        color: PDFLib.rgb(0, 0, 0),
      });
    }

    // Settings!C1 directly above the sheet-name in the footer.
    if (headerC1) {
      const c1FontSize = 8.5;
      const c1Width = helveticaBold.widthOfTextAtSize(headerC1, c1FontSize);
      newPage.drawText(headerC1, {
        x: templatePage.getWidth() - c1Width - 50,
        y: 40,
        size: c1FontSize,
        font: helveticaBold,
        color: PDFLib.rgb(0, 0, 0),
      });
    }

    // Top-right header (date + C3) — matches the appendix overlay style.
    const headerFontSize = 10;
    if (headerDate) {
      const dateWidth = helvetica.widthOfTextAtSize(headerDate, headerFontSize);
      newPage.drawText(headerDate, {
        x: templatePage.getWidth() - dateWidth - 50,
        y: templatePage.getHeight() - 30,
        size: headerFontSize,
        font: helvetica,
        color: PDFLib.rgb(0, 0, 0),
      });
    }
    if (headerC3) {
      const c3Width = helvetica.widthOfTextAtSize(headerC3, headerFontSize);
      newPage.drawText(headerC3, {
        x: templatePage.getWidth() - c3Width - 50,
        y: templatePage.getHeight() - 42,
        size: headerFontSize,
        font: helvetica,
        color: PDFLib.rgb(0, 0, 0),
      });
    }
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