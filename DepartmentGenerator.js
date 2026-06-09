/** * PART 1: THE MODIFIED GENERATOR
 */
function generateDeptMimicLobby() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName("Department Template");
  const standardSheet = ss.getSheetByName("Standard Room Size");
  const reportSheet = ss.getSheetByName("Non-Standard Report");

  const deptName = template.getRange("A5").getValue();
  if (!deptName || deptName === "" || deptName === "Department Title") {
    SpreadsheetApp.getUi().alert("Please enter a Department Title in cell A5!");
    return;
  }

  // --- REQ 1: CHECK FOR CATEGORY IN H1 BEFORE GENERATING ---
  const categoryName = template.getRange("H1").getValue();
  if (!categoryName || categoryName === "") {
    SpreadsheetApp.getUi().alert("Please select a Category from the dropdown in H1!");
    return;
  }

  // --- DYNAMICALLY FIND TEMPLATE BOUNDS ---
  // Scan column A to locate "Total NSM", which anchors the entire summary block.
  // This makes the generator resilient when users insert rows in the data area (9 to N).
  const startRow = 9;
  const tmplColAVals = template.getRange("A:A").getValues();
  let tmplLabelRowNSM = 0;
  for (let i = startRow - 1; i < tmplColAVals.length; i++) {
    if (tmplColAVals[i][0].toString().trim() === "Total NSM") {
      tmplLabelRowNSM = i + 1; // convert 0-indexed to 1-indexed row number
      break;
    }
  }
  const maxRoomRow      = tmplLabelRowNSM > 0 ? tmplLabelRowNSM - 1 : 28; // last data row
  const tmplCircFactorRow = tmplLabelRowNSM > 0 ? tmplLabelRowNSM + 1 : 30; // "Circ Factor" row

  // --- CHECK FOR CIRCULATION FACTOR (dynamic row) ---
  const factorVal = template.getRange(tmplCircFactorRow, 7).getValue();
  if (factorVal === "" || factorVal === null) {
    SpreadsheetApp.getUi().alert("Kindly indicate Circulation-net/gross factor in the Circulation Factor cell before generating");
    return;
  }

  const newSheet = template.copyTo(ss).setName(deptName);

  // --- REQ 2: CLEAR G1 AND H1 ON THE GENERATED SHEET ---
  newSheet.getRange("G1:H1").clearDataValidations().clearContent();

  // Remove buttons/drawings from the new sheet
  const drawings = newSheet.getDrawings();
  for (let i = 0; i < drawings.length; i++) {
    drawings[i].remove();
  }

  const standards = standardSheet.getRange("A2:E" + standardSheet.getLastRow()).getValues();

  let lastUsedRow = startRow;

  // 1. Process Rooms, Area (F), and Total NSM (G) — loop covers all data rows dynamically
  for (let i = startRow; i <= maxRoomRow; i++) {
    const roomName = newSheet.getRange(i, 1).getValue();
    const qty = newSheet.getRange(i, 2).getValue();

    if (roomName && !roomName.toString().includes("Total NSM")) {

      // --- REQ 8: SUB-HEADER / EMPTY LINE HANDLING ---
      if (qty === "" || qty === null || qty === 0) {
        newSheet.getRange(i, 1, 1, 8).setFontWeight("bold");
        newSheet.getRange(i, 6, 1, 2).clearContent(); // Clear formulas so it doesn't show 0
        continue;
      }

      lastUsedRow = i;
      newSheet.getRange(i, 1, 1, 8).setFontWeight("normal");
      newSheet.getRange(i, 4).setValue("x");
      newSheet.getRange(i, 6).setFormula(`=C${i} * E${i}`);
      newSheet.getRange(i, 7).setFormula(`=B${i} * F${i}`);

      const match = standards.find(s => s[1].toString().trim().toLowerCase() === roomName.toString().trim().toLowerCase());
      if (match) {
        const stdLen = match[2];
        const stdWid = match[3];
        const remarks = (match[4]);

        if (newSheet.getRange(i, 3).isBlank() || newSheet.getRange(i, 5).isBlank()) {
          newSheet.getRange(i, 3).setValue(stdLen).setFontColor("black"); // REQ 2: Default to black
          newSheet.getRange(i, 5).setValue(stdWid).setFontColor("black"); // REQ 2: Default to black
          newSheet.getRange(i, 8).setValue(remarks);
        } else {
          let cLen = newSheet.getRange(i, 3).getValue();
          let cWid = newSheet.getRange(i, 5).getValue();

          // REQ 1 & 2: Ensure color is set correctly on the new generated sheet
          newSheet.getRange(i, 3).setFontColor((cLen !== "" && cLen != stdLen) ? "red" : "black");
          newSheet.getRange(i, 5).setFontColor((cWid !== "" && cWid != stdWid) ? "red" : "black");

          if (cLen != stdLen || cWid != stdWid) {
            reportSheet.appendRow([new Date(), deptName, roomName, cLen + " x " + cWid, stdLen + " x " + stdWid]);
          }
        }
      }
    }
  }

  // 2. Find the Summary Table on the GENERATED sheet (re-scan since it's a copy of the template)
  const newSheetColAVals = newSheet.getRange("A:A").getValues();
  let labelRowNSM = 0;
  for (let i = 0; i < newSheetColAVals.length; i++) {
    if (newSheetColAVals[i][0].toString().trim() === "Total NSM") {
      labelRowNSM = i + 1;
      break;
    }
  }

  const rowCircFactor = labelRowNSM + 1;
  const rowCircNSM    = labelRowNSM + 2;
  const rowTotalGSM   = labelRowNSM + 3;
  const lastDataRow   = labelRowNSM - 1; // last actual data row (row just above "Total NSM")

  // 4. Summary Section Formulas — SUMIFS range is now fully dynamic
  newSheet.getRange(labelRowNSM, 7).setFormula(`=SUMIFS(G${startRow}:G${lastDataRow}, H${startRow}:H${lastDataRow}, "<>*Not a room*")`);

  const factorCell = newSheet.getRange(rowCircFactor, 7);
  factorCell.setNumberFormat("0%");
  factorCell.setFontWeight("bold");

  newSheet.getRange(rowCircNSM, 7).setFormula(`=ROUND(G${labelRowNSM} * G${rowCircFactor}, 0)`);
  newSheet.getRange(rowTotalGSM, 7).setFormula(`=G${labelRowNSM} + G${rowCircNSM}`);

  // Final Polish
  // --- REQ 5: .00 Area and NSM FORMATTING (covers all data rows dynamically) ---
  newSheet.getRange(startRow, 6, (lastDataRow - startRow + 1), 2).setNumberFormat("#,##0.00");
  newSheet.getRange(labelRowNSM, 7, 3, 1).setNumberFormat("#,##0.00");

  newSheet.getRange(rowTotalGSM, 1, 1, 7).setFontWeight("bold");
  newSheet.getRange(labelRowNSM, 1, 4, 7).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  // Enforce circ factor cell as percentage on the generated sheet (dynamic row)
  newSheet.getRange(rowCircFactor, 7).setNumberFormat("0%");

  // --- TEMPLATE CLEANUP (dynamic ranges) ---
  template.getRange("A5").setValue("Department Title").setFontColor("red").setFontWeight("bold");
  template.getRange(`A${startRow}:H${maxRoomRow}`).clearContent();
  template.getRange(`A${startRow}:H${maxRoomRow}`).setFontWeight("normal");
  template.getRange(`C${startRow}:E${maxRoomRow}`).setFontColor("black");

  // --- REQ 7 & NEW REQ 2: Reset Circ Factor back to 0% and Red in the Template (dynamic row) ---
  const tmplFactorCell = template.getRange(tmplCircFactorRow, 7);
  tmplFactorCell.setValue(0);
  tmplFactorCell.setFontColor("red");
  tmplFactorCell.setNumberFormat("0%");

  // --- RESTORE TEMPLATE TO DEFAULT LAYOUT (rows 9-28) ---
  // Delete only the extra rows (row 29 up to the row before "Total NSM").
  // Iterates in reverse so row numbers stay valid, and skips rows that contain
  // drawings (e.g. the generate button) so they are never accidentally removed.
  if (maxRoomRow > 28) {
    const drawings = template.getDrawings();
    const drawnRows = new Set(drawings.map(d => d.getContainerInfo().getAnchorRow()));
    for (let r = maxRoomRow; r >= 29; r--) {
      if (!drawnRows.has(r)) {
        template.deleteRow(r);
      }
    }
  }

  // --- REQ 3: UPDATE 'SETTINGS' AUTOMATICALLY ---
  const settingsSheet = ss.getSheetByName("Settings");
  if (settingsSheet) {
    // Isolating the check to Column G to prevent other columns from pushing the row down
    const colGValues = settingsSheet.getRange("G:G").getValues();
    let insertRow = 5; // Default starting row (assuming rows 1-4 are headers)

    // Loop backwards from the bottom of Column G to find the first actual value
    for (let i = colGValues.length - 1; i >= 4; i--) {
      if (colGValues[i][0] !== "") {
        insertRow = i + 2; // i is 0-indexed (so row is i+1), and we want the row AFTER that (+1)
        break;
      }
    }

    settingsSheet.getRange(insertRow, 7).setValue(categoryName); // Col G
    settingsSheet.getRange(insertRow, 8).setValue(deptName);     // Col H
    // Col I removed per user request
  }

  // --- REQ 4: AUTOMATED SUMMARY CREATION ---
  autoUpdateSummarySheet(deptName);

  // --- NEW REQUIREMENT: HIDE ROW 1 ---
  newSheet.hideRows(1);

  // --- NEW REQUIREMENT: REMOVE EXTRA COLUMNS (I TO END) ---
  const totalCols = newSheet.getMaxColumns();
  if (totalCols > 8) {
    newSheet.deleteColumns(9, totalCols - 8);
  }

  // TAG THE SHEET AS GENERATED (Placed securely in H1 so it doesn't auto-expand columns to ZZ!)
  newSheet.getRange("H1").setValue("GENERATED_DEPT").setFontColor("white");

  // --- NEW: Track this sheet's id->name so future renames can be detected and propagated ---
  _saveGenSheetSnapshot(_buildGenSheetSnapshot());

  ss.setActiveSheet(newSheet);
  SpreadsheetApp.getUi().alert("Finished! Sheet '" + deptName + "' is ready. Post-generation edits will be tracked, and the Summary has been automatically updated.");
}

/** * PART 2: THE CHANGE TRACKER (onEdit)
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();
  const ss = e.source;

  // ==========================================
  // BRANCH A: DEPARTMENT TEMPLATE LOGIC
  // ==========================================
  if (sheetName === "Department Template") {

    // NEW REQ: Cell A5 Tracker for Department Title
    if (row === 5 && col === 1) {
      range.setFontWeight("bold"); // Always keep bold
      const val = range.getValue().toString().trim();
      if (val === "" || val === "Department Title") {
        range.setFontColor("red");
      } else {
        range.setFontColor("black");
      }
      return;
    }

    // --- DYNAMICALLY FIND TEMPLATE BOUNDS FOR THIS EDIT ---
    // Scan column A once to locate "Total NSM" — this anchors all row references
    // so that any rows inserted in the data area (9 to N) are correctly handled.
    const tmplColA = sheet.getRange("A:A").getValues();
    let labelRowNSM = 0;
    const startRow = 9;
    for (let i = startRow - 1; i < tmplColA.length; i++) {
      if (tmplColA[i][0].toString().trim() === "Total NSM") {
        labelRowNSM = i + 1;
        break;
      }
    }
    const lastDataRow   = labelRowNSM > 0 ? labelRowNSM - 1 : 28;
    const rowCircFactor = labelRowNSM > 0 ? labelRowNSM + 1 : 30;
    const rowCircNSM    = labelRowNSM > 0 ? labelRowNSM + 2 : 31;
    const rowTotalGSM   = labelRowNSM > 0 ? labelRowNSM + 3 : 32;

    // REQ 2, 3, 5, 6, 8: Room Auto-fill, Live Formula Injection, and Custom Remarks formatting
    // Upper bound uses lastDataRow (dynamic) instead of hardcoded 28
    if (row >= startRow && row <= lastDataRow) {

      // Auto-populate ONLY when Column A is selected
      if (col === 1) {
        const roomName = range.getValue().toString().trim();
        if (roomName !== "") {
          const standardSheet = ss.getSheetByName("Standard Room Size");
          if (standardSheet) {
            const standards = standardSheet.getRange("A2:H" + standardSheet.getLastRow()).getValues();
            const match = standards.find(s => s[1] && s[1].toString().trim().toLowerCase() === roomName.toLowerCase());

            if (match) {
              sheet.getRange(row, 3).setValue(match[2]).setFontColor("black"); // Length (reset to black)
              sheet.getRange(row, 4).setValue("x");      // Auto-fill "x"
              sheet.getRange(row, 5).setValue(match[3]).setFontColor("black"); // Width (reset to black)
              sheet.getRange(row, 8).setValue(match[4]); // Remarks
            }
          }
        }
      }

      // Check Column A and B for live calculations (Triggered on Col 1 or 2 edit)
      if (col === 1 || col === 2) {
        const rName = sheet.getRange(row, 1).getValue();
        const qty = sheet.getRange(row, 2).getValue();

        if (rName !== "") {
          if (qty === "" || qty === null || qty === 0) {
            // It's a sub-header
            sheet.getRange(row, 1, 1, 8).setFontWeight("bold");
            sheet.getRange(row, 6, 1, 2).clearContent();
          } else {
            // It's a standard room
            sheet.getRange(row, 1, 1, 8).setFontWeight("normal");
            sheet.getRange(row, 6).setFormula(`=C${row} * E${row}`).setNumberFormat("#,##0.00");
            sheet.getRange(row, 7).setFormula(`=B${row} * F${row}`).setNumberFormat("#,##0.00");
          }
        }

        // Live calculation for the bottom Summary Block (fully dynamic range)
        if (labelRowNSM > 0) {
          sheet.getRange(labelRowNSM, 7).setFormula(`=SUMIFS(G${startRow}:G${lastDataRow}, H${startRow}:H${lastDataRow}, "<>*Not a room*")`).setNumberFormat("#,##0.00");
          sheet.getRange(rowCircNSM, 7).setFormula(`=ROUND(G${labelRowNSM} * G${rowCircFactor}, 0)`).setNumberFormat("#,##0.00");
          sheet.getRange(rowTotalGSM, 7).setFormula(`=G${labelRowNSM} + G${rowCircNSM}`).setNumberFormat("#,##0.00");
        }
      }

      // REQ 2: TEMPLATE DIMENSION COLOR TRACKER (Isolating Columns)
      if (col === 3 || col === 5) {
        const roomName = sheet.getRange(row, 1).getValue().toString().trim();
        if (roomName !== "") {
          const standardSheet = ss.getSheetByName("Standard Room Size");
          if (standardSheet) {
            const standards = standardSheet.getRange("A2:E" + standardSheet.getLastRow()).getValues();
            const match = standards.find(s => s[1] && s[1].toString().trim().toLowerCase() === roomName.toLowerCase());

            if (match) {
              const stdLen = match[2];
              const stdWid = match[3];

              if (col === 3) {
                const cLen = sheet.getRange(row, 3).getValue();
                sheet.getRange(row, 3).setFontColor((cLen !== "" && cLen != stdLen) ? "red" : "black");
              } else if (col === 5) {
                const cWid = sheet.getRange(row, 5).getValue();
                sheet.getRange(row, 5).setFontColor((cWid !== "" && cWid != stdWid) ? "red" : "black");
              }
            }
          }
        }
      }

      // NEW REQ: Manage ||| Marker based on ALL edits (Dimensions & Remarks)
      if (col === 3 || col === 5 || col === 8) {
        const roomName = sheet.getRange(row, 1).getValue().toString().trim();
        let stdRemark = "";
        let stdLen = "N/A";
        let stdWid = "N/A";

        const standardSheet = ss.getSheetByName("Standard Room Size");
        if (standardSheet && roomName !== "") {
          const standards = standardSheet.getRange("A2:E" + standardSheet.getLastRow()).getValues();
          const match = standards.find(s => s[1] && s[1].toString().trim().toLowerCase() === roomName.toLowerCase());
          if (match) {
            stdLen = match[2];
            stdWid = match[3];
            stdRemark = match[4].toString().trim();
          }
        }

        const cLen = sheet.getRange(row, 3).getValue();
        const cWid = sheet.getRange(row, 5).getValue();

        let isDimEdited = false;
        if (cLen !== "" && cLen != stdLen) isDimEdited = true;
        if (cWid !== "" && cWid != stdWid) isDimEdited = true;

        const remarksCell = sheet.getRange(row, 8);
        let currentRemarks = remarksCell.getValue().toString();
        const marker = "||| ";

        // Auto-Insert | only if user is actively typing in Col 8
        if (col === 8 && currentRemarks !== "") {
          if (currentRemarks.indexOf(marker) !== 0) {
            currentRemarks = marker + currentRemarks;
          }

          const textWithoutMarker = currentRemarks.substring(marker.length).trim();
          if (stdRemark !== "" && currentRemarks.indexOf("|", 3) === -1 && textWithoutMarker !== stdRemark) {
            if (textWithoutMarker.startsWith(stdRemark)) {
              const extra = textWithoutMarker.substring(stdRemark.length).trim();
              currentRemarks = marker + stdRemark + " | " + extra;
            } else {
              currentRemarks = marker + stdRemark + " | " + textWithoutMarker;
            }
          }
        }

        // Check text without marker to see if there is a custom comment
        let textWithoutMarker = currentRemarks.indexOf(marker) === 0 ? currentRemarks.substring(marker.length).trim() : currentRemarks.trim();

        // If they deleted the comment but left the pipe, clean it up
        if (textWithoutMarker === stdRemark + " |" || textWithoutMarker === stdRemark + "|") {
            textWithoutMarker = stdRemark;
        }

        let isRemarkEdited = false;
        if (textWithoutMarker !== stdRemark && textWithoutMarker !== "") {
            isRemarkEdited = true;
        } else if (textWithoutMarker.includes("|")) {
            isRemarkEdited = true;
        }

        // Add or remove marker based on whether the row has any edits
        if (isDimEdited || isRemarkEdited) {
            if (textWithoutMarker !== "") {
               currentRemarks = marker + textWithoutMarker;
            } else {
               currentRemarks = marker + stdRemark;
            }
        } else {
            // Nothing is edited, remove marker
            currentRemarks = textWithoutMarker;
        }

        // Apply formatting & Red text to custom comment
        const customPipeIndex = currentRemarks.indexOf("|", 3);

        if (customPipeIndex !== -1) {
          const richText = SpreadsheetApp.newRichTextValue().setText(currentRemarks);
          const defaultStyle = SpreadsheetApp.newTextStyle().setForegroundColor("black").build();
          const redStyle = SpreadsheetApp.newTextStyle().setForegroundColor("red").build();

          richText.setTextStyle(0, customPipeIndex, defaultStyle);
          richText.setTextStyle(customPipeIndex, currentRemarks.length, redStyle);
          remarksCell.setRichTextValue(richText.build());
        } else {
          const richText = SpreadsheetApp.newRichTextValue().setText(currentRemarks);
          const defaultStyle = SpreadsheetApp.newTextStyle().setForegroundColor("black").build();
          richText.setTextStyle(0, currentRemarks.length, defaultStyle);
          remarksCell.setRichTextValue(richText.build());
        }
      }
    }

    // REQ 7: Circulation Factor Color Toggle — uses dynamic rowCircFactor instead of hardcoded 30
    if (row === rowCircFactor && col === 7) {
      const val = range.getValue();
      // Keeps text red if it is exactly 0% (0) or empty. Otherwise turns black.
      if (val === "" || val === 0 || val === "0%") {
        range.setFontColor("red");
      } else {
        range.setFontColor("black");
      }
    }
    return; // Stop execution here for the Template
  }

  // ==========================================
  // BRANCH B: GENERATED SHEET LOGIC
  // ==========================================
  let isGenerated = false;
  try { if (sheet.getRange("ZZ1").getValue() === "GENERATED_DEPT") isGenerated = true; } catch(e) {}
  try { if (sheet.getRange("H1").getValue() === "GENERATED_DEPT") isGenerated = true; } catch(e) {}

  if (!isGenerated) return;
  const excludedSheets = ["Department Template", "Standard Room Size", "Non-Standard Report", "Instructions"];
  if (excludedSheets.indexOf(sheetName) > -1) return;

  // --- DYNAMICALLY FIND GENERATED SHEET BOUNDS FOR THIS EDIT ---
  // Scan column A to locate "Total NSM" — handles rows inserted by users after generation.
  const genColA = sheet.getRange("A:A").getValues();
  let genLabelRowNSM = 0;
  const genStartRow = 9;
  for (let i = genStartRow - 1; i < genColA.length; i++) {
    if (genColA[i][0].toString().trim() === "Total NSM") {
      genLabelRowNSM = i + 1;
      break;
    }
  }
  const genLastDataRow   = genLabelRowNSM > 0 ? genLabelRowNSM - 1 : 28;
  const genRowCircFactor = genLabelRowNSM > 0 ? genLabelRowNSM + 1 : 30;
  const genRowCircNSM    = genLabelRowNSM > 0 ? genLabelRowNSM + 2 : 31;
  const genRowTotalGSM   = genLabelRowNSM > 0 ? genLabelRowNSM + 3 : 32;

  // Track changes only within the room data range (dynamic upper bound)
  if (row >= genStartRow && row <= genLastDataRow && col >= 1 && col <= 8) {

    // REQ 1: COPIED TEMPLATE LOGIC SO GENERATED SHEET CALCULATES WHEN QTY IS EDITED
    if (col === 1) {
      const roomName = range.getValue().toString().trim();
      if (roomName !== "") {
        const standardSheet = ss.getSheetByName("Standard Room Size");
        if (standardSheet) {
          const standards = standardSheet.getRange("A2:E" + standardSheet.getLastRow()).getValues();
          const match = standards.find(s => s[1] && s[1].toString().trim().toLowerCase() === roomName.toLowerCase());

          if (match) {
            sheet.getRange(row, 3).setValue(match[2]).setFontColor("black"); // Length
            sheet.getRange(row, 4).setValue("x");      // Auto-fill "x"
            sheet.getRange(row, 5).setValue(match[3]).setFontColor("black"); // Width
            sheet.getRange(row, 8).setValue(match[4]); // Remarks
          }
        }
      }
    }

    // Check Column A and B for live calculations in Generated Sheet
    if (col === 1 || col === 2) {
      const rName = sheet.getRange(row, 1).getValue();
      const qty = sheet.getRange(row, 2).getValue();

      if (rName !== "") {
        if (qty === "" || qty === null || qty === 0) {
          // It's a sub-header
          sheet.getRange(row, 1, 1, 8).setFontWeight("bold");
          sheet.getRange(row, 6, 1, 2).clearContent();
        } else {
          // It's a standard room
          sheet.getRange(row, 1, 1, 8).setFontWeight("normal");
          sheet.getRange(row, 6).setFormula(`=C${row} * E${row}`).setNumberFormat("#,##0.00");
          sheet.getRange(row, 7).setFormula(`=B${row} * F${row}`).setNumberFormat("#,##0.00");
        }
      }

      // Live calculation for the bottom Summary Block (fully dynamic range)
      if (genLabelRowNSM > 0) {
        sheet.getRange(genLabelRowNSM, 7).setFormula(`=SUMIFS(G${genStartRow}:G${genLastDataRow}, H${genStartRow}:H${genLastDataRow}, "<>*Not a room*")`).setNumberFormat("#,##0.00");
        sheet.getRange(genRowCircNSM, 7).setFormula(`=ROUND(G${genLabelRowNSM} * G${genRowCircFactor}, 0)`).setNumberFormat("#,##0.00");
        sheet.getRange(genRowTotalGSM, 7).setFormula(`=G${genLabelRowNSM} + G${genRowCircNSM}`).setNumberFormat("#,##0.00");
      }
    }

    // NEW REQ: Manage ||| Marker based on ALL edits (Dimensions & Remarks) in Generated Sheet
    if (col === 3 || col === 5 || col === 8) {
      const roomName = sheet.getRange(row, 1).getValue().toString().trim();
      let stdRemark = "";
      let stdLen = "N/A";
      let stdWid = "N/A";

      const standardSheet = ss.getSheetByName("Standard Room Size");
      if (standardSheet && roomName !== "") {
        const standards = standardSheet.getRange("A2:E" + standardSheet.getLastRow()).getValues();
        const match = standards.find(s => s[1] && s[1].toString().trim().toLowerCase() === roomName.toLowerCase());
        if (match) {
          stdLen = match[2];
          stdWid = match[3];
          stdRemark = match[4].toString().trim();
        }
      }

      const cLen = sheet.getRange(row, 3).getValue();
      const cWid = sheet.getRange(row, 5).getValue();

      let isDimEdited = false;
      let bStdLen = stdLen !== "N/A" ? Math.round(stdLen) : "N/A";
      let bStdWid = stdWid !== "N/A" ? Math.round(stdWid) : "N/A";

      if (cLen !== "" && cLen != bStdLen) isDimEdited = true;
      if (cWid !== "" && cWid != bStdWid) isDimEdited = true;

      const remarksCell = sheet.getRange(row, 8);
      let currentRemarks = remarksCell.getValue().toString();
      const marker = "||| ";

      // Auto-Insert | only if user is actively typing in Col 8
      if (col === 8 && currentRemarks !== "") {
        if (currentRemarks.indexOf(marker) !== 0) {
          currentRemarks = marker + currentRemarks;
        }

        const textWithoutMarker = currentRemarks.substring(marker.length).trim();
        if (stdRemark !== "" && currentRemarks.indexOf("|", 3) === -1 && textWithoutMarker !== stdRemark) {
          if (textWithoutMarker.startsWith(stdRemark)) {
            const extra = textWithoutMarker.substring(stdRemark.length).trim();
            currentRemarks = marker + stdRemark + " | " + extra;
          } else {
            currentRemarks = marker + stdRemark + " | " + textWithoutMarker;
          }
        }
      }

      // Check text without marker to see if there is a custom comment
      let textWithoutMarker = currentRemarks.indexOf(marker) === 0 ? currentRemarks.substring(marker.length).trim() : currentRemarks.trim();

      // If they deleted the comment but left the pipe, clean it up
      if (textWithoutMarker === stdRemark + " |" || textWithoutMarker === stdRemark + "|") {
          textWithoutMarker = stdRemark;
      }

      let isRemarkEdited = false;
      if (textWithoutMarker !== stdRemark && textWithoutMarker !== "") {
          isRemarkEdited = true;
      } else if (textWithoutMarker.includes("|")) {
          isRemarkEdited = true;
      }

      // Add or remove marker based on whether the row has any edits
      if (isDimEdited || isRemarkEdited) {
          if (textWithoutMarker !== "") {
             currentRemarks = marker + textWithoutMarker;
          } else {
             currentRemarks = marker + stdRemark;
          }
      } else {
          // Nothing is edited, remove marker
          currentRemarks = textWithoutMarker;
      }

      // Apply formatting & Red text to custom comment
      const customPipeIndex = currentRemarks.indexOf("|", 3);

      if (customPipeIndex !== -1) {
        const richText = SpreadsheetApp.newRichTextValue().setText(currentRemarks);
        const defaultStyle = SpreadsheetApp.newTextStyle().setForegroundColor("black").build();
        const redStyle = SpreadsheetApp.newTextStyle().setForegroundColor("red").build();

        richText.setTextStyle(0, customPipeIndex, defaultStyle);
        richText.setTextStyle(customPipeIndex, currentRemarks.length, redStyle);
        remarksCell.setRichTextValue(richText.build());
      } else {
        const richText = SpreadsheetApp.newRichTextValue().setText(currentRemarks);
        const defaultStyle = SpreadsheetApp.newTextStyle().setForegroundColor("black").build();
        richText.setTextStyle(0, currentRemarks.length, defaultStyle);
        remarksCell.setRichTextValue(richText.build());
      }
    }

    // 4. UPDATE REPORT & DIMENSION COLOR (Triggered ONLY if Column 3 or 5 is edited)
    if (col === 3 || col === 5) {
      const reportSheet = ss.getSheetByName("Non-Standard Report");
      const standardSheet = ss.getSheetByName("Standard Room Size");

      const deptName = sheet.getRange("A5").getValue(); // Reverted to A5
      const roomName = sheet.getRange(row, 1).getValue();

      if (roomName && roomName.toString().trim() !== "") {
        let stdLen = "N/A";
        let stdWid = "N/A";

        if (standardSheet) {
          const standards = standardSheet.getRange("A2:C" + standardSheet.getLastRow()).getValues();
          const match = standards.find(s => s[0].toString().trim().toLowerCase() === roomName.toString().trim().toLowerCase());
          if (match) {
            stdLen = Math.round(match[1]);
            stdWid = Math.round(match[2]);
          }
        }

        // REQ 2: GENERATED SHEET DIMENSION COLOR TRACKER (Isolating Columns)
        let cLen = sheet.getRange(row, 3).getValue();
        let cWid = sheet.getRange(row, 5).getValue();

        if (col === 3) {
          sheet.getRange(row, 3).setFontColor((cLen !== "" && cLen != stdLen) ? "red" : "black");
        } else if (col === 5) {
          sheet.getRange(row, 5).setFontColor((cWid !== "" && cWid != stdWid) ? "red" : "black");
        }

        if (cLen != stdLen || cWid != stdWid) {
          if (reportSheet) {
            const reportData = reportSheet.getDataRange().getValues();
            let foundRow = -1;

            for (let i = 0; i < reportData.length; i++) {
              if (reportData[i][1] === deptName && reportData[i][2] === roomName) {
                foundRow = i + 1;
                break;
              }
            }

            if (foundRow !== -1) {
              reportSheet.getRange(foundRow, 1).setValue(new Date());
              reportSheet.getRange(foundRow, 4).setValue(cLen + " x " + cWid);
              reportSheet.getRange(foundRow, 5).setValue(stdLen + " x " + stdWid);
            } else {
              reportSheet.appendRow([new Date(), deptName, roomName, cLen + " x " + cWid, stdLen + " x " + stdWid]);
            }
          }
        }
      }
    }
  }

  // REQ 7 (Mimicked for Generated Sheet): Circulation Factor Color Toggle — uses dynamic genRowCircFactor
  if (row === genRowCircFactor && col === 7) {
    const val = range.getValue();
    // Keeps text red if it is exactly 0% (0) or empty. Otherwise turns black.
    if (val === "" || val === 0 || val === "0%") {
      range.setFontColor("red");
    } else {
      range.setFontColor("black");
    }
  }

  // --- REQ 4: REAL-TIME SUMMARY UPDATE ON GENERATED SHEET EDIT ---
  if (isGenerated) {
    SpreadsheetApp.flush(); // Ensure sheet calculations are finished
    autoUpdateSummarySheet(); // Re-build the summary to reflect new changes instantly
  }
}

/** * PART 3: SHEET DELETION LISTENER
 * onChange is a reserved simple trigger name — Google fires it automatically
 * on any structural change (row/column inserts, sheet add/delete, etc.).
 * No installation or setup needed.
 */
function onChange(e) {
  if (!e || !e.changeType) return;

  // NEW: Detect when a generated sheet was renamed (e.g., "Shane" -> "Shane First").
  // If so, update Settings col H to match the new name and refresh the Summary.
  const renameHandled = detectGeneratedSheetRenames();

  // Existing behavior: refresh summary when a sheet is deleted
  if (e.changeType === "REMOVE_GRID" && !renameHandled) autoUpdateSummarySheet();
}

/**
 * Snapshot helpers — track {sheetId: name} for sheets tagged as GENERATED_DEPT.
 * Sheet IDs are stable across renames, so comparing the current name against
 * the stored name for the same ID lets us detect a rename event.
 */
function _GEN_SNAPSHOT_KEY() { return "GEN_SHEET_NAME_SNAPSHOT_V1"; }

function _getGenSheetSnapshot() {
  const raw = PropertiesService.getDocumentProperties().getProperty(_GEN_SNAPSHOT_KEY());
  return raw ? JSON.parse(raw) : null;
}

function _saveGenSheetSnapshot(map) {
  PropertiesService.getDocumentProperties().setProperty(_GEN_SNAPSHOT_KEY(), JSON.stringify(map));
}

function _buildGenSheetSnapshot() {
  const ss = SpreadsheetApp.getActive();
  const map = {};
  ss.getSheets().forEach(sh => {
    let tag = "";
    try { tag = sh.getRange("H1").getValue(); } catch (e) {}
    if (tag === "GENERATED_DEPT") {
      map[String(sh.getSheetId())] = sh.getName();
    }
  });
  return map;
}

/**
 * NEW: Seed the rename snapshot. Safe to call anytime (e.g., from onOpen).
 * This guarantees the {sheetId: name} baseline exists for EXISTING generated
 * sheets, so the very first rename you make is detected and propagated even if
 * you have not generated a new department since installing this code.
 */
function seedGenSheetSnapshot() {
  _saveGenSheetSnapshot(_buildGenSheetSnapshot());
}

/**
 * Detect renames of generated sheets, propagate to Settings col H, and refresh Summary.
 * Returns true if a rename was handled (so onChange can skip a duplicate summary refresh).
 */
function detectGeneratedSheetRenames() {
  const ss = SpreadsheetApp.getActive();
  let snapshot = _getGenSheetSnapshot();

  // First-run bootstrap: no snapshot yet — populate and exit silently
  if (snapshot === null) {
    _saveGenSheetSnapshot(_buildGenSheetSnapshot());
    return false;
  }

  const settings = ss.getSheetByName("Settings");
  let renamed = false;

  ss.getSheets().forEach(sh => {
    let tag = "";
    try { tag = sh.getRange("H1").getValue(); } catch (e) {}
    if (tag !== "GENERATED_DEPT") return;

    const id = String(sh.getSheetId());
    const currentName = sh.getName();
    const oldName = snapshot[id];

    if (oldName && oldName !== currentName) {
      // Update Settings col H wherever the old name appears
      if (settings) {
        const lastRow = settings.getLastRow();
        if (lastRow >= 5) {
          const colH = settings.getRange(5, 8, lastRow - 4, 1).getValues();
          for (let i = 0; i < colH.length; i++) {
            if (String(colH[i][0]).trim() === oldName) {
              settings.getRange(5 + i, 8).setValue(currentName);
            }
          }
        }
      }
      renamed = true;
    }

    snapshot[id] = currentName;
  });

  // Prune snapshot entries for sheets that no longer exist
  const currentIds = new Set(ss.getSheets().map(s => String(s.getSheetId())));
  Object.keys(snapshot).forEach(id => {
    if (!currentIds.has(id)) delete snapshot[id];
  });

  _saveGenSheetSnapshot(snapshot);

  if (renamed) {
    SpreadsheetApp.flush();
    autoUpdateSummarySheet();
  }

  return renamed;
}

/**
 * Installs an installable onChange trigger as a fallback in case the simple
 * trigger does not fire in certain deployment contexts.
 * Run once from Report Automation > Setup Auto-Refresh.
 */
function setupOnChangeTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alreadySet = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "onChange");
  if (!alreadySet) {
    ScriptApp.newTrigger("onChange").forSpreadsheet(ss).onChange().create();
    // NEW: seed the baseline immediately so the first rename after setup is caught
    seedGenSheetSnapshot();
    SpreadsheetApp.getUi().alert("Done! The Summary will now auto-refresh whenever a sheet is renamed or deleted.");
  } else {
    // NEW: refresh the baseline in case sheets changed before this re-run
    seedGenSheetSnapshot();
    SpreadsheetApp.getUi().alert("Auto-refresh is already set up.");
  }
}

/** * PART 4: AUTOMATED SUMMARY GENERATOR (Mimics code.gs but purely for sheet update)
 */
function autoUpdateSummarySheet(newSheetName) {
  const ss = SpreadsheetApp.getActive();
  const template = ss.getSheetByName("Summary Template");
  const coverTemplate = ss.getSheetByName("Cover Template");

  if (!template || !coverTemplate) return;

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

  // --- UPDATED: Fetch Gross Multiplier from D2 ---
  const grossMultiplier = Number(settings.getRange("D2").getValue()) || 1.25;
  // --------------------------------------------------------

  // Read Columns B to I
  let settingsData = (lastRow < 5) ? [] : settings.getRange(5, 2, lastRow - 4, 8).getValues().filter(row => String(row[6]).trim() !== "");

  // Check if newly generated sheet is inside the Settings list
  if (newSheetName) {
    const isSheetInSettings = settingsData.some(row => String(row[6]).trim() === String(newSheetName).trim());
    if (!isSheetInSettings) {
      // If missing, append to the end under "Unassigned"
      // Indices map to: 0(B), 1(C:Checkbox), 5(G:Section), 6(H:Sheet), 7(I:Display)
      settingsData.push(["", false, "", "", "", "Unassigned", newSheetName, ""]);
    }
  }

  // --- FIX: Filter out deleted/ghost sheets so they don't produce empty gap rows! ---
  settingsData = settingsData.filter(row => ss.getSheetByName(String(row[6]).trim()));

  // --- FIX: Group identical categories together to prevent duplicate headers ---
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

  let lastSection = "";
  let grandTotalC = 0, grandTotalE = 0;

  let sectionTotalC = 0, sectionTotalE = 0;
  let lastSectionNeedsSubtotal = false;

  settingsData.forEach((row) => {
    const needsSubtotal = (row[1] === true || String(row[1]).toLowerCase() === 'true');
    const sectionFull = String(row[5]).trim();
    const sheetName = String(row[6]).trim();
    const displayRoomName = row[7] ? String(row[7]).trim() : sheetName;

    if (sectionFull !== lastSection) {
      if (lastSection !== "") {
        if (lastSectionNeedsSubtotal) {
          currentRowInPage++;
          checkAndTile();

          const subRow = (currentPage * PAGE_HEIGHT) + DATA_START_REL + currentRowInPage;
          target.getRange(subRow, 2).setValue("Sub Total").setFontWeight("normal").setHorizontalAlignment("right");
          target.getRange(subRow, 3).setValue(sectionTotalC).setFontWeight("bold");
          target.getRange(subRow, 5).setValue(sectionTotalE).setFontWeight("bold");
          currentRowInPage++;
          checkAndTile();

          currentRowInPage++;
        } else {
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

    // Because we filtered settingsData above, we are 100% sure this source exists.
    const source = ss.getSheetByName(sheetName);

    if (source) {
      const data = source.getRange(1, 1, source.getLastRow() || 1, 7).getValues();
      let sqm = 0;
      let gsm = "";

      for (let i = 0; i < data.length; i++) {
        const colA = String(data[i][0]).trim();
        if (colA.includes("Total Departmental GSM")) sqm = Number(data[i][6]) || 0;
        if (colA.includes("Circulation-net/gross factor")) gsm = data[i][6];
      }

      sectionTotalC += sqm;
      sectionTotalE += (sqm * grossMultiplier); // <-- UPDATED

      grandTotalC += sqm;
      grandTotalE += (sqm * grossMultiplier); // <-- UPDATED

      target.getRange(targetRow, 2).setValue(displayRoomName);
      target.getRange(targetRow, 3).setValue(sqm);
      target.getRange(targetRow, 4).setValue(gsm);
      target.getRange(targetRow, 5).setValue(sqm * grossMultiplier); // <-- UPDATED
    }
    currentRowInPage++;
  });

  if (lastSection !== "" && lastSectionNeedsSubtotal) {
    currentRowInPage++;
    checkAndTile();

    const subRow = (currentPage * PAGE_HEIGHT) + DATA_START_REL + currentRowInPage;
    target.getRange(subRow, 2).setValue("Sub Total").setFontWeight("normal").setHorizontalAlignment("right");
    target.getRange(subRow, 3).setValue(sectionTotalC).setFontWeight("bold");
    target.getRange(subRow, 5).setValue(sectionTotalE).setFontWeight("bold");
    currentRowInPage++;
    checkAndTile();

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
}