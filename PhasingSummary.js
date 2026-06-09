/**
 * PHASING SUMMARY — SELF-CONTAINED MODULE
 *
 * Builds a "Phasing Summary" sheet that is LIVE-LINKED to each department sheet
 * by formula, separate from the regular "Summary" flow. Construction teams
 * distribute each department's total area across construction phases (Phase 1,
 * Phase 2) in the editable phase columns; the area numbers themselves stay
 * live, so any later edit to a department reflects automatically.
 *
 * Column layout (matches the team's printed template):
 *   A: DEPARTMENT             (and section names as bold header rows)
 *   B: DEPARTMENTAL AREA (SQM)   <- live formula, pulls "Total Departmental GSM"
 *   C: GROSS FLOOR AREA (SQM)    <- = B × Settings!D2 (fallback 1.25)
 *   D: PHASE 1                <- editable
 *   E: PHASE 2                <- editable
 *   F: REMARKS                <- editable
 *
 * This file is intentionally isolated from the rest of the codebase:
 *   - Does NOT modify Code.js, DepartmentGenerator.js, or any other file.
 *   - Does NOT touch the regular "Summary" sheet or its build flow.
 *   - Adds its own "Phasing" menu via an installable onOpen trigger so the
 *     existing simple onOpen in Code.js is not disturbed.
 *
 * ONE-TIME SETUP
 *   1. clasp push.
 *   2. Open the Apps Script editor (Extensions > Apps Script).
 *   3. Run setupPhasingMenu once (approve any permission prompt).
 *   4. Reload the spreadsheet. A new "Phasing" menu appears next to
 *      "Report Automation".
 *
 * EVERYDAY USE
 *   Phasing > Build Phasing Summary
 *     - Creates (or rebuilds) the "Phasing Summary" sheet.
 *     - Existing PHASE values AND REMARKS are PRESERVED across rebuilds,
 *       keyed by the department name in column A.
 */

const PHASING_SHEET_NAME = "Phasing Summary";
const PHASING_NUM_PHASES = 2;

/* ===========================================================================
 * MENU INSTALLATION
 * ========================================================================== */

function setupPhasingMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log("No active spreadsheet bound to this script. Open the script via the spreadsheet's Extensions > Apps Script menu and try again.");
    return;
  }

  // The main job: install the installable onOpen trigger so the Phasing menu
  // appears every time this spreadsheet is opened. Everything below is just
  // convenience and may fail silently when invoked from the editor.
  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === "_phasingInstallableOnOpen"
  );
  if (!exists) {
    ScriptApp.newTrigger("_phasingInstallableOnOpen")
      .forSpreadsheet(ss).onOpen().create();
  }

  // SpreadsheetApp.getUi() throws "Cannot call ... from this context" when run
  // from the Apps Script editor's Run button, so wrap both UI calls.
  try { _phasingInstallableOnOpen(); } catch (e) {}
  try {
    SpreadsheetApp.getUi().alert(
      "Phasing menu installed.\n\n" +
      "Reload the spreadsheet — a new 'Phasing' menu will appear next to " +
      "'Report Automation'."
    );
  } catch (e) {
    Logger.log("Phasing menu trigger installed. Reload the spreadsheet to see the 'Phasing' menu.");
  }
}

function _phasingInstallableOnOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Phasing")
      .addItem("Build Phasing Summary", "buildPhasingSummary")
      .addToUi();
  } catch (e) { /* no UI during automated runs */ }
}

/* ===========================================================================
 * MAIN BUILDER
 * ========================================================================== */

function buildPhasingSummary() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const settings = ss.getSheetByName("Settings");
  if (!settings) { ui.alert("'Settings' sheet not found."); return; }

  const lastRow = settings.getLastRow();
  if (lastRow < 5) { ui.alert("Settings has no department rows (row 5+)."); return; }

  // Column layout (1-indexed)
  const COL_DEPT = 1;
  const COL_DEPT_AREA = 2;
  const COL_GROSS = 3;
  const COL_PHASE_FIRST = 4;
  const COL_PHASE_LAST = COL_PHASE_FIRST + PHASING_NUM_PHASES - 1;
  const COL_REMARKS = COL_PHASE_LAST + 1;
  const TOTAL_COLS = COL_REMARKS;

  // Preserve previously-typed Phase values and Remarks, keyed by department.
  const preserved = _readPhasingPreserved(ss);
  const existing = ss.getSheetByName(PHASING_SHEET_NAME);
  if (existing) ss.deleteSheet(existing);

  // Read & sort the Settings rows the same way the regular Summary does.
  let data = settings.getRange(5, 2, lastRow - 4, 8).getValues()
    .filter(r => String(r[6]).trim() && ss.getSheetByName(String(r[6]).trim()));
  data.forEach((r, idx) => r.push(idx));
  data.sort((a, b) => {
    const catA = String(a[5] || "").trim();
    const catB = String(b[5] || "").trim();
    if (catA === catB) return a[a.length - 1] - b[b.length - 1];
    if (catA === "Unassigned") return 1;
    if (catB === "Unassigned") return -1;
    return catA.localeCompare(catB, undefined, { numeric: true, sensitivity: "base" });
  });

  if (data.length === 0) {
    ui.alert("No department sheets found in Settings to build a phasing summary from.");
    return;
  }

  const sheet = ss.insertSheet(PHASING_SHEET_NAME);

  // Title row (no merge — see freeze-line constraint comment in earlier history)
  sheet.getRange(1, 1, 1, TOTAL_COLS).setBackground("#cfe2f3");
  sheet.getRange(1, 1).setValue("PHASING SUMMARY")
    .setFontWeight("bold").setFontSize(14);

  // Instructions row
  sheet.getRange(2, 1, 1, TOTAL_COLS).setBackground("#ffffff");
  sheet.getRange(2, 1).setValue(
    "Type each department's allocation in the PHASE columns. " +
    "DEPARTMENTAL AREA and GROSS FLOOR AREA update automatically from the department sheets."
  ).setFontStyle("italic");

  // Header row 4
  const headers = ["DEPARTMENT", "DEPARTMENTAL AREA (SQM)", "GROSS FLOOR AREA (SQM)"];
  for (let p = 0; p < PHASING_NUM_PHASES; p++) headers.push("PHASE " + (p + 1));
  headers.push("REMARKS");
  sheet.getRange(4, 1, 1, TOTAL_COLS).setValues([headers])
    .setFontWeight("bold").setBackground("#d9d9d9")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true);

  let row = 5;
  let dataStart = -1;
  let lastSection = "";

  data.forEach(item => {
    const sectionFull = String(item[5]).trim();
    const sheetName = String(item[6]).trim();
    const displayName = item[7] ? String(item[7]).trim() : sheetName;

    // Section header (bold row above first dept of each section)
    if (sectionFull !== lastSection) {
      sheet.getRange(row, COL_DEPT).setValue(sectionFull);
      sheet.getRange(row, 1, 1, TOTAL_COLS)
        .setFontWeight("bold").setBackground("#f3f3f3");
      row++;
      lastSection = sectionFull;
    }

    if (dataStart === -1) dataStart = row;

    // Department name
    sheet.getRange(row, COL_DEPT).setValue(displayName);

    // Live formulas
    const safe = sheetName.replace(/'/g, "''");

    // DEPARTMENTAL AREA = the "Total Departmental GSM" value on the dept sheet.
    sheet.getRange(row, COL_DEPT_AREA).setFormula(
      `=IFERROR(INDEX('${safe}'!G:G, MATCH("Total Departmental GSM*", '${safe}'!A:A, 0)), 0)`
    );
    // GROSS FLOOR AREA = DEPT AREA × Settings!D2 (falls back to 1.25 when D2 is
    // blank/zero/non-numeric, matching autoUpdateSummarySheet's behavior).
    const deptAreaLetter = _phasingColLetter(COL_DEPT_AREA);
    sheet.getRange(row, COL_GROSS).setFormula(
      `=${deptAreaLetter}${row} * IF(N(Settings!$D$2)>0, Settings!$D$2, 1.25)`
    );

    // Restore preserved phase values + remarks (keyed by department name).
    const pres = preserved[displayName];
    if (pres) {
      for (let p = 0; p < PHASING_NUM_PHASES; p++) {
        const v = pres.phases ? pres.phases[p] : null;
        if (v !== "" && v !== null && v !== undefined) {
          sheet.getRange(row, COL_PHASE_FIRST + p).setValue(v);
        }
      }
      if (pres.remarks !== "" && pres.remarks !== null && pres.remarks !== undefined) {
        sheet.getRange(row, COL_REMARKS).setValue(pres.remarks);
      }
    }

    row++;
  });

  const lastDataRow = row - 1;

  // Grand total row (blank row above for breathing room)
  const grandRow = row + 1;
  sheet.getRange(grandRow, COL_DEPT).setValue("GRAND TOTAL");
  const deptAreaLetter = _phasingColLetter(COL_DEPT_AREA);
  const grossLetter = _phasingColLetter(COL_GROSS);
  sheet.getRange(grandRow, COL_DEPT_AREA).setFormula(
    `=SUM(${deptAreaLetter}${dataStart}:${deptAreaLetter}${lastDataRow})`
  );
  sheet.getRange(grandRow, COL_GROSS).setFormula(
    `=SUM(${grossLetter}${dataStart}:${grossLetter}${lastDataRow})`
  );
  for (let p = 0; p < PHASING_NUM_PHASES; p++) {
    const col = COL_PHASE_FIRST + p;
    const letter = _phasingColLetter(col);
    sheet.getRange(grandRow, col).setFormula(`=SUM(${letter}${dataStart}:${letter}${lastDataRow})`);
  }
  sheet.getRange(grandRow, 1, 1, TOTAL_COLS)
    .setFontWeight("bold").setBackground("#666666").setFontColor("white")
    .setBorder(true, true, true, true, true, true);

  // Number formats
  const dataRowCount = lastDataRow - dataStart + 1;
  if (dataRowCount > 0) {
    sheet.getRange(dataStart, COL_DEPT_AREA, dataRowCount, 2).setNumberFormat("#,##0.00");
    sheet.getRange(dataStart, COL_PHASE_FIRST, dataRowCount, PHASING_NUM_PHASES).setNumberFormat("#,##0.00");
    // Highlight editable phase cells so users see what to fill in.
    sheet.getRange(dataStart, COL_PHASE_FIRST, dataRowCount, PHASING_NUM_PHASES).setBackground("#fff9e6");
  }
  sheet.getRange(grandRow, COL_DEPT_AREA).setNumberFormat("#,##0.00");
  sheet.getRange(grandRow, COL_GROSS).setNumberFormat("#,##0.00");
  for (let p = 0; p < PHASING_NUM_PHASES; p++) {
    sheet.getRange(grandRow, COL_PHASE_FIRST + p).setNumberFormat("#,##0.00");
  }

  // Column widths
  sheet.setColumnWidth(COL_DEPT, 300);
  sheet.setColumnWidth(COL_DEPT_AREA, 170);
  sheet.setColumnWidth(COL_GROSS, 170);
  for (let p = 0; p < PHASING_NUM_PHASES; p++) sheet.setColumnWidth(COL_PHASE_FIRST + p, 110);
  sheet.setColumnWidth(COL_REMARKS, 280);

  // Freeze headers + the Department label column.
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);

  ss.setActiveSheet(sheet);

  const preservedCount = Object.keys(preserved).length;
  ui.alert(
    "Phasing Summary built.\n\n" +
    "Yellow PHASE cells and the REMARKS column are editable.\n" +
    "DEPARTMENTAL AREA and GROSS FLOOR AREA are live formulas — they update " +
    "automatically when any department changes." +
    (preservedCount > 0
      ? "\n\n" + preservedCount + " row(s) of phase values / remarks were preserved."
      : "")
  );
}

/* ===========================================================================
 * INTERNAL HELPERS
 * ========================================================================== */

/**
 * Reads any existing "Phasing Summary" sheet and returns a map of
 * { departmentName -> { phases: [phase1, phase2, ...], remarks: "..." } }
 * for rows that had at least one non-empty phase value or non-empty remark.
 * Department column and phase columns are located by header text so the
 * preservation survives layout tweaks.
 */
function _readPhasingPreserved(ss) {
  const existing = ss.getSheetByName(PHASING_SHEET_NAME);
  const out = {};
  if (!existing) return out;

  const lastRow = existing.getLastRow();
  const lastCol = existing.getLastColumn();
  if (lastRow < 5 || lastCol < 1) return out;

  const headers = existing.getRange(4, 1, 1, lastCol).getValues()[0];
  let deptColIdx = 0;
  let remarksColIdx = -1;
  const phaseColIdxs = [];
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]).trim().toUpperCase();
    if (h === "DEPARTMENT") deptColIdx = i;
    else if (/^PHASE \d+$/.test(h)) phaseColIdxs.push(i);
    else if (h === "REMARKS") remarksColIdx = i;
  }
  if (phaseColIdxs.length === 0 && remarksColIdx === -1) return out;

  const allData = existing.getRange(5, 1, lastRow - 4, lastCol).getValues();
  for (const r of allData) {
    const name = String(r[deptColIdx] || "").trim();
    if (!name || name === "GRAND TOTAL") continue;
    const phases = phaseColIdxs.map(idx => r[idx]);
    const remarks = remarksColIdx >= 0 ? r[remarksColIdx] : "";
    const hasPhases = phases.some(v => v !== "" && v !== 0 && v != null);
    const hasRemarks = remarks !== "" && remarks !== null && remarks !== undefined;
    if (hasPhases || hasRemarks) {
      out[name] = { phases, remarks };
    }
  }
  return out;
}

/**
 * 1-indexed column number -> letter (1 -> A, 26 -> Z, 27 -> AA, ...).
 * Prefixed to avoid colliding with similar helpers in other files.
 */
function _phasingColLetter(num) {
  let letter = "";
  while (num > 0) {
    const rem = (num - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}
