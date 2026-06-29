/**
 * USER MANUAL — IN-SHEET REFERENCE
 *
 * showUserManual() builds (or rebuilds) a "User Manual" sheet in the active
 * spreadsheet with a styled, plain-English explanation of every menu item and
 * the main workflow. Intended to be called from the Report Automation menu so
 * non-technical users always have a built-in reference.
 *
 * The same content lives in USER_MANUAL.md at the repo root for developer
 * reference. Keep the two in sync when editing.
 */

const USER_MANUAL_SHEET = "User Manual";

/**
 * Manual content as a flat array of lines. Each entry has a `type` that
 * determines styling, plus content fields. Keep entries short so the on-sheet
 * cells stay readable without horizontal scrolling.
 *
 * Types:
 *   title       — big banner row
 *   subtitle    — small italic line under the title
 *   heading     — section header (gray bg)
 *   para        — plain wrapped paragraph
 *   numbered    — auto-numbered step inside the current "How it works" block
 *   item        — bold name + " — " + description on one wrapped line
 *   tip         — bullet line, lighter color
 *   spacer      — blank row
 */
const USER_MANUAL_LINES = [
  { type: "title",    text: "REPORT AUTOMATION — USER MANUAL" },
  { type: "subtitle", text: "A quick guide to every button and how the pieces fit together." },
  { type: "spacer" },

  { type: "heading",  text: "HOW THE SYSTEM WORKS" },
  { type: "numbered", text: "Standard Room Size holds the default Length / Width / Remarks for every room name." },
  { type: "numbered", text: "Fill out a Department Template (name in A5, category in H1, rooms below) and click Generate Department — a new sheet is created and registered in Settings." },
  { type: "numbered", text: "Settings controls the order, category, and display name of each department in the Summary." },
  { type: "numbered", text: "Create Summary builds the Summary + cover + appendix PDFs and emails you the combined report." },
  { type: "spacer" },

  { type: "heading",  text: "MENU: REPORT AUTOMATION" },
  { type: "item", name: "Create Summary",
    desc: "Rebuilds the Summary from Settings, generates appendix PDFs for every department, and emails the combined PDF to you. Use when you're ready to send the report." },
  { type: "item", name: "Build Appendix",
    desc: "Generates a Google Docs appendix with images from the Drive folder that match the rooms in your departments. Opens the new doc at the end." },
  { type: "item", name: "Refresh Sheet Names",
    desc: "Updates Settings column A with the current list of all sheets in the file. Run it if a newly added or renamed sheet isn't showing up." },
  { type: "item", name: "Refresh Rooms from Master List",
    desc: "Walks every generated department sheet and re-pulls Length / Width / Remarks from Standard Room Size for any matching room. Custom remarks marked with ||| are preserved." },
  { type: "item", name: "Wrap Remarks (All Sheets)",
    desc: "Forces text wrap on the Remarks column for every generated department sheet so long remarks don't get clipped." },
  { type: "item", name: "Show Total GSM",
    desc: "Opens a sidebar showing the total Departmental GSM across the file with a per-sheet breakdown." },
  { type: "item", name: "Setup Auto-Refresh",
    desc: "Installs the trigger that keeps the Summary in sync when sheets are renamed or deleted. Run this once per file." },
  { type: "item", name: "Grant Permissions",
    desc: "Re-runs the permission flow if the script reports missing access." },
  { type: "item", name: "Reset File (Clean Start)",
    desc: "Deletes all generated department sheets plus Summary, Phasing Summary, and Appendix. Clears Non-Standard Report data and Settings cols G & H. Templates and the legend in Settings col B are kept. Then refreshes the sheet names." },
  { type: "item", name: "User Manual",
    desc: "Rebuilds this sheet for quick reference." },
  { type: "spacer" },

  { type: "heading",  text: "MENU: PHASING" },
  { type: "item", name: "Build Phasing Summary",
    desc: "Creates a Phasing Summary sheet listing every department with live area / gross floor area columns plus editable Phase 1, Phase 2, and Remarks columns. Existing Phase values and Remarks are preserved when you rebuild." },
  { type: "spacer" },

  { type: "heading",  text: "ON-SHEET BUTTONS / BEHAVIORS" },
  { type: "item", name: "Department Template → Generate Department",
    desc: "Reads A5 (department name), H1 (category), and the rooms below, then creates a new department sheet and registers it in Settings. Tabs are color-coded by the category's font color in the Settings legend." },
  { type: "item", name: "Editing a generated department sheet",
    desc: "Just edit it. The Summary auto-updates. Changing A5 also renames the tab. Custom dimensions or remarks get marked with ||| so a later master-list refresh doesn't overwrite them." },
  { type: "item", name: "Non-Standard Report",
    desc: "Audit log. The script appends a row here every time you change a room's dimensions away from the Standard Room Size defaults." },
  { type: "spacer" },

  { type: "heading",  text: "IMPORTANT SHEETS AND WHAT THEY DO" },
  { type: "item", name: "Settings",
    desc: "Configuration. C1 = project name, C3 = project header, C4 = date, D2 = gross multiplier. Col B (rows 5+) is the category legend with colors. Cols G & H (rows 5+) is the per-department category + sheet reference list the Summary reads from." },
  { type: "item", name: "Cover Template",
    desc: "Page 1 of the Summary PDF. Project header in I10 (version) and I15 (date). 34 rows total." },
  { type: "item", name: "Summary Template",
    desc: "Pages 2+ of the Summary PDF. Hospital name in F33. 34 rows total." },
  { type: "item", name: "Department Template",
    desc: "The blank you fill in to generate a new department." },
  { type: "item", name: "Standard Room Size",
    desc: "Master list. Editing here is how you change defaults for every future department." },
  { type: "spacer" },

  { type: "heading",  text: "CATEGORY LEGEND (SETTINGS COLS B + C)" },
  { type: "item", name: "Col B — Category name + tab color",
    desc: "Each unique category you write in col B is a legend entry. The font color you give that cell becomes the tab color of every department assigned to this category (via col G). Black or default font = no tab color. Change the font color in col B and all matching tabs recolor next time a Summary refresh runs." },
  { type: "item", name: "Col C — Sub Total checkbox",
    desc: "Tick the checkbox next to a category in col B and that category gets a Sub Total row at the bottom of its block in the Summary (blank row → Sub Total → blank row → next category). Untick it and the Sub Total disappears." },
  { type: "spacer" },

  { type: "heading",  text: "TIPS" },
  { type: "tip", text: "Want a fresh project with the same templates? → Reset File (Clean Start)." },
  { type: "tip", text: "Renamed a generated sheet and Settings looks stale? → Refresh Sheet Names, then Create Summary." },
  { type: "tip", text: "Updated Standard Room Size after generating departments? → Refresh Rooms from Master List picks up the new defaults everywhere (custom ||| rows stay)." },
  { type: "tip", text: "Don't see the final report? → It's emailed to your active Google account after Create Summary finishes." }
];

/**
 * Build / rebuild the User Manual sheet from USER_MANUAL_LINES.
 * Wired into the Report Automation menu in onOpen (Code.js).
 */
function showUserManual() {
  const ss = SpreadsheetApp.getActive();

  // Always rebuild so edits to USER_MANUAL_LINES are picked up next run.
  let sheet = ss.getSheetByName(USER_MANUAL_SHEET);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(USER_MANUAL_SHEET);

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 920);

  // Reset numbered-list counter at every "heading" so each section's list
  // starts at 1 again.
  let stepCounter = 0;
  let row = 1;

  USER_MANUAL_LINES.forEach(line => {
    if (line.type === "heading") stepCounter = 0;
    if (line.type === "numbered") stepCounter++;

    const cell = sheet.getRange(row, 1);

    switch (line.type) {
      case "title":
        cell.setValue(line.text)
          .setFontSize(20).setFontWeight("bold")
          .setBackground("#1a73e8").setFontColor("white")
          .setHorizontalAlignment("center").setVerticalAlignment("middle");
        sheet.setRowHeight(row, 50);
        break;

      case "subtitle":
        cell.setValue(line.text)
          .setFontSize(11).setFontStyle("italic").setFontColor("#5f6368")
          .setHorizontalAlignment("center");
        sheet.setRowHeight(row, 24);
        break;

      case "heading":
        cell.setValue(line.text)
          .setFontSize(13).setFontWeight("bold").setFontColor("#1a73e8")
          .setBackground("#e8f0fe")
          .setVerticalAlignment("middle");
        sheet.setRowHeight(row, 30);
        break;

      case "numbered": {
        const text = stepCounter + ". " + line.text;
        cell.setValue(text).setFontSize(11).setWrap(true).setVerticalAlignment("top");
        break;
      }

      case "item": {
        // Bold the item name; rest of the line is regular weight.
        const fullText = line.name + " — " + line.desc;
        const bold = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(11).build();
        const regular = SpreadsheetApp.newTextStyle().setBold(false).setFontSize(11).build();
        const rich = SpreadsheetApp.newRichTextValue()
          .setText(fullText)
          .setTextStyle(0, line.name.length, bold)
          .setTextStyle(line.name.length, fullText.length, regular)
          .build();
        cell.setRichTextValue(rich).setWrap(true).setVerticalAlignment("top");
        break;
      }

      case "para":
        cell.setValue(line.text).setFontSize(11).setWrap(true).setVerticalAlignment("top");
        break;

      case "tip":
        cell.setValue("• " + line.text)
          .setFontSize(11).setFontColor("#5f6368")
          .setWrap(true).setVerticalAlignment("top");
        break;

      case "spacer":
        sheet.setRowHeight(row, 8);
        break;
    }
    row++;
  });

  sheet.setFrozenRows(1);
  ss.setActiveSheet(sheet);
}
