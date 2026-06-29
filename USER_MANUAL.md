# Report Automation — User Manual

A quick guide to every button and how the system fits together.

## How the system works

1. **Standard Room Size** holds the default Length / Width / Remarks for every room name.
2. Fill out a **Department Template** (department name in `A5`, category in `H1`, room list below) and click the **Generate Department** button on that sheet — the script creates a new sheet for that department and registers it in **Settings**.
3. **Settings** controls the order, category, and display name of each department in the final Summary.
4. **Create Summary** builds the Summary sheet, the cover page, the appendix PDFs for every department, and emails you the combined PDF.

## Menu: Report Automation

- **Create Summary** — Rebuilds the Summary from Settings, generates appendix PDFs for every department listed there, and emails the combined PDF to you. Use this when you're ready to send the report.
- **Build Appendix** — Generates a Google Docs appendix with images from the Drive folder that match the rooms in your departments. Opens the new document at the end.
- **Refresh Sheet Names** — Updates Settings column A with the current list of all sheets in the file. Run it if a newly added/renamed sheet isn't showing up.
- **Refresh Rooms from Master List** — Walks every generated department sheet and re-pulls Length / Width / Remarks from Standard Room Size for any room that matches by name. Custom remarks marked with `|||` are preserved.
- **Wrap Remarks (All Sheets)** — Forces text wrap on the Remarks column for every generated department sheet so long remarks don't get clipped.
- **Show Total GSM** — Opens a sidebar showing the total Departmental GSM across the file with a per-sheet breakdown.
- **Setup Auto-Refresh** — Installs the trigger that keeps the Summary in sync when sheets are renamed or deleted. Run once per file.
- **Grant Permissions** — Re-runs the permission flow if the script reports missing access.
- **Reset File (Clean Start)** — Deletes all generated department sheets, the Summary, the Phasing Summary, and the Appendix. Clears Non-Standard Report data and Settings cols G & H. Templates and the legend in Settings col B are kept. Then refreshes the sheet names list.
- **User Manual** — Rebuilds this sheet inside the spreadsheet for quick reference.

## Menu: Phasing

- **Build Phasing Summary** — Creates a Phasing Summary sheet listing every department with live area / gross floor area columns plus editable Phase 1, Phase 2, and Remarks columns. Existing Phase values and Remarks are preserved when you rebuild.

## On-sheet buttons / behaviors

- **Department Template → Generate Department** — Reads A5 (department name), H1 (category), and the rooms below, then creates a new department sheet and registers it in Settings. Tabs are color-coded by the category's font color in the Settings legend (col B).
- **Editing a generated department sheet** — Just edit it; the Summary auto-updates. Changing the department name in A5 also renames the tab. Custom dimensions or remarks get marked with `|||` so a later master-list refresh doesn't overwrite them.
- **Non-Standard Report** — Audit log; the script appends a row here every time you change a room's dimensions away from the Standard Room Size defaults.

## Important sheets and what they're for

- **Settings** — Configuration. Row 1 has the project name (`C1`), row 3 the project header (`C3`), row 4 the date (`C4`), `D2` the gross multiplier. Col B (rows 5+) is the category legend with colors. Cols G & H (rows 5+) is the per-department category + sheet reference list that the Summary reads from.
- **Cover Template** — Page 1 of the Summary PDF. Project header lives in `I10` (version) and `I15` (date). 34 rows total.
- **Summary Template** — Pages 2+ of the Summary PDF. Hospital name in `F33`. 34 rows total.
- **Department Template** — The blank you fill in to generate a new department.
- **Standard Room Size** — Master list. Editing this is how you change the defaults for every future department.

## Category legend (Settings cols B + C)

- **Col B — Category name + tab color** — Each unique category you write in col B is a legend entry. The **font color** you give that cell becomes the **tab color** of every department assigned to this category (via col G). Black or default font color = no tab color. Change the font color in col B and all matching tabs recolor next time a Summary refresh runs.
- **Col C — Sub Total checkbox** — Tick the checkbox next to a category in col B and that category gets a **Sub Total row** at the bottom of its block in the Summary (blank row → `Sub Total` → blank row → next category). Untick it and the Sub Total disappears.

## Tips

- Want a fresh project with the same templates? → **Reset File (Clean Start)**.
- Renamed a generated sheet and Settings looks stale? → **Refresh Sheet Names**, then **Create Summary**.
- Updated Standard Room Size after generating departments? → **Refresh Rooms from Master List** picks up the new defaults everywhere (custom `|||` rows stay).
- Don't see the final report? → It's emailed to your active Google account after **Create Summary** finishes.
