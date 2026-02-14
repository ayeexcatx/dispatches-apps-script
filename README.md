# Dispatch System: Safe Change Guide (Beginner-Friendly)

This project is a Google Apps Script dispatch pipeline. If you are changing behavior, **make small edits in the right function** and avoid touching shared naming/format rules unless you truly need to.

## Before you edit anything

1. Find the feature area below.
2. Change only the function listed for that area.
3. Keep file naming, placeholders, and label text stable unless the change requires it.
4. If you do change naming/formatting, verify downstream features still work (archive sorting, web page grouping, link labels).

---

## 1) Form submit handling

### Exact function responsible
- `onFormSubmit(e)` in `dispatch.js`

### What it modifies
- Reads the submitted form row from `e.values`.
- Splits selected truck numbers and runs processing once per truck.
- Triggers downstream work for each truck (archive copy creation, company page refresh).

### What **not** to change unless necessary
- The **index positions** in `e.values` (these must match the exact Google Form question order).
- Truck parsing logic (`split(',').map(trim)`) unless form output format changed.
- The call path that eventually updates company pages (`updateCompanyDispatchPage(...)`).

### Where formatting decisions are made
- Date file-name format (`yyyy-MM-dd`) is set inside `onFormSubmit`.
- Sortable time (`HHMM`) for archive names is built by `convertToSortableTime(...)` (nested helper).
- Human-readable times (like `6:15 AM`) are set by `formatTime(...)` (nested helper).

---

## 2) Archive dispatch creation

### Exact function responsible
- `onFormSubmit(e)` in `dispatch.js` (archive-copy section inside it)
- Helper: `boldLabel(body, labelText)`

### What it modifies
- Copies the archive template doc (`template.makeCopy(...)`).
- Names archive files using date/time/truck/job pattern.
- Saves copy into truck-specific archive folder (or fallback folder).
- Replaces placeholders and applies styling in archive copy, then saves and closes.

### What **not** to change unless necessary
- Archive file naming pattern (`YYYY-MM-DD_HHMM_Dispatch_TRUCK_JOB`) because page parsing/sorting depends on it.
- Truck archive folder mappings (`truckArchiveFolders`) and fallback folder usage.
- Placeholder token names (for example `{{START LOCATION}}`) unless template and replacement code are updated together.

### Where formatting decisions are made
- Archive file name composition is inside `onFormSubmit`.
- Archive text replacement for `MCRC` / `CMPM` is in archive-specific regex replacement blocks.
- Archive label styling is applied with `boldLabel(...)` and phrase styling (`setBold`, `setUnderline`).
- Optional section display logic is in the `if (field.trim()) ... else remove paragraph` blocks.

## 3) Company web page rebuilding

### Exact function responsible
- `updateCompanyDispatchPage(companyName)` in `dispatch.js`
- Batch trigger helper: `updateAllCompanyPages()` in `dispatch.js`

### What it modifies
- Reads dispatch files from each company’s truck folders in Drive.
- Filters to recent files (currently last 18 days).
- Builds grouped HTML sections: Upcoming, Today, Past.
- Creates/updates `*_dispatch_list.html` in company folder and sets link sharing.

### What **not** to change unless necessary
- Expected dispatch file-name parse format (date/time extraction uses regex).
- Company-to-truck mapping (`folderMap`) and company folder naming assumptions.
- `*_dispatch_list.html` naming convention (used by web app lookup).

### Where formatting decisions are made
- Date/time display labels are created in `toLocaleDateString(...)` and `toLocaleTimeString(...)`.
- Status badges (`CANCELLED`, `AMENDMENT`) and link label text are built in label HTML logic.
- Visual styling (colors, borders, typography) is defined in the inline CSS inside the HTML template string.

---

## 4) Web app serving (read-only delivery path)

### Exact function responsible
- `doGet(e)` in `webApp.js`
- Helper: `findFileRecursively(folder, targetFileName)`

### What it modifies
- Does not create dispatch data.
- Reads the matching company HTML file from Drive and returns it as web output.

### What **not** to change unless necessary
- URL parameter contract (`?company=...`) and resulting file-name pattern `${company}_dispatch_list.html`.
- Recursive lookup behavior unless folder layout changes.

### Where formatting decisions are made
- Almost none here; this path mainly serves already-generated HTML.
- Browser page title is set in `doGet(e)` via `.setTitle(...)`.

---

## 5) Maintenance cleanup (safety for storage growth)

### Exact function responsible
- `cleanupOldDispatches()` in `maintenance.js`
- `deleteOldReplacedDispatches()` in `maintenance.js`

### What it modifies
- Moves old archive files (and old files in "Replaced Dispatches") to Drive trash.

### What **not** to change unless necessary
- Date parsing assumptions from file names.
- Age threshold logic unless policy changed.
- Skip rule for `.html` files in archive cleanup.

### Where formatting decisions are made
- Not a formatting area; mostly age cutoff and folder traversal logic.

---

## Safe-change checklist (quick practical steps)

- If adding a form field:
  - Update `e.values` index mapping in `onFormSubmit(e)`.
  - Add placeholder replacement in the archive section.
  - Add optional removal logic if field can be blank.
- If changing label text:
  - Update text in the template doc and any related `boldLabel(...)` calls.
- If changing file naming:
  - Update both creation and parsing logic in page rebuild functions.
- If adding a truck/company:
  - Update archive folder IDs, truck-to-company mapping, and company folder map.

When unsure, prefer **minimal edits in one feature block** and keep naming/format conventions stable.
