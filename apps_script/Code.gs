/**
 * @OnlyCurrentDoc  Limits the script's permission to this one spreadsheet file.
 */

/**
 * Anderson Internal Database: "Add entry" bridge.
 *
 * Receives a new entry from the lookup page (variant_lookup.html) and writes it
 * as a new row at the bottom of the curation tab. Only requests carrying the
 * team passcode are accepted, if a passcode is set (optional).
 *
 * Setup (once):
 *   1. Open the Google Sheet → Extensions → Apps Script. Paste this file in as Code.gs.
 *   2. Optional: Project Settings (gear) → Script properties → Add property
 *        PASSCODE = <the team passcode>   (without it, anyone with the page can add rows)
 *   3. Deploy → New deployment → type "Web app"
 *        Execute as: Me      Who has access: Anyone
 *      Copy the Web app URL (ends in /exec) and put it in ADD_URL in variant_lookup.html.
 *
 * To change the passcode later, edit the script property; no redeploy needed.
 * After editing this code, use Deploy → Manage deployments → Edit → New version
 * so the same URL keeps working.
 */

const SHEET_GID = 964100659;      // the curation tab the page reads
const COLS = 5;                   // A Gene | B OMIM Phenotype | C PMID | D Variant | E Remarks
const MAX_PER_HOUR = 60;          // safety cap on new rows per hour
const WEB_FILL = "#FDEBD3";       // light orange fill marking rows added from the page

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    // Passcode is optional: only checked when a PASSCODE script property exists.
    const pass = PropertiesService.getScriptProperties().getProperty("PASSCODE");
    if (pass && String(d.passcode || "") !== pass) {
      return reply({ ok: false, needPass: true, error: d.passcode ? "Wrong passcode." : "This sheet needs the team passcode." });
    }

    // Safety cap so a stray bot can't flood the sheet.
    const cache = CacheService.getScriptCache();
    const hourKey = "adds-" + Math.floor(Date.now() / 3600000);
    const used = Number(cache.get(hourKey) || 0);
    if (used >= MAX_PER_HOUR) return reply({ ok: false, error: "Too many entries this hour. Please try again later." });

    const gene = clean(d.gene).toUpperCase().replace(/\s+/g, "");
    const source = clean(d.pmid);
    const variant = clean(d.variant) || "Gene paper";
    if (!/^[A-Z0-9][A-Z0-9-]{0,19}$/.test(gene)) return reply({ ok: false, error: "The gene symbol doesn't look right." });
    if (!source) return reply({ ok: false, error: "A PMID or DOI is required." });

    const row = [gene, clean(d.pheno), /^\d+$/.test(source) ? Number(source) : source, variant, clean(d.remarks)];

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = SpreadsheetApp.getActive().getSheets().filter(s => s.getSheetId() === SHEET_GID)[0];
      if (!sh) return reply({ ok: false, error: "Couldn't find the curation tab in the sheet." });

      // Next empty row after the last filled cell in column A. Other columns are not used
      // to decide this, so side lists elsewhere in the tab are never shifted or overwritten.
      const colA = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
      let last = colA.length;
      while (last > 1 && String(colA[last - 1][0]).trim() === "") last--;
      const r = last + 1;
      if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);

      // Mark rows that came from the page: light orange fill + a note on the gene cell.
      // Neither changes the cell values, so the lookup page reads the row as usual.
      sh.getRange(r, 1, 1, COLS).setValues([row]).setBackground(WEB_FILL);
      sh.getRange(r, 1).setNote("Added via Anderson Internal Database\n" +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm"));
      cache.put(hourKey, String(used + 1), 3700);
      return reply({ ok: true, row: r });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return reply({ ok: false, error: String((err && err.message) || err) });
  }
}

// Lets you open the /exec URL in a browser to check the deployment is live.
function doGet() {
  return reply({ ok: true, service: "Anderson Internal Database add-entry" });
}

// Trims, caps length, and stops text being read as a formula (=, +, -, @).
function clean(v) {
  let s = String(v == null ? "" : v).replace(/\r/g, "").trim().slice(0, 2000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
