/**
 * @OnlyCurrentDoc  Limits the script's permission to this one spreadsheet file.
 */

/**
 * Anderson Internal Database: "Add entry" / "Edit" bridge.
 *
 * Receives a new entry from the lookup page (variant_lookup.html) and writes it
 * as a new row at the bottom of the curation tab, or edits an existing row
 * (action: "edit"). Only requests carrying the team passcode are accepted, if a
 * passcode is set (optional).
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

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = SpreadsheetApp.getActive().getSheets().filter(s => s.getSheetId() === SHEET_GID)[0];
      if (!sh) return reply({ ok: false, error: "Couldn't find the curation tab in the sheet." });

      if (d.action === "edit") {
        const res = editRow(sh, d);
        if (res.ok && !res.unchanged) cache.put(hourKey, String(used + 1), 3700);
        return reply(res);
      }

      // PMID/DOI is optional: an entry can start as gene + variant and be completed later.
      const gene = clean(d.gene).toUpperCase().replace(/\s+/g, "");
      const source = clean(d.pmid);
      const variant = clean(d.variant) || "Gene paper";
      if (!/^[A-Z0-9][A-Z0-9-]{0,19}$/.test(gene)) return reply({ ok: false, error: "The gene symbol doesn't look right." });
      const row = [gene, clean(d.pheno), sourceValue(source), variant, clean(d.remarks)];

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

/**
 * Edits one existing row. The page sends the row number, the values it showed the
 * user (d.orig) and the new values. The row is changed only if it still holds exactly
 * what the user saw, so a sorted sheet or someone else's edit is never overwritten.
 * Only cells whose value changed are written, so hyperlinks on untouched cells stay.
 */
function editRow(sh, d) {
  const r = Number(d.row);
  if (!(r >= 2 && r <= sh.getLastRow())) return { ok: false, conflict: true, error: "That row no longer exists. Reload the page and try again." };

  const range = sh.getRange(r, 1, 1, COLS);
  const now = range.getDisplayValues()[0];
  const orig = d.orig || {};
  const seen = [orig.gene, orig.pheno, orig.pmid, orig.variant, orig.remarks];
  if (now.some((v, i) => same(v) !== same(seen[i]))) {
    return { ok: false, conflict: true, error: "This row was changed in the sheet since the page loaded (it may have been sorted or edited). Reload the page and try again." };
  }

  let gene = clean(d.gene).replace(/\s+/g, " ");
  if (same(gene) !== same(now[0])) {                   // only a changed gene name is checked
    if (!/\s/.test(gene)) gene = gene.toUpperCase();   // keep labels like "VWF type 3" as written
    if (!/^[A-Za-z0-9][A-Za-z0-9 .\-]{0,29}$/.test(gene)) return { ok: false, error: "The gene symbol doesn't look right." };
  }
  // The variant is written as sent (the page keeps an empty or free-text cell as it was).
  const next = [gene, clean(d.pheno), clean(d.pmid), clean(d.variant), clean(d.remarks)];

  const changed = [];
  next.forEach((v, i) => {
    if (same(v) === same(now[i])) return;
    sh.getRange(r, i + 1).setValue(i === 2 ? sourceValue(v) : v);
    changed.push(HEADERS[i]);
  });
  if (!changed.length) return { ok: true, row: r, unchanged: true };

  const cell = sh.getRange(r, 1), old = cell.getNote();
  const line = "Edited via Anderson Internal Database (" + changed.join(", ") + ")\n" +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm");
  cell.setNote(old ? old + "\n\n" + line : line);
  return { ok: true, row: r, changed: changed };
}

const HEADERS = ["gene", "OMIM phenotype", "PMID", "variant", "remarks"];

// A plain PMID is stored as a number, as in the rest of the sheet; anything else as text.
function sourceValue(s) { return /^\d+$/.test(s) ? Number(s) : s; }

// Compares cell text the way the page does: whitespace-insensitive at the edges and inside.
function same(v) { return String(v == null ? "" : v).replace(/\s+/g, " ").trim(); }

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
