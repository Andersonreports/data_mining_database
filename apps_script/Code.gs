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
    const cell = sh.getRange(r, i + 1);
    const links = i === 1 ? phenoLinks(d) : [];
    if (links.length) cell.setRichTextValue(linkedText(String(d.pheno).trim(), links));
    else cell.setValue(i === 2 ? sourceValue(v) : v);
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

// OMIM phenotypes the user picked from the page's suggestions: [{text, mim}].
function phenoLinks(d) {
  return (Array.isArray(d.phenoLinks) ? d.phenoLinks : [])
    .filter(l => l && /^\d{6}$/.test(String(l.mim)) && String(l.text || "").trim())
    .slice(0, 20);
}

// Phenotype text with each picked name linked to its OMIM entry, as curators do by hand.
// Rich text is never read as a formula, so it needs no apostrophe guard.
function linkedText(text, links) {
  text = text.replace(/\r/g, "").slice(0, 2000);
  const b = SpreadsheetApp.newRichTextValue().setText(text);
  links.forEach(l => {
    const t = String(l.text).trim(), s = text.indexOf(t);
    if (s >= 0) b.setLinkUrl(s, s + t.length, "https://www.omim.org/entry/" + l.mim);
  });
  return b.build();
}

/**
 * Papers LitVar2 links to a gene + variant. The page asks through here because
 * LitVar2 doesn't allow requests straight from a browser. Results are cached 6 hours.
 */
const LITVAR = "https://www.ncbi.nlm.nih.gov/research/litvar2-api/variant";
function litvar(gene, variant) {
  gene = String(gene || "").toUpperCase().replace(/\s+/g, "");
  variant = String(variant || "").replace(/\s+/g, "");
  if (!/^[A-Z0-9-]{1,20}$/.test(gene) || !/^[cp]\.\S{1,80}$/i.test(variant)) return { ok: false, error: "Needs a gene and a c. or p. variant." };
  const cache = CacheService.getScriptCache(), key = "lv-" + gene + "-" + variant;
  const kept = cache.get(key);
  if (kept) return JSON.parse(kept);
  const get = u => {
    const res = UrlFetchApp.fetch(u, { muteHttpExceptions: true });
    return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : null;
  };
  const out = { ok: true, query: gene + " " + variant, n: 0, pmids: [] };
  const hit = (get(LITVAR + "/autocomplete/?query=" + encodeURIComponent(gene + " " + variant)) || [])
    .filter(h => (h.gene || []).map(g => String(g).toUpperCase()).indexOf(gene) >= 0)[0];
  if (hit) {
    const pm = ((get(LITVAR + "/get/" + encodeURIComponent(hit._id) + "/publications") || {}).pmids || []).map(String);
    out.n = pm.length; out.pmids = pm.slice(0, 20);
  }
  cache.put(key, JSON.stringify(out), 21600);
  return out;
}

// A plain PMID is stored as a number, as in the rest of the sheet; anything else as text.
function sourceValue(s) { return /^\d+$/.test(s) ? Number(s) : s; }

// Compares cell text the way the page does: whitespace-insensitive at the edges and inside.
function same(v) { return String(v == null ? "" : v).replace(/\s+/g, " ").trim(); }

// ?litvar=1&gene=TP53&variant=c.573del returns LitVar2 paper suggestions.
// Without parameters: opening the /exec URL in a browser checks the deployment is live.
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.litvar) {
    try { return reply(litvar(p.gene, p.variant)); }
    catch (err) { return reply({ ok: false, error: String((err && err.message) || err) }); }
  }
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

// Run this once from the editor (function dropdown → testLitvar → Run) to grant the
// "connect to an external service" permission the LitVar2 lookup needs. Writes nothing.
function testLitvar() {
  Logger.log(JSON.stringify(litvar("TP53", "c.573del")));
}
