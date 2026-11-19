"""Refresh the OMIM and source hyperlinks used by variant_lookup.html.

Google's live CSV feed (which the page reads in the browser) drops the
hyperlinks that are set on cells in the sheet. This script downloads the
sheet once as .xlsx, reads the link on each OMIM phenotype cell (column B)
and on each non-PMID source cell (column C, e.g. DOIs), and writes them
into variant_lookup.html between the LINKS-START / LINKS-END markers.
It also embeds OMIM's gene symbol -> gene MIM number list (mim2gene.txt),
which the page uses for rows that have no OMIM phenotype link in the sheet.

It only READS the Google Sheet. Run it whenever new rows with links have
been added to the sheet:

    python update_links.py
"""
import io
import json
import re
import sys
import urllib.parse
import urllib.request
from pathlib import Path

import openpyxl

SHEET_ID = "13eexUpiAMbHRal9C2SV15DcvhmtDzh3QSVD8vO_nnKc"
TAB_NAME = "Gene Papers"  # the tab with gid=964100659
HTML = Path(__file__).with_name("variant_lookup.html")
MIM2GENE_URL = "https://omim.org/static/omim/data/mim2gene.txt"


def norm(s):
    """Must match linkKey() in variant_lookup.html."""
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


LITVAR = "https://www.ncbi.nlm.nih.gov/research/litvar2-api/variant"


def _get_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def litvar_suggest(gene, variant):
    """PMIDs that LitVar2 links to this gene + variant, or None.

    LitVar2 recognises the different ways a variant is written (c./p./rsID),
    so it finds papers a plain PubMed text search misses. These are only
    suggestions: they are not necessarily the paper the analyst used.
    """
    gene = re.sub(r"\s+", "", str(gene or "")).upper()
    v = next((s.strip() for s in str(variant or "").splitlines() if re.match(r"[cp]\.", s.strip(), re.I)), "")
    if not gene or not v:
        return None
    try:
        hits = _get_json(f"{LITVAR}/autocomplete/?query={urllib.parse.quote(gene + ' ' + v)}")
        hit = next((h for h in hits if gene in [g.upper() for g in h.get("gene", [])]), None)
        if not hit:
            return None
        pubs = _get_json(f"{LITVAR}/get/{urllib.parse.quote(hit['_id'], safe='')}/publications")
    except Exception as e:  # network problems shouldn't stop the link update
        print(f"  LitVar2 lookup failed for {gene} {v}: {e}")
        return None
    pmids = [str(p) for p in pubs.get("pmids", [])]
    return {"n": len(pmids), "p": pmids[:20], "q": f"{gene} {v}"} if pmids else None


def main():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
    with urllib.request.urlopen(url, timeout=60) as r:
        wb = openpyxl.load_workbook(io.BytesIO(r.read()))
    if TAB_NAME not in wb.sheetnames:
        sys.exit(f"Tab '{TAB_NAME}' not found. Tabs: {wb.sheetnames}")
    ws = wb[TAB_NAME]

    links = {}
    omim = src = suggested = 0
    for row in ws.iter_rows(min_row=2):
        gene, pheno, pmid, variant = (row[i].value if i < len(row) else None for i in range(4))
        if not (gene or pheno or pmid):
            continue
        entry = {}
        if row[1].hyperlink and row[1].hyperlink.target:
            entry["o"] = row[1].hyperlink.target
            omim += 1
        # PMIDs are linked by the page itself; keep links only for DOI/URL sources
        if row[2].hyperlink and row[2].hyperlink.target and not re.fullmatch(r"\d+(\.0)?", str(pmid or "").strip()):
            entry["s"] = row[2].hyperlink.target
            src += 1
        # No source at all in the sheet: ask LitVar2 for papers on this gene + variant
        if not str(pmid or "").strip() and not entry.get("s"):
            sug = litvar_suggest(gene, variant)
            if sug:
                entry["lv"] = sug
                suggested += 1
        if entry:
            links["|".join([norm(gene), norm(pheno), norm(variant)])] = entry

    # Gene symbol -> gene MIM number, from OMIM's public mim2gene.txt. Used by the
    # page when the sheet has no OMIM phenotype link for a row.
    req = urllib.request.Request(MIM2GENE_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        text = r.read().decode("utf-8")
    pairs = []
    for line in text.splitlines():
        if line.startswith("#"):
            continue
        cols = line.split("\t")
        if len(cols) >= 4 and "gene" in cols[1] and cols[3].strip():
            pairs.append(f"{cols[3].strip().upper()}:{cols[0].strip()}")
    genemim = ",".join(sorted(pairs))

    html = HTML.read_text(encoding="utf-8")
    for start, end, value, label in [
        ("LINKS-START", "LINKS-END", json.dumps(links, separators=(",", ":")), "links"),
        ("GENEMIM-START", "GENEMIM-END", json.dumps(genemim), "gene MIM numbers"),
    ]:
        html, n = re.subn(rf"/\*{start}\*/.*?/\*{end}\*/",
                          lambda m: f"/*{start}*/{value}/*{end}*/", html, flags=re.S)
        if n != 1:
            sys.exit(f"Could not find the {start}/{end} markers ({label}) in variant_lookup.html")
    HTML.write_text(html, encoding="utf-8", newline="\n")
    print(f"Updated {HTML.name}: {omim} OMIM links, {src} DOI/URL source links, "
          f"{len(pairs)} gene MIM numbers, {suggested} LitVar2 paper suggestions for rows without a source.")


if __name__ == "__main__":
    main()
