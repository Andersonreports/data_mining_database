"""Refresh the OMIM and source hyperlinks used by variant_lookup.html.

Google's live CSV feed (which the page reads in the browser) drops the
hyperlinks that are set on cells in the sheet. This script downloads the
sheet once as .xlsx, reads the link on each OMIM phenotype cell (column B)
and on each non-PMID source cell (column C, e.g. DOIs), and writes them
into variant_lookup.html between the LINKS-START / LINKS-END markers.

It only READS the Google Sheet. Run it whenever new rows with links have
been added to the sheet:

    python update_links.py
"""
import io
import json
import re
import sys
import urllib.request
from pathlib import Path

import openpyxl

SHEET_ID = "13eexUpiAMbHRal9C2SV15DcvhmtDzh3QSVD8vO_nnKc"
TAB_NAME = "Gene Papers"  # the tab with gid=964100659
HTML = Path(__file__).with_name("variant_lookup.html")


def norm(s):
    """Must match linkKey() in variant_lookup.html."""
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def main():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
    with urllib.request.urlopen(url, timeout=60) as r:
        wb = openpyxl.load_workbook(io.BytesIO(r.read()))
    if TAB_NAME not in wb.sheetnames:
        sys.exit(f"Tab '{TAB_NAME}' not found. Tabs: {wb.sheetnames}")
    ws = wb[TAB_NAME]

    links = {}
    omim = src = 0
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
        if entry:
            links["|".join([norm(gene), norm(pheno), norm(variant)])] = entry

    html = HTML.read_text(encoding="utf-8")
    new, n = re.subn(r"/\*LINKS-START\*/.*?/\*LINKS-END\*/",
                     lambda m: "/*LINKS-START*/" + json.dumps(links, separators=(",", ":")) + "/*LINKS-END*/",
                     html, flags=re.S)
    if n != 1:
        sys.exit("Could not find the LINKS-START/LINKS-END markers in variant_lookup.html")
    HTML.write_text(new, encoding="utf-8", newline="\n")
    print(f"Updated {HTML.name}: {omim} OMIM links, {src} DOI/URL source links.")


if __name__ == "__main__":
    main()
