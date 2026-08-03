#!/usr/bin/env python3
"""Pull consolidated annual P&L / balance-sheet rows from screener.in for the coverage universe.

Screener.in is a SECONDARY aggregator of Indian stock-exchange filings. It is used here only to
(a) build a consistent multi-year backbone and (b) cross-check figures taken from primary filings.
Every figure that ends up in the workbook is labelled with the source actually relied upon.
"""
import html
import os
import re
import subprocess
import sys

RAW = os.path.join(os.path.dirname(os.path.abspath(__file__)), "raw")
os.makedirs(RAW, exist_ok=True)

TICKERS = [
    "TATASTEEL", "JSWSTEEL", "SAIL", "JINDALSTEL", "JSL",
    "SHYAMMETL", "GPIL", "APLAPOLLO", "MUKANDLTD", "GALLANTT",
    "UTTAMSTL", "KIRLFER", "NSLNISP", "JSWHL",
]

ROWS = [
    "Sales", "Expenses", "Operating Profit", "OPM %", "Other Income", "Interest",
    "Depreciation", "Profit before tax", "Net Profit", "EPS in Rs",
    "Borrowings", "Cash Equivalents", "Total Assets", "Equity Capital", "Reserves",
]


def get(ticker: str, consolidated: bool = True) -> str:
    suffix = "consolidated/" if consolidated else ""
    path = os.path.join(RAW, "scr_%s%s.html" % (ticker, "_c" if consolidated else "_s"))
    if not os.path.exists(path) or os.path.getsize(path) < 20000:
        url = "https://www.screener.in/company/%s/%s" % (ticker, suffix)
        subprocess.run(
            ["curl", "-sSLk", "-m", "90", "-A", "Mozilla/5.0 (X11; Linux x86_64) Chrome/126",
             "-o", path, url], check=False)
    with open(path, encoding="utf-8", errors="ignore") as fh:
        return fh.read()


def flatten(seg: str) -> str:
    seg = re.sub(r"<[^>]+>", "|", seg)
    seg = html.unescape(seg)
    seg = re.sub(r"\|+", "|", seg)
    return re.sub(r"[ \t\r\n]+", " ", seg)


def parse_section(text: str, section: str):
    i = text.find('id="%s"' % section)
    if i < 0:
        return None, {}
    j = text.find("</section>", i)
    seg = flatten(text[i:j if j > 0 else i + 40000])
    cells = [c.strip() for c in seg.split("|")]
    # header periods
    periods = [c for c in cells if re.fullmatch(r"(Mar|Jun|Sep|Dec) \d{4}", c or "")]
    out = {}
    for row in ROWS:
        # find row label then take following numeric cells
        for k, c in enumerate(cells):
            if c == row or c == row + " ":
                vals = []
                for c2 in cells[k + 1:]:
                    if not c2 or c2 in ("+", "%"):
                        continue
                    if re.fullmatch(r"-?[\d,]+\.?\d*%?", c2):
                        vals.append(c2)
                    elif vals:
                        break
                    elif c2 in ROWS:
                        break
                if vals:
                    out[row] = vals
                break
    return periods, out


def main():
    for t in TICKERS:
        text = get(t)
        if len(text) < 20000:
            print("### %s : FETCH FAILED" % t)
            continue
        print("=" * 70)
        print("### %s (consolidated)" % t)
        for section in ("profit-loss", "balance-sheet", "quarters"):
            periods, rows = parse_section(text, section)
            if not periods:
                continue
            print("--- %s\nPERIODS: %s" % (section, " ".join(periods)))
            for r, v in rows.items():
                print("%-20s %s" % (r, " ".join(v)))


if __name__ == "__main__":
    sys.exit(main())
