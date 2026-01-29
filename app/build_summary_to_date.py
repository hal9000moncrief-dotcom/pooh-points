import os
import re
import sys
from collections import defaultdict
from bs4 import BeautifulSoup

DOCS_DIR = os.path.join(os.path.dirname(__file__), "..", "docs")

def parse_cap_pd(argv) -> int | None:
    if len(argv) < 2:
        return None
    s = argv[1].strip().upper()
    m = re.fullmatch(r"PD(\d+)", s)
    if not m:
        raise SystemExit("Usage: python app/build_summary_to_date.py [PD7]")
    return int(m.group(1))

def pd_num_from_filename(fn: str) -> int | None:
    m = re.search(r"Final_Owners_PD(\d+)\.html$", fn)
    return int(m.group(1)) if m else None

def read_owner_totals_from_final_owners_html(path: str) -> dict[str, int]:
    """
    Reads docs/Final_Owners_PDx.html which contains a table like:
      Owner | Starter Pooh Total | Starters Count So Far
    Returns {owner: starter_pooh_total}.
    """
    with open(path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    table = soup.find("table")
    if not table:
        return {}

    rows = table.find_all("tr")
    out: dict[str, int] = {}

    # skip header
    for tr in rows[1:]:
        tds = tr.find_all("td")
        if len(tds) < 2:
            continue
        owner = tds[0].get_text(strip=True)
        total_txt = tds[1].get_text(strip=True)
        try:
            out[owner] = int(total_txt)
        except:
            out[owner] = 0

    return out

def main():
    cap_pd = parse_cap_pd(sys.argv)

    # Collect Final_Owners_PD*.html in docs
    files = []
    for fn in os.listdir(DOCS_DIR):
        n = pd_num_from_filename(fn)
        if n is None:
            continue
        if cap_pd is not None and n > cap_pd:
            continue
        files.append((n, fn))

    files.sort(key=lambda x: x[0])  # PD1..PDN

    # per_owner_per_pd[owner][pd] = starter_total_that_pd
    per_owner_per_pd: dict[str, dict[int, int]] = defaultdict(dict)
    owners_set = set()

    for pd, fn in files:
        path = os.path.join(DOCS_DIR, fn)
        totals = read_owner_totals_from_final_owners_html(path)

        # Important: Final_Owners_PDx.html is already "starter total for that day"
        # (not cumulative), per your generation logic.
        for owner, v in totals.items():
            owners_set.add(owner)
            per_owner_per_pd[owner][pd] = int(v)

    owners = sorted(list(owners_set))

    # Build totals and sort owners by total desc
    owner_total = {o: sum(per_owner_per_pd[o].get(pd, 0) for pd, _ in files) for o in owners}
    owners_sorted = sorted(owners, key=lambda o: owner_total.get(o, 0), reverse=True)

    # Output
    out_path = os.path.join(DOCS_DIR, "SummaryToDate.html")
    title = "Summary To Date"
    if cap_pd is not None:
        title += f" (through PD{cap_pd})"

    with open(out_path, "w", encoding="utf-8") as out:
        out.write("<!doctype html><html><head><meta charset='utf-8'>")
        out.write(f"<title>{title}</title>")
        out.write(
            "<style>"
            "body{font-family:Arial}"
            "table{border-collapse:collapse;font-size:14px}"
            "th,td{border:1px solid #ccc;padding:4px 6px}"
            "th{background:#eee}"
            "</style>"
        )
        out.write("</head><body>")
        out.write(f"<h2>{title}</h2>")

        # Header: Owner, Total, PD1..PDN
        out.write("<table><thead><tr>")
        out.write("<th>Owner</th><th>Total</th>")
        for pd, _fn in files:
            out.write(f"<th>PD{pd}</th>")
        out.write("</tr></thead><tbody>")

        for owner in owners_sorted:
            out.write("<tr>")
            out.write(f"<td>{owner}</td>")
            out.write(f"<td>{owner_total.get(owner, 0)}</td>")
            for pd, _fn in files:
                out.write(f"<td>{per_owner_per_pd[owner].get(pd, 0)}</td>")
            out.write("</tr>")

        out.write("</tbody></table></body></html>")

    print(f"Wrote: {out_path}")

if __name__ == "__main__":
    main()
