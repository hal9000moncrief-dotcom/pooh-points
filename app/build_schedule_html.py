import os
import html
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles.colors import Color

REPO_ROOT = os.path.join(os.path.dirname(__file__), "..")
DOCS_DIR = os.path.join(REPO_ROOT, "docs")

XLSX_PATH = os.path.join(DOCS_DIR, "Schedule 2026.xlsx")
OUT_HTML = os.path.join(DOCS_DIR, "Schedule.html")


def _rgb_from_color(c: Color):
    """
    Return CSS hex color like #RRGGBB or None.
    Handles ARGB (e.g. 'FF112233') and rgb (e.g. '112233').
    Ignores theme/indexed unless explicitly rgb is present.
    """
    if c is None:
        return None
    rgb = getattr(c, "rgb", None)
    if not rgb:
        return None
    rgb = str(rgb).strip()
    # ARGB -> drop alpha
    if len(rgb) == 8:
        rgb = rgb[2:]
    if len(rgb) != 6:
        return None
    return f"#{rgb}"


def _cell_style_to_css(cell):
    styles = []

    # Font
    f = cell.font
    if f is not None:
        if f.bold:
            styles.append("font-weight:700")
        if f.italic:
            styles.append("font-style:italic")
        if f.underline:
            styles.append("text-decoration:underline")
        if f.color is not None:
            col = _rgb_from_color(f.color)
            if col:
                styles.append(f"color:{col}")

    # Fill (background)
    fill = cell.fill
    if fill is not None and getattr(fill, "patternType", None) == "solid":
        fg = getattr(fill, "fgColor", None)
        col = _rgb_from_color(fg)
        if col:
            styles.append(f"background:{col}")

    # Alignment
    a = cell.alignment
    if a is not None:
        if a.horizontal:
            styles.append(f"text-align:{a.horizontal}")
        if a.vertical:
            styles.append(f"vertical-align:{a.vertical}")
        if a.wrap_text:
            styles.append("white-space:pre-wrap")

    return ";".join(styles)


def _escape_cell_value(v):
    if v is None:
        return "&nbsp;"
    # preserve newlines
    s = str(v)
    s = html.escape(s)
    s = s.replace("\n", "<br>")
    if s.strip() == "":
        return "&nbsp;"
    return s


def main():
    if not os.path.isfile(XLSX_PATH):
        raise SystemExit(f"ERROR: Missing file: {XLSX_PATH}")

    wb = load_workbook(XLSX_PATH, data_only=True)
    ws = wb.active

    # Determine the used range robustly
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    # Build merged-cell maps: top-left -> (rowspan, colspan)
    merged_top_left = {}  # (r,c) -> (rs, cs)
    merged_covered = set()  # cells inside merges that are not top-left
    for m in ws.merged_cells.ranges:
        min_row = m.min_row
        min_col = m.min_col
        max_row_m = m.max_row
        max_col_m = m.max_col
        rs = max_row_m - min_row + 1
        cs = max_col_m - min_col + 1
        merged_top_left[(min_row, min_col)] = (rs, cs)
        for r in range(min_row, max_row_m + 1):
            for c in range(min_col, max_col_m + 1):
                if (r, c) != (min_row, min_col):
                    merged_covered.add((r, c))

    updated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write("<!doctype html><html><head><meta charset='utf-8'>")
        f.write("<title>Schedule</title>")
        f.write(
            "<style>"
            "body{font-family:Calibri,Arial;background:#ffffff}"
            ".wrap{width:1200px;margin:20px auto;border:3px solid #000;background:#FFFFCC;padding:12px}"
            "h1{margin:0 0 10px 0;text-align:center;background:#C0C0C0;border:1px solid #000;padding:10px}"
            "table{border-collapse:collapse;width:100%;background:#ffffff}"
            "td,th{border:1px solid #000;padding:6px 8px;font-size:12pt}"
            "</style>"
        )
        f.write("</head><body><div class='wrap'>")
        f.write("<h1>Schedule</h1>")
        f.write(f"<div style='font-size:10pt;margin-bottom:10px;'><b>Last updated:</b> {html.escape(updated)}</div>")
        f.write("<table>")

        for r in range(1, max_row + 1):
            f.write("<tr>")
            for c in range(1, max_col + 1):
                if (r, c) in merged_covered:
                    continue

                cell = ws.cell(row=r, column=c)

                tag = "th" if r <= 2 else "td"  # your first 2 rows are header rows
                attrs = []
                if (r, c) in merged_top_left:
                    rs, cs = merged_top_left[(r, c)]
                    if rs > 1:
                        attrs.append(f"rowspan='{rs}'")
                    if cs > 1:
                        attrs.append(f"colspan='{cs}'")

                css = _cell_style_to_css(cell)
                if css:
                    attrs.append(f"style='{css}'")

                val = _escape_cell_value(cell.value)

                f.write(f"<{tag} {' '.join(attrs)}>{val}</{tag}>")

            f.write("</tr>")

        f.write("</table>")

        f.write("<div style='margin-top:10px;font-size:10pt;'>")
        f.write(f"<a href='{html.escape(os.path.basename(XLSX_PATH))}'>Download the Excel version</a>")
        f.write("</div>")

        f.write("</div></body></html>")

    print(f"Wrote: {OUT_HTML}")


if __name__ == "__main__":
    main()
