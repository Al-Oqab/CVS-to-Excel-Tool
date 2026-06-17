"""
Convert Excel file(s) to PDF using openpyxl + matplotlib (no Excel or LibreOffice needed).

Usage:
    python excel_to_pdf.py                        # converts all xlsx in ./output
    python excel_to_pdf.py file.xlsx              # converts a single file
    python excel_to_pdf.py -f some_folder         # converts all xlsx in a given folder
"""

import os
import sys
import argparse
import glob

import openpyxl
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


# ── colours ──────────────────────────────────────────────────────────────────

def _hex_to_rgb(hex_str: str) -> tuple[float, float, float]:
    hex_str = hex_str.lstrip("#")
    if len(hex_str) == 8:       # ARGB → RGB
        hex_str = hex_str[2:]
    if len(hex_str) != 6:
        return (1.0, 1.0, 1.0)
    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
    return (r / 255, g / 255, b / 255)


def _cell_bg(cell) -> tuple[float, float, float] | None:
    try:
        fill = cell.fill
        if fill and fill.fill_type not in (None, "none"):
            fg = fill.fgColor
            if fg.type == "rgb" and fg.rgb not in ("00000000", "FFFFFFFF", "FF000000"):
                return _hex_to_rgb(fg.rgb)
    except Exception:
        pass
    return None


def _cell_fg(cell) -> tuple[float, float, float]:
    try:
        color = cell.font.color
        if color and color.type == "rgb":
            return _hex_to_rgb(color.rgb)
    except Exception:
        pass
    return (0.0, 0.0, 0.0)


# ── single sheet → one matplotlib figure ─────────────────────────────────────

def _sheet_to_figure(ws) -> plt.Figure | None:
    rows = list(ws.iter_rows(values_only=False))
    if not rows:
        return None

    # Trim trailing empty rows
    while rows and all(c.value is None for c in rows[-1]):
        rows.pop()
    if not rows:
        return None

    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)

    # Build cell-text / colour matrices
    cell_text = []
    cell_colors = []
    cell_fcolors = []

    for row in rows:
        row_text, row_bg, row_fg = [], [], []
        for i in range(n_cols):
            cell = row[i] if i < len(row) else None
            val = cell.value if cell else None
            row_text.append("" if val is None else str(val))
            bg = _cell_bg(cell) if cell else None
            row_bg.append(bg if bg else (1.0, 1.0, 1.0))
            row_fg.append(_cell_fg(cell) if cell else (0.0, 0.0, 0.0))
        cell_text.append(row_text)
        cell_colors.append(row_bg)
        cell_fcolors.append(row_fg)

    # Figure sizing: ~0.35 in per row, ~1.2 in per col, min 6 in wide
    fig_w = max(8, n_cols * 1.5)
    fig_h = max(3, n_rows * 0.38)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis("off")
    fig.suptitle(ws.title, fontsize=10, fontweight="bold", y=0.98)

    tbl = ax.table(
        cellText=cell_text,
        cellLoc="left",
        loc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(6)
    tbl.auto_set_column_width(list(range(n_cols)))

    # Apply colours
    for (r, c), table_cell in tbl.get_celld().items():
        if 0 <= r < n_rows and 0 <= c < n_cols:
            table_cell.set_facecolor(cell_colors[r][c])
            table_cell.get_text().set_color(cell_fcolors[r][c])
        table_cell.set_edgecolor("#cccccc")
        table_cell.set_linewidth(0.4)

    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        fig.tight_layout(pad=0.5)
    return fig


# ── public API ────────────────────────────────────────────────────────────────

def excel_to_pdf(excel_path: str, pdf_path: str | None = None) -> str:
    excel_path = os.path.abspath(excel_path)
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    if pdf_path is None:
        pdf_path = os.path.splitext(excel_path)[0] + ".pdf"
    pdf_path = os.path.abspath(pdf_path)

    wb = openpyxl.load_workbook(excel_path, data_only=True)

    with PdfPages(pdf_path) as pdf:
        for ws in wb.worksheets:
            fig = _sheet_to_figure(ws)
            if fig is None:
                continue
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

    print(f"  Saved: {pdf_path}")
    return pdf_path


def convert_folder(folder: str) -> list[str]:
    pattern = os.path.join(os.path.abspath(folder), "*.xlsx")
    files = glob.glob(pattern)

    if not files:
        print(f"No .xlsx files found in: {folder}")
        return []

    print(f"Found {len(files)} file(s) in '{folder}':")
    results = []
    for f in files:
        print(f"  Converting: {os.path.basename(f)}")
        try:
            results.append(excel_to_pdf(f))
        except Exception as e:
            print(f"  Error: {e}", file=sys.stderr)

    return results


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert Excel (.xlsx) file(s) to PDF (no Excel needed)."
    )
    parser.add_argument(
        "excel_file", nargs="?", default=None,
        help="Single Excel file to convert (optional)",
    )
    parser.add_argument(
        "-o", "--output", default=None,
        help="Output PDF path (single file mode only)",
    )
    parser.add_argument(
        "-f", "--folder", default=None,
        help="Folder with .xlsx files (default: output)",
    )
    args = parser.parse_args()

    try:
        if args.excel_file:
            excel_to_pdf(args.excel_file, args.output)
        else:
            convert_folder(args.folder or "output")
    except (FileNotFoundError, Exception) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
