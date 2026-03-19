"""
Unified Excel Data Extraction Pipeline
=======================================
Auto-discovers all spreadsheet files xl_pdf the script's directory and, for each
workbook, exports every visible non-empty sheet as:

  - CSV file (raw data)
  - Chunked Markdown files (<=100 data rows per chunk, for AI consumption)
  - Named-range Markdown files (one per named range on the sheet)
  - PDF file (visual layout via Excel COM, for AI visual context)

Also produces a _manifest.json per workbook with metadata.

Output structure:
    ./pdfs/{workbook_stem}/
        _manifest.json
        {SanitizedSheet}.pdf                       ← sheet PDF
        NR_{RangeName}_{SanitizedSheet}.pdf        ← named-range PDF
        {SanitizedSheet}/
            {SanitizedSheet}.csv
            {SanitizedSheet}_001.md
            {SanitizedSheet}_002.md
            NR_{RangeName}.md

Supported: .xlsx, .xlsm, .xlsb, .xls, .ods

Requirements:
    - Windows with Excel installed
    - pip install pywin32 openpyxl

Usage:
    python excel_to_pdf.py
"""

import copy
import csv
import json
import math
import os
import re
import shutil
import sys
import time
import traceback
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from gcs_io import GcsIO, is_gcs_uri, parse_gcs_uri

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls", ".ods"}

CHUNK_SIZE = 100  # data rows per markdown chunk (excludes header row)
MIN_FILE_SIZE_KB = 85  # skip files smaller than this (no extract/PDF generated)
MIN_COL_FILL_RATIO = 0.01  # 1% — columns filled below this are phantom/sparse

# Realistic estimate of how many average-width columns fit on one page at
# 100% zoom with zero side margins.  A4 landscape ~11.7" usable,
# portrait ~8.3".  Assuming ~0.78" per column on average.
COLS_PER_PAGE_LANDSCAPE = 15
COLS_PER_PAGE_PORTRAIT = 11


# ---------------------------------------------------------------------------
# Data class
# ---------------------------------------------------------------------------

@dataclass
class SheetResult:
    sheet_name: str
    sanitized_name: str
    sheet_dir: str
    csv_path: str | None = None
    md_paths: list[str] = field(default_factory=list)
    named_range_md_paths: list[str] = field(default_factory=list)
    pdf_path: str | None = None
    named_range_pdf_paths: list[str] = field(default_factory=list)
    row_count: int = 0
    max_row: int = 0        # same as row_count
    max_col: int = 0        # number of columns
    chunk_count: int = 0


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def col_letter_to_number(col_str: str) -> int:
    """Convert Excel column letter(s) to 1-based number. A=1, Z=26, AA=27."""
    result = 0
    for ch in col_str.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def number_to_col_letter(n: int) -> str:
    """Convert 1-based column number to Excel letter(s). 1=A, 26=Z, 27=AA."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def parse_range_string(range_str: str) -> dict | None:
    """Parse a range like '$A$1:$M$50' or 'Sheet!$A$1:$M$50' into components."""
    match = re.search(
        r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", range_str, re.IGNORECASE
    )
    if not match:
        return None
    col_start = col_letter_to_number(match.group(1))
    row_start = int(match.group(2))
    col_end = col_letter_to_number(match.group(3))
    row_end = int(match.group(4))
    return {
        "range": match.group(0),
        "col_start": col_start,
        "row_start": row_start,
        "col_end": col_end,
        "row_end": row_end,
        "num_cols": col_end - col_start + 1,
        "num_rows": row_end - row_start + 1,
    }


def infer_orientation(num_cols: int, num_rows: int) -> str:
    """Infer page orientation from dimensions. ~0.78"/col, ~0.2"/row."""
    est_width = num_cols * 0.78
    est_height = num_rows * 0.2
    return "landscape" if est_width > est_height else "portrait"


def calculate_fit_pages_wide(num_cols: int, orientation: str) -> int:
    """How many pages wide to keep columns readable (not microscopic)."""
    cols_per_page = (
        COLS_PER_PAGE_LANDSCAPE if orientation == "landscape"
        else COLS_PER_PAGE_PORTRAIT
    )
    return max(1, math.ceil(num_cols / cols_per_page))


def sanitize_sheet_name(name: str) -> str:
    """Make a sheet name safe for use as a filename/directory name."""
    sanitized = re.sub(r'[<>:"/\\|?*]', "_", name)
    sanitized = sanitized.replace(" ", "_")
    sanitized = re.sub(r"_+", "_", sanitized).strip("_ ")
    return sanitized or "sheet"


def _upload_dir_to_gcs(local_dir: Path, gcs_uri: str, gcs: GcsIO) -> None:
    g = parse_gcs_uri(gcs_uri)
    for root, _dirs, files in os.walk(local_dir):
        for fname in files:
            file_path = Path(root) / fname
            rel = os.path.relpath(file_path, local_dir).replace("\\", "/")
            obj = f"{g.prefix}/{rel}".strip("/")
            gcs.upload_file(str(file_path), g.bucket, obj)


def _list_gcs_inputs(gcs_uri: str) -> list[tuple[str, int]]:
    gcs = GcsIO()
    g = parse_gcs_uri(gcs_uri)
    results: list[tuple[str, int]] = []
    for name, size in gcs.list_with_sizes(g.bucket, g.prefix):
        ext = os.path.splitext(name)[1].lower()
        if ext in SUPPORTED_EXTENSIONS and not name.endswith("/"):
            results.append((name, size))
    return results


def is_sheet_empty(ws) -> bool:
    """Check if a COM worksheet is truly empty (UsedRange is just A1 and blank)."""
    try:
        used = ws.UsedRange
        if used.Rows.Count == 1 and used.Columns.Count == 1:
            val = used.Cells(1, 1).Value
            return val is None or (isinstance(val, str) and val.strip() == "")
        return False
    except Exception:
        return False


def format_value(val) -> str:
    """Format a cell value as a clean string for CSV/Markdown output."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if isinstance(val, float):
        if val == int(val) and abs(val) < 1e15:
            return str(int(val))
        return f"{val:.10g}"
    # Dates — COM returns pywintypes.datetime (subclass of datetime.datetime)
    if hasattr(val, "strftime") and hasattr(val, "hour"):
        try:
            if val.hour == 0 and val.minute == 0 and val.second == 0:
                return val.strftime("%Y-%m-%d")
            return val.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
    return str(val)


# ---------------------------------------------------------------------------
# Data reading from COM
# ---------------------------------------------------------------------------

def read_sheet_data(ws) -> list[list[str]]:
    """Read all data from a COM worksheet's UsedRange as formatted strings."""
    try:
        used = ws.UsedRange
        data = used.Value
    except Exception:
        return []

    if data is None:
        return []

    # Single cell → scalar
    if not isinstance(data, tuple):
        return [[format_value(data)]]

    # Tuple of tuples (multiple rows) or flat tuple (single row)
    if len(data) > 0 and isinstance(data[0], tuple):
        return [[format_value(v) for v in row] for row in data]

    # Single row
    return [[format_value(v) for v in data]]


def strip_sparse_trailing_columns(
    data: list[list[str]],
    min_fill: float = MIN_COL_FILL_RATIO,
) -> list[list[str]]:
    """Remove trailing columns whose fill ratio is below *min_fill*.

    A pyxlsb artifact can produce 16k phantom columns with just a couple
    of stray zeros.  We scan from the right and drop every column that has
    fewer non-empty cells than ``min_fill * total_rows``, stopping at the
    first column that meets the threshold.
    """
    if not data:
        return data

    total_rows = len(data)
    max_cols = max(len(row) for row in data)
    if max_cols == 0:
        return data

    threshold = max(1, total_rows * min_fill)

    # Find rightmost column that meets the fill threshold
    keep_cols = 0
    for col_idx in range(max_cols - 1, -1, -1):
        filled = sum(
            1 for row in data
            if col_idx < len(row) and row[col_idx] != ""
        )
        if filled >= threshold:
            keep_cols = col_idx + 1
            break

    if keep_cols == 0:
        keep_cols = 1  # always keep at least one column

    if keep_cols < max_cols:
        data = [row[:keep_cols] for row in data]

    return data


def _is_blank_or_zero(val: str) -> bool:
    """Return True if *val* is empty or numerically zero."""
    if val == "":
        return True
    try:
        return float(val) == 0
    except (ValueError, TypeError):
        return False


def strip_trailing_blank_rows(data: list[list[str]]) -> list[list[str]]:
    """Remove trailing rows where every cell is blank or zero."""
    while data and all(_is_blank_or_zero(c) for c in data[-1]):
        data.pop()
    return data


def strip_trailing_blank_cols(data: list[list[str]]) -> list[list[str]]:
    """Remove trailing columns where every cell is blank or zero."""
    if not data:
        return data

    max_cols = max(len(row) for row in data)
    if max_cols == 0:
        return data

    keep_cols = max_cols
    for col_idx in range(max_cols - 1, -1, -1):
        if all(
            _is_blank_or_zero(row[col_idx]) if col_idx < len(row) else True
            for row in data
        ):
            keep_cols = col_idx
        else:
            break

    keep_cols = max(keep_cols, 1)  # always keep at least one column
    if keep_cols < max_cols:
        data = [row[:keep_cols] for row in data]

    return data


def build_print_area(ws, num_rows: int, num_cols: int) -> str:
    """Construct an Excel range string from UsedRange start plus trimmed dims.

    E.g. ``$A$1:$M$42``.
    """
    try:
        start_row = ws.UsedRange.Row
        start_col = ws.UsedRange.Column
    except Exception:
        start_row, start_col = 1, 1

    end_row = start_row + num_rows - 1
    end_col = start_col + num_cols - 1
    col_start_letter = number_to_col_letter(start_col)
    col_end_letter = number_to_col_letter(end_col)
    return f"${col_start_letter}${start_row}:${col_end_letter}${end_row}"


def read_range_data(ws, range_addr: str) -> list[list[str]]:
    """Read data from a specific range address on a worksheet."""
    try:
        rng = ws.Range(range_addr)
        data = rng.Value
    except Exception:
        return []

    if data is None:
        return []

    if not isinstance(data, tuple):
        return [[format_value(data)]]

    if len(data) > 0 and isinstance(data[0], tuple):
        return [[format_value(v) for v in row] for row in data]

    return [[format_value(v) for v in data]]


# ---------------------------------------------------------------------------
# CSV export
# ---------------------------------------------------------------------------

def write_csv(data: list[list[str]], csv_path: str) -> None:
    """Write sheet data as CSV."""
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(data)


# ---------------------------------------------------------------------------
# Markdown export
# ---------------------------------------------------------------------------

def escape_md(val: str) -> str:
    """Escape pipe characters and newlines for Markdown table cells."""
    return val.replace("|", "\\|").replace("\n", " ").replace("\r", "")


def format_md_table(headers: list[str], rows: list[list[str]]) -> str:
    """Format data as a Markdown table."""
    if not headers:
        return ""

    n_cols = len(headers)
    lines = []

    # Header row
    lines.append("| " + " | ".join(escape_md(h) for h in headers) + " |")
    # Separator
    lines.append("|" + "|".join(" --- " for _ in range(n_cols)) + "|")
    # Data rows
    for row in rows:
        padded = list(row) + [""] * max(0, n_cols - len(row))
        lines.append("| " + " | ".join(escape_md(v) for v in padded[:n_cols]) + " |")

    return "\n".join(lines)


def write_chunked_md(
    data: list[list[str]],
    sheet_dir: Path,
    sanitized_name: str,
    sheet_name: str,
) -> list[str]:
    """
    Write sheet data as chunked Markdown files.
    First row is treated as headers and repeated xl_pdf each chunk.
    Returns list of filenames (relative to sheet_dir).
    """
    if not data:
        return []

    headers = data[0]
    body = data[1:]

    total_chunks = max(1, math.ceil(len(body) / CHUNK_SIZE)) if body else 1
    filenames = []

    for chunk_idx in range(total_chunks):
        start = chunk_idx * CHUNK_SIZE
        end = start + CHUNK_SIZE
        chunk_rows = body[start:end]

        chunk_num = chunk_idx + 1
        filename = f"{sanitized_name}_{chunk_num:03d}.md"
        filepath = sheet_dir / filename

        part_label = f"Part {chunk_num} of {total_chunks}"
        if chunk_rows:
            row_label = f"Rows {start + 1}\u2013{start + len(chunk_rows)} of {len(body)} data rows"
        else:
            row_label = "No data rows"

        content_lines = [
            f"# {sheet_name} ({part_label})",
            "",
            f"*{row_label}*",
            "",
            format_md_table(headers, chunk_rows),
            "",
        ]

        with open(filepath, "w", encoding="utf-8") as f:
            f.write("\n".join(content_lines))

        filenames.append(filename)

    return filenames


# ---------------------------------------------------------------------------
# Named-range helpers
# ---------------------------------------------------------------------------

def collect_named_ranges(wb) -> list[dict]:
    """Collect all named ranges from the COM workbook's Names collection."""
    result = []
    try:
        for i in range(1, wb.Names.Count + 1):
            nm = wb.Names.Item(i)
            entry = {
                "name": nm.Name,
                "refers_to": str(nm.RefersTo),
                "visible": bool(nm.Visible),
            }

            full_name = nm.Name
            if "!" in full_name:
                entry["scope_sheet"] = full_name.split("!")[0]
            else:
                entry["scope_sheet"] = None

            ref = str(nm.RefersTo)
            if "#REF!" in ref:
                entry["error"] = True
                entry["range_parsed"] = None
            else:
                ref_clean = ref.lstrip("=")
                if "!" in ref_clean:
                    sheet_part, range_part = ref_clean.rsplit("!", 1)
                    entry["scope_sheet"] = entry["scope_sheet"] or sheet_part.strip("'")
                else:
                    range_part = ref_clean
                entry["range_parsed"] = parse_range_string(range_part)

            result.append(entry)
    except Exception as e:
        print(f"    Warning: could not read named ranges: {e}")
    return result


def get_sheet_named_ranges(all_ranges: list[dict], sheet_name: str) -> list[dict]:
    """Filter named ranges that belong to (or reference) a specific sheet."""
    result = []
    for nr in all_ranges:
        ref = nr.get("refers_to", "")
        scope = nr.get("scope_sheet", "")
        if (scope and scope.strip("'") == sheet_name) or (
            f"{sheet_name}!" in ref or f"'{sheet_name}'!" in ref
        ):
            result.append(nr)
    return result


def write_named_range_md(
    ws,
    named_ranges: list[dict],
    sheet_dir: Path,
    sheet_name: str,
) -> list[str]:
    """
    Write a Markdown file for each named range on this sheet.
    Returns list of filenames (relative to sheet_dir).
    """
    filenames = []

    for nr in named_ranges:
        if nr.get("error"):
            continue

        parsed = nr.get("range_parsed")
        if not parsed:
            continue

        name = nr["name"]
        display_name = name.split("!")[-1] if "!" in name else name
        safe_name = sanitize_sheet_name(display_name)
        filename = f"NR_{safe_name}.md"
        filepath = sheet_dir / filename

        range_addr = parsed["range"]
        data = read_range_data(ws, range_addr)

        content_lines = [
            f"# Named Range: {display_name}",
            "",
            f"- **Sheet:** {sheet_name}",
            f"- **Range:** {nr['refers_to']}",
            f"- **Size:** {parsed['num_rows']} rows \u00d7 {parsed['num_cols']} columns",
            "",
        ]

        if data and len(data) > 1:
            content_lines.append(format_md_table(data[0], data[1:]))
        elif data and len(data) == 1:
            n = len(data[0])
            headers = [number_to_col_letter(parsed["col_start"] + j) for j in range(n)]
            content_lines.append(format_md_table(headers, data))
        else:
            content_lines.append("*No data*")

        content_lines.append("")

        with open(filepath, "w", encoding="utf-8") as f:
            f.write("\n".join(content_lines))

        filenames.append(filename)

    return filenames


# ---------------------------------------------------------------------------
# PDF export
# ---------------------------------------------------------------------------

def determine_title_row(effective_range_addr: str | None) -> str | None:
    """Return a PrintTitleRows string like '$1:$1' from the effective range."""
    parsed = parse_range_string(effective_range_addr) if effective_range_addr else None
    if parsed:
        r = parsed["row_start"]
        return f"${r}:${r}"
    return "$1:$1"


def determine_print_range(ws) -> tuple[str | None, str | None, str | None]:
    """Return (print_area, used_range, effective_range) addresses."""
    try:
        pa_raw = ws.PageSetup.PrintArea
    except Exception:
        pa_raw = ""
    print_area = pa_raw if pa_raw else None

    try:
        used_range = ws.UsedRange.Address
    except Exception:
        used_range = None

    effective = print_area or used_range
    return print_area, used_range, effective


def export_sheet_to_pdf(
    ws,
    excel_app,
    pdf_path: str,
    effective_range: str | None,
    orientation: str,
    num_cols: int,
    sheet_name: str,
) -> float | None:
    """
    Configure PageSetup and export a single sheet to PDF.
    Returns the file size xl_pdf KB, or None on failure.
    """
    ps = ws.PageSetup

    if effective_range:
        ps.PrintArea = effective_range

    # xlLandscape = 2, xlPortrait = 1
    ps.Orientation = 2 if orientation == "landscape" else 1

    pages_wide = calculate_fit_pages_wide(num_cols, orientation)
    ps.Zoom = False
    ps.FitToPagesWide = pages_wide
    ps.FitToPagesTall = False

    ps.LeftMargin = 0
    ps.RightMargin = 0
    ps.TopMargin = excel_app.InchesToPoints(0.3)
    ps.BottomMargin = excel_app.InchesToPoints(0.3)
    ps.HeaderMargin = excel_app.InchesToPoints(0.15)
    ps.FooterMargin = excel_app.InchesToPoints(0.15)

    ps.PrintGridlines = True
    ps.PrintHeadings = False
    ps.CenterHorizontally = False
    ps.CenterVertically = False

    title_row = determine_title_row(effective_range)
    if title_row:
        try:
            ps.PrintTitleRows = title_row
        except Exception:
            pass

    ps.PaperSize = 9  # A4
    ps.LeftHeader = f"&F - {sheet_name}"
    ps.CenterFooter = "Page &P of &N"

    # xlTypePDF = 0, xlQualityStandard = 0
    ws.ExportAsFixedFormat(
        Type=0,
        Filename=pdf_path,
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False,
    )

    if os.path.exists(pdf_path):
        return os.path.getsize(pdf_path) / 1024.0
    return None


def export_named_range_pdfs(
    ws,
    excel_app,
    named_ranges: list[dict],
    wb_dir: Path,
    safe_sheet_name: str,
    sheet_name: str,
) -> list[str]:
    """
    Export each named range on this sheet as a separate PDF.
    PDFs go into wb_dir as NR_{RangeName}_{SheetName}.pdf.
    Returns list of PDF filenames.
    """
    filenames = []

    for nr in named_ranges:
        if nr.get("error"):
            continue

        parsed = nr.get("range_parsed")
        if not parsed:
            continue

        name = nr["name"]
        display_name = name.split("!")[-1] if "!" in name else name
        safe_range_name = sanitize_sheet_name(display_name)

        pdf_filename = f"NR_{safe_range_name}_{safe_sheet_name}.pdf"
        pdf_path = str(wb_dir / pdf_filename)

        range_addr = parsed["range"]
        num_cols = parsed["num_cols"]
        num_rows = parsed["num_rows"]
        orientation = infer_orientation(num_cols, num_rows)

        try:
            export_sheet_to_pdf(
                ws, excel_app, pdf_path, range_addr, orientation,
                num_cols, f"{sheet_name} - {display_name}",
            )
            filenames.append(pdf_filename)
            print(f"      NR PDF: {pdf_filename}")
        except Exception as e:
            print(f"      NR PDF error ({display_name}): {e}")

    return filenames


# ---------------------------------------------------------------------------
# Manifest helpers
# ---------------------------------------------------------------------------

def _sheet_to_manifest(
    sr: SheetResult,
    skipped: bool = False,
    skip_reason: str = "",
) -> dict:
    """Convert a SheetResult to a manifest dict entry."""
    entry = {
        "sheet_name": sr.sheet_name,
        "sanitized_name": sr.sanitized_name,
        "sheet_dir": sr.sheet_dir,
        "row_count": sr.row_count,
        "max_row": sr.max_row,
        "max_col": sr.max_col,
        "chunk_count": sr.chunk_count,
        "csv": sr.csv_path,
        "md_files": sr.md_paths,
        "named_range_files": sr.named_range_md_paths,
        "pdf": sr.pdf_path,
        "named_range_pdfs": sr.named_range_pdf_paths,
    }
    if skipped:
        entry["skipped"] = True
        entry["skip_reason"] = skip_reason
    return entry


# ---------------------------------------------------------------------------
# Process one workbook
# ---------------------------------------------------------------------------

def process_workbook(filepath: Path, output_dir: Path, pdf_mode: str = "both") -> dict:
    """
    Open a workbook via Excel COM, iterate visible sheets, export CSV +
    Markdown + PDF, collect metadata.  Returns manifest dict.

    *pdf_mode*: ``"standard"`` — current full-range PDF only;
    ``"trimmed"`` — trailing-blank-stripped PDF only;
    ``"both"`` (default) — emit both variants per sheet.
    """
    import win32com.client as win32

    input_path = str(filepath.resolve())
    stem = filepath.stem
    wb_dir = output_dir / stem
    wb_dir.mkdir(parents=True, exist_ok=True)

    excel = None
    wb = None

    manifest = {
        "source": filepath.name,
        "extracted_at": datetime.now().isoformat(),
        "workbook_info": {},
        "sheets": [],
    }

    try:
        print("  Starting Excel (headless) ...")
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        print(f"  Opening: {filepath.name}")
        wb = excel.Workbooks.Open(
            input_path,
            ReadOnly=True,
            UpdateLinks=0,
            IgnoreReadOnlyRecommended=True,
        )

        sheet_count = wb.Sheets.Count

        # --- Named ranges (workbook level) ---
        all_named_ranges = collect_named_ranges(wb)

        manifest["workbook_info"] = {
            "sheet_count": sheet_count,
            "named_ranges": [
                {
                    "name": nr["name"],
                    "refers_to": nr["refers_to"],
                    "visible": nr["visible"],
                    "scope_sheet": nr.get("scope_sheet"),
                    "error": nr.get("error", False),
                    "range_parsed": nr.get("range_parsed"),
                }
                for nr in all_named_ranges
            ],
        }

        # --- Track sanitised names for collision avoidance ---
        used_names: dict[str, int] = {}

        for i in range(1, sheet_count + 1):
            ws = wb.Sheets(i)
            sheet_name = ws.Name
            visible = ws.Visible == -1  # xlSheetVisible = -1

            # --- Sanitise name (collision-safe) ---
            safe_name = sanitize_sheet_name(sheet_name)
            if safe_name in used_names:
                used_names[safe_name] += 1
                safe_name = f"{safe_name}_{used_names[safe_name]}"
            else:
                used_names[safe_name] = 1

            sr = SheetResult(
                sheet_name=sheet_name,
                sanitized_name=safe_name,
                sheet_dir=safe_name,
            )

            # --- Skip hidden sheets ---
            if not visible:
                manifest["sheets"].append(
                    _sheet_to_manifest(sr, skipped=True, skip_reason="hidden")
                )
                print(f"    [{i}/{sheet_count}] {sheet_name} -- skipped (hidden)", flush=True)
                continue

            # --- Skip empty sheets ---
            if is_sheet_empty(ws):
                manifest["sheets"].append(
                    _sheet_to_manifest(sr, skipped=True, skip_reason="empty")
                )
                print(f"    [{i}/{sheet_count}] {sheet_name} -- skipped (empty)", flush=True)
                continue

            # --- Read data ---
            data = read_sheet_data(ws)
            raw_cols = max((len(row) for row in data), default=0)
            data = strip_sparse_trailing_columns(data)
            stripped_cols = raw_cols - max((len(row) for row in data), default=0)
            if stripped_cols > 0:
                print(f"      Stripped {stripped_cols} phantom columns "
                      f"(fill < {MIN_COL_FILL_RATIO:.0%})", flush=True)
            sr.row_count = len(data)
            sr.max_row = len(data)
            sr.max_col = max((len(row) for row in data), default=0)

            if sr.row_count == 0:
                manifest["sheets"].append(
                    _sheet_to_manifest(sr, skipped=True, skip_reason="empty")
                )
                print(f"    [{i}/{sheet_count}] {sheet_name} -- skipped (no data)", flush=True)
                continue

            # --- Create sheet subdirectory ---
            sheet_dir = wb_dir / safe_name
            sheet_dir.mkdir(parents=True, exist_ok=True)

            # --- CSV ---
            csv_filename = f"{safe_name}.csv"
            try:
                write_csv(data, str(sheet_dir / csv_filename))
                sr.csv_path = csv_filename
            except Exception as e:
                print(f"      CSV error: {e}")

            # --- Chunked Markdown ---
            try:
                md_files = write_chunked_md(data, sheet_dir, safe_name, sheet_name)
                sr.md_paths = md_files
                sr.chunk_count = len(md_files)
            except Exception as e:
                print(f"      Markdown error: {e}")

            # --- Named-range Markdown ---
            sheet_nrs = get_sheet_named_ranges(all_named_ranges, sheet_name)
            if sheet_nrs:
                try:
                    nr_files = write_named_range_md(ws, sheet_nrs, sheet_dir, sheet_name)
                    sr.named_range_md_paths = nr_files
                except Exception as e:
                    print(f"      Named-range MD error: {e}")

            # --- Sheet PDF (into wb_dir, not sheet_dir) ---
            print_area, used_range, effective = determine_print_range(ws)
            eff_parsed = parse_range_string(effective) if effective else None
            num_cols = eff_parsed["num_cols"] if eff_parsed else sr.max_col
            num_rows = eff_parsed["num_rows"] if eff_parsed else sr.max_row
            orientation = infer_orientation(num_cols, num_rows)
            pages_wide = calculate_fit_pages_wide(num_cols, orientation)

            # -- Standard PDF --
            if pdf_mode in ("standard", "both"):
                pdf_filename = f"{safe_name}.pdf"
                pdf_path = str(wb_dir / pdf_filename)
                try:
                    print(
                        f"    [{i}/{sheet_count}] {sheet_name} "
                        f"-> {pdf_filename}  ({sr.max_col}c x {sr.max_row}r, "
                        f"{sr.chunk_count} chunks, {orientation}, {pages_wide}p wide)",
                        flush=True,
                    )
                    sheet_print_range = print_area  # None if no explicit PrintArea
                    size_kb = export_sheet_to_pdf(
                        ws, excel, pdf_path, sheet_print_range, orientation,
                        num_cols, sheet_name,
                    )
                    sr.pdf_path = pdf_filename
                    print(f"      OK  {size_kb:.0f} KB", flush=True)
                except Exception as e:
                    print(f"      PDF error: {e}", flush=True)
                    traceback.print_exc()

            # -- Trimmed PDF --
            if pdf_mode in ("trimmed", "both"):
                trimmed_data = strip_trailing_blank_rows(copy.deepcopy(data))
                trimmed_data = strip_trailing_blank_cols(trimmed_data)
                trim_rows = len(trimmed_data)
                trim_cols = max((len(r) for r in trimmed_data), default=0)

                if trim_rows != sr.max_row or trim_cols != sr.max_col:
                    trim_orientation = infer_orientation(trim_cols, trim_rows)
                    trim_range = build_print_area(ws, trim_rows, trim_cols)
                    trim_pages_wide = calculate_fit_pages_wide(trim_cols, trim_orientation)
                    trim_pdf_filename = f"{safe_name}_trimmed.pdf"
                    trim_pdf_path = str(wb_dir / trim_pdf_filename)
                    try:
                        print(
                            f"    [{i}/{sheet_count}] {sheet_name} "
                            f"-> {trim_pdf_filename}  "
                            f"(trimmed {sr.max_col}c x {sr.max_row}r "
                            f"-> {trim_cols}c x {trim_rows}r, "
                            f"{trim_orientation}, {trim_pages_wide}p wide)",
                            flush=True,
                        )
                        size_kb = export_sheet_to_pdf(
                            ws, excel, trim_pdf_path, trim_range,
                            trim_orientation, trim_cols, sheet_name,
                        )
                        print(f"      OK  {size_kb:.0f} KB", flush=True)
                    except Exception as e:
                        print(f"      Trimmed PDF error: {e}", flush=True)
                        traceback.print_exc()
                else:
                    print(
                        f"    [{i}/{sheet_count}] {sheet_name} "
                        f"-- no trailing blanks/zeros to trim",
                        flush=True,
                    )

            # If only trimmed mode, still need a standard-style log line and pdf_path
            if pdf_mode == "trimmed":
                pdf_filename = f"{safe_name}_trimmed.pdf"
                trim_pdf_path = str(wb_dir / pdf_filename)
                if os.path.exists(trim_pdf_path):
                    sr.pdf_path = pdf_filename

            # --- Named-range PDFs (into wb_dir) ---
            if sheet_nrs:
                try:
                    nr_pdf_files = export_named_range_pdfs(
                        ws, excel, sheet_nrs, wb_dir, safe_name, sheet_name,
                    )
                    sr.named_range_pdf_paths = nr_pdf_files
                except Exception as e:
                    print(f"      Named-range PDF error: {e}")

            manifest["sheets"].append(_sheet_to_manifest(sr))

    except Exception as e:
        print(f"  ERROR processing {filepath.name}: {e}")
        traceback.print_exc()
        manifest["error"] = str(e)

    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel:
                excel.Quit()
        except Exception:
            pass
        wb = None
        excel = None
        time.sleep(1)

    # --- Write manifest ---
    manifest_path = wb_dir / "_manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    print(f"  Manifest: {manifest_path.name}")

    return manifest


# ---------------------------------------------------------------------------
# Main — auto-discover and orchestrate
# ---------------------------------------------------------------------------

def main():
    script_dir = Path(os.path.abspath(__file__)).parent
    input_root = "./data/xl_pdf"
    output_root = "./data/out"
    input_is_gcs = is_gcs_uri(input_root)
    output_is_gcs = is_gcs_uri(output_root)
    gcs = GcsIO() if (input_is_gcs or output_is_gcs) else None

    # Tee stdout to a log file so we can see progress even when PowerShell buffers
    import io

    class TeeWriter:
        def __init__(self, *streams):
            self.streams = streams
        def write(self, data):
            for s in self.streams:
                s.write(data)
                s.flush()
        def flush(self):
            for s in self.streams:
                s.flush()

    log_file = open(script_dir / "run.log", "w", encoding="utf-8")
    sys.stdout = TeeWriter(sys.__stdout__, log_file)
    sys.stderr = TeeWriter(sys.__stderr__, log_file)

    if output_is_gcs:
        temp_output = tempfile.TemporaryDirectory()
        pdfs_dir = Path(temp_output.name) / "pdfs"
    else:
        temp_output = None
        pdfs_dir = Path(output_root)

    # Clean output directory (retry on Windows lock)
    if output_is_gcs:
        g = parse_gcs_uri(output_root)
        print(f"Deleting gs://{g.bucket}/{g.prefix} ...")
        gcs.delete_prefix(g.bucket, g.prefix)
    else:
        if pdfs_dir.exists():
            print(f"Deleting {pdfs_dir} ...")
            for attempt in range(3):
                try:
                    shutil.rmtree(pdfs_dir)
                    break
                except PermissionError:
                    if attempt < 2:
                        print(f"  Directory locked, retrying xl_pdf 2s ...")
                        time.sleep(2)
                    else:
                        print(f"  Could not delete {pdfs_dir}, clearing contents instead ...")
                        for child in pdfs_dir.rglob("*"):
                            try:
                                if child.is_file():
                                    child.unlink()
                            except Exception:
                                pass
    pdfs_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output directory: {pdfs_dir}\n")

    # Discover spreadsheet files
    if input_is_gcs:
        files = _list_gcs_inputs(input_root)
    else:
        input_dir = Path(input_root)
        files = sorted(
            f for f in input_dir.iterdir()
            if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS
        )

    if not files:
        print(f"No spreadsheet files found xl_pdf {input_root}")
        print(f"Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}")
        sys.exit(0)

    print(f"Found {len(files)} spreadsheet file(s):\n")

    all_results = []
    temp_input = tempfile.TemporaryDirectory() if input_is_gcs else None

    for idx, item in enumerate(files, start=1):
        if input_is_gcs:
            obj_name, size_bytes = item
            file_size_kb = size_bytes / 1024.0
            g = parse_gcs_uri(input_root)
            rel = obj_name[len(g.prefix):].lstrip("/") if g.prefix else obj_name
            local_path = Path(temp_input.name) / rel
            local_path.parent.mkdir(parents=True, exist_ok=True)
            gcs.download_to_file(g.bucket, obj_name, str(local_path))
            filepath = local_path
        else:
            filepath = item
            file_size_kb = filepath.stat().st_size / 1024.0

        if file_size_kb < MIN_FILE_SIZE_KB:
            print(
                f"[{idx}/{len(files)}] {filepath.name} -- skipped "
                f"({file_size_kb:.0f} KB < {MIN_FILE_SIZE_KB} KB)\n"
            )
            continue

        print(f"[{idx}/{len(files)}] {filepath.name}")
        print(f"  {'=' * 50}")

        try:
            result = process_workbook(filepath, pdfs_dir)
        except Exception as e:
            print(f"  FATAL ERROR: {e}\n")
            traceback.print_exc()
            continue

        if output_is_gcs:
            workbook_dir = pdfs_dir / filepath.stem
            gcs_prefix = f"{output_root.rstrip('/')}/{filepath.stem}"
            _upload_dir_to_gcs(workbook_dir, gcs_prefix, gcs)
            shutil.rmtree(workbook_dir, ignore_errors=True)

        all_results.append(result)

        exported = sum(
            1 for s in result.get("sheets", []) if not s.get("skipped")
        )
        total = len(result.get("sheets", []))
        print(f"  Summary: {exported}/{total} sheets exported\n")

    # Final summary
    total_sheets = sum(len(r.get("sheets", [])) for r in all_results)
    total_exported = sum(
        sum(1 for s in r.get("sheets", []) if not s.get("skipped"))
        for r in all_results
    )
    print("=" * 60)
    print(
        f"Done: {len(all_results)} file(s), "
        f"{total_exported}/{total_sheets} sheets exported"
    )
    print(f"Output: {output_root if output_is_gcs else pdfs_dir}")

    if temp_input is not None:
        temp_input.cleanup()
    if temp_output is not None:
        temp_output.cleanup()


if __name__ == "__main__":
    main()
