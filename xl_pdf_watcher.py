"""
xl_pdf Watcher Service
======================
Continuously monitors an output directory for ``**/xl_pdf/.extract`` flag
files.  When one is found the Excel file xl_pdf that directory is opened via
COM, every visible non-empty sheet is exported as a PDF, and the
``.extract`` flag is written so the upstream lana-converter pipeline
knows the conversion is complete.

Usage:
    python xl_pdf_watcher.py <output_dir> [--interval <seconds>]

Requirements:
    - Windows with Excel installed
    - pip install pywin32
"""

import argparse
import copy
import glob
import os
import shutil
import sys
import time
import traceback
import tempfile
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Imports from the existing extraction pipeline
# ---------------------------------------------------------------------------
from gcs_io import GcsIO, format_gcs_uri, is_gcs_uri, parse_gcs_uri
from excel_to_pdf import (
    SUPPORTED_EXTENSIONS,
    sanitize_sheet_name,
    is_sheet_empty,
    determine_print_range,
    export_sheet_to_pdf,
    infer_orientation,
    calculate_fit_pages_wide,
    parse_range_string,
    strip_sparse_trailing_columns,
    strip_trailing_blank_rows,
    strip_trailing_blank_cols,
    build_print_area,
    read_sheet_data,
)


def log(msg: str = "...") -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def find_excel_file(xl_pdf_dir: Path | str, gcs: GcsIO | None = None) -> Path | str | None:
    """Return the first spreadsheet file found xl_pdf *xl_pdf_dir*, or None."""
    if isinstance(xl_pdf_dir, str) and is_gcs_uri(xl_pdf_dir):
        if gcs is None:
            gcs = GcsIO()
        g = parse_gcs_uri(xl_pdf_dir)
        for obj in gcs.list(g.bucket, g.prefix):
            ext = os.path.splitext(obj)[1].lower()
            if ext in SUPPORTED_EXTENSIONS:
                return obj
        return None

    for child in Path(xl_pdf_dir).iterdir():
        if child.is_file() and child.suffix.lower() in SUPPORTED_EXTENSIONS:
            return child
    return None


def process_xl_pdf_dir(
    xl_pdf_dir: Path | str,
    pdf_mode: str = "both",
    test: bool = False,
    gcs: GcsIO | None = None,
) -> None:
    """Convert the Excel file xl_pdf *xl_pdf_dir* to per-sheet PDFs.

    *pdf_mode*: ``"standard"`` — current full-range PDF only;
    ``"trimmed"`` — trailing-blank-stripped PDF only;
    ``"both"`` (default) — emit both variants per sheet.

    *test*: when True, skip all PDF processing — just write ``.extract``
    and remove ``.extract`` (useful for pipeline testing without Excel).

    On success the Excel file is deleted, ``.extract`` is written, and
    ``.extract`` is removed.  On failure the flags are left untouched so the
    upstream pipeline can detect the timeout.
    """
    is_gcs = isinstance(xl_pdf_dir, str) and is_gcs_uri(xl_pdf_dir)
    if is_gcs and gcs is None:
        gcs = GcsIO()

    if is_gcs:
        g = parse_gcs_uri(xl_pdf_dir)
        extract_flag_obj = f"{g.prefix}/.extract".strip("/")
    else:
        xl_pdf_dir = Path(xl_pdf_dir)
        extract_flag = xl_pdf_dir / ".extract"

    if test:
        log(f"  TEST MODE: skipping PDF export for {xl_pdf_dir}")
        script_dir = Path(os.path.abspath(__file__)).parent
        test_pdf_dir = script_dir / "pdfs" / "Copy of 47.1.4 02_ProjectPEC_Financial Model_v0.617"
        if is_gcs:
            excel_obj = find_excel_file(xl_pdf_dir, gcs=gcs)
            if test_pdf_dir.is_dir():
                copied = 0
                for pdf_file in test_pdf_dir.glob("*.pdf"):
                    gcs.upload_file(
                        str(pdf_file), g.bucket, f"{g.prefix}/{pdf_file.name}".strip("/")
                    )
                    copied += 1
                log(f"  Uploaded {copied} test PDF(s) to {xl_pdf_dir}")
            else:
                log(f"  WARNING: test PDF folder not found: {test_pdf_dir}")
            if excel_obj:
                try:
                    gcs.delete(g.bucket, excel_obj)
                except Exception as e:
                    log(f"  WARNING: could not delete {os.path.basename(excel_obj)}: {e}")
            gcs.write_text(g.bucket, f"{g.prefix}/.extract".strip("/"), datetime.now().isoformat())
            try:
                gcs.delete(g.bucket, extract_flag_obj)
            except Exception as e:
                log(f"  WARNING: could not delete .extract: {e}")
        else:
            excel_file = find_excel_file(xl_pdf_dir)
            if test_pdf_dir.is_dir():
                copied = 0
                for pdf_file in test_pdf_dir.glob("*.pdf"):
                    shutil.copy2(pdf_file, xl_pdf_dir / pdf_file.name)
                    copied += 1
                log(f"  Copied {copied} test PDF(s) to {xl_pdf_dir}")
            else:
                log(f"  WARNING: test PDF folder not found: {test_pdf_dir}")
            if excel_file:
                try:
                    excel_file.unlink()
                except Exception as e:
                    log(f"  WARNING: could not delete {excel_file.name}: {e}")
            processed_flag = xl_pdf_dir / ".extract"
            processed_flag.write_text(datetime.now().isoformat(), encoding="utf-8")
            try:
                extract_flag.unlink()
            except Exception as e:
                log(f"  WARNING: could not delete .extract: {e}")
        log(f"  Done (test): {xl_pdf_dir}")
        return

    temp_dir = None
    import win32com.client as win32

    if is_gcs:
        excel_obj = find_excel_file(xl_pdf_dir, gcs=gcs)
        if excel_obj is None:
            log(f"  WARNING: no spreadsheet xl_pdf {xl_pdf_dir}, skipping")
            return
        temp_dir = tempfile.TemporaryDirectory()
        xl_pdf_dir_local = Path(temp_dir.name)
        excel_file = xl_pdf_dir_local / os.path.basename(excel_obj)
        gcs.download_to_file(g.bucket, excel_obj, str(excel_file))
        log(f"  Processing: {excel_file.name} xl_pdf {xl_pdf_dir}")
    else:
        xl_pdf_dir_local = xl_pdf_dir
        excel_file = find_excel_file(xl_pdf_dir_local)
        if excel_file is None:
            log(f"  WARNING: no spreadsheet xl_pdf {xl_pdf_dir_local}, skipping")
            return
        log(f"  Processing: {excel_file.name} xl_pdf {xl_pdf_dir_local}")

    excel = None
    wb = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        wb = excel.Workbooks.Open(
            str(excel_file.resolve()),
            ReadOnly=True,
            UpdateLinks=0,
            IgnoreReadOnlyRecommended=True,
        )

        sheet_count = wb.Sheets.Count
        used_names: dict[str, int] = {}
        exported = 0

        for i in range(1, sheet_count + 1):
            ws = wb.Sheets(i)
            sheet_name = ws.Name
            visible = ws.Visible == -1  # xlSheetVisible

            if not visible:
                log(f"    [{i}/{sheet_count}] {sheet_name} -- skipped (hidden)")
                continue

            if is_sheet_empty(ws):
                log(f"    [{i}/{sheet_count}] {sheet_name} -- skipped (empty)")
                continue

            # Read data to detect truly-empty sheets and get dimensions
            data = read_sheet_data(ws)
            data = strip_sparse_trailing_columns(data)
            if not data or all(all(c == "" for c in row) for row in data):
                log(f"    [{i}/{sheet_count}] {sheet_name} -- skipped (no data)")
                continue

            max_col = max((len(row) for row in data), default=0)
            max_row = len(data)

            # Collision-safe sanitised name
            safe_name = sanitize_sheet_name(sheet_name)
            if safe_name in used_names:
                used_names[safe_name] += 1
                safe_name = f"{safe_name}_{used_names[safe_name]}"
            else:
                used_names[safe_name] = 1

            # Orientation and fit-to-pages
            print_area, used_range, effective = determine_print_range(ws)
            eff_parsed = parse_range_string(effective) if effective else None
            num_cols = eff_parsed["num_cols"] if eff_parsed else max_col
            num_rows = eff_parsed["num_rows"] if eff_parsed else max_row
            orientation = infer_orientation(num_cols, num_rows)
            pages_wide = calculate_fit_pages_wide(num_cols, orientation)

            sheet_print_range = print_area  # None if no explicit PrintArea

            # -- Standard PDF --
            if pdf_mode in ("standard", "both"):
                pdf_filename = f"{safe_name}.pdf"
                pdf_path = str(xl_pdf_dir_local / pdf_filename)

                log(
                    f"    [{i}/{sheet_count}] {sheet_name} -> {pdf_filename}  "
                    f"({num_cols}c x {num_rows}r, {orientation}, {pages_wide}p wide)"
                )

                size_kb = export_sheet_to_pdf(
                    ws, excel, pdf_path, sheet_print_range,
                    orientation, num_cols, sheet_name,
                )
                if size_kb is not None:
                    log(f"      OK  {size_kb:.0f} KB")
                    exported += 1
                else:
                    log(f"      WARNING: PDF not created")

            # -- Trimmed PDF --
            if pdf_mode in ("trimmed", "both"):
                trimmed_data = strip_trailing_blank_rows(copy.deepcopy(data))
                trimmed_data = strip_trailing_blank_cols(trimmed_data)
                trim_rows = len(trimmed_data)
                trim_cols = max((len(r) for r in trimmed_data), default=0)

                if trim_rows != max_row or trim_cols != max_col:
                    trim_orientation = infer_orientation(trim_cols, trim_rows)
                    trim_range = build_print_area(ws, trim_rows, trim_cols)
                    trim_pages_wide = calculate_fit_pages_wide(trim_cols, trim_orientation)
                    trim_pdf_filename = f"{safe_name}_trimmed.pdf"
                    trim_pdf_path = str(xl_pdf_dir_local / trim_pdf_filename)

                    log(
                        f"    [{i}/{sheet_count}] {sheet_name} -> {trim_pdf_filename}  "
                        f"(trimmed {max_col}c x {max_row}r -> {trim_cols}c x {trim_rows}r, "
                        f"{trim_orientation}, {trim_pages_wide}p wide)"
                    )

                    size_kb = export_sheet_to_pdf(
                        ws, excel, trim_pdf_path, trim_range,
                        trim_orientation, trim_cols, sheet_name,
                    )
                    if size_kb is not None:
                        log(f"      OK  {size_kb:.0f} KB")
                        exported += 1
                    else:
                        log(f"      WARNING: trimmed PDF not created")
                else:
                    log(
                        f"    [{i}/{sheet_count}] {sheet_name} "
                        f"-- no trailing blanks/zeros to trim"
                    )

        log(f"  Exported {exported}/{sheet_count} sheets")

    except Exception as e:
        log(f"  ERROR converting {excel_file.name}: {e}")
        traceback.print_exc()
        # Leave .extract xl_pdf place — do NOT write .extract
        if temp_dir is not None:
            temp_dir.cleanup()
            temp_dir = None
        return

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

    # --- Success: clean up and signal completion ---
    if is_gcs:
        for pdf_file in xl_pdf_dir_local.glob("*.pdf"):
            gcs.upload_file(
                str(pdf_file), g.bucket, f"{g.prefix}/{pdf_file.name}".strip("/")
            )
        try:
            gcs.delete(g.bucket, excel_obj)
        except Exception as e:
            log(f"  WARNING: could not delete {os.path.basename(excel_obj)}: {e}")
        try:
            gcs.write_text(
                g.bucket, f"{g.prefix}/.extract".strip("/"), datetime.now().isoformat()
            )
        except Exception as e:
            log(f"  WARNING: could not write .extract: {e}")
        try:
            gcs.delete(g.bucket, extract_flag_obj)
        except Exception as e:
            log(f"  WARNING: could not delete .extract: {e}")
    else:
        try:
            excel_file.unlink()
        except Exception as e:
            log(f"  WARNING: could not delete {excel_file.name}: {e}")
        processed_flag = xl_pdf_dir / ".extract"
        processed_flag.write_text(datetime.now().isoformat(), encoding="utf-8")
        try:
            extract_flag.unlink()
        except Exception as e:
            log(f"  WARNING: could not delete .extract: {e}")

    if temp_dir is not None:
        temp_dir.cleanup()
        temp_dir = None

    log(f"  Done: {xl_pdf_dir}")


def scan_and_process(output_dir: Path | str, pdf_mode: str = "both", test: bool = False) -> int:
    """Scan *output_dir* for ``.extract`` flags and process each one.

    Returns the number of directories processed.
    """
    if isinstance(output_dir, str) and is_gcs_uri(output_dir):
        gcs = GcsIO()
        g = parse_gcs_uri(output_dir)
        count = 0
        for obj in sorted(gcs.list(g.bucket, g.prefix)):
            if not obj.endswith("/xl_pdf/.extract"):
                continue
            xl_pdf_prefix = obj.rsplit("/", 1)[0]
            xl_pdf_dir = format_gcs_uri(g.bucket, xl_pdf_prefix)
            log(f"Found .extract xl_pdf {xl_pdf_dir}")
            process_xl_pdf_dir(xl_pdf_dir, pdf_mode=pdf_mode, test=test, gcs=gcs)
            count += 1
        return count

    output_dir = Path(output_dir)
    count = 0
    for extract_flag in sorted(output_dir.rglob("xl_pdf/.extract")):
        xl_pdf_dir = extract_flag.parent
        log(f"Found .extract xl_pdf {xl_pdf_dir}")
        process_xl_pdf_dir(xl_pdf_dir, pdf_mode=pdf_mode, test=test)
        count += 1
    return count


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Watch for xl_pdf .extract flags and convert Excel to PDFs",
    )
    parser.add_argument(
        "output_dir",
        type=str,
        help="Root output directory containing {stem}/xl_pdf/ subdirectories",
    )
    parser.add_argument(
        "--interval",
        type=float,
        default=5.0,
        help="Polling interval xl_pdf seconds (default: 5)",
    )
    parser.add_argument(
        "--pdf-mode",
        choices=["standard", "trimmed", "both"],
        default="both",
        dest="pdf_mode",
        help="PDF export mode: standard (full range), trimmed (strip trailing blanks/zeros), both (default)",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        default=False,
        help="Test mode: skip PDF processing, just write .extract and delete .extract + Excel file",
    )
    args = parser.parse_args()

    output_dir = args.output_dir
    if not is_gcs_uri(output_dir):
        output_dir = Path(output_dir).resolve()
        if not output_dir.is_dir():
            print(f"ERROR: {output_dir} is not a directory", file=sys.stderr)
            sys.exit(1)

    log(f"xl_pdf watcher started")
    log(f"  Watching: {output_dir}")
    log(f"  Interval: {args.interval}s")
    log(f"  PDF mode: {args.pdf_mode}")
    if args.test:
        log(f"  TEST MODE: PDF processing disabled")
    log(f"  Press Ctrl+C to stop")

    try:
        i = 0
        while True:
            count = scan_and_process(output_dir, pdf_mode=args.pdf_mode, test=args.test)
            if count:
                log(f"Processed {count} xl_pdf dir(s). Resuming watch ...")
            else:
                log()
            time.sleep(args.interval)
            i+=1
            if i==3:
                break

    except KeyboardInterrupt:
        log("Stopped.")


if __name__ == "__main__":
    main()
