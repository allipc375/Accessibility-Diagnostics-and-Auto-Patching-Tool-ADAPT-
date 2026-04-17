# main.py

import sys
import os

from docx_checker import check_docx
from pptx_checker import run_pptx_accessibility_check
from pdf_checker import check_pdf

SUPPORTED_EXTENSIONS = {".docx", ".pptx", ".pdf"}


# ============================================================
# REPORTING
# ============================================================

def print_report(result, title):
    print(f"\n{title}")
    print("-" * 50)
    print(f"Score: {result['score']} ({result['band']})")
    print("Details:")
    for k, v in result["details"].items():
        print(f"  - {k}: {v}")
    if result["issues"]:
        print("\nIssues:")
        for i, (problem, location) in enumerate(result["issues"], 1):
            print(f"{i}. {problem} ({location})")
    else:
        print("\nNo issues found.")


def print_suite_summary(rows):
    """
    Print a compact before/after table for every file processed in suite mode.

    Each row dict contains:
        file, before_score, before_band, after_score, after_band,
        fixed_path, error
    """
    print("\n" + "=" * 78)
    print(" SUITE SUMMARY")
    print("=" * 78)

    col_file = max((len(r["file"]) for r in rows), default=4)
    col_file = max(col_file, 4)

    header = (
        f"{'File':<{col_file}}  "
        f"{'Before':>7}  {'Band':<12}  "
        f"{'After':>7}  {'Band':<12}  "
        f"Fixed Path"
    )
    print(header)
    print("-" * len(header))

    for r in rows:
        if r.get("error"):
            print(f"{r['file']:<{col_file}}  ERROR: {r['error']}")
            continue

        after_score = r["after_score"] if r["after_score"] is not None else "-"
        after_band  = r["after_band"]  if r["after_band"]  is not None else "-"
        fixed       = r["fixed_path"]  if r["fixed_path"]  is not None else "(not fixed)"

        print(
            f"{r['file']:<{col_file}}  "
            f"{r['before_score']:>7}  {r['before_band']:<12}  "
            f"{str(after_score):>7}  {str(after_band):<12}  "
            f"{fixed}"
        )

    print("=" * 78)


# ============================================================
# CHECKER DISPATCH
# ============================================================

def run_checker(file_path, fix=False):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        result = check_docx(file_path, apply_fix=fix)
        result.setdefault("fixed_path", None)
        return result
    elif ext == ".pptx":
        return run_pptx_accessibility_check(file_path, fix=fix)
    elif ext == ".pdf":
        result = check_pdf(file_path, apply_fix=fix)
        result.setdefault("fixed_path", None)
        return result
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def _fixed_path_for(file_path):
    """Derive the expected _fixed path from an original file path."""
    base, ext = os.path.splitext(file_path)
    return f"{base}_fixed{ext}"


# ============================================================
# SINGLE-FILE MODE  (original behaviour, completely unchanged)
# ============================================================

def run_single(file_path, fix):
    # -------- BEFORE FIX --------
    before = run_checker(file_path, fix=False)
    print_report(before, "Before Fix")

    # -------- APPLY FIX --------
    if fix:
        print("\n--- Applying Fixes ---")
        fix_result = run_checker(file_path, fix=True)

        fixed_path = fix_result.get("fixed_path") or _fixed_path_for(file_path)

        # -------- AFTER FIX --------
        after = run_checker(fixed_path, fix=False)
        print_report(after, "After Fix")


# ============================================================
# SUITE MODE
# ============================================================

def collect_files(targets):
    """
    Expand *targets* (a list of paths) into an ordered flat list of supported
    files.

    - Directories are scanned one level deep (not recursive) in sorted order.
    - Individual file paths are accepted as-is if their extension is supported.
    - Unsupported or missing paths emit a warning and are skipped.
    """
    files = []
    for target in targets:
        if os.path.isdir(target):
            for name in sorted(os.listdir(target)):
                if os.path.splitext(name)[1].lower() in SUPPORTED_EXTENSIONS:
                    files.append(os.path.join(target, name))
        elif os.path.isfile(target):
            if os.path.splitext(target)[1].lower() in SUPPORTED_EXTENSIONS:
                files.append(target)
            else:
                print(f"[suite] Skipping unsupported file: {target}")
        else:
            print(f"[suite] Path not found, skipping: {target}")
    return files


def run_suite(targets, fix):
    """
    Run accessibility checks (and optionally fixes) on every collected file.

    For each file:
      1. Print the full "Before Fix" report (identical format to single-file
         mode).
      2. If --fix: apply all fixes, save the _fixed document alongside the
         original, then print the full "After Fix" report.
      3. Accumulate one summary row for the final table that is printed after
         all files have been processed.
    """
    files = collect_files(targets)

    if not files:
        print("No supported files found in the specified targets.")
        sys.exit(1)

    print(f"\nRunning suite on {len(files)} file(s)...\n")

    summary_rows = []

    for file_path in files:
        print("\n" + "=" * 78)
        print(f" FILE: {file_path}")
        print("=" * 78)

        row = {
            "file":         os.path.basename(file_path),
            "before_score": None,
            "before_band":  None,
            "after_score":  None,
            "after_band":   None,
            "fixed_path":   None,
            "error":        None,
        }

        try:
            # ---- Before ----
            before = run_checker(file_path, fix=False)
            print_report(before, "Before Fix")
            row["before_score"] = before["score"]
            row["before_band"]  = before["band"]

            # ---- Fix + After ----
            if fix:
                print("\n--- Applying Fixes ---")
                fix_result = run_checker(file_path, fix=True)

                fixed_path = fix_result.get("fixed_path") or _fixed_path_for(file_path)
                row["fixed_path"] = fixed_path

                if os.path.exists(fixed_path):
                    after = run_checker(fixed_path, fix=False)
                    print_report(after, "After Fix")
                    row["after_score"] = after["score"]
                    row["after_band"]  = after["band"]
                else:
                    print(f"[suite] Warning: fixed file not found at {fixed_path}")

        except Exception as exc:
            row["error"] = str(exc)
            print(f"[suite] ERROR processing {file_path}: {exc}")

        summary_rows.append(row)

    print_suite_summary(summary_rows)


# ============================================================
# ENTRY POINT
# ============================================================

def usage():
    print(
        "Usage:\n"
        "  Single file : python main.py <file.(docx|pptx|pdf)> [--fix]\n"
        "  Test suite  : python main.py --suite <file_or_dir>"
        " [<file_or_dir> ...] [--fix]\n"
        "\n"
        "Options:\n"
        "  --fix    Apply fixes and produce a _fixed document for each input.\n"
        "  --suite  Batch mode — accepts any mix of individual files and\n"
        "           directories (scanned one level deep). Prints a summary\n"
        "           table of before/after scores at the end.\n"
    )


def main():
    args = sys.argv[1:]

    if not args:
        usage()
        sys.exit(1)

    fix   = "--fix"   in args
    suite = "--suite" in args

    # Strip flag tokens so only file/directory paths remain
    paths = [a for a in args if a not in ("--fix", "--suite")]

    if suite:
        # ---- SUITE MODE ----
        if not paths:
            print("Error: --suite requires at least one file or directory path.")
            usage()
            sys.exit(1)
        run_suite(paths, fix)

    else:
        # ---- SINGLE-FILE MODE (original behaviour) ----
        if len(paths) != 1:
            usage()
            sys.exit(1)
        file_path = paths[0]
        if not os.path.exists(file_path):
            print(f"Error: file not found — {file_path}")
            sys.exit(1)
        run_single(file_path, fix)


if __name__ == "__main__":
    main()
