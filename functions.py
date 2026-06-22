"""Multiple functions to support Excel automation.
© Thiha Aung (infowizard@gmail.com)
For the excel, the last row technically is 1048576.
However, I have hard-limited this to 1500 rows.
The code will need to be updated if more rows are needed.
"""

import getpass
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import xlwings as xw  # type: ignore
import string

import hide
import checklist_collections as cc

LEGEND = {
    "UC": "Unit cost in original (buying) currency",
    "SC": "Subtotal cost in original (buying) currency",
    "Discount": "Discount in percentage from the supplier",
    "UCD": "Unit cost after discount in original (buying) currency",
    "SCD": "Subtotal cost after discount in original (buying) currency",
    "UCDQ": "Unit cost after discount in quoted currency (follows main contract quoted currency)",
    "SCDQ": "Subtotal cost after discount in quoted currency (follows main contract quoted currency)",
    "SCDQL": "Subtotal cost after discount in quoted currency lumpsum (follows main contract quoted currency). If lumpsum, indicates lumpsum cost.",
    "TCDQL": "Total cost after discount in quoted currency lumpsum (follows main contract quoted currency). Total lumpsum cost.",
    "BSCQL": "Base subtotal cost in quoted currency lumpsum (follows main contract quoted currency). If lumpsum, indicates lumpsum cost.",
    "BTCQL": "Base total cost in quoted currency lumpsum (follows main contract quoted currency). Total lumpsum cost.",
    "NOTE": "BTCQL is overall system cost in the related excel sheet.",
}

# MACRO_NB references PERSONAL.XLSB which is auto-loaded by Excel from XLSTART folder
# Lazy loading avoids import-time errors when Excel is not running
_MACRO_NB = None

# Cache for PERSONAL.XLSB range references.
# Since PERSONAL.XLSB doesn't change during an Excel session, we cache range
# references to avoid repeated workbook/sheet/range lookups on each copy operation.
_PERSONAL_RANGE_CACHE = {}

# Cache for SharePoint workbook directory lookups
_WORKBOOK_DIR_CACHE: dict[str, tuple[str, bool]] = {}


def get_macro_nb():
    """
    Get the PERSONAL.XLSB workbook (auto-loaded by Excel from XLSTART folder).
    Works on both Mac and Windows.
    """
    global _MACRO_NB
    if _MACRO_NB is None:
        _MACRO_NB = xw.Book("PERSONAL.XLSB")
    return _MACRO_NB


def run_macro(macro_name):
    """Run a VBA macro from PERSONAL.XLSB."""
    get_macro_nb().macro(macro_name)()


def get_macro_sheet(sheet_name):
    """Get a sheet from PERSONAL.XLSB."""
    return get_macro_nb().sheets[sheet_name]


def get_cached_range(sheet_name, range_addr):
    """
    Get a cached range reference from PERSONAL.XLSB.

    Caches range references to avoid repeated workbook/sheet/range lookups.
    The cache is valid for the entire Excel session since PERSONAL.XLSB
    doesn't change during a session.

    Args:
        sheet_name: Name of the sheet in PERSONAL.XLSB (e.g., "Design", "Data")
        range_addr: Range address (e.g., "5:5", "A28:E36", "B1")

    Returns:
        xlwings Range object from the cached reference
    """
    cache_key = f"{sheet_name}:{range_addr}"
    if cache_key not in _PERSONAL_RANGE_CACHE:
        _PERSONAL_RANGE_CACHE[cache_key] = get_macro_sheet(sheet_name).range(range_addr)
    return _PERSONAL_RANGE_CACHE[cache_key]


def copy_design_row(pwb, row_num, dest_range):
    """
    Copy a design row from PERSONAL.XLSB.

    Args:
        pwb: The PERSONAL.XLSB workbook (from get_macro_nb()) - kept for API compatibility
        row_num: The row number to copy from Design sheet (e.g., "5:5" or "21:21")
        dest_range: The destination range object
    """
    get_cached_range("Design", row_num).copy(dest_range)


def apply_lastrow_border(row_range):
    """
    Apply thin blue (#0432FF) top and bottom borders to a subtotal row range.
    Mac: calls apply_subtotal_borders VBA macro (faster than clipboard copy).
    Windows: applies borders directly via COM API.
    """
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlContinuous = 1
    xlThin = 2
    color_bgr = (255 << 16) | (50 << 8) | 4  # #0432FF in BGR long (Windows COM)

    if sys.platform == "win32":
        for edge in [xlEdgeTop, xlEdgeBottom]:
            border = row_range.api.Borders(edge)
            border.LineStyle = xlContinuous
            border.Weight = xlThin
            border.Color = color_bgr
    else:
        row_num = row_range.row
        row_range.sheet.activate()
        get_macro_nb().macro("apply_subtotal_borders")(row_num)


def _has_problematic_path_chars(path: Path) -> bool:
    """Check if path contains characters that cause issues with macOS AppleScript."""
    problematic_chars = ["@", "#", "%"]
    path_str = str(path)
    return any(char in path_str for char in problematic_chars)


def save_workbook_safe(wb, full_path: Path, password: str = "") -> Path:
    """
    Save workbook handling macOS AppleScript path limitations.

    On macOS, paths with special characters (like @) cause AppleScript errors.
    This function saves to ~/Downloads (which Excel has access to), then
    moves to the final destination using Python.

    Args:
        wb: xlwings Workbook object
        full_path: Target path for the saved file
        password: Optional password for the saved file

    Returns:
        The final path where the file was saved
    """
    if sys.platform == "darwin" and _has_problematic_path_chars(full_path):
        # macOS with problematic path - save to Downloads, then move.
        # Downloads folder is accessible by Excel without permission dialogs.
        # Use the real filename (not a temp-prefixed name) so that wb.name in
        # Excel stays correct for subsequent operations (e.g. technical() needs
        # to see "Commercial ..." to pick the right code path).
        downloads = Path.home() / "Downloads"
        temp_path = downloads / full_path.name
        if temp_path.exists():
            temp_path.unlink()
        # Always pass password to SaveAs — even password="" explicitly clears
        # any inherited open-password from the source workbook. The earlier
        # "skip if empty" approach caused Commercial/Technical files to inherit
        # hide.legacy from the source. The reopen dialog that prompted that change
        # was caused by the osascript open lacking the password, not by saving
        # with password=""; that reopen is now fixed via _find_or_open_workbook.
        wb.save(temp_path, password=password)
        # Move to final destination using Python (handles special chars fine)
        if full_path.exists():
            full_path.unlink()  # Remove existing file if present
        shutil.move(str(temp_path), str(full_path))
        return full_path
    else:
        # On Windows with OneDrive/SharePoint, SaveAs to the same path the
        # workbook is already open at fails with COM error -2146827284
        # ("Cannot access").  Detect this by comparing the open workbook's
        # filename to the target; if they match and no password change is
        # needed, use in-place Save() which always works.
        try:
            already_at_target = (wb.name == full_path.name)
        except Exception:
            already_at_target = False

        if already_at_target and not password:
            wb.save()
        else:
            wb.save(full_path, password=password)
        return full_path


def to_pdf_safe(wb, pdf_path: Path, show: bool = True) -> None:
    """
    Export workbook to PDF handling macOS AppleScript path limitations.

    On macOS, paths with special characters (like @) cause AppleScript -50
    errors in wb.to_pdf(). This function exports to ~/Downloads with a temp
    name, moves the file to the final destination using Python, then
    optionally opens the PDF at its real path.

    Args:
        wb: xlwings Workbook object
        pdf_path: Target path for the PDF file
        show: Whether to open the PDF after saving
    """
    if sys.platform == "darwin" and _has_problematic_path_chars(pdf_path):
        downloads = Path.home() / "Downloads"
        temp_name = f"~xltemp_{os.getpid()}_{pdf_path.name}"
        temp_path = downloads / temp_name
        wb.to_pdf(path=str(temp_path), show=False)
        if pdf_path.exists():
            pdf_path.unlink()
        shutil.move(str(temp_path), str(pdf_path))
        if show:
            subprocess.Popen(["open", str(pdf_path)])
    else:
        wb.to_pdf(path=str(pdf_path), show=show)


def _find_or_open_workbook(app, src_path: Path):
    """
    Return the workbook if already open; if not, open it from disk.

    Workbooks created by this tool are password-protected (hide.legacy).
    We pass that password so Excel opens them silently without a dialog.
    Returns None (never raises) — callers use the reference only for activate().
    """
    # Check all open books first — iterate to avoid exact-name lookup failures
    # from path/encoding differences after a Mac SaveAs.
    try:
        for book in app.books:
            try:
                if book.name == src_path.name:
                    return book
            except Exception:
                continue
    except Exception:
        pass
    if not src_path.exists():
        return None
    try:
        import hide as _hide
        pwd = _hide.legacy
    except Exception:
        pwd = ""
    try:
        return open_workbook_safe(app, src_path, password=pwd)
    except Exception:
        return None


def open_workbook_safe(app, full_path: Path, password: str = ""):
    """
    Open a workbook handling macOS AppleScript path limitations.

    On macOS, xlwings' Books.open() uses the appscript bridge which can fail on
    paths with special characters (@, #, %) or when Excel is in certain states.
    This function uses osascript subprocess on Mac (more reliable) and direct
    Books.open() on Windows.
    """
    if sys.platform == "darwin":
        return _open_workbook_mac(app, full_path, password=password)
    if password:
        return app.books.open(str(full_path), password=password)
    return app.books.open(str(full_path))


def _open_workbook_mac(app, full_path: Path, password: str = ""):
    """
    Open a workbook on macOS using osascript, bypassing the xlwings appscript bridge.

    Tries the direct POSIX path first — osascript handles '@' and other special
    characters in file paths just fine (the Downloads copy workaround was only
    needed for the xlwings appscript bridge).  Falls back to a ~/Downloads copy
    only if the direct open fails (e.g. Excel version quirk).
    """
    def _try_open(open_path: Path):
        posix = str(open_path).replace('"', '\\"')
        if password:
            pw = password.replace('"', '\\"')
            script = (
                'tell application "Microsoft Excel"\n'
                f'  set wb to open workbook workbook file name POSIX file "{posix}" password "{pw}"\n'
                '  return name of wb\n'
                'end tell'
            )
        else:
            script = (
                'tell application "Microsoft Excel"\n'
                f'  set wb to open workbook workbook file name POSIX file "{posix}"\n'
                '  return name of wb\n'
                'end tell'
            )
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=60,
        )
        return result

    result = _try_open(full_path)

    if result.returncode != 0 and _has_problematic_path_chars(full_path):
        # Direct open failed — fall back to a Downloads copy
        downloads = Path.home() / "Downloads"
        copy_path = downloads / full_path.name
        if copy_path.exists():
            copy_path.unlink()
        shutil.copy2(str(full_path), str(copy_path))
        result = _try_open(copy_path)
        open_path = copy_path
    else:
        open_path = full_path

    if result.returncode != 0:
        raise RuntimeError(
            f"Cannot open workbook '{open_path}': {result.stderr.strip()}"
        )

    wb_name = result.stdout.strip()
    for name_try in (wb_name, open_path.stem, open_path.name):
        try:
            return app.books[name_try]
        except (KeyError, Exception):
            continue
    # File was opened successfully via osascript but xlwings can't locate it yet.
    # Return None rather than raising — callers that don't need the reference
    # (commercial, technical) work fine; the book is visible in Excel.
    return None


def _get_rfq_base_path() -> Path | None:
    """
    Get the user-specific @rfqs base path based on the current user.

    Returns:
        Path to the @rfqs folder, or None if it doesn't exist.
    """
    # Mac: OneDrive for Business syncs to ~/Library/CloudStorage/OneDrive-*/
    # Search all OneDrive folders there for "Bid Proposal - Documents/@rfqs".
    if sys.platform == "darwin":
        cloud_base = Path.home() / "Library" / "CloudStorage"
        if cloud_base.exists():
            try:
                for od_dir in sorted(cloud_base.iterdir()):
                    candidate = od_dir / "Bid Proposal - Documents" / "@rfqs"
                    if candidate.exists():
                        return candidate
            except PermissionError:
                pass

    username = getpass.getuser()

    # Windows user-specific path configurations
    if username == "oliver":
        base = (
            Path.home()
            / "OneDrive - Jason Electronics Pte Ltd"
            / "Shared Documents"
            / "@rfqs"
        )
    else:
        # Default path for carol_lim and others
        base = (
            Path.home()
            / "Jason Electronics Pte Ltd"
            / "Bid Proposal - Documents"
            / "@rfqs"
        )

    return base if base.exists() else None


def _find_workbook_in_rfqs(workbook_name: str, base_path: Path) -> Path | None:
    """
    Search for a workbook in the @rfqs folder structure.

    Searches year subfolders (2024/, 2025/, 2026/, etc.) up to 5 levels deep.
    Returns the shallowest match if multiple are found.

    Args:
        workbook_name: The filename to search for (e.g., "JEC-2026-001-v1.xlsx")
        base_path: The @rfqs base path to search in

    Returns:
        Path to the directory containing the workbook, or None if not found.
    """
    matches: list[tuple[int, Path]] = []  # (depth, parent_dir)
    workbook_name_lower = workbook_name.lower()
    max_depth = 5

    # Get current year to search recent years first
    current_year = datetime.now().year
    year_folders = []

    # Check years from current down to 2020
    for year in range(current_year, 2019, -1):
        year_path = base_path / str(year)
        if year_path.is_dir() and not year_path.is_symlink():
            year_folders.append(year_path)

    # Use iterative BFS instead of recursion to avoid stack overflow
    # Each item is (directory, depth)
    for year_folder in year_folders:
        queue: list[tuple[Path, int]] = [(year_folder, 1)]

        while queue:
            directory, depth = queue.pop(0)

            if depth > max_depth:
                continue

            try:
                for entry in directory.iterdir():
                    # Skip symbolic links to avoid cycles
                    if entry.is_symlink():
                        continue

                    if entry.is_file() and entry.name.lower() == workbook_name_lower:
                        matches.append((depth, directory))
                    elif entry.is_dir():
                        queue.append((entry, depth + 1))
            except (PermissionError, OSError):
                # Skip directories we can't access
                continue

    if not matches:
        return None

    # Sort by depth (shallowest first)
    matches.sort(key=lambda x: x[0])

    if len(matches) > 1:
        # Multiple matches found - alert user, use shallowest
        print(
            f"Note: Found {len(matches)} locations for '{workbook_name}'. Using: {matches[0][1]}"
        )

    return matches[0][1]


def _resolve_workbook_path(wb) -> "Path | None":
    """Return the local .xlsx path for wb, handling SharePoint/OneDrive URLs."""
    try:
        p = Path(wb.fullname)
        if p.exists() and p.suffix.lower() == ".xlsx":
            return p
    except Exception:
        pass
    try:
        rfq_base = _get_rfq_base_path()
        if rfq_base is not None:
            found_dir = _find_workbook_in_rfqs(wb.name, rfq_base)
            if found_dir is not None:
                p = Path(found_dir) / wb.name
                if p.exists():
                    return p
    except Exception:
        pass
    return None


def get_workbook_directory(wb):
    """
    Get the directory path for a workbook, handling SharePoint/OneDrive URLs.

    When a workbook is opened from SharePoint or OneDrive, wb.fullname may return
    a URL instead of a local file path, or may fail entirely. This function
    first attempts to locate the workbook in the user's synced @rfqs folder,
    then falls back to the Downloads folder.

    Args:
        wb: xlwings Workbook object

    Returns:
        tuple: (directory_path, is_cloud) where:
            - directory_path: Local directory path for saving files
            - is_cloud: True if the file is on SharePoint/OneDrive or fullname failed
    """
    try:
        fullname = wb.fullname
    except Exception:
        # wb.fullname can fail on SharePoint/OneDrive files
        fullname = None

    # Check if it's a SharePoint/OneDrive URL or if fullname failed
    is_cloud = fullname is None or fullname.startswith(("http://", "https://"))

    # Also treat as cloud/unknown if the path no longer exists on disk.
    # This happens when save_workbook_safe() saved to ~/Downloads and moved
    # the file to the real destination — the wb.fullname is then stale.
    # Falling through to the @rfqs search recovers the correct directory.
    if not is_cloud and fullname is not None and not Path(fullname).exists():
        is_cloud = True

    if is_cloud:
        # Try to find the workbook in the @rfqs folder first
        rfq_base = _get_rfq_base_path()
        if rfq_base is not None:
            found_dir = _find_workbook_in_rfqs(wb.name, rfq_base)
            if found_dir is not None:
                return (str(found_dir), True)

        # Fall back to Downloads folder for cloud files
        downloads = Path.home() / "Downloads"

        # Create a subdirectory for the project if possible
        project_name = wb.name[:-5] if wb.name.endswith(".xlsx") else wb.name[:-4]
        project_dir = downloads / project_name

        # Create the directory if it doesn't exist
        project_dir.mkdir(parents=True, exist_ok=True)

        print(f"Note: '{wb.name}' not found in @rfqs. Using: {project_dir}")
        return (str(project_dir), True)
    else:
        assert fullname is not None  # Guaranteed by is_cloud check above
        return (os.path.dirname(fullname), False)


# Accounting number format
ACCOUNTING = "_(* #,##0.00_);_(* (#,##0.00);_(* " "-" "??_);_(@_)"
EXCNANGE_RATE = '_(* #,##0.0000_);_(* (#,##0.0000);_(* "-"????_);_(@_)'

RESOURCES = os.path.join(
    os.path.dirname(os.path.realpath(__file__)),
    "resources/",
)

# To update the value upon updating of the template.
LATEST_WB_VERSION = "R2"
LATEST_MINOR_REVISION = "M3"
UPDATE_MESSAGE = "Now you can choose the number scheme. Single or Double."

# Skipped sheets (includes TN as alias for Technical_Notes)
# Note: "Scratch" is handled case-insensitively via should_skip_sheet()
SKIP_SHEETS = ["Config", "Cover", "Summary", "Technical_Notes", "TN", "T&C", "Scratch", "Proposal"]


def should_skip_sheet(sheet_name):
    """
    Check if a sheet should be skipped during processing.

    Handles case-insensitive matching for "Scratch" sheets, allowing users
    to name their scratch sheet "scratch", "SCRATCH", "Scratch", etc.

    Args:
        sheet_name: The name of the sheet to check (string)

    Returns:
        True if the sheet should be skipped, False otherwise.
    """
    if sheet_name in SKIP_SHEETS:
        return True
    # Case-insensitive check for "scratch"
    if sheet_name.lower() == "scratch":
        return True
    return False


# Sheet name aliases - maps alternative names to canonical sheet names
# Format: {"alias": "canonical_name"}
SHEET_ALIASES = {
    "TN": "Technical_Notes",
}

# Reverse mapping: canonical → list of aliases (built automatically)
_CANONICAL_TO_ALIASES = {}
for alias, canonical in SHEET_ALIASES.items():
    _CANONICAL_TO_ALIASES.setdefault(canonical, []).append(alias)


def resolve_sheet_name(name):
    """
    Resolve a sheet name alias to its canonical name.

    Args:
        name: Sheet name or alias (e.g., "TN" or "Technical_Notes")

    Returns:
        The canonical sheet name (e.g., "Technical_Notes")
    """
    return SHEET_ALIASES.get(name, name)


def get_sheet(wb, name, required=True):
    """
    Get a sheet from a workbook, supporting aliases.

    Tries the canonical name first, then any aliases. Works whether the
    actual sheet in Excel is named "Technical_Notes" or "TN".

    Args:
        wb: xlwings Workbook object
        name: Sheet name or alias (e.g., "TN" or "Technical_Notes")
        required: If False, returns None when sheet doesn't exist instead of raising.

    Returns:
        xlwings Sheet object, or None if required=False and sheet doesn't exist.
    """
    canonical_name = resolve_sheet_name(name)
    sheet_names = wb.sheet_names

    # Try canonical name first
    if canonical_name in sheet_names:
        return wb.sheets[canonical_name]

    # Try aliases if canonical not found
    for alias in _CANONICAL_TO_ALIASES.get(canonical_name, []):
        if alias in sheet_names:
            return wb.sheets[alias]

    # Sheet not found
    if not required:
        return None

    # Fall back to original name (will raise KeyError if not found)
    return wb.sheets[name]


def delete_scratch_sheet(wb):
    """Delete any Scratch sheet (case-insensitive) from the workbook."""
    for sheet_name in list(wb.sheet_names):
        if sheet_name.lower() == "scratch":
            wb.sheets[sheet_name].delete()
            break


def sheet_exists(wb, name):
    """
    Check if a sheet exists in workbook (considering aliases).

    Args:
        wb: xlwings Workbook object
        name: Sheet name or alias (e.g., "TN" or "Technical_Notes")

    Returns:
        True if the sheet exists, False otherwise.
    """
    canonical = resolve_sheet_name(name)
    if canonical in wb.sheet_names:
        return True
    for alias in _CANONICAL_TO_ALIASES.get(canonical, []):
        if alias in wb.sheet_names:
            return True
    return False


def is_sheet_name(name, canonical):
    """
    Check if a sheet name matches a canonical name (including aliases).

    Args:
        name: Sheet name to check (could be an alias)
        canonical: The canonical sheet name to match against

    Returns:
        True if name matches canonical (directly or via alias)
    """
    return resolve_sheet_name(name) == canonical


def set_nitty_gritty(text):
    """Fix annoying text"""
    # Strip EOL
    text = text.strip()
    # Strip 2 or more spaces
    text = re.sub(" {2,}", " ", text)
    # Put bullet point for Sub-subitem preceded by '-' or '~'.
    text = re.sub("^(-|~)", "•", text)
    # Put bullet point for Sub-subitem preceded by a single * followed by space.
    text = re.sub(r"^[*?]\s", " • ", text)
    # Instead of ';' at the end of line, use ':' instead.
    text = re.sub(";$", ":", text)
    text = set_comma_space(text)
    text = set_x(text)
    return text


def set_comma_space(text):
    """Fix having space before comma and not having space after comma"""
    # fix word+space+, to word+,
    x = re.compile(r"\w+\s,")
    if x.search(text):
        substring = re.findall(r"\w+\s,", text)
        for word in substring:
            text = re.sub(word, word[:-2] + ",", text)

    # Fix word+,+no-space to word+,+space
    x = re.compile(r",\d?\w+")
    if x.search(text):
        # Ignores format like 1,200 but matches 1,w
        substring = re.findall(r"(?<![0-9]),\w+", text)
        for word in substring:
            text = re.sub(word, ", " + word[1:], text)
    return text


_TITLE_CASE_LOWER = frozenset({
    "a", "an", "the",
    "and", "but", "or", "nor", "for", "yet", "so",
    "at", "by", "in", "of", "on", "to", "up", "as",
})


def title_case_ignore_double_char(text):
    words = text.split()
    last_idx = len(words) - 1
    titled_words = []
    for i, word in enumerate(words):
        core = word.strip(string.punctuation).lower()
        if i != 0 and i != last_idx and core in _TITLE_CASE_LOWER:
            # Articles, conjunctions, short prepositions stay lowercase
            # regardless of how the user typed them, unless first or last word
            titled_words.append(word.lower())
        elif len(word.strip(string.punctuation)) > 2:
            # To prevent cases like 'mm)' from becoming 'Mm)'
            titled_words.append(word.capitalize())
        else:
            titled_words.append(word)
    return " ".join(titled_words)


def set_case_preserve_acronym(text, title=False, capitalize=False, upper=False):
    """Maintaion acronyms case when using title or sentence"""
    # The regex below essentially ignore the letters in lower case letter.
    # Now cases such as iPhone, mPower, c/w are recognized.
    # acronym_regex = re.compile(r'\b([a-z0-9\.]?[A-Z0-9\/][A-Z0-9a-z-]*)(?=\b|[^a-z])')
    # Remove matching hyphen
    acronym_regex = re.compile(r"\b([a-z0-9\.]?[A-Z0-9\/][A-Z0-9a-z]*)(?=\b|[^a-z])")
    # acronym_regex = re.compile(r'\b([a-z]?[A-Z0-9][A-Z0-9-]*)(?=\b|[^a-z])')
    acronyms = [a for a in acronym_regex.findall(text) if a.lower() not in _TITLE_CASE_LOWER]

    if title:
        text = title_case_ignore_double_char(text)
        # Restore acronyms with IGNORECASE so that dot-separated part numbers like
        # "BTD.NH.TE.TD.NR.RC" are preserved. capitalize() lowercases every character
        # after the first, so ".NH" becomes ".nh" — a case-insensitive pattern is
        # required to match it back.
        for acronym in acronyms:
            pattern = rf"\b{re.escape(acronym)}\b"
            text = re.sub(pattern, acronym, text, flags=re.IGNORECASE)
        return text

    elif capitalize:
        # First change all to lower case
        text = text.lower()
        for acronym in acronyms:
            acronym_regex = acronym.lower()
            pattern = rf"\b{acronym_regex}\b"
            text = re.sub(pattern, acronym, text)
        text = text.capitalize()  # Has not handle the first word
        return text

    elif upper:
        text = text.upper()
        return text


def set_x(text):
    """Normalize quantity notation to 'N ×' format using the multiplication sign.

    All variants (1x, 20X, x1, X 20, x20, etc.) are converted to 'N ×',
    e.g. '2x items' and 'x2 items' both become '2 × items'.
    Hyphenated cases like '20x-connector' are left unchanged.
    """
    # Number-first: 20x, 30X (not followed by -)
    text = re.sub(r"(\d+)[xX](?!-)", r"\1 ×", text)
    # Symbol-first: x20, X30 — flip to number-first
    text = re.sub(r"[xX](\d+)", r"\1 ×", text)
    # Number-first with space: 20 x, 20 X
    text = re.sub(r"(\d+) [xX](?!\S)", r"\1 ×", text)
    # Symbol-first with space: x 20, X 20 — flip to number-first
    text = re.sub(r"[xX] (\d+)", r"\1 ×", text)
    return text


def fill_formula(sheet):
    """
    Fill formulas in a sheet for pricing calculations.

    Optimized to batch adjacent column formula assignments, reducing COM calls
    from ~30 to ~10 for significant performance improvement.
    """
    if not should_skip_sheet(sheet.name):
        # Formula to cells
        # Increase the last row by 1 so that the cells are not left empty
        last_row = sheet.range("C1500").end("up").row + 1
        lr = str(last_row)

        # A1: Reference formula (single cell)
        sheet.range("A1").formula = (
            '= "JASON REF: " & Config!B29 &  ", REVISION: " &  Config!B30 & ", PROJECT: " & Config!B26'
        )

        # B: Serial Numbering (single column)
        sheet.range("B3:B" + lr).formula = (
            '=IF(AND(A3="", ISNUMBER(D3), ISNUMBER(K3)), COUNT(B2:INDEX($B$1:B2, XMATCH("Title", $AL$1:AL2, 0, -1))) + 1 , "")'
        )

        # BATCH 1: Columns N, O (2 adjacent columns) - Cost calculations
        sheet.range("N3:O" + lr).formula = [
            [
                '=IF(K3<>"",K3*(1-M3),"")',  # N: UCD
                '=IF(AND(D3<>"", K3<>"",H3<>"OPTION",H3<>"REMOVED"),D3*N3,"")',  # O: SCD
            ]
        ]

        # BATCH 2: Columns Q through AA (11 adjacent columns) - Exchange rates & escalations
        sheet.range("Q3:AA" + lr).formula = [
            [
                # Q: Exchange rate
                '=IF(J3<>"", INDEX(Config!$B$2:$B$10, XMATCH(J3, Config!$A$2:$A$10, 0))/INDEX(Config!$B$2:$B$10, XMATCH(Config!$B$12, Config!$A$2:$A$10, 0)), "")',
                # R: UCDQ
                '=IF(AND(D3<>"", K3<>""), N3*Q3,"")',
                # S: SCDQ
                '=IF(AND(D3>0, K3<>"", H3<>"OPTION", H3<>"REMOVED", INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*R3, "")',
                # T: BUCQ
                '=IF(AND(D3<>"",K3<>""), (R3*(1+$L$1+$N$1+$P$1+$R$1))/(1-0.05),"")',
                # U: BSCQ
                '=IF(AND(D3>0,K3<>"",H3<>"OPTION",H3<>"REMOVED",INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*T3, "")',
                # V: Default escalation
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION", H3<>"REMOVED"), AQ3*$L$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION", H3<>"REMOVED"), S3*$L$1, ""))',
                # W: Warranty
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION", H3<>"REMOVED"), AQ3*$N$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION", H3<>"REMOVED"), S3*$N$1, ""))',
                # X: Freight
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION", H3<>"REMOVED"), AQ3*$P$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION", H3<>"REMOVED"), S3*$P$1, ""))',
                # Y: Special
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION", H3<>"REMOVED"), AQ3*$R$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION", H3<>"REMOVED"), S3*$R$1, ""))',
                # Z: Risk
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION", H3<>"REMOVED"), AS3-(AQ3+V3+W3+X3+Y3), IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION", H3<>"REMOVED"), U3-(S3+V3+W3+X3+Y3), ""))',
                # AA: Margin reference
                '=IF(AND(D3<>"",K3<>""),$J$1,"")',
            ]
        ]

        # BATCH 3: Columns AC through AI (7 adjacent columns) - Pricing calculations
        sheet.range("AC3:AI" + lr).formula = [
            [
                # AC: RUPQ
                '=IF(AND(D3<>"",K3<>""),CEILING(T3/(1-AA3), 1),"")',
                # AD: RSPQ
                '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED", H3<>"REMOVED",INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*AC3,"")',
                # AE: UPLS
                '=IF(AND(D3<>"",K3<>""), IF(AB3<>"", AB3, AC3),"")',
                # AF: SPLS
                '=IF(AND(D3>0,K3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED", H3<>"REMOVED",INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*AE3,"")',
                # AG: Profit
                '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"REMOVED",AF3<>""),AF3-U3,"")',
                # AH: Margin %
                '=IF(AND(AG3<>"", AG3<>0), AG3/AF3, "")',
                # AI: Total price
                '=IF(AND(D3>0,K3<>"", H3<>"OPTION", H3<>"REMOVED"), D3*AE3, "")',
            ]
        ]

        # BATCH 4: Columns F, G (2 adjacent columns) - Unit/Subtotal Price
        sheet.range("F3:G" + lr).formula = [
            [
                # F: Unit Price
                '=IF(AND(AL3="Title", ISNUMBER(AJ3)), AJ3, IF(H3="REMOVED", "", IF(AND(AL3="Lineitem", AK3="Lumpsum", H3<>"OPTION"), "", AE3)))',
                # G: Subtotal Price
                '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED", H3<>"REMOVED"), D3*F3,"")',
            ]
        ]

        # L: Subtotal Cost (single column)
        sheet.range("L3:L" + lr).formula = (
            '=IF(AND(D3<>"",K3<>"",H3<>"OPTION",H3<>"REMOVED"),D3*K3,"")'
        )

        # AL: Format field (special handling - values and formulas)
        sheet.range("AL1").value = "Title"
        sheet.range("AL3").value = "System"
        sheet.range("AL4:AL" + lr).formula = (
            '=IF(C4<>"",IF(AND(A4<>"",C4<>""),"Title", IF(B4<>"","Lineitem", IF(LEFT(C4,3)="***","Comment", IF(AND(A4="",B4="",C3="", C5<>"",D5<>""), "Subtitle", IF(AND(A4="",B4="",C3="", C5=""), "Subsystem", "Description"))))),"")'
        )
        sheet.range("AL" + str(last_row + 1)).value = "Title"

        # BATCH 5: Columns AJ, AK (2 adjacent columns) - Lumpsum flags
        sheet.range("AJ3:AK" + lr).formula = [
            [
                # AJ: Lumpsum total
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(AI4:INDEX(AI4:AI1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), "")',
                # AK: Lumpsum/Unit Price flag
                '=IF(AL3="Lineitem", IF(ISNUMBER(INDEX($AJ$1:AJ2, XMATCH("Title", $AL$1:AL2, 0, -1))), "Lumpsum", "Unit Price"), "")',
            ]
        ]

        # BATCH 6: Columns AP through AW (8 adjacent columns) - Lumpsum calculations
        sheet.range("AP3:AW" + lr).formula = [
            [
                # AP: SCDQL
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(S4:INDEX(S4:S1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), IF(AND(AL3="Lineitem", AK3="Unit Price"), R3, ""))',
                # AQ: TCDQL (material cost)
                '=IF(AND(ISNUMBER(D3), ISNUMBER(AP3), H3<>"OPTION", H3<>"REMOVED"), D3*AP3, "")',
                # AR: BSCQL
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(U4:INDEX(U4:U1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), IF(AND(AL3="Lineitem", AK3="Unit Price"), T3, ""))',
                # AS: BTCQL (base cost)
                '=IF(AND(ISNUMBER(D3), ISNUMBER(AR3), H3<>"OPTION", H3<>"REMOVED"), D3*AR3, "")',
                # AT: SSPL
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(AF4:INDEX(AF4:AF1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), IF(AND(AL3="Lineitem", AK3="Unit Price"), AE3, ""))',
                # AU: TSPL (selling price)
                '=IF(AND(ISNUMBER(D3), H3<>"WAIVED", H3<>"INCLUDED", H3<>"OPTION", H3<>"REMOVED", ISNUMBER(AT3)), D3*AT3, "")',
                # AV: Total Profit
                '=IF(AND(ISNUMBER(D3), ISNUMBER(AS3), ISNUMBER(AU3)), AU3-AS3, "")',
                # AW: Grand Margin
                '=IF(AND(H3<>"OPTION", H3<>"REMOVED", ISNUMBER(D3), ISNUMBER(AU3), AU3<>0, ISNUMBER(AV3)), AV3/AU3, "")',
            ]
        ]


def sanitize_config_string(value):
    """Sanitize Config string: remove newlines, collapse spaces, strip whitespace."""
    if not isinstance(value, str):
        return value
    text = value.replace("\n", " ").replace("\r", " ")
    text = re.sub(" {2,}", " ", text)
    return text.strip()


def sanitize_config_date(value):
    """Convert date to ISO format (yyyy-mm-dd). Uses day-first for ambiguous dates."""
    if value is None or value == "":
        return value
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if not isinstance(value, str):
        return value
    if re.match(r"^\d{4}-\d{2}-\d{2}$", value.strip()):
        return value.strip()
    from dateutil import parser as date_parser

    try:
        parsed = date_parser.parse(value, dayfirst=True)
        return parsed.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return value


def sanitize_config_sheet(wb):
    """Sanitize Config sheet cells B21-B32 before filling formulas."""
    try:
        config = wb.sheets["Config"]
    except KeyError:
        return
    for row in range(21, 32):  # B21-B31 strings
        cell = config.range(f"B{row}")
        original = cell.value
        sanitized = sanitize_config_string(original)
        if sanitized != original:
            cell.value = sanitized
    # B32 date
    cell = config.range("B32")
    original = cell.value
    sanitized = sanitize_config_date(original)
    if sanitized != original:
        cell.value = sanitized
    # Disable word wrap for B21:B32
    config.range("B21:B32").wrap_text = False


def fill_formula_wb(wb):
    sanitize_config_sheet(wb)
    for sheet in wb.sheets:
        fill_formula(sheet)


def fill_lastrow(wb):
    for sheet in wb.sheets:
        if not should_skip_sheet(sheet.name):
            fill_lastrow_sheet(wb, sheet)


def fill_lastrow_sheet(wb, sheet):  # type: ignore
    if not should_skip_sheet(sheet.name):
        last_row = sheet.range("C1500").end("up").row
        sr = last_row + 2  # subtotal row (last_row+1 is spacer)
        row_range = sheet.range(f"{sr}:{sr}")
        apply_lastrow_border(row_range)
        sheet.range(f"F{sr}").formula = '="Subtotal(" & Config!B12 & ")"'
        sheet.range(f"F{sr}").font.size = 9
        sheet.range(f"G{sr}").formula = f"=SUM(G3:G{last_row + 1})"
        # Default
        sheet.range(f"V{sr}").formula = f"=SUM(V3:V{last_row + 1})"
        # Warranty
        sheet.range(f"W{sr}").formula = f"=SUM(W3:W{last_row + 1})"
        # Freight (Inbound)
        sheet.range(f"X{sr}").formula = f"=SUM(X3:X{last_row + 1})"
        # Special (Conditions)
        sheet.range(f"Y{sr}").formula = f"=SUM(Y3:Y{last_row + 1})"
        # Risk
        sheet.range(f"Z{sr}").formula = f"=SUM(Z3:Z{last_row + 1})"
        sheet.range(f"AL{sr}").value = "Title"
        # TCDQL — material cost
        sheet.range(f"AQ{sr}").formula = f"=SUM(AQ3:AQ{last_row + 1})"
        # BTCQL — base price after escalation
        sheet.range(f"AS{sr}").formula = f"=SUM(AS3:AS{last_row + 1})"
        # TSPL — actual selling price
        sheet.range(f"AU{sr}").formula = f"=SUM(AU3:AU{last_row + 1})"
        # TP — total profit
        sheet.range(f"AV{sr}").formula = f"=SUM(AV3:AV{last_row + 1})"
        # Total margin
        sheet.range(f"AW{sr}").formula = f'=IF(AU{sr}<>0,AV{sr}/AU{sr},"")'
        sheet.range(f"AW{sr}").number_format = "0.00%"
        # Format subtotal row
        sheet.range(f"V{sr}:Z{sr}").font.color = (0, 144, 81)
        sheet.range(f"{sr}:{sr}").font.bold = True

        sheet.page_setup.print_area = f"A1:H{sr}"


def unhide_columns(sheet):
    """Unhide all columns while setting the width for selected columns"""
    if not should_skip_sheet(sheet.name):
        sheet.range("A:A").column_width = 5
        sheet.range("B:B").autofit()
        sheet.range("C:C").column_width = 55
        sheet.range("C:C").rows.autofit()
        sheet.range("C:C").wrap_text = True
        sheet.range("D:H").autofit()
        sheet.range("I:AQ").wrap_text = False
        sheet.range("I:I").column_width = 10
        sheet.range("I:I").wrap_text = False
        sheet.range("J:O").autofit()
        sheet.range("P:P").column_width = 20
        sheet.range("P:P").wrap_text = False
        sheet.range("Q:AP").autofit()


def unhide_columns_wb(wb):
    for sheet in wb.sheets:
        unhide_columns(sheet)


def _set_wrap_row_heights(sheet, col_width=55):
    """Set col-C row heights using ReportLab metrics — bypasses Excel's conservative autofit.

    Excel word-wrap is conservative: text that just barely fits on one line often
    causes Excel to allocate a blank second line, creating ugly gaps in PDFs.
    We calculate line count independently with ReportLab (Helvetica ≈ Arial) and
    set row height precisely, so no blank lines form.

    _SP_MDW_PX is platform-calibrated: Mac ≈ 8 (72 DPI effective),
    Windows ≈ 7 (96 DPI) — see constant definition for details.
    """
    last_row = sheet.range("C1500").end("up").row
    if last_row < 2:
        return
    sheet.range(f"2:{last_row}").row_height = _SP_ROW_H
    c_vals = sheet.range(f"C2:C{last_row}").value
    if not isinstance(c_vals, list):
        c_vals = [c_vals]
    for i, val in enumerate(c_vals):
        text = str(val).strip() if val else ""
        if not text:
            continue
        lines = _sp_wrap_lines(text, col_width)
        if lines > 1:
            row_num = i + 2
            sheet.range(f"{row_num}:{row_num}").row_height = _SP_ROW_H * lines


def adjust_columns(sheet):
    """Unhide all columns while setting the width for selected columns"""
    if not should_skip_sheet(sheet.name):
        sheet.range("A:A").column_width = 5
        sheet.range("B:B").autofit()
        sheet.range("C:C").column_width = 55
        sheet.range("C:C").wrap_text = True
        _set_wrap_row_heights(sheet)
        sheet.range("D:H").autofit()


def adjust_columns_wb(wb):
    for sheet in wb.sheets:
        adjust_columns(sheet)


def hide_columns(sheet):
    if not should_skip_sheet(sheet.name):
        sheet.activate()
        run_macro("hide_proposal_columns")


def hide_columns_wb(wb):
    for sheet in wb.sheets:
        hide_columns(sheet)


def set_row_heights_wb(wb):
    for sheet in wb.sheets:
        if not should_skip_sheet(sheet.name):
            _set_wrap_row_heights(sheet)


def summary(wb, discount=False, detail=False, simulation=True, discount_level=15):
    # Calculate first to ensure we read fresh values (not stale)
    wb.app.calculate()

    summary_formula = []
    collect = []  # Collect formula to be put in summary page.
    # formula_fragment = '=IF(OR(Config!B13="COMMERCIAL PROPOSAL", Config!B13="BUDGETARY PROPOSAL"),'
    # The design will now be taken from PERSONAL.XLSB (Windows only)
    pwb = get_macro_nb()

    # Initialize counters
    start_row = 19
    count = 1
    offset = 20
    sheet = wb.sheets["Summary"]

    # Need to collect information if already exists so that it can be repopulated
    # Count actual system sheets (excluding skipped sheets like Config, Cover, Scratch, etc.)
    system_count = sum(1 for s in wb.sheet_names if not should_skip_sheet(s))
    remarks = {}
    discount_price = 0

    # Collect the remarks on summary sheet, such as 'OPTION'
    # It will collect without checking whether data exists or not
    for item in range(system_count):
        remarks[sheet.range(f"C{start_row+1+item}").value] = [
            sheet.range(f"E{start_row+1+item}").value
        ]

    # Collect discount
    if sheet.range(f"C{system_count+start_row+3}").value in [
        "SPECIAL DISCOUNT",
        "SPECIAL PROJECT DISCOUNT",
    ]:
        discount_price = sheet.range(f"D{system_count+start_row+3}").value

    if detail:
        # Collect formula
        for sheet in wb.sheet_names:
            if not should_skip_sheet(sheet):
                sheet = wb.sheets[sheet]
                last_row = sheet.range("G1500").end("up").row
                collect = [
                    "='" + sheet.name + "'!$C$3",
                    # Selling price
                    "='" + sheet.name + "'!$G$" + str(last_row),
                    # "='" + sheet.name + "'!$S$" + str(last_row),
                    # Material cost
                    "='" + sheet.name + "'!$AQ$" + str(last_row),
                    # Escalations
                    "='" + sheet.name + "'!$V$" + str(last_row),
                    "='" + sheet.name + "'!$W$" + str(last_row),
                    "='" + sheet.name + "'!$X$" + str(last_row),
                    "='" + sheet.name + "'!$Y$" + str(last_row),
                    "='" + sheet.name + "'!$Z$" + str(last_row),
                    # Base cost after escalations
                    "='" + sheet.name + "'!$AS$" + str(last_row),
                ]
                summary_formula.extend(collect)
                collect = []

        # Reverse the order of collected items
        odered_summary_formula = summary_formula[::-1]

        # Set sheet to summary
        sheet = wb.sheets["Summary"]
        # Clear summary page
        sheet.range("A18:Z1000").clear()
        # Set format
        sheet.range("C:C").column_width = 55
        # sheet.range('E20:E1000').horizontal_alignment = 'center'

        for system in wb.sheet_names:
            if not should_skip_sheet(system):
                copy_design_row(
                    pwb, "21:21", sheet.range(str(offset) + ":" + str(offset))
                )
                sheet.range("B" + str(offset)).value = str(count) + " ‣ "
                sheet.range("C" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("D" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range(f"G{offset}").formula = (
                    f'=IF(E{offset}<>"OPTION", IF(D{start_row+system_count+2}>0.00001, D{offset}/D{start_row+system_count+2}, ""), "")'  # For scope percentage
                )
                sheet.range("H" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("I" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("J" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("K" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("L" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("M" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("N" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("O" + str(offset)).formula = (
                    "=IF(N"
                    + str(offset)
                    + '<>"",D'
                    + str(offset)
                    + "- N"
                    + str(offset)
                    + ',"")'
                )
                sheet.range("P" + str(offset)).formula = (
                    "=IF(OR(D"
                    + str(offset)
                    + ">0.00001, D"
                    + str(offset)
                    + "<-0.00001), O"
                    + str(offset)
                    + "/D"
                    + str(offset)
                    + ", 0)"
                )
                count += 1
                offset += 1

        # Drawing lines
        copy_design_row(
            pwb, "15:15", sheet.range(str(start_row) + ":" + str(start_row))
        )
        copy_design_row(pwb, "11:11", sheet.range(str(offset) + ":" + str(offset)))
        copy_design_row(
            pwb, "17:17", sheet.range(str(offset + 1) + ":" + str(offset + 1))
        )

        # sheet = wb.sheets['Summary']
        sheet.range("C" + str(offset + 1)).value = (
            '="TOTAL PROJECT (" & Config!B12 & ")"'
        )
        sheet.range("D" + str(offset + 1)).formula = (
            "=SUMIFS(D20:D" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("E" + str(offset + 1)).formula = (
            "=IF(COUNTIF(E20:E" + str(offset) + ',"OPTION"), "Excluding Option", "")'
        )
        sheet.range("H" + str(offset + 1)).formula = (
            "=SUMIFS(H20:H" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("I" + str(offset + 1)).formula = (
            "=SUMIFS(I20:I" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("J" + str(offset + 1)).formula = (
            "=SUMIFS(J20:J" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("K" + str(offset + 1)).formula = (
            "=SUMIFS(K20:K" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("L" + str(offset + 1)).formula = (
            "=SUMIFS(L20:L" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("M" + str(offset + 1)).formula = (
            "=SUMIFS(M20:M" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("N" + str(offset + 1)).formula = (
            "=SUMIFS(N20:N" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("O" + str(offset + 1)).formula = (
            "=IF(N"
            + str(offset + 1)
            + '<>"", D'
            + str(offset + 1)
            + "- N"
            + str(offset + 1)
            + ',"")'
        )
        sheet.range("P" + str(offset + 1)).formula = (
            "=IF(OR(D"
            + str(offset + 1)
            + ">0.00001, D"
            + str(offset + 1)
            + "<-0.00001), O"
            + str(offset + 1)
            + "/D"
            + str(offset + 1)
            + ", 0)"
        )

        # Format
        sheet.range(f"D20:O{offset+1}").number_format = ACCOUNTING
        sheet.range(f"H20:H{offset+1}").font.color = (4, 50, 255)
        sheet.range(f"I20:M{offset+1}").font.color = (148, 55, 255)
        sheet.range(f"P20:P{offset+1}").number_format = "0.00%"
        sheet.range(f"G20:G{offset+1}").number_format = "0.00%"  # For scope percentage
        sheet.range(f"G20:G{offset+1}").font.color = (0, 128, 0)  # Teal

        # Write back remarks
        for item in range(system_count):
            if sheet.range(f"C{start_row+1+item}").value in remarks:
                sheet.range(f"E{start_row+1+item}").value = remarks[
                    sheet.range(f"C{start_row+1+item}").value
                ]

        if discount:
            copy_design_row(
                pwb, "18:18", sheet.range(str(offset + 2) + ":" + str(offset + 2))
            )
            copy_design_row(
                pwb, "19:19", sheet.range(str(offset + 3) + ":" + str(offset + 3))
            )
            sheet.range("C" + str(offset + 3)).formula = (
                '="TOTAL PROJECT PRICE AFTER DISCOUNT (" & Config!B12 & ")"'
            )
            sheet.range("D" + str(offset + 3)).formula = (
                "=SUM(D" + str(offset + 1) + ":D" + str(offset + 2) + ")"
            )
            # Number format for discout field
            sheet.range("D" + str(offset + 2)).number_format = ACCOUNTING
            sheet.range("D" + str(offset + 3)).number_format = ACCOUNTING
            sheet.range("N" + str(offset + 3)).formula = "=$N$" + str(offset + 1)
            sheet.range("N" + str(offset + 3)).number_format = ACCOUNTING
            sheet.range("O" + str(offset + 3)).formula = (
                "=IF(N"
                + str(offset + 3)
                + '<>"", D'
                + str(offset + 3)
                + "- N"
                + str(offset + 3)
                + ',"")'
            )
            sheet.range("O" + str(offset + 3)).number_format = ACCOUNTING
            sheet.range("P" + str(offset + 3)).formula = (
                "=IF(OR(D"
                + str(offset + 3)
                + ">0.00001, D"
                + str(offset + 3)
                + "<-0.00001), O"
                + str(offset + 3)
                + "/D"
                + str(offset + 3)
                + ", 0)"
            )
            sheet.range("P" + str(offset + 3)).number_format = "0.00%"
            sheet.range("C" + str(offset + 5)).formula = (
                '="• All the prices are in " & Config!B12 & " excluding GST."'
            )
            sheet.range("C" + str(offset + 6)).value = (
                "• Total project price does not include prices for optional items set out in the detailed bill of material."
            )
            sheet.range("C" + str(offset + 7)).value = (
                "• Items marked as 'INCLUDED' or 'WAIVED' are included in the scope of supply without price impact."
            )

            # Write back the discount
            if sheet.range(f"C{system_count+start_row+3}").value in [
                "SPECIAL DISCOUNT",
                "SPECIAL PROJECT DISCOUNT",
            ]:
                sheet.range(f"D{system_count+start_row+3}").value = discount_price

            # Discount percentages simulation
            if simulation:
                sheet.range(f"H{offset+5}").value = "Actual Dis"
                sheet.range(f"I{offset+5}").formula = f"=-D{offset+2}/D{offset+1}"
                sheet.range(f"I{offset+5}").number_format = "0.00%"
                sheet.range(f"H{offset+6}").value = "Price"
                sheet.range(f"I{offset+6}").value = "D%"
                sheet.range(f"J{offset+6}").value = "Discount"
                sheet.range(f"K{offset+6}").value = "D Price"
                sheet.range(f"L{offset+6}").value = "Cost"
                sheet.range(f"M{offset+6}").value = "Profit"
                sheet.range(f"N{offset+6}").value = "MU"
                for i in range(discount_level):
                    sheet.range(f"H{offset+7+i}").formula = f"=D{offset+1}"
                    sheet.range(f"I{offset+7+i}").value = (i + 1) / 100
                    sheet.range(f"J{offset+7+i}").formula = (
                        f"=CEILING(H{offset+7+i}*I{offset+7+i},1)"
                    )
                    sheet.range(f"K{offset+7+i}").formula = (
                        f"=H{offset+7+i}-J{offset+7+i}"
                    )
                    sheet.range(f"L{offset+7+i}").formula = f"=N{offset+1}"
                    sheet.range(f"M{offset+7+i}").formula = (
                        f"=K{offset+7+i}-L{offset+7+i}"
                    )
                    sheet.range(f"N{offset+7+i}").formula = (
                        f"=M{offset+7+i}/K{offset+7+i}"
                    )
                # Format
                sheet.range(f"H{offset+7}:H{offset+7+discount_level}").number_format = (
                    ACCOUNTING
                )
                sheet.range(f"I{offset+7}:I{offset+7+discount_level}").number_format = (
                    "0.00%"
                )
                sheet.range(f"J{offset+7}:M{offset+7+discount_level}").number_format = (
                    ACCOUNTING
                )
                sheet.range(f"N{offset+7}:N{offset+7+discount_level}").number_format = (
                    "0.00%"
                )

        else:
            sheet.range("C" + str(offset + 3)).formula = (
                '="• All the prices are in " & Config!B12 & " excluding GST."'
            )
            sheet.range("C" + str(offset + 4)).value = (
                "• Total project price does not include items marked 'OPTION' or 'REMOVED' in the detailed bill of material."
            )
            sheet.range("C" + str(offset + 5)).value = (
                "• Items marked as 'INCLUDED' or 'WAIVED' are included in the scope of supply without price impact."
            )

    else:
        for sheet in wb.sheet_names:
            if not should_skip_sheet(sheet):
                sheet = wb.sheets[sheet]
                last_row = sheet.range("G1500").end("up").row
                collect = [
                    "='" + sheet.name + "'!$C$3",
                    "='" + sheet.name + "'!$G$" + str(last_row),
                    "='" + sheet.name + "'!$AS$" + str(last_row),
                ]
                summary_formula.extend(collect)
                collect = []

        # Reverse the order of collected items
        odered_summary_formula = summary_formula[::-1]

        # Set sheet to summary
        sheet = wb.sheets["Summary"]
        # Clear summary page
        sheet.range("A18:Z1000").clear()
        # Set format
        sheet.range("C:C").column_width = 55
        # sheet.range('E20:E1000').horizontal_alignment = 'center'

        for system in wb.sheet_names:
            if not should_skip_sheet(system):
                copy_design_row(
                    pwb, "21:21", sheet.range(str(offset) + ":" + str(offset))
                )
                sheet.range("B" + str(offset)).value = str(count) + " ‣ "
                sheet.range("C" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("D" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range(f"G{offset}").formula = (
                    f'=IF(E{offset}<>"OPTION", IF(D{start_row+system_count+2}>0.00001, D{offset}/D{start_row+system_count+2}, ""), "")'  # For scope percentage
                )
                sheet.range("H" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("I" + str(offset)).formula = (
                    "=IF(H"
                    + str(offset)
                    + '<>"",D'
                    + str(offset)
                    + "- H"
                    + str(offset)
                    + ',"")'
                )
                sheet.range("J" + str(offset)).formula = (
                    "=IF(OR(D"
                    + str(offset)
                    + ">0.00001, D"
                    + str(offset)
                    + "<-0.00001), I"
                    + str(offset)
                    + "/D"
                    + str(offset)
                    + ", 0)"
                )
                count += 1
                offset += 1

        # Drawing lines
        copy_design_row(
            pwb, "13:13", sheet.range(str(start_row) + ":" + str(start_row))
        )
        copy_design_row(pwb, "11:11", sheet.range(str(offset) + ":" + str(offset)))
        copy_design_row(
            pwb, "7:7", sheet.range(str(offset + 1) + ":" + str(offset + 1))
        )

        # sheet = wb.sheets['Summary']
        sheet.range("C" + str(offset + 1)).value = (
            '="TOTAL PROJECT (" & Config!B12 & ")"'
        )
        sheet.range("D" + str(offset + 1)).formula = (
            "=SUMIFS(D20:D" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("E" + str(offset + 1)).formula = (
            "=IF(COUNTIF(E20:E" + str(offset) + ',"OPTION"), "Excluding Option", "")'
        )
        sheet.range("H" + str(offset + 1)).formula = (
            "=SUMIFS(H20:H" + str(offset) + ",E20:E" + str(offset) + ',"<>OPTION",E20:E' + str(offset) + ',"<>REMOVED")'
        )
        sheet.range("I" + str(offset + 1)).formula = (
            "=IF(H"
            + str(offset + 1)
            + '<>"", D'
            + str(offset + 1)
            + "- H"
            + str(offset + 1)
            + ',"")'
        )
        sheet.range("J" + str(offset + 1)).formula = (
            "=IF(OR(D"
            + str(offset + 1)
            + ">0.00001, D"
            + str(offset + 1)
            + "<-0.00001), I"
            + str(offset + 1)
            + "/D"
            + str(offset + 1)
            + ", 0)"
        )

        # Format
        sheet.range(f"D20:I{offset+1}").number_format = ACCOUNTING
        sheet.range(f"G20:G{offset+1}").number_format = "0.00%"  # For scope percentage
        sheet.range(f"G20:G{offset+1}").font.color = (0, 128, 0)  # Teal
        sheet.range(f"J20:J{offset+1}").number_format = "0.00%"

        # Write back remarks
        for item in range(system_count):
            if sheet.range(f"C{start_row+1+item}").value in remarks:
                sheet.range(f"E{start_row+1+item}").value = remarks[
                    sheet.range(f"C{start_row+1+item}").value
                ]

        if discount:
            copy_design_row(
                pwb, "8:8", sheet.range(str(offset + 2) + ":" + str(offset + 2))
            )
            copy_design_row(
                pwb, "9:9", sheet.range(str(offset + 3) + ":" + str(offset + 3))
            )
            sheet.range("C" + str(offset + 3)).formula = (
                '="TOTAL PROJECT PRICE AFTER DISCOUNT (" & Config!B12 & ")"'
            )
            sheet.range("D" + str(offset + 3)).formula = (
                "=SUM(D" + str(offset + 1) + ":D" + str(offset + 2) + ")"
            )
            # Number format for discout field
            sheet.range("D" + str(offset + 2)).number_format = ACCOUNTING
            sheet.range("D" + str(offset + 3)).number_format = ACCOUNTING
            sheet.range("H" + str(offset + 3)).formula = "=$H$" + str(offset + 1)
            sheet.range("H" + str(offset + 3)).number_format = ACCOUNTING
            sheet.range("I" + str(offset + 3)).formula = (
                "=IF(H"
                + str(offset + 3)
                + '<>"", D'
                + str(offset + 3)
                + "- H"
                + str(offset + 3)
                + ',"")'
            )
            sheet.range("I" + str(offset + 3)).number_format = ACCOUNTING
            # sheet.range('J' + str(offset+3)).formula = '=IF(I' + str(offset+3) + '<>0,I' + str(offset+3) + '/D' + str(offset+3) + ',"")'
            sheet.range("J" + str(offset + 3)).formula = (
                "=IF(OR(D"
                + str(offset + 3)
                + ">0.00001, D"
                + str(offset + 3)
                + "<-0.00001), I"
                + str(offset + 3)
                + "/D"
                + str(offset + 3)
                + ", 0)"
            )
            sheet.range("J" + str(offset + 3)).number_format = "0.00%"
            sheet.range("C" + str(offset + 5)).formula = (
                '="• All the prices are in " & Config!B12 & " excluding GST."'
            )
            sheet.range("C" + str(offset + 6)).value = (
                "• Total project price does not include prices for optional items set out in the detailed bill of material."
            )
            sheet.range("C" + str(offset + 7)).value = (
                "• Items marked as 'INCLUDED' or 'WAIVED' are included in the scope of supply without price impact."
            )

            # Write back the discount
            if sheet.range(f"C{system_count+start_row+3}").value in [
                "SPECIAL DISCOUNT",
                "SPECIAL PROJECT DISCOUNT",
            ]:
                sheet.range(f"D{system_count+start_row+3}").value = discount_price

            # Discount percentages simulation
            # Discount percentages simulation
            if simulation:
                sheet.range(f"H{offset+5}").value = "Actual Dis"
                sheet.range(f"I{offset+5}").formula = f"=-D{offset+2}/D{offset+1}"
                sheet.range(f"I{offset+5}").number_format = "0.00%"
                sheet.range(f"H{offset+6}").value = "Price"
                sheet.range(f"I{offset+6}").value = "D%"
                sheet.range(f"J{offset+6}").value = "Discount"
                sheet.range(f"K{offset+6}").value = "D Price"
                sheet.range(f"L{offset+6}").value = "Cost"
                sheet.range(f"M{offset+6}").value = "Profit"
                sheet.range(f"N{offset+6}").value = "MU"
                for i in range(discount_level):
                    sheet.range(f"H{offset+7+i}").formula = f"=D{offset+1}"
                    sheet.range(f"I{offset+7+i}").value = (i + 1) / 100
                    sheet.range(f"J{offset+7+i}").formula = (
                        f"=CEILING(H{offset+7+i}*I{offset+7+i},1)"
                    )
                    sheet.range(f"K{offset+7+i}").formula = (
                        f"=H{offset+7+i}-J{offset+7+i}"
                    )
                    sheet.range(f"L{offset+7+i}").formula = f"=H{offset+1}"
                    sheet.range(f"M{offset+7+i}").formula = (
                        f"=K{offset+7+i}-L{offset+7+i}"
                    )
                    sheet.range(f"N{offset+7+i}").formula = (
                        f"=M{offset+7+i}/K{offset+7+i}"
                    )
                # Format
                sheet.range(f"H{offset+7}:H{offset+7+discount_level}").number_format = (
                    ACCOUNTING
                )
                sheet.range(f"I{offset+7}:I{offset+7+discount_level}").number_format = (
                    "0.00%"
                )
                sheet.range(f"J{offset+7}:M{offset+7+discount_level}").number_format = (
                    ACCOUNTING
                )
                sheet.range(f"N{offset+7}:N{offset+7+discount_level}").number_format = (
                    "0.00%"
                )

        else:
            sheet.range("C" + str(offset + 3)).formula = (
                '="• All the prices are in " & Config!B12 & " excluding GST."'
            )
            sheet.range("C" + str(offset + 4)).value = (
                "• Total project price does not include items marked 'OPTION' or 'REMOVED' in the detailed bill of material."
            )
            sheet.range("C" + str(offset + 5)).value = (
                "• Items marked as 'INCLUDED' or 'WAIVED' are included in the scope of supply without price impact."
            )

    # Calculate all formulas written to summary sheet to avoid stale values
    wb.app.calculate()

    sheet.range("D:D").autofit()
    sheet.range("E:E").autofit()
    sheet.range("F:P").autofit()
    last_row = sheet.range("C1500").end("up").row
    sheet.page_setup.print_area = "A1:F" + str(last_row + 3)


def get_num_scheme(wb):
    """
    Get numbering scheme parameters from Config B16
    - Double (or empty/None): count=10, step=10
    """
    scheme = wb.sheets["Config"].range("B16").value
    if scheme and str(scheme).strip().upper() == "SINGLE":
        return 1, 1
    return 10, 10  # Default to Double


def number_title(wb, count=10, step=10):
    """
    For the main numbering. It will fix as long as it is a number.
    Need to look for only the systems and engineering services.
    Takes a work book, then start number and step.

    Optimized to use vectorized pandas operations instead of row-by-row iteration.
    """
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheets:
        if not should_skip_sheet(sheet.name):
            system_names.append(str.upper(sheet.name))
            ws = wb.sheets[sheet]
            last_row = ws.range("C1500").end("up").row
            data = (
                ws.range("A2:C" + str(last_row))
                .options(pd.DataFrame, index=False)
                .value
            )
            data["System"] = str.upper(sheet.name)
            systems = pd.concat([systems, data], join="outer")

    # Now that I have collect the data, let us do the numbering
    # Index is reset so that index number is continuous
    systems = systems.reset_index(drop=True)
    # Reindexing will remove columns that are not named.
    systems = systems.reindex(columns=["NO", "Description", "System"])

    # Vectorized approach:
    # 1. Identify numeric values (main titles)
    # 2. Identify strings not starting with A-Z (sub-items)
    # 3. Use cumsum to group sub-items under their parent title
    # 4. Assign numbers using vectorized operations

    # Convert to string for consistent checking, handle NaN
    no_col = systems["NO"].fillna("")

    # Check if each value can be converted to int (is a main title number)
    def is_numeric(x):
        try:
            return bool(int(x)) if x != "" else False
        except (ValueError, TypeError):
            return False

    is_main_title = no_col.apply(is_numeric)

    # Check if string starts with A-Z (should be kept as-is)
    def starts_with_letter(x):
        if isinstance(x, str) and x.strip():
            return bool(re.match(r"^[A-Z]", x.strip()))
        return False

    starts_with_az = no_col.apply(starts_with_letter)

    # Identify sub-items: strings that don't start with A-Z and are not main titles
    is_sub_item = (
        (~is_main_title) & (~starts_with_az) & (no_col.astype(str).str.strip() != "")
    )

    # Assign main title numbers
    # cumsum of is_main_title gives us the title count at each position
    title_cumsum = is_main_title.cumsum()
    # For main titles: count + (cumsum - 1) * step = 10, 20, 30, ...
    systems.loc[is_main_title, "NO"] = count + (title_cumsum[is_main_title] - 1) * step

    # Assign sub-item numbers within each title group
    # Group by the cumulative title count to get sub-items under each title
    if is_sub_item.any():
        # Create group ID based on which title each row belongs to
        group_id = title_cumsum
        # Within each group, count sub-items
        sub_item_count = (
            systems[is_sub_item].groupby(group_id[is_sub_item]).cumcount() + 1
        )
        systems.loc[is_sub_item, "NO"] = "⠠" + sub_item_count.astype(str)

    # Now is the matter of writing to the required sheets
    for system in system_names:
        sheet = wb.sheets[system]
        system_data = systems[systems["System"] == system]
        sheet.range("A2").options(index=False).value = system_data["NO"]


def prepare_to_print_technical(wb):
    """Takes a work book, set horizantal borders at pagebreaks."""
    current_sheet = wb.sheets.active
    page_setup(wb)
    for sheet in wb.sheet_names:
        if not should_skip_sheet(sheet):
            ws = wb.sheets[sheet]
            last_row = ws.range("C1500").end("up").row
            ws.activate()
            ws.range("C:C").column_width = 60
            ws.range("C:C").wrap_text = True
            _set_wrap_row_heights(ws, col_width=60)
            ws.range("D:F").autofit()
            # Adjust the last two rows so that unwanted pagebreak can be prevented
            ws.range(f"{last_row+1}:{last_row+1}").delete()
            ws.range(f"{last_row+1}:{last_row+1}").row_height = 2
            try:
                apply_conditional_format(ws)
            except Exception:
                run_macro("conditional_format")
            apply_remove_h_borders(ws)
            apply_teal_border(ws, "A", "left")
            apply_teal_border(ws, "F", "right")
            ws.activate()  # pagebreak_borders VBA needs active sheet
            run_macro("pagebreak_borders")
    wb.sheets[current_sheet].activate()


def _missing_proposal_fields(wb):
    """Return list of required Config field labels that are blank or '-'."""
    config = wb.sheets["Config"]
    cfg   = config.range("B21:B32").options(ndim=1).value or []
    right = config.range("A28:B35").options(ndim=2).value or []

    def _blank(v):
        return v is None or str(v).strip() in ("", "-")

    def _right_val(*keys):
        for row in right:
            if row[0] and str(row[0]).strip().lower().rstrip(": ") in keys:
                return row[1]
        return None

    required = [
        ("Attention to:",  cfg[0]  if len(cfg) > 0  else None),
        ("Customer:",      cfg[2]  if len(cfg) > 2  else None),
        ("Project Name:",  cfg[5]  if len(cfg) > 5  else None),
        ("Sales Manager:", _right_val("sales manager")),
        ("Jason Ref:",     _right_val("jason ref", "jason ref num")),
        ("Revision Num:",  _right_val("revision num")),
        ("Date:",          cfg[11] if len(cfg) > 11 else None),
    ]
    return [lbl for lbl, val in required if _blank(val)]


def technical(wb, show_pdf=True):
    app = wb.app
    directory, is_cloud = get_workbook_directory(wb)
    src_path = Path(directory) / wb.name
    # Check if Technical PDF already exist
    temp_file_name = Path(directory, "Technical " + wb.name[:-4] + "pdf")
    if temp_file_name.is_file():
        xw.apps.active.alert(  # type: ignore
            "The Technical PDF file already exists!\n Please delete the file and try again."
        )
        return

    missing = _missing_proposal_fields(wb)
    if missing:
        xw.apps.active.alert(  # type: ignore
            "Cannot generate proposal — the following required fields are empty in Config:\n\n"
            + "\n".join(f"  • {lbl}" for lbl in missing)
        )
        return

    wb.sheets["Cover"].range("D39").value = "TECHNICAL PROPOSAL"
    wb.sheets["Cover"].range("D40").value = wb.sheets["Cover"].range("D40").value

    wb.sheets["Summary"].range("D20:D100").value = ""
    wb.sheets["Summary"].range("C20:C100").value = (
        wb.sheets["Summary"].range("C20:C100").raw_value
    )

    if wb.name[:9] == "Technical":
        xw.apps.active.alert("The file already seems to be technical.")  # type: ignore
        return

    _stage = "Preparing technical proposal..."
    app.status_bar = _stage

    def _restore():
        app.status_bar = _stage

    if wb.name[:10] == "Commercial":
        for sheet in wb.sheet_names:
            ws = wb.sheets[sheet]
            wb.sheets[2].activate()
            if not should_skip_sheet(sheet):
                # Require to remove h_borders as these willl not be detected
                # when columns are removed and page setup changed.
                apply_remove_h_borders(ws)
                last_row = ws.range("C1500").end("up").row
                ws.range("F:G").delete()
                ws.range("AL3:AL" + str(last_row)).value = ws.range(
                    "AL3:AL" + str(last_row)
                ).raw_value
                # To reduce visual clutter
                ws.range(f"AM1:AM{last_row}").value = ws.range(
                    f"AJ1:AJ{last_row}"
                ).raw_value
                ws.range("AJ:AJ").delete()
                ws.range("AL:AL").column_width = 0
        if "T&C" in wb.sheet_names:
            wb.sheets["T&C"].delete()
        delete_scratch_sheet(wb)
        prepare_to_print_technical(wb)
        wb.sheets["Summary"].activate()
        file_name = "Technical " + wb.name[11:-4] + "xlsx"
        full_path = Path(directory, file_name)
        app.status_bar = "Saving technical proposal..."
        save_workbook_safe(wb, full_path, password="")
        pdf_path = full_path.with_suffix(".pdf")
        app.status_bar = "Generating PDF..."
        print_technical(wb, pdf_path=str(pdf_path), show_pdf=show_pdf)
        app.status_bar = "Reopening source..."
        app.calculation = "automatic"
        _src_wb = _find_or_open_workbook(app, src_path)
        wb.close()
        if _src_wb is not None:
            try:
                _src_wb.activate()
            except Exception:
                pass
        elif not src_path.exists():
            xw.apps.active.alert(f"Proposal generated but could not reopen:\n{src_path.name}")  # type: ignore
    else:
        wb.sheets["Cover"].range("C42:C47").value = (
            wb.sheets["Cover"].range("C42:C47").raw_value
        )
        wb.sheets["Cover"].range("D6:D8").value = (
            wb.sheets["Cover"].range("D6:D8").raw_value
        )
        wb.sheets["Summary"].range("G:S").delete()
        for sheet in wb.sheet_names:
            ws = wb.sheets[sheet]
            ws.range("A1").value = ws.range("A1").raw_value  # Remove formula
            ws.range("A1").wrap_text = False
            if not should_skip_sheet(sheet):
                last_row = ws.range("C1500").end("up").row
                ws.range("B3:B" + str(last_row)).value = ws.range(
                    "B3:B" + str(last_row)
                ).raw_value
                ws.range("AL3:AL" + str(last_row)).value = ws.range(
                    "AL3:AL" + str(last_row)
                ).raw_value
                ws.range("AM:BD").delete()
                ws.range("I:AK").delete()
                ws.range("F:G").delete()
                # To reduce visual clutter
                ws.range(f"AM1:AM{last_row}").value = ws.range(
                    f"G1:G{last_row}"
                ).raw_value
                ws.range("G:G").delete()
                ws.range("AL:AL").column_width = 0
        wb.sheets["Config"].delete()
        tn_sheet = get_sheet(wb, "Technical_Notes", required=False)
        if tn_sheet:
            tn_sheet.range("F:I").delete()

        # If T&C does not exist, do nothing.
        try:
            wb.sheets["T&C"].delete()
        except Exception:
            pass
        delete_scratch_sheet(wb)
        prepare_to_print_technical(wb)
        file_name = "Technical " + wb.name[:-4] + "xlsx"
        full_path = Path(directory, file_name)
        app.status_bar = "Saving technical proposal..."
        save_workbook_safe(wb, full_path, password="")
        pdf_path = full_path.with_suffix(".pdf")
        app.status_bar = "Generating PDF..."
        print_technical(wb, pdf_path=str(pdf_path), show_pdf=show_pdf)
        app.status_bar = "Reopening source..."
        app.calculation = "automatic"
        _src_wb = _find_or_open_workbook(app, src_path)
        wb.close()
        if _src_wb is not None:
            try:
                _src_wb.activate()
            except Exception:
                pass
        elif not src_path.exists():
            xw.apps.active.alert(f"Proposal generated but could not reopen:\n{src_path.name}")  # type: ignore


def prepare_to_print_commercial(wb):
    """Apply print formatting and VBA border/CF to each commercial sheet (second pass)."""
    current_sheet = wb.sheets.active
    page_setup(wb)
    for sheet in wb.sheet_names:
        if not should_skip_sheet(sheet):
            ws = wb.sheets[sheet]
            ws.activate()
            ws.range("A:A").column_width = 4
            ws.range("B:B").autofit()
            ws.range("C:C").column_width = 55
            ws.range("C:C").wrap_text = True
            _set_wrap_row_heights(ws)
            ws.range("D:H").autofit()
            try:
                apply_conditional_format(ws)
            except Exception:
                ws.activate()
                run_macro("conditional_format")
            apply_remove_h_borders(ws)
            apply_teal_border(ws, "A", "left")
            apply_teal_border(ws, "H", "right")
            ws.activate()
            run_macro("pagebreak_borders")
    wb.sheets[current_sheet].activate()


def _commercial_prepare_sheet_py(wb, sheet_name):
    """Python fallback for the commercial_prepare_sheet VBA macro.
    Used when PERSONAL.XLSB is older than the macro's introduction.
    Mirrors the VBA logic exactly: freeze A:H values, rebuild G formula,
    delete extra columns, and relocate the AL row-type labels.
    """
    ws = wb.sheets[sheet_name]
    last_row = ws.range("G1500").end("up").row
    if last_row < 3:
        return
    ws.range(f"A3:H{last_row}").value = ws.range(f"A3:H{last_row}").raw_value
    ws.range(f"G3:G{last_row - 1}").formula = (
        '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED"), D3*F3,"")'
    )
    ws.range(f"G{last_row}").formula = f"=SUM(G3:G{last_row - 1})"
    ws.range("AM:BD").delete()
    ws.range("I:AK").delete()
    ws = wb.sheets[sheet_name]  # Refresh reference after column deletions
    col_i_vals = ws.range(f"I1:I{last_row}").options(ndim=1).value
    ws.range("I:I").delete()
    if col_i_vals:
        ws.range(f"AL1:AL{last_row}").value = [[v] for v in col_i_vals]
    ws.range("AL:AL").column_width = 0


def commercial(wb, show_pdf=True):
    app = wb.app
    directory, is_cloud = get_workbook_directory(wb)
    src_path = Path(directory) / wb.name
    # Check if Commercial PDF already exists
    temp_file_name = Path(directory, "Commercial " + wb.name[:-4] + "pdf")
    if temp_file_name.is_file():
        xw.apps.active.alert(  # type: ignore
            "The Commercial PDF file already exists!\n Please delete the file and try again."
        )
        return

    missing = _missing_proposal_fields(wb)
    if missing:
        xw.apps.active.alert(  # type: ignore
            "Cannot generate proposal — the following required fields are empty in Config:\n\n"
            + "\n".join(f"  • {lbl}" for lbl in missing)
        )
        return

    app.status_bar = "Preparing commercial proposal..."
    wb.sheets["Cover"].range("D6:D8").value = (
        wb.sheets["Cover"].range("D6:D8").raw_value
    )
    wb.sheets["Cover"].range("D39").value = wb.sheets["Config"].range("B13").value
    wb.sheets["Cover"].range("D40").value = wb.sheets["Config"].range("B14").value
    wb.sheets["Cover"].range("C42:C47").value = (
        wb.sheets["Cover"].range("C42:C47").raw_value
    )
    wb.sheets["Summary"].range("C20:C100").value = (
        wb.sheets["Summary"].range("C20:C100").raw_value
    )
    for sheet in wb.sheet_names:
        ws = wb.sheets[sheet]
        ws.range("A1").value = ws.range("A1").raw_value  # Remove formula
        ws.range("A1").wrap_text = False
        if not should_skip_sheet(sheet):
            try:
                get_macro_nb().macro("commercial_prepare_sheet")(sheet)
            except Exception:
                # Fallback: Python equivalent of the VBA macro (older PERSONAL.XLSB)
                _commercial_prepare_sheet_py(wb, sheet)
    app.status_bar = "Formatting for print..."
    prepare_to_print_commercial(wb)

    wb.sheets["Summary"].range("G:X").delete()
    wb.sheets["Config"].delete()
    delete_scratch_sheet(wb)
    tn_sheet = get_sheet(wb, "Technical_Notes", required=False)
    if tn_sheet:
        tn_sheet.range("F:I").delete()
    wb.sheets["Summary"].activate()
    file_name = "Commercial " + wb.name[:-4] + "xlsx"
    full_path = Path(directory, file_name)
    app.status_bar = "Saving commercial proposal..."
    save_workbook_safe(wb, full_path, password="")
    pdf_path = full_path.with_suffix(".pdf")
    app.status_bar = "Generating PDF..."
    try:
        to_pdf_safe(wb, pdf_path, show=show_pdf)
    except Exception as e:
        xw.apps.active.alert(  # type: ignore
            f"This error is encountered {e}. The PDF file already exists?"
        )
    app.status_bar = "Reopening source..."
    app.calculation = "automatic"  # Restore before reopening so source opens without stale warnings
    _src_wb = _find_or_open_workbook(app, src_path)
    wb.close()
    if _src_wb is not None:
        try:
            _src_wb.activate()
        except Exception:
            pass
    elif not src_path.exists():
        xw.apps.active.alert(f"Proposal generated but could not reopen:\n{src_path.name}")  # type: ignore


def prepare_to_print_internal(wb):
    """Takes a work book, set horizantal borders at pagebreaks."""
    current_sheet = wb.sheets.active
    page_setup(wb)
    for sheet in wb.sheet_names:
        if not should_skip_sheet(sheet):
            ws = wb.sheets[sheet]
            ws.activate()
            run_macro("conditional_format_internal_costing")  # Phase 2: replace with Python
            apply_remove_h_borders(ws)
            # Below is commented out so that blue lines do not show
            # MACRO_NB.macro('pagebreak_borders')()
    wb.sheets[current_sheet].activate()


def print_technical(wb, pdf_path=None, show_pdf=True):
    """The technical proposal will be written to the specified path or cwd."""
    try:
        if pdf_path:
            to_pdf_safe(wb, Path(pdf_path), show=show_pdf)
        else:
            wb.to_pdf(show=show_pdf)
    except Exception:
        # The program does not override the existing file. The file needs to be removed if it exists.
        xw.apps.active.alert(  # type: ignore
            "The PDF file already exists!\n Please delete the file and try again."
        )


# ---------------------------------------------------------------------------
# Simple Proposal helpers
# ---------------------------------------------------------------------------

# Entity-specific company information keyed by Config!B14 / Cover!B14 dropdown value
_SIMPLE_ENTITY_DATA = {
    "Jason Energy Pte. Ltd.": {
        "name": "JASON ENERGY PTE. LTD.",
        "address": "194 Pandan Loop · #06-05 PanTech Business Hub · Singapore 128383",
        "contact": "TEL: +65 6477 7700 · FAX: +65 6872 1800 · www.jason.com.sg · Co. Reg. No. 201304398E",
    },
    "Jason Electronics (Pte) Ltd": {
        "name": "JASON ELECTRONICS (PTE) LTD",
        "address": "194 Pandan Loop · #06-04 PanTech Business Hub · Singapore 128383",
        "contact": "TEL: +65 6477 7700 · FAX: +65 6872 1800 · www.jason.com.sg · Co. Reg. No. 197800377K",
    },
}
_SIMPLE_ENTITY_DEFAULT = "Jason Electronics (Pte) Ltd"

# Fixed row anchors in Template_simple.xlsx (must match create_simple_template.py)
# Rows 1–5 are the repeat block (print_title_rows="1:5"):
#   1=entity name, 2=address, 3=contact, 4=blue rule, 5=spacer
_ST_ENTITY_ROW   = 1
_ST_ADDRESS_ROW  = 2
_ST_CONTACT_ROW  = 3
_ST_TYPE_ROW     = 7   # proposal type ("COMMERCIAL PROPOSAL") — data area, page 1 only
_ST_META_START   = 9   # Metadata block starts here

# Metadata field map: (template_row, label, Config cell B21–B32)
# All 12 Config fields B21–B32 are now mapped; B25/B27/B31 were previously skipped.
_ST_META_FIELDS = [
    (_ST_META_START +  0, "Attention to:",    "B21"),
    (_ST_META_START +  1, "Designation:",     "B22"),
    (_ST_META_START +  2, "Customer:",        "B23"),
    (_ST_META_START +  3, "Client Ref:",      "B24"),
    (_ST_META_START +  4, "Ref Doc No:",      "B25"),
    (_ST_META_START +  5, "Project:",         "B26"),
    (_ST_META_START +  6, "Sales:",           "B28"),
    (_ST_META_START +  7, "Jason Ref:",       "B29"),
    (_ST_META_START +  8, "Revision Num:",    "B30"),
    (_ST_META_START +  9, "Date:",            "B32"),
]
# First 6 fields go in left column (C); last 4 are replaced by dynamic read from Config A28:B35.
_ST_META_LEFT_COUNT = 6
# Maximum header/data positions (all 6 left fields filled, 1 spacer, header).
# Actual positions are computed at runtime after filtering empty fields.
_ST_TABLE_HDR  = _ST_META_START + _ST_META_LEFT_COUNT + 1   # row 16 (max)
_ST_DATA_START = _ST_TABLE_HDR + 1                           # row 17 (max)

ACCOUNTING_COMMA = "#,##0.00"
ACCOUNTING_PAREN = "#,##0.00;(#,##0.00)"   # negative shown as (111), not -111

# Simple-Proposal layout constants — change font/size here to update everywhere.
_SP_BODY_FONT = "Helvetica"  # ReportLab name; metrically equivalent to Excel's Arial
_SP_BODY_PT   = 12           # BOQ and totals body font size
_SP_TC_PT     = 10           # T&C lines font size (subordinate to BOQ)
_SP_ROW_H     = 18.0         # single-line row height for Arial 12pt; 18pt clears descenders on Mac Excel
                             # (Mac top-padding ≥ 2.5pt; 15.75 and 16.5 both still clip)
# Excel col_width → available text width (pt): avail = (col_width × _SP_MDW_PX + 1) × 0.75
# _SP_MDW_PX is the MaxDigitWidth of the workbook's Normal-style font (Calibri 11pt default).
# Mac (72 DPI effective): Calibri 11pt MDW ≈ 8px.  Windows (96 DPI): ≈ 7px.
_SP_MDW_PX    = 7.0 if sys.platform == "win32" else 8.0


def _format_iso_date(val):
    """Return val as YYYY-MM-DD string if it is a date/datetime; else return as-is."""
    if val is not None and hasattr(val, "strftime"):
        return val.strftime("%Y-%m-%d")
    return val


def _get_tools_path() -> Path | None:
    """Return the local sync path of SharePoint @tools, or None if not found."""
    username = getpass.getuser()
    if username == "oliver":
        p = Path.home() / "OneDrive - Jason Electronics Pte Ltd" / "Shared Documents" / "@tools"
    elif username == "carol_lim":
        p = Path.home() / "Jason Electronics Pte Ltd" / "Bid Proposal - @tools"
    else:
        p = Path.home() / "Jason Electronics Pte Ltd" / "Bid Proposal - Documents" / "@tools"
    return p if p.exists() else None


def _find_simple_template(source_dir: str) -> Path | None:
    """
    Locate Template_simple.xlsx.

    Search order:
      1. Same directory as the source project file
      2. SharePoint @tools/resources/
    """
    local = Path(source_dir) / "Template_simple.xlsx"
    if local.exists():
        return local
    tools = _get_tools_path()
    if tools:
        shared = tools / "resources" / "Template_simple.xlsx"
        if shared.exists():
            return shared
    return None


def _sp_cell(ws, row, col_letter):
    """Return an xlwings Range for (row, col_letter) without sheet activation."""
    return ws.range(f"{col_letter}{row}")


_JASON_BLUE = (0, 91, 191)     # #005BBF — Jason Blue
_COMMENT_GREY = (127, 127, 127)  # mid-grey for comment rows


def _sp_wrap_lines(text, col_width, font=_SP_BODY_FONT, pt=_SP_BODY_PT):
    """Return the number of word-wrapped lines *text* occupies in an Excel column.

    Matches Excel's WrapText word-break behaviour using ReportLab font metrics.
    Hard newlines (\\n in the cell value) are treated as forced line breaks before
    word-wrapping is applied within each segment — matching Excel's rendering.
    Helvetica is metrically equivalent to Arial; change *font* and *pt* to match
    whatever font is actually written to the sheet.  *col_width* is the Excel
    column_width value (same units as Range.column_width).
    """
    from reportlab.pdfbase.pdfmetrics import stringWidth as _sw
    avail_pt = (col_width * _SP_MDW_PX + 1) * 0.75
    text = str(text).strip()
    if not text:
        return 1
    sp_w = _sw(" ", font, pt)
    total = 0
    for segment in text.split("\n"):
        if not segment.strip():
            total += 1   # blank line still occupies a row in Excel
            continue
        words = segment.split()
        seg_lines, cur = 1, 0.0
        for w in words:
            ww = _sw(w, font, pt)
            if cur == 0:
                cur = ww
            elif cur + sp_w + ww > avail_pt:
                seg_lines += 1
                cur = ww
            else:
                cur += sp_w + ww
        total += seg_lines
    return max(1, total)


def _sp_apply_row_fmt(ws, row, fmt_type, mode, desc=None):
    """Apply per-row formatting based on AL format type."""
    row_range = ws.range(f"{row}:{row}")
    if fmt_type == "System":
        row_range.font.bold = True
        row_range.font.italic = False
        row_range.font.color = _JASON_BLUE
    elif fmt_type == "Subsystem":
        row_range.font.bold = True
        row_range.font.italic = False
        row_range.font.color = (0, 0, 0)
    elif fmt_type == "Title":
        row_range.font.bold = True
        row_range.font.italic = False
        row_range.font.color = (0, 0, 0)
    elif fmt_type == "Subtitle":
        row_range.font.italic = True
        if sys.platform == "win32":
            row_range.api.Font.Underline = 2  # xlUnderlineStyleSingle
    elif fmt_type == "Comment":
        row_range.font.italic = True
        if desc and str(desc).startswith("***"):
            row_range.font.color = _JASON_BLUE
        else:
            row_range.font.color = _COMMENT_GREY


def _sp_write_column_header(ps, hdr_row, mode, currency, has_scope=True):
    """Write and style the BOQ column header row (blue fill, white bold text)."""
    labels = [
        "No.", "SN", "Description", "Qty", "Unit",
        f"Unit Price ({currency})" if mode == "commercial" else "",
        f"Total ({currency})" if mode == "commercial" else "",
        "Scope" if has_scope else "",
    ]
    rng = ps.range(f"A{hdr_row}:H{hdr_row}")
    rng.value = [labels]
    rng.color = _JASON_BLUE
    rng.font.color = (255, 255, 255)
    rng.font.bold = True
    rng.font.name = "Aptos"
    rng.font.size = 9
    rng.row_height = 17
    rng.vertical_alignment = "center"   # Windows; Mac relies on template pre-styling
    # Right-align Unit Price and Total headers to match the numbers below them.
    # Mac relies on template pre-styling (row 16 excluded from left override).
    if sys.platform == "win32":
        try:
            ps.range(f"F{hdr_row}").api.HorizontalAlignment = -4152  # xlRight
            ps.range(f"G{hdr_row}").api.HorizontalAlignment = -4152
        except Exception:
            pass


def simple_proposal(wb, mode="commercial", show_pdf=True):
    """
    Generate a compact single-page proposal using Template_simple.xlsx.

    Reads from the source workbook without modifying it.  Fills a copy of
    Template_simple.xlsx with entity header, metadata, BOQ, and T&C, then
    exports XLSX + PDF to the same directory as the source file.

    mode: 'commercial' (with pricing, auto-detects discount) or 'technical'
    Output: "Commercial [name].xlsx/.pdf" or "Technical [name].xlsx/.pdf"
    """
    directory, is_cloud = get_workbook_directory(wb)

    prefix = "Technical " if mode == "technical" else "Commercial "
    base_name = wb.name[:-5] if wb.name.endswith(".xlsx") else wb.name[:-4]
    output_xlsx = Path(directory) / f"{prefix}{base_name}.xlsx"
    output_pdf  = output_xlsx.with_suffix(".pdf")

    existing = [f.name for f in (output_xlsx, output_pdf) if f.is_file()]
    if existing:
        xw.apps.active.alert(  # type: ignore
            "Output file(s) already exist — please delete before regenerating:\n\n"
            + "\n".join(f"  • {name}" for name in existing)
        )
        return

    system_sheets = [s for s in wb.sheet_names if not should_skip_sheet(s)]
    if not (1 <= len(system_sheets) <= 2):
        xw.apps.active.alert(  # type: ignore
            f"Simple Proposal supports one or two system sheets. "
            f"Found {len(system_sheets)}: {', '.join(system_sheets) or 'none'}.\n"
            "Use Commercial Proposal or Technical Proposal for workbooks with more sheets."
        )
        return

    # Locate Template_simple.xlsx
    tmpl_path = _find_simple_template(directory)
    if tmpl_path is None:
        xw.apps.active.alert(  # type: ignore
            "Template_simple.xlsx not found.\n"
            "Place it in the same folder as the project file, "
            "or in SharePoint @tools/resources/."
        )
        return

    wb.app.calculate()

    # -----------------------------------------------------------------------
    # Read source data
    # -----------------------------------------------------------------------
    config  = wb.sheets["Config"]

    currency = config.range("B12").value or "SGD"
    if mode == "technical":
        proposal_title = "TECHNICAL PROPOSAL"
    else:
        proposal_title = config.range("B13").value or "COMMERCIAL PROPOSAL"

    # Entity: Config!B14 dropdown
    try:
        entity_key = config.range("B14").value or _SIMPLE_ENTITY_DEFAULT
    except Exception:
        entity_key = _SIMPLE_ENTITY_DEFAULT
    entity_info = _SIMPLE_ENTITY_DATA.get(entity_key, _SIMPLE_ENTITY_DATA[_SIMPLE_ENTITY_DEFAULT])

    # Collect BOQ data and font colors per sheet.
    # Maps (row_idx, col_offset) -> (R, G, B); col_offset 0=B,1=C,2=D,3=E.
    # Optimisation: check column C first; only read B/D/E for rows where C is colored.
    _SRC_COLOR_OUT_COLS = ["B", "C", "D", "E"]
    _special_fmts_color = {"System", "Subsystem", "Title", "Subtitle", "Comment"}

    def _xlw_to_rgb(v):
        if v is None:
            return None
        if isinstance(v, (tuple, list)) and len(v) == 3:
            r, g, b = int(v[0]), int(v[1]), int(v[2])
        else:
            try:
                n = int(v)
                if n == 0:
                    return None
                r, g, b = (n >> 16) & 0xFF, (n >> 8) & 0xFF, n & 0xFF
            except Exception:
                return None
        return None if (r, g, b) in ((0, 0, 0), (255, 255, 255)) else (r, g, b)

    # Strikethrough is not exposed by xlwings Font on Mac, so we read it via
    # openpyxl directly from the source file. Load once before the loop.
    # Fails gracefully for OneDrive placeholder files — strike just won't carry over.
    _oxl_wb_strike = None
    _src_file_path = _resolve_workbook_path(wb)
    if _src_file_path is not None:
        try:
            import openpyxl as _oxl
            _oxl_wb_strike = _oxl.load_workbook(str(_src_file_path), data_only=True)
        except Exception:
            pass

    sheets_data = []   # [(rows_ah, al_vals, sheet_colors, sheet_strike), ...]
    for _sname in system_sheets:
        _src_ws = wb.sheets[_sname]
        _last_row = max(_src_ws.range("C1500").end("up").row, _src_ws.range("G1500").end("up").row)
        _rows_ah = _src_ws.range(f"A3:H{_last_row}").options(ndim=2).value or []
        _al_vals = _src_ws.range(f"AL3:AL{_last_row}").options(ndim=1).value or []

        # Font colors via xlwings (works on Mac and Windows, including OneDrive files)
        _sheet_colors = {}
        try:
            for _ri, _row_data in enumerate(_rows_ah):
                _no, _desc = _row_data[0], _row_data[2]
                if not _desc and not _no:
                    continue
                _al = _al_vals[_ri] if _al_vals else None
                _no_has_val = _no is not None and (
                    (isinstance(_no, str) and _no.strip()) or
                    (isinstance(_no, (int, float)) and _no)
                )
                if _al in ("Comment", "Subtitle") and _no_has_val:
                    _al = "Title"
                if _al in _special_fmts_color:
                    continue
                _c_rgb = _xlw_to_rgb(_src_ws.range(f"C{_ri + 3}").font.color)
                if _c_rgb is None:
                    continue  # C is default — skip B/D/E too (saves calls per row)
                _sheet_colors[(_ri, 1)] = _c_rgb
                for _ci, _ltr in [(0, "B"), (2, "D"), (3, "E")]:
                    _rgb = _xlw_to_rgb(_src_ws.range(f"{_ltr}{_ri + 3}").font.color)
                    if _rgb:
                        _sheet_colors[(_ri, _ci)] = _rgb
        except Exception:
            pass

        # Strikethrough via openpyxl (xlwings Font doesn't expose strikethrough on Mac)
        _sheet_strike = {}
        if _oxl_wb_strike is not None:
            try:
                _oxl_ws = _oxl_wb_strike[_sname]
                for _ri in range(len(_rows_ah)):
                    for _ci, _col_num in enumerate([2, 3, 4, 5]):  # B–E
                        if _oxl_ws.cell(row=_ri + 3, column=_col_num).font.strikethrough:
                            _sheet_strike[(_ri, _ci)] = True
            except Exception:
                pass

        sheets_data.append((_rows_ah, _al_vals, _sheet_colors, _sheet_strike))

    if _oxl_wb_strike is not None:
        try:
            _oxl_wb_strike.close()
        except Exception:
            pass

    # T&C lines: column B = letter (A, B, C…), column C = text; starts at row 5
    tc_lines = []
    tc_sheet = get_sheet(wb, "T&C", required=False)
    if tc_sheet:
        tc_last = tc_sheet.range("C1500").end("up").row
        if tc_last >= 5:
            bc_raw = tc_sheet.range(f"B5:C{tc_last}").options(ndim=2).value
            for row_bc in bc_raw:
                letter, text = row_bc[0], row_bc[1]
                if text:
                    tc_lines.append(f"({letter})  {text}" if letter else text)

    # Auto-detect discount from Summary (commercial mode only)
    discount_amount = None
    has_discount = False
    if mode == "commercial" and "Summary" in wb.sheet_names:
        disc_row = len(system_sheets) + 19 + 3
        label_cell = wb.sheets["Summary"].range(f"C{disc_row}").value
        if label_cell in ("SPECIAL DISCOUNT", "SPECIAL PROJECT DISCOUNT"):
            val = wb.sheets["Summary"].range(f"D{disc_row}").value
            if val is not None and val != 0:
                discount_amount = val
                has_discount = True

    has_scope = any(
        row_data[7] not in (None, "")
        for rows_ah, al_vals, src_colors, src_strike in sheets_data
        for row_data in rows_ah
    )

    # -----------------------------------------------------------------------
    # Validate required metadata before doing any file I/O
    # -----------------------------------------------------------------------
    def _is_blank(v):
        return v is None or str(v).strip() in ("", "-")

    _cfg_b21_b32 = config.range("B21:B32").options(ndim=1).value or []
    _right_check  = config.range("A28:B35").options(ndim=2).value or []

    def _right_val(target_keys):
        for _row in _right_check:
            if _row[0] and str(_row[0]).strip().lower().rstrip(": ") in target_keys:
                return _row[1]
        return None

    _required = [
        ("Attention to:",  _cfg_b21_b32[0]  if len(_cfg_b21_b32) > 0  else None),
        ("Customer:",      _cfg_b21_b32[2]  if len(_cfg_b21_b32) > 2  else None),
        ("Sales Manager:", _right_val({"sales manager"})),
        ("Jason Ref:",     _right_val({"jason ref", "jason ref num"})),
        ("Revision Num:",  _right_val({"revision num"})),
        ("Date:",          _cfg_b21_b32[11] if len(_cfg_b21_b32) > 11 else None),
    ]
    _missing = [lbl for lbl, val in _required if _is_blank(val)]
    if _missing:
        xw.apps.active.alert(  # type: ignore
            "Cannot generate proposal — the following required fields are empty in Config:\n\n"
            + "\n".join(f"  • {lbl}" for lbl in _missing)
        )
        return

    # -----------------------------------------------------------------------
    # Copy template and open it in the same Excel instance
    # -----------------------------------------------------------------------
    shutil.copy2(str(tmpl_path), str(output_xlsx))

    app = wb.app
    out_wb = open_workbook_safe(app, output_xlsx)
    try:
        ps = out_wb.sheets["Proposal"]

        # Snapshot logo positions before any column-width changes. Column width
        # changes can drift cell-anchored shapes, so we restore these after all
        # widths are finalised.
        _shape_pos = {s.name: (s.left, s.top) for s in ps.shapes}

        # -------------------------------------------------------------------
        # Fill header rows 1–3: batch write (1 AppleScript call)
        # -------------------------------------------------------------------
        ps.range(f"A{_ST_ENTITY_ROW}:A{_ST_CONTACT_ROW}").value = [
            [entity_info["name"]],
            [entity_info["address"]],
            [entity_info["contact"]],
        ]

        # -------------------------------------------------------------------
        # Proposal type (row 7, data area — page 1 only)
        # -------------------------------------------------------------------
        ps.range(f"C{_ST_TYPE_ROW}").value = proposal_title

        # -------------------------------------------------------------------
        # Metadata: two-column layout.
        #   Left  (C): Attention to, Designation, Customer, Client Ref, Ref Doc No, Project Name
        #   Right (F): dynamic — reads Config A28:B35, skips blanks and "Comm Site"
        # Header row written dynamically after the taller of the two columns.
        # -------------------------------------------------------------------
        ps.range(f"A{_ST_META_START}:H{_ST_TABLE_HDR}").clear_contents()
        ps.range(f"A{_ST_TABLE_HDR}:H{_ST_TABLE_HDR}").color = None

        # Left column: hardcoded fields from Config B21–B26
        cfg_block = config.range("B21:B32").options(ndim=1).value or []
        left_active = []
        for (_, lbl, cfg_cell) in _ST_META_FIELDS[:_ST_META_LEFT_COUNT]:
            cfg_row_idx = int(cfg_cell[1:]) - 21
            raw = cfg_block[cfg_row_idx] if cfg_row_idx < len(cfg_block) else None
            if raw is None:
                continue
            val_str = str(raw).strip()
            if not val_str or val_str == "-":
                continue
            left_active.append((lbl, val_str))

        if left_active:
            left_end = _ST_META_START + len(left_active) - 1
            left_rng = ps.range(f"C{_ST_META_START}:C{left_end}")
            left_rng.value = [[f"{lbl} {val}"] for lbl, val in left_active]
            left_rng.number_format = "@"
            left_rng.wrap_text = False

        # Right column: dynamic from Config A28:B35 (label from A, value from B)
        _EXCLUDE_RIGHT = {"comn site", "comm site"}
        _RENAME_RIGHT  = {
            "sales manager": "Sales:",
            "jason ref num": "Jason Ref:",
            "jason ref":     "Jason Ref:",
            "revision num":  "Revision:",
        }
        right_raw = config.range("A28:B35").options(ndim=2).value or []
        right_active = []
        for row_ab in right_raw:
            a_lbl, b_val = row_ab[0], row_ab[1]
            if not a_lbl or b_val is None:
                continue
            lbl_str = str(a_lbl).strip()
            if not lbl_str:
                continue
            key = lbl_str.lower().rstrip(": ")
            if key in _EXCLUDE_RIGHT:
                continue
            lbl_str = _RENAME_RIGHT.get(key, lbl_str.rstrip())
            val_str = _format_iso_date(b_val) or str(b_val).strip()
            if not val_str or val_str == "-":
                continue
            right_active.append((lbl_str, val_str))

        right_meta_col = "D" if mode == "technical" else "F"
        if right_active:
            right_end = _ST_META_START + len(right_active) - 1
            right_rng = ps.range(f"{right_meta_col}{_ST_META_START}:{right_meta_col}{right_end}")
            right_rng.value = [[f"{lbl} {val}"] for lbl, val in right_active]
            right_rng.number_format = "@"
            right_rng.wrap_text = False

        # Dynamic header: 1 spacer row after the taller column, then blue header
        hdr_row = _ST_META_START + max(len(left_active), len(right_active)) + 1
        _sp_write_column_header(ps, hdr_row, mode, currency, has_scope=has_scope)
        # Narrow spacer row between column header and first BOQ row for breathing room
        ps.range(f"{hdr_row + 1}:{hdr_row + 1}").row_height = 4
        data_start = hdr_row + 2

        # Repeat entity header (rows 1–5) on every page (Windows COM only)
        if sys.platform == "win32":
            try:
                ps.api.PageSetup.PrintTitleRows = "$1:$5"
            except Exception:
                pass

        # -------------------------------------------------------------------
        # BOQ: build data in Python first, then write in one batch call.
        # Column layout: A=No B=SN C=Description D=Qty E=Unit F=UP G=Total H=Scope
        # Rows needing special font (System/Title/Subtitle/Comment) are tracked
        # separately and formatted after the bulk write.
        # -------------------------------------------------------------------
        r = data_start
        num_fmt = ACCOUNTING_PAREN

        ps.range(f"C{r}").clear_contents()   # clear template placeholder

        boq_rows = []        # 2-D list for batch write
        fmt_pending = []     # [(row, fmt_type, desc)] for rows needing font changes
        color_pending = []   # [(row, col_letter, rgb)] source colors to carry over
        strike_pending = []  # [(row, col_letter)] source strikethrough to carry over

        for sheet_idx, (rows_ah, al_vals, src_colors, src_strike) in enumerate(sheets_data):
            if sheet_idx > 0:
                boq_rows.append([None] * 8)   # one empty row between sheets
                r += 1

            for i, row_data in enumerate(rows_ah):
                no, sn, desc, qty, unit, up, sp, scope = row_data
                fmt = al_vals[i] if al_vals else None

                # Preserve intentional empty rows (spacers between items)
                if not desc and not no:
                    boq_rows.append([None] * 8)
                    r += 1
                    continue

                # Column A in source: integers = main item number; strings = sub-number (.1, .2)
                if no is None:
                    no_disp = None
                elif isinstance(no, str):
                    no_disp = no.strip() or None          # ".1", ".2", ".3" etc.
                elif isinstance(no, (int, float)) and no:
                    no_disp = str(int(no))                # 1, 2, 3 etc.
                else:
                    no_disp = None

                # Column B (SN): integer sub-item numbers within a group
                if sn is None:
                    sn_disp = None
                elif isinstance(sn, str):
                    sn_disp = sn.strip() or None
                elif isinstance(sn, (int, float)) and sn:
                    sn_disp = str(int(sn))
                else:
                    sn_disp = None

                if mode == "commercial":
                    boq_rows.append([no_disp, sn_disp, desc, qty, unit, up, sp, scope])
                else:
                    boq_rows.append([no_disp, sn_disp, desc, qty, unit, None, None, scope])

                # Numbered rows (No. column has a value) that the source marks as
                # Comment or Subtitle are section-level headers — render as Title
                # (bold black) so they look consistent in the simple proposal.
                if fmt in ("Comment", "Subtitle") and no_disp:
                    fmt = "Title"
                if fmt in ("System", "Subsystem", "Title", "Subtitle", "Comment"):
                    fmt_pending.append((r, fmt, desc))
                else:
                    if src_colors:
                        for _ci, _col_letter in enumerate(_SRC_COLOR_OUT_COLS):
                            if (i, _ci) in src_colors:
                                color_pending.append((r, _col_letter, src_colors[(i, _ci)]))
                    if src_strike:
                        for _ci, _col_letter in enumerate(_SRC_COLOR_OUT_COLS):
                            if (i, _ci) in src_strike:
                                strike_pending.append((r, _col_letter))

                r += 1

        data_end = r - 1

        if boq_rows:
            # One call writes all BOQ rows at once
            ps.range(f"A{data_start}:H{data_end}").value = boq_rows
            # One call applies price format to both price columns
            if mode == "commercial":
                ps.range(f"F{data_start}:G{data_end}").number_format = ACCOUNTING_PAREN

        # Apply font/colour only for rows that need it (skips Description/Lineitem)
        for (row_r, fmt, desc) in fmt_pending:
            _sp_apply_row_fmt(ps, row_r, fmt, mode, desc=desc)

        # -------------------------------------------------------------------
        # Totals block (commercial only)
        # Labels written to D so they sit adjacent to the amounts in G,
        # overflowing naturally through the empty E and F cells.
        # -------------------------------------------------------------------
        if mode == "commercial":
            total_row = r
            d_total = ps.range(f"D{total_row}")
            d_total.value = f"TOTAL  ({currency})"
            d_total.font.bold = True
            d_total.wrap_text = False
            ps.range(f"G{total_row}").formula = f"=SUM(G{data_start}:G{data_end})"
            ps.range(f"G{total_row}").number_format = num_fmt
            ps.range(f"G{total_row}").font.bold = True
            r += 1

            if has_discount:
                disc_row_r = r
                d_disc = ps.range(f"D{disc_row_r}")
                d_disc.value = "SPECIAL DISCOUNT"
                d_disc.wrap_text = False
                ps.range(f"G{disc_row_r}").value = discount_amount
                ps.range(f"G{disc_row_r}").number_format = ACCOUNTING_PAREN
                r += 1

                after_row = r
                d_after = ps.range(f"D{after_row}")
                d_after.value = f"AFTER DISCOUNT  ({currency})"
                d_after.font.bold = True
                d_after.wrap_text = False
                ps.range(f"G{after_row}").formula = f"=G{total_row}+G{disc_row_r}"
                ps.range(f"G{after_row}").number_format = num_fmt
                ps.range(f"G{after_row}").font.bold = True
                r += 1

        # -------------------------------------------------------------------
        # T&C section — batch write title + all lines in two range calls
        # -------------------------------------------------------------------
        if tc_lines:
            r += 1  # spacer
            tc_title_row = r
            ps.range(f"C{tc_title_row}").value = "TERMS & CONDITIONS"
            ps.range(f"C{tc_title_row}").font.bold = True
            r += 1
            tc_start = r
            ps.range(f"C{tc_start}").value = [[line] for line in tc_lines]
            tc_end = tc_start + len(tc_lines) - 1
            r = tc_end + 1

        # -------------------------------------------------------------------
        # Standardize body font: Arial 12 for BOQ/totals/T&C, Arial 10 for metadata.
        # Row 19 (column header) and rows 1–5 (entity branding) excluded.
        # Setting name/size independently preserves bold/italic/color already applied.
        # -------------------------------------------------------------------
        # Proposal type (row 7) + metadata spacer area → Arial 12
        ps.range(f"A{_ST_TYPE_ROW}:H{hdr_row - 1}").font.name = "Arial"
        ps.range(f"A{_ST_TYPE_ROW}:H{hdr_row - 1}").font.size = 12
        # BOQ + totals + T&C → Arial 12
        ps.range(f"A{data_start}:H{r - 1}").font.name = "Arial"
        ps.range(f"A{data_start}:H{r - 1}").font.size = 12
        # Metadata rows: Arial 10, not bold
        if left_active:
            _left_end = _ST_META_START + len(left_active) - 1
            ps.range(f"C{_ST_META_START}:C{_left_end}").font.name = "Arial"
            ps.range(f"C{_ST_META_START}:C{_left_end}").font.size = 10
            ps.range(f"C{_ST_META_START}:C{_left_end}").font.bold = False
        if right_active:
            _right_end = _ST_META_START + len(right_active) - 1
            _right_meta_rng = ps.range(f"{right_meta_col}{_ST_META_START}:{right_meta_col}{_right_end}")
            _right_meta_rng.font.name = "Arial"
            _right_meta_rng.font.size = 10
            _right_meta_rng.font.bold = False
        # T&C items → Arial 10 (smaller than BOQ to subordinate them)
        if tc_lines:
            ps.range(f"C{tc_start}:C{tc_end}").font.size = 10

        # Carry over source font colors and strikethrough for Description/Lineitem rows.
        # Applied after all batch font.name/size writes — setting font.name via COM
        # on Windows resets other font properties (including color) to defaults.
        for (row_r, col_letter, rgb) in color_pending:
            ps.range(f"{col_letter}{row_r}").font.color = rgb
        for (row_r, col_letter) in strike_pending:
            _rng = ps.range(f"{col_letter}{row_r}")
            try:
                if sys.platform == "win32":
                    _rng.api.Font.Strikethrough = True
                else:
                    _rng.font.api.strikethrough.set(True)
            except Exception:
                pass

        # Top-align all columns so numbers/qty/scope sit at the top of wrapped rows
        ps.range(f"A{data_start}:H{r - 1}").vertical_alignment = "top"
        # Right-align No. column (A) so sub-numbers (.1) and integers align flush right
        if sys.platform == "win32":
            try:
                ps.range(f"A{data_start}:A{r - 1}").api.HorizontalAlignment = -4152
            except Exception:
                pass

        # Reset all column widths to their final values BEFORE autofit so that row
        # heights are calculated at the correct widths. On Mac, xlwings inflates
        # columns when values are written; autofitting before resetting would
        # produce row heights based on the wrong (wider) column, leaving blank
        # space below single-line descriptions after the column is narrowed.
        ps.range("A:A").column_width = 5
        ps.range("B:B").column_width = 4
        ps.range("C:C").column_width = 55
        ps.range("D:D").column_width = 5
        ps.range("E:E").column_width = 5
        ps.range("H:H").column_width = 8
        if mode == "commercial":
            # Width based on header label with a floor that fits 7-digit totals
            # in parentheses format e.g. "(9,999,999.00)" needs ~14 units.
            f_width = max(round(len(f"Unit Price ({currency})") * 0.82), 14)
            g_width = max(round(len(f"Total ({currency})") * 0.82), 14)
            ps.range("F:F").column_width = f_width
            ps.range("G:G").column_width = g_width

        # -------------------------------------------------------------------
        # Technical mode: remove price columns and redistribute their width.
        # F(11) + G(12) = 23 units freed; given to C(+13→68) and H(+10→18)
        # so total page width stays at 105 — logos remain within print boundary.
        # -------------------------------------------------------------------
        if mode == "technical":
            ps.range("F:G").column_width = 0
            ps.range("C:C").column_width = 68
            ps.range("H:H").column_width = 18

        if data_end >= data_start:
            ps.range(f"C{data_start}:C{data_end}").wrap_text = True

        # Set row heights using ReportLab font metrics (_sp_wrap_lines) instead of
        # rows.autofit(). On Mac, autofit lags behind column-width changes made via
        # AppleScript, causing single-line rows to get 2-line height.  We derive
        # heights from the Python data already in memory — no Excel reads needed.
        _c_w = 55 if mode == "commercial" else 68
        if data_end >= data_start and boq_rows:
            ps.range(f"{data_start}:{data_end}").row_height = _SP_ROW_H
            for _ri, _brow in enumerate(boq_rows):
                _desc = _brow[2] if _brow and len(_brow) > 2 else None
                if _desc is None:
                    continue
                _lines = _sp_wrap_lines(_desc, _c_w)
                if _lines > 1:
                    ps.range(f"{data_start + _ri}:{data_start + _ri}").row_height = _SP_ROW_H * _lines

        if tc_lines:
            for _i, _tc in enumerate(tc_lines):
                if not _tc:
                    continue
                _lines = _sp_wrap_lines(str(_tc), _c_w, pt=_SP_TC_PT)
                if _lines > 1:
                    ps.range(f"{tc_start + _i}:{tc_start + _i}").row_height = _SP_ROW_H * _lines

        # Restore logo positions to their template coordinates. Autofit and
        # column-width changes may have drifted cell-anchored shapes.
        for _s in ps.shapes:
            if _s.name in _shape_pos:
                _s.left, _s.top = _shape_pos[_s.name]

        # -------------------------------------------------------------------
        # Print area and page setup
        # -------------------------------------------------------------------
        ps.page_setup.print_area = f"A1:H{r - 1}"
        ps.page_setup.fit_to_width = True
        ps.page_setup.center_horizontally = True
        # Footer: "Page X of Y" centered, Arial 10 (template also carries this)
        try:
            ps.api.PageSetup.CenterFooter = '&"Arial,Regular"&10Page &P of &N'
        except Exception:
            pass

        # -------------------------------------------------------------------
        # Save XLSX and export PDF
        # -------------------------------------------------------------------
        save_workbook_safe(out_wb, output_xlsx)
        try:
            to_pdf_safe(out_wb, output_pdf, show=show_pdf)
        except Exception as e:
            xw.apps.active.alert(f"PDF export error: {e}")  # type: ignore

    finally:
        try:
            out_wb.close()
        except Exception:
            pass


def apply_conditional_format(sheet):
    """
    Apply conditional formatting to column C (row type styles) and D:G (Title bold).
    Uses xlwings API - no sheet activation required.
    """
    xlExpression = 2
    xlUnderlineStyleSingle = 2

    # --- Column C: row-type styles ---
    col_c = sheet.range("C:C")
    col_c.api.FormatConditions.Delete()

    # Rules in reverse priority order (last added = highest priority via SetFirstPriority)
    rules = [
        ("System", {"bold": True, "color": -7137279}),
        ("Subsystem", {"bold": True, "color": -7137279}),
        ("Title", {"bold": True}),
        ("Subtitle", {"italic": True, "underline": xlUnderlineStyleSingle}),
        ("Comment", {"italic": True, "color": -52732}),
        ("Deleted", {"strikethrough": True}),
    ]

    for al_value, fmt in rules:
        formula = f'=AL1="{al_value}"'
        fc = col_c.api.FormatConditions.Add(Type=xlExpression, Formula1=formula)
        fc.SetFirstPriority()
        if fmt.get("bold"):
            fc.Font.Bold = True
        if fmt.get("italic"):
            fc.Font.Italic = True
        if fmt.get("underline"):
            fc.Font.Underline = fmt["underline"]
        if fmt.get("strikethrough"):
            fc.Font.Strikethrough = True
        if "color" in fmt:
            fc.Font.Color = fmt["color"]
        fc.StopIfTrue = True

    # --- Columns D:G: bold when Title row has a number in D ---
    col_dg = sheet.range("D:G")
    col_dg.api.FormatConditions.Delete()
    fc = col_dg.api.FormatConditions.Add(
        Type=xlExpression,
        Formula1='=AND($AL1="Title",ISNUMBER($D1))',
    )
    fc.SetFirstPriority()
    fc.Font.Bold = True
    fc.StopIfTrue = False


def apply_teal_border(sheet, col_letter, edge):
    """Apply teal edge border to a column via VBA macro (both Mac and Windows).
    VBA operates on ActiveSheet — caller must activate the sheet first.
    """
    run_macro(f"format_col_{col_letter.lower()}_{edge}_border")


def apply_remove_h_borders(sheet):
    """
    Remove horizontal inside borders from the data range.
    Windows: Python/xlwings API.  Mac: VBA (appscript bridge doesn't expose Borders).
    """
    if sys.platform == "darwin":
        sheet.activate()
        run_macro("remove_h_borders")
        return
    xlInsideHorizontal = 12
    xlNone = -4142
    last_row = sheet.range("C1500").end("up").row
    if last_row > 4:
        sheet.range(f"A3:H{last_row - 2}").api.Borders(xlInsideHorizontal).LineStyle = xlNone


def apply_format_column_border(sheet):
    """
    Apply column border formatting to the sheet.
    Windows: Python/xlwings API.  Mac: VBA (appscript bridge doesn't expose Borders).
    """
    if sys.platform == "darwin":
        sheet.activate()
        run_macro("format_column_border")
        return
    # Excel constants
    xlContinuous = 1
    xlNone = -4142
    xlThin = 2
    xlDiagonalDown = 5
    xlDiagonalUp = 6
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    xlInsideVertical = 11
    xlInsideHorizontal = 12

    COLOR_TEAL = -52732  # Dark teal color used in template

    def clear_diagonals(rng):
        rng.api.Borders(xlDiagonalDown).LineStyle = xlNone
        rng.api.Borders(xlDiagonalUp).LineStyle = xlNone

    def set_border(
        rng, edge, color=None, theme_color=None, tint: float = 0.0, weight=xlThin
    ):
        border = rng.api.Borders(edge)
        border.LineStyle = xlContinuous
        if color is not None:
            border.Color = color
            border.TintAndShade = 0
        elif theme_color is not None:
            border.ThemeColor = theme_color
            border.TintAndShade = tint
        border.Weight = weight

    def clear_border(rng, edge):
        rng.api.Borders(edge).LineStyle = xlNone

    # Column A: left=teal, right=theme4
    col_a = sheet.range("A:A")
    clear_diagonals(col_a)
    set_border(col_a, xlEdgeLeft, color=COLOR_TEAL)
    clear_border(col_a, xlEdgeTop)
    clear_border(col_a, xlEdgeBottom)
    set_border(col_a, xlEdgeRight, theme_color=4, tint=0.599993896298105)
    clear_border(col_a, xlInsideVertical)

    # Columns B-G: left and right = theme4
    for col in ["B:B", "C:C", "D:D", "E:E", "F:F", "G:G"]:
        rng = sheet.range(col)
        clear_diagonals(rng)
        set_border(rng, xlEdgeLeft, theme_color=4, tint=0.599993896298105)
        clear_border(rng, xlEdgeTop)
        clear_border(rng, xlEdgeBottom)
        set_border(rng, xlEdgeRight, theme_color=4, tint=0.599993896298105)
        clear_border(rng, xlInsideVertical)

    # Column H: left=theme4, right=teal
    col_h = sheet.range("H:H")
    clear_diagonals(col_h)
    set_border(col_h, xlEdgeLeft, theme_color=4, tint=0.599993896298105)
    clear_border(col_h, xlEdgeTop)
    clear_border(col_h, xlEdgeBottom)
    set_border(col_h, xlEdgeRight, color=COLOR_TEAL)
    clear_border(col_h, xlInsideVertical)

    # Columns I:BD: inside borders with theme3
    cols_ibd = sheet.range("I:BD")
    clear_diagonals(cols_ibd)
    clear_border(cols_ibd, xlEdgeRight)
    border_v = cols_ibd.api.Borders(xlInsideVertical)
    border_v.LineStyle = xlContinuous
    border_v.ThemeColor = 3
    border_v.TintAndShade = -0.249946592608417
    border_v.Weight = xlThin
    border_h = cols_ibd.api.Borders(xlInsideHorizontal)
    border_h.LineStyle = xlContinuous
    border_h.ThemeColor = 3
    border_h.TintAndShade = -0.249946592608417
    border_h.Weight = xlThin

    # Row 1: clear all borders
    row1 = sheet.range("1:1")
    for edge in [
        xlDiagonalDown,
        xlDiagonalUp,
        xlEdgeLeft,
        xlEdgeTop,
        xlEdgeBottom,
        xlEdgeRight,
        xlInsideVertical,
        xlInsideHorizontal,
    ]:
        clear_border(row1, edge)

    # Row 2: left, top, bottom = teal
    row2 = sheet.range("2:2")
    clear_diagonals(row2)
    set_border(row2, xlEdgeLeft, color=COLOR_TEAL)
    set_border(row2, xlEdgeTop, color=COLOR_TEAL)
    set_border(row2, xlEdgeBottom, color=COLOR_TEAL)
    clear_border(row2, xlEdgeRight)
    clear_border(row2, xlInsideHorizontal)


def conditional_format_wb(wb, app=None):
    """
    Apply conditional formatting to all sheets using Python/xlwings API.
    Falls back to VBA conditional_format macro if the API call fails (e.g. older Mac Excel).
    remove_h_borders and format_column_border are now pure Python on all platforms.

    Pass app to restore the status bar if the VBA fallback resets it.
    """
    current_sheet = wb.sheets.active
    _status = "Applying conditional formatting..."

    def _restore():
        if app:
            app.status_bar = _status

    for sheet_name in wb.sheet_names:
        if not should_skip_sheet(sheet_name):
            sheet = wb.sheets[sheet_name]

            try:
                apply_conditional_format(sheet)
            except Exception:
                sheet.activate()
                run_macro("conditional_format")
                _restore()

            apply_remove_h_borders(sheet)
            apply_format_column_border(sheet)

    current_sheet.activate()


def fix_unit_price(wb):
    """
    Fix unit prices, normally done for subsequent revisions.
    """
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheets:
        if not should_skip_sheet(sheet.name):
            system_names.append(str.upper(sheet.name))
            ws = wb.sheets[sheet]
            last_row = ws.range("C1500").end("up").row
            data = (
                ws.range("AE2:AE" + str(last_row))
                .options(pd.DataFrame, index=False)
                .value
            )
            data["System"] = str.upper(sheet.name)
            systems = pd.concat([systems, data], join="outer")

            # Set font color for FUP column AB2
            sheet.range(f"AB3:AB{str(last_row)}").font.color = (4, 50, 255)

    systems = systems.reset_index(
        drop=True
    )  # Otherwise separate sheet will have own index.
    systems.columns = ["FUP", "System"]

    # Write fixed unit price in FUP field
    for system in system_names:
        sheet = wb.sheets[system]
        system = systems[systems["System"] == system]
        sheet.range("AB2").options(index=False).value = system["FUP"]


def format_text(
    wb,
    indent_description=False,
    bullet_description=False,
    title_lineitem_or_description=False,
    upper_title=False,
    upper_system=True,
):
    """
    Format text in the workbook to remove inconsistencies.

    Optimized to use vectorized pandas operations instead of row-by-row iteration.
    """
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheets:
        if not should_skip_sheet(sheet.name):
            system_names.append(str.upper(sheet.name))
            ws = wb.sheets[sheet]
            last_row = ws.range("C1500").end("up").row
            data = (
                ws.range("C2:AL" + str(last_row))
                .options(pd.DataFrame, empty="", index=False)
                .value
            )
            data["System"] = str.upper(sheet.name)
            systems = pd.concat([systems, data], join="outer")

    systems = systems.reset_index(drop=True)
    systems = systems.reindex(
        columns=["Description", "Unit", "Scope", "Format", "System"]
    )

    # Vectorized processing of Description column
    # Apply set_nitty_gritty using vectorized apply (faster than row iteration)
    systems["Description"] = (
        systems["Description"]
        .astype(str)
        .str.strip()
        .str.lstrip("• ")
        .apply(set_nitty_gritty)
    )

    # Vectorized Unit processing
    systems["Unit"] = systems["Unit"].astype(str).str.strip().str.lower()
    # Replace "nos" and "no" with "ea"
    systems.loc[systems["Unit"].isin(["nos", "no"]), "Unit"] = "ea"
    # Remove trailing 's' (but not if it's the only character)
    mask_trailing_s = (systems["Unit"].str.len() > 1) & (systems["Unit"].str[-1] == "s")
    systems.loc[mask_trailing_s, "Unit"] = systems.loc[mask_trailing_s, "Unit"].str[:-1]

    # Vectorized Scope processing
    systems["Scope"] = systems["Scope"].astype(str).str.strip().str.lower()
    systems.loc[
        systems["Scope"].isin(["inclusive", "include", "included"]), "Scope"
    ] = "INCLUDED"
    systems.loc[systems["Scope"].isin(["option", "optional"]), "Scope"] = "OPTION"
    systems.loc[systems["Scope"] == "waived", "Scope"] = "WAIVED"
    systems.loc[systems["Scope"] == "tba", "Scope"] = "TBA"
    systems.loc[systems["Scope"] == "removed", "Scope"] = "REMOVED"

    # Apply title case to Lineitem and Description rows with short descriptions
    if title_lineitem_or_description:
        mask = systems["Format"].isin(["Lineitem", "Description"]) & (
            systems["Description"].str.len() <= 60
        )
        if mask.any():
            systems.loc[mask, "Description"] = (
                systems.loc[mask, "Description"]
                .str.strip()
                .str.lstrip("• ")
                .apply(lambda x: set_case_preserve_acronym(x, title=True))
            )

    # Upper case for Title rows
    if upper_title:
        mask = systems["Format"] == "Title"
        systems.loc[mask, "Description"] = (
            systems.loc[mask, "Description"].str.strip().str.upper()
        )

    # Upper case for System rows
    if upper_system:
        mask = systems["Format"] == "System"
        systems.loc[mask, "Description"] = (
            systems.loc[mask, "Description"].str.strip().str.upper()
        )

    # Indent and bullet Description rows
    if indent_description:
        mask = systems["Format"] == "Description"
        if mask.any():
            desc_col = systems.loc[mask, "Description"].str.strip().str.lstrip("• ")

            if bullet_description:
                # Handle # prefix -> ‣ bullet
                starts_hash = desc_col.str.startswith("#")
                # Handle ‣ prefix -> ‣ bullet
                starts_triangle = desc_col.str.startswith("‣")
                # Default -> • bullet

                result = pd.Series(index=desc_col.index, dtype=str)
                result[starts_hash] = "      ‣ " + desc_col[starts_hash].str.lstrip(
                    "# "
                )
                result[starts_triangle] = "      ‣ " + desc_col[
                    starts_triangle
                ].str.lstrip("‣ ")
                result[~starts_hash & ~starts_triangle] = (
                    "   • " + desc_col[~starts_hash & ~starts_triangle]
                )
                systems.loc[mask, "Description"] = result
            else:
                systems.loc[mask, "Description"] = "   " + desc_col

    # Write formatted description to Description field
    for system in system_names:
        sheet = wb.sheets[system]
        system_data = systems[systems["System"] == system]
        sheet.range("C2").options(index=False).value = system_data["Description"]
        sheet.range("E2").options(index=False).value = system_data["Unit"]
        sheet.range("H2").options(index=False).value = system_data["Scope"]


def indent_description(wb):
    """
    Depricated.
    Indent description
    This function works but slow. Replaced with 'format_text' function
    """
    for sheet in wb.sheets:
        if not should_skip_sheet(sheet.name):
            ws = wb.sheets[sheet]
            last_row = ws.range("C1500").end("up").row
            for format in ws.range("AL3:AL" + str(last_row)):
                if format.value == "Subtitle":
                    ws.range("C" + str(format.row)).value = str(
                        ws.range("C" + str(format.row)).value
                    ).strip()
                    ws.range("C" + str(format.row)).value = str(
                        ws.range("C" + str(format.row)).value
                    ).lstrip("• ")
                elif format.value == "Description":
                    ws.range("C" + str(format.row)).value = str(
                        ws.range("C" + str(format.row)).value
                    ).strip()
                    ws.range("C" + str(format.row)).value = str(
                        ws.range("C" + str(format.row)).value
                    ).lstrip("• ")
                    ws.range("C" + str(format.row)).value = (
                        "   • " + ws.range("C" + str(format.row)).value
                    )


def shaded(wb, shaded=True):
    """Added Shaded region"""
    # current_sheet = wb.sheets.active
    for sheet in wb.sheet_names:
        if not should_skip_sheet(sheet):
            wb.sheets[sheet].activate()
            if shaded:
                run_macro("shaded")
            else:
                run_macro("unshaded")
                # pass
    # wb.sheets[current_sheet].activate()


def internal_costing(wb):
    directory, is_cloud = get_workbook_directory(wb)

    wb.sheets["Cover"].range("D39").value = "INTERNAL COSTING"
    wb.sheets["Cover"].range("D40").value = wb.sheets["Cover"].range("D40").raw_value
    wb.sheets["Cover"].range("C42:C47").value = (
        wb.sheets["Cover"].range("C42:C47").raw_value
    )
    wb.sheets["Cover"].range("D6:D8").value = (
        wb.sheets["Cover"].range("D6:D8").raw_value
    )

    summary_last_row = wb.sheets["Summary"].range("D1500").end("up").row
    wb.sheets["Summary"].range("D20:D100").value = ""
    wb.sheets["Summary"].range("C20:C100").value = (
        wb.sheets["Summary"].range("C20:C100").raw_value
    )
    wb.sheets["Summary"].range(f"H20:H{summary_last_row}").value = (
        wb.sheets["Summary"].range(f"H20:H{summary_last_row}").raw_value
    )
    wb.sheets["Summary"].range(
        f"H{summary_last_row+1}:H{summary_last_row+50}"
    ).value = ""
    wb.sheets["Summary"].range("I:P").value = ""

    # Write out exchange rates
    wb.sheets["Summary"].range("H7:I16").value = (
        wb.sheets["Config"].range("A1:B10").raw_value
    )
    wb.sheets["Summary"].range("I8:I16").number_format = "0.0000"
    wb.sheets["Summary"].range("K7").value = "Legend"
    wb.sheets["Summary"].range("K9").value = LEGEND
    wb.sheets["Summary"].range("K:L").clear_formats()

    for sheet in wb.sheet_names:
        ws = wb.sheets[sheet]
        ws.range("A1").value = ws.range("A1").raw_value  # Remove formula
        if not should_skip_sheet(sheet):
            # Collect escalation
            escalation = ws.range("K1:R1").value
            ws.range("I1:R1").value = ""
            # Construct as dictionary
            escalation = dict(zip(escalation[::2], escalation[1::2]))

            # Work on columns
            last_row = ws.range("G1500").end("up").row
            ws.range("B3:B" + str(last_row)).value = ws.range(
                "B3:B" + str(last_row)
            ).raw_value
            ws.range("F3:G" + str(last_row)).value = ""
            # ws.range('K3:Q'+ str(last_row)).value = ws.range('K3:Q'+ str(last_row)).raw_value
            ws.range("Q3:Q" + str(last_row)).value = ws.range(
                "Q3:Q" + str(last_row)
            ).raw_value
            # Copy Flag
            ws.range(f"AK2:AK{last_row}").value = ws.range(
                f"AK2:AK{last_row}"
            ).raw_value
            ws.range("AK:AK").copy(ws.range("BB:BB"))
            ws.range("AP:AW").delete()
            ws.range("R:AK").delete()

            # Copy row first to get formatting right
            ws.range("K:K").copy(ws.range("V:V"))
            ws.range("W:AB").insert("right")
            ws.range("V:V").delete()
            # Insert Escalation
            ws.range("V2").value = "Escalation"
            ws.range("V3:V" + str(last_row - 1)).formula = (
                '=IF(AND(D3<>"", J3<>"",K3<>""), $AD$7, "")'
            )
            ws.range("V3:V" + str(last_row)).number_format = "0.00%"

            # Insert UCDQ
            ws.range("W2").value = "UCDQ"
            ws.range(f"W3:W{last_row - 1}").formula = (
                '=IF(AND(D3<>"", K3<>""), N3*Q3,"")'
            )

            # Insert SCDQ
            ws.range("X2").value = "SCDQ"
            ws.range(f"X3:X{last_row - 1}").formula = (
                '=IF(AND(D3<>"", K3<>"", H3<>"OPTION",INDEX($H$1:H2, XMATCH("Title", $R$1:R2, 0, -1))<>"OPTION"), D3*W3, "")'
            )

            # Insert SCDQL
            ws.range("Y2").value = "SCDQL"
            ws.range(f"Y3:Y{last_row - 1}").formula = (
                '=IF(AND(R3="Title", ISNUMBER(D3), E3<>""), SUM(X4:INDEX(X4:X1500, XMATCH("Title", R4:R1500, 0, 1)-1)), IF(AND(R3="Lineitem", AE3="Unit Price"), W3, ""))'
            )

            # Insert TCDQL
            ws.range("Z2").value = "TCDQL"
            ws.range(f"Z3:Z{last_row - 1}").formula = (
                '=IF(AND(ISNUMBER(D3), ISNUMBER(Y3), H3<>"OPTION"), D3*Y3, "")'
            )

            # Insert BSCQL
            ws.range("AA2").value = "BSCQL"
            ws.range(f"AA3:AA{last_row - 1}").formula = (
                '=IF(ISNUMBER(Y3), Y3*(1+$AD$7)/(1-0.05), "")'
            )

            # Insert BTCQL
            ws.range("AB2").value = "BTCQL"
            ws.range(f"AB3:AB{last_row - 1}").formula = (
                '=IF(AND(ISNUMBER(D3), ISNUMBER(AA3), H3<>"OPTION"), D3*AA3, "")'
            )

            ws.range(f"AB{last_row}").formula = "=SUM(AB3:AB" + str(last_row - 1) + ")"
            ws.range(f"W3:AB{last_row}").number_format = ACCOUNTING
            # Consolidated escalation
            ws.range("AC3").value = escalation
            ws.range("AC7").value = "Total"
            ws.range("AD7").formula = "=SUM(AD3:AD6)"
            ws.range("AD3:AD7").number_format = "0.00%"

            # To reduce visual clutter
            ws.range("D:X").autofit()
            ws.range("I:I").column_width = 20
            ws.range("P:P").column_width = 20
            ws.range("F:G").column_width = 0
            ws.range("R:R").column_width = 0
            ws.range("AE:AG").column_width = 0
    wb.sheets["Config"].delete()
    # wb.sheets['T&C'].delete()
    prepare_to_print_internal(wb)
    wb.sheets["Summary"].activate()
    file_name = "Internal " + wb.name[:-4] + "xlsx"
    save_workbook_safe(wb, Path(directory, file_name), password="")


def convert_legacy(wb):
    import requests

    directory, is_cloud = get_workbook_directory(wb)

    if wb.name[-4:] == "xlsm":
        # Read and initialize values
        # Differentiate between new and legacy template
        # visible_sheets = [sht.name for sht in wb.sheets if sht.visible]
        full_column_list = [
            "NO",
            "SN",
            "Description",
            "Qty",
            "Unit",
            "Unit Price",
            "Subtotal Price",
            "Scope",
            "Model",
            "Cur",
            "UC",
            "SC",
            "Discount",
            "UCD",
            "SCD",
            "Remark",
            "Rate",
            "UCDQ",
            "SCDQ",
            "BUCQ",
            "BSCQ",
            "Default",
            "Warranty",
            "Freight",
            "Special",
            "Risk",
            "MU",
            "FUP",
            "RUPQ",
            "RSPQ",
            "UPLS",
            "SPLS",
            "Profit",
            "Margin",
            "Auxiliary",
            "Lumpsum",
            "Flag",
            "Format",
            "Category",
            "System",
        ]
        # skip_sheets_lg = ['FX', 'Cover', 'Intro', 'ES', 'T&C']
        skip_sheets_lg = [
            "A1",
            "A2",
            "A3",
            "A4",
            "A5",
            "A6",
            "A7",
            "A8",
            "A9",
            "A10",
            "A11",
            "A12",
            "A13",
            "A14",
            "A15",
            "A16",
            "A17",
            "A18",
            "A19",
            "SUM",
            "FX",
            "Cover",
            "Intro",
            "ES",
            "T&C",
        ]
        df = pd.DataFrame(columns=full_column_list)  # type: ignore
        # risk = 0.05
        # Read and set currency from FX sheet
        fx = wb.sheets["FX"]
        exchange_rates = dict(fx.range("A2:B9").value)
        quoted_currency = fx.range("B12").value
        project_info = dict(fx.range("A36:B46").value)
        try:
            project_info = {key: value.upper() for key, value in project_info.items()}
        except Exception:
            xw.apps.active.alert("Project Info items cannot be empty value.")  # type: ignore
            return
        # Read system sheets
        cols = [
            "NO",
            "Qty",
            "Unit",
            "Description",
            "Unit Price",
            "Subtotal Price",
            "Model",
            "Cur",
            "UC",
            "SC",
            "Discount",
        ]
        systems = pd.DataFrame()
        defaults = {}
        system_names = []
        for sheet in wb.sheet_names:
            if sheet not in skip_sheets_lg:
                system_names.append(sheet.upper())
                ws = wb.sheets[sheet]
                escalation = dict(ws.range("K2:L5").value)
                default_mu = ws.range("H5").value
                escalation["default_mu"] = default_mu
                defaults[sheet.upper()] = escalation
                last_row = ws.range("D1500").end("up").row  # Returns a number
                data = (
                    ws.range("A8:K" + str(last_row))
                    .options(pd.DataFrame, index=False)
                    .value
                )
                data.columns = cols
                data["System"] = str(sheet.upper())
                data["Category"] = "Product"
                systems = pd.concat([systems, data], join="outer")
        systems = pd.concat([systems, df], join="outer")

        # Read Engineering Services
        es_cols = [
            "NO",
            "Qty",
            "Unit",
            "Description",
            "Unit Price",
            "Subtotal Price",
            "Model",
            "Cur",
            "UC",
            "SC",
            "Discount",
        ]
        es = wb.sheets["ES"]
        es_last_row = es.range("D1500").end("up").row
        eng_service = (
            es.range("A8:K" + str(es_last_row)).options(pd.DataFrame, index=False).value
        )
        eng_service.columns = es_cols
        eng_service = pd.concat([eng_service, df], join="outer")
        eng_service = eng_service.reindex(columns=full_column_list)
        eng_service["Discount"] = np.nan
        eng_service["System"] = "ENGINEERING SERVICES"
        # eng_service['Category'] = 'Service'
        systems = pd.concat([systems, eng_service], join="outer")
        systems = systems.reindex(columns=full_column_list)
        system_names.append("ENGINEERING SERVICES")

        # Set font case for some columns
        systems["Unit"] = systems["Unit"].str.lower()

        # Remove lineitem numbers
        systems = systems.reset_index(drop=True)
        for idx in systems.index:
            if str(systems.loc[idx, "NO"]).count(".") == 2:
                systems.loc[idx, "NO"] = np.nan

        for idx in systems.index:
            if pd.notna(systems.loc[idx, "NO"]) and not pd.notna(
                systems.loc[idx, "Qty"]
            ):
                systems.loc[idx, "NO"] = np.nan

        # Let's take care of the main numbering
        systems["Format"] = np.nan
        item_count = 10
        for idx in systems.index:
            if pd.notna(systems.loc[idx, "NO"]):
                try:
                    systems.at[idx, "NO"] = item_count
                    systems.at[idx, "Format"] = "Title"
                    item_count += 10
                except Exception as e:
                    print(str(e))
                    pass

        # Move Option and Included to scope
        for idx in systems.index:
            if str(systems.loc[idx, "Subtotal Price"]).lower() in [
                "option",
                "optional",
            ]:
                systems.at[idx, "Scope"] = "OPTION"
            if str(systems.loc[idx, "Subtotal Price"]).lower() in [
                "included",
                "inclusive",
            ]:
                systems.at[idx, "Scope"] = "INCLUDED"
            if str(systems.loc[idx, "Subtotal Price"]).lower() == "removed":
                systems.at[idx, "Scope"] = "REMOVED"

        # Cleaning data
        for idx in systems.index:
            # if set_nitty_gritty(str(systems.loc[idx, 'Description'])) != 'None':
            #     systems.at[idx, 'Description'] = set_nitty_gritty(str(systems.loc[idx, 'Description']))
            if (
                str(systems.loc[idx, "Model"]).lower().strip()
                == "start line:  delete forbidden"
            ):
                systems.at[idx, "Model"] = np.nan
            if (
                str(systems.loc[idx, "UC"]).lower().strip() == "true"
                or str(systems.loc[idx, "UC"]).lower().strip() == "false"
            ):
                systems.at[idx, "UC"] = np.nan
            if (
                str(systems.loc[idx, "SC"]).lower().strip() == "true"
                or str(systems.loc[idx, "SC"]).lower().strip() == "false"
            ):
                systems.at[idx, "SC"] = np.nan
            if (
                str(systems.loc[idx, "Model"]).lower().strip() == "true"
                or str(systems.loc[idx, "Model"]).lower().strip() == "false"
            ):
                systems.at[idx, "Model"] = np.nan

        # Previoulsy using Proposal_Template.xlsx
        # url = "https://filedn.com/liTeg81ShEXugARC7cg981h/Proposal_Template.xlsx"
        # Now using Template.xlsx
        url = "https://filedn.com/liTeg81ShEXugARC7cg981h/Template.xlsx"
        resp = requests.get(url)

        with open(Path(directory, "Template.xlsx"), "wb") as fd:
            for chunk in resp.iter_content(chunk_size=8192):
                fd.write(chunk)

        # Copy sheet from template to new workbook
        nb = xw.Book()
        # xl_app = xw.App(visible=False)
        # template = xl_app.books.open(Path(directory, "Template.xlsx"), password=hide.legacy)
        template = xw.Book(Path(directory, "Template.xlsx"), password=hide.legacy)
        template.sheets["Config"].copy(after=nb.sheets[0])
        nb.sheets["Sheet1"].delete()
        template.sheets["Cover"].copy(after=nb.sheets["config"])

        # Set date in Config
        nb.sheets["Config"].range("B32").value = datetime.today().strftime("%Y-%m-%d")

        # Set up formula in Cover sheet
        nb.sheets["Cover"].range("D7").formula = "=Config!B26"
        nb.sheets["Cover"].range("C42").formula = "=Config!B21"
        nb.sheets["Cover"].range("C43").formula = "=Config!B23"
        nb.sheets["Cover"].range("C44").formula = "=Config!B24"
        nb.sheets["Cover"].range("C45").formula = "=Config!B29"
        nb.sheets["Cover"].range("C46").formula = "=Config!B30"
        nb.sheets["Cover"].range("C47").formula = "=Config!B32"
        nb.sheets["Cover"].range("D39").formula = "=Config!B13"

        for system in system_names[::-1]:
            sheet_name = "Cover"
            template.sheets["System"].copy(after=nb.sheets[sheet_name])
            sheet_name = system
            nb.sheets["System"].name = sheet_name
            # Set formula to reference Config.
            # nb.sheets[sheet_name].range('C1').formula = '=Config!B29'
            # nb.sheets[sheet_name].range('C2').formula = '=Config!B30'
            # nb.sheets[sheet_name].range('C3').formula = '=Config!B32'
            # nb.sheets[sheet_name].range('C4').formula = '=Config!B26'
            nb.sheets[sheet_name].range(
                "A1"
            ).formula = '= "JASON REF: " & Config!B29 &  ", REVISION: " &  Config!B30 & ", PROJECT: " & Config!B26'
        template.sheets["Summary"].copy(after=nb.sheets["Cover"])
        template.sheets["Technical_Notes"].copy(after=nb.sheets[-1])
        template.sheets["T&C"].copy(after=nb.sheets[-1])
        for sheet in nb.sheet_names:
            if sheet in ["Summary", "Technical_Notes", "TN", "T&C"]:
                # nb.sheets[sheet].range('C1').formula = '=Config!B29'
                # nb.sheets[sheet].range('C2').formula = '=Config!B30'
                # nb.sheets[sheet].range('C3').formula = '=Config!B32'
                # nb.sheets[sheet].range('C4').formula = '=Config!B26'
                nb.sheets[sheet].range(
                    "A1"
                ).formula = '= "JASON REF: " & Config!B29 &  ", REVISION: " &  Config!B30 & ", PROJECT: " & Config!B26'
        template.close()
        os.remove(Path(directory, "Template.xlsx"))

        # Write data to sheet
        for system in system_names:
            sheet = nb.sheets[system]
            system = systems[systems["System"] == system]
            sheet.range("A2").options(index=False).value = system

        # Set exchange rates
        sheet = nb.sheets["Config"]
        exchange = pd.DataFrame([exchange_rates])
        exchange = exchange.T
        sheet.range("A2").value = exchange

        # Quoted currency
        sheet.range("B12").value = quoted_currency

        # Project info
        sheet.range("B21").value = project_info["Attend to: "]
        sheet.range("B22").value = project_info["Designation: "]
        sheet.range("B23").value = project_info["Client Name: "]
        sheet.range("B24").value = project_info["Client RFQ No: "]
        sheet.range("B25").value = project_info["Ref Doc No: "]
        sheet.range("B26").value = project_info["Project Name: "]
        sheet.range("B27").value = project_info["Prepared By: "]
        sheet.range("B28").value = project_info["Sales Manager: "]
        sheet.range("B29").value = project_info["Jason Ref: "]
        sheet.range("B30").value = project_info["Revision Num: "]
        # sheet.range('B31').value = project_info['']
        sheet.range("B32").value = datetime.today().strftime("%Y-%m-%d")

        # Write necessary formula to excel
        for system in system_names:
            sheet = nb.sheets[system]
            # Set default values
            if system != "ENGINEERING SERVICES":
                sheet.range("J1").value = defaults[system]["default_mu"]
                sheet.range("L1").value = defaults[system]["Default"]
                sheet.range("N1").value = defaults[system]["Warranty"]
                sheet.range("P1").value = defaults[system]["Inbound Freight"]
                sheet.range("R1").value = defaults[system]["Special Terms"]
                sheet.range("AL3").value = "System"
            else:
                sheet.range("J1").value = 0.3
                sheet.range("L1").value = 0
                sheet.range("N1").value = 0
                sheet.range("P1").value = 0
                sheet.range("R1").value = 0
                sheet.range("AL3").value = "System"

            # fill_formula(sheet)

        # Setup print area
        for system in system_names:
            sheet = nb.sheets[system]
            unhide_columns(sheet)
            last_row = sheet.range("G1500").end("up").row
            sheet.range("AL" + str(last_row)).value = "Title"
            sheet.page_setup.print_area = "A1:H" + str(last_row)

        fill_formula_wb(nb)
        nb.sheets["Summary"].activate()
        format_text(nb, title_lineitem_or_description=True, upper_system=True)
        format_text(nb, indent_description=True, bullet_description=True)
        conditional_format_wb(nb)
        fill_lastrow(nb)
        unhide_columns_wb(nb)
        summary(nb)
        page_setup(nb)

        file_name = wb.name[:-4] + "xlsx"
        try:
            nb.save(Path(directory, file_name), password=hide.legacy)
        except Exception:
            xw.apps.active.alert("The file already exists. Please save manually.")  # type: ignore

    else:
        xw.apps.active.alert("The excel file does not seem to be legacy template.")  # type: ignore


def page_setup(wb):
    for sheet in wb.sheets:
        sheet.page_setup.center_horizontally = True
        sheet.page_setup.center_vertically = True
        sheet.page_setup.left_margin = 0.7  # in inches
        sheet.page_setup.right_margin = 0.7  # in inches
        sheet.page_setup.top_margin = 0.75  # in inches
        sheet.page_setup.bottom_margin = 0.75  # in inches
        sheet.page_setup.header_margin = 0.3  # in inches
        sheet.page_setup.footer_margin = 0.3  # in inches
        sheet.page_setup.fit_to_width = True
        if sheet.name in ["Technical_Notes", "TN", "T&C"]:
            sheet.range("A:A").column_width = 2
            sheet.range("B:B").autofit()
            sheet.range("C:C").column_width = 70
            sheet.range("C:C").rows.autofit()
            sheet.range("C:C").wrap_text = True


def fill_formula_active_row(wb, ws):
    if not should_skip_sheet(ws.name):
        active_row = wb.app.selection.row
        ws.range("B4").copy(ws.range("B" + str(active_row)))
        ws.range("F4:G4").copy(ws.range("F" + str(active_row) + ":G" + str(active_row)))
        ws.range("L4").copy(ws.range("L" + str(active_row)))
        ws.range("N4:O4").copy(ws.range("N" + str(active_row) + ":O" + str(active_row)))
        ws.range("Q4:AA4").copy(
            ws.range("Q" + str(active_row) + ":AA" + str(active_row))
        )
        ws.range("AC4:AL4").copy(
            ws.range("AC" + str(active_row) + ":AL" + str(active_row))
        )


def delete_extra_empty_row(ws):
    """
    Delete consecutive empty rows (2 or more) from a worksheet.

    Optimized to read all data at once instead of row-by-row COM calls.
    """
    c_column = ws.range("C1500").end("up").row
    g_column = ws.range("G1500").end("up").row
    last_row = max(c_column, g_column)

    if last_row <= 1:
        return

    # Read all data at once (single COM call instead of row-by-row)
    data = ws.range(f"A1:H{last_row}").value

    # Handle single row case (value is a list, not list of lists)
    if last_row == 1:
        data = [data]

    # Find empty rows using numpy for speed
    empty_mask = np.array(
        [all(cell is None or cell == "" for cell in row) for row in data]
    )

    # Find ranges of consecutive empty rows (2 or more)
    ranges_to_delete = []
    i = 0
    while i < len(empty_mask):
        if empty_mask[i]:
            # Found start of empty region
            start = i
            while i < len(empty_mask) and empty_mask[i]:
                i += 1
            end = i
            # Only delete if 2 or more consecutive empty rows
            # Keep one empty row, delete the rest
            if end - start >= 2:
                # Delete from start+1 to end (keep one empty row)
                ranges_to_delete.append((start + 2, end))  # +2 for 1-based Excel row
        else:
            i += 1

    # Delete ranges from bottom to top to avoid index shifting
    for start_row, end_row in reversed(ranges_to_delete):
        ws.range(f"{start_row}:{end_row}").delete(shift="up")


def delete_extra_empty_row_wb(wb):

    for sheet in wb.sheets:
        if not should_skip_sheet(sheet.name):
            delete_extra_empty_row(sheet)


def format_cell_data_sheet(sheet):
    """
    Set the cell font and font size for a single sheet.
    Resets font to Arial size 12 for data rows, size 9 for headers.

    Optimized to format only used range instead of entire columns.
    """
    if not should_skip_sheet(sheet.name):
        last_row = sheet.range("C1500").end("up").row + 1
        lr = str(last_row)

        # Set cell font and size for data range only
        data_range = sheet.range(f"A3:BD{lr}")
        data_range.font.name = "Arial"
        data_range.font.size = 12
        sheet.range("2:2").font.size = 9
        sheet.range("C3").font.size = 14

        # Set cell number formats - optimized to use used range instead of entire columns
        # Integer format columns
        sheet.range(f"A1:B{lr}").number_format = "0"
        sheet.range(f"D1:D{lr}").number_format = "0"

        # Accounting format columns - batch adjacent columns together
        sheet.range(f"F1:G{lr}").number_format = ACCOUNTING
        sheet.range(f"K1:L{lr}").number_format = ACCOUNTING
        sheet.range(f"N1:O{lr}").number_format = ACCOUNTING
        sheet.range(f"R1:Z{lr}").number_format = ACCOUNTING
        sheet.range(f"AB1:AG{lr}").number_format = ACCOUNTING
        sheet.range(f"AI1:AJ{lr}").number_format = ACCOUNTING

        # Percentage format columns
        sheet.range(f"M1:M{lr}").number_format = "0.00%"
        sheet.range(f"AH1:AH{lr}").number_format = "0.00%"
        sheet.range("I1:R1").number_format = "0.00%"

        # Exchange rate format
        sheet.range(f"Q1:Q{lr}").number_format = EXCNANGE_RATE

        # Delete 'Category' and 'System' fields to avoid visual clutter.
        if sheet.range("AN2").value == "System":
            sheet.range("AN:AN").delete()
        if sheet.range("AM2").value == "Category":
            sheet.range("AM:AM").delete()
        sheet.range("AM2").value = "Leadtime"
        sheet.range("AN2").value = "Supplier"
        sheet.range("AO2").value = "Maker"


def format_cell_data(wb):
    """
    Set the cell font and font size for all sheets in workbook.
    Format the cell data to correct number or text representation.
    E.g. 1,000.00 or 1.00%
    """
    for sheet in wb.sheets:
        format_cell_data_sheet(sheet)


def download_file(path, filename, url):
    """
    path: directory
    filename: filename with extension
    url: url to download
    """
    import requests

    local_file_path = Path(path, filename)
    if not os.path.exists(local_file_path):
        response = requests.get(url)
        if response.status_code == 200:
            with open(local_file_path, "wb") as fd:
                for chunk in response.iter_content(chunk_size=8192):
                    fd.write(chunk)
            print(f"Downloaded {local_file_path}")
        else:
            print("Download is not necessary.")


# Download necessary files to local machine in 'Documents' folder
def download_logo():
    try:
        bid = os.path.join(os.path.expanduser("~/Documents"), "Bid")
        if not os.path.exists(bid):
            os.makedirs(bid)
        # Download Jason Logo
        download_file(
            bid,
            "Jason_Transparent_Logo_SS.png",
            "https://filedn.com/liTeg81ShEXugARC7cg981h/Bid/Jason_Transparent_Logo_SS.png",
        )
    except Exception as e:
        print(f"{e} has occured.")


# Can be done as tempfile
def download_template():
    try:
        bid = os.path.join(os.path.expanduser("~/Documents"), "Bid")
        filename = "Template.xlsx"
        file_path = Path(bid, filename)
        # Delete the file if exists
        if os.path.exists(file_path):
            os.remove(file_path)
        download_file(
            bid, filename, "https://filedn.com/liTeg81ShEXugARC7cg981h/Template.xlsx"
        )
        wb = xw.Book.caller()
        wb.app.books.open(file_path.absolute(), password=hide.legacy)
    except Exception as e:
        print(f"Failed to download template -> {e}")


def create_new_template():
    filename = "Template.xlsx"
    file_path = Path(RESOURCES, filename)
    wb = xw.Book.caller()
    wb.app.books.open(file_path.absolute(), password=hide.legacy)
    try:
        wb.app.books.active.save(
            Path("~/Downloads/Template.xlsx").expanduser(), password=hide.legacy
        )
        xw.apps.active.alert("Saved to Downloads as 'Template.xlsx'. Rename as required.")  # type: ignore
    except Exception:
        xw.apps.active.alert("Cannot save workbook. Save manually.")  # type: ignore


# Can be done as tempfile
def download_planner():
    try:
        bid = os.path.join(os.path.expanduser("~/Documents"), "Bid")
        filename = "Planner.xlsx"
        file_path = Path(bid, filename)
        # Delete the file if exists
        if os.path.exists(file_path):
            os.remove(file_path)
        download_file(
            bid,
            filename,
            "https://filedn.com/liTeg81ShEXugARC7cg981h/Project_Planner_R0.xlsx",
        )
        wb = xw.Book.caller()
        wb.app.books.open(file_path.absolute())
    except Exception as e:
        print(f"Failed to download template -> {e}")


def create_new_planner():
    filename = "Planner.xltx"
    file_path = Path(RESOURCES, filename)
    wb = xw.Book.caller()
    wb.app.books.open(file_path.absolute())


def update_template_version(wb):
    current_sheet = wb.sheets.active
    flag = 0
    try:
        current_wb_revision = int(wb.sheets["Config"].range("B15").value[1:])
        current_minor_revision = int(wb.sheets["Config"].range("C15").value[1:])
    except Exception:
        current_wb_revision = None
        current_minor_revision = None
    if current_wb_revision is None or current_wb_revision < int(LATEST_WB_VERSION[1:]):
        wb.sheets["Config"].range("D1:I20").clear()
        wb.sheets["Config"].range("95:106").delete()
        # Copy design elements from PERSONAL.XLSB (using cached ranges)
        get_cached_range("Design", "A28:E36").copy(wb.sheets["Config"].range("D2"))
        get_cached_range("Data", "C1:C2").copy(wb.sheets["Config"].range("B95"))
        get_cached_range("Data", "D1:D2").copy(wb.sheets["Config"].range("C95"))
        wb.sheets["Config"].range("A15").value = "Template Version"
        wb.sheets["Config"].range("B15").value = LATEST_WB_VERSION
        # Put currency and proposal type validation
        wb.sheets["Config"].activate()
        run_macro("put_currency_proposal_validation_formula")
        flag += 1

    if current_minor_revision is None or current_minor_revision < int(
        LATEST_MINOR_REVISION[1:]
    ):
        update_checklist(wb)  # Enabled the update checklist
        # xw.apps.active.alert("Called")  # type: ignore
        update_format(wb)
        summary(wb, discount=True)
        wb.sheets["Config"].range("C15").value = LATEST_MINOR_REVISION
        flag += 1

    if flag:
        wb.sheets[current_sheet].activate()
        xw.apps.active.alert(  # type: ignore
            f"The template has been updated to {LATEST_WB_VERSION}.{LATEST_MINOR_REVISION} {UPDATE_MESSAGE}"
        )
    # else:
    #     message = """
    #     No update is required. If you want to force an update, delete "Template Version" in cell "B15" & "C15" in "Config" sheet.
    #     Advisable to force an update if system or checklist is not available in dropdown list in "Technical_Notes".
    #     If item is not available in dropdown after forced update, there is no checklist or checklist is not ready.
    #     """
    #     xw.apps.active.alert(f"{message}")  # type: ignore


def update_checklist(wb):
    "Update checklist"
    wb.sheets["Config"].range("C15").value = LATEST_MINOR_REVISION
    # Clear previous data if any
    last_row = wb.sheets["Config"].range("A1500").end("up").row
    if last_row > 95:
        wb.sheets["Config"].range(f"A95:A{last_row}").clear()
    wb.sheets["Config"].range("A95").value = "SYSTEMS"
    # Write data from list
    cc.available_system_checklist_register.sort()
    wb.sheets["Config"].range("A96").options(transpose=True).value = [
        system.upper() for system in cc.available_system_checklist_register
    ]

    # Get Technical_Notes sheet (optional - may not exist in all workbooks)
    tn_sheet = get_sheet(wb, "Technical_Notes", required=False)

    # Test if value "Systems" is already there (only if Technical_Notes exists)
    if tn_sheet:
        cell_value = tn_sheet.range("F3")
        # if cell_value is None:
        if cell_value != "Systems".upper():
            # Copy from PERSONAL.XLSB (using cached range)
            get_cached_range("Data", "B1").copy(tn_sheet.range("F3"))
            # Call macro to fill in the dropdown formula
            tn_sheet.activate()
            run_macro("put_systems_validation_formula")

    # For general checklist
    # Clear previous data if any
    last_row = wb.sheets["Config"].range("E1500").end("up").row
    if last_row > 95:
        wb.sheets["Config"].range(f"E95:E{last_row}").clear()
    wb.sheets["Config"].range("E95").value = "CHECKLISTS"
    # Write data from list
    cc.available_checklist_register.sort()
    wb.sheets["Config"].range("E96").options(transpose=True).value = [
        system.upper() for system in cc.available_checklist_register
    ]

    # Test if value "Checklists" is already there (only if Technical_Notes exists)
    if tn_sheet:
        cell_value = tn_sheet.range("G3")
        # if cell_value is None:
        if cell_value != "Checklists".upper():
            # Copy from PERSONAL.XLSB (using cached range)
            get_cached_range("Data", "E1").copy(tn_sheet.range("G3"))
            # Call macro to fill in the dropdown formula
            tn_sheet.activate()
            run_macro("put_checklists_validation_formula")

            tn_sheet.range("F:G").autofit()

    # Add Num Scheme setting
    config = wb.sheets["Config"]
    config.range("A16").value = "Num Scheme"
    # Only set default if cell is empty (preserve user's existing choice)
    if config.range("B16").value is None:
        config.range("B16").value = "Single"
    if sys.platform == "win32":
        # Windows: Add dropdown validation (delete existing first to avoid error)
        try:
            config.range("B16").api.Validation.Delete()
        except Exception:
            pass  # No existing validation to delete
        config.range("B16").api.Validation.Add(Type=3, Formula1="Single,Double")


def update_format(wb):
    "Update cell formatting for sheet"
    "Separate out here because it needs to run only once and not everytime"
    for sheet in wb.sheets:
        if not should_skip_sheet(sheet.name):
            # Write titles
            # xw.apps.active.alert("Updating formats")
            sheet.range("AP2").value = "SCDQL"
            sheet.range("AQ2").value = "TCDQL"
            sheet.range("AR2").value = "BSCQL"
            sheet.range("AS2").value = "BTCQL"
            sheet.range("AT2").value = "SSPL"
            sheet.range("AU2").value = "TSPL"
            sheet.range("AV2").value = "TP"
            sheet.range("AW2").value = "TM"

            sheet.range("AP:AV").number_format = ACCOUNTING
            sheet.range("AW:AW").number_format = "0.00%"
            sheet.range("AP:AW").autofit()


if __name__ == "__main__":
    pass
