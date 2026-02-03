"""Multiple functions to support Excel automation.
© Thiha Aung (infowizard@gmail.com)
For the excel, the last row technically is 1048576.
However, I have hard-limited this to 1500 rows.
The code will need to be updated if more rows are needed.
"""

import getpass
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import requests
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
    Apply top and bottom border with color #0332FF to a row range.
    Compatible with both Mac and Windows platforms.

    Args:
        row_range: xlwings range object for the row to style
    """
    # Excel border edge constants
    xlEdgeTop = 8
    xlEdgeBottom = 9
    # Line style constants
    xlContinuous = 1
    # Weight constants
    xlThin = 2

    # Color #0332FF: R=3, G=50, B=255
    if sys.platform == "win32":
        # Windows: Use COM API directly (pure Python)
        # Color in BGR long integer format for Windows
        color = 255 * 65536 + 50 * 256 + 3  # 16724483
        for edge in [xlEdgeTop, xlEdgeBottom]:
            border = row_range.api.Borders(edge)
            border.LineStyle = xlContinuous
            border.Weight = xlThin
            border.Color = color
    else:
        # macOS: AppleScript/VBA limitations prevent direct border manipulation.
        # Fall back to copying border style from PERSONAL.XLSB Design sheet.
        get_cached_range("Design", "5:5").copy(row_range)


def _get_rfq_base_path() -> Path | None:
    """
    Get the user-specific @rfqs base path based on the current user.

    Returns:
        Path to the @rfqs folder, or None if it doesn't exist.
    """
    username = getpass.getuser()

    # User-specific path configurations
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
SKIP_SHEETS = ["Config", "Cover", "Summary", "Technical_Notes", "TN", "T&C", "Scratch"]


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


def title_case_ignore_double_char(text):
    words = text.split()
    titled_words = []
    for word in words:
        if (
            len(word.strip(string.punctuation)) > 2
        ):  # So that two letter words are ignored without punctuation mark
            # To prevent cases like 'mm)' from becoming 'Mm)'
            titled_words.append(word.title())
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
    acronyms = acronym_regex.findall(text)

    if title:
        text = title_case_ignore_double_char(text)
        # Improved function to restore acronyms
        for acronym in acronyms:
            acronym_regex = acronym.title()
            pattern = rf"\b{acronym_regex}\b"
            text = re.sub(pattern, acronym, text)
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
    """Function to replace description such as 1x, 20x, 10X ,
    x1, x20, X20 into 1 x, 20 x, 10 x, x 1, x 20, X 10 etc."""
    # For cases such as 20x, 30X. Allows if followed by -
    x = re.compile(r"\d+x(?!-)|\d+X(?!-)")
    if x.search(text):
        substring = re.findall(r"(\d+x|\d+X)", text)
        for word in substring:
            text = re.sub(word, (word[:-1] + " x"), text)
    # For cases such as x20, X30
    x = re.compile(r"(x\d+|X\d+)")
    if x.search(text):
        substring = re.findall(r"(x\d+|X\d+)", text)
        for word in substring:
            text = re.sub(word, ("x " + word[1:]), text)
    # For cases such as 20 X, 30 X
    x = re.compile(r"(\d+ X)")
    if x.search(text):
        substring = re.findall(r"(\d+ X)", text)
        for word in substring:
            text = re.sub(word, (word[:-1] + "x"), text)
    # For cases such as X 20, X 30
    x = re.compile(r"(X \d+)")
    if x.search(text):
        substring = re.findall(r"(X \d+)", text)
        for word in substring:
            text = re.sub(word, ("x" + word[1:]), text)
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
                '=IF(AND(D3<>"", K3<>"",H3<>"OPTION"),D3*N3,"")',  # O: SCD
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
                '=IF(AND(D3<>"", K3<>"", H3<>"OPTION", INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*R3, "")',
                # T: BUCQ
                '=IF(AND(D3<>"",K3<>""), (R3*(1+$L$1+$N$1+$P$1+$R$1))/(1-0.05),"")',
                # U: BSCQ
                '=IF(AND(D3<>"",K3<>"",H3<>"OPTION",INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*T3, "")',
                # V: Default escalation
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AQ3*$L$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), S3*$L$1, ""))',
                # W: Warranty
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AQ3*$N$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), S3*$N$1, ""))',
                # X: Freight
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AQ3*$P$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), S3*$P$1, ""))',
                # Y: Special
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AQ3*$R$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), S3*$R$1, ""))',
                # Z: Risk
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AS3-(AQ3+V3+W3+X3+Y3), IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), U3-(S3+V3+W3+X3+Y3), ""))',
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
                '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED",INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*AC3,"")',
                # AE: UPLS
                '=IF(AND(D3<>"",K3<>""), IF(AB3<>"", AB3, AC3),"")',
                # AF: SPLS
                '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED",INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*AE3,"")',
                # AG: Profit
                '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED",AF3<>""),AF3-U3,"")',
                # AH: Margin %
                '=IF(AND(AG3<>"", AG3<>0), AG3/AF3, "")',
                # AI: Total price
                '=IF(AND(D3<>"",K3<>"", H3<>"OPTION"), D3*AE3, "")',
            ]
        ]

        # BATCH 4: Columns F, G (2 adjacent columns) - Unit/Subtotal Price
        sheet.range("F3:G" + lr).formula = [
            [
                # F: Unit Price
                '=IF(AND(AL3="Title", ISNUMBER(AJ3)), AJ3, IF(AND(AL3="Lineitem", AK3="Lumpsum", H3<>"OPTION"), "", AE3))',
                # G: Subtotal Price
                '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED"), D3*F3,"")',
            ]
        ]

        # L: Subtotal Cost (single column)
        sheet.range("L3:L" + lr).formula = (
            '=IF(AND(D3<>"",K3<>"",H3<>"OPTION"),D3*K3,"")'
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
                '=IF(AND(ISNUMBER(D3), ISNUMBER(AP3), H3<>"OPTION"), D3*AP3, "")',
                # AR: BSCQL
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(U4:INDEX(U4:U1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), IF(AND(AL3="Lineitem", AK3="Unit Price"), T3, ""))',
                # AS: BTCQL (base cost)
                '=IF(AND(ISNUMBER(D3), ISNUMBER(AR3), H3<>"OPTION"), D3*AR3, "")',
                # AT: SSPL
                '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(AF4:INDEX(AF4:AF1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), IF(AND(AL3="Lineitem", AK3="Unit Price"), AE3, ""))',
                # AU: TSPL (selling price)
                '=IF(AND(ISNUMBER(D3), H3<>"WAIVED", H3<>"INCLUDED", H3<>"OPTION", ISNUMBER(AT3)), D3*AT3, "")',
                # AV: Total Profit
                '=IF(AND(ISNUMBER(D3), ISNUMBER(AS3), ISNUMBER(AU3)), AU3-AS3, "")',
                # AW: Grand Margin
                '=IF(AND(H3<>"OPTION", ISNUMBER(D3), ISNUMBER(AU3), AU3<>0, ISNUMBER(AV3)), AV3/AU3, "")',
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
        # Apply top and bottom border with color #0332FF (pure Python, cross-platform)
        row_range = sheet.range(f"{last_row + 2}:{last_row + 2}")
        apply_lastrow_border(row_range)
        sheet.range("F" + str(last_row + 2)).formula = '="Subtotal(" & Config!B12 & ")"'
        sheet.range("F" + str(last_row + 2)).font.size = 9
        sheet.range("G" + str(last_row + 2)).formula = (
            "=SUM(G3:G" + str(last_row + 1) + ")"
        )
        # SCDQ: Subtotal cost after discount in quoted currency
        # sheet.range("S" + str(last_row + 2)).formula = (
        #     "=SUM(S3:S" + str(last_row + 1) + ")"
        # )
        # BSCQ: Base subtotal cost in quoted currency
        # sheet.range("U" + str(last_row + 2)).formula = (
        #     "=SUM(U3:U" + str(last_row + 1) + ")"
        # )
        # Default
        sheet.range("V" + str(last_row + 2)).formula = (
            "=SUM(V3:V" + str(last_row + 1) + ")"
        )
        # Warranty
        sheet.range("W" + str(last_row + 2)).formula = (
            "=SUM(W3:W" + str(last_row + 1) + ")"
        )
        # Freight (Inbound)
        sheet.range("X" + str(last_row + 2)).formula = (
            "=SUM(X3:X" + str(last_row + 1) + ")"
        )
        # Special (Conditions)
        sheet.range("Y" + str(last_row + 2)).formula = (
            "=SUM(Y3:Y" + str(last_row + 1) + ")"
        )
        # Risk
        sheet.range("Z" + str(last_row + 2)).formula = (
            "=SUM(Z3:Z" + str(last_row + 1) + ")"
        )
        # sheet.range("AF" + str(last_row + 2)).formula = (
        #     "=SUM(AF3:AF" + str(last_row + 1) + ")"
        # )
        # sheet.range("AG" + str(last_row + 2)).formula = (
        #     "=SUM(AG3:AG" + str(last_row + 1) + ")"
        # )
        # sheet.range("AH" + str(last_row + 2)).formula = (
        #     "=AG" + str(last_row + 2) + "/AF" + str(last_row + 2)
        # )
        sheet.range("AL" + str(last_row + 2)).value = "Title"
        # TCDQL(Total Cost after Discount in Quoted Currency Lumpsum)
        # Material cost
        sheet.range(f"AQ{str(last_row + 2)}").formula = f"=SUM(AQ3:AQ{last_row+1})"
        # BTCQL (Base Total Cost in Quoted Currency Lumpsum)
        # Base price after escalation
        sheet.range(f"AS{str(last_row + 2)}").formula = f"=SUM(AS3:AS{last_row+1})"
        # TSPL (Total Selling Price Lumpsum)
        # Actual selling price
        sheet.range(f"AU{str(last_row + 2)}").formula = f"=SUM(AU3:AU{last_row+1})"
        # TP (Total Profit)
        sheet.range(f"AV{str(last_row + 2)}").formula = f"=SUM(AV3:AV{last_row+1})"
        # Total Margin
        sheet.range(f"AW{str(last_row + 2)}").formula = (
            f'=IF(AU{str(last_row+2)}<>0,AV{str(last_row + 2)}/AU{str(last_row + 2)}, "")'
        )
        # The formatting for added row.
        sheet.range(f"AW{str(last_row + 2)}").number_format = "0.00%"
        # Format
        # sheet.range(f"S{last_row+2}:S{last_row+2}").font.color = (0, 144, 81)
        sheet.range(f"V{last_row+2}:Z{last_row+2}").font.color = (0, 144, 81)
        sheet.range(f"{last_row+2}:{last_row+2}").font.bold = True

        # Set-up print area
        sheet.page_setup.print_area = "A1:H" + str(last_row + 2)


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


def adjust_columns(sheet):
    """Unhide all columns while setting the width for selected columns"""
    if not should_skip_sheet(sheet.name):
        sheet.range("A:A").column_width = 5
        sheet.range("B:B").autofit()
        sheet.range("C:C").column_width = 55
        sheet.range("C:C").rows.autofit()
        sheet.range("C:C").wrap_text = True
        sheet.range("D:H").autofit()


def adjust_columns_wb(wb):
    for sheet in wb.sheets:
        adjust_columns(sheet)


def hide_columns(sheet):
    if not should_skip_sheet(sheet.name):
        sheet.range("AI:AL").column_width = 0
        sheet.range("AC:AD").column_width = 0
        sheet.range("AF:AF").column_width = 0
        sheet.range("S:AA").column_width = 0
        sheet.range("Q:Q").column_width = 0
        sheet.range("P:P").column_width = 20
        sheet.range("P:P").wrap_text = False
        sheet.range("R:R").autofit()
        sheet.range("O:O").column_width = 0
        sheet.range("L:L").autofit()
        sheet.range("T:T").autofit()
        sheet.range("AB:AB").autofit()
        sheet.range("AE:AE").autofit()
        sheet.range("AG:AH").autofit()
        sheet.range("AM:AP").autofit()
        sheet.range("M:N").autofit()
        sheet.range("D:H").autofit()
        sheet.range("I:I").column_width = 10
        sheet.range("I:I").wrap_text = False
        sheet.range("J:K").autofit()
        sheet.range("C:C").column_width = 55
        sheet.range("C:C").rows.autofit()
        sheet.range("C:C").wrap_text = True
        sheet.range("B:B").autofit()
        sheet.range("A:A").column_width = 5


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
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",D20:D' + str(offset) + ")"
        )
        sheet.range("E" + str(offset + 1)).formula = (
            "=IF(COUNTIF(E20:E" + str(offset) + ',"OPTION"), "Excluding Option", "")'
        )
        sheet.range("H" + str(offset + 1)).formula = (
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",H20:H' + str(offset) + ")"
        )
        sheet.range("I" + str(offset + 1)).formula = (
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",I20:I' + str(offset) + ")"
        )
        sheet.range("J" + str(offset + 1)).formula = (
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",J20:J' + str(offset) + ")"
        )
        sheet.range("K" + str(offset + 1)).formula = (
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",K20:K' + str(offset) + ")"
        )
        sheet.range("L" + str(offset + 1)).formula = (
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",L20:L' + str(offset) + ")"
        )
        sheet.range("M" + str(offset + 1)).formula = (
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",M20:M' + str(offset) + ")"
        )
        sheet.range("N" + str(offset + 1)).formula = (
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",N20:N' + str(offset) + ")"
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
                "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."
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
                "• Total project price does not include items marked 'OPTION' in the detailed bill of material."
            )
            sheet.range("C" + str(offset + 5)).value = (
                "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."
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
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",D20:D' + str(offset) + ")"
        )
        sheet.range("E" + str(offset + 1)).formula = (
            "=IF(COUNTIF(E20:E" + str(offset) + ',"OPTION"), "Excluding Option", "")'
        )
        sheet.range("H" + str(offset + 1)).formula = (
            "=SUMIF(E20:E" + str(offset) + ',"<>OPTION",H20:H' + str(offset) + ")"
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
                "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."
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
                "• Total project price does not include items marked 'OPTION' in the detailed bill of material."
            )
            sheet.range("C" + str(offset + 5)).value = (
                "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."
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
            last_row = wb.sheets[sheet].range("C1500").end("up").row
            wb.sheets[sheet].activate()
            wb.sheets[sheet].range("C:C").autofit()
            wb.sheets[sheet].range("C:C").column_width = 60
            wb.sheets[sheet].range("C:C").wrap_text = True
            wb.sheets[sheet].range("D:F").autofit()
            # Adjust the last two rows so that unwanted pagebreak can be prevented
            wb.sheets[sheet].range(f"{last_row+1}:{last_row+1}").delete()
            wb.sheets[sheet].range(f"{last_row+1}:{last_row+1}").row_height = 2
            run_macro("conditional_format")
            run_macro("remove_h_borders")
            run_macro("pagebreak_borders")
    wb.sheets[current_sheet].activate()


def technical(wb):
    directory, is_cloud = get_workbook_directory(wb)
    # Check if Technical PDF already exist
    temp_file_name = Path(directory, "Technical " + wb.name[:-4] + "pdf")
    if temp_file_name.is_file():
        xw.apps.active.alert(  # type: ignore
            "The Technical PDF file already exists!\n Please delete the file and try again."
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

    if wb.name[:10] == "Commercial":
        for sheet in wb.sheet_names:
            ws = wb.sheets[sheet]
            wb.sheets[2].activate()
            if not should_skip_sheet(sheet):
                # Require to remove h_borders as these willl not be detected
                # when columns are removed and page setup changed.
                run_macro("remove_h_borders")
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
        prepare_to_print_technical(wb)
        wb.sheets["Summary"].activate()
        file_name = "Technical " + wb.name[11:-4] + "xlsx"
        full_path = Path(directory, file_name)
        wb.save(full_path, password="")
        pdf_path = full_path.with_suffix(".pdf")
        print_technical(wb, pdf_path=str(pdf_path))
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
        prepare_to_print_technical(wb)
        # wb.sheets["Summary"].activate()
        file_name = "Technical " + wb.name[:-4] + "xlsx"
        full_path = Path(directory, file_name)
        wb.save(full_path, password="")
        pdf_path = full_path.with_suffix(".pdf")
        print_technical(wb, pdf_path=str(pdf_path))


def commercial(wb):
    directory, is_cloud = get_workbook_directory(wb)
    # Check if Commercial PDF already exists
    temp_file_name = Path(directory, "Commercial " + wb.name[:-4] + "pdf")
    if temp_file_name.is_file():
        xw.apps.active.alert(  # type: ignore
            "The Commercial PDF file already exists!\n Please delete the file and try again."
        )
        return

    """Takes a work book, set horizantal borders at pagebreaks."""
    # current_sheet = wb.sheets.active
    wb.sheets["Cover"].range("D6:D8").value = (
        wb.sheets["Cover"].range("D6:D8").raw_value
    )
    wb.sheets["Cover"].range("D39").value = wb.sheets["Config"].range("B13").value
    wb.sheets["Cover"].range("D40").value = wb.sheets["Config"].range("B14").value
    wb.sheets["Cover"].range("C42:C47").value = (
        wb.sheets["Cover"].range("C42:C47").raw_value
    )
    last_row = wb.sheets["Summary"].range("D1500").end("up").row
    wb.sheets["Summary"].range(f"G20:P{last_row}").value = (
        wb.sheets["Summary"].range(f"G20:P{last_row}").raw_value
    )
    wb.sheets["Summary"].range("C20:C100").value = (
        wb.sheets["Summary"].range("C20:C100").raw_value
    )
    page_setup(wb)
    for sheet in wb.sheet_names:
        ws = wb.sheets[sheet]
        ws.range("A1").value = ws.range("A1").raw_value  # Remove formula
        ws.range("A1").wrap_text = False
        if not should_skip_sheet(sheet):
            last_row = ws.range("G1500").end("up").row
            ws.activate()
            # Adjust column width as sometimes, the long value does not show.
            ws.range(f"A3:AL{last_row}").value = ws.range(f"A3:AL{last_row}").raw_value
            ws.range("A:A").column_width = 4
            ws.range("B:B").autofit()
            ws.range("C:C").autofit()
            ws.range("C:C").column_width = 55
            # wb.sheets[sheet].range('C:C').wrap_text =
            ws.range(f"G3:G{last_row-1}").formula = (
                '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED"), D3*F3,"")'
            )
            ws.range(f"G{last_row}").formula = "=SUM(G3:G" + str(last_row - 1) + ")"
            wb.sheets[sheet].range("D:H").autofit()
            ws = wb.sheets[sheet]  # Refresh stale reference before column deletions
            ws.range("AM:BD").delete()
            ws.range("I:AK").delete()
            ws = wb.sheets[sheet]  # Refresh again after column deletions
            col_i_values = ws.range(f"I1:I{last_row}").options(ndim=1).value
            ws.range("I:I").delete()
            if col_i_values:
                ws.range(f"AL1:AL{last_row}").value = [[v] for v in col_i_values]
            ws.range("AL:AL").column_width = 0
            # Call macros
            run_macro("conditional_format")
            run_macro("remove_h_borders")
            run_macro("pagebreak_borders")

    wb.sheets["Summary"].range("G:X").delete()
    wb.sheets["Config"].delete()
    tn_sheet = get_sheet(wb, "Technical_Notes", required=False)
    if tn_sheet:
        tn_sheet.range("F:I").delete()
    wb.sheets["Summary"].activate()
    file_name = "Commercial " + wb.name[:-4] + "xlsx"
    full_path = Path(directory, file_name)
    wb.save(full_path, password="")
    # Explicitly specify PDF path - don't rely on xlwings to figure it out
    # (SharePoint sync can cause stale workbook path references)
    pdf_path = full_path.with_suffix(".pdf")
    try:
        wb.to_pdf(path=str(pdf_path), show=True)
    except Exception as e:
        # The program does not override the existing file. Therefore, the file needs to be removed if it exists.
        # xw.apps.active.alert('The PDF file already exists!\n Please delete the file and try again.')
        xw.apps.active.alert(  # type: ignore
            f"This error is encountered {e}. The PDF file already exists?"
        )


def prepare_to_print_internal(wb):
    """Takes a work book, set horizantal borders at pagebreaks."""
    current_sheet = wb.sheets.active
    page_setup(wb)
    for sheet in wb.sheet_names:
        if not should_skip_sheet(sheet):
            wb.sheets[sheet].activate()
            run_macro("conditional_format_internal_costing")
            run_macro("remove_h_borders")
            # Below is commented out so that blue lines do not show
            # MACRO_NB.macro('pagebreak_borders')()
    wb.sheets[current_sheet].activate()


def print_technical(wb, pdf_path=None):
    """The technical proposal will be written to the specified path or cwd."""
    try:
        if pdf_path:
            wb.to_pdf(path=pdf_path, show=True)
        else:
            wb.to_pdf(show=True)
    except Exception:
        # The program does not override the existing file. The file needs to be removed if it exists.
        xw.apps.active.alert(  # type: ignore
            "The PDF file already exists!\n Please delete the file and try again."
        )


def apply_conditional_format(sheet):
    """
    Apply conditional formatting to column C based on AL values.
    Uses xlwings API - no sheet activation required.
    """
    col_c = sheet.range("C:C")

    # Excel constants
    xlExpression = 2
    xlUnderlineStyleSingle = 2

    # Clear existing conditional formats
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


def apply_remove_h_borders(sheet):
    """
    Remove horizontal inside borders from the data range.
    """
    xlInsideHorizontal = 12
    xlNone = -4142

    # Find last row with data
    last_row = sheet.range("A1").end("down").row
    if last_row > 5:  # Ensure we have data
        data_range = sheet.range(f"A3:H{last_row - 2}")
        data_range.api.Borders(xlInsideHorizontal).LineStyle = xlNone


def apply_format_column_border(sheet):
    """
    Apply column border formatting to the sheet.
    """
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

    def set_border(rng, edge, color=None, theme_color=None, tint=0, weight=xlThin):
        border = rng.api.Borders(edge)
        border.LineStyle = xlContinuous
        border.Weight = weight
        if color is not None:
            border.Color = color
            border.TintAndShade = 0
        elif theme_color is not None:
            border.ThemeColor = theme_color
            border.TintAndShade = tint

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


def conditional_format_wb(wb):
    """
    Apply conditional formatting to all sheets.
    On Windows: Uses Python/xlwings API (no sheet activation, no focus stealing).
    On macOS: Uses VBA macros (AppleScript doesn't support FormatConditions/Borders API).
    """
    current_sheet = wb.sheets.active
    is_windows = sys.platform == "win32"

    for sheet_name in wb.sheet_names:
        if not should_skip_sheet(sheet_name):
            sheet = wb.sheets[sheet_name]

            if is_windows:
                # Windows: use Python API (no sheet activation required)
                try:
                    apply_conditional_format(sheet)
                    apply_remove_h_borders(sheet)
                    apply_format_column_border(sheet)
                except Exception:
                    # Fallback to VBA if API fails
                    sheet.activate()
                    run_macro("conditional_format")
                    run_macro("remove_h_borders")
                    run_macro("format_column_border")
            else:
                # macOS: use VBA directly (AppleScript doesn't support these APIs)
                sheet.activate()
                run_macro("conditional_format")
                run_macro("remove_h_borders")
                run_macro("format_column_border")

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
    wb.save(Path(directory, file_name), password="")


def convert_legacy(wb):
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
