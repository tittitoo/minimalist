"""
Unit tests for the proposal automation functions.
Run with: python tests.py

Tests cover:
- Text formatting functions (set_nitty_gritty, set_x, set_comma_space, etc.)
- number_title vectorized logic
- format_text vectorized logic

These tests don't require Excel - they test the pure Python/pandas logic.
"""

import unittest
import pandas as pd
import re
import tempfile
from pathlib import Path

# Import functions to test
from functions import (
    set_nitty_gritty,
    set_comma_space,
    set_x,
    set_case_preserve_acronym,
    title_case_ignore_double_char,
    SKIP_SHEETS,
    SHEET_ALIASES,
    resolve_sheet_name,
    is_sheet_name,
    get_sheet,
    sheet_exists,
    should_skip_sheet,
    _find_workbook_in_rfqs,
    sanitize_config_string,
    sanitize_config_date,
)
from datetime import datetime


class TestSetNittyGritty(unittest.TestCase):
    """Tests for set_nitty_gritty text cleanup function."""

    def test_strips_whitespace(self):
        self.assertEqual(set_nitty_gritty("  hello  "), "hello")

    def test_removes_multiple_spaces(self):
        self.assertEqual(set_nitty_gritty("hello   world"), "hello world")

    def test_converts_dash_to_bullet(self):
        self.assertEqual(set_nitty_gritty("- item"), "• item")

    def test_converts_tilde_to_bullet(self):
        self.assertEqual(set_nitty_gritty("~ item"), "• item")

    def test_converts_asterisk_space_to_bullet(self):
        self.assertEqual(set_nitty_gritty("* item"), " • item")

    def test_semicolon_to_colon_at_end(self):
        self.assertEqual(set_nitty_gritty("includes;"), "includes:")

    def test_preserves_normal_text(self):
        self.assertEqual(set_nitty_gritty("normal text"), "normal text")


class TestSetCommaSpace(unittest.TestCase):
    """Tests for set_comma_space function."""

    def test_removes_space_before_comma(self):
        self.assertEqual(set_comma_space("hello , world"), "hello, world")

    def test_adds_space_after_comma(self):
        self.assertEqual(set_comma_space("hello,world"), "hello, world")

    def test_preserves_numbers_with_commas(self):
        # Numbers like 1,200 should not be affected
        self.assertEqual(set_comma_space("price is 1,200"), "price is 1,200")

    def test_handles_multiple_commas(self):
        result = set_comma_space("a,b,c")
        self.assertIn(", ", result)


class TestSetX(unittest.TestCase):
    """Tests for set_x function that normalizes 'x' notation."""

    def test_20x_becomes_20_x(self):
        self.assertEqual(set_x("20x items"), "20 x items")

    def test_30X_becomes_30_x(self):
        self.assertEqual(set_x("30X items"), "30 x items")

    def test_x20_becomes_x_20(self):
        self.assertEqual(set_x("x20 items"), "x 20 items")

    def test_X30_becomes_x_30(self):
        self.assertEqual(set_x("X30 items"), "x 30 items")

    def test_20_X_becomes_20_x(self):
        self.assertEqual(set_x("20 X items"), "20 x items")

    def test_X_20_becomes_x_20(self):
        self.assertEqual(set_x("X 20 items"), "x 20 items")

    def test_preserves_hyphenated(self):
        # 20x- should not be changed (hyphen follows)
        self.assertEqual(set_x("20x-connector"), "20x-connector")


class TestSetCasePreserveAcronym(unittest.TestCase):
    """Tests for set_case_preserve_acronym function."""

    def test_title_case_preserves_acronyms(self):
        result = set_case_preserve_acronym("IP camera system", title=True)
        self.assertIn("IP", result)

    def test_title_case_preserves_mixed_case(self):
        result = set_case_preserve_acronym("iPhone charger", title=True)
        self.assertIn("iPhone", result)

    def test_title_case_basic(self):
        result = set_case_preserve_acronym("hello world", title=True)
        self.assertEqual(result, "Hello World")

    def test_upper_case(self):
        result = set_case_preserve_acronym("hello world", upper=True)
        self.assertEqual(result, "HELLO WORLD")


class TestTitleCaseIgnoreDoubleChar(unittest.TestCase):
    """Tests for title_case_ignore_double_char function."""

    def test_title_cases_long_words(self):
        result = title_case_ignore_double_char("hello world")
        self.assertEqual(result, "Hello World")

    def test_ignores_two_letter_words(self):
        result = title_case_ignore_double_char("it is ok")
        # Two letter words should not be title-cased
        self.assertEqual(result, "it is ok")

    def test_mixed_length_words(self):
        result = title_case_ignore_double_char("the IP camera is on")
        self.assertIn("The", result)
        self.assertIn("Camera", result)


class TestNumberTitleLogic(unittest.TestCase):
    """Tests for the vectorized number_title logic."""

    def test_main_title_numbering(self):
        """Test that numeric values get sequential numbers."""
        test_data = {
            "NO": [10, "a", 20, "b", 30],
            "Description": ["Sys A", "Item 1", "Sys B", "Item 2", "Sys C"],
            "System": ["TEST"] * 5,
        }
        systems = pd.DataFrame(test_data)

        # Apply vectorized logic
        count, step = 10, 10
        no_col = systems["NO"].fillna("")

        def is_numeric(x):
            try:
                return bool(int(x)) if x != "" else False
            except (ValueError, TypeError):
                return False

        is_main_title = no_col.apply(is_numeric)
        title_cumsum = is_main_title.cumsum()
        systems.loc[is_main_title, "NO"] = count + (title_cumsum[is_main_title] - 1) * step

        # Verify main titles got sequential numbers
        self.assertEqual(systems.loc[0, "NO"], 10)
        self.assertEqual(systems.loc[2, "NO"], 20)
        self.assertEqual(systems.loc[4, "NO"], 30)

    def test_sub_item_numbering(self):
        """Test that sub-items get braille markers."""
        test_data = {
            "NO": [10, "a", "b", 20, "x"],
            "Description": ["Sys A", "Item 1", "Item 2", "Sys B", "Item 3"],
            "System": ["TEST"] * 5,
        }
        systems = pd.DataFrame(test_data)

        # Apply vectorized logic
        count, step = 10, 10
        no_col = systems["NO"].fillna("")

        def is_numeric(x):
            try:
                return bool(int(x)) if x != "" else False
            except (ValueError, TypeError):
                return False

        def starts_with_letter(x):
            if isinstance(x, str) and x.strip():
                return bool(re.match(r"^[A-Z]", x.strip()))
            return False

        is_main_title = no_col.apply(is_numeric)
        starts_with_az = no_col.apply(starts_with_letter)
        is_sub_item = (~is_main_title) & (~starts_with_az) & (no_col.astype(str).str.strip() != "")

        title_cumsum = is_main_title.cumsum()
        systems.loc[is_main_title, "NO"] = count + (title_cumsum[is_main_title] - 1) * step

        if is_sub_item.any():
            group_id = title_cumsum
            sub_item_count = systems[is_sub_item].groupby(group_id[is_sub_item]).cumcount() + 1
            systems.loc[is_sub_item, "NO"] = "⠠" + sub_item_count.astype(str)

        # Verify sub-items got braille markers
        self.assertEqual(systems.loc[1, "NO"], "⠠1")
        self.assertEqual(systems.loc[2, "NO"], "⠠2")
        self.assertEqual(systems.loc[4, "NO"], "⠠1")  # Resets after new title

    def test_preserves_uppercase_letters(self):
        """Test that values starting with A-Z are preserved."""
        test_data = {
            "NO": [10, "A", "B", 20],
            "Description": ["Sys A", "Note A", "Note B", "Sys B"],
            "System": ["TEST"] * 4,
        }
        systems = pd.DataFrame(test_data)

        no_col = systems["NO"].fillna("")

        def starts_with_letter(x):
            if isinstance(x, str) and x.strip():
                return bool(re.match(r"^[A-Z]", x.strip()))
            return False

        starts_with_az = no_col.apply(starts_with_letter)

        # Values starting with A-Z should be identified
        self.assertTrue(starts_with_az[1])
        self.assertTrue(starts_with_az[2])
        self.assertFalse(starts_with_az[0])
        self.assertFalse(starts_with_az[3])


class TestFormatTextLogic(unittest.TestCase):
    """Tests for the vectorized format_text logic."""

    def test_unit_normalization_nos_to_ea(self):
        """Test that 'nos' and 'no' become 'ea'."""
        systems = pd.DataFrame({"Unit": ["NOS", "no", "pcs", "ea"]})
        systems["Unit"] = systems["Unit"].astype(str).str.strip().str.lower()
        systems.loc[systems["Unit"].isin(["nos", "no"]), "Unit"] = "ea"

        self.assertEqual(systems.loc[0, "Unit"], "ea")
        self.assertEqual(systems.loc[1, "Unit"], "ea")
        self.assertEqual(systems.loc[2, "Unit"], "pcs")
        self.assertEqual(systems.loc[3, "Unit"], "ea")

    def test_unit_removes_trailing_s(self):
        """Test that trailing 's' is removed from units."""
        systems = pd.DataFrame({"Unit": ["meters", "pcs", "lots", "m"]})
        systems["Unit"] = systems["Unit"].astype(str).str.strip().str.lower()
        mask_trailing_s = (systems["Unit"].str.len() > 1) & (systems["Unit"].str[-1] == "s")
        systems.loc[mask_trailing_s, "Unit"] = systems.loc[mask_trailing_s, "Unit"].str[:-1]

        self.assertEqual(systems.loc[0, "Unit"], "meter")
        self.assertEqual(systems.loc[1, "Unit"], "pc")
        self.assertEqual(systems.loc[2, "Unit"], "lot")
        self.assertEqual(systems.loc[3, "Unit"], "m")  # Single char, not changed

    def test_scope_normalization(self):
        """Test scope values are normalized to INCLUDED/OPTION/WAIVED."""
        systems = pd.DataFrame({
            "Scope": ["included", "INCLUSIVE", "optional", "option", "waived", ""]
        })
        systems["Scope"] = systems["Scope"].astype(str).str.strip().str.lower()
        systems.loc[systems["Scope"].isin(["inclusive", "include", "included"]), "Scope"] = "INCLUDED"
        systems.loc[systems["Scope"].isin(["option", "optional"]), "Scope"] = "OPTION"
        systems.loc[systems["Scope"] == "waived", "Scope"] = "WAIVED"

        self.assertEqual(systems.loc[0, "Scope"], "INCLUDED")
        self.assertEqual(systems.loc[1, "Scope"], "INCLUDED")
        self.assertEqual(systems.loc[2, "Scope"], "OPTION")
        self.assertEqual(systems.loc[3, "Scope"], "OPTION")
        self.assertEqual(systems.loc[4, "Scope"], "WAIVED")

    def test_description_indentation(self):
        """Test that Description rows get proper indentation."""
        systems = pd.DataFrame({
            "Description": ["Item A", "Sub item 1", "Sub item 2"],
            "Format": ["Title", "Description", "Description"],
        })

        mask = systems["Format"] == "Description"
        desc_col = systems.loc[mask, "Description"].str.strip().str.lstrip("• ")

        # Default bullet
        result = "   • " + desc_col
        systems.loc[mask, "Description"] = result

        self.assertTrue(systems.loc[1, "Description"].startswith("   • "))
        self.assertTrue(systems.loc[2, "Description"].startswith("   • "))
        self.assertFalse(systems.loc[0, "Description"].startswith("   • "))

    def test_hash_prefix_becomes_triangle_bullet(self):
        """Test that # prefix becomes ‣ bullet."""
        systems = pd.DataFrame({
            "Description": ["# Note item", "Regular item"],
            "Format": ["Description", "Description"],
        })

        mask = systems["Format"] == "Description"
        desc_col = systems.loc[mask, "Description"].str.strip().str.lstrip("• ")
        starts_hash = desc_col.str.startswith("#")

        result = pd.Series(index=desc_col.index, dtype=str)
        result[starts_hash] = "      ‣ " + desc_col[starts_hash].str.lstrip("# ")
        result[~starts_hash] = "   • " + desc_col[~starts_hash]
        systems.loc[mask, "Description"] = result

        self.assertTrue(systems.loc[0, "Description"].startswith("      ‣ "))
        self.assertTrue(systems.loc[1, "Description"].startswith("   • "))


class TestSkipSheets(unittest.TestCase):
    """Test that SKIP_SHEETS constant is defined correctly."""

    def test_skip_sheets_contains_expected(self):
        expected = ["Config", "Cover", "Summary", "Technical_Notes", "TN", "T&C", "Scratch"]
        for sheet in expected:
            self.assertIn(sheet, SKIP_SHEETS)

    def test_skip_sheets_is_list(self):
        self.assertIsInstance(SKIP_SHEETS, list)


class TestShouldSkipSheet(unittest.TestCase):
    """Test should_skip_sheet helper function for case-insensitive Scratch handling."""

    def test_skips_standard_sheets(self):
        """Standard sheets in SKIP_SHEETS should be skipped."""
        for sheet in ["Config", "Cover", "Summary", "Technical_Notes", "TN", "T&C"]:
            self.assertTrue(should_skip_sheet(sheet), f"{sheet} should be skipped")

    def test_skips_scratch_exact_case(self):
        """Scratch with exact case should be skipped."""
        self.assertTrue(should_skip_sheet("Scratch"))

    def test_skips_scratch_lowercase(self):
        """scratch (lowercase) should be skipped."""
        self.assertTrue(should_skip_sheet("scratch"))

    def test_skips_scratch_uppercase(self):
        """SCRATCH (uppercase) should be skipped."""
        self.assertTrue(should_skip_sheet("SCRATCH"))

    def test_skips_scratch_mixed_case(self):
        """ScRaTcH (mixed case) should be skipped."""
        self.assertTrue(should_skip_sheet("ScRaTcH"))

    def test_does_not_skip_system_sheets(self):
        """System/product sheets should not be skipped."""
        for sheet in ["CCTV", "Access Control", "Fire Alarm", "System1"]:
            self.assertFalse(should_skip_sheet(sheet), f"{sheet} should NOT be skipped")

    def test_does_not_skip_partial_scratch_match(self):
        """Sheet names containing 'scratch' but not exactly 'scratch' should not be skipped."""
        self.assertFalse(should_skip_sheet("Scratch2"))
        self.assertFalse(should_skip_sheet("MyScratch"))
        self.assertFalse(should_skip_sheet("Scratch_Notes"))


class TestSheetAliases(unittest.TestCase):
    """Test sheet name aliasing functionality."""

    def test_technical_notes_alias_defined(self):
        """TN should be an alias for Technical_Notes."""
        self.assertEqual(SHEET_ALIASES.get("TN"), "Technical_Notes")

    def test_resolve_alias(self):
        """resolve_sheet_name should convert TN to Technical_Notes."""
        self.assertEqual(resolve_sheet_name("TN"), "Technical_Notes")

    def test_resolve_canonical_unchanged(self):
        """resolve_sheet_name should return canonical names unchanged."""
        self.assertEqual(resolve_sheet_name("Technical_Notes"), "Technical_Notes")
        self.assertEqual(resolve_sheet_name("Config"), "Config")
        self.assertEqual(resolve_sheet_name("Summary"), "Summary")

    def test_resolve_unknown_unchanged(self):
        """resolve_sheet_name should return unknown names unchanged."""
        self.assertEqual(resolve_sheet_name("Unknown_Sheet"), "Unknown_Sheet")

    def test_is_sheet_name_with_alias(self):
        """is_sheet_name should match alias to canonical name."""
        self.assertTrue(is_sheet_name("TN", "Technical_Notes"))

    def test_is_sheet_name_with_canonical(self):
        """is_sheet_name should match canonical name to itself."""
        self.assertTrue(is_sheet_name("Technical_Notes", "Technical_Notes"))

    def test_is_sheet_name_mismatch(self):
        """is_sheet_name should return False for non-matching names."""
        self.assertFalse(is_sheet_name("Config", "Technical_Notes"))
        self.assertFalse(is_sheet_name("TN", "Config"))


class MockWorkbook:
    """Mock workbook for testing get_sheet and sheet_exists without Excel."""

    def __init__(self, sheet_names_list):
        self._sheet_names = sheet_names_list
        self._sheets = {name: f"Sheet:{name}" for name in sheet_names_list}

    @property
    def sheet_names(self):
        return self._sheet_names

    @property
    def sheets(self):
        return self._sheets


class TestGetSheetOptional(unittest.TestCase):
    """Tests for get_sheet with required=False parameter."""

    def test_get_sheet_returns_sheet_when_exists(self):
        """get_sheet should return the sheet when it exists."""
        wb = MockWorkbook(["Config", "Technical_Notes", "Summary"])
        result = get_sheet(wb, "Technical_Notes")
        self.assertEqual(result, "Sheet:Technical_Notes")

    def test_get_sheet_returns_sheet_via_alias(self):
        """get_sheet should find sheet via alias."""
        wb = MockWorkbook(["Config", "TN", "Summary"])
        result = get_sheet(wb, "Technical_Notes")
        self.assertEqual(result, "Sheet:TN")

    def test_get_sheet_required_true_raises_on_missing(self):
        """get_sheet with required=True should raise KeyError when sheet missing."""
        wb = MockWorkbook(["Config", "Summary"])
        with self.assertRaises(KeyError):
            get_sheet(wb, "Technical_Notes", required=True)

    def test_get_sheet_required_false_returns_none_on_missing(self):
        """get_sheet with required=False should return None when sheet missing."""
        wb = MockWorkbook(["Config", "Summary"])
        result = get_sheet(wb, "Technical_Notes", required=False)
        self.assertIsNone(result)

    def test_get_sheet_required_false_returns_sheet_when_exists(self):
        """get_sheet with required=False should still return sheet when it exists."""
        wb = MockWorkbook(["Config", "Technical_Notes", "Summary"])
        result = get_sheet(wb, "Technical_Notes", required=False)
        self.assertEqual(result, "Sheet:Technical_Notes")


class TestSheetExists(unittest.TestCase):
    """Tests for sheet_exists helper function."""

    def test_sheet_exists_true_for_canonical_name(self):
        """sheet_exists should return True when sheet exists by canonical name."""
        wb = MockWorkbook(["Config", "Technical_Notes", "Summary"])
        self.assertTrue(sheet_exists(wb, "Technical_Notes"))

    def test_sheet_exists_true_for_alias(self):
        """sheet_exists should return True when sheet exists by alias."""
        wb = MockWorkbook(["Config", "TN", "Summary"])
        self.assertTrue(sheet_exists(wb, "Technical_Notes"))

    def test_sheet_exists_true_when_query_by_alias(self):
        """sheet_exists should return True when queried by alias for existing sheet."""
        wb = MockWorkbook(["Config", "Technical_Notes", "Summary"])
        self.assertTrue(sheet_exists(wb, "TN"))

    def test_sheet_exists_false_when_missing(self):
        """sheet_exists should return False when sheet doesn't exist."""
        wb = MockWorkbook(["Config", "Summary"])
        self.assertFalse(sheet_exists(wb, "Technical_Notes"))

    def test_sheet_exists_false_for_unknown_sheet(self):
        """sheet_exists should return False for unknown sheet names."""
        wb = MockWorkbook(["Config", "Summary"])
        self.assertFalse(sheet_exists(wb, "Unknown_Sheet"))


class TestFindWorkbookInRfqs(unittest.TestCase):
    """Tests for _find_workbook_in_rfqs SharePoint folder lookup."""

    def setUp(self):
        """Create a temporary directory structure mimicking @rfqs."""
        self.temp_dir = tempfile.TemporaryDirectory()
        self.base_path = Path(self.temp_dir.name)

    def tearDown(self):
        """Clean up temporary directory."""
        self.temp_dir.cleanup()

    def _create_structure(self, *paths):
        """Create files at given relative paths."""
        for path in paths:
            full_path = self.base_path / path
            full_path.parent.mkdir(parents=True, exist_ok=True)
            full_path.touch()

    def test_finds_workbook_in_commercial_folder(self):
        """Should find workbook in standard 01-Commercial location."""
        self._create_structure("2026/ProjectABC/01-Commercial/JEC-2026-001-v1.xlsx")
        result = _find_workbook_in_rfqs("JEC-2026-001-v1.xlsx", self.base_path)
        self.assertEqual(result, self.base_path / "2026/ProjectABC/01-Commercial")

    def test_finds_workbook_case_insensitive(self):
        """Should match filenames case-insensitively."""
        self._create_structure("2026/ProjectABC/01-Commercial/JEC-2026-001-v1.xlsx")
        result = _find_workbook_in_rfqs("jec-2026-001-v1.XLSX", self.base_path)
        self.assertEqual(result, self.base_path / "2026/ProjectABC/01-Commercial")

    def test_returns_none_when_not_found(self):
        """Should return None when workbook doesn't exist."""
        self._create_structure("2026/ProjectABC/01-Commercial/other-file.xlsx")
        result = _find_workbook_in_rfqs("nonexistent.xlsx", self.base_path)
        self.assertIsNone(result)

    def test_returns_shallowest_match(self):
        """Should return shallowest folder when file exists at multiple depths."""
        self._create_structure(
            "2026/ProjectABC/test.xlsx",  # depth 2
            "2026/ProjectABC/01-Commercial/test.xlsx",  # depth 3
            "2026/ProjectABC/01-Commercial/subfolder/test.xlsx",  # depth 4
        )
        result = _find_workbook_in_rfqs("test.xlsx", self.base_path)
        self.assertEqual(result, self.base_path / "2026/ProjectABC")

    def test_searches_multiple_years(self):
        """Should search across multiple year folders."""
        self._create_structure("2025/OldProject/01-Commercial/legacy.xlsx")
        result = _find_workbook_in_rfqs("legacy.xlsx", self.base_path)
        self.assertEqual(result, self.base_path / "2025/OldProject/01-Commercial")

    def test_handles_empty_base_path(self):
        """Should return None when base path has no year folders."""
        result = _find_workbook_in_rfqs("test.xlsx", self.base_path)
        self.assertIsNone(result)

    def test_respects_max_depth(self):
        """Should not search beyond max depth of 5."""
        # Create file at depth 6 (beyond max)
        self._create_structure("2026/a/b/c/d/e/deep.xlsx")
        result = _find_workbook_in_rfqs("deep.xlsx", self.base_path)
        self.assertIsNone(result)

    def test_finds_at_max_depth(self):
        """Should find file at exactly max depth (5)."""
        self._create_structure("2026/a/b/c/d/file.xlsx")  # depth 5
        result = _find_workbook_in_rfqs("file.xlsx", self.base_path)
        self.assertEqual(result, self.base_path / "2026/a/b/c/d")


class TestSanitizeConfigString(unittest.TestCase):
    """Tests for sanitize_config_string function."""

    def test_removes_newlines(self):
        self.assertEqual(sanitize_config_string("Hello\nWorld"), "Hello World")

    def test_removes_carriage_returns(self):
        self.assertEqual(sanitize_config_string("Hello\rWorld"), "Hello World")

    def test_collapses_double_spaces(self):
        self.assertEqual(sanitize_config_string("Hello  World"), "Hello World")

    def test_strips_whitespace(self):
        self.assertEqual(sanitize_config_string("  Hello  "), "Hello")

    def test_handles_none(self):
        self.assertIsNone(sanitize_config_string(None))

    def test_handles_non_string(self):
        self.assertEqual(sanitize_config_string(123), 123)

    def test_combined(self):
        self.assertEqual(sanitize_config_string("  Hi\n  There  "), "Hi There")


class TestSanitizeConfigDate(unittest.TestCase):
    """Tests for sanitize_config_date function."""

    def test_datetime_object(self):
        self.assertEqual(sanitize_config_date(datetime(2024, 1, 15)), "2024-01-15")

    def test_iso_unchanged(self):
        self.assertEqual(sanitize_config_date("2024-01-15"), "2024-01-15")

    def test_iso_with_whitespace(self):
        self.assertEqual(sanitize_config_date("  2024-01-15  "), "2024-01-15")

    def test_european_format(self):
        self.assertEqual(sanitize_config_date("15/01/2024"), "2024-01-15")

    def test_handles_none(self):
        self.assertIsNone(sanitize_config_date(None))

    def test_handles_empty_string(self):
        self.assertEqual(sanitize_config_date(""), "")

    def test_handles_non_string(self):
        self.assertEqual(sanitize_config_date(12345), 12345)


if __name__ == "__main__":
    # Run tests with verbosity
    unittest.main(verbosity=2)
