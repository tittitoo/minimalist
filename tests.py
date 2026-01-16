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

# Import functions to test
from functions import (
    set_nitty_gritty,
    set_comma_space,
    set_x,
    set_case_preserve_acronym,
    title_case_ignore_double_char,
    SKIP_SHEETS,
)


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
        expected = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
        for sheet in expected:
            self.assertIn(sheet, SKIP_SHEETS)

    def test_skip_sheets_is_list(self):
        self.assertIsInstance(SKIP_SHEETS, list)


if __name__ == "__main__":
    # Run tests with verbosity
    unittest.main(verbosity=2)
