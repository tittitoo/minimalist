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
from unittest.mock import patch

# Import functions to test
from functions import (
    set_nitty_gritty,
    set_comma_space,
    set_paren_spacing,
    set_x,
    set_case_preserve_acronym,
    title_case_ignore_double_char,
    normalize_standard_tokens,
    set_dimension_unit_chain,
    format_description_text,
    collapse_spaced_cat_standard,
    set_degree_unit,
    set_spaced_voltage_type,
    expand_shorthand,
    _sp_wrap_lines,
    _SP_MDW_PX_WIN,
    _SP_MDW_PX_MAC,
    _SP_MDW_PX_PRINT,
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
    apply_scope_style,
)
from datetime import datetime

import functions


class WindowsWrapCalibrationMixin:
    """Pin _SP_MDW_PX to the Windows value for the duration of the test.

    _SP_MDW_PX is platform-dependent (Windows and Mac render different amounts of
    text per line — see the constant's comment in functions.py). Every wrap case
    pinned in this file was verified against a Windows-generated PDF, so they must
    be asserted against the Windows constant regardless of the machine running the
    suite, or the results would flip depending on the developer's OS.
    """

    def setUp(self):
        super().setUp()
        p = patch.object(functions, "_SP_MDW_PX", _SP_MDW_PX_WIN)
        p.start()
        self.addCleanup(p.stop)


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

    def test_inserts_space_when_token_before_comma_is_not_numeric(self):
        # A 3-digit follow-side alone isn't enough to call it a thousands
        # separator — the token before the comma must also end in a digit,
        # otherwise "Cable,450V" would wrongly be left unchanged.
        self.assertEqual(set_comma_space("Cable,450V"), "Cable, 450V")
        self.assertEqual(set_comma_space("Blue,112A"), "Blue, 112A")


class TestSetX(unittest.TestCase):
    """Tests for set_x function that normalizes quantity notation to 'N ×' format."""

    def test_number_first_lowercase(self):
        self.assertEqual(set_x("20x items"), "20 × items")

    def test_number_first_uppercase(self):
        self.assertEqual(set_x("30X items"), "30 × items")

    def test_symbol_first_lowercase_flipped(self):
        self.assertEqual(set_x("x20 items"), "20 × items")

    def test_symbol_first_uppercase_flipped(self):
        self.assertEqual(set_x("X30 items"), "30 × items")

    def test_number_first_with_space(self):
        self.assertEqual(set_x("20 X items"), "20 × items")

    def test_symbol_first_with_space_flipped(self):
        self.assertEqual(set_x("X 20 items"), "20 × items")

    def test_preserves_hyphenated(self):
        # 20x- should not be changed (hyphen follows)
        self.assertEqual(set_x("20x-connector"), "20x-connector")

    def test_preserves_x_as_last_letter_of_word(self):
        # 'x' ending a word (Max) is not a multiplication sign
        self.assertEqual(set_x("Max 11.7 watts"), "Max 11.7 watts")

    def test_preserves_x_as_last_letter_of_word_no_space(self):
        self.assertEqual(set_x("Flex 10G ports"), "Flex 10G ports")

    def test_preserves_cisco_part_number_digit_x_letter(self):
        # X directly followed by another letter is part of a part number
        self.assertEqual(
            set_x("WS-C2960X-24TS-L"), "WS-C2960X-24TS-L"
        )

    def test_preserves_cisco_part_number_x_digit_hyphen(self):
        # X followed by digits then a hyphen is part of a part number
        self.assertEqual(set_x("X2-10GB-SR"), "X2-10GB-SR")

    def test_preserves_cisco_part_number_letter_x_digit_letter(self):
        self.assertEqual(set_x("N9K-X9736C-EX"), "N9K-X9736C-EX")

    def test_preserves_part_number_ending_in_digits_plus_x(self):
        # Regression: the number-first regex only checked the single character
        # immediately before the digit run, so backtracking let it start matching
        # mid-token (the digit before "02X" is itself a digit, not a letter, so the
        # old lookbehind passed) — corrupting "LTD002X" into "LTD002 ×".
        self.assertEqual(set_x("LTD002X"), "LTD002X")

    def test_pads_already_present_multiplication_sign_glued_to_digit(self):
        # A "×" pasted from a supplier spec, glued to a digit on one side only,
        # is left untouched by every x/X-keyed pattern above — padded separately.
        self.assertEqual(set_x("4× 256"), "4 × 256")
        self.assertEqual(set_x("2×1.2"), "2 × 1.2")

    def test_preserves_nema_enclosure_rating_suffix(self):
        # "NEMA 4X" / "NEMA-4X" (enclosure rating — trailing X is a
        # corrosion-resistance suffix letter, not a multiplier) has the exact same
        # local shape as a real "4X" quantity (nothing glued after the X) — excluded
        # by name via lookbehind since there's no other way to tell them apart.
        self.assertEqual(set_x("NEMA 4X enclosure"), "NEMA 4X enclosure")
        self.assertEqual(set_x("NEMA-4X enclosure"), "NEMA-4X enclosure")
        # An ordinary multiplier right after "NEMA" text (not the enclosure rating
        # shape) still normalizes.
        self.assertEqual(set_x("4X zoom"), "4 × zoom")


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

    def test_dot_separated_part_number_preserved(self):
        # "BTD.NH.TE.TD.NR.RC" must not be lowercased to "BTD.nh.te.td.nr.rc".
        # capitalize() lowercases every char after the first, so restoration
        # must use IGNORECASE to recover inner dot-separated segments.
        result = set_case_preserve_acronym(
            "Part Number BTD.NH.TE.TD.NR.RC", title=True
        )
        self.assertIn("BTD.NH.TE.TD.NR.RC", result)

    def test_upper_case(self):
        result = set_case_preserve_acronym("hello world", upper=True)
        self.assertEqual(result, "HELLO WORLD")

    def test_title_case_preserves_ohm_symbol(self):
        # Ω is a cased Unicode letter; str.capitalize()/lower() silently
        # turn it into ω since it has no ASCII acronym-regex match to restore it.
        result = set_case_preserve_acronym(
            "Pigtail Cable RG214, 50Ω, 1m N-Male Crimp Connector 50Ω", title=True
        )
        self.assertIn("50Ω", result)
        self.assertNotIn("50ω", result)

    def test_title_case_preserves_diameter_and_delta_symbols(self):
        result = set_case_preserve_acronym("Diameter Φ50 mm Hole", title=True)
        self.assertIn("Φ50", result)
        result = set_case_preserve_acronym("Temperature Rise ΔT 10K", title=True)
        self.assertIn("ΔT", result)


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

    def test_apostrophe_possessive(self):
        # str.title() incorrectly capitalises the 's' after an apostrophe;
        # capitalize() must be used instead so "manufacturer's" stays lowercase.
        result = title_case_ignore_double_char("manufacturer's product")
        self.assertEqual(result, "Manufacturer's Product")


class TestSetParenSpacing(unittest.TestCase):
    """Tests for set_paren_spacing function."""

    def test_adds_space_before_paren(self):
        self.assertEqual(set_paren_spacing("unit(bracket)"), "unit (bracket)")

    def test_no_space_before_paren_at_start(self):
        self.assertEqual(set_paren_spacing("(bracket) unit"), "(bracket) unit")

    def test_hugs_following_punctuation(self):
        self.assertEqual(set_paren_spacing("unit(bracket),next"), "unit (bracket),next")

    def test_single_space_before_letter_or_digit(self):
        self.assertEqual(set_paren_spacing("unit(bracket)next"), "unit (bracket) next")


class TestNormalizeStandardTokens(unittest.TestCase):
    """Tests for normalize_standard_tokens: UOM and standard-designator normalization.

    Rules ported from the `hote` web app's normalizeStandardTokens() so free-text
    descriptions normalize the same way across both tools.
    """

    def test_normalizes_units_regardless_of_source_casing(self):
        self.assertEqual(normalize_standard_tokens("27MM bracket"), "27 mm bracket")
        self.assertEqual(normalize_standard_tokens("100 FT cable"), "100 ft cable")

    def test_normalizes_kw_both_casings(self):
        self.assertEqual(normalize_standard_tokens("5kw supply"), "5 kW supply")
        self.assertEqual(normalize_standard_tokens("5KW supply"), "5 kW supply")

    def test_normalizes_voltage_units(self):
        self.assertEqual(normalize_standard_tokens("230Vac supply"), "230 VAC supply")
        self.assertEqual(normalize_standard_tokens("24vdc unit"), "24 VDC unit")

    def test_normalizes_ah_without_colliding_with_bare_a(self):
        self.assertEqual(normalize_standard_tokens("205ah battery"), "205 Ah battery")

    def test_normalizes_month_hour_year_to_database_standard_codes(self):
        self.assertEqual(normalize_standard_tokens("36mths support"), "36 mth support")
        self.assertEqual(normalize_standard_tokens("36 months support"), "36 mth support")
        self.assertEqual(normalize_standard_tokens("2hrs labour"), "2 hr labour")
        self.assertEqual(normalize_standard_tokens("3 years warranty"), "3 yr warranty")

    def test_normalizes_rack_units_with_no_space(self):
        self.assertEqual(normalize_standard_tokens("2u rack"), "2U rack")

    def test_converts_area_volume_exponent_to_superscript(self):
        self.assertEqual(normalize_standard_tokens("10mm2 wire"), "10 mm² wire")
        self.assertEqual(normalize_standard_tokens("5m3 tank"), "5 m³ tank")

    def test_normalizes_cat_family_standards(self):
        self.assertEqual(normalize_standard_tokens("cat6a patch cord"), "Cat6A patch cord")
        self.assertEqual(normalize_standard_tokens("CAT6A patch cord"), "Cat6A patch cord")

    def test_collapses_spaced_cat_standard(self):
        # collapse_spaced_cat_standard only removes the internal spacing/punctuation —
        # it preserves whatever casing the input had, since normalize_standard_tokens
        # (which consumes its output) matches case-insensitively.
        self.assertEqual(collapse_spaced_cat_standard("Cat. 6 A patch cord"), "cat6A patch cord")
        self.assertEqual(collapse_spaced_cat_standard("Cat 6A patch cord"), "cat6A patch cord")
        self.assertEqual(collapse_spaced_cat_standard("cat. 5 e cable"), "cat5e cable")
        self.assertEqual(collapse_spaced_cat_standard("Cat 6 keystone"), "cat6 keystone")

    def test_format_description_text_normalizes_spaced_cat_standard(self):
        self.assertEqual(format_description_text("Cat. 6 A patch cord"), "Cat6A patch cord")
        self.assertEqual(format_description_text("Cat 6 A patch cord"), "Cat6A patch cord")

    def test_normalizes_ip_ratings(self):
        self.assertEqual(normalize_standard_tokens("ip65 rated"), "IP65 rated")
        self.assertEqual(normalize_standard_tokens("ipx6 rated"), "IPX6 rated")
        self.assertEqual(normalize_standard_tokens("ip69k washdown"), "IP69K washdown")

    def test_normalizes_spelled_out_meter_keeping_space(self):
        # Space preserved like every other unit here — a prior version folded
        # meter/metre into their own space-collapsing dict, inconsistent with
        # mm/kg/Hz below.
        self.assertEqual(normalize_standard_tokens("40 meter"), "40 m")
        self.assertEqual(normalize_standard_tokens("0.2 meter cable"), "0.2 m cable")
        self.assertEqual(normalize_standard_tokens("40 Metre"), "40 m")

    def test_normalizes_spelled_out_ohm_to_symbol_keeping_space(self):
        self.assertEqual(normalize_standard_tokens("50 ohm resistor"), "50 Ω resistor")
        self.assertEqual(normalize_standard_tokens("50 Ohms"), "50 Ω")

    def test_normalizes_already_literal_ohm_symbol_spacing(self):
        # A "50Ω" pasted straight from a datasheet wasn't getting the space
        # enforced before — only the spelled-out "ohm" word was.
        self.assertEqual(normalize_standard_tokens("50Ω resistor"), "50 Ω resistor")

    def test_does_not_corrupt_ordinary_words_with_dictionary_substrings(self):
        self.assertEqual(
            normalize_standard_tokens("format the description"), "format the description"
        )


class TestSetDimensionUnitChain(unittest.TestCase):
    """Tests for set_dimension_unit_chain: glued multi-number dimension chains."""

    def test_normalizes_glued_dimension_chain(self):
        self.assertEqual(set_dimension_unit_chain("600x746x673mm"), "600 × 746 × 673 mm")
        self.assertEqual(set_dimension_unit_chain("600X746X673MM"), "600 × 746 × 673 mm")

    def test_leaves_single_unit_mention_untouched(self):
        self.assertEqual(set_dimension_unit_chain("27mm bracket"), "27mm bracket")


class _MockScopeFont:
    """Records bold/color assignments for one range address."""

    def __init__(self, sink, addr):
        self._sink = sink
        self._addr = addr

    @property
    def bold(self):
        raise NotImplementedError

    @bold.setter
    def bold(self, value):
        self._sink.append((self._addr, "bold", value))

    @property
    def color(self):
        raise NotImplementedError

    @color.setter
    def color(self, value):
        self._sink.append((self._addr, "color", value))


class _MockScopeRange:
    def __init__(self, addr, value=None, sink=None):
        self.addr = addr
        self._value = value
        self.font = _MockScopeFont(sink if sink is not None else [], addr)

    @property
    def value(self):
        return self._value

    def end(self, direction):
        return self


class MockScopeSheet:
    """Minimal sheet double supporting only the .range(...) calls apply_scope_style
    makes, so its batching logic is testable without Excel."""

    def __init__(self, h_values, al_values=None):
        self.h_values = h_values
        # Defaults to "Title" for every row when omitted, so tests that don't care
        # about the AL split still exercise the bold path.
        self.al_values = al_values if al_values is not None else ["Title"] * len(h_values)
        self.calls = []  # (addr, prop, value) in call order

    def range(self, addr):
        last_row = len(self.h_values) + 2
        if addr == "C1500":
            r = _MockScopeRange(addr)
            r.row = last_row
            return r
        if addr == f"H3:H{last_row}":
            return _MockScopeRange(addr, value=list(self.h_values))
        if addr == f"AL3:AL{last_row}":
            return _MockScopeRange(addr, value=list(self.al_values))
        return _MockScopeRange(addr, sink=self.calls)


class TestApplyScopeStyle(unittest.TestCase):
    """Tests for apply_scope_style's per-Scope-value coloring and Title-only bold."""

    def test_batches_contiguous_runs_and_sets_bold_blue_for_option(self):
        # H3="", H4="OPTION", H5="OPTION", H6="", H7="OPTION" — all Title rows here,
        # so this only exercises the option/non-option split, not the AL-bold split.
        sheet = MockScopeSheet(["", "OPTION", "OPTION", "", "OPTION"])
        apply_scope_style(sheet)

        bold_true = {addr for addr, prop, val in sheet.calls if prop == "bold" and val is True}
        bold_false = {addr for addr, prop, val in sheet.calls if prop == "bold" and val is False}
        color_calls = {addr: val for addr, prop, val in sheet.calls if prop == "color"}

        self.assertEqual(bold_true, {"H4:H5", "H7:H7"})
        self.assertEqual(bold_false, {"H3:H3", "H6:H6"})
        self.assertEqual(color_calls["H4:H5"], (4, 50, 255))
        self.assertEqual(color_calls["H7:H7"], (4, 50, 255))
        self.assertEqual(color_calls["H3:H3"], (0, 0, 0))
        self.assertEqual(color_calls["H6:H6"], (0, 0, 0))

    def test_bolds_option_only_on_title_rows_not_sub_item_rows(self):
        # H3="OPTION" on a Title row (bold+blue); H4="OPTION" on a Description
        # sub-item row (blue only, not bold) — mirrors the real BOQ layout where an
        # OPTION Title has an OPTION Description nested under it.
        sheet = MockScopeSheet(
            h_values=["OPTION", "OPTION"],
            al_values=["Title", "Description"],
        )
        apply_scope_style(sheet)

        bold_calls = {addr: val for addr, prop, val in sheet.calls if prop == "bold"}
        color_calls = {addr: val for addr, prop, val in sheet.calls if prop == "color"}

        self.assertEqual(bold_calls["H3:H3"], True)
        self.assertEqual(bold_calls["H4:H4"], False)
        self.assertEqual(color_calls["H3:H3"], (4, 50, 255))
        self.assertEqual(color_calls["H4:H4"], (4, 50, 255))

    def test_colors_each_scope_value_distinctly(self):
        sheet = MockScopeSheet(
            h_values=["INCLUDED", "WAIVED", "TBA"],
            al_values=["Description", "Description", "Description"],
        )
        apply_scope_style(sheet)
        color_calls = {addr: val for addr, prop, val in sheet.calls if prop == "color"}
        self.assertEqual(color_calls["H3:H3"], (0, 128, 0))
        self.assertEqual(color_calls["H4:H4"], (255, 140, 0))
        self.assertEqual(color_calls["H5:H5"], (127, 127, 127))

    def test_removed_gets_red_and_strikethrough_regardless_of_row_type(self):
        # REMOVED strikethrough applies on every row type (unlike bold, which is
        # Title-only) — it marks content as voided wherever it appears.
        sheet = MockScopeSheet(
            h_values=["REMOVED", "REMOVED"],
            al_values=["Title", "Description"],
        )
        with patch("functions.set_range_strikethrough") as mock_strike:
            apply_scope_style(sheet)

        strike_calls = {c.args[0].addr: c.args[1] for c in mock_strike.call_args_list}
        self.assertEqual(strike_calls, {"H3:H3": True, "H4:H4": True})

        color_calls = {addr: val for addr, prop, val in sheet.calls if prop == "color"}
        bold_calls = {addr: val for addr, prop, val in sheet.calls if prop == "bold"}
        self.assertEqual(color_calls["H3:H3"], (192, 0, 0))
        self.assertEqual(color_calls["H4:H4"], (192, 0, 0))
        self.assertEqual(bold_calls["H3:H3"], True)
        self.assertEqual(bold_calls["H4:H4"], False)

    def test_non_removed_rows_get_strikethrough_cleared(self):
        sheet = MockScopeSheet(h_values=["OPTION"], al_values=["Title"])
        with patch("functions.set_range_strikethrough") as mock_strike:
            apply_scope_style(sheet)
        mock_strike.assert_called_once_with(mock_strike.call_args.args[0], False)

    def test_no_op_when_sheet_has_no_data_rows(self):
        sheet = MockScopeSheet([])
        apply_scope_style(sheet)
        self.assertEqual(sheet.calls, [])


class TestSetDegreeUnit(unittest.TestCase):
    """Tests for set_degree_unit: spelled-out Deg C/F and literal ° symbol spacing."""

    def test_normalizes_spelled_out_deg_c(self):
        self.assertEqual(set_degree_unit("-40 Deg C"), "-40 °C")
        self.assertEqual(set_degree_unit("55 Deg F"), "55 °F")

    def test_normalizes_already_glued_degree_symbol(self):
        self.assertEqual(set_degree_unit("55°C"), "55 °C")

    def test_normalizes_stray_spacing_around_degree_symbol(self):
        self.assertEqual(set_degree_unit("55 ° C"), "55 °C")

    def test_leaves_bare_deg_with_no_value_alone(self):
        self.assertEqual(set_degree_unit("Deg C rating"), "Deg C rating")


class TestSetSpacedVoltageType(unittest.TestCase):
    """Tests for set_spaced_voltage_type: collapsing spaced "V AC"/"V DC" into VAC/VDC."""

    def test_collapses_spaced_v_ac(self):
        self.assertEqual(set_spaced_voltage_type("110 V AC supply"), "110 VAC supply")

    def test_collapses_spaced_v_dc(self):
        self.assertEqual(set_spaced_voltage_type("24 V DC unit"), "24 VDC unit")

    def test_leaves_already_glued_form_untouched(self):
        self.assertEqual(set_spaced_voltage_type("230VAC supply"), "230VAC supply")


class TestExpandShorthand(unittest.TestCase):
    """Tests for expand_shorthand: c/w, w/, w/o, Equiv, Incl spec-sheet shorthand."""

    def test_expands_c_w(self):
        self.assertEqual(
            expand_shorthand("Bracket c/w mounting screws"),
            "Bracket complete with mounting screws",
        )

    def test_expands_w_slash_glued_to_next_word(self):
        self.assertEqual(expand_shorthand("MT74H52A w/FLX2 cable"), "MT74H52A with FLX2 cable")

    def test_does_not_expand_w_slash_followed_by_digit(self):
        # That shape is dimension-chain notation (e.g. "W/800"), not the shorthand.
        self.assertEqual(expand_shorthand("D/1200 x W/800 x H/2100"), "D/1200 x W/800 x H/2100")

    def test_expands_equiv_and_incl(self):
        self.assertEqual(expand_shorthand("Equiv. to OEM part"), "equivalent. to OEM part")
        self.assertEqual(expand_shorthand("Incl: mounting kit"), "including: mounting kit")


class TestFormatDescriptionText(unittest.TestCase):
    """Tests for format_description_text, the combined cleanup/normalization/title-case
    pipeline ported from the `hote` web app's formatDescriptionText(), so free-text
    descriptions format identically across both tools.
    """

    def test_title_cases_short_text(self):
        self.assertEqual(
            format_description_text("cable tray for antenna", title_case=True),
            "Cable Tray for Antenna",
        )

    def test_leaves_long_text_unchanged_by_title_casing(self):
        long_text = (
            "this is a very long description that definitely exceeds one "
            "hundred characters in length, well past it"
        )
        self.assertGreater(len(long_text), 100)
        self.assertEqual(format_description_text(long_text, title_case=True), long_text)

    def test_title_cases_label_style_text_up_to_100_chars(self):
        # Comma-separated attribute lists (not prose) stay title-cased up to the
        # raised 100-char cutoff — data-driven from the products catalog, where
        # label-style names top out around 94 chars.
        self.assertEqual(
            format_description_text(
                "Motorola XiR P6600i NON-TIA (No-Display, No-Keypad, Non-I.S) "
                "Portable Radio - APAC model",
                title_case=True,
            ),
            "Motorola XiR P6600i NON-TIA (No-Display, No-Keypad, Non-I.S) "
            "Portable Radio - APAC Model",
        )

    def test_preserves_hyphenated_part_numbers_under_title_case(self):
        self.assertEqual(
            format_description_text("switch WS-C2960X-24TS-L unit", title_case=True),
            "Switch WS-C2960X-24TS-L Unit",
        )

    def test_preserves_hyphen_flanked_sku_segment_colliding_with_article(self):
        # "-A" reduces to "a" (normally lowercased mid-string) but is a part-number
        # segment here, not the English article, and must survive untouched.
        self.assertEqual(
            format_description_text(
                "Catalyst 9300 24-port PoE+, Network Advantage · C9300-24P-A",
                title_case=True,
            ),
            "Catalyst 9300 24-port PoE+, Network Advantage · C9300-24P-A",
        )

    def test_does_not_bleed_allcaps_sku_casing_into_unrelated_word(self):
        self.assertEqual(
            format_description_text(
                "Single Pack Option · CW9174I-SINGLE", title_case=True
            ),
            "Single Pack Option · CW9174I-SINGLE",
        )

    def test_normalizes_quantity_x_notation(self):
        self.assertEqual(
            format_description_text("20x cable ties", title_case=True), "20 × Cable Ties"
        )

    def test_keeps_value_unit_x_count_suffix_in_order(self):
        # Regression: the leading-multiplier pattern (for "x2 items" -> "2 × Items")
        # used to blindly match just the "x 2" fragment here too, discarding "8GB"
        # entirely and reordering into "2 ×" — producing "8GB 2 ×" instead of keeping
        # the value and count together as "8GB × 2".
        self.assertEqual(
            format_description_text("Memory 8GB x 2, total 16GB", title_case=True),
            "Memory 8 GB × 2, Total 16 GB",
        )
        self.assertEqual(
            format_description_text("1080p and 4K x 2K at 60Hz", title_case=True),
            "1080p and 4K × 2K at 60 Hz",
        )

    def test_does_not_misread_qty_unit_bullet_separator_as_leading_multiplier(self):
        # Regression: the leading-multiplier pattern (for "x 2 items" -> "2 × Items")
        # used to fire on "lot x 6-Way ..." too, reading the "6" as the multiplier
        # count and reordering into "1 lot 6 ×-Way Universal PDU" — corrupting a real
        # catalog bullet.
        self.assertEqual(
            format_description_text("1 lot x 6-Way Universal PDU", title_case=True),
            "1 lot x 6-Way Universal PDU",
        )
        self.assertEqual(
            format_description_text("x 2 items", title_case=True), "2 × Items"
        )

    def test_normalizes_qty_unit_codes_glued_to_a_number(self):
        # ea/set/lot/trp/md are the app's established qty-unit codes (see the UNITS
        # constants in the `hote` web app) — glued to a number the same spacing rule
        # applies as any other unit of measure.
        self.assertEqual(
            format_description_text("1lot x Cable Management Unit", title_case=True),
            "1 lot x Cable Management Unit",
        )
        self.assertEqual(
            format_description_text("1lot x 6-Way Universal PDU", title_case=True),
            "1 lot x 6-Way Universal PDU",
        )
        self.assertEqual(
            format_description_text("2ea spare fuses", title_case=True),
            "2 ea Spare Fuses",
        )
        self.assertEqual(
            format_description_text("3sets of connectors", title_case=True),
            "3 set of Connectors",
        )
        self.assertEqual(
            format_description_text("5md assembly", title_case=True), "5 md Assembly"
        )
        self.assertEqual(
            format_description_text("1trp site survey", title_case=True),
            "1 trp Site Survey",
        )

    def test_keeps_qty_unit_codes_lowercase_mid_string_even_when_already_spaced(self):
        # Regression: title-casing ran before normalize_standard_tokens, so an
        # already-spaced unit word like "1 lot x ..." briefly became "1 Lot x ..." —
        # usually corrected back by normalize_standard_tokens' case-insensitive pass,
        # but now held lowercase directly by the title-casing step itself as a stated
        # exception, same as "a/of/to/...".
        self.assertEqual(
            format_description_text("1 lot x 6-Way Universal PDU", title_case=True),
            "1 lot x 6-Way Universal PDU",
        )
        self.assertEqual(
            format_description_text("2 ea spare fuses", title_case=True),
            "2 ea Spare Fuses",
        )
        self.assertEqual(
            format_description_text(
                "5 md assembly of connectors", title_case=True
            ),
            "5 md Assembly of Connectors",
        )

    def test_normalizes_asterisk_multiplier_without_corrupting_markdown_italics(self):
        # A bare "*" is a CommonMark emphasis delimiter — left untouched, a second
        # "*" later in the same string (e.g. a repeated multiplier) would italicize
        # everything between the two, not just the multiplier itself.
        self.assertEqual(
            format_description_text(
                "2*200G/400G board (2*100G capacity included)", title_case=True
            ),
            "2 × 200G/400G Board (2 × 100G Capacity Included)",
        )

    def test_fixes_comma_spacing_without_breaking_thousands_separator(self):
        self.assertEqual(
            format_description_text(
                "cable ,connector and 1,200 units", title_case=True
            ),
            "Cable, Connector and 1,200 Units",
        )

    def test_inserts_comma_space_even_after_a_digit_that_is_not_a_thousands_separator(self):
        self.assertEqual(
            format_description_text(
                "9172H(W7,3 radio,3 band 2x2),Global", title_case=True
            ),
            "9172H (W7, 3 Radio, 3 Band 2x2), Global",
        )

    def test_adds_paren_spacing_and_capitalizes_past_leading_punctuation(self):
        self.assertEqual(
            format_description_text("unit(bracket),next", title_case=True),
            "Unit (Bracket), Next",
        )

    def test_normalizes_double_single_quote_inch_mark(self):
        self.assertEqual(
            format_description_text("Storage 2.5'' 1TB SSD", title_case=True),
            'Storage 2.5" 1 TB SSD',
        )
        # A genuine double-quote inch mark elsewhere is untouched.
        self.assertEqual(
            format_description_text('27" Monitor', title_case=True), '27" Monitor'
        )

    def test_strips_optional_plural_paren_after_any_word(self):
        self.assertEqual(
            format_description_text(
                "2 COM Port(s), 1 VGA port(s)", title_case=True
            ),
            "2 COM Port, 1 VGA Port",
        )

    def test_expands_with_shorthand(self):
        self.assertEqual(
            format_description_text(
                "Air-conditioner unit w/ duct kit", title_case=True
            ),
            "Air-conditioner Unit With Duct Kit",
        )
        self.assertEqual(
            format_description_text(
                "Enclosure w/o External JB", title_case=True
            ),
            "Enclosure Without External JB",
        )
        # Must not corrupt a genuine part number/fraction with a bare "w" and slash.
        self.assertEqual(
            format_description_text('27" Monitor', title_case=True), '27" Monitor'
        )

    def test_expands_c_w_equiv_incl_shorthand(self):
        self.assertEqual(
            format_description_text(
                "Bracket c/w mounting screws", title_case=True
            ),
            "Bracket Complete With Mounting Screws",
        )
        self.assertEqual(
            format_description_text("Equiv. to OEM part", title_case=True),
            "Equivalent. to OEM Part",
        )
        self.assertEqual(
            format_description_text("Incl: mounting kit", title_case=True),
            "Including: Mounting Kit",
        )

    def test_expands_shorthand_glued_directly_to_next_word(self):
        # "w/FLX2 Cable" (no space after the slash — common in real product names,
        # e.g. "MT74H52A w/FLX2 cable") must still expand; only a following DIGIT is
        # excluded (that shape is dimension-chain notation, e.g. "W/800", handled
        # separately by protect_dimension_suffix_chains).
        self.assertEqual(
            format_description_text("MT74H52A w/FLX2 cable", title_case=True),
            "MT74H52A With FLX2 Cable",
        )

    def test_shorthand_expansion_trims_trailing_space_at_end_of_string(self):
        # expand_shorthand's slash-form replacements always end in a space (needed
        # to properly separate a glued-on following word) — must be trimmed back off
        # if the source string ends right on one of those forms with nothing after.
        self.assertEqual(format_description_text("w/", title_case=True), "With")

    def test_normalizes_dimension_letter_chain_without_misreading_w_as_watts(self):
        self.assertEqual(
            format_description_text(
                "Cabinet 800W X 1200D X 2100H", title_case=True
            ),
            "Cabinet 800W × 1200D × 2100H",
        )
        # Lowercase dimension letters must survive title-casing untouched — this only
        # holds if the chain is protected from title-casing, not merely normalized
        # afterwards.
        self.assertEqual(
            format_description_text("800w x 1200d x 2100h", title_case=True),
            "800w × 1200d × 2100h",
        )

    def test_lone_dimension_letter_is_still_read_as_its_unit(self):
        self.assertEqual(
            format_description_text("800w fan", title_case=True), "800 W Fan"
        )

    def test_normalizes_naked_dimension_chain_with_no_trailing_unit(self):
        # A chain of 3+ numbers multiplied together is unambiguous even with no
        # trailing unit at all (e.g. a junction box's "160x160x91", implied mm).
        self.assertEqual(
            format_description_text("JB (160x160x91)", title_case=True),
            "JB (160 × 160 × 91)",
        )
        self.assertEqual(
            format_description_text("Junction box (160X160X91)", title_case=True),
            "Junction Box (160 × 160 × 91)",
        )
        # A bare two-number chain stays untouched — same ambiguity as set_x's own
        # guard (could be a resolution or a part-number-style code).
        self.assertEqual(
            format_description_text("20x30 enclosure", title_case=True),
            "20x30 Enclosure",
        )

    def test_normalizes_range_tilde_to_en_dash(self):
        # A tilde used as a numeric "to" range separator would otherwise corrupt
        # markdown rendering (GFM reads a tilde pair as strikethrough) in `hote`'s
        # preview — normalized here too so both tools agree on the output text.
        self.assertEqual(
            format_description_text("Gain 20~31dB, Max 21.5dBm", title_case=True),
            "Gain 20–31 dB, Max 21.5 dBm",
        )
        self.assertEqual(
            format_description_text("190.65THz~196.675THz range", title_case=True),
            "190.65THz–196.675THz Range",
        )

    def test_converts_caret_exponent_notation_to_superscript(self):
        self.assertEqual(
            format_description_text("25mm^2 wire", title_case=True), "25 mm² Wire"
        )
        self.assertEqual(
            format_description_text("5m^3 tank", title_case=True), "5 m³ Tank"
        )

    def test_strips_parenthesized_plural_marker_when_normalizing_month(self):
        self.assertEqual(
            format_description_text("60Month(s) support", title_case=True),
            "60 mth Support",
        )
        # Glued to a preceding underscore (a field-delimiter artifact in some
        # imported data) must still be recognized.
        self.assertEqual(
            format_description_text("Basic Chassis_60Month(s)", title_case=True),
            "Basic Chassis_60 mth",
        )

    def test_does_not_treat_comma_before_nonnumeric_token_as_thousands_separator(self):
        self.assertEqual(
            format_description_text(
                "Power Cable,450V/750V,25mm^2,Blue,112A,CCC,CE", title_case=True
            ),
            "Power Cable, 450 V/750 V, 25 mm², Blue, 112 A, CCC, CE",
        )

    def test_normalizes_bit_rate_slash_notation_distinct_from_byte_storage(self):
        # "Xb/s" (bit rate) looks identical to "XB" (byte storage) once case is
        # folded — the trailing "/s" is the only signal distinguishing the two.
        self.assertEqual(
            format_description_text("8.5Gb/s-11.1Gb/s with CDR", title_case=True),
            "8.5 Gb/s-11.1 Gb/s With CDR",
        )
        self.assertEqual(
            format_description_text("10Mb/s uplink", title_case=True),
            "10 Mb/s Uplink",
        )
        self.assertEqual(
            format_description_text("1.5tb/s backbone", title_case=True),
            "1.5 Tb/s Backbone",
        )
        # Byte storage and bit-rate-via-slash must resolve independently in the
        # same string.
        self.assertEqual(
            format_description_text(
                "500gb storage over an 8gb/s link", title_case=True
            ),
            "500 GB Storage Over an 8 Gb/s Link",
        )

    def test_normalizes_flashes_per_minute_unit(self):
        self.assertEqual(
            format_description_text("120FPM xenon beacon", title_case=True),
            "120 fpm Xenon Beacon",
        )
        self.assertEqual(
            format_description_text("60fpm xenon beacon", title_case=True),
            "60 fpm Xenon Beacon",
        )

    def test_inserts_space_in_glued_ex_protection_type_markings(self):
        cases = [
            ("Exd IIC T6 enclosure", "Ex d IIC T6 Enclosure"),
            ("Exde IIC T4 junction box", "Ex de IIC T4 Junction Box"),
            ("Exeb IIC T6 terminal box", "Ex eb IIC T6 Terminal Box"),
            ("Exdb IIC Gb rated", "Ex db IIC Gb Rated"),
            ("Exia IIC T4 barrier", "Ex ia IIC T4 Barrier"),
            # Already spaced — must pass through unchanged.
            ("Ex d IIC T6 enclosure", "Ex d IIC T6 Enclosure"),
            ("Ex db eb IIC T6 Gb", "Ex db eb IIC T6 Gb"),
        ]
        for text, expected in cases:
            self.assertEqual(format_description_text(text, title_case=True), expected)

    def test_ex_protection_spacing_does_not_corrupt_ordinary_ex_words(self):
        cases = [
            ("Express delivery available", "Express Delivery Available"),
            ("Extra bracket included", "Extra Bracket Included"),
            ("Extreme temperature rating", "Extreme Temperature Rating"),
            ("Exempt from certification", "Exempt From Certification"),
            ("Exercise caution", "Exercise Caution"),
            ("Exodus of legacy units", "Exodus of Legacy Units"),
            ("Exotic materials used", "Exotic Materials Used"),
            ("Exhaust fan included", "Exhaust Fan Included"),
            ("Exist in inventory", "Exist in Inventory"),
            ("Exact dimensions given", "Exact Dimensions Given"),
        ]
        for text, expected in cases:
            self.assertEqual(format_description_text(text, title_case=True), expected)

    def test_preserves_atex_iecex_certificate_numbers_ending_in_x(self):
        cases = [
            (
                "ATEX ITS18ATEX103970X / IECEx ITS 18.0052X · Ex db op is IIC T6 Gb",
                "ATEX ITS18ATEX103970X / IECEx ITS 18.0052X · Ex db op is IIC T6 Gb",
            ),
            (
                "Atex Certificate: SIRA06ATEX1097X",
                "Atex Certificate: SIRA06ATEX1097X",
            ),
            (
                "IECEX Certifcate: IECEx CML 18.0177X, IECEx SIM 15.0002X",
                "IECEX Certifcate: IECEx CML 18.0177X, IECEx SIM 15.0002X",
            ),
        ]
        for text, expected in cases:
            self.assertEqual(format_description_text(text, title_case=True), expected)

    def test_normalizes_microsecond_unit_both_micro_sign_variants(self):
        # µ (MICRO SIGN U+00B5) and μ (GREEK SMALL LETTER MU U+03BC) look identical
        # but are different codepoints — both must canonicalize to the same form.
        self.assertEqual(
            format_description_text("inrush 70A/120µs", title_case=True),
            "Inrush 70 A/120 µs",
        )
        self.assertEqual(
            format_description_text("inrush 70A/120μs", title_case=True),
            "Inrush 70 A/120 µs",
        )

    def test_normalizes_bare_micron_unit_without_colliding_with_microseconds(self):
        # "50µ" (coating/anodizing thickness) is distinct from "50µs" (microseconds,
        # tested above) — the trailing boundary check keeps the two units from
        # colliding regardless of dict iteration order.
        self.assertEqual(
            format_description_text(
                "anodized aluminum 50µ, PMMA lens", title_case=True
            ),
            "Anodized Aluminum 50 µ, PMMA Lens",
        )
        self.assertEqual(
            format_description_text(
                "anodized aluminum 50μ, PMMA lens", title_case=True
            ),
            "Anodized Aluminum 50 µ, PMMA Lens",
        )

    def test_normalizes_nautical_mile_unit(self):
        # Uppercase "NM" is unambiguous. Lowercase "nm" is disambiguated from the
        # nanometre unit by magnitude: visibility ratings are 1-2 digits, while
        # wavelength specs (nanometres) are always 3 digits and must stay untouched.
        self.assertEqual(
            format_description_text(
                "6nm dbl masthead, black anodized", title_case=True
            ),
            "6 NM Dbl Masthead, Black Anodized",
        )
        self.assertEqual(
            format_description_text("3nm 225°", title_case=True), "3 NM 225°"
        )
        self.assertEqual(
            format_description_text("visibility 6nm", title_case=True),
            "Visibility 6 NM",
        )
        self.assertEqual(
            format_description_text(
                "Light Colour: Green, 530nm", title_case=True
            ),
            "Light Colour: Green, 530nm",
        )
        self.assertEqual(
            format_description_text("1550nm fiber wavelength", title_case=True),
            "1550nm Fiber Wavelength",
        )

    def test_normalizes_hectopascal_unit(self):
        self.assertEqual(
            format_description_text("500-1100 hpa pressure range", title_case=True),
            "500-1100 hPa Pressure Range",
        )

    def test_normalizes_knots_and_mph_units(self):
        self.assertEqual(
            format_description_text("wind speed 40kt gust", title_case=True),
            "Wind Speed 40 kt Gust",
        )
        self.assertEqual(
            format_description_text("40kts gust", title_case=True), "40 kt Gust"
        )
        self.assertEqual(
            format_description_text("speed 30knots", title_case=True),
            "Speed 30 kt",
        )
        self.assertEqual(
            format_description_text("60mph rated", title_case=True), "60 mph Rated"
        )

    def test_normalizes_milliwatt_without_colliding_with_megawatt(self):
        # "mW" (milliwatt) must be spaced, but the case-sensitive match must not
        # touch a genuine "MW" (megawatt) spec — six orders of magnitude apart.
        self.assertEqual(
            format_description_text(
                "275 mW average (10 W peak)", title_case=True
            ),
            "275 mW Average (10 W Peak)",
        )
        self.assertEqual(
            format_description_text("5 MW generator", title_case=True),
            "5 MW Generator",
        )

    def test_normalizes_spelled_out_degree_unit(self):
        # A leading "-" is not used here (it would otherwise trigger the unrelated
        # leading-dash-to-bullet-marker rule at the very start of the pipeline).
        self.assertEqual(
            format_description_text(
                "Operating range: -40 Deg C to +55 Deg C", title_case=False
            ),
            "Operating range: -40 °C to +55 °C",
        )

    def test_normalizes_already_literal_degree_symbol_spacing(self):
        # A source pasted straight from a datasheet may already contain "°C" glued
        # with no space, or stray spacing like "55 ° C" — both forms are rewritten
        # to the same canonical "N °C" spacing, not just the spelled-out "Deg C".
        self.assertEqual(
            format_description_text("55°C rated", title_case=True), "55 °C Rated"
        )
        self.assertEqual(
            format_description_text("55 ° C rated", title_case=True),
            "55 °C Rated",
        )

    def test_collapses_spaced_voltage_type_to_vac_vdc(self):
        # The industrial-standard symbol is the combined "VAC"/"VDC" (no internal
        # space), but a source description may have "V" and "AC"/"DC" typed as
        # separate words.
        self.assertEqual(
            format_description_text(
                "110/220 V AC to 24 V DC", title_case=True
            ),
            "110/220 VAC to 24 VDC",
        )

    def test_normalizes_ma_mah_spacing(self):
        self.assertEqual(
            format_description_text("500mA draw", title_case=True), "500 mA Draw"
        )
        self.assertEqual(
            format_description_text("2075mAh capacity", title_case=True),
            "2075 mAh Capacity",
        )

    def test_normalizes_candela_unit(self):
        self.assertEqual(
            format_description_text("32 Cd beacon", title_case=True),
            "32 cd Beacon",
        )

    def test_normalizes_minute_unit(self):
        self.assertEqual(
            format_description_text("15 min assembly", title_case=True),
            "15 min Assembly",
        )
        self.assertEqual(
            format_description_text("15 minutes assembly", title_case=True),
            "15 min Assembly",
        )

    def test_strips_trailing_comma(self):
        # "LSOH" -> "LSZH" here too since standardize_lsoh_acronym runs in the same
        # pipeline — confirms trailing-comma stripping composes correctly with it.
        self.assertEqual(
            format_description_text(
                "Cat6 UTP Patch Cord, LSOH, 1 m Length, 4P,", title_case=True
            ),
            "Cat6 UTP Patch Cord, LSZH, 1 m Length, 4P",
        )

    def test_always_normalizes_units_regardless_of_title_case_flag(self):
        # Units/standards normalize even when title_case=False (long text, or a
        # Format type that's never title-cased).
        self.assertEqual(
            format_description_text("27MM bracket", title_case=False), "27 mm bracket"
        )

    def test_returns_empty_string_for_falsy_input(self):
        self.assertEqual(format_description_text(None, title_case=True), "")
        self.assertEqual(format_description_text("", title_case=True), "")


class TestSpWrapLinesRealPdfRegression(WindowsWrapCalibrationMixin, unittest.TestCase):
    """Regression cases for _sp_wrap_lines pinned against real generated PDFs.

    _SP_MDW_PX predicts how many lines Excel's actual print/export renderer
    (Windows: "Microsoft Print to PDF", fixed at 600 DPI) will wrap a
    Description cell to — it does NOT predict on-screen/autofit wrapping,
    which is display-scaling-dependent and varies per machine. Every case
    here was confirmed against an actual generated Commercial/Technical PDF
    (col_width=55/60, Arial 12pt), not just eyeballed in Excel. When a real
    PDF shows a new phantom-blank-line or clipped-text case, add it here
    with the source file/row noted, so recalibrating _SP_MDW_PX later can't
    silently regress a case already fixed.

    IMPORTANT: as of the J12632 recalibration (see _SP_MDW_PX and
    TestSpWrapLinesJ12632Regression below), several of these J12815-era
    "confirmed single line" cases now predict 2 lines (a phantom blank line)
    instead of 1. That's an accepted, deliberate trade-off — under-predicting
    a wrap silently drops real content (proven to happen with the old MDW),
    over-predicting just adds a harmless visible blank line. Do not "fix" this
    by raising MDW back up without re-proving it against the J12632 cases too.

    TREAT THESE AS SUSPECT, NOT AUTHORITATIVE. Their "true" line counts were read
    off `pdftotext -layout`, which reports hyphen-wrapped words and trailing
    punctuation as though text were missing — the exact false positives that drove
    several bad recalibrations. They are kept as change-detectors so a future MDW
    move is visible, not as evidence of correct rendering. The trustworthy pins are
    in TestSpWrapLinesGroundTruthCalibration, measured from real glyph geometry.
    """

    def test_confirmed_single_line_cases_do_not_get_a_phantom_second_line(self):
        # Commercial J12815 GEV CCTV R2, sheet CCTV, "Cisco IE-9320-24P4X-E"
        # line items — confirmed single-line in the real Windows Commercial PDF
        # (col_width=55), but mispredicted as 2 lines (phantom blank line)
        # under the old MDW=8.0.
        #
        # Under the safety-biased MDW=7.9 (post-J12632), 4 of these 5 now
        # predict 2 lines (phantom blank line) again — the same tension that
        # motivated raising MDW to 9.2 originally. This time it's accepted:
        # 9.2 was proven to silently drop real text in a different real
        # document (see TestSpWrapLinesJ12632Regression), which is worse than
        # a harmless extra blank line here.
        still_one_line = [
            "Software for Catalyst IE9300 Rugged Series · IE9300_sw",
        ]
        for text in still_one_line:
            self.assertEqual(_sp_wrap_lines(text, 55), 1, f"expected 1 line: {text!r}")

        now_phantom_two_lines = [
            "Cisco DNA Essentials License for IE9300 Series · IE9300-DNA-E",
            "IE 9300 DNA Essentials, 3 yr Term License · IE9300-DNA-E-3Y",
            "Digital Download Code for Software License · DIGITAL-DL-CODE",
            "Not Related to an IoT Solution; for Tracking Only. · IOT-OTHER",
        ]
        for text in now_phantom_two_lines:
            self.assertEqual(_sp_wrap_lines(text, 55), 2, f"expected 2 lines (accepted phantom line): {text!r}")

    def test_confirmed_two_line_cases_still_wrap(self):
        # Same source/sheet — confirmed genuinely 2 lines in the real Commercial
        # PDF (col_width=55), must not be pushed down to 1 line when
        # _SP_MDW_PX is raised to fix the phantom-line cases above.
        cases = [
            "24 Port PoE+ Downlinks With 4x10G Uplinks (720 W) · IE-9320-24P4X-E",
            "SNTC-8X5XNBD 24 Port PoE+ Downlinks With 4x10G Uplink, 36 mth · CON-SNT-IE932PXE",
            "Higher PoE, 400 W PSU for IE9300, 100-240 VAC/100-250 VDC · PWR-RGD-AC-DC-400",
            "Cisco CAB-STK-0.5 m 50 cm Stacking Cable for Catalyst IE9300 · CAB-STK-0.5 m",
            "Network Plug-n-Play Connect for Zero-touch Device Deployment · NETWORK-PNP-LIC",
        ]
        for text in cases:
            self.assertEqual(_sp_wrap_lines(text, 55), 2, f"expected 2 lines: {text!r}")

    def test_confirmed_at_technical_col_width_too(self):
        # Same workbook/rows, re-verified live at col_width=60 (the Technical
        # proposal width — wider than Commercial's 55, so a case that's 2
        # lines at 55 can legitimately become 1 line at 60; only re-assert
        # the subset actually re-checked at this width, not all 10 cases).
        #
        # The three "one_line" cases now predict 2 under the ground-truth-derived
        # MDW=7.50 — the same accepted phantom-line trade-off described in this
        # class's docstring. These pins came from the old pdftotext method, which
        # is exactly what the J12632 geometry measurement showed to be unreliable;
        # they are kept as change-detectors, not as proof of correct behaviour.
        now_phantom_two_line = [
            "Cisco DNA Essentials License for IE9300 Series · IE9300-DNA-E",
            "IE 9300 DNA Essentials, 3 yr Term License · IE9300-DNA-E-3Y",
            "Digital Download Code for Software License · DIGITAL-DL-CODE",
        ]
        two_line = [
            "SNTC-8X5XNBD 24 Port PoE+ Downlinks With 4x10G Uplink, 36 mth · CON-SNT-IE932PXE",
            "Network Plug-n-Play Connect for Zero-touch Device Deployment · NETWORK-PNP-LIC",
        ]
        for text in now_phantom_two_line:
            self.assertEqual(_sp_wrap_lines(text, 60), 2, f"expected 2 lines (accepted phantom) at col_width=60: {text!r}")
        for text in two_line:
            self.assertEqual(_sp_wrap_lines(text, 60), 2, f"expected 2 lines at col_width=60: {text!r}")

    def test_confirmed_tn_sheet_items_at_wider_col_width(self):
        # Same workbook, sheet TN ("Technical Notes and Clarifications" A-E),
        # confirmed against the real generated Technical PDF at col_width=68.43
        # (TN/T&C sheets use a wider column than the BOQ Description column).
        # Item D originally needed the low end of the old [9.15, 9.3] MDW
        # window; under the safety-biased MDW=7.9 (post-J12632) it now predicts
        # one extra phantom line (3 instead of 2) — accepted trade-off, see
        # TestSpWrapLinesJ12632Regression.
        cases = [
            (
                "Coating and painting as per manufacturers' standard unless "
                "specifically mentioned in the proposal.",
                2,
            ),
            (
                "Inclusions:\n"
                "- Central Rack including server, network equipment, internal "
                "wiring, patch cables, and accessories\n"
                "- CCTV Cameras each with Wall Mount Bracket and Junction Box\n"
                "- CCTV Workstations\n"
                "- CCTV VMS and Client Station Software",
                6,
            ),
            (
                "Exclusions (to be provided by client):\n"
                "- External Field Cables\n"
                "- Cable Supports\n"
                "- Mounting Poles (if necessary)",
                4,
            ),
            (
                "All civil works such as running of cables, carpentry, "
                "foundational works or any hot works, equipment installation "
                "and field cable termination are to be provided by the Client.",
                3,  # was 2 under MDW=9.2; now a phantom extra line under MDW=7.9
            ),
            (
                "Work permits required are to be provided by the Client. Work "
                "Permit refers to the permit given to our engineer/technician "
                "that allows him to do work on-site after completing the "
                "required site safety training. This is different from the "
                "work visa, which is already included in the Mob/Demob fee.",
                5,  # was 4 under MDW=9.2/7.9; phantom extra line under MDW=7.50
            ),
        ]
        for text, expected in cases:
            self.assertEqual(
                _sp_wrap_lines(text, 68.43), expected,
                f"expected {expected} lines at col_width=68.43: {text[:50]!r}",
            )


class TestSpWrapLinesItalicRegression(WindowsWrapCalibrationMixin, unittest.TestCase):
    """Regression cases for italic-comment wrapping, pinned against real PDFs.

    "*** ..." clarification comments render in Arial Italic, which wraps to more
    lines than the non-italic Helvetica metric predicts — so the row, sized for the
    smaller count, clips its last line.  _sp_wrap_lines(..., italic=True) applies
    _SP_ITALIC_INFLATE to correct this.  Every case here was confirmed against the
    actual generated Windows PDF (via `pdftotext -layout`) for
    "J12824 EKIUM - VENUS FPSO - NAVCOM B1" — Commercial (col_width=55) and Technical
    (col_width=60).  Across all 62 matchable comment rows the italic prediction matched
    the true PDF line count exactly; these are the boundary cases that pin the factor.

    The true_italic values below are unaffected by the MDW=7.9 recalibration (see
    _SP_MDW_PX / TestSpWrapLinesJ12632Regression) — they still match exactly. Only the
    non_italic baselines shifted up (the base prediction now needs one more line before
    the italic factor is even applied), so those are updated here.
    """

    def test_confirmed_clip_cases_need_the_extra_italic_line(self):
        # (text, col_width, non_italic_pred, true_italic_lines). Each clipped in the
        # real PDF because the non-italic prediction was one line short.
        cases = [
            # Technical VSAT r211 — the "Satellite phone" note (screenshot)
            ("*** Satellite phone is not in the specification. Therefore, propose as an option.", 60, 2, 2),
            # Technical Berthing Aids r49
            ("*** Two (2) Powerbankc can keep a CAT MAX System Running for 45 hours.", 60, 2, 2),
            # Technical VSAT r44 — real PDF also hyphen-breaks "Ka-band"; inflation covers it
            ("*** The specification does not mention any BUC requirement. Therefore, only "
             "C-band BUC is included in the offer. The other two Ku-band and Ka-band will "
             "add only dummy BUC. Please advise the require band BUC and its power.", 60, 4, 4),
            # Commercial RADAR r52 — the "X-Band" note (screenshot)
            ("*** Current Assumption is X-band Antenna and the Processor Unit distance is "
             "within 50mtr. Client to advise if more than 50mtr is required.", 55, 3, 3),
            # Commercial ES r23
            ("*** Commissioning man-days quantity is an estimate based on past experience. "
             "Extra man-days are billable at the man-day rates indicated.", 55, 3, 3),
        ]
        for text, cw, non_italic, true_italic in cases:
            self.assertEqual(
                _sp_wrap_lines(text, cw), non_italic,
                f"non-italic baseline changed: {text[:40]!r}",
            )
            self.assertEqual(
                _sp_wrap_lines(text, cw, italic=True), true_italic,
                f"expected {true_italic} italic lines at col={cw}: {text[:40]!r}",
            )

    def test_stable_italic_cases_are_not_over_inflated(self):
        # Genuinely-fitting italic comments that must NOT gain a phantom blank line
        # when the inflation factor is applied (true count == non-italic count here).
        #
        # The first two cases (col_width=60) DID have true count == non-italic count
        # under MDW=9.2 (both 2). Under the safety-biased MDW=7.9 they now pick up a
        # phantom third line — an accepted regression (see _SP_MDW_PX rationale):
        # eliminating the J12632 clipping required narrowing the base MDW itself, and
        # that narrowing applies before the italic factor is ever considered.
        cases = [
            # Technical VSAT r95 — nearest to the phantom edge (flips only at factor 1.043)
            ("*** FO cores are not specify in the Block Diagram. Assume that each FO cable "
             "has 24 cores. Please advise the FO cable information.", 60, 3),
            # Technical RADAR r47 — the S-band twin of the clipped X-band note; fits at 2
            ("*** Current Assumption is S-band Radar Antenna and the Processor Unit distance "
             "is within 50mtr. Client to advise if more than 50mtr is required.", 60, 3),
            # Commercial VSAT r129
            ("*** DWDM equipment and all other subsea communication connection methods are "
             "not included in JEN's scope of supply and shall be provided by others.", 55, 3),
        ]
        for text, cw, expected in cases:
            self.assertEqual(
                _sp_wrap_lines(text, cw, italic=True), expected,
                f"italic over-inflated to a phantom line at col={cw}: {text[:40]!r}",
            )

    def test_italic_factor_never_reduces_line_count(self):
        # Italic can only ever ADD lines vs the non-italic prediction, never remove.
        for text in [
            "*** Short note",
            "*** A longer clarification comment that wraps across two full lines in the column",
            "*** Current Assumption is X-band Antenna and the Processor Unit distance is "
            "within 50mtr. Client to advise if more than 50mtr is required.",
        ]:
            for cw in (55, 60, 68.43):
                self.assertGreaterEqual(
                    _sp_wrap_lines(text, cw, italic=True),
                    _sp_wrap_lines(text, cw),
                )


class TestSpWrapLinesJ12632Regression(WindowsWrapCalibrationMixin, unittest.TestCase):
    """Regression cases that forced the MDW=9.2 -> 7.9 recalibration.

    Confirmed against the real Windows-generated PDFs for "J12632 SPL - 2GW
    TENNET HVDC BETA OSS - CCTV ITEM CHANGES" (Commercial + Technical). Under
    the old MDW=9.2 these all predicted one fewer line than Excel's real
    renderer produced, and — because overflow text is DROPPED rather than
    visually clipped — the words below were confirmed completely absent from
    the generated PDF via full-text search (not just eyeballed):
    "10G Uplinks", "rugged series", "IE9300 Series", "Rugged SFP",
    "(Safe Area)", and "(REMOVED)". These are real spec/scope words missing
    from a client-facing proposal, which is why the calibration now biases
    toward over-wrapping instead of "centered in the window" — see the
    _SP_MDW_PX comment for the full rationale.
    """

    def test_commercial_boq_items_that_were_silently_clipped(self):
        # Commercial J12632, Simple Proposal "Proposal" sheet, col_width=55.
        cases = [
            # Row 64 — "10G Uplinks" was completely absent from the PDF.
            "Cisco IE-9320-22S2C4X-E 24 Port SFP Downlinks with 4 10G Uplinks",
            # Row 65 — trailing "series" absent.
            "Cisco IE9300_SW Software for Catalyst IE9300 rugged series",
            # Row 68 — trailing "IE9300 Series" absent.
            "Cisco IE9300-DNA-E Cisco DNA Essentials license for IE9300 Series",
            # Row 72 — trailing "Rugged SFP" absent.
            "Cisco GLC-FE-100LX-RGD= 100Mbps Single Mode Rugged SFP",
            # Rows 110/135/159/181 — trailing "(Safe Area)" absent (same text,
            # repeated across 4 different camera sections in the same document).
            "10m, 3 Core (1.5mm FLEX) Cable c/w Cable Glands (Safe Area)",
        ]
        for text in cases:
            self.assertEqual(
                _sp_wrap_lines(text, 55), 2,
                f"expected 2 lines (was clipped to 1 under old MDW): {text!r}",
            )

    def test_technical_cctv_title_row_that_was_silently_clipped(self):
        # Technical J12632, "CCTV" sheet (prepare_to_print_technical flow,
        # col_width=60) — row 181. "(REMOVED)" was completely absent from the
        # PDF, dropping the one word that flags this item as a removed scope
        # item on a document specifically about item changes.
        text = "FIXED INDOOR CCTV CAMERA STATIONS - EMC TYPE (REMOVED)"
        self.assertEqual(
            _sp_wrap_lines(text, 60), 2,
            f"expected 2 lines (was clipped to 1 under old MDW): {text!r}",
        )


class TestSpWrapLinesGroundTruthCalibration(WindowsWrapCalibrationMixin, unittest.TestCase):
    """Pins the MDW=7.50 calibration derived from real rendered PDF geometry.

    Unlike every earlier pinned case in this file — which read "true" line counts off
    `pdftotext -layout` and were repeatedly fooled by hyphen-wrapped words rendering as
    "electro-" / "polished" — these come from actual glyph coordinates in the Windows
    PDFs for "J12632 SPL - 2GW TENNET HVDC BETA OSS" (tools/extract_wrap_ground_truth.py),
    matched back to their source cells.  216 rows across both documents; MDW=7.50 gives
    zero mismatches at BOTH col_width=55 and col_width=68.

    Regenerate with:
        python tools/extract_wrap_ground_truth.py <proposal.xlsx> <proposal.pdf> <col_width>
    """

    def test_measured_available_width_matches_windows_calibration(self):
        # Measured directly from the PDF: usable text width is ~310pt at col_width=55
        # (the old MDW=7.9 assumed 326.6pt — a ~5% over-estimate that caused the
        # clipping). Both column widths must land inside their measured windows.
        # Refined against a second round of real files from TWO Windows machines
        # (J12632 and J12838/"Baker"), which agree with each other. Measured with
        # bold-aware wrapping, without which col=55 has no perfect window at all:
        #   col=55 perfect 309-311pt, col=68 perfect 380-386pt
        avail_55 = (55 * _SP_MDW_PX_WIN + 1) * 0.75
        avail_68 = (68 * _SP_MDW_PX_WIN + 1) * 0.75
        self.assertTrue(309 <= avail_55 <= 311, f"col=55 avail {avail_55:.1f}pt outside measured 309-311pt")
        self.assertTrue(380 <= avail_68 <= 386, f"col=68 avail {avail_68:.1f}pt outside measured 380-386pt")
        # "(REMOVED)" is 384.71pt wide and MUST wrap — clipped rows are excluded from
        # the ground-truth fit, so this one needs asserting separately.
        self.assertLess(avail_68, 384.71, "col=68 avail too wide — '(REMOVED)' would clip again")

    def test_measured_available_width_matches_mac_calibration(self):
        # Mac renders ~5% more text per line than Windows at the same nominal column
        # width. Measured on two Mac-generated PDFs of the same workbook:
        #   col=55 (Simple Commercial, 225 rows): no clipping at avail <= 326pt,
        #          325-326pt minimises phantom lines
        #   col=68 (Simple Technical,  224 rows): perfect at avail 399-407pt
        # Setting the Windows value on Mac put a blank line under most wrapped rows.
        avail_55 = (55 * _SP_MDW_PX_MAC + 1) * 0.75
        avail_68 = (68 * _SP_MDW_PX_MAC + 1) * 0.75
        self.assertTrue(324 <= avail_55 <= 326,
                        f"col=55 Mac avail {avail_55:.1f}pt outside measured 324-326pt")
        self.assertTrue(399 <= avail_68 <= 407,
                        f"col=68 Mac avail {avail_68:.1f}pt outside measured 399-407pt")
        self.assertGreater(_SP_MDW_PX_MAC, _SP_MDW_PX_WIN,
                           "Mac fits more text per line than Windows — do not collapse these")

    def test_print_prep_calibration_is_separate_from_simple_proposal(self):
        # The normal Commercial/Technical flow renders a different document. The CELL
        # font is Arial 12 in both flows; what differs is the workbook's Normal style
        # (ArialMT 12 here vs Calibri 11 in the simple template). Excel sizes a column
        # in units of the Normal font's "0", so the same column_width buys more room
        # here. Constraint intersection across every measured document:
        #     J12831 col=55 fit       8.879-9.145  (+ clipped row below: < 9.066)
        #     J12632 col=55 / col=60  8.952-9.291 / 8.850-9.361
        #     J12838 col=55 / col=60  8.952-9.388 / 8.850-9.894
        #   => usable window 8.952-9.066
        avail_60 = (60 * _SP_MDW_PX_PRINT + 1) * 0.75
        avail_55 = (55 * _SP_MDW_PX_PRINT + 1) * 0.75
        self.assertTrue(399 <= avail_60 <= 422,
                        f"print-prep col=60 avail {avail_60:.1f}pt outside measured 399-422pt")
        self.assertTrue(370 <= avail_55 <= 378,
                        f"print-prep col=55 avail {avail_55:.1f}pt outside measured 370-378pt")
        # Sharing one constant between the two flows is what caused the long
        # 8.0/8.8/9.2/7.9/7.5 oscillation. Keep them apart.
        self.assertNotEqual(_SP_MDW_PX_PRINT, _SP_MDW_PX_WIN)
        self.assertNotEqual(_SP_MDW_PX_PRINT, _SP_MDW_PX_MAC)

    def test_j12831_nominal_voltage_row_wraps(self):
        # J12831 (BALWIN 5 HVADC OSS - ACS), Windows: this row is 374.70pt wide and must
        # wrap, but at MDW=9.2 avail was 380.25pt so it was predicted to fit on one line
        # and "Optional)" was silently dropped from the PDF.
        #
        # It is pinned explicitly because the ground-truth fit CANNOT see it: a clipped
        # row's rendered text no longer matches its source cell, so it is excluded from
        # the matched set. The fit looked perfect at 9.2 on this very document while this
        # row was broken. Every known clip case needs an assertion of its own.
        text = "   • Nominal Voltage 230 VAC ±10%, 50 Hz (115 VAC, 60 Hz Optional)"
        self.assertEqual(_sp_wrap_lines(text, 55, mdw=_SP_MDW_PX_PRINT), 2)

    def test_non_ascii_glyphs_are_measured_not_ignored(self):
        # The J12831 row contains '±' and '•'. Measuring against the real PDF showed our
        # Helvetica widths reproduce the actual Arial rendering to within 0.4%, so these
        # glyphs are handled correctly — the clipping was purely an avail_pt problem.
        # Guard against a future font/encoding change silently measuring them as zero.
        from reportlab.pdfbase.pdfmetrics import stringWidth
        for ch in ("±", "•", "×", "%"):
            self.assertGreater(stringWidth(ch, "Helvetica", 12), 0,
                               f"glyph {ch!r} measured as zero width")

    def test_bold_headings_are_measured_with_bold_metrics(self):
        # Title/System/Subsystem rows render bold, and Arial Bold is genuinely wider.
        # This heading is the real case that clipped in a Windows PDF: 3 lines in the
        # PDF, but only 2 were predicted while everything was measured as regular.
        text = "PTZ OUTDOOR CCTV CAMERA STATIONS (CHANGED FROM TRIMODE TO OUTDOOR, ADDITIONAL)"
        self.assertEqual(_sp_wrap_lines(text, 55), 2)
        self.assertEqual(_sp_wrap_lines(text, 55, bold=True), 3)

    def test_bold_never_predicts_fewer_lines_than_regular(self):
        # Helvetica-Bold is wider than Helvetica for every glyph, so enabling bold can
        # only ever add lines. A row sized from bold metrics is therefore never shorter
        # than the regular prediction — it cannot introduce clipping.
        for text in [
            "PTZ OUTDOOR CCTV CAMERA STATIONS (CHANGED FROM TRIMODE TO OUTDOOR, ADDITIONAL)",
            "FIXED INDOOR CCTV CAMERA STATIONS - EMC TYPE (REMOVED)",
            "   • Constructed in 316L stainless steel with electro-polished sunshield",
            "Short heading",
        ]:
            for cw in (55, 60, 68):
                self.assertGreaterEqual(
                    _sp_wrap_lines(text, cw, bold=True),
                    _sp_wrap_lines(text, cw),
                    f"bold predicted fewer lines at col={cw}: {text[:40]!r}",
                )

    def test_only_bold_row_types_are_treated_as_bold(self):
        # Subtitle and Comment are italic, not bold — they are handled by
        # _SP_ITALIC_INFLATE and must not be swept into the bold set.
        self.assertEqual(functions._SP_BOLD_FMTS, ("System", "Subsystem", "Title"))

    def test_print_prep_flow_uses_the_print_calibration(self):
        # _set_wrap_row_heights must pass mdw=_SP_MDW_PX_PRINT. Guards against a future
        # edit dropping the argument and silently falling back to the simple-proposal
        # value, which is ~55pt too narrow here and reintroduces the phantom lines.
        text = "   • Camera Station Has an Ingress Protection Rating of IP66 & IP67"
        self.assertNotEqual(
            _sp_wrap_lines(text, 60, mdw=_SP_MDW_PX_PRINT),
            _sp_wrap_lines(text, 60),
            "pick a case where the two calibrations actually differ",
        )
        self.assertEqual(_sp_wrap_lines(text, 60, mdw=_SP_MDW_PX_PRINT), 1)

    def test_hyphen_broken_compound_word_row_gets_enough_lines(self):
        # Windows breaks "electro-polished" after the hyphen; Mac keeps it whole. We do
        # not model hyphen-splitting (it would wrongly split part numbers like
        # WS-C2960X-24TS-L), but the correct LINE COUNT — all a row height needs — still
        # falls out of the corrected width. Real PDF: 3 lines.
        text = ("   • Constructed in 316L stainless steel with electro-polished "
                "sunshield, Equipped with a Pre-Terminated 3 m cable tail")
        self.assertEqual(_sp_wrap_lines(text, 55), 3)

    def test_leading_indent_is_measured_not_discarded(self):
        # Sub-item rows are indented in the source and Excel renders that indent, so it
        # consumes real width. The old segment.split() dropped it, under-counting
        # exactly the bullet rows most likely to wrap. A 3-space indent is ~10pt at
        # Arial 12, so a string sized to sit just under the boundary must gain a line
        # once the indent is counted.
        body = ("Constructed in 316L stainless steel with electro-polished sunshield, "
                "Equipped with a Pre-Terminated 3 m cable tail")
        self.assertEqual(_sp_wrap_lines(body, 55), 2)
        self.assertEqual(_sp_wrap_lines("   • " + body, 55), 3)


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

    def test_double_hash_and_grandchild_prefixes_become_grandchild_bullet(self):
        """Test that ## and an existing ◦/▹ prefix all become the ◦ grandchild bullet.

        ▹ (U+25B9 WHITE RIGHT-POINTING SMALL TRIANGLE) was hote's original third-level
        marker; it rendered visibly larger than ‣ at the same font size, so hote switched
        to ◦ (U+25E6 WHITE BULLET). Both are still recognized here for any content typed
        or pasted before that change — see functions.py's format_text.
        """
        systems = pd.DataFrame({
            "Description": ["## Deep item", "◦ Already grandchild", "▹ Old-style grandchild", "Regular item"],
            "Format": ["Description", "Description", "Description", "Description"],
        })

        mask = systems["Format"] == "Description"
        desc_col = systems.loc[mask, "Description"].str.strip().str.lstrip("• ")
        starts_double_hash = desc_col.str.startswith("##")
        starts_grandchild = desc_col.str.startswith("◦") | desc_col.str.startswith("▹")

        result = pd.Series(index=desc_col.index, dtype=str)
        result[starts_double_hash] = "         ◦ " + desc_col[starts_double_hash].str.lstrip("# ")
        result[starts_grandchild] = "         ◦ " + desc_col[starts_grandchild].str.lstrip("◦▹ ")
        other = ~starts_double_hash & ~starts_grandchild
        result[other] = "   • " + desc_col[other]
        systems.loc[mask, "Description"] = result

        self.assertEqual(systems.loc[0, "Description"], "         ◦ Deep item")
        self.assertEqual(systems.loc[1, "Description"], "         ◦ Already grandchild")
        self.assertEqual(systems.loc[2, "Description"], "         ◦ Old-style grandchild")
        self.assertTrue(systems.loc[3, "Description"].startswith("   • "))


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
