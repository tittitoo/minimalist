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
    set_paren_spacing,
    set_x,
    set_case_preserve_acronym,
    title_case_ignore_double_char,
    normalize_standard_tokens,
    set_dimension_unit_chain,
    format_description_text,
    _sp_wrap_lines,
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
        self.assertEqual(normalize_standard_tokens("cat6a patch cord"), "Cat6a patch cord")
        self.assertEqual(normalize_standard_tokens("CAT6A patch cord"), "Cat6a patch cord")

    def test_normalizes_ip_ratings(self):
        self.assertEqual(normalize_standard_tokens("ip65 rated"), "IP65 rated")
        self.assertEqual(normalize_standard_tokens("ipx6 rated"), "IPX6 rated")
        self.assertEqual(normalize_standard_tokens("ip69k washdown"), "IP69K washdown")

    def test_normalizes_spelled_out_meter_collapsing_space(self):
        self.assertEqual(normalize_standard_tokens("40 meter"), "40m")
        self.assertEqual(normalize_standard_tokens("0.2 meter cable"), "0.2m cable")

    def test_normalizes_spelled_out_ohm_to_symbol(self):
        self.assertEqual(normalize_standard_tokens("50 ohm resistor"), "50Ω resistor")

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

    def test_always_normalizes_units_regardless_of_title_case_flag(self):
        # Units/standards normalize even when title_case=False (long text, or a
        # Format type that's never title-cased).
        self.assertEqual(
            format_description_text("27MM bracket", title_case=False), "27 mm bracket"
        )

    def test_returns_empty_string_for_falsy_input(self):
        self.assertEqual(format_description_text(None, title_case=True), "")
        self.assertEqual(format_description_text("", title_case=True), "")


class TestSpWrapLinesRealPdfRegression(unittest.TestCase):
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
    """

    def test_confirmed_single_line_cases_do_not_get_a_phantom_second_line(self):
        # Commercial J12815 GEV CCTV R2, sheet CCTV, "Cisco IE-9320-24P4X-E"
        # line items — confirmed single-line in the real Windows Commercial PDF
        # (col_width=55), but mispredicted as 2 lines (phantom blank line)
        # under the old MDW=8.0.
        cases = [
            "Cisco DNA Essentials License for IE9300 Series · IE9300-DNA-E",
            "IE 9300 DNA Essentials, 3 yr Term License · IE9300-DNA-E-3Y",
            "Digital Download Code for Software License · DIGITAL-DL-CODE",
            "Not Related to an IoT Solution; for Tracking Only. · IOT-OTHER",
            "Software for Catalyst IE9300 Rugged Series · IE9300_sw",
        ]
        for text in cases:
            self.assertEqual(_sp_wrap_lines(text, 55), 1, f"expected 1 line: {text!r}")

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
        one_line = [
            "Cisco DNA Essentials License for IE9300 Series · IE9300-DNA-E",
            "IE 9300 DNA Essentials, 3 yr Term License · IE9300-DNA-E-3Y",
            "Digital Download Code for Software License · DIGITAL-DL-CODE",
        ]
        two_line = [
            "SNTC-8X5XNBD 24 Port PoE+ Downlinks With 4x10G Uplink, 36 mth · CON-SNT-IE932PXE",
            "Network Plug-n-Play Connect for Zero-touch Device Deployment · NETWORK-PNP-LIC",
        ]
        for text in one_line:
            self.assertEqual(_sp_wrap_lines(text, 60), 1, f"expected 1 line at col_width=60: {text!r}")
        for text in two_line:
            self.assertEqual(_sp_wrap_lines(text, 60), 2, f"expected 2 lines at col_width=60: {text!r}")

    def test_confirmed_tn_sheet_items_at_wider_col_width(self):
        # Same workbook, sheet TN ("Technical Notes and Clarifications" A-E),
        # confirmed against the real generated Technical PDF at col_width=68.43
        # (TN/T&C sheets use a wider column than the BOQ Description column).
        # Item D specifically needed the low end of the MDW range still
        # compatible with the col_width=55/60 cases above — this is what
        # pinned MDW down to [9.15, 9.3] instead of a wider range.
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
                2,
            ),
            (
                "Work permits required are to be provided by the Client. Work "
                "Permit refers to the permit given to our engineer/technician "
                "that allows him to do work on-site after completing the "
                "required site safety training. This is different from the "
                "work visa, which is already included in the Mob/Demob fee.",
                4,
            ),
        ]
        for text, expected in cases:
            self.assertEqual(
                _sp_wrap_lines(text, 68.43), expected,
                f"expected {expected} lines at col_width=68.43: {text[:50]!r}",
            )


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
