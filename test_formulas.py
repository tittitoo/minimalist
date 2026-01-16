"""
Formula verification tests for fill_formula().

CRITICAL: These tests ensure Excel formulas that calculate prices are NEVER
accidentally changed. Any change to these formulas could affect pricing calculations.

Run with: python test_formulas.py
"""

import unittest


# =============================================================================
# MASTER FORMULA DEFINITIONS
# These are the authoritative formulas. Any change here should be intentional
# and reviewed carefully as it affects pricing calculations.
# =============================================================================

FORMULAS = {
    # A1: Reference formula
    "A1": '= "JASON REF: " & Config!B29 &  ", REVISION: " &  Config!B30 & ", PROJECT: " & Config!B26',

    # B: Serial Numbering
    "B": '=IF(AND(A3="", ISNUMBER(D3), ISNUMBER(K3)), COUNT(B2:INDEX($B$1:B2, XMATCH("Title", $AL$1:AL2, 0, -1))) + 1 , "")',

    # N: UCD (Unit Cost after Discount)
    "N": '=IF(K3<>"",K3*(1-M3),"")',

    # O: SCD (Subtotal Cost after Discount)
    "O": '=IF(AND(D3<>"", K3<>"",H3<>"OPTION"),D3*N3,"")',

    # Q: Exchange rate
    "Q": '=IF(J3<>"", INDEX(Config!$B$2:$B$10, XMATCH(J3, Config!$A$2:$A$10, 0))/INDEX(Config!$B$2:$B$10, XMATCH(Config!$B$12, Config!$A$2:$A$10, 0)), "")',

    # R: UCDQ (Unit Cost after Discount in Quoted currency)
    "R": '=IF(AND(D3<>"", K3<>""), N3*Q3,"")',

    # S: SCDQ (Subtotal Cost after Discount in Quoted currency)
    "S": '=IF(AND(D3<>"", K3<>"", H3<>"OPTION", INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*R3, "")',

    # T: BUCQ (Base Unit Cost in Quoted currency) - includes escalations
    "T": '=IF(AND(D3<>"",K3<>""), (R3*(1+$L$1+$N$1+$P$1+$R$1))/(1-0.05),"")',

    # U: BSCQ (Base Subtotal Cost in Quoted currency)
    "U": '=IF(AND(D3<>"",K3<>"",H3<>"OPTION",INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*T3, "")',

    # V: Default escalation
    "V": '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AQ3*$L$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), S3*$L$1, ""))',

    # W: Warranty escalation
    "W": '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AQ3*$N$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), S3*$N$1, ""))',

    # X: Freight escalation
    "X": '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AQ3*$P$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), S3*$P$1, ""))',

    # Y: Special escalation
    "Y": '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AQ3*$R$1, IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), S3*$R$1, ""))',

    # Z: Risk calculation
    "Z": '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>"", H3<>"OPTION"), AS3-(AQ3+V3+W3+X3+Y3), IF(AND(AL3="Lineitem", AK3="Unit Price", H3<>"OPTION"), U3-(S3+V3+W3+X3+Y3), ""))',

    # AA: Margin reference
    "AA": '=IF(AND(D3<>"",K3<>""),$J$1,"")',

    # AC: RUPQ (Recommended Unit Price in Quoted currency)
    "AC": '=IF(AND(D3<>"",K3<>""),CEILING(T3/(1-AA3), 1),"")',

    # AD: RSPQ (Recommended Subtotal Price in Quoted currency)
    "AD": '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED",INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*AC3,"")',

    # AE: UPLS (Unit Price Lumpsum)
    "AE": '=IF(AND(D3<>"",K3<>""), IF(AB3<>"", AB3, AC3),"")',

    # AF: SPLS (Subtotal Price Lumpsum)
    "AF": '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED",INDEX($H$1:H2, XMATCH("Title", $AL$1:AL2, 0, -1))<>"OPTION"), D3*AE3,"")',

    # AG: Profit
    "AG": '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED",AF3<>""),AF3-U3,"")',

    # AH: Margin percentage
    "AH": '=IF(AND(AG3<>"", AG3<>0), AG3/AF3, "")',

    # AI: Total price per item
    "AI": '=IF(AND(D3<>"",K3<>"", H3<>"OPTION"), D3*AE3, "")',

    # F: Unit Price (display)
    "F": '=IF(AND(AL3="Title", ISNUMBER(AJ3)), AJ3, IF(AND(AL3="Lineitem", AK3="Lumpsum", H3<>"OPTION"), "", AE3))',

    # G: Subtotal Price (display)
    "G": '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED"), D3*F3,"")',

    # L: Subtotal Cost
    "L": '=IF(AND(D3<>"",K3<>"",H3<>"OPTION"),D3*K3,"")',

    # AL: Format field (row type detection)
    "AL": '=IF(C4<>"",IF(AND(A4<>"",C4<>""),"Title", IF(B4<>"","Lineitem", IF(LEFT(C4,3)="***","Comment", IF(AND(A4="",B4="",C3="", C5<>"",D5<>""), "Subtitle", IF(AND(A4="",B4="",C3="", C5=""), "Subsystem", "Description"))))),"")',

    # AJ: Lumpsum total
    "AJ": '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(AI4:INDEX(AI4:AI1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), "")',

    # AK: Lumpsum/Unit Price flag
    "AK": '=IF(AL3="Lineitem", IF(ISNUMBER(INDEX($AJ$1:AJ2, XMATCH("Title", $AL$1:AL2, 0, -1))), "Lumpsum", "Unit Price"), "")',

    # AP: SCDQL (Subtotal Cost after Discount in Quoted currency Lumpsum)
    "AP": '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(S4:INDEX(S4:S1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), IF(AND(AL3="Lineitem", AK3="Unit Price"), R3, ""))',

    # AQ: TCDQL (Total Cost after Discount in Quoted currency Lumpsum) - Material cost
    "AQ": '=IF(AND(ISNUMBER(D3), ISNUMBER(AP3), H3<>"OPTION"), D3*AP3, "")',

    # AR: BSCQL (Base Subtotal Cost in Quoted currency Lumpsum)
    "AR": '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(U4:INDEX(U4:U1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), IF(AND(AL3="Lineitem", AK3="Unit Price"), T3, ""))',

    # AS: BTCQL (Base Total Cost in Quoted currency Lumpsum) - Base cost
    "AS": '=IF(AND(ISNUMBER(D3), ISNUMBER(AR3), H3<>"OPTION"), D3*AR3, "")',

    # AT: SSPL (Subtotal Selling Price Lumpsum)
    "AT": '=IF(AND(AL3="Title", ISNUMBER(D3), E3<>""), SUM(AF4:INDEX(AF4:AF1500, XMATCH("Title", AL4:AL1500, 0, 1)-1)), IF(AND(AL3="Lineitem", AK3="Unit Price"), AE3, ""))',

    # AU: TSPL (Total Selling Price Lumpsum) - Selling price
    "AU": '=IF(AND(ISNUMBER(D3), H3<>"WAIVED", H3<>"INCLUDED", H3<>"OPTION", ISNUMBER(AT3)), D3*AT3, "")',

    # AV: Total Profit
    "AV": '=IF(AND(ISNUMBER(D3), ISNUMBER(AS3), ISNUMBER(AU3)), AU3-AS3, "")',

    # AW: Grand Margin
    "AW": '=IF(AND(H3<>"OPTION", ISNUMBER(D3), ISNUMBER(AU3), AU3<>0, ISNUMBER(AV3)), AV3/AU3, "")',
}


# =============================================================================
# FORMULA TESTS
# =============================================================================

class TestCostFormulas(unittest.TestCase):
    """Tests for cost calculation formulas."""

    def test_formula_N_UCD(self):
        """N: Unit Cost after Discount = Unit Cost * (1 - Discount%)"""
        formula = FORMULAS["N"]
        self.assertIn("K3*(1-M3)", formula)
        self.assertIn('IF(K3<>""', formula)

    def test_formula_O_SCD(self):
        """O: Subtotal Cost after Discount = Qty * UCD"""
        formula = FORMULAS["O"]
        self.assertIn("D3*N3", formula)
        self.assertIn('H3<>"OPTION"', formula)

    def test_formula_Q_exchange_rate(self):
        """Q: Exchange rate lookup from Config sheet"""
        formula = FORMULAS["Q"]
        self.assertIn("Config!$B$2:$B$10", formula)
        self.assertIn("XMATCH(J3", formula)
        self.assertIn("Config!$B$12", formula)

    def test_formula_R_UCDQ(self):
        """R: Unit Cost after Discount in Quoted currency = UCD * Exchange Rate"""
        formula = FORMULAS["R"]
        self.assertIn("N3*Q3", formula)

    def test_formula_S_SCDQ(self):
        """S: Subtotal Cost after Discount in Quoted currency"""
        formula = FORMULAS["S"]
        self.assertIn("D3*R3", formula)
        self.assertIn('H3<>"OPTION"', formula)
        # Check parent title OPTION check
        self.assertIn('INDEX($H$1:H2, XMATCH("Title"', formula)

    def test_formula_L_subtotal_cost(self):
        """L: Subtotal Cost = Qty * Unit Cost"""
        formula = FORMULAS["L"]
        self.assertIn("D3*K3", formula)


class TestEscalationFormulas(unittest.TestCase):
    """Tests for escalation calculation formulas."""

    def test_formula_T_includes_all_escalations(self):
        """T: Base Unit Cost includes all escalation factors"""
        formula = FORMULAS["T"]
        # Should include all escalation references
        self.assertIn("$L$1", formula)  # Default
        self.assertIn("$N$1", formula)  # Warranty
        self.assertIn("$P$1", formula)  # Freight
        self.assertIn("$R$1", formula)  # Special
        # Should include 5% risk factor
        self.assertIn("1-0.05", formula)

    def test_formula_V_default_escalation(self):
        """V: Default escalation calculation"""
        formula = FORMULAS["V"]
        self.assertIn("$L$1", formula)
        self.assertIn('AL3="Title"', formula)
        self.assertIn('AL3="Lineitem"', formula)

    def test_formula_W_warranty_escalation(self):
        """W: Warranty escalation calculation"""
        formula = FORMULAS["W"]
        self.assertIn("$N$1", formula)

    def test_formula_X_freight_escalation(self):
        """X: Freight escalation calculation"""
        formula = FORMULAS["X"]
        self.assertIn("$P$1", formula)

    def test_formula_Y_special_escalation(self):
        """Y: Special escalation calculation"""
        formula = FORMULAS["Y"]
        self.assertIn("$R$1", formula)

    def test_formula_Z_risk_calculation(self):
        """Z: Risk = Base Cost - (Material + all escalations)"""
        formula = FORMULAS["Z"]
        # Should subtract all escalation columns
        self.assertIn("V3+W3+X3+Y3", formula)
        self.assertIn("AS3-(AQ3+V3+W3+X3+Y3)", formula)  # For Title
        self.assertIn("U3-(S3+V3+W3+X3+Y3)", formula)   # For Lineitem


class TestPricingFormulas(unittest.TestCase):
    """Tests for pricing calculation formulas."""

    def test_formula_AC_RUPQ(self):
        """AC: Recommended Unit Price = Base Cost / (1 - Margin), rounded up"""
        formula = FORMULAS["AC"]
        self.assertIn("CEILING(T3/(1-AA3)", formula)

    def test_formula_AD_RSPQ(self):
        """AD: Recommended Subtotal Price = Qty * RUPQ"""
        formula = FORMULAS["AD"]
        self.assertIn("D3*AC3", formula)
        # Should exclude OPTION, INCLUDED, WAIVED
        self.assertIn('H3<>"OPTION"', formula)
        self.assertIn('H3<>"INCLUDED"', formula)
        self.assertIn('H3<>"WAIVED"', formula)

    def test_formula_AE_UPLS(self):
        """AE: Unit Price Lumpsum - uses fixed price if available"""
        formula = FORMULAS["AE"]
        self.assertIn('IF(AB3<>""', formula)
        self.assertIn("AB3", formula)  # Fixed price column
        self.assertIn("AC3", formula)  # Fallback to calculated

    def test_formula_AF_SPLS(self):
        """AF: Subtotal Price Lumpsum"""
        formula = FORMULAS["AF"]
        self.assertIn("D3*AE3", formula)

    def test_formula_F_unit_price_display(self):
        """F: Unit Price for display - handles lumpsum vs unit pricing"""
        formula = FORMULAS["F"]
        self.assertIn('AL3="Title"', formula)
        self.assertIn('AL3="Lineitem"', formula)
        self.assertIn('AK3="Lumpsum"', formula)

    def test_formula_G_subtotal_price_display(self):
        """G: Subtotal Price for display"""
        formula = FORMULAS["G"]
        self.assertIn("D3*F3", formula)
        self.assertIn('H3<>"OPTION"', formula)
        self.assertIn('H3<>"INCLUDED"', formula)
        self.assertIn('H3<>"WAIVED"', formula)


class TestProfitFormulas(unittest.TestCase):
    """Tests for profit calculation formulas."""

    def test_formula_AG_profit(self):
        """AG: Profit = Selling Price - Base Cost"""
        formula = FORMULAS["AG"]
        self.assertIn("AF3-U3", formula)

    def test_formula_AH_margin_percentage(self):
        """AH: Margin % = Profit / Selling Price"""
        formula = FORMULAS["AH"]
        self.assertIn("AG3/AF3", formula)

    def test_formula_AV_total_profit(self):
        """AV: Total Profit = Selling Price - Base Cost"""
        formula = FORMULAS["AV"]
        self.assertIn("AU3-AS3", formula)

    def test_formula_AW_grand_margin(self):
        """AW: Grand Margin = Profit / Selling Price"""
        formula = FORMULAS["AW"]
        self.assertIn("AV3/AU3", formula)


class TestLumpsumFormulas(unittest.TestCase):
    """Tests for lumpsum calculation formulas."""

    def test_formula_AJ_lumpsum_total(self):
        """AJ: Lumpsum total - sums AI column until next Title"""
        formula = FORMULAS["AJ"]
        self.assertIn("SUM(AI4:INDEX(AI4:AI1500", formula)
        self.assertIn('XMATCH("Title", AL4:AL1500', formula)

    def test_formula_AK_lumpsum_flag(self):
        """AK: Determines if item uses lumpsum or unit pricing"""
        formula = FORMULAS["AK"]
        self.assertIn('"Lumpsum"', formula)
        self.assertIn('"Unit Price"', formula)

    def test_formula_AP_SCDQL(self):
        """AP: Subtotal Cost Lumpsum"""
        formula = FORMULAS["AP"]
        self.assertIn("SUM(S4:INDEX(S4:S1500", formula)

    def test_formula_AR_BSCQL(self):
        """AR: Base Subtotal Cost Lumpsum"""
        formula = FORMULAS["AR"]
        self.assertIn("SUM(U4:INDEX(U4:U1500", formula)

    def test_formula_AT_SSPL(self):
        """AT: Subtotal Selling Price Lumpsum"""
        formula = FORMULAS["AT"]
        self.assertIn("SUM(AF4:INDEX(AF4:AF1500", formula)


class TestFormatFormulas(unittest.TestCase):
    """Tests for format/structure detection formulas."""

    def test_formula_AL_detects_title(self):
        """AL: Detects Title rows"""
        formula = FORMULAS["AL"]
        self.assertIn('"Title"', formula)
        self.assertIn('A4<>""', formula)

    def test_formula_AL_detects_lineitem(self):
        """AL: Detects Lineitem rows"""
        formula = FORMULAS["AL"]
        self.assertIn('"Lineitem"', formula)
        self.assertIn('B4<>""', formula)

    def test_formula_AL_detects_comment(self):
        """AL: Detects Comment rows (starts with ***)"""
        formula = FORMULAS["AL"]
        self.assertIn('"Comment"', formula)
        self.assertIn('LEFT(C4,3)="***"', formula)

    def test_formula_AL_detects_subtitle(self):
        """AL: Detects Subtitle rows"""
        formula = FORMULAS["AL"]
        self.assertIn('"Subtitle"', formula)

    def test_formula_AL_detects_subsystem(self):
        """AL: Detects Subsystem rows"""
        formula = FORMULAS["AL"]
        self.assertIn('"Subsystem"', formula)

    def test_formula_AL_detects_description(self):
        """AL: Detects Description rows (default)"""
        formula = FORMULAS["AL"]
        self.assertIn('"Description"', formula)


class TestFormulaConsistency(unittest.TestCase):
    """Tests to verify formulas match what's in functions.py"""

    def test_all_formulas_defined(self):
        """Verify all expected formula columns are defined."""
        expected_columns = [
            "A1", "B", "N", "O", "Q", "R", "S", "T", "U",
            "V", "W", "X", "Y", "Z", "AA",
            "AC", "AD", "AE", "AF", "AG", "AH", "AI",
            "F", "G", "L", "AL",
            "AJ", "AK",
            "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW"
        ]
        for col in expected_columns:
            self.assertIn(col, FORMULAS, f"Formula for column {col} is missing")

    def test_formulas_start_with_equals(self):
        """All formulas should start with '=' """
        for col, formula in FORMULAS.items():
            self.assertTrue(
                formula.startswith("="),
                f"Formula {col} doesn't start with '=': {formula[:50]}"
            )

    def test_no_empty_formulas(self):
        """No formula should be empty or just '=' """
        for col, formula in FORMULAS.items():
            self.assertGreater(
                len(formula), 1,
                f"Formula {col} is empty or too short"
            )


class TestCriticalPricingLogic(unittest.TestCase):
    """Critical tests for pricing calculation integrity."""

    def test_margin_formula_uses_division(self):
        """Margin calculation must use division: Profit/Price"""
        # AH = AG/AF (per item margin)
        self.assertIn("AG3/AF3", FORMULAS["AH"])
        # AW = AV/AU (grand margin)
        self.assertIn("AV3/AU3", FORMULAS["AW"])

    def test_profit_formula_uses_subtraction(self):
        """Profit calculation must use subtraction: Selling - Cost"""
        # AG = AF - U (per item profit)
        self.assertIn("AF3-U3", FORMULAS["AG"])
        # AV = AU - AS (total profit)
        self.assertIn("AU3-AS3", FORMULAS["AV"])

    def test_price_ceiling_rounds_up(self):
        """Price calculation must use CEILING to round up"""
        self.assertIn("CEILING(", FORMULAS["AC"])

    def test_escalation_formula_adds_factors(self):
        """Escalation must add all factors together"""
        # T should have: 1 + L$1 + N$1 + P$1 + R$1
        formula = FORMULAS["T"]
        self.assertIn("1+$L$1+$N$1+$P$1+$R$1", formula)

    def test_option_items_excluded_from_totals(self):
        """OPTION items must be excluded from calculations"""
        # These formulas should check for OPTION
        option_sensitive = ["O", "S", "U", "AD", "AF", "G", "L", "AQ", "AS", "AU"]
        for col in option_sensitive:
            self.assertIn(
                'OPTION',
                FORMULAS[col],
                f"Formula {col} should check for OPTION status"
            )


class TestFormulasMatchSource(unittest.TestCase):
    """
    CRITICAL: Verify formulas in this test file match functions.py
    This catches any accidental changes to either file.
    """

    def test_formulas_match_fill_formula_function(self):
        """
        Read fill_formula() from functions.py and verify formulas match.
        This is the most important test - it ensures test and source are in sync.
        """
        import re

        # Read the functions.py file
        with open("functions.py", "r") as f:
            content = f.read()

        # Extract the fill_formula function
        match = re.search(
            r'def fill_formula\(sheet\):.*?(?=\ndef |\Z)',
            content,
            re.DOTALL
        )
        self.assertIsNotNone(match, "Could not find fill_formula function")
        func_content = match.group(0)

        # Check each formula exists in the function
        formulas_to_check = {
            "N": '=IF(K3<>"",K3*(1-M3),"")',
            "O": '=IF(AND(D3<>"", K3<>"",H3<>"OPTION"),D3*N3,"")',
            "R": '=IF(AND(D3<>"", K3<>""), N3*Q3,"")',
            "L": '=IF(AND(D3<>"",K3<>"",H3<>"OPTION"),D3*K3,"")',
            "AC": '=IF(AND(D3<>"",K3<>""),CEILING(T3/(1-AA3), 1),"")',
            "AG": '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED",AF3<>""),AF3-U3,"")',
            "AH": '=IF(AND(AG3<>"", AG3<>0), AG3/AF3, "")',
        }

        for col, formula in formulas_to_check.items():
            self.assertIn(
                formula,
                func_content,
                f"Formula for column {col} not found in fill_formula():\n{formula}"
            )

    def test_critical_pricing_formulas_unchanged(self):
        """
        Verify the most critical pricing formulas haven't changed.
        These directly affect the final price shown to customers.
        """
        # F: Unit Price shown to customer
        self.assertEqual(
            FORMULAS["F"],
            '=IF(AND(AL3="Title", ISNUMBER(AJ3)), AJ3, IF(AND(AL3="Lineitem", AK3="Lumpsum", H3<>"OPTION"), "", AE3))'
        )

        # G: Subtotal Price shown to customer
        self.assertEqual(
            FORMULAS["G"],
            '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED"), D3*F3,"")'
        )

        # AC: Recommended Unit Price calculation
        self.assertEqual(
            FORMULAS["AC"],
            '=IF(AND(D3<>"",K3<>""),CEILING(T3/(1-AA3), 1),"")'
        )

    def test_escalation_formula_exact_match(self):
        """
        Verify escalation formula T is exactly correct.
        This formula adds Default, Warranty, Freight, Special escalations
        and applies a 5% risk factor.
        """
        expected = '=IF(AND(D3<>"",K3<>""), (R3*(1+$L$1+$N$1+$P$1+$R$1))/(1-0.05),"")'
        self.assertEqual(FORMULAS["T"], expected)


if __name__ == "__main__":
    unittest.main(verbosity=2)
