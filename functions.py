""" Multiple functions to support Excel automation.
    © Thiha Aung (infowizard@gmail.com)
"""

import os
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import requests
import xlwings as xw  # type: ignore

import hide
import checklist_collections as cc

LEGEND = {
    "UC": "Unit cost in original (buying) currency",
    "SC": "Subtotal cost in original (buying) currency",
    "Discount": "Discount in percentage from the supplier",
    "UCD": "Unit cost after discount in original (buying) currency",
    "SCD": "Subtotal cost after discount in original (buying) currency",
}

MACRO_NB = xw.Book("PERSONAL.XLSB")

# Accounting number format
ACCOUNTING = "_(* #,##0.00_);_(* (#,##0.00);_(* " "-" "??_);_(@_)"
EXCNANGE_RATE = '_(* #,##0.0000_);_(* (#,##0.0000);_(* "-"????_);_(@_)'

RESOURCES = os.path.join(
    os.path.dirname(os.path.realpath(__file__)),
    "resources/",
)

# To update the value upon updating of the template.
LATEST_WB_VERSION = "R2"


def set_nitty_gritty(text):
    """Fix annoying text"""
    # Strip EOL
    text = text.strip()
    # Strip 2 or more spaces
    text = re.sub(" {2,}", " ", text)
    # Put bullet point for Sub-subitem preceded by '-' or '~'.
    text = re.sub("^(-|~)", "•", text)
    # Put bullet point for Sub-subitem preceded by a single * followed by space.
    text = re.sub("^[*?]\s", " • ", text)
    # Instead of ';' at the end of line, use ':' instead.
    text = re.sub(";$", ":", text)
    text = set_comma_space(text)
    text = set_x(text)
    return text


def set_comma_space(text):
    """Fix having space before comma and not having space after comma"""
    # fix word+space+, to word+,
    x = re.compile("\w+\s,")
    if x.search(text):
        substring = re.findall("\w+\s,", text)
        for word in substring:
            text = re.sub(word, word[:-2] + ",", text)

    # Fix word+,+no-space to word+,+space
    x = re.compile(",\d?\w+")
    if x.search(text):
        # Ignores format like 1,200 but matches 1,w
        substring = re.findall("(?<![0-9]),\w+", text)
        for word in substring:
            text = re.sub(word, ", " + word[1:], text)
    return text


def title_case_ignore_double_char(text):
    words = text.split()
    titled_words = []
    for word in words:
        if len(word) > 2:  # So that two letter words are ignored.
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
        # Restore acronyms
        for acronym in acronyms:
            text = text.replace(acronym.title(), acronym)
        return text

    elif capitalize:
        text = text.lower()
        for acronym in acronyms:
            print(acronym)
            print(acronym.lower())
            text = text.replace(acronym.lower(), acronym)
        # text = text.capitalize()
        return text
        # return acronyms

    elif upper:
        text = text.upper()
        return text


def set_x(text):
    """Function to replace description such as 1x, 20x, 10X ,
    x1, x20, X20 into 1 x, 20 x, 10 x, x 1, x 20, X 10 etc."""
    # For cases such as 20x, 30X. Allows if followed by -
    x = re.compile("\d+x(?!-)|\d+X(?!-)")
    if x.search(text):
        substring = re.findall("(\d+x|\d+X)", text)
        for word in substring:
            text = re.sub(word, (word[:-1] + " x"), text)
    # For cases such as x20, X30
    x = re.compile("(x\d+|X\d+)")
    if x.search(text):
        substring = re.findall("(x\d+|X\d+)", text)
        for word in substring:
            text = re.sub(word, ("x " + word[1:]), text)
    # For cases such as 20 X, 30 X
    x = re.compile("(\d+ X)")
    if x.search(text):
        substring = re.findall("(\d+ X)", text)
        for word in substring:
            text = re.sub(word, (word[:-1] + "x"), text)
    # For cases such as X 20, X 30
    x = re.compile("(X \d+)")
    if x.search(text):
        substring = re.findall("(X \d+)", text)
        for word in substring:
            text = re.sub(word, ("x" + word[1:]), text)
    return text


def fill_formula(sheet):
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    if sheet.name not in skip_sheets:
        # Formula to cells
        # Increase the last row by 1 so that the cells are not left empty
        last_row = sheet.range("C1048576").end("up").row + 1
        sheet.range(
            "A1"
        ).formula = '= "JASON REF: " & Config!B29 &  ", REVISION: " &  Config!B30 & ", PROJECT: " & Config!B26'
        # Serail Numbering (SN)
        sheet.range(
            "B3:B" + str(last_row)
        ).formula = '=IF(AND(ISNUMBER(D3), ISNUMBER(K3), XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)), COUNT(INDIRECT(CONCAT("B",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1),":B",ROW()-1))) + 1, "")'
        sheet.range("N3:N" + str(last_row)).formula = '=IF(K3<>"",K3*(1-M3),"")'
        sheet.range(
            "O3:O" + str(last_row)
        ).formula = '=IF(AND(D3<>"", K3<>"",H3<>"OPTION"),D3*N3,"")'
        # Exchange rates
        sheet.range(
            "Q3:Q" + str(last_row)
        ).formula = '=IF(Config!$B$12="SGD",IF(J3<>"",VLOOKUP(J3,Config!$A$2:$B$10,2,FALSE),""),IF(J3<>"",VLOOKUP(J3,Config!$A$2:$B$10,2,FALSE)/VLOOKUP(Config!$B$12,Config!$A$2:$B$10,2,FALSE),""))'
        sheet.range(
            "R3:R" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"") ,N3*Q3,"")'
        # sheet.range('S3:S' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION") ,D3*R3,"")'
        sheet.range(
            "S3:S" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION",INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))<>"OPTION"),D3*R3,"")'
        sheet.range(
            "T3:T" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>""), (R3*(1+$L$1+$N$1+$P$1+$R$1))/(1-0.05),"")'
        sheet.range(
            "U3:U" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION",INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))<>"OPTION"),D3*T3,"")'
        # Default
        sheet.range(
            "V3:V" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),D3*R3*$L$1,"")'
        # Warranty
        sheet.range(
            "W3:W" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),D3*R3*$N$1,"")'
        # Freight (Inbound)
        sheet.range(
            "X3:X" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),D3*R3*$P$1,"")'
        # Special (Condition)
        sheet.range(
            "Y3:Y" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),D3*R3*$R$1,"")'
        # Risk
        sheet.range(
            "Z3:Z" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),U3-(S3+V3+W3+X3+Y3),"")'
        sheet.range(
            "AA3:AA" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>""),$J$1,"")'
        sheet.range(
            "AC3:AC" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>""),CEILING(T3/(1-AA3), 1),"")'
        # sheet.range('AD3:AD' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION",H3<>"INCLUDED"),D3*AC3,"")'
        sheet.range(
            "AD3:AD" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION",H3<>"INCLUDED", H3<>"WAIVED",(INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))) <>"OPTION"),D3*AC3,"")'
        sheet.range(
            "AE3:AE" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>""),IF(AB3<>"",AB3,AC3),"")'
        # sheet.range('AF3:AF' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED"),D3*AE3,"")'
        sheet.range(
            "AF3:AF" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED",(INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))) <>"OPTION"),D3*AE3,"")'
        # sheet.range('AF3:AF' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),D3*AE3,"")'
        sheet.range(
            "AG3:AG" + str(last_row)
        ).formula = (
            '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED",AF3<>""),AF3-U3,"")'
        )
        sheet.range(
            "AH3:AH" + str(last_row)
        ).formula = '=IF(AND(AG3<>"",AG3<>0),AG3/AF3,"")'
        sheet.range(
            "AI3:AI" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION"),D3*AE3,"")'
        # sheet.range('AL3:AL' + str(last_row)).formula = '=IF(A3<>"","Title",IF(B3<>"","Lineitem",IF(LEFT(C3,3)="***","Comment",IF(AND(A3="",B3="",C2="", C4<>"",D4<>""), "Subtitle",""))))'
        # Unit Price
        sheet.range(
            "F3:F" + str(last_row)
        ).formula = '=IF(AND(AL3="Title", ISNUMBER(AJ3)), AJ3, IF(AND(AL3="Lineitem", AK3="Lumpsum", H3<>"OPTION"), "", AE3))'
        # sheet.range('F3:F' + str(last_row)).formula = '=IF(AE3<>"", AE3,"")'
        sheet.range(
            "G3:G" + str(last_row)
        ).formula = (
            '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED"), D3*F3,"")'
        )
        sheet.range(
            "L3:L" + str(last_row)
        ).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION"),D3*K3,"")'
        # For Format field
        sheet.range("AL1").value = "Title"
        sheet.range("AL3").value = "System"
        # sheet.range('AL4:AL' + str(last_row)).formula = '=IF(C4<>"",IF(AND(A4<>"",C4<>""),"Title", IF(B4<>"","Lineitem", IF(LEFT(C4,3)="***","Comment", IF(AND(A4="",B4="",C3="", C5<>"",D5<>""), "Subtitle","Description")))),"")'
        # Implement "Subsystem"
        sheet.range(
            "AL4:AL" + str(last_row)
        ).formula = '=IF(C4<>"",IF(AND(A4<>"",C4<>""),"Title", IF(B4<>"","Lineitem", IF(LEFT(C4,3)="***","Comment", IF(AND(A4="",B4="",C3="", C5<>"",D5<>""), "Subtitle", IF(AND(A4="",B4="",C3="", C5=""), "Subsystem", "Description"))))),"")'
        sheet.range("AL" + str(last_row + 1)).value = "Title"

        # For Lumpsum
        # sheet.range('AJ3:AJ' + str(last_row)).formula = '=IF(AND(AL3="Title", D3=1, E3="lot"), SUM(INDIRECT(CONCAT("AF", ROW()+1, ":AF",((MATCH("Title",INDIRECT(CONCAT("AL", ROW()+1, ":AL", MATCH(REPT("z",50),AL:AL))),0)) + ROW())))), "")'
        sheet.range(
            "AJ3:AJ" + str(last_row)
        ).formula = '=IF(AND(AL3="Title", D3=1, E3="lot"), SUM(INDIRECT(CONCAT("AI", ROW()+1, ":AI",((MATCH("Title",INDIRECT(CONCAT("AL", ROW()+1, ":AL", MATCH(REPT("z",50),AL:AL))),0)) + ROW())))), "")'
        sheet.range(
            "AK3:AK" + str(last_row)
        ).formula = '=IF(AL3="Lineitem", IF(ISNUMBER(INDIRECT(CONCAT("AJ",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))), "Lumpsum", "Unit Price"), "")'


def fill_formula_wb(wb):
    for sheet in wb.sheets:
        fill_formula(sheet)


def fill_lastrow(wb):
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]

    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            fill_lastrow_sheet(wb, sheet)


def fill_lastrow_sheet(wb, sheet):
    pwb = xw.books("PERSONAL.XLSB")
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    if sheet.name not in skip_sheets:
        last_row = sheet.range("C1048576").end("up").row
        (pwb.sheets["Design"].range("5:5")).copy(
            sheet.range(str(last_row + 2) + ":" + str(last_row + 2))
        )
        sheet.range("F" + str(last_row + 2)).formula = '="Subtotal(" & Config!B12 & ")"'
        sheet.range("F" + str(last_row + 2)).font.size = 9
        sheet.range("G" + str(last_row + 2)).formula = (
            "=SUM(G3:G" + str(last_row + 1) + ")"
        )
        # SCDQ: Subtotal cost after discount in quoted currency
        sheet.range("S" + str(last_row + 2)).formula = (
            "=SUM(S3:S" + str(last_row + 1) + ")"
        )
        # BSCQ: Base subtotal cost in quoted currency
        sheet.range("U" + str(last_row + 2)).formula = (
            "=SUM(U3:U" + str(last_row + 1) + ")"
        )
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
        sheet.range("AF" + str(last_row + 2)).formula = (
            "=SUM(AF3:AF" + str(last_row + 1) + ")"
        )
        sheet.range("AG" + str(last_row + 2)).formula = (
            "=SUM(AG3:AG" + str(last_row + 1) + ")"
        )
        sheet.range("AH" + str(last_row + 2)).formula = (
            "=AG" + str(last_row + 2) + "/AF" + str(last_row + 2)
        )
        sheet.range("AL" + str(last_row + 2)).value = "Title"

        # Format
        sheet.range(f"S{last_row+2}:S{last_row+2}").font.color = (0, 144, 81)
        sheet.range(f"V{last_row+2}:Z{last_row+2}").font.color = (0, 144, 81)
        sheet.range(f"{last_row+2}:{last_row+2}").font.bold = True

        # Set-up print area
        sheet.page_setup.print_area = "A1:H" + str(last_row + 2)


def unhide_columns(sheet):
    """Unhide all columns while setting the width for selected columns"""
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    if sheet.name not in skip_sheets:
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
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    if sheet.name not in skip_sheets:
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
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    if sheet.name not in skip_sheets:
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
    summary_formula = []
    collect = []  # Collect formula to be put in summary page.
    # formula_fragment = '=IF(OR(Config!B13="COMMERCIAL PROPOSAL", Config!B13="BUDGETARY PROPOSAL"),'
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    # The design will now be taken from PERSONAL.XLSB
    pwb = xw.books("PERSONAL.XLSB")

    # Initialize counters
    start_row = 19
    count = 1
    offset = 20
    sheet = wb.sheets["Summary"]

    # Need to collect information if already exists so that it can be repopulated
    system_count = wb.sheets.count - len(
        skip_sheets
    )  # Previously hard-coded to 5,number of default skip_sheets
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
            if sheet not in skip_sheets:
                sheet = wb.sheets[sheet]
                last_row = sheet.range("G1048576").end("up").row
                collect = [
                    "='" + sheet.name + "'!$C$3",
                    "='" + sheet.name + "'!$G$" + str(last_row),
                    "='" + sheet.name + "'!$S$" + str(last_row),
                    "='" + sheet.name + "'!$V$" + str(last_row),
                    "='" + sheet.name + "'!$W$" + str(last_row),
                    "='" + sheet.name + "'!$X$" + str(last_row),
                    "='" + sheet.name + "'!$Y$" + str(last_row),
                    "='" + sheet.name + "'!$Z$" + str(last_row),
                    "='" + sheet.name + "'!$U$" + str(last_row),
                ]
                #    "='" + sheet.name + "'!$AF$" + str(last_row)]
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
            if system not in skip_sheets:
                (pwb.sheets["Design"].range("21:21")).copy(
                    sheet.range(str(offset) + ":" + str(offset))
                )
                sheet.range("B" + str(offset)).value = str(count) + " ‣ "
                sheet.range("C" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("D" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range(
                    f"G{offset}"
                ).formula = f'=IF(E{offset}<>"OPTION", IF(D{start_row+system_count+2}>0.00001, D{offset}/D{start_row+system_count+2}, ""), "")'  # For scope percentage
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
        (pwb.sheets["Design"].range("15:15")).copy(
            sheet.range(str(start_row) + ":" + str(start_row))
        )
        (pwb.sheets["Design"].range("11:11")).copy(
            sheet.range(str(offset) + ":" + str(offset))
        )
        (pwb.sheets["Design"].range("17:17")).copy(
            sheet.range(str(offset + 1) + ":" + str(offset + 1))
        )

        # sheet = wb.sheets['Summary']
        sheet.range(
            "C" + str(offset + 1)
        ).value = '="TOTAL PROJECT (" & Config!B12 & ")"'
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
            (pwb.sheets["Design"].range("18:18")).copy(
                sheet.range(str(offset + 2) + ":" + str(offset + 2))
            )
            (pwb.sheets["Design"].range("19:19")).copy(
                sheet.range(str(offset + 3) + ":" + str(offset + 3))
            )
            sheet.range(
                "C" + str(offset + 3)
            ).formula = '="TOTAL PROJECT PRICE AFTER DISCOUNT (" & Config!B12 & ")"'
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
            sheet.range(
                "C" + str(offset + 5)
            ).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'
            sheet.range(
                "C" + str(offset + 6)
            ).value = "• Total project price does not include prices for optional items set out in the detailed bill of material."
            sheet.range(
                "C" + str(offset + 7)
            ).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."

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
                    sheet.range(
                        f"J{offset+7+i}"
                    ).formula = f"=CEILING(H{offset+7+i}*I{offset+7+i},1)"
                    sheet.range(
                        f"K{offset+7+i}"
                    ).formula = f"=H{offset+7+i}-J{offset+7+i}"
                    sheet.range(f"L{offset+7+i}").formula = f"=N{offset+1}"
                    sheet.range(
                        f"M{offset+7+i}"
                    ).formula = f"=K{offset+7+i}-L{offset+7+i}"
                    sheet.range(
                        f"N{offset+7+i}"
                    ).formula = f"=M{offset+7+i}/K{offset+7+i}"
                # Format
                sheet.range(
                    f"H{offset+7}:H{offset+7+discount_level}"
                ).number_format = ACCOUNTING
                sheet.range(
                    f"I{offset+7}:I{offset+7+discount_level}"
                ).number_format = "0.00%"
                sheet.range(
                    f"J{offset+7}:M{offset+7+discount_level}"
                ).number_format = ACCOUNTING
                sheet.range(
                    f"N{offset+7}:N{offset+7+discount_level}"
                ).number_format = "0.00%"

        else:
            sheet.range(
                "C" + str(offset + 3)
            ).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'
            sheet.range(
                "C" + str(offset + 4)
            ).value = "• Total project price does not include items marked 'OPTION' in the detailed bill of material."
            sheet.range(
                "C" + str(offset + 5)
            ).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."

    else:
        for sheet in wb.sheet_names:
            if sheet not in skip_sheets:
                sheet = wb.sheets[sheet]
                last_row = sheet.range("G1048576").end("up").row
                collect = [
                    "='" + sheet.name + "'!$C$3",
                    "='" + sheet.name + "'!$G$" + str(last_row),
                    "='" + sheet.name + "'!$U$" + str(last_row),
                ]
                #    "='" + sheet.name + "'!$AF$" + str(last_row
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
            if system not in skip_sheets:
                (pwb.sheets["Design"].range("21:21")).copy(
                    sheet.range(str(offset) + ":" + str(offset))
                )
                sheet.range("B" + str(offset)).value = str(count) + " ‣ "
                sheet.range("C" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range("D" + str(offset)).formula = odered_summary_formula.pop()
                sheet.range(
                    f"G{offset}"
                ).formula = f'=IF(E{offset}<>"OPTION", IF(D{start_row+system_count+2}>0.00001, D{offset}/D{start_row+system_count+2}, ""), "")'  # For scope percentage
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
        (pwb.sheets["Design"].range("13:13")).copy(
            sheet.range(str(start_row) + ":" + str(start_row))
        )
        (pwb.sheets["Design"].range("11:11")).copy(
            sheet.range(str(offset) + ":" + str(offset))
        )
        (pwb.sheets["Design"].range("7:7")).copy(
            sheet.range(str(offset + 1) + ":" + str(offset + 1))
        )

        # sheet = wb.sheets['Summary']
        sheet.range(
            "C" + str(offset + 1)
        ).value = '="TOTAL PROJECT (" & Config!B12 & ")"'
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
            (pwb.sheets["Design"].range("8:8")).copy(
                sheet.range(str(offset + 2) + ":" + str(offset + 2))
            )
            (pwb.sheets["Design"].range("9:9")).copy(
                sheet.range(str(offset + 3) + ":" + str(offset + 3))
            )
            sheet.range(
                "C" + str(offset + 3)
            ).formula = '="TOTAL PROJECT PRICE AFTER DISCOUNT (" & Config!B12 & ")"'
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
            sheet.range(
                "C" + str(offset + 5)
            ).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'
            sheet.range(
                "C" + str(offset + 6)
            ).value = "• Total project price does not include prices for optional items set out in the detailed bill of material."
            sheet.range(
                "C" + str(offset + 7)
            ).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."

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
                    sheet.range(
                        f"J{offset+7+i}"
                    ).formula = f"=CEILING(H{offset+7+i}*I{offset+7+i},1)"
                    sheet.range(
                        f"K{offset+7+i}"
                    ).formula = f"=H{offset+7+i}-J{offset+7+i}"
                    sheet.range(f"L{offset+7+i}").formula = f"=H{offset+1}"
                    sheet.range(
                        f"M{offset+7+i}"
                    ).formula = f"=K{offset+7+i}-L{offset+7+i}"
                    sheet.range(
                        f"N{offset+7+i}"
                    ).formula = f"=M{offset+7+i}/K{offset+7+i}"
                # Format
                sheet.range(
                    f"H{offset+7}:H{offset+7+discount_level}"
                ).number_format = ACCOUNTING
                sheet.range(
                    f"I{offset+7}:I{offset+7+discount_level}"
                ).number_format = "0.00%"
                sheet.range(
                    f"J{offset+7}:M{offset+7+discount_level}"
                ).number_format = ACCOUNTING
                sheet.range(
                    f"N{offset+7}:N{offset+7+discount_level}"
                ).number_format = "0.00%"

        else:
            sheet.range(
                "C" + str(offset + 3)
            ).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'
            sheet.range(
                "C" + str(offset + 4)
            ).value = "• Total project price does not include items marked 'OPTION' in the detailed bill of material."
            sheet.range(
                "C" + str(offset + 5)
            ).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."

    sheet.range("D:D").autofit()
    sheet.range("E:E").autofit()
    sheet.range("F:P").autofit()
    last_row = sheet.range("C1048576").end("up").row
    sheet.page_setup.print_area = "A1:F" + str(last_row + 3)


def number_title(wb, count=10, step=10):
    """
    For the main numbering. It will fix as long as it is a number.
    Need to look for only the systems and engineering services.
    Takes a work book, then start number and step."""
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            system_names.append(str.upper(sheet.name))
            ws = wb.sheets[sheet]
            last_row = ws.range("C1048576").end("up").row
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

    # Need to do try-except as the float type can return nan
    for idx, item in systems["NO"].items():
        try:
            if int(item):
                systems.at[idx, "NO"] = count
                count += step
        except Exception:
            pass

    # Now is the matter of writing to the required sheets
    for system in system_names:
        # print(system)
        sheet = wb.sheets[system]
        system = systems[systems["System"] == system]
        sheet.range("A2").options(index=False).value = system["NO"]


def prepare_to_print_technical(wb):
    """Takes a work book, set horizantal borders at pagebreaks."""
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    # macro_nb = xw.Book('PERSONAL.XLSB')
    current_sheet = wb.sheets.active
    page_setup(wb)
    for sheet in wb.sheet_names:
        if sheet not in skip_sheets:
            last_row = wb.sheets[sheet].range("C1048576").end("up").row
            wb.sheets[sheet].activate()
            wb.sheets[sheet].range("C:C").autofit()
            wb.sheets[sheet].range("C:C").column_width = 60
            wb.sheets[sheet].range("C:C").wrap_text = True
            wb.sheets[sheet].range("D:F").autofit()
            # Adjust the last two rows so that unwanted pagebreak can be prevented
            wb.sheets[sheet].range(f"{last_row+1}:{last_row+1}").delete()
            wb.sheets[sheet].range(f"{last_row+1}:{last_row+1}").row_height = 2
            MACRO_NB.macro("conditional_format")()
            MACRO_NB.macro("remove_h_borders")()
            MACRO_NB.macro("pagebreak_borders")()
    wb.sheets[current_sheet].activate()


def technical(wb):
    directory = os.path.dirname(wb.fullname)
    wb.sheets["Cover"].range("D39").value = "TECHNICAL PROPOSAL"
    wb.sheets["Summary"].range("D20:D100").value = ""
    wb.sheets["Summary"].range("C20:C100").value = (
        wb.sheets["Summary"].range("C20:C100").raw_value
    )

    if wb.name[:9] == "Technical":
        xw.apps.active.alert("The file already seems to be technical.")
        return

    elif wb.name[:10] == "Commercial":
        for sheet in wb.sheet_names:
            ws = wb.sheets[sheet]
            skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
            wb.sheets[2].activate()
            if sheet not in skip_sheets:
                # Require to remove h_borders as these willl not be detected
                # when columns are removed and page setup changed.
                MACRO_NB.macro("remove_h_borders")()
                last_row = ws.range("C1048576").end("up").row
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
        wb.sheets["T&C"].delete()
        prepare_to_print_technical(wb)
        wb.sheets["Summary"].activate()
        file_name = "Technical " + wb.name[11:-4] + "xlsx"
        wb.save(Path(directory, file_name), password="")
        technical_wb = xw.Book(file_name)
        print_technical(technical_wb)
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
            skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
            if sheet not in skip_sheets:
                last_row = ws.range("C1048576").end("up").row
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

        # If T&C does not exist, do nothing.
        try:
            wb.sheets["T&C"].delete()
        except Exception as e:
            pass
        prepare_to_print_technical(wb)
        wb.sheets["Summary"].activate()
        file_name = "Technical " + wb.name[:-4] + "xlsx"
        wb.save(Path(directory, file_name), password="")
        technical_wb = xw.Book(file_name)
        print_technical(technical_wb)


def commercial(wb):
    directory = os.path.dirname(wb.fullname)
    """Takes a work book, set horizantal borders at pagebreaks."""
    # macro_nb = xw.Book('PERSONAL.XLSB')
    # current_sheet = wb.sheets.active
    wb.sheets["Cover"].range("D6:D8").value = (
        wb.sheets["Cover"].range("D6:D8").raw_value
    )
    wb.sheets["Cover"].range("D39").value = wb.sheets["Config"].range("B13").value
    wb.sheets["Cover"].range("C42:C47").value = (
        wb.sheets["Cover"].range("C42:C47").raw_value
    )
    last_row = wb.sheets["Summary"].range("D1048576").end("up").row
    wb.sheets["Summary"].range(f"G20:P{last_row}").value = (
        wb.sheets["Summary"].range(f"G20:P{last_row}").raw_value
    )
    wb.sheets["Summary"].range("C20:C100").value = (
        wb.sheets["Summary"].range("C20:C100").raw_value
    )
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    page_setup(wb)
    for sheet in wb.sheet_names:
        ws = wb.sheets[sheet]
        ws.range("A1").value = ws.range("A1").raw_value  # Remove formula
        ws.range("A1").wrap_text = False
        if sheet not in skip_sheets:
            last_row = ws.range("G1048576").end("up").row
            ws.activate()
            # Adjust column width as sometimes, the long value does not show.
            ws.range(f"A3:AL{last_row}").value = ws.range(f"A3:AL{last_row}").raw_value
            ws.range("A:A").column_width = 4
            ws.range("B:B").autofit()
            ws.range("C:C").autofit()
            ws.range("C:C").column_width = 55
            # wb.sheets[sheet].range('C:C').wrap_text =
            ws.range(
                f"G3:G{last_row-1}"
            ).formula = (
                '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED", H3<>"WAIVED"), D3*F3,"")'
            )
            ws.range(f"G{last_row}").formula = "=SUM(G3:G" + str(last_row - 1) + ")"
            wb.sheets[sheet].range("D:H").autofit()
            ws.range("AM:BD").delete()
            ws.range("I:AK").delete()
            ws.range(f"AM1:AM{last_row}").value = ws.range(f"I1:I{last_row}").raw_value
            ws.range("I:I").delete()
            ws.range("AL:AL").column_width = 0
            # Call macros
            MACRO_NB.macro("conditional_format")()
            MACRO_NB.macro("remove_h_borders")()
            MACRO_NB.macro("pagebreak_borders")()

    wb.sheets["Summary"].range("G:X").delete()
    wb.sheets["Config"].delete()
    wb.sheets["Summary"].activate()
    file_name = "Commercial " + wb.name[:-4] + "xlsx"
    wb.save(Path(directory, file_name), password="")
    commercial_wb = xw.Book(file_name)
    try:
        commercial_wb.to_pdf(show=True)
    except Exception as e:
        # The program does not override the existing file. Therefore, the file needs to be removed if it exists.
        # xw.apps.active.alert('The PDF file already exists!\n Please delete the file and try again.')
        xw.apps.active.alert(
            f"This error is encountered {e}. The PDF file already exists?"
        )


def prepare_to_print_internal(wb):
    """Takes a work book, set horizantal borders at pagebreaks."""
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    # macro_nb = xw.Book('PERSONAL.XLSB')
    current_sheet = wb.sheets.active
    page_setup(wb)
    for sheet in wb.sheet_names:
        if sheet not in skip_sheets:
            wb.sheets[sheet].activate()
            MACRO_NB.macro("conditional_format_internal_costing")()
            MACRO_NB.macro("remove_h_borders")()
            # Below is commented out so that blue lines do not show
            # MACRO_NB.macro('pagebreak_borders')()
    wb.sheets[current_sheet].activate()


def print_technical(wb):
    """The technical proposal will be written to the cwd."""
    try:
        wb.to_pdf(show=True)
    except Exception:
        # The program does not override the existing file. The file needs to be removed if it exists.
        xw.apps.active.alert(
            "The PDF file already exists!\n Please delete the file and try again."
        )


def conditional_format_wb(wb):
    """
    Takes a workbook, and do conditional formatting.
    Rely on excel macro for conditional format.
    """
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    # macro_nb = xw.Book('PERSONAL.XLSB')
    current_sheet = wb.sheets.active
    for sheet in wb.sheet_names:
        if sheet not in skip_sheets:
            wb.sheets[sheet].activate()
            MACRO_NB.macro("conditional_format")()
            # Remove H borders in original excel
            MACRO_NB.macro("remove_h_borders")()
            # Fix the columns border
            MACRO_NB.macro("format_column_border")()
    wb.sheets[current_sheet].activate()


def fix_unit_price(wb):
    """
    Fix unit prices, normally done for subsequent revisions.
    """
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            system_names.append(str.upper(sheet.name))
            ws = wb.sheets[sheet]
            last_row = ws.range("C1048576").end("up").row
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
    """
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            system_names.append(str.upper(sheet.name))
            ws = wb.sheets[sheet]
            last_row = ws.range("C1048576").end("up").row
            data = (
                ws.range("C2:AL" + str(last_row))
                .options(pd.DataFrame, empty="", index=False)
                .value
            )
            data["System"] = str.upper(sheet.name)
            # format_type = ws.range('AL2:AL' + str(last_row)).options(pd.DataFrame, empty='', index=False).value
            systems = pd.concat([systems, data], join="outer")
            # systems = pd.concat([systems, format_type], join='outer')

    systems = systems.reset_index(
        drop=True
    )  # Otherwise separate sheet will have own index.
    systems = systems.reindex(
        columns=["Description", "Unit", "Scope", "Format", "System"]
    )

    for idx, item in systems["Description"].items():
        systems.at[idx, "Description"] = set_nitty_gritty(
            str(systems.loc[idx, "Description"]).strip().lstrip("• ")
        )
        # Set unit description
        # Set to lower case
        systems.at[idx, "Unit"] = str(systems.loc[idx, "Unit"]).strip().lower()
        # Change to ea
        if str(systems.loc[idx, "Unit"]) in ["nos", "no"]:
            systems.at[idx, "Unit"] = "ea"
        if str(systems.loc[idx, "Unit"])[-1:] == "s":
            systems.at[idx, "Unit"] = str(systems.loc[idx, "Unit"])[:-1]

        systems.at[idx, "Scope"] = str(systems.loc[idx, "Scope"]).strip().lower()
        if str(systems.loc[idx, "Scope"]) in ["inclusive", "include", "included"]:
            systems.at[idx, "Scope"] = "INCLUDED"
        if str(systems.loc[idx, "Scope"]) in ["option", "optional"]:
            systems.at[idx, "Scope"] = "OPTION"
        if str(systems.loc[idx, "Scope"]) in ["waived"]:
            systems.at[idx, "Scope"] = "WAIVED"

        if indent_description:
            if systems.at[idx, "Format"] == "Description":
                systems.at[idx, "Description"] = "   " + (
                    str(systems.loc[idx, "Description"]).strip()
                ).lstrip("• ")
                if bullet_description:
                    if str(systems.loc[idx, "Description"]).strip().startswith("#"):
                        systems.at[idx, "Description"] = "      ‣ " + str(
                            systems.loc[idx, "Description"]
                        ).strip().lstrip("# ")
                    elif str(systems.loc[idx, "Description"]).strip().startswith("‣"):
                        systems.at[idx, "Description"] = "      ‣ " + str(
                            systems.loc[idx, "Description"]
                        ).strip().lstrip("‣ ")
                    else:
                        systems.at[idx, "Description"] = (
                            "   • " + str(systems.loc[idx, "Description"]).strip()
                        )

        if title_lineitem_or_description:
            if systems.at[idx, "Format"] == "Lineitem":
                if len(str(systems.loc[idx, "Description"])) <= 60:
                    systems.at[idx, "Description"] = set_case_preserve_acronym(
                        (str(systems.loc[idx, "Description"]).strip()).lstrip("• "),
                        title=True,
                    )

            if systems.at[idx, "Format"] == "Description":
                if len(str(systems.loc[idx, "Description"])) <= 60:
                    systems.at[idx, "Description"] = set_case_preserve_acronym(
                        (str(systems.loc[idx, "Description"]).strip()).lstrip("• "),
                        title=True,
                    )

        if upper_title:
            if systems.at[idx, "Format"] == "Title":
                systems.at[idx, "Description"] = (
                    str(systems.loc[idx, "Description"]).strip().upper()
                )

        if upper_system:
            if systems.at[idx, "Format"] == "System":
                systems.at[idx, "Description"] = (
                    str(systems.loc[idx, "Description"]).strip().upper()
                )

    # Write fomatted description to Description field
    for system in system_names:
        sheet = wb.sheets[system]
        system = systems[systems["System"] == system]
        # sheet.range('C2').value = sheet.range('C2').options(empty='')
        sheet.range("C2").options(index=False).value = system["Description"]
        sheet.range("E2").options(index=False).value = system["Unit"]
        sheet.range("H2").options(index=False).value = system["Scope"]


def indent_description(wb):
    """
    Depricated.
    Indent description
    This function works but slow. Replaced with 'format_text' function
    """
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            ws = wb.sheets[sheet]
            last_row = ws.range("C1048576").end("up").row
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
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    # macro_nb = xw.Book('PERSONAL.XLSB')
    current_sheet = wb.sheets.active
    for sheet in wb.sheet_names:
        if sheet not in skip_sheets:
            wb.sheets[sheet].activate()
            if shaded:
                MACRO_NB.macro("shaded")()
            else:
                MACRO_NB.macro("unshaded")()
    wb.sheets[current_sheet].activate()


def internal_costing(wb):
    directory = os.path.dirname(wb.fullname)

    wb.sheets["Cover"].range("D39").value = "INTERNAL COSTING"
    wb.sheets["Cover"].range("C42:C47").value = (
        wb.sheets["Cover"].range("C42:C47").raw_value
    )
    wb.sheets["Cover"].range("D6:D8").value = (
        wb.sheets["Cover"].range("D6:D8").raw_value
    )

    summary_last_row = wb.sheets["Summary"].range("D1048576").end("up").row
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

    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    for sheet in wb.sheet_names:
        ws = wb.sheets[sheet]
        ws.range("A1").value = ws.range("A1").raw_value  # Remove formula
        if sheet not in skip_sheets:
            # Collect escalation
            escalation = ws.range("K1:R1").value
            ws.range("I1:R1").value = ""
            # Construct as dictionary
            escalation = dict(zip(escalation[::2], escalation[1::2]))

            # Work on columns
            last_row = ws.range("G1048576").end("up").row
            ws.range("B3:B" + str(last_row)).value = ws.range(
                "B3:B" + str(last_row)
            ).raw_value
            ws.range("F3:G" + str(last_row)).value = ""
            # ws.range('K3:Q'+ str(last_row)).value = ws.range('K3:Q'+ str(last_row)).raw_value
            ws.range("Q3:Q" + str(last_row)).value = ws.range(
                "Q3:Q" + str(last_row)
            ).raw_value
            # ws.range('AM:AM').delete()
            ws.range("R:AK").delete()
            ws.range("W3").value = escalation
            ws.range("W7").value = "Total"
            ws.range("X7").formula = "=SUM(X3:X6)"
            ws.range("X3:X7").number_format = "0.00%"
            # Insert Escalation column
            ws.range("S:S").insert("right")
            ws.range("S2").value = "Escalation"
            ws.range(
                "S3:S" + str(last_row)
            ).formula = '=IF(AND(D3<>"", J3<>"",K3<>""), $Y$7, "")'
            ws.range("S3:S" + str(last_row)).number_format = "0.00%"

            # To reduce visual clutter
            ws.range("D:X").autofit()
            ws.range("I:I").column_width = 20
            ws.range("P:P").column_width = 20
            ws.range("F:G").column_width = 0
            ws.range("R:R").column_width = 0
    wb.sheets["Config"].delete()
    # wb.sheets['T&C'].delete()
    prepare_to_print_internal(wb)
    wb.sheets["Summary"].activate()
    file_name = "Internal " + wb.name[:-4] + "xlsx"
    wb.save(Path(directory, file_name), password="")


def convert_legacy(wb):
    directory = os.path.dirname(wb.fullname)

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
        # skip_sheets = ['FX', 'Cover', 'Intro', 'ES', 'T&C']
        skip_sheets = [
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
        df = pd.DataFrame(columns=full_column_list)
        risk = 0.05
        # Read and set currency from FX sheet
        fx = wb.sheets["FX"]
        exchange_rates = dict(fx.range("A2:B9").value)
        quoted_currency = fx.range("B12").value
        project_info = dict(fx.range("A36:B46").value)
        try:
            project_info = {key: value.upper() for key, value in project_info.items()}
        except:
            xw.apps.active.alert("Project Info items cannot be empty value.")
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
            if sheet not in skip_sheets:
                system_names.append(sheet.upper())
                ws = wb.sheets[sheet]
                escalation = dict(ws.range("K2:L5").value)
                default_mu = ws.range("H5").value
                escalation["default_mu"] = default_mu
                defaults[sheet.upper()] = escalation
                last_row = ws.range("D100000").end("up").row  # Returns a number
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
        es_last_row = es.range("D100000").end("up").row
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
            if sheet in ["Summary", "Technical_Notes", "T&C"]:
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
            last_row = sheet.range("G100000").end("up").row
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
        except:
            xw.apps.active.alert("The file already exists. Please save manually.")

    else:
        xw.apps.active.alert("The excel file does not seem to be legacy template.")


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
        if sheet.name in ["Technical_Notes", "T&C"]:
            sheet.range("A:A").column_width = 2
            sheet.range("B:B").autofit()
            sheet.range("C:C").column_width = 70
            sheet.range("C:C").rows.autofit()
            sheet.range("C:C").wrap_text = True


def fill_formula_active_row(wb, ws):
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    if ws.name not in skip_sheets:
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
    c_column = ws.range("C1048576").end("up").row
    g_column = ws.range("G1048576").end("up").row
    if c_column > g_column:
        last_row = c_column
    else:
        last_row = g_column

    empty_row = 0
    start_row = 0

    for i in range(1, last_row + 1):
        # Check if the below columns are empty
        if all(cell.value is None for cell in ws.range(f"A{i}:H{i}")):
            empty_row += 1
            if empty_row == 1:
                # Set start_row to empty_row
                start_row = i + 1
        else:
            # Check if there is more empty rows
            if empty_row >= 2:
                # Delete the empty rows
                ws.range(f"A{start_row}:XFD{i-1}").delete(shift="up")
                # Make adjustment
                last_row -= empty_row
                i -= empty_row
            # Reset empty_row and start_row
            empty_row = 0
            start_row = 0


def delete_extra_empty_row_wb(wb):
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]

    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            delete_extra_empty_row(sheet)


def format_cell_data(wb):
    """
    Set the cell font and font size
    Format the cell data to correct number or text representation.
    E.g. 1,000.00 or 1.00%
    """
    skip_sheets = ["Config", "Cover", "Summary", "Technical_Notes", "T&C"]
    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            last_row = sheet.range("C1048576").end("up").row + 1
            # Set cell font and size
            sheet.range(f"A3:BD{last_row}").font.name = "Arial"
            sheet.range("2:2").font.size = 9
            sheet.range(f"A3:BD{last_row}").font.size = 12
            sheet.range("C3").font.size = 14
            # Set cell format
            sheet.range("A:B").number_format = "0"
            sheet.range("D:D").number_format = "0"
            sheet.range("F:G").number_format = ACCOUNTING
            sheet.range("K:L").number_format = ACCOUNTING
            sheet.range("M:M").number_format = "0.00%"
            sheet.range("N:O").number_format = ACCOUNTING
            sheet.range("Q:Q").number_format = EXCNANGE_RATE
            sheet.range("R:Z").number_format = ACCOUNTING
            sheet.range("MU:MU").number_format = "0.00%"
            sheet.range("AB:AG").number_format = ACCOUNTING
            sheet.range("AH:AH").number_format = "0.00%"
            sheet.range("AI:AJ").number_format = ACCOUNTING
            sheet.range("I1:R1").number_format = "0.00%"
            # Delet 'Catergory' and 'System' fields to avoid visual clutter.
            if sheet.range("AN2").value == "System":
                sheet.range("AN:AN").delete()
            if sheet.range("AM2").value == "Category":
                sheet.range("AM:AM").delete()
            sheet.range("AM2").value = "Leadtime"
            sheet.range("AN2").value = "Supplier"
            sheet.range("AO2").value = "Maker"


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


def creat_new_template():
    filename = "Template.xltx"
    file_path = Path(RESOURCES, filename)
    wb = xw.Book.caller()
    wb.app.books.open(file_path.absolute(), password=hide.legacy)


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


def creat_new_planner():
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
    except:
        current_wb_revision = None
        current_minor_revision = None
    if current_wb_revision is None or current_wb_revision < int(LATEST_WB_VERSION[1:]):
        wb.sheets["Config"].range("D1:I20").clear()
        wb.sheets["Config"].range("95:106").delete()
        MACRO_NB.sheets["Design"].range("A28:E36").copy(wb.sheets["Config"].range("D2"))
        MACRO_NB.sheets["Data"].range("C1:C2").copy(wb.sheets["Config"].range("B95"))
        MACRO_NB.sheets["Data"].range("D1:D2").copy(wb.sheets["Config"].range("C95"))
        wb.sheets["Config"].range("A15").value = "Template Version"
        wb.sheets["Config"].range("B15").value = LATEST_WB_VERSION
        # Put currency and proposal type validation
        wb.sheets["Config"].activate()
        MACRO_NB.macro("put_currency_proposal_validation_formula")()
        flag += 1

    if current_minor_revision is None or current_minor_revision < int(cc.LATEST_MINOR_REVISION[1:]):
        wb.sheets["Config"].range("C15").value = cc.LATEST_MINOR_REVISION
        # Clear previous data if any
        last_row = wb.sheets["Config"].range("A1048576").end("up").row
        if last_row > 95:
            wb.sheets["Config"].range(f"A95:A{last_row}").clear()
        wb.sheets["Config"].range("A95").value = "SYSTEMS"
        # Write data from list
        cc.available_system_checklist_register.sort()
        wb.sheets["Config"].range("A96").options(transpose=True).value = [
            system.upper() for system in cc.available_system_checklist_register
        ]

        # Test if value "Systems" is already there
        cell_value = wb.sheets["Technical_Notes"].range("F3")
        # if cell_value is None:
        if cell_value != "Systems".upper():
            MACRO_NB.sheets["Data"].range("B1").copy(
                wb.sheets["Technical_Notes"].range("F3")
            )
            # Call macro to fill in the dropbown formula
            wb.sheets["Technical_Notes"].activate()
            MACRO_NB.macro("put_systems_validation_formula")()

        # For general checklist
        # Clear previous data if any
        last_row = wb.sheets["Config"].range("E1048576").end("up").row
        if last_row > 95:
            wb.sheets["Config"].range(f"E95:E{last_row}").clear()
        wb.sheets["Config"].range("E95").value = "CHECKLISTS"
        # Write data from list
        cc.available_checklist_register.sort()
        wb.sheets["Config"].range("E96").options(transpose=True).value = [
            system.upper() for system in cc.available_checklist_register
        ]

        # Test if value "Systems" is already there
        cell_value = wb.sheets["Technical_Notes"].range("G3")
        # if cell_value is None:
        if cell_value != "Checklists".upper():
            MACRO_NB.sheets["Data"].range("E1").copy(
                wb.sheets["Technical_Notes"].range("G3")
            )
            # Call macro to fill in the dropbown formula
            wb.sheets["Technical_Notes"].activate()
            MACRO_NB.macro("put_checklists_validation_formula")()

            wb.sheets["Technical_Notes"].range("F:G").autofit()
        flag += 1

    if flag:
        wb.sheets[current_sheet].activate()
        xw.apps.active.alert(f"The template has been updated to {LATEST_WB_VERSION}.{cc.LATEST_MINOR_REVISION}")
    else:
        message = """           
        No update is required. If you want to force an update, delete "Template Version" in cell "B15" & "C15" in "Config" sheet.
        Advisable to force an update if system or checklist is not available in dropdown list in "Technical_Notes".
        If item is not available in dropdown after forced update, there is no checklist or checklist is not ready.
        """
        xw.apps.active.alert(f"{message}")
