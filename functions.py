""" Multiple functions to support Excel automation.
    © Thiha Aung
"""

import re
import os
from pathlib import Path
import pandas as pd
import xlwings as xw

def set_nitty_gritty(text):
    """Fix annoying text"""
    # Strip EOL
    text = text.strip()
    # Strip 2 or more spaces
    text = re.sub(' {2,}', ' ',  text)
    # Put bullet point for Sub-subitem preceded by '-' or '~'.
    text = re.sub('^(-|~)', '•', text)
    # Put bullet point for Sub-subitem preceded by a single * followed by space.
    text = re.sub('^[*?]\s', ' • ', text)
    # Instead of ';' at the end of line, use ':' instead.
    text = re.sub(';$', ':', text)
    text = set_comma_space(text)
    text = set_x(text)
    return text

def set_comma_space(text):
    """Fix having space before comma and not having space after comma"""
    # fix word+space+, to word+,
    x = re.compile('\w+\s,')
    if x.search(text):
        substring = re.findall('\w+\s,', text)
        for word in substring:
            text = re.sub(word, word[:-2] + ',', text)
    
    # Fix word+,+no-space to word+,+space
    x = re.compile(',\w+')
    if x.search(text):
        substring = re.findall(',\w+', text)
        for word in substring:
            text = re.sub(word, ', ' + word[1:], text)
    return text

def title_case_ignore_single_char(text):
    words = text.split()
    titled_words = []
    for word in words:
        if len(word) > 2: # So that two letter words are ignored.
            titled_words.append(word.title())
        else:
            titled_words.append(word)
    return " ".join(titled_words)

def set_case_preserve_acronym(text, title=False, capitalize=False, upper=False):
    """ Maintaion acronyms case when using title or sentence"""
    # The regex below essentially ignore the letters in lower case letter.
    # Now cases such as iPhone, mPower are recognized.
    acronym_regex = re.compile(r'\b([a-z0-9\.]?[A-Z0-9][A-Z0-9a-z-]*)(?=\b|[^a-z])')
    # acronym_regex = re.compile(r'\b([a-z]?[A-Z0-9][A-Z0-9-]*)(?=\b|[^a-z])')
    acronyms = acronym_regex.findall(text)

    if title:
        text = title_case_ignore_single_char(text)
        # Restore acronyms
        for acronym in acronyms:
            text = text.replace(acronym.title(), acronym)
        return text
   
    elif capitalize:
        text = text.capitalize()
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
    # For cases such as 20x, 30X
    x = re.compile('(\d+x|\d+X)')
    if x.search(text):
        substring = re.findall('(\d+x|\d+X)', text)
        for word in substring:
            text = re.sub(word, (word[:-1] + ' x'), text)
    # For cases such as x20, X30
    x = re.compile('(x\d+|X\d+)')
    if x.search(text):
        substring = re.findall('(x\d+|X\d+)', text)
        for word in substring:
            text = re.sub(word, ('x ' + word[1:]), text)
    # For cases such as 20 X, 30 X
    x = re.compile('(\d+ X)')
    if x.search(text):
        substring = re.findall('(\d+ X)', text)
        for word in substring:
            text = re.sub(word, (word[:-1] + 'x'), text)
    # For cases such as X 20, X 30
    x = re.compile('(X \d+)')
    if x.search(text):
        substring = re.findall('(X \d+)', text)
        for word in substring:
            text = re.sub(word, ('x' + word[1:]), text)
    return text

def fill_formula(sheet):
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    if sheet.name not in skip_sheets:
        # Formula to cells
        last_row = sheet.range('C100000').end('up').row
        sheet.range('A1').formula = '= "JASON REF: " & Config!B29 &  ", REVISION: " &  Config!B30 & ", PROJECT: " & Config!B26'  # noqa: E501
        # Serail Numbering (SN)
        sheet.range('B3:B' + str(last_row)).formula = '=IF(AND(ISNUMBER(D3), ISNUMBER(K3), XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)), COUNT(INDIRECT(CONCAT("B",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1),":B",ROW()-1))) + 1, "")'  # noqa: E501
        sheet.range('N3:N' + str(last_row)).formula = '=IF(K3<>"",K3*(1-M3),"")'
        sheet.range('O3:O' + str(last_row)).formula = '=IF(AND(D3<>"", K3<>"",H3<>"OPTION"),D3*N3,"")'  # noqa: E501
        # Exchange rates
        sheet.range('Q3:Q' + str(last_row)).formula = '=IF(Config!B12="SGD",IF(J3<>"",VLOOKUP(J3,Config!$A$2:$B$10,2,FALSE),""),IF(J3<>"",VLOOKUP(J3,Config!$A$2:$B$10,2,FALSE)/VLOOKUP(Config!$B$12,Config!$A$2:$B$10,2,FALSE),""))'  # noqa: E501
        sheet.range('R3:R' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"") ,N3*Q3,"")'  # noqa: E501
        # sheet.range('S3:S' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION") ,D3*R3,"")'  # noqa: E501
        sheet.range('S3:S' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION",INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))<>"OPTION"),D3*R3,"")'  # noqa: E501
        sheet.range('T3:T' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""), (R3*(1+$L$1+$N$1+$P$1+$R$1))/(1-0.05),"")'  # noqa: E501
        sheet.range('U3:U' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION",INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))<>"OPTION"),D3*T3,"")'  # noqa: E501
        # Default
        sheet.range('V3:V' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),D3*R3*$L$1,"")'  # noqa: E501
        # Warranty
        sheet.range('W3:W' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),D3*R3*$N$1,"")'  # noqa: E501
        # Freight (Inbound)
        sheet.range('X3:X' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),D3*R3*$P$1,"")'  # noqa: E501
        # Special (Condition)
        sheet.range('Y3:Y' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),R3*$R$1,"")'  # noqa: E501
        # Risk
        sheet.range('Z3:Z' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",U3<>""),U3-(S3+V3+W3+X3+Y3),"")'  # noqa: E501
        sheet.range('AA3:AA' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),$J$1,"")'  # noqa: E501
        sheet.range('AC3:AC' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),CEILING(T3/(1-AA3), 1),"")'  # noqa: E501
        # sheet.range('AD3:AD' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION",H3<>"INCLUDED"),D3*AC3,"")'  # noqa: E501
        sheet.range('AD3:AD' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION",H3<>"INCLUDED",(INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))) <>"OPTION"),D3*AC3,"")'  # noqa: E501
        sheet.range('AE3:AE' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),IF(AB3<>"",AB3,AC3),"")'  # noqa: E501
        # sheet.range('AF3:AF' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED"),D3*AE3,"")'  # noqa: E501
        sheet.range('AF3:AF' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED",(INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))) <>"OPTION"),D3*AE3,"")'  # noqa: E501
        # sheet.range('AF3:AF' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),D3*AE3,"")'  # noqa: E501
        sheet.range('AG3:AG' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED",AF3<>""),AF3-U3,"")'  # noqa: E501
        sheet.range('AH3:AH' + str(last_row)).formula = '=IF(AND(AG3<>"",AG3<>0),AG3/AF3,"")'  # noqa: E501
        sheet.range('AI3:AI' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION"),D3*AE3,"")'  # noqa: E501
        # sheet.range('AL3:AL' + str(last_row)).formula = '=IF(A3<>"","Title",IF(B3<>"","Lineitem",IF(LEFT(C3,3)="***","Comment",IF(AND(A3="",B3="",C2="", C4<>"",D4<>""), "Subtitle",""))))'  # noqa: E501
        # Unit Price
        sheet.range('F3:F' + str(last_row)).formula = '=IF(AND(AL3="Title", ISNUMBER(AJ3)), AJ3, IF(AND(AL3="Lineitem", AK3="Lumpsum", H3<>"OPTION"), "", AE3))'  # noqa: E501
        # sheet.range('F3:F' + str(last_row)).formula = '=IF(AE3<>"", AE3,"")'
        sheet.range('G3:G' + str(last_row)).formula = '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED"), D3*F3,"")'  # noqa: E501
        sheet.range('L3:L' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION"),D3*K3,"")'  # noqa: E501
        # For Format field
        sheet.range('AL1').value = "Title"
        sheet.range('AL4:AL' + str(last_row)).formula = '=IF(C4<>"",IF(AND(A4<>"",C4<>""),"Title", IF(B4<>"","Lineitem", IF(LEFT(C4,3)="***","Comment", IF(AND(A4="",B4="",C3="", C5<>"",D5<>""), "Subtitle","Description")))),"")'  # noqa: E501
        sheet.range('AL' + str(last_row+1)).value = "Title"

        # For Lumpsum
        # sheet.range('AJ3:AJ' + str(last_row)).formula = '=IF(AND(AL3="Title", D3=1, E3="lot"), SUM(INDIRECT(CONCAT("AF", ROW()+1, ":AF",((MATCH("Title",INDIRECT(CONCAT("AL", ROW()+1, ":AL", MATCH(REPT("z",50),AL:AL))),0)) + ROW())))), "")'  # noqa: E501
        sheet.range('AJ3:AJ' + str(last_row)).formula = '=IF(AND(AL3="Title", D3=1, E3="lot"), SUM(INDIRECT(CONCAT("AI", ROW()+1, ":AI",((MATCH("Title",INDIRECT(CONCAT("AL", ROW()+1, ":AL", MATCH(REPT("z",50),AL:AL))),0)) + ROW())))), "")'  # noqa: E501
        sheet.range('AK3:AK' + str(last_row)).formula = '=IF(AL3="Lineitem", IF(ISNUMBER(INDIRECT(CONCAT("AJ",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))), "Lumpsum", "Unit Price"), "")'  # noqa: E501

def fill_formula_wb(wb):
    for sheet in wb.sheets:
        fill_formula(sheet)

def fill_lastrow (wb):
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']

    for sheet in wb.sheets:
       if sheet.name not in skip_sheets:
            fill_lastrow_sheet(wb, sheet)

def fill_lastrow_sheet(wb, sheet):
    pwb = xw.books('PERSONAL.XLSB')
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    if sheet.name not in skip_sheets:
        last_row = sheet.range('C100000').end('up').row
        (pwb.sheets['Design'].range('5:5')).copy(sheet.range(str(last_row+2) + ':' + str(last_row+2)))  # noqa: E501
        sheet.range('F'+ str(last_row+2)).formula = '="Subtotal(" & Config!B12 & ")"'
        sheet.range('F'+ str(last_row+2)).font.bold = True
        sheet.range('F'+ str(last_row+2)).font.size = 9
        sheet.range('G' + str(last_row+2)).formula = '=SUM(G3:G' + str(last_row+1) + ')'
        sheet.range('G' + str(last_row+2)).font.bold = True
        # SCDQ: Subtotal cost after discount in quoted currency
        sheet.range('S' + str(last_row+2)).formula = '=SUM(S3:S' + str(last_row+1) + ')'
        sheet.range('S' + str(last_row+2)).font.bold = True
        sheet.range('S' + str(last_row+2)).font.color = (0, 144, 81)
        # BSCQ: Base subtotal cost in quoted currency
        sheet.range('U' + str(last_row+2)).formula = '=SUM(U3:U' + str(last_row+1) + ')'
        sheet.range('U' + str(last_row+2)).font.bold = True
        # Default
        sheet.range('V' + str(last_row+2)).formula = '=SUM(V3:V' + str(last_row+1) + ')'
        sheet.range('V' + str(last_row+2)).font.bold = True
        sheet.range('V' + str(last_row+2)).font.color = (0, 144, 81)
        # Warranty
        sheet.range('W' + str(last_row+2)).formula = '=SUM(W3:W' + str(last_row+1) + ')'
        sheet.range('W' + str(last_row+2)).font.bold = True
        sheet.range('W' + str(last_row+2)).font.color = (0, 144, 81)
        # Freight (Inbound)
        sheet.range('X' + str(last_row+2)).formula = '=SUM(X3:X' + str(last_row+1) + ')'
        sheet.range('X' + str(last_row+2)).font.bold = True
        sheet.range('X' + str(last_row+2)).font.color = (0, 144, 81)
        # Special (Conditions)
        sheet.range('Y' + str(last_row+2)).formula = '=SUM(Y3:Y' + str(last_row+1) + ')'
        sheet.range('Y' + str(last_row+2)).font.bold = True
        sheet.range('Y' + str(last_row+2)).font.color = (0, 144, 81)
        # Risk
        sheet.range('Z' + str(last_row+2)).formula = '=SUM(Z3:Z' + str(last_row+1) + ')'
        sheet.range('Z' + str(last_row+2)).font.bold = True
        sheet.range('Z' + str(last_row+2)).font.color = (0, 144, 81)
        sheet.range('AF' + str(last_row+2)).formula = '=SUM(AF3:AF' + str(last_row+1) + ')'  # noqa: E501
        sheet.range('AF' + str(last_row+2)).font.bold = True
        sheet.range('AG' + str(last_row+2)).formula = '=SUM(AG3:AG' + str(last_row+1) + ')'  # noqa: E501
        sheet.range('AG' + str(last_row+2)).font.bold = True
        sheet.range('AH' + str(last_row+2)).formula = '=AG' + str(last_row+2) + '/AF' + str(last_row+2)  # noqa: E501
        sheet.range('AH' + str(last_row+2)).font.bold = True

        # Set-up print area
        sheet.page_setup.print_area = 'A1:H' + str(last_row+2)

def unhide_columns(sheet):
    """Unhide all columns while setting the width for selected columns"""
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    if sheet.name not in skip_sheets:
        sheet.range('A:A').column_width = 4
        sheet.range('B:B').autofit()
        sheet.range('C:C').column_width = 55
        sheet.range('C:C').rows.autofit()
        sheet.range('C:C').wrap_text = True
        sheet.range('D:H').autofit()
        sheet.range('J:O').autofit()
        sheet.range('Q:AN').autofit()

def hide_columns(sheet):
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    if sheet.name not in skip_sheets:
        sheet.range('AI:AN').column_width = 0
        sheet.range('AC:AD').column_width = 0
        sheet.range('AF:AF').column_width = 0
        sheet.range('AB:AB').column_width = 10
        sheet.range('S:AA').column_width = 0
        sheet.range('Q:Q').column_width = 0
        sheet.range('P:P').column_width = 20
        sheet.range('O:O').column_width = 0
        sheet.range('L:L').column_width = 0
        sheet.range('T:T').autofit()

def summary(wb, discount=False, detail=False):
    summary_formula = []
    collect = [] # Collect formula to be put in summary page.
    formula_fragment = '=IF(OR(Config!B13="COMMERCIAL PROPOSAL", Config!B13="BUDGETARY PROPOSAL"),'  # noqa: E501
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    # The design will now be taken from PERSONAL.XLSB
    pwb = xw.books('PERSONAL.XLSB')
    n_format = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    if not detail:
        for sheet in wb.sheet_names:
            if sheet not in skip_sheets:
                sheet = wb.sheets[sheet]
                last_row = sheet.range('G100000').end('up').row
                collect = [ formula_fragment + "'" + sheet.name + "'!$G$" + str(last_row) + ', "")',  # noqa: E501
                            formula_fragment + "'" + sheet.name + "'!$U$" + str(last_row) + ', "")']  # noqa: E501
                    #    "='" + sheet.name + "'!$AF$" + str(last_row)]
                summary_formula.extend(collect)
                collect = []

        start_row = 19
        count = 1
        offset = 20
        odered_summary_formula = summary_formula[::-1]
        sheet = wb.sheets['Summary']
        # Clear summary page first
        sheet.range('A18:Z1000').clear()
        # Set format
        sheet.range('C:C').column_width = 55
        # sheet.range('E20:E1000').horizontal_alignment = 'center'
 
        for system in wb.sheet_names:
            if system not in skip_sheets:
                (pwb.sheets['Design'].range('21:21')).copy(sheet.range(str(offset) + ':' + str(offset)))  # noqa: E501
                sheet.range('B' + str(offset)).value = count
                sheet.range('C' + str(offset)).value = system
                sheet.range('D' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('D' + str(offset)).number_format = n_format
                sheet.range('H' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('H' + str(offset)).number_format = n_format
                sheet.range('I' + str(offset)).formula = '=IF(H'+ str(offset) + '<>"",D' + str(offset) + '- H' + str(offset) + ',"")'  # noqa: E501
                sheet.range('I' + str(offset)).number_format = n_format
                sheet.range('J' + str(offset)).formula = '=IF(I' + str(offset) + '<>"",I' + str(offset) + '/D' + str(offset) + ',"")'  # noqa: E501
                sheet.range('J' + str(offset)).number_format = '0.00%'
                count += 1
                offset += 1
        
        # Drawing lines
        (pwb.sheets['Design'].range('13:13')).copy(sheet.range(str(start_row) + ':' + str(start_row)))  # noqa: E501
        (pwb.sheets['Design'].range('11:11')).copy(sheet.range(str(offset) + ':' + str(offset)))  # noqa: E501
        (pwb.sheets['Design'].range('7:7')).copy(sheet.range(str(offset+1) + ':' + str(offset+1)))  # noqa: E501
        # sheet = wb.sheets['Summary']
        sheet.range('C' + str(offset+1)).value = '="TOTAL PROJECT (" & Config!B12 & ")"'
        sheet.range('D' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",D20:D' + str(offset) + ')'  # noqa: E501
        sheet.range('E' + str(offset+1)).formula = '=IF(COUNTIF(E20:E' + str(offset) + ',"OPTION"), "Excluding Option", "")'  # noqa: E501
        sheet.range('H' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",H20:H' + str(offset) + ')'  # noqa: E501
        sheet.range('I' + str(offset+1)).formula = '=IF(H'+ str(offset+1) + '<>"", D' + str(offset+1) + '- H' + str(offset+1) + ',"")'  # noqa: E501
        sheet.range('J' + str(offset+1)).formula = '=IF(I' + str(offset+1) + '<>0,I' + str(offset+1) + '/D' + str(offset+1) + ',"")'  # noqa: E501
        if discount:
            (pwb.sheets['Design'].range('8:8')).copy(sheet.range(str(offset+2) + ':' + str(offset+2)))  # noqa: E501
            (pwb.sheets['Design'].range('9:9')).copy(sheet.range(str(offset+3) + ':' + str(offset+3)))  # noqa: E501
            sheet.range('C' + str(offset+3)).formula = '="TOTAL PROJECT PRICE AFTER DISCOUNT (" & Config!B12 & ")"'  # noqa: E501
            sheet.range('D' + str(offset+3)).formula = '=SUM(D' +str(offset+1) + ':D' + str(offset+2) + ')'  # noqa: E501
            sheet.range('D' + str(offset+3)).number_format = n_format
            sheet.range('H' + str(offset+3)).formula = '=$H$' +str(offset+1)
            sheet.range('H' + str(offset+3)).number_format = n_format
            sheet.range('I' + str(offset+3)).formula = '=IF(H'+ str(offset+3) + '<>"", D' + str(offset+3) + '- H' + str(offset+3) + ',"")'  # noqa: E501
            sheet.range('I' + str(offset+3)).number_format = n_format
            sheet.range('J' + str(offset+3)).formula = '=IF(I' + str(offset+3) + '<>0,I' + str(offset+3) + '/D' + str(offset+3) + ',"")'  # noqa: E501
            sheet.range('J' + str(offset+3)).number_format = '0.00%'
            sheet.range('C' + str(offset+5)).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'  # noqa: E501
            sheet.range('C' + str(offset+6)).value = "• Total project price does not include prices for optional items set out in the detailed bill of material."  # noqa: E501
            sheet.range('C' + str(offset+7)).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."  # noqa: E501
        else:
            sheet.range('C' + str(offset+3)).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'  # noqa: E501
            sheet.range('C' + str(offset+4)).value = "• Total project price does not include items marked 'OPTION' in the detailed bill of material."  # noqa: E501
            sheet.range('C' + str(offset+5)).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."  # noqa: E501
    else: #detail
        for sheet in wb.sheet_names:
            if sheet not in skip_sheets:
                sheet = wb.sheets[sheet]
                last_row = sheet.range('G100000').end('up').row
                collect = [ formula_fragment + "'" + sheet.name + "'!$G$" + str(last_row) + ', "")',  # noqa: E501
                            formula_fragment + "'" + sheet.name + "'!$S$" + str(last_row) + ', "")',  # noqa: E501
                            formula_fragment + "'" + sheet.name + "'!$V$" + str(last_row) + ', "")',  # noqa: E501
                            formula_fragment + "'" + sheet.name + "'!$W$" + str(last_row) + ', "")',  # noqa: E501
                            formula_fragment + "'" + sheet.name + "'!$X$" + str(last_row) + ', "")',  # noqa: E501
                            formula_fragment + "'" + sheet.name + "'!$Y$" + str(last_row) + ', "")',  # noqa: E501
                            formula_fragment + "'" + sheet.name + "'!$Z$" + str(last_row) + ', "")',  # noqa: E501
                            formula_fragment + "'" + sheet.name + "'!$U$" + str(last_row) + ', "")']  # noqa: E501
                    #    "='" + sheet.name + "'!$AF$" + str(last_row)]
                summary_formula.extend(collect)
                collect = []

        start_row = 19
        count = 1
        offset = 20
        odered_summary_formula = summary_formula[::-1]
        sheet = wb.sheets['Summary']
        # Clear summary page first
        sheet.range('A18:Z1000').clear()
        # Set format
        sheet.range('C:C').column_width = 55
        # sheet.range('E20:E1000').horizontal_alignment = 'center'
 
        for system in wb.sheet_names:
            if system not in skip_sheets:
                (pwb.sheets['Design'].range('21:21')).copy(sheet.range(str(offset) + ':' + str(offset)))  # noqa: E501
                sheet.range('B' + str(offset)).value = count
                sheet.range('B' + str(offset)).value = count
                sheet.range('C' + str(offset)).value = system
                sheet.range('D' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('D' + str(offset)).number_format = n_format
                sheet.range('H' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('H' + str(offset)).number_format = n_format
                sheet.range('H' + str(offset)).font.color = (4, 50, 255)
                sheet.range('I' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('I' + str(offset)).number_format = n_format
                sheet.range('I' + str(offset)).font.color = (148, 55, 255)
                sheet.range('J' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('J' + str(offset)).number_format = n_format
                sheet.range('J' + str(offset)).font.color = (148, 55, 255)
                sheet.range('K' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('K' + str(offset)).number_format = n_format
                sheet.range('K' + str(offset)).font.color = (148, 55, 255)
                sheet.range('L' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('L' + str(offset)).number_format = n_format
                sheet.range('L' + str(offset)).font.color = (148, 55, 255)
                sheet.range('M' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('M' + str(offset)).number_format = n_format
                sheet.range('M' + str(offset)).font.color = (148, 55, 255)
                sheet.range('N' + str(offset)).formula = odered_summary_formula.pop()
                sheet.range('N' + str(offset)).number_format = n_format
                sheet.range('O' + str(offset)).formula = '=IF(N'+ str(offset) + '<>"",D' + str(offset) + '- N' + str(offset) + ',"")'  # noqa: E501
                sheet.range('O' + str(offset)).number_format = n_format
                sheet.range('P' + str(offset)).formula = '=IF(O' + str(offset) + '<>"",O' + str(offset) + '/D' + str(offset) + ',"")'  # noqa: E501
                sheet.range('P' + str(offset)).number_format = '0.00%'
                count += 1
                offset += 1
        
        # Drawing lines
        (pwb.sheets['Design'].range('15:15')).copy(sheet.range(str(start_row) + ':' + str(start_row)))  # noqa: E501
        (pwb.sheets['Design'].range('11:11')).copy(sheet.range(str(offset) + ':' + str(offset)))  # noqa: E501
        (pwb.sheets['Design'].range('17:17')).copy(sheet.range(str(offset+1) + ':' + str(offset+1)))  # noqa: E501
        # sheet = wb.sheets['Summary']
        sheet.range('C' + str(offset+1)).value = '="TOTAL PROJECT (" & Config!B12 & ")"'
        sheet.range('D' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",D20:D' + str(offset) + ')'  # noqa: E501
        sheet.range('D' + str(offset+1)).number_format = n_format
        sheet.range('E' + str(offset+1)).formula = '=IF(COUNTIF(E20:E' + str(offset) + ',"OPTION"), "Excluding Option", "")'  # noqa: E501
        sheet.range('H' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",H20:H' + str(offset) + ')'  # noqa: E501
        sheet.range('H' + str(offset+1)).number_format = n_format
        sheet.range('H' + str(offset+1)).font.color = (4, 50, 255)
        sheet.range('I' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",I20:I' + str(offset) + ')'  # noqa: E501
        sheet.range('I' + str(offset+1)).number_format = n_format
        sheet.range('I' + str(offset+1)).font.color = (148, 55, 255)
        sheet.range('J' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",J20:J' + str(offset) + ')'  # noqa: E501
        sheet.range('J' + str(offset+1)).number_format = n_format
        sheet.range('J' + str(offset+1)).font.color = (148, 55, 255)
        sheet.range('K' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",K20:K' + str(offset) + ')'  # noqa: E501
        sheet.range('K' + str(offset+1)).number_format = n_format
        sheet.range('K' + str(offset+1)).font.color = (148, 55, 255)
        sheet.range('L' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",L20:L' + str(offset) + ')'  # noqa: E501
        sheet.range('L' + str(offset+1)).number_format = n_format
        sheet.range('L' + str(offset+1)).font.color = (148, 55, 255)
        sheet.range('M' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",M20:M' + str(offset) + ')'  # noqa: E501
        sheet.range('M' + str(offset+1)).number_format = n_format
        sheet.range('M' + str(offset+1)).font.color = (148, 55, 255)
        sheet.range('N' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",N20:N' + str(offset) + ')'  # noqa: E501
        sheet.range('N' + str(offset+1)).number_format = n_format
        sheet.range('O' + str(offset+1)).formula = '=IF(N'+ str(offset+1) + '<>"", D' + str(offset+1) + '- N' + str(offset+1) + ',"")'  # noqa: E501
        sheet.range('O' + str(offset+1)).number_format = n_format
        sheet.range('P' + str(offset+1)).formula = '=IF(O' + str(offset+1) + '<>0,O' + str(offset+1) + '/D' + str(offset+1) + ',"")'  # noqa: E501
        sheet.range('P' + str(offset+1)).number_format = '0.00%'
        if discount:
            (pwb.sheets['Design'].range('18:18')).copy(sheet.range(str(offset+2) + ':' + str(offset+2)))  # noqa: E501
            (pwb.sheets['Design'].range('19:19')).copy(sheet.range(str(offset+3) + ':' + str(offset+3)))  # noqa: E501
            sheet.range('C' + str(offset+3)).formula = '="TOTAL PROJECT PRICE AFTER DISCOUNT (" & Config!B12 & ")"'  # noqa: E501
            sheet.range('D' + str(offset+3)).formula = '=SUM(D' +str(offset+1) + ':D' + str(offset+2) + ')'  # noqa: E501
            sheet.range('D' + str(offset+3)).number_format = n_format
            sheet.range('N' + str(offset+3)).formula = '=$N$' +str(offset+1)
            sheet.range('N' + str(offset+3)).number_format = n_format
            sheet.range('O' + str(offset+3)).formula = '=IF(N'+ str(offset+3) + '<>"", D' + str(offset+3) + '- N' + str(offset+3) + ',"")'  # noqa: E501
            sheet.range('O' + str(offset+3)).number_format = n_format
            sheet.range('P' + str(offset+3)).formula = '=IF(O' + str(offset+3) + '<>0,O' + str(offset+3) + '/D' + str(offset+3) + ',"")'  # noqa: E501
            sheet.range('P' + str(offset+3)).number_format = '0.00%'
            sheet.range('C' + str(offset+5)).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'  # noqa: E501
            sheet.range('C' + str(offset+6)).value = "• Total project price does not include prices for optional items set out in the detailed bill of material."  # noqa: E501
            sheet.range('C' + str(offset+7)).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."  # noqa: E501
        else:
            sheet.range('C' + str(offset+3)).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'  # noqa: E501
            sheet.range('C' + str(offset+4)).value = "• Total project price does not include items marked 'OPTION' in the detailed bill of material."  # noqa: E501
            sheet.range('C' + str(offset+5)).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."  # noqa: E501

    sheet.range('D:D').autofit()
    sheet.range('E:E').column_width = 15
    sheet.range('F:P').autofit()
    last_row = sheet.range('C100000').end('up').row
    sheet.page_setup.print_area = 'A1:F' + str(last_row+3)

def number_title(wb, count=10, step=10):
    """
    For the main numbering. It will fix as long as it is a number.
    Need to look for only the systems and engineering services.
    Takes a work book, then start number and step."""
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            system_names.append(str.upper(sheet.name))
            ws = wb.sheets[sheet]
            last_row = ws.range('C100000').end('up').row
            data = ws.range('A2:C' + str(last_row)).options(pd.DataFrame, index=False).value  # noqa: E501
            data['System'] = str.upper(sheet.name)
            systems = pd.concat([systems, data], join='outer')
    # Now that I have collect the data, let us do the numbering
    # Index is reset so that index number is continuous
    systems = systems.reset_index(drop=True)
    # Reindexing will remove columns that are not named.
    systems = systems.reindex(columns=['NO', 'Description', 'System'])

    # Need to do try-except as the float type can return nan
    for idx, item in systems['NO'].items():
        try:
            if int(item):
                systems.at[idx, 'NO'] = count
                count += step
        except Exception:
            pass
    
    # Now is the matter of writing to the required sheets
    for system in system_names:
        # print(system)
        sheet = wb.sheets[system]
        system = systems[systems['System'] == system]
        sheet.range('A2').options(index=False).value = system['NO']

def technical(wb):
    directory = os.path.dirname(wb.fullname)

    wb.sheets['Cover'].range('D39').value = 'TECHNICAL PROPOSAL'
    wb.sheets['Cover'].range('C42:C47').value = wb.sheets['Cover'].range('C42:C47').raw_value  # noqa: E501
    wb.sheets['Cover'].range('D6:D8').value = wb.sheets['Cover'].range('D6:D8').raw_value  # noqa: E501
    wb.sheets['Summary'].range('D20:D100').value = ''
    wb.sheets['Summary'].range('C20:C100').value = wb.sheets['Summary'].range('C20:C100').raw_value  # noqa: E501
    wb.sheets['Summary'].range('G:P').delete()
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    for sheet in wb.sheet_names:
        ws = wb.sheets[sheet]
        ws.range('A1').value = ws.range('A1').raw_value #Remove formula
        if sheet not in skip_sheets:
            last_row = ws.range('B100000').end('up').row
            ws.range('B3:B' + str(last_row)).value = ws.range('B3:B' + str(last_row)).raw_value  # noqa: E501
            ws.range('AM:AN').delete()
            ws.range('I:AK').delete()
            ws.range('F:G').delete()
            # To reduce visual clutter
            ws.range('G:H').column_width = 0
    wb.sheets['Config'].delete()
    wb.sheets['T&C'].delete()
    prepare_to_print_technical(wb)
    wb.sheets['Summary'].activate()
    file_name = 'Technical ' + wb.name[:-4] + 'xlsx'
    wb.save(Path(directory, file_name), password='')
    technical_wb = xw.Book(file_name)
    print_technical(technical_wb)

def prepare_to_print_commercial(wb):
    """Takes a work book, set horizantal borders at pagebreaks."""
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    macro_nb = xw.Book('PERSONAL.XLSB')
    current_sheet = wb.sheets.active
    for sheet in wb.sheet_names:
        if sheet not in skip_sheets:
            wb.sheets[sheet].activate()
            # Adjust column width as sometimes, the long value does not show.
            wb.sheets[sheet].range('A:A').column_width = 4
            wb.sheets[sheet].range('B:B').autofit()
            wb.sheets[sheet].range('C:C').autofit()
            wb.sheets[sheet].range('C:C').column_width = 55
            # wb.sheets[sheet].range('C:C').wrap_text = True
            wb.sheets[sheet].range('D:H').autofit()
            # Call macros
            macro_nb.macro('conditional_format')()
            macro_nb.macro('remove_h_borders')()
            macro_nb.macro('pagebreak_borders')()
    wb.sheets[current_sheet].activate()

def prepare_to_print_technical(wb):
    """Takes a work book, set horizantal borders at pagebreaks."""
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    macro_nb = xw.Book('PERSONAL.XLSB')
    current_sheet = wb.sheets.active
    for sheet in wb.sheet_names:
        if sheet not in skip_sheets:
            wb.sheets[sheet].activate()
            macro_nb.macro('conditional_format_technical')()
            macro_nb.macro('remove_h_borders')()
            macro_nb.macro('pagebreak_borders')()
    wb.sheets[current_sheet].activate()

def print_commercial(wb):
    """The commercial proposal will be written to the cwd."""
    prepare_to_print_commercial(wb)
    try:
        wb.to_pdf(exclude='Config', show=True)
    except Exception:
            # The program does not override the existing file. Therefore, the file needs to be removed if it exists.  # noqa: E501
            xw.apps.active.alert('The PDF file already exists!\n Please delete the file and try again.')  # noqa: E501

def print_technical(wb):
    """The technical proposal will be written to the cwd."""
    try:
        wb.to_pdf(show=True)
    except Exception:
            # The program does not override the existing file. The file needs to be removed if it exists.  # noqa: E501
            xw.apps.active.alert('The PDF file already exists!\n Please delete the file and try again.')  # noqa: E501

def conditional_format_wb(wb):  
    """
    Takes a workbook, and do conditional formatting.
    Rely on excel macro for conditional format.
    """
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    macro_nb = xw.Book('PERSONAL.XLSB')
    current_sheet = wb.sheets.active
    for sheet in wb.sheet_names:
        if sheet not in skip_sheets:
            wb.sheets[sheet].activate()
            macro_nb.macro('conditional_format')()
    wb.sheets[current_sheet].activate()

def fix_unit_price(wb):
    """ 
    Fix unit prices, normally done for subsequent revisions.
    """
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            system_names.append(str.upper(sheet.name))
            ws = wb.sheets[sheet]
            last_row = ws.range('C100000').end('up').row
            data = ws.range('AE2:AE' + str(last_row)).options(pd.DataFrame, index=False).value  # noqa: E501
            data['System'] = str.upper(sheet.name)
            systems = pd.concat([systems, data], join='outer')
    
    systems = systems.reset_index(drop=True) # Otherwise separate sheet will have own index.  # noqa: E501
    systems.columns = ['FUP', 'System']

    # Write fixed unit price in FUP field
    for system in system_names:
        sheet = wb.sheets[system]
        system = systems[systems['System'] == system]
        sheet.range('AB2').options(index=False).value = system['FUP']

def format_text(wb, indent_description=False, bullet_description=False, title_lineitem=False):  # noqa: E501
    """ 
    Format text in the workbook to remove inconsistencies.
    """
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            system_names.append(str.upper(sheet.name))
            ws = wb.sheets[sheet]
            last_row = ws.range('C100000').end('up').row
            data = ws.range('C2:AL' + str(last_row)).options(pd.DataFrame, empty='', index=False).value  # noqa: E501
            data['System'] = str.upper(sheet.name)
            # format_type = ws.range('AL2:AL' + str(last_row)).options(pd.DataFrame, empty='', index=False).value  # noqa: E501
            systems = pd.concat([systems, data], join='outer')
            # systems = pd.concat([systems, format_type], join='outer')
    
    systems = systems.reset_index(drop=True) # Otherwise separate sheet will have own index.  # noqa: E501
    systems = systems.reindex(columns=['Description', 'Format', 'System'])

    for idx, item in systems['Description'].items():
        systems.at[idx, 'Description'] = set_nitty_gritty(str(systems.loc[idx, 'Description']).strip().lstrip('• '))  # noqa: E501
        if indent_description:
            if systems.at[idx, 'Format'] == 'Description':
                systems.at[idx, 'Description'] = '   ' + (str(systems.loc[idx, 'Description']).strip()).lstrip('• ')  # noqa: E501
                if bullet_description:
                    systems.at[idx, 'Description'] = '   • ' + str(systems.loc[idx, 'Description']).strip()  # noqa: E501
            
        if title_lineitem:
            if systems.at[idx, 'Format'] == 'Lineitem':
                systems.at[idx, 'Description'] = set_case_preserve_acronym(
                    (str(systems.loc[idx, 'Description']).strip()).lstrip('• '), title=True)  # noqa: E501

    # Write fomatted description to Description field
    for system in system_names:
        sheet = wb.sheets[system]
        system = systems[systems['System'] == system]
        # sheet.range('C2').value = sheet.range('C2').options(empty='')
        sheet.range('C2').options(index=False).value = system['Description']

def indent_description(wb):
    """ 
    Indent description
    This function works but slow. Replaced with 'format_text' function
    """
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    for sheet in wb.sheets:
        if sheet.name not in skip_sheets:
            ws = wb.sheets[sheet]
            last_row = ws.range('C100000').end('up').row
            for format in ws.range('AL3:AL' + str(last_row)):
                if format.value == 'Subtitle':
                    ws.range('C'+ str(format.row)).value = str(ws.range('C'+ str(format.row)).value).strip()  # noqa: E501
                    ws.range('C'+ str(format.row)).value = str(ws.range('C'+ str(format.row)).value).lstrip('• ')  # noqa: E501
                elif format.value == 'Description':
                    ws.range('C'+ str(format.row)).value = str(ws.range('C'+ str(format.row)).value).strip()  # noqa: E501
                    ws.range('C'+ str(format.row)).value = str(ws.range('C'+ str(format.row)).value).lstrip('• ')  # noqa: E501
                    ws.range('C'+ str(format.row)).value = '   • ' + ws.range('C'+ str(format.row)).value  # noqa: E501