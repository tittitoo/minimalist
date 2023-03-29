""" Multiple functions to support Excel automation.
    © Thiha Aung
"""

import re
import os
from pathlib import Path
import pandas as pd
import xlwings as xw
import hide

# Fix annoying text
def set_nitty_gritty(text):
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

# Set space after comma
def set_comma_space(text):
    x = re.compile(',\w+')
    if x.search(text):
        substring = re.findall(',\w+', text)
        for word in substring:
            text = re.sub(word, ', ' + word[1:], text)
    return text

# Function to replace description such as 1x, 20x, 10X , x1, x20, X20 into 1 x, 20 x, 10 x, x 1, x 20, X 10 etc.
def set_x(text):
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

# Take in sheet and workbook
def fill_formula(sheet):
    # Formula to cells
    last_row = sheet.range('C100000').end('up').row
    sheet.range('A1').formula = '= "JASON REF: " & Config!B29 &  ", REVISION: " &  Config!B30 & ", PROJECT: " & Config!B26'
    # Serail Numbering (SN)
    sheet.range('B3:B' + str(last_row)).formula = '=IF(AND(ISNUMBER(D3), ISNUMBER(K3), XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)), COUNT(INDIRECT(CONCAT("B",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1),":B",ROW()-1))) + 1, "")'
    sheet.range('N3:N' + str(last_row)).formula = '=IF(K3<>"",K3*(1-M3),"")'
    sheet.range('O3:O' + str(last_row)).formula = '=IF(AND(D3<>"", K3<>"",H3<>"OPTION"),D3*N3,"")'
    # Exchange rates
    sheet.range('Q3:Q' + str(last_row)).formula = '=IF(Config!B12="SGD",IF(J3<>"",VLOOKUP(J3,Config!$A$2:$B$10,2,FALSE),""),IF(J3<>"",VLOOKUP(J3,Config!$A$2:$B$10,2,FALSE)/VLOOKUP(Config!$B$12,Config!$A$2:$B$10,2,FALSE),""))'
    sheet.range('R3:R' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"") ,N3*Q3,"")'
    sheet.range('S3:S' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION") ,D3*R3,"")'
    sheet.range('T3:T' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""), (R3*(1+$L$1+$N$1+$P$1+$R$1))/(1-0.05),"")'
    sheet.range('U3:U' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION",INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))<>"OPTION"),D3*T3,"")'
    sheet.range('V3:V' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),R3*$L$1,"")'
    sheet.range('W3:W' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),R3*$N$1,"")'
    sheet.range('X3:X' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),R3*$P$1,"")'
    sheet.range('Y3:Y' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),R3*$R$1,"")'
    sheet.range('Z3:Z' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),T3-(R3+V3+W3+X3+Y3),"")'
    sheet.range('AA3:AA' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),$J$1,"")'
    sheet.range('AC3:AC' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),CEILING(T3/(1-AA3), 1),"")'
    # sheet.range('AD3:AD' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION",H3<>"INCLUDED"),D3*AC3,"")'
    sheet.range('AD3:AD' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION",H3<>"INCLUDED",(INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))) <>"OPTION"),D3*AC3,"")'
    sheet.range('AE3:AE' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),IF(AB3<>"",AB3,AC3),"")'
    # sheet.range('AF3:AF' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED"),D3*AE3,"")'
    sheet.range('AF3:AF' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED",(INDIRECT(CONCAT("H",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))) <>"OPTION"),D3*AE3,"")'
    # sheet.range('AF3:AF' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>""),D3*AE3,"")'
    sheet.range('AG3:AG' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION", H3<>"INCLUDED",AF3<>""),AF3-U3,"")'
    sheet.range('AH3:AH' + str(last_row)).formula = '=IF(AND(AG3<>"",AG3<>0),AG3/AF3,"")'
    sheet.range('AI3:AI' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"", H3<>"OPTION"),D3*AE3,"")'
    # sheet.range('AL3:AL' + str(last_row)).formula = '=IF(A3<>"","Title",IF(B3<>"","Lineitem",IF(LEFT(C3,3)="***","Comment",IF(AND(A3="",B3="",C2="", C4<>"",D4<>""), "Subtitle",""))))'
    # Unit Price
    sheet.range('F3:F' + str(last_row)).formula = '=IF(AND(AL3="Title", ISNUMBER(AJ3)), AJ3, IF(AND(AL3="Lineitem", AK3="Lumpsum", H3<>"OPTION"), "", AE3))'
    # sheet.range('F3:F' + str(last_row)).formula = '=IF(AE3<>"", AE3,"")'
    sheet.range('G3:G' + str(last_row)).formula = '=IF(AND(F3<>"", H3<>"OPTION", H3<>"INCLUDED"), D3*F3,"")'
    sheet.range('L3:L' + str(last_row)).formula = '=IF(AND(D3<>"",K3<>"",H3<>"OPTION"),D3*K3,"")'
    # For Format field
    sheet.range('AL1').value = "Title"
    sheet.range('AL3:AL' + str(last_row)).formula = '=IF(A3<>"","Title", IF(B3<>"","Lineitem", IF(LEFT(C3,3)="***","Comment", IF(AND(A3="",B3="",C2="", C4<>"",D4<>""), "Subtitle",""))))'
    sheet.range('AL' + str(last_row+1)).value = "Title"

    # For Lumpsum
    # sheet.range('AJ3:AJ' + str(last_row)).formula = '=IF(AND(AL3="Title", D3=1, E3="lot"), SUM(INDIRECT(CONCAT("AF", ROW()+1, ":AF",((MATCH("Title",INDIRECT(CONCAT("AL", ROW()+1, ":AL", MATCH(REPT("z",50),AL:AL))),0)) + ROW())))), "")'
    sheet.range('AJ3:AJ' + str(last_row)).formula = '=IF(AND(AL3="Title", D3=1, E3="lot"), SUM(INDIRECT(CONCAT("AI", ROW()+1, ":AI",((MATCH("Title",INDIRECT(CONCAT("AL", ROW()+1, ":AL", MATCH(REPT("z",50),AL:AL))),0)) + ROW())))), "")'
    sheet.range('AK3:AK' + str(last_row)).formula = '=IF(AL3="Lineitem", IF(ISNUMBER(INDIRECT(CONCAT("AJ",XMATCH("Title",(INDIRECT(CONCAT("AL1:","AL",ROW()-1))),0,-1)))), "Lumpsum", "Unit Price"), "")'

def fill_lastrow (wb):
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']

    for sheet in wb.sheet_names:
       if sheet not in skip_sheets:
            sheet = wb.sheets[sheet]
            last_row = sheet.range('G100000').end('up').row
            (wb.sheets['Config'].range('100:100')).copy(sheet.range(str(last_row+2) + ':' + str(last_row+2)))
            sheet.range('F'+ str(last_row+2)).formula = '="Subtotal(" & Config!B12 & ")"'
            sheet.range('F'+ str(last_row+2)).font.bold = True
            sheet.range('F'+ str(last_row+2)).font.size = 9
            sheet.range('G' + str(last_row+2)).formula = '=SUM(G3:G' + str(last_row+1) + ')'
            sheet.range('G' + str(last_row+2)).font.bold = True
            sheet.range('U' + str(last_row+2)).formula = '=SUM(U3:U' + str(last_row+1) + ')'
            sheet.range('U' + str(last_row+2)).font.bold = True
            sheet.range('AF' + str(last_row+2)).formula = '=SUM(AF3:AF' + str(last_row+1) + ')'
            # sheet.range('AF' + str(last_row+2)).formula = '=SUMIF(H3:H' + str(last_row+1) + ',"<>OPTION",AF3:AF' + str(last_row+1) + ')'
            sheet.range('AF' + str(last_row+2)).font.bold = True
            sheet.range('AG' + str(last_row+2)).formula = '=SUM(AG3:AG' + str(last_row+1) + ')'
            sheet.range('AG' + str(last_row+2)).font.bold = True
            sheet.range('AH' + str(last_row+2)).formula = '=AG' + str(last_row+2) + '/AF' + str(last_row+2)
            sheet.range('AH' + str(last_row+2)).font.bold = True

            # Set-up print area
            sheet.page_setup.print_area = 'A1:H' + str(last_row+2)

def fill_lastrow_sheet(wb, sheet):
    last_row = sheet.range('C100000').end('up').row
    (wb.sheets['Config'].range('100:100')).copy(sheet.range(str(last_row+2) + ':' + str(last_row+2)))
    sheet.range('F'+ str(last_row+2)).formula = '="Subtotal(" & Config!B12 & ")"'
    sheet.range('F'+ str(last_row+2)).font.bold = True
    sheet.range('F'+ str(last_row+2)).font.size = 9
    sheet.range('G' + str(last_row+2)).formula = '=SUM(G3:G' + str(last_row+1) + ')'
    sheet.range('G' + str(last_row+2)).font.bold = True
    sheet.range('U' + str(last_row+2)).formula = '=SUM(U3:U' + str(last_row+1) + ')'
    sheet.range('U' + str(last_row+2)).font.bold = True
    sheet.range('AF' + str(last_row+2)).formula = '=SUM(AF3:AF' + str(last_row+1) + ')'
    # sheet.range('AF' + str(last_row+2)).formula = '=SUMIF(H3:H' + str(last_row+1) + ',"<>OPTION",AF3:AF' + str(last_row+1) + ')'
    sheet.range('AF' + str(last_row+2)).font.bold = True
    sheet.range('AG' + str(last_row+2)).formula = '=SUM(AG3:AG' + str(last_row+1) + ')'
    sheet.range('AG' + str(last_row+2)).font.bold = True
    sheet.range('AH' + str(last_row+2)).formula = '=AG' + str(last_row+2) + '/AF' + str(last_row+2)
    sheet.range('AH' + str(last_row+2)).font.bold = True

    # Set-up print area
    sheet.page_setup.print_area = 'A1:H' + str(last_row+2)


    # For formatting
def format(sheet):
    sheet.range('A:A').column_width = 3
    sheet.range('B:B').autofit()
    sheet.range('C:C').autofit()
    sheet.range('C:C').column_width = 55
    sheet.range('C:C').wrap_text = True
    sheet.range('D:D').autofit()
    sheet.range('E:E').autofit()
    sheet.range('F:F').autofit()
    sheet.range('G:G').autofit()
    sheet.range('H:H').autofit()
    sheet.range('J:J').autofit()
    sheet.range('K:K').autofit()
    sheet.range('L:L').autofit()
    sheet.range('M:M').autofit()
    sheet.range('N:N').autofit()
    sheet.range('O:O').autofit()
    sheet.range('Q:Q').autofit()
    sheet.range('R:R').autofit()
    sheet.range('S:S').autofit()
    sheet.range('T:T').autofit()
    sheet.range('U:U').autofit()
    sheet.range('V:V').autofit()
    sheet.range('W:W').autofit()
    sheet.range('X:X').autofit()
    sheet.range('Y:Y').autofit()
    sheet.range('Z:Z').autofit()
    sheet.range('AA:AA').autofit()
    sheet.range('AB:AB').autofit()
    sheet.range('AC:AC').autofit()
    sheet.range('AD:AD').autofit()
    sheet.range('AE:AE').autofit()
    sheet.range('AF:AF').autofit()
    sheet.range('AG:AG').autofit()
    sheet.range('AH:AH').autofit()
    sheet.range('AI:AI').autofit()
    sheet.range('AJ:AJ').autofit()
    sheet.range('AK:AK').autofit()
    sheet.range('AL:AL').autofit()
    sheet.range('AM:AM').autofit()
    sheet.range('AN:AN').autofit()

def hide_columns(sheet):
    sheet.range('AI:AN').column_width = 0
    sheet.range('AC:AF').column_width = 0
    sheet.range('AB:AB').column_width = 10
    sheet.range('S:AA').column_width = 0
    sheet.range('Q:Q').column_width = 0
    sheet.range('P:P').column_width = 20
    sheet.range('O:O').column_width = 0
    sheet.range('L:L').column_width = 0
    # sheet.range('F:G').column_width = 0
    # sheet.range('B:B').column_width = 0

def summary(wb, discount=True):
    summary_formula = []
    collect = [] # Collect formula to be put in summary page.
    formula_fragment = '=IF(OR(Config!B13="COMMERCIAL PROPOSAL", Config!B13="BUDGETARY PROPOSAL"),'
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']

    for sheet in wb.sheet_names:
       if sheet not in skip_sheets:
            sheet = wb.sheets[sheet]
            last_row = sheet.range('G100000').end('up').row
            collect = [ formula_fragment + "'" + sheet.name + "'!$G$" + str(last_row) + ', "")',
                        formula_fragment + "'" + sheet.name + "'!$U$" + str(last_row) + ', "")']
                #    "='" + sheet.name + "'!$AF$" + str(last_row)]
            summary_formula.extend(collect)
            collect = []

    count = 1
    offset = 20
    odered_summary_formula = summary_formula[::-1]
    sheet = wb.sheets['Summary']
    for system in wb.sheet_names:
        if system not in skip_sheets:
            sheet.range('B' + str(offset)).value = count
            sheet.range('C' + str(offset)).value = system
            sheet.range('D' + str(offset)).formula = odered_summary_formula.pop()
            sheet.range('H' + str(offset)).formula = odered_summary_formula.pop()
            sheet.range('I' + str(offset)).formula = '=IF(H'+ str(offset) + '<>"",D' + str(offset) + '- H' + str(offset) + ',"")'
            sheet.range('J' + str(offset)).formula = '=IF(I' + str(offset) + '<>"",I' + str(offset) + '/D' + str(offset) + ',"")'
            count += 1
            offset += 1
    
    # Drawing lines
    (wb.sheets['Config'].range('106:106')).copy(sheet.range(str(offset) + ':' + str(offset)))
    (wb.sheets['Config'].range('102:102')).copy(sheet.range(str(offset+1) + ':' + str(offset+1)))
    # sheet = wb.sheets['Summary']
    sheet.range('C' + str(offset+1)).value = '="TOTAL PROJECT (" & Config!B12 & ")"'
    sheet.range('D' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",D20:D' + str(offset) + ')'
    sheet.range('E' + str(offset+1)).formula = '=IF(COUNTIF(E20:E' + str(offset) + ',"OPTION"), "Excluding Option", "")'
    sheet.range('H' + str(offset+1)).formula = '=SUMIF(E20:E' + str(offset) + ',"<>OPTION",H20:H' + str(offset) + ')'
    sheet.range('I' + str(offset+1)).formula = '=IF(H'+ str(offset+1) + '<>"", D' + str(offset+1) + '- H' + str(offset+1) + ',"")'
    sheet.range('J' + str(offset+1)).formula = '=IF(I' + str(offset+1) + '<>0,I' + str(offset+1) + '/D' + str(offset+1) + ',"")'
    if discount:
        (wb.sheets['Config'].range('103:103')).copy(sheet.range(str(offset+2) + ':' + str(offset+2)))
        (wb.sheets['Config'].range('104:104')).copy(sheet.range(str(offset+3) + ':' + str(offset+3)))
        sheet.range('C' + str(offset+3)).formula = '="TOTAL PROJECT PRICE AFTER DISCOUNT (" & Config!B12 & ")"'
        sheet.range('D' + str(offset+3)).formula = '=SUM(D' +str(offset+1) + ':D' + str(offset+2) + ')'
        sheet.range('H' + str(offset+3)).formula = '=$H$' +str(offset+1)
        sheet.range('I' + str(offset+3)).formula = '=IF(H'+ str(offset+3) + '<>"", D' + str(offset+3) + '- H' + str(offset+3) + ',"")'
        sheet.range('J' + str(offset+3)).formula = '=IF(I' + str(offset+3) + '<>0,I' + str(offset+3) + '/D' + str(offset+3) + ',"")'
        sheet.range('C' + str(offset+5)).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'
        sheet.range('C' + str(offset+6)).value = "• Total project price does not include prices for optional items set out in the detailed bill of material."
        sheet.range('C' + str(offset+7)).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."
    else:
        sheet.range('C' + str(offset+3)).formula = '="• All the prices are in " & Config!B12 & " excluding GST."'
        sheet.range('C' + str(offset+4)).value = "• Total project price does not include items marked 'OPTION' in the detailed bill of material."
        sheet.range('C' + str(offset+5)).value = "• Items marked as 'INCLUDED' are included in the scope of supply without price impact."

    last_row = sheet.range('C100000').end('up').row
    sheet.page_setup.print_area = 'A1:F' + str(last_row+3)

# For the main numbering. It will fix as long as it is a number.
# Need to look for only the systems and engineering services.
def number_title(wb, count=10, step=10):
    """Takes a work book, then start number and step."""
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    # Collect system_names and data
    systems = pd.DataFrame()
    system_names = []
    for sheet in wb.sheet_names:
        if sheet not in skip_sheets:
            system_names.append(sheet.upper())
            ws = wb.sheets[sheet]
            last_row = ws.range('C100000').end('up').row
            data = ws.range('A2:C' + str(last_row)).options(pd.DataFrame, index=False).value
            data['System'] = str(sheet.upper())
            systems = pd.concat([systems, data], join='outer')
    # Now that I have collect the data, let us do the numbering
    systems = systems.reset_index(drop=True)
    systems = systems.reindex(columns=['NO', 'Description', 'System'])

    # Need to do try-except as the float type can return nan
    for idx, item in systems['NO'].items():
        try:
            if int(item):
                systems.at[idx, 'NO'] = count
                count += step
        except Exception as e:
            pass
    
    # Now is the matter of writing to the required sheets
    for system in system_names:
        # print(system)
        sheet = wb.sheets[system]
        system = systems[systems['System'] == system]
        sheet.range('A2').options(index=False).value = system['NO']

# For formatting text for consistency
# Need to look for only the systems and engineering services.
def format_description(ws):
    """Takes a work book or sheet, and format text."""
    systems = pd.DataFrame()
    system_names = []
    if isinstance(ws, xw.main.Sheet):
        last_row = ws.range('C100000').end('up').row
        data = ws.range('A2:C' + str(last_row)).options(pd.DataFrame, index=False).value

    # skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    # # Collect system_names and data
    # for sheet in wb.sheet_names:
    #     if sheet not in skip_sheets:
    #         system_names.append(sheet.upper())
    #         ws = wb.sheets[sheet]
    #         last_row = ws.range('C100000').end('up').row
    #         data = ws.range('A2:C' + str(last_row)).options(pd.DataFrame, index=False).value
    #         data['System'] = str(sheet.upper())
    #         systems = pd.concat([systems, data], join='outer')
    # # Now that I have collect the data, let us do the numbering
    # systems = systems.reset_index(drop=True)
    # systems = systems.reindex(columns=['NO', 'Description', 'System'])

    # # Need to do try-except as the float type can return nan
    # for idx, item in systems['NO'].items():
    #     try:
    #         if int(item):
    #             systems.at[idx, 'NO'] = count
    #             count += step
    #     except Exception as e:
    #         pass
    
    # # Now is the matter of writing to the required sheets
    # for system in system_names:
    #     # print(system)
    #     sheet = wb.sheets[system]
    #     system = systems[systems['System'] == system]
    #     sheet.range('A2').options(index=False).value = system['NO']

def technical(wb):
    file = wb.name
    downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
    file_path = Path(downloads_folder, file)
    try:
        nb = xw.Book(file_path, password=hide.legacy)
    except:
        nb = xw.Book(file_path, password=hide.new)

    nb.sheets['Cover'].range('D39').value = 'TECHNICAL PROPOSAL'
    nb.sheets['Cover'].range('C42:C47').value = nb.sheets['Cover'].range('C42:C47').raw_value
    nb.sheets['Cover'].range('D6:D8').value = nb.sheets['Cover'].range('D6:D8').raw_value
    nb.sheets['Summary'].range('D20:D100').value = ''
    nb.sheets['Summary'].range('C20:C100').value = nb.sheets['Summary'].range('C20:C100').raw_value
    nb.sheets['Summary'].range('G:K').delete()
    skip_sheets = ['Config', 'Cover', 'Summary', 'Technical_Notes', 'T&C']
    for sheet in nb.sheet_names:
        ws = nb.sheets[sheet]
        ws.range('A1').value = ws.range('A1').raw_value
        if sheet not in skip_sheets:
            last_row = ws.range('B100000').end('up').row
            ws.range('B3:B' + str(last_row)).value = ws.range('B3:B' + str(last_row)).raw_value
            ws.range('AM:AN').delete()
            ws.range('I:AK').delete()
            ws.range('F:G').delete()
    nb.sheets['Config'].delete()
    nb.sheets['T&C'].delete()
    nb.sheets['Summary'].activate()
    nb.save( downloads_folder + '/' + 'Technical ' + file[:-4] + 'xlsx', password='') 
