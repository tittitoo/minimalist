""" Created so that python fucntions are available in Excel.
    © Thiha Aung
"""

import xlwings as xw
import functions


def fill_formula():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_formula(ws)

def fill_formula_wb():
    wb = xw.Book.caller()
    functions.fill_formula_wb(wb)

def subtotal():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_lastrow_sheet(wb, ws)

def subtotal_wb():
    wb = xw.Book.caller()
    functions.fill_lastrow(wb)

def unhide_columns():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.unhide_columns(ws)

def summary():
    wb = xw.Book.caller()
    functions.summary(wb, discount=False)

def summary_discount():
    wb = xw.Book.caller()
    functions.summary(wb, discount=True)

def summary_detail():
    wb = xw.Book.caller()
    functions.summary(wb, discount=False, detail=True)

def summary_detail_discount():
    wb = xw.Book.caller()
    functions.summary(wb, discount=True, detail=True)

def number_title():
    wb = xw.Book.caller()
    functions.number_title(wb)

def hide_columns():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.hide_columns(ws)

def technical():
    wb = xw.Book.caller()
    functions.technical(wb)

def prepare_to_print_commercial():
    wb = xw.Book.caller()
    functions.prepare_for_print_commercial(wb)

def print_commercial():
    wb = xw.Book.caller()
    functions.print_commercial(wb)

def conditional_format_wb():
    wb = xw.Book.caller()
    functions.conditional_format_wb(wb)

def fix_unit_price():
    wb = xw.Book.caller()
    functions.fix_unit_price(wb)

def format_text():
    wb = xw.Book.caller()
    functions.format_text(wb, title_lineitem_or_description=True)

def indent_description():
    wb = xw.Book.caller()
    functions.format_text(wb, indent_description=True, bullet_description=True)