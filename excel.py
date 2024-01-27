""" Created so that python fucntions are available in Excel.
    © Thiha Aung
"""

import xlwings as xw  # type: ignore
import functions
import checklists
import checklist_collections as cc

# from reportlab.lib.colors import lightcyan, black, white, lightyellow, blue


def fill_formula():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_formula(ws)


# Fix the whole workbook. The function name will later change to fix_workbook
# Now is tied ot Fix Wrokbook in Excel
def fill_formula_wb():
    wb = xw.Book.caller()
    functions.delete_extra_empty_row_wb(wb)
    # Calling twice as sometimes some rows are missed.
    functions.delete_extra_empty_row_wb(wb)
    functions.number_title(wb)
    functions.fill_formula_wb(wb)
    functions.format_text(wb, title_lineitem_or_description=True)
    functions.format_text(wb, indent_description=True, bullet_description=True)
    functions.format_cell_data(wb)
    functions.adjust_columns_wb(wb)
    functions.conditional_format_wb(wb)
    functions.fill_lastrow(wb)


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
    functions.commercial(wb)


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


def shaded():
    wb = xw.Book.caller()
    functions.shaded(wb, shaded=True)


def unshaded():
    wb = xw.Book.caller()
    functions.shaded(wb, shaded=False)


def internal_costing():
    wb = xw.Book.caller()
    functions.internal_costing(wb)


def convert_legacy():
    wb = xw.Book.caller()
    functions.convert_legacy(wb)


def fill_formula_active_row():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_formula_active_row(wb, ws)


def delete_extra_empty_row():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.delete_extra_empty_row(ws)


def leave_application_checklist():
    checklists.leave_application_checklist()


def download_template():
    functions.creat_new_template()


def download_planner():
    functions.creat_new_planner()


def generate_sales_checklist():
    checklists.generate_sales_checklist()


def generate_firmed_proposal_checklist():
    wb = xw.Book.caller()
    checklists.generate_proposal_checklist(wb)


def generate_budgetary_proposal_checklist():
    wb = xw.Book.caller()
    checklists.generate_proposal_checklist(
        wb, proposal_type="budgetary", title="Budgetary Proposal Checklist"
    )


def generate_handover_checklist():
    pass


def update_template_version():
    wb = xw.Book.caller()
    functions.update_template_version(wb)
