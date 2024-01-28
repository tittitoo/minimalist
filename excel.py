""" Created so that python fucntions are available in Excel.
    © Thiha Aung
"""

import xlwings as xw  # type: ignore
import functions
import checklists
import checklist_collections as cc

# from reportlab.lib.colors import lightcyan, black, white, lightyellow, blue


class IsNotTemplateException(BaseException):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)


# My first use of decotator
def check_if_template(func):
    def wrapper(*args, **kwargs):
        try:
            wb = xw.Book.caller()
            if "Config" not in wb.sheet_names:
                raise IsNotTemplateException(
                    "The excel file is not a recognized template."
                )
            func(*args, **kwargs)
        except IsNotTemplateException as e:
            xw.apps.active.alert(f"{e}")

    return wrapper


@check_if_template
def fill_formula():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_formula(ws)


# Fix the whole workbook. The function name will later change to fix_workbook
# Now is tied ot Fix Wrokbook in Excel
@check_if_template
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


@check_if_template
def subtotal():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_lastrow_sheet(wb, ws)


@check_if_template
def subtotal_wb():
    wb = xw.Book.caller()
    functions.fill_lastrow(wb)


@check_if_template
def unhide_columns():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.unhide_columns(ws)


@check_if_template
def summary():
    wb = xw.Book.caller()
    functions.summary(wb, discount=False)


@check_if_template
def summary_discount():
    wb = xw.Book.caller()
    functions.summary(wb, discount=True)


@check_if_template
def summary_detail():
    wb = xw.Book.caller()
    functions.summary(wb, discount=False, detail=True)


@check_if_template
def summary_detail_discount():
    wb = xw.Book.caller()
    functions.summary(wb, discount=True, detail=True)


@check_if_template
def number_title():
    wb = xw.Book.caller()
    functions.number_title(wb)


@check_if_template
def hide_columns():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.hide_columns(ws)


def technical():
    wb = xw.Book.caller()
    functions.technical(wb)


@check_if_template
def prepare_to_print_commercial():
    wb = xw.Book.caller()
    functions.prepare_for_print_commercial(wb)


@check_if_template
def print_commercial():
    wb = xw.Book.caller()
    functions.commercial(wb)


@check_if_template
def conditional_format_wb():
    wb = xw.Book.caller()
    functions.conditional_format_wb(wb)


@check_if_template
def fix_unit_price():
    wb = xw.Book.caller()
    functions.fix_unit_price(wb)


@check_if_template
def format_text():
    wb = xw.Book.caller()
    functions.format_text(wb, title_lineitem_or_description=True)


@check_if_template
def indent_description():
    wb = xw.Book.caller()
    functions.format_text(wb, indent_description=True, bullet_description=True)


@check_if_template
def shaded():
    wb = xw.Book.caller()
    functions.shaded(wb, shaded=True)


@check_if_template
def unshaded():
    wb = xw.Book.caller()
    functions.shaded(wb, shaded=False)


@check_if_template
def internal_costing():
    wb = xw.Book.caller()
    functions.internal_costing(wb)


def convert_legacy():
    wb = xw.Book.caller()
    functions.convert_legacy(wb)


@check_if_template
def fill_formula_active_row():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_formula_active_row(wb, ws)


@check_if_template
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


@check_if_template
def generate_firmed_proposal_checklist():
    wb = xw.Book.caller()
    checklists.generate_proposal_checklist(wb)


@check_if_template
def generate_budgetary_proposal_checklist():
    wb = xw.Book.caller()
    checklists.generate_proposal_checklist(
        wb, proposal_type="budgetary", title="Budgetary Proposal Checklist"
    )


@check_if_template
def update_template_version():
    wb = xw.Book.caller()
    functions.update_template_version(wb)


def generate_handover_checklist():
    pass
