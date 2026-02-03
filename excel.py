"""
Created so that python fucntions are available in Excel.
© Thiha Aung (infowizard@gmail.com)
"""

import sys
import time
import tempfile
from pathlib import Path
import xlwings as xw  # type: ignore
import functions
import checklists

# Lock file to prevent multiple simultaneous executions
LOCK_FILE = Path(tempfile.gettempdir()) / "minimalist_running.lock"
WARN_LOCK_FILE = Path(tempfile.gettempdir()) / "minimalist_warning.lock"
LOCK_TIMEOUT = 300  # 5 minutes - ignore stale locks older than this
WARN_LOCK_TIMEOUT = 30  # 30 seconds - warning popup should be dismissed by then


def is_lock_stale(lock_path, timeout):
    """Check if a lock file is stale (older than timeout seconds)."""
    try:
        age = time.time() - lock_path.stat().st_mtime
        return age > timeout
    except Exception:
        return True


def is_script_running():
    """Check if another instance of the script is already running."""
    if not LOCK_FILE.exists():
        return False
    if is_lock_stale(LOCK_FILE, LOCK_TIMEOUT):
        # Stale lock from crash - remove it
        LOCK_FILE.unlink(missing_ok=True)
        return False
    return True


def is_warning_showing():
    """Check if a warning popup is already being displayed."""
    if not WARN_LOCK_FILE.exists():
        return False
    if is_lock_stale(WARN_LOCK_FILE, WARN_LOCK_TIMEOUT):
        # Stale warning lock - remove it
        WARN_LOCK_FILE.unlink(missing_ok=True)
        return False
    return True


def acquire_lock():
    """Create lock file to indicate script is running."""
    try:
        LOCK_FILE.touch()
        return True
    except Exception:
        return False


def release_lock():
    """Remove lock file when script finishes."""
    try:
        LOCK_FILE.unlink(missing_ok=True)
    except Exception:
        pass


def acquire_warn_lock():
    """Create warning lock file to prevent multiple warning popups."""
    try:
        WARN_LOCK_FILE.touch()
        return True
    except Exception:
        return False


def release_warn_lock():
    """Remove warning lock file after popup is dismissed."""
    try:
        WARN_LOCK_FILE.unlink(missing_ok=True)
    except Exception:
        pass


# import checklist_collections as cc

# from reportlab.lib.colors import lightcyan, black, white, lightyellow, blue


# Progress indicator helper functions
def update_status(app, message):
    """Update Excel status bar with message (cross-platform)"""
    app.status_bar = message


def set_busy_cursor(app, busy=True):
    """Set cursor to hourglass (busy=True) or default (busy=False). Windows only."""
    if sys.platform == "win32":
        app.api.Cursor = 2 if busy else -4143  # xlWait=2, xlDefault=-4143


class IsNotTemplateException(Exception):
    """Raise if the excel file is not recognized as a template."""

    # def __init__(self, *args: object) -> None:
    #     super().__init__(*args)


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
            xw.apps.active.alert(f"{e}")  # type: ignore
        except KeyError:
            # Workbook was renamed during operation (e.g., commercial PDF) - ignore
            pass

    return wrapper


def is_excel_available(app):
    """Check if Excel app is still available."""
    try:
        _ = app.version
        return True
    except Exception:
        return False


def retry_com_operation(operation, max_retries=3, delay=0.5):
    """Retry a COM operation if Excel is busy (error 0x800ac472)."""
    for attempt in range(max_retries):
        try:
            return operation()
        except Exception as e:
            error_str = str(e).lower()
            # Check for "Excel is busy" COM error
            if "800ac472" in error_str or "-2146777998" in str(e):
                if attempt < max_retries - 1:
                    time.sleep(delay * (attempt + 1))  # Exponential backoff
                    continue
            # Check if Excel quit/crashed - don't retry
            if (
                "disconnected" in error_str
                or "rpc" in error_str
                or "server" in error_str
            ):
                return None
            raise
    return None


def disable_screen_updating(func):
    "Disable excel screen updating and automatic calculation to improve performance"

    def wrapper(*args, **kwargs):
        # Check if another instance is already running
        if is_script_running():
            # Only show warning if not already showing (prevents multiple popups)
            if not is_warning_showing():
                acquire_warn_lock()
                try:
                    xw.apps.active.alert(
                        "A script is already running. Please wait for it to finish."
                    )
                except Exception:
                    pass
                finally:
                    release_warn_lock()
            return

        # Acquire lock before proceeding
        acquire_lock()

        try:
            try:
                app = xw.Book.caller().app
            except KeyError:
                # Workbook was renamed - try to get app from active instance
                app = xw.apps.active
            if app is None:
                release_lock()
                return
            # Store original settings with retry (default to safe values if None)
            original_calculation = (
                retry_com_operation(lambda: app.calculation) or "automatic"
            )
            original_screen_updating = retry_com_operation(lambda: app.screen_updating)
            if original_screen_updating is None:
                original_screen_updating = True
            success = False
            try:
                retry_com_operation(lambda: setattr(app, "screen_updating", False))
                retry_com_operation(lambda: setattr(app, "calculation", "manual"))
                set_busy_cursor(app, busy=True)
                update_status(app, "Running please wait ...")
                func(*args, **kwargs)
                success = True
            except KeyError:
                # Workbook was renamed during operation (e.g., commercial PDF) - this is expected
                success = True
            except Exception as e:
                print(f"Error during function execution -> {e}")
                raise
            finally:
                # Only restore settings if Excel is still available
                if is_excel_available(app):
                    try:
                        # Restore calculation mode first, then recalculate
                        retry_com_operation(
                            lambda: setattr(app, "calculation", original_calculation)
                        )
                        # Force full recalculation to avoid stale value errors
                        retry_com_operation(lambda: app.calculate())
                        # Restore screen updating last
                        retry_com_operation(
                            lambda: setattr(
                                app, "screen_updating", original_screen_updating
                            )
                        )
                        set_busy_cursor(app, busy=False)
                        if success:
                            update_status(app, "Ready")
                    except Exception:
                        # Cleanup failed but main operation succeeded - ignore
                        pass
        finally:
            # Always release the lock, even if an error occurred
            release_lock()

    return wrapper


@check_if_template
@disable_screen_updating
def fill_formula():
    wb = xw.Book.caller()
    app = wb.app
    ws = wb.sheets.active
    update_status(app, "Filling formulas...")
    functions.fill_formula(ws)
    # Added number_title so that it is also tied to ctrl+e shortcut
    update_status(app, "Numbering titles...")
    count, step = functions.get_num_scheme(wb)
    functions.number_title(wb, count=count, step=step)
    # Reset font sizes to default (Arial 12 for data, 9 for headers)
    update_status(app, "Formatting cells...")
    functions.format_cell_data_sheet(ws)


# Fix the whole workbook. The function name will later change to fix_workbook
# Now is tied ot Fix Wrokbook in Excel
@check_if_template
@disable_screen_updating
def fill_formula_wb():
    wb = xw.Book.caller()
    app = wb.app
    update_status(app, "Updating template version check...")
    functions.update_template_version(wb)
    update_status(app, "Cleaning up empty rows...")
    functions.delete_extra_empty_row_wb(wb)
    # Calling twice as sometimes some rows are missed.
    functions.delete_extra_empty_row_wb(wb)
    update_status(app, "Numbering titles...")
    count, step = functions.get_num_scheme(wb)
    functions.number_title(wb, count=count, step=step)
    update_status(app, "Filling formulas...")
    functions.fill_formula_wb(wb)
    update_status(app, "Formatting text...")
    functions.format_text(
        wb,
        indent_description=True,
        bullet_description=True,
        title_lineitem_or_description=True,
    )
    update_status(app, "Formatting cells...")
    functions.format_cell_data(wb)
    update_status(app, "Adjusting columns...")
    functions.adjust_columns_wb(wb)
    update_status(app, "Applying conditional formatting...")
    functions.conditional_format_wb(wb)
    update_status(app, "Filling subtotals...")
    functions.fill_lastrow(wb)
    update_status(app, "Recalculating...")
    # Force recalculation at the end to avoid stale value errors
    wb.app.calculate()


@check_if_template
@disable_screen_updating
def subtotal():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_lastrow_sheet(wb, ws)


@check_if_template
@disable_screen_updating
def subtotal_wb():
    wb = xw.Book.caller()
    functions.fill_lastrow(wb)


@check_if_template
@disable_screen_updating
def unhide_columns():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.unhide_columns(ws)


@check_if_template
@disable_screen_updating
def summary():
    wb = xw.Book.caller()
    functions.summary(wb, discount=False)


@check_if_template
@disable_screen_updating
def summary_discount():
    wb = xw.Book.caller()
    functions.summary(wb, discount=True)


@check_if_template
@disable_screen_updating
def summary_detail():
    wb = xw.Book.caller()
    functions.summary(wb, discount=False, detail=True)


@check_if_template
@disable_screen_updating
def summary_detail_discount():
    wb = xw.Book.caller()
    functions.summary(wb, discount=True, detail=True)


@check_if_template
@disable_screen_updating
def number_title():
    wb = xw.Book.caller()
    count, step = functions.get_num_scheme(wb)
    functions.number_title(wb, count=count, step=step)


@check_if_template
@disable_screen_updating
def hide_columns():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.hide_columns(ws)


@disable_screen_updating
# TODO: does not work on Sharepoint
def technical():
    wb = xw.Book.caller()
    functions.technical(wb)


@check_if_template
@disable_screen_updating
# TODO: does not work on Sharepoint
def print_commercial():
    wb = xw.Book.caller()
    functions.commercial(wb)


@check_if_template
@disable_screen_updating
def conditional_format_wb():
    wb = xw.Book.caller()
    functions.conditional_format_wb(wb)


@check_if_template
@disable_screen_updating
def fix_unit_price():
    wb = xw.Book.caller()
    functions.fix_unit_price(wb)


@check_if_template
@disable_screen_updating
def format_text():
    wb = xw.Book.caller()
    functions.format_text(
        wb,
        title_lineitem_or_description=True,
        indent_description=True,
        bullet_description=True,
    )


@check_if_template
@disable_screen_updating
# TODO: intend_description now do fill_formula_wb, to change to consistent function name
# To change to fill_formula_wb
def indent_description():
    wb = xw.Book.caller()
    functions.fill_formula_wb(wb)
    # functions.format_text(wb, indent_description=True, bullet_description=True)


@check_if_template
@disable_screen_updating
def shaded():
    wb = xw.Book.caller()
    functions.shaded(wb, shaded=True)


@check_if_template
@disable_screen_updating
def unshaded():
    wb = xw.Book.caller()
    functions.shaded(wb, shaded=False)


@check_if_template
@disable_screen_updating
# TODO: does not work on Sharepoint
def internal_costing():
    wb = xw.Book.caller()
    functions.internal_costing(wb)


def convert_legacy():
    wb = xw.Book.caller()
    functions.convert_legacy(wb)


@check_if_template
@disable_screen_updating
def fill_formula_active_row():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_formula_active_row(wb, ws)


@check_if_template
@disable_screen_updating
def delete_extra_empty_row():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.delete_extra_empty_row(ws)


def leave_application_checklist():
    checklists.leave_application_checklist()


def download_template():
    functions.create_new_template()


def download_planner():
    functions.create_new_planner()


def generate_sales_checklist():
    checklists.generate_sales_checklist()


@check_if_template
@disable_screen_updating
def generate_firmed_proposal_checklist():
    wb = xw.Book.caller()
    checklists.generate_proposal_checklist(wb)


@check_if_template
@disable_screen_updating
def generate_budgetary_proposal_checklist():
    wb = xw.Book.caller()
    checklists.generate_proposal_checklist(
        wb, proposal_type="budgetary", title="Budgetary Proposal Checklist"
    )


@check_if_template
@disable_screen_updating
def update_template_version():
    wb = xw.Book.caller()
    functions.update_template_version(wb)


@check_if_template
@disable_screen_updating
def generate_handover_checklist():
    wb = xw.Book.caller()
    checklists.generate_handover_checklist(wb)


@check_if_template
@disable_screen_updating
def generate_general_checklist():
    wb = xw.Book.caller()
    checklists.generate_general_checklist(wb)
