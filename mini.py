#!/usr/bin/env -S uv run --quiet --no-upgrade --script
# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "click",
#     "xlwings",
#     "pandas",
#     "requests",
#     "reportlab",
# ]
# ///
"""
CLI tool for minimalist Excel automation.
Run operations without Excel stealing focus.

Usage:
    ./mini.py fix <file>                # Fill formulas and fix workbook
    ./mini.py commercial <file>         # Generate commercial PDF (prompts to run fix first)
    ./mini.py technical <file>          # Generate technical PDF
"""

import sys
import time
from pathlib import Path

import click
import xlwings as xw

import functions
import hide
from excel import (
    is_script_running,
    acquire_lock,
    release_lock,
)

# CLI Mode Alert Handling


def cli_alert(message, title=None, buttons="ok", mode=None, callback=None):
    """Terminal-based alert replacement for xw.apps.active.alert."""
    click.echo(f"\n[INFO] {message}\n")
    return True


def enable_cli_mode():
    """Patch xlwings alert to use terminal output."""
    xw.App.alert = lambda self, *args, **kwargs: cli_alert(*args, **kwargs)


# Workbook Operations


def open_workbook(filepath: str):
    """Open workbook using existing Excel app or create new instance."""
    # Use existing Excel app if available to avoid PERSONAL.XLSB conflict
    if xw.apps:
        app = xw.apps.active
        created_app = False
        original_screen_updating = app.screen_updating
    else:
        # Run invisible to avoid distracting user (benchmarked 1.8x faster on macOS)
        app = xw.App(visible=False)
        created_app = True
        original_screen_updating = True

    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(filepath, password=hide.legacy)

    return app, wb, created_app, original_screen_updating


def run_fix_workbook(filepath: str) -> bool:
    """Run fill_formula_wb operation (Fix Workbook)."""
    start_time = time.perf_counter()
    click.echo(f"Opening: {filepath}")
    app, wb, created_app, original_screen_updating = open_workbook(filepath)

    try:
        if "Config" not in wb.sheet_names:
            click.echo("[ERROR] The excel file is not a recognized template.", err=True)
            return False

        click.echo("Updating template version...")
        functions.update_template_version(wb)

        click.echo("Cleaning up empty rows...")
        functions.delete_extra_empty_row_wb(wb)
        functions.delete_extra_empty_row_wb(wb)

        click.echo("Numbering titles...")
        count, step = functions.get_num_scheme(wb)
        functions.number_title(wb, count=count, step=step)

        click.echo("Filling formulas...")
        functions.fill_formula_wb(wb)

        click.echo("Formatting text...")
        functions.format_text(
            wb,
            indent_description=True,
            bullet_description=True,
            title_lineitem_or_description=True,
        )

        click.echo("Formatting cells...")
        functions.format_cell_data(wb)

        click.echo("Adjusting columns...")
        functions.adjust_columns_wb(wb)

        click.echo("Applying conditional formatting...")
        functions.conditional_format_wb(wb)

        click.echo("Filling subtotals...")
        functions.fill_lastrow(wb)

        click.echo("Recalculating...")
        app.calculate()

        wb.save()
        elapsed = time.perf_counter() - start_time
        click.echo(f"[SUCCESS] Workbook saved: {filepath}")
        click.echo(f"[TIME] fix_workbook completed in {elapsed:.2f}s")
        return True

    except Exception as e:
        click.echo(f"[ERROR] {e}", err=True)
        return False
    finally:
        wb.close()
        app.screen_updating = original_screen_updating
        if created_app:
            app.quit()


def run_commercial(filepath: str) -> bool:
    """Run commercial PDF generation."""
    start_time = time.perf_counter()
    click.echo(f"Opening: {filepath}")
    app, wb, created_app, original_screen_updating = open_workbook(filepath)

    try:
        if "Config" not in wb.sheet_names:
            click.echo("[ERROR] The excel file is not a recognized template.", err=True)
            return False

        click.echo("Generating commercial PDF...")
        functions.commercial(wb, show_pdf=False)
        elapsed = time.perf_counter() - start_time
        click.echo("[SUCCESS] Commercial PDF generated.")
        click.echo(f"[TIME] commercial completed in {elapsed:.2f}s")
        return True

    except Exception as e:
        click.echo(f"[ERROR] {e}", err=True)
        return False
    finally:
        wb.close()
        app.screen_updating = original_screen_updating
        if created_app:
            app.quit()


def run_technical(filepath: str) -> bool:
    """Run technical PDF generation."""
    start_time = time.perf_counter()
    click.echo(f"Opening: {filepath}")
    app, wb, created_app, original_screen_updating = open_workbook(filepath)

    try:
        click.echo("Generating technical PDF...")
        functions.technical(wb, show_pdf=False)
        elapsed = time.perf_counter() - start_time
        click.echo("[SUCCESS] Technical PDF generated.")
        click.echo(f"[TIME] technical completed in {elapsed:.2f}s")
        return True

    except Exception as e:
        click.echo(f"[ERROR] {e}", err=True)
        return False
    finally:
        wb.close()
        app.screen_updating = original_screen_updating
        if created_app:
            app.quit()


def run_with_lock(operation, filepath: str) -> bool:
    """Run operation with lock to prevent concurrent execution."""
    if is_script_running():
        click.echo(
            "[ERROR] Another operation is already running. Please wait.", err=True
        )
        return False

    acquire_lock()
    try:
        return operation(filepath)
    finally:
        release_lock()


# CLI Commands


@click.group()
def cli() -> None:
    """CLI tool for minimalist Excel automation."""
    enable_cli_mode()


@cli.command("fix_workbook")
@click.argument("file", type=click.Path(exists=True))
def fix_workbook_cmd(file):
    """Fill formulas and fix workbook."""
    filepath = str(Path(file).resolve())
    success = run_with_lock(run_fix_workbook, filepath)
    sys.exit(0 if success else 1)


@cli.command("fix")
@click.argument("file", type=click.Path(exists=True))
def fix_cmd(file):
    """Alias for fix_workbook."""
    filepath = str(Path(file).resolve())
    success = run_with_lock(run_fix_workbook, filepath)
    sys.exit(0 if success else 1)


@cli.command("commercial")
@click.argument("file", type=click.Path(exists=True))
@click.option("--fix", is_flag=True, help="Run fix_workbook first without prompting")
def commercial_cmd(file, fix):
    """Generate commercial PDF proposal."""
    filepath = str(Path(file).resolve())

    run_fix = fix or click.confirm("Run fix_workbook first?", default=True)
    if run_fix:
        click.echo("Running fix_workbook first...")
        if not run_with_lock(run_fix_workbook, filepath):
            sys.exit(1)

    success = run_with_lock(run_commercial, filepath)
    sys.exit(0 if success else 1)


@cli.command("technical")
@click.argument("file", type=click.Path(exists=True))
def technical_cmd(file):
    """Generate technical PDF proposal."""
    filepath = str(Path(file).resolve())
    success = run_with_lock(run_technical, filepath)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    cli()
