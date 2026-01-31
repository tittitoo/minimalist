# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Minimalist is a proposal writing and tendering tool that uses Excel as the UI and Python (via xlwings) for automation. The tool automates formula filling, text formatting, price calculations with currency conversion, and PDF checklist generation.

## Commands

### Run Tests
```bash
python tests.py           # Text formatting unit tests
python test_formulas.py   # Formula validation/regression tests
```

### Setup (one-time)
```bash
xlwings addin install     # Install Excel add-in
```

Tests run without Excel dependency (pure Python/pandas logic).

## Architecture

### Module Overview

- **excel.py** - xlwings bridge exposing Python functions to Excel. All functions wrapped with `@check_if_template` (validates template) and `@disable_screen_updating` (performance).
- **functions.py** (~2,400 lines) - Core business logic: formula filling, text processing, pricing calculations, row/column management.
- **checklists.py** - PDF checklist generation using ReportLab.
- **checklist_collections.py** - Predefined choice lists and checklist templates.
- **test_formulas.py** - Master formula registry used as regression tests to prevent formula changes.

### Key Patterns

**Excel Integration Flow:**
1. Excel (UI) → xlwings decorators (excel.py) → Business logic (functions.py)
2. PERSONAL.XLSB macro workbook lazy-loaded to avoid errors when Excel isn't running
3. Performance: manual calculation mode and screen updating disabled during operations

**Pricing Calculation Chain:**
UC (unit cost) → UCD (after discount) → UCDQ (converted to quote currency) → apply escalations (default, warranty, freight, risk) → RUPQ (recommended unit price)

**Formula Architecture:**
- Dynamic formulas using XMATCH/INDEX for robust references
- Master formula definitions in test_formulas.py serve as source of truth
- Batch processing for performance

### Excel Keyboard Shortcuts (via VBA)

| Shortcut | Action |
|----------|--------|
| Ctrl+E | Fill formula |
| Ctrl+J | Hide rows |
| Ctrl+M | Unhide rows |
| Ctrl+W | Insert rows |
| Ctrl+Q | Delete rows |

## Column Reference

Key Excel columns (see About.md for full list):
- UC/SC: Unit/Subtotal Cost in original currency
- UCD/SCD: After discount
- UCDQ/SCDQ: Converted to quoted currency
- RUPQ/RSPQ: Recommended prices
- Escalation factors: Default, Warranty, Freight, Special Terms, Risk

## Dependencies

Python 3.12.x required. Package management via `uv`.

Core: xlwings (Excel bridge), pandas, numpy, reportlab (PDF generation)
