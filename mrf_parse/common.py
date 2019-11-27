"""Commom functions
"""

import sys
from subprocess import call
from importlib import import_module

# Install dependency (openpyxl) and import
try:
    xl = import_module('openpyxl')
except ModuleNotFoundError:
    print('Installing dependency (openpyxl)')
    call((sys.executable, '-m', 'pip', 'install', '--user', 'openpyxl'))
    try:
        xl = import_module('openpyxl')
        print()
    except ModuleNotFoundError:
        print('Cannot install dependencies!')
        sys.exit()


def read_excel_file(fn):
    """Read Excel file

    Args:
        fn (str): File path

    Returns:
        wb (Workbook): Workbook object
    """
    wb = xl.load_workbook(fn, read_only=True)
    return wb


def read_excel_sheet(wb, sheetn):
    """Read Excel sheet

    Args:
        wb     (Workbook): Workbook object
        sheetn (str)     : Sheet name

    Returns:
        ws (ReadOnlyWorksheet): Read-only worksheet object
    """
    ws = wb[sheetn]
    return ws


def gen_empty_excel():
    """Generate new empty Excel file

    Returns:
        wb (Workbook) : Workbook (Single worksheet)
        ws (Worksheet): Worksheet object (Empty)
    """
    wb = xl.Workbook()
    ws = wb.active
    return (wb, ws)
