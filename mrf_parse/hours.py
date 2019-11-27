"""Generate the Excel files
"""

import os

from common import read_excel_file, read_excel_sheet, gen_empty_excel

OUT_FNS = ('DISTRICT BOARD FEEDBACK.xlsx',
           'TOTAL HOURS PER TENET.xlsx',
           'TOTAL FUNDS RAISED PER DFI.xlsx')


def calc_hours(ws, ws_mo, r, c):
    """Calculate monthly hours for a category

    Args:
        ws    (Worksheet): 'Annual Totals' Worksheet Object
        ws_mo (Worksheet): Respective month Worksheet Object
        r     (int)      : Row number
        c     (str)      : Column letter

    Returns:
        result (float): Monthly hours for the category
    """
    formula = ws[f'{c}{r}'].value.split('!')[-1]
    total = ws_mo[formula].value[5:-1]
    result = 0.0
    for row in ws_mo[total]:
        if not row:
            break
        v = row[0].value
        if v is not None:
            result += row[0].value
    return result


def month_hours(wb, ws, r):
    """Read monthly hours

    Args:
        wb (Workbook) : Workbook Object
        ws (Worksheet): 'Annual Totals' Worksheet Object
        r  (int)      : Row number

    Returns:
        serv (float): Service hours for the month
        lead (float): Leadership hours for the month
        fell (float): Fellowship hours for the month
    """
    mo = ws[f'C{r}'].value
    ws_mo = read_excel_sheet(wb, mo)
    serv = calc_hours(ws, ws_mo, r, 'F')
    lead = calc_hours(ws, ws_mo, r, 'H')
    fell = calc_hours(ws, ws_mo, r, 'J')
    return (serv, lead, fell)


def total_hours(folder, month, year):
    """Get total hours and put into new Excel file

    Args:
        folder (str): Folder path
        month  (str): Month name
        year   (str): Year

    Returns:
        wb_new (Workbook): Workbook Object
    """
    wb_new, ws_new = gen_empty_excel()
    ws_new.append((f'TOTAL HOURS AS OF ({month} {year})',))
    ws_new.merge_cells('A1:D1')
    ws_new.append(('CLUB NAME', 'TOTAL SERVICE HOURS',
                   'TOTAL LEADERSHIP HOURS', 'TOTAL FELLOWSHIP HOURS'))
    # Iterate through MRFs
    count = 0
    for file in os.listdir(folder):
        # Preliminary file check
        if file.split('.')[-1] not in ('xls', 'xlsx', 'xlsm') or os.path.isdir(file) or file in OUT_FNS:
            continue
        print(f'Reading {file}...')
        file_path = os.path.join(folder, file)
        wb = read_excel_file(file_path)
        ws = read_excel_sheet(wb, 'Annual Totals')
        school_name = read_excel_sheet(wb, 'Club Administration')['A12'].value
        # Calculate hours
        serv = lead = fell = 0.0
        for r in range(39, 51):
            serv_m, lead_m, fell_m = month_hours(wb, ws, r)
            serv += serv_m
            lead += lead_m
            fell += fell_m
        ws_new.append((school_name, serv, lead, fell))
        count += 1
    print('Post-processing...')
    if not count:
        print('W: No Excel file processed!')
    nrows = len(ws_new['A'])
    ws_new.append(
        ('TOTAL HOURS', f'=SUM(B3:B{nrows})', f'=SUM(C3:C{nrows})', f'=SUM(D3:D{nrows})'))
    for row in ws_new[f'B3:D{nrows+1}']:
        for cell in row:
            cell.number_format = '0.00'
    return wb_new
