"""Generate funds Excel file
"""

import os

from common import read_excel_file, read_excel_sheet, gen_empty_excel

OUT_FNS = ('DISTRICT BOARD FEEDBACK.xlsx',
           'TOTAL HOURS PER TENET.xlsx',
           'TOTAL FUNDS RAISED PER DFI.xlsx')


def calc_funds(ws, ws_mo, r, c):
    """Calculate monthly funds for a project

    Args:
        ws    (Worksheet): 'Annual Totals' Worksheet Object
        ws_mo (Worksheet): Respective month Worksheet Object
        r     (int)      : Row number
        c     (str)      : Column letter

    Returns:
        result (float): Monthly funds for the project
    """
    formula = ws[f'{c}{r}'].value.split('!')[-1]
    reference = ws_mo[formula].value[1:]
    total = ws_mo[reference].value[5:-1]
    result = 0.0
    for row in ws_mo[total]:
        if not row:
            break
        v = row[0].value
        if v is not None:
            result += row[0].value
    return result


def month_funds(wb, ws, r):
    """Read monthly funds

    Args:
        wb (Workbook) : Workbook Object
        ws (Worksheet): 'Annual Totals' Worksheet Object
        r  (int)      : Row number

    Returns:
        ptp    (float): PTP for the month
        trevor (float): Trevor project for the month
        kfh    (float): Kiwanis Family House for the month
        others (float): Others for the month
    """
    mo = ws[f'A{r}'].value
    ws_mo = read_excel_sheet(wb, mo)
    ptp = calc_funds(ws, ws_mo, r, 'C')
    trevor = calc_funds(ws, ws_mo, r, 'E')
    kfh = calc_funds(ws, ws_mo, r, 'G')
    others = calc_funds(ws, ws_mo, r, 'M')
    return (ptp, trevor, kfh, others)


def total_funds(folder, month, year):
    """Get total funds and put into new Excel file

    Args:
        folder (str): Folder path
        month  (str): Month name
        year   (str): Year

    Returns:
        wb_new (Workbook): Workbook Object
    """
    wb_new, ws_new = gen_empty_excel()
    ws_new.append((f'TOTAL FUNDS AS OF ({month} {year})',))
    ws_new.merge_cells('A1:E1')
    ws_new.append(('CLUB NAME', 'PTP', 'TREVOR PROJECT',
                   'KFH', 'OTHER CHARITIES'))
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
        # Calculate funds
        ptp = trevor = kfh = others = 0.0
        for r in range(60, 72):
            ptp_m, trevor_m, kfh_m, others_m = month_funds(wb, ws, r)
            ptp += ptp_m
            trevor += trevor_m
            kfh += kfh_m
            others += others_m
        ws_new.append((school_name, ptp, trevor, kfh, others))
        count += 1
    print('Post-processing...')
    if not count:
        print('W: No Excel file processed!')
    nrows = len(ws_new['A'])
    ws_new.append(
        ('TOTAL FUNDS', f'=SUM(B3:B{nrows})', f'=SUM(C3:C{nrows})', f'=SUM(D3:D{nrows})', f'=SUM(E3:E{nrows})'))
    for row in ws_new[f'B3:E{nrows+1}']:
        for cell in row:
            cell.number_format = '$#,##0.00;[Red]-$#,##0.00'
    return wb_new
