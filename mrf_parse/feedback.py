"""Generate feedback Excel file
"""

import os

from common import read_excel_file, gen_empty_excel, OUT_FNS


def grab_feedback(folder, month, year):
    """Get feedbacks and put into new Excel file

    Args:
        folder (str): Folder path
        month  (str): Month name
        year   (str): Year

    Returns:
        wb_new (Workbook): Workbook Object
    """
    wb_new, ws_new = gen_empty_excel()
    ws_new.append(('CLUB NAME', f'DISTRICT BOARD FEEDBACK ({month} {year})'))
    # Iterate through MRFs
    count = 0
    for file in os.listdir(folder):
        # Preliminary file check
        if (file.split('.')[-1] not in ('xls', 'xlsx', 'xlsm') or
                os.path.isdir(file) or file[:-20] in OUT_FNS):
            continue
        print(f'Reading {file}...')
        file_path = os.path.join(folder, file)
        wb = read_excel_file(file_path)
        try:
            ws = wb[month]
        except KeyError:
            continue
        school_name = wb['Club Administration']['A12'].value
        feedback = ws['A64'].value
        ws_new.append((school_name, feedback))
        count += 1
    print()
    print('Post-processing...')
    if not count:
        print('W: No Excel file processed!')
    return wb_new
