"""Generate the Excel files
"""

import os.path
from sys import exit
from time import time
from datetime import datetime as dt
from argparse import Namespace, ArgumentTypeError, ArgumentParser as AP

from feedback import grab_feedback
from funds import total_funds
from hours import total_hours

MONTHS = ('January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December')
OUT_FNS = {'f': 'DISTRICT BOARD FEEDBACK',
           'h': 'TOTAL HOURS PER TENET',
           'm': 'TOTAL FUNDS RAISED PER DFI'}


def in_ok(path):
    """Check if indir is a folder

    Args:
        path (str): Indir path

    Returns:
        path (str): Valid directory
    """
    if not os.path.isdir(path):
        raise ArgumentTypeError(f'{path} is not an existing folder!')
    return path


def parse_sys_args(argv):
    """Argument parsing (semi-automatic)

    Args:
        argv (List[str]): An array of arguments, excluding the file name
                          (E.g., sys.argv[1:])

    Returns:
        args (argparse.Namespace): Parsed args
    """
    # Parse arguments
    par = AP(prog='mrf_parse', add_help=False,
             description=('Summarize MRFs\' feedbacks, hours, '
                          'and money (funds) data into another Excel file'),
             usage='%(prog)s month year procedure [-i [INDIR]] [-o [OUTDIR]]')
    par.add_argument('proc', metavar='procedure', choices=('f', 'h', 'm'),
                     help='Procedure ([f]eedbacks/[h]ours/[m]oney)')
    par.add_argument('-m', '--month', type=int, choices=range(1, 13),
                     default=dt.now().month,
                     help='Month number (1-12). Default current month')
    par.add_argument('-y', '--year', type=int,
                     help=('Year number. Default current year (Jan-Feb) '
                           'or last year (Mar-Dec). See README'))
    par.add_argument('-i', '--indir', nargs='?', default='.', type=in_ok,
                     help=('Directory of MRF Excel files. '
                           'Default current folder'))
    par.add_argument('-o', '--outdir', nargs='?', default='.',
                     help=('Directory to put the summary in.'
                           'Default current folder'))
    if not argv:
        par.print_help()
        exit(1)
    args = par.parse_args(argv)
    if args.year is None:
        t = dt.now()
        args.year = t.year - (1 if args.month not in (1, 2)
                              and t.month in (1, 2) else 0)
    return args


def prompt_args():
    """Prompt user for args so that this is more user-friendly

    Returns:
        args (argparse.Namespace): Args from user input
    """
    # Procedure
    while True:
        print()
        print('\n'.join(('Procedures:',
                         '    f: Feedback',
                         '    h: Hours',
                         '    m: Funds (Money)')))
        proc = input('Please choose a procedure to be run (f/h/m): ')
        if proc not in ('f', 'h', 'm'):
            print(f'{proc} is not a valid procedure (f/h/m)!')
            continue
        break
    # Month
    t = dt.now()
    mo_cur = t.month
    yr_cur = t.year
    while True:
        print()
        print('Please input a month (1-12).', end=' ')
        mo_str = input(f'Leave empty for current month ({mo_cur}): ')
        if mo_str:
            try:
                mo = int(mo_str)
            except ValueError:
                print(f'{mo_str} is not a number!')
                continue
            if mo not in range(1, 13):
                print(f'{mo} is not a month number (1-12)!')
                continue
        else:
            mo = mo_cur
        break
    # Year
    while True:
        print()
        print('Please input a year. Leave empty for', end=' ')
        if mo not in (1, 2) and mo_cur in (1, 2):
            print('last', end=' ')
            yr_def = yr_cur - 1
        else:
            print('current', end=' ')
            yr_def = yr_cur
        yr_str = input(f'year ({yr_def}). This is only used for Excel title: ')
        if yr_str:
            try:
                yr = int(yr_str)
            except ValueError:
                print(f'{yr_str} is not a number!')
                continue
        else:
            yr = yr_def
        break
    # Input directory
    while True:
        print()
        print('Please input the folder containing MRF files.', end=' ')
        indir = input('Leave empty for current folder: ')
        if not indir:
            indir = '.'
        if not os.path.isdir(indir):
            print(f'{indir} is not a folder!')
            continue
        break
    # Output directory
    print()
    print('Please input the folder to put the summary in.', end=' ')
    outdir = input('Leave empty for current folder: ')
    if not outdir:
        outdir = '.'
    print()
    return Namespace(month=mo, year=yr, proc=proc, indir=indir, outdir=outdir)


def main(argv):
    """Parse arguments, do calculations, and write results to Excel file

    Args:
        argv (List[str]): An array of arguments, excluding the file name
                          (E.g., sys.argv[1:])
    """
    args = parse_sys_args(argv) if argv else prompt_args()
    # Run calculations
    mo = MONTHS[args.month - 1]
    yr = str(args.year)
    begin = time()
    if args.proc == 'f':
        wb = grab_feedback(args.indir, mo, yr)
    elif args.proc == 'h':
        wb = total_hours(args.indir, mo, yr, (args.month + 9) % 12 + 1)
    elif args.proc == 'm':
        wb = total_funds(args.indir, mo, yr, (args.month + 9) % 12 + 1)
    else:
        raise AssertionError('You are messing with it, aren\'t you?')
    # Save file and writing unnecessary properties
    proc = OUT_FNS[args.proc]
    wb.properties.creator = wb.properties.lastModifiedBy = 'Stella Liang'
    wb.properties.title = proc
    # Try writing to file
    outdir = args.outdir
    while True:
        t = dt.now().strftime('%Y%m%d%H%M%S')
        out_fn = f'{os.path.join(outdir, proc)}-{t}.xlsx'
        if os.path.exists(out_fn):
            if os.path.isfile(out_fn):
                print('Output file already exists.', end=' ')
                cont = input('Overwrite? [y/N] ')
                if cont.lower() in ('y', 'yes'):
                    print('Overwriting...')
                else:
                    print('Overwrite cancelled!')
                    print('Cannot save to this folder.', end=' ')
                    outdir = input('Please choose a new folder: ')
                    continue
            else:
                print(f'{out_fn} already exists and cannot be overwritten!')
                print('Cannot save to this folder.', end=' ')
                outdir = input('Please choose a new folder: ')
                continue
        try:
            os.makedirs(outdir, exist_ok=True)
            wb.save(out_fn)
        except PermissionError:
            print('Cannot save to this folder.', end=' ')
            outdir = input('Please choose a new folder: ')
            continue
        break
    print(f'Summary saved to {out_fn}')
    print(f'Data extraction took {time()-begin:.3f} seconds')
