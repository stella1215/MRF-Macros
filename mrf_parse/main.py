"""Generate the Excel files
"""

import os.path
from sys import exit
from time import time
from argparse import ArgumentParser, RawDescriptionHelpFormatter, RawTextHelpFormatter

from feedback import grab_feedback
from funds import total_funds
from hours import total_hours

MONTHS = ('January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December')
OUT_FNS = {'f': 'DISTRICT BOARD FEEDBACK',
           'h': 'TOTAL HOURS PER TENET', 'm': 'TOTAL FUNDS RAISED PER DFI'}


class CustomFormatter(RawTextHelpFormatter):
    """Custom argparse formatter with both RawDescriptionHelpFormatter and RawTextHelpFormatter and a few changes

    Args: See `argparse.HelpFormatter` class
    """

    def __init__(self, prog, indent_increment=4, max_help_position=35, width=1000):
        super().__init__(prog, indent_increment, max_help_position, width)


def main(argv):
    """Main function: Parse arguments, call functions to do calculations, and write results to Excel file

    Args:
        argv (Iterable[str]): An array of arguments, excluding the file name (E.g., sys.argv[1:])
    """
    # Parse arguments
    eg = '''examples:
    %(prog)s 12 2019 f                          - Extract feedback from all Excel files in this folder and put the output in this folder
    %(prog)s 1 2020 m -i input/files -o output  - Extract money (funds) information from all Excel files in `input/files` folder and put the output in `output` folder'''
    par = ArgumentParser(prog='mrf_parse', add_help=False, formatter_class=CustomFormatter,
                         description='Summarize MRFs\' feedbacks, hours, and money(funds) data into another Excel file',
                         usage='%(prog)s month year procedure [-i [INDIR]] [-o [OUTDIR]]', epilog=eg)
    par.add_argument('month', metavar='month', type=int, choices=range(1, 13),
                     help='Month number (Choose from 1-12)')
    par.add_argument('year', type=int,
                     help='Year number (No use, only for reference)')
    par.add_argument('proc', metavar='procedure', choices=('f', 'h', 'm'),
                     help='Procedure name (Choose from [f]eedbacks, [h]ours, [m]oney)')
    par.add_argument('-i', '--indir', nargs='?', default='.',
                     help='Directory of MRF Excel files. Default is current folder')
    par.add_argument('-o', '--outdir', nargs='?', default='.',
                     help='Directory to put the summary in. Default is current folder')
    if not argv:
        par.print_help()
        exit(1)
    args = par.parse_args(argv)
    # Check indir existence
    if not os.path.isdir(args.indir):
        print('The `indir` is not a folder!')
        return
    # Check outdir viability
    out_fn = os.path.join(args.outdir, OUT_FNS[args.proc]) + '.xlsx'
    if os.path.exists(out_fn):
        if os.path.isfile(out_fn):
            cont = input(
                'The output file already exists in the directory specified. Overwrite? [y/N] ')
            if cont.lower() == 'y':
                print('Overwriting...')
            else:
                print('Overwrite cancelled!')
                exit(-1)
        else:
            print(f'{out_fn} already exists and cannot be overwritten!')
            exit(1337)
    # Run calculations
    mo = MONTHS[args.month - 1]
    yr = str(args.year)
    begin = time()
    if args.proc == 'f':
        wb = grab_feedback(args.indir, mo, yr)
    elif args.proc == 'h':
        wb = total_hours(args.indir, mo, yr)
    elif args.proc == 'm':
        wb = total_funds(args.indir, mo, yr)
    else:
        raise AssertionError('You are messing with it, aren\'t you?')
    # Save file and writing unnecessary properties
    wb.properties.creator = wb.properties.lastModifiedBy = 'Stella Liang'
    wb.properties.title = OUT_FNS[args.proc]
    wb.save(out_fn)
    print(f'Summary saved to {out_fn}')
    print(f'Data extraction took {time()-begin:.3f} seconds')
