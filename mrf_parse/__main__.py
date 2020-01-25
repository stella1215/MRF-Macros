import sys

from main import main

if sys.version_info.major < 3:
    print('This program requires Python 3.',
          'Download latest release at https://www.python.org/downloads/')

try:
    main(sys.argv[1:])
except KeyboardInterrupt:
    print()
    print('Terminated by user input')
