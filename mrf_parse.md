# MRF Parse

This is a collection of scripts that extract data from MRF files.

## Prerequisites

* Python 3 (from [https://www.python.org/downloads/](https://www.python.org/downloads/)).
* OpenPyXl (Automatically installed upon first use).
  * See [OpenPyXl Manual Installation](#openpyxl-manual-installation) if automatic installation fails.

### OpenPyXl Manual Installation

#### Windows OpenPyXl Installation

1. Press `Win` and `R` key at the same time. A dialog called **Run** should pop up.
2. Type `cmd` and press **OK**. A terminal window should pop up.
3. In the terminal, type `pip install --user openpyxl` and press **Enter**/**Return** on your keyboard.
4. Wait for installation to complete.

#### macOS OpenPyXl Installation

1. Press `Command` and `Space` at the same time to bring up **Spotlight Search**.
2. Type `Terminal`.
3. Click on **Terminal - Utilities**. A terminal window should pop up.
4. In the terminal, type `pip3 install --user openpyxl` and press `Enter`/`Return` on your keyboard.
5. Wait for installation to complete.

## Basic Usage

### Windows Basic Usage

1. Double-click on `mrf_parse.bat` in this folder.
2. Proceed as the program instructs.

### macOS Basic Usage Method 1 (Unverified)

1. Double-click on `mrf_parse.command` in this folder.
2. Proceed as the program instructs.

### macOS Basic Usage Method 2, Linux Basic Usage

1. If you have done steps **2-7** before, skip to step **8**.
2. *(One-time)* Press `Command` and `Space` at the same time to bring up **Spotlight Search**.
3. *(One-time)* Type `Terminal`.
4. *(One-time)* Click on **Terminal - Utilities**. A terminal window should pop up.
5. *(One-time)* Drag `mrf_parse.sh` onto the Terminal window. A string should appear, e.g., `/Users/cnhcirclek/Downloads/MRF-Macros/mrf_parse.sh`.
6. *(One-time)* Type `chmod a+x` and a `Space` **before** the string, e.g., `chmod a+x /Users/cnhcirclek/Downloads/MRF-Macros/mrf_parse.sh`. User *arrow keys* to navigate cursor as necessary.
7. *(One-time)* Press `Enter`/`Return` on your keyboard. Close the terminal after completion.
8. Double-click `mrf_parse.sh` in this folder.
9. Proceed as the program instructs.

## Advanced Usage

1. Open a terminal of your choosing in this directory.
2. Execute `py -3 mrf_parse` on Windows, or `python3 mrf_parse` on macOS and Linux with appropriate arguments. The script will guide you through the process if no arguments are given.

```text
Usage: mrf_parse procedure [-m (1-12)] [-y YEAR] [-i INDIR] [-o OUTDIR]

Required arguments:
  proc                            Procedure (f/h/m)
                                    f: Feedbacks
                                    h: Hours
                                    m: Funds (money)

Optional arguments:
  -m (1-12), --month (1-12)       Month number (1-12). Default current month.
  -y YEAR  , --year YEAR          Year number. Default current year or last year.
                                    See below for explanation.
  -i INDIR , --indir INDIR        Folder with MRF Excel files. Default current
                                    folder.
  -o OUTDIR, --outdir OUTDIR      Folder to put the summary in. Default current
                                    folder.
```

### Default Year

The default year is determined using the following critera:

* If the *current* month is **Mar.** through **Dec.**, the default year is the **current** year.
* If the *current* month *and* the *requested* month are *both* **Jan.** or **Feb.**, the default year is the **current** year.
* If the *current* month is **Jan.** or **Feb.** *but* the *requested* month is **Mar.** through **Dec.**, the default year is the **last** year.

This mechanism is implemented such that in most cases, the generated report is for a month in the past and in the same CKI year.
