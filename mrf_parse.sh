#!/bin/bash
if command -v python3 >/dev/null 2>&1
then
    python3 mrf_parse
elif command -v python >/dev/null 2>&1 && [[ "$(python -V 2>&1)" =~ [[:space:]]3\. ]]
then
    python mrf_parse
else
    echo "Python 3 not installed"
    echo "Install latest Python 3 release from https://www.python.org/downloads/"
fi
read -n 1 -s -r -p "Press any key to continue . . ."
echo
