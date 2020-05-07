#!/bin/bash
echo -e "posix_setup.sh assumes that you are running Python 3+ and have pip 19+ installed. If not, ctrl+c now!\n\n\n"
rm -rf venv
python3 -m venv venv
source venv/bin/activate
pip3 install -r requirements.txt
echo -e -n "\n\n\nRun HyperionCSV.py with the following commands: \n"
echo -e -n "    source venv/bin/activate\n"
echo -e -n "    python3 HyperionCSV.py PATH_TO_ECHO_XLSX\n"
echo -e -n "See more information in README.md\n"
