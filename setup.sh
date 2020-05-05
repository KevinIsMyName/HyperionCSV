#!/bin/bash
echo -e "setup.sh assumes that you are running Python 3+ and have pip 19+ installed. If not, ctrl+c now!\n\n\n\n\n\n"
rm -rf venv
python -m venv venv
./venv/Scripts/activate
pip install -r requirements.txt
echo -e -n "\n\n\n\n\n\nRun ExcelGrabber.py with the following commands: \n"
echo -e -n "\t.begin.sh  <--Done automatically after setup.sh\n"
echo -e -n "\tpython ExcelGrabber.py <PATH_TO_ECHO_XLSX>\n"
echo -e -n "See more information in README.md\n"
