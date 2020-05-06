@echo off
echo win_setup.sh assumes that you are running Python 3+ and have pip 19+ installed. If not, ctrl+c now!
echo Run HyperionCSV.py with the following commands:
echo .\venv\Scripts\activate
echo python HyperionCSV.py PATH_TO_ECHO_XLSX
echo See more information in README.md
rd /s /q venv
python3 -m pip install venv
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt