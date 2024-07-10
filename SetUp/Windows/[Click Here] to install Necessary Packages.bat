@echo off
rem Install required packages using pip

rem Ensure pip is up-to-date
python -m pip install --upgrade pip

rem Install specific versions of packages
python -m pip install openpyxl==3.1.4
python -m pip install pandas==2.2.2
python -m pip install ttkthemes==3.2.2
python -m pip install Pillow==10.4.0

rem Optionally, install python-dotenv if needed
rem python -m pip install python-dotenv

rem Pause to see the output
pause
