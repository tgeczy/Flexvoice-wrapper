@echo off
echo Building FlexVoice 32-bit host executable...
py -3.14-32 -m PyInstaller --onefile --noconsole --name flexvoice_host32 synthDrivers\host_flexvoice32.py
echo.
echo Done. Output: dist\flexvoice_host32.exe
echo Copy to synthDrivers\ before running build.py:
echo   copy dist\flexvoice_host32.exe synthDrivers\
