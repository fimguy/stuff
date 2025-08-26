@echo off
echo === AAD Connect XML Parser for Power BI ===
echo.
echo This script will parse your AAD Connect XML file and create Power BI ready data files.
echo Make sure today-RunStep.xml is in the same folder as this batch file.
echo.
pause

echo Starting PowerShell script...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Run-XMLParser-Windows.ps1"

echo.
echo Script completed. Check the folder for generated CSV files.
pause
