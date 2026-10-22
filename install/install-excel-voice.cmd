@echo off
rem Installs Excel Voice for the current Windows user. No admin rights needed.
rem Downloads the add-in manifest and tells Excel where it is (the same
rem registry entry `npm start` uses to load an add-in during development).
setlocal
set "DIR=%LOCALAPPDATA%\ExcelVoice"
set "ID=b873a624-7ccd-4833-acff-e1bd5c27d5c9"
set "URL=https://patharearya.github.io/excel-voice/manifest.xml"

echo Installing Excel Voice...
if not exist "%DIR%" mkdir "%DIR%"
curl -fsSL -o "%DIR%\manifest.xml" "%URL%"
if errorlevel 1 (
  echo.
  echo Could not download Excel Voice. Check your internet connection and try again.
  pause
  exit /b 1
)
reg add "HKCU\Software\Microsoft\Office\16.0\Wef\Developer" /v "%ID%" /t REG_SZ /d "%DIR%\manifest.xml" /f >nul
if errorlevel 1 (
  echo.
  echo Could not register Excel Voice with Excel.
  pause
  exit /b 1
)
echo.
echo Excel Voice is installed.
echo.
echo  1. Close Excel if it is open, then open it again.
echo  2. Open any workbook.
echo  3. On the Home tab, click "Excel Voice" at the far right.
echo.
pause
