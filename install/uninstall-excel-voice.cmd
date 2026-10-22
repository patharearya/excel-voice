@echo off
rem Removes Excel Voice for the current Windows user.
setlocal
set "DIR=%LOCALAPPDATA%\ExcelVoice"
set "ID=b873a624-7ccd-4833-acff-e1bd5c27d5c9"

reg delete "HKCU\Software\Microsoft\Office\16.0\Wef\Developer" /v "%ID%" /f >nul 2>&1
if exist "%DIR%" rmdir /s /q "%DIR%"
echo Excel Voice is removed. Close and reopen Excel to finish.
pause
