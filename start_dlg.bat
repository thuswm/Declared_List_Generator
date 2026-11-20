@echo off
rem Start the Declared List Generator: double-click this file, or point a
rem desktop shortcut at it. The browser opens by itself; stop the tool with
rem the Exit button in the browser, Ctrl+C here, or by closing this window.
rem Extra arguments are passed on, e.g.  start_dlg.bat --port 8001
setlocal
rem Run from the project folder, wherever the shortcut or double-click came from.
cd /d "%~dp0"

set "DLG_PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%DLG_PYTHON%" goto :no_venv

title Declared List Generator
"%DLG_PYTHON%" -m dlg %*
if errorlevel 1 goto :failed
exit /b 0

:no_venv
echo The virtual environment .venv was not found in:
echo   %~dp0
echo.
echo Set it up once from this folder (see README.md, "Setup"):
echo   py -3 -m venv .venv
echo   .venv\Scripts\pip install -e .
echo.
pause
exit /b 1

:failed
rem Keep the window open so a double-click user can read the error.
echo.
echo The Declared List Generator stopped with an error (see above).
pause
exit /b 1
