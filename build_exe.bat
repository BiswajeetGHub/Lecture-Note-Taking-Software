@echo off
title Building Lecnotes
cd /d "%~dp0"

echo ==========================================
echo   Building Lecnotes.exe  (takes 1-3 minutes)
echo ==========================================
echo.

rem --- find Python: try "python" first, then "py"
set PY=python
where python >nul 2>nul
if errorlevel 1 set PY=py
%PY% --version >nul 2>nul
if errorlevel 1 goto nopython

echo Step 1 of 2: installing what the app needs...
%PY% -m pip install customtkinter mss keyboard python-pptx python-docx pillow pyinstaller
if errorlevel 1 goto failed

echo.
echo Step 2 of 2: creating the exe...
%PY% -m PyInstaller --onefile --noconsole --name Lecnotes --collect-all customtkinter --collect-data pptx --collect-data docx lecture_capture.py
if errorlevel 1 goto failed

copy /Y "dist\Lecnotes.exe" "Lecnotes.exe" >nul
rmdir /S /Q build 2>nul
rmdir /S /Q dist 2>nul
del /Q Lecnotes.spec 2>nul

echo.
echo ==========================================
echo   Done! Lecnotes.exe is in this folder.
echo ==========================================
explorer /select,"%~dp0Lecnotes.exe"
pause
exit /b 0

:nopython
echo Python was not found.
echo Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH", then run this file again.
pause
exit /b 1

:failed
echo.
echo Something went wrong. Scroll up to read the error message.
pause
exit /b 1
