@echo off
title PDF Merge Tool
setlocal

rem =====================================================================
rem  Find a Python that has BOTH tkinter and pypdf, to run the merge tool.
rem
rem  Order: paths from data\local_python.txt first (the user's own choice),
rem  then common install locations, then whatever is on PATH.
rem  Every candidate is PROBED with a real import, so an interpreter that
rem  only looks right never gets picked.
rem
rem  Want to see which interpreter it picks?  Run this file with  --probe
rem =====================================================================

set "PY="
set "PYW="
set "NEED=tkinter, pypdf"

rem ---- 1) paths the user wrote into data\local_python.txt (highest) ----
if exist "%~dp0data\local_python.txt" (
    for /f "usebackq eol=# delims=" %%L in ("%~dp0data\local_python.txt") do (
        if not defined PY (
            set "CAND=%%~L"
            call :try_cand
        )
    )
)

rem ---- 2) common install locations (no usernames baked in) ----
for %%P in (
"%USERPROFILE%\miniconda3\python.exe"
"%USERPROFILE%\anaconda3\python.exe"
"C:\ProgramData\miniconda3\python.exe"
"C:\ProgramData\Anaconda3\python.exe"
"%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
"%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
"%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
) do (
    if not defined PY if exist "%%~P" (
        set "CAND=%%~P"
        call :try_cand
    )
)

rem ---- 3) whatever is on PATH ----
if not defined PY (
    python -c "import %NEED%" >nul 2>&1
    if not errorlevel 1 (
        set "PY=python"
        set "PYW=pythonw"
    )
)
if not defined PY (
    py -c "import %NEED%" >nul 2>&1
    if not errorlevel 1 (
        set "PY=py"
        set "PYW=pyw"
    )
)

if not defined PY (
    echo.
    echo   [ERROR] No Python with tkinter + pypdf was found.
    echo.
    echo   Install one, or write its full path into data\local_python.txt
    echo   ^(one path per line, for example: D:\somewhere\python.exe^)
    echo.
    pause
    exit /b 1
)

rem pythonw.exe = run without an extra console window
if not exist "%PYW%" set "PYW=%PY%"

if /i "%~1"=="--probe" (
    echo Selected Python: %PY%
    echo Window-less exe : %PYW%
    exit /b 0
)

start "" "%PYW%" "%~dp0pdf_merge_gui.py" %*
exit /b 0


rem ---------------------------------------------------------------------
rem  Probe one candidate: must exist AND must import tkinter + pypdf.
rem  Reads %CAND%, sets %PY% / %PYW% on success, returns quietly otherwise.
rem ---------------------------------------------------------------------
:try_cand
if not exist "%CAND%" goto :eof
"%CAND%" -c "import %NEED%" >nul 2>&1
if errorlevel 1 goto :eof
set "PY=%CAND%"
for %%D in ("%CAND%") do set "PYW=%%~dpDpythonw.exe"
goto :eof
