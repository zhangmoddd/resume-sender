@echo off
title Resume Batch Sender
setlocal

rem =====================================================================
rem  Pick a Python to run the main app.
rem
rem  Two rules, in this order:
rem
rem   1) CAPABILITY, not existence. A candidate is accepted only if it can
rem      really import what the app needs. "The file exists" is not enough:
rem      a Python without openpyxl starts the app fine and then dies the
rem      moment you click "export Excel".
rem
rem   2) THE USER'S OWN CHOICE WINS. Paths listed in data\local_python.txt
rem      are tried first. That file is machine specific and git-ignored,
rem      so nothing in here is tied to one particular PC.
rem
rem  Want to see which interpreter it picks?  Run:  Start.bat --probe
rem =====================================================================

set "PY="
set "FB="
rem openpyxl is mandatory (see README); pypdf + pillow only matter for merging
set "NEED=openpyxl, pypdf, PIL"

rem ---- 1) paths the user wrote into data\local_python.txt (highest) ----
if exist "%~dp0data\local_python.txt" (
    for /f "usebackq eol=# delims=" %%L in ("%~dp0data\local_python.txt") do (
        if not defined PY (
            set "CAND=%%~L"
            call :try_cand
        )
    )
)

rem ---- 2) a virtualenv sitting inside the project folder ----
if not defined PY if exist "%~dp0.venv\Scripts\python.exe" (
    set "CAND=%~dp0.venv\Scripts\python.exe"
    call :try_cand
)

rem ---- 3) whatever is on PATH, then the official Windows launcher ----
if not defined PY (
    python -c "import %NEED%" >nul 2>&1
    if not errorlevel 1 set PY="python"
)
if not defined PY (
    py -c "import %NEED%" >nul 2>&1
    if not errorlevel 1 set PY="py"
)

rem ---- 4) common install locations (no usernames baked in) ----
for %%P in (
"C:\ProgramData\miniconda3\python.exe"
"C:\ProgramData\Anaconda3\python.exe"
"%USERPROFILE%\miniconda3\python.exe"
"%USERPROFILE%\anaconda3\python.exe"
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

rem ---- 5) last resort: any Python that runs at all ----
rem  Sending mail only needs the standard library, so the app is still
rem  usable. Excel and PDF merging will not be. Say so, loudly.
if not defined PY (
    python -c "import sys" >nul 2>&1
    if not errorlevel 1 (
        set PY="python"
        set "FB=1"
    )
)
if not defined PY (
    py -c "import sys" >nul 2>&1
    if not errorlevel 1 (
        set PY="py"
        set "FB=1"
    )
)

if not defined PY (
    echo.
    echo   [ERROR] No Python was found on this computer.
    echo.
    echo   1. Install Python 3.10 or newer -^> https://www.python.org/downloads/
    echo      ^(during setup, tick "Add python.exe to PATH"^)
    echo   2. Then run:   pip install openpyxl pypdf pillow
    echo   3. Or write your Python's full path into:  data\local_python.txt
    echo.
    pause
    goto :eof
)

if defined FB (
    echo.
    echo   [WARN] No Python with openpyxl was found. Excel import / export is OFF.
    echo   [WARN] Fix it with:   pip install openpyxl pypdf pillow
    echo.
)

if /i "%~1"=="--probe" (
    echo Selected Python: %PY%
    goto :eof
)

echo   Python: %PY%
%PY% "%~dp0app.py"
pause
goto :eof


rem ---------------------------------------------------------------------
rem  Probe one candidate: must exist AND must import what the app needs.
rem  Reads %CAND%, sets %PY% on success, returns quietly on failure.
rem ---------------------------------------------------------------------
:try_cand
if not exist "%CAND%" goto :eof
"%CAND%" -c "import %NEED%" >nul 2>&1
if errorlevel 1 goto :eof
set PY="%CAND%"
goto :eof
