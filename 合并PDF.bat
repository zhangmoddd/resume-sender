@echo off
title PDF Merge Tool
setlocal

rem ---- Find a Python that has BOTH tkinter and pypdf. Order matters. ----

rem 1) A path you wrote into data\local_python.txt
rem    (machine-specific, one line, NOT committed to Git)
set "PY="
set "PYW="
set "LOCALPY="
if exist "%~dp0data\local_python.txt" set /p LOCALPY=<"%~dp0data\local_python.txt"
if defined LOCALPY if exist "%LOCALPY%" (
    "%LOCALPY%" -c "import tkinter, pypdf" >nul 2>&1
    if not errorlevel 1 (
        set "PY=%LOCALPY%"
        set "PYW=%LOCALPY%"
    )
)

rem 2) Common install locations (no usernames baked in, uses %USERPROFILE%)
for %%P in (
"%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
"%USERPROFILE%\miniconda3\python.exe"
"%USERPROFILE%\anaconda3\python.exe"
"C:\ProgramData\miniconda3\python.exe"
"C:\ProgramData\Anaconda3\python.exe"
) do (
    if not defined PY if exist "%%~P" (
        "%%~P" -c "import tkinter, pypdf" >nul 2>&1
        if not errorlevel 1 (
            set "PY=%%~P"
            set "PYW=%%~dpPpythonw.exe"
        )
    )
)

rem 3) Whatever is on PATH
if not defined PY (
    python -c "import tkinter, pypdf" >nul 2>&1
    if not errorlevel 1 (
        set "PY=python"
        set "PYW=pythonw"
    )
)

if not defined PY (
    echo.
    echo   [ERROR] No Python with tkinter + pypdf was found.
    echo.
    echo   Install one, or write its full path into data\local_python.txt
    echo   ^(one line, for example: D:\somewhere\python.exe^)
    echo.
    pause
    exit /b 1
)

rem pythonw.exe = run without an extra console window
if not exist "%PYW%" set "PYW=%PY%"

start "" "%PYW%" "%~dp0pdf_merge_gui.py" %*
exit /b 0
