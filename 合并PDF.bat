@echo off
title PDF Merge Tool
setlocal

rem ---- Find a Python that has BOTH tkinter and pypdf. Order matters. ----
set "PY="
set "PYW="
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
    echo   Do NOT close this window. Take a screenshot and send it to the assistant.
    echo.
    pause
    exit /b 1
)

rem pythonw.exe = run without an extra console window
if not exist "%PYW%" set "PYW=%PY%"

start "" "%PYW%" "%~dp0pdf_merge_gui.py" %*
exit /b 0
