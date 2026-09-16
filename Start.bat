@echo off
title Resume Batch Sender
setlocal

rem ---- Pick a Python that has openpyxl (used for Excel import/export) ----
rem Try the per-user environment first, then fall back to whatever is on PATH.
set "PY=%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" "%~dp0app.py"
pause
