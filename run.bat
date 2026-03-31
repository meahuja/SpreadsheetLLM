@echo off
set PY=C:\Users\vatsal.gaur\AppData\Local\Programs\Python\Python314\python.exe
set OUT=C:\Users\vatsal.gaur\Desktop\SpreadsheetLLm\test_results.txt

echo START > %OUT%
%PY% --version >> %OUT% 2>&1
if errorlevel 1 (
    echo PYTHON NOT FOUND >> %OUT%
    goto :eof
)

cd /d C:\Users\vatsal.gaur\Desktop\SpreadsheetLLm
%PY% -m pip install openpyxl pytest >> %OUT% 2>&1
%PY% -m pytest tests/ -v --tb=short >> %OUT% 2>&1
echo EXIT_CODE=%errorlevel% >> %OUT%
