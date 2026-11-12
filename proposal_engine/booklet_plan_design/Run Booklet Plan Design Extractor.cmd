@echo off
setlocal

set "SCRIPT=%~dp0booklet_plan_design_extractor.py"
set "PYTHON_EXE=%BOOKLET_PYTHON_EXE%"

if not "%PYTHON_EXE%"=="" (
    if exist "%PYTHON_EXE%" goto run
)

set "CODEX_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%CODEX_PYTHON%" (
    set "PYTHON_EXE=%CODEX_PYTHON%"
    goto run
)

where python.exe >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_EXE=python.exe"
    goto run
)

where py.exe >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_EXE=py.exe"
    goto run
)

echo Python was not found.
echo Install Python 3, or set BOOKLET_PYTHON_EXE to a Python executable with openpyxl and pypdf available.
exit /b 1

:run
"%PYTHON_EXE%" -c "import cryptography" >nul 2>nul
if errorlevel 1 (
    echo Installing encrypted PDF support. This is needed for some carrier booklets...
    "%PYTHON_EXE%" -m pip install cryptography
    if errorlevel 1 (
        echo Could not install encrypted PDF support.
        echo The extractor may not be able to read AES-encrypted PDFs.
        exit /b 1
    )
)

"%PYTHON_EXE%" "%SCRIPT%" %*
exit /b %ERRORLEVEL%
