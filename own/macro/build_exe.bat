@echo off
REM Build pdf_sign_inserter.exe from pdf_sign_inserter.py
REM Requirements: Python 3.8+ with pip

echo Installing required Python packages...
pip install pymupdf pyinstaller

echo.
echo Building EXE...
pyinstaller pdf_sign_inserter.spec

echo.
if exist dist\pdf_sign_inserter.exe (
    echo SUCCESS: dist\pdf_sign_inserter.exe was created.
    echo Copy pdf_sign_inserter.exe from the dist\ folder to use it.
) else (
    echo ERROR: Build failed. Check the output above for details.
    exit /b 1
)
