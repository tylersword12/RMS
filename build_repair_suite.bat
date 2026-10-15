@echo off
echo ============================================
echo  Repair Management Suite - EXE Builder
echo  (Using Virtual Environment)
echo ============================================
echo.

:: Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

:: Clean previous builds
echo Cleaning previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

:: Create fresh virtual environment
echo.
echo Creating fresh virtual environment...
if exist "build_venv" rmdir /s /q build_venv
python -m venv build_venv

:: Activate virtual environment
call build_venv\Scripts\activate.bat

:: Upgrade pip
echo.
echo Upgrading pip...
python -m pip install --upgrade pip

:: Install all dependencies
echo.
echo Installing dependencies...

echo   Installing pyinstaller...
pip install pyinstaller
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing Pillow...
pip install Pillow
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing python-barcode...
pip install python-barcode
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing brother_ql...
pip install brother_ql
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing pyusb...
pip install pyusb
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing libusb1 (provides libusb DLLs)...
pip install libusb1
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing libusb-package (bundled libusb DLLs)...
pip install libusb-package
if errorlevel 1 ( echo Continuing without libusb-package )

echo   Installing openpyxl...
pip install openpyxl
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing reportlab...
pip install reportlab
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing xlrd...
pip install xlrd
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing opencv-python (webcam support)...
pip install opencv-python
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing pyzbar (barcode decoding)...
pip install pyzbar
if errorlevel 1 ( echo FAILED & pause & exit /b 1 )

echo   Installing pywin32 (Outlook email integration)...
pip install pywin32
if errorlevel 1 ( echo Continuing without pywin32 - Outlook will use command-line fallback )

:: Verify all imports work
echo.
echo Verifying all imports...
python -c "from brother_ql.conversion import convert; print('  OK: brother_ql')"
python -c "from PIL import Image, ImageDraw, ImageFont; print('  OK: Pillow')"
python -c "import barcode; from barcode.writer import ImageWriter; print('  OK: python-barcode')"
python -c "import usb.core; print('  OK: pyusb')"
python -c "import openpyxl; print('  OK: openpyxl')"
python -c "from reportlab.pdfgen import canvas; print('  OK: reportlab')"
python -c "import xlrd; print('  OK: xlrd')"
python -c "import cv2; print('  OK: opencv-python')"
python -c "from pyzbar import pyzbar; print('  OK: pyzbar')"
python -c "import win32com.client; print('  OK: pywin32')" 2>nul || echo   WARN: pywin32 missing - Outlook attach falls back to command line

:: Locate package paths
echo.
echo Locating package data files...
for /f "delims=" %%i in ('python -c "import brother_ql; import os; print(os.path.dirname(brother_ql.__file__))"') do set BQL_PATH=%%i
echo   brother_ql at: %BQL_PATH%

for /f "delims=" %%i in ('python -c "import barcode; import os; print(os.path.dirname(barcode.__file__))"') do set BARCODE_PATH=%%i
echo   barcode at: %BARCODE_PATH%

:: Locate libusb DLLs
echo.
echo Locating libusb DLLs...
for /f "delims=" %%i in ('python -c "import libusb_package; import os; print(os.path.dirname(libusb_package.__file__))" 2^>nul') do set LIBUSB_PKG_PATH=%%i
if defined LIBUSB_PKG_PATH (
    echo   libusb-package at: %LIBUSB_PKG_PATH%
)

for /f "delims=" %%i in ('python -c "import usb1; import os; print(os.path.dirname(usb1.__file__))" 2^>nul') do set USB1_PATH=%%i
if defined USB1_PATH (
    echo   libusb1 at: %USB1_PATH%
)

:: Locate pyzbar DLLs (required for Windows barcode decoding)
for /f "delims=" %%i in ('python -c "import pyzbar; import os; print(os.path.dirname(pyzbar.__file__))" 2^>nul') do set PYZBAR_PATH=%%i
if defined PYZBAR_PATH (
    echo   pyzbar at: %PYZBAR_PATH%
)

:: Locate cv2 DLLs
for /f "delims=" %%i in ('python -c "import cv2; import os; print(os.path.dirname(cv2.__file__))" 2^>nul') do set CV2_PATH=%%i
if defined CV2_PATH (
    echo   cv2 at: %CV2_PATH%
)

echo.
echo Building EXE (this may take a couple minutes)...
echo.

pyinstaller --onefile --noconsole --name "RepairManagementSuite" ^
    --hidden-import=brother_ql ^
    --hidden-import=brother_ql.conversion ^
    --hidden-import=brother_ql.backends ^
    --hidden-import=brother_ql.backends.helpers ^
    --hidden-import=brother_ql.backends.pyusb ^
    --hidden-import=brother_ql.raster ^
    --hidden-import=brother_ql.devicedependent ^
    --hidden-import=brother_ql.models ^
    --hidden-import=brother_ql.labels ^
    --hidden-import=PIL ^
    --hidden-import=PIL.Image ^
    --hidden-import=PIL.ImageDraw ^
    --hidden-import=PIL.ImageFont ^
    --hidden-import=PIL._imaging ^
    --hidden-import=barcode ^
    --hidden-import=barcode.codex ^
    --hidden-import=barcode.code128 ^
    --hidden-import=barcode.writer ^
    --hidden-import=barcode.base ^
    --hidden-import=barcode.charsets ^
    --hidden-import=barcode.charsets.code128 ^
    --hidden-import=usb ^
    --hidden-import=usb.core ^
    --hidden-import=usb.util ^
    --hidden-import=usb.backend ^
    --hidden-import=usb.backend.libusb1 ^
    --hidden-import=usb.backend.libusb0 ^
    --hidden-import=usb.backend.openusb ^
    --hidden-import=usb._debug ^
    --hidden-import=usb._interop ^
    --hidden-import=usb._lookup ^
    --hidden-import=usb._objfinalizer ^
    --hidden-import=openpyxl ^
    --hidden-import=reportlab ^
    --hidden-import=reportlab.pdfgen ^
    --hidden-import=reportlab.pdfgen.canvas ^
    --hidden-import=reportlab.lib ^
    --hidden-import=reportlab.lib.pagesizes ^
    --hidden-import=reportlab.lib.colors ^
    --hidden-import=reportlab.lib.units ^
    --hidden-import=reportlab.platypus ^
    --hidden-import=xlrd ^
    --hidden-import=cv2 ^
    --hidden-import=pyzbar ^
    --hidden-import=pyzbar.pyzbar ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    --hidden-import=win32api ^
    --hidden-import=libusb_package ^
    --hidden-import=usb1 ^
    --collect-all=brother_ql ^
    --collect-all=barcode ^
    --collect-all=usb ^
    --collect-all=reportlab ^
    --collect-all=libusb_package ^
    --collect-all=usb1 ^
    --collect-all=cv2 ^
    --collect-all=pyzbar ^
    --collect-all=win32com ^
    --collect-submodules=brother_ql ^
    --collect-submodules=barcode ^
    --collect-submodules=usb ^
    --collect-submodules=PIL ^
    --collect-submodules=reportlab ^
    --collect-submodules=xlrd ^
    --collect-data=brother_ql ^
    --collect-data=barcode ^
    --collect-data=reportlab ^
    --add-data="%BQL_PATH%;brother_ql" ^
    --add-data="%BARCODE_PATH%;barcode" ^
    --add-data="%PYZBAR_PATH%;pyzbar" ^
    --add-data="%CV2_PATH%;cv2" ^
    repair_management_suite.py

:: Deactivate virtual environment
call build_venv\Scripts\deactivate.bat 2>nul

echo.
if exist "dist\RepairManagementSuite.exe" (
    echo ============================================
    echo  SUCCESS! EXE built at:
    echo  dist\RepairManagementSuite.exe
    echo ============================================
    echo.
    for %%A in ("dist\RepairManagementSuite.exe") do echo   File size: %%~zA bytes
    echo.
    echo You can move RepairManagementSuite.exe anywhere.
    echo No Python or other installs needed to run it.
    echo.
    echo Optional: Delete the build_venv folder to free space.
) else (
    echo ============================================
    echo  BUILD FAILED - Check errors above
    echo ============================================
)

echo.
pause
