@echo off
REM DOCX to Asana CSV Generator Installer
REM This batch file installs the DOCX to Asana CSV Generator application and its dependencies.

echo === DOCX to Asana CSV Generator Installer ===
echo.

REM Set installation directory
set INSTALL_DIR=%USERPROFILE%\DOCX-to-Asana

REM Check if Python is installed
python --version > nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Python is not installed or not in PATH.
    echo Please install Python 3.8 or higher from https://www.python.org/downloads/
    echo After installing Python, run this installer again.
    goto :exit
)

REM Check Python version
for /f "tokens=2" %%V in ('python --version 2^>^&1') do set PYTHON_VERSION=%%V
for /f "tokens=1,2 delims=." %%a in ("%PYTHON_VERSION%") do (
    set PYTHON_MAJOR=%%a
    set PYTHON_MINOR=%%b
)

if %PYTHON_MAJOR% LSS 3 (
    echo Python %PYTHON_VERSION% found, but version 3.8 or higher is required.
    goto :exit
) else (
    if %PYTHON_MAJOR% EQU 3 (
        if %PYTHON_MINOR% LSS 8 (
            echo Python %PYTHON_VERSION% found, but version 3.8 or higher is required.
            goto :exit
        )
    )
)

echo Python %PYTHON_VERSION% found. Requirement satisfied.
echo.

REM Create installation directory if it doesn't exist
if not exist "%INSTALL_DIR%" (
    echo Creating installation directory: %INSTALL_DIR%
    mkdir "%INSTALL_DIR%"
)

REM Determine source directory
set SOURCE_DIR=%~dp0..
echo Using source directory: %SOURCE_DIR%

REM Verify source directory contains required files
if not exist "%SOURCE_DIR%\src" (
    echo Warning: Source directory does not contain 'src' folder. Installation may be incomplete.
)

echo Copying application files from %SOURCE_DIR% to %INSTALL_DIR%
echo.

REM Create necessary directories
mkdir "%INSTALL_DIR%\src" 2>nul
mkdir "%INSTALL_DIR%\src\utils" 2>nul
mkdir "%INSTALL_DIR%\tests" 2>nul
mkdir "%INSTALL_DIR%\tests\sample_documents" 2>nul

REM Copy files
set COPY_ERRORS=0

if exist "%SOURCE_DIR%\requirements.txt" (
    copy "%SOURCE_DIR%\requirements.txt" "%INSTALL_DIR%\" > nul
    echo Copied: requirements.txt
) else (
    echo Warning: Source file not found: %SOURCE_DIR%\requirements.txt
    set /a COPY_ERRORS+=1
)

if exist "%SOURCE_DIR%\config.json" (
    copy "%SOURCE_DIR%\config.json" "%INSTALL_DIR%\" > nul
    echo Copied: config.json
) else (
    echo Warning: Source file not found: %SOURCE_DIR%\config.json
    set /a COPY_ERRORS+=1
)

if exist "%SOURCE_DIR%\README.md" (
    copy "%SOURCE_DIR%\README.md" "%INSTALL_DIR%\" > nul
    echo Copied: README.md
) else (
    echo Warning: Source file not found: %SOURCE_DIR%\README.md
    set /a COPY_ERRORS+=1
)

REM Copy source files
set SRC_FILES=docx_to_asana.py document_parser.py csv_generator.py gui_handler.py
for %%F in (%SRC_FILES%) do (
    if exist "%SOURCE_DIR%\src\%%F" (
        copy "%SOURCE_DIR%\src\%%F" "%INSTALL_DIR%\src\" > nul
        echo Copied: src\%%F
    ) else (
        echo Warning: Source file not found: %SOURCE_DIR%\src\%%F
        set /a COPY_ERRORS+=1
    )
)

REM Copy utility files
set UTIL_FILES=logger.py config.py validator.py
for %%F in (%UTIL_FILES%) do (
    if exist "%SOURCE_DIR%\src\utils\%%F" (
        copy "%SOURCE_DIR%\src\utils\%%F" "%INSTALL_DIR%\src\utils\" > nul
        echo Copied: src\utils\%%F
    ) else (
        echo Warning: Source file not found: %SOURCE_DIR%\src\utils\%%F
        set /a COPY_ERRORS+=1
    )
)

REM Copy test files
set TEST_FILES=test_parser.py test_csv_generator.py
for %%F in (%TEST_FILES%) do (
    if exist "%SOURCE_DIR%\tests\%%F" (
        copy "%SOURCE_DIR%\tests\%%F" "%INSTALL_DIR%\tests\" > nul
        echo Copied: tests\%%F
    ) else (
        echo Warning: Source file not found: %SOURCE_DIR%\tests\%%F
        set /a COPY_ERRORS+=1
    )
)

if %COPY_ERRORS% GTR 0 (
    echo Warning: %COPY_ERRORS% files could not be copied. Installation may be incomplete.
)

REM Create requirements.txt if it doesn't exist
if not exist "%INSTALL_DIR%\requirements.txt" (
    echo Creating default requirements.txt file
    (
        echo python-docx^>=0.8.11
        echo pandas^>=1.3.0
        echo openpyxl^>=3.0.7
        echo pytest^>=6.2.5
    ) > "%INSTALL_DIR%\requirements.txt"
)

REM Install dependencies
echo.
echo Installing Python dependencies...
python -m pip install --upgrade pip
python -m pip install -r "%INSTALL_DIR%\requirements.txt"
if %ERRORLEVEL% neq 0 (
    echo Warning: Some dependencies may not have been installed correctly.
) else (
    echo Dependencies installed successfully.
)

REM Create desktop shortcut
echo.
echo Creating desktop shortcut...
set SHORTCUT_PATH=%USERPROFILE%\Desktop\DOCX to Asana.lnk
powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%SHORTCUT_PATH%'); $Shortcut.TargetPath = 'pythonw.exe'; $Shortcut.Arguments = '\"%INSTALL_DIR:\=\\%\\src\\docx_to_asana.py\" --gui'; $Shortcut.WorkingDirectory = '%INSTALL_DIR:\=\\%'; $Shortcut.IconLocation = 'shell32.dll,23'; $Shortcut.Description = 'DOCX to Asana CSV Generator'; $Shortcut.Save()"
if %ERRORLEVEL% neq 0 (
    echo Warning: Could not create desktop shortcut.
    echo You can still run the application manually.
) else (
    echo Desktop shortcut created: %SHORTCUT_PATH%
)

echo.
echo Installation completed successfully!
echo You can run the application using the desktop shortcut or by running:
echo python -m src.docx_to_asana --gui
echo.

:exit
echo Press any key to exit...
pause > nul
