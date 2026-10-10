@echo off
setlocal enabledelayedexpansion

echo ======================================================================
echo    HOSTEL REGISTER GENERATOR - WINDOWS EXE BUILD SCRIPT
echo ======================================================================
echo.

:: 1. Check Python installation
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    where py >nul 2>nul
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Python was not found in your PATH!
        echo Please install Python 3.8 or higher from https://www.python.org/
        echo Make sure to check "Add Python to PATH" during installation.
        echo.
        pause
        exit /b 1
    ) else (
        set PYTHON_CMD=py -3
    )
) else (
    set PYTHON_CMD=python
)

echo [*] Using Python command: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

:: 2. Create or verify virtual environment
set VENV_DIR=.venv_build
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo [*] Creating build virtual environment in "%VENV_DIR%"...
    %PYTHON_CMD% -m venv %VENV_DIR%
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
) else (
    echo [*] Reusing existing virtual environment in "%VENV_DIR%".
)

:: 3. Activate virtual environment
echo [*] Activating virtual environment...
call "%VENV_DIR%\Scripts\activate.bat"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to activate virtual environment.
    pause
    exit /b 1
)

:: 4. Install dependencies
echo [*] Installing runtime and build dependencies...
python -m pip install --upgrade pip --quiet
if exist requirements-build.txt (
    python -m pip install -r requirements-build.txt --quiet
) else (
    python -m pip install -r requirements.txt pyinstaller --quiet
)
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install required dependencies.
    pause
    exit /b 1
)
echo [OK] Dependencies installed successfully.
echo.

:: 5. Clean prior build artifacts
if exist build rmdir /s /q build
if not exist dist mkdir dist

:: 6. Build the executable with PyInstaller
echo [*] Building standalone Windows executable with PyInstaller...
if exist HostelRegisterGenerator.spec (
    echo [*] Using HostelRegisterGenerator.spec configuration...
    pyinstaller --clean -y HostelRegisterGenerator.spec
) else (
    pyinstaller --clean -y ^
        --name "HostelRegisterGenerator" ^
        --onefile ^
        --windowed ^
        --collect-all openpyxl ^
        main.py
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ======================================================================
    echo [ERROR] PyInstaller build failed! Please check the logs above.
    echo ======================================================================
    pause
    exit /b 1
)

:: 7. Verify the output
echo.
if exist "dist\HostelRegisterGenerator.exe" (
    echo ======================================================================
    echo [SUCCESS] Standalone executable created successfully!
    echo.
    echo  Location: dist\HostelRegisterGenerator.exe
    echo.
    echo  You can now copy "dist\HostelRegisterGenerator.exe" to any Windows
    echo  computer and double-click to run it without installing Python.
    echo ======================================================================
) else (
    echo [ERROR] Expected output "dist\HostelRegisterGenerator.exe" was not found.
    pause
    exit /b 1
)

echo.
pause
