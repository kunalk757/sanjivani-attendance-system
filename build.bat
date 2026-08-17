@echo off
setlocal
echo ======================================================================
echo   SANJIVANI ATTENDANCE SYSTEM - BUILD STANDALONE EXECUTABLE
echo ======================================================================
echo.

where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not found in PATH.
    echo Please install Python 3.10+ and add it to PATH.
    pause
    exit /b 1
)

echo [1/3] Installing / Verifying requirements...
python -m pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo [2/4] Generating application icon (sanjivani.ico)...
python create_icon.py

echo.
echo [3/4] Building standalone executable with PyInstaller...
python -m PyInstaller --clean --noconfirm app.spec
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo.
echo [4/4] Building Inno Setup Installer...
set "ISCC_EXE="
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if exist "%LOCALAPPDATA%\Programs\Inno Setup 7\ISCC.exe" set "ISCC_EXE=%LOCALAPPDATA%\Programs\Inno Setup 7\ISCC.exe"
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if defined ISCC_EXE (
    "%ISCC_EXE%" installer.iss
    if %ERRORLEVEL% neq 0 (
        echo [WARNING] Installer compilation failed.
    ) else (
        echo [OK] Installer built successfully.
    )
) else (
    echo [NOTE] ISCC.exe not found in standard paths; skipping installer.
)

echo.
echo ======================================================================
echo   BUILD SUCCESSFUL!
echo   Output binary: dist\Sanjivani Attendance.exe
echo   Installer    : dist\Sanjivani-Attendance-Setup.exe
echo ======================================================================
echo.
pause

