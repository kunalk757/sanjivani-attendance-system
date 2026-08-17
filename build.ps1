# ======================================================================
# Sanjivani Attendance System - PowerShell Build Script
# ======================================================================

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "  SANJIVANI ATTENDANCE SYSTEM - BUILD STANDALONE EXECUTABLE" -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Check Python
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    Write-Host "[ERROR] Python was not found in PATH." -ForegroundColor Red
    Write-Host "Please install Python 3.10+ and ensure it is added to your PATH." -ForegroundColor Yellow
    exit 1
}

# 2. Check / Install dependencies
Write-Host "[1/4] Checking dependencies from requirements.txt..." -ForegroundColor Yellow
python -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Failed to install dependencies." -ForegroundColor Red
    exit 1
}

# 3. Generate icon
Write-Host "[2/5] Generating and verifying icon assets..." -ForegroundColor Yellow
python create_icon.py

# 4. Run PyInstaller build
Write-Host "[3/5] Running PyInstaller compilation..." -ForegroundColor Yellow
python -m PyInstaller --clean --noconfirm app.spec
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Compilation failed." -ForegroundColor Red
    exit 1
}

# 5. Build Installer with Inno Setup
Write-Host "[4/5] Building Windows Setup Installer..." -ForegroundColor Yellow
$isccCandidates = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 7\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
)

$isccPath = $null
foreach ($p in $isccCandidates) {
    if (Test-Path $p) {
        $isccPath = $p
        break
    }
}
if (-not $isccPath) {
    $cmd = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($cmd) { $isccPath = $cmd.Source }
}

if ($isccPath) {
    & $isccPath ".\installer.iss"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[WARNING] Inno Setup compilation exited with code $LASTEXITCODE" -ForegroundColor Yellow
    } else {
        Write-Host "  [OK] Inno Setup Installer created successfully." -ForegroundColor Green
    }
} else {
    Write-Host "  [NOTE] Inno Setup compiler (ISCC.exe) not found; skipping installer creation." -ForegroundColor Yellow
}

# 6. Output check
Write-Host "[5/5] Checking generated binaries..." -ForegroundColor Yellow
$exePath = ".\dist\Sanjivani Attendance.exe"
$setupPath = ".\dist\Sanjivani-Attendance-Setup.exe"

if (Test-Path $exePath) {
    $fileItem = Get-Item $exePath
    $sizeMB = [math]::Round($fileItem.Length / 1MB, 2)
    Write-Host ""
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host "  BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host "  Application Executable : $exePath ($sizeMB MB)" -ForegroundColor Green
    if (Test-Path $setupPath) {
        $setupItem = Get-Item $setupPath
        $setupSizeMB = [math]::Round($setupItem.Length / 1MB, 2)
        Write-Host "  Setup Installer        : $setupPath ($setupSizeMB MB)" -ForegroundColor Green
    }
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host ""
} else {
    Write-Host "[ERROR] Target executable was not found at $exePath." -ForegroundColor Red
    exit 1
}

