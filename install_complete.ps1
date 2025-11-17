# ========================================
# BADR LTA Automation - Complete Installation Script
# ========================================
# Run with: .\install_complete.ps1
# Or: powershell -ExecutionPolicy Bypass -File .\install_complete.ps1

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  BADR LTA Automation - Installation" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Function to check if running as Administrator
function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Check for admin rights (recommended but not required)
if (-not (Test-Administrator)) {
    Write-Host "⚠️  Note: Not running as Administrator" -ForegroundColor Yellow
    Write-Host "   Installation will continue, but some features may be limited" -ForegroundColor Yellow
    Write-Host ""
}

# Step 1: Check Python Installation
Write-Host "🔍 Step 1: Checking Python installation..." -ForegroundColor Green
Write-Host ""

try {
    $pythonVersion = python --version 2>&1
    Write-Host "   ✅ Python found: $pythonVersion" -ForegroundColor Green
    
    # Extract version number
    if ($pythonVersion -match "Python (\d+)\.(\d+)\.(\d+)") {
        $major = [int]$matches[1]
        $minor = [int]$matches[2]
        
        if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 10)) {
            Write-Host "   ⚠️  Python version is too old. Please upgrade to Python 3.10 or higher." -ForegroundColor Red
            Write-Host "   Download from: https://www.python.org/downloads/" -ForegroundColor Yellow
            exit 1
        }
    }
} catch {
    Write-Host "   ❌ Python not found!" -ForegroundColor Red
    Write-Host "   Please install Python 3.10 or higher from https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host "   IMPORTANT: Check 'Add Python to PATH' during installation!" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Step 2: Check pip
Write-Host "🔍 Step 2: Checking pip..." -ForegroundColor Green
Write-Host ""

try {
    $pipVersion = pip --version 2>&1
    Write-Host "   ✅ pip found: $pipVersion" -ForegroundColor Green
} catch {
    Write-Host "   ❌ pip not found!" -ForegroundColor Red
    Write-Host "   Installing pip..." -ForegroundColor Yellow
    
    try {
        python -m ensurepip --upgrade
        Write-Host "   ✅ pip installed successfully" -ForegroundColor Green
    } catch {
        Write-Host "   ❌ Failed to install pip. Please install manually." -ForegroundColor Red
        exit 1
    }
}

Write-Host ""

# Step 3: Upgrade pip
Write-Host "📦 Step 3: Upgrading pip to latest version..." -ForegroundColor Green
Write-Host ""

try {
    python -m pip install --upgrade pip
    Write-Host "   ✅ pip upgraded successfully" -ForegroundColor Green
} catch {
    Write-Host "   ⚠️  Could not upgrade pip, continuing with current version" -ForegroundColor Yellow
}

Write-Host ""

# Step 4: Check if requirements.txt exists
Write-Host "🔍 Step 4: Checking requirements.txt..." -ForegroundColor Green
Write-Host ""

if (-not (Test-Path "requirements.txt")) {
    Write-Host "   ❌ requirements.txt not found in current directory!" -ForegroundColor Red
    Write-Host "   Current directory: $(Get-Location)" -ForegroundColor Yellow
    Write-Host "   Please make sure you're running this script from the correct folder." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "   ✅ requirements.txt found" -ForegroundColor Green
Write-Host ""

# Step 5: Install dependencies
Write-Host "📦 Step 5: Installing Python dependencies..." -ForegroundColor Green
Write-Host "   This may take a few minutes..." -ForegroundColor Yellow
Write-Host ""

try {
    pip install -r requirements.txt
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "   ✅ All dependencies installed successfully!" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "   ⚠️  Some dependencies may have failed to install" -ForegroundColor Yellow
        Write-Host "   Please check the output above for errors" -ForegroundColor Yellow
    }
} catch {
    Write-Host ""
    Write-Host "   ❌ Installation failed!" -ForegroundColor Red
    Write-Host "   Error: $_" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Step 6: Verify installations
Write-Host "✅ Step 6: Verifying installations..." -ForegroundColor Green
Write-Host ""

$packages = @(
    "selenium",
    "openpyxl",
    "pdfplumber",
    "PyPDF2",
    "Pillow",
    "google-generativeai",
    "rapidfuzz",
    "python-dotenv"
)

$allInstalled = $true

foreach ($package in $packages) {
    try {
        $null = pip show $package 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "   ✅ $package" -ForegroundColor Green
        } else {
            Write-Host "   ❌ $package - NOT INSTALLED" -ForegroundColor Red
            $allInstalled = $false
        }
    } catch {
        Write-Host "   ❌ $package - NOT INSTALLED" -ForegroundColor Red
        $allInstalled = $false
    }
}

Write-Host ""

# Step 7: Check Microsoft Edge
Write-Host "🌐 Step 7: Checking Microsoft Edge..." -ForegroundColor Green
Write-Host ""

$edgePaths = @(
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
)

$edgeFound = $false
foreach ($path in $edgePaths) {
    if (Test-Path $path) {
        Write-Host "   ✅ Microsoft Edge found: $path" -ForegroundColor Green
        $edgeFound = $true
        break
    }
}

if (-not $edgeFound) {
    Write-Host "   ⚠️  Microsoft Edge not found in standard locations" -ForegroundColor Yellow
    Write-Host "   Please verify Edge is installed on this system" -ForegroundColor Yellow
}

Write-Host ""

# Step 8: Check Edge WebDriver (optional check)
Write-Host "🔧 Step 8: Checking Edge WebDriver..." -ForegroundColor Green
Write-Host ""

if (Test-Path "msedgedriver.exe") {
    Write-Host "   ✅ msedgedriver.exe found in current directory" -ForegroundColor Green
} else {
    Write-Host "   ℹ️  msedgedriver.exe not found locally" -ForegroundColor Cyan
    Write-Host "   The script will download it automatically when needed" -ForegroundColor Cyan
}

Write-Host ""

# Final Summary
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  Installation Summary" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

if ($allInstalled -and $edgeFound) {
    Write-Host "✅ Installation completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "  1. Review INSTALL_GUIDE.md for detailed instructions" -ForegroundColor White
    Write-Host "  2. Configure your .env file (optional, for Gemini API)" -ForegroundColor White
    Write-Host "  3. Run the script: python badr_login_test.py" -ForegroundColor White
    Write-Host ""
    Write-Host "📚 For help, see INSTALL_GUIDE.md" -ForegroundColor Cyan
} else {
    Write-Host "⚠️  Installation completed with warnings" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please review the messages above and fix any issues." -ForegroundColor Yellow
    Write-Host "See INSTALL_GUIDE.md for troubleshooting help." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Press Enter to exit..." -ForegroundColor Gray
Read-Host
