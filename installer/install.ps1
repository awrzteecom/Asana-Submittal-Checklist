# DOCX to Asana CSV Generator Installer
# This script installs the DOCX to Asana CSV Generator application
# and its dependencies.

# Script parameters
param (
    [switch]$CreateShortcut = $true,
    [switch]$InstallDependencies = $true,
    [string]$InstallDir = "$env:USERPROFILE\DOCX-to-Asana"
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Function to check if Python is installed
function Check-Python {
    try {
        $pythonVersion = python --version 2>&1
        if ($pythonVersion -match "Python (\d+)\.(\d+)\.(\d+)") {
            $major = [int]$Matches[1]
            $minor = [int]$Matches[2]
            
            if ($major -ge 3 -and $minor -ge 8) {
                Write-Host "Python $major.$minor.$($Matches[3]) found. Requirement satisfied." -ForegroundColor Green
                return $true
            } else {
                Write-Host "Python $major.$minor.$($Matches[3]) found, but version 3.8 or higher is required." -ForegroundColor Yellow
                return $false
            }
        } else {
            Write-Host "Python version could not be determined." -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Host "Python is not installed or not in PATH." -ForegroundColor Yellow
        return $false
    }
}

# Function to install Python dependencies
function Install-Dependencies {
    Write-Host "Installing Python dependencies..." -ForegroundColor Cyan
    try {
        python -m pip install --upgrade pip
        python -m pip install -r "$InstallDir\requirements.txt"
        Write-Host "Dependencies installed successfully." -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Error installing dependencies. Please check if Python is installed correctly." -ForegroundColor Red
        return $false
    }
}

# Function to create desktop shortcut
function Create-Shortcut {
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath "DOCX to Asana.lnk"
    $targetPath = "pythonw.exe"
    $arguments = """$InstallDir\src\docx_to_asana.py"" --gui"
    $workingDirectory = $InstallDir
    
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = $targetPath
    $Shortcut.Arguments = $arguments
    $Shortcut.WorkingDirectory = $workingDirectory
    $Shortcut.IconLocation = "shell32.dll,23"  # Document icon
    $Shortcut.Description = "DOCX to Asana CSV Generator"
    $Shortcut.Save()
    
    Write-Host "Desktop shortcut created: $shortcutPath" -ForegroundColor Green
}

# Main installation process
function Install-Application {
    Write-Host "=== DOCX to Asana CSV Generator Installer ===" -ForegroundColor Cyan
    
    # Check if Python is installed
    if (-not (Check-Python)) {
        Write-Host "Please install Python 3.8 or higher from https://www.python.org/downloads/" -ForegroundColor Yellow
        Write-Host "After installing Python, run this installer again." -ForegroundColor Yellow
        return $false
    }
    
    # Create installation directory if it doesn't exist
    if (-not (Test-Path $InstallDir)) {
        Write-Host "Creating installation directory: $InstallDir" -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    }
    
    # Determine source directory
    $sourceDir = $null
    try {
        # Try to get the script directory
        $scriptPath = $MyInvocation.MyCommand.Path
        if ($scriptPath) {
            # If script path is available, use it to determine source directory
            $sourceDir = Split-Path -Parent (Split-Path -Parent $scriptPath)
            Write-Host "Detected source directory: $sourceDir" -ForegroundColor Cyan
        }
    } catch {
        Write-Host "Could not determine source directory from script path." -ForegroundColor Yellow
    }
    
    # If source directory is still null, try current directory
    if (-not $sourceDir -or -not (Test-Path $sourceDir)) {
        $sourceDir = Get-Location
        Write-Host "Using current directory as source: $sourceDir" -ForegroundColor Cyan
    }
    
    # Verify source directory contains required files
    if (-not (Test-Path (Join-Path $sourceDir "src"))) {
        Write-Host "Warning: Source directory does not contain 'src' folder. Installation may be incomplete." -ForegroundColor Yellow
    }
    
    Write-Host "Copying application files from $sourceDir to $InstallDir" -ForegroundColor Cyan
    
    # Create necessary directories
    $directories = @(
        "src",
        "src\utils",
        "tests",
        "tests\sample_documents"
    )
    
    foreach ($dir in $directories) {
        $targetDir = Join-Path $InstallDir $dir
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
    }
    
    # Copy files
    $filesToCopy = @(
        "requirements.txt",
        "config.json",
        "README.md",
        "src\docx_to_asana.py",
        "src\document_parser.py",
        "src\csv_generator.py",
        "src\gui_handler.py",
        "src\utils\logger.py",
        "src\utils\config.py",
        "src\utils\validator.py",
        "tests\test_parser.py",
        "tests\test_csv_generator.py"
    )
    
    $copyErrors = 0
    foreach ($file in $filesToCopy) {
        $sourcePath = Join-Path $sourceDir $file
        $targetPath = Join-Path $InstallDir $file
        
        # Ensure target directory exists
        $targetDir = Split-Path -Parent $targetPath
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        
        if (Test-Path $sourcePath) {
            try {
                Copy-Item -Path $sourcePath -Destination $targetPath -Force
                Write-Host "Copied: $file" -ForegroundColor Gray
            } catch {
                Write-Host "Error copying file $file" -ForegroundColor Red
                $copyErrors++
            }
        } else {
            Write-Host "Warning: Source file not found: $sourcePath" -ForegroundColor Yellow
            $copyErrors++
        }
    }
    
    if ($copyErrors -gt 0) {
        Write-Host "Warning: $copyErrors files could not be copied. Installation may be incomplete." -ForegroundColor Yellow
    }
    
    # Create requirements.txt if it doesn't exist
    $requirementsPath = Join-Path $InstallDir "requirements.txt"
    if (-not (Test-Path $requirementsPath)) {
        Write-Host "Creating default requirements.txt file" -ForegroundColor Cyan
        @"
python-docx>=0.8.11
pandas>=1.3.0
openpyxl>=3.0.7
pytest>=6.2.5
"@ | Out-File -FilePath $requirementsPath -Encoding utf8
    }
    
    # Install dependencies if requested
    if ($InstallDependencies) {
        $dependenciesInstalled = Install-Dependencies
        if (-not $dependenciesInstalled) {
            Write-Host "Warning: Some dependencies may not have been installed correctly." -ForegroundColor Yellow
        }
    }
    
    # Create desktop shortcut if requested
    if ($CreateShortcut) {
        try {
            Create-Shortcut
        } catch {
            Write-Host "Error creating desktop shortcut." -ForegroundColor Yellow
            Write-Host "You can still run the application manually." -ForegroundColor Cyan
        }
    }
    
    Write-Host "Installation completed successfully!" -ForegroundColor Green
    Write-Host "You can run the application using the desktop shortcut or by running:" -ForegroundColor Cyan
    Write-Host "python $InstallDir\src\docx_to_asana.py --gui" -ForegroundColor Cyan
    
    return $true
}

# Run the installation
try {
    $result = Install-Application
    if ($result) {
        Write-Host "Press any key to exit..." -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 0
    } else {
        Write-Host "Installation did not complete successfully." -ForegroundColor Red
        Write-Host "Press any key to exit..." -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
} catch {
    Write-Host "Error during installation. Please check your system configuration." -ForegroundColor Red
    Write-Host "Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
