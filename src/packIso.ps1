#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Unpacks a Windows ISO, adds autounattend.xml, and repacks it to a specified location.

.DESCRIPTION
    This script takes a Windows ISO file, extracts its contents, adds an autounattend.xml file
    for unattended installation, and creates a new ISO file at the specified output location.
    
    REQUIRES ADMINISTRATOR PRIVILEGES:
    - Mounting/unmounting ISO files
    - Installing Windows ADK (if needed)
    - Accessing system directories
    
    Run PowerShell as Administrator before executing this script.

.PARAMETER InputIso
    Path to the input Windows ISO file.

.PARAMETER OutputIso
    Path where the modified ISO will be created.

.PARAMETER AutounattendXml
    Path to the autounattend.xml file to add to the ISO. Defaults to "autounattend.xml" in the same directory as this script.

.PARAMETER WorkingDirectory
    Temporary directory for extracting ISO contents. Defaults to a temp directory.

.PARAMETER KeepWorkingDirectory
    If specified, the working directory will not be cleaned up after completion.

.PARAMETER SkipAutoInstall
    If specified, the script will not attempt to automatically install Windows ADK if oscdimg is not found.

.EXAMPLE
    .\packIso.ps1 -InputIso "C:\ISOs\Windows11.iso" -OutputIso "C:\ISOs\Windows11_Unattended.iso"

.EXAMPLE
    .\packIso.ps1 -InputIso "C:\ISOs\Windows11.iso" -OutputIso "C:\ISOs\Windows11_Unattended.iso" -AutounattendXml "C:\Custom\autounattend.xml"

.EXAMPLE
    .\packIso.ps1 -InputIso "C:\ISOs\Windows11.iso" -OutputIso "C:\ISOs\Windows11_Unattended.iso" -SkipAutoInstall
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "Input ISO file does not exist: $_"
        }
        if ($_ -notmatch '\.iso$') {
            throw "Input file must have .iso extension: $_"
        }
        $true
    })]
    [string]$InputIso,

    [Parameter(Mandatory = $true)]
    [ValidateScript({
        $parentDir = Split-Path $_ -Parent
        if (-not (Test-Path $parentDir -PathType Container)) {
            throw "Output directory does not exist: $parentDir"
        }
        if ($_ -notmatch '\.iso$') {
            throw "Output file must have .iso extension: $_"
        }
        $true
    })]
    [string]$OutputIso,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "Autounattend XML file does not exist: $_"
        }
        $true
    })]
    [string]$AutounattendXml = (Join-Path $PSScriptRoot "autounattend.xml"),

    [Parameter(Mandatory = $false)]
    [string]$WorkingDirectory = (Join-Path $env:TEMP "WindowsIsoRepack_$(Get-Date -Format 'yyyyMMdd_HHmmss')"),

    [Parameter(Mandatory = $false)]
    [switch]$KeepWorkingDirectory,

    [Parameter(Mandatory = $false)]
    [switch]$SkipAutoInstall
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to check if required tools are available
function Test-RequiredTools {
    Write-ColorOutput "Checking for required tools..." "Yellow"
    
    # Check for oscdimg (Windows ADK tool)
    $oscdimgPath = Get-Command "oscdimg.exe" -ErrorAction SilentlyContinue
    if (-not $oscdimgPath) {
        # Try common installation paths
        $commonPaths = @(
            "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
            "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
            "${env:ProgramFiles(x86)}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
            "${env:ProgramFiles}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
        )
        
        foreach ($path in $commonPaths) {
            if (Test-Path $path) {
                $script:oscdimgPath = $path
                break
            }
        }
        
        if (-not $script:oscdimgPath) {
            if ($SkipAutoInstall) {
                throw "oscdimg.exe not found. Please install Windows ADK (Assessment and Deployment Kit) or ensure oscdimg.exe is in your PATH. Use -SkipAutoInstall:$false to enable automatic installation."
            } else {
                Write-ColorOutput "oscdimg.exe not found. Attempting to install Windows ADK..." "Yellow"
                Install-WindowsADK
            }
        }
    } else {
        $script:oscdimgPath = $oscdimgPath.Source
    }
    
    Write-ColorOutput "Found oscdimg.exe at: $script:oscdimgPath" "Green"
}

# Function to check if Chocolatey is installed
function Test-Chocolatey {
    return (Get-Command "choco" -ErrorAction SilentlyContinue) -ne $null
}
function Install-Chocolatey {
    Write-ColorOutput "Installing Chocolatey package manager..." "Yellow"
    try {
        
        Write-Log "Installing Chocolatey..."
        # Install Chocolatey
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        
        if (Test-Chocolatey) {
            Write-ColorOutput "✓ Chocolatey installed successfully!" "Green"
            return $true
        } else {
            Write-ColorOutput "Chocolatey installation failed to become available on PATH." "Red"
            return $false
        }
    } catch {
        Write-ColorOutput "Chocolatey installation failed: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Install-WindowsADK {
    Write-ColorOutput "=== Windows ADK Installation ===" "Cyan"

    if (-not (Test-Chocolatey)) {
        Write-ColorOutput "Chocolatey not found. Installing Chocolatey..." "Yellow"
        if (-not (Install-Chocolatey)) {
            Write-ColorOutput "Failed to install Chocolatey. Showing manual installation instructions..." "Red"
            Show-ManualInstallInstructions
            throw "Windows ADK installation aborted (Chocolatey missing)."
        }
    } else {
        Write-ColorOutput "✓ Chocolatey found" "Green"
    }

    Write-ColorOutput "Installing Windows ADK via Chocolatey..." "Yellow"
    try {
        $result = Start-Process -FilePath "choco" -ArgumentList @("install","windows-adk","-y") -Wait -PassThru -NoNewWindow
        if ($result.ExitCode -eq 0) {
            Write-ColorOutput "✓ Windows ADK installed successfully via Chocolatey!" "Green"

            # Refresh PATH so newly installed tools are discoverable
            $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User')
            return
        } else {
            Write-ColorOutput "Windows ADK installation via Chocolatey failed with exit code: $($result.ExitCode)" "Red"
        }
    } catch {
        Write-ColorOutput "Windows ADK installation via Chocolatey threw an error: $($_.Exception.Message)" "Red"
    }

    # If we’re here, Chocolatey path failed — fall back to manual instructions
    Write-ColorOutput "Falling back to manual installation instructions..." "Yellow"
    Show-ManualInstallInstructions
    throw "Windows ADK installation required. Please install it and run this script again."
}

# Function to show manual installation instructions
function Show-ManualInstallInstructions {
    Write-ColorOutput "=== Manual Installation Instructions ===" "Cyan"
    Write-ColorOutput "Please install Windows ADK manually:" "Yellow"
    Write-ColorOutput "1. Download Windows ADK from: https://docs.microsoft.com/en-us/windows-hardware/get-started/adk-install" "White"
    Write-ColorOutput "2. Run the installer and select 'Deployment Tools'" "White"
    Write-ColorOutput "3. Restart this script after installation" "White"
    Write-ColorOutput "" "White"
    Write-ColorOutput "Alternative: Install via Chocolatey:" "White"
    Write-ColorOutput "  choco install windows-adk -y" "Cyan"
    Write-ColorOutput "" "White"
    Write-ColorOutput "Or install Chocolatey first, then Windows ADK:" "White"
    Write-ColorOutput "  Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))" "Cyan"
    Write-ColorOutput "  choco install windows-adk -y" "Cyan"
    
    # Ask user if they want to continue with manual download
    $response = Read-Host "Would you like to open the ADK download page? (y/n)"
    if ($response -eq 'y' -or $response -eq 'Y') {
        Start-Process "https://docs.microsoft.com/en-us/windows-hardware/get-started/adk-install"
    }
    
    throw "Windows ADK installation required. Please install it and run this script again."
}

# Function to mount ISO and extract contents
function Extract-IsoContents {
    param(
        [string]$IsoPath,
        [string]$ExtractPath
    )
    
    Write-ColorOutput "Mounting ISO: $IsoPath" "Yellow"
    
    # Mount the ISO
    $mountResult = Mount-DiskImage -ImagePath $IsoPath -PassThru
    $driveLetter = ($mountResult | Get-Volume).DriveLetter
    
    if (-not $driveLetter) {
        throw "Failed to mount ISO or get drive letter"
    }
    
    $mountedPath = "${driveLetter}:\"
    Write-ColorOutput "ISO mounted at: $mountedPath" "Green"
    
    try {
        Write-ColorOutput "Extracting ISO contents to: $ExtractPath" "Yellow"
        
        # Create extraction directory
        if (Test-Path $ExtractPath) {
            Remove-Item $ExtractPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null
        
        # Copy all contents from mounted ISO
        robocopy $mountedPath $ExtractPath /E /COPY:DAT /R:3 /W:10 /NFL /NDL /NJH /NJS /nc /ns /np
        
        if ($LASTEXITCODE -gt 7) {
            throw "Failed to extract ISO contents. Robocopy exit code: $LASTEXITCODE"
        }
        
        Write-ColorOutput "ISO contents extracted successfully" "Green"
        
    } finally {
        # Dismount the ISO
        Write-ColorOutput "Dismounting ISO..." "Yellow"
        Dismount-DiskImage -ImagePath $IsoPath
        Write-ColorOutput "ISO dismounted" "Green"
    }
}

# Function to add autounattend.xml to ISO contents
function Add-AutounattendXml {
    param(
        [string]$ExtractPath,
        [string]$AutounattendXmlPath
    )
    
    Write-ColorOutput "Adding autounattend.xml to ISO contents..." "Yellow"
    
    # Copy autounattend.xml to the root of the extracted ISO
    $destinationPath = Join-Path $ExtractPath "autounattend.xml"
    Copy-Item $AutounattendXmlPath $destinationPath -Force
    
    Write-ColorOutput "autounattend.xml added to: $destinationPath" "Green"
}

# Function to create new ISO
function New-IsoFromDirectory {
    param(
        [string]$SourcePath,
        [string]$OutputPath,
        [string]$OscdimgPath
    )
    
    Write-ColorOutput "Creating new ISO from directory: $SourcePath" "Yellow"
    
    # Prepare oscdimg command arguments
    $arguments = @(
        "-m"                    # Allow files larger than 2GB
        "-o"                    # Optimize storage by encoding duplicate files only once
        "-u2"                   # Use UDF file system
        "-udfver102"            # Use UDF version 1.02
        "-l"                    # Set volume label
        "Windows"               # Volume label
        "`"$SourcePath`""       # Source directory
        "`"$OutputPath`""       # Output ISO file
    )
    
    Write-ColorOutput "Running oscdimg with arguments: $($arguments -join ' ')" "Cyan"
    
    # Execute oscdimg
    $process = Start-Process -FilePath $OscdimgPath -ArgumentList $arguments -Wait -PassThru -NoNewWindow
    
    if ($process.ExitCode -ne 0) {
        throw "oscdimg failed with exit code: $($process.ExitCode)"
    }
    
    Write-ColorOutput "ISO created successfully: $OutputPath" "Green"
}

# Function to cleanup working directory
function Remove-WorkingDirectory {
    param([string]$Path)
    
    if (-not $KeepWorkingDirectory -and (Test-Path $Path)) {
        Write-ColorOutput "Cleaning up working directory: $Path" "Yellow"
        Remove-Item $Path -Recurse -Force
        Write-ColorOutput "Working directory cleaned up" "Green"
    } elseif ($KeepWorkingDirectory) {
        Write-ColorOutput "Keeping working directory: $Path" "Cyan"
    }
}

# Main execution
try {
    # Check if running as administrator FIRST - before any other operations
    Write-ColorOutput "=== Windows ISO Repack Script ===" "Cyan"
    Write-ColorOutput "Checking administrator privileges..." "Yellow"
    
    if (-not (Test-Administrator)) {
        Write-ColorOutput "ERROR: This script must be run as Administrator!" "Red"
        Write-ColorOutput "" "White"
        Write-ColorOutput "Required for:" "White"
        Write-ColorOutput "  - Mounting/unmounting ISO files" "White"
        Write-ColorOutput "  - Installing Windows ADK (if needed)" "White"
        Write-ColorOutput "  - Accessing system directories" "White"
        Write-ColorOutput "" "White"
        Write-ColorOutput "To run as Administrator:" "White"
        Write-ColorOutput "  1. Right-click PowerShell" "White"
        Write-ColorOutput "  2. Select 'Run as Administrator'" "White"
        Write-ColorOutput "  3. Navigate to script directory" "White"
        Write-ColorOutput "  4. Run the script again" "White"
        Write-ColorOutput "" "White"
        Write-ColorOutput "Example:" "White"
        Write-ColorOutput "  cd 'C:\path\to\script'" "Cyan"
        Write-ColorOutput "  .\packIso.ps1 -InputIso 'C:\ISOs\Windows11.iso' -OutputIso 'C:\ISOs\Windows11_Unattended.iso'" "Cyan"
        throw "Administrator privileges required. Please run PowerShell as Administrator and try again."
    }
    
    Write-ColorOutput "✓ Administrator privileges confirmed" "Green"
    Write-ColorOutput "" "White"
    Write-ColorOutput "Script Configuration:" "White"
    Write-ColorOutput "Input ISO: $InputIso" "White"
    Write-ColorOutput "Output ISO: $OutputIso" "White"
    Write-ColorOutput "Autounattend XML: $AutounattendXml" "White"
    Write-ColorOutput "Working Directory: $WorkingDirectory" "White"
    Write-ColorOutput "Skip Auto Install: $SkipAutoInstall" "White"
    Write-ColorOutput "=================================" "Cyan"
    
    # Check for required tools
    Test-RequiredTools
    
    # Validate input files
    Write-ColorOutput "Validating input files..." "Yellow"
    if (-not (Test-Path $InputIso -PathType Leaf)) {
        throw "Input ISO file not found: $InputIso"
    }
    if (-not (Test-Path $AutounattendXml -PathType Leaf)) {
        throw "Autounattend XML file not found: $AutounattendXml"
    }
    Write-ColorOutput "Input files validated" "Green"
    
    # Check if output file already exists
    if (Test-Path $OutputIso) {
        Write-ColorOutput "Output ISO already exists. Removing..." "Yellow"
        Remove-Item $OutputIso -Force
    }
    
    # Extract ISO contents
    Extract-IsoContents -IsoPath $InputIso -ExtractPath $WorkingDirectory
    
    # Add autounattend.xml
    Add-AutounattendXml -ExtractPath $WorkingDirectory -AutounattendXmlPath $AutounattendXml
    
    # Create new ISO
    New-IsoFromDirectory -SourcePath $WorkingDirectory -OutputPath $OutputIso -OscdimgPath $script:oscdimgPath
    
    # Verify output file was created
    if (Test-Path $OutputIso) {
        $fileSize = (Get-Item $OutputIso).Length
        $fileSizeGB = [math]::Round($fileSize / 1GB, 2)
        Write-ColorOutput "Output ISO created successfully!" "Green"
        Write-ColorOutput "File size: $fileSizeGB GB" "Green"
    } else {
        throw "Output ISO was not created successfully"
    }
    
} catch {
    Write-ColorOutput "Error: $($_.Exception.Message)" "Red"
    exit 1
} finally {
    # Cleanup
    Remove-WorkingDirectory -Path $WorkingDirectory
}

Write-ColorOutput "=== Script completed successfully! ===" "Green"
