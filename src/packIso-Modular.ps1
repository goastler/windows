# #Requires -RunAsAdministrator

# Windows ISO Repack Script - Modular Version
# This script repacks Windows ISOs with autounattend.xml, OEM directory, and VirtIO drivers

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
    [string]$OemDirectory = (Join-Path (Split-Path $PSScriptRoot -Parent) '$OEM$'),

    [Parameter(Mandatory = $false)]
    [string]$WorkingDirectory = "C:\WinIsoRepack_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [switch]$KeepWorkingDirectory,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeVirtioDrivers,

    [Parameter(Mandatory = $false)]
    [ValidateSet("stable", "latest")]
    [string]$VirtioVersion = "stable",

    [Parameter(Mandatory = $false)]
    [string]$VirtioCacheDirectory = (Join-Path $env:TEMP "virtio-cache"),

    [Parameter(Mandatory = $false)]
    [switch]$OverwriteOutputIso

)

$ErrorActionPreference = "Stop"

# Dot source all modules
$modulePath = Join-Path $PSScriptRoot "modules"

Write-ColorOutput "=== Loading Modules ===" -Color "Cyan"
Write-ColorOutput "Loading Common utilities..." -Color "Yellow" -Indent 1
. (Join-Path $modulePath "Common.ps1")

Write-ColorOutput "Loading Tools and Prerequisites..." -Color "Yellow" -Indent 1
$toolsPath = Join-Path $modulePath "tools"
. (Join-Path $toolsPath "Chocolatey.ps1")
. (Join-Path $toolsPath "WindowsADK.ps1")
. (Join-Path $toolsPath "OSCDIMG.ps1")
. (Join-Path $toolsPath "DISM.ps1")
. (Join-Path $toolsPath "ToolsOrchestrator.ps1")

Write-ColorOutput "Loading ISO Operations..." -Color "Yellow" -Indent 1
. (Join-Path $modulePath "ISO.ps1")

Write-ColorOutput "Loading WIM Analysis..." -Color "Yellow" -Indent 1
. (Join-Path $modulePath "WIM.ps1")

Write-ColorOutput "Loading VirtIO Drivers..." -Color "Yellow" -Indent 1
. (Join-Path $modulePath "VirtIO.ps1")

Write-ColorOutput "All modules loaded successfully" -Color "Green"

try {
    Write-ColorOutput "=== Windows ISO Repack Script ===" -Color "Cyan"
    Write-ColorOutput "Checking administrator privileges..." -Color "Yellow"
    
    if (-not (Test-Administrator)) {
        Write-ColorOutput "ERROR: This script must be run as Administrator!" -Color "Red"
        throw "Administrator privileges required."
    }
    
    Write-ColorOutput "Administrator privileges confirmed" -Color "Green"
    Write-ColorOutput "Input ISO: $InputIso" -Color "White"
    Write-ColorOutput "Output ISO: $OutputIso" -Color "White"
    Write-ColorOutput "Autounattend XML: $AutounattendXml" -Color "White"
    Write-ColorOutput "OEM Directory: $OemDirectory" -Color "White"
    Write-ColorOutput "Working Directory: $WorkingDirectory" -Color "White"
    Write-ColorOutput "Include VirtIO Drivers: $IncludeVirtioDrivers" -Color "White"
    if ($IncludeVirtioDrivers) {
        Write-ColorOutput "VirtIO Version: $VirtioVersion" -Color "White"
        Write-ColorOutput "VirtIO Cache Directory: $VirtioCacheDirectory" -Color "White"
    }
    
    Test-RequiredTools
    
    Write-ColorOutput "Validating input files..." -Color "Yellow"
    
    # Resolve and validate ISO paths
    try {
        $resolvedInputIso = Resolve-Path $InputIso -ErrorAction Stop
        Write-ColorOutput "Resolved input ISO path: $resolvedInputIso" -Color "Cyan"
    } catch {
        throw "Cannot resolve input ISO file path: $InputIso. Error: $($_.Exception.Message)"
    }
    
    try {
        $resolvedOutputIso = Resolve-Path $OutputIso -ErrorAction SilentlyContinue
        if (-not $resolvedOutputIso) {
            # If path doesn't exist, resolve the parent directory and create the full path
            $outputDir = Split-Path $OutputIso -Parent
            $outputFile = Split-Path $OutputIso -Leaf
            $resolvedOutputDir = Resolve-Path $outputDir -ErrorAction Stop
            $resolvedOutputIso = Join-Path $resolvedOutputDir $outputFile
        }
        Write-ColorOutput "Resolved output ISO path: $resolvedOutputIso" -Color "Cyan"
    } catch {
        throw "Cannot resolve output ISO file path: $OutputIso. Error: $($_.Exception.Message)"
    }
    
    if (-not (Test-Path $resolvedInputIso -PathType Leaf)) {
        throw "Input ISO file not found: $resolvedInputIso"
    }
    if (-not (Test-Path $AutounattendXml -PathType Leaf)) {
        throw "Autounattend XML file not found: $AutounattendXml"
    }
    Write-ColorOutput "Input files validated" -Color "Green"
    
    if (Test-Path $resolvedOutputIso) {
        if ($OverwriteOutputIso) {
            Write-ColorOutput "Output ISO already exists. Overwriting..." -Color "Yellow"
            Remove-Item $resolvedOutputIso -Force
        } else {
            Write-ColorOutput "ERROR: Output ISO already exists: $resolvedOutputIso" -Color "Red"
            Write-ColorOutput "Use -OverwriteOutputIso parameter to overwrite the existing file." -Color "Red"
            throw "Output ISO file already exists. Use -OverwriteOutputIso to overwrite."
        }
    }
    
    # Extract ISO contents
    Write-ColorOutput "=== Extracting ISO Contents ===" -Color "Cyan"
    Extract-IsoContents -IsoPath $resolvedInputIso -ExtractPath $WorkingDirectory
    
    # Add autounattend.xml and OEM directory
    Write-ColorOutput "=== Adding Configuration Files ===" -Color "Cyan"
    Add-AutounattendXml -ExtractPath $WorkingDirectory -AutounattendXmlPath $AutounattendXml
    Add-OemDirectory -ExtractPath $WorkingDirectory -OemSourcePath $OemDirectory
    
    # Add VirtIO drivers to WIM files
    if ($IncludeVirtioDrivers) {
        Write-ColorOutput "=== Adding VirtIO Drivers ===" -Color "Cyan"
        Add-VirtioDrivers -ExtractPath $WorkingDirectory -VirtioVersion $VirtioVersion -VirtioCacheDirectory $VirtioCacheDirectory
    } else {
        Write-ColorOutput "=== Skipping VirtIO Drivers (not requested) ===" -Color "Cyan"
    }
    
    # Create new ISO
    Write-ColorOutput "=== Creating New ISO ===" -Color "Cyan"
    New-IsoFromDirectory -SourcePath $WorkingDirectory -OutputPath $resolvedOutputIso -OscdimgPath $script:oscdimgPath
    
    if (Test-Path $resolvedOutputIso) {
        $fileSize = (Get-Item $resolvedOutputIso).Length
        $fileSizeGB = [math]::Round($fileSize / 1GB, 2)
        Write-ColorOutput "Output ISO created successfully!" -Color "Green"
        Write-ColorOutput "File size: $fileSizeGB GB" -Color "Green" -Indent 1
    } else {
        throw "Output ISO was not created successfully"
    }
} catch {
    Write-ColorOutput "Error: $($_.Exception.Message)" -Color "Red"
    exit 1
} finally {
    Remove-WorkingDirectory -Path $WorkingDirectory
}

Write-ColorOutput "=== Script completed successfully! ===" -Color "Green"
