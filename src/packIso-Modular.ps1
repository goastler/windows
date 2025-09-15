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
    [string]$VirtioCacheDirectory = (Join-Path $env:TEMP "virtio-cache")

)

$ErrorActionPreference = "Stop"

# Dot source all modules
$modulePath = Join-Path $PSScriptRoot "modules"

Write-ColorOutput "=== Loading Modules ===" "Cyan"
Write-ColorOutput "Loading Common utilities..." "Yellow" -Indent 1
. (Join-Path $modulePath "Common.ps1")

Write-ColorOutput "Loading Tools and Prerequisites..." "Yellow" -Indent 1
$toolsPath = Join-Path $modulePath "tools"
. (Join-Path $toolsPath "Chocolatey.ps1")
. (Join-Path $toolsPath "WindowsADK.ps1")
. (Join-Path $toolsPath "OSCDIMG.ps1")
. (Join-Path $toolsPath "DISM.ps1")
. (Join-Path $toolsPath "ToolsOrchestrator.ps1")

Write-ColorOutput "Loading ISO Operations..." "Yellow" -Indent 1
. (Join-Path $modulePath "ISO.ps1")

Write-ColorOutput "Loading WIM Analysis..." "Yellow" -Indent 1
. (Join-Path $modulePath "WIM.ps1")

Write-ColorOutput "Loading VirtIO Drivers..." "Yellow" -Indent 1
. (Join-Path $modulePath "VirtIO.ps1")

Write-ColorOutput "All modules loaded successfully" "Green"

try {
    Write-ColorOutput "=== Windows ISO Repack Script ===" "Cyan"
    Write-ColorOutput "Checking administrator privileges..." "Yellow"
    
    if (-not (Test-Administrator)) {
        Write-ColorOutput "ERROR: This script must be run as Administrator!" "Red"
        throw "Administrator privileges required."
    }
    
    Write-ColorOutput "Administrator privileges confirmed" "Green"
    Write-ColorOutput "Input ISO: $InputIso" "White"
    Write-ColorOutput "Output ISO: $OutputIso" "White"
    Write-ColorOutput "Autounattend XML: $AutounattendXml" "White"
    Write-ColorOutput "OEM Directory: $OemDirectory" "White"
    Write-ColorOutput "Working Directory: $WorkingDirectory" "White"
    Write-ColorOutput "Include VirtIO Drivers: $IncludeVirtioDrivers" "White"
    if ($IncludeVirtioDrivers) {
        Write-ColorOutput "VirtIO Version: $VirtioVersion" "White"
        Write-ColorOutput "VirtIO Cache Directory: $VirtioCacheDirectory" "White"
        Write-ColorOutput "Processing Mode: Per-WIM (architecture and version inferred from each WIM)" "White"
    }
    
    Test-RequiredTools
    
    Write-ColorOutput "Validating input files..." "Yellow"
    
    # Resolve and validate ISO paths
    try {
        $resolvedInputIso = Resolve-Path $InputIso -ErrorAction Stop
        Write-ColorOutput "Resolved input ISO path: $resolvedInputIso" "Cyan"
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
        Write-ColorOutput "Resolved output ISO path: $resolvedOutputIso" "Cyan"
    } catch {
        throw "Cannot resolve output ISO file path: $OutputIso. Error: $($_.Exception.Message)"
    }
    
    if (-not (Test-Path $resolvedInputIso -PathType Leaf)) {
        throw "Input ISO file not found: $resolvedInputIso"
    }
    if (-not (Test-Path $AutounattendXml -PathType Leaf)) {
        throw "Autounattend XML file not found: $AutounattendXml"
    }
    Write-ColorOutput "Input files validated" "Green"
    
    if (Test-Path $resolvedOutputIso) {
        Write-ColorOutput "Output ISO already exists. Removing..." "Yellow"
        Remove-Item $resolvedOutputIso -Force
    }
    
    # Step 1: Extract ISO contents
    Write-ColorOutput "=== Step 1: Extracting ISO Contents ===" "Cyan"
    Extract-IsoContents -IsoPath $resolvedInputIso -ExtractPath $WorkingDirectory
    
    # Step 2: Add autounattend.xml and OEM directory
    Write-ColorOutput "=== Step 2: Adding Configuration Files ===" "Cyan"
    Add-AutounattendXml -ExtractPath $WorkingDirectory -AutounattendXmlPath $AutounattendXml
    Add-OemDirectory -ExtractPath $WorkingDirectory -OemSourcePath $OemDirectory
    
    # Step 3: Add VirtIO drivers to WIM files
    if ($IncludeVirtioDrivers) {
        Write-ColorOutput "=== Step 3: Adding VirtIO Drivers ===" "Cyan"
        Add-VirtioDrivers -ExtractPath $WorkingDirectory -VirtioVersion $VirtioVersion -VirtioCacheDirectory $VirtioCacheDirectory
    } else {
        Write-ColorOutput "=== Step 3: Skipping VirtIO Drivers (not requested) ===" "Cyan"
    }
    
    # Step 4: Create new ISO
    Write-ColorOutput "=== Step 4: Creating New ISO ===" "Cyan"
    New-IsoFromDirectory -SourcePath $WorkingDirectory -OutputPath $resolvedOutputIso -OscdimgPath $script:oscdimgPath
    
    if (Test-Path $resolvedOutputIso) {
        $fileSize = (Get-Item $resolvedOutputIso).Length
        $fileSizeGB = [math]::Round($fileSize / 1GB, 2)
        Write-ColorOutput "Output ISO created successfully!" "Green"
        Write-ColorOutput "File size: $fileSizeGB GB" "Green" -Indent 1
    } else {
        throw "Output ISO was not created successfully"
    }
} catch {
    Write-ColorOutput "Error: $($_.Exception.Message)" "Red"
    exit 1
} finally {
    Remove-WorkingDirectory -Path $WorkingDirectory
}

Write-ColorOutput "=== Script completed successfully! ===" "Green"
