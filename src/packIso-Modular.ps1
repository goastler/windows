# #Requires -RunAsAdministrator

# Windows ISO Repack Script - Modular Version
# This script repacks Windows ISOs with autounattend.xml, OEM directory, and VirtIO drivers

param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        # Resolve to absolute path
        $resolvedPath = if ([System.IO.Path]::IsPathRooted($_)) {
            $_
        } else {
            Join-Path (Get-Location) $_
        }
        
        if (-not (Test-Path $resolvedPath -PathType Leaf)) {
            throw "Input ISO file does not exist: $resolvedPath (resolved from: $_)"
        }
        if ($resolvedPath -notmatch '\.iso$') {
            throw "Input file must have .iso extension: $resolvedPath"
        }
        $true
    })]
    [string]$InputIso,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        # Resolve to absolute path
        $resolvedPath = if ([System.IO.Path]::IsPathRooted($_)) {
            $_
        } else {
            Join-Path (Get-Location) $_
        }
        
        $parentDir = Split-Path $resolvedPath -Parent
        if (-not (Test-Path $parentDir -PathType Container)) {
            throw "Output directory does not exist: $parentDir (resolved from: $_)"
        }
        if ($resolvedPath -notmatch '\.iso$') {
            throw "Output file must have .iso extension: $resolvedPath"
        }
        $true
    })]
    [string]$OutputIso,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        # Resolve to absolute path
        $resolvedPath = if ([System.IO.Path]::IsPathRooted($_)) {
            $_
        } else {
            Join-Path (Get-Location) $_
        }
        
        if (-not (Test-Path $resolvedPath -PathType Leaf)) {
            throw "Autounattend XML file does not exist: $resolvedPath (resolved from: $_)"
        }
        if ($resolvedPath -notmatch '\.xml$') {
            throw "Autounattend file must have .xml extension: $resolvedPath"
        }
        $true
    })]
    [string]$AutounattendXml = (Join-Path $PSScriptRoot "autounattend.xml"),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        # Resolve to absolute path
        $resolvedPath = if ([System.IO.Path]::IsPathRooted($_)) {
            $_
        } else {
            Join-Path (Get-Location) $_
        }
        
        # Note: We don't check if directory exists here as it might be created during processing
        $true
    })]
    [string]$OemDirectory = (Join-Path (Split-Path $PSScriptRoot -Parent) '$OEM$'),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        # Resolve to absolute path
        $resolvedPath = if ([System.IO.Path]::IsPathRooted($_)) {
            $_
        } else {
            Join-Path (Get-Location) $_
        }
        
        # Note: We don't check if directory exists here as it will be created during processing
        $true
    })]
    [string]$WorkingDirectory = "C:\WinIsoRepack_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [switch]$KeepWorkingDirectory,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeVirtioDrivers,

    [Parameter(Mandatory = $false)]
    [ValidateSet("stable", "latest")]
    [string]$VirtioVersion = "stable",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        # Resolve to absolute path
        $resolvedPath = if ([System.IO.Path]::IsPathRooted($_)) {
            $_
        } else {
            Join-Path (Get-Location) $_
        }
        
        # Note: We don't check if directory exists here as it will be created during processing
        $true
    })]
    [string]$VirtioCacheDirectory = (Join-Path $env:TEMP "virtio-cache"),

    [Parameter(Mandatory = $false)]
    [switch]$OverwriteOutputIso

)

$ErrorActionPreference = "Stop"

# Dot source all modules
$modulePath = Join-Path $PSScriptRoot "modules"

# Import Common.ps1 first (contains Write-ColorOutput)
. (Join-Path $modulePath "Common.ps1")

Write-ColorOutput "=== Importing Modules ===" -Color "Cyan" -CurrentIndent 0 -InheritedIndent 0
Write-ColorOutput "Importing Tools and Prerequisites..." -Color "Yellow" -CurrentIndent 1 -InheritedIndent 0
$toolsPath = Join-Path $modulePath "tools"
. (Join-Path $toolsPath "Chocolatey.ps1")
. (Join-Path $toolsPath "WindowsADK.ps1")
. (Join-Path $toolsPath "OSCDIMG.ps1")
. (Join-Path $toolsPath "DISM.ps1")
. (Join-Path $toolsPath "ToolsOrchestrator.ps1")

Write-ColorOutput "Importing ISO Operations..." -Color "Yellow" -CurrentIndent 1 -InheritedIndent 0
. (Join-Path $modulePath "ISO.ps1")

Write-ColorOutput "Importing WIM Analysis..." -Color "Yellow" -CurrentIndent 1 -InheritedIndent 0
. (Join-Path $modulePath "WIM.ps1")

Write-ColorOutput "Importing VirtIO Drivers..." -Color "Yellow" -CurrentIndent 1 -InheritedIndent 0
. (Join-Path $modulePath "VirtIO.ps1")

Write-ColorOutput "All modules imported successfully" -Color "Green"

# Resolve all paths to absolute paths
Write-ColorOutput "=== Resolving Paths ===" -Color "Cyan" -CurrentIndent 0 -InheritedIndent 0

# Resolve InputIso
$InputIso = if ([System.IO.Path]::IsPathRooted($InputIso)) {
    $InputIso
} else {
    $resolved = Join-Path (Get-Location) $InputIso
    Write-ColorOutput "Resolved InputIso: $InputIso -> $resolved" -Color "Cyan" -CurrentIndent 1 -InheritedIndent 0
    $resolved
}

# Resolve OutputIso
$OutputIso = if ([System.IO.Path]::IsPathRooted($OutputIso)) {
    $OutputIso
} else {
    $resolved = Join-Path (Get-Location) $OutputIso
    Write-ColorOutput "Resolved OutputIso: $OutputIso -> $resolved" -Color "Cyan" -CurrentIndent 1 -InheritedIndent 0
    $resolved
}

# Resolve AutounattendXml
$AutounattendXml = if ([System.IO.Path]::IsPathRooted($AutounattendXml)) {
    $AutounattendXml
} else {
    $resolved = Join-Path (Get-Location) $AutounattendXml
    Write-ColorOutput "Resolved AutounattendXml: $AutounattendXml -> $resolved" -Color "Cyan" -CurrentIndent 1 -InheritedIndent 0
    $resolved
}

# Resolve OemDirectory
$OemDirectory = if ([System.IO.Path]::IsPathRooted($OemDirectory)) {
    $OemDirectory
} else {
    $resolved = Join-Path (Get-Location) $OemDirectory
    Write-ColorOutput "Resolved OemDirectory: $OemDirectory -> $resolved" -Color "Cyan" -CurrentIndent 1 -InheritedIndent 0
    $resolved
}

# Resolve WorkingDirectory
$WorkingDirectory = if ([System.IO.Path]::IsPathRooted($WorkingDirectory)) {
    $WorkingDirectory
} else {
    $resolved = Join-Path (Get-Location) $WorkingDirectory
    Write-ColorOutput "Resolved WorkingDirectory: $WorkingDirectory -> $resolved" -Color "Cyan" -CurrentIndent 1 -InheritedIndent 0
    $resolved
}

# Resolve VirtioCacheDirectory
$VirtioCacheDirectory = if ([System.IO.Path]::IsPathRooted($VirtioCacheDirectory)) {
    $VirtioCacheDirectory
} else {
    $resolved = Join-Path (Get-Location) $VirtioCacheDirectory
    Write-ColorOutput "Resolved VirtioCacheDirectory: $VirtioCacheDirectory -> $resolved" -Color "Cyan" -CurrentIndent 1 -InheritedIndent 0
    $resolved
}

try {
    Write-ColorOutput "=== Windows ISO Repack Script ===" -Color "Cyan"
    Write-ColorOutput "Checking administrator privileges..." -Color "Yellow"
    
    if (-not (Test-Administrator)) {
        throw "This script must be run as Administrator!"
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
    
    # Paths are already resolved to absolute paths at the beginning of the script
    $resolvedInputIso = $InputIso
    $resolvedOutputIso = $OutputIso
    
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
            throw "Output ISO file already exists: $resolvedOutputIso. Use -OverwriteOutputIso parameter to overwrite the existing file."
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
        Write-ColorOutput "File size: $fileSizeGB GB" -Color "Green" -CurrentIndent 1 -InheritedIndent 0
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
