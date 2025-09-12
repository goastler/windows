# #Requires -RunAsAdministrator

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
    [switch]$KeepWorkingDirectory

)

$ErrorActionPreference = "Stop"

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-RequiredTools {
    Write-ColorOutput "Installing Windows ADK via Chocolatey..." "Yellow"
    
    # Always install Windows ADK via Chocolatey (it won't reinstall if already present)
    Install-WindowsADK
    
    # After installation, find oscdimg.exe
    $oscdimgPath = Get-Command "oscdimg.exe" -ErrorAction SilentlyContinue
    if ($oscdimgPath) {
        $script:oscdimgPath = $oscdimgPath.Source
        Write-ColorOutput "Found oscdimg.exe at: $script:oscdimgPath" "Green"
        return
    }
    
    throw "oscdimg.exe not found after Windows ADK installation. Please check the installation."
}

function Test-Chocolatey {
    return (Get-Command "choco" -ErrorAction SilentlyContinue) -ne $null
}

function Install-Chocolatey {

    if (Test-Chocolatey) {
        Write-ColorOutput "[OK] Chocolatey already installed!" "Green"
        return
    }

    Write-ColorOutput "Installing Chocolatey package manager..." "Yellow"
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = `
        [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    
    Start-Sleep -Seconds 1

    if (Test-Chocolatey) {
        Write-ColorOutput "[OK] Chocolatey installed successfully!" "Green"
        return
    } else {
        throw "Chocolatey installation failed to become available on PATH."            
    }
}

function Install-WindowsADK {
    Write-ColorOutput "=== Windows ADK Installation ===" "Cyan"

    Install-Chocolatey

    Write-ColorOutput "Installing Windows ADK via Chocolatey..." "Yellow"

    $result = Start-Process -FilePath "choco" -ArgumentList @("install","windows-adk","-y") -Wait -PassThru -NoNewWindow
    if ($result.ExitCode -eq 0) {
        Write-ColorOutput "[OK] Windows ADK installed successfully via Chocolatey!" "Green"
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + `
                    [System.Environment]::GetEnvironmentVariable('Path','User')
        return
    }

    throw "Windows ADK installation failed"
}

# function Extract-IsoContents {
#     param(
#         [string]$IsoPath,
#         [string]$ExtractPath
#     )
    
#     Write-ColorOutput "Mounting ISO: $IsoPath" "Yellow"
#     $mountResult = Mount-DiskImage -ImagePath $IsoPath -PassThru
#     $driveLetter = ($mountResult | Get-Volume).DriveLetter
    
#     if (-not $driveLetter) {
#         throw "Failed to mount ISO or get drive letter"
#     }
    
#     $mountedPath = "${driveLetter}:\"
#     Write-ColorOutput "ISO mounted at: $mountedPath" "Green"
    
#     try {
#         Write-ColorOutput "Extracting ISO contents to: $ExtractPath" "Yellow"
        
#         if (Test-Path $ExtractPath) {
#             Remove-Item $ExtractPath -Recurse -Force
#         }
#         New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null
        
#         robocopy $mountedPath $ExtractPath /E /COPY:DAT /R:3 /W:10 /NFL /NDL /NJH /NJS /nc /ns /np
        
#         if ($LASTEXITCODE -gt 7) {
#             throw "Failed to extract ISO contents. Robocopy exit code: $LASTEXITCODE"
#         }
        
#         Write-ColorOutput "ISO contents extracted successfully" "Green"
#     } finally {
#         Write-ColorOutput "Dismounting ISO..." "Yellow"
#         Dismount-DiskImage -ImagePath $IsoPath
#         Write-ColorOutput "ISO dismounted" "Green"
#     }
# }

# function Add-AutounattendXml {
#     param(
#         [string]$ExtractPath,
#         [string]$AutounattendXmlPath
#     )
#     Write-ColorOutput "Adding autounattend.xml to ISO contents..." "Yellow"
#     $destinationPath = Join-Path $ExtractPath "autounattend.xml"
#     Copy-Item $AutounattendXmlPath $destinationPath -Force
#     Write-ColorOutput "autounattend.xml added to: $destinationPath" "Green"
# }

# function New-IsoFromDirectory {
#     param(
#         [string]$SourcePath,
#         [string]$OutputPath,
#         [string]$OscdimgPath
#     )
#     Write-ColorOutput "Creating new ISO from directory: $SourcePath" "Yellow"
#     $arguments = @(
#         "-m"
#         "-o"
#         "-u2"
#         "-udfver102"
#         "-l"
#         "Windows"
#         "`"$SourcePath`""
#         "`"$OutputPath`""
#     )
#     Write-ColorOutput "Running oscdimg with arguments: $($arguments -join ' ')" "Cyan"
#     $process = Start-Process -FilePath $OscdimgPath -ArgumentList $arguments -Wait -PassThru -NoNewWindow
#     if ($process.ExitCode -ne 0) {
#         throw "oscdimg failed with exit code: $($process.ExitCode)"
#     }
#     Write-ColorOutput "ISO created successfully: $OutputPath" "Green"
# }

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
    Write-ColorOutput "Working Directory: $WorkingDirectory" "White"
    
    Test-RequiredTools
    
    # Write-ColorOutput "Validating input files..." "Yellow"
    # if (-not (Test-Path $InputIso -PathType Leaf)) {
    #     throw "Input ISO file not found: $InputIso"
    # }
    # if (-not (Test-Path $AutounattendXml -PathType Leaf)) {
    #     throw "Autounattend XML file not found: $AutounattendXml"
    # }
    # Write-ColorOutput "Input files validated" "Green"
    
    # if (Test-Path $OutputIso) {
    #     Write-ColorOutput "Output ISO already exists. Removing..." "Yellow"
    #     Remove-Item $OutputIso -Force
    # }
    
    # Extract-IsoContents -IsoPath $InputIso -ExtractPath $WorkingDirectory
    # Add-AutounattendXml -ExtractPath $WorkingDirectory -AutounattendXmlPath $AutounattendXml
    # New-IsoFromDirectory -SourcePath $WorkingDirectory -OutputPath $OutputIso -OscdimgPath $script:oscdimgPath
    
    # if (Test-Path $OutputIso) {
    #     $fileSize = (Get-Item $OutputIso).Length
    #     $fileSizeGB = [math]::Round($fileSize / 1GB, 2)
    #     Write-ColorOutput "Output ISO created successfully!" "Green"
    #     Write-ColorOutput "File size: $fileSizeGB GB" "Green"
    # } else {
    #     throw "Output ISO was not created successfully"
    # }
} catch {
    Write-ColorOutput "Error: $($_.Exception.Message)" "Red"
    exit 1
} finally {
    Remove-WorkingDirectory -Path $WorkingDirectory
}

Write-ColorOutput "=== Script completed successfully! ===" "Green"
