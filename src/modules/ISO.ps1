# ISO operations for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot) "Common.ps1"
. $commonPath

function Extract-IsoContents {
    param(
        [string]$IsoPath,
        [string]$ExtractPath
    )
    
    Write-ColorOutput "Mounting ISO: $IsoPath" "Yellow"
    $mountResult = Mount-DiskImage -ImagePath $IsoPath -PassThru
    $driveLetter = ($mountResult | Get-Volume).DriveLetter
    
    if (-not $driveLetter) {
        throw "Failed to mount ISO or get drive letter"
    }
    
    $mountedPath = "${driveLetter}:\"
    Write-ColorOutput "ISO mounted at: $mountedPath" "Green" -Indent 1
    
    try {
        Write-ColorOutput "Extracting ISO contents to: $ExtractPath" "Yellow" -Indent 1
        
        if (Test-Path $ExtractPath) {
            Remove-Item $ExtractPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null
        
        robocopy $mountedPath $ExtractPath /E /COPY:DT /R:3 /W:10 /NFL /NDL /NJH /NJS /nc /ns /np
        
        if ($LASTEXITCODE -gt 7) {
            throw "Failed to extract ISO contents. Robocopy exit code: $LASTEXITCODE"
        }

        Write-ColorOutput "ISO contents extracted successfully" "Green" -Indent 1
    } finally {
        Write-ColorOutput "Dismounting ISO..." "Yellow" -Indent 1
        Dismount-DiskImage -ImagePath $IsoPath
        Write-ColorOutput "ISO dismounted" "Green" -Indent 1
    }
}

function Add-AutounattendXml {
    param(
        [string]$ExtractPath,
        [string]$AutounattendXmlPath
    )
    Write-ColorOutput "Adding autounattend.xml to ISO contents..." "Yellow"
    $destinationPath = Join-Path $ExtractPath "autounattend.xml"
    Copy-Item $AutounattendXmlPath $destinationPath -Force
    Write-ColorOutput "autounattend.xml added to: $destinationPath" "Green" -Indent 1
}

function Add-OemDirectory {
    param(
        [string]$ExtractPath,
        [string]$OemSourcePath
    )
    Write-ColorOutput "Adding $OEM$ directory to ISO contents..." "Yellow"
    
    if (-not (Test-Path $OemSourcePath -PathType Container)) {
        Write-ColorOutput "Warning: $OEM$ directory not found at: $OemSourcePath" "Yellow" -Indent 1
        return
    }
    
    $destinationPath = Join-Path $ExtractPath '$OEM$'
    
    # Remove existing $OEM$ directory if it exists
    if (Test-Path $destinationPath) {
        Remove-Item $destinationPath -Recurse -Force
    }
    
    # Copy the entire $OEM$ directory structure
    Copy-Item $OemSourcePath $destinationPath -Recurse -Force
    Write-ColorOutput "$OEM$ directory added to: $destinationPath" "Green" -Indent 1
}

function New-IsoFromDirectory {
    param(
        [string]$SourcePath,
        [string]$OutputPath,
        [string]$OscdimgPath
    )
    Write-ColorOutput "Creating new ISO from directory: $SourcePath" "Yellow"

    # Resolve absolute paths
    $absSrc    = (Resolve-Path $SourcePath).ProviderPath
    $absOutDir = (Resolve-Path (Split-Path $OutputPath -Parent)).ProviderPath
    $absOutIso = Join-Path $absOutDir (Split-Path $OutputPath -Leaf)

    $etfsbootPath  = "$absSrc\boot\etfsboot.com"
    $efisysPath    = "$absSrc\efi\microsoft\boot\efisys.bin"

    Write-ColorOutput "Using source directly: $absSrc" "Cyan" -Indent 1

    $arguments = @(
        "-m"
        "-u2"
        "-udfver102"
        "-bootdata:2#p0,e,b`"$etfsbootPath`"#pEF,e,b`"$efisysPath`""
        "`"$absSrc`""
        "`"$absOutIso`""
    )

    Write-ColorOutput "Current working directory: $(Get-Location)" "Cyan" -Indent 1
    Write-ColorOutput "Running oscdimg with arguments: $($arguments -join ' ')" "Cyan" -Indent 1
    Write-ColorOutput "Full command: & `"$OscdimgPath`" $($arguments -join ' ')" "Cyan" -Indent 1
    
    & $OscdimgPath $arguments
    if ($LASTEXITCODE -ne 0) { throw "oscdimg failed with exit code: $LASTEXITCODE" }

    Write-ColorOutput "ISO created successfully: $absOutIso" "Green" -Indent 1
}
