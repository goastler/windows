# ISO operations for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path $PSScriptRoot "Common.ps1"
. $commonPath

function Extract-IsoContents {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "ISO file does not exist: $_"
            }
            $true
        })]
        [string]$IsoPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            # Validate that the path is a valid path format, but allow non-existent directories
            try {
                $null = [System.IO.Path]::GetFullPath($_)
                $true
            } catch {
                throw "Extract path is not a valid path format: $_"
            }
        })]
        [string]$ExtractPath
    )
    
    Write-ColorOutput "Mounting ISO: $IsoPath" -Color "Yellow"
    
    $mountResult = Mount-DiskImage -ImagePath $IsoPath -PassThru
    $mountResult = Assert-Defined -VariableName "mountResult" -Value $mountResult -ErrorMessage "Failed to mount ISO"
    
    $volume = $mountResult | Get-Volume
    $volume = Assert-Defined -VariableName "volume" -Value $volume -ErrorMessage "Failed to get volume information from mounted ISO"
    
    $driveLetter = $volume.DriveLetter
    $driveLetter = Assert-NotEmpty -VariableName "driveLetter" -Value $driveLetter -ErrorMessage "Failed to get drive letter from mounted ISO"
    
    $mountedPath = "${driveLetter}:\"
    $mountedPath = Assert-ValidPath -VariableName "mountedPath" -Path $mountedPath -ErrorMessage "Generated mounted path is invalid: $mountedPath"
    Write-ColorOutput "ISO mounted at: $mountedPath" -Color "Green" -Indent 1
    
    Write-Host ""
    try {
        Write-ColorOutput "Extracting ISO contents to: $ExtractPath" -Color "Yellow" -Indent 1         
        if (Test-Path $ExtractPath) {
            Remove-Item $ExtractPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null
        
        robocopy $mountedPath $ExtractPath /E /COPY:DT /R:3 /W:10 /NFL /NDL /NJH /NJS /nc /ns /np
        
        if ($LASTEXITCODE -gt 7) {
            throw "Failed to extract ISO contents. Robocopy exit code: $LASTEXITCODE"
        }

        Write-ColorOutput "ISO contents extracted successfully" -Color "Green" -Indent 1
        
        Write-Host ""
    } finally {
        Write-ColorOutput "Dismounting ISO..." -Color "Yellow" -Indent 1
        Dismount-DiskImage -ImagePath $IsoPath
        Write-ColorOutput "ISO dismounted" -Color "Green" -Indent 1
    }
}

function Add-AutounattendXml {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Container)) {
                throw "Extract path does not exist: $_"
            }
            $true
        })]
        [string]$ExtractPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "Autounattend.xml file does not exist: $_"
            }
            $true
        })]
        [string]$AutounattendXmlPath
    )
    
    Write-ColorOutput "Adding autounattend.xml to ISO contents..." -Color "Yellow"
    $destinationPath = Join-Path $ExtractPath "autounattend.xml"
    Copy-Item $AutounattendXmlPath $destinationPath -Force
    Write-ColorOutput "autounattend.xml added to: $destinationPath" -Color "Green" -Indent 1
}

function Add-OemDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Container)) {
                throw "Extract path does not exist: $_"
            }
            $true
        })]
        [string]$ExtractPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Container)) {
                throw "$OEM$ directory not found at: $_"
            }
            $true
        })]
        [string]$OemSourcePath
    )
    
    Write-ColorOutput "Adding $OEM$ directory to ISO contents..." -Color "Yellow"
    
    $destinationPath = Join-Path $ExtractPath '$OEM$'
    
    # Remove existing $OEM$ directory if it exists
    if (Test-Path $destinationPath) {
        Remove-Item $destinationPath -Recurse -Force
    }
    
    # Copy the entire $OEM$ directory structure
    Copy-Item $OemSourcePath $destinationPath -Recurse -Force
    Write-ColorOutput "$OEM$ directory added to: $destinationPath" -Color "Green" -Indent 1
}

function New-IsoFromDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Container)) {
                throw "Source directory does not exist: $_"
            }
            $true
        })]
        [string]$SourcePath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            try {
                $resolvedPath = Resolve-Path $_ -ErrorAction Stop
                $true
            } catch {
                throw "Output path is invalid: $_"
            }
        })]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "OSCDIMG executable does not exist: $_"
            }
            $true
        })]
        [string]$OscdimgPath
    )
    
    Write-ColorOutput "Creating new ISO from directory: $SourcePath" -Color "Yellow"

    # Resolve absolute paths
    $absSrc    = (Resolve-Path $SourcePath).ProviderPath
    $absSrc = Assert-ValidPath -VariableName "absSrc" -Path $absSrc -ErrorMessage "Failed to resolve source path: $SourcePath"
    
    $absOutDir = (Resolve-Path (Split-Path $OutputPath -Parent)).ProviderPath
    $absOutDir = Assert-ValidPath -VariableName "absOutDir" -Path $absOutDir -ErrorMessage "Failed to resolve output directory path: $OutputPath"
    
    $absOutIso = Join-Path $absOutDir (Split-Path $OutputPath -Leaf)
    $absOutIso = Assert-ValidPath -VariableName "absOutIso" -Path $absOutIso -ErrorMessage "Generated output ISO path is invalid: $absOutIso"

    $etfsbootPath  = "$absSrc\boot\etfsboot.com"
    $etfsbootPath = Assert-ValidPath -VariableName "etfsbootPath" -Path $etfsbootPath -ErrorMessage "Generated etfsboot path is invalid: $etfsbootPath"
    
    $efisysPath    = "$absSrc\efi\microsoft\boot\efisys.bin"
    $efisysPath = Assert-ValidPath -VariableName "efisysPath" -Path $efisysPath -ErrorMessage "Generated efisys path is invalid: $efisysPath"

    Write-ColorOutput "Using source directly: $absSrc" -Color "Cyan" -Indent 1 
    $arguments = @(
        "-m"
        "-u2"
        "-udfver102"
        "-bootdata:2#p0,e,b`"$etfsbootPath`"#pEF,e,b`"$efisysPath`""
        "`"$absSrc`""
        "`"$absOutIso`""
    )

    Write-ColorOutput "Current working directory: $(Get-Location)" -Color "Cyan" -Indent 1
    Invoke-CommandWithExitCode -Command $OscdimgPath -Arguments $arguments -Description "Create ISO with oscdimg" -WorkingDirectory (Get-Location)

    Write-ColorOutput "ISO created successfully: $absOutIso" -Color "Green" -Indent 1
}
