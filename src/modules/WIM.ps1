# WIM analysis and information extraction for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path $PSScriptRoot "Common.ps1"
. $commonPath

# Load DISM module dependency
$toolsPath = Join-Path $PSScriptRoot "tools"
. (Join-Path $toolsPath "DISM.ps1")

function Get-WimImageDetails {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WimPath,
        [Parameter(Mandatory = $true)]
        [int]$ImageIndex,
        [Parameter(Mandatory = $true)]
        [string]$DismPath,
        [Parameter(Mandatory = $false)]
        [bool]$ShowDebugOutput = $true,
        [Parameter(Mandatory = $false)]
        [int]$InheritedIndent = 0
    )
    
    # Get detailed image information using DISM /Get-WimInfo with specific index
    $result = Start-Process -FilePath $DismPath -ArgumentList @(
        "/Get-WimInfo",
        "/WimFile:`"$WimPath`"",
        "/Index:$ImageIndex"
    ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_image_details.txt"
    
    if ($result.ExitCode -ne 0) {
        $errorOutput = Get-Content "temp_image_details.txt" -ErrorAction SilentlyContinue
        Remove-Item "temp_image_details.txt" -ErrorAction SilentlyContinue
        throw "DISM failed with exit code $($result.ExitCode) for image index $ImageIndex. Error output: $($errorOutput -join ' ')"
    }
    
    $imageDetails = Get-Content "temp_image_details.txt" -ErrorAction SilentlyContinue
    
    # Check if DISM output is empty
    if (-not $imageDetails -or ($imageDetails | Where-Object { $_.Trim() -ne "" }).Count -eq 0) {
        Remove-Item "temp_image_details.txt" -ErrorAction SilentlyContinue
        throw "DISM returned empty output for image index $ImageIndex in WIM file: $WimPath"
    }
    
    # Debug: Show the image-specific DISM output if requested
    if ($ShowDebugOutput) {
        Write-ColorOutput "DISM image details for index $($ImageIndex):" -Color "Cyan" -Indent 2 -InheritedIndent $InheritedIndent
        foreach ($line in $imageDetails) {
            if ($line.Trim() -ne "") {
                Write-ColorOutput "  $line" -Color "Gray" -Indent 0 -InheritedIndent ($InheritedIndent + 2)
            }
        }
        Write-Host ""
    }
    
    # Clean up the temporary file
    Remove-Item "temp_image_details.txt" -ErrorAction SilentlyContinue
    
    # Parse the output into a structured format
    $parsedData = @{}
    foreach ($line in $imageDetails) {
        if ($line -match "^\s*(\w+)\s*:\s*(.+)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $parsedData[$key] = $value
        }
    }
    
    return @{
        RawOutput = $imageDetails
        ParsedData = $parsedData
    }
}


function Get-WimImageInfo {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "WIM file does not exist: $_"
            }
            if ($_ -notmatch '\.wim$') {
                throw "File must have .wim extension: $_"
            }
            $true
        })]
        [string]$WimPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "DISM executable does not exist: $_"
            }
            $true
        })]
        [string]$DismPath,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$InheritedIndent = 0
    )
    
    try {
        Write-ColorOutput "Getting WIM image information from: $WimPath" -Color "Yellow" -Indent 0 -InheritedIndent $InheritedIndent
        
        # Use DISM to get image information
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Get-WimInfo",
            "/WimFile:`"$WimPath`""
        ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_wim_info.txt"
        
        if ($result.ExitCode -ne 0) {
            throw "Failed to get WIM image information. DISM exit code: $($result.ExitCode)"
        }
        
        # Parse the output to extract image information
        $wimInfo = Get-Content "temp_wim_info.txt" -ErrorAction SilentlyContinue
        
        # Check if DISM output is empty or contains only whitespace
        if (-not $wimInfo -or ($wimInfo | Where-Object { $_.Trim() -ne "" }).Count -eq 0) {
            Write-ColorOutput "Warning: DISM returned empty output for WIM file: $WimPath" -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
            Write-ColorOutput "This may indicate the WIM file is corrupted or inaccessible" -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
            Remove-Item "temp_wim_info.txt" -ErrorAction SilentlyContinue
            return @()
        }
        
        # Debug: Log the raw DISM output for troubleshooting
        Write-ColorOutput "Raw DISM output from temp_wim_info.txt:" -Color "Cyan" -Indent 1 -InheritedIndent $InheritedIndent
        Write-Host ""
        foreach ($line in $wimInfo) {
            if ($line.Trim() -ne "") {
                Write-ColorOutput "  $line" -Color "Gray" -Indent 0 -InheritedIndent ($InheritedIndent + 1)
            }
        }
        Write-Host ""
        
        # Clean up the temporary file
        Remove-Item "temp_wim_info.txt" -ErrorAction SilentlyContinue
        
        $images = @()
        $currentImage = $null
        
        foreach ($line in $wimInfo) {
            if ($line -match "Index\s*:\s*(\d+)") {
                if ($currentImage) {
                    $images += $currentImage
                }
                $currentImage = @{
                    Index = [int]$matches[1]
                    Name = ""
                    Description = ""
                    Architecture = ""
                    Version = ""
                }
            } elseif ($currentImage -and $line -match "Name\s*:\s*(.+)") {
                $currentImage.Name = $matches[1].Trim()
                
                # Detect version from name
                if ($currentImage.Name -match "Windows 11" -or $currentImage.Name -match "Windows 1[1-9]") {
                    $currentImage.Version = "w11"
                } elseif ($currentImage.Name -match "Windows 10") {
                    $currentImage.Version = "w10"
                }
            } elseif ($currentImage -and $line -match "Description\s*:\s*(.+)") {
                $currentImage.Description = $matches[1].Trim()
                
                # Detect version from description if not already found
                if (-not $currentImage.Version) {
                    if ($currentImage.Description -match "Windows 11" -or $currentImage.Description -match "Windows 1[1-9]") {
                        $currentImage.Version = "w11"
                    } elseif ($currentImage.Description -match "Windows 10") {
                        $currentImage.Version = "w10"
                    }
                }
            } elseif ($line -match "Architecture\s*:\s*(.+)") {
                $arch = $matches[1].Trim()
                
                # Map DISM architecture names to our expected values
                switch ($arch.ToLower()) {
                    "x64" { $arch = "amd64" }
                    "x86" { $arch = "x86" }
                    "arm64" { $arch = "arm64" }
                    default {
                        Write-ColorOutput "Warning: Unknown architecture detected: $arch, defaulting to amd64" -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
                        $arch = "amd64"
                    }
                }
                
                # Set architecture for current image if we have one
                if ($currentImage) {
                    $currentImage.Architecture = $arch
                }
            }
        }
        
        if ($currentImage) {
            $images += $currentImage
        }
        
        # Post-process images to add fallback architecture detection
        foreach ($image in $images) {
            if (-not $image.Architecture) {
                Write-ColorOutput "No architecture found for image $($image.Index), attempting fallback detection..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
                
                # Get architecture using DISM
                $imageDetails = Get-WimImageDetails -WimPath $WimPath -ImageIndex $image.Index -DismPath $DismPath -ShowDebugOutput $true -InheritedIndent $InheritedIndent
                
                if ($imageDetails.ParsedData.ContainsKey("Architecture")) {
                    $arch = $imageDetails.ParsedData["Architecture"].ToLower()
                    switch ($arch) {
                        "x64" { $detectedArch = "amd64" }
                        "x86" { $detectedArch = "x86" }
                        "arm64" { $detectedArch = "arm64" }
                        default { 
                            throw "Unknown architecture '$arch' found in DISM output for image index $($image.Index). Expected: x64, x86, or arm64"
                        }
                    }
                    Write-ColorOutput "Architecture detected via DISM: $detectedArch" -Color "Green" -Indent 1 -InheritedIndent $InheritedIndent
                    $image.Architecture = $detectedArch
                } else {
                    throw "Architecture information not found in DISM output for image index $($image.Index)"
                }
            }
        }
        
        return $images
        
    } catch {
        throw "Failed to get WIM image information: $($_.Exception.Message)"
    }
}



function Get-AllWimInfo {
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
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$InheritedIndent = 0
    )
    
    Write-ColorOutput "=== Analyzing All WIM Files ===" -Color "Cyan" -Indent 0 -InheritedIndent $InheritedIndent
    
    Write-Host ""
    $wims = @()
    $installWimPath = Join-Path $ExtractPath "sources\install.wim"
    $bootWimPath = Join-Path $ExtractPath "sources\boot.wim"
    
    # Analyze install.wim if it exists
    if (Test-Path $installWimPath) {
        Write-ColorOutput "Analyzing install.wim..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
        $installWimInfo = Get-WimImageInfo -WimPath $installWimPath -DismPath (Get-DismPath) -InheritedIndent ($InheritedIndent + 1)
        
        if ($installWimInfo) {
            foreach ($image in $installWimInfo) {
                $wimInfo = @{
                    Path = $installWimPath
                    Type = "install"
                    Index = $image.Index
                    Name = $image.Name
                    Description = $image.Description
                    Architecture = $image.Architecture
                    Version = $image.Version
                }
                
                $wims += $wimInfo
                Write-ColorOutput "Found install image: $($image.Name) (Arch: $($image.Architecture), Version: $($image.Version))" -Color "Green" -Indent 2 -InheritedIndent $InheritedIndent
            }
        }
    }
    
    # Analyze boot.wim if it exists
    if (Test-Path $bootWimPath) {
        Write-ColorOutput "Analyzing boot.wim..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
        $bootWimInfo = Get-WimImageInfo -WimPath $bootWimPath -DismPath (Get-DismPath) -InheritedIndent ($InheritedIndent + 1)
        
        if ($bootWimInfo) {
            foreach ($image in $bootWimInfo) {
                $wimInfo = @{
                    Path = $bootWimPath
                    Type = "boot"
                    Index = $image.Index
                    Name = $image.Name
                    Description = $image.Description
                    Architecture = $image.Architecture
                    Version = $image.Version
                }
                
                $wims += $wimInfo
                Write-ColorOutput "Found boot image: $($image.Name) (Arch: $($image.Architecture), Version: $($image.Version))" -Color "Green" -Indent 2 -InheritedIndent $InheritedIndent
            }
        }
    }
    
    Write-Host ""
    if ($wims.Count -eq 0) {
        throw "No WIM files found in the ISO"
    }
    
    Write-ColorOutput "Found $($wims.Count) WIM image(s) total" -Color "Green" -Indent 0 -InheritedIndent $InheritedIndent
    return $wims
}

