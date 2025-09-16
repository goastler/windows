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
    & $DismPath /Get-WimInfo /WimFile:$WimPath /Index:$ImageIndex > temp_image_details.txt 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        $errorOutput = Get-Content "temp_image_details.txt" -ErrorAction SilentlyContinue
        Remove-Item "temp_image_details.txt" -ErrorAction SilentlyContinue
        throw "DISM failed with exit code $LASTEXITCODE for image index $ImageIndex. Error output: $($errorOutput -join ' ')"
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
        & $dismPath /Get-WimInfo /WimFile:$WimPath > temp_wim_info.txt 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to get WIM image information. DISM exit code: $LASTEXITCODE"
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
                }
            } elseif ($currentImage -and $line -match "^\s*(\w+(?:\s+\w+)*)\s*:\s*(.+)") {
                $fieldName = $matches[1].Trim()
                $fieldValue = $matches[2].Trim()
                
                # Store field directly with original DISM name
                $currentImage[$fieldName] = $fieldValue
            }
        }
        
        if ($currentImage) {
            $images += $currentImage
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
    
    # Analyze boot.wim if it exists
    if (Test-Path $bootWimPath) {
        Write-ColorOutput "Analyzing boot.wim..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
        $bootWimInfo = Get-WimImageInfo -WimPath $bootWimPath -DismPath (Get-DismPath) -InheritedIndent ($InheritedIndent + 1)
        
        if ($bootWimInfo) {
            foreach ($image in $bootWimInfo) {
                # Start with all DISM fields from the image
                $wimInfo = $image.Clone()
                
                # Add our custom fields
                $wimInfo.Path = $bootWimPath
                $wimInfo.Type = "boot"
                
                $wims += $wimInfo
                Write-ColorOutput "Found boot image: $($image.Name) (Arch: $($image.Architecture), Version: $($image.Version))" -Color "Green" -Indent 2 -InheritedIndent $InheritedIndent
            }
        }
    }
    
    # Analyze install.wim if it exists
    if (Test-Path $installWimPath) {
        Write-ColorOutput "Analyzing install.wim..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
        $installWimInfo = Get-WimImageInfo -WimPath $installWimPath -DismPath (Get-DismPath) -InheritedIndent ($InheritedIndent + 1)
        
        if ($installWimInfo) {
            foreach ($image in $installWimInfo) {
                # Start with all DISM fields from the image
                $wimInfo = $image.Clone()
                
                # Add our custom fields
                $wimInfo.Path = $installWimPath
                $wimInfo.Type = "install"
                
                $wims += $wimInfo
                Write-ColorOutput "Found install image: $($image.Name) (Arch: $($image.Architecture), Version: $($image.Version))" -Color "Green" -Indent 2 -InheritedIndent $InheritedIndent
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

function Filter-InstallWimImages {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExtractPath,
        [Parameter(Mandatory = $false)]
        [string[]]$IncludeTargets = $null,
        [Parameter(Mandatory = $false)]
        [string[]]$ExcludeTargets = $null
    )
    
    $installWimPath = Join-Path $ExtractPath "sources\install.wim"
    
    if (-not (Test-Path $installWimPath)) {
        Write-ColorOutput "No install.wim found at: $installWimPath" -Color "Yellow" -Indent 1
        return
    }
    
    Write-ColorOutput "Filtering install.wim images..." -Color "Yellow" -Indent 1
    
    # Get current WIM information
    $dismPath = Get-DismPath
    $wimInfo = Get-WimImageInfo -WimPath $installWimPath -DismPath $dismPath -InheritedIndent 1
    
    if (-not $wimInfo -or $wimInfo.Count -eq 0) {
        Write-ColorOutput "No images found in install.wim" -Color "Yellow" -Indent 1
        return
    }
    
    # Determine which images to keep
    $imagesToKeep = @()
    $imagesToRemove = @()
    
    foreach ($image in $wimInfo) {
        $imageName = $image.Name
        $shouldKeep = $true
        
        # Check include targets
        if ($IncludeTargets -and $IncludeTargets.Count -gt 0) {
            $matchesInclude = $false
            foreach ($includeTarget in $IncludeTargets) {
                if ($imageName -eq $includeTarget) {
                    $matchesInclude = $true
                    break
                }
            }
            if (-not $matchesInclude) {
                $shouldKeep = $false
            }
        }
        
        # Check exclude targets
        if ($ExcludeTargets -and $ExcludeTargets.Count -gt 0) {
            foreach ($excludeTarget in $ExcludeTargets) {
                if ($imageName -eq $excludeTarget) {
                    $shouldKeep = $false
                    break
                }
            }
        }
        
        if ($shouldKeep) {
            $imagesToKeep += $image
            Write-ColorOutput "Keeping: $imageName" -Color "Green" -Indent 2
        } else {
            $imagesToRemove += $image
            Write-ColorOutput "Removing: $imageName" -Color "Red" -Indent 2
        }
    }
    
    if ($imagesToKeep.Count -eq 0) {
        throw "No install.wim images would be kept after filtering. Cannot create ISO with no install images."
    }
    
    if ($imagesToRemove.Count -eq 0) {
        Write-ColorOutput "No images to remove" -Color "Green" -Indent 1
        return
    }
    
    Write-ColorOutput "Removing $($imagesToRemove.Count) image(s) from install.wim..." -Color "Yellow" -Indent 1
    
    # Create a new install.wim with only the images we want to keep
    $tempWimPath = Join-Path (Split-Path $installWimPath -Parent) "install_temp.wim"
    
    try {
        
        # Create new WIM with only the images we want to keep
        $newIndex = 1
        foreach ($imageToKeep in $imagesToKeep) {
            Write-ColorOutput "Adding image $($newIndex): $($imageToKeep.Name)" -Color "Cyan" -Indent 2
            
            # Build DISM command with proper quoting for image names that might contain spaces
            $sourceIndex = $imageToKeep.Index
            $imageName = $imageToKeep.Name
            $dismCommand = @(
                $dismPath,
                "/Export-Image",
                "/SourceImageFile:`"$installWimPath`"",
                "/SourceIndex:$sourceIndex",
                "/DestinationImageFile:`"$tempWimPath`"",
                "/DestinationName:`"$imageName`""
            )
            
            if ($newIndex -gt 1) {
                # Add compression for subsequent images
                $dismCommand += "/Compress:maximum"
            }
            
            Write-ColorOutput "DISM Command: $($dismCommand -join ' ')" -Color "Gray" -Indent 3
            
            # Execute DISM command
            & $dismPath /Export-Image /SourceImageFile:"$installWimPath" /SourceIndex:$sourceIndex /DestinationImageFile:"$tempWimPath" /DestinationName:"$imageName" @(if ($newIndex -gt 1) { "/Compress:maximum" })
            
            if ($LASTEXITCODE -ne 0) {
                $errorDetails = "DISM exit code: $LASTEXITCODE"
                if (Test-Path "C:\WINDOWS\Logs\DISM\dism.log") {
                    $logContent = Get-Content "C:\WINDOWS\Logs\DISM\dism.log" -Tail 10 -ErrorAction SilentlyContinue
                    if ($logContent) {
                        $errorDetails += ". Last 10 lines of DISM log:`n$($logContent -join "`n")"
                    }
                }
                throw "Failed to export image '$imageName' (index $sourceIndex) to new WIM file. $errorDetails"
            }
            
            $newIndex++
        }
        
        # Replace original with filtered WIM
        Remove-Item $installWimPath -Force
        Move-Item $tempWimPath $installWimPath
        
        Write-ColorOutput "Successfully filtered install.wim" -Color "Green" -Indent 1
        Write-ColorOutput "Kept $($imagesToKeep.Count) image(s), removed $($imagesToRemove.Count) image(s)" -Color "Green" -Indent 2
        
    } catch {
        
        # Clean up temp file
        if (Test-Path $tempWimPath) {
            Remove-Item $tempWimPath -Force
        }
        
        throw "Failed to filter install.wim: $($_.Exception.Message)"
    }
}

