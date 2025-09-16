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
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "WIM file does not exist: $_"
            }
            $true
        })]
        [string]$WimPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if ($_ -le 0) {
                throw "ImageIndex must be a positive number"
            }
            $true
        })]
        [int]$ImageIndex,
        
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
        [bool]$ShowDebugOutput = $true,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$InheritedIndent = 0
    )
    
    # Get detailed image information using DISM /Get-WimInfo with specific index
    try {
        Invoke-CommandWithExitCode -Command $DismPath -Arguments @("/Get-WimInfo", "/WimFile:$WimPath", "/Index:$ImageIndex") -Description "Get detailed WIM image information for index $ImageIndex" -OutputFile "temp_image_details.txt"
    } catch {
        $errorOutput = Get-Content "temp_image_details.txt" -ErrorAction SilentlyContinue
        Remove-Item "temp_image_details.txt" -ErrorAction SilentlyContinue
        throw "DISM failed for image index $ImageIndex. Error output: $($errorOutput -join ' ') $_"
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
            $matches = Assert-Defined -VariableName "matches" -Value $matches -ErrorMessage "Regex match failed unexpectedly"
            $key = Assert-NotEmpty -VariableName "matches[1]" -Value $matches[1].Trim() -ErrorMessage "Regex match group 1 is empty"
            $value = Assert-NotEmpty -VariableName "matches[2]" -Value $matches[2].Trim() -ErrorMessage "Regex match group 2 is empty"
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
        [ValidatePattern('\.wim$')]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "WIM file does not exist: $_"
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
        
        # Use DISM to get image information for all images
        Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Get-WimInfo", "/WimFile:$WimPath") -Description "Get WIM image information" -OutputFile "temp_wim_info.txt"
        
        # Parse the output to extract image information
        $wimInfo = Get-Content "temp_wim_info.txt" -ErrorAction SilentlyContinue
        $wimInfo = Assert-Defined -VariableName "wimInfo" -Value $wimInfo -ErrorMessage "Failed to read DISM output from temp_wim_info.txt"
        
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
                $matches = Assert-Defined -VariableName "matches" -Value $matches -ErrorMessage "Index regex match failed unexpectedly"
                $indexValue = Assert-NotEmpty -VariableName "matches[1]" -Value $matches[1] -ErrorMessage "Index regex match group 1 is empty"
                $currentImage = @{
                    Index = [int]$indexValue
                }
            } elseif ($currentImage -and $line -match "^\s*(\w+(?:\s+\w+)*)\s*:\s*(.+)") {
                $matches = Assert-Defined -VariableName "matches" -Value $matches -ErrorMessage "Field regex match failed unexpectedly"
                $fieldName = Assert-NotEmpty -VariableName "matches[1]" -Value $matches[1].Trim() -ErrorMessage "Field name regex match group 1 is empty"
                $fieldValue = Assert-NotEmpty -VariableName "matches[2]" -Value $matches[2].Trim() -ErrorMessage "Field value regex match group 2 is empty"
                
                # Store field directly with original DISM name
                $currentImage[$fieldName] = $fieldValue
            }
        }
        
        if ($currentImage) {
            $images += $currentImage
        }
        
        # Now get detailed information for each image
        $detailedImages = @()
        foreach ($basicImage in $images) {
            Write-ColorOutput "Getting detailed info for image index $($basicImage.Index)..." -Color "Cyan" -Indent 1 -InheritedIndent $InheritedIndent
            
            try {
                $detailedInfo = Get-WimImageDetails -WimPath $WimPath -ImageIndex $basicImage.Index -DismPath $DismPath -ShowDebugOutput $true -InheritedIndent $InheritedIndent
                
                # Start with all detailed DISM fields from the image
                $detailedInfo = Assert-Defined -VariableName "detailedInfo" -Value $detailedInfo -ErrorMessage "Failed to get detailed image information"
                $detailedInfo.ParsedData = Assert-Defined -VariableName "detailedInfo.ParsedData" -Value $detailedInfo.ParsedData -ErrorMessage "Parsed data is not available from detailed image information"
                $detailedImage = $detailedInfo.ParsedData.Clone()
                
                # Ensure Index is preserved
                $basicImage.Index = Assert-PositiveNumber -VariableName "basicImage.Index" -Value $basicImage.Index -ErrorMessage "Basic image index must be a positive number"
                $detailedImage.Index = $basicImage.Index
                
                $detailedImages += $detailedImage
            } catch {
                Write-ColorOutput "Warning: Failed to get detailed info for image index $($basicImage.Index): $($_.Exception.Message)" -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
                # Fall back to basic image info if detailed info fails
                $detailedImages += $basicImage
            }
        }
        
        return $detailedImages
        
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
                # Validate image properties before using them
                $image = Assert-Defined -VariableName "image" -Value $image -ErrorMessage "Boot WIM image is null"
                $imageName = Assert-NotEmpty -VariableName "image.Name" -Value $image.Name -ErrorMessage "Boot WIM image name is not defined"
                $imageArch = Assert-NotEmpty -VariableName "image.Architecture" -Value $image.Architecture -ErrorMessage "Boot WIM image architecture is not defined"
                $imageVersion = Assert-NotEmpty -VariableName "image.Version" -Value $image.Version -ErrorMessage "Boot WIM image version is not defined"
                
                # Start with all detailed DISM fields from the image
                $wimInfo = $image.Clone()
                
                # Add our custom fields
                $wimInfo.Path = $bootWimPath
                $wimInfo.Type = "boot"
                
                $wims += $wimInfo
                Write-ColorOutput "Found boot image: $imageName (Arch: $imageArch, Version: $imageVersion)" -Color "Green" -Indent 2 -InheritedIndent $InheritedIndent
            }
        }
    }
    
    # Analyze install.wim if it exists
    if (Test-Path $installWimPath) {
        Write-ColorOutput "Analyzing install.wim..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
        $installWimInfo = Get-WimImageInfo -WimPath $installWimPath -DismPath (Get-DismPath) -InheritedIndent ($InheritedIndent + 1)
        
        if ($installWimInfo) {
            foreach ($image in $installWimInfo) {
                # Validate image properties before using them
                $image = Assert-Defined -VariableName "image" -Value $image -ErrorMessage "Install WIM image is null"
                $imageName = Assert-NotEmpty -VariableName "image.Name" -Value $image.Name -ErrorMessage "Install WIM image name is not defined"
                $imageArch = Assert-NotEmpty -VariableName "image.Architecture" -Value $image.Architecture -ErrorMessage "Install WIM image architecture is not defined"
                $imageVersion = Assert-NotEmpty -VariableName "image.Version" -Value $image.Version -ErrorMessage "Install WIM image version is not defined"
                
                # Start with all detailed DISM fields from the image
                $wimInfo = $image.Clone()
                
                # Add our custom fields
                $wimInfo.Path = $installWimPath
                $wimInfo.Type = "install"
                
                $wims += $wimInfo
                Write-ColorOutput "Found install image: $imageName (Arch: $imageArch, Version: $imageVersion)" -Color "Green" -Indent 2 -InheritedIndent $InheritedIndent
            }
        }
    }
    
    Write-Host ""
    $wims = Assert-ArrayNotEmpty -VariableName "wims" -Value $wims -ErrorMessage "No WIM files found in the ISO"
    
    Write-ColorOutput "Found $($wims.Count) WIM image(s) total" -Color "Green" -Indent 0 -InheritedIndent $InheritedIndent
    return $wims
}

function Filter-InstallWimImages {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Container)) {
                throw "Extract path does not exist: $_"
            }
            $true
        })]
        [string]$ExtractPath,
        [Parameter(Mandatory = $false)]
        [string[]]$IncludeTargets = $null,
        [Parameter(Mandatory = $false)]
        [string[]]$ExcludeTargets = $null
    )
    
    $installWimPath = Join-Path $ExtractPath "sources\install.wim"
    
    # Validate install.wim exists
    try {
        $installWimPath = Assert-FileExists -FilePath $installWimPath -ErrorMessage "No install.wim found at: $installWimPath"
    } catch {
        Write-ColorOutput "No install.wim found at: $installWimPath" -Color "Yellow" -Indent 1
        return
    }
    
    Write-ColorOutput "Filtering install.wim images..." -Color "Yellow" -Indent 1
    
    # Get current WIM information
    $dismPath = Get-DismPath
    $wimInfo = Get-WimImageInfo -WimPath $installWimPath -DismPath $dismPath -InheritedIndent 1
    
    # Validate WIM has images
    try {
        $wimInfo = Assert-ArrayNotEmpty -VariableName "wimInfo" -Value $wimInfo -ErrorMessage "No images found in install.wim"
    } catch {
        Write-ColorOutput "No images found in install.wim" -Color "Yellow" -Indent 1
        return
    }
    
    # Determine which images to keep
    $imagesToKeep = @()
    $imagesToRemove = @()
    
    foreach ($image in $wimInfo) {
        # Validate image properties before using them
        $image = Assert-Defined -VariableName "image" -Value $image -ErrorMessage "WIM image is null"
        $imageName = Assert-NotEmpty -VariableName "image.Name" -Value $image.Name -ErrorMessage "WIM image name is not defined"
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
    
    # Validate filtering results
    $imagesToKeep = Assert-ArrayNotEmpty -VariableName "imagesToKeep" -Value $imagesToKeep -ErrorMessage "No install.wim images would be kept after filtering. Cannot create ISO with no install images."
    
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
            # Validate image data
            $sourceIndex = Assert-PositiveNumber -VariableName "imageToKeep.Index" -Value $imageToKeep.Index -ErrorMessage "Image index must be a positive number"
            $imageName = Assert-NotEmpty -VariableName "imageToKeep.Name" -Value $imageToKeep.Name -ErrorMessage "Image name cannot be empty"
            
            Write-ColorOutput "Adding image $($newIndex): $imageName" -Color "Cyan" -Indent 2
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
            
            # Build DISM arguments
            $dismArgs = @(
                "/Export-Image",
                "/SourceImageFile:`"$installWimPath`"",
                "/SourceIndex:$sourceIndex",
                "/DestinationImageFile:`"$tempWimPath`"",
                "/DestinationName:`"$imageName`""
            )
            
            if ($newIndex -gt 1) {
                $dismArgs += "/Compress:maximum"
            }
            
            Write-ColorOutput "DISM Command: $dismPath $($dismArgs -join ' ')" -Color "Gray" -Indent 3
            
            # Execute DISM command
            try {
                Invoke-CommandWithExitCode -Command $dismPath -Arguments $dismArgs -Description "Export image '$imageName' (index $sourceIndex)" -SuppressOutput
            } catch {
                $errorDetails = $_
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

