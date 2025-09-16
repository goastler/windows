# Common utilities and helper functions for Windows ISO repack script

function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow", "White")]
        [string]$Color = "White",
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$Indent = 0,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$InheritedIndent = 0
    )
    
    # Additional validation for edge cases
    if ($Message.Trim().Length -eq 0) {
        throw "Message cannot be empty or contain only whitespace"
    }
    
    # Calculate total indentation (inherited + current)
    $totalIndent = $InheritedIndent + $Indent
    
    # Validate total indentation doesn't exceed reasonable limits
    if ($totalIndent -gt 40) {
        throw "Total indentation ($totalIndent) exceeds maximum allowed (40). Indent: $Indent, InheritedIndent: $InheritedIndent"
    }
    
    # Create indentation string (2 spaces per indent level)
    $indentString = "  " * $totalIndent
    
    # Combine indentation with message
    $indentedMessage = $indentString + $Message
    
    Write-Host $indentedMessage -ForegroundColor $Color
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Assert-Administrator {
    param(
        [string]$ErrorMessage = "Administrator privileges are required. Please run PowerShell as Administrator."
    )
    
    if (-not (Test-Administrator)) {
        throw $ErrorMessage
    }
}

function Invoke-WebRequestWithCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern('^https?://')]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            $parentDir = Split-Path $_ -Parent
            if ($parentDir -and -not (Test-Path $parentDir -PathType Container)) {
                throw "Output directory does not exist: $parentDir"
            }
            $true
        })]
        [string]$OutFile,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Description = "file",
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 10)]
        [int]$ProgressId = 3
    )
    
    Write-ColorOutput "Downloading $Description from: $Uri" -Color "White"     
    $webRequest = $null
    try {
        # Create a web client for progress tracking
        $webClient = New-Object System.Net.WebClient
        
        # Set up progress tracking
        $totalBytes = 0
        $downloadedBytes = 0
        
        # Register for download progress
        Register-ObjectEvent -InputObject $webClient -EventName "DownloadProgressChanged" -Action {
            $global:downloadProgress = $Event.SourceEventArgs
        } | Out-Null
        
        # Get file size first
        try {
            $headRequest = [System.Net.WebRequest]::Create($Uri)
            $headRequest = Assert-Defined -VariableName "headRequest" -Value $headRequest -ErrorMessage "Failed to create web request"
            $headRequest.Method = "HEAD"
            $response = $headRequest.GetResponse()
            $response = Assert-Defined -VariableName "response" -Value $response -ErrorMessage "Failed to get response from web request"
            $totalBytes = $response.ContentLength
            $response.Close()
        } catch {
            Write-ColorOutput "Could not determine file size, progress tracking may be limited" -Color "Yellow"
        }
        
        # Start download with progress tracking
        $webClient = Assert-Defined -VariableName "webClient" -Value $webClient -ErrorMessage "Web client is not initialized"
        $downloadTask = $webClient.DownloadFileTaskAsync($Uri, $OutFile)
        $downloadTask = Assert-Defined -VariableName "downloadTask" -Value $downloadTask -ErrorMessage "Failed to start download task"
        
        # Monitor progress
        while (-not $downloadTask.IsCompleted) {
            Start-Sleep -Milliseconds 100
            
            if ($global:downloadProgress) {
                $percentComplete = if ($totalBytes -gt 0) {
                    [math]::Round(($global:downloadProgress.BytesReceived / $totalBytes) * 100)
                } else {
                    [math]::Round(($global:downloadProgress.BytesReceived / ($global:downloadProgress.BytesReceived + 1)) * 100)
                }
                
                $downloadedMB = [math]::Round($global:downloadProgress.BytesReceived / 1MB, 2)
                $totalMB = if ($totalBytes -gt 0) { [math]::Round($totalBytes / 1MB, 2) } else { "Unknown" }
                
                Write-ProgressWithPercentage -Activity "Downloading $Description" -Status "Downloaded $downloadedMB MB of $totalMB MB" -PercentComplete $percentComplete -Id $ProgressId
            }
        }
        
        # Wait for completion and handle any exceptions
        $downloadTask.Wait()
        if ($downloadTask.Exception) {
            throw $downloadTask.Exception
        }
        
        Write-Progress -Activity "Downloading $Description" -Completed -Id $ProgressId
        Write-Host "" # Clear the progress line
        Write-ColorOutput "$Description downloaded successfully" -Color "Green"
    }
    finally {
        # Clean up event subscription
        try {
            Get-EventSubscriber | Where-Object { $_.SourceObject -eq $webClient } | Unregister-Event
        } catch {
            # Ignore cleanup errors
        }
        
        # Properly dispose of web client
        if ($webClient) {
            try {
                $webClient.Dispose()
            }
            catch {
                # Ignore disposal errors
            }
        }
        
        # Properly dispose of web request object
        if ($webRequest) {
            try {
                $webRequest.Dispose()
            }
            catch {
                # Ignore disposal errors
            }
        }
        
        # Force garbage collection to ensure resources are released
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Brief pause to ensure file handles are released
        Start-Sleep -Seconds 1
    }
}

function Remove-WorkingDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern('^([a-zA-Z]:\\|\\\\)')]
        [string]$Path
    )
    
    if (-not $KeepWorkingDirectory -and (Test-Path $Path)) {
        Write-ColorOutput "Cleaning up working directory: $Path" -Color "Yellow"
        Remove-Item $Path -Recurse -Force
        Write-ColorOutput "Working directory cleaned up" -Color "Green"
    } elseif ($KeepWorkingDirectory) {
        Write-ColorOutput "Keeping working directory: $Path" -Color "Cyan"
    }
}

# Command execution helper function
function Invoke-CommandWithExitCode {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Command,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Arguments = @(),
        
        [Parameter(Mandatory = $false)]
        [string]$WorkingDirectory,
        
        [Parameter(Mandatory = $false)]
        [string]$Description = "Command execution",
        
        [Parameter(Mandatory = $false)]
        [int[]]$ExpectedExitCodes = @(0),
        
        [Parameter(Mandatory = $false)]
        [switch]$SuppressOutput,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$InheritedIndent = 0
    )
    
    $Command = Assert-NotEmpty -VariableName "Command" -Value $Command -ErrorMessage "Command cannot be empty"
    
    Write-ColorOutput "Running: $Command $($Arguments -join ' ')" -Color "Cyan" -Indent 1 -InheritedIndent $InheritedIndent
    
    try {
        # Set working directory if specified
        $originalLocation = $null
        if ($WorkingDirectory) {
            $WorkingDirectory = Assert-ValidPath -VariableName "WorkingDirectory" -Path $WorkingDirectory -ErrorMessage "Working directory path is invalid: $WorkingDirectory"
            $originalLocation = Get-Location
            Set-Location $WorkingDirectory
        }
        
        # Execute the command with or without output redirection
        if ($OutputFile) {
            # Validate output file path format but allow non-existent files (they will be created)
            try {
                $null = [System.IO.Path]::GetFullPath($OutputFile)
            } catch {
                throw "Output file path is not a valid path format: $OutputFile"
            }
            & $Command @Arguments > $OutputFile 2>&1
        } elseif ($SuppressOutput) {
            & $Command @Arguments 2>&1 | Out-Null
        } else {
            & $Command @Arguments
        }
        
        $exitCode = $LASTEXITCODE
        $exitCode = Assert-Defined -VariableName "exitCode" -Value $exitCode -ErrorMessage "Failed to get exit code from command execution"
        
        # Check if exit code is expected
        if ($exitCode -notin $ExpectedExitCodes) {
            throw "$Description failed with exit code: $exitCode. Expected: $($ExpectedExitCodes -join ', ')"
        }
        
        Write-ColorOutput "$Description completed successfully (exit code: $exitCode)" -Color "Green" -Indent 1 -InheritedIndent $InheritedIndent
        
    } catch {
        throw "$Description failed: $_"
    } finally {
        # Restore original location if changed
        if ($originalLocation) {
            Set-Location $originalLocation
        }
    }
}

# Validation helper functions
function Assert-Defined {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VariableName,
        
        [Parameter(Mandatory = $true)]
        [object]$Value,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "Variable '$VariableName' is not defined or is null"
    )
    
    if ($null -eq $Value) {
        throw $ErrorMessage
    }
    
    return $Value
}

function Assert-NotEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VariableName,
        
        [Parameter(Mandatory = $true)]
        [string]$Value,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "Variable '$VariableName' cannot be empty or contain only whitespace"
    )
    
    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw $ErrorMessage
    }
    
    return $Value
}

function Assert-PositiveNumber {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VariableName,
        
        [Parameter(Mandatory = $true)]
        [object]$Value,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "Variable '$VariableName' must be a positive number"
    )
    
    if ($null -eq $Value -or -not ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal])) {
        throw "Variable '$VariableName' must be a number. Got: $($Value.GetType().Name)"
    }
    
    if ($Value -le 0) {
        throw $ErrorMessage
    }
    
    return $Value
}

function Assert-NonNegativeNumber {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VariableName,
        
        [Parameter(Mandatory = $true)]
        [object]$Value,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "Variable '$VariableName' must be a non-negative number"
    )
    
    if ($null -eq $Value -or -not ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal])) {
        throw "Variable '$VariableName' must be a number. Got: $($Value.GetType().Name)"
    }
    
    if ($Value -lt 0) {
        throw $ErrorMessage
    }
    
    return $Value
}

function Assert-ArrayNotEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VariableName,
        
        [Parameter(Mandatory = $true)]
        [array]$Value,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "Array '$VariableName' cannot be empty"
    )
    
    if ($null -eq $Value -or $Value.Count -eq 0) {
        throw $ErrorMessage
    }
    
    return $Value
}

function Assert-FileExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "File does not exist: $FilePath"
    )
    
    if (-not (Test-Path $FilePath -PathType Leaf)) {
        throw $ErrorMessage
    }
    
    return $FilePath
}

function Assert-DirectoryExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "Directory does not exist: $DirectoryPath"
    )
    
    if (-not (Test-Path $DirectoryPath -PathType Container)) {
        throw $ErrorMessage
    }
    
    return $DirectoryPath
}

function Assert-ValidPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VariableName,
        
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "Variable '$VariableName' contains an invalid path: $Path"
    )
    
    try {
        $resolvedPath = Resolve-Path $Path -ErrorAction Stop
        return $resolvedPath.Path
    } catch {
        throw "$ErrorMessage. Error: $($_.Exception.Message)"
    }
}
