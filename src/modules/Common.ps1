# Common utilities and helper functions for Windows ISO repack script

function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow", "White")]
        [string]$Color = "White",
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$CurrentIndent = 0,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$InheritedIndent = 0
    )
    
    # Additional validation for edge cases
    if ($Message.Trim().Length -eq 0) {
        throw "Message cannot be empty or contain only whitespace"
    }
    
    # Calculate total indentation (inherited + current)
    $totalIndent = $InheritedIndent + $CurrentIndent
    
    # Validate total indentation doesn't exceed reasonable limits
    if ($totalIndent -gt 40) {
        throw "Total indentation ($totalIndent) exceeds maximum allowed (40). CurrentIndent: $CurrentIndent, InheritedIndent: $InheritedIndent"
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

function Invoke-WebRequestWithCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if ($_ -notmatch '^https?://') {
                throw "URI must be a valid HTTP or HTTPS URL"
            }
            $true
        })]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            $parentDir = Split-Path $_ -Parent
            if (-not (Test-Path $parentDir -PathType Container)) {
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
    
    Write-ColorOutput "Downloading $Description from: $Uri" -Color "White" -CurrentIndent 0 -InheritedIndent 0
    
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
            $headRequest.Method = "HEAD"
            $response = $headRequest.GetResponse()
            $totalBytes = $response.ContentLength
            $response.Close()
        } catch {
            Write-ColorOutput "Could not determine file size, progress tracking may be limited" -Color "Yellow" -CurrentIndent 0 -InheritedIndent 0
        }
        
        # Start download with progress tracking
        $downloadTask = $webClient.DownloadFileTaskAsync($Uri, $OutFile)
        
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
        Write-ColorOutput "$Description downloaded successfully" -Color "Green" -CurrentIndent 0 -InheritedIndent 0
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
        [ValidateScript({
            if ($_ -notmatch '^[a-zA-Z]:\\' -and $_ -notmatch '^\\\\') {
                throw "Path must be a valid absolute path (drive letter or UNC path)"
            }
            $true
        })]
        [string]$Path
    )
    if (-not $KeepWorkingDirectory -and (Test-Path $Path)) {
        Write-ColorOutput "Cleaning up working directory: $Path" -Color "Yellow" -CurrentIndent 0 -InheritedIndent 0
        Remove-Item $Path -Recurse -Force
        Write-ColorOutput "Working directory cleaned up" -Color "Green" -CurrentIndent 0 -InheritedIndent 0
    } elseif ($KeepWorkingDirectory) {
        Write-ColorOutput "Keeping working directory: $Path" -Color "Cyan" -CurrentIndent 0 -InheritedIndent 0
    }
}
