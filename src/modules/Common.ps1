# Common utilities and helper functions for Windows ISO repack script

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White",
        [string]$Colour = $null,
        [int]$Indent = 0
    )
    
    # Use Colour parameter if provided, otherwise use Color parameter
    $actualColor = if ($Colour) { $Colour } else { $Color }
    
    # Create indentation string (2 spaces per indent level)
    $indentString = "  " * $Indent
    
    # Combine indentation with message
    $indentedMessage = $indentString + $Message
    
    Write-Host $indentedMessage -ForegroundColor $actualColor
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-WebRequestWithCleanup {
    param(
        [string]$Uri,
        [string]$OutFile,
        [string]$Description = "file",
        [int]$ProgressId = 3
    )
    
    Write-ColorOutput "Downloading $Description from: $Uri"
    
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
            Write-ColorOutput "Could not determine file size, progress tracking may be limited"
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
        Write-ColorOutput "$Description downloaded successfully"
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
    param([string]$Path)
    if (-not $KeepWorkingDirectory -and (Test-Path $Path)) {
        Write-ColorOutput "Cleaning up working directory: $Path" "Yellow"
        Remove-Item $Path -Recurse -Force
        Write-ColorOutput "Working directory cleaned up" "Green"
    } elseif ($KeepWorkingDirectory) {
        Write-ColorOutput "Keeping working directory: $Path" "Cyan"
    }
}
