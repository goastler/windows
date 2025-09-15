#Requires -RunAsAdministrator

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
    [string]$OemDirectory = (Join-Path (Split-Path $PSScriptRoot -Parent) '$OEM$'),

    [Parameter(Mandatory = $false)]
    [string]$WorkingDirectory = "C:\WinIsoRepack_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [switch]$KeepWorkingDirectory,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeVirtioDrivers,

    [Parameter(Mandatory = $false)]
    [ValidateSet("stable", "latest")]
    [string]$VirtioVersion = "stable",

    [Parameter(Mandatory = $false)]
    [string]$VirtioCacheDirectory = (Join-Path $env:TEMP "virtio-cache")
)

$ErrorActionPreference = "Stop"

function Invoke-WebRequestWithCleanup {
    param(
        [string]$Uri,
        [string]$OutFile,
        [string]$Description = "download file",
        [int]$ProgressId = 3
    )

    Write-Log "Downloading $Description from: $Uri" "White" 0

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
            Write-Log "Could not determine file size, progress tracking may be limited" "Yellow" 0
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
            throw "Download failed: $($downloadTask.Exception.Message)"
        }

        Write-Progress -Activity "Downloading $Description" -Completed -Id $ProgressId
        Write-Host "" # Clear the progress line
        Write-Log "$Description downloaded successfully" "Green" 0
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

function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White",
        [int]$Indent = 0
    )

    $indentString = "  " * $Indent
    $fullMessage = $indentString + $Message
    Write-Host $fullMessage -ForegroundColor $Color
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
