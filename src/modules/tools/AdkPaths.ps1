# ADK Path utilities for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
. $commonPath

function Get-AdkPaths {
    <#
    .SYNOPSIS
    Gets the standard ADK installation paths for both Windows ADK 10 and 8.1.
    
    .DESCRIPTION
    Returns an array of common ADK installation paths for both x86 and x64 Program Files locations.
    This function centralizes the ADK path definitions to avoid duplication.
    
    .OUTPUTS
    [string[]] Array of ADK installation paths
    #>
    
    $adkPaths = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg",
        "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg",
        "${env:ProgramFiles(x86)}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg",
        "${env:ProgramFiles}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg"
    )
    $adkPaths = Assert-ArrayNotEmpty -VariableName "adkPaths" -Value $adkPaths -ErrorMessage "ADK paths array is empty"
    
    return $adkPaths
}

function Get-AdkExecutablePaths {
    <#
    .SYNOPSIS
    Gets the standard ADK executable paths for oscdimg.exe.
    
    .DESCRIPTION
    Returns an array of common ADK executable paths for oscdimg.exe in both x86 and x64 Program Files locations.
    This function centralizes the ADK executable path definitions to avoid duplication.
    
    .OUTPUTS
    [string[]] Array of ADK executable paths
    #>
    
    $adkPaths = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles(x86)}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
    )
    $adkPaths = Assert-ArrayNotEmpty -VariableName "adkPaths" -Value $adkPaths -ErrorMessage "ADK executable paths array is empty"
    
    return $adkPaths
}
