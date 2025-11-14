<#
.SYNOPSIS
    Wraps the script package as a Win32 App for Intune deployment using the Microsoft Win32 Content Prep Tool.

.NOTES
    Tim Welch

    Reference: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/
    Prerequisites
    .NET Framework 4.7.2
    
    This script installs the Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe) to C:\tools\IntuneWin32AppUtil
#>

#requires -RunAsAdministrator

$dest = "C:\tools\IntuneWin32AppUtil"

$NetFramework = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue

if (-not $NetFramework -or $NetFramework.Release -lt 461808) {
    Write-Host ".NET Framework 4.7.2 or later not detected — installing..."
    $ProgressPreference = 'SilentlyContinue'    # Subsequent calls do not display UI. - to suppress the output of the file downloading...
    Invoke-WebRequest -uri 'https://go.microsoft.com/fwlink/?linkid=2088631' -OutFile 'ndp48-x86-x64-allos-enu.exe'
    $ProgressPreference = 'Continue'            # Subsequent calls do display UI. - reenable output
    Start-process -filepath .\ndp48-x86-x64-allos-enu.exe -ArgumentList '/q /norestart' -NoNewWindow -Wait -PassThru
}

If (-not(Test-Path "$($dest)\intuneWinAppUtil.exe" -ErrorAction SilentlyContinue)) {
    Write-Host "intuneWinAppUtil.exe not found in PATH. Installing to C:\tools\IntuneWin32AppUtil"
    $release = Invoke-RestMethod "https://api.github.com/repos/microsoft/Microsoft-Win32-Content-Prep-Tool/releases/latest"
    $zipPath = Join-Path $env:TEMP "IntuneWinAppUtil$($release.tag_name).zip"
    Invoke-WebRequest -Uri $release.zipball_url -OutFile $zipPath

  
    If (-not (Test-Path $dest)) {
        New-Item -Path $dest -ItemType Directory | Out-Null
    }

    # Extract to a temporary folder first so we can collapse any single top-level folder
    $tempExtract = Join-Path $env:TEMP "IntuneWinAppUtil_extracted_$($release.tag_name)"
    if (Test-Path $tempExtract) { Remove-Item -Path $tempExtract -Recurse -Force }
    Expand-Archive -Path $zipPath -DestinationPath $tempExtract -Force

    # If the archive creates a single top-level directory, use its contents; otherwise use the root of the extraction
    $children = Get-ChildItem -Path $tempExtract -Force
    if ($children.Count -eq 1 -and $children[0].PSIsContainer) {
        $source = $children[0].FullName
    } else {
        $source = $tempExtract
    }

    # Move all files and folders from source into destination (overwrite existing)
    Get-ChildItem -Path $source -Force | ForEach-Object {
        $target = Join-Path $dest $_.Name
        if (Test-Path $target) { Remove-Item -Path $target -Recurse -Force }
        Move-Item -Path $_.FullName -Destination $dest -Force
    }

    # Cleanup temporary extraction folder
    if (Test-Path $tempExtract) { Remove-Item -Path $tempExtract -Recurse -Force }

} Else {
    Write-Host "intuneWinAppUtil.exe found. Skipping installation."
}

# Build the argument array for IntuneWinAppUtil.exe. Use an array to avoid extra quoting issues.
$AppBuildArgs = @(
    '-c', "$PSScriptRoot\Win32App\",
    '-s', 'ServiceUI.exe',
    '-o', $env:TEMP,
    '-q'
)

Write-Host "Running IntuneWinAppUtil.exe with args: $($AppBuildArgs -join ' ')"

# Prefer the call operator so we get the tool's stdout/stderr in this session and a proper exit code.
$IntuneWinAppUtil = Join-Path $dest 'IntuneWinAppUtil.exe'
if (-not (Test-Path $IntuneWinAppUtil)) { throw "Executable not found: $IntuneWinAppUtil" }

& $IntuneWinAppUtil @AppBuildArgs

if ($LASTEXITCODE -ne 0) {
    throw "IntuneWinAppUtil.exe exited with code $LASTEXITCODE"
} else {
    Try {
        $IntuneWinFile = (gci $env:TEMP -file *.intunewin -ErrorAction Stop | Sort LastWriteTime | Select -Last 1).FullName
    } Catch {
        throw "Intune Win32 App package not found in $env:TEMP after IntuneWinAppUtil.exe completed."
    }
    Write-Host "IntuneWinAppUtil.exe completed successfully."
    Write-Host "The Intune Win32 App package has been created at $($IntuneWinFile)"
}

