<#
.SYNOPSIS
    PSADT v4 wrapper for an interactive Winget upgrade utility.
.DESCRIPTION
    Allows the user to select and upgrade applications available via Winget.
    Designed for deployment as a user-available Win32 app in Intune.

.NOTES
    Run in Intune with
    ServiceUI.exe -process:explorer.exe %SYSTEMROOT%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Run-WingetUpgrade.ps1"
#>

#Requires -RunAsAdministrator

#region Initialization
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Import PSADT module
Import-Module "$ScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1" -Force

# Function to get the path to winget.exe
function Get-WingetPath {
    $resolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue
    
    if ($resolveWingetPath) {
        return $resolveWingetPath[-1].Path
    }
    
    # Fallback: search manually
    $appxPackage = Get-AppxPackage -Name Microsoft.DesktopAppInstaller -AllUsers | 
        Sort-Object Version -Descending | 
        Select-Object -First 1
    
    if ($appxPackage) {
        return Join-Path $appxPackage.InstallLocation "winget.exe"
    }
    
    return $null
}

# Open a toolkit session for logging, prompts, and progress bars
$SessionParameters = @{
    SessionState = $ExecutionContext.SessionState
    PassThru      = $true
    DeployMode  = "Interactive"
    AppVendor  = "Ricoh"
    AppName    = "Application Upgrader"
    # AppVersion = "1.0.0" - Optionally specify an application version but this tends to overflow popups
    AppScriptAuthor = "Tim Welch"
    RequireAdmin = $true
}
$adtSession = Open-ADTSession @SessionParameters -ErrorAction Stop
#endregion Initialization

#region Detection File
$DetectionFile = 'C:\ProgramData\CompanyPortalTools\WingetUpgradeTool.installed'
#endregion Detection File

#region Winget Upgrade Logic
Write-ADTLogEntry -Message '=== Starting Winget Upgrade Tool ===' -Severity 1

# Ensure Winget PowerShell module
if (-not (Get-Module -ListAvailable Microsoft.WinGet.Client)) {
    try {
        Write-ADTLogEntry -Message 'Installing Microsoft.WinGet.Client module...' -Severity 1
        Install-Module Microsoft.WinGet.Client -Scope AllUsers -Force -ErrorAction Stop
    }
    catch {
        Show-ADTInstallationPrompt -Message "Failed to install Microsoft.WinGet.Client module." -ButtonRightText "OK"
        Throw "Winget module install failed: $_"
    }
}

Import-Module Microsoft.WinGet.Client -Force

# Get apps with available upgrades (including version information)
try {
    Write-ADTLogEntry -Message "Querying installed WinGet packages (in user context)..." -Severity 1

    # Prepare temp paths in detection directory so SYSTEM can read results
    $detectionDir = Split-Path $DetectionFile
    if (-not (Test-Path $detectionDir)) { New-Item -Path $detectionDir -ItemType Directory -Force | Out-Null }
    $tempJson = Join-Path $detectionDir 'winget-packages.json'
    $tempScript = Join-Path $detectionDir 'winget-query.ps1'

    if (Test-Path $tempJson) { Remove-Item -Path $tempJson -Force -ErrorAction SilentlyContinue }
    if (Test-Path $tempScript) { Remove-Item -Path $tempScript -Force -ErrorAction SilentlyContinue }

    # Create a small PowerShell script to run in the interactive user session
    $scriptTemplate = @'
Import-Module Microsoft.WinGet.Client
Get-WinGetPackage | Where-Object { $_.IsUpdateAvailable } | Select-Object Name,Id,InstalledVersion,@{Name="AvailableVersion";Expression={$_.AvailableVersions[0]}} | ConvertTo-Json -Depth 5 | Out-File -FilePath "<<TEMP_JSON>>" -Encoding UTF8
'@

    $scriptContent = $scriptTemplate -replace '<<TEMP_JSON>>', ($tempJson -replace "'", "''")
    # Write script file that will be executed in user context
    $scriptContent | Out-File -FilePath $tempScript -Encoding UTF8 -Force

    # Execute the script in the interactive user session and wait
    # Use Start-ADTProcessAsUser so the WinGet client runs in the user's session
    $psArgs = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$tempScript`""
    Start-ADTProcessAsUser -FilePath "powershell.exe" -ArgumentList $psArgs -ErrorAction Stop

    # Read results produced by user context
    if (-not (Test-Path $tempJson)) {
        throw "User-context query did not produce result file: $tempJson"
    }

    $jsonRaw = Get-Content -Path $tempJson -Raw -ErrorAction Stop
    $upgradeablePackages = if ($jsonRaw) { $jsonRaw | ConvertFrom-Json } else { @() }

    # Normalize single-object vs array from ConvertFrom-Json
    if ($upgradeablePackages -and ($upgradeablePackages -isnot [System.Collections.IEnumerable])) {
        $upgradeablePackages = ,$upgradeablePackages
    }

    if (-not $upgradeablePackages -or $upgradeablePackages.Count -eq 0) {
        Show-ADTInstallationPrompt -Message "No upgrades available via Winget." -ButtonRightText "Close"
        Write-ADTLogEntry -Message "No Winget upgrades found." -Severity 1
        Close-ADTSession -ExitCode 0
        return
    }

    Write-ADTLogEntry -Message "Found $($upgradeablePackages.Count) upgradeable packages." -Severity 1

    # Cleanup temporary script (keep JSON for troubleshooting until later cleanup)
    try { Remove-Item -Path $tempScript -Force -ErrorAction SilentlyContinue } catch {}
}
catch {
    Show-ADTInstallationPrompt -Message "Unable to query Winget packages: $_" -ButtonRightText "OK"
    Throw "Winget query failed: $_"
}

if (-not $upgradeablePackages) {
    Show-ADTInstallationPrompt -Message "No upgrades available via Winget." -ButtonRightText "Close"
    Write-ADTLogEntry -Message "No Winget upgrades found." -Severity 1
    Close-ADTSession -ExitCode 0
    return
}

#region WPF Selection Window with Checkboxes 

Add-Type -AssemblyName PresentationFramework 
Add-Type -AssemblyName PresentationCore 
Add-Type -AssemblyName WindowsBase 

# Import WPF namespaces 
$null = [System.Windows.Window] 
$null = [System.Windows.Controls.Grid] 
$null = [System.Windows.Controls.ColumnDefinition] 

$window = New-Object Windows.Window 
$window.Title = "Select applications to upgrade" 
$window.Width = 500 
$window.Height = 600 
$window.WindowStartupLocation = 'CenterScreen' 

$scrollViewer = New-Object Windows.Controls.ScrollViewer 
$scrollViewer.VerticalScrollBarVisibility = "Auto" 
$scrollViewer.Margin = '15' 

$stackPanel = New-Object Windows.Controls.StackPanel 
$scrollViewer.Content = $stackPanel 

# Create checkboxes with version information for each app 
$checkboxes = @() 
$upgradeablePackages | ForEach-Object { 
    # Create checkbox with app name and version in the same control
    $checkbox = New-Object Windows.Controls.CheckBox
    $checkbox.Content = "$($_.Name): ($($_.InstalledVersion) $([char]0x2192) $($_.AvailableVersion))"
    $checkbox.Tag = $_.Id
    $checkbox.Margin = '5'
    $checkbox.Padding = '5'
    $checkbox.VerticalContentAlignment = 'Center'
    
    # Add checkbox to main vertical stack panel
    $stackPanel.Children.Add($checkbox)
    $checkboxes += $checkbox
} 

# Add Select All checkbox with matching margin and padding
$selectAllCheckbox = New-Object Windows.Controls.CheckBox 
$selectAllCheckbox.Content = "Select All" 
$selectAllCheckbox.Margin = '5'  
$selectAllCheckbox.Padding = '5'  
$selectAllCheckbox.FontWeight = 'Bold'
$selectAllCheckbox.VerticalContentAlignment = 'Center'  # Added for consistency

# Handle Select All functionality 
$selectAllCheckbox.Add_Click({ 
    $isChecked = $selectAllCheckbox.IsChecked 
    $checkboxes | ForEach-Object { 
        $_.IsChecked = $isChecked 
    } 
}) 

$button = New-Object Windows.Controls.Button 
$button.Content = "Upgrade Selected" 
$button.Margin = '10' 
$button.Padding = '10' 
$button.HorizontalAlignment = 'Center' 

$mainStack = New-Object Windows.Controls.StackPanel 
$mainStack.Children.Add($selectAllCheckbox) 
$mainStack.Children.Add($scrollViewer) 
$mainStack.Children.Add($button) 
$window.Content = $mainStack 

$button.Add_Click({ 
    $window.DialogResult = $true 
    $window.Close() 
})

# Test for winget
If (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
    $winget = Get-WingetPath
} else {
    $winget = (Get-Command winget.exe).Source
}

if ($window.ShowDialog()) {
    $selectedApps = $checkboxes | Where-Object { $_.IsChecked } | ForEach-Object {
        @{
            Name = $_.Content.ToString()
            Id = $_.Tag.ToString()
        }
    }

    if ($selectedApps) {
        $totalApps = $selectedApps.Count
        $currentApp = 0

        foreach ($app in $selectedApps) {
            $currentApp++
            Write-ADTLogEntry -Message "Upgrading $($app.Name) (Id: $($app.Id)) - App $currentApp of $totalApps" -Severity 1
            Show-ADTInstallationProgress -StatusMessage "Upgrading $($app.Name) ($currentApp of $totalApps)..."

            if ($app.Id) {
                try {
                    # Treat known false-negative Winget exit codes as success
                    $acceptableCodes = @(0, -1978335226, 0x8A15010E)

                    $Arguments = "upgrade --id `"$($app.Id)`" --silent --disable-interactivity --nowarn --accept-source-agreements --accept-package-agreements"
                    Write-ADTLogEntry -Message "Executing $winget with arguments: $($Arguments)" -Severity 1

                    # Use the PSAppDeployToolkit wrapper so ADT handles logging and environment correctly.
                    # Request a passthru object so we can inspect ExitCode.
                    $ExecuteResult = Start-ADTProcess -FilePath $winget -ArgumentList $Arguments -PassThru -ErrorAction Stop

                    if ($acceptableCodes -contains $ExecuteResult.ExitCode) {
                        Write-ADTLogEntry -Message "$($app.Name) upgraded (ExitCode: $($ExecuteResult.ExitCode))." -Severity 1
                    }
                    else {
                        throw "winget exited with code $($ExecuteResult.ExitCode)"
                    }
                }
                catch {
                    Show-ADTInstallationPrompt -Message "Failed to upgrade $($app.Name)." -ButtonRightText "OK"
                    Write-ADTLogEntry -Message "Upgrade failed for $($app.Name): $_" -Severity 3
                }
            } else {
                Show-ADTInstallationPrompt -Message "Could not locate package ID for $($app.Name)." -ButtonRightText "OK"
                Write-ADTLogEntry -Message "Could not find package ID for $($app.Name)." -Severity 2
            }
        }

        Show-ADTInstallationPrompt -Message "Completed upgrading $totalApps applications." -ButtonRightText "OK"
    } else {
        Show-ADTInstallationPrompt -Message "No applications were selected for upgrade." -ButtonRightText "OK"
    }
}
#endregion WPF Dropdown
#endregion Winget Upgrade Logic


#region Detection File Creation
try {
    if (-not (Test-Path (Split-Path $DetectionFile))) {
        New-Item -ItemType Directory -Force -Path (Split-Path $DetectionFile) | Out-Null
    }
    New-Item -ItemType File -Force -Path $DetectionFile | Out-Null
}
catch {
    Write-Warning "Failed to create detection file: $_"
}
#endregion Detection File Creation


# Close the toolkit session and exit cleanly
Close-ADTSession -ExitCode 0
