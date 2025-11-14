<#
.SYNOPSIS
    PSADT v4 wrapper for an interactive Winget upgrade utility.
.DESCRIPTION
    Allows the user to select and upgrade applications available via Winget.
    Designed for deployment as a user-available Win32 app in Intune.

.NOTES
    Run in Intune with Install Command...
    Troubleshooting: ServiceUI.exe -process:explorer.exe %SYSTEMROOT%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Run-WingetUpgrade.ps1"
    Production: ServiceUI.exe -process:explorer.exe %SYSTEMROOT%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "Run-WingetUpgrade.ps1"
#>  

#Requires -RunAsAdministrator

#region Initialization
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Import PSADT module
Import-Module "$ScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1" -Force

# Function to load configuration from config.json
function Get-ConfigurationSettings {
    [CmdletBinding()]
    param(
        [string]$ConfigPath = (Join-Path $ScriptRoot 'config.json')
    )
    
    $defaultConfig = @{
        SessionParameters = @{
            PassThru         = $true
            DeployMode       = 'Interactive'
            AppVendor        = 'Ricoh'
            AppName          = 'Application Upgrader'
            AppScriptAuthor  = 'Tim Welch'
            RequireAdmin     = $true
        }
        DetectionFile = 'C:\ProgramData\CompanyPortalTools\WingetUpgradeTool.installed'
        AcceptableCodes = @(0, -1978335226, -1979189490)
        SkipApplicationIds = @()
        WingetArguments = @{
            upgrade = '--silent --disable-interactivity --nowarn --accept-source-agreements --accept-package-agreements'
        }
        UI = @{
            SelectionWindowTitle = 'Select applications to upgrade'
            SelectionWindowWidth = 500
            SelectionWindowHeight = 600
        }
    }
    
    # Try to load config.json if it exists
    if (Test-Path $ConfigPath) {
        try {
            $jsonConfig = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json
            Write-ADTLogEntry -Message "Loaded configuration from $ConfigPath" -Severity 1
            
            # Merge JSON config with defaults (JSON values override defaults)
            if ($jsonConfig.SessionParameters) {
                foreach ($key in $jsonConfig.SessionParameters.PSObject.Properties.Name) {
                    $defaultConfig.SessionParameters[$key] = $jsonConfig.SessionParameters.$key
                }
            }
            if ($jsonConfig.DetectionFile) { $defaultConfig.DetectionFile = $jsonConfig.DetectionFile }
            if ($jsonConfig.AcceptableCodes) { $defaultConfig.AcceptableCodes = $jsonConfig.AcceptableCodes }
            if ($jsonConfig.SkipApplicationIds) { $defaultConfig.SkipApplicationIds = $jsonConfig.SkipApplicationIds }
            if ($jsonConfig.WingetArguments) { $defaultConfig.WingetArguments = $jsonConfig.WingetArguments }
            if ($jsonConfig.UI) {
                foreach ($key in $jsonConfig.UI.PSObject.Properties.Name) {
                    $defaultConfig.UI[$key] = $jsonConfig.UI.$key
                }
            }
        }
        catch {
            Write-ADTLogEntry -Message "Failed to parse config.json, using defaults: $_" -Severity 2
        }
    }
    else {
        Write-ADTLogEntry -Message "Config file not found at $ConfigPath, using defaults" -Severity 2
    }
    
    return $defaultConfig
}

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

# Load configuration from config.json
$Config = Get-ConfigurationSettings -ConfigPath (Join-Path $ScriptRoot 'config.json')

# Open a toolkit session for logging, prompts, and progress bars
# Prepare SessionParameters with required runtime values
$SessionParameters = $Config.SessionParameters.Clone()
$SessionParameters['SessionState'] = $ExecutionContext.SessionState
$adtSession = Open-ADTSession @SessionParameters -ErrorAction Stop
#endregion Initialization

#region Detection File
$DetectionFile = $Config.DetectionFile
#endregion Detection File

#region Winget Upgrade Logic
Write-ADTLogEntry -Message '=== Starting Winget Upgrade Tool ===' -Severity 1

# Ensure Winget PowerShell module
if (-not (Get-Module -ListAvailable Microsoft.WinGet.Client)) {
    try {
        Write-ADTLogEntry -Message 'Installing Microsoft.WinGet.Client module...' -Severity 1
        Find-PackageProvider -Name NuGet -ForceBootstrap
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
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

    # Filter out skipped application IDs from config
    if ($Config.SkipApplicationIds -and $Config.SkipApplicationIds.Count -gt 0) {
        Write-ADTLogEntry -Message "Filtering out $($Config.SkipApplicationIds.Count) skipped application IDs" -Severity 1
        $upgradeablePackages = $upgradeablePackages | Where-Object { $_.Id -notin $Config.SkipApplicationIds }
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
$window.Title = $Config.UI.SelectionWindowTitle
$window.Width = $Config.UI.SelectionWindowWidth
$window.Height = $Config.UI.SelectionWindowHeight 
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
                    # Use acceptable exit codes from config
                    $acceptableCodes = $Config.AcceptableCodes

                    $Arguments = "upgrade --id `"$($app.Id)`" $($Config.WingetArguments.upgrade)"
                    Write-ADTLogEntry -Message "Executing $winget with arguments: $($Arguments)" -Severity 1

                    # Use the PSAppDeployToolkit wrapper so ADT handles logging and environment correctly.
                    # Request a passthru object so we can inspect ExitCode.
                    # Use -CreateNoWindow to avoid flashing a console window but remove it for troubleshooting if needed
                    $ExecuteResult = Start-ADTProcess -FilePath $winget -ArgumentList $Arguments -CreateNoWindow -PassThru -ErrorAction Stop

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

#region Scheduled Task for Detection File Cleanup
try {
    # Create a scheduled task to delete the detection file after 5 minutes
    Write-ADTLogEntry -Message "Creating scheduled task to delete detection file in 5 minutes" -Severity 1
    
    $taskName = "WingetUpgradeTool_DetectionFileCleanup"
    $taskPath = "\WingetUpgradeTool\"
    $taskDescription = "Automatically deletes the WingetUpgradeTool detection file 5 minutes after creation"
    
    # Remove existing task if it exists to avoid conflicts
    if ((Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue)) {
        Unregister-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Confirm:$false -ErrorAction SilentlyContinue
        Write-ADTLogEntry -Message "Removed existing scheduled task: $taskName" -Severity 1
    }
    
    # Create the cleanup script content
    $cleanupScript = @"
`$DetectionFile = '$DetectionFile'
if (Test-Path `$DetectionFile) {
    Remove-Item -Path `$DetectionFile -Force -ErrorAction SilentlyContinue
    Write-Host "Deleted detection file: `$DetectionFile"
}
"@
    
    # Create action to run PowerShell with the cleanup script
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -Command `"$cleanupScript`""
    
    # Create trigger for 5 minutes from now
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5)
    
    # Create task settings
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
    
    # Register the scheduled task (runs as SYSTEM since this script runs as admin)
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    
    Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description $taskDescription -Force
    
    Write-ADTLogEntry -Message "Successfully created scheduled task to delete detection file in 5 minutes" -Severity 1
}
catch {
    Write-ADTLogEntry -Message "Failed to create scheduled task for detection file cleanup: $_" -Severity 2
    # Don't throw - this is a non-critical operation
}
#endregion Scheduled Task for Detection File Cleanup


# Close the toolkit session and exit cleanly
Close-ADTSession -ExitCode 0
