# wingetupgrade

Small helper repository to run Windows Package Manager (winget) upgrades via a simple PowerShell wrapper and the PSAppDeployToolkit.

## Overview

This repository provides a script to automate checking for and applying upgrades using winget. The primary entry point is `Run-WingetUpgrade.ps1`, which orchestrates the upgrade process and uses the bundled PSAppDeployToolkit for UI, logging, and elevation handling.

## Prerequisites

- Windows 10 / Windows 11
- PowerShell (built-in)
- Windows Package Manager (`winget`) installed and available in PATH
- Administrator privileges for system-wide upgrades (PSAppDeployToolkit will help request elevation)

## Usage (Intune Win32 app)

This repository is intended to be packaged and deployed as a Win32 App in Microsoft Intune. The preferred deployment pattern is to include all files and subfolders to create a `.intunewin` package using the Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe).

Packaging and deployment summary:

1. Prepare the package content: Include all content in the Win32App folder.
2. Run the Win32 Content Prep Tool to create a `.intunewin` file.
3. In the Microsoft Intune console, add a new Win32 app and upload the `.intunewin` package.
4. Set the Install command (Intune 'Program' > 'Install command') to the following so the PSAppDeployToolkit UI is shown for interactive runs:

```
ServiceUI.exe -process:explorer.exe %SYSTEMROOT%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Run-WingetUpgrade.ps1"
```

Notes and recommendations:

- Run context: Set the app to run as System (recommended for system-wide upgrades) since all users running this will be standard users and unable to elevate.
- Detection rule: Manually configure detection rules - File - Path "C:\ProgramData\CompanyPortalTools\" File "WingetUpgradeTool.installed"
- Requirements: target Windows 10 / Windows 11 devices with `winget` available. If `winget` may not be present, consider adding a pre-install detection or provide instructions to ensure the App Installer/winget is installed.
- Logging: PSAppDeployToolkit will produce logs; configure/introspect the log path during testing to ensure Intune detection and reporting behave as expected.


## Configuration

Configuration values (if present) live under the `Config/` folder. See `Config/config.psd1` for any script options or settings used by the wrapper.

## Files of interest

- `Run-WingetUpgrade.ps1` - Main script to run upgrades (primary logic and invocation).
- `Invoke-AppDeployToolkit.ps1` & `PSAppDeployToolkit/` - Toolkit and helper functions for UI, elevation, and logging.
- `Config/config.psd1` - Configuration (if used) for the script/toolkit.
- `Assets/`, `Files/`, `SupportFiles/` - Place additional assets, supporting files, or setup files here.

## Logs

PSAppDeployToolkit produces logs in its configured log folder (typically a Logs subfolder next to the toolkit execution). Check PSAppDeployToolkit documentation or the script's logging lines for the exact path used in this repo.

## Contributing

Small improvements, bug fixes, or documentation updates are welcome. If you change behavior, update this README with new usage notes.

## License

This repository follows the licensing included with PSAppDeployToolkit. If you add new files or scripts, please include a license or confirm compatibility with the toolkit license.
