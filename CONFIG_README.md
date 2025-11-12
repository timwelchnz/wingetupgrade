# Configuration Guide for Run-WingetUpgrade.ps1

The `Run-WingetUpgrade.ps1` script now supports a `config.json` configuration file for easy customization without editing the PowerShell script directly.

## Configuration File Location

- **Default**: `config.json` (in the same directory as `Run-WingetUpgrade.ps1`)
- The script automatically detects and loads `config.json` if present
- If not found, built-in defaults are used

## Configuration Sections

### SessionParameters
Controls the PSAppDeployToolkit session behavior and UI appearance:

```json
"SessionParameters": {
  "PassThru": true,
  "DeployMode": "Interactive",
  "AppVendor": "Ricoh",
  "AppName": "Application Upgrader",
  "AppScriptAuthor": "Tim Welch",
  "RequireAdmin": true
}
```

**Key fields:**
- `DeployMode`: "Interactive" (shows UI) or "Silent" (suppresses UI)
- `AppVendor`: Company/vendor name displayed in PSAppDeployToolkit UI
- `AppName`: Application name displayed in logs and UI
- `RequireAdmin`: Set to `true` to require administrator privileges

### DetectionFile
Specifies the path where the detection marker file is created (used by Intune to detect successful deployment):

```json
"DetectionFile": "C:\\ProgramData\\CompanyPortalTools\\WingetUpgradeTool.installed"
```

Customize this path to match your organization's detection requirements.

### AcceptableCodes
An array of exit codes that winget can return while still being considered a success. This is useful for known false-negative scenarios:

```json
"AcceptableCodes": [0, -1978335226, -1979189490]
```

- `0` = Success
- `-1978335226` (0x8A15010E) = Known false-negative
- `-1979189490` = Another known false-negative

Add or remove codes as needed based on your environment.

### SkipApplicationIds
An array of application IDs to exclude from the upgrade checklist. These apps will be filtered out and not presented to the user:

```json
"SkipApplicationIds": ["Microsoft.Teams", "Microsoft.OneDrive"]
```

Use the exact package ID as returned by `winget list`. Common examples:
- `Microsoft.Teams` - Microsoft Teams
- `Microsoft.OneDrive` - OneDrive
- `Microsoft.Office` - Microsoft Office

Leave empty (`[]`) to show all available upgrades.

### WingetArguments
Global arguments passed to every winget upgrade command:

```json
"WingetArguments": {
  "upgrade": "--silent --disable-interactivity --nowarn --accept-source-agreements --accept-package-agreements"
}
```

Customize these flags to control winget's behavior (e.g., remove `--silent` to show installation progress).

### UI
Settings for the WPF selection window:

```json
"UI": {
  "SelectionWindowTitle": "Select applications to upgrade",
  "SelectionWindowWidth": 500,
  "SelectionWindowHeight": 600
}
```

Adjust width and height to suit your screen resolution and user preferences.

## Example Customizations

### Example 1: Silent Deployment (for Intune automation)

Modify `config.json` to suppress the selection UI:

```json
{
  "SessionParameters": {
    "PassThru": true,
    "DeployMode": "Silent",
    "AppVendor": "YourCompany",
    "AppName": "Winget Auto Upgrade",
    "AppScriptAuthor": "IT Admin",
    "RequireAdmin": true
  },
  "DetectionFile": "C:\\ProgramData\\YourCompany\\WingetAutoUpgrade.marker"
}
```

### Example 2: Custom Company Branding

```json
{
  "SessionParameters": {
    "PassThru": true,
    "DeployMode": "Interactive",
    "AppVendor": "Acme Corp",
    "AppName": "Software Upgrade Tool",
    "AppScriptAuthor": "IT Ops Team",
    "RequireAdmin": true
  },
  "DetectionFile": "C:\\ProgramData\\AcmeCorp\\SoftwareUpgradeTool.installed"
}
```

### Example 3: Less Restrictive Winget Arguments

If you want to see progress during upgrades:

```json
{
  "WingetArguments": {
    "upgrade": "--disable-interactivity --accept-source-agreements --accept-package-agreements"
  }
}
```

### Example 4: Skip Specific Applications

If you want to exclude Teams and OneDrive from the upgrade list:

```json
{
  "SkipApplicationIds": ["Microsoft.Teams", "Microsoft.OneDrive"]
}
```

These applications will be filtered out and won't appear in the user's checklist, even if they have updates available.

## How Configuration is Loaded

1. The script calls `Get-ConfigurationSettings` during initialization
2. Built-in defaults are loaded first
3. If `config.json` exists, it is parsed and its values override the defaults
4. Values not specified in `config.json` retain their defaults
5. If `config.json` has a parse error, the script logs a warning and continues with defaults

## Best Practices

1. **Backup defaults**: Keep the original `config.json` as reference in case you need to revert changes
2. **Version control**: If you're using Git, consider committing `config.json` so your settings are versioned
3. **Environment-specific configs**: You can create multiple configs (e.g., `config-prod.json`, `config-test.json`) and rename as needed before deployment
4. **Test changes**: Always test config changes on a test machine before deploying via Intune
5. **Log validation**: Check PSAppDeployToolkit logs to confirm your settings were applied correctly

## Troubleshooting

**"Config file not found" warning**: This is normal if you delete `config.json`. The script will use hardcoded defaults.

**Config not being read**: Ensure `config.json` is in the same directory as `Run-WingetUpgrade.ps1` and verify JSON syntax (use an online JSON validator if needed).

**Invalid JSON**: If the JSON is malformed, the script will log an error and fall back to defaults. Use a JSON linter to validate syntax.

## PowerShell Execution Policy

When running with `config.json`, ensure you're still using the proper execution policy:

```powershell
& .\Run-WingetUpgrade.ps1
```

Or with explicit policy override (as in Intune):

```powershell
ServiceUI.exe -process:explorer.exe %SYSTEMROOT%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Run-WingetUpgrade.ps1"
```

Both will read and apply `config.json` automatically.
