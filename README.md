# OneDriveMapper

**Map OneDrive for Business and SharePoint Online document libraries as Windows drive letters - without the sync client.**

OneDriveMapper uses WebDAV (NET USE) to present SharePoint Online and OneDrive for Business document libraries as regular Windows network drives. Users get a familiar drive letter in Explorer (e.g. `X:`) that points directly to their cloud storage, with no sync engine, no OneDrive client installation, and no local disk space consumed.

- **Author:** Jos Lieben ([Lieben Consultancy](https://www.lieben.nu))
- **License:** [Commercial use policy](https://www.lieben.nu/liebensraum/commercial-use/)
- **Documentation & FAQ:** [lieben.nu/liebensraum/onedrivemapper](https://www.lieben.nu/liebensraum/onedrivemapper/)
- **Enterprise alternative:** [OneDriveMapper Cloud](https://www.lieben.nu/liebensraum/onedrivemapper/onedrivemapper-cloud/)

---

## Table of Contents

- [How It Works](#how-it-works)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Configuration Reference](#configuration-reference)
  - [Required Settings](#required-settings)
  - [Drive Mappings](#drive-mappings)
  - [Folder Redirection](#folder-redirection)
  - [Authentication](#authentication)
  - [Behavior Options](#behavior-options)
  - [Progress Bar](#progress-bar)
  - [Logging](#logging)
- [Features](#features)
- [Deployment](#deployment)
- [Version History](#version-history)
  - [v6 (Current)](#v6-current---cdp-based-authentication)
  - [v5](#v5---selenium-based-authentication)
  - [v4](#v4---internet-explorer-based-authentication)
  - [v3](#v3---internet-explorer--native-authentication)
  - [Comparison Table](#comparison-table)
- [Troubleshooting](#troubleshooting)

---

## How It Works

OneDriveMapper authenticates to SharePoint Online / OneDrive for Business, extracts the session cookies (`FedAuth` and `rtFa`), injects them into the Windows Internet cookie store (WinINET), and then uses `NET USE` to create a WebDAV drive mapping. Because Windows' WebClient service reads from the same cookie store, the WebDAV connection is authenticated transparently.

**v6 authentication flow (current):**

```
1. Launch headless Microsoft Edge -> navigates to SharePoint/OneDrive URL
2. Edge uses the device's Primary Refresh Token (PRT) via BrowserCore for silent SSO
3. Entra ID completes OAuth2 -> SharePoint issues FedAuth/rtFa cookies in Edge
4. Cookies are extracted via Chrome DevTools Protocol (CDP) WebSocket
5. Cookies are injected into WinINET (InternetSetCookie)
6. NET USE maps the WebDAV path using the injected cookies
```

If silent SSO fails (device not Entra ID joined, no PRT, conditional access blocking headless), a visible Edge window opens for the user to sign in manually. Cookies are still extracted via CDP - no browser automation framework is needed.

---

## Requirements

| Requirement | Details |
|---|---|
| **Operating System** | Windows 10 / 11 (or Windows Server 2016+) |
| **PowerShell** | 5.1 or later |
| **Browser** | Microsoft Edge (ships with Windows by default) |
| **WebClient Service** | Must be running (started automatically by the script) |
| **Trusted Sites** | SharePoint URLs must be in the Local Intranet or Trusted Sites zone |
| **Silent SSO** | Entra ID joined (or hybrid joined) device with active PRT |
| **Manual fallback** | Any device with Edge - user signs in via browser window |

No app registrations, no client secrets, no Graph API permissions, no Selenium, and no WebDriver downloads are needed.

---

## Quick Start

1. **Download** `OneDriveMapper.ps1` (see [Version History](#version-history) for download links).

2. **Edit the configuration** at the top of the script:
   ```powershell
   $O365CustomerName = 'contoso'  # your tenant name (contoso.onmicrosoft.com -> 'contoso')
   ```

3. **Configure drive mappings** (default maps OneDrive to `X:`):
   ```powershell
   $desiredMappings = @(
       @{
           displayName             = 'OneDrive for Business'
           targetLocationType      = 'driveletter'
           targetLocationPath      = 'X:'
           sourceLocationPath      = 'autodetect'
           mapOnlyForSpecificGroup = ''
       }
   )
   ```

4. **Run the script:**
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\OneDriveMapper.ps1
   ```

5. **Deploy** via login script, scheduled task, Intune, or any other mechanism (see [Deployment](#deployment)).

---

## Configuration Reference

All configuration is done by editing variables at the top of `OneDriveMapper.ps1`. There is no external config file.

### Required Settings

| Variable | Default | Description |
|---|---|---|
| `$O365CustomerName` | `'lieben'` | Your Microsoft 365 tenant name (the part before `.onmicrosoft.com`). |

### Drive Mappings

The `$desiredMappings` array defines what to map. Each entry is a hashtable with:

| Key | Values | Description |
|---|---|---|
| `displayName` | Any string | Label shown in Explorer for the drive or shortcut. |
| `targetLocationType` | `'driveletter'`, `'networklocation'`, `'converged'` | How to expose the mapping. Drive letter, network shortcut, or converged (multiple sites as links under one drive letter). |
| `targetLocationPath` | `'X:'`, folder path | The drive letter to use (for `driveletter`/`converged`) or the folder to create the shortcut in (for `networklocation`). |
| `sourceLocationPath` | `'autodetect'` or full URL | Use `'autodetect'` for OneDrive for Business. For SharePoint, provide the full URL to the document library. |
| `mapOnlyForSpecificGroup` | AD group CN or `''` | Only map if the user is a member of this Active Directory group. Requires domain-joined device with DC connectivity. Leave empty to map for all users. |

**Examples:**

```powershell
$desiredMappings = @(
    # OneDrive for Business -> X: drive
    @{
        displayName             = 'OneDrive for Business'
        targetLocationType      = 'driveletter'
        targetLocationPath      = 'X:'
        sourceLocationPath      = 'autodetect'
        mapOnlyForSpecificGroup = ''
    }
    # SharePoint document library -> Z: drive
    @{
        displayName             = 'Team Documents'
        targetLocationType      = 'driveletter'
        targetLocationPath      = 'Z:'
        sourceLocationPath      = 'https://contoso.sharepoint.com/sites/TeamSite/Shared%20Documents'
        mapOnlyForSpecificGroup = ''
    }
    # SharePoint site as a network shortcut (appears in Explorer sidebar)
    @{
        displayName             = 'Project Files'
        targetLocationType      = 'networklocation'
        targetLocationPath      = "$env:APPDATA\Microsoft\Windows\Network Shortcuts"
        sourceLocationPath      = 'https://contoso.sharepoint.com/sites/Project/Shared%20Documents'
        mapOnlyForSpecificGroup = ''
    }
    # Multiple sites consolidated as links under one drive letter
    @{
        displayName             = 'Marketing Docs'
        targetLocationType      = 'converged'
        targetLocationPath      = 'Y:'
        sourceLocationPath      = 'https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents'
        mapOnlyForSpecificGroup = ''
    }
    @{
        displayName             = 'Sales Docs'
        targetLocationType      = 'converged'
        targetLocationPath      = 'Y:'
        sourceLocationPath      = 'https://contoso.sharepoint.com/sites/Sales/Shared%20Documents'
        mapOnlyForSpecificGroup = 'SalesTeam'
    }
)
```

### Folder Redirection

Redirect Windows known folders (Desktop, Documents, Pictures) into the mapped OneDrive drive:

```powershell
$redirectFolders = $true
$listOfFoldersToRedirect = @(
    @{ knownFolderInternalName = 'Desktop';     knownFolderInternalIdentifier = 'Desktop';   desiredTargetPath = 'X:\Desktop';      copyExistingFiles = 'true' }
    @{ knownFolderInternalName = 'MyDocuments';  knownFolderInternalIdentifier = 'Documents'; desiredTargetPath = 'X:\My Documents'; copyExistingFiles = 'true' }
    @{ knownFolderInternalName = 'MyPictures';   knownFolderInternalIdentifier = 'Pictures';  desiredTargetPath = 'X:\My Pictures';  copyExistingFiles = 'false' }
)
```

Set `$restartExplorer = $true` when using folder redirection to ensure Explorer picks up the changes immediately.

### Authentication

| Variable | Default | Description |
|---|---|---|
| `$edgeWaitSeconds` | `10` | Maximum seconds to wait for headless Edge to complete silent SSO via PRT. The script polls and finishes as soon as authentication cookies appear, so this is a ceiling, not a fixed delay. Increase on slow networks. |
| `$fallbackToVisibleAuth` | `$true` | When silent SSO fails, open a visible Edge window for the user to sign in manually. |
| `$visibleAuthTimeout` | `300` | Maximum seconds to wait for manual sign-in before giving up (5 minutes). |
| `$reauthIntervalHours` | `12` | Proactively re-authenticate every N hours so SharePoint cookies are refreshed before they expire (prevents stale drives after sleep/resume or overnight). Set `0` to disable. Requires `$autoRemapMethod` to be enabled. |

### Behavior Options

| Variable | Default | Description |
|---|---|---|
| `$showConsoleOutput` | `$true` | Display log messages in the console window. |
| `$autoRemapMethod` | `'Path'` | Monitor and remap disconnected drives. `'Path'` checks the underlying WebDAV path, `'Link'` checks the drive letter exists, `'Disabled'` turns off monitoring. |
| `$restartExplorer` | `$false` | Restart Explorer after mapping to refresh drive visibility. Primarily needed with folder redirection. |
| `$libraryName` | `'Documents'` | OneDrive document library name. Almost always `'Documents'`. |
| `$displayErrors` | `$true` | Show a dialog box to the user when errors occur. |
| `$persistentMapping` | `$true` | Use `/PERSISTENT:YES` with NET USE so mappings survive logoff. |
| `$urlOpenAfter` | `''` | URL to open in Edge after mapping completes. |
| `$removeExistingMaps` | `$true` | Remove existing SharePoint/OneDrive drive mappings before remapping. |
| `$removeEmptyMaps` | `$true` | Remove dead/empty drive mappings. |
| `$autoDetectProxy` | `$false` | Disable IE/Windows auto-proxy detection (speeds up WebDAV significantly). |
| `$addShellLink` | `$false` | Create a Favorites shortcut for OneDrive (Windows 7/8 style). |
| `$createUserFolderOn` | `''` | Drive letter(s) on which to automatically create a per-user subfolder. |
| `$convergedDriveLabel` | `'SharePoint and Team sites'` | Label for the converged drive letter. |

### Progress Bar

| Variable | Default | Description |
|---|---|---|
| `$showProgressBar` | `$true` | Show a visual progress bar during mapping. |
| `$progressBarColor` | `'#CC99FF'` | Progress bar accent color (HTML hex). |
| `$progressBarText` | `'OneDriveMapper v<version> is connecting your drives...'` | Text displayed in the progress bar window. |

### Logging

| Variable | Default | Description |
|---|---|---|
| `$logfile` | `%APPDATA%\OneDriveMapper.log` | Log file path. |
| `$maxLocalLogSizeMB` | `2` | Maximum log file size in MB before rotation. |

### Central Configuration (Intune / GPO)

All scalar settings above (and the mapping tables as JSON) can be overridden without editing the script by deploying registry values:

| Registry key | Scope |
|---|---|
| `HKCU\Software\Policies\Lieben\OneDriveMapper` | User policy |
| `HKLM\Software\Policies\Lieben\OneDriveMapper` | Machine policy (takes precedence) |

Value name = variable name (e.g. `O365CustomerName`). Booleans as DWORD `0`/`1`, numbers as DWORD, text as REG_SZ. `desiredMappings` and `listOfFoldersToRedirect` can be provided as a REG_SZ containing a JSON array, e.g.:

```json
[{"displayName":"Onedrive for Business","targetLocationType":"driveletter","targetLocationPath":"X:","sourceLocationPath":"autodetect","mapOnlyForSpecificGroup":""}]
```

Applied overrides are written to the log at startup, so you can always see which settings came from policy. This lets you deploy one unmodified script to your entire estate and manage tenant name, mappings, and behavior centrally.

---

## Features

- **Zero dependencies** - No Selenium, no WebDriver, no app registration, no client secrets. Only Edge (preinstalled on Windows) and PowerShell 5.1.
- **Silent SSO** - Fully automatic authentication on Entra ID joined devices using the Primary Refresh Token.
- **Manual fallback** - If silent SSO fails, a visible Edge window opens for the user to sign in. Works on any device.
- **OneDrive autodetect** - Automatically resolves the user's personal OneDrive path from their UPN.
- **SharePoint document libraries** - Map any number of SharePoint Online document libraries to drive letters.
- **Converged drives** - Combine multiple SharePoint sites as links under a single drive letter.
- **Network locations** - Create Explorer network shortcuts instead of drive letters.
- **Folder redirection** - Redirect Desktop, Documents, and Pictures into the OneDrive drive.
- **Auto-remap** - Monitors mapped drives and automatically remaps them if they disconnect.
- **AD group filtering** - Conditionally map drives based on Active Directory group membership.
- **Elevation bypass** - If accidentally run as Administrator, re-launches as the standard user via scheduled task.
- **Progress bar** - Visual progress indicator for end users.
- **Per-user folders** - Optionally create individual user folders on shared drives.
- **Central configuration** - Override any setting via registry policies (Intune/GPO) without editing the script.
- **Proactive cookie refresh** - Re-authenticates silently before cookies expire so drives never go stale.
- **Diagnostics mode** - `-Diagnose` produces a full environment health report without making changes.

---

## Deployment

OneDriveMapper is a single `.ps1` file with no external dependencies. Common deployment methods:

### Intune (Recommended for Entra ID joined devices)

Deploy as a PowerShell script via Intune:
- **Script:** `OneDriveMapper.ps1`
- **Run as:** User context (not System)
- **Execution policy:** Bypass
- **Arguments:** `-HideConsole`

### Group Policy Login Script

```
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "\\server\share\OneDriveMapper.ps1" -HideConsole
```

### Scheduled Task

The script includes built-in elevation bypass - if it detects it's running elevated, it creates a scheduled task to re-run as the standard user. You can also pre-create a scheduled task:

```powershell
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Scripts\OneDriveMapper.ps1" -HideConsole'
$trigger = New-ScheduledTaskTrigger -AtLogOn
Register-ScheduledTask -TaskName 'OneDriveMapper' -Action $action -Trigger $trigger -RunLevel Limited
```

### Parameters

| Parameter | Description |
|---|---|
| `-HideConsole` | Hides the PowerShell console window. Use in production deployments. |
| `-AsTask` | Internal flag indicating the script was launched from the elevation bypass scheduled task. |
| `-Diagnose` | Reports environment health (Edge, PRT, WebClient service, tenant connectivity, existing mappings, effective configuration) and exits without making any changes. Ideal for helpdesk triage: `powershell -ExecutionPolicy Bypass -File OneDriveMapper.ps1 -Diagnose` |

---

## Version History

OneDriveMapper has gone through several major architectural changes over the years as Microsoft's authentication landscape evolved. Each major version represents a fundamentally different authentication approach.

### v6 (Current) - CDP-based Authentication

**Download:** [`OneDriveMapper.ps1`](OneDriveMapper.ps1) (main branch)

The current version. Uses headless Microsoft Edge and Chrome DevTools Protocol (CDP) to extract cookies after PRT-based silent SSO. Zero external dependencies.

**Key characteristics:**
- Authentication via headless Edge + CDP WebSocket
- Silent SSO using the device's Primary Refresh Token (PRT) via Edge's native BrowserCore integration
- Falls back to visible Edge window for manual login
- No Selenium, no WebDriver, no IE, no browser automation framework
- No app registrations or API permissions needed
- ~1,580 lines, cleanly structured with regions and proper function documentation
- PowerShell 5.1+ required, Strict Mode enabled

**Added in v6.10:**
- **Central configuration via registry** - manage all settings (including mappings as JSON) through Intune/GPO under `Software\Policies\Lieben\OneDriveMapper`, no script editing needed
- **`-Diagnose` parameter** - one-shot environment health report (Edge, PRT, WebClient, tenant connectivity, existing mappings) for helpdesk triage, makes no changes
- **Faster sign-in** - authentication now polls for cookies and completes as soon as SSO finishes (typically 3-6s) instead of always waiting the full `$edgeWaitSeconds`; also eliminates the race where cookies were read before they were set
- **Proactive cookie refresh** (`$reauthIntervalHours`) - silently re-authenticates before SharePoint cookies expire, so drives no longer go stale overnight or after sleep/resume
- **Tenant-based connectivity checks** - drive monitoring now tests TCP 443 to your own SharePoint host instead of pinging 8.8.8.8 (which many corporate firewalls block)
- **Converged drive fix** - converged mappings now correctly create their shortcuts inside the converged drive and are health-checked correctly (previously they could trigger an endless remap loop)
- **Elevation bypass fix** - the scheduled-task fallback now survives script paths containing spaces
- Cookies are injected with the `secure` attribute (https-only), auth cookie values never touch the log file
- Stable log file name (`OneDriveMapper.log`) across versions, version still logged in the banner

---

### v5 - Selenium-based Authentication

**Download:** [`Releases/v5.16.ps1`](Releases/v5.16.ps1)

Replaced Internet Explorer with Microsoft Edge controlled via Selenium WebDriver. This fixed the authentication issues caused by Microsoft removing IE support, but introduced a dependency on Selenium and the Edge WebDriver binary.

**Key characteristics:**
- Authentication via Selenium-controlled Edge browser
- Required `WebDriver.dll` (.NET Selenium library) and `msedgedriver.exe`
- Auto-downloaded and auto-updated the Edge driver to match the installed Edge version
- Extracted cookies from the Selenium-controlled browser session
- `$forceHideEdge` option to suppress the browser window entirely
- ~1,650 lines

**Drawbacks addressed by v6:**
- Selenium WebDriver broke frequently with Edge updates (version mismatch)
- `msedgedriver.exe` download could be blocked by firewalls/proxies
- Running a visible browser window was sometimes unavoidable
- Heavy dependency chain for what is essentially "get two cookies"

---

### v4 - Internet Explorer-based Authentication

**Download:** [`Releases/v4.08.ps1`](Releases/v4.08.ps1)

Used Internet Explorer's COM automation (`InternetExplorer.Application`) for authentication. This was the most widely deployed version but became obsolete as Microsoft deprecated and removed IE.

**Key characteristics:**
- Authentication via Internet Explorer COM object
- Required IE Protected Mode to be disabled (temporarily managed by the script)
- Supported Azure AD Connect SSO (`$useAzAdConnectSSO`)
- Auto-mapped favorited SharePoint sites (`$autoMapFavoriteSites`)
- Required `Keep Me Signed In` (KMSI) to be enabled tenant-wide
- ~1,710 lines

**Why it was replaced:**
- Internet Explorer removed from Windows 11 and later Windows 10 builds
- IE COM automation increasingly unreliable with modern auth flows
- Protected Mode toggling caused issues in locked-down environments

---

### v3 - Native Authentication

**Download:** [`Releases/v3.30.ps1`](Releases/v3.30.ps1)

The original architecture. Uses direct http calls with extensive support for different user identity lookup methods (AD UPN, AD email, Azure AD join, interactive prompt, registry key, full login form, `whoami /upn`) and ADFS authentication modes including certificate-based auth.

**Key characteristics:**
- Seven different `$userLookupMode` options for determining user identity
- Three ADFS authentication modes including client certificate matching
- Cached credentials (encrypted password stored in AppData)
- Cookie caching for silent re-authentication
- Supported auto-mapping of favorited sites
- ~3,150 lines (largest version due to all the authentication permutations)

**Why it was replaced:**
- Extremely complex credential management logic
- Many authentication modes made troubleshooting difficult and lack native MFA support
- Password caching raised security concerns

---

### Comparison Table

| Feature | v3 | v4 | v5 | v6 |
|---|:---:|:---:|:---:|:---:|
| **Authentication** | credentials/ADFS | IE COM | Selenium + Edge | Headless Edge + CDP |
| **Browser dependency** | None | Internet Explorer | Edge + WebDriver | Edge (built-in) |
| **External files needed** | None | None | WebDriver.dll + msedgedriver.exe | None |
| **Silent SSO** | Cookie cache | KMSI + IE | Selenium auto-login | PRT via BrowserCore |
| **Manual login fallback** | Built-in form / IE | IE prompt | Edge window | Edge window |
| **App registration** | No | No | No | No |
| **Works without IE** | No | No | Yes | Yes |
| **Works on Windows 11** | No | No | Yes | Yes |
| **Entra ID joined support** | Limited | Limited | Yes | Native (PRT) |
| **Auto-remap** | No | Yes | Yes | Yes |
| **Folder redirection** | Yes | Yes | Yes | Yes |
| **Converged drives** | Yes (v3.28+) | Yes | Yes | Yes |
| **Favorited sites auto-map** | Yes | Yes | No | No |
| **Credential caching** | Encrypted file | IE cookies | Selenium cookies | None needed (PRT) |
| **Approximate size** | ~3,150 lines | ~1,710 lines | ~1,650 lines | ~1,580 lines |
| **PowerShell version** | 3.0+ | 3.0+ | 3.0+ | 5.1+ |
| **Status** | Legacy | Deprecated | Deprecated | **Current** |

---

## Troubleshooting

### Common Issues

**"WebClient service is not running"**
The WebClient (WebDAV) service must be running. OneDriveMapper attempts to start it automatically. If it fails, start it manually or ensure it's set to `Manual` or `Automatic` startup:
```powershell
Set-Service WebClient -StartupType Manual
Start-Service WebClient
```

**"Silent SSO did not complete"**
- Verify the device is Entra ID joined: `dsregcmd /status` -> `AzureAdJoined: YES`
- Verify PRT is present: `dsregcmd /status` -> `AzureAdPrt: YES`
- Increase `$edgeWaitSeconds` on slow networks
- Check that Conditional Access policies don't block headless browsers
- The script will fall back to visible Edge for manual login

**"WebDAV file locking is enabled"**
This is a warning, not an error. File locking can cause issues with some applications. To disable:
```
HKLM\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\SupportLocking = 0
```

**"NET USE error 67"**
The WebDAV URL is not trusted. Ensure the SharePoint URLs are added to the Trusted Sites or Local Intranet zone. OneDriveMapper adds them automatically, but GPO may override this.

**Drive mapping succeeds but drive is empty or inaccessible**
- Check `$libraryName` matches your OneDrive library name (usually `'Documents'`)
- Verify the WebDAV max file size: `HKLM\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\FileSizeLimitInBytes`
- Default is 50MB - increase if needed for large files

### Log File

Check the log file at `%APPDATA%\OneDriveMapper.log` for detailed diagnostic information. For a quick environment health check, run the script with `-Diagnose`. Each line is timestamped and categorized as INFO, WARNING, or ERROR.

---

## Links

- **Documentation & FAQ:** https://www.lieben.nu/liebensraum/onedrivemapper/
- **Enterprise version:** https://www.lieben.nu/liebensraum/onedrivemapper/onedrivemapper-cloud/
- **License:** https://www.lieben.nu/liebensraum/commercial-use/
- **Author:** Jos Lieben - [Lieben Consultancy](https://www.lieben.nu)