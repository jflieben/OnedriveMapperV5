#Requires -Version 5.1
<#
.SYNOPSIS
    OneDriveMapper v6.00 - Maps OneDrive for Business and SharePoint libraries as network drives.

.DESCRIPTION
    Maps OneDrive for Business and/or SharePoint Online document libraries as Windows drive letters
    or network locations using WebDAV. Authentication is fully automatic via the device's Primary
    Refresh Token (PRT) using headless Microsoft Edge and Chrome DevTools Protocol (CDP).

    NO Selenium. NO WebDriver. NO app registration. NO user interaction on Entra ID joined devices.

    Authentication flow:
    1. Launch headless Edge navigating to the SharePoint/OneDrive URL
    2. Edge uses PRT via its native BrowserCore integration for silent SSO
    3. Entra ID completes OAuth2 → SharePoint issues FedAuth/rtFa cookies
    4. Cookies extracted via CDP WebSocket and injected into WinINET
    5. NET USE maps the WebDAV path using the injected cookies

    If silent SSO fails (non-Entra joined device, no PRT), falls back to a visible Edge window
    where the user can sign in manually. Cookies are still extracted via CDP - no Selenium needed.

    Requirements:
    - Windows 10/11 with Microsoft Edge (ships by default)
    - PowerShell 5.1+
    - For silent SSO: Entra ID joined device with active PRT
    - For manual fallback: user signs in via Edge window

.PARAMETER AsTask
    Indicates the script is running from a scheduled task (used for elevation bypass).

.PARAMETER HideConsole
    Hides the PowerShell console window on startup.

.NOTES
    Copyright/License: https://www.lieben.nu/liebensraum/commercial-use/
    Author:            Jos Lieben (Lieben Consultancy)
    Script help:       https://www.lieben.nu/liebensraum/onedrivemapper/
    Enterprise users:  https://www.lieben.nu/liebensraum/onedrivemapper/onedrivemapper-cloud/
#>

[CmdletBinding()]
param(
    [switch]$AsTask,
    [switch]$HideConsole
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

#region ===== CONFIGURATION =====

$version = '6.00'

# REQUIRED - set your tenant name (e.g. 'contoso' from contoso.onmicrosoft.com)
$O365CustomerName = 'lieben'

<#
DRIVE MAPPING CONFIGURATION
  displayName             = Label for the drive (shown in Explorer)
  targetLocationType      = 'driveletter', 'converged', or 'networklocation'
  targetLocationPath      = Drive letter (e.g. 'X:') or folder path for network shortcuts
  sourceLocationPath      = 'autodetect' for OneDrive, or full SharePoint document library URL
  mapOnlyForSpecificGroup = AD group CN (domain-joined only); leave empty to map for all users
#>
$desiredMappings = @(
    @{
        displayName             = 'Onedrive for Business'
        targetLocationType      = 'driveletter'
        targetLocationPath      = 'X:'
        sourceLocationPath      = 'autodetect'
        mapOnlyForSpecificGroup = ''
    }
    #@{
    #    displayName             = 'Sharepoint Site'
    #    targetLocationType      = 'driveletter'
    #    targetLocationPath      = 'Z:'
    #    sourceLocationPath      = 'https://lieben.sharepoint.com/sites/testing/Gedeelde%20documenten'
    #    mapOnlyForSpecificGroup = ''
    #}
)

# Folder redirection (moves Desktop/Documents/Pictures into the OneDrive drive)
$redirectFolders = $false
$listOfFoldersToRedirect = @(
    @{ knownFolderInternalName = 'Desktop';    knownFolderInternalIdentifier = 'Desktop';   desiredTargetPath = 'X:\Desktop';      copyExistingFiles = 'true' }
    @{ knownFolderInternalName = 'MyDocuments'; knownFolderInternalIdentifier = 'Documents'; desiredTargetPath = 'X:\My Documents'; copyExistingFiles = 'true' }
    @{ knownFolderInternalName = 'MyPictures';  knownFolderInternalIdentifier = 'Pictures';  desiredTargetPath = 'X:\My Pictures';  copyExistingFiles = 'false' }
)

# OPTIONAL - Behavior
$showConsoleOutput     = $true                # Show log output in console window
$autoRemapMethod       = 'Path'               # 'Path', 'Link', or 'Disabled' - monitor and remap disconnected drives
$restartExplorer       = $false               # Restart Explorer after mapping (refreshes drive visibility)
$libraryName           = 'Documents'          # OneDrive library name (usually 'Documents')
$displayErrors         = $true                # Show error dialog to user on failure
$persistentMapping     = $true                # Use /PERSISTENT:YES for NET USE
$urlOpenAfter          = ''                   # URL to open in Edge after mapping completes
$removeExistingMaps    = $true                # Remove existing SP/OneDrive mappings before remapping
$removeEmptyMaps       = $true                # Remove dead/empty drive mappings
$autoDetectProxy       = $false               # Disable IE auto-proxy detection (speeds up WebDAV)
$addShellLink          = $false               # Create a Favorites shortcut for OneDrive
$createUserFolderOn    = ''                   # Drive letter on which to create per-user folder (e.g. 'Q:')
$convergedDriveLabel   = 'SharePoint and Team sites'

# OPTIONAL - Progress bar
$showProgressBar       = $true
$progressBarColor      = '#CC99FF'
$progressBarText       = "OneDriveMapper v$version is connecting your drives..."
$showSystemTrayIcon    = $true                # Show status icon in system tray

# OPTIONAL - Authentication
$edgeWaitSeconds       = 10                   # Seconds to wait for silent PRT SSO
$fallbackToVisibleAuth = $true                # Show visible Edge if silent SSO fails
$visibleAuthTimeout    = 300                  # Max seconds to wait for manual login (5 min)

# OPTIONAL - Logging
$logfile               = Join-Path $env:APPDATA "OneDriveMapper_$version.log"
$maxLocalLogSizeMB     = 2

# OPTIONAL - Icon paths (set to valid .ico paths if available)
$onedriveIconPath      = ''
$sharepointIconPath    = ''

#endregion

#region ===== CONSTANTS =====

$privateSuffix         = '-my'
$script:errorsForUser  = ''
$script:traySync       = $null
$script:trayRunspace   = $null
$script:trayPS         = $null
$O365CustomerName      = $O365CustomerName.ToLower() -replace '\.onmicrosoft\.com', ''

#endregion

#region ===== NATIVE CODE =====

# Hide console window API
$winApiType = Add-Type -Name 'Win32Window' -Namespace 'ODM' -PassThru -MemberDefinition @'
[DllImport("user32.dll")]
public static extern bool ShowWindow(int handle, int state);
'@

if ($HideConsole) {
    try {
        $proc = [System.Diagnostics.Process]::GetCurrentProcess()
        $null = $winApiType::ShowWindow($proc.MainWindowHandle, 0)
    }
    catch { <# Non-critical #> }
}

# WinINET cookie interop
$cookieInteropSource = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class WinInetCookies {
    [DllImport("wininet.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool InternetSetCookie(string url, string cookieName, string cookieData);

    [DllImport("wininet.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool InternetGetCookieEx(
        string url, string cookieName, StringBuilder cookieData,
        ref int size, int flags, IntPtr reserved);

    public static bool SetCookie(string url, string name, string value) {
        bool result = InternetSetCookie(url, name, value);
        if (!result) {
            int err = Marshal.GetLastWin32Error();
            throw new Exception("InternetSetCookie failed. Win32 error: " + err);
        }
        return result;
    }

    public static string GetCookies(string url) {
        int size = 8192;
        var sb = new StringBuilder(size);
        if (InternetGetCookieEx(url, null, sb, ref size, 0x2000, IntPtr.Zero)) {
            return sb.ToString();
        }
        return null;
    }
}
"@

try { [WinInetCookies] | Out-Null } catch {
    Add-Type -TypeDefinition $cookieInteropSource -Language CSharp -ErrorAction Stop
}

# Shell32 notification for Explorer refresh
try {
    Add-Type -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("Shell32.dll")]
private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
public static void Refresh() {
    SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);
}
'@ -Namespace WinAPI -Name Explorer -ErrorAction Stop
}
catch { <# Already loaded or non-critical #> }

# Foreground window API
try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class Win32SetWindow {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
'@ -ErrorAction Stop
}
catch { <# Already loaded #> }

# Load Windows Forms for progress bar
try {
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Web')
}
catch {
    $showProgressBar = $false
}

# Enforce modern TLS
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
}
catch {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
}

#endregion

#region ===== LOGGING =====

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Text,
        [switch]$IsError,
        [switch]$IsWarning
    )

    $prefix = switch ($true) {
        $IsError   { 'ERROR'   }
        $IsWarning { 'WARNING' }
        default    { 'INFO'    }
    }

    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $prefix | $Text"

    try { Add-Content -Path $logfile -Value $entry -ErrorAction SilentlyContinue } catch { }

    if ($showConsoleOutput) {
        $color = switch ($true) {
            $IsError   { 'Red'    }
            $IsWarning { 'Yellow' }
            default    { 'Green'  }
        }
        Write-Host $entry -ForegroundColor $color
    }
}

function Reset-LogFile {
    [CmdletBinding()]
    param()
    if (-not (Test-Path $logfile)) { return }
    try {
        $size = (Get-Item $logfile -ErrorAction Stop).Length
        if (($size / 1MB) -gt $maxLocalLogSizeMB) {
            $old = "$logfile.old"
            if (Test-Path $old) { Remove-Item $old -Force -Confirm:$false }
            Rename-Item -Path $logfile -NewName $old -Force -Confirm:$false
            Write-Log -Text 'Log file rotated (size limit reached)'
        }
    }
    catch { }
}

#endregion

#region ===== SYSTEM UTILITIES =====

function Get-RegistryValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$EntryName
    )
    try {
        return (Get-ItemProperty -Path "$BasePath\" -Name $EntryName -ErrorAction Stop).$EntryName
    }
    catch { return -1 }
}

function Add-SiteToIEZone {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [int]$Mode = 2  # 1 = Intranet, 2 = Trusted Sites
    )
    try {
        $parts = $SiteUrl.Split('.')
        if ($parts.Count -gt 3) {
            $sub = ($parts[0..($parts.Count - 3)]) -join '.'
            $parts = @($sub, $parts[$parts.Count - 2], $parts[$parts.Count - 1])
        }
        $base = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains'
        if ($parts.Count -gt 2) {
            $domain = "$($parts[1]).$($parts[2])"
            $null = New-Item "$base\$domain" -ErrorAction SilentlyContinue
            $null = New-Item "$base\$domain\$($parts[0])" -ErrorAction SilentlyContinue
            $null = New-ItemProperty "$base\$domain\$($parts[0])" -Name 'https' -Value $Mode -ErrorAction Stop
        }
        else {
            $domain = "$($parts[0]).$($parts[1])"
            $null = New-Item "$base\$domain" -ErrorAction SilentlyContinue
            $null = New-ItemProperty "$base\$domain" -Name 'https' -Value $Mode -ErrorAction Stop
        }
        return $true
    }
    catch { return -1 }
}

function Get-ExplorerProcessForCurrentUser {
    [CmdletBinding()]
    param()
    try {
        $procs = Get-CimInstance -ClassName Win32_Process -Filter "name LIKE 'explorer%'" -ErrorAction Stop
    }
    catch { return -1 }
    if ($null -eq $procs) { return 0 }

    $owned = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $procs) {
        try {
            $owner = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction Stop
            if ($owner.User -eq $env:USERNAME) {
                $p | Add-Member -MemberType NoteProperty -Name 'UserName' -Value $owner.User -Force
                $owned.Add($p)
            }
        }
        catch { }
    }
    return $(if ($owned.Count -gt 0) { $owned.ToArray() } else { 0 })
}

function Restart-ExplorerProcess {
    [CmdletBinding()]
    param()
    Write-Log -Text 'Restarting Explorer.exe to refresh drive visibility'
    $procs = @(Get-ExplorerProcessForCurrentUser)
    foreach ($p in $procs) {
        if ($p -is [CimInstance]) {
            try { Stop-Process -Id $p.handle -Force -ErrorAction Stop } catch { }
        }
    }
}

function Start-WebDavClient {
    [CmdletBinding()]
    param()

    $Source = @"
using System;
using System.Text;
using System.Security;
using System.Collections.Generic;
using System.Runtime.Versioning;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
namespace JosL.WebClient{
    public static class Starter{
        [StructLayout(LayoutKind.Explicit, Size=16)]
        public class EVENT_DESCRIPTOR{
            [FieldOffset(0)]ushort Id = 1;
            [FieldOffset(2)]byte Version = 0;
            [FieldOffset(3)]byte Channel = 0;
            [FieldOffset(4)]byte Level = 4;
            [FieldOffset(5)]byte Opcode = 0;
            [FieldOffset(6)]ushort Task = 0;
            [FieldOffset(8)]long Keyword = 0;
        }
        [StructLayout(LayoutKind.Explicit, Size = 16)]
        public struct EventData{
            [FieldOffset(0)]internal UInt64 DataPointer;
            [FieldOffset(8)]internal uint Size;
            [FieldOffset(12)]internal int Reserved;
        }
        public static void startService(){
            Guid webClientTrigger = new Guid(0x22B6D684, 0xFA63, 0x4578, 0x87, 0xC9, 0xEF, 0xFC, 0xBE, 0x66, 0x43, 0xC7);
            long handle = 0;
            uint output = EventRegister(ref webClientTrigger, IntPtr.Zero, IntPtr.Zero, ref handle);
            if (output == 0){
                EVENT_DESCRIPTOR desc = new EVENT_DESCRIPTOR();
                unsafe{
                    EventWrite(handle, ref desc, 0, null);
                    EventUnregister(handle);
                }
            }
        }
        [DllImport("Advapi32.dll", SetLastError = true)]
        public static extern uint EventRegister(ref Guid guid, [Optional] IntPtr EnableCallback, [Optional] IntPtr CallbackContext, [In][Out] ref long RegHandle);
        [DllImport("Advapi32.dll", SetLastError = true)]
        public static extern unsafe uint EventWrite(long RegHandle, ref EVENT_DESCRIPTOR EventDescriptor, uint UserDataCount, EventData* UserData);
        [DllImport("Advapi32.dll", SetLastError = true)]
        public static extern uint EventUnregister(long RegHandle);
    }
}
"@
    try {
        Write-Log -Text 'Starting WebClient service via ETW trigger...'
        $cp = New-Object System.CodeDom.Compiler.CompilerParameters
        $cp.CompilerOptions = '/unsafe'
        $cp.GenerateInMemory = $true
        if ($PSVersionTable.PSVersion.Major -eq 5) {
            Add-Type -TypeDefinition $Source -Language CSharp -CompilerParameters $cp
        }
        else {
            Add-Type -TypeDefinition $Source -Language CSharp -CompilerOptions $cp
        }
        [JosL.WebClient.Starter]::startService()
        Start-Sleep -Seconds 5
        if ((Get-Service -Name WebClient).Status -eq 'Running') {
            Write-Log -Text 'WebClient service started successfully'
        }
        else {
            Write-Log -Text 'WebClient service failed to start. Set it to Automatic start.' -IsError
        }
    }
    catch {
        Write-Log -Text "Failed to start WebClient: $_" -IsError
    }
}

function Test-WebClient {
    [CmdletBinding()]
    param()
    if ((Get-Service -Name WebClient).Status -ne 'Running') {
        Write-Log -Text 'WebClient service is not running' -IsWarning
        if ($script:isElevated) {
            Start-Service WebClient -ErrorAction SilentlyContinue
        }
        else {
            Start-WebDavClient
        }
        if ((Get-Service -Name WebClient).Status -ne 'Running') {
            Write-Log -Text 'CRITICAL: WebClient service could not be started!' -IsError
            $script:errorsForUser += "Drive mapping failed: WebClient service is not running`n"
        }
    }
    else {
        Write-Log -Text 'WebClient service is running'
    }
}

#endregion

#region ===== AUTHENTICATION (CDP) =====

function Test-DeviceState {
    <#
    .SYNOPSIS
        Verifies whether the device is Entra ID joined and has an active PRT.
    .OUTPUTS
        Hashtable with AadJoined (bool) and HasPrt (bool).
    #>
    [CmdletBinding()]
    param()

    $output = & dsregcmd /status 2>&1 | Out-String
    $joined = $output -match 'AzureAdJoined\s*:\s*YES'
    $prt    = $output -match 'AzureAdPrt\s*:\s*YES'

    if ($joined) { Write-Log -Text 'Device is Entra ID joined' }
    else         { Write-Log -Text 'Device is NOT Entra ID joined - silent SSO may not work' -IsWarning }

    if ($prt) { Write-Log -Text 'Primary Refresh Token (PRT) is active' }
    else      { Write-Log -Text 'No active PRT detected - silent SSO may not work' -IsWarning }

    if ($output -match 'UserEmail\s*:\s*(\S+)') {
        Write-Log -Text "Signed-in user: $($Matches[1])"
    }

    return @{ AadJoined = $joined; HasPrt = $prt }
}

function Find-EdgePath {
    <#
    .SYNOPSIS
        Locates the Microsoft Edge executable.
    .OUTPUTS
        Full path to msedge.exe, or $null if not found.
    #>
    [CmdletBinding()]
    param()

    foreach ($candidate in @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
        "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
    )) {
        if (Test-Path $candidate) { return $candidate }
    }
    return $null
}

function Get-SharePointCookiesViaCDP {
    <#
    .SYNOPSIS
        Authenticates to a SharePoint/OneDrive URL using headless Edge and extracts cookies via CDP.
    .DESCRIPTION
        Launches Edge (headless or visible), waits for authentication, then extracts FedAuth/rtFa
        cookies via Chrome DevTools Protocol WebSocket. Returns cookies and the final page URL.
    .OUTPUTS
        Hashtable with FedAuth, rtFa, Host, FinalUrl - or $null on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$EdgePath,
        [int]$WaitSeconds = 10,
        [switch]$Headless,
        [int]$ManualLoginTimeout = 300
    )

    $spHost    = ([Uri]$Url).Host
    $debugPort = Get-Random -Minimum 9200 -Maximum 9400
    $tempDir   = Join-Path $env:TEMP "odm_edge_$(Get-Random)"
    $edgeProc  = $null

    try {
        # Build launch arguments
        $args = @(
            "--remote-debugging-port=$debugPort"
            "--user-data-dir=`"$tempDir`""
            "--no-first-run"
            "--no-default-browser-check"
            "--disable-features=msEdgeInterstitialRedirectToNTP"
        )
        if ($Headless) {
            $args += '--headless'
            $args += '--disable-gpu'
        }
        $args += "`"$Url`""

        $windowStyle = if ($Headless) { 'Hidden' } else { 'Normal' }
        $edgeProc = Start-Process -FilePath $EdgePath -ArgumentList ($args -join ' ') -PassThru -WindowStyle $windowStyle
        Write-Log -Text "Edge launched (PID $($edgeProc.Id), port $debugPort, headless=$Headless)"

        # Wait for CDP to become available
        Start-Sleep -Seconds 2
        $cdpReady = $false
        $wc = New-Object System.Net.WebClient
        for ($i = 0; $i -lt 15; $i++) {
            try {
                $null = $wc.DownloadString("http://localhost:$debugPort/json/version")
                $cdpReady = $true
                break
            }
            catch { Start-Sleep -Seconds 1 }
        }

        if (-not $cdpReady) {
            Write-Log -Text 'CDP not available after 17 seconds' -IsError
            return $null
        }

        # Wait for authentication to complete
        $spPage = $null

        if ($Headless) {
            Write-Log -Text "Waiting ${WaitSeconds}s for silent SSO..."
            Start-Sleep -Seconds $WaitSeconds

            $pages = $wc.DownloadString("http://localhost:$debugPort/json") | ConvertFrom-Json
            $spPage = $pages | Where-Object { $_.url -like "*$spHost*" -and $_.webSocketDebuggerUrl } | Select-Object -First 1

            if (-not $spPage) {
                # SSO may need more time - check if still authenticating
                $loginPage = $pages | Where-Object { $_.url -like "*login.microsoftonline.com*" -or $_.url -like "*login.microsoft.com*" }
                if ($loginPage) {
                    Write-Log -Text 'SSO still in progress, waiting 15 more seconds...'
                    Start-Sleep -Seconds 15
                    $pages = $wc.DownloadString("http://localhost:$debugPort/json") | ConvertFrom-Json
                    $spPage = $pages | Where-Object { $_.url -like "*$spHost*" -and $_.webSocketDebuggerUrl } | Select-Object -First 1
                }
            }

            if (-not $spPage) {
                Write-Log -Text "Silent SSO did not complete for $spHost" -IsWarning
                return $null
            }
        }
        else {
            # Visible mode - wait for user to complete authentication
            Write-Log -Text "Waiting for user to sign in (timeout: ${ManualLoginTimeout}s)..."
            $waited = 0
            while ($waited -lt $ManualLoginTimeout) {
                try {
                    $pages = $wc.DownloadString("http://localhost:$debugPort/json") | ConvertFrom-Json
                    $spPage = $pages | Where-Object { $_.url -like "*$spHost*" -and $_.webSocketDebuggerUrl } | Select-Object -First 1
                    if ($spPage) {
                        Write-Log -Text 'User completed authentication'
                        Start-Sleep -Seconds 2  # Let cookies settle
                        break
                    }
                }
                catch { }
                Start-Sleep -Seconds 2
                $waited += 2
            }

            if (-not $spPage) {
                Write-Log -Text "Authentication timed out after ${ManualLoginTimeout}s" -IsError
                return $null
            }
        }

        Write-Log -Text "Authenticated: $($spPage.url)"

        # Extract cookies via CDP WebSocket
        $ws  = New-Object System.Net.WebSockets.ClientWebSocket
        $cts = New-Object System.Threading.CancellationTokenSource
        $cts.CancelAfter(15000)
        $ws.ConnectAsync([Uri]$spPage.webSocketDebuggerUrl, $cts.Token).Wait()

        $cmd = "{`"id`":1,`"method`":`"Network.getCookies`",`"params`":{`"urls`":[`"https://$spHost/`"]}}"
        $buf = [System.Text.Encoding]::UTF8.GetBytes($cmd)
        $ws.SendAsync([ArraySegment[byte]]::new($buf), [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).Wait()

        $response  = ''
        $recvBuf   = New-Object byte[] 131072
        do {
            $recv = $ws.ReceiveAsync([ArraySegment[byte]]::new($recvBuf), $cts.Token).GetAwaiter().GetResult()
            $response += [System.Text.Encoding]::UTF8.GetString($recvBuf, 0, $recv.Count)
        } while (-not $recv.EndOfMessage)

        try { $ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, '', [System.Threading.CancellationToken]::None).Wait() } catch { }

        $data    = $response | ConvertFrom-Json
        $cookies = $data.result.cookies
        if (-not $cookies) {
            Write-Log -Text 'No cookies returned from CDP' -IsError
            return $null
        }

        $fedAuth = ($cookies | Where-Object { $_.name -eq 'FedAuth' }).value
        $rtFa    = ($cookies | Where-Object { $_.name -eq 'rtFa' }).value

        if (-not $fedAuth -or -not $rtFa) {
            Write-Log -Text "FedAuth/rtFa not found among $($cookies.Count) cookies" -IsError
            foreach ($c in $cookies) {
                Write-Log -Text "  Cookie: $($c.name) ($("$($c.value)".Length) chars, $($c.domain))"
            }
            return $null
        }

        Write-Log -Text "Cookies obtained: FedAuth=$($fedAuth.Length) chars, rtFa=$($rtFa.Length) chars"

        # Re-check page URL to capture any redirects that occurred during cookie extraction
        try {
            $updatedPages = $wc.DownloadString("http://localhost:$debugPort/json") | ConvertFrom-Json
            $updatedPage  = $updatedPages | Where-Object { $_.url -like "*$spHost*" -and $_.webSocketDebuggerUrl } | Select-Object -First 1
            if ($updatedPage -and $updatedPage.url -ne $spPage.url) {
                Write-Log -Text "Page redirected to: $($updatedPage.url)"
                $spPage = $updatedPage
            }
        }
        catch { <# Non-critical, use original URL #> }

        return @{
            FedAuth  = $fedAuth
            rtFa     = $rtFa
            Host     = $spHost
            FinalUrl = $spPage.url
        }
    }
    catch {
        Write-Log -Text "CDP authentication error: $_" -IsError
        return $null
    }
    finally {
        if ($ws)  { try { $ws.Dispose()  } catch { } }
        if ($cts) { try { $cts.Dispose() } catch { } }
        if ($wc)  { try { $wc.Dispose()  } catch { } }
        if ($edgeProc -and -not $edgeProc.HasExited) {
            try { $edgeProc.Kill() } catch { }
            Start-Sleep -Milliseconds 500
        }
        if (Test-Path $tempDir) {
            try { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        }
    }
}

function Invoke-Authentication {
    <#
    .SYNOPSIS
        Authenticates to a SharePoint/OneDrive URL. Tries silent SSO first, then falls back to visible.
    .OUTPUTS
        Hashtable with FedAuth, rtFa, Host, FinalUrl - or $null on total failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$EdgePath
    )

    # Attempt 1: Silent SSO via headless Edge + PRT
    $result = Get-SharePointCookiesViaCDP -Url $Url -EdgePath $EdgePath -WaitSeconds $edgeWaitSeconds -Headless
    if ($result) { return $result }

    # Attempt 2: Visible Edge for manual login
    if ($fallbackToVisibleAuth) {
        Write-Log -Text 'Falling back to visible Edge for manual authentication'
        $result = Get-SharePointCookiesViaCDP -Url $Url -EdgePath $EdgePath -ManualLoginTimeout $visibleAuthTimeout
        if ($result) { return $result }
    }

    Write-Log -Text "Authentication failed for $Url" -IsError
    return $null
}

function Set-WinINETCookies {
    <#
    .SYNOPSIS
        Injects FedAuth and rtFa cookies into the WinINET cookie jar for WebDAV.
    .OUTPUTS
        $true if both cookies were set successfully, $false otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpHost,
        [Parameter(Mandatory)][string]$FedAuth,
        [Parameter(Mandatory)][string]$RtFa
    )

    $expiry = (Get-Date).AddDays(5).ToString('R')

    try {
        # FedAuth: host-specific
        $fedAuthUrl  = "https://$SpHost"
        $fedAuthData = "$FedAuth;Expires=$expiry"
        $r1 = [WinInetCookies]::SetCookie($fedAuthUrl, 'FedAuth', $fedAuthData)
        Write-Log -Text "FedAuth cookie set for $fedAuthUrl ($($FedAuth.Length) chars)"

        # rtFa: domain-wide (.sharepoint.com) - extract base domain for cross-site cookie
        $rtFaDomain = ($SpHost -split '\.' | Select-Object -Last 2) -join '.'
        $rtFaUrl  = "https://$rtFaDomain"
        $rtFaData = "$RtFa;Expires=$expiry"
        $r2 = [WinInetCookies]::SetCookie($rtFaUrl, 'rtFa', $rtFaData)
        Write-Log -Text "rtFa cookie set for $rtFaUrl ($($RtFa.Length) chars)"

        # Also set rtFa on the specific host for WebDAV
        $r3 = [WinInetCookies]::SetCookie($fedAuthUrl, 'rtFa', $rtFaData)

        # Verify
        $existing = [WinInetCookies]::GetCookies($fedAuthUrl)
        if ($existing -and $existing -like '*FedAuth=*' -and $existing -like '*rtFa=*') {
            Write-Log -Text "Cookie injection verified for $SpHost"
        }
        else {
            Write-Log -Text "Cookie verification inconclusive for $SpHost" -IsWarning
        }

        return ($r1 -and $r2 -and $r3)
    }
    catch {
        Write-Log -Text "Cookie injection failed for $SpHost : $_" -IsError
        return $false
    }
}

#endregion

#region ===== DRIVE MAPPING =====

function Add-NetworkLocation {
    [CmdLetBinding()]
    param(
        [string]$NetworkLocationPath = "$env:APPDATA\Microsoft\Windows\Network Shortcuts",
        [Parameter(Mandatory)][string]$NetworkLocationName,
        [Parameter(Mandatory)][string]$NetworkLocationTarget,
        [string]$IconPath
    )

    $desktopIni = "[.ShellClassInfo]`r`nCLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}`r`nFlags=2"
    $dir = Join-Path $NetworkLocationPath $NetworkLocationName

    if (-not (Test-Path $NetworkLocationPath -PathType Container)) {
        Write-Error "'$NetworkLocationPath' is not a valid directory."
        return $false
    }

    try {
        if (-not (Test-Path $dir -PathType Container)) {
            $null = New-Item -Path $dir -ItemType Directory -ErrorAction Stop
            Set-ItemProperty -Path $dir -Name Attributes -Value ([IO.FileAttributes]::System) -ErrorAction Stop
        }

        $iniPath = Join-Path $dir 'desktop.ini'
        if (-not (Test-Path $iniPath)) { $null = New-Item -Path $iniPath -ItemType File -ErrorAction Stop }
        Set-Content -Path $iniPath -Value $desktopIni -ErrorAction Stop

        $shell = New-Object -ComObject WScript.Shell
        $lnk   = $shell.CreateShortcut((Join-Path $dir 'target.lnk'))
        $lnk.TargetPath = $NetworkLocationTarget
        if ($IconPath -and [IO.File]::Exists($IconPath)) { $lnk.IconLocation = "$IconPath, 0" }
        $lnk.Description = "Created $(Get-Date -Format 's') by OneDriveMapper"
        $lnk.Save()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)

        return $true
    }
    catch {
        Write-Error "Failed to create network location '$dir': $_"
        return $false
    }
}

function New-FavoritesShortcut {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$TargetLocation)

    $regPath  = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
    $guid     = '{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}'
    $linksDir = (Get-ItemProperty -Path $regPath -Name $guid -ErrorAction Stop).$guid
    $lnkFile  = Join-Path $linksDir "OneDrive - $O365CustomerName.lnk"

    if ([IO.File]::Exists($lnkFile)) {
        Write-Log -Text 'Favorites shortcut already exists'
        return
    }

    $shell = New-Object -ComObject WScript.Shell
    $lnk   = $shell.CreateShortcut($lnkFile)
    $lnk.TargetPath = $TargetLocation
    if ($onedriveIconPath -and [IO.File]::Exists($onedriveIconPath)) { $lnk.IconLocation = "$onedriveIconPath, 0" }
    $lnk.Description = 'OneDrive for Business'
    $lnk.Save()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
}

function Set-DriveLabel {
    [CmdletBinding()]
    param(
        [string]$DriveLetter,
        [string]$MapURL,
        [string]$DriveLabel
    )
    if ([string]::IsNullOrWhiteSpace($DriveLabel)) { return }

    try {
        $regURL   = ($MapURL.TrimEnd('\') -replace [regex]::Escape('\'), '#')
        $basePath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2'
        foreach ($url in @($regURL, ($regURL -replace [regex]::Escape('DavWWWRoot#'), ''))) {
            $regPath = "$basePath\$url"
            $null = New-Item -Path $regPath -Force -ErrorAction SilentlyContinue
            $null = New-ItemProperty -Path $regPath -Name '_CommentFromDesktopINI' -ErrorAction SilentlyContinue
            $null = New-ItemProperty -Path $regPath -Name '_LabelFromDesktopINI' -ErrorAction SilentlyContinue
            $null = New-ItemProperty -Path $regPath -Name '_LabelFromReg' -Value $DriveLabel -ErrorAction SilentlyContinue
        }
        Write-Log -Text "Drive label set: $DriveLetter = '$DriveLabel'"
        [WinAPI.Explorer]::Refresh()
    }
    catch {
        Write-Log -Text "Failed to set drive label: $_" -IsError
    }
}

function Test-NetUseErrors {
    [CmdletBinding()]
    param([object]$Output, [string]$DriveLetter, [string]$WebDavPath)

    if ($Output -like '*error 67*') {
        Write-Log -Text 'NET USE error 67: WebClient service not running or URL not trusted' -IsError
    }
    if ($Output -like '*error 224*') {
        Write-Log -Text 'NET USE error 224: trusted sites misconfigured or KB2846960 missing' -IsError
    }
    if ($LASTEXITCODE -ne 0) {
        Write-Log -Text "NET USE failed: $DriveLetter -> $WebDavPath (exit $LASTEXITCODE) $Output" -IsError
        $script:errorsForUser += "$DriveLetter could not be mapped (error $LASTEXITCODE)`n"
    }
}

function Invoke-DriveMapping {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$DriveMapping)

    $persist = if ($persistentMapping) { '/PERSISTENT:YES' } else { '/PERSISTENT:NO' }

    if ($DriveMapping.targetLocationType -eq 'driveletter') {
        return Invoke-DriveLetterMapping -DriveMapping $DriveMapping -PersistenceFlag $persist
    }
    else {
        return Invoke-NetworkLocationMapping -DriveMapping $DriveMapping -PersistenceFlag $persist
    }
}

function Invoke-DriveLetterMapping {
    [CmdletBinding()]
    param([hashtable]$DriveMapping, [string]$PersistenceFlag)

    Write-Log -Text "Mapping $($DriveMapping.targetLocationPath) -> $($DriveMapping.webDavPath)"

    # Remove existing mapping for this letter
    $LASTEXITCODE = 0
    try { $null = NET USE $($DriveMapping.targetLocationPath) /DELETE /Y 2>&1 } catch { }

    # First pass: prime the WebDAV connection
    try { $out = NET USE $($DriveMapping.webDavPath) $PersistenceFlag 2>&1 } catch { }
    Test-NetUseErrors -Output $out -DriveLetter $DriveMapping.targetLocationPath -WebDavPath $DriveMapping.webDavPath

    # Create per-user folder if configured
    if ($createUserFolderOn -and $createUserFolderOn -like "*$($DriveMapping.targetLocationPath)*") {
        $userFolder = Join-Path $DriveMapping.webDavPath $env:USERNAME
        if (-not (Test-Path $userFolder)) {
            Write-Log -Text "Creating user folder: $userFolder"
            $null = New-Item -Path $userFolder -ItemType Directory -Force -Confirm:$false
        }
        $DriveMapping.webDavPath = $userFolder
    }

    # Second pass: map with drive letter
    $LASTEXITCODE = 0
    try { $null = NET USE $($DriveMapping.targetLocationPath) /DELETE /Y 2>&1 } catch { }
    try { $out = NET USE $($DriveMapping.targetLocationPath) $($DriveMapping.webDavPath) $PersistenceFlag 2>&1 } catch { }
    Test-NetUseErrors -Output $out -DriveLetter $DriveMapping.targetLocationPath -WebDavPath $DriveMapping.webDavPath

    if (Test-Path $DriveMapping.webDavPath) {
        Set-DriveLabel -DriveLetter $DriveMapping.targetLocationPath -MapURL $DriveMapping.webDavPath -DriveLabel $DriveMapping.displayName
        Write-Log -Text "$($DriveMapping.targetLocationPath) mapped successfully"
        return $true
    }
    else {
        Write-Log -Text "Mapping verification failed: $($DriveMapping.targetLocationPath)" -IsError
        return $false
    }
}

function Invoke-NetworkLocationMapping {
    [CmdletBinding()]
    param([hashtable]$DriveMapping, [string]$PersistenceFlag)

    try {
        $icon = if ($DriveMapping.sourceLocationPath -eq 'autodetect') { $onedriveIconPath } else { $sharepointIconPath }

        Write-Log -Text "Mapping network location: $($DriveMapping.webDavPath)"
        try { $null = NET USE $($DriveMapping.webDavPath) /DELETE /Y 2>&1 } catch { }
        try { $null = NET USE $($DriveMapping.webDavPath) $PersistenceFlag 2>&1 } catch { }

        $null = Add-NetworkLocation `
            -NetworkLocationPath $DriveMapping.targetLocationPath `
            -NetworkLocationName $DriveMapping.displayName `
            -NetworkLocationTarget $DriveMapping.webDavPath `
            -IconPath $icon -ErrorAction Stop

        if (Test-Path $DriveMapping.webDavPath) {
            Write-Log -Text "Network location added: '$($DriveMapping.displayName)'"
            return $true
        }
        else {
            Write-Log -Text "Network location verification failed: $($DriveMapping.webDavPath)" -IsError
            return $false
        }
    }
    catch {
        Write-Log -Text "Failed to add network location: $_" -IsError
        return $false
    }
}

#endregion

#region ===== FOLDER REDIRECTION =====

function Set-KnownFolderPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Desktop', 'Documents', 'Downloads', 'Music', 'Pictures', 'Videos',
            'Favorites', 'Contacts', 'Links', 'SavedGames', 'SavedSearches')]
        [string]$KnownFolder,
        [Parameter(Mandatory)][string]$Path
    )

    $KnownFolders = @{
        'Desktop'       = @('B4BFCC3A-DB2C-424C-B029-7FE99A87C641')
        'Documents'     = @('FDD39AD0-238F-46AF-ADB4-6C85480369C7','f42ee2d3-909f-4907-8871-4c22fc0bf756')
        'Downloads'     = @('374DE290-123F-4565-9164-39C4925E467B','7d83ee9b-2244-4e70-b1f5-5393042af1e4')
        'Music'         = @('4BD8D571-6D19-48D3-BE97-422220080E43','a0c69a99-21c8-4671-8703-7934162fcf1d')
        'Pictures'      = @('33E28130-4E1E-4676-835A-98395C3BC3BB','0ddd015d-b06c-45d5-8c4c-f59713854639')
        'Videos'        = @('18989B1D-99B5-455B-841C-AB7C74E4DDFC','35286a68-3c57-41a1-bbb1-0eae73d76c95')
        'Favorites'     = @('1777F761-68AD-4D8A-87BD-30B759FA33DD')
        'Contacts'      = @('56784854-C6CB-462b-8169-88E350ACB882')
        'Links'         = @('bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968')
        'SavedGames'    = @('4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4')
        'SavedSearches' = @('7d1d3a04-debb-4115-95cf-2f29da2920da')
    }

    $Type = ([System.Management.Automation.PSTypeName]'KnownFolders').Type
    if (-not $Type) {
        $Sig = @'
[DllImport("shell32.dll")]
public extern static int SHSetKnownFolderPath(ref Guid folderId, uint flags, IntPtr token, [MarshalAs(UnmanagedType.LPWStr)] string path);
'@
        $Type = Add-Type -MemberDefinition $Sig -Name 'KnownFolders' -Namespace 'SHSetKnownFolderPath' -PassThru
    }

    if (-not (Test-Path $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }

    foreach ($guid in $KnownFolders[$KnownFolder]) {
        $hr = $Type::SHSetKnownFolderPath([ref]$guid, 0, 0, $Path)
        if ($hr -ne 0) {
            throw "SHSetKnownFolderPath($KnownFolder) HRESULT: $hr - $(([ComponentModel.Win32Exception]$hr).Message)"
        }
    }

    [WinAPI.Explorer]::Refresh()
    attrib +r $Path
    return $Path
}

function Get-KnownFolderPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Desktop', 'DesktopDirectory', 'MyDocuments', 'MyMusic', 'MyPictures', 'MyVideos',
            'Favorites', 'Personal', 'UserProfile')]
        [string]$KnownFolder
    )
    return [Environment]::GetFolderPath($KnownFolder)
}

function Invoke-FolderRedirection {
    [CmdletBinding()]
    param(
        [string]$GetFolder,
        [string]$SetFolder,
        [string]$Target,
        [string]$CopyExistingFiles
    )

    $currentPath = Get-KnownFolderPath -KnownFolder $GetFolder
    if ($currentPath -ne $Target) {
        Set-KnownFolderPath -KnownFolder $SetFolder -Path $Target
        if ($CopyExistingFiles -eq 'true' -and $currentPath) {
            Get-ChildItem -Path $currentPath -ErrorAction Continue |
                Copy-Item -Destination $Target -Recurse -Container -Force -Confirm:$false -ErrorAction Continue
        }
        attrib +h $currentPath
    }
}

#endregion

#region ===== ELEVATION BYPASS =====

function Invoke-ElevationBypass {
    [CmdletBinding()]
    param()

    $windowArg = if ($showConsoleOutput) { '' } else { ' -WindowStyle Hidden' }
    $taskCmd   = "Powershell.exe -NoProfile -ExecutionPolicy ByPass$windowArg -File '$scriptPath\OneDriveMapper.ps1' -AsTask"
    $result    = schtasks "/Create /SC ONCE /TN OneDriveMapper /IT /RL LIMITED /F /TR `"$taskCmd`" /st 00:00".Split(' ') 2>&1

    if ($result -notmatch 'ERROR') {
        Write-Log -Text 'Created scheduled task for unelevated execution'
        $run = schtasks /Run /TN OneDriveMapper /I 2>&1
        if ($run -notmatch 'ERROR') {
            Write-Log -Text 'Scheduled task started successfully'
        }
        else {
            Write-Log -Text "Failed to start task: $run" -IsError
        }
    }
    else {
        Write-Log -Text "Failed to create task: $result" -IsError
    }
}

#endregion

#region ===== MAIN EXECUTION =====

# Determine script path
$scriptPath = if ($PSScriptRoot) { $PSScriptRoot }
              elseif (Test-Path Variable:psISE) { Split-Path $psISE.CurrentFile.FullPath }
              else { $PWD.Path }
$scriptPath = $scriptPath -replace 'Microsoft\.PowerShell\.Core\\FileSystem::', ''

Reset-LogFile
Write-Log -Text "===== OneDriveMapper v$version - $env:USERNAME on $env:COMPUTERNAME ====="
Write-Log -Text "Script path: $scriptPath"

# Elevation check
$script:isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

if ($script:isElevated) {
    Write-Log -Text 'Running elevated (Administrator)' -IsWarning

    $scheduleTask = $true
    $uac = Get-RegistryValue -BasePath 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System' -EntryName 'EnableLUA'
    if ($uac -eq 0)  { Write-Log -Text 'UAC is disabled'; $scheduleTask = $false }
    if ($AsTask)      { Write-Log -Text 'Already running as task'; $scheduleTask = $false }

    Test-WebClient

    if ($scheduleTask) {
        Invoke-ElevationBypass
        exit
    }
}
else {
    Write-Log -Text 'Running as standard user'
    Test-WebClient
}

# Find Microsoft Edge
$edgePath = Find-EdgePath
if (-not $edgePath) {
    Write-Log -Text 'Microsoft Edge not found! Cannot authenticate.' -IsError
    $script:errorsForUser += "Microsoft Edge not found`n"
    exit
}
$edgeVersion = try { (Get-ItemProperty $edgePath -ErrorAction Stop).VersionInfo.ProductVersion } catch { 'unknown' }
Write-Log -Text "Edge: $edgePath (v$edgeVersion)"
Write-Log -Text "PowerShell: $($PSVersionTable.PSVersion) | Windows: $([Environment]::OSVersion.Version)"

# Verify device state for PRT-based SSO
$deviceState = Test-DeviceState

# WebDAV configuration check
$lockingVal = Get-RegistryValue -BasePath 'HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters' -EntryName 'SupportLocking'
if ($lockingVal -ne 0) { Write-Log -Text 'WebDAV file locking is enabled (can cause issues)' -IsWarning }

$fileSizeLimit = Get-RegistryValue -BasePath 'HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters' -EntryName 'FileSizeLimitInBytes'
if ($fileSizeLimit -gt 0) { Write-Log -Text "WebDAV max file size: $([Math]::Round($fileSizeLimit / 1MB)) MB" }

# IE trusted sites check and auto-add
$regHKLM = Get-RegistryValue -BasePath 'HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings' -EntryName 'Security_HKLM_only'

foreach ($suffix in @('', $privateSuffix)) {
    $siteName = "$O365CustomerName$suffix.sharepoint.com"
    $found = $false
    foreach ($path in @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$O365CustomerName$suffix",
        "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$O365CustomerName$suffix",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$O365CustomerName$suffix",
        "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$O365CustomerName$suffix"
    )) {
        if ($regHKLM -eq 1 -and $path.StartsWith('HKCU:')) { continue }
        if ((Get-RegistryValue -BasePath $path -EntryName 'https') -match '^[1-2]+$') {
            $found = $true
            break
        }
    }
    if ($found) {
        Write-Log -Text "$siteName found in trusted sites"
    }
    else {
        Write-Log -Text "$siteName not in trusted sites - adding..." -IsWarning
        if ((Add-SiteToIEZone -SiteUrl $siteName) -eq $true) {
            Write-Log -Text "Added $siteName to trusted sites"
        }
        else {
            Write-Log -Text "Failed to add $siteName to trusted sites" -IsError
        }
    }
}

# Prevent copy/paste warnings
try {
    $null = New-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com@SSL" -ErrorAction SilentlyContinue
    $null = New-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com@SSL\$O365CustomerName" -ErrorAction SilentlyContinue
    $null = New-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com@SSL\$O365CustomerName" -Name 'file' -Value 1 -ErrorAction SilentlyContinue
}
catch { }

# Disable IE auto-proxy if configured
if (-not $autoDetectProxy) {
    $proxyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
    if ((Get-RegistryValue -BasePath $proxyPath -EntryName 'AutoDetect') -ne 0) {
        try {
            $null = New-ItemProperty $proxyPath -Name 'AutoDetect' -Value 0 -ErrorAction Stop
            Write-Log -Text 'Disabled IE auto-proxy detection'
        }
        catch { }
    }
}

# Explorer status check
$explorerStatus = Get-ExplorerProcessForCurrentUser
if ($explorerStatus -eq 0) { Write-Log -Text 'No Explorer instances running' -IsWarning }

# Remove existing SharePoint/OneDrive mappings
if ($removeExistingMaps) {
    Write-Log -Text 'Removing existing SharePoint/OneDrive mappings...'
    subst 2>$null | ForEach-Object { subst $_.Substring(0, 2) /D 2>$null }
    Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot } | ForEach-Object {
        if ($_.DisplayRoot.StartsWith("\\$O365CustomerName.sharepoint.com") -or
            $_.DisplayRoot.StartsWith("\\$O365CustomerName-my.sharepoint.com")) {
            try { $null = NET USE "$($_.Name):" /DELETE /Y 2>&1 } catch { }
        }
    }
}

# Remove empty/dead drive mappings
if ($removeEmptyMaps) {
    Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -eq 0 -and $null -eq $_.Free } | ForEach-Object {
        try { $_ | Remove-PSDrive -Force } catch { }
    }
}

# Resolve AD group membership for conditional mappings
$groups = $null
if ($desiredMappings | Where-Object { $_.mapOnlyForSpecificGroup.Length -gt 0 }) {
    try {
        $searcher = [adsisearcher]"samaccountname=$env:USERNAME"
        $userDN = $searcher.FindOne().Properties.distinguishedname
        $groupSearcher = [adsisearcher]"(member:1.2.840.113556.1.4.1941:=$userDN)"
        $groups = $groupSearcher.FindAll().Properties.distinguishedname -replace '^CN=([^,]+).+$', '$1'
        Write-Log -Text 'Resolved user group membership for conditional mappings'
    }
    catch {
        Write-Log -Text "Failed to resolve group membership: $_" -IsWarning
        $desiredMappings = $desiredMappings | Where-Object { $_.mapOnlyForSpecificGroup.Length -eq 0 }
    }
}

# Build WebDAV paths and filter mappings
$mapURLpersonal = "\\$O365CustomerName-my.sharepoint.com@SSL\DavWWWRoot\personal\"
$intendedMappings = [System.Collections.Generic.List[hashtable]]::new()

foreach ($mapping in $desiredMappings) {
    if ($mapping.sourceLocationPath -ne 'autodetect') {
        $webDav = [System.Web.HttpUtility]::UrlDecode($mapping.sourceLocationPath)
        $webDav = $webDav.Replace('https://', '\\')
        $webDav = $webDav.Replace('/_layouts/15/start.aspx#', '')
        $webDav = $webDav.Replace('sharepoint.com', 'sharepoint.com@SSL\DavWWWRoot')
        $webDav = $webDav.Replace('/Forms/AllItems.aspx', '')
        $webDav = $webDav.Replace("%27", "'")
        $webDav = $webDav.Replace('/', '\')
        $mapping.webDavPath = $webDav
    }
    else {
        $mapping.webDavPath = $mapURLpersonal  # Will be updated after autodetect
    }

    # Check group filter
    if ($mapping.mapOnlyForSpecificGroup -and $groups) {
        if ($groups -notcontains $mapping.mapOnlyForSpecificGroup) { continue }
        Write-Log -Text "Including '$($mapping.displayName)': user is member of '$($mapping.mapOnlyForSpecificGroup)'"
    }

    $intendedMappings.Add($mapping)
}

Write-Log -Text "Prepared $($intendedMappings.Count) drive mapping(s)"

# Prepare converged drives
$convergedDrives = @($intendedMappings | Where-Object { $_.targetLocationType -eq 'converged' })
if ($convergedDrives.Count -gt 0) {
    $convergedLetters = $convergedDrives.targetLocationPath | Select-Object -Unique
    foreach ($letter in $convergedLetters) {
        $folder = Join-Path $env:TEMP "OneDriveMapperLinks $($letter.Substring(0, 1))"
        if (-not [IO.Directory]::Exists($folder)) {
            $null = New-Item -Path $folder -ItemType Directory -Force
        }
        else {
            Get-ChildItem $folder | Remove-Item -Force -Recurse -Confirm:$false -ErrorAction SilentlyContinue
        }
        $null = subst $letter $folder 2>$null
        Set-DriveLabel -DriveLetter $letter -MapURL $letter.Substring(0, 1) -DriveLabel $convergedDriveLabel
    }
}

#endregion

#region ===== MAIN LOOP =====

while ($true) {

    # ---- Progress bar (modern dark toast) ----
    $form1 = $null
    $progressFill = $null
    $label1 = $null
    $progressTrackW = 0
    $accentColor = [Drawing.ColorTranslator]::FromHtml($progressBarColor)

    if ($showProgressBar) {
        $w = 380; $h = 56; $pad = 16; $trackH = 4
        $progressTrackW = $w - $pad * 2

        $form1 = New-Object Windows.Forms.Form
        $form1.Text = "OneDriveMapper v$version"
        $form1.Size = $form1.MaximumSize = $form1.MinimumSize = New-Object Drawing.Size($w, $h)
        $form1.BackColor = [Drawing.Color]::FromArgb(45, 45, 48)
        $form1.ControlBox = $false
        $form1.FormBorderStyle = 'None'
        $form1.ShowInTaskbar = $false
        $form1.StartPosition = 'Manual'
        $form1.Location = New-Object Drawing.Point(-9999, -9999)  # Start off-screen
        $form1.TopMost = $true
        $form1.Opacity = 0.95

        # Rounded corners
        $radius = 8
        $gp = New-Object Drawing.Drawing2D.GraphicsPath
        $gp.AddArc(0, 0, $radius * 2, $radius * 2, 180, 90)
        $gp.AddArc($w - $radius * 2 - 1, 0, $radius * 2, $radius * 2, 270, 90)
        $gp.AddArc($w - $radius * 2 - 1, $h - $radius * 2 - 1, $radius * 2, $radius * 2, 0, 90)
        $gp.AddArc(0, $h - $radius * 2 - 1, $radius * 2, $radius * 2, 90, 90)
        $gp.CloseFigure()
        $form1.Region = New-Object Drawing.Region($gp)
        $gp.Dispose()

        $label1 = New-Object Windows.Forms.Label
        $label1.Text = $progressBarText
        $label1.Location = New-Object Drawing.Point($pad, 12)
        $label1.Size = New-Object Drawing.Size(($w - $pad * 2), 20)
        $label1.Font = New-Object Drawing.Font('Segoe UI', 9)
        $label1.ForeColor = [Drawing.Color]::FromArgb(240, 240, 240)
        $label1.BackColor = [Drawing.Color]::Transparent

        $progressTrack = New-Object Windows.Forms.Panel
        $progressTrack.Location = New-Object Drawing.Point($pad, 38)
        $progressTrack.Size = New-Object Drawing.Size($progressTrackW, $trackH)
        $progressTrack.BackColor = [Drawing.Color]::FromArgb(70, 70, 70)

        $progressFill = New-Object Windows.Forms.Panel
        $progressFill.Location = New-Object Drawing.Point(0, 0)
        $progressFill.Size = New-Object Drawing.Size(0, $trackH)
        $progressFill.BackColor = $accentColor

        $progressTrack.Controls.Add($progressFill)
        $form1.Controls.AddRange(@($label1, $progressTrack))
        [void]$form1.Show()
        # Position after Show - all layout resets have already happened at (-9999,-9999)
        $screen = ([Windows.Forms.Screen]::AllScreens | Where-Object { $_.Primary }).WorkingArea
        $form1.SetDesktopLocation(($screen.Right - $w - 12), ($screen.Bottom - $h - 12))
        [void]$form1.Focus()
        $progressFill.Width = [int]($progressTrackW * 5 / 100)
        $form1.Refresh()
    }

    # ---- System tray icon (runs in its own runspace with a proper message loop) ----
    if ($showSystemTrayIcon -and -not $script:traySync) {
        try {
            $script:traySync = [hashtable]::Synchronized(@{
                Text          = "OneDriveMapper v$version"
                BalloonTitle  = ''
                BalloonMsg    = ''
                BalloonIcon   = 'Info'
                ShowBalloon   = $false
                ExitRequested = $false
                LogFile       = $logfile
            })
            $script:trayRunspace = [runspacefactory]::CreateRunspace()
            $script:trayRunspace.ApartmentState = 'STA'
            $script:trayRunspace.ThreadOptions = 'ReuseThread'
            $script:trayRunspace.Open()
            $script:trayRunspace.SessionStateProxy.SetVariable('sync', $script:traySync)

            $script:trayPS = [powershell]::Create().AddScript({
                param($accentHtml, $ver)
                [void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')
                [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

                $accentColor = [Drawing.ColorTranslator]::FromHtml($accentHtml)
                $icon = New-Object Windows.Forms.NotifyIcon

                # Draw cloud icon with upload arrow
                $bmp = New-Object Drawing.Bitmap(16, 16)
                $g = [Drawing.Graphics]::FromImage($bmp)
                $g.SmoothingMode = 'AntiAlias'
                $g.Clear([Drawing.Color]::Transparent)
                $fill = New-Object Drawing.SolidBrush($accentColor)
                $g.FillEllipse($fill, 2, 5, 12, 9)
                $g.FillEllipse($fill, 4, 2, 8, 8)
                $g.FillEllipse($fill, 1, 6, 6, 7)
                $g.FillEllipse($fill, 9, 6, 6, 7)
                $pen = New-Object Drawing.Pen([Drawing.Color]::White, 1.6)
                $pen.StartCap = $pen.EndCap = [Drawing.Drawing2D.LineCap]::Round
                $g.DrawLine($pen, 8, 12, 8, 7)
                $g.DrawLine($pen, 5.5, 9.5, 8, 7)
                $g.DrawLine($pen, 10.5, 9.5, 8, 7)
                $pen.Dispose(); $fill.Dispose(); $g.Dispose()
                $icon.Icon = [Drawing.Icon]::FromHandle($bmp.GetHicon())
                $bmp.Dispose()

                $icon.Text = "OneDriveMapper v$ver"
                $icon.Visible = $true

                # Context menu
                $menu = New-Object Windows.Forms.ContextMenuStrip
                $logItem = New-Object Windows.Forms.ToolStripMenuItem('Open log file')
                $logItem.Add_Click({
                    $lf = $sync.LogFile
                    if ($lf -and (Test-Path $lf)) { Start-Process notepad.exe $lf }
                })
                $sep = New-Object Windows.Forms.ToolStripSeparator
                $exitItem = New-Object Windows.Forms.ToolStripMenuItem('Exit OneDriveMapper')
                $exitItem.Add_Click({
                    $sync.ExitRequested = $true
                    $icon.Visible = $false; $icon.Dispose()
                    [Windows.Forms.Application]::ExitThread()
                })
                [void]$menu.Items.Add($logItem)
                [void]$menu.Items.Add($sep)
                [void]$menu.Items.Add($exitItem)
                $icon.ContextMenuStrip = $menu

                # Timer polls the sync hashtable for updates from the main script
                $timer = New-Object Windows.Forms.Timer
                $timer.Interval = 250
                $timer.Add_Tick({
                    if ($sync.ExitRequested) {
                        $timer.Stop()
                        $icon.Visible = $false; $icon.Dispose()
                        [Windows.Forms.Application]::ExitThread()
                        return
                    }
                    $icon.Text = $sync.Text
                    if ($sync.ShowBalloon) {
                        $sync.ShowBalloon = $false
                        $tipIcon = switch ($sync.BalloonIcon) {
                            'Warning' { [Windows.Forms.ToolTipIcon]::Warning }
                            'Error'   { [Windows.Forms.ToolTipIcon]::Error }
                            default   { [Windows.Forms.ToolTipIcon]::Info }
                        }
                        $icon.ShowBalloonTip(3000, $sync.BalloonTitle, $sync.BalloonMsg, $tipIcon)
                    }
                })
                $timer.Start()

                [Windows.Forms.Application]::Run()
            }).AddArgument($progressBarColor).AddArgument($version)

            $script:trayPS.Runspace = $script:trayRunspace
            $null = $script:trayPS.BeginInvoke()
        }
        catch { <# Non-critical if tray icon fails #> }
    }
    if ($script:traySync) { $script:traySync.Text = "OneDriveMapper v$version - Connecting..." }

    # ---- Determine unique hosts that need authentication ----
    $authTargets = [ordered]@{}  # host -> URL

    foreach ($mapping in $intendedMappings) {
        if ($mapping.sourceLocationPath -eq 'autodetect') {
            $spHost = "$O365CustomerName$privateSuffix.sharepoint.com"
            $url    = "https://$spHost/my"
        }
        else {
            $uri    = [Uri]$mapping.sourceLocationPath
            $spHost = $uri.Host
            $url    = $mapping.sourceLocationPath
        }
        if (-not $authTargets.Contains($spHost)) {
            $authTargets[$spHost] = $url
        }
    }

    Write-Log -Text "Authenticating to $($authTargets.Count) unique host(s): $($authTargets.Keys -join ', ')"

    if ($showProgressBar) { $progressFill.Width = [int]($progressTrackW * 10 / 100); $form1.Refresh() }
    [Windows.Forms.Application]::DoEvents()

    # ---- Authenticate and inject cookies for each host ----
    $authResults     = @{}
    $progressPerHost = if ($authTargets.Count -gt 0) { [int](60 / $authTargets.Count) } else { 60 }
    $hostIndex       = 0

    foreach ($authHost in $authTargets.Keys) {
        $authUrl = $authTargets[$authHost]
        Write-Log -Text "Authenticating to $authHost..."

        $result = Invoke-Authentication -Url $authUrl -EdgePath $edgePath

        if ($result) {
            $authResults[$authHost] = $result
            $injected = Set-WinINETCookies -SpHost $result.Host -FedAuth $result.FedAuth -RtFa $result.rtFa
            if (-not $injected) {
                Write-Log -Text "Cookie injection failed for $authHost" -IsError
            }
        }
        else {
            Write-Log -Text "Authentication failed for $authHost - drives for this host will not be mapped" -IsError
            $script:errorsForUser += "Authentication failed for $authHost`n"
        }

        $hostIndex++
        if ($showProgressBar) {
            $progressFill.Width = [int]($progressTrackW * [Math]::Min(10 + ($hostIndex * $progressPerHost), 70) / 100)
            $form1.Refresh()
        }
        [Windows.Forms.Application]::DoEvents()
    }

    # ---- Resolve OneDrive autodetect user slug ----
    $odHost = "$O365CustomerName$privateSuffix.sharepoint.com"
    if ($authResults.ContainsKey($odHost)) {
        $finalUrl = $authResults[$odHost].FinalUrl
        $userSlug = $null

        # Method 1: Extract from redirect URL
        $m = [regex]::Match($finalUrl, '/personal/([^/?#]+)')
        if ($m.Success) {
            $userSlug = $m.Groups[1].Value
            Write-Log -Text "OneDrive user detected from URL: $userSlug"
        }

        # Method 2: Derive from UPN (e.g. user@domain.com -> user_domain_com)
        if (-not $userSlug) {
            try {
                $upn = (whoami /upn 2>$null)
                if ($upn) {
                    $upn = $upn.Trim()
                    $userSlug = $upn -replace '@', '_' -replace '\.', '_'
                    Write-Log -Text "OneDrive user derived from UPN ($upn): $userSlug"
                }
            }
            catch { }
        }

        if ($userSlug) {
            $personalWebDav = "${mapURLpersonal}${userSlug}\$libraryName"
            Write-Log -Text "OneDrive WebDAV path: $personalWebDav"

            # Update all autodetect mappings with the resolved path
            foreach ($mapping in $intendedMappings) {
                if ($mapping.sourceLocationPath -eq 'autodetect') {
                    $mapping.webDavPath = $personalWebDav
                }
            }
        }
        else {
            Write-Log -Text "Could not determine OneDrive user slug" -IsError
            $script:errorsForUser += "Could not detect your OneDrive username`n"
        }
    }

    if ($showProgressBar) { $progressFill.Width = [int]($progressTrackW * 75 / 100); $form1.Refresh() }
    [Windows.Forms.Application]::DoEvents()

    # ---- Map drives ----
    for ($i = 0; $i -lt $intendedMappings.Count; $i++) {
        # Skip if the host authentication failed
        $mappingHost = if ($intendedMappings[$i].sourceLocationPath -eq 'autodetect') {
            $odHost
        }
        else {
            ([Uri]$intendedMappings[$i].sourceLocationPath).Host
        }

        if (-not $authResults.ContainsKey($mappingHost)) {
            Write-Log -Text "Skipping '$($intendedMappings[$i].displayName)': no auth for $mappingHost" -IsWarning
            $intendedMappings[$i].mapped = $false
            continue
        }

        $intendedMappings[$i].mapped = Invoke-DriveMapping -DriveMapping $intendedMappings[$i]

        # Favorites shortcut for OneDrive
        if ($intendedMappings[$i].sourceLocationPath -eq 'autodetect' -and
            $addShellLink -and
            $intendedMappings[$i].targetLocationType -eq 'driveletter' -and
            [IO.Directory]::Exists($intendedMappings[$i].targetLocationPath)) {
            try { New-FavoritesShortcut -TargetLocation $intendedMappings[$i].targetLocationPath }
            catch { Write-Log -Text "Failed to create favorites shortcut: $_" -IsError }
        }

        if ($showProgressBar) {
            $progressFill.Width = [int]($progressTrackW * [Math]::Min(75 + ($i * 5), 90) / 100)
            $form1.Refresh()
        }
        [Windows.Forms.Application]::DoEvents()
    }

    # ---- Close progress bar ----
    if ($showProgressBar -and $form1) {
        $progressFill.Width = $progressTrackW
        $label1.Text = 'Done!'
        $form1.Refresh()
        Start-Sleep -Milliseconds 600
        $form1.Close()
        $form1.Dispose()
        $form1 = $null
    }

    # ---- Folder redirection ----
    if ($redirectFolders) {
        foreach ($folder in $listOfFoldersToRedirect) {
            Write-Log -Text "Redirecting $($folder.knownFolderInternalName) -> $($folder.desiredTargetPath)"
            try {
                Invoke-FolderRedirection `
                    -GetFolder $folder.knownFolderInternalName `
                    -SetFolder $folder.knownFolderInternalIdentifier `
                    -Target $folder.desiredTargetPath `
                    -CopyExistingFiles $folder.copyExistingFiles
                Write-Log -Text "Redirected $($folder.knownFolderInternalName)"
            }
            catch {
                Write-Log -Text "Folder redirection failed for $($folder.knownFolderInternalName): $_" -IsError
            }
        }
    }

    # ---- Summary ----
    $successCount = @($intendedMappings | Where-Object { $_.mapped }).Count
    $failCount    = $intendedMappings.Count - $successCount
    Write-Log -Text "Mapping complete: $successCount succeeded, $failCount failed"
    if ($script:traySync) {
        $trayMsg = if ($failCount -eq 0) { "$successCount drive(s) connected" } else { "$successCount OK, $failCount failed" }
        $script:traySync.Text = "OneDriveMapper v$version - $trayMsg"
        $script:traySync.BalloonTitle = 'OneDriveMapper'
        $script:traySync.BalloonMsg = $trayMsg
        $script:traySync.BalloonIcon = if ($failCount -eq 0) { 'Info' } else { 'Warning' }
        $script:traySync.ShowBalloon = $true
    }

    foreach ($mapping in $intendedMappings) {
        $status = if ($mapping.mapped) { 'OK' } else { 'FAILED' }
        Write-Log -Text "  [$status] $($mapping.displayName) -> $($mapping.webDavPath)"
    }

    # ---- Cleanup ----

    if ($restartExplorer) {
        Restart-ExplorerProcess
    }
    elseif ($redirectFolders) {
        Write-Log -Text 'Tip: set $restartExplorer=$true for redirected folders to appear immediately' -IsWarning
    }

    if ($urlOpenAfter.Length -gt 10) {
        Start-Process msedge.exe $urlOpenAfter
    }

    if ($displayErrors -and $script:errorsForUser) {
        [void][Windows.Forms.MessageBox]::Show($script:errorsForUser, 'OneDriveMapper Error', 0)
        [void][Windows.Forms.MessageBox]::Show(
            'You can always access your files at https://portal.office.com',
            'Need a workaround?', 0
        )
    }

    # ---- Auto-remap monitoring ----
    if ($autoRemapMethod -eq 'Disabled') { break }

    $successfulMappings = @($intendedMappings | Where-Object { $_.mapped })
    if ($successfulMappings.Count -eq 0) {
        Write-Log -Text "All mappings failed - auto-remap disabled. Exiting." -IsError
        break
    }

    Write-Log -Text "Auto-remap enabled ($autoRemapMethod). Monitoring drive health..."
    if ($script:traySync) { $script:traySync.Text = "OneDriveMapper v$version - Monitoring $($successfulMappings.Count) drive(s)" }
    $script:errorsForUser = ''  # Reset for next cycle

    :escape while ($true) {
        foreach ($mapping in $intendedMappings) {
            if (-not $mapping.mapped) { continue }

            $healthy = $false

            # First check: does the drive letter / shortcut / link still exist?
            $linkExists = switch ($mapping.targetLocationType) {
                'networklocation' { Test-Path (Join-Path $mapping.targetLocationPath $mapping.displayName) }
                'driveletter'     { Test-Path $mapping.targetLocationPath }
                'converged'       { Test-Path (Join-Path $mapping.targetLocationPath $mapping.displayName) }
                default           { $true }
            }

            if ($linkExists) {
                # Second check depends on mode
                switch ($autoRemapMethod) {
                    'Path' {
                        # Also verify the underlying WebDAV connection is alive
                        $healthy = Test-Path $mapping.webDavPath
                    }
                    'Link' {
                        # Link exists = healthy (lightweight check)
                        $healthy = $true
                    }
                }
            }
            # else: $healthy stays $false - link/drive is gone

            if (-not $healthy) {
                Write-Log -Text "'$($mapping.displayName)' appears disconnected - checking connectivity..."
                if (-not (Test-Connection 8.8.8.8 -Count 1 -Quiet)) {
                    Write-Log -Text 'No internet - waiting 10s before retry' -IsWarning
                    Start-Sleep -Seconds 10
                    break
                }

                Write-Log -Text 'Internet confirmed - triggering remap'
                if ($script:traySync) {
                    $script:traySync.BalloonTitle = 'OneDriveMapper'
                    $script:traySync.BalloonMsg = "'$($mapping.displayName)' disconnected - remapping..."
                    $script:traySync.BalloonIcon = 'Warning'
                    $script:traySync.ShowBalloon = $true
                }
                $mapping.mapped = $false
                Start-Sleep -Seconds 2
                break escape
            }

            # Check if user requested exit via tray icon
            if ($script:traySync -and $script:traySync.ExitRequested) { break escape }

            $sleep = Get-Random -Minimum 5 -Maximum 20
            Start-Sleep -Seconds $sleep
        }

        # Check for exit between scan cycles
        if ($script:traySync -and $script:traySync.ExitRequested) { break escape }
    }

    # If user clicked Exit, stop the whole script
    if ($script:traySync -and $script:traySync.ExitRequested) { break }

    Write-Log -Text 'Auto-remap triggered - re-authenticating and remapping...'
}

#endregion

if ($script:traySync) {
    $script:traySync.ExitRequested = $true
    Start-Sleep -Milliseconds 500
}
if ($script:trayRunspace) {
    try { $script:trayRunspace.Close() } catch { }
}
Write-Log -Text "===== OneDriveMapper v$version finished ====="
