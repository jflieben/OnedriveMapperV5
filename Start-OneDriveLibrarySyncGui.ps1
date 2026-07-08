#Requires -Version 5.1
<#
.SYNOPSIS
    GUI helper to configure modern OneDrive sync for a SharePoint document library.

.DESCRIPTION
    Uses Microsoft Graph PowerShell with interactive/MFA sign-in to resolve a SharePoint
    site and document library, then writes the documented OneDrive TenantAutoMount
    policy value for the current user. It can also try to launch the OneDrive sync
    protocol so the user completes account/sign-in/MFA in the official OneDrive UI.

    This script does not replace OneDriveMapper.ps1. It is a modern sync-client path
    for Windows 11 and Windows Server 2022 pilots.
#>

[CmdletBinding()]
param(
    [string]$DefaultUserPrincipalName = 'victor.gonzalez@geexsa.com',
    [string]$DefaultLibraryUrl = 'https://gesex.sharepoint.com/sites/Sharepoint_Test_Informatica/archivos/Forms/AllItems.aspx'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$script:LogFile = Join-Path $env:APPDATA 'OneDriveLibrarySyncGui.log'

try {
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
}
catch {
    throw "Windows Forms could not be loaded. Run this script on Windows with Windows PowerShell 5.1."
}

function Write-SyncLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [System.Windows.Forms.TextBox]$StatusBox,
        [switch]$Warning,
        [switch]$Error
    )

    $level = if ($Error) { 'ERROR' } elseif ($Warning) { 'WARN' } else { 'INFO' }
    $line = '{0} | {1} | {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $level, $Message
    try { Add-Content -LiteralPath $script:LogFile -Value $line -ErrorAction SilentlyContinue } catch { }

    if ($StatusBox) {
        $StatusBox.AppendText($line + [Environment]::NewLine)
        $StatusBox.SelectionStart = $StatusBox.TextLength
        $StatusBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Get-ObjectPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }
    $prop = $InputObject.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

function ConvertFrom-QueryString {
    [CmdletBinding()]
    param([string]$Query)

    $result = @{}
    if ([string]::IsNullOrWhiteSpace($Query)) { return $result }

    $trimmed = $Query.TrimStart('?')
    foreach ($pair in $trimmed -split '&') {
        if ([string]::IsNullOrWhiteSpace($pair)) { continue }
        $parts = $pair -split '=', 2
        $key = [Uri]::UnescapeDataString($parts[0].Replace('+', ' '))
        $value = ''
        if ($parts.Count -gt 1) {
            $value = [Uri]::UnescapeDataString($parts[1].Replace('+', ' '))
        }
        $result[$key] = $value
    }
    return $result
}

function Normalize-UrlPath {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '/' }
    $decoded = [Uri]::UnescapeDataString($Path)
    $decoded = $decoded -replace '\\', '/'
    $decoded = $decoded.TrimEnd('/')
    if (-not $decoded.StartsWith('/')) { $decoded = "/$decoded" }
    return $decoded.ToLowerInvariant()
}

function Find-OneDriveExe {
    [CmdletBinding()]
    param()

    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft OneDrive\OneDrive.exe')
    )

    if (${env:ProgramFiles(x86)}) {
        $candidates += (Join-Path ${env:ProgramFiles(x86)} 'Microsoft OneDrive\OneDrive.exe')
    }

    foreach ($candidate in $candidates | Select-Object -Unique) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $null
}

function Get-LocalCompatibilityReport {
    [CmdletBinding()]
    param()

    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $oneDrivePath = Find-OneDriveExe
    $oneDriveVersion = $null
    if ($oneDrivePath) {
        $oneDriveVersion = (Get-Item -LiteralPath $oneDrivePath -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
    }

    [pscustomobject]@{
        WindowsCaption  = if ($os) { $os.Caption } else { [Environment]::OSVersion.VersionString }
        WindowsVersion  = if ($os) { $os.Version } else { [Environment]::OSVersion.Version.ToString() }
        WindowsBuild    = if ($os) { $os.BuildNumber } else { [Environment]::OSVersion.Version.Build }
        ProductType     = if ($os) { $os.ProductType } else { $null }
        PowerShell      = $PSVersionTable.PSVersion.ToString()
        OneDrivePath    = $oneDrivePath
        OneDriveVersion = $oneDriveVersion
    }
}

function Test-GraphPrerequisites {
    [CmdletBinding()]
    param([System.Windows.Forms.TextBox]$StatusBox)

    $required = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Sites'
    )

    foreach ($moduleName in $required) {
        $module = Get-Module -ListAvailable -Name $moduleName |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $module) {
            throw "Missing module $moduleName. Install with: Install-Module $moduleName -Scope CurrentUser"
        }

        Write-SyncLog -Message "Found $moduleName $($module.Version)" -StatusBox $StatusBox
    }

    return $true
}

function Resolve-SharePointSyncTarget {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url)

    try { $uri = [Uri]$Url }
    catch { throw "Invalid SharePoint URL: $Url" }

    if ($uri.Scheme -ne 'https') {
        throw "The SharePoint URL must use https."
    }
    if ($uri.Host -notlike '*.sharepoint.com') {
        throw "The URL host does not look like SharePoint Online: $($uri.Host)"
    }

    $path = [Uri]::UnescapeDataString($uri.AbsolutePath).TrimEnd('/')
    $query = ConvertFrom-QueryString -Query $uri.Query
    $serverRelativeTarget = $path

    if ($query.ContainsKey('id') -and -not [string]::IsNullOrWhiteSpace($query['id'])) {
        $serverRelativeTarget = $query['id'].TrimEnd('/')
    }
    elseif ($serverRelativeTarget -match '(?i)/Forms/[^/]+\.aspx$') {
        $serverRelativeTarget = ($serverRelativeTarget -replace '(?i)/Forms/[^/]+\.aspx$', '').TrimEnd('/')
    }

    $segments = @($path.Trim('/') -split '/' | Where-Object { $_ })
    if ($segments.Count -ge 2 -and ($segments[0] -eq 'sites' -or $segments[0] -eq 'teams')) {
        $siteRelative = "/$($segments[0])/$($segments[1])"
    }
    else {
        $siteRelative = '/'
    }

    $libraryServerRelative = $serverRelativeTarget
    if ($siteRelative -ne '/' -and $serverRelativeTarget.StartsWith($siteRelative + '/', [StringComparison]::OrdinalIgnoreCase)) {
        $tail = $serverRelativeTarget.Substring($siteRelative.Length).Trim('/')
        if ($tail) {
            $librarySegment = @($tail -split '/')[0]
            $libraryServerRelative = "$siteRelative/$librarySegment"
        }
    }
    elseif ($siteRelative -eq '/') {
        $tail = $serverRelativeTarget.Trim('/')
        if ($tail) {
            $libraryServerRelative = '/' + (@($tail -split '/')[0])
        }
    }

    [pscustomobject]@{
        OriginalUrl               = $Url
        SharePointHost           = $uri.Host
        SiteRelativePath         = $siteRelative
        ServerRelativeTargetPath = $serverRelativeTarget
        LibraryServerRelativeUrl = $libraryServerRelative
        LibraryNameHint          = (($libraryServerRelative.Trim('/') -split '/') | Select-Object -Last 1)
    }
}

function Connect-GraphInteractive {
    [CmdletBinding()]
    param([System.Windows.Forms.TextBox]$StatusBox)

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Sites -ErrorAction Stop

    $existing = Get-MgContext -ErrorAction SilentlyContinue
    if ($existing -and $existing.Account) {
        Write-SyncLog -Message "Graph already connected as $($existing.Account)" -StatusBox $StatusBox
        return $existing
    }

    Write-SyncLog -Message 'Opening Microsoft Graph sign-in. Complete account and MFA in the browser window.' -StatusBox $StatusBox
    Connect-MgGraph -Scopes @('Sites.Read.All', 'User.Read') -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null

    $context = Get-MgContext -ErrorAction Stop
    if (-not $context -or -not $context.Account) {
        throw 'Microsoft Graph sign-in completed but no account context was returned.'
    }

    Write-SyncLog -Message "Graph connected as $($context.Account), tenant $($context.TenantId)" -StatusBox $StatusBox
    return $context
}

function Invoke-GraphGetAll {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri

    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -ErrorAction Stop
        $values = Get-ObjectPropertyValue -InputObject $page -Name 'value'
        if ($values) {
            foreach ($item in @($values)) { $items.Add($item) }
        }
        else {
            $items.Add($page)
        }
        $next = Get-ObjectPropertyValue -InputObject $page -Name '@odata.nextLink'
    }

    return $items.ToArray()
}

function Get-GraphSiteByPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Target,
        [System.Windows.Forms.TextBox]$StatusBox
    )

    if ($Target.SiteRelativePath -eq '/') {
        $siteUri = "https://graph.microsoft.com/v1.0/sites/$($Target.SharePointHost)"
    }
    else {
        $siteUri = "https://graph.microsoft.com/v1.0/sites/$($Target.SharePointHost):$($Target.SiteRelativePath)"
    }

    Write-SyncLog -Message "Resolving site through Graph: $siteUri" -StatusBox $StatusBox
    $site = Invoke-MgGraphRequest -Method GET -Uri $siteUri -OutputType PSObject -ErrorAction Stop
    $siteId = Get-ObjectPropertyValue -InputObject $site -Name 'id'
    if (-not $siteId) { throw 'Graph did not return a site id.' }

    Write-SyncLog -Message "Resolved site id: $siteId" -StatusBox $StatusBox
    return $site
}

function Get-GraphDocumentLibraries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Site,
        [System.Windows.Forms.TextBox]$StatusBox
    )

    $siteId = [Uri]::EscapeDataString((Get-ObjectPropertyValue -InputObject $Site -Name 'id'))
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists?`$select=id,displayName,name,webUrl,list,system"
    Write-SyncLog -Message 'Reading document libraries through Graph.' -StatusBox $StatusBox

    $allLists = @(Invoke-GraphGetAll -Uri $uri)
    $libraries = @()

    foreach ($list in $allLists) {
        $listFacet = Get-ObjectPropertyValue -InputObject $list -Name 'list'
        $template = Get-ObjectPropertyValue -InputObject $listFacet -Name 'template'
        $systemFacet = Get-ObjectPropertyValue -InputObject $list -Name 'system'
        if ($template -eq 'documentLibrary' -and -not $systemFacet) {
            $libraries += $list
        }
    }

    Write-SyncLog -Message "Found $($libraries.Count) visible document libraries." -StatusBox $StatusBox
    return $libraries
}

function Select-DocumentLibrary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Libraries,
        [Parameter(Mandatory)][string]$Hint
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Select SharePoint library'
    $form.Size = New-Object System.Drawing.Size(720, 180)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Choose the document library to sync. Hint: $Hint"
    $label.AutoSize = $false
    $label.Location = New-Object System.Drawing.Point(12, 12)
    $label.Size = New-Object System.Drawing.Size(680, 24)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = 'DropDownList'
    $combo.Location = New-Object System.Drawing.Point(12, 45)
    $combo.Size = New-Object System.Drawing.Size(680, 24)

    foreach ($library in $Libraries) {
        $displayName = Get-ObjectPropertyValue -InputObject $library -Name 'displayName'
        $webUrl = Get-ObjectPropertyValue -InputObject $library -Name 'webUrl'
        [void]$combo.Items.Add(('{0} - {1}' -f $displayName, $webUrl))
    }
    if ($combo.Items.Count -gt 0) { $combo.SelectedIndex = 0 }

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'
    $ok.Location = New-Object System.Drawing.Point(516, 92)
    $ok.Size = New-Object System.Drawing.Size(82, 30)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(610, 92)
    $cancel.Size = New-Object System.Drawing.Size(82, 30)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.Controls.AddRange(@($label, $combo, $ok, $cancel))
    $form.AcceptButton = $ok
    $form.CancelButton = $cancel

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $Libraries[$combo.SelectedIndex]
}

function Resolve-DocumentLibrary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Libraries,
        [Parameter(Mandatory)][pscustomobject]$Target,
        [System.Windows.Forms.TextBox]$StatusBox
    )

    if ($Libraries.Count -eq 0) { throw 'No visible document libraries were found in this site.' }

    $targetPath = Normalize-UrlPath -Path $Target.LibraryServerRelativeUrl
    $matches = @()

    foreach ($library in $Libraries) {
        $webUrl = Get-ObjectPropertyValue -InputObject $library -Name 'webUrl'
        if (-not $webUrl) { continue }
        try {
            $libraryPath = Normalize-UrlPath -Path ([Uri]$webUrl).AbsolutePath
            if ($libraryPath -eq $targetPath) { $matches += $library }
        }
        catch { }
    }

    if ($matches.Count -eq 1) {
        Write-SyncLog -Message "Matched library by URL: $((Get-ObjectPropertyValue $matches[0] 'displayName'))" -StatusBox $StatusBox
        return $matches[0]
    }

    $hintMatches = @()
    foreach ($library in $Libraries) {
        $displayName = Get-ObjectPropertyValue -InputObject $library -Name 'displayName'
        $name = Get-ObjectPropertyValue -InputObject $library -Name 'name'
        if ($displayName -eq $Target.LibraryNameHint -or $name -eq $Target.LibraryNameHint) {
            $hintMatches += $library
        }
    }

    if ($hintMatches.Count -eq 1) {
        Write-SyncLog -Message "Matched library by name: $((Get-ObjectPropertyValue $hintMatches[0] 'displayName'))" -StatusBox $StatusBox
        return $hintMatches[0]
    }

    Write-SyncLog -Message 'Could not match exactly. Asking user to choose a library.' -StatusBox $StatusBox -Warning
    return Select-DocumentLibrary -Libraries $Libraries -Hint $Target.LibraryNameHint
}

function Format-OneDriveGuid {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value)

    $trimmed = $Value.Trim().Trim('{', '}')
    return "{$trimmed}"
}

function New-OneDriveLibraryId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Site,
        [Parameter(Mandatory)][pscustomobject]$Library,
        [Parameter(Mandatory)][string]$TenantId
    )

    $siteCompositeId = Get-ObjectPropertyValue -InputObject $Site -Name 'id'
    $siteParts = @($siteCompositeId -split ',', 3)
    if ($siteParts.Count -lt 3) {
        throw "Unexpected Graph site id format: $siteCompositeId"
    }

    $siteCollectionId = Format-OneDriveGuid -Value $siteParts[1]
    $webId = Format-OneDriveGuid -Value $siteParts[2]
    $listId = Format-OneDriveGuid -Value (Get-ObjectPropertyValue -InputObject $Library -Name 'id')
    $siteWebUrl = Get-ObjectPropertyValue -InputObject $Site -Name 'webUrl'

    return 'tenantId={0}&siteId={1}&webId={2}&listId={3}&webUrl={4}&version=1' -f `
        (Format-OneDriveGuid -Value $TenantId), $siteCollectionId, $webId, $listId, $siteWebUrl
}

function Set-OneDriveTenantAutoMount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LibraryName,
        [Parameter(Mandatory)][string]$LibraryId,
        [System.Windows.Forms.TextBox]$StatusBox
    )

    $regPath = 'HKCU:\Software\Policies\Microsoft\OneDrive\TenantAutoMount'
    $null = New-Item -Path $regPath -Force -ErrorAction Stop
    $existing = Get-ItemProperty -Path $regPath -Name $LibraryName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-SyncLog -Message "Replacing existing TenantAutoMount value for '$LibraryName'." -StatusBox $StatusBox -Warning
    }

    $null = New-ItemProperty -Path $regPath -Name $LibraryName -Value $LibraryId -PropertyType String -Force -ErrorAction Stop
    Write-SyncLog -Message "TenantAutoMount configured: HKCU\Software\Policies\Microsoft\OneDrive\TenantAutoMount\$LibraryName" -StatusBox $StatusBox
}

function New-OdOpenSyncUri {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LibraryId,
        [string]$UserEmail,
        [string]$LibraryTitle
    )

    $parameters = ConvertFrom-QueryString -Query $LibraryId
    if ($UserEmail) { $parameters['userEmail'] = $UserEmail }
    if ($LibraryTitle) { $parameters['listTitle'] = $LibraryTitle }

    $pairs = foreach ($key in $parameters.Keys) {
        '{0}={1}' -f [Uri]::EscapeDataString($key), [Uri]::EscapeDataString([string]$parameters[$key])
    }

    return 'odopen://sync/?' + ($pairs -join '&')
}

function Start-OneDriveClient {
    [CmdletBinding()]
    param([System.Windows.Forms.TextBox]$StatusBox)

    $oneDrivePath = Find-OneDriveExe
    if (-not $oneDrivePath) {
        throw 'OneDrive.exe was not found. Install the current OneDrive sync app first.'
    }

    Write-SyncLog -Message "Starting OneDrive: $oneDrivePath /background" -StatusBox $StatusBox
    Start-Process -FilePath $oneDrivePath -ArgumentList '/background' -ErrorAction Stop
    return $oneDrivePath
}

function Invoke-SyncConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][string]$LibraryUrl,
        [Parameter(Mandatory)][bool]$WriteAutoMount,
        [Parameter(Mandatory)][bool]$LaunchOneDriveSync,
        [Parameter(Mandatory)][bool]$OpenLibraryInBrowser,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$StatusBox
    )

    Write-SyncLog -Message 'Starting modern OneDrive sync configuration.' -StatusBox $StatusBox

    $compat = Get-LocalCompatibilityReport
    Write-SyncLog -Message "OS: $($compat.WindowsCaption) build $($compat.WindowsBuild), PowerShell $($compat.PowerShell)" -StatusBox $StatusBox
    if ($compat.OneDrivePath) {
        Write-SyncLog -Message "OneDrive: $($compat.OneDrivePath) version $($compat.OneDriveVersion)" -StatusBox $StatusBox
    }
    else {
        Write-SyncLog -Message 'OneDrive.exe was not found.' -StatusBox $StatusBox -Error
        throw 'OneDrive.exe was not found.'
    }

    if ($compat.ProductType -and [int]$compat.ProductType -ne 1) {
        Write-SyncLog -Message 'Windows Server detected. For RDS/VDI, Microsoft supports OneDrive with per-machine install and FSLogix/profile guidance.' -StatusBox $StatusBox -Warning
    }

    Test-GraphPrerequisites -StatusBox $StatusBox | Out-Null
    $target = Resolve-SharePointSyncTarget -Url $LibraryUrl
    Write-SyncLog -Message "Target site path: $($target.SiteRelativePath); library path: $($target.LibraryServerRelativeUrl)" -StatusBox $StatusBox

    $context = Connect-GraphInteractive -StatusBox $StatusBox
    if ($UserPrincipalName -and $context.Account -and $context.Account -ne $UserPrincipalName) {
        Write-SyncLog -Message "Signed-in Graph account ($($context.Account)) differs from requested sync account ($UserPrincipalName)." -StatusBox $StatusBox -Warning
    }

    $site = Get-GraphSiteByPath -Target $target -StatusBox $StatusBox
    $libraries = @(Get-GraphDocumentLibraries -Site $site -StatusBox $StatusBox)
    $library = Resolve-DocumentLibrary -Libraries $libraries -Target $target -StatusBox $StatusBox
    if (-not $library) { throw 'No document library was selected.' }

    $libraryName = Get-ObjectPropertyValue -InputObject $library -Name 'displayName'
    $libraryWebUrl = Get-ObjectPropertyValue -InputObject $library -Name 'webUrl'
    $libraryId = New-OneDriveLibraryId -Site $site -Library $library -TenantId $context.TenantId

    Write-SyncLog -Message "Library selected: $libraryName" -StatusBox $StatusBox
    Write-SyncLog -Message "Library URL: $libraryWebUrl" -StatusBox $StatusBox
    Write-SyncLog -Message "Library ID: $libraryId" -StatusBox $StatusBox

    if ($WriteAutoMount) {
        Set-OneDriveTenantAutoMount -LibraryName $libraryName -LibraryId $libraryId -StatusBox $StatusBox
    }

    Start-OneDriveClient -StatusBox $StatusBox | Out-Null

    if ($LaunchOneDriveSync) {
        $syncUri = New-OdOpenSyncUri -LibraryId $libraryId -UserEmail $UserPrincipalName -LibraryTitle $libraryName
        Write-SyncLog -Message 'Launching OneDrive sync dialog. Complete any account or MFA prompts in the OneDrive/Microsoft UI.' -StatusBox $StatusBox
        Start-Process -FilePath $syncUri -ErrorAction Stop
    }

    if ($OpenLibraryInBrowser) {
        Write-SyncLog -Message 'Opening library in browser as a fallback/manual verification path.' -StatusBox $StatusBox
        Start-Process -FilePath $libraryWebUrl -ErrorAction SilentlyContinue
    }

    Write-SyncLog -Message 'Configuration completed. If TenantAutoMount was written, OneDrive may apply it the next time the user signs in and within the Microsoft documented sync window.' -StatusBox $StatusBox
}

function Show-OneDriveSyncGui {
    [CmdletBinding()]
    param()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'OneDrive SharePoint Sync - Modern GUI'
    $form.Size = New-Object System.Drawing.Size(820, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(760, 520)
    $form.StartPosition = 'CenterScreen'

    $accountLabel = New-Object System.Windows.Forms.Label
    $accountLabel.Text = 'Cuenta a sincronizar'
    $accountLabel.Location = New-Object System.Drawing.Point(16, 18)
    $accountLabel.Size = New-Object System.Drawing.Size(180, 22)

    $accountBox = New-Object System.Windows.Forms.TextBox
    $accountBox.Location = New-Object System.Drawing.Point(200, 16)
    $accountBox.Size = New-Object System.Drawing.Size(580, 24)
    $accountBox.Text = $DefaultUserPrincipalName

    $urlLabel = New-Object System.Windows.Forms.Label
    $urlLabel.Text = 'URL biblioteca/carpeta'
    $urlLabel.Location = New-Object System.Drawing.Point(16, 55)
    $urlLabel.Size = New-Object System.Drawing.Size(180, 22)

    $urlBox = New-Object System.Windows.Forms.TextBox
    $urlBox.Location = New-Object System.Drawing.Point(200, 52)
    $urlBox.Size = New-Object System.Drawing.Size(580, 24)
    $urlBox.Text = $DefaultLibraryUrl

    $autoMountCheck = New-Object System.Windows.Forms.CheckBox
    $autoMountCheck.Location = New-Object System.Drawing.Point(200, 88)
    $autoMountCheck.Size = New-Object System.Drawing.Size(580, 22)
    $autoMountCheck.Text = 'Configurar TenantAutoMount documentado por Microsoft'
    $autoMountCheck.Checked = $true

    $launchCheck = New-Object System.Windows.Forms.CheckBox
    $launchCheck.Location = New-Object System.Drawing.Point(200, 116)
    $launchCheck.Size = New-Object System.Drawing.Size(580, 22)
    $launchCheck.Text = 'Abrir dialogo de sincronizacion de OneDrive ahora'
    $launchCheck.Checked = $true

    $browserCheck = New-Object System.Windows.Forms.CheckBox
    $browserCheck.Location = New-Object System.Drawing.Point(200, 144)
    $browserCheck.Size = New-Object System.Drawing.Size(580, 22)
    $browserCheck.Text = 'Abrir la biblioteca en el navegador al terminar'
    $browserCheck.Checked = $true

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = 'Resolver y sincronizar'
    $runButton.Location = New-Object System.Drawing.Point(200, 180)
    $runButton.Size = New-Object System.Drawing.Size(170, 34)

    $checkButton = New-Object System.Windows.Forms.Button
    $checkButton.Text = 'Verificar requisitos'
    $checkButton.Location = New-Object System.Drawing.Point(382, 180)
    $checkButton.Size = New-Object System.Drawing.Size(150, 34)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = 'Cerrar'
    $closeButton.Location = New-Object System.Drawing.Point(642, 180)
    $closeButton.Size = New-Object System.Drawing.Size(138, 34)
    $closeButton.Add_Click({ $form.Close() })

    $statusBox = New-Object System.Windows.Forms.TextBox
    $statusBox.Location = New-Object System.Drawing.Point(16, 230)
    $statusBox.Size = New-Object System.Drawing.Size(764, 270)
    $statusBox.Multiline = $true
    $statusBox.ScrollBars = 'Vertical'
    $statusBox.ReadOnly = $true
    $statusBox.Font = New-Object System.Drawing.Font('Consolas', 9)

    $checkButton.Add_Click({
        try {
            Write-SyncLog -Message 'Checking local prerequisites.' -StatusBox $statusBox
            $compat = Get-LocalCompatibilityReport
            Write-SyncLog -Message "OS: $($compat.WindowsCaption) build $($compat.WindowsBuild), PowerShell $($compat.PowerShell)" -StatusBox $statusBox
            if ($compat.OneDrivePath) {
                Write-SyncLog -Message "OneDrive found: $($compat.OneDrivePath) version $($compat.OneDriveVersion)" -StatusBox $statusBox
            }
            else {
                Write-SyncLog -Message 'OneDrive.exe not found.' -StatusBox $statusBox -Error
            }
            Test-GraphPrerequisites -StatusBox $statusBox | Out-Null
            Write-SyncLog -Message 'Prerequisite check completed.' -StatusBox $statusBox
        }
        catch {
            Write-SyncLog -Message $_.Exception.Message -StatusBox $statusBox -Error
            [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Prerequisite error', 'OK', 'Error')
        }
    })

    $runButton.Add_Click({
        $runButton.Enabled = $false
        $checkButton.Enabled = $false
        try {
            Invoke-SyncConfiguration `
                -UserPrincipalName $accountBox.Text.Trim() `
                -LibraryUrl $urlBox.Text.Trim() `
                -WriteAutoMount ([bool]$autoMountCheck.Checked) `
                -LaunchOneDriveSync ([bool]$launchCheck.Checked) `
                -OpenLibraryInBrowser ([bool]$browserCheck.Checked) `
                -StatusBox $statusBox
            [void][System.Windows.Forms.MessageBox]::Show('Proceso completado. Revisa OneDrive y el log para confirmar el estado.', 'OneDrive sync', 'OK', 'Information')
        }
        catch {
            Write-SyncLog -Message $_.Exception.Message -StatusBox $statusBox -Error
            [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'OneDrive sync error', 'OK', 'Error')
        }
        finally {
            $runButton.Enabled = $true
            $checkButton.Enabled = $true
        }
    })

    $form.Controls.AddRange(@(
        $accountLabel, $accountBox,
        $urlLabel, $urlBox,
        $autoMountCheck, $launchCheck, $browserCheck,
        $runButton, $checkButton, $closeButton,
        $statusBox
    ))

    Write-SyncLog -Message "Log file: $script:LogFile" -StatusBox $statusBox
    [void]$form.ShowDialog()
}

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    [void][System.Windows.Forms.MessageBox]::Show(
        'Para una experiencia GUI estable, ejecuta con Windows PowerShell en STA: powershell.exe -STA -ExecutionPolicy Bypass -File .\Start-OneDriveLibrarySyncGui.ps1',
        'OneDrive sync GUI',
        'OK',
        'Warning'
    )
}

Show-OneDriveSyncGui
