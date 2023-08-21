######## 
#OneDriveMapper
#Copyright:          Free to use, please leave this header intact 
#Author:             Jos Lieben
#Company:            Lieben Consultancy (http://www.lieben.nu) 
#Script help:        http://www.lieben.nu, please provide a decrypted Fiddler Trace Log if you're using Native auth and having issues
#Copyright/License:  https://www.lieben.nu/liebensraum/commercial-use/ (Commercial (re)use not allowed without prior written consent by the author, otherwise free to use/modify as long as header are kept intact)
#Purpose:            This script maps Onedrive for Business and/or maps a configurable number of Sharepoint Libraries
#Enterprise users:   This script is not recommended for Enterprise use as no dedicated support is available. Check www.lieben.nu for enterprise options.

param(
    [Switch]$asTask,
    [Switch]$hideConsole
)

$version = "3.29"

####MANDATORY MANUAL CONFIGURATION
$O365CustomerName      = "lieben"                  #This should be the name of your tenant (example, ogd as in ogd.onmicrosoft.com) 
$userLookupMode        = 4                         #1 = Active Directory UPN, 2 = Active Directory Email, 3 = Azure AD Joined Windows 10, 4 = query user for his/her login, 5 = lookup by registry key, 6 = display full form (ask for both username and login if no cached versions can be found), 7 = whoami /upn
$adfsMode              = 1                         #1 = use whatever came out of userLookupMode, 2 = use only the part before the @ in the upn, 3 = use user certificate (local user store) and match Subject to Username
$showConsoleOutput     = $True                     #Set this to $False to hide console output
$showElevatedConsole   = $True

<#if you wish to add more, add more lines to the below (copy the first above itself). Parameter explanation:
displayName = the label of the driveletter, or name of the shortcut we'll create to the target site/library
targetLocationType = driveletter OR networklocation, if you use driveletter, enter a driveletter in targetLocationPath. If you use networklocation, enter a path to a folder where you want the shortcut to be created
targetLocationPath = enter a driveletter if mapping to a driveletter, enter a folder path if just creating shortcuts, type 'autodetect' if you want the script to automatically find a free driveletter
sourceLocationPath = autodetect or the full URL to the sharepoint / groups site. Autodetect automatically makes this a mapping to Onedrive For Business
mapOnlyForSpecificGroup = this only works for DOMAIN JOINED devices that can reach a domain controller and means that the mapping will only be made if the user is a member of the group you specify here
#>

#DEFAULT SETTINGS: (onedrive only, to the X: drive)
$desiredMappings =  @(
    @{"displayName"="Onedrive for Business";"targetLocationType"="driveletter";"targetLocationPath"="X:";"sourceLocationPath"="autodetect";"mapOnlyForSpecificGroup"=""}
)

#EXAMPLE SETTINGS (Onedrive for Business, two Sharepoint sites, one mapped to a driveletter, one to a shortcut, the last only when a member of the Active Directory group SEC-SHAREPOINTA)
#$desiredMappings =  @(
#    @{"displayName"="Onedrive for Business";"targetLocationType"="driveletter";"targetLocationPath"="X:";"sourceLocationPath"="autodetect";"mapOnlyForSpecificGroup"=""},
#    @{"displayName"="Sharepoint Site A";"targetLocationType"="networklocation";"targetLocationPath"="$env:APPDATA\Microsoft\Windows\Network Shortcuts";"sourceLocationPath"="https://ogd.sharepoint.com/sites/OGDWerkplek/Gedeelde%20%20documenten/Forms/AllItems.aspx";"mapOnlyForSpecificGroup"="SEC-SHAREPOINTA"},
#    @{"displayName"="Sharepoint Site A";"targetLocationType"="driveletter";"targetLocationPath"="Z:";"sourceLocationPath"="https://ogd.sharepoint.com/sites/OGDWerkplek/Gedeelde%20%20documenten/Forms/AllItems.aspx";"mapOnlyForSpecificGroup"=""} #note that the last entry does NOT end with a comma!
#)

$redirectFolders       = $false #Set to TRUE and configure below hashtable to redirect folders
$listOfFoldersToRedirect = @(#One line for each folder you want to redirect, only works if redirectFolders=$True. For knownFolderInternalName choose from Get-KnownFolderPath function, for knownFolderInternalIdentifier choose from Set-KnownFolderPath function
    @{"knownFolderInternalName" = "Desktop";"knownFolderInternalIdentifier"="Desktop";"desiredTargetPath"="X:\Desktop";"copyExistingFiles"="true"},
    @{"knownFolderInternalName" = "MyDocuments";"knownFolderInternalIdentifier"="Documents";"desiredTargetPath"="X:\My Documents";"copyExistingFiles"="true"},
    @{"knownFolderInternalName" = "MyPictures";"knownFolderInternalIdentifier"="Pictures";"desiredTargetPath"="X:\My Pictures";"copyExistingFiles"="false"} #note that the last entry does NOT end with a comma
)

###OPTIONAL CONFIGURATION
$autoMapFavoriteSites  = $False                     #Set to $True to automatically map any sites/teams/groups the user has favorited (https://yourtenantname.sharepoint.com/_layouts/15/sharepoint.aspx?v=following)
$autoMapFavoritesMode  = "Converged"                  #Normal = map each detected site to a free driveletter, Onedrive = map to Onedrive subfolder (Links), Converged = single dummy mapping with all links in it
$autoMapFavoritesDrive = "T"                       #Driveletter when using automapFavoritesMode = "Converged"
$autoMapFavoritesLabel = "Teams"                   #Label of favorites container, ie; folder name if automapFavoritesMode = "Onedrive", drive label if automapFavoritesMode = "Converged"
$autoMapFavoritesDrvLetterList = "DEFGHIJKLMNOPQRSTUVWXYZ" #List of driveletters that shall be used (you can ommit some of yours "reserved" letters)
$favoriteSitesDLName   = "Gedeelde  Documenten"    #Normally autodetected, default document library name in Teams/Groups/Sites to map in conjunction with $autoMapFavoriteSites, note the double spaces! Use Shared  Documents for english language tenants
$restartExplorer       = $True                     #You can safely set this to False if you're not redirecting folders
$autoResetIE           = $False                    #always clear all Internet Explorer cookies before running (prevents certain occasional issues with IE)
$authenticateToProxy   = $False                    #use system proxy settings and authenticate automatically
$libraryName           = "Documents"               #leave this default, unless you wish to map a non-default onedrive library you've created 
$adfsSmartLink         = $Null                     #If set, the ADFS smartlink will be used to log in to Office 365. For more info, read the FAQ at http://http://www.lieben.nu/liebensraum/onedrivemapper/onedrivemapper-faq/
$displayErrors         = $True                     #show errors to user in visual popups
$persistentMapping     = $True                     #If set to $False, the mapping will go away when the user logs off
$buttonText            = "Login"                   #Text of the button on the password input popup box
$loginformTitleText    = "OneDriveMapper"          #Used as the window title for input popup boxes (userLookupMode is set to 4) and login forms (userLookupMode is set to 6)
$loginformIntroText    = "Welcome to COMPANY NAME`r`nPlease enter your login and password" #used as introduction text when you set userLookupMode to 6
$loginFieldText        = "Please enter your login in the form of xxx@xxx.com" #used as label above the login text field when you set userLookupMode to 6
$passwordFieldText     = "Please enter your password" #used as label above the password text field when you set userLookupMode to 6
$urlOpenAfter          = ""                        #This URL will be opened by the script after running if you configure it
$showProgressBar       = $True                     #will show a progress bar to the user
$progressBarColor      = "#CC99FF"
$script:progressBarText       = "OnedriveMapper v$version is connecting your drives..."
$autoDetectProxy       = $False                    #if set to $False, unchecks the 'Automatically detect proxy settings' setting in IE; this greatly enhanced WebDav performance, set to true to not modify this IE setting (leave as is)
$forceUserName         = ''                        #if anything is entered here, userLookupMode is ignored
$forcePassword         = ''                        #if anything is entered here, the user won't be prompted for a password. This function is not recommended, as your password could be stolen from this file 
$addShellLink          = $False                    #Adds a link to Onedrive to the Shell under Favorites (Windows 7, 8 / 2008R2 and 2012R2 only) If you use a remote path, google EnableShellShortcutIconRemotePath
$cacheCookies          = $True                     #caches user cookies in appdata to sign in silently while still possible/valid
$cookieCacheFilePath   = ($env:APPDATA + "\OneDriveMapper.v3c")    #Logfile to log to 
$logfile               = ($env:APPDATA + "\OneDriveMapper_$version.log")    #Logfile to log to 
$pwdCache              = ($env:APPDATA + "\OneDriveMapper.tmp")    #file to store encrypted password into, change to $Null to disable
$loginCache            = ($env:APPDATA + "\OneDriveMapper.tmp2")    #file to store encrypted login into, change to $Null to disable

if($hideConsole){
    $showConsoleOutput     = $False
    $showElevatedConsole   = $False
}

if($showConsoleOutput -eq $False){
    $t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    try{
        add-type -name win -member $t -namespace native
        [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
    }catch{$Null}
}

######## 
#Required resources and some customizations you'll probably not use
######## 
$privateSuffix = "-my"
$script:errorsForUser = ""
$userLoginRegistryKey = "HKCU:\System\CurrentControlSet\Control\CustomUID"
$onedriveIconPath = "C:\GitRepos\OnedriveMapper\onedrive.ico" #if this file exists, and you've set addShellLink to True, it will be used as icon for the shortcut
$teamsIconPath = "C:\GitRepos\OnedriveMapper\teams.ico" #if this file exists, and you've set addShellLink to True, it will be used as icon for the shortcut
$sharepointIconPath = "C:\GitRepos\OnedriveMapper\sharepoint.ico" #if this file exists, and you've set addShellLink to True, it will be used as icon for the shortcut
$i_MaxLocalLogSize = 2 #max local log size in MB
$certificateMatchMethod = 1 #used with adfsMode = 3, when set to 1 it'll match based on the local username, if set to 2 it'll use the following variable to match to a template name
$certificateTemplateName  = "Office365_Client_Authentication"
$O365CustomerName = $O365CustomerName.ToLower() 
#for people that don't RTFM, fix wrongly entered customer names:
$O365CustomerName = $O365CustomerName -Replace ".onmicrosoft.com",""
$forceUserName = $forceUserName.ToLower() 
$finalURLs = @("https://portal.office.com","https://outlook.office365.com","https://outlook.office.com","https://$($O365CustomerName)-my.sharepoint.com","https://$($O365CustomerName).sharepoint.com","https://www.office.com")

function log{
    param (
        [Parameter(Mandatory=$true)][String]$text,
        [Switch]$fout,
        [Switch]$warning
    )
    if($fout){
        $text = "ERROR | $text"
    }
    elseif($warning){
        $text = "WARNING | $text"
    }
    else{
        $text = "INFO | $text"
    }
    try{
        Add-Content $logfile "$(Get-Date) | $text"
    }catch{$Null}
    if($showConsoleOutput){
        if($fout){
            Write-Host $text -ForegroundColor Red
        }elseif($warning){
            Write-Host $text -ForegroundColor Yellow
        }else{
            Write-Host $text -ForegroundColor Green
        }
    }
}

function ResetLog{
    <#
    -------------------------------------------------------------------------------------------
    Manage the local log file size
    Always keep a backup
    #credits to Steven Heimbecker
    -------------------------------------------------------------------------------------------
    #>
    #Restart the local log file if it exists and is bigger than $i_MaxLocalLogSize MB as defined below
    [int]$i_LocalLogSize
    if ((Test-Path $logfile) -eq $True){
        #The log file exists
        try{
            $i_LocalLogSize=(Get-Item $logfile).Length
            if($i_LocalLogSize / 1Mb -gt $i_MaxLocalLogSize){
                #The log file is greater than the defined maximum.  Let's back it up / rename it
                #Blank line in the old log
                log -text ""
                log -text "******** End of log - maximum size ********"
                #Save the current log as a .old.  If one already exists, delete it.
                if ((Test-Path ($logfile + ".old")) -eq $True){
                    #Already a backup file, delete it
                    Remove-Item ($logfile + ".old") -Force -Confirm:$False
                }
                #Now lets rename 
                Rename-Item -path $logfile -NewName ($logfile + ".old") -Force -Confirm:$False
                #Start a new log
                log -text "******** Log file reset after reaching maximum size ********`n"
            }
        }catch{
            log -text "there was an issue resetting the logfile! $($Error[0])" -fout
        }
    }
}

$scriptPath = $MyInvocation.MyCommand.Definition
ResetLog
log -text "-----$(Get-Date) OneDriveMapper v$version - $($env:USERNAME) on $($env:COMPUTERNAME) starting-----" 

###THIS ONLY HAS TO BE CONFIGURED IF YOU WANT TO MAP USER SECURITY GROUPS TO SHAREPOINT SITES
if($desiredMappings.mapOnlyForSpecificGroup | Where-Object{$_.Length -gt 0}){
    try{
        $groups = ([ADSISEARCHER]"(member:1.2.840.113556.1.4.1941:=$(([ADSISEARCHER]"samaccountname=$($env:USERNAME)").FindOne().Properties.distinguishedname))").FindAll().Properties.distinguishedname -replace '^CN=([^,]+).+$','$1'
        log -text "cached user group membership because you have configured mappings where the mapOnlyForSpecificGroup option was configured"   
    }catch{
        log -text "failed to cache user group membership, ignoring these mappings because of: $($Error[0])" -fout
        $desiredMappings = $desiredMappings | Where-Object{$_.mapOnlyForSpecificGroup.Length -eq 0}
    }
}

#Find a driveletter for any drivemappings that have autotect as targetlocationpath
$drvlist=(Get-PSDrive -PSProvider filesystem).Name
for($i=0;$i -lt $desiredMappings.Count;$i++){
    if($desiredMappings[$i].targetLocationPath -eq "autodetect"){
        Foreach ($drvletter in $autoMapFavoritesDrvLetterList.ToCharArray()) {
            If ($drvlist -notcontains $drvletter) {
                $drvlist += $drvletter
                $desiredMappings[$i].targetLocationPath = "$($drvletter):"
                log -text "automatically selected drive $drvletter for Onedrive mapping"
                break
            }
        }
    }
}

function Add-NetworkLocation
<#
    Author: Tom White, 2015.
    Description:
        Creates a network location shortcut using the specified path, name and target.
        Replicates the behaviour of the 'Add Network Location' wizard, creating a special folder as opposed to a simple shortcut.
        Returns $true on success and $false on failure.
        Use -Verbose for extended output.
    Example:
        Add-NetworkLocation -networkLocationPath "$env:APPDATA\Microsoft\Windows\Network Shortcuts" -networkLocationName "Network Location" -networkLocationTarget "\\server\share" -Verbose
#>
{
    [CmdLetBinding()]
    param
    (
        [string]$networkLocationPath="$env:APPDATA\Microsoft\Windows\Network Shortcuts",
        [Parameter(Mandatory=$true)][string]$networkLocationName ,
        [Parameter(Mandatory=$true)][string]$networkLocationTarget,
        [String]$iconPath
    )
    Begin
    {
        Write-Verbose -Message "Network location path: `"$networkLocationPath`"."
        Write-Verbose -Message "Network location name: `"$networkLocationName`"."
        Write-Verbose -Message "Network location target: `"$networkLocationTarget`"."
        Set-Variable -Name desktopIniContent -Option ReadOnly -value ([string]"[.ShellClassInfo]`r`nCLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}`r`nFlags=2")
    }
    Process
    {
        Write-Verbose -Message "Checking that `"$networkLocationPath`" is a valid directory..."
        if(Test-Path -Path $networkLocationPath -PathType Container)
        {
            try
            {
                Write-Verbose -Message "Creating `"$networkLocationPath\$networkLocationName`"."
                [void]$(New-Item -Path "$networkLocationPath\$networkLocationName" -ItemType Directory -ErrorAction Stop)
                Write-Verbose -Message "Setting system attribute on `"$networkLocationPath\$networkLocationName`"."
                Set-ItemProperty -Path "$networkLocationPath\$networkLocationName" -Name Attributes -Value ([System.IO.FileAttributes]::System) -ErrorAction Stop
            }
            catch [Exception]
            {
                Write-Error -Message "Cannot create or set attributes on `"$networkLocationPath\$networkLocationName`". Check your access and/or permissions."
                return $false
            }
        }
        else
        {
            Write-Error -Message "`"$networkLocationPath`" is not a valid directory path."
            return $false
        }
        try
        {
            Write-Verbose -Message "Creating `"$networkLocationPath\$networkLocationName\desktop.ini`"."
            [object]$desktopIni = New-Item -Path "$networkLocationPath\$networkLocationName\desktop.ini" -ItemType File
            Write-Verbose -Message "Writing to `"$($desktopIni.FullName)`"."
            Add-Content -Path $desktopIni.FullName -Value $desktopIniContent
        }
        catch [Exception]
        {
            Write-Error -Message "Error while creating or writing to `"$networkLocationPath\$networkLocationName\desktop.ini`". Check your access and/or permissions."
            return $false
        }
        try
        {
            $WshShell = New-Object -ComObject WScript.Shell
            Write-Verbose -Message "Creating shortcut to `"$networkLocationTarget`" at `"$networkLocationPath\$networkLocationName\target.lnk`"."
            $Shortcut = $WshShell.CreateShortcut("$networkLocationPath\$networkLocationName\target.lnk")
            $Shortcut.TargetPath = $networkLocationTarget
            if([System.IO.File]::Exists($iconPath)){
                $Shortcut.IconLocation = "$($iconPath), 0"
            }            
            $Shortcut.Description = "Created $(Get-Date -Format s) by $($MyInvocation.MyCommand)."
            $Shortcut.Save()
        }
        catch [Exception]
        {
            Write-Error -Message "Error while creating shortcut @ `"$networkLocationPath\$networkLocationName\target.lnk`". Check your access and permissions."
            return $false
        }
        return $true
    }
}

function handleMFArequest{
    Param(
        $res,
        $clientId
    )
    $mfaArray = returnEnclosedFormValue -res $res -searchString "`"arrUserProofs`":" -endString "}],"
    if($mfaArray -ne -1){
        $mfaArray = "$($mfaArray)}]" | ConvertFrom-Json
        $mfaMethod = @($mfaArray | Where-Object {$_.isDefault})[0].authMethodId
    }else{
        $mfaMethod = returnEnclosedFormValue -res $res -searchString "`"authMethodId`":`""
    }

    if($mfaMethod -eq -1){
        Throw "No MFA method detected"
    }
    if($mfaMethod -ne "PhoneAppNotification" -and $mfaMethod -ne "TwoWayVoiceMobile"){
        Throw "Unsupported MFA method detected: $mfaMethod"
    }
    $canary = returnEnclosedFormValue -res $res -searchString "`",`"canary`":`""
    $apiCanary = returnEnclosedFormValue -res $res -searchString "ConvergedTFA`",`"apiCanary`":`""
    $ctx = returnEnclosedFormValue -res $res -searchString "sFTName`":`"flowToken`",`"sCtx`":`""
    $sFT = returnEnclosedFormValue -res $res -searchString "`",`"sFT`":`""
    $body = @{"AuthMethodId"=$mfaMethod;"Method"="BeginAuth";"ctx"=$ctx;"flowToken"=$sFT}
    $customHeaders = @{"canary" = $apiCanary;"hpgrequestid" = $res.Headers["x-ms-request-id"];"client-request-id"=$clientId;"hpgid"=1114;"hpgact"=2000}
    try{
        $res = New-WebRequest -url "https://login.microsoftonline.com/common/SAS/BeginAuth" -Method POST -customHeaders $customHeaders -body ($body | ConvertTo-Json) -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/json"
        $result = $res.Content | convertfrom-json
        $entropy = $result.Entropy
        $sFT = $result.FlowToken
        $ctx = $result.Ctx
        if($result.Success -eq $False){
            Throw
        }
        $sessionId = $result.SessionId
    }catch{
        Throw "SAS BeginAuth failure, MFA initiation not accepted $($result.Message)"
    }

    if($result.Entropy){
        [System.Windows.Forms.MessageBox]::Show("Please enter this number in your MFA Authenticator: " + $entropy, "OnedriveMapper") 
    }

    
    $waitedForMFA = 0
    $p=0
    while($true){
        if($waitedForMFA -ge 60){
            Throw "Waited longer than 60 seconds for MFA request to be validated, aborting"
        }
        try{
            $p++
            $body = @{"Method"="EndAuth";"SessionId"=$sessionId;"FlowToken"=$sFT;"Ctx"=$ctx;"AuthMethodId"=$mfaMethod;"PollCount"=$p}
            $res2 = New-WebRequest -url "https://login.microsoftonline.com/common/SAS/EndAuth" -Method POST -customHeaders $customHeaders -body ($body | ConvertTo-Json) -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/json"
            $result = $res2.Content | convertfrom-json
            if($result.Success){
                break
            }
        }catch{
            Throw "SAS EndAuth failure, MFA initiation not accepted"
        }
        Start-Sleep -s 5
        $waitedForMFA+=5
    }

    try{
        $ctx = [System.Web.HttpUtility]::UrlEncode($ctx)
        $canary = [System.Web.HttpUtility]::UrlEncode($canary)
        $sFT = [System.Web.HttpUtility]::UrlEncode($result.FlowToken)
        if($mfaMethod -eq "PhoneAppNotification"){
            $mfaAuthMethod = "PhoneAppOTP"
            $type=22
        }
        if($mfaMethod -eq "TwoWayVoiceMobile"){
            $mfaAuthMethod = "TwoWayVoiceMobile"
            $type=1
        }
        $body = "type=$type&request=$ctx&mfaAuthMethod=$mfaAuthMethod&canary=$canary&login=$userUPN&flowToken=$sFT&hpgrequestid=$($customHeaders["hpgrequestid"])&sacxt=&i2=&i17=&i18=&i19=7406"
        $res = New-WebRequest -url "https://login.microsoftonline.com/common/SAS/ProcessAuth" -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri
        return $res
    }catch{
        Throw "SAS ProcessAuth failure"
    }
}

function createFavoritesShortcutToO4B{
    Param(
        $targetLocation
    )
    $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    try{
        $linksPath = (Get-ItemProperty -Path $regPath -Name "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}")."{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}"
        log -text "Path to links folder determined: $linksPath"
    }catch{
        Throw "Failed to determine path to Links folder: $($Error[0])"
    }
    $shortcutName = "Onedrive - $O365CustomerName.lnk"
    $shortcutPath = Join-Path $linksPath -ChildPath $shortcutName
    if([System.IO.Directory]::Exists($linksPath)){
        if([System.IO.File]::Exists($shortcutPath)){
            log -text "Shortcut already exists"
            return
        }else{
            try{
                $WshShell = New-Object -ComObject WScript.Shell
                $Shortcut = $WshShell.CreateShortcut($shortcutPath)
                $Shortcut.TargetPath = $targetLocation
                if([System.IO.File]::Exists($onedriveIconPath)){
                    $Shortcut.IconLocation = "$($onedriveIconPath), 0"
                }
                $Shortcut.Description ="Onedrive for Business"
                $Shortcut.Save()
            }catch{
                Throw
            } 
        }
    }else{
        Throw "Links folder does not exist"
    }
}

Function Set-KnownFolderPath {
    Param (
            [Parameter(Mandatory = $true)][ValidateSet('AddNewPrograms', 'AdminTools', 'AppUpdates', 'CDBurning', 'ChangeRemovePrograms', 'CommonAdminTools', 'CommonOEMLinks', 'CommonPrograms', `
            'CommonStartMenu', 'CommonStartup', 'CommonTemplates', 'ComputerFolder', 'ConflictFolder', 'ConnectionsFolder', 'Contacts', 'ControlPanelFolder', 'Cookies', `
            'Desktop', 'Documents', 'Downloads', 'Favorites', 'Fonts', 'Games', 'GameTasks', 'History', 'InternetCache', 'InternetFolder', 'Links', 'LocalAppData', `
            'LocalAppDataLow', 'LocalizedResourcesDir', 'Music', 'NetHood', 'NetworkFolder', 'OriginalImages', 'PhotoAlbums', 'Pictures', 'Playlists', 'PrintersFolder', `
            'PrintHood', 'Profile', 'ProgramData', 'ProgramFiles', 'ProgramFilesX64', 'ProgramFilesX86', 'ProgramFilesCommon', 'ProgramFilesCommonX64', 'ProgramFilesCommonX86', `
            'Programs', 'Public', 'PublicDesktop', 'PublicDocuments', 'PublicDownloads', 'PublicGameTasks', 'PublicMusic', 'PublicPictures', 'PublicVideos', 'QuickLaunch', `
            'Recent', 'RecycleBinFolder', 'ResourceDir', 'RoamingAppData', 'SampleMusic', 'SamplePictures', 'SamplePlaylists', 'SampleVideos', 'SavedGames', 'SavedSearches', `
            'SEARCH_CSC', 'SEARCH_MAPI', 'SearchHome', 'SendTo', 'SidebarDefaultParts', 'SidebarParts', 'StartMenu', 'Startup', 'SyncManagerFolder', 'SyncResultsFolder', `
            'SyncSetupFolder', 'System', 'SystemX86', 'Templates', 'TreeProperties', 'UserProfiles', 'UsersFiles', 'Videos', 'Windows')]
            [string]$KnownFolder,
            [Parameter(Mandatory = $true)][string]$Path
    )

    # Define known folder GUIDs
    $KnownFolders = @{
        'AddNewPrograms' = 'de61d971-5ebc-4f02-a3a9-6c82895e5c04';'AdminTools' = '724EF170-A42D-4FEF-9F26-B60E846FBA4F';'AppUpdates' = 'a305ce99-f527-492b-8b1a-7e76fa98d6e4';
        'CDBurning' = '9E52AB10-F80D-49DF-ACB8-4330F5687855';'ChangeRemovePrograms' = 'df7266ac-9274-4867-8d55-3bd661de872d';'CommonAdminTools' = 'D0384E7D-BAC3-4797-8F14-CBA229B392B5';
        'CommonOEMLinks' = 'C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D';'CommonPrograms' = '0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8';'CommonStartMenu' = 'A4115719-D62E-491D-AA7C-E74B8BE3B067';
        'CommonStartup' = '82A5EA35-D9CD-47C5-9629-E15D2F714E6E';'CommonTemplates' = 'B94237E7-57AC-4347-9151-B08C6C32D1F7';'ComputerFolder' = '0AC0837C-BBF8-452A-850D-79D08E667CA7';
        'ConflictFolder' = '4bfefb45-347d-4006-a5be-ac0cb0567192';'ConnectionsFolder' = '6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD';'Contacts' = '56784854-C6CB-462b-8169-88E350ACB882';
        'ControlPanelFolder' = '82A74AEB-AEB4-465C-A014-D097EE346D63';'Cookies' = '2B0F765D-C0E9-4171-908E-08A611B84FF6';'Desktop' = @('B4BFCC3A-DB2C-424C-B029-7FE99A87C641');
        'Documents' = @('FDD39AD0-238F-46AF-ADB4-6C85480369C7','f42ee2d3-909f-4907-8871-4c22fc0bf756');'Downloads' = @('374DE290-123F-4565-9164-39C4925E467B','7d83ee9b-2244-4e70-b1f5-5393042af1e4');
        'Favorites' = '1777F761-68AD-4D8A-87BD-30B759FA33DD';'Fonts' = 'FD228CB7-AE11-4AE3-864C-16F3910AB8FE';'Games' = 'CAC52C1A-B53D-4edc-92D7-6B2E8AC19434';
        'GameTasks' = '054FAE61-4DD8-4787-80B6-090220C4B700';'History' = 'D9DC8A3B-B784-432E-A781-5A1130A75963';'InternetCache' = '352481E8-33BE-4251-BA85-6007CAEDCF9D';
        'InternetFolder' = '4D9F7874-4E0C-4904-967B-40B0D20C3E4B';'Links' = 'bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968';'LocalAppData' = 'F1B32785-6FBA-4FCF-9D55-7B8E7F157091';
        'LocalAppDataLow' = 'A520A1A4-1780-4FF6-BD18-167343C5AF16';'LocalizedResourcesDir' = '2A00375E-224C-49DE-B8D1-440DF7EF3DDC';'Music' = @('4BD8D571-6D19-48D3-BE97-422220080E43','a0c69a99-21c8-4671-8703-7934162fcf1d');
        'NetHood' = 'C5ABBF53-E17F-4121-8900-86626FC2C973';'NetworkFolder' = 'D20BEEC4-5CA8-4905-AE3B-BF251EA09B53';'OriginalImages' = '2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39';
        'PhotoAlbums' = '69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C';'Pictures' = @('33E28130-4E1E-4676-835A-98395C3BC3BB','0ddd015d-b06c-45d5-8c4c-f59713854639');
        'Playlists' = 'DE92C1C7-837F-4F69-A3BB-86E631204A23';'PrintersFolder' = '76FC4E2D-D6AD-4519-A663-37BD56068185';'PrintHood' = '9274BD8D-CFD1-41C3-B35E-B13F55A758F4';
        'Profile' = '5E6C858F-0E22-4760-9AFE-EA3317B67173';'ProgramData' = '62AB5D82-FDC1-4DC3-A9DD-070D1D495D97';'ProgramFiles' = '905e63b6-c1bf-494e-b29c-65b732d3d21a';
        'ProgramFilesX64' = '6D809377-6AF0-444b-8957-A3773F02200E';'ProgramFilesX86' = '7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E';'ProgramFilesCommon' = 'F7F1ED05-9F6D-47A2-AAAE-29D317C6F066';
        'ProgramFilesCommonX64' = '6365D5A7-0F0D-45E5-87F6-0DA56B6A4F7D';'ProgramFilesCommonX86' = 'DE974D24-D9C6-4D3E-BF91-F4455120B917';'Programs' = 'A77F5D77-2E2B-44C3-A6A2-ABA601054A51';
        'Public' = 'DFDF76A2-C82A-4D63-906A-5644AC457385';'PublicDesktop' = 'C4AA340D-F20F-4863-AFEF-F87EF2E6BA25';'PublicDocuments' = 'ED4824AF-DCE4-45A8-81E2-FC7965083634';
        'PublicDownloads' = '3D644C9B-1FB8-4f30-9B45-F670235F79C0';'PublicGameTasks' = 'DEBF2536-E1A8-4c59-B6A2-414586476AEA';'PublicMusic' = '3214FAB5-9757-4298-BB61-92A9DEAA44FF';
        'PublicPictures' = 'B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5';'PublicVideos' = '2400183A-6185-49FB-A2D8-4A392A602BA3';'QuickLaunch' = '52a4f021-7b75-48a9-9f6b-4b87a210bc8f';
        'Recent' = 'AE50C081-EBD2-438A-8655-8A092E34987A';'RecycleBinFolder' = 'B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC';'ResourceDir' = '8AD10C31-2ADB-4296-A8F7-E4701232C972';
        'RoamingAppData' = '3EB685DB-65F9-4CF6-A03A-E3EF65729F3D';'SampleMusic' = 'B250C668-F57D-4EE1-A63C-290EE7D1AA1F';'SamplePictures' = 'C4900540-2379-4C75-844B-64E6FAF8716B';
        'SamplePlaylists' = '15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5';'SampleVideos' = '859EAD94-2E85-48AD-A71A-0969CB56A6CD';'SavedGames' = '4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4';
        'SavedSearches' = '7d1d3a04-debb-4115-95cf-2f29da2920da';'SEARCH_CSC' = 'ee32e446-31ca-4aba-814f-a5ebd2fd6d5e';'SEARCH_MAPI' = '98ec0e18-2098-4d44-8644-66979315a281';
        'SearchHome' = '190337d1-b8ca-4121-a639-6d472d16972a';'SendTo' = '8983036C-27C0-404B-8F08-102D10DCFD74';'SidebarDefaultParts' = '7B396E54-9EC5-4300-BE0A-2482EBAE1A26';
        'SidebarParts' = 'A75D362E-50FC-4fb7-AC2C-A8BEAA314493';'StartMenu' = '625B53C3-AB48-4EC1-BA1F-A1EF4146FC19';'Startup' = 'B97D20BB-F46A-4C97-BA10-5E3608430854';
        'SyncManagerFolder' = '43668BF8-C14E-49B2-97C9-747784D784B7';'SyncResultsFolder' = '289a9a43-be44-4057-a41b-587a76d7e7f9';'SyncSetupFolder' = '0F214138-B1D3-4a90-BBA9-27CBC0C5389A';
        'System' = '1AC14E77-02E7-4E5D-B744-2EB1AE5198B7';'SystemX86' = 'D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27';'Templates' = 'A63293E8-664E-48DB-A079-DF759E0509F7';
        'TreeProperties' = '5b3749ad-b49f-49c1-83eb-15370fbd4882';'UserProfiles' = '0762D272-C50A-4BB0-A382-697DCD729B80';'UsersFiles' = 'f3ce0f7c-4901-4acc-8648-d5d44b04ef8f';
        'Videos' = @('18989B1D-99B5-455B-841C-AB7C74E4DDFC','35286a68-3c57-41a1-bbb1-0eae73d76c95');'Windows' = 'F38BF404-1D43-42F2-9305-67DE0B28FC23';
    }

    $Type = ([System.Management.Automation.PSTypeName]'KnownFolders').Type
    If (-not $Type) {
        $Signature = @'
[DllImport("shell32.dll")]
public extern static int SHSetKnownFolderPath(ref Guid folderId, uint flags, IntPtr token, [MarshalAs(UnmanagedType.LPWStr)] string path);
'@
        $Type = Add-Type -MemberDefinition $Signature -Name 'KnownFolders' -Namespace 'SHSetKnownFolderPath' -PassThru
    }

	If (!(Test-Path $Path -PathType Container)) {
		New-Item -Path $Path -Type Directory -Force -Verbose
    }

    If (Test-Path $Path -PathType Container) {
        ForEach ($guid in $KnownFolders[$KnownFolder]) {
            $result = $Type::SHSetKnownFolderPath([ref]$guid, 0, 0, $Path)
            If ($result -ne 0) {
                $errormsg = "Error redirecting $($KnownFolder). Return code $($result) = $((New-Object System.ComponentModel.Win32Exception($result)).message)"
                Throw $errormsg
            }
        }
    } Else {
        Throw New-Object System.IO.DirectoryNotFoundException "Could not find part of the path $Path."
    }
	
	Attrib +r $Path
    Return $Path
}

Function Get-KnownFolderPath {
    Param (
            [Parameter(Mandatory = $true)]
            [ValidateSet('AdminTools','ApplicationData','CDBurning','CommonAdminTools','CommonApplicationData','CommonDesktopDirectory','CommonDocuments','CommonMusic',`
            'CommonOemLinks','CommonPictures','CommonProgramFiles','CommonProgramFilesX86','CommonPrograms','CommonStartMenu','CommonStartup','CommonTemplates',`
            'CommonVideos','Cookies','Downloads','Desktop','DesktopDirectory','Favorites','Fonts','History','InternetCache','LocalApplicationData','LocalizedResources','MyComputer',`
            'MyDocuments','MyMusic','MyPictures','MyVideos','NetworkShortcuts','Personal','PrinterShortcuts','ProgramFiles','ProgramFilesX86','Programs','Recent',`
            'Resources','SendTo','StartMenu','Startup','System','SystemX86','Templates','UserProfile','Windows')]
            [string]$KnownFolder
    )
    if($KnownFolder -eq "Downloads"){
        Return $Null
    }else{
        Return [Environment]::GetFolderPath($KnownFolder)
    }
}

Function Redirect-Folder {
    Param (
        $GetFolder,
        $SetFolder,
        $Target,
		$copyExistingFiles
    )

    $Folder = Get-KnownFolderPath -KnownFolder $GetFolder
    If ($Folder -ne $Target) {
        Set-KnownFolderPath -KnownFolder $SetFolder -Path $Target
        if($copyExistingFiles -and $Folder){
            Get-ChildItem -Path $Folder -ErrorAction Continue | Copy-Item -Destination $Target -Recurse -Container -Force -Confirm:$False -ErrorAction Continue
        }
        Attrib +h $Folder
    }
}

function New-WebRequest{
    Param(
        $url,
        $method="GET",
        $body,
        $trySSO=1,
        $customHeaders,
        $accept = "application/json",
        $referer = $Null,
        $contentType = "application/x-www-form-urlencoded"
    )
    $maxAttempts = 3
    $attempts=0
    while($true){
        $attempts++
        try{
            $retVal = @{}
            $request = [System.Net.WebRequest]::Create($url)
            $request.KeepAlive = $True
            if($adfsMode -eq 3){
                #Find the FIRST certificate that matches the user's username and append it to all requests
                if($certificateMatchMethod -eq 1){
                    $userCert = (Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {$_.Subject -like "*$($Env:USERNAME)*"})[0]
                }
                #Find the FIRST certificate that matches the template name specified in the script configuration
                if($certificateMatchMethod -eq 2){
                    $attempts = 0
                    while($true){
                        try{
                            $userCert = (Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {($_.extensions.Format(1)[0].split('(')[0]).Split('=')[-1] -like "*$($certificateTemplateName)*"})[0]       
                            if(!$userCert){Throw}
                            break
                        }catch{
                            $attempts++
                            if($attemps -le 3){
                                certutil -user -pulse
                                Start-Sleep -s 10
                            }else{
                                log -text "Failed to find a local certificate to authenticate with" -fout
                                abort_OM
                            }
                        }
                    }
                    
                }                
                $request.ClientCertificates.AddRange($userCert)
            }

            #add device auth certificate
            if($url.startsWith("https://device.login.microsoftonline.com")){
                try{
                    $cert = dir Cert:\LocalMachine\My\ | where { $_.Issuer -match "CN=MS-Organization-Access" }
                    $request.ClientCertificates.AddRange($cert)
                    log -text "detected device authentication prompt and used $($cert.Subject)"
                }catch{
                    log -text "detected device authentication prompt and failed to use/retrieve certificate" -fout
                }
            }

            $request.TimeOut = 10000
            $request.Method = $method
            $request.Referer = $referer
            $request.Accept = $accept
            if($trySSO -eq 1){
                $request.UseDefaultCredentials = $True
            }
            if($customHeaders){
                $customHeaders.Keys | % { 
                    $request.Headers[$_] = $customHeaders.Item($_)
                }
            }
            if($authenticateToProxy){
                $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
                $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
                $request.proxy = $proxy
            }
            $request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.71 Safari/537.36 Edge/12.0 OneDriveMapper/$version"
            $request.ContentType = $contentType
            $request.CookieContainer = $script:cookiejar
            if($method -eq "POST"){
                $body = [byte[]][char[]]$body
                $upStream = $request.GetRequestStream()
                $upStream.Write($body, 0, $body.Length)
                $upStream.Flush()
                $upStream.Close()
            }
            $response = $request.GetResponse()
            $retVal.rawResponse = $response
            $retVal.StatusCode = $response.StatusCode
            $retVal.StatusDescription = $response.StatusDescription
            $retVal.Headers = $response.Headers
            $stream = $response.GetResponseStream()
            $streamReader = [System.IO.StreamReader]($stream)
            $retVal.Content = $streamReader.ReadToEnd()
            $streamReader.Close()
            $response.Close()
            $response = $Null
            $request = $Null
            return $retVal
        }catch{
            if($attempts -ge $maxAttempts){Throw}else{Start-Sleep -s 2}
        }
    }
}

function returnEnclosedFormValue{
    Param(
        $res,
        $searchString,
        $endString = "`"",
        [Switch]$includeEndString,
        [Switch]$decode
    )
    try{
        if($res.Content.Length -le 0){Throw "no request given"}
        if($searchString){$start = $searchString}else{Throw "empty search string"}
        $startLoc = $res.Content.IndexOf($start)+$start.Length
        if($startLoc -eq $start.Length-1){
            return -1
        }
        $searchLength = $res.Content.IndexOf($endString,$startLoc)-$startLoc
        if($searchLength -le 0){
            return -1
        }
        if($includeEndString){
            $searchLength += $endString.Length
        }
        if($decode){
            return([System.Web.HttpUtility]::UrlDecode($res.Content.Substring($startLoc,$searchLength)))
        }else{
            return($res.Content.Substring($startLoc,$searchLength))
        }
    }catch{Throw}
}

function storeSecureString{
    Param(
        $filePath,
        $string
    )
    try{
        $stringForFile = $string | ConvertTo-SecureString -AsPlainText -Force -ErrorAction Stop | ConvertFrom-SecureString -ErrorAction Stop
        Set-Content -Path $filePath -Value $stringForFile -Force -ErrorAction Stop | Out-Null
    }catch{
        Throw "Failed to store string: $($Error[0] | out-string)"
    }
}

function loadSecureString{
    Param(
        $filePath
    )
    try{
        $string = Get-Content $filePath -ErrorAction Stop | ConvertTo-SecureString -ErrorAction Stop
        $string = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($string)
        $string = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($string)
        if($string.Length -lt 3){throw "no valid string returned from cache"}
        return $string
    }catch{
        Throw
    }
}

function startWebDavClient{
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
                [FieldOffset(0)]
                internal UInt64 DataPointer;
                [FieldOffset(8)]
                internal uint Size;
                [FieldOffset(12)]
                internal int Reserved;
            }

            public static void startService(){
                Guid webClientTrigger = new Guid(0x22B6D684, 0xFA63, 0x4578, 0x87, 0xC9, 0xEF, 0xFC, 0xBE, 0x66, 0x43, 0xC7);
                long handle = 0;
                uint output = EventRegister(ref webClientTrigger, IntPtr.Zero, IntPtr.Zero, ref handle);
                bool success = false;
                if (output == 0){
                    EVENT_DESCRIPTOR desc = new EVENT_DESCRIPTOR();
                    unsafe{
                        uint writeOutput = EventWrite(handle, ref desc, 0, null);
                        success = writeOutput == 0;
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
    try{
        log -text "Attempting to automatically start the WebDav client without elevation..."
        $compilerParameters = New-Object System.CodeDom.Compiler.CompilerParameters
        $compilerParameters.CompilerOptions="/unsafe"
        $compilerParameters.GenerateInMemory = $True
        Add-Type -TypeDefinition $Source -Language CSharp -CompilerParameters $compilerParameters
        [JosL.WebClient.Starter]::startService()
        log -text "Start Service Command completed without errors"
        Start-Sleep -s 5
        if((Get-Service -Name WebClient).status -eq "Running"){
            log -text "detected that the webdav client is now running!"
        }else{
            log -text "but the webdav client is still not running! Please set the client to automatically start!" -fout
        }
    }catch{
        Throw "Failed to start the webdav client :( $($Error[0])"
    }
}

function restart_explorer{ 
    log -text "Refreshing Explorer to make the drive(s) visible" 
$definition = @'
[System.Runtime.InteropServices.DllImport("Shell32.dll")] 
private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
public static void Refresh() {
    SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);    
}
'@
    try{
        Add-Type -MemberDefinition $definition -Namespace WinAPI -Name Explorer
        [WinAPI.Explorer]::Refresh()
    }catch{
        log -text "Failed to refresh Explorer" -fout
    }
}  
function queryForAllCreds {
    Param(
        [Parameter(Mandatory=$true)]$titleText,
        [Parameter(Mandatory=$true)]$introText,
        [Parameter(Mandatory=$true)]$buttonText,
        [Parameter(Mandatory=$true)]$loginLabel,
        [Parameter(Mandatory=$true)]$passwordLabel
    )
    $objBalloon = New-Object System.Windows.Forms.NotifyIcon  
    $objBalloon.BalloonTipIcon = "Info" 
    $objBalloon.BalloonTipTitle = $titleText 
    $objBalloon.BalloonTipText = "OneDriveMapper - www.lieben.nu" 
    $objBalloon.Visible = $True  
    $objBalloon.ShowBalloonTip(10000) 
 
    $userForm = New-Object 'System.Windows.Forms.Form' 
    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState' 
    $Form_StateCorrection_Load= {$userForm.WindowState = $InitialFormWindowState}  
    $userForm.Text = $titleText 
    $userForm.Size = New-Object System.Drawing.Size(400,380) 
    $userForm.StartPosition = "CenterScreen" 
    $userForm.AutoSize = $False 
    $userForm.MinimizeBox = $False 
    $userForm.MaximizeBox = $False 
    $userForm.SizeGripStyle= "Hide" 
    $userForm.WindowState = "Normal"
    $userForm.FormBorderStyle="Fixed3D" 
    $userForm.KeyPreview = $True 
    $userForm.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$userForm.Close()}})   

    $OKButton = New-Object System.Windows.Forms.Button 
    $OKButton.Location = New-Object System.Drawing.Size(290,300) 
    $OKButton.Size = New-Object System.Drawing.Size(90,30) 
    $OKButton.Text = $buttonText
    $OKButton.Add_Click({$userForm.Close()}) 
    $userForm.Controls.Add($OKButton)  

    $objTextBox = New-Object System.Windows.Forms.TextBox 
    $objTextBox.Location = New-Object System.Drawing.Size(10,10) 
    $objTextBox.Size = New-Object System.Drawing.Size(370,180) 
    $objTextBox.Multiline = $True
    $objTextBox.ReadOnly = $True
    $objTextBox.Text = $introText
    $userForm.Controls.Add($objTextBox)  

    $userLabel = New-Object System.Windows.Forms.Label 
    $userLabel.Location = New-Object System.Drawing.Size(10,200) 
    $userLabel.Size = New-Object System.Drawing.Size(370,20) 
    $userLabel.Text = $loginLabel
    $userForm.Controls.Add($userLabel)  

    $objTextBox2 = New-Object System.Windows.Forms.TextBox 
    $objTextBox2.Location = New-Object System.Drawing.Size(10,220) 
    $objTextBox2.Size = New-Object System.Drawing.Size(370,20) 
    $objTextBox2.Name = "loginField"
    if($userUPN.Length -gt 5 -and $userUPN.IndexOf("@") -ne -1){
        $objTextBox2.Text = $userUPN
    }else{
        $objTextBox2.Text = ""
    }
    $userForm.Controls.Add($objTextBox2) 

    $userLabel2 = New-Object System.Windows.Forms.Label 
    $userLabel2.Location = New-Object System.Drawing.Size(10,250) 
    $userLabel2.Size = New-Object System.Drawing.Size(370,20) 
    $userLabel2.Text = $passwordLabel 
    $userForm.Controls.Add($userLabel2)  

    $objTextBox3 = New-Object System.Windows.Forms.TextBox 
    $objTextBox3.UseSystemPasswordChar = $True
    $objTextBox3.Location = New-Object System.Drawing.Size(10,270) 
    $objTextBox3.Size = New-Object System.Drawing.Size(370,20) 
    $objTextBox3.Text = ""
    $objTextBox3.Name = "passwordField"
    $userForm.Controls.Add($objTextBox3) 

    $userForm.Topmost = $True 
    $userForm.TopLevel = $True 
    $userForm.ShowIcon = $True 
    $userForm.Add_Shown({$userForm.Activate();
    if($userUPN.Length -gt 5 -and $userUPN.IndexOf("@") -ne -1){
        $objTextBox3.focus()
    }else{
        $objTextBox2.focus()
    }}) 
    $InitialFormWindowState = $userForm.WindowState 
    $userForm.add_Load($Form_StateCorrection_Load) 
    [void]$userForm.ShowDialog() 

    $objTextBox2.Text
    $objTextBox3.Text
}

function checkIfAtO365URL{
    param(
        $userUPN,
        [Array]$finalURLs
    )
    $url = $script:ie.LocationURL
    foreach($item in $finalURLs){
        if($url.StartsWith($item)){
            return $True
        }
    }
    $lookupQuery = $userUPN -replace "@","_"
    $lookupQuery = $lookupQuery -replace "\.","_"
    $attempts = 0
    while($true){
        $attempts++
        try{
            try{
                $userTile = getElementById -id $lookupQuery
                log -text "detected user logged in Tile in IE"
                $userTile.Click()
                waitForIE
                Start-Sleep -m 500
                waitForIE
            }catch{$Null}
            $url = $script:ie.LocationURL
            foreach($item in $finalURLs){
                if($url.StartsWith($item)){
                    return $True
                }
            }
        }catch{
            log -text "Failed to detect or use logged in Tile in IE: $($Error[0])" -fout
        }
        if($attempts -gt 3){
            break
        }
    }
    return $False
}

#region basicFunctions
function lookupLoginFromAD{
    Param(
        [Switch]$lookupEmail #otherwise lookup UPN
    ) 
    if($lookupEmail){
        try{
            $userMail = ([ADSISEARCHER]"samaccountname=$($env:USERNAME)").Findone().Properties.mail
            if($userMail){
                return $userMail
            }else{Throw $Null}
        }catch{
            log -text "Failed to lookup email, active directory connection failed, please change userLookupMode" -fout
            Throw
        }
    }else{
        try{ 
            $objDomain = New-Object System.DirectoryServices.DirectoryEntry 
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
            $objSearcher.SearchRoot = $objDomain 
            $objSearcher.Filter = "(&(objectCategory=User)(SAMAccountName=$Env:USERNAME))"
            $objSearcher.SearchScope = "Subtree"
            $objSearcher.PropertiesToLoad.Add("userprincipalname") | Out-Null 
            $results = $objSearcher.FindAll() 
            return $results[0].Properties.userprincipalname 
        }catch{ 
            log -text "Failed to lookup username, active directory connection failed, please change userLookupMode" -fout
            $script:errorsForUser += "Could not connect to your corporate network.`n"
            Throw 
        }
    }
}

function CustomInputBox(){ 
    Param(
        [String]$title,
        [String]$message,
        [Switch]$password
    )
    if($forcePassword.Length -gt 2 -and $password) { 
        return $forcePassword 
    } 
    $objBalloon = New-Object System.Windows.Forms.NotifyIcon  
    $objBalloon.BalloonTipIcon = "Info" 
    $objBalloon.BalloonTipTitle = "OneDriveMapper"  
    $objBalloon.BalloonTipText = "OneDriveMapper - www.lieben.nu" 
    $objBalloon.Visible = $True  
    $objBalloon.ShowBalloonTip(10000) 
 
    $userForm = New-Object 'System.Windows.Forms.Form' 
    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState' 
    $Form_StateCorrection_Load= 
    { 
        $userForm.WindowState = $InitialFormWindowState 
    }  
    $userForm.Text = "$title" 
    $userForm.Size = New-Object System.Drawing.Size(350,200) 
    $userForm.StartPosition = "CenterScreen" 
    $userForm.AutoSize = $False 
    $userForm.MinimizeBox = $False 
    $userForm.MaximizeBox = $False 
    $userForm.SizeGripStyle= "Hide" 
    $userForm.WindowState = "Normal" 
    $userForm.FormBorderStyle="Fixed3D" 
    $userForm.KeyPreview = $True 
    $userForm.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$userForm.Close()}})   
    $OKButton = New-Object System.Windows.Forms.Button 
    $OKButton.Location = New-Object System.Drawing.Size(105,110) 
    $OKButton.Size = New-Object System.Drawing.Size(95,23) 
    $OKButton.Text = $buttonText 
    $OKButton.Add_Click({$userForm.Close()}) 
    $userForm.Controls.Add($OKButton) 
    $userLabel = New-Object System.Windows.Forms.Label 
    $userLabel.Location = New-Object System.Drawing.Size(10,20) 
    $userLabel.Size = New-Object System.Drawing.Size(300,50) 
    $userLabel.Text = "$message" 
    $userForm.Controls.Add($userLabel)  
    $objTextBox = New-Object System.Windows.Forms.TextBox 
    if($password) {$objTextBox.UseSystemPasswordChar = $True }
    $objTextBox.Location = New-Object System.Drawing.Size(70,75) 
    $objTextBox.Size = New-Object System.Drawing.Size(180,20) 
    $userForm.Controls.Add($objTextBox)  
    $userForm.Topmost = $True 
    $userForm.TopLevel = $True 
    $userForm.ShowIcon = $True 
    $userForm.Add_Shown({$userForm.Activate();$objTextBox.focus()}) 
    $InitialFormWindowState = $userForm.WindowState 
    $userForm.add_Load($Form_StateCorrection_Load) 
    [void] $userForm.ShowDialog() 
    return $objTextBox.Text 
} 
 
function labelDrive{ 
    Param( 
    [String]$lD_DriveLetter, 
    [String]$lD_MapURL, 
    [String]$lD_DriveLabel 
    ) 
 
    #try to set the drive label 
    if($lD_DriveLabel.Length -gt 0){ 
        log -text "A drive label has been specified, attempting to set the label for $($lD_DriveLetter)"
        try{ 
            $regURL = $lD_MapURL.TrimEnd("\") -Replace [regex]::escape("\"),"#"
            $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($regURL)" 
            $Null = New-Item -Path $path -Force -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_CommentFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel -ErrorAction SilentlyContinue
            $regURL = $regURL -Replace [regex]::escape("DavWWWRoot#"),"" 
            $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($regURL)" 
            $Null = New-Item -Path $path -Force -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_CommentFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel -ErrorAction SilentlyContinue
            log -text "Label has been set to $($lD_DriveLabel)" 
 
        }catch{ 
            log -text "Failed to set the drive label registry keys: $($Error[0]) " -fout
        } 
 
    } 
} 

function fixElevationVisibility{
    #check if a task already exists for this script
    if($showElevatedConsole){
        $createTask = "schtasks /Create /SC ONCE /TN OnedriveMapper /IT /RL LIMITED /F /TR `"Powershell.exe -NoProfile -ExecutionPolicy ByPass -File '$scriptPath' -asTask`" /st 00:00"    
    }else{
        $createTask = "schtasks /Create /SC ONCE /TN OnedriveMapper /IT /RL LIMITED /F /TR `"Powershell.exe -NoProfile -ExecutionPolicy ByPass -WindowStyle Hidden -File '$scriptPath' -asTask`" /st 00:00"
    }
    $res = Invoke-Expression $createTask
    if($res -NotMatch "ERROR"){
        log -text "Scheduled a task to run OnedriveMapper unelevated because this script cannot run elevated"
        $runTask = "schtasks /Run /TN OnedriveMapper /I"
        $res = Invoke-Expression $runTask
        if($res -NotMatch "ERROR"){
            log -text "Scheduled task started"
        }else{
            log -text "Failed to start a scheduled task to run OnedriveMapper without elevation because: $res" -fout
        }
    }else{
        log -text "Failed to schedule a task to run OnedriveMapper without elevation because: $res" -fout
    }
}

function MapDrive{ 
    Param( 
        $driveMapping
    )
    if($driveMapping.targetLocationType -eq "driveletter"){
        $LASTEXITCODE = 0
        log -text "Mapping target: $($driveMapping.webDavPath)" 
        try{$del = NET USE $($driveMapping.targetLocationPath) /DELETE /Y 2>&1}catch{$Null}
        if($persistentMapping){
            try{$out = NET USE $($driveMapping.targetLocationPath) $($driveMapping.webDavPath) /PERSISTENT:YES 2>&1}catch{$Null}
        }else{
            try{$out = NET USE $($driveMapping.targetLocationPath) $($driveMapping.webDavPath) /PERSISTENT:NO 2>&1}catch{$Null}
        }
        if($out -like "*error 67*"){
            log -text "ERROR: detected string error 67 in return code of net use command, this usually means the WebClient isn't running" -fout
        }
        if($out -like "*error 224*"){
            log -text "ERROR: detected string error 224 in return code of net use command, this usually means your trusted sites are misconfigured or KB2846960 is missing or Internet Explorer needs a reset" -fout
        }
        if($LASTEXITCODE -ne 0){ 
            log -text "Failed to map $($driveMapping.targetLocationPath) to $($driveMapping.webDavPath), error: $($LASTEXITCODE) $($out) $del" -fout
            $script:errorsForUser += "$($driveMapping.targetLocationPath) could not be mapped because of error $($LASTEXITCODE) $($out) d$del`n"
            return $False 
        } 
        if([System.IO.Directory]::Exists($driveMapping.targetLocationPath)){ 
            #set drive label 
            $Null = labelDrive $($driveMapping.targetLocationPath) $($driveMapping.webDavPath) $($driveMapping.displayName)
            log -text "$($driveMapping.targetLocationPath) mapped successfully`n" 
            return $True 
        }else{ 
            log -text "failed to contact $($driveMapping.targetLocationPath) after mapping it to $($driveMapping.webDavPath), check if the URL is valid. Error: $($error[0]) $out" -fout
            return $False 
        }
    }else{
        try{
            if($driveMapping.sourceLocationPath -eq "autodetect"){
                $desiredIconPath = $onedriveIconPath
            }elseif($driveMapping.mapOnlyForSpecificGroup -eq "favoritesPlaceholder"){
                $desiredIconPath = $teamsIconPath
            }else{
                $desiredIconPath = $sharepointIconPath
            }
            Add-NetworkLocation -networkLocationPath $($driveMapping.targetLocationPath) -networkLocationName $($driveMapping.displayName) -networkLocationTarget $($driveMapping.webDavPath) -iconPath $desiredIconPath -Verbose
            log -text "Added network location $($driveMapping.displayName)"
        }catch{
            log -text "failed to add network location: $($Error[0])" -fout
        }
    }
} 

function handleDuoMFA{
    Param(
        $res,
        $clientId
    )
    ##First redirect
    $federationUrl = returnEnclosedFormValue -res $res -searchString "action=`"" -endString "`""
    if($federationUrl -notlike "*/federation/redirecttoexternal*"){
        return $res
    }

    try{
        $id_token_hint = returnEnclosedFormValue -res $res -searchString "name=`"id_token_hint`" value=`"" -endString "`""
        $claims = "%7B%22id_token%22%3A%7B%22DuoMfa%22%3A%7B%22essential%22%3Atrue%2C%22value%22%3A%22MfaDone%22%7D%7D%7D&"
        $client_request_id = returnEnclosedFormValue -res $res -searchString "name=`"client-request-id`" value=`"" -endString "`""
        $client_id = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"client_id`" value=`"" -endString "`""))
        $nonce = returnEnclosedFormValue -res $res -searchString "name=`"nonce`" value=`"" -endString "`""
        $ExternalClaimsProviderAuthorizeEndpointUri = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"ExternalClaimsProviderAuthorizeEndpointUri`" value=`"" -endString "`""))
        $state = returnEnclosedFormValue -res $res -searchString "name=`"state`" value=`"" -endString "`""
        $canary = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"canary`" value=`"" -endString "`""))
        $flowtoken = returnEnclosedFormValue -res $res -searchString "name=`"flowtoken`" value=`"" -endString "`""
        $body = "scope=openid&response_mode=form_post&id_token_hint=$id_token_hint&response_type=id_token&client_id=$client_id&redirect_uri=$([System.Web.HttpUtility]::UrlEncode("https://login.microsoftonline.com/common/federation/OAuth2ClaimsProvider"))&claims=$claims&client-request-id=$client_request_id&nonce=$nonce&ExternalClaimsProviderAuthorizeEndpointUri=$ExternalClaimsProviderAuthorizeEndpointUri&state=$state&canary=$canary&flowtoken=$flowtoken"
        $res = New-WebRequest -url $federationUrl -Method POST -body $body
    }catch{
        Throw "Failed to follow federation redirect because of $_"
    }

    #second redirect
    try{
        $redirectToDuoUrl = returnEnclosedFormValue -res $res -searchString "action=`"" -endString "`""
        $state = returnEnclosedFormValue -res $res -searchString "name=`"state`" value=`"" -endString "`""
        $id_token_hint = returnEnclosedFormValue -res $res -searchString "name=`"id_token_hint`" value=`"" -endString "`""
        $id_token = returnEnclosedFormValue -res $res -searchString "name=`"id_token`" value=`"" -endString "`""
        $client_request_id = returnEnclosedFormValue -res $res -searchString "name=`"client-request-id`" value=`"" -endString "`""
        $canary = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"canary`" value=`"" -endString "`""))
        $nonce = returnEnclosedFormValue -res $res -searchString "name=`"nonce`" value=`"" -endString "`""
        $customHeaders = @{"Sec-Fetch-Site"="cross-site";"Sec-Fetch-Mode"="navigate";"Sec-Fetch-Dest"="document";"Accept-Language"="en-US,en;q=0.9";"Upgrade-Insecure-Requests" = 1;"DNT" = 1}
        $body = "state=$state&scope=openid&response_mode=form_post&id_token_hint=$id_token_hint&response_type=id_token&client_id=$client_id&redirect_uri=$([System.Web.HttpUtility]::UrlEncode("https://login.microsoftonline.com/common/federation/OAuth2ClaimsProvider"))&claims=$claims&client-request-id=$client_request_id&canary=$canary&nonce=$nonce&cxh_inclusive=False"
        $res = New-WebRequest -url $redirectToDuoUrl -Method POST -body $body  -customHeaders $customHeaders -contentType "application/x-www-form-urlencoded" -referer "https://login.microsoftonline.com/" -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
    }catch{
        Throw "Failed to follow federation to Duo redirect because of $_"
    }

    #third redirect
    try{
        $redirectToDuoUrl2 = returnEnclosedFormValue -res $res -searchString "action=`"" -endString "`""
        $client_id = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"client_id`" value=`"" -endString "`""))
        $request = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"request`" value=`"" -endString "`""))
        $req_trace_group = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"req-trace-group`" value=`"" -endString "`""))
        $redirectToDuoUrl2 = "https://$($redirectToDuoUrl2.Split(";")[3].Split(".")[0]).duosecurity.com/oauth/v1/authorize?response_type=code&client_id=$client_id&request=$request&req-trace-group=$req_trace_group"
        $api = $redirectToDuoUrl2.Split("/")[2]
        $res = New-WebRequest -url $redirectToDuoUrl2 -Method GET -customHeaders $customHeaders -referer $redirectToDuoUrl -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
    }catch{
        Throw "Failed to follow Duo redirect #2 because of $_"
    }


    #frame redirect follower
    try{
        $xsrf = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"_xsrf`" value=`"" -endString "`""))
        $body = "tx=$request&parent=None&_xsrf=$xsrf&java_version=&flash_version=&screen_resolution_width=1920&screen_resolution_height=1080&color_depth=24&ch_ua_error=&client_hints=eyJicmFuZHMiOlt7ImJyYW5kIjoiQ2hyb21pdW0iLCJ2ZXJzaW9uIjoiMTEyIn0seyJicmFuZCI6Ikdvb2dsZSBDaHJvbWUiLCJ2ZXJzaW9uIjoiMTEyIn0seyJicmFuZCI6Ik5vdDpBLUJyYW5kIiwidmVyc2lvbiI6Ijk5In1dLCJmdWxsVmVyc2lvbkxpc3QiOlt7ImJyYW5kIjoiQ2hyb21pdW0iLCJ2ZXJzaW9uIjoiMTEyLjAuNTYxNS4xMzgifSx7ImJyYW5kIjoiR29vZ2xlIENocm9tZSIsInZlcnNpb24iOiIxMTIuMC41NjE1LjEzOCJ9LHsiYnJhbmQiOiJOb3Q6QS1CcmFuZCIsInZlcnNpb24iOiI5OS4wLjAuMCJ9XSwibW9iaWxlIjpmYWxzZSwicGxhdGZvcm0iOiJXaW5kb3dzIiwicGxhdGZvcm1WZXJzaW9uIjoiMTUuMC4wIiwidWFGdWxsVmVyc2lvbiI6IjExMi4wLjU2MTUuMTM4In0%3D&is_cef_browser=false&is_ipad_os=false&is_ie_compatibility_mode=&is_user_verifying_platform_authenticator_available=false&user_verifying_platform_authenticator_available_error=&acting_ie_version=&react_support=true&react_support_error_message="
        $res = New-WebRequest -url $res.rawResponse.ResponseUri.AbsoluteUri -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
    }catch{
        Throw "Failed to follow frame redirect #1 because of $_"
    }

    #follow redirect token to O365 if already given
    try{
        $id_token = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"id_token`" value=`"" -endString "`""))
        if($id_token -and $id_token -ne -1){
            $expires_in = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"expires_in`" value=`"" -endString "`""))
            $state = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"state`" value=`"" -endString "`""))
            $body = "token_type=Bearer&id_token=$id_token&expires_in=$expires_in&state=$state&scope=openid"
            $res = New-WebRequest -url "https://login.microsoftonline.com/common/federation/OAuth2ClaimsProvider" -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
            return $res
        }
    }catch{
        Throw "Failed to use token for O365 because of $_"
    }

    #healthcheck follower
    try{
        $uri = $res.rawResponse.ResponseUri.AbsoluteUri.Replace("?sid","/data?sid")
        $res = New-WebRequest -url $uri -Method GET -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
    }catch{
        Throw "Failed to follow healthcheck redirect #1 because of $_"
    }

    #return
    try{
        $sid = $res.rawResponse.ResponseUri.AbsoluteUri.Split("?")[1]
        $uri = "https://$api/frame/v4/return?$sid"
        $res = New-WebRequest -url $uri -Method GET -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
    }catch{
        Throw "Failed to follow healthcheck redirect #1 because of $_"
    }

    #frame redirect follower #2
    try{
        $xsrf = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"_xsrf`" value=`"" -endString "`""))
        $body = "tx=$request&parent=None&_xsrf=$xsrf&java_version=&flash_version=&screen_resolution_width=1920&screen_resolution_height=1080&color_depth=24&ch_ua_error=&client_hints=eyJicmFuZHMiOlt7ImJyYW5kIjoiQ2hyb21pdW0iLCJ2ZXJzaW9uIjoiMTEyIn0seyJicmFuZCI6Ikdvb2dsZSBDaHJvbWUiLCJ2ZXJzaW9uIjoiMTEyIn0seyJicmFuZCI6Ik5vdDpBLUJyYW5kIiwidmVyc2lvbiI6Ijk5In1dLCJmdWxsVmVyc2lvbkxpc3QiOlt7ImJyYW5kIjoiQ2hyb21pdW0iLCJ2ZXJzaW9uIjoiMTEyLjAuNTYxNS4xMzgifSx7ImJyYW5kIjoiR29vZ2xlIENocm9tZSIsInZlcnNpb24iOiIxMTIuMC41NjE1LjEzOCJ9LHsiYnJhbmQiOiJOb3Q6QS1CcmFuZCIsInZlcnNpb24iOiI5OS4wLjAuMCJ9XSwibW9iaWxlIjpmYWxzZSwicGxhdGZvcm0iOiJXaW5kb3dzIiwicGxhdGZvcm1WZXJzaW9uIjoiMTUuMC4wIiwidWFGdWxsVmVyc2lvbiI6IjExMi4wLjU2MTUuMTM4In0%3D&is_cef_browser=false&is_ipad_os=false&is_ie_compatibility_mode=&is_user_verifying_platform_authenticator_available=true&user_verifying_platform_authenticator_available_error=&acting_ie_version=&react_support=true&react_support_error_message="
        $res = New-WebRequest -url $res.rawResponse.ResponseUri.AbsoluteUri -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
    }catch{
        Throw "Failed to follow frame redirect #1 because of $_"
    }

    #follow redirect token to O365 if already given
    try{
        $id_token = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"id_token`" value=`"" -endString "`""))
        if($id_token -and $id_token -ne -1){
            $expires_in = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"expires_in`" value=`"" -endString "`""))
            $state = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"state`" value=`"" -endString "`""))
            $body = "token_type=Bearer&id_token=$id_token&expires_in=$expires_in&state=$state&scope=openid"
            $res = New-WebRequest -url "https://login.microsoftonline.com/common/federation/OAuth2ClaimsProvider" -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
            return $res
        }
    }catch{
        Throw "Failed to use token for O365 because of $_"
    }
    
    #get auth method
    try{
        $sid = $res.rawResponse.ResponseUri.AbsoluteUri.Split("=")[1].Split("&")[0]
        $initiatePostUri = "https://$($redirectToDuoUrl2.Split("/")[2])/frame/v4/auth/prompt/data?post_auth_action=OIDC_EXIT&sid=$sid"
        $res = New-WebRequest -url $initiatePostUri -Method GET -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
        $mfaInfo = $res.content | convertfrom-json
        $mfaPhone = if($mfaInfo.response.phones){$mfaInfo.response.phones[0]}else{$False}
        $mfaHwToken = $mfaInfo.response.auth_method_order[0].factor -eq "Hardware Token"
        if(!$mfaPhone -and !$mfaHwToken){
            Throw "No MFA phone or HW token registered in DUO"
        }
    }catch{
        Throw "Failed to get DUO auth method for this user because of $_"
    }

    if($mfaPhone){
        #prompt user's first device
        try{
            $script:label1.text="PLEASE APPROVE CISCO DUO REQUEST TO CONTINUE"
            $script:form1.Refresh()
            $body = "device=$($mfaPhone.index)&factor=Duo+Push&sid=$sid"
            $res = New-WebRequest -url "https://$($api)/frame/v4/prompt" -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
            $txid = ($res.content | convertfrom-json).response.txid
        }catch{
            Throw "Failed to send push because of $_"
        }
    
        #wait until user completes login
        $timeSpentInSeconds = 0
        while($true){
            #check status
            try{
                $body = "txid=$txid&sid=$sid"
                $res = New-WebRequest -url "https://$($api)/frame/v4/status" -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
                $mfaResult = $res.content | convertfrom-json
            }catch{
                log -text "Failed to send push to user's device because of $_" -fout
                continue
            }
            if($mfaResult.response.status_code -eq "allow"){
                break #good
            }
            if($mfaResult.response.status_code -eq "deny"){
                log -text "user denied MFA request on mobile device" -fout
                continue
            }

            Start-Sleep -s 1
            $timeSpentInSeconds += 2
            if($timeSpentInSeconds -gt 120){
                Throw "timeout waiting for user to approve (120seconds)"
            }
        }

        $script:label1.text=$script:progressBarText
        $script:form1.Refresh()

        #close MFA session and get redirect
        try{
            $body = "txid=$txid&sid=$sid&factor=Duo+Push&device_key=$($mfaPhone.key)&_xsrf=$xsrf&dampen_choice=true"
            $res = New-WebRequest -url "https://$($api)/frame/v4/oidc/exit" -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
        }catch{
            Throw "Failed to close OIDC session because of $_"
        }
    }elseif($mfaHwToken){
        #wait until user completes login
        $attempts = 0
        while($true){
            $attempts++
            if($attempts -gt 5){
                Throw "User failed to auth using a DUO passcode for more than 5 times"
            }

            #Query user for passcode and send
            try{
                $script:label1.text="PLEASE PROVIDE CISCO DUO CODE TO CONTINUE"
                $script:form1.Refresh()
                $duoCode = CustomInputBox -title "DUO Challenge" -message "Please enter the 6-digit code from your DUO token to continue"
                if(!$duoCode){
                    log -text "User did not enter a DUO passcode" -fout
                    continue
                }else{
                    $body = "device=null&factor=Passcode&sid=$sid&passcode=$duoCode"
                    $res = New-WebRequest -url "https://$($api)/frame/v4/prompt" -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
                    $txid = ($res.content | convertfrom-json).response.txid
                }
            }catch{
                log -text "Failed to send code because of $_" -fout
            }

            #check status
            try{
                $body = "txid=$txid&sid=$sid"
                $res = New-WebRequest -url "https://$($api)/frame/v4/status" -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
                $mfaResult = $res.content | convertfrom-json
            }catch{
                log -text "Failed to check code validity because of $_" -fout
            }
            if($mfaResult.response.status_code -eq "allow"){
                break #good
            }
            if($mfaResult.response.status_code -eq "deny"){
                log -text "user did not enter a correct DUO code" -fout
            }
        }

        $script:label1.text=$script:progressBarText
        $script:form1.Refresh()

        #close MFA session and get redirect
        try{
            $body = "txid=$txid&sid=$sid&factor=Hardware+Token&device_key=&_xsrf=$xsrf&dampen_choice=true"
            $res = New-WebRequest -url "https://$($api)/frame/v4/oidc/exit" -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
        }catch{
            Throw "Failed to close OIDC session because of $_"
        }
    }

    #follow redirect token to O365
    try{
        $id_token = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"id_token`" value=`"" -endString "`""))
        $expires_in = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"expires_in`" value=`"" -endString "`""))
        $state = [System.Web.HttpUtility]::UrlEncode((returnEnclosedFormValue -res $res -searchString "name=`"state`" value=`"" -endString "`""))
        $body = "token_type=Bearer&id_token=$id_token&expires_in=$expires_in&state=$state&scope=openid"
        $res = New-WebRequest -url "https://login.microsoftonline.com/common/federation/OAuth2ClaimsProvider" -Method POST -body $body -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri -accept "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
    }catch{
        Throw "Failed to use token for O365 because of $_"
    }

    return $res
}
 
function abort_OM{ 
    if($showProgressBar) {
        $script:progressbar1.Value = 100
        $script:label1.text="Done!"
        Start-Sleep -Milliseconds 500
        $form1.Close()
    }
    if($restartExplorer){ 
        restart_explorer 
    }else{ 
        log -text "restartExplorer is set to False, if you're redirecting folders they may not show up" -warning 
    }     
    log -text "OnedriveMapper has finished running"
    if($urlOpenAfter.Length -gt 10){Start-Process iexplore.exe $urlOpenAfter}
    if($displayErrors){
        if($errorsForUser){ 
            $OUTPUT= [System.Windows.Forms.MessageBox]::Show($errorsForUser, "Onedrivemapper Error" , 0) 
            $OUTPUT2= [System.Windows.Forms.MessageBox]::Show("You can always use https://portal.office.com to access your data", "Need a workaround?" , 0) 
        }
    }
    Exit 
} 
 
function askForPassword{
    $askAttempts = 0
    do{ 
        $askAttempts++ 
        log -text "asking user for password" 
        try{ 
            $password = CustomInputBox $loginformTitleText $passwordFieldText -password
        }catch{ 
            log -text "failed to display a password input box, exiting. $($Error[0])" -fout
            abort_OM              
        } 
    } 
    until($password.Length -gt 0 -or $askAttempts -gt 2) 
    if($askAttempts -gt 2) { 
        log -text "user refused to enter a password, exiting" -fout
        $script:errorsForUser += "You did not enter a password, we will not be able to connect to Onedrive`n"
        abort_OM 
    }else{ 
        return $password 
    }
}

function retrievePassword{ 
    Param(
        [switch]$forceNewPassword,
        [String]$cachePassword,
        [Switch]$noQuery
    )
    if($forcePassword){
        return $forcePassword
    }

    if($forceNewPassword){
        $password = askForPassword
        if($pwdCache){
            try{
                $res = storeSecureString -filePath $pwdCache -string $password
                log -text "Stored user's new password to user password cache file $pwdCache"
            }catch{
                log -text "Error storing user password to user password cache file ($($Error[0] | out-string)" -fout
            }
        }    
        return $password
    }
    if($pwdCache -and $cachePassword.Length -le 0){
        try{
            $res = loadSecureString -filePath $pwdCache
            log -text "Retrieved user password from cache $pwdCache"
            return $res
        }catch{
            log -text "Failed to retrieve user password from cache: $($Error[0])" -fout
            if($noQuery){
                return -1
            }
        }
    }
    if($cachePassword.Length -gt 0){
        $password = $cachePassword
        log -text "Password for caching supplied, not querying user for password"
    }else{
        $password = askForPassword
    }
    if($pwdCache){
        try{
            $res = storeSecureString -filePath $pwdCache -string $password
            log -text "Stored user's new password to user password cache file $pwdCache"
        }catch{
            log -text "Error storing user password to user password cache file ($($Error[0] | out-string)" -fout
        }
    }
    return $password
} 
 
function askForUserName{ 
    $askAttempts = 0
    do{ 
        $askAttempts++ 
        log -text "asking user for login" 
        try{ 
            $login = CustomInputBox $loginformTitleText $loginFieldText
        }catch{ 
            log -text "failed to display a login input box, exiting $($Error[0])" -fout
            abort_OM              
        } 
    } 
    until($login.Length -gt 0 -or $askAttempts -gt 2) 
    if($askAttempts -gt 2) { 
        log -text "user refused to enter a login name, exiting" -fout
        $script:errorsForUser += "You did not enter a login name, script cannot continue`n"
        abort_OM 
    }else{ 
        return $login 
    } 
}

function retrieveLogin{ 
    Param(
        [switch]$forceNewLogin,
        [string]$cacheLogin,
        [Switch]$noQuery
    )
    if($forceUserName){
        return $forceUserName
    }
    if($forceNewLogin){
        $login = askForUserName
        if($loginCache){
            try{
                $res = storeSecureString -filePath $loginCache -string $login
                log -text "Stored user's new login to user login cache file $loginCache"
            }catch{
                log -text "Error storing user login to user login cache file ($($Error[0] | out-string)" -fout
            }
        }
        return [String]$login.ToLower()
    }
    if($loginCache -and $cacheLogin.Length -le 0){
        try{
            $res = loadSecureString -filePath $loginCache
            log -text "Retrieved user login from cache $loginCache"
            return [String]$res.ToLower()
        }catch{
            log -text "Failed to retrieve user login from cache: $($Error[0])" -fout
            if($noQuery){
                return -1
            }
        }
    }
    if($cacheLogin.Length -gt 0){
        log -text "Login for caching supplied, not querying user for login"
        $login = $cacheLogin
    }else{
        $login = askForUserName
    }
    if($loginCache){
        try{
            $res = storeSecureString -filePath $loginCache -string $login
            log -text "Stored user's new login to user login cache file $loginCache"
        }catch{
            log -text "Error storing user login to user login cache file ($($Error[0] | out-string)" -fout
        }
    }
    return [String]$login.ToLower()
} 



#cookie setter function, only works for rtFA and FedAuth cookies
function setCookies{
    [DateTime]$dateTime = Get-Date
    $dateTime = $dateTime.AddDays(5)
    $str = $dateTime.ToString("R")
    $relevantCookies += $script:cookiejar.GetCookies("https://$O365CustomerName-my.sharepoint.com")
    $relevantCookies += $script:cookiejar.GetCookies("https://$O365CustomerName.sharepoint.com")
    foreach($cookie in $relevantCookies){
        [String]$cookieValue = [String]$cookie.Value.Trim()
        [String]$cookieDomain = [String]$cookie.Domain.Trim()
        try{
            if($cookie.Name -eq "rtFa"){
                $cookieDomain = "https://$($cookieDomain)"
                log -text "Setting rtFA cookie for $cookieDomain...."
                $res = [Cookies.setter]::SetWinINETCookieString($cookieDomain,"rtFa","$cookieValue;Expires=$str")
            }
            if($cookie.Name -eq "FedAuth"){
                $cookieDomain = "https://$($cookieDomain)"
                log -text "Setting FedAuth cookie for $cookieDomain...."
                $res = [Cookies.setter]::SetWinINETCookieString($cookieDomain,"FedAuth","$cookieValue;Expires=$($str)")
            }
        }catch{
            log -text "Failed to set a cookie: $($Error[0])" -fout
        }
    }
}

function askForCode{ 
    $askAttempts = 0
    do{ 
        $askAttempts++ 
        log -text "asking user for SMS or App code" 
        try{ 
            $login = CustomInputBox "Microsoft Office 365 OneDrive" "Please enter the SMS or Authenticator App code you have received on your cellphone"
        }catch{ 
            log -text "failed to display a code input box, exiting $($Error[0])" -fout
            abort_OM              
        } 
    } 
    until($login.Length -gt 0 -or $askAttempts -gt 2) 
    if($askAttempts -gt 2) { 
        log -text "user refused to enter an SMS code, exiting" -fout
        $script:errorsForUser += "You did not enter an SMS code, script cannot continue`n"
        abort_OM 
    }else{ 
        return $login 
    } 
}

function checkIfAtFhmPage{
    Param(
        [parameter(mandatory=$true)]$res
    )
    $nextURL = returnEnclosedFormValue -res $res -searchString "form name=`"fmHF`" id=`"fmHF`" action=`"" -decode
    $nextURL = [System.Web.HttpUtility]::HtmlDecode($nextURL)
    $value = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"t`" id=`"t`" value=`""
    if($value -eq -1){
        Throw "Not at Fhm page"
    }else{
        $body = "t=$value"
        try{
            $res = New-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/x-www-form-urlencoded" -accept "text/html, application/xhtml+xml, image/jxr, */*"       
            $value = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"t`" id=`"t`" value=`""
            if($value -ne -1){
                return $res
            }else{
                Throw "Was at Fhm page, but no new t value received"
            }
        }catch{
            Throw
        }
    }
}
function Get-ProcessWithOwner { 
    param( 
        [parameter(mandatory=$true,position=0)]$ProcessName 
    ) 
    $ComputerName=$env:COMPUTERNAME 
    $UserName=$env:USERNAME 
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($(New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$('ProcessName','UserName','Domain','ComputerName','handle')))) 
    try { 
        $Processes = Get-wmiobject -Class Win32_Process -ComputerName $ComputerName -Filter "name LIKE '$ProcessName%'" 
    } catch { 
        return -1 
    } 
    if ($Processes -ne $null) { 
        $OwnedProcesses = @() 
        foreach ($Process in $Processes) { 
            if($Process.GetOwner().User -eq $UserName){ 
                $Process |  
                Add-Member -MemberType NoteProperty -Name 'Domain' -Value $($Process.getowner().domain) 
                $Process | 
                Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value $ComputerName  
                $Process | 
                Add-Member -MemberType NoteProperty -Name 'UserName' -Value $($Process.GetOwner().User)  
                $Process |  
                Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers 
                $OwnedProcesses += $Process 
            } 
        } 
        return $OwnedProcesses 
    } else { 
        return 0 
    } 
 
} 
#endregion

function handleO365Redirect{
    Param(
        $res
    )

    $redirectFollowed = $False
    $nextURL = returnEnclosedFormValue -res $res -searchString "form method=`"POST`" name=`"hiddenform`" action=`""
    $nextURL = [System.Web.HttpUtility]::HtmlDecode($nextURL)
    $code = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"code`" value=`""
    $id_token = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"id_token`" value=`""
    $session_state = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"session_state`" value=`""
    if($nextURL -like "*sharepoint.com*"){
        $correlation_id = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"correlation_id`" value=`""
        $body = "code=$([System.Web.HttpUtility]::UrlEncode($code))&id_token=$([System.Web.HttpUtility]::UrlEncode($id_token))&correlation_id=$([System.Web.HttpUtility]::UrlEncode($correlation_id))&session_state=$([System.Web.HttpUtility]::UrlEncode($session_state))"
    }else{
        $state = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"state`" value=`""
        $body = "code=$([System.Web.HttpUtility]::UrlEncode($code))&id_token=$([System.Web.HttpUtility]::UrlEncode($id_token))&state=$([System.Web.HttpUtility]::UrlEncode($state))&session_state=$([System.Web.HttpUtility]::UrlEncode($session_state))"
    }
       
    if($nextURL -ne -1 -and $id_token -ne -1){
        log -text "Detected a id_token redirect, following..."
        try{
            $res = New-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/x-www-form-urlencoded" -accept "text/html, application/xhtml+xml, image/jxr, */*"  
            $redirectFollowed=$True
        }catch{
            log -text "Error detected while following id_token redirect, check the FAQ for help" -fout
            Throw "Failed to follow id_token redirect"           
        }
    }     
    $nextURL = returnEnclosedFormValue -res $res -searchString "form name=`"fmHF`" id=`"fmHF`" action=`"" -decode
    $nextURL = [System.Web.HttpUtility]::HtmlDecode($nextURL)
    $value = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"t`" id=`"t`" value=`""
    if($value -ne -1){
        $body = "t=$value"                
        log -text "Detected fmHF redirect, following"
        try{
            $res = New-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/x-www-form-urlencoded" -accept "text/html, application/xhtml+xml, image/jxr, */*"       
            $redirectFollowed=$True
        }catch{
            log -text "Error detected while following fmHF redirect, check the FAQ for help" -fout
            Throw "Failed to follow fmHF redirect" 
        }    
    }    

    $nextURL = returnEnclosedFormValue -res $res -searchString "form method=`"POST`" name=`"hiddenform`" action=`""
    $nextURL = [System.Web.HttpUtility]::HtmlDecode($nextURL)
    $ctx = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"ctx`" value=`""
    $flowtoken = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"flowtoken`" value=`""
    if($ctx -ne -1){
        $body = "ctx=$([System.Web.HttpUtility]::UrlEncode($ctx))&flowtoken=$([System.Web.HttpUtility]::UrlEncode($flowtoken))"                
        log -text "Detected DeviceAuth redirect, following"
        try{
            $res = New-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/x-www-form-urlencoded" -accept "text/html, application/xhtml+xml, image/jxr, */*"       
            $redirectFollowed=$True
        }catch{
            log -text "Error detected while following DeviceAuth redirect, check the FAQ for help" -fout
            Throw "Failed to follow DeviceAuth redirect" 
        }    
    }   

    return $res,$redirectFollowed     
}

function loginV2(){
    $script:cookiejar = New-Object System.Net.CookieContainer
    if($cacheCookies -and (Test-Path $cookieCacheFilePath)){
        log -text "retrieving cookies from cache"
        try{
            $cookies = Import-Clixml -Path $cookieCacheFilePath
            foreach($cookie in $cookies){
                $script:cookiejar.Add(
                [System.Net.Cookie]@{
                    "Comment" = $cookie.Comment
                    "CommentUri" = $cookie.CommentUri
                    "HttpOnly" = $cookie.HttpOnly
                    "Discard" = $cookie.Discard
                    "Domain" = $cookie.Domain
                    "Expired" = $cookie.Expired
                    "Expires" = $cookie.Expires
                    "Name" = $cookie.Name
                    "Path" = $cookie.Path
                    "Port" = $cookie.Port
                    "Secure" = $cookie.Secure
                    "Value" = $cookie.Value
                    "Version" = $cookie.Version
                }
                )
            } 
        }catch{
            log -text "failed to retrieve cookies from cache $_" -fout
        }
    }

    try{
        $res = New-WebRequest -url "https://$($O365CustomerName)-my.sharepoint.com" -method GET
    }catch{$Null}
    if($res.rawResponse.ResponseUri.AbsoluteUri -like "*/personal/*"){
        log -text "Already logged in, using existing session silently"
        Return $True
    }

    log -text "Login attempt using native method at tenant $O365CustomerName"
    $uidEnc = [System.Web.HttpUtility]::UrlEncode($userUPN)
    #stel allereerste cookie in om websessie te beginnen
    try{
        if($adfsSmartLink){
            log -text "Using ADFS Smartlink"
            $res = New-WebRequest -url $adfsSmartLink -Method Get
            $mode = "Federated"
        }else{
            $res = New-WebRequest -url https://login.microsoftonline.com -Method Get -contentType ""
            $stsRequest = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"ctx`" value=`""
            if($stsRequest -eq -1){ #we're seeing a new forward method, so we should find the login URL Microsoft suggests
                log -text "no STS request detected in response, checking for urlLogin parameter..."
                $urlLogin = returnEnclosedFormValue -res $res -searchString "`"urlLogin`":`""
                try{
                    if($urlLogin.StartsWith("https://")){
                        log -text "urlLogin parameter found, following once...."
                        $res = New-WebRequest -url $urlLogin -Method GET            
                    }
                }catch{$Null}
            }
            $apiCanary = returnEnclosedFormValue -res $res -searchString "`"apiCanary`":`""
            $clientId = returnEnclosedFormValue -res $res -searchString "correlationId`":`""
            #vind session code en gebruik deze om het realm van de user te vinden
            $stsRequest = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"ctx`" value=`""
            $cstsRequest = returnEnclosedFormValue -res $res -searchString "`",`"sCtx`":`""
            $flowToken = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"flowToken`" value=`""
            $sFT = returnEnclosedFormValue -res $res -searchString "`"sFT`":`""
            $canary = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"canary`" value=`""
            $newCanary = returnEnclosedFormValue -res $res -searchString "`"canary`":`""
            #URL encode the canary for POST request
            $canary = [System.Web.HttpUtility]::UrlEncode($canary)
            $newCanary = [System.Web.HttpUtility]::UrlEncode($newCanary)
            #this is the new realm discovery endpoint:
            $customHeaders = @{"canary" = $apiCanary;"hpgid" = "1104";"hpgact" = "1800";"client-request-id"=$clientId}
            $JSON = @{"username"="$userUPN";"isOtherIdpSupported"=$true;"checkPhones"=$false;"isRemoteNGCSupported"=$false;"isCookieBannerShown"=$false;"isFidoSupported"=$false;"originalRequest"="$cstsRequest"}
            $JSON = ConvertTo-Json $JSON
            try{
                $res = New-WebRequest -url "https://login.microsoftonline.com/common/GetCredentialType" -Method POST -body $JSON -customHeaders $customHeaders -referer $res.rawResponse.ResponseUri.AbsoluteUri
                log -text "New realm discovery method succeeded"
            }catch{
                log -text "New realm discovery method failed" -warning
                $res = New-WebRequest -url "https://login.microsoftonline.com/common/userrealm?user=$uidEnc&api-version=2.1&stsRequest=$stsRequest&checkForMicrosoftAccount=false" -Method GET
                log -text "Old realm discovery method succeeded"
            }
        }
    }catch{
        log -text "Unable to find user realm due to $($Error[0])" -fout
        return $False
    }

    if(!$adfsSmartLink){
        $jsonRealmConfig = ConvertFrom-Json $res.Content
        $mode = $Null
        if($jsonRealmConfig.Credentials.FederationRedirectUrl){
            $mode = "Federated"
            $nextURL = $jsonRealmConfig.Credentials.FederationRedirectUrl
            log -text "Received API response for authentication method: Federated (new style)"
            log -text "Authentication target: $nextURL"
        }elseif($jsonRealmConfig.NameSpaceType -eq "Federated"){
            $mode = "Federated"
            $nextURL = $jsonRealmConfig.AuthURL
            log -text "Received API response for authentication method: Federated"
            log -text "Authentication target: $nextURL"
        }elseif($jsonRealmConfig.NameSpaceType -eq "Managed"){
            $mode = "Managed"
            log -text "Received API response for authentication method: Managed"
            if($jsonRealmConfig.is_dsso_enabled){
                $apiCanary = $jsonRealmConfig.apiCanary
                $azureADSSOEnabled = $True
                log -text "Additionally, Azure AD SSO and/or PassThrough is enabled for your tenant"
            }
            $nextURL = "https://login.microsoftonline.com/common/login"            
        }else{
            $mode = "New_Managed"
            log -text "Received API response for authentication method: Managed, new style"
            if($jsonRealmConfig.EstsProperties.DesktopSsoEnabled){
                $azureADSSOEnabled = $True
                log -text "Additionally, Azure AD SSO and/or PassThrough is enabled for your tenant"                
            }
        }
    }

    #authenticate using New Managed Mode
    if($mode -eq "New_Managed"){
       #if azure AD SSO is enable, we need to trigger a session with the backend
        if($azureADSSOEnabled){
            $nextURL2 = "https://autologon.microsoftazuread-sso.com/$($userUPN.Split("@")[1])/winauth/sso?desktopsso=true&isAdalRequest=False&client-request-id=$clientId"
            log -text "Authentication target: $nextURL2"
            try{
                $res = New-WebRequest -url $nextURL2 -trySSO 1 -method GET -accept "text/html, application/xhtml+xml, image/jxr, */*" -referer "https://login.microsoftonline.com/" 
                log -text "Azure AD SSO response received: $($res.Content)"
                $ssoToken = $res.Content
            }catch{
                log -text "no SSO token received from AzureAD, did you add autologon.microsoftazuread-sso.com to the local intranet sites?" -warning
            }
            $nextURL2 = "https://login.microsoftonline.com/common/instrumentation/dssostatus"
            $customHeaders = @{"canary" = $apiCanary;"hpgid" = "1104";"hpgact" = "1800";"client-request-id"=$clientId}
            $JSON = @{"resultCode"="0";"ssoDelay"="200";"log"=$Null}
            $JSON = ConvertTo-Json $JSON
            $res = New-WebRequest -url $nextURL2 -method POST -body $JSON -customHeaders $customHeaders
            $JSON = ConvertFrom-Json $res.Content
            if($JSON.apiCanary){
                log -text "AADC SSO step 1 completed"
            }else{
                log -text "Failed to retrieve AADC SSO status token" -fout
            }

        }
        $attempts = 0
        while($true){
            if($attempts -gt 2){
                log -text "Failed to log you in with the supplied credentials after 3 attempts, aborting" -fout
                return $False
            }
            if($attempts -gt 0){
                #we didn't get logged in
                log -text "Failed to log you in with the supplied password, asking for (new) password" -fout
                $password = retrievePassword -forceNewPassword
            }else{
                if(!$azureADSSOEnabled -or ($azureADSSOEnabled -and $ssoToken.Length -lt 10)){
                    log -text "Managed authentication requires a password, retrieving it now..."
                    $password = retrievePassword
                }
            }
            $passwordEnc = [System.Web.HttpUtility]::UrlEncode($password)
            try{
                if($azureADSSOEnabled -and $ssoToken.Length -gt 10){
                    $body = "login=$userUPN&passwd=&ctx=$cstsRequest&flowToken=$sFT&canary=$newCanary&dssoToken=$ssoToken"
                }else{
                    $body = "i13=0&login=$userUPN&loginfmt=$userUPN&type=11&LoginOptions=3&passwd=$passwordEnc&ps=2&canary=$newCanary&ctx=$cstsRequest&flowToken=$sFT&NewUser=1&fspost=0&i21=0&CookieDisclosure=0&i2=1&i19=41303"
                }
                log -text "authenticating using new managed mode as $userUPN"
                $res = New-WebRequest -url "https://login.microsoftonline.com/common/login" -Method POST -body $body      
            }catch{
                log -text "error received while posting to login page" -fout
            }

            #DUO MFA check
            try{
                log -text "Checking for DUO MFA...."
                $res = handleDuoMFA -res $res -clientId $clientId
                log -text "DUO MFA checked"
            }catch{
                log -text "DUO MFA check result: $_" -fout
            }

            #MFA check
            try{
                $res = handleMFArequest -res $res -clientId $clientId
                log -text "MFA challenge completed"
            }catch{
                log -text "MFA check result: $_"
            }

            if($res.rawResponse.ResponseUri.AbsoluteUri.StartsWith("https://login.microsoftonline.com")){
                #still at login page
                if($res.Content.IndexOf("<meta name=`"PageID`" content=`"KmsiInterrupt`"") -ne -1){
                    #we're at the KMSI prompt, let's handle that
                    log -text "KMSI prompt detected"
                    $cstsRequest = returnEnclosedFormValue -res $res -searchString "`",`"sCtx`":`""
                    $cstsRequest = [System.Web.HttpUtility]::UrlEncode($cstsRequest)
                    $sFT = returnEnclosedFormValue -res $res -searchString "`",`"sFT`":`""
                    $sFT = [System.Web.HttpUtility]::UrlEncode($sFT)
                    $newCanary = returnEnclosedFormValue -res $res -searchString "`",`"apiCanary`":`""
                    $newCanary = [System.Web.HttpUtility]::UrlEncode($newCanary)
                    $body = "LoginOptions=1&ctx=$cstsRequest&flowToken=$sFT&canary=$newCanary&DontShowAgain=true&i19=2759"
                    try{
                        log -text "sending cookie persistence request"
                        $res = New-WebRequest -url "https://login.microsoftonline.com/kmsi" -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri        
                    }catch{$Null}
                }
            }

            if($res.rawResponse.ResponseUri.AbsoluteUri.StartsWith("https://device.login.microsoftonline.com")){
                log -text "Deviceauth encountered, breaking auth loop"
                break
            }

            $code = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"code`" value=`""
            if($res.rawResponse.ResponseUri.AbsoluteUri.StartsWith("https://login.microsoftonline.com") -and $code -ne -1){
                #we reached a landing page, but there is a redirect left, which is handled later in this script
                log -text "Redirect detected..."
                break
            }else{
                log -text "Failure during attempt $attempts, fiddler logs may be required" -fout
            }
            $attempts++
        }
    }

    #authenticate using Managed Mode
    if($mode -eq "Managed"){
        $attempts = 0
        #if azure AD SSO is enable, we need to trigger a session with the backend
        if($azureADSSOEnabled){
            $nextURL2 = "https://autologon.microsoftazuread-sso.com/$($jsonRealmConfig.DomainName)/winauth/sso?desktopsso=true&isAdalRequest=False&client-request-id=$clientId"
            log -text "Authentication target: $nextURL2"
            try{
                $res = New-WebRequest -url $nextURL2 -trySSO 1 -method GET -accept "text/html, application/xhtml+xml, image/jxr, */*" -referer "https://login.microsoftonline.com/" 
                log -text "Azure AD SSO response received: $($res.Content)"
                $ssoToken = $res.Content
            }catch{
                log -text "no SSO token received from AzureAD, did you add autologon.microsoftazuread-sso.com to the local intranet sites?" -warning
            }
            $nextURL2 = "https://login.microsoftonline.com/common/instrumentation/dssostatus"
            $customHeaders = @{"canary" = $apiCanary;"hpgid" = "1002";"hpgact" = "2101";"client-request-id"=$clientId}
            $JSON = @{"resultCode"="107";"ssoDelay"="200";"log"=$Null}
            $JSON = ConvertTo-Json $JSON
            $res = New-WebRequest -url $nextURL2 -method POST -body $JSON -customHeaders $customHeaders
            $JSON = ConvertFrom-Json $res.Content
            if($JSON.apiCanary){
                log -text "AADC SSO step 1 completed"
            }else{
                log -text "Failed to retrieve AADC SSO status token" -fout
            }

        }
        while($true){
            if($attempts -gt 2){
                log -text "Failed to log you in with the supplied credentials after 3 attempts, aborting" -fout
                return $False
            }
            if($attempts -gt 0){
                $loginFormPos = returnEnclosedFormValue -res $res -searchString "<form id=`"credentials`" method=`"post`" action=`"/common/login`">"
                if($loginFormPos -ne -1){
                    #we didn't get logged in
                    log -text "Failed to log you in with the supplied password, asking for (new) password" -fout
                    $password = retrievePassword -forceNewPassword
                    $passwordEnc = [System.Web.HttpUtility]::UrlEncode($password)
                }else{
                    break
                }
            }else{
                if(!$azureADSSOEnabled -or ($azureADSSOEnabled -and $ssoToken.Length -lt 10)){
                    log -text "Managed authentication requires a password, retrieving it now..."
                    $password = retrievePassword
                    $passwordEnc = [System.Web.HttpUtility]::UrlEncode($password)
                }
            }
            if($azureADSSOEnabled -and $ssoToken.Length -gt 10){
                $body = "login=$userUPN&passwd=&ctx=$stsRequest&flowToken=$flowToken&canary=$canary&dssoToken=$ssoToken"
            }else{
                $body = "login=$userUPN&passwd=$passwordEnc&ctx=$stsRequest&flowToken=$flowToken&canary=$canary"
            }
            log -text "Requesting session..."
            try{
                $res = New-WebRequest -url $nextURL -Method POST -body $body
            }catch{
                log -text "failure at logon attempt: $($Error[0])" -fout
                continue
            }
            $errorDetection = returnEnclosedFormValue -res $res -searchString "<td id=`"service_exception_message`"" -endString "</td>"
            if($errorDetection -ne -1){
                log -text "Possible error detected at signin page: $errorDetection" -fout
                if($errorDetection.IndexOf("AADSTS165000") -ne -1){
                    log -text "This is a know issue, try again in a few minutes" -fout
                    return $False
                }
            }
            if($azureADSSOEnabled){
                try{
                    log -text "Checking if we've been Signed in automatically by Azure AD PassThrough..."
                    $res2 = checkIfAtFhmPage -res $res
                    log -text "SSO completed"
                    return $True
                }catch{
                    log -text "We do not seem to have been properly redirected for SSO yet" -warning
                }
                try{
                    $stsRequest = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"ctx`" value=`""
                    if($stsRequest -eq -1){
                        Throw "No sts request found in response"
                    }else{
                        log -text "New sts request retrieved"
                    }
                    $flowToken = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"flowToken`" value=`""
                    if($flowToken -eq -1){
                        Throw "No flowToken found in response"
                    }else{
                        log -text "New flowToken retrieved"
                    }                    
                    $canary = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"canary`" value=`""
                    if($canary -eq -1){
                        Throw "No canary found in response"
                    }else{
                        log -text "New canary retrieved"
                    } 
                    $apiCanary = returnEnclosedFormValue -res $res -searchString "`"apiCanary`":`""
                    if($apiCanary -eq -1){
                        Throw "No apiCanary found in response"
                    }else{
                        log -text "New apiCanary retrieved"
                    }  
                    $customHeaders = @{"canary" = $apiCanary;"hpgid" = "1002";"hpgact" = "2000";"client-request-id"=$clientId}
                    $nextURL = "https://login.microsoftonline.com/common/onpremvalidation/Poll"
                    $JSON = @{"flowToken"=$flowToken;"ctx"=$stsRequest}
                    $JSON = ConvertTo-Json $JSON
                    $res = New-WebRequest -url $nextURL -Method POST -body $JSON -customHeaders $customHeaders
                    $response = ConvertFrom-Json $res.Content
                    $body = "flowToken=$($response.flowToken)&ctx=$stsRequest"
                    $nextURL =  "https://login.microsoftonline.com/common/onpremvalidation/End"
                    $res = New-WebRequest -url $nextURL -Method POST -body $body
                    log -text "AADC SSO step 2 completed"
                }catch{
                    log -text "Error trying to request AzureAD SSO token: $($Error[0])" -fout
                }
                try{
                    log -text "Checking if we've been Signed in by Azure AD"
                    $res2 = checkIfAtFhmPage -res $res
                    log -text "SSO completed"
                    return $True
                }catch{
                    log -text "We do not seem to have been properly redirected for SSO yet" -warning
                }
                try{
                    $JSON = ConvertFrom-Json $res.Content
                    if($JSON.error.message -eq "AADSTS50012"){
                        log -text "Your password was incorrect according to AADC SSO" -fout
                    }elseif($JSON.error){
                        log -text "There was an error at AzureAD Connect SSO: $($JSON.error.message)" -fout
                        return $False
                    }
                }catch{$Null}
            }
            $attempts++
        }
    }

    #authenticate using Federated Mode
    if($mode -eq "Federated"){
        log -text "Contacting Federation server and attempting Single SignOn..."
        if(!$adfsSmartLink){
            try{
                $res = New-WebRequest -url $nextURL -Method GET -referer $res.rawResponse.ResponseUri.AbsoluteUri
            }catch{
                log -text "Error received from ADFS server: $($Error[0])" -fout
                return $False
            }
        }

        #check if there is an F5 in between and handle accordingly
        $SAMLRequest = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"SAMLRequest`" value=`""
        if($SAMLRequest -ne -1){
            $nextURL = returnEnclosedFormValue -res $res -searchString "<form method=`"POST`" name=`"hiddenform`" action=`""
            log -text "Detected F5 SAML Request, forwarding it to $nextURL"
            $SAMLRequest = [System.Web.HttpUtility]::HtmlDecode($SAMLRequest)
            $SAMLRequest = [System.Web.HttpUtility]::UrlEncode($SAMLRequest)
            $RelayState = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"RelayState`" value=`""
            $body = "SAMLRequest=$SAMLRequest&RelayState=$RelayState"
            try{
                $res = New-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri
            }catch{
                log -text "Error received from F5 server: $($Error[0])" -fout
                return $False
            }
            $nextURL = returnEnclosedFormValue -res $res -searchString "<form method=`"POST`" action=`""        
            $nextURL = "https://$($res.rawResponse.ResponseUri.Host)$nextURL"
            $dummy = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"dummy`" value=`""
            $body = "dummy=$dummy"
            try{
                if($dummy -eq -1){
                    Throw "No redirect code detected from F5, but this may not be fatal"
                }else{
                    log -text "Retrieved forward code from F5 server, forwarding to $nextURL"
                }
                $res = New-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri
            }catch{
                log -text "Error received from F5 server: $($Error[0])" -fout
            }

            $nextURL = returnEnclosedFormValue -res $res -searchString "<form action=`""
            $SAMLResponse = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"SAMLResponse`" value=`""
            $SAMLResponse = [System.Web.HttpUtility]::HtmlDecode($SAMLResponse)
            $SAMLResponse = [System.Web.HttpUtility]::UrlEncode($SAMLResponse)
            $RelayState = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"RelayState`" value=`""
            $body = "SAMLResponse=$SAMLResponse&RelayState=$RelayState"
            try{
                if($nextURL -ne -1){
                    log -text "SAML response retrieved from F5, sending to endpoint..."
                }else{
                    Throw "No (readable) SAML response retrieved from F5! Use Fiddler to debug"
                }               
                $res = New-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri
            }catch{
                log -text "Error received when getting SAML response from F5 server, script will likely fail: $($Error[0])" -fout
            }
        }
        #\END F5 logic

        ##if we get a SAML token, we've been signed in automatically, otherwise, we will have to post our credentials
        $wResult = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"wresult`" value=`""
        if($wResult -eq -1){
            log -text "Federation Services did not sign us in automatically, retrieving user credentials.." -warning
            $password = retrievePassword
            $passwordEnc = [System.Web.HttpUtility]::HtmlEncode($password)
            $ADFShost = $res.rawResponse.ResponseUri.host
            $nextURL = returnEnclosedFormValue -res $res -searchString "<form method=`"post`" id=`"loginForm`" autocomplete=`"off`" novalidate=`"novalidate`" onKeyPress=`"if (event && event.keyCode == 13) Login.submitLoginRequest();`" action=`"/" -decode
            if($nextURL.IndexOf("https:") -eq -1){
                $nextURL = "https://$($ADFShost)/$($nextURL)"
            }
            $userName = $userUPN
            $attempts = 0
            while($true){
                if($attempts -gt 2){
                    log -text "Failed to log you in with the supplied credentials after 3 attempts, aborting" -fout
                    return $False
                }
                if($attempts -gt 0){
                    $wResult = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"wresult`" value=`""
                    if($wResult -eq -1){
                        #we didn't get logged in
                        log -text "Failed to log you in with the supplied password, asking for (new) password" -fout
                        $password = retrievePassword -forceNewPassword
                        $passwordEnc = [System.Web.HttpUtility]::UrlEncode($password)
                    }else{
                        break
                    }
                }
                #handle ESET plugin
                if($res.Content.IndexOf("esa_message_push") -ne -1){
                    $nextURL = returnEnclosedFormValue -res $res -searchString "<form id=`"options`"  method=`"post`" action=`""
                    $context = returnEnclosedFormValue -res $res -searchString "<input id=`"context`" type=`"hidden`" name=`"Context`" value=`""
                    $body = "AuthMethod=esa_adfs_adapter&Context=$context&submit_auto="
                }else{
                    $body = "UserName=$userName&Password=$passwordEnc&Kmsi=true&AuthMethod=FormsAuthentication"
                }
                $res = New-WebRequest -url $nextURL -Method POST -body $body
                $attempts++
            }

        }
        $nextURL = returnEnclosedFormValue -res $res -searchString "<form method=`"POST`" name=`"hiddenform`" action=`""
        $wResult = [System.Web.HttpUtility]::HtmlDecode($wResult)
        $wResult = [System.Web.HttpUtility]::UrlEncode($wResult)
        $body = "wa=wsignin1.0&wresult=$wResult&amp;LoginOptions=1"
        $res = New-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/x-www-form-urlencoded" -accept "text/html, application/xhtml+xml, image/jxr, */*"
        
        #check for double redirect which will happen if ADFS is itself federated
        $wResult = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"wresult`" value=`""
		if($wResult -ne -1){
			$nextURL = returnEnclosedFormValue -res $res -searchString "<form method=`"POST`" name=`"hiddenform`" action=`""
			log -text "Federation Services has a second wresult step.." -warning
			$wResult = [System.Web.HttpUtility]::HtmlDecode($wResult)
			$wResult = [System.Web.HttpUtility]::UrlEncode($wResult)
			$body = "wa=wsignin1.0&wresult=$wResult&amp;LoginOptions=1"
			$res = New-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/x-www-form-urlencoded" -accept "text/html, application/xhtml+xml, image/jxr, */*" 
		}

        if($res.Content.IndexOf("<meta name=`"PageID`" content=`"KmsiInterrupt`"") -ne -1){
            #we're at the KMSI prompt, let's handle that
            log -text "KMSI prompt detected"
            $cstsRequest = returnEnclosedFormValue -res $res -searchString "`",`"sCtx`":`""
            $cstsRequest = [System.Web.HttpUtility]::UrlEncode($cstsRequest)
            $sFT = returnEnclosedFormValue -res $res -searchString "`",`"sFT`":`""
            $sFT = [System.Web.HttpUtility]::UrlEncode($sFT)
            $newCanary = returnEnclosedFormValue -res $res -searchString "`",`"apiCanary`":`""
            $newCanary = [System.Web.HttpUtility]::UrlEncode($newCanary)
            $body = "LoginOptions=1&ctx=$cstsRequest&flowToken=$sFT&canary=$newCanary&DontShowAgain=true&i19=2759"
            try{
                log -text "sending cookie persistence request"
                $res = New-WebRequest -url "https://login.microsoftonline.com/kmsi" -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri        
            }catch{$Null}
        }

    }

    ##AT this point, authentication should have succeeded, but redirects need to be followed and may differ per type of tenant

    #some customers have a redirect active to Onedrive for Business, check if we're already there and return true if so
    if($res.rawResponse.ResponseUri.OriginalString.IndexOf("/personal/") -ne -1){
        log -text "Logged into Office 365! Already redirected to Onedrive for Business"
        return $True
    }

    #follow first 1-2 redirects, fail if none are detected or if redirects are detected but fail (abnormal flow)
    try{
        $res = handleO365Redirect -res $res
        if($res[1] -eq $False){
            return $False
        }
    }catch{
        return $False
    }

    #MFA check
    try{
        $res = handleMFArequest -res $res[0] -clientId $clientId
        log -text "MFA challenge completed"
    }catch{
        log -text "MFA check result: $_"
    }    

    #sometimes additional redirects are needed, fail if redirects fail, succeed if none are detected or if they are followed
    try{
        $res = handleO365Redirect -res $res
    }catch{
        return $False
    }    

    #MFA check
    try{
        $res = handleMFArequest -res $res[0] -clientId $clientId
        log -text "MFA challenge completed"
    }catch{
        log -text "MFA check result: $_"
    }   
    
    return $True
}

#return -1 if nothing found, or value if found
function checkRegistryKeyValue{
    Param(
        [String]$basePath,
        [String]$entryName
    )
    try{$value = (Get-ItemProperty -Path "$($basePath)\" -Name $entryName -ErrorAction Stop).$entryName
        return $value
    }catch{
        return -1
    }
}

function addSiteToIEZoneThroughRegistry{
    Param(
        [String]$siteUrl,
        [Int]$mode=2 #1=intranet, 2=trusted sites
    )
    try{
        $components = $siteUrl.Split(".")
        $count = $components.Count
        if($count -gt 3){
            $old = $components
            $components = @()
            $subDomainString = ""
            for($i=0;$i -le $count-3;$i++){
                if($i -lt $count-3){$subDomainString += "$($old[$i])."}else{$subDomainString += "$($old[$i])"}
            }
            $components += $subDomainString
            $components += $old[$count-2]
            $components += $old[$count-1]    
        } 
        if($count -gt 2){
            $res = New-Item "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[1]).$($components[2])" -ErrorAction SilentlyContinue 
            $res = New-Item "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[1]).$($components[2])\$($components[0])" -ErrorAction SilentlyContinue
            $res = New-ItemProperty "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[1]).$($components[2])\$($components[0])" -Name "https" -value $mode -ErrorAction Stop
        }else{
            $res = New-Item "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[0]).$($components[1])" -ErrorAction SilentlyContinue 
            $res = New-ItemProperty "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[0]).$($components[1])" -Name "https" -value $mode -ErrorAction Stop
        }
    }catch{
        return -1
    }
    return $True
}

function getUserLogin{
    Param(
        [Switch]$reType
    )
    #get user login 
    if($forceUserName.Length -gt 2){ 
        log -text "A username was already specified in the script configuration: $($forceUserName)" 
        $userUPN = $forceUserName 
        $userLookupMode = 0
    }else{
        switch($userLookupMode){
            1 {    
                log -text "userLookupMode is set to 1 -> checking Active Directory UPN" 
                try{
                    $userUPN = (lookupLoginFromAD).ToLower()
                    $Null = retrieveLogin -cacheLogin $userUPN
                }catch{
                    $userUPN = retrieveLogin
                }
            }
            2 {
                log -text "userLookupMode is set to 2 -> checking Active Directory email address" 
                try{    
                    $userUPN = (lookupLoginFromAD -lookupEmail).ToLower()  
                    $Null = retrieveLogin -cacheLogin $userUPN
                }catch{
                    $userUPN = retrieveLogin
                }
            }
            3 {
            #Windows 10
                 try{
                    log -text "userLookupMode is set to 3, using SID discovery method"
                    $objUser = New-Object System.Security.Principal.NTAccount($Env:USERNAME)
                    $strSID = ($objUser.Translate([System.Security.Principal.SecurityIdentifier])).Value
                    $basePath = "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$strSID\IdentityCache\$strSID"
                    if((test-path $basePath) -eq $False){
                        log -text "userLookupMode is set to 3, but the registry path $basePath does not exist! All lookup modes exhausted, exiting" -fout
                        abort_OM   
                    }
                    $userId = (Get-ItemProperty -Path $basePath -Name UserName).UserName
                    if($userId -and $userId -like "*@*"){
                        log -text "userLookupMode is set to 3, we detected $userId in $basePath"
                        $userUPN = ($userId).ToLower()
                    }else{
                        log -text "userLookupMode is set to 3, but we failed to detect a username at $basePath" -fout
                        abort_OM
                    }
                 }catch{
                    log -text "userLookupMode is set to 3, but we failed to detect a proper username" -fout
                    abort_OM
                 }
            }
            4 {
                if($reType){
                    $userUPN = (retrieveLogin -forceNewLogin)
                }else{
                    $userUPN = (retrieveLogin)
                }
            }
            5 {
                try{
                    if((test-path $userLoginRegistryKey)){
                        $userId = (Get-ItemProperty -Path $userLoginRegistryKey -Name Office365Login).Office365Login   
                    }else{
                        Throw "$userLoginRegistryKey path does not exist"
                    }
                }catch{
                    log -text "failed to detect username in $userLoginRegistryKey path due to $($Error[0])"
                    abort_OM
                } 
                if($userId -and $userId -like "*@*"){
                    log -text "userLookupMode is set to 5, we detected $userId in $userLoginRegistryKey"
                    $userUPN = ($userId).ToLower()
                }else{
                    log -text "userLookupMode is set to 5, but we failed to detect a username at $userLoginRegistryKey" -fout
                    abort_OM
                }
            }
            6 {
                log -text "userLookupMode is set to 6"
                $login = retrieveLogin -noQuery
                $password = retrievePassword -noQuery
                if($password -ne -1 -and $login -ne -1){
                    log -text "$login and matching password detected in cache, no need to query user"
                    $userUPN = $login
                }else{
                    if($login -ne -1){
                        $script:userUPN = $login
                    }
                    try{
                        $res = queryForAllCreds -titleText $loginformTitleText -introText $loginformIntroText -buttonText $buttonText -loginLabel $loginFieldText -passwordLabel $passwordFieldText
                        if($res[0]){
                            log -text "User login entered: $($res[0]), storing..."
                            $rez = retrieveLogin -cacheLogin $res[0]
                            $userUPN = $res[0]
                        }else{
                            log -text "User did not enter a login, aborting"
                            abort_OM
                        }
                        if($res[1]){
                            log -text "User password entered, storing..."
                            $rez = retrievePassword -cachePassword $res[1]
                        }else{
                            log -text "User did not enter a password, aborting"
                            abort_OM
                        }
                    }catch{
                        log -text "Failed to query user for credentials: $($Error[0])" -fout
                        abort_OM
                    }
                }
            }
            7 {
                log -text "userLookupMode is set to 7 -> getting user from whoami /upn"
                $userUPN = whoami /upn
            }
            default {
                log -text "userLookupMode is not properly configured" -fout
                abort_OM
            }
        }
    }
    return $userUPN
}

function checkWebClient{
    if((Get-Service -Name WebClient).Status -ne "Running"){ 
        #attempt to auto-start if user is admin
        if($isElevated){
            Start-Service WebClient -ErrorAction SilentlyContinue | Out-Null
        }else{
            #use another trick to autostart the client
            try{
                startWebDavClient
            }catch{
                log -text "CRITICAL ERROR: OneDriveMapper detected that the WebClient service was not started, please ensure this service is always running!`n" -fout
                $script:errorsForUser += "$MD_DriveLetter could not be mapped because the WebClient service is not running`n"
            }
        }
    } 
}

#check if the script is running elevated, run via scheduled task if UAC is not disabled
If (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){   
    log -text "Script elevation level: Administrator" -fout
    $scheduleTask = $True
    $isElevated = $True
    if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -entryName "EnableLUA") -eq 0){
        log -text "NOTICE: $($BaseKeypath)\EnableLua found in registry and set to 0, you have disabled UAC, the script does not need to bypass by using a scheduled task"    
        $scheduleTask = $False                
    }    
    if($asTask){
        log -text "Already running as task, but still elevated, will attempt to map normally but drives may not be visible" -fout
        $scheduleTask = $False
    }
    checkWebClient
    if($scheduleTask){
        $Null = fixElevationVisibility
        Exit
    }
}else{
    log -text "Script elevation level: User"
    $isElevated = $False
    checkWebClient
}

#load windows libraries that we require
try{ 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
}catch{ 
    log -text "Error loading windows forms libraries, script will not be able to display a password input box" -fout
} 

$WebAssemblyloaded = $True
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Web")
if(-NOT [appdomain]::currentdomain.getassemblies() -match "System.Web"){
    log -text "Error loading System.Web library to decode sharepoint URL's, mapped sharepoint URL's may become read-only. $($Error[0])" -fout
    $WebAssemblyloaded = $False
}

#try to set TLS to v1.2, Powershell defaults to v1.0
try{
    $res = [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    log -text "Set TLS protocol version to prefer v1.2"
}catch{
    log -text "Failed to set TLS protocol to perfer v1.2 $($Error[0])" -fout
}

#get OSVersion
$windowsVersion = ([System.Environment]::OSVersion.Version).Major
$objUser = New-Object System.Security.Principal.NTAccount($Env:USERNAME)
$strSID = ($objUser.Translate([System.Security.Principal.SecurityIdentifier])).Value
log -text "You are $strSID running on Windows $windowsVersion and Powershell version $($PSVersionTable.PSVersion.Major)"

if($showConsoleOutput -eq $False){
    $t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    try{
        add-type -name win -member $t -namespace native
        [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
    }catch{$Null}
}

if($PSVersionTable.PSVersion.Major -le 2){
    log -text "ERROR: you're trying to use Native auth on Powershell V2 or lower" -fout
}

#show a progress bar if set to True
if($showProgressBar) {
    #title for the winform
    $Title = "OnedriveMapper v$version"
    #winform dimensions
    $height=39
    $width=400
    #winform background color
    $color = "White"

    #create the form
    $form1 = New-Object System.Windows.Forms.Form
    $form1.Text = $title
    $form1.Height = $height
    $form1.Width = $width
    $form1.BackColor = $color
    $form1.ControlBox = $false
    $form1.MaximumSize = New-Object System.Drawing.Size($width,$height)
    $form1.MinimumSize = new-Object System.Drawing.Size($width,$height)
    $form1.Size = new-Object System.Drawing.Size($width,$height)

    $form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None 
    #display center screen
    $form1.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $screen = ([System.Windows.Forms.Screen]::AllScreens | Where-Object {$_.Primary}).WorkingArea
    $form1.Location = New-Object System.Drawing.Size(($screen.Right - $width), ($screen.Bottom - $height))
    $form1.Topmost = $True 
    $form1.TopLevel = $True 

    # create label
    $script:label1 = New-Object system.Windows.Forms.Label
    $script:label1.text=$script:progressBarText
    $script:label1.Name = "label1"
    $script:label1.Left=0
    $script:label1.Top= 9
    $script:label1.Width= $width
    $script:label1.Height=17
    $script:label1.Font= "Verdana"
    # create label
    $label2 = New-Object system.Windows.Forms.Label
    $label2.Name = "label2"
    $label2.Left=0
    $label2.Top= 0
    $label2.Width= $width
    $label2.Height=7
    $label2.backColor= $progressBarColor

    #add the label to the form
    $form1.controls.add($script:label1) 
    $form1.controls.add($label2) 
    $script:progressBar1 = New-Object System.Windows.Forms.ProgressBar
    $script:progressBar1.Name = 'progressBar1'
    $script:progressBar1.Value = 0
    $script:progressBar1.Style="Continuous" 
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = $width
    $System_Drawing_Size.Height = 10
    $progressBar1.Size = $System_Drawing_Size   
    
    $script:progressBar1.Left = 0
    $script:progressBar1.Top = 29
    $form1.Controls.Add($script:progressBar1)
    $form1.Show()| out-null  
    $form1.Focus() | out-null 
    $script:progressbar1.Value = 5
    $form1.Refresh()
}

#load cookie code and test-set a cookie

log -text "Loading CookieSetter..."
$source=@"
using System.Runtime.InteropServices;
using System;
namespace Cookies
{
    public static class setter
    {
        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool InternetSetCookie(string url, string name, string data);

        public static bool SetWinINETCookieString(string url, string name, string data)
        {
            bool res = setter.InternetSetCookie(url, name, data);
            if (!res)
            {
                throw new Exception("Exception setting cookie: Win32 Error code="+Marshal.GetLastWin32Error());
            }else{
                return res;
            }
        }
    }
}
"@
try{
    $compilerParameters = New-Object System.CodeDom.Compiler.CompilerParameters
    $compilerParameters.CompilerOptions="/unsafe"
    $compilerParameters.GenerateInMemory = $True
    Add-Type -TypeDefinition $source -Language CSharp -CompilerParameters $compilerParameters
    [DateTime]$dateTime = Get-Date
    $dateTime = $dateTime.AddDays(1)
    $str = $dateTime.ToString("R")
    $res = [Cookies.setter]::SetWinINETCookieString("https://testdomainthatdoesnotexist.com","none","Data=nothing;Expires=$str")
    log -text "Test cookie set successfully"
}catch{
    log -text "ERROR: Failed to set test cookie, script will fail: $($Error[0])" -fout
}

#Check if Zone Configuration is on a per machine or per user basis, then check the zones 
$privateZoneFound = $False
$publicZoneFound = $False

#check if zone enforcement is set to machine only
$reg_HKLM = checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings" -entryName "Security HKLM only"
if($reg_HKLM -eq -1){
    log -text "NOTICE: HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Security HKLM only not found in registry, your zone configuration could be set on both levels" 
}elseif($reg_HKLM -eq 1){
    log -text "NOTICE: HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Security HKLM only found in registry and set to 1, your zone configuration is set on a machine level"    
}

#check if sharepoint and onedrive are set as safe sites at the user level
if($reg_HKLM -ne 1){
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -match '[1-2]' -or (checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "*") -match '[1-2]'){
        log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level"  
        $privateZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -match '[1-2]' -or (checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "*") -match '[1-2]'){
        log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level (through GPO)"  
        $privateZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -match '[1-2]' -or (checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "*") -match '[1-2]'){
        log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on user level"  
        $publicZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -match '[1-2]' -or (checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "*") -match '[1-2]'){
        log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on user level (through GPO)" 
        $publicZoneFound = $True        
    }
}

#check if sharepoint and onedrive are set as safe sites at the machine level
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -match '[1-2]' -or (checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "*") -match '[1-2]'){
    log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level"
    $privateZoneFound = $True 
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -match '[1-2]' -or (checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "*") -match '[1-2]'){
    log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level (through GPO)"  
    $privateZoneFound = $True        
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -match '[1-2]' -or (checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "*") -match '[1-2]'){
    log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on machine level"  
    $publicZoneFound = $True    
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -match '[1-2]' -or (checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "*") -match '[1-2]'){
    log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on machine level (through GPO)"  
    $publicZoneFound = $True    
}

#add an entry to prevent file copy paste warnings
try{
    $res = New-Item "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\sharepoint.com@SSL" -ErrorAction SilentlyContinue 
    $res = New-Item "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\sharepoint.com@SSL\$($O365CustomerName)" -ErrorAction SilentlyContinue
    $res = New-ItemProperty "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\sharepoint.com@SSL\$($O365CustomerName)" -Name "file" -value 1 -ErrorAction SilentlyContinue
}catch{$Null}

#log results, try to automatically add trusted sites to user trusted sites if not yet added
if($publicZoneFound -eq $False){
    log -text "Possible critical error: $($O365CustomerName).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail" -fout
    if((addSiteToIEZoneThroughRegistry -siteUrl "$O365CustomerName.sharepoint.com") -eq $True){log -text "Automatically added $O365CustomerName.sharepoint.com to trusted sites for this user"}
}
if($privateZoneFound -eq $False){
    log -text "Possible critical error: $($O365CustomerName)$($privateSuffix).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail" -fout
    if((addSiteToIEZoneThroughRegistry -siteUrl "$($O365CustomerName)$($privateSuffix).sharepoint.com") -eq $True){log -text "Automatically added $($O365CustomerName)$($privateSuffix).sharepoint.com to trusted sites for this user"}
}

if($autoDetectProxy -eq $False){
    $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    $val = checkRegistryKeyValue -basePath $path -entryName "AutoDetect"
    if($val -eq 0){
        log -text "IE Automatic Proxy Detection is already disabled"
    }else{
        log -text "IE Automatic Proxy Detection is not yet disabled, attempting to disable..."
        try{
            $res = New-ItemProperty $path -Name "AutoDetect" -value 0 -ErrorAction Stop
            log -text "IE Automatic Proxy Detection disabled"    
        }catch{
            log -text "Failed to disable IE automatic proxy detection: $($Error[0])" -fout
        }
    }
}

$userUPN = getUserLogin

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 10
    $script:form1.Refresh()
}

#Check if webdav locking is enabled
if((checkRegistryKeyValue -basePath "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\" -entryName "SupportLocking") -ne 0){
    log -text "ERROR: WebDav File Locking support is enabled, this could cause files to become locked in your OneDrive or Sharepoint site" -fout 
} 

#report/warn file size limit
$sizeLimit = [Math]::Round((checkRegistryKeyValue -basePath "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\" -entryName "FileSizeLimitInBytes")/1024/1024)
log -text "Maximum file upload size is set to $sizeLimit MB" -warning

$baseURL = ("https://$($O365CustomerName)-my.sharepoint.com/_layouts/15/MySite.aspx?MySiteRedirect=AllDocuments") 
$mapURLpersonal = "\\$O365CustomerName-my.sharepoint.com@SSL\DavWWWRoot\personal\"

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 15
    $script:form1.Refresh()
}

for($count=0;$count -lt $desiredMappings.Count;$count++){
    #replace funky sharepoint URL stuff and turn into webdav path
    if($desiredMappings[$count].sourceLocationPath -ne "autodetect"){
        if($WebAssemblyloaded){
            $desiredMappings[$count].webDavPath = [System.Web.HttpUtility]::UrlDecode($desiredMappings[$count].sourceLocationPath)
        }
        $desiredMappings[$count].webDavPath = $desiredMappings[$count].webDavPath.Replace("https://","\\").Replace("/_layouts/15/start.aspx#","").Replace("sharepoint.com/","sharepoint.com@SSL\DavWWWRoot\").Replace("/Forms/AllItems.aspx","")
        $desiredMappings[$count].webDavPath = $desiredMappings[$count].webDavPath.Replace("/","\")  
    }else{
        $desiredMappings[$count].webDavPath = $mapURLpersonal
    }

    if($desiredMappings[$count].mapOnlyForSpecificGroup -and $groups){
        $group = $groups -contains $desiredMappings[$count].mapOnlyForSpecificGroup
        if($group){ 
            log -text "adding a sharepoint mapping because the user is a member of $group" 
            $desiredMappings[$count].alreadyMapped = $False 
        }else{
            $desiredMappings[$count].alreadyMapped = $True
        }
    }
}
 
#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 20
    $script:form1.Refresh()
}

log -text "Base URL: $($baseURL) `n" 

if($autoResetIE){
    & RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
}

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 30
    $script:form1.Refresh()
}

$res = loginV2
if($res -eq $False){
    log -text "native auth login mode failed, aborting script" -fout
    abort_OM
}else{
    log -text "Login succeeded"
}


#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 45
    $script:form1.Refresh()
}

#clean up any existing mappings
subst | % {subst $_.SubString(0,2) /D}
Get-PSDrive -PSProvider filesystem | Where-Object {$_.DisplayRoot} | % {
    if($_.DisplayRoot.StartsWith("\\$($O365CustomerName).sharepoint.com") -or $_.DisplayRoot.StartsWith("\\$($O365CustomerName)-my.sharepoint.com")){
        try{$del = NET USE "$($_.Name):" /DELETE /Y 2>&1}catch{$Null}     
    }
}

#clean up empty mappings
Get-PSDrive -PSProvider filesystem | Where-Object {($_.Used -eq 0 -and $_.Free -eq $Null)} | % {
    try{$_ | Remove-PSDRive -Force}catch{$Null}     
}

$suffixCounter = $Null #used in case converged mappings with the same name are detected
if($autoMapFavoriteSites){
    #get drives already in use
    $drvlist=(Get-PSDrive -PSProvider filesystem).Name
    #add already planned mappings to in use list
    foreach($mapping in $desiredMappings){
        if($mapping.targetLocationType -eq "driveletter"){
            if($drvlist -notcontains $($mapping.targetLocationPath.Substring(0,1))){
                $drvList += $($mapping.targetLocationPath.Substring(0,1))
            }
        }
    }
    if($autoMapFavoritesMode -eq "Converged"){ 
        #get first free driveletter for a converged fake mapping to contain all links
        if($drvlist -contains $autoMapFavoritesDrive){
            Foreach ($drvletter in $autoMapFavoritesDrvLetterList.ToCharArray()) {
                If ($drvlist -notcontains $drvletter) {
                    log -text "You set $autoMapFavoritesDrive as converged driveletter, but it is not available, using $drvletter instead" -warning
                    $drvlist += $drvletter
                    $autoMapFavoritesDriveletter = $drvletter
                    break
                }
            }
        }else{
            $drvlist += $autoMapFavoritesDrive
            $autoMapFavoritesDriveletter = $autoMapFavoritesDrive            
        }   
        $targetFolder = Join-Path $Env:TEMP -ChildPath "OnedriveMapperLinks" 
        if(![System.IO.Directory]::Exists($targetFolder)){
            log -text "Desired path for Team site links: $targetFolder does not exist, creating"
            try{
                $res = New-Item -Path $targetFolder -ItemType Directory -Force
            }catch{
                log -text "Failed to create folder $targetFolder! $($Error[0])" -fout
            }
        }else{
            try{
                Get-ChildItem $targetFolder | Remove-Item -Force -Confirm:$False -Recurse
            }catch{$Null}
        }
        $res = subst "$($autoMapFavoritesDriveletter):" $targetFolder
        labelDrive "$($autoMapFavoritesDriveLetter):" $autoMapFavoritesDriveLetter $autoMapFavoritesLabel
    }

    try{
        log -text "Retrieving favorited sites because autoMapFavoriteSites is set to TRUE"
        $res = New-WebRequest -url "https://$O365CustomerName.sharepoint.com/_layouts/15/sharepoint.aspx?v=following" -method GET
        $res = (handleO365Redirect -res $res)[0]
        $favoritesURL = "https://$O365CustomerName.sharepoint.com/_api/v2.1/favorites/followedSites?`$expand=contentTypes&`$top=100"
        $res = New-WebRequest -url $favoritesURL -method GET -accept "application/json;odata=verbose"
    }catch{
        log -text "error retrieving favorited sites $($Error[0])" -fout
    }
    try{
        $res = (handleO365Redirect -res $res)[0]
        $results = ($res.Content | convertfrom-json).value
    }catch{
        continue
    }   
    foreach($result in $results){
        $desiredUrl = $result.webUrl.Replace("https://","\\").Replace("/_layouts/15/start.aspx#","").Replace("sharepoint.com/","sharepoint.com@SSL\DavWWWRoot\").Replace("/Forms/AllItems.aspx","").Replace("/","\")
        if($autoMapFavoritesMode -eq "Normal"){
            Foreach ($drvletter in $autoMapFavoritesDrvLetterList.ToCharArray()) {
                If ($drvlist -notcontains $drvletter) {
                    $drvlist += $drvletter
                    break
                }
            }
                
            $desiredMappings +=   @{"displayName"=$($result.title);"targetLocationType"="driveletter";"targetLocationPath"="$($drvletter):";"sourceLocationPath" = $result.webUrl; "webDavPath"=$desiredUrl;"mapOnlyForSpecificGroup"="favoritesPlaceholder"}
            log -text "Adding $($result.webUrl) as $($result.title) to mapping list as drive $drvletter"
        }
        if(@($desiredMappings | Where-Object {$_.displayName -eq $result.title}).Count -gt 0){
            $suffixCounter++    
        }            
        if($autoMapFavoritesMode -eq "Onedrive"){
            [Array]$odMapping = @($desiredMappings | Where-Object{$_.sourceLocationPath -eq "autodetect"})
            if($odMapping.Count -le 0){
                log -text "you set automapFavoritesMode to Onedrive, but have not mapped Onedrive!" -fout
            }
            $path  = "$($odMapping[0].targetLocationPath)\$autoMapFavoritesLabel"
            $desiredMappings +=   @{"displayName"="$($result.title)$suffixCounter";"targetLocationType"="networklocation";"targetLocationPath"="$($path)";"sourceLocationPath" = $result.webUrl; "webDavPath"=$desiredUrl;"mapOnlyForSpecificGroup"="favoritesPlaceholder"}
            log -text "Adding $($result.webUrl) as $($result.title)$suffixCounter to mapping list as network shortcut in $path"
        } 
        if($autoMapFavoritesMode -eq "Converged"){ 
            $desiredMappings +=   @{"displayName"="$($result.title)$suffixCounter";"targetLocationType"="networklocation";"targetLocationPath"="$($autoMapFavoritesDriveletter):";"sourceLocationPath" = $result.webUrl; "webDavPath"=$desiredUrl;"mapOnlyForSpecificGroup"="favoritesPlaceholder"}
            log -text "Adding $($result.webUrl) as $($result.title)$suffixCounter to mapping list as network shortcut in a converged drive with letter $autoMapFavoritesDriveletter"               
        }
    } 
}

#generate cookies
for($count=0;$count -lt $desiredMappings.Count;$count++){
    if($desiredMappings[$count].alreadyMapped){continue}
    if($desiredMappings[$count].sourceLocationPath -eq "autodetect"){
        log -text "Retrieving Onedrive for Business cookie step 1..." 
        #trigger forced authentication to SpO O4B and follow the redirect
        try{
            $res = New-WebRequest -url $baseURL -method GET
        }catch{
            log -text "Failed to retrieve cookie for Onedrive for Business: $($Error[0])" -fout
        }

        #follow first 1-2 redirects, fail if none are detected or if redirects are detected but fail (abnormal flow)
        try{
            $res = (handleO365Redirect -res $res)[0]
        }catch{
            continue
        }

        #MFA check
        try{
            $res = handleMFArequest -res $res -clientId $clientId
            log -text "MFA challenge completed"
        }catch{
            log -text "MFA check result: $_"
        }    

        #sometimes additional redirects are needed, fail if redirects fail, succeed if none are detected or if they are followed
        try{
            $res = (handleO365Redirect -res $res)[0]
        }catch{
            continue
        } 

        $stillProvisioning = $True
        $timeWaited = 0
        while($stillProvisioning){
            if($timeWaited -gt 180){
                $stillProvisioning = $False
                log -text "Failed to auto provision onedrive and/or retrieve username from the response URL. Is this user licensed?" -fout
                $userURL = $userUPN.Replace("@","_").Replace(".","_")
                $mapURL = $mapURLpersonal + $userURL + "\" + $libraryName
                log -text "Will attempt to use auto-guessed value of $mapURL"
            }
            try{
                if($res.rawResponse.ResponseUri.OriginalString.IndexOf("/personal/") -ne -1){
                    $url = $res.rawResponse.ResponseUri.OriginalString
                    $stillProvisioning = $False
                    $start = $url.IndexOf("/personal/")+10 
                    $end = $url.IndexOf("/",$start) 
                    $userURL = $url.Substring($start,$end-$start) 
                    $mapURL = $mapURLpersonal + $userURL + "\" + $libraryName
                    log -text "username detected, your onedrive should be at $mapURL"
                    break
                }else{
                    Throw "No username detected in response string"
                }  
            }catch{
                log -text "Waited for $timeWaited seconds for O4b auto provisioning..."
            }
            if($timeWaited -gt 0){
                $res = New-WebRequest -url "https://$($O365CustomerName)-my.sharepoint.com/_layouts/15/MyBraryFirstRun.aspx?FirstRunStage=waiting" -method GET
            }
            Start-Sleep -s 10
            $res = New-WebRequest -url $baseURL -method GET
            $timeWaited += 10
        }
        $desiredMappings[$count].webDavPath = $mapURL 
        log -text "Onedrive cookie generated"              
    }else{
        log -text "Initiating Sharepoint session with: $($desiredMappings[$count].sourceLocationPath)"
        $spURL = $desiredMappings[$count].sourceLocationPath #URL to browse to
        log -text "Retrieving Sharepoint cookie step 1..." 
        #trigger forced authentication to SpO and follow the redirect if needed
        try{
            $res = New-WebRequest -url $spURL -method GET
        }catch{
            log -text "Failed to retrieve cookie for SpO, will not map this site: $($Error[0])" -fout
            $desiredMappings[$count].alreadyMapped = $True
            continue
        }

        #follow first 1-2 redirects, fail if none are detected or if redirects are detected but fail (abnormal flow)
        try{
            $res = (handleO365Redirect -res $res)[0]
        }catch{
            continue
        }

        #MFA check
        try{
            $res = handleMFArequest -res $res -clientId $clientId
            log -text "MFA challenge completed"
        }catch{
            log -text "MFA check result: $_"
        }    

        #sometimes additional redirects are needed, fail if redirects fail, succeed if none are detected or if they are followed
        try{
            $res = (handleO365Redirect -res $res)[0]
        }catch{
            continue
        } 

        if($desiredMappings[$count]."mapOnlyForSpecificGroup" -eq "favoritesPlaceholder"){
            try{
                try{
                    $documentLibrary = @((returnEnclosedFormValue -res $res -searchString "`"navigationInfo`":" -endString ",`"appBarParams`"" | convertfrom-json).quickLaunch | Where-Object{$_.IsDocLib})[0]
                }catch{
                    try{
                        $documentLibrary = @((returnEnclosedFormValue -res $res -searchString "`"navigationInfo`":" -endString ",`"guestsEnabled`"" | convertfrom-json).quickLaunch | Where-Object{$_.IsDocLib})[0]
                    }catch{
                        try{
                            $documentLibrary = @((returnEnclosedFormValue -res $res -searchString "`"navigationInfo`":" -endString ",`"clientPersistedCacheKey`"" | convertfrom-json).quickLaunch | Where-Object{$_.IsDocLib})[0]
                        }catch{
                            Throw
                        }
                    }
                }
                    
                if(!$documentLibrary){
                    Throw
                }else{
                    $prefix = $spURL.SubString($spURL.IndexOf(".com")+4)
                    $startLoc = $prefix.Length+1
                    $endLoc = ([regex]::Unescape($documentLibrary.Url)).IndexOf("/", $startLoc)
                    $dlName = ([regex]::Unescape($documentLibrary.Url)).SubString($startLoc,$endLoc-$startLoc)
                    $desiredMappings[$count].webDavPath = "$($desiredMappings[$count].webDavPath)\$($dlName)"
                    log -text "auto detected document library url: $($desiredMappings[$count].webDavPath)"
                }
            }catch{
                log -text "Failed to auto detect document library name for $($desiredMappings[$count].displayName), defaulting to $($desiredMappings[$count].webDavPath)\$($favoriteSitesDLName)" -fout
                $desiredMappings[$count].webDavPath = "$($desiredMappings[$count].webDavPath)\$($favoriteSitesDLName)"                   
            }
        }
        #update progress bar
        if($showProgressBar) {
            $script:progressbar1.Value += 2
            $script:form1.Refresh()
        }
        log -text "SpO cookie generated"
    }
}

try{
    setCookies
}catch{
    log -text "Failed to set cookies, error received: $($Error[0])" -fout
}

if($cacheCookies){
    log -text "Caching sessions"
    try{
        $cookies = @()
        foreach($cookie in $script:cookiejar.GetCookies("https://$($O365CustomerName).sharepoint.com")){
            $cookies += $cookie
        }
        foreach($cookie in $script:cookiejar.GetCookies("https://$($O365CustomerName)-my.sharepoint.com")){
            $cookies += $cookie
        }
        foreach($cookie in $script:cookiejar.GetCookies("https://login.microsoftonline.com")){
            $cookies += $cookie
        }
        $cookies | Export-Clixml -Depth 5 -Path $cookieCacheFilePath -Force -Encoding UTF8
    }catch{
        log -text "Caching sessions failed $_" -fout
    }
}

#map the drives
foreach($mapping in $desiredMappings){
    if($mapping.alreadyMapped){continue}
    $mapresult = MapDrive $mapping
    if($mapping.sourceLocationPath -eq "autodetect"){   
        if($autoMapFavoritesMode -eq "Onedrive"){
            $path  = "$($mapping.targetLocationPath)\$autoMapFavoritesLabel"
            if(![System.IO.Directory]::Exists($path)){
                log -text "Desired path for Team site links: $path does not exist, creating"
                try{
                    $res = New-Item -Path $path -ItemType Directory -Force
                }catch{
                    log -text "Failed to create folder $path! $($Error[0])" -fout
                }
            }else{
                try{
                    Get-ChildItem $path | Remove-Item -Force -Confirm:$False -Recurse
                }catch{$Null}    
            }
        }      
        if($addShellLink -and $windowsVersion -eq 6 -and $mapping.targetLocationType -eq "driveletter" -and [System.IO.Directory]::Exists($mapping.targetLocationPath)){
            try{
                $res = createFavoritesShortcutToO4B -targetLocation $mapping.targetLocationPath
            }catch{
                log -text "Failed to create a shortcut to the mapped drive for Onedrive for Business because of: $($Error[0])" -fout
            }
        }
    }
} 

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 90
    $script:form1.Refresh()
}

if($redirectFolders){
    $listOfFoldersToRedirect | % {
        log -text "Redirecting $($_.knownFolderInternalName) to $($_.desiredTargetPath)"
        try{
            Redirect-Folder -GetFolder $_.knownFolderInternalName -SetFolder $_.knownFolderInternalIdentifier -Target $_.desiredTargetPath -copyExistingFiles $_.copyExistingFiles
            log -text "Redirected $($_.knownFolderInternalName) to $($_.desiredTargetPath)"
        }catch{
            log -text "Failed to redirect $($_.knownFolderInternalName) to $($_.desiredTargetPath): $($Error[0])" -fout
        }
    }
}

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 100
    $script:form1.Refresh()
}

abort_OM
