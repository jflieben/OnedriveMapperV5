######## 
#OneDriveMapper
#Copyright/License: https://www.lieben.nu/liebensraum/commercial-use/ (Commercial (re)use not allowed without prior written consent by the author, otherwise free to use/modify as long as header are kept intact)
#Author:            Jos Lieben (Lieben Consultancy)
#Script help:       https://www.lieben.nu/liebensraum/onedrivemapper/
#Purpose:           This script maps Onedrive for Business and/or maps a configurable number of Sharepoint Libraries
#Enterprise users:  This script is not recommended for business critical Enterprise use, see https://www.lieben.nu/liebensraum/onedrivemapper/onedrivemapper-cloud/ for alternatives
#Requirements:      Keep Me Signed in (sign in acceleration) allowed (Tenant wide). Trusted sites already configured or user allowed to configure them. PowerShell v3 or higher

param(
    [Switch]$asTask,
    [Switch]$hideConsole
)

$version = "5.15"

####REQUIRED MANUAL CONFIGURATION
$O365CustomerName      = "lieben"          #This should be the name of your tenant (example, lieben as in lieben.onmicrosoft.com) 
$showConsoleOutput     = $True             #Set this to $False to hide console output
$useAzAdConnectSSO     = $false            #Set to true if using Azure Ad Connect SSO. Do NOT set the aadg.windows.net.nsatc.net and autologon.microsoftazuread-sso.com zones forcibly through GPO as ODM will temporarily remove them for mapping and then readd them

<#
HELPTEXT: if you wish to add more, add more lines to the below (copy the first above itself). Parameter explanation:
displayName = the label of the driveletter, or name of the shortcut we'll create to the target site/library
targetLocationType = driveletter, converged OR networklocation, if you use driveletter, enter a driveletter in targetLocationPath. If you use networklocation, enter a path to a folder where you want the shortcut to be created. Converged driveletters are a collection of links (fake driveletter with links to all mappings you want)
targetLocationPath = enter a driveletter if mapping to a driveletter, enter a folder path if just creating shortcuts
sourceLocationPath = autodetect or the full URL to the sharepoint / groups site. Autodetect automatically makes this a mapping to Onedrive For Business
mapOnlyForSpecificGroup = this only works for DOMAIN JOINED devices that can reach a domain controller and means that the mapping will only be made if the user is a member of the group you specify here
#>

#DEFAULT SETTINGS: (onedrive only, to the X: drive)
$desiredMappings =  @(
    @{"displayName"="Onedrive for Business";"targetLocationType"="driveletter";"targetLocationPath"="X:";"sourceLocationPath"="autodetect";"mapOnlyForSpecificGroup"=""}
    @{"displayName"="Sharepoint Site A";"targetLocationType"="driveletter";"targetLocationPath"="Z:";"sourceLocationPath"="https://lieben.sharepoint.com/sites/groep30/Gedeelde%20%20documenten/Forms/AllItems.aspx";"mapOnlyForSpecificGroup"=""}
    @{"displayName"="Sharepoint Site A";"targetLocationType"="driveletter";"targetLocationPath"="Q:";"sourceLocationPath"="https://lieben.sharepoint.com/sites/groep30/Gedeelde%20%20documenten/Brondata";"mapOnlyForSpecificGroup"=""}
)

<#
EXAMPLE SETTINGS (Onedrive for Business, two Sharepoint sites, one mapped to a driveletter, one to a shortcut, the last only when a member of the Active Directory group SEC-SHAREPOINTA and two sharepoint sites mapped as links (converged) into a fake driveletter Y)
$desiredMappings =  @(
    @{"displayName"="Onedrive for Business";"targetLocationType"="driveletter";"targetLocationPath"="X:";"sourceLocationPath"="autodetect";"mapOnlyForSpecificGroup"=""},
    @{"displayName"="Sharepoint Site A";"targetLocationType"="networklocation";"targetLocationPath"="$env:APPDATA\Microsoft\Windows\Network Shortcuts";"sourceLocationPath"="https://lieben.sharepoint.com/sites/lieben/Gedeelde%20%20documenten/Forms/AllItems.aspx";"mapOnlyForSpecificGroup"="SEC-SHAREPOINTA"}
    @{"displayName"="Sharepoint Site A";"targetLocationType"="driveletter";"targetLocationPath"="Z:";"sourceLocationPath"="https://lieben.sharepoint.com/sites/groep30/Gedeelde%20%20documenten/Forms/AllItems.aspx";"mapOnlyForSpecificGroup"=""}
    @{"displayName"="Sharepoint Site B";"targetLocationType"="converged";"targetLocationPath"="Y:";"sourceLocationPath"="https://lieben.sharepoint.com/sites/groep30/Gedeelde%20%20documenten/Forms/AllItems.aspx";"mapOnlyForSpecificGroup"="AD Group SPSB"} 
    @{"displayName"="Sharepoint Site C";"targetLocationType"="converged";"targetLocationPath"="Y:";"sourceLocationPath"="https://lieben.sharepoint.com/sites/groep30/Gedeelde%20%20documenten/Forms/AllItems.aspx";"mapOnlyForSpecificGroup"="AD Group SPSC"} 
)
#>

$redirectFolders       = $false #Set to TRUE and configure below hashtable to redirect folders to locations you're mapping (e.g. onedrive, teams, sharepoint)
$listOfFoldersToRedirect = @(#One line for each folder you want to redirect, only works if redirectFolders=$True. For knownFolderInternalName choose from Get-KnownFolderPath function, for knownFolderInternalIdentifier choose from Set-KnownFolderPath function
    @{"knownFolderInternalName" = "Desktop";"knownFolderInternalIdentifier"="Desktop";"desiredTargetPath"="X:\Desktop";"copyExistingFiles"="true"}
    @{"knownFolderInternalName" = "MyDocuments";"knownFolderInternalIdentifier"="Documents";"desiredTargetPath"="X:\My Documents";"copyExistingFiles"="true"}
    @{"knownFolderInternalName" = "MyPictures";"knownFolderInternalIdentifier"="Pictures";"desiredTargetPath"="X:\My Pictures";"copyExistingFiles"="false"}
)

###OPTIONAL CONFIGURATION
$autoUpdateEdgeDriver  = $True                     #Automatically update msedgedriver (otherwise you need to do this manually/frequently!)
$autoRemapMethod       = "Path"                    #automatically rerun if a connection is dropped / lost but an active internet connection exists. Options: "Path" (checks underlying webdav connection), "Link" (checks existence of driveletter or shortcut as well, only works for drivemappings and converged drives), "Disabled" (no reruns)
$restartExplorer       = $False                    #You can safely set this to False if you're not redirecting folders, if used with autoRemapMethod this can be very intrusive for users.
$libraryName           = "Documents"               #leave this default, unless you wish to map a non-default onedrive library you've created. Only used if it cannot be autodetected for some reason
$displayErrors         = $True                     #show errors to user in visual popups
$persistentMapping     = $True                     #If set to $False, the mapping will go away when the user logs off
$urlOpenAfter          = ""                        #This URL will be opened by the script after running if you configure it
$showProgressBar       = $True                     #will show a progress bar to the user
$progressBarColor      = "#CC99FF"
$progressBarText       = "OnedriveMapper v$version is (re)connecting your drives..."
$convergedDriveLabel   = "Sharepoint and Team sites" #used only if you're doing converged drive mappings
$autoDetectProxy       = $False                    #if set to $False, unchecks the 'Automatically detect proxy settings' setting; this greatly enhanced WebDav performance, set to true to not modify this setting (leave as is)
$addShellLink          = $False                    #Adds a link to Onedrive to the Shell under Favorites (Windows 7, 8 / 2008R2 and 2012R2 only) If you use a remote path, google EnableShellShortcutIconRemotePath
$removeExistingMaps    = $True                     #Removes any existing drive mappings if $True ($false to disable)
$removeEmptyMaps       = $True                     #Removes any existing empty drive maps if $True ($false to disable)
$logfile               = ($env:APPDATA + "\OneDriveMapper_$version.log")    #Logfile to log to 
$driversLocation       = ($env:APPDATA)            #location where the Edge and Selenium drivers are located. If not present, these are downloaded automatically 
$forceHideEdge         = $false                    #Forcibly ensures the user never sees edge/ps windows. Warning: also does not show authentication dialogs, so only useful if your SSO method is working 100%
$autoClearAllCookies   = $False                    #always clear all cookies before running (prevents/fixes certain occasional issues with cookies)
$createUserFolderOn    = "Q:"                      #creates a user folder if not already present and maps that instead for the given driveletter(s). If multiple drives, separate them with a comma (e.g. Q:,P:,Z:)

$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native

if($hideConsole){
    try{
        [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
    }catch{$Null}
}

######## 
#Required resources and some customizations you'll probably not use
########
$privateSuffix = "-my"
$script:errorsForUser = ""
$onedriveIconPath = "C:\GitRepos\OnedriveMapper\onedrive.ico" #if this file exists, and you've set addShellLink to True, it will be used as icon for the shortcut
$sharepointIconPath = "C:\GitRepos\OnedriveMapper\sharepoint.ico" #if this file exists, and you've set addShellLink to True, it will be used as icon for the shortcut
$i_MaxLocalLogSize = 2 #max local log size in MB
$maxWaitSecondsForSpO  = 5                        #Maximum seconds the script waits for Sharepoint Online to load before mapping
$o365loginURL = "https://login.microsoftonline.com/login.srf?msafed=0"

$O365CustomerName = $O365CustomerName.ToLower() 
#for people that don't RTFM, fix wrongly entered customer names:
$O365CustomerName = $O365CustomerName -Replace ".onmicrosoft.com",""
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
    param(
        [string]$networkLocationPath="$env:APPDATA\Microsoft\Windows\Network Shortcuts",
        [Parameter(Mandatory=$true)][string]$networkLocationName ,
        [Parameter(Mandatory=$true)][string]$networkLocationTarget,
        [String]$iconPath
    )
    Begin{
        Write-Verbose -Message "Network location path: `"$networkLocationPath`"."
        Write-Verbose -Message "Network location name: `"$networkLocationName`"."
        Write-Verbose -Message "Network location target: `"$networkLocationTarget`"."
        Set-Variable -Name desktopIniContent -Option ReadOnly -value ([string]"[.ShellClassInfo]`r`nCLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}`r`nFlags=2")
    }Process{
        Write-Verbose -Message "Checking that `"$networkLocationPath`" is a valid directory..."
        if(Test-Path -Path $networkLocationPath -PathType Container){
            try{
                if(Test-Path -Path $networkLocationPath\$networkLocationName -PathType Container){
                    Write-Verbose -Message "`"$networkLocationPath\$networkLocationName`". already exists"
                }else{
                    Write-Verbose -Message "Creating `"$networkLocationPath\$networkLocationName`"."
                    [void]$(New-Item -Path "$networkLocationPath\$networkLocationName" -ItemType Directory -ErrorAction Stop)
                    Write-Verbose -Message "Setting system attribute on `"$networkLocationPath\$networkLocationName`"."
                    Set-ItemProperty -Path "$networkLocationPath\$networkLocationName" -Name Attributes -Value ([System.IO.FileAttributes]::System) -ErrorAction Stop
                }
            }catch [Exception]{
                Write-Error -Message "Cannot create or set attributes on `"$networkLocationPath\$networkLocationName`". Check your access and/or permissions."
                return $false
            }
        }else{
            Write-Error -Message "`"$networkLocationPath`" is not a valid directory path."
            return $false
        }

        try{
            if(Test-Path -Path "$networkLocationPath\$networkLocationName\desktop.ini" -PathType Leaf){
                Write-Verbose -Message "`"$networkLocationPath\$networkLocationName\desktop.ini`". already exists"
            }else{
                Write-Verbose -Message "Creating `"$networkLocationPath\$networkLocationName\desktop.ini`"."
                $Null = New-Item -Path "$networkLocationPath\$networkLocationName\desktop.ini" -ItemType File
            }
            Write-Verbose -Message "Writing to $networkLocationPath\$networkLocationName\desktop.ini"
            Set-Content -Path "$networkLocationPath\$networkLocationName\desktop.ini" -Value $desktopIniContent
        }catch [Exception]{
            Write-Error -Message "Error while creating or writing to `"$networkLocationPath\$networkLocationName\desktop.ini`". Check your access and/or permissions."
            return $false
        }

        try{
            $WshShell = New-Object -ComObject WScript.Shell
            Write-Verbose -Message "Creating shortcut to `"$networkLocationTarget`" at `"$networkLocationPath\$networkLocationName\target.lnk`"."
            $Shortcut = $WshShell.CreateShortcut("$networkLocationPath\$networkLocationName\target.lnk")
            $Shortcut.TargetPath = $networkLocationTarget
            if([System.IO.File]::Exists($iconPath)){
                $Shortcut.IconLocation = "$($iconPath), 0"
            }            
            $Shortcut.Description = "Created $(Get-Date -Format s) by $($MyInvocation.MyCommand)."
            $Shortcut.Save()
        }catch [Exception]{
            Write-Error -Message "Error while creating shortcut @ `"$networkLocationPath\$networkLocationName\target.lnk`". Check your access and permissions."
            return $false
        }
        return $true
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
	[WinAPI.Explorer]::Refresh()
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

function getElementById{
    Param(
        [Parameter(Mandatory=$true)]$id
    )
    $localObject = $Null
    try{
        $localObject = $global:edgeDriver.FindElementById($id)
        if($Null -eq $localObject.tagName){Throw "The element $id was not found (1) or had no tagName"}
        return $localObject
    }catch{$localObject = $Null}
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
    log -text "Restarting Explorer.exe to make the drive(s) visible"  
    #kill all running explorer instances of this user  
    $explorerStatus = Get-ProcessWithOwner explorer  
    if($explorerStatus -eq 0){  
        log -text "no instances of Explorer running yet, at least one should be running" -warning 
    }elseif($explorerStatus -eq -1){  
        log -text "ERROR Checking status of Explorer.exe: unable to query WMI" -fout 
    }else{  
        log -text "Detected running Explorer processes, attempting to shut them down..."  
        foreach($Process in $explorerStatus){  
            try{  
                Stop-Process $Process.handle | Out-Null  
                log -text "Stopped process with handle $($Process.handle)"  
            }catch{  
                log -text "Failed to kill process with handle $($Process.handle)" -fout 
            }  
        }  
    }  
}  

function checkIfAtO365URL{
    param(
        [Array]$finalURLs
    )
    $url = $global:edgeDriver.Url
    foreach($item in $finalURLs){
        if($url.StartsWith($item)){
            return $True
        }
    }
    return $False
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
            [WinAPI.Explorer]::Refresh()
        }catch{ 
            log -text "Failed to set the drive label: $_ " -fout
        } 
 
    } 
} 

function fixElevationVisibility{
    if($showConsoleOutput){
        $createTask = "schtasks /Create /SC ONCE /TN OnedriveMapper /IT /RL LIMITED /F /TR `"Powershell.exe -NoProfile -ExecutionPolicy ByPass -File '$scriptPath\OnedriveMapper.ps1' -asTask`" /st 00:00"    
    }else{
        $createTask = "schtasks /Create /SC ONCE /TN OnedriveMapper /IT /RL LIMITED /F /TR `"Powershell.exe -NoProfile -ExecutionPolicy ByPass -WindowStyle Hidden -File '$scriptPath\OnedriveMapper.ps1' -asTask`" /st 00:00"
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
            try{$out = NET USE $($driveMapping.webDavPath) /PERSISTENT:YES 2>&1}catch{$Null}
        }else{
            try{$out = NET USE $($driveMapping.webDavPath) /PERSISTENT:NO 2>&1}catch{$Null}
        }
        if($out -like "*error 67*"){
            log -text "ERROR: detected string error 67 in return code of net use command, this usually means the WebClient isn't running" -fout
        }
        if($out -like "*error 224*"){
            log -text "ERROR: detected string error 224 in return code of net use command, this usually means your trusted sites are misconfigured or KB2846960 is missing or Edge needs a reset" -fout
        }
        if($LASTEXITCODE -ne 0){ 
            log -text "Failed to map $($driveMapping.targetLocationPath) to $($driveMapping.webDavPath), error: $($LASTEXITCODE) $($out) $del" -fout
            $script:errorsForUser += "$($driveMapping.targetLocationPath) could not be mapped because of error $($LASTEXITCODE) $($out) d$del`n"
        } 

        #check if we need to create a user folder:
        if($createUserFolderOn -like "*$($driveMapping.targetLocationPath)*"){
            log -text "this is a mapping we should create a user folder in if it doesn't exist, checking..."
            $targetUserfolderPath = $Null; $targetUserfolderPath = (Join-Path $driveMapping.webDavPath -ChildPath $env:USERNAME)
            if(!(Test-Path $targetUserfolderPath)){
                log -text "creating $targetUserfolderPath ...."
                $Null = New-Item -Path $targetUserfolderPath -ItemType Directory -Force -Confirm:$False
                log -text "$targetUserfolderPath created!"
            }else{
                log -text "$targetUserfolderPath already exists"
            }
            $driveMapping.webDavPath = $targetUserfolderPath
        }
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
            log -text "ERROR: detected string error 224 in return code of net use command, this usually means your trusted sites are misconfigured or KB2846960 is missing or Edge needs a reset" -fout
        }
        if($LASTEXITCODE -ne 0){ 
            log -text "Failed to map $($driveMapping.targetLocationPath) to $($driveMapping.webDavPath), error: $($LASTEXITCODE) $($out) $del" -fout
            $script:errorsForUser += "$($driveMapping.targetLocationPath) could not be mapped because of error $($LASTEXITCODE) $($out) d$del`n"
        } 
        if((Test-Path $($driveMapping.webDavPath))){ 
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
            }else{
                $desiredIconPath = $sharepointIconPath
            }
            log -text "Mapping target: $($driveMapping.webDavPath)" 
            try{$del = NET USE $($driveMapping.webDavPath) /DELETE /Y 2>&1}catch{$Null}
            if($persistentMapping){
                try{$out = NET USE $($driveMapping.webDavPath) /PERSISTENT:YES 2>&1}catch{$Null}
            }else{
                try{$out = NET USE $($driveMapping.webDavPath) /PERSISTENT:NO 2>&1}catch{$Null}
            }            
            $res = Add-NetworkLocation -networkLocationPath $($driveMapping.targetLocationPath) -networkLocationName $($driveMapping.displayName) -networkLocationTarget $($driveMapping.webDavPath) -iconPath $desiredIconPath -Verbose -ErrorAction Stop
            if((Test-Path $($driveMapping.webDavPath))){ 
                log -text "Added network location $($driveMapping.displayName)"
                return $True
            }else{
                log -text "failed to contact $($driveMapping.targetLocationPath) after mapping it to $($driveMapping.webDavPath), check if the URL is valid. Error: $($error[0]) $out" -fout
                return $False
            }
        }catch{
            log -text "failed to add network location: $($Error[0])" -fout
            return $False
        }
    }
} 
 
function run-CleanUp{ 
    $global:edgeDriver.Quit()
    
    if($showProgressBar) {
        $progressbar1.Value = 100
        $label1.text="Done!"
        Start-Sleep -Milliseconds 500
        $form1.Close()
    }

    if($useAzAdConnectSSO){
        if((addSiteToIEZoneThroughRegistry -siteUrl "aadg.windows.net.nsatc.net" -mode 1) -eq $True){
            log -text "Automatically added aadg.windows.net.nsatc.net to intranet sites for this user"
        }
        if((addSiteToIEZoneThroughRegistry -siteUrl "autologon.microsoftazuread-sso.com" -mode 1) -eq $True){
            log -text "Automatically added autologon.microsoftazuread-sso.com to intranet sites for this user"
        }
    }

    if($restartExplorer){ 
        restart_explorer 
    }else{ 
        #Show warning only if redirecting folders is requested
        if ($redirectFolders){        
            log -text "restartExplorer is set to False, if you're redirecting folders they may not show up" -warning 
        }
    }     
    if($urlOpenAfter.Length -gt 10){Start-Process msedge.exe $urlOpenAfter}
    if($displayErrors){
        if($errorsForUser){ 
            $OUTPUT= [System.Windows.Forms.MessageBox]::Show($errorsForUser, "Onedrivemapper Error" , 0) 
            $OUTPUT2= [System.Windows.Forms.MessageBox]::Show("You can always use https://portal.office.com to access your data", "Need a workaround?" , 0) 
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

function start-AuthCheck(){
    $kmsiDetected = $False
    if((checkIfAtO365URL -finalURLs $finalURLs) -eq $True){
        log -text "You're already logged in! No need to display login dialog" 
    }else{
        log -text "Encountered a dialog, showing dialog to user" 
        $global:cachedHwnds | % {[native.win]::ShowWindow($_,5)}
        [Win32SetWindow]::SetForegroundWindow($global:cachedHwnds)
        $waited = 0
        while((checkIfAtO365URL -finalURLs $finalURLs) -ne $True){ 
            $waited += 0.2
            Start-Sleep -Milliseconds 200
            if($waited -gt 300){
                log -text "User did not sign in to $($global:edgeDriver.Url) within 5 minutes, aborting"
                run-CleanUp
                Exit
            }
            try{
                $checkBox = getElementById -id "KmsiCheckboxField"
                if($checkbox.Displayed){
                    if(!$checkbox.Enabled){
                        $checkBox.Click()
                    }
                    $kmsiDetected = $True
                    (getElementById -id "idSIButton9").Click()
                }
            }catch{$Null}
            try{
                $checkBox = getElementById -id "idChkBx_SAOTCC_TD"
                if($checkbox.Displayed){
                    if(!$checkbox.Enabled){
                        $checkBox.Click()
                    }
                    $kmsiDetected = $True
                }
            }catch{$Null}
        }
        if($kmsiDetected){
            log -text "KMSI prompt detected"
        }else{
            log -text "KMSI prompt not detected, check FAQ if sign in fails!" -warning
        }
        log -text "User completed dialog" 
        $global:cachedHwnds | % {[native.win]::ShowWindow($_,0)}     
    }
}

function add-cookies{
    [DateTime]$dateTime = Get-Date
    $dateTime = $dateTime.AddDays(5)
    $str = $dateTime.ToString("R")
    foreach($cookie in $global:edgeDriver.Manage().Cookies.AllCookies){
        [String]$cookieValue = [String]$cookie.Value.Trim()
        [String]$cookieDomain = [String]$cookie.Domain.Trim()
        try{
            if($cookie.Name -eq "rtFa"){
                $cookieDomain = "https://$($cookieDomain.replace(".sharepoint","sharepoint"))"
                log -text "Setting rtFA cookie for $cookieDomain...."
                $res = [LiebenConsultancy.cookieSetter]::SetWinINETCookieString($cookieDomain,"rtFa","$cookieValue;Expires=$str")
            }
            if($cookie.Name -eq "FedAuth"){
                $cookieDomain = "https://$($cookieDomain)"
                log -text "Setting FedAuth cookie for $cookieDomain...."
                $res = [LiebenConsultancy.cookieSetter]::SetWinINETCookieString($cookieDomain,"FedAuth","$cookieValue;Expires=$($str)")
            }
            log -text "$cookieDomain cookie stored"
        }catch{
            log -text "Failed to set a cookie: $($Error[0])" -fout
        }
    }
}

$scriptPath=$(if ($psISE) {Split-Path -Path $psISE.CurrentFile.FullPath} else {$(if ($global:PSScriptRoot.Length -gt 0) {$global:PSScriptRoot} else {$global:pwd.Path})})

$scriptPath = $scriptPath.Replace("Microsoft.PowerShell.Core\FileSystem::","")

ResetLog
log -text "-----$(Get-Date) OneDriveMapper v$version - $($env:USERNAME) on $($env:COMPUTERNAME) starting-----" 

log -text "OnedriveMapper is running from $scriptPath"

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
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Web")
$definition = @'
[System.Runtime.InteropServices.DllImport("Shell32.dll")] 
private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
public static void Refresh() {
    SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);    
}
'@
    Add-Type -MemberDefinition $definition -Namespace WinAPI -Name Explorer
}catch{ 
    log -text "Error loading windows libraries, script will likely fail" -fout
} 

#try to set TLS to v1.2, Powershell defaults to v1.0
try{
    $res = [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    log -text "Set TLS protocol version to prefer v1.2"
}catch{
    log -text "Failed to set TLS protocol to perfer v1.2 $($Error[0])" -fout
}

#load the Selenium component
try{
    log -text "Loading Selenium driver ($($driversLocation)\WebDriver.dll)"
    $seleniumDriverPath = "$($driversLocation)\WebDriver.dll"
    if(!(Test-Path -Path $seleniumDriverPath)){
        log -text "Selenium driver not present at $seleniumDriverPath, trying to download automatically from trusted source" -warning
        try{
            Invoke-WebRequest -uri "https://gitlab.com/Lieben/OnedriveMapper_V3/-/raw/master/WebDriver.dll?inline=false" -OutFile $seleniumDriverPath -Method Get -UseBasicParsing -ErrorAction Stop
            if(!(Test-Path -Path $seleniumDriverPath)){
                log -text "Failed to download Selenium Driver because of $($Error[0])" -fout
                Throw
            }
            log -text "Selenium driver downloaded"
        }catch{
            Throw "Ensure WebDriver.dll is present in the same folder as OnedriveMapper.ps1. You can download WebDriver.dll from https://www.nuget.org/packages/Selenium.WebDriver or https://gitlab.com/Lieben/OnedriveMapper_V3/-/raw/master/WebDriver.dll?inline=false"
        }
    }
    try{
        $driverBlocked = Get-Item $seleniumDriverPath -Stream "Zone.Identifier" -ErrorAction Stop
    }catch{
        log -text "Selenium driver present and not blocked by zoning, loading..."
    }
    if($driverBlocked){
        log -text "Selenium driver was downloaded from the internet, so we need to run Unblock-File"
        try{
            Unblock-File -Path $seleniumDriverPath -Confirm:$False
            log -text "Selenium driver automatically unblocked"
        }catch{
            Throw "Selenium driver not trusted by windows OS, right click WebDriver.dll and unblock it in Properties or run Unblock-File"
        }
    }
    $bytes = [System.IO.File]::ReadAllBytes($seleniumDriverPath)
    [System.Reflection.Assembly]::Load($bytes)
    log -text "Selenium loaded successfully"
}catch{
    log -text "Failed to load Selenium driver, cannot continue! Error details: $($Error[0])" -fout
    Exit
}

#load cookie code and test-set a cookie

log -text "Loading CookieSetter..."
$source=@"
using System.Runtime.InteropServices;
using System;
namespace LiebenConsultancy
{
    public static class cookieSetter
    {
        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool InternetSetCookie(string url, string name, string data);

        public static bool SetWinINETCookieString(string url, string name, string data)
        {
            bool res = cookieSetter.InternetSetCookie(url, name, data);
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
    $res = [LiebenConsultancy.cookieSetter]::SetWinINETCookieString("https://testdomainthatdoesnotexist.com","none","Data=nothing;Expires=$str")
    log -text "Test cookie set successfully"
}catch{
    log -text "ERROR: Failed to set test cookie, script will fail: $($Error[0])" -fout
}


#get OSVersion
$windowsVersion = ([System.Environment]::OSVersion.Version).Major

#get Edge version
try{
    $edgeVersion = (Get-ItemProperty -Path (Join-Path ${env:ProgramFiles(x86)} -ChildPath "Microsoft\Edge\Application\msedge.exe")).VersionInfo.ProductVersion
}catch{
    $edgeVersion = $null
}

if(!$edgeVersion){
    try{
        $edgeVersion = (Get-ItemProperty -Path (Join-Path $env:ProgramFiles -ChildPath "Microsoft\Edge\Application\msedge.exe")).VersionInfo.ProductVersion
    }catch{
        $edgeVersion = $null
    }
}

if(!$edgeVersion){
    try{
        $edgeVersion = (Get-AppxPackage -Name "Microsoft.MicrosoftEdge.Stable" -ErrorAction Stop).Version
    }catch{
        $edgeVersion = $null
    }
}

if($edgeVersion -is [Array]){
    [String]$edgeVersion = $edgeVersion[-1]        
}

$edgeVersion = $edgeVersion.Split(" ")[-1]

if(!$edgeVersion){
    log -text "Could not detect Edge version on this system." -fout
}

log -text "You are $($Env:USERNAME) running on Windows $windowsVersion with Powershell version $($PSVersionTable.PSVersion.Major) and Edge version $edgeVersion"

if($windowsVersion -eq 6){
    log -text "Windows 2012 R2 is not supported, please use an older version of ODM or update your OS!" -fout
}

#get .NET versions
try{
    $v4Client = get-itempropertyvalue "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Client" -Name Version -ErrorAction Stop
}catch{
    $v4Client = $Null
}
try{
    $v4Full = get-itempropertyvalue "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name Version -ErrorAction Stop
}catch{
    $v4Full = $Null
}

$netVersion = $Null

if($v4Client){
    $netVersion = $v4Client
    log -text ".NET V4 client version: $netVersion"
}elseif($v4Full){
    $netVersion = $v4Client
    log -text ".NET V4 full version: $netVersion"
}

if(!$netVersion -or $netVersion -lt 4.8){
    log -text ".NET 4.8 or higher is required to run OnedriveMapper! Bypass at your own risk by downloading a supported version of Selenium from https://www.nuget.org/packages/Selenium.WebDriver and modifying this code. No free support will be provided." -fout
    Exit
}

if($showConsoleOutput -eq $False){
    log -text "hiding console window..."
    try{
        [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
        log -text "console hidden"
    }catch{$Null}
}

log -text "loading interop service"
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    public class Win32SetWindow {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
"@
log -text "interop service loaded"
 
#Check if webdav locking is enabled
if((checkRegistryKeyValue -basePath "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\" -entryName "SupportLocking") -ne 0){
    log -text "ERROR: WebDav File Locking support is enabled, this could cause files to become locked in your OneDrive or Sharepoint site" -fout 
} 

#report/warn file size limit
$sizeLimit = [Math]::Round((checkRegistryKeyValue -basePath "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\" -entryName "FileSizeLimitInBytes")/1024/1024)
log -text "Maximum file upload size is set to $sizeLimit MB" -warning

#check if zone enforcement is set to machine only
$reg_HKLM = checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings" -entryName "Security_HKLM_only"
if($reg_HKLM -eq -1){
    log -text "NOTICE: IE Security zones ambiguous - checking both computer and user" 
}elseif($reg_HKLM -eq 1){
    log -text "NOTICE: IE Security zones configured via computer policy"    
}

#Check if Zone Configuration is on a per machine or per user basis, then check the zones 
$privateZoneFound = $False
$publicZoneFound = $False

#check if sharepoint and onedrive are set as safe sites at the user level
if($reg_HKLM -ne 1){
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -match '^[1-2]+$'){
        log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level"  
        $privateZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -match '^[1-2]+$'){
        log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level (through GPO)"  
        $privateZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -match '^[1-2]+$'){
        log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on user level"  
        $publicZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -match '^[1-2]+$'){
        log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on user level (through GPO)" 
        $publicZoneFound = $True        
    }
}

#check if sharepoint and onedrive are set as safe sites at the machine level
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -match '^[1-2]+$'){
    log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level"
    $privateZoneFound = $True 
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -match '^[1-2]+$'){
    log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level (through GPO)"  
    $privateZoneFound = $True        
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -match '^[1-2]+$'){
    log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on machine level"  
    $publicZoneFound = $True    
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -match '^[1-2]+$'){
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
    log -text "Possible critical error: $($O365CustomerName).sharepoint.com not found in IE Trusted Sites on user or machine level"
    if((addSiteToIEZoneThroughRegistry -siteUrl "$O365CustomerName.sharepoint.com") -eq $True){
        log -text "Automatically added $O365CustomerName.sharepoint.com to trusted sites for this user"
    }else{
        log -text "Failed to automatically add $O365CustomerName.sharepoint.com to trusted sites for this user, the script will likely fail" -fout
    }
}
if($privateZoneFound -eq $False){
    log -text "Possible critical error: $($O365CustomerName)$($privateSuffix).sharepoint.com not found in IE Trusted Sites on user or machine level"
    if((addSiteToIEZoneThroughRegistry -siteUrl "$($O365CustomerName)$($privateSuffix).sharepoint.com") -eq $True){
        log -text "Automatically added $($O365CustomerName)$($privateSuffix).sharepoint.com to trusted sites for this user"
    }else{
        log -text "Failed to automatically add $($O365CustomerName)$($privateSuffix).sharepoint.com to trusted sites for this user, the script will likely fail" -fout
    }
}

#Check and log if Explorer is running 
$explorerStatus = Get-ProcessWithOwner explorer 
if($explorerStatus -eq 0){ 
    log -text "no instances of Explorer running yet, expected at least one running" -warning
}elseif($explorerStatus -eq -1){ 
    log -text "Checking status of explorer.exe: unable to query WMI" -fout
}else{ 
    log -text "Detected running explorer process" 
} 

#clean up any existing mappings
if ($removeExistingMaps){
    subst | % {subst $_.SubString(0,2) /D}
    Get-PSDrive -PSProvider filesystem | Where-Object {$_.DisplayRoot} | % {
        if($_.DisplayRoot.StartsWith("\\$($O365CustomerName).sharepoint.com") -or $_.DisplayRoot.StartsWith("\\$($O365CustomerName)-my.sharepoint.com")){
            try{$del = NET USE "$($_.Name):" /DELETE /Y 2>&1}catch{$Null}     
        }
    }
}

#clean up empty mappings
if ($removeEmptyMaps){
    Get-PSDrive -PSProvider filesystem | Where-Object {($_.Used -eq 0 -and $_.Free -eq $Null)} | % {
        try{$_ | Remove-PSDrive -Force}catch{$Null}     
    }
}

#check which mappings require a group membership and add/remove. Note: check is done through the CN (NAME) of the group, not the DisplayName
if($desiredMappings | Where-Object{$_.mapOnlyForSpecificGroup.Length -gt 0}){
    try{
        $groups = ([ADSISEARCHER]"(member:1.2.840.113556.1.4.1941:=$(([ADSISEARCHER]"samaccountname=$($env:USERNAME)").FindOne().Properties.distinguishedname))").FindAll().Properties.distinguishedname -replace '^CN=([^,]+).+$','$1'
        log -text "cached user group membership because you have configured mappings where the mapOnlyForSpecificGroup option was configured"   
    }catch{
        log -text "failed to cache user group membership, ignoring these mappings because of: $($Error[0])" -fout
        $desiredMappings = $desiredMappings | Where-Object{$_.mapOnlyForSpecificGroup.Length -eq 0}
    }
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

if($autoClearAllCookies){
    log -text "Clearing cookies..."
    & RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
    Start-Sleep -s 10
}

$baseURL = ("https://$($O365CustomerName)-my.sharepoint.com/_layouts/15/MySite.aspx?MySiteRedirect=AllDocuments") 
$mapURLpersonal = "\\$O365CustomerName-my.sharepoint.com@SSL\DavWWWRoot\personal\"

$intendedmappings = @()
for($count=0;$count -lt $desiredMappings.Count;$count++){
    #replace funky sharepoint URL stuff and turn into webdav path
    if($desiredMappings[$count].sourceLocationPath -ne "autodetect"){
        $desiredMappings[$count].webDavPath = [System.Web.HttpUtility]::UrlDecode($desiredMappings[$count].sourceLocationPath)
        $desiredMappings[$count].webDavPath = $desiredMappings[$count].webDavPath.Replace("https://","\\").Replace("/_layouts/15/start.aspx#","").Replace("sharepoint.com","sharepoint.com@SSL\DavWWWRoot").Replace("/Forms/AllItems.aspx","").Replace("%27","'")
        $desiredMappings[$count].webDavPath = $desiredMappings[$count].webDavPath.Replace("/","\")  
    }else{
        $desiredMappings[$count].webDavPath = $mapURLpersonal
    }

    if($desiredMappings[$count].mapOnlyForSpecificGroup -and $groups){
        $group = $groups -contains $desiredMappings[$count].mapOnlyForSpecificGroup
        if($group){ 
            log -text "adding a sharepoint mapping because the user is a member of $($desiredMappings[$count].mapOnlyForSpecificGroup)" 
        }else{
            continue
        }
    }
    $intendedmappings += $desiredMappings[$count]
}

#prepare converged drives if configured
$convergedDrives = @($intendedMappings | Where-Object {$_.targetLocationType -eq "converged"})
if($convergedDrives){
    $convergedDriveLetters = $convergedDrives.targetLocationPath | Select-Object -Unique
    foreach($convergedDriveletter in $convergedDriveLetters){
        $targetFolder = Join-Path $Env:TEMP -ChildPath "OnedriveMapperLinks $($convergedDriveletter.SubString(0,1))" 
        if(![System.IO.Directory]::Exists($targetFolder)){
            log -text "Converged drive source folder $targetFolder does not exist, creating"
            try{
                $res = New-Item -Path $targetFolder -ItemType Directory -Force
                log -text "Converged drive $convergedDriveletter created in $targetFolder"
            }catch{
                log -text "Failed to create folder $targetFolder! $($Error[0])" -fout
            }
        }else{
            try{
                Get-ChildItem $targetFolder | Remove-Item -Force -Confirm:$False -Recurse
            }catch{$Null}
        }
        $res = subst "$($convergedDriveletter)" $targetFolder
        labelDrive "$($convergedDriveletter)" $convergedDriveletter.SubString(0,1) $convergedDriveLabel
    }
}

while($true){
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
        $label1 = New-Object system.Windows.Forms.Label
        $label1.text=$progressBarText
        $label1.Name = "label1"
        $label1.Left=0
        $label1.Top= 9
        $label1.Width= $width
        $label1.Height=17
        $label1.Font= "Verdana"
        # create label
        $label2 = New-Object system.Windows.Forms.Label
        $label2.Name = "label2"
        $label2.Left=0
        $label2.Top= 0
        $label2.Width= $width
        $label2.Height=7
        $label2.backColor= $progressBarColor

        #add the label to the form
        $form1.controls.add($label1) 
        $form1.controls.add($label2) 
        $progressBar1 = New-Object System.Windows.Forms.ProgressBar
        $progressBar1.Name = 'progressBar1'
        $progressBar1.Value = 0
        $progressBar1.Style="Continuous" 
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = $width
        $System_Drawing_Size.Height = 10
        $progressBar1.Size = $System_Drawing_Size   
        
        $progressBar1.Left = 0
        $progressBar1.Top = 29
        $form1.Controls.Add($progressBar1)
        $form1.Show()| out-null  
        $form1.Focus() | out-null 
        $progressbar1.Value = 10
        $form1.Refresh()
    }

    #load edge driver
    try{
        $edgeDriverPath = "$($driversLocation)\msedgedriver.exe"
        log -text "Loading Edge driver $edgeDriverPath"
        if(!(Test-Path -Path $edgeDriverPath)){
            log -text "Edge driver not present at $edgeDriverPath, will try to download automatically from trusted source" -warning
            $autoUpdateEdgeDriver = $True
            $curEdgeDriverVersion = $Null
        }else{
            $curEdgeDriverVersion = (Get-ItemProperty -Path $edgeDriverPath).VersionInfo.ProductVersion
            log -text "Discovered Edge driver v$($curEdgeDriverVersion) at $edgeDriverPath"
        }

        if($autoUpdateEdgeDriver){
            if($edgeVersion){
                if($curEdgeDriverVersion -and $edgeVersion.Split(".")[0] -eq $curEdgeDriverVersion.Split(".")[0]){
                    log -text "Your Edge version and Edge Driver version match"
                }else{
                    log -text "Your Edge version ($edgeVersion) and Edge Driver version ($curEdgeDriverVersion) do not match, attempting to auto update"
                    $edgeDriverDownloadUrl = "https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/$edgeVersion/edgedriver_win64.zip"
                    log -text "Downloading latest version from $edgeDriverDownloadUrl to $driversLocation"
                    $tempZipPath = Join-Path $ENV:TEMP -ChildPath "edgedriver_win64.zip"
                    Invoke-WebRequest -uri $edgeDriverDownloadUrl -OutFile $tempZipPath -Method Get -UseBasicParsing -ErrorAction Stop
                    Expand-Archive -Path $tempZipPath -DestinationPath $driversLocation -Force -ErrorAction Stop
                    Remove-Item $tempZipPath -Force
                    Remove-Item "$($driversLocation)\Driver_Notes" -Recurse -Force -ErrorAction SilentlyContinue
                    log -text "Updated to version $latestEdgeDriverVersion from $edgeDriverDownloadUrl to $driversLocation"
                }
            }else{
                log -text "We don't know the version of your local Edge installation, so will not attempt to auto update the Edge driver" -warning
            }
        }
        $driverBlocked = $Null
        try{
            $driverBlocked = Get-Item $edgeDriverPath -Stream "Zone.Identifier" -ErrorAction Stop
        }catch{
            log -text "Edge driver present and not blocked by zoning, loading..."
        }
        if($driverBlocked){
            log -text "Edge driver was downloaded from the internet, so we need to run Unblock-File"
            try{
                Unblock-File -Path $edgeDriverPath -Confirm:$False
                log -text "Edge driver automatically unblocked"
            }catch{
                Throw "Edge driver not trusted by windows OS, right click msedgedriver.exe and unblock it in Properties or run Unblock-File"
            }
        }      
       
        #update progress bar
        if($showProgressBar) {
            $progressbar1.Value = 15
            $form1.Refresh()
        }
        $global:edgeOptions = [OpenQA.Selenium.Edge.EdgeOptions]::new()
        $global:edgeOptions.AddAdditionalOption("useAutomationExtension",$False)
        $global:edgeOptions.AddExcludedArgument("enable-automation")
        $global:edgeOptions.addArguments("user-data-dir=$($Env:appdata)\Lieben Consultancy\OnedriveMapper\Profile")
        $global:edgeOptions.addArguments("proxy-server='direct://'")
        $global:edgeOptions.addArguments("proxy-bypass-list=*")
        $global:edgeOptions.addArguments("disk-cache-size=262144")
        $global:edgeOptions.addArguments("--user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.71 Safari/537.36 Edge/12.0 OneDriveMapper/$version")
        if($forceHideEdge){
            $global:edgeOptions.addArguments("--headless=new")
        }         
        $global:edgeDriverService = [OpenQA.Selenium.Edge.EdgeDriverService]::CreateDefaultService($driversLocation,"msedgedriver.exe")
        $global:edgeDriverService.HideCommandPromptWindow = $true
        $global:edgeDriver = [OpenQA.Selenium.Edge.EdgeDriver]::new($global:edgeDriverService,$global:edgeOptions)
        $global:edgeDriver.Manage().Window.Size = [System.Drawing.Size]::New(600,600)
        $global:edgeDriver.Manage().Window.Position = [System.Drawing.Point]::new(([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width-600)/2,([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height-600)/2)
        log -text "Edge loaded successfully"
    }catch{
        log -text "Failed to load Edge driver, cannot continue.  error details: $($Error[0])" -fout
        run-CleanUp
        Exit
    }

    #update progress bar
    if($showProgressBar) {
        $progressbar1.Value = 20
        $form1.Refresh()
    }

    #cache HWND's of the new Edge window and then hide it until we need user-input
    try{
        $global:cachedHwnds = (Get-Process -ErrorAction SilentlyContinue -Id (gwmi win32_process | ? parentprocessid -eq $((gwmi win32_process | ? {$_.parentprocessid -eq $PID -and $_.name -eq "msedgedriver.exe"})).ProcessId).ProcessId).MainWindowHandle | Where-Object{$_ -ne 0}
    }catch{
        log -text "Failed to cache Edge Window Handles $($Error[0])" -fout
    }

    #update progress bar
    if($showProgressBar) {
        $progressbar1.Value = 25
        $form1.Refresh()
    }

    #hide edge window, sometimes needs extra attempts
    $attempts = 0
    while($true){
        if($attempts -gt 5){break}
        try{
            $res = $global:cachedHwnds | % {[native.win]::ShowWindow($_,0)}
            break
        }catch{
            $attempts++
            Start-Sleep -s 1
        }
    }

    #update progress bar
    if($showProgressBar) {
        $progressbar1.Value = 30
        $form1.Refresh()
    }

    #navigate to the o365 login url
    try{ 
        $global:edgeDriver.Navigate().GoToURL($o365loginURL)
    }catch{ 
        log -text "Failed to browse to the Office 365 Sign in page, this is a fatal error $($Error[0])`n" -fout
        $errorsForUser += "Mapping cannot continue because we could not contact Office 365`n"
        run-CleanUp
        Exit
    } 

    #update progress bar
    if($showProgressBar) {
        $progressbar1.Value = 35
        $form1.Refresh()
    }

    #generate cookies
    for($count=0;$count -lt $intendedMappings.Count;$count++){
        #update progress bar
        if($showProgressBar -and $script:progressbar1.Value -lt 90) {
            $script:progressbar1.Value += 5
            $script:form1.Refresh()
        }
        if($intendedMappings[$count].mapped){continue}
        if($intendedMappings[$count].sourceLocationPath -eq "autodetect"){
            $timeSpent = 0
            while($global:edgeDriver.Url.IndexOf("/personal/") -eq -1){
                Start-Sleep -s 2
                $timeSpent+=2
                log -text "Attempting to detect username at $($global:edgeDriver.Url), waited for $timeSpent seconds" 
                $global:edgeDriver.Navigate().GoToUrl($baseURL)
                start-AuthCheck
                if($timeSpent -gt 60){
                    log -text "Failed to get the username from the URL for over $timeSpent seconds while at $($global:edgeDriver.Url), aborting" -fout 
                    $errorsForUser += "Mapping cannot continue because we cannot detect your username`n"
                    run-CleanUp
                    Exit 
                }
            }
            try{
                $start = $global:edgeDriver.Url.IndexOf("/personal/")+10 
                $end = $global:edgeDriver.Url.IndexOf("/",$start) 
                $userURL = $global:edgeDriver.Url.Substring($start,$end-$start).Replace("%27","'")
                $mapURL = $mapURLpersonal + $userURL + "\" + $libraryName 
            }catch{
                log -text "Failed to get the username while at $($global:edgeDriver.Url), aborting" -fout
                $errorsForUser += "Mapping cannot continue because we cannot detect your username`n"
                run-CleanUp
                Exit 
            }
            $intendedMappings[$count].webDavPath = $mapURL 
            log -text "Detected user: $($userURL)"
            log -text "Onedrive cookie generated"
            add-cookies
        }else{
            log -text "Initiating Sharepoint session with: $($intendedMappings[$count].sourceLocationPath)"
            $spURL = $intendedMappings[$count].sourceLocationPath #URL to browse to
            if($spURL.IndexOf($privateSuffix) -ne -1){ #hack to make mappings to other user's Onedrive work
                $spURL = $spURL.Substring(0,$spURL.LastIndexOf("/"))
            }
            log -text "Current location: $($global:edgeDriver.Url)" 
            $global:edgeDriver.Navigate().GoToUrl($spURL) #check the URL
            $waited = 0
            while(!$($global:edgeDriver.Url.StartsWith("https://$($O365CustomerName).sharepoint.com"))){
                start-AuthCheck
                Start-Sleep -s 1
                $waited++

                log -text "waited $waited seconds to load $spURL, currently at $($global:edgeDriver.Url)"
                if($waited -ge $maxWaitSecondsForSpO){
                    log -text "waited longer than $maxWaitSecondsForSpO seconds to load $spURL! This mapping may fail" -fout
                    break
                }
            }
            log -text "Current location: $($global:edgeDriver.Url)" 
            log -text "SpO cookie generated"
            add-cookies
        }
    }

    for($count=0;$count -lt $intendedMappings.Count;$count++){
        #map the drive
        $intendedMappings[$count].mapped = MapDrive $intendedMappings[$count]

        if($intendedMappings[$count].sourceLocationPath -eq "autodetect"){       
            if($addShellLink -and $windowsVersion -eq 6 -and $intendedMappings[$count].targetLocationType -eq "driveletter" -and [System.IO.Directory]::Exists($intendedMappings[$count].targetLocationPath)){
                try{
                    $res = createFavoritesShortcutToO4B -targetLocation $intendedMappings[$count].targetLocationPath
                }catch{
                    log -text "Failed to create a shortcut to the mapped drive for Onedrive for Business because of: $($Error[0])" -fout
                }
            }
        }
    }

    #update progress bar
    if($showProgressBar) {
        $script:progressbar1.Value = 95
        $script:form1.Refresh()
    }

    if($redirectFolders){
        $listOfFoldersToRedirect | ForEach-Object {
            log -text "Redirecting $($_.knownFolderInternalName) to $($_.desiredTargetPath)"
            try{
                Redirect-Folder -GetFolder $_.knownFolderInternalName -SetFolder $_.knownFolderInternalIdentifier -Target $_.desiredTargetPath -copyExistingFiles $_.copyExistingFiles
                log -text "Redirected $($_.knownFolderInternalName) to $($_.desiredTargetPath)"
            }catch{
                log -text "Failed to redirect $($_.knownFolderInternalName) to $($_.desiredTargetPath): $($Error[0])" -fout
            }
        }
    }

    if($showProgressBar) {
        $progressbar1.Value = 100
        $label1.text="Done!"
        $script:form1.Refresh()
        Start-Sleep -Milliseconds 1000
    }      

    run-CleanUp

    if($autoRemapMethod -ne "Disabled"){
        if(@($intendedMappings | where {$_.mapped}).count -gt 0){
            log "autoRemapMethod is set to $autoRemapMethod, OnedriveMapper will continue to monitor your mappings and remap if they get disconnected"
        }else{
            log "autoRemapMethod is set to $autoRemapMethod, but all mappings failed, OnedriveMapper will exit" -fout
            break
        }
        :escape while($true){
            for($count=0;$count -lt $intendedMappings.Count;$count++){
                $mappingConnected = $False
                if(($autoRemapMethod -eq "Path" -and !(Test-Path $intendedMappings[$count].webDavPath))){
                    Write-Host "UNHEALTHY: $($intendedMappings[$count].webDavPath)" -ForegroundColor Red
                }elseif($autoRemapMethod -eq "Link" -and $intendedMappings[$count].targetLocationType -eq "networklocation" -and !(Test-Path (Join-Path $intendedMappings[$count].targetLocationPath -ChildPath "$($intendedMappings[$count].displayName).lnk"))){
                    Write-Host "UNHEALTHY: $($intendedMappings[$count].targetLocationPath)\$($intendedMappings[$count].displayName).lnk" -ForegroundColor Red
                }elseif($autoRemapMethod -eq "Link" -and $intendedMappings[$count].targetLocationType -eq "driveletter" -and !(Test-Path $intendedMappings[$count].targetLocationPath)){
                    Write-Host "UNHEALTHY: $($intendedMappings[$count].targetLocationPath)" -ForegroundColor Red
                }elseif($autoRemapMethod -eq "Link" -and $intendedMappings[$count].targetLocationType -eq "converged" -and !(Test-Path (Join-Path $intendedMappings[$count].targetLocationPath -ChildPath $($intendedMappings[$count].displayName)))){
                    Write-Host "UNHEALTHY: $($intendedMappings[$count].targetLocationPath)\$($intendedMappings[$count].displayName)" -ForegroundColor Red
                }else{
                    $mappingConnected = $True
                    Write-Host "HEALTHY: $($intendedMappings[$count].webDavPath) " -ForegroundColor Green
                }
                if(!$mappingConnected){
                    log "autoRerun is set to True and $($intendedMappings[$count].displayName) seems to be unavailable...checking internet connectivity"
                    $internetConnectivity = $False
                    $internetConnectivity = (test-connection 8.8.8.8 -Count 1 -Quiet)
                    if(!$internetConnectivity){
                        log "Internet connectivity to 8.8.8.8 failed, waiting..." -fout
                        Start-Sleep 5
                        break
                    }else{
                        log "Internet connectivity to 8.8.8.8 tested positive"
                        $intendedMappings[$count].mapped = $False
                        Start-Sleep -s 2
                        break escape
                    }                    
                }else{
                    $secs = (Get-Random -Minimum 5 -Maximum 20)
                    Write-Host "Sleeping for $secs seconds" -ForegroundColor Green
                    Start-Sleep -s $secs
                } 
            }
        }
        log "autoRemap triggered, closing and reconnecting broken mappings"
    }else{
        break
    }
}

log -text "OnedriveMapper has finished running"
Exit