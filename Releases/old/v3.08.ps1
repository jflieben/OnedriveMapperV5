######## 
#OneDriveMapper
#Copyright:         Free to use, please leave this header intact 
#Author:            Jos Lieben (OGD)
#Company:           OGD (http://www.ogd.nl) 
#Script help:       http://www.lieben.nu, please provide a decrypted Fiddler Trace Log if you're using Native auth and having issues
#Purpose:           This script maps Onedrive for Business and maps a configurable number of Sharepoint Libraries
#Enterprise users:  A special version of OnedriveMapper per MSI (or ps1) is available that allows central (re)configuration of all options set below
#Enterprise users:  In addition, the cloud version can automatically update itself and you can centrally switch the authentication methode without a new rollout
#Enterprise users:  Finally, enterprise users receive email notifications of new versions

#TODO:
#explorer restart only if logon process complete? https://gallery.technet.microsoft.com/scriptcenter/Analyze-Session-Logon-63e02691
#optionally, use PSDrive
#AzureADSSO for IE retest
#handling user based distribution vs device based
#sso for AzureAD Joined devices using native method
#explanation video of all settings
#optionally don't display version check information
#adfs login velden robuuster opzoeken
#http://stackoverflow.com/questions/7530734/creating-a-cookie-outside-of-a-web-browser-e-g-with-vbscript
#decrypt stored password on different pc's

param(
    [Switch]$asTask,
    [Switch]$fallbackMode
)

######## 
#Configuration 
######## 
$version = "3.08"
$configurationID       = "00000000-0000-0000-0000-000000000000"#Don't modify this, unless you are using OnedriveMapper Cloud edition

###If you set a ConfigurationID and are using OnedriveMapper Cloud, no further configuration is required. If you're not using OnedriveMapper Cloud, please finish below configuration.

$authMethod            = "native"                  #Uses IE automation (old method) when set to ie, uses new native method when set to 'native'
$allowFallbackMode     = $True                     #if set to True, and the selected authentication method fails, onedrivemapper will try again using the other authentication method
$domain                = "liebensraum.NL"          #This should be your domain name in O365, and your UPN in Active Directory, for example: ogd.nl 
$driveLetter           = "X:"                      #This is the driveletter you'd like to use for OneDrive, for example: Z: 
$redirectMyDocs        = $False                    #will redirect mydocuments to the mounted drive if set to $True, does not properly 'undo' when disabled after being enabled
$redirectDesktop       = $False
$redirectFavorites     = $False
$redirectToSubfolderName  = "Documents"               #This is the folder to which we will redirect under the given $driveletter, leave empty to redirect to the Root (may cause odd labels for special folders in Windows)
$driveLabel            = "onedrive"                #If you enter a name here, the script will attempt to label the drive with this value 
$O365CustomerName      = "onedrivemapper"          #This should be the name of your tenant (example, ogd as in ogd.onmicrosoft.com) 
$logfile               = ($env:APPDATA + "\OneDriveMapper_$version.log")    #Logfile to log to 
$pwdCache              = ($env:APPDATA + "\OneDriveMapper.tmp")    #file to store encrypted password into, change to $Null to disable
$loginCache            = ($env:APPDATA + "\OneDriveMapper.tmp2")    #file to store encrypted login into, change to $Null to disable
$settingsCache         = ($env:APPDATA + "\OneDriveMapper.cache")    #file to store encrypted settings in case server isn't reachable, change to $Null to disable
$dontMapO4B            = $False                    #If you're only using Sharepoint Online mappings (see below), set this to True to keep the script from mapping the user's O4B
$addShellLink          = $False                    #Adds a link to Onedrive to the Shell under Favorites (Windows 7, 8 / 2008R2 and 2012R2 only) If you use a remote path, google EnableShellShortcutIconRemotePath
$deleteUnmanagedDrives = $True                     #If set to $True, OnedriveMapper checks if there are 'other' mapped drives to Sharepoint Online/Onedrive that OnedriveMapper does not manage, and disconnects them. This is useful if you change a driveletter.
$debugmode             = $False                    #Set to $True for debugging purposes. You'll be able to see the script navigate in Internet Explorer if you're using IE auth mode
$userLookupMode        = 1                         #1 = Active Directory UPN, 2 = Active Directory Email, 3 = Azure AD Joined Windows 10, 4 = query user for his/her login, 5 = lookup by registry key, 6 = display full form (ask for both username and login if no cached versions can be found)
$AzureAADConnectSSO    = $False                    #NOT NEEDED FOR NATIVE AUTH, if set to True, will automatically remove AzureADSSO registry key before mapping, and then readd them after mapping. Otherwise, mapping fails because AzureADSSO creates a non-persistent cookie
$lookupUserGroups      = $False                    #Set this to $True if you want to map user security groups to Sharepoint Sites (read below for additional required configuration)
$forceUserName         = ''                        #if anything is entered here, userLookupMode is ignored
$forcePassword         = ''                        #if anything is entered here, the user won't be prompted for a password. This function is not recommended, as your password could be stolen from this file 
$restartExplorer       = $False                    #Set to $True if you're having any issues with drive visibility
$autoProtectedMode     = $True                     #Automatically temporarily disable IE Protected Mode if it is enabled. ProtectedMode has to be disabled for the script to function 
$adfsWaitTime          = 10                        #Amount of seconds to allow for SSO (ADFS or AzureAD or any other configured SSO provider) redirects, if set too low, the script may fail while just waiting for a slow redirect, this is because the IE object will report being ready even though it is not.  Set to 0 if using passwords to sign in.
$libraryName           = "Documents"               #leave this default, unless you wish to map a non-default library you've created 
$autoKillIE            = $True                     #Kill any running Internet Explorer processes prior to running the script to prevent security errors when mapping 
$abortIfNoAdfs         = $False                    #If set to True, will stop the script if no ADFS server has been detected during login
$adfsMode              = 1                         #1 = use whatever came out of userLookupMode, 2 = use only the part before the @ in the upn
$adfsSmartLink         = $Null                     #If set, the ADFS smartlink will be used to log in to Office 365. For more info, read the FAQ at http://http://www.lieben.nu/liebensraum/onedrivemapper/onedrivemapper-faq/
$displayErrors         = $True                     #show errors to user in visual popups
$persistentMapping     = $True                     #If set to $False, the mapping will go away when the user logs off
$buttonText            = "Login"                   #Text of the button on the password input popup box
$loginformIntroText    = "Welcome to COMPANY NAME`r`nPlease enter your login and password" #used as introduction text when you set userLookupMode to 6
$loginFieldText        = "Please enter your login in the form of xxx@xxx.com" #used as label above the login text field when you set userLookupMode to 6
$passwordFieldText     = "Please enter your password" #used as label above the password text field when you set userLookupMode to 6
$adfsLoginInput        = "userNameInput"           #change to user-signin if using Okta, username2Txt if using RMUnify
$adfsPwdInput          = "passwordInput"           #change to pass-signin if using Okta, passwordTxt if using RMUnify
$adfsButton            = "submitButton"            #change to singin-button if using Okta, Submit if using RMUnify
$urlOpenAfter          = ""                        #This URL will be opened by the script after running if you configure it
$showConsoleOutput     = $True                     #Set this to $False to hide console output
$showElevatedConsole   = $True
$sharepointMappings    = @()
$sharepointMappings    += "https://ogd.sharepoint.com/site1/documentsLibrary,ExampleLabel,Y:"
$showProgressBar       = $True                     #will show a progress bar to the user
$versionCheck          = $True                     #will check if running the latest version, if not, this will be logged to the logfile, no personal data is transmitted.
$autoDetectProxy       = $False                    #if set to $False, unchecks the 'Automatically detect proxy settings' setting in IE; this greatly enhanced WebDav performance, set to true to not modify this IE setting (leave as is)
#for each sharepoint site you wish to map 3 comma seperated values are required, the 'clean' url to the library (see example), the desired drive label, and the driveletter
#if you wish to add more, copy the example as you see above, if you don't wish to map any sharepoint sites, simply leave as is

if($showConsoleOutput -eq $False){
    $t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    try{
        add-type -name win -member $t -namespace native
        [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
    }catch{$Null}
}

######## 
#Required resources, it's highly unlikely you need to change any of this
######## 
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
$mapresult = $False 
$protectedModeValues = @{} 
$privateSuffix = "-my"
$script:errorsForUser = ""
$userLoginRegistryKey = "HKCU:\System\CurrentControlSet\Control\CustomUID"
$onedriveIconPath = "C:\GitRepos\OnedriveMapper\onedrive.ico" #if this file exists, and you've set addShellLink to True, it will be used as icon for the shortcut
$i_MaxLocalLogSize = 2 #max local log size in MB
$maxWaitSecondsForSpO  = 5                        #Maximum seconds the script waits for Sharepoint Online to load before mapping
if($adfsSmartLink){
    $o365loginURL = $adfsSmartLink
}else{
    $o365loginURL = "https://login.microsoftonline.com/login.srf?msafed=0"
}
if($sharepointMappings[0] -eq "https://ogd.sharepoint.com/site1/documentsLibrary,ExampleLabel,Y:"){           ##DO NOT CHANGE THIS
    $sharepointMappings = @()
}

$domain = $domain.ToLower() 
$debugInfo = $Null
$O365CustomerName = $O365CustomerName.ToLower() 
#for people that don't RTFM, fix wrongly entered customer names:
$O365CustomerName = $O365CustomerName -Replace ".onmicrosoft.com",""
$forceUserName = $forceUserName.ToLower() 
$finalURLs = @()
$finalURLs += "https://portal.office.com"
$finalURLs += "https://outlook.office365.com"
$finalURLs += "https://outlook.office.com"
$finalURLs += "https://$($O365CustomerName)-my.sharepoint.com"
$finalURLs += "https://$($O365CustomerName).sharepoint.com"
$finalURLs += "https://www.office.com"

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
        ac $logfile "$(Get-Date) | $text"
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
if($lookupUserGroups -and $configurationID -eq "00000000-0000-0000-0000-000000000000"){
    try{
        $groups = ([ADSISEARCHER]"samaccountname=$($env:USERNAME)").Findone().Properties.memberof -replace '^CN=([^,]+).+$','$1'
        log -text "cached user group membership because lookupUserGroups was set to True"
        #####################FOR EACH GROUP YOU WISH TO MAP TO A SHAREPOINT LIBRARY, UNCOMMENT AND REPEAT BELOW EXAMPLE, NOTE: THIS MAY FAIL IF THERE ARE REGEX CHARACTERS IN THE NAME
        #    $group = $groups -contains "DLG_West District School A - Sharepoint"
        #    if($group){
        #       ###REMEMBER, THE BELOW LINE SHOULD CONTAIN 2 COMMA's to distinguish between URL, LABEL and DRIVELETTER
        #       $sharepointMappings += "https://ogd.sharepoint.com/district_west/DocumentLibraryName,West District,Y:"
        #       log -text "adding a sharepoint mapping because the user is a member of $group"
        #    }  
    }catch{
        log -text "failed to cache user group membership because of: $($Error[0])" -fout
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

function getElementById{
    Param(
        [Parameter(Mandatory=$true)]$id
    )
    $localObject = $Null
    try{
        $localObject = $script:ie.document.getElementById($id)
        if($localObject.tagName -eq $Null){Throw "The element $id was not found (1) or had no tagName"}
        return $localObject
    }catch{$localObject = $Null}
    try{
        $localObject = $script:ie.document.IHTMLDocument3_getElementById($id)
        if($localObject.tagName -eq $Null){Throw "The element $id was not found (2) or had no tagName"}
        return $localObject
    }catch{
        Throw
    }
}

function ConvertTo-Json20([object] $item){
    add-type -assembly system.web.extensions
    $ps_js=new-object system.web.script.serialization.javascriptSerializer
    return $ps_js.Serialize($item)
}

function ConvertFrom-Json20([object] $item){ 
    add-type -assembly system.web.extensions
    $ps_js=new-object system.web.script.serialization.javascriptSerializer

    #The comma operator is the array construction operator in PowerShell
    return ,$ps_js.DeserializeObject($item)
}

function Create-Cookie{
    Param(
        $name,
        $value,
        $domain,
        $path="/",
        $HttpOnly=$True,
        $Secure=$True,
        $Expires
    )
    $c=New-Object System.Net.Cookie;
    $c.Name=$name;
    $c.Path=$path;
    $c.Value = $value
    $c.Domain =$domain;
    $c.HttpOnly = $HttpOnly;
    $c.Secure = $Secure;
    if($Expires){
        $c.Expires = $Expires
    }
    return $c;
}

function JosL-WebRequest{
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
            $request.TimeOut = 30000
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
            $request.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E)"
            $request.ContentType = $contentType
            $request.CookieContainer = $script:cookiejar
            $script:debugInfo += "JOSL-REQUEST "
            $script:debugInfo += $request.Method
            $script:debugInfo += $url
            $script:debugInfo += "`r"
            $script:debugInfo += $body
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
            if($PSVersionTable.PSVersion.Major -le 2){
                $cookies = $response.Headers["Set-Cookie"]
                if($cookies){
                    $marker = 0
                    while($true){
                        if($cookies.IndexOf("expires=",$marker) -ne -1){
                            $start = $cookies.IndexOf("expires=",$marker)
                            $end = $cookies.IndexOf(",",$start)-$start+1
                            if($end -gt 0){
                                $replaceFrom = $cookies.SubString($start,$end)
                                $replaceTo = "$($replaceFrom.SubString(0,$replaceFrom.Length-1))|JOS|"
                                $cookies = $cookies.Replace($replaceFrom,$replaceTo)
                            }
                        }else{break}
                        $marker=$start+10                        
                    }
                    
                    $cookies = $cookies.Split(",")
                    if($cookies.Count -gt 0){
                        foreach($cookie in $cookies){
                            $cookieParts = $cookie.Split(";")
                            $index = 0
                            $name = $Null
                            $value = $Null
                            $path = $null
                            $secure = $Null
                            $domain = $Null
                            $expires = $Null
                            foreach($part in $cookieParts){
                                if($index -eq 0){
                                    $name = $part.Split("=")[0]
                                    $value = $part.Split("=")[1]
                                }
                                if($part.Split("=")[0] -eq "path"){
                                    $path = $part.Split("=")[1]
                                }
                                if($part.Split("=")[0] -eq "expires"){
                                    $expires = $part.Split("=")[1].Replace("|JOS|",",")
                                }
                                if($part.Split("=")[0] -eq "secure"){
                                    $secure = $True
                                } 
                                if($part.Split("=")[0] -eq "domain"){
                                    $domain = $part.Split("=")[1]
                                    if($domain.IndexOf(".") -eq 0){
                                        $domain = $domain.SubString(1)
                                    }
                                }else{
                                    $domain = $request.RequestUri.host
                                }                                        
                                $index++
                            }
                            $script:cookiejar.Add((Create-Cookie -name $name -value $value -domain $domain -path $path -HttpOnly $True -Secure $True -Expires $expires))
                        }
                    }
                }
            }
            $retVal.Headers = $response.Headers
            $stream = $response.GetResponseStream()
            $streamReader = [System.IO.StreamReader]($stream)
            $retVal.Content = $streamReader.ReadToEnd()
            $script:debugInfo += "JOSL-RESPONSE "
            $script:debugInfo += $retVal.StatusCode
            $script:debugInfo += " $($response.ResponseUri )"
            $script:debugInfo += "`r"
            $script:debugInfo += $retVal.Content
            $streamReader.Close()
            $response.Close()
            $response = $Null
            $request = $Null

            return $retVal
        }catch{
            if($attempts -ge $maxAttempts){Throw}else{sleep -s 2}
        }
    }
}

function returnEnclosedFormValue{
    Param(
        $res,
        $searchString,
		$endString = "`"",
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
        if($searchLength -eq $startLoc-1){
            return -1
        }
        if($decode){
            return([System.Web.HttpUtility]::UrlDecode($res.Content.Substring($startLoc,$searchLength)))
        }else{
            return($res.Content.Substring($startLoc,$searchLength))
        }
    }catch{Throw}
}

function handleAzureADConnectSSO{
    Param(
        [Switch]$initial
    )
    $failed = $False
    if($script:AzureAADConnectSSO -and $authMethod -ne "native"){
        if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon" -entryName "https") -eq 1){
            log -text "ERROR: https://autologon.microsoftazuread-sso.com found in IE Local Intranet sites, Azure AD Connect SSO is only supported if you let OnedriveMapper set the registry keys! Don't set this site through GPO" -fout
            $failed = $True
        }
        if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nsatc.net\aadg.windows.net" -entryName "https") -eq 1){
            log -text "ERROR: https://aadg.windows.net.nsatc.net found in IE Local Intranet sites, Azure AD Connect SSO is only supported if you let OnedriveMapper set the registry keys! Don't set this site through GPO" -fout
            $failed = $True
        } 
        if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon" -entryName "https") -eq 1){
            log -text "ERROR: https://autologon.microsoftazuread-sso.com found in IE Local Intranet sites, Azure AD Connect SSO is only supported if you let OnedriveMapper set the registry keys! Don't set this site through GPO" -fout
            $failed = $True
        }
        if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nsatc.net\aadg.windows.net" -entryName "https") -eq 1){
            log -text "ERROR: https://aadg.windows.net.nsatc.net found in IE Local Intranet sites, Azure AD Connect SSO is only supported if you let OnedriveMapper set the registry keys! Don't set this site through GPO" -fout
            $failed = $True
        } 
        if($failed -eq $False){
            if($initial){
                #check AzureADConnect SSO intranet sites
                if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon" -entryName "https") -eq 1){
                    $res = remove-item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon"    
                    log -text "Automatically removed autologon.microsoftazuread-sso.com from intranet sites for this user"
                }
                if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nsatc.net\aadg.windows.net" -entryName "https") -eq 1){
                    $res = remove-item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nsatc.net\aadg.windows.net" 
                    log -text "Automatically removed aadg.windows.net.nsatc.net from intranet sites for this user"   
                }
            }else{
                #log results, try to automatically add trusted sites to user trusted sites if not yet added
                if((addSiteToIEZoneThroughRegistry -siteUrl "aadg.windows.net.nsatc.net" -mode 1) -eq $True){log -text "Automatically added aadg.windows.net.nsatc.net to intranet sites for this user"}
                if((addSiteToIEZoneThroughRegistry -siteUrl "autologon.microsoftazuread-sso.com" -mode 1) -eq $True){log -text "Automatically added autologon.microsoftazuread-sso.com to intranet sites for this user"}   
            }
        }
    }
}

function storeSecureString{
    Param(
        $filePath,
        $string
    )
    try{
        $stringForFile = $string | ConvertTo-SecureString -AsPlainText -Force -ErrorAction Stop | ConvertFrom-SecureString -ErrorAction Stop
        $res = Set-Content -Path $filePath -Value $stringForFile -Force -ErrorAction Stop
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

function versionCheck{
    Param(
        $currentVersion
    )
    $apiURL = "http://om.lieben.nu/lieben_api.php?script=OnedriveMapper&version=$currentVersion"
    $apiKeyword = "latestOnedriveMapperVersion"
    try{
        $result = JosL-WebRequest -Url $apiURL
    }catch{
        Throw "Failed to connect to API url for version check: $apiURL $($Error[0])"
    }
    try{
        $keywordIndex = $result.Content.IndexOf($apiKeyword)
        if($keywordIndex -lt 1){
            Throw ""
        }
    }catch{
        Throw "Connected to API url for version check, but invalid API response"
    }
    $latestVersion = $result.Content.SubString($keywordIndex+$apiKeyword.Length+1,4)
    if($latestVersion -ne $currentVersion){
        Throw "OnedriveMapper version mismatch, current version: v$currentVersion, latest version: v$latestVersion"
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
        sleep -s 5
        if((Get-Service -Name WebClient).status -eq "Running"){
            log -text "detected that the webdav client is now running!"
        }else{
            log -text "but the webdav client is still not running! Please set the client to automatically start!" -fout
        }
    }catch{
        Throw "Failed to start the webdav client :( $($Error[0])"
    }
}

function Pause{
   Read-Host 'Press Enter to continue...' | Out-Null
}

function storeSettingsToCache{
    Param(
        [Parameter(Mandatory=$true)]$settingsCache,
        [Parameter(Mandatory=$true)]$settings
    )
    try{
        Export-Clixml -Depth 6 -Path $settingsCache -InputObject $settings -Force -Encoding UTF8 -Confirm:$False -ErrorAction Stop
    }catch{
        Throw
    }
}

function queryForAllCreds {
    Param(
        [Parameter(Mandatory=$true)]$introText,
        [Parameter(Mandatory=$true)]$buttonText,
        [Parameter(Mandatory=$true)]$loginLabel,
        [Parameter(Mandatory=$true)]$passwordLabel
    )
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
    $userForm.Text = "OnedriveMapper" 
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

function retrieveSettingsFromCache{
    Param(
        [Parameter(Mandatory=$true)]$settingsCache
    )
    try{
        $settings = Import-Clixml -Path $settingsCache -ErrorAction Stop
        return $settings
    }catch{
        Throw
    }
}

function redirectMyDocuments{
    Param(
        $driveLetter
    )
    $dl = "$($driveLetter)\"
    $myDocumentsNewPath = Join-Path -Path $dl -ChildPath $redirectToSubfolderName
    $myDesktopNewPath = Join-Path -Path $myDocumentsNewPath -ChildPath "Desktop"
    $myPicturesNewPath = Join-Path -Path $myDocumentsNewPath -ChildPath "Pictures"
    $myVideosNewPath = Join-Path -Path $myDocumentsNewPath -ChildPath "Videos"
    $myMusicNewPath = Join-Path -Path $myDocumentsNewPath -ChildPath "Music"
    $myFavoritesNewPath = Join-Path -Path $myDocumentsNewPath -ChildPath "Favorites"
    $myDownloadsNewPath = Join-Path -Path $myDocumentsNewPath -ChildPath "Downloads"
    #create folders if necessary
    $waitedTime = 0    
    while($true){
        try{
            if(![System.IO.Directory]::Exists($myDocumentsNewPath)){
                $res = New-Item $myDocumentsNewPath -ItemType Directory -ErrorAction Stop
                Sleep -Milliseconds 200
            } 
            if(![System.IO.Directory]::Exists($myDesktopNewPath) -and $redirectDesktop){
                $res = New-Item $myDesktopNewPath -ItemType Directory -ErrorAction Stop
                Sleep -Milliseconds 200
            }               
            if(![System.IO.Directory]::Exists($myPicturesNewPath) -and $redirectMyDocs){
                $res = New-Item $myPicturesNewPath -ItemType Directory -ErrorAction Stop
                Sleep -Milliseconds 200
            } 
            if(![System.IO.Directory]::Exists($myVideosNewPath) -and $redirectMyDocs){
                $res = New-Item $myVideosNewPath -ItemType Directory -ErrorAction Stop
                Sleep -Milliseconds 200
            } 
            if(![System.IO.Directory]::Exists($myMusicNewPath) -and $redirectMyDocs){
                $res = New-Item $myMusicNewPath -ItemType Directory -ErrorAction Stop
            } 
            if(![System.IO.Directory]::Exists($myFavoritesNewPath) -and $redirectFavorites){
                $res = New-Item $myFavoritesNewPath -ItemType Directory -ErrorAction Stop
            }
            if(![System.IO.Directory]::Exists($myDownloadsNewPath) -and $redirectMyDocs){
                $res = New-Item $myDownloadsNewPath -ItemType Directory -ErrorAction Stop
            }
            break
        }catch{
            sleep -s 2
            $waitedTime+=2
            if($waitedTime -gt 15){
                log -text "Failed to redirect document libraries because we could not create folders in the target path $dl $($Error[0])" -fout
                return $False              
            }      
        }
    }
    try{
        log -text "Retrieving current document library configuration"
        $lib = "$Env:appdata\Microsoft\Windows\Libraries\Documents.library-ms"
        $content = get-content -LiteralPath $lib
    }catch{
        log -text "Failed to retrieve document library configuration, will not be able to redirect $($Error[0])" -fout
        return $False
    }
    #Method 1 (works for Win7/8/2008R2)
    if($redirectMyDocs){
        try{
            $strip = $false
            $count = 0
            foreach($line in $content){
                if($line -like "*<searchConnectorDescriptionList>*"){$strip = $True}
                if($strip){$content[$count]=$Null}
                $count++
            }
            $content+="<searchConnectorDescriptionList>"
            $content+="<searchConnectorDescription>"
            $content+="<isDefaultSaveLocation>true</isDefaultSaveLocation>"
            $content+="<isSupported>false</isSupported>"
            $content+="<simpleLocation>"
            $content+="<url>$myDocumentsNewPath</url>"
            $content+="</simpleLocation>"
            $content+="</searchConnectorDescription>"
            $content+="</searchConnectorDescriptionList>"
            $content+="</libraryDescription>"
            Set-Content -Value $content -Path $lib -Force -ErrorAction Stop
            log -text "Modified $lib"
        }catch{
            log -text "Failed to redirect document library $($Error[0])" -fout
            return $False
        }
    }
    #Method 2 (Windows 10+)
    try{   
        if($redirectMyDocs){
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "Personal" -value $myDocumentsNewPath -ErrorAction Stop        
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "{F42EE2D3-909F-4907-8871-4C22FC0BF756}" -value $myDocumentsNewPath -ErrorAction Stop  
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "My Video" -value $myVideosNewPath -ErrorAction SilentlyContinue
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "{35286A68-3C57-41A1-BBB1-0EAE73D76C95}" -value $myVideosNewPath -ErrorAction SilentlyContinue
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "My Music" -value $myMusicNewPath -ErrorAction SilentlyContinue
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "{A0C69A99-21C8-4671-8703-7934162FCF1D}" -value $myMusicNewPath -ErrorAction SilentlyContinue
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "My Pictures" -value $myPicturesNewPath -ErrorAction SilentlyContinue
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "{0DDD015D-B06C-45D5-8C4C-F59713854639}" -value $myPicturesNewPath -ErrorAction SilentlyContinue
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}" -value $myDownloadsNewPath -ErrorAction SilentlyContinue
        }
        if($redirectFavorites){
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "Favorites" -value $myFavoritesNewPath -ErrorAction SilentlyContinue
        }
        if($redirectDesktop){
            $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "Desktop" -value $myDesktopNewPath -ErrorAction SilentlyContinue
        }
        
        log -text "Modified explorer shell registry entries"
    }catch{
        log -text "Failed to redirect document library $($Error[0])" -fout
        return $False
    }
    log -text "Redirection complete"
    return $True
}

function checkIfAtO365URL{
    param(
        [String]$url,
        [Array]$finalURLs
    )
    foreach($item in $finalURLs){
        if($url.StartsWith($item)){
            return $True
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
            $Null = New-Item -Path $path –Force -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_CommentFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel -ErrorAction SilentlyContinue
            $regURL = $regURL -Replace [regex]::escape("DavWWWRoot#"),"" 
            $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($regURL)" 
            $Null = New-Item -Path $path –Force -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_CommentFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel -ErrorAction SilentlyContinue
            log -text "Label has been set to $($lD_DriveLabel)" 
 
        }catch{ 
            log -text "Failed to set the drive label registry keys: $($Error[0]) " -fout
        } 
 
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

function fixElevationVisibility{
    #check if a task already exists for this script
    if($showElevatedConsole){
        $createTask = "schtasks /Create /SC ONCE /TN OnedriveMapper /IT /RL LIMITED /F /TR `"Powershell.exe -ExecutionPolicy ByPass -File '$scriptPath' -asTask`" /st 00:00"    
    }else{
        $createTask = "schtasks /Create /SC ONCE /TN OnedriveMapper /IT /RL LIMITED /F /TR `"Powershell.exe -ExecutionPolicy ByPass -WindowStyle Hidden -File '$scriptPath' -asTask`" /st 00:00"
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
        [String]$MD_DriveLetter, 
        [String]$MD_MapURL, 
        [String]$MD_DriveLabel 
    ) 
    $LASTEXITCODE = 0
    log -text "Mapping target: $($MD_MapURL)" 
    try{$del = NET USE $MD_DriveLetter /DELETE /Y 2>&1}catch{$Null}
    if($persistentMapping){
        try{$out = NET USE $MD_DriveLetter $MD_MapURL /PERSISTENT:YES 2>&1}catch{$Null}
    }else{
        try{$out = NET USE $MD_DriveLetter $MD_MapURL /PERSISTENT:NO 2>&1}catch{$Null}
    }
    if($out -like "*error 67*"){
        log -text "ERROR: detected string error 67 in return code of net use command, this usually means the WebClient isn't running" -fout
    }
    if($out -like "*error 224*"){
        log -text "ERROR: detected string error 224 in return code of net use command, this usually means your trusted sites are misconfigured or KB2846960 is missing" -fout
    }
    if($LASTEXITCODE -ne 0){ 
        log -text "Failed to map $($MD_DriveLetter) to $($MD_MapURL), error: $($LASTEXITCODE) $($out) $del" -fout
        $script:errorsForUser += "$MD_DriveLetter could not be mapped because of error $($LASTEXITCODE) $($out) d$del`n"
        return $False 
    } 
    if([System.IO.Directory]::Exists($MD_DriveLetter)){ 
        #set drive label 
        $Null = labelDrive $MD_DriveLetter $MD_MapURL $MD_DriveLabel
        log -text "$($MD_DriveLetter) mapped successfully`n" 
        if(($redirectMyDocs -or $redirectDesktop -or $redirectFavorites) -and $driveLetter -eq $MD_DriveLetter){
            $res = redirectMyDocuments -driveLetter $MD_DriveLetter
        }
        return $True 
    }else{ 
        log -text "failed to contact $($MD_DriveLetter) after mapping it to $($MD_MapURL), check if the URL is valid. Error: $($error[0]) $out" -fout
        return $False 
    } 
} 
 
function revertProtectedMode(){ 
    log -text "autoProtectedMode is set to True, reverting to old settings" 
    try{ 
        for($i=0; $i -lt 5; $i++){ 
            if($protectedModeValues[$i] -ne $Null){ 
                log -text "Setting zone $i back to $($protectedModeValues[$i])" 
                Set-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500"  -Value $protectedModeValues[$i] -Type Dword -ErrorAction SilentlyContinue 
            } 
        } 
    } 
    catch{ 
        log -text "Failed to modify registry keys to change ProtectedMode back to the original settings: $($Error[0])" -fout
    } 
} 

function abort_OM{ 
    if($showProgressBar) {
        $progressbar1.Value = 100
        $label1.text="Done!"
        Sleep -Milliseconds 500
        $form1.Close()
    }
    #find and kill all active COM objects for IE
    if($authMethod -ne "native"){
        try{
            $script:ie.Quit() | Out-Null
        }catch{}
        $shellapp = New-Object -ComObject "Shell.Application"
        $ShellWindows = $shellapp.Windows()
        for ($i = 0; $i -lt $ShellWindows.Count; $i++)
        {
          if ($ShellWindows.Item($i).FullName -like "*iexplore.exe")
          {
            $del = $ShellWindows.Item($i)
            try{
                $Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($del)  2>&1 
            }catch{}
          }
        }
        try{
            $Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shellapp) 
        }catch{}
        if($autoProtectedMode){ 
            revertProtectedMode 
        } 
    }else{
        if($debugMode -and $authMethod -eq "native"){
            try{
                $debugFilePath = Join-path (split-path $logfile -Parent) -ChildPath "OnedriveMapper.debug"
                $debugInfo | Out-File -FilePath $debugFilePath -Force -Confirm:$False -ErrorAction Stop -Encoding UTF8
            }catch{
                log -text "Error writing debug file: $($Error[0])" -fout
            }
        }
    }
    handleAzureADConnectSSO
    log -text "OnedriveMapper has finished running"
    if($restartExplorer){
        restart_explorer
    }else{
        log -text "restartExplorer is set to False, if you're redirecting My Documents, it won't show until next logon" -warning
    }
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
            $password = CustomInputBox "Microsoft Office 365 OneDrive" "Please enter your password for Office 365" -password
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
            $login = CustomInputBox "Microsoft Office 365 OneDrive" "Please enter your login name for Office 365"
        }catch{ 
            log -text "failed to display a login input box, exiting $($Error[0])" -fout
            abort_OM              
        } 
    } 
    until($login.Length -gt 0 -or $askAttempts -gt 2) 
    if($askAttempts -gt 3) { 
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
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    $failed = $False
    try{
        $folder = (get-itemproperty $regPath -Name Cookies -ErrorAction Stop).Cookies
    }catch{
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        $failed = $True
    }
    if($failed){
        try{
            $folder = (get-itemproperty $regPath -Name Cookies).Cookies
            $failed = $False
        }catch{
            log -text "Unable to find out where your cookies are stored! This means we can't double-check if they were really created when we attempt to create them and won't be able to delete existing cookies. $($Error[0])"
        }
    }
    if(!$failed){
        $cookies = get-childitem $folder -include * -recurse -force | where {$_.Extension -eq ".cookie" -or $_.Extension -eq ".txt"}
        foreach($cookie in $cookies){
            if(Get-Content -Path $cookie.FullName | Select-String -Pattern "sharepoint.com"){
                Remove-Item $cookie -Force -ErrorAction SilentlyContinue
            }
        }
    }

    [DateTime]$dateTime = Get-Date
    $dateTime = $dateTime.AddDays(5)
    $str = $dateTime.ToString("R")
    $relevantCookies += $script:cookiejar.GetCookies("https://$O365CustomerName-my.sharepoint.com")
    $relevantCookies += $script:cookiejar.GetCookies("https://$O365CustomerName.sharepoint.com")
    $findMe = @()
    foreach($cookie in $relevantCookies){
        [String]$cookieValue = [String]$cookie.Value.Trim()
        [String]$cookieDomain = [String]$cookie.Domain.Trim()
        try{
            if($cookie.Name -eq "rtFa"){
                $findMe+=$cookieDomain
                $cookieDomain = "https://$($cookieDomain)"
                log -text "Setting rtFA cookie for $cookieDomain...."
                $res = [Cookies.setter]::SetWinINETCookieString($cookieDomain,"rtFa","$cookieValue;Expires=$str")
            }
            if($cookie.Name -eq "FedAuth"){
                $findMe+=$cookieDomain
                $cookieDomain = "https://$($cookieDomain)"
                log -text "Setting FedAuth cookie for $cookieDomain...."
                $res = [Cookies.setter]::SetWinINETCookieString($cookieDomain,"FedAuth","$cookieValue;Expires=$($str)")
            }
        }catch{
            log -text "Failed to set a cookie: $($Error[0])" -fout
        }
    }

    #check if it was properly created
    if(!$failed){
        $cookies = get-childitem $folder -include * -recurse -force | where {$_.Extension -eq ".cookie" -or $_.Extension -eq ".txt"}
        $found = 0
        foreach($cookie in $cookies){
            foreach($find in $findMe){
                if(Get-Content -Path $cookie.FullName | Select-String -Pattern $find){
                    $found++
                }
            }
        }
        if($found -lt $findMe.Count){
            Throw "Found $found cookies out of the expected $($findMe.Count) cookies in $folder"
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
    if($askAttempts -gt 3) { 
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
            $res = JosL-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/x-www-form-urlencoded" -accept "text/html, application/xhtml+xml, image/jxr, */*"       
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

function addMapping(){
    Param(
        [String]$driveLetter,
        [String]$url,
        [String]$label
    )
    $mapping = "" | Select-Object driveLetter, URL, Label, alreadyMapped
    $mapping.driveLetter = $driveLetter
    $mapping.url = $url
    $mapping.label = $label
    $mapping.alreadyMapped = $False
    log -text "Adding to mapping list: $driveLetter ($url)"
    return $mapping
}

#this function checks if a given drivemapper is properly mapped to the given location, returns true if it is, otherwise false
function checkIfLetterIsMapped(){
    Param(
        [String]$driveLetter,
        [String]$url
    )
    if([System.IO.Directory]::Exists($driveLetter)){ 
        #check if mapped path is to at least the personal folder on Onedrive for Business, username detection would require a full login and slow things down
        #Ignore DavWWWRoot, as this does not consistently appear in the actual URL
        try{
            [string]$mapped_URL = @(Get-WMIObject -query "Select * from Win32_NetworkConnection Where LocalName = '$driveLetter'")[0].RemoteName.Replace("DavWWWRoot\","").Replace("@SSL","")
        }catch{
            log -text "problem detecting network path for $driveLetter, $($Error[0])" -fout
        }
        [String]$url = $url.Replace("DavWWWRoot\","").Replace("@SSL","")
        if($mapped_URL.StartsWith($url)){
            log -text "the mapped url for $driveLetter ($mapped_URL) matches the expected URL of $url, no need to remap"
            return $True
        }else{
            log -text "the mapped url for $driveLetter ($mapped_URL) does not match the expected partial URL of $url"
            return $False
        } 
    }else{
        log -text "$driveLetter is not yet mapped"
        return $False
    }
}

function waitForIE{
    do {sleep -m 100} until (-not ($script:ie.Busy))
}

function checkIfMFASetupIsRequired{
    try{
        $found_Tfa = (getElementById -id "tfa_setupnow_button").tagName
        #two factor was required but not yet set up
        log -text "Failed to log in at $($script:ie.LocationURL) because you have not set up two factor authentication methods while you are required to." -fout
        $script:errorsForUser += "Cannot continue: you have not yet set up two-factor authentication at portal.office.com"
        abort_OM 
    }catch{$Null}    
}

function checkIfCOMObjectIsHealthy{
    #check if the COM object is healthy, otherwise we're running into issues 
    if($script:ie.HWND -eq $null){ 
        log -text "ERROR: the browser object was Nulled during login, this means IE ProtectedMode or other security settings are blocking the script, check if you have correctly configure Trusted Sites." -fout
        $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
        abort_OM 
    } 
}

#Returns True if there was an error (and logs the error), returns False if no error was detected
function checkErrorAtLoginValue{
    Param(
        [String]$mode #msonline = microsoft, #adfs = adfs of client
    )
    if($mode -eq "msonline"){
        try{
            $found_ErrorControl = (getElementById -id "error_code").value
        }catch{$Null}
    }elseif($mode -eq "adfs"){
        try{
            $found_ErrorControl = (getElementById -id "errorText").innerHTML
        }catch{$Null}
    }
    if($found_ErrorControl){
        if($mode -eq "msonline"){
            switch($found_ErrorControl){
                "InvalidUserNameOrPassword" {
                    log -text "Detected an error at $($ie.LocationURL): invalidUsernameOrPassword" -fout
                    $script:errorsForUser += "The password or login you're trying to use is invalid`n"
                }
                default{                
                    log -text "Detected an error at $($ie.LocationURL): $found_ErrorControl" -fout
                    $script:errorsForUser += "Office365 reported an error: $found_ErrorControl`n"
                }
            }
            return $True
        }elseif($mode -eq "adfs"){
            if($found_ErrorControl.Length -gt 1){
                log -text "Detected an error at $($ie.LocationURL): $found_ErrorControl" -fout
                return $True
            }else{
                return $False
            }
        }
    }else{
        return $False
    }
}

function checkIfMFAControlsArePresent{
    Param(
        [Switch]$withoutADFS
    )
    try{
        $found_TfaWaiting = (getElementById -id "tfa_results_container").tagName
    }catch{$found_TfaWaiting = $Null}
    if($found_TfaWaiting){return $True}else{return $False}
}

function loginV2(){
    Param(
        $tryAgainRes
    )
    $script:cookiejar = New-Object System.Net.CookieContainer
###TODO
#re-enter usename wanneer nodig
#detecteer niet met forms maar op url waar mogelijk, verschillende ADFS/Okta versies
    log -text "Login attempt using native method at tenant $O365CustomerName"
    $uidEnc = [System.Web.HttpUtility]::UrlEncode($userUPN)
    #stel allereerste cookie in om websessie te beginnen
    try{
        if($adfsSmartLink){
            $res = JosL-WebRequest -url $adfsSmartLink -Method Get
            $mode = "Federated"
        }else{
            if(!$tryAgainRes) {
                $res = JosL-WebRequest -url https://login.microsoftonline.com -Method Get
            }else{
                $res = $tryAgainRes
            }
            $apiCanary = returnEnclosedFormValue -res $res -searchString "`"apiCanary`":`""
            $iwaEndpoint = returnEnclosedFormValue -res $res -searchString "iwaEndpointUrlFormat: `""
            $clientId = returnEnclosedFormValue -res $res -searchString "correlationId`":`""
            #vind session code en gebruik deze om het realm van de user te vinden
            $stsRequest = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"ctx`" value=`""
            $flowToken = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"flowToken`" value=`""
            $canary = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"canary`" value=`""
            #URL encode the canary for POST request
            $canary = [System.Web.HttpUtility]::UrlEncode($canary)
            $res = JosL-WebRequest -url "https://login.microsoftonline.com/common/userrealm?user=$uidEnc&api-version=2.1&stsRequest=$stsRequest&checkForMicrosoftAccount=false" -Method GET
        }
    }catch{
        log -text "Unable to find user realm due to $($Error[0])" -fout
        return $False
    }
    if(!$adfsSmartLink){
        $jsonRealmConfig = ConvertFrom-Json20 -item $res.Content
        $iwaEndpoint = $iwaEndpoint -replace  [regex]::escape("{0}"),$jsonRealmConfig.DomainName
        $mode = $Null
        if($jsonRealmConfig.NameSpaceType -eq "Managed"){
            $mode = "Managed"
            log -text "Received API response for authentication method: Managed"
            if($jsonRealmConfig.is_dsso_enabled){
                $azureADSSOEnabled = $True
                log -text "Additionally, Azure AD SSO and/or PassThrough is enabled for your tenant"
            }
            $nextURL = "https://login.microsoftonline.com/common/login"
        }
        if($jsonRealmConfig.NameSpaceType -eq "Federated"){
            $mode = "Federated"
            $nextURL = $jsonRealmConfig.AuthURL
            log -text "Received API response for authentication method: Federated"
            log -text "Authentication target: $nextURL"
        }
        $uidEnc = [System.Web.HttpUtility]::HtmlEncode($jsonRealmConfig.Login)
    }

    #authenticate using Managed Mode
    if($mode -eq "Managed"){
        $attempts = 0
        #if azure AD SSO is enable, we need to trigger a session with the backend
        if($azureADSSOEnabled){
            $nextURL2 = "https://autologon.microsoftazuread-sso.com/$($jsonRealmConfig.DomainName)/winauth/sso?desktopsso=true&isAdalRequest=False&client-request-id=$clientId"
            log -text "Authentication target: $nextURL2"
            try{
                $res = JosL-WebRequest -url $nextURL2 -trySSO 1 -method GET -accept "text/html, application/xhtml+xml, image/jxr, */*" -referer "https://login.microsoftonline.com/" 
                log -text "Azure AD SSO response received: $($res.Content)"
                $ssoToken = $res.Content
            }catch{
                log -text "no SSO token received from AzureAD, did you add autologon.microsoftazuread-sso.com to the local intranet sites?" -warning
            }
            $nextURL2 = "https://login.microsoftonline.com/common/instrumentation/dssostatus"
            $customHeaders = @{"canary" = $apiCanary;"hpgid" = "1002";"hpgact" = "2101";"client-request-id"=$clientId}
            $JSON = @{"resultCode"="107";"ssoDelay"="200";"log"=$Null}
            $JSON = ConvertTo-Json20 -item $JSON
            $res = JosL-WebRequest -url $nextURL2 -method POST -body $JSON -customHeaders $customHeaders
            $JSON = ConvertFrom-Json20 -item $res.Content
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
                $res = JosL-WebRequest -url $nextURL -Method POST -body $body
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
                    $JSON = ConvertTo-Json20 -item $JSON
                    $res = JosL-WebRequest -url $nextURL -Method POST -body $JSON -customHeaders $customHeaders
                    $response = ConvertFrom-Json20 -item $res.Content
                    $body = "flowToken=$($response.flowToken)&ctx=$stsRequest"
                    $nextURL =  "https://login.microsoftonline.com/common/onpremvalidation/End"
                    $res = JosL-WebRequest -url $nextURL -Method POST -body $body
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
                    $JSON = ConvertFrom-Json20 -item $res.Content
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
                $res = JosL-WebRequest -url $nextURL -Method GET
            }catch{
                log -text "Error received from ADFS server: $($Error[0])" -fout
                return $False
            }
        }
        ##if we get a SAML token, we've been signed in automatically, otherwise, we will have to post our credentials
        $wResult = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"wresult`" value=`""
        if($wResult -eq -1){
            log -text "Federation Services did not sign us in automatically, retrieving user credentials.." -warning
            $password = retrievePassword
            $passwordEnc = [System.Web.HttpUtility]::HtmlEncode($password)
            $ADFShost = $jsonRealmConfig.AuthURL.SubString(0,$jsonRealmConfig.AuthURL.IndexOf("adfs/ls"))
            $nextURL = returnEnclosedFormValue -res $res -searchString "<form method=`"post`" id=`"loginForm`" autocomplete=`"off`" novalidate=`"novalidate`" onKeyPress=`"if (event && event.keyCode == 13) Login.submitLoginRequest();`" action=`"/" -decode
            if($nextURL.IndexOf("https:") -eq -1){
                $nextURL = "$($ADFShost)$($nextURL)"
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
                $body = "UserName=$userName&Password=$passwordEnc&Kmsi=true&AuthMethod=FormsAuthentication"
                $res = JosL-WebRequest -url $nextURL -Method POST -body $body
                $attempts++
            }

        }
        $nextURL = returnEnclosedFormValue -res $res -searchString "<form method=`"POST`" name=`"hiddenform`" action=`""
        $wctx = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"wctx`" value=`""
        if($wctx -ne -1){
            $wctx = "$($wctx)&amp;LoginOptions=1"
            $wctx = [System.Web.HttpUtility]::htmldecode($wctx)
            $wctx = [System.Web.HttpUtility]::UrlEncode($wctx)
            $wctx = "&wctx=$($wctx)"
        }else{
            $wctx = $Null
        }
        $wResult = [System.Web.HttpUtility]::HtmlDecode($wResult)
        $wResult = [System.Web.HttpUtility]::UrlEncode($wResult)
        $body = "wa=wsignin1.0&wresult=$wResult"
        $res = JosL-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/x-www-form-urlencoded" -accept "text/html, application/xhtml+xml, image/jxr, */*"       
    }

    #some customers have a redirect active to Onedrive for Business, check if we're already there and return true if so
    if($res.rawResponse.ResponseUri.OriginalString.IndexOf("/personal/") -ne -1){
        log -text "Logged into Office 365! Already redirected to Onedrive for Business"
        return $True
    }

    #we should be back at an O365 page now, signed in but still having to do a POST
    $nextURL = returnEnclosedFormValue -res $res -searchString "form name=`"fmHF`" id=`"fmHF`" action=`"" -decode
    $nextURL = [System.Web.HttpUtility]::HtmlDecode($nextURL)
    $value = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"t`" id=`"t`" value=`""
    if($value -eq -1){
        log -text "We do not seem to have been properly redirected after signing in." -fout
        return $False
    }
    $body = "t=$value"
    try{
        $res = JosL-WebRequest -url $nextURL -Method POST -body $body -referer $res.rawResponse.ResponseUri.AbsoluteUri -contentType "application/x-www-form-urlencoded" -accept "text/html, application/xhtml+xml, image/jxr, */*"       
    }catch{
        $Null
    }
    $value = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"t`" id=`"t`" value=`""
    if($value -eq -1){
        log -text "You don't seem to be logged in properly" -warning
        return $False
    }else{
        log -text "Logged into Office 365!"
        return $True
    }
}

#region loginFunction
function login(){ 
    log -text "Login attempt using IE method"
    #AzureAD SSO check if a tile exists for this user
    $skipNormalLogin = $False
    if($userLookupMode -eq 3){
        try{
            $lookupQuery = $userUPN -replace "@","_"
            $lookupQuery = $lookupQuery -replace "\.","_"
            $userTile = getElementById -id $lookupQuery
            $skipNormalLogin = $True
            log -text "detected SSO option for OnedriveMapper through AzureAD, attempting to login automatically"
            $userTile.Click()
            waitForIE
            Sleep -m 500
            waitForIE
            if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
                #we've been logged in, we can abort the login function 
                log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
                return $True             
            }else{
                $skipNormalLogin = $False
                log -text "failed to use Azure AD SSO for Workplace Joined devices because we are not yet signed in while we should be" -fout
            }
        }catch{
            $skipNormalLogin = $False
            log -text "failed to use Azure AD SSO for Workplace Joined devices" -fout
        }
    }
    
    if(!$skipNormalLogin){
        #click to open up the login menu 
        try{
            (getElementById -id "use_another_account").Click()
            log -text "Found sign in elements type 1 on Office 365 login page, proceeding" 
        }catch{
            log -text "Failed to find signin element type 1 on Office 365 login page, trying next method. Error details: $($Error[0])"
        }
        try{
            (getElementById -id "use_another_account_link").click() 
            log -text "Found sign in elements type 2 on Office 365 login page, proceeding" 
        }catch{
            log -text "Failed to find signin element type 2 on Office 365 login page, trying next method. Error details: $($Error[0])"
        }
        try{
            $Null = getElementById -id "cred_keep_me_signed_in_checkbox"
            log -text "Found sign in elements type 3 on Office 365 login page, proceeding" 
        }catch{
            log -text "Failed to find signin element type 3 at $($script:ie.LocationURL). You may have to upgrade to a later Powershell version or Install Office. Attempting to log in anyway, this will likely fail. Error details: $($Error[0])" -fout
        }
        waitForIE 
        if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
            #we've been logged in, we can abort the login function 
            log -text "user detected as logged in, login function succeeded but mapping will probably fail, final url: $($script:ie.LocationURL)" 
            return $True             
        }
 
        $userName = $userUPN

        log -text "Will use $userName as login"

        #attempt to trigger redirect to detect if we're using ADFS automatically 
        try{ 
            log -text "attempting to trigger a redirect to SSO Provider using method 1" 
            $checkBox = getElementById -id "cred_keep_me_signed_in_checkbox"
            if($checkBox.checked -eq $False){
                $checkBox.click() 
                log -text "Signin Option persistence selected"
            }else{
                log -text "Signin Option persistence was already selected"
            }
            if($checkBox.checked -eq $False){
                log -text "the cred_keep_me_signed_in_checkbox is not selected! This may result in error 224" -fout
            }
            (getElementById -id "cred_userid_inputtext").value = $userName       
            waitForIE 
            (getElementById -id "cred_password_inputtext").click() 
            waitForIE
        }catch{ 
            log -text "Failed to find the correct controls at $($script:ie.LocationURL) to log in by script, check your browser and proxy settings or check for an update of this script. $($Error[0])" -fout
            $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
            return $False
        } 
    }
    sleep -s 2 

    #update progress bar
    if($showProgressBar) {
        $script:progressbar1.Value = 35
        $script:form1.Refresh()
    }

    $redirWaited = 0 
    while($True){ 
        sleep -m 500 

        checkIfMFASetupIsRequired

        checkIfCOMObjectIsHealthy

        #If ADFS or Azure automatically signs us on, this will trigger
        if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
            log -text "Detected an url that indicates we've been signed in automatically: $($script:ie.LocationURL)"
            $useADFS = $True
            break            
        }

        #this is the ADFS login control ID, modify this in the script setup if you have a custom IdP
        try{
            $found_ADFSControl = getElementById -id $adfsLoginInput
        }catch{
            $found_ADFSControl = $Null
            log -text "Waited $redirWaited of $adfsWaitTime seconds for SSO redirect. While looking for $adfsLoginInput at $($script:ie.LocationURL). If you're not using SSO this message is expected."
        }

        $redirWaited += 0.5 
        #found ADFS control
        if($found_ADFSControl){
            log -text "ADFS Control found, we were redirected to: $($script:ie.LocationURL)" 
            $useADFS = $True
            break            
        } 

        if($redirWaited -ge $adfsWaitTime){ 
            log -text "waited for more than $adfsWaitTime to get redirected by SSO provider, attempting normal signin" 
            $useADFS = $False    
            break 
        } 
    }  
    
    #update progress bar
    if($showProgressBar) {
        $script:progressbar1.Value = 40
        $script:form1.Refresh()
    }       

    #if not using ADFS, sign in 
    if($useADFS -eq $False){ 
        if($abortIfNoAdfs){
            log -text "abortIfNoAdfs was set to true, SSO provider was not detected, script is exiting" -fout
            $script:errorsForUser += "Onedrivemapper cannot login because SSO provider is not available`n"
            return $False
        }
        if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
            #we've been logged in, we can abort the login function 
            log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
            return $True             
        }
        $pwdAttempts = 0
        while($pwdAttempts -lt 3){
            $pwdAttempts++
            try{ 
                $checkBox = getElementById -id "cred_keep_me_signed_in_checkbox"
                if($checkBox.checked -eq $False){
                    $checkBox.click() 
                    log -text "Signin Option persistence selected"
                }else{
                    log -text "Signin Option persistence was already selected"
                }
                if($pwdAttempts -gt 1){
                    if($userLookupMode -eq 4){
                        $userName = (retrieveLogin -forceNewUsername)
                        (getElementById -id "cred_userid_inputtext").value = $userName
                    }
                    (getElementById -id "cred_password_inputtext").value = retrievePassword -forceNewPassword
                }else{
                    (getElementById -id "cred_password_inputtext").value = retrievePassword 
                }
                (getElementById -id "cred_sign_in_button").click() 
                waitForIE
            }catch{ 
                if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
                    #we've been logged in, we can abort the login function 
                    log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
                    return $True 
                } 
                log -text "Failed to find the correct controls at $($ie.LocationURL) to log in by script, check your browser and proxy settings or check for an update of this script (2). $($Error[0])" -fout
                return $False
            }
            Sleep -s 1
            waitForIE
            #check if the error field does not appear, if it does not our attempt was succesfull
            if((checkErrorAtLoginValue -mode "msonline") -eq $False){
                break   
            }else{
                log -text "There was an issue while trying to log in during attempt $pwdAttempts" -fout
            }
            $script:errorsForUser = $Null
        }

        #Office 365 two factor is required (SMS NOT YET SUPPORTED)
        if((checkIfMFAControlsArePresent -withoutADFS)){ 
            $waited = 0
            $maxWait = 90
            $loop = $True
            $MfaCodeAsked = $False
            while($loop){
                Sleep -s 2
                $waited+=2
                #check if on the MFA page, otherwise we're past the page already
                if((checkIfMFAControlsArePresent -withoutADFS)){ 
                    log -text "Waited for $waited seconds for user to complete Multi-Factor Authentication $found_TfaWaiting"
                }else{
                    log -text "Multi-Factor Authentication completed in $waited seconds"
                    $loop = $False
                }
                #check for SMS/App input field container
                try{
                    $found_MfaCode = getElementById -id "tfa_code_container"
                }catch{
                    $found_MfaCode = $Null
                }
                #if field is visible and we haven't asked before, ask for the text/app message code, otherwise user is probably using the phonecall method
                if($found_MfaCode -ne $Null -and $found_MfaCode.ariaHidden -ne $True -and $MfaCodeAsked -eq $False){
                    $MfaCodeAsked = $True
                    $code = askForCode
                    (getElementById -id "tfa_code_inputtext").value = $code
                    waitForIE
                    (getElementById -id "tfa_signin_button").click() 
                    waitForIE
                }
                if($waited -ge $maxWait){
                    log -text "Failed to log in at $($script:ie.LocationURL) because multi-factor authentication was not completed in time." -fout
                    $script:errorsForUser += "Cannot continue: you have not completed multi-factor authentication in the maximum alotted time"
                    return $False 
                }
            }

        }
    }else{ 
        #check if logged in now automatically after ADFS redirect 
        if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
            #we've been logged in, we can abort the login function 
            log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
            return $True 
        } 
    } 
 
    waitForIE
 
    #Check if we arrived at a 404, or an actual page 
    if($script:ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") { 
        log -text "We received a 404 error after our signin attempt, retrying...." -fout
        $script:ie.navigate("https://portal.office.com")
        waitForIE
        if($script:ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") { 
            log -text "We received a 404 error again, aborting" -fout
            $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
            return $False    
        }     
    } 

    #check if logged in now 
    if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
        #we've been logged in, we can abort the login function 
        log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
        return $True 
    }else{ 
        if($useADFS){ 
            log -text "ADFS did not automatically sign us on, attempting to enter credentials at $($script:ie.LocationURL)" 
            $pwdAttempts = 0
            while($pwdAttempts -lt 3){
                $pwdAttempts++
                try{ 
                    if($userLookupMode -eq 4 -and $pwdAttempts -gt 1){
                        $userName = (retrieveLogin -forceNewLogin)
                    }
                    if($adfsMode -eq 1){
                        (getElementById -id $adfsLoginInput).value = $userName
                    }else{
                        (getElementById -id $adfsLoginInput).value = ($userName.Split("@")[0])
                    }
                    if($pwdAttempts -gt 1){
                        (getElementById -id $adfsPwdInput).value = retrievePassword -forceNewPassword
                    }else{
                        (getElementById -id $adfsPwdInput).value = retrievePassword 
                    }
                    (getElementById -id $adfsButton).click() 
                    waitForIE 
                    sleep -s 1 
                    waitForIE  
                    if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
                        #we've been logged in, we can abort the login function 
                        log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
                        return $True 
                    } 
                }catch{ 
                    log -text "Failed to find the correct controls at $($script:ie.LocationURL) to log in by script, check your browser and proxy settings or modify this script to match your ADFS form. Error details: $($Error[0])" -fout
                    $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
                    return $False 
                }
                #check if the error field does not appear, if it does not our attempt was succesfull
                if((checkErrorAtLoginValue -mode "adfs") -eq $False){
                    break   
                }
            }
 
            waitForIE   
            #check if logged in now         
            if((checkIfAtO365URL -url $ie.LocationURL -finalURLs $finalURLs)){
                #we've been logged in, we can abort the login function 
                log -text "login detected, login function succeeded, final url: $($ie.LocationURL)" 
                return $True 
            }else{ 
                log -text "We attempted to login with ADFS, but did not end up at the expected location. Detected url: $($ie.LocationURL), expected URL: $($baseURL)" -fout
                $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
                return $False
            } 
        }else{ 
            log -text "We attempted to login without using ADFS, but did not end up at the expected location. Detected url: $($ie.LocationURL), expected URL: $($baseURL)" -fout
            $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
            return $False 
        } 
    } 
} 
#endregion

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
                $failed = $False
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
                    $basePath = "HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\AADNGC"
                    if((test-path $basePath) -eq $False){
                        log -text "userLookupMode is set to 3, but the registry path $basePath does not exist! Using method 2" -fout
                        $basePath = "HKCU:\Software\Classes\Local Settings\Software\Microsoft\SettingSyncHost.exe\WinMSIPC"
                        if((test-path $basePath) -eq $False){
                            log -text "userLookupMode is set to 3, but the registry path $basePath does not exist! Using method 3" -fout 
                            $objUser = New-Object System.Security.Principal.NTAccount($Env:USERNAME)
                            $strSID = ($objUser.Translate([System.Security.Principal.SecurityIdentifier])).Value
                            $basePath = "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$strSID\IdentityCache\$strSID"
                            if((test-path $basePath) -eq $False){
                                log -text "userLookupMode is set to 3, but the registry path $basePath does not exist! All lookup modes exhausted, exiting" -fout
                                abort_OM   
                            }
                            $userId = (Get-ItemProperty -Path $basePath -Name UserName).UserName
                        }else{
                            $userId = @(Get-ChildItem $basePath)[0].Name | Split-Path -Leaf
                        }
                    }else{
                        $basePath = @(Get-ChildItem $basePath)[0].Name -Replace "HKEY_CURRENT_USER","HKCU:"
                        $userId = (Get-ItemProperty -Path $basePath -Name UserId).UserId
                    }
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
                        $res = queryForAllCreds -introText $loginformIntroText -buttonText $buttonText -loginLabel $loginFieldText -passwordLabel $passwordFieldText
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
            $res = Start-Service WebClient -ErrorAction SilentlyContinue
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

function autoUpdate{
    param(
        [Parameter(Mandatory=$true)]$desiredVersion,
        [Parameter(Mandatory=$true)]$newVersionPath
    )
    if($desiredVersion -gt $version){
        #a newer version of the script is set to deploy
        $pathToSelf = $script:MyInvocation.MyCommand.Path
        log -text "New version detected: $desiredVersion, auto updating from $version at $pathToSelf"
        log -text "Attempting to download from $newVersionPath"
        try{
            $req = JosL-WebRequest -url $newVersionPath -method GET
            log -text "New version retrieved succesfully, replacing..."
        }catch{
            log -text "Failed to download new version: $($Error[0])" -fout
            Throw
        }
        
        $req.Content = $req.Content -replace ".+#Don't modify this, unless you are using OnedriveMapper Cloud edition", "`$configurationID       = `"$($configurationID)`"#Don't modify this, unless you are using OnedriveMapper Cloud edition"
        try{
            $req.Content | Out-File -FilePath $pathToSelf -Force -Confirm:$False -ErrorAction Stop
            log -text "Replaced old version, update succesfull!"
        }catch{
            log -text "failed to overwrite old version of the script, reason: $($Error[0])" -fout
            Throw
        }
        #start new version
        restartMe
    }
}

function restartMe{
    Param(
        [Switch]$fallBackMode
    )
    if($fallBackMode){
        $arguments = "$($arguments) -fallbackMode"
        log -text "restarting OnedriveMapper in different authentication mode"
    }else{
        log -text "starting newer version of OnedriveMapper"
    }
    try{$form1.Close()}catch{$Null}
    try{          
        Start-Process powershell -ArgumentList $arguments -Wait
        Exit
    }catch{
        log -text "Failed to restart Onedrivemapper: $($Error[0])" -fout
        Exit        
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
    $res = [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    log -text "Set TLS protocol version to 1.2"
}catch{
    log -text "Failed to set TLS protocol to version 1.2 $($Error[0])" -fout
}

#get IE version on this machine
$ieVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Internet Explorer').svcVersion
if($ieVersion -eq $Null){
    $ieVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Internet Explorer').Version
    $ieVersion = $ieVersion.Split(".")[1]
}else{
    $ieVersion = $ieVersion.Split(".")[0]
}

#get OSVersion
$windowsVersion = ([System.Environment]::OSVersion.Version).Major
$objUser = New-Object System.Security.Principal.NTAccount($Env:USERNAME)
$strSID = ($objUser.Translate([System.Security.Principal.SecurityIdentifier])).Value
log -text "You are $strSID running on Windows $windowsVersion with IE $ieVersion and Powershell version $($PSVersionTable.PSVersion.Major)"

#load settings from OnedriveMapper Cloud Configurator if licensed
if($configurationID -ne "00000000-0000-0000-0000-000000000000"){
    $loadFromFile = $False
    try{
        log -text "configurationID set to $configurationID, retrieving associated settings from lieben.nu..."
        $rawSettingsResponse = JosL-WebRequest -url "http://om.lieben.nu/lieben_api.php?cid=$configurationID&ieVersion=$ieVersion&winVersion=$windowsVersion&oVersion=$version&uid=$strSID"
        $configuratorSettings = ConvertFrom-Json20 $rawSettingsResponse.Content
        log -text "settings retrieved, processing..."
    }catch{
        $loadFromFile = $True
        log -text "failed to retrieve settings from lieben.nu using $configurationID because of $($Error[0]), content of request: $($rawSettingsReponse.Content)" -fout
    }
    if($settingsCache -and !$loadFromFile){
        try{
            log -text "Caching settings to file..."
            storeSettingsToCache -settingsCache $settingsCache -settings $configuratorSettings
            log -text "Cached settings to file"
        }catch{
            log -text "Failed to cache settings to file: $($Error[0])" -fout
        }
    }
    if($loadFromFile){
        log -text "attempting to load cached settings from a previous request"
        try{
            if($settingsCache){
                $configuratorSettings = retrieveSettingsFromCache -settingsCache $settingsCache
            }else{
                Throw "Settings cache was disabled in script configuration"
            }
        }catch{
            log -text "failed to retrieve settings from cache $($Error[0])" -fout
            $script:errorsForUser="could not retrieve settings, check your connection"
            abort_OM
        }
    }
    $domain = $configuratorSettings.upnSuffix
    $O365CustomerName = $configuratorSettings.O365CustomerName
    $O365CustomerName = $O365CustomerName.Split(".")[0]
    if($configuratorSettings.deleteUnmanagedDrives -eq "No") {$deleteUnmanagedDrives = $False}else{$deleteUnmanagedDrives = $True}
    if($configuratorSettings.redirectMyDocuments -eq "Yes") {$redirectMyDocs = $True}else{$redirectMyDocs = $False}
    if($configuratorSettings.redirectFavorites -eq "Yes") {$redirectFavorites = $True}else{$redirectFavorites = $False}
    if($configuratorSettings.redirectDesktop -eq "Yes") {$redirectDesktop = $True}else{$redirectDesktop = $False}
    $redirectToSubfolderName = $configuratorSettings.redirectMyDocumentsName
    if($configuratorSettings.authMethod -eq "IE") {$authMethod = "IE"}else{$authMethod = "native"}
    if($configuratorSettings.allowFallbackMode -eq "No") {$allowFallbackMode = $False}else{$allowFallbackMode = $True}    
    if($configuratorSettings.debugMode -eq "Yes") {$debugmode = $True}else{$debugmode = $False}
    $userLookupMode = $configuratorSettings.userLookupMode
    if($configuratorSettings.AzureAADConnectSSO -eq "No") {$AzureAADConnectSSO = $False}else{$AzureAADConnectSSO = $True}
    if($configuratorSettings.lookupUserGroups -eq "Yes") {
        try{
            $groups = ([ADSISEARCHER]"samaccountname=$($env:USERNAME)").Findone().Properties.memberof -replace '^CN=([^,]+).+$','$1'
            log -text "cached user group membership because lookupUserGroups was set to True"   
            $lookupUserGroups = $True 
        }catch{
            log -text "Failed to cache user group membership because of $($Error[0])" -fout
        }
    }else{$lookupUserGroups = $False}
    $forceUserName = $configuratorSettings.forceUserName
    $forcePassword = $configuratorSettings.forcePassword
    if($configuratorSettings.restartExplorer -eq "Yes") {$restartExplorer = $True}else{$restartExplorer = $False}
    if($configuratorSettings.addShellLink -eq "Yes") {$addShellLink = $True}else{$addShellLink = $False}
    if($configuratorSettings.persistentMapping -eq "No") {$persistentMapping = $False}else{$persistentMapping = $True}
    if($configuratorSettings.autoProtectedMode -eq "No") {$autoProtectedMode = $False}else{$autoProtectedMode = $True}
    $adfsWaitTime = $configuratorSettings.adfsWaitTime
    $libraryName = $configuratorSettings.libraryName
    if($configuratorSettings.autoKillIE -eq "No") {$autoKillIE = $False}else{$autoKillIE = $True}
    if($configuratorSettings.abortIfNoAdfs -eq "Yes") {$abortIfNoAdfs = $True}else{$abortIfNoAdfs = $False}
    $adfsMode = $configuratorSettings.adfsMode
    if($configuratorSettings.adfsSmartLink.Length -gt 5){
        $adfsSmartLink = $configuratorSettings.adfsSmartLink
    }
    if($configuratorSettings.displayErrors -eq "No") {$displayErrors = $False}else{$displayErrors = $True}
    $buttonText = $configuratorSettings.buttonText
    $loginformIntroText = $configuratorSettings.loginformIntroText
    $loginFieldText = $configuratorSettings.loginFieldText
    $passwordFieldText = $configuratorSettings.passwordFieldText
    $adfsLoginInput = $configuratorSettings.adfsLoginInput
    $adfsPwdInput = $configuratorSettings.adfsPwdInput
    $adfsButton = $configuratorSettings.adfsButton
    $urlOpenAfter = $configuratorSettings.urlOpenAfter
    if($configuratorSettings.showConsoleOutput -eq "No") {$showConsoleOutput = $False}else{$showConsoleOutput = $True}
    if($configuratorSettings.showElevatedConsole -eq "No") {$showElevatedConsole = $False}else{$showElevatedConsole = $True}
    if($configuratorSettings.showProgressBar -eq "No") {$showProgressBar = $False}else{$showProgressBar = $True}
    if($configuratorSettings.autoDetectProxy -eq "Yes") {$autoDetectProxy = $True}else{$autoDetectProxy = $False}
    log -text "Settings determined, starting..."
    if($configuratorSettings.allowAutoUpdate -eq "Yes"){
        try{
            autoUpdate -newVersionPath $configuratorSettings.newVersionPath -desiredVersion $configuratorSettings.desiredVersion
        }catch{
            log -text "failed to update to new OnedriveMapper version" -fout
        }
    }
}

if($showConsoleOutput -eq $False){
    $t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    try{
        add-type -name win -member $t -namespace native
        [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
    }catch{$Null}
}

if($allowFallbackMode -and $fallBackMode){
    if($authMethod -eq "native"){
        $authMethod = "ie"
        log -text "detected we are running in fallback mode, switching from native to ie" -warning
    }else{
        $authMethod = "native"
        log -text "detected we are running in fallback mode, switching from ie to native" -warning
    }
}

if($authMethod -eq "native" -and $PSVersionTable.PSVersion.Major -le 2){
    log -text "ERROR: you're trying to use Native auth on Powershell V2 or lower, switching to IE mode" -fout
    $authMethod = "IE"
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
    $screen = ([System.Windows.Forms.Screen]::AllScreens | where {$_.Primary}).WorkingArea
    $form1.Location = New-Object System.Drawing.Size(($screen.Right - $width), ($screen.Bottom - $height))
    $form1.Topmost = $True 
    $form1.TopLevel = $True 

    # create label
    $label1 = New-Object system.Windows.Forms.Label
    $label1.text="OnedriveMapper v$version is connecting your drives..."
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
    $label2.backColor= "#CC99FF"

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
    $progressbar1.Value = 5
    $form1.Refresh()
}

#do a version check if allowed
if($versionCheck){
    #update progress bar
    try{
        versionCheck -currentVersion $version
        log -text "NOTICE: you are running the latest (v$version) version of OnedriveMapper"
    }catch{
        if($showProgressBar) {
            $form1.controls["Label1"].Text = "New OnedriveMapper version available!"
            $form1.Refresh()
            Sleep -s 1
            $form1.controls["Label1"].Text = "OnedriveMapper v$version is connecting your drives..."
            $form1.Refresh()
        }
        log -text "ERROR: $($Error[0])" -fout
    }
}

#load cookie code and test-set a cookie
if($authMethod -eq "native"){
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
}

#check if KB2846960 is installed or not
if($authMethod -ne "native"){
    try{
        $res = get-hotfix -id kb2846960 -erroraction stop
        log -text "KB2846960 detected as installed"
    }catch{
        if($ieVersion -eq 10 -and $windowsVersion -lt 10){
            log -text "KB2846960 is not installed on this machine, if you're running IE 10 on anything below Windows 10, you may not be able to map your drives until you install KB2846960" -warning
        }
    }
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
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -eq 2 -or (checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "*") -eq 2){
        log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level"  
        $privateZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -eq 2 -or (checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "*") -eq 2){
        log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level (through GPO)"  
        $privateZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -eq 2 -or (checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "*") -eq 2){
        log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on user level"  
        $publicZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -eq 2 -or (checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "*") -eq 2){
        log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on user level (through GPO)" 
        $publicZoneFound = $True        
    }
}

#check if sharepoint and onedrive are set as safe sites at the machine level
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -eq 2 -or (checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "*") -eq 2){
    log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level"
    $privateZoneFound = $True 
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -eq 2 -or (checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "*") -eq 2){
    log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level (through GPO)"  
    $privateZoneFound = $True        
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -eq 2 -or (checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "*") -eq 2){
    log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on machine level"  
    $publicZoneFound = $True    
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -eq 2 -or (checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "*") -eq 2){
    log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on machine level (through GPO)"  
    $publicZoneFound = $True    
}

#log results, try to automatically add trusted sites to user trusted sites if not yet added
if($publicZoneFound -eq $False){
    log -text "Possible critical error: $($O365CustomerName).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail" -fout
    if((addSiteToIEZoneThroughRegistry -siteUrl "$O365CustomerName.sharepoint.com") -eq $True){log -text "Automatically added $O365CustomerName.sharepoint.com to trusted sites for this user"}
}
if($privateZoneFound -eq $False){
    log -text "Possible critical error: $($O365CustomerName)$($privateSuffix).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail" -fout
    if((addSiteToIEZoneThroughRegistry -siteUrl "$($O365CustomerName)$($privateSuffix).sharepoint.com") -eq $True){log -text "Automatically added $($O365CustomerName)$($privateSuffix).sharepoint.com to trusted sites for this user"}
}

#Check if IE FirstRun is disabled at computer level
if($authMethod -ne "native" -and (checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main" -entryName "DisableFirstRunCustomize") -ne 1){
    log -text "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\DisableFirstRunCustomize not found or not set to 1 registry, if script hangs this may be due to the First Run popup in IE, checking at user level..." -warning
    #Check if IE FirstRun is disabled at user level
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Internet Explorer\Main" -entryName "DisableFirstRunCustomize") -ne 1){
        log -text "HKCU:\Software\Microsoft\Internet Explorer\Main\DisableFirstRunCustomize not found or not set to 1 registry, if script hangs this may be due to the First Run popup in IE, attempting to autocorrect..." -warning
        try{
            $res = New-ItemProperty "hkcu:\software\microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -value 1 -ErrorAction Stop
            log -text "automatically prevented IE firstrun Popup"
        }catch{
            log -text "failed to automatically add a registry key to prevent the IE firstrun wizard from starting"
        }
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

$userUPN = getUserLogin

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 10
    $form1.Refresh()
}

#region flightChecks

#Check if required HTML parsing libraries have been installed 
if([System.IO.File]::Exists("$(${env:ProgramFiles(x86)})\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll") -eq $False){ 
    log -text "Microsoft Office installation not detected"
}  
 
#Check if webdav locking is enabled
if((checkRegistryKeyValue -basePath "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\" -entryName "SupportLocking") -ne 0){
    log -text "ERROR: WebDav File Locking support is enabled, this could cause files to become locked in your OneDrive or Sharepoint site" -fout 
} 

#report/warn file size limit
$sizeLimit = [Math]::Round((checkRegistryKeyValue -basePath "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\" -entryName "FileSizeLimitInBytes")/1024/1024)
log -text "Maximum file upload size is set to $sizeLimit MB" -warning

#check if any zones are configured with Protected Mode through group policy (which we can't modify) 
if($authMethod -ne "native"){
    $BaseKeypath = "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" 
    for($i=0; $i -lt 5; $i++){ 
        $curr = Get-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500" -ErrorAction SilentlyContinue | select -exp 2500 
        if($curr -ne $Null -and $curr -ne 3){ 
            log -text "IE Zone $i protectedmode is enabled through group policy, autoprotectedmode cannot disable it. This will likely cause the script to fail." -fout
        }
    }
}

#endregion

#translate to URLs 
$mapURLpersonal = ("\\"+$O365CustomerName+"$($privateSuffix).sharepoint.com@SSL\DavWWWRoot\personal\") 

$desiredMappings = @() #array with mappings to be made

#add the O4B mapping first, with an incorrect URL that will be updated later on because we haven't logged in yet and can't be sure of the URL
if($configurationID -ne "00000000-0000-0000-0000-000000000000"){
    [Array]$o4bMappings = @($configuratorSettings.mappings | where{$_.Type -eq "Onedrive" -and $_})
    if($o4bMappings.Count -eq 1){
        $dontMapO4B = $False
        $desiredMappings += addMapping -driveLetter $o4bMappings[0].Driveletter -url $mapURLpersonal -label $o4bMappings[0].Label
        $driveLetter = $o4bMappings[0].Driveletter
        $driveLabel = $o4bMappings[0].Label
    }else{
        $dontMapO4B = $True
        log -text "0 or more than 1 mappings to Onedrive returned by the web service, will not map Onedrive for Business"
    }
}else{
    if($dontMapO4B){
        log -text "Not mapping O4B because dontMapO4B is set to True"
    }else{
        $desiredMappings += addMapping -driveLetter $driveLetter -url $mapURLpersonal -label $driveLabel
    }
}

if($dontMapO4B){
    $baseURL = ("https://"+$O365CustomerName+".sharepoint.com") 
}else{
    $baseURL = ("https://$($O365CustomerName)-my.sharepoint.com/_layouts/15/MySite.aspx?MySiteRedirect=AllDocuments") 
}

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 15
    $form1.Refresh()
}

if($configurationID -ne "00000000-0000-0000-0000-000000000000"){
    $sharepointMappings = @()
    [Array]$returnedSharepointMappings = @($configuratorSettings.mappings | where{$_.Type -eq "Sharepoint" -and $_.Url -and $_.Driveletter -and $_.Label})
    if($returnedSharepointMappings.Count -eq 0){
        log -text "No sharepoint mappings detected from web service for configurationID $configurationID"
    }else{
        foreach($spOMapping in $returnedSharepointMappings){
            if($lookupUserGroups -and $groups){
                if($spOMapping.SecurityGroup.Length -gt 2 -and $spOMapping.SecurityGroup -ne "N/A" -and $spOMapping.SecurityGroup -ne "If you wish to restrict this mapping to a specific security group, type the EXACT group name here"){
                    $group = $groups -contains $spOMapping.SecurityGroup  
                    if($group){
                        $sharepointMappings+="$($spOMapping.Url),$($spOMapping.Label),$($spOMapping.Driveletter)"
                    }                    
                }
            }else{
                $sharepointMappings+="$($spOMapping.Url),$($spOMapping.Label),$($spOMapping.Driveletter)"
            }
        }
        log -text "Loaded $($returnedSharepointMappings.Count) Sharepoint mappings from web service with id $configurationID"
    }
}

#add any desired Sharepoint Mappings
$sharepointMappings | % {
    $data = $_.Split(",")
    if($data[0] -and $data[1] -and $data[2]){
        if($WebAssemblyloaded){
            $add = [System.Web.HttpUtility]::UrlDecode($data[0])
        }else{
            $add = $data[0]
        }
        $add = $add.Replace("https://","\\") 
        $add = $add.Replace("/_layouts/15/start.aspx#","")
        $add = $add.Replace("sharepoint.com/","sharepoint.com@SSL\DavWWWRoot\") 
        $add = $add.Replace("/","\") 
        $desiredMappings += addMapping -driveLetter $data[2] -url $add -label $data[1]
    }
}

$continue = $False
$countMapping = 0
#check if any of the mappings we should make is already mapped and update the corresponding property
$desiredMappings | % {
    if((checkIfLetterIsMapped -driveLetter $_.driveletter -url $_.url)){
        $desiredMappings[$countMapping].alreadyMapped = $True
        if(($redirectMyDocs -or $redirectDesktop -or $redirectFavorites) -and $_.driveletter -eq $driveLetter) {
            $res = redirectMyDocuments -driveLetter $driveLetter
        }
    }
    $countMapping++
}
 
if(@($desiredMappings | where-object{$_.alreadyMapped -eq $False}).Count -le 0){
    log -text "no unmapped or incorrectly mapped drives detected"
    abort_OM    
}

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 20
    $form1.Refresh()
}

handleAzureADConnectSSO -initial

log -text "Base URL: $($baseURL) `n" 

#Start IE and stop it once to make sure IE sets default registry keys 
if($authMethod -ne "native" -and $autoKillIE){ 
    #start invisible IE instance 
    $tempIE = new-object -com InternetExplorer.Application 
    $tempIE.visible = $debugmode 
    sleep 2 
 
    #kill all running IE instances of this user 
    $ieStatus = Get-ProcessWithOwner iexplore 
    if($ieStatus -eq 0){ 
        log -text "no instances of Internet Explorer running yet, at least one should be running" -warning
    }elseif($ieStatus -eq -1){ 
        log -text "Checking status of iexplore.exe: unable to query WMI" -fout
    }else{ 
        log -text "autoKillIE enabled, stopping IE processes" 
        foreach($Process in $ieStatus){ 
                Stop-Process $Process.handle -ErrorAction SilentlyContinue
                log -text "Stopped process with handle $($Process.handle)"
        } 
    } 
}elseif($authMethod -eq "ie"){ 
    log -text "ERROR: autoKillIE disabled, IE processes not stopped. This may cause the script to fail for users with a clean/new profile" -fout
} 

if($authMethod -ne "native" -and $autoProtectedMode){ 
    log -text "autoProtectedMode is set to True, disabling ProtectedMode temporarily" 
    $BaseKeypath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" 
     
    #store old values and change new ones 
    try{ 
        for($i=0; $i -lt 5; $i++){ 
            $curr = Get-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500" -ErrorAction SilentlyContinue| select -exp 2500 
            if($curr -ne $Null){ 
                $protectedModeValues[$i] = $curr 
                log -text "Zone $i was set to $curr, setting it to 3" 
            }else{
                $protectedModeValues[$i] = 0 
                log -text "Zone $i was not yet set, setting it to 3" 
            }
            Set-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500"  -Value "3" -Type Dword -ErrorAction Stop
        } 
    } 
    catch{ 
        log -text "Failed to modify registry keys to autodisable ProtectedMode $($error[0])" -fout
    } 
}elseif($authMethod -ne "native"){
    log -text "autoProtectedMode is set to False, IE ProtectedMode will not be disabled temporarily" -fout
}

#start invisible IE instance 
if($authMethod -ne "native"){
    $COMFailed = $False
    try{ 
        $script:ie = new-object -com InternetExplorer.Application -ErrorAction Stop
        $script:ie.visible = $debugmode 
    }catch{ 
        log -text "failed to start Internet Explorer COM Object, check user permissions or already running instances. Will retry in 30 seconds. $($Error[0])" -fout
        $COMFailed = $True
    } 

    #retry above if failed
    if($COMFailed){
        Sleep -s 30
        try{ 
            $script:ie = new-object -com InternetExplorer.Application -ErrorAction Stop
            $script:ie.visible = $debugmode 
        }catch{ 
            log -text "failed to start Internet Explorer COM Object a second time, check user permissions or already running instances $($Error[0])" -fout
            $errorsForUser += "Mapping cannot continue because we could not start your browser"
            abort_OM 
        }
    }
}

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 25
    $form1.Refresh()
}

if($authMethod -ne "native"){
    #navigate to the base URL of the tenant's Sharepoint to check if it exists 
    try{ 
        $script:ie.navigate("https://login.microsoftonline.com/logout.srf")
        waitForIE
        sleep -s 1
        waitForIE
        $script:ie.navigate($o365loginURL) 
        waitForIE
        sleep -s 1
    }catch{ 
        log -text "Failed to browse to the Office 365 Sign in page, this is a fatal error $($Error[0])`n" -fout
        $errorsForUser += "Mapping cannot continue because we could not contact Office 365`n"
        abort_OM 
    } 
 
    #check if we got a 404 not found 
    if($script:ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") { 
        log -text "Failed to browse to the Office 365 Sign in page, 404 error detected, exiting script" -fout
        $errorsForUser += "Mapping cannot continue because we could not start the browser`n"
        abort_OM 
    } 
 
    checkIfCOMObjectIsHealthy

    if($script:ie.LocationURL.StartsWith($o365loginURL)){
        log -text "Starting logon process at: $($script:ie.LocationURL)" 
    }elseif(!(checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
        log -text "For some reason we're not at the logon page, even though we tried to browse there, we'll probably fail now but let's try one final time. URL: $($script:ie.LocationURL)" -fout
        $script:ie.navigate($o365loginURL) 
    }
}

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 30
    $form1.Refresh()
}

if($authMethod -ne "native"){
    #log in 
    if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
        log -text "Detected an url that indicates we've been signed in automatically: $($script:ie.LocationURL), but we did not select sign in persistence, this may cause an error when mapping" -fout
    }else{ 
        #Check and log if Explorer is running 
        $explorerStatus = Get-ProcessWithOwner explorer 
        if($explorerStatus -eq 0){ 
            log -text "no instances of Explorer running yet, expected at least one running" -warning
        }elseif($explorerStatus -eq -1){ 
            log -text "Checking status of explorer.exe: unable to query WMI" -fout
        }else{ 
            log -text "Detected running explorer process" 
        } 
        $res = login
        if($False -ne $res){
            log -text "IE login function succeeded"
        }else{
            if($allowFallbackMode -and !$fallbackMode){
                log -text "fallback mode is enabled, and login failed. Attempting native auth mode..." -fout
                restartMe -fallBackMode
            }else{
                log -text "IE auth login mode failed, aborting script" -fout
                abort_OM
            }
        }
        $script:ie.navigate($baseURL) 
        waitForIE
        do {sleep -m 100} until ($script:ie.ReadyState -eq 4 -or $script:ie.ReadyState -eq 0)  
        Sleep -s 2
    } 
}else{
    $res = loginV2
    if($res -eq $False){
        if($allowFallbackMode -and !$fallbackMode){
            log -text "fallback mode is enabled, and login failed. Attempting IE auth mode..." -fout
            restartMe -fallBackMode
        }else{
            log -text "native auth login mode failed, aborting script" -fout
            abort_OM
        }
    }
}

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 45
    $script:form1.Refresh()
}

#delete mappings to Sharepoint / Onedrive that aren't managed by OnedriveMapper
if($deleteUnmanagedDrives){
    [Array]$currentMappings = @(Get-WMIObject -query "Select * from Win32_NetworkConnection" | where {$_})
    foreach($currentMapping in $currentMappings){
        if($desiredMappings.driveletter -contains $currentMapping.LocalName){continue}
        $searchStringSpO = "\\$($O365CustomerName).sharepoint.com"
        $searchStringO4B = "\\$($O365CustomerName)-my.sharepoint.com"
        if($currentMapping.RemoteName.StartsWith($searchStringSpO) -or $currentMapping.RemoteName.StartsWith($searchStringO4B)){
            try{$del = NET USE $currentMapping.LocalName /DELETE /Y 2>&1}catch{$Null}
        }
    }
}

#username detection method
if($dontMapO4B -eq $False -and !$desiredMappings[0].alreadyMapped){
    if($authMethod -ne "native"){
        #find username
        $url = $script:ie.LocationURL 
        $timeSpent = 0
        while($url.IndexOf("/personal/") -eq -1){
            log -text "Attempting to detect username at $url, waited for $timeSpent seconds" 
            $script:ie.navigate($baseURL)
            waitForIE
            if($timeSpent -gt 60){
                log -text "Failed to get the username from the URL for over $timeSpent seconds while at $url, aborting" -fout 
                $errorsForUser += "Mapping cannot continue because we cannot detect your username`n"
                abort_OM 
            }
            Sleep -s 2
            $timeSpent+=2
            $url = $script:ie.LocationURL
        }
        try{
            $start = $url.IndexOf("/personal/")+10 
            $end = $url.IndexOf("/",$start) 
            $userURL = $url.Substring($start,$end-$start) 
            $mapURL = $mapURLpersonal + $userURL + "\" + $libraryName 
        }catch{
            log -text "Failed to get the username while at $url, aborting" -fout
            $errorsForUser += "Mapping cannot continue because we cannot detect your username`n"
            abort_OM 
        }
        $desiredMappings[0].url = $mapURL 
        log -text "Detected user: $($userURL)"
        log -text "Onedrive cookie generated, mapping drive..."
        $mapresult = MapDrive $desiredMappings[0].driveLetter $desiredMappings[0].url $desiredMappings[0].label
    }else{
        log -text "Retrieving Onedrive for Business cookie step 1..." 
        #trigger forced authentication to SpO O4B and follow the redirect
        try{
            $res = JosL-WebRequest -url $baseURL -method GET
            $nextURL = [System.Web.HttpUtility]::HtmlDecode((returnEnclosedFormValue -res $res -searchString "form name=`"fmHF`" id=`"fmHF`" action=`"" -decode))
            $value = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"t`" id=`"t`" value=`""
            $body = "t=$value"
        }catch{
            log -text "Failed to retrieve cookie for Onedrive for Business: $($Error[0])" -fout
        }
        try{
            if((returnEnclosedFormValue -res $res -searchString "<form method=`"POST`" name=`"hiddenform`" action=`"" -decode) -ne -1){
                $nextURL = returnEnclosedFormValue -res $res -searchString "<form method=`"POST`" name=`"hiddenform`" action=`"" -decode
                $code = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"code`" value=`""
                $id_token = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"id_token`" value=`""
                $session_state = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"session_state`" value=`""
                 $body = "code=$code&id_token=$id_token&session_state=$session_state"
            }
            if($nextURL.Length -gt 10){
                log -text "Retrieving Onedrive for Business cookie step 2 at $nextURL"
                $res = JosL-WebRequest -url $nextURL -Method POST -body $body
            }else{
                throw "no next url detected: $nextURL"
            }
        }catch{
            log -text "Problem reported during step 2: $($Error[0])" -fout
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
                    continue
                }else{
                    Throw "No username detected in response string"
                }  
            }catch{
                log -text "Waited for $timeWaited seconds for O4b auto provisioning..."
            }
            if($timeWaited -gt 0){
                $res = JosL-WebRequest -url "https://$($O365CustomerName)-my.sharepoint.com/_layouts/15/MyBraryFirstRun.aspx?FirstRunStage=waiting" -method GET
            }
            sleep -s 10
            $res = JosL-WebRequest -url $baseURL -method GET
            $timeWaited += 10
        }
        try{
            setCookies
        }catch{
            log -text "Failed to set cookies, error received: $($Error[0])" -fout
        }
        $desiredMappings[0].url = $mapURL
        log -text "Onedrive cookie loop finished, mapping drive..."
        $mapresult = MapDrive $desiredMappings[0].driveLetter $desiredMappings[0].url $desiredMappings[0].label
    }
    if($addShellLink -and $windowsVersion -eq 6 -and [System.IO.Directory]::Exists($desiredMappings[0].driveLetter)){
        try{
            $res = createFavoritesShortcutToO4B -targetLocation $desiredMappings[0].driveLetter
        }catch{
            log -text "Failed to create a shortcut to the mapped drive for Onedrive for Business because of: $($Error[0])" -fout
        }
    }
}

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 50
    $script:form1.Refresh()
    $maxAdded = 40
    $added = 0
}

foreach($spMapping in $sharepointMappings){
    $data = $spMapping.Split(",")
    $desiredMapping = $Null
    [Array]$desiredMapping = @($desiredMappings | where{$_.alreadyMapped -eq $False -and $_.driveLetter -eq $data[2] -and $_})
    if($desiredMapping.Count -ne 1){
        continue
    }
    log -text "Initiating session with: $($data[0])"
    if($data[0] -and $data[1] -and $data[2]){
        $spURL = $data[0] #URL to browse to
        #original IE method to set cookies
        if($authMethod -ne "native"){
            log -text "Current location: $($script:ie.LocationURL)" 
            $script:ie.navigate($spURL) #check the URL
            $waited = 0
            waitForIE
            while($($ie.LocationURL) -notlike "$spURL*"){
                sleep -s 1
                $waited++
                log -text "waited $waited seconds to load $spURL, currently at $($ie.LocationURL)"
                if($waited -ge $maxWaitSecondsForSpO){
                    log -text "waited longer than $maxWaitSecondsForSpO seconds to load $spURL! This mapping may fail" -fout
                    break
                }
            }
            if($script:ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*" -or $script:ie.HWND -eq $null) { 
                log -text "Failed to browse to Sharepoint URL $spURL.`n" -fout
            } 
            log -text "Current location: $($script:ie.LocationURL)" 
        }else{#new method to set cookies
            log -text "Retrieving Sharepoint cookie step 1..." 
            #trigger forced authentication to SpO and follow the redirect if needed
            try{
                $res = JosL-WebRequest -url $data[0] -method GET
                $nextURL = [System.Web.HttpUtility]::HtmlDecode((returnEnclosedFormValue -res $res -searchString "form name=`"fmHF`" id=`"fmHF`" action=`"" -decode))               
                $value = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"t`" id=`"t`" value=`""
                $body = "t=$value"
            }catch{
                log -text "Failed to retrieve cookie for SpO: $($Error[0])" -fout
            }
            try{
                if((returnEnclosedFormValue -res $res -searchString "<form method=`"POST`" name=`"hiddenform`" action=`"" -decode) -ne -1){
                    $nextURL = returnEnclosedFormValue -res $res -searchString "<form method=`"POST`" name=`"hiddenform`" action=`"" -decode
                    $code = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"code`" value=`""
                    $id_token = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"id_token`" value=`""
                    $session_state = returnEnclosedFormValue -res $res -searchString "<input type=`"hidden`" name=`"session_state`" value=`""
                    $body = "code=$code&id_token=$id_token&session_state=$session_state"                    
                }
                if($nextURL.Length -gt 10){
                    log -text "Retrieving Sharepoint cookie step 2 at $nextURL"
                    $res = JosL-WebRequest -url $nextURL -Method POST -body $body
                }else{
                    throw "no next url detected: $nextURL"
                }
            }catch{
                log -text "Problem reported during step 2: $($Error[0])" -fout
            }
            try{
                setCookies
            }catch{
                log -text "Failed to set cookies, error received: $($Error[0])" -fout
            }
        }
    }
    #update progress bar
    if($showProgressBar) {
        if($added -le $maxAdded){
            $script:progressbar1.Value += 10
            $script:form1.Refresh()
        }
        $added+=10
    }
    log -text "SpO cookie generated, attempting to map drive"
    $mapresult = MapDrive $desiredMapping[0].driveLetter $desiredMapping[0].url $desiredMapping[0].label
}

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 100
    $script:form1.Refresh()
}

abort_OM