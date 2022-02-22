$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Test the path C:\Migration is there, create it if it deosnt exist
$Migrationfolder = Test-Path C:\Migration
If($Migrationfolder -eq $False) {
    Try {
        $ErrorActionPreference = 'stop'
        New-Item -Path "C:\" -Name "Migration" -ItemType "directory"
    }
    Catch {
        # If there is an error, log the error
  	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Host $error[0].Exception.GetType().FullName
        Write-Host $errormessage -ForegroundColor Red
    }
}

# Begin Teams sign out
" " | out-file -append "C:\migration\_log.txt"
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING: Teams sign out"
Write-Host $output
$output | out-file -append "C:\migration\_log.txt"

# Get MS Teams process. Only using 'SilentlyContinue' as we test this below
$TeamsProcess = Get-Process Teams -ErrorAction SilentlyContinue

# Get Outlook process. Only using 'SilentlyContinue' as we test this below
$OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue

If ($TeamsProcess) {
    # If 'Teams' process is running, stop it else do nothing
    $TeamsProcess | Stop-Process -Force
    Start-Sleep 3
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Teams process was running, so we stopped it"
    Write-Host $output
	$output | out-file -append "C:\migration\_log.txt"
}

If ($OutlookProcess) {
    # If 'Outlook' process is running, stop it else do nothing
    $OutlookProcess | Stop-Process -Force | Wait-Process
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Outlook process was running, so we stopped it"
    Write-Host $output
	$output | out-file -append "C:\migration\_log.txt"
}

# Clear Teams cached folders under %appdata%\Microsoft\Teams
# Check if teams folder cache was already cleared by script
$TeamsCache = Test-Path -Path C:\Migration\teams-cache-cleared.txt
If ($TeamsCache -eq $false){
    $TeamsFolderCheck = Test-Path -Path $env:APPDATA\"Microsoft\Teams"
    If ($TeamsFolderCheck -eq $true) {
        # Check if 'Teams' folder exists in %APPDATA%\Microsoft\Teams
        Try { 
            $ErrorActionPreference = 'stop'
            $Blob = Test-Path -Path $env:APPDATA\"Microsoft\teams\blob_storage"
            If( $Blob -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\blob_storage" | Remove-Item -Recurse -ErrorAction SilentlyContinue
            }
            $db = Test-Path -Path $env:APPDATA\"Microsoft\teams\databases"
            If( $db -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\databases" | Remove-Item -Recurse -ErrorAction SilentlyContinue
            }
            $cache = Test-Path -Path $env:APPDATA\"Microsoft\teams\cache"
            If( $cache -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\cache" | Remove-Item -Recurse -ErrorAction SilentlyContinue
            }
            $gpucache = Test-Path -Path $env:APPDATA\"Microsoft\teams\gpucache"
            If( $gpucache -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\gpucache" | Remove-Item -Recurse -ErrorAction SilentlyContinue
            }
            $Indexeddb = Test-Path -Path $env:APPDATA\"Microsoft\teams\Indexeddb"
            If( $Indexeddb -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Indexeddb" | Remove-Item -Recurse -ErrorAction SilentlyContinue
            }
            $local = Test-Path -Path $env:APPDATA\"Microsoft\teams\Local Storage"
            If( $local -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Local Storage" | Remove-Item -Recurse -ErrorAction SilentlyContinue
            }
            $tmp = Test-Path -Path $env:APPDATA\"Microsoft\teams\tmp"
            If( $tmp -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\tmp" | Remove-Item -Recurse -ErrorAction SilentlyContinue
            }
            $cookies = Test-Path -Path $env:APPDATA\"Microsoft\teams\Cookies"
            If( $cookies -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Cookies" | Remove-Item -ErrorAction SilentlyContinue
            }
            $storage = Test-Path -Path $env:APPDATA\"Microsoft\teams\storage.json"
            If( $storage -eq $True) {
                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\storage.json" | Remove-Item -ErrorAction SilentlyContinue
            }
            # If no errors renaming folder, log success and write file to c:\Migration\teams-folder-rename.txt so script doesnt repeat action
	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $output = $timestamp + " Teams cache was cleared and user was signed out"
            Write-Host $output
	        $output | out-file -append "C:\migration\_log.txt"
            New-Item C:\Migration\teams-cache-cleared.txt
        }
        Catch {
            # If there is an error, log the error
  	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $errormessage=$timestamp + " ERROR: " + $_.ToString()
	        Write-Host $error[0].Exception.GetType().FullName
            Write-Host $errormessage -ForegroundColor Red
	        $errormessage | out-file -append "C:\migration\_log.txt"
        }
    }
    Else {
        # If 'Teams' folder does not exist notify user then break
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Teams folder not found"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
    }
}
Else {
    # If 'Teams' folder was already renamed then skip
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Teams folder was already renamed "
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
}

# Check if HomeUserUpn for teams was cleared by script
$TeamsUserClear = Test-Path -Path C:\Migration\teams-homeuserupn-cleared.txt
If ($TeamsUserClear -eq $false){
    Try { 
        $ErrorActionPreference = 'stop'
        Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Teams" -Name "HomeUserUpn"
        # If no errors clearing HomeUserUpn, log success and write file to c:\Migration\teams-username-cleared.txt so script doesnt repeat action
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Registry entry deleted: HKCU:\Software\Microsoft\Office\Teams\HomeUserUpn"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\teams-homeuserupn-cleared.txt
    }
    Catch {
        # If there is an error, log the error
  	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Host $error[0].Exception.GetType().FullName
        Write-Host $errormessage -ForegroundColor Red
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}
# If HomeUserUpn was cleared before then skip
Else {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Teams HomeUserUpn was already cleared "
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
}
# Modify desktop-Config.Json under %appdata%\Microsoft\Teams
# Check if modify desktop-Config.json was already done by script
$TeamsModifyDesktopConfig = Test-Path -Path C:\Migration\teams-modify-desktop-config.txt
If($TeamsModifyDesktopConfig -eq $false) {
    Try { 
        # Import desktop-Config.json
        $ErrorActionPreference = 'stop'
        $TeamsFolder = "$env:APPDATA\Microsoft\Teams"
        $SourceDesktopConfigFile = "$TeamsFolder\desktop-config.json"
	    $desktopConfig = (Get-Content -Path $SourceDesktopConfigFile | ConvertFrom-Json)
        # If no errors importing desktop-config, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " " + $TeamsFolder + "\desktop-Config/json imported successfuly"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        # Modify desktop-Config.json
        $desktopConfig.isLoggedOut = $true
	    $desktopConfig.upnWindowUserUpn =""; #The email used to sign in
	    $desktopConfig.userUpn ="";
	    $desktopConfig.userOid ="";
	    $desktopConfig.userTid = "";
	    $desktopConfig.homeTenantId ="";
	    $desktopConfig.webAccountId="";
	    $desktopConfig | ConvertTo-Json -Compress | Set-Content -Path $SourceDesktopConfigFile -Force
        # If no errors modifying desktop-config, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " " + $TeamsFolder + "\desktop-Config/json modified successfuly"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\teams-modify-desktop-config.txt
    }
    Catch {
        # If there is an error, log the error
  	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Host $error[0].Exception.GetType().FullName
        Write-Host $errormessage -ForegroundColor Red
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}
# If script desktop-Config.json was already modified then skip
Else {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: desktop-Config.json was already modified"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
}

# Begin Outlook new profile creation
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING: New profile creation for Outlook "
Write-Host $output
$output | out-file -append "C:\migration\_log.txt"
# Check if new profile was created
$NewOutlookProfile = Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook-New
If ($NewOutlookProfile -eq $True) {
    # If profile was created, skip profile creation
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Outlook profile 'Outlook-New' was already created"
    Write-Host $output
	$output | out-file -append "C:\migration\_log.txt"
}
Else {
    # If profile was was not created, create new profile
    # If 'Outlook' process is running, stop it else do nothing
    # Get Outlook process. Only using 'SilentlyContinue' as we test this below
    $OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue
    If ($OutlookProcess) {
        $OutlookProcess | Stop-Process -Force  | Wait-Process
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " Outlook process was running, so we stopped it"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
    }
    Try {
        # Create new Outlook profile named 'Outlook-New'
        $ErrorActionPreference = 'stop'
        New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook-New
        # If no errors creating the new profile, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " RegKey created: HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook-New"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}
# Check if default profile has been set to Outlook-New
$DefaultProfile = Get-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook -Name DefaultProfile -ErrorAction SilentlyContinue
If($DefaultProfile.defaultprofile -eq "Outlook-New") {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " SKIPPING: Default Outlook profile already set to Outlook-New"
    Write-Host $output
	$output | out-file -append "C:\migration\_log.txt"
}
Else {
    # If 'Outlook' process is running, stop it else do nothing
    # Get Outlook process. Only using 'SilentlyContinue' as we test this below
    $OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue
    If ($OutlookProcess) {
        $OutlookProcess | Stop-Process -Force | Wait-Process
        Write-Host "Outlook process was running, so we stopped it" -ForegroundColor Green
    }
    Try {
        # Set default Outlook profile to 'Outlook-New'
        $ErrorActionPreference = 'stop'
        New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook -Name DefaultProfile -PropertyType String -Value "Outlook-New" -Force
        # If no errors setting the default Outlook profile, log success
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Reg string created: 'HKCU:\Software\Microsoft\Office\16.0\Outlook\DefaultProfile' with value 'Outlook-New'"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
    }
    # If there was an error, log the error
    Catch {
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

# Begin OneDrive sign out and unlink
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING: OneDrive sign out and unlink"
Write-Host $output
$output | out-file -append "C:\migration\_log.txt"
# Clear registry path HKCU:\Software\Microsoft\OneDrive\Accounts\* to unlink OneDrive account 
$OneDriveUnlinked1 = Test-Path -Path C:\Migration\onedrive-unlinked-1.txt
If ($OneDriveUnlinked1 -eq $false){
    # If OneDrive is open, close it
    $OneDriveProcess = Get-Process onedrive -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process
    # Delete regkeys to unlink OneDrive account
    Try {
        $ErrorActionPreference = 'stop'
        Remove-Item -Path HKCU:\Software\Microsoft\OneDrive\Accounts\* -Recurse
        # If no errors deleting the registry keys, log success
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Registry cleared: HKCU:\Software\Microsoft\OneDrive\Accounts\*"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item -Path C:\Migration\onedrive-unlinked-1.txt
    }
    # If there was an error, log the error
    Catch {
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}
Else {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	$output = $timestamp + " SKIPPING: Registry already cleared: HKCU:\Software\Microsoft\OneDrive\Accounts\*"
    Write-Host $output
	$output | out-file -append "C:\migration\_log.txt"
}
# Clear registry path HKCU:\software\microsoft\windows\currentversion\explorer\desktop\namespace\* to unlink OneDrive account 
$OneDriveUnlinked2 = Test-Path -Path C:\Migration\onedrive-unlinked-2.txt
If ($OneDriveUnlinked2 -eq $false){
    # If OneDrive is open, close it
    $OneDriveProcess = Get-Process onedrive -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process
    # Delete regkeys to unlink OneDrive account
    Try {
        $ErrorActionPreference = 'stop'
        Remove-Item -Path HKCU:\software\microsoft\windows\currentversion\explorer\desktop\namespace\* -Recurse
        # If no errors deleting the registry keys, log success
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Registry cleared: HKCU:\software\microsoft\windows\currentversion\explorer\desktop\namespace\*"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item -Path C:\Migration\onedrive-unlinked-2.txt
    }
    # If there was an error, log the error
    Catch {
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}
Else {
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	$output = $timestamp + " SKIPPING: Registry already cleared: HKCU:\software\microsoft\windows\currentversion\explorer\desktop\namespace\*"
    Write-Host $output
	$output | out-file -append "C:\migration\_log.txt"
}

# Clear OneDrive credentials unless script has been run previously
If ($OneDriveUnlinked1 -eq $false -and $OneDriveUnlinked2 -eq $false){
    Install-Module -Name pscredentialmanager -Scope CurrentUser
    Install-Module -Name CredentialManager -Scope CurrentUser
    $onedrive = Get-CachedCredential | where {$_.name -like "*onedrive*"}
    If($onedrive -ne $null) {
        remove-storedcredential -target $onedrive.name
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Cleared Cached OneDrive Credentials"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
    }
    Else {
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " No Cached OneDrive Credentials to clear"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
    }
}

# Begin Office Activation logout
#We need to make sure all the Office programs are closed, otherwise the IDentities Keys will be recreated and the user not logged out
$TeamProcess = Get-Process -ProcessName Teams -ErrorAction SilentlyContinue
If ($TeamsProcess) {
    # If 'Teams' process is running, stop it else do nothing
    $TeamsProcess | Stop-Process -Force
    Start-Sleep 3
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Teams process was running, so we stopped it"
    Write-Host $output
	$output | out-file -append "C:\migration\_log.txt"
}Start-Sleep 3
Get-Process -ProcessName EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process
Get-Process -ProcessName WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process
Get-Process -ProcessName POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process
Get-Process -ProcessName ONENOTE -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process
Get-Process -ProcessName MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process
Get-Process -ProcessName MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process
Get-Process -ProcessName Outlook -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process
Get-Process -ProcessName OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force | Wait-Process

$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING: Clearing Office Activation sign in"
Write-Host $output
$output | out-file -append "C:\migration\_log.txt"

$OfficeCommon = Get-Item Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common -ErrorAction SilentlyContinue

#Get the Any Identities of users.
#Iterate through each of the Identities under Common\Identity\Identities
#From each Identity, grab the ProviderID value, Put each into  the Identity Keys Value : SignedOutWAMUsers, seperate using semicolons
$IdentityKey = Get-Item Registry::$OfficeCommon\Identity -ErrorAction SilentlyContinue
$UserIdentityKeys = Get-ChildItem Registry::$IdentityKey\Identities -ErrorAction SilentlyContinue

$UserIdentityKeysCheck = Test-Path -Path Registry::$IdentityKey\Identities
if($UserIdentityKeysCheck -eq $False)
{
    $UserIdentityKeys = Get-ChildItem Registry::$IdentityKey\Identities -ErrorAction SilentlyContinue
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " No users logged into Office "
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
}
else
{
	$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " UserIdentityKeys: " + $UserIdentityKeys + "\Identities"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	# Iterates each Key under the Identity, and pulls the ProviderID which is an ID for each user logged in.
	$UserIdentityKeys | foreach {
		$CurrentProviderID = (Get-ItemProperty Registry::$_ -Name ProviderID -ErrorAction SilentlyContinue).ProviderID
		$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " CurrentProviderID: " + $CurrentProviderID
        Write-Host $output
        $output | out-file -append "C:\migration\_log.txt"
		if([string]::IsNullOrWhiteSpace($SignedOutWAMUsers))
		{
			$SignedOutWAMUsers = $CurrentProviderID
		}
		else
		{
			$SignedOutWAMUsers = $SignedOutWAMUsers +";"+ $CurrentProviderID
		}
		#Optionally, we can set a DWORD =1 for the value SignedOut on Identity SubKey
		#But they get removed next time you launch an Office program (the whole Subkey is removed)
	}
	$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Generated SignedOutWAMUsers: " + $SignedOutWAMUsers
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"

	#Compare the WAM to whats in the IdentityKey
	$ExistingWAM = $IdentityKey.GetValue("SignedOutWAMUsers")
	if(!($ExistingWAM -eq $null))
	{

		if(!($ExistingWAM))
		{
			#With the above null check, being here means empty
			$NewWam = ($SignedOutWAMUsers)

		}
		else
		{
			$NewWam = (($ExistingWAM +';'+ $SignedOutWAMUsers) -split ';' | Select -Unique)-join ';'
		}
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " Prior WAM Existed: " + $ExistingWAM
        Write-Host $output
        $output | out-file -append "C:\migration\_log.txt"
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " New WAM: " + $NewWam
        Write-Host $output
        $output | out-file -append "C:\migration\_log.txt"
		Set-ItemProperty Registry::$IdentityKey -Name SignedOutWAMUsers -Value $NewWam -ErrorAction SilentlyContinue
	}
	else
	{
		# there wasnt one here before, so we create a new registry value and put our SignedOutWAMUsers in it.
		New-ItemProperty Registry::$IdentityKey -Name SignedOutWAMUsers -Value $SignedOutWAMUsers -ErrorAction SilentlyContinue
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " No Prior WAM, Creating new Value From Generated SignOutWAMUsers"
        Write-Host $output
        $output | out-file -append "C:\migration\_log.txt"		
	}
}
    # Check if Identities Subkey was already deleted by script
    $IdentitiesKey = Test-Path -Path C:\Migration\office-identities-cleared.txt
    If($IdentitiesKey -eq $False) { 
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " Removing Identities Registry Key: " + $IdentityKey + "\Identities"
        Write-Host $output
        $output | out-file -append "C:\migration\_log.txt"
	    Try {
            $ErrorActionPreference = 'stop'
            Remove-Item Registry::$IdentityKey\Identities -Recurse
            # If no errors, log success
	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $output = $timestamp + " Removed: " + $IdentityKey + "\Identities"
            Write-Host $output
	        $output | out-file -append "C:\migration\_log.txt"
            New-Item C:\Migration\office-identities-cleared.txt
        }
        Catch {
            # If there is an error, log the error
            $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $errormessage=$timestamp + " ERROR: " + $_.ToString()
	        Write-Warning $errormessage
	        $errormessage | out-file -append "C:\migration\_log.txt"
        }
    }
    # Check if Profiles Subkey was already deleted by script
    $ProfileKey = Test-Path -Path C:\Migration\office-profiles-cleared.txt
    If($ProfileKey -eq $False) { 
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " Removing Profiles Registry Key: " + $IdentityKey + "\Profiles"
        Write-Host $output
        $output | out-file -append "C:\migration\_log.txt"
	    Try {
            $ErrorActionPreference = 'stop'
            Remove-Item Registry::$IdentityKey\Profiles -Recurse
            # If no errors, log success
	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $output = $timestamp + " Removed: " + $IdentityKey + "\Profiles"
            Write-Host $output
	        $output | out-file -append "C:\migration\_log.txt"
            New-Item C:\Migration\office-profiles-cleared.txt
        }
        Catch {
            # If there is an error, log the error
            $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $errormessage=$timestamp + " ERROR: " + $_.ToString()
	        Write-Warning $errormessage
	        $errormessage | out-file -append "C:\migration\_log.txt"
        }
    }
	
	#This one, I noticed starts to populate with URLS if you save to OneDrive
	#It also creates keys when Outlook the app is logged into and you select "Remember My Credentials".
    # Check if DocToIdMapping Subkey was already deleted by script
    $DocToIdMappingKey = Test-Path -Path C:\Migration\office-DocToIdMapping-cleared.txt
    If($DocToIdMappingKey -eq $False) { 
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
        $output = $timestamp + " Removing DocToIdMapping Registry Key: " + $IdentityKey + "\Profiles"
        Write-Host $output
        $output | out-file -append "C:\migration\_log.txt"
	    Try {
            $ErrorActionPreference = 'stop'
            Remove-Item Registry::$IdentityKey\DocToIdMapping\* -Recurse
             # If no errors, log success
	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $output = $timestamp + " Removed: " + $IdentityKey + "\DocToIdMapping\*"
            Write-Host $output
	        $output | out-file -append "C:\migration\_log.txt"
            New-Item C:\Migration\office-DocToIdMapping-cleared.txt
        }
        Catch {
            # If there is an error, log the error
            $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $errormessage=$timestamp + " ERROR: " + $_.ToString()
	        Write-Warning $errormessage
	        $errormessage | out-file -append "C:\migration\_log.txt"
        }
    }	

#Start Clearing out remanants that the user was logged in before#
$CloudPolicyKey = Test-Path -Path C:\Migration\office-cloudpolicy-cleared.txt
If($CloudPolicyKey -eq $False) { 
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing CloudPolicy Registry Key: " + $OfficeCommon + "\CloudPolicy"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	Try {
        $ErrorActionPreference = 'stop'
        $CloudPolicy = Get-Item Registry::$OfficeCommon\CloudPolicy
        Remove-Item Registry::$CloudPolicy\* -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $CloudPolicy
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-cloudpolicy-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    } 
}

$LicensingNextKey = Test-Path -Path C:\Migration\office-licensingnext-cleared.txt
If($LicensingNextKey -eq $False) { 
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing LicensingNext Registry Key: " + $OfficeCommon + "\Licensing\LicensingNext"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	Try {
        $ErrorActionPreference = 'stop'
        $LicensingNext = Get-Item Registry::$OfficeCommon\Licensing\LicensingNext
        Remove-Item Registry::$LicensingNext\* -Exclude "CIDtoLicenseIdsMapping" -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $LicensingNext
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-licensingnext-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    } 
}

$TemplatesKey = Test-Path -Path C:\Migration\office-templates-cleared.txt
If($TemplatesKey -eq $False) { 
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing Templates Registry Key: " + $OfficeCommon + "\OfficeStart\Web\Templates"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	Try {
        $ErrorActionPreference = 'stop'
        $WebTemplates = Get-Item Registry::$OfficeCommon\OfficeStart\Web\Templates
        Remove-Item Registry::$WebTemplates\* -Exclude "Anonymous" -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $WebTemplates
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-templates-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

$SettingsStoreKey = Test-Path -Path C:\Migration\office-settingsstore-cleared.txt
If($SettingsStoreKey -eq $False) { 
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing SettingsStore Registry Key: " + $OfficeCommon + "\Privacy\SettingsStore"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	Try {
        $ErrorActionPreference = 'stop'
        $SettingsStore =  Get-Item Registry::$OfficeCommon\Privacy\SettingsStore
        Remove-Item Registry::$SettingsStore\* -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $SettingsStore
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-settingsstore-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

$RoamIdKey = Test-Path -Path C:\Migration\office-roamid-cleared.txt
If($RoamIdKey -eq $False) { 
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing RoamId Registry Key: " + $OfficeCommon + "\Roaming\Identities"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	Try {
        $ErrorActionPreference = 'stop'
        $RoamId  = Get-Item Registry::$OfficeCommon\Roaming\Identities
        Remove-Item Registry::$RoamId\* -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $RoamId
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-roamid-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

$SerManCacheKey = Test-Path -Path C:\Migration\office-sermancache-cleared.txt
If($SerManCacheKey -eq $False) { 
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing SerManCache Registry Key: " + $OfficeCommon + "\ServicesManagerCache\Identities and \ServicesManagerCache\OnPremises"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	Try {
        $ErrorActionPreference = 'stop'
        $SerManCacheCheck = Test-Path -Path Registry::$OfficeCommon\ServicesManagerCache
        If($SerManCacheCheck -eq $true) {
            $SerManCache = Get-Item Registry::$OfficeCommon\ServicesManagerCache
            Remove-Item Registry::$SerManCache\Identities\* -Recurse
            Remove-Item Registry::$SerManCache\OnPremises\* -Recurse
            }
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $SerManCache + "\Identities and \OnPremises"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-sermancache-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

Write-Host "-- Clearing TargetMessaging --"
$TargetedMsgServKey = Test-Path -Path C:\Migration\office-targetedmsgserv-cleared.txt
If($TargetedMsgServKey -eq $False) { 
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing TargetedMsgServ Registry Key: " + $OfficeCommon + "\TargetedMessagingService\MessageData and \TargetedMessagingService\MessageMetaData"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	Try {
        $ErrorActionPreference = 'stop'
        $TargetedMsgServCheck = Test-Path -Path Registry::$OfficeCommon\TargetedMessagingService
        If($TargetedMsgServCheck -eq $true) {
            $TargetedMsgServ = Get-Item Registry::$OfficeCommon\TargetedMessagingService
            Remove-Item Registry::$TargetedMsgServ\MessageData\* -Recurse
            Remove-Item Registry::$TargetedMsgServ\MessageMetaData\* -Recurse
            }
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $TargetedMsgServ + "\MessageData and \MessageMetaData"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-targetedmsgserv-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

$UrlRepkey = Test-Path -Path C:\Migration\office-urlrep-cleared.txt
If($UrlRepkey -eq $False) { 
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing UrlRep Registry Key: " + $OfficeCommon + "\UrlReputation\UserPolicy"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	Try {
        $ErrorActionPreference = 'stop'
        $UrlRep = Get-Item Registry::$OfficeCommon\UrlReputation\UserPolicy
        Remove-Item Registry::$UrlRep\* -Recurse
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: " + $UrlRep
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-urlrep-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

#There are Registry Keys inside Office\16.0 (above Common) for each individual Office software.
#The following is used to clear user data each of these may have stored of a previous user.
Write-Host "-- Clearing App Specific Keys  --"
$AppsKey = Test-Path -Path C:\Migration\office-apps-cleared.txt
If($AppsKey -eq $False) { 
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Removing App Specific Registry Keys"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
	Try {
        $ErrorActionPreference = 'stop'
        $Apps = "Access;Excel;PowerPoint;Publisher;Word;OneNote"
        $Apps -split ";" | foreach{
	        Write-Host "-- -- Processing $_ -- --"
	        $CurAppKey = Get-Item Registry::"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\"$_ -ErrorAction SilentlyContinue
	        $FMRU = Get-Item Registry::$CurAppKey\'File MRU' -ErrorAction SilentlyContinue
	        $UMRU = Get-Item Registry::$CurAppKey\'User MRU' -ErrorAction SilentlyContinue
	        $PMRU = Get-Item Registry::$CurAppKey\'Place MRU' -ErrorAction SilentlyContinue
	        
	        If($FMRU -ne $NULL) {
                Remove-ItemProperty Registry::$FMRU -Name *
            }
            If($PMRU -ne$NULL) {
	            Remove-ItemProperty Registry::$PMRU -Name *
            }
	        If($UMRU -ne $NULL) {
                Remove-Item Registry::$UMRU\* -Recurse
            }

	        if($_ -eq "Word")
	        {
		        $RLCheck = Test-Path -Path Registry::$CurAppKey\"Reading Locations"
                If($RLCheck -eq $True) {
                    Remove-Item Registry::$CurAppKey\"Reading Locations"\* -Recurse
                }
	        }
	        if($_ -eq "OneNote")
	        {
		        $ONCheck = Test-Path -Path Registry::$CurAppKey\"OpenNotebooks"
                If($ONCheck -eq $True) {
                    Remove-Item Registry::$CurAppKey\"OpenNotebooks"\* -Recurse
                }
                $RNCheck = Test-Path -Path Registry::$CurAppKey\"RecentNotebooks"
                If($RNCheck -eq $True) {
		            Remove-Item Registry::$CurAppKey\"RecentNotebooks"\* -Recurse
                }
	        }
        }
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: App Specific Registry Keys"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-apps-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}	

#Clear  Files Associated with Cached  User  data inside LocalAppData
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING: Clearing LocalAppData"
Write-Host $output
$output | out-file -append "C:\migration\_log.txt"

#Most of these folders will be completely emptied, a few will have a single sub folder left behind.
$FoldersTest = Test-Path -Path C:\Migration\office-folders-cleared.txt
If($FoldersTest -eq $False) { 
    $OfficeDataDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0"
    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " Clearing Office Folders: " + $OfficeDataDir
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
    
	Try {
        $ErrorActionPreference = 'stop'
        $Folders = "aggmru;BackstageInAppNavCache;TapCache;MruServiceCache;Personalization\Content;Personalization\Governance"
        $Folders -split ";" | foreach {
            if($_ -like "Personalization\*") {
		        $SubFolderTest = Test-Path "$OfficeDataDir\$_\*"
                If($SubFolderTest -eq $True) {
                    Remove-Item "$OfficeDataDir\$_\*" -Recurse -Exclude "Anonymous"
                }
	        }
	        else {	
		        $SubFolderTest = Test-Path "$OfficeDataDir\$_\*"
                If($SubFolderTest -eq $True) {
                    Remove-Item "$OfficeDataDir\$_\*" -Recurse
                }
	        }
        }
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed: Office Folders " + $OfficeDataDir
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\office-folders-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

if($RmOneNoteFiles.IsPresent)
{
	$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
    $output = $timestamp + " ***STARTING: Deleting OneNote Local Folder"
    Write-Host $output
    $output | out-file -append "C:\migration\_log.txt"
    $OneNoteLocal = Test-Path -Path C:\Migration\onenote-local-cleared.txt
    If($OneNoteLocal -eq $False) { 
        Try {
            $ErrorActionPreference = 'stop'
            $ONRegKey = Get-Item Registry::"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\OneNote\Options\Save"
            $ONFolder = (Get-ItemProperty Registry::$ONRegKey).'Last Local Notebook Path'
            Remove-Item $ONFolder\* -Recurse -Force
            # If no errors, log success
	        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $output = $timestamp + " Removed: " + $ONFolder
            Write-Host $output
	        $output | out-file -append "C:\migration\_log.txt"
            New-Item C:\Migration\onenote-local-cleared.txt
        }
        Catch {
            # If there is an error, log the error
            $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	        $errormessage=$timestamp + " ERROR: " + $_.ToString()
	        Write-Warning $errormessage
	        $errormessage | out-file -append "C:\migration\_log.txt"
        }
    }
}

$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING: Clearing AppData Roaming"
Write-Host $output
$output | out-file -append "C:\migration\_log.txt"

$OfficeData = Test-Path -Path C:\Migration\officedata-recent-cleared.txt
If($OfficeData -eq $False) { 
    Try {
        $ErrorActionPreference = 'stop'
        $OfficeDataDir = "$env:APPDATA\Microsoft\Office\Recent"
        Remove-Item $OfficeDataDir\* -Recurse 
        # If no errors, log success
	    $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Removed recent folder: " + $OfficeDataDir
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item C:\Migration\officedata-recent-cleared.txt
    }
    Catch {
        # If there is an error, log the error
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING: Clearing Windows Credentials"
Write-Host $output
$output | out-file -append "C:\migration\_log.txt"

$Credentials = (cmdkey /list | Where-Object {$_ -like "*Target=MicrosoftOffice16_Data*"})
Foreach ($Target in $Credentials) {
    $Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
    $Argument = "/delete:" + $Target
    Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
    }

# Clear saved credentials
$Credentials = (cmdkey /list | Where-Object {$_ -like "*SSO_POP_Device*"})
Foreach ($Target in $Credentials) {
    $Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
    $Argument = "/delete:" + $Target
    Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
    }
Write-Host "****** Clearing SSO_POP_Device Credentials : Complete ******"

$Credentials = (cmdkey /list | Where-Object {$_ -like "*virtualapp/didlogical*"})
Foreach ($Target in $Credentials) {
    $Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
    $Argument = "/delete:" + $Target
    Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
    }
Write-Host "****** Clearing virtualapp/didlogical Credentials : Complete ******"

$Credentials = (cmdkey /list | Where-Object {$_ -like "*msteams*"})
Foreach ($Target in $Credentials) {
    $Target = ($Target -split (":", 2) | Select-Object -Skip 1).substring(1)
    $Argument = "/delete:" + $Target
    Start-Process Cmdkey -ArgumentList $Argument -NoNewWindow -RedirectStandardOutput $False
    }
Write-Host "****** Clearing Teams Credentials : Complete ******"

# Try to remove the Link School/Work account if there was one. It can be created if the first time you sign in, the user all
$timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
$output = $timestamp + " ***STARTING: Removal of 'Link School/Work account"
Write-Host $output
$output | out-file -append "C:\migration\_log.txt"

# Check if School/Work account was already removed by script
$SchoolWorkAccount = Test-Path -Path C:\Migration\school-work-account-cleared.txt
If ($SchoolWorkAccount -eq $false){
    # Delete folders to remove School/Work account link
    $LocalPackagesFolder ="$env:LOCALAPPDATA\Packages"
    $AADBrokerFolder = Get-ChildItem -Path $LocalPackagesFolder -Recurse -Include "Microsoft.AAD.BrokerPlugin_*";
    $AADBrokerFolder = $AADBrokerFolder[0];
    Try {
        $ErrorActionPreference = 'stop'
        Get-ChildItem "$AADBrokerFolder\AC\TokenBroker\Accounts" | Remove-Item -Recurse -Force
        # If no errors, log success
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $output = $timestamp + " Folder deleted: " + $AADBrokerFolder + "\AC\TokenBroker\Accounts"
        Write-Host $output
	    $output | out-file -append "C:\migration\_log.txt"
        New-Item -Path C:\Migration\school-work-account-cleared.txt
    }
    # If there was an error, log the error
    Catch {
        $timestamp=Get-Date -Format "MM/dd/yyyy HH:mm"
	    $errormessage=$timestamp + " ERROR: " + $_.ToString()
	    Write-Warning $errormessage
	    $errormessage | out-file -append "C:\migration\_log.txt"
    }
}

$stopwatch.Elapsed.TotalSeconds

$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Migration complete, please open Outlook, Teams, and OneDrive and sign in")

Read-Host -Prompt "Press any key to continue"

