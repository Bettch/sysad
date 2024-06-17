############################################################################################################
<#---------------------------------------------INITIAL SETUP----------------------------------------------#>
############################################################################################################

#no errors throughout
$ErrorActionPreference = 'silentlycontinue'


#Create Folder for logging output
$DebloatFolder = "C:\ProgramData\Debloat"
If (Test-Path $DebloatFolder) {
    Write-Output "$DebloatFolder exists. Skipping."
}
Else {
    Write-Output "The folder '$DebloatFolder' doesn't exist. This folder will be used for storing logs created after the script runs. Creating now."
    Start-Sleep 1
    New-Item -Path "$DebloatFolder" -ItemType Directory
    Write-Output "The folder $DebloatFolder was successfully created."
}

Start-Transcript -Path "C:\ProgramData\Debloat\Debloat.log"

############################################################################################################
<#------------------------------------------Remove AppX Packages------------------------------------------#>
############################################################################################################


    #Removes AppxPackages
    $WhitelistedApps = @(
        'Microsoft.Appconnector',
        "Microsoft.MicrosoftPowerBIForWindows",
        "Microsoft.MicrosoftStickyNotes",
        "Microsoft.Office.OneNote",
        "Microsoft.Windows.Photos",
        "Microsoft.WindowsCalculator",
        "Microsoft.WindowsStore",
        "Microsoft.MSPaint",
        'Microsoft.WindowsNotepad',
        'Microsoft.CompanyPortal',
        'Microsoft.ScreenSketch',
        'Microsoft.Paint3D',
        '.NET',
        'Framework',
        'Microsoft.HEIFImageExtension',
        'Microsoft.StorePurchaseApp',
        'Microsoft.VP9VideoExtensions',
        'Microsoft.WebMediaExtensions',
        'Microsoft.WebpImageExtension',
        'Microsoft.DesktopAppInstaller',
        'WindSynthBerry',
        'MIDIBerry'
    )
    ##If $customwhitelist is set, split on the comma and add to whitelist
    if ($customwhitelist) {
        $customWhitelistApps = $customwhitelist -split ","
        foreach ($whitelistapp in $customwhitelistapps) {
            ##Add to the array
            $WhitelistedApps += $whitelistapp
        }
    }
    
    #NonRemovable Apps that where getting attempted and the system would reject the uninstall, speeds up debloat and prevents 'initalizing' overlay when removing apps
    $NonRemovable = @(
        '1527c705-839a-4832-9118-54d4Bd6a0c89',
        'c5e2524a-ea46-4f67-841f-6a9465d9d515',
        'E2A4F912-2574-4A75-9BB0-0D023378592B',
        'F46D4000-FD22-4DB4-AC8E-4E1DDDE828FE',
        'InputApp',
        'Microsoft.AAD.BrokerPlugin',
        'Microsoft.AccountsControl',
        'Microsoft.BioEnrollment',
        'Microsoft.CredDialogHost',
        'Microsoft.ECApp',
        'Microsoft.LockApp',
        'Microsoft.MicrosoftEdgeDevToolsClient',
        'Microsoft.MicrosoftEdge',
        'Microsoft.PPIProjection',
        'Microsoft.Win32WebViewHost',
        'Microsoft.Windows.Apprep.ChxApp',
        'Microsoft.Windows.AssignedAccessLockApp',
        'Microsoft.Windows.CapturePicker',
        'Microsoft.Windows.CloudExperienceHost',
        'Microsoft.Windows.ContentDeliveryManager',
        'Microsoft.Windows.Cortana',
        'Microsoft.Windows.NarratorQuickStart',
        'Microsoft.Windows.ParentalControls',
        'Microsoft.Windows.PeopleExperienceHost',
        'Microsoft.Windows.PinningConfirmationDialog',
        'Microsoft.Windows.SecHealthUI',
        'Microsoft.Windows.SecureAssessmentBrowser',
        'Microsoft.Windows.ShellExperienceHost',
        'Microsoft.Windows.XGpuEjectDialog',
        'Microsoft.XboxGameCallableUI',
        'Windows.CBSPreview',
        'windows.immersivecontrolpanel',
        'Windows.PrintDialog',
        'Microsoft.VCLibs.140.00',
        'Microsoft.Services.Store.Engagement',
        'Microsoft.UI.Xaml.2.0',
        '*Nvidia*'
    )

    ##Combine the two arrays
    $appstoignore = $WhitelistedApps += $NonRemovable


    Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -notin $appstoignore} | Remove-AppxProvisionedPackage -Online
    Get-AppxPackage -AllUsers | Where-Object {$_.Name -notin $appstoignore} | Remove-AppxPackage


##Remove bloat
$Bloatware = @(
    #Unnecessary Windows 10/11 AppX Apps
    "Microsoft.549981C3F5F10"
    "Microsoft.BingNews"
    "Microsoft.BingFoodAndDrink",
    "Microsoft.BingHealthAndFitness",
    "Microsoft.BingTravel",
    "Microsoft.WindowsReadingList",
    "Microsoft.GetHelp",
    "Microsoft.Getstarted",
    "Microsoft.Messaging",
    "Microsoft.Microsoft3DViewer",
    "Microsoft.MicrosoftOfficeHub",
    "Microsoft.MicrosoftSolitaireCollection",
    "Microsoft.NetworkSpeedTest",
    "Microsoft.MixedReality.Portal",
    "Microsoft.News",
    "Microsoft.Office.Lens",
    "Microsoft.Office.Sway",
    "Microsoft.OneConnect",
    "Microsoft.People",
    "Microsoft.Print3D",
    "Microsoft.RemoteDesktop",
    "*Skype*",
    "Microsoft.StorePurchaseApp",
    "Microsoft.Office.Todo.List",
    "Microsoft.Whiteboard",
    "Microsoft.WindowsAlarms",
    "Microsoft.WindowsCamera",
    "Microsoft.windowscommunicationsapps",
    "Microsoft.WindowsFeedbackHub",
    "Microsoft.WindowsMaps",
    "Microsoft.WindowsSoundRecorder",
    "Microsoft.Xbox.TCUI",
    "Microsoft.XboxApp",
    "Microsoft.XboxGameOverlay",
    "Microsoft.XboxIdentityProvider",
    "Microsoft.XboxSpeechToTextOverlay",
    "Microsoft.ZuneMusic",
    "Microsoft.ZuneVideo",
    "Microsoft.YourPhone",
    "Microsoft.XboxGamingOverlay",
    "Microsoft.GamingApp",
    "Microsoft.Todos",
    "Microsoft.PowerAutomateDesktop",
    "Microsoft.Wallet",
    "SpotifyAB.SpotifyMusic",
    "Microsoft.MicrosoftJournal",
    "Disney.37853FC22B2CE",
    "*EclipseManager*",
    "*ActiproSoftwareLLC*",
    "*AdobeSystemsIncorporated.AdobePhotoshopExpress*",
    "*Duolingo-LearnLanguagesforFree*",
    "*PandoraMediaInc*",
    "*CandyCrush*",
    "*BubbleWitch3Saga*",
    "*Wunderlist*",
    "*Flipboard*",
    "*Twitter*",
    "*Facebook*",
    "*Spotify*",
    "*Minecraft*",
    "*Royal Revolt*",
    "*Sway*",
    "*Speed Test*",
    "*Dolby*",
    "*Disney*",
    "clipchamp.clipchamp",
    "*gaming*",
    "MicrosoftCorporationII.MicrosoftFamily",
    "C27EB4BA.DropboxOEM*",
    "*DevHome*",
    "MicrosoftCorporationII.QuickAssist",

    #Other non-windows apps
    "2FE3CB00.PicsArt-PhotoStudio",
    "46928bounde.EclipseManager",
    "4DF9E0F8.Netflix",
    "613EBCEA.PolarrPhotoEditorAcademicEdition",
    "6Wunderkinder.Wunderlist",
    "7EE7776C.LinkedInforWindows",
    "89006A2E.AutodeskSketchBook",
    "9E2F88E3.Twitter",
    "A278AB0D.DisneyMagicKingdoms",
    "A278AB0D.MarchofEmpires",
    "ActiproSoftwareLLC.562882FEEB491",
    "CAF9E577.Plex",
    "ClearChannelRadioDigital.iHeartRadio",
    "D52A8D61.FarmVille2CountryEscape",
    "D5EA27B7.Duolingo-LearnLanguagesforFree",
    "DB6EA5DB.CyberLinkMediaSuiteEssentials",
    "DolbyLaboratories.DolbyAccess",
    "DolbyLaboratories.DolbyAccess",
    "Drawboard.DrawboardPDF",
    "Facebook.Facebook",
    "Fitbit.FitbitCoach",
    "Flipboard.Flipboard",
    "GAMELOFTSA.Asphalt8Airborne",
    "KeeperSecurityInc.Keeper",
    "NORDCURRENT.COOKINGFEVER",
    "PandoraMediaInc.29680B314EFC2",
    "Playtika.CaesarsSlotsFreeCasino",
    "ShazamEntertainmentLtd.Shazam",
    "SlingTVLLC.SlingTV",
    "SpotifyAB.SpotifyMusic",
    "TheNewYorkTimes.NYTCrossword",
    "ThumbmunkeysLtd.PhototasticCollage",
    "TuneIn.TuneInRadio",
    "WinZipComputing.WinZipUniversal",
    "XINGAG.XING",
    "flaregamesGmbH.RoyalRevolt2",
    "king.com.*",
    "king.com.BubbleWitch3Saga",
    "king.com.CandyCrushSaga",
    "king.com.CandyCrushSodaSaga",
    "ZhuhaiKingsoftOfficeSoftw.WPSOffice2019",
    "A025C540.Yandex.Music"
)

foreach ($Bloat in $Bloatware) {
        
        if (Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $Bloat -ErrorAction SilentlyContinue) {
            Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $Bloat | Remove-AppxProvisionedPackage -Online
            Write-Host "Removed provisioned package for $Bloat."
        } else {
            Write-Host "Provisioned package for $Bloat not found."
        }

        if (Get-AppxPackage -Name $Bloat -ErrorAction SilentlyContinue) {
            Get-AppxPackage -allusers -Name $Bloat | Remove-AppxPackage -AllUsers
            Write-Host "Removed $Bloat."
        } else {
            Write-Host "$Bloat not found."
        }
}

############################################################################################################
<#------------------------------------------Remove REGISTRY KEYS------------------------------------------#>
############################################################################################################

##We need to grab all SIDs to remove at user level
$UserSIDs = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | Select-Object -ExpandProperty PSChildName

    
    #These are the registry keys that it will delete.
            
    $Keys = @(
            
        #Remove Background Tasks
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\46928bounde.EclipseManager_2.2.4.51_neutral__a5h4egax66k6y"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.MicrosoftOfficeHub_17.7909.7600.0_x64__8wekyb3d8bbwe"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.PPIProjection_10.0.15063.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.XboxGameCallableUI_1000.15063.0.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.XboxGameCallableUI_1000.16299.15.0_neutral_neutral_cw5n1h2txyewy"
            
        #Windows File
        "HKCR:\Extensions\ContractId\Windows.File\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
            
        #Registry keys to delete if they aren't uninstalled by RemoveAppXPackage/RemoveAppXProvisionedPackage
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\46928bounde.EclipseManager_2.2.4.51_neutral__a5h4egax66k6y"
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\Microsoft.PPIProjection_10.0.15063.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\Microsoft.XboxGameCallableUI_1000.15063.0.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\Microsoft.XboxGameCallableUI_1000.16299.15.0_neutral_neutral_cw5n1h2txyewy"
            
        #Scheduled Tasks to delete
        "HKCR:\Extensions\ContractId\Windows.PreInstalledConfigTask\PackageId\Microsoft.MicrosoftOfficeHub_17.7909.7600.0_x64__8wekyb3d8bbwe"
            
        #Windows Protocol Keys
        "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
        "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\Microsoft.PPIProjection_10.0.15063.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\Microsoft.XboxGameCallableUI_1000.15063.0.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\Microsoft.XboxGameCallableUI_1000.16299.15.0_neutral_neutral_cw5n1h2txyewy"
               
        #Windows Share Target
        "HKCR:\Extensions\ContractId\Windows.ShareTarget\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
    )
        
    #This writes the output of each key it is removing and also removes the keys listed above.
    ForEach ($Key in $Keys) {
        Write-Host "Removing $Key from registry"
        Remove-Item $Key -Recurse
    }


    #Disables Windows Feedback Experience
    Write-Host "Disabling Windows Feedback Experience program"
    $Advertising = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo"
    If (!(Test-Path $Advertising)) {
        New-Item $Advertising
    }
    If (Test-Path $Advertising) {
        Set-ItemProperty $Advertising Enabled -Value 0 
    }
            
    #Stops Cortana from being used as part of your Windows Search Function
    Write-Host "Stopping Cortana from being used as part of your Windows Search Function"
    $Search = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
    If (!(Test-Path $Search)) {
        New-Item $Search
    }
    If (Test-Path $Search) {
        Set-ItemProperty $Search AllowCortana -Value 0 
    }

    #Disables Web Search in Start Menu
    Write-Host "Disabling Bing Search in Start Menu"
    $WebSearch = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
    If (!(Test-Path $WebSearch)) {
        New-Item $WebSearch
    }
    Set-ItemProperty $WebSearch DisableWebSearch -Value 1 
    ##Loop through all user SIDs in the registry and disable Bing Search
    foreach ($sid in $UserSIDs) {
        $WebSearch = "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Search"
        If (!(Test-Path $WebSearch)) {
            New-Item $WebSearch
        }
        Set-ItemProperty $WebSearch BingSearchEnabled -Value 0
    }
    
    Set-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" BingSearchEnabled -Value 0 

            
    #Stops the Windows Feedback Experience from sending anonymous data
    Write-Host "Stopping the Windows Feedback Experience program"
    $Period = "HKCU:\Software\Microsoft\Siuf\Rules"
    If (!(Test-Path $Period)) { 
        New-Item $Period
    }
    Set-ItemProperty $Period PeriodInNanoSeconds -Value 0 

    ##Loop and do the same
    foreach ($sid in $UserSIDs) {
        $Period = "Registry::HKU\$sid\Software\Microsoft\Siuf\Rules"
        If (!(Test-Path $Period)) { 
            New-Item $Period
        }
        Set-ItemProperty $Period PeriodInNanoSeconds -Value 0 
    }

    #Prevents bloatware applications from returning and removes Start Menu suggestions               
    Write-Host "Adding Registry key to prevent bloatware apps from returning"
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
    $registryOEM = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    If (!(Test-Path $registryPath)) { 
        New-Item $registryPath
    }
    Set-ItemProperty $registryPath DisableWindowsConsumerFeatures -Value 1 

    If (!(Test-Path $registryOEM)) {
        New-Item $registryOEM
    }
    Set-ItemProperty $registryOEM  ContentDeliveryAllowed -Value 0 
    Set-ItemProperty $registryOEM  OemPreInstalledAppsEnabled -Value 0 
    Set-ItemProperty $registryOEM  PreInstalledAppsEnabled -Value 0 
    Set-ItemProperty $registryOEM  PreInstalledAppsEverEnabled -Value 0 
    Set-ItemProperty $registryOEM  SilentInstalledAppsEnabled -Value 0 
    Set-ItemProperty $registryOEM  SystemPaneSuggestionsEnabled -Value 0  
    
    ##Loop through users and do the same
    foreach ($sid in $UserSIDs) {
        $registryOEM = "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
        If (!(Test-Path $registryOEM)) {
            New-Item $registryOEM
        }
        Set-ItemProperty $registryOEM  ContentDeliveryAllowed -Value 0 
        Set-ItemProperty $registryOEM  OemPreInstalledAppsEnabled -Value 0 
        Set-ItemProperty $registryOEM  PreInstalledAppsEnabled -Value 0 
        Set-ItemProperty $registryOEM  PreInstalledAppsEverEnabled -Value 0 
        Set-ItemProperty $registryOEM  SilentInstalledAppsEnabled -Value 0 
        Set-ItemProperty $registryOEM  SystemPaneSuggestionsEnabled -Value 0 
    }
    
    #Preping mixed Reality Portal for removal    
    Write-Host "Setting Mixed Reality Portal value to 0 so that you can uninstall it in Settings"
    $Holo = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Holographic"    
    If (Test-Path $Holo) {
        Set-ItemProperty $Holo  FirstRunSucceeded -Value 0 
    }

    ##Loop through users and do the same
    foreach ($sid in $UserSIDs) {
        $Holo = "Registry::HKU\$sid\Software\Microsoft\Windows\CurrentVersion\Holographic"    
        If (Test-Path $Holo) {
            Set-ItemProperty $Holo  FirstRunSucceeded -Value 0 
        }
    }

    #Disables Wi-fi Sense
    Write-Host "Disabling Wi-Fi Sense"
    $WifiSense1 = "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting"
    $WifiSense2 = "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowAutoConnectToWiFiSenseHotspots"
    $WifiSense3 = "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config"
    If (!(Test-Path $WifiSense1)) {
        New-Item $WifiSense1
    }
    Set-ItemProperty $WifiSense1  Value -Value 0 
    If (!(Test-Path $WifiSense2)) {
        New-Item $WifiSense2
    }
    Set-ItemProperty $WifiSense2  Value -Value 0 
    Set-ItemProperty $WifiSense3  AutoConnectAllowedOEM -Value 0 
        
    #Disables live tiles
    Write-Host "Disabling live tiles"
    $Live = "HKCU:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\PushNotifications"    
    If (!(Test-Path $Live)) {      
        New-Item $Live
    }
    Set-ItemProperty $Live  NoTileApplicationNotification -Value 1 

    ##Loop through users and do the same
    foreach ($sid in $UserSIDs) {
        $Live = "Registry::HKU\$sid\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\PushNotifications"    
        If (!(Test-Path $Live)) {      
            New-Item $Live
        }
        Set-ItemProperty $Live  NoTileApplicationNotification -Value 1 
    }
        
###Enable location tracking for "find my device",

    #Disabling Location Tracking
    #Write-Host "Disabling Location Tracking"
    #$SensorState = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}"
    #$LocationConfig = "HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration"
    #If (!(Test-Path $SensorState)) {
    #    New-Item $SensorState
    #}
    #Set-ItemProperty $SensorState SensorPermissionState -Value 0 
    #If (!(Test-Path $LocationConfig)) {
    #    New-Item $LocationConfig
    #}
    #Set-ItemProperty $LocationConfig Status -Value 0 
        
    #Disables People icon on Taskbar
    Write-Host "Disabling People icon on Taskbar"
    $People = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People'
    If (Test-Path $People) {
        Set-ItemProperty $People -Name PeopleBand -Value 0
    }

    #Loop through users and do the same
    foreach ($sid in $UserSIDs) {
        $People = "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People"
        If (Test-Path $People) {
            Set-ItemProperty $People -Name PeopleBand -Value 0
        }
    }

    Write-Host "Disabling Cortana"
    $Cortana1 = "HKCU:\SOFTWARE\Microsoft\Personalization\Settings"
    $Cortana2 = "HKCU:\SOFTWARE\Microsoft\InputPersonalization"
    $Cortana3 = "HKCU:\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore"
    If (!(Test-Path $Cortana1)) {
        New-Item $Cortana1
    }
    Set-ItemProperty $Cortana1 AcceptedPrivacyPolicy -Value 0 
    If (!(Test-Path $Cortana2)) {
        New-Item $Cortana2
    }
    Set-ItemProperty $Cortana2 RestrictImplicitTextCollection -Value 1 
    Set-ItemProperty $Cortana2 RestrictImplicitInkCollection -Value 1 
    If (!(Test-Path $Cortana3)) {
        New-Item $Cortana3
    }
    Set-ItemProperty $Cortana3 HarvestContacts -Value 0

    ##Loop through users and do the same
    foreach ($sid in $UserSIDs) {
        $Cortana1 = "Registry::HKU\$sid\SOFTWARE\Microsoft\Personalization\Settings"
        $Cortana2 = "Registry::HKU\$sid\SOFTWARE\Microsoft\InputPersonalization"
        $Cortana3 = "Registry::HKU\$sid\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore"
        If (!(Test-Path $Cortana1)) {
            New-Item $Cortana1
        }
        Set-ItemProperty $Cortana1 AcceptedPrivacyPolicy -Value 0 
        If (!(Test-Path $Cortana2)) {
            New-Item $Cortana2
        }
        Set-ItemProperty $Cortana2 RestrictImplicitTextCollection -Value 1 
        Set-ItemProperty $Cortana2 RestrictImplicitInkCollection -Value 1 
        If (!(Test-Path $Cortana3)) {
            New-Item $Cortana3
        }
        Set-ItemProperty $Cortana3 HarvestContacts -Value 0
    }


    #Removes 3D Objects from the 'My Computer' submenu in explorer
    Write-Host "Removing 3D Objects from explorer 'My Computer' submenu"
    $Objects32 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}"
    $Objects64 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}"
    If (Test-Path $Objects32) {
        Remove-Item $Objects32 -Recurse 
    }
    If (Test-Path $Objects64) {
        Remove-Item $Objects64 -Recurse 
    }

   
    ##Removes the Microsoft Feeds from displaying
$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds"
$Name = "EnableFeeds"
$value = "0"

if (!(Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
}

else {
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
}

##Kill Cortana again
Get-AppxPackage - allusers Microsoft.549981C3F5F10 | Remove AppxPackage
    
############################################################################################################
<#-----------------------------------------REMOVE SCHEDULED TASKS-----------------------------------------#>
############################################################################################################

    #Disables scheduled tasks that are considered unnecessary 
    Write-Host "Disabling scheduled tasks"
    $task1 = Get-ScheduledTask -TaskName XblGameSaveTaskLogon -ErrorAction SilentlyContinue
    if ($null -ne $task1) {
    Get-ScheduledTask  XblGameSaveTaskLogon | Disable-ScheduledTask -ErrorAction SilentlyContinue
    }
    $task2 = Get-ScheduledTask -TaskName XblGameSaveTask -ErrorAction SilentlyContinue
    if ($null -ne $task2) {
    Get-ScheduledTask  XblGameSaveTask | Disable-ScheduledTask -ErrorAction SilentlyContinue
    }
    $task3 = Get-ScheduledTask -TaskName Consolidator -ErrorAction SilentlyContinue
    if ($null -ne $task3) {
    Get-ScheduledTask  Consolidator | Disable-ScheduledTask -ErrorAction SilentlyContinue
    }
    $task4 = Get-ScheduledTask -TaskName UsbCeip -ErrorAction SilentlyContinue
    if ($null -ne $task4) {
    Get-ScheduledTask  UsbCeip | Disable-ScheduledTask -ErrorAction SilentlyContinue
    }
    $task5 = Get-ScheduledTask -TaskName DmClient -ErrorAction SilentlyContinue
    if ($null -ne $task5) {
    Get-ScheduledTask  DmClient | Disable-ScheduledTask -ErrorAction SilentlyContinue
    }
    $task6 = Get-ScheduledTask -TaskName DmClientOnScenarioDownload -ErrorAction SilentlyContinue
    if ($null -ne $task6) {
    Get-ScheduledTask  DmClientOnScenarioDownload | Disable-ScheduledTask -ErrorAction SilentlyContinue
    }


############################################################################################################
<#--------------------------------------------DISABLE SERVICES--------------------------------------------#>
############################################################################################################

    Write-Host "Stopping and disabling Diagnostics Tracking Service"
    #Disabling the Diagnostics Tracking Service
    Stop-Service "DiagTrack"
    Set-Service "DiagTrack" -StartupType Disabled

############################################################################################################
<#-----------------------------------------DISABLE WINDOWS BACKUP-----------------------------------------#>
############################################################################################################

# $version = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption
# if ($version -like "*Windows 10*") {
#     write-host "Removing Windows Backup"
#     $filepath = "C:\Windows\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\WindowsBackup\Assets"
# if (Test-Path $filepath) {
# Remove-WindowsPackage -Online -PackageName "Microsoft-Windows-UserExperience-Desktop-Package~31bf3856ad364e35~amd64~~10.0.19041.3393"

# ##Add back snipping tool functionality
# write-host "Adding Windows Shell Components"
# DISM /Online /Add-Capability /CapabilityName:Windows.Client.ShellComponents~~~~0.0.1.0
# write-host "Components Added"
# }
# write-host "Removed"
# }

############################################################################################################
<#-----------------------------------------REMOVE WINDOWS COPILOT-----------------------------------------#>
############################################################################################################

# $version = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption
# if ($version -like "*Windows 11*") {
#     write-host "Removing Windows Copilot"
# # Define the registry key and value
# $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"
# $propertyName = "TurnOffWindowsCopilot"
# $propertyValue = 1

# # Check if the registry key exists
# if (!(Test-Path $registryPath)) {
#     # If the registry key doesn't exist, create it
#     New-Item -Path $registryPath -Force | Out-Null
# }

# # Get the property value
# $currentValue = Get-ItemProperty -Path $registryPath -Name $propertyName -ErrorAction SilentlyContinue

# # Check if the property exists and if its value is different from the desired value
# if ($null -eq $currentValue -or $currentValue.$propertyName -ne $propertyValue) {
#     # If the property doesn't exist or its value is different, set the property value
#     Set-ItemProperty -Path $registryPath -Name $propertyName -Value $propertyValue
# }


# ##Grab the default user as well
# $registryPath = "HKEY_USERS\.DEFAULT\Software\Policies\Microsoft\Windows\WindowsCopilot"
# $propertyName = "TurnOffWindowsCopilot"
# $propertyValue = 1

# # Check if the registry key exists
# if (!(Test-Path $registryPath)) {
#     # If the registry key doesn't exist, create it
#     New-Item -Path $registryPath -Force | Out-Null
# }

# # Get the property value
# $currentValue = Get-ItemProperty -Path $registryPath -Name $propertyName -ErrorAction SilentlyContinue

# # Check if the property exists and if its value is different from the desired value
# if ($null -eq $currentValue -or $currentValue.$propertyName -ne $propertyValue) {
#     # If the property doesn't exist or its value is different, set the property value
#     Set-ItemProperty -Path $registryPath -Name $propertyName -Value $propertyValue
# }


# ##Load the default hive from c:\users\Default\NTUSER.dat
# reg load HKU\temphive "c:\users\default\ntuser.dat"
# $registryPath = "registry::hku\temphive\Software\Policies\Microsoft\Windows\WindowsCopilot"
# $propertyName = "TurnOffWindowsCopilot"
# $propertyValue = 1

# # Check if the registry key exists
# if (!(Test-Path $registryPath)) {
#     # If the registry key doesn't exist, create it
#     [Microsoft.Win32.RegistryKey]$HKUCoPilot = [Microsoft.Win32.Registry]::Users.CreateSubKey("temphive\Software\Policies\Microsoft\Windows\WindowsCopilot", [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree)
#     $HKUCoPilot.SetValue("TurnOffWindowsCopilot", 0x1, [Microsoft.Win32.RegistryValueKind]::DWord)
# }

#     $HKUCoPilot.Flush()
#     $HKUCoPilot.Close()
# [gc]::Collect()
# [gc]::WaitForPendingFinalizers()
# reg unload HKU\temphive


# write-host "Removed"


# foreach ($sid in $UserSIDs) {
#     $registryPath = "Registry::HKU\$sid\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"
#     $propertyName = "TurnOffWindowsCopilot"
#     $propertyValue = 1
    
#     # Check if the registry key exists
#     if (!(Test-Path $registryPath)) {
#         # If the registry key doesn't exist, create it
#         New-Item -Path $registryPath -Force | Out-Null
#     }
    
#     # Get the property value
#     $currentValue = Get-ItemProperty -Path $registryPath -Name $propertyName -ErrorAction SilentlyContinue
    
#     # Check if the property exists and if its value is different from the desired value
#     if ($null -eq $currentValue -or $currentValue.$propertyName -ne $propertyValue) {
#         # If the property doesn't exist or its value is different, set the property value
#         Set-ItemProperty -Path $registryPath -Name $propertyName -Value $propertyValue
#     }
# }
# }

############################################################################################################
<#-------------------------------------------REMOVE XBOX GAMING-------------------------------------------#>
############################################################################################################

New-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\xbgm" -Name "Start" -PropertyType DWORD -Value 4 -Force
Set-Service -Name XblAuthManager -StartupType Disabled
Set-Service -Name XblGameSave -StartupType Disabled
Set-Service -Name XboxGipSvc -StartupType Disabled
Set-Service -Name XboxNetApiSvc -StartupType Disabled
$task = Get-ScheduledTask -TaskName "Microsoft\XblGameSave\XblGameSaveTask" -ErrorAction SilentlyContinue
if ($null -ne $task) {
Set-ScheduledTask -TaskPath $task.TaskPath -Enabled $false
}

##Check if GamePresenceWriter.exe exists
if (Test-Path "$env:WinDir\System32\GameBarPresenceWriter.exe") {
    write-host "GamePresenceWriter.exe exists"
    C:\Windows\Temp\SetACL.exe -on  "$env:WinDir\System32\GameBarPresenceWriter.exe" -ot file -actn setowner -ownr "n:$everyone"
C:\Windows\Temp\SetACL.exe -on  "$env:WinDir\System32\GameBarPresenceWriter.exe" -ot file -actn ace -ace "n:$everyone;p:full"

#Take-Ownership -Path "$env:WinDir\System32\GameBarPresenceWriter.exe"
$NewAcl = Get-Acl -Path "$env:WinDir\System32\GameBarPresenceWriter.exe"
# Set properties
$identity = "$builtin\Administrators"
$fileSystemRights = "FullControl"
$type = "Allow"
# Create new rule
$fileSystemAccessRuleArgumentList = $identity, $fileSystemRights, $type
$fileSystemAccessRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $fileSystemAccessRuleArgumentList
# Apply new rule
$NewAcl.SetAccessRule($fileSystemAccessRule)
Set-Acl -Path "$env:WinDir\System32\GameBarPresenceWriter.exe" -AclObject $NewAcl
Stop-Process -Name "GameBarPresenceWriter.exe" -Force
Remove-Item "$env:WinDir\System32\GameBarPresenceWriter.exe" -Force -Confirm:$false

}
else {
    write-host "GamePresenceWriter.exe does not exist"
}

New-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\GameDVR" -Name "AllowgameDVR" -PropertyType DWORD -Value 0 -Force
New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "SettingsPageVisibility" -PropertyType String -Value "hide:gaming-gamebar;gaming-gamedvr;gaming-broadcasting;gaming-gamemode;gaming-xboxnetworking" -Force
Remove-Item C:\Windows\Temp\SetACL.exe -recurse

############################################################################################################
<#------------------------------------------REMOVE EDGE SURF GAME-----------------------------------------#>
############################################################################################################

$surf = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge"
If (!(Test-Path $surf)) {
    New-Item $surf
}
New-ItemProperty -Path $surf -Name 'AllowSurfGame' -Value 0 -PropertyType DWord

############################################################################################################
<#---------------------------------------GRAB ALL UNINSTALL STRINGS---------------------------------------#>
############################################################################################################

#LOCAL MACHINE
write-host "Checking 32-bit System Registry"
#Search for 32-bit versions and list them
$allstring = @()
$path1 =  "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
#Loop Through the apps
$32apps = Get-ChildItem -Path $path1 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

foreach ($32app in $32apps) {
#Get uninstall string
$string1 =  $32app.uninstallstring
#Check if it's an MSI install
if ($string1 -match "^msiexec*") {
#MSI install, replace the I with an X and make it quiet
$string2 = $string1 + " /quiet /norestart"
$string2 = $string2 -replace "/I", "/X "
#Create custom object with name and string
$allstring += New-Object -TypeName PSObject -Property @{
    Name = $32app.DisplayName
    String = $string2
}
}
else {
#Exe installer, run straight path
$string2 = $string1
$allstring += New-Object -TypeName PSObject -Property @{
    Name = $32app.DisplayName
    String = $string2
}
}

}
write-host "32-bit check complete"

write-host "Checking 64-bit System registry"
##Search for 64-bit versions and list them

$path2 =  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
#Loop Through the apps
$64apps = Get-ChildItem -Path $path2 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

foreach ($64app in $64apps) {
#Get uninstall string
$string1 =  $64app.uninstallstring
#Check if it's an MSI install
if ($string1 -match "^msiexec*") {
#MSI install, replace the I with an X and make it quiet
$string2 = $string1 + " /quiet /norestart"
$string2 = $string2 -replace "/I", "/X "
#Uninstall with string2 params
$allstring += New-Object -TypeName PSObject -Property @{
    Name = $64app.DisplayName
    String = $string2
}
}
else {
#Exe installer, run straight path
$string2 = $string1
$allstring += New-Object -TypeName PSObject -Property @{
    Name = $64app.DisplayName
    String = $string2
}
}

}
write-host "64-bit checks complete"

#CURRENT USER
write-host "Checking 32-bit User Registry"
#Search for 32-bit versions and list them
$path1 =  "HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
#Check if path exists
if (Test-Path $path1) {
#Loop Through the apps
$32apps = Get-ChildItem -Path $path1 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

foreach ($32app in $32apps) {
#Get uninstall string
$string1 =  $32app.uninstallstring
#Check if it's an MSI install
if ($string1 -match "^msiexec*") {
#MSI install, replace the I with an X and make it quiet
$string2 = $string1 + " /quiet /norestart"
$string2 = $string2 -replace "/I", "/X "
#Create custom object with name and string
$allstring += New-Object -TypeName PSObject -Property @{
    Name = $32app.DisplayName
    String = $string2
}
}
else {
#Exe installer, run straight path
$string2 = $string1
$allstring += New-Object -TypeName PSObject -Property @{
    Name = $32app.DisplayName
    String = $string2
}
}
}
}
write-host "32-bit check complete"

write-host "Checking 64-bit Use registry"
#Search for 64-bit versions and list them

$path2 =  "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
#Loop Through the apps
$64apps = Get-ChildItem -Path $path2 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

foreach ($64app in $64apps) {
#Get uninstall string
$string1 =  $64app.uninstallstring
#Check if it's an MSI install
if ($string1 -match "^msiexec*") {
#MSI install, replace the I with an X and make it quiet
$string2 = $string1 + " /quiet /norestart"
$string2 = $string2 -replace "/I", "/X "
#Uninstall with string2 params
$allstring += New-Object -TypeName PSObject -Property @{
    Name = $64app.DisplayName
    String = $string2
}
}
else {
#Exe installer, run straight path
$string2 = $string1
$allstring += New-Object -TypeName PSObject -Property @{
    Name = $64app.DisplayName
    String = $string2
}
}

}

# ALL USERS
write-host "Checking 32-bit User Registry"
#Search for 32-bit versions and list them
foreach ($sid in $UserSIDs) {
    $path1 =  "Registry::HKU\$sid\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    #Check if path exists
    if (Test-Path $path1) {
        #Loop Through the apps
        $32apps = Get-ChildItem -Path $path1 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

        foreach ($32app in $32apps) {
            #Get uninstall string
            $string1 =  $32app.uninstallstring
            #Check if it's an MSI install
            if ($string1 -match "^msiexec*") {
            #MSI install, replace the I with an X and make it quiet
            $string2 = $string1 + " /quiet /norestart"
            $string2 = $string2 -replace "/I", "/X "
            #Create custom object with name and string
            $allstring += New-Object -TypeName PSObject -Property @{
                Name = $32app.DisplayName
                String = $string2
            }
            }
            else {
            #Exe installer, run straight path
            $string2 = $string1
            $allstring += New-Object -TypeName PSObject -Property @{
                Name = $32app.DisplayName
                String = $string2
            }
            }
        }
    }
}
write-host "32-bit check complete"

write-host "Checking 64-bit Use registry"
#Search for 64-bit versions and list them
foreach ($sid in $UserSIDs) {
    $path2 =  "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    #Loop Through the apps
    $64apps = Get-ChildItem -Path $path2 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

    foreach ($64app in $64apps) {
        #Get uninstall string
        $string1 =  $64app.uninstallstring
        #Check if it's an MSI install
        if ($string1 -match "^msiexec*") {
        #MSI install, replace the I with an X and make it quiet
        $string2 = $string1 + " /quiet /norestart"
        $string2 = $string2 -replace "/I", "/X "
        #Uninstall with string2 params
        $allstring += New-Object -TypeName PSObject -Property @{
            Name = $64app.DisplayName
            String = $string2
        }
        }
        else {
        #Exe installer, run straight path
        $string2 = $string1
        $allstring += New-Object -TypeName PSObject -Property @{
            Name = $64app.DisplayName
            String = $string2
        }
        }
    }
}

############################################################################################################
<#------------------------------------------REMOVE SELECTED APPS------------------------------------------#>
############################################################################################################

##Apps to remove - NOTE: Some applications like Chrome, Opera and WPS Office seem to have an unusual uninstall so needed to sort on their own

$blacklistapps = @(
    "*WPS Office*",
    "*Opera*",
    "*Firefox*",
    "*LibreOffice*",
    "*Acrobat*",
    "*Chrome*"
)

foreach ($app in $allstring) {
    $displayName = $app.Name
    $uninstallString = $app.String

    foreach ($blacklistedApp in $blacklistapps) {
        if ($displayName -like $blacklistedApp -and $uninstallString -like "*msiexec*") {
            $msiCode = $uninstallString.Split("{")[1]
            $uninstallCommand = "msiexec.exe /X {$msiCode} /qn"
            Write-Host "Uninstalling $displayName"
            Get-Process -Name "*$displayName*" -ErrorAction SilentlyContinue | Stop-Process -Force
            Get-Process -Name "*$blacklistedApp*" -ErrorAction SilentlyContinue | Stop-Process -Force
            Start-Sleep -Seconds 2
            Start-Process "cmd.exe" -ArgumentList "/c $uninstallCommand"
            break
        }
        elseif ($displayName -like $blacklistedApp -and $uninstallString -like "*.exe*") {
            $uninstallCommand = "$uninstallString /quiet --uninstall --runimmediately /S --multi-install --chrome --system-level --force-uninstall"
            Write-Host "Uninstalling $displayName"
            Get-Process -Name "*$displayName*" -ErrorAction SilentlyContinue | Stop-Process -Force
            Get-Process -Name "*$blacklistedApp*" -ErrorAction SilentlyContinue | Stop-Process -Force
            Start-Sleep -Seconds 2
            Start-Process "cmd.exe" -ArgumentList "/c $uninstallCommand"
            break
        }
    }
}

Start-Sleep -Seconds 10

# #Remove Chrome
# $chromeLM32path = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"

# if ($null -ne $chromeLM32path) {

# $versions = (Get-ItemProperty -path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome').version
# ForEach ($version in $versions) {
# write-host "Found Chrome version $version"
# $directory = ${env:ProgramFiles(x86)}
# write-host "Removing Chrome"
# Start-Process "$directory\Google\Chrome\Application\$version\Installer\setup.exe" -argumentlist  "--uninstall --multi-install --chrome --system-level --force-uninstall"
# }

# }

# $chromeLM64path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"

# if ($null -ne $chromeLM64path) {

# $versions = (Get-ItemProperty -path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome').version
# ForEach ($version in $versions) {
# write-host "Found Chrome version $version"
# $directory = ${env:ProgramFiles}
# write-host "Removing Chrome"
# Start-Process "$directory\Google\Chrome\Application\$version\Installer\setup.exe" -argumentlist  "--uninstall --multi-install --chrome --system-level --force-uninstall"
# }
# }

# Remove Opera Browser (All User Profiles)

$PotentialInstallLocations = New-Object System.Collections.Generic.List[Object]

$ProgramFiles = Get-ChildItem $env:ProgramFiles -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer -eq $true -and $_.Name -like "*Opera*" } | Select-Object Fullname
if ($ProgramFiles) { $ProgramFiles | ForEach-Object { $PotentialInstallLocations.Add($_) } }

$ProgramFilesX86 = Get-ChildItem ${env:ProgramFiles(x86)} -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer -eq $true -and $_.Name -like "*Opera*" } | Select-Object Fullname
if ($ProgramFilesX86) { $ProgramFilesX86 | ForEach-Object { $PotentialInstallLocations.Add($_) } }

$AppData = Get-ChildItem C:\Users\*\AppData\Local\Programs -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer -eq $true -and $_.Name -like "*Opera*" } | Select-Object Fullname
if ($AppData) { $AppData | ForEach-Object { $PotentialInstallLocations.Add($_) } }

if ($installDirectory) {
  $Directory = Get-Item $installDirectory -ErrorAction SilentlyContinue
  if ($Directory) { $PotentialInstallLocations.Add($Directory) }
}

if ($PotentialInstallLocations) {
  $OperaExe = $PotentialInstallLocations | ForEach-Object { Get-ChildItem $_.FullName | Where-Object { $_.Name -like "opera.exe" } }
  $LauncherExe = $OperaExe | ForEach-Object { "$($_.Directory)\opera.exe"}
}

if ($LauncherExe) {
  Write-Host "Opera installations found! Below are the install locations."
  $LauncherExe | Write-Host

  # Killing All Opera Processes
  Write-Warning "Killing all Opera Processes for uninstall."
  Get-Process "opera" -ErrorAction SilentlyContinue | Stop-Process -Force
  Get-Process "OperaSetup" -ErrorAction SilentlyContinue | Stop-Process -Force

  if ($DeleteUserProfile) {
    Write-Warning "Delete User Browser Profile Selected!"
    $Arguments = "--uninstall", "--runimmediately", "--deleteuserprofile=1"
  }
  else {
    $Arguments = "--uninstall", "--runimmediately", "--deleteuserprofile=0"
  }

  $Process = $LauncherExe | ForEach-Object {
    Start-Process -Wait $_ -ArgumentList $Arguments -PassThru 
  }
  Write-Host "Exit Code(s): $($Process.ExitCode)"

  $Process | ForEach-Object {
    switch ($_.ExitCode) {
      0 { Write-Host "Opera removal Success! Please note that Opera will still be visible in the Control Panel however it won't prompt for admin to remove it from the control panel when clicking uninstall." }
      default {
        Write-Error "Exit code does not indicate success"
        exit 1
      }
    }
  }

  $DesktopIcons = Get-ChildItem -Recurse "C:\Users\*\Desktop" -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "Opera Browser.lnk" } | Select-Object FullName

  if ($DesktopIcons) {
    Write-Host "Cleaning up Desktop icons."
    Write-Host "### Desktop Icon Locations ###"
    $DesktopIcons.FullName | Write-Host
    $DesktopIcons | ForEach-Object {
      Remove-Item $_.FullName
    }
  }
  else {
    Write-Host "No Desktop Icons found!"
  }

  $StartMenu = Get-ChildItem -Recurse "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs" -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "Opera Browser.lnk" } | Select-Object FullName

  if ($StartMenu) {
    Write-Host "Cleaning up Start Menu."
    Write-Host "### StartMenu Locations ###"
    $StartMenu.FullName | Write-Host
    $StartMenu | ForEach-Object {
      Remove-Item $_.FullName
    }
  }
  else {
    Write-Host "No Start Menu Entries found!"
  }
}
else {
  Write-Error "No installations found in C:\Users\*\AppData\Local\Programs, C:\Program Files or C:\Program Files(x86). Maybe give an installation directory using -InstallDirectory 'C:\ReplaceMe' ?"
  Exit 1
}

# Remove WPS Office (All User Profiles)
$Users = Get-ChildItem C:\Users
foreach ($user in $Users){
$WPSLocal = "$($user.fullname)\AppData\Local\Kingsoft\WPS Office"
If (Test-Path $WPSLocal) {
$UninstPath = Get-ChildItem -Path "$WPSLocal\*" -Include uninst.exe -Recurse -ErrorAction SilentlyContinue
If($UninstPath.Exists)
{
Write-Log -Message "Found $($UninstPath.FullName), now attempting to uninstall $installTitle."
Execute-ProcessAsUser -Path "$UninstPath" -Parameters "/S" -Wait
Get-Process -Name "Au_" -ErrorAction SilentlyContinue | Wait-Process
## Cleanup User Profile Registry
[scriptblock]$HKCURegistrySettings = {
Remove-RegistryKey -Key 'HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\Kingsoft Office' -SID $UserProfile.SID
}
Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings -ErrorAction SilentlyContinue
}
}
}
## Restart Explorer.exe
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
$ExplorerProcess = Get-Process Explorer -ErrorAction SilentlyContinue
If($null -eq $ExplorerProcess)
{
Execute-ProcessAsUser -Path "$envSystemRoot\explorer.exe"
}
## Sleep for 5 Seconds
Start-Sleep -Seconds 5
$Users = Get-ChildItem C:\Users
foreach ($user in $Users){
## Cleanup Kingsoft (Local User Profile) Directory
$KingsoftLocal = "$($user.fullname)\AppData\Local\Kingsoft"
If (Test-Path $KingsoftLocal) {
Write-Log -Message "Cleanup ($KingsoftLocal) Directory."
Remove-Item -Path "$KingsoftLocal" -Force -Recurse -ErrorAction SilentlyContinue 
}
## Cleanup Kingsoft (Roaming User Profile) Directory
$KingsoftRoaming = "$($user.fullname)\AppData\Roaming\kingsoft"
If (Test-Path $KingsoftRoaming) {
Write-Log -Message "Cleanup ($KingsoftRoaming) Directory."
Remove-Item -Path "$KingsoftRoaming" -Force -Recurse -ErrorAction SilentlyContinue 
}
## Remove WPS Office Start Menu Shortcut From All Profiles
$StartMenuSC = "$($user.fullname)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\WPS*"
If (Test-Path $StartMenuSC) {
Remove-Item $StartMenuSC -Recurse -Force -ErrorAction SilentlyContinue
}
## Remove WPS Office Desktop Shortcuts From All Profiles
$DesktopSC = "$($user.fullname)\Desktop\WPS*.lnk"
If (Test-Path $DesktopSC) {
Remove-Item $DesktopSC -Recurse -Force -ErrorAction SilentlyContinue
}
## Remove WPS Office TaskBar Shortcut From All Profiles
$TaskBarSC = "$($user.fullname)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\WPS Office.lnk"
If (Test-Path $TaskBarSC) {
Remove-Item $TaskBarSC -Recurse -Force -ErrorAction SilentlyContinue
}
}

write-host "Completed"

Stop-Transcript