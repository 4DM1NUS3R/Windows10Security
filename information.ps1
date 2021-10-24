#Report path
$reportPath = $PSScriptRoot + "\report.html"

############################# Display general information #############################

#Get Status 
if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $env:computername -Quiet) {
        $Status = "Offline"
    } else {
        $Status = "Online"
    }

#Get OS Informations
$OSInfo = Get-WmiObject -ComputerName $env:computername -ClassName Win32_OperatingSystem | Select-Object Caption, Version, OSArchitecture


#Get VM
$InfoModel = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $env:computername | Select Manufacturer, Model
if  ($InfoModel){
        $resultvm= 'False'
    } else {
        $resultvm = 'True'
    }


#Get DateBuilt
$DateBuilt =([WMI]'').ConvertToDateTime((Get-WmiObject -ComputerName $env:computername -ClassName Win32_OperatingSystem).InstallDate)

#Get LastBootTime
$LastBootTime = Get-CimInstance -ClassName Win32_OperatingSystem | Select lastbootuptime

#Get IP Adress
$adresses = Get-WmiObject win32_networkadapterconfiguration | Select-Object -Property @{
Name = 'IPAddress'
    Expression = {($PSItem.IPAddress[0])}
},MacAddress | 
Where IPAddress -NE $null

#Creating IP adresses and MAC adresses objects
$adresssesHtml = $adresses | ConvertTo-Html -Fragment -PreContent "<h2>Adresses IP et adresses MAC</h2>"

#Creating Computer object
$computerProps = @{
	'Name'= $env:computername;
	'Operating System'= $OSInfo.Version;
    'OS Caption'= $OSInfo.Caption
	'Architecture'= $OSInfo.OSArchitecture;
    'Status' = $Status;
	'VM'= $resultvm;
    'Model Name'=$InfoModel.Model;
    'Manufacturer'=$InfoModel.Manufacturer;
    'Date Built'= $DateBuilt
	'Last Boot'= $LastBootTime.lastbootuptime
}
$computer = New-Object -TypeName PSObject -Prop $computerProps
$computerHtml = $computer | ConvertTo-Html -Fragment -PreContent "<h1>Rapport d'audit de sécurité Windows 10 </h1><h2>Informations générales</h2>"

############################# Display information about local accounts #############################

#create user informations as array of objetcs containing informations
$localUserData = @()
$localUserDataHtml = @()
foreach($user in Get-LocalUser){
    $localUserData += $user | Select-Object Name,FullName,SID,LastLogon,PasswordExpire,PasswordLastSet,PasswordChangeableDate,AccountExpires,Enabled
    #Write-Host "File permissions :`n"
    #.\accesschk.exe $user c:\* -s
}
#convert array to html tables 
$localUserData | ForEach-Object {$localUserDataHtml += $PSItem | ConvertTo-Html -Fragment -PostContent "<a><br></a>"}

############################# Getting private life properties #############################
#The try catch methode was used for some parameters because for some of them, the property causes an error when the parameter is activated.

#Advertising Info
$NameAdInfo = "Advertising ID"
$StateAdInfo = Get-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo | Select-Object -ExpandProperty Enabled

if($StateAdInfo -eq 1){
    $StateAdInfo = "Le paramètre est activé"
}    
else {
        $StateAdInfo = "Le paramètre est désactivé"
    }



#Creation of hash table containing all name and status about parameters
$hash = @{}
$hash.Add($NameAdInfo, $StateAdInfo)

#Location
$NameLocation = "Location"
$StateLocation = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}" -Name "SensorPermissionState" | select-object -ExpandProperty SensorPermissionState

if($StateLocation -eq 1){
    $StateLocation = "Le paramètre est activé"
    }
    else {
        $StateLocation = "Le paramètre est désactivé"
    }

$hash.Add($NameLocation, $StateLocation)

#Tracking
$NameDiagTrack = "DiagTrack"
$StateDiagTrack = Get-service DiagTrack | Select-Object -ExpandProperty Status
if($StateDiagTrack -eq "Running"){
    $StateDiagTrack = "Le paramètre est activé"
    }
    else {
        $StateDiagTrack = "Le paramètre est désactivé"
    }

$hash.Add($NameDiagTrack, $StateDiagTrack)

#Function to get status when privates parameters activated
function GetPropertyEnabled ($path, $property){
    Try {
        Get-ItemProperty -Path $path -Name $property -ErrorAction stop
    
    }
    Catch {
        $State = "Le paramètre est activé"
        return $State
    }
    $State = "Le paramètre est désactivé"
    return $State
}

#SmartScreen Filter
$NameSmartScrenn = "SmartScreen"
$StateSmartScreen = GetPropertyEnabled("HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppHost","EnableWebContentEvaluation")

$hash.Add($NameSmartScrenn, $StateSmartScreen)

#Feedback
$NameFeedback = "Feedback"
$StateFeedback = GetPropertyEnabled("HKCU:\Software\Microsoft\Siuf\Rules","NumberOfSIUFInPeriod")

$hash.Add($NameFeedback, $StateFeedback)

#Windows P2P
$NameP2P = "Windows Update P2P"
$StateP2P = GetPropertyEnabled("HKLM:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config", "DODownloadMode")

$hash.Add($NameP2P, $StateP2P)

#Cortana
$NameCortana = "Cortana"
$StateCortana = Get-ItemProperty -Path HKCU:\Software\Microsoft\Personalization\Settings | Select-Object -ExpandProperty AcceptedPrivacyPolicy
if($StateCortana -eq 1){
    $StateCortana = "Le paramètre est activé"
    }
    else {
        $StateCortana = "Le paramètre est désactivé"
    }

$hash.Add($NameCortana, $StateCortana)

#WifiSense
$NameWifiSense = "Wifi Sense"
$Version = [environment]::OSVersion.Version |Select-Object -ExpandProperty Build
if($Version -lt "1709"){
    $StateWifiSense = GetPropertyEnabled("HKLM:\Software\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting", "value")
    }else{
       $StateWifiSense = "Le paramètre est désactivé"
       }

$hash.Add($NameWifiSense, $StateWifiSense)

#Caméras
$cameras = Get-PnpDevice  -Class Camera | Select-Object Class, FriendlyName, Description, Status

#Microphones
$microphone = Get-PnpDevice -FriendlyName *microphone* -Class AudioEndpoint | Select-Object Class, FriendlyName, Description, Status

#Creating privacy parameters object
$PrivacyParametersHtml = [pscustomobject]$hash | ConvertTo-Html -Fragment -PreContent "<h2>Status des paramètres de la vie privée</h2>"
$camerasHtml = $cameras | ConvertTo-Html -Fragment -PreContent "<h2>Cameras</h2>"
$microphoneHtml = $microphone | ConvertTo-Html -Fragment -PreContent "<h2>Microphones</h2>"

############################# Display useless services #############################
#Thoses services where determinated according to the recommendations of the ANSSI, and assuming that this audit was running on human resources employee. 

#Get useless services
$UselessServices = Get-Service -Name "ScardSvr","PcaSvc","NetTcpPortSharing", "WerSvc","WdiServiceHost","SharedAccess","TapiSrv", "WMPNetworkSvc", "DPS", "SCPolicySvc", "diagnosticshub*", "DiagTrack", "dmwappush*", "lfsvc", "RetailDemo", "WbioSrvc", "Xbl*", "Xbox*", "MapsBroker", "TabletInputService" |Select Name, Status, DisplayName |Sort-Object Name

#Creating useless services objects
$UselessServicesHtml = $UselessServices | ConvertTo-Html -Fragment -PreContent "<h2>Services inutiles</h2>"

############################# Create HTML file #############################

$head = @"
	<title>Rapport d'audit </title>
	<style>
 
		body {
			background-color: #F7FAFC;
			font-family: monospace;
		}
		h1 {
			color: #1AB6E6;
		}
		h2 {
			color: #1AB6E6;
		}
		table {
			background-color: #D1D9DB;
            border-collapse: collapse;
		}
		td {
			border: 2px solid #EEEEEE;
			background-color: #D1D9DB;
			color: #1AB6E6;
			padding: 5px;
		}
		th {
			border: 2px solid #EEEEEE;
			background-color: #1AB6E6;
			color: #FFFFFF;
			text-align: left;
			padding: 5px;
		}
		div.localUsersData table {
			width: 100%;
		}
	</style>
"@

# Output to file
ConvertTo-Html -Head $head -Body "$computerHtml $adresssesHtml 
									<div class=`"localUsersData`">
										<h2>Local user informations</h2>
										$localUserDataHtml
									</div> 
									$PrivacyParametersHtml 
									$camerasHtml 
									$microphoneHtml 
									$UselessServicesHtml" | Out-File $reportPath


