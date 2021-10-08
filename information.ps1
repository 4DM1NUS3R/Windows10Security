#Report path
$reportPath = $PSScriptRoot + "\report.html"

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

#Get useless services
$UselessServices = Get-Service -Name "ScardSvr","PcaSvc","NetTcpPortSharing", "WerSvc","WdiServiceHost","RasAuto","SharedAccess","TapiSrv", "WMPNetworkSvc", "DPS", "SCPolicySvc","seclogon", "diagnosticshub*", "DiagTrack", "dmwappush*", "lfsvc", "RetailDemo", "WbioSrvc", "Xbl*", "Xbox*", "MapsBroker", "TabletInputService" |Select Name, Status, DisplayName |Sort-Object Name

#Creating useless services objects
$UselessServicesHtml = $UselessServices | ConvertTo-Html -Fragment -PreContent "<h2>Services inutiles</h2>"

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


# Create HTML file
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
	</style>
"@
# Output to file
ConvertTo-Html -Head $head -Body "$computerHtml $adresssesHtml $UselessServicesHtml" | Out-File $reportPath


