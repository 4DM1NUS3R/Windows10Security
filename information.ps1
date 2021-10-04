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
$adresssesHtml = $adresses | ConvertTo-Html -Fragment -PreContent "<h2>IP and MAC Adresses</h2>"

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
$computerHtml = $computer | ConvertTo-Html -Fragment -PreContent "<h1>Computer Report </h1><h2>General Informations</h2>"


# Create HTML file
$head = @"
	<title>Computer Report</title>
	<style>
		body {
			background-color: #282A36;
			font-family: sans-serif;
		}
		h1 {
			color: #FF7575;
		}
		h2 {
			color: #E56969;
		}
		table {
			background-color: #363949;
            border-collapse: collapse;
		}
		td {
			border: 2px solid #282A36;
			background-color: #363949;
			color: #FF7575;
			padding: 5px;
		}
		th {
			border: 2px solid #282A36;
			background-color: #363949;
			color: #FF7575;
			text-align: left;
			padding: 5px;
		}
	</style>
"@
# Output to file
ConvertTo-Html -Head $head -Body "$computerHtml $adresssesHtml" | Out-File $reportPath


