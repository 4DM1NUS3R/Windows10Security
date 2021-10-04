#Get computerName
Write-host 'ComputerName :' $env:computername

#Get Status 
if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $env:computername -Quiet) {
        $Status = "Online"
    } else {
        $Status = "Offline"
    }
Write-Host 'Status :' $Status

#Get OS Informations
$OSInfo = Get-WmiObject -ComputerName $env:computername -ClassName Win32_OperatingSystem | Select-Object Caption, Version, OSArchitecture

#Get OS Version
Write-Host 'OS version :' $OSInfo.Version

#Get OS Caption
Write-Host 'OSCaption :' $OSInfo.Caption

#Get OS Architecture
Write-Host 'OS Architecture :' $OSInfo.OSArchitecture

#Get VM
$InfoModel = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $env:computername | Select Manufacturer, Model
if  ($InfoModel){
        $resultvm= 'False'
    } else {
        $resultvm = 'True'
    }
Write-Host 'VM :' $resultvm

#Get Model
Write-host 'Model :' $InfoModel.Model

#Get Manufacturer
Write-Host 'Manufacturer :' $InfoModel.Manufacturer

#Get DateBuilt
$DateBuilt =([WMI]'').ConvertToDateTime((Get-WmiObject -ComputerName $env:computername -ClassName Win32_OperatingSystem).InstallDate)
Write-host 'DateBuilt:' $DateBuilt

#Get LastBootTime
$LastBootTime = Get-CimInstance -ClassName Win32_OperatingSystem | Select lastbootuptime
Write-Host 'LastBootTime :' $LastBootTime.lastbootuptime

#Get IP Adress
Get-WmiObject win32_networkadapterconfiguration | 
Select-Object -Property @{
    Name = 'IPAddress'
    Expression = {($PSItem.IPAddress[0])}
},MacAddress | 
Where IPAddress -NE $null