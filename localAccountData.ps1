Write-Host "Local account data : `n"
$localUserData = @()
foreach($user in Get-LocalUser){
    $localUserData += $user | Select-Object Name,FullName,SID,LastLogon,PasswordExpire,PasswordLastSet,PasswordChangeableDate,AccountExpires,Enabled
    #Write-Host "File permissions :`n"
    #.\accesschk.exe $user c:\* -s
}

$localUserData | ForEach-Object {$PSItem}
#$localUserData[0] | Get-Member
