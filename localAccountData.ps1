Write-Host "Local account data : `n"
foreach($user in Get-LocalUser){
    $user | Select-Object Name,FullName,SID,LastLogon,PasswordExpire,PasswordLastSet,PasswordChangeableDate,AccountExpires,Enabled | Format-List
    Write-Host "File permissions :`n"
    .\accesschk.exe $user c:\* -s
}