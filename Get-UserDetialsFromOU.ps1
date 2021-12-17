$TimeStamp = Get-Date -Format "yyyyMMddTHHmmssffff"
$CSVOutputPath = "$PSScriptRoot\Output\Get-UserDetialsFromOU-$TimeStamp.csv"

$OU = Read-Host "Enter the OU path"

Get-ADUser -Filter "Enabled -eq '$true'" -Properties * -SearchBase $OU `
    | Select-Object Name, EmailAddress, UserPrincipalName, Department, Title, Manager, Office, Country `
    | Sort-Object Country, Office, Department, Title, Name `
    | Export-Csv -Path $CSVOutputPath -NoTypeInformation

Invoke-Item $CSVOutputPath
