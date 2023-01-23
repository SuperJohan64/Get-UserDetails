$TimeStamp = Get-Date -Format "yyyyMMddTHHmmssffff"
$CSVOutputPath = "$PSScriptRoot\Get-UserDetailsFromCsvName-$TimeStamp.csv"

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $PSScriptRoot
$OpenFileDialog.filter = "CSV (*.csv)| *.csv"
$OpenFileDialog.ShowDialog() | Out-Null
$OpenFileDialog.filename
$CSVImportPath = $OpenFileDialog.filename
$CSVImport = Import-Csv -Path $CSVImportPath

$CSVOutput = @()
foreach ($User in $CSVImport) {
    $UserName = $User.Name

    Try {
        $ADQuery = Get-ADUser -Filter 'Name -eq $UserName' -Properties EmailAddress, UserPrincipalName, Department, Title, Manager, Office, Country -ErrorAction Stop
        $ADManager = Get-ADUser $ADQuery.Manager
    
        $CSVOutput += New-Object -TypeName psobject -Property @{
            Name = $User.Name
            EmailAddress = $ADQuery.EmailAddress 
            UserPrincipalName = $ADQuery.UserPrincipalName
            Department = $ADQuery.Department
            Title = $ADQuery.Title
            Manager = $ADManager.Name
            Office = $ADQuery.Office
            Country = $ADQuery.Country
        }
    }
    Catch {
        $CSVOutput += New-Object -TypeName psobject -Property @{Name = $User.Name}
    }
}

$CSVOutput `
    | Select-Object Name, EmailAddress, UserPrincipalName, Department, Title, Manager, Office, Country `
    | Sort-Object Country, Office, Department, Title, Name `
    | Export-Csv -Path $CSVOutputPath -NoTypeInformation

Invoke-Item $CSVOutputPath
