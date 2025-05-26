param (
    [string]$ComputerName
)


$objComputer = Get-ADComputer $ComputerName
$Bitlocker_Object = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $objComputer.DistinguishedName -Properties 'msFVE-RecoveryPassword'
echo $Bitlocker_Object 

Write-Host "Appuyer sur entree pour continuer"