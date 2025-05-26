param (
    [string]$ComputerName
)


$objComputer = Get-ADComputer $ComputerName
$Bitlocker_Object = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $objComputer.DistinguishedName -Properties 'msFVE-RecoveryPassword'
$Bitlocker_Object 