if((Get-BitLockerVolume -MountPoint $env:SystemDrive).ProtectionStatus -eq "Off"){

echo *********
echo Decrypte
echo *********

}
if((Get-BitLockerVolume -MountPoint $env:SystemDrive).ProtectionStatus -eq "On"){

echo     **********
echo      crypte 
echo     **********

}