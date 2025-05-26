$ADBegup = Get-BitLockerVolume -MountPoint "C:"
Backup-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $ADBegup.KeyProtector[1].KeyProtectorId