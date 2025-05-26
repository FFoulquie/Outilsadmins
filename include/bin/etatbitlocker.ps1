


$bitlockerStatus =  {Get-BitLockerVolume |Select-Object -ExpandProperty ProtectionStatus -First 1}

$bitlockerStatus 