#include <FileConstants.au3>
#include <MsgBoxConstants.au3>



#cs
$opale = regread("HKEY_LOCAL_MACHINE\SOFTWARE\CD47","opale")


$logopale = "C:\ProgramData\CD47\Logs\opale.txt"
FileOpen($logopale, 10)
FileWriteLine($logopale, $opale)
FileClose($logopale)
#ce

$logopale = "\\p000-t4181\c$\ProgramData\CD47\Logs\opale.txt"

Local $hFileOpen = FileOpen($logopale, $FO_READ)
 Local $sFileRead = FileReadLine($hFileOpen, 1)
  MsgBox($MB_SYSTEMMODAL, "", "Première ligne du fichier:" & @CRLF & $sFileRead)
 FileClose($hFileOpen)