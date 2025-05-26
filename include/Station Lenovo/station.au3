#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\DOCKMGR.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


#cs
********************************************************************
V1 26/09/2023
Affiche la version du Firmware et le numéro de série des stations d'accueil LENOVO USB-C

#Ce

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Station", "Version", "REG_SZ", $version)

$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems = ""
$strComputer = "localhost"

$OutputTitle = ""
$Output = ""
$OutputTitle &= "Computer: " & $strComputer  & @CRLF
$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\Lenovo\Dock_Manager")
$colItems = $objWMIService.ExecQuery("SELECT * FROM DockDevice", "WQL", _
                                          $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

If IsObj($colItems) then
   Local $Object_Flag = 0
   For $objItem In $colItems
      $Object_Flag = 1
	  $Output &= "   "
      $Output &= "FWVersion: " & $objItem.FWVersion & @CRLF
      $Output &= "SerialNumber: " & StringRight($objItem.SerialNumber,8) & @CRLF
	   $Output &= "MACAddress: " & $objItem.MACAddress & @CRLF
   Next
  If $Object_Flag = 0 Then Msgbox(1,"Sortie WMI",$OutputTitle)
   ConsoleWrite($Output)
  ; FileWrite("c:\temp\DockDevice.TXT", $Output )
  ; Run(@Comspec & " /c start " & @TempDir & "\DockDevice.TXT" )
Else
   Msgbox(0,"Sortie WMI","Aucun objets WMI trouvés pour la Classe :" & "DockDevice" )
Endif


