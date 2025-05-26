#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Fileversion=1.0.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include "..\au3\ficconf.au3"

; paramétrage Général
Opt("GUICloseOnESC", 0) ; 0=pas de sortie par la touche ESC 1= Fermeture du programme via esc
;Opt("TrayAutoPause", 0) ;0=pas de pause, 1=pause
Opt("TrayIconHide", 1) ;0=montre, 1=masque l'icône de Notification

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\InfoLenovo", "Version", "REG_SZ", $version)

FileInstall(".\InfoLenovo.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)

$ippc = IniRead($chemin_dossier_bin & "conf\param.flq", "PC", "Nom", "")


#Region ### START Koda GUI section ### Form=C:\CD47Users\frfoulqu\Documents\CD47\Scripts\Autoit\Outils Admin\include\Station Lenovo\infostation.kxf
Global $InfoStation = GUICreate("Info Station Lenovo     v:" & $version, 362, 346, 763, 276)
GUISetBkColor(0xFFFFFF)
Global $Label1 = GUICtrlCreateLabel($ippc, 24, 8, 36, 17)
Global $Bt_Info = GUICtrlCreateButton("Info Version Station", 48, 144, 105, 33)
Global $BT_DscDis = GUICtrlCreateButton("Pb scintillement", 200, 144, 105, 33)
Global $ico_lenovo = GUICtrlCreateIcon($chemin_dossier_bin & "conf\OA.dll", 7, 80, 64, 199, 65)
Global $BCancel = GUICtrlCreateIcon($chemin_dossier_bin & "conf\OA.dll", 4, 256, 264, 32, 32)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###





While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $BCancel
			FileDelete("c:\temp\dsc.exe")
			FileDelete($chemin_dossier_bin & "InfoLenovo.au3")
			Exit
		Case $BT_DscDis
							RunAs($sUsername, $Domain, $sPassword, 0, "xcopy " & $chemin_dossier_bin & "dsc.exe \\" & $IPPc & "\c$\temp\ /Y")
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
				WinWaitActive("\\" & $IPPc & ": cmd", "")
				Sleep(200)
				Send("cd \temp {Enter}")
				Sleep(200)
				Send("dsc.exe q {Enter}")
				;Send("station.exe {Enter}")
				Sleep(200)
				Send("{Enter}")
				Sleep(2000)
		Case $Bt_Info
			cryptage($Versionini, 2, $Key)
			If IniRead($Versionini, "Station", "version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Station", "Version") Then
				FileCopy($Sources & "station.exe", $chemin_dossier_bin & "", $FC_OVERWRITE)
			EndIf
			cryptage($Versionini, 1, $Key)
			_FileWriteLog($Cheminlog, "Recherche des informations sation LENOVO sur " & $ippc)
			Run("xcopy " & $chemin_dossier_bin & "station.exe \\" & $ippc & "\c$\temp\ /Y")
			ConsoleWrite( "copie sation LENOVO sur " & @CRLF)
			Run($chemin_dossier_bin & "PsExec.exe /accepteula \\" & $ippc & " -s cmd")
			ConsoleWrite( "psexec " & @error & " "  & @CRLF)
			WinWaitActive("\\" & $ippc & ": cmd", "")
			ConsoleWrite( "wait " & @CRLF)
			Sleep(200)
			Send("cd \temp {Enter}")
			ConsoleWrite( "send temp " & @CRLF)
			Sleep(200)
			Send("station.exe {Enter}")
			Sleep(200)
			Send("{Enter}")
			Sleep(2000)
	EndSwitch
WEnd

