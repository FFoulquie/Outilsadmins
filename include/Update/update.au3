#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\..\include\DLL\Icones\System Restore.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.3
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <APIShellExConstants.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>


#include "..\au3\ficconf.au3"
#include "..\au3\fonctions.au3"

FileInstall(".\Update.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)
$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Update", "Version", "REG_SZ", $version)

#Region ### START Koda GUI section ### Form=
Global $Update = GUICreate("Update  version : " & $version, 478, 268, 784, 354)
GUISetFont(10, 400, 0, "MV Boli")
GUISetBkColor(0xFFFFFF)
Global $Label1 = GUICtrlCreateLabel("Mise à jour d'une application", 88, 24, 303, 31)
GUICtrlSetFont(-1, 16, 400, 0, "Elephant")
GUICtrlSetColor(-1, 0x0066CC)
Global $Pic1 = GUICtrlCreateIcon($dlloa, 9, 178, 92, 60, 60)
Global $bt_transfert = GUICtrlCreateButton("Transfert exe", 48, 200, 105, 33)
Global $BT_MAJini = GUICtrlCreateButton("Mise à jour INI", 224, 200, 105, 33)
Global $quitButton = GUICtrlCreateIcon($dlloa, 4, 400, 200, 32, 32)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $quitButton
			FileDelete($chemin_dossier_bin & "update.au3")
			Exit
		Case $bt_transfert
			Global $SourceFile = FileOpenDialog("Selectionner le fichier à copier", @ScriptDir & "\", "All files (*.*)", $FD_FILEMUSTEXIST)
			If @error Then
				MsgBox($MB_SYSTEMMODAL, "", "Aucun fichier sélectionné.")
			Else

				Global $DestFile = FileSelectFolder("Choisir le dossier de Destination", "\\srv-mdt2016\gpo$\")
				If @error Then
					MsgBox($MB_SYSTEMMODAL, "", "Pas de dossier sélectionné.")
					FileChangeDir(@ScriptDir)
				Else

					_FileCopy($SourceFile, $DestFile)
				EndIf
			EndIf
		Case $BT_MAJini
			ShellExecute("\\srv-mdt2016\gpo$\AutoUpdate\version.ini")
	EndSwitch
WEnd
