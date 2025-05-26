#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\DLL\Nouveau dossier\btl.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.8
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include "..\au3\ficconf.au3"

FileInstall(".\bitlocker.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Bitlocker", "Version", "REG_SZ", $version)

$ippc = IniRead($chemin_dossier_bin & "conf\param.flq", "PC", "Nom", "")
$ippc = InputBox("Nom de l'ordinateur", "saisir le nom de l'ordinateur a analyser", $ippc)



cryptage($userini, 2, $Key)
$sUsername = IniRead($userini, "Credential", "Nom", "")
$sPassword = IniRead($userini, "Credential", "Pass", "")
$Domain = IniRead($userini, "Credential", "Domaine", "")
cryptage($userini, 1, $Key)

#Region ### START Koda GUI section ### Form=
Global $Form1 = GUICreate("Gestion BitLocker    version : " & $version, 525, 437, 320, 190)
GUISetBkColor(0xFFFFFF)

Global $Bt_CheckAD = GUICtrlCreateButton("Vérification sur l'AD", 24, 150, 105, 49, $BS_MULTILINE)
GUICtrlSetTip($Bt_CheckAD, "Controle l'existance de la sauvegarde de la clé de chiffrement sur l'AD")
Global $Bt_checkCrypt = GUICtrlCreateButton("1 - Vérification cryptage", 24, 208, 105, 49, $BS_MULTILINE)
GUICtrlSetTip($Bt_checkCrypt, "Vérifie si le cryptage est actif sur le PC")
Global $Bt_CheckTpm = GUICtrlCreateButton("2 - Vérification TPM + Mot de passe", 168, 208, 105, 49, $BS_MULTILINE)
GUICtrlSetTip($Bt_CheckTpm, "Vérifie si le TPM et la clé de chiffrement sont actifs sur le PC")
Global $Bt_SavAD = GUICtrlCreateButton("3 - Sauvegarde sur l'AD", 328, 208, 105, 49, $BS_MULTILINE)
GUICtrlSetTip($Bt_SavAD, "Sauvegarde la clé de cryptage sur l'AD")

Global $Icon1 = GUICtrlCreateIcon($dlloa, 7, 240, 56, 65, 65)
Global $Label1 = GUICtrlCreateLabel($ippc, 232, 144, 134, 24)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetColor(-1, 0xFF0000)
Global $Label2 = GUICtrlCreateLabel("Gestion BitLocker", 176, 8, 194, 33)
GUICtrlSetFont(-1, 18, 400, 0, "MS Sans Serif")
Global $Bt_Doc = GUICtrlCreateButton("Documentation", 24, 336, 89, 33)
$BT_QUIT = GUICtrlCreateIcon($dlloa, 4, 432, 336, 33, 33)
GUICtrlSetTip($BT_QUIT, "Quitter")

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $Bt_Doc
			FileInstall("bitlocker.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
			ShellExecute($chemin_dossier_bin & "doc\bitlocker.pdf", "", "", "")
		Case $GUI_EVENT_CLOSE, $BT_QUIT
			FileDelete($chemin_dossier_bin & "bitlocker.au3")
			Exit

		Case $Bt_CheckAD
						Local $iPing = Ping($ippc, 250)
			If $iPing Then
			FileInstall("CheckBitlocker.ps1", $chemin_dossier_bin & "", $FC_OVERWRITE)
			RunAs($sUsername, $Domain, $sPassword, 0, "cmd")
			_FileWriteLog($Cheminlog, "****Bitlocker****Controle clé sur l'AD du PC : " & $ippc)
			Sleep(1000)
			Send("powershell -ExecutionPolicy ByPass -File " & $chemin_dossier_bin & "CheckBitlocker.ps1 -ComputerName " & $ippc & "{Enter}")
			Else
        MsgBox(48, "Erreur", "Le PC " & $ippc & " n'est pas joignable.")
    EndIf
		Case $Bt_checkCrypt
			Local $iPing = Ping($ippc, 250)
			If $iPing Then
				FileInstall("etatbitlocker.ps1", $chemin_dossier_bin & "", $FC_OVERWRITE)
				_FileWriteLog($Cheminlog, "****Bitlocker**** Controle etat sur " & $ippc)
				RunAsWait($sUsername, $Domain, $sPassword, 0, "xcopy " & $chemin_dossier_bin & "etatbitlocker.ps1 \\" & $ippc & "\c$\temp\ /Y")
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $ippc & " -s cmd")
				Sleep(1000)
				WinWaitActive("\\" & $ippc & ": cmd", "")
				Sleep(200)
				Send("powershell -ExecutionPolicy ByPass -File c:\temp\etatbitlocker.ps1 {Enter}")
				Else
				MsgBox(48, "Poste " & $ippc & " non joignable", "Le PC n'est pas joignable." & @CRLF & "Action interrompue")
			EndIf

		Case $Bt_CheckTpm
			Local $iPing = Ping($ippc, 250)
			If $iPing Then
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $ippc & " -s cmd")
				WinWaitActive("\\" & $ippc & ": cmd", "")
				_FileWriteLog($Cheminlog, "****Bitlocker****Controle clé sur " & $ippc)
				Sleep(200)
				Send("manage-bde -protectors -get c: {Enter}")
			Else
				MsgBox(48, "Poste " & $ippc & " non joignable", "Le PC n'est pas joignable." & @CRLF & "Action interrompue")
			EndIf
		Case $Bt_SavAD
			Local $iPing = Ping($ippc, 250)
			If $iPing Then
				FileInstall("savcle.ps1", $chemin_dossier_bin & "", $FC_OVERWRITE)
				Sleep(1000)
				_FileWriteLog($Cheminlog, "****Bitlocker**** sauvegarde clé pour " & $ippc)
				RunAsWait($sUsername, $Domain, $sPassword, 0, "xcopy " & $chemin_dossier_bin & "savcle.ps1 \\" & $ippc & "\c$\temp\ /Y")
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $ippc & " -s cmd")
				WinWaitActive("\\" & $ippc & ": cmd", "")
				Sleep(200)
				Send("powershell -ExecutionPolicy ByPass -File c:\temp\savcle.ps1 {Enter}")
			Else
				MsgBox(48, "Poste " & $ippc & " non joignable", "Le PC n'est pas joignable." & @CRLF & "Action interrompue")
			EndIf
	EndSwitch
WEnd
