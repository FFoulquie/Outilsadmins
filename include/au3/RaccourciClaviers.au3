#include <File.au3>
#include <FileConstants.au3>

global $sIPAddress

Func startssi()
	Run("C:\Progwindev\ssi\ssi.exe")

	WinWaitActive("[Title:Systèmes d'information]")
	WinActivate("[Title:Systèmes d'information]")
EndFunc   ;==>startssi

Func verifIP($IPPc)

	adresseIP($IPPc)

	Local $iPing = Ping($IPPc, 250)

	If $iPing Then
		$labip = GUICtrlCreateLabel("IP : " & $sIPAddress, 16, 96, 150, 17)
		GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		GUICtrlSetColor(-1, 0x33FF33)
	Else
		$labnip = GUICtrlCreateLabel("Poste non joignable", 16, 96, 150, 17)
		GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		GUICtrlSetColor(-1, 0xFF0000)
	EndIf
EndFunc   ;==>verifIP


Func ssiM($IPPc)
	ControlClick("[Title:Systèmes d'information]", "", "[CLASS:Button; INSTANCE:3]")
	WinActivate("[Title:Gestion des matériels]]")
	MouseClick("left", 741, 63)
	Send($IPPc)
	Send("{ENTER}")
EndFunc   ;==>ssiM

Func ssiO($IPPc)
	;Clique sur le bouton ordinateur
	ControlClick("[Title:Systèmes d'information]", "", "[CLASS:Button; INSTANCE:5]")
	WinWaitActive("[Title:Gestion des stations]")
	;Clique dans la fenêtre N°Ordinateur
	MouseClick("left", 148, 39)
	Send($IPPc)
	Send("{ENTER}")
	;Double clic sur la première ligne du tableau de résultat
	MouseClick("left", 202, 262, 2)
EndFunc   ;==>ssiO


Func _GestionUser()
	cryptage($Versionini, 2, $Key)
	If IniRead($Versionini, "Mdp", "version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Msp", "Version") Then
		FileCopy($Sources & "mdp.exe", $chemin_dossier_bin & "", $FC_OVERWRITE)
	EndIf
	cryptage($Versionini, 1, $Key)
	_FileWriteLog($Cheminlog, "Lancement de l'outil de gestion des utilisateurs MDP.exe")
	RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "mdp.exe ")
EndFunc   ;==>_GestionUser

Func _GestADPC()
	cryptage($Versionini, 2, $Key)
	If IniRead($Versionini, "GestPcAD", "version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\GestPcAD", "Version") Then
		FileCopy($Sources & "GestPcAd.exe", $chemin_dossier_bin & "", $FC_OVERWRITE)
	EndIf
	cryptage($Versionini, 1, $Key)
	_FileWriteLog($Cheminlog, "Lancement de la gestion PC AD")
	RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "GestPcAd.exe")
EndFunc   ;==>_GestADPC


Func _Cyber()
	_FileWriteLog($Cheminlog, "Lancement de Cyber 47")
	cryptage($Versionini, 2, $Key)
	If IniRead($Versionini, "Cyber", "version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Cyber", "Version") Then
		FileCopy($Sources & "Desactivation.exe", $chemin_dossier_bin & "", $FC_OVERWRITE)
	EndIf
	cryptage($Versionini, 1, $Key)
	_FileWriteLog($Cheminlog, "Lancement de l'outil Cyber47")
	RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "Desactivation.exe ")
EndFunc   ;==>_Cyber

Func _applis($IPPc)
	cryptage($Versionini, 2, $Key)
	If IniRead($Versionini, "infoappli", "version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\infoappli", "Version") Then
		FileCopy($Sources & "ApplisInstall.exe", $chemin_dossier_bin & "", $FC_OVERWRITE)
	EndIf
	cryptage($Versionini, 1, $Key)
	_FileWriteLog($Cheminlog, "Lancement d'infos applis sur " & $IPPc)
	RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "ApplisInstall.exe ")
EndFunc   ;==>_applis

Func _infopc($IPPc)
	_FileWriteLog($Cheminlog, "Lancement d'info PC sur : " & $IPPc)
	cryptage($Versionini, 2, $Key)
	If IniRead($Versionini, "infoPC", "version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\infoPC", "Version") Then
		FileCopy($Sources & "infoPC.exe", $chemin_dossier_bin & "", $FC_OVERWRITE)
	EndIf
	cryptage($Versionini, 1, $Key)
	Run($chemin_dossier_bin & "infoPC.exe")
EndFunc   ;==>_infopc


Func adresseIP($IPPc)
	; Démarre le service TCP.
	TCPStartup()

	; Inscrit OnAutoItExit qui sera appelée lorsque le script se fermera.
	OnAutoItExitRegister("OnAutoItExit")

	; Assigne une variable locale avec l'adresse IP:
	Global $sIPAddress = TCPNameToIP($IPPc)

	; Si une erreur s'est produite, affiche le code d'erreur et retourne False.
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "", "Code de l'erreur: " & @error)
		Return False
	EndIf
EndFunc   ;==>adresseIP


