#include <Date.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\..\..\..\..\..\..\..\..\Icones\Metro Icons\System Icons\OTHERS\OTHERS\ICO\2404.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.5
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include "..\au3\ficconf.au3"

Global $DiagDepart = 0
Global $LogFile = @ScriptDir & "\Logs\telephone_diag.log"
If Not FileExists(@ScriptDir & "\Logs") Then DirCreate(@ScriptDir & "\Logs")

ConsoleWrite(@error & @CRLF)
Global $step = 0
Global $GUI, $LabelQuestion, $BtnOui, $BtnNon, $EditLog, $BtnReset, $LabelAction, $Versionini
Global $questions[19] = [ _
		"Les appels entrants fonctionnent-ils ? Oui pour un soucis d'appels sortants", _ ;0
		"S'agit-il d'un problème uniquement au niveau du standard ?", _ ;1
		"Plusieurs téléphones sont-ils concernés ?", _ ;2
		"Avez-vous un softphone ?", _ ;3
		"Les téléphones redémarrent-ils ?", _ ;4
		"Les communications venant de l'extérieur sont-elles hachurées, coupées ?", _ ;5
		"Les téléphones HD sont-ils concernés ?", _ ;6
		"L'incident est-il résolu ?", _ ;7
		"Clôture de la demande", _ ;8
		"Remplacer le téléphone", _ ;9
		"Fin du diagnostic", _ ;10
		"Le probléme est résolu en changeant de téléphone", _ ;11
		"S'agit-il d'un problème uniquement au niveau du standard ?", _ ;12
		"Plusieurs téléphones sont-ils concernés ?", _ ;13
		"Les téléphones redémarrent-ils ?", _ ;14
		"Les communications venant de l'extérieur sont-elles hachurées, coupées ?", _ ;15
		"Avez-vous un softphone ?", _ ;16
		"Site non géré ?", _ ;17
		"Le téléphone est bien connecté au réseau?" _ ;18
		]

Global $questionPath = [] ; pour garder l'historique des décisions
FileInstall(".\telephone.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)
If Not FileExists($chemin_dossier_bin & "conf\OA.dll") Then
_FileCopy("O:\Techs\outiladmins\bin\conf\OA.dll", $chemin_dossier_bin & "conf\")
EndIf
If Not FileExists($chemin_dossier_bin & "conf\config.flq") Then
_FileCopy("O:\Techs\outiladmins\bin\conf\Config.Flq", $chemin_dossier_bin & "conf\")
endif
; Appel initial
ShowSelectionGUI()

Func ShowSelectionGUI()
	Local $guiSelect = GUICreate("Choix du site", 400, 150, -1, -1)
	Local $label = GUICtrlCreateLabel("Sélectionnez le site :", 30, 20, 200, 20)
	Local $combo = GUICtrlCreateCombo("", 30, 50, 340, 25)
	GUICtrlSetData($combo, $listeNas, $listeNas) ; suppose que $listeNas contient une liste séparée par "|"
	Local $btnOK = GUICtrlCreateButton("Valider", 150, 90, 100, 30)
	GUISetState()

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				Exit
			Case $btnOK
				Global $ConnNas = GUICtrlRead($combo)
				If $ConnNas <> "" And $ConnNas <> "Choix" Then
					GUIDelete($guiSelect)
					_MainGUI()
					Return
				Else
					MsgBox(16, "Erreur", "Veuillez sélectionner un site valide.")
				EndIf
		EndSwitch
	WEnd
EndFunc


Func _MainGUI()

	; paramétrage Général
	Opt("GUICloseOnESC", 0) ; 0=pas de sortie par la touche ESC 1= Fermeture du programme via esc
	;Opt("TrayAutoPause", 0) ;0=pas de pause, 1=pause
	Opt("TrayIconHide", 1) ;0=montre, 1=masque l'icône de Notification

	Global $version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Telephone", "Version", "REG_SZ", $version)

	#Region ### START Koda GUI section ### Form=
	Global $GUI = GUICreate("Dépannage Téléphones    v: " & $version, 623, 478, 849, 191)
	GUISetBkColor(0xFFFFFF)
	;GUICtrlCreateLabel("Diagnostic Incidents Téléphones", 100, 10, 300, 30)
	;Global $CB_NAS = GUICtrlCreateCombo("Choix", 130, 18, 185, 25)
	;GUICtrlSetData($CB_NAS, $listeNas)
	Global $MenuPrincipal = GUICtrlCreateMenu("&Schémas")
	Global $MenuScSJ = GUICtrlCreateMenuItem("Schéma sites agenais, CE, CMS Casteljaloux", $MenuPrincipal)
	Global $MenuScDistant = GUICtrlCreateMenuItem("Schéma CMS, MD, Parc, UD Ma et Vsl", $MenuPrincipal)

	Global $LabelQuestion = GUICtrlCreateLabel($questions[$step], 20, 60, 560, 40)
	GUICtrlSetFont($LabelQuestion, 10, 600, 0, "MS Sans Serif")
	Global $BtnOui = GUICtrlCreateButton("Oui", 100, 120, 100, 40)
	Global $BtnNon = GUICtrlCreateButton("Non", 250, 120, 100, 40)
	Global $BtnReset = GUICtrlCreateButton("Réinitialiser", 400, 120, 100, 40)
	Global $EditLog = GUICtrlCreateEdit("", 20, 214, 560, 180)
	GUICtrlSetData(-1, "")
	GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
	Global $LabelAction = GUICtrlCreateLabel("", 20, 176, 560, 30)
	GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
	GUICtrlSetColor(-1, 0xCC0000)
	Global $quitButton = GUICtrlCreateIcon($dlloa, 4, 512, 416, 32, 32)
	GUICtrlSetTip($quitButton, "Quitter")
	Global $Pic1 = GUICtrlCreateIcon($dlloa, 1, 64, 408, 200, 38)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	; Détermination du point de départ du diagnostic

				cryptage($configini, 2, $Key)
				_FileWriteLog($Cheminlog, "decryptage de  " & $configini & " code erreur " & @error)
				$DiagDepart = Number(IniRead($configini, $ConnNas, "diagdepart", "0"))
				cryptage($configini, 1, $Key)


	$step = $DiagDepart
	GUICtrlSetData($LabelQuestion, $questions[$step])

	_log("Début du diagnostic")

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $quitButton
				FileDelete($chemin_dossier_bin & "telephone.au3")
				Exit
			case $MenuScDistant
				FileInstall("..\..\Documentations\Schema telephonie distant.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\Schema telephonie distant.pdf", "", "", "")
			case $MenuScSJ
				FileInstall("..\..\Documentations\Schema telephonie agenais.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\Schema telephonie agenais.pdf", "", "", "")
			Case $BtnOui
				ProcessResponse(True)
			Case $BtnNon
				ProcessResponse(False)
			Case $BtnReset
				ResetDiag()
		EndSwitch
	WEnd
EndFunc   ;==>_MainGUI

Func ProcessResponse($yes)
	Local $msg = $questions[$step]
	Local $reponse = $yes ? "Oui" : "Non"
	_log("Q: " & $msg)
	_log("R: " & $reponse)

	; Logique du chemin en fonction de l'étape actuelle et de la réponse
	Switch $step
		Case 0 ; Appels entrants ?
			If Not $yes Then
				$step = 6 ; Téléphones HD concernés ?
			Else
				$step = 1
			EndIf
		Case 1 ; Problème standard ?
			If $yes Then
				$step = 11
			Else
				$step = 2
			EndIf
		Case 2 ; Tous les téléphones concernés ?
			If Not $yes Then
				$step = 3 ; Softphone ?
			Else
				$step = 4
			EndIf
		Case 3 ; Softphone ?
			If $yes Then
				ShowAction(" : Test Ping + Redémarrage PC + Vérification connectique")
			Else
				$step = 18 ; Téléphone connecté au réseau
			EndIf

		Case 4 ; Les téléphones redémarrent-ils ?
			If $yes Then
				ShowAction(" : Transférer la demande aux Admins")
				;$step = 5 ; Comms hachurées ?
			Else
				$step = 5 ; Comms hachurées ?
			EndIf
		Case 5 ; Comms hachurées ?
			If $yes Then
				ShowAction(" : Ticket Orange pour vérifier T0 du site"); + Ouvrir un ticket NXO")
			Else
				ShowAction(" : Transférer la demande aux Admins")
				;$step = 7 ; Incident résolu ?
			EndIf

		Case 6 ; Téléphones HD concernés ?
			If $yes Then
				ShowAction(" : Ticket SFR pour vérifier T2 HD")
				EndDiag()
			Else
				ShowAction(" : Transférer aux Admins avec n° ticket NXO")
				EndDiag()
				;$step = 2
			EndIf
		Case 7 ; Incident résolu ?
			If $yes Then
				$step = 8
			Else
				ShowAction(" : Transférer aux Admins avec n° ticket NXO")
				EndDiag()
			EndIf
		Case 8 ; Clôture
			_log("Clôture de la demande")
			EndDiag()
		Case 11 ; Changement de téléphone
			If $yes Then
				EndDiag()
			Else
				ShowAction(" : Ouvrir un ticket NXO")
				EndDiag()
			EndIf
		Case 12 ; Problème standard ?
			If $yes Then
				$step = 11
			Else
				$step = 13
			EndIf
		Case 13 ; Tous les téléphones concernés ?
			If Not $yes Then
				$step = 16 ; reboot ?
			Else
				$step = 14 ; Softphone
			EndIf
		Case 14 ; Les téléphones redémarrent-ils ?
			If $yes Then
				ShowAction(" : Transférer la demande aux Admins")
				EndDiag()

			Else
				$step = 15 ; Comms hachurées ?
			EndIf
		Case 15 ; Comms hachurées ?
			If $yes Then
				ShowAction(" : Ticket SFR pour vérifier T2 de l'HD"); + Ouvrir un ticket NXO")

			Else
				ShowAction(" : Test Ping sur routeur BVPN + Ouvrir un ticket chez Orange")

			EndIf
			EndDiag()
		Case 16 ; Softphone ?
			If $yes Then
				ShowAction(" : Test Ping + Redémarrage PC + Vérification connectique")
			Else
				$step = 18 ; Téléphone connecté au réseau
			EndIf

		Case 17 ; Non géré
			ShowAction(" : Site non géré actuellement. Contacter le prestataire")
			EndDiag()
		Case 18 ; Téléphone connecté au réseau
			If $yes Then
				ShowAction(" : Remplacer le téléphone")

			Else
				ShowAction(" :  Vérification connectique + test Ping")

			EndIf
EndDiag()
	EndSwitch

	If $step < UBound($questions) Then
		GUICtrlSetData($LabelQuestion, $questions[$step])
	Else
		GUICtrlSetData($LabelQuestion, "Fin du diagnostic")
		GUICtrlSetState($BtnOui, $GUI_DISABLE)
		GUICtrlSetState($BtnNon, $GUI_DISABLE)
	EndIf
EndFunc   ;==>ProcessResponse

Func _log($msg)
	Local $timestamp = _NowCalc()
	FileWriteLine($LogFile, "[" & $timestamp & "] " & $msg)
	GUICtrlSetData($EditLog, GUICtrlRead($EditLog) & "[" & $timestamp & "] " & $msg & @CRLF)
EndFunc   ;==>_log

Func EndDiag()
	$step = 10
	GUICtrlSetData($LabelQuestion, $questions[$step])
	GUICtrlSetState($BtnOui, $GUI_DISABLE)
	GUICtrlSetState($BtnNon, $GUI_DISABLE)
EndFunc   ;==>EndDiag

Func ResetDiag()
	GUIDelete($gui)
	ShowSelectionGUI()
	#cs
	$step = 0
	GUICtrlSetData($LabelQuestion, $questions[$step])
	GUICtrlSetData($LabelAction, "")
	GUICtrlSetData($EditLog, "")
	GUICtrlSetState($BtnOui, $GUI_ENABLE)
	GUICtrlSetState($BtnNon, $GUI_ENABLE)
	#ce
	Log("Diagnostic réinitialisé")
EndFunc   ;==>ResetDiag

Func ShowAction($action)
	GUICtrlSetData($LabelAction, "=> " & $action)
	Log("Action : " & $action)
EndFunc   ;==>ShowAction

Func _FileCopy($pFrom, $pTo, $fFlags = 0)
    Local $FOF_NOCONFIRMMKDIR = 512
    Local $FOF_NOCONFIRMATION = 16
    If Not FileExists($pTo) Then DirCreate($pTo)
    $winShell = ObjCreate("shell.application")
    ;$winShell.namespace($pTo).CopyHere($pFrom, BitOR(272))
	$winShell.namespace($pTo).CopyHere($pFrom, BitOR($FOF_NOCONFIRMMKDIR,$FOF_NOCONFIRMATION))
EndFunc
