#include <EditConstants.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\7532.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.6
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include "..\au3\ficconf.au3"

Opt("GUIOnEventMode", 1)

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Crypt", "Version", "REG_SZ", $version)

Global $sFilePath = ""
Global $sKey = ""
Global $sOperation = ""

; Création de la fenêtre principale
Global $hGUI = GUICreate("Cryptage / Décryptage de fichier", 500, 200)
GUISetOnEvent($GUI_EVENT_CLOSE, "OnClose")

; Label fichier
GUICtrlCreateLabel("Fichier :", 20, 20, 60, 20)
Global $inputFile = GUICtrlCreateInput("", 100, 20, 250, 20)
Global $btnBrowse = GUICtrlCreateButton("Parcourir", 360, 20, 80, 20)
GUICtrlSetOnEvent($btnBrowse, "BrowseFile")

; Label clé de chiffrement
GUICtrlCreateLabel("Clé de chiffrement :", 20, 60, 120, 20)
Global $inputKey = GUICtrlCreateInput("", 150, 60, 250, 20,$ES_PASSWORD)

; Radio boutons pour choisir l'opération
GUICtrlCreateLabel("Opération :", 20, 100, 60, 20)
Global $radioCrypt = GUICtrlCreateRadio("Crypter", 100, 100, 80, 20)
Global $radioDecrypt = GUICtrlCreateRadio("Décrypter", 200, 100, 80, 20)
GUICtrlSetState($radioCrypt, $GUI_CHECKED)
GUICtrlSetOnEvent($radioCrypt, "SetOperation")
GUICtrlSetOnEvent($radioDecrypt, "SetOperation")

; Bouton d'exécution
Global $btnExecute = GUICtrlCreateButton("Exécuter", 150, 140, 100, 30)
GUICtrlSetOnEvent($btnExecute, "ExecuteOperation")

Global $Label10 = GUICtrlCreateLabel($version, 400, 160, 50, 17)
GUISetState(@SW_SHOW, $hGUI)

;Global $btnQuit = GUICtrlCreateButton("Quitter", 270, 140, 100, 30)
Global $btnQuit = GUICtrlCreateIcon($chemin_dossier_bin & "conf\OA.dll", 4, 270, 140, 32, 32)
GUICtrlSetTip($btnQuit, "Quitter")
GUICtrlSetOnEvent($btnQuit, "OnClose")

While 1
	Sleep(100)
WEnd

Func OnClose()
	Exit
EndFunc   ;==>OnClose

Func BrowseFile()

	$sFilePath = FileOpenDialog("Choisir un fichier", @ScriptDir, "Fichiers (*.*)", 1)
	If @error <> 0 Then Return
	GUICtrlSetData($inputFile, $sFilePath)


EndFunc   ;==>BrowseFile

Func SetOperation()
	If GUICtrlRead($radioCrypt) = $GUI_CHECKED Then
		$sOperation = "Crypter"
	ElseIf GUICtrlRead($radioDecrypt) = $GUI_CHECKED Then
		$sOperation = "Décrypter"
	EndIf
EndFunc   ;==>SetOperation

Func ExecuteOperation()
	$sFilePath = GUICtrlRead($inputFile)
	$sKey = GUICtrlRead($inputKey)

	If $sFilePath = "" Or $sKey = "" Then
		MsgBox(16, "Erreur", "Veuillez spécifier un fichier et une clé de chiffrement.")
		Return
	EndIf

	If $sOperation = "Crypter" Then
		_FileWriteLog($Cheminlog, "***CRYPT***  Cryptage du fichier : " & $sFilePath)

		cryptage($sFilePath, 1, $sKey)

		; Récupération du l'emplacement du fichier
		Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
		Local $aPathSplit = _PathSplit($sFilePath, $sDrive, $sDir, $sFileName, $sExtension)

		ShellExecute($aPathSplit[1] & $aPathSplit[2])

		MsgBox($MB_ICONINFORMATION, "Succès", "Le fichier a été crypté avec succès.")

	ElseIf $sOperation = "Décrypter" Then
		_FileWriteLog($Cheminlog, "***CRYPT***  Decryptage du fichier : " & $sFilePath)
		$sFilePath = GUICtrlRead($inputFile)
		$sKey = GUICtrlRead($inputKey)

		If $sFilePath = "" Or $sKey = "" Then
			MsgBox(16, "Erreur", "Veuillez spécifier un fichier et une clé de chiffrement.")
			Return
		EndIf
		cryptage($sFilePath, 2, $sKey)
		ShellExecute($sFilePath, "", "", "")
		_FileWriteLog($Cheminlog, "Déryptage du fichier : " & $sFilePath)
		; À compléter pour le décryptage
	EndIf
EndFunc   ;==>ExecuteOperation
