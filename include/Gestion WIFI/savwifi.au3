#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\Wifi Connect.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.7
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include "..\au3\ficconf.au3"

; paramÃ©trage GÃ©nÃ©ral
Opt("GUICloseOnESC", 0) ; 0=pas de sortie par la touche ESC 1= Fermeture du programme via esc
;Opt("TrayAutoPause", 0) ;0=pas de pause, 1=pause
Opt("TrayIconHide", 1) ;0=montre, 1=masque l'icÃ´ne de Notification



Global $ippc = IniRead($paramini, "PC", "Nom", "")
_FileWriteLog($Cheminlog, "Controle du WIFI sur le poste : " & $ippc)
$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Wifi", "Version", "REG_SZ", $version)

#Region ### START Koda GUI section ### Form=C:\CD47Users\frfoulqu\Documents\CD47\Scripts\Autoit\Outils Admin\include\Gestion WIFI\wifi.kxf
Global $FormGestWifi = GUICreate("Gestion du Wifi", 326, 286, 424, 354)
GUISetIcon("C:\Icones\Metro Icons\System Icons\OTHERS\WIFI\ICO\Wifi Connect.ico", -1)
Global $BT_SAVWIFI = GUICtrlCreateButton("Sauvegarder", 24, 184, 105, 25)
Global $BT_RestoWIFI = GUICtrlCreateButton("Restaurer", 168, 184, 105, 25)
Global $BT_QuitGestWifi = GUICtrlCreateButton("Quitter", 120, 240, 81, 25)
Global $Label1 = GUICtrlCreateLabel("Gestion des profils Wifi", 56, 32, 190, 24)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
Global $Label2 = GUICtrlCreateLabel("sur le PC " & $ippc, 66, 80, 180, 20)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
Global $Label3 = GUICtrlCreateLabel($version, 248, 264, 39, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $BT_QuitGestWifi
			Exit
		Case $BT_SAVWIFI
			Global $sFileSelectFolder = "\\" & $ippc & "\c$\temp\wifi\" & $ippc
			DirCreate($sFileSelectFolder)
			If @error Then
				; Affiche le message d'erreur.
				MsgBox($MB_SYSTEMMODAL, "", "Aucun dossier n'a été sélectionné.")
			Else
				;copie du script vers le PC distant
				FileCopy($chemin_dossier_bin & "savWIFI.bat", $sFileSelectFolder)
				;exÃ©cution du script
				RunWait($chemin_dossier_bin & "PsExec.exe /accepteula \\" & $ippc & " -s " & $sFileSelectFolder & "\savwifi.bat")
				;suppression du script
				FileDelete($sFileSelectFolder & "\savwifi.bat")
				;******************************
				; Cryptage des fichiers
				;******************************
				cryptwifi()

				;rapatriement des donnÃ©es sur le pc local
				FileCopy($chemin_dossier_bin & "restoWiFi.bat", $sFileSelectFolder & "\wifi\", $FC_OVERWRITE)
				DirMove("\\" & $ippc & "\c$\temp\wifi\", "c:\temp\", $FC_OVERWRITE + $FC_CREATEPATH)
				_FileWriteLog($Cheminlog, "Sauvegarde des paramètres wifi du poste : " & $ippc & " terminée")
				MsgBox($MB_SYSTEMMODAL, "", "Les paramètres WIFI sont sauvegardés sous C:\temp\wifi" & @CRLF & $ippc & "\Wifi")


			EndIf

		Case $BT_RestoWIFI
			; choix du dossier Ã  restaurer
			MsgBox(64, "Information", "Sélectionner le fichier restoWifi.bat du PC.")
			Local $sFileOpenDialog = FileOpenDialog("choix du fichier", "c:\temp\wifi\", "bat (restoWiFi.bat)", $FD_FILEMUSTEXIST + $FD_MULTISELECT)
			If @error Then
				; Affiche le message d'erreur.
				MsgBox($MB_SYSTEMMODAL, "", "Aucun fichier sélectionné.")
			Else
				#Region --- CodeWizard generated code Start ---
				;MsgBox features: Title=Yes, Text=Yes, Buttons=OK and Cancel, Icon=Warning
				If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
				$iMsgBoxAnswer = MsgBox(49, "choix", "Vous avez sélectionner le dossier suivant : " & @CRLF & @CRLF & $sFileOpenDialog & @CRLF & @CRLF & "Voulez vous continuer?")
				Select
					Case $iMsgBoxAnswer = 1 ;OK

						Local $DossierLocal = StringTrimRight($sFileOpenDialog, 13) ; récupère le nom de dossier en enlevant restoWifi.bat
						FileCopy($DossierLocal & "*.*", "\\" & $ippc & "\c$\temp\wifi\", 9) ; transfère les données sur l'ordinateur distant 9 pour remplacer et créer l'arborescence
						;***************************
						; décrypte les données
						;***************************
						decryptwifi()
						;Copie du script de restauration


						RunWait($chemin_dossier_bin & "PsExec.exe /accepteula \\" & $ippc & " -s c:\temp\wifi\restoWifi.bat") ; lance le script de resto
						_FileWriteLog($Cheminlog, "Restauration des paramètres wifi du poste : " & $ippc & " terminée")
						DirRemove("\\" & $ippc & "\c$\temp\wifi\", 1)
						MsgBox(64, "Terminé", "Restauration effectuée.")
					Case $iMsgBoxAnswer = 2 ;Cancel
						Exit
				EndSelect
				#EndRegion --- CodeWizard generated code Start ---


			EndIf
	EndSwitch
WEnd

Func cryptwifi()

	; Chemin du dossier contenant les fichiers XML cryptés
	Local $sFolder = $sFileSelectFolder & "\wifi\"

	; Clé de cryptage (la même que celle utilisée pour le cryptage)
	Local $sKey = "MaCleDeCryptage"

	; Parcours de tous les fichiers XML dans le dossier
	Local $aFileList = _FileListToArray($sFolder, "*.xml", $FLTA_FILES)

	If @error Then
		MsgBox(16, "Erreur", $sFileSelectFolder & "\wifi\ Aucun fichier XML trouvé dans le dossier spécifié.")
		Exit
	EndIf

	For $i = 1 To $aFileList[0]
		Local $sFilePath = $sFolder & "\" & $aFileList[$i]

		; Lecture du contenu du fichier XML
		Local $sContent = FileRead($sFilePath)

		; Cryptage du contenu
		Local $sEncryptedContent = _Crypt_EncryptData($sContent, $sKey, $CALG_AES_256)

		; Écriture du contenu crypté dans le même fichier
		FileDelete($sFilePath)
		FileWrite($sFilePath, $sEncryptedContent)
	Next

EndFunc   ;==>cryptwifi


Func decryptwifi()

	; Chemin du dossier contenant les fichiers XML cryptés
	Local $sFolder = "\\" & $ippc & "\c$\temp\wifi\"

	; Clé de cryptage (la même que celle utilisée pour le cryptage)
	Local $sKey = "MaCleDeCryptage"

	; Parcours de tous les fichiers XML dans le dossier
	Local $aFileList = _FileListToArray($sFolder, "*.xml", $FLTA_FILES)

	If @error Then
		MsgBox(16, "Erreur", "Aucun fichier XML trouvé dans le dossier spécifié.")
		Exit
	EndIf

	For $i = 1 To $aFileList[0]
		Local $sFilePath = $sFolder & "\" & $aFileList[$i]

		; Lecture du contenu crypté du fichier XML
		Local $sEncryptedContent = FileRead($sFilePath)

		; Décryptage du contenu
		Local $sDecryptedContent = _Crypt_DecryptData($sEncryptedContent, $sKey, $CALG_AES_256)

		; Écriture du contenu décrypté dans le même fichier
		FileDelete($sFilePath)
		FileWrite($sFilePath, $sDecryptedContent)
	Next



EndFunc   ;==>decryptwifi
