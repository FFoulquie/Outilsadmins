#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\macadress.ico
#AutoIt3Wrapper_Outfile=cherche PC.exe
#AutoIt3Wrapper_Res_Fileversion=1.0.0.2
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstantsEx.au3>
#include <File.au3>
#include <GuiListView.au3>
#include <GUIConstants.au3>
#include "..\au3\ficconf.au3"

;Func AdmlocauxPC()
Global $elements
Global $numLigne = ""
Global $resultList = ""
Global $rep_du_site = "\\vulcain\fitrans$\dsian\LOG"
Global $rep_MAC = "\\vulcain\fitrans$\dsian\mac"


$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\CherchePC", "Version", "REG_SZ", $version)

#Region ### START Koda GUI section ### Form=
Global $_1 = GUICreate("Recherche des ordinateurs  version : " & $version, 402, 302, 553, 186)
Global $label = GUICtrlCreateLabel("Saisir l'adresse MAC, le N° de série ou le code OPALE", 10, 10, 350, 20)
Global $input = GUICtrlCreateInput("", 10, 30, 250, 21)
Global $searchButton = GUICtrlCreateButton("Rechercher MAC", 10, 60, 100, 30)
Global $BT_SearchSN = GUICtrlCreateButton("Rechercher S/N et Opale", 138, 60, 100, 30,$BS_MULTILINE)
;
Global $quitButton = GUICtrlCreateIcon($chemin_dossier_bin & "conf\OA.dll", 4 , 280, 60, 32, 32)
	GUICtrlSetTip($quitButton, "Quitter")
Global $results = GUICtrlCreateEdit("", 10, 100, 380, 190)
;GUICtrlSetData(-1, "resultList")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###



While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $quitButton
			Exit
			;Case $quitButton
			;Exit

		Case $searchButton
			SplashTextOn("Recherche en cours", "Merci de  patienter", 250, 120)
			$resultList = ""
			$searchString = GUICtrlRead($input)
			$files = _FileListToArray($rep_MAC, "*.*", $FLTA_FILES)

			For $i = 1 To $files[0]
				$filePath = $rep_MAC & "\" & $files[$i]
				If _findStringInFile($filePath, $searchString) Then
					$resultList &= $files[$i] & @CRLF
				EndIf
			Next
			$arr = StringSplit($resultList, ".")
			$gauche = $arr[1]
			GUICtrlSetData($results, $gauche)
			SplashOff()

		Case $BT_SearchSN
			SplashTextOn("Recherche en cours", "Merci de  patienter", 250, 120)
			$resultList = ""
			$searchString = GUICtrlRead($input)
			$files = _FileListToArray($rep_du_site, "*.txt", $FLTA_FILES)
			For $i = 1 To $files[0]
				$filePath = $rep_du_site & "\" & $files[$i]
				If _findStringInFile($filePath, $searchString) Then

					SplashOff()
					$file = FileOpen($rep_du_site & "\" & $files[$i], 0)

					; Vérifier si le fichier a été ouvert avec succès.
					If $file = -1 Then
						MsgBox(0, "Erreur", "Impossible d'ouvrir le fichier.")
						Exit
					EndIf

					; Initialiser le compteur de lignes.
					$numeroLigne = 0

					; Parcourir le fichier ligne par ligne.
					While 1
						$ligne = FileReadLine($file)
						If @error = -1 Then ExitLoop ; Sortir de la boucle si nous atteignons la fin du fichier.

						; Incrémenter le compteur de lignes.
						$numeroLigne += 1

						; Vérifier si la ligne contient la variable recherchée.
						If StringInStr($ligne, $searchString) Then

							$numLigne += 1
							; Utiliser StringSplit pour diviser la ligne en sous-chaînes en utilisant le point-virgule comme délimiteur.
							$elements = StringSplit($ligne, ";")
							ExitLoop
						EndIf
					WEnd
					; Fermer le fichier.
					FileClose($file)

					$resultList &= $files[$i] & @CRLF
				EndIf
			Next
			$arr = StringSplit($resultList, ".")
			$gauche = $arr[1]
			SplashOff()
			If $gauche <> "" Then
				GUICtrlSetData($results, $elements[4])
			Else
				MsgBox(0, "", "PC non trouvé")
			EndIf

	EndSwitch
WEnd
;EndFunc


Func _findStringInFile($filePath, $searchString)
	$fileHandle = FileOpen($filePath, $FO_READ)
	If $fileHandle = -1 Then
		Return False
	EndIf
	$fileContents = FileRead($fileHandle)
	$pos = StringInStr($fileContents, $searchString)
	FileClose($fileHandle)
	If $pos = 0 Then
		Return False
	EndIf
	Return True

EndFunc   ;==>_findStringInFile
