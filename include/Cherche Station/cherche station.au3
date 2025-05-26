#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\icones\Lenovo.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.3
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
Global $rep_MAC = "\\vulcain\fitrans$\dsian\log\dock"
FileInstall(".\cherche station.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)


$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Cherchestation", "Version", "REG_SZ", $version)

#Region ### START Koda GUI section ### Form=
Global $_1 = GUICreate("Recherche des station LENOVO  version : " & $version, 402, 302, 553, 186)
Global $label = GUICtrlCreateLabel("Saisir le N° de série de la station", 10, 10, 250, 20)
Global $input = GUICtrlCreateInput("", 10, 30, 250, 21)
Global $searchButton = GUICtrlCreateButton("  Rechercher", 10, 60, 100, 36)
GUICtrlSetImage($searchButton, $dlloa, 8, 1)
Global $quitButton =GUICtrlCreateButton("     Quitter", 280, 60, 100, 36)
GUICtrlSetImage($quitButton, $dlloa, 4, 1)
Global $results = GUICtrlCreateEdit("", 10, 100, 380, 190)
;GUICtrlSetData(-1, "resultList")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###



While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $quitButton
			;Exit
			;Case $quitButton
			Exit

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
