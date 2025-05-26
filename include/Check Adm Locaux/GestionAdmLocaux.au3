#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\Graduation alt.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.3
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include "..\au3\ficconf.au3"
#include <File.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\GestAdminslocaux", "Version", "REG_SZ", $version)

FileInstall(".\GestionAdmLocaux.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)
;Func AdmlocauxPC()
Global $fichier_Complet = "c:\temp\admin.txt"
Global $fichier_temp1 = "c:\temp\admins1.txt"
Global $fichier_temp2 = "c:\temp\admins2.txt"
Global $fichier_final = "c:\temp\admins.txt"

Global $resultList = ""
Global $rep_du_site = "\\vulcain\fitrans$\dsian\AdminLocal"
GUICreate("Recherche des ordinateurs", 400, 300)

#Region ### START Koda GUI section ### Form=
Global $_1 = GUICreate("Recherche des comptes admins", 401, 301, 553, 186, $WS_CAPTION)
Global $label = GUICtrlCreateLabel("Saisir le nom d'ouverture de session :", 10, 10, 250, 20)
Global $input = GUICtrlCreateInput("", 10, 30, 250, 21)
Global $searchButton = GUICtrlCreateButton("Rechercher", 154, 60, 100, 30)
;Global $quitButton = GUICtrlCreateButton("Quitter", 288, 60, 100, 30)
Global $results = GUICtrlCreateEdit("", 10, 100, 380, 190)
GUICtrlSetData(-1, "")
Global $Genebutton = GUICtrlCreateButton("Générer Liste", 10, 60, 100, 30)

	$quitButton = GUICtrlCreateIcon($dlloa, 4, 288, 60, 32, 32)
	GUICtrlSetTip($quitButton, "Quitter")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###



While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $quitButton
			FileDelete($chemin_dossier_bin & "GestionAdmLocaux.au3")
			FileDelete($fichier_Complet)
			FileDelete($fichier_temp1)
			FileDelete($fichier_temp2)
			FileDelete($fichier_final)

			Exit
		Case $searchButton
			SplashTextOn("", "Traitement en cours", 250, 120)

			$resultList = ""
			$searchString = GUICtrlRead($input)
			$files = _FileListToArray($rep_du_site, "*.*", $FLTA_FILES)

			For $i = 1 To $files[0]
				$filePath = $rep_du_site & "\" & $files[$i]
				If _findStringInFile($filePath, $searchString) Then
					$resultList &= $files[$i] & @CRLF
				EndIf
			Next
			GUICtrlSetData($results, $resultList)

			SplashOff()

		Case $Genebutton

			SplashTextOn("Traitement en cours", "Merci de  patienter", 250, 120)

			concatene()
			supp()
			suppChiffre()
			supplignevide()

			SplashOff()
			ShellExecute($fichier_final)

			;MsgBox($MB_ICONINFORMATION, "Info", "Fin du traitement")
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

Func concatene()

	Local $dossier = "\\vulcain\fitrans$\dsian\AdminLocal"


	Local $contenu_final = ""

	Local $fichiers = _FileListToArray($dossier, "*.txt", $FLTA_FILES)
	If @error Then
		MsgBox($MB_ICONERROR, "Erreur", "Erreur lors de la recherche des fichiers dans le dossier.")
		Exit
	EndIf

	For $i = 1 To $fichiers[0]
		Local $contenu_fichier = FileRead($dossier & "\" & $fichiers[$i])
		$contenu_final &= $contenu_fichier & @CRLF
	Next

	Local $fileHandle = FileOpen($fichier_Complet, $FO_OVERWRITE)
	If $fileHandle = -1 Then
		MsgBox($MB_ICONERROR, "Erreur", "Impossible d'ouvrir le fichier pour l'écriture.")
		Exit
	EndIf

	FileWrite($fileHandle, $contenu_final)
	FileClose($fileHandle)
EndFunc   ;==>concatene

Func supp()

	Local $contenu = FileRead($fichier_Complet)
	Local $valeurs_a_supprimer[9] = ["administrateur", "Admins du domaine", "Adm_job", "Admin_locaux_pc", "admapophis", "adm_loc_sig", "adm_iris", "adm_sig", "AdmLocal"]

	For $i = 0 To UBound($valeurs_a_supprimer) - 1
		$contenu = StringReplace($contenu, $valeurs_a_supprimer[$i], "")
	Next

	Local $fileHandle = FileOpen($fichier_temp1, $FO_OVERWRITE)
	If $fileHandle = -1 Then
		MsgBox($MB_ICONERROR, "Erreur", "Impossible d'ouvrir le fichier pour l'écriture.")
		Exit
	EndIf

	FileWrite($fileHandle, $contenu)
	FileClose($fileHandle)
EndFunc   ;==>supp

Func suppChiffre()



	Local $contenu = FileRead($fichier_temp1)
	Local $nouveau_contenu = ""

	For $i = 1 To StringLen($contenu)
		Local $caractere = StringMid($contenu, $i, 1)
		If Not StringIsDigit($caractere) Then
			$nouveau_contenu &= $caractere
		EndIf
	Next

	Local $fileHandle = FileOpen($fichier_temp2, $FO_OVERWRITE)
	If $fileHandle = -1 Then
		MsgBox($MB_ICONERROR, "Erreur", "Impossible d'ouvrir le fichier pour l'écriture.")
		Exit
	EndIf

	FileWrite($fileHandle, $nouveau_contenu)
	FileClose($fileHandle)

EndFunc   ;==>suppChiffre

Func supplignevide()

	Local $contenu = FileRead($fichier_temp2)
	Local $nouveau_contenu = ""

	Local $lines = StringSplit($contenu, @CRLF, 1)
	For $i = 1 To $lines[0]
		If StringStripWS($lines[$i], 3) <> "" Then
			$nouveau_contenu &= $lines[$i] & @CRLF
		EndIf
	Next

	Local $fileHandle = FileOpen($fichier_final, $FO_OVERWRITE)
	If $fileHandle = -1 Then
		MsgBox($MB_ICONERROR, "Erreur", "Impossible d'ouvrir le fichier pour l'écriture.")
		Exit
	EndIf

	FileWrite($fileHandle, $nouveau_contenu)
	FileClose($fileHandle)


EndFunc   ;==>supplignevide
