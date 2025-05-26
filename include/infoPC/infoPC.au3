#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\Info.ico
#AutoIt3Wrapper_Res_Fileversion=2.0.0.2
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_ProductVersion=2
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

	AutoIt Version : 3.3.14.5
	Auteur:         F.Foulquié

	Fonction du Script :
	Le script permet de mettre en page les informations principale concernant l'ordinateur sélectionné.
	Nécessite l'exécution préalable du script BDD qui au démmarage de l'ordinateur créée un ficgier contenant les informations à transmettre

#ce ----------------------------------------------------------------------------
#include <ButtonConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include "..\au3\ficconf.au3"
#include <EditConstants.au3>



FileInstall(".\infoPC.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)

$ippc = IniRead($chemin_dossier_bin & "conf\param.flq", "PC", "Nom", "")
$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\infoPC", "Version", "REG_SZ", $version)

Global $Processeur = ""
Global $Opale = ""
Global $marque = ""
Global $Type = ""
Global $Ram = ""
Global $user = ""
Global $datec = ""
Global $IP = ""
Global $Mac = ""
Global $OS = ""
Global $SN = ""
Global $DispoC = ""
Global $TotalC = ""


Global $numLigne, $elements

$searchString = InputBox("Nom de l'ordinateur", "saisir le nom de l'ordinateur a analyser", $ippc)

; Ouvrir le fichier BDD en mode lecture
Global $rep_du_site = "T:\dsian\Log\bdd\" & $searchString & ".txt"
If @error = 1 Then
	Exit
EndIf


SplashTextOn("Recherche en cours", "Merci de patienter", 250, 120, -1, -1, 16)
$resultList = ""

$file = FileOpen($rep_du_site)
If $file = -1 Then

	$Opale = ""
	$marque = ""
	$Type = ""
	$Ram = ""
	$user = ""
	$datec = ""
	$IP = ""
	$Mac = ""
	$OS = ""
	$SN = ""
	$Processeur = ""
	$DispoC = ""
	$TotalC = ""
	SplashOff()

	MsgBox(48, "Erreur", "Ordinateur inconnu")

	Exit
EndIf

lectureFichier($file, $searchString)
FileClose($file)

;If UBound($elements) > 13 Then

$Opale = $elements[7]
$marque = $elements[8]
$Type = $elements[9]
$Ram = $elements[6]
$user = $elements[3]
$datec = $elements[2]
$IP = $elements[4]
$SN = $elements[5]
$Mac = $elements[10]
$OS = $elements[11]
$Processeur = $elements[12]
$DispoC = $elements[13]
$TotalC = $elements[14]
#cs
	Else
	$Opale = $elements[7]
	$marque = $elements[8]
	$Type = $elements[9]
	$Ram = $elements[6]
	$user = $elements[3]
	$datec = $elements[2]
	$IP = $elements[4]
	$SN = $elements[5]
	$Mac = $elements[10]
	$OS = $elements[11]
	$Processeur = $elements[12]
	$DispoC = ""
	$TotalC = ""
	EndIf
#ce


Global $finalString = ""

; Gestion du fichier Dock

Local $filePathE = "T:\dsian\Log\ecran\" & $searchString & ".txt"
Local $fileHandleE = FileOpen($filePathE, 0)

If $fileHandleE = -1 Then
	$finalString = "N/C"
Else
	; Initialiser le tableau pour stocker les chaînes extraites
	Local $stringsArray[1]
	Local $arrayIndex = 0

	; Lire le fichier ligne par ligne
	While Not @error
		Local $line = FileReadLine($fileHandleE)
		If @error = -1 Then ExitLoop

		; Chercher le deuxième ":"
		Local $pos2 = StringInStr($line, ":", 0, 2)
		If $pos2 > 0 Then
			; Extraire et nettoyer la chaîne après le deuxième ":"
			Local $string = StringStripWS(StringTrimLeft($line, $pos2), 3)

			; Redimensionner et ajouter la chaîne au tableau
			ReDim $stringsArray[$arrayIndex + 1]
			$stringsArray[$arrayIndex] = $string
			$arrayIndex += 1
		EndIf
	WEnd

	; Fermer le fichier
	FileClose($fileHandleE)

	; Concaténer toutes les chaînes manuellement
	For $i = 0 To UBound($stringsArray) - 1
		$finalString &= $stringsArray[$i] & @CRLF
	Next

	; Si rien n'a été extrait, mettre "N/C"
	If $finalString = "" Then
		$finalString = "N/C"
	EndIf
EndIf


; Gestion du fichier Dock

Local $filePathD = "T:\dsian\Log\dock\" & $searchString & ".txt"

Local $fileHandleD = FileOpen($filePathD, 0)

If $fileHandleD = -1 Then
	$SNDock = "N/C"
Else
	; Lire le fichier ligne par ligne
	While Not @error
		Local $lineD = FileReadLine($fileHandleD)
		If @error = -1 Then ExitLoop

		; Chercher le  ":"
		Local $pos2 = StringInStr($lineD, ":", 0, 1)
		If $pos2 > 0 Then
			; Extraire et nettoyer la chaîne après le  ":"
			Local $SNDock = StringStripWS(StringTrimLeft($lineD, $pos2), 3)
		EndIf
	WEnd
EndIf

SplashOff()


#Region ### START Koda GUI section ### Form=C:\CD47Users\frfoulqu\Documents\CD47\Scripts\Autoit\Outils Admin\include\infoPC\info.kxf
Global $guiresultats = GUICreate("info PC    v:" & $version, 736, 500)
GUISetBkColor(0xFFFFFF)
Global $LNomPc = GUICtrlCreateLabel($searchString, 160, 24, 146, 33)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
Global $LMarque = GUICtrlCreateLabel($marque, 32, 80, 103, 17)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
Global $BtQuitter = GUICtrlCreateButton("     Quitter", 360, 448, 100, 36)
GUICtrlSetImage($BtQuitter, $dlloa, 4, 1)
GUICtrlSetTip($BtQuitter, "Quitter")
Global $LRef = GUICtrlCreateLabel($Type, 232, 80, 223, 33)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
Global $LRAM = GUICtrlCreateLabel($Ram, 32, 120, 30, 17)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
Global $Label13 = GUICtrlCreateLabel("GO de RAM", 70, 120, 103, 17)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
Global $LOpale = GUICtrlCreateInput($Opale, 264, 120, 120, 24, -1, 0)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
Global $Label14 = GUICtrlCreateLabel("Opale :", 188, 120, 63, 20)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
Global $Group1 = GUICtrlCreateGroup("Gestion des Accès", 24, 312, 441, 116, BitOR($GUI_SS_DEFAULT_GROUP, $BS_RIGHTBUTTON, $BS_CENTER), $WS_EX_TRANSPARENT)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xFF0000)
Global $Label2 = GUICtrlCreateLabel("Dernier accès", 32, 344, 183, 17)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $Label3 = GUICtrlCreateLabel("par :", 32, 392, 33, 20)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
;GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $LUser = GUICtrlCreateLabel($user, 70, 392, 183, 33)
GUICtrlSetFont(-1, 12, 0, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $LDate = GUICtrlCreateLabel($datec, 55, 368, 200, 17)
GUICtrlSetFont(-1, 12, 0, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $Label7 = GUICtrlCreateLabel("le", 32, 368, 23, 17)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xFFFFFF)
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $Label11 = GUICtrlCreateLabel($Processeur, 32, 160, 400, 24)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
Global $LIP = GUICtrlCreateLabel("IP :", 32, 216, 17, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
Global $RIP = GUICtrlCreateInput($IP, 64, 216, 90, 21, -1, 0)
Global $LMAC = GUICtrlCreateLabel("MAC :", 160, 216, 27, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
Global $RMAC = GUICtrlCreateInput($Mac, 190, 216, 225, 21, -1, 0)
Global $ROS = GUICtrlCreateLabel($OS, 336, 8, 129, 21)
;Global $LSN = GUICtrlCreateLabel("S/N", 32, 188, 24, 17)
;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
;Global $RSerie = GUICtrlCreateInput($SN, 72, 188, 225, 21, -1, 0)
Global $LabC = GUICtrlCreateLabel("C: ", 32, 248, 20, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
Global $LabDispoC = GUICtrlCreateLabel($DispoC, 56, 248, 57, 21)
Global $LabCtexte = GUICtrlCreateLabel("Go disponible sur ", 120, 248, 88, 17)
Global $LabCtotal = GUICtrlCreateLabel($TotalC, 208, 248, 49, 17)
Global $Labtextec2 = GUICtrlCreateLabel("Go au total", 264, 248, 56, 17)

Global $LabEcran = GUICtrlCreateLabel("Numéros de séries", 515, 86, 167, 17)
GUICtrlSetFont($LabEcran, 12, 800, 0, "MS Sans Serif")
Global $ISNecran = GUICtrlCreateEdit("", 515, 112, 167, 280, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN))
GUICtrlSetData($ISNecran, "*********" & @CRLF & "PC" & @CRLF & "*********" & @CRLF & $SN & @CRLF & @CRLF & "*********" & @CRLF & "écrans" & @CRLF & "*********" & @CRLF & $finalString & @CRLF & @CRLF & "*********" & @CRLF & "Station " & @CRLF & "*********" & @CRLF & $SNDock, 1)
GUICtrlSetFont(-1, 8, 400, 0, "MS Sans Serif")

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $BtQuitter
			GUICtrlSetState($guiresultats, $GUI_DISABLE)
			FileDelete($chemin_dossier_bin & "infoPC.au3")
			Exit
	EndSwitch
WEnd


Func lectureFichier($file, $searchString)
	Local $numeroLigne = _FileCountLines($file)

	While $numeroLigne > 0
		$ligne = FileReadLine($file, $numeroLigne)
		If StringInStr($ligne, $searchString) Then
			$elements = StringSplit($ligne, ",")

			ExitLoop
		EndIf
		$numeroLigne -= 1
	WEnd
EndFunc   ;==>lectureFichier


