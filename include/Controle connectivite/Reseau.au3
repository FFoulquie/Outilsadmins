#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\iCloud.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.8
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include "..\au3\ficconf.au3"
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <cryptage.au3>

;Opt("MustDeclareVars", 1)




Global $PathIni = $configini

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]

RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Reseau", "Version", "REG_SZ", $version)

FileInstall("C:\CD47Users\frfoulqu\Documents\CD47\Scripts\Autoit\Outils Admin\include\Controle connectivite\reseau.au3", $chemin_dossier_bin, $FC_OVERWRITE)
Global $Key = "5aQXvTCr4dhBF48zusUC"

#Region ### START Koda GUI section ### Form=
Global $Form1 = GUICreate("Test réseau", 443, 438, 442, 260)
Global $Label1 = GUICtrlCreateLabel("Vérification de connectivité", 120, 24, 332, 17)
GUICtrlSetFont(-1, 12, 800, 0, "MV Boli")
Global $CB_Connect = GUICtrlCreateCombo("", 128, 72, 193, 25)
GUICtrlSetData($CB_Connect, $listeNas)
Global $BtVal = GUICtrlCreateButton("Valider", 144, 112, 129, 33)
Global $Label2 = GUICtrlCreateLabel("PCS", 88, 288, 25, 17)
Global $LabelSwitch = GUICtrlCreateLabel("Switch", 88, 192, 36, 17)
Global $LabelNas = GUICtrlCreateLabel("Nas", 88, 224, 36, 17)
Global $LabelIPGD = GUICtrlCreateLabel("IPGD", 88, 256, 30, 17)
Global $LabelRouteur = GUICtrlCreateLabel("Routeur", 88, 160, 42, 17)
Global $LArouteur = GUICtrlCreateLabel("", 168, 160, 40, 17)
Global $LASwitch = GUICtrlCreateLabel("", 168, 192, 40, 17)
Global $LAnbswitch = GUICtrlCreateLabel("", 210, 192, 100, 17)
Global $LANas = GUICtrlCreateLabel("", 168, 224, 40, 17)
Global $LAIPGD = GUICtrlCreateLabel("", 168, 256, 40, 17)
Global $LAPCS = GUICtrlCreateLabel("", 168, 288, 40, 17)
Global $btquit = GUICtrlCreateButton("Quitter", 112, 336, 81, 33)
Global $btstop = GUICtrlCreateButton("Arrêter", 227, 336, 81, 33, $WS_DISABLED)
Global $label3 = GUICtrlCreateLabel($version, 10, 420, 81, 33)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $btquit

			FileDelete($chemin_dossier_bin & "reseau.au3")
			Exit
		Case $BtVal
			;GUICtrlSetState($btquit, $GUI_DISABLED)
			GUICtrlSetState($btstop, $GUI_ENABLE)


			VerifyConnectivity()
		Case $btstop
			GUICtrlSetState($btquit, $GUI_ENABLE)

			ExitLoop




	EndSwitch
WEnd

Func VerifyConnectivity()
	Local $Site = GUICtrlRead($CB_Connect)
	_FileWriteLog($Cheminlog, "**Reseau**  Vérification connection réseau sur le site " & $Site & "'")
	GUICtrlSetData($LArouteur, "En cours...")
	GUICtrlSetData($LASwitch, "En cours...")
	GUICtrlSetData($LANas, "En cours...")
	GUICtrlSetData($LAIPGD, "En cours...")
	GUICtrlSetData($LAPCS, "En cours...")

	cryptage($configini, 2, $Key)
	Global $IPRouteur = IniRead($PathIni, $Site, "Routeur", "")
	Global $IPSwitch = IniRead($PathIni, $Site, "Switch", "")
	Global $IPnas = IniRead($PathIni, $Site, "Nas", "")
	Global $IPIPGD = IniRead($PathIni, $Site, "IPGD", "")
	Global $IPPCS = IniRead($PathIni, $Site, "PCS", "")
	Global $nbswitch = IniRead($PathIni, $Site, "NBSwitch", "")
	cryptage($configini, 1, $Key)
	While 1


		;msgbox(0,"",$IPRouteur)
		$LAnbswitch = GUICtrlCreateLabel("Nb switch : " & $nbswitch, 208, 192, 100, 17)
		Local $PingR = Ping($IPRouteur, 250)
		;msgbox(0,"",$PingR)
		If $PingR > "4" Then
			$LArouteur = GUICtrlCreateLabel(" OK", 168, 160, 40, 17)
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			GUICtrlSetColor(-1, 0x33FF33)
		Else
			$LArouteur = GUICtrlCreateLabel("KO", 168, 160, 40, 17)
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			GUICtrlSetColor(-1, 0xFF0000)
		EndIf

		If $IPSwitch = "" Then
			$LASwitch = GUICtrlCreateLabel(" N/C", 168, 192, 40, 17)
		Else
			Local $PingS = Ping($IPSwitch, 250)
			If $PingS > "4" Then
				$LASwitch = GUICtrlCreateLabel(" OK", 168, 192, 40, 17)
				GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
				GUICtrlSetColor(-1, 0x33FF33)
			Else
				$LASwitch = GUICtrlCreateLabel("KO", 168, 192, 40, 17)
				GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
				GUICtrlSetColor(-1, 0xFF0000)
			EndIf
		EndIf

		Local $Pingn = Ping($IPnas, 250)
		If $Pingn > "4" Then
			$LANas = GUICtrlCreateLabel(" OK", 168, 224, 40, 17)
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			GUICtrlSetColor(-1, 0x33FF33)
		Else
			$LANas = GUICtrlCreateLabel("KO", 168, 224, 40, 17)
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			GUICtrlSetColor(-1, 0xFF0000)
		EndIf
		If $IPIPGD = "" Then
			$LAIPGD = GUICtrlCreateLabel(" N/C", 168, 256, 40, 17)
		Else
			Local $PingI = Ping($IPIPGD, 250)
			If $PingI > "4" Then
				$LAIPGD = GUICtrlCreateLabel(" OK", 168, 256, 40, 17)
				GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
				GUICtrlSetColor(-1, 0x33FF33)
			Else
				$LAIPGD = GUICtrlCreateLabel("KO", 168, 256, 40, 17)
				GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
				GUICtrlSetColor(-1, 0xFF0000)
			EndIf
		EndIf
		If $IPPCS = "" Then
			$LAPCS = GUICtrlCreateLabel(" N/C", 168, 288, 40, 17)
		Else
			Local $PingP = Ping($IPPCS, 250)
			If $PingP > "4" Then
				$LAPCS = GUICtrlCreateLabel(" OK", 168, 288, 40, 17)
				GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
				GUICtrlSetColor(-1, 0x33FF33)
			Else
				$LAPCS = GUICtrlCreateLabel("KO", 168, 288, 40, 17)
				GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
				GUICtrlSetColor(-1, 0xFF0000)
			EndIf
		EndIf
		Sleep(100)
		ConsoleWrite($btstop & @CRLF)
		; Vérifie si le bouton "Arrêter" est cliqué
		If GUIGetMsg() = $btstop Then
			ConsoleWrite($btstop & @CRLF)

			ExitLoop
		EndIf
	WEnd

EndFunc   ;==>VerifyConnectivity
