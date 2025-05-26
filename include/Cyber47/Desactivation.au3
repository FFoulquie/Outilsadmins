#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\DNS.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <AD.au3>
#include <Date.au3>

; paramétrage Général
Opt("GUICloseOnESC", 0) ; 0=pas de sortie par la touche ESC 1= Fermeture du programme via esc
;Opt("TrayAutoPause", 0) ;0=pas de pause, 1=pause
Opt("TrayIconHide", 1) ;0=montre, 1=masque l'icône de Notification

FileInstall("C:\CD47Users\frfoulqu\Documents\Mes images\cyber47.jpg", @ScriptDir & "\cyber47.jpg", $FC_OVERWRITE)
FileInstall("K:\Reseau\100-Procedures urgence\schema alerte admins.pdf", @ScriptDir & "\admin.pdf", $FC_OVERWRITE)
FileInstall("K:\Reseau\100-Procedures urgence\schema alerte tech.pdf", @ScriptDir & "\tech.pdf", $FC_OVERWRITE)

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Cyber", "Version", "REG_SZ", $version)

Local $MENDOC, $AdminDoc, $TechDoc, $Gencode

_fenetre()

Func _fenetre()
	#Region ### START Koda GUI section ### Form=
	Global $Form1 = GUICreate("Désactivation d'urgence", 385, 392)
	GUISetBkColor(0xFFFFFF)
	$MENDOC = GUICtrlCreateMenu("Procédures")
	$AdminDoc = GUICtrlCreateMenuItem("Admins", $MENDOC)
	$TechDoc = GUICtrlCreateMenuItem("Techs", $MENDOC)
	$MENCode = GUICtrlCreateMenu("Code Retour")
	$Gencode = GUICtrlCreateMenuItem("Génération du code", $MENCode)
	Global $LUser = GUICtrlCreateLabel("Nom de l'utilisateur", 40, 48, 92, 17)
	Global $IUser = GUICtrlCreateInput("", 160, 48, 169, 21)
	Global $LOrdi = GUICtrlCreateLabel("Nom de l'ordinateur", 40, 88, 95, 17)
	Global $IOrdi = GUICtrlCreateInput("", 160, 88, 169, 21)
	Global $BtValider = GUICtrlCreateButton("Valider", 64, 136, 97, 33)
	Global $BtAnnuler = GUICtrlCreateButton("Annuler", 208, 136, 97, 33)
	Global $Pic1 = GUICtrlCreatePic(@ScriptDir & "\cyber47.jpg", 64, 272, 233, 65)
	Global $Label1 = GUICtrlCreateLabel("- Désactivation du compte PC et utilisateur" & @CRLF & "- Modification du mot de passe utilisateur" & @CRLF & "- Suppresion des groupes VPN et activeSync", 24, 200, 386, 65)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###
	_AD_Open()

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $BtAnnuler
				FileDelete(@ScriptDir & "\cyber47.jpg")
				filedelete(@ScriptDir & "\tech.pdf")
				FileDelete(@ScriptDir & "\admin.pdf")
				Exit
			Case $BtValider
				Global $nomordi = GUICtrlRead($IOrdi)
				;************************************
				; Gestion Utilisateur
				;************************************
				$user = GUICtrlRead($IUser)
				$sUser = GUICtrlRead($IUser)
				$sUser = _AD_SamAccountNameToFQDN($sUser)
				$pass = _PassWordGenerate(12, 1)
				$iValuePass = _AD_SetPassword($sUser, $pass)
				;global $iValue = _AD_SetPassword($iuser, $pass)
				Global $logfile = "\\mars\dgs_si$\Reseau\0-Admin&Tech\Utilisateurs\Trace Cyber\" & $user & "_" & $nomordi & ".csv"
				If $iValuePass = 1 Then
					MsgBox(0, "mot de passe ", "Le mot de passe pour '" & $user & "' a été changé")

					FileOpen($logfile)
					FileWriteLine($logfile, @MDAY & "/" & @MON & "/" & @YEAR & "; Changement du mot de passe ")
					FileClose($logfile)
					Global $user = GUICtrlRead($IUser)
					;récupère le DN de l'utilisateur
					Global $sFQDN = _AD_SamAccountNameToFQDN($user)
					; Suppression des groupes VPN

					Global $iValueF5Util = _AD_RemoveUserFromGroup("CN=VPN_F5_UTIL,OU=Groupes Globaux,OU=Groupes,OU=ADMINISTRATION,DC=dptlg,DC=fr", $sFQDN)

					If $iValueF5Util = 1 Then
						MsgBox(64, "Gestion des groupes", "Supression du groupe VPN_F5_UTIL effectuée ")
						FileOpen($logfile)
						FileWriteLine($logfile, @MDAY & "/" & @MON & "/" & @YEAR & "; Supression du groupe VPN_F5_UTIL effectuée ")
						FileClose($logfile)
					EndIf

					Global $iValueF5DSIAN = _AD_RemoveUserFromGroup("CN=VPN_F5_DSIAN,OU=Groupes Globaux,OU=Groupes,OU=ADMINISTRATION,DC=dptlg,DC=fr", $sFQDN)

					If $iValueF5DSIAN = 1 Then
						MsgBox(64, "Gestion des groupes", "Supression du groupe VPN_F5_DSIAN effectuée ")
						FileOpen($logfile)
						FileWriteLine($logfile, @MDAY & "/" & @MON & "/" & @YEAR & "; Supression du groupe VPN_F5_DSIAN effectuée ")
						FileClose($logfile)
					EndIf

					Global $iValueF5VIP = _AD_RemoveUserFromGroup("CN=VPN_F5_VIP,OU=Groupes Globaux,OU=Groupes,OU=ADMINISTRATION,DC=dptlg,DC=fr", $sFQDN)

					If $iValueF5VIP = 1 Then
						MsgBox(64, "Gestion des groupes", "Supression du groupe VPN_F5_VIP effectuée ")
						FileOpen($logfile)
						FileWriteLine($logfile, @MDAY & "/" & @MON & "/" & @YEAR & "; Supression du groupe VPN_F5_VIP effectuée ")
						FileClose($logfile)
					EndIf

					Global $iValueF5sync = _AD_RemoveUserFromGroup("CN=F5_ACTIVESYNC_2016,OU=Groupes Globaux,OU=Groupes,OU=ADMINISTRATION,DC=dptlg,DC=fr", $sFQDN)

					If $iValueF5sync = 1 Then
						MsgBox(64, "Gestion des groupes", "Supression du groupe F5_ACTIVESYNC_2010 effectuée ")
						FileOpen($logfile)
						FileWriteLine($logfile, @MDAY & "/" & @MON & "/" & @YEAR & "; Supression du groupe F5_ACTIVESYNC_2010 effectuée ")
						FileClose($logfile)
					EndIf


					Global $idisableUser = _AD_DisableObject($user)
					If $idisableUser = 1 Then
						MsgBox(64, "Gestion de l'utilisateur", "L'utilisateur '" & $user & "' est désactivé")
						FileOpen($logfile)
						FileWriteLine($logfile, @MDAY & "/" & @MON & "/" & @YEAR & "; L'utilisateur '" & $user & "' est désactivé")
						FileClose($logfile)
					EndIf
					_GroupMembership()

				ElseIf @error = 1 Then
					MsgBox(0, "Erreur", "l'utilisateur, " & $sUser & ", n'existe pas ou n'a pas été saisi.")

				EndIf
				;***********************************
				; Gestion du PC
				;***********************************
				Global $nompc = GUICtrlRead($IOrdi) & "$"
				Global $nomordi = GUICtrlRead($IOrdi)
				Global $sFQDNordi = _AD_SamAccountNameToFQDN($nompc)

				Global $iValue = _AD_DisableObject($nompc)

				If $iValue = 1 Then
					MsgBox(64, "Gestion du PC", "le PC '" & $nompc & "' est désactivé")
					FileOpen($logfile)
					FileWriteLine($logfile, @MDAY & "/" & @MON & "/" & @YEAR & "; le PC '" & $nompc & "' est désactivé")
					FileClose($logfile)

					Global $iValuePC = _AD_RemoveUserFromGroup("CN=GPO_VPN_F5_O,OU=Groupes GPO Ordi,OU=Ordinateurs,OU=SEVEN,OU=CLIENTS CG47,DC=dptlg,DC=fr", $sFQDNordi)

					If $iValuePC = 1 Then
						MsgBox(64, "Gestion des groupes", "Suppression du groupe GPO_VPN_F5_O effectuée ")
						FileOpen($logfile)
						FileWriteLine($logfile, @MDAY & "/" & @MON & "/" & @YEAR & "; Suppression du groupe GPO_VPN_F5_O effectuée")
						FileClose($logfile)
					EndIf
				ElseIf @error = 1 Then
					MsgBox(0, "Erreur", "l'ordinateur, " & $nompc & ", n'existe pas ou n'a pas été saisi.")
				EndIf

			Case $Gencode
				$calcul = Number(@YDAY) * Number(@MDAY) * Number(@HOUR)
				MsgBox(0, 'Code Retour Cyber', "Le code est :  " & $calcul)

			Case $TechDoc
				ShellExecute(@ScriptDir & "\tech.pdf", "", "", "")

			Case $AdminDoc
				ShellExecute(@ScriptDir & "\admin.pdf", "", "", "")
		EndSwitch
	WEnd

	_AD_Close()
EndFunc   ;==>_fenetre


Func _PassWordGenerate($PassWordLen, $CharFlag)
	Local $PassWrd, $Global
	Local $Char = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
	Local $SpecChar = "&'_:;,.!@#$%*()-=+[]\/?<>{}^~¤"

	If $CharFlag = 1 Then
		$Array = StringSplit($Char & $SpecChar, "")
	Else
		$Array = StringSplit($Char, "")
	EndIf


	For $X = 1 To $PassWordLen
		$PassWrd &= $Array[Random(1, $Array[0], 1)]
	Next

	Return $PassWrd
EndFunc   ;==>_PassWordGenerate

Func _GroupMembership()
	Local $groupMembership = _AD_GetUserGroups($user)
	$frmList = GUICreate("liste des groupes de l'utilisateur", 556, 294, 192, 124)
	$listGroups = GUICtrlCreateList("", 8, 8, 537, 279)
	GUISetState(@SW_SHOW)

	For $i = 1 To $groupMembership[0]
		GUICtrlSetData($listGroups, $groupMembership[$i])
	Next
	While 2
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($frmList)
				ExitLoop
		EndSwitch
	WEnd
EndFunc   ;==>_GroupMembership

