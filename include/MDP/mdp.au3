#include <AD.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GuiComboBox.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>


#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\User Accounts.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.9
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_LegalCopyright=@FFOULQUIE
#AutoIt3Wrapper_Res_Language=1036
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include "..\au3\ficconf.au3"

Opt("TrayIconHide", 1) ; pas d'icone dans la zone de notification

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Mdp", "Version", "REG_SZ", $version)

_AD_OPEN()

$GpoUtil = _AD_GetObjectsInOU("", "(&(objectcategory=group)(cn=gpo*_U))", 2)

Global $groupToDel
Global $Com_DelGroup

#Region ### START Koda GUI section ### Form=CC:\CD47Users\frfoulqu\Documents\CD47\Scripts\Autoit\Outils Admin\include\MDP\mdp.kxf
Global $Form1 = GUICreate("Gestion des utilisateurs sur l'AD   V : " & $version, 600, 400)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $IUser = GUICtrlCreateInput("", 24, 40, 153, 21)
Global $Label1 = GUICtrlCreateLabel("Nom de l'utilisateur", 24, 16, 92, 17)
$Bt_SSI = GUICtrlCreateButton("Fiche SSI", 200, 40, 105, 21, $BS_MULTILINE)

Global $BtQuitter = GUICtrlCreateButton("   Quitter", 400, 30, 100, 40)
GUICtrlSetImage($BtQuitter, $chemin_dossier_bin & "conf\OA.dll", 4, 1)
GUICtrlSetTip($BtQuitter, "Quitter")

Global $Tab1 = GUICtrlCreateTab(24, 72, 521, 241)

Global $TabSheetconn = GUICtrlCreateTabItem("Connexion")
$Bt_opensession = GUICtrlCreateButton("Date ouverture session", 48, 150, 105, 41, $BS_MULTILINE)
GUICtrlSetTip($Bt_opensession, "Affiche la dernière ouverture de session de l'utilisateur")

Global $TabSheet1 = GUICtrlCreateTabItem("Mot de passe")
Global $Label2 = GUICtrlCreateLabel("Mot de Passe", 33, 108, 69, 17)
Global $IMDP = GUICtrlCreateInput("", 33, 132, 153, 21)
Global $btnCheckPWInfo = GUICtrlCreateButton("Vérification MDP", 212, 132, 177, 24)
Global $NB_BadPwd = GUICtrlCreateButton("Nombre mauvais mot de passes", 212, 169, 177, 24)
Global $bt_datelastbadpwd = GUICtrlCreateButton("Date dernier mauvais mot de passe", 212, 206, 177, 24)
Global $BtForceMdp = GUICtrlCreateButton("Forcer le changement", 212, 243, 177, 24)
;Global $IMDP = GUICtrlCreateInput("", 33, 132, 153, 21, BitOR($GUI_SS_DEFAULT_INPUT, $ES_PASSWORD))
Global $Group3 = GUICtrlCreateGroup("Génération", 34, 156, 153, 145)
Global $bttech = GUICtrlCreateButton("Technicien" & @CRLF & "pas de forcage", 50, 212, 105, 40, $BS_MULTILINE)
Global $Btuser = GUICtrlCreateButton("Utilisateur" & @CRLF & "Force le changement", 50, 252, 105, 40, $BS_MULTILINE)
Global $Label5 = GUICtrlCreateLabel("Type de mot de passe", 58, 180, 109, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)


Global $TabSheet2 = GUICtrlCreateTabItem("Déverrouillage")
Global $btnCheck = GUICtrlCreateButton("Vérification", 32, 112, 177, 25)
Global $btnUnlock = GUICtrlCreateButton("Dévérrouiller", 32, 144, 177, 41)

Global $TabSheet3 = GUICtrlCreateTabItem("Groupes")
Global $btnCheckgrp = GUICtrlCreateButton("Groupes membres de", 32, 112, 177, 24)
Global $Bt_AduserToGroup = GUICtrlCreateButton("Ajouter à un groupe", 28, 173, 177, 24)
GUICtrlSetFont(-1, 8, 400, 0, "Arial")
Global $Bt_DeluserToGroup = GUICtrlCreateButton("Lister les groupes utilisateur", 28, 245, 177, 24)
GUICtrlSetFont(-1, 8, 400, 0, "Arial")
Global $Com_AddGroup = GUICtrlCreateCombo("Ajout", 28, 144, 305, 25, BitOR($CBS_DROPDOWN, $WS_VSCROLL))
GUICtrlSetFont(-1, 8, 400, 0, "Arial")
GUICtrlSetData($Com_AddGroup, "|" & _ArrayToString($GpoUtil))
Global $BT_ValSupp = GUICtrlCreateButton("Supprimer", 224, 245, 81, 25)
GUICtrlSetFont(-1, 8, 400, 0, "Arial")
Global $BT_Actualiser = GUICtrlCreateButton("Actualiser", 224, 285, 81, 25)
GUICtrlSetFont(-1, 8, 400, 0, "Arial")
$Com_DelGroup = GUICtrlCreateCombo("Suppression", 28, 208, 305, 25, BitOR($CBS_DROPDOWN, $WS_VSCROLL))
GUICtrlSetFont(-1, 8, 400, 0, "Arial")







Global $TabSheet5 = GUICtrlCreateTabItem("Divers")
Global $BT_SID = GUICtrlCreateButton("Récupérer SID", 48, 120, 105, 25)

GUICtrlCreateTabItem("")

Global $Edit1 = GUICtrlCreateEdit("", 24, 315, 521, 50, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN))
GUICtrlSetData(-1, "")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


While 1

	$Msg = GUIGetMsg()
	Select
		Case $Msg = $BT_SID
			GUICtrlSetData($Edit1, "")
			GUICtrlSetData($Edit1, SID(), 1)
		Case $Msg = $BtQuitter
			_AD_Close()
			Exit
		Case $Msg = $GUI_EVENT_CLOSE
			_AD_Close()
			Exit
		Case $Msg = $Btuser
			$vUsername1 = GUICtrlRead($IUser)
			$vPassword = GUICtrlRead($IMDP)
			Global $vForce = 1
			_pwReset()

			$pwValue = 0
			$vUsername1 = ""
			$vPassword = ""

			GUICtrlSetData($IMDP, "")
		Case $Msg = $Bt_opensession
			$user = GUICtrlRead($IUser)
			;_AD_Open()
			$aListProp = _AD_GetObjectProperties($user, "lastlogon")
			If @error Then
				MsgBox(16, "Erreur Active Directory", "Utilisateur inconnu")
			Else
				GUICtrlSetData($Edit1, "")
				GUICtrlSetData($Edit1, "Dernière connexion de " & $user & " : " & @CRLF & WMIDateStringToDate($aListProp[1][1]), 1)
				;MsgBox(0, "", "Dernière connexion de " & $user & " : " & @CRLF & WMIDateStringToDate($aListProp[1][1]))

			EndIf
			;_AD_Close()

		Case $Msg = $Bt_AduserToGroup
			$repcombo = GUICtrlRead($Com_AddGroup)

			$sObject = GUICtrlRead($IUser)
			_FileWriteLog($Cheminlog, "**MDP**  Ajout de groupes GPO pour  : " & $sObject)
			Global $iValue = _AD_AddUserToGroup($repcombo, $sObject)
			If $iValue = 1 Then
				MsgBox(64, "Gestion des Groupes", "L'utilisateur '" & $sObject & "' a été ajouté au groupe '" & $repcombo & "'")
			ElseIf @error = 1 Then
				MsgBox(64, "Gestion des Groupes", "Le groupe '" & $repcombo & "' n'existe pas")
			ElseIf @error = 2 Then
				MsgBox(64, "Gestion des Groupes", "L'utilisateur '" & $sObject & "' n'existe pas")
			ElseIf @error = 3 Then
				MsgBox(64, "Gestion des Groupes", "L'utilisateur '" & $sObject & "' est déjà membre du groupe '" & $repcombo & "'")
			Else
				MsgBox(64, "Gestion des Groupes", "Erreur '" & @error & "' de l'Active Directory")
			EndIf

		Case $Msg = $NB_BadPwd
			$sObject = GUICtrlRead($IUser)
			$aListProp = _AD_GetObjectProperties($sObject, "badPwdCount")
			If @error Then
				MsgBox(16, "Erreur Active Directory", "Utilisateur inconnu")
			Else
				GUICtrlSetData($Edit1, "")
				;MsgBox(0, "", "Nombre de mauvais mot de passe " & $sObject & " : " & @CRLF & $aListProp[1][1])
				GUICtrlSetData($Edit1, "Nombre de mauvais mot de passe : " & @CRLF & $aListProp[1][1], 1)
			EndIf
		Case $Msg = $bt_datelastbadpwd
			$sObject = GUICtrlRead($IUser)
			$aListProp = _AD_GetObjectProperties($sObject, "badPasswordTime")
			If @error Then
				MsgBox(16, "Erreur Active Directory", "Utilisateur inconnu")
			Else
				;MsgBox(0, "", "Date du dernier mauvais mot de passe de  " & $sObject & " : " & @CRLF & $aListProp[1][1])
				GUICtrlSetData($Edit1, "")
				GUICtrlSetData($Edit1, "Date dernier mauvais mot de passe : " & @CRLF & WMIDateStringToDate($aListProp[1][1]), 1)
			EndIf
		Case $Msg = $Bt_DeluserToGroup

			$vUsername1 = GUICtrlRead($IUser)
			$groupToDel = _AD_GetUserGroups($vUsername1)

			GUICtrlSetData($Com_DelGroup, "|" & _ArrayToString($groupToDel))
		Case $Msg = $BT_ValSupp
			$repcombodel = GUICtrlRead($Com_DelGroup)
			$sObject = GUICtrlRead($IUser)
			_FileWriteLog($Cheminlog, "**MDP**  Suppression de GPO pour : " & $vUsername1)
			; Vérification si le groupe à sortir est un groupe GPO (commence par GPO)
			; Extrait les premiers 6 caractères de la variable
			$sousChaine = StringLeft($repcombodel, 6)

			; Vérifie si la sous-chaîne est égale à "CN=GPO"
			If $sousChaine = "CN=GPO" Then
				Global $iValue = _AD_RemoveUserFromGroup($repcombodel, $sObject)
				If $iValue = 1 Then
					MsgBox(64, "Suppression", "L'utilisateur '" & $sObject & "' est sorti du groupe '" & $repcombodel & "'")
				ElseIf @error = 1 Then
					MsgBox(64, "Suppression", "le groupe '" & $repcombodel & "' n'existe pas")
				ElseIf @error = 2 Then
					MsgBox(64, "Suppression", "L'utilisateur '" & $sObject & "' n'existe pas")
				ElseIf @error = 3 Then
					MsgBox(64, "Suppression", "L'utilisateur '" & $sObject & "' n'est pas membre du groupe '" & $repcombodel & "'")
				Else
					MsgBox(64, "Suppression", "Return code '" & @error & "' from Active Directory")
				EndIf
			Else
				MsgBox(0, "Avertissement", "Vous n'êtes pas autoriser à sortir l'utilisateur de ce groupe")
			EndIf

		Case $Msg = $BtForceMdp
			$vUsername1 = GUICtrlRead($IUser)
			$iValue = _AD_SetPasswordExpire($vUsername1, "0")
			If $iValue = 1 Then
				MsgBox(64, "Réussite", "L'utilisateur '" & $vUsername1 & "' devra changer son mot de passe à sa prochaine ouverture de session")
			ElseIf @error = 1 Then
				MsgBox(64, "Erreur", "L'utilisateur '" & $vUsername1 & "' n'existe pas")
			Else
				MsgBox(64, "Erreur", "Code '" & @error)
			EndIf

		Case $Msg = $bttech
			$vUsername1 = GUICtrlRead($IUser)
			$vPassword = GUICtrlRead($IMDP)
			Global $vForce = 0
			_pwReset()

			$pwValue = 0
			$vUsername1 = ""
			$vPassword = ""

			GUICtrlSetData($IMDP, "")
		Case $Msg = $btnCheck
			;Read data that has been inputted from the user
			$vUsername2 = GUICtrlRead($IUser)

			_LockOutStatus()

		Case $Msg = $btnUnlock
			$vUsername2 = GUICtrlRead($IUser)
			_FileWriteLog($Cheminlog, "**MDP**  Déverrouillage de  : " & $vUsername2)
			_UnlockAcct()

		Case $Msg = $btnCheckgrp
			$vUsername3 = GUICtrlRead($IUser)
			GUISetState(@SW_DISABLE, $Form1)
			_GroupMembership()
			GUISetState(@SW_ENABLE, $Form1)
			WinActivate($Form1)


		Case $Msg = $btnCheckPWInfo
			$vUsername4 = GUICtrlRead($IUser)
			_PasswordInfo()

		Case $Msg = $BT_Actualiser
			$vUsername1 = ""
			$groupToDel = ""
			_GUICtrlComboBox_Destroy($Com_DelGroup)
			$vUsername1 = GUICtrlRead($IUser)
			$groupToDel = _AD_GetUserGroups($vUsername1)
			$Com_DelGroup = GUICtrlCreateCombo("Suppression", 28, 208, 305, 25, BitOR($CBS_DROPDOWN, $WS_VSCROLL))
			GUICtrlSetFont(-1, 8, 400, 0, "Arial")
			GUICtrlSetData($Com_DelGroup, "|" & _ArrayToString($groupToDel))
		Case $Msg = $Bt_SSI

			$vUsername1 = GUICtrlRead($IUser)
			_FileWriteLog($Cheminlog, "**MDP**  lancement SSI pour  : " & $vUsername1)
			Run("C:\Progwindev\ssi\ssi.exe")

			WinWaitActive("[Title:Systèmes d'information]")
			WinActivate("[Title:Systèmes d'information]")
			ControlClick("[Title:Systèmes d'information]", "", "[CLASS:Button; INSTANCE:10]")
			WinWaitActive("[Title:Agents]")
			MouseClick("left", 1159, 111)
			Send($vUsername1)
			Send("{ENTER}")
			MouseClick($MOUSE_CLICK_LEFT, 941, 490, 2)
	EndSelect
WEnd

Func _UnlockAcct()
	Local $lockout = _AD_UnlockObject($vUsername2)
	If $lockout = 1 Then
		MsgBox(0, "OK", "Le compte utilisateur " & $vUsername2 & " a été déverrouillé.")
	ElseIf @error = 1 Then
		MsgBox(0, "Erreur", "l'utilisateur, " & $vUsername2 & ", n'existe pas.")
	Else
		MsgBox(64, "Active Directory Erreur", " code '" & @error & "' de l'AD")
	EndIf
EndFunc   ;==>_UnlockAcct

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


;Performs the function that resets the user password
Func _pwReset()
	$IMDPReset = _AD_SetPassword($vUsername1, $vPassword, $vForce)
	If $IMDPReset = 1 Then
		MsgBox(0, "OK", "Le changement de mot de passe " & $vPassword & " pour " & $vUsername1 & " a réussi.")
	ElseIf @error = 1 Then
		MsgBox(0, "Erreur", "L'utilisateur, " & $vUsername1 & ", n'existe pas.")
	Else
		MsgBox(64, "Active Directory Erreur", "code '" & @error & "' de l'AD")
	EndIf
EndFunc   ;==>_pwReset

Func _PasswordInfo()
	Local $pwInfo = _AD_GetPasswordInfo($vUsername4)
	;_ArrayDisplay($pwinfo, "Groupes AD pour l'ordinateur ")
	;MsgBox(0, "Information sur le mot de passe", "Dernier changement: " & $pwInfo[8] & @CRLF & "Expiration: " & $pwInfo[9] & @CRLF)
	GUICtrlSetData($Edit1, "")
	GUICtrlSetData($Edit1, "Dernier changement: " & WMIDateStringToDate($pwInfo[8]) & @CRLF & "Expiration: " & $pwInfo[9] & @CRLF, 0)
EndFunc   ;==>_PasswordInfo

Func _LockOutStatus()
	Local $LockoutStatus = _AD_IsObjectLocked($vUsername2)
	If $LockoutStatus = 1 Then
		GUICtrlSetData($Edit1, "")
		GUICtrlSetData($Edit1, "Le compte est verrouillé", 0)
		;MsgBox(8208, "Compte verrouillé", "Le compte est verrouillé")
		;$lblLockoutStatus = GUICtrlCreateLabel("Verrouillé", 120, 85, 81, 20)
		;GUICtrlSetFont(-1, 10, 800, 0, "Arial")
		;GUICtrlSetColor(-1, 0xff0000)
		;GUICtrlSetBkColor(-1, 0xffffff)
	Else
		GUICtrlSetData($Edit1, "")
		GUICtrlSetData($Edit1, "Le compte est déverrouillé", 0)
		;MsgBox(8256, "Compte deverrouillé", "Le compte est déverrouillé")
		;$lblLockoutStatus = GUICtrlCreateLabel("Déverrouillé", 120, 85, 81, 20)
		;GUICtrlSetFont(-1, 10, 800, 0, "Arial")
		;GUICtrlSetColor(-1, 0x33cc00)
		;GUICtrlSetBkColor(-1, 0xffffff)
	EndIf
	#comments-start
		$checkcontrol = _AD_GetObjectProperties($vUsername2, "useraccountControl")
		$aListProp = StringRegExp("*ACCOUNTDISABLE*" , $checkcontrol[1][1],3)

		if $aListProp = 1 Then
		$lblLockoutStatus = GUICtrlCreateLabel("Actif", 60, 85, 41, 20)
		GUICtrlSetFont(-1, 10, 800, 0, "Arial")
		GUICtrlSetColor(-1, 0x33cc00)
		GUICtrlSetBkColor(-1, 0xffffff)
		Else
		$lblLockoutStatus = GUICtrlCreateLabel("Inactif", 60, 85, 41, 20)
		GUICtrlSetFont(-1, 10, 800, 0, "Arial")
		GUICtrlSetColor(-1, 0xff0000)
		GUICtrlSetBkColor(-1, 0xffffff)
		EndIf
	#comments-end
EndFunc   ;==>_LockOutStatus


Func _GroupMembership()
	Local $groupMembership = _AD_GetUserGroups($vUsername3)

	$frmList = GUICreate("membre de", 356, 294, 192, 124)
	$listGroups = GUICtrlCreateList("", 8, 8, 337, 279)
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

Func SID()
	$sObject = GUICtrlRead($IUser)
	; Chemin du fichier texte
	Local $filePath = "c:\temp\sid.txt"
	RunWait(@ComSpec & " /c wmic useraccount where name='" & $sObject & "' get sid > c:\temp\sid.txt ")

	; Ouvrir le fichier en mode lecture
	Local $file = FileOpen($filePath, 0)

	; Vérifier si le fichier a été ouvert avec succès
	If $file = -1 Then
		MsgBox(0, "Erreur", "Impossible d'ouvrir le fichier.")
		Exit
	EndIf

	; Lire la deuxième ligne
	Local $line
	For $i = 1 To 2
		$line = FileReadLine($file)
		If @error Then ; Atteint la fin du fichier avant la deuxième ligne
			MsgBox(0, "Attention", "Le fichier n'a que " & $i - 1 & " ligne(s).")
			ExitLoop
		EndIf
	Next

	; Fermer le fichier
	FileClose($file)

	; Afficher le contenu de la deuxième ligne
	;MsgBox(0, "SID", "Le SID est : " & $line)
	Return $line
	FileDelete($filePath)
EndFunc   ;==>SID



Func WMIDateStringToDate($dtmDate)
	Return (StringMid($dtmDate, 9, 2) & "/" & _
			StringMid($dtmDate, 6, 2) & "/" & StringLeft($dtmDate, 4) _
			 & " " & StringMid($dtmDate, 12, 2) & ":" & StringMid($dtmDate, 15, 2))
	;& ":" & StringMid($dtmDate,13, 2))
EndFunc   ;==>WMIDateStringToDate
