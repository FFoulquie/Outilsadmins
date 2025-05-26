#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\User Accounts.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.2
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <AD.au3>
#include "..\au3\ficconf.au3"

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]

#Region ### START Koda GUI section ### Form=
Global $Form1_1 = GUICreate("Gestion du groupe ADMIN_LOCAUX_PC     V : " & $version, 515, 125, 525, 432)
GUICtrlCreateLabel("Nom  d'ouverture de session de l'utilisateur", 8, 10, 231, 17)
Global $User = GUICtrlCreateInput(@UserName, 241, 8, 259, 21)
Global $bt_Ajout = GUICtrlCreateButton("Ajouter l'utilisateur au groupe", 208, 64, 162, 33)
;Global $Bt_Cancel = GUICtrlCreateButton("Annuler", 412, 64, 73, 33, $BS_DEFPUSHBUTTON)
Global $Bt_Supp = GUICtrlCreateButton("Supprimer l'utilisateur du groupe", 32, 64, 162, 33)
	$Bt_Cancel = GUICtrlCreateIcon($dlloa, 4, 412, 64, 32, 32)
	GUICtrlSetTip($Bt_Cancel, "Quitter")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\GrpAdminslocaux", "Version", "REG_SZ", $version)

_AD_Open()

Global $sGroup = _AD_SamAccountNameToFQDN("CN=ADMIN_LOCAUX_PC,OU=Groupes Globaux,OU=Groupes,OU=ADMINISTRATION,DC=dptlg,DC=fr")


While 1
 $nMsg = GUIGetMsg()
 Switch $nMsg
  Case $GUI_EVENT_CLOSE, $Bt_Cancel
   Exit
	Case $bt_Ajout
			Global $sObject = _AD_SamAccountNameToFQDN(GUICtrlRead($User))
			AjoutUser()
			ExitLoop
	Case $Bt_Supp
			Global $sObject = _AD_SamAccountNameToFQDN(GUICtrlRead($User))
			SuppUser()
			ExitLoop
 EndSwitch
WEnd


Func AjoutUser()
	Global $iValue = _AD_AddUserToGroup($sGroup, $sObject)
If $iValue = 1 Then
	MsgBox(64, "Gestion Admin locaux", "L'utilisateur '" & $sObject & "' a été ajouté au groupe '" & $sGroup & "'")
ElseIf @error = 1 Then
	MsgBox(64, "Gestion Admin locaux", "Le groupe '" & $sGroup & "' n'existe pas")
ElseIf @error = 2 Then
	MsgBox(64, "Gestion Admin locaux", "L'utilisateur '" & $sObject & "' n'existe pas")
ElseIf @error = 3 Then
	MsgBox(64, "Gestion Admin locaux", "L'utilisateur '" & $sObject & "' est déjà membre du groupe '" & $sGroup & "'")
Else
	MsgBox(64, "Gestion Admin locaux", "Retour erreur '" & @error & "' de l' Active Directory")
EndIf
EndFunc

Func SuppUser()
	Global $iValue = _AD_RemoveUserFromGroup($sGroup, $sObject)
If $iValue = 1 Then
	MsgBox(64, "Gestion Admin locaux", "L'utilisateur '" & $sObject & "' a été supprimé du groupe '" & $sGroup & "'")
ElseIf @error = 1 Then
	MsgBox(64, "Gestion Admin locaux", "Le groupe '" & $sGroup & "' n'existe pas")
ElseIf @error = 2 Then
	MsgBox(64, "Gestion Admin locaux", "L'utilisateur '" & $sObject & "' n'existe pas")
ElseIf @error = 3 Then
	MsgBox(64, "Gestion Admin locaux", "L'utilisateur '" & $sObject & "' ne fait pas parti du groupe '" & $sGroup & "'")
Else
	MsgBox(64, "Gestion Admin locaux", "Retour erreur '" & @error & "' de l' Active Directory")
EndIf
EndFunc


_AD_Close()