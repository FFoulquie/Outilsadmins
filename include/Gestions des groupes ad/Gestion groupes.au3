#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\AD 2.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.2
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <AD.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

Global $chemin_dossier_bin = "C:\Program Files\DSIAN\outilsAdmin\Bin\"
_AD_Open()

FileInstall(".\Gestion Groupes.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\GestGroupe", "Version", "REG_SZ", $version)

$GpoPC = _AD_GetObjectsInOU("", "(&(objectcategory=group)(cn=gpo*_O))", 2)
$GpoUser = _AD_GetObjectsInOU("", "(&(objectcategory=group)(cn=gpo*_U))", 2)

#Region ### START Koda GUI section ### Form=C:\CD47Users\frfoulqu\Documents\CD47\Scripts\Autoit\Outils Admin\include\Gestions des groupes ad\Form1.kxf
Global $Form1 = GUICreate("Suivi des groupes AD", 556, 270, 574, 239)
GUISetBkColor(0xFFFFFF)
Global $CBordi = GUICtrlCreateCombo("", 24, 120, 201, 25, BitOR($CBS_DROPDOWN, $WS_VSCROLL))
GUICtrlSetData($CBordi, "|" & _ArrayToString($GpoPC))
Global $CBUser = GUICtrlCreateCombo("", 332, 120, 201, 25, BitOR($CBS_DROPDOWN, $WS_VSCROLL))
GUICtrlSetData($CBUser, "|" & _ArrayToString($GpoUser))
Global $BtListPC = GUICtrlCreateButton("Lister", 40, 157, 105, 25)
Global $BtListUser = GUICtrlCreateButton("Lister", 356, 157, 105, 25)
Global $Label1 = GUICtrlCreateLabel("Groupes Ordinateurs", 48, 88, 114, 19)
GUICtrlSetFont(-1, 9, 400, 0, "@Malgun Gothic")
Global $Label2 = GUICtrlCreateLabel("Groupes Utilisateurs", 356, 88, 110, 19)
GUICtrlSetFont(-1, 9, 400, 0, "@Malgun Gothic")
;Global $BtQuit = GUICtrlCreateButton("Quitter", 232, 232, 89, 25)
Global $BtQuit = GUICtrlCreateIcon($chemin_dossier_bin & "conf\OA.dll", 4, 232, 232, 32, 32)
GUICtrlSetTip($BtQuit, "Quitter")
Global $Label3 = GUICtrlCreateLabel("Lister les membres des groupes AD", 144, 16, 272, 25)
GUICtrlSetFont(-1, 12, 400, 0, "@Malgun Gothic")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $BTQUIT
			FileDelete($chemin_dossier_bin & "Gestion Groupes.au3")
			_AD_Close()
			Exit
		case $BtListPC
			$repcombo = GUICtrlRead($CBordi)
			;msgbox (0,"",$repcombo)
			$aMemberOf = _AD_GetGroupMembers($repcombo)
			;$aMemberOf = _AD_SamAccountNameToFQDN($aMemberOffqdn)
_ArraySort($aMemberOf, 0, 1)
_ArrayDisplay($aMemberOf, "Membres du groupe :  '" & $repcombo & "'")
		case $BtListUser
		$repcombo = GUICtrlRead($CBUser)
			;msgbox (0,"",$repcombo)
			$aMemberOf = _AD_GetGroupMembers($repcombo)
			;$aMemberOf = _AD_SamAccountNameToFQDN($aMemberOffqdn)
_ArraySort($aMemberOf, 0, 1)
_ArrayDisplay($aMemberOf, "Active Directory Functions - Example 1 - Membership for group '" & $repcombo & "'")
	EndSwitch
WEnd




