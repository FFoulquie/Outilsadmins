#include <AD.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <FileConstants.au3>
#include <GuiComboBox.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\bin\Icones\Computer.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.11
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include "..\au3\ficconf.au3"

; paramétrage Général
Opt("GUICloseOnESC", 0) ; 0=pas de sortie par la touche ESC 1= Fermeture du programme via esc
;Opt("TrayAutoPause", 0) ;0=pas de pause, 1=pause
Opt("TrayIconHide", 1) ;0=montre, 1=masque l'icône de Notification

$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\GestPcAD", "Version", "REG_SZ", $version)

FileInstall(".\gestPcAd.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)
_AD_Open()

;Global $iReply = MsgBox(308, "Avertissement", "ATTENTION!!!! Vous allez gérer un ordinateur sur le domaine." & @CRLF & @CRLF & _
;		"Etes-vous sûr de vouloir continuer?")
;If $iReply <> 6 Then Exit
$GpoPC = _AD_GetObjectsInOU("", "(&(objectcategory=group)(cn=gpo*_O))", 2)
$ippc = IniRead($paramini, "PC", "Nom", "")


#Region ### START Koda GUI section ### Form=C:\CD47Users\frfoulqu\Documents\CD47\Scripts\Autoit\Outils Admin\include\Gestion PC\Form1_1.kxf
Global $Form1_1 = GUICreate("Gestion AD d'un PC", 540, 352, 415, 227)
GUICtrlCreateLabel("Saisir le nom du PC:", 8, 10, 127, 34)
Global $IObject = GUICtrlCreateInput($ippc, 137, 10, 163, 21)
Global $BOK = GUICtrlCreateButton("Supprimer le PC", 8, 72, 160, 33)
Global $BT_SuppPCAD = GUICtrlCreateButton("Réinstallation du PC", 8, 120, 160, 33)
Global $BT_CheckVer = GUICtrlCreateButton("Contrôle verrouillage", 384, 64, 129, 33)
Global $BT_Dever = GUICtrlCreateButton("Activer", 384, 112, 129, 33)
Global $Bt_Transfert = GUICtrlCreateButton("Transfert des groupes", 384, 192, 129, 33)
Global $COM_grADPc = GUICtrlCreateCombo("COM_grADPc", 8, 160, 361, 25, BitOR($CBS_DROPDOWN, $WS_VSCROLL))
GUICtrlSetData($COM_grADPc, "|" & _ArrayToString($GpoPC))
Global $BT_AdPctoGroup = GUICtrlCreateButton("Ajout du PC à un groupe", 8, 192, 153, 33)
Global $Bt_SupGrpPC = GUICtrlCreateButton("Lister les groupes du PC", 8, 272, 153, 33)
Global $BT_ValSuppGrp = GUICtrlCreateButton("Valider suppression", 176, 272, 113, 33)
Global $BT_Actualise = GUICtrlCreateButton("Actualiser", 320, 272, 97, 33)
;Global $Pic1 = GUICtrlCreatePic("", 240, 72, 73, 73)
$PLogo = GUICtrlCreateIcon($dlloa, 6, 240, 72, 48, 48)
;Global $BCancel = GUICtrlCreateButton("Quitter", 292, 72, 73, 33, $BS_DEFPUSHBUTTON)
Global $BCancel = GUICtrlCreateIcon($chemin_dossier_bin & "conf\OA.dll", 4, 452, 272, 32, 32)
GUICtrlSetTip($BCancel, "Quitter")

Global $Label1 = GUICtrlCreateLabel($version, 16, 320, 44, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $BCancel
			FileDelete($chemin_dossier_bin & "gestPcAd.au3")
			_AD_Close()
			Exit
		Case $BT_CheckVer
			Global $pc = GUICtrlRead($IObject)
			$pc &= "$"
			; Vérifier si l'objet PC est désactivé
			$rep = _AD_IsObjectDisabled($pc)
			Switch @error
				Case 0
					If $rep = 0 Then
						MsgBox(64, "État du PC", "Le PC est activé")
					Else
						MsgBox(64, "État du PC", "Le PC est désactivé")
					EndIf
				Case 1
					MsgBox(48, "Erreur", "PC introuvable")
				Case Else
					MsgBox(16, "Erreur", "Erreur AD")
			EndSwitch
			; Fermer la connexion Active Directory

		Case $BT_Dever
			Global $pc = GUICtrlRead($IObject)
			$pc &= "$"
			_FileWriteLog($Cheminlog, "**GestPC**  Deverrouillage du PC  : " & $pc)
			Global $iValue = _AD_EnableObject($pc)
			If $iValue = 1 Then
				MsgBox(64, "État du PC", "Le PC '" & $pc & "' est activé")
			ElseIf @error = 1 Then
				MsgBox(48, "État du PC", "Le PC '" & $pc & "' n'existe pas")
			Else
				MsgBox(16, "Erreur", "Erreur AD")
			EndIf

		Case $BOK
			Global $sObject = GUICtrlRead($IObject)
			_FileWriteLog($Cheminlog, "**GestPC**  Desactivation du PC  : " & $pc)
			deplace()
		Case $BT_SuppPCAD
			Global $sObject = GUICtrlRead($IObject)
			_FileWriteLog($Cheminlog, "**GestPC**  Suppression du PC  : " & $pc)
			supPC()
		Case $Bt_SupGrpPC

			$vUsername1 = GUICtrlRead($IObject)
			$groupToDel = _AD_GetUserGroups($vUsername1 & "$")
			$CoM_DelGRP = GUICtrlCreateCombo("Groupe à supprimer", 8, 240, 361, 25, BitOR($CBS_DROPDOWN, $WS_VSCROLL))
			GUICtrlSetFont(-1, 8, 400, 0, "Arial")
			GUICtrlSetData($CoM_DelGRP, "|" & _ArrayToString($groupToDel))
		Case $BT_ValSuppGrp
			$repcombodel = GUICtrlRead($CoM_DelGRP)
			$sObject = GUICtrlRead($IObject)

			; Vérification si le groupe à sortir est un groupe GPO (commence par GPO)
			; Extrait les premiers 6 caractères de la variable
			$sousChaine = StringLeft($repcombodel, 6)

			; Vérifie si la sous-chaîne est égale à "CN=GPO"
			If $sousChaine = "CN=GPO" Then
				Global $iValue = _AD_RemoveUserFromGroup($repcombodel, $sObject & "$")
				If $iValue = 1 Then
					MsgBox(64, "Suppression", "L'ordinateur '" & $sObject & "' est sorti du groupe '" & $repcombodel & "'")
					_FileWriteLog($Cheminlog, "**GestPC**  L'ordinateur '" & $sObject & "' est sorti du groupe '" & $repcombodel & "'")
				ElseIf @error = 1 Then
					MsgBox(64, "Suppression", "le groupe '" & $repcombodel & "' n'existe pas")
				ElseIf @error = 2 Then
					MsgBox(64, "Suppression", "L'ordinateur '" & $sObject & "' n'existe pas")
				ElseIf @error = 3 Then
					MsgBox(64, "Suppression", "L'ordinateur '" & $sObject & "' n'est pas membre du groupe '" & $repcombodel & "'")
				Else
					MsgBox(64, "Suppression", "Erreur '" & @error & "' de Active Directory")
				EndIf

			Else
				MsgBox(0, "Avertissement", "Vous n'êtes pas autoriser à sortir l'ordinateur de ce groupe")
			EndIf

		Case $BT_Actualise
			$vUsername1 = ""
			$groupToDel = ""
			$CoM_DelGRP = GUICtrlCreateCombo("Groupe à supprimer", 8, 240, 361, 25, BitOR($CBS_DROPDOWN, $WS_VSCROLL))
			_GUICtrlComboBox_Destroy($CoM_DelGRP)
			$vUsername1 = GUICtrlRead($IObject)
			$groupToDel = _AD_GetUserGroups($vUsername1 & "$")
			$CoM_DelGRP = GUICtrlCreateCombo("Suppression", 8, 240, 361, 25, BitOR($CBS_DROPDOWN, $WS_VSCROLL))
			GUICtrlSetFont(-1, 8, 400, 0, "Arial")
			GUICtrlSetData($CoM_DelGRP, "|" & _ArrayToString($groupToDel))
		Case $BT_AdPctoGroup
			$repcombo = GUICtrlRead($COM_grADPc)

			$sObject = GUICtrlRead($IObject)
			Global $iValue = _AD_AddUserToGroup($repcombo, $sObject & "$")
			If $iValue = 1 Then
				MsgBox(64, "Gestion des Groupes", "L'ordinateur '" & $sObject & "' a été ajouté au groupe '" & $repcombo & "'")
				_FileWriteLog($Cheminlog, "**GestPC**  L'ordinateur '" & $sObject & "' est ajoute du groupe '" & $repcombo & "'")
			ElseIf @error = 1 Then
				MsgBox(64, "Gestion des Groupes", "Le groupe '" & $repcombo & "' n'existe pas")
			ElseIf @error = 2 Then
				MsgBox(64, "Gestion des Groupes", "L'ordinateur '" & $sObject & "' n'existe pas")
			ElseIf @error = 3 Then
				MsgBox(64, "Gestion des Groupes", "L'ordinateur '" & $sObject & "' est déjà membre du groupe '" & $repcombo & "'")
			Else
				MsgBox(64, "Gestion des Groupes", "Erreur '" & @error & "' de l'Active Directory")
			EndIf
		Case $Bt_Transfert

			Global $aProperties[1][2]
			;Recherche l'attribut membre de sur l'AD
			$aProperties = _AD_GetObjectProperties($ippc & "$", "memberof")

			$NewPC = InputBox("Nouveau PC", "Merci de saisir le nom du nouveau PC", "P000-T")

			Select
				Case @error = 0 ;OK -
					;pour chaques lignes trouvées dans le GetObjectProperties je l'ajoute dans le nouveau PC
					For $i = 1 To UBound($aProperties) - 1
						Global $iValue = _AD_AddUserToGroup($aProperties[$i][1], $NewPC & "$")

						If @error = 2 Then
							MsgBox(64, "Gestion des Groupes", "L'ordinateur '" & $NewPC & "' n'existe pas")
						EndIf
					Next
					MsgBox(64, "Gestion des Groupes", "Transfert terminé")
					$aNewProperties = _AD_GetObjectProperties($NewPC & "$", "memberof")
					_ArrayDisplay($aNewProperties, "Groupes AD pour l'ordinateur '" & $NewPC & "'", "1:|2:", 64, Default, "Membre de|Nom du groupe")
				Case @error = 1 ;Annulé
					MsgBox(64, "Gestion des Groupes", "Transfert Annulé")
			EndSelect
	EndSwitch
WEnd

Func deplace()
	Global $sTargetOU = "OU=A_Supp,OU=Ordinateurs,OU=W10,OU=CLIENTS CG47,DC=dptlg,DC=fr"
	_AD_DisableObject($sObject & "$")
	Global $iValue = _AD_MoveObject($sTargetOU, $sObject & "$")
	If $iValue = 1 Then
		MsgBox(64, "Sortie", "Le PC '" & $sObject & "' Est correctement sorti")
	ElseIf @error = 1 Then
		MsgBox(64, "Sortie", "Le PC '" & $sObject & "' n'existe pas")
	Else
		MsgBox(64, "Sortie", "Erreur '" & @error & "' de l'Active Directory")
	EndIf

EndFunc   ;==>deplace

Func supPC()
	; Check if object exists
	If Not _AD_ObjectExists() Then Exit MsgBox(16, "Réinstallation", "Le PC'" & $sObject & "' n'existe pas")
	; Delete object
	Global $iValue = _AD_DeleteObject($sObject & "$", _AD_GetObjectClass($sObject & "$"))
	If $iValue = 1 Then
		MsgBox(64, "Réinstallation", "Le PC'" & $sObject & "' a été supprimé et peut être réinstallé")
	ElseIf @error = 1 Then
		MsgBox(16, "Réinstallation", "Le PC'" & $sObject & "'n'existe pas")
	Else
		MsgBox(16, "Réinstallation", "Erreur '" & @error & "' de l'Active Directory")
	EndIf

	; Close Connection to the Active Directory

EndFunc   ;==>supPC

