#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=.\include\bin\icones\system_utilities.ico
#AutoIt3Wrapper_Outfile=outiladmins.exe
#AutoIt3Wrapper_Outfile_x64=outiladmins.exe
#AutoIt3Wrapper_Res_Fileversion=4.0.0.9
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_ProductName=Outils Admin
#AutoIt3Wrapper_Res_ProductVersion=4.0
#AutoIt3Wrapper_Res_LegalCopyright=@FFoulquié
#AutoIt3Wrapper_Res_Language=1036
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <AD.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <Date.au3>
#include <Excel.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <Misc.au3>
#include <MsgBoxConstants.au3>
#include <StaticConstants.au3>
#include <StructureConstants.au3>
#include <WindowsConstants.au3>
#include <Word.au3>
;#include <Menu.au3>

#include ".\include\au3\ficconf.au3"
#include ".\include\au3\authentif.au3"
#include ".\include\au3\gui.au3"
#include ".\include\au3\fonctions.au3"
#include ".\include\au3\editini.au3"
#include ".\include\au3\VerifSauvegarde.au3"
#include ".\include\au3\RaccourciClaviers.au3"

; paramétrage Général
Opt("GUICloseOnESC", 0) ; 0=pas de sortie par la touche ESC 1= Fermeture du programme via esc
;Opt("TrayAutoPause", 0) ;0=pas de pause, 1=pause
Opt("TrayIconHide", 1) ;0=montre, 1=masque l'icône de Notification


;Global $ini_configure = $configini
Global $hListView, $ArrayCrypt[10], $sections, $id, $rep, $a, $c, $iIDFrom, $addPc, $Labelmeteo, $Labelmeteo5
Global $version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
_Connect()

Func _fenetre()
	;FileDelete($Cheminlog)


	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins", "Version", "REG_SZ", $version)

	#Region Mise à jour globale

	$verattendue = IniRead($Versionini, "outilsadmins", "Version", "")
	_FileWriteLog($Cheminlog, "**Lanceur**  attendu  " & $verattendue)

	If Not FileExists($chemin_dossier_bin & "icones\MAJ.jpg") Then
		FileInstall(".\include\bin\icones\MAJ.jpg", $chemin_dossier_bin & "icones\", $FC_OVERWRITE)
	EndIf

	;SplashImageOn("Merci de patienter", $chemin_dossier_bin & "icones\MAJ.jpg", 428, 145)

	DirCreate($chemin_dossier_bin)
	DirCreate($chemin_dossier_bin & "icones")
	DirCreate($chemin_dossier_bin & "doc")
	DirCreate($chemin_dossier_bin & "conf")


	DirRemove("c:\temp\rapport", 1)
	DirCreate("c:\temp\rapport\")


	_FileWriteLog($Cheminlog, "Démarrage OA par : " & $sUsername)
	FileInstall("C:\CD47Users\frfoulqu\Documents\CD47\Scripts\Autoit\DLL\user.dll", $chemin_dossier_bin & "conf\", $FC_OVERWRITE)
	_FileCopy("O:\Techs\outiladmins\bin\conf\Param.Flq", $chemin_dossier_bin & "conf\")
	_FileWriteLog($Cheminlog, "Copie param O:\Techs\outiladmins\bin\conf\Param.Flq sous ", $chemin_dossier_bin & "conf\")
	_FileCopy("O:\Techs\outiladmins\bin\conf\version.Flq", $chemin_dossier_bin & "conf\")
	_FileWriteLog($Cheminlog, "Copie version O:\Techs\outiladmins\bin\conf\version.Flq sous " & $chemin_dossier_bin & "conf\")
	_FileCopy("O:\Techs\outiladmins\bin\conf\liens.Flq", $chemin_dossier_bin & "conf\")
	_FileWriteLog($Cheminlog, "Copie liens O:\Techs\outiladmins\bin\conf\liens.Flq sous " & $chemin_dossier_bin & "conf\")
	_FileCopy("O:\Techs\outiladmins\bin\conf\Config.Flq", $chemin_dossier_bin & "conf\")
	_FileWriteLog($Cheminlog, "Copie Config O:\Techs\outiladmins\bin\conf\Config.Flq sous " & $chemin_dossier_bin & "conf\")
	_FileCopy("O:\Techs\outiladmins\bin\conf\OA.dll", $chemin_dossier_bin & "conf\")
	_FileWriteLog($Cheminlog, "Copie DLL O:\Techs\outiladmins\bin\conf\OA.dll sous " & $chemin_dossier_bin & "conf\")

	If Not FileExists($chemin_dossier_bin & "icones\encours.jpg") Then
		_FileCopy(".\include\bin\icones\encours.jpg", $chemin_dossier_bin & "icones\")
	EndIf
	If Not FileExists($chemin_dossier_bin & "Lister_Sessions_Ouvertes.vbs") Then
		_FileCopy("\\nas-000\progtech$\Techs\outiladmins\Bin\Lister_Sessions_Ouvertes.vbs", $chemin_dossier_bin & "")
	EndIf
	If Not FileExists($chemin_dossier_bin & "PsExec.exe") Then
		_FileCopy("\\nas-000\progtech$\Techs\outiladmins\Bin\PsExec.exe", $chemin_dossier_bin & "")
	EndIf
	If Not FileExists($chemin_dossier_bin & "Psloggedon64.exe") Then
		_FileCopy("\\nas-000\progtech$\Techs\outiladmins\Bin\Psloggedon64.exe", $chemin_dossier_bin & "")
	EndIf
	If Not FileExists($chemin_dossier_bin & "pskill.exe") Then
		_FileCopy("\\nas-000\progtech$\Techs\outiladmins\Bin\pskill.exe", $chemin_dossier_bin & "")
	EndIf
	If Not FileExists($chemin_dossier_bin & "PsList.exe") Then
		_FileCopy("\\nas-000\progtech$\Techs\outiladmins\Bin\PsList.exe", $chemin_dossier_bin & "")
	EndIf
	If Not FileExists($chemin_dossier_bin & "WolCmd.exe") Then
		_FileCopy("\\nas-000\progtech$\Techs\outiladmins\Bin\WolCmd.exe", $chemin_dossier_bin & "")
	EndIf
	If Not FileExists($chemin_dossier_bin & "\conf\liens.txt") Then
		FileInstall(".\include\CONF\liens.txt", $chemin_dossier_bin & "conf\")
	EndIf

	Global $sFichierlien = $chemin_dossier_bin & "conf\liens.txt"

	FileInstall(".\outiladmins.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)

	;enregistrement des credential dans le fichier de conf
	cryptage($userini, 2, $Key)
	IniWrite($userini, "Credential", "Nom", $sUsername)
	IniWrite($userini, "Credential", "Pass", $sPassword)
	IniWrite($userini, "Credential", "Domaine", $Domain)
	cryptage($userini, 1, $Key)

	;SplashOff()
	#EndRegion Mise à jour globale


	_GuiCreate($version, $sFichierlien)

	Global $sIPAddress = TCPNameToIP(@ComputerName)
	Global $Result


	While 1
		; Mise à jour de l'heure
		GUICtrlSetData($LaTime, @HOUR & ":" & @MIN & ":" & @SEC)
		$msg = GUIGetMsg()

		; Filtrer uniquement les événements liés aux boutons/menus (éviter -11, -3, etc.)
		If $msg <= 0 Then ContinueLoop

		Sleep(10)
		; Lecture du PC saisi
		Global $IPPc = GUICtrlRead($addPc)
		IniWrite($paramini, "PC", "Nom", $IPPc)

		If $IPPc = "" Then
			$IPPc = @ComputerName
		EndIf

		; Vérifier si l'action correspond à un lien
		If VerifierLiencase($msg, $aliensmenu) Then
			ContinueLoop ; Si un lien est trouvé, on évite d'entrer dans le Switch
		EndIf

		; Gestion des autres boutons
		Switch $msg
			Case $msg = $iToOpenLien
				ShellExecute($sFichierlien)
			Case $BT_ExportversionWin
				_ExportVWin()
			Case $Bt_ScanRZO

				If Not FileExists($chemin_dossier_bin & "apps\netscan\netscanlkl.exe") Then
					SplashImageOn("Merci de patienter", $chemin_dossier_bin & "icones\MAJ.jpg", 428, 145, -1, 0)
					_filecopy($Sources & "\apps\Netscan\*.*", $chemin_dossier_bin & "apps\netscan\")
					_Filecopy($Sources & "\LiberKeyTools\*.*", $chemin_dossier_bin & "LiberKeyTools\")
					SplashOff()
				EndIf
				_FileWriteLog($Cheminlog, "lancement du scan Réseau")
				Run($chemin_dossier_bin & "apps\netscan\NetscanLKL.exe")
			Case $BT_ValDNS
				Switch GUICtrlRead($CBDNS)
					Case "Vide le DNS"
						RunAs($sUsername, $Domain, $sPassword, 0, @ComSpec & " /k " & '"' & $chemin_dossier_bin & '\PsExec.exe"' & " \\" & $IPPc & " ipconfig /flushdns")
						GUICtrlDelete($CBDNS)
						$CBDNS = GUICtrlCreateCombo("Choix", 164, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
						GUICtrlSetData($CBDNS, "Vide le DNS|Force enregistrement")
					Case "Force enregistrement"
						RunAs($sUsername, $Domain, $sPassword, 0, @ComSpec & " /k " & '"' & $chemin_dossier_bin & '\PsExec.exe"' & " \\" & $IPPc & " ipconfig /registerdns")
						GUICtrlDelete($CBDNS)
						$CBDNS = GUICtrlCreateCombo("Choix", 164, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
						GUICtrlSetData($CBDNS, "Réparation|Force MAJ")
					Case Else
						MsgBox(0, "", "Choisir une action")
				EndSwitch
			Case $BTValAlim
				Switch GUICtrlRead($CBAlim)
					Case "Démarrage"
						$Adip = InputBox('Machine', "saisir l'IP de l'ordi")
						$MAC = InputBox('Machine', "saisir l'adresse MAC de l'ordi")
						$SReseau = InputBox('Machine', "saisir le masque de sous réseau")
						If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
						$iMsgBoxAnswer = MsgBox(36, "outiladmins", "Le PC se trouve t-il sur un site DISTANT")
						Select
							Case $iMsgBoxAnswer = 6 ;Yes
								$PosteRelais = InputBox('WOL', "saisir le nom du poste relais")
								$MdpADMDATA = InputBox('MOt de passe', "Saisir le mot de passe ADMDATA")
								RunAs("Admdata", $Domain, $MdpADMDATA, 0, "xcopy " & $chemin_dossier_bin & "wolcmd.exe \\" & $PosteRelais & "\c$\temp\ /Y")
								ConsoleWrite(@error)
								RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $PosteRelais & " -s cmd")
								WinWaitActive("\\" & $PosteRelais & ": cmd", "")
								Sleep(200)
								Send('c:\temp\Wolcmd.exe ' & $MAC & ' ' & $Adip & ' ' & $SReseau & ' 7')
								Sleep(200)
								Send("{enter}")
							Case $iMsgBoxAnswer = 7 ;No
								RunAs($sUsername, $Domain, $sPassword, 0, 'C:\Windows\System32\cmd.exe')
								Sleep(200)
								Send($chemin_dossier_bin & 'Wolcmd.exe ' & $MAC & ' ' & $Adip & ' ' & $SReseau & ' 7')
								Send("{enter}")
						EndSelect
						GUICtrlDelete($CBAlim)
						$CBAlim = GUICtrlCreateCombo("Choix", 27, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
						GUICtrlSetData($CBAlim, "Reboot|Démarrage")
					Case "Reboot"
						RunAs($sUsername, $Domain, $sPassword, 0, "C:\Windows\System32\cmd.exe /K shutdown -r  -t 20 -m \\" & $IPPc)
						GUICtrlDelete($CBAlim)
						$CBAlim = GUICtrlCreateCombo("Choix", 27, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
						GUICtrlSetData($CBAlim, "Reboot|Démarrage")
					Case Else
						MsgBox(0, "", "Choisir une action")
				EndSwitch
			Case $BTValWsus
				Switch GUICtrlRead($CBwsus)
					Case "Réparation"
						FileInstall("\\nas-000\progtech$\Techs\outiladmins\Bin\reinitialisationWSUS.bat", $chemin_dossier_bin & "", $FC_OVERWRITE)
						RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd") ;& $IPPc&"\c$\temp\reinitialisationWSUS.bat")
						Sleep(2000)
						Send("net stop wuauserv")
						Send("{ENTER}")
						Sleep(2000)
						Send("rmdir c:\windows\SoftwareDistribution /s /q")
						Send("{ENTER}")
						Sleep(1000)
						Send("REG DELETE HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v SusClientId /f")
						Send("{ENTER}")
						Sleep(1000)
						Send("REG DELETE HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v SusClientIdValidation /f")
						Send("{ENTER}")
						Sleep(1000)
						Send("REG DELETE HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v PingID /f")
						Send("{ENTER}")
						Sleep(1000)
						Send("REG DELETE HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v AccountDomainSid /f")
						Send("{ENTER}")
						Sleep(1000)
						Send("net start wuauserv")
						Send("{ENTER}")
						Sleep(2000)
						Send("wuauclt.exe /resetauthorization /detectnow")
						Sleep(2000)
						Send("{ENTER}")
						Sleep(2000)
						Send("exit")
						Send("{ENTER}")
					Case "Force MAJ"
						GUICtrlDelete($CBwsus)
						$CBwsus = GUICtrlCreateCombo("Choix", 301, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
						GUICtrlSetData($CBwsus, "Réparation|Force MAJ")
						RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s c:\temp\instalwsus.bat")
					Case Else
						MsgBox(0, "", "Choisir une action")
				EndSwitch
			Case $BTValProcess
				Switch GUICtrlRead($CBProcess)
					Case "Liste"
						RunAs($sUsername, $Domain, $sPassword, 0, 'C:\Windows\System32\cmd.exe /K "' & $chemin_dossier_bin & 'PsList.exe" /accepteula -t -s -r 2 \\' & $IPPc)
						GUICtrlDelete($CBProcess)
						$CBProcess = GUICtrlCreateCombo("Choix", 438, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
						GUICtrlSetData($CBProcess, "Liste|Arrêt")
					Case "Arrêt"
						Local $processtokill = InputBox("*** Quel Process ***", "Quel est le nom ou l'ID du Process à Killer ?", "Excel.exe", "", 280, 130)
						If @error <> 1 Then
							RunAs($sUsername, $Domain, $sPassword, 0, 'C:\Windows\System32\cmd.exe /K "' & $chemin_dossier_bin & "pskill.exe /accepteula -t \\" & $IPPc & ' "' & $processtokill & '"')
							GUICtrlDelete($CBProcess)
							$CBProcess = GUICtrlCreateCombo("Choix", 438, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
							GUICtrlSetData($CBProcess, "Liste|Arrêt")
						EndIf
					Case Else
						MsgBox(0, "", "Choisir une action")
				EndSwitch
			Case $BTFicSSI
				If GUICtrlRead($RadOrdiSSI) = $GUI_CHECKED Then
					startssi()
					ssiO($IPPc)
				ElseIf GUICtrlRead($RadMatSSI) = $GUI_CHECKED Then
					startssi()
					ssiM($IPPc)
				Else
					MsgBox(0, "", "sélectionner un type de fiche")
				EndIf
			Case $Bt_check_Dock
				FileInstall(".\include\bin\SNStation.ps1", $chemin_dossier_bin & "", $FC_OVERWRITE)
				RunAs($sUsername, $Domain, $sPassword, 0, "powershell -ExecutionPolicy ByPass -File " & $chemin_dossier_bin & "\SNStation.ps1")
			Case $Bt_check_ecran
				FileInstall(".\include\bin\bddecran.ps1", $chemin_dossier_bin & "", $FC_OVERWRITE)
				RunAs($sUsername, $Domain, $sPassword, 0, "powershell -ExecutionPolicy ByPass -File " & $chemin_dossier_bin & "\bddecran.ps1")
			Case $Icohotline
				ShellExecute("http://apollon/dgs/sit/asp_formulaires/incident/consultation_m.asp?param=id_incident")
			Case $Visioitem
				ShellExecute("K:\Reseau\5-Divers (Techniciens)\Fiches procédures - Visio et audio conférences")
			Case $PhotosItem
				ShellExecute("K:\Reseau\2-Batiments")
			Case $MenuResultGPO
				FileInstall(".\Documentations\GPO.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\GPO.pdf", "", "", "")
			Case $MenucheckBTL
				FileInstall(".\include\Bitlocker\bitlocker.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\bitlocker.pdf", "", "", "")
			Case $MenudocLiensPersos
				FileInstall(".\Documentations\Modification des liens persos.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\Modification des liens persos.pdf", "", "", "")
			Case $MenuRaccourci
				FileInstall(".\Documentations\raccourcis.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\raccourcis.pdf", "", "", "")
			Case $OrangeBItem
				ShellExecute("https://www.orange-business.com/fr")
			Case $Majconfig
				cryptage($configini, 2, $Key)
				_GuiEDITini($configini)
				cryptage($configini, 1, $Key)
			Case $MajIniitem
				_CheckAndUpdateComponent("Crypt", "crypt.exe")
			Case $Bt_update
				_CheckAndUpdateComponent("update", "update.exe")
			Case $Bt_GestWifi
				IniWrite($paramini, "PC", "Nom", $IPPc)
				_FileWriteLog($Cheminlog, "lancement de la sauvegarde Wifi sur : " & $IPPc)
				FileInstall(".\include\Gestion WIFI\restoWiFi.bat", $chemin_dossier_bin & "", $FC_OVERWRITE)
				FileInstall(".\include\Gestion WIFI\savwifi.bat", $chemin_dossier_bin & "", $FC_OVERWRITE)
				_CheckAndUpdateComponent("Wifi", "savwifi.exe")
			Case $BT_SuiviRZO
				_CheckAndUpdateComponent("Reseau", "Reseau.exe")
			Case $BT_CheckBitlocker
				_CheckAndUpdateComponent("Bitlocker", "Bitlocker.exe")
			case $BTUserLoggin
				SetLastLoggedOnUser($IPPc)
			Case $BT_ScintEcran
				_FileCopy($Sources & "DSC.exe", $chemin_dossier_bin & "")
				_FileWriteLog($Cheminlog, "Correction du scintillement sur " & $IPPc)
				RunAs($sUsername, $Domain, $sPassword, 0, "xcopy " & $chemin_dossier_bin & "dsc.exe \\" & $IPPc & "\c$\temp\ /Y")
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
				WinWaitActive("\\" & $IPPc & ": cmd", "")
				Sleep(200)
				Send("cd \temp {Enter}")
				Sleep(200)
				Send("dsc.exe 1 {Enter}")
				Sleep(200)
				Send("{Enter}")
				Sleep(2000)
			Case $BT_DockLenovo
				cryptage($Versionini, 2, $Key)
				If IniRead($Versionini, "Station", "version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\Station", "Version") Then
					_FileCopy($Sources & "station.exe", $chemin_dossier_bin & "")
				EndIf
				cryptage($Versionini, 1, $Key)
				_FileWriteLog($Cheminlog, "Recherche des informations sation LENOVO sur " & $IPPc)
				RunAs($sUsername, $Domain, $sPassword, 0, "xcopy " & $chemin_dossier_bin & "station.exe \\" & $IPPc & "\c$\temp\ /Y")
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
				WinWaitActive("\\" & $IPPc & ": cmd", "")
				Sleep(200)
				Send("cd \temp {Enter}")
				Sleep(200)
				Send("station.exe {Enter}")
				Sleep(200)
				Send("{Enter}")
				Sleep(2000)
			Case $BT_Dism
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
				WinWaitActive("\\" & $IPPc & ": cmd", "")
				Sleep(200)
				Send("cd \ {Enter}")
				Sleep(200)
				Send("DISM /Online /Cleanup-image /Restorehealth {Enter}")
			Case $TVBOUDONitem
				_Word_DocOpen(_Word_Create(), "K:\Reseau\5-Divers (Techniciens)\Fiches procédures - Logistique\LOGISTIQUE - Projection_BOUDON 2024.docx")
			Case $menuRefEtude
				_Excel_BookOpen(_Excel_Open(), "K:\Reseau\0-Admin&Tech\Utilisateurs\Référents études pour habilitations applicatives.xls")
			Case $BT_Member
				_CheckAndUpdateComponent("GestGroup", "Gestion groupes.exe")
			Case $BT_SerialEcran
				FileInstall(".\include\bin\Ecran.ps1", $chemin_dossier_bin & "", $FC_OVERWRITE)
				_AD_Open($sUsername, $sPassword)
				_FileWriteLog($Cheminlog, "ouverture AD : " & @error)
				RunAsWait($sUsername, $Domain, $sPassword, 0, "powershell -ExecutionPolicy ByPass -File " & $chemin_dossier_bin & "ecran.ps1 -ComputerName " & $IPPc)
				_FileWriteLog($Cheminlog, "Execution Script code retour : " & @error)
				ShellExecute("c:\temp\" & $IPPc & ".txt")
				_ad_close()
				_FileCopy("c:\temp\" & $IPPc & ".txt", "\\vulcain\fitrans$\dsian\Log\ecran\" & $IPPc & ".txt")
						case $Bt_DepTel
				_CheckAndUpdateComponent("Telephone", "Telephone.exe")
			Case $BT_Dock
				_CheckAndUpdateComponent("Cherchestation", "cherche station.exe")
			Case $BT_Ecran
				_CheckAndUpdateComponent("ChercheEcran", "ecran.exe")
			Case $BT_MAC_PC
				_CheckAndUpdateComponent("CherchePC", "cherche PC.exe")
			Case $BT_GestADPC
				_GestADPC()
			Case $IDtoDrivers
				ShellExecute("https://driverpack.io/fr/catalog")
			Case $Cyber
				_Cyber()
			Case $BT_GestionUser
				_GestionUser()
						Case $menuModifINI
				FileInstall(".\Documentations\MajIni.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\MajIni.pdf", "", "", "")

			Case $DocResolvpb
				FileInstall(".\Documentations\résolution des PBs.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\résolution des PBs.pdf", "", "", "")
			Case $AlerteOrangeItem
				FileInstall(".\Documentations\mail orange.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\mail orange.pdf", "", "", "")
			Case $menuJAII
				FileInstall(".\Documentations\jaii.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\jaii.pdf", "", "", "")
			Case $MenuInfoPC
				FileInstall(".\include\infoPC\info PC.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\info PC.pdf", "", "", "")
			Case $Adobeitem
				ShellExecute("https://adminconsole.adobe.com/")
			Case $Autocaditem
				ShellExecute("https://manage.autodesk.com/user-access/users/user-list")
			Case $MenuLenovo
				FileInstall(".\include\Station Lenovo\Station Lenovo.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\Station Lenovo.pdf", "", "", "")
			Case $MenuConnectSharp
				FileInstall(".\Documentations\Connexion Imprimantes.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\connexion Imprimantes.pdf", "", "", "")
			Case $MenuPresent
				FileInstall(".\Documentations\Outils Admins.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\Outils Admins.pdf", "", "", "")
			Case $MenuAutoUpdate
				FileInstall(".\include\update\update.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\update.pdf", "", "", "")
			Case $MenucheckSAV
				FileInstall(".\include\Sauvegardes\suivi sauvegardes.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\suivi sauvegardes.pdf", "", "", "")
			Case $Majitem
				FileInstall(".\Documentations\A propos.txt", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\A propos.txt", "", "", "")
				_FileWriteLog($Cheminlog, "++++++Lancement de A PROPOS +++++++++")
			Case $MenuGestWifi
				FileInstall(".\include\Gestion WIFI\GestionWifi.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\GestionWifi.pdf", "", "", "")
			Case $Menuapplisinstall
				FileInstall(".\include\applis\info applis.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\info applis.pdf", "", "", "")
			Case $MenuConectRzo
				FileInstall(".\include\Controle connectivite\controle reseau.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\controle reseau.pdf", "", "", "")
			Case $MenuNAS
				FileInstall(".\include\bin\Panne NAS.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\Panne NAS.pdf", "", "", "")
			Case $MenuGestUserAD
				FileInstall(".\include\MDP\Gestion des Utilisateurs.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\Gestion des Utilisateurs.pdf", "", "", "")
			Case $MenuGestPCAD
				FileInstall(".\include\Gestion PC\GestionPC.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\GestionPC.pdf", "", "", "")
			Case $SuppMAJitem
				FileInstall(".\Documentations\Désinstallation des Mises à jours de sécurité.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\Désinstallation des Mises à jours de sécurité.pdf", "", "", "")
			Case $MenuAdminLocaux
				FileInstall(".\include\Check Adm Locaux\AdminLocaux.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\AdminLocaux.pdf", "", "", "")
			Case $MenuCyber
				FileInstall(".\include\Cyber47\Cyber47.pdf", $chemin_dossier_bin & "doc\", $FC_OVERWRITE)
				ShellExecute($chemin_dossier_bin & "doc\cyber47.pdf", "", "", "")
			Case $Bt_LastlogonAD
				; Connection à l'Active Directory
				_AD_Open()
				_FileWriteLog($Cheminlog, "Recherche de la dernière connexion AD du PC " & $IPPc)
				If @error Then Exit MsgBox(16, "Erreur Active Directory", "Erreur de connexion à l'AD. @error = " & @error & ", @extended = " & @extended)

				Global $aPropertiesO[1][2]
				;Recherche l'attribut lastlogon sur l'AD
				$aPropertiesO = _AD_GetObjectProperties($IPPc & "$", "lastlogon")
				If @error Then
					MsgBox(16, "Erreur Active Directory", "Ordinateur inconnu")
				Else
					;Affichage du résultat
					MsgBox(0, "", "Dernière connexion à l'AD " & @CRLF & WMIDateStringToDate($aPropertiesO[1][1]))
				EndIf
				_AD_Close()
			Case $Bt_GrpADPc
				_AD_Open()
				If @error Then Exit MsgBox(16, "Erreur Active Directory", "Erreur de connexion à l'AD. @error = " & @error & ", @extended = " & @extended)
				Global $aProperties[1][2]
				;Recherche l'attribut membre de sur l'AD
				$aProperties = _AD_GetObjectProperties($IPPc & "$", "memberof")
				;Affichage du tableau de résultat
				If @error Then
					MsgBox(16, "Erreur Active Directory", "Ordinateur inconnu")
				Else
					_ArrayDisplay($aProperties, "Groupes AD pour l'ordinateur '" & $IPPc & "'", "1:|2:", 64, Default, "Membre de|Nom du groupe")
				EndIf
				_AD_Close()
			Case $BT_InfoPC
				_infopc($IPPc)
			Case $BT_checkSAV
				$Site = GUICtrlRead($CB_NAS)
				_FileWriteLog($Cheminlog, "Check SAV sur " & $Site)
				IniWrite($paramini, "NAS", "Nom", $Site)
				If $Site = "choix" Then
					MsgBox(0, "Attention", "Veuillez choisir un site à contrôler")
				ElseIf $Site = "000_Saint-Jacques" Or $Site = "000-COM" Or $Site = "000-SAV_Saint-Jacques" Or $Site = "000-Backup_Saint-Jacques" Or $Site = "000-COM-SAV" Or $Site = "001_Scaliger" Or $Site = "100_Dolet" Or $Site = "100_Verdun" Or $Site = "101_Jean Bru" Or $Site = "101B_Jean Bru" Then
					CheckSAVPCSJ($Site)
				Else
					CheckSAVPC($Site)
				EndIf
			Case $BT_ValNAS
				$ConnNas = GUICtrlRead($CB_NAS)
				cryptage($configini, 2, $Key)
				_FileWriteLog($Cheminlog, "decryptage de  " & $configini & " code erreur " & @error)
				$Nomnas = IniRead($configini, $ConnNas, "Nas", "")
				$port = IniRead($configini, $ConnNas, "Port", "")
				cryptage($configini, 1, $Key)
				_FileWriteLog($Cheminlog, "Connexion sur le Nas  " & $ConnNas & " IP " & $Nomnas)
				ShellExecute("http:" & $Nomnas & ":" & $port & "/#/signin")
			Case $GUI_EVENT_CLOSE, $BT_QUIT
				FileDelete($chemin_dossier_bin & "outiladmins.au3")
				FileDelete($configini)
				FileDelete($paramini)
				FileDelete($liensini)
				;FileDelete($Versionini)
				FileDelete($userini)
				FileDelete($dlloa)
				_FileWriteLog($Cheminlog, "Fermeture Outils Admins")
				; Si l'ordinateur utilisé n'est pas celui d'un des techniciens du pole on supprime le dossier $chemin_dossier_bin pour ne laisser aucune trace des scripts et des docs
				; Liste des ordinateurs autorisés
				$ordinateursTech = "P000-T4563|P000-T4181|P000-T4133|P000-T4612|P000-T4187|P000-T4171|P000-T4183|P000-T4680|P000-T4682" ; Remplacez par les noms d'ordinateur autorisés
;~ 				; Obtenez le nom de l'ordinateur actuel
				$nomOrdinateur = @ComputerName
				; Vérifiez si le nom de l'ordinateur fait partie de la liste autorisée
				If StringInStr($ordinateursTech, $nomOrdinateur) Then
					; L'ordinateur fait partie de la liste autorisée, ne rien faire
				Else
					; Suppression du dossier
					If DirRemove($chemin_dossier_bin, 1) Then
					Else
						MsgBox(16, "Erreur", "La suppression du dossier Bin a échoué." & @CRLF & @CRLF & "Merci de le faire manuellement")
					EndIf
				EndIf
				Exit
			Case $BT_ConecIMP
				$IPCOPIEUR = InputBox("IP du copieur", "IP du copieur", "10.128.50.")
				_FileWriteLog($Cheminlog, "Prise en main sur le copieur " & $IPCOPIEUR)
				ShellExecute("http://129.47.0.138:8085/VNCConverter/" & $IPCOPIEUR & ":5900/?locale=en&modelName=SHARPMX-M5070N&ipAddress=" & $IPCOPIEUR)
			Case $BT_applis
				_applis($IPPc)
			Case $BT_AdminLocAD
				cryptage($Versionini, 2, $Key)
				If IniRead($Versionini, "GrpAdminslocaux", "version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\GrpAdminslocaux", "Version") Then
					_FileCopy($Sources & "groupeadminslocaux.exe", $chemin_dossier_bin & "")
				EndIf
				cryptage($Versionini, 1, $Key)
				$passadm = InputBox("Mot de passe", "Saisir le mot de passe ADMTECH", "", "*")
				_FileWriteLog($Cheminlog, "Lancement de la gestion des groupes admins locaux")
				RunAs($admtech, $Domain, $passadm, 0, $chemin_dossier_bin & "groupeadminslocaux.exe ")
			Case $BT_AdminLoc
				cryptage($Versionini, 2, $Key)
				If IniRead($Versionini, "GestAdminslocaux", "version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\GrpAdminslocaux", "Version") Then
					_FileCopy($Sources & "GestionAdmLocaux.exe", $chemin_dossier_bin & "")
				EndIf
				cryptage($Versionini, 1, $Key)
				$passadm = InputBox('Mot de passe', 'Saisir le mot de passe ADMTECH', '', '*')
				_FileWriteLog($Cheminlog, "Lancement de la gestion des admins locaux")
				RunAs($admtech, $Domain, $passadm, 0, $chemin_dossier_bin & "GestionAdmLocaux.exe ")
			Case $BTPing
				RunAs($sUsername, $Domain, $sPassword, 0, "C:\Windows\System32\ping.exe " & $IPPc & " -t")
			Case $BT_Whosconnect
				RunAs($sUsername, $Domain, $sPassword, 0, 'C:\Windows\System32\cmd.exe /K ' & $chemin_dossier_bin & 'Psloggedon64.exe -nobanner \\' & $IPPc)
			Case $BT_GPupdate
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s gpupdate ")
			Case $BT_RSOP
				adresseIP($IPPc)
				Local $iPing = Ping($IPPc, 250) ; vérifie la connexion du PC au réesau
				If $iPing Then
					$sUserRsop = InputBox('User', "Saisir le nom de l'utilisateur pour lequel vérifier les GPOs")
					_FileWriteLog($Cheminlog, "Vérification des GPOs sur le poste " & $IPPc & " pour l'utilisateur : " & $sUserRsop)
					SplashImageOn("Merci de patienter", $chemin_dossier_bin & "icones\encours.jpg", 220, 220, -1, -1, 16)
					If FileExists("\\" & $IPPc & "\C$\temp\rapport.html") Then FileDelete("\\" & $IPPc & "\C$\temp\rapport.html")
					RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
					WinWaitActive("\\" & $IPPc & ": cmd ")
					Send(" gpresult /user:" & $sUserRsop & " /H C:\temp\rapport.html ")
					Send("{enter}")
					Send("Exit")
					Send("{enter}")
					WinWaitClose("\\" & $IPPc & ": cmd ")
					_FileWriteLog($Cheminlog, "fermetur de la fenêtre")
					_filecopy("\\" & $IPPc & "\c$\temp\rapport.html", "c:\temp\rapport\rapport" & $IPPc & $sUserRsop & @HOUR & @MDAY & ".html")
					_FileWriteLog($Cheminlog, "copie du rapport")
					SplashOff()
					ShellExecute("c:\temp\rapport\rapport" & $IPPc & $sUserRsop & @HOUR & @MDAY & ".html")
				Else
					MsgBox(48, "Avertissement", "Le PC " & $IPPc & " n'est pas joignable. Il doit être connecté au réseau pour lancer la commande")
				EndIf
			Case $BT_C
				RunAs($sUsername, $Domain, $sPassword, 0, "C:\Windows\explorer.exe \\" & $IPPc & "\c$")
				_FileWriteLog($Cheminlog, "Connexion au partage administratif sur " & $IPPc)
			Case $BT_GestinPC
				RunAs($sUsername, $Domain, $sPassword, 0, 'C:\Windows\System32\cmd.exe /C C:\Windows\System32\mmc.exe compmgmt.msc /computer="' & $IPPc & '"')
			Case $icoVNC
				_FileWriteLog($Cheminlog, "Lancement de la prise en main sur " & $IPPc)
				Run("C:\Program Files\TightVnc\tvnviewer.exe " & $IPPc)
			Case $BT_CMDDISTANT
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
			Case $BT_Infosys
				_FileWriteLog($Cheminlog, "Lancement d'info système sur " & $IPPc)
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
				WinWaitActive("\\" & $IPPc & ": cmd", "")
				Sleep(2000)
				Send("systeminfo {Enter}")
			Case $BT_Registre
				RunWait("sc \\" & $IPPc & " start remoteregistry")
				Sleep(1000)
				RunAs($sUsername, $Domain, $sPassword, 0, "C:\Windows\System32\cmd.exe /C C:\Windows\regedit.exe")
				Local $err = WinWaitActive("Éditeur du Registre", "", 5)
				Sleep(1000)
				Send("!f")
				Sleep(200)
				Send("c")
				Sleep(2000)
				Send($IPPc)
				Sleep(200)
				Send("{ENTER}")
			Case $BT_Printer
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
				Sleep(3000)
				Send("wmic printer get caption")
				Send("{ENTER}")
			Case $BT_VerifIP
				verifIP($IPPc)
			Case $bt_FlushDNS
				RunWait(@ComSpec & " /c ipconfig /flushdns")
			Case $BT_Opale
				; Si aucun nom d'ordinateur n'est saisi $IPPC prend le nom du pc local, la requête ne marche pas
				If $IPPc = @ComputerName Then
					MsgBox(0, "", "Pour saisir le code opale de l'ordinateur sur lequel vous êtes connecté, " & @CRLF & " Utilisez Outils Tech", $MB_RIGHT)
				Else
					; récupération du numéro Opale dans le fichier créé par le script de connexion
					; Ouverture du log du poste
					Local $rep_du_site = "T:\dsian\Log\bdd\" & $IPPc & ".txt"
					If @error = 1 Then
						Exit
					EndIf
					$file = FileOpen($rep_du_site)
					$sOpale = ""
					Local $numeroLigne = _FileCountLines($file)
					;Récupération des informations
					While $numeroLigne > 0
						$ligne = FileReadLine($file, $numeroLigne)
						If StringInStr($ligne, $IPPc) Then
							$elements = StringSplit($ligne, ",")
							ExitLoop
						EndIf
						$numeroLigne -= 1
					WEnd
					FileClose($file)
					;Isolation du code Opale
					If UBound($elements) > 0 Then
						$sOpale = $elements[7]
					EndIf
					MsgBox($MB_SYSTEMMODAL, "", "l'opale actuel est : " & @CRLF & $sOpale)
					If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
					$iMsgBoxAnswer = MsgBox(4, "Numéro Opale", "Voulez-vous changer le numéro Opale")
					Select
						Case $iMsgBoxAnswer = 6 ;Oui
							$saisieopale = InputBox("Avertissement", "Saisir le Numéro Opale de l'ordinateur", "", " M8")
							If @error = 1 Then
								MsgBox(64, "Opale", "le Numéro Opale n'a pas été changé")
							Else
								RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
								Sleep(2000)
								Send("REG add HKEY_LOCAL_MACHINE\SOFTWARE\CD47 /v opale /d " & $saisieopale & " /f")
								Send("{ENTER}")
								Sleep(2000)
								Send("exit")
								Send("{ENTER}")
								RunAs($sUsername, $Domain, $sPassword, 0, "\\srv-mdt2016\gpo$\script\lanceur.exe")
								_FileWriteLog($Cheminlog, "saisie du code OPALE " & $saisieopale & " sur le PC " & $IPPc)
							EndIf
						Case $iMsgBoxAnswer = 7 ;Non
							MsgBox(64, "Opale", "le Numéro Opale n'a pas été changé")
					EndSelect
				EndIf
		EndSwitch
		Sleep(10)

		;Utilisation de raccourci clavier si la fenêtre Outils Admins est active
		If WinActive($Form1) Then
			; Appuis sur la touche entrée vérifie si le PC est en ligne
			If _IsPressed("0D") Then
				verifIP($IPPc)
			EndIf
			; Appuis sur la touche F1 pour lancer les menus de documentations en fonction de l'onglet sur lequel on se trouve
			If _IsPressed("70") Then
				$msg2 = GUIGetMsg()
				Switch GUICtrlRead($PageControl1)
					Case $PageControl1
					Case 0
						Send("{alt}{enter}{down}{right}")
					Case 1
						Send("{alt}{enter}{down}{down}{right}")
					Case 2
						Send("{alt}{enter}{down}{down}{down}{right}")
					Case 3
						Send("{alt}{enter}{down}{down}{down}{down}{right}")
				EndSwitch
			EndIf
			If _IsPressed("71") Then ;F2
				startssi()
				ssiO($IPPc)
			EndIf
			If _IsPressed("72") Then ;F3
				startssi()
				ssiM($IPPc)
			EndIf
			If _IsPressed("73") Then ;F4
				_FileWriteLog($Cheminlog, "Lancement de la prise en main sur " & $IPPc)
				Run("C:\Program Files\TightVnc\tvnviewer.exe " & $IPPc)
			EndIf
			If _IsPressed("74") Then ;F5
				_GestADPC()

			EndIf
			If _IsPressed("75") Then ;F6
				_GestionUser()
			EndIf
			If _IsPressed("77") Then ;F8
				_infopc($IPPc)
			EndIf
			If _IsPressed("76") Then ;F7
				RunAs($sUsername, $Domain, $sPassword, 0, 'C:\Windows\System32\cmd.exe /K ' & $chemin_dossier_bin & 'Psloggedon64.exe -nobanner \\' & $IPPc)
			EndIf
			If _IsPressed("78") Then ;F9
				RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $IPPc & " -s cmd")
			EndIf
			If _IsPressed("79") Then ;F10
				RunAs($sUsername, $Domain, $sPassword, 0, "C:\Windows\explorer.exe \\" & $IPPc & "\c$")
				_FileWriteLog($Cheminlog, "Connexion au partage administratif sur " & $IPPc)
			EndIf
			If _IsPressed("7A") Then ;F11
				_AD_Open()
				_FileWriteLog($Cheminlog, "Recherche de la dernière connexion AD du PC " & $IPPc)
				If @error Then Exit MsgBox(16, "Erreur Active Directory", "Erreur de connexion à l'AD. @error = " & @error & ", @extended = " & @extended)

				Global $aPropertiesO[1][2]
				;Recherche l'attribut lastlogon sur l'AD
				$aPropertiesO = _AD_GetObjectProperties($IPPc & "$", "lastlogon")
				If @error Then
					MsgBox(16, "Erreur Active Directory", "Ordinateur inconnu")
				Else
					;Affichage du résultat
					MsgBox(0, "", "Dernière connexion à l'AD " & @CRLF & WMIDateStringToDate($aPropertiesO[1][1]))

				EndIf
				_AD_Close()
			EndIf
			If _IsPressed("7B") Then ;F12
				_Cyber()
			EndIf
			If _IsPressed("11") = 1 And _IsPressed("41") = 1 Then ;Ctrl + A
				_applis($IPPc)
			EndIf
		EndIf
	WEnd

EndFunc   ;==>_fenetre
