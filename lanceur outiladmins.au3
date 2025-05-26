#include <File.au3>
#include <FileConstants.au3>
#include <cryptage.au3>
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\..\..\..\..\..\..\Icones\system_utilities.ico
#AutoIt3Wrapper_Res_ProductName=Outils Admin
#AutoIt3Wrapper_Res_LegalCopyright=@FFoulquié
#AutoIt3Wrapper_Res_Language=1036
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include ".\include\au3\ficconf.au3"
#include ".\include\au3\fonctions.au3"

DirCreate($chemin_dossier_bin)
DirCreate($chemin_dossier_bin & "icones")
DirCreate($chemin_dossier_bin & "doc")
DirCreate($chemin_dossier_bin & "conf")


FileInstall(".\include\bin\icones\MAJ.jpg", $chemin_dossier_bin & "icones\", $FC_OVERWRITE)

splashImageOn("Merci de patienter", $chemin_dossier_bin & "icones\MAJ.jpg", 428, 145)

_FileCopy("\\nas-000\progtech$\Techs\outiladmins\bin\conf\version.Flq", $chemin_dossier_bin & "conf\")
_FileWriteLog($Cheminlog, "**Lanceur**    Copie Version  sous " & $chemin_dossier_bin & "conf\")


cryptage($Versionini, 2, $Key)
_FileWriteLog($Cheminlog, "**Lanceur**    "&  @error)
$verattendue = IniRead($Versionini, "outilsadmins", "Version", "")


_FileWriteLog($Cheminlog, "**Lanceur**  attendu  "&  $verattendue)
If IniRead($Versionini, "outilsadmins", "Version", "") <> RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins", "Version") Then
	_FileWriteLog($Cheminlog, "**Lanceur** Mise à jour en version " & IniRead($Versionoaini, "outilsadmins", "Version", ""))
	ConsoleWrite("pas vu")
	cryptage($Versionini, 1, $Key)

	_FileCopy("\\nas-000\progtech$\Techs\outiladmins\bin\outiladmins.exe", $chemin_dossier_bin & "")
	Run($chemin_dossier_bin & "outiladmins.exe")

	SplashOff()
Else
	cryptage($Versionini, 1, $Key)
	_FileWriteLog($Cheminlog, RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins", "Version"))
	_FileWriteLog($Cheminlog, "**Lanceur** version OK " & IniRead($Versionini, "outilsadmins", "Version", ""))
	Run($chemin_dossier_bin & "outiladmins.exe")
	SplashOff()

EndIf
