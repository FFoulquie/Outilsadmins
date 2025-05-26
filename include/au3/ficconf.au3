
#cs ----------------------------------------------------------------------------

 AutoIt Version : 3.3.14.5
 Auteur:         FFOULQUIE

 Fonction du Script :
	Gestion des variables communes à outils admins et ses sous exe.

#ce ----------------------------------------------------------------------------

#include <AutoItConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <cryptage.au3>

Global $Key = "5aQXvTCr4dhBF48zusUC"

Global $chemin_dossier_bin = "C:\Progra~1\DSIAN\outilsAdmin\Bin\"
Global $Cheminlog = "C:\ProgramData\CD47\Logs\OA_" & @MON & StringRight(@YEAR, 4) & ".log"

global $Versionini = $chemin_dossier_bin & "conf\version.Flq"
global $configini = $chemin_dossier_bin & "conf\config.Flq"
global $paramini = $chemin_dossier_bin & "conf\param.Flq"
global $liensini = $chemin_dossier_bin & "conf\liens.Flq"
global $userini = $chemin_dossier_bin & "conf\user.Flq"
Global $Sources = "\\nas-000\progtech$\Techs\outiladmins\bin\"
global $admtech = "admtech"

global $listeNas = "000_Saint-Jacques|000-COM|000-SAV_Saint-Jacques|000-Backup_Saint-Jacques|000-COM-SAV|001_Scaliger|011_CE Bon-Encontre|012_CE Tonneins|013_CE Duras|014_CE Cancon|015_CE Port Ste-Marie|016_CE Miramont|017_CE Monflanquin|020_CMS Miramont|021_CE Casteljaloux|022_UD Marmande|024_CE Condezaygues|025_UD Villeneuve|026_UD Agen|027_CE Nérac|029_CE Navigation|030_Parc Routier|042_CMS Montanou|043_FEB|045_CMS Louis-Vivent|050_CMS Fumel|051_CMS Villeneuve|052_CMS Marmande|053_CMS Tonneins|054_CMS Nérac|055_CMS Tapie|057_MDC-OPP|058_CMS Casteljaloux|060_Annexe CMS Villeneuve|100_Dolet|100_Verdun|101_Jean Bru|101B_Jean Bru|103_Médiathèque|"

global $DLLOA = $chemin_dossier_bin & "conf\OA.dll"
global $userdll = $chemin_dossier_bin & "conf\user.dll"

