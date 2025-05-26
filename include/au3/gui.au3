#include <ButtonConstants.au3>
#include <ColorConstants.au3>
#include <ComboConstants.au3>
#include <Date.au3>
#include <GDIPlus.au3>
#include <StaticConstants.au3>
#include <GuiPerso.au3>

Func _GuiCreate($version, $sFichierlien)
	_GDIPlus_Startup()

	Global Const $COLOR_BG = 0xA0A0A4
	Global Const $FONT_NAME = "Calibri"
	Global Const $FONT_SIZE = 9

	Global $Form1 = GUICreate("Outils Admins    v: " & $version, 661, 764, 515, 154)


	#Region gestion du menu
	Global $MenuItem1 = GUICtrlCreateMenu("&Documentations Outils Admins")
	Global $MenuItem3 = GUICtrlCreateMenu("&Liens")
	Global $MenuItem2 = GUICtrlCreateMenu("&Documentations Diverses")
	Global $liensmenu = GUICtrlCreateMenu("Liens Persos")
	Global $MenuItem5 = GUICtrlCreateMenu("&?")

	Global $MenuPresent = GUICtrlCreateMenuItem("Présentation", $MenuItem1)
	Global $MenuRaccourci = GUICtrlCreateMenuItem("Raccourcis Clavier", $MenuItem1)
	Global $MenuGestPC = GUICtrlCreateMenu("&Onglet Gestion du PC", $MenuItem1)
	Global $MenuGestAD = GUICtrlCreateMenu("&Onglet Gestion AD", $MenuItem1)
	Global $MenuGestGrpAdm = GUICtrlCreateMenu("&Onglet Gestion Groupes Admins", $MenuItem1)
	Global $MenuGestDivers = GUICtrlCreateMenu("&Onglet Gestion Divers", $MenuItem1)
	Global $Menulienspersos = GUICtrlCreateMenu("&Onglet Liens Persos", $MenuItem1)
	Global $MenuliensAide = GUICtrlCreateMenu("&Onglet ?", $MenuItem1)

	Global $menuModifINI = GUICtrlCreateMenuItem("Modification des fichiers de Conf", $MenuliensAide)
	Global $MenudocLiensPersos = GUICtrlCreateMenuItem("Modification des liens persos", $Menulienspersos)
	Global $AppliInternetitem = GUICtrlCreateMenu("Gestions applis", $MenuItem3)
	Global $SuiviOrangeitem = GUICtrlCreateMenu("Suivi Orange", $MenuItem3)
	Global $Diversitem = GUICtrlCreateMenu("Drivers", $MenuItem3)
	Global $Logicielsitem = GUICtrlCreateMenu("Logiciels", $MenuItem3)

	Global $Resopbsitem = GUICtrlCreateMenu("Aide à la résolution des pannes", $MenuItem2)
	Global $impimantesitem = GUICtrlCreateMenu("imprimantes", $MenuItem2)
	Global $Wsusitem = GUICtrlCreateMenu("WSUS", $MenuItem2)
	Global $SalleReuitem = GUICtrlCreateMenu("Salles de réunions", $MenuItem2)

	Global $MenuLenovo = GUICtrlCreateMenuItem("Gestion des stations Lenovo", $MenuGestPC)
	Global $MenuGestWifi = GUICtrlCreateMenuItem("Gestion Wifi", $MenuGestPC)
	Global $MenuResultGPO = GUICtrlCreateMenuItem("Résultat GPO", $MenuGestPC)
	Global $Menuapplisinstall = GUICtrlCreateMenuItem("Infos applis installées", $MenuGestPC)
	Global $MenucheckBTL = GUICtrlCreateMenuItem("Controle Bitlocker", $MenuGestPC)
	Global $MenuCyber = GUICtrlCreateMenuItem("Cyber47&&Moi", $MenuItem1)

	Global $MenuAdminLocaux = GUICtrlCreateMenuItem("Gestion admins locaux", $MenuGestGrpAdm)
	Global $MenuGestPCAD = GUICtrlCreateMenuItem("Gestion des PC sur l'AD", $MenuGestAD)
	Global $MenuGestUserAD = GUICtrlCreateMenuItem("Gestion des utilisateurs sur l'AD", $MenuGestAD)
	Global $MenuInfoPC = GUICtrlCreateMenuItem("Infos PC", $MenuGestPC)
	Global $MenuConnectSharp = GUICtrlCreateMenuItem("Connexion au copieurs Sharp", $MenuGestDivers)
	Global $MenuNAS = GUICtrlCreateMenuItem("Panne d'un NAS", $MenuGestDivers)
	Global $MenuConectRzo = GUICtrlCreateMenuItem("Vérification réseau des sites distants", $MenuGestDivers)
	Global $MenucheckSAV = GUICtrlCreateMenuItem("Vérification des sauvegardes utilisateurs", $MenuGestDivers)
	Global $MenuAutoUpdate = GUICtrlCreateMenuItem("Mise à jour AutoUpdate", $MenuGestDivers)


	Global $IDtoDrivers = GUICtrlCreateMenuItem("Trouver drivers depuis le numéro d'identification", $Diversitem)
	Global $Adobeitem = GUICtrlCreateMenuItem("Adobe", $AppliInternetitem)

	Global $OrangeBItem = GUICtrlCreateMenuItem("Orange Business", $SuiviOrangeitem)
	Global $AlerteOrangeItem = GUICtrlCreateMenuItem("Réception d'un mail ", $SuiviOrangeitem)
	Global $PhotosItem = GUICtrlCreateMenuItem("Photos des Baies", $SuiviOrangeitem)
	Global $Autocaditem = GUICtrlCreateMenuItem("Autocad", $AppliInternetitem)

	Global $SuppMAJitem = GUICtrlCreateMenuItem("Suppression des mises à jours", $Wsusitem)
	Global $TVBOUDONitem = GUICtrlCreateMenuItem("Réglage TV Boudon", $SalleReuitem)
	Global $Visioitem = GUICtrlCreateMenuItem("Docs visios", $SalleReuitem)
	Global $menuJAII = GUICtrlCreateMenuItem("JAII Crédit utilisateur", $impimantesitem)
	Global $DocResolvpb = GUICtrlCreateMenuItem("Aide à la résolution", $Resopbsitem)

	Global $Majitem = GUICtrlCreateMenuItem("A propos", $MenuItem5)
	Global $Majconfig = GUICtrlCreateMenuItem("Visualisation fichier config", $MenuItem5)
	Global $MajIniitem = GUICtrlCreateMenuItem("MAJ ini", $MenuItem5)

	Global $menuRefEtude = GUICtrlCreateMenuItem("Référents études", $Logicielsitem)


	;///////////////////////////////////////////
	; Créer un menu Lien
	; global $liensmenu = GUICtrlCreateMenu("Liens Perso")

	; Charger les liens depuis le fichier liens.txt
	Global $aliensmenu = ChargerLiens($sFichierlien)

	; Création des menus pour liens.txt
	Global $aSublienMenus1[10][2]
	CreerMenu($liensmenu, $aliensmenu, $aSublienMenus1)

	; Séparateur + item pour modification des logiciels
	GUICtrlCreateMenuItem("", $liensmenu)
	Global $iToOpenLien = GUICtrlCreateMenuItem("Modifier les liens", $liensmenu)


	#EndRegion gestion du menu

	#Region Interface graphique
	Global $PageControl1 = GUICtrlCreateTab(8, 144, 636, 480)
	GUISetBkColor(0xFFFFFF)
	;******************************************************************
	;		Onglet gestion du PC
	;******************************************************************
	Global $TabSheet1 = GUICtrlCreateTabItem("Gestion du PC")
	Global $BT_Infosys = _CreateBouton("Info Système", 27, 242, 113, 41, $FONT_SIZE, $FONT_NAME, $COLOR_GREEN, "Information Système")
	Global $BT_GestinPC = _CreateBouton("Gestion de l'ordinateur", 301, 356, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Gestion de l'ordinateur")
	Global $BTPing = _CreateBouton("Ping", 27, 297, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Ping du PC")
	Global $BT_Whosconnect = _CreateBouton("Qui est connecté" & @CRLF & "(F7)", 27, 185, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Vérifie la session ouverte")
	Global $BT_Printer = _CreateBouton("Imprimante Installées", 164, 413, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Vérifie les imprimantes installées")
	Global $BT_GPupdate = _CreateBouton("GPupdate", 301, 299, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "GPupdate sur le poste sélectionné")
	Global $BT_ValDNS = _CreateBouton("Validation DNS", 164, 573, 116, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_Registre = _CreateBouton("Accés registre", 438, 242, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_CMDDISTANT = _CreateBouton("CMD Distant" & @CRLF & "(F9)", 301, 185, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Lance un CMD Distant")
	Global $BT_C = _CreateBouton("C$ " & @CRLF & "(F10)", 301, 242, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_DockLenovo = _CreateBouton("Info Station Lenovo", 164, 242, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_ScintEcran = _CreateBouton("PB Scintillement écrans", 164, 356, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_SerialEcran = _CreateBouton("Numéro de Série écrans", 164, 299, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_RSOP = _CreateBouton("Résultat GPO", 301, 413, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_Opale = _CreateBouton("Saisie Opale", 438, 299, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_applis = _CreateBouton("Applis installées" & @CRLF & "(Ctrl + A)", 27, 356, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $Bt_GestWifi = _CreateBouton("Gestion Wifi", 438, 185, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Sauvegarde et restauration des paramètres wifi du PC")
	Global $BT_InfoPC = _CreateBouton("Infos PC" & @CRLF & "(F8)", 164, 185, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Affiche les principales informations de l'ordinateur" & @CRLF & "les dernières infos de connexion")
	Global $BT_Dism = _CreateBouton("Réparation fichiers systèmes Windows", 438, 356, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Répare les fichiers systèmes Windows " & @CRLF & "via la commande dism")
	Global $BT_CheckBitlocker = _CreateBouton("Controle Bitlocker", 27, 413, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BTValAlim = _CreateBouton("Validation Alimentation", 27, 573, 116, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BTValWsus = _CreateBouton("Validation WSUS", 301, 573, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BTValProcess = _CreateBouton("Validation Processus", 438, 573, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BTUserLoggin = _CreateBouton("User Loggin", 438, 413, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")

	Global $Label2 = GUICtrlCreateLabel("DNS", 164, 515, 116, 17, $SS_CENTER)
	GUICtrlSetFont(-1, 12, 800, 0, $FONT_NAME)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	Global $CBDNS = GUICtrlCreateCombo("Choix", 164, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	GUICtrlSetData(-1, "Vide le DNS|Force enregistrement")
	GUICtrlSetFont(-1, 9, 400, 0, $FONT_NAME)

	Global $CBAlim = GUICtrlCreateCombo("Choix", 27, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	GUICtrlSetData(-1, "Reboot|Démarrage")
	Global $Label4 = GUICtrlCreateLabel("Alimentation", 27, 515, 116, 17, $SS_CENTER)
	GUICtrlSetFont($Label4, 12, 800, 0, $FONT_NAME)
	GUICtrlSetBkColor($Label4, 0xFFFFFF)
	
	Global $Label5 = GUICtrlCreateLabel("WSUS", 301, 515, 116, 17, $SS_CENTER)
	GUICtrlSetFont(-1, 12, 800, 0, $FONT_NAME)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	Global $CBwsus = GUICtrlCreateCombo("Choix", 301, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	GUICtrlSetData(-1, "Réparation|Force MAJ")
	GUICtrlSetFont(-1, 9, 400, 0, $FONT_NAME)

	Global $Label6 = GUICtrlCreateLabel("Processus", 438, 515, 116, 17, $SS_CENTER)
	GUICtrlSetFont(-1, 12, 800, 0, $FONT_NAME)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	Global $CBProcess = GUICtrlCreateCombo("Choix", 438, 544, 116, 17, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	GUICtrlSetData(-1, "Liste|Arrêt")
	GUICtrlSetFont(-1, 9, 400, 0, $FONT_NAME)



	;***************************************************************************************
	;                     Onglet Gestion AD
	;***************************************************************************************
	Global $TabGestAD = GUICtrlCreateTabItem("Gestion AD")

	$Group2 = GUICtrlCreateGroup("Ordinateur", 24, 176, 265, 273)
	GUICtrlSetFont(-1, 8, 400, 0, "Arial")
	GUICtrlSetBkColor(-1, 0xFFFF00)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Group3 = GUICtrlCreateGroup("Groupes", 328, 184, 305, 145)
	GUICtrlSetFont(-1, 8, 400, 0, "Arial")
	GUICtrlSetBkColor(-1, 0x00FF00)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Group4 = GUICtrlCreateGroup("Utilisateurs", 328, 336, 305, 137)
	GUICtrlSetFont(-1, 8, 400, 0, "Arial")
	GUICtrlSetBkColor(-1, 0xFF0000)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	Global $Bt_GrpADPc = _CreateBouton("Groupes Ordis AD", 29, 220, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Affiche les groupes auxquels appartient le PC")
	Global $Bt_LastlogonAD = _CreateBouton("Dernières connexion AD", 29, 302, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Affiche la dernière connexion à l'AD de l'ordinateur")
	Global $BT_GestionUser = _CreateBouton("Gestion utilisateur sur AD (F6)", 498, 402, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_GestADPC = _CreateBouton("Gestion du PC sur l'AD (F5)", 29, 402, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_ExportversionWin = _CreateBouton("Liste des PC par version de Windows", 162, 302, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Créé un fichier Excel énumérant les ordinateurs de l'AD et leur version de Windows")
	Global $BT_Member = _CreateBouton("Visualiser membres des groupes  AD", 365, 220, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Affiche les membres des groupes GPO ordis ou utilisateurs choisis")

	;**********************************************************************************************
	;                    Onglet Gestion admins Locaux
	;**********************************************************************************************
	Global $TabSheet3 = GUICtrlCreateTabItem("Gestion Groupes Admins")

	Global $BT_AdminLocAD = _CreateBouton("Admin Locaux PC", 230, 205, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Gestion du groupe AD ADMIN_Locaux_PC")
	Global $BT_AdminLoc = _CreateBouton("Gestion groupes admin Locaux", 28, 205, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "Gestion du groupe Administrateur local PC")

	;*********************************************************************************************
	;                  Onglet Gestion Divers
	;*********************************************************************************************
	Global $TabSheet2 = GUICtrlCreateTabItem("Gestion Divers")

	Global $BT_ConecIMP = _CreateBouton("Connexion Imprimante", 32, 192, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $Label3 = GUICtrlCreateLabel("Gestion NAS", 284, 193, 180, 17)
	GUICtrlSetFont(-1, 12, 800, 0, $FONT_NAME)
	Global $CB_NAS = GUICtrlCreateCombo("Choix", 282, 216, 185, 25)
	GUICtrlSetData($CB_NAS, $listeNas)
	Global $BT_ValNAS = _CreateBouton("Connexion au Nas", 282, 248, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_SuiviRZO = _CreateBouton("Suivi Connexions réseaux", 32, 248, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_checkSAV = _CreateBouton("Controle sauvegardes PC", 283, 304, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $Bt_ScanRZO = _CreateBouton("Scan réseau", 32, 304, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $Bt_DepTel = _CreateBouton("Diagnostic téléphonie", 32, 360, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $Bt_update = _CreateBouton("Mise à jour AutoUpdate", 512, 504, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")

	;*********************************************************************************************
	;                  Onglet Gestion Matériel
	;*********************************************************************************************
	Global $TabSheet4 = GUICtrlCreateTabItem("Gestion matériel")

	Global $BT_MAC_PC = _CreateBouton("Cherche PC depuis MAC / SN / Opale", 35, 209, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_Ecran = _CreateBouton("Recherche d'écrans", 171, 209, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $Bt_check_ecran = _CreateBouton("Création liste écrans connectés", 171, 392, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $BT_Dock = _CreateBouton("Recherche station LENOVO", 307, 209, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")
	Global $Bt_check_Dock = _CreateBouton("Création liste station LENOVO", 307, 392, 113, 41, $FONT_SIZE, $FONT_NAME, 0xFF0000, "")


	GUICtrlCreateTabItem("")

	Global $LabUser = GUICtrlCreateLabel($sUsername, 15, 7, 44, 17)
	Global $PLogo = GUICtrlCreateIcon($dlloa, 1, 426, 14, 200, 38)
	;$PLogo = GUICtrlCreatePic($chemin_dossier_bin & "icones\logo.jpg", 426, 14, 200, 38)
	Global $Label1 = GUICtrlCreateLabel("NOM/ IP du PC", 52, 32, 106, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetColor(-1, 0x0066CC)
	Global $addPc = GUICtrlCreateInput("", 12, 64, 217, 21)
	Global $BT_VerifIP = GUICtrlCreateButton("Vérification", 236, 64, 145, 25)
	Global $icoVNC = GUICtrlCreateIcon($dlloa, 3, 416, 64, 32, 32)

	;$BT_VerifIP = GUICtrlCreateButton("Vérification", 236, 64, 145, 25)

	Global $Icohotline = GUICtrlCreateIcon($userdll, 5, 480, 64, 33, 33)
	GUICtrlSetCursor(-1, 0)
	GUICtrlSetTip($Icohotline, "Hotline DSIAN")

	Global $BTFicSSI = GUICtrlCreateButton("Fiche SSI", 356, 104, 105, 25)
	GUICtrlSetTip($BTFicSSI, "Ouvre la fiche SSI de l'ordinateur ou du matériel")
	Global $RadOrdiSSI = GUICtrlCreateRadio("Ordinateur (F2)", 232, 96, 113, 17)
	Global $RadMatSSI = GUICtrlCreateRadio("Matériel (F3)", 232, 120, 113, 17)
	Global $Group1 = GUICtrlCreateGroup("", 224, 88, 257, 49)
	GUICtrlSetBkColor(-1, 0xC0DCC0)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	;$BTFicSSIO = GUICtrlCreateButton("Fiche SSI Ordinateur", 236, 104, 145, 25)
	Global $BT_QUIT = GUICtrlCreateButton("     Quitter", 17, 658, 100, 36)
	GUICtrlSetImage($BT_QUIT, $dlloa, 4, 1)
	GUICtrlSetTip($BT_QUIT, "Quitter")
	;$BT_QUIT = GUICtrlCreateButton("Quitter", 17, 658, 65, 33)
	;GUICtrlSetBkColor(-1, 0xFF0000)
	Global $Cyber = GUICtrlCreateIcon($dlloa, 2, 200, 656, 100, 28)

	GUICtrlSetCursor(-1, 0)
	;$labip = GUICtrlCreateLabel("", 16, 96, 4, 4)
	Global $bt_FlushDNS = GUICtrlCreateButton("Vider DNS local", 469, 658, 113, 33)

	$LAversion = GUICtrlCreateLabel("V : " & $version, 472, 693, 60, 33)

	Global $LaDate = GUICtrlCreateLabel(_NowDate(), 190, 693, 82, 17)
	GUICtrlSetFont($LaDate, 10, 800, 0, "MS Sans Serif")
	Global $LaTime = GUICtrlCreateLabel("Label6", 290, 693, 62, 17)
	GUICtrlSetFont($LaTime, 10, 800, 0, "MS Sans Serif")
	;GUICtrlSetColor(-1, 0x0080FF)
	GUICtrlSetData($LaTime, @HOUR & ":" & @MIN & ":" & @SEC)
	;GUICtrlSetColor(-1, 0x0000FF)

	GUISetState(@SW_SHOW)
EndFunc   ;==>_GuiCreate


