#include <AD.au3>
#include <APIShellExConstants.au3>
#include <Excel.au3>
#include <ExcelConstants.au3>
#include <File.au3>
#include <Word.au3>

;===============================================================================
;
; Function Name:    _FileCopy()
; Description:      Copies Files and/or Folders with optional animated Windows dialog box
; Parameter(s):     $pFrom - Source files/folders to copy from
;                           $pTo - Destination folder to copy to
;                           $fFlags - [optional] Defines the flag for CopyHere action
;
; Option Flags:        $FOF_DEFAULT = 0          Default. No options specified.
;                       $FOF_SILENT = 4              Do not display a progress dialog box.
;                       $FOF_RENAMEONCOLLISION = 8 Rename the target file if a file exists at the target location with the same name.
;                       $FOF_NOCONFIRMATION = 16    Click "Yes to All" in any dialog box displayed.
;                       $FOF_ALLOWUNDO = 64          Preserve undo information, if possible.
;                       $FOF_FILESONLY = 128          Perform the operation only if a wildcard file name (*.*) is specified.
;                       $FOF_SIMPLEPROGRESS = 256   Display a progress dialog box but do not show the file names.
;                       $FOF_NOCONFIRMMKDIR = 512   Do not confirm the creation of a new directory if the operation requires one to be created.
;                       $FOF_NOERRORUI = 1024          Do not display a user interface if an error occurs.
;                       $FOF_NORECURSION = 4096     Disable recursion.
;                       $FOF_SELECTFILES = 9182 Do not copy connected files as a group. Only copy the specified files.
;
; Requirement(s):   None.
; Return Value(s):  No return values per MSDN article.
; Author(s):        Jos, commented & modified by ssubirias3 to create $pTo if not present
; Note(s):          Must declare $fFlag variables or use their numeric equivalents without declaration. Use BitOR() to combine
;                       Flags. See http://support.microsoft.com/kb/151799 &
;                       http://msdn2.microsoft.com/en-us/library/ms723207.aspx for more details
;
;===============================================================================


Func _FileCopy($pFrom, $pTo, $fFlags = 0)
    Local $FOF_NOCONFIRMMKDIR = 512
    Local $FOF_NOCONFIRMATION = 16
    If Not FileExists($pTo) Then DirCreate($pTo)
    $winShell = ObjCreate("shell.application")
    ;$winShell.namespace($pTo).CopyHere($pFrom, BitOR(272))
	$winShell.namespace($pTo).CopyHere($pFrom, BitOR($FOF_NOCONFIRMMKDIR,$FOF_NOCONFIRMATION))
EndFunc

Func _ExportVWin()
	_AD_Open()

	; Requête LDAP pour obtenir les informations nécessaires
	Local $aComputers = _AD_GetObjectsInOU("OU=CLIENTS CG47,DC=dptlg,DC=fr", "(&(objectCategory=computer)(name=*))", 2, "name,lastLogonTimestamp,operatingSystem,operatingSystemVersion")

	; Vérifier si la requête a réussi
	If @error Then
		MsgBox(16, "Erreur", "Impossible de récupérer les informations des ordinateurs.")
		Exit
	EndIf

	; Créer un nouvel objet Excel
	Local $oExcel = _Excel_Open()
	If @error Then
		MsgBox(16, "Erreur", "Impossible d'ouvrir Excel.")
		Exit
	EndIf

	; Créer un nouveau classeur
	Local $oWorkbook = _Excel_BookNew($oExcel)
	If @error Then
		MsgBox(16, "Erreur", "Impossible de créer un nouveau classeur.")
		_Excel_Close($oExcel)
		Exit
	EndIf

	; Écrire l'en-tête du fichier Excel
	_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "Nom, Version de Windows, Version de l'OS", "A1")

	; Parcourir les résultats et écrire dans le fichier Excel
	For $i = 1 To UBound($aComputers) - 1
		Local $sName = $aComputers[$i][0]
		;Local $sLastLogon = _ConvertLDAPTime($aComputers[$i][1])
		Local $sOS = $aComputers[$i][2]
		Local $sOSVersion = $aComputers[$i][3]
		_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $sName & "," & $sOS & "," & $sOSVersion, "A" & ($i + 1))
	Next

	; Enregistrer le fichier Excel
	_Excel_BookSaveAs($oWorkbook, "c:\temp\ListeOS.xlsx", $xlWorkbookDefault, True)
	_Excel_Close($oExcel)

	_AD_Close()

	_Excel_BookOpen(_Excel_Open(), "c:\temp\ListeOS.xlsx")

	MsgBox(64, "Rappel des Versions", "Version attendue 22H2" & @CRLF & "" & @CRLF & "22H2 = 19045" & @CRLF & "22H1 = 19044" & @CRLF & "21H2 = 19043" & @CRLF & "21H1 = 19042")


EndFunc   ;==>_ExportVWin

Func _GetUserName($IPPc)
	Local $objWMIService, $objItem, $colItems, $strUser, $strDomain
	$objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & $IPPc)
	$colItems = $objWMIService.InstancesOf("Win32_Process")
	If IsObj($colItems) Then
		For $objItem In $colItems
			If ($objItem.Caption = "explorer.exe") Then
				$Result = $objItem.GetOwner($strUser, $strDomain)
				If (Not @error) And ($Result = 0) Then Return $strUser
			EndIf
		Next
	EndIf

EndFunc   ;==>_GetUserName



Func WMIDateStringToDate($dtmDate)
	Return (StringMid($dtmDate, 9, 2) & "/" & _
			StringMid($dtmDate, 6, 2) & "/" & StringLeft($dtmDate, 4) _
			 & " " & StringMid($dtmDate, 12, 2) & ":" & StringMid($dtmDate, 15, 2))
	;& ":" & StringMid($dtmDate,13, 2))
EndFunc   ;==>WMIDateStringToDate

Func OnAutoItExit()
	TCPShutdown() ; Ferme le service TCP.
EndFunc   ;==>OnAutoItExit


#Region Gestion du menu perso via le fichier TXT

Func ChargerLiens($sFichierlien)
    Local $aLiens[0][4] ; Tableau vide initialement avec 4 colonnes
    Local $sFile = $sFichierlien  ; Le chemin vers le fichier contenant les liens

    ; Vérification si le fichier existe
    If Not FileExists($sFile) Then
        MsgBox(0, "Erreur", "Le fichier liens.txt n'a pas été trouvé.")
        Return $aLiens
    EndIf

    Local $sSection = ""  ; Variable pour identifier les sections (par exemple [CD47])
    Local $i = 0  ; Indice pour ajouter de nouvelles lignes au tableau

    ; Lecture du fichier ligne par ligne
    For $line In FileReadToArray($sFile)
        $line = StringStripWS($line, 8)

        ; Si la ligne commence par une section, on la sauvegarde
        If StringLeft($line, 1) == "[" Then
            $sSection = StringTrimLeft($line, 1)

        ; Si la ligne contient un lien (symbole |), on la traite
        ElseIf StringInStr($line, "|") Then
            Local $aParts = StringSplit($line, "|")

            ; Si la ligne contient bien des parties séparées par "|"
            If UBound($aParts) > 1 Then
                ; Vérifier si la ligne contient bien deux éléments après la séparation
                If StringLen($aParts[1]) > 0 And StringLen($aParts[2]) > 0 Then
                    ; Redimensionner le tableau pour ajouter une nouvelle ligne
                    ReDim $aLiens[$i + 1][4]  ; On augmente la taille du tableau de 1 ligne avec 4 colonnes

                    ; Ajouter les données dans la nouvelle ligne
                    $aLiens[$i][0] = $aParts[1]   ; Nom du lien
                    $aLiens[$i][1] = $aParts[2]   ; URL
                    $aLiens[$i][2] = $sSection    ; Section (optionnelle)

                    ; Incrémenter l'indice pour la prochaine ligne
                    $i += 1
                EndIf
            EndIf
        EndIf
    Next

    Return $aLiens
EndFunc

Func CreerMenu($hMainMenu, ByRef $aLiens, ByRef $aSubMenus)
    Local $iMax = UBound($aSubMenus)

    For $i = 0 To UBound($aLiens) - 1
        Local $sCategorie = StringTrimRight($aLiens[$i][2], 1) ; Supprimer le crochet fermant ']'
        Local $iIndex = -1

        ; Vérifier si la catégorie existe déjà
        For $j = 0 To UBound($aSubMenus) - 1
            If $aSubMenus[$j][0] == $sCategorie Then
                $iIndex = $j
                ExitLoop
            EndIf
        Next

        ; Si la catégorie n'existe pas, l'ajouter
        If $iIndex == -1 Then
            ; Si on atteint la limite, agrandir le tableau
            If UBound($aSubMenus) = $iMax Then
                $iMax += 10 ; Augmenter la capacité par blocs de 10
                ReDim $aSubMenus[$iMax][2]
            EndIf

            ; Ajouter la nouvelle catégorie
            $iIndex = UBound($aSubMenus) - 1
            $aSubMenus[$iIndex][0] = $sCategorie
            $aSubMenus[$iIndex][1] = GUICtrlCreateMenu($sCategorie, $hMainMenu)
        EndIf

        ; Créer l'élément du menu et stocker l'identifiant
        $aLiens[$i][3] = GUICtrlCreateMenuItem($aLiens[$i][0], $aSubMenus[$iIndex][1])  ; Nom du lien
    Next
EndFunc

Func VerifierLiencase($msg, ByRef $aLiensAssocies)
    If $msg = 0 Then Return False  ; Évite de traiter les événements non pertinents

    For $i = 0 To UBound($aLiensAssocies) - 1
        If $msg == $aLiensAssocies[$i][3] Then
            Local $filePath = $aLiensAssocies[$i][1]
            Local $fileExtension = StringLower(StringRight($filePath, 5))

            ; Vérifier l'extension du fichier et ouvrir avec l'application appropriée
            Switch $fileExtension
                Case ".docx" , ".doc"
                   _Word_DocOpen(_Word_Create(), $filePath)
                Case ".xlsx" , ".xls" , ".csv"
                    _Excel_BookOpen(_Excel_Open(), $filePath)
                Case Else
                    ShellExecute($filePath)  ; Ouvrir le lien avec l'application par défaut
            EndSwitch
            Return True
        EndIf
    Next
    Return False
EndFunc
 #EndRegion

 Func _CheckAndUpdateComponent($sComponentName, $sExeName)
    ; Décrypte temporairement le fichier de versions

    cryptage($Versionini, 2, $Key)

    ; Compare version distante vs locale
    Local $sDistantVersion = IniRead($Versionini, $sComponentName, "version", "")
    Local $sLocalVersion = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\" & $sComponentName, "Version")

    ; Met à jour si nécessaire
    If $sDistantVersion <> $sLocalVersion Then
        _FileWriteLog($Cheminlog, "Mise à jour du composant " & $sComponentName & " (" & $sLocalVersion & " -> " & $sDistantVersion & ")")
        _FileCopy($Sources & $sExeName, $chemin_dossier_bin & "")
        RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\" & $sComponentName, "Version", "REG_SZ", $sDistantVersion)
    EndIf

    ; Re-crypte le fichier de versions
    cryptage($Versionini, 1, $Key)

    ; Lance l'application
    _FileWriteLog($Cheminlog, "Lancement de " & $sExeName)
    Return RunAs($sUsername, $Domain, $sPassword, 0, $chemin_dossier_bin & $sExeName)
 EndFunc

 Func SetLastLoggedOnUser($sComputerName)
    ; Demande le nom d'utilisateur
    Local $sNomUtilisateur = InputBox("Dernier utilisateur connecté", "Entrez le nom de l'utilisateur (ex: JeDupont)")
    If @error Or $sNomUtilisateur = "" Then
        MsgBox(48, "Annulé", "Aucun utilisateur saisi.")
        Return
    EndIf

    ; Récupère le domaine local ou défini manuellement
    Local $sDomaine = "dptlg"

    ; Clé cible
    Local $sKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"

    ; Construit les commandes REG
    Local $sCmd1 = 'reg add "' & $sKey & '" /v "LastLoggedOnSAMUser" /t REG_SZ /d "' & $sDomaine & '\' & $sNomUtilisateur & '" /f'
    Local $sCmd2 = 'reg add "' & $sKey & '" /v "LastLoggedOnUser" /t REG_SZ /d "' & $sDomaine & '\' & $sNomUtilisateur & '" /f'
    Local $sCmd3 = 'reg add "' & $sKey & '" /v "LastLoggedOnDisplayName" /t REG_SZ /d "' & $sNomUtilisateur & '" /f'

    ; Exécute les commandes sur le PC distant via PsExec
    RunWait(@ComSpec & ' /c ' & $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $sComputerName & ' -s ' & $sCmd1, "", @SW_HIDE)
    RunWait(@ComSpec & ' /c ' & $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $sComputerName & ' -s ' & $sCmd2, "", @SW_HIDE)
    RunWait(@ComSpec & ' /c ' & $chemin_dossier_bin & "PsExec.exe /accepteula \\" & $sComputerName & ' -s ' & $sCmd3, "", @SW_HIDE)

    MsgBox(64, "Succès", "Dernier utilisateur mis à jour sur " & $sComputerName)
EndFunc
