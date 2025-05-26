#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <StructureConstants.au3>
#include <WindowsConstants.au3>

;===============================================================
; Fonction _GuiEditIni
;
; avec l'appel des fonctions WM_NOTIFY et _GuiEDITiniSet
; permet d'afficher dans une Gui les informations d'un fichier ini
;
; Ce fichier doit respecter la norme d'un fichier ini
; [Section]
; clé=Valeur
;
;
; Appel de la fonction _GuiEDITini(NomDuFichierDeConf)
;=================================================================

Global $ArrayCrypt, $hListView, $ini_configure, $iC, $iS

Func _GuiEDITini($ini_configure)
    ; Déclaration des variables locales pour la taille et la position de la fenêtre
    Local $w = 800, $h = 600, $int = 3

    ; Création de la fenêtre GUI
    GUICreate("Fichier de configuration", $w, $h)

    ; Création d'une ListView pour afficher les sections et les clés du fichier INI
    $hListView = GUICtrlCreateListView("", $int, $int, $w - 2 * $int, $h - 2 * $int, BitOR($LVS_EDITLABELS, $LVS_REPORT))
    _GUICtrlListView_SetUnicodeFormat($hListView, False) ; Utilisation du format ANSI

    ; Activation de l'état de la GUI
    GUISetState()

    ; Ajout des colonnes à la ListView
    _GUICtrlListView_InsertColumn($hListView, 0, "Valeur", 1200)
    _GUICtrlListView_InsertColumn($hListView, 1, "Site", $w / 2.7)
    _GUICtrlListView_SetColumnOrder($hListView, "1|0")
    _GUICtrlListView_SetColumnWidth($hListView, 0, $LVSCW_AUTOSIZE_USEHEADER)
    _GUICtrlListView_SetColumnWidth($hListView, 0, _GUICtrlListView_GetColumnWidth($hListView, 0) - 20)

    ; Activation de l'affichage des groupes dans la ListView
    _GUICtrlListView_EnableGroupView($hListView)

    ; Lecture des sections du fichier INI
    $sections = IniReadSectionNames($ini_configure)

    ; Vérification des erreurs lors de la lecture des sections
    If @error Then
        MsgBox(16, "Erreur", "Erreur lors de la lecture des sections : " & @error)
    ElseIf IsArray($sections) Then
        ; Parcours des sections lues
        For $i = 1 To $sections[0]
            ; Ajout de chaque section en tant que groupe dans la ListView
            _GUICtrlListView_InsertGroup($hListView, $iS, $iS, $sections[$i])
            $array = IniReadSection($ini_configure, $sections[$i])
            If IsArray($array) Then
                ; Parcours des clés dans chaque section
                For $j = 1 To $array[0][0]
                    Local $iIndex = _GUICtrlListView_AddItem($hListView, $array[$j][1])
_GUICtrlListView_AddSubItem($hListView, $iIndex, $array[$j][0], 1)
_GUICtrlListView_SetItemGroupID($hListView, $iIndex, $iS)
                    ; Gestion des clés cryptées
                    If $sections[$i] = "Crypt" Then
                        ReDim $ArrayCrypt[$array[0][0]]
                        $ArrayCrypt[$j - 1] = $iC
                    EndIf
                    $iC += 1
                Next
                $iS += 1
            Else
                MsgBox(0, "", "Erreur de lecture des paramètres dans la section : " & $sections[$i])
                Return 0
            EndIf
        Next
    Else
        MsgBox(0, "", "Erreur de lecture du fichier de configuration")
        Return 0
    EndIf

    ; Enregistrement des messages de notification pour la ListView
    GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

    ; Boucle principale de la GUI
    Do
        Sleep(20)
    Until GUIGetMsg() = $GUI_EVENT_CLOSE

    ; Suppression de la GUI
    GUIDelete()
    Return 1
EndFunc   ;==>_GuiEDITini

Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)
    ; Forcer la référence des variables pour éviter les avertissements de compilation
    #forceref $hWnd, $iMsg, $iwParam

    ; Déclaration des variables locales
    Local $hWndFrom, $iCode, $tNMHDR, $hWndListView, $tInfo

    ; Récupération du handle de la ListView
    $hWndListView = $hListView
    If Not IsHWnd($hListView) Then $hWndListView = GUICtrlGetHandle($hListView)

    ; Création d'une structure pour stocker les informations de l'en-tête de notification
    $tNMHDR = DllStructCreate($tagNMHDR, $ilParam)

    ; Récupération des informations de la structure
    $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
    $iIDFrom = DllStructGetData($tNMHDR, "IDFrom")
    $iCode = DllStructGetData($tNMHDR, "Code")

    ; Vérification de la source de la notification
    Switch $hWndFrom
        Case $hWndListView
            ; Gestion des différents codes de notification pour la ListView
            Switch $iCode
                ; Début de l'édition d'une étiquette d'un élément
                Case $LVN_BEGINLABELEDITA, $LVN_BEGINLABELEDITW
                    $tInfo = DllStructCreate($tagNMLVDISPINFO, $ilParam)
                    $ID = _ArraySearch($ArrayCrypt, Int(DllStructGetData($tInfo, "Item")))
                    If $ID <> -1 Then
                        ; Si l'élément est crypté, demander une nouvelle valeur
                        $rep = InputBox("", "Ce champ est crypté. Veuillez donner la nouvelle valeur ci-dessous.")
                        If $rep <> "" Then
                            _GUICtrlListView_SetItemText($hListView, Int(DllStructGetData($tInfo, "Item")), "12*/#Max987²", 2)
                            _GUICtrlListView_CancelEditLabel($hListView)
                            _GuiEDITiniSet(Int(DllStructGetData($tInfo, "Item")))
                        EndIf
                        Return False
                    EndIf

                ; Fin de l'édition d'une étiquette d'un élément
                Case $LVN_ENDLABELEDITA, $LVN_ENDLABELEDITW
                    $tInfo = DllStructCreate($tagNMLVDISPINFO, $ilParam)
                    Local $tBuffer = DllStructCreate("char Text[" & DllStructGetData($tInfo, "TextMax") & "]", DllStructGetData($tInfo, "Text"))
                    If StringLen(DllStructGetData($tBuffer, "Text")) Then
                        _GuiEDITiniSet(Int(DllStructGetData($tInfo, "Item")), DllStructGetData($tBuffer, "Text"))
                        Return True
                    EndIf

                ; Double-clic sur un élément de la ListView
                Case $NM_DBLCLK
                    $tInfo = DllStructCreate($tagNMITEMACTIVATE, $ilParam)
                    _GUICtrlListView_EditLabel($hListView, Int(DllStructGetData($tInfo, "Index")))
                    $ID = _ArraySearch($ArrayCrypt, Int(DllStructGetData($tInfo, "Index")))
                    If $ID <> -1 Then
                        ; Si l'élément est crypté, demander une nouvelle valeur
                        $rep = InputBox("", "Ce champ est crypté. Veuillez donner la nouvelle valeur ci-dessous.", $hListView, "12*/#Max987²", 2)
                        If $rep <> "" Then
                            _GUICtrlListView_SetItemText($hListView, $rep, "12*/#Max987²", 2)
                            _GUICtrlListView_CancelEditLabel($hListView)
                            _GuiEDITiniSet(Int(DllStructGetData($tInfo, "Index")))
                        EndIf
                    EndIf
            EndSwitch
    EndSwitch

    ; Retourne le message par défaut pour continuer le traitement
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

Func _GuiEDITiniSet($ID, $b = "")
    ; Récupération des informations du groupe de l'élément modifié
    $a = _GUICtrlListView_GetGroupInfo($hListView, _GUICtrlListView_GetItemGroupID($hListView, $ID))

    ; Activation du format Unicode pour la ListView
    _GUICtrlListView_SetUnicodeFormat($hListView, True)

    ; Si aucune nouvelle valeur n'est fournie, utiliser la valeur actuelle de l'élément
    If $b = "" Then $b = _GUICtrlListView_GetItemText($hListView, $ID)

    ; Récupération du texte de la sous-item (clé) de l'élément
    $c = _GUICtrlListView_GetItemText($hListView, $ID, 1)

    ; Désactivation du format Unicode pour la ListView
    _GUICtrlListView_SetUnicodeFormat($hListView, False)

    ; Écriture de la nouvelle valeur dans le fichier INI
    IniWrite($ini_configure, $a[0], $c, $b)
EndFunc   ;==>_GuiEDITiniSet
