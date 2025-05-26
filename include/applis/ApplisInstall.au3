#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\DLL\Icones\Puzzle.ico
#AutoIt3Wrapper_Res_Fileversion=1.0.0.3
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <Constants.au3>
#include <Date.au3>
#include <String.au3>
#include "..\au3\ficconf.au3"

FileInstall(".\ApplisInstall.au3", $chemin_dossier_bin & "", $FC_OVERWRITE)
$searchString = IniRead($chemin_dossier_bin & "conf\param.flq", "PC", "Nom", "")
$version = @Compiled ? FileGetVersion(@ScriptName) : StringRegExp(FileRead(@ScriptName), "(?i)#AutoIt3Wrapper_Res_Fileversion=(.*?)\h*[;|\v]", 1)[0]
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\CD47\outilsAdmins\infoappli", "Version", "REG_SZ", $version)



GUI()

Exit (0)

Func GUI()
	Global $nameCol = 0
	Global $verCol = 1
	Global $pubCol = 2
	Global $archCol = 3
	Global $msiCol = 4
	Global $guidCol = 5
	Global $uninstCol = 6
	Global $dateCol = 7
	Global $locCol = 8
	Global $srcCol = 9
	Global $keyCol = 10

	Global $computer = ""
	Global $osArch = ""

	Global $GUI = GUICreate("Information Logiciels    V: " & $version, 1000, 500, -1, -1, BitOR($WS_MAXIMIZEBOX, $WS_MINIMIZEBOX, $WS_SYSMENU, $WS_SIZEBOX, $WS_CAPTION), BitOR($WS_EX_STATICEDGE, $WS_EX_WINDOWEDGE))

	Global $iComp = GUICtrlCreateLabel($searchString, 10, 9, 100, 21)
		GUICtrlSetColor($iComp, 0x0000FF)
	Global $rAll = GUICtrlCreateRadio("Tous", 495, 10, 45, 20)
	Global $r32 = GUICtrlCreateRadio("x86", 540, 10, 35, 20)
	Global $r64 = GUICtrlCreateRadio("x64", 580, 10, 35, 20)
	Global $bExport = GUICtrlCreateButton("Export CSV", 783, 5, 90, 28)
	Global $BT_QUIT = GUICtrlCreateIcon($chemin_dossier_bin & "conf\OA.dll", 4 , 920, 5, 32, 32)
	GUICtrlSetTip($BT_QUIT, "Quitter")
	GUICtrlSetState($rAll, $GUI_CHECKED)

	If $CmdLine[0] > 0 And $CmdLine[0] < 3 Then
		If $CmdLine[0] == 2 Then
			GUICtrlSetData($iComp, $CmdLine[2])
		Else
		EndIf
	EndIf

	Global $lResults = GUICtrlCreateLabel("", 8, 485, 500, 20)

	Global $ListView = GUICtrlCreateListView("", 0, 38, 1000, 445, -1, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $WS_EX_CLIENTEDGE))

	_GUICtrlListView_AddColumn($ListView, "Nom", 170)
	_GUICtrlListView_AddColumn($ListView, "Version", 65)
	_GUICtrlListView_AddColumn($ListView, "Fabriquant", 105)
	_GUICtrlListView_AddColumn($ListView, "Arch", 40, 2)
	_GUICtrlListView_AddColumn($ListView, "MSI", 35, 2)
	_GUICtrlListView_AddColumn($ListView, "GUID", 85)
	_GUICtrlListView_AddColumn($ListView, "Commande désinstall", 95)
	_GUICtrlListView_AddColumn($ListView, "Date installation", 70)
	_GUICtrlListView_AddColumn($ListView, "Emplacement", 100)
	_GUICtrlListView_AddColumn($ListView, "Source", 85)
	_GUICtrlListView_AddColumn($ListView, "Clé désinstall", 125)


	Global $rcMenu = GUICtrlCreateContextMenu($ListView)
	Global $rcCopy = GUICtrlCreateMenu("Copie", $rcMenu)
	Global $rcCopyLine = GUICtrlCreateMenuItem("Ligne", $rcCopy)
	Global $rcCopyName = GUICtrlCreateMenuItem("Nom", $rcCopy)
	Global $rcCopyVer = GUICtrlCreateMenuItem("Version", $rcCopy)
	Global $rcCopyGUID = GUICtrlCreateMenuItem("GUID", $rcCopy)
	Global $rcCopyLoc = GUICtrlCreateMenuItem("Emplacement", $rcCopy)
	Global $rcCancel = 0

	_GUICtrlListView_RegisterSortCallBack($ListView)

	GUISetState(@SW_SHOW)

	FillListView()

	While 1
		ReadGUI(GUIGetMsg())
	WEnd
EndFunc   ;==>GUI

Func FillListView()
	Local $i = 0, $c = 0
	Local $key = "", $value = "", $name = ""

	Local $regRoot = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	Local $regRoot32 = "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
	Local $regArch = "HKLM64\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"

	$computer = $searchString
	local $srchTrm  = "" ;GUICtrlRead($iSrch)

	ToggleControlAvailability(0)

	_GUICtrlListView_DeleteAllItems($ListView)
	_GUICtrlListView_BeginUpdate($ListView)

	Local $pingErr = 0
	If $computer <> "" And StringCompare($computer, @ComputerName) <> 0 And StringCompare($computer, "localhost") <> 0 And $computer <> "127.0.0.1" Then
		$regRoot = "\\" & $computer & "\" & $regRoot
		$regRoot32 = "\\" & $computer & "\" & $regRoot32
		$regArch = "\\" & $computer & "\" & $regArch
		GUICtrlSetData($lResults, "Connecting to " & $computer & "...")
		Ping($computer, 1000)
		$pingErr = @error
	Else
		GUICtrlSetData($lResults, "")
	EndIf

	Local $dataBlob = ""

	ToggleRcMenus(1)

	If $pingErr <> 0 Then
		GUICtrlSetData($lResults, "Unable to connect to " & $computer & ".")
	Else
		$osArch = StringUpper(RegRead($regArch, "PROCESSOR_ARCHITECTURE"))

		If $osArch == "AMD64" And GUICtrlRead($r32) == $GUI_CHECKED Then
			$regRoot = $regRoot32
		ElseIf $osArch <> "AMD64" And GUICtrlRead($r64) == $GUI_CHECKED Then
			GUICtrlSetState($rAll, $GUI_CHECKED)
		EndIf

		GUICtrlSetData($lResults, "Searching...")
		Local $hideUpdates = 1

		While 1
			If ReadGUI(GUIGetMsg($GUI)) == False Then
				GUICtrlSetData($lResults, $c & " results found (Search cancelled)")
				ExitLoop
			EndIf
			$i += 1
			$key = $regRoot & "\" & RegEnumKey($regRoot, $i)
			Switch @error
				Case 0
					$name = RegRead($key, "DisplayName")
					If @error == 0 And $name <> "" Then
						$pub = RegRead($key, "Publisher")
						$uninst = RegRead($key, "UninstallString")

						If ($srchTrm == "" Or StringInStr($name, $srchTrm) Or StringInStr($pub, $srchTrm) Or StringInStr($uninst, $srchTrm)) Then
							If $hideUpdates Then
								If Not (StringInStr($pub, "Microsoft", 0, 1) And (StringInStr($name, "Update", 0, 1) Or StringInStr($name, "Hotfix", 0, 1))) Then
									$dataBlob &= BuildLine($ListView, $key, $name, RegRead($key, "DisplayVersion"), $pub, $uninst, RegRead($key, "InstallDate"), RegRead($key, "InstallLocation"), RegRead($key, "InstallSource"))
									If $computer <> "" Then AddListEntry($computer, $dataBlob)
									$c += 1
								EndIf
							Else
								$dataBlob &= BuildLine($ListView, $key, $name, RegRead($key, "DisplayVersion"), $pub, $uninst, RegRead($key, "InstallDate"), RegRead($key, "InstallLocation"), RegRead($key, "InstallSource"))
								If $computer <> "" Then AddListEntry($computer, $dataBlob)
								$c += 1
							EndIf
						EndIf

					EndIf
				Case -1
					If $osArch == "AMD64" And $regRoot <> $regRoot32 And GUICtrlRead($r64) <> $GUI_CHECKED Then
						$regRoot = $regRoot32
						$i = 0
					Else
						GUICtrlSetData($lResults, $c & " result(s) found")
						ExitLoop
					EndIf
				Case 3
					GUICtrlSetData($lResults, "Unable to connect to remote registry.")
					ExitLoop
			EndSwitch
		WEnd
	EndIf

	If $computer == "" Then
		$dataBlob = StringStripWS($dataBlob, $STR_STRIPTRAILING)
		Local $lines = StringSplit($dataBlob, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)
		If $dataBlob <> "" Then
			_ArraySort($lines)
			Local $listArray[UBound($lines)][_GUICtrlListView_GetColumnCount($ListView)]
			For $i = 0 To UBound($lines) - 1
				Local $temp = StringSplit($lines[$i], "|", $STR_NOCOUNT)
				For $j = 0 To UBound($temp) - 1
					$listArray[$i][$j] = $temp[$j]
				Next
			Next
			_GUICtrlListView_DeleteAllItems($ListView)
			_GUICtrlListView_AddArray($ListView, $listArray)
			_GUICtrlListView_SetColumnWidth($ListView, 0, 170)
		EndIf
	Else
		Local $sortSense = False
		_GUICtrlListView_SimpleSort($ListView, $sortSense, 0)
	EndIf

	_GUICtrlListView_EndUpdate($ListView)

	ToggleRcMenus(0)

	ToggleControlAvailability(1)
EndFunc   ;==>FillListView

Func ToggleRcMenus($inProgress)
	If $inProgress Then
		GUICtrlDelete($rcMenu)
		$rcMenu = GUICtrlCreateContextMenu($ListView)
		$rcCancel = GUICtrlCreateMenuItem("Cancel Search", $rcMenu)
	Else

		GUICtrlDelete($rcMenu)

	Global $rcMenu = GUICtrlCreateContextMenu($ListView)
	Global $rcCopy = GUICtrlCreateMenu("Copie", $rcMenu)
	Global $rcCopyLine = GUICtrlCreateMenuItem("Ligne", $rcCopy)
	Global $rcCopyName = GUICtrlCreateMenuItem("Nom", $rcCopy)
	Global $rcCopyVer = GUICtrlCreateMenuItem("Version", $rcCopy)

	Global $rcCopyGUID = GUICtrlCreateMenuItem("GUID", $rcCopy)
	Global $rcCopyLoc = GUICtrlCreateMenuItem("Emplacement", $rcCopy)
EndIf
EndFunc   ;==>ToggleRcMenus

Func ToggleControlAvailability($enable)
	If $enable Then
		GUICtrlSetState($iComp, $GUI_ENABLE)
	GUICtrlSetState($rAll, $GUI_ENABLE)
		GUICtrlSetState($r32, $GUI_ENABLE)
		GUICtrlSetState($r64, $GUI_ENABLE)
		GUICtrlSetState($bExport, $GUI_ENABLE)
	Else
		GUICtrlSetState($iComp, $GUI_DISABLE)
		GUICtrlSetState($rAll, $GUI_DISABLE)
		GUICtrlSetState($r32, $GUI_DISABLE)
		GUICtrlSetState($r64, $GUI_DISABLE)
		GUICtrlSetState($bExport, $GUI_DISABLE)
	EndIf
EndFunc   ;==>ToggleControlAvailability

Func ReadGUI($guiMsg)
	Switch $guiMsg
		Case $rAll To $r64
			FillListView()
		Case $bExport
			ExportCSV()
		Case $ListView
			_GUICtrlListView_SortItems($ListView, GUICtrlGetState($ListView))
		Case $rcCopyLine
			ClipPut(_GUICtrlListView_GetItemTextString($ListView, -1))
		Case $rcCopyName
			ClipPut(_GUICtrlListView_GetItemText($ListView, Int(_GUICtrlListView_GetSelectedIndices($ListView), 1), $nameCol))
		Case $rcCopyVer
			ClipPut(_GUICtrlListView_GetItemText($ListView, Int(_GUICtrlListView_GetSelectedIndices($ListView), 1), $verCol))
		Case $rcCopyGUID
			ClipPut(_GUICtrlListView_GetItemText($ListView, Int(_GUICtrlListView_GetSelectedIndices($ListView), 1), $guidCol))
		Case $rcCopyLoc
			ClipPut(_GUICtrlListView_GetItemText($ListView, Int(_GUICtrlListView_GetSelectedIndices($ListView), 1), $locCol))
		Case $rcCancel
			Return False ; Stop search operation.
		Case $GUI_EVENT_CLOSE, $BT_QUIT
			ExportCSV()
			Exit
	EndSwitch
	Return True ; Continue search operation.
EndFunc   ;==>ReadGUI

Func JumpToDir($dir)
	If $computer <> "" And StringInStr($dir, ":\", 0, 1) Then
		$dir = StringReplace($dir, ":", "$", 1)
		$dir = "\\" & $computer & "\" & $dir
	EndIf
	If Not FileExists($dir) Then
		MsgBox($MB_ICONWARNING, "", "Unable to locate:" & @CRLF & @CRLF & $dir, 10)
	Else
		Run("explorer.exe """ & $dir & """", @SystemDir)
	EndIf
EndFunc   ;==>JumpToDir

Func ExpandKey($key)
	If StringLeft($key, 4) == "HKLM" Then
		$key = StringReplace($key, "HKLM", "HKEY_LOCAL_MACHINE", 1)
	ElseIf StringLeft($key, 4) == "HKCU" Then
		$key = StringReplace($key, "HKCU", "HKEY_CURRENT_USER", 1)
	ElseIf StringLeft($key, 3) == "HKU" Then
		$key = StringReplace($key, "HKU", "HKEY_USERS", 1)
	EndIf
	Return $key
EndFunc   ;==>ExpandKey

Func AddListEntry($computer, ByRef $dataBlob)
	If $dataBlob <> "" Then
		$dataBlob = StringStripWS($dataBlob, $STR_STRIPTRAILING)
		Local $temp = StringSplit($dataBlob, "|", $STR_NOCOUNT)
		Local $lineArray[1][UBound($temp)]
		For $i = 0 To UBound($temp) - 1
			$lineArray[0][$i] = $temp[$i]
		Next
		_GUICtrlListView_AddArray($ListView, $lineArray)
		$dataBlob = ""
	EndIf
EndFunc   ;==>AddListEntry

Func BuildLine($ListView, $key, $name, $ver, $pub, $uninst, $date, $loc, $src)
	Local $guid, $msi, $arch, $temp

	If $date <> "" Then
		$date = _StringInsert($date, "/", 4)
		$date = _StringInsert($date, "/", 7)
		$date = _DateTimeFormat($date, 2)
	EndIf

	If StringInStr($uninst, "msiexec", 0) > 0 Then
		$temp = _StringBetween($uninst, "{", "}")
		$guid = "{" & $temp[0] & "}"
		$msi = "Y"
	Else
		$guid = ""
		$msi = "N"
	EndIf

	If $osArch == "AMD64" Then
		If StringInStr($key, "HKLM\SOFTWARE\WOW6432NODE") > 0 Then
			$arch = "x86"
		Else
			$key = StringReplace($key, "HKLM64", "HKLM", 1)
			$arch = "x64"
		EndIf
	Else
		$key = StringReplace($key, "HKLM64", "HKLM", 1)
		$arch = "x86"
	EndIf

	If $computer <> "" Then
		$key = StringReplace($key, "\\" & $computer & "\", "", 1, 0)
	EndIf

	Return $name & "|" & $ver & "|" & $pub & "|" & $arch & "|" & $msi & "|" & $guid & "|" & $uninst & "|" & $date & "|" & $loc & "|" & $src & "|" & $key & @CRLF
EndFunc   ;==>BuildLine


Func ExportCSV()
    ; Définition du chemin du fichier CSV à enregistrer
    Local $filePath = "\\VULCAIN\fitrans$\dsian\Log\applis\BDD\" & $searchString & ".csv"
    ; Alternative désactivée : boîte de dialogue pour enregistrer le fichier CSV
    ; Local $filePath = FileSaveDialog("Saisir le nom du fichier", @DesktopDir, "CSV(*.csv)", $FD_PROMPTOVERWRITE)

    ; Vérifie si une erreur s'est produite précédemment
    If @error == 0 Then
        ; Vérifie si le fichier a bien l'extension .csv, sinon l'ajoute
        $compare = StringCompare(StringRight($filePath, 4), ".csv", 0)
        If $compare <> 0 Then
            $filePath &= ".csv"
        EndIf

        ; Récupération du nombre de colonnes et de lignes de la ListView
        Local $lvColCount = _GUICtrlListView_GetColumnCount($ListView) - 1
        Local $lvRowCount = _GUICtrlListView_GetItemCount($ListView) - 1
        Local $csvString = ""

        ; Parcours de chaque ligne du ListView
        For $i = 0 To $lvRowCount
            For $j = 0 To $lvColCount
                ; Récupère le texte de la cellule et remplace les guillemets par des guillemets échappés
                $csvString &= '"' & StringReplace(_GUICtrlListView_GetItemText($ListView, $i, $j), '"', '"' & '"') & '"'

                ; Ajoute une virgule entre les colonnes (sauf pour la dernière)
                If $j < $lvColCount Then $csvString &= ','
            Next

            ; Ajoute un retour à la ligne entre chaque ligne (sauf pour la dernière)
            If $i < $lvRowCount Then $csvString &= @CRLF
        Next

        ; Ouvre le fichier en mode écriture (10 = création + écrasement si existant)
        Local $fileHandle = FileOpen($filePath, 10)
        If $fileHandle <> -1 Then
            ; Définit l'en-tête du fichier CSV avec les noms des colonnes
            Local $header = "Nom,Version,Fabriquant,Arch,MSI,GUID,Commande desinstall,Date Installation,Source,emplacement,clé désinstall"

            ; Écrit l'en-tête suivi des données extraites du ListView
            FileWrite($fileHandle, $header & @CRLF & StringStripWS($csvString, $STR_STRIPTRAILING))

            ; Ferme le fichier
            FileClose($fileHandle)

            ; Affiche un message de confirmation
            ;MsgBox($MB_ICONINFORMATION, "", "Export Effectué", 10)
        Else
            ; Affiche un message d'erreur en cas d'échec de l'ouverture du fichier
            MsgBox(16, "", "Export échoué", 10)
        EndIf
    EndIf
EndFunc   ;==>ExportCSV
