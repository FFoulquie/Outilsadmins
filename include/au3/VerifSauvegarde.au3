Func CheckSAVPC($SiteD)

	SplashTextOn("Traitement en cours", "Merci de  patienter", 250, 120, -1, -1, 48, "Cambria", "", 800)
	_FileWriteLog($Cheminlog, "Fonction CheckSavPC")

	cryptage($configini, 2, $Key)
	_FileWriteLog($Cheminlog, "decryptage de  " & $configini & " code erreur " & @error)

	_FileWriteLog($Cheminlog, "emplacement config " & $configini)
	Global $NomSite = IniRead($configini, $SiteD, "filtre", "")
	Global $NasSite = IniRead($configini, $SiteD, "Backup", "")

	cryptage($configini, 1, $Key)
	_FileWriteLog($Cheminlog, "Chek sauvegarde de " & $NomSite)
	Sleep(100)

	; Chemin du fichier texte pour sauvegarder la liste
	Local $PathFilePC = "c:\temp\liste_pc2.txt"

	; Ouvrir le fichier de sortie en écriture
	Local $PathFilePCHandle = FileOpen($PathFilePC, 2)

	; Vérifier si le fichier de sortie a pu être ouvert
	If $PathFilePCHandle = -1 Then
		SplashOff()
		MsgBox(16, "Erreur", "Impossible d'ouvrir le fichier de sortie.")
		Exit
	EndIf

	; Ouvrir la connexion à l'Active Directory
	_AD_Open()

	; Récupérer les objets ordinateurs
	Local $aComputers = _AD_GetObjectsInOU("OU=W10,OU=CLIENTS CG47,DC=dptlg,DC=fr", "(&(objectClass=computer)(sAMAccountName=*" & $NomSite & "*))", 2, "name")

	_FileWriteLog($Cheminlog, "Nombre de PC trouvés " & UBound($aComputers))
	; Vérifier s'il y a des ordinateurs correspondants
	If UBound($aComputers) = 0 Then
		SplashOff()
		MsgBox(64, "Information", "Aucun ordinateur trouvé.")
	Else
		; Écrire les noms d'ordinateurs dans la console
		For $i = 1 To UBound($aComputers) - 1

			FileWriteLine($PathFilePCHandle, $aComputers[$i])
		Next
	EndIf

	; Fermer le handle du fichier de sortie
	FileClose($PathFilePCHandle)


	; Lire le chemin du fichier des PC
	Local $inputFilePath = $PathFilePC

	; Vérifier si le fichier texte existe
	If Not FileExists($inputFilePath) Then
		SplashOff()
		MsgBox(16, "Erreur", "Le fichier liste n'existe pas.")
		Exit
	EndIf

	; Ouvrir le fichier texte en lecture
	Local $fileHandle = FileOpen($inputFilePath, 0)

	; Vérifier si le fichier a pu être ouvert
	If $fileHandle = -1 Then
		SplashOff()
		MsgBox(16, "Erreur", "Impossible d'ouvrir le fichier liste")
		Exit
	EndIf

	; Chemin du fichier de résultats
	Local $outputFilePath = "c:\temp\suiviSAV.txt"

	; Ouvrir le fichier de résultats en écriture
	Local $outputFileHandle = FileOpen($outputFilePath, 2)

	; Vérifier si le fichier de résultats a pu être ouvert
	If $outputFileHandle = -1 Then
		SplashOff()
		MsgBox(16, "Erreur", "Impossible d'ouvrir le fichier de résultats.")
		Exit
	EndIf

	; Lire le contenu du fichier texte ligne par ligne
	While 1
		Local $line = FileReadLine($fileHandle)
		; Sortir de la boucle si la fin du fichier est atteinte
		If @error = -1 Or $line = "" Then ExitLoop

		; Construire le chemin complet du sous-répertoire
		Local $subDirPath = "\\nas-" & $NasSite & "\*" & $line & "*"
		Local $subDirPath2 = "\\nas-" & $NasSite & "\W10\*" & $line & "*"
		Local $subDirPath3 = "\\nas-" & $NasSite & "\CloudStation\*" & $line & "*"
		Local $subDirPath4 = "\\nas-" & $NasSite & "\Synology Drive\*" & $line & "*"
		Local $subDirPath5 = "\\nas-" & $NasSite & "\SynologyDrive\*" & $line & "*"
		Local $subDirPath6 = "\\nas-" & $NasSite & "\Backups\*" & $line & "*"
		Local $subDirPath7 = "\\nas-" & $NasSite & "\SynolgyDrive\*" & $line & "*"
		; Vérifier si le sous-répertoire existe
		If Not FileExists($subDirPath) Then
			;ConsoleWrite("Pas 1" & @CRLF)
			If Not FileExists($subDirPath2) Then
				;ConsoleWrite("Pas 2" & @CRLF)
				If Not FileExists($subDirPath3) Then
					;ConsoleWrite("Pas 3" & @CRLF)
					If Not FileExists($subDirPath4) Then
						;ConsoleWrite("Pas 4" & @CRLF)
						If Not FileExists($subDirPath5) Then
							If Not FileExists($subDirPath6) Then
								If Not FileExists($subDirPath7) Then
									;ConsoleWrite("Pas 5" & @CRLF)
									; Récupération de la dernière connexion du PC
									$aProperties = _AD_GetObjectProperties($line & "$", "lastlogon")

									; Vérification des erreurs
									If @error Then
										MsgBox($MB_OK, "Erreur", "Erreur lors de la récupération des propriétés : " & @error)
									Else
									EndIf

									; Vérification si la propriété lastlogon existe
									If IsArray($aProperties) And _ArraySearch($aProperties, "lastlogon", 0, 0, 0, 1) <> -1 Then
										; Affichage de la dernière connexion
										FileWriteLine($outputFileHandle, $line & ", n'est pas sauvegardé. Sa dernière connexion AD est le : " & $aProperties[1][1])
									Else
										; La propriété lastlogon n'est pas renseignée
										FileWriteLine($outputFileHandle, $line & ", n'est pas sauvegardé. Sa dernière connexion AD est le : ")
									EndIf

									;;ConsoleWrite($line & " La propriété lastlogon n'est pas renseignée." & @CRLF)
								Else
									; Chemin du répertoire à inspecter
									Local $sFolderPath = "\\nas-" & $NasSite & "\SynolgyDrive\"
									;ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
									; Obtenir la liste des sous-dossiers de PC
									Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)
									; Vérifier s'il y a eu une erreur lors de la récupération
									If @error Then
										;
										FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
									Else
										For $u = 1 To $aSubfolders[0]
											$Cduser = $sFolderPath & $aSubfolders[$u] & "\c\CD47Users"

											;ConsoleWrite("cdusers " & $Cduser & @CRLF)
											Local $aNomSession = _FileListToArray($Cduser, "*", 2)

											;_FileListToArray(

											If @error = 1 Then

												FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $Cduser)
												ExitLoop
											EndIf
											; Afficher la liste des sous-dossiers
											For $i = 1 To $aNomSession[0]

												;date de modification du fichier de données outlook

												$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
												If @error Then
													;
													FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du Log pour " & $aNomSession[$i])
												Else
													$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
													;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
													;date de dernière connexion de l'utilisateur en cours
													$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
													If @error <> 0 Then
														FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
														;ConsoleWrite(@error & @CRLF)
														ExitLoop
													EndIf
													;suppression de l'heure dans la date de connexion
													$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)


													$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
													If $nbr_jours_de_difference > 20 Then
														FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
													EndIf

												EndIf

											Next
										Next
									EndIf
								EndIf

							Else
								; Chemin du répertoire à inspecter
								Local $sFolderPath = "\\nas-" & $NasSite & "\Backups\"
								;ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
								; Obtenir la liste des sous-dossiers de PC
								Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)
								; Vérifier s'il y a eu une erreur lors de la récupération
								If @error Then
									;
									FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
								Else
									For $u = 1 To $aSubfolders[0]
										$Cduser = $sFolderPath & $aSubfolders[$u] & "\c\CD47Users"

										;ConsoleWrite("cdusers " & $Cduser & @CRLF)
										Local $aNomSession = _FileListToArray($Cduser, "*", 2)

										;_FileListToArray(

										If @error = 1 Then

											FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $Cduser)
											ExitLoop
										EndIf
										; Afficher la liste des sous-dossiers
										For $i = 1 To $aNomSession[0]

											;date de modification du fichier de données outlook

											$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
											If @error Then
												;
												FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du Log pour " & $aNomSession[$i])
											Else
												$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
												;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
												;date de dernière connexion de l'utilisateur en cours
												$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
												If @error <> 0 Then
													FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
													;ConsoleWrite(@error & @CRLF)
													ExitLoop
												EndIf
												;suppression de l'heure dans la date de connexion
												$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)


												$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
												If $nbr_jours_de_difference > 20 Then
													FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
												EndIf

											EndIf

										Next
									Next
								EndIf
							EndIf
						Else

							; Vérifier si le sous-répertoire existe
							If Not FileExists($subDirPath5) Then
								;ConsoleWrite("pastrouvé" & @CRLF)
							Else
								;ConsoleWrite($subDirPath5 & @CRLF)
								; Chemin du répertoire à inspecter
								Local $sFolderPath = "\\nas-" & $NasSite & "\SynologyDrive\"
								;ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
								; Obtenir la liste des sous-dossiers de PC
								Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)


								; Vérifier s'il y a eu une erreur lors de la récupération
								If @error Then
									;
									FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
								Else
									For $u = 1 To $aSubfolders[0]
										$Cduser = $sFolderPath & $aSubfolders[$u] & "\c\CD47Users"

										;ConsoleWrite("cdusers " & $Cduser & @CRLF)
										Local $aNomSession = _FileListToArray($Cduser, "*", 2)

										;_FileListToArray(

										If @error = 1 Then

											FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $subDirPath5)
											ExitLoop
										EndIf
										; Afficher la liste des sous-dossiers
										For $i = 1 To $aNomSession[0]

											;date de modification du fichier de données outlook

											$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
											If @error Then
												;
												FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du Log pour " & $aNomSession[$i])
											Else
												$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
												;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
												;date de dernière connexion de l'utilisateur en cours
												$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
												If @error <> 0 Then
													FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
													;ConsoleWrite(@error & @CRLF)
													ExitLoop
												EndIf
												;suppression de l'heure dans la date de connexion
												$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)


												$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
												If $nbr_jours_de_difference > 20 Then
													FileWriteLine($outputFileHandle, "le log de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
												EndIf

											EndIf

										Next
									Next
								EndIf
							EndIf
						EndIf
						; Fin $subDirPath5
					Else

						; Vérifier si le sous-répertoire existe
						If Not FileExists($subDirPath4) Then
							;ConsoleWrite("pastrouvé" & @CRLF)
						Else
							;ConsoleWrite($subDirPath4 & @CRLF)
							; Chemin du répertoire à inspecter
							Local $sFolderPath = "\\nas-" & $NasSite & "\Synology Drive\"
							;ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
							; Obtenir la liste des sous-dossiers de PC
							Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)

							; Vérifier s'il y a eu une erreur lors de la récupération
							If @error Then
								;
								FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
							ElseIf $aSubfolders = 0 Then
								FileWriteLine($outputFileHandle, $sFolderPath & " erreur dans la structure des sous dossiers ")
							Else
								For $u = 1 To $aSubfolders[0]
									$Cduser = $sFolderPath & $aSubfolders[$u] & "\c\CD47Users\"

									Local $aNomSession = _FileListToArray($Cduser, "*", 2)
									;;ConsoleWrite(@error &@CRLF)
									;_FileListToArray(

									If @error = 1 Then
										FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $subDirPath4)
										ExitLoop
									EndIf



									; Afficher la liste des sous-dossiers
									For $i = 1 To $aNomSession[0]
										;ConsoleWrite($aNomSession[$i] & @CRLF)
										;date de modification du fichier de données outlook

										$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
										If @error Then
											;
											FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du log pour " & $aNomSession[$i])
										Else
											$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
											;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
											;date de dernière connexion de l'utilisateur en cours
											$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
											If @error <> 0 Then
												FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
												;ConsoleWrite(@error & @CRLF)
												ExitLoop
											EndIf
											;suppression de l'heure dans la date de connexion
											$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)


											$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
											If $nbr_jours_de_difference > 20 Then
												FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
											EndIf

										EndIf

									Next
								Next
							EndIf
						EndIf
					EndIf
					; Fin $subDirPath4
				Else

					; Vérifier si le sous-répertoire existe
					If Not FileExists($subDirPath3) Then
						;ConsoleWrite("pas trouvé 3" & @CRLF)
					Else
						ConsoleWrite($subDirPath3 & @CRLF)
						; Chemin du répertoire à inspecter
						Local $sFolderPath = "\\nas-" & $NasSite & "\CloudStation\"
						ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
						; Obtenir la liste des sous-dossiers de PC
						Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)
						; Vérifier s'il y a eu une erreur lors de la récupération
						If @error Then
							;
							FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
						Else
							For $u = 1 To $aSubfolders[0]
								$Cduser = $sFolderPath & $aSubfolders[$u] & "\c\CD47Users"

								ConsoleWrite("cdusers " & $Cduser & @CRLF)
								Local $aNomSession = _FileListToArray($Cduser, "*", 2)

								;_FileListToArray(
								If @error = 1 Then
									ConsoleWrite("nom de session : " & $aNomSession & @CRLF)
									FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $subDirPath3)
									ExitLoop
								EndIf

								; Afficher la liste des sous-dossiers
								For $i = 1 To $aNomSession[0]
									ConsoleWrite("session " & $aNomSession[$i] & @CRLF)
									;date de modification du fichier de données outlook

									$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
									If @error Then
										;
										FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du LOG pour " & $aNomSession[$i])
									Else
										$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]

										;date de dernière connexion de l'utilisateur en cours
										$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
										If @error <> 0 Then
											FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
											ConsoleWrite(@error & @CRLF)
											ExitLoop
										EndIf
										;suppression de l'heure dans la date de connexion
										$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)


										$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
										If $nbr_jours_de_difference > 20 Then
											FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
										EndIf

									EndIf

								Next
							Next
						EndIf
					EndIf
				EndIf
				; Fin $subDirPath3
			Else
				; Vérifier si le sous-répertoire existe
				If Not FileExists($subDirPath2) Then
					;ConsoleWrite("pas trouvé 2" & @CRLF)
				Else
					;ConsoleWrite($subDirPath2 & @CRLF)
					; Chemin du répertoire à inspecter
					Local $sFolderPath = "\\nas-" & $NasSite & "\W10\"
					;ConsoleWrite("nas W10 " & $NasSite & @CRLF)
					;ConsoleWrite("folderPath W10 " & $sFolderPath & @CRLF)
					; Obtenir la liste des sous-dossiers de PC
					Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)
					; Vérifier s'il y a eu une erreur lors de la récupération
					If @error Then
						;
						FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
					Else
						For $u = 1 To $aSubfolders[0]
							$Cduser = $sFolderPath & $aSubfolders[$u] & "\c\CD47Users\"
							;ConsoleWrite("cdusers " & $Cduser & @CRLF)
							Local $aNomSession = _FileListToArray($Cduser, "*", 2)

							;_FileListToArray(
							If @error = 1 Then
								;ConsoleWrite("nom de session : " & $aNomSession & @CRLF)
								;ConsoleWrite("subfolder : " & $aSubfolders[$u] & @CRLF)
								FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $subDirPath2 & " " & $Cduser)
								ExitLoop
							EndIf

							; Afficher la liste des sous-dossiers
							For $i = 1 To $aNomSession[0]
								;ConsoleWrite($aNomSession[$i] & @CRLF)
								;date de modification du fichier de données outlook

								$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
								If @error Then
									;
									FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du LOG pour " & $aNomSession[$i])
								Else
									$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
									;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
									;date de dernière connexion de l'utilisateur en cours
									$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
									If @error <> 0 Then
										FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
										;ConsoleWrite(@error & @CRLF)
										ExitLoop
									EndIf
									;suppression de l'heure dans la date de connexion
									$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)


									$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
									If $nbr_jours_de_difference > 20 Then
										FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
									EndIf

								EndIf

							Next
						Next
					EndIf
				EndIf
			EndIf
			; Fin $subDirPath2
		Else

			; Vérifier si le sous-répertoire existe
			If Not FileExists($subDirPath) Then
				;ConsoleWrite("pas trouvé" & @CRLF)
			Else
				;ConsoleWrite($subDirPath & @CRLF)
				; Chemin du répertoire à inspecter
				Local $sFolderPath = "\\nas-" & $NasSite & "\"
				;ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
				; Obtenir la liste des sous-dossiers de PC
				Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)

				; Vérifier s'il y a eu une erreur lors de la récupération
				If @error Then
					;
					FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
				Else
					For $u = 1 To $aSubfolders[0]
						$Cduser = $sFolderPath & $aSubfolders[$u] & "\c\CD47Users"

						;ConsoleWrite("cdusers " & $Cduser & @CRLF)
						Local $aNomSession = _FileListToArray($Cduser, "*", 2)

						;_FileListToArray(
						If @error = 1 Then
							;ConsoleWrite("nom de session : " & $aNomSession & @CRLF)
							FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $subDirPath)
							ExitLoop
						EndIf

						; Afficher la liste des sous-dossiers
						For $i = 1 To $aNomSession[0]
							;ConsoleWrite($aNomSession[$i] & @CRLF)
							;date de modification du fichier de données outlook

							$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
							If @error Then
								;
								FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du LOG pour " & $aNomSession[$i])
							Else
								$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
								;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
								;date de dernière connexion de l'utilisateur en cours
								$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
								If @error = 1 Then
									FileWriteLine($outputFileHandle, $aNomSession[$i] & " Inconnu ")
									ExitLoop
								EndIf
								;suppression de l'heure dans la date de connexion
								$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)


								$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
								If $nbr_jours_de_difference > 20 Then
									FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
								EndIf

							EndIf

						Next
					Next
				EndIf
			EndIf

		EndIf
		;fin $subDirPath

		Sleep(10)
	WEnd

	SplashOff()
	; Fermer les handles des fichiers
	FileClose($fileHandle)
	FileClose($outputFileHandle)

	; Fermer la connexion à l'Active Directory
	_AD_Close()
	ShellExecute($outputFilePath)
EndFunc   ;==>CheckSAVPC

Func CheckSAVPCSJ($SiteD)
	_FileWriteLog($Cheminlog, "Fonction CheckSavPCSJ")
	SplashTextOn("Traitement en cours", "Merci de  patienter", 250, 120, -1, -1, 48, "Cambria", "", 800)
	cryptage($configini, 2, $Key)


	_FileWriteLog($Cheminlog, "PathIni " & $configini)
	Global $NomSite = IniRead($configini, $SiteD, "filtre", "")
	Global $NasSite = IniRead($configini, $SiteD, "Backup", "")

	_FileWriteLog($Cheminlog, "Filtre site " & $NomSite)
	_FileWriteLog($Cheminlog, "Dossier sur le Nas " & $NasSite)
	;Global $DossierSite = IniRead($configini, $SiteD, "dossier", "")
	cryptage($configini, 1, $Key)
	Sleep(100)

	; Chemin du fichier texte pour sauvegarder la liste
	Local $PathFilePC = "c:\temp\liste_pc2.txt"

	; Ouvrir le fichier de sortie en écriture
	Local $PathFilePCHandle = FileOpen($PathFilePC, 2)

	; Vérifier si le fichier de sortie a pu être ouvert
	If $PathFilePCHandle = -1 Then
		SplashOff()
		MsgBox(16, "Erreur", "Impossible d'ouvrir le fichier de sortie.")
		Exit
	EndIf

	; Ouvrir la connexion à l'Active Directory
	_AD_Open()

	; Récupérer les objets ordinateurs
	Local $aComputers = _AD_GetObjectsInOU("OU=W10,OU=CLIENTS CG47,DC=dptlg,DC=fr", "(&(objectClass=computer)(sAMAccountName=*" & $NomSite & "*))", 2, "name")

	_FileWriteLog($Cheminlog, "Nombre de Pc détecté : " & UBound($aComputers))
	; Vérifier s'il y a des ordinateurs correspondants
	If UBound($aComputers) = 0 Then
		SplashOff()
		MsgBox(64, "Information", "Aucun ordinateur trouvé.")

	Else
		; Écrire les noms d'ordinateurs dans la console
		For $i = 1 To UBound($aComputers) - 1
			;;ConsoleWrite($aComputers[$i] & @CRLF)
			FileWriteLine($PathFilePCHandle, $aComputers[$i])
		Next
	EndIf

	;ConsoleWrite("Fin de récupération des PC sur l'AD" & @CRLF)

	; Fermer le handle du fichier de sortie
	FileClose($PathFilePCHandle)

	; Lire le chemin du fichier des PC
	Local $inputFilePath = $PathFilePC

	; Vérifier si le fichier texte existe
	If Not FileExists($inputFilePath) Then
		SplashOff()
		MsgBox(16, "Erreur", "Le fichier liste n'existe pas.")
		Exit
	EndIf

	; Ouvrir le fichier texte en lecture
	Local $fileHandle = FileOpen($inputFilePath, 0)

	; Vérifier si le fichier a pu être ouvert
	If $fileHandle = -1 Then
		SplashOff()
		MsgBox(16, "Erreur", "Impossible d'ouvrir le fichier liste")
		Exit
	EndIf

	; Chemin du fichier de résultats
	Local $outputFilePath = "c:\temp\suiviSAV.txt"

	; Ouvrir le fichier de résultats en écriture
	Local $outputFileHandle = FileOpen($outputFilePath, 2)

	; Vérifier si le fichier de résultats a pu être ouvert
	If $outputFileHandle = -1 Then
		SplashOff()
		MsgBox(16, "Erreur", "Impossible d'ouvrir le fichier de résultats.")
		Exit
	EndIf

	; Lire le contenu du fichier texte ligne par ligne
	While 1
		Local $line = FileReadLine($fileHandle)
		; Sortir de la boucle si la fin du fichier est atteinte
		If @error = -1 Or $line = "" Then ExitLoop

		;ConsoleWrite("Debut du traitement" & @CRLF)

		; Construire le chemin complet du sous-répertoire
		Local $subDirPath = "\\nas-" & $NasSite & "\*" & $line & "*"
		Local $subDirPath2 = "\\nas-" & $NasSite & "\W10\*" & $line & "*"
		Local $subDirPath3 = "\\nas-" & $NasSite & "\CloudStation\*" & $line & "*"
		Local $subDirPath4 = "\\nas-" & $NasSite & "\Synology Drive\*" & $line & "*"
		Local $subDirPath5 = "\\nas-" & $NasSite & "\SynologyDrive\*" & $line & "*"
		; Vérifier si le sous-répertoire existe
		If Not FileExists($subDirPath) Then
			If Not FileExists($subDirPath2) Then
				If Not FileExists($subDirPath3) Then
					If Not FileExists($subDirPath4) Then
						If Not FileExists($subDirPath5) Then

							; Récupération de la dernière connexion du PC
							$aProperties = _AD_GetObjectProperties($line & "$", "lastlogon")

							; Vérification des erreurs
							If @error Then
								;MsgBox($MB_OK, "Erreur", "Erreur lors de la récupération des propriétés : " & @error)
								FileWriteLine($outputFileHandle, "Erreur lors de la récupération des propriétés : " & $line)
							Else
							EndIf

							; Vérification si la propriété lastlogon existe
							If IsArray($aProperties) And _ArraySearch($aProperties, "lastlogon", 0, 0, 0, 1) <> -1 Then
								; Affichage de la dernière connexion
								; ;ConsoleWrite($line & "Last Logon : " & $aProperties[1][1] & @CRLF)
								FileWriteLine($outputFileHandle, $line & ", n'est pas sauvegardé. Sa dernière connexion AD est le : " & $aProperties[1][1])
							Else
								; La propriété lastlogon n'est pas renseignée
								FileWriteLine($outputFileHandle, $line & ", n'est pas sauvegardé. Sa dernière connexion AD est le : ")
							EndIf

							;;ConsoleWrite($line & " La propriété lastlogon n'est pas renseignée." & @CRLF)
						Else

							; Vérifier si le sous-répertoire existe
							If Not FileExists($subDirPath5) Then
								_FileWriteLog($Cheminlog, $subDirPath5 & "N'existe pas ")
								;ConsoleWrite("pastrouvé" & @CRLF)
							Else
								;ConsoleWrite($subDirPath5 & @CRLF)
								; Chemin du répertoire à inspecter
								Local $sFolderPath = "\\nas-" & $NasSite & "\SynologyDrive\"
								;ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
								; Obtenir la liste des sous-dossiers de PC
								Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)

								; Vérifier s'il y a eu une erreur lors de la récupération
								If @error Then
									;
									FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
								Else
									For $u = 1 To $aSubfolders[0]
										$Cduser = $sFolderPath & $aSubfolders[$u] & "\c$\CD47Users"

										;ConsoleWrite("cdusers " & $Cduser & @CRLF)
										Local $aNomSession = _FileListToArray($Cduser, "*", 2)

										;_FileListToArray(

										If @error = 1 Then
											;ConsoleWrite("nom de session : " & $aNomSession & @CRLF)
											FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $subDirPath5)
											ExitLoop
										EndIf
										; Afficher la liste des sous-dossiers
										For $i = 1 To $aNomSession[0]
											;ConsoleWrite($aNomSession[$i] & @CRLF)
											;date de modification du fichier de données outlook

											$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
											If @error Then
												;
												FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du LOG pour " & $aNomSession[$i])
											Else
												$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
												;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
												;date de dernière connexion de l'utilisateur en cours
												$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
												If @error <> 0 Then
													FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
													;ConsoleWrite(@error & @CRLF)
													ExitLoop
												EndIf
												;suppression de l'heure dans la date de connexion
												$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)

												$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
												If $nbr_jours_de_difference > 20 Then
													FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
												EndIf
											EndIf
										Next
									Next
								EndIf
							EndIf
						EndIf
						; Fin $subDirPath5
					Else

						; Vérifier si le sous-répertoire existe
						If Not FileExists($subDirPath4) Then
							_FileWriteLog($Cheminlog, $subDirPath4 & "N'existe pas ")
							;ConsoleWrite("pastrouvé" & @CRLF)
						Else
							ConsoleWrite($subDirPath4 & @CRLF)
							; Chemin du répertoire à inspecter
							Local $sFolderPath = "\\nas-" & $NasSite & "\Synology Drive\"
							;ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
							; Obtenir la liste des sous-dossiers de PC
							Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)


							; Vérifier s'il y a eu une erreur lors de la récupération
							If @error Then
								;
								FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
							Else
								For $u = 1 To $aSubfolders[0]
									$Cduser = $sFolderPath & $aSubfolders[$u] & "\c$\CD47Users\"

									Local $aNomSession = _FileListToArray($Cduser, "*", 2)
									;;ConsoleWrite(@error &@CRLF)
									;_FileListToArray(

									If @error = 1 Then
										FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $subDirPath4)
										ExitLoop
									EndIf

									; Afficher la liste des sous-dossiers
									For $i = 1 To $aNomSession[0]
										;ConsoleWrite($aNomSession[$i] & @CRLF)
										;date de modification du fichier de données outlook

										$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
										If @error Then
											;
											FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du LOG pour " & $aNomSession[$i])
										Else
											$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
											;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
											;date de dernière connexion de l'utilisateur en cours
											$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
											If @error <> 0 Then
												FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
												;ConsoleWrite(@error & @CRLF)
												ExitLoop
											EndIf
											;suppression de l'heure dans la date de connexion
											$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)

											$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
											If $nbr_jours_de_difference > 20 Then
												FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
											EndIf
										EndIf
									Next
								Next
							EndIf
						EndIf
					EndIf
					; Fin $subDirPath4
				Else

					; Vérifier si le sous-répertoire existe
					If Not FileExists($subDirPath3) Then
						_FileWriteLog($Cheminlog, $subDirPath3 & "N'existe pas ")
						;ConsoleWrite("pastrouvé" & @CRLF)
					Else
						;ConsoleWrite($subDirPath3 & @CRLF)
						; Chemin du répertoire à inspecter
						Local $sFolderPath = "\\nas-" & $NasSite & "\CloudStation\"
						;ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
						; Obtenir la liste des sous-dossiers de PC
						Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)

						; Vérifier s'il y a eu une erreur lors de la récupération
						If @error Then
							;
							FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
						Else
							For $u = 1 To $aSubfolders[0]
								$Cduser = $sFolderPath & $aSubfolders[$u] & "\c$\CD47Users"

								;ConsoleWrite("cdusers " & $Cduser & @CRLF)
								Local $aNomSession = _FileListToArray($Cduser, "*", 2)

								;_FileListToArray(
								If @error = 1 Then
									;ConsoleWrite("nom de session : " & $aNomSession & @CRLF)
									FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC  " & $subDirPath3)
									ExitLoop
								EndIf

								; Afficher la liste des sous-dossiers
								If IsArray($aNomSession) Then
									For $i = 1 To $aNomSession[0]
										;ConsoleWrite($aNomSession[$i] & @CRLF)
										;date de modification du fichier de données outlook

										$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
										If @error Then
											;
											FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du LOG pour " & $aNomSession[$i])
										Else
											$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
											;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
											;date de dernière connexion de l'utilisateur en cours
											$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
											If @error <> 0 Then
												FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
												;ConsoleWrite(@error & @CRLF)
												ExitLoop
											EndIf
											;suppression de l'heure dans la date de connexion
											$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)

											$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
											If $nbr_jours_de_difference > 20 Then
												FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
											EndIf
										EndIf
									Next
								EndIf
							Next
						EndIf
					EndIf
				EndIf
				; Fin $subDirPath3
			Else
				; Vérifier si le sous-répertoire existe
				If Not FileExists($subDirPath2) Then
					_FileWriteLog($Cheminlog, $subDirPath2 & "N'existe pas ")
					;ConsoleWrite("pastrouvé" & @CRLF)
				Else
					;ConsoleWrite($subDirPath2 & @CRLF)
					; Chemin du répertoire à inspecter
					Local $sFolderPath = "\\nas-" & $NasSite & "\W10\"
					;ConsoleWrite("nas W10 " & $NasSite & @CRLF)
					;ConsoleWrite("folderPath W10 " & $sFolderPath & @CRLF)
					; Obtenir la liste des sous-dossiers de PC
					Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)

					; Vérifier s'il y a eu une erreur lors de la récupération
					If @error Then
						;
						FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
					Else
						For $u = 1 To $aSubfolders[0]
							$Cduser = $sFolderPath & $aSubfolders[$u] & "\c$\CD47Users\"
							;ConsoleWrite("cdusers " & $Cduser & @CRLF)
							Local $aNomSession = _FileListToArray($Cduser, "*", 2)

							;_FileListToArray(
							If @error = 1 Then
								;ConsoleWrite("nom de session : " & $aNomSession & @CRLF)
								;ConsoleWrite("subfolder : " & $aSubfolders[$u] & @CRLF)
								FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $subDirPath2)
								ExitLoop
							EndIf
							If IsArray($aNomSession) Then
								; Afficher la liste des sous-dossiers
								For $i = 1 To $aNomSession[0]
									;ConsoleWrite($aNomSession[$i] & @CRLF)
									;date de modification du fichier de données outlook

									$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
									If @error Then
										;
										FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du LOG pour " & $aNomSession[$i])
									Else
										$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
										;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
										;date de dernière connexion de l'utilisateur en cours
										$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
										If @error <> 0 Then
											FileWriteLine($outputFileHandle, $aNomSession[$i] & " inconnu sur l'AD pour le PC " & $aSubfolders[$u])
											;ConsoleWrite(@error & @CRLF)
											ExitLoop
										EndIf
										;suppression de l'heure dans la date de connexion
										If IsArray($adateconnexionU) And UBound($adateconnexionU, 0) >= 1 And UBound($adateconnexionU, 1) >= 2 Then
											$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)
										Else
											;ConsoleWrite("Invalid array or out of bounds: $adateconnexionU" & @CRLF)
										EndIf

										$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
										If $nbr_jours_de_difference > 20 Then
											FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
										EndIf
									EndIf
								Next
							EndIf
						Next
					EndIf
				EndIf
			EndIf
			; Fin $subDirPath2
		Else

			; Vérifier si le sous-répertoire existe
			If Not FileExists($subDirPath) Then
				_FileWriteLog($Cheminlog, $subDirPath & "N'existe pas ")
				;ConsoleWrite("pastrouvé" & @CRLF)
			Else
				;ConsoleWrite($subDirPath & @CRLF)
				; Chemin du répertoire à inspecter
				Local $sFolderPath = "\\nas-" & $NasSite & "\"
				;ConsoleWrite("folderPath" & $sFolderPath & @CRLF)
				; Obtenir la liste des sous-dossiers de PC
				Local $aSubfolders = _FileListToArray($sFolderPath, "*" & $line & "*", 2)

				; Vérifier s'il y a eu une erreur lors de la récupération
				If @error Then
					;
					FileWriteLine($outputFileHandle, $sFolderPath & " impossible de récupérer la liste des sous dossiers ")
				Else
					For $u = 1 To $aSubfolders[0]
						$Cduser = $sFolderPath & $aSubfolders[$u] & "\c$\CD47Users"

						;ConsoleWrite("cdusers " & $Cduser & @CRLF)
						Local $aNomSession = _FileListToArray($Cduser, "*", 2)

						;_FileListToArray(
						If @error = 1 Then
							;ConsoleWrite("nom de session : " & $aNomSession & @CRLF)
							FileWriteLine($outputFileHandle, $aSubfolders[$u] & " Erreur de traitement du PC " & $subDirPath)
							ExitLoop
						EndIf
						If IsArray($aNomSession) Then
							; Afficher la liste des sous-dossiers
							For $i = 1 To $aNomSession[0]
								;ConsoleWrite($aNomSession[$i] & @CRLF)
								;date de modification du fichier de données outlook

								$datePst = FileGetTime($Cduser & "\" & $aNomSession[$i] & "\Documents\docappli\controleSAV.log", 0)
								If @error Then
									;
									FileWriteLine($outputFileHandle, $aSubfolders[$u] & " impossible de récupérer la date du LOG pour " & $aNomSession[$i])
								Else
									$datepstf = $datePst[0] & "/" & $datePst[1] & "/" & $datePst[2]
									;ConsoleWrite("Date PST " & $datepstf & " pour : " & $aNomSession[$i] & "sur " & $aSubfolders[$u] & @CRLF)
									;date de dernière connexion de l'utilisateur en cours
									$adateconnexionU = _AD_GetObjectProperties($aNomSession[$i], "lastlogon")
									If @error = 1 Then
										FileWriteLine($outputFileHandle, $aNomSession[$i] & " Inconnu ")
										ExitLoop
									EndIf
									;suppression de l'heure dans la date de connexion
									If IsArray($adateconnexionU) And UBound($adateconnexionU, 0) >= 1 And UBound($adateconnexionU, 1) >= 2 Then
										$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)
									Else
										;ConsoleWrite("Invalid array or out of bounds: $adateconnexionU" & $aNomSession & @CRLF)
									EndIf

									;									$adateconnexionU = StringLeft($adateconnexionU[1][1], 10)


									$nbr_jours_de_difference = _DateDiff("D", $datepstf, $adateconnexionU)
									If $nbr_jours_de_difference > 20 Then
										FileWriteLine($outputFileHandle, "le LOG de " & $aNomSession[$i] & " Sur " & $aSubfolders[$u] & " n a pas été sauvegardé depuis " & $nbr_jours_de_difference & " jours soit le " & $datepstf & @CRLF)
									EndIf

								EndIf

							Next
						EndIf
					Next
				EndIf
			EndIf
		EndIf
		;fin $subDirPath

		Sleep(10)
	WEnd

	SplashOff()
	; Fermer les handles des fichiers
	FileClose($fileHandle)
	FileClose($outputFileHandle)

	; Fermer la connexion à l'Active Directory
	_AD_Close()
	ShellExecute($outputFilePath)
EndFunc   ;==>CheckSAVPCSJ

