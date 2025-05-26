func _Connect()
	Global $Domain = "dptlg"
	Global $sUsername = InputBox('User', "Nom d'utilisateur")
	Select
		Case @error = 0 ;OK -
			Global $sPassword = InputBox('Password', 'Mot de passe', '', '*')
			Select
				Case @error = 0
					MsgBox(0, 'Test nom loggin et mot de passe', _CheckUserPass($sUsername, $sPassword, $Domain))
				Case @error = 1
					MsgBox(0, "_", "Annulé par l'utilisateur")
			EndSelect

		Case @error = 1
			MsgBox(0, "_", "Annulé par l'utilisateur")
			exit
	EndSelect

EndFunc
;_CheckUserPass($sUsername, $sPassword, $Domain))

Func _CheckUserPass($sUsername, $sPassword, $Domain)
	Local $iCheck = True
	;Return $iCheck
	;NE fonctionne pas avec les comptes Techs Pourtant plus rapide et sans fenetre Dos
	;RunAs($sUsername, $Domain, $sPassword, 0, @ComSpec & " /c  echo vérification du mot de passe..", @TempDir, @SW_HIDE)

	RunAs($sUsername, $Domain, $sPassword, 0, "ping localhost n 1")

	If @error Then
		$iCheck = "utilisateur ou mot de passe incorrect"
		Return $iCheck

		Exit

	Else
		CheckGroup()
	EndIf
EndFunc   ;==>_CheckUserPass

Func CheckGroup()
	$utilisateur = $sUsername
	$nom_groupe = "ADMIN_LOCAUX_PC"
	$nom_groupe2 = "TECH_SI_AD"
	$nom_groupe3 = "Admins du domaine"


	; Obtenir les informations sur le groupe du domaine
	$group_path = "WinNT://dptlg/" & $nom_groupe & ",group"
	$domain_group = ObjGet($group_path)

	$group_path2 = "WinNT://dptlg/" & $nom_groupe2 & ",group"
	$domain_group2 = ObjGet($group_path2)

	$group_path3 = "WinNT://dptlg/" & $nom_groupe3 & ",group"
	$domain_group3 = ObjGet($group_path3)

	; Vérifier si l'utilisateur est membre du groupe du domaine
	For $membre In $domain_group.Members
		If $membre.Name = $utilisateur Then
			_fenetre()
			Exit

		EndIf
	Next

	For $membre2 In $domain_group2.Members
		If $membre2.Name = $utilisateur Then
			_fenetre()
			Exit

		EndIf
	Next

	For $membre3 In $domain_group3.Members
		If $membre3.Name = $utilisateur Then
			_fenetre()
			Exit

		EndIf
	Next
	$Subject = "Utilisation interdite outil Admin " & @ComputerName ; subject from the email - can be anything you want it to be
	$Body = "Outil admin a été lancé sans succès sur le poste " & @ComputerName & " par " & @UserName ; the messagebody from the mail - can be left blank but then you get a blank mail
	envoiMail($Subject, $Body)

	MsgBox(0, "Attention", "Vous n'êtes pas autorisé à lancer cette application. Un Email a été envoyé au gestionnaire de l'application.")
	Exit
EndFunc   ;==>CheckGroup

#Region ENVOIS DU MAIL
Func envoiMail($Subject, $Body)
	;_______________________Variables___________________________
	$SmtpServer = "smtp.dptlg.fr" ; address for the smtp-server to use - REQUIRED
	$FromName = "Outil Admin  " ; name from who the email was sent
	$FromAddress = "frederic.foulquie@lotetgaronne.fr" ; address from where the mail should come
	$ToAddress = "frederic.foulquie@lotetgaronne.fr " ; destination address of the email - REQUIRED
	$AttachFiles = "" ; the file(s) you want to attach seperated with a ; (Semicolon)
	$CcAddress = "" ; address for cc - leave blank if not needed
	$BccAddress = "" ; address for bcc - leave blank if not needed
	$Importance = "Normal" ; Send message priority: "High", "Normal", "Low"
	;$Username = "notifinfra" ; username for the account used from where the mail gets sent - REQUIRED
	$Username = "" ; username for the account used from where the mail gets sent - REQUIRED
	;$Password = "PiRasd41" ; password for the account used from where the mail gets sent - REQUIRED
	$Password = "" ; password for the account used from where the mail gets sent - REQUIRED
	$IPPort = 25 ; port used for sending the mail
	$ssl = 0 ; enables/disables secure socket layer sending - put to 1 if using httpS

	;_______________________Script___________________________

	Global $oMyRet[2]
	Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
	$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $Username, $Password, $IPPort, $ssl)
	If @error Then
		MsgBox(0, "Error sending message", "Error code:" & @error & "  Description:" & $rc)
		ClipPut($rc)
	EndIf
	;desactive()
EndFunc   ;==>envoiMail
;_______________________UDF___________________________

Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance = "Normal", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)
	Local $objEmail = ObjCreate("CDO.Message")
	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress
	Local $i_Error = 0
	Local $i_Error_desciption = ""
	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
	$objEmail.Subject = $s_Subject
	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		$objEmail.HTMLBody = $as_Body
	Else
		$objEmail.Textbody = $as_Body & @CRLF
	EndIf

	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
	If Number($IPPort) = 0 Then $IPPort = 25
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort ;Authenticated SMTP
	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf
	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	EndIf
	$objEmail.Configuration.Fields.Update ;Update settings
	Switch $s_Importance ; Set Email Importance
		Case "High"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "High"
		Case "Normal"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Normal"
		Case "Low"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Low"
	EndSwitch
	$objEmail.Fields.Update
	$objEmail.Send ; Sent the Message
	If @error Then
		SetError(2)
		Return $oMyRet[1]
	EndIf
	$objEmail = ""
EndFunc   ;==>_INetSmtpMailCom
