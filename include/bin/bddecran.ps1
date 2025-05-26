# Changement du répertoire de travail
Set-Location $([Environment]::GetFolderPath("MyDocuments"))

# Importation du module Active Directory
ipmo activedirectory

# Récupération de la liste des ordinateurs du domaine avec lastLogonTimestamp
$ordinateurs = Get-ADComputer -Filter * -SearchBase "OU=Ordinateurs,OU=W10,OU=CLIENTS CG47,DC=dptlg,DC=fr" -Properties lastLogonTimestamp | Select-Object -Property name, lastLogonTimestamp

# Conversion de lastLogonTimestamp en date lisible
$ordinateurs | ForEach-Object {
    $_ | Add-Member -MemberType NoteProperty -Name 'LastLogonDate' -Value ([DateTime]::FromFileTime($_.lastLogonTimestamp))
}

# Filtrage des ordinateurs qui se sont connectés dans les dernières 24 heures
$ordinateurs = $ordinateurs | Where-Object { $_.LastLogonDate -gt (Get-Date).AddDays(-7) }

# Création d'un tableau pour stocker les résultats
$tableau = New-Object System.Collections.ArrayList
$progress = 0

foreach ($ordinateur in $ordinateurs) {
    $progress++
    Write-Progress -Activity "Recherche d ecrans sur $($ordinateur.name)" -PercentComplete ($progress / $ordinateurs.Count * 100) -Status "$progress / $($ordinateurs.count)"
    $objet = New-Object Psobject
    $objet | Add-Member -MemberType NoteProperty -Name "Ordinateur" -Value $ordinateur.name

    # Test de connexion (ping)
    $ping = Test-Connection -ComputerName $ordinateur.name -Count 1 -ErrorAction SilentlyContinue

    # Vérification du système d'exploitation Windows et de l'adresse IP
    if ($ping -and $ping.ResponseTimeToLive -gt 64 -and $ping.ResponseTimeToLive -le 128) {
        $ipAddress = ($ping | Select-Object -First 1).Address

        if ($ipAddress -notlike "172.16.*") {
            $objet | Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Oui"
            $moniteurs = $null

            # Requête WMI pour récupérer les informations des moniteurs
            $moniteurs = Get-WmiObject -Namespace root\wmi -Class wmiMonitorID -ComputerName $($ordinateur.name) -ErrorAction SilentlyContinue
            if (!$?) {
                Write-Host "$($ordinateur.name) Erreur!"
                $objet | Add-Member -MemberType NoteProperty -Name "WMI" -Value "ERREUR!!"
            } else {
                $objet | Add-Member -MemberType NoteProperty -Name "WMI" -Value ""
            }

            $i = 0
            $moniteursInfo = @()
            foreach ($moniteur in $moniteurs) {
                $i++
                $modele = $moniteur.instancename.split("\")[1]
                $sn = $null
                $moniteur.serialnumberid | % { $sn += [char]$_ }
                $objet | Add-Member -MemberType NoteProperty -Name "Modèle écran $i" -Value $modele
                $objet | Add-Member -MemberType NoteProperty -Name "SN écran $i" -Value $sn
                $moniteursInfo += "Modèle écran $i : $modele, SN écran $i : $sn"
            }

            $i++
            for ($i; $i -le 3; $i++) {
                $objet | Add-Member -MemberType NoteProperty -Name "Modèle écran $i" -Value ""
                $objet | Add-Member -MemberType NoteProperty -Name "SN écran $i" -Value ""
            }

            # Écriture des informations des moniteurs dans un fichier texte
            $filePath = "\\vulcain\fitrans$\dsian\log\ecran\$($ordinateur.name).txt"
            $moniteursInfo | Out-File -FilePath $filePath -Encoding UTF8
        } else {
            $objet | Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Oui (IP filtrée)"
            Write-Host "$($ordinateur.name) IP commence par 172.16, WMI ignoré."
        }
    } elseif ($ping.ResponseTimeToLive -ne $null -and $ping.ResponseTimeToLive -le 64) {
        $objet | Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Non Windows"
    } else {
        $objet | Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Non"
    }
    $null = $tableau.add($objet)
}

# Exportation des résultats
#$tableau | Export-Clixml -Path "c:\temp\ecrans.xml"
#$tableau | Out-File -FilePath "c:\temp\ecrans.txt" -Encoding UTF8
