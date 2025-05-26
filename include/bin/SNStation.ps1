# Changement du répertoire de travail
Set-Location $([Environment]::GetFolderPath("MyDocuments"))

# Importation du module Active Directory
ipmo activedirectory

# Récupération de la liste des ordinateurs du domaine avec lastLogonTimestamp
$ordinateurs = Get-ADComputer -Filter 'Name -like "P*"' -SearchBase "OU=Ordinateurs,OU=W10,OU=CLIENTS CG47,DC=dptlg,DC=fr" -Properties lastLogonTimestamp | Select-Object -Property name, lastLogonTimestamp

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
    Write-Progress -Activity "Recherche station accueil sur $($ordinateur.name)" -PercentComplete ($progress / $ordinateurs.Count * 100) -Status "$progress / $($ordinateurs.count)"
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

           

            # Requête WMI pour récupérer les informations des stations d'accueil Lenovo
            $dockDevices = Get-WmiObject -Namespace root\Lenovo\Dock_Manager -Class DockDevice -ComputerName $($ordinateur.name) -ErrorAction SilentlyContinue
            if ($dockDevices) {
                $dockInfo = @()
                foreach ($dockDevice in $dockDevices) {
                    $dockSerialNumber = $dockDevice.SerialNumber
                  #  $dockMacAddress = $dockDevice.MACAddress
                    $dockInfo += "SerialNumber: $dockSerialNumber" #, MACAddress: $dockMacAddress"
                    $objet | Add-Member -MemberType NoteProperty -Name "Dock Serial Number" -Value $dockSerialNumber
                    #$objet | Add-Member -MemberType NoteProperty -Name "Dock MAC Address" -Value $dockMacAddress
                }

                # Écrire les informations des docks dans un fichier texte
                $dockFilePath = "\\vulcain\fitrans$\dsian\log\dock\$($ordinateur.name).txt"
                $dockInfo | Out-File -FilePath $dockFilePath -Encoding UTF8
            } else {
                #Write-Host "$($ordinateur.name) Erreur lors de la récupération des informations de la station d'accueil!"
                $objet | Add-Member -MemberType NoteProperty -Name "Dock Info" -Value "ERREUR!!"
            }

            
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

