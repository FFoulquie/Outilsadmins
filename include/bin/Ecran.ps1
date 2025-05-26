param (
    [string]$ComputerName
)

# Vérifiez si un nom d'ordinateur a été fourni
if (-not $ComputerName) {
    Write-Host "Veuillez fournir le nom de l'ordinateur avec le paramètre -ComputerName"
    exit
}

# Déplacement dans le répertoire "Mes documents" de l'utilisateur actuel
Set-Location $([Environment]::GetFolderPath("MyDocuments"))

# Import du module Active Directory
ipmo activedirectory

# Récupération de l'ordinateur spécifié du domaine, en prenant également la dernière ouverture de session
$ordinateur = Get-ADComputer -Identity $ComputerName -Properties lastlogondate

if ($null -eq $ordinateur) {
    Write-Host "L'ordinateur spécifié n'a pas été trouvé dans le domaine."
    exit
}

# Création d'un tableau
$tableau = New-Object System.Collections.ArrayList
$progress = 0

$progress++
$objet = New-Object Psobject
$objet | Add-Member -MemberType NoteProperty -Name "Ordinateur" -Value $ordinateur.Name
$ping = $null
# Nous pingons le poste
$ping = Test-Connection -ComputerName $ordinateur.Name -Count 1 -ErrorAction SilentlyContinue

# Si le TTL du poste est entre 65 et 128, c'est un poste Windows. Nous continuons donc le script
if ($ping -and $ping.ResponsetimeToLive -gt 64 -and $ping.ResponsetimeToLive -le 128) {
    $objet | Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Oui"
    $moniteurs = $null

    # On lance la requête WMI
    $moniteurs = Get-WmiObject -Namespace root\wmi -Class wmiMonitorID -ComputerName $($ordinateur.Name) -ErrorAction SilentlyContinue
    # Si le résultat de la commande précédente est une erreur, alors...
    if (!$?) {
        Write-Host "Erreur!"
        $objet | Add-Member -MemberType NoteProperty -Name "WMI" -Value "ERREUR!!"
    } else {
        Write-Host "Aucune erreur"
        $objet | Add-Member -MemberType NoteProperty -Name "WMI" -Value ""
    }

    $i = 0
    $moniteursInfo = @()
    foreach ($moniteur in $moniteurs) {
        $i++
        # Le modèle est la deuxième valeur de la propriété instancename.
        $modele = $moniteur.instancename.split("\")[1]
        $sn = $null
        # Récupération du SN puis ajout dans l'objet
        $moniteur.serialnumberid | % { $sn += [char]$_ }
        $objet | Add-Member -MemberType NoteProperty -Name "Modele ecran $i" -Value $modele
        $objet | Add-Member -MemberType NoteProperty -Name "SN ecran $i" -Value $sn
        $moniteursInfo += "Modele ecran $i : $modele, SN ecran $i : $sn"
    }

    $i++
    for ($i; $i -le 3; $i++) {
        $objet | Add-Member -MemberType NoteProperty -Name "Modèle écran $i" -Value ""
        $objet | Add-Member -MemberType NoteProperty -Name "SN écran $i" -Value ""
    }

    # Écrire les informations des moniteurs dans un fichier texte
    $filePath = "c:\temp\$($ordinateur.Name).txt"
    $moniteursInfo | Out-File -FilePath $filePath -Encoding UTF8
} elseif ($ping.ResponsetimeToLive -ne $null -and $ping.ResponsetimeToLive -le 64) {
    $objet | Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Non Windows"
} else {
    $objet | Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Non"
}



