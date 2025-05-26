#param($ordinateurs)
#déplacement dans le repertoires "mes documents" de l'utilisateur actuel
Set-Location $([Environment]::GetFolderPath("Mydocuments"))
#import du module activedirectory
ipmo activedirectory
#nous récupérons la liste des ordinateurs du domaine, en prenant également la dernière ouverture de session
#$ordinateurs = "p000-t4187$"
$ordinateurs = Get-ADComputer -Filter * -SearchBase "OU=Test, OU=Ordinateurs, OU=W10, OU=CLIENTS CG47,DC=dptlg,DC=fr" -Properties lastlogondate | Select-Object -Property name,lastlogondate
#nous filtrons à 4 mois en arrière
$ordinateurs = $ordinateurs | Where-Object {$_.lastlogondate -gt $((Get-Date).addmonths(-4))}
#creation d'un tableau
$tableau = New-Object System.Collections.ArrayList
$progress = 0
foreach ($ordinateur in $ordinateurs)
 {
 
  $progress++
  #Write-Progress -Activity "Recherche d'écrans sur $($ordinateur.name) " -PercentComplete $($progress/$($ordinateurs.Count)*100) -Status "$progress / $($ordinateurs.count)"
  $objet = New-Object Psobject
  $objet |Add-Member -MemberType NoteProperty -Name "Ordinateur" -Value $ordinateur.name
  $ping = $null
  #nous pingons le poste.
  $ping = Test-Connection -ComputerName $ordinateur.name -Count 1
  #Si le TTL du poste est entre 65 et 128, c'est un poste windows. Nous continuons donc le script
  if ($ping -and $ping.responsetimetolive -gt 64 -and $ping.responsetimetolive -le 128)
  {
   $objet |Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Oui"
   $moniteurs = $null
 
    $moniteurs = $null
    #on lance la requete WMI
    $moniteurs = Get-WmiObject -Namespace root\wmi -Class wmiMonitorID -ComputerName $($ordinateur.name)
   #Si le resultat de la commande precedente est une erreur, alors...    
   if (!$?)
 
    {
    Write-Host "Erreur!"
    $objet | Add-Member -MemberType NoteProperty -Name "WMI" -Value "ERREUR!!"
    }
   #Sinon...
   else
    {
    Write-Host "Aucune erreur"
    $objet | Add-Member -MemberType NoteProperty -Name "WMI" -Value ""
    }
    $i=0
   foreach ($moniteur in $moniteurs)
    {
    $i++
    #Le modele est la deuxième valeur de la propriété instancename.
    $modele = $moniteur.instancename.split("\")[1]
    $sn = $null
    #Récupération du SN puis ajout dans l'objet
    $moniteur.serialnumberid | % {$sn += [char]$_}
    $objet | Add-Member -MemberType NoteProperty -Name "Modèle ecran $i" -Value $modele
    $objet | Add-Member -MemberType NoteProperty -Name "SN ecran $i" -Value $sn
 
    }
    $i++
    for($i; $i -le 3;$i++)
    {
    $objet | Add-Member -MemberType NoteProperty -Name "Modèle ecran $i" -Value ""
    $objet | Add-Member -MemberType NoteProperty -Name "SN ecran $i" -Value ""
 
    }
 
  }
  #si le ttl est inférieur ou égal à 64, c'est un non windows
  elseif ($ping.responsetimetolive -ne $null -and $ping.responsetimetolive -le 64)
   {
   $objet |Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Non Windows"
   }
  else 
   {
   $objet | Add-Member -MemberType NoteProperty -Name "Répond au ping" -Value "Non"
   }
 $null = $tableau.add($objet)
 
 }
$tableau | Export-Clixml -Path "c:\temp\ecrans.xml" 
#$tableau | Export-Csv -Path "c:\temp\ecrans.csv" -NoTypeInformation -Encoding UTF8 -Force -Delimiter ";"
$tableau | Out-File -FilePath "c:\temp\ecrans.txt" -Encoding UTF8
