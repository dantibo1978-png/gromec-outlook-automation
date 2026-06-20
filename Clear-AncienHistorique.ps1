<#
.SYNOPSIS
    Nettoie le noeud "historique" de Firebase (gromec-outlook-vba) en supprimant
    les entrées plus vieilles que X jours.
.DESCRIPTION
    Ce script est conçu pour rouler de façon autonome via GitHub Actions (cron).
    Il n'a besoin d'aucun accès à Outlook, au lecteur U:\, ni à aucun poste
    de travail spécifique — seulement une connexion Internet vers Firebase.
    Logique :
      1. Lit TOUT le noeud "historique" sous gromec_vba (GET sur l'URL du noeud)
      2. Pour chaque entrée, compare le champ "date" (format ISO "YYYY-MM-DDTHH:MM:SS")
         à la date limite de rétention
      3. Supprime (DELETE) les entrées plus vieilles que la limite, une par une
.NOTES
    Comme "date" est en format ISO, une simple comparaison de chaînes de
    caractères suffit (pas besoin de parser en objet DateTime pour comparer),
    mais on parse quand même pour plus de robustesse / lisibilité des logs.
    Le noeud historique vit sous gromec_vba (pas a la racine de la base).
#>
param(
    [int]$JoursRetention = 14,
    [string]$FirebaseUrl = $env:FIREBASE_DB_URL
)
$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($FirebaseUrl)) {
    throw "FIREBASE_DB_URL n'est pas défini. Configure-le comme secret GitHub Actions (sans slash final), ex: https://gromec-outlook-vba-default-rtdb.firebaseio.com"
}
$NoeudHistorique = "$FirebaseUrl/gromec_vba/historique.json"
$DateLimite = (Get-Date).AddDays(-$JoursRetention)
Write-Host "=== Nettoyage historique Firebase ==="
Write-Host "Date limite de rétention : $($DateLimite.ToString('yyyy-MM-ddTHH:mm:ss')) (garder $JoursRetention jours)"
Write-Host "Noeud interrogé : $NoeudHistorique"
Write-Host ""
# 1. Lire tout le noeud historique
try {
    $Historique = Invoke-RestMethod -Uri $NoeudHistorique -Method Get
} catch {
    throw "Échec de la lecture du noeud historique : $($_.Exception.Message)"
}
if ($null -eq $Historique) {
    Write-Host "Le noeud historique est vide ou inexistant. Rien à faire."
    exit 0
}
# $Historique est un objet PSCustomObject où chaque propriété est une push-key Firebase
$Cles = $Historique.PSObject.Properties.Name
Write-Host "Nombre total d'entrées trouvées : $($Cles.Count)"
$ASupprimer = @()
$Conservees = 0
$DatesIllisibles = 0
foreach ($Cle in $Cles) {
    $Entree = $Historique.$Cle
    $DateTexte = $Entree.date
    if ([string]::IsNullOrWhiteSpace($DateTexte)) {
        Write-Warning "Entrée '$Cle' sans champ 'date' - conservée par prudence."
        $Conservees++
        continue
    }
    try {
        $DateEntree = [DateTime]::Parse($DateTexte, [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        Write-Warning "Entrée '$Cle' a une date illisible ('$DateTexte') - conservée par prudence."
        $DatesIllisibles++
        continue
    }
    if ($DateEntree -lt $DateLimite) {
        $ASupprimer += $Cle
    } else {
        $Conservees++
    }
}
Write-Host ""
Write-Host "À supprimer : $($ASupprimer.Count)"
Write-Host "Conservées (récentes) : $Conservees"
if ($DatesIllisibles -gt 0) {
    Write-Host "Conservées (date illisible) : $DatesIllisibles"
}
Write-Host ""
if ($ASupprimer.Count -eq 0) {
    Write-Host "Aucune entrée à supprimer. Terminé."
    exit 0
}
# 2. Supprimer les entrées trop vieilles, une par une
$SupprimeesOk = 0
$SupprimeesErreur = 0
foreach ($Cle in $ASupprimer) {
    $UrlSuppression = "$FirebaseUrl/gromec_vba/historique/$Cle.json"
    try {
        Invoke-RestMethod -Uri $UrlSuppression -Method Delete | Out-Null
        $SupprimeesOk++
    } catch {
        Write-Warning "Échec de la suppression de '$Cle' : $($_.Exception.Message)"
        $SupprimeesErreur++
    }
}
Write-Host "=== Résumé ==="
Write-Host "Supprimées avec succès : $SupprimeesOk"
if ($SupprimeesErreur -gt 0) {
    Write-Host "Échecs de suppression : $SupprimeesErreur"
}
Write-Host "Nettoyage terminé."
