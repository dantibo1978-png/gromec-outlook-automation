# =============================================================================
# RelanceFournisseur.ps1
# Tourne en arriere-plan sur le poste Windows (meme machine qu'Outlook/SyncDTW).
# Poll Firebase toutes les 60s pour la file gromec_vba/relance_auto, remplie par
# l'onglet outils/relance-fournisseur.html (bouton "Mettre en file d'envoi auto").
#
# SECURITE : par defaut ce script tourne en mode ESSAI (-Envoyer non fourni).
# En mode essai, AUCUN courriel n'est envoye : le script journalise ce qu'il
# aurait fait (destinataire, sujet, PJ) et laisse le statut Firebase a
# "attente" pour qu'on puisse verifier avant d'activer l'envoi reel.
#
# Lancement en essai (ne fait qu'observer/journaliser, aucun envoi) :
#   powershell -ExecutionPolicy Bypass -File "U:\GromecOutlook\RelanceFournisseur.ps1"
#
# Lancement reel (envoie vraiment les courriels via Outlook) -- A N'UTILISER
# QU'APRES VALIDATION EXPLICITE :
#   powershell -ExecutionPolicy Bypass -File "U:\GromecOutlook\RelanceFournisseur.ps1" -Envoyer
# =============================================================================

param(
    [switch]$Envoyer,          # sans ce flag : mode essai, aucun envoi reel
    [switch]$UneFois,          # traite la file une seule fois puis quitte (utile pour tester)
    [int]$IntervalleSecondes = 60
)

# ── Auto-update depuis GitHub ─────────────────────────────────────────────────
$GitHubRawUrl = "https://raw.githubusercontent.com/dantibo1978-png/gromec-outlook-automation/main/RelanceFournisseur.ps1"

function Update-ScriptSiNecessaire {
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.Encoding = [System.Text.Encoding]::UTF8
        $remoteContent = $webClient.DownloadString($GitHubRawUrl)
        $webClient.Dispose()
    } catch {
        return
    }

    if ([string]::IsNullOrWhiteSpace($remoteContent)) { return }

    $scriptPath = $PSCommandPath
    if ([string]::IsNullOrEmpty($scriptPath)) { return }

    try {
        $localContent = Get-Content -Path $scriptPath -Raw -Encoding UTF8
    } catch {
        return
    }

    $normLocal  = ($localContent  -replace "`r`n", "`n").Trim()
    $normRemote = ($remoteContent -replace "`r`n", "`n").Trim()

    if ($normLocal -eq $normRemote) { return }

    try {
        $utf8AvecBom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($scriptPath, $remoteContent, $utf8AvecBom)
    } catch {
        return
    }

    $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $scriptPath)
    if ($Envoyer) { $argList += "-Envoyer" }
    if ($UneFois) { $argList += "-UneFois" }
    $argList += @("-IntervalleSecondes", $IntervalleSecondes)

    Start-Process -FilePath "powershell.exe" -ArgumentList $argList -WindowStyle Hidden
    exit 0
}

Update-ScriptSiNecessaire

# ── Configuration ─────────────────────────────────────────────────────────────
$FirebaseUrl   = "https://gromec-outlook-vba-default-rtdb.firebaseio.com"
$DossierTemp   = "U:\GromecOutlook\RelanceTemp"

if (-not (Test-Path $DossierTemp)) {
    New-Item -ItemType Directory -Path $DossierTemp -Force | Out-Null
}

function Write-Log {
    param([string]$Message, [string]$Niveau = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ligne = "[$ts] [$Niveau] $Message"
    Write-Host $ligne
    try {
        $body = @{ ts = $ts; msg = $Message; niveau = $Niveau; source = "RelanceFournisseur" } | ConvertTo-Json -Compress
        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/logs.json" -Method Post -Body $body -ContentType "application/json" -TimeoutSec 10 | Out-Null
    } catch {}
}

function Set-StatutJob {
    param([string]$Cle, [string]$Statut, [string]$Erreur = "")
    $patch = @{ statut = $Statut; dateTraitement = (Get-Date).ToString("o") }
    if ($Erreur) { $patch.erreur = $Erreur }
    $json = $patch | ConvertTo-Json -Compress
    try {
        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/relance_auto/$Cle.json" -Method Patch -Body $json -ContentType "application/json" -TimeoutSec 15 | Out-Null
    } catch {
        Write-Log "Echec ecriture statut Firebase pour $Cle : $($_.Exception.Message)" "ERREUR"
    }
}

function Update-JournalRelance {
    <#
    Incremente le compteur de relances pour ce fournisseur dans gromec_vba/relance_log,
    en miroir de markRelance() cote JS -- que la relance vienne du bouton manuel
    "Telecharger PDF + ouvrir Outlook" ou de l'envoi automatique.
    #>
    param([string]$Cle)
    try {
        $actuel = Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/relance_log/$Cle.json" -Method Get -TimeoutSec 10
    } catch { $actuel = $null }
    $count = 0
    if ($actuel -and $actuel.count) { $count = [int]$actuel.count }
    $nouveau = @{ count = $count + 1; lastSent = (Get-Date).ToString("o"); escalade = $false } | ConvertTo-Json -Compress
    try {
        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/relance_log/$Cle.json" -Method Put -Body $nouveau -ContentType "application/json" -TimeoutSec 15 | Out-Null
    } catch {
        Write-Log "Echec mise a jour du journal de relance pour $Cle : $($_.Exception.Message)" "ERREUR"
    }
}

function Invoke-TraitementFile {
    param($Outlook)

    $file = $null
    try {
        $file = Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/relance_auto.json" -Method Get -TimeoutSec 20
    } catch {
        Write-Log "Impossible de lire la file Firebase : $($_.Exception.Message)" "ERREUR"
        return
    }
    if ($null -eq $file) { return }

    $jobs = $file.PSObject.Properties | Where-Object { $_.Value.statut -eq "attente" }
    if (-not $jobs -or $jobs.Count -eq 0) { return }

    Write-Log "$($jobs.Count) relance(s) en attente dans la file."

    foreach ($jobProp in $jobs) {
        $cle = $jobProp.Name
        $job = $jobProp.Value

        $dest = $job.email
        $sujet = $job.subject
        $corps = $job.body
        $fournisseur = $job.fournisseur

        if (-not $dest) {
            Write-Log "Job '$cle' ($fournisseur) sans courriel destinataire -- ignore, marque en erreur." "AVERT"
            Set-StatutJob -Cle $cle -Statut "erreur" -Erreur "Pas de courriel destinataire"
            continue
        }

        if (-not $Envoyer) {
            Write-Log "[ESSAI - AUCUN ENVOI] Relance prete pour '$fournisseur' <$dest> -- sujet: $sujet -- PJ: $($job.pdfNom)" "ESSAI"
            continue
        }

        # ── Envoi reel via Outlook (uniquement si -Envoyer est fourni) ──
        $cheminPdf = $null
        try {
            $bytes = [Convert]::FromBase64String($job.pdfBase64)
            $nomPdf = if ($job.pdfNom) { $job.pdfNom } else { "Relance_$cle.pdf" }
            $cheminPdf = Join-Path $DossierTemp $nomPdf
            [System.IO.File]::WriteAllBytes($cheminPdf, $bytes)
        } catch {
            Write-Log "Echec decodage/ecriture PDF pour '$fournisseur' : $($_.Exception.Message)" "ERREUR"
            Set-StatutJob -Cle $cle -Statut "erreur" -Erreur "PDF invalide : $($_.Exception.Message)"
            continue
        }

        try {
            $mail = $Outlook.CreateItem(0)  # olMailItem
            $mail.To = $dest
            $mail.Subject = $sujet
            $mail.Body = $corps
            if ($cheminPdf -and (Test-Path $cheminPdf)) {
                $mail.Attachments.Add($cheminPdf) | Out-Null
            }
            $mail.Send()
            Write-Log "Relance envoyee a '$fournisseur' <$dest> (BC: $($job.orders -join ', '))." "OK"
            Set-StatutJob -Cle $cle -Statut "envoye"
            Update-JournalRelance -Cle $cle
        } catch {
            Write-Log "Echec envoi Outlook pour '$fournisseur' : $($_.Exception.Message)" "ERREUR"
            Set-StatutJob -Cle $cle -Statut "erreur" -Erreur $_.Exception.Message
        } finally {
            if ($cheminPdf -and (Test-Path $cheminPdf)) { Remove-Item $cheminPdf -Force -ErrorAction SilentlyContinue }
        }
    }
}

# =====================================================================
# PROGRAMME PRINCIPAL
# =====================================================================

if ($Envoyer) {
    Write-Log "RelanceFournisseur.ps1 demarre en mode ENVOI REEL. Les courriels seront envoyes via Outlook." "AVERT"
} else {
    Write-Log "RelanceFournisseur.ps1 demarre en mode ESSAI (aucun envoi). Lancer avec -Envoyer pour activer l'envoi reel." "INFO"
}

$Script:MutexInstance = New-Object System.Threading.Mutex($false, "Global\GromecRelanceFournisseur_InstanceUnique")
if (-not $Script:MutexInstance.WaitOne(0)) {
    Write-Log "Une autre instance de RelanceFournisseur.ps1 tourne deja -- arret de cette copie." "AVERT"
    exit 0
}

try {
    $outlook = $null
    if ($Envoyer) {
        $outlook = New-Object -ComObject Outlook.Application
    }

    do {
        Invoke-TraitementFile -Outlook $outlook
        if (-not $UneFois) { Start-Sleep -Seconds $IntervalleSecondes }
    } while (-not $UneFois)
} finally {
    $Script:MutexInstance.ReleaseMutex()
}
