# =====================================================================
# VerifierConfirmation.ps1
# Verification automatique des confirmations de commande fournisseurs
# Lance de maniere asynchrone depuis Outlook VBA (Shell, sans attente)
# -- Outlook ne gele jamais, tout le travail se fait ici, en dehors
# de son fil d'execution.
#
# Utilisation:
#   powershell -ExecutionPolicy Bypass -File VerifierConfirmation.ps1 -EntryID "..." -StoreID "..."
#   powershell -ExecutionPolicy Bypass -File VerifierConfirmation.ps1 -Interactive
#   powershell -ExecutionPolicy Bypass -File VerifierConfirmation.ps1 -EntryID "..." -StoreID "..." -Force
# =====================================================================

param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [switch]$Force,
    [switch]$Interactive
)

Add-Type -AssemblyName System.Windows.Forms

# =====================================================================
# MISE A JOUR AUTOMATIQUE depuis GitHub
# Verifie a chaque lancement si une version plus recente du script existe.
# Si oui: remplace ce fichier, relance avec les memes parametres, sort.
# En cas d'echec (pas de connexion, repo indisponible): continue
# normalement avec la version locale -- jamais bloquant.
# =====================================================================

$GitHubRawUrl = "https://raw.githubusercontent.com/dantibo1978-png/gromec-outlook-automation/main/VerifierConfirmation.ps1"

function Update-ScriptSiNecessaire {
    try {
        # Telechargement force en UTF-8 explicite (plutot que de se fier a
        # l'en-tete de reponse GitHub, qui peut etre ambigu) pour eviter toute
        # corruption des caracteres accentues avant meme la comparaison.
        $webClient = New-Object System.Net.WebClient
        $webClient.Encoding = [System.Text.Encoding]::UTF8
        $remoteContent = $webClient.DownloadString($GitHubRawUrl)
        $webClient.Dispose()
    } catch {
        return  # Pas de connexion ou GitHub indisponible -- on continue avec la version locale
    }

    if ([string]::IsNullOrWhiteSpace($remoteContent)) { return }

    $scriptPath = $PSCommandPath
    if ([string]::IsNullOrEmpty($scriptPath)) { return }  # Securite si lance autrement qu'en fichier

    try {
        $localContent = Get-Content -Path $scriptPath -Raw -Encoding UTF8
    } catch {
        return
    }

    # Normaliser les fins de ligne avant comparaison (evite les faux positifs)
    $normLocal  = ($localContent  -replace "`r`n", "`n").Trim()
    $normRemote = ($remoteContent -replace "`r`n", "`n").Trim()

    if ($normLocal -eq $normRemote) { return }  # Deja a jour

    try {
        # Ecriture en UTF-8 AVEC BOM explicite et unique (System.Text.Encoding::UTF8
        # genere toujours un seul BOM correct) -- evite le bug de double-BOM possible
        # avec Set-Content -Encoding UTF8 sur PowerShell 5.1.
        $utf8AvecBom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($scriptPath, $remoteContent, $utf8AvecBom)
    } catch {
        return  # Fichier verrouille ou inaccessible en ecriture -- continuer avec version locale
    }

    # Relancer avec les memes parametres que cet appel
    $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $scriptPath)
    if ($EntryID -ne "")  { $argList += @("-EntryID", $EntryID) }
    if ($StoreID -ne "")  { $argList += @("-StoreID", $StoreID) }
    if ($Force)           { $argList += "-Force" }
    if ($Interactive)     { $argList += "-Interactive" }

    Start-Process -FilePath "powershell.exe" -ArgumentList $argList -WindowStyle Hidden
    exit 0
}

Update-ScriptSiNecessaire

# --- Chemin du dossier de donnees (meme que la version VBA) ---
$DataFolder = "U:\GromecOutlook\"
$ConfigPath = Join-Path $DataFolder "config.json"

if (-not (Test-Path $ConfigPath)) {
    [System.Windows.Forms.MessageBox]::Show("Fichier config.json introuvable dans $DataFolder", "Erreur de configuration", "OK", "Error") | Out-Null
    exit 1
}

$Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

$ClaudeApiKey   = $Config.ClaudeApiKey
$ClaudeModel    = $Config.ClaudeModel
$ClaudeApiUrl   = $Config.ClaudeApiUrl
$FirebaseUrl    = $Config.FirebaseUrl
$JoursRecherche = $Config.JoursRechercheEnvoyes
$SeuilConfiance = $Config.SeuilConfiance

$FichierFournisseurs  = Join-Path $DataFolder "fournisseurs_appris.csv"
$FichierConversations = Join-Path $DataFolder "conversations_traitees.csv"
$FichierJournal       = Join-Path $DataFolder "journal_confirmations.csv"
$FichierRapportExcel  = Join-Path $DataFolder "Rapport_Confirmations.xlsx"
$FichierComportementCorps = Join-Path $DataFolder "comportement_corps.csv"

# Nombre de confirmations identiques consecutives (par adresse exacte) avant que
# le script applique automatiquement "verifier le corps" sans repasser par le
# dialogue manuel. Une reponse contraire remet le compteur du cote oppose a zero
# (voir Set-ReponseCorps).
$SeuilConfianceCorps = 5


# =====================================================================
# FONCTIONS - Firebase
# =====================================================================

function Write-Log {
    param([string]$Message)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$ts] $Message"
    try {
        $niveau = if ($Message -like "ERREUR*") { "erreur" } elseif ($Message -like "WARN*") { "warn" } else { "info" }
        $body = @{ ts = $ts; msg = $Message; niveau = $niveau; source = "Verifier" } | ConvertTo-Json -Compress
        Invoke-RestMethod -Uri "${FirebaseUrl}gromec_vba/logs.json" -Method Post -Body $body -ContentType "application/json" -TimeoutSec 3 | Out-Null
    } catch {}
}

function Get-FirebaseValue {
    param([string]$Chemin)
    try {
        $url = "$FirebaseUrl$Chemin.json"
        $rep = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 15
        return $rep
    } catch {
        return $null
    }
}

function Set-FirebaseValue {
    param([string]$Chemin, [string]$Valeur)
    try {
        $url = "$FirebaseUrl$Chemin.json"
        $jsonBody = $Valeur | ConvertTo-Json -Compress
        Invoke-RestMethod -Uri $url -Method Put -Body $jsonBody -ContentType "application/json; charset=utf-8" -TimeoutSec 15 | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Write-FirebaseHistorique {
    param([string]$Fournisseur, [string]$Sujet, [string]$StatutGlobal, $Resultats, [string]$Devise, [string]$NumeroCommande, [string]$EntryID = "", [string]$StoreID = "", [string]$HistoriqueId = "")

    if ($Resultats.Count -eq 0) { return }

    $articles = @()
    foreach ($r in $Resultats) {
        $articles += @{
            sapLigne   = $r.SapLineNbr
            sapArticle = $r.SapArticle
            sapCode    = $r.SapCodeManuf
            sapDesc    = $r.SapDesc
            sapQty     = $r.SapQty
            sapPrix    = $r.SapPrice
            pdfCode    = $r.PdfCode
            pdfQty     = $r.PdfQty
            pdfPrix    = $r.PdfUnit
            diffQty    = $r.DiffQty
            diffUnit   = $r.DiffUnit
            diffTotal  = $r.DiffTotal
            statut     = $r.Statut
            methode    = $r.Methode
            confiance  = $r.Confiance
        }
    }

    $nbEcarts = ($Resultats | Where-Object { $_.Statut -eq "ECART" }).Count
    $nbNonTrouves = ($Resultats | Where-Object { $_.Statut -eq "NON_TROUVE" }).Count

    $entree = @{
        date         = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
        fournisseur  = $Fournisseur
        sujet        = $Sujet
        statut       = $StatutGlobal
        devise       = $Devise
        numeroCommande = $NumeroCommande
        nbEcarts     = $nbEcarts
        nbNonTrouves = $nbNonTrouves
        articles     = $articles
        entryID      = $EntryID
        storeID      = $StoreID
        resolu       = $false
        syncedResolu = $false
    }

    try {
        if ($HistoriqueId -ne "") {
            # Remplace une entree existante (ex: NON_APPARIE corrigee manuellement)
            # plutot que d'en creer une nouvelle -- evite les doublons sur le dashboard
            $url = "${FirebaseUrl}gromec_vba/historique/$HistoriqueId.json"
            $jsonBody = $entree | ConvertTo-Json -Depth 10 -Compress
            Invoke-RestMethod -Uri $url -Method Put -Body $jsonBody -ContentType "application/json; charset=utf-8" -TimeoutSec 15 | Out-Null
        } else {
            $url = "${FirebaseUrl}gromec_vba/historique.json"
            $jsonBody = $entree | ConvertTo-Json -Depth 10 -Compress
            Invoke-RestMethod -Uri $url -Method Post -Body $jsonBody -ContentType "application/json; charset=utf-8" -TimeoutSec 15 | Out-Null
        }
    } catch {
        # Echec silencieux -- le rapport Excel local reste la source de verite principale
    }
}

function Update-FirebaseChamp {
    param([string]$Chemin, [hashtable]$Champs)
    try {
        $url = "$FirebaseUrl$Chemin.json"
        $jsonBody = $Champs | ConvertTo-Json -Compress
        Invoke-RestMethod -Uri $url -Method Patch -Body $jsonBody -ContentType "application/json; charset=utf-8" -TimeoutSec 15 | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Sync-ResolusVersOutlook {
    param($Namespace)

    try {
        $url = "${FirebaseUrl}gromec_vba/historique.json"
        $historique = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 15
    } catch {
        return  # Pas de connexion -- on reessaiera au prochain passage du script
    }

    if ($null -eq $historique) { return }

    foreach ($cle in $historique.PSObject.Properties.Name) {
        $entree = $historique.$cle

        $resolu = [bool]$entree.resolu
        $syncedResolu = [bool]$entree.syncedResolu

        if ($resolu -eq $syncedResolu) { continue }  # Deja synchronise, rien a faire

        $entryID = $entree.entryID
        $storeID = $entree.storeID
        if ([string]::IsNullOrEmpty($entryID)) { continue }  # Ancienne entree sans EntryID -- ignoree

        try {
            $mail = $Namespace.GetItemFromID($entryID, $storeID)
            if ($null -ne $mail) {
                Set-CategorieConfirmation $mail $resolu
            }
            # Marquer comme synchronise, que le courriel ait ete trouve ou non
            # (s'il a ete deplace/supprime, on evite de reessayer indefiniment)
            Update-FirebaseChamp "gromec_vba/historique/$cle" @{ syncedResolu = $resolu } | Out-Null
        } catch {
            # Courriel introuvable ou erreur Outlook -- on marque quand meme synchronise
            # pour eviter une boucle de tentatives infructueuses
            Update-FirebaseChamp "gromec_vba/historique/$cle" @{ syncedResolu = $resolu } | Out-Null
        }
    }
}

function Sync-ReessaisManuels {
    param($Namespace)

    # Cherche les entrees NON_APPARIE pour lesquelles un numero de BC a ete
    # fourni manuellement depuis le dashboard (aReessayer=true), et relance
    # la comparaison complete avec ce numero -- remplace l'entree existante
    # par le vrai resultat (succes ou nouvel echec documente)
    try {
        $url = "${FirebaseUrl}gromec_vba/historique.json"
        $historique = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 15
    } catch {
        return
    }

    if ($null -eq $historique) { return }

    foreach ($cle in $historique.PSObject.Properties.Name) {
        $entree = $historique.$cle

        if ($entree.statut -ne "NON_APPARIE") { continue }
        if ([bool]$entree.aReessayer -ne $true) { continue }

        $numeroBCManuel = $entree.numeroBCManuel
        if ([string]::IsNullOrEmpty($numeroBCManuel)) { continue }

        $entryID = $entree.entryID
        $storeID = $entree.storeID
        if ([string]::IsNullOrEmpty($entryID)) { continue }

        try {
            $mail = $Namespace.GetItemFromID($entryID, $storeID)
        } catch {
            $mail = $null
        }

        if ($null -eq $mail) {
            # Le courriel original n'existe plus -- on documente l'echec et on arrete d'essayer
            Update-FirebaseChamp "gromec_vba/historique/$cle" @{
                raisonEchec      = "COURRIEL_ORIGINAL_INTROUVABLE"
                aReessayer       = $false
                dateDernierEssai = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
            } | Out-Null
            continue
        }

        # Relance la comparaison complete avec le numero fourni manuellement,
        # en reutilisant le meme identifiant Firebase (remplace plutot que duplique)
        Invoke-TraiterComparaison $Namespace $mail $numeroBCManuel $cle
    }
}

function Get-DictionnaireFournisseurs {
    $dict = Get-FirebaseValue "gromec_vba/dict_fournisseurs"
    $regles = Get-FirebaseValue "gromec_vba/regles_generales"
    if ([string]::IsNullOrEmpty($dict)) {
        # Fallback minimal si Firebase inaccessible
        $dict = "* ASC: netUnit=colonne NET, qty=QUANTITY+B/O QTY`n* WIKA Instruments: code=colonne Item, qty=QtyUnit avant pcs, netUnit=Unit price`n* Boshart (Confirmation): code=Item Number/Customer PN, qty=Quantity, netUnit=Net Price`n* Goulds Pumps: prix souvent USD -- indiquer USD"
        $regles = "- qty = quantite commandee TOTALE (inclure back-order si present)`n- netUnit = prix UNITAIRE net final`n- code = code article du FOURNISSEUR (pas le code SAP Gromec)`n- Si devise USD detectee: ecrire USD, sinon CAD"
    }
    return @{ Dict = $dict; Regles = $regles }
}

function Add-FournisseurAuDictionnaire {
    param([string]$NouvelleLigne)
    $dictActuel = Get-FirebaseValue "gromec_vba/dict_fournisseurs"
    if ([string]::IsNullOrEmpty($dictActuel)) { return $false }
    $nouveauDict = "$dictActuel`n$NouvelleLigne"
    return (Set-FirebaseValue "gromec_vba/dict_fournisseurs" $nouveauDict)
}

# =====================================================================
# FONCTIONS - API Claude
# =====================================================================

function ConvertTo-Base64File {
    param([string]$Chemin)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Chemin)
        return [System.Convert]::ToBase64String($bytes)
    } catch {
        return ""
    }
}

function Invoke-ClaudeMessage {
    param([string]$SystemPrompt, [string]$UserPrompt)
    $body = @{
        model      = $ClaudeModel
        max_tokens = 2048
        system     = $SystemPrompt
        messages   = @(@{ role = "user"; content = $UserPrompt })
    } | ConvertTo-Json -Depth 10

    try {
        $headers = @{
            "x-api-key"         = $ClaudeApiKey
            "anthropic-version" = "2023-06-01"
        }
        $rep = Invoke-RestMethod -Uri $ClaudeApiUrl -Method Post -Headers $headers -Body $body -ContentType "application/json; charset=utf-8" -TimeoutSec 90
        $texte = ($rep.content | Where-Object { $_.type -eq "text" } | Select-Object -First 1).text
        return $texte
    } catch {
        return "ERREUR: $($_.Exception.Message)"
    }
}

function Invoke-ClaudeDocument {
    param([string]$Base64PDF, [string]$TextPrompt)
    $body = @{
        model      = $ClaudeModel
        max_tokens = 16000
        messages   = @(
            @{
                role    = "user"
                content = @(
                    @{
                        type   = "document"
                        source = @{
                            type       = "base64"
                            media_type = "application/pdf"
                            data       = $Base64PDF
                        }
                    },
                    @{
                        type = "text"
                        text = $TextPrompt
                    }
                )
            }
        )
    } | ConvertTo-Json -Depth 10

    try {
        $headers = @{
            "x-api-key"         = $ClaudeApiKey
            "anthropic-version" = "2023-06-01"
        }
        $rep = Invoke-RestMethod -Uri $ClaudeApiUrl -Method Post -Headers $headers -Body $body -ContentType "application/json; charset=utf-8" -TimeoutSec 120
        $texte = ($rep.content | Where-Object { $_.type -eq "text" } | Select-Object -First 1).text
        return $texte
    } catch {
        return "ERREUR: $($_.Exception.Message)"
    }
}

# =====================================================================
# FONCTIONS - Apprentissage (CSV) -- meme format que la version VBA
# =====================================================================

function Get-CompteursFournisseur {
    param([string]$Adresse)
    $nbOui = 0; $nbNon = 0
    if (Test-Path $FichierFournisseurs) {
        foreach ($ligne in Get-Content $FichierFournisseurs) {
            $champs = $ligne -split ","
            if ($champs.Count -ge 3 -and $champs[0].Trim().ToLower() -eq $Adresse.ToLower()) {
                $nbOui = [int]$champs[1]
                $nbNon = [int]$champs[2]
                break
            }
        }
    }
    return @{ Oui = $nbOui; Non = $nbNon }
}

function Get-StatutFournisseurConnu {
    param([string]$Adresse)
    $c = Get-CompteursFournisseur $Adresse
    if ($c.Oui -ge $SeuilConfiance -and $c.Non -eq 0) { return "OUI" }
    if ($c.Non -ge $SeuilConfiance -and $c.Oui -eq 0) { return "NON" }
    return "INCERTAIN"
}

function Set-ReponseFournisseur {
    param([string]$Adresse, [bool]$EstConfirmation)
    $c = Get-CompteursFournisseur $Adresse
    if ($EstConfirmation) { $c.Oui++ } else { $c.Non++ }

    $lignes = @()
    $trouve = $false
    if (Test-Path $FichierFournisseurs) {
        foreach ($ligne in Get-Content $FichierFournisseurs) {
            $champs = $ligne -split ","
            if ($champs.Count -ge 3 -and $champs[0].Trim().ToLower() -eq $Adresse.ToLower()) {
                $lignes += "$Adresse,$($c.Oui),$($c.Non)"
                $trouve = $true
            } else {
                $lignes += $ligne
            }
        }
    }
    if (-not $trouve) { $lignes += "$Adresse,$($c.Oui),$($c.Non)" }
    $lignes | Out-File -FilePath $FichierFournisseurs -Encoding ASCII -Force
}

# =====================================================================
# FONCTIONS - Apprentissage "verifier le corps" (par adresse EXACTE)
# Important: deux expediteurs du meme fournisseur peuvent avoir un
# comportement different (ex: une boite aux lettres d'accuses de
# reception automatiques vs une representante qui repond en texte
# libre dans le corps) -- la cle est donc l'adresse complete, jamais
# le domaine ou le nom du fournisseur.
# =====================================================================

function Get-CompteursCorps {
    param([string]$Adresse)
    $nbCorps = 0; $nbPdf = 0
    if (Test-Path $FichierComportementCorps) {
        foreach ($ligne in Get-Content $FichierComportementCorps) {
            $champs = $ligne -split ","
            if ($champs.Count -ge 3 -and $champs[0].Trim().ToLower() -eq $Adresse.ToLower()) {
                $nbCorps = [int]$champs[1]
                $nbPdf = [int]$champs[2]
                break
            }
        }
    }
    return @{ Corps = $nbCorps; Pdf = $nbPdf }
}

function Get-StatutCorpsConnu {
    param([string]$Adresse)
    $c = Get-CompteursCorps $Adresse
    if ($c.Corps -ge $SeuilConfianceCorps -and $c.Pdf -eq 0) { return "CORPS" }
    if ($c.Pdf -ge $SeuilConfianceCorps -and $c.Corps -eq 0) { return "PDF" }
    return "INCERTAIN"
}

function Set-ReponseCorps {
    param([string]$Adresse, [bool]$VerifierCorps)
    $c = Get-CompteursCorps $Adresse
    # Une reponse qui contredit le pattern etabli remet l'autre compteur a
    # zero plutot que de laisser les deux s'accumuler en parallele -- on
    # veut refleter le comportement le plus RECENT de cet expediteur.
    if ($VerifierCorps) { $c.Corps++; $c.Pdf = 0 } else { $c.Pdf++; $c.Corps = 0 }

    $lignes = @()
    $trouve = $false
    if (Test-Path $FichierComportementCorps) {
        foreach ($ligne in Get-Content $FichierComportementCorps) {
            $champs = $ligne -split ","
            if ($champs.Count -ge 3 -and $champs[0].Trim().ToLower() -eq $Adresse.ToLower()) {
                $lignes += "$Adresse,$($c.Corps),$($c.Pdf)"
                $trouve = $true
            } else {
                $lignes += $ligne
            }
        }
    }
    if (-not $trouve) { $lignes += "$Adresse,$($c.Corps),$($c.Pdf)" }
    $lignes | Out-File -FilePath $FichierComportementCorps -Encoding ASCII -Force
}

function Test-ConversationTraitee {
    param([string]$ConvID)
    if (-not (Test-Path $FichierConversations)) { return $false }
    $contenu = Get-Content $FichierConversations
    return ($contenu -contains $ConvID)
}

function Set-ConversationTraitee {
    param([string]$ConvID)
    Add-Content -Path $FichierConversations -Value $ConvID -Encoding ASCII
}

function Write-JournalEntry {
    param([string]$Expediteur, [string]$Statut, [string]$Details)
    $detailsPropre = $Details -replace "[\r\n]+", " | " -replace ",", ";"
    $horodatage = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $FichierJournal -Value "$horodatage,$Expediteur,$Statut,$detailsPropre" -Encoding ASCII
}

# =====================================================================
# FONCTIONS - Extraction PDF via Claude
# =====================================================================

function Get-NumeroBC {
    param([string]$Texte)
    # Cherche 7-8 chiffres commencant par 90 (ex: 9006545, 9010000)
    $m = [regex]::Match($Texte, '(?<![0-9])90[0-9]{5,6}(?![0-9])')
    if ($m.Success) { return $m.Value }
    return ""
}

function Get-ItemsFournisseur {
    param([string]$CheminPDF)
    $dictInfo = Get-DictionnaireFournisseurs
    $b64 = ConvertTo-Base64File $CheminPDF
    if ([string]::IsNullOrEmpty($b64)) { return $null }

    $prompt = @"
Tu es expert en verification de commandes industrielles pour Gromec Inc.
Lis ce bon de confirmation fournisseur et extrais TOUS les articles.

DICTIONNAIRE PAR FOURNISSEUR:
$($dictInfo.Dict)

REGLES GENERALES:
$($dictInfo.Regles)

Cherche aussi le numero de bon de commande GROMEC (pas le numero de confirmation
du fournisseur) -- format typique 90XXXXX (7-8 chiffres commencant par 90).
Ce numero peut apparaitre sous plusieurs etiquettes selon le fournisseur, par exemple:
"N bon de commande", "Bon de commande", "No de commande", "PO Number", "Purchase Order",
"Customer PO", "Order No", "Numero de commande", "Notre commande", "Votre commande".
CHAMP OBLIGATOIRE: le 5e champ de la ligne FOURNISSEUR doit TOUJOURS etre present,
meme vide (laisse-le vide entre les deux derniers | si vraiment introuvable, mais
cherche attentivement avant de conclure que c'est absent -- il est presque toujours
imprime quelque part sur le document, souvent pres du nom du client ou en en-tete.

Reponds STRICTEMENT dans ce format, rien d'autre:
FOURNISSEUR|NomFournisseur|NumConfirmationFournisseur|CAD_ou_USD|NumeroBCGromec
ARTICLE|1|CODE|5|10.5800|description complete de la ligne
ARTICLE|2|CODE2|3|25.0000|description
Si aucun article trouve: ecrire AUCUN_ARTICLE
"@

    $reponse = Invoke-ClaudeDocument $b64 $prompt
    if ($reponse -like "ERREUR:*") { return @{ Erreur = $reponse } }

    $items = @()
    $nomFournisseur = ""
    $devise = "CAD"
    $bcGromec = ""

    foreach ($ligne in ($reponse -split "`r`n|`n")) {
        $l = $ligne.Trim()
        if ($l.StartsWith("FOURNISSEUR|")) {
            $champs = $l -split '\|'
            if ($champs.Count -ge 5) {
                $nomFournisseur = $champs[1]
                $devise = $champs[3]
                $bcGromec = $champs[4].Trim()
            } elseif ($champs.Count -ge 4) {
                $nomFournisseur = $champs[1]
                $devise = $champs[3]
            }
        } elseif ($l.StartsWith("ARTICLE|")) {
            $champs = $l -split '\|'
            if ($champs.Count -ge 5) {
                $items += [PSCustomObject]@{
                    LineNbr = [int]$champs[1]
                    Code    = $champs[2].Trim()
                    Qty     = [int]$champs[3]
                    NetUnit = [double]($champs[4] -replace ",", ".")
                    RawLine = if ($champs.Count -ge 6) { $champs[5] } else { "" }
                }
            }
        }
    }

    return @{
        Items          = $items
        NomFournisseur = $nomFournisseur
        Devise         = $devise
        BCGromec       = $bcGromec
        Erreur         = $null
    }
}

function Get-ItemsFournisseurDepuisCorps {
    # Meme contrat de sortie que Get-ItemsFournisseur (Items/NomFournisseur/
    # Devise/BCGromec/Erreur) pour pouvoir reutiliser Find-TousLesMatches sans
    # aucune modification. Utilisee quand le fournisseur ecrit les ecarts en
    # texte libre dans le corps du courriel plutot que dans un PDF chiffre
    # (le PDF joint, s'il y en a un, n'est alors qu'une copie du PO original
    # et ne contient PAS les nouveaux prix/quantites).
    #
    # Important: on analyse SEULEMENT le dernier message recu (avant la
    # chaine citee en dessous, du genre "--- Forwarded Message ---" ou
    # "De: ... Envoye: ..."), jamais toute la chaine -- le format de
    # citation varie trop d'un client courriel a l'autre pour etre fiable,
    # et le dernier message est de toute facon celui qui contient la
    # reponse pertinente.
    param([string]$CorpsMessage, [string]$Sujet)

    $dictInfo = Get-DictionnaireFournisseurs

    $prompt = @"
Tu es expert en verification de commandes industrielles pour Gromec Inc.
Un fournisseur a repondu en TEXTE LIBRE dans le corps d'un courriel (pas dans
un PDF chiffre) pour signaler des ecarts sur une commande -- typiquement des
corrections de prix ou de quantite sur certaines lignes seulement.

Voici UNIQUEMENT le dernier message recu (ignore tout texte de citation/chaine
en dessous, comme "Forwarded Message", "De: ... Envoye: ...", ou des lignes
commencant par ">"):

SUJET: $Sujet

CORPS:
$CorpsMessage

DICTIONNAIRE PAR FOURNISSEUR:
$($dictInfo.Dict)

REGLES GENERALES:
$($dictInfo.Regles)

Cherche le numero de bon de commande GROMEC (format 90XXXXX, 7-8 chiffres
commencant par 90) s'il apparait dans ce texte.

Pour chaque article mentionne avec un ecart, identifie le code (souvent le
CODE MANUF du fournisseur, pas le code SAP Gromec), la quantite si mentionnee
(laisse vide/0 si seul le prix est corrige), et le nouveau prix unitaire si
mentionne (laisse vide si seule la quantite est corrigee).

Reponds STRICTEMENT dans ce format, rien d'autre:
FOURNISSEUR|NomFournisseur||CAD_ou_USD|NumeroBCGromec
ARTICLE|1|CODE|QTE|PRIX|texte original de la ligne
ARTICLE|2|CODE2|QTE|PRIX|texte original de la ligne
Si aucun article avec ecart trouve dans ce texte: ecrire AUCUN_ARTICLE
Si la quantite n'est pas mentionnee pour une ligne, ecris 0 dans le champ QTE.
Si le prix n'est pas mentionne pour une ligne, ecris 0 dans le champ PRIX.
"@

    $reponse = Invoke-ClaudeMessage "" $prompt
    if ($reponse -like "ERREUR:*") { return @{ Erreur = $reponse } }

    $items = @()
    $nomFournisseur = ""
    $devise = "CAD"
    $bcGromec = ""

    foreach ($ligne in ($reponse -split "`r`n|`n")) {
        $l = $ligne.Trim()
        if ($l.StartsWith("FOURNISSEUR|")) {
            $champs = $l -split '\|'
            if ($champs.Count -ge 5) {
                $nomFournisseur = $champs[1]
                $devise = $champs[3]
                $bcGromec = $champs[4].Trim()
            }
        } elseif ($l.StartsWith("ARTICLE|")) {
            $champs = $l -split '\|'
            if ($champs.Count -ge 5) {
                $items += [PSCustomObject]@{
                    LineNbr = [int]$champs[1]
                    Code    = $champs[2].Trim()
                    Qty     = [int]([double]($champs[3] -replace ",", "."))
                    NetUnit = [double]($champs[4] -replace ",", ".")
                    RawLine = if ($champs.Count -ge 6) { $champs[5] } else { "" }
                }
            }
        }
    }

    return @{
        Items          = $items
        NomFournisseur = $nomFournisseur
        Devise         = $devise
        BCGromec       = $bcGromec
        Erreur         = $null
    }
}

function Get-ItemsCommandeGromec {
    param([string]$CheminPDF)
    $b64 = ConvertTo-Base64File $CheminPDF
    if ([string]::IsNullOrEmpty($b64)) { return $null }

    $prompt = @"
Tu lis un bon de commande Gromec Inc. envoye a un fournisseur.
Extrais chaque ligne article.

Format STRICT -- une ligne par article:
LIGNE|1|NumArticleSAP|CodeManuf|Description|Quantite|PrixUnitaire
- NumArticleSAP: numero SAP 10 chiffres (ex: 0001234567)
- CodeManuf: code fabricant si present, sinon laisser vide
- PrixUnitaire: prix unitaire 4 decimales
- Rien d'autre que les lignes LIGNE|
"@

    $reponse = Invoke-ClaudeDocument $b64 $prompt
    if ($reponse -like "ERREUR:*") { return @{ Erreur = $reponse } }

    $items = @()
    foreach ($ligne in ($reponse -split "`r`n|`n")) {
        $l = $ligne.Trim()
        if ($l.StartsWith("LIGNE|")) {
            $champs = $l -split '\|'
            if ($champs.Count -ge 7) {
                $items += [PSCustomObject]@{
                    LineNbr   = [int]$champs[1]
                    Article   = $champs[2].Trim()
                    CodeManuf = $champs[3].Trim()
                    Desc      = $champs[4].Trim()
                    Qty       = [int]$champs[5]
                    Price     = [double]($champs[6] -replace ",", ".")
                    RawLine   = $l
                }
            }
        }
    }
    return @{ Items = $items; Erreur = $null }
}

# =====================================================================
# FONCTIONS - Matching (deterministe + IA en secours)
# =====================================================================

function Format-CodeNormalise {
    param([string]$Code)
    if ([string]::IsNullOrEmpty($Code)) { return "" }
    $r = $Code.ToUpper().Trim() -replace "-", "" -replace " ", ""
    $r = $r.TrimStart("0")
    if ($r -eq "") { $r = "0" }
    return $r
}

function Find-MatchesDeterministe {
    param($SapItems, $PdfItems)

    $resultats = @()
    $pdfUtilises = @{}
    $sapNonTrouves = @()
    $tolerance = 0.015

    foreach ($sap in $SapItems) {
        $matchIdx = -1

        # 1. Code exact
        for ($j = 0; $j -lt $PdfItems.Count; $j++) {
            if ($pdfUtilises.ContainsKey($j)) { continue }
            if ($sap.CodeManuf -ne "" -and $PdfItems[$j].Code -ne "") {
                if ($sap.CodeManuf.ToUpper() -eq $PdfItems[$j].Code.ToUpper()) { $matchIdx = $j; break }
            }
        }

        # 2. Code normalise
        if ($matchIdx -eq -1) {
            $nSap = Format-CodeNormalise $sap.CodeManuf
            if ($nSap.Length -ge 4) {
                for ($j = 0; $j -lt $PdfItems.Count; $j++) {
                    if ($pdfUtilises.ContainsKey($j)) { continue }
                    if ((Format-CodeNormalise $PdfItems[$j].Code) -eq $nSap) { $matchIdx = $j; break }
                }
            }
        }

        # 3. Code PDF trouve dans la ligne brute / description SAP
        if ($matchIdx -eq -1) {
            for ($j = 0; $j -lt $PdfItems.Count; $j++) {
                if ($pdfUtilises.ContainsKey($j)) { continue }
                $nCode = Format-CodeNormalise $PdfItems[$j].Code
                if ($nCode.Length -ge 6) {
                    if ($sap.RawLine.ToUpper().Contains($nCode) -or $sap.Desc.ToUpper().Contains($nCode)) {
                        $matchIdx = $j; break
                    }
                }
            }
        }

        # 4. Suffix (7 derniers caracteres)
        if ($matchIdx -eq -1) {
            $nSap = Format-CodeNormalise $sap.CodeManuf
            if ($nSap.Length -ge 7) {
                $suffSap = $nSap.Substring($nSap.Length - 7)
                for ($j = 0; $j -lt $PdfItems.Count; $j++) {
                    if ($pdfUtilises.ContainsKey($j)) { continue }
                    $nPdf = Format-CodeNormalise $PdfItems[$j].Code
                    if ($nPdf.Length -ge 7 -and $nPdf.Substring($nPdf.Length - 7) -eq $suffSap) { $matchIdx = $j; break }
                }
            }
        }

        if ($matchIdx -ge 0) {
            $pdfUtilises[$matchIdx] = $true
            $pdf = $PdfItems[$matchIdx]
            $dU = $pdf.NetUnit - $sap.Price
            $dQ = $pdf.Qty - $sap.Qty
            $statut = if ([Math]::Abs($dU) -gt $tolerance -or $dQ -ne 0) { "ECART" } else { "OK" }
            $resultats += [PSCustomObject]@{
                SapLineNbr = $sap.LineNbr; SapArticle = $sap.Article; SapCodeManuf = $sap.CodeManuf
                SapDesc = $sap.Desc; SapQty = $sap.Qty; SapPrice = $sap.Price
                PdfCode = $pdf.Code; PdfQty = $pdf.Qty; PdfUnit = $pdf.NetUnit
                DiffUnit = $dU; DiffTotal = $dU * $sap.Qty; DiffQty = $dQ
                Statut = $statut; Methode = "Code"; Confiance = 1.0; EstIA = $false
            }
        } else {
            $sapNonTrouves += $sap
        }
    }

    $pdfNonUtilises = @()
    for ($j = 0; $j -lt $PdfItems.Count; $j++) {
        if (-not $pdfUtilises.ContainsKey($j)) { $pdfNonUtilises += $PdfItems[$j] }
    }

    return @{ Resultats = $resultats; SapNonTrouves = $sapNonTrouves; PdfNonUtilises = $pdfNonUtilises }
}

function Find-MatchesIA {
    param($SapNonTrouves, $PdfNonUtilises)

    if ($SapNonTrouves.Count -eq 0 -or $PdfNonUtilises.Count -eq 0) { return @() }

    $sl = ($SapNonTrouves | ForEach-Object { $i = $SapNonTrouves.IndexOf($_); "SAP[$i] Art:$($_.Article) Code:$($_.CodeManuf) Desc:$($_.Desc.Substring(0,[Math]::Min(50,$_.Desc.Length))) Prix:$($_.Price)" }) -join "`n"
    $pl = ($PdfNonUtilises | ForEach-Object { $i = $PdfNonUtilises.IndexOf($_); "PDF[$i] Code:$($_.Code) Qte:$($_.Qty) Unit:$($_.NetUnit) Desc:$($_.RawLine.Substring(0,[Math]::Min(40,$_.RawLine.Length)))" }) -join "`n"

    $prompt = @"
Expert commandes industrielles. Trouve correspondances SAP<->PDF pour les articles NON APPARIES.

STRATEGIES:
1. Code similaire (zeros en tete, tirets ignores)
2. Description semantique: TEE=TE, ELBOW=COUDE, REDUCER=REDUCTEUR, CAP=CAPUCHON, COUPLING=MANCHON, BUSHING=REDUCTION, WYE=YEE, BEND=COUDE, SS316=INOX316, BLK=NOIR, GALV=GALVANISE
3. Si 1 SAP non trouve ET 1 PDF non utilise -> forcer match, confiance 0.82 minimum
4. Cherche le numero de piece dans la DESCRIPTION SAP

SAP:
$sl

PDF:
$pl

Reponds UNIQUEMENT dans ce format, une ligne par match:
MATCH|sap_index|pdf_index|confiance|raison courte
Exemple: MATCH|0|2|0.95|meme code normalise
"@

    $reponse = Invoke-ClaudeMessage "" $prompt
    $matches = @()
    foreach ($ligne in ($reponse -split "`r`n|`n")) {
        $l = $ligne.Trim()
        if ($l.StartsWith("MATCH|")) {
            $champs = $l -split '\|'
            if ($champs.Count -ge 5) {
                $si = [int]$champs[1]; $pi = [int]$champs[2]
                $conf = [double]($champs[3] -replace ",", ".")
                if ($conf -ge 0.6 -and $si -ge 0 -and $si -lt $SapNonTrouves.Count -and $pi -ge 0 -and $pi -lt $PdfNonUtilises.Count) {
                    $sap = $SapNonTrouves[$si]; $pdf = $PdfNonUtilises[$pi]
                    $dU = $pdf.NetUnit - $sap.Price
                    $dQ = $pdf.Qty - $sap.Qty
                    $tolerance = 0.015
                    $statut = if ([Math]::Abs($dU) -gt $tolerance -or $dQ -ne 0) { "ECART" } else { "OK" }
                    $matches += [PSCustomObject]@{
                        SapLineNbr = $sap.LineNbr; SapArticle = $sap.Article; SapCodeManuf = $sap.CodeManuf
                        SapDesc = $sap.Desc; SapQty = $sap.Qty; SapPrice = $sap.Price
                        PdfCode = $pdf.Code; PdfQty = $pdf.Qty; PdfUnit = $pdf.NetUnit
                        DiffUnit = $dU; DiffTotal = $dU * $sap.Qty; DiffQty = $dQ
                        Statut = $statut; Methode = "IA: $($champs[4])"; Confiance = $conf; EstIA = $true
                    }
                }
            }
        }
    }
    return $matches
}

function Find-TousLesMatches {
    param($SapItems, $PdfItems)
    $det = Find-MatchesDeterministe $SapItems $PdfItems
    $resultats = @($det.Resultats)
    $iaMatches = Find-MatchesIA $det.SapNonTrouves $det.PdfNonUtilises
    $resultats += $iaMatches

    # Ajouter les SAP toujours non trouves apres IA (en NON_TROUVE)
    $sapMatches = $resultats.SapArticle
    foreach ($sap in $det.SapNonTrouves) {
        if ($sapMatches -notcontains $sap.Article) {
            $resultats += [PSCustomObject]@{
                SapLineNbr = $sap.LineNbr; SapArticle = $sap.Article; SapCodeManuf = $sap.CodeManuf
                SapDesc = $sap.Desc; SapQty = $sap.Qty; SapPrice = $sap.Price
                PdfCode = ""; PdfQty = 0; PdfUnit = 0
                DiffUnit = 0; DiffTotal = 0; DiffQty = 0
                Statut = "NON_TROUVE"; Methode = ""; Confiance = 0; EstIA = $false
            }
        }
    }
    return $resultats
}

# =====================================================================
# FONCTIONS - Outlook (recherche Envoyes, categories)
# =====================================================================

function Find-CourrielEnvoyeCorrespondant {
    param($Namespace, $MailConfirmation, [string]$NumeroBC)

    $sentFolder = $Namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
    $items = $sentFolder.Items
    $items.Sort("[SentOn]", $true)  # plus recent en premier

    $limiteDate = (Get-Date).AddDays(-$JoursRecherche)

    if ($NumeroBC -ne "") {
        # Etape 1: chercher TOUS les courriels avec le BC dans sujet/corps des envoyes
        # et retourner le PLUS VIEUX (premier courriel de la chaine = BC original avec PDF)
        $candidats = @()
        foreach ($item in $items) {
            if ($item.Class -ne 43) { continue }  # 43 = olMail
            if ($item.SentOn -lt $limiteDate) { break }
            if ($item.Attachments.Count -gt 0) {
                if ($item.Subject -like "*$NumeroBC*" -or $item.Body -like "*$NumeroBC*") {
                    $candidats += $item
                }
            }
        }
        if ($candidats.Count -gt 0) {
            # Prendre le plus vieux (dernier dans la liste triee recent->vieux)
            return $candidats[-1]
        }
    }

    # Etape 2: fallback par adresse email du fournisseur
    $adresseFournisseur = $MailConfirmation.SenderEmailAddress
    foreach ($item in $items) {
        if ($item.Class -ne 43) { continue }
        if ($item.SentOn -lt $limiteDate) { break }
        if ($item.Attachments.Count -eq 0) { continue }
        foreach ($dest in $item.Recipients) {
            if ($dest.Address -eq $adresseFournisseur) { return $item }
        }
    }

    return $null
}

function Save-PremierePDF {
    param($MailItem)
    foreach ($piece in $MailItem.Attachments) {
        if ($piece.FileName -like "*.pdf") {
            $chemin = Join-Path $env:TEMP "ps_$([guid]::NewGuid().ToString('N').Substring(0,8))_$($piece.FileName)"
            $piece.SaveAsFile($chemin)
            return $chemin
        }
    }
    return ""
}

function Save-PDFConfirmationFournisseur {
    # Retourne la PJ PDF qui est la confirmation fournisseur (pas le BC Gromec).
    # Logique : si une seule PJ PDF -> la prendre.
    # Si plusieurs PJ PDF -> exclure celle dont le nom ressemble a un BC Gromec
    # (contient "Commande fournisseur" ou "Purchase Order" ou "Bon de commande").
    # En dernier recours, prendre la derniere PJ PDF (le BC Gromec est generalement le premier).
    param($MailItem)

    # Patterns pour identifier les BC Gromec (a exclure)
    # Exemples: "Commande fournisseur - 9006904.pdf", "Purchase Order - 9006904.pdf"
    $motsCleBC = @("Commande fournisseur", "Purchase Order", "Bon de commande", "PO_", "BC_")

    $pdfs = @()
    foreach ($piece in $MailItem.Attachments) {
        if ($piece.FileName -like "*.pdf") {
            $pdfs += $piece
        }
    }

    if ($pdfs.Count -eq 0) { return "" }
    if ($pdfs.Count -eq 1) {
        $chemin = Join-Path $env:TEMP "ps_$([guid]::NewGuid().ToString('N').Substring(0,8))_$($pdfs[0].FileName)"
        $pdfs[0].SaveAsFile($chemin)
        return $chemin
    }

    # Plusieurs PDFs : chercher celui qui n'est PAS un BC Gromec
    $candidats = @()
    foreach ($piece in $pdfs) {
        $estBC = $false
        foreach ($mot in $motsCleBC) {
            if ($piece.FileName -like "*$mot*") { $estBC = $true; break }
        }
        if (-not $estBC) { $candidats += $piece }
    }

    # Si on a trouve des candidats non-BC, prendre le premier
    $choix = if ($candidats.Count -gt 0) { $candidats[0] } else { $pdfs[$pdfs.Count - 1] }
    $chemin = Join-Path $env:TEMP "ps_$([guid]::NewGuid().ToString('N').Substring(0,8))_$($choix.FileName)"
    $choix.SaveAsFile($chemin)
    Write-Log "INFO  PDF confirmation selectionne : $($choix.FileName)"
    return $chemin
}

function Set-CategorieConfirmation {
    param($MailItem, [bool]$EstOK)

    $catOK = "Confirmation OK"
    $catEcart = "Confirmation - Ecart"

    $catsActuelles = $MailItem.Categories
    if ([string]::IsNullOrEmpty($catsActuelles)) { $catsActuelles = "" }
    $catsActuelles = $catsActuelles -replace [regex]::Escape($catOK), "" -replace [regex]::Escape($catEcart), ""
    $catsActuelles = ($catsActuelles -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }) -join ", "

    $nouvelleCategorie = if ($EstOK) { $catOK } else { $catEcart }
    if ($catsActuelles -ne "") {
        $MailItem.Categories = "$catsActuelles, $nouvelleCategorie"
    } else {
        $MailItem.Categories = $nouvelleCategorie
    }
    $MailItem.Save()
}

# =====================================================================
# FONCTIONS - Rapport Excel (3 feuilles: Verification, Sommaire, Donnees_SAP)
# =====================================================================

function Write-RapportExcel {
    param([string]$Fournisseur, [string]$Sujet, [string]$StatutGlobal, $Resultats, [string]$Devise, [string]$NumeroCommande = "")

    if ($Resultats.Count -eq 0) { return }

    $excel = $null
    $classeur = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $fichierExistait = Test-Path $FichierRapportExcel

        if ($fichierExistait) {
            $classeur = $excel.Workbooks.Open($FichierRapportExcel)
            $fsVerif = $classeur.Sheets.Item("Verification")
            $fsSommaire = $classeur.Sheets.Item("Sommaire")
            $fsDonnees = $classeur.Sheets.Item("Donnees_SAP")
        } else {
            $classeur = $excel.Workbooks.Add()
            $fsVerif = $classeur.Sheets.Item(1)
            $fsVerif.Name = "Verification"
            $hdrsV = @("Date","Fournisseur","Sujet","Ligne","Article SAP","Code Manuf.","Description","Qte SAP","Qte PDF","Diff Qte","Prix SAP","Prix PDF","Diff Unit.","Diff Total","Methode","Confiance","Statut")
            for ($c = 0; $c -lt $hdrsV.Count; $c++) { $fsVerif.Cells.Item(1, $c + 1) = $hdrsV[$c] }
            $fsVerif.Range("A1:Q1").Font.Bold = $true
            $fsVerif.Range("A1:Q1").Interior.Color = 0x5C3A1A

            $fsSommaire = $classeur.Sheets.Add([System.Reflection.Missing]::Value, $fsVerif)
            $fsSommaire.Name = "Sommaire"
            $hdrsS = @("Date","Fournisseur","Sujet","Statut","Nb ecarts","Nb non trouves","Total SAP","Total PDF","Devise")
            for ($c = 0; $c -lt $hdrsS.Count; $c++) { $fsSommaire.Cells.Item(1, $c + 1) = $hdrsS[$c] }
            $fsSommaire.Range("A1:I1").Font.Bold = $true
            $fsSommaire.Range("A1:I1").Interior.Color = 0x5C3A1A

            $fsDonnees = $classeur.Sheets.Add([System.Reflection.Missing]::Value, $fsSommaire)
            $fsDonnees.Name = "Donnees_SAP"
            $hdrsD = @("Ligne SAP","Article","Prix PDF (a coller)","Qte PDF (a coller)","Statut","No Commande SAP")
            for ($c = 0; $c -lt $hdrsD.Count; $c++) { $fsDonnees.Cells.Item(1, $c + 1) = $hdrsD[$c] }
            $fsDonnees.Range("A1:F1").Font.Bold = $true
            $fsDonnees.Range("A1:F1").Interior.Color = 0x6B1A4A

            # Supprimer feuilles superflues
            for ($k = $classeur.Sheets.Count; $k -ge 1; $k--) {
                $nomF = $classeur.Sheets.Item($k).Name
                if ($nomF -ne "Verification" -and $nomF -ne "Sommaire" -and $nomF -ne "Donnees_SAP") {
                    $classeur.Sheets.Item($k).Delete()
                }
            }
        }

        # --- Lignes Verification (triees: ECART, NON_TROUVE, OK) ---
        $ligneV = $fsVerif.Cells.Item($fsVerif.Rows.Count, 1).End(-4162).Row + 1
        $nbEcarts = 0; $nbNonTrouves = 0; $totalSAP = 0.0; $totalPDF = 0.0

        $tries = @($Resultats | Where-Object { $_.Statut -eq "ECART" }) +
                 @($Resultats | Where-Object { $_.Statut -eq "NON_TROUVE" }) +
                 @($Resultats | Where-Object { $_.Statut -eq "OK" })

        foreach ($r in $tries) {
            if ($r.Statut -eq "ECART") { $nbEcarts++ }
            if ($r.Statut -eq "NON_TROUVE") { $nbNonTrouves++ }
            $totalSAP += $r.SapPrice * $r.SapQty
            if ($r.PdfUnit -gt 0) { $totalPDF += $r.PdfUnit * $r.SapQty }

            $fsVerif.Cells.Item($ligneV, 1)  = (Get-Date)
            $fsVerif.Cells.Item($ligneV, 2)  = $Fournisseur
            $fsVerif.Cells.Item($ligneV, 3)  = $Sujet
            $fsVerif.Cells.Item($ligneV, 4)  = $r.SapLineNbr
            $fsVerif.Cells.Item($ligneV, 5)  = $r.SapArticle
            $fsVerif.Cells.Item($ligneV, 6)  = $r.SapCodeManuf
            $fsVerif.Cells.Item($ligneV, 7)  = $r.SapDesc
            $fsVerif.Cells.Item($ligneV, 8)  = $r.SapQty
            if ($r.PdfQty -gt 0) { $fsVerif.Cells.Item($ligneV, 9) = $r.PdfQty }
            if ($r.Statut -ne "NON_TROUVE") { $fsVerif.Cells.Item($ligneV, 10) = $r.DiffQty }
            $fsVerif.Cells.Item($ligneV, 11) = $r.SapPrice
            if ($r.PdfUnit -gt 0) { $fsVerif.Cells.Item($ligneV, 12) = $r.PdfUnit }
            if ($r.Statut -ne "NON_TROUVE") {
                $fsVerif.Cells.Item($ligneV, 13) = $r.DiffUnit
                $fsVerif.Cells.Item($ligneV, 14) = $r.DiffTotal
            }
            $fsVerif.Cells.Item($ligneV, 15) = $r.Methode
            if ($r.Confiance -gt 0) { $fsVerif.Cells.Item($ligneV, 16) = $r.Confiance }
            $fsVerif.Cells.Item($ligneV, 17) = $r.Statut

            $fsVerif.Range("K$ligneV`:N$ligneV").NumberFormat = "#,##0.0000"

            $couleur = switch ($r.Statut) {
                "ECART"      { 15001066 }  # rouge clair (BGR)
                "NON_TROUVE" { 14803425 }  # jaune clair
                default      { 15264488 }  # vert clair
            }
            $fsVerif.Range("A$ligneV`:Q$ligneV").Interior.Color = $couleur
            $ligneV++
        }
        $fsVerif.Columns.AutoFit() | Out-Null

        # --- Ligne Sommaire ---
        $ligneS = $fsSommaire.Cells.Item($fsSommaire.Rows.Count, 1).End(-4162).Row + 1
        $fsSommaire.Cells.Item($ligneS, 1) = (Get-Date)
        $fsSommaire.Cells.Item($ligneS, 2) = $Fournisseur
        $fsSommaire.Cells.Item($ligneS, 3) = $Sujet
        $fsSommaire.Cells.Item($ligneS, 4) = $StatutGlobal
        $fsSommaire.Cells.Item($ligneS, 5) = $nbEcarts
        $fsSommaire.Cells.Item($ligneS, 6) = $nbNonTrouves
        $fsSommaire.Cells.Item($ligneS, 7) = $totalSAP
        $fsSommaire.Cells.Item($ligneS, 8) = $totalPDF
        $fsSommaire.Cells.Item($ligneS, 9) = $Devise
        $couleurS = if ($StatutGlobal -eq "OK") { 15264488 } else { 15001066 }
        $fsSommaire.Range("A$ligneS`:I$ligneS").Interior.Color = $couleurS
        $fsSommaire.Columns.AutoFit() | Out-Null

        # --- Feuille Donnees_SAP (triee par ligne SAP) ---
        $ligneD = $fsDonnees.Cells.Item($fsDonnees.Rows.Count, 1).End(-4162).Row + 1
        $triesParLigne = $Resultats | Sort-Object SapLineNbr
        foreach ($r in $triesParLigne) {
            $fsDonnees.Cells.Item($ligneD, 1) = $r.SapLineNbr
            $fsDonnees.Cells.Item($ligneD, 2) = $r.SapArticle
            $fsDonnees.Cells.Item($ligneD, 3) = if ($r.PdfUnit -gt 0) { $r.PdfUnit } else { $r.SapPrice }
            $fsDonnees.Cells.Item($ligneD, 4) = if ($r.PdfQty -gt 0) { $r.PdfQty } else { $r.SapQty }
            $fsDonnees.Cells.Item($ligneD, 5) = $r.Statut
            $fsDonnees.Cells.Item($ligneD, 6) = $NumeroCommande
            $fsDonnees.Cells.Item($ligneD, 3).NumberFormat = "0.0000"
            if ($r.Statut -eq "ECART") { $fsDonnees.Range("A$ligneD`:F$ligneD").Interior.Color = 15001066 }
            $ligneD++
        }
        $fsDonnees.Columns.AutoFit() | Out-Null

        if ($fichierExistait) {
            $classeur.Save()
        } else {
            $classeur.SaveAs($FichierRapportExcel, 51)  # 51 = xlOpenXMLWorkbook
        }
        $classeur.Close($false)
        $excel.Quit()
    } catch {
        Write-JournalEntry "" "ERREUR_EXCEL" $_.Exception.Message
        if ($classeur) { try { $classeur.Close($false) } catch {} }
        if ($excel) { try { $excel.Quit() } catch {} }
    } finally {
        if ($excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
        [System.GC]::Collect()
    }
}

# =====================================================================
# FONCTIONS - Apprentissage automatique du dictionnaire (fournisseur inconnu)
# =====================================================================

function Get-PropositionDictionnaire {
    param([string]$Base64PDF)

    $prompt = @"
Tu analyses un PDF de confirmation de commande d'un fournisseur INCONNU.
Determine le nom du fournisseur et la structure de ses colonnes.

Exemples du format de dictionnaire existant (pour que ta proposition soit dans le meme style):
* ASC: netUnit=colonne NET, qty=QUANTITY+B/O QTY
* Dahl Valve: netUnit=Unit Price, qty=Order Qty
* WIKA Instruments: code=colonne Item, qty=QtyUnit avant pcs, netUnit=Unit price

Identifie pour CE document:
- code: quelle colonne contient le code article du FOURNISSEUR
- qty: quelle colonne contient la quantite commandee
- netUnit: quelle colonne contient le prix unitaire net final
- devise: CAD ou USD

Reponds STRICTEMENT dans ce format, rien d'autre:
NOM|NomDuFournisseur
ENTREE|* NomDuFournisseur: netUnit=..., qty=..., code=...
Si tu n'es pas en mesure de determiner le format avec confiance, reponds uniquement: IMPOSSIBLE
"@

    $reponse = Invoke-ClaudeDocument $Base64PDF $prompt
    if ($reponse -match "IMPOSSIBLE") { return $null }

    $nom = ""; $entree = ""
    foreach ($ligne in ($reponse -split "`r`n|`n")) {
        $l = $ligne.Trim()
        if ($l.StartsWith("NOM|")) { $nom = $l.Substring(4) }
        if ($l.StartsWith("ENTREE|")) { $entree = $l.Substring(7) }
    }
    if ($entree -eq "") { return $null }
    return @{ Nom = $nom; Entree = $entree }
}

function Invoke-ProposerFournisseur {
    param([string]$CheminPDF)
    $b64 = ConvertTo-Base64File $CheminPDF
    if ([string]::IsNullOrEmpty($b64)) { return $false }

    $prop = Get-PropositionDictionnaire $b64
    if ($null -eq $prop) { return $false }

    $question = "Fournisseur inconnu detecte: $($prop.Nom)`n`nClaude propose d'ajouter cette regle au dictionnaire:`n`n$($prop.Entree)`n`nAjouter cette regle de facon permanente (Firebase)?"
    $rep = [System.Windows.Forms.MessageBox]::Show($question, "Nouveau fournisseur - proposition", "YesNo", "Question")

    if ($rep -eq "Yes") {
        $ok = Add-FournisseurAuDictionnaire $prop.Entree
        if ($ok) {
            [System.Windows.Forms.MessageBox]::Show("Regle ajoutee. Nouvelle tentative d'extraction en cours...", "OK", "OK", "Information") | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("Echec de l'ecriture dans Firebase. Verifiez la connexion.", "Erreur", "OK", "Warning") | Out-Null
        }
        return $ok
    }
    return $false
}

# =====================================================================
# FONCTION - Orchestration complete de la comparaison
# =====================================================================

function Write-FirebaseEchec {
    param($MailConfirmation, [string]$RaisonEchec, [string]$NumeroBCRecherche = "", [string]$NomFournisseurDetecte = "", [string]$HistoriqueId = "")

    if ($HistoriqueId -ne "") {
        # Mise a jour d'une entree existante (re-tentative manuelle echouee)
        # -- on garde l'historique de l'essai sans creer de doublon
        Update-FirebaseChamp "gromec_vba/historique/$HistoriqueId" @{
            raisonEchec       = $RaisonEchec
            numeroBCRecherche = $NumeroBCRecherche
            aReessayer        = $false
            dateDernierEssai  = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
        } | Out-Null
        return
    }

    $entree = @{
        date            = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
        fournisseur     = if ($NomFournisseurDetecte -ne "") { $NomFournisseurDetecte } else { $MailConfirmation.SenderEmailAddress }
        sujet           = $MailConfirmation.Subject
        statut          = "NON_APPARIE"
        raisonEchec     = $RaisonEchec
        numeroBCRecherche = $NumeroBCRecherche
        entryID         = $MailConfirmation.EntryID
        storeID         = $MailConfirmation.Parent.StoreID
        articles        = @()
        aReessayer      = $false
        numeroBCManuel  = ""
    }

    try {
        $url = "${FirebaseUrl}gromec_vba/historique.json"
        $jsonBody = $entree | ConvertTo-Json -Depth 10 -Compress
        Invoke-RestMethod -Uri $url -Method Post -Body $jsonBody -ContentType "application/json; charset=utf-8" -TimeoutSec 15 | Out-Null
    } catch {
        # Echec silencieux -- le journal local CSV reste la source de verite principale
    }
}

function Invoke-TraiterComparaison {
    param($Namespace, $MailConfirmation, [string]$NumeroBCOverride = "", [string]$HistoriqueId = "", [Nullable[bool]]$VerifierCorps = $null)

    $sujet = $MailConfirmation.Subject
    $expediteur = $MailConfirmation.SenderEmailAddress

    # Si l'appelant n'a pas precise explicitement le mode (ex: re-tentative
    # manuelle depuis le dashboard via Sync-ReessaisManuels), on se fie au
    # comportement appris pour cette adresse exacte -- par defaut PDF si
    # encore incertain, pour ne rien changer au comportement existant.
    if ($null -eq $VerifierCorps) {
        $VerifierCorps = (Get-StatutCorpsConnu $expediteur) -eq "CORPS"
    }

    if ($VerifierCorps) {
        # --- MODE CORPS: les ecarts sont dans le texte du courriel, pas dans
        # un PDF chiffre. Le PDF joint (s'il y en a un) n'est qu'une copie de
        # reference et n'est PAS utilise pour l'extraction des ecarts. ---

        $corpsMessage = $MailConfirmation.Body
        if ($corpsMessage.Length -gt 4000) { $corpsMessage = $corpsMessage.Substring(0, 4000) }

        $resFourn = Get-ItemsFournisseurDepuisCorps $corpsMessage $sujet
        if ($resFourn.Erreur) {
            Write-JournalEntry $expediteur "ERREUR_EXTRACTION_CORPS" $resFourn.Erreur
            Write-FirebaseEchec $MailConfirmation "ERREUR_EXTRACTION_CORPS" "" "" $HistoriqueId
            return
        }
        $itemsFourn = @($resFourn.Items)
        $nomFourn = $resFourn.NomFournisseur
        $devise = $resFourn.Devise
        $bcGromec = $resFourn.BCGromec

        if ($itemsFourn.Count -eq 0) {
            Write-JournalEntry $expediteur "AUCUN_ECART_DANS_CORPS" "Mode corps actif mais aucun article extrait du texte"
            Write-FirebaseEchec $MailConfirmation "AUCUN_ECART_DANS_CORPS" $bcGromec $nomFourn $HistoriqueId
            return
        }

        # --- Trouver le numero BC pour la recherche dans Envoyes ---
        if ($NumeroBCOverride -ne "") {
            $numeroBC = $NumeroBCOverride
        } else {
            $numeroBC = Get-NumeroBC "$sujet $($MailConfirmation.Body)"
            if ($numeroBC -eq "" -and $bcGromec -ne "") { $numeroBC = $bcGromec }
        }

        $mailEnvoye = Find-CourrielEnvoyeCorrespondant $Namespace $MailConfirmation $numeroBC
        if ($null -eq $mailEnvoye) {
            Write-JournalEntry $expediteur "AUCUNE_COMMANDE_TROUVEE" "BC recherche: $numeroBC"
            Write-FirebaseEchec $MailConfirmation "AUCUNE_COMMANDE_TROUVEE" $numeroBC $nomFourn $HistoriqueId
            return
        }

        $cheminCommande = Save-PremierePDF $mailEnvoye
        if ($cheminCommande -eq "") {
            Write-JournalEntry $expediteur "PIECE_JOINTE_PDF_MANQUANTE" "PDF commande Gromec introuvable"
            Write-FirebaseEchec $MailConfirmation "PIECE_JOINTE_PDF_MANQUANTE_COMMANDE" $numeroBC $nomFourn $HistoriqueId
            return
        }

        $resSAP = Get-ItemsCommandeGromec $cheminCommande
        $itemsSAP = @($resSAP.Items)
        Remove-Item $cheminCommande -Force -ErrorAction SilentlyContinue

        if ($itemsSAP.Count -eq 0) {
            Write-JournalEntry $expediteur "ARTICLES_NON_EXTRAITS" "Fourn(corps):$($itemsFourn.Count) SAP:$($itemsSAP.Count)"
            Write-FirebaseEchec $MailConfirmation "ARTICLES_NON_EXTRAITS" $numeroBC $nomFourn $HistoriqueId
            return
        }

        # Completer Qty/NetUnit manquants (0) avec les valeurs du PO Gromec --
        # le fournisseur, en mode corps, ne corrige souvent qu'UN seul des
        # deux champs (juste le prix, ou juste la quantite) sur une ligne.
        # Sans ce remplissage, un champ a 0 creerait un faux ecart.
        $itemsFournRempli = @()
        $codesMentionnes = @{}
        foreach ($itemFourn in $itemsFourn) {
            $sapCorrespondant = $itemsSAP | Where-Object {
                $_.CodeManuf -ne "" -and $itemFourn.Code -ne "" -and
                $_.CodeManuf.ToUpper() -eq $itemFourn.Code.ToUpper()
            } | Select-Object -First 1
            if ($null -eq $sapCorrespondant) {
                $nCode = Format-CodeNormalise $itemFourn.Code
                $sapCorrespondant = $itemsSAP | Where-Object {
                    (Format-CodeNormalise $_.CodeManuf) -eq $nCode -and $nCode.Length -ge 4
                } | Select-Object -First 1
            }
            if ($sapCorrespondant) {
                if ($itemFourn.Qty -eq 0) { $itemFourn.Qty = $sapCorrespondant.Qty }
                if ($itemFourn.NetUnit -eq 0) { $itemFourn.NetUnit = $sapCorrespondant.Price }
                $codesMentionnes[(Format-CodeNormalise $sapCorrespondant.CodeManuf)] = $true
            }
            $itemsFournRempli += $itemFourn
        }

        # Pour les lignes du PO que le fournisseur n'a PAS commentees dans le
        # corps (silence = pas d'ecart), on ajoute une "confirmation" synthetique
        # avec les valeurs SAP d'origine telles quelles. Find-TousLesMatches va
        # alors les matcher en statut OK automatiquement (diff = 0). Sans ca,
        # ces lignes seraient absentes de $Resultats et donc de Firebase -- ce
        # qui decale les boutons "Copier prix/quantites vers SAP" du dashboard,
        # qui s'attendent a UNE valeur par ligne du PO complet, dans l'ordre.
        foreach ($sap in $itemsSAP) {
            $nCodeSap = Format-CodeNormalise $sap.CodeManuf
            if (-not $codesMentionnes.ContainsKey($nCodeSap)) {
                $itemsFournRempli += [PSCustomObject]@{
                    LineNbr = 0
                    Code    = $sap.CodeManuf
                    Qty     = $sap.Qty
                    NetUnit = $sap.Price
                    RawLine = "(non mentionne dans le corps -- valeur SAP d'origine)"
                }
            }
        }
        $itemsFourn = $itemsFournRempli

    } else {
        # --- MODE PDF (comportement original, inchange) ---

        $cheminConfirmation = Save-PDFConfirmationFournisseur $MailConfirmation
        if ($cheminConfirmation -eq "") {
            Write-JournalEntry $expediteur "PIECE_JOINTE_PDF_MANQUANTE" "PDF confirmation introuvable"
            Write-FirebaseEchec $MailConfirmation "PIECE_JOINTE_PDF_MANQUANTE_CONFIRMATION" "" "" $HistoriqueId
            return
        }

        # --- Extraction fournisseur (1ere tentative) ---
        $resFourn = Get-ItemsFournisseur $cheminConfirmation
        $itemsFourn = @($resFourn.Items)
        $nomFourn = $resFourn.NomFournisseur
        $devise = $resFourn.Devise
        $bcGromec = $resFourn.BCGromec

        # --- Apprentissage automatique si fournisseur inconnu ---
        if ($itemsFourn.Count -eq 0) {
            $ajoute = Invoke-ProposerFournisseur $cheminConfirmation
            if ($ajoute) {
                $resFourn = Get-ItemsFournisseur $cheminConfirmation
                $itemsFourn = @($resFourn.Items)
                $nomFourn = $resFourn.NomFournisseur
                $devise = $resFourn.Devise
                $bcGromec = $resFourn.BCGromec
            }
        }

        # --- Trouver le numero BC pour la recherche dans Envoyes ---
        # Si un numero a ete fourni manuellement (re-tentative depuis le dashboard),
        # il a priorite absolue sur la detection automatique
        if ($NumeroBCOverride -ne "") {
            $numeroBC = $NumeroBCOverride
        } else {
            $numeroBC = Get-NumeroBC "$sujet $($MailConfirmation.Body)"
            if ($numeroBC -eq "" -and $bcGromec -ne "") { $numeroBC = $bcGromec }
        }

        $mailEnvoye = Find-CourrielEnvoyeCorrespondant $Namespace $MailConfirmation $numeroBC
        if ($null -eq $mailEnvoye) {
            Write-JournalEntry $expediteur "AUCUNE_COMMANDE_TROUVEE" "BC recherche: $numeroBC"
            Write-FirebaseEchec $MailConfirmation "AUCUNE_COMMANDE_TROUVEE" $numeroBC $nomFourn $HistoriqueId
            Remove-Item $cheminConfirmation -Force -ErrorAction SilentlyContinue
            return
        }

        $cheminCommande = Save-PremierePDF $mailEnvoye
        if ($cheminCommande -eq "") {
            Write-JournalEntry $expediteur "PIECE_JOINTE_PDF_MANQUANTE" "PDF commande Gromec introuvable"
            Write-FirebaseEchec $MailConfirmation "PIECE_JOINTE_PDF_MANQUANTE_COMMANDE" $numeroBC $nomFourn $HistoriqueId
            Remove-Item $cheminConfirmation -Force -ErrorAction SilentlyContinue
            return
        }

        $resSAP = Get-ItemsCommandeGromec $cheminCommande
        $itemsSAP = @($resSAP.Items)

        Remove-Item $cheminConfirmation -Force -ErrorAction SilentlyContinue
        Remove-Item $cheminCommande -Force -ErrorAction SilentlyContinue

        if ($itemsFourn.Count -eq 0 -or $itemsSAP.Count -eq 0) {
            Write-JournalEntry $expediteur "ARTICLES_NON_EXTRAITS" "Fourn:$($itemsFourn.Count) SAP:$($itemsSAP.Count)"
            Write-FirebaseEchec $MailConfirmation "ARTICLES_NON_EXTRAITS" $numeroBC $nomFourn $HistoriqueId
            return
        }
    }

    # --- Matching (commun aux deux modes) ---
    $resultats = Find-TousLesMatches $itemsSAP $itemsFourn

    # Tri par numero de ligne SAP -- respecte l'ordre original du PO envoye.
    # Important pour les boutons "Copier prix/quantites vers SAP" du dashboard,
    # qui collent une valeur par ligne dans l'ordre attendu par SAP. En mode
    # PDF c'est deja l'ordre naturel (no-op); en mode corps, les lignes
    # completees automatiquement (non mentionnees par le fournisseur) doivent
    # se remettre a leur place plutot que de rester a la fin.
    $resultats = @($resultats | Sort-Object SapLineNbr)

    $nbEcarts = ($resultats | Where-Object { $_.Statut -eq "ECART" }).Count
    $nbNonTrouves = ($resultats | Where-Object { $_.Statut -eq "NON_TROUVE" }).Count
    $estOK = ($nbEcarts -eq 0 -and $nbNonTrouves -eq 0)


    Set-CategorieConfirmation $MailConfirmation $estOK
    Write-JournalEntry $expediteur $(if ($estOK) { "OK" } else { "ECART" }) "Ecarts:$nbEcarts NonTrouves:$nbNonTrouves"
    Write-RapportExcel $nomFourn $sujet $(if ($estOK) { "OK" } else { "ECART" }) $resultats $devise $numeroBC
    Write-FirebaseHistorique $nomFourn $sujet $(if ($estOK) { "OK" } else { "ECART" }) $resultats $devise $numeroBC $MailConfirmation.EntryID $MailConfirmation.Parent.StoreID $HistoriqueId
}

# =====================================================================
# FONCTION - Traitement d'un nouveau courriel (classification + apprentissage)
# =====================================================================

function Test-DomaineExclu {
    param([string]$Adresse)
    # Toujours exclure les courriels internes Gromec
    if ($Adresse -like "*@gromec.com") { return $true }
    # Verifier le fichier domaines_exclus.csv
    $fichierExclus = Join-Path $DataFolder "domaines_exclus.csv"
    if (-not (Test-Path $fichierExclus)) { return $false }
    $domaine = ($Adresse -split "@")[-1].ToLower().Trim()
    foreach ($ligne in Get-Content $fichierExclus) {
        if ($ligne.Trim().ToLower() -eq $domaine) { return $true }
    }
    return $false
}

function Add-DomaineExclu {
    param([string]$Adresse)
    $fichierExclus = Join-Path $DataFolder "domaines_exclus.csv"
    $domaine = ($Adresse -split "@")[-1].ToLower().Trim()
    $existants = @()
    if (Test-Path $fichierExclus) { $existants = Get-Content $fichierExclus | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ -ne "" } }
    if ($existants -notcontains $domaine) {
        Add-Content -Path $fichierExclus -Value $domaine -Encoding UTF8
    }
    return $domaine
}

function Invoke-ClassifierCourriel {
    param($MailItem)

    $adresseExp = $MailItem.SenderEmailAddress

    # Filtre rapide -- domaines exclus et @gromec.com
    if (Test-DomaineExclu $adresseExp) {
        return @{ EstConfirmation = $false; Confiance = 1.0; VerifierCorps = $false; TexteBrut = "DOMAINE EXCLU"; Exclu = $true }
    }

    # Historique comme contexte seulement (pas comme filtre)
    $compteurs = Get-CompteursFournisseur $adresseExp
    $contexteHistorique = ""
    $totalConnu = $compteurs.Oui + $compteurs.Non
    if ($totalConnu -gt 0) {
        $contexteHistorique = "HISTORIQUE $adresseExp : $($compteurs.Oui) confirmation(s) passees, $($compteurs.Non) non-confirmation(s)."
    }

    $sysPrompt = @"
Tu es un classificateur de courriels pour Gromec Inc. (distributeur industriel, Quebec).
Tu dois determiner si un courriel est une CONFIRMATION DE COMMANDE fournisseur.

Analyse le sujet, le corps ET toutes les pieces jointes fournies.

Reponds a ces 5 questions par OUI ou NON, puis donne ta conclusion:
Q1_NUMERO_BC: Y a-t-il un numero de commande Gromec (format 9XXXXXX, ex: 9006906)?
Q2_PRIX_QTE: Y a-t-il des prix ou quantites confirmes en lien avec une commande?
Q3_DATE_LIVRAISON: Y a-t-il une date de livraison ou delai de livraison confirme?
Q4_ACCUSÉ_RECEPTION: Le fournisseur confirme-t-il explicitement avoir recu la commande?
Q5_DOCUMENT_COMMANDE: Les pieces jointes contiennent-elles un bon de commande ou une confirmation?

REGLE: C'est une confirmation si (Q1=OUI ou Q5=OUI) ET (Q4=OUI ou Q2=OUI).
N'EST PAS une confirmation: questions, devis seuls, factures seules, avis expedition seuls, newsletters, courriels generaux sans reference a une commande specifique.

Reponds EXACTEMENT en ce format:
Q1_NUMERO_BC: OUI/NON
Q2_PRIX_QTE: OUI/NON
Q3_DATE_LIVRAISON: OUI/NON
Q4_ACCUSÉ_RECEPTION: OUI/NON
Q5_DOCUMENT_COMMANDE: OUI/NON
CONFIRMATION: OUI/NON
CONFIANCE: 0.00
SOURCE: PDF/CORPS
"@

    $corps = $MailItem.Body
    if ($corps.Length -gt 3000) { $corps = $corps.Substring(0, 3000) }
    $nomsPJ = ($MailItem.Attachments | ForEach-Object { $_.FileName }) -join ", "

    $usrPrompt = "$contexteHistorique`n`nExpediteur: $($MailItem.SenderName) <$adresseExp>`nSujet: $($MailItem.Subject)`nPieces jointes: $nomsPJ`n`nCorps du courriel:`n$corps"

    # Construire les messages avec PJ PDF incluses
    $contenuMessages = @(@{ type = "text"; text = $usrPrompt })

    foreach ($pj in $MailItem.Attachments) {
        if ($pj.FileName -like "*.pdf") {
            try {
                $cheminTemp = Join-Path $env:TEMP "classif_$([guid]::NewGuid().ToString('N').Substring(0,8))_$($pj.FileName)"
                $pj.SaveAsFile($cheminTemp)
                $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($cheminTemp))
                Remove-Item $cheminTemp -Force -ErrorAction SilentlyContinue
                $contenuMessages += @{
                    type   = "document"
                    source = @{ type = "base64"; media_type = "application/pdf"; data = $b64 }
                }
            } catch {}
        }
    }

    $body = @{
        model      = "claude-haiku-4-5-20251001"
        max_tokens = 150
        system     = $sysPrompt
        messages   = @(@{ role = "user"; content = $contenuMessages })
    } | ConvertTo-Json -Depth 15

    try {
        $headers = @{ "x-api-key" = $ClaudeApiKey; "anthropic-version" = "2023-06-01"; "anthropic-beta" = "pdfs-2024-09-25" }
        $rep = Invoke-RestMethod -Uri $ClaudeApiUrl -Method Post -Headers $headers -Body $body -ContentType "application/json; charset=utf-8" -TimeoutSec 45
        $texte = ($rep.content | Where-Object { $_.type -eq "text" } | Select-Object -First 1).text

        $confirmation = if ($texte -match "CONFIRMATION:\s*(OUI|NON)") { $Matches[1] } else { "NON" }
        $confiance    = if ($texte -match "CONFIANCE:\s*([\d.]+)")     { [double]$Matches[1] } else { 0.5 }
        $source       = if ($texte -match "SOURCE:\s*(PDF|CORPS)")     { $Matches[1] } else { "PDF" }

        return @{ EstConfirmation = ($confirmation -eq "OUI"); Confiance = $confiance; VerifierCorps = ($source -eq "CORPS"); TexteBrut = $texte; Exclu = $false }
    } catch {
        return @{ EstConfirmation = $false; Confiance = 0.0; VerifierCorps = $false; TexteBrut = "ERREUR: $($_.Exception.Message)"; Exclu = $false }
    }
}

function Invoke-TraiterNouveauCourriel {
    param($Namespace, $MailItem, [bool]$ForcerTraitement = $false)

    $convID = $MailItem.ConversationID
    if (-not $ForcerTraitement) {
        if (Test-ConversationTraitee $convID) { return }
    }

    $adresseExp = $MailItem.SenderEmailAddress
    $seuilAuto  = 0.85

    # Filtre domaine exclu avant d'appeler Haiku
    if (Test-DomaineExclu $adresseExp) {
        Write-Log "INFO  Domaine exclu -- courriel ignore : $adresseExp"
        return
    }

    # Claude Haiku analyse le courriel avec PJ incluses
    $analyse = Invoke-ClassifierCourriel $MailItem

    # Prefixe indicateur dans le sujet du courriel (visible dans la liste Outlook)
    # [OK 94%] = confiant, confirmation | [X 97%] = confiant, pas une confirmation | [? 71%] = incertain
    $pctAffiche = [math]::Round($analyse.Confiance * 100)
    $prefixe = if (-not $analyse.EstConfirmation) { "[X $pctAffiche%]" } `
               elseif ($analyse.Confiance -ge $seuilAuto) { "[OK $pctAffiche%]" } `
               else { "[? $pctAffiche%]" }
    if (-not $MailItem.Subject.StartsWith("[")) {
        try { $MailItem.Subject = "$prefixe $($MailItem.Subject)"; $MailItem.Save() } catch {}
    }

    $estConfirmation = $false
    $verifierCorps   = $false

    if ($ForcerTraitement) {
        # Mode reclassification manuelle : ignorer le verdict de Claude,
        # toujours traiter comme confirmation (Dan a dit que c'en est une)
        Write-Log "INFO  Mode -Force : traitement force comme confirmation (Claude: $(if($analyse.EstConfirmation){'OUI'}else{'NON'}) $([math]::Round($analyse.Confiance*100))%)"
        $estConfirmation = $true
        $verifierCorps   = $analyse.VerifierCorps
        $statutCorpsConnu = Get-StatutCorpsConnu $adresseExp
        if ($statutCorpsConnu -eq "CORPS") { $verifierCorps = $true }
        if ($statutCorpsConnu -eq "PDF")   { $verifierCorps = $false }

    } elseif ($analyse.Confiance -ge $seuilAuto) {
        # Claude est confiant -- agir automatiquement sans deranger Dan
        $estConfirmation = $analyse.EstConfirmation
        $verifierCorps   = $analyse.VerifierCorps

        if (-not $estConfirmation) { return }  # Skip silencieux

        # Pour le mode corps/PDF, privilegier l'apprentissage existant si disponible
        $statutCorpsConnu = Get-StatutCorpsConnu $adresseExp
        if ($statutCorpsConnu -eq "CORPS") { $verifierCorps = $true }
        if ($statutCorpsConnu -eq "PDF")   { $verifierCorps = $false }

    } else {
        # Claude est incertain
        $suggestionOui = $analyse.EstConfirmation
        $pctConfiance  = [math]::Round($analyse.Confiance * 100)

        if ($ForcerTraitement) {
            # Ne devrait plus arriver (gere en haut) mais securite
            $estConfirmation = $true
            $verifierCorps   = $analyse.VerifierCorps
        } else {
            # Mode normal : poser la question a Dan
            $q = "Courriel de: $($MailItem.SenderName)`nSujet: $($MailItem.Subject)`n`nClaude pense que c'est $(if($suggestionOui){'UNE confirmation de commande'}else{'PAS une confirmation de commande'}) (confiance: $pctConfiance%).`n`nEst-ce bien une confirmation de commande?`n`n[Oui] = confirmation, ecarts dans le PDF joint`n[Oui, mais verifier le corps] = confirmation, ecarts ecrits dans le texte du courriel (PDF = juste reference)`n[Non] = pas une confirmation de commande"
            $rep = [System.Windows.Forms.MessageBox]::Show($q, "Confirmation de commande?", "YesNoCancel", "Question")

            if ($rep -eq "Cancel") {
                $estConfirmation = $true; $verifierCorps = $true
                if (-not $suggestionOui) { Set-ReponseFournisseur $adresseExp $true }
                Set-ReponseCorps $adresseExp $true
            } elseif ($rep -eq "Yes") {
                $estConfirmation = $true; $verifierCorps = $false
                if (-not $suggestionOui) { Set-ReponseFournisseur $adresseExp $true }
                Set-ReponseCorps $adresseExp $false
            } else {
                $estConfirmation = $false
                if ($suggestionOui) { Set-ReponseFournisseur $adresseExp $false }
            }
        }
    }

    if (-not $estConfirmation) { return }
    if (-not $ForcerTraitement) { Set-ConversationTraitee $convID }

    Invoke-TraiterComparaison $Namespace $MailItem -VerifierCorps $verifierCorps
}

# =====================================================================
# SYNCHRONISATION SAP (DTW)
# =====================================================================
# Synchronise les commandes "conformes" (statut=OK) vers SAP B1 en ecrivant
# U_NWR_ConfirmPO = "Y" sur le PO via DTW (Data Transfer Workbench), en
# ligne de commande. Appelee a la toute fin du programme principal.
#
# Traitement un PO a la fois (securitaire, facile a deboguer). Le fichier
# source DOIT etre en UTF-16 (Unicode) avec tabulation comme separateur,
# format attendu par le scenario ConfirmPO_PROD.xml.

$Script:DTW_Exe           = "C:\Program Files\sap\Data Transfer Workbench\DTW.exe"
$Script:DTW_ScenarioXml   = "U:\GromecOutlook\DTW\ConfirmPO_PROD.xml"
$Script:DTW_FichierSource = "U:\GromecOutlook\DTW\template.txt"
$Script:DTW_FirebaseUrl   = "https://gromec-outlook-vba-default-rtdb.firebaseio.com"

function Write-FichierSourceDTW {
    <#
    Ecrit le fichier source attendu par DTW pour UN SEUL PO.
    Format : 2 lignes d'en-tete identiques (DocNum / DocEntry / U_NWR_ConfirmPO),
    puis 1 ligne de donnees. Tabulation comme separateur, encodage Unicode (UTF-16 LE).

    Inclut des tentatives repetees en cas de verrou de fichier temporaire (DTW peut
    garder le fichier source "ouvert" brievement apres sa fermeture apparente).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DocNum,
        [int]$NombreTentatives = 8,
        [int]$DelaiSecondes = 3
    )

    $Tab = "`t"
    $Lignes = @(
        "DocNum${Tab}DocEntry${Tab}U_NWR_ConfirmPO",
        "DocNum${Tab}DocEntry${Tab}U_NWR_ConfirmPO",
        "${DocNum}${Tab}${Tab}Y"
    )
    $Contenu = ($Lignes -join "`r`n") + "`r`n"

    $DerniereErreur = $null
    for ($i = 1; $i -le $NombreTentatives; $i++) {
        try {
            [System.IO.File]::WriteAllText($Script:DTW_FichierSource, $Contenu, [System.Text.Encoding]::Unicode)
            return
        } catch {
            $DerniereErreur = $_
            Start-Sleep -Seconds $DelaiSecondes
        }
    }

    throw "Impossible d'ecrire le fichier source apres $NombreTentatives tentatives : $($DerniereErreur.Exception.Message)"
}

function Invoke-DTWImport {
    <#
    Lance DTW.exe en ligne de commande avec le scenario de production.
    Retourne $true si le code de sortie indique un succes, $false sinon.

    Verifie d'abord qu'aucune instance de DTW n'est deja ouverte manuellement
    (ex: Dan en train de l'utiliser) pour eviter tout conflit.
    #>
    $DejaOuvert = Get-Process -Name "DTW" -ErrorAction SilentlyContinue
    if ($DejaOuvert) {
        return @{ Succes = $false; Erreur = "DTW est deja ouvert manuellement (probablement en cours d'utilisation) -- synchronisation reportee au prochain passage." }
    }

    try {
        $Process = Start-Process -FilePath $Script:DTW_Exe `
            -ArgumentList "-s `"$Script:DTW_ScenarioXml`"" `
            -Wait -PassThru -WindowStyle Hidden

        if ($Process.ExitCode -eq 0) {
            Start-Sleep -Seconds 5
            return @{ Succes = $true; Erreur = $null }
        } else {
            Start-Sleep -Seconds 5
            return @{ Succes = $false; Erreur = "DTW a retourne le code de sortie $($Process.ExitCode)" }
        }
    } catch {
        return @{ Succes = $false; Erreur = $_.Exception.Message }
    }
}

function Get-CommandesAConfirmerSAP {
    <#
    Interroge Firebase et retourne les entrees historique avec statut="OK"
    et syncSAP non encore a true.
    #>
    $NoeudHistorique = "$Script:DTW_FirebaseUrl/gromec_vba/historique.json"

    try {
        $Historique = Invoke-RestMethod -Uri $NoeudHistorique -Method Get
    } catch {
        Write-Warning "Invoke-SyncDTW : echec de la lecture de l'historique Firebase : $($_.Exception.Message)"
        return @()
    }

    if ($null -eq $Historique) { return @() }

    $Resultats = @()
    foreach ($Cle in $Historique.PSObject.Properties.Name) {
        $Entree = $Historique.$Cle

        $DejaSync = $false
        if ($Entree.PSObject.Properties.Name -contains 'syncSAP') {
            $DejaSync = [bool]$Entree.syncSAP
        }

        if ($Entree.statut -eq "OK" -and -not $DejaSync) {
            $Resultats += [PSCustomObject]@{
                Cle             = $Cle
                NumeroCommande  = $Entree.numeroCommande
            }
        }
    }

    return $Resultats
}

function Set-StatutSyncSAP {
    <#
    Met a jour le statut de synchronisation SAP pour une entree Firebase donnee.
    #>
    param(
        [Parameter(Mandatory)] [string]$Cle,
        [Parameter(Mandatory)] [bool]$Succes,
        [string]$MessageErreur = $null
    )

    $Url = "$Script:DTW_FirebaseUrl/gromec_vba/historique/$Cle.json"
    $Maintenant = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")

    $Body = @{
        syncSAP     = $Succes
        syncSAPDate = $Maintenant
    }
    if (-not $Succes -and $MessageErreur) {
        $Body["syncSAPErreur"] = $MessageErreur
    }

    try {
        Invoke-RestMethod -Uri $Url -Method Patch -Body ($Body | ConvertTo-Json) -ContentType "application/json" | Out-Null
    } catch {
        Write-Warning "Invoke-SyncDTW : echec de la mise a jour du statut syncSAP pour '$Cle' : $($_.Exception.Message)"
    }
}

function Invoke-SyncDTW {
    <#
    Fonction principale appelee a la toute fin du programme principal.
    #>
    $ACofirmer = Get-CommandesAConfirmerSAP

    if ($ACofirmer.Count -eq 0) {
        Write-Log "Invoke-SyncDTW : aucune commande conforme en attente de synchronisation SAP."
        return
    }

    Write-Log "Invoke-SyncDTW : $($ACofirmer.Count) commande(s) a synchroniser vers SAP."

    foreach ($Item in $ACofirmer) {
        Write-Log "  -> PO $($Item.NumeroCommande) (cle Firebase: $($Item.Cle))..."

        try {
            Write-FichierSourceDTW -DocNum $Item.NumeroCommande
        } catch {
            Write-Warning "     Echec de la generation du fichier source : $($_.Exception.Message)"
            Set-StatutSyncSAP -Cle $Item.Cle -Succes $false -MessageErreur "Echec generation fichier source : $($_.Exception.Message)"
            continue
        }

        $Resultat = Invoke-DTWImport

        if ($Resultat.Succes) {
            Write-Log "     OK - synchronise avec succes."
            Set-StatutSyncSAP -Cle $Item.Cle -Succes $true
        } elseif ($Resultat.Erreur -like "*deja ouvert manuellement*") {
            Write-Log "     DTW est occupe (ouvert manuellement) - synchronisation reportee, arret du traitement pour ce passage."
            return
        } else {
            Write-Warning "     ECHEC - $($Resultat.Erreur)"
            Set-StatutSyncSAP -Cle $Item.Cle -Succes $false -MessageErreur $Resultat.Erreur
        }
    }

    Write-Log "Invoke-SyncDTW : termine."
}

# =====================================================================
# PROGRAMME PRINCIPAL
# =====================================================================

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    # Synchronise vers Outlook les cases "resolu" cochees/decochees depuis le
    # dashboard web depuis le dernier passage du script -- piggyback sur
    # chaque execution, qu'il s'agisse d'un nouveau courriel ou d'un test manuel
    Sync-ResolusVersOutlook $namespace

    # Relance les comparaisons "non appariees" pour lesquelles un numero de BC
    # a ete fourni manuellement depuis le dashboard
    Sync-ReessaisManuels $namespace

    $mail = $null

    if ($Interactive -or ($EntryID -eq "" -and -not $Force)) {
        # Mode interactif: choisir un courriel dans la Boite de reception
        $inbox = $namespace.GetDefaultFolder(6)  # olFolderInbox
        $items = $inbox.Items
        $items.Sort("[ReceivedTime]", $true)

        $liste = @()
        $compte = 0
        foreach ($item in $items) {
            if ($item.Class -ne 43) { continue }
            if ($item.Attachments.Count -eq 0) { continue }
            $liste += [PSCustomObject]@{
                Sujet       = $item.Subject
                Expediteur  = $item.SenderName
                Recu        = $item.ReceivedTime
                EntryID     = $item.EntryID
                StoreID     = $item.Parent.StoreID
            }
            $compte++
            if ($compte -ge 50) { break }
        }

        $choix = $liste | Out-GridView -Title "Choisissez un courriel a traiter (PJ requise)" -OutputMode Single
        if ($null -eq $choix) { exit 0 }

        $mail = $namespace.GetItemFromID($choix.EntryID, $choix.StoreID)
    } else {
        # Tentative 1 : GetItemFromID direct avec le StoreID fourni
        $mail = $null
        try { $mail = $namespace.GetItemFromID($EntryID, $StoreID) } catch { $mail = $null }

        # Tentative 2 : si null, iterer sur tous les stores Outlook (robustesse VBA/StoreID mismatch)
        if ($null -eq $mail) {
            Write-Log "INFO  GetItemFromID direct a echoue, tentative sur tous les stores..."
            foreach ($store in $namespace.Stores) {
                try {
                    $mail = $namespace.GetItemFromID($EntryID, $store.StoreID)
                    if ($null -ne $mail) {
                        Write-Log "INFO  Courriel trouve dans le store : $($store.DisplayName)"
                        break
                    }
                } catch { $mail = $null }
            }
        }
    }

    if ($null -eq $mail) {
        Write-Log "ERREUR  Courriel introuvable dans tous les stores Outlook (EntryID invalide ou courriel deplace/supprime)"
        Write-JournalEntry "" "ERREUR" "Courriel introuvable (EntryID invalide)"
        exit 1
    }

    Invoke-TraiterNouveauCourriel $namespace $mail $Force.IsPresent

    # Synchronisation SAP via DTW (confirmation des commandes "conformes").
    # Isolee dans son propre try/catch : un probleme ici ne doit jamais
    # empecher la liberation de l'objet Outlook ni etre confondu avec
    # une erreur de traitement du courriel courant.
    try {
        Invoke-SyncDTW
    } catch {
        Write-JournalEntry "" "ERREUR_SYNC_DTW" $_.Exception.Message
    }

} catch {
    Write-JournalEntry "" "ERREUR_FATALE" $_.Exception.Message
} finally {
    if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
    [System.GC]::Collect()
}
