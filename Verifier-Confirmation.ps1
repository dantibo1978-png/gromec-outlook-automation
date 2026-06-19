# =====================================================================
# Verifier-Confirmation.ps1
# Verification automatique des confirmations de commande fournisseurs
# Lance de maniere asynchrone depuis Outlook VBA (Shell, sans attente)
# -- Outlook ne gele jamais, tout le travail se fait ici, en dehors
# de son fil d'execution.
#
# Utilisation:
#   powershell -ExecutionPolicy Bypass -File Verifier-Confirmation.ps1 -EntryID "..." -StoreID "..."
#   powershell -ExecutionPolicy Bypass -File Verifier-Confirmation.ps1 -Interactive
#   powershell -ExecutionPolicy Bypass -File Verifier-Confirmation.ps1 -EntryID "..." -StoreID "..." -Force
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

$GitHubRawUrl = "https://raw.githubusercontent.com/dantibo1978-png/gromec-outlook-automation/main/Verifier-Confirmation.ps1"

function Update-ScriptSiNecessaire {
    try {
        $remoteContent = Invoke-RestMethod -Uri $GitHubRawUrl -TimeoutSec 10
    } catch {
        return  # Pas de connexion ou GitHub indisponible -- on continue avec la version locale
    }

    if ([string]::IsNullOrWhiteSpace($remoteContent)) { return }

    $scriptPath = $PSCommandPath
    if ([string]::IsNullOrEmpty($scriptPath)) { return }  # Securite si lance autrement qu'en fichier

    try {
        $localContent = Get-Content -Path $scriptPath -Raw
    } catch {
        return
    }

    # Normaliser les fins de ligne avant comparaison (evite les faux positifs)
    $normLocal  = ($localContent  -replace "`r`n", "`n").Trim()
    $normRemote = ($remoteContent -replace "`r`n", "`n").Trim()

    if ($normLocal -eq $normRemote) { return }  # Deja a jour

    try {
        Set-Content -Path $scriptPath -Value $remoteContent -Encoding UTF8 -Force
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


# =====================================================================
# FONCTIONS - Firebase
# =====================================================================

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
        # Etape 1: chercher le BC dans sujet/corps des envoyes
        foreach ($item in $items) {
            if ($item.Class -ne 43) { continue }  # 43 = olMail
            if ($item.SentOn -lt $limiteDate) { break }
            if ($item.Attachments.Count -gt 0) {
                if ($item.Subject -like "*$NumeroBC*" -or $item.Body -like "*$NumeroBC*") {
                    return $item
                }
            }
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

function Invoke-TraiterComparaison {
    param($Namespace, $MailConfirmation)

    $sujet = $MailConfirmation.Subject
    $expediteur = $MailConfirmation.SenderEmailAddress

    $cheminConfirmation = Save-PremierePDF $MailConfirmation
    if ($cheminConfirmation -eq "") {
        Write-JournalEntry $expediteur "PIECE_JOINTE_PDF_MANQUANTE" "PDF confirmation introuvable"
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
    $numeroBC = Get-NumeroBC "$sujet $($MailConfirmation.Body)"
    if ($numeroBC -eq "" -and $bcGromec -ne "") { $numeroBC = $bcGromec }

    $mailEnvoye = Find-CourrielEnvoyeCorrespondant $Namespace $MailConfirmation $numeroBC
    if ($null -eq $mailEnvoye) {
        Write-JournalEntry $expediteur "AUCUNE_COMMANDE_TROUVEE" "BC recherche: $numeroBC"
        Remove-Item $cheminConfirmation -Force -ErrorAction SilentlyContinue
        return
    }

    $cheminCommande = Save-PremierePDF $mailEnvoye
    if ($cheminCommande -eq "") {
        Write-JournalEntry $expediteur "PIECE_JOINTE_PDF_MANQUANTE" "PDF commande Gromec introuvable"
        Remove-Item $cheminConfirmation -Force -ErrorAction SilentlyContinue
        return
    }

    $resSAP = Get-ItemsCommandeGromec $cheminCommande
    $itemsSAP = @($resSAP.Items)

    Remove-Item $cheminConfirmation -Force -ErrorAction SilentlyContinue
    Remove-Item $cheminCommande -Force -ErrorAction SilentlyContinue

    if ($itemsFourn.Count -eq 0 -or $itemsSAP.Count -eq 0) {
        Write-JournalEntry $expediteur "ARTICLES_NON_EXTRAITS" "Fourn:$($itemsFourn.Count) SAP:$($itemsSAP.Count)"
        return
    }

    # --- Matching ---
    $resultats = Find-TousLesMatches $itemsSAP $itemsFourn

    $nbEcarts = ($resultats | Where-Object { $_.Statut -eq "ECART" }).Count
    $nbNonTrouves = ($resultats | Where-Object { $_.Statut -eq "NON_TROUVE" }).Count
    $estOK = ($nbEcarts -eq 0 -and $nbNonTrouves -eq 0)

    Set-CategorieConfirmation $MailConfirmation $estOK
    Write-JournalEntry $expediteur $(if ($estOK) { "OK" } else { "ECART" }) "Ecarts:$nbEcarts NonTrouves:$nbNonTrouves"
    Write-RapportExcel $nomFourn $sujet $(if ($estOK) { "OK" } else { "ECART" }) $resultats $devise $numeroBC
}

# =====================================================================
# FONCTION - Traitement d'un nouveau courriel (classification + apprentissage)
# =====================================================================

function Invoke-TraiterNouveauCourriel {
    param($Namespace, $MailItem, [bool]$ForcerTraitement = $false)

    $convID = $MailItem.ConversationID
    if (-not $ForcerTraitement) {
        if (Test-ConversationTraitee $convID) { return }
    }

    $adresseExp = $MailItem.SenderEmailAddress
    $statutConnu = Get-StatutFournisseurConnu $adresseExp
    $estConfirmation = $false

    switch ($statutConnu) {
        "OUI" { $estConfirmation = $true }
        "NON" { $estConfirmation = $false }
        default {
            $sysPrompt = "Tu analyses des courriels pour determiner si c'est une CONFIRMATION DE COMMANDE FOURNISSEUR (document formel avec numero de commande, articles, quantites, prix ou delai). Reponds UNIQUEMENT: REPONSE: OUI  ou  REPONSE: NON"
            $usrPrompt = "Expediteur: $($MailItem.SenderName)`nSujet: $($MailItem.Subject)`nCorps:`n$($MailItem.Body.Substring(0,[Math]::Min(2000,$MailItem.Body.Length)))"
            $suggestion = Invoke-ClaudeMessage $sysPrompt $usrPrompt
            $suggestionOui = $suggestion -match "OUI"

            $q = "Courriel de: $($MailItem.SenderName)`nSujet: $($MailItem.Subject)`n`nClaude pense que c'est $(if($suggestionOui){'UNE confirmation de commande.'}else{'PAS une confirmation de commande.'})`n`nEst-ce bien une confirmation de commande?"
            $rep = [System.Windows.Forms.MessageBox]::Show($q, "Confirmation de commande?", "YesNo", "Question")
            $estConfirmation = ($rep -eq "Yes")
            Set-ReponseFournisseur $adresseExp $estConfirmation
        }
    }

    if (-not $estConfirmation) { return }
    if (-not $ForcerTraitement) { Set-ConversationTraitee $convID }

    Invoke-TraiterComparaison $Namespace $MailItem
}

# =====================================================================
# PROGRAMME PRINCIPAL
# =====================================================================

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

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
        $mail = $namespace.GetItemFromID($EntryID, $StoreID)
    }

    if ($null -eq $mail) {
        Write-JournalEntry "" "ERREUR" "Courriel introuvable (EntryID invalide)"
        exit 1
    }

    Invoke-TraiterNouveauCourriel $namespace $mail $Force.IsPresent

} catch {
    Write-JournalEntry "" "ERREUR_FATALE" $_.Exception.Message
} finally {
    if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
    [System.GC]::Collect()
}
