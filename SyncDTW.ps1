# =============================================================================
# SyncDTW.ps1
# Tourne en arriere-plan sur le poste Windows.
# Poll Firebase toutes les 30s, detecte les flags dtw_copierPrix / dtw_copierQty
# poses par le dashboard web, genere les fichiers source DTW et lance DTW.exe.
#
# Lancement (une seule instance, en arriere-plan) :
#   powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "U:\GromecOutlook\SyncDTW.ps1"
#
# Pour arreter : fermer la fenetre PowerShell ou tuer le processus.
# =============================================================================

# ── Auto-update depuis GitHub ─────────────────────────────────────────────────
$GitHubRawUrl = "https://raw.githubusercontent.com/dantibo1978-png/gromec-outlook-automation/main/SyncDTW.ps1"

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

    Start-Process -FilePath "powershell.exe" -ArgumentList @(
        "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $scriptPath
    ) -WindowStyle Hidden
    exit 0
}

Update-ScriptSiNecessaire

# ── Configuration ─────────────────────────────────────────────────────────────
$FirebaseUrl        = "https://gromec-outlook-vba-default-rtdb.firebaseio.com"
$DTW_Exe            = "C:\Program Files\sap\Data Transfer Workbench\DTW.exe"
$DTW_Dossier        = "U:\GromecOutlook\DTW"
$DTW_FichierLignes     = "$DTW_Dossier\modif_prix_dans_po.txt"
$DTW_FichierPrixL2     = "$DTW_Dossier\modif_prix_fourn.txt"
$DTW_FichierConfirmPO  = "$DTW_Dossier\template.txt"
$DTW_ScenarioLignes    = "$DTW_Dossier\UpdatePOLines_PROD.xml"
$DTW_ScenarioPrixL2    = "$DTW_Dossier\UpdatePriceList2_PROD.xml"
$DTW_ScenarioConfirmPO = "$DTW_Dossier\ConfirmPO_PROD.xml"
$DTW_FichierPrix       = "$DTW_Dossier\import_prix.txt"
$DTW_ScenarioPrix      = "$DTW_Dossier\UpdatePriceList_PROD.xml"
$IntervalleSecondes = 15
$Logo_Gromec        = "U:\GromecOutlook\logo_gromec.png"
$Temp_Dossier       = "U:\GromecOutlook\temp"
# ──────────────────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Cle = "", [string]$BC = "")
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$ts] $Message"
    try {
        $niveau = if ($Message -like "ERREUR*") { "erreur" } elseif ($Message -like "WARN*") { "warn" } else { "info" }
        $body = @{ ts = $ts; msg = $Message; niveau = $niveau; source = "SyncDTW" }
        if ($Cle -ne "") { $body["cle"] = $Cle }
        if ($BC -ne "")  { $body["bc"]  = $BC }
        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/logs.json" -Method Post -Body ($body | ConvertTo-Json -Compress) -ContentType "application/json" -TimeoutSec 3 | Out-Null
    } catch {}
}

function Invoke-PurgeVieuxLogs {
    # Supprime les entrees de gromec_vba/logs de plus de 3 jours pour limiter
    # la taille du noeud (et donc la bande passante Firebase consommee par
    # le dashboard qui relit ce noeud regulierement).
    try {
        $cutoff = (Get-Date).AddDays(-3).ToString("yyyy-MM-dd HH:mm:ss")
        $logs = Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/logs.json?orderBy=%22ts%22&endAt=%22$cutoff%22" -Method Get -TimeoutSec 20
        if ($null -eq $logs -or $logs -eq "null") { return }
        $nb = 0
        foreach ($cle in $logs.PSObject.Properties.Name) {
            Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/logs/$cle.json" -Method Delete -TimeoutSec 10 | Out-Null
            $nb++
        }
        if ($nb -gt 0) { Write-Log "INFO  Purge logs : $nb entree(s) de plus de 3 jours supprimee(s)." }
    } catch {}
}

function Get-Historique {
    try {
        $rep = Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/historique.json" -Method Get -TimeoutSec 15
        return $rep
    } catch {
        Write-Log "WARN  Impossible de lire Firebase : $($_.Exception.Message)"
        return $null
    }
}

function Set-StatutDTW {
    param(
        [string]$Cle,
        [string]$Statut,        # 'ok' | 'erreur'
        [string]$Erreur = $null
    )
    $Maintenant = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    $body = @{
        dtw_statut         = $Statut
        dtw_traiteLe       = $Maintenant
        # Remettre les flags a false pour ne pas retraiter au prochain poll
        dtw_copierPrix     = $false
        dtw_copierQty      = $false
    }
    # Si succes DTW, marquer la commande comme conforme dans Firebase
    if ($Statut -eq 'ok') {
        $body["statut"]  = "OK"
        $body["confirme"] = $true
        $body["confirme_le"] = $Maintenant
    }
    if ($Erreur) { $body["dtw_erreur"] = $Erreur }

    try {
        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/historique/$Cle.json" `
            -Method Patch `
            -Body ($body | ConvertTo-Json -Compress) `
            -ContentType "application/json" `
            -TimeoutSec 15 | Out-Null
    } catch {
        Write-Log "WARN  Impossible de mettre a jour Firebase pour '$Cle' : $($_.Exception.Message)"
    }
}

function Write-FichierLignesDTW {
    <#
    Genere le fichier source pour DTW - Document_Lines (UnitPrice et/ou Quantity).
    Format : 2 lignes d'en-tete identiques, puis 1 ligne par article.
    Encodage : UTF-16 LE (Unicode), separateur tabulation.

    Colonnes : ParentKey (DocNum), LineNum (0-indexed), ItemCode, UnitPrice, Quantity
    On inclut toujours les 2 colonnes meme si on n'en modifie qu'une -
    DTW ne modifie que les champs mappes dans le scenario XML.
    #>
    param(
        [string]$DocNum,
        [array]$Articles,       # objets avec sapLigne, sapArticle, pdfPrix, pdfQty, sapPrix, sapQty
        [bool]$CopierPrix,
        [bool]$CopierQty
    )

    $Tab = "`t"
    $Header = "ParentKey${Tab}LineNum${Tab}ItemCode${Tab}ItemDescription${Tab}Quantity${Tab}ShipDate${Tab}UnitPrice"
    $Lignes = @($Header, $Header)

    foreach ($a in $Articles) {
        $lineNum = [int]$a.sapLigne - 1   # 0-indexed

        $prix = if ($CopierPrix) {
            $v = if ($a.pdfPrix -and $a.pdfPrix -gt 0) { $a.pdfPrix } else { $a.sapPrix }
            [string][math]::Round([double]$v, 4)
        } else { "" }

        $qty = if ($CopierQty) {
            $v = if ($a.pdfQty -and $a.pdfQty -gt 0) { $a.pdfQty } else { $a.sapQty }
            [string][int]$v
        } else { "" }

        $Lignes += "${DocNum}${Tab}${lineNum}${Tab}${Tab}${Tab}${qty}${Tab}${Tab}${prix}"
    }

    $Contenu = ($Lignes -join "`r`n") + "`r`n"

    for ($i = 1; $i -le 8; $i++) {
        try {
            [System.IO.File]::WriteAllText($DTW_FichierLignes, $Contenu, [System.Text.Encoding]::Unicode)
            return
        } catch {
            Start-Sleep -Seconds 3
        }
    }
    throw "Impossible d'ecrire $DTW_FichierLignes apres 8 tentatives"
}

function Write-FichierPrixL2DTW {
    <#
    Genere le fichier source pour DTW - Price List 2 (Items/ItemPrices).
    Format : 2 lignes d'en-tete identiques, puis 1 ligne par article.
    Colonnes : ItemCode (A), PriceList=1 (B), ListNum=2 (C), Price (D), Currency=vide (E)
    #>
    param([array]$Articles)

    $Tab = "`t"
    $Header = "ParentKey${Tab}LineNum${Tab}PriceList${Tab}Price"
    $Lignes = @($Header, $Header)

    foreach ($a in $Articles) {
        $itemCode = ($a.sapArticle -replace '^0+', '')
        if ($itemCode -eq '') { $itemCode = $a.sapArticle }

        $prix = if ($a.pdfPrix -and $a.pdfPrix -gt 0) { $a.pdfPrix } else { $a.sapPrix }
        $prixStr = [string][math]::Round([double]$prix, 4)

        $Lignes += "${itemCode}${Tab}1${Tab}2${Tab}${prixStr}"
    }

    $Contenu = ($Lignes -join "`r`n") + "`r`n"

    for ($i = 1; $i -le 8; $i++) {
        try {
            [System.IO.File]::WriteAllText($DTW_FichierPrixL2, $Contenu, [System.Text.Encoding]::Unicode)
            return
        } catch {
            Start-Sleep -Seconds 3
        }
    }
    throw "Impossible d'ecrire $DTW_FichierPrixL2 apres 8 tentatives"
}

function Invoke-DTW {
    param([string]$ScenarioXml)

    # Attendre que tout processus DTW residuel soit ferme (jusqu'a 15s)
    $attenteInit = 0
    while ((Get-Process -Name "DTW" -ErrorAction SilentlyContinue) -and $attenteInit -lt 15) {
        Start-Sleep -Seconds 2
        $attenteInit += 2
    }

    # Verifier qu'aucune instance manuelle de DTW n'est ouverte
    $dejaOuvert = Get-Process -Name "DTW" -ErrorAction SilentlyContinue
    if ($dejaOuvert) {
        return @{ Succes = $false; Erreur = "DTW est deja ouvert manuellement - operation reportee." }
    }

    try {
        $proc = Start-Process -FilePath $DTW_Exe `
            -ArgumentList "-s `"$ScenarioXml`"" `
            -Wait -PassThru -WindowStyle Hidden

        # Attendre que DTW soit completement ferme (jusqu'a 30s)
        $attente = 0
        while ((Get-Process -Name "DTW" -ErrorAction SilentlyContinue) -and $attente -lt 30) {
            Start-Sleep -Seconds 2
            $attente += 2
        }
        Start-Sleep -Seconds 3

        if ($proc.ExitCode -eq 0) {
            return @{ Succes = $true; Erreur = $null }
        } else {
            return @{ Succes = $false; Erreur = "DTW code de sortie $($proc.ExitCode)" }
        }
    } catch {
        return @{ Succes = $false; Erreur = $_.Exception.Message }
    }
}

function Invoke-TraiterEntree {
    param(
        [string]$Cle,
        [object]$Entree
    )

    # Relire l'entree fraiche depuis Firebase pour avoir dtw_lignesCochees a jour
    try {
        $EntreeFraiche = Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/historique/$Cle.json" -Method Get -TimeoutSec 10
        if ($null -ne $EntreeFraiche) { $Entree = $EntreeFraiche }
    } catch {
        Write-Log "WARN  Impossible de relire l'entree fraiche, utilisation du cache."
    }

    # Verifier que l'entree est encore en attente (evite double-traitement)
    if ($Entree.dtw_statut -ne 'en_attente') {
        Write-Log "INFO  Entree $Cle deja traitee (statut=$($Entree.dtw_statut)) - abandon."
        return
    }

    # Marquer immediatement comme en cours pour eviter double-traitement
    try {
        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/historique/$Cle.json" `
            -Method Patch `
            -Body '{"dtw_statut":"en_cours"}' `
            -ContentType "application/json" `
            -TimeoutSec 10 | Out-Null
    } catch {}

    $copierPrix     = [bool]$Entree.dtw_copierPrix
    $copierQty      = [bool]$Entree.dtw_copierQty
    $docNum         = $Entree.numeroCommande
    $articlesTotal  = @($Entree.articles)

    # Filtrer selon les lignes cochees (si dtw_lignesCochees est fourni par le dashboard)
    $lignesCochees = @()
    if ($null -ne $Entree.dtw_lignesCochees) {
        try {
            $lignesCochees = @($Entree.dtw_lignesCochees | ForEach-Object { [int]"$_" })
        } catch {
            $lignesCochees = @()
        }
    }

    # Aussi prendre en compte copierPrix/copierQty depuis Firebase
    # (pour compatibilite avec l'ancien comportement sans cases a cocher)
    if ($lignesCochees.Count -gt 0) {
        # Dashboard a envoye des lignes cochees specifiques
        $articles = @($articlesTotal | Where-Object { $lignesCochees -contains [int]"$($_.sapLigne)" })
        Write-Log "INFO  Lignes filtrees selon selection ($($lignesCochees.Count) cochees -> $($articles.Count) articles)"
    } else {
        # Pas de selection specifique -> prendre seulement les ecarts
        $articles = @($articlesTotal | Where-Object { $_.statut -eq 'ECART' })
        Write-Log "INFO  Aucune selection specifique -> ecarts seulement ($($articles.Count) articles)"
    }
    # Si aucun flag copier defini, activer les deux par defaut
    if (-not $copierPrix -and -not $copierQty) {
        $copierPrix = $true
        $copierQty  = $true
    }

    Write-Log "INFO  Traitement PO $docNum (cle: $Cle) - prix=$copierPrix qty=$copierQty articles=$($articles.Count)" -Cle $Cle -BC $docNum
    if ($articles.Count -gt 0) {
        Write-Log "INFO  Premier article: sapLigne=$($articles[0].sapLigne) type=$($articles[0].sapLigne.GetType().Name)" -Cle $Cle -BC $docNum
    }

    if ($articles.Count -eq 0) {
        Write-Log "WARN  Aucun article dans l'entree Firebase - abandon." -Cle $Cle -BC $docNum
        Set-StatutDTW -Cle $Cle -Statut 'erreur' -Erreur "Aucun article dans l'entree Firebase."
        return
    }

    # ── Etape 1 : Document_Lines (UnitPrice et/ou Quantity) ─────────────────
    try {
        Write-FichierLignesDTW -DocNum $docNum -Articles $articles -CopierPrix $copierPrix -CopierQty $copierQty
    } catch {
        Set-StatutDTW -Cle $Cle -Statut 'erreur' -Erreur "Erreur generation fichier lignes : $($_.Exception.Message)"
        return
    }

    $res = Invoke-DTW -ScenarioXml $DTW_ScenarioLignes
    if (-not $res.Succes) {
        Write-Log "ERREUR DTW lignes : $($res.Erreur)" -Cle $Cle -BC $docNum
        Set-StatutDTW -Cle $Cle -Statut 'erreur' -Erreur "DTW lignes : $($res.Erreur)"
        return
    }
    Write-Log "INFO  DTW lignes OK." -Cle $Cle -BC $docNum

    # ── Etape 2 : Price List 2 (seulement si on copie les prix) ─────────────
    if ($copierPrix) {
        try {
            Write-FichierPrixL2DTW -Articles $articles
        } catch {
            Write-Log "WARN  Erreur generation fichier Price List 2 : $($_.Exception.Message)" -Cle $Cle -BC $docNum
            Set-StatutDTW -Cle $Cle -Statut 'erreur' -Erreur "DTW lignes OK, mais erreur Price List 2 : $($_.Exception.Message)"
            return
        }

        $res2 = Invoke-DTW -ScenarioXml $DTW_ScenarioPrixL2
        if (-not $res2.Succes) {
            Write-Log "WARN  DTW Price List 2 : $($res2.Erreur)" -Cle $Cle -BC $docNum
            Set-StatutDTW -Cle $Cle -Statut 'erreur' -Erreur "DTW lignes OK, mais DTW Price List 2 : $($res2.Erreur)"
            return
        }
        Write-Log "INFO  DTW Price List 2 OK." -Cle $Cle -BC $docNum
    }

    # ── Etape 3 : U_NWR_ConfirmPO = Y (confirmer la commande dans SAP) ────────
    try {
        $contenuConfirm = "DocNum`tDocEntry`tU_NWR_ConfirmPO`r`nDocNum`tDocEntry`tU_NWR_ConfirmPO`r`n$docNum`t`tY"
        [System.IO.File]::WriteAllText($DTW_FichierConfirmPO, $contenuConfirm, [System.Text.Encoding]::Unicode)
    } catch {
        Write-Log "WARN  Erreur generation fichier ConfirmPO : $($_.Exception.Message)" -Cle $Cle -BC $docNum
        Set-StatutDTW -Cle $Cle -Statut 'ok'
        Set-CategorieOutlook -EntryID $Entree.entryID -StoreID $Entree.storeID -EstOK $true
        Write-Log "INFO  PO $docNum traite avec succes (ConfirmPO non mis a jour)." -Cle $Cle -BC $docNum
        return
    }

    $res3 = Invoke-DTW -ScenarioXml $DTW_ScenarioConfirmPO
    if (-not $res3.Succes) {
        Write-Log "WARN  DTW ConfirmPO : $($res3.Erreur)" -Cle $Cle -BC $docNum
        Set-StatutDTW -Cle $Cle -Statut 'ok'
        Set-CategorieOutlook -EntryID $Entree.entryID -StoreID $Entree.storeID -EstOK $true
        Write-Log "INFO  PO $docNum traite avec succes (ConfirmPO non mis a jour)." -Cle $Cle -BC $docNum
        return
    }
    Write-Log "INFO  DTW ConfirmPO OK." -Cle $Cle -BC $docNum

    Set-StatutDTW -Cle $Cle -Statut 'ok'
    Set-CategorieOutlook -EntryID $Entree.entryID -StoreID $Entree.storeID -EstOK $true
    Write-Log "INFO  PO $docNum traite avec succes." -Cle $Cle -BC $docNum
}


function Set-CategorieOutlook {
    param([string]$EntryID, [string]$StoreID, [bool]$EstOK)
    if (-not $EntryID -or -not $StoreID) { return }
    $outlook = $null
    try {
        $outlook   = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $mail      = $namespace.GetItemFromID($EntryID, $StoreID)
        if ($null -eq $mail) { return }

        $catOK    = "Confirmation OK"
        $catEcart = "Confirmation - Ecart"
        $cats = $mail.Categories
        if ([string]::IsNullOrEmpty($cats)) { $cats = "" }
        $cats = $cats -replace [regex]::Escape($catOK), "" -replace [regex]::Escape($catEcart), ""
        $cats = ($cats -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }) -join ", "
        $nouvelleCat = if ($EstOK) { $catOK } else { $catEcart }
        $mail.Categories = if ($cats -ne "") { "$cats, $nouvelleCat" } else { $nouvelleCat }
        $mail.Save()
        Write-Log "INFO  Categorie Outlook mise a jour : $nouvelleCat"
    } catch {
        Write-Log "WARN  Impossible de mettre a jour la categorie Outlook : $($_.Exception.Message)"
    } finally {
        if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
    }
}

function Set-StatutConfirmation {
    param([string]$Cle, [string]$Statut, [string]$Erreur = $null)
    $body = @{ envoyer_confirmation = $Statut; envoyer_confirmation_le = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss") }
    if ($Erreur) { $body["envoyer_confirmation_erreur"] = $Erreur }
    try {
        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/historique/$Cle.json" `
            -Method Patch `
            -Body ($body | ConvertTo-Json -Compress) `
            -ContentType "application/json" `
            -TimeoutSec 15 | Out-Null
    } catch {
        Write-Log "WARN  Impossible de mettre a jour statut confirmation pour '$Cle' : $($_.Exception.Message)"
    }
}

function Get-PDFOriginalDepuisOutlook {
    <#
    Cherche la premiere piece jointe PDF dans le courriel original (Sent Items)
    identifie par EntryID/StoreID, et la sauvegarde dans CheminDestination.
    Retourne $true si trouve, $false sinon.
    #>
    param([string]$EntryID, [string]$StoreID, [string]$CheminDestination)

    $outlook = $null; $namespace = $null; $mail = $null
    try {
        $outlook   = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $mail      = $namespace.GetItemFromID($EntryID, $StoreID)

        foreach ($pj in $mail.Attachments) {
            if ($pj.FileName -match '\.pdf$') {
                $pj.SaveAsFile($CheminDestination)
                return $true
            }
        }
        return $false
    } catch {
        Write-Log "WARN  Impossible de recuperer le PDF original : $($_.Exception.Message)"
        return $false
    } finally {
        if ($mail)      { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail)      | Out-Null } catch {} }
        if ($namespace) { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null } catch {} }
        if ($outlook)   { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)   | Out-Null } catch {} }
        [System.GC]::Collect()
    }
}

function New-PDFAnnote {
    <#
    Prend le PDF SAP original et annote les lignes modifiees en rouge.
    Pour chaque article avec ecart de prix ou de quantite, ajoute un tampon
    rouge dans la marge droite avec les nouvelles valeurs.
    Utilise iTextSharp 5.x.
    #>
    param([string]$CheminPDFSource, [string]$CheminPDFDest, [array]$Articles)

    $iTextDLL      = "U:\GromecOutlook\lib\itextsharp\lib\itextsharp.dll"
    $bouncyCastleDLL = "U:\GromecOutlook\lib\itextsharp\lib\BouncyCastle.Crypto.dll"
    Add-Type -Path $bouncyCastleDLL
    Add-Type -Path $iTextDLL

    $reader  = New-Object iTextSharp.text.pdf.PdfReader($CheminPDFSource)
    $stamper = New-Object iTextSharp.text.pdf.PdfStamper($reader, [System.IO.File]::Create($CheminPDFDest))

    try {
        $pageWidth  = $reader.GetPageSize(1).Width
        $pageHeight = $reader.GetPageSize(1).Height

        # Police et couleurs
        $fontRouge  = [iTextSharp.text.pdf.BaseFont]::CreateFont([iTextSharp.text.pdf.BaseFont]::HELVETICA_BOLD, [iTextSharp.text.pdf.BaseFont]::CP1252, $false)
        $rouge      = [iTextSharp.text.BaseColor]::RED
        $blanc      = [iTextSharp.text.BaseColor]::WHITE
        $jauneLight = New-Object iTextSharp.text.BaseColor(255, 255, 180)

        # Position Y de depart du tableau articles (approximation depuis le bas de page)
        # Le PDF SAP place les lignes articles a environ 55-75% de la hauteur de page
        # On positionne les annotations en fonction du numero de ligne SAP
        $yBase      = $pageHeight * 0.62   # Y du haut du tableau articles (approx)
        $hautLigne  = 18.5                  # hauteur approximative d'une ligne en points

        $cb = $stamper.GetOverContent(1)

        # Banniere REVISION en haut du document
        $cb.SetColorFill($jauneLight)
        $cb.Rectangle(30, $pageHeight - 55, $pageWidth - 60, 18)
        $cb.Fill()
        $cb.SetColorFill($rouge)
        $cb.BeginText()
        $cb.SetFontAndSize($fontRouge, 9)
        $cb.ShowTextAligned([iTextSharp.text.Element]::ALIGN_CENTER, "** BON DE COMMANDE REVISE -- PRIX ET/OU QUANTITES MIS A JOUR **", $pageWidth / 2, $pageHeight - 46, 0)
        $cb.EndText()

        # Annotations par ligne modifiee
        $ligneNum = 0
        foreach ($a in $Articles) {
            $ligneNum++
            $prixSap = [double]$a.sapPrix
            $prixPdf = if ($a.pdfPrix -and [double]$a.pdfPrix -gt 0) { [double]$a.pdfPrix } else { $prixSap }
            $qtySap  = [int]$a.sapQty
            $qtyPdf  = if ($a.pdfQty  -and [int]$a.pdfQty    -gt 0) { [int]$a.pdfQty    } else { $qtySap  }

            $prixModif = [math]::Abs($prixPdf - $prixSap) -gt 0.001
            $qtyModif  = $qtyPdf -ne $qtySap

            if (-not $prixModif -and -not $qtyModif) { continue }

            $yLigne = $yBase - ($ligneNum - 1) * $hautLigne

            # Rectangle rose translucide sur la ligne
            $cb.SaveState()
            $couleurLigne = New-Object iTextSharp.text.BaseColor(255, 200, 200)
            $cb.SetColorFill($couleurLigne)
            $cb.Rectangle(25, $yLigne - 3, $pageWidth - 50, $hautLigne)
            $cb.Fill()
            $cb.RestoreState()

            # Texte annotation a droite
            $annotation = ""
            if ($prixModif) { $annotation += "Prix: $([math]::Round($prixSap,2))$ -> $([math]::Round($prixPdf,2))$  " }
            if ($qtyModif)  { $annotation += "Qte: $qtySap -> $qtyPdf" }

            $cb.SetColorFill($rouge)
            $cb.BeginText()
            $cb.SetFontAndSize($fontRouge, 7.5)
            $cb.ShowTextAligned([iTextSharp.text.Element]::ALIGN_RIGHT, $annotation.Trim(), $pageWidth - 30, $yLigne + 3, 0)
            $cb.EndText()
        }

    } finally {
        $stamper.Close()
        $reader.Close()
    }
}

function New-BrouillonOutlook {
    param([object]$Entree, [string]$CheminPDF)

    $entryID = $Entree.entryID
    $storeID = $Entree.storeID
    $bc      = $Entree.numeroCommande

    $outlook = $null; $namespace = $null; $mailOrig = $null; $brouillon = $null
    try {
        $outlook   = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $mailOrig  = $namespace.GetItemFromID($entryID, $storeID)

        $brouillon = $mailOrig.ReplyAll()
        $brouillon.Subject  = "Bon de commande revise -- BC $bc"
        $brouillon.HTMLBody = "<html><body><p>Bonjour,</p><p>Veuillez trouver ci-joint notre bon de commande revise (BC $bc).</p><p>##POLITESSE##</p></body></html>"
        if ($CheminPDF -and (Test-Path $CheminPDF)) {
            $brouillon.Attachments.Add($CheminPDF) | Out-Null
        }
        $brouillon.Save()

        Write-Log "INFO  Brouillon cree dans Outlook pour BC $bc."
    } catch {
        throw "Erreur creation brouillon Outlook : $($_.Exception.Message)"
    } finally {
        if ($brouillon)  { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($brouillon)  | Out-Null } catch {} }
        if ($mailOrig)   { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mailOrig)   | Out-Null } catch {} }
        if ($namespace)  { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace)  | Out-Null } catch {} }
        if ($outlook)    { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)    | Out-Null } catch {} }
        [System.GC]::Collect()
    }
}

function Invoke-EnvoyerConfirmationFournisseur {
    param([string]$Cle, [object]$Entree)

    $bc = $Entree.numeroCommande
    Write-Log "INFO  Confirmation fournisseur BC $bc (cle: $Cle)." -Cle $Cle -BC $bc

    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.Windows.Forms  -ErrorAction SilentlyContinue

        # ── Etape 1 : Ouvrir le BC dans SAP via SendKeys ─────────────────────
        $sap = Get-Process | Where-Object { $_.MainWindowTitle -like "*SAP Business One*" } | Select-Object -First 1
        if ($sap) {
            [Microsoft.VisualBasic.Interaction]::AppActivate($sap.Id)
            Start-Sleep -Milliseconds 500
            [System.Windows.Forms.SendKeys]::SendWait("{F9}")
            Start-Sleep -Seconds 4
            [System.Windows.Forms.SendKeys]::SendWait("^f")
            Start-Sleep -Milliseconds 800
            [System.Windows.Forms.SendKeys]::SendWait($bc)
            Start-Sleep -Milliseconds 500
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Start-Sleep -Seconds 2
            Write-Log "INFO  BC $bc ouvert dans SAP."
        } else {
            Write-Log "WARN  SAP Business One non trouve -- etape SAP ignoree."
        }

        # ── Etape 2 : Ouvrir le courriel en Repondre a tous dans Outlook ─────
        $outlook  = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $mailOrig = $namespace.GetItemFromID($Entree.entryID, $Entree.storeID)
        if ($mailOrig) {
            $reponse = $mailOrig.ReplyAll()
            $reponse.Display()
            Write-Log "INFO  Repondre a tous ouvert dans Outlook pour BC $bc."
        } else {
            Write-Log "WARN  Courriel original introuvable dans Outlook."
        }

        Set-StatutConfirmation -Cle $Cle -Statut 'envoye'
        Write-Log "INFO  BC $bc -- confirmation fournisseur preparee avec succes."

    } catch {
        Write-Log "ERREUR confirmation fournisseur BC $bc : $($_.Exception.Message)"
        Set-StatutConfirmation -Cle $Cle -Statut 'erreur' -Erreur $_.Exception.Message
    }
}



function Set-ReponseFournisseur {
    param([string]$Adresse, [bool]$EstConfirmation)
    $fichier = "U:\GromecOutlook\fournisseurs_appris.csv"
    $nbOui = 0; $nbNon = 0
    $lignes = @(); $trouve = $false

    if (Test-Path $fichier) {
        foreach ($ligne in Get-Content $fichier) {
            $champs = $ligne -split ","
            if ($champs.Count -ge 3 -and $champs[0].Trim().ToLower() -eq $Adresse.ToLower()) {
                $nbOui = [int]$champs[1]; $nbNon = [int]$champs[2]
                if ($EstConfirmation) { $nbOui++ } else { $nbNon++ }
                $lignes += "$Adresse,$nbOui,$nbNon"
                $trouve = $true
            } else {
                $lignes += $ligne
            }
        }
    }
    if (-not $trouve) {
        if ($EstConfirmation) { $lignes += "$Adresse,1,0" } else { $lignes += "$Adresse,0,1" }
    }
    $lignes | Out-File -FilePath $fichier -Encoding ASCII -Force
    Write-Log "INFO  fournisseurs_appris.csv mis a jour : $Adresse -> $(if ($EstConfirmation) { 'OUI' } else { 'NON' })"
}

function Invoke-TraiterReclassification {
    param([object]$Reclassif, [string]$CleFirebase)

    $expediteur      = $Reclassif.expediteur
    $estConfirmation = [bool]$Reclassif.estConfirmation
    $classification  = $Reclassif.classification
    $sujet           = $Reclassif.sujet

    Write-Log "INFO  Reclassification manuelle : $expediteur -> $classification"

    try {
        # Mettre a jour fournisseurs_appris.csv
        Set-ReponseFournisseur -Adresse $expediteur -EstConfirmation $estConfirmation

        # Si c est une confirmation, relancer VerifierConfirmation.ps1 avec -Force
        if ($estConfirmation) {
            $entryID = $Reclassif.entryID
            $storeID = $Reclassif.storeID
            $ps1Path = "U:\GromecOutlook\VerifierConfirmation.ps1"

            Write-Log "INFO  Retraitement force : lancement de VerifierConfirmation.ps1 pour $sujet"
            Start-Process -FilePath "powershell.exe" `
                -ArgumentList "-ExecutionPolicy Bypass -File `"$ps1Path`" -EntryID `"$entryID`" -StoreID `"$storeID`" -Force" `
                -WindowStyle Hidden `
                -Wait
            Write-Log "INFO  Retraitement termine pour : $sujet"
        }

        # Effacer cette entree specifique dans Firebase (pas toute la liste)
        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/reclassifications/$CleFirebase.json" `
            -Method Delete `
            -TimeoutSec 15 | Out-Null

        Write-Log "INFO  Reclassification '$CleFirebase' traitee avec succes."

    } catch {
        Write-Log "ERREUR reclassification '$CleFirebase' : $($_.Exception.Message)"
    }
}


# ── Boucle principale ─────────────────────────────────────────────────────────
# Synchroniser domaines_exclus depuis Firebase au demarrage
try {
    $domainesFirebase = Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/domaines_exclus.json" -Method Get -TimeoutSec 10
    if ($null -ne $domainesFirebase -and $domainesFirebase -ne "null") {
        $fichierExclus = "U:\GromecOutlook\domaines_exclus.csv"
        $domaines = @()
        $domainesFirebase.PSObject.Properties | ForEach-Object { $domaines += $_.Value }
        $domaines | Where-Object { $_ -ne "" } | Sort-Object -Unique | Set-Content -Path $fichierExclus -Encoding UTF8
        Write-Log "INFO  domaines_exclus.csv synchronise ($($domaines.Count) domaine(s))."
    }
} catch {
    Write-Log "WARN  Impossible de synchroniser domaines_exclus : $($_.Exception.Message)"
}

Write-Log "INFO  SyncDTW.ps1 demarre. Poll toutes les ${IntervalleSecondes}s."

$DerniereePurgeLogs = Get-Date "2000-01-01"

while ($true) {
    if (((Get-Date) - $DerniereePurgeLogs).TotalHours -ge 1) {
        Invoke-PurgeVieuxLogs
        $DerniereePurgeLogs = Get-Date
    }

    # Verifier reclassifications manuelles (VBA Outlook) -- liste pour supporter plusieurs en meme temps
    try {
        $reclassifs = Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/reclassifications.json" -Method Get -TimeoutSec 10
        if ($null -ne $reclassifs -and $reclassifs -ne "null") {
            foreach ($cleR in $reclassifs.PSObject.Properties.Name) {
                $reclassif = $reclassifs.$cleR
                if ($null -ne $reclassif -and $reclassif.expediteur) {
                    Invoke-TraiterReclassification -Reclassif $reclassif -CleFirebase $cleR
                }
            }
        }
    } catch {
        Write-Log "WARN  Erreur lecture noeud reclassifications : $($_.Exception.Message)"
    }

    # Verifier les entrees a reessayer (NON_TROUVE avec numeroBCManuel fourni)
    try {
        $historiquePlat = Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/historique.json" -Method Get -TimeoutSec 10
        if ($null -ne $historiquePlat -and $historiquePlat -ne "null") {
            foreach ($cleR in $historiquePlat.PSObject.Properties.Name) {
                $entreeR = $historiquePlat.$cleR
                if ($entreeR.aReessayer -eq $true -and $entreeR.numeroBCManuel -and $entreeR.entryID) {
                    Write-Log "INFO  Reessai demande pour BC $($entreeR.numeroBCManuel) (cle: $cleR)"
                    # Remettre aReessayer a false pour eviter boucle
                    Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/historique/$cleR.json" `
                        -Method Patch -Body '{"aReessayer":false}' -ContentType "application/json" -TimeoutSec 5 | Out-Null
                    # Relancer VerifierConfirmation.ps1 avec EntryID + NumeroBC + Force
                    $scriptPath = "U:\GromecOutlook\VerifierConfirmation.ps1"
                    $cmd = "powershell -ExecutionPolicy Bypass -File `"$scriptPath`" -EntryID `"$($entreeR.entryID)`" -StoreID `"$($entreeR.storeID)`" -NumeroBC `"$($entreeR.numeroBCManuel)`" -Force"
                    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`" -EntryID `"$($entreeR.entryID)`" -StoreID `"$($entreeR.storeID)`" -NumeroBC `"$($entreeR.numeroBCManuel)`" -Force" -WindowStyle Hidden
                }
            }
        }
    } catch {
        Write-Log "WARN  Erreur verification aReessayer : $($_.Exception.Message)"
    }

    $historique = Get-Historique

    if ($null -ne $historique) {
        foreach ($cle in $historique.PSObject.Properties.Name) {
            $entree = $historique.$cle

            $statut = $entree.dtw_statut
            $statutConfirm = $entree.envoyer_confirmation
            # Traiter seulement les entrees avec quelque chose a faire
            if ($statut -ne 'en_attente' -and $statut -ne 'en_cours' -and $statutConfirm -ne 'en_attente') { continue }

            # ── Confirmation fournisseur (PDF + brouillon Outlook) ──────────────
            if ($entree.envoyer_confirmation -eq 'en_attente') {
                try {
                    Invoke-EnvoyerConfirmationFournisseur -Cle $cle -Entree $entree
                } catch {
                    Write-Log "ERREUR confirmation fournisseur '$cle' : $($_.Exception.Message)"
                    Set-StatutConfirmation -Cle $cle -Statut 'erreur' -Erreur $_.Exception.Message
                }
                continue
            }

            # ── Import DTW (prix / quantites) ────────────────────────────────────
            $copierPrix    = [bool]$entree.dtw_copierPrix
            $copierQty     = [bool]$entree.dtw_copierQty
            $aLignesCochees = ($null -ne $entree.dtw_lignesCochees)
            if (-not $copierPrix -and -not $copierQty -and -not $aLignesCochees) { continue }

            try {
                Invoke-TraiterEntree -Cle $cle -Entree $entree
            } catch {
                Write-Log "ERREUR non geree pour '$cle' : $($_.Exception.Message)"
                Set-StatutDTW -Cle $cle -Statut 'erreur' -Erreur $_.Exception.Message
            }
        }
    }

    # ── ConfirmPO automatique : pousser U_NWR_ConfirmPO=Y pour les commandes conformes ──
    if ($null -ne $historique) {
        foreach ($cle in $historique.PSObject.Properties.Name) {
            $entree = $historique.$cle
            $estOK     = ($entree.statut -eq 'OK')
            $estResolu = ($entree.statut -eq 'ECART' -and $entree.resolu -eq $true)
            if (-not $estOK -and -not $estResolu) { continue }
            $dejaSync = $false
            if ($entree.PSObject.Properties.Name -contains 'syncSAP') {
                $dejaSync = [bool]$entree.syncSAP
            }
            if ($dejaSync) { continue }

            $docNum = $entree.numeroCommande
            if ([string]::IsNullOrEmpty($docNum)) { continue }

            Write-Log "INFO  ConfirmPO auto : PO $docNum (cle: $cle)" -Cle $cle -BC $docNum
            try {
                $contenuConfirm = "DocNum`tDocEntry`tU_NWR_ConfirmPO`r`nDocNum`tDocEntry`tU_NWR_ConfirmPO`r`n$docNum`t`tY"
                [System.IO.File]::WriteAllText($DTW_FichierConfirmPO, $contenuConfirm, [System.Text.Encoding]::Unicode)
            } catch {
                Write-Log "WARN  ConfirmPO auto : erreur ecriture fichier pour PO $docNum : $($_.Exception.Message)" -Cle $cle -BC $docNum
                continue
            }

            $resConfirm = Invoke-DTW -ScenarioXml $DTW_ScenarioConfirmPO
            $maintenant = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
            if ($resConfirm.Succes) {
                Write-Log "INFO  ConfirmPO auto : PO $docNum OK." -Cle $cle -BC $docNum
                try {
                    Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/historique/$cle.json" `
                        -Method Patch `
                        -Body (@{ syncSAP = $true; syncSAPDate = $maintenant; confirme = $true; confirme_le = $maintenant } | ConvertTo-Json -Compress) `
                        -ContentType "application/json" -TimeoutSec 15 | Out-Null
                } catch {}
            } else {
                Write-Log "WARN  ConfirmPO auto : PO $docNum echec : $($resConfirm.Erreur)" -Cle $cle -BC $docNum
                if ($resConfirm.Erreur -like "*deja ouvert*") { break }
                try {
                    Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/historique/$cle.json" `
                        -Method Patch `
                        -Body (@{ syncSAP = $false; syncSAPDate = $maintenant; syncSAPErreur = $resConfirm.Erreur } | ConvertTo-Json -Compress) `
                        -ContentType "application/json" -TimeoutSec 15 | Out-Null
                } catch {}
            }
        }
    }

    # ── Import prix depuis le dashboard (drop zone) ────────────────────────────
    try {
        $importsPrix = Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/imports_prix.json" -Method Get -TimeoutSec 10
        if ($null -ne $importsPrix -and $importsPrix -ne "null") {
            foreach ($cleIP in $importsPrix.PSObject.Properties.Name) {
                $imp = $importsPrix.$cleIP
                if ($imp.statut -ne 'en_attente') { continue }

                $nomFichier = $imp.fichier
                Write-Log "INFO  Import prix dashboard : $nomFichier (cle: $cleIP)"

                try {
                    Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/imports_prix/$cleIP.json" `
                        -Method Patch -Body '{"statut":"en_cours"}' -ContentType "application/json" -TimeoutSec 5 | Out-Null

                    [System.IO.File]::WriteAllText($DTW_FichierPrix, $imp.contenu, [System.Text.Encoding]::Unicode)

                    $resPrix = Invoke-DTW -ScenarioXml $DTW_ScenarioPrix
                    $maintenant = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")

                    if ($resPrix.Succes) {
                        Write-Log "INFO  Import prix dashboard : $nomFichier OK."
                        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/imports_prix/$cleIP.json" `
                            -Method Patch `
                            -Body (@{ statut = 'ok'; date_traitement = $maintenant } | ConvertTo-Json -Compress) `
                            -ContentType "application/json" -TimeoutSec 5 | Out-Null
                    } else {
                        Write-Log "WARN  Import prix dashboard : $nomFichier echec : $($resPrix.Erreur)"
                        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/imports_prix/$cleIP.json" `
                            -Method Patch `
                            -Body (@{ statut = 'erreur'; erreur = $resPrix.Erreur; date_traitement = $maintenant } | ConvertTo-Json -Compress) `
                            -ContentType "application/json" -TimeoutSec 5 | Out-Null
                    }
                } catch {
                    Write-Log "ERREUR Import prix dashboard '$nomFichier' : $($_.Exception.Message)"
                    try {
                        Invoke-RestMethod -Uri "$FirebaseUrl/gromec_vba/imports_prix/$cleIP.json" `
                            -Method Patch `
                            -Body (@{ statut = 'erreur'; erreur = $_.Exception.Message } | ConvertTo-Json -Compress) `
                            -ContentType "application/json" -TimeoutSec 5 | Out-Null
                    } catch {}
                }
            }
        }
    } catch {
        Write-Log "WARN  Erreur lecture noeud imports_prix : $($_.Exception.Message)"
    }

    Start-Sleep -Seconds $IntervalleSecondes
}
