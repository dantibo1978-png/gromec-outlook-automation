# =============================================================================
# Sync-DTW.ps1
# Tourne en arriere-plan sur le poste Windows.
# Poll Firebase toutes les 30s, detecte les flags dtw_copierPrix / dtw_copierQty
# poses par le dashboard web, genere les fichiers source DTW et lance DTW.exe.
#
# Lancement (une seule instance, en arriere-plan) :
#   powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "U:\GromecOutlook\Sync-DTW.ps1"
#
# Pour arreter : fermer la fenetre PowerShell ou tuer le processus.
# =============================================================================

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
$IntervalleSecondes = 30
$Logo_Gromec        = "U:\GromecOutlook\logo_gromec.png"
$Temp_Dossier       = "U:\GromecOutlook\temp"
# ──────────────────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$ts] $Message"
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
        dtw_lignesCochees  = $null
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
    Genere le fichier source pour DTW — Document_Lines (UnitPrice et/ou Quantity).
    Format : 2 lignes d'en-tete identiques, puis 1 ligne par article.
    Encodage : UTF-16 LE (Unicode), separateur tabulation.

    Colonnes : ParentKey (DocNum), LineNum (0-indexed), ItemCode, UnitPrice, Quantity
    On inclut toujours les 2 colonnes meme si on n'en modifie qu'une —
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
    Genere le fichier source pour DTW — Price List 2 (Items/ItemPrices).
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

    Write-Log "INFO  Traitement PO $docNum (cle: $Cle) — prix=$copierPrix qty=$copierQty"

    if ($articles.Count -eq 0) {
        Write-Log "WARN  Aucun article dans l'entree Firebase — abandon."
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
        Write-Log "ERREUR DTW lignes : $($res.Erreur)"
        Set-StatutDTW -Cle $Cle -Statut 'erreur' -Erreur "DTW lignes : $($res.Erreur)"
        return
    }
    Write-Log "INFO  DTW lignes OK."

    # ── Etape 2 : Price List 2 (seulement si on copie les prix) ─────────────
    if ($copierPrix) {
        try {
            Write-FichierPrixL2DTW -Articles $articles
        } catch {
            # Non bloquant : les lignes PO sont deja mises a jour
            Write-Log "WARN  Erreur generation fichier Price List 2 : $($_.Exception.Message)"
            Set-StatutDTW -Cle $Cle -Statut 'erreur' -Erreur "DTW lignes OK, mais erreur Price List 2 : $($_.Exception.Message)"
            return
        }

        $res2 = Invoke-DTW -ScenarioXml $DTW_ScenarioPrixL2
        if (-not $res2.Succes) {
            Write-Log "WARN  DTW Price List 2 : $($res2.Erreur)"
            Set-StatutDTW -Cle $Cle -Statut 'erreur' -Erreur "DTW lignes OK, mais DTW Price List 2 : $($res2.Erreur)"
            return
        }
        Write-Log "INFO  DTW Price List 2 OK."
    }

    # ── Etape 3 : U_NWR_ConfirmPO = Y (confirmer la commande dans SAP) ────────
    try {
        $contenuConfirm = "DocNum`tDocEntry`tU_NWR_ConfirmPO`r`nDocNum`tDocEntry`tU_NWR_ConfirmPO`r`n$docNum`t`tY"
        [System.IO.File]::WriteAllText($DTW_FichierConfirmPO, $contenuConfirm, [System.Text.Encoding]::Unicode)
    } catch {
        Write-Log "WARN  Erreur generation fichier ConfirmPO : $($_.Exception.Message)"
        Set-StatutDTW -Cle $Cle -Statut 'ok'
        Write-Log "INFO  PO $docNum traite avec succes (ConfirmPO non mis a jour)."
        return
    }

    $res3 = Invoke-DTW -ScenarioXml $DTW_ScenarioConfirmPO
    if (-not $res3.Succes) {
        Write-Log "WARN  DTW ConfirmPO : $($res3.Erreur)"
        Set-StatutDTW -Cle $Cle -Statut 'ok'
        Write-Log "INFO  PO $docNum traite avec succes (ConfirmPO non mis a jour)."
        return
    }
    Write-Log "INFO  DTW ConfirmPO OK."

    Set-StatutDTW -Cle $Cle -Statut 'ok'
    Write-Log "INFO  PO $docNum traite avec succes."
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

    $iTextDLL = "U:\GromecOutlook\lib\itextsharp\lib\itextsharp.dll"
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
    Write-Log "INFO  Confirmation fournisseur BC $bc (cle: $Cle)."

    if (-not (Test-Path $Temp_Dossier)) { New-Item -ItemType Directory -Path $Temp_Dossier | Out-Null }

    $cheminPDFOrig  = Join-Path $Temp_Dossier "BC_Original_$bc.pdf"
    $cheminPDFAnnote = Join-Path $Temp_Dossier "BC_Revise_$bc.pdf"

    try {
        # 1. Recuperer le PDF SAP original depuis la PJ du courriel
        $pdfTrouve = Get-PDFOriginalDepuisOutlook -EntryID $Entree.entryID -StoreID $Entree.storeID -CheminDestination $cheminPDFOrig

        if ($pdfTrouve) {
            # 2. Annoter le PDF original avec iTextSharp
            New-PDFAnnote -CheminPDFSource $cheminPDFOrig -CheminPDFDest $cheminPDFAnnote -Articles @($Entree.articles)
            $cheminPJFinal = $cheminPDFAnnote
            Write-Log "INFO  PDF original annote avec succes."
        } else {
            # Fallback : pas de PDF original trouve -- on avertit mais on continue sans PJ
            Write-Log "WARN  PDF original introuvable dans la PJ -- brouillon sans PJ PDF."
            $cheminPJFinal = $null
        }

        # 3. Creer le brouillon Outlook
        New-BrouillonOutlook -Entree $Entree -CheminPDF $cheminPJFinal
        Set-StatutConfirmation -Cle $Cle -Statut 'envoye'
        Write-Log "INFO  BC $bc -- brouillon cree avec succes."

    } catch {
        Write-Log "ERREUR confirmation fournisseur BC $bc : $($_.Exception.Message)"
        Set-StatutConfirmation -Cle $Cle -Statut 'erreur' -Erreur $_.Exception.Message
    } finally {
        try { if (Test-Path $cheminPDFOrig)   { Remove-Item $cheminPDFOrig   -Force } } catch {}
        try { if (Test-Path $cheminPDFAnnote) { Remove-Item $cheminPDFAnnote -Force } } catch {}
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

        # Si c est une confirmation, relancer Verifier-Confirmation.ps1 avec -Force
        if ($estConfirmation) {
            $entryID = $Reclassif.entryID
            $storeID = $Reclassif.storeID
            $ps1Path = "U:\GromecOutlook\Verifier-Confirmation.ps1"

            Write-Log "INFO  Retraitement force : lancement de Verifier-Confirmation.ps1 pour $sujet"
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
Write-Log "INFO  Sync-DTW.ps1 demarre. Poll toutes les ${IntervalleSecondes}s."

while ($true) {
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

    $historique = Get-Historique

    if ($null -ne $historique) {
        foreach ($cle in $historique.PSObject.Properties.Name) {
            $entree = $historique.$cle

            $statut = $entree.dtw_statut
            $statutConfirm = $entree.envoyer_confirmation
            # Traiter seulement les entrees avec quelque chose a faire
            if ($statut -ne 'en_attente' -and $statutConfirm -ne 'en_attente') { continue }

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

    Start-Sleep -Seconds $IntervalleSecondes
}
