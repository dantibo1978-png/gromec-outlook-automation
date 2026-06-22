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
$DTW_FichierLignes  = "$DTW_Dossier\modif_prix_dans_po.txt"
$DTW_FichierPrixL2  = "$DTW_Dossier\modif_prix_fourn.txt"
$DTW_ScenarioLignes = "$DTW_Dossier\UpdatePOLines_PROD.xml"
$DTW_ScenarioPrixL2 = "$DTW_Dossier\UpdatePriceList2_PROD.xml"
$IntervalleSecondes = 30
$Edge_Exe           = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
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
        dtw_statut    = $Statut
        dtw_traiteLe  = $Maintenant
        # Remettre les flags a false pour ne pas retraiter au prochain poll
        dtw_copierPrix = $false
        dtw_copierQty  = $false
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

    # Verifier qu'aucune instance manuelle de DTW n'est ouverte
    $dejaOuvert = Get-Process -Name "DTW" -ErrorAction SilentlyContinue
    if ($dejaOuvert) {
        return @{ Succes = $false; Erreur = "DTW est deja ouvert manuellement - operation reportee." }
    }

    try {
        $proc = Start-Process -FilePath $DTW_Exe `
            -ArgumentList "-s `"$ScenarioXml`"" `
            -Wait -PassThru -WindowStyle Hidden

        Start-Sleep -Seconds 5

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

    $copierPrix = [bool]$Entree.dtw_copierPrix
    $copierQty  = [bool]$Entree.dtw_copierQty
    $docNum     = $Entree.numeroCommande
    $articles   = @($Entree.articles)

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

function Get-LogoBase64 {
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Logo_Gromec)
        return [System.Convert]::ToBase64String($bytes)
    } catch {
        return ""
    }
}

function New-BonCommandeHTML {
    param([object]$Entree)

    $articles  = @($Entree.articles)
    $bc        = $Entree.numeroCommande
    $date      = (Get-Date).ToString("d/M/yyyy")
    $fourn     = $Entree.fournisseur
    $devise    = if ($Entree.devise) { $Entree.devise } else { "CAD" }
    $logoB64   = Get-LogoBase64
    $logoHtml  = if ($logoB64 -ne "") { "<img src='data:image/png;base64,$logoB64' style='height:60px;'>" } else { "<span style='font-size:24px;font-weight:bold;color:#2d6da3;'>GROMEC</span>" }

    $lignesHtml = ""
    $sousTotal  = 0.0
    $num        = 1
    foreach ($a in $articles) {
        $prixUnit  = if ($a.pdfPrix -and [double]$a.pdfPrix -gt 0) { [double]$a.pdfPrix } else { [double]$a.sapPrix }
        $qty       = if ($a.pdfQty  -and [int]$a.pdfQty    -gt 0) { [int]$a.pdfQty    } else { [int]$a.sapQty    }
        $total     = $prixUnit * $qty
        $sousTotal += $total
        $desc      = if ($a.sapDesc)    { $a.sapDesc    } else { "" }
        $codeManuf = if ($a.sapCode)    { $a.sapCode    } else { "" }
        $article   = if ($a.sapArticle) { $a.sapArticle } else { "" }
        $prixStr   = [string][math]::Round($prixUnit, 2)
        $totalStr  = [string][math]::Round($total, 2)

        $lignesHtml += "<tr>
            <td style='text-align:center;padding:6px 4px;border-bottom:1px solid #e0e0e0;'>$num</td>
            <td style='padding:6px 4px;border-bottom:1px solid #e0e0e0;'>$article</td>
            <td style='padding:6px 4px;border-bottom:1px solid #e0e0e0;'>$codeManuf</td>
            <td style='padding:6px 4px;border-bottom:1px solid #e0e0e0;'>$desc</td>
            <td style='text-align:center;padding:6px 4px;border-bottom:1px solid #e0e0e0;'>$qty</td>
            <td style='text-align:right;padding:6px 4px;border-bottom:1px solid #e0e0e0;'>$prixStr $</td>
            <td style='text-align:right;padding:6px 4px;border-bottom:1px solid #e0e0e0;'>$totalStr $</td>
        </tr>"
        $num++
    }

    $tps        = [math]::Round($sousTotal * 0.05,    2)
    $tvq        = [math]::Round($sousTotal * 0.09975, 2)
    $totalFinal = [math]::Round($sousTotal + $tps + $tvq, 2)
    $sousTotalStr  = [string][math]::Round($sousTotal,  2)
    $tpsStr        = [string][math]::Round($tps,         2)
    $tvqStr        = [string][math]::Round($tvq,         2)
    $totalFinalStr = [string][math]::Round($totalFinal,  2)

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset='UTF-8'>
<style>
  body{font-family:Arial,sans-serif;font-size:11px;margin:20px;color:#222;}
  .header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px;}
  .titre{font-size:28px;font-weight:bold;color:#333;}
  .adresses{display:flex;justify-content:space-between;margin-bottom:20px;}
  .adresse-bloc{width:48%;}
  .adresse-titre{font-weight:bold;border-bottom:2px solid #333;padding-bottom:4px;margin-bottom:8px;}
  table.articles{width:100%;border-collapse:collapse;margin-top:10px;}
  table.articles th{background:#2d6da3;color:white;padding:7px 4px;text-align:left;font-size:10px;}
  table.articles th:nth-child(1),table.articles th:nth-child(5){text-align:center;}
  table.articles th:nth-child(6),table.articles th:nth-child(7){text-align:right;}
  .totaux{margin-top:10px;float:right;width:280px;}
  .totaux table{width:100%;border-collapse:collapse;}
  .totaux td{padding:4px 8px;}
  .totaux .total-final{font-weight:bold;font-size:13px;background:#f0f0f0;}
  .note{margin-top:60px;font-size:10px;border-top:1px solid #ccc;padding-top:10px;}
  .revised-banner{background:#fff3cd;border:1px solid #ffc107;padding:6px 12px;margin-bottom:16px;font-weight:bold;font-size:12px;color:#856404;}
</style>
</head>
<body>
  <div class='header'>
    <div>
      $logoHtml
      <div style='margin-top:10px;font-size:10px;color:#555;'>
        1911 Rue des Outardes<br>Chicoutimi QC, G7K 1C3<br>
        Tel: 418-549-5961<br>Email: gromec@gromec.com
      </div>
    </div>
    <div style='text-align:right;'>
      <div class='titre'>Bon de commande</div>
      <table style='margin-top:10px;font-size:11px;'>
        <tr><td style='color:#555;padding-right:10px;'>No. bon de commande:</td><td><strong>$bc</strong></td></tr>
        <tr><td style='color:#555;'>Date du document:</td><td>$date</td></tr>
        <tr><td style='color:#555;'>Conditions de paiement:</td><td>NET 30 JOURS</td></tr>
        <tr><td style='color:#555;'>Acheteur:</td><td>Daniel Thibault</td></tr>
      </table>
    </div>
  </div>
  <div class='revised-banner'>REVISION -- Prix et/ou quantites mis a jour</div>
  <div class='adresses'>
    <div class='adresse-bloc'>
      <div class='adresse-titre'>Achete a :</div>
      $fourn
    </div>
    <div class='adresse-bloc'>
      <div class='adresse-titre'>Expedie a :</div>
      1911 Rue des Outardes<br>Chicoutimi QC G7K 1C3<br>CANADA
    </div>
  </div>
  <table class='articles'>
    <thead><tr>
      <th style='width:4%;'>#</th>
      <th style='width:12%;'># Produit</th>
      <th style='width:13%;'># Code manuf.</th>
      <th style='width:40%;'>Description</th>
      <th style='width:7%;'>Qte</th>
      <th style='width:12%;'>Prix</th>
      <th style='width:12%;'>Total</th>
    </tr></thead>
    <tbody>$lignesHtml</tbody>
  </table>
  <div class='totaux'>
    <table>
      <tr><td>Sous-Total</td><td style='text-align:right;'>$sousTotalStr $</td></tr>
      <tr><td>TPS 5.000</td><td style='text-align:right;'>$tpsStr $</td></tr>
      <tr><td>TVQ 9.975</td><td style='text-align:right;'>$tvqStr $</td></tr>
      <tr class='total-final'><td>Total $devise</td><td style='text-align:right;'>$totalFinalStr $</td></tr>
    </table>
  </div>
  <div style='clear:both;'></div>
  <div class='note'>NOTE AU FOURNISSEUR: S.V.P. NOUS CONFIRMER LES PRIX ET DELAIS DE LIVRAISON POUR CHAQUE LIGNE DE COMMANDE AU: DTHIBAULT@GROMEC.COM.</div>
  <div style='margin-top:30px;font-size:10px;color:#777;'>USAGER: Daniel Thibault &nbsp;&nbsp; Page: 1 de 1</div>
</body>
</html>
"@
    return $html
}

function New-PDFDepuisHTML {
    param([string]$HtmlContent, [string]$CheminPDF)

    if (-not (Test-Path $Temp_Dossier)) { New-Item -ItemType Directory -Path $Temp_Dossier | Out-Null }
    $htmlTemp = Join-Path $Temp_Dossier "bc_revise_temp.html"

    try {
        [System.IO.File]::WriteAllText($htmlTemp, $HtmlContent, [System.Text.Encoding]::UTF8)
        $urlFichier = "file:///" + $htmlTemp.Replace("\", "/")
        $args = "--headless --disable-gpu --no-sandbox --print-to-pdf=`"$CheminPDF`" --print-to-pdf-no-header `"$urlFichier`""
        $proc = Start-Process -FilePath $Edge_Exe -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
        Start-Sleep -Seconds 3
        if (-not (Test-Path $CheminPDF)) { throw "Edge n'a pas genere le fichier PDF." }
    } finally {
        try { Remove-Item $htmlTemp -Force -ErrorAction SilentlyContinue } catch {}
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
        $brouillon.Attachments.Add($CheminPDF) | Out-Null
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
    $cheminPDF = Join-Path $Temp_Dossier "BC_Revise_$bc.pdf"

    try {
        $html = New-BonCommandeHTML -Entree $Entree
        New-PDFDepuisHTML -HtmlContent $html -CheminPDF $cheminPDF
        New-BrouillonOutlook -Entree $Entree -CheminPDF $cheminPDF
        Set-StatutConfirmation -Cle $Cle -Statut 'envoye'
        Write-Log "INFO  BC $bc -- brouillon cree avec succes."
    } catch {
        Write-Log "ERREUR confirmation fournisseur BC $bc : $($_.Exception.Message)"
        Set-StatutConfirmation -Cle $Cle -Statut 'erreur' -Erreur $_.Exception.Message
    } finally {
        try { if (Test-Path $cheminPDF) { Remove-Item $cheminPDF -Force } } catch {}
    }
}


# ── Boucle principale ─────────────────────────────────────────────────────────
Write-Log "INFO  Sync-DTW.ps1 demarre. Poll toutes les ${IntervalleSecondes}s."

while ($true) {
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
            $copierPrix = [bool]$entree.dtw_copierPrix
            $copierQty  = [bool]$entree.dtw_copierQty
            if (-not $copierPrix -and -not $copierQty) { continue }

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
