function New-PDFBonDeCommande {
    <#
    .SYNOPSIS
        Generates a professional PDF Purchase Order in SAP Business One style.
    .DESCRIPTION
        Creates a revised PO PDF ("Bon de commande revise") using raw PDF format.
        No external libraries required - works with pure PowerShell on Windows.
        Modified lines (price/quantity changes) are highlighted in yellow with
        old values shown struck through in red.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$NumeroCommande,

        [Parameter(Mandatory=$true)]
        [string]$Fournisseur,

        [Parameter(Mandatory=$true)]
        [array]$Articles,

        [string]$Devise = "CAD",

        [Parameter(Mandatory=$true)]
        [string]$CheminDestination,

        [string]$DateDocument = (Get-Date -Format "yyyy-MM-dd"),
        [string]$DateLivraison = "",
        [string]$ConditionsPaiement = "",
        [string]$AdresseFournisseur = "",
        [string]$NoCompteFournisseur = "",
        [string]$LogoPath = "U:\GromecOutlook\logo_gromec.png",
        [string]$NomAcheteur = "",
        [string]$AdresseLigne1 = "",
        [string]$AdresseLigne2 = "",
        [string]$CourrielContact = "",
        [double]$TauxTPS = 0,
        [double]$TauxTVQ = 0
    )

    # Charger les parametres depuis Firebase si pas fournis en argument
    if (-not $NomAcheteur -or -not $AdresseLigne1) {
        try {
            $fbUrl = "https://gromec-outlook-vba-default-rtdb.firebaseio.com"
            $pv = Invoke-RestMethod -Uri "$fbUrl/gromec_vba/parametres/valeurs.json" -Method Get -TimeoutSec 10
            if ($pv) {
                if (-not $NomAcheteur     -and $pv.nom_acheteur)     { $NomAcheteur     = $pv.nom_acheteur }
                if (-not $AdresseLigne1   -and $pv.adresse_ligne1)   { $AdresseLigne1   = $pv.adresse_ligne1 }
                if (-not $AdresseLigne2   -and $pv.adresse_ligne2)   { $AdresseLigne2   = $pv.adresse_ligne2 }
                if (-not $CourrielContact -and $pv.courriel_contact) { $CourrielContact = $pv.courriel_contact }
                if ($TauxTPS -eq 0 -and $pv.tps) { $TauxTPS = [double]$pv.tps / 100.0 }
                if ($TauxTVQ -eq 0 -and $pv.tvq) { $TauxTVQ = [double]$pv.tvq / 100.0 }
            }
        } catch {}
    }
    if (-not $NomAcheteur)     { $NomAcheteur     = "Daniel Thibault" }
    if (-not $AdresseLigne1)   { $AdresseLigne1   = "1911 Rue des Outardes" }
    if (-not $AdresseLigne2)   { $AdresseLigne2   = "Chicoutimi QC G7K 1C3" }
    if (-not $CourrielContact) { $CourrielContact = "DTHIBAULT@GROMEC.COM" }
    if ($TauxTPS -eq 0) { $TauxTPS = 0.05 }
    if ($TauxTVQ -eq 0) { $TauxTVQ = 0.09975 }

    # ── Encoding ──────────────────────────────────────────────────────
    $latin1 = [System.Text.Encoding]::GetEncoding("ISO-8859-1")

    # Accented characters compatible with PowerShell 5.1
    $eAcute = [char]0xe9    # é
    $eGrave = [char]0xe8    # è
    $aGrave = [char]0xe0    # à
    $eCirc  = [char]0xea    # ê
    $EAcute = [char]0xc9    # É

    function Esc([string]$s) {
        if (-not $s) { return "" }
        return $s.Replace('\','\\').Replace('(', '\(').Replace(')', '\)')
    }

    # ── Mark modified lines ───────────────────────────────────────────
    foreach ($a in $Articles) {
        $mod = $false
        if ($null -eq $a.pdfQty)  { $a | Add-Member -NotePropertyName pdfQty  -NotePropertyValue $a.sapQty  -Force }
        if ($null -eq $a.pdfPrix) { $a | Add-Member -NotePropertyName pdfPrix -NotePropertyValue $a.sapPrix -Force }
        if ([double]$a.pdfPrix -ne [double]$a.sapPrix) { $mod = $true }
        if ([double]$a.pdfQty  -ne [double]$a.sapQty)  { $mod = $true }
        $a | Add-Member -NotePropertyName isModified -NotePropertyValue $mod -Force
    }

    # ── Page constants (Letter) ───────────────────────────────────────
    $pw = 612; $ph = 792
    $ml = 40; $mr = 572

    # ── Calculate totals ──────────────────────────────────────────────
    $sousTotal = 0.0
    foreach ($a in $Articles) { $sousTotal += [double]$a.pdfQty * [double]$a.pdfPrix }
    $sousTotal = [Math]::Round($sousTotal, 2)
    $tps = [Math]::Round($sousTotal * $TauxTPS, 2)
    $tvq = [Math]::Round($sousTotal * $TauxTVQ, 2)
    $grandTotal = [Math]::Round($sousTotal + $tps + $tvq, 2)

    # ── Pagination ────────────────────────────────────────────────────
    $maxFirst = 18; $maxNext = 30
    $pageSlices = @()
    $si = 0
    while ($si -lt $Articles.Count -or $pageSlices.Count -eq 0) {
        $cap = if ($pageSlices.Count -eq 0) { $maxFirst } else { $maxNext }
        $n = [Math]::Min($cap, [Math]::Max(0, $Articles.Count - $si))
        $pageSlices += ,@($si, $n)
        $si += [Math]::Max($n, 1)
        if ($n -eq 0) { break }
    }
    $totalPages = $pageSlices.Count

    # ── Load logo ─────────────────────────────────────────────────────
    $hasLogo = $false
    $logoCompBytes = $null; $logoRawLen = 0; $logoW = 0; $logoH = 0
    if (Test-Path $LogoPath -ErrorAction SilentlyContinue) {
        try {
            $ms = New-Object System.IO.MemoryStream(,[System.IO.File]::ReadAllBytes($LogoPath))
            $bmp = New-Object System.Drawing.Bitmap($ms)
            $logoW = $bmp.Width; $logoH = $bmp.Height
            $rgb = New-Object byte[] ($logoW * $logoH * 3)
            $bi = 0
            for ($py = 0; $py -lt $logoH; $py++) {
                for ($px = 0; $px -lt $logoW; $px++) {
                    $c = $bmp.GetPixel($px, $py)
                    $rgb[$bi++] = $c.R; $rgb[$bi++] = $c.G; $rgb[$bi++] = $c.B
                }
            }
            $bmp.Dispose(); $ms.Dispose()
            $msC = New-Object System.IO.MemoryStream
            $ds = New-Object System.IO.Compression.DeflateStream($msC, [System.IO.Compression.CompressionMode]::Compress, $true)
            $ds.Write($rgb, 0, $rgb.Length); $ds.Close()
            $logoCompBytes = $msC.ToArray(); $logoRawLen = $rgb.Length
            $msC.Dispose()
            $hasLogo = $true
        } catch { $hasLogo = $false }
    }

    # ── Build all content streams first ───────────────────────────────
    $streamDataList = @()

    for ($pi = 0; $pi -lt $totalPages; $pi++) {
        $sIdx = $pageSlices[$pi][0]
        $cnt  = $pageSlices[$pi][1]
        $pg1  = ($pi -eq 0)
        $pgNum = $pi + 1

        $sb = [System.Text.StringBuilder]::new(8192)

        # Helpers write into $sb
        function T([double]$x,[double]$y,[string]$t,[string]$f,[double]$sz) {
            [void]$sb.AppendLine("BT /$f $sz Tf $x $y Td ($(Esc $t)) Tj ET")
        }
        function TR([double]$xR,[double]$y,[string]$t,[string]$f,[double]$sz) {
            # Right-aligned: approximate char width
            $cw = if ($f -eq 'F2') { $sz * 0.56 } else { $sz * 0.52 }
            $x = $xR - ($t.Length * $cw)
            T $x $y $t $f $sz
        }
        function LN([double]$x1,[double]$y1,[double]$x2,[double]$y2,[double]$w) {
            [void]$sb.AppendLine("$w w $x1 $y1 m $x2 $y2 l S")
        }
        function RF([double]$x,[double]$y,[double]$w,[double]$h,[double]$r,[double]$g,[double]$b) {
            [void]$sb.AppendLine("$r $g $b rg $x $y $w $h re f 0 0 0 rg")
        }
        function RFS([double]$x,[double]$y,[double]$w,[double]$h,[double]$r,[double]$g,[double]$b) {
            [void]$sb.AppendLine("$r $g $b rg 0 0 0 RG 0.5 w $x $y $w $h re B 0 0 0 rg")
        }
        function RS([double]$x,[double]$y,[double]$w,[double]$h) {
            [void]$sb.AppendLine("0.7 0.7 0.7 RG 0.3 w $x $y $w $h re S 0 0 0 RG")
        }

        [void]$sb.AppendLine("0 0 0 RG 0 0 0 rg 1 w")

        $cy = 752

        if ($pg1) {
            # ── Logo ──
            if ($hasLogo) {
                $dw = 150; $dh = [Math]::Round(150 * $logoH / $logoW)
                [void]$sb.AppendLine("q $dw 0 0 $dh $ml $($cy - $dh) cm /Im1 Do Q")
            }

            # ── Title ──
            T 355 ($cy - 10) "Bon de commande r${eAcute}vis${eAcute}" "F2" 16

            # ── Company info ──
            $iy = $cy - 80
            T $ml $iy (Esc $AdresseLigne1) "F1" 8
            T $ml ($iy-10) (Esc $AdresseLigne2) "F1" 8
            T $ml ($iy-20) "T${eAcute}l: 418-549-5961" "F1" 8
            T $ml ($iy-30) "Email: gromec@gromec.com" "F1" 8

            # ── Info box (right) ──
            $bx = 340; $bw2 = 232; $lx = $bx+5; $vx = $bx+$bw2-5
            $by = $cy - 35; $rh = 13

            RFS $bx $by $bw2 14 0.17 0.24 0.45
            [void]$sb.AppendLine("1 1 1 rg")
            T $lx ($by+3) "Informations" "F2" 9
            [void]$sb.AppendLine("0 0 0 rg")

            $fields = @(
                @("No. bon de commande:", $NumeroCommande),
                @("Date du document:", $DateDocument),
                @("Conditions de paiement:", $ConditionsPaiement),
                @("Date de livraison:", $DateLivraison),
                @("Acheteur:", $NomAcheteur),
                @("F.A.B.:", ""),
                @("No Compte:", $NoCompteFournisseur)
            )
            $fy = $by - $rh; $alt = $false
            foreach ($fld in $fields) {
                if ($alt) { RF $bx $fy $bw2 $rh 0.93 0.93 0.93 }
                RS $bx $fy $bw2 $rh
                T $lx ($fy+3) $fld[0] "F2" 7.5
                TR $vx ($fy+3) $fld[1] "F1" 7.5
                $fy -= $rh; $alt = -not $alt
            }

            # ── Addresses ──
            $ay = $fy - 10; $mid = 306
            RFS $ml $ay ($mid-$ml-5) 14 0.17 0.24 0.45
            [void]$sb.AppendLine("1 1 1 rg"); T ($ml+5) ($ay+3) "Achet${eAcute} ${aGrave}:" "F2" 9; [void]$sb.AppendLine("0 0 0 rg")
            RFS ($mid+5) $ay ($mr-$mid-5) 14 0.17 0.24 0.45
            [void]$sb.AppendLine("1 1 1 rg"); T ($mid+10) ($ay+3) "Exp${eAcute}di${eAcute} ${aGrave}:" "F2" 9; [void]$sb.AppendLine("0 0 0 rg")

            # Supplier
            $sa = $ay - 13
            T ($ml+5) $sa $Fournisseur "F2" 8
            if ($AdresseFournisseur) {
                foreach ($ln in ($AdresseFournisseur -split "`n")) {
                    $sa -= 11; T ($ml+5) $sa $ln.Trim() "F1" 8
                }
            }

            # Ship-to
            $sa2 = $ay - 13
            T ($mid+10) $sa2 "GROMEC" "F2" 8
            $sa2 -= 11; T ($mid+10) $sa2 (Esc $AdresseLigne1) "F1" 8
            $sa2 -= 11; T ($mid+10) $sa2 (Esc $AdresseLigne2) "F1" 8
            $sa2 -= 11; T ($mid+10) $sa2 "CANADA" "F1" 8
            $sa2 -= 11; T ($mid+10) $sa2 "T${eAcute}l:" "F1" 8

            $abh = 70
            RS $ml ($ay-$abh) ($mid-$ml-5) $abh
            RS ($mid+5) ($ay-$abh) ($mr-$mid-5) $abh

            $tableY = $ay - $abh - 20
        } else {
            T $ml ($cy-5) "GROMEC - Bon de commande r${eAcute}vis${eAcute}" "F2" 12
            TR $mr ($cy-5) "No: $NumeroCommande" "F1" 10
            $tableY = $cy - 30
        }

        # ── Table header ──
        $tw = $mr - $ml
        $cols = @($ml, 65, 150, 240, 420, 465, 525)
        $hdrs = @("#","# Produit","# Code manuf.","Description.","Qt${eAcute}","Prix","Total")
        $thH = 16

        RFS $ml $tableY $tw $thH 0.17 0.24 0.45
        [void]$sb.AppendLine("1 1 1 rg")
        for ($ci=0; $ci -lt $hdrs.Count; $ci++) { T ($cols[$ci]+3) ($tableY+4) $hdrs[$ci] "F2" 7 }
        [void]$sb.AppendLine("0 0 0 rg")

        # ── Table rows ──
        $rowH = 14; $ry = $tableY - $rowH

        for ($ri=0; $ri -lt $cnt; $ri++) {
            $a = $Articles[$sIdx + $ri]
            $lineNum = $sIdx + $ri + 1
            $qty = [double]$a.pdfQty; $prix = [double]$a.pdfPrix
            $lt = [Math]::Round($qty * $prix, 2)

            if ($a.isModified) {
                RF $ml $ry $tw $rowH 1.0 0.95 0.80
            } elseif ($ri % 2 -eq 1) {
                RF $ml $ry $tw $rowH 0.96 0.96 0.96
            }

            [void]$sb.AppendLine("0.85 0.85 0.85 RG 0.3 w $ml $ry $tw $rowH re S 0 0 0 RG")

            $ty = $ry + 3
            T ($cols[0]+3) $ty "$lineNum" "F1" 7

            $artCode = if ($a.sapArticle) { $a.sapArticle } else { "" }
            T ($cols[1]+3) $ty $artCode "F1" 7

            $cm = if ($a.sapCodeManuf) { $a.sapCodeManuf } else { "" }
            T ($cols[2]+3) $ty $cm "F1" 7

            $desc = if ($a.description) { $a.description } else { "" }
            if ($desc.Length -gt 30) { $desc = $desc.Substring(0,30) + "..." }
            T ($cols[3]+3) $ty $desc "F1" 7

            # Qty
            TR ($cols[5]-5) $ty $qty.ToString("N2") "F1" 7
            if ($a.isModified -and [double]$a.pdfQty -ne [double]$a.sapQty) {
                $oq = ([double]$a.sapQty).ToString("N2")
                [void]$sb.AppendLine("0.7 0.0 0.0 rg")
                T ($cols[4]+2) ($ty-1) "($oq)" "F3" 5.5
                $oxw = $oq.Length * 2.8 + 8
                LN ($cols[4]+2) ($ty+1) ($cols[4]+2+$oxw) ($ty+1) 0.4
                [void]$sb.AppendLine("0 0 0 rg")
            }

            # Prix
            TR ($cols[6]-5) $ty $prix.ToString("N4") "F1" 7
            if ($a.isModified -and [double]$a.pdfPrix -ne [double]$a.sapPrix) {
                $op = ([double]$a.sapPrix).ToString("N4")
                [void]$sb.AppendLine("0.7 0.0 0.0 rg")
                T ($cols[5]+2) ($ty-1) "($op)" "F3" 5.5
                $opw = $op.Length * 2.8 + 8
                LN ($cols[5]+2) ($ty+1) ($cols[5]+2+$opw) ($ty+1) 0.4
                [void]$sb.AppendLine("0 0 0 rg")
            }

            # Total
            TR ($mr-5) $ty $lt.ToString("N2") "F1" 7

            $ry -= $rowH
        }

        # ── Footer on last page ──
        if ($pi -eq $totalPages - 1) {
            $fy2 = $ry - 15; $tx = 400; $tl = $tx+5; $tv = $mr-5; $trh = 14

            # Sous-Total
            RF $tx $fy2 ($mr-$tx) $trh 0.93 0.93 0.93; RS $tx $fy2 ($mr-$tx) $trh
            T $tl ($fy2+3) "Sous-Total" "F2" 8; TR $tv ($fy2+3) $sousTotal.ToString("N2") "F1" 8

            $fy2 -= $trh; RS $tx $fy2 ($mr-$tx) $trh
            T $tl ($fy2+3) "TPS $([Math]::Round($TauxTPS * 100, 3).ToString('N3'))" "F1" 8; TR $tv ($fy2+3) $tps.ToString("N2") "F1" 8

            $fy2 -= $trh; RF $tx $fy2 ($mr-$tx) $trh 0.93 0.93 0.93; RS $tx $fy2 ($mr-$tx) $trh
            T $tl ($fy2+3) "TVQ $([Math]::Round($TauxTVQ * 100, 3).ToString('N3'))" "F1" 8; TR $tv ($fy2+3) $tvq.ToString("N2") "F1" 8

            $fy2 -= $trh; RFS $tx $fy2 ($mr-$tx) $trh 0.17 0.24 0.45
            [void]$sb.AppendLine("1 1 1 rg")
            T $tl ($fy2+3) "Total $Devise" "F2" 9; TR $tv ($fy2+3) $grandTotal.ToString("N2") "F2" 9
            [void]$sb.AppendLine("0 0 0 rg")

            # Signature
            $sy = $fy2 - 5
            LN $ml ($sy-30) 250 ($sy-30) 0.5
            T $ml ($sy-42) "Signataire autoris${eAcute}" "F3" 8

            # Note
            $ny = $sy - 65
            [void]$sb.AppendLine("0.15 0.15 0.15 rg")
            T $ml $ny      "NOTE AU FOURNISSEUR: S.V.P. NOUS CONFIRMER LES PRIX ET D${EAcute}LAIS DE LIVRAISON POUR CHAQUE LIGNE DE" "F2" 6
            T $ml ($ny-8)  "COMMANDE AU: $(Esc $CourrielContact). TOUTE MODIFICATION DE PRIX DEVRA ${eCirc}TRE COMMUNIQU${EAcute}E DANS LES" "F2" 6
            T $ml ($ny-16) "24 HEURES. SANS QUOI LE PAIEMENT SERA FAIT SELON LES PRIX DE NOTRE COMMANDE." "F2" 6
            [void]$sb.AppendLine("0 0 0 rg")

            # Revision legend
            [void]$sb.AppendLine("0.6 0.0 0.0 rg")
            T $ml ($ny-30) "* Les lignes surlign`u{00e9}es en jaune ont `u{00e9}t`u{00e9} modifi`u{00e9}es. Les anciennes valeurs sont affich`u{00e9}es barr`u{00e9}es en rouge." "F3" 6.5
            [void]$sb.AppendLine("0 0 0 rg")
        }

        # Page number
        T $ml 33 "USAGER: $(Esc $NomAcheteur)" "F1" 7
        TR $mr 33 "Page: $pgNum de $totalPages" "F1" 7

        $streamDataList += ,$latin1.GetBytes($sb.ToString())
    }

    # ── Assemble PDF ──────────────────────────────────────────────────
    # Object layout:
    #   1 = Catalog
    #   2 = Pages
    #   3 = Font Helvetica
    #   4 = Font Helvetica-Bold
    #   5 = Font Helvetica-Oblique
    #   6 = Logo image (if present)
    #   then for each page: content stream obj, page obj

    $nextO = 6
    $logoObj = 0
    if ($hasLogo) { $logoObj = $nextO; $nextO++ }

    $contentObjs = @()
    $pageObjs = @()
    for ($p = 0; $p -lt $totalPages; $p++) {
        $contentObjs += $nextO; $nextO++
        $pageObjs += $nextO; $nextO++
    }

    $totalObjs = $nextO - 1
    $offsets = New-Object long[] ($totalObjs + 1) # 1-indexed

    # Open file
    $fs = [System.IO.FileStream]::new($CheminDestination, [System.IO.FileMode]::Create)

    function WB([byte[]]$b) { $fs.Write($b, 0, $b.Length) }
    function WS([string]$s) { WB ($latin1.GetBytes($s)) }

    WS "%PDF-1.4`n"
    WB ([byte[]]@(0x25, 0xC0, 0xC1, 0xC2, 0xC3, 0x0A))

    # Obj 1: Catalog
    $offsets[1] = $fs.Position
    WS "1 0 obj`n<< /Type /Catalog /Pages 2 0 R >>`nendobj`n"

    # Obj 2: Pages
    $kids = ($pageObjs | ForEach-Object { "$_ 0 R" }) -join " "
    $offsets[2] = $fs.Position
    WS "2 0 obj`n<< /Type /Pages /Kids [$kids] /Count $totalPages >>`nendobj`n"

    # Obj 3-5: Fonts
    $offsets[3] = $fs.Position
    WS "3 0 obj`n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>`nendobj`n"
    $offsets[4] = $fs.Position
    WS "4 0 obj`n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>`nendobj`n"
    $offsets[5] = $fs.Position
    WS "5 0 obj`n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Oblique /Encoding /WinAnsiEncoding >>`nendobj`n"

    # Obj 6: Logo (if present)
    if ($hasLogo) {
        $offsets[$logoObj] = $fs.Position
        WS "$logoObj 0 obj`n"
        WS "<< /Type /XObject /Subtype /Image /Width $logoW /Height $logoH /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length $($logoCompBytes.Length) >>`n"
        WS "stream`n"
        WB $logoCompBytes
        WS "`nendstream`n"
        WS "endobj`n"
    }

    # Page content streams + page objects
    for ($p = 0; $p -lt $totalPages; $p++) {
        $cObj = $contentObjs[$p]
        $pObj = $pageObjs[$p]
        $sData = $streamDataList[$p]
        $sLen = $sData.Length

        # Content stream
        $offsets[$cObj] = $fs.Position
        WS "$cObj 0 obj`n<< /Length $sLen >>`nstream`n"
        WB $sData
        WS "`nendstream`nendobj`n"

        # Page object
        $res = "/Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >>"
        if ($hasLogo -and $p -eq 0) {
            $res = "/Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >> /XObject << /Im1 $logoObj 0 R >>"
        }
        $offsets[$pObj] = $fs.Position
        WS "$pObj 0 obj`n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 $pw $ph] /Contents $cObj 0 R /Resources << $res >> >>`nendobj`n"
    }

    # XRef
    $xrefPos = $fs.Position
    WS "xref`n0 $($totalObjs + 1)`n"
    WS "0000000000 65535 f `r`n"
    for ($o = 1; $o -le $totalObjs; $o++) {
        WS "$($offsets[$o].ToString('D10')) 00000 n `r`n"
    }

    WS "trailer`n<< /Size $($totalObjs + 1) /Root 1 0 R >>`nstartxref`n$xrefPos`n%%EOF`n"

    $fs.Close()
    Write-Host "PDF generated: $CheminDestination" -ForegroundColor Green
    return $CheminDestination
}
