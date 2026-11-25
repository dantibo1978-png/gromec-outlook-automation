<#
.SYNOPSIS
    Recherche de commandes d'achat — Vérifie quelles commandes SAP ont été envoyées par courriel.
.DESCRIPTION
    1. Se connecte à SAP via DI API (SAPbobsCOM)
    2. Récupère les commandes d'achat ouvertes de OPOR
    3. Vérifie dans les éléments envoyés d'Outlook si un courriel contenant le numéro a été envoyé
    4. Génère un fichier Excel avec le résultat
.NOTES
    Exécuter avec:  & "U:\GromecOutlook\RechercheCommandes.ps1"
#>

# ── Auto-relance en 64-bit si nécessaire ──────────────────────────────────────
if (-not [Environment]::Is64BitProcess) {
    $ps64 = "$env:SystemRoot\SysNative\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $ps64) {
        Start-Process $ps64 -ArgumentList @(
            "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath
        ) -WindowStyle Normal
        exit 0
    }
}

# ── Configuration ─────────────────────────────────────────────────────────────
$SAP_Server    = "NDB@hana:30013"
$SAP_DbType    = 9
$SAP_CompanyDB = "SAP_GROMEC_PROD"
$SAP_User      = ""
$SAP_Password  = ""

try {
    $configPath = Join-Path (Split-Path $PSCommandPath) "config.json"
    if (Test-Path $configPath) {
        $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
        if ($cfg.sap_user)     { $SAP_User     = [string]$cfg.sap_user }
        if ($cfg.sap_password) { $SAP_Password = [string]$cfg.sap_password }
        if ($cfg.sap_db_prod)  { $SAP_CompanyDB = [string]$cfg.sap_db_prod }
    }
} catch {}

# ── Paramètres par défaut ─────────────────────────────────────────────────────
$JoursRecherche      = 90       # Chercher les PO des X derniers jours
$ExclureCreateurs    = @("Gagnon Marie-Christine", "St-Laurent Charles")
$ExclureFournisseurs = @()
$ExclureStatuts      = @("C")   # C = Cloture
$ExclurePOs          = @()

# ── Charger les exclusions sauvegardees ──────────────────────────────────────
$exclusionFile = Join-Path (Split-Path $PSCommandPath) "RechercheCommandes_exclusions.json"
try {
    if (Test-Path $exclusionFile) {
        $excData = Get-Content $exclusionFile -Raw | ConvertFrom-Json
        if ($excData.po_exclus) {
            $ExclurePOs = @($excData.po_exclus)
        }
        if ($excData.createurs_exclus) {
            $ExclureCreateurs = @($excData.createurs_exclus)
        }
    }
} catch {}

# ── Fonctions utilitaires ─────────────────────────────────────────────────────
function Write-Log {
    param([string]$msg)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$ts  $msg"
}

function Show-LoginDialog {
    $tempHta    = "$env:TEMP\sap_login_rc.hta"
    $resultFile = "$env:TEMP\sap_login_rc_result.txt"
    if (Test-Path $resultFile) { Remove-Item $resultFile -Force }

    $hta = @"
<html><head><title>Connexion SAP - Recherche Commandes</title>
<hta:application border="dialog" maximizeButton="no" minimizeButton="no" scroll="no" />
<style>body{font:12px Tahoma;margin:20px}td{padding:5px}input{width:160px;padding:3px}
button{width:85px;padding:5px;cursor:pointer}</style>
<script language="VBScript">
Sub DoOK()
  Dim f : Set f = CreateObject("Scripting.FileSystemObject").CreateTextFile("$resultFile", True)
  f.WriteLine document.getElementById("usr").value
  f.WriteLine document.getElementById("pwd").value
  f.Close : window.close
End Sub
Sub DoCancel() : window.close : End Sub
Sub DoKeyPress(k) : If k = 13 Then Call DoOK() : End Sub
Sub Window_OnLoad() : window.resizeTo 350, 220 : document.getElementById("usr").focus : End Sub
</script></head><body>
<table><tr><td><b>Usager</b></td><td><input id="usr" value="$SAP_User"></td></tr>
<tr><td><b>Mot de passe</b></td><td><input id="pwd" type="password" onkeypress="DoKeyPress(window.event.keyCode)"></td></tr></table><br>
<center><button onclick="DoOK()">OK</button>&nbsp;&nbsp;
<button onclick="DoCancel()">ANNULER</button></center></body></html>
"@
    Set-Content -Path $tempHta -Value $hta -Encoding UTF8
    $proc = Start-Process "mshta.exe" -ArgumentList """$tempHta""" -PassThru -Wait
    if (Test-Path $resultFile) {
        $lines = Get-Content $resultFile
        Remove-Item $resultFile -Force
        if (Test-Path $tempHta) { Remove-Item $tempHta -Force }
        return @{ User = $lines[0]; Pass = $lines[1] }
    }
    if (Test-Path $tempHta) { Remove-Item $tempHta -Force }
    return $null
}

# ── Connexion SAP ─────────────────────────────────────────────────────────────
Write-Log "Connexion a SAP..."
if (-not $SAP_User -or -not $SAP_Password) {
    $login = Show-LoginDialog
    if (-not $login) { Write-Log "Annule par l'utilisateur."; exit 1 }
    $SAP_User     = $login.User
    $SAP_Password = $login.Pass
}

$company = New-Object -ComObject SAPbobsCOM.Company
$company.Server       = $SAP_Server
$company.DbServerType = $SAP_DbType
$company.CompanyDB    = $SAP_CompanyDB
$company.UseTrusted   = $false
$company.UserName     = $SAP_User
$company.Password     = $SAP_Password

$ret = $company.Connect()
if ($ret -ne 0) {
    $errMsg = $company.GetLastErrorDescription()
    Write-Log "ERREUR connexion SAP ($ret): $errMsg"
    [System.Windows.Forms.MessageBox]::Show("Connexion SAP echouee ($ret): $errMsg", "Erreur", 0, 16) | Out-Null
    exit 1
}
Write-Log "Connecte a SAP ($SAP_CompanyDB)"

function Run-SAPQuery {
    param([string]$sql)
    $rs = $company.GetBusinessObject(300)  # BoRecordset
    $rs.DoQuery($sql)
    return $rs
}

# ── Récupérer les commandes d'achat ──────────────────────────────────────────
$dateDebut = (Get-Date).AddDays(-$JoursRecherche).ToString("yyyyMMdd")
Write-Log "Requete OPOR - commandes depuis $dateDebut..."

$sqlPO = 'SELECT T0."DocNum", T0."DocDate", T0."CardCode", T0."CardName", T0."DocTotal", T0."DocCur", T0."DocStatus", IFNULL(T1."U_NAME", ' + "''" + ') AS "Createur" FROM "OPOR" T0 LEFT JOIN "OUSR" T1 ON T0."UserSign" = T1."INTERNAL_K" WHERE T0."DocDate" >= ' + "'" + $dateDebut + "'" + ' ORDER BY T0."DocNum" DESC'

$rsPO = Run-SAPQuery $sqlPO

$commandes = @()
while (-not $rsPO.EoF) {
    $docNum    = $rsPO.Fields.Item("DocNum").Value
    $docDate   = $rsPO.Fields.Item("DocDate").Value
    $cardCode  = $rsPO.Fields.Item("CardCode").Value
    $cardName  = $rsPO.Fields.Item("CardName").Value
    $total     = $rsPO.Fields.Item("DocTotal").Value
    $devise    = $rsPO.Fields.Item("DocCur").Value
    $status    = $rsPO.Fields.Item("DocStatus").Value
    $createur  = $rsPO.Fields.Item("Createur").Value

    # Exclure cloturees
    if ($ExclureStatuts -contains $status) { $rsPO.MoveNext(); continue }

    # Exclure POs marques
    if ($ExclurePOs -contains [string]$docNum) { $rsPO.MoveNext(); continue }

    # Exclure créateurs
    $skip = $false
    foreach ($exc in $ExclureCreateurs) {
        if ($createur -and $createur -like "*$exc*") { $skip = $true; break }
    }
    if ($skip) { $rsPO.MoveNext(); continue }

    # Exclure fournisseurs
    foreach ($exc in $ExclureFournisseurs) {
        if ($cardName -and $cardName -like "*$exc*") { $skip = $true; break }
    }
    if ($skip) { $rsPO.MoveNext(); continue }

    $commandes += [PSCustomObject]@{
        DocNum    = $docNum
        DocDate   = $docDate
        CardCode  = $cardCode
        CardName  = $cardName
        Total     = $total
        Devise    = $devise
        Status    = $status
        Createur  = $createur
        Envoye    = $false
        DateEnvoi = ""
        Dest      = ""
        Source    = ""
    }
    $rsPO.MoveNext()
}

Write-Log "$($commandes.Count) commande(s) trouvee(s) dans SAP"

# ── Vérifier dans Outlook (envoyés + réception) ─────────────────────────────
Write-Log "Verification dans Outlook..."
try {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch {
    try {
        $outlook = New-Object -ComObject Outlook.Application
    } catch {
        Write-Log "ATTENTION: Outlook non disponible - colonne envoi sera vide"
        $outlook = $null
    }
}

if ($outlook) {
    $ns = $outlook.GetNamespace("MAPI")
    $dateFilter = (Get-Date).AddDays(-$JoursRecherche).ToString("MM/dd/yyyy")

    # Index rapide: scanner tous les courriels UNE SEULE FOIS, extraire les numeros
    $sentFolder  = $ns.GetDefaultFolder(5)  # olFolderSentMail
    $inboxFolder = $ns.GetDefaultFolder(6)  # olFolderInbox

    # Set des numeros de PO a chercher (pour filtrer vite)
    $poSet = @{}
    foreach ($cmd in $commandes) { $poSet[[string]$cmd.DocNum] = $true }

    # Hash: numero PO -> info du premier courriel trouve
    $emailIndex = @{}

    # Regex pour extraire tous les nombres de 7+ chiffres (numeros de PO SAP)
    $numRegex = [regex]'\d{7,}'

    # Scanner les envoyes
    Write-Log "  Scan des envoyes (sujet + corps)..."
    $sentItems = $sentFolder.Items
    $sentItems.Sort("[SentOn]", $true)
    $sentItems = $sentItems.Restrict("[SentOn] >= '$dateFilter'")
    $scanCount = 0
    $item = $sentItems.GetFirst()
    while ($item -ne $null) {
        $scanCount++
        if ($scanCount % 200 -eq 0) { Write-Log "    $scanCount courriels scannes..." }
        try {
            $text = ""
            try { $text = [string]$item.Subject } catch {}
            try { $text += " " + [string]$item.Body } catch {}
            try { $text += " " + [string]$item.HTMLBody } catch {}
            $matches = $numRegex.Matches($text)
            foreach ($m in $matches) {
                $num = $m.Value
                if ($poSet.ContainsKey($num) -and -not $emailIndex.ContainsKey($num)) {
                    $eDate = ""; try { $eDate = $item.SentOn.ToString("yyyy-MM-dd HH:mm") } catch {}
                    $eDest = ""; try { $eDest = [string]$item.To } catch {}
                    $emailIndex[$num] = @{ Source = "Envoye"; Date = $eDate; Dest = $eDest }
                }
            }
        } catch {}
        $prev = $item
        $item = $sentItems.GetNext()
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($prev) | Out-Null } catch {}
    }
    Write-Log "    $scanCount envoye(s) scannes, $($emailIndex.Count) trace(s)"

    # Scanner la reception
    Write-Log "  Scan de la reception (sujet + corps)..."
    $inboxItems = $inboxFolder.Items
    $inboxItems.Sort("[ReceivedTime]", $true)
    $inboxItems = $inboxItems.Restrict("[ReceivedTime] >= '$dateFilter'")
    $scanCount = 0
    $item = $inboxItems.GetFirst()
    while ($item -ne $null) {
        $scanCount++
        if ($scanCount % 200 -eq 0) { Write-Log "    $scanCount courriels scannes..." }
        try {
            $text = ""
            try { $text = [string]$item.Subject } catch {}
            try { $text += " " + [string]$item.Body } catch {}
            try { $text += " " + [string]$item.HTMLBody } catch {}
            $matches = $numRegex.Matches($text)
            foreach ($m in $matches) {
                $num = $m.Value
                if ($poSet.ContainsKey($num) -and -not $emailIndex.ContainsKey($num)) {
                    $eDate = ""; try { $eDate = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm") } catch {}
                    $eDest = ""; try { $eDest = [string]$item.SenderName } catch {}
                    $emailIndex[$num] = @{ Source = "Recu"; Date = $eDate; Dest = $eDest }
                }
            }
        } catch {}
        $prev = $item
        $item = $inboxItems.GetNext()
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($prev) | Out-Null } catch {}
    }
    Write-Log "    $scanCount recu(s) scannes, $($emailIndex.Count) trace(s) au total"

    # Appliquer l'index aux commandes
    $found = 0
    foreach ($cmd in $commandes) {
        $numStr = [string]$cmd.DocNum
        if ($emailIndex.ContainsKey($numStr)) {
            $info = $emailIndex[$numStr]
            $cmd.Envoye    = $true
            $cmd.DateEnvoi = $info.Date
            $cmd.Dest      = $info.Dest
            $cmd.Source    = $info.Source
            $found++
        }
    }
    Write-Log "$found trace(s) trouvee(s) sur $($commandes.Count) commande(s)"

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sentFolder) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($inboxFolder) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns) | Out-Null
}

$nbEnvoye    = ($commandes | Where-Object { $_.Envoye }).Count
$nbNonEnvoye = ($commandes | Where-Object { -not $_.Envoye }).Count
Write-Log "Resultat: $nbEnvoye envoyee(s), $nbNonEnvoye non envoyee(s)"

# ── Déconnecter SAP ──────────────────────────────────────────────────────────
$company.Disconnect()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($company) | Out-Null
Write-Log "SAP deconnecte"

# ── Générer le fichier Excel ─────────────────────────────────────────────────
Write-Log "Generation du fichier Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()
$ws = $wb.Sheets.Item(1)
$ws.Name = "Recherche Commandes"

# ── En-tête ──────────────────────────────────────────────────────────────────
$ws.Cells.Item(1,1).Value2 = "RECHERCHE DE COMMANDES D'ACHAT"
$ws.Cells.Item(1,1).Font.Size = 14
$ws.Cells.Item(1,1).Font.Bold = $true
$ws.Range("A1:I1").Merge()

$dateGen = (Get-Date).ToString("yyyy-MM-dd HH:mm")
$ws.Cells.Item(2,1).Value2 = "Genere le $dateGen - Periode: $JoursRecherche derniers jours"
$ws.Cells.Item(2,1).Font.Size = 9
$ws.Cells.Item(2,1).Font.Color = 8421504
$ws.Range("A2:I2").Merge()

# Résumé
$ws.Cells.Item(3,1).Value2 = "Total: $($commandes.Count)"
$ws.Cells.Item(3,3).Value2 = "Avec trace: $nbEnvoye"
$ws.Cells.Item(3,3).Font.Color = 32768  # vert
$ws.Cells.Item(3,5).Value2 = "Sans trace: $nbNonEnvoye"
$ws.Cells.Item(3,5).Font.Color = 255    # rouge

# ── En-têtes colonnes ────────────────────────────────────────────────────────
$headers = @("# Commande", "Date", "Code Fourn.", "Fournisseur", "Total", "Devise", "Statut", "Createur", "Trace?", "Source", "Date trace", "De/A", "Exclure")
$row = 5
for ($c = 0; $c -lt $headers.Count; $c++) {
    $cell = $ws.Cells.Item($row, $c + 1)
    $cell.Value2 = $headers[$c]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 15395562  # gris clair
    $cell.Borders.Item(9).LineStyle = 1  # xlEdgeBottom
}

# ── Données ──────────────────────────────────────────────────────────────────
$dataRow = $row + 1
foreach ($cmd in $commandes) {
    $ws.Cells.Item($dataRow, 1).Value2 = [string]$cmd.DocNum
    $ws.Cells.Item($dataRow, 2).Value2 = if ($cmd.DocDate) { try { ([datetime]$cmd.DocDate).ToString("yyyy-MM-dd") } catch { [string]$cmd.DocDate } } else { "" }
    $ws.Cells.Item($dataRow, 3).Value2 = [string]$cmd.CardCode
    $ws.Cells.Item($dataRow, 4).Value2 = [string]$cmd.CardName
    $ws.Cells.Item($dataRow, 5).Value2 = [double]$cmd.Total
    $ws.Cells.Item($dataRow, 5).NumberFormat = "#,##0.00"
    $ws.Cells.Item($dataRow, 6).Value2 = [string]$cmd.Devise
    $ws.Cells.Item($dataRow, 7).Value2 = switch ([string]$cmd.Status) { "O" { "Ouvert" } "C" { "Cloture" } default { [string]$cmd.Status } }
    $ws.Cells.Item($dataRow, 8).Value2 = [string]$cmd.Createur

    if ($cmd.Envoye) {
        $ws.Cells.Item($dataRow, 9).Value2 = "Oui"
        $ws.Cells.Item($dataRow, 9).Font.Color = 32768
    } else {
        $ws.Cells.Item($dataRow, 9).Value2 = "NON"
        $ws.Cells.Item($dataRow, 9).Font.Color = 255
        $ws.Cells.Item($dataRow, 9).Font.Bold = $true
    }
    $ws.Cells.Item($dataRow, 10).Value2 = [string]$cmd.Source
    $ws.Cells.Item($dataRow, 11).Value2 = [string]$cmd.DateEnvoi
    $ws.Cells.Item($dataRow, 12).Value2 = [string]$cmd.Dest

    $dataRow++
}

# ── Mise en forme ────────────────────────────────────────────────────────────
$lastRow = $dataRow - 1
if ($lastRow -ge $row + 1) {
    # Filtre auto
    $ws.Range("A${row}:M${lastRow}").AutoFilter() | Out-Null

    # Mise en forme conditionnelle - lignes sans trace en rose
    $dataRange = $ws.Range("A$($row+1):M${lastRow}")
    $fc = $dataRange.FormatConditions.Add(2, 0, '=$I' + ($row+1) + '="NON"')  # xlExpression=2
    $fc.Interior.Color = 13421823  # rose pale
    $fc.StopIfTrue = $false

    # Dropdown "OUI" dans la colonne Exclure
    $exclRange = $ws.Range("M$($row+1):M${lastRow}")
    $exclRange.Validation.Delete()
    $exclRange.Validation.Add(3, 1, 1, "OUI")  # xlValidateList=3
    $exclRange.Validation.InCellDropdown = $true
    $exclRange.Validation.IgnoreBlank = $true
}

# Largeur des colonnes
$ws.Columns.Item(1).ColumnWidth = 14   # DocNum
$ws.Columns.Item(2).ColumnWidth = 12   # Date
$ws.Columns.Item(3).ColumnWidth = 12   # CardCode
$ws.Columns.Item(4).ColumnWidth = 30   # CardName
$ws.Columns.Item(5).ColumnWidth = 14   # Total
$ws.Columns.Item(6).ColumnWidth = 8    # Devise
$ws.Columns.Item(7).ColumnWidth = 10   # Statut
$ws.Columns.Item(8).ColumnWidth = 22   # Createur
$ws.Columns.Item(9).ColumnWidth = 10   # Trace?
$ws.Columns.Item(10).ColumnWidth = 10  # Source
$ws.Columns.Item(11).ColumnWidth = 18  # Date trace
$ws.Columns.Item(12).ColumnWidth = 30  # De/A
$ws.Columns.Item(13).ColumnWidth = 10  # Exclure

# ── Feuille Parametres ──────────────────────────────────────────────────────
$wsParam = $wb.Sheets.Add([System.Reflection.Missing]::Value, $ws)
$wsParam.Name = "Parametres"
$wsParam.Cells.Item(1,1).Value2 = "PARAMETRES - RECHERCHE COMMANDES"
$wsParam.Cells.Item(1,1).Font.Size = 14
$wsParam.Cells.Item(1,1).Font.Bold = $true
$wsParam.Range("A1:C1").Merge()

$wsParam.Cells.Item(3,1).Value2 = "Createurs exclus (un par ligne):"
$wsParam.Cells.Item(3,1).Font.Bold = $true
$pRow = 4
foreach ($nom in $ExclureCreateurs) {
    $wsParam.Cells.Item($pRow, 1).Value2 = $nom
    $pRow++
}

$wsParam.Cells.Item(3,3).Value2 = "Commandes exclues:"
$wsParam.Cells.Item(3,3).Font.Bold = $true
$pRow = 4
foreach ($po in $ExclurePOs) {
    $wsParam.Cells.Item($pRow, 3).Value2 = [string]$po
    $pRow++
}

$wsParam.Cells.Item(2,1).Value2 = "Modifiez les listes ci-dessous puis cliquez sur [Sauvegarder exclusions]"
$wsParam.Cells.Item(2,1).Font.Size = 9
$wsParam.Cells.Item(2,1).Font.Color = 8421504
$wsParam.Range("A2:D2").Merge()

$wsParam.Columns.Item(1).ColumnWidth = 30
$wsParam.Columns.Item(2).ColumnWidth = 5
$wsParam.Columns.Item(3).ColumnWidth = 20

# ── Macro VBA pour sauvegarder les exclusions ───────────────────────────────
$scriptDir = Split-Path $PSCommandPath
$vbaCode = @"
Sub SauvegarderExclusions()
    Dim wsRC As Worksheet, wsP As Worksheet
    Dim json As String, i As Long, v As String
    Dim poList As String, crList As String
    Dim lastRow As Long
    Dim filePath As String

    filePath = "$($scriptDir.Replace('\','\\'))\\RechercheCommandes_exclusions.json"

    Set wsRC = ThisWorkbook.Sheets("Recherche Commandes")
    Set wsP = ThisWorkbook.Sheets("Parametres")

    ' Collecter les PO exclus depuis la colonne M + feuille Parametres
    poList = ""
    lastRow = wsRC.Cells(wsRC.Rows.Count, 1).End(xlUp).Row
    For i = 6 To lastRow
        If UCase(Trim(wsRC.Cells(i, 13).Value)) = "OUI" Then
            v = Trim(wsRC.Cells(i, 1).Value)
            If Len(v) > 0 Then
                If Len(poList) > 0 Then poList = poList & ","
                poList = poList & """" & v & """"
            End If
        End If
    Next i
    ' Ajouter ceux de la feuille Parametres
    lastRow = wsP.Cells(wsP.Rows.Count, 3).End(xlUp).Row
    For i = 4 To lastRow
        v = Trim(wsP.Cells(i, 3).Value)
        If Len(v) > 0 Then
            If Len(poList) > 0 Then poList = poList & ","
            poList = poList & """" & v & """"
        End If
    Next i

    ' Collecter les createurs exclus
    crList = ""
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    For i = 4 To lastRow
        v = Trim(wsP.Cells(i, 1).Value)
        If Len(v) > 0 Then
            If Len(crList) > 0 Then crList = crList & ","
            crList = crList & """" & v & """"
        End If
    Next i

    json = "{" & vbCrLf
    json = json & "  ""po_exclus"": [" & poList & "]," & vbCrLf
    json = json & "  ""createurs_exclus"": [" & crList & "]" & vbCrLf
    json = json & "}"

    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, json
    Close #f

    MsgBox "Exclusions sauvegardees!" & vbCrLf & vbCrLf & _
           "PO exclus: " & IIf(Len(poList) > 0, poList, "(aucun)") & vbCrLf & _
           "Createurs exclus: " & IIf(Len(crList) > 0, crList, "(aucun)"), _
           vbInformation, "Recherche Commandes"
End Sub
"@

try {
    $vbModule = $wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
    $vbModule.CodeModule.AddFromString($vbaCode)
} catch {
    Write-Log "ATTENTION: Impossible d'ajouter la macro VBA (verifier les parametres de securite Excel)"
}

# Bouton sur la feuille principale
$ws.Activate()
try {
    $btn = $ws.Buttons.Add(
        $ws.Columns.Item(9).Left,   # colonne I
        $ws.Rows.Item(3).Top,       # ligne 3
        180, 22
    )
    $btn.Caption = "Sauvegarder exclusions"
    $btn.OnAction = "SauvegarderExclusions"
    $btn.Font.Size = 9
} catch {
    Write-Log "ATTENTION: Impossible d'ajouter le bouton macro"
}

# Figer les volets
$ws.Cells.Item($row + 1, 1).Select()
$excel.ActiveWindow.FreezePanes = $true

# ── Afficher ─────────────────────────────────────────────────────────────────
$excel.Visible = $true
Write-Log "Fichier Excel genere et ouvert. $($commandes.Count) commande(s)."
if ($ExclurePOs.Count -gt 0) { Write-Log "  $($ExclurePOs.Count) commande(s) exclue(s) par les exclusions sauvegardees" }
Write-Log "Termine!"
