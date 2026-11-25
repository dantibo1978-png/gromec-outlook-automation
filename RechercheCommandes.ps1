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
$ExclureStatuts      = @("C")   # C = Clôturé

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
<html><head><title>Connexion SAP — Recherche Commandes</title>
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
Write-Log "Requete OPOR — commandes depuis $dateDebut..."

$sqlPO = @"
SELECT T0."DocNum", T0."DocDate", T0."CardCode", T0."CardName",
       T0."DocTotal", T0."DocCur", T0."DocStatus",
       IFNULL(T0."U_CreePar", '') AS "Createur"
FROM "OPOR" T0
WHERE T0."DocDate" >= '$dateDebut'
ORDER BY T0."DocNum" DESC
"@

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

    # Exclure clôturées
    if ($ExclureStatuts -contains $status) { $rsPO.MoveNext(); continue }

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
    }
    $rsPO.MoveNext()
}

Write-Log "$($commandes.Count) commande(s) trouvee(s) dans SAP"

# ── Vérifier dans Outlook les éléments envoyés ───────────────────────────────
Write-Log "Verification dans Outlook (elements envoyes)..."
try {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch {
    try {
        $outlook = New-Object -ComObject Outlook.Application
    } catch {
        Write-Log "ATTENTION: Outlook non disponible — colonne envoi sera vide"
        $outlook = $null
    }
}

if ($outlook) {
    $ns = $outlook.GetNamespace("MAPI")
    $sentFolder = $ns.GetDefaultFolder(5)  # olFolderSentMail

    $dateFilter = (Get-Date).AddDays(-$JoursRecherche).ToString("MM/dd/yyyy")
    $filter = "[SentOn] >= '$dateFilter'"
    $sentItems = $sentFolder.Items.Restrict($filter)

    # Construire un index numéro -> { date, dest, sujet }
    $indexEnvois = @{}
    $count = 0
    foreach ($item in $sentItems) {
        try {
            $sujet = $item.Subject
            if (-not $sujet) { continue }
            # Extraire un numéro de 4+ chiffres du sujet
            if ($sujet -match '(\d{4,})') {
                $num = $Matches[1]
                if (-not $indexEnvois.ContainsKey($num)) {
                    $indexEnvois[$num] = @{
                        Date = $item.SentOn.ToString("yyyy-MM-dd HH:mm")
                        Dest = $item.To
                    }
                }
            }
            $count++
        } catch {}
    }
    Write-Log "$count element(s) envoye(s) scannes, $($indexEnvois.Count) numero(s) uniques trouves"

    # Croiser avec les commandes
    foreach ($cmd in $commandes) {
        $numStr = [string]$cmd.DocNum
        if ($indexEnvois.ContainsKey($numStr)) {
            $cmd.Envoye    = $true
            $cmd.DateEnvoi = $indexEnvois[$numStr].Date
            $cmd.Dest      = $indexEnvois[$numStr].Dest
        }
    }

    # Libérer COM
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sentItems) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sentFolder) | Out-Null
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
$ws.Cells.Item(2,1).Value2 = "Genere le $dateGen — Periode: $JoursRecherche derniers jours"
$ws.Cells.Item(2,1).Font.Size = 9
$ws.Cells.Item(2,1).Font.Color = 8421504
$ws.Range("A2:I2").Merge()

# Résumé
$ws.Cells.Item(3,1).Value2 = "Total: $($commandes.Count)"
$ws.Cells.Item(3,3).Value2 = "Envoyees: $nbEnvoye"
$ws.Cells.Item(3,3).Font.Color = 32768  # vert
$ws.Cells.Item(3,5).Value2 = "Non envoyees: $nbNonEnvoye"
$ws.Cells.Item(3,5).Font.Color = 255    # rouge

# ── En-têtes colonnes ────────────────────────────────────────────────────────
$headers = @("# Commande", "Date", "Code Fourn.", "Fournisseur", "Total", "Devise", "Statut", "Createur", "Envoye?", "Date envoi", "Destinataire")
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
    $ws.Cells.Item($dataRow, 1).Value2 = $cmd.DocNum
    $ws.Cells.Item($dataRow, 2).Value2 = if ($cmd.DocDate) { ([datetime]$cmd.DocDate).ToString("yyyy-MM-dd") } else { "" }
    $ws.Cells.Item($dataRow, 3).Value2 = $cmd.CardCode
    $ws.Cells.Item($dataRow, 4).Value2 = $cmd.CardName
    $ws.Cells.Item($dataRow, 5).Value2 = $cmd.Total
    $ws.Cells.Item($dataRow, 5).NumberFormat = "#,##0.00"
    $ws.Cells.Item($dataRow, 6).Value2 = $cmd.Devise
    $ws.Cells.Item($dataRow, 7).Value2 = switch ($cmd.Status) { "O" { "Ouvert" } "C" { "Cloture" } default { $cmd.Status } }
    $ws.Cells.Item($dataRow, 8).Value2 = $cmd.Createur

    if ($cmd.Envoye) {
        $ws.Cells.Item($dataRow, 9).Value2 = "Oui"
        $ws.Cells.Item($dataRow, 9).Font.Color = 32768
    } else {
        $ws.Cells.Item($dataRow, 9).Value2 = "NON"
        $ws.Cells.Item($dataRow, 9).Font.Color = 255
        $ws.Cells.Item($dataRow, 9).Font.Bold = $true
    }
    $ws.Cells.Item($dataRow, 10).Value2 = $cmd.DateEnvoi
    $ws.Cells.Item($dataRow, 11).Value2 = $cmd.Dest

    $dataRow++
}

# ── Mise en forme ────────────────────────────────────────────────────────────
$lastRow = $dataRow - 1
if ($lastRow -ge $row + 1) {
    # Filtre auto
    $ws.Range("A${row}:K${lastRow}").AutoFilter() | Out-Null

    # Mise en forme conditionnelle — lignes non envoyées en rose
    $dataRange = $ws.Range("A$($row+1):K${lastRow}")
    $fc = $dataRange.FormatConditions.Add(2, 0, '=$I' + ($row+1) + '="NON"')  # xlExpression=2
    $fc.Interior.Color = 13421823  # rose pâle
    $fc.StopIfTrue = $false
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
$ws.Columns.Item(9).ColumnWidth = 10   # Envoyé
$ws.Columns.Item(10).ColumnWidth = 18  # Date envoi
$ws.Columns.Item(11).ColumnWidth = 30  # Dest

# Figer les volets
$ws.Cells.Item($row + 1, 1).Select()
$excel.ActiveWindow.FreezePanes = $true

# ── Afficher ─────────────────────────────────────────────────────────────────
$excel.Visible = $true
Write-Log "Fichier Excel genere et ouvert. $($commandes.Count) commande(s)."
Write-Log "Termine!"
