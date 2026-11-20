<#
  GestionPrix.ps1 — Outil local de gestion des prix SAP
  Connexion directe via DI API (zero latence)
  Usage: clic-droit > Executer avec PowerShell
         ou: powershell -ExecutionPolicy Bypass -File "U:\GromecOutlook\GestionPrix.ps1"
#>

$ErrorActionPreference = "Stop"
$configPath = "U:\GromecOutlook\config.json"

# ── Charger config ──
if (-not (Test-Path $configPath)) {
    [System.Windows.Forms.MessageBox]::Show("config.json introuvable: $configPath", "Erreur", "OK", "Error")
    exit 1
}
$config = Get-Content $configPath -Raw | ConvertFrom-Json

# ── Charger DI API ──
$dllPaths = @(
    "C:\Program Files\SAP\SAP Business One DI API\Interop.SAPbobsCOM.dll",
    "C:\Program Files (x86)\SAP\SAP Business One DI API\Interop.SAPbobsCOM.dll",
    "C:\Program Files\sap\Data Transfer Workbench\Interop.SAPbobsCOM.dll"
)
$dllLoaded = $false
foreach ($dll in $dllPaths) {
    if (Test-Path $dll) {
        [System.Reflection.Assembly]::LoadFile($dll) | Out-Null
        $dllLoaded = $true
        break
    }
}
if (-not $dllLoaded) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("DI API DLL introuvable", "Erreur", "OK", "Error")
    exit 1
}

# ── Connexion SAP ──
function Connect-SAPPrix {
    $company = New-Object -ComObject SAPbobsCOM.Company
    $company.Server = "NDB@hana:30013"
    $company.DbServerType = 9
    $company.CompanyDB = "SAP_GROMEC_PROD"
    $company.UseTrusted = $false
    $company.UserName = $config.sap_user
    $company.Password = $config.sap_password
    $ret = $company.Connect()
    if ($ret -ne 0) {
        $err = $company.GetLastErrorDescription()
        throw "Connexion SAP echouee ($ret): $err"
    }
    return $company
}

function Run-SAPQuery {
    param($company, [string]$sql)
    $rs = $company.GetBusinessObject(300)
    [System.__ComObject].InvokeMember("DoQuery", "InvokeMethod", $null, $rs, @($sql))
    return $rs
}

function Get-RSField {
    param($rs, [string]$fieldName)
    try {
        $fields = [System.__ComObject].InvokeMember("Fields", "InvokeMethod", $null, $rs, $null)
        $field = [System.__ComObject].InvokeMember("Item", "InvokeMethod", $null, $fields, @($fieldName))
        return [System.__ComObject].InvokeMember("Value", "GetProperty", $null, $field, $null)
    } catch { return "" }
}

function Get-RSCount {
    param($rs)
    return [int][System.__ComObject].InvokeMember("RecordCount", "GetProperty", $null, $rs, $null)
}

function Move-RSNext {
    param($rs)
    [System.__ComObject].InvokeMember("MoveNext", "InvokeMethod", $null, $rs, $null)
}

# ── Interface Excel ──
Add-Type -AssemblyName System.Windows.Forms

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false

$wb = $xl.Workbooks.Add()

# Supprimer les feuilles en trop, garder 1
while ($wb.Sheets.Count -gt 1) { $wb.Sheets($wb.Sheets.Count).Delete() }

$wsClient = $wb.Sheets(1)
$wsClient.Name = "Par client"

$wsArticle = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsClient)
$wsArticle.Name = "Par article"

$wsListes = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsArticle)
$wsListes.Name = "Listes de prix"

# ── Style headers ──
function Set-Header {
    param($ws, [int]$row, [string[]]$headers)
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $ws.Cells($row, $i + 1)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 0x2E231B  # RGB(27,35,46) en BGR
        $cell.Font.Color = 0xAE9B8C      # RGB(140,155,174) en BGR
    }
}

# ── Onglet Par client ──
$wsClient.Cells(1,1).Value2 = "Code client:"
$wsClient.Cells(1,2).Value2 = ""
$wsClient.Cells(2,1).Value2 = "Nom:"
$wsClient.Cells(2,2).Value2 = ""
$wsClient.Cells(1,1).Font.Bold = $true
$wsClient.Cells(2,1).Font.Bold = $true
Set-Header $wsClient 4 @("Article","Description","Prix","Devise","Escompte %","Liste #","Auto MAJ","Valide","Debut","Fin")
$wsClient.Columns("A:A").ColumnWidth = 16
$wsClient.Columns("B:B").ColumnWidth = 35
$wsClient.Columns("C:C").ColumnWidth = 12
$wsClient.Columns("D:D").ColumnWidth = 8
$wsClient.Columns("E:E").ColumnWidth = 12
$wsClient.Columns("F:F").ColumnWidth = 8
$wsClient.Columns("G:G").ColumnWidth = 10
$wsClient.Columns("H:H").ColumnWidth = 8
$wsClient.Columns("I:J").ColumnWidth = 12

# ── Onglet Par article ──
$wsArticle.Cells(1,1).Value2 = "Code article:"
$wsArticle.Cells(1,2).Value2 = ""
$wsArticle.Cells(2,1).Value2 = "Description:"
$wsArticle.Cells(2,2).Value2 = ""
$wsArticle.Cells(1,1).Font.Bold = $true
$wsArticle.Cells(2,1).Font.Bold = $true
$wsArticle.Cells(3,1).Value2 = "PRIX PAR LISTE"
$wsArticle.Cells(3,1).Font.Bold = $true
Set-Header $wsArticle 4 @("#","Liste","Prix","Devise","Base #","Facteur","Brut")
$wsArticle.Columns("A:A").ColumnWidth = 6
$wsArticle.Columns("B:B").ColumnWidth = 28
$wsArticle.Columns("C:C").ColumnWidth = 12
$wsArticle.Columns("D:D").ColumnWidth = 8
$wsArticle.Columns("E:E").ColumnWidth = 8
$wsArticle.Columns("F:F").ColumnWidth = 10
$wsArticle.Columns("G:G").ColumnWidth = 6

# ── Onglet Listes de prix ──
Set-Header $wsListes 1 @("#","Nom","Base #","Facteur","Brut","Devise","Valide","Debut","Fin")
$wsListes.Columns("A:A").ColumnWidth = 6
$wsListes.Columns("B:B").ColumnWidth = 30
$wsListes.Columns("C:C").ColumnWidth = 8
$wsListes.Columns("D:D").ColumnWidth = 10
$wsListes.Columns("E:E").ColumnWidth = 6
$wsListes.Columns("F:F").ColumnWidth = 8
$wsListes.Columns("G:G").ColumnWidth = 8
$wsListes.Columns("H:I").ColumnWidth = 12

$wsClient.Activate()

# ── Connexion SAP ──
Write-Host "Connexion SAP..."
try {
    $company = Connect-SAPPrix
    Write-Host "Connecte a SAP OK"
} catch {
    [System.Windows.Forms.MessageBox]::Show("Erreur connexion SAP: $($_.Exception.Message)", "Erreur", "OK", "Error")
    exit 1
}

# ── Menu interactif dans la console ──
Write-Host ""
Write-Host "============================================"
Write-Host "  GESTION PRIX SAP — Commandes disponibles"
Write-Host "============================================"
Write-Host "  client <code>    — Chercher prix speciaux d'un client"
Write-Host "  article <code>   — Chercher prix d'un article"
Write-Host "  listes           — Charger toutes les listes de prix"
Write-Host "  save-client      — Sauvegarder modifications client"
Write-Host "  save-article     — Sauvegarder modifications article"
Write-Host "  refresh          — Reconnecter SAP"
Write-Host "  quit             — Quitter"
Write-Host "============================================"
Write-Host ""

while ($true) {
    $input = Read-Host "GestionPrix>"
    if ([string]::IsNullOrWhiteSpace($input)) { continue }

    $parts = $input.Trim() -split "\s+", 2
    $cmd = $parts[0].ToLower()
    $arg = if ($parts.Count -gt 1) { $parts[1] } else { "" }

    switch ($cmd) {
        "client" {
            if ($arg -eq "") { Write-Host "Usage: client <CodeClient>"; continue }
            try {
                $wsClient.Cells(1,2).Value2 = $arg
                Write-Host "Recherche client $arg..."

                $rsInfo = Run-SAPQuery $company "SELECT ""CardCode"",""CardName"" FROM ""OCRD"" WHERE ""CardCode"" = '$($arg -replace "'","''")'"
                if ((Get-RSCount $rsInfo) -eq 0) {
                    Write-Host "Client introuvable: $arg"
                    $wsClient.Cells(2,2).Value2 = "INTROUVABLE"
                    continue
                }
                $wsClient.Cells(2,2).Value2 = Get-RSField $rsInfo "CardName"
                Write-Host "Client: $(Get-RSField $rsInfo 'CardName')"

                # Effacer anciennes donnees
                $lastRow = $wsClient.Cells($wsClient.Rows.Count, 1).End(-4162).Row  # xlUp
                if ($lastRow -gt 4) {
                    $wsClient.Range($wsClient.Cells(5,1), $wsClient.Cells($lastRow, 10)).ClearContents()
                }

                $rsSP = Run-SAPQuery $company "SELECT s.""ItemCode"", i.""ItemName"", s.""Price"", s.""Currency"", s.""Discount"", s.""ListNum"", s.""AutoUpdt"", s.""Valid"", s.""ValidFrom"", s.""ValidTo"" FROM ""OSPP"" s LEFT JOIN ""OITM"" i ON s.""ItemCode"" = i.""ItemCode"" WHERE s.""CardCode"" = '$($arg -replace "'","''")' ORDER BY s.""ItemCode"""
                $cnt = Get-RSCount $rsSP
                Write-Host "$cnt prix speciaux trouves"

                for ($r = 0; $r -lt $cnt -and $r -lt 5000; $r++) {
                    $row = 5 + $r
                    $wsClient.Cells($row, 1).Value2 = Get-RSField $rsSP "ItemCode"
                    $wsClient.Cells($row, 2).Value2 = Get-RSField $rsSP "ItemName"
                    $wsClient.Cells($row, 3).Value2 = Get-RSField $rsSP "Price"
                    $wsClient.Cells($row, 4).Value2 = Get-RSField $rsSP "Currency"
                    $wsClient.Cells($row, 5).Value2 = Get-RSField $rsSP "Discount"
                    $wsClient.Cells($row, 6).Value2 = Get-RSField $rsSP "ListNum"
                    $wsClient.Cells($row, 7).Value2 = Get-RSField $rsSP "AutoUpdt"
                    $wsClient.Cells($row, 8).Value2 = Get-RSField $rsSP "Valid"
                    $wsClient.Cells($row, 9).Value2 = Get-RSField $rsSP "ValidFrom"
                    $wsClient.Cells($row, 10).Value2 = Get-RSField $rsSP "ValidTo"
                    Move-RSNext $rsSP
                }
                Write-Host "OK — $cnt lignes chargees dans 'Par client'"
            } catch {
                Write-Host "ERREUR: $($_.Exception.Message)"
            }
        }

        "article" {
            if ($arg -eq "") { Write-Host "Usage: article <CodeArticle>"; continue }
            try {
                $wsArticle.Cells(1,2).Value2 = $arg
                Write-Host "Recherche article $arg..."

                $rsInfo = Run-SAPQuery $company "SELECT ""ItemCode"",""ItemName"" FROM ""OITM"" WHERE ""ItemCode"" = '$($arg -replace "'","''")'"
                if ((Get-RSCount $rsInfo) -eq 0) {
                    Write-Host "Article introuvable: $arg"
                    $wsArticle.Cells(2,2).Value2 = "INTROUVABLE"
                    continue
                }
                $wsArticle.Cells(2,2).Value2 = Get-RSField $rsInfo "ItemName"
                Write-Host "Article: $(Get-RSField $rsInfo 'ItemName')"

                $lastRow = $wsArticle.Cells($wsArticle.Rows.Count, 1).End(-4162).Row
                if ($lastRow -gt 4) {
                    $wsArticle.Range($wsArticle.Cells(5,1), $wsArticle.Cells($lastRow, 7)).ClearContents()
                }

                $rsPL = Run-SAPQuery $company "SELECT p.""PriceList"", l.""ListName"", p.""Price"", p.""Currency"", l.""BASE_NUM"", l.""Factor"", l.""IsGrossPrc"" FROM ""ITM1"" p LEFT JOIN ""OPLN"" l ON p.""PriceList"" = l.""ListNum"" WHERE p.""ItemCode"" = '$($arg -replace "'","''")' ORDER BY p.""PriceList"""
                $cnt = Get-RSCount $rsPL

                for ($r = 0; $r -lt $cnt; $r++) {
                    $row = 5 + $r
                    $wsArticle.Cells($row, 1).Value2 = Get-RSField $rsPL "PriceList"
                    $wsArticle.Cells($row, 2).Value2 = Get-RSField $rsPL "ListName"
                    $wsArticle.Cells($row, 3).Value2 = Get-RSField $rsPL "Price"
                    $wsArticle.Cells($row, 4).Value2 = Get-RSField $rsPL "Currency"
                    $wsArticle.Cells($row, 5).Value2 = Get-RSField $rsPL "BASE_NUM"
                    $wsArticle.Cells($row, 6).Value2 = Get-RSField $rsPL "Factor"
                    $wsArticle.Cells($row, 7).Value2 = Get-RSField $rsPL "IsGrossPrc"
                    Move-RSNext $rsPL
                }
                Write-Host "OK — $cnt listes de prix chargees dans 'Par article'"
            } catch {
                Write-Host "ERREUR: $($_.Exception.Message)"
            }
        }

        "listes" {
            try {
                Write-Host "Chargement listes de prix..."
                $wsListes.Activate()

                $lastRow = $wsListes.Cells($wsListes.Rows.Count, 1).End(-4162).Row
                if ($lastRow -gt 1) {
                    $wsListes.Range($wsListes.Cells(2,1), $wsListes.Cells($lastRow, 9)).ClearContents()
                }

                $rsL = Run-SAPQuery $company "SELECT ""ListNum"",""ListName"",""BASE_NUM"",""Factor"",""IsGrossPrc"",""PrimCurr"",""ValidFor"",""ValidFrom"",""ValidTo"" FROM ""OPLN"" ORDER BY ""ListNum"""
                $cnt = Get-RSCount $rsL

                for ($r = 0; $r -lt $cnt; $r++) {
                    $row = 2 + $r
                    $wsListes.Cells($row, 1).Value2 = Get-RSField $rsL "ListNum"
                    $wsListes.Cells($row, 2).Value2 = Get-RSField $rsL "ListName"
                    $wsListes.Cells($row, 3).Value2 = Get-RSField $rsL "BASE_NUM"
                    $wsListes.Cells($row, 4).Value2 = Get-RSField $rsL "Factor"
                    $wsListes.Cells($row, 5).Value2 = Get-RSField $rsL "IsGrossPrc"
                    $wsListes.Cells($row, 6).Value2 = Get-RSField $rsL "PrimCurr"
                    $wsListes.Cells($row, 7).Value2 = Get-RSField $rsL "ValidFor"
                    $wsListes.Cells($row, 8).Value2 = Get-RSField $rsL "ValidFrom"
                    $wsListes.Cells($row, 9).Value2 = Get-RSField $rsL "ValidTo"
                    Move-RSNext $rsL
                }
                Write-Host "OK — $cnt listes chargees"
            } catch {
                Write-Host "ERREUR: $($_.Exception.Message)"
            }
        }

        "save-client" {
            try {
                $code = $wsClient.Cells(1,2).Value2
                if ([string]::IsNullOrWhiteSpace($code)) { Write-Host "Pas de client charge"; continue }
                Write-Host "Sauvegarde prix client $code..."
                $cnt = 0
                $r = 5
                while ($wsClient.Cells($r, 1).Value2 -ne "" -and $wsClient.Cells($r, 1).Value2 -ne $null) {
                    $itemCode = $wsClient.Cells($r, 1).Value2
                    $price = $wsClient.Cells($r, 3).Value2
                    $discount = $wsClient.Cells($r, 5).Value2
                    $listNum = $wsClient.Cells($r, 6).Value2
                    $sql = "UPDATE ""OSPP"" SET ""Price"" = $price, ""Discount"" = $discount, ""ListNum"" = $listNum"
                    $validTo = $wsClient.Cells($r, 10).Value2
                    if ($validTo -ne "" -and $validTo -ne $null) {
                        $sql += ", ""ValidTo"" = '$([DateTime]$validTo | Get-Date -Format 'yyyy-MM-dd')'"
                    }
                    $sql += " WHERE ""CardCode"" = '$($code -replace "'","''")' AND ""ItemCode"" = '$($itemCode -replace "'","''")'"
                    Run-SAPQuery $company $sql
                    $cnt++
                    $r++
                }
                Write-Host "OK — $cnt prix mis a jour"
            } catch {
                Write-Host "ERREUR: $($_.Exception.Message)"
            }
        }

        "save-article" {
            try {
                $code = $wsArticle.Cells(1,2).Value2
                if ([string]::IsNullOrWhiteSpace($code)) { Write-Host "Pas d'article charge"; continue }
                Write-Host "Sauvegarde prix article $code..."
                $cnt = 0
                $r = 5
                while ($wsArticle.Cells($r, 1).Value2 -ne "" -and $wsArticle.Cells($r, 1).Value2 -ne $null) {
                    $listNum = $wsArticle.Cells($r, 1).Value2
                    $price = $wsArticle.Cells($r, 3).Value2
                    $sql = "UPDATE ""ITM1"" SET ""Price"" = $price WHERE ""ItemCode"" = '$($code -replace "'","''")' AND ""PriceList"" = $listNum"
                    Run-SAPQuery $company $sql
                    $cnt++
                    $r++
                }
                Write-Host "OK — $cnt prix de liste mis a jour"
            } catch {
                Write-Host "ERREUR: $($_.Exception.Message)"
            }
        }

        "refresh" {
            try {
                $company.Disconnect()
                $company = Connect-SAPPrix
                Write-Host "Reconnecte a SAP OK"
            } catch {
                Write-Host "ERREUR reconnexion: $($_.Exception.Message)"
            }
        }

        "quit" {
            try { $company.Disconnect() } catch {}
            Write-Host "Deconnecte. Fermeture..."
            break
        }

        default {
            Write-Host "Commande inconnue: $cmd"
            Write-Host "Commandes: client, article, listes, save-client, save-article, refresh, quit"
        }
    }
}
