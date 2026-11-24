<#
  GenererVentesPerte.ps1 - Genere VentesPerte.xlsm
  Detecte les ventes a perte : tout prix de vente SAP < cout moyen magasin 10 (OITW.AvgPrice, WhsCode=10)
  Sources verifiees :
    - Listes de prix          (ITM1)
    - Prix speciaux client     (OSPP)
    - Prix speciaux par dates  (SPP1, non expires)
    - Prix speciaux par qte    (SPP2, periode non expiree)
  Exclus : codes commencant par PT, TA, TB, TC, TD, PR, liste de prix 11 (Coutant moyen), articles inactifs
  Filtre optionnel par groupe d'articles (onglet Groupes) et par liste (onglet Listes)
#>

$xlsmPath = "U:\GromecOutlook\VentesPerte.xlsm"

if (Test-Path $xlsmPath) { Remove-Item $xlsmPath -Force }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false

$wb = $xl.Workbooks.Add()
while ($wb.Sheets.Count -gt 1) { $wb.Sheets($wb.Sheets.Count).Delete() }

# -- Feuilles --
$wsRes = $wb.Sheets(1)
$wsRes.Name = "Ventes a perte"
$wsGrp = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsRes)
$wsGrp.Name = "Groupes"
$wsLst = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsGrp)
$wsLst.Name = "Listes"

function Set-Header($ws, $row, $headers, $color) {
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $c = $ws.Cells($row, $i + 1)
        $c.Value2 = $headers[$i]
        $c.Font.Bold = $true
        $c.Interior.Color = $color
        $c.Font.Color = 0xFFFFFF
    }
}

# ===== ONGLET VENTES A PERTE =====
$wsRes.Cells(1,1).Value2 = "VENTES A PERTE - Prix de vente < Cout moyen magasin 10"
$wsRes.Cells(1,1).Font.Bold = $true
$wsRes.Cells(1,1).Font.Size = 14
$wsRes.Cells(2,1).Value2 = "1) Charger groupes/listes  2) Mettre X dans onglets Groupes/Listes (aucun X = tout)  3) Analyser"
$wsRes.Cells(2,1).Font.Italic = $true
$wsRes.Cells(2,1).Font.Color = 0x808080
$wsRes.Cells(3,1).Value2 = "Lignes orange = devise differente de `$ (cout en devise locale, comparer avec prudence)"
$wsRes.Cells(3,1).Font.Italic = $true
$wsRes.Cells(3,1).Font.Color = 0x808080

$hdrR = @("Source","Groupe","Article","Description","Liste #","Nom liste","Client","Nom client",
          "Date debut","Date fin","Qte min","Devise","Prix vente","Cout mag 10","Perte `$","Marge %")
Set-Header $wsRes 5 $hdrR 0x00008B
$widths = @(14,28,16,40,7,24,12,30,11,11,8,7,12,12,11,9)
for ($i = 0; $i -lt $widths.Count; $i++) { $wsRes.Columns($i + 1).ColumnWidth = $widths[$i] }
$wsRes.Columns("C:C").NumberFormat = "@"
$wsRes.Columns("G:G").NumberFormat = "@"
$wsRes.Columns("I:J").NumberFormat = "yyyy-mm-dd"
$wsRes.Columns("M:O").NumberFormat = "#,##0.00"
$wsRes.Columns("P:P").NumberFormat = "0.0%"
$wsRes.Rows(5).Font.Bold = $true

# Boutons
try {
    $b1 = $wsRes.Buttons().Add($wsRes.Range("J1").Left, 2, 160, 26)
    $b1.Text = "1. Charger groupes/listes"
    $b1.OnAction = "ChargerGroupesEtListes"
    $b2 = $wsRes.Buttons().Add($wsRes.Range("J1").Left + 170, 2, 160, 26)
    $b2.Text = "2. Analyser ventes a perte"
    $b2.OnAction = "AnalyserVentesPerte"
    $wsRes.Rows(1).RowHeight = 30
} catch {
    Write-Host "Boutons non crees (utiliser Alt+F8): $($_.Exception.Message)"
}

# ===== ONGLET GROUPES =====
$wsGrp.Cells(1,1).Value2 = "Mettre X en colonne A pour analyser seulement ces groupes. Aucun X = tous les groupes."
$wsGrp.Cells(1,1).Font.Italic = $true
Set-Header $wsGrp 3 @("Inclure (X)","Code groupe","Nom du groupe","Nb articles") 0x8B4513
$wsGrp.Columns("A:A").ColumnWidth = 11
$wsGrp.Columns("B:B").ColumnWidth = 12
$wsGrp.Columns("C:C").ColumnWidth = 45
$wsGrp.Columns("D:D").ColumnWidth = 11

# ===== ONGLET LISTES =====
$wsLst.Cells(1,1).Value2 = "Mettre X en colonne A pour verifier seulement ces listes de prix. Aucun X = toutes les listes."
$wsLst.Cells(1,1).Font.Italic = $true
Set-Header $wsLst 3 @("Inclure (X)","Liste #","Nom de la liste") 0x003366
$wsLst.Columns("A:A").ColumnWidth = 11
$wsLst.Columns("B:B").ColumnWidth = 8
$wsLst.Columns("C:C").ColumnWidth = 40

$wsRes.Activate()

# ========== INJECTION VBA ==========
Write-Host "Injection du code VBA..."

$vbaCode = @'
Option Explicit

Private mSapUser As String
Private mSapPass As String
Private mErreurs As String
Private Const LOG_PATH As String = "U:\GromecOutlook\VentesPerte.log"
Private Const NB_COLS As Long = 16
Private Const ROW_HDR As Long = 5
Private Const WHS_COUT As String = "10"   ' magasin de reference pour le cout
Private Const LISTES_EXCLUES As String = "11"   ' listes de prix ignorees (11 = Coutant moyen)
Private Const PREFIXES_EXCLUS As String = "'PT','TA','TB','TC','TD','PR'"   ' codes articles exclus (2 premiers caracteres)

Private Sub WriteLog(msg As String)
    On Error Resume Next
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(LOG_PATH, 8, True)
    f.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & mSapUser & "] " & msg
    f.Close
End Sub

Private Function GetSAPLogin() As Boolean
    Dim fso As Object, tempHta As String, resultFile As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempHta = Environ("TEMP") & "\sap_login.hta"
    resultFile = Environ("TEMP") & "\sap_login_result.txt"

    If fso.FileExists(resultFile) Then fso.DeleteFile resultFile

    Dim h As String
    h = "<html><head><title>Connexion SAP</title>" & _
        "<hta:application border=""dialog"" maximizeButton=""no"" minimizeButton=""no"" scroll=""no"" />" & _
        "<style>body{font:12px Tahoma;margin:20px}td{padding:5px}input{width:160px;padding:3px}" & _
        "button{width:85px;padding:5px;cursor:pointer}</style>" & _
        "<script language=""VBScript"">" & vbCrLf & _
        "Sub DoOK()" & vbCrLf & _
        "  Dim f : Set f = CreateObject(""Scripting.FileSystemObject"").CreateTextFile(""" & resultFile & """, True)" & vbCrLf & _
        "  f.WriteLine document.getElementById(""usr"").value" & vbCrLf & _
        "  f.WriteLine document.getElementById(""pwd"").value" & vbCrLf & _
        "  f.Close : window.close" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Sub DoCancel() : window.close : End Sub" & vbCrLf & _
        "Sub DoKeyPress(k)" & vbCrLf & "  If k = 13 Then" & vbCrLf & "    Call DoOK()" & vbCrLf & "  End If" & vbCrLf & "End Sub" & vbCrLf & _
        "Sub Window_OnLoad() : window.resizeTo 350, 220 : document.getElementById(""usr"").focus : End Sub" & vbCrLf & _
        "</script></head><body>" & _
        "<table><tr><td><b>Usager</b></td><td><input id=""usr"" value=""" & mSapUser & """></td></tr>" & _
        "<tr><td><b>Mot de passe</b></td><td><input id=""pwd"" type=""password"" onkeypress=""DoKeyPress(window.event.keyCode)""></td></tr></table><br>" & _
        "<center><button onclick=""DoOK()"">OK</button>&nbsp;&nbsp;" & _
        "<button onclick=""DoCancel()"">ANNULER</button></center></body></html>"

    Dim f As Object
    Set f = fso.CreateTextFile(tempHta, True)
    f.Write h
    f.Close

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & tempHta & """", 1, True

    If fso.FileExists(resultFile) Then
        Set f = fso.OpenTextFile(resultFile, 1)
        mSapUser = f.ReadLine
        mSapPass = f.ReadLine
        f.Close
        fso.DeleteFile resultFile
        GetSAPLogin = True
    Else
        GetSAPLogin = False
    End If
    If fso.FileExists(tempHta) Then fso.DeleteFile tempHta
End Function

Private Function ConnectSAP() As Object
    Dim company As Object
    Set company = CreateObject("SAPbobsCOM.Company")

    If mSapUser = "" Or mSapPass = "" Then
        If Not GetSAPLogin() Then
            Set ConnectSAP = Nothing: Exit Function
        End If
    End If

    company.Server = "NDB@hana:30013"
    company.DbServerType = 9
    company.CompanyDB = "SAP_GROMEC_PROD"
    company.UseTrusted = False
    company.UserName = mSapUser
    company.Password = mSapPass

    Dim ret As Long
    ret = company.Connect()
    If ret <> 0 Then
        MsgBox "Connexion SAP echouee (" & ret & "): " & company.GetLastErrorDescription(), vbCritical
        mSapPass = ""
        Set ConnectSAP = Nothing: Exit Function
    End If
    Set ConnectSAP = company
End Function

Private Function RunQuery(company As Object, sql As String) As Object
    Dim rs As Object
    Set rs = company.GetBusinessObject(300)
    CallByName rs, "DoQuery", VbMethod, sql
    Set RunQuery = rs
End Function

Private Function GF(rs As Object, fn As String) As Variant
    On Error Resume Next
    Dim flds As Object, fld As Object
    Set flds = CallByName(rs, "Fields", VbGet)
    Set fld = CallByName(flds, "Item", VbMethod, fn)
    GF = CallByName(fld, "Value", VbGet)
    If Err.Number <> 0 Then GF = ""
    On Error GoTo 0
End Function

Private Function RC(rs As Object) As Long
    RC = CLng(CallByName(rs, "RecordCount", VbGet))
End Function

Private Sub MN(rs As Object)
    CallByName rs, "MoveNext", VbMethod
End Sub

Private Function Txt(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Txt = "" Else Txt = CStr(v)
End Function

Private Function Num(v As Variant) As Double
    If IsNumeric(v) Then Num = CDbl(v) Else Num = 0
End Function

Private Function DateVal(v As Variant) As Variant
    DateVal = ""
    If IsDate(v) Then
        If Year(CDate(v)) > 1901 Then DateVal = CDate(v)
    End If
End Function

' Liste des codes marques X (colonne A) dans un onglet, valeur en colonne colVal
Private Function Filtre(wsName As String, colVal As Long) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(wsName)
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, colVal).End(xlUp).Row
    Dim r As Long, s As String
    For r = 4 To lr
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "X" Then
            If IsNumeric(ws.Cells(r, colVal).Value) Then
                If s <> "" Then s = s & ","
                s = s & CLng(ws.Cells(r, colVal).Value)
            End If
        End If
    Next r
    Filtre = s
End Function

' ==========================================
'  CHARGER GROUPES + LISTES
' ==========================================
Public Sub ChargerGroupesEtListes()
    Application.StatusBar = "Connexion SAP..."
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Exit Sub
    Application.ScreenUpdating = False
    On Error GoTo ErrCGL

    Dim wsG As Worksheet: Set wsG = ThisWorkbook.Sheets("Groupes")
    Dim wsL As Worksheet: Set wsL = ThisWorkbook.Sheets("Listes")
    Dim lr As Long, r As Long

    lr = wsG.Cells(wsG.Rows.Count, 2).End(xlUp).Row
    If lr > 3 Then wsG.Range(wsG.Cells(4, 1), wsG.Cells(lr, 4)).ClearContents
    Application.StatusBar = "Chargement groupes..."
    Dim rsG As Object
    Set rsG = RunQuery(co, "SELECT g.""ItmsGrpCod"", g.""ItmsGrpNam"", COUNT(i.""ItemCode"") AS ""NbItems"" " & _
        "FROM ""OITB"" g LEFT JOIN ""OITM"" i ON g.""ItmsGrpCod"" = i.""ItmsGrpCod"" " & _
        "GROUP BY g.""ItmsGrpCod"", g.""ItmsGrpNam"" ORDER BY g.""ItmsGrpNam""")
    Dim cntG As Long: cntG = RC(rsG)
    For r = 1 To cntG
        wsG.Cells(3 + r, 2).Value = GF(rsG, "ItmsGrpCod")
        wsG.Cells(3 + r, 3).Value = Txt(GF(rsG, "ItmsGrpNam"))
        wsG.Cells(3 + r, 4).Value = GF(rsG, "NbItems")
        MN rsG
    Next r

    lr = wsL.Cells(wsL.Rows.Count, 2).End(xlUp).Row
    If lr > 3 Then wsL.Range(wsL.Cells(4, 1), wsL.Cells(lr, 3)).ClearContents
    Application.StatusBar = "Chargement listes de prix..."
    Dim rsL As Object
    Set rsL = RunQuery(co, "SELECT ""ListNum"", ""ListName"" FROM ""OPLN"" ORDER BY ""ListNum""")
    Dim cntL As Long: cntL = RC(rsL)
    For r = 1 To cntL
        wsL.Cells(3 + r, 2).Value = GF(rsL, "ListNum")
        wsL.Cells(3 + r, 3).Value = Txt(GF(rsL, "ListName"))
        MN rsL
    Next r

    WriteLog "Charge " & cntG & " groupes, " & cntL & " listes"
    MsgBox cntG & " groupes et " & cntL & " listes charges." & vbCrLf & _
           "Mettre X dans les onglets Groupes / Listes pour filtrer (aucun X = tout).", vbInformation
    GoTo CleanCGL
ErrCGL:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanCGL:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' Execute une requete et ajoute les lignes au tableau resultat
Private Function AjouterResultats(co As Object, sql As String, src As String, ws As Worksheet, ByRef nextRow As Long) As Long
    Dim rs As Object
    Application.StatusBar = "Requete: " & src & "..."
    On Error Resume Next
    Set rs = RunQuery(co, sql)
    If Err.Number <> 0 Then
        mErreurs = mErreurs & vbCrLf & "- " & src & ": " & Err.Description
        WriteLog "ERREUR " & src & ": " & Err.Description & " | SQL: " & sql
        Err.Clear
        On Error GoTo 0
        AjouterResultats = 0
        Exit Function
    End If
    On Error GoTo 0

    Dim n As Long: n = RC(rs)
    AjouterResultats = n
    If n = 0 Then Exit Function

    Dim arr() As Variant: ReDim arr(1 To n, 1 To NB_COLS)
    Dim r As Long, prix As Double, cout As Double, q As Variant
    For r = 1 To n
        prix = Num(GF(rs, "Prix"))
        cout = Num(GF(rs, "Cout"))
        arr(r, 1) = src
        arr(r, 2) = Txt(GF(rs, "GrpNam"))
        arr(r, 3) = Txt(GF(rs, "ItemCode"))
        arr(r, 4) = Txt(GF(rs, "ItemName"))
        arr(r, 5) = GF(rs, "ListNum")
        arr(r, 6) = Txt(GF(rs, "ListName"))
        arr(r, 7) = Txt(GF(rs, "CardCode"))
        arr(r, 8) = Txt(GF(rs, "CardName"))
        arr(r, 9) = DateVal(GF(rs, "FromDate"))
        arr(r, 10) = DateVal(GF(rs, "ToDate"))
        q = GF(rs, "Qte")
        If IsNumeric(q) And Txt(q) <> "" Then arr(r, 11) = CDbl(q) Else arr(r, 11) = ""
        arr(r, 12) = Txt(GF(rs, "Devise"))
        arr(r, 13) = prix
        arr(r, 14) = cout
        arr(r, 15) = Round(cout - prix, 4)
        If prix <> 0 Then arr(r, 16) = (prix - cout) / prix Else arr(r, 16) = ""
        MN rs
        If r Mod 500 = 0 Then
            Application.StatusBar = src & ": " & r & " / " & n
            DoEvents
        End If
    Next r

    ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow + n - 1, NB_COLS)).Value = arr
    nextRow = nextRow + n
End Function

' ==========================================
'  ANALYSER VENTES A PERTE
' ==========================================
Public Sub AnalyserVentesPerte()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Ventes a perte")
    Dim grp As String: grp = Filtre("Groupes", 2)
    Dim lst As String: lst = Filtre("Listes", 2)

    Dim msg As String
    msg = "Groupes: " & IIf(grp = "", "TOUS", grp) & vbCrLf & _
          "Listes de prix: " & IIf(lst = "", "TOUTES", lst) & vbCrLf & vbCrLf & "Lancer l'analyse?"
    If MsgBox(msg, vbQuestion + vbYesNo, "Ventes a perte") <> vbYes Then Exit Sub

    Application.StatusBar = "Connexion SAP..."
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Exit Sub

    Dim t0 As Double: t0 = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrAVP
    mErreurs = ""

    ' Vider resultats precedents
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr > ROW_HDR Then
        ws.Range(ws.Cells(ROW_HDR + 1, 1), ws.Cells(lr, NB_COLS)).ClearContents
        ws.Range(ws.Cells(ROW_HDR + 1, 1), ws.Cells(lr, NB_COLS)).Interior.ColorIndex = xlNone
    End If

    ' --- Morceaux SQL communs ---
    Dim selItem As String, joinItem As String, wBase As String, wLst As String
    selItem = " b.""ItmsGrpNam"" AS ""GrpNam"", i.""ItemCode"", i.""ItemName"", w.""AvgPrice"" AS ""Cout"""
    joinItem = " JOIN ""OITM"" i ON i.""ItemCode"" = x.""ItemCode""" & _
               " JOIN ""OITB"" b ON b.""ItmsGrpCod"" = i.""ItmsGrpCod""" & _
               " JOIN ""OITW"" w ON w.""ItemCode"" = i.""ItemCode"" AND w.""WhsCode"" = '" & WHS_COUT & "'"
    wBase = " AND w.""AvgPrice"" > 0 AND i.""SellItem"" = 'Y' AND IFNULL(i.""frozenFor"", 'N') <> 'Y'" & _
            " AND UPPER(LEFT(i.""ItemCode"", 2)) NOT IN (" & PREFIXES_EXCLUS & ")"
    If grp <> "" Then wBase = wBase & " AND i.""ItmsGrpCod"" IN (" & grp & ")"
    If lst <> "" Then wLst = " AND x.""PriceList"" IN (" & lst & ")"

    Dim sql As String, nextRow As Long
    Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long
    nextRow = ROW_HDR + 1

    ' 1) Listes de prix (ITM1)
    sql = "SELECT" & selItem & ", x.""PriceList"" AS ""ListNum"", l.""ListName"", x.""Currency"" AS ""Devise"", x.""Price"" AS ""Prix""" & _
          " FROM ""ITM1"" x" & joinItem & _
          " LEFT JOIN ""OPLN"" l ON l.""ListNum"" = x.""PriceList""" & _
          " WHERE x.""Price"" > 0 AND x.""Price"" < w.""AvgPrice""" & _
          " AND x.""PriceList"" NOT IN (" & LISTES_EXCLUES & ")" & wBase & wLst
    n1 = AjouterResultats(co, sql, "Liste de prix", ws, nextRow)

    ' 2) Prix speciaux client (OSPP)
    sql = "SELECT" & selItem & ", x.""CardCode"", c.""CardName"", x.""ListNum"", l.""ListName"", x.""Currency"" AS ""Devise"", x.""Price"" AS ""Prix""" & _
          " FROM ""OSPP"" x" & joinItem & _
          " LEFT JOIN ""OCRD"" c ON c.""CardCode"" = x.""CardCode""" & _
          " LEFT JOIN ""OPLN"" l ON l.""ListNum"" = x.""ListNum""" & _
          " WHERE x.""Price"" > 0 AND x.""Price"" < w.""AvgPrice""" & wBase
    n2 = AjouterResultats(co, sql, "Prix special", ws, nextRow)

    ' 3) Prix speciaux par periode (SPP1, non expires)
    sql = "SELECT" & selItem & ", x.""CardCode"", c.""CardName"", x.""ListNum"", l.""ListName"", x.""FromDate"", x.""ToDate""," & _
          " x.""Currency"" AS ""Devise"", x.""Price"" AS ""Prix""" & _
          " FROM ""SPP1"" x" & joinItem & _
          " LEFT JOIN ""OCRD"" c ON c.""CardCode"" = x.""CardCode""" & _
          " LEFT JOIN ""OPLN"" l ON l.""ListNum"" = x.""ListNum""" & _
          " WHERE x.""Price"" > 0 AND x.""Price"" < w.""AvgPrice""" & _
          " AND (x.""ToDate"" IS NULL OR x.""ToDate"" >= CURRENT_DATE)" & wBase
    n3 = AjouterResultats(co, sql, "Special periode", ws, nextRow)

    ' 4) Prix speciaux par quantite (SPP2, periode non expiree)
    sql = "SELECT" & selItem & ", x.""CardCode"", c.""CardName"", d.""ListNum"", l.""ListName"", d.""FromDate"", d.""ToDate""," & _
          " x.""Amount"" AS ""Qte"", x.""Currency"" AS ""Devise"", x.""Price"" AS ""Prix""" & _
          " FROM ""SPP2"" x" & _
          " JOIN ""SPP1"" d ON d.""ItemCode"" = x.""ItemCode"" AND d.""CardCode"" = x.""CardCode"" AND d.""LINENUM"" = x.""SPP1LNum""" & _
          joinItem & _
          " LEFT JOIN ""OCRD"" c ON c.""CardCode"" = x.""CardCode""" & _
          " LEFT JOIN ""OPLN"" l ON l.""ListNum"" = d.""ListNum""" & _
          " WHERE x.""Price"" > 0 AND x.""Price"" < w.""AvgPrice""" & _
          " AND (d.""ToDate"" IS NULL OR d.""ToDate"" >= CURRENT_DATE)" & wBase
    n4 = AjouterResultats(co, sql, "Special quantite", ws, nextRow)

    Dim total As Long: total = n1 + n2 + n3 + n4
    Dim lastRow As Long: lastRow = ROW_HDR + total

    If total > 0 Then
        Dim rng As Range
        Set rng = ws.Range(ws.Cells(ROW_HDR, 1), ws.Cells(lastRow, NB_COLS))
        rng.Sort Key1:=ws.Cells(ROW_HDR, 15), Order1:=xlDescending, Header:=xlYes
        ' Devise etrangere en orange
        Dim r As Long, dv As String
        For r = ROW_HDR + 1 To lastRow
            dv = Trim(CStr(ws.Cells(r, 12).Value))
            If dv <> "" And dv <> "$" Then ws.Range(ws.Cells(r, 1), ws.Cells(r, NB_COLS)).Interior.Color = RGB(255, 204, 153)
        Next r
        rng.AutoFilter
    End If

    ws.Cells(4, 1).Value = total & " ventes a perte  |  Listes: " & n1 & "  Speciaux: " & n2 & _
        "  Periodes: " & n3 & "  Quantites: " & n4 & "  |  Groupes: " & IIf(grp = "", "tous", grp) & _
        "  |  " & Format(Now, "yyyy-mm-dd hh:nn")
    ws.Cells(4, 1).Font.Bold = True

    WriteLog "Analyse: " & total & " lignes (grp=" & grp & " lst=" & lst & ") en " & Format(Timer - t0, "0") & "s"

    Dim fin As String
    fin = total & " ventes a perte trouvees en " & Format(Timer - t0, "0") & " s."
    If mErreurs <> "" Then fin = fin & vbCrLf & vbCrLf & "Requetes en erreur (voir VentesPerte.log):" & mErreurs
    MsgBox fin, IIf(mErreurs = "", vbInformation, vbExclamation), "Ventes a perte"
    GoTo CleanAVP
ErrAVP:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanAVP:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ws.Activate
End Sub
'@

try {
    $vbMod = $wb.VBProject.VBComponents.Add(1)
    $vbMod.Name = "ModVentesPerte"
    $vbMod.CodeModule.AddFromString($vbaCode)
    Write-Host "VBA injecte avec succes!"
} catch {
    Write-Host "ERREUR injection VBA: $($_.Exception.Message)"
    Write-Host "Verifiez que 'Acces approuve au modele objet VBA' est active dans Excel."
    exit 1
}

$wb.SaveAs($xlsmPath, 52)
Write-Host "Fichier sauvegarde: $xlsmPath"
Write-Host ""
Write-Host "====================================="
Write-Host "  VentesPerte.xlsm cree et ouvert!"
Write-Host "  Boutons (ou Alt+F8):"
Write-Host "    - ChargerGroupesEtListes : Charge groupes + listes de SAP"
Write-Host "    - AnalyserVentesPerte    : Trouve prix vente < cout moyen"
Write-Host "====================================="
