<#
  GenererVentesPerte.ps1 - Genere VentesPerte.xlsm
  Detecte les ventes a perte : tout prix de vente SAP < cout moyen magasin 10 (OITW.AvgPrice, WhsCode=10)
  Sources verifiees :
    - Listes de prix          (ITM1)
    - Prix speciaux client     (OSPP)
    - Prix speciaux par dates  (SPP1, non expires)
    - Prix speciaux par qte    (SPP2, periode non expiree)
  Onglet Recherche article : meme analyse filtree par numero d'article
  Exclus : codes commencant par PT, TA, TB, TC, TD, PR, listes de prix 1, 2, 3, 11 (Coutant moyen), articles inactifs
  Filtre groupe Gromec (3 premiers car. du nom de groupe SAP) en B2 de l'onglet Ventes a perte
  Onglet Groupes = liste de reference (code a taper en B2). Filtre optionnel par liste (onglet Listes)
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
$wsArt = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsRes)
$wsArt.Name = "Recherche article"

function Set-Header($ws, $row, $headers, $color) {
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $c = $ws.Cells($row, $i + 1)
        $c.Value2 = $headers[$i]
        $c.Font.Bold = $true
        $c.Interior.Color = $color
        $c.Font.Color = 0xFFFFFF
    }
}

# ===== MISE EN PAGE ONGLET RESULTATS (Ventes a perte / Recherche article) =====
function Set-ResultSheet($ws, $titre, $labelB2, $aideB2) {
    $ws.Cells(1,1).Value2 = $titre
    $ws.Cells(1,1).Font.Bold = $true
    $ws.Cells(1,1).Font.Size = 14
    $ws.Cells(2,1).Value2 = $labelB2
    $ws.Cells(2,1).Font.Bold = $true
    $ws.Cells(2,2).NumberFormat = "@"
    $ws.Cells(2,2).Interior.Color = 0xCCFFFF
    $ws.Cells(2,2).Borders.LineStyle = 1
    $ws.Cells(2,2).Font.Bold = $true
    $ws.Cells(2,3).Value2 = $aideB2
    $ws.Cells(2,3).Font.Italic = $true
    $ws.Cells(2,3).Font.Color = 0x808080
    $ws.Cells(3,1).Value2 = "Profit min %:"
    $ws.Cells(3,1).Font.Bold = $true
    $ws.Cells(3,2).Interior.Color = 0xCCFFFF
    $ws.Cells(3,2).Borders.LineStyle = 1
    $ws.Cells(3,2).Font.Bold = $true
    $ws.Cells(3,3).Value2 = "<- % de profit minimum sur cout (ex: 15 = prix < cout x 1.15). Vide = 0 (vente a perte). Orange = devise autre que `$. ROUGE = prix vente < dernier prix d'achat"
    $ws.Cells(3,3).Font.Italic = $true
    $ws.Cells(3,3).Font.Color = 0x808080

    $hdrR = @("Source","Groupe","Article","Description","Liste #","Lie a","Facteur","Prix Net","Prix base","% profit voulu","Nom liste","Client","Nom client",
              "Date debut","Date fin","Qte min","Devise","Prix vente","Cout mag 10","Prix min","Manque `$","Profit % /cout","Dern. prix achat","Date dern. achat","Sous achat")
    Set-Header $ws 5 $hdrR 0x00008B
    $widths = @(14,28,16,40,7,7,8,8,10,9,24,12,30,11,11,8,7,12,12,12,11,10,12,11,9)
    for ($i = 0; $i -lt $widths.Count; $i++) { $ws.Columns($i + 1).ColumnWidth = $widths[$i] }
    $ws.Columns("C:C").NumberFormat = "@"
    $ws.Columns("L:L").NumberFormat = "@"
    $ws.Columns("G:G").NumberFormat = "0.0000"
    $ws.Columns("H:H").HorizontalAlignment = -4108
    $ws.Columns("I:I").NumberFormat = "#,##0.00"
    $ws.Columns("R:U").NumberFormat = "#,##0.00"
    $ws.Columns("V:V").NumberFormat = "0.0%"
    $ws.Columns("W:W").NumberFormat = "#,##0.00"
    $ws.Rows(5).Font.Bold = $true
    $ws.Rows(1).RowHeight = 30
}

Set-ResultSheet $wsRes "VENTES A PERTE - Prix de vente < Cout moyen magasin 10" "Groupe Gromec:" `
    "<- code groupe, 3 premiers chiffres (ex: 716 ou 716,101). Vide = tous. Liste des codes: onglet Groupes."
Set-ResultSheet $wsArt "RECHERCHE PAR ARTICLE - Prix de vente < Cout moyen magasin 10" "No article:" `
    "<- numero d'article (ex: 701096 ou 701096,701097). Memes regles que Ventes a perte, mais liste 1 incluse."
# Colonnes modifiables (F Lie a, G Facteur, H Prix Net) en vert
$wsArt.Range("F5:H5").Interior.Color = 0x008000
$wsArt.Range("R5").Interior.Color = 0x008000
$wsArt.Range("J5").Interior.Color = 0x008000
$wsRes.Range("F5:H5").Interior.Color = 0x008000
$wsRes.Range("R5").Interior.Color = 0x008000
$wsRes.Range("J5").Interior.Color = 0x008000
$wsArt.Cells(4,6).Value2 = "Modifiable (vert) -> bouton Appliquer. N = Lie a x Facteur. Y = prix net saisi en R."
$wsArt.Cells(4,6).Font.Italic = $true
$wsArt.Cells(4,6).Font.Color = 0x008000

# Boutons
try {
    $b1 = $wsRes.Buttons().Add($wsRes.Range("J1").Left, 2, 160, 26)
    $b1.Text = "1. Charger groupes/listes"
    $b1.OnAction = "ChargerGroupesEtListes"
    $b2 = $wsRes.Buttons().Add($wsRes.Range("J1").Left + 170, 2, 160, 26)
    $b2.Text = "2. Analyser ventes a perte"
    $b2.OnAction = "AnalyserVentesPerte"
    $b3 = $wsArt.Buttons().Add($wsArt.Range("J1").Left, 2, 160, 26)
    $b3.Text = "Rechercher article"
    $b3.OnAction = "RechercherArticle"
    $b4 = $wsArt.Buttons().Add($wsArt.Range("J1").Left + 170, 2, 200, 26)
    $b4.Text = "Appliquer modifs dans SAP"
    $b4.OnAction = "AppliquerModifsArticle"
    $b5 = $wsArt.Buttons().Add($wsArt.Range("J1").Left + 380, 2, 160, 26)
    $b5.Text = "Annuler derniere modif"
    $b5.OnAction = "AnnulerDerniereModif"
    $b6 = $wsRes.Buttons().Add($wsRes.Range("J1").Left + 340, 2, 200, 26)
    $b6.Text = "Appliquer modifs dans SAP"
    $b6.OnAction = "AppliquerModifsVentesPerte"
    $b7 = $wsRes.Buttons().Add($wsRes.Range("J1").Left + 550, 2, 160, 26)
    $b7.Text = "Annuler derniere modif"
    $b7.OnAction = "AnnulerDerniereModif"
} catch {
    Write-Host "Boutons non crees (utiliser Alt+F8): $($_.Exception.Message)"
}

# ===== ONGLET GROUPES =====
$wsGrp.Cells(1,1).Value2 = "Liste des groupes. Taper le code (colonne A) dans la case B2 de l'onglet Ventes a perte."
$wsGrp.Cells(1,1).Font.Italic = $true
Set-Header $wsGrp 3 @("Code","Nom du groupe","Nb articles") 0x8B4513
$wsGrp.Columns("A:A").ColumnWidth = 8
$wsGrp.Columns("A:A").NumberFormat = "@"
$wsGrp.Columns("B:B").ColumnWidth = 45
$wsGrp.Columns("C:C").ColumnWidth = 11

# ===== ONGLET LISTES =====
$wsLst.Cells(1,1).Value2 = "Colonne A: X = analyser seulement ces listes (aucun X = toutes). N = ne jamais analyser cette liste."
$wsLst.Cells(1,1).Font.Italic = $true
Set-Header $wsLst 3 @("X / N","Liste #","Nom de la liste") 0x003366
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
Private mSkipManuel As Long
Private mFacteur As Double   ' 1 + profit min / 100
Private Const LOG_PATH As String = "U:\GromecOutlook\VentesPerte.log"
Private Const NB_COLS As Long = 25   ' colonne 10 (J) = % profit voulu (saisie, vide au chargement)
Private Const COL_KEY As Long = 40    ' colonne cachee = cle de ligne (Source|Article|Liste|Client) pour detecter un tri
Private Const COL_ORIG As Long = 41   ' colonnes cachees 41-43 = originaux F-G-H, 44 = original R (prix vente)
Private Const ROW_HDR As Long = 5
Private Const WHS_COUT As String = "10"   ' magasin de reference pour le cout
Private Const LISTES_EXCLUES As String = "1,2,3,11"       ' listes ignorees - Ventes a perte (11 = Coutant moyen)
Private Const LISTES_EXCLUES_ART As String = "2,3,11"    ' listes ignorees - Recherche article (liste 1 incluse)
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
Private Function Filtre(wsName As String, colVal As Long, Optional marque As String = "X") As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(wsName)
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, colVal).End(xlUp).Row
    Dim r As Long, s As String
    For r = 4 To lr
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = marque Then
            If IsNumeric(ws.Cells(r, colVal).Value) Then
                If s <> "" Then s = s & ","
                s = s & CLng(ws.Cells(r, colVal).Value)
            End If
        End If
    Next r
    Filtre = s
End Function

' Groupes Gromec saisis (ex: "101, 205") -> "'101','205'" (alphanumerique seulement)
Private Function GroupesGromec(v As Variant) As String
    Dim s As String: s = Txt(v)
    s = Replace(Replace(s, ";", ","), " ", ",")
    Dim parts() As String: parts = Split(s, ",")
    Dim k As Long, j As Long, t As String, ok As Boolean, res As String
    For k = LBound(parts) To UBound(parts)
        t = Trim(parts(k))
        ok = (t <> "")
        For j = 1 To Len(t)
            If Not (Mid(t, j, 1) Like "[0-9A-Za-z]") Then ok = False
        Next j
        If ok Then
            If res <> "" Then res = res & ","
            res = res & "'" & UCase(Left(t, 3)) & "'"
        End If
    Next k
    GroupesGromec = res
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
    Dim gNam As String
    Application.StatusBar = "Chargement groupes..."
    Dim rsG As Object
    Set rsG = RunQuery(co, "SELECT g.""ItmsGrpCod"", g.""ItmsGrpNam"", COUNT(i.""ItemCode"") AS ""NbItems"" " & _
        "FROM ""OITB"" g LEFT JOIN ""OITM"" i ON g.""ItmsGrpCod"" = i.""ItmsGrpCod"" " & _
        "GROUP BY g.""ItmsGrpCod"", g.""ItmsGrpNam"" ORDER BY g.""ItmsGrpNam""")
    Dim cntG As Long: cntG = RC(rsG)
    For r = 1 To cntG
        gNam = Txt(GF(rsG, "ItmsGrpNam"))
        wsG.Cells(3 + r, 1).Value = Left(gNam, 3)
        wsG.Cells(3 + r, 2).Value = gNam
        wsG.Cells(3 + r, 3).Value = GF(rsG, "NbItems")
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
           "Taper le code groupe en B2 (onglet Ventes a perte). X dans onglet Listes pour filtrer les listes.", vbInformation
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
    Dim r As Long, prix As Double, cout As Double, dern As Double, q As Variant
    For r = 1 To n
        prix = Num(GF(rs, "Prix"))
        cout = Num(GF(rs, "Cout"))
        arr(r, 1) = src
        arr(r, 2) = Txt(GF(rs, "GrpNam"))
        arr(r, 3) = Txt(GF(rs, "ItemCode"))
        arr(r, 4) = Txt(GF(rs, "ItemName"))
        arr(r, 5) = GF(rs, "ListNum")
        arr(r, 6) = GF(rs, "LieA")
        arr(r, 7) = GF(rs, "Facteur")
        arr(r, 8) = Txt(GF(rs, "PrixNet"))
        arr(r, 11) = Txt(GF(rs, "ListName"))
        arr(r, 12) = Txt(GF(rs, "CardCode"))
        arr(r, 13) = Txt(GF(rs, "CardName"))
        arr(r, 14) = DateVal(GF(rs, "FromDate"))
        arr(r, 15) = DateVal(GF(rs, "ToDate"))
        q = GF(rs, "Qte")
        If IsNumeric(q) And Txt(q) <> "" Then arr(r, 16) = CDbl(q) Else arr(r, 16) = ""
        arr(r, 17) = Txt(GF(rs, "Devise"))
        arr(r, 18) = prix
        arr(r, 19) = cout
        arr(r, 20) = Round(cout * mFacteur, 4)
        arr(r, 21) = Round(cout * mFacteur - prix, 4)
        arr(r, 22) = (prix - cout) / cout
        dern = Num(GF(rs, "DernAchat"))
        If dern > 0 Then arr(r, 23) = dern Else arr(r, 23) = ""
        arr(r, 24) = DateVal(GF(rs, "DernAchatDate"))
        If dern > 0 And prix < dern Then arr(r, 25) = "OUI" Else arr(r, 25) = ""
        q = GF(rs, "PrixBase")
        If IsNumeric(q) And Txt(q) <> "" Then arr(r, 9) = CDbl(q) Else arr(r, 9) = ""
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
    LancerAnalyse ThisWorkbook.Sheets("Ventes a perte"), False
End Sub

Public Sub RechercherArticle()
    LancerAnalyse ThisWorkbook.Sheets("Recherche article"), True
End Sub

' Articles saisis (ex: "701096, 701097") -> "'701096','701097'"
Private Function ArticlesSaisis(v As Variant) As String
    Dim s As String: s = Txt(v)
    s = Replace(Replace(s, ";", ","), " ", ",")
    Dim parts() As String: parts = Split(s, ",")
    Dim k As Long, t As String, res As String
    For k = LBound(parts) To UBound(parts)
        t = Trim(parts(k))
        If t <> "" Then
            If res <> "" Then res = res & ","
            res = res & "'" & Replace(t, "'", "''") & "'"
        End If
    Next k
    ArticlesSaisis = res
End Function

Private Sub LancerAnalyse(ws As Worksheet, parArticle As Boolean, Optional silencieux As Boolean = False)
    Dim gro As String, arts As String, filtreTxt As String
    If parArticle Then
        arts = ArticlesSaisis(ws.Range("B2").Value)
        If arts = "" Then MsgBox "Entrer un numero d'article en B2.", vbExclamation: Exit Sub
        filtreTxt = "Article(s): " & Replace(arts, "'", "")
    Else
        gro = GroupesGromec(ws.Range("B2").Value)
        filtreTxt = "Groupe: " & IIf(gro = "", "tous", Replace(gro, "'", ""))
    End If
    Dim pct As Double: pct = Val(Replace(Txt(ws.Range("B3").Value), ",", "."))
    If pct < 0 Then pct = 0
    mFacteur = 1 + pct / 100
    Dim sFact As String: sFact = Trim(Str(mFacteur))
    Dim lst As String: lst = Filtre("Listes", 2)
    Dim lstN As String: lstN = Filtre("Listes", 2, "N")

    Dim msg As String
    msg = filtreTxt & vbCrLf & _
          "Profit min: " & pct & " % sur cout" & vbCrLf & _
          "Listes de prix: " & IIf(lst = "", "TOUTES", lst) & IIf(lstN = "", "", "  sauf " & lstN) & vbCrLf & vbCrLf & "Lancer l'analyse?"
    If Not silencieux Then
        If MsgBox(msg, vbQuestion + vbYesNo, "Ventes a perte") <> vbYes Then Exit Sub
    End If

    Application.StatusBar = "Connexion SAP..."
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Exit Sub

    Dim t0 As Double: t0 = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrAVP
    mErreurs = ""

    ' Vider resultats precedents
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr > ROW_HDR Then
        ws.Range(ws.Cells(ROW_HDR + 1, 1), ws.Cells(lr, COL_ORIG + 3)).ClearContents
        ws.Range(ws.Cells(ROW_HDR + 1, 1), ws.Cells(lr, NB_COLS)).Interior.ColorIndex = xlNone
    End If

    ' --- Morceaux SQL communs ---
    Dim selItem As String, joinItem As String, wBase As String, wLst As String
    selItem = " b.""ItmsGrpNam"" AS ""GrpNam"", i.""ItemCode"", i.""ItemName"", w.""AvgPrice"" AS ""Cout"", i.""LastPurPrc"" AS ""DernAchat"", i.""LastPurDat"" AS ""DernAchatDate"""
    joinItem = " JOIN ""OITM"" i ON i.""ItemCode"" = x.""ItemCode""" & _
               " JOIN ""OITB"" b ON b.""ItmsGrpCod"" = i.""ItmsGrpCod""" & _
               " JOIN ""OITW"" w ON w.""ItemCode"" = i.""ItemCode"" AND w.""WhsCode"" = '" & WHS_COUT & "'"
    wBase = " AND w.""AvgPrice"" > 0 AND i.""SellItem"" = 'Y' AND IFNULL(i.""frozenFor"", 'N') <> 'Y'" & _
            " AND UPPER(LEFT(i.""ItemCode"", 2)) NOT IN (" & PREFIXES_EXCLUS & ")"
    If gro <> "" Then wBase = wBase & " AND LEFT(b.""ItmsGrpNam"", 3) IN (" & gro & ")"
    If arts <> "" Then wBase = wBase & " AND i.""ItemCode"" IN (" & arts & ")"
    If lst <> "" Then wLst = " AND x.""PriceList"" IN (" & lst & ")"
    If lstN <> "" Then wLst = wLst & " AND x.""PriceList"" NOT IN (" & lstN & ")"
    ' Meme filtre X / N pour les prix speciaux, sur leur liste (colonne Liste #)
    Dim wLstSp As String, wLstQte As String
    If lst <> "" Then wLstSp = " AND x.""ListNum"" IN (" & lst & ")": wLstQte = " AND d.""ListNum"" IN (" & lst & ")"
    If lstN <> "" Then
        wLstSp = wLstSp & " AND x.""ListNum"" NOT IN (" & lstN & ")"
        wLstQte = wLstQte & " AND d.""ListNum"" NOT IN (" & lstN & ")"
    End If

    Dim sql As String, nextRow As Long
    Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long
    nextRow = ROW_HDR + 1

    ' 1) Listes de prix (ITM1)
    sql = "SELECT" & selItem & ", x.""PriceList"" AS ""ListNum"", x.""BasePLNum"" AS ""LieA"", x.""Factor"" AS ""Facteur"", x.""Ovrwritten"" AS ""PrixNet"", CASE WHEN x.""BasePLNum"" = -1 THEN i.""LastPurPrc"" WHEN x.""BasePLNum"" = -2 THEN i.""LstEvlPric""" & _
          " ELSE (SELECT pb.""Price"" FROM ""ITM1"" pb WHERE pb.""ItemCode"" = x.""ItemCode"" AND pb.""PriceList"" = x.""BasePLNum"") END AS ""PrixBase"", l.""ListName"", x.""Currency"" AS ""Devise"", x.""Price"" AS ""Prix""" & _
          " FROM ""ITM1"" x" & joinItem & _
          " LEFT JOIN ""OPLN"" l ON l.""ListNum"" = x.""PriceList""" & _
          " WHERE x.""Price"" > 0 AND x.""Price"" < w.""AvgPrice"" * " & sFact & _
          " AND x.""PriceList"" NOT IN (" & IIf(parArticle, LISTES_EXCLUES_ART, LISTES_EXCLUES) & ")" & wBase & wLst
    n1 = AjouterResultats(co, sql, "Liste de prix", ws, nextRow)

    ' 2) Prix speciaux client (OSPP)
    sql = "SELECT" & selItem & ", x.""CardCode"", c.""CardName"", x.""ListNum"", l.""ListName"", x.""Currency"" AS ""Devise"", x.""Price"" AS ""Prix""" & _
          " FROM ""OSPP"" x" & joinItem & _
          " LEFT JOIN ""OCRD"" c ON c.""CardCode"" = x.""CardCode""" & _
          " LEFT JOIN ""OPLN"" l ON l.""ListNum"" = x.""ListNum""" & _
          " WHERE x.""Price"" > 0 AND x.""Price"" < w.""AvgPrice"" * " & sFact & wBase & wLstSp
    n2 = AjouterResultats(co, sql, "Prix special", ws, nextRow)

    ' 3) Prix speciaux par periode (SPP1, non expires)
    sql = "SELECT" & selItem & ", x.""CardCode"", c.""CardName"", x.""ListNum"", l.""ListName"", x.""FromDate"", x.""ToDate""," & _
          " x.""Currency"" AS ""Devise"", x.""Price"" AS ""Prix""" & _
          " FROM ""SPP1"" x" & joinItem & _
          " LEFT JOIN ""OCRD"" c ON c.""CardCode"" = x.""CardCode""" & _
          " LEFT JOIN ""OPLN"" l ON l.""ListNum"" = x.""ListNum""" & _
          " WHERE x.""Price"" > 0 AND x.""Price"" < w.""AvgPrice"" * " & sFact & _
          " AND (x.""ToDate"" IS NULL OR x.""ToDate"" >= CURRENT_DATE)" & wBase & wLstSp
    n3 = AjouterResultats(co, sql, "Special periode", ws, nextRow)

    ' 4) Prix speciaux par quantite (SPP2, periode non expiree)
    sql = "SELECT" & selItem & ", x.""CardCode"", c.""CardName"", d.""ListNum"", l.""ListName"", d.""FromDate"", d.""ToDate""," & _
          " x.""Amount"" AS ""Qte"", x.""Currency"" AS ""Devise"", x.""Price"" AS ""Prix""" & _
          " FROM ""SPP2"" x" & _
          " JOIN ""SPP1"" d ON d.""ItemCode"" = x.""ItemCode"" AND d.""CardCode"" = x.""CardCode"" AND d.""LINENUM"" = x.""SPP1LNum""" & _
          joinItem & _
          " LEFT JOIN ""OCRD"" c ON c.""CardCode"" = x.""CardCode""" & _
          " LEFT JOIN ""OPLN"" l ON l.""ListNum"" = d.""ListNum""" & _
          " WHERE x.""Price"" > 0 AND x.""Price"" < w.""AvgPrice"" * " & sFact & _
          " AND (d.""ToDate"" IS NULL OR d.""ToDate"" >= CURRENT_DATE)" & wBase & wLstQte
    n4 = AjouterResultats(co, sql, "Special quantite", ws, nextRow)

    Dim total As Long: total = n1 + n2 + n3 + n4
    Dim lastRow As Long: lastRow = ROW_HDR + total

    If total > 0 Then
        Dim rng As Range
        Set rng = ws.Range(ws.Cells(ROW_HDR, 1), ws.Cells(lastRow, NB_COLS))
        rng.Sort Key1:=ws.Cells(ROW_HDR, 21), Order1:=xlDescending, Header:=xlYes
        ' Devise etrangere en orange, prix < dernier achat en rouge
        Dim r As Long, dv As String, nSous As Long
        For r = ROW_HDR + 1 To lastRow
            dv = Trim(CStr(ws.Cells(r, 17).Value))
            If dv <> "" And dv <> "$" Then ws.Range(ws.Cells(r, 1), ws.Cells(r, NB_COLS)).Interior.Color = RGB(255, 204, 153)
            If ws.Cells(r, 25).Value = "OUI" Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, NB_COLS)).Interior.Color = RGB(255, 150, 150)
                nSous = nSous + 1
            End If
        Next r
        ' Format date applique en VBA (NumberFormat via PowerShell = interprete en francais)
        ws.Range(ws.Cells(ROW_HDR + 1, 14), ws.Cells(lastRow, 15)).NumberFormat = "yyyy-mm-dd"
        ws.Range(ws.Cells(ROW_HDR + 1, 24), ws.Cells(lastRow, 24)).NumberFormat = "yyyy-mm-dd"
        ' Copie cachee des valeurs originales F-G-H (detection des modifs)
        ws.Range(ws.Cells(ROW_HDR + 1, COL_ORIG), ws.Cells(lastRow, COL_ORIG + 2)).Value = _
            ws.Range(ws.Cells(ROW_HDR + 1, 6), ws.Cells(lastRow, 8)).Value
        ws.Range(ws.Cells(ROW_HDR + 1, COL_ORIG + 3), ws.Cells(lastRow, COL_ORIG + 3)).Value = _
            ws.Range(ws.Cells(ROW_HDR + 1, 18), ws.Cells(lastRow, 18)).Value
        Dim cles() As Variant: ReDim cles(1 To total, 1 To 1)
        For r = ROW_HDR + 1 To lastRow
            cles(r - ROW_HDR, 1) = CleLigne(ws, r)
        Next r
        ws.Range(ws.Cells(ROW_HDR + 1, COL_KEY), ws.Cells(lastRow, COL_KEY)).Value = cles
        ws.Range(ws.Cells(ROW_HDR, COL_KEY), ws.Cells(ROW_HDR, COL_ORIG + 3)).Value = Array("cle", "orig F", "orig G", "orig H", "orig P")
        ws.Range(ws.Cells(1, COL_KEY), ws.Cells(1, COL_ORIG + 3)).EntireColumn.Hidden = True
        ' Filtre/tri sur tout le tableau, colonnes cachees incluses (elles suivent le tri)
        ws.Range(ws.Cells(ROW_HDR, 1), ws.Cells(lastRow, COL_ORIG + 3)).AutoFilter
    End If

    ws.Cells(4, 1).Value = total & " prix sous cout +" & pct & "%  |  Listes: " & n1 & "  Speciaux: " & n2 & _
        "  Periodes: " & n3 & "  Quantites: " & n4 & "  |  Sous dern. achat: " & nSous & "  |  " & filtreTxt & _
        "  |  " & Format(Now, "yyyy-mm-dd hh:nn")
    ws.Cells(4, 1).Font.Bold = True

    WriteLog "Analyse: " & total & " lignes (" & filtreTxt & " lst=" & lst & ") en " & Format(Timer - t0, "0") & "s"

    Dim fin As String
    fin = total & " prix sous cout +" & pct & "% trouves en " & Format(Timer - t0, "0") & " s." & vbCrLf & _
          nSous & " sous le dernier prix d'achat (lignes rouges)."
    If mErreurs <> "" Then fin = fin & vbCrLf & vbCrLf & "Requetes en erreur (voir VentesPerte.log):" & mErreurs
    If Not silencieux Or mErreurs <> "" Then MsgBox fin, IIf(mErreurs = "", vbInformation, vbExclamation), "Ventes a perte"
    GoTo CleanAVP
ErrAVP:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanAVP:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ws.Activate
End Sub

' ==========================================
'  APPLIQUER MODIFS F (Lie a) / G (Facteur) / H (Prix Net) DANS SAP (ITM1)
' ==========================================
Public Sub AppliquerModifsArticle()
    ' Pas de recalcul des listes liees : seule la liste de la ligne change
    AppliquerModifs ThisWorkbook.Sheets("Recherche article"), False
End Sub

' Onglet Ventes a perte : modifie seulement la liste de la ligne (pas de recalcul des listes liees)
Public Sub AppliquerModifsVentesPerte()
    AppliquerModifs ThisWorkbook.Sheets("Ventes a perte"), False
End Sub

Private Sub AppliquerModifs(ws As Worksheet, avecCascade As Boolean)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastRow <= ROW_HDR Then MsgBox "Aucun resultat. Faire une recherche d'abord.", vbExclamation: Exit Sub

    Dim ops() As Variant, nb As Long, r As Long, k As Long
    Dim art As String, liste As Long, lie As Variant, fac As Variant, pn As String, pv As Variant
    Dim erreurs As String, ignores As Long, resumeTxt As String
    ReDim ops(1 To 6, 1 To lastRow)

    ' Verifier que les lignes n'ont pas ete deplacees (tri) par rapport aux originaux caches
    For r = ROW_HDR + 1 To lastRow
        If CStr(ws.Cells(r, COL_KEY).Value) <> CleLigne(ws, r) Then
            MsgBox "Le tableau a ete trie ou deplace (ligne " & r & ")." & vbCrLf & _
                   "Relancer 'Rechercher article', puis refaire les modifications.", vbExclamation
            Exit Sub
        End If
    Next r

    For r = ROW_HDR + 1 To lastRow
        If Different(ws.Cells(r, 6).Value2, ws.Cells(r, COL_ORIG).Value2) Or _
           Different(ws.Cells(r, 7).Value2, ws.Cells(r, COL_ORIG + 1).Value2) Or _
           Different(ws.Cells(r, 8).Value2, ws.Cells(r, COL_ORIG + 2).Value2) Or _
           Different(ws.Cells(r, 18).Value2, ws.Cells(r, COL_ORIG + 3).Value2) Then

            If CStr(ws.Cells(r, 1).Value) <> "Liste de prix" Then
                ignores = ignores + 1
            Else
                art = CStr(ws.Cells(r, 3).Value)
                liste = CLng(ws.Cells(r, 5).Value)
                lie = ws.Cells(r, 6).Value
                fac = ws.Cells(r, 7).Value
                pn = UCase(Trim(CStr(ws.Cells(r, 8).Value)))
                pv = ws.Cells(r, 18).Value
                If pn <> "Y" And pn <> "N" Then
                    erreurs = erreurs & vbCrLf & "Ligne " & r & ": Prix Net doit etre Y ou N"
                ElseIf pn = "Y" Then
                    If Not IsNumeric(pv) Or Trim(CStr(pv)) = "" Then
                        erreurs = erreurs & vbCrLf & "Ligne " & r & ": Prix vente (R) invalide"
                    ElseIf CDbl(pv) <= 0 Then
                        erreurs = erreurs & vbCrLf & "Ligne " & r & ": Prix vente (R) doit etre > 0"
                    Else
                        nb = nb + 1
                        ops(1, nb) = art: ops(2, nb) = liste: ops(3, nb) = 0
                        ops(4, nb) = 0: ops(5, nb) = "Y": ops(6, nb) = CDbl(pv)
                        If nb <= 25 Then resumeTxt = resumeTxt & vbCrLf & art & "  liste " & liste & _
                            " -> PRIX NET " & Format(CDbl(pv), "0.00")
                    End If
                ElseIf Different(ws.Cells(r, 18).Value2, ws.Cells(r, COL_ORIG + 3).Value2) Then
                    erreurs = erreurs & vbCrLf & "Ligne " & r & ": prix modifie en R -> mettre Y en H (prix net)"
                ElseIf Not IsNumeric(lie) Or Trim(CStr(lie)) = "" Then
                    erreurs = erreurs & vbCrLf & "Ligne " & r & ": 'Lie a' invalide"
                ElseIf CLng(lie) < -2 Or CLng(lie) = 0 Then
                    erreurs = erreurs & vbCrLf & "Ligne " & r & ": 'Lie a' doit etre un no de liste, -1 (dern. prix achat) ou -2 (dern. prix determine)"
                ElseIf CLng(lie) = liste Then
                    erreurs = erreurs & vbCrLf & "Ligne " & r & ": 'Lie a' = la liste elle-meme"
                ElseIf Not IsNumeric(fac) Or Trim(CStr(fac)) = "" Then
                    erreurs = erreurs & vbCrLf & "Ligne " & r & ": Facteur invalide"
                ElseIf CDbl(fac) <= 0 Then
                    erreurs = erreurs & vbCrLf & "Ligne " & r & ": Facteur doit etre > 0"
                Else
                    nb = nb + 1
                    ops(1, nb) = art: ops(2, nb) = liste: ops(3, nb) = CLng(lie)
                    ops(4, nb) = CDbl(fac): ops(5, nb) = "N": ops(6, nb) = 0
                    If nb <= 25 Then resumeTxt = resumeTxt & vbCrLf & art & "  liste " & liste & " -> lie a " & CLng(lie) & _
                        "  x" & CDbl(fac)
                End If
            End If
        End If
    Next r

    If erreurs <> "" Then MsgBox "Corriger avant d'appliquer:" & erreurs, vbExclamation: Exit Sub
    If nb = 0 Then
        MsgBox "Aucune modification detectee dans F-G-H-R (lignes 'Liste de prix')." & _
            IIf(ignores > 0, vbCrLf & ignores & " modif(s) sur prix speciaux ignoree(s).", ""), vbInformation
        Exit Sub
    End If
    If nb > 25 Then resumeTxt = resumeTxt & vbCrLf & "... et " & (nb - 25) & " autres"
    If MsgBox(nb & " prix a modifier dans SAP:" & resumeTxt & vbCrLf & vbCrLf & _
              "Prix Net N = prix recalcule (prix liste de base x facteur)." & vbCrLf & _
              "Prix Net Y = Lie a 0, Facteur 0, prix = colonne R." & _
              IIf(ignores > 0, vbCrLf & ignores & " modif(s) sur prix speciaux ignoree(s).", "") & vbCrLf & vbCrLf & _
              "CONFIRMER?", vbYesNo + vbQuestion, "Modifier SAP") <> vbYes Then Exit Sub

    Application.StatusBar = "Connexion SAP..."
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Exit Sub

    ' Fichier d'annulation (ecrase a chaque lot) : anciennes valeurs de chaque ligne touchee
    Dim fso As Object, fU As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fU = fso.CreateTextFile(UndoPath(), True)
    fU.WriteLine "# " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & mSapUser
    fU.WriteLine "@" & ws.Name
    mSkipManuel = 0

    Dim sql As String, sFac As String, ok As Long, ko As String, rsU As Object, nCasc As Long, nc As Long
    For k = 1 To nb
        Application.StatusBar = "Mise a jour " & k & "/" & nb & " - " & ops(1, k) & " liste " & ops(2, k)
        sFac = Trim(Str(ops(4, k)))
        If ops(5, k) = "N" Then
            sql = "UPDATE ""ITM1"" AS T0 SET ""BasePLNum"" = " & ops(3, k) & ", ""Factor"" = " & sFac & _
                  ", ""Ovrwritten"" = 'N'" & _
                  ", ""Price"" = " & PrixBaseSQL(CLng(ops(3, k))) & " * " & sFac & _
                  " WHERE T0.""ItemCode"" = '" & Replace(ops(1, k), "'", "''") & "' AND T0.""PriceList"" = " & ops(2, k)
        Else
            sql = "UPDATE ""ITM1"" SET ""BasePLNum"" = 0, ""Factor"" = 0, ""Ovrwritten"" = 'Y'" & _
                  ", ""Price"" = " & Trim(Str(ops(6, k))) & _
                  " WHERE ""ItemCode"" = '" & Replace(ops(1, k), "'", "''") & "' AND ""PriceList"" = " & ops(2, k)
        End If
        On Error Resume Next
        SauverAncien co, fU, CStr(ops(1, k)), CLng(ops(2, k))
        If Err.Number <> 0 Then
            ' Pas de sauvegarde possible -> on ne modifie pas cette ligne
            ko = ko & vbCrLf & ops(1, k) & " liste " & ops(2, k) & ": sauvegarde impossible (" & Err.Description & "), non modifie"
            Err.Clear
            GoTo LigneSuivante
        End If
        If ops(5, k) = "N" Then
            ' Securite : ne jamais ecrire un prix a 0 si le prix de base est vide
            Set rsU = RunQuery(co, "SELECT " & PrixBaseSQL(CLng(ops(3, k))) & " AS ""PB"" FROM ""ITM1"" T0" & _
                " WHERE T0.""ItemCode"" = '" & Replace(ops(1, k), "'", "''") & "' AND T0.""PriceList"" = " & ops(2, k))
            If Num(GF(rsU, "PB")) <= 0 Then
                ko = ko & vbCrLf & ops(1, k) & " liste " & ops(2, k) & ": prix de base (lie a " & ops(3, k) & ") = 0, non modifie"
                Err.Clear
                GoTo LigneSuivante
            End If
        End If
        Set rsU = co.GetBusinessObject(300)
        CallByName rsU, "DoQuery", VbMethod, sql
        If Err.Number <> 0 Then
            ko = ko & vbCrLf & ops(1, k) & " liste " & ops(2, k) & ": " & Err.Description
            WriteLog "MODIF ERREUR " & ops(1, k) & " liste " & ops(2, k) & ": " & Err.Description & " | " & sql
            Err.Clear
        Else
            ok = ok + 1
            WriteLog "MODIF " & ops(1, k) & " liste " & ops(2, k) & " lie=" & ops(3, k) & " fac=" & sFac & " net=" & ops(5, k) & " prix=" & ops(6, k)
            ' Recalcul des listes liees a cette liste (en chaine) - Recherche article seulement
            nc = 0
            If avecCascade Then nc = Cascade(co, fU, CStr(ops(1, k)), CLng(ops(2, k)))
            If Err.Number <> 0 Then
                ko = ko & vbCrLf & ops(1, k) & " listes liees a " & ops(2, k) & ": " & Err.Description
                WriteLog "CASCADE ERREUR " & ops(1, k) & " base " & ops(2, k) & ": " & Err.Description
                Err.Clear
            End If
            nCasc = nCasc + nc
        End If
LigneSuivante:
        On Error GoTo 0
        Set rsU = Nothing
    Next k
    fU.Close

    On Error Resume Next
    co.Disconnect: Set co = Nothing
    On Error GoTo 0
    Application.StatusBar = False

    MsgBox ok & " prix modifies dans SAP." & vbCrLf & _
        IIf(avecCascade, nCasc & " liste(s) liee(s) recalculee(s) automatiquement.", "Seule la liste de chaque ligne a ete modifiee (listes liees non touchees).") & _
        IIf(mSkipManuel > 0, vbCrLf & mSkipManuel & " liste(s) liee(s) en prix manuel non touchee(s).", "") & _
        IIf(ko <> "", vbCrLf & vbCrLf & "Erreurs:" & ko, "") & vbCrLf & vbCrLf & _
        "Bouton 'Annuler derniere modif' pour revenir en arriere.", _
        IIf(ko = "", vbInformation, vbExclamation), "Modifier SAP"

    ' Recharger pour afficher les nouveaux prix
    LancerAnalyse ws, (ws.Name = "Recherche article"), True
End Sub

' Cle d'identification d'une ligne de resultat
Private Function CleLigne(ws As Worksheet, r As Long) As String
    CleLigne = CStr(ws.Cells(r, 1).Value) & "|" & CStr(ws.Cells(r, 3).Value) & "|" & _
               CStr(ws.Cells(r, 5).Value) & "|" & CStr(ws.Cells(r, 12).Value)
End Function

' Compare deux valeurs : numerique avec tolerance, sinon texte (majuscules, sans espaces)
Private Function Different(a As Variant, b As Variant) As Boolean
    If IsNumeric(a) And IsNumeric(b) And Trim(CStr(a)) <> "" And Trim(CStr(b)) <> "" Then
        Different = Abs(CDbl(a) - CDbl(b)) > 0.000001
    Else
        Different = UCase(Trim(CStr(a))) <> UCase(Trim(CStr(b)))
    End If
End Function

' Expression SQL du prix de base pour un UPDATE "ITM1" AS T0.
' -1 = Dernier prix d'achat (OITM.LastPurPrc), -2 = Dernier prix determine (OITM.LstEvlPric), sinon liste ITM1.
Private Function PrixBaseSQL(base As Long) As String
    Select Case base
        Case -1
            PrixBaseSQL = "IFNULL((SELECT I.""LastPurPrc"" FROM ""OITM"" I WHERE I.""ItemCode"" = T0.""ItemCode""), 0)"
        Case -2
            PrixBaseSQL = "IFNULL((SELECT I.""LstEvlPric"" FROM ""OITM"" I WHERE I.""ItemCode"" = T0.""ItemCode""), 0)"
        Case Else
            PrixBaseSQL = "IFNULL((SELECT T1.""Price"" FROM ""ITM1"" T1 WHERE T1.""ItemCode"" = T0.""ItemCode""" & _
                          " AND T1.""PriceList"" = " & base & "), 0)"
    End Select
End Function

' ==========================================
'  ANNULATION / CASCADE - utilitaires
' ==========================================
Private Function UndoPath() As String
    UndoPath = "U:\GromecOutlook\VentesPerte_annuler_" & Environ("USERNAME") & ".txt"
End Function

' Ecrit l'etat actuel d'une ligne ITM1 dans le fichier d'annulation
Private Sub SauverAncien(co As Object, fU As Object, art As String, liste As Long)
    Dim rs As Object
    Set rs = RunQuery(co, "SELECT ""BasePLNum"", ""Factor"", IFNULL(""Ovrwritten"", 'N') AS ""Ovr"", ""Price"" FROM ""ITM1""" & _
        " WHERE ""ItemCode"" = '" & Replace(art, "'", "''") & "' AND ""PriceList"" = " & liste)
    If RC(rs) > 0 Then
        fU.WriteLine art & vbTab & liste & vbTab & CLng(Num(GF(rs, "BasePLNum"))) & vbTab & _
            Trim(Str(Num(GF(rs, "Factor")))) & vbTab & Txt(GF(rs, "Ovr")) & vbTab & Trim(Str(Num(GF(rs, "Price"))))
    End If
End Sub

' Recalcule (en chaine) les listes dont "Lie a" = listeBase pour cet article.
' Les listes en prix manuel (Ovrwritten = Y) ne sont pas touchees (comptees dans mSkipManuel).
Private Function Cascade(co As Object, fU As Object, art As String, listeBase As Long) As Long
    Dim q(1 To 500) As Long, head As Long, tail As Long, vus As String
    Dim L As Long, rs As Object, n As Long, i As Long, deps() As Long, cnt As Long
    Dim a As String: a = Replace(art, "'", "''")
    head = 1: tail = 1: q(1) = listeBase: vus = "," & listeBase & ","
    Do While head <= tail
        L = q(head): head = head + 1
        Set rs = RunQuery(co, "SELECT COUNT(*) AS ""N"" FROM ""ITM1"" WHERE ""ItemCode"" = '" & a & "'" & _
            " AND ""BasePLNum"" = " & L & " AND ""PriceList"" <> " & L & " AND IFNULL(""Ovrwritten"", 'N') = 'Y'")
        mSkipManuel = mSkipManuel + CLng(Num(GF(rs, "N")))
        Set rs = RunQuery(co, "SELECT ""PriceList"" FROM ""ITM1"" WHERE ""ItemCode"" = '" & a & "'" & _
            " AND ""BasePLNum"" = " & L & " AND ""PriceList"" <> " & L & _
            " AND IFNULL(""Ovrwritten"", 'N') <> 'Y' AND ""Factor"" > 0")
        n = RC(rs)
        If n > 0 Then
            ReDim deps(1 To n)
            For i = 1 To n
                deps(i) = CLng(GF(rs, "PriceList"))
                MN rs
            Next i
            For i = 1 To n
                If InStr(vus, "," & deps(i) & ",") = 0 And tail < 500 Then
                    vus = vus & deps(i) & ","
                    SauverAncien co, fU, art, deps(i)
                    RunQuery co, "UPDATE ""ITM1"" AS T0 SET ""Price"" = IFNULL((SELECT T1.""Price"" FROM ""ITM1"" T1" & _
                        " WHERE T1.""ItemCode"" = T0.""ItemCode"" AND T1.""PriceList"" = " & L & "), 0) * T0.""Factor""" & _
                        " WHERE T0.""ItemCode"" = '" & a & "' AND T0.""PriceList"" = " & deps(i)
                    WriteLog "CASCADE " & art & " liste " & deps(i) & " recalculee (base " & L & ")"
                    cnt = cnt + 1
                    tail = tail + 1: q(tail) = deps(i)
                End If
            Next i
        End If
    Loop
    Cascade = cnt
End Function

' ==========================================
'  ANNULER LA DERNIERE MODIFICATION (restaure les anciennes valeurs ITM1)
' ==========================================
Public Sub AnnulerDerniereModif()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim p As String: p = UndoPath()
    If Not fso.FileExists(p) Then MsgBox "Aucune modification a annuler.", vbInformation: Exit Sub

    Dim f As Object, ligne As String, entete As String, lignes() As String, n As Long, feuille As String
    ReDim lignes(1 To 1)
    Set f = fso.OpenTextFile(p, 1)
    Do While Not f.AtEndOfStream
        ligne = f.ReadLine
        If Left(ligne, 1) = "#" Then
            entete = Mid(ligne, 3)
        ElseIf Left(ligne, 1) = "@" Then
            feuille = Mid(ligne, 2)
        ElseIf UBound(Split(ligne, vbTab)) = 5 Then
            n = n + 1
            ReDim Preserve lignes(1 To n)
            lignes(n) = ligne
        End If
    Loop
    f.Close
    If n = 0 Then MsgBox "Aucune modification a annuler.", vbInformation: Exit Sub

    If MsgBox("Restaurer " & n & " prix tels qu'avant la modification du " & entete & " ?", _
              vbYesNo + vbQuestion, "Annuler derniere modif") <> vbYes Then Exit Sub

    Application.StatusBar = "Connexion SAP..."
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Exit Sub

    Dim i As Long, c() As String, ok As Long, ko As String
    ' Ordre inverse : la 1re valeur sauvegardee d'une ligne (l'originale) est appliquee en dernier
    For i = n To 1 Step -1
        c = Split(lignes(i), vbTab)
        Application.StatusBar = "Restauration " & (n - i + 1) & "/" & n
        On Error Resume Next
        RunQuery co, "UPDATE ""ITM1"" SET ""BasePLNum"" = " & CLng(Val(c(2))) & ", ""Factor"" = " & Trim(Str(Val(c(3)))) & _
            ", ""Ovrwritten"" = '" & IIf(c(4) = "Y", "Y", "N") & "', ""Price"" = " & Trim(Str(Val(c(5)))) & _
            " WHERE ""ItemCode"" = '" & Replace(c(0), "'", "''") & "' AND ""PriceList"" = " & CLng(Val(c(1)))
        If Err.Number <> 0 Then
            ko = ko & vbCrLf & c(0) & " liste " & c(1) & ": " & Err.Description
            Err.Clear
        Else
            ok = ok + 1
            WriteLog "ANNULE " & c(0) & " liste " & c(1) & " -> lie=" & c(2) & " fac=" & c(3) & " net=" & c(4) & " prix=" & c(5)
        End If
        On Error GoTo 0
    Next i

    On Error Resume Next
    co.Disconnect: Set co = Nothing
    If fso.FileExists(p & ".annule") Then fso.DeleteFile p & ".annule"
    fso.MoveFile p, p & ".annule"
    On Error GoTo 0
    Application.StatusBar = False

    MsgBox ok & " prix restaures." & IIf(ko <> "", vbCrLf & vbCrLf & "Erreurs:" & ko, ""), _
        IIf(ko = "", vbInformation, vbExclamation), "Annuler derniere modif"

    ' Recharger l'onglet d'ou venait la modification
    If feuille = "" Then feuille = "Recherche article"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(feuille)
    If feuille = "Ventes a perte" Then
        LancerAnalyse ws, False, True
    ElseIf Trim(CStr(ws.Range("B2").Value)) <> "" Then
        LancerAnalyse ws, True, True
    End If
End Sub
'@

try {
    $vbMod = $wb.VBProject.VBComponents.Add(1)
    $vbMod.Name = "ModVentesPerte"
    $vbMod.CodeModule.AddFromString($vbaCode)
    Write-Host "VBA injecte avec succes!"
    # Code de feuille "Recherche article" : surlignage jaune des cellules modifiees
    $sheetCode = @'
' Colonnes : F=6 Lie a, G=7 Facteur, H=8 Prix Net, I=9 Prix base, J=10 % profit voulu, R=18 Prix vente, S=19 Cout
' Originaux caches : 41=F, 42=G, 43=H, 44=R
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim zone As Range, c As Range, r As Long
    On Error GoTo Fin

    ' 1) % profit voulu (J) -> calcule le facteur (G) ou le prix net (R)
    Set zone = Application.Intersect(Target, Me.Range("J6:J100000"))
    If Not zone Is Nothing Then
        If zone.Cells.Count <= 5000 Then
            For Each c In zone.Cells
                r = c.Row
                If Trim(CStr(Me.Cells(r, 3).Value)) <> "" And IsNumeric(c.Value2) And Trim(CStr(c.Value2)) <> "" Then
                    If CStr(Me.Cells(r, 1).Value) <> "Liste de prix" Then
                        MsgBox "Ligne " & r & ": seulement pour les lignes 'Liste de prix'.", vbExclamation
                    ElseIf CStr(Me.Cells(r, 6).Value2) <> CStr(Me.Cells(r, 41).Value2) Then
                        MsgBox "Ligne " & r & ": 'Lie a' a ete modifie, le prix de base (X) n'est plus a jour." & vbCrLf & _
                               "Appliquer le changement de 'Lie a' d'abord, puis entrer le %.", vbExclamation
                    Else
                        Dim cible As Double, base As Double
                        cible = CDbl(Me.Cells(r, 19).Value2) * (1 + CDbl(c.Value2) / 100)
                        If UCase(Trim(CStr(Me.Cells(r, 8).Value2))) = "Y" Then
                            Me.Cells(r, 18).Value = Round(cible, 2)
                        Else
                            base = 0
                            If IsNumeric(Me.Cells(r, 9).Value2) And Trim(CStr(Me.Cells(r, 9).Value2)) <> "" Then base = CDbl(Me.Cells(r, 9).Value2)
                            If base > 0 Then
                                Me.Cells(r, 7).Value = Round(cible / base, 6)
                            Else
                                MsgBox "Ligne " & r & ": prix de base (colonne I) inconnu, facteur non calcule.", vbExclamation
                            End If
                        End If
                    End If
                End If
            Next c
        End If
    End If

    ' 2) Surligne en jaune F-G-H-R modifiees
    Set zone = Application.Intersect(Target, Me.Range("F6:H100000,R6:R100000"))
    If zone Is Nothing Then Exit Sub
    If zone.Cells.Count > 5000 Then Exit Sub
    Dim col As Long, a As Variant, o As Variant, dif As Boolean
    For Each c In zone.Cells
        r = c.Row
        If Trim(CStr(Me.Cells(r, 3).Value)) <> "" Then
            Select Case c.Column
                Case 6: col = 41
                Case 7: col = 42
                Case 8: col = 43
                Case Else: col = 44
            End Select
            a = c.Value2: o = Me.Cells(r, col).Value2
            If IsNumeric(a) And IsNumeric(o) And Trim(CStr(a)) <> "" And Trim(CStr(o)) <> "" Then
                dif = Abs(CDbl(a) - CDbl(o)) > 0.000001
            Else
                dif = UCase(Trim(CStr(a))) <> UCase(Trim(CStr(o)))
            End If
            If dif Then
                c.Interior.Color = RGB(255, 255, 0)
            ElseIf Me.Cells(r, 1).Interior.ColorIndex = xlNone Then
                c.Interior.ColorIndex = xlNone
            Else
                c.Interior.Color = Me.Cells(r, 1).Interior.Color
            End If
        End If
    Next c
Fin:
End Sub
'@
    foreach ($c in $wb.VBProject.VBComponents) {
        if ($c.Type -eq 100) {
            $nm = ""
            try { $nm = $c.Properties.Item("Name").Value } catch {}
            if ($nm -eq "Recherche article" -or $nm -eq "Ventes a perte") { $c.CodeModule.AddFromString($sheetCode) }
        }
    }
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
Write-Host "    - RechercherArticle      : Meme analyse pour un/des articles"
Write-Host "    - AppliquerModifsArticle : Ecrit F-G-H-R dans SAP (liste de la ligne seulement)"
Write-Host "    - AppliquerModifsVentesPerte : Ecrit F-G-H-R dans SAP (liste de la ligne seulement)"
Write-Host "    - AnnulerDerniereModif   : Restaure les prix d'avant la derniere modif"
Write-Host "====================================="
