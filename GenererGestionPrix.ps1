<#
  GenererGestionPrix.ps1 - Genere GestionPrix.xlsm
  Scenario 1: Appliquer facteurs par groupe de produits sur listes de prix clients
#>

$xlsmPath = "U:\GromecOutlook\GestionPrix.xlsm"

if (Test-Path $xlsmPath) { Remove-Item $xlsmPath -Force }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false

$wb = $xl.Workbooks.Add()

while ($wb.Sheets.Count -gt 1) { $wb.Sheets($wb.Sheets.Count).Delete() }

# -- Feuilles --
$wsFact = $wb.Sheets(1)
$wsFact.Name = "Facteurs"
$wsListes = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsFact)
$wsListes.Name = "Listes de prix"
$wsGroupes = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsListes)
$wsGroupes.Name = "Groupes"

# ===== ONGLET FACTEURS =====
# Zone selection
$wsFact.Cells(1,1).Value2 = "APPLIQUER FACTEURS PAR GROUPE DE PRODUITS"
$wsFact.Cells(1,1).Font.Bold = $true
$wsFact.Cells(1,1).Font.Size = 14

$wsFact.Cells(3,1).Value2 = "Liste de prix cible:"
$wsFact.Cells(3,1).Font.Bold = $true
$wsFact.Cells(3,3).Value2 = "(Entrez le # de liste, ex: 4)"
$wsFact.Cells(3,3).Font.Color = 0x808080

$wsFact.Cells(4,1).Value2 = "Liste de base:"
$wsFact.Cells(4,1).Font.Bold = $true
$wsFact.Cells(4,3).Value2 = "(Entrez le # de la liste de base, ex: 1)"
$wsFact.Cells(4,3).Font.Color = 0x808080

$wsFact.Cells(6,1).Value2 = "GROUPES DE PRODUITS ET FACTEURS"
$wsFact.Cells(6,1).Font.Bold = $true
$wsFact.Cells(6,1).Font.Size = 11

$hdrF = @("Code Grp","Nom du groupe","Facteur","Nb articles","Apercu prix base","Apercu prix calcule")
for ($i = 0; $i -lt $hdrF.Count; $i++) {
    $c = $wsFact.Cells(7, $i+1)
    $c.Value2 = $hdrF[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x8B4513
    $c.Font.Color = 0xFFFFFF
}
$wsFact.Columns("A:A").ColumnWidth = 10
$wsFact.Columns("B:B").ColumnWidth = 35
$wsFact.Columns("C:C").ColumnWidth = 12
$wsFact.Columns("D:D").ColumnWidth = 12
$wsFact.Columns("E:E").ColumnWidth = 16
$wsFact.Columns("F:F").ColumnWidth = 18

$wsFact.Cells(3,2).Interior.Color = 0xE0FFFF
$wsFact.Cells(4,2).Interior.Color = 0xE0FFFF

# ===== ONGLET LISTES DE PRIX =====
$hdrL = @("#","Nom","Base #","Facteur","Brut","Devise","Valide","Debut","Fin")
for ($i = 0; $i -lt $hdrL.Count; $i++) {
    $c = $wsListes.Cells(1, $i+1)
    $c.Value2 = $hdrL[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x8B4513
    $c.Font.Color = 0xFFFFFF
}
$wsListes.Columns("A:A").ColumnWidth = 6
$wsListes.Columns("B:B").ColumnWidth = 35
$wsListes.Columns("C:C").ColumnWidth = 8
$wsListes.Columns("D:D").ColumnWidth = 10
$wsListes.Columns("E:E").ColumnWidth = 6
$wsListes.Columns("F:F").ColumnWidth = 8
$wsListes.Columns("G:G").ColumnWidth = 8
$wsListes.Columns("H:I").ColumnWidth = 12

# ===== ONGLET GROUPES =====
$hdrG = @("Code","Nom du groupe","Nb articles")
for ($i = 0; $i -lt $hdrG.Count; $i++) {
    $c = $wsGroupes.Cells(1, $i+1)
    $c.Value2 = $hdrG[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x8B4513
    $c.Font.Color = 0xFFFFFF
}
$wsGroupes.Columns("A:A").ColumnWidth = 8
$wsGroupes.Columns("B:B").ColumnWidth = 35
$wsGroupes.Columns("C:C").ColumnWidth = 12

$wsFact.Activate()

# ========== INJECTION VBA ==========
Write-Host "Injection du code VBA..."

$vbaCode = @'
Option Explicit

Private mSapUser As String
Private mSapPass As String

Private Function ConnectSAP() As Object
    Dim company As Object
    Set company = CreateObject("SAPbobsCOM.Company")

    If mSapUser = "" Then
        mSapUser = InputBox("Nom d'utilisateur SAP:", "Connexion SAP")
        If mSapUser = "" Then
            Set ConnectSAP = Nothing: Exit Function
        End If
    End If

    If mSapPass = "" Then
        mSapPass = InputBox("Mot de passe SAP:", "Connexion SAP")
        If mSapPass = "" Then
            mSapUser = ""
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

' ==========================================
'  CHARGER GROUPES DE PRODUITS
' ==========================================
Public Sub ChargerGroupes()
    Application.StatusBar = "Connexion SAP..."
    Application.ScreenUpdating = False
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Application.ScreenUpdating = True: Exit Sub

    On Error GoTo ErrG

    ' Charger dans onglet Groupes
    Dim wsG As Worksheet: Set wsG = ThisWorkbook.Sheets("Groupes")
    Dim lr As Long
    lr = wsG.Cells(wsG.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then wsG.Range(wsG.Cells(2, 1), wsG.Cells(lr, 3)).ClearContents

    Application.StatusBar = "Chargement groupes de produits..."
    Dim rsG As Object
    Set rsG = RunQuery(co, "SELECT g.""ItmsGrpCod"", g.""ItmsGrpNam"", COUNT(i.""ItemCode"") AS ""NbItems"" FROM ""OITB"" g LEFT JOIN ""OITM"" i ON g.""ItmsGrpCod"" = i.""ItmsGrpCod"" GROUP BY g.""ItmsGrpCod"", g.""ItmsGrpNam"" ORDER BY g.""ItmsGrpCod""")
    Dim cnt As Long: cnt = RC(rsG)
    Dim r As Long
    For r = 1 To cnt
        wsG.Cells(1 + r, 1).Value = GF(rsG, "ItmsGrpCod")
        wsG.Cells(1 + r, 2).Value = GF(rsG, "ItmsGrpNam")
        wsG.Cells(1 + r, 3).Value = GF(rsG, "NbItems")
        MN rsG
    Next r

    ' Aussi charger dans onglet Facteurs (lignes 8+)
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("Facteurs")
    lr = wsF.Cells(wsF.Rows.Count, 1).End(xlUp).Row
    If lr > 7 Then wsF.Range(wsF.Cells(8, 1), wsF.Cells(lr, 6)).ClearContents

    For r = 1 To cnt
        wsF.Cells(7 + r, 1).Value = wsG.Cells(1 + r, 1).Value
        wsF.Cells(7 + r, 2).Value = wsG.Cells(1 + r, 2).Value
        wsF.Cells(7 + r, 3).Value = ""
        wsF.Cells(7 + r, 3).Interior.Color = 15792383
        wsF.Cells(7 + r, 4).Value = wsG.Cells(1 + r, 3).Value
    Next r

    Application.StatusBar = cnt & " groupes charges - Entrez les facteurs en colonne C"
    GoTo CleanG
ErrG:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanG:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.ScreenUpdating = True
End Sub

' ==========================================
'  CHARGER LISTES DE PRIX
' ==========================================
Public Sub ChargerListes()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Listes de prix")
    Application.StatusBar = "Connexion SAP..."
    Application.ScreenUpdating = False
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Application.ScreenUpdating = True: Exit Sub

    On Error GoTo ErrL

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then ws.Range(ws.Cells(2, 1), ws.Cells(lr, 9)).ClearContents

    Dim rsL As Object
    Set rsL = RunQuery(co, "SELECT ""ListNum"",""ListName"",""BASE_NUM"",""Factor"",""IsGrossPrc"",""PrimCurr"",""ValidFor"",""ValidFrom"",""ValidTo"" FROM ""OPLN"" ORDER BY ""ListNum""")
    Dim cnt As Long: cnt = RC(rsL)
    Dim r As Long
    For r = 1 To cnt
        ws.Cells(1 + r, 1).Value = GF(rsL, "ListNum")
        ws.Cells(1 + r, 2).Value = GF(rsL, "ListName")
        ws.Cells(1 + r, 3).Value = GF(rsL, "BASE_NUM")
        ws.Cells(1 + r, 4).Value = GF(rsL, "Factor")
        ws.Cells(1 + r, 5).Value = GF(rsL, "IsGrossPrc")
        ws.Cells(1 + r, 6).Value = GF(rsL, "PrimCurr")
        ws.Cells(1 + r, 7).Value = GF(rsL, "ValidFor")
        ws.Cells(1 + r, 8).Value = GF(rsL, "ValidFrom")
        ws.Cells(1 + r, 9).Value = GF(rsL, "ValidTo")
        MN rsL
    Next r

    Application.StatusBar = cnt & " listes chargees"
    GoTo CleanL
ErrL:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanL:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' ==========================================
'  APERCU - Calculer prix avant application
' ==========================================
Public Sub ApercuFacteurs()
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("Facteurs")
    Dim listeCible As Variant: listeCible = wsF.Cells(3, 2).Value
    Dim listeBase As Variant: listeBase = wsF.Cells(4, 2).Value

    If listeCible = "" Or listeBase = "" Then
        MsgBox "Entrez le # de liste cible (B3) et liste de base (B4)", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "Connexion SAP pour apercu..."
    Application.ScreenUpdating = False
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Application.ScreenUpdating = True: Exit Sub

    On Error GoTo ErrAp

    Dim r As Long, grpCode As Variant, facteur As Double
    r = 8
    Do While wsF.Cells(r, 1).Value <> ""
        facteur = 0
        If wsF.Cells(r, 3).Value <> "" Then
            facteur = CDbl(wsF.Cells(r, 3).Value)
        End If

        If facteur <> 0 Then
            grpCode = wsF.Cells(r, 1).Value
            Dim rsAp As Object
            Set rsAp = RunQuery(co, "SELECT AVG(p.""Price"") AS ""AvgBase"" FROM ""ITM1"" p INNER JOIN ""OITM"" i ON p.""ItemCode"" = i.""ItemCode"" WHERE i.""ItmsGrpCod"" = " & grpCode & " AND p.""PriceList"" = " & listeBase & " AND p.""Price"" > 0")
            If RC(rsAp) > 0 Then
                Dim avgBase As Double
                avgBase = CDbl("0" & GF(rsAp, "AvgBase"))
                wsF.Cells(r, 5).Value = Format(avgBase, "#,##0.00")
                wsF.Cells(r, 6).Value = Format(avgBase * facteur, "#,##0.00")
            End If
            Set rsAp = Nothing
        Else
            wsF.Cells(r, 5).Value = ""
            wsF.Cells(r, 6).Value = ""
        End If
        r = r + 1
    Loop

    Application.StatusBar = "Apercu calcule"
    GoTo CleanAp
ErrAp:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanAp:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' ==========================================
'  APPLIQUER FACTEURS - Via DI API oItems
' ==========================================
Public Sub AppliquerFacteurs()
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("Facteurs")
    Dim listeCible As Variant: listeCible = wsF.Cells(3, 2).Value
    Dim listeBase As Variant: listeBase = wsF.Cells(4, 2).Value

    If listeCible = "" Or listeBase = "" Then
        MsgBox "Entrez le # de liste cible (B3) et liste de base (B4)", vbExclamation
        Exit Sub
    End If

    ' Collecter groupes avec facteurs
    Dim grpCodes() As String, grpFactors() As Double
    Dim nbGrp As Long: nbGrp = 0
    Dim r As Long: r = 8
    Do While wsF.Cells(r, 1).Value <> ""
        If wsF.Cells(r, 3).Value <> "" And CDbl("0" & wsF.Cells(r, 3).Value) <> 0 Then
            nbGrp = nbGrp + 1
            ReDim Preserve grpCodes(1 To nbGrp)
            ReDim Preserve grpFactors(1 To nbGrp)
            grpCodes(nbGrp) = CStr(wsF.Cells(r, 1).Value)
            grpFactors(nbGrp) = CDbl(wsF.Cells(r, 3).Value)
        End If
        r = r + 1
    Loop

    If nbGrp = 0 Then
        MsgBox "Aucun facteur saisi. Entrez des facteurs en colonne C.", vbExclamation
        Exit Sub
    End If

    ' Compter articles concernes
    Application.StatusBar = "Comptage des articles..."
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Exit Sub

    On Error GoTo ErrApp

    Dim inList As String, g As Long
    inList = ""
    For g = 1 To nbGrp
        If inList <> "" Then inList = inList & ","
        inList = inList & grpCodes(g)
    Next g

    Dim rsCount As Object
    Set rsCount = RunQuery(co, "SELECT COUNT(*) AS ""Total"" FROM ""OITM"" WHERE ""ItmsGrpCod"" IN (" & inList & ")")
    Dim totalItems As Long: totalItems = CLng(GF(rsCount, "Total"))

    ' Resume et confirmation
    Dim msg As String
    msg = "RESUME DE L'OPERATION" & vbCrLf & vbCrLf
    msg = msg & "Liste cible: " & listeCible & vbCrLf
    msg = msg & "Liste de base: " & listeBase & vbCrLf
    msg = msg & "Articles concernes: " & totalItems & vbCrLf & vbCrLf
    For g = 1 To nbGrp
        msg = msg & "Groupe " & grpCodes(g) & " -> facteur " & grpFactors(g) & vbCrLf
    Next g
    msg = msg & vbCrLf & "CONFIRMER l'application?"

    If MsgBox(msg, vbYesNo + vbQuestion, "Confirmer") <> vbYes Then
        GoTo CleanApp
    End If

    ' Appliquer via DI API - Update Items prices
    Dim cntOK As Long, cntErr As Long
    cntOK = 0: cntErr = 0

    For g = 1 To nbGrp
        Application.StatusBar = "Groupe " & grpCodes(g) & " - Chargement articles..."
        Dim rsItems As Object
        Set rsItems = RunQuery(co, "SELECT ""ItemCode"" FROM ""OITM"" WHERE ""ItmsGrpCod"" = " & grpCodes(g) & " AND LEFT(""ItemCode"",2) NOT IN ('PT','PR','TA','TB','TC','TD')")
        Dim nbItems As Long: nbItems = RC(rsItems)
        Dim i As Long

        For i = 1 To nbItems
            Dim itemCode As String
            itemCode = CStr(GF(rsItems, "ItemCode"))
            Application.StatusBar = "Groupe " & grpCodes(g) & " - " & i & "/" & nbItems & " - " & itemCode

            ' Charger l'item via DI API (oItems = 4)
            Dim oItem As Object
            Set oItem = co.GetBusinessObject(4)
            Dim found As Boolean
            found = CallByName(oItem, "GetByKey", VbMethod, itemCode)

            If found Then
                ' Acceder a Prices (collection de lignes de prix)
                Dim oPrices As Object
                Set oPrices = CallByName(oItem, "PriceList", VbGet)

                ' Trouver la ligne de la liste cible via SetCurrentLine
                CallByName oPrices, "SetCurrentLine", VbMethod, CLng(listeCible) - 1

                ' Setter BasePriceList et Factor
                CallByName oPrices, "BasePriceList", VbLet, CLng(listeBase)
                CallByName oPrices, "Factor", VbLet, grpFactors(g)

                Dim retU As Long
                retU = CallByName(oItem, "Update", VbMethod)
                If retU = 0 Then
                    cntOK = cntOK + 1
                Else
                    cntErr = cntErr + 1
                End If
            Else
                cntErr = cntErr + 1
            End If

            Set oItem = Nothing
            MN rsItems
            DoEvents
        Next i
        Set rsItems = Nothing
    Next g

    MsgBox "Termine!" & vbCrLf & "Succes: " & cntOK & vbCrLf & "Erreurs: " & cntErr, vbInformation
    Application.StatusBar = "Termine - " & cntOK & " articles mis a jour, " & cntErr & " erreurs"
    GoTo CleanApp

ErrApp:
    MsgBox "Erreur: " & Err.Description, vbCritical
    Application.StatusBar = False
CleanApp:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.ScreenUpdating = True
End Sub
'@

try {
    $vbMod = $wb.VBProject.VBComponents.Add(1)
    $vbMod.Name = "ModGestionPrix"
    $vbMod.CodeModule.AddFromString($vbaCode)
    Write-Host "VBA injecte avec succes!"
} catch {
    Write-Host "ERREUR injection VBA: $($_.Exception.Message)"
    Write-Host "Verifiez que 'Acces approuve au modele objet VBA' est active dans Excel."
    $xl.Quit()
    exit 1
}

# Creer les boutons ActiveX sur l'onglet Facteurs
Write-Host "Creation des boutons..."

try {
    $m = [System.Reflection.Missing]::Value
    $btn1 = $wsFact.OLEObjects().Add("Forms.CommandButton.1", $m, $false, $false, $m, $m, $m, 560, 10, 140, 30)
    $btn1.Object.Caption = "Charger les donnees"
    $btn1.Object.Font.Size = 10
    Write-Host "  Bouton 1 cree"

    $btn2 = $wsFact.OLEObjects().Add("Forms.CommandButton.1", $m, $false, $false, $m, $m, $m, 560, 50, 140, 30)
    $btn2.Object.Caption = "Apercus"
    $btn2.Object.Font.Size = 10
    Write-Host "  Bouton 2 cree"

    $btn3 = $wsFact.OLEObjects().Add("Forms.CommandButton.1", $m, $false, $false, $m, $m, $m, 560, 90, 140, 30)
    $btn3.Object.Caption = "Appliquer facteur"
    $btn3.Object.Font.Size = 10
    Write-Host "  Bouton 3 cree"

    $vbSheet = $wb.VBProject.VBComponents($wsFact.CodeName)
    $codeMod = $vbSheet.CodeModule
    $lines = $codeMod.CountOfLines
    if ($lines -gt 0) { $codeMod.DeleteLines(1, $lines) }
    $codeMod.AddFromString(
@'
Private Sub CommandButton1_Click()
    Call ChargerGroupes
    Call ChargerListes
End Sub

Private Sub CommandButton2_Click()
    Call ApercuFacteurs
End Sub

Private Sub CommandButton3_Click()
    Call AppliquerFacteurs
End Sub
'@
    )
    Write-Host "  Code VBA injecte"
    Write-Host "Boutons crees avec succes!"
} catch {
    Write-Host "AVERT Boutons non crees: $($_.Exception.Message) -- utilisez Alt+F8"
}

# Sauvegarder en xlsm
$wb.SaveAs($xlsmPath, 52)
Write-Host "Fichier sauvegarde: $xlsmPath"
$wb.Close($false)
$xl.Quit()

Write-Host ""
Write-Host "====================================="
Write-Host "  GestionPrix.xlsm cree avec succes!"
Write-Host "  Ouvrez: $xlsmPath"
Write-Host "  Macros (Alt+F8):"
Write-Host "    - ChargerGroupes : Charge les groupes de produits"
Write-Host "    - ChargerListes  : Charge les listes de prix"
Write-Host "    - ApercuFacteurs : Calcule un apercu avant application"
Write-Host "    - AppliquerFacteurs : Applique les facteurs via DI API"
Write-Host "====================================="
