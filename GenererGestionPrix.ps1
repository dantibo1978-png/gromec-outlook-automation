<#
  GenererGestionPrix.ps1 - Genere GestionPrix.xlsm (une seule fois)
  Apres generation, ouvrir GestionPrix.xlsm directement - tout est dedans.
#>

$xlsmPath = "U:\GromecOutlook\GestionPrix.xlsm"

# Supprimer ancien fichier
if (Test-Path $xlsmPath) { Remove-Item $xlsmPath -Force }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false

$wb = $xl.Workbooks.Add()

# Supprimer feuilles en trop
while ($wb.Sheets.Count -gt 1) { $wb.Sheets($wb.Sheets.Count).Delete() }

$wsClient = $wb.Sheets(1)
$wsClient.Name = "Par client"
$wsArticle = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsClient)
$wsArticle.Name = "Par article"
$wsListes = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsArticle)
$wsListes.Name = "Listes de prix"

# -- Onglet Par client --
$wsClient.Cells(1,1).Value2 = "Code client:"
$wsClient.Cells(2,1).Value2 = "Nom:"
$wsClient.Cells(1,1).Font.Bold = $true
$wsClient.Cells(2,1).Font.Bold = $true
$hdrC = @("Article","Description","Prix","Devise","Escompte %","Liste #","Auto MAJ","Valide","Debut","Fin")
for ($i = 0; $i -lt $hdrC.Count; $i++) {
    $c = $wsClient.Cells(4, $i+1)
    $c.Value2 = $hdrC[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x2E231B
    $c.Font.Color = 0xAE9B8C
}
$wsClient.Columns("A:A").ColumnWidth = 16
$wsClient.Columns("B:B").ColumnWidth = 35
$wsClient.Columns("C:C").ColumnWidth = 12
$wsClient.Columns("D:D").ColumnWidth = 8
$wsClient.Columns("E:E").ColumnWidth = 12
$wsClient.Columns("F:F").ColumnWidth = 8
$wsClient.Columns("G:G").ColumnWidth = 10
$wsClient.Columns("H:H").ColumnWidth = 8
$wsClient.Columns("I:J").ColumnWidth = 12

# -- Onglet Par article --
$wsArticle.Cells(1,1).Value2 = "Code article:"
$wsArticle.Cells(2,1).Value2 = "Description:"
$wsArticle.Cells(1,1).Font.Bold = $true
$wsArticle.Cells(2,1).Font.Bold = $true
$wsArticle.Cells(3,1).Value2 = "PRIX PAR LISTE"
$wsArticle.Cells(3,1).Font.Bold = $true
$hdrA = @("#","Liste","Prix","Devise","Base #","Facteur","Brut")
for ($i = 0; $i -lt $hdrA.Count; $i++) {
    $c = $wsArticle.Cells(4, $i+1)
    $c.Value2 = $hdrA[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x2E231B
    $c.Font.Color = 0xAE9B8C
}
$wsArticle.Columns("A:A").ColumnWidth = 6
$wsArticle.Columns("B:B").ColumnWidth = 28
$wsArticle.Columns("C:C").ColumnWidth = 12
$wsArticle.Columns("D:D").ColumnWidth = 8
$wsArticle.Columns("E:E").ColumnWidth = 8
$wsArticle.Columns("F:F").ColumnWidth = 10
$wsArticle.Columns("G:G").ColumnWidth = 6

# -- Onglet Listes de prix --
$hdrL = @("#","Nom","Base #","Facteur","Brut","Devise","Valide","Debut","Fin")
for ($i = 0; $i -lt $hdrL.Count; $i++) {
    $c = $wsListes.Cells(1, $i+1)
    $c.Value2 = $hdrL[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x2E231B
    $c.Font.Color = 0xAE9B8C
}
$wsListes.Columns("A:A").ColumnWidth = 6
$wsListes.Columns("B:B").ColumnWidth = 30
$wsListes.Columns("C:C").ColumnWidth = 8
$wsListes.Columns("D:D").ColumnWidth = 10
$wsListes.Columns("E:E").ColumnWidth = 6
$wsListes.Columns("F:F").ColumnWidth = 8
$wsListes.Columns("G:G").ColumnWidth = 8
$wsListes.Columns("H:I").ColumnWidth = 12

# -- Boutons --
$btn1 = $wsClient.Buttons.Add($wsClient.Cells(1,4).Left, $wsClient.Cells(1,4).Top, 90, 22)
$btn1.Text = "Rechercher"
$btn1.OnAction = "RechercherClient"

$btn2 = $wsClient.Buttons.Add($wsClient.Cells(1,6).Left, $wsClient.Cells(1,6).Top, 90, 22)
$btn2.Text = "Sauvegarder"
$btn2.OnAction = "SauvegarderClient"

$btn3 = $wsArticle.Buttons.Add($wsArticle.Cells(1,4).Left, $wsArticle.Cells(1,4).Top, 90, 22)
$btn3.Text = "Rechercher"
$btn3.OnAction = "RechercherArticle"

$btn4 = $wsArticle.Buttons.Add($wsArticle.Cells(1,6).Left, $wsArticle.Cells(1,6).Top, 90, 22)
$btn4.Text = "Sauvegarder"
$btn4.OnAction = "SauvegarderArticle"

$btn5 = $wsListes.Buttons.Add($wsListes.Cells(1,10).Left, $wsListes.Cells(1,10).Top, 110, 22)
$btn5.Text = "Charger listes"
$btn5.OnAction = "ChargerListes"

$wsClient.Activate()

# ========== INJECTION VBA ==========
Write-Host "Injection du code VBA..."

$vbaCode = @'
Option Explicit

Private Function ConnectSAP() As Object
    Dim company As Object
    Set company = CreateObject("SAPbobsCOM.Company")

    Dim fso As Object, f As Object, jsonText As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim cfgPath As String
    cfgPath = "U:\GromecOutlook\config.json"
    If fso.FileExists(cfgPath) Then
        Set f = fso.OpenTextFile(cfgPath, 1)
        jsonText = f.ReadAll
        f.Close
    Else
        MsgBox "config.json introuvable: " & cfgPath, vbCritical
        Set ConnectSAP = Nothing: Exit Function
    End If

    company.Server = "NDB@hana:30013"
    company.DbServerType = 9
    company.CompanyDB = "SAP_GROMEC_PROD"
    company.UseTrusted = False
    company.UserName = ExtractJsonValue(jsonText, "sap_user")
    company.Password = ExtractJsonValue(jsonText, "sap_password")

    If company.UserName = "" Then
        MsgBox "sap_user manquant dans config.json", vbCritical
        Set ConnectSAP = Nothing: Exit Function
    End If

    Dim ret As Long
    ret = company.Connect()
    If ret <> 0 Then
        MsgBox "Connexion SAP echouee (" & ret & "): " & company.GetLastErrorDescription(), vbCritical
        Set ConnectSAP = Nothing: Exit Function
    End If
    Set ConnectSAP = company
End Function

Private Function ExtractJsonValue(jsonText As String, key As String) As String
    Dim pos As Long, pos2 As Long, pos3 As Long
    pos = InStr(1, jsonText, Chr(34) & key & Chr(34))
    If pos = 0 Then ExtractJsonValue = "": Exit Function
    pos2 = InStr(pos, jsonText, ":")
    If pos2 = 0 Then ExtractJsonValue = "": Exit Function
    pos2 = InStr(pos2, jsonText, Chr(34))
    pos3 = InStr(pos2 + 1, jsonText, Chr(34))
    ExtractJsonValue = Mid(jsonText, pos2 + 1, pos3 - pos2 - 1)
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

Public Sub RechercherClient()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Par client")
    Dim cd As String: cd = Trim(ws.Cells(1, 2).Value)
    If cd = "" Then MsgBox "Entrez un code client en B1", vbExclamation: Exit Sub

    Application.StatusBar = "Connexion SAP..."
    Application.ScreenUpdating = False
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Application.ScreenUpdating = True: Exit Sub

    On Error GoTo ErrC

    Application.StatusBar = "Recherche client..."
    Dim ri As Object
    Set ri = RunQuery(co, "SELECT ""CardCode"",""CardName"" FROM ""OCRD"" WHERE ""CardCode"" = '" & Replace(cd, "'", "''") & "'")
    If RC(ri) = 0 Then
        MsgBox "Client introuvable: " & cd, vbExclamation
        GoTo CleanC
    End If
    ws.Cells(2, 2).Value = GF(ri, "CardName")
    Set ri = Nothing

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr > 4 Then ws.Range(ws.Cells(5, 1), ws.Cells(lr, 10)).ClearContents

    Application.StatusBar = "Chargement prix speciaux..."
    Dim rsSP As Object
    Set rsSP = RunQuery(co, "SELECT s.""ItemCode"", i.""ItemName"", s.""Price"", s.""Currency"", s.""Discount"", s.""ListNum"", s.""AutoUpdt"", s.""Valid"", s.""ValidFrom"", s.""ValidTo"" FROM ""OSPP"" s LEFT JOIN ""OITM"" i ON s.""ItemCode"" = i.""ItemCode"" WHERE s.""CardCode"" = '" & Replace(cd, "'", "''") & "' ORDER BY s.""ItemCode""")
    Dim cnt As Long: cnt = RC(rsSP)
    Dim r As Long
    For r = 1 To cnt
        If r > 5000 Then Exit For
        ws.Cells(4 + r, 1).Value = GF(rsSP, "ItemCode")
        ws.Cells(4 + r, 2).Value = GF(rsSP, "ItemName")
        ws.Cells(4 + r, 3).Value = GF(rsSP, "Price")
        ws.Cells(4 + r, 4).Value = GF(rsSP, "Currency")
        ws.Cells(4 + r, 5).Value = GF(rsSP, "Discount")
        ws.Cells(4 + r, 6).Value = GF(rsSP, "ListNum")
        ws.Cells(4 + r, 7).Value = GF(rsSP, "AutoUpdt")
        ws.Cells(4 + r, 8).Value = GF(rsSP, "Valid")
        ws.Cells(4 + r, 9).Value = GF(rsSP, "ValidFrom")
        ws.Cells(4 + r, 10).Value = GF(rsSP, "ValidTo")
        MN rsSP
    Next r

    Application.StatusBar = cnt & " prix speciaux charges"
    GoTo CleanC
ErrC:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanC:
    On Error Resume Next
    co.Disconnect
    Set co = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub SauvegarderClient()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Par client")
    Dim cd As String: cd = Trim(ws.Cells(1, 2).Value)
    If cd = "" Then Exit Sub

    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Exit Sub

    On Error GoTo ErrSC
    Dim r As Long, cnt As Long, sql As String
    r = 5
    Do While ws.Cells(r, 1).Value <> ""
        Dim ic As String: ic = ws.Cells(r, 1).Value
        sql = "UPDATE ""OSPP"" SET " _
            & """Price"" = " & ws.Cells(r, 3).Value _
            & ", ""Discount"" = " & ws.Cells(r, 5).Value _
            & ", ""ListNum"" = " & ws.Cells(r, 6).Value
        If ws.Cells(r, 10).Value <> "" Then sql = sql & ", ""ValidTo"" = '" & Format(ws.Cells(r, 10).Value, "yyyy-MM-dd") & "'"
        sql = sql & " WHERE ""CardCode"" = '" & Replace(cd, "'", "''") & "' AND ""ItemCode"" = '" & Replace(ic, "'", "''") & "'"
        Dim rsU As Object
        Set rsU = RunQuery(co, sql)
        cnt = cnt + 1
        r = r + 1
    Loop
    MsgBox cnt & " prix mis a jour", vbInformation
    GoTo CleanSC
ErrSC:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanSC:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
End Sub

Public Sub RechercherArticle()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Par article")
    Dim cd As String: cd = Trim(ws.Cells(1, 2).Value)
    If cd = "" Then MsgBox "Entrez un code article en B1", vbExclamation: Exit Sub

    Application.StatusBar = "Connexion SAP..."
    Application.ScreenUpdating = False
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Application.ScreenUpdating = True: Exit Sub

    On Error GoTo ErrA

    Dim ri As Object
    Set ri = RunQuery(co, "SELECT ""ItemCode"",""ItemName"" FROM ""OITM"" WHERE ""ItemCode"" = '" & Replace(cd, "'", "''") & "'")
    If RC(ri) = 0 Then
        MsgBox "Article introuvable: " & cd, vbExclamation
        GoTo CleanA
    End If
    ws.Cells(2, 2).Value = GF(ri, "ItemName")
    Set ri = Nothing

    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr > 4 Then ws.Range(ws.Cells(5, 1), ws.Cells(lr, 7)).ClearContents

    Application.StatusBar = "Chargement prix par liste..."
    Dim rsPL As Object
    Set rsPL = RunQuery(co, "SELECT p.""PriceList"", l.""ListName"", p.""Price"", p.""Currency"", l.""BASE_NUM"", l.""Factor"", l.""IsGrossPrc"" FROM ""ITM1"" p LEFT JOIN ""OPLN"" l ON p.""PriceList"" = l.""ListNum"" WHERE p.""ItemCode"" = '" & Replace(cd, "'", "''") & "' ORDER BY p.""PriceList""")
    Dim cnt As Long: cnt = RC(rsPL)
    Dim r As Long
    For r = 1 To cnt
        ws.Cells(4 + r, 1).Value = GF(rsPL, "PriceList")
        ws.Cells(4 + r, 2).Value = GF(rsPL, "ListName")
        ws.Cells(4 + r, 3).Value = GF(rsPL, "Price")
        ws.Cells(4 + r, 4).Value = GF(rsPL, "Currency")
        ws.Cells(4 + r, 5).Value = GF(rsPL, "BASE_NUM")
        ws.Cells(4 + r, 6).Value = GF(rsPL, "Factor")
        ws.Cells(4 + r, 7).Value = GF(rsPL, "IsGrossPrc")
        MN rsPL
    Next r

    Application.StatusBar = cnt & " listes de prix chargees"
    GoTo CleanA
ErrA:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanA:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub SauvegarderArticle()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Par article")
    Dim cd As String: cd = Trim(ws.Cells(1, 2).Value)
    If cd = "" Then Exit Sub

    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Exit Sub

    On Error GoTo ErrSA
    Dim r As Long, cnt As Long
    r = 5
    Do While ws.Cells(r, 1).Value <> ""
        Dim ln As String: ln = ws.Cells(r, 1).Value
        Dim rsU As Object
        Set rsU = RunQuery(co, "UPDATE ""ITM1"" SET ""Price"" = " & ws.Cells(r, 3).Value & " WHERE ""ItemCode"" = '" & Replace(cd, "'", "''") & "' AND ""PriceList"" = " & ln)
        cnt = cnt + 1
        r = r + 1
    Loop
    MsgBox cnt & " prix de liste mis a jour", vbInformation
    GoTo CleanSA
ErrSA:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanSA:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
End Sub

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

# Sauvegarder en xlsm
$wb.SaveAs($xlsmPath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
Write-Host "Fichier sauvegarde: $xlsmPath"
$wb.Close($false)
$xl.Quit()

Write-Host ""
Write-Host "====================================="
Write-Host "  GestionPrix.xlsm cree avec succes!"
Write-Host "  Ouvrez: $xlsmPath"
Write-Host "====================================="
