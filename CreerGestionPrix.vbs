' ============================================================================
' CreerGestionPrix.vbs
' Genere le fichier Excel "GestionPrix.xlsm" avec macros VBA integrees
' pour gerer les prix SAP directement via DI API (zero latence).
' Usage: double-cliquer sur ce fichier ou: cscript CreerGestionPrix.vbs
' ============================================================================
Option Explicit

Dim xlApp, wb, wsClient, wsArticle, wsListes, wsConfig
Dim fso, basePath, xlsmPath

basePath = "U:\GromecOutlook\"
xlsmPath = basePath & "GestionPrix.xlsm"

' Supprimer l'ancien fichier s'il existe
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(xlsmPath) Then fso.DeleteFile xlsmPath, True

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set wb = xlApp.Workbooks.Add

' ── Creer les onglets ──
Set wsClient = wb.Sheets(1)
wsClient.Name = "Par client"

Set wsArticle = wb.Sheets.Add(, wsClient)
wsArticle.Name = "Par article"

Set wsListes = wb.Sheets.Add(, wsArticle)
wsListes.Name = "Listes de prix"

Set wsConfig = wb.Sheets.Add(, wsListes)
wsConfig.Name = "Config"

' ── Onglet Config (parametres connexion) ──
wsConfig.Cells(1,1).Value = "Parametres SAP (lus depuis config.json)"
wsConfig.Cells(2,1).Value = "Server:"
wsConfig.Cells(2,2).Value = "NDB@hana:30013"
wsConfig.Cells(3,1).Value = "DbType:"
wsConfig.Cells(3,2).Value = 9
wsConfig.Cells(4,1).Value = "CompanyDB:"
wsConfig.Cells(4,2).Value = "SAP_GROMEC_PROD"
wsConfig.Cells(5,1).Value = "UseTrusted:"
wsConfig.Cells(5,2).Value = False
wsConfig.Cells(6,1).Value = "User:"
wsConfig.Cells(6,2).Value = ""
wsConfig.Cells(7,1).Value = "Password:"
wsConfig.Cells(7,2).Value = ""
wsConfig.Cells(9,1).Value = "Les credentials sont lus de config.json au demarrage."
wsConfig.Cells(10,1).Value = "Ne pas entrer de mot de passe ici."
wsConfig.Columns("A:A").ColumnWidth = 15
wsConfig.Columns("B:B").ColumnWidth = 30

' ── Onglet Par client ──
wsClient.Cells(1,1).Value = "Code client:"
wsClient.Cells(1,3).Value = "Rechercher"
wsClient.Cells(2,1).Value = "Nom:"
' Headers ligne 4
Dim hdrC: hdrC = Array("Article","Description","Prix","Devise","Escompte %","Liste #","Auto MAJ","Valide","Debut","Fin")
Dim i
For i = 0 To UBound(hdrC)
    wsClient.Cells(4, i+1).Value = hdrC(i)
    wsClient.Cells(4, i+1).Font.Bold = True
    wsClient.Cells(4, i+1).Interior.Color = RGB(27,35,46)
    wsClient.Cells(4, i+1).Font.Color = RGB(140,155,174)
Next
wsClient.Columns("A:A").ColumnWidth = 15
wsClient.Columns("B:B").ColumnWidth = 35
wsClient.Columns("C:C").ColumnWidth = 12
wsClient.Columns("D:D").ColumnWidth = 8
wsClient.Columns("E:E").ColumnWidth = 12
wsClient.Columns("I:J").ColumnWidth = 12

' Bouton rechercher
Dim btnSearch
Set btnSearch = wsClient.Buttons.Add(wsClient.Cells(1,3).Left, wsClient.Cells(1,3).Top, 80, 22)
btnSearch.Text = "Rechercher"
btnSearch.OnAction = "RechercherClient"

' Bouton sauvegarder
Dim btnSaveC
Set btnSaveC = wsClient.Buttons.Add(wsClient.Cells(1,5).Left, wsClient.Cells(1,5).Top, 100, 22)
btnSaveC.Text = "Sauvegarder"
btnSaveC.OnAction = "SauvegarderClient"

' ── Onglet Par article ──
wsArticle.Cells(1,1).Value = "Code article:"
wsArticle.Cells(1,3).Value = "Rechercher"
wsArticle.Cells(2,1).Value = "Description:"
wsArticle.Cells(4,1).Value = "PRIX PAR LISTE"
wsArticle.Cells(4,1).Font.Bold = True
Dim hdrPL: hdrPL = Array("#","Liste","Prix","Devise","Base #","Facteur","Brut")
For i = 0 To UBound(hdrPL)
    wsArticle.Cells(5, i+1).Value = hdrPL(i)
    wsArticle.Cells(5, i+1).Font.Bold = True
Next

Set btnSearch = wsArticle.Buttons.Add(wsArticle.Cells(1,3).Left, wsArticle.Cells(1,3).Top, 80, 22)
btnSearch.Text = "Rechercher"
btnSearch.OnAction = "RechercherArticle"

Dim btnSaveA
Set btnSaveA = wsArticle.Buttons.Add(wsArticle.Cells(1,5).Left, wsArticle.Cells(1,5).Top, 100, 22)
btnSaveA.Text = "Sauvegarder"
btnSaveA.OnAction = "SauvegarderArticle"

wsArticle.Columns("A:A").ColumnWidth = 8
wsArticle.Columns("B:B").ColumnWidth = 25
wsArticle.Columns("C:C").ColumnWidth = 12

' ── Onglet Listes de prix ──
wsListes.Cells(1,1).Value = "Charger toutes les listes"
Set btnSearch = wsListes.Buttons.Add(wsListes.Cells(1,1).Left, wsListes.Cells(1,1).Top, 140, 22)
btnSearch.Text = "Charger listes"
btnSearch.OnAction = "ChargerListes"
Dim hdrL: hdrL = Array("#","Nom","Base #","Facteur","Brut","Devise","Valide","Debut","Fin")
For i = 0 To UBound(hdrL)
    wsListes.Cells(3, i+1).Value = hdrL(i)
    wsListes.Cells(3, i+1).Font.Bold = True
Next

' ── Ajouter le code VBA ──
Dim vbProj, vbMod
Set vbProj = wb.VBProject
Set vbMod = vbProj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
vbMod.Name = "ModGestionPrix"

Dim code
code = ""
code = code & "Option Explicit" & vbCrLf & vbCrLf

' ── ConnectSAP ──
code = code & "Private Function ConnectSAP() As Object" & vbCrLf
code = code & "    Dim company As Object" & vbCrLf
code = code & "    Set company = CreateObject(""SAPbobsCOM.Company"")" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    ' Lire config.json" & vbCrLf
code = code & "    Dim fso As Object, f As Object, jsonText As String" & vbCrLf
code = code & "    Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
code = code & "    Dim cfgPath As String" & vbCrLf
code = code & "    cfgPath = ""U:\GromecOutlook\config.json""" & vbCrLf
code = code & "    If fso.FileExists(cfgPath) Then" & vbCrLf
code = code & "        Set f = fso.OpenTextFile(cfgPath, 1)" & vbCrLf
code = code & "        jsonText = f.ReadAll" & vbCrLf
code = code & "        f.Close" & vbCrLf
code = code & "    End If" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    company.Server = ""NDB@hana:30013""" & vbCrLf
code = code & "    company.DbServerType = 9" & vbCrLf
code = code & "    company.CompanyDB = ""SAP_GROMEC_PROD""" & vbCrLf
code = code & "    company.UseTrusted = False" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    ' Parser user/password du JSON" & vbCrLf
code = code & "    Dim usr As String, pwd As String" & vbCrLf
code = code & "    usr = ExtractJsonValue(jsonText, ""sap_user"")" & vbCrLf
code = code & "    pwd = ExtractJsonValue(jsonText, ""sap_password"")" & vbCrLf
code = code & "    If usr = """" Then" & vbCrLf
code = code & "        MsgBox ""Credentials SAP manquants dans config.json"", vbCritical" & vbCrLf
code = code & "        Set ConnectSAP = Nothing: Exit Function" & vbCrLf
code = code & "    End If" & vbCrLf
code = code & "    company.UserName = usr" & vbCrLf
code = code & "    company.Password = pwd" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    Dim ret As Long" & vbCrLf
code = code & "    ret = company.Connect()" & vbCrLf
code = code & "    If ret <> 0 Then" & vbCrLf
code = code & "        MsgBox ""Connexion SAP echouee ("" & ret & ""): "" & company.GetLastErrorDescription(), vbCritical" & vbCrLf
code = code & "        Set ConnectSAP = Nothing: Exit Function" & vbCrLf
code = code & "    End If" & vbCrLf
code = code & "    Set ConnectSAP = company" & vbCrLf
code = code & "End Function" & vbCrLf & vbCrLf

' ── ExtractJsonValue ──
code = code & "Private Function ExtractJsonValue(jsonText As String, key As String) As String" & vbCrLf
code = code & "    Dim pos As Long, pos2 As Long, pos3 As Long" & vbCrLf
code = code & "    pos = InStr(1, jsonText, Chr(34) & key & Chr(34))" & vbCrLf
code = code & "    If pos = 0 Then ExtractJsonValue = """": Exit Function" & vbCrLf
code = code & "    pos2 = InStr(pos, jsonText, "":"")" & vbCrLf
code = code & "    If pos2 = 0 Then ExtractJsonValue = """": Exit Function" & vbCrLf
code = code & "    pos2 = InStr(pos2, jsonText, Chr(34))" & vbCrLf
code = code & "    pos3 = InStr(pos2 + 1, jsonText, Chr(34))" & vbCrLf
code = code & "    ExtractJsonValue = Mid(jsonText, pos2 + 1, pos3 - pos2 - 1)" & vbCrLf
code = code & "End Function" & vbCrLf & vbCrLf

' ── RunQuery ──
code = code & "Private Function RunQuery(company As Object, sql As String) As Object" & vbCrLf
code = code & "    Dim rs As Object" & vbCrLf
code = code & "    Set rs = company.GetBusinessObject(300)" & vbCrLf
code = code & "    CallByName rs, ""DoQuery"", VbMethod, sql" & vbCrLf
code = code & "    Set RunQuery = rs" & vbCrLf
code = code & "End Function" & vbCrLf & vbCrLf

' ── RunUpdate ──
code = code & "Private Sub RunUpdate(company As Object, sql As String)" & vbCrLf
code = code & "    Dim rs As Object" & vbCrLf
code = code & "    Set rs = company.GetBusinessObject(300)" & vbCrLf
code = code & "    CallByName rs, ""DoQuery"", VbMethod, sql" & vbCrLf
code = code & "    Set rs = Nothing" & vbCrLf
code = code & "End Sub" & vbCrLf & vbCrLf

' ── GetFieldValue ──
code = code & "Private Function GetFieldValue(rs As Object, fieldName As String) As Variant" & vbCrLf
code = code & "    On Error Resume Next" & vbCrLf
code = code & "    Dim flds As Object, fld As Object" & vbCrLf
code = code & "    Set flds = CallByName(rs, ""Fields"", VbGet)" & vbCrLf
code = code & "    Set fld = CallByName(flds, ""Item"", VbMethod, fieldName)" & vbCrLf
code = code & "    GetFieldValue = CallByName(fld, ""Value"", VbGet)" & vbCrLf
code = code & "    On Error GoTo 0" & vbCrLf
code = code & "End Function" & vbCrLf & vbCrLf

' ── RecordCount ──
code = code & "Private Function GetRecordCount(rs As Object) As Long" & vbCrLf
code = code & "    GetRecordCount = CLng(CallByName(rs, ""RecordCount"", VbGet))" & vbCrLf
code = code & "End Function" & vbCrLf & vbCrLf

' ── MoveNext ──
code = code & "Private Sub MoveNextRS(rs As Object)" & vbCrLf
code = code & "    CallByName rs, ""MoveNext"", VbMethod" & vbCrLf
code = code & "End Sub" & vbCrLf & vbCrLf

' ── RechercherClient ──
code = code & "Public Sub RechercherClient()" & vbCrLf
code = code & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""Par client"")" & vbCrLf
code = code & "    Dim code As String: code = Trim(ws.Cells(1, 2).Value)" & vbCrLf
code = code & "    If code = """" Then MsgBox ""Entrez un code client"", vbExclamation: Exit Sub" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    Application.StatusBar = ""Connexion SAP..."": Application.ScreenUpdating = False" & vbCrLf
code = code & "    Dim company As Object: Set company = ConnectSAP()" & vbCrLf
code = code & "    If company Is Nothing Then Application.StatusBar = False: Exit Sub" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    On Error GoTo ErrHandler" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    ' Info client" & vbCrLf
code = code & "    Application.StatusBar = ""Recherche client..."" " & vbCrLf
code = code & "    Dim rsInfo As Object" & vbCrLf
code = code & "    Set rsInfo = RunQuery(company, ""SELECT """"CardCode"""",""""CardName"""" FROM """"OCRD"""" WHERE """"CardCode"""" = '"" & Replace(code, ""'"", ""''"") & ""'"")" & vbCrLf
code = code & "    If GetRecordCount(rsInfo) = 0 Then" & vbCrLf
code = code & "        MsgBox ""Client introuvable: "" & code, vbExclamation" & vbCrLf
code = code & "        GoTo Cleanup" & vbCrLf
code = code & "    End If" & vbCrLf
code = code & "    ws.Cells(2, 2).Value = GetFieldValue(rsInfo, ""CardName"")" & vbCrLf
code = code & "    Set rsInfo = Nothing" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    ' Effacer anciennes donnees" & vbCrLf
code = code & "    If ws.Cells(5, 1).Value <> """" Then ws.Range(ws.Cells(5, 1), ws.Cells(ws.Rows.Count, 10).End(3)).ClearContents" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    ' Prix speciaux" & vbCrLf
code = code & "    Application.StatusBar = ""Chargement prix speciaux...""" & vbCrLf
code = code & "    Dim rsSP As Object" & vbCrLf
code = code & "    Set rsSP = RunQuery(company, ""SELECT s.""""ItemCode"""", i.""""ItemName"""", s.""""Price"""", s.""""Currency"""", s.""""Discount"""", s.""""ListNum"""", s.""""AutoUpdt"""", s.""""Valid"""", s.""""ValidFrom"""", s.""""ValidTo"""" FROM """"OSPP"""" s LEFT JOIN """"OITM"""" i ON s.""""ItemCode"""" = i.""""ItemCode"""" WHERE s.""""CardCode"""" = '"" & Replace(code, ""'"", ""''"") & ""' ORDER BY s.""""ItemCode"""""")" & vbCrLf
code = code & "    Dim cnt As Long: cnt = GetRecordCount(rsSP)" & vbCrLf
code = code & "    Dim r As Long" & vbCrLf
code = code & "    For r = 1 To cnt" & vbCrLf
code = code & "        If r > 5000 Then Exit For" & vbCrLf
code = code & "        ws.Cells(4 + r, 1).Value = GetFieldValue(rsSP, ""ItemCode"")" & vbCrLf
code = code & "        ws.Cells(4 + r, 2).Value = GetFieldValue(rsSP, ""ItemName"")" & vbCrLf
code = code & "        ws.Cells(4 + r, 3).Value = GetFieldValue(rsSP, ""Price"")" & vbCrLf
code = code & "        ws.Cells(4 + r, 4).Value = GetFieldValue(rsSP, ""Currency"")" & vbCrLf
code = code & "        ws.Cells(4 + r, 5).Value = GetFieldValue(rsSP, ""Discount"")" & vbCrLf
code = code & "        ws.Cells(4 + r, 6).Value = GetFieldValue(rsSP, ""ListNum"")" & vbCrLf
code = code & "        ws.Cells(4 + r, 7).Value = GetFieldValue(rsSP, ""AutoUpdt"")" & vbCrLf
code = code & "        ws.Cells(4 + r, 8).Value = GetFieldValue(rsSP, ""Valid"")" & vbCrLf
code = code & "        ws.Cells(4 + r, 9).Value = GetFieldValue(rsSP, ""ValidFrom"")" & vbCrLf
code = code & "        ws.Cells(4 + r, 10).Value = GetFieldValue(rsSP, ""ValidTo"")" & vbCrLf
code = code & "        MoveNextRS rsSP" & vbCrLf
code = code & "    Next r" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    Application.StatusBar = cnt & "" prix speciaux charges""" & vbCrLf
code = code & "    GoTo Cleanup" & vbCrLf
code = code & "ErrHandler:" & vbCrLf
code = code & "    MsgBox ""Erreur: "" & Err.Description, vbCritical" & vbCrLf
code = code & "Cleanup:" & vbCrLf
code = code & "    On Error Resume Next" & vbCrLf
code = code & "    company.Disconnect" & vbCrLf
code = code & "    Set company = Nothing" & vbCrLf
code = code & "    Application.ScreenUpdating = True" & vbCrLf
code = code & "    Application.StatusBar = False" & vbCrLf
code = code & "End Sub" & vbCrLf & vbCrLf

' ── SauvegarderClient ──
code = code & "Public Sub SauvegarderClient()" & vbCrLf
code = code & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""Par client"")" & vbCrLf
code = code & "    Dim code As String: code = Trim(ws.Cells(1, 2).Value)" & vbCrLf
code = code & "    If code = """" Then Exit Sub" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    Dim company As Object: Set company = ConnectSAP()" & vbCrLf
code = code & "    If company Is Nothing Then Exit Sub" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    On Error GoTo ErrHandler" & vbCrLf
code = code & "    Dim r As Long, cnt As Long, sql As String" & vbCrLf
code = code & "    r = 5" & vbCrLf
code = code & "    Do While ws.Cells(r, 1).Value <> """"" & vbCrLf
code = code & "        Dim itemCode As String: itemCode = ws.Cells(r, 1).Value" & vbCrLf
code = code & "        sql = ""UPDATE """"OSPP"""" SET "" _" & vbCrLf
code = code & "            & """"""Price"""" = "" & ws.Cells(r, 3).Value _" & vbCrLf
code = code & "            & "", """"Discount"""" = "" & ws.Cells(r, 5).Value _" & vbCrLf
code = code & "            & "", """"ListNum"""" = "" & ws.Cells(r, 6).Value" & vbCrLf
code = code & "        If ws.Cells(r, 9).Value <> """" Then sql = sql & "", """"ValidFrom"""" = '"" & Format(ws.Cells(r, 9).Value, ""yyyy-MM-dd"") & ""'""" & vbCrLf
code = code & "        If ws.Cells(r, 10).Value <> """" Then sql = sql & "", """"ValidTo"""" = '"" & Format(ws.Cells(r, 10).Value, ""yyyy-MM-dd"") & ""'""" & vbCrLf
code = code & "        sql = sql & "" WHERE """"CardCode"""" = '"" & Replace(code, ""'"", ""''"") & ""' AND """"ItemCode"""" = '"" & Replace(itemCode, ""'"", ""''"") & ""'""" & vbCrLf
code = code & "        RunUpdate company, sql" & vbCrLf
code = code & "        cnt = cnt + 1" & vbCrLf
code = code & "        r = r + 1" & vbCrLf
code = code & "    Loop" & vbCrLf
code = code & "    MsgBox cnt & "" prix mis a jour"", vbInformation" & vbCrLf
code = code & "    GoTo Cleanup" & vbCrLf
code = code & "ErrHandler:" & vbCrLf
code = code & "    MsgBox ""Erreur: "" & Err.Description, vbCritical" & vbCrLf
code = code & "Cleanup:" & vbCrLf
code = code & "    On Error Resume Next" & vbCrLf
code = code & "    company.Disconnect: Set company = Nothing" & vbCrLf
code = code & "End Sub" & vbCrLf & vbCrLf

' ── RechercherArticle ──
code = code & "Public Sub RechercherArticle()" & vbCrLf
code = code & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""Par article"")" & vbCrLf
code = code & "    Dim code As String: code = Trim(ws.Cells(1, 2).Value)" & vbCrLf
code = code & "    If code = """" Then MsgBox ""Entrez un code article"", vbExclamation: Exit Sub" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    Application.StatusBar = ""Connexion SAP..."": Application.ScreenUpdating = False" & vbCrLf
code = code & "    Dim company As Object: Set company = ConnectSAP()" & vbCrLf
code = code & "    If company Is Nothing Then Application.StatusBar = False: Exit Sub" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    On Error GoTo ErrHandler" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    ' Info article" & vbCrLf
code = code & "    Dim rsInfo As Object" & vbCrLf
code = code & "    Set rsInfo = RunQuery(company, ""SELECT """"ItemCode"""",""""ItemName"""" FROM """"OITM"""" WHERE """"ItemCode"""" = '"" & Replace(code, ""'"", ""''"") & ""'"")" & vbCrLf
code = code & "    If GetRecordCount(rsInfo) = 0 Then" & vbCrLf
code = code & "        MsgBox ""Article introuvable: "" & code, vbExclamation" & vbCrLf
code = code & "        GoTo Cleanup" & vbCrLf
code = code & "    End If" & vbCrLf
code = code & "    ws.Cells(2, 2).Value = GetFieldValue(rsInfo, ""ItemName"")" & vbCrLf
code = code & "    Set rsInfo = Nothing" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    ' Effacer" & vbCrLf
code = code & "    If ws.Cells(6, 1).Value <> """" Then ws.Range(ws.Cells(6, 1), ws.Cells(ws.Rows.Count, 7).End(3)).ClearContents" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    ' Prix par liste" & vbCrLf
code = code & "    Application.StatusBar = ""Chargement prix par liste...""" & vbCrLf
code = code & "    Dim rsPL As Object" & vbCrLf
code = code & "    Set rsPL = RunQuery(company, ""SELECT p.""""PriceList"""", l.""""ListName"""", p.""""Price"""", p.""""Currency"""", l.""""BASE_NUM"""", l.""""Factor"""", l.""""IsGrossPrc"""" FROM """"ITM1"""" p LEFT JOIN """"OPLN"""" l ON p.""""PriceList"""" = l.""""ListNum"""" WHERE p.""""ItemCode"""" = '"" & Replace(code, ""'"", ""''"") & ""' ORDER BY p.""""PriceList"""""")" & vbCrLf
code = code & "    Dim cnt As Long: cnt = GetRecordCount(rsPL)" & vbCrLf
code = code & "    Dim r As Long" & vbCrLf
code = code & "    For r = 1 To cnt" & vbCrLf
code = code & "        If r > 100 Then Exit For" & vbCrLf
code = code & "        ws.Cells(5 + r, 1).Value = GetFieldValue(rsPL, ""PriceList"")" & vbCrLf
code = code & "        ws.Cells(5 + r, 2).Value = GetFieldValue(rsPL, ""ListName"")" & vbCrLf
code = code & "        ws.Cells(5 + r, 3).Value = GetFieldValue(rsPL, ""Price"")" & vbCrLf
code = code & "        ws.Cells(5 + r, 4).Value = GetFieldValue(rsPL, ""Currency"")" & vbCrLf
code = code & "        ws.Cells(5 + r, 5).Value = GetFieldValue(rsPL, ""BASE_NUM"")" & vbCrLf
code = code & "        ws.Cells(5 + r, 6).Value = GetFieldValue(rsPL, ""Factor"")" & vbCrLf
code = code & "        ws.Cells(5 + r, 7).Value = GetFieldValue(rsPL, ""IsGrossPrc"")" & vbCrLf
code = code & "        MoveNextRS rsPL" & vbCrLf
code = code & "    Next r" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    Application.StatusBar = cnt & "" listes de prix chargees""" & vbCrLf
code = code & "    GoTo Cleanup" & vbCrLf
code = code & "ErrHandler:" & vbCrLf
code = code & "    MsgBox ""Erreur: "" & Err.Description, vbCritical" & vbCrLf
code = code & "Cleanup:" & vbCrLf
code = code & "    On Error Resume Next" & vbCrLf
code = code & "    company.Disconnect: Set company = Nothing" & vbCrLf
code = code & "    Application.ScreenUpdating = True" & vbCrLf
code = code & "    Application.StatusBar = False" & vbCrLf
code = code & "End Sub" & vbCrLf & vbCrLf

' ── SauvegarderArticle ──
code = code & "Public Sub SauvegarderArticle()" & vbCrLf
code = code & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""Par article"")" & vbCrLf
code = code & "    Dim code As String: code = Trim(ws.Cells(1, 2).Value)" & vbCrLf
code = code & "    If code = """" Then Exit Sub" & vbCrLf
code = code & "    Dim company As Object: Set company = ConnectSAP()" & vbCrLf
code = code & "    If company Is Nothing Then Exit Sub" & vbCrLf
code = code & "    On Error GoTo ErrHandler" & vbCrLf
code = code & "    Dim r As Long, cnt As Long" & vbCrLf
code = code & "    r = 6" & vbCrLf
code = code & "    Do While ws.Cells(r, 1).Value <> """"" & vbCrLf
code = code & "        Dim listNum As String: listNum = ws.Cells(r, 1).Value" & vbCrLf
code = code & "        RunUpdate company, ""UPDATE """"ITM1"""" SET """"Price"""" = "" & ws.Cells(r, 3).Value & "" WHERE """"ItemCode"""" = '"" & Replace(code, ""'"", ""''"") & ""' AND """"PriceList"""" = "" & listNum" & vbCrLf
code = code & "        cnt = cnt + 1" & vbCrLf
code = code & "        r = r + 1" & vbCrLf
code = code & "    Loop" & vbCrLf
code = code & "    MsgBox cnt & "" prix de liste mis a jour"", vbInformation" & vbCrLf
code = code & "    GoTo Cleanup" & vbCrLf
code = code & "ErrHandler:" & vbCrLf
code = code & "    MsgBox ""Erreur: "" & Err.Description, vbCritical" & vbCrLf
code = code & "Cleanup:" & vbCrLf
code = code & "    On Error Resume Next" & vbCrLf
code = code & "    company.Disconnect: Set company = Nothing" & vbCrLf
code = code & "End Sub" & vbCrLf & vbCrLf

' ── ChargerListes ──
code = code & "Public Sub ChargerListes()" & vbCrLf
code = code & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""Listes de prix"")" & vbCrLf
code = code & "    Application.StatusBar = ""Connexion SAP..."": Application.ScreenUpdating = False" & vbCrLf
code = code & "    Dim company As Object: Set company = ConnectSAP()" & vbCrLf
code = code & "    If company Is Nothing Then Application.StatusBar = False: Exit Sub" & vbCrLf
code = code & "    On Error GoTo ErrHandler" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    If ws.Cells(4, 1).Value <> """" Then ws.Range(ws.Cells(4, 1), ws.Cells(ws.Rows.Count, 9).End(3)).ClearContents" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    Dim rsL As Object" & vbCrLf
code = code & "    Set rsL = RunQuery(company, ""SELECT """"ListNum"""",""""ListName"""",""""BASE_NUM"""",""""Factor"""",""""IsGrossPrc"""",""""PrimCurr"""",""""ValidFor"""",""""ValidFrom"""",""""ValidTo"""" FROM """"OPLN"""" ORDER BY """"ListNum"""""")" & vbCrLf
code = code & "    Dim cnt As Long: cnt = GetRecordCount(rsL)" & vbCrLf
code = code & "    Dim r As Long" & vbCrLf
code = code & "    For r = 1 To cnt" & vbCrLf
code = code & "        ws.Cells(3 + r, 1).Value = GetFieldValue(rsL, ""ListNum"")" & vbCrLf
code = code & "        ws.Cells(3 + r, 2).Value = GetFieldValue(rsL, ""ListName"")" & vbCrLf
code = code & "        ws.Cells(3 + r, 3).Value = GetFieldValue(rsL, ""BASE_NUM"")" & vbCrLf
code = code & "        ws.Cells(3 + r, 4).Value = GetFieldValue(rsL, ""Factor"")" & vbCrLf
code = code & "        ws.Cells(3 + r, 5).Value = GetFieldValue(rsL, ""IsGrossPrc"")" & vbCrLf
code = code & "        ws.Cells(3 + r, 6).Value = GetFieldValue(rsL, ""PrimCurr"")" & vbCrLf
code = code & "        ws.Cells(3 + r, 7).Value = GetFieldValue(rsL, ""ValidFor"")" & vbCrLf
code = code & "        ws.Cells(3 + r, 8).Value = GetFieldValue(rsL, ""ValidFrom"")" & vbCrLf
code = code & "        ws.Cells(3 + r, 9).Value = GetFieldValue(rsL, ""ValidTo"")" & vbCrLf
code = code & "        MoveNextRS rsL" & vbCrLf
code = code & "    Next r" & vbCrLf
code = code & "    " & vbCrLf
code = code & "    Application.StatusBar = cnt & "" listes chargees""" & vbCrLf
code = code & "    GoTo Cleanup" & vbCrLf
code = code & "ErrHandler:" & vbCrLf
code = code & "    MsgBox ""Erreur: "" & Err.Description, vbCritical" & vbCrLf
code = code & "Cleanup:" & vbCrLf
code = code & "    On Error Resume Next" & vbCrLf
code = code & "    company.Disconnect: Set company = Nothing" & vbCrLf
code = code & "    Application.ScreenUpdating = True" & vbCrLf
code = code & "    Application.StatusBar = False" & vbCrLf
code = code & "End Sub" & vbCrLf

vbMod.CodeModule.AddFromString code

' ── Sauvegarder en xlsm ──
wb.SaveAs xlsmPath, 52  ' 52 = xlOpenXMLWorkbookMacroEnabled
wb.Close False
xlApp.Quit
Set xlApp = Nothing

MsgBox "GestionPrix.xlsm cree avec succes!" & vbCrLf & vbCrLf & xlsmPath, vbInformation, "Gestion Prix SAP"
