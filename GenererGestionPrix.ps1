<#
  GenererGestionPrix.ps1 - Genere GestionPrix.xlsm
  Scenario 1: Appliquer facteurs par groupe de produits sur listes de prix clients
  MODE TABLEAU: saisie libre Code Grp / Liste cible / Liste base / Facteur
  Le nom du groupe se remplit automatiquement (VLOOKUP sur onglet Groupes)
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
$wsGroupes = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsFact)
$wsGroupes.Name = "Groupes"
$wsListes = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsGroupes)
$wsListes.Name = "Listes"

# ===== ONGLET FACTEURS =====
$wsFact.Cells(1,1).Value2 = "GESTION DES PRIX - FACTEURS PAR GROUPE"
$wsFact.Cells(1,1).Font.Bold = $true
$wsFact.Cells(1,1).Font.Size = 14

$wsFact.Cells(3,1).Value2 = "Inscrivez vos operations ci-dessous. Le nom du groupe se remplit automatiquement."
$wsFact.Cells(3,1).Font.Color = 0x808080
$wsFact.Cells(3,1).Font.Italic = $true

# Headers row 5
$hdrF = @("Code Grp","Nom du groupe","Liste cible","Liste base","Facteur")
for ($i = 0; $i -lt $hdrF.Count; $i++) {
    $c = $wsFact.Cells(5, $i+1)
    $c.Value2 = $hdrF[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x8B4513
    $c.Font.Color = 0xFFFFFF
}
$wsFact.Columns("A:A").ColumnWidth = 12
$wsFact.Columns("B:B").ColumnWidth = 38
$wsFact.Columns("C:C").ColumnWidth = 14
$wsFact.Columns("D:D").ColumnWidth = 14
$wsFact.Columns("E:E").ColumnWidth = 12

# Pre-fill 50 input rows with alternating colors and VLOOKUP formula
for ($r = 6; $r -le 55; $r++) {
    # Col A: Code Grp (saisie) - jaune clair
    $wsFact.Cells($r, 1).Interior.Color = 0xE0FFFF
    # Col B: Nom du groupe (VLOOKUP auto)
    $wsFact.Cells($r, 2).Formula = "=IF(A$r="""","""",IFERROR(VLOOKUP(A$r,Groupes!A:B,2,FALSE),""Groupe inconnu""))"
    $wsFact.Cells($r, 2).Font.Color = 0x404040
    # Col C: Liste cible (saisie) - jaune clair
    $wsFact.Cells($r, 3).Interior.Color = 0xE0FFFF
    # Col D: Liste base (saisie) - jaune clair
    $wsFact.Cells($r, 4).Interior.Color = 0xE0FFFF
    # Col E: Facteur (saisie) - jaune clair
    $wsFact.Cells($r, 5).Interior.Color = 0xE0FFFF

    # Lignes alternees
    if (($r - 6) % 2 -eq 1) {
        for ($c = 1; $c -le 5; $c++) {
            $wsFact.Cells($r, $c).Interior.Color = 0xFFF8F0
        }
    }
}

# Reference listes de prix (colonnes G-J)
$hdrRef = @("#","Nom de la liste","Base #","Facteur SAP")
for ($i = 0; $i -lt $hdrRef.Count; $i++) {
    $c = $wsFact.Cells(5, 7 + $i)
    $c.Value2 = $hdrRef[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x003366
    $c.Font.Color = 0xFFFFFF
}
$wsFact.Columns("G:G").ColumnWidth = 6
$wsFact.Columns("H:H").ColumnWidth = 30
$wsFact.Columns("I:I").ColumnWidth = 8
$wsFact.Columns("J:J").ColumnWidth = 12

# Separator column F
$wsFact.Columns("F:F").ColumnWidth = 3

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

# ===== ONGLET LISTES =====
$hdrL = @("#","Nom de la liste","Base #","Facteur")
for ($i = 0; $i -lt $hdrL.Count; $i++) {
    $c = $wsListes.Cells(1, $i+1)
    $c.Value2 = $hdrL[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x003366
    $c.Font.Color = 0xFFFFFF
}
$wsListes.Columns("A:A").ColumnWidth = 6
$wsListes.Columns("B:B").ColumnWidth = 35
$wsListes.Columns("C:C").ColumnWidth = 8
$wsListes.Columns("D:D").ColumnWidth = 10

$wsFact.Activate()
$wsFact.Cells(6,1).Select()

# ========== INJECTION VBA ==========
Write-Host "Injection du code VBA..."

$vbaCode = @'
Option Explicit

Private mSapUser As String
Private mSapPass As String

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
        "Sub Window_OnLoad() : window.resizeTo 350, 220 : document.getElementById(""usr"").focus : End Sub" & vbCrLf & _
        "</script></head><body>" & _
        "<table><tr><td><b>Usager</b></td><td><input id=""usr"" value=""" & mSapUser & """></td></tr>" & _
        "<tr><td><b>Mot de passe</b></td><td><input id=""pwd"" type=""password""></td></tr></table><br>" & _
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

' ==========================================
'  CHARGER DONNEES SAP (groupes + listes)
' ==========================================
Public Sub ChargerGroupesEtListes()
    Application.StatusBar = "Connexion SAP..."
    Application.ScreenUpdating = False
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Application.ScreenUpdating = True: Exit Sub

    On Error GoTo ErrCGL

    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("Facteurs")
    Dim wsG As Worksheet: Set wsG = ThisWorkbook.Sheets("Groupes")
    Dim wsL As Worksheet: Set wsL = ThisWorkbook.Sheets("Listes")

    ' --- Charger groupes ---
    Dim lr As Long
    lr = wsG.Cells(wsG.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then wsG.Range(wsG.Cells(2, 1), wsG.Cells(lr, 3)).ClearContents

    Application.StatusBar = "Chargement groupes de produits..."
    Dim rsG As Object
    Set rsG = RunQuery(co, "SELECT g.""ItmsGrpCod"", g.""ItmsGrpNam"", COUNT(i.""ItemCode"") AS ""NbItems"" FROM ""OITB"" g LEFT JOIN ""OITM"" i ON g.""ItmsGrpCod"" = i.""ItmsGrpCod"" GROUP BY g.""ItmsGrpCod"", g.""ItmsGrpNam"" ORDER BY g.""ItmsGrpCod""")
    Dim cntG As Long: cntG = RC(rsG)
    Dim r As Long
    For r = 1 To cntG
        wsG.Cells(1 + r, 1).Value = GF(rsG, "ItmsGrpCod")
        wsG.Cells(1 + r, 2).Value = GF(rsG, "ItmsGrpNam")
        wsG.Cells(1 + r, 3).Value = GF(rsG, "NbItems")
        MN rsG
    Next r

    ' --- Charger listes de prix ---
    Application.StatusBar = "Chargement listes de prix..."
    lr = wsL.Cells(wsL.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then wsL.Range(wsL.Cells(2, 1), wsL.Cells(lr, 4)).ClearContents

    ' Reference sur Facteurs (G-J)
    lr = wsF.Cells(wsF.Rows.Count, 7).End(xlUp).Row
    If lr > 5 Then wsF.Range(wsF.Cells(6, 7), wsF.Cells(lr, 10)).ClearContents

    Dim rsL As Object
    Set rsL = RunQuery(co, "SELECT ""ListNum"",""ListName"",""BASE_NUM"",""Factor"" FROM ""OPLN"" ORDER BY ""ListNum""")
    Dim cntL As Long: cntL = RC(rsL)
    For r = 1 To cntL
        wsL.Cells(1 + r, 1).Value = GF(rsL, "ListNum")
        wsL.Cells(1 + r, 2).Value = GF(rsL, "ListName")
        wsL.Cells(1 + r, 3).Value = GF(rsL, "BASE_NUM")
        wsL.Cells(1 + r, 4).Value = GF(rsL, "Factor")
        wsF.Cells(5 + r, 7).Value = GF(rsL, "ListNum")
        wsF.Cells(5 + r, 8).Value = GF(rsL, "ListName")
        wsF.Cells(5 + r, 9).Value = GF(rsL, "BASE_NUM")
        wsF.Cells(5 + r, 10).Value = GF(rsL, "Factor")
        MN rsL
    Next r

    Application.StatusBar = cntG & " groupes, " & cntL & " listes chargees"
    GoTo CleanCGL
ErrCGL:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanCGL:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub ChargerGroupes()
    Call ChargerGroupesEtListes
End Sub

Public Sub ChargerListes()
    Call ChargerGroupesEtListes
End Sub

' ==========================================
'  APPLIQUER FACTEURS - Tableau simple
' ==========================================
Public Sub AppliquerFacteurs()
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("Facteurs")

    Dim ops() As String
    Dim nbOps As Long: nbOps = 0
    Dim r As Long

    r = 6
    Do While r <= 55
        Dim grpCode As Variant: grpCode = wsF.Cells(r, 1).Value
        Dim listeCible As Variant: listeCible = wsF.Cells(r, 3).Value
        Dim listeBase As Variant: listeBase = wsF.Cells(r, 4).Value
        Dim facteur As Variant: facteur = wsF.Cells(r, 5).Value

        If grpCode <> "" And listeCible <> "" And facteur <> "" Then
            If listeBase = "" Then listeBase = 1
            nbOps = nbOps + 1
            ReDim Preserve ops(1 To 4, 1 To nbOps)
            ops(1, nbOps) = CStr(grpCode)
            ops(2, nbOps) = CStr(listeCible)
            ops(3, nbOps) = CStr(facteur)
            ops(4, nbOps) = CStr(listeBase)
        End If
        r = r + 1
    Loop

    If nbOps = 0 Then
        MsgBox "Aucune operation saisie." & vbCrLf & vbCrLf & _
               "Remplissez les colonnes:" & vbCrLf & _
               "A = Code Grp, C = Liste cible, D = Liste base, E = Facteur", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "Connexion SAP..."
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Exit Sub

    On Error GoTo ErrApp

    Dim msg As String, g As Long
    msg = "RESUME - " & nbOps & " operations:" & vbCrLf & vbCrLf
    For g = 1 To nbOps
        Dim grpNom As String: grpNom = CStr(wsF.Cells(5 + g, 2).Value)
        If Len(grpNom) > 20 Then grpNom = Left(grpNom, 20) & "..."
        msg = msg & "Grp " & ops(1, g) & " (" & grpNom & ") -> Liste " & ops(2, g) & " (base " & ops(4, g) & ") x" & ops(3, g) & vbCrLf
        If g >= 20 And nbOps > 20 Then
            msg = msg & "... et " & (nbOps - 20) & " autres" & vbCrLf
            Exit For
        End If
    Next g
    msg = msg & vbCrLf & "CONFIRMER l'application?"

    If MsgBox(msg, vbYesNo + vbQuestion, "Confirmer") <> vbYes Then
        GoTo CleanApp
    End If

    Dim cntOK As Long: cntOK = 0

    For g = 1 To nbOps
        Application.StatusBar = "Operation " & g & "/" & nbOps & " - Grp " & ops(1, g) & " Liste " & ops(2, g) & "..."
        Dim sql As String
        sql = "UPDATE ""ITM1"" SET ""BasePLNum"" = " & ops(4, g) & ", ""Factor"" = " & Replace(ops(3, g), ",", ".") & _
              " WHERE ""PriceList"" = " & ops(2, g) & _
              " AND ""ItemCode"" IN (SELECT ""ItemCode"" FROM ""OITM"" WHERE ""ItmsGrpCod"" = " & ops(1, g) & _
              " AND LEFT(""ItemCode"",2) NOT IN ('PT','PR','TA','TB','TC','TD'))"
        Dim rsUpd As Object
        Set rsUpd = co.GetBusinessObject(300)
        CallByName rsUpd, "DoQuery", VbMethod, sql
        Set rsUpd = Nothing
        cntOK = cntOK + 1
    Next g

    MsgBox "Termine! " & cntOK & " operations executees.", vbInformation
    Application.StatusBar = "Termine - " & cntOK & " operations executees"
    GoTo CleanApp

ErrApp:
    MsgBox "Erreur: " & Err.Description, vbCritical
    Application.StatusBar = False
CleanApp:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.ScreenUpdating = True
End Sub

Public Sub TestUpdateSQL()
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Exit Sub

    Dim rs As Object: Set rs = co.GetBusinessObject(300)
    On Error Resume Next
    CallByName rs, "DoQuery", VbMethod, "UPDATE ""ITM1"" SET ""Factor"" = 1 WHERE ""ItemCode"" = 'ZZZZTEST' AND 1=0"
    If Err.Number <> 0 Then
        MsgBox "UPDATE refuse: " & Err.Description, vbExclamation
    Else
        MsgBox "UPDATE SQL permis! On peut optimiser.", vbInformation
    End If
    On Error GoTo 0
    co.Disconnect: Set co = Nothing
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
$wb.SaveAs($xlsmPath, 52)
Write-Host "Fichier sauvegarde: $xlsmPath"
$wb.Close($false)
$xl.Quit()

Write-Host ""
Write-Host "====================================="
Write-Host "  GestionPrix.xlsm cree avec succes!"
Write-Host "  Ouvrez: $xlsmPath"
Write-Host "  Macros (Alt+F8):"
Write-Host "    - ChargerGroupesEtListes : Charge groupes + listes de SAP"
Write-Host "    - AppliquerFacteurs      : Applique les facteurs saisis"
Write-Host "====================================="
