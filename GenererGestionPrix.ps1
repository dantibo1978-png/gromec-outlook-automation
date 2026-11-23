<#
  GenererGestionPrix.ps1 - Genere GestionPrix.xlsm
  Scenario 1: Appliquer facteurs par groupe de produits sur listes de prix clients
  MODE MATRICE: lignes=groupes, colonnes=listes de prix, cellules=facteurs
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

# ===== ONGLET FACTEURS =====
$wsFact.Cells(1,1).Value2 = "APPLIQUER FACTEURS PAR GROUPE DE PRODUITS"
$wsFact.Cells(1,1).Font.Bold = $true
$wsFact.Cells(1,1).Font.Size = 14

$wsFact.Cells(3,1).Value2 = "Liste de base:"
$wsFact.Cells(3,1).Font.Bold = $true
$wsFact.Cells(3,2).Interior.Color = 0xE0FFFF
$wsFact.Cells(3,3).Value2 = "(# de la liste de base, ex: 1)"
$wsFact.Cells(3,3).Font.Color = 0x808080

$wsFact.Cells(5,1).Value2 = "MATRICE GROUPES x LISTES DE PRIX"
$wsFact.Cells(5,1).Font.Bold = $true
$wsFact.Cells(5,1).Font.Size = 11
$wsFact.Cells(5,3).Value2 = "(Inscrivez le facteur aux intersections voulues, vide = ignorer)"
$wsFact.Cells(5,3).Font.Color = 0x808080

# Headers: A=Code Grp, B=Nom du groupe, C=Nb articles, D+ = listes de prix (dynamiques)
$hdrF = @("Code Grp","Nom du groupe","Nb articles")
for ($i = 0; $i -lt $hdrF.Count; $i++) {
    $c = $wsFact.Cells(6, $i+1)
    $c.Value2 = $hdrF[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x8B4513
    $c.Font.Color = 0xFFFFFF
}
$wsFact.Columns("A:A").ColumnWidth = 10
$wsFact.Columns("B:B").ColumnWidth = 35
$wsFact.Columns("C:C").ColumnWidth = 12

# ===== LISTES DE PRIX (reference, colonnes J-M sur Facteurs) =====
$hdrL = @("#","Nom","Base #","Facteur")
for ($i = 0; $i -lt $hdrL.Count; $i++) {
    $c = $wsFact.Cells(1, 10 + $i)
    $c.Value2 = $hdrL[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 0x003366
    $c.Font.Color = 0xFFFFFF
}
$wsFact.Columns("J:J").ColumnWidth = 6
$wsFact.Columns("K:K").ColumnWidth = 35
$wsFact.Columns("L:L").ColumnWidth = 8
$wsFact.Columns("M:M").ColumnWidth = 10

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
'  CHARGER GROUPES ET LISTES (matrice)
' ==========================================
Public Sub ChargerGroupesEtListes()
    Application.StatusBar = "Connexion SAP..."
    Application.ScreenUpdating = False
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Application.StatusBar = False: Application.ScreenUpdating = True: Exit Sub

    On Error GoTo ErrCGL

    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("Facteurs")
    Dim wsG As Worksheet: Set wsG = ThisWorkbook.Sheets("Groupes")

    ' --- Charger groupes dans onglet Groupes ---
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

    ' --- Charger groupes dans Facteurs (lignes 7+) ---
    lr = wsF.Cells(wsF.Rows.Count, 1).End(xlUp).Row
    If lr > 6 Then wsF.Range(wsF.Cells(7, 1), wsF.Cells(lr, 3)).ClearContents

    For r = 1 To cntG
        wsF.Cells(6 + r, 1).Value = wsG.Cells(1 + r, 1).Value
        wsF.Cells(6 + r, 2).Value = wsG.Cells(1 + r, 2).Value
        wsF.Cells(6 + r, 3).Value = wsG.Cells(1 + r, 3).Value
    Next r

    ' --- Charger listes de prix (colonnes D+ en matrice ET J-M en reference) ---
    Application.StatusBar = "Chargement listes de prix..."

    ' Reference J-M
    lr = wsF.Cells(wsF.Rows.Count, 10).End(xlUp).Row
    If lr > 1 Then wsF.Range(wsF.Cells(2, 10), wsF.Cells(lr, 13)).ClearContents

    Dim rsL As Object
    Set rsL = RunQuery(co, "SELECT ""ListNum"",""ListName"",""BASE_NUM"",""Factor"" FROM ""OPLN"" ORDER BY ""ListNum""")
    Dim cntL As Long: cntL = RC(rsL)

    ' Nettoyer anciennes colonnes matrice (D jusqu'a I max ou plus)
    Dim lastCol As Long
    lastCol = wsF.Cells(6, wsF.Columns.Count).End(-4159).Column
    If lastCol > 3 And lastCol < 10 Then
        wsF.Range(wsF.Cells(5, 4), wsF.Cells(wsF.Rows.Count, lastCol)).ClearContents
        wsF.Range(wsF.Cells(5, 4), wsF.Cells(wsF.Rows.Count, lastCol)).Interior.ColorIndex = 0
    End If

    Dim col As Long
    For r = 1 To cntL
        Dim listNum As Variant: listNum = GF(rsL, "ListNum")
        Dim listName As Variant: listName = GF(rsL, "ListName")

        ' Reference J-M
        wsF.Cells(1 + r, 10).Value = listNum
        wsF.Cells(1 + r, 11).Value = listName
        wsF.Cells(1 + r, 12).Value = GF(rsL, "BASE_NUM")
        wsF.Cells(1 + r, 13).Value = GF(rsL, "Factor")

        ' Matrice: colonne D + (r-1) = header avec # et nom de liste
        col = 3 + r
        wsF.Cells(5, col).Value = "Liste " & listNum
        wsF.Cells(5, col).Font.Bold = True
        wsF.Cells(5, col).Font.Size = 8
        wsF.Cells(6, col).Value = listName
        wsF.Cells(6, col).Font.Bold = True
        wsF.Cells(6, col).Font.Size = 8
        wsF.Cells(6, col).Interior.Color = 6697728
        wsF.Cells(6, col).Font.Color = 16777215
        wsF.Cells(5, col).Interior.Color = 6697728
        wsF.Cells(5, col).Font.Color = 16777215
        wsF.Columns(col).ColumnWidth = 12

        ' Colorer les cellules de saisie
        Dim rr As Long
        For rr = 1 To cntG
            wsF.Cells(6 + rr, col).Interior.Color = 15792383
        Next rr

        MN rsL
    Next r

    Application.StatusBar = cntG & " groupes, " & cntL & " listes chargees - Remplissez les facteurs dans la matrice"
    GoTo CleanCGL
ErrCGL:
    MsgBox "Erreur: " & Err.Description, vbCritical
CleanCGL:
    On Error Resume Next
    co.Disconnect: Set co = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' ==========================================
'  CHARGER GROUPES (seul)
' ==========================================
Public Sub ChargerGroupes()
    Call ChargerGroupesEtListes
End Sub

' ==========================================
'  CHARGER LISTES (seul)
' ==========================================
Public Sub ChargerListes()
    Call ChargerGroupesEtListes
End Sub

' ==========================================
'  APPLIQUER FACTEURS - Matrice
' ==========================================
Public Sub AppliquerFacteurs()
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("Facteurs")
    Dim listeBase As Variant: listeBase = wsF.Cells(3, 2).Value

    If listeBase = "" Then
        MsgBox "Entrez le # de liste de base en B3", vbExclamation
        Exit Sub
    End If

    ' Detecter colonnes de listes (D+, row 5 contient "Liste X")
    Dim col As Long, maxCol As Long
    maxCol = wsF.Cells(5, wsF.Columns.Count).End(-4159).Column
    If maxCol < 4 Then
        MsgBox "Aucune liste de prix chargee. Executez ChargerGroupesEtListes d'abord.", vbExclamation
        Exit Sub
    End If

    ' Collecter toutes les operations (groupe, liste cible, facteur)
    Dim ops() As String
    Dim nbOps As Long: nbOps = 0
    Dim r As Long, grpCode As Variant, facteur As Variant, listeCible As Variant

    r = 7
    Do While wsF.Cells(r, 1).Value <> ""
        grpCode = CStr(wsF.Cells(r, 1).Value)
        For col = 4 To maxCol
            facteur = wsF.Cells(r, col).Value
            If facteur <> "" And CDbl("0" & facteur) <> 0 Then
                ' Extraire le # de liste du header "Liste X"
                Dim hdr As String: hdr = CStr(wsF.Cells(5, col).Value)
                listeCible = Mid(hdr, 7)
                nbOps = nbOps + 1
                ReDim Preserve ops(1 To 3, 1 To nbOps)
                ops(1, nbOps) = grpCode
                ops(2, nbOps) = listeCible
                ops(3, nbOps) = CStr(facteur)
            End If
        Next col
        r = r + 1
    Loop

    If nbOps = 0 Then
        MsgBox "Aucun facteur saisi dans la matrice.", vbExclamation
        Exit Sub
    End If

    ' Connexion SAP
    Application.StatusBar = "Connexion SAP..."
    Dim co As Object: Set co = ConnectSAP()
    If co Is Nothing Then Exit Sub

    On Error GoTo ErrApp

    ' Resume et confirmation
    Dim msg As String, g As Long
    msg = "RESUME DE L'OPERATION" & vbCrLf & vbCrLf
    msg = msg & "Liste de base: " & listeBase & vbCrLf
    msg = msg & nbOps & " operations a effectuer:" & vbCrLf & vbCrLf
    For g = 1 To nbOps
        msg = msg & "Grp " & ops(1, g) & " -> Liste " & ops(2, g) & " x" & ops(3, g) & vbCrLf
        If g >= 15 And nbOps > 15 Then
            msg = msg & "... et " & (nbOps - 15) & " autres" & vbCrLf
            Exit For
        End If
    Next g
    msg = msg & vbCrLf & "CONFIRMER l'application?"

    If MsgBox(msg, vbYesNo + vbQuestion, "Confirmer") <> vbYes Then
        GoTo CleanApp
    End If

    ' Executer les UPDATE SQL
    Dim cntOK As Long: cntOK = 0

    For g = 1 To nbOps
        Application.StatusBar = "Operation " & g & "/" & nbOps & " - Grp " & ops(1, g) & " Liste " & ops(2, g) & "..."
        Dim sql As String
        sql = "UPDATE ""ITM1"" SET ""BasePLNum"" = " & listeBase & ", ""Factor"" = " & Replace(ops(3, g), ",", ".") & _
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

' ==========================================
'  APERCU FACTEURS - Matrice
' ==========================================
Public Sub ApercuFacteurs()
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("Facteurs")
    Dim listeBase As Variant: listeBase = wsF.Cells(3, 2).Value

    If listeBase = "" Then
        MsgBox "Entrez le # de liste de base en B3", vbExclamation
        Exit Sub
    End If

    MsgBox "L'apercu n'est pas disponible en mode matrice." & vbCrLf & _
           "Utilisez AppliquerFacteurs - un resume s'affiche avant confirmation.", vbInformation
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
Write-Host "    - ChargerGroupesEtListes : Charge groupes + listes en matrice"
Write-Host "    - AppliquerFacteurs      : Applique les facteurs de la matrice"
Write-Host "    - TestUpdateSQL          : Teste si HANA accepte les UPDATE"
Write-Host "====================================="
