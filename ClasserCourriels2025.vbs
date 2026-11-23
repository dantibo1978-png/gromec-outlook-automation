' ClasserCourriels2025.vbs
' Deplace tous les courriels de 2025 de la boite de reception vers un sous-dossier "2025"
'
' INSTRUCTIONS:
' 1. Ouvrir Outlook
' 2. Alt+F11 pour ouvrir l'editeur VBA
' 3. Coller ce code dans ThisOutlookSession (ou un nouveau Module)
' 4. F5 pour executer ClasserCourriels2025
' OU
' 1. Copier le fichier sur le bureau
' 2. Double-cliquer pour executer (via WScript)

Sub ClasserCourriels2025()
    Dim olApp As Object
    Dim olNs As Object
    Dim inbox As Object
    Dim dossier2025 As Object
    Dim items As Object
    Dim item As Object
    Dim i As Long
    Dim cnt As Long
    Dim total As Long

    Set olApp = Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set inbox = olNs.GetDefaultFolder(6) ' 6 = olFolderInbox

    ' Creer le sous-dossier 2025 s'il n'existe pas
    On Error Resume Next
    Set dossier2025 = inbox.Folders("2025")
    On Error GoTo 0

    If dossier2025 Is Nothing Then
        Set dossier2025 = inbox.Folders.Add("2025")
    End If

    ' Filtrer les courriels de 2025
    Set items = inbox.items
    items.Sort "[ReceivedTime]", False

    Dim sFilter As String
    sFilter = "[ReceivedTime] >= '2025-01-01' AND [ReceivedTime] < '2026-01-01'"

    Set items = items.Restrict(sFilter)
    total = items.Count

    If total = 0 Then
        MsgBox "Aucun courriel de 2025 dans la boite de reception.", vbInformation
        Exit Sub
    End If

    If MsgBox("Deplacer " & total & " courriels de 2025 vers le sous-dossier '2025'?" & vbCrLf & vbCrLf & _
              "Ca peut prendre quelques minutes.", vbYesNo + vbQuestion, "Classer courriels 2025") <> vbYes Then
        Exit Sub
    End If

    cnt = 0
    ' Parcourir a l'envers pour eviter les problemes d'index
    For i = total To 1 Step -1
        Set item = items(i)
        item.Move dossier2025
        cnt = cnt + 1
        If cnt Mod 100 = 0 Then
            DoEvents
        End If
    Next i

    MsgBox "Termine! " & cnt & " courriels deplaces vers Boite de reception\2025", vbInformation
End Sub
