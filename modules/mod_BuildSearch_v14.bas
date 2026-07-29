Attribute VB_Name = "mod_BuildSearch"
Option Explicit

Public Sub BuildSearchForm()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    
    ' Remove old form if it exists
    On Error Resume Next
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    vbProj.VBComponents.Remove vbProj.VBComponents("frmSearch")
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    On Error GoTo ErrHandler
    
    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)
    
    With frm
        .Properties("Name") = "frmSearch"
        .Properties("Caption") = "Recherche"
        .Properties("Width") = 600
        .Properties("Height") = 420
        .Properties("StartUpPosition") = 1
    End With
    
    frm.CodeModule.AddFromString GetSearchFormCode()
    
    Dim ctrl As Object, t As Single, lc As Single
    lc = 15
    
    t = 10
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTitle": .Caption = "Recherche Universelle": .Left = lc: .Top = t: .Width = 550: .Height = 28: .Font.Size = 16: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 45
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblSearch": .Caption = "Rechercher:": .Left = lc: .Top = t: .Width = 80: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtSearch": .Left = lc + 85: .Top = t: .Width = 300: .Height = 22: .Font.Size = 11: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl: .Name = "cboScope": .Left = lc + 400: .Top = t: .Width = 140: .Height = 22: .Style = 2: End With
    
    t = 80
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblResultsLabel": .Caption = "Resultats (0):": .Left = lc: .Top = t: .Width = 200: .Height = 22: .Font.Size = 10: .Font.Bold = True: End With
    
    t = 105
    Set ctrl = frm.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl: .Name = "lstResults": .Left = lc: .Top = t: .Width = 550: .Height = 240: .ColumnCount = 4: .ColumnWidths = "80;200;120;120": End With
    
    t = 360
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnView": .Caption = "View": .Left = lc: .Top = t: .Width = 80: .Height = 30: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClear": .Caption = "Clear": .Left = lc + 90: .Top = t: .Width = 80: .Height = 30: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClose": .Caption = "Close": .Left = lc + 470: .Top = t: .Width = 80: .Height = 30: End With
    
    MsgBox "frmSearch created!" & vbCrLf & "Run: frmSearch.Show", vbInformation, "Done"
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetSearchFormCode() As String
    Dim c As String
    
    c = "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboScope.Clear" & vbCrLf
    c = c & "    cboScope.AddItem ""All""" & vbCrLf
    c = c & "    cboScope.AddItem ""Articles""" & vbCrLf
    c = c & "    cboScope.AddItem ""Fournisseurs""" & vbCrLf
    c = c & "    cboScope.AddItem ""Mouvements""" & vbCrLf
    c = c & "    cboScope.ListIndex = 0" & vbCrLf
    c = c & "    txtSearch.SetFocus" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub txtSearch_Change()" & vbCrLf
    c = c & "    Call DoSearch" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub cboScope_Change()" & vbCrLf
    c = c & "    Call DoSearch" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub DoSearch()" & vbCrLf
    c = c & "    Dim q As String: q = LCase(Trim(txtSearch.Text))" & vbCrLf
    c = c & "    Dim scope As String: scope = cboScope.Text" & vbCrLf
    c = c & "    lstResults.Clear" & vbCrLf
    c = c & "    If Len(q) = 0 Then lblResultsLabel.Caption = ""Resultats (0)"": Exit Sub" & vbCrLf
    c = c & "    Dim count As Long: count = 0" & vbCrLf
    c = c & "    If scope = ""All"" Or scope = ""Articles"" Then Call SearchArticles(q, count)" & vbCrLf
    c = c & "    If scope = ""All"" Or scope = ""Fournisseurs"" Then Call SearchSuppliers(q, count)" & vbCrLf
    c = c & "    If scope = ""All"" Or scope = ""Mouvements"" Then Call SearchMovements(q, count)" & vbCrLf
    c = c & "    lblResultsLabel.Caption = ""Resultats ("" & count & "")""" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub SearchArticles(q As String, count As Long)" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next: Set ws = ThisWorkbook.Sheets(""ARTICLES""): On Error GoTo 0" & vbCrLf
    c = c & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim code As String: code = LCase(ws.Cells(i, 1).Value)" & vbCrLf
    c = c & "        Dim desc As String: desc = LCase(ws.Cells(i, 2).Value)" & vbCrLf
    c = c & "        Dim cat As String: cat = LCase(ws.Cells(i, 5).Value)" & vbCrLf
    c = c & "        If InStr(1, code, q) > 0 Or InStr(1, desc, q) > 0 Or InStr(1, cat, q) > 0 Then" & vbCrLf
    c = c & "            lstResults.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lstResults.List(lstResults.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            lstResults.List(lstResults.ListCount - 1, 2) = Format(Val(ws.Cells(i, 3).Value), ""#,##0"") & "" """ & vbCrLf
    c = c & "            lstResults.List(lstResults.ListCount - 1, 3) = Format(Val(ws.Cells(i, 8).Value), ""#,##0"") & "" DZD""" & vbCrLf
    c = c & "            count = count + 1" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub SearchSuppliers(q As String, count As Long)" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next: Set ws = ThisWorkbook.Sheets(""FOURNISSEURS""): On Error GoTo 0" & vbCrLf
    c = c & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim code As String: code = LCase(ws.Cells(i, 1).Value)" & vbCrLf
    c = c & "        Dim name As String: name = LCase(ws.Cells(i, 2).Value)" & vbCrLf
    c = c & "        Dim tel As String: tel = LCase(ws.Cells(i, 4).Value)" & vbCrLf
    c = c & "        If InStr(1, code, q) > 0 Or InStr(1, name, q) > 0 Or InStr(1, tel, q) > 0 Then" & vbCrLf
    c = c & "            lstResults.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lstResults.List(lstResults.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            lstResults.List(lstResults.ListCount - 1, 2) = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            lstResults.List(lstResults.ListCount - 1, 3) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            count = count + 1" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub SearchMovements(q As String, count As Long)" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next: Set ws = ThisWorkbook.Sheets(""MOUVEMENTS""): On Error GoTo 0" & vbCrLf
    c = c & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim code As String: code = LCase(ws.Cells(i, 3).Value)" & vbCrLf
    c = c & "        Dim movType As String: movType = LCase(ws.Cells(i, 5).Value)" & vbCrLf
    c = c & "        If InStr(1, code, q) > 0 Or InStr(1, movType, q) > 0 Then" & vbCrLf
    c = c & "            lstResults.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lstResults.List(lstResults.ListCount - 1, 1) = ws.Cells(i, 3).Value & "" - "" & ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            lstResults.List(lstResults.ListCount - 1, 2) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            lstResults.List(lstResults.ListCount - 1, 3) = Format(Val(ws.Cells(i, 6).Value), ""#,##0"") & "" """ & vbCrLf
    c = c & "            count = count + 1" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnView_Click()" & vbCrLf
    c = c & "    If lstResults.ListIndex < 0 Then MsgBox ""Select an item"", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    Dim selVal As String: selVal = lstResults.Value" & vbCrLf
    c = c & "    Dim selCol1 As String: selCol1 = lstResults.List(lstResults.ListIndex, 0)" & vbCrLf
    c = c & "    Dim scope As String: scope = cboScope.Text" & vbCrLf
    c = c & "    Dim ws As Worksheet, lr As Long, i As Long" & vbCrLf
    c = c & "    ' Try ARTICLES" & vbCrLf
    c = c & "    If scope = ""All"" Or scope = ""Articles"" Then" & vbCrLf
    c = c & "        On Error Resume Next: Set ws = ThisWorkbook.Sheets(""ARTICLES""): On Error GoTo 0" & vbCrLf
    c = c & "        If Not ws Is Nothing Then" & vbCrLf
    c = c & "            lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "            For i = 2 To lr" & vbCrLf
    c = c & "                If ws.Cells(i, 1).Value = selVal Then ws.Activate: ws.Cells(i, 1).Select: Exit Sub" & vbCrLf
    c = c & "            Next i" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    ' Try FOURNISSEURS" & vbCrLf
    c = c & "    If scope = ""All"" Or scope = ""Fournisseurs"" Then" & vbCrLf
    c = c & "        Set ws = Nothing" & vbCrLf
    c = c & "        On Error Resume Next: Set ws = ThisWorkbook.Sheets(""FOURNISSEURS""): On Error GoTo 0" & vbCrLf
    c = c & "        If Not ws Is Nothing Then" & vbCrLf
    c = c & "            lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "            For i = 2 To lr" & vbCrLf
    c = c & "                If ws.Cells(i, 1).Value = selVal Then ws.Activate: ws.Cells(i, 1).Select: Exit Sub" & vbCrLf
    c = c & "            Next i" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    ' Try MOUVEMENTS" & vbCrLf
    c = c & "    If scope = ""All"" Or scope = ""Mouvements"" Then" & vbCrLf
    c = c & "        Set ws = Nothing" & vbCrLf
    c = c & "        On Error Resume Next: Set ws = ThisWorkbook.Sheets(""MOUVEMENTS""): On Error GoTo 0" & vbCrLf
    c = c & "        If Not ws Is Nothing Then" & vbCrLf
    c = c & "            lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "            For i = 2 To lr" & vbCrLf
    c = c & "                If CStr(ws.Cells(i, 1).Value) = selVal Then ws.Activate: ws.Cells(i, 1).Select: Exit Sub" & vbCrLf
    c = c & "            Next i" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClear_Click()" & vbCrLf
    c = c & "    txtSearch.Text = """": lstResults.Clear: lblResultsLabel.Caption = ""Resultats (0)""" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    
    GetSearchFormCode = c
End Function
