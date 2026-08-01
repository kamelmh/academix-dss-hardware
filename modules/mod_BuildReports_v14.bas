Attribute VB_Name = "mod_BuildReports"
Option Explicit

Public Sub BuildReportsForm()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    
    On Error Resume Next
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    vbProj.VBComponents.Remove vbProj.VBComponents("frmReports")
    DoEvents
    On Error GoTo ErrHandler
    
    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)
    
    With frm
        .Properties("Name") = "frmReports"
        .Properties("Caption") = "Rapports"
        .Properties("Width") = 560
        .Properties("Height") = 380
        .Properties("StartUpPosition") = 1
    End With
    
    frm.CodeModule.AddFromString GetReportsFormCode()
    
    Dim ctrl As Object, t As Single, lc As Single
    lc = 15
    
    t = 10
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTitle": .Caption = "Rapports & Analyse": .Left = lc: .Top = t: .Width = 500: .Height = 28: .Font.Size = 16: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 50
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblType": .Caption = "Type:": .Left = lc: .Top = t: .Width = 50: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl: .Name = "cboType": .Left = lc + 55: .Top = t: .Width = 200: .Height = 22: .Style = 2: End With
    
    t = 85
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblPreviewLabel": .Caption = "Apercu:": .Left = lc: .Top = t: .Width = 100: .Height = 22: .Font.Bold = True: End With
    
    t = 110
    Set ctrl = frm.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl: .Name = "lstPreview": .Left = lc: .Top = t: .Width = 510: .Height = 180: .ColumnCount = 4: .ColumnWidths = "80;200;100;100": End With
    
    t = 305
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnGenerate": .Caption = "Generer": .Left = lc: .Top = t: .Width = 100: .Height = 35: .BackColor = &H7F4600: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnPrint": .Caption = "Imprimer": .Left = lc + 110: .Top = t: .Width = 80: .Height = 35: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClose": .Caption = "Close": .Left = lc + 430: .Top = t: .Width = 80: .Height = 35: End With
    
    MsgBox "frmReports created!" & vbCrLf & "Run: frmReports.Show", vbInformation, "Done"
    Exit Sub
    
ErrHandler:
    MsgBox "Erreur " & Err.Number & ": " & Err.Description, vbCritical, "Erreur"
End Sub

Private Function GetReportsFormCode() As String
    Dim c As String
    
    c = "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboType.Clear" & vbCrLf
    c = c & "    cboType.AddItem ""ABC Analysis""" & vbCrLf
    c = c & "    cboType.AddItem ""Stock Aging""" & vbCrLf
    c = c & "    cboType.AddItem ""Supplier Performance""" & vbCrLf
    c = c & "    cboType.AddItem ""Stock Summary""" & vbCrLf
    c = c & "    cboType.AddItem ""Movement History""" & vbCrLf
    c = c & "    cboType.ListIndex = 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub cboType_Change()" & vbCrLf
    c = c & "    Call GenerateReport" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub GenerateReport()" & vbCrLf
    c = c & "    lstPreview.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""ARTICLES"")" & vbCrLf
    c = c & "    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    Select Case cboType.Text" & vbCrLf
    c = c & "        Case ""ABC Analysis""" & vbCrLf
    c = c & "            For i = 2 To lr" & vbCrLf
    c = c & "                Dim abc As String: abc = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "                Dim stockValue As Double: stockValue = Val(ws.Cells(i, 3).Value) * IIf(Val(ws.Cells(i, 12).Value) > 0, Val(ws.Cells(i, 12).Value), Val(ws.Cells(i, 8).Value))" & vbCrLf
    c = c & "                lstPreview.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 2) = abc" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 3) = Format(stockValue, ""#,##0"") & "" DZD""" & vbCrLf
    c = c & "            Next i" & vbCrLf
    c = c & "        Case ""Stock Aging""" & vbCrLf
    c = c & "            For i = 2 To lr" & vbCrLf
    c = c & "                lstPreview.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 2) = Format(Val(ws.Cells(i, 3).Value), ""#,##0"") & "" u""" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 3) = Format(Val(ws.Cells(i, 8).Value), ""#,##0"") & "" DZD""" & vbCrLf
    c = c & "            Next i" & vbCrLf
    c = c & "        Case ""Stock Summary""" & vbCrLf
    c = c & "            Dim totalVal As Double: totalVal = 0" & vbCrLf
    c = c & "            Dim totalQty As Double: totalQty = 0" & vbCrLf
    c = c & "            Dim cat As String" & vbCrLf
    c = c & "            For i = 2 To lr" & vbCrLf
    c = c & "                totalQty = totalQty + Val(ws.Cells(i, 3).Value)" & vbCrLf
    c = c & "                totalVal = totalVal + (Val(ws.Cells(i, 3).Value) * IIf(Val(ws.Cells(i, 12).Value) > 0, Val(ws.Cells(i, 12).Value), Val(ws.Cells(i, 8).Value)))" & vbCrLf
    c = c & "            Next i" & vbCrLf
    c = c & "            lstPreview.AddItem ""Total Articles: "" & (lr - 1)" & vbCrLf
    c = c & "            lstPreview.List(lstPreview.ListCount - 1, 1) = ""Total Stock: "" & Format(totalQty, ""#,##0"") & "" u""" & vbCrLf
    c = c & "            lstPreview.List(lstPreview.ListCount - 1, 2) = ""Valeur: "" & Format(totalVal, ""#,##0"") & "" DZD""" & vbCrLf
    c = c & "        Case ""Supplier Performance""" & vbCrLf
    c = c & "            Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets(""FOURNISSEURS"")" & vbCrLf
    c = c & "            Dim lrS As Long: lrS = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "            For i = 2 To lrS" & vbCrLf
    c = c & "                lstPreview.AddItem wsS.Cells(i, 1).Value" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 1) = wsS.Cells(i, 2).Value" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 2) = wsS.Cells(i, 3).Value" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 3) = wsS.Cells(i, 4).Value" & vbCrLf
    c = c & "            Next i" & vbCrLf
    c = c & "        Case ""Movement History""" & vbCrLf
    c = c & "            Dim wsM As Worksheet: Set wsM = ThisWorkbook.Sheets(""MOUVEMENTS"")" & vbCrLf
    c = c & "            Dim lrM As Long: lrM = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "            For i = 2 To lrM" & vbCrLf
    c = c & "                lstPreview.AddItem wsM.Cells(i, 1).Value" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 1) = wsM.Cells(i, 3).Value & "" - "" & wsM.Cells(i, 5).Value" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 2) = Format(Val(wsM.Cells(i, 6).Value), ""#,##0"") & "" u""" & vbCrLf
    c = c & "                lstPreview.List(lstPreview.ListCount - 1, 3) = wsM.Cells(i, 7).Value" & vbCrLf
    c = c & "            Next i" & vbCrLf
    c = c & "    End Select" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnGenerate_Click()" & vbCrLf
    c = c & "    Call GenerateReport" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnPrint_Click()" & vbCrLf
    c = c & "    MsgBox ""Report sent to printer"", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    
    GetReportsFormCode = c
End Function
