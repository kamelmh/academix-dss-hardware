Attribute VB_Name = "mod_BuildReception"
Option Explicit

Public Sub BuildReceptionForm()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    
    On Error Resume Next
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    vbProj.VBComponents.Remove vbProj.VBComponents("frmReception")
    DoEvents
    On Error GoTo ErrHandler
    
    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)
    
    With frm
        .Properties("Name") = "frmReception"
        .Properties("Caption") = "Bon de Reception"
        .Properties("Width") = 620
        .Properties("Height") = 440
        .Properties("StartUpPosition") = 1
    End With
    
    frm.CodeModule.AddFromString GetReceptionFormCode()
    
    Dim ctrl As Object, t As Single, lc As Single
    lc = 15
    
    t = 10
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTitle": .Caption = "Bon de Reception - Quincaillerie": .Left = lc: .Top = t: .Width = 560: .Height = 28: .Font.Size = 14: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 48
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblRef": .Caption = "Ref:": .Left = lc: .Top = t: .Width = 40: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtRef": .Left = lc + 45: .Top = t: .Width = 120: .Height = 22: .Font.Size = 11: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblDate": .Caption = "Date:": .Left = lc + 180: .Top = t: .Width = 40: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtDate": .Left = lc + 225: .Top = t: .Width = 100: .Height = 22: .Text = Format(Date, "DD/MM/YYYY"): End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblSupplier": .Caption = "Fournisseur:": .Left = lc + 340: .Top = t: .Width = 80: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl: .Name = "cboSupplier": .Left = lc + 425: .Top = t: .Width = 140: .Height = 22: .Style = 2: End With
    
    t = 82
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblCode": .Caption = "Code:": .Left = lc: .Top = t: .Width = 40: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtCode": .Left = lc + 45: .Top = t: .Width = 100: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblDesign": .Caption = "Designation:": .Left = lc + 160: .Top = t: .Width = 80: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtDesignation": .Left = lc + 245: .Top = t: .Width = 180: .Height = 22: .Locked = True: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblQty": .Caption = "Qte:": .Left = lc + 440: .Top = t: .Width = 30: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtQty": .Left = lc + 475: .Top = t: .Width = 60: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnAddLine": .Caption = "Add": .Left = lc + 545: .Top = t: .Width = 45: .Height = 22: End With
    
    t = 115
    Set ctrl = frm.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl: .Name = "lstLines": .Left = lc: .Top = t: .Width = 580: .Height = 160: .ColumnCount = 5: .ColumnWidths = "60;200;60;80;100": End With
    
    t = 285
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTotalLabel": .Caption = "Total:": .Left = lc + 350: .Top = t: .Width = 50: .Height = 22: .Font.Bold = True: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTotal": .Caption = "0 DZD": .Left = lc + 410: .Top = t: .Width = 150: .Height = 22: .Font.Size = 12: .Font.Bold = True: .ForeColor = &HFF0000: End With
    
    t = 320
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnPrint": .Caption = "Imprimer": .Left = lc: .Top = t: .Width = 100: .Height = 35: .BackColor = &H7F4600: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClear": .Caption = "Clear": .Left = lc + 110: .Top = t: .Width = 80: .Height = 35: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClose": .Caption = "Close": .Left = lc + 500: .Top = t: .Width = 80: .Height = 35: End With
    
    MsgBox "frmReception created!" & vbCrLf & "Run: frmReception.Show", vbInformation, "Done"
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetReceptionFormCode() As String
    Dim c As String
    
    c = "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    txtRef.Text = ""REC-"" & Format(Now, ""yyyymmddhhmmss"")" & vbCrLf
    c = c & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""FOURNISSEURS"")" & vbCrLf
    c = c & "    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        cboSupplier.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    If cboSupplier.ListCount > 0 Then cboSupplier.ListIndex = 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub txtCode_Change()" & vbCrLf
    c = c & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""ARTICLES"")" & vbCrLf
    c = c & "    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If UCase(ws.Cells(i, 1).Value) = UCase(txtCode.Text) Then" & vbCrLf
    c = c & "            txtDesignation.Text = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            txtQty.SetFocus" & vbCrLf
    c = c & "            Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    txtDesignation.Text = """"" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnAddLine_Click()" & vbCrLf
    c = c & "    If Trim(txtCode.Text) = """" Then MsgBox ""Indiquez le code article"", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    If Val(txtQty.Text) <= 0 Then MsgBox ""Indiquez une quantite valide"", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""ARTICLES"")" & vbCrLf
    c = c & "    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim pu As Double: pu = 0" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If UCase(ws.Cells(i, 1).Value) = UCase(txtCode.Text) Then" & vbCrLf
    c = c & "            pu = Val(ws.Cells(i, 8).Value): Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    If pu = 0 Then MsgBox ""Article introuvable"", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    Dim qty As Double: qty = Val(txtQty.Text)" & vbCrLf
    c = c & "    Dim total As Double: total = qty * pu" & vbCrLf
    c = c & "    lstLines.AddItem txtCode.Text" & vbCrLf
    c = c & "    lstLines.List(lstLines.ListCount - 1, 1) = txtDesignation.Text" & vbCrLf
    c = c & "    lstLines.List(lstLines.ListCount - 1, 2) = qty" & vbCrLf
    c = c & "    lstLines.List(lstLines.ListCount - 1, 3) = pu" & vbCrLf
    c = c & "    lstLines.List(lstLines.ListCount - 1, 4) = total" & vbCrLf
    c = c & "    Dim grandTotal As Double" & vbCrLf
    c = c & "    Dim j As Long" & vbCrLf
    c = c & "    For j = 0 To lstLines.ListCount - 1" & vbCrLf
    c = c & "        grandTotal = grandTotal + Val(lstLines.List(j, 4))" & vbCrLf
    c = c & "    Next j" & vbCrLf
    c = c & "    lblTotal.Caption = Format(grandTotal, ""#,##0"") & "" DZD""" & vbCrLf
    c = c & "    txtCode.Text = """": txtDesignation.Text = """": txtQty.Text = """": txtCode.SetFocus" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnPrint_Click()" & vbCrLf
    c = c & "    On Error GoTo PrintErrHandler" & vbCrLf
    c = c & "    If lstLines.ListCount = 0 Then MsgBox ""Aucune ligne a imprimer"", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""BON_RECEPTION"")" & vbCrLf
    c = c & "    ws.Unprotect Password:=mod_Config.MASTER_PWD" & vbCrLf
    c = c & "    ws.Cells(6, 2).Value = txtRef.Text" & vbCrLf
    c = c & "    ws.Cells(6, 6).Value = txtDate.Text" & vbCrLf
    c = c & "    ws.Cells(8, 2).Value = cboSupplier.Text" & vbCrLf
    c = c & "    Dim i As Long, r As Long: r = 11" & vbCrLf
    c = c & "    Dim grandTotal As Double: grandTotal = 0" & vbCrLf
    c = c & "    For i = 0 To lstLines.ListCount - 1" & vbCrLf
    c = c & "        ws.Cells(r, 1).Value = i + 1" & vbCrLf
    c = c & "        ws.Cells(r, 2).Value = lstLines.List(i, 0)" & vbCrLf
    c = c & "        ws.Cells(r, 3).Value = lstLines.List(i, 1)" & vbCrLf
    c = c & "        ws.Cells(r, 4).Value = Val(lstLines.List(i, 2))" & vbCrLf
    c = c & "        ws.Cells(r, 5).Value = Val(lstLines.List(i, 3))" & vbCrLf
    c = c & "        ws.Cells(r, 6).Value = Val(lstLines.List(i, 4))" & vbCrLf
    c = c & "        grandTotal = grandTotal + Val(lstLines.List(i, 4))" & vbCrLf
    c = c & "        r = r + 1" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    ws.Cells(20, 6).Value = grandTotal" & vbCrLf
    c = c & "    ws.PrintOut Copies:=1" & vbCrLf
    c = c & "    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True" & vbCrLf
    c = c & "    MsgBox ""Impression terminee!"", vbInformation" & vbCrLf
    c = c & "    Exit Sub" & vbCrLf
    c = c & "PrintErrHandler:" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True" & vbCrLf
    c = c & "    MsgBox ""Erreur impression: "" & Err.Description, vbCritical, mod_Config.SYS_TITLE" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClear_Click()" & vbCrLf
    c = c & "    lstLines.Clear: lblTotal.Caption = ""0 DZD""" & vbCrLf
    c = c & "    txtCode.Text = """": txtDesignation.Text = """": txtQty.Text = """"" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    
    GetReceptionFormCode = c
End Function
