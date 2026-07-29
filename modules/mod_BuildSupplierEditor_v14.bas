Attribute VB_Name = "mod_BuildSupplierEditor"
Option Explicit

Public Sub BuildSupplierEditorForm()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    
    On Error Resume Next
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    vbProj.VBComponents.Remove vbProj.VBComponents("frmSupplierEditor")
    DoEvents
    On Error GoTo ErrHandler
    
    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)
    
    With frm
        .Properties("Name") = "frmSupplierEditor"
        .Properties("Caption") = "Supplier Editor"
        .Properties("Width") = 560
        .Properties("Height") = 420
        .Properties("StartUpPosition") = 1
    End With
    
    frm.CodeModule.AddFromString GetSupplierFormCode()
    
    Dim ctrl As Object, t As Single, lc As Single, rc As Single
    lc = 10: rc = 280
    
    t = 8
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTitle": .Caption = "Gestion des Fournisseurs": .Left = lc: .Top = t: .Width = 520: .Height = 28: .Font.Size = 14: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 42
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblCode": .Caption = "Code:": .Left = lc: .Top = t: .Width = 45: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtCode": .Left = 60: .Top = t: .Width = 100: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnSearch": .Caption = "...": .Left = 165: .Top = t: .Width = 30: .Height = 22: End With
    
    t = 70
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblRaison": .Caption = "Raison Sociale:": .Left = lc: .Top = t: .Width = 80: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtRaisonSociale": .Left = 95: .Top = t: .Width = 420: .Height = 22: End With
    
    t = 98
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblAdresse": .Caption = "Adresse:": .Left = lc: .Top = t: .Width = 55: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtAdresse": .Left = 70: .Top = t: .Width = 445: .Height = 22: End With
    
    t = 126
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTel": .Caption = "Telephone:": .Left = lc: .Top = t: .Width = 60: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtTelephone": .Left = 75: .Top = t: .Width = 150: .Height = 22: End With
    
    t = 156
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblNIF": .Caption = "NIF:": .Left = lc: .Top = t: .Width = 30: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtNIF": .Left = 45: .Top = t: .Width = 150: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblNIS": .Caption = "NIS:": .Left = rc: .Top = t: .Width = 30: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtNIS": .Left = rc + 35: .Top = t: .Width = 150: .Height = 22: End With
    
    t = 184
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblRC": .Caption = "RC:": .Left = lc: .Top = t: .Width = 30: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtRC": .Left = 45: .Top = t: .Width = 150: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblArtImpot": .Caption = "Art. Impot:": .Left = rc: .Top = t: .Width = 60: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtArtImpot": .Left = rc + 65: .Top = t: .Width = 150: .Height = 22: End With
    
    t = 220
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnSave": .Caption = "Save": .Left = lc: .Top = t: .Width = 100: .Height = 35: .BackColor = &HD48700: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnNew": .Caption = "New": .Left = 120: .Top = t: .Width = 100: .Height = 35: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnDelete": .Caption = "Delete": .Left = 230: .Top = t: .Width = 100: .Height = 35: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClose": .Caption = "Close": .Left = 420: .Top = t: .Width = 100: .Height = 35: End With
    
    MsgBox "frmSupplierEditor created!" & vbCrLf & "Run: frmSupplierEditor.Show", vbInformation, "Done"
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetSupplierFormCode() As String
    Dim c As String
    
    c = "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Me.Caption = ""Gestion des Fournisseurs""" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnSearch_Click()" & vbCrLf
    c = c & "    Dim s As String" & vbCrLf
    c = c & "    s = InputBox(""Enter supplier code (SUP-xxx):"", ""Search"")" & vbCrLf
    c = c & "    If Len(Trim(s)) > 0 Then" & vbCrLf
    c = c & "        txtCode.Value = UCase(Trim(s))" & vbCrLf
    c = c & "        Call LoadSupplier" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub LoadSupplier()" & vbCrLf
    c = c & "    Dim code As String" & vbCrLf
    c = c & "    code = UCase(Trim(txtCode.Value))" & vbCrLf
    c = c & "    If Len(code) = 0 Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(""FOURNISSEURS"")" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim r As Variant" & vbCrLf
    c = c & "    r = Application.Match(code, ws.Range(""A:A""), 0)" & vbCrLf
    c = c & "    If IsError(r) Then MsgBox ""Supplier not found."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    txtRaisonSociale.Value = ws.Cells(r, 2).Value" & vbCrLf
    c = c & "    txtAdresse.Value = ws.Cells(r, 3).Value" & vbCrLf
    c = c & "    txtTelephone.Value = ws.Cells(r, 4).Value" & vbCrLf
    c = c & "    txtNIF.Value = ws.Cells(r, 5).Value" & vbCrLf
    c = c & "    txtNIS.Value = ws.Cells(r, 6).Value" & vbCrLf
    c = c & "    txtRC.Value = ws.Cells(r, 7).Value" & vbCrLf
    c = c & "    txtArtImpot.Value = ws.Cells(r, 8).Value" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnSave_Click()" & vbCrLf
    c = c & "    If Len(Trim(txtCode.Value)) = 0 Then MsgBox ""Enter code."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    If Len(Trim(txtRaisonSociale.Value)) = 0 Then MsgBox ""Enter raison sociale."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    If MsgBox(""Save supplier?"", vbQuestion + vbYesNo) = vbNo Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(""FOURNISSEURS"")" & vbCrLf
    c = c & "    ws.Unprotect Password:=mod_Config.MASTER_PWD" & vbCrLf
    c = c & "    Dim code As String: code = UCase(Trim(txtCode.Value))" & vbCrLf
    c = c & "    Dim r As Variant" & vbCrLf
    c = c & "    r = Application.Match(code, ws.Range(""A:A""), 0)" & vbCrLf
    c = c & "    Dim nr As Long" & vbCrLf
    c = c & "    If IsError(r) Then" & vbCrLf
    c = c & "        nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        nr = r" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    ws.Cells(nr, 1).Value = code" & vbCrLf
    c = c & "    ws.Cells(nr, 2).Value = Trim(txtRaisonSociale.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 3).Value = Trim(txtAdresse.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 4).Value = Trim(txtTelephone.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 5).Value = Trim(txtNIF.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 6).Value = Trim(txtNIS.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 7).Value = Trim(txtRC.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 8).Value = Trim(txtArtImpot.Value)" & vbCrLf
    c = c & "    ws.Protect Password:=mod_Config.MASTER_PWD" & vbCrLf
    c = c & "    MsgBox ""Supplier saved!"", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnNew_Click()" & vbCrLf
    c = c & "    txtCode.Value = """"" & vbCrLf
    c = c & "    txtRaisonSociale.Value = """"" & vbCrLf
    c = c & "    txtAdresse.Value = """"" & vbCrLf
    c = c & "    txtTelephone.Value = """"" & vbCrLf
    c = c & "    txtNIF.Value = """"" & vbCrLf
    c = c & "    txtNIS.Value = """"" & vbCrLf
    c = c & "    txtRC.Value = """"" & vbCrLf
    c = c & "    txtArtImpot.Value = """"" & vbCrLf
    c = c & "    txtCode.SetFocus" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnDelete_Click()" & vbCrLf
    c = c & "    Dim code As String: code = UCase(Trim(txtCode.Value))" & vbCrLf
    c = c & "    If Len(code) = 0 Then MsgBox ""Enter code first."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    If MsgBox(""Delete supplier "" & code & ""?"", vbQuestion + vbYesNo) = vbNo Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(""FOURNISSEURS"")" & vbCrLf
    c = c & "    Dim r As Variant" & vbCrLf
    c = c & "    r = Application.Match(code, ws.Range(""A:A""), 0)" & vbCrLf
    c = c & "    If IsError(r) Then MsgBox ""Not found."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    ws.Unprotect Password:=mod_Config.MASTER_PWD" & vbCrLf
    c = c & "    ws.Rows(r).Delete" & vbCrLf
    c = c & "    ws.Protect Password:=mod_Config.MASTER_PWD" & vbCrLf
    c = c & "    Call btnNew_Click" & vbCrLf
    c = c & "    MsgBox ""Supplier deleted."", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    
    GetSupplierFormCode = c
End Function
