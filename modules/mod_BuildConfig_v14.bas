Attribute VB_Name = "mod_BuildConfig"
Option Explicit

Public Sub BuildConfigForm()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    
    On Error Resume Next
    DoEvents
    Dim existingConfig As Object
    Set existingConfig = vbProj.VBComponents("frmConfig")
    If Not existingConfig Is Nothing Then
        vbProj.VBComponents.Remove existingConfig
    End If
    DoEvents
    On Error GoTo ErrHandler
    
    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)
    
    With frm
        .Properties("Name") = "frmConfig"
        .Properties("Caption") = "Configuration"
        .Properties("Width") = 480
        .Properties("Height") = 320
        .Properties("StartUpPosition") = 1
    End With
    
    frm.CodeModule.AddFromString GetConfigFormCode()
    
    Dim ctrl As Object, t As Single, lc As Single
    lc = 15
    
    t = 10
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTitle": .Caption = "Configuration Systeme": .Left = lc: .Top = t: .Width = 420: .Height = 28: .Font.Size = 16: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 50
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblWDays": .Caption = "Working Days/Year:": .Left = lc: .Top = t: .Width = 120: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtWorkingDays": .Left = lc + 130: .Top = t: .Width = 60: .Height = 22: End With
    
    t = 80
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblOCost": .Caption = "Order Cost (DZD):": .Left = lc: .Top = t: .Width = 120: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtOrderCost": .Left = lc + 130: .Top = t: .Width = 60: .Height = 22: End With
    
    t = 110
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblHRate": .Caption = "Holding Rate:": .Left = lc: .Top = t: .Width = 120: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtHoldingRate": .Left = lc + 130: .Top = t: .Width = 60: .Height = 22: End With
    
    t = 140
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblLead": .Caption = "Lead Time (days):": .Left = lc: .Top = t: .Width = 120: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtLeadTime": .Left = lc + 130: .Top = t: .Width = 60: .Height = 22: End With
    
    t = 170
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTax": .Caption = "Tax Rate:": .Left = lc: .Top = t: .Width = 120: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtTaxRate": .Left = lc + 130: .Top = t: .Width = 60: .Height = 22: End With
    
    t = 210
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnSave": .Caption = "Enregistrer": .Left = lc: .Top = t: .Width = 100: .Height = 35: .BackColor = &H7F4600: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClose": .Caption = "Close": .Left = lc + 350: .Top = t: .Width = 80: .Height = 35: End With
    
    MsgBox "frmConfig created!" & vbCrLf & "Run: frmConfig.Show", vbInformation, "Done"
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetConfigFormCode() As String
    Dim c As String
    
    c = "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    c = c & "    txtWorkingDays.Text = ws.Range(""B2"").Value" & vbCrLf
    c = c & "    txtOrderCost.Text = ws.Range(""B4"").Value" & vbCrLf
    c = c & "    txtHoldingRate.Text = ws.Range(""B5"").Value" & vbCrLf
    c = c & "    txtLeadTime.Text = ws.Range(""B6"").Value" & vbCrLf
    c = c & "    txtTaxRate.Text = ws.Range(""B7"").Value" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnSave_Click()" & vbCrLf
    c = c & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    c = c & "    ws.Range(""B2"").Value = Val(txtWorkingDays.Text)" & vbCrLf
    c = c & "    ws.Range(""B4"").Value = Val(txtOrderCost.Text)" & vbCrLf
    c = c & "    ws.Range(""B5"").Value = Val(txtHoldingRate.Text)" & vbCrLf
    c = c & "    ws.Range(""B6"").Value = Val(txtLeadTime.Text)" & vbCrLf
    c = c & "    ws.Range(""B7"").Value = Val(txtTaxRate.Text)" & vbCrLf
    c = c & "    MsgBox ""Saved!"", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    
    GetConfigFormCode = c
End Function
