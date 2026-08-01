Attribute VB_Name = "mod_BuildDashboard"
Option Explicit

Public Sub BuildDashboardForm()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    
    On Error Resume Next
    DoEvents
    Dim existingDash As Object
    Set existingDash = vbProj.VBComponents("frmDashboard")
    If Not existingDash Is Nothing Then
        vbProj.VBComponents.Remove existingDash
    End If
    DoEvents
    On Error GoTo ErrHandler
    
    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)
    
    With frm
        .Properties("Name") = "frmDashboard"
        .Properties("Caption") = "Dashboard - DSS v14.0"
        .Properties("Width") = 680
        .Properties("Height") = 480
        .Properties("StartUpPosition") = 1
    End With
    
    frm.CodeModule.AddFromString GetDashboardFormCode()
    
    Dim ctrl As Object, t As Single, lc As Single
    lc = 15
    
    t = 8
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTitle": .Caption = "DSS Quincaillerie": .Left = lc: .Top = t: .Width = 640: .Height = 32: .Font.Size = 18: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 48
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblSub": .Caption = "v14.0 - Tableau de Bord": .Left = lc: .Top = t: .Width = 300: .Height = 20: .Font.Size = 10: .ForeColor = &H808080: End With
    
    t = 80
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblArtLabel": .Caption = "Articles:": .Left = lc: .Top = t: .Width = 70: .Height = 22: .Font.Size = 11: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblArtCount": .Caption = "0": .Left = lc + 75: .Top = t: .Width = 60: .Height = 22: .Font.Size = 14: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblSupLabel": .Caption = "Fournisseurs:": .Left = lc + 160: .Top = t: .Width = 90: .Height = 22: .Font.Size = 11: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblSupCount": .Caption = "0": .Left = lc + 255: .Top = t: .Width = 60: .Height = 22: .Font.Size = 14: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblMvtLabel": .Caption = "Mouvements:": .Left = lc + 340: .Top = t: .Width = 90: .Height = 22: .Font.Size = 11: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblMvtCount": .Caption = "0": .Left = lc + 435: .Top = t: .Width = 60: .Height = 22: .Font.Size = 14: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 115
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblValLabel": .Caption = "Valeur Stock:": .Left = lc: .Top = t: .Width = 90: .Height = 22: .Font.Size = 11: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblStockValue": .Caption = "0 DZD": .Left = lc + 95: .Top = t: .Width = 150: .Height = 22: .Font.Size = 14: .Font.Bold = True: .ForeColor = &HFF0000: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblRuptLabel": .Caption = "Ruptures:": .Left = lc + 280: .Top = t: .Width = 70: .Height = 22: .Font.Size = 11: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblRuptCount": .Caption = "0": .Left = lc + 355: .Top = t: .Width = 60: .Height = 22: .Font.Size = 14: .Font.Bold = True: .ForeColor = &HFF: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblAlertLabel": .Caption = "Alertes:": .Left = lc + 440: .Top = t: .Width = 60: .Height = 22: .Font.Size = 11: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblAlertCount": .Caption = "0": .Left = lc + 505: .Top = t: .Width = 60: .Height = 22: .Font.Size = 14: .Font.Bold = True: .ForeColor = &HFF8000: End With
    
    t = 160
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblQuickTitle": .Caption = "Quick Actions:": .Left = lc: .Top = t: .Width = 100: .Height = 22: .Font.Size = 11: .Font.Bold = True: End With
    
    t = 190
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnStockEntry": .Caption = "Stock Entry": .Left = lc: .Top = t: .Width = 120: .Height = 40: .BackColor = &HD48700: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnArticleEditor": .Caption = "Articles": .Left = lc + 130: .Top = t: .Width = 120: .Height = 40: .BackColor = &H7F4600: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnSupplierEditor": .Caption = "Fournisseurs": .Left = lc + 260: .Top = t: .Width = 120: .Height = 40: .BackColor = &H7F4600: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnConfig": .Caption = "Config": .Left = lc + 390: .Top = t: .Width = 100: .Height = 40: End With
    
    t = 245
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblAlertsTitle": .Caption = "Alertes:": .Left = lc: .Top = t: .Width = 100: .Height = 22: .Font.Size = 11: .Font.Bold = True: End With
    
    t = 270
    Set ctrl = frm.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl: .Name = "lstAlerts": .Left = lc: .Top = t: .Width = 620: .Height = 120: End With
    
    t = 400
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnRefresh": .Caption = "Refresh": .Left = lc: .Top = t: .Width = 100: .Height = 30: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClose": .Caption = "Close": .Left = lc + 530: .Top = t: .Width = 100: .Height = 30: End With
    
    MsgBox "frmDashboard created!" & vbCrLf & "Run: frmDashboard.Show", vbInformation, "Done"
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetDashboardFormCode() As String
    Dim c As String
    
    c = "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Call RefreshData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub RefreshData()" & vbCrLf
    c = c & "    Dim wsA As Worksheet, wsM As Worksheet, wsS As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set wsA = ThisWorkbook.Sheets(""ARTICLES"")" & vbCrLf
    c = c & "    Set wsM = ThisWorkbook.Sheets(""MOUVEMENTS"")" & vbCrLf
    c = c & "    Set wsS = ThisWorkbook.Sheets(""FOURNISSEURS"")" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    Dim artCount As Long, mvtCount As Long, supCount As Long" & vbCrLf
    c = c & "    Dim stockVal As Double, ruptCount As Long, alertCount As Long" & vbCrLf
    c = c & "    If Not wsA Is Nothing Then artCount = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row - 1" & vbCrLf
    c = c & "    If Not wsM Is Nothing Then mvtCount = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row - 1" & vbCrLf
    c = c & "    If Not wsS Is Nothing Then supCount = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row - 1" & vbCrLf
    c = c & "    If Not wsA Is Nothing Then" & vbCrLf
    c = c & "        Dim lr As Long: lr = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "        Dim i As Long" & vbCrLf
    c = c & "        For i = 2 To lr" & vbCrLf
    c = c & "            stockVal = stockVal + (Val(wsA.Cells(i, 3).Value) * IIf(Val(wsA.Cells(i, 12).Value) > 0, Val(wsA.Cells(i, 12).Value), Val(wsA.Cells(i, 8).Value)))" & vbCrLf
    c = c & "            If Val(wsA.Cells(i, 3).Value) <= Val(wsA.Cells(i, 4).Value) Then ruptCount = ruptCount + 1" & vbCrLf
    c = c & "            If Val(wsA.Cells(i, 3).Value) <= Val(wsA.Cells(i, 10).Value) Then alertCount = alertCount + 1" & vbCrLf
    c = c & "        Next i" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    lblArtCount.Caption = artCount" & vbCrLf
    c = c & "    lblSupCount.Caption = supCount" & vbCrLf
    c = c & "    lblMvtCount.Caption = mvtCount" & vbCrLf
    c = c & "    lblStockValue.Caption = Format(stockVal, ""#,##0"") & "" DZD""" & vbCrLf
    c = c & "    lblRuptCount.Caption = ruptCount" & vbCrLf
    c = c & "    lblAlertCount.Caption = alertCount" & vbCrLf
    c = c & "    lstAlerts.Clear" & vbCrLf
    c = c & "    If Not wsA Is Nothing Then" & vbCrLf
    c = c & "        For i = 2 To lr" & vbCrLf
    c = c & "            Dim stk As Double: stk = Val(wsA.Cells(i, 3).Value)" & vbCrLf
    c = c & "            Dim seuil As Double: seuil = Val(wsA.Cells(i, 4).Value)" & vbCrLf
    c = c & "            Dim sec As Double: sec = Val(wsA.Cells(i, 10).Value)" & vbCrLf
    c = c & "            If stk <= seuil Then" & vbCrLf
    c = c & "                lstAlerts.AddItem wsA.Cells(i, 1).Value & "" - "" & wsA.Cells(i, 2).Value & "" [RUPTURE] Stock: "" & Format(stk, ""#,##0"") & "" < Seuil: "" & Format(seuil, ""#,##0"")" & vbCrLf
    c = c & "            ElseIf stk <= sec Then" & vbCrLf
    c = c & "                lstAlerts.AddItem wsA.Cells(i, 1).Value & "" - "" & wsA.Cells(i, 2).Value & "" [ALERTE] Stock: "" & Format(stk, ""#,##0"") & "" < Securite: "" & Format(sec, ""#,##0"")" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        Next i" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If lstAlerts.ListCount = 0 Then lstAlerts.AddItem ""All articles OK""" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnStockEntry_Click()" & vbCrLf
    c = c & "    VBA.UserForms.Add(""frmStockEntry"").Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnArticleEditor_Click()" & vbCrLf
    c = c & "    VBA.UserForms.Add(""frmArticleEditor"").Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnSupplierEditor_Click()" & vbCrLf
    c = c & "    VBA.UserForms.Add(""frmSupplierEditor"").Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnConfig_Click()" & vbCrLf
    c = c & "    VBA.UserForms.Add(""frmConfig"").Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnRefresh_Click()" & vbCrLf
    c = c & "    Call RefreshData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    
    GetDashboardFormCode = c
End Function
