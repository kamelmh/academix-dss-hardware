Attribute VB_Name = "mod_BuildArticleEditor"
Option Explicit

Public Sub BuildArticleEditorForm()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    
    On Error Resume Next
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    vbProj.VBComponents.Remove vbProj.VBComponents("frmArticleEditor")
    DoEvents
    On Error GoTo ErrHandler
    
    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)
    
    With frm
        .Properties("Name") = "frmArticleEditor"
        .Properties("Caption") = "Article Editor"
        .Properties("Width") = 580
        .Properties("Height") = 480
        .Properties("StartUpPosition") = 1
    End With
    
    frm.CodeModule.AddFromString GetArticleFormCode()
    
    Dim ctrl As Object
    Dim t As Single, lc As Single, rc As Single
    lc = 10: rc = 290
    
    t = 8
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTitle": .Caption = "Gestion des Articles": .Left = lc: .Top = t: .Width = 540: .Height = 28: .Font.Size = 14: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 42
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblCode": .Caption = "Code:": .Left = lc: .Top = t: .Width = 50: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtCode": .Left = 65: .Top = t: .Width = 100: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnSearch": .Caption = "...": .Left = 170: .Top = t: .Width = 30: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblCatLabel": .Caption = "Categorie:": .Left = rc: .Top = t: .Width = 60: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl: .Name = "cboCategorie": .Left = rc + 65: .Top = t: .Width = 100: .Height = 22: End With
    
    t = 70
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblDesig": .Caption = "Designation:": .Left = lc: .Top = t: .Width = 70: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtDesignation": .Left = 85: .Top = t: .Width = 440: .Height = 22: End With
    
    t = 98
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblSupLabel": .Caption = "Fournisseur:": .Left = lc: .Top = t: .Width = 70: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl: .Name = "cboFournisseur": .Left = 85: .Top = t: .Width = 200: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblUniteLabel": .Caption = "Unite:": .Left = rc: .Top = t: .Width = 40: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtUnite": .Left = rc + 45: .Top = t: .Width = 60: .Height = 22: End With
    
    t = 126
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblPrixLabel": .Caption = "PU (DZD):": .Left = lc: .Top = t: .Width = 60: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtPU": .Left = 75: .Top = t: .Width = 100: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblStockActuelLabel": .Caption = "Stock Actuel:": .Left = 185: .Top = t: .Width = 70: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblStockActuel": .Caption = "0": .Left = 260: .Top = t: .Width = 80: .Height = 20: .ForeColor = &HFF0000: End With
    
    t = 154
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblSeuilLabel": .Caption = "Seuil Min:": .Left = lc: .Top = t: .Width = 60: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtSeuilMin": .Left = 75: .Top = t: .Width = 80: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblSecurLabel": .Caption = "Stock Securite:": .Left = 165: .Top = t: .Width = 80: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtStockSecurite": .Left = 250: .Top = t: .Width = 80: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblDelaiLabel": .Caption = "Delai (j):": .Left = rc: .Top = t: .Width = 50: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtDelai": .Left = rc + 55: .Top = t: .Width = 60: .Height = 22: End With
    
    t = 182
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblMethLabel": .Caption = "Methode:": .Left = lc: .Top = t: .Width = 55: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl: .Name = "cboMethode": .Left = 70: .Top = t: .Width = 120: .Height = 22: End With
    
    t = 210
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblNotes": .Caption = "Notes:": .Left = lc: .Top = t: .Width = 45: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtNotes": .Left = 60: .Top = t: .Width = 465: .Height = 40: .MultiLine = True: End With
    
    t = 260
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblCMUPLabel": .Caption = "CMUP:": .Left = lc: .Top = t: .Width = 40: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblCMUP": .Caption = "0.00": .Left = 55: .Top = t: .Width = 100: .Height = 20: .ForeColor = &H7F4600: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblABCLabel": .Caption = "ABC:": .Left = 165: .Top = t: .Width = 30: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblABC": .Caption = "-": .Left = 200: .Top = t: .Width = 40: .Height = 20: .ForeColor = &H7F4600: End With
    
    t = 295
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnSave": .Caption = "Save": .Left = lc: .Top = t: .Width = 100: .Height = 35: .BackColor = &HD48700: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnNew": .Caption = "New": .Left = 120: .Top = t: .Width = 100: .Height = 35: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnDelete": .Caption = "Delete": .Left = 230: .Top = t: .Width = 100: .Height = 35: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClose": .Caption = "Close": .Left = 440: .Top = t: .Width = 100: .Height = 35: End With
    
    MsgBox "frmArticleEditor created!" & vbCrLf & "Run: frmArticleEditor.Show", vbInformation, "Done"
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetArticleFormCode() As String
    Dim c As String
    
    c = "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Call LoadCategories" & vbCrLf
    c = c & "    Call LoadSuppliers" & vbCrLf
    c = c & "    cboMethode.Clear" & vbCrLf
    c = c & "    cboMethode.AddItem ""COMMANDER""" & vbCrLf
    c = c & "    cboMethode.AddItem ""COMMANDE_AUTO""" & vbCrLf
    c = c & "    cboMethode.ListIndex = 0" & vbCrLf
    c = c & "    lblStockActuel.Caption = ""0""" & vbCrLf
    c = c & "    lblCMUP.Caption = ""0.00""" & vbCrLf
    c = c & "    lblABC.Caption = ""-""" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub LoadCategories()" & vbCrLf
    c = c & "    cboCategorie.Clear" & vbCrLf
    c = c & "    cboCategorie.AddItem ""FER""" & vbCrLf
    c = c & "    cboCategorie.AddItem ""CEM""" & vbCrLf
    c = c & "    cboCategorie.AddItem ""PVC""" & vbCrLf
    c = c & "    cboCategorie.AddItem ""ELEC""" & vbCrLf
    c = c & "    cboCategorie.AddItem ""OUTIL""" & vbCrLf
    c = c & "    cboCategorie.AddItem ""PEINT""" & vbCrLf
    c = c & "    cboCategorie.AddItem ""CARO""" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub LoadSuppliers()" & vbCrLf
    c = c & "    cboFournisseur.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(""FOURNISSEURS"")" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        cboFournisseur.AddItem ws.Cells(i, 1).Value & "" - "" & ws.Cells(i, 2).Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnSearch_Click()" & vbCrLf
    c = c & "    Dim s As String" & vbCrLf
    c = c & "    s = InputBox(""Enter article code:"", ""Search"")" & vbCrLf
    c = c & "    If Len(Trim(s)) > 0 Then" & vbCrLf
    c = c & "        txtCode.Value = UCase(Trim(s))" & vbCrLf
    c = c & "        Call LoadArticle" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub LoadArticle()" & vbCrLf
    c = c & "    Dim code As String" & vbCrLf
    c = c & "    code = UCase(Trim(txtCode.Value))" & vbCrLf
    c = c & "    If Len(code) = 0 Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(""ARTICLES"")" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim r As Variant" & vbCrLf
    c = c & "    r = Application.Match(code, ws.Range(""A:A""), 0)" & vbCrLf
    c = c & "    If IsError(r) Then" & vbCrLf
    c = c & "        lblStockActuel.Caption = ""NEW""" & vbCrLf
    c = c & "        lblCMUP.Caption = ""0.00""" & vbCrLf
    c = c & "        lblABC.Caption = ""-""" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    txtDesignation.Value = ws.Cells(r, 2).Value" & vbCrLf
    c = c & "    txtSeuilMin.Value = ws.Cells(r, 4).Value" & vbCrLf
    c = c & "    Dim cat As String: cat = ws.Cells(r, 5).Value" & vbCrLf
    c = c & "    Dim ci As Long: ci = cboCategorie.ListCount - 1" & vbCrLf
    c = c & "    Dim j As Long" & vbCrLf
    c = c & "    For j = 0 To ci" & vbCrLf
    c = c & "        If UCase(cboCategorie.List(j)) = UCase(cat) Then cboCategorie.ListIndex = j: Exit For" & vbCrLf
    c = c & "    Next j" & vbCrLf
    c = c & "    lblABC.Caption = ws.Cells(r, 6).Value" & vbCrLf
    c = c & "    lblStockActuel.Caption = Format(ws.Cells(r, 3).Value, ""#,##0"")" & vbCrLf
    c = c & "    txtPU.Value = Format(ws.Cells(r, 8).Value, ""#,##0.00"")" & vbCrLf
    c = c & "    Dim sup As String: sup = ws.Cells(r, 9).Value" & vbCrLf
    c = c & "    Dim si As Long: si = cboFournisseur.ListCount - 1" & vbCrLf
    c = c & "    For j = 0 To si" & vbCrLf
    c = c & "        If UCase(Left(cboFournisseur.List(j), Len(sup))) = UCase(sup) Then cboFournisseur.ListIndex = j: Exit For" & vbCrLf
    c = c & "    Next j" & vbCrLf
    c = c & "    txtStockSecurite.Value = ws.Cells(r, 10).Value" & vbCrLf
    c = c & "    txtNotes.Value = ws.Cells(r, 11).Value" & vbCrLf
    c = c & "    lblCMUP.Caption = Format(ws.Cells(r, 12).Value, ""#,##0.00"")" & vbCrLf
    c = c & "    Dim meth As String: meth = ws.Cells(r, 13).Value" & vbCrLf
    c = c & "    If meth = ""COMMANDE_AUTO"" Then cboMethode.ListIndex = 1 Else cboMethode.ListIndex = 0" & vbCrLf
    c = c & "    txtDelai.Value = ws.Cells(r, 14).Value" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnSave_Click()" & vbCrLf
    c = c & "    If Len(Trim(txtCode.Value)) = 0 Then MsgBox ""Indiquez le code article."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    If Len(Trim(txtDesignation.Value)) = 0 Then MsgBox ""Indiquez la designation."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    If MsgBox(""Save article?"", vbQuestion + vbYesNo) = vbNo Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(""ARTICLES"")" & vbCrLf
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
    c = c & "    ws.Cells(nr, 2).Value = Trim(txtDesignation.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 4).Value = Val(Replace(txtSeuilMin.Value, "","", """"))" & vbCrLf
    c = c & "    ws.Cells(nr, 5).Value = cboCategorie.Value" & vbCrLf
    c = c & "    ws.Cells(nr, 8).Value = Val(Replace(txtPU.Value, "","", """"))" & vbCrLf
    c = c & "    Dim supCode As String: supCode = Left(cboFournisseur.Value, InStr(cboFournisseur.Value, "" - "") - 1)" & vbCrLf
    c = c & "    ws.Cells(nr, 9).Value = supCode" & vbCrLf
    c = c & "    ws.Cells(nr, 10).Value = Val(Replace(txtStockSecurite.Value, "","", """"))" & vbCrLf
    c = c & "    ws.Cells(nr, 11).Value = Trim(txtNotes.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 13).Value = cboMethode.Value" & vbCrLf
    c = c & "    ws.Cells(nr, 14).Value = Val(Replace(txtDelai.Value, "","", """"))" & vbCrLf
    c = c & "    ws.Protect Password:=mod_Config.MASTER_PWD" & vbCrLf
    c = c & "    MsgBox ""Article enregistre!"", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnNew_Click()" & vbCrLf
    c = c & "    txtCode.Value = """"" & vbCrLf
    c = c & "    txtDesignation.Value = """"" & vbCrLf
    c = c & "    cboCategorie.ListIndex = 0" & vbCrLf
    c = c & "    cboFournisseur.ListIndex = 0" & vbCrLf
    c = c & "    txtPU.Value = """"" & vbCrLf
    c = c & "    txtSeuilMin.Value = """"" & vbCrLf
    c = c & "    txtStockSecurite.Value = """"" & vbCrLf
    c = c & "    txtDelai.Value = """"" & vbCrLf
    c = c & "    txtNotes.Value = """"" & vbCrLf
    c = c & "    cboMethode.ListIndex = 0" & vbCrLf
    c = c & "    lblStockActuel.Caption = ""0""" & vbCrLf
    c = c & "    lblCMUP.Caption = ""0.00""" & vbCrLf
    c = c & "    lblABC.Caption = ""-""" & vbCrLf
    c = c & "    txtCode.SetFocus" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnDelete_Click()" & vbCrLf
    c = c & "    Dim code As String: code = UCase(Trim(txtCode.Value))" & vbCrLf
    c = c & "    If Len(code) = 0 Then MsgBox ""Indiquez le code d'abord."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    If MsgBox(""Supprimer l'article "" & code & ""?"", vbQuestion + vbYesNo) = vbNo Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(""ARTICLES"")" & vbCrLf
    c = c & "    Dim r As Variant" & vbCrLf
    c = c & "    r = Application.Match(code, ws.Range(""A:A""), 0)" & vbCrLf
    c = c & "    If IsError(r) Then MsgBox ""Article introuvable."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    ws.Unprotect Password:=mod_Config.MASTER_PWD" & vbCrLf
    c = c & "    ws.Rows(r).Delete" & vbCrLf
    c = c & "    ws.Protect Password:=mod_Config.MASTER_PWD" & vbCrLf
    c = c & "    Call btnNew_Click" & vbCrLf
    c = c & "    MsgBox ""Article supprime."", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    
    GetArticleFormCode = c
End Function
