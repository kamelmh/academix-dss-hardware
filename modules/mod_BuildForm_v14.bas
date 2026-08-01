Attribute VB_Name = "mod_BuildForm"
Option Explicit

Public Sub BuildStockEntryForm()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    
    On Error Resume Next
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    vbProj.VBComponents.Remove vbProj.VBComponents("frmStockEntry")
    DoEvents
    On Error GoTo ErrHandler
    
    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)
    
    With frm
        .Properties("Name") = "frmStockEntry"
        .Properties("Caption") = "Mouvement de Stock"
        .Properties("Width") = 560
        .Properties("Height") = 450
        .Properties("StartUpPosition") = 1
    End With
    
    Dim code As String
    code = GetFormCode()
    frm.CodeModule.AddFromString code
    
    Dim ctrl As Object
    Dim t As Single
    
    t = 8
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTitle": .Caption = "Mouvement de Stock": .Left = 10: .Top = t: .Width = 520: .Height = 28: .Font.Size = 14: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 42
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblBarcode": .Caption = "Barcode:": .Left = 10: .Top = t: .Width = 55: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtBarcode": .Left = 70: .Top = t: .Width = 200: .Height = 22: .BackColor = &HFFFFC0: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblScanHint": .Caption = "(Scan or type barcode)": .Left = 280: .Top = t: .Width = 150: .Height = 20: .ForeColor = &H808080: .Font.Size = 8: End With
    
    t = 72
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblDate": .Caption = "Date:": .Left = 10: .Top = t: .Width = 50: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtDate": .Left = 65: .Top = t: .Width = 100: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblType": .Caption = "Type:": .Left = 280: .Top = t: .Width = 40: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl: .Name = "cboType": .Left = 325: .Top = t: .Width = 100: .Height = 22: End With
    
    t = 102
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblCode": .Caption = "Code:": .Left = 10: .Top = t: .Width = 50: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtCode": .Left = 65: .Top = t: .Width = 100: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnSearch": .Caption = "...": .Left = 170: .Top = t: .Width = 30: .Height = 22: End With
    
    t = 132
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblDesLabel": .Caption = "Designation:": .Left = 10: .Top = t: .Width = 70: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblDesignation": .Caption = "": .Left = 85: .Top = t: .Width = 440: .Height = 22: .ForeColor = &H7F4600: .BackColor = &HF0F0F0: End With
    
    t = 158
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblStockLabel": .Caption = "Stock:": .Left = 10: .Top = t: .Width = 50: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblStockInfo": .Caption = "-": .Left = 65: .Top = t: .Width = 440: .Height = 20: .ForeColor = &HFF0000: .BackColor = &HF0F0F0: End With
    
    t = 156
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblQty": .Caption = "Quantite:": .Left = 10: .Top = t: .Width = 60: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtQuantite": .Left = 75: .Top = t: .Width = 100: .Height = 22: End With
    
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblPU": .Caption = "PU:": .Left = 280: .Top = t: .Width = 30: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtPU": .Left = 315: .Top = t: .Width = 100: .Height = 22: End With
    
    t = 186
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblValLabel": .Caption = "Valeur:": .Left = 10: .Top = t: .Width = 50: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblValeur": .Caption = "0.00 DZD": .Left = 65: .Top = t: .Width = 440: .Height = 20: .Font.Bold = True: .ForeColor = &H7F4600: End With
    
    t = 216
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblRef": .Caption = "Ref Doc:": .Left = 10: .Top = t: .Width = 55: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtRefDoc": .Left = 70: .Top = t: .Width = 455: .Height = 22: End With
    
    t = 246
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTiers": .Caption = "Tiers:": .Left = 10: .Top = t: .Width = 50: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtTiers": .Left = 65: .Top = t: .Width = 460: .Height = 22: End With
    
    t = 276
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblNotes": .Caption = "Notes:": .Left = 10: .Top = t: .Width = 50: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtNotes": .Left = 65: .Top = t: .Width = 460: .Height = 22: End With
    
    t = 310
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnSave": .Caption = "Save": .Left = 10: .Top = t: .Width = 100: .Height = 35: .BackColor = &HD48700: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnNew": .Caption = "New": .Left = 120: .Top = t: .Width = 100: .Height = 35: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnDelete": .Caption = "Delete": .Left = 230: .Top = t: .Width = 100: .Height = 35: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClose": .Caption = "Close": .Left = 420: .Top = t: .Width = 100: .Height = 35: End With
    
    MsgBox "frmStockEntry created!" & vbCrLf & "Double-click it to view." & vbCrLf & "Then run: frmStockEntry.Show", vbInformation, "Done"
    Exit Sub
    
ErrHandler:
    MsgBox "Erreur " & Err.Number & ": " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

Private Function GetFormCode() As String
    Dim c As String
    c = "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    txtDate.Value = Format(Date, ""DD/MM/YYYY"")" & vbCrLf
    c = c & "    cboType.Clear" & vbCrLf
    c = c & "    cboType.AddItem ""ENTREE""" & vbCrLf
    c = c & "    cboType.AddItem ""SORTIE""" & vbCrLf
    c = c & "    cboType.ListIndex = 0" & vbCrLf
    c = c & "    lblDesignation.Caption = """"" & vbCrLf
    c = c & "    lblStockInfo.Caption = ""-""" & vbCrLf
    c = c & "    lblValeur.Caption = ""0.00 DZD""" & vbCrLf
    c = c & "    txtBarcode.SetFocus" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub txtBarcode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    c = c & "    If KeyCode = 13 Then" & vbCrLf
    c = c & "        Dim bc As String: bc = Trim(txtBarcode.Value)" & vbCrLf
    c = c & "        If Len(bc) = 0 Then Exit Sub" & vbCrLf
    c = c & "        Dim ws As Worksheet" & vbCrLf
    c = c & "        On Error Resume Next: Set ws = ThisWorkbook.Sheets(""BARCODES""): On Error GoTo 0" & vbCrLf
    c = c & "        If ws Is Nothing Then" & vbCrLf
    c = c & "            MsgBox ""Feuille BARCODES introuvable. Generez d'abord les codes-barres."", vbExclamation" & vbCrLf
    c = c & "            Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "        Dim i As Long" & vbCrLf
    c = c & "        For i = 2 To lr" & vbCrLf
    c = c & "            If CStr(ws.Cells(i, 1).Value) = bc Then" & vbCrLf
    c = c & "                txtCode.Value = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "                Call txtCode_AfterUpdate" & vbCrLf
    c = c & "                txtQuantite.SetFocus" & vbCrLf
    c = c & "                txtBarcode.Value = """"" & vbCrLf
    c = c & "                Exit Sub" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        Next i" & vbCrLf
    c = c & "        MsgBox ""Code-barres introuvable: "" & bc, vbExclamation" & vbCrLf
    c = c & "        txtBarcode.Value = """"" & vbCrLf
    c = c & "        txtBarcode.SetFocus" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnSearch_Click()" & vbCrLf
    c = c & "    Dim s As String" & vbCrLf
    c = c & "    s = InputBox(""Enter article code:"", ""Search"")" & vbCrLf
    c = c & "    If Len(Trim(s)) > 0 Then" & vbCrLf
    c = c & "        txtCode.Value = UCase(Trim(s))" & vbCrLf
    c = c & "        Call txtCode_AfterUpdate" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub txtCode_AfterUpdate()" & vbCrLf
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
    c = c & "        lblDesignation.Caption = ""Article not found""" & vbCrLf
    c = c & "        lblStockInfo.Caption = ""-""" & vbCrLf
    c = c & "        txtCode.SetFocus" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    lblDesignation.Caption = ws.Cells(r, 2).Value" & vbCrLf
    c = c & "    txtPU.Value = Format(ws.Cells(r, 8).Value, ""#,##0.00"")" & vbCrLf
    c = c & "    Dim stock As Double" & vbCrLf
    c = c & "    stock = mod_StockEngine.GetArticleStock(code)" & vbCrLf
    c = c & "    lblStockInfo.Caption = ""Stock: "" & Format(stock, ""#,##0"")" & vbCrLf
    c = c & "    Call CalcValeur" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub txtQuantite_AfterUpdate()" & vbCrLf
    c = c & "    Call CalcValeur" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub txtPU_AfterUpdate()" & vbCrLf
    c = c & "    Call CalcValeur" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub CalcValeur()" & vbCrLf
    c = c & "    Dim v As Double, q As Double, p As Double" & vbCrLf
    c = c & "    q = Val(Replace(txtQuantite.Value, "","", """"))" & vbCrLf
    c = c & "    p = Val(Replace(txtPU.Value, "","", """"))" & vbCrLf
    c = c & "    v = q * p" & vbCrLf
    c = c & "    lblValeur.Caption = Format(v, ""#,##0.00"") & "" DZD""" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnSave_Click()" & vbCrLf
    c = c & "    On Error GoTo ErrorHandler" & vbCrLf
    c = c & "    Dim qty As Double, pu As Double" & vbCrLf
    c = c & "    qty = Val(Replace(txtQuantite.Value, "","", """"))" & vbCrLf
    c = c & "    pu = Val(Replace(txtPU.Value, "","", """"))" & vbCrLf
    c = c & "    If Len(Trim(txtCode.Value)) = 0 Then MsgBox ""Entrez le code article."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    If qty <= 0 Then MsgBox ""La quantite doit etre > 0."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    If Not mod_FormHelpers.ArticleExists(Trim(txtCode.Value)) Then" & vbCrLf
    c = c & "        MsgBox ""Article inexistant: "" & Trim(txtCode.Value), vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If cboType.Value = ""SORTIE"" Then" & vbCrLf
    c = c & "        Dim curStock As Double: curStock = mod_StockEngine.GetArticleStock(Trim(txtCode.Value))" & vbCrLf
    c = c & "        If qty > curStock Then" & vbCrLf
    c = c & "            MsgBox ""Stock insuffisant. Disponibilite: "" & Format(curStock, ""#,##0"") & "" unite(s)."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(""Enregistrer le mouvement?"", vbQuestion + vbYesNo) = vbNo Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(""MOUVEMENTS"")" & vbCrLf
    c = c & "    ws.Unprotect Password:=mod_Config.MASTER_PWD" & vbCrLf
    c = c & "    Dim nr As Long" & vbCrLf
    c = c & "    nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    c = c & "    ws.Cells(nr, 1).Value = CDate(txtDate.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 1).NumberFormat = ""DD/MM/YYYY""" & vbCrLf
    c = c & "    ws.Cells(nr, 2).Value = UCase(Trim(txtCode.Value))" & vbCrLf
    c = c & "    ws.Cells(nr, 3).Value = lblDesignation.Caption" & vbCrLf
    c = c & "    ws.Cells(nr, 4).Value = cboType.Value" & vbCrLf
    c = c & "    ws.Cells(nr, 5).Value = qty" & vbCrLf
    c = c & "    ws.Cells(nr, 6).Value = qty * pu" & vbCrLf
    c = c & "    ws.Cells(nr, 7).Value = Trim(txtRefDoc.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 8).Value = pu" & vbCrLf
    c = c & "    ws.Cells(nr, 9).Value = Trim(txtTiers.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 10).Value = Trim(txtNotes.Value)" & vbCrLf
    c = c & "    ws.Cells(nr, 11).Value = Environ(""USERNAME"")" & vbCrLf
    c = c & "    ws.Cells(nr, 12).Value = Now" & vbCrLf
    c = c & "    ws.Cells(nr, 12).NumberFormat = ""DD/MM/YYYY HH:MM:SS""" & vbCrLf
    c = c & "    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True" & vbCrLf
    c = c & "    If cboType.Value = ""ENTREE"" Then" & vbCrLf
    c = c & "        Call mod_StockEngine.UpdateArticleStockBalance(txtCode.Value, ""IN"", CLng(qty))" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        Call mod_StockEngine.UpdateArticleStockBalance(txtCode.Value, ""OUT"", CLng(qty))" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    MsgBox ""Enregistre! Stock: "" & Format(mod_StockEngine.GetArticleStock(txtCode.Value), ""#,##0""), vbInformation" & vbCrLf
    c = c & "    Call btnNew_Click" & vbCrLf
    c = c & "    Exit Sub" & vbCrLf
    c = c & "ErrorHandler:" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True" & vbCrLf
    c = c & "    MsgBox ""Erreur ecriture: "" & Err.Description, vbCritical, mod_Config.SYS_TITLE" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnNew_Click()" & vbCrLf
    c = c & "    txtDate.Value = Format(Date, ""DD/MM/YYYY"")" & vbCrLf
    c = c & "    cboType.ListIndex = 0" & vbCrLf
    c = c & "    txtCode.Value = """"" & vbCrLf
    c = c & "    lblDesignation.Caption = """"" & vbCrLf
    c = c & "    lblStockInfo.Caption = ""-""" & vbCrLf
    c = c & "    txtQuantite.Value = """"" & vbCrLf
    c = c & "    txtPU.Value = """"" & vbCrLf
    c = c & "    lblValeur.Caption = ""0.00 DZD""" & vbCrLf
    c = c & "    txtRefDoc.Value = """"" & vbCrLf
    c = c & "    txtTiers.Value = """"" & vbCrLf
    c = c & "    txtNotes.Value = """"" & vbCrLf
    c = c & "    txtCode.SetFocus" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnDelete_Click()" & vbCrLf
    c = c & "    MsgBox ""Selectionnez une ligne dans le tableau a supprimer."", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & vbCrLf
    c = c & "Private Sub btnClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    
    GetFormCode = c
End Function
