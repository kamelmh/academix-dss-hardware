Attribute VB_Name = "mod_BuildFirstRun"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Builds frmFirstRun - the two-page setup wizard shown on the first open
' ============================================================================
'
' Same approach as the other mod_Build* modules: the form is created at runtime
' through the VBIDE extensibility model, so there is no .frm in the repo and
' "Trust access to the VBA project object model" must be enabled.
'
' The code module below is assembled by GetFirstRunFormCode. That function is
' generated rather than hand-written - see tools/genbuilder.py - because getting
' the doubled quotes wrong by hand produces a form whose code module will not
' compile, and there is no way to notice that until it is opened.
'
' Page 1 is the business identity, page 2 the operating parameters. All controls
' sit flat on the form and the two groups are swapped by toggling Visible; no
' container control is used, matching the rest of this project.
' ============================================================================

Option Explicit

Public Sub BuildFirstRunForm(Optional ByVal silent As Boolean = False)
    On Error GoTo ErrHandler

    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject

    ' Remove any previous build. Guarded because the component legitimately does
    ' not exist the first time round.
    On Error Resume Next
    DoEvents
    vbProj.VBComponents.Remove vbProj.VBComponents("frmFirstRun")
    DoEvents
    On Error GoTo ErrHandler

    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)

    With frm
        .Properties("Name") = "frmFirstRun"
        .Properties("Caption") = "Configuration initiale - DSS Quincaillerie"
        .Properties("Width") = 500
        .Properties("Height") = 345
        .Properties("StartUpPosition") = 1
    End With

    frm.CodeModule.AddFromString GetFirstRunFormCode()

    Dim ctrl As Object

    ' ---- header ----
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTitle": .Caption = "Configuration initiale": .Left = 18: .Top = 14: .Width = 460: .Height = 26: .Font.Size = 15: .Font.Bold = True: .ForeColor = &H7F4600: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblStep": .Caption = "Etape 1 sur 2": .Left = 18: .Top = 42: .Width = 460: .Height = 18: .Font.Bold = True: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblHint": .Caption = "": .Left = 18: .Top = 60: .Width = 460: .Height = 18: .ForeColor = &H808080: End With

    ' ---- page 1 fields ----
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblName": .Caption = "Nom commercial *": .Left = 18: .Top = 88: .Width = 140: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtName": .Left = 165: .Top = 88: .Width = 320: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblAddr": .Caption = "Adresse": .Left = 18: .Top = 116: .Width = 140: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtAddr": .Left = 165: .Top = 116: .Width = 320: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblPhone": .Caption = "Telephone": .Left = 18: .Top = 144: .Width = 140: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtPhone": .Left = 165: .Top = 144: .Width = 320: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblNIF": .Caption = "NIF (15 chiffres)": .Left = 18: .Top = 172: .Width = 140: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtNIF": .Left = 165: .Top = 172: .Width = 320: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblNIS": .Caption = "NIS (15 chiffres)": .Left = 18: .Top = 200: .Width = 140: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtNIS": .Left = 165: .Top = 200: .Width = 320: .Height = 22: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblRC": .Caption = "RC (00/00-0000000A00)": .Left = 18: .Top = 228: .Width = 140: .Height = 20: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtRC": .Left = 165: .Top = 228: .Width = 320: .Height = 22: End With

    ' ---- page 2 fields ----
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblWDays": .Caption = "Jours ouvres par an": .Left = 18: .Top = 88: .Width = 140: .Height = 20: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtWDays": .Left = 165: .Top = 88: .Width = 320: .Height = 22: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblOCost": .Caption = "Cout d'une commande": .Left = 18: .Top = 116: .Width = 140: .Height = 20: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtOCost": .Left = 165: .Top = 116: .Width = 320: .Height = 22: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblHRate": .Caption = "Taux de possession": .Left = 18: .Top = 144: .Width = 140: .Height = 20: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtHRate": .Left = 165: .Top = 144: .Width = 320: .Height = 22: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblLead": .Caption = "Delai de livraison (jours)": .Left = 18: .Top = 172: .Width = 140: .Height = 20: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtLead": .Left = 165: .Top = 172: .Width = 320: .Height = 22: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblTax": .Caption = "Taux de TVA": .Left = 18: .Top = 200: .Width = 140: .Height = 20: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtTax": .Left = 165: .Top = 200: .Width = 320: .Height = 22: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Name = "lblCurr": .Caption = "Devise": .Left = 18: .Top = 228: .Width = 140: .Height = 20: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtCurr": .Left = 165: .Top = 228: .Width = 320: .Height = 22: .Visible = False: End With

    ' ---- navigation ----
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnBack": .Caption = "< Precedent": .Left = 18: .Top = 274: .Width = 90: .Height = 32: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnNext": .Caption = "Suivant >": .Left = 385: .Top = 274: .Width = 100: .Height = 32: .BackColor = &H7F4600: .ForeColor = &HFFFFFF: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnStart": .Caption = "Demarrer a vide": .Left = 335: .Top = 274: .Width = 150: .Height = 32: .BackColor = &H7F4600: .ForeColor = &HFFFFFF: .Visible = False: End With
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnDemo": .Caption = "Charger donnees demo": .Left = 150: .Top = 274: .Width = 175: .Height = 32: .Visible = False: End With

    If Not silent Then
        MsgBox "frmFirstRun cree." & vbCrLf & _
               "Lancez mod_FirstRun.FirstRunSetup pour l'ouvrir.", _
               vbInformation, "Done"
    End If
    Exit Sub

ErrHandler:
    If Not silent Then
        MsgBox "Erreur " & Err.Number & ": " & Err.Description, vbCritical, "Erreur"
    Else
        Debug.Print "[BuildFirstRun] " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Function GetFirstRunFormCode() As String
    Dim c As String

    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "' Page 1 collects the business identity, page 2 the operating parameters." & vbCrLf
    c = c & "' There is no container control: every field sits directly on the form and the" & vbCrLf
    c = c & "' two groups are swapped by toggling Visible, which is the same flat-control" & vbCrLf
    c = c & "' approach the other mod_Build* modules in this project use." & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    ' Prefill from the current configuration. On a genuinely fresh workbook the" & vbCrLf
    c = c & "    ' identity keys are blank, so these come back as the mod_Config fallbacks" & vbCrLf
    c = c & "    ' and the owner overwrites them." & vbCrLf
    c = c & "    txtName.Text = mod_Config.BUSINESS_NAME" & vbCrLf
    c = c & "    txtAddr.Text = mod_Config.BUSINESS_ADDRESS" & vbCrLf
    c = c & "    txtPhone.Text = mod_Config.BUSINESS_PHONE" & vbCrLf
    c = c & "    txtNIF.Text = mod_Config.BUSINESS_NIF" & vbCrLf
    c = c & "    txtNIS.Text = mod_Config.BUSINESS_NIS" & vbCrLf
    c = c & "    txtRC.Text = mod_Config.BUSINESS_RC" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "    txtWDays.Text = mod_Config.WORKING_DAYS_PER_YEAR" & vbCrLf
    c = c & "    txtOCost.Text = mod_Config.ORDER_COST" & vbCrLf
    c = c & "    txtHRate.Text = mod_Config.HOLDING_RATE" & vbCrLf
    c = c & "    txtLead.Text = mod_Config.LEAD_TIME_DEFAULT" & vbCrLf
    c = c & "    txtTax.Text = mod_Config.TAX_RATE" & vbCrLf
    c = c & "    txtCurr.Text = mod_Config.CURRENCY_SYMBOL" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "    ShowPage 1" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub ShowPage(ByVal n As Long)" & vbCrLf
    c = c & "    Dim p1 As Boolean, p2 As Boolean" & vbCrLf
    c = c & "    p1 = (n = 1)" & vbCrLf
    c = c & "    p2 = (n = 2)" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "    lblName.Visible = p1" & vbCrLf
    c = c & "    txtName.Visible = p1" & vbCrLf
    c = c & "    lblAddr.Visible = p1" & vbCrLf
    c = c & "    txtAddr.Visible = p1" & vbCrLf
    c = c & "    lblPhone.Visible = p1" & vbCrLf
    c = c & "    txtPhone.Visible = p1" & vbCrLf
    c = c & "    lblNIF.Visible = p1" & vbCrLf
    c = c & "    txtNIF.Visible = p1" & vbCrLf
    c = c & "    lblNIS.Visible = p1" & vbCrLf
    c = c & "    txtNIS.Visible = p1" & vbCrLf
    c = c & "    lblRC.Visible = p1" & vbCrLf
    c = c & "    txtRC.Visible = p1" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "    lblWDays.Visible = p2" & vbCrLf
    c = c & "    txtWDays.Visible = p2" & vbCrLf
    c = c & "    lblOCost.Visible = p2" & vbCrLf
    c = c & "    txtOCost.Visible = p2" & vbCrLf
    c = c & "    lblHRate.Visible = p2" & vbCrLf
    c = c & "    txtHRate.Visible = p2" & vbCrLf
    c = c & "    lblLead.Visible = p2" & vbCrLf
    c = c & "    txtLead.Visible = p2" & vbCrLf
    c = c & "    lblTax.Visible = p2" & vbCrLf
    c = c & "    txtTax.Visible = p2" & vbCrLf
    c = c & "    lblCurr.Visible = p2" & vbCrLf
    c = c & "    txtCurr.Visible = p2" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "    btnNext.Visible = p1" & vbCrLf
    c = c & "    btnBack.Visible = p2" & vbCrLf
    c = c & "    btnStart.Visible = p2" & vbCrLf
    c = c & "    btnDemo.Visible = p2" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "    If p1 Then" & vbCrLf
    c = c & "        lblStep.Caption = ""Etape 1 sur 2 - Identite de l'etablissement""" & vbCrLf
    c = c & "        lblHint.Caption = ""Ces informations figurent sur vos factures et vos bons.""" & vbCrLf
    c = c & "        txtName.SetFocus" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        lblStep.Caption = ""Etape 2 sur 2 - Parametres de gestion""" & vbCrLf
    c = c & "        lblHint.Caption = ""Les valeurs proposees conviennent a une quincaillerie. Modifiables plus tard.""" & vbCrLf
    c = c & "        txtWDays.SetFocus" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub btnNext_Click()" & vbCrLf
    c = c & "    If Len(Trim(txtName.Text)) = 0 Then" & vbCrLf
    c = c & "        MsgBox ""Indiquez le nom commercial avant de continuer."", vbExclamation, mod_Config.SYS_TITLE" & vbCrLf
    c = c & "        txtName.SetFocus" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    ShowPage 2" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub btnBack_Click()" & vbCrLf
    c = c & "    ShowPage 1" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "' Validation, normalisation and writing all live in mod_FirstRun so that the" & vbCrLf
    c = c & "' form and the no-form fallback wizard cannot drift apart. False means something" & vbCrLf
    c = c & "' was rejected and a message has already been shown, so the form stays open." & vbCrLf
    c = c & "Private Function PersistConfig() As Boolean" & vbCrLf
    c = c & "    PersistConfig = mod_FirstRun.SaveFirstRunConfig( _" & vbCrLf
    c = c & "        txtName.Text, txtAddr.Text, txtPhone.Text, _" & vbCrLf
    c = c & "        txtNIF.Text, txtNIS.Text, txtRC.Text, _" & vbCrLf
    c = c & "        txtWDays.Text, txtOCost.Text, txtHRate.Text, _" & vbCrLf
    c = c & "        txtLead.Text, txtTax.Text, txtCurr.Text)" & vbCrLf
    c = c & "End Function" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub btnStart_Click()" & vbCrLf
    c = c & "    If Not PersistConfig() Then Exit Sub" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    mod_FirstRun.CompleteFirstRun False" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub btnDemo_Click()" & vbCrLf
    c = c & "    Dim m As String" & vbCrLf
    c = c & "    m = ""Charger les 40 articles de demonstration ?"" & vbCrLf & vbCrLf & _" & vbCrLf
    c = c & "        ""Utile pour se former ou presenter le systeme."" & vbCrLf & _" & vbCrLf
    c = c & "        ""Pour un usage reel, preferez Demarrer a vide.""" & vbCrLf
    c = c & "    If MsgBox(m, vbQuestion + vbYesNo + vbDefaultButton2, mod_Config.SYS_TITLE) = vbNo Then Exit Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "    If Not PersistConfig() Then Exit Sub" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    mod_FirstRun.CompleteFirstRun True" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "' Closing the wizard leaves FIRST_RUN set, so it comes back on the next open" & vbCrLf
    c = c & "' rather than leaving the system half configured." & vbCrLf
    c = c & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    c = c & "    If CloseMode = vbFormControlMenu Then" & vbCrLf
    c = c & "        MsgBox ""Configuration interrompue."" & vbCrLf & _" & vbCrLf
    c = c & "               ""L'assistant sera propose de nouveau a la prochaine ouverture."", _" & vbCrLf
    c = c & "               vbInformation, mod_Config.SYS_TITLE" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf

    GetFirstRunFormCode = c
End Function
