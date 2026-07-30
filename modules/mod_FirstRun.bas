Attribute VB_Name = "mod_FirstRun"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' First-run flow - turns the thesis demo into a system a real store can open
' ============================================================================
'
' Opening the workbook used to present the thesis case study: 40 pre-loaded
' articles, 90 days of generated movements, and a placeholder business identity.
' A store owner needs the opposite - an empty ledger carrying their own
' identity - while the demo data still has to stay reachable for training and
' for showing the system to someone.
'
' This module owns that decision. On the first open it collects the business
' identity and the operating parameters, then either leaves the data sheets
' empty (the default) or loads the demo set on request.
'
' Two things worth knowing before editing:
'
' 1. Forms in this project are built at runtime by the mod_Build* modules, so
'    no .frm exists in the repo. Naming a form directly in code - frmFirstRun.Show -
'    makes the whole project fail to compile on any workbook where that form has
'    not been built yet. Every launcher here therefore resolves the form by
'    string through VBA.UserForms.Add, which is the same approach
'    mod_AccueilButtons uses.
'
' 2. Config is only ever written through mod_Config.WriteConfig. It upserts on
'    the key and handles the unprotect/reprotect, so nothing here depends on a
'    row position or on UserInterfaceOnly surviving a workbook reopen.
' ============================================================================

Option Explicit

Private Const FORM_NAME As String = "frmFirstRun"

' Workbook_Open and MasterSetup can both reach FirstRunCheck. Without this the
' owner would answer the wizard twice in one session.
Private mCheckedThisSession As Boolean

' ============================================================================
' ENTRY POINTS
' ============================================================================

' Called from ThisWorkbook.Workbook_Open and from the end of MasterSetup.
' Silent and cheap when setup is already done.
Public Sub FirstRunCheck()
    On Error GoTo ErrHandler

    If mCheckedThisSession Then Exit Sub

    ' Brings a workbook created before this flow existed up to date, and creates
    ' the CONFIG rows on a blank sheet. Never overwrites an existing value.
    mod_Config.SeedDefaultConfig

    If Not mod_Config.IS_FIRST_RUN Then
        mCheckedThisSession = True
        Exit Sub
    End If

    mCheckedThisSession = True
    ShowWizard

    Exit Sub

ErrHandler:
    ' A failure here must never block the workbook from opening.
    Debug.Print "[FirstRun] FirstRunCheck error " & Err.Number & ": " & Err.Description
End Sub

' Manual entry point - reopens the wizard at any time, whether or not first run
' has been completed. Wired to a macro button and safe to run mid-life.
Public Sub FirstRunSetup()
    mCheckedThisSession = True
    mod_Config.SeedDefaultConfig
    ShowWizard
End Sub

Public Function IsFirstRun() As Boolean
    IsFirstRun = mod_Config.IS_FIRST_RUN
End Function

' ============================================================================
' WIZARD LAUNCH
' ============================================================================
' Prefers the built form. Falls back to a prompt sequence so that a workbook
' whose forms have not been built is still configurable by its owner rather
' than being told to run a developer macro.
Private Sub ShowWizard()
    If FormExists(FORM_NAME) Then
        On Error GoTo FormFailed
        VBA.UserForms.Add(FORM_NAME).Show
        Exit Sub
    End If

    ' Try to build it once - harmless when mod_BuildFirstRun is absent.
    On Error Resume Next
    Application.Run "mod_BuildFirstRun.BuildFirstRunForm", True
    On Error GoTo 0
    DoEvents

    If FormExists(FORM_NAME) Then
        On Error GoTo FormFailed
        VBA.UserForms.Add(FORM_NAME).Show
        Exit Sub
    End If

    PromptWizard
    Exit Sub

FormFailed:
    Debug.Print "[FirstRun] form launch failed (" & Err.Description & "); using prompts"
    PromptWizard
End Sub

Private Function FormExists(ByVal formName As String) As Boolean
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(formName)
    FormExists = Not (vbComp Is Nothing)
    On Error GoTo 0
End Function

' ============================================================================
' FALLBACK WIZARD - no UserForm required
' ============================================================================
Private Sub PromptWizard()
    Dim msg As String
    msg = "Configuration initiale" & vbCrLf & vbCrLf & _
          "Quelques informations sur votre commerce, puis le systeme est pret." & vbCrLf & _
          "Laissez vide pour conserver la valeur actuelle."
    MsgBox msg, vbInformation, mod_Config.SYS_TITLE

    Dim bName As String, bAddr As String, bPhone As String
    Dim bNIF As String, bNIS As String, bRC As String

    bName = AskText("Nom commercial de l'etablissement :", mod_Config.BUSINESS_NAME)
    If Trim(bName) = "" Then
        MsgBox "Configuration annulee. L'assistant sera propose a la prochaine ouverture.", _
               vbExclamation, mod_Config.SYS_TITLE
        Exit Sub
    End If

    bAddr = AskText("Adresse :", mod_Config.BUSINESS_ADDRESS)
    bPhone = AskText("Telephone :", mod_Config.BUSINESS_PHONE)
    bNIF = AskText("NIF (15 chiffres) :", mod_Config.BUSINESS_NIF)
    bNIS = AskText("NIS (15 chiffres) :", mod_Config.BUSINESS_NIS)
    bRC = AskText("RC (format 00/00-0000000A00) :", mod_Config.BUSINESS_RC)

    Dim wDays As String, oCost As String, hRate As String
    Dim lTime As String, tRate As String, curr As String

    wDays = AskText("Jours ouvres par an :", CStr(mod_Config.WORKING_DAYS_PER_YEAR))
    oCost = AskText("Cout d'une commande (" & mod_Config.CURRENCY_SYMBOL & ") :", CStr(mod_Config.ORDER_COST))
    hRate = AskText("Taux de possession du stock (0,2 pour 20%) :", CStr(mod_Config.HOLDING_RATE))
    lTime = AskText("Delai de livraison (jours) :", CStr(mod_Config.LEAD_TIME_DEFAULT))
    tRate = AskText("Taux de TVA (0,19 pour 19%) :", CStr(mod_Config.TAX_RATE))
    curr = AskText("Devise :", mod_Config.CURRENCY_SYMBOL)

    If Not SaveFirstRunConfig(bName, bAddr, bPhone, bNIF, bNIS, bRC, _
                              wDays, oCost, hRate, lTime, tRate, curr) Then
        Exit Sub
    End If

    Dim wantDemo As VbMsgBoxResult
    wantDemo = MsgBox("Charger les 40 articles de demonstration ?" & vbCrLf & vbCrLf & _
                      "Non  - demarrer avec un stock vide (recommande pour un usage reel)." & vbCrLf & _
                      "Oui  - charger les donnees de demonstration pour se former.", _
                      vbQuestion + vbYesNo + vbDefaultButton2, mod_Config.SYS_TITLE)

    Dim useDemo As Boolean: useDemo = (wantDemo = vbYes)
    CompleteFirstRun useDemo
End Sub

Private Function AskText(ByVal prompt As String, ByVal current As String) As String
    Dim answer As String
    answer = InputBox(prompt, mod_Config.SYS_TITLE, current)
    If StrPtr(answer) = 0 Then          ' Cancel pressed
        AskText = current
    ElseIf Trim(answer) = "" Then       ' left blank - keep what was there
        AskText = current
    Else
        AskText = Trim(answer)
    End If
End Function

' ============================================================================
' SAVE - called by frmFirstRun and by the fallback wizard
' ============================================================================
' Every value arrives as text because both callers are text inputs. Returns
' False when validation rejected something, in which case nothing was written
' and the caller should keep the wizard open.
Public Function SaveFirstRunConfig(ByVal businessName As String, _
                                   ByVal businessAddress As String, _
                                   ByVal businessPhone As String, _
                                   ByVal businessNIF As String, _
                                   ByVal businessNIS As String, _
                                   ByVal businessRC As String, _
                                   ByVal workingDays As String, _
                                   ByVal orderCost As String, _
                                   ByVal holdingRate As String, _
                                   ByVal leadTime As String, _
                                   ByVal taxRate As String, _
                                   ByVal currencySymbol As String) As Boolean

    SaveFirstRunConfig = False

    ' ---- identity ----
    businessName = Trim(businessName)
    If businessName = "" Then
        MsgBox "Le nom commercial est obligatoire : il figure sur les factures et les bons.", _
               vbExclamation, mod_Config.SYS_TITLE
        Exit Function
    End If

    If Not CheckFiscalId(businessNIF, "NIF") Then Exit Function
    If Not CheckFiscalId(businessNIS, "NIS") Then Exit Function

    ' ---- operating parameters ----
    Dim nWorkingDays As Double, nOrderCost As Double, nHoldingRate As Double
    Dim nLeadTime As Double, nTaxRate As Double

    nWorkingDays = ParseNumber(workingDays)
    nOrderCost = ParseNumber(orderCost)
    nHoldingRate = ParseNumber(holdingRate)
    nLeadTime = ParseNumber(leadTime)
    nTaxRate = ParseNumber(taxRate)

    ' WORKING_DAYS divides annual demand in mod_StockEngine (avgDaily =
    ' AnnualDemand / WORKING_DAYS_PER_YEAR), which then feeds every reorder
    ' point. Zero would raise a division error on the next calculation.
    If nWorkingDays < 1 Or nWorkingDays > 366 Then
        MsgBox "Jours ouvres par an : indiquez une valeur entre 1 et 366." & vbCrLf & _
               "Un commerce ferme le vendredi ouvre environ 300 jours.", _
               vbExclamation, mod_Config.SYS_TITLE
        Exit Function
    End If

    If nOrderCost <= 0 Then
        MsgBox "Le cout de commande doit etre superieur a zero : c'est le numerateur de la formule de Wilson.", _
               vbExclamation, mod_Config.SYS_TITLE
        Exit Function
    End If

    ' HOLDING_RATE is the EOQ divisor by way of holdingCostH = PU * HOLDING_RATE.
    ' Zero produces a division by zero on the next reorder calculation, so it is
    ' refused here rather than at the point of failure.
    nHoldingRate = NormaliseRate(nHoldingRate, "taux de possession")
    If nHoldingRate <= 0 Then
        MsgBox "Le taux de possession doit etre superieur a zero : il divise la formule de Wilson.", _
               vbExclamation, mod_Config.SYS_TITLE
        Exit Function
    End If

    If nLeadTime < 0 Then
        MsgBox "Le delai de livraison ne peut pas etre negatif.", vbExclamation, mod_Config.SYS_TITLE
        Exit Function
    End If

    nTaxRate = NormaliseRate(nTaxRate, "taux de TVA")
    If nTaxRate < 0 Or nTaxRate >= 1 Then
        MsgBox "Taux de TVA invalide. Indiquez 0,19 (ou 19) pour 19%.", _
               vbExclamation, mod_Config.SYS_TITLE
        Exit Function
    End If

    currencySymbol = Trim(currencySymbol)
    If currencySymbol = "" Then currencySymbol = "DZD"

    ' ---- write ----
    mod_Config.WriteConfig "BUSINESS_NAME", businessName
    mod_Config.WriteConfig "BUSINESS_ADDRESS", Trim(businessAddress)
    mod_Config.WriteConfig "BUSINESS_PHONE", Trim(businessPhone)
    mod_Config.WriteConfig "BUSINESS_NIF", DigitsOnly(businessNIF)
    mod_Config.WriteConfig "BUSINESS_NIS", DigitsOnly(businessNIS)
    mod_Config.WriteConfig "BUSINESS_RC", Trim(businessRC)

    mod_Config.WriteConfig "WORKING_DAYS", CLng(nWorkingDays)
    mod_Config.WriteConfig "ORDER_COST", nOrderCost
    mod_Config.WriteConfig "HOLDING_RATE", nHoldingRate
    mod_Config.WriteConfig "LEAD_TIME", CLng(nLeadTime)
    mod_Config.WriteConfig "TAX_RATE", nTaxRate
    mod_Config.WriteConfig "CURRENCY", currencySymbol

    SaveFirstRunConfig = True
End Function

' ============================================================================
' COMPLETE - clean start or demo, then close out first run
' ============================================================================
Public Sub CompleteFirstRun(ByVal loadDemoData As Boolean)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    If loadDemoData Then
        ' GenerateDemoData clears CONFIG as part of its reset, so it snapshots
        ' and restores the parameters around that clear. The values just
        ' collected therefore survive; see mod_DemoData.
        mod_DemoData.GenerateDemoData
    Else
        PrepareCleanStart True
    End If

    ' Set last: if anything above failed, first run stays pending and the owner
    ' gets the wizard again rather than a half-configured system.
    mod_Config.MarkFirstRunComplete

    On Error Resume Next
    mod_Dashboard.RefreshDashboard
    mod_AccueilButtons.RefreshAccueilKPIs
    ThisWorkbook.Sheets(mod_Config.SHEET_ACCUEIL).Activate
    On Error GoTo 0

    Application.ScreenUpdating = True

    If loadDemoData Then
        MsgBox "Configuration enregistree." & vbCrLf & vbCrLf & _
               "40 articles de demonstration ont ete charges." & vbCrLf & _
               "Modifiez-les pour refleter votre inventaire reel, ou relancez" & vbCrLf & _
               "l'assistant et choisissez un demarrage a vide.", _
               vbInformation, mod_Config.SYS_TITLE
    Else
        MsgBox "Configuration enregistree pour " & mod_Config.BUSINESS_NAME & "." & vbCrLf & vbCrLf & _
               "Le stock est vide. Commencez par saisir vos fournisseurs," & vbCrLf & _
               "puis vos articles, puis vos receptions.", _
               vbInformation, mod_Config.SYS_TITLE
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur pendant l'initialisation : " & Err.Description & vbCrLf & vbCrLf & _
           "La configuration n'a pas ete validee ; l'assistant sera propose de nouveau.", _
           vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' CLEAN START - empty the ledger, keep the headers and the configuration
' ============================================================================
' Clears the transactional and derived sheets only. CONFIG is deliberately
' untouched: it holds the identity that was just entered.
Public Sub PrepareCleanStart(Optional ByVal silent As Boolean = False)
    On Error GoTo ErrHandler

    If Not silent Then
        Dim m As String
        m = "Vider le stock et l'historique ?" & vbCrLf & vbCrLf & _
            "Les articles, mouvements et fournisseurs seront effaces." & vbCrLf & _
            "La configuration et l'identite de l'etablissement sont conservees."
        If MsgBox(m, vbExclamation + vbYesNo + vbDefaultButton2, mod_Config.SYS_TITLE) = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Ledgers and document registers only - every one of these carries exactly
    ' one header row, so deleting from row 2 down is safe.
    '
    ' DASHBOARD, ALERTE_DASHBOARD, CALCULS_EOQ and RAPPORTS are deliberately
    ' absent. They are rendered rather than recorded, and mod_Dashboard writes a
    ' merged title on row 1 with the column headers on row 2, so deleting from
    ' row 2 would eat the header. They are regenerated from the ledgers instead -
    ' CompleteFirstRun calls RefreshDashboard once the clearing is done.
    Dim targets As Variant
    targets = Array(mod_Config.SHEET_ARTICLES, mod_Config.SHEET_MOUVEMENTS, _
                    mod_Config.SHEET_FOURNISSEURS, mod_Config.SHEET_STAGING, _
                    mod_Config.SHEET_AUDIT_LOG, _
                    "BON_RECEPTION", "FACTURES", "BARCODES", "BONS_COMMANDE")

    Dim i As Long
    For i = LBound(targets) To UBound(targets)
        ClearDataRows CStr(targets(i))
    Next i

    Application.ScreenUpdating = True

    If Not silent Then
        MsgBox "Stock et historique vides. La configuration est conservee.", _
               vbInformation, mod_Config.SYS_TITLE
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur pendant le vidage : " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' Deletes rows 2 and below, leaving row 1 and its formatting intact. Missing
' sheets are skipped: deployments do not all carry the same optional sheets.
Private Sub ClearDataRows(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Unprotect Password:=mod_Config.MASTER_PWD

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete

    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0

    Debug.Print "[FirstRun] cleared " & sheetName
End Sub

' ============================================================================
' DEMO ON DEMAND - reachable after first run without re-running the wizard
' ============================================================================
Public Sub LoadDemoDataKeepingIdentity()
    Dim m As String
    m = "Charger les donnees de demonstration ?" & vbCrLf & vbCrLf & _
        "40 articles, 9 fournisseurs et 90 jours de mouvements remplaceront" & vbCrLf & _
        "le contenu actuel. Votre configuration et votre identite sont conservees."
    If MsgBox(m, vbExclamation + vbYesNo + vbDefaultButton2, mod_Config.SYS_TITLE) = vbNo Then Exit Sub

    mod_DemoData.GenerateDemoData
End Sub

' ============================================================================
' INPUT HELPERS
' ============================================================================

' Val() reads only a period as the decimal separator, so on the French locale
' this system runs on, a typed 0,2 would come back as 0 - and that value is the
' EOQ divisor. Normalise the separator before converting.
Public Function ParseNumber(ByVal s As String) As Double
    s = Trim(Replace(Replace(s, " ", ""), ",", "."))
    If s = "" Then
        ParseNumber = 0
    Else
        ParseNumber = Val(s)
    End If
End Function

' Accepts a rate written either as a fraction (0,19) or as a percentage (19),
' which is how people actually type these. Confirms the reading rather than
' converting silently.
Private Function NormaliseRate(ByVal value As Double, ByVal label As String) As Double
    NormaliseRate = value
    If value <= 1 Then Exit Function

    Dim asFraction As Double: asFraction = value / 100
    Dim m As String
    m = "Vous avez saisi " & Format(value, "0.####") & " pour le " & label & "." & vbCrLf & vbCrLf & _
        "Faut-il le lire comme " & Format(asFraction, "0.####") & " (" & Format(value, "0.##") & "%) ?"
    If MsgBox(m, vbQuestion + vbYesNo + vbDefaultButton1, mod_Config.SYS_TITLE) = vbYes Then
        NormaliseRate = asFraction
    End If
End Function

' Digits are required; length is only advisory. Algerian identifiers are
' normally 15 digits, but refusing anything else outright would block an owner
' whose paperwork disagrees with our assumption.
Private Function CheckFiscalId(ByVal raw As String, ByVal label As String) As Boolean
    CheckFiscalId = True
    Dim cleaned As String: cleaned = DigitsOnly(raw)

    If cleaned = "" Then Exit Function      ' optional

    If cleaned <> Replace(Trim(raw), " ", "") Then
        MsgBox "Le " & label & " ne doit contenir que des chiffres.", _
               vbExclamation, mod_Config.SYS_TITLE
        CheckFiscalId = False
        Exit Function
    End If

    If Len(cleaned) <> 15 Then
        Dim m As String
        m = "Le " & label & " comporte " & Len(cleaned) & " chiffres ; " & _
            "il en compte habituellement 15." & vbCrLf & vbCrLf & "Conserver cette valeur ?"
        If MsgBox(m, vbQuestion + vbYesNo + vbDefaultButton1, mod_Config.SYS_TITLE) = vbNo Then
            CheckFiscalId = False
        End If
    End If
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    DigitsOnly = out
End Function
