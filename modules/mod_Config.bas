Attribute VB_Name = "mod_Config"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Hardware Store Deployment - All parameters configurable
' ============================================================================

Option Explicit

' ============================================================================
' SHEET NAMES - Fixed for all deployments
' ============================================================================
Public Const SHEET_MOUVEMENTS As String = "MOUVEMENTS"
Public Const SHEET_ARTICLES As String = "ARTICLES"
Public Const SHEET_SYS_STRINGS As String = "SYS_STRINGS"
Public Const SHEET_AUDIT_LOG As String = "AUDIT_LOG"
Public Const SHEET_STAGING As String = "STAGING_BUFFER"
Public Const SHEET_FOURNISSEURS As String = "FOURNISSEURS"
Public Const SHEET_ACCUEIL As String = "ACCUEIL"
Public Const SHEET_CONFIG As String = "CONFIG"
Public Const SHEET_DASHBOARD As String = "DASHBOARD"
Public Const SHEET_ALERTE As String = "ALERTE_DASHBOARD"
Public Const SHEET_CALCULS As String = "CALCULS_EOQ"
Public Const SHEET_HISTORIQUE As String = "HISTORIQUE"
Public Const SHEET_RAPPORTS As String = "RAPPORTS"
Public Const SHEET_BUDGET As String = "BUDGET"
Public Const SHEET_BUDGET_REPORT As String = "BUDGET_REPORT"

' ============================================================================
' ARTICLES SHEET COLUMN INDICES
' ============================================================================
Public Const COL_ART_CODE As Long = 1
Public Const COL_ART_DESIGNATION As Long = 2
Public Const COL_ART_STOCK As Long = 3
Public Const COL_ART_SEUIL_MIN As Long = 4
Public Const COL_ART_CATEGORIE As Long = 5
Public Const COL_ART_CLASSE_ABC As Long = 6
Public Const COL_ART_STOCK_ACTUEL As Long = 7
Public Const COL_ART_PU As Long = 8
Public Const COL_ART_FOURNISSEUR As Long = 9
Public Const COL_ART_STOCK_SECURITE As Long = 10
Public Const COL_ART_NOTES As Long = 11
Public Const COL_ART_CMUP As Long = 12
Public Const COL_ART_METHODEAppro As Long = 13
Public Const COL_ART_DELAI As Long = 14

' ============================================================================
' MOUVEMENTS SHEET COLUMN INDICES
' ============================================================================
Public Const COL_MOUV_DATE As Long = 1
Public Const COL_MOUV_CODE_ARTICLE As Long = 2
Public Const COL_MOUV_DESIGNATION As Long = 3
Public Const COL_MOUV_TYPE As Long = 4
Public Const COL_MOUV_QTE As Long = 5
Public Const COL_MOUV_VALEUR As Long = 6
Public Const COL_MOUV_REF_DOC As Long = 7
Public Const COL_MOUV_PU As Long = 8
Public Const COL_MOUV_THIRD_PARTY As Long = 9
Public Const COL_MOUV_NOTES As Long = 10
Public Const COL_MOUV_USER As Long = 11
Public Const COL_MOUV_TIMESTAMP As Long = 12

' ============================================================================
' FOURNISSEURS SHEET COLUMN INDICES
' ============================================================================
Public Const COL_FOU_CODE As Long = 1
Public Const COL_FOU_RAISON_SOCIALE As Long = 2
Public Const COL_FOU_ADRESSE As Long = 3
Public Const COL_FOU_TELEPHONE As Long = 4
Public Const COL_FOU_NIF As Long = 5
Public Const COL_FOU_NIS As Long = 6
Public Const COL_FOU_RC As Long = 7
Public Const COL_FOU_ARTICLE_IMPOSITION As Long = 8

' ============================================================================
' SYSTEM PROPERTIES
' ============================================================================

Public Property Get SYS_TITLE() As String
    SYS_TITLE = "DSS Quincaillerie"
End Property

Public Property Get DOC_TYPE_BR() As String
    DOC_TYPE_BR = "Bon de R" & Chr(201) & "ception"
End Property

Public Property Get DOC_TYPE_BC() As String
    DOC_TYPE_BC = "Bon de Commande"
End Property

Public Property Get DOC_TYPE_BS As String
    DOC_TYPE_BS = "Bon de Sortie"
End Property

Public Property Get DOC_TYPE_DA As String
    DOC_TYPE_DA = "Demande d'Achat"
End Property

Public Property Get MASTER_PWD() As String
    MASTER_PWD = "erp_secure_pwd_2026"
End Property

Public Property Get APP_VERSION() As String
    APP_VERSION = "v14.0"
End Property

' ============================================================================
' BACKWARD COMPATIBILITY - Old method names used by original v13 modules
' ============================================================================

Public Function GetWorkingDaysPerYear() As Integer
    GetWorkingDaysPerYear = WORKING_DAYS_PER_YEAR
End Function

Public Function GetOrderCost() As Double
    GetOrderCost = ORDER_COST
End Function

Public Function GetHoldingRate() As Double
    GetHoldingRate = HOLDING_RATE
End Function

Public Function GetLeadTimeDays() As Integer
    GetLeadTimeDays = LEAD_TIME_DEFAULT
End Function

Public Function GetTaxRate() As Double
    GetTaxRate = TAX_RATE
End Function

Public Function GetCurrencySymbol() As String
    GetCurrencySymbol = CURRENCY_SYMBOL
End Function

Public Function GetBusinessName() As String
    GetBusinessName = BUSINESS_NAME
End Function

Public Function GetMasterPassword() As String
    GetMasterPassword = MASTER_PWD
End Function

' ============================================================================
' CONFIG SHEET PARAMETERS - Read from CONFIG sheet (fallback to defaults)
' ============================================================================

Public Property Get WORKING_DAYS_PER_YEAR() As Integer
    WORKING_DAYS_PER_YEAR = ReadConfigInt("WORKING_DAYS", 300)
End Property

Public Property Get OBSERVATION_DAYS() As Integer
    OBSERVATION_DAYS = ReadConfigInt("OBSERVATION_DAYS", 90)
End Property

Public Property Get ORDER_COST() As Double
    ORDER_COST = ReadConfigDouble("ORDER_COST", 300)
End Property

Public Property Get HOLDING_RATE() As Double
    HOLDING_RATE = ReadConfigDouble("HOLDING_RATE", 0.2)
End Property

Public Property Get LEAD_TIME_DEFAULT() As Integer
    LEAD_TIME_DEFAULT = ReadConfigInt("LEAD_TIME", 2)
End Property

Public Property Get TAX_RATE() As Double
    TAX_RATE = ReadConfigDouble("TAX_RATE", 0.19)
End Property

Public Property Get PU_INCLUDES_TVA() As Boolean
    PU_INCLUDES_TVA = ReadConfigBool("PU_INCLUDES_TVA", True)
End Property

Public Property Get INCLUDE_FREIGHT_IN_CMUP() As Boolean
    INCLUDE_FREIGHT_IN_CMUP = ReadConfigBool("INCLUDE_FREIGHT_IN_CMUP", False)
End Property

Public Property Get CURRENCY_SYMBOL() As String
    CURRENCY_SYMBOL = ReadConfigString("CURRENCY", "DZD")
End Property

Public Property Get BUSINESS_NAME() As String
    BUSINESS_NAME = ReadConfigString("BUSINESS_NAME", "Quincaillerie")
End Property

Public Property Get BUSINESS_ADDRESS() As String
    BUSINESS_ADDRESS = ReadConfigString("BUSINESS_ADDRESS", "Algerie")
End Property

Public Property Get BUSINESS_PHONE() As String
    BUSINESS_PHONE = ReadConfigString("BUSINESS_PHONE", "049 00 00 00")
End Property

Public Property Get BUSINESS_NIF() As String
    BUSINESS_NIF = ReadConfigString("BUSINESS_NIF", "000100000000000")
End Property

Public Property Get BUSINESS_NIS() As String
    BUSINESS_NIS = ReadConfigString("BUSINESS_NIS", "00100000000000")
End Property

Public Property Get BUSINESS_RC() As String
    BUSINESS_RC = ReadConfigString("BUSINESS_RC", "00/00-0000000A00")
End Property

' ----------------------------------------------------------------------------
' FIRST RUN STATE
' ----------------------------------------------------------------------------
' TRUE until the setup wizard has been completed once. Defaults to TRUE so that
' a workbook whose CONFIG sheet has no FIRST_RUN row - every build shipped
' before this flow existed - still presents the wizard on the next open.
Public Property Get IS_FIRST_RUN() As Boolean
    IS_FIRST_RUN = ReadConfigBool("FIRST_RUN", True)
End Property

Public Sub MarkFirstRunComplete()
    WriteConfig "FIRST_RUN", "FALSE"
End Sub

Public Sub MarkFirstRunPending()
    WriteConfig "FIRST_RUN", "TRUE"
End Sub

' ----------------------------------------------------------------------------
' EFFECTIVE OBSERVATION WINDOW
' ----------------------------------------------------------------------------
' OBSERVATION_DAYS is a thesis-era constant: the demo generator produces
' exactly 90 days of movements, so dividing total outflows by 90 was correct
' by construction. For a real store it is wrong from day one - a shop three
' weeks into using the system would have its daily consumption understated
' roughly fourfold, and the stockout projection on the DASHBOARD would report
' about four times more runway than it actually has.
'
' This derives the window from the data instead: the span actually covered by
' MOUVEMENTS. Falls back to the configured value when there is nothing to
' measure, so demo workbooks and empty workbooks both behave as before.
Public Function ObservationDaysEffective() As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_MOUVEMENTS)
    On Error GoTo 0
    If ws Is Nothing Then ObservationDaysEffective = OBSERVATION_DAYS: Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_MOUV_DATE).End(xlUp).Row
    If lastRow < 2 Then ObservationDaysEffective = OBSERVATION_DAYS: Exit Function

    Dim minD As Double, maxD As Double, v As Variant, i As Long, seen As Boolean
    For i = 2 To lastRow
        v = ws.Cells(i, COL_MOUV_DATE).Value
        If IsDate(v) Then
            If Not seen Then
                minD = CDbl(CDate(v)): maxD = minD: seen = True
            Else
                If CDbl(CDate(v)) < minD Then minD = CDbl(CDate(v))
                If CDbl(CDate(v)) > maxD Then maxD = CDbl(CDate(v))
            End If
        End If
    Next i

    If Not seen Then ObservationDaysEffective = OBSERVATION_DAYS: Exit Function

    ' Inclusive span: a single day of movements is one day of observation.
    Dim span As Long: span = CLng(maxD - minD) + 1
    If span < 1 Then span = 1
    ObservationDaysEffective = span
End Function

' ============================================================================
' CONFIG SHEET READERS - Read from CONFIG sheet with fallback
' ============================================================================
' Public rather than Private: mod_FirstRun reads arbitrary keys to prefill the
' setup wizard and to test FIRST_RUN, and a caller outside this module cannot
' reach a Private member. Every one of these resolves its key with
' Application.Match on column A, so none of them depends on row position.
' ============================================================================

Public Function ReadConfigInt(ByVal key As String, ByVal defaultVal As Integer) As Integer
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then ReadConfigInt = defaultVal: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(key, ws.Range("A:A"), 0)
    If IsError(foundRow) Then
        ReadConfigInt = defaultVal
    Else
        Dim val As Variant: val = ws.Cells(foundRow, 2).Value
        If IsNumeric(val) Then ReadConfigInt = CInt(val) Else ReadConfigInt = defaultVal
    End If
End Function

Public Function ReadConfigDouble(ByVal key As String, ByVal defaultVal As Double) As Double
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then ReadConfigDouble = defaultVal: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(key, ws.Range("A:A"), 0)
    If IsError(foundRow) Then
        ReadConfigDouble = defaultVal
    Else
        Dim val As Variant: val = ws.Cells(foundRow, 2).Value
        If IsNumeric(val) Then ReadConfigDouble = CDbl(val) Else ReadConfigDouble = defaultVal
    End If
End Function

Public Function ReadConfigString(ByVal key As String, ByVal defaultVal As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then ReadConfigString = defaultVal: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(key, ws.Range("A:A"), 0)
    If IsError(foundRow) Then
        ReadConfigString = defaultVal
    Else
        Dim val As String: val = Trim(CStr(ws.Cells(foundRow, 2).Value))
        If val <> "" Then ReadConfigString = val Else ReadConfigString = defaultVal
    End If
End Function

Public Function ReadConfigBool(ByVal key As String, ByVal defaultVal As Boolean) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then ReadConfigBool = defaultVal: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(key, ws.Range("A:A"), 0)
    If IsError(foundRow) Then
        ReadConfigBool = defaultVal
    Else
        Dim val As String: val = UCase(Trim(CStr(ws.Cells(foundRow, 2).Value)))
        If val = "TRUE" Or val = "1" Or val = "YES" Or val = "OUI" Then
            ReadConfigBool = True
        ElseIf val = "FALSE" Or val = "0" Or val = "NO" Or val = "NON" Then
            ReadConfigBool = False
        Else
            ReadConfigBool = defaultVal
        End If
    End If
End Function

' ============================================================================
' WRITE CONFIG - Write configuration to CONFIG sheet
' ============================================================================

Public Sub WriteConfig(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ws.Unprotect Password:=MASTER_PWD
    
    Dim foundRow As Variant
    foundRow = Application.Match(key, ws.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(lastRow, 1).Value = key
        ws.Cells(lastRow, 2).Value = value
    Else
        ws.Cells(foundRow, 2).Value = value
    End If
    
    ws.Protect Password:=MASTER_PWD, UserInterfaceOnly:=True
End Sub

' ============================================================================
' SEED DEFAULT CONFIG - Fill in any missing parameter, preserve every existing
' one
' ============================================================================
' This used to delete rows 2:last and rewrite all sixteen keys, which meant
' calling it after setup silently destroyed the business identity the owner had
' just typed in. That made it a routine nobody could safely call, and in fact
' nothing did - it was dead code.
'
' It is now an upsert that only fills gaps: an existing value is never
' overwritten and no row is ever deleted, so it is safe to call on every open
' and it is what brings a pre-existing workbook up to date when new parameters
' are added. For a deliberate factory reset, use ResetConfigToDefaults.
'
' OBSERVATION_DAYS and SEASON are deliberately absent. Both were thesis
' parameters: SEASON was never read anywhere, and OBSERVATION_DAYS is now
' derived from the movement history by ObservationDaysEffective.
' ============================================================================

Public Sub SeedDefaultConfig()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    EnsureConfigHeader ws

    ' Business identity is seeded blank on purpose. Shipping a plausible but
    ' fictitious NIF/NIS/RC is how the thesis placeholders ended up on
    ' documents; an empty row makes the parameter discoverable in the CONFIG
    ' sheet while leaving it obvious that it still needs filling in.
    SeedIfMissing "FIRST_RUN", "TRUE"
    SeedIfMissing "WORKING_DAYS", 300
    SeedIfMissing "ORDER_COST", 300
    SeedIfMissing "HOLDING_RATE", 0.2
    SeedIfMissing "LEAD_TIME", 2
    SeedIfMissing "TAX_RATE", 0.19
    SeedIfMissing "CURRENCY", "DZD"
    SeedIfMissing "PU_INCLUDES_TVA", "TRUE"
    SeedIfMissing "INCLUDE_FREIGHT_IN_CMUP", "FALSE"
    SeedIfMissing "BUSINESS_NAME", ""
    SeedIfMissing "BUSINESS_ADDRESS", ""
    SeedIfMissing "BUSINESS_PHONE", ""
    SeedIfMissing "BUSINESS_NIF", ""
    SeedIfMissing "BUSINESS_NIS", ""
    SeedIfMissing "BUSINESS_RC", ""

    Debug.Print "[Config] Missing defaults seeded; existing values preserved"
End Sub

' Writes the key only when it has no row yet. WriteConfig handles the
' unprotect/reprotect and the keyed upsert.
Private Sub SeedIfMissing(ByVal key As String, ByVal defaultVal As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim foundRow As Variant
    foundRow = Application.Match(key, ws.Range("A:A"), 0)
    If IsError(foundRow) Then WriteConfig key, defaultVal
End Sub

Private Sub EnsureConfigHeader(ByVal ws As Worksheet)
    If Trim(CStr(ws.Cells(1, 1).Value)) <> "" Then Exit Sub

    ws.Unprotect Password:=MASTER_PWD
    ws.Cells(1, 1).Value = "Parameter"
    ws.Cells(1, 2).Value = "Value"
    ws.Cells(1, 3).Value = "Description"
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(0, 70, 127)
    ws.Range("A1:C1").Font.Color = RGB(255, 255, 255)
    ws.Columns("A:C").AutoFit
    ws.Protect Password:=MASTER_PWD, UserInterfaceOnly:=True
End Sub

' ============================================================================
' RESET CONFIG TO DEFAULTS - Destructive, explicit, and re-arms the wizard
' ============================================================================
' The old wipe-and-reseed behaviour, kept available but named for what it does
' so it cannot be reached for by accident. Clears every parameter, reseeds the
' defaults, and sets FIRST_RUN back to TRUE so the wizard runs again.
' ============================================================================

Public Sub ResetConfigToDefaults(Optional ByVal confirm As Boolean = True)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If confirm Then
        Dim m As String
        m = "Reinitialiser toute la configuration ?" & vbCrLf & vbCrLf & _
            "Le nom commercial, l'adresse, le NIF, le NIS et le RC seront effaces." & vbCrLf & _
            "L'assistant de configuration sera relance a la prochaine ouverture." & vbCrLf & vbCrLf & _
            "Les articles et les mouvements ne sont pas touches."
        If MsgBox(m, vbExclamation + vbYesNo + vbDefaultButton2, SYS_TITLE) = vbNo Then Exit Sub
    End If

    ws.Unprotect Password:=MASTER_PWD
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr > 1 Then ws.Rows("2:" & lr).Delete
    ws.Protect Password:=MASTER_PWD, UserInterfaceOnly:=True

    SeedDefaultConfig
    MarkFirstRunPending
    Debug.Print "[Config] Configuration reset to defaults; FIRST_RUN re-armed"
End Sub
