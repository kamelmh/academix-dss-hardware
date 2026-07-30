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

Public Function GetObservationDays() As Integer
    GetObservationDays = OBSERVATION_DAYS
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

Public Property Get SEASON() As String
    SEASON = ReadConfigString("SEASON", "Printemps")
End Property

' ============================================================================
' CONFIG SHEET READERS - Read from CONFIG sheet with fallback
' ============================================================================

Private Function ReadConfigInt(ByVal key As String, ByVal defaultVal As Integer) As Integer
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

Private Function ReadConfigDouble(ByVal key As String, ByVal defaultVal As Double) As Double
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

Private Function ReadConfigString(ByVal key As String, ByVal defaultVal As String) As String
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

Private Function ReadConfigBool(ByVal key As String, ByVal defaultVal As Boolean) As Boolean
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
' SEED DEFAULT CONFIG - Initialize CONFIG sheet with default values
' ============================================================================

Public Sub SeedDefaultConfig()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ws.Unprotect Password:=MASTER_PWD
    
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr > 1 Then ws.Rows("2:" & lr).Delete
    
    Dim r As Long: r = 2
    ws.Cells(r, 1).Value = "WORKING_DAYS": ws.Cells(r, 2).Value = 300: r = r + 1
    ws.Cells(r, 1).Value = "OBSERVATION_DAYS": ws.Cells(r, 2).Value = 90: r = r + 1
    ws.Cells(r, 1).Value = "ORDER_COST": ws.Cells(r, 2).Value = 300: r = r + 1
    ws.Cells(r, 1).Value = "HOLDING_RATE": ws.Cells(r, 2).Value = 0.2: r = r + 1
    ws.Cells(r, 1).Value = "LEAD_TIME": ws.Cells(r, 2).Value = 2: r = r + 1
    ws.Cells(r, 1).Value = "TAX_RATE": ws.Cells(r, 2).Value = 0.19: r = r + 1
    ws.Cells(r, 1).Value = "CURRENCY": ws.Cells(r, 2).Value = "DZD": r = r + 1
    ws.Cells(r, 1).Value = "BUSINESS_NAME": ws.Cells(r, 2).Value = "Quincaillerie": r = r + 1
    ws.Cells(r, 1).Value = "BUSINESS_ADDRESS": ws.Cells(r, 2).Value = "Algerie": r = r + 1
    ws.Cells(r, 1).Value = "BUSINESS_PHONE": ws.Cells(r, 2).Value = "049 00 00 00": r = r + 1
    ws.Cells(r, 1).Value = "BUSINESS_NIF": ws.Cells(r, 2).Value = "000100000000000": r = r + 1
    ws.Cells(r, 1).Value = "BUSINESS_NIS": ws.Cells(r, 2).Value = "00100000000000": r = r + 1
    ws.Cells(r, 1).Value = "BUSINESS_RC": ws.Cells(r, 2).Value = "00/00-0000000A00": r = r + 1
    ws.Cells(r, 1).Value = "SEASON": ws.Cells(r, 2).Value = "Printemps": r = r + 1
    ws.Cells(r, 1).Value = "PU_INCLUDES_TVA": ws.Cells(r, 2).Value = "TRUE": r = r + 1
    ws.Cells(r, 1).Value = "INCLUDE_FREIGHT_IN_CMUP": ws.Cells(r, 2).Value = "FALSE": r = r + 1
    
    ws.Protect Password:=MASTER_PWD, UserInterfaceOnly:=True
    Debug.Print "[Config] Default configuration seeded"
End Sub
