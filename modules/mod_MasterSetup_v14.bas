Attribute VB_Name = "mod_MasterSetup"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Master Setup - One-click system initialization
' ============================================================================

Option Explicit

' ============================================================================
' MASTER SETUP - Run this to build everything
' Creates all UserForms, sets up ACCUEIL, refreshes data
' ============================================================================
Public Sub MasterSetup()
    On Error GoTo ErrorHandler
    
    Dim startTime As Double: startTime = Timer
    Dim msg As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    msg = "Starting Master Setup..." & vbCrLf & vbCrLf
    
    ' === STEP 1: Build all UserForms ===
    msg = msg & "Building UserForms..." & vbCrLf
    
    ' 1. frmStockEntry
    DoEvents
    On Error Resume Next
    mod_BuildForm.BuildStockEntryForm
    If Err.Number = 0 Then
        msg = msg & "  [OK] frmStockEntry" & vbCrLf
    Else
        msg = msg & "  [FAIL] frmStockEntry: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' 2. frmArticleEditor
    DoEvents
    On Error Resume Next
    mod_BuildArticleEditor.BuildArticleEditorForm
    If Err.Number = 0 Then
        msg = msg & "  [OK] frmArticleEditor" & vbCrLf
    Else
        msg = msg & "  [FAIL] frmArticleEditor: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' 3. frmSupplierEditor
    DoEvents
    On Error Resume Next
    mod_BuildSupplierEditor.BuildSupplierEditorForm
    If Err.Number = 0 Then
        msg = msg & "  [OK] frmSupplierEditor" & vbCrLf
    Else
        msg = msg & "  [FAIL] frmSupplierEditor: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' 4. frmDashboard
    DoEvents
    On Error Resume Next
    mod_BuildDashboard.BuildDashboardForm
    If Err.Number = 0 Then
        msg = msg & "  [OK] frmDashboard" & vbCrLf
    Else
        msg = msg & "  [FAIL] frmDashboard: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' 5. frmSearch
    DoEvents
    On Error Resume Next
    mod_BuildSearch.BuildSearchForm
    If Err.Number = 0 Then
        msg = msg & "  [OK] frmSearch" & vbCrLf
    Else
        msg = msg & "  [FAIL] frmSearch: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' 6. frmReception
    DoEvents
    On Error Resume Next
    mod_BuildReception.BuildReceptionForm
    If Err.Number = 0 Then
        msg = msg & "  [OK] frmReception" & vbCrLf
    Else
        msg = msg & "  [FAIL] frmReception: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' 7. frmReports
    DoEvents
    On Error Resume Next
    mod_BuildReports.BuildReportsForm
    If Err.Number = 0 Then
        msg = msg & "  [OK] frmReports" & vbCrLf
    Else
        msg = msg & "  [FAIL] frmReports: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' 8. frmConfig
    DoEvents
    On Error Resume Next
    mod_BuildConfig.BuildConfigForm
    If Err.Number = 0 Then
        msg = msg & "  [OK] frmConfig" & vbCrLf
    Else
        msg = msg & "  [FAIL] frmConfig: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' 9. frmFirstRun
    DoEvents
    On Error Resume Next
    mod_BuildFirstRun.BuildFirstRunForm True
    If Err.Number = 0 Then
        msg = msg & "  [OK] frmFirstRun" & vbCrLf
    Else
        msg = msg & "  [FAIL] frmFirstRun: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' === STEP 2: Setup ACCUEIL buttons ===
    msg = msg & vbCrLf & "Setting up ACCUEIL..." & vbCrLf
    DoEvents
    On Error Resume Next
    mod_AccueilButtons.SetupAccueilButtons
    If Err.Number = 0 Then
        msg = msg & "  [OK] ACCUEIL buttons" & vbCrLf
    Else
        msg = msg & "  [FAIL] ACCUEIL: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' === STEP 3: Refresh data ===
    msg = msg & vbCrLf & "Refreshing data..." & vbCrLf
    DoEvents
    On Error Resume Next
    mod_Dashboard.RefreshDashboard
    If Err.Number = 0 Then
        msg = msg & "  [OK] Dashboard KPIs" & vbCrLf
    Else
        msg = msg & "  [FAIL] Dashboard: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' === STEP 4: Go to ACCUEIL ===
    DoEvents
    On Error Resume Next
    ThisWorkbook.Sheets("ACCUEIL").Activate
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    ' === SUMMARY ===
    Dim elapsed As Double: elapsed = Timer - startTime
    msg = msg & vbCrLf & String(40, "=") & vbCrLf
    msg = msg & "Setup complete in " & Format(elapsed, "0.0") & " seconds" & vbCrLf
    msg = msg & "All 8 UserForms built" & vbCrLf
    msg = msg & "ACCUEIL buttons ready" & vbCrLf
    msg = msg & "Dashboard KPIs refreshed" & vbCrLf
    
    msg = msg & vbCrLf
    If mod_FirstRun.IsFirstRun() Then
        msg = msg & "Configuration initiale requise - l'assistant va s'ouvrir." & vbCrLf
    End If

    MsgBox msg, vbInformation, mod_Config.SYS_TITLE

    ' Deliberately after the summary and after step 1: the wizard is a form, so
    ' it can only be shown once the forms have been built. Silent no-op when
    ' setup has already been completed.
    mod_FirstRun.FirstRunCheck
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    MsgBox "Setup error: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' QUICK REBUILD - Rebuild forms only (faster than full setup)
' ============================================================================
Public Sub QuickRebuildForms()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Dim msg As String: msg = "Rebuilding forms..." & vbCrLf
    
    DoEvents: mod_BuildForm.BuildStockEntryForm: msg = msg & "frmStockEntry OK" & vbCrLf
    DoEvents: mod_BuildArticleEditor.BuildArticleEditorForm: msg = msg & "frmArticleEditor OK" & vbCrLf
    DoEvents: mod_BuildSupplierEditor.BuildSupplierEditorForm: msg = msg & "frmSupplierEditor OK" & vbCrLf
    DoEvents: mod_BuildDashboard.BuildDashboardForm: msg = msg & "frmDashboard OK" & vbCrLf
    DoEvents: mod_BuildSearch.BuildSearchForm: msg = msg & "frmSearch OK" & vbCrLf
    DoEvents: mod_BuildReception.BuildReceptionForm: msg = msg & "frmReception OK" & vbCrLf
    DoEvents: mod_BuildReports.BuildReportsForm: msg = msg & "frmReports OK" & vbCrLf
    DoEvents: mod_BuildConfig.BuildConfigForm: msg = msg & "frmConfig OK" & vbCrLf
    DoEvents: mod_BuildFirstRun.BuildFirstRunForm True: msg = msg & "frmFirstRun OK" & vbCrLf
    
    Application.ScreenUpdating = True
    MsgBox msg, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur setup: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SYSTEM STATUS - Check what's built and what's missing
' ============================================================================
Public Sub SystemStatus()
    On Error GoTo ErrorHandler
    
    Dim msg As String
    msg = "=== DSS v14.0 System Status ===" & vbCrLf & vbCrLf
    
    ' Check sheets
    msg = msg & "SHEETS:" & vbCrLf
    Dim sheetNames As Variant
    sheetNames = Array("ACCUEIL", "ARTICLES", "MOUVEMENTS", "FOURNISSEURS", _
                       "CONFIG", "DASHBOARD", "BON_RECEPTION", "AUDIT_LOG", _
                       "FACTURES", "BARCODES", "BONS_COMMANDE")
    Dim i As Long
    For i = LBound(sheetNames) To UBound(sheetNames)
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetNames(i)))
        On Error GoTo ErrorHandler
        If Not ws Is Nothing Then
            Dim rowCount As Long: rowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            msg = msg & "  [OK] " & sheetNames(i) & " (" & rowCount & " rows)" & vbCrLf
        Else
            msg = msg & "  [MISSING] " & sheetNames(i) & vbCrLf
        End If
    Next i
    
    ' Check forms
    msg = msg & vbCrLf & "USERFORMS:" & vbCrLf
    Dim formNames As Variant
    formNames = Array("frmStockEntry", "frmArticleEditor", "frmSupplierEditor", _
                      "frmDashboard", "frmSearch", "frmReception", "frmReports", "frmConfig", _
                      "frmFirstRun")
    For i = LBound(formNames) To UBound(formNames)
        Dim vbComp As Object
        Set vbComp = Nothing
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(CStr(formNames(i)))
        On Error GoTo ErrorHandler
        If Not vbComp Is Nothing Then
            msg = msg & "  [OK] " & formNames(i) & vbCrLf
        Else
            msg = msg & "  [NOT BUILT] " & formNames(i) & " - Run MasterSetup" & vbCrLf
        End If
    Next i
    
    ' Check modules
    msg = msg & vbCrLf & "VBA MODULES:" & vbCrLf
    Dim modNames As Variant
    modNames = Array("mod_Config", "mod_DemoData", "mod_StockEngine", "mod_Dashboard", _
                     "mod_SupplierRegistry", "mod_Cleanup", "mod_Reports", "mod_Invoice", _
                     "mod_Barcode", "mod_FormHelpers", "mod_AccueilButtons", "mod_AccueilDesign", _
                     "mod_Backup", "mod_BuildForm", "mod_BuildArticleEditor", _
                     "mod_BuildSupplierEditor", "mod_BuildDashboard", "mod_BuildSearch", _
                     "mod_BuildReception", "mod_BuildReports", "mod_BuildConfig", _
                     "mod_MasterSetup", "mod_StockEntryHelpers", "mod_PurchaseOrder", _
                     "mod_FirstRun", "mod_BuildFirstRun")
    For i = LBound(modNames) To UBound(modNames)
        Dim mComp As Object
        Set mComp = Nothing
        On Error Resume Next
        Set mComp = ThisWorkbook.VBProject.VBComponents(CStr(modNames(i)))
        On Error GoTo ErrorHandler
        If Not mComp Is Nothing Then
            msg = msg & "  [OK] " & modNames(i) & vbCrLf
        Else
            msg = msg & "  [MISSING] " & modNames(i) & vbCrLf
        End If
    Next i
    
    ' Configuration state - the first thing to check when a store reports that
    ' it is looking at demo articles.
    msg = msg & vbCrLf & "CONFIGURATION:" & vbCrLf
    If mod_Config.IS_FIRST_RUN Then
        msg = msg & "  [PENDING] setup wizard not yet completed" & vbCrLf
    Else
        msg = msg & "  [OK] setup completed" & vbCrLf
    End If
    msg = msg & "  Etablissement : " & mod_Config.BUSINESS_NAME & vbCrLf
    msg = msg & "  Fenetre d'observation : " & mod_Config.ObservationDaysEffective() & " jour(s)" & vbCrLf

    MsgBox msg, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Error checking status: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' AUTO_OPEN - Runs when workbook opens (if macros enabled)
' ============================================================================
' Note on hooks: ThisWorkbook is a document module and cannot be imported as a
' .bas, so a repo that ships only modules cannot install Workbook_Open for the
' user. Auto_Open can be imported, which makes it the hook that works out of the
' box. docs/FIRST_RUN.md carries the Workbook_Open snippet for anyone who wants
' it as well - FirstRunCheck guards itself against running twice in a session,
' so having both hooks is harmless.
Public Sub Auto_Open()
    On Error GoTo ErrorHandler
    
    ' Before the KPI refresh: on a first open there is nothing worth reporting
    ' until the store has said who it is.
    mod_FirstRun.FirstRunCheck
    
    ' Only refresh ACCUEIL KPIs on open (don't rebuild forms)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ACCUEIL")
    If Not ws Is Nothing Then
        mod_AccueilButtons.RefreshAccueilKPIs
        ws.Activate
    End If
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    ' Silent fail on auto-open is acceptable
    On Error GoTo 0
End Sub
