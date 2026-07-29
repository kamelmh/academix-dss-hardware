Attribute VB_Name = "MAIN_MACROS"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Main entry points - Clean version without old module references
' ============================================================================

Option Explicit

' ============================================================================
' HELPER: Check if a form exists
' ============================================================================
Private Function MainMacrosFormExists(ByVal formName As String) As Boolean
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(formName)
    MainMacrosFormExists = Not (vbComp Is Nothing)
    On Error GoTo 0
End Function

' ============================================================================
' STOCK ENTRY
' ============================================================================
Public Sub AjouterMouvement()
    On Error GoTo ErrorHandler
    Dim frmName As String
    frmName = "frmStockEntry"
    
    If MainMacrosFormExists(frmName) Then
        VBA.UserForms.Add(frmName).Show
    Else
        MsgBox "frmStockEntry not available.", vbExclamation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' NAVIGATION - Go to ACCUEIL
' ============================================================================
Public Sub GoToAccueil()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ACCUEIL)
    If Not ws Is Nothing Then
        ws.Activate
    Else
        MsgBox "Accueil sheet not found.", vbCritical, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' DASHBOARD
' ============================================================================
Public Sub ShowDashboard()
    On Error GoTo ErrorHandler
    Dim frmName As String
    frmName = "frmDashboard"
    
    If MainMacrosFormExists(frmName) Then
        VBA.UserForms.Add(frmName).Show
    Else
        MsgBox "Dashboard form not available.", vbExclamation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' REPORTS
' ============================================================================
Public Sub ShowReports()
    On Error GoTo ErrorHandler
    Dim frmName As String
    frmName = "frmReports"
    
    If MainMacrosFormExists(frmName) Then
        VBA.UserForms.Add(frmName).Show
    Else
        MsgBox "Reports form not available.", vbExclamation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SEARCH
' ============================================================================
Public Sub ShowSearch()
    On Error GoTo ErrorHandler
    Dim frmName As String
    frmName = "frmSearch"
    
    If MainMacrosFormExists(frmName) Then
        VBA.UserForms.Add(frmName).Show
    Else
        MsgBox "Search form not available.", vbExclamation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' CONFIGURATION
' ============================================================================
Public Sub ShowConfig()
    On Error GoTo ErrorHandler
    Dim frmName As String
    frmName = "frmConfig"
    
    If MainMacrosFormExists(frmName) Then
        VBA.UserForms.Add(frmName).Show
    Else
        MsgBox "Config form not available.", vbExclamation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' ARTICLE EDITOR
' ============================================================================
Public Sub ShowArticleEditor()
    On Error GoTo ErrorHandler
    Dim frmName As String
    frmName = "frmArticleEditor"
    
    If MainMacrosFormExists(frmName) Then
        VBA.UserForms.Add(frmName).Show
    Else
        MsgBox "Article Editor form not available.", vbExclamation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUPPLIER EDITOR
' ============================================================================
Public Sub ShowSupplierEditor()
    On Error GoTo ErrorHandler
    Dim frmName As String
    frmName = "frmSupplierEditor"
    
    If MainMacrosFormExists(frmName) Then
        VBA.UserForms.Add(frmName).Show
    Else
        MsgBox "Supplier Editor form not available.", vbExclamation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' RECEPTION
' ============================================================================
Public Sub ShowReception()
    On Error GoTo ErrorHandler
    Dim frmName As String
    frmName = "frmReception"
    
    If MainMacrosFormExists(frmName) Then
        VBA.UserForms.Add(frmName).Show
    Else
        MsgBox "Reception form not available.", vbExclamation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' REFRESH ALL DATA
' ============================================================================
Public Sub RefreshAllData()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Refresh CMUP
    mod_StockEngine.RefreshAllCMUP
    
    ' Refresh ABC classifications
    mod_StockEngine.UpdateAllABCClassifications silent:=True
    
    Application.ScreenUpdating = True
    
    MsgBox "All data refreshed successfully.", vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' GENERATE DEMO DATA
' ============================================================================
Public Sub RunGenerateDemoData()
    On Error GoTo ErrorHandler
    
    If MsgBox("Generate demo data? This will reset all articles and movements.", _
              vbQuestion + vbYesNo, mod_Config.SYS_TITLE) = vbNo Then
        Exit Sub
    End If
    
    mod_DemoData.GenerateDemoData
    
    MsgBox "Demo data generated successfully." & vbCrLf & _
           "40 articles, 9 suppliers, 90 days of movements.", _
           vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' CLEANUP WORKBOOK
' ============================================================================
Public Sub RunCleanupWorkbook()
    On Error GoTo ErrorHandler
    
    If MsgBox("This will delete all data sheets and recreate them." & vbCrLf & _
              "Are you sure?", vbQuestion + vbYesNo + vbDefaultButton2, _
              mod_Config.SYS_TITLE) = vbNo Then
        Exit Sub
    End If
    
    mod_Cleanup.CleanupWorkbook
    
    MsgBox "Workbook cleaned and reset.", vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' AUTOFIT ALL SHEETS - Adjust columns and rows for readability
' ============================================================================
Public Sub AutoFitAllSheets()
    On Error Resume Next
    Dim ws As Worksheet
    Dim originalSheet As String
    originalSheet = ActiveSheet.Name
    
    For Each ws In ThisWorkbook.Sheets
        ws.Activate
        ws.Unprotect Password:=mod_Config.MASTER_PWD
        ws.Cells.EntireColumn.AutoFit
        ws.Cells.EntireRow.AutoFit
        ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    Next ws
    
    ThisWorkbook.Sheets(originalSheet).Activate
    On Error GoTo 0
    MsgBox "All sheets autofitted.", vbInformation, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' BACKUP NOW - Create timestamped backup
' ============================================================================
Public Sub BackupNow()
    On Error GoTo ErrorHandler
    mod_Backup.ManualBackup
    Exit Sub
    
ErrorHandler:
    MsgBox "Backup error: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' RESTORE FROM BACKUP - Show backup list and restore
' ============================================================================
Public Sub RestoreFromBackup()
    On Error GoTo ErrorHandler
    mod_Backup.RestoreBackup
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur restauration: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' EXPORT TO CSV - Export all sheets to CSV
' ============================================================================
Public Sub ExportAllToCSV()
    On Error GoTo ErrorHandler
    mod_Backup.ExportToCSV
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur export: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' REDESIGN ACCUEIL - Modern dashboard with gradient + KPI cards
' ============================================================================
Public Sub RedesignAccueil()
    On Error GoTo ErrorHandler
    mod_AccueilDesign.RedesignAccueil
    Exit Sub
    
ErrorHandler:
    MsgBox "Redesign error: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' RESET ACCUEIL - Reset to default layout
' ============================================================================
Public Sub ResetAccueil()
    On Error GoTo ErrorHandler
    mod_AccueilDesign.ResetAccueil
    Exit Sub
    
ErrorHandler:
    MsgBox "Reset error: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub
