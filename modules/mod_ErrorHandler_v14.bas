Attribute VB_Name = "mod_ErrorHandler"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Error Handler Module - Consistent error handling across all modules
' ============================================================================

Option Explicit

' ============================================================================
' TYPE: AppState
' Stores application state for save/restore
' ============================================================================
Private Type AppState
    ScreenUpdating As Boolean
    Calculation As XlCalculation
    EnableEvents As Boolean
    DisplayAlerts As Boolean
End Type

Private savedState As AppState
Private isProtectedSheet As Worksheet

' ============================================================================
' SUB: SaveAppState
' Saves current application state
' ============================================================================
Public Sub SaveAppState()
    With savedState
        .ScreenUpdating = Application.ScreenUpdating
        .Calculation = Application.Calculation
        .EnableEvents = Application.EnableEvents
        .DisplayAlerts = Application.DisplayAlerts
    End With
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

' ============================================================================
' SUB: RestoreAppState
' Restores saved application state
' ============================================================================
Public Sub RestoreAppState()
    With savedState
        Application.ScreenUpdating = .ScreenUpdating
        Application.Calculation = .Calculation
        Application.EnableEvents = .EnableEvents
        Application.DisplayAlerts = .DisplayAlerts
    End With
End Sub

' ============================================================================
' FUNCTION: UnprotectSheet
' Safely unprotects a sheet and returns it
' ============================================================================
Public Function UnprotectSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set UnprotectSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If UnprotectSheet Is Nothing Then Exit Function
    
    On Error Resume Next
    UnprotectSheet.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0
    
    Set isProtectedSheet = UnprotectSheet
End Function

' ============================================================================
' SUB: ProtectSheet
' Re-protects the last unprotected sheet
' ============================================================================
Public Sub ProtectSheet()
    If Not isProtectedSheet Is Nothing Then
        On Error Resume Next
        isProtectedSheet.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
        On Error GoTo 0
        Set isProtectedSheet = Nothing
    End If
End Sub

' ============================================================================
' FUNCTION: SafeOpenSheet
' Opens a sheet without protection, returns Nothing if not found
' ============================================================================
Public Function SafeOpenSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set SafeOpenSheet = ThisWorkbook.Sheets(sheetName)
    If Not SafeOpenSheet Is Nothing Then
        SafeOpenSheet.Unprotect Password:=mod_Config.MASTER_PWD
    End If
    On Error GoTo 0
End Function

' ============================================================================
' SUB: SafeCloseSheet
' Re-protects a sheet
' ============================================================================
Public Sub SafeCloseSheet(ByVal ws As Worksheet)
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

' ============================================================================
' FUNCTION: HandleError
' Displays user-friendly error message in French
' Returns: True if user clicked OK, False otherwise
' ============================================================================
Public Function HandleError(ByVal moduleName As String, ByVal errNum As Long, ByVal errDesc As String) As Boolean
    Dim msg As String
    
    ' Build user-friendly message in French
    msg = "Erreur dans " & moduleName & vbCrLf & vbCrLf
    
    Select Case errNum
        Case 9
            msg = msg & "Index hors limites. Verifiez les donnees."
        Case 13
            msg = msg & "Type de donnee incompatible. Verifiez les valeurs entrees."
        Case 91
            msg = msg & "Objet non defini. La feuille ou forme est introuvable."
        Case 1004
            msg = msg & "Operation impossible. Verifiez que le fichier n'est pas en lecture seule."
        Case Else
            msg = msg & "Erreur " & errNum & ": " & errDesc
    End Select
    
    msg = msg & vbCrLf & vbCrLf & "Veuillez reessayer ou contacter le support."
    
    MsgBox msg, vbExclamation, mod_Config.SYS_TITLE
    
    ' Log error
    LogError moduleName, errNum, errDesc
    
    HandleError = True
End Function

' ============================================================================
' SUB: LogError
' Logs error to AUDIT_LOG sheet
' ============================================================================
Private Sub LogError(ByVal moduleName As String, ByVal errNum As Long, ByVal errDesc As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_AUDIT_LOG)
    
    ' Create AUDIT_LOG if it doesn't exist
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = mod_Config.SHEET_AUDIT_LOG
        
        ws.Range("A1").Value = "Timestamp"
        ws.Range("B1").Value = "Action"
        ws.Range("C1").Value = "Details"
        ws.Range("D1").Value = "User"
        
        ws.Range("A1:D1").Font.Bold = True
        ws.Range("A1:D1").Interior.Color = RGB(30, 60, 114)
        ws.Range("A1:D1").Font.Color = vbWhite
    End If
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "DD/MM/YYYY HH:MM:SS"
    ws.Cells(nextRow, 2).Value = "ERROR_" & moduleName
    ws.Cells(nextRow, 3).Value = "Err " & errNum & ": " & errDesc
    ws.Cells(nextRow, 4).Value = Environ("USERNAME")
    
    ws.Columns("A:D").AutoFit
    
    On Error GoTo 0
End Sub

' ============================================================================
' MACRO: SafeMacroWrapper
' Wraps any macro with proper error handling and state management
' Usage: SafeMacroWrapper "MacroName", Address Of MyMacro
' ============================================================================
Public Sub SafeMacroWrapper(ByVal macroName As String, ByVal macroPtr As LongPtr)
    On Error GoTo ErrorHandler
    
    SaveAppState
    
    ' Call the macro
    Application.Run macroPtr
    
    RestoreAppState
    Exit Sub
    
ErrorHandler:
    RestoreAppState
    HandleError macroName, Err.Number, Err.Description
End Sub
