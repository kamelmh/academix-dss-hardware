Attribute VB_Name = "mod_AuditTrail"
' ================================================================================
' mod_AuditTrail - Lightweight audit trail stub
' ================================================================================
Option Explicit

Private m_initialized As Boolean

Public Property Get AuditLogInitialized() As Boolean
    AuditLogInitialized = m_initialized
End Property

Public Sub InitAuditLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("AUDIT_LOG")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "AUDIT_LOG"
        ws.Cells(1, 1).Value = "Date"
        ws.Cells(1, 2).Value = "Action"
        ws.Cells(1, 3).Value = "Module"
        ws.Cells(1, 4).Value = "Sub"
        ws.Cells(1, 5).Value = "User"
        ws.Range("A1:E1").Font.Bold = True
    End If
    m_initialized = True
End Sub

Public Sub LogAction(ByVal actionType As String, ByVal description As String, _
                     ByVal moduleName As String, ByVal subName As String)
    If Not m_initialized Then InitAuditLog
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("AUDIT_LOG")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = actionType & ": " & description
    ws.Cells(r, 3).Value = moduleName
    ws.Cells(r, 4).Value = subName
    On Error Resume Next
    ws.Cells(r, 5).Value = Environ("USERNAME")
    On Error GoTo 0
End Sub
