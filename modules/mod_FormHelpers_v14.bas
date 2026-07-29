Attribute VB_Name = "mod_FormHelpers"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Form Helpers - Shared functions for UserForms
' ============================================================================

Option Explicit

' ============================================================================
' FUNCTION: ArticleExists
' Check if article code exists in ARTICLES sheet
' ============================================================================
Public Function ArticleExists(ByVal artCode As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If ws Is Nothing Then ArticleExists = False: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(UCase(artCode), ws.Range("A:A"), 0)
    ArticleExists = Not IsError(foundRow)
End Function

' ============================================================================
' FUNCTION: GetArticleDesignation
' Returns designation for given article code
' ============================================================================
Public Function GetArticleDesignation(ByVal artCode As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If ws Is Nothing Then GetArticleDesignation = "": Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(UCase(artCode), ws.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetArticleDesignation = ""
    Else
        GetArticleDesignation = ws.Cells(foundRow, mod_Config.COL_ART_DESIGNATION).Value
    End If
End Function

' ============================================================================
' FUNCTION: GetArticlePU
' Returns unit price for given article code
' ============================================================================
Public Function GetArticlePU(ByVal artCode As String) As Double
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If ws Is Nothing Then GetArticlePU = 0: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(UCase(artCode), ws.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetArticlePU = 0
    Else
        GetArticlePU = Val(ws.Cells(foundRow, mod_Config.COL_ART_PU).Value)
    End If
End Function

' ============================================================================
' FUNCTION: GetArticleCategory
' Returns category for given article code
' ============================================================================
Public Function GetArticleCategory(ByVal artCode As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If ws Is Nothing Then GetArticleCategory = "": Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(UCase(artCode), ws.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetArticleCategory = ""
    Else
        GetArticleCategory = ws.Cells(foundRow, mod_Config.COL_ART_CATEGORIE).Value
    End If
End Function

' ============================================================================
' FUNCTION: SupplierExists
' Check if supplier code exists in FOURNISSEURS sheet
' ============================================================================
Public Function SupplierExists(ByVal supCode As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_FOURNISSEURS)
    On Error GoTo 0
    
    If ws Is Nothing Then SupplierExists = False: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(UCase(supCode), ws.Range("A:A"), 0)
    SupplierExists = Not IsError(foundRow)
End Function

' ============================================================================
' FUNCTION: GetNextArticleCode
' Generates next available article code for given category
' ============================================================================
Public Function GetNextArticleCode(ByVal category As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If ws Is Nothing Then GetNextArticleCode = category & "-001": Exit Function
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, mod_Config.COL_ART_CODE).End(xlUp).Row
    
    Dim maxNum As Long
    maxNum = 0
    
    Dim i As Long
    For i = 2 To lastRow
        Dim code As String
        code = CStr(ws.Cells(i, mod_Config.COL_ART_CODE).Value)
        If Left(code, Len(category)) = category Then
            Dim numPart As String
            numPart = Mid(code, Len(category) + 2)
            If IsNumeric(numPart) Then
                If CLng(numPart) > maxNum Then maxNum = CLng(numPart)
            End If
        End If
    Next i
    
    GetNextArticleCode = category & "-" & Format(maxNum + 1, "000")
End Function

' ============================================================================
' FUNCTION: GetNextSupplierCode
' Generates next available supplier code
' ============================================================================
Public Function GetNextSupplierCode() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_FOURNISSEURS)
    On Error GoTo 0
    
    If ws Is Nothing Then GetNextSupplierCode = "SUP-001": Exit Function
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, mod_Config.COL_FOU_CODE).End(xlUp).Row
    
    Dim maxNum As Long
    maxNum = 0
    
    Dim i As Long
    For i = 2 To lastRow
        Dim code As String
        code = CStr(ws.Cells(i, mod_Config.COL_FOU_CODE).Value)
        If Left(code, 4) = "SUP-" Then
            Dim numPart As String
            numPart = Mid(code, 5)
            If IsNumeric(numPart) Then
                If CLng(numPart) > maxNum Then maxNum = CLng(numPart)
            End If
        End If
    Next i
    
    GetNextSupplierCode = "SUP-" & Format(maxNum + 1, "000")
End Function

' ============================================================================
' FUNCTION: FormatDZD
' Formats number as Algerian Dinar
' ============================================================================
Public Function FormatDZD(ByVal amount As Double) As String
    FormatDZD = Format(amount, "#,##0.00") & " DZD"
End Function

' ============================================================================
' FUNCTION: GetTodaysDate
' Returns today's date in DD/MM/YYYY format
' ============================================================================
Public Function GetTodaysDate() As String
    GetTodaysDate = Format(Date, "DD/MM/YYYY")
End Function
