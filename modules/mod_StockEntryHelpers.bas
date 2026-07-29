Attribute VB_Name = "mod_StockEntryHelpers"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' mod_StockEntryHelpers - Helper functions for frmStockEntry (MOUVEMENTS
' CRUD + ARTICLES lookup). Stock-balance math and dashboard refresh are
' delegated to the existing mod_StockEngine / mod_Dashboard modules -
' never duplicated here.
' ============================================================================

Option Explicit

' ============================================================================
' FUNCTION: FindArticleRow
' Returns the ARTICLES row for an exact code match, or 0 if not found.
' ============================================================================
Public Function FindArticleRow(ByVal artCode As String) As Long
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    If wsArt Is Nothing Then FindArticleRow = 0: Exit Function

    Dim foundRow As Variant
    foundRow = Application.Match(Trim(artCode), wsArt.Range("A:A"), 0)
    If IsError(foundRow) Then
        FindArticleRow = 0
    Else
        FindArticleRow = CLng(foundRow)
    End If
End Function
