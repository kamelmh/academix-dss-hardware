Attribute VB_Name = "mod_Barcode"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Barcode Module - USB scanner + article lookup
' P2 Priority (Hyperagent recommendation)
' ============================================================================

Option Explicit

' ============================================================================
' CONSTANTS
' ============================================================================
Private Const BARCODE_SHEET As String = "BARCODES"

' ============================================================================
' SUB: CreateBarcodeSheet
' Creates BARCODES sheet for barcode-to-article mapping
' ============================================================================
Public Sub CreateBarcodeSheet()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean: sheetExists = False
    
    Dim s As Worksheet
    For Each s In ThisWorkbook.Sheets
        If s.Name = BARCODE_SHEET Then
            sheetExists = True
            Set ws = s
            Exit For
        End If
    Next s
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = BARCODE_SHEET
    End If
    
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Cells.Clear
    
    ' Headers
    ws.Cells(1, 1).Value = "Code Barre"
    ws.Cells(1, 2).Value = "Code Article"
    ws.Cells(1, 3).Value = "Designation"
    ws.Cells(1, 4).Value = "PU"
    ws.Cells(1, 5).Value = "Stock"
    ws.Cells(1, 6).Value = "Fournisseur"
    
    ' Format
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("A1:F1").Interior.Color = RGB(30, 60, 114)
    ws.Range("A1:F1").Font.Color = vbWhite
    ws.Columns("A:F").AutoFit
    
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    MsgBox "BARCODES sheet created.", vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur code-barres: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUB: GenerateBarcodesFromArticles
' Auto-generates barcodes from article codes (simple numeric mapping)
' ============================================================================
Public Sub GenerateBarcodesFromArticles()
    On Error GoTo ErrorHandler
    
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets("ARTICLES")
    
    Dim wsBar As Worksheet
    Set wsBar = ThisWorkbook.Sheets(BARCODE_SHEET)
    wsBar.Unprotect Password:=mod_Config.MASTER_PWD
    
    ' Clear existing barcodes
    Dim lastRowBar As Long: lastRowBar = wsBar.Cells(wsBar.Rows.Count, "A").End(xlUp).Row
    If lastRowBar > 1 Then wsBar.Range("A2:F" & lastRowBar).Clear
    
    ' Get articles
    Dim lastRowArt As Long: lastRowArt = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    
    Dim nr As Long: nr = 2
    Dim i As Long
    
    For i = 2 To lastRowArt
        Dim artCode As String: artCode = Trim(wsArt.Cells(i, 1).Value)
        Dim designation As String: designation = Trim(wsArt.Cells(i, 2).Value)
        Dim pu As Double: pu = Val(wsArt.Cells(i, 8).Value)
        Dim stock As Double: stock = Val(wsArt.Cells(i, 3).Value)
        Dim fournisseur As String: fournisseur = Trim(wsArt.Cells(i, 9).Value)
        
        If artCode <> "" Then
            ' Generate barcode: article code with check digit
            Dim barcode As String: barcode = GenerateCheckDigit(artCode)
            
            wsBar.Cells(nr, 1).Value = barcode
            wsBar.Cells(nr, 2).Value = artCode
            wsBar.Cells(nr, 3).Value = designation
            wsBar.Cells(nr, 4).Value = pu
            wsBar.Cells(nr, 5).Value = stock
            wsBar.Cells(nr, 6).Value = fournisseur
            
            nr = nr + 1
        End If
    Next i
    
    wsBar.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    MsgBox (nr - 2) & " barcodes generated from articles.", vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur code-barres: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' FUNCTION: GenerateCheckDigit
' Generates a simple check digit for barcode
' ============================================================================
Private Function GenerateCheckDigit(ByVal code As String) As String
    Dim i As Long, sum As Long
    
    ' Simple checksum
    For i = 1 To Len(code)
        sum = sum + Asc(Mid(code, i, 1))
    Next i
    
    ' Add prefix and check digit
    GenerateCheckDigit = "20" & code & (sum Mod 10)
End Function

' ============================================================================
' FUNCTION: LookupBarcode
' Returns article code for a scanned barcode
' ============================================================================
Public Function LookupBarcode(ByVal barcode As String) As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(BARCODE_SHEET)
    If ws Is Nothing Then
        LookupBarcode = ""
        Exit Function
    End If
    
    Dim foundRow As Variant
    foundRow = Application.Match(barcode, ws.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        LookupBarcode = ""
    Else
        LookupBarcode = CStr(ws.Cells(foundRow, 2).Value)
    End If
    
    On Error GoTo 0
End Function

' ============================================================================
' SUB: ScanBarcode
' Simulates barcode scan - looks up article and shows info
' ============================================================================
Public Sub ScanBarcode()
    On Error GoTo ErrorHandler
    
    Dim barcode As String
    barcode = InputBox("Scan or enter barcode:", "Barcode Scanner")
    
    If Len(Trim(barcode)) = 0 Then Exit Sub
    
    ' Look up barcode
    Dim artCode As String
    artCode = LookupBarcode(barcode)
    
    If artCode = "" Then
        MsgBox "Barcode not found: " & barcode, vbExclamation, mod_Config.SYS_TITLE
        Exit Sub
    End If
    
    ' Get article info
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets("ARTICLES")
    
    Dim foundRow As Variant
    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        MsgBox "Article not found: " & artCode, vbExclamation, mod_Config.SYS_TITLE
        Exit Sub
    End If
    
    ' Display article info
    Dim msg As String
    msg = "BARCODE SCANNED" & vbCrLf & vbCrLf & _
          "Barcode: " & barcode & vbCrLf & _
          "Code: " & artCode & vbCrLf & _
          "Article: " & wsArt.Cells(foundRow, 2).Value & vbCrLf & _
          "Stock: " & wsArt.Cells(foundRow, 3).Value & vbCrLf & _
          "PU: " & Format(Val(wsArt.Cells(foundRow, 8).Value), "#,##0.00") & " DZD" & vbCrLf & _
          "Fournisseur: " & wsArt.Cells(foundRow, 9).Value
    
    MsgBox msg, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur code-barres: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUB: QuickScanEntry
' Quick stock entry using barcode scan - most common workflow
' ============================================================================
Public Sub QuickScanEntry()
    On Error GoTo ErrorHandler
    
    Dim barcode As String
    barcode = InputBox("Scan article barcode:", "Quick Entry")
    
    If Len(Trim(barcode)) = 0 Then Exit Sub
    
    ' Look up
    Dim artCode As String: artCode = LookupBarcode(barcode)
    If artCode = "" Then
        MsgBox "Unknown barcode.", vbExclamation, mod_Config.SYS_TITLE
        Exit Sub
    End If
    
    ' Get article
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets("ARTICLES")
    Dim foundRow As Variant
    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
    If IsError(foundRow) Then Exit Sub
    
    Dim designation As String: designation = wsArt.Cells(foundRow, 2).Value
    Dim pu As Double: pu = Val(wsArt.Cells(foundRow, 8).Value)
    Dim currentStock As Double: currentStock = Val(wsArt.Cells(foundRow, 3).Value)
    
    ' Ask quantity
    Dim qtyStr As String
    qtyStr = InputBox("Article: " & designation & vbCrLf & _
                      "Current Stock: " & currentStock & vbCrLf & _
                      "PU: " & Format(pu, "#,##0.00") & " DZD" & vbCrLf & vbCrLf & _
                      "Enter quantity (negative for sortie):", _
                      "Quick Entry", "1")
    
    If Len(Trim(qtyStr)) = 0 Then Exit Sub
    
    Dim qty As Double: qty = Val(qtyStr)
    If qty = 0 Then Exit Sub
    
    ' Determine type
    Dim mvtType As String
    If qty > 0 Then
        mvtType = "ENTREE"
    Else
        mvtType = "SORTIE"
        qty = Abs(qty)
    End If
    
    ' Save movement
    Dim wsMouv As Worksheet
    Set wsMouv = ThisWorkbook.Sheets("MOUVEMENTS")
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim nr As Long: nr = wsMouv.Cells(wsMouv.Rows.Count, "A").End(xlUp).Row + 1
    
    wsMouv.Cells(nr, 1).Value = Date
    wsMouv.Cells(nr, 1).NumberFormat = "DD/MM/YYYY"
    wsMouv.Cells(nr, 2).Value = artCode
    wsMouv.Cells(nr, 3).Value = designation
    wsMouv.Cells(nr, 4).Value = mvtType
    wsMouv.Cells(nr, 5).Value = qty
    wsMouv.Cells(nr, 6).Value = qty * pu
    wsMouv.Cells(nr, 7).Value = "SCAN"
    wsMouv.Cells(nr, 8).Value = pu
    wsMouv.Cells(nr, 11).Value = Environ("USERNAME")
    wsMouv.Cells(nr, 12).Value = Now
    wsMouv.Cells(nr, 12).NumberFormat = "DD/MM/YYYY HH:MM:SS"
    
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    ' Update stock
    If mvtType = "ENTREE" Then
        mod_StockEngine.UpdateArticleStockBalance artCode, "IN", CLng(qty)
    Else
        mod_StockEngine.UpdateArticleStockBalance artCode, "OUT", CLng(qty)
    End If
    
    MsgBox mvtType & " saved!" & vbCrLf & _
           "Article: " & designation & vbCrLf & _
           "Quantity: " & qty & vbCrLf & _
           "New Stock: " & mod_StockEngine.GetArticleStock(artCode), _
           vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    MsgBox "Erreur scan: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
    On Error GoTo 0
End Sub
