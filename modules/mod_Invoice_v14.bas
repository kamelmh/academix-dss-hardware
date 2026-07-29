Attribute VB_Name = "mod_Invoice"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Invoice Module - Simplified version
' ============================================================================

Option Explicit

' ============================================================================
' SUB: CreateInvoiceSheet
' Creates the FACTURES sheet
' ============================================================================
Public Sub CreateInvoiceSheet()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean: sheetExists = False
    
    ' Check if sheet exists
    Dim s As Worksheet
    For Each s In ThisWorkbook.Sheets
        If s.Name = "FACTURES" Then
            sheetExists = True
            Set ws = s
            Exit For
        End If
    Next s
    
    ' Create sheet if needed
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "FACTURES"
    End If
    
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Cells.Clear
    
    ' Headers
    ws.Cells(1, 1).Value = "Numero"
    ws.Cells(1, 2).Value = "Date"
    ws.Cells(1, 3).Value = "Client"
    ws.Cells(1, 4).Value = "Adresse"
    ws.Cells(1, 5).Value = "NIF"
    ws.Cells(1, 6).Value = "NIS"
    ws.Cells(1, 7).Value = "RC"
    ws.Cells(1, 8).Value = "Total HT"
    ws.Cells(1, 9).Value = "Total TVA"
    ws.Cells(1, 10).Value = "Total TTC"
    ws.Cells(1, 11).Value = "Reglement"
    ws.Cells(1, 12).Value = "Notes"
    ws.Cells(1, 13).Value = "User"
    ws.Cells(1, 14).Value = "Timestamp"
    
    ' Format
    ws.Range("A1:N1").Font.Bold = True
    ws.Range("A1:N1").Interior.Color = RGB(30, 60, 114)
    ws.Range("A1:N1").Font.Color = vbWhite
    ws.Columns("A:N").AutoFit
    
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    MsgBox "FACTURES sheet created.", vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur creation facture: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' FUNCTION: GenerateInvoiceNumber
' Returns: FAC-YYYYMMDD-NNN
' ============================================================================
Public Function GenerateInvoiceNumber() As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("FACTURES")
    If ws Is Nothing Then
        GenerateInvoiceNumber = "FAC-" & Format(Date, "YYYYMMDD") & "-001"
        Exit Function
    End If
    
    Dim today As String: today = Format(Date, "YYYYMMDD")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim seq As Long: seq = 1
    Dim i As Long
    
    For i = 2 To lastRow
        If Left(CStr(ws.Cells(i, 1).Value), 12) = "FAC-" & today Then
            seq = seq + 1
        End If
    Next i
    
    GenerateInvoiceNumber = "FAC-" & today & "-" & Format(seq, "000")
    On Error GoTo 0
End Function

' ============================================================================
' SUB: SaveInvoice
' Saves invoice data (pipe-delimited: Numero|Date|Client|Adresse|NIF|NIS|RC|HT|TVA|TTC|Reglement|Notes)
' ============================================================================
Public Sub SaveInvoice(ByVal invoiceData As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("FACTURES")
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim nr As Long: nr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    Dim parts() As String: parts = Split(invoiceData, "|")
    
    ' Save data
    ws.Cells(nr, 1).Value = parts(0)
    ws.Cells(nr, 2).Value = CDate(parts(1))
    ws.Cells(nr, 2).NumberFormat = "DD/MM/YYYY"
    ws.Cells(nr, 3).Value = parts(2)
    ws.Cells(nr, 4).Value = parts(3)
    ws.Cells(nr, 5).Value = parts(4)
    ws.Cells(nr, 6).Value = parts(5)
    ws.Cells(nr, 7).Value = parts(6)
    ws.Cells(nr, 8).Value = CDbl(parts(7))
    ws.Cells(nr, 9).Value = CDbl(parts(8))
    ws.Cells(nr, 10).Value = CDbl(parts(9))
    ws.Cells(nr, 11).Value = parts(10)
    ws.Cells(nr, 12).Value = parts(11)
    ws.Cells(nr, 13).Value = Environ("USERNAME")
    ws.Cells(nr, 14).Value = Now
    ws.Cells(nr, 14).NumberFormat = "DD/MM/YYYY HH:MM:SS"
    
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    MsgBox "Erreur sauvegarde facture: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' FUNCTION: GetInvoiceSummary
' Returns summary of invoices
' ============================================================================
Public Function GetInvoiceSummary() As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("FACTURES")
    If ws Is Nothing Then
        GetInvoiceSummary = "No invoices."
        Exit Function
    End If
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        GetInvoiceSummary = "No invoices."
        Exit Function
    End If
    
    Dim totalHT As Double, totalTVA As Double, totalTTC As Double
    Dim i As Long
    
    For i = 2 To lastRow
        totalHT = totalHT + Val(ws.Cells(i, 8).Value)
        totalTVA = totalTVA + Val(ws.Cells(i, 9).Value)
        totalTTC = totalTTC + Val(ws.Cells(i, 10).Value)
    Next i
    
    GetInvoiceSummary = "Invoices: " & (lastRow - 1) & vbCrLf & _
                        "Total HT: " & Format(totalHT, "#,##0") & " DZD" & vbCrLf & _
                        "Total TVA: " & Format(totalTVA, "#,##0") & " DZD" & vbCrLf & _
                        "Total TTC: " & Format(totalTTC, "#,##0") & " DZD"
    On Error GoTo 0
End Function
