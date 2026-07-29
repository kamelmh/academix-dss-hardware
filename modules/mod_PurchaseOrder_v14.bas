Attribute VB_Name = "mod_PurchaseOrder"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Purchase Order Module - Auto-generate POs from ROP alerts
' P2 Priority (Hyperagent recommendation)
' ============================================================================

Option Explicit

' ============================================================================
' CONSTANTS
' ============================================================================
Private Const PO_SHEET As String = "BONS_COMMANDE"
Private Const PO_PREFIX As String = "BC"

' ============================================================================
' TYPE: PurchaseOrderLine
' ============================================================================
Private Type PurchaseOrderLine
    CodeArticle As String
    Designation As String
    QuantiteCommander As Double
    PU As Double
    Fournisseur As String
    MontantTotal As Double
End Type

' ============================================================================
' TYPE: PurchaseOrder
' ============================================================================
Private Type PurchaseOrder
    Numero As String
    DateCommande As Date
    Fournisseur As String
    AdresseFournisseur As String
    Lines() As PurchaseOrderLine
    LineCount As Long
    TotalHT As Double
    TotalTVA As Double
    TotalTTC As Double
    Statut As String  ' En attente, Confirmee, Livree
End Type

' ============================================================================
' SUB: CreatePOSheet
' Creates BONS_COMMANDE sheet
' ============================================================================
Public Sub CreatePOSheet()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean: sheetExists = False
    
    Dim s As Worksheet
    For Each s In ThisWorkbook.Sheets
        If s.Name = PO_SHEET Then
            sheetExists = True
            Set ws = s
            Exit For
        End If
    Next s
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = PO_SHEET
    End If
    
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Cells.Clear
    
    ' Headers
    ws.Cells(1, 1).Value = "Numero BC"
    ws.Cells(1, 2).Value = "Date"
    ws.Cells(1, 3).Value = "Fournisseur"
    ws.Cells(1, 4).Value = "Adresse"
    ws.Cells(1, 5).Value = "NIF"
    ws.Cells(1, 6).Value = "NIS"
    ws.Cells(1, 7).Value = "RC"
    ws.Cells(1, 8).Value = "Total HT"
    ws.Cells(1, 9).Value = "Total TVA"
    ws.Cells(1, 10).Value = "Total TTC"
    ws.Cells(1, 11).Value = "Statut"
    ws.Cells(1, 12).Value = "User"
    ws.Cells(1, 13).Value = "Timestamp"
    
    ' Format
    ws.Range("A1:M1").Font.Bold = True
    ws.Range("A1:M1").Interior.Color = RGB(30, 60, 114)
    ws.Range("A1:M1").Font.Color = vbWhite
    ws.Columns("A:M").AutoFit
    
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    MsgBox "BONS_COMMANDE sheet created.", vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur commande: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUB: GeneratePOsFromROP
' Auto-generates purchase orders for items below ROP
' Groups by supplier for efficiency
' ============================================================================
Public Sub GeneratePOsFromROP()
    On Error GoTo ErrorHandler
    
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets("ARTICLES")
    
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    
    ' Collect items needing reorder
    Dim poCount As Long: poCount = 0
    Dim i As Long
    
    For i = 2 To lastRow
        Dim artCode As String: artCode = Trim(wsArt.Cells(i, 1).Value)
        If artCode <> "" Then
            Dim currentStock As Double: currentStock = Val(wsArt.Cells(i, 3).Value)
            Dim rop As Double: rop = mod_StockEngine.ComputeROP( _
                mod_StockEngine.GetAnnualDemandFromHistory(artCode) / mod_Config.WORKING_DAYS_PER_YEAR, _
                artCode)
            
            If currentStock <= rop Then
                poCount = poCount + 1
            End If
        End If
    Next i
    
    If poCount = 0 Then
        MsgBox "No items below ROP. No purchase orders needed.", vbInformation, mod_Config.SYS_TITLE
        Exit Sub
    End If
    
    ' Confirm
    If MsgBox(poCount & " items need reorder." & vbCrLf & _
              "Generate purchase orders by supplier?", _
              vbQuestion + vbYesNo, mod_Config.SYS_TITLE) = vbNo Then
        Exit Sub
    End If
    
    ' Generate POs grouped by supplier
    Dim suppliers As Object
    Set suppliers = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        artCode = Trim(wsArt.Cells(i, 1).Value)
        If artCode <> "" Then
            currentStock = Val(wsArt.Cells(i, 3).Value)
            Dim annualDemand As Double: annualDemand = mod_StockEngine.GetAnnualDemandFromHistory(artCode)
            Dim avgDaily As Double: avgDaily = annualDemand / mod_Config.WORKING_DAYS_PER_YEAR
            rop = mod_StockEngine.ComputeROP(avgDaily, artCode)
            
            If currentStock <= rop Then
                Dim fou As String: fou = Trim(wsArt.Cells(i, 9).Value)
                If fou = "" Then fou = "NON DECLARE"
                
                If Not suppliers.Exists(fou) Then
                    suppliers.Add fou, CreateObject("System.Collections.ArrayList")
                End If
                
                ' Add item to supplier list
                Dim item As Object
                Set item = CreateObject("Scripting.Dictionary")
                item.Add "Code", artCode
                item.Add "Designation", wsArt.Cells(i, 2).Value
                item.Add "CurrentStock", currentStock
                item.Add "ROP", rop
                item.Add "EOQ", mod_StockEngine.ComputeEOQ(annualDemand, Val(wsArt.Cells(i, 8).Value))
                item.Add "PU", Val(wsArt.Cells(i, 8).Value)
                
                suppliers(fou).Add item
            End If
        End If
    Next i
    
    ' Create PO for each supplier
    Dim wsPO As Worksheet
    Set wsPO = ThisWorkbook.Sheets(PO_SHEET)
    wsPO.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim fouKey As Variant
    For Each fouKey In suppliers.Keys
        Dim items As Object: Set items = suppliers(fouKey)
        
        ' Generate PO number
        Dim poNum As String: poNum = GeneratePONumber()
        
        ' Get supplier info
        Dim wsFou As Worksheet
        Set wsFou = ThisWorkbook.Sheets("FOURNISSEURS")
        Dim fouRow As Variant
        fouRow = Application.Match(fouKey, wsFou.Range("B:B"), 0)
        
        Dim fouAdresse As String, fouNIF As String, fouNIS As String, fouRC As String
        If Not IsError(fouRow) Then
            fouAdresse = wsFou.Cells(fouRow, 3).Value
            fouNIF = wsFou.Cells(fouRow, 5).Value
            fouNIS = wsFou.Cells(fouRow, 6).Value
            fouRC = wsFou.Cells(fouRow, 7).Value
        End If
        
        ' Calculate totals
        Dim totalHT As Double: totalHT = 0
        Dim j As Long
        For j = 0 To items.Count - 1
            Dim qty As Double: qty = items(j)("EOQ")
            Dim pu As Double: pu = items(j)("PU")
            totalHT = totalHT + (qty * pu)
        Next j
        
        Dim totalTVA As Double: totalTVA = totalHT * 0.19
        Dim totalTTC As Double: totalTTC = totalHT + totalTVA
        
        ' Save PO header
        Dim nr As Long: nr = wsPO.Cells(wsPO.Rows.Count, "A").End(xlUp).Row + 1
        wsPO.Cells(nr, 1).Value = poNum
        wsPO.Cells(nr, 2).Value = Date
        wsPO.Cells(nr, 2).NumberFormat = "DD/MM/YYYY"
        wsPO.Cells(nr, 3).Value = fouKey
        wsPO.Cells(nr, 4).Value = fouAdresse
        wsPO.Cells(nr, 5).Value = fouNIF
        wsPO.Cells(nr, 6).Value = fouNIS
        wsPO.Cells(nr, 7).Value = fouRC
        wsPO.Cells(nr, 8).Value = totalHT
        wsPO.Cells(nr, 9).Value = totalTVA
        wsPO.Cells(nr, 10).Value = totalTTC
        wsPO.Cells(nr, 11).Value = "En attente"
        wsPO.Cells(nr, 12).Value = Environ("USERNAME")
        wsPO.Cells(nr, 13).Value = Now
        wsPO.Cells(nr, 13).NumberFormat = "DD/MM/YYYY HH:MM:SS"
        
        poCount = poCount + 1
    Next fouKey
    
    wsPO.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    MsgBox poCount & " purchase orders generated." & vbCrLf & _
           "Check BONS_COMMANDE sheet.", _
           vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur commande: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' FUNCTION: GeneratePONumber
' Returns: BC-YYYYMMDD-NNN
' ============================================================================
Private Function GeneratePONumber() As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PO_SHEET)
    
    Dim today As String: today = Format(Date, "YYYYMMDD")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim seq As Long: seq = 1
    Dim i As Long
    
    For i = 2 To lastRow
        If Left(CStr(ws.Cells(i, 1).Value), 13) = PO_PREFIX & "-" & today Then
            seq = seq + 1
        End If
    Next i
    
    GeneratePONumber = PO_PREFIX & "-" & today & "-" & Format(seq, "000")
    On Error GoTo 0
End Function

' ============================================================================
' SUB: GetReorderAlert
' Shows items currently below ROP
' ============================================================================
Public Sub GetReorderAlert()
    On Error GoTo ErrorHandler
    
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets("ARTICLES")
    
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    
    Dim alertItems As String: alertItems = ""
    Dim alertCount As Long: alertCount = 0
    Dim i As Long
    
    For i = 2 To lastRow
        Dim artCode As String: artCode = Trim(wsArt.Cells(i, 1).Value)
        If artCode <> "" Then
            Dim currentStock As Double: currentStock = Val(wsArt.Cells(i, 3).Value)
            Dim annualDemand As Double: annualDemand = mod_StockEngine.GetAnnualDemandFromHistory(artCode)
            Dim avgDaily As Double: avgDaily = annualDemand / mod_Config.WORKING_DAYS_PER_YEAR
            Dim rop As Double: rop = mod_StockEngine.ComputeROP(avgDaily, artCode)
            Dim ss As Double: ss = mod_StockEngine.GetSafetyStock(artCode)
            
            If currentStock <= rop Then
                alertCount = alertCount + 1
                Dim status As String
                If currentStock <= ss Then
                    status = "CRITIQUE"
                Else
                    status = "ALERTE"
                End If
                
                alertItems = alertItems & artCode & " - " & wsArt.Cells(i, 2).Value & _
                             " (Stock: " & currentStock & ", ROP: " & Round(rop, 0) & _
                             ", " & status & ")" & vbCrLf
            End If
        End If
    Next i
    
    If alertCount = 0 Then
        MsgBox "All items above ROP. No reorder needed.", vbInformation, mod_Config.SYS_TITLE
    Else
        MsgBox alertCount & " items need reorder:" & vbCrLf & vbCrLf & alertItems, _
               vbExclamation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur commande: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub
