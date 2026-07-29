Attribute VB_Name = "mod_StockEngine"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Stock Engine - Reads from CONFIG sheet (no hardcoded values)
' ============================================================================

Option Explicit

' ============================================================================
' FUNCTION: GetOrderCost
' Reads order cost from CONFIG sheet (no hardcoded value)
' ============================================================================
Public Function GetOrderCost() As Double
    GetOrderCost = mod_Config.ORDER_COST
End Function

' ============================================================================
' FUNCTION: GetHoldingRate
' Reads holding rate from CONFIG sheet (no hardcoded value)
' ============================================================================
Public Function GetHoldingRate() As Double
    GetHoldingRate = mod_Config.HOLDING_RATE
End Function

' ============================================================================
' FUNCTION: GetLeadTime
' Reads lead time from CONFIG sheet (no hardcoded value)
' ============================================================================
Public Function GetLeadTime() As Integer
    GetLeadTime = mod_Config.LEAD_TIME_DEFAULT
End Function

' ============================================================================
' FUNCTION: GetSafetyStock
' Reads safety stock from ARTICLES sheet column 10
' ============================================================================
Public Function GetSafetyStock(ByVal sku As String) As Double
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then GetSafetyStock = 50: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetSafetyStock = 50
    Else
        Dim ss As Double: ss = Val(wsArt.Cells(foundRow, COL_ART_STOCK_SECURITE).Value)
        If ss > 0 Then GetSafetyStock = ss Else GetSafetyStock = 50
    End If
End Function

' ============================================================================
' FUNCTION: ComputeEOQ
' Formula: Q* = SQRT(2 x D x S / (P x t))
' All parameters read from CONFIG sheet
' ============================================================================
Public Function ComputeEOQ(ByVal AnnualDemand As Double, _
                            ByVal unitPrice As Double) As Double
    If unitPrice <= 0 Or AnnualDemand <= 0 Then
        ComputeEOQ = 0
        Exit Function
    End If

    Dim orderCost As Double: orderCost = GetOrderCost()
    Dim holdingRate As Double: holdingRate = GetHoldingRate()
    Dim holdingCostH As Double: holdingCostH = unitPrice * holdingRate

    If holdingCostH = 0 Then ComputeEOQ = 0: Exit Function
    ComputeEOQ = Sqr((2 * AnnualDemand * orderCost) / holdingCostH)
End Function

' ============================================================================
' FUNCTION: ComputeROP
' Formula: ROP = (avg_daily_demand x lead_time) + safety_stock
' Lead time read from CONFIG sheet
' ============================================================================
Public Function ComputeROP(ByVal AvgDailyDemand As Double, _
                            ByVal sku As String, _
                            Optional ByVal LeadTimeDays As Integer = -1) As Double
    If LeadTimeDays = -1 Then LeadTimeDays = GetLeadTime()
    ComputeROP = (AvgDailyDemand * LeadTimeDays) + GetSafetyStock(sku)
End Function

' ============================================================================
' SUB: ValidateStockLevel
' Fires a UI alert if current stock breaches ROP.
' ============================================================================
Public Sub ValidateStockLevel(ByVal sku As String, _
                               ByVal CurrentStock As Double, _
                               ByVal AnnualDemand As Double, _
                               ByVal unitPrice As Double)
    If AnnualDemand <= 0 Then Exit Sub

    Dim avgDaily As Double: avgDaily = AnnualDemand / mod_Config.WORKING_DAYS_PER_YEAR
    Dim rop As Double: rop = ComputeROP(avgDaily, sku)
    Dim ss As Double: ss = GetSafetyStock(sku)

    If CurrentStock <= rop Then
        Dim eoq As Double: eoq = ComputeEOQ(AnnualDemand, unitPrice)
        Dim alertLevel As String
        If CurrentStock <= ss Then
            alertLevel = "RUPTURE IMMINENTE"
        Else
            alertLevel = "SEUIL D'ALERTE ATTEINT"
        End If

        MsgBox alertLevel & vbCrLf & vbCrLf & _
               "Article  : " & sku & vbCrLf & _
               "Stock    : " & CurrentStock & " unites" & vbCrLf & _
               "ROP      : " & Round(rop, 1) & " unites" & vbCrLf & _
               "SS       : " & ss & " unites" & vbCrLf & _
               "EOQ (Q*) : " & Round(eoq, 0) & " unites a commander", _
               vbExclamation, mod_Config.SYS_TITLE
    End If
End Sub

' ============================================================================
' FUNCTION: GetArticleStock
' Returns current stock quantity for an article (reads from ARTICLES sheet)
' ============================================================================
Public Function GetArticleStock(ByVal sku As String) As Double
    Dim wsArt As Worksheet
    Dim foundRow As Variant
    
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then GetArticleStock = 0: Exit Function
    
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetArticleStock = 0
    Else
        GetArticleStock = Val(wsArt.Cells(foundRow, COL_ART_STOCK).Value)
    End If
End Function

' ============================================================================
' SUB: UpdateArticleStockBalance
' Updates stock quantity in ARTICLES sheet based on movements
' Error handling: Guaranteed sheet re-protection on any failure
' ============================================================================
Public Sub UpdateArticleStockBalance(ByVal artCode As String, ByVal mvtSign As String, ByVal qty As Long)
    On Error GoTo ErrorHandler
    
    Dim wsArt As Worksheet
    Dim foundRow As Variant
    
    ' Safely open sheet
    Set wsArt = mod_ErrorHandler.SafeOpenSheet(mod_Config.SHEET_ARTICLES)
    If wsArt Is Nothing Then Exit Sub
    
    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
    
    If Not IsError(foundRow) Then
        Dim currentQty As Double: currentQty = Val(wsArt.Cells(foundRow, COL_ART_STOCK).Value)
        
        If mvtSign = "IN" Then
            wsArt.Cells(foundRow, COL_ART_STOCK).Value = currentQty + qty
        Else
            wsArt.Cells(foundRow, COL_ART_STOCK).Value = currentQty - qty
        End If
    End If
    
    ' Always re-protect sheet
    mod_ErrorHandler.SafeCloseSheet wsArt
    Exit Sub
    
ErrorHandler:
    ' Guarantee sheet re-protection even on error
    mod_ErrorHandler.SafeCloseSheet wsArt
    mod_ErrorHandler.HandleError "UpdateArticleStockBalance", Err.Number, Err.Description
End Sub

' ============================================================================
' FUNCTION: GetAnnualDemandFromHistory
' Aggregates annual demand from MOUVEMENTS sheet for a given SKU
' ============================================================================
Public Function GetAnnualDemandFromHistory(ByVal sku As String) As Double
    On Error Resume Next
    Dim wsMouv As Worksheet
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    If wsMouv Is Nothing Then GetAnnualDemandFromHistory = 0: Exit Function
    
    Dim currentYear As Integer: currentYear = Year(Date)
    
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    GetAnnualDemandFromHistory = WorksheetFunction.SumIfs( _
        wsMouv.Range("E:E"), _
        wsMouv.Range("B:B"), sku, _
        wsMouv.Range("D:D"), "SORTIE", _
        wsMouv.Range("A:A"), ">=" & DateSerial(currentYear, 1, 1))
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Function

' ============================================================================
' FUNCTION: CalculateCMUP
' SCF-compliant moving weighted average (arrete 26/07/2008, points 123-6/123-7, IAS 2)
'
' Walks MOUVEMENTS chronologically:
'   ENTREE: qtyOnHand += qty; valueOnHand += qty * unitCostExclTVA
'   SORTIE: qtyOnHand -= qty; valueOnHand -= qty * currentCMUP (CMUP unchanged)
'
' Returns: valueOnHand / qtyOnHand (or 0 if no stock)
'
' Key fix: openingQty is derived from current balance minus net movements,
' NOT read directly from col 3 (which is the live balance).
' ============================================================================
Public Function CalculateCMUP(ByVal sku As String) As Double
    On Error GoTo ErrorHandler
    
    ' --- Step 1: Get opening stock from ARTICLES sheet ---
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    If wsArt Is Nothing Then CalculateCMUP = 0: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    If IsError(foundRow) Then CalculateCMUP = 0: Exit Function
    
    ' Read current live balance and PU (both excl TVA)
    Dim currentBalance As Double: currentBalance = Val(wsArt.Cells(foundRow, mod_Config.COL_ART_STOCK).Value)
    Dim rawUnitPrice As Double: rawUnitPrice = Val(wsArt.Cells(foundRow, mod_Config.COL_ART_PU).Value)
    
    ' Exclude TVA from PU if needed
    Dim puExclTVA As Double
    If mod_Config.PU_INCLUDES_TVA And mod_Config.TAX_RATE > 0 Then
        puExclTVA = rawUnitPrice / (1 + mod_Config.TAX_RATE)
    Else
        puExclTVA = rawUnitPrice
    End If
    
    ' --- Step 2: Get movements for this SKU ---
    Dim wsMouv As Worksheet
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    If wsMouv Is Nothing Then
        CalculateCMUP = IIf(currentBalance > 0, puExclTVA, 0)
        Exit Function
    End If
    
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim lastRow As Long: lastRow = wsMouv.Cells(wsMouv.Rows.Count, mod_Config.COL_MOUV_DATE).End(xlUp).Row
    Dim totalInQty As Double, totalOutQty As Double
    Dim i As Long
    
    ' First pass: compute net movements to derive initial stock
    ' initialStock = currentBalance - totalIn + totalOut
    For i = 2 To lastRow
        If Trim(wsMouv.Cells(i, mod_Config.COL_MOUV_CODE_ARTICLE).Value) = sku Then
            Dim movType As String: movType = UCase(Trim(wsMouv.Cells(i, mod_Config.COL_MOUV_TYPE).Value))
            Dim movQty As Double: movQty = Val(wsMouv.Cells(i, mod_Config.COL_MOUV_QTE).Value)
            If movType = "ENTREE" Then
                totalInQty = totalInQty + movQty
            ElseIf movType = "SORTIE" Then
                totalOutQty = totalOutQty + movQty
            End If
        End If
    Next i
    
    ' Derive initial stock (before any movements)
    Dim initialQty As Double: initialQty = currentBalance - totalInQty + totalOutQty
    If initialQty < 0 Then initialQty = 0
    
    ' Initialize running totals
    Dim qtyOnHand As Double: qtyOnHand = initialQty
    Dim valueOnHand As Double: valueOnHand = initialQty * puExclTVA
    Dim currentCMUP As Double
    If qtyOnHand > 0 Then currentCMUP = puExclTVA Else currentCMUP = 0
    
    ' --- Collect SKU movements into array for date-sort ---
    Dim mouvDate() As Double, mouvTypeArr() As String, mouvQtyArr() As Double
    Dim mouvPUArr() As Double, mouvValArr() As Double
    Dim mouvCount As Long: mouvCount = 0
    
    For i = 2 To lastRow
        If Trim(wsMouv.Cells(i, mod_Config.COL_MOUV_CODE_ARTICLE).Value) = sku Then
            mouvCount = mouvCount + 1
            ReDim Preserve mouvDate(1 To mouvCount)
            ReDim Preserve mouvTypeArr(1 To mouvCount)
            ReDim Preserve mouvQtyArr(1 To mouvCount)
            ReDim Preserve mouvPUArr(1 To mouvCount)
            ReDim Preserve mouvValArr(1 To mouvCount)
            mouvDate(mouvCount) = CDbl(wsMouv.Cells(i, mod_Config.COL_MOUV_DATE).Value)
            mouvTypeArr(mouvCount) = UCase(Trim(wsMouv.Cells(i, mod_Config.COL_MOUV_TYPE).Value))
            mouvQtyArr(mouvCount) = Val(wsMouv.Cells(i, mod_Config.COL_MOUV_QTE).Value)
            mouvPUArr(mouvCount) = Val(wsMouv.Cells(i, mod_Config.COL_MOUV_PU).Value)
            mouvValArr(mouvCount) = Val(wsMouv.Cells(i, mod_Config.COL_MOUV_VALEUR).Value)
        End If
    Next i
    
    ' Bubble-sort by date (ascending) for chronological order
    Dim j As Long, tmpD As Double, tmpT As String, tmpQ As Double, tmpP As Double, tmpV As Double
    If mouvCount > 1 Then
        For i = 1 To mouvCount - 1
            For j = i + 1 To mouvCount
                If mouvDate(j) < mouvDate(i) Then
                    tmpD = mouvDate(i): mouvDate(i) = mouvDate(j): mouvDate(j) = tmpD
                    tmpT = mouvTypeArr(i): mouvTypeArr(i) = mouvTypeArr(j): mouvTypeArr(j) = tmpT
                    tmpQ = mouvQtyArr(i): mouvQtyArr(i) = mouvQtyArr(j): mouvQtyArr(j) = tmpQ
                    tmpP = mouvPUArr(i): mouvPUArr(i) = mouvPUArr(j): mouvPUArr(j) = tmpP
                    tmpV = mouvValArr(i): mouvValArr(i) = mouvValArr(j): mouvValArr(j) = tmpV
                End If
            Next j
        Next i
    End If
    
    ' Second pass: walk movements in DATE order (chronological moving average)
    Dim mType As String, mQty As Double, movPU As Double, movValue As Double
    Dim movCostExclTVA As Double
    
    For i = 1 To mouvCount
        mType = mouvTypeArr(i)
        mQty = mouvQtyArr(i)
        
        If mType = "ENTREE" Then
            movPU = mouvPUArr(i)
            movValue = mouvValArr(i)
            
            ' If PU includes TVA, exclude it
            If mod_Config.PU_INCLUDES_TVA And mod_Config.TAX_RATE > 0 And movPU > 0 Then
                movCostExclTVA = movPU / (1 + mod_Config.TAX_RATE)
            ElseIf movPU > 0 Then
                movCostExclTVA = movPU
            ElseIf mQty > 0 Then
                movCostExclTVA = movValue / mQty
            Else
                movCostExclTVA = 0
            End If
            
            qtyOnHand = qtyOnHand + mQty
            valueOnHand = valueOnHand + (mQty * movCostExclTVA)
            
        ElseIf mType = "SORTIE" Then
            If qtyOnHand > 0 Then
                ' Subtract at current CMUP (cost doesn't change on sale)
                valueOnHand = valueOnHand - (mQty * currentCMUP)
                qtyOnHand = qtyOnHand - mQty
                If qtyOnHand < 0 Then qtyOnHand = 0
            End If
        End If
        
        ' Update CMUP after each movement
        If qtyOnHand > 0 Then
            currentCMUP = valueOnHand / qtyOnHand
        Else
            currentCMUP = 0
        End If
    Next i
    
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    CalculateCMUP = currentCMUP
    On Error GoTo 0
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    CalculateCMUP = 0
    On Error GoTo 0
End Function

' ============================================================================
' FUNCTION: GetUnitPriceExclTVA
' Returns unit price excluding TVA for costing purposes
' SCF requires TVA to be excluded from inventory valuation
' ============================================================================
Public Function GetUnitPriceExclTVA(ByVal sku As String) As Double
    On Error Resume Next
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    If wsArt Is Nothing Then GetUnitPriceExclTVA = 0: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    If IsError(foundRow) Then
        GetUnitPriceExclTVA = 0
    Else
        Dim pu As Double: pu = Val(wsArt.Cells(foundRow, mod_Config.COL_ART_PU).Value)
        ' Use CONFIG flag to determine if PU includes TVA
        Dim taxRate As Double: taxRate = mod_Config.TAX_RATE
        If mod_Config.PU_INCLUDES_TVA And taxRate > 0 Then
            GetUnitPriceExclTVA = pu / (1 + taxRate)
        Else
            GetUnitPriceExclTVA = pu
        End If
    End If
    On Error GoTo 0
End Function

' ============================================================================
' SUB: RefreshAllCMUP
' Recalculates CMUP for all articles in ARTICLES sheet
' Uses corrected SCF-compliant formula with opening stock
' Error handling: State save/restore, sheet re-protection
' ============================================================================
Public Sub RefreshAllCMUP()
    On Error GoTo ErrorHandler
    
    ' Save application state
    mod_ErrorHandler.SaveAppState
    
    Dim wsArt As Worksheet
    Set wsArt = mod_ErrorHandler.SafeOpenSheet(mod_Config.SHEET_ARTICLES)
    If wsArt Is Nothing Then GoTo Cleanup
    
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.Count, COL_ART_CODE).End(xlUp).Row
    
    Dim i As Long, cmup As Double
    For i = 2 To lastRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, COL_ART_CODE).Value)
        If sku <> "" Then
            cmup = CalculateCMUP(sku)
            If cmup > 0 Then wsArt.Cells(i, COL_ART_CMUP).Value = cmup
        End If
    Next i
    
Cleanup:
    ' Always restore state
    mod_ErrorHandler.SafeCloseSheet wsArt
    mod_ErrorHandler.RestoreAppState
    
    If Application.UserControl Then
        MsgBox "CMUP (Prix Moyen Pondere) mis a jour." & vbCrLf & vbCrLf & _
               "Formule conforme SCF:" & vbCrLf & _
               "Moyenne mobile chronologique sur les mouvements", _
               vbInformation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    ' Restore state on error
    mod_ErrorHandler.SafeCloseSheet wsArt
    mod_ErrorHandler.RestoreAppState
    mod_ErrorHandler.HandleError "RefreshAllCMUP", Err.Number, Err.Description
End Sub

' ============================================================================
' SUB: UpdateAllABCClassifications
' ABC classification: A: Top 80%, B: 15%, C: 5%
' Error handling: State save/restore, sheet re-protection
' ============================================================================
Public Sub UpdateAllABCClassifications(Optional ByVal silent As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' Save application state
    mod_ErrorHandler.SaveAppState
    
    Dim wsArt As Worksheet
    Set wsArt = mod_ErrorHandler.SafeOpenSheet(mod_Config.SHEET_ARTICLES)
    If wsArt Is Nothing Then GoTo Cleanup
    
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.Count, COL_ART_CODE).End(xlUp).Row
    If lastRow < 2 Then GoTo Cleanup

    Dim i As Long
    Dim totalValue As Double: totalValue = 0
    Dim articleValues() As Double: ReDim articleValues(2 To lastRow)
    Dim articleCodes() As String: ReDim articleCodes(2 To lastRow)

    For i = 2 To lastRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, COL_ART_CODE).Value)
        If sku <> "" Then
            Dim AnnualDemand As Double: AnnualDemand = GetAnnualDemandFromHistory(sku)
            Dim pu As Double: pu = Val(wsArt.Cells(i, COL_ART_PU).Value)
            articleValues(i) = AnnualDemand * pu
            articleCodes(i) = sku
            totalValue = totalValue + articleValues(i)
        End If
    Next i

    If totalValue = 0 Then GoTo Cleanup

    Dim j As Long, tempVal As Double, tempCode As String
    For i = 2 To lastRow - 1
        For j = i + 1 To lastRow
            If articleValues(i) < articleValues(j) Then
                tempVal = articleValues(i): articleValues(i) = articleValues(j): articleValues(j) = tempVal
                tempCode = articleCodes(i): articleCodes(i) = articleCodes(j): articleCodes(j) = tempCode
            End If
        Next j
    Next i

    Dim cumulativeValue As Double: cumulativeValue = 0
    For i = 2 To lastRow
        cumulativeValue = cumulativeValue + articleValues(i)
        Dim ratio As Double: ratio = cumulativeValue / totalValue
        Dim abcClass As String
        
        If ratio <= 0.8 Then
            abcClass = "A"
        ElseIf ratio <= 0.95 Then
            abcClass = "B"
        Else
            abcClass = "C"
        End If

        Dim foundRow As Variant
        foundRow = Application.Match(articleCodes(i), wsArt.Range("A:A"), 0)
        If Not IsError(foundRow) Then
            wsArt.Cells(foundRow, COL_ART_CLASSE_ABC).Value = abcClass
        End If
    Next i

Cleanup:
    ' Always restore state
    mod_ErrorHandler.SafeCloseSheet wsArt
    mod_ErrorHandler.RestoreAppState
    
    If Not silent Then
        MsgBox "Classifications ABC mises a jour.", vbInformation, mod_Config.SYS_TITLE
    End If
    Exit Sub
    
ErrorHandler:
    ' Restore state on error
    mod_ErrorHandler.SafeCloseSheet wsArt
    mod_ErrorHandler.RestoreAppState
    mod_ErrorHandler.HandleError "UpdateAllABCClassifications", Err.Number, Err.Description
End Sub

' ============================================================================
' FUNCTION: GetOrderRecommendation
' Returns recommended order quantity based on EOQ
' ============================================================================
Public Function GetOrderRecommendation(ByVal sku As String) As Double
    Dim annualDemand As Double: annualDemand = GetAnnualDemandFromHistory(sku)
    Dim pu As Double
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then GetOrderRecommendation = 0: Exit Function
    
    Dim foundRow As Variant
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    If IsError(foundRow) Then GetOrderRecommendation = 0: Exit Function
    
    pu = Val(wsArt.Cells(foundRow, COL_ART_PU).Value)
    GetOrderRecommendation = ComputeEOQ(annualDemand, pu)
End Function

' ============================================================================
' FUNCTION: GetStockStatus
' Returns stock status: OK, ALERT, CRITICAL, OUT_OF_STOCK
' ============================================================================
Public Function GetStockStatus(ByVal sku As String) As String
    Dim currentStock As Double: currentStock = GetArticleStock(sku)
    Dim annualDemand As Double: annualDemand = GetAnnualDemandFromHistory(sku)
    
    If annualDemand <= 0 Then GetStockStatus = "OK": Exit Function
    
    Dim avgDaily As Double: avgDaily = annualDemand / mod_Config.WORKING_DAYS_PER_YEAR
    Dim rop As Double: rop = ComputeROP(avgDaily, sku)
    Dim ss As Double: ss = GetSafetyStock(sku)
    
    If currentStock <= 0 Then
        GetStockStatus = "OUT_OF_STOCK"
    ElseIf currentStock <= ss Then
        GetStockStatus = "CRITICAL"
    ElseIf currentStock <= rop Then
        GetStockStatus = "ALERT"
    Else
        GetStockStatus = "OK"
    End If
End Function
