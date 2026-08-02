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

' ================================================================================
' FUNCTION: CalculateTurnoverRatio
' Formula: Turnover = COGS / Average Inventory
' BTS Ref: Semester 3 — Rotation des stocks
' ================================================================================
Public Function CalculateTurnoverRatio(ByVal sku As String) As Double
    On Error Resume Next
    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    If wsMouv Is Nothing Then CalculateTurnoverRatio = 0: Exit Function

    Dim currentYear As Integer: currentYear = Year(Date)
    Dim totalOutQty As Double

    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    totalOutQty = WorksheetFunction.SumIfs( _
        wsMouv.Columns(COL_MOUV_QTE), _
        wsMouv.Columns(COL_MOUV_CODE_ARTICLE), sku, _
        wsMouv.Columns(COL_MOUV_TYPE), "SORTIE", _
        wsMouv.Columns(COL_MOUV_DATE), ">=" & DateSerial(currentYear, 1, 1))
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    Dim cogs As Double: cogs = totalOutQty * GetCMUP(sku)
    Dim avgInventory As Double
    avgInventory = GetArticleStock(sku) * GetCMUP(sku)

    If avgInventory > 0 Then
        CalculateTurnoverRatio = cogs / avgInventory
    Else
        CalculateTurnoverRatio = 0
    End If
    On Error GoTo 0
End Function

' ================================================================================
' FUNCTION: CalculateDSI
' Formula: DSI = (Average Inventory / COGS) x 365
' BTS Ref: Semester 3 — Jours de stockage (Days Sales of Inventory)
' ================================================================================
Public Function CalculateDSI(ByVal sku As String) As Double
    Dim turnover As Double: turnover = CalculateTurnoverRatio(sku)
    If turnover > 0 Then
        CalculateDSI = 365 / turnover
    Else
        CalculateDSI = 999
    End If
End Function

' ================================================================================
' FUNCTION: CalculateStockCoverage
' Formula: Coverage = Current Stock / Average Daily Usage
' BTS Ref: Semester 3 — Couverture de stock
' ================================================================================
Public Function CalculateStockCoverage(ByVal sku As String) As Double
    Dim annualDemand As Double: annualDemand = GetAnnualDemandFromHistory(sku)
    If annualDemand <= 0 Then CalculateStockCoverage = 999: Exit Function

    Dim avgDaily As Double: avgDaily = annualDemand / mod_Config.WORKING_DAYS_PER_YEAR
    Dim currentStock As Double: currentStock = GetArticleStock(sku)

    If avgDaily > 0 Then
        CalculateStockCoverage = currentStock / avgDaily
    Else
        CalculateStockCoverage = 999
    End If
End Function

' ================================================================================
' FUNCTION: CalculatePriceVariance
' Formula: (Actual Price - Standard Price) x Actual Quantity
' BTS Ref: Semester 3 — Analyse des ecarts (Budget Management)
' ================================================================================
Public Function CalculatePriceVariance(ByVal sku As String, _
                                       ByVal stdPrice As Double) As Double
    Dim actualPrice As Double: actualPrice = GetUnitPrice(sku)
    Dim qtyPurchased As Double: qtyPurchased = GetTotalPurchasedQty(sku)
    CalculatePriceVariance = (actualPrice - stdPrice) * qtyPurchased
End Function

' ================================================================================
' FUNCTION: CalculateQuantityVariance
' Formula: (Actual Quantity - Standard Quantity) x Standard Price
' BTS Ref: Semester 3 — Analyse des ecarts (Budget Management)
' ================================================================================
Public Function CalculateQuantityVariance(ByVal sku As String, _
                                           ByVal stdQty As Double) As Double
    Dim actualQty As Double: actualQty = GetAnnualDemandFromHistory(sku)
    Dim unitPrice As Double: unitPrice = GetUnitPrice(sku)
    CalculateQuantityVariance = (actualQty - stdQty) * unitPrice
End Function

' ================================================================================
' HELPER: GetUnitPrice
' ================================================================================
Private Function GetUnitPrice(ByVal sku As String) As Double
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    If wsArt Is Nothing Then GetUnitPrice = 0: Exit Function

    Dim foundRow As Variant
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    If IsError(foundRow) Then
        GetUnitPrice = 0
    Else
        GetUnitPrice = Val(wsArt.Cells(foundRow, COL_ART_PU).Value)
    End If
End Function

' HELPER: GetCMUP
' Returns Weighted Average Cost (CMUP) for an article, falls back to PU if CMUP=0
' ================================================================================
Private Function GetCMUP(ByVal sku As String) As Double
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    If wsArt Is Nothing Then GetCMUP = 0: Exit Function

    Dim foundRow As Variant
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    If IsError(foundRow) Then
        GetCMUP = 0
    Else
        Dim cmup As Double: cmup = Val(wsArt.Cells(foundRow, COL_ART_CMUP).Value)
        If cmup <= 0 Then cmup = Val(wsArt.Cells(foundRow, COL_ART_PU).Value)
        GetCMUP = cmup
    End If
End Function

' ================================================================================
' HELPER: GetTotalPurchasedQty
' ================================================================================
Private Function GetTotalPurchasedQty(ByVal sku As String) As Double
    On Error Resume Next
    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    If wsMouv Is Nothing Then GetTotalPurchasedQty = 0: Exit Function

    Dim currentYear As Integer: currentYear = Year(Date)
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    GetTotalPurchasedQty = WorksheetFunction.SumIfs( _
        wsMouv.Columns(COL_MOUV_QTE), _
        wsMouv.Columns(COL_MOUV_CODE_ARTICLE), sku, _
        wsMouv.Columns(COL_MOUV_TYPE), "IN", _
        wsMouv.Columns(COL_MOUV_DATE), ">=" & DateSerial(currentYear, 1, 1))
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Function

' ================================================================================
' FUNCTION: ComputeDynamicSafetyStock
' Formula: SS = Z × σ_LT × √LT
' Z = service level factor (1.65 for 95%, 1.28 for 90%)
' σ_LT = standard deviation of demand during lead time
' LT = lead time in days
' BTS Ref: Semester 3 — Stock de securite dynamique
' ================================================================================
Public Function ComputeDynamicSafetyStock(ByVal sku As String, _
                                         Optional ByVal serviceLevel As Double = 0.95) As Double
    On Error Resume Next
    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    If wsMouv Is Nothing Then ComputeDynamicSafetyStock = GetSafetyStock(sku): Exit Function

    Dim lt As Integer: lt = GetLeadTime()
    If lt <= 0 Then lt = 2

    ' Z-score for service level (95% = 1.65, 90% = 1.28)
    Dim z As Double
    If serviceLevel >= 0.95 Then
        z = 1.65
    ElseIf serviceLevel >= 0.9 Then
        z = 1.28
    Else
        z = 1.0
    End If

    ' Calculate standard deviation of daily demand from recent history
    Dim currentYear As Integer: currentYear = Year(Date)
    Dim demandValues() As Double
    Dim count As Long: count = 0

    ' Count sortie movements for this article
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    Dim lastRow As Long: lastRow = wsMouv.Cells(wsMouv.Rows.Count, COL_MOUV_DATE).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If CStr(wsMouv.Cells(i, COL_MOUV_CODE_ARTICLE).Value) = sku Then
            If CStr(wsMouv.Cells(i, COL_MOUV_TYPE).Value) = "SORTIE" Then
                If Year(CDate(wsMouv.Cells(i, COL_MOUV_DATE).Value)) = currentYear Then
                    count = count + 1
                End If
            End If
        End If
    Next i

    ' Need at least 2 data points for standard deviation
    If count < 2 Then
        wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
        ComputeDynamicSafetyStock = GetSafetyStock(sku)
        Exit Function
    End If

    ' Collect demand values
    ReDim demandValues(1 To count)
    Dim idx As Long: idx = 0
    For i = 2 To lastRow
        If CStr(wsMouv.Cells(i, COL_MOUV_CODE_ARTICLE).Value) = sku Then
            If CStr(wsMouv.Cells(i, COL_MOUV_TYPE).Value) = "SORTIE" Then
                If Year(CDate(wsMouv.Cells(i, COL_MOUV_DATE).Value)) = currentYear Then
                    idx = idx + 1
                    demandValues(idx) = Val(wsMouv.Cells(i, COL_MOUV_QTE).Value)
                End If
            End If
        End If
    Next i
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    ' Calculate mean and standard deviation
    Dim mean As Double: mean = 0
    For i = 1 To count
        mean = mean + demandValues(i)
    Next i
    mean = mean / count

    Dim variance As Double: variance = 0
    For i = 1 To count
        variance = variance + (demandValues(i) - mean) ^ 2
    Next i
    variance = variance / (count - 1)
    Dim stdDev As Double: stdDev = Sqr(variance)

    ' SS = Z × σ × √LT
    ComputeDynamicSafetyStock = z * stdDev * Sqr(lt)
    On Error GoTo 0
End Function

' ================================================================================
' FUNCTION: CalculateStockVariation
' Formula: Variation = (Stock_final - Stock_initial) / Stock_initial × 100
' BTS Ref: Semester 3 — Variation de stock
' ================================================================================
Public Function CalculateStockVariation(ByVal sku As String, _
                                       Optional ByVal periodMonths As Integer = 1) As Double
    On Error Resume Next
    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    If wsMouv Is Nothing Then CalculateStockVariation = 0: Exit Function

    Dim cutoffDate As Date: cutoffDate = DateAdd("m", -periodMonths, Date)
    Dim initialQty As Double: initialQty = 0
    Dim finalQty As Double: finalQty = 0

    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    Dim lastRow As Long: lastRow = wsMouv.Cells(wsMouv.Rows.Count, COL_MOUV_DATE).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If CStr(wsMouv.Cells(i, COL_MOUV_CODE_ARTICLE).Value) = sku Then
            Dim mDate As Date: mDate = CDate(wsMouv.Cells(i, COL_MOUV_DATE).Value)
            Dim mType As String: mType = CStr(wsMouv.Cells(i, COL_MOUV_TYPE).Value)
            Dim mQty As Double: mQty = Val(wsMouv.Cells(i, COL_MOUV_QTE).Value)

            If mDate < cutoffDate Then
                ' Period start: IN adds, SORTIE subtracts
                If mType = "IN" Then initialQty = initialQty + mQty
                If mType = "SORTIE" Then initialQty = initialQty - mQty
            Else
                ' Period end: same logic
                If mType = "IN" Then finalQty = finalQty + mQty
                If mType = "SORTIE" Then finalQty = finalQty - mQty
            End If
        End If
    Next i
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    ' Variation = (final - initial) / initial × 100
    If initialQty > 0 Then
        CalculateStockVariation = ((finalQty - initialQty) / initialQty) * 100
    Else
        CalculateStockVariation = 0
    End If
    On Error GoTo 0
End Function

' ================================================================================
' FUNCTION: CalculateCommercialMargin
' Formula: Marge = (PV_HT - PA_HT) / PV_HT × 100
' PV_HT = selling price excluding TVA
' PA_HT = purchase price excluding TVA (CMUP)
' BTS Ref: Semester 3 — Marge commerciale
' ================================================================================
Public Function CalculateCommercialMargin(ByVal sku As String) As Double
    On Error Resume Next
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    If wsArt Is Nothing Then CalculateCommercialMargin = 0: Exit Function

    Dim foundRow As Variant
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    If IsError(foundRow) Then CalculateCommercialMargin = 0: Exit Function

    Dim pu As Double: pu = Val(wsArt.Cells(foundRow, COL_ART_PU).Value)
    Dim cmup As Double: cmup = Val(wsArt.Cells(foundRow, COL_ART_CMUP).Value)
    If cmup <= 0 Then cmup = pu

    ' PU is selling price (TTC or HT depending on config)
    ' CMUP is cost price
    If pu > 0 Then
        CalculateCommercialMargin = ((pu - cmup) / pu) * 100
    Else
        CalculateCommercialMargin = 0
    End If
    On Error GoTo 0
End Function

' ================================================================================
' FUNCTION: CalculateStockoutRate
' Formula: Taux_rupture = (Nb_jours_rupture / Nb_jours_total) × 100
' Counts days where stock was at or below safety stock
' BTS Ref: Semester 3 — Taux de rupture / service level
' ================================================================================
Public Function CalculateStockoutRate(ByVal sku As String, _
                                     Optional ByVal periodMonths As Integer = 3) As Double
    On Error Resume Next
    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    If wsMouv Is Nothing Then CalculateStockoutRate = 0: Exit Function

    Dim ss As Double: ss = GetSafetyStock(sku)
    Dim cutoffDate As Date: cutoffDate = DateAdd("m", -periodMonths, Date)
    Dim totalDays As Long: totalDays = 0
    Dim stockoutDays As Long: stockoutDays = 0
    Dim dailyStock As Double: dailyStock = 0

    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    Dim lastRow As Long: lastRow = wsMouv.Cells(wsMouv.Rows.Count, COL_MOUV_DATE).End(xlUp).Row
    Dim i As Long

    ' Build daily stock movement map
    Dim stockMap As Object: Set stockMap = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        If CStr(wsMouv.Cells(i, COL_MOUV_CODE_ARTICLE).Value) = sku Then
            Dim mDate As Date: mDate = CDate(wsMouv.Cells(i, COL_MOUV_DATE).Value)
            If mDate >= cutoffDate Then
                Dim mType As String: mType = CStr(wsMouv.Cells(i, COL_MOUV_TYPE).Value)
                Dim mQty As Double: mQty = Val(wsMouv.Cells(i, COL_MOUV_QTE).Value)
                Dim dateKey As String: dateKey = Format(mDate, "YYYY-MM-DD")

                If stockMap.Exists(dateKey) Then
                    If mType = "IN" Then
                        stockMap(dateKey) = stockMap(dateKey) + mQty
                    ElseIf mType = "SORTIE" Then
                        stockMap(dateKey) = stockMap(dateKey) - mQty
                    End If
                Else
                    If mType = "IN" Then
                        stockMap.Add dateKey, mQty
                    ElseIf mType = "SORTIE" Then
                        stockMap.Add dateKey, -mQty
                    End If
                End If
            End If
        End If
    Next i
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    ' Get current stock as starting point
    dailyStock = GetArticleStock(sku)

    ' Walk through dates in order
    Dim keys() As String: keys = stockMap.keys
    Dim sortedKeys() As String
    ' Simple sort (bubble sort for small datasets)
    ReDim sortedKeys(LBound(keys) To UBound(keys))
    For i = LBound(keys) To UBound(keys)
        sortedKeys(i) = keys(i)
    Next i
    Dim j As Long, temp As String
    For i = LBound(sortedKeys) To UBound(sortedKeys)
        For j = i + 1 To UBound(sortedKeys)
            If sortedKeys(i) > sortedKeys(j) Then
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i

    ' Count days
    For i = LBound(sortedKeys) To UBound(sortedKeys)
        totalDays = totalDays + 1
        dailyStock = dailyStock + stockMap(sortedKeys(i))
        If dailyStock <= ss Then stockoutDays = stockoutDays + 1
    Next i

    ' Calculate rate
    If totalDays > 0 Then
        CalculateStockoutRate = (stockoutDays / totalDays) * 100
    Else
        CalculateStockoutRate = 0
    End If
    On Error GoTo 0
End Function
