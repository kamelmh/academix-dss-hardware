Attribute VB_Name = "mod_Reports"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Report Pack Module - ABC, Stock Aging, Supplier Scorecards, Demand Forecast
' P2 Priority (Hyperagent recommendation)
' ============================================================================

Option Explicit

' ============================================================================
' SUB: GenerateABCReport
' Creates ABC classification report sorted by consumption value
' ============================================================================
Public Sub GenerateABCReport()
    On Error GoTo ErrorHandler
    
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets("ARTICLES")
    
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No articles found.", vbExclamation, mod_Config.SYS_TITLE
        Exit Sub
    End If
    
    ' Collect data
    Dim data() As Variant
    ReDim data(1 To lastRow - 1, 1 To 6)
    
    Dim totalValue As Double: totalValue = 0
    Dim i As Long
    
    For i = 2 To lastRow
        Dim artCode As String: artCode = Trim(wsArt.Cells(i, 1).Value)
        If artCode <> "" Then
            Dim annualDemand As Double: annualDemand = mod_StockEngine.GetAnnualDemandFromHistory(artCode)
            Dim cmup As Double: cmup = Val(wsArt.Cells(i, 12).Value)
            If cmup <= 0 Then cmup = Val(wsArt.Cells(i, 8).Value)
            Dim consumptionValue As Double: consumptionValue = annualDemand * cmup
            
            data(i - 1, 1) = artCode
            data(i - 1, 2) = wsArt.Cells(i, 2).Value
            data(i - 1, 3) = annualDemand
            data(i - 1, 4) = cmup
            data(i - 1, 5) = consumptionValue
            data(i - 1, 6) = wsArt.Cells(i, 6).Value  ' ABC class
            
            totalValue = totalValue + consumptionValue
        End If
    Next i
    
    ' Sort by consumption value (descending)
    Dim j As Long, temp As Variant
    For i = 1 To lastRow - 2
        For j = i + 1 To lastRow - 1
            If data(i, 5) < data(j, 5) Then
                Dim k As Long
                For k = 1 To 6
                    temp = data(i, k)
                    data(i, k) = data(j, k)
                    data(j, k) = temp
                Next k
            End If
        Next j
    Next i
    
    ' Build report
    Dim report As String
    report = "ABC ANALYSIS REPORT" & vbCrLf
    report = report & "Generated: " & Format(Now, "DD/MM/YYYY HH:MM") & vbCrLf
    report = report & "Total Consumption Value: " & Format(totalValue, "#,##0") & " DZD" & vbCrLf
    report = report & String(60, "-") & vbCrLf
    report = report & "Code" & vbTab & "Designation" & vbTab & "Demand" & vbTab & "PU" & vbTab & "Value" & vbTab & "Class" & vbCrLf
    report = report & String(60, "-") & vbCrLf
    
    Dim cumValue As Double: cumValue = 0
    For i = 1 To lastRow - 1
        cumValue = cumValue + data(i, 5)
        Dim ratio As Double: ratio = cumValue / totalValue
        Dim abcClass As String
        
        If ratio <= 0.8 Then
            abcClass = "A"
        ElseIf ratio <= 0.95 Then
            abcClass = "B"
        Else
            abcClass = "C"
        End If
        
        report = report & data(i, 1) & vbTab & _
                 Left(data(i, 2), 20) & vbTab & _
                 Format(data(i, 3), "#,##0") & vbTab & _
                 Format(data(i, 4), "#,##0") & vbTab & _
                 Format(data(i, 5), "#,##0") & vbTab & _
                 abcClass & vbCrLf
    Next i
    
    report = report & String(60, "-") & vbCrLf
    report = report & "A: Top 80% | B: Next 15% | C: Last 5%" & vbCrLf
    
    ' Show report
    Dim wsReport As Worksheet
    Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsReport.Name = "REPORT_ABC_" & Format(Date, "YYYYMMDD")
    
    wsReport.Range("A1").Value = report
    wsReport.Range("A1").Font.Name = "Consolas"
    wsReport.Range("A1").Font.Size = 10
    wsReport.Columns("A").ColumnWidth = 100
    
    MsgBox "ABC Report created: " & wsReport.Name, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur rapport: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUB: GenerateStockAgingReport
' Shows stock age based on last movement date
' ============================================================================
Public Sub GenerateStockAgingReport()
    On Error GoTo ErrorHandler
    
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets("ARTICLES")
    
    Dim wsMouv As Worksheet
    Set wsMouv = ThisWorkbook.Sheets("MOUVEMENTS")
    
    Dim lastRowArt As Long: lastRowArt = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    
    ' Build report
    Dim report As String
    report = "STOCK AGING REPORT" & vbCrLf
    report = report & "Generated: " & Format(Now, "DD/MM/YYYY HH:MM") & vbCrLf
    report = report & String(70, "-") & vbCrLf
    report = report & "Code" & vbTab & "Designation" & vbTab & "Stock" & vbTab & "PU" & vbTab & "Value" & vbTab & "Last Move" & vbTab & "Days" & vbCrLf
    report = report & String(70, "-") & vbCrLf
    
    Dim totalValue As Double: totalValue = 0
    Dim i As Long
    
    For i = 2 To lastRowArt
        Dim artCode As String: artCode = Trim(wsArt.Cells(i, 1).Value)
        If artCode <> "" Then
            Dim stock As Double: stock = Val(wsArt.Cells(i, 3).Value)
            Dim cmup As Double: cmup = Val(wsArt.Cells(i, 12).Value)
            If cmup <= 0 Then cmup = Val(wsArt.Cells(i, 8).Value)
            Dim value As Double: value = stock * cmup
            
            ' Find last movement date
            Dim lastMoveDate As Date: lastMoveDate = DateSerial(2020, 1, 1)
            Dim foundRow As Variant
            foundRow = Application.Match(artCode, wsMouv.Range("B:B"), 0)
            
            If Not IsError(foundRow) Then
                Dim lastMouvRow As Long: lastMouvRow = wsMouv.Cells(wsMouv.Rows.Count, "B").End(xlUp).Row
                Dim r As Long
                For r = lastMouvRow To 2 Step -1
                    If wsMouv.Cells(r, 2).Value = artCode Then
                        lastMoveDate = wsMouv.Cells(r, 1).Value
                        Exit For
                    End If
                Next r
            End If
            
            Dim daysSinceMove As Long: daysSinceMove = Date - lastMoveDate
            
            ' Aging category
            Dim agingCategory As String
            If daysSinceMove <= 30 Then
                agingCategory = "0-30 days"
            ElseIf daysSinceMove <= 90 Then
                agingCategory = "31-90 days"
            ElseIf daysSinceMove <= 180 Then
                agingCategory = "91-180 days"
            ElseIf daysSinceMove <= 365 Then
                agingCategory = "181-365 days"
            Else
                agingCategory = "Over 1 year"
            End If
            
            report = report & artCode & vbTab & _
                     Left(wsArt.Cells(i, 2).Value, 18) & vbTab & _
                     Format(stock, "#,##0") & vbTab & _
                     Format(pu, "#,##0") & vbTab & _
                     Format(value, "#,##0") & vbTab & _
                     Format(lastMoveDate, "DD/MM/YYYY") & vbTab & _
                     daysSinceMove & " (" & agingCategory & ")" & vbCrLf
            
            totalValue = totalValue + value
        End If
    Next i
    
    report = report & String(70, "-") & vbCrLf
    report = report & "Total Stock Value: " & Format(totalValue, "#,##0") & " DZD" & vbCrLf
    
    ' Show report
    Dim wsReport As Worksheet
    Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsReport.Name = "REPORT_AGING_" & Format(Date, "YYYYMMDD")
    
    wsReport.Range("A1").Value = report
    wsReport.Range("A1").Font.Name = "Consolas"
    wsReport.Range("A1").Font.Size = 10
    wsReport.Columns("A").ColumnWidth = 120
    
    MsgBox "Stock Aging Report created: " & wsReport.Name, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur rapport: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUB: GenerateSupplierScorecard
' Shows supplier performance metrics
' ============================================================================
Public Sub GenerateSupplierScorecard()
    On Error GoTo ErrorHandler
    
    Dim wsFou As Worksheet
    Set wsFou = ThisWorkbook.Sheets("FOURNISSEURS")
    
    Dim wsMouv As Worksheet
    Set wsMouv = ThisWorkbook.Sheets("MOUVEMENTS")
    
    Dim lastRowFou As Long: lastRowFou = wsFou.Cells(wsFou.Rows.Count, "A").End(xlUp).Row
    
    ' Build report
    Dim report As String
    report = "SUPPLIER SCORECARD" & vbCrLf
    report = report & "Generated: " & Format(Now, "DD/MM/YYYY HH:MM") & vbCrLf
    report = report & String(80, "-") & vbCrLf
    report = report & "Supplier" & vbTab & "Orders" & vbTab & "Total Qty" & vbTab & "Total Value" & vbTab & "Avg Lead Time" & vbCrLf
    report = report & String(80, "-") & vbCrLf
    
    Dim i As Long
    For i = 2 To lastRowFou
        Dim fouName As String: fouName = Trim(wsFou.Cells(i, 2).Value)
        If fouName <> "" Then
            ' Count orders and totals
            Dim orderCount As Long: orderCount = 0
            Dim totalQty As Double: totalQty = 0
            Dim totalValue As Double: totalValue = 0
            
            Dim lastRowMouv As Long: lastRowMouv = wsMouv.Cells(wsMouv.Rows.Count, "A").End(xlUp).Row
            Dim r As Long
            For r = 2 To lastRowMouv
                If wsMouv.Cells(r, 9).Value = fouName And wsMouv.Cells(r, 4).Value = "ENTREE" Then
                    orderCount = orderCount + 1
                    totalQty = totalQty + Val(wsMouv.Cells(r, 5).Value)
                    totalValue = totalValue + Val(wsMouv.Cells(r, 6).Value)
                End If
            Next r
            
            report = report & Left(fouName, 20) & vbTab & _
                     orderCount & vbTab & _
                     Format(totalQty, "#,##0") & vbTab & _
                     Format(totalValue, "#,##0") & " DZD" & vbCrLf
        End If
    Next i
    
    report = report & String(80, "-") & vbCrLf
    
    ' Show report
    Dim wsReport As Worksheet
    Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsReport.Name = "REPORT_SUPPLIERS_" & Format(Date, "YYYYMMDD")
    
    wsReport.Range("A1").Value = report
    wsReport.Range("A1").Font.Name = "Consolas"
    wsReport.Range("A1").Font.Size = 10
    wsReport.Columns("A").ColumnWidth = 100
    
    MsgBox "Supplier Scorecard created: " & wsReport.Name, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur rapport: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUB: GenerateDemandForecast
' Simple moving average demand forecast
' ============================================================================
Public Sub GenerateDemandForecast()
    On Error GoTo ErrorHandler
    
    Dim wsArt As Worksheet
    Set wsArt = ThisWorkbook.Sheets("ARTICLES")
    
    Dim wsMouv As Worksheet
    Set wsMouv = ThisWorkbook.Sheets("MOUVEMENTS")
    
    Dim lastRowArt As Long: lastRowArt = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    
    ' Build report
    Dim report As String
    report = "DEMAND FORECAST (30-day moving average)" & vbCrLf
    report = report & "Generated: " & Format(Now, "DD/MM/YYYY HH:MM") & vbCrLf
    report = report & String(80, "-") & vbCrLf
    report = report & "Code" & vbTab & "Designation" & vbTab & "Avg Daily" & vbTab & "Forecast 30d" & vbTab & "Forecast 90d" & vbTab & "Suggested EOQ" & vbCrLf
    report = report & String(80, "-") & vbCrLf
    
    Dim totalForecast As Double: totalForecast = 0
    Dim i As Long
    
    For i = 2 To lastRowArt
        Dim artCode As String: artCode = Trim(wsArt.Cells(i, 1).Value)
        If artCode <> "" Then
            Dim annualDemand As Double: annualDemand = mod_StockEngine.GetAnnualDemandFromHistory(artCode)
            Dim avgDaily As Double: avgDaily = annualDemand / mod_Config.WORKING_DAYS_PER_YEAR
            
            ' Simple forecast: assume demand continues at same rate
            Dim forecast30 As Double: forecast30 = avgDaily * 30
            Dim forecast90 As Double: forecast90 = avgDaily * 90
            
            ' EOQ suggestion
            Dim pu As Double: pu = Val(wsArt.Cells(i, 8).Value)
            Dim eoq As Double: eoq = mod_StockEngine.ComputeEOQ(annualDemand, pu)
            
            report = report & artCode & vbTab & _
                     Left(wsArt.Cells(i, 2).Value, 18) & vbTab & _
                     Format(avgDaily, "#,##0.0") & vbTab & _
                     Format(forecast30, "#,##0") & vbTab & _
                     Format(forecast90, "#,##0") & vbTab & _
                     Format(eoq, "#,##0") & vbCrLf
            
            totalForecast = totalForecast + forecast30
        End If
    Next i
    
    report = report & String(80, "-") & vbCrLf
    report = report & "Total 30-day forecast: " & Format(totalForecast, "#,##0") & " units" & vbCrLf
    report = report & vbCrLf & "Note: Simple moving average. For seasonal patterns, extend historical data." & vbCrLf
    
    ' Show report
    Dim wsReport As Worksheet
    Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsReport.Name = "REPORT_FORECAST_" & Format(Date, "YYYYMMDD")
    
    wsReport.Range("A1").Value = report
    wsReport.Range("A1").Font.Name = "Consolas"
    wsReport.Range("A1").Font.Size = 10
    wsReport.Columns("A").ColumnWidth = 120
    
    MsgBox "Demand Forecast created: " & wsReport.Name, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur rapport: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub
