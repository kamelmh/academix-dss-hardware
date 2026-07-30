Attribute VB_Name = "mod_Dashboard"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Hardware Store Dashboard - 40 articles, real-time KPIs
' ============================================================================

Option Explicit

'--------------------------------------------------------------------------------------
' MAIN: RefreshDashboard
'--------------------------------------------------------------------------------------
Public Sub RefreshDashboard()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetOrCreateDashboardSheet()
    
    Call UpdateKPIs(ws)
    Call UpdateCriticalTable(ws)
    Call UpdateABCXYZSummary(ws)
    Call UpdateProjection(ws)
    
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Range("B1").Value = "Derniere actualisation : " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    ws.Range("B1").Font.Size = 8
    ws.Range("B1").Font.Italic = True
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    MsgBox "Tableau de bord actualise avec succes !", vbInformation, "DSS Quincaillerie"
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    MsgBox "Erreur actualisation: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------------
' KPIs
'--------------------------------------------------------------------------------------
Private Sub UpdateKPIs(ws As Worksheet)
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    If wsArt Is Nothing Then Exit Sub
    
    Dim totalSKUs As Long
    Dim countRupture As Long
    Dim countAlert As Long
    Dim totalValue As Double
    
    totalSKUs = wsArt.Cells(wsArt.Rows.Count, COL_ART_CODE).End(xlUp).Row - 1
    If totalSKUs < 0 Then totalSKUs = 0
    
    Dim i As Long
    On Error Resume Next
    For i = 2 To totalSKUs + 1
        Dim stock As Double: stock = Val(wsArt.Cells(i, COL_ART_STOCK).Value)
        Dim pu As Double: pu = Val(wsArt.Cells(i, COL_ART_PU).Value)
        
        totalValue = totalValue + (stock * pu)
        
        Dim sku As String: sku = Trim(wsArt.Cells(i, COL_ART_CODE).Value)
        Dim ss As Double: ss = mod_StockEngine.GetSafetyStock(sku)
        Dim annualDemand As Double: annualDemand = mod_StockEngine.GetAnnualDemandFromHistory(sku)
        Dim rop As Double: rop = mod_StockEngine.ComputeROP(annualDemand / mod_Config.WORKING_DAYS_PER_YEAR, sku)
        
        If stock <= 0 Then
            countRupture = countRupture + 1
        ElseIf stock <= rop Then
            countAlert = countAlert + 1
        End If
    Next i
    On Error GoTo 0
    
    ws.Range("B2").Value = "Total Articles"
    ws.Range("C2").Value = totalSKUs
    
    ws.Range("B3").Value = "Articles en RUPTURE"
    ws.Range("C3").Value = countRupture
    ws.Range("C3").Font.Color = RGB(200, 0, 0)
    
    ws.Range("B4").Value = "Articles en ALERTE"
    ws.Range("C4").Value = countAlert
    ws.Range("C4").Font.Color = RGB(200, 100, 0)
    
    ws.Range("B5").Value = "Valeur Total Stock"
    ws.Range("C5").Value = totalValue
    On Error Resume Next
    ws.Range("C5").NumberFormat = "#,##0.00"
    On Error GoTo 0
    
    ws.Range("B6").Value = "Rotation Moyenne (ITR)"
    ws.Range("C6").Value = "Calcul Local"
    
    With ws.Range("B2:B6")
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("C2:C5")
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'--------------------------------------------------------------------------------------
' CRITICAL TABLE
'--------------------------------------------------------------------------------------
Private Sub UpdateCriticalTable(ws As Worksheet)
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Range("D2:G2").Value = Array("SKU", "Designation", "Stock", "Etat")
    ws.Range("D2:G2").Interior.Color = RGB(0, 70, 127)
    ws.Range("D2:G2").Font.Color = RGB(255, 255, 255)
    ws.Range("D2:G2").Font.Bold = True
    ws.Range("D2:G2").HorizontalAlignment = xlCenter
    
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    If wsArt Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.Count, COL_ART_CODE).End(xlUp).Row
    
    Dim criticalList(1 To 1000, 1 To 4) As Variant
    Dim countCrit As Integer: countCrit = 0
    
    Dim i As Long
    For i = 2 To lastRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, COL_ART_CODE).Value)
        Dim stock As Double: stock = Val(wsArt.Cells(i, COL_ART_STOCK).Value)
        
        Dim annualDemand As Double: annualDemand = mod_StockEngine.GetAnnualDemandFromHistory(sku)
        Dim rop As Double: rop = mod_StockEngine.ComputeROP(annualDemand / mod_Config.WORKING_DAYS_PER_YEAR, sku)
        
        If stock <= rop Then
            countCrit = countCrit + 1
            If countCrit > 1000 Then Exit For
            
            criticalList(countCrit, 1) = sku
            criticalList(countCrit, 2) = wsArt.Cells(i, COL_ART_DESIGNATION).Value
            criticalList(countCrit, 3) = stock
            criticalList(countCrit, 4) = IIf(stock <= 0, "RUPTURE", "ALERTE")
        End If
    Next i
    
    Dim rowNum As Integer: rowNum = 3
    For i = 1 To countCrit
        If i > 10 Then Exit For  ' Show top 10 instead of 5
        ws.Cells(rowNum, 4).Value = criticalList(i, 1)
        ws.Cells(rowNum, 5).Value = criticalList(i, 2)
        ws.Cells(rowNum, 6).Value = criticalList(i, 3)
        ws.Cells(rowNum, 7).Value = criticalList(i, 4)
        
        If criticalList(i, 3) <= 0 Then
            ws.Range("D" & rowNum & ":G" & rowNum).Interior.Color = RGB(255, 200, 200)
        End If
        rowNum = rowNum + 1
    Next i
    
    If rowNum > 3 Then
        ws.Range("D3:G" & rowNum - 1).Borders.LineStyle = xlContinuous
    End If
    ws.Columns("D:G").AutoFit
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'--------------------------------------------------------------------------------------
' ABC SUMMARY
'--------------------------------------------------------------------------------------
Private Sub UpdateABCXYZSummary(ws As Worksheet)
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Range("I2").Value = "Classe"
    ws.Range("J2").Value = "Nombre d'articles"
    ws.Range("I2:J2").Interior.Color = RGB(0, 70, 127)
    ws.Range("I2:J2").Font.Color = RGB(255, 255, 255)
    ws.Range("I2:J2").Font.Bold = True
    
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    If wsArt Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.Count, COL_ART_CODE).End(xlUp).Row
    
    Dim classes As Variant: classes = Array("A", "B", "C")
    Dim rowNum As Integer: rowNum = 3
    
    Dim c As Integer
    For c = 0 To UBound(classes)
        Dim cls As String: cls = classes(c)
        Dim countCls As Long: countCls = 0
        
        Dim i As Long
        For i = 2 To lastRow
            If wsArt.Cells(i, COL_ART_CLASSE_ABC).Value = cls Then
                countCls = countCls + 1
            End If
        Next i
        
        ws.Cells(rowNum, 9).Value = cls
        ws.Cells(rowNum, 10).Value = countCls
        rowNum = rowNum + 1
    Next c
    
    ws.Range("I3:J" & rowNum - 1).Borders.LineStyle = xlContinuous
    ws.Columns("I:J").AutoFit
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'--------------------------------------------------------------------------------------
' PROJECTION
'--------------------------------------------------------------------------------------
Private Sub UpdateProjection(ws As Worksheet)
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    Dim wsMouv As Worksheet
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    If wsMouv Is Nothing Or wsArt Is Nothing Then Exit Sub
    
    Dim lastMouv As Long: lastMouv = wsMouv.Cells(wsMouv.Rows.Count, COL_MOUV_DATE).End(xlUp).Row
    Dim lastArt As Long: lastArt = wsArt.Cells(wsArt.Rows.Count, COL_ART_CODE).End(xlUp).Row
    Dim i As Long
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastMouv
        If Trim(wsMouv.Cells(i, COL_MOUV_TYPE).Value) = "SORTIE" Then
            Dim artCode As String: artCode = Trim(wsMouv.Cells(i, COL_MOUV_CODE_ARTICLE).Value)
            Dim qty As Double: qty = Val(wsMouv.Cells(i, COL_MOUV_QTE).Value)
            If dict.Exists(artCode) Then
                dict(artCode) = CDbl(dict(artCode)) + qty
            Else
                dict.Add artCode, qty
            End If
        End If
    Next i
    
    Dim startCol As Long: startCol = 12
    ws.Cells(1, startCol).Value = "Projection des Ruptures"
    ws.Range(ws.Cells(1, startCol), ws.Cells(1, startCol + 3)).Merge
    ws.Cells(1, startCol).Font.Bold = True
    ws.Cells(1, startCol).Font.Size = 10
    
    ws.Range(ws.Cells(2, startCol), ws.Cells(2, startCol + 3)).Value = _
        Array("Article", "Stock", "Cons./Jour", "Jours Rest.")
    With ws.Range(ws.Cells(2, startCol), ws.Cells(2, startCol + 3))
        .Interior.Color = RGB(0, 70, 127)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Font.Size = 8
    End With
    
    ' Observation window, resolved once. ObservationDaysEffective walks
    ' MOUVEMENTS, so calling it inside the per-article loop below would rescan
    ' the whole sheet for every article.
    Dim obsDays As Long: obsDays = mod_Config.ObservationDaysEffective()

    Dim rowNum As Long: rowNum = 3
    For i = 2 To lastArt
        artCode = Trim(wsArt.Cells(i, COL_ART_CODE).Value)
        If artCode <> "" Then
            Dim stock As Double: stock = Val(wsArt.Cells(i, COL_ART_STOCK).Value)
            Dim totalOut As Double
            If dict.Exists(artCode) Then totalOut = CDbl(dict(artCode)) Else totalOut = 0
            
            ' Was: totalOut / OBSERVATION_DAYS, a fixed 90. That is correct for
            ' the demo set, which really does span 90 days, but wrong for a real
            ' store from its first week - three weeks of history divided by 90
            ' understates daily consumption about fourfold and reports roughly
            ' four times more runway than the shelf actually holds.
            Dim dailyCons As Double
            If obsDays > 0 Then
                dailyCons = totalOut / obsDays
            Else
                dailyCons = 0
            End If
            
            ws.Cells(rowNum, startCol).Value = artCode
            ws.Cells(rowNum, startCol).Font.Size = 8
            ws.Cells(rowNum, startCol + 1).Value = stock
            ws.Cells(rowNum, startCol + 1).Font.Size = 8
            
            If dailyCons > 0 Then
                Dim daysLeft As Double: daysLeft = stock / dailyCons
                ws.Cells(rowNum, startCol + 2).Value = Round(dailyCons, 2)
                ws.Cells(rowNum, startCol + 2).Font.Size = 8
                ws.Cells(rowNum, startCol + 3).Value = Round(daysLeft, 0)
                ws.Cells(rowNum, startCol + 3).Font.Size = 8
                
                If daysLeft < 30 Then
                    ws.Range(ws.Cells(rowNum, startCol), ws.Cells(rowNum, startCol + 3)) _
                        .Interior.Color = RGB(255, 200, 200)
                ElseIf daysLeft < 60 Then
                    ws.Range(ws.Cells(rowNum, startCol), ws.Cells(rowNum, startCol + 3)) _
                        .Interior.Color = RGB(255, 243, 205)
                End If
            Else
                ws.Cells(rowNum, startCol + 2).Value = 0
                ws.Cells(rowNum, startCol + 2).Font.Size = 8
                ws.Cells(rowNum, startCol + 3).Value = "N/A"
                ws.Cells(rowNum, startCol + 3).Font.Size = 8
            End If
            
            rowNum = rowNum + 1
        End If
    Next i
    
    If rowNum > 3 Then
        ws.Range(ws.Cells(3, startCol), ws.Cells(rowNum - 1, startCol + 3)).Borders.LineStyle = xlContinuous
    End If
    ws.Columns("L:O").AutoFit
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'--------------------------------------------------------------------------------------
' GET OR CREATE
'--------------------------------------------------------------------------------------
Private Function GetOrCreateDashboardSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("DASHBOARD")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "DASHBOARD"
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(245, 245, 245)
    
    Set GetOrCreateDashboardSheet = ws
End Function
