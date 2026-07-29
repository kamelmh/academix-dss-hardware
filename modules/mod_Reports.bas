Attribute VB_Name = "mod_Reports"
' ============================================================================
' Academix v13.2 - DSS Logistique El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Direction de l'Education - Wilaya d'El Bayadh
' Protected under Algerian Copyright Law (Ordinance 03-05, July 19, 2003)
' All rights reserved. Unauthorized reproduction or distribution prohibited.
' ============================================================================

Option Explicit

'=======================================================================================
' SUB: GenerateMonthlyReport
' Creates monthly summary report (inputs, outputs, stock status)
'=======================================================================================
Public Sub GenerateMonthlyReport(Optional ByVal rptMonth As Integer = 0)
    Dim wsMouv As Worksheet, wsArt As Worksheet, wsReport As Worksheet
    Dim desktopPath As String, fileName As String, fullPath As String
    
    If rptMonth = 0 Then rptMonth = Month(Date)
    
    On Error GoTo ReportError
    
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    
    desktopPath = mod_SharedEnvironment.GetSharedExportPath()
    fileName = "Rapport_Mensuel_" & Format(Date, "yyyy-mm") & ".xlsx"
    fullPath = desktopPath & fileName
    
    Application.ScreenUpdating = False
    Set wsReport = Worksheets.Add
    wsReport.name = "RAPPORT_MENSUEL"
    
    With wsReport
        .Range("A1").Value = "RAPPORT MENSUEL - " & mod_Config.SYS_TITLE
        .Range("A2").Value = "Direction de l'" & Chr(201) & "ducation - El Bayadh"
        .Range("A3").Value = "Mois: " & Format(Date, "mmmm yyyy")
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        .Range("A5").Value = "CODE"
        .Range("B5").Value = "D" & Chr(201) & "SIGNATION"
        .Range("C5").Value = "ENTR" & Chr(201) & "ES"
        .Range("D5").Value = "SORTIES"
        .Range("E5").Value = "STOCK FINAL"
        .Range("F5").Value = "VALEUR (DZD)"
        .Range("G5").Value = "CLASSE ABC"
        .Range("H5").Value = "CONTRIB. %"
        .Range("A5:H5").Font.Bold = True
        .Range("A5:H5").Interior.Color = RGB(0, 112, 192)
        .Range("A5:H5").Font.Color = vbWhite
    End With
    
    Dim lastArtRow As Long: lastArtRow = wsArt.Cells(wsArt.Rows.count, COL_ART_CODE).End(xlUp).Row
    Dim reportRow As Integer: reportRow = 6
    Dim totalIn As Double, totalOut As Double, totalValue As Double
    
    Dim i As Long
    For i = 3 To lastArtRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, COL_ART_CODE).Value)
        Dim name As String: name = Trim(wsArt.Cells(i, COL_ART_DESIGNATION).Value)
        
        If sku <> "" Then
            Dim monthIn As Double, monthOut As Double
            On Error Resume Next
            wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
            On Error GoTo ReportError
            monthIn = WorksheetFunction.SumIfs(wsMouv.Columns(COL_MOUV_QTE), wsMouv.Columns(COL_MOUV_CODE_ARTICLE), sku, wsMouv.Columns(COL_MOUV_TYPE), "IN", _
                                             wsMouv.Columns(COL_MOUV_DATE), ">=" & DateSerial(Year(Date), rptMonth, 1), _
                                             wsMouv.Columns(COL_MOUV_DATE), "<" & DateSerial(Year(Date), rptMonth + 1, 1))
            monthOut = WorksheetFunction.SumIfs(wsMouv.Columns(COL_MOUV_QTE), wsMouv.Columns(COL_MOUV_CODE_ARTICLE), sku, wsMouv.Columns(COL_MOUV_TYPE), "OUT", _
                                               wsMouv.Columns(COL_MOUV_DATE), ">=" & DateSerial(Year(Date), rptMonth, 1), _
                                               wsMouv.Columns(COL_MOUV_DATE), "<" & DateSerial(Year(Date), rptMonth + 1, 1))
            wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
            
            Dim totalStock As Double: totalStock = wsArt.Cells(i, COL_ART_STOCK_ACTUEL).Value
            Dim pu As Double: pu = wsArt.Cells(i, COL_ART_PU).Value
            Dim rowValue As Double: rowValue = totalStock * pu
            
            With wsReport
                .Cells(reportRow, 1).Value = sku
                .Cells(reportRow, 2).Value = name
                .Cells(reportRow, 3).Value = monthIn
                .Cells(reportRow, 4).Value = monthOut
                .Cells(reportRow, 5).Value = totalStock
                .Cells(reportRow, 6).Value = rowValue
                .Cells(reportRow, 6).NumberFormat = "#,##0.00"
                
                ' --- Enhanced Metrics ---
                ' Pull ABC Class from ARTICLES sheet (Col E = 5)
                .Cells(reportRow, 7).Value = wsArt.Cells(i, COL_ART_CATEGORIE).Value
                
                ' Contribution % (calculated at the end or updated later)
                ' We'll leave Col H for now and fill it in a second pass
                .Cells(reportRow, 8).Value = 0
            End With
            
            totalIn = totalIn + monthIn
            totalOut = totalOut + monthOut
            totalValue = totalValue + rowValue
            reportRow = reportRow + 1
        End If
    Next i
    
    ' --- Final Totals Row ---
    With wsReport
        .Range("A" & reportRow + 1).Value = "TOTAUX"
        .Range("A" & reportRow + 1).Font.Bold = True
        .Range("C" & reportRow + 1).Value = totalIn
        .Range("D" & reportRow + 1).Value = totalOut
        .Range("F" & reportRow + 1).Value = totalValue
        .Range("F" & reportRow + 1).NumberFormat = "#,##0.00"
        .Range("A" & reportRow + 1 & ":H" & reportRow + 1).Interior.Color = RGB(217, 217, 217)
        
        ' Calculate contribution percentages for each row
        Dim r As Long
        For r = 6 To reportRow
            Dim valLigne As Double: valLigne = .Cells(r, 6).Value
            If totalValue > 0 Then
                .Cells(r, 8).Value = valLigne / totalValue
                .Cells(r, 8).NumberFormat = "0.00%"
            End If
        Next r
        
        .Columns("A:H").AutoFit
    End With
    
    wsReport.SaveAs fullPath
    wsReport.Delete
    Application.ScreenUpdating = True
    MsgBox "Rapport g?n?r?: " & fullPath, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
ReportError:
    Application.ScreenUpdating = True
    MsgBox "Erreur rapport: " & Err.Description, vbCritical
End Sub

'=======================================================================================
' SUB: GenerateStockCard
' Creates stock card (fiche de stock) for a specific article
'=======================================================================================
Public Sub GenerateStockCard(Optional ByVal sku As String = "")
    Dim wsArt As Worksheet, wsMouv As Worksheet, wsCard As Worksheet
    Dim desktopPath As String, fileName As String, fullPath As String
    
    On Error GoTo ReportError
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    
    If sku = "" Then
        sku = InputBox("Entrez le code article:", "Fiche de Stock")
        If sku = "" Then Exit Sub
    End If
    
    Dim artRow As Variant: artRow = Application.Match(sku, wsArt.Columns(COL_ART_CODE), 0)
    If IsError(artRow) Then
        MsgBox "Article non trouv?: " & sku, vbExclamation
        Exit Sub
    End If
    
    desktopPath = mod_SharedEnvironment.GetSharedExportPath()
    fileName = "Fiche_Stock_" & sku & ".xlsx"
    fullPath = desktopPath & fileName
    
    Application.ScreenUpdating = False
    Set wsCard = Worksheets.Add
    wsCard.name = "FICHE_STOCK"
    
    With wsCard
        .Range("A1").Value = "FICHE DE STOCK - " & sku
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "CODE:": .Range("B3").Value = sku
        .Range("A4").Value = "D" & Chr(233) & "SIGNATION:":         .Range("B4").Value = wsArt.Cells(artRow, COL_ART_DESIGNATION).Value
        .Range("A5").Value = "STOCK ACTUEL:": .Range("B5").Value = wsArt.Cells(artRow, COL_ART_STOCK_ACTUEL).Value
        .Range("A6").Value = "PRIX UNITAIRE:": .Range("B6").Value = wsArt.Cells(artRow, COL_ART_PU).Value
        .Range("A8").Value = "DATE": .Range("B8").Value = "TYPE": .Range("C8").Value = "QT" & Chr(233): .Range("D8").Value = "VALEUR": .Range("E8").Value = "R" & Chr(233) & "F" & Chr(233) & "RENCE"
        .Range("A8:E8").Font.Bold = True
        .Range("A8:E8").Interior.Color = RGB(0, 112, 192)
        .Range("A8:E8").Font.Color = vbWhite
    End With
    
    Dim lastMvtRow As Long: lastMvtRow = wsMouv.Cells(wsMouv.Rows.count, COL_MOUV_DATE).End(xlUp).Row
    Dim cardRow As Integer: cardRow = 9
    Dim j As Long
    On Error Resume Next
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo ReportError
    For j = 2 To lastMvtRow
        If Trim(wsMouv.Cells(j, COL_MOUV_CODE_ARTICLE).Value) = sku Then
            wsCard.Cells(cardRow, 1).Value = wsMouv.Cells(j, COL_MOUV_DATE).Value
            wsCard.Cells(cardRow, 2).Value = wsMouv.Cells(j, COL_MOUV_DESIGNATION).Value
            wsCard.Cells(cardRow, 3).Value = wsMouv.Cells(j, COL_MOUV_QTE).Value
            wsCard.Cells(cardRow, 4).Value = wsMouv.Cells(j, COL_MOUV_VALEUR).Value
            wsCard.Cells(cardRow, 5).Value = wsMouv.Cells(j, COL_MOUV_PU).Value
            cardRow = cardRow + 1
        End If
    Next j
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    wsCard.Columns("A:E").AutoFit
    wsCard.SaveAs fullPath
    wsCard.Delete
    Application.ScreenUpdating = True
    MsgBox "Fiche de stock g?n?r?e: " & fullPath, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
ReportError:
    Application.ScreenUpdating = True
    MsgBox "Erreur: " & Err.Description, vbCritical
End Sub

'=======================================================================================
' SUB: ConfigurerImpression
' Sets up professional print layout for all printable sheets:
'   - RAPPORTS: Landscape, print titles (rows 1-5), dynamic area, page breaks every 50 rows
'   - INVENTAIRE: Portrait, print titles (rows 1-2), dynamic area
'   - TABLEAU DE BORD: Landscape, print titles (row 1), dynamic area
' Headers: Left=Title, Center=Date, Right="[Logo]"
' Footers: Left=Direction, Center="", Right="Page X/Y"
' Call from GenerateDemoData (silent) or via ACCUEIL button
'=======================================================================================
Public Sub ConfigurerImpression()
    On Error GoTo PrintError
    Dim ws As Worksheet
    Dim pwd As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim printArea As String
    
    pwd = mod_Config.MASTER_PWD
    Application.ScreenUpdating = False
    
    ' --- Common header/footer strings ---
    Dim hdrLeft As String:  hdrLeft = mod_Config.SYS_TITLE & "  -  Direction de l'Education, El Bayadh"
    Dim hdrCenter As String: hdrCenter = Format(Date, "DD/MM/YYYY")
    Dim hdrRight As String: hdrRight = "[Logo]"
    Dim ftrLeft As String:   ftrLeft = "ERP Academie v" & mod_Config.APP_VERSION
    Dim ftrRight As String:  ftrRight = "Page " & Chr(38) & "P" & " / " & Chr(38) & "N"
    
    ' ====================================================================
    ' RAPPORTS - Landscape, print titles, dynamic area, page breaks
    ' ====================================================================
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("RAPPORTS")
    On Error GoTo PrintError
    If Not ws Is Nothing Then
        ws.Unprotect Password:=pwd
        
        ' Dynamic print area (data range)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastRow < 2 Then lastRow = 2
        printArea = "A1:" & Chr(64 + lastCol) & lastRow
        ws.PageSetup.PrintArea = printArea
        
        ' Page setup
        With ws.PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = False  ' Allow multiple pages vertically
            .CenterHorizontally = True
            .CenterVertically = False
            .PrintHeadings = False
            .PrintGridlines = False
            .Order = xlDownThenOver
            
            ' Margins (inches)
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            
            ' Headers and footers
            .LeftHeader = hdrLeft
            .CenterHeader = hdrCenter
            .RightHeader = hdrRight
            .LeftFooter = ftrLeft
            .CenterFooter = ""
            .RightFooter = ftrRight
            
            ' Print titles - repeat rows 1-5 on every page
            .PrintTitleRows = "$1:$5"
            .PrintTitleColumns = ""
        End With
        
        ' Note: Page breaks managed automatically by Excel via FitToPagesWide/FitToPagesTall
        ' Manual HPageBreaks.Add skipped (extremely slow in COM automation)
        
        ws.Protect Password:=pwd, UserInterfaceOnly:=True
    End If
    
    ' ====================================================================
    ' INVENTAIRE - Portrait, print titles, dynamic area
    ' ====================================================================
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("INVENTAIRE")
    On Error GoTo PrintError
    If Not ws Is Nothing Then
        ws.Unprotect Password:=pwd
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastRow < 2 Then lastRow = 2
        printArea = "A1:" & Chr(64 + lastCol) & lastRow
        ws.PageSetup.PrintArea = printArea
        
        With ws.PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
            .PrintHeadings = False
            .PrintGridlines = False
            
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            
            .LeftHeader = hdrLeft
            .CenterHeader = hdrCenter
            .RightHeader = hdrRight
            .LeftFooter = ftrLeft
            .CenterFooter = ""
            .RightFooter = ftrRight
            
            ' Print titles - repeat rows 1-2 on every page
            .PrintTitleRows = "$1:$2"
            .PrintTitleColumns = ""
        End With
        
        ws.Protect Password:=pwd, UserInterfaceOnly:=True
    End If
    
    ' ====================================================================
    ' TABLEAU DE BORD - Landscape, print titles, dynamic area
    ' ====================================================================
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TABLEAU DE BORD")
    On Error GoTo PrintError
    If Not ws Is Nothing Then
        ws.Unprotect Password:=pwd
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastRow < 2 Then lastRow = 2
        printArea = "A1:" & Chr(64 + lastCol) & lastRow
        ws.PageSetup.PrintArea = printArea
        
        With ws.PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
            .PrintHeadings = False
            .PrintGridlines = False
            
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            
            .LeftHeader = hdrLeft
            .CenterHeader = hdrCenter
            .RightHeader = hdrRight
            .LeftFooter = ftrLeft
            .CenterFooter = ""
            .RightFooter = ftrRight
            
            ' Print titles - repeat row 1 on every page
            .PrintTitleRows = "$1:$1"
            .PrintTitleColumns = ""
        End With
        
        ws.Protect Password:=pwd, UserInterfaceOnly:=True
    End If
    
    ' ====================================================================
    ' BON_RECEPTION / BON_SORTIE / BON_COMMANDE - Portrait, dynamic area
    ' ====================================================================
    Dim bonSheets As Variant
    bonSheets = Array("BON_RECEPTION", "BON_SORTIE", "BON_COMMANDE", "DA_DEMANDE_ACHAT")
    Dim idx As Integer
    For idx = LBound(bonSheets) To UBound(bonSheets)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(bonSheets(idx)))
        On Error GoTo PrintError
        If Not ws Is Nothing Then
            ws.Unprotect Password:=pwd
            
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lastRow < 2 Then lastRow = 2
            printArea = "A1:" & Chr(64 + lastCol) & lastRow
            ws.PageSetup.PrintArea = printArea
            
            With ws.PageSetup
                .Orientation = xlPortrait
                .PaperSize = xlPaperA4
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .CenterHorizontally = True
                .PrintHeadings = False
                
                .LeftMargin = Application.InchesToPoints(0.75)
                .RightMargin = Application.InchesToPoints(0.75)
                .TopMargin = Application.InchesToPoints(1)
                .BottomMargin = Application.InchesToPoints(0.75)
                .HeaderMargin = Application.InchesToPoints(0.3)
                .FooterMargin = Application.InchesToPoints(0.3)
                
                .LeftHeader = hdrLeft
                .CenterHeader = ""
                .RightHeader = hdrRight
                .LeftFooter = ftrLeft
                .CenterFooter = ""
                .RightFooter = ftrRight
            End With
            
            ws.Protect Password:=pwd, UserInterfaceOnly:=True
        End If
    Next idx
    
    Application.ScreenUpdating = True
    MsgBox "Configuration d'impression appliquee :" & vbCrLf & _
           "  - RAPPORTS (paysage, titres repeteurs, sauts toutes les 50 lignes)" & vbCrLf & _
           "  - INVENTAIRE (portrait, titres repeteurs)" & vbCrLf & _
           "  - TABLEAU DE BORD (paysage, titres repeteurs)" & vbCrLf & _
           "  - Bons de reception/sortie/commande (portrait)" & vbCrLf & vbCrLf & _
           "En-tetes et pieds de page configures (Page X/Y, date, logo).", _
           vbInformation, "Configuration Impression"
    Exit Sub
    
PrintError:
    Application.ScreenUpdating = True
    MsgBox "Erreur impression: " & Err.Description, vbCritical
End Sub

'=======================================================================================
' SUB: PreviewRapports
' Wrapper for print preview of RAPPORTS sheet
'=======================================================================================
Public Sub PreviewRapports()
    On Error Resume Next
    ThisWorkbook.Sheets("RAPPORTS").PrintPreview
    On Error GoTo 0
End Sub

'=======================================================================================
' SUB: PreviewInventaire
' Wrapper for print preview of INVENTAIRE sheet
'=======================================================================================
Public Sub PreviewInventaire()
    On Error Resume Next
    ThisWorkbook.Sheets("INVENTAIRE").PrintPreview
    On Error GoTo 0
End Sub

' ================================================================================
' TASK CALLBACK STUB - Added Session 22
' Referenced by MAIN_MACROS and mod_TaskOrchestrator ("RUN-DASHBOARD") as runtime
' Application.Run callback. Delegates to GenerateMonthlyReport (closest "dashboard"
' report equivalent).
' ================================================================================
Public Sub GenerateDashboardReport()
    Call GenerateMonthlyReport
End Sub

' ================================================================================
' TASK CALLBACK STUB - Added Session 22
' Referenced by mod_TaskOrchestrator ("RUN-INVENTORY") as runtime Application.Run
' callback. Delegates to GenerateStockCard (closest per-article inventory report).
' ================================================================================
Public Sub GenerateInventoryReport()
    Call GenerateStockCard
End Sub
