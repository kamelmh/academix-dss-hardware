Attribute VB_Name = "mod_AccueilDesign"
' ============================================================================
' Academix v14.0 - ACCUEIL Sheet Redesign v2
' Clean modern dashboard with proper colors and layout
' ============================================================================

Option Explicit

' ============================================================================
' MAIN REDESIGN PROCEDURE
' ============================================================================
Public Sub RedesignAccueil()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ACCUEIL)
    
    ' Unprotect
    On Error Resume Next
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0
    
    ' Clear existing content
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone
    ws.Cells.Borders.LineStyle = xlNone
    
    ' Remove existing shapes
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' === SET LAYOUT ===
    SetLayout ws
    
    ' === DRAW SECTIONS ===
    DrawHeaderSection ws
    DrawLeftPanel ws
    DrawMiddlePanel ws
    DrawRightPanel ws
    DrawFooterSection ws
    
    ' === ADD BUTTONS ===
    AddNavigationButtons ws
    
    ' Freeze and activate
    ws.Activate
    ws.Range("A1").Select
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "ACCUEIL redesigned!", vbInformation, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SET LAYOUT - Column widths and row heights
' ============================================================================
Private Sub SetLayout(ws As Worksheet)
    ' Column widths - wider separation between panels
    ws.Columns("A").ColumnWidth = 1     ' Left margin
    ws.Columns("B").ColumnWidth = 18    ' Left panel (buttons)
    ws.Columns("C").ColumnWidth = 1     ' Spacer
    ws.Columns("D").ColumnWidth = 15    ' Middle panel (KPIs)
    ws.Columns("E").ColumnWidth = 1     ' Spacer
    ws.Columns("F").ColumnWidth = 12    ' Right panel labels
    ws.Columns("G").ColumnWidth = 15    ' Right panel values
    ws.Columns("H").ColumnWidth = 1     ' Right margin
    
    ' Row heights (total 25 rows)
    ws.Rows(1).RowHeight = 8     ' Top margin
    ws.Rows(2).RowHeight = 40    ' Title
    ws.Rows(3).RowHeight = 20    ' Subtitle
    ws.Rows(4).RowHeight = 4     ' Accent bar
    ws.Rows(5).RowHeight = 12    ' Spacer
    ws.Rows(6).RowHeight = 18    ' Section headers
    ws.Rows(7).RowHeight = 6     ' Spacer
    ws.Rows(8).RowHeight = 26    ' Button 1 / KPI 1
    ws.Rows(9).RowHeight = 26    ' KPI 1 value
    ws.Rows(10).RowHeight = 26   ' Button 2
    ws.Rows(11).RowHeight = 26   ' KPI 2 / Button 3
    ws.Rows(12).RowHeight = 26   ' KPI 2 value
    ws.Rows(13).RowHeight = 26   ' Button 4
    ws.Rows(14).RowHeight = 26   ' KPI 3 / Button 5
    ws.Rows(15).RowHeight = 26   ' KPI 3 value
    ws.Rows(16).RowHeight = 26   ' Button 6
    ws.Rows(17).RowHeight = 20   ' Spacer
    ws.Rows(18).RowHeight = 26   ' Button 7
    ws.Rows(19).RowHeight = 26   ' Button 8
    ws.Rows(20).RowHeight = 20   ' Spacer
    ws.Rows(21).RowHeight = 22   ' Backup/Restore
    ws.Rows(22).RowHeight = 8    ' Spacer
    ws.Rows(23).RowHeight = 18   ' Version info
    ws.Rows(24).RowHeight = 8    ' Spacer
    ws.Rows(25).RowHeight = 24   ' Footer
End Sub

' ============================================================================
' DRAW HEADER SECTION (Rows 1-4)
' ============================================================================
Private Sub DrawHeaderSection(ws As Worksheet)
    Dim i As Long
    Dim j As Long
    
    ' Title background - DARK BLUE
    For i = 1 To 4
        For j = 1 To 8
            ws.Cells(i, j).Interior.Color = RGB(26, 60, 114)  ' #1A3C72
        Next j
    Next i
    
    ' Title text
    With ws.Range("B2:G2")
        .Merge
        .Value = "QUINCAILLERIE"
        .Font.Name = "Segoe UI"
        .Font.Size = 26
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Subtitle
    With ws.Range("B3:G3")
        .Merge
        .Value = "Systeme de Gestion d'Inventaire v14.0"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Color = RGB(180, 195, 220)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Accent bar - RED
    For j = 1 To 8
        ws.Cells(4, j).Interior.Color = RGB(233, 69, 96)  ' #E94560
    Next j
End Sub

' ============================================================================
' DRAW LEFT PANEL (Navigation - Column B)
' ============================================================================
Private Sub DrawLeftPanel(ws As Worksheet)
    ' Panel background
    Dim i As Long
    For i = 5 To 25
        ws.Cells(i, 2).Interior.Color = RGB(245, 247, 250)  ' Light gray
    Next i
    
    ' Section header
    With ws.Range("B6")
        .Value = "Navigation"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(26, 60, 114)
        .HorizontalAlignment = xlLeft
    End With
End Sub

' ============================================================================
' DRAW MIDDLE PANEL (KPIs - Column D)
' ============================================================================
Private Sub DrawMiddlePanel(ws As Worksheet)
    ' Panel background
    Dim i As Long
    For i = 5 To 25
        ws.Cells(i, 4).Interior.Color = RGB(250, 250, 252)
    Next i
    
    ' Section header
    With ws.Range("D6")
        .Value = "Etat du Systeme"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(26, 60, 114)
        .HorizontalAlignment = xlLeft
    End With
    
    ' KPI 1: Articles (Rows 8-9)
    With ws.Range("D8")
        .Value = "Articles"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = RGB(100, 110, 130)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With
    
    With ws.Range("D9")
        .Value = GetArticleCount()
        .Font.Name = "Segoe UI"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(26, 60, 114)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With
    
    ' Border for KPI 1
    With ws.Range("D8:D9")
        .Borders(xlEdgeTop).Color = RGB(219, 226, 239)
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(219, 226, 239)
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).Color = RGB(219, 226, 239)
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(219, 226, 239)
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    
    ' KPI 2: Suppliers (Rows 11-12)
    With ws.Range("D11")
        .Value = "Fournisseurs"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = RGB(100, 110, 130)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With
    
    With ws.Range("D12")
        .Value = GetSupplierCount()
        .Font.Name = "Segoe UI"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(26, 60, 114)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With
    
    ' Border for KPI 2
    With ws.Range("D11:D12")
        .Borders(xlEdgeTop).Color = RGB(219, 226, 239)
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(219, 226, 239)
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).Color = RGB(219, 226, 239)
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(219, 226, 239)
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    
    ' KPI 3: Stock Value (Rows 14-15)
    With ws.Range("D14")
        .Value = "Valeur Stock"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = RGB(100, 110, 130)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With
    
    With ws.Range("D15")
        .Value = GetStockValue()
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(233, 69, 96)  ' Red accent
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With
    
    ' Border for KPI 3
    With ws.Range("D14:D15")
        .Borders(xlEdgeTop).Color = RGB(219, 226, 239)
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(219, 226, 239)
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).Color = RGB(219, 226, 239)
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(219, 226, 239)
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    
    ' System info
    With ws.Range("D22")
        .Value = "Version"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Color = RGB(150, 160, 170)
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("D23")
        .Value = "v14.0"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(26, 60, 114)
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ============================================================================
' DRAW RIGHT PANEL (Company Info - Columns F-G)
' ============================================================================
Private Sub DrawRightPanel(ws As Worksheet)
    ' Panel background
    Dim i As Long
    For i = 5 To 25
        ws.Cells(i, 6).Interior.Color = RGB(250, 250, 252)
        ws.Cells(i, 7).Interior.Color = RGB(250, 250, 252)
    Next i
    
    ' Section header
    With ws.Range("F6")
        .Value = "Entreprise"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(26, 60, 114)
        .HorizontalAlignment = xlLeft
    End With
    
    ' Company details
    Dim labels As Variant
    Dim values As Variant
    labels = Array("Nom:", "Adresse:", "Tel:", "NIF:", "RC:")
    values = Array("Quincaillerie", "Algerie", "049 00 00 00", "12345678901", "00/00-0000000A00")
    
    Dim r As Long
    For i = 0 To 4
        r = 8 + i
        
        ' Label
        With ws.Cells(r, 6)
            .Value = labels(i)
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(26, 60, 114)
            .HorizontalAlignment = xlRight
        End With
        
        ' Value
        With ws.Cells(r, 7)
            .Value = values(i)
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Color = RGB(64, 64, 64)
            .HorizontalAlignment = xlLeft
        End With
    Next i
End Sub

' ============================================================================
' DRAW FOOTER SECTION (Row 25)
' ============================================================================
Private Sub DrawFooterSection(ws As Worksheet)
    Dim j As Long
    
    ' Footer background
    For j = 1 To 8
        ws.Cells(25, j).Interior.Color = RGB(26, 60, 114)  ' Dark blue
    Next j
    
    ' Footer text
    With ws.Range("B25:G25")
        .Merge
        .Value = "DSS Quincaillerie v14.0"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = RGB(180, 195, 220)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' ============================================================================
' ADD NAVIGATION BUTTONS
' ============================================================================
Private Sub AddNavigationButtons(ws As Worksheet)
    ' Button dimensions - fit strictly within column B
    Dim btnW As Double: btnW = 105  ' Width within column B
    Dim btnH As Double: btnH = 20   ' Height fits in row
    Dim btnL As Double: btnL = 12   ' Left edge within column B
    
    ' Row positions match SetLayout
    ' Rows 8,10,11,13,14,16,18,19 are for buttons
    ' Each row height = 26
    
    ' Button 1: Dashboard (Row 8)
    CreateButton ws, "btnDashboard", "Dashboard", btnL, 72, btnW, btnH, RGB(26, 60, 114), "MAIN_MACROS.ShowDashboard"
    
    ' Button 2: Stock Entry (Row 10)
    CreateButton ws, "btnStockEntry", "Mouvements Stock", btnL, 124, btnW, btnH, RGB(0, 70, 127), "MAIN_MACROS.AjouterMouvement"
    
    ' Button 3: Articles (Row 11)
    CreateButton ws, "btnArticles", "Articles", btnL, 150, btnW, btnH, RGB(0, 90, 150), "MAIN_MACROS.ShowArticleEditor"
    
    ' Button 4: Suppliers (Row 13)
    CreateButton ws, "btnFournisseurs", "Fournisseurs", btnL, 202, btnW, btnH, RGB(0, 110, 170), "MAIN_MACROS.ShowSupplierEditor"
    
    ' Button 5: Search (Row 14)
    CreateButton ws, "btnSearch", "Recherche", btnL, 228, btnW, btnH, RGB(0, 80, 140), "MAIN_MACROS.ShowSearch"
    
    ' Button 6: Reception (Row 16)
    CreateButton ws, "btnReception", "Bon de Reception", btnL, 280, btnW, btnH, RGB(0, 100, 160), "MAIN_MACROS.ShowReception"
    
    ' Button 7: Reports (Row 18)
    CreateButton ws, "btnReports", "Rapports", btnL, 332, btnW, btnH, RGB(0, 80, 140), "MAIN_MACROS.ShowReports"
    
    ' Button 8: Config (Row 19)
    CreateButton ws, "btnConfig", "Configuration", btnL, 358, btnW, btnH, RGB(0, 60, 120), "MAIN_MACROS.ShowConfig"
    
    ' Backup/Restore buttons (Row 21)
    CreateButton ws, "btnBackup", "Backup", btnL, 410, 50, 18, RGB(233, 69, 96), "MAIN_MACROS.BackupNow"
    CreateButton ws, "btnRestore", "Restaurer", btnL + 55, 410, 50, 18, RGB(200, 50, 80), "MAIN_MACROS.RestoreFromBackup"
End Sub

' ============================================================================
' CREATE SINGLE BUTTON
' ============================================================================
Private Sub CreateButton(ws As Worksheet, name As String, caption As String, _
                         left As Double, top As Double, width As Double, height As Double, _
                         bgColor As Long, action As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)
    
    With shp
        .Name = name
        .Fill.ForeColor.RGB = bgColor
        .Line.Visible = msoFalse
        
        ' Subtle shadow
        .Shadow.Visible = msoTrue
        .Shadow.Type = msoShadow21
        .Shadow.Blur = 6
        .Shadow.OffsetX = 1
        .Shadow.OffsetY = 1
        .Shadow.Transparency = 0.75
        
        ' Text
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Name = "Segoe UI"
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        
        ' OnAction
        .OnAction = action
    End With
End Sub

' ============================================================================
' RESET TO DEFAULT
' ============================================================================
Public Sub ResetAccueil()
    If MsgBox("Reset ACCUEIL to default?", vbQuestion + vbYesNo, mod_Config.SYS_TITLE) = vbNo Then
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ACCUEIL)
    
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone
    ws.Cells.Borders.LineStyle = xlNone
    
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    MsgBox "ACCUEIL reset. Run RedesignAccueil to recreate.", vbInformation, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' HELPER FUNCTIONS - Dynamic KPI calculations
' ============================================================================

' Get article count from ARTICLES sheet
Private Function GetArticleCount() As String
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        GetArticleCount = "0"
    Else
        GetArticleCount = CStr(lastRow - 1)
    End If
    Exit Function
    
ErrorHandler:
    GetArticleCount = "0"
End Function

' Get supplier count from FOURNISSEURS sheet
Private Function GetSupplierCount() As String
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_FOURNISSEURS)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        GetSupplierCount = "0"
    Else
        GetSupplierCount = CStr(lastRow - 1)
    End If
    Exit Function
    
ErrorHandler:
    GetSupplierCount = "0"
End Function

' Get total stock value from ARTICLES sheet (Stock * PU)
Private Function GetStockValue() As String
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        GetStockValue = "0 DZD"
        Exit Function
    End If
    
    Dim totalValue As Double
    Dim i As Long
    
    For i = 2 To lastRow
        Dim stock As Double
        Dim pu As Double
        
        stock = Val(ws.Cells(i, mod_Config.COL_ART_STOCK).Value)
        pu = Val(ws.Cells(i, mod_Config.COL_ART_PU).Value)
        
        totalValue = totalValue + (stock * pu)
    Next i
    
    ' Format with comma separator
    GetStockValue = Format(totalValue, "#,##0") & " DZD"
    Exit Function
    
ErrorHandler:
    GetStockValue = "0 DZD"
End Function
