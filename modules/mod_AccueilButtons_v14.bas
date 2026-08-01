Attribute VB_Name = "mod_AccueilButtons"
Option Explicit

Public Sub SetupAccueilButtons()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ACCUEIL")
    If ws Is Nothing Then Exit Sub
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    
    ' Clear everything
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    ws.Cells.Clear
    
    ' Colors
    Dim cPrimary As Long: cPrimary = RGB(30, 60, 114)
    Dim cAccent As Long: cAccent = RGB(70, 90, 130)
    Dim cDark As Long: cDark = RGB(48, 48, 48)
    Dim cBg As Long: cBg = RGB(234, 234, 234)
    Dim cWhite As Long: cWhite = RGB(255, 255, 255)
    Dim cBlack As Long: cBlack = RGB(0, 0, 0)
    
    ' Sheet background
    ws.Cells.Interior.Color = cBg
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    
    ' Column widths
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 26
    ws.Columns("G").ColumnWidth = 20
    ws.Columns("H").ColumnWidth = 20
    
    ' === HEADER ===
    ws.Range("B1").Value = "DSS Quincaillerie"
    ws.Range("B1").Font.Size = 20
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Color = cPrimary
    
    ws.Range("B2").Value = "Systeme de Gestion d'Inventaire v14.0"
    ws.Range("B2").Font.Size = 11
    ws.Range("B2").Font.Color = cAccent
    
    ' Separator
    ws.Range("B3:H3").Interior.Color = cPrimary
    
    ' === BUSINESS INFO ===
    ws.Range("E4").Value = "Entreprise"
    ws.Range("E4").Font.Bold = True
    ws.Range("E4").Font.Color = cPrimary
    
    ws.Range("E5").Value = "Nom:"
    ws.Range("E5").Font.Bold = True
    ws.Range("F5").Value = "Quincaillerie"
    
    ws.Range("E6").Value = "Adresse:"
    ws.Range("E6").Font.Bold = True
    ws.Range("F6").Value = "Algerie"
    
    ws.Range("E7").Value = "Tel:"
    ws.Range("E7").Font.Bold = True
    ws.Range("F7").Value = "049 00 00 00"
    
    ws.Range("E8").Value = "NIF:"
    ws.Range("E8").Font.Bold = True
    ws.Range("F8").Value = "12345678901"
    
    ws.Range("E9").Value = "RC:"
    ws.Range("E9").Font.Bold = True
    ws.Range("F9").Value = "00/00-0000000A00"
    
    ' === NAVIGATION SECTION ===
    ws.Range("B12").Value = "Navigation"
    ws.Range("B12").Font.Bold = True
    ws.Range("B12").Font.Color = cPrimary
    ws.Range("B12").Font.Size = 11
    
    ' === BUTTONS (left side) ===
    Dim topPos As Double: topPos = 105
    Dim leftPos As Double: leftPos = 35
    Dim btnW As Double: btnW = 190
    Dim btnH As Double: btnH = 30
    Dim gap As Double: gap = 5
    Dim btn As Shape
    
    ' Button 1: Dashboard
    Set btn = ws.Shapes.AddShape(5, leftPos, topPos, btnW, btnH)
    btn.Name = "btnDashboard"
    btn.TextFrame.Characters.Text = "Dashboard"
    btn.TextFrame.Characters.Font.Size = 11
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cBlack
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = cPrimary
    btn.Line.Visible = 0
    btn.OnAction = "LaunchDashboard"
    
    ' Button 2: Stock Entry
    topPos = topPos + btnH + gap
    Set btn = ws.Shapes.AddShape(5, leftPos, topPos, btnW, btnH)
    btn.Name = "btnStockEntry"
    btn.TextFrame.Characters.Text = "Mouvements Stock"
    btn.TextFrame.Characters.Font.Size = 11
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cBlack
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = cAccent
    btn.Line.Visible = 0
    btn.OnAction = "LaunchStockEntry"
    
    ' Button 3: Articles
    topPos = topPos + btnH + gap
    Set btn = ws.Shapes.AddShape(5, leftPos, topPos, btnW, btnH)
    btn.Name = "btnArticles"
    btn.TextFrame.Characters.Text = "Articles"
    btn.TextFrame.Characters.Font.Size = 11
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cBlack
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = cAccent
    btn.Line.Visible = 0
    btn.OnAction = "LaunchArticles"
    
    ' Button 4: Fournisseurs
    topPos = topPos + btnH + gap
    Set btn = ws.Shapes.AddShape(5, leftPos, topPos, btnW, btnH)
    btn.Name = "btnFournisseurs"
    btn.TextFrame.Characters.Text = "Fournisseurs"
    btn.TextFrame.Characters.Font.Size = 11
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cBlack
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = cAccent
    btn.Line.Visible = 0
    btn.OnAction = "LaunchFournisseurs"
    
    ' Button 5: Search
    topPos = topPos + btnH + gap
    Set btn = ws.Shapes.AddShape(5, leftPos, topPos, btnW, btnH)
    btn.Name = "btnSearch"
    btn.TextFrame.Characters.Text = "Recherche"
    btn.TextFrame.Characters.Font.Size = 11
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cBlack
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = cAccent
    btn.Line.Visible = 0
    btn.OnAction = "LaunchSearch"
    
    ' Button 6: Reception
    topPos = topPos + btnH + gap
    Set btn = ws.Shapes.AddShape(5, leftPos, topPos, btnW, btnH)
    btn.Name = "btnReception"
    btn.TextFrame.Characters.Text = "Bon de Reception"
    btn.TextFrame.Characters.Font.Size = 11
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cBlack
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = cAccent
    btn.Line.Visible = 0
    btn.OnAction = "LaunchReception"
    
    ' Button 7: Reports
    topPos = topPos + btnH + gap
    Set btn = ws.Shapes.AddShape(5, leftPos, topPos, btnW, btnH)
    btn.Name = "btnReports"
    btn.TextFrame.Characters.Text = "Rapports"
    btn.TextFrame.Characters.Font.Size = 11
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cBlack
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = cAccent
    btn.Line.Visible = 0
    btn.OnAction = "LaunchReports"
    
    ' Button 8: Config
    topPos = topPos + btnH + gap
    Set btn = ws.Shapes.AddShape(5, leftPos, topPos, btnW, btnH)
    btn.Name = "btnConfig"
    btn.TextFrame.Characters.Text = "Configuration"
    btn.TextFrame.Characters.Font.Size = 11
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cBlack
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = cDark
    btn.Line.Visible = 0
    btn.OnAction = "LaunchConfig"
    
    ' === BACKUP BUTTONS (small, below main buttons) ===
    Dim backupTop As Double: backupTop = topPos + btnH + 20
    Dim smallBtnW As Double: smallBtnW = 80
    Dim smallBtnH As Double: smallBtnH = 28
    
    ' Button: Manual Backup
    Set btn = ws.Shapes.AddShape(5, leftPos, backupTop, smallBtnW, smallBtnH)
    btn.Name = "btnBackup"
    btn.TextFrame.Characters.Text = "Backup"
    btn.TextFrame.Characters.Font.Size = 9
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cWhite
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = RGB(40, 120, 40)  ' Green
    btn.Line.Visible = 0
    btn.OnAction = "ManualBackup"
    
    ' Button: Restore
    Set btn = ws.Shapes.AddShape(5, leftPos + smallBtnW + 8, backupTop, smallBtnW, smallBtnH)
    btn.Name = "btnRestore"
    btn.TextFrame.Characters.Text = "Restore"
    btn.TextFrame.Characters.Font.Size = 9
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Color = cWhite
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.Fill.ForeColor.RGB = RGB(180, 100, 30)  ' Orange
    btn.Line.Visible = 0
    btn.OnAction = "RestoreBackup"
    
    ' === KPI SECTION (right side, below business info) ===
    Dim kpiLeft As Double: kpiLeft = 500
    Dim kpiTop As Double: kpiTop = 150
    Dim kpiW As Double: kpiW = 130
    Dim kpiH As Double: kpiH = 45
    Dim kpi As Shape
    
    ws.Range("E15").Value = "Etat du Systeme"
    ws.Range("E15").Font.Bold = True
    ws.Range("E15").Font.Color = cPrimary
    ws.Range("E15").Font.Size = 11
    
    ' KPI 1: Articles
    Set kpi = ws.Shapes.AddShape(1, kpiLeft, kpiTop, kpiW, kpiH)
    kpi.Name = "kpiArticles"
    kpi.TextFrame.Characters.Text = "Articles: 0"
    kpi.TextFrame.Characters.Font.Size = 10
    kpi.TextFrame.Characters.Font.Bold = True
    kpi.TextFrame.Characters.Font.Color = cBlack
    kpi.TextFrame.Characters.Font.Name = "Calibri"
    kpi.Fill.ForeColor.RGB = cWhite
    kpi.Line.Visible = 0
    
    ' KPI 2: Suppliers
    Set kpi = ws.Shapes.AddShape(1, kpiLeft + kpiW + 8, kpiTop, kpiW, kpiH)
    kpi.Name = "kpiSuppliers"
    kpi.TextFrame.Characters.Text = "Fournisseurs: 0"
    kpi.TextFrame.Characters.Font.Size = 10
    kpi.TextFrame.Characters.Font.Bold = True
    kpi.TextFrame.Characters.Font.Color = cBlack
    kpi.TextFrame.Characters.Font.Name = "Calibri"
    kpi.Fill.ForeColor.RGB = cWhite
    kpi.Line.Visible = 0
    
    ' KPI 3: Stock Value (full width)
    Set kpi = ws.Shapes.AddShape(1, kpiLeft, kpiTop + kpiH + 8, kpiW * 2 + 8, kpiH)
    kpi.Name = "kpiStockValue"
    kpi.TextFrame.Characters.Text = "Valeur Stock: 0 DZD"
    kpi.TextFrame.Characters.Font.Size = 11
    kpi.TextFrame.Characters.Font.Bold = True
    kpi.TextFrame.Characters.Font.Color = cBlack
    kpi.TextFrame.Characters.Font.Name = "Calibri"
    kpi.Fill.ForeColor.RGB = cWhite
    kpi.Line.Visible = 0
    
    ' Footer
    Dim footerRow As Double: footerRow = 22
    ws.Range("B" & footerRow & ":H" & footerRow).Interior.Color = cPrimary
    ws.Range("B" & (footerRow + 1)).Value = "DSS Quincaillerie v14.0"
    ws.Range("B" & (footerRow + 1)).Font.Size = 8
    ws.Range("B" & (footerRow + 1)).Font.Color = cAccent
    
    ' Refresh KPIs with live data
    Call RefreshAccueilKPIs
    
    ' DO NOT PROTECT - buttons won't work if protected
    On Error GoTo 0
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    MsgBox "Erreur setup ACCUEIL: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
    On Error GoTo 0
End Sub

Public Sub RefreshAccueilKPIs()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ACCUEIL")
    If ws Is Nothing Then
        Debug.Print "[KPI] ACCUEIL sheet not found"
        Exit Sub
    End If
    
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets("ARTICLES")
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("FOURNISSEURS")
    
    Dim artCount As Long: artCount = 0
    Dim supCount As Long: supCount = 0
    Dim stockVal As Double: stockVal = 0
    
    If Not wsA Is Nothing Then
        Dim lr As Long: lr = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
        Debug.Print "[KPI] ARTICLES last row: " & lr
        If lr > 1 Then
            artCount = lr - 1
            Dim i As Long
            For i = 2 To lr
                stockVal = stockVal + (CDbl(wsA.Cells(i, 3).Value) * CDbl(wsA.Cells(i, 8).Value))
            Next i
        End If
    End If
    If Not wsS Is Nothing Then
        Dim lrS As Long: lrS = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
        Debug.Print "[KPI] FOURNISSEURS last row: " & lrS
        If lrS > 1 Then supCount = lrS - 1
    End If
    
    Debug.Print "[KPI] Counts: articles=" & artCount & " suppliers=" & supCount & " stock=" & stockVal
    
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "kpiArticles" Then
            shp.TextFrame.Characters.Text = "Articles: " & artCount
            Debug.Print "[KPI] Updated kpiArticles to " & artCount
        ElseIf shp.Name = "kpiSuppliers" Then
            shp.TextFrame.Characters.Text = "Fournisseurs: " & supCount
            Debug.Print "[KPI] Updated kpiSuppliers to " & supCount
        ElseIf shp.Name = "kpiStockValue" Then
            shp.TextFrame.Characters.Text = "Valeur Stock: " & Format(stockVal, "#,##0") & " DZD"
            Debug.Print "[KPI] Updated kpiStockValue to " & stockVal
        End If
    Next shp
    On Error GoTo 0
    Exit Sub
    
ErrorHandler:
    Debug.Print "[KPI] Error: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub LaunchDashboard()
    On Error Resume Next
    Dim frmName As String: frmName = "frmDashboard"
    If Not FormExists(frmName) Then
        MsgBox frmName & " pas encore construit. Lancez MasterSetup.", vbExclamation, mod_Config.SYS_TITLE
    Else
        VBA.UserForms.Add(frmName).Show
    End If
    On Error GoTo 0
End Sub

Public Sub LaunchStockEntry()
    On Error Resume Next
    Dim frmName As String: frmName = "frmStockEntry"
    If Not FormExists(frmName) Then
        MsgBox frmName & " pas encore construit. Lancez MasterSetup.", vbExclamation, mod_Config.SYS_TITLE
    Else
        VBA.UserForms.Add(frmName).Show
    End If
    On Error GoTo 0
End Sub

Public Sub LaunchArticles()
    On Error Resume Next
    Dim frmName As String: frmName = "frmArticleEditor"
    If Not FormExists(frmName) Then
        MsgBox frmName & " pas encore construit. Lancez MasterSetup.", vbExclamation, mod_Config.SYS_TITLE
    Else
        VBA.UserForms.Add(frmName).Show
    End If
    On Error GoTo 0
End Sub

Public Sub LaunchFournisseurs()
    On Error Resume Next
    Dim frmName As String: frmName = "frmSupplierEditor"
    If Not FormExists(frmName) Then
        MsgBox frmName & " pas encore construit. Lancez MasterSetup.", vbExclamation, mod_Config.SYS_TITLE
    Else
        VBA.UserForms.Add(frmName).Show
    End If
    On Error GoTo 0
End Sub

Public Sub LaunchSearch()
    On Error Resume Next
    Dim frmName As String: frmName = "frmSearch"
    If Not FormExists(frmName) Then
        MsgBox frmName & " pas encore construit. Lancez MasterSetup.", vbExclamation, mod_Config.SYS_TITLE
    Else
        VBA.UserForms.Add(frmName).Show
    End If
    On Error GoTo 0
End Sub

Public Sub LaunchReception()
    On Error Resume Next
    Dim frmName As String: frmName = "frmReception"
    If Not FormExists(frmName) Then
        MsgBox frmName & " pas encore construit. Lancez MasterSetup.", vbExclamation, mod_Config.SYS_TITLE
    Else
        VBA.UserForms.Add(frmName).Show
    End If
    On Error GoTo 0
End Sub

Public Sub LaunchReports()
    On Error Resume Next
    Dim frmName As String: frmName = "frmReports"
    If Not FormExists(frmName) Then
        MsgBox frmName & " pas encore construit. Lancez MasterSetup.", vbExclamation, mod_Config.SYS_TITLE
    Else
        VBA.UserForms.Add(frmName).Show
    End If
    On Error GoTo 0
End Sub

Public Sub LaunchConfig()
    On Error Resume Next
    Dim frmName As String: frmName = "frmConfig"
    If Not FormExists(frmName) Then
        MsgBox frmName & " pas encore construit. Lancez MasterSetup.", vbExclamation, mod_Config.SYS_TITLE
    Else
        VBA.UserForms.Add(frmName).Show
    End If
    On Error GoTo 0
End Sub

Private Function FormExists(ByVal formName As String) As Boolean
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(formName)
    FormExists = Not (vbComp Is Nothing)
    On Error GoTo 0
End Function
