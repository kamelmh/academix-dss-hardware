Attribute VB_Name = "mod_DemoData"
' ============================================================================
' Academix v13.2 - DSS Logistique El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Direction de l'Education - Wilaya d'El Bayadh
' Protected under Algerian Copyright Law (Ordinance 03-05, July 19, 2003)
' All rights reserved. Unauthorized reproduction or distribution prohibited.
' ============================================================================

'   - ~150 movements (mix of IN/OUT across all 15 articles)
'   - Real values from MOUVEMENTS: D=789, ROP=206, SS=200, Q*=37, LT=2
'==============================================================================

Option Explicit

'================================================================================
' DEMO DATA GENERATOR - 38-Day Observation Period
'================================================================================

Public Sub GenerateDemoData()
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call SeedArticles:          If Err.Number <> 0 Then Err.Clear
    Call SeedSuppliers:         If Err.Number <> 0 Then Err.Clear
    Call SeedMovements:         If Err.Number <> 0 Then Err.Clear
    ' Test individual: SeedBarcodesSilent only added next
    Call SeedInitialStock:      If Err.Number <> 0 Then Err.Clear
    Call SeedBarcodesSilent:    If Err.Number <> 0 Then Err.Clear
    Call ConfigurerImpressionSilent: If Err.Number <> 0 Then Err.Clear
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then Err.Clear
End Sub

'================================================================================
' SEED ARTICLES - Ensure all 15 articles exist with initial stock
'================================================================================

Private Sub SeedArticles()
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        MsgBox "Feuille ARTICLES introuvable.", vbCritical
        Exit Sub
    End If
    
    wsArt.Unprotect Password:=mod_Config.MASTER_PWD
    
    ' Clear all existing content
    Dim lastRow As Long
    lastRow = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    If lastRow > 1 Then wsArt.Rows("2:" & lastRow).Delete
    
    ' 15 articles (incl. ART-013/014/015 for BORDEREAU_COMMANDE compatibility)
    Dim articles As Variant
    articles = Array( _
        Array("ART-001", "Rame papier A4 80g/m²", "Fournitures Bureau", "F-002", 400, 400, 400, "A", "Papeterie standard"), _
        Array("ART-002", "Toner imprimante G030 (noir)", "Informatique", "F-001", 3000, 3000, 3, "B", "Fournitures d'impression"), _
        Array("ART-003", "Rame papier A3 80g/m2", "Fournitures Bureau", "F-002", 400, 1200, 80, "B", "Papier grand format"), _
        Array("ART-004", "Boite archives carton", "Admin", "F-003", 300, 350, 50, "B", "Archivage physique"), _
        Array("ART-005", "Agrafeuse de bureau", "Fournitures Bureau", "F-003", 150, 280, 30, "C", "Petit materiel"), _
        Array("ART-006", "Stylos bille boite/50", "Fournitures Bureau", "F-002", 500, 420, 60, "B", "Consommables ecriture"), _
        Array("ART-007", "Registre grand format 5m", "Admin", "F-003", 200, 680, 40, "C", "Registres officiels"), _
        Array("ART-008", "Encre tampon", "Fournitures Bureau", "F-001", 100, 180, 20, "C", "Consommables tampon"), _
        Array("ART-009", "Sous-chemise carton", "Fournitures Bureau", "F-002", 600, 95, 70, "B", "Chemises classement"), _
        Array("ART-010", "Chemise cartonnee", "Fournitures Bureau", "F-002", 450, 120, 50, "B", "Chemises documents"), _
        Array("ART-011", "Rouleau papier fax", "Informatique", "F-001", 80, 550, 15, "C", "Consommables fax"), _
        Array("ART-012", "Marqueur permanent noir", "Fournitures Bureau", "F-003", 350, 230, 40, "C", "Marquage etatiquetage"), _
        Array("ART-013", "Encre pour cachets", "Fournitures Bureau", "F-001", 10, 450, 5, "C", "Encre sceaux officiels"), _
        Array("ART-014", "Classeur a levier", "Admin", "F-003", 12, 550, 5, "C", "Classement archives"), _
        Array("ART-015", "Cartouche toner generique", "Informatique", "F-001", 39, 4500, 10, "B", "Toner compatible HP") _
    )
    
    Dim i As Long
    For i = 0 To UBound(articles)
        Dim rowIdx As Long
        rowIdx = 2 + i
        
        wsArt.Cells(rowIdx, COL_ART_CODE).Value = articles(i)(0)   ' CODE
        wsArt.Cells(rowIdx, COL_ART_DESIGNATION).Value = articles(i)(1)   ' DESIGNATION
        wsArt.Cells(rowIdx, COL_ART_STOCK).Value = articles(i)(4)   ' STOCK INITIAL
        wsArt.Cells(rowIdx, COL_ART_SEUIL_MIN).Value = ""               ' SEUIL_MIN (auto)
        wsArt.Cells(rowIdx, COL_ART_CATEGORIE).Value = articles(i)(2)   ' CATEGORIE
        wsArt.Cells(rowIdx, COL_ART_CLASSE_ABC).Value = articles(i)(7)   ' CLASSE ABC
        wsArt.Cells(rowIdx, COL_ART_STOCK_ACTUEL).Value = articles(i)(4)   ' STOCK ACTUEL (same as initial)
        wsArt.Cells(rowIdx, COL_ART_PU).Value = articles(i)(5)   ' PU (DZD)
        wsArt.Cells(rowIdx, COL_ART_FOURNISSEUR).Value = articles(i)(3)   ' FOURNISSEUR
        wsArt.Cells(rowIdx, COL_ART_STOCK_SECURITE).Value = articles(i)(6)  ' STOCK SECURITE
        wsArt.Cells(rowIdx, COL_ART_NOTES).Value = articles(i)(8)  ' NOTES
        wsArt.Cells(rowIdx, COL_ART_CMUP).Value = ""              ' CMUP (auto)
    Next i
    
    wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'================================================================================
' SEED SUPPLIERS - 3 suppliers with DGI tax IDs
'================================================================================

Private Sub SeedSuppliers()
    Dim wsFou As Worksheet
    On Error Resume Next
    Set wsFou = ThisWorkbook.Sheets("FOURNISSEURS")
    On Error GoTo 0
    
    If wsFou Is Nothing Then
        Exit Sub  ' Supplier sheet optional - mod_SupplierRegistry has hardcoded data
    End If
    
    wsFou.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim lastRow As Long
    lastRow = wsFou.Cells(wsFou.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then wsFou.Rows("3:" & lastRow).Delete
    
    ' Column order matches actual FOURNISSEURS sheet headers (9-col evaluation layout):
    ' A=Code, B=Nom abrege, C=Raison sociale, D=Wilaya, E=Telephone, F=Classe, G=Delai, H=Note, I=Specialite
    ' Data order: Code, NomAbrege, RaisonSociale, Wilaya, Telephone, Classe, Delai, Note, Specialite
    Dim suppliers As Variant
    suppliers = Array( _
        Array("F-001", "ENAP Alger", "Entreprise Nationale des Arts Plastiques, Alger", "Alger", "023-XXX-XXXX", "A", 5, 92, "Arts Graphiques & Papeterie"), _
        Array("F-002", "Bureautique Oran", "Societe de Fournitures de Bureau, Oran", "Oran", "041-XXX-XXXX", "B", 8, 78, "Fournitures de Bureau"), _
        Array("F-003", "Bureau Plus", "Bureau Plus Distribution, El Bayadh", "El Bayadh", "049-XXX-XXXX", "A", 3, 95, "Materiel Scolaire & Informatique") _
    )
    
    Dim i As Long
    For i = 0 To UBound(suppliers)
        Dim rowIdx As Long
        rowIdx = 3 + i
        
        wsFou.Cells(rowIdx, COL_FOU_CODE).Value = suppliers(i)(0)        ' Code
        wsFou.Cells(rowIdx, COL_FOU_NOM_ABREGE).Value = suppliers(i)(1)  ' Nom abrege
        wsFou.Cells(rowIdx, COL_FOU_RAISON_SOCIALE).Value = suppliers(i)(2) ' Raison sociale
        wsFou.Cells(rowIdx, COL_FOU_WILAYA).Value = suppliers(i)(3)     ' Wilaya
        wsFou.Cells(rowIdx, COL_FOU_TELEPHONE).Value = suppliers(i)(4)  ' Telephone
        wsFou.Cells(rowIdx, COL_FOU_CLASSE).Value = suppliers(i)(5)     ' Classe
        wsFou.Cells(rowIdx, COL_FOU_DELAI).Value = suppliers(i)(6)      ' Delai
        wsFou.Cells(rowIdx, COL_FOU_NOTE).Value = suppliers(i)(7)       ' Note /100
        wsFou.Cells(rowIdx, COL_FOU_SPECIALITE).Value = suppliers(i)(8) ' Specialite
    Next i
    
    wsFou.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'================================================================================
' SEED MOVEMENTS - ~150 realistic movements over 38 days
'================================================================================

Private Sub SeedMovements()
    Dim wsMouv As Worksheet
    On Error Resume Next
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    On Error GoTo 0
    
    If wsMouv Is Nothing Then
        MsgBox "Feuille MOUVEMENTS introuvable.", vbCritical
        Exit Sub
    End If
    
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    
    ' Clear existing data
    Dim lastRow As Long
    lastRow = wsMouv.Cells(wsMouv.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then wsMouv.Rows("3:" & lastRow).Delete
    
    Dim startDate As Date
    Dim endDate As Date
    Dim dayNum As Long
    Dim mvtDate As Date
    Dim rowIdx As Long
    rowIdx = 3
    
    startDate = DateSerial(2026, 3, 1)
    endDate = DateSerial(2026, 4, 7)
    
    ' Predefined movement patterns per article (realistic consumption)
    ' Format: (ART_CODE, OUT_DAY_PATTERN, OUT_QTY_PATTERN, IN_DAY_PATTERN, IN_QTY)
    Dim patterns As Variant
    patterns = Array( _
        Array("ART-001", Array(3, 7, 12, 18, 25, 32), Array(15, 20, 18, 22, 25, 20), Array(10, 28), Array(80, 100)), _
        Array("ART-002", Array(2, 5, 8, 12, 15, 18, 22, 25, 28, 32, 35), Array(30, 25, 35, 28, 30, 32, 28, 30, 35, 25, 30), Array(8, 24), Array(150, 200)), _
        Array("ART-003", Array(4, 10, 17, 24, 31), Array(12, 15, 10, 18, 14), Array(14), Array(80)), _
        Array("ART-004", Array(6, 16, 26), Array(20, 15, 25), Array(20), Array(60)), _
        Array("ART-005", Array(11, 27), Array(5, 8), Array(4), Array(15)), _
        Array("ART-006", Array(3, 9, 14, 20, 27, 33), Array(15, 18, 20, 12, 22, 15), Array(5, 15), Array(55, 50)), _
        Array("ART-007", Array(8, 22), Array(10, 12), Array(3), Array(25)), _
        Array("ART-008", Array(15, 30), Array(3, 5), Array(4), Array(10)), _
        Array("ART-009", Array(5, 12, 19, 26, 34), Array(25, 30, 20, 28, 22), Array(6, 18), Array(65, 65)), _
        Array("ART-010", Array(7, 14, 21, 28, 35), Array(18, 22, 15, 20, 18), Array(20), Array(100)), _
        Array("ART-011", Array(10, 25), Array(2, 3), Array(3), Array(8)), _
        Array("ART-012", Array(9, 18, 28), Array(8, 12, 10), Array(3), Array(35)), _
        Array("ART-013", Array(6, 20), Array(2, 3), Array(3), Array(8)), _
        Array("ART-014", Array(12, 26), Array(3, 5), Array(3), Array(10)), _
        Array("ART-015", Array(10, 22, 34), Array(5, 8, 6), Array(18), Array(30)) _
    )
    
    Dim docCounters As Object
    Set docCounters = CreateObject("Scripting.Dictionary")
    docCounters("BS") = 1
    docCounters("BR") = 1
    
    Dim pIdx As Long
    For pIdx = 0 To UBound(patterns)
        Dim artCode As String
        artCode = CStr(patterns(pIdx)(0))
        
        ' OUT movements
        Dim outDays As Variant
        Dim outQtys As Variant
        outDays = patterns(pIdx)(1)
        outQtys = patterns(pIdx)(2)
        
        Dim dIdx As Long
        For dIdx = 0 To UBound(outDays)
            If outDays(dIdx) > 0 Then
                mvtDate = startDate + outDays(dIdx) - 1
                If mvtDate <= endDate Then
                    Dim bsNum As Long
                    bsNum = docCounters("BS")
                    docCounters("BS") = bsNum + 1
                    
                    Dim bsRef As String
                    bsRef = REFDOC_PREFIX & "2026-" & Format(bsNum, "0000")
                    
                    wsMouv.Cells(rowIdx, COL_MOUV_DATE).Value = mvtDate
                    wsMouv.Cells(rowIdx, COL_MOUV_CODE_ARTICLE).Value = artCode
                    wsMouv.Cells(rowIdx, COL_MOUV_TYPE).Value = "OUT"
                    wsMouv.Cells(rowIdx, COL_MOUV_QTE).Value = outQtys(dIdx)
                    wsMouv.Cells(rowIdx, COL_MOUV_REF_DOC).Value = bsRef
                    
                    ' Get PU from ARTICLES
                    wsMouv.Cells(rowIdx, COL_MOUV_PU).Value = GetArticlePU(artCode)
                    
                    ' Random service
                    wsMouv.Cells(rowIdx, COL_MOUV_THIRD_PARTY).Value = RandomService()
                    
                    rowIdx = rowIdx + 1
                End If
            End If
        Next dIdx
        
        ' IN movements (reorders)
        Dim inDays As Variant
        Dim inQtys As Variant
        inDays = patterns(pIdx)(3)
        inQtys = patterns(pIdx)(4)
        
        For dIdx = 0 To UBound(inDays)
            If inDays(dIdx) > 0 Then
                mvtDate = startDate + inDays(dIdx) - 1
                If mvtDate <= endDate Then
                    Dim brNum As Long
                    brNum = docCounters("BR")
                    docCounters("BR") = brNum + 1
                    
                    Dim brRef As String
                    brRef = "BR-2026-" & Format(brNum, "0000")
                    
                    wsMouv.Cells(rowIdx, COL_MOUV_DATE).Value = mvtDate
                    wsMouv.Cells(rowIdx, COL_MOUV_CODE_ARTICLE).Value = artCode
                    wsMouv.Cells(rowIdx, COL_MOUV_TYPE).Value = "IN"
                    wsMouv.Cells(rowIdx, COL_MOUV_QTE).Value = inQtys(dIdx)
                    wsMouv.Cells(rowIdx, COL_MOUV_REF_DOC).Value = brRef
                    
                    wsMouv.Cells(rowIdx, COL_MOUV_PU).Value = GetArticlePU(artCode)
                    
                    ' Assign supplier based on article
                    wsMouv.Cells(rowIdx, COL_MOUV_THIRD_PARTY).Value = GetArticleSupplier(artCode)
                    
                    rowIdx = rowIdx + 1
                End If
            End If
        Next dIdx
    Next pIdx
    
    ' Calculate LINE_VALUE for all rows
    Dim i As Long
    For i = 3 To rowIdx - 1
        Dim qty As Double
        Dim pu As Double
        qty = wsMouv.Cells(i, COL_MOUV_QTE).Value
        pu = wsMouv.Cells(i, COL_MOUV_PU).Value
        wsMouv.Cells(i, COL_MOUV_VALEUR).Value = qty * pu
    Next i
    
    ' Sort by date
    Dim sortRange As Range
    Set sortRange = wsMouv.Range("A3:L" & (rowIdx - 1))
    sortRange.Sort Key1:=wsMouv.Range("A3"), Order1:=xlAscending, Header:=xlNo
    
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    Debug.Print "[DemoData] Generated " & (rowIdx - 3) & " movements over 38 days"
End Sub

'================================================================================
' UPDATE INITIAL STOCK - Calculate current stock from movements
'================================================================================

Private Sub SeedInitialStock()
    ' ... existing code ...
End Sub

'================================================================================
' SILENT PRINT LAYOUT - called during build, no UI
' Mirrors mod_Reports.ConfigurerImpression settings
'================================================================================
Private Sub ConfigurerImpressionSilent()
    On Error Resume Next
    Dim ws As Worksheet
    Dim pwd As String
    Dim lastRow As Long, lastCol As Long
    Dim printArea As String
    pwd = mod_Config.MASTER_PWD
    
    ' Common header/footer
    Dim hdrLeft As String:  hdrLeft = mod_Config.SYS_TITLE & "  -  Direction de l'Education, El Bayadh"
    Dim hdrRight As String: hdrRight = "[Logo]"
    Dim ftrLeft As String:   ftrLeft = "ERP Academie v" & mod_Config.APP_VERSION
    Dim ftrRight As String:  ftrRight = "Page " & Chr(38) & "P" & " / " & Chr(38) & "N"
    
    ' RAPPORTS - Landscape, print titles, dynamic area, page breaks
    Set ws = ThisWorkbook.Sheets("RAPPORTS")
    If Not ws Is Nothing Then
        ws.Unprotect Password:=pwd
        If Err.Number <> 0 Then Err.Clear
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastRow < 2 Then lastRow = 2
        printArea = "A1:" & Chr(64 + lastCol) & lastRow
        ws.PageSetup.PrintArea = printArea
        If Err.Number <> 0 Then Err.Clear
        
        With ws.PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
            .PrintHeadings = False
            .PrintGridlines = False
            .Order = xlDownThenOver
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .LeftHeader = hdrLeft
            .CenterHeader = Format(Date, "DD/MM/YYYY")
            .RightHeader = hdrRight
            .LeftFooter = ftrLeft
            .RightFooter = ftrRight
            .PrintTitleRows = "$1:$5"
            .PrintTitleColumns = ""
        End With
        If Err.Number <> 0 Then Err.Clear
        
        ' Note: Page breaks are managed automatically by Excel via FitToPagesWide/FitToPagesTall
        ' Manual HPageBreaks.Add is extremely slow in COM automation (skipped for build speed)
        
        ws.Protect Password:=pwd, UserInterfaceOnly:=True
        If Err.Number <> 0 Then Err.Clear
    End If
    
    ' INVENTAIRE - Portrait, print titles, dynamic area
    Set ws = ThisWorkbook.Sheets("INVENTAIRE")
    If Not ws Is Nothing Then
        ws.Unprotect Password:=pwd
        If Err.Number <> 0 Then Err.Clear
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastRow < 2 Then lastRow = 2
        printArea = "A1:" & Chr(64 + lastCol) & lastRow
        ws.PageSetup.PrintArea = printArea
        If Err.Number <> 0 Then Err.Clear
        
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
            .CenterHeader = Format(Date, "DD/MM/YYYY")
            .RightHeader = hdrRight
            .LeftFooter = ftrLeft
            .RightFooter = ftrRight
            .PrintTitleRows = "$1:$2"
            .PrintTitleColumns = ""
        End With
        If Err.Number <> 0 Then Err.Clear
        
        ws.Protect Password:=pwd, UserInterfaceOnly:=True
        If Err.Number <> 0 Then Err.Clear
    End If
    
    ' TABLEAU DE BORD - Landscape, print titles, dynamic area
    Set ws = ThisWorkbook.Sheets("TABLEAU DE BORD")
    If Not ws Is Nothing Then
        ws.Unprotect Password:=pwd
        If Err.Number <> 0 Then Err.Clear
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastRow < 2 Then lastRow = 2
        printArea = "A1:" & Chr(64 + lastCol) & lastRow
        ws.PageSetup.PrintArea = printArea
        If Err.Number <> 0 Then Err.Clear
        
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
            .CenterHeader = Format(Date, "DD/MM/YYYY")
            .RightHeader = hdrRight
            .LeftFooter = ftrLeft
            .RightFooter = ftrRight
            .PrintTitleRows = "$1:$1"
            .PrintTitleColumns = ""
        End With
        If Err.Number <> 0 Then Err.Clear
        
        ws.Protect Password:=pwd, UserInterfaceOnly:=True
        If Err.Number <> 0 Then Err.Clear
    End If
    
    ' BON sheets - Portrait, dynamic area
    Dim bonSheets As Variant
    bonSheets = Array("BON_RECEPTION", "BON_SORTIE", "BON_COMMANDE", "DA_DEMANDE_ACHAT")
    Dim idx As Integer
    For idx = LBound(bonSheets) To UBound(bonSheets)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(bonSheets(idx)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Unprotect Password:=pwd
            If Err.Number <> 0 Then Err.Clear
            
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lastRow < 2 Then lastRow = 2
            printArea = "A1:" & Chr(64 + lastCol) & lastRow
            ws.PageSetup.PrintArea = printArea
            If Err.Number <> 0 Then Err.Clear
            
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
                .LeftHeader = hdrLeft
                .RightHeader = hdrRight
                .LeftFooter = ftrLeft
                .RightFooter = ftrRight
            End With
            If Err.Number <> 0 Then Err.Clear
            
            ws.Protect Password:=pwd, UserInterfaceOnly:=True
            If Err.Number <> 0 Then Err.Clear
        End If
    Next idx
End Sub

'================================================================================
' SILENT BARCODE SETUP  -  called during build, no MsgBox
'================================================================================
Private Sub SeedBarcodesSilent()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("STAGING_BUFFER")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "STAGING_BUFFER"
    End If
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0
    
    ' Clear old barcode data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 0 Then ws.Rows("1:" & lastRow).Delete
    
    ' Write header
    ws.Cells(1, 1).Value = "BARCODE_MAP"
    ws.Cells(1, 2).Value = "Default barcode mapping"
    
    ' Write all 15 articles
    Dim i As Long
    For i = 0 To 14
        ws.Cells(2 + i, 1).Value = Format(i + 1, "000")
        ws.Cells(2 + i, 2).Value = "ART-" & Format(i + 1, "000")
    Next i
    
    On Error Resume Next
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

'================================================================================
' HELPERS
'================================================================================

Private Function GetArticlePU(ByVal artCode As String) As Double
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        GetArticlePU = 500
        Exit Function
    End If
    
    Dim foundRow As Variant
    foundRow = Application.Match(artCode, wsArt.Columns(COL_ART_CODE), 0)
    
    If IsError(foundRow) Then
        GetArticlePU = 500
    Else
        GetArticlePU = wsArt.Cells(foundRow, COL_ART_PU).Value
    End If
End Function

Private Function GetArticleSupplier(ByVal artCode As String) As String
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        GetArticleSupplier = "F-001"
        Exit Function
    End If
    
    Dim foundRow As Variant
    foundRow = Application.Match(artCode, wsArt.Columns(COL_ART_CODE), 0)
    
    If IsError(foundRow) Then
        GetArticleSupplier = "F-001"
    Else
        GetArticleSupplier = wsArt.Cells(foundRow, COL_ART_FOURNISSEUR).Value
    End If
End Function

Private Function RandomService() As String
    Dim services As Variant
    services = Array("Service Comptabilite", "Service Archives", "Service Informatique", "Direction", "Service Juridique")
    Randomize
    RandomService = services(Int(Rnd * 5))
End Function

'================================================================================
' BUILD FINALIZE - Re-protect all sheets (build-time only, no UserInterfaceOnly)
' Called by build.ps1 after GenerateDemoData
'================================================================================

Public Sub FinalizeBuildProtection()
    Dim ws As Worksheet
    Dim pwd As String
    pwd = mod_Config.MASTER_PWD
    
    For Each ws In ThisWorkbook.Sheets
        If ws.ProtectContents Then ws.Unprotect Password:=pwd
        ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                   AllowFormattingCells:=False, AllowFormattingColumns:=False, _
                   AllowFormattingRows:=False, AllowInsertingColumns:=False, _
                   AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
                   AllowDeletingColumns:=False, AllowDeletingRows:=False, _
                   AllowSorting:=False, AllowFiltering:=True, _
                   AllowUsingPivotTables:=False
    Next ws
End Sub

'================================================================================
' END -- mod_DemoData.bas
'================================================================================

