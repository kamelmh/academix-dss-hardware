Attribute VB_Name = "mod_SupplierRegistry"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Hardware Store Supplier Registry - 9 suppliers
' ============================================================================

Option Explicit

'================================================================================
' SUPPLIER STRUCTURE
'================================================================================

Public Type SupplierInfo
    Code            As String
    Name            As String
    Address         As String
    Phone           As String
    NIF             As String
    NIS             As String
    RC              As String
    ArticleImpot    As String
    Category        As String
    IsActive        As Boolean
    Rating          As Double
End Type

'================================================================================
' SUPPLIER DATABASE - Hardware Store Suppliers
'================================================================================

Public Function GetSupplierInfo(ByVal supplierCode As String) As SupplierInfo
    Dim s As SupplierInfo
    
    Select Case UCase(Trim(supplierCode))
        Case "SIDERAL", "FER-001"
            s.Code = "SIDERAL"
            s.Name = "SIDERAL SPA - Distribution Acier"
            s.Address = "Zone Industrielle"
            s.Phone = "049 00 00 01"
            s.NIF = "000116010002500"
            s.NIS = "0161600100250"
            s.RC = "32/00-0012345B67"
            s.ArticleImpot = "250"
            s.Category = "Acier et metaux"
            s.IsActive = True
            s.Rating = 4.5
            
        Case "CIMENTAL", "CIM-001"
            s.Code = "CIMENTAL"
            s.Name = "CIMENTAL SPA - Ciment Algerie"
            s.Address = "Bab Ezzouar, Alger"
            s.Phone = "023 00 00 02"
            s.NIF = "000231010003400"
            s.NIS = "0313100100340"
            s.RC = "16/00-0023456B12"
            s.ArticleImpot = "340"
            s.Category = "Ciment"
            s.IsActive = True
            s.Rating = 4.8
            
        Case "GRANULATS", "SAB-001"
            s.Code = "GRANULATS"
            s.Name = "GRANULATS SA - Sable et Gravier"
            s.Address = "Tipaza, Alger"
            s.Phone = "024 00 00 03"
            s.NIF = "000416010004500"
            s.NIS = "0161600100450"
            s.RC = "42/00-0034567B89"
            s.ArticleImpot = "450"
            s.Category = "Granulats"
            s.IsActive = True
            s.Rating = 4.3
            
        Case "PLASTIQUE", "PVC-001"
            s.Code = "PLASTIQUE"
            s.Name = "PLASTIQUE PRO - Tuyauterie PVC"
            s.Address = "Es Senia, Oran"
            s.Phone = "041 00 00 04"
            s.NIF = "000531010005600"
            s.NIS = "0313100100560"
            s.RC = "31/00-0045678B90"
            s.ArticleImpot = "560"
            s.Category = "Plomberie"
            s.IsActive = True
            s.Rating = 4.1
            
        Case "ELECTRO", "CAB-001"
            s.Code = "ELECTRO"
            s.Name = "ELECTRO PLUS - Cablage"
            s.Address = "Constantine"
            s.Phone = "031 00 00 05"
            s.NIF = "000632010006700"
            s.NIS = "0323200100670"
            s.RC = "25/00-0056789B01"
            s.ArticleImpot = "670"
            s.Category = "Electricite"
            s.IsActive = True
            s.Rating = 4.4
            
        Case "DISTRIBUTION", "DIS-001"
            s.Code = "DISTRIBUTION"
            s.Name = "DISTRIBUTION ELECTRIQUE SPA"
            s.Address = "Setif"
            s.Phone = "036 00 00 06"
            s.NIF = "000719010007800"
            s.NIS = "0191900100780"
            s.RC = "19/00-0067890B12"
            s.ArticleImpot = "780"
            s.Category = "Electricite"
            s.IsActive = True
            s.Rating = 4.0
            
        Case "OUTILMAG", "OUT-001"
            s.Code = "OUTILMAG"
            s.Name = "OUTILMAG - Materiaux de Batiment"
            s.Address = "Bab Ezzouar, Alger"
            s.Phone = "023 00 00 07"
            s.NIF = "000816010008900"
            s.NIS = "0161600100890"
            s.RC = "16/00-0078901B23"
            s.ArticleImpot = "890"
            s.Category = "Outillage"
            s.IsActive = True
            s.Rating = 4.2
            
        Case "PEINTURE", "PEI-001"
            s.Code = "PEINTURE"
            s.Name = "PEINTURE PLUS - Enduits et Peintures"
            s.Address = "Blida"
            s.Phone = "025 00 00 08"
            s.NIF = "000909010009000"
            s.NIS = "0090900100900"
            s.RC = "09/00-0089012B34"
            s.ArticleImpot = "900"
            s.Category = "Peinture"
            s.IsActive = True
            s.Rating = 4.6
            
        Case "CERAMIQUE", "CAR-001"
            s.Code = "CERAMIQUE"
            s.Name = "CERAMIQUE EL BAYADH"
            s.Address = "Centre"
            s.Phone = "049 00 00 09"
            s.NIF = "001032010010100"
            s.NIS = "0323200101010"
            s.RC = "32/00-0090123B45"
            s.ArticleImpot = "101"
            s.Category = "Carrelage"
            s.IsActive = True
            s.Rating = 4.7
            
        Case Else
            s.Code = supplierCode
            s.IsActive = False
    End Select
    
    GetSupplierInfo = s
End Function

'================================================================================
' INDIVIDUAL FIELD LOOKUPS
'================================================================================

Public Function GetSupplierNIF(ByVal supplierCode As String) As String
    GetSupplierNIF = GetSupplierInfo(supplierCode).NIF
End Function

Public Function GetSupplierNIS(ByVal supplierCode As String) As String
    GetSupplierNIS = GetSupplierInfo(supplierCode).NIS
End Function

Public Function GetSupplierRC(ByVal supplierCode As String) As String
    GetSupplierRC = GetSupplierInfo(supplierCode).RC
End Function

Public Function GetSupplierArticle(ByVal supplierCode As String) As String
    GetSupplierArticle = GetSupplierInfo(supplierCode).ArticleImpot
End Function

Public Function GetSupplierName(ByVal supplierCode As String) As String
    GetSupplierName = GetSupplierInfo(supplierCode).Name
End Function

Public Function GetSupplierAddress(ByVal supplierCode As String) As String
    GetSupplierAddress = GetSupplierInfo(supplierCode).Address
End Function

Public Function GetSupplierPhone(ByVal supplierCode As String) As String
    GetSupplierPhone = GetSupplierInfo(supplierCode).Phone
End Function

'================================================================================
' VALIDATION
'================================================================================

Public Function IsSupplierValid(ByVal supplierCode As String) As Boolean
    Dim s As SupplierInfo
    s = GetSupplierInfo(supplierCode)
    IsSupplierValid = s.IsActive And Len(s.NIF) > 0 And Len(s.NIS) > 0
End Function

Public Function ValidateSupplier(ByVal supplierCode As String) As String
    Dim s As SupplierInfo
    s = GetSupplierInfo(supplierCode)
    
    If Not s.IsActive Then
        ValidateSupplier = "Fournisseur '" & supplierCode & "' non actif"
    ElseIf Len(s.NIF) = 0 Then
        ValidateSupplier = "NIF manquant pour " & supplierCode
    ElseIf Len(s.NIS) = 0 Then
        ValidateSupplier = "NIS manquant pour " & supplierCode
    ElseIf Len(s.RC) = 0 Then
        ValidateSupplier = "RC manquant pour " & supplierCode
    Else
        ValidateSupplier = ""
    End If
End Function

'================================================================================
' FOURNISSEURS SHEET - Setup
'================================================================================

Public Sub SetupFournisseursSheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim s As Worksheet
    
    sheetExists = False
    For Each s In ThisWorkbook.Sheets
        If s.Name = mod_Config.SHEET_FOURNISSEURS Then
            sheetExists = True
            Set ws = s
            Exit For
        End If
    Next s
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = mod_Config.SHEET_FOURNISSEURS
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Cells.Interior.ColorIndex = xlNone
    
    ws.Range("A1:I1").Merge
    ws.Cells(1, 1).Value = "CATALOGUE FOURNISSEURS - " & mod_Config.BUSINESS_NAME
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(1, 1).Font.Color = RGB(0, 70, 127)
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Rows(1).RowHeight = 25
    
    Dim headers As Variant
    headers = Array("Code", "Raison Sociale", "Adresse", "Telephone", _
                    "NIF (15 chiffres)", "NIS (13 chiffres)", "Registre Commerce", _
                    "Art. Imposition", "Categorie")
    
    Dim i As Integer
    For i = 0 To UBound(headers)
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Size = 9
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 70, 127)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
    Next i
    ws.Rows(2).RowHeight = 30
    
    Dim suppliers As Variant
    suppliers = Array("SIDERAL", "CIMENTAL", "GRANULATS", "PLASTIQUE", _
                      "ELECTRO", "DISTRIBUTION", "OUTILMAG", "PEINTURE", "CERAMIQUE")
    
    Dim r As Long
    For r = 0 To 8
        Dim sInfo As SupplierInfo
        sInfo = GetSupplierInfo(suppliers(r))
        
        Dim dataRow As Long
        dataRow = r + 3
        
        ws.Cells(dataRow, 1).Value = sInfo.Code
        ws.Cells(dataRow, 2).Value = sInfo.Name
        ws.Cells(dataRow, 3).Value = sInfo.Address
        ws.Cells(dataRow, 4).Value = sInfo.Phone
        ws.Cells(dataRow, 5).Value = sInfo.NIF
        ws.Cells(dataRow, 6).Value = sInfo.NIS
        ws.Cells(dataRow, 7).Value = sInfo.RC
        ws.Cells(dataRow, 8).Value = sInfo.ArticleImpot
        ws.Cells(dataRow, 9).Value = sInfo.Category
        
        Dim c As Integer
        For c = 1 To 9
            With ws.Cells(dataRow, c)
                .Font.Size = 9
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With
        Next c
        
        ws.Cells(dataRow, 1).HorizontalAlignment = xlCenter
        ws.Cells(dataRow, 1).Font.Name = "Courier New"
        ws.Cells(dataRow, 5).Font.Name = "Courier New"
        ws.Cells(dataRow, 6).Font.Name = "Courier New"
        ws.Cells(dataRow, 7).Font.Name = "Courier New"
        ws.Cells(dataRow, 8).Font.Name = "Courier New"
        
        If r Mod 2 = 0 Then
            ws.Range("A" & dataRow & ":I" & dataRow).Interior.Color = RGB(245, 245, 255)
        End If
        
        ws.Rows(dataRow).RowHeight = 24
    Next r
    
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 38
    ws.Columns("C").ColumnWidth = 32
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 22
    
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    Debug.Print "[SupplierRegistry] 9 hardware suppliers loaded"
End Sub

'================================================================================
' PDF EXPORT INTEGRATION
'================================================================================

Public Function GetSupplierTaxIDsForPDF(ByVal supplierCode As String) As String
    Dim s As SupplierInfo
    s = GetSupplierInfo(supplierCode)
    
    If s.IsActive And Len(s.NIF) > 0 Then
        GetSupplierTaxIDsForPDF = "NIF : " & s.NIF & "  |  NIS : " & s.NIS & _
                                  "  |  RC : " & s.RC & "  |  Art. : " & s.ArticleImpot
    Else
        GetSupplierTaxIDsForPDF = "NIF : ____________________  |  NIS : ____________________" & _
                                  "  |  RC : ____________________  |  Art. : ____________________"
    End If
End Function

Public Function GetSupplierLegalName(ByVal supplierCode As String) As String
    Dim s As SupplierInfo
    s = GetSupplierInfo(supplierCode)
    If s.IsActive Then
        GetSupplierLegalName = s.Name
    Else
        GetSupplierLegalName = supplierCode
    End If
End Function

Public Function GetAllSupplierCodes() As Variant
    GetAllSupplierCodes = Array("SIDERAL", "CIMENTAL", "GRANULATS", "PLASTIQUE", _
                                "ELECTRO", "DISTRIBUTION", "OUTILMAG", "PEINTURE", "CERAMIQUE")
End Function

Public Function GetAllSupplierNames() As Variant
    GetAllSupplierNames = Array( _
        "SIDERAL - Acier et metaux", _
        "CIMENTAL - Ciment", _
        "GRANULATS - Sable et Gravier", _
        "PLASTIQUE - Tuyauterie PVC", _
        "ELECTRO - Cablage", _
        "DISTRIBUTION - Electricite", _
        "OUTILMAG - Outillage", _
        "PEINTURE - Peinture", _
        "CERAMIQUE - Carrelage" _
    )
End Function

'================================================================================
' SUPPLIER PERFORMANCE
'================================================================================

Public Function GetSupplierRating(ByVal supplierCode As String) As Double
    GetSupplierRating = GetSupplierInfo(supplierCode).Rating
End Function

Public Function GetSupplierCategory(ByVal supplierCode As String) As String
    GetSupplierCategory = GetSupplierInfo(supplierCode).Category
End Function
