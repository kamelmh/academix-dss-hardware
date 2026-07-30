Attribute VB_Name = "mod_DemoData"
Option Explicit

Public Sub GenerateDemoData()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrHandler
    Call NuclearClear
    Call SeedConfig
    Call SeedArticles40
    Call SeedSuppliers9
    Call SeedMovements90
    Call UpdateStockFromMovements
    Call mod_StockEngine.UpdateAllABCClassifications(silent:=True)
    Call mod_StockEngine.RefreshAllCMUP
    Call AutoFitAllData
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Done! 40 articles, 9 suppliers, 90 days." & vbCrLf & "ABC + CMUP calculated.", vbInformation, "DSS v14"
    Exit Sub
ErrHandler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur: " & Err.Description, vbCritical
End Sub

Private Sub NuclearClear()
    Dim ws As Worksheet
    On Error Resume Next
    
    ' ARTICLES - 14 columns matching mod_Config constants
    Set ws = ThisWorkbook.Sheets("ARTICLES")
    If Not ws Is Nothing Then
        ws.Unprotect Password:=mod_Config.MASTER_PWD
        ws.Cells.Clear
        ws.Cells.MergeCells = False
        ' Row 1: Headers
        ws.Cells(1, 1).Value = "Code"
        ws.Cells(1, 2).Value = "Designation"
        ws.Cells(1, 3).Value = "Stock"
        ws.Cells(1, 4).Value = "Seuil Min"
        ws.Cells(1, 5).Value = "Categorie"
        ws.Cells(1, 6).Value = "Classe ABC"
        ws.Cells(1, 7).Value = "Stock Actuel"
        ws.Cells(1, 8).Value = "PU"
        ws.Cells(1, 9).Value = "Fournisseur"
        ws.Cells(1, 10).Value = "Stock Securite"
        ws.Cells(1, 11).Value = "Notes"
        ws.Cells(1, 12).Value = "CMUP"
        ws.Cells(1, 13).Value = "Methode Appro"
        ws.Cells(1, 14).Value = "Delai"
        ws.Range("A1:N1").Font.Bold = True
        ws.Range("A1:N1").Interior.Color = RGB(0, 70, 127)
        ws.Range("A1:N1").Font.Color = RGB(255, 255, 255)
        ws.Columns("A:N").AutoFit
    End If
    Set ws = Nothing
    
    ' MOUVEMENTS - 12 columns matching mod_Config constants (with audit trail)
    Set ws = ThisWorkbook.Sheets("MOUVEMENTS")
    If Not ws Is Nothing Then
        ws.Unprotect Password:=mod_Config.MASTER_PWD
        ws.Cells.Clear
        ws.Cells.MergeCells = False
        ws.Cells(1, 1).Value = "Date"
        ws.Cells(1, 2).Value = "Code Article"
        ws.Cells(1, 3).Value = "Designation"
        ws.Cells(1, 4).Value = "Type"
        ws.Cells(1, 5).Value = "Quantite"
        ws.Cells(1, 6).Value = "Valeur"
        ws.Cells(1, 7).Value = "Ref Document"
        ws.Cells(1, 8).Value = "PU"
        ws.Cells(1, 9).Value = "Tiers"
        ws.Cells(1, 10).Value = "Notes"
        ws.Cells(1, 11).Value = "Utilisateur"
        ws.Cells(1, 12).Value = "Horodatage"
        ws.Range("A1:L1").Font.Bold = True
        ws.Range("A1:L1").Interior.Color = RGB(0, 70, 127)
        ws.Range("A1:L1").Font.Color = RGB(255, 255, 255)
        ws.Columns("A:L").AutoFit
    End If
    Set ws = Nothing
    
    ' FOURNISSEURS - 8 columns matching mod_Config constants
    Set ws = ThisWorkbook.Sheets("FOURNISSEURS")
    If Not ws Is Nothing Then
        ws.Unprotect Password:=mod_Config.MASTER_PWD
        ws.Cells.Clear
        ws.Cells.MergeCells = False
        ws.Cells(1, 1).Value = "Code"
        ws.Cells(1, 2).Value = "Raison Sociale"
        ws.Cells(1, 3).Value = "Adresse"
        ws.Cells(1, 4).Value = "Telephone"
        ws.Cells(1, 5).Value = "NIF"
        ws.Cells(1, 6).Value = "NIS"
        ws.Cells(1, 7).Value = "RC"
        ws.Cells(1, 8).Value = "Art. Impot"
        ws.Range("A1:H1").Font.Bold = True
        ws.Range("A1:H1").Interior.Color = RGB(0, 70, 127)
        ws.Range("A1:H1").Font.Color = RGB(255, 255, 255)
        ws.Columns("A:H").AutoFit
    End If
    Set ws = Nothing
    
    ' CONFIG - clean slate for SeedConfig
    Set ws = ThisWorkbook.Sheets("CONFIG")
    If Not ws Is Nothing Then
        ws.Unprotect Password:=mod_Config.MASTER_PWD
        ws.Cells.Clear
        ws.Cells(1, 1).Value = "Parameter"
        ws.Cells(1, 2).Value = "Value"
        ws.Cells(1, 3).Value = "Description"
        ws.Range("A1:C1").Font.Bold = True
        ws.Range("A1:C1").Interior.Color = RGB(0, 70, 127)
        ws.Range("A1:C1").Font.Color = RGB(255, 255, 255)
        ws.Columns("A:C").AutoFit
    End If
    Set ws = Nothing
    
    ' === FEATURE SHEETS (create if missing) ===
    ' FACTURES
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FACTURES")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "FACTURES"
    End If
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Numero"
    ws.Cells(1, 2).Value = "Date"
    ws.Cells(1, 3).Value = "Client"
    ws.Cells(1, 4).Value = "Adresse"
    ws.Cells(1, 5).Value = "NIF"
    ws.Cells(1, 6).Value = "NIS"
    ws.Cells(1, 7).Value = "RC"
    ws.Cells(1, 8).Value = "Total HT"
    ws.Cells(1, 9).Value = "Total TVA"
    ws.Cells(1, 10).Value = "Total TTC"
    ws.Cells(1, 11).Value = "Reglement"
    ws.Cells(1, 12).Value = "Notes"
    ws.Cells(1, 13).Value = "Utilisateur"
    ws.Cells(1, 14).Value = "Horodatage"
    ws.Range("A1:N1").Font.Bold = True
    ws.Range("A1:N1").Interior.Color = RGB(0, 70, 127)
    ws.Range("A1:N1").Font.Color = RGB(255, 255, 255)
    ws.Columns("A:N").AutoFit
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    Set ws = Nothing
    
    ' BARCODES
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("BARCODES")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "BARCODES"
    End If
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Barcode"
    ws.Cells(1, 2).Value = "Code Article"
    ws.Cells(1, 3).Value = "Designation"
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(0, 70, 127)
    ws.Range("A1:C1").Font.Color = RGB(255, 255, 255)
    ws.Columns("A:C").AutoFit
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    Set ws = Nothing
    
    ' BONS_COMMANDE
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("BONS_COMMANDE")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "BONS_COMMANDE"
    End If
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Numero"
    ws.Cells(1, 2).Value = "Date"
    ws.Cells(1, 3).Value = "Fournisseur"
    ws.Cells(1, 4).Value = "Code Article"
    ws.Cells(1, 5).Value = "Designation"
    ws.Cells(1, 6).Value = "Quantite"
    ws.Cells(1, 7).Value = "PU"
    ws.Cells(1, 8).Value = "Valeur"
    ws.Cells(1, 9).Value = "Statut"
    ws.Cells(1, 10).Value = "Utilisateur"
    ws.Cells(1, 11).Value = "Horodatage"
    ws.Range("A1:K1").Font.Bold = True
    ws.Range("A1:K1").Interior.Color = RGB(0, 70, 127)
    ws.Range("A1:K1").Font.Color = RGB(255, 255, 255)
    ws.Columns("A:K").AutoFit
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    Set ws = Nothing
    
    On Error GoTo 0
End Sub

Private Sub SeedConfig()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CONFIG")
    Dim r As Long
    r = 2
    ws.Cells(r, 1) = "WORKING_DAYS": ws.Cells(r, 2) = 300: ws.Cells(r, 3) = "Jours ouverts/an": r = r + 1
    ws.Cells(r, 1) = "OBSERVATION_DAYS": ws.Cells(r, 2) = 90: ws.Cells(r, 3) = "Jours d'observation": r = r + 1
    ws.Cells(r, 1) = "ORDER_COST": ws.Cells(r, 2) = 300: ws.Cells(r, 3) = "Cout commande (DZD)": r = r + 1
    ws.Cells(r, 1) = "HOLDING_RATE": ws.Cells(r, 2) = 0.2: ws.Cells(r, 3) = "Taux de stockage (20%)": r = r + 1
    ws.Cells(r, 1) = "LEAD_TIME": ws.Cells(r, 2) = 2: ws.Cells(r, 3) = "Delai livraison (jours)": r = r + 1
    ws.Cells(r, 1) = "TAX_RATE": ws.Cells(r, 2) = 0.19: ws.Cells(r, 3) = "TVA (19%)": r = r + 1
    ws.Cells(r, 1) = "CURRENCY": ws.Cells(r, 2) = "DZD": ws.Cells(r, 3) = "Devise": r = r + 1
    ws.Cells(r, 1) = "BUSINESS_NAME": ws.Cells(r, 2) = "Quincaillerie": ws.Cells(r, 3) = "Nom commercial": r = r + 1
    ws.Cells(r, 1) = "BUSINESS_ADDRESS": ws.Cells(r, 2) = "Algerie": ws.Cells(r, 3) = "Adresse": r = r + 1
    ws.Cells(r, 1) = "BUSINESS_PHONE": ws.Cells(r, 2) = "049 00 00 00": ws.Cells(r, 3) = "Telephone": r = r + 1
    ws.Cells(r, 1) = "BUSINESS_NIF": ws.Cells(r, 2) = "000100000000000": ws.Cells(r, 3) = "NIF": r = r + 1
    ws.Cells(r, 1) = "BUSINESS_NIS": ws.Cells(r, 2) = "00100000000000": ws.Cells(r, 3) = "NIS": r = r + 1
    ws.Cells(r, 1) = "BUSINESS_RC": ws.Cells(r, 2) = "00/00-0000000A00": ws.Cells(r, 3) = "RC": r = r + 1
    ws.Cells(r, 1) = "SEASON": ws.Cells(r, 2) = "Printemps": ws.Cells(r, 3) = "Saison": r = r + 1
    Debug.Print "[DemoData] Config seeded"
End Sub

Private Sub SeedArticles40()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ARTICLES")
    Dim r As Long
    r = 2
    ' Columns: 1=Code, 2=Designation, 3=Stock, 4=SeuilMin, 5=Categorie, 6=ClasseABC, 7=StockActuel, 8=PU, 9=Fournisseur, 10=StockSecurite, 11=Notes
    ws.Cells(r, 1) = "FER-001": ws.Cells(r, 2) = "Fer a beton 10mm 12m": ws.Cells(r, 3) = 2500: ws.Cells(r, 4) = 200: ws.Cells(r, 5) = "Fer": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 2500: ws.Cells(r, 8) = 4000: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 12: ws.Cells(r, 11) = "Acier": r = r + 1
    ws.Cells(r, 1) = "FER-002": ws.Cells(r, 2) = "Fer a beton 12mm 12m": ws.Cells(r, 3) = 1800: ws.Cells(r, 4) = 150: ws.Cells(r, 5) = "Fer": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 1800: ws.Cells(r, 8) = 6000: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 10: ws.Cells(r, 11) = "Acier": r = r + 1
    ws.Cells(r, 1) = "FER-003": ws.Cells(r, 2) = "Fer a beton 14mm 12m": ws.Cells(r, 3) = 1200: ws.Cells(r, 4) = 100: ws.Cells(r, 5) = "Fer": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 1200: ws.Cells(r, 8) = 8000: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 8: ws.Cells(r, 11) = "Acier": r = r + 1
    ws.Cells(r, 1) = "FER-004": ws.Cells(r, 2) = "Fer a beton 16mm 12m": ws.Cells(r, 3) = 800: ws.Cells(r, 4) = 80: ws.Cells(r, 5) = "Fer": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 800: ws.Cells(r, 8) = 10500: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 6: ws.Cells(r, 11) = "Acier": r = r + 1
    ws.Cells(r, 1) = "ACI-001": ws.Cells(r, 2) = "Acier filete 12mm 6m": ws.Cells(r, 3) = 1500: ws.Cells(r, 4) = 120: ws.Cells(r, 5) = "Fer": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 1500: ws.Cells(r, 8) = 3500: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 10: ws.Cells(r, 11) = "Acier": r = r + 1
    ws.Cells(r, 1) = "ACI-002": ws.Cells(r, 2) = "Acier filete 16mm 6m": ws.Cells(r, 3) = 1000: ws.Cells(r, 4) = 90: ws.Cells(r, 5) = "Fer": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 1000: ws.Cells(r, 8) = 5200: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 8: ws.Cells(r, 11) = "Acier": r = r + 1
    ws.Cells(r, 1) = "TOLE-001": ws.Cells(r, 2) = "Tole acier 2mm 1x2m": ws.Cells(r, 3) = 800: ws.Cells(r, 4) = 60: ws.Cells(r, 5) = "Fer": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 800: ws.Cells(r, 8) = 4800: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 6: ws.Cells(r, 11) = "Acier": r = r + 1
    ws.Cells(r, 1) = "TOLE-002": ws.Cells(r, 2) = "Tole acier 3mm 1x2m": ws.Cells(r, 3) = 500: ws.Cells(r, 4) = 40: ws.Cells(r, 5) = "Fer": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 500: ws.Cells(r, 8) = 7200: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 4: ws.Cells(r, 11) = "Acier": r = r + 1
    ws.Cells(r, 1) = "CIM-001": ws.Cells(r, 2) = "Ciment CPA 55 35kg": ws.Cells(r, 3) = 500: ws.Cells(r, 4) = 100: ws.Cells(r, 5) = "Ciment": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 500: ws.Cells(r, 8) = 550: ws.Cells(r, 9) = "CIMENTAL": ws.Cells(r, 10) = 100: ws.Cells(r, 11) = "Ciment": r = r + 1
    ws.Cells(r, 1) = "CIM-002": ws.Cells(r, 2) = "Ciment CPA 45 35kg": ws.Cells(r, 3) = 400: ws.Cells(r, 4) = 80: ws.Cells(r, 5) = "Ciment": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 400: ws.Cells(r, 8) = 480: ws.Cells(r, 9) = "CIMENTAL": ws.Cells(r, 10) = 80: ws.Cells(r, 11) = "Ciment": r = r + 1
    ws.Cells(r, 1) = "CIM-003": ws.Cells(r, 2) = "Ciment CRC 35kg": ws.Cells(r, 3) = 300: ws.Cells(r, 4) = 60: ws.Cells(r, 5) = "Ciment": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 300: ws.Cells(r, 8) = 620: ws.Cells(r, 9) = "CIMENTAL": ws.Cells(r, 10) = 60: ws.Cells(r, 11) = "Ciment": r = r + 1
    ws.Cells(r, 1) = "SAB-001": ws.Cells(r, 2) = "Sable fin 1 tonne": ws.Cells(r, 3) = 50: ws.Cells(r, 4) = 15: ws.Cells(r, 5) = "Granulat": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 50: ws.Cells(r, 8) = 3500: ws.Cells(r, 9) = "GRANULATS": ws.Cells(r, 10) = 15: ws.Cells(r, 11) = "Granulat": r = r + 1
    ws.Cells(r, 1) = "SAB-002": ws.Cells(r, 2) = "Sable concasse 1 tonne": ws.Cells(r, 3) = 40: ws.Cells(r, 4) = 12: ws.Cells(r, 5) = "Granulat": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 40: ws.Cells(r, 8) = 4200: ws.Cells(r, 9) = "GRANULATS": ws.Cells(r, 10) = 12: ws.Cells(r, 11) = "Granulat": r = r + 1
    ws.Cells(r, 1) = "GRA-001": ws.Cells(r, 2) = "Gravier 5-25mm 1 tonne": ws.Cells(r, 3) = 35: ws.Cells(r, 4) = 10: ws.Cells(r, 5) = "Granulat": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 35: ws.Cells(r, 8) = 4800: ws.Cells(r, 9) = "GRANULATS": ws.Cells(r, 10) = 10: ws.Cells(r, 11) = "Granulat": r = r + 1
    ws.Cells(r, 1) = "GRA-002": ws.Cells(r, 2) = "Gravier 0-5mm 1 tonne": ws.Cells(r, 3) = 30: ws.Cells(r, 4) = 10: ws.Cells(r, 5) = "Granulat": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 30: ws.Cells(r, 8) = 3800: ws.Cells(r, 9) = "GRANULATS": ws.Cells(r, 10) = 10: ws.Cells(r, 11) = "Granulat": r = r + 1
    ws.Cells(r, 1) = "PVC-001": ws.Cells(r, 2) = "Tuyau PVC 32mm 6m": ws.Cells(r, 3) = 150: ws.Cells(r, 4) = 30: ws.Cells(r, 5) = "Plomberie": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 150: ws.Cells(r, 8) = 450: ws.Cells(r, 9) = "PLASTIQUE": ws.Cells(r, 10) = 30: ws.Cells(r, 11) = "PVC": r = r + 1
    ws.Cells(r, 1) = "PVC-002": ws.Cells(r, 2) = "Tuyau PVC 50mm 6m": ws.Cells(r, 3) = 120: ws.Cells(r, 4) = 25: ws.Cells(r, 5) = "Plomberie": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 120: ws.Cells(r, 8) = 780: ws.Cells(r, 9) = "PLASTIQUE": ws.Cells(r, 10) = 25: ws.Cells(r, 11) = "PVC": r = r + 1
    ws.Cells(r, 1) = "PVC-003": ws.Cells(r, 2) = "Tuyau PVC 110mm 6m": ws.Cells(r, 3) = 80: ws.Cells(r, 4) = 20: ws.Cells(r, 5) = "Plomberie": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 80: ws.Cells(r, 8) = 1800: ws.Cells(r, 9) = "PLASTIQUE": ws.Cells(r, 10) = 20: ws.Cells(r, 11) = "PVC": r = r + 1
    ws.Cells(r, 1) = "PVC-004": ws.Cells(r, 2) = "Tuyau PVC 160mm 6m": ws.Cells(r, 3) = 40: ws.Cells(r, 4) = 10: ws.Cells(r, 5) = "Plomberie": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 40: ws.Cells(r, 8) = 3200: ws.Cells(r, 9) = "PLASTIQUE": ws.Cells(r, 10) = 10: ws.Cells(r, 11) = "PVC": r = r + 1
    ws.Cells(r, 1) = "FER-010": ws.Cells(r, 2) = "Tuyau galva 25mm 6m": ws.Cells(r, 3) = 60: ws.Cells(r, 4) = 15: ws.Cells(r, 5) = "Plomberie": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 60: ws.Cells(r, 8) = 2200: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 15: ws.Cells(r, 11) = "Metal": r = r + 1
    ws.Cells(r, 1) = "FER-011": ws.Cells(r, 2) = "Tuyau galva 32mm 6m": ws.Cells(r, 3) = 50: ws.Cells(r, 4) = 12: ws.Cells(r, 5) = "Plomberie": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 50: ws.Cells(r, 8) = 3500: ws.Cells(r, 9) = "SIDERAL": ws.Cells(r, 10) = 12: ws.Cells(r, 11) = "Metal": r = r + 1
    ws.Cells(r, 1) = "RAC-001": ws.Cells(r, 2) = "Raccords PVC lot 100": ws.Cells(r, 3) = 100: ws.Cells(r, 4) = 20: ws.Cells(r, 5) = "Plomberie": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 100: ws.Cells(r, 8) = 850: ws.Cells(r, 9) = "PLASTIQUE": ws.Cells(r, 10) = 20: ws.Cells(r, 11) = "PVC": r = r + 1
    ws.Cells(r, 1) = "ROB-001": ws.Cells(r, 2) = "Robinetterie lot 10": ws.Cells(r, 3) = 50: ws.Cells(r, 4) = 10: ws.Cells(r, 5) = "Plomberie": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 50: ws.Cells(r, 8) = 2500: ws.Cells(r, 9) = "PLASTIQUE": ws.Cells(r, 10) = 10: ws.Cells(r, 11) = "PVC": r = r + 1
    ws.Cells(r, 1) = "CAB-001": ws.Cells(r, 2) = "Cable 2.5mm2 100m": ws.Cells(r, 3) = 80: ws.Cells(r, 4) = 20: ws.Cells(r, 5) = "Electricite": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 80: ws.Cells(r, 8) = 3800: ws.Cells(r, 9) = "ELECTRO": ws.Cells(r, 10) = 20: ws.Cells(r, 11) = "Cable": r = r + 1
    ws.Cells(r, 1) = "CAB-002": ws.Cells(r, 2) = "Cable 4mm2 100m": ws.Cells(r, 3) = 60: ws.Cells(r, 4) = 15: ws.Cells(r, 5) = "Electricite": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 60: ws.Cells(r, 8) = 5500: ws.Cells(r, 9) = "ELECTRO": ws.Cells(r, 10) = 15: ws.Cells(r, 11) = "Cable": r = r + 1
    ws.Cells(r, 1) = "CAB-003": ws.Cells(r, 2) = "Cable 6mm2 100m": ws.Cells(r, 3) = 40: ws.Cells(r, 4) = 10: ws.Cells(r, 5) = "Electricite": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 40: ws.Cells(r, 8) = 8200: ws.Cells(r, 9) = "ELECTRO": ws.Cells(r, 10) = 10: ws.Cells(r, 11) = "Cable": r = r + 1
    ws.Cells(r, 1) = "DIS-001": ws.Cells(r, 2) = "Disjoncteur 20A lot 10": ws.Cells(r, 3) = 50: ws.Cells(r, 4) = 12: ws.Cells(r, 5) = "Electricite": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 50: ws.Cells(r, 8) = 4500: ws.Cells(r, 9) = "DISTRIBUTION": ws.Cells(r, 10) = 12: ws.Cells(r, 11) = "Electrique": r = r + 1
    ws.Cells(r, 1) = "DIS-002": ws.Cells(r, 2) = "Disjoncteur 32A lot 10": ws.Cells(r, 3) = 30: ws.Cells(r, 4) = 8: ws.Cells(r, 5) = "Electricite": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 30: ws.Cells(r, 8) = 6800: ws.Cells(r, 9) = "DISTRIBUTION": ws.Cells(r, 10) = 8: ws.Cells(r, 11) = "Electrique": r = r + 1
    ws.Cells(r, 1) = "PRI-001": ws.Cells(r, 2) = "Prise murale lot 20": ws.Cells(r, 3) = 100: ws.Cells(r, 4) = 25: ws.Cells(r, 5) = "Electricite": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 100: ws.Cells(r, 8) = 2200: ws.Cells(r, 9) = "DISTRIBUTION": ws.Cells(r, 10) = 25: ws.Cells(r, 11) = "Electrique": r = r + 1
    ws.Cells(r, 1) = "INT-001": ws.Cells(r, 2) = "Interrupteur lot 20": ws.Cells(r, 3) = 100: ws.Cells(r, 4) = 25: ws.Cells(r, 5) = "Electricite": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 100: ws.Cells(r, 8) = 1800: ws.Cells(r, 9) = "DISTRIBUTION": ws.Cells(r, 10) = 25: ws.Cells(r, 11) = "Electrique": r = r + 1
    ws.Cells(r, 1) = "OUT-001": ws.Cells(r, 2) = "Perceuse Makita": ws.Cells(r, 3) = 10: ws.Cells(r, 4) = 3: ws.Cells(r, 5) = "Outillage": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 10: ws.Cells(r, 8) = 15000: ws.Cells(r, 9) = "OUTILMAG": ws.Cells(r, 10) = 3: ws.Cells(r, 11) = "Outillage": r = r + 1
    ws.Cells(r, 1) = "OUT-002": ws.Cells(r, 2) = "Meuleuse 125mm": ws.Cells(r, 3) = 15: ws.Cells(r, 4) = 5: ws.Cells(r, 5) = "Outillage": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 15: ws.Cells(r, 8) = 8500: ws.Cells(r, 9) = "OUTILMAG": ws.Cells(r, 10) = 5: ws.Cells(r, 11) = "Outillage": r = r + 1
    ws.Cells(r, 1) = "OUT-003": ws.Cells(r, 2) = "Tournevis isole lot 5": ws.Cells(r, 3) = 30: ws.Cells(r, 4) = 10: ws.Cells(r, 5) = "Outillage": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 30: ws.Cells(r, 8) = 1200: ws.Cells(r, 9) = "OUTILMAG": ws.Cells(r, 10) = 10: ws.Cells(r, 11) = "Outillage": r = r + 1
    ws.Cells(r, 1) = "OUT-004": ws.Cells(r, 2) = "Marteau 500g": ws.Cells(r, 3) = 20: ws.Cells(r, 4) = 8: ws.Cells(r, 5) = "Outillage": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 20: ws.Cells(r, 8) = 1800: ws.Cells(r, 9) = "OUTILMAG": ws.Cells(r, 10) = 8: ws.Cells(r, 11) = "Outillage": r = r + 1
    ws.Cells(r, 1) = "PEI-001": ws.Cells(r, 2) = "Peinture blanche 25L": ws.Cells(r, 3) = 20: ws.Cells(r, 4) = 5: ws.Cells(r, 5) = "Peinture": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 20: ws.Cells(r, 8) = 12000: ws.Cells(r, 9) = "PEINTURE": ws.Cells(r, 10) = 5: ws.Cells(r, 11) = "Peinture": r = r + 1
    ws.Cells(r, 1) = "PEI-002": ws.Cells(r, 2) = "Peinture grise 25L": ws.Cells(r, 3) = 15: ws.Cells(r, 4) = 5: ws.Cells(r, 5) = "Peinture": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 15: ws.Cells(r, 8) = 11000: ws.Cells(r, 9) = "PEINTURE": ws.Cells(r, 10) = 5: ws.Cells(r, 11) = "Peinture": r = r + 1
    ws.Cells(r, 1) = "PEI-003": ws.Cells(r, 2) = "Sous-couche 25L": ws.Cells(r, 3) = 10: ws.Cells(r, 4) = 3: ws.Cells(r, 5) = "Peinture": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 10: ws.Cells(r, 8) = 8500: ws.Cells(r, 9) = "PEINTURE": ws.Cells(r, 10) = 3: ws.Cells(r, 11) = "Peinture": r = r + 1
    ws.Cells(r, 1) = "CAR-001": ws.Cells(r, 2) = "Carrelage 30x30 m2": ws.Cells(r, 3) = 200: ws.Cells(r, 4) = 40: ws.Cells(r, 5) = "Carrelage": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 200: ws.Cells(r, 8) = 850: ws.Cells(r, 9) = "CERAMIQUE": ws.Cells(r, 10) = 40: ws.Cells(r, 11) = "Carrelage": r = r + 1
    ws.Cells(r, 1) = "CAR-002": ws.Cells(r, 2) = "Carrelage 60x60 m2": ws.Cells(r, 3) = 150: ws.Cells(r, 4) = 30: ws.Cells(r, 5) = "Carrelage": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 150: ws.Cells(r, 8) = 1800: ws.Cells(r, 9) = "CERAMIQUE": ws.Cells(r, 10) = 30: ws.Cells(r, 11) = "Carrelage": r = r + 1
    ws.Cells(r, 1) = "CAR-003": ws.Cells(r, 2) = "Carrelage ext m2": ws.Cells(r, 3) = 80: ws.Cells(r, 4) = 15: ws.Cells(r, 5) = "Carrelage": ws.Cells(r, 6) = "": ws.Cells(r, 7) = 80: ws.Cells(r, 8) = 1200: ws.Cells(r, 9) = "CERAMIQUE": ws.Cells(r, 10) = 15: ws.Cells(r, 11) = "Carrelage": r = r + 1
    Debug.Print "[DemoData] 40 articles seeded"
End Sub

Private Sub SeedSuppliers9()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("FOURNISSEURS")
    Dim r As Long
    r = 2
    ws.Cells(r, 1) = "SIDERAL": ws.Cells(r, 2) = "SIDERAL SPA": ws.Cells(r, 3) = "Zone Industrielle": ws.Cells(r, 4) = "049 00 00 01": ws.Cells(r, 5) = "000116010002500": ws.Cells(r, 6) = "0161600100250": ws.Cells(r, 7) = "32/00-0012345B67": ws.Cells(r, 8) = "250": r = r + 1
    ws.Cells(r, 1) = "CIMENTAL": ws.Cells(r, 2) = "CIMENTAL SPA": ws.Cells(r, 3) = "Bab Ezzouar Alger": ws.Cells(r, 4) = "023 00 00 02": ws.Cells(r, 5) = "000231010003400": ws.Cells(r, 6) = "0313100100340": ws.Cells(r, 7) = "16/00-0023456B12": ws.Cells(r, 8) = "340": r = r + 1
    ws.Cells(r, 1) = "GRANULATS": ws.Cells(r, 2) = "GRANULATS SA": ws.Cells(r, 3) = "Tipaza Alger": ws.Cells(r, 4) = "024 00 00 03": ws.Cells(r, 5) = "000416010004500": ws.Cells(r, 6) = "0161600100450": ws.Cells(r, 7) = "42/00-0034567B89": ws.Cells(r, 8) = "450": r = r + 1
    ws.Cells(r, 1) = "PLASTIQUE": ws.Cells(r, 2) = "PLASTIQUE PRO": ws.Cells(r, 3) = "Es Senia Oran": ws.Cells(r, 4) = "041 00 00 04": ws.Cells(r, 5) = "000531010005600": ws.Cells(r, 6) = "0313100100560": ws.Cells(r, 7) = "31/00-0045678B90": ws.Cells(r, 8) = "560": r = r + 1
    ws.Cells(r, 1) = "ELECTRO": ws.Cells(r, 2) = "ELECTRO PLUS": ws.Cells(r, 3) = "Constantine": ws.Cells(r, 4) = "031 00 00 05": ws.Cells(r, 5) = "000632010006700": ws.Cells(r, 6) = "0323200100670": ws.Cells(r, 7) = "25/00-0056789B01": ws.Cells(r, 8) = "670": r = r + 1
    ws.Cells(r, 1) = "DISTRIBUTION": ws.Cells(r, 2) = "DISTRIBUTION ELECTRIQUE": ws.Cells(r, 3) = "Setif": ws.Cells(r, 4) = "036 00 00 06": ws.Cells(r, 5) = "000719010007800": ws.Cells(r, 6) = "0191900100780": ws.Cells(r, 7) = "19/00-0067890B12": ws.Cells(r, 8) = "780": r = r + 1
    ws.Cells(r, 1) = "OUTILMAG": ws.Cells(r, 2) = "OUTILMAG": ws.Cells(r, 3) = "Bab Ezzouar Alger": ws.Cells(r, 4) = "023 00 00 07": ws.Cells(r, 5) = "000816010008900": ws.Cells(r, 6) = "0161600100890": ws.Cells(r, 7) = "16/00-0078901B23": ws.Cells(r, 8) = "890": r = r + 1
    ws.Cells(r, 1) = "PEINTURE": ws.Cells(r, 2) = "PEINTURE PLUS": ws.Cells(r, 3) = "Blida": ws.Cells(r, 4) = "025 00 00 08": ws.Cells(r, 5) = "000909010009000": ws.Cells(r, 6) = "0090900100900": ws.Cells(r, 7) = "09/00-0089012B34": ws.Cells(r, 8) = "900": r = r + 1
    ws.Cells(r, 1) = "CERAMIQUE": ws.Cells(r, 2) = "CERAMIQUE": ws.Cells(r, 3) = "Centre": ws.Cells(r, 4) = "049 00 00 09": ws.Cells(r, 5) = "001032010010100": ws.Cells(r, 6) = "0323200101010": ws.Cells(r, 7) = "32/00-0090123B45": ws.Cells(r, 8) = "101": r = r + 1
    Debug.Print "[DemoData] 9 suppliers seeded"
End Sub

Private Sub SeedMovements90()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MOUVEMENTS")
    Dim r As Long
    r = 2
    Dim docBS As Long: docBS = 1
    Dim docBR As Long: docBR = 1
    Dim baseDate As Date: baseDate = DateSerial(2026, 4, 1)
    
    ' SORTIE (outgoing sales) - columns: Date, CodeArt, Designation, Type, Qte, Valeur, RefDoc, PU, Tiers, Notes
    ' ENTREE (incoming purchases) - same columns
    
    ' FER-001: Sell 25 every 3 days, buy 120 every 12 days
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "FER-001", "Fer a beton 10mm 12m", 3, 25, 4000, 12, 120, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "FER-002", "Fer a beton 12mm 12m", 4, 30, 6000, 15, 80, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "FER-003", "Fer a beton 14mm 12m", 7, 35, 8000, 18, 60, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "FER-004", "Fer a beton 16mm 12m", 10, 40, 10500, 22, 40, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "ACI-001", "Acier filete 12mm 6m", 6, 30, 3500, 14, 60, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "ACI-002", "Acier filete 16mm 6m", 12, 25, 5200, 18, 30, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "TOLE-001", "Tole acier 2mm 1x2m", 8, 20, 4800, 16, 25, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "TOLE-002", "Tole acier 3mm 1x2m", 15, 12, 7200, 20, 15, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "CIM-001", "Ciment CPA 55 35kg", 3, 35, 550, 8, 120, "CIMENTAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "CIM-002", "Ciment CPA 45 35kg", 4, 30, 480, 10, 150, "CIMENTAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "CIM-003", "Ciment CRC 35kg", 8, 18, 620, 12, 80, "CIMENTAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "SAB-001", "Sable fin 1 tonne", 8, 5, 3500, 10, 15, "GRANULATS")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "SAB-002", "Sable concasse 1 tonne", 10, 4, 4200, 12, 10, "GRANULATS")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "GRA-001", "Gravier 5-25mm 1 tonne", 9, 4, 4800, 11, 10, "GRANULATS")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "GRA-002", "Gravier 0-5mm 1 tonne", 15, 2, 3800, 14, 8, "GRANULATS")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "PVC-001", "Tuyau PVC 32mm 6m", 4, 10, 450, 6, 40, "PLASTIQUE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "PVC-002", "Tuyau PVC 50mm 6m", 6, 7, 780, 8, 30, "PLASTIQUE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "PVC-003", "Tuyau PVC 110mm 6m", 12, 4, 1800, 10, 20, "PLASTIQUE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "PVC-004", "Tuyau PVC 160mm 6m", 20, 2, 3200, 12, 10, "PLASTIQUE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "FER-010", "Tuyau galva 25mm 6m", 8, 5, 2200, 10, 20, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "FER-011", "Tuyau galva 32mm 6m", 15, 3, 3500, 12, 15, "SIDERAL")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "RAC-001", "Raccords PVC lot 100", 6, 6, 850, 8, 30, "PLASTIQUE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "ROB-001", "Robinetterie lot 10", 12, 3, 2500, 10, 15, "PLASTIQUE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "CAB-001", "Cable 2.5mm2 100m", 4, 6, 3800, 8, 25, "ELECTRO")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "CAB-002", "Cable 4mm2 100m", 6, 4, 5500, 10, 15, "ELECTRO")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "CAB-003", "Cable 6mm2 100m", 12, 2, 8200, 12, 10, "ELECTRO")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "DIS-001", "Disjoncteur 20A lot 10", 7, 4, 4500, 8, 15, "DISTRIBUTION")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "DIS-002", "Disjoncteur 32A lot 10", 12, 2, 6800, 10, 10, "DISTRIBUTION")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "PRI-001", "Prise murale lot 20", 4, 6, 2200, 6, 25, "DISTRIBUTION")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "INT-001", "Interrupteur lot 20", 4, 6, 1800, 6, 25, "DISTRIBUTION")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "OUT-001", "Perceuse Makita", 20, 1, 15000, 25, 3, "OUTILMAG")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "OUT-002", "Meuleuse 125mm", 15, 1, 8500, 20, 4, "OUTILMAG")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "OUT-003", "Tournevis isole lot 5", 8, 2, 1200, 10, 8, "OUTILMAG")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "OUT-004", "Marteau 500g", 12, 2, 1800, 14, 6, "OUTILMAG")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "PEI-001", "Peinture blanche 25L", 10, 3, 12000, 12, 6, "PEINTURE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "PEI-002", "Peinture grise 25L", 15, 2, 11000, 15, 4, "PEINTURE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "PEI-003", "Sous-couche 25L", 20, 1, 8500, 18, 3, "PEINTURE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "CAR-001", "Carrelage 30x30 m2", 6, 12, 850, 10, 30, "CERAMIQUE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "CAR-002", "Carrelage 60x60 m2", 10, 6, 1800, 12, 15, "CERAMIQUE")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call AddMvtBatch(ws, r, docBS, docBR, baseDate, "CAR-003", "Carrelage ext m2", 15, 3, 1200, 14, 8, "CERAMIQUE")
    
    Debug.Print "[DemoData] Movements generated"
End Sub

Private Sub AddMvtBatch(ws As Worksheet, startRow As Long, docBS As Long, docBR As Long, baseDate As Date, artCode As String, artName As String, outInt As Long, outQty As Long, outPU As Long, inInt As Long, inQty As Long, supplier As String)
    Dim r As Long: r = startRow
    Dim d As Long
    
    ' SORTIE rows
    d = outInt
    Do While d <= 89
        ws.Cells(r, 1) = baseDate + d - 1
        ws.Cells(r, 2) = artCode
        ws.Cells(r, 3) = artName
        ws.Cells(r, 4) = "SORTIE"
        ws.Cells(r, 5) = outQty
        ws.Cells(r, 6) = outQty * outPU
        ws.Cells(r, 7) = "BS-" & Format(docBS, "00000")
        ws.Cells(r, 8) = outPU
        ws.Cells(r, 9) = "Vente"
        ws.Cells(r, 10) = ""
        docBS = docBS + 1
        r = r + 1
        d = d + outInt
    Loop
    
    ' ENTREE rows
    d = inInt
    Do While d <= 89
        ws.Cells(r, 1) = baseDate + d - 1
        ws.Cells(r, 2) = artCode
        ws.Cells(r, 3) = artName
        ws.Cells(r, 4) = "ENTREE"
        ws.Cells(r, 5) = inQty
        ws.Cells(r, 6) = inQty * outPU
        ws.Cells(r, 7) = "BR-" & Format(docBR, "00000")
        ws.Cells(r, 8) = outPU
        ws.Cells(r, 9) = supplier
        ws.Cells(r, 10) = ""
        docBR = docBR + 1
        r = r + 1
        d = d + inInt
    Loop
End Sub

Private Sub UpdateStockFromMovements()
    Dim wsArt As Worksheet
    Dim wsMouv As Worksheet
    Set wsArt = ThisWorkbook.Sheets("ARTICLES")
    Set wsMouv = ThisWorkbook.Sheets("MOUVEMENTS")
    Dim lastArt As Long: lastArt = wsArt.Cells(wsArt.Rows.Count, 1).End(xlUp).Row
    Dim lastMouv As Long: lastMouv = wsMouv.Cells(wsMouv.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastArt
        Dim code As String: code = wsArt.Cells(i, 1).Value
        If code <> "" Then
            Dim initStk As Long: initStk = wsArt.Cells(i, 3).Value
            Dim totalIn As Long: totalIn = 0
            Dim totalOut As Long: totalOut = 0
            Dim m As Long
            For m = 2 To lastMouv
                If wsMouv.Cells(m, 2).Value = code Then
                    If wsMouv.Cells(m, 4).Value = "ENTREE" Then
                        totalIn = totalIn + wsMouv.Cells(m, 5).Value
                    ElseIf wsMouv.Cells(m, 4).Value = "SORTIE" Then
                        totalOut = totalOut + wsMouv.Cells(m, 5).Value
                    End If
                End If
            Next m
            Dim finalStk As Long: finalStk = initStk + totalIn - totalOut
            If finalStk < 0 Then finalStk = 0
            wsArt.Cells(i, 3).Value = finalStk
            wsArt.Cells(i, 7).Value = finalStk
        End If
    Next i
End Sub

Public Sub RegenerateAllDemoData()
    If MsgBox("Regenerate all demo data?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        Call GenerateDemoData
    End If
End Sub

Private Sub AutoFitAllData()
    Dim ws As Worksheet
    
    ' AutoFit ARTICLES columns
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ARTICLES")
    If Not ws Is Nothing Then
        ws.Columns("A:F").AutoFit
        ws.Columns("G:G").ColumnWidth = 14
        ws.Columns("H:H").ColumnWidth = 16
        ws.Columns("I:I").ColumnWidth = 18
        ws.Columns("J:J").ColumnWidth = 12
        ws.Columns("K:K").ColumnWidth = 14
        ws.Columns("A1:K1").HorizontalAlignment = xlCenter
        ws.Range("A2:A500").HorizontalAlignment = xlLeft
        ws.Range("B2:B500").HorizontalAlignment = xlLeft
        ws.Range("C2:C500").HorizontalAlignment = xlCenter
        ws.Range("D2:D500").HorizontalAlignment = xlCenter
        ws.Range("E2:E500").HorizontalAlignment = xlCenter
        ws.Range("F2:F500").HorizontalAlignment = xlCenter
        ws.Range("G2:G500").NumberFormat = "#,##0"
        ws.Range("H2:H500").NumberFormat = "#,##0"
        ws.Range("I2:I500").NumberFormat = "#,##0.00"
        ws.Range("J2:J500").HorizontalAlignment = xlCenter
        ws.Range("K2:K500").HorizontalAlignment = xlCenter
    End If
    
    ' AutoFit FOURNISSEURS columns
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("FOURNISSEURS")
    If Not ws Is Nothing Then
        ws.Columns("A:H").AutoFit
        ws.Columns("G:G").ColumnWidth = 18
        ws.Range("A1:H1").HorizontalAlignment = xlCenter
        ws.Range("A2:A100").HorizontalAlignment = xlLeft
        ws.Range("B2:B100").HorizontalAlignment = xlLeft
        ws.Range("C2:C100").HorizontalAlignment = xlLeft
        ws.Range("D2:D100").HorizontalAlignment = xlLeft
        ws.Range("E2:E100").HorizontalAlignment = xlLeft
        ws.Range("F2:F100").HorizontalAlignment = xlLeft
        ws.Range("G2:G100").HorizontalAlignment = xlCenter
        ws.Range("H2:H100").HorizontalAlignment = xlCenter
    End If
    
    ' AutoFit MOUVEMENTS columns
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("MOUVEMENTS")
    If Not ws Is Nothing Then
        ws.Columns("A:L").AutoFit
        ws.Columns("A:A").ColumnWidth = 14
        ws.Columns("B:B").ColumnWidth = 16
        ws.Columns("C:C").ColumnWidth = 14
        ws.Columns("D:D").HorizontalAlignment = xlCenter
        ws.Columns("E:E").HorizontalAlignment = xlCenter
        ws.Columns("F:F").NumberFormat = "#,##0"
        ws.Columns("G:G").NumberFormat = "#,##0.00"
        ws.Columns("H:H").NumberFormat = "#,##0.00"
        ws.Columns("I:I").HorizontalAlignment = xlCenter
        ws.Columns("J:J").HorizontalAlignment = xlCenter
        ws.Columns("K:K").HorizontalAlignment = xlLeft
        ws.Columns("L:L").HorizontalAlignment = xlLeft
        ws.Range("A1:L1").HorizontalAlignment = xlCenter
    End If
    
    ' AutoFit CONFIG columns
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("CONFIG")
    If Not ws Is Nothing Then
        ws.Columns("A:A").ColumnWidth = 24
        ws.Columns("B:B").ColumnWidth = 20
        ws.Range("A1:B1").HorizontalAlignment = xlCenter
        ws.Range("A2:A20").HorizontalAlignment = xlLeft
        ws.Range("B2:B20").HorizontalAlignment = xlLeft
    End If
    
    ' AutoFit DASHBOARD columns
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("DASHBOARD")
    If Not ws Is Nothing Then
        ws.Columns("A:L").AutoFit
        ws.Columns("A1:L1").HorizontalAlignment = xlCenter
    End If
    
    ' AutoFit FACTURES columns
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("FACTURES")
    If Not ws Is Nothing Then
        ws.Columns("A:N").AutoFit
        ws.Columns("G:G").NumberFormat = "#,##0.00"
        ws.Columns("H:H").NumberFormat = "#,##0.00"
        ws.Columns("I:I").NumberFormat = "#,##0.00"
        ws.Columns("J:J").NumberFormat = "#,##0.00"
        ws.Columns("K:K").NumberFormat = "#,##0.00"
        ws.Range("A1:N1").HorizontalAlignment = xlCenter
    End If
    
    ' AutoFit BARCODES columns
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("BARCODES")
    If Not ws Is Nothing Then
        ws.Columns("A:D").AutoFit
        ws.Columns("A1:D1").HorizontalAlignment = xlCenter
    End If
    
    ' AutoFit BONS_COMMANDE columns
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("BONS_COMMANDE")
    If Not ws Is Nothing Then
        ws.Columns("A:H").AutoFit
        ws.Columns("F:F").NumberFormat = "#,##0"
        ws.Columns("G:G").NumberFormat = "#,##0.00"
        ws.Columns("H:H").NumberFormat = "#,##0.00"
        ws.Range("A1:H1").HorizontalAlignment = xlCenter
    End If
    
    ' AutoFit AUDIT_LOG columns
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("AUDIT_LOG")
    If Not ws Is Nothing Then
        ws.Columns("A:H").AutoFit
        ws.Columns("A1:H1").HorizontalAlignment = xlCenter
    End If
    
    ' AutoFit ACCUEIL labels
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("ACCUEIL")
    If Not ws Is Nothing Then
        ws.Columns("A:C").AutoFit
    End If
    
    On Error GoTo 0
End Sub
