Attribute VB_Name = "mod_Cleanup"
Option Explicit

' Sheets to KEEP (hardware store deployment)
Private Const KEEP_SHEETS As String = "ACCUEIL,ARTICLES,MOUVEMENTS,FOURNISSEURS,CONFIG,DASHBOARD,BON_RECEPTION,FACTURES,BARCODES,BONS_COMMANDE,AUDIT_LOG"

Public Sub CleanupWorkbook()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call DeleteUnnecessarySheets
    Call UpdateAccueilBranding
    Call UpdateBonReceptionBranding
    Call ReorderSheets
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Workbook cleaned up!" & vbCrLf & "Sheets remaining: " & ThisWorkbook.Sheets.Count, vbInformation, "DSS v14"
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Erreur cleanup: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

Public Sub ListAllSheets()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim msg As String: msg = "All sheets (" & ThisWorkbook.Sheets.Count & "):" & vbCrLf & vbCrLf
    For Each ws In ThisWorkbook.Sheets
        msg = msg & ws.Name & vbCrLf
    Next ws
    MsgBox msg, vbInformation, "Sheet List"
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

Private Sub DeleteUnnecessarySheets()
    Dim ws As Worksheet
    Dim keepList As Object: Set keepList = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In Split(KEEP_SHEETS, ",")
        keepList(Trim(CStr(k))) = True
    Next k
    
    Dim deleted As Long: deleted = 0
    Dim i As Long
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)
        If Not keepList.Exists(ws.Name) Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            deleted = deleted + 1
        End If
    Next i
    
    Debug.Print "[Cleanup] Deleted " & deleted & " sheets, " & ThisWorkbook.Sheets.Count & " remaining"
End Sub

Private Sub ReorderSheets()
    Dim order As Variant
    order = Array("ACCUEIL", "ARTICLES", "MOUVEMENTS", "FOURNISSEURS", "CONFIG", "DASHBOARD", "BON_RECEPTION")
    Dim i As Long
    For i = LBound(order) To UBound(order)
        On Error Resume Next
        ThisWorkbook.Sheets(CStr(order(i))).Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        On Error GoTo 0
    Next i
End Sub

Private Sub UpdateAccueilBranding()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ACCUEIL")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(245, 245, 245)
    ws.Cells.ColumnWidth = 15
    
    ws.Range("A1").Value = mod_Config.SYS_TITLE
    ws.Range("A1").Font.Size = 24
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(0, 70, 127)
    ws.Range("A1:H1").Merge
    
    ws.Range("A2").Value = "Version " & mod_Config.APP_VERSION
    ws.Range("A2").Font.Size = 14
    ws.Range("A2").Font.Color = RGB(100, 100, 100)
    ws.Range("A2:H2").Merge
    
    ws.Range("A4").Value = "Articles:"
    ws.Range("B4").Value = "=COUNTA(ARTICLES!A:A)-1"
    ws.Range("A5").Value = "Fournisseurs:"
    ws.Range("B5").Value = "=COUNTA(FOURNISSEURS!A:A)-1"
    ws.Range("A6").Value = "Mouvements:"
    ws.Range("B6").Value = "=COUNTA(MOUVEMENTS!A:A)-1"
    ws.Range("A4:A6").Font.Bold = True
    ws.Range("B4:B6").Font.Bold = True
    ws.Range("B4:B6").Font.Size = 12
    
    ws.Range("E4").Value = "Nom:"
    ws.Range("F4").Value = mod_Config.BUSINESS_NAME
    ws.Range("E5").Value = "Adresse:"
    ws.Range("F5").Value = mod_Config.BUSINESS_ADDRESS
    ws.Range("E6").Value = "Tel:"
    ws.Range("F6").Value = mod_Config.BUSINESS_PHONE
    ws.Range("E7").Value = "NIF:"
    ws.Range("F7").Value = mod_Config.BUSINESS_NIF
    ws.Range("E8").Value = "NIS:"
    ws.Range("F8").Value = mod_Config.BUSINESS_NIS
    ws.Range("E9").Value = "RC:"
    ws.Range("F9").Value = mod_Config.BUSINESS_RC
    ws.Range("E4:E9").Font.Bold = True
    
    ws.Range("A8").Value = "Utilisez les boutons en haut pour naviguer."
    ws.Range("A8").Font.Italic = True
    ws.Range("A8").Font.Color = RGB(100, 100, 100)
    ws.Range("A8:H8").Merge
    
    ws.Columns("A:A").ColumnWidth = 14
    ws.Columns("B:B").ColumnWidth = 12
    ws.Columns("E:E").ColumnWidth = 12
    ws.Columns("F:F").ColumnWidth = 30
    
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

Private Sub UpdateBonReceptionBranding()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("BON_RECEPTION")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    Dim cell As Range
    For Each cell In ws.UsedRange
        If IsEmpty(cell) Then GoTo NextCell
        If VarType(cell.Value) = vbString Then
            Dim txt As String: txt = cell.Value
            If InStr(txt, "v8") > 0 Or InStr(txt, "v11") > 0 Or InStr(txt, "v13") > 0 Then
                cell.Value = Replace(Replace(Replace(txt, "v8", mod_Config.APP_VERSION), "v11", mod_Config.APP_VERSION), "v13", mod_Config.APP_VERSION)
            End If
            If InStr(txt, "directeur") > 0 Or InStr(txt, "Directeur") > 0 Then
                cell.Value = Replace(Replace(txt, "directeur", "Gerant"), "Directeur", "Gerant")
            End If
        End If
NextCell:
    Next cell
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub
