Attribute VB_Name = "mod_Backup"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Backup & Restore Module - P0 Priority (Hyperagent recommendation)
' ============================================================================

Option Explicit

' ============================================================================
' CONSTANTS
' ============================================================================
Private Const BACKUP_FOLDER As String = "Academix_Backups"
Private Const MAX_BACKUPS As Integer = 10
Private Const BACKUP_PREFIX As String = "DSS_Backup_"

' ============================================================================
' SUB: AutoBackupOnClose
' Called from ThisWorkbook_BeforeClose - creates timestamped backup
' ============================================================================
Public Sub AutoBackupOnClose()
    On Error Resume Next
    
    ' Only backup if workbook has been modified
    If Not ThisWorkbook.Saved Then
        CreateBackup "AutoClose"
    End If
    
    On Error GoTo 0
End Sub

' ============================================================================
' SUB: AutoBackupOnOpen
' Called from ThisWorkbook_AfterOpen - creates backup of clean version
' ============================================================================
Public Sub AutoBackupOnOpen()
    On Error Resume Next
    
    ' Create a backup on open (in case previous session didn't close properly)
    CreateBackup "AutoOpen"
    
    ' Clean old backups
    CleanOldBackups
    
    On Error GoTo 0
End Sub

' ============================================================================
' FUNCTION: CreateBackup
' Creates a timestamped backup copy of the workbook
' Returns: Full path to backup file
' ============================================================================
Public Function CreateBackup(ByVal backupType As String) As String
    On Error GoTo ErrorHandler
    
    Dim backupDir As String
    Dim backupPath As String
    Dim timestamp As String
    Dim fileName As String
    
    ' Get backup directory
    backupDir = GetBackupDirectory()
    If backupDir = "" Then
        CreateBackup = ""
        Exit Function
    End If
    
    ' Create timestamp: YYYYMMDD_HHMMSS
    timestamp = Format(Now, "YYYYMMDD_HHMMSS")
    
    ' Build filename: DSS_Backup_Type_YYYYMMDD_HHMMSS.xlsm
    fileName = BACKUP_PREFIX & backupType & "_" & timestamp & ".xlsm"
    backupPath = backupDir & "\" & fileName
    
    ' Save copy
    ThisWorkbook.SaveCopyAs backupPath
    
    ' Log backup
    LogBackup backupPath, backupType
    
    CreateBackup = backupPath
    
    Debug.Print "[Backup] Created: " & fileName
    Exit Function
    
ErrorHandler:
    Debug.Print "[Backup] Error: " & Err.Description
    CreateBackup = ""
End Function

' ============================================================================
' FUNCTION: GetBackupDirectory
' Returns the backup directory path, creates if needed
' ============================================================================
Private Function GetBackupDirectory() As String
    On Error GoTo ErrorHandler
    
    Dim basePath As String
    Dim backupDir As String
    
    ' Use same location as workbook
    basePath = ThisWorkbook.Path
    
    ' If workbook hasn't been saved yet, use My Documents
    If basePath = "" Then
        basePath = Environ("USERPROFILE") & "\Documents"
    End If
    
    backupDir = basePath & "\" & BACKUP_FOLDER
    
    ' Create directory if it doesn't exist
    If Dir(backupDir, vbDirectory) = "" Then
        On Error Resume Next
        MkDir backupDir
        If Err.Number <> 0 Then
            ' If can't create subfolder, use base path
            backupDir = basePath
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Verify directory exists
    If Dir(backupDir, vbDirectory) = "" Then
        GetBackupDirectory = basePath
    Else
        GetBackupDirectory = backupDir
    End If
    Exit Function
    
ErrorHandler:
    ' Fallback to workbook path
    GetBackupDirectory = ThisWorkbook.Path
End Function

' ============================================================================
' SUB: CleanOldBackups
' Keeps only the last MAX_BACKUPS backups
' ============================================================================
Private Sub CleanOldBackups()
    On Error Resume Next
    
    Dim backupDir As String
    backupDir = GetBackupDirectory()
    If backupDir = "" Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(backupDir)
    
    ' Collect backup files
    Dim files() As String
    Dim fileCount As Long
    fileCount = 0
    
    Dim f As Object
    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "xlsm" Then
            If Left(f.Name, Len(BACKUP_PREFIX)) = BACKUP_PREFIX Then
                fileCount = fileCount + 1
                ReDim Preserve files(1 To fileCount)
                files(fileCount) = f.Path
            End If
        End If
    Next f
    
    ' If more than MAX_BACKUPS, delete oldest
    If fileCount > MAX_BACKUPS Then
        ' Sort by name (timestamp in name = chronological order)
        Dim i As Long, j As Long
        Dim temp As String
        For i = 1 To fileCount - 1
            For j = i + 1 To fileCount
                If files(i) > files(j) Then
                    temp = files(i)
                    files(i) = files(j)
                    files(j) = temp
                End If
            Next j
        Next i
        
        ' Delete oldest
        For i = 1 To fileCount - MAX_BACKUPS
            Kill files(i)
            Debug.Print "[Backup] Cleaned: " & files(i)
        Next i
    End If
    
    Set fso = Nothing
    On Error GoTo 0
End Sub

' ============================================================================
' SUB: LogBackup
' Logs backup to AUDIT_LOG sheet
' ============================================================================
Private Sub LogBackup(ByVal backupPath As String, ByVal backupType As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_AUDIT_LOG)
    On Error GoTo 0
    
    ' Create AUDIT_LOG if it doesn't exist
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = mod_Config.SHEET_AUDIT_LOG
        
        ' Add headers
        ws.Range("A1").Value = "Timestamp"
        ws.Range("B1").Value = "Action"
        ws.Range("C1").Value = "Details"
        ws.Range("D1").Value = "User"
        
        ' Format header
        ws.Range("A1:D1").Font.Bold = True
        ws.Range("A1:D1").Interior.Color = RGB(30, 60, 114)
        ws.Range("A1:D1").Font.Color = vbWhite
    End If
    
    ' Unprotect sheet before writing
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    
    ' Find next empty row
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Log entry
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "BACKUP_" & UCase(backupType)
    ws.Cells(nextRow, 3).Value = backupPath
    ws.Cells(nextRow, 4).Value = Environ("USERNAME")
    
    ' Format timestamp
    ws.Cells(nextRow, 1).NumberFormat = "DD/MM/YYYY HH:MM:SS"
    
    ' Auto-fit columns
    ws.Columns("A:D").AutoFit
    
    ' Re-protect sheet
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    On Error GoTo 0
End Sub

' ============================================================================
' SUB: ManualBackup
' User-triggered backup from menu/button
' ============================================================================
Public Sub ManualBackup()
    On Error GoTo ErrorHandler
    
    Dim backupPath As String
    backupPath = CreateBackup("Manual")
    
    If backupPath <> "" Then
        MsgBox "Backup created successfully!" & vbCrLf & vbCrLf & _
               "Location: " & backupPath & vbCrLf & vbCrLf & _
               "Last 10 backups are kept automatically.", _
               vbInformation, mod_Config.SYS_TITLE
    Else
        MsgBox "Backup failed. Check write permissions.", _
               vbExclamation, mod_Config.SYS_TITLE
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Backup error: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUB: RestoreBackup
' Restores a backup from the backup folder
' ============================================================================
Public Sub RestoreBackup()
    On Error GoTo ErrorHandler
    
    Dim backupDir As String
    backupDir = GetBackupDirectory()
    If backupDir = "" Then
        MsgBox "Backup directory not found.", vbExclamation, mod_Config.SYS_TITLE
        Exit Sub
    End If
    
    ' Show file dialog
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker = 3
    
    With fd
        .Title = "Select Backup to Restore"
        .InitialFileName = backupDir & "\"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsm"
        .Filters.Add "All Files", "*.*"
        
        If .Show = -1 Then
            Dim selectedFile As String
            selectedFile = .SelectedItems(1)
            
            ' Confirm restore
            If MsgBox("Restore from backup?" & vbCrLf & vbCrLf & _
                      "File: " & Dir(selectedFile) & vbCrLf & vbCrLf & _
                      "This will replace the current workbook.", _
                      vbQuestion + vbYesNo + vbDefaultButton2, _
                      mod_Config.SYS_TITLE) = vbYes Then
                
                ' Create backup of current state first
                CreateBackup "PreRestore"
                
                ' Open the backup
                Workbooks.Open selectedFile
                
                MsgBox "Backup restored successfully." & vbCrLf & _
                       "Save the workbook to keep changes.", _
                       vbInformation, mod_Config.SYS_TITLE
            End If
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Restore error: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUB: ExportToCSV
' Exports data sheets to CSV for portability/backup
' ============================================================================
Public Sub ExportToCSV()
    On Error GoTo ErrorHandler
    
    Dim exportDir As String
    exportDir = ThisWorkbook.Path & "\Academix_Exports"
    
    ' Create export directory
    If Dir(exportDir, vbDirectory) = "" Then
        MkDir exportDir
    End If
    
    Dim timestamp As String
    timestamp = Format(Now, "YYYYMMDD_HHMMSS")
    
    ' Export ARTICLES
    ExportSheetToCSV "ARTICLES", exportDir & "\Articles_" & timestamp & ".csv"
    
    ' Export MOUVEMENTS
    ExportSheetToCSV "MOUVEMENTS", exportDir & "\Mouvements_" & timestamp & ".csv"
    
    ' Export FOURNISSEURS
    ExportSheetToCSV "FOURNISSEURS", exportDir & "\Fournisseurs_" & timestamp & ".csv"
    
    MsgBox "Data exported to CSV files." & vbCrLf & vbCrLf & _
           "Location: " & exportDir, _
           vbInformation, mod_Config.SYS_TITLE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Export error: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

' ============================================================================
' SUB: ExportSheetToCSV
' Exports a single sheet to CSV format
' ============================================================================
Private Sub ExportSheetToCSV(ByVal sheetName As String, ByVal filePath As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then Exit Sub
    
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Copy range
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Save as CSV
    rng.Copy
    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.Sheets(1).Range("A1").PasteSpecial xlPasteValues
    wb.SaveAs filePath, xlCSV
    wb.Close False
    
    Application.CutCopyMode = False
    
    On Error GoTo 0
End Sub

' ============================================================================
' FUNCTION: GetBackupInfo
' Returns info about available backups
' ============================================================================
Public Function GetBackupInfo() As String
    On Error Resume Next
    
    Dim backupDir As String
    backupDir = GetBackupDirectory()
    If backupDir = "" Then
        GetBackupInfo = "No backup directory found."
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(backupDir)
    
    Dim count As Long
    Dim totalSize As Double
    Dim newest As String
    Dim oldest As String
    
    Dim f As Object
    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "xlsm" Then
            If Left(f.Name, Len(BACKUP_PREFIX)) = BACKUP_PREFIX Then
                count = count + 1
                totalSize = totalSize + f.Size
                If newest = "" Or f.DateLastModified > CDate(newest) Then
                    newest = f.DateLastModified
                End If
                If oldest = "" Or f.DateLastModified < CDate(oldest) Then
                    oldest = f.DateLastModified
                End If
            End If
        End If
    Next f
    
    GetBackupInfo = "Backups: " & count & " files (" & Format(totalSize / 1024, "#,##0") & " KB)" & vbCrLf & _
                    "Newest: " & newest & vbCrLf & _
                    "Oldest: " & oldest & vbCrLf & _
                    "Location: " & backupDir
    
    Set fso = Nothing
    On Error GoTo 0
End Function
