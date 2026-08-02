Attribute VB_Name = "mod_Splash"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Splash screen and loading screen
' ============================================================================
'
' This module provides:
' - Splash screen shown during startup
' - Loading indicator for long operations
' - Progress bar support
'
' Usage:
'   mod_Splash.ShowSplash        ' Show splash during startup
'   mod_Splash.ShowLoading "Loading..."  ' Show loading indicator
'   mod_Splash.HideLoading       ' Hide loading indicator
'
' ============================================================================

Option Explicit

' ---- Private State ----
Private m_splashShown As Boolean
Private m_loadingForm As Object

' ============================================================================
' ShowSplash - Shows splash screen during startup
' ============================================================================
Public Sub ShowSplash()
    If m_splashShown Then Exit Sub
    m_splashShown = True
    
    ' Update status bar with splash info
    Application.StatusBar = mod_Branding.GetVersionString() & " | Chargement..."
    
    ' DoEvents to allow Excel to update
    DoEvents
End Sub

' ============================================================================
' HideSplash - Hides splash screen
' ============================================================================
Public Sub HideSplash()
    If Not m_splashShown Then Exit Sub
    m_splashShown = False
    
    ' Clear status bar
    Application.StatusBar = False
    
    DoEvents
End Sub

' ============================================================================
' ShowLoading - Shows loading indicator
' ============================================================================
Public Sub ShowLoading(ByVal message As String)
    Application.StatusBar = mod_Branding.APP_NAME_SHORT & " | " & message
    DoEvents
End Sub

' ============================================================================
' HideLoading - Hides loading indicator
' ============================================================================
Public Sub HideLoading()
    Application.StatusBar = False
    DoEvents
End Sub

' ============================================================================
' ShowProgress - Shows progress in status bar
' ============================================================================
Public Sub ShowProgress(ByVal current As Long, ByVal total As Long, Optional ByVal message As String = "")
    Dim pct As Long
    If total > 0 Then
        pct = (current / total) * 100
    Else
        pct = 0
    End If
    
    Dim bar As String
    Dim i As Long
    For i = 1 To 20
        If i <= (pct / 5) Then
            bar = bar & Chr(9608)  ' [block]
        Else
            bar = bar & Chr(9617)  ' [light]
        End If
    Next i
    
    Dim status As String
    If Len(message) > 0 Then
        status = message & " "
    End If
    status = status & bar & " " & pct & "% (" & current & "/" & total & ")"
    
    Application.StatusBar = status
    DoEvents
End Sub

' ============================================================================
' ShowWelcome - Shows welcome message after splash
' ============================================================================
Public Sub ShowWelcome()
    Dim msg As String
    msg = "Bienvenue dans " & mod_Branding.APP_NAME & "!" & vbCrLf & vbCrLf
    msg = msg & "Version " & mod_Branding.APP_VERSION & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "Developpe par " & mod_Branding.APP_AUTHOR & vbCrLf
    msg = msg & mod_Branding.APP_EMAIL & vbCrLf
    
    MsgBox msg, vbInformation, mod_Branding.APP_TITLE
End Sub

' ============================================================================
' ShowGoodbye - Shows goodbye message
' ============================================================================
Public Sub ShowGoodbye()
    Dim msg As String
    msg = "Merci d'utiliser " & mod_Branding.APP_NAME & "!" & vbCrLf & vbCrLf
    msg = msg & "N'oubliez pas de sauvegarder votre travail." & vbCrLf
    msg = msg & vbCrLf
    msg = msg & mod_Branding.APP_COPYRIGHT
    
    MsgBox msg, vbInformation, mod_Branding.APP_TITLE
End Sub

' ============================================================================
' ShowError - Shows branded error message
' ============================================================================
Public Sub ShowError(ByVal errorNumber As Long, ByVal errorDescription As String)
    Dim msg As String
    msg = "Erreur " & errorNumber & ": " & errorDescription & vbCrLf & vbCrLf
    msg = msg & "Veuillez contacter le support technique:" & vbCrLf
    msg = msg & "  " & mod_Branding.APP_EMAIL & vbCrLf
    msg = msg & "  " & mod_Branding.APP_PHONE
    
    MsgBox msg, vbCritical, mod_Branding.APP_TITLE & " - Erreur"
End Sub

' ============================================================================
' ShowSuccess - Shows branded success message
' ============================================================================
Public Sub ShowSuccess(ByVal message As String)
    Dim msg As String
    msg = message & vbCrLf & vbCrLf
    msg = msg & mod_Branding.GetVersionString()
    
    MsgBox msg, vbInformation, mod_Branding.APP_TITLE
End Sub

' ============================================================================
' ShowWarning - Shows branded warning message
' ============================================================================
Public Sub ShowWarning(ByVal message As String)
    Dim msg As String
    msg = message & vbCrLf & vbCrLf
    msg = msg & "Consultez la documentation pour plus d'informations."
    
    MsgBox msg, vbExclamation, mod_Branding.APP_TITLE & " - Attention"
End Sub
