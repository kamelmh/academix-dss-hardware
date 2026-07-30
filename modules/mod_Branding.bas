Attribute VB_Name = "mod_Branding"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Branding and version information
' ============================================================================
'
' This module provides:
' - Version constants and build information
' - About dialog
' - Splash screen
' - Branding customization support
'
' Usage:
'   mod_Branding.ShowAbout        ' Show about dialog
'   mod_Branding.ShowSplash       ' Show splash screen
'   mod_Branding.GetVersionString ' Get version string
'
' ============================================================================

Option Explicit

' ---- Version Constants ----
Public Const APP_NAME As String = "Academix DSS"
Public Const APP_NAME_SHORT As String = "Academix"
Public Const APP_VERSION As String = "14.0"
Public Const APP_BUILD As String = "2026.07.30"
Public Const APP_AUTHOR As String = "Mahi Kamel Abdelghani"
Public Const APP_EMAIL As String = "kamelmahi71@gmail.com"
Public Const APP_PHONE As String = "+213 676 77 38 92"
Public Const APP_GITHUB As String = "github.com/kamelmh/academix-dss-hardware"
Public Const APP_COPYRIGHT As String = "Copyright (c) 2025-2026 Mahi Kamel Abdelghani"
Public Const APP_LICENSE As String = "Proprietary - All rights reserved"

' ---- Branding Colors ----
Public Const BRAND_PRIMARY As Long = &H7F4600     ' Burnt orange (primary)
Public Const BRAND_SECONDARY As Long = &HFFFFFF   ' White (secondary)
Public Const BRAND_ACCENT As Long = &H404040      ' Dark gray (accent)
Public Const BRAND_SUCCESS As Long = &H00C000     ' Green (success)
Public Const BRAND_WARNING As Long = &H00C0FF     ' Orange (warning)
Public Const BRAND_DANGER As Long = &H0000FF      ' Red (danger)

' ---- Application Info ----
Public Const APP_TITLE As String = "Academix DSS - Gestion de Stock"
Public Const APP_DESCRIPTION As String = _
    "Decision Support System for Inventory Management" & vbCrLf & _
    "Systeme d'Aide a la Decision pour la Gestion de Stock"

' ============================================================================
' GetVersionString - Returns formatted version string
' ============================================================================
Public Function GetVersionString() As String
    GetVersionString = APP_NAME & " v" & APP_VERSION & " (Build " & APP_BUILD & ")"
End Function

' ============================================================================
' GetShortVersion - Returns short version string
' ============================================================================
Public Function GetShortVersion() As String
    GetShortVersion = "v" & APP_VERSION
End Function

' ============================================================================
' GetAboutText - Returns about information
' ============================================================================
Public Function GetAboutText() As String
    Dim s As String
    s = APP_NAME & vbCrLf
    s = s & "Version " & APP_VERSION & " (Build " & APP_BUILD & ")" & vbCrLf
    s = s & vbCrLf
    s = s & APP_DESCRIPTION & vbCrLf
    s = s & vbCrLf
    s = s & "Developpe par:" & vbCrLf
    s = s & "  " & APP_AUTHOR & vbCrLf
    s = s & "  " & APP_EMAIL & vbCrLf
    s = s & "  " & APP_PHONE & vbCrLf
    s = s & vbCrLf
    s = s & "GitHub: " & APP_GITHUB & vbCrLf
    s = s & vbCrLf
    s = s & APP_COPYRIGHT & vbCrLf
    s = s & APP_LICENSE
    GetAboutText = s
End Function

' ============================================================================
' ShowAbout - Shows about dialog
' ============================================================================
Public Sub ShowAbout()
    Dim msg As String
    msg = GetAboutText()
    MsgBox msg, vbInformation, APP_TITLE & " - A propos"
End Sub

' ============================================================================
' ShowSplash - Shows splash screen (2 seconds)
' ============================================================================
Public Sub ShowSplash()
    Dim startTime As Double
    startTime = Timer
    
    ' Simple splash using status bar
    Application.StatusBar = GetVersionString() & " | " & APP_AUTHOR
    
    ' Wait 2 seconds
    Do While Timer - startTime < 2
        DoEvents
    Loop
    
    ' Clear status bar
    Application.StatusBar = False
End Sub

' ============================================================================
' GetBrandedCaption - Returns branded caption for forms
' ============================================================================
Public Function GetBrandedCaption(Optional ByVal formName As String = "") As String
    If Len(formName) = 0 Then
        GetBrandedCaption = APP_TITLE
    Else
        GetBrandedCaption = APP_TITLE & " - " & formName
    End If
End Function

' ============================================================================
' GetBrandedMsgBox - Shows branded message box
' ============================================================================
Public Function GetBrandedMsgBox( _
    ByVal prompt As String, _
    Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional ByVal title As String = "" _
) As VbMsgBoxResult
    If Len(title) = 0 Then
        title = APP_TITLE
    End If
    GetBrandedMsgBox = MsgBox(prompt, buttons, title)
End Function

' ============================================================================
' ApplyTheme - Applies branding theme to a form
' ============================================================================
Public Sub ApplyTheme(ByVal frm As Object)
    On Error Resume Next
    
    ' Apply to form
    frm.BackColor = BRAND_SECONDARY
    
    ' Apply to controls
    Dim ctrl As Object
    For Each ctrl In frm.Designer.Controls
        Select Case True
            Case TypeOf ctrl Is MSForms.Label
                If ctrl.Font.Bold Then
                    ctrl.ForeColor = BRAND_PRIMARY
                End If
            Case TypeOf ctrl Is MSForms.CommandButton
                If ctrl.Default Then
                    ctrl.BackColor = BRAND_PRIMARY
                    ctrl.ForeColor = BRAND_SECONDARY
                End If
        End Select
    Next ctrl
    
    On Error GoTo 0
End Sub

' ============================================================================
' GetCurrencySymbol - Returns currency symbol
' ============================================================================
Public Function GetCurrencySymbol() As String
    GetCurrencySymbol = mod_Config.CURRENCY_SYMBOL
End Function

' ============================================================================
' FormatCurrency - Formats amount with currency
' ============================================================================
Public Function FormatCurrency(ByVal amount As Double) As String
    FormatCurrency = Format(amount, "#,##0.00") & " " & GetCurrencySymbol()
End Function

' ============================================================================
' GetCompanyInfo - Returns company information from config
' ============================================================================
Public Function GetCompanyInfo() As String
    Dim s As String
    s = mod_Config.BUSINESS_NAME
    If Len(mod_Config.BUSINESS_ADDRESS) > 0 Then
        s = s & vbCrLf & mod_Config.BUSINESS_ADDRESS
    End If
    If Len(mod_Config.BUSINESS_PHONE) > 0 Then
        s = s & vbCrLf & "Tel: " & mod_Config.BUSINESS_PHONE
    End If
    If Len(mod_Config.BUSINESS_NIF) > 0 Then
        s = s & vbCrLf & "NIF: " & mod_Config.BUSINESS_NIF
    End If
    GetCompanyInfo = s
End Function

' ============================================================================
' IsRegistered - Checks if software is registered
' ============================================================================
Public Function IsRegistered() As Boolean
    ' Check for registration key in CONFIG sheet
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CONFIG")
    On Error GoTo 0
    
    If ws Is Nothing Then
        IsRegistered = False
        Exit Function
    End If
    
    ' Check for registration key
    Dim regKey As String
    regKey = ""
    On Error Resume Next
    regKey = CStr(ws.Range("B30").Value)
    On Error GoTo 0
    
    ' For now, always return True (no licensing)
    IsRegistered = True
End Function

' ============================================================================
' GetRegistrationInfo - Returns registration information
' ============================================================================
Public Function GetRegistrationInfo() As String
    If IsRegistered() Then
        GetRegistrationInfo = "Enregistre - " & APP_NAME & " " & APP_VERSION
    Else
        GetRegistrationInfo = "Version d'evaluation - " & APP_NAME & " " & APP_VERSION
    End If
End Function
