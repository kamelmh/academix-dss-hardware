Attribute VB_Name = "mod_Config"
' ============================================================================
' Academix v13.4 - DSS Logistique El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Direction de l'Education - Wilaya d'El Bayadh
' Protected under Algerian Copyright Law (Ordinance 03-05, July 19, 2003)
' All rights reserved. Unauthorized reproduction or distribution prohibited.
' ============================================================================

Option Explicit

' Barcode symbology types (used by mod_BarcodeSim and mod_BarcodeEncoder)
Public Enum BarcodeSymbology
    bcCode128 = 0
    bcEAN13 = 1
    bcCode39 = 2
    bcInterleaved2of5 = 3
    bcQR_Visual = 4
End Enum

' Sheet names
Public Const SHEET_MOUVEMENTS As String = "MOUVEMENTS"
Public Const SHEET_ARTICLES As String = "ARTICLES"
Public Const SHEET_SYS_STRINGS As String = "SYS_STRINGS"
Public Const SHEET_AUDIT_LOG As String = "AUDIT_LOG"
Public Const SHEET_STAGING As String = "STAGING_BUFFER"
Public Const SHEET_FOURNISSEURS As String = "FOURNISSEURS"
Public Const SHEET_ACCUEIL As String = "ACCUEIL"

' Document types
Public Const DOC_TYPE_BS As String = "Bon de Sortie"
Public Const DOC_TYPE_DA As String = "Demande d'Achat"

' System constants
Public Const WORKING_DAYS_PER_YEAR As Integer = 250
Public Const OBSERVATION_DAYS As Integer = 38
Public Const REFDOC_PREFIX As String = "BS-"

' Column names for AUDIT_LOG sheet
Public Const COL_AUDIT_DATE As Long = 1
Public Const COL_AUDIT_TIME As Long = 2
Public Const COL_AUDIT_USER As Long = 3
Public Const COL_AUDIT_ACTION As Long = 4
Public Const COL_AUDIT_REF As Long = 5

' ARTICLES sheet column indices
Public Const COL_ART_CODE As Long = 1
Public Const COL_ART_DESIGNATION As Long = 2
Public Const COL_ART_STOCK As Long = 3
Public Const COL_ART_SEUIL_MIN As Long = 4
Public Const COL_ART_CATEGORIE As Long = 5
Public Const COL_ART_CLASSE_ABC As Long = 6
Public Const COL_ART_STOCK_ACTUEL As Long = 7
Public Const COL_ART_PU As Long = 8
Public Const COL_ART_FOURNISSEUR As Long = 9
Public Const COL_ART_STOCK_SECURITE As Long = 10
Public Const COL_ART_NOTES As Long = 11
Public Const COL_ART_CMUP As Long = 12

' MOUVEMENTS sheet column indices
Public Const COL_MOUV_DATE As Long = 1
Public Const COL_MOUV_CODE_ARTICLE As Long = 2
Public Const COL_MOUV_DESIGNATION As Long = 3
Public Const COL_MOUV_TYPE As Long = 4
Public Const COL_MOUV_QTE As Long = 5
Public Const COL_MOUV_VALEUR As Long = 6
Public Const COL_MOUV_REF_DOC As Long = 7
Public Const COL_MOUV_PU As Long = 8
Public Const COL_MOUV_THIRD_PARTY As Long = 9
Public Const COL_MOUV_NOTES As Long = 10

' FOURNISSEURS sheet column indices (matches golden master 9-col evaluation layout)
Public Const COL_FOU_CODE As Long = 1
Public Const COL_FOU_NOM_ABREGE As Long = 2
Public Const COL_FOU_RAISON_SOCIALE As Long = 3
Public Const COL_FOU_WILAYA As Long = 4
Public Const COL_FOU_TELEPHONE As Long = 5
Public Const COL_FOU_CLASSE As Long = 6
Public Const COL_FOU_DELAI As Long = 7
Public Const COL_FOU_NOTE As Long = 8
Public Const COL_FOU_SPECIALITE As Long = 9

' SYS_STRINGS sheet column indices
Public Const COL_SYS_ID As Long = 1
Public Const COL_SYS_VALUE As Long = 2

' Properties (must come after all Const declarations)
Public Property Get SYS_TITLE() As String
    SYS_TITLE = "ERP Acad" & Chr(233) & "mie"
End Property

Public Property Get DOC_TYPE_BR() As String
    DOC_TYPE_BR = "Bon de R" & Chr(201) & "ception"
End Property

Public Property Get DOC_TYPE_BC() As String
    DOC_TYPE_BC = "Bon de Commande"
End Property

Public Property Get MASTER_PWD() As String
    MASTER_PWD = "erp_secure_pwd_2026"
End Property

Public Property Get APP_VERSION() As String
    APP_VERSION = "v13.4"
End Property
