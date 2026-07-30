#!/usr/bin/env python3
"""Generate modules/mod_BuildFirstRun.bas.

Run from anywhere:  python3 tools/genbuilder.py

Edit FORM_CODE below to change frmFirstRun's behaviour, then re-run this script.
Do not hand-edit the generated GetFirstRunFormCode function in the .bas file.

The form's code module has to be emitted as VBA string literals ("c = c & ..."),
which means every embedded quote must be doubled. Hand-writing that is how you
ship a form whose code module does not compile, so the VBA below is written once
in plain form and the escaping is generated. render_back() reverses the process
so the result can be checked with the same validator used on real modules.
"""
import pathlib

# ---------------------------------------------------------------------------
# The form's code module, written plainly. This is what must end up inside
# frmFirstRun after BuildFirstRunForm runs.
# ---------------------------------------------------------------------------
FORM_CODE = r'''
Option Explicit

' Page 1 collects the business identity, page 2 the operating parameters.
' There is no container control: every field sits directly on the form and the
' two groups are swapped by toggling Visible, which is the same flat-control
' approach the other mod_Build* modules in this project use.
Private Sub UserForm_Initialize()
    ' Prefill from the current configuration. On a genuinely fresh workbook the
    ' identity keys are blank, so these come back as the mod_Config fallbacks
    ' and the owner overwrites them.
    txtName.Text = mod_Config.BUSINESS_NAME
    txtAddr.Text = mod_Config.BUSINESS_ADDRESS
    txtPhone.Text = mod_Config.BUSINESS_PHONE
    txtNIF.Text = mod_Config.BUSINESS_NIF
    txtNIS.Text = mod_Config.BUSINESS_NIS
    txtRC.Text = mod_Config.BUSINESS_RC

    txtWDays.Text = mod_Config.WORKING_DAYS_PER_YEAR
    txtOCost.Text = mod_Config.ORDER_COST
    txtHRate.Text = mod_Config.HOLDING_RATE
    txtLead.Text = mod_Config.LEAD_TIME_DEFAULT
    txtTax.Text = mod_Config.TAX_RATE
    txtCurr.Text = mod_Config.CURRENCY_SYMBOL

    ShowPage 1
End Sub

Private Sub ShowPage(ByVal n As Long)
    Dim p1 As Boolean, p2 As Boolean
    p1 = (n = 1)
    p2 = (n = 2)

    lblName.Visible = p1
    txtName.Visible = p1
    lblAddr.Visible = p1
    txtAddr.Visible = p1
    lblPhone.Visible = p1
    txtPhone.Visible = p1
    lblNIF.Visible = p1
    txtNIF.Visible = p1
    lblNIS.Visible = p1
    txtNIS.Visible = p1
    lblRC.Visible = p1
    txtRC.Visible = p1

    lblWDays.Visible = p2
    txtWDays.Visible = p2
    lblOCost.Visible = p2
    txtOCost.Visible = p2
    lblHRate.Visible = p2
    txtHRate.Visible = p2
    lblLead.Visible = p2
    txtLead.Visible = p2
    lblTax.Visible = p2
    txtTax.Visible = p2
    lblCurr.Visible = p2
    txtCurr.Visible = p2

    btnNext.Visible = p1
    btnBack.Visible = p2
    btnStart.Visible = p2
    btnDemo.Visible = p2

    If p1 Then
        lblStep.Caption = "Etape 1 sur 2 - Identite de l'etablissement"
        lblHint.Caption = "Ces informations figurent sur vos factures et vos bons."
        txtName.SetFocus
    Else
        lblStep.Caption = "Etape 2 sur 2 - Parametres de gestion"
        lblHint.Caption = "Les valeurs proposees conviennent a une quincaillerie. Modifiables plus tard."
        txtWDays.SetFocus
    End If
End Sub

Private Sub btnNext_Click()
    If Len(Trim(txtName.Text)) = 0 Then
        MsgBox "Indiquez le nom commercial avant de continuer.", vbExclamation, mod_Config.SYS_TITLE
        txtName.SetFocus
        Exit Sub
    End If
    ShowPage 2
End Sub

Private Sub btnBack_Click()
    ShowPage 1
End Sub

' Validation, normalisation and writing all live in mod_FirstRun so that the
' form and the no-form fallback wizard cannot drift apart. False means something
' was rejected and a message has already been shown, so the form stays open.
Private Function PersistConfig() As Boolean
    PersistConfig = mod_FirstRun.SaveFirstRunConfig( _
        txtName.Text, txtAddr.Text, txtPhone.Text, _
        txtNIF.Text, txtNIS.Text, txtRC.Text, _
        txtWDays.Text, txtOCost.Text, txtHRate.Text, _
        txtLead.Text, txtTax.Text, txtCurr.Text)
End Function

Private Sub btnStart_Click()
    If Not PersistConfig() Then Exit Sub
    Me.Hide
    mod_FirstRun.CompleteFirstRun False
    Unload Me
End Sub

Private Sub btnDemo_Click()
    Dim m As String
    m = "Charger les 40 articles de demonstration ?" & vbCrLf & vbCrLf & _
        "Utile pour se former ou presenter le systeme." & vbCrLf & _
        "Pour un usage reel, preferez Demarrer a vide."
    If MsgBox(m, vbQuestion + vbYesNo + vbDefaultButton2, mod_Config.SYS_TITLE) = vbNo Then Exit Sub

    If Not PersistConfig() Then Exit Sub
    Me.Hide
    mod_FirstRun.CompleteFirstRun True
    Unload Me
End Sub

' Closing the wizard leaves FIRST_RUN set, so it comes back on the next open
' rather than leaving the system half configured.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "Configuration interrompue." & vbCrLf & _
               "L'assistant sera propose de nouveau a la prochaine ouverture.", _
               vbInformation, mod_Config.SYS_TITLE
    End If
End Sub
'''


def vba_escape(line: str) -> str:
    """One plain line -> one 'c = c & "..." & vbCrLf' builder statement."""
    return '    c = c & "' + line.replace('"', '""') + '" & vbCrLf'


def build_codegen_function(form_code: str) -> str:
    lines = form_code.strip("\n").split("\n")
    out = ["Private Function GetFirstRunFormCode() As String", "    Dim c As String", ""]
    out.append('    c = "' + lines[0].replace('"', '""') + '" & vbCrLf')
    for l in lines[1:]:
        out.append(vba_escape(l))
    out += ["", "    GetFirstRunFormCode = c", "End Function"]
    return "\n".join(out)


def render_back(codegen: str) -> str:
    """Reverse the escaping, to confirm what VBA will actually assemble."""
    import re
    rendered = []
    for line in codegen.split("\n"):
        m = re.match(r'\s*c = (?:c & )?"(.*)" & vbCrLf\s*$', line)
        if m:
            rendered.append(m.group(1).replace('""', '"'))
    return "\n".join(rendered)


# ---------------------------------------------------------------------------
# Control layout. (kind, name, caption, left, top, width, height, extra)
# Page-1 and page-2 fields deliberately share the same top coordinates: only
# one group is ever visible, so the form stays compact.
# ---------------------------------------------------------------------------
LBL_L, FLD_L, FLD_W = 18, 165, 320
ROWS = [88, 116, 144, 172, 200, 228]

P1 = [
    ("lblName", "Nom commercial *", "txtName"),
    ("lblAddr", "Adresse", "txtAddr"),
    ("lblPhone", "Telephone", "txtPhone"),
    ("lblNIF", "NIF (15 chiffres)", "txtNIF"),
    ("lblNIS", "NIS (15 chiffres)", "txtNIS"),
    ("lblRC", "RC (00/00-0000000A00)", "txtRC"),
]
P2 = [
    ("lblWDays", "Jours ouvres par an", "txtWDays"),
    ("lblOCost", "Cout d'une commande", "txtOCost"),
    ("lblHRate", "Taux de possession", "txtHRate"),
    ("lblLead", "Delai de livraison (jours)", "txtLead"),
    ("lblTax", "Taux de TVA", "txtTax"),
    ("lblCurr", "Devise", "txtCurr"),
]


def emit_controls():
    L = []
    A = L.append
    A("    Dim ctrl As Object")
    A("")
    A("    ' ---- header ----")
    A('    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")')
    A('    With ctrl: .Name = "lblTitle": .Caption = "Configuration initiale": '
      '.Left = 18: .Top = 14: .Width = 460: .Height = 26: '
      '.Font.Size = 15: .Font.Bold = True: .ForeColor = &H7F4600: End With')
    A('    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")')
    A('    With ctrl: .Name = "lblStep": .Caption = "Etape 1 sur 2": '
      '.Left = 18: .Top = 42: .Width = 460: .Height = 18: .Font.Bold = True: End With')
    A('    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")')
    A('    With ctrl: .Name = "lblHint": .Caption = "": '
      '.Left = 18: .Top = 60: .Width = 460: .Height = 18: .ForeColor = &H808080: End With')

    for page, spec in (("1", P1), ("2", P2)):
        A("")
        A(f"    ' ---- page {page} fields ----")
        vis = "" if page == "1" else ": .Visible = False"
        for (lbl, cap, txt), top in zip(spec, ROWS):
            A('    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")')
            A(f'    With ctrl: .Name = "{lbl}": .Caption = "{cap}": '
              f'.Left = {LBL_L}: .Top = {top}: .Width = 140: .Height = 20{vis}: End With')
            A('    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")')
            A(f'    With ctrl: .Name = "{txt}": '
              f'.Left = {FLD_L}: .Top = {top}: .Width = {FLD_W}: .Height = 22{vis}: End With')

    A("")
    A("    ' ---- navigation ----")
    A('    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")')
    A('    With ctrl: .Name = "btnBack": .Caption = "< Precedent": '
      '.Left = 18: .Top = 274: .Width = 90: .Height = 32: .Visible = False: End With')
    A('    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")')
    A('    With ctrl: .Name = "btnNext": .Caption = "Suivant >": '
      '.Left = 385: .Top = 274: .Width = 100: .Height = 32: '
      '.BackColor = &H7F4600: .ForeColor = &HFFFFFF: End With')
    A('    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")')
    A('    With ctrl: .Name = "btnStart": .Caption = "Demarrer a vide": '
      '.Left = 335: .Top = 274: .Width = 150: .Height = 32: '
      '.BackColor = &H7F4600: .ForeColor = &HFFFFFF: .Visible = False: End With')
    A('    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")')
    A('    With ctrl: .Name = "btnDemo": .Caption = "Charger donnees demo": '
      '.Left = 150: .Top = 274: .Width = 175: .Height = 32: .Visible = False: End With')
    return "\n".join(L)


MODULE = '''Attribute VB_Name = "mod_BuildFirstRun"
' ============================================================================
' Academix v14.0 - DSS Quincaillerie El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' Builds frmFirstRun - the two-page setup wizard shown on the first open
' ============================================================================
'
' Same approach as the other mod_Build* modules: the form is created at runtime
' through the VBIDE extensibility model, so there is no .frm in the repo and
' "Trust access to the VBA project object model" must be enabled.
'
' The code module below is assembled by GetFirstRunFormCode. That function is
' generated rather than hand-written - see tools/genbuilder.py - because getting
' the doubled quotes wrong by hand produces a form whose code module will not
' compile, and there is no way to notice that until it is opened.
'
' Page 1 is the business identity, page 2 the operating parameters. All controls
' sit flat on the form and the two groups are swapped by toggling Visible; no
' container control is used, matching the rest of this project.
' ============================================================================

Option Explicit

Public Sub BuildFirstRunForm(Optional ByVal silent As Boolean = False)
    On Error GoTo ErrHandler

    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject

    ' Remove any previous build. Guarded because the component legitimately does
    ' not exist the first time round.
    On Error Resume Next
    DoEvents
    vbProj.VBComponents.Remove vbProj.VBComponents("frmFirstRun")
    DoEvents
    On Error GoTo ErrHandler

    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3)

    With frm
        .Properties("Name") = "frmFirstRun"
        .Properties("Caption") = "Configuration initiale - DSS Quincaillerie"
        .Properties("Width") = 500
        .Properties("Height") = 345
        .Properties("StartUpPosition") = 1
    End With

    frm.CodeModule.AddFromString GetFirstRunFormCode()

__CONTROLS__

    If Not silent Then
        MsgBox "frmFirstRun cree." & vbCrLf & _
               "Lancez mod_FirstRun.FirstRunSetup pour l'ouvrir.", _
               vbInformation, "Done"
    End If
    Exit Sub

ErrHandler:
    If Not silent Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Else
        Debug.Print "[BuildFirstRun] " & Err.Number & ": " & Err.Description
    End If
End Sub

__CODEGEN__
'''


def main():
    codegen = build_codegen_function(FORM_CODE)
    module = MODULE.replace("__CONTROLS__", emit_controls()).replace("__CODEGEN__", codegen)
    repo_root = pathlib.Path(__file__).resolve().parent.parent
    out = repo_root / "modules" / "mod_BuildFirstRun.bas"
    out.write_text(module, encoding="ascii")
    print(f"wrote {out} ({len(module)} bytes)")

    # Reverse the escaping and compare, so a botched quote cannot ship silently.
    rendered = render_back(codegen)
    original = FORM_CODE.strip("\n")
    ok = rendered == original
    print("round-trip identical to source form code:", ok)
    if not ok:
        import difflib
        for l in difflib.unified_diff(original.split("\n"), rendered.split("\n"), lineterm=""):
            print(l)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
