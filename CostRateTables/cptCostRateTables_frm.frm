VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptCostRateTables_frm 
   Caption         =   "CostRateTables"
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4830
   OleObjectBlob   =   "cptCostRateTables_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptCostRateTables_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.1.0</cpt_version>
Option Explicit

Private Sub cboStatusField_AfterUpdate()
  cptSaveSetting "CostRateTables", "cboStatusField", Me.cboStatusField.Value
End Sub

Private Sub cboStatusField_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 27 Then Me.Hide
End Sub

Private Sub chkAddNew_Click()
  cptSaveSetting "CostRateTables", "chkAddNew", CBool(Me.chkAddNew.Value)
End Sub

Private Sub chkOverwrite_Click()
  cptSaveSetting "CostRateTables", "chkOverwrite", CBool(Me.chkOverwrite.Value)
End Sub

Private Sub cmdGo_Click()

  Me.txtCostRateTables.BorderColor = 8421504
  Me.cboStatusField.BorderColor = 8421504
  If Me.tglExport Then
    If Len(Me.txtCostRateTables) = 0 Then
      Me.txtCostRateTables.BorderColor = 192
      Me.txtCostRateTables.SetFocus
    Else
      Call cptExportCostRateTables(Me, Me.txtCostRateTables)
    End If
  ElseIf Me.tglImport Then
    If IsNull(Me.cboStatusField.Value) Or Me.cboStatusField.Value = "" Then
      Me.cboStatusField.BorderColor = 192
      Me.cboStatusField.SetFocus
    Else
      Call cptImportCostRateTables(Me, Me.cboStatusField.Value)
    End If
  End If
  'save user settings
  cptSaveSetting "CostRateTables", "txtCostRateTables", Me.txtCostRateTables.Text
  cptSaveSetting "CostRateTables", "cboStatusField", Me.cboStatusField.Value
  cptSaveSetting "CostRateTables", "chkAddNew", IIf(Me.chkAddNew, "1", "0")
  cptSaveSetting "CostRateTables", "chkOverwrite", IIf(Me.chkOverwrite, "1", "0")

End Sub

Private Sub cmdGo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 27 Then Me.Hide
End Sub

Private Sub tglExport_Click()
  'apply toggle
  Me.tglImport = Not Me.tglExport
  'enable export controls
  Me.lblExportTables.Enabled = Me.tglExport
  Me.txtCostRateTables.Enabled = Me.tglExport
  'disable import controls
  Me.chkOverwrite.Enabled = Me.tglImport
  Me.chkAddNew.Enabled = Me.tglImport
  Me.lblImportField.Enabled = Me.tglImport
  Me.cboStatusField.Enabled = Me.tglImport
  If Me.tglExport Then
    Me.txtCostRateTables.SetFocus
    Me.txtCostRateTables.Text = "A,B,C,D,E"
  End If
End Sub

Private Sub tglExport_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 27 Then Me.Hide
End Sub

Private Sub tglImport_Click()
  'apply toggle
  Me.tglExport = Not Me.tglImport
  'enable export controls
  Me.lblExportTables.Enabled = Me.tglExport
  Me.txtCostRateTables.Enabled = Me.tglExport
  'enable import controls
  Me.chkOverwrite.Enabled = Me.tglImport
  Me.chkAddNew.Enabled = Me.tglImport
  Me.lblImportField.Enabled = Me.tglImport
  Me.cboStatusField.Enabled = Me.tglImport
  If Me.tglImport Then
    Me.cboStatusField.SetFocus
    Me.cboStatusField.DropDown
  End If
  'get saved setting
  Dim strStatusField As String
  strStatusField = cptGetSetting("CostRateTables", "cboStatusField")
  If Len(strStatusField) > 0 Then
    Me.cboStatusField.Value = CLng(strStatusField)
  End If
End Sub

Private Sub tglImport_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 27 Then Me.Hide
End Sub

Private Sub txtCostRateTables_Change()
  Me.txtCostRateTables.Text = cptRegEx(Me.txtCostRateTables.Text, "([A-E],{0,1}){1,5}")
End Sub

Private Sub txtCostRateTables_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 27 Then Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub
