VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptResetAll_frm 
   Caption         =   "How would you like to Reset All?"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6180
   OleObjectBlob   =   "cptResetAll_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptResetAll_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.4.0</cpt_version>
Option Explicit

Private Sub chkKeepPosition_Click()
  cptSaveSetting "ResetAll", "KeepPosition", IIf(Me.chkKeepPosition, "1", "0")
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Sub cmdDoIt_Click()
  'objects
  'strings
  'longs
  Dim lngSettings As Long
  Dim lngOutlineLevel As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True
  
  'capture bitwise value
  If Me.chkActiveOnly Then lngSettings = 1
  If Me.chkGroup Then lngSettings = lngSettings + 2
  If Me.chkSummaries Then lngSettings = lngSettings + 4
  If Me.optShowAllTasks Then
    lngSettings = lngSettings + 8
  ElseIf Me.optOutlineLevel Then
    lngOutlineLevel = Me.cboOutlineLevel
  End If
  If Me.chkSort Then lngSettings = lngSettings + 16
  If Me.chkFilter Then lngSettings = lngSettings + 32
  If Me.chkIndent Then lngSettings = lngSettings + 64
  If Me.chkOutlineSymbols Then lngSettings = lngSettings + 128
  'save settings
  cptSaveSetting "ResetAll", "DefaultView", Me.cboViews.Value
  cptSaveSetting "ResetAll", "Settings", CStr(lngSettings)
  cptSaveSetting "ResetAll", "OutlineLevel", CStr(lngOutlineLevel)
  cptSaveSetting "ResetAll", "KeepPosition", IIf(Me.chkKeepPosition, "1", "0")
  Me.Hide
  'apply
  cptResetAll
  
exit_here:
  On Error Resume Next
  If Me.Visible Then Me.Hide
  cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptResetAll_frm", "cmdDoIt_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub optOutlineLevel_Click()
  Me.cboOutlineLevel.Enabled = True
End Sub

Private Sub optShowAllTasks_Click()
  Me.cboOutlineLevel.Enabled = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub UserForm_Terminate()
  Unload Me
End Sub
