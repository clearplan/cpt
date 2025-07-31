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
'<cpt_version>v1.3.1</cpt_version>
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
  Dim strFilter As String
  Dim strFile As String
  'longs
  Dim lngSettings As Long
  Dim lngOutlineLevel As Long
  Dim lngLevel As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnApplyOutlineLevel As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True
  
  'first apply the view
  If Me.cboViews.Value <> "<None>" Then
    If cptViewExists(Me.cboViews.Value) Then
      ViewApply Me.cboViews.Value
    End If
  End If
  cptSaveSetting "ResetAll", "DefaultView", Me.cboViews.Value
  
  'capture bitwise value
  If Me.chkActiveOnly Then
    If Edition = pjEditionProfessional Then SetAutoFilter "Active", pjAutoFilterFlagYes
    lngSettings = 1
  End If
  If Me.chkGroup Then
    GroupClear
    lngSettings = lngSettings + 2
  End If
  If Me.chkSummaries Then
    OptionsViewEx DisplaySummaryTasks:=True
    lngSettings = lngSettings + 4
  End If
  'outline options
  OptionsViewEx DisplaySummaryTasks:=True
  On Error Resume Next
  blnApplyOutlineLevel = True
  If Not OutlineShowAllTasks Then
    If Not Me.chkSort Then
      If MsgBox("Outline Structure must be retained in order to expand all tasks. OK to re-sort?", vbExclamation + vbYesNo, "Sort Conflict") = vbYes Then
        Sort "ID", , , , , , False, True
        OutlineShowAllTasks
        blnApplyOutlineLevel = True
      Else
        MsgBox "Cannot apply Outline Level option until Sort includes 'Retain Outline Structure' option.", vbInformation + vbOKOnly, "Sort Conflict"
        blnApplyOutlineLevel = False
      End If
    Else
      Sort "ID", , , , , , False, True
      OutlineShowAllTasks
      blnApplyOutlineLevel = True
    End If
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Me.optShowAllTasks Then
    If ActiveProject.Subprojects.Count > 0 Then
      OptionsViewEx DisplaySummaryTasks:=True
      If Not Me.chkFilter Then
        strFilter = ActiveProject.CurrentFilter
      End If
      FilterClear
      SelectAll
      OutlineShowAllTasks
      If Not Me.chkSummaries Then OptionsViewEx DisplaySummaryTasks:=False
      If Len(strFilter) > 0 Then FilterApply strFilter
    End If
    If Not Me.chkSummaries Then
      OptionsViewEx DisplaySummaryTasks:=False
    End If
    lngSettings = lngSettings + 8
  ElseIf Me.optOutlineLevel Then
    lngOutlineLevel = Me.cboOutlineLevel
    If blnApplyOutlineLevel Then
      OutlineShowTasks pjTaskOutlineShowLevelMax
      For lngLevel = 20 To lngOutlineLevel Step -1
        OutlineShowTasks lngLevel
      Next lngLevel
    End If
  End If
  If Me.chkSort Then
    Sort "ID", , , , , , False, True
    lngSettings = lngSettings + 16
  End If
  If Me.chkFilter Then
    FilterClear
    If Me.chkActiveOnly And Edition = pjEditionProfessional Then SetAutoFilter "Active", pjAutoFilterFlagYes
    lngSettings = lngSettings + 32
  End If
  If Me.chkIndent Then
    OptionsViewEx DisplayNameIndent:=True
    lngSettings = lngSettings + 64
  End If
  If Me.chkOutlineSymbols Then
    OptionsViewEx displayoutlinesymbols:=True
    lngSettings = lngSettings + 128
  End If
  cptSaveSetting "ResetAll", "Settings", CStr(lngSettings)
  cptSaveSetting "ResetAll", "OutlineLevel", CStr(lngOutlineLevel)

exit_here:
  On Error Resume Next
  cptSpeed False
  Me.Hide
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
