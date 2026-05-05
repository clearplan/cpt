VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptAgeDates_frm 
   Caption         =   "Age Dates"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "cptAgeDates_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptAgeDates_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.1.1</cpt_version>
Option Explicit

Private Sub cboWeeks_Change()
  'objects
  'strings
  Dim strControlName As String
  'longs
  Dim lngControl As Long
  Dim lngWeeks As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'If Not Me.Visible Or Me.cboWeeks.Value = "" Then GoTo exit_here
  lngWeeks = CLng(Replace(Replace(Me.cboWeeks.Value, "weeks", ""), "week", ""))
  For lngControl = 1 To 10
    Me.Controls("cboWeek" & lngControl).BorderColor = -2147483642
    Me.Controls("cboWeek" & lngControl).BackColor = -2147483643
    strControlName = Me.Controls("cboWeek" & lngControl).Name
    If CLng(Replace(strControlName, "cboWeek", "")) <= lngWeeks Then
      Me.Controls("cboWeek" & lngControl).Enabled = True
      Me.Controls("cboWeek" & lngControl).Locked = False
    Else
      Me.Controls("cboWeek" & lngControl).Value = Null
      Me.Controls("cboWeek" & lngControl).Enabled = False
      Me.Controls("cboWeek" & lngControl).Locked = True
    End If
  Next lngControl

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAgeDates_frm", "cboWeeks_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdRun_Click()
  'objects
  'strings
  'longs
  Dim lngControl As Long
  'integers
  'doubles
  'booleans
  Dim blnMissing As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'save user settings
  cptSaveSetting "AgeDates", "cboWeeks", Me.cboWeeks
  For lngControl = 1 To 10
    Me.Controls("cboWeek" & lngControl).BorderColor = -2147483642
    Me.Controls("cboWeek" & lngControl).BackColor = -2147483643
    If Me.Controls("cboWeek" & lngControl).Enabled Then
      If Me.Controls("cboWeek" & lngControl).Value <> "" Then
        cptSaveSetting "AgeDates", "cboWeek" & lngControl, Me.Controls("cboWeek" & lngControl).Value
      Else
        Me.Controls("cboWeek" & lngControl).BorderColor = 192
        Me.Controls("cboWeek" & lngControl).BackColor = 12648447
        blnMissing = True
      End If
    End If
  Next lngControl
  cptSaveSetting "AgeDates", "chkIncludeDurations", IIf(Me.chkIncludeDurations, 1, 0)
  cptSaveSetting "AgeDates", "chkUpdateCustomFieldNames", IIf(Me.chkUpdateCustomFieldNames, 1, 0)
  
  If Not blnMissing Then cptAgeDates Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAgeDates_frm", "cmdRun_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub lblStatusDate_Click()
  Application.ChangeStatusDate
  Me.lblStatus = "(" & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & ")"
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
