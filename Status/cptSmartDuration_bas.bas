Attribute VB_Name = "cptSmartDuration_bas"
'<cpt_version>v2.1.0</cpt_version>

Sub cptShowSmartDuration_frm()
  'objects
  Dim mySmartDuration_frm As cptSmartDuration_frm
  'strings
  Dim strSetting As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set mySmartDuration_frm = New cptSmartDuration_frm
  cptUpdateSmartDurationForm mySmartDuration_frm
  With mySmartDuration_frm
    .Caption = "Smart Duration (" & cptGetVersion("cptSmartDuration_frm") & ")"
    strSetting = cptGetSetting("SmartDuration", "chkMarkOnTrack")
    If Len(strSetting) = 0 Then
      .chkMarkOnTrack = False 'default to false
    Else
      .chkMarkOnTrack = CBool(strSetting)
    End If
    strSetting = cptGetSetting("SmartDuration", "chkRetainETC")
    If Len(strSetting) = 0 Then
      .chkRetainETC = True 'default to true
    Else
      .chkRetainETC = CBool(strSetting)
    End If
    strSetting = cptGetSetting("SmartDuration", "chkKeepOpen")
    If Len(strSetting) = 0 Then
      .chkKeepOpen = False 'default to false
    Else
      .chkKeepOpen = CBool(strSetting)
    End If
    .Show False
    If .txtTargetFinish.Enabled Then .txtTargetFinish.SetFocus
  End With
  
  cptCore_bas.cptStartEvents

exit_here:
  On Error Resume Next
  Set mySmartDuration_frm = Nothing
  
  Exit Sub

err_here:
  Call cptHandleErr("cptSmartDuration_bas", "cptShowSmartDuration_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateSmartDurationForm(ByRef mySmartDuration_frm As cptSmartDuration_frm)
  'objects
  Dim oTasks As MSProject.Tasks
  'strings
  'longs
  'integers
  'doubles
  'booleans
  Dim blnValid As Boolean
  'variants
  'dates
  
  On Error Resume Next
  Set oTasks = Nothing
  Set oTasks = ActiveSelection.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnValid = True
  With mySmartDuration_frm
    If oTasks Is Nothing Then
      blnValid = False
    ElseIf oTasks.Count = 0 Then 'Group By Summary
      .txtTargetFinish = "-"
      .lblWeekday.Caption = "< invalid >"
      .lblWeekday.ControlTipText = "Cannot adjust Group By Summary tasks."
      blnValid = False
    ElseIf oTasks(1) Is Nothing Then 'newly inserted task
      .txtTargetFinish = "-"
      .lblWeekday.Caption = "< invalid >"
      .lblWeekday.ControlTipText = "Cannot adjust newly inserted tasks."
      blnValid = False
    ElseIf oTasks.Count > 1 Then 'too many
      .txtTargetFinish = ""
      .lblWeekday.Caption = "< focus >"
      .lblWeekday.ControlTipText = "Please select a single task."
      blnValid = False
    ElseIf oTasks(1).Summary Then
      .txtTargetFinish = ""
      .lblWeekday.Caption = "< summary >"
      .lblWeekday.ControlTipText = "Please select a Non-summary task."
      blnValid = False
    ElseIf Not oTasks(1).Active Then
      .txtTargetFinish = ""
      .lblWeekday.Caption = "< inactive >"
      .lblWeekday.ControlTipText = "Please select an Active task."
      blnValid = False
    ElseIf IsDate(oTasks(1).ActualFinish) Then
      .txtTargetFinish = ""
      .lblWeekday.Caption = "< complete >"
      .lblWeekday.ControlTipText = "Please select an incomplete task."
      blnValid = False
    End If
  
    If blnValid Then
      .lngUID = oTasks(1).UniqueID
      .StartDate = oTasks(1).Start
      .txtTargetFinish = FormatDateTime(oTasks(1).Finish, vbShortDate)
      .lblWeekday.Caption = Format(.txtTargetFinish.Text, "dddd")
      .lblWeekday.ControlTipText = ""
      '.txtTargetFinish.SetFocus 'this steals focus when user may not want it to
    End If
    .txtTargetFinish.Enabled = blnValid
    .chkMarkOnTrack.Enabled = blnValid
    .chkRetainETC.Enabled = blnValid
    .cmdApply.Enabled = blnValid
  End With

exit_here:
  On Error Resume Next
  Set oTasks = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSmartDuration_bas", "cptUpdateSmartDurationForm", Err, Erl)
  Resume exit_here
End Sub

