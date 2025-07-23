VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptDynamicFilter_frm 
   Caption         =   "Dynamic Filter"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   OleObjectBlob   =   "cptDynamicFilter_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptDynamicFilter_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.6.4</cpt_version>
Option Explicit

Private Sub cboField_Change()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub cboOperator_Change()
  If Me.Visible Then
    If Me.ActiveControl.Name = "tglRegEx" Then Exit Sub
    Me.txtFilter_Change
  End If
End Sub

Private Sub chkActiveOnly_Click()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub chkHideSummaries_Click()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub chkHighlight_Click()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub chkKeepSelected_Click()
  Dim Task As MSProject.Task
  Dim blnErrorTrapping As Boolean
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Me.chkKeepSelected = True Then
    On Error Resume Next
    Set Task = ActiveSelection.Tasks(1)
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Task Is Nothing Then Me.chkKeepSelected = False
    Set Task = Nothing
  End If
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_frm", "chkKeepSelected_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub chkShowRelatedSummaries_Click()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdClear_Click()
  If ActiveWindow.ActivePane <> ActiveWindow.TopPane Then ActiveWindow.TopPane.Activate
  FilterClear
End Sub

Private Sub cmdGoRegEx_Click()
  Dim strMsg As String, lngResponse As Long
  'todo: this logic doesn't work right
  If cptGetSetting("DynamicFilter", "IgnoreOverwriteWarning") = "" Then
    strMsg = "OK to overwrite the 'Marked' field?" & vbCrLf & vbCrLf
    strMsg = strMsg & "Abort = No it is not OK" & vbCrLf
    strMsg = strMsg & "Retry = Yes this is fine" & vbCrLf
    strMsg = strMsg & "Ignore = Yes and stop bugging me about it"
    lngResponse = MsgBox(strMsg, vbQuestion + vbAbortRetryIgnore, "geeks of the world, unite")
    If lngResponse = vbAbort Then
      'todo: switch to normie view
      GoTo exit_here
    ElseIf lngResponse = vbIgnore Then
      If cptSaveSetting("DynamicFilter", "IgnoreOverwriteWarning", "1") Then Debug.Print "that worked fine"
    End If
  End If
  Call cptGoRegEx(Me, Me.txtFilter)
exit_here:
  Me.txtFilter.SetFocus
End Sub

Private Sub cmdUndo_Click()
  Dim oTask As MSProject.Task
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If oTask.Marked Then oTask.Marked = False
next_task:
  Next oTask
  FilterClear
  Me.txtFilter.SetFocus
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_frm", "cmdUndo_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub lblHelp_Click()
    
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then
    If MsgBox("'regex' stands for regular expresssions, or 'pattern matching.' would you like to see a tutorial?", vbYesNo + vbInformation, "pretty '/^[abds]{6}$/g' stuff") = vbYes Then
      Application.FollowHyperlink ("https://ryanstutorials.net/regular-expressions-tutorial/")
    End If
  Else
    MsgBox "Public Internett access seems to be blocked. Search somewhere else for 'Regular Expression Tutorial'", vbExclamation + vbOKOnly, "sadly..."
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_frm", "lblHelp_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink ("http://" & Me.lblURL.Caption)

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_frm", "lblURL", Err, Erl)
  Resume exit_here

End Sub

Private Sub tglRegEx_Click()
'color pattern credit: https://github.com/julianlatest/material-windows-terminal/blob/master/material-darker.json
'converter credit: https://www.rapidtables.com/convert/color/hex-to-rgb.html

  If Me.tglRegEx Then
    Me.Caption = "dynamicFilter.geekMode() - " & cptGetVersion("cptDynamicFilter_bas")
    Me.BackColor = RGB(33, 33, 33)
    
    Me.cboField.Font.Name = "Consolas"
    Me.cboField.ForeColor = RGB(195, 232, 141) 'RGB(130, 170, 255)
    Me.cboField.BackColor = RGB(84, 84, 84)
    
    Me.cboOperator.Clear
    Me.cboOperator.AddItem "matches"
    Me.cboOperator.Value = "matches"
    Me.cboOperator.Font.Name = "Consolas"
    Me.cboOperator.ForeColor = RGB(195, 232, 141)
    Me.cboOperator.BackColor = RGB(84, 84, 84)
    
    Me.txtFilter.Font.Name = "Consolas"
    Me.txtFilter.ForeColor = RGB(195, 232, 141)
    Me.txtFilter.BackColor = RGB(84, 84, 84)
    Me.txtFilter.Width = 134
    
    'font is always consolas
    Me.cmdGoRegEx.ForeColor = RGB(195, 232, 141)
    Me.cmdGoRegEx.BackColor = RGB(84, 84, 84)
    Me.cmdGoRegEx.Visible = True
    
    Me.cmdUndo.Font.Name = "Consolas"
    Me.cmdUndo.ForeColor = RGB(195, 232, 141)
    Me.cmdUndo.BackColor = RGB(84, 84, 84)
    Me.cmdUndo.Visible = True
    
    Me.lblHelp.ForeColor = RGB(255, 203, 107)
    Me.lblHelp.BackColor = RGB(33, 33, 33)
    Me.lblHelp.Visible = True
    
    Me.chkKeepSelected.Font.Name = "Consolas"
    Me.chkKeepSelected.ForeColor = RGB(195, 232, 141)
    Me.chkKeepSelected.BackColor = RGB(84, 84, 84)
    
    Me.chkHideSummaries.Font.Name = "Consolas"
    Me.chkHideSummaries.ForeColor = RGB(195, 232, 141)
    Me.chkHideSummaries.BackColor = RGB(84, 84, 84)
    
    Me.chkShowRelatedSummaries.Font.Name = "Consolas"
    Me.chkShowRelatedSummaries.ForeColor = RGB(195, 232, 141)
    Me.chkShowRelatedSummaries.BackColor = RGB(84, 84, 84)
    
    Me.chkHighlight.Font.Name = "Consolas"
    Me.chkHighlight.ForeColor = RGB(195, 232, 141)
    Me.chkHighlight.BackColor = RGB(84, 84, 84)
    Me.chkHighlight.Visible = False
    
    Me.chkActiveOnly.Font.Name = "Consolas"
    Me.chkActiveOnly.ForeColor = RGB(195, 232, 141)
    Me.chkActiveOnly.BackColor = RGB(84, 84, 84)
    Me.chkActiveOnly.Left = 12
    
    'font is always consolas
    'Me.tglRegEx.ForeColor = RGB(195, 232, 141)
    'Me.tglRegEx.BackColor = RGB(84, 84, 84)
    
    Me.cmdClear.Font.Name = "Consolas"
    Me.cmdClear.ForeColor = RGB(195, 232, 141)
    Me.cmdClear.BackColor = RGB(84, 84, 84)
    
    Me.cmdCancel.Font.Name = "Consolas"
    Me.cmdCancel.ForeColor = RGB(195, 232, 141)
    Me.cmdCancel.BackColor = RGB(84, 84, 84)
    
    Me.lblURL.ForeColor = RGB(255, 203, 107)
  Else
    Me.Caption = "Dynamic Filter (" & cptGetVersion("cptDynamicFilter_bas") & ")"
    Me.BackColor = -2147483633 'default light grey
    
    Me.cboField.Font.Name = "Tahoma"
    Me.cboField.ForeColor = -2147483640
    Me.cboField.BackColor = -2147483643 'default white
    
    Me.cboOperator.Font.Name = "Tahoma"
    Me.cboOperator.ForeColor = -2147483640
    Me.cboOperator.BackColor = -2147483643
    Me.cboOperator.Clear
    Me.cboOperator.AddItem "equals"
    Me.cboOperator.AddItem "does not equal"
    Me.cboOperator.AddItem "contains"
    Me.cboOperator.AddItem "does not contain"
    Me.cboOperator.Value = "contains"
    
    Me.txtFilter.Font.Name = "Tahoma"
    Me.txtFilter.ForeColor = -2147483640 'black?
    Me.txtFilter.BackColor = -2147483643
    Me.cmdGoRegEx.Visible = False
    Me.txtFilter.Width = 198
    
    Me.chkKeepSelected.Font.Name = "Tahoma"
    Me.chkKeepSelected.ForeColor = -2147483640
    Me.chkKeepSelected.BackColor = -2147483633
    
    Me.chkHideSummaries.Font.Name = "Tahoma"
    Me.chkHideSummaries.ForeColor = -2147483640
    Me.chkHideSummaries.BackColor = -2147483633
    
    Me.chkShowRelatedSummaries.Font.Name = "Tahoma"
    Me.chkShowRelatedSummaries.ForeColor = -2147483640
    Me.chkShowRelatedSummaries.BackColor = -2147483633
    
    Me.chkHighlight.Font.Name = "Tahoma"
    Me.chkHighlight.ForeColor = -2147483640
    Me.chkHighlight.BackColor = -2147483633
    Me.chkHighlight.Visible = True
    
    Me.chkActiveOnly.Font.Name = "Tahoma"
    Me.chkActiveOnly.ForeColor = -2147483640
    Me.chkActiveOnly.BackColor = -2147483633
    Me.chkActiveOnly.Left = 66
    
    Me.tglRegEx.ForeColor = -2147483640
    Me.tglRegEx.BackColor = -2147483633
    
    Me.cmdClear.Font.Name = "Tahoma"
    Me.cmdClear.ForeColor = -2147483640
    Me.cmdClear.BackColor = -2147483633
    
    Me.cmdCancel.Font.Name = "Tahoma"
    Me.cmdCancel.ForeColor = -2147483640
    Me.cmdCancel.BackColor = -2147483633
    
    Me.cmdUndo.Visible = False
    Me.lblHelp.Visible = False
    
    Me.lblURL.ForeColor = 16711680
  End If
  Me.txtFilter.SetFocus
End Sub

Sub txtFilter_Change()
  'strings
  Dim strField As String, strOperator As String, strFilterText As String, strFilter As String
  'booleans
  Dim blnActiveOnly As Boolean
  Dim blnHideSummaryTasks As Boolean, blnHighlight As Boolean, blnKeepSelected As Boolean
  Dim blnShowRelatedSummaries As Boolean
  'longs
  Dim lngOriginalUID As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Me.tglRegEx Then Exit Sub
  If Me.ActiveControl.Name = "cmdClear" Then Exit Sub

  'assign values to variables
  On Error Resume Next
  lngOriginalUID = ActiveSelection.Tasks(1).UniqueID
  strField = Me.cboField
  strOperator = Me.cboOperator
  blnHideSummaryTasks = Not Me.chkHideSummaries
  blnShowRelatedSummaries = Me.chkShowRelatedSummaries
  blnHighlight = Me.chkHighlight
  blnKeepSelected = Me.chkKeepSelected
  If lngOriginalUID = 0 Then blnKeepSelected = False
  If blnHighlight Then
    strFilter = "Dynamic Highlight"
  Else
    strFilter = "Dynamic Filter"
  End If
  strFilterText = Me.txtFilter.Text
  'field Active only accessible in Professional
  If Application.Edition = pjEditionProfessional Then
    blnActiveOnly = Me.chkActiveOnly
  Else
    blnActiveOnly = True
  End If
  
  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Dynamic Filter"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===

  'capture formatting that resembles a field name "[x]" and add a space "[x] "
  If Left(strFilterText, 1) = "[" And Right(strFilterText, 1) = "]" Then strFilterText = strFilterText & " "

  'capture wildcard - not allowed
  If InStr(strFilterText, "*") > 0 Or InStr(strFilterText, "%") > 0 Then
    MsgBox "Wildcards ('*') not allowed.", vbExclamation + vbOKOnly, "Error"
    strFilterText = Replace(strFilterText, "*", "")
    strFilterText = Replace(strFilterText, "%", "")
    Me.txtFilter = strFilterText
    Me.Show False
    Me.txtFilter.SetFocus
    GoTo exit_here
  End If

  cptSpeed True 'speed up

  'build custom filter on the fly and apply it
  If Len(strFilterText) > 0 And Len(strOperator) > 0 Then
    If strField = "Task Name" Then strField = "Name"
    FilterEdit Name:=strFilter, TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:=strField, test:=strOperator, Value:=strFilterText, operation:=IIf(blnKeepSelected Or blnHideSummaryTasks, "Or", "None"), ShowInMenu:=False, ShowSummaryTasks:=blnShowRelatedSummaries
  Else
    If strField = "Task Name" Then strField = "Name"
    FilterEdit Name:=strFilter, TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:=strField, test:="equals", Value:=strFilterText, operation:=IIf(blnKeepSelected Or blnHideSummaryTasks, "Or", "None"), ShowInMenu:=False, ShowSummaryTasks:=blnShowRelatedSummaries
  End If
  If blnKeepSelected Then
    FilterEdit Name:=strFilter, TaskFilter:=True, NewFieldName:="Unique ID", test:="equals", Value:=lngOriginalUID, operation:="Or"
  End If
  If blnHideSummaryTasks Then
    FilterEdit Name:=strFilter, TaskFilter:=True, NewFieldName:="Summary", test:="equals", Value:="No", operation:="And", parenthesis:=blnKeepSelected
    OptionsViewEx DisplaySummaryTasks:=True
  End If

  If Len(strFilterText) > 0 Then
    FilterEdit Name:=strFilter, ShowSummaryTasks:=blnShowRelatedSummaries
    If blnShowRelatedSummaries Then OptionsViewEx DisplaySummaryTasks:=True
  Else
    'build a sterile filter to retain existing autofilters
    FilterEdit Name:=strFilter, TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Summary", test:="equals", Value:="Yes", ShowInMenu:=False, ShowSummaryTasks:=True
    FilterEdit Name:=strFilter, TaskFilter:=True, FieldName:="", NewFieldName:="Summary", test:="equals", Value:="No", operation:="Or", ShowSummaryTasks:=True
  End If
  
  If Application.Edition = pjEditionProfessional Then
    If blnActiveOnly Then
      FilterEdit Name:=strFilter, TaskFilter:=True, parenthesis:=True, NewFieldName:="Active", test:="equals", Value:="Yes", operation:="And"
    End If
  End If
  
  If blnHighlight And Len(strFilterText) = 0 Then
    FilterEdit Name:=strFilter, TaskFilter:=True, parenthesis:=True, FieldName:="", NewFieldName:="Name", test:="equals", Value:="xxx_force-no-match_xxx", operation:="And", ShowSummaryTasks:=True
  End If
  
  FilterClear
  FilterApply strFilter, blnHighlight

  On Error Resume Next
  If lngOriginalUID > 0 And blnKeepSelected Then Application.Find "Unique ID", "equals", lngOriginalUID

exit_here:
  On Error Resume Next
  cptSpeed False 'slow down
  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_frm", "txtFilter_Change", Err, Erl)
  Resume exit_here

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub UserForm_Terminate()
  If Not cptSaveSetting("DynamicFilter", "Operator", Me.cboOperator.Value) Then Debug.Print "Operator not saved."
  If Not cptSaveSetting("DynamicFilter", "KeepSelected", CStr(IIf(Me.chkKeepSelected, "1", "0"))) Then Debug.Print "KeepSelected not saved."
  If Not cptSaveSetting("DynamicFilter", "IncludeSummaries", CStr(IIf(Me.chkHideSummaries, "1", "0"))) Then Debug.Print "IncludeSummaries not saved."
  If Not cptSaveSetting("DynamicFilter", "RelatedSummaries", CStr(IIf(Me.chkShowRelatedSummaries, "1", "0"))) Then Debug.Print "RelatedSummaries not saved."
  If Not cptSaveSetting("DynamicFilter", "Highlight", CStr(IIf(Me.chkHighlight, "1", "0"))) Then Debug.Print "Highlight not saved."
  If Not cptSaveSetting("DynamicFilter", "geekMode", CStr(IIf(Me.tglRegEx, "1", "0"))) Then Debug.Print "geekMode not saved."
  If Not cptSaveSetting("DynamicFilter", "ActiveOnly", CStr(IIf(Me.chkActiveOnly, "1", "0"))) Then Debug.Print "Active Only not saved."
  Unload Me
End Sub
