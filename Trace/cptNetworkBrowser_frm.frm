VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptNetworkBrowser_frm 
   Caption         =   "Network Browser"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885.001
   OleObjectBlob   =   "cptNetworkBrowser_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptNetworkBrowser_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.2.2</cpt_version>
Option Explicit

Private Sub cboSortPredecessorsBy_Change()
  If Me.Visible Then
    cptSaveSetting "NetworkBrowser", "cboSortPredecessorsBy", Me.cboSortPredecessorsBy.Value
    cptSortNetworkBrowserLinks Me, "p", Me.chkSortPredDescending.Value
  End If
End Sub

Private Sub cboSortSuccessorsBy_Change()
  If Me.Visible Then
    cptSaveSetting "NetworkBrowser", "cboSortSuccessorsBy", Me.cboSortSuccessorsBy.Value
    cptSortNetworkBrowserLinks Me, "s", Me.chkSortSuccDescending.Value
  End If
End Sub

Private Sub chkHideInactive_Click()
  If Me.Visible Then
    cptSaveSetting "NetworkBrowser", "chkHideInactive", IIf(Me.chkHideInactive, 1, 0)
    cptShowPreds Me
  End If
End Sub

Private Sub chkSortPredDescending_Click()
  If Me.Visible Then
    cptSaveSetting "NetworkBrowser", "chkSortPredDescending", IIf(Me.chkSortPredDescending, "1", "0")
    cptSortNetworkBrowserLinks Me, "p", Me.chkSortPredDescending.Value
  End If
End Sub

Private Sub chkSortSuccDescending_Click()
  If Me.Visible Then
    cptSaveSetting "NetworkBrowser", "chkSortSuccDescending", IIf(Me.chkSortSuccDescending, "1", "0")
    cptSortNetworkBrowserLinks Me, "s", Me.chkSortSuccDescending.Value
  End If
End Sub

Private Sub cmdBack_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.lboHistory.SetFocus
  
  If IsNull(Me.lboHistory.Value) Then Me.lboHistory.ListIndex = -1

  If Me.lboHistory.ListCount > 1 Then
    If Me.lboHistory.ListIndex < Me.lboHistory.ListCount - 1 Then
      Me.lboHistory.ListIndex = Me.lboHistory.ListIndex + 1
      cptHistoryDoubleClick Me
    End If
  End If

exit_here:
  Exit Sub
err_here:
  If Err.Number = 380 Then
    Err.Clear
  Else
    Call cptHandleErr("cptNetworkBrowser_frm", "cmdBack_Click", Err, Erl)
  End If
  Resume exit_here
  
End Sub

Private Sub cmdClearHistory_Click()
  If Me.lboHistory.ListCount > 0 Then
    If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Confirm") = vbYes Then Me.lboHistory.Clear
  End If
End Sub

Private Sub cmdClose_Click()
  Set oSubMap = Nothing
  Me.Hide
End Sub

Private Sub cmdFwd_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.lboHistory.SetFocus
  
  If IsNull(Me.lboHistory.Value) Then Me.lboHistory.ListIndex = 0

  If Me.lboHistory.ListCount > 0 And Me.lboHistory.ListIndex > 0 Then
    Me.lboHistory.ListIndex = Me.lboHistory.ListIndex - 1
    cptHistoryDoubleClick Me
  End If

exit_here:
  Exit Sub
err_here:
  If Err.Number = 380 Then
    Err.Clear
  Else
    Call cptHandleErr("cptNetworkBrowser_frm", "cmdFwd_Click", Err, Erl)
  End If
  Resume exit_here

End Sub

Private Sub cmdMark_Click()
  'objects
  Dim oTask As MSProject.Task
  'strings
  'longs
  Dim lngUID As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'cptSpeed True
  If ActiveSelection.Tasks.Count = 1 Then
    lngUID = ActiveSelection.Tasks(1).UniqueID
    If Not ActiveSelection.Tasks(1).Marked Then ActiveSelection.Tasks(1).Marked = True
    For lngItem = 0 To Me.lboPredecessors.ListCount - 1
      If Me.lboPredecessors.Selected(lngItem) Then
        If Me.lboPredecessors.Column(0, lngItem) = "UID" Then GoTo exit_here
        Set oTask = ActiveProject.Tasks.UniqueID(CLng(Me.lboPredecessors.Column(0, lngItem)))
        If Not oTask.ExternalTask Then
          oTask.Marked = True
        Else
          MsgBox "Cannot Mark an External Task", vbExclamation + vbOKOnly, "Unavailable"
          GoTo exit_here
        End If
      End If
    Next lngItem
    For lngItem = 0 To Me.lboSuccessors.ListCount - 1
      If Me.lboSuccessors.Selected(lngItem) Then
        If Me.lboSuccessors.Column(0, lngItem) = "UID" Then GoTo exit_here
        Set oTask = ActiveProject.Tasks.UniqueID(CLng(Me.lboSuccessors.Column(0, lngItem)))
        If Not oTask.ExternalTask Then
          oTask.Marked = True
        Else
          MsgBox "Cannot Mark an External Task", vbExclamation + vbOKOnly, "Unavailable"
          GoTo exit_here
        End If
      End If
    Next lngItem
  Else
    MsgBox "Please select only one task.", vbInformation + vbOKOnly, "Error"
    Exit Sub
  End If
  ActiveWindow.TopPane.Activate
  If Not cptFilterExists("Marked") Then cptCreateFilter ("Marked")
  FilterApply "Marked"
  If ActiveWindow.TopPane.View.Name <> "Network Diagram" Then
    Sort "Start", True, "Duration", True
    SelectAll
    If ActiveWindow.BottomPane Is Nothing Then
      Application.DetailsPaneToggle False
    End If
    ActiveWindow.BottomPane.Activate
    ViewApply "Network Diagram"
  End If
  'Find "Unique ID", "equals", lngUID
exit_here:
  On Error Resume Next
  Set oTask = Nothing
  cptSpeed False
  Exit Sub
err_here:
  On Error Resume Next
  Call cptHandleErr("cptNetworkBrowser_frm", "cmdMark_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdUnmark_Click()
  'objects
  Dim oTask As MSProject.Task
  'strings
  'longs
  Dim lngUID As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'cptSpeed True
  If ActiveSelection.Tasks.Count = 1 Then
    lngUID = ActiveSelection.Tasks(1).UniqueID
    For lngItem = 0 To Me.lboPredecessors.ListCount - 1
      If Me.lboPredecessors.Selected(lngItem) Then
         If Me.lboPredecessors.Column(0, lngItem) = "UID" Then GoTo exit_here
        Set oTask = ActiveProject.Tasks.UniqueID(CLng(Me.lboPredecessors.Column(0, lngItem)))
        If Not oTask.ExternalTask Then
          oTask.Marked = False
        Else
          MsgBox "Cannot Mark an External Task", vbExclamation + vbOKOnly, "Unavailable"
          GoTo exit_here
        End If
      End If
    Next lngItem
    For lngItem = 0 To Me.lboSuccessors.ListCount - 1
      If Me.lboSuccessors.Selected(lngItem) Then
        If Me.lboSuccessors.Column(0, lngItem) = "UID" Then GoTo exit_here
        Set oTask = ActiveProject.Tasks.UniqueID(CLng(Me.lboSuccessors.Column(0, lngItem)))
        If Not oTask.ExternalTask Then
          oTask.Marked = False
        Else
          MsgBox "Cannot Mark an External Task", vbExclamation + vbOKOnly, "Unavailable"
          GoTo exit_here
        End If
      End If
    Next lngItem
  Else
    MsgBox "Please select only one task.", vbInformation + vbOKOnly, "Error"
    Exit Sub
  End If
  ActiveWindow.TopPane.Activate
  If Not cptFilterExists("Marked") Then cptCreateFilter ("Marked")
  FilterApply "Marked"
  If ActiveWindow.TopPane.View.Name <> "Network Diagram" Then
    Sort "Start", True, "Duration", True
    SelectAll
    If ActiveWindow.BottomPane Is Nothing Then
      Application.DetailsPaneToggle False
    End If
    ActiveWindow.BottomPane.Activate
    ViewApply "Network Diagram"
  End If
  'Find "Unique ID", "equals", lngUID
    
exit_here:
  On Error Resume Next
  Set oTask = Nothing
  cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_frm", "cmdUnmark_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdUnmarkAll_Click()
  Dim oTask As MSProject.Task

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True
  ActiveWindow.BottomPane.Activate
  On Error Resume Next
  If ActiveSelection.Tasks.Count = 0 Then Exit Sub
  For Each oTask In ActiveSelection.Tasks
    oTask.Marked = False
  Next oTask
  If Not cptFilterExists("Marked") Then cptCreateFilter ("Marked")
  FilterApply "Marked"
  If ActiveWindow.BottomPane Is Nothing Then
    Application.DetailsPaneToggle False
  End If
  ActiveWindow.TopPane.Activate
  Sort "Start", True, "Duration", True
  SelectAll
  ActiveWindow.BottomPane.Activate

exit_here:
  On Error Resume Next
  cptSpeed False

  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_frm", "cmdUnmarkAll_Click", Err, Erl)
  Resume exit_here
  
End Sub

Sub lboHistory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  cptHistoryDoubleClick Me
End Sub

Sub lboPredecessors_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim lngTaskUID As Long
  Dim blnErrorTrapping As Boolean
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Me.lboPredecessors.ListIndex <= 0 Then GoTo exit_here
  With Me.lboHistory
    .AddItem ActiveSelection.Tasks.Item(1).UniqueID, 0
  End With
  lngTaskUID = CLng(Me.lboPredecessors.List(Me.lboPredecessors.ListIndex, 0))
  If lngTaskUID > 0 Then
    WindowActivate TopPane:=True
    On Error Resume Next
    If Not Find("Unique ID", "equals", lngTaskUID) Then
      If ActiveWindow.TopPane.View.Name = "Network Diagram" Then GoTo exit_here
      If MsgBox("Task is currently hidden - remove filters and show it?", vbQuestion + vbYesNo, "Confirm Apocalypse") = vbYes Then
        FilterClear
        OptionsViewEx DisplaySummaryTasks:=True
        On Error Resume Next
        If Not OutlineShowAllTasks Then
          If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
            Sort "ID", , , , , , False, True
            OutlineShowAllTasks
          Else
            SelectBeginning
            GoTo exit_here
          End If
        End If
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If Not Find("Unique ID", "equals", lngTaskUID) Then
          MsgBox "Task not found.", vbExclamation + vbOKOnly, "Missing Task?"
        End If
      Else
        GoTo exit_here
      End If
    End If
    Me.lboHistory.AddItem lngTaskUID, 0
    Me.lboHistory.ListIndex = Me.lboHistory.TopIndex
    cptShowPreds Me
  End If
  
exit_here:
  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_frm", "lboPredecesors_DblClick", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboSuccessors_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim lngTaskUID As Long, oTask As MSProject.Task

  On Error Resume Next
  If Me.lboSuccessors.ListIndex <= 0 Then GoTo exit_here
  Set oTask = ActiveSelection.Tasks(1)

  On Error GoTo err_here
  
  With Me.lboHistory
    If Not oTask Is Nothing Then
      If Me.lboHistory.ListCount > 0 Then
        If Me.lboHistory.List(0, 0) <> oTask.UniqueID Then .AddItem oTask.UniqueID, 0
      Else
        .AddItem oTask.UniqueID, 0
      End If
    End If
  End With
  lngTaskUID = CLng(Me.lboSuccessors.List(Me.lboSuccessors.ListIndex, 0))
  WindowActivate TopPane:=True
  On Error Resume Next
  If Not Find("Unique ID", "equals", lngTaskUID) Then
    If ActiveWindow.TopPane.View.Name = "Network Diagram" Then
      ActiveProject.Tasks(lngTaskUID).Marked = True
      FilterApply "Marked"
      GoTo exit_here
    End If
    If MsgBox("Task may be hidden - remove filters and show it?", vbQuestion + vbYesNo, "Please Confirm") = vbYes Then
      FilterClear
      OptionsViewEx DisplaySummaryTasks:=True
      On Error Resume Next
      If Not OutlineShowAllTasks Then
        If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
          Sort "ID", , , , , , False, True
          OutlineShowAllTasks
        Else
          SelectBeginning
          GoTo exit_here
        End If
      End If
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If Not Find("Unique ID", "equals", lngTaskUID) Then
        MsgBox "Task not found.", vbExclamation + vbOKOnly, "Missing Task?"
      End If
    End If
  Else
    GoTo exit_here
  End If
  Me.lboHistory.AddItem lngTaskUID, 0
  Me.lboHistory.ListIndex = Me.lboHistory.TopIndex
  cptShowPreds Me
  
exit_here:
  On Error Resume Next
  Set oTask = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_frm", "lboSuccessors_DblClick", Err, Erl)
  Resume exit_here
End Sub

Private Sub tglTrace_Click()
  If Not Me.tglTrace Then
    Me.tglTrace.Caption = "Jump"
    Me.cmdMark.Enabled = False
    Me.cmdUnmark.Enabled = False
    Me.lboPredecessors.MultiSelect = fmMultiSelectSingle
    Me.lboPredecessors.ControlTipText = "Double-click to Jump"
    Me.lboSuccessors.MultiSelect = fmMultiSelectSingle
    Me.lboSuccessors.ControlTipText = "Double-click to Jump"
  Else
    Me.tglTrace.Caption = "Trace"
    Me.cmdMark.Enabled = True
    Me.cmdUnmark.Enabled = True
    Me.lboPredecessors.MultiSelect = fmMultiSelectMulti
    Me.lboPredecessors.ControlTipText = "Select and then Mark/Unmark"
    Me.lboSuccessors.MultiSelect = fmMultiSelectMulti
    Me.lboSuccessors.ControlTipText = "Select and then Mark/Unmark"
  End If
End Sub

Private Sub UserForm_Initialize()
  
  Call cptResizeWindowSettings(Me, True)
  
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call cptCore_bas.cptStartEvents
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub UserForm_Resize()
  
  If Me.Width < 506.25 Then 'protect minimum width
    Me.Width = 506.25
  Else
    Me.lboPredecessors.Width = 416 + (Me.Width - 506.25)
    Me.lboSuccessors.Width = 416 + (Me.Width - 506.25)
  End If
  If Me.Height <> 345.75 Then 'protect height (for now)
    Me.Height = 345.75
  End If
  
End Sub

Private Sub UserForm_Terminate()
  Unload Me
End Sub
