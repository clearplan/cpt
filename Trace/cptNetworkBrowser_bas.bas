Attribute VB_Name = "cptNetworkBrowser_bas"
'<cpt_version>v1.2.2</cpt_version>
Option Explicit
'=====================================
Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000

#If VBA7 Then
    Public Declare PtrSafe Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function DrawMenuBar _
        Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
#Else
    Public Declare Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
    Public Declare Function DrawMenuBar _
        Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
#End If
'=====================================

Public oSubMap As Scripting.Dictionary

Sub cptResizeWindowSettings(frm As Object, Show As Boolean)

  Dim windowStyle As Long
  Dim windowHandle As Long
  
  'Get the references to window and style position within the Windows memory
  windowHandle = FindWindowA(vbNullString, frm.Caption)
  windowStyle = GetWindowLong(windowHandle, GWL_STYLE)
  
  'Determine the style to apply based
  If Show = False Then
      windowStyle = windowStyle And (Not WS_THICKFRAME)
  Else
      windowStyle = windowStyle + (WS_THICKFRAME)
  End If
  
  'Apply the new style
  SetWindowLong windowHandle, GWL_STYLE, windowStyle
  
  'Recreate the UserForm window with the new style
  DrawMenuBar windowHandle

End Sub

Sub cptShowNetworkBrowser_frm()
  'objects
  Dim myNetworkBrowser_frm As cptNetworkBrowser_frm
  'strings
  Dim strDescending As String
  Dim strSortBy As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
    
  'prevent spawning
  If Not cptGetUserForm("cptNetworkBrowser_frm") Is Nothing Then Exit Sub

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not cptFilterExists("Marked") Then cptCreateFilter ("Marked")
  
  Call cptStartEvents
  Set myNetworkBrowser_frm = New cptNetworkBrowser_frm
  With myNetworkBrowser_frm
    .Caption = "Network Browser (" & cptGetVersion("cptNetworkBrowser_frm") & ")"
    .tglTrace = False
    .tglTrace.Caption = "Jump"
    .lboPredecessors.MultiSelect = fmMultiSelectSingle
    .lboSuccessors.MultiSelect = fmMultiSelectSingle
    With .cboSortPredecessorsBy
      .Clear
      .AddItem "ID"
      .AddItem "Finish"
      .AddItem "Total Slack"
      strSortBy = cptGetSetting("NetworkBrowser", "cboSortPredecessorsBy")
      If Len(strSortBy) > 0 Then
        .Value = strSortBy
      Else
        .Value = "Total Slack"
      End If
    End With
    strDescending = cptGetSetting("NetworkBrowser", "chkSortPredDescending")
    If Len(strDescending) > 0 Then
      .chkSortPredDescending.Value = CBool(strDescending)
    Else
      .chkSortPredDescending.Value = False
    End If
    With .cboSortSuccessorsBy
      .Clear
      .AddItem "ID"
      .AddItem "Start"
      .AddItem "Total Slack"
      strSortBy = cptGetSetting("NetworkBrowser", "cboSortSuccessorsBy")
      If Len(strSortBy) > 0 Then
        .Value = strSortBy
      Else
        .Value = "Total Slack"
      End If
    End With
    strDescending = cptGetSetting("NetworkBrowser", "chkSortSuccDescending")
    If Len(strDescending) > 0 Then
      .chkSortSuccDescending.Value = CBool(strDescending)
    Else
      .chkSortSuccDescending.Value = False
    End If
    cptResizeWindowSettings myNetworkBrowser_frm, True
    .Show False 'VBA.FormShowConstants.vbModeless
    cptShowPreds myNetworkBrowser_frm
  End With

exit_here:
  On Error Resume Next
  Set myNetworkBrowser_frm = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_bas", "cptShowNetworkBrowser_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowPreds(Optional myNetworkBrowser_frm As cptNetworkBrowser_frm)
  'objects
  Dim oTaskDependencies As TaskDependencies
  Dim oSubproject As Subproject
  Dim oLink As TaskDependency, oTask As MSProject.Task
  'strings
  Dim strHideInactive As String
  Dim strProject As String
  'longs
  Dim lngLinkUID As Long
  Dim lngItem As Long
  Dim lngItems As Long
  Dim lngFactor As Long
  Dim lngTasks As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnHideInactive As Boolean
  Dim blnSubprojects As Boolean
  'variants
  Dim vControl As Variant
  'dates
  
  On Error Resume Next
  Set oTask = ActiveSelection.Tasks(1)
  If oTask Is Nothing Then GoTo exit_here
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  lngTasks = ActiveSelection.Tasks.Count
  'determine if there are subprojects loaded (this affects displayed UIDs)
  blnSubprojects = ActiveProject.Subprojects.Count > 0
  
  If blnSubprojects Then
    If oSubMap Is Nothing Then
      Set oSubMap = CreateObject("Scripting.Dictionary")
    Else
      oSubMap.RemoveAll
    End If
    For Each oSubproject In ActiveProject.Subprojects
      If InStr(oSubproject.Path, "<>") = 0 Then 'offline
        oSubMap.Add Replace(Dir(oSubproject.Path), ".mpp", ""), 0
      ElseIf Left(oSubproject.Path, 2) = "<>" Then 'online
        oSubMap.Add Replace(oSubproject.Path, "<>\", ""), 0
      End If
      If oSubproject.IsLoaded = False Then
        Application.OpenUndoTransaction "cpt - load subproject"
        FilterClear
        GroupClear
        SelectAll
        OutlineShowAllTasks
        Application.CloseUndoTransaction
        If Application.GetUndoListCount > 0 Then
          If Application.GetUndoListItem(1) = "cpt - load subproject" Then
            Application.Undo
          End If
        End If
      End If
    Next oSubproject
    For Each oTask In ActiveProject.Tasks
      If oSubMap.Exists(oTask.Project) Then
        If oSubMap(oTask.Project) > 0 Then GoTo next_mapping_task
        oSubMap.Item(oTask.Project) = CLng(oTask.UniqueID / 4194304)
      End If
next_mapping_task:
    Next oTask
  End If
  
  'reset after mapping
  Set oTask = ActiveSelection.Tasks(1)
  If myNetworkBrowser_frm Is Nothing Then Set myNetworkBrowser_frm = New cptNetworkBrowser_frm
  
  With myNetworkBrowser_frm
    If Not .Visible Then .Show False
    Select Case lngTasks
      Case Is < 1
        .lboCurrent.Clear
        .lboPredecessors.Clear
        .lboPredecessors.ColumnCount = 1
        .lboPredecessors.AddItem "Please select a task."
        .lboSuccessors.Clear
        .lboSuccessors.ColumnCount = 1
        .lboSuccessors.AddItem "Please select a task."
        GoTo exit_here
      Case Is > 1
        .lboCurrent.Clear
        .lboPredecessors.Clear
        .lboPredecessors.ColumnCount = 1
        .lboPredecessors.AddItem "Please select only one task."
        .lboSuccessors.Clear
        .lboSuccessors.ColumnCount = 1
        .lboSuccessors.AddItem "Please select only one task."
        GoTo exit_here
    End Select
    If .tglTrace Then
      .tglTrace.Caption = "Trace"
      .lboPredecessors.MultiSelect = fmMultiSelectMulti
      .lboPredecessors.MultiSelect = fmMultiSelectMulti
    Else
      .tglTrace.Caption = "Jump"
      .lboSuccessors.MultiSelect = fmMultiSelectSingle
      .lboSuccessors.MultiSelect = fmMultiSelectSingle
    End If
    With .lboCurrent
      .Clear
      .ColumnCount = 4
      .AddItem
      If blnSubprojects Then
        .ColumnWidths = "50 pt;35 pt;24.95 pt"
      Else
        .ColumnWidths = "24.95 pt;0 pt;24.95 pt"
      End If
      .Column(0, .ListCount - 1) = oTask.UniqueID
      .Column(1, .ListCount - 1) = oTask.UniqueID Mod 4194304
      .Column(2, .ListCount - 1) = oTask.ID
      .Column(3, .ListCount - 1) = IIf(oTask.Marked, "[m] ", "") & oTask.Name
    End With
    strHideInactive = cptGetSetting("NetworkBrowser", "chkHideInactive")
    If Len(strHideInactive) > 0 Then
      .chkHideInactive.Value = CBool(strHideInactive)
    Else
      .chkHideInactive.Value = True 'defaults to true
    End If
    blnHideInactive = .chkHideInactive.Value
  End With
    
  'only 1 is selected
  On Error Resume Next
  Set oTaskDependencies = oTask.TaskDependencies
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTaskDependencies Is Nothing Then
    myNetworkBrowser_frm.lboPredecessors.Clear
    myNetworkBrowser_frm.lboSuccessors.Clear
    GoTo exit_here
  End If
    
  'reset both lbos once in an array here
  For Each vControl In Array("lboPredecessors", "lboSuccessors")
    With myNetworkBrowser_frm.Controls(vControl)
      .Clear
      .ColumnCount = 9
      .AddItem
      If blnSubprojects Then
        .ColumnWidths = "50 pt;35 pt;24.95 pt;24.95 pt;24.95 pt;55 pt;35 pt;225 pt;35 pt"
        .Column(0, .ListCount - 1) = "UID[M]"
        .Column(1, .ListCount - 1) = "UID[S]"
      Else
        .ColumnWidths = "35 pt;0 pt;24.95 pt;24.95 pt;24.95 pt;55 pt;35 pt;225 pt;35 pt"
        .Column(0, .ListCount - 1) = "UID"
      End If
      .Column(2, .ListCount - 1) = "ID"
      .Column(3, .ListCount - 1) = "Type"
      .Column(4, .ListCount - 1) = "Lag"
      .Column(5, .ListCount - 1) = IIf(vControl = "lboPredecessors", "Finish", "Start")
      .Column(6, .ListCount - 1) = "Slack"
      .Column(7, .ListCount - 1) = "Task"
      .Column(8, .ListCount - 1) = "Critical"
    End With
  Next vControl
  
  'capture list of preds with valid native UIDs
  lngItems = oTask.TaskDependencies.Count
  lngItem = 0
  For Each oLink In oTask.TaskDependencies
    'limit to only predecessors
    If oLink.To.Guid = oTask.Guid Then 'it's a predecessor to selected task
      If blnHideInactive And Not oLink.From.Active Then GoTo next_link
      'handle external tasks
      If blnSubprojects And oLink.From.ExternalTask Then
        'fix the returned UID
        lngLinkUID = oLink.From.GetField(185073906) Mod 4194304
        strProject = oLink.From.Project
        If InStr(oLink.From.Project, "\") > 0 Then
          strProject = Replace(strProject, ".mpp", "")
          strProject = Mid(strProject, InStrRev(strProject, "\") + 1)
        End If
        lngFactor = oSubMap(strProject)
        lngLinkUID = (lngFactor * 4194304) + lngLinkUID
      Else
        If blnSubprojects Then
          lngFactor = Round(oTask / 4194304, 0)
          lngLinkUID = (lngFactor * 4194304) + oLink.From.UniqueID
        Else
          lngLinkUID = oLink.From.UniqueID
        End If
      End If
      With myNetworkBrowser_frm.lboPredecessors
        .AddItem
        .Column(0, .ListCount - 1) = lngLinkUID
        .Column(1, .ListCount - 1) = lngLinkUID Mod 4194304
        If blnSubprojects And oLink.From.ExternalTask Then
          .Column(2, .ListCount - 1) = ActiveProject.Tasks.UniqueID(lngLinkUID).ID
          .Column(7, .ListCount - 1) = "<>\" & IIf(ActiveProject.Tasks.UniqueID(lngLinkUID).Marked, "[m] ", "") & IIf(Len(oLink.From.Name) > 65, Left(oLink.From.Name, 65) & "... ", oLink.From.Name)
        ElseIf Not blnSubprojects And oLink.From.ExternalTask Then
          .Column(2, .ListCount - 1) = oLink.From.ID
          .Column(7, .ListCount - 1) = "<>\" & IIf(Len(oLink.From.Name) > 65, Left(oLink.From.Name, 65) & "... ", oLink.From.Name)
        Else
          .Column(2, .ListCount - 1) = oLink.From.ID
          .Column(7, .ListCount - 1) = IIf(ActiveProject.Tasks.UniqueID(lngLinkUID).Marked, "[m] ", "") & IIf(Len(oLink.From.Name) > 65, Left(oLink.From.Name, 65) & "... ", oLink.From.Name)
        End If
        .Column(3, .ListCount - 1) = Choose(oLink.Type + 1, "FF", "FS", "SF", "SS") & IIf(oLink.Type <> pjFinishToStart, "*", "")
        .Column(4, .ListCount - 1) = Round(oLink.Lag / (ActiveProject.HoursPerDay * 60), 2) & "d"
        Select Case oLink.From.ConstraintType
          Case pjFNET
            If oLink.From.Finish > oLink.From.ConstraintDate Then
              .Column(5, .ListCount - 1) = FormatDateTime(oLink.From.Finish, vbShortDate)
            Else
              .Column(5, .ListCount - 1) = "<" & FormatDateTime(oLink.From.Finish, vbShortDate)
            End If
          Case pjFNLT
            If oLink.From.Finish < oLink.From.ConstraintDate Then
              .Column(5, .ListCount - 1) = FormatDateTime(oLink.From.Finish, vbShortDate)
            Else
              .Column(5, .ListCount - 1) = ">" & FormatDateTime(oLink.From.Finish, vbShortDate)
            End If
          Case pjMFO
            If oLink.From.Finish = oLink.From.ConstraintDate Then
              .Column(5, .ListCount - 1) = "=" & FormatDateTime(oLink.From.Finish, vbShortDate)
            Else
              .Column(5, .ListCount - 1) = FormatDateTime(oLink.From.Finish, vbShortDate)
            End If
          Case Else
            .Column(5, .ListCount - 1) = FormatDateTime(oLink.From.Finish, vbShortDate)
        End Select
        .Column(6, .ListCount - 1) = Round(oLink.From.TotalSlack / (ActiveProject.HoursPerDay * 60), 2) & "d"
        .Column(8, .ListCount - 1) = IIf(oLink.From.Critical, "X", "")
      End With
    ElseIf oLink.To.Guid <> oTask.Guid Then 'it's a successor
      If blnHideInactive And Not oLink.From.Active Then GoTo next_link
      'handle external tasks
      If blnSubprojects And oLink.To.ExternalTask Then
        'fix the returned UID
        lngLinkUID = oLink.To.GetField(185073906) Mod 4194304
        strProject = oLink.To.Project
        If InStr(oLink.To.Project, "\") > 0 Then
          strProject = Replace(strProject, ".mpp", "")
          strProject = Mid(strProject, InStrRev(strProject, "\") + 1)
        End If
        lngFactor = oSubMap(strProject)
        lngLinkUID = (lngFactor * 4194304) + lngLinkUID
      Else
        If blnSubprojects Then
          lngFactor = Round(oTask / 4194304, 0)
          lngLinkUID = (lngFactor * 4194304) + oLink.To.UniqueID
        Else
          lngLinkUID = oLink.To.UniqueID
        End If
      End If
      With myNetworkBrowser_frm.lboSuccessors
        .AddItem
        .Column(0, .ListCount - 1) = lngLinkUID
        .Column(1, .ListCount - 1) = lngLinkUID Mod 4194304
        If blnSubprojects And oLink.To.ExternalTask Then
          .Column(2, .ListCount - 1) = ActiveProject.Tasks.UniqueID(lngLinkUID).ID
          .Column(7, .ListCount - 1) = "<>\" & IIf(ActiveProject.Tasks.UniqueID(lngLinkUID).Marked, "[m] ", "") & IIf(Len(oLink.To.Name) > 65, Left(oLink.To.Name, 65) & "... ", oLink.To.Name)
        ElseIf Not blnSubprojects And oLink.To.ExternalTask Then
          .Column(2, .ListCount - 1) = oLink.To.ID
          .Column(7, .ListCount - 1) = "<>\" & IIf(Len(oLink.To.Name) > 65, Left(oLink.To.Name, 65) & "... ", oLink.To.Name)
        Else
          .Column(2, .ListCount - 1) = oLink.To.ID
          .Column(7, .ListCount - 1) = IIf(ActiveProject.Tasks.UniqueID(lngLinkUID).Marked, "[m] ", "") & IIf(Len(oLink.To.Name) > 65, Left(oLink.To.Name, 65) & "... ", oLink.To.Name)
        End If
        .Column(3, .ListCount - 1) = Choose(oLink.Type + 1, "FF", "FS", "SF", "SS") & IIf(oLink.Type <> pjFinishToStart, "*", "")
        .Column(4, .ListCount - 1) = Round(oLink.Lag / (ActiveProject.HoursPerDay * 60), 2) & "d"
        Select Case oLink.To.ConstraintType
          Case pjSNET
            If oLink.To.ConstraintDate > oLink.To.Start Then
              .Column(5, .ListCount - 1) = ">" & FormatDateTime(oLink.To.Start, vbShortDate)
            Else
              .Column(5, .ListCount - 1) = FormatDateTime(oLink.To.Start, vbShortDate)
            End If
          Case pjSNLT
            If oLink.To.ConstraintDate = oLink.To.Start Then
              .Column(5, .ListCount - 1) = "<" & FormatDateTime(oLink.To.Start, vbShortDate)
            Else
              .Column(5, .ListCount - 1) = FormatDateTime(oLink.To.Start, vbShortDate)
            End If
          Case pjMSO
            .Column(5, .ListCount - 1) = "=" & FormatDateTime(oLink.To.Start, vbShortDate)
          Case Else
            .Column(5, .ListCount - 1) = FormatDateTime(oLink.To.Start, vbShortDate)
        End Select
        .Column(6, .ListCount - 1) = Round(oLink.To.TotalSlack / (ActiveProject.HoursPerDay * 60), 2) & "d"
        .Column(8, .ListCount - 1) = IIf(oLink.To.Critical, "X", "")
      End With
    End If
next_link:
    lngItem = lngItem + 1
    myNetworkBrowser_frm.lblPreds.Caption = "Predecessors (" & Format(lngItem / lngItems, "0%") & ")"
    myNetworkBrowser_frm.lblSuccs.Caption = "Successors (" & Format(lngItem / lngItems, "0%") & ")"
    If lngItem = 1 Or lngItems > 300 Then DoEvents
  Next oLink
  
  With myNetworkBrowser_frm
    If .Visible Then
      If .lboPredecessors.ListCount > 2 Then cptSortNetworkBrowserLinks myNetworkBrowser_frm, "p", myNetworkBrowser_frm.chkSortPredDescending.Value
      If .lboSuccessors.ListCount > 2 Then cptSortNetworkBrowserLinks myNetworkBrowser_frm, "s", myNetworkBrowser_frm.chkSortSuccDescending.Value
      If Not oTask Is Nothing Then
        .lblPreds.Caption = "Predecessors: (" & Format(oTask.PredecessorTasks.Count, "#,##0") & ")"
        .lblSuccs.Caption = "Successors: (" & Format(oTask.SuccessorTasks.Count, "#,##0") & ")"
      End If
    Else
      .lblPreds.Caption = "Predecessors:"
      .lblSuccs.Caption = "Successors:"
    End If
  End With
  
exit_here:
  On Error Resume Next
  cptSpeed False
  'Set myNetworkBrowser_frm = Nothing 'do not do this
  Set oTaskDependencies = Nothing
  Set oSubproject = Nothing
  Set oLink = Nothing
  Set oTask = Nothing
  Exit Sub
err_here:
  If Err.Number <> 424 Then Call cptHandleErr("cptNetworkBrowser_bas", "cptShowPreds", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptMarkSelected()
  'todo: separate network browser and make it cptMarkSelected(Optional blnRefilter as Boolean)
  Dim oTask As MSProject.Task, oTasks As MSProject.Tasks
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If Not oTasks Is Nothing Then
    For Each oTask In oTasks
      oTask.Marked = True
    Next oTask
  End If
  If ActiveWindow.TopPane.View.Name = "Network Diagram" Then
    'todo: call cptFilterReapply
    'todo: "Highlight Marked tasks in the current view?"
    cptSpeed True
    FilterApply "All Tasks"
    FilterApply "Marked"
    cptSpeed False
  Else
    'todo
  End If
  Set oTask = Nothing
  Set oTasks = Nothing
End Sub

Sub cptUnmarkSelected(Optional myNetworkBrowser_frm As cptNetworkBrowser_frm)
  'todo: make cptMark(blnMark as Boolean)
  'todo: separate network browser and make it cptUnmarkSelected(Optional blnRefilter as Boolean)
  Dim oTask As MSProject.Task

  cptSpeed True
  For Each oTask In ActiveSelection.Tasks
    If Not oTask Is Nothing Then oTask.Marked = False
  Next oTask
  cptSpeed False
  
  If Not myNetworkBrowser_frm Is Nothing Then
    'todo: from here down from network browser only
    ActiveWindow.TopPane.Activate
    FilterApply "Marked"
    If ActiveWindow.TopPane.View.Name <> "Network Diagram" Then
      SelectAll
      ActiveWindow.BottomPane.Activate
      ViewApply "Network Diagram"
    Else
      'todo: call cptFilterReapply
      cptSpeed True
      FilterApply "All Tasks"
      FilterApply "Marked"
      cptSpeed False
    End If
  End If
  
  Set oTask = Nothing
  'Set myNetworkBrowser_frm = Nothing 'do not do this
End Sub

Sub cptMarked()
  ActiveWindow.TopPane.Activate
  On Error Resume Next
  If Not FilterApply("Marked") Then
    FilterEdit "Marked", True, True, True, , , "Marked", , "equals", "Yes", , True, False
  End If
  FilterApply "Marked"
End Sub

Sub cptClearMarked()
  Dim oTask As MSProject.Task
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSpeed True
  
  'todo: what about master/sub?
  
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    'If Not oTask.Active Then GoTo next_task
    If oTask.Marked Then oTask.Marked = False
next_task:
  Next oTask
  ActiveProject.Tasks.UniqueID(0).Marked = False
  'todo: fix this
  If ActiveWindow.TopPane.View.Name = "Network Diagram" Then
    cptSpeed True
    If Edition = pjEditionProfessional Then
      If Not cptFilterExists("Active Tasks") Then
        FilterEdit Name:="Active Tasks", TaskFilter:=True, Create:=True, OverwriteExisting:=False, FieldName:="Active", test:="equals", Value:="Yes", ShowInMenu:=True, ShowSummaryTasks:=True
      End If
      FilterApply "Active Tasks"
    ElseIf Edition = pjEditionStandard Then
      FilterApply "All Tasks"
    End If
    FilterApply "Marked"
    cptSpeed False
  Else
    'todo: if lower pane
  End If

exit_here:
  On Error Resume Next
  cptSpeed False
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_bas", "cptClearMarked", Err, Erl)
  Resume exit_here
End Sub

Sub cptHistoryDoubleClick(Optional myNetworkBrowser_frm As cptNetworkBrowser_frm)
  Dim lngTaskUID As Long
  Dim blnErrorTrapping As Boolean
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  lngTaskUID = CLng(myNetworkBrowser_frm.lboHistory.Value)
  WindowActivate TopPane:=True
  If IsNumeric(lngTaskUID) Then
    On Error Resume Next
    If Not Find("Unique ID", "equals", lngTaskUID) Then
      If ActiveWindow.TopPane.View.Name = "Network Diagram" Then
        ActiveProject.Tasks.UniqueID(lngTaskUID).Marked = True
        FilterApply "Marked"
        GoTo exit_here
      End If
      If MsgBox("Task is hidden - remove filters and show it?", vbQuestion + vbYesNo, "Confirm Apocalypse") = vbYes Then
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
          MsgBox "Unable to find Task UID " & lngTaskUID & "...", vbExclamation + vbOKOnly, "Task Not Found"
        End If
      End If
    End If
  End If
  
exit_here:
  'Set myNetworkBrowser_frm = Nothing 'do not do this
  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_bas", "cptHistoryDoubleClick", Err, Erl)
  Resume exit_here
End Sub

Sub cptSortNetworkBrowserLinks(ByRef myNetworkBrowser_frm As cptNetworkBrowser_frm, strWhich As String, Optional blnDescending = False)
  'objects
  Dim oComboBox As Object 'MSForms.ComboBox
  Dim oListBox As Object 'MSForms.ListBox
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strIndicator As String
  Dim strUID As String
  Dim strSortBy As String
  'longs
  Dim lngUID As Long
  Dim lngCol As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If strWhich = "p" Then
    Set oListBox = myNetworkBrowser_frm.lboPredecessors
    Set oComboBox = myNetworkBrowser_frm.cboSortPredecessorsBy
  ElseIf strWhich = "s" Then
    Set oListBox = myNetworkBrowser_frm.lboSuccessors
    Set oComboBox = myNetworkBrowser_frm.cboSortSuccessorsBy
  End If

  If oListBox.ListCount <= 2 Then GoTo exit_here

  Set oRecordset = CreateObject("ADODB.Recordset")
  'UID,ID,Type,Lag,Date,Slack,Task,Critical
  With oRecordset
    .Fields.Append "UID_M", adInteger
    .Fields.Append "UID_S", adInteger
    .Fields.Append "ID", adInteger
    .Fields.Append "Type", adVarChar, 3
    .Fields.Append "Lag", adVarChar, 255
    .Fields.Append "Date", adDate
    .Fields.Append "Slack", adInteger
    .Fields.Append "Task", adVarChar, 255
    .Fields.Append "Critical", adBoolean
    .Fields.Append "indicator", adVarChar, 1
    .Open
    For lngItem = oListBox.ListCount - 1 To 1 Step -1
      .AddNew
      For lngCol = 0 To oListBox.ColumnCount - 1
        If .Fields(lngCol).Name = "Slack" Then
          .Fields(lngCol) = CInt(Replace(oListBox.List(lngItem, lngCol), "d", ""))
        ElseIf .Fields(lngCol).Name = "Critical" Then
          If IsNull(oListBox.List(lngItem, lngCol)) Then
            .Fields(lngCol) = False
          Else
            .Fields(lngCol) = True
          End If
        ElseIf .Fields(lngCol).Name = "Date" Then
          If Len(cptRegEx(oListBox.List(lngItem, lngCol), "<|>|=")) > 0 Then
            strIndicator = Left(oListBox.List(lngItem, lngCol), 1)
            'indicates a constraint on the date
            '< = SNET
            '> = FNLT
            '= = MSO/MFO
            .Fields(lngCol) = Replace(oListBox.List(lngItem, lngCol), strIndicator, "")
            .Fields("indicator") = strIndicator
          Else
            .Fields(lngCol) = oListBox.List(lngItem, lngCol)
          End If
        Else
          .Fields(lngCol) = oListBox.List(lngItem, lngCol)
        End If
      Next lngCol
      oListBox.RemoveItem lngItem
    Next lngItem
    strSortBy = oComboBox.Value
    If strSortBy = "Start" Or strSortBy = "Finish" Then strSortBy = "Date"
    If strSortBy = "Total Slack" Then strSortBy = "Slack"
    .Sort = strSortBy & IIf(blnDescending, " desc", "")
    .MoveFirst
    Do While Not .EOF
      oListBox.AddItem
      For lngCol = 0 To .Fields.Count - 2
        If .Fields(lngCol).Name = "Slack" Then
          oListBox.List(oListBox.ListCount - 1, lngCol) = .Fields(lngCol) & "d"
        ElseIf .Fields(lngCol).Name = "Critical" Then
          If .Fields(lngCol) Then
            oListBox.List(oListBox.ListCount - 1, lngCol) = "X"
          End If
        ElseIf .Fields(lngCol).Name = "Date" And Not IsNull(.Fields("indicator")) Then
          oListBox.List(oListBox.ListCount - 1, lngCol) = .Fields("indicator") & .Fields(lngCol)
        Else
          oListBox.List(oListBox.ListCount - 1, lngCol) = .Fields(lngCol)
        End If
      Next lngCol
      .MoveNext
    Loop
    .Close
  End With

exit_here:
  On Error Resume Next
  'Set myNetworkBrowser_frm = Nothing 'do not do this
  Set oComboBox = Nothing
  Set oListBox = Nothing
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_bas", "cptSortNetworkBrowserLinks", Err, Erl)
  Resume exit_here
End Sub
