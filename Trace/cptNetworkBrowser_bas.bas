Attribute VB_Name = "cptNetworkBrowser_bas"
'<cpt_version>v1.2.4</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptNetworkBrowser_bas"
'=====================================
Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000
#If VBA7 Then
    Public Declare PtrSafe Function cptGetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function cptSetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function cptDrawMenuBar _
        Lib "user32" Alias "DrawMenuBar" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function cptFindWindow _
        Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
#Else
    Public Declare Function cptGetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function cptSetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
    Public Declare Function cptDrawMenuBar _
        Lib "user32" Alias "DrawMenuBar" (ByVal hWnd As Long) As Long
    Public Declare Function cptFindWindow _
        Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
#End If
'=====================================
Public oSubMap As Scripting.Dictionary

Sub cptResizeWindowSettings(frm As Object, Show As Boolean)

  Dim windowStyle As Long
  Dim windowHandle As Long
  
  'Get the references to window and style position within the Windows memory
  windowHandle = cptFindWindow(vbNullString, frm.Caption)
  windowStyle = cptGetWindowLong(windowHandle, GWL_STYLE)
  
  'Determine the style to apply based
  If Show = False Then
      windowStyle = windowStyle And (Not WS_THICKFRAME)
  Else
      windowStyle = windowStyle + (WS_THICKFRAME)
  End If
  
  'Apply the new style
  cptSetWindowLong windowHandle, GWL_STYLE, windowStyle
  
  'Recreate the UserForm window with the new style
  cptDrawMenuBar windowHandle

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
    .Caption = "Network Browser (" & cptGetVersion("cptNetworkBrowser_frm") & ") - ClearPlan Toolbar"
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
  Dim oSubproject As SubProject
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
      If Left(oSubproject.Path, 2) = "<>" Then 'PWA
        oSubMap.Add Replace(oSubproject.Path, "<>\", ""), 0
      Else 'mpp (local or remote)
        oSubMap.Add Replace(cptRegEx(oSubproject.Path, "[^\\/]*.mpp$"), ".mpp", ""), 0
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
        If Left(strProject, 2) = "<>" Then
          strProject = Replace(strProject, "<>\", "")
        Else
          strProject = Replace(cptRegEx(strProject, "[^\\/]*.mpp$"), ".mpp", "")
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
        'todo: TrueFloat
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
        If Left(strProject, 2) = "<>" Then
          strProject = Replace(strProject, "<>\", "")
        Else
          strProject = Replace(cptRegEx(strProject, "[^\\/]*.mpp$"), ".mpp", "")
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
        'todo: TrueFloat
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
        FilterEdit Name:="Active Tasks", TaskFilter:=True, Create:=True, OverwriteExisting:=False, FieldName:="Active", Test:="equals", Value:="Yes", ShowInMenu:=True, ShowSummaryTasks:=True
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

Sub cptExportCrossProjectLinks()
  'objects
  Dim oTaskMap As Scripting.Dictionary
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oSubproject As MSProject.SubProject
  Dim oTask As MSProject.Task
  Dim oFrom As MSProject.Task
  Dim oTo As MSProject.Task
  Dim oPred As MSProject.Task
  Dim oLink As MSProject.TaskDependency
  Dim oCodeModule As VBIDE.CodeModule
  'longs
  Dim lngCount As Long
  Dim lngFactor As Long
  Dim lngSourceUID As Long
  Dim lngMasterUID As Long
  Dim lngPUID As Long
  Dim lngTask As Long
  Dim lngTaskCount As Long
  'strings
  Dim strProject As String
  Dim strProjectUID As String
  Dim strCode As String
  Dim strFromPUID As String
  Dim strToPUID As String
  'variants
  Dim vCol As Variant
  Dim vCPL() As Variant
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnMaster As Boolean
  Const CHUNK_SIZE As Long = 1000
  
  blnMaster = ActiveProject.Subprojects.Count > 0
  If Not blnMaster Then
    MsgBox "This project has no subprojects.", vbExclamation + vbOKOnly, "Export CPLs"
    GoTo exit_here
  End If
  
  strProjectUID = "PUID" 'use listbox?
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oSubMap Is Nothing Then
    Application.StatusBar = "Building SubMap..."
    cptGetSubMap
    Application.StatusBar = ""
  End If
  
  cptSpeed True
  
  'todo: add to Master Toolset on Ribbon
  ActiveWindow.TopPane.Activate
  OptionsViewEx DisplayNameIndent:=True, DisplaySummaryTasks:=True, DisplayExternalSuccessors:=True, DisplayExternalPredecessors:=True
  Sort "ID", , , , , , False, True
  FilterClear
  SelectAll
  OutlineShowAllTasks
  OptionsViewEx DisplayNameIndent:=False, DisplaySummaryTasks:=False, DisplayExternalSuccessors:=True, DisplayExternalPredecessors:=True
  SetAutoFilter "Unique ID Predecessors", pjAutoFilterCustom, "contains", ":", "or", "contains", "<>"
  SelectAll
  
  'build taskindex
  Application.StatusBar = "Building TaskMap..."
  Set oTaskMap = CreateObject("Scripting.Dictionary")
  For Each oTask In ActiveProject.Tasks
    If Not oTask Is Nothing Then
      oTaskMap.Add oTask.UniqueID, oTask
    End If
  Next oTask
  Application.StatusBar = ""
  
  ReDim vCPL(0 To 17, 0 To 0)
  lngCount = 0
  lngPUID = FieldNameToFieldConstant(strProjectUID, pjTask)
  lngTask = 0
  lngTaskCount = ActiveSelection.Tasks.Count
  Application.StatusBar = "EXPORTING CPLs: Tasks " & Format(lngTask, "#,##0") & "/" & Format(lngTaskCount, "#,##0") & " (" & Format(lngTask / lngTaskCount, "0%") & ") | " & Format(lngCount, "#,##0") & " CPLs found"
  For Each oTask In ActiveSelection.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask = True Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    For Each oLink In oTask.TaskDependencies
      Set oFrom = Nothing
      Set oFrom = oLink.From
      Set oTo = Nothing
      Set oTo = oLink.To
      If oTo.Guid = oTask.Guid And oFrom.ExternalTask = True Then 'preds only
        If Not oFrom.Active Then GoTo next_link
        If lngCount > UBound(vCPL, 2) Then
          ReDim Preserve vCPL(0 To 17, 0 To UBound(vCPL, 2) + CHUNK_SIZE)
        End If
        'fix the returned UID
        lngSourceUID = oFrom.GetField(185073906) Mod 4194304
        'strProject = Replace(cptRxMatch(oFrom.Project, "[^\\/]+$"), ".mpp", "") 'KEEP: in case https://file.mpp
        strProject = Replace(Mid$(oFrom.Project, InStrRev(oFrom.Project, "\") + 1), ".mpp", "")
        lngFactor = oSubMap(strProject)
        lngMasterUID = (lngFactor * 4194304) + lngSourceUID
        vCPL(0, lngCount) = strProject
        vCPL(1, lngCount) = lngMasterUID
        vCPL(2, lngCount) = lngSourceUID
        Set oPred = Nothing
        On Error Resume Next
        Set oPred = oTaskMap(lngMasterUID)
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        strToPUID = oTask.GetField(lngPUID)
        If Not oPred Is Nothing Then
          strFromPUID = oPred.GetField(lngPUID)
          vCPL(3, lngCount) = strFromPUID
          vCPL(5, lngCount) = strFromPUID & "-" & strToPUID
        Else
          strFromPUID = "<<< GHOST >>>"
          vCPL(3, lngCount) = strFromPUID
          vCPL(5, lngCount) = strFromPUID
        End If
        vCPL(4, lngCount) = oFrom.Name
        vCPL(6, lngCount) = Choose(oLink.Type + 1, "FF", "FS", "SF", "SS")
        vCPL(7, lngCount) = Round(oLink.Lag / 480, 1)
        vCPL(8, lngCount) = oTask.Project
        vCPL(9, lngCount) = oTask.UniqueID
        vCPL(10, lngCount) = oTo.UniqueID
        vCPL(11, lngCount) = strToPUID
        vCPL(12, lngCount) = oTask.Name
        vCPL(13, lngCount) = oTask.ActualFinish
        vCPL(14, lngCount) = Choose(oTask.ConstraintType + 1, "ASAP", "ALAP", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT")
        vCPL(15, lngCount) = oTask.ConstraintDate
        vCPL(16, lngCount) = oTask.Start
        vCPL(17, lngCount) = oTask.PredecessorTasks.Count
        lngCount = lngCount + 1
      ElseIf oFrom.Guid = oTask.Guid And oTo.ExternalTask = True Then 'succs too tho
        'only export if ghost...
        If Not oTo.Active Then GoTo next_link
        If lngCount > UBound(vCPL, 2) Then
          ReDim Preserve vCPL(0 To 17, 0 To UBound(vCPL, 2) + CHUNK_SIZE)
        End If
        'fix the returned UID
        lngSourceUID = oTo.GetField(185073906) Mod 4194304
        'strProject = Replace(cptRxMatch(oFrom.Project, "[^\\/]+$"), ".mpp", "") 'KEEP: in case https://file.mpp
        strProject = Replace(Mid$(oTo.Project, InStrRev(oTo.Project, "\") + 1), ".mpp", "")
        lngFactor = oSubMap(strProject)
        lngMasterUID = (lngFactor * 4194304) + lngSourceUID
        Set oPred = Nothing 'todo: change to oSucc
        On Error Resume Next
        Set oPred = oTaskMap(lngMasterUID)
        If oPred Is Nothing Then
          vCPL(0, lngCount) = oTask.Project
          vCPL(1, lngCount) = oTask.UniqueID
          vCPL(2, lngCount) = oFrom.UniqueID
          strFromPUID = oTask.GetField(lngPUID)
          vCPL(3, lngCount) = strFromPUID
          vCPL(4, lngCount) = oTask.Name
          strToPUID = "<<< GHOST >>>"
          vCPL(5, lngCount) = strToPUID
          vCPL(6, lngCount) = Choose(oLink.Type + 1, "FF", "FS", "SF", "SS")
          vCPL(7, lngCount) = Round(oLink.Lag / 480, 1)
          vCPL(8, lngCount) = strProject
          vCPL(9, lngCount) = lngMasterUID
          vCPL(10, lngCount) = lngSourceUID
          vCPL(11, lngCount) = strToPUID
          vCPL(12, lngCount) = oTo.Name
          vCPL(13, lngCount) = oTo.ActualFinish 'todo: change this?
          vCPL(14, lngCount) = Choose(oTo.ConstraintType + 1, "ASAP", "ALAP", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT")
          vCPL(15, lngCount) = oTo.ConstraintDate
          vCPL(16, lngCount) = oTo.Start 'todo: change this?
          vCPL(17, lngCount) = oTo.PredecessorTasks.Count 'todo: change this?
          lngCount = lngCount + 1
        End If
      End If
next_link:
    Next oLink
next_task:
    lngTask = lngTask + 1
    If lngTask Mod 100 = 0 Then
      Application.StatusBar = "EXPORTING CPLs: Tasks " & Format(lngTask, "#,##0") & "/" & Format(lngTaskCount, "#,##0") & " (" & Format(lngTask / lngTaskCount, "0%") & ") | " & Format(lngCount, "#,##0") & " CPLs found"
    End If
    DoEvents
  Next oTask
  
  ReDim Preserve vCPL(0 To 17, 0 To lngCount - 1)
  
  'export to Excel
  Application.StatusBar = "Exporting to Excel..."
  DoEvents
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  oExcel.Visible = True
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.[A2].Resize(, UBound(vCPL, 1) + 1) = Split("PROJECT,UID[M],UID[S],PUID,TASK NAME,GRUID,TYPE,LAG(DAYS),PROJECT,UID[M],UID[S],PUID,TASK NAME,ACTUAL FINISH,CONSTRAINT TYPE,CONSTRAINT DATE,FORECAST START,PRED COUNT", ",")
  oWorksheet.[A3].Resize(UBound(vCPL, 2) + 1, UBound(vCPL, 1) + 1) = oExcel.WorksheetFunction.Transpose(vCPL)
  'conditional formatting
  With oWorksheet.Range(oWorksheet.[D3], oWorksheet.[D3].End(xlDown))
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""<<< GHOST >>>"""
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Font
      .Color = -16383844
      .TintAndShade = 0
    End With
    With .FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 13551615
      .TintAndShade = 0
    End With
    .FormatConditions(1).StopIfTrue = False
  End With
  With oWorksheet.Range(oWorksheet.[F3], oWorksheet.[F3].End(xlDown))
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""<<< GHOST >>>"""
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Font
      .Color = -16383844
      .TintAndShade = 0
    End With
    With .FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 13551615
      .TintAndShade = 0
    End With
    .FormatConditions(1).StopIfTrue = False
  End With
  With oWorksheet.Range(oWorksheet.[L3], oWorksheet.[L3].End(xlDown))
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""<<< GHOST >>>"""
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Font
      .Color = -16383844
      .TintAndShade = 0
    End With
    With .FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 13551615
      .TintAndShade = 0
    End With
    .FormatConditions(1).StopIfTrue = False
  End With
  
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.DisplayGridlines = False
  oWorksheet.[A1] = "GIVER"
  oWorksheet.[A1:E1].HorizontalAlignment = xlCenterAcrossSelection
  oWorksheet.[F1] = "LINK"
  oWorksheet.[F1:H1].HorizontalAlignment = xlCenterAcrossSelection
  oWorksheet.[i1] = "RECEIVER"
  oWorksheet.[I1:R1].HorizontalAlignment = xlCenterAcrossSelection
  oWorksheet.[A1].Resize(2, UBound(vCPL, 1) + 1).Font.Bold = True
  oWorksheet.[A1].Resize(1, UBound(vCPL, 1) + 1).Font.Size = 20
  oWorksheet.[A2].AutoFilter
  For Each vCol In Array(1, 5, 9, 13)
    With oWorksheet.Columns(vCol)
      If vCol = 1 Or vCol = 9 Then .ColumnWidth = 40
      If vCol = 5 Or vCol = 13 Then .ColumnWidth = 70
      .WrapText = True
    End With
  Next vCol
  For Each vCol In Array(14, 16, 17)
    oWorksheet.Range(oWorksheet.Cells(oWorksheet.[A3].Row, vCol), oWorksheet.Cells(oWorksheet.[A3].End(xlDown).Row, vCol)).NumberFormat = "m/d/yyyy"
  Next vCol
  oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight)).HorizontalAlignment = xlCenter
  For Each vCol In Array(1, 5, 9, 13)
    oWorksheet.Range(oWorksheet.Cells(oWorksheet.[A3].Row, vCol), oWorksheet.Cells(oWorksheet.[A3].End(xlDown).Row, vCol)).HorizontalAlignment = xlLeft
  Next vCol
  oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight)).VerticalAlignment = xlCenter
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.SplitRow = 2
  oExcel.ActiveWindow.FreezePanes = True
  oWorksheet.Columns.AutoFit
  cptAddBorders oWorksheet.Range(oWorksheet.[A2].End(xlToRight).Offset(-1, 0), oWorksheet.[A2].End(xlDown))
  cptAddBorders oWorksheet.Range(oWorksheet.[F1], oWorksheet.[F2].End(xlDown).Offset(0, 2))
  cptAddShading oWorksheet.Range(oWorksheet.[A2].Offset(-1, 0), oWorksheet.[A2].End(xlToRight))
  Application.StatusBar = Format(lngCount, "#,##0") & " cross-project links exported."
  DoEvents
  
  oWorkbook.VBProject.VBComponents("Sheet1").CodeModule.DeleteLines 1, 2
  strCode = "Private Const BLN_FILTER As Boolean = False" & vbCrLf
  strCode = strCode & "Option Explicit" & vbCrLf
  strCode = strCode & "" & vbCrLf
  strCode = strCode & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf
  strCode = strCode & "  If Not BLN_FILTER Then Exit Sub" & vbCrLf
  strCode = strCode & "  Dim strG_PUID As String" & vbCrLf
  strCode = strCode & "  Dim strR_PUID As String" & vbCrLf
  strCode = strCode & "  If Target.Cells.Count > 1 Then Exit Sub" & vbCrLf
  strCode = strCode & "  strG_PUID = Me.Cells(Target.Row, 4)" & vbCrLf
  strCode = strCode & "  strR_PUID = Me.Cells(Target.Row, 11)" & vbCrLf
  strCode = strCode & "  Dim oMSPROJ As Object 'MSProject.Application" & vbCrLf
  strCode = strCode & "  Dim oProject As Object 'MSProject.Project" & vbCrLf
  strCode = strCode & "  Set oMSPROJ = GetObject(, ""MSProject.Application"")" & vbCrLf
  strCode = strCode & "  Set oProject = oMSPROJ.ActiveProject" & vbCrLf
  strCode = strCode & "  If Len(strG_PUID) > 0 Or Len(strR_PUID) > 0 Then" & vbCrLf
  strCode = strCode & "    oMSPROJ.SetAutoFilter ""PUID"", 1, ""equals"", strG_PUID, ""or"", ""equals"", strR_PUID" & vbCrLf
  strCode = strCode & "  Else" & vbCrLf
  strCode = strCode & "    oMSPROJ.FilterClear" & vbCrLf
  strCode = strCode & "  End If" & vbCrLf
  strCode = strCode & "  oMSPROJ.SelectBeginning" & vbCrLf
  strCode = strCode & "  oMSPROJ.SelectAll" & vbCrLf
  strCode = strCode & "  Set oProject = Nothing" & vbCrLf
  strCode = strCode & "  Set oMSPROJ = Nothing" & vbCrLf
  strCode = strCode & "End Sub" & vbCrLf
  oWorkbook.VBProject.VBComponents("Sheet1").CodeModule.AddFromString strCode
  
  MsgBox Format(lngCount, "#,##0") & " cross-project links exported.", vbInformation + vbOKOnly, "CPLs"
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Application.StatusBar = ""
  oSubMap.RemoveAll
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oSubMap = Nothing
  Set oLink = Nothing
  Set oTask = Nothing
  Set oFrom = Nothing
  Set oTo = Nothing
  Set oPred = Nothing
  Set oTaskMap = Nothing
  Set oSubproject = Nothing
  Exit Sub
err_here:
  cptHandleErr THIS_MODULE, "cptExportCrossProjectLinks", Err, Erl
  Resume exit_here
End Sub

Sub cptGetSubMap()
  'objects
  Dim oSubproject As MSProject.SubProject
  Dim oTask As MSProject.Task
  'strings
  'longs
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oSubMap Is Nothing Then
    Set oSubMap = CreateObject("Scripting.Dictionary")
  Else
    oSubMap.RemoveAll
  End If
  Application.Calculation = pjManual
  For Each oSubproject In ActiveProject.Subprojects
    If Left(oSubproject.Path, 2) = "<>" Then 'PWA
      oSubMap.Add Replace(oSubproject.Path, "<>\", ""), 0
    Else 'mpp (local or remote)
      oSubMap.Add Replace(cptRegEx(oSubproject.Path, "[^\\/]*.mpp$"), ".mpp", ""), 0
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
  Application.CalculateProject
  For Each oTask In ActiveProject.Tasks
    If oSubMap.Exists(oTask.Project) Then
      If oSubMap(oTask.Project) > 0 Then GoTo next_mapping_task
      oSubMap.Item(oTask.Project) = CLng(oTask.UniqueID / 4194304)
    End If
next_mapping_task:
  Next oTask
  
exit_here:
  On Error Resume Next
  Application.Calculation = pjAutomatic
  Set oTask = Nothing
  Set oSubproject = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptGetSubMap", Err, Erl)
  Resume exit_here
  
End Sub
