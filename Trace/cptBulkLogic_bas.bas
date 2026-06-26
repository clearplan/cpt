Attribute VB_Name = "cptBulkLogic_bas"
'<cpt_version>v0.0.1</cpt_version>
Private Const THIS_MODULE As String = "cptBulkLogic_bas"
Option Explicit

Sub cptBulkLogicAddCommonPredecessor()
  cptShowBulkLogic_frm 0
End Sub

Sub cptBulkLogicAddCommonSuccessor()
  cptShowBulkLogic_frm 1
End Sub

Sub cptBulkLogicRemoveCommon()
  cptShowBulkLogic_frm 2
End Sub

Sub cptShowBulkLogic_frm(Optional lngPage As Long = 0)
  'objects
  Dim myBulkLogic_frm As New cptBulkLogic_frm
  'strings
  Dim strSetting As String
  Dim strColumnWidths As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vHeader As Variant
  Dim vLinkType As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If ActiveProject.Tasks.Count = 0 Then
    MsgBox "There are no tasks in this schedule.", vbExclamation + vbOKOnly, "cptBulkLogic"
    GoTo exit_here
  End If
  
  cptMakeBulkLogicDataset

  With myBulkLogic_frm
    .Caption = "cptBulkLogic (" & cptGetVersion(THIS_MODULE) & ")"
    'populate lboHeaders
    vHeader = Split("UID,ID,TASK NAME", ",")
    .lboHeaderFrom.IntegralHeight = False
    .lboFrom.IntegralHeight = True
    .lboHeaderFrom.List = cptTranspose(vHeader)
    .lboHeaderTo.IntegralHeight = False
    .lboTo.IntegralHeight = True
    .lboHeaderTo.List = cptTranspose(vHeader)
    strSetting = cptGetSetting("BulkLogic", "chkID")
    If Len(strSetting) > 0 Then
      .chkID = CBool(strSetting)
    Else
      .chkID = False
    End If
    If .chkID Then
      strColumnWidths = "0 pt;25 pt"
    Else
      strColumnWidths = "25 pt;0 pt"
    End If
    .lboHeaderFrom.ColumnWidths = strColumnWidths
    .lboHeaderTo.ColumnWidths = strColumnWidths
    .lboFrom.ColumnWidths = strColumnWidths
    .lboTo.ColumnWidths = strColumnWidths
    .lboHeaderFrom.Height = 12
    .lboHeaderTo.Height = 12
    'reposition lbo
    .lboHeaderFrom.Top = .txtFilterFrom.Top + .txtFilterFrom.Height + 1
    .lboHeaderTo.Top = .txtFilterTo.Top + .txtFilterTo.Height + 1
    .lboFrom.Top = .lboHeaderFrom.Top + .lboHeaderFrom.Height - 1
    .lboTo.Top = .lboHeaderTo.Top + .lboHeaderTo.Height - 1
    'populate cboLinkTypes
    .cboLinkType.Clear
    For Each vLinkType In Split("0:Finish to Finish,1:Finish to Start,2:Start to Finish,3:Start to Start", ",")
      .cboLinkType.AddItem
      .cboLinkType.List(.cboLinkType.ListCount - 1, 0) = Split(vLinkType, ":")(0)
      .cboLinkType.List(.cboLinkType.ListCount - 1, 1) = Split(vLinkType, ":")(1)
    Next vLinkType
    .cboLinkType.Value = 1  'FS default
    'set the page
    .MultiPage1.Value = lngPage
    cptUpdateBulkLogicForm myBulkLogic_frm
    .Show 'eventually make this false and update on selection
  End With

exit_here:
  On Error Resume Next
  Set myBulkLogic_frm = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptShowBulkLogic_frm", Err, Erl)
  Resume exit_here

End Sub

Sub cptUpdateBulkLogicForm(ByRef myBulkLogic_frm As cptBulkLogic_frm)
  'objects
  Dim oPreds As Scripting.Dictionary
  Dim oSuccs As Scripting.Dictionary
  Dim oLink As MSProject.TaskDependency
  Dim oTask As MSProject.Task
  Dim oTasks As MSProject.Tasks
  Dim oListBox As MSForms.ListBox
  Dim oTextBox As MSForms.TextBox
  'strings
  Dim strColumnWidths As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  Debug.Print "triggered!"
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  With myBulkLogic_frm
    .Caption = "cptBulkLogic (" & cptGetVersion(THIS_MODULE) & ")"
    .lboFrom.Clear
    .lboTo.Clear
    If .chkID Then
      strColumnWidths = "0 pt;25 pt"
    Else
      strColumnWidths = "25 pt;0 pt"
    End If
    .lboHeaderFrom.ColumnWidths = strColumnWidths
    .lboHeaderTo.ColumnWidths = strColumnWidths
    .lboFrom.ColumnWidths = strColumnWidths
    .lboTo.ColumnWidths = strColumnWidths
    'populate listbox
    If .MultiPage1.Value < 2 Then
      .cmdApply.Caption = "Add"
      .cmdApply.Enabled = False 'until at least one item is selected
      .txtFilterFrom.Enabled = True
      .txtFilterTo.Enabled = True
      .cboLinkType.Enabled = True
      .lboFrom.Enabled = True
      .lboTo.Enabled = True
      .txtLag.Enabled = True
      If .MultiPage1.Value = 0 Then  'common pred
        Set oListBox = .lboTo
        Set oTextBox = .txtFilterFrom
        .txtFilterTo.Enabled = False
        .lboTo.Enabled = False
      ElseIf .MultiPage1.Value = 1 Then  'common succ
        Set oListBox = .lboFrom
        Set oTextBox = .txtFilterTo
        .txtFilterFrom.Enabled = False
        .lboFrom.Enabled = False
      End If
      DoEvents
      Set oTasks = ActiveSelection.Tasks
      If oTasks Is Nothing Then
        'notify in listbox
        GoTo skip_load
      ElseIf oTasks.Count = 0 Then
        'notify in listbox
        GoTo skip_load
      End If
      oListBox.Clear
      For Each oTask In oTasks
        If oTask Is Nothing Then GoTo next_task
        If oTask.Summary Then GoTo next_task
        If Not oTask.Active Then GoTo next_task
        If oTask.ExternalTask Then GoTo next_task 'todo: yes no?
        oListBox.AddItem
        oListBox.List(oListBox.ListCount - 1, 0) = oTask.UniqueID
        oListBox.List(oListBox.ListCount - 1, 1) = oTask.ID
        oListBox.List(oListBox.ListCount - 1, 2) = Replace(oTask.Name, ",", "-")
next_task:
      Next oTask
      .txtLag.Value = 0
skip_load:
      .Controls(oTextBox.Name).SetFocus
    Else
      .cmdApply.Caption = "Remove"
      .cmdApply.Enabled = False 'until at least one item is selected
      .txtFilterFrom.Enabled = False
      .lboFrom.Enabled = True
      .txtFilterTo.Enabled = False
      .lboTo.Enabled = True
      .cboLinkType.Enabled = False
      .txtLag.Enabled = False
      Set oPreds = CreateObject("Scripting.Dictionary")
      Set oSuccs = CreateObject("Scripting.Dictionary")
      Set oTasks = ActiveSelection.Tasks
      For Each oTask In oTasks
        If oTask Is Nothing Then GoTo next_task2
        If oTask.Summary Then GoTo next_task2
        If Not oTask.Active Then GoTo next_task2
        If oTask.ExternalTask Then GoTo next_task2
        For Each oLink In oTask.TaskDependencies
          If oLink.To = oTask Then
            'capture preds
            If oPreds.Exists(oLink.From.UniqueID) Then
              oPreds(oLink.From.UniqueID) = oPreds(oLink.From.UniqueID) + 1
            Else
              oPreds.Add oLink.From.UniqueID, 1
            End If
          ElseIf oLink.From = oTask Then
            'capture succs
            If oSuccs.Exists(oLink.To.UniqueID) Then
              oSuccs(oLink.To.UniqueID) = oSuccs(oLink.To.UniqueID) + 1
            Else
              oSuccs.Add oLink.To.UniqueID, 1
            End If
          End If
        Next oLink
next_task2:
      Next oTask
      For lngItem = 0 To oPreds.Count - 1
        If oPreds.Items(lngItem) > 1 Then
          .lboFrom.AddItem
          .lboFrom.List(.lboFrom.ListCount - 1, 0) = oPreds.Keys(lngItem)
          .lboFrom.List(.lboFrom.ListCount - 1, 1) = ActiveProject.Tasks.UniqueID(oPreds.Keys(lngItem)).ID
          .lboFrom.List(.lboFrom.ListCount - 1, 2) = Replace(ActiveProject.Tasks.UniqueID(oPreds.Keys(lngItem)).Name, ",", "-")
        End If
      Next lngItem
      For lngItem = 0 To oSuccs.Count - 1
        If oSuccs.Items(lngItem) > 1 Then
          .lboTo.AddItem
          .lboTo.List(.lboTo.ListCount - 1, 0) = oSuccs.Keys(lngItem)
          .lboTo.List(.lboTo.ListCount - 1, 1) = ActiveProject.Tasks.UniqueID(oSuccs.Keys(lngItem)).ID
          .lboTo.List(.lboTo.ListCount - 1, 2) = Replace(ActiveProject.Tasks.UniqueID(oSuccs.Keys(lngItem)).Name, ",", "-")
        End If
      Next lngItem
    End If
  End With
  
exit_here:
  On Error Resume Next
  Set oPreds = Nothing
  Set oSuccs = Nothing
  Set oLink = Nothing
  Set oTask = Nothing
  Set oTasks = Nothing
  Set oTextBox = Nothing
  Set oListBox = Nothing
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptUpdateBulkLogicForm", Err, Erl)
  Resume exit_here
End Sub

Sub cptMakeBulkLogicDataset()
  'objects
  Dim oLink As MSProject.TaskDependency
  Dim oDict As Scripting.Dictionary
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  'strings
  Dim strFileName As String
  'longs
  Dim lngBLT As Long 'Bulk Logic Tasks
  Dim lngBLL As Long 'Bulk Logic Links
  Dim lngFile As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'create Schema.ini
  strFileName = Environ("tmp") & "\Schema.ini"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  lngFile = FreeFile
  Open strFileName For Output As #lngFile
  Print #lngFile, "[cpt-blt.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=UID Integer"
  Print #lngFile, "Col2=ID Integer"
  Print #lngFile, "Col3=TASK_NAME Text Width 100"
  Print #lngFile, "[cpt-bll.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=FROM_UID Integer"
  Print #lngFile, "Col2=TO_UID Integer"
  Close #lngFile
  'create bulk-logic-tasks.csv
  strFileName = Environ("tmp") & "\cpt-blt.csv"
  lngBLT = FreeFile
  Open strFileName For Output As #lngBLT
  Print #lngFile, "UID,ID,TASK_NAME"
  'create bulk-logic-links.csv
  strFileName = Environ("tmp") & "\cpt-bll.csv"
  lngBLL = FreeFile
  Open strFileName For Output As #lngBLL
  Print #lngBLL, "FROM_UID,TO_UID"
  'dump it
  Set oTasks = ActiveProject.Tasks
  lngTasks = oTasks.Count
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task 'todo: yes no?
    If Not oTask.Active Then GoTo next_task
    Print #lngBLT, oTask.UniqueID & "," & oTask.ID & "," & Replace(oTask.Name, ",", "-")
    For Each oLink In oTask.TaskDependencies
      If oLink.To = oTask Then 'only do preds
        Print #lngBLL, oLink.From.UniqueID & "," & oLink.To.UniqueID
      End If
    Next oLink
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Building task dataset...(" & Format(lngTask / lngTasks, "0%") & ")"
  Next oTask
  Close #lngBLL
  Close #lngBLT
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oLink = Nothing
  Set oTask = Nothing
  Set oTasks = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptMakeBulkLogicDataset", Err, Erl)
  Resume exit_here

End Sub

Sub cptBulkLogicUpdateCommand(ByRef myBulkLogic_frm As cptBulkLogic_frm)
  Dim lngPreds As Long
  Dim lngSuccs As Long
  Dim lngItem As Long
  With myBulkLogic_frm
    For lngItem = 0 To .lboFrom.ListCount - 1
      If .lboFrom.Selected(lngItem) Then lngPreds = lngPreds + 1
    Next lngItem
    For lngItem = 0 To .lboTo.ListCount - 1
      If .lboTo.Selected(lngItem) Then lngSuccs = lngSuccs + 1
    Next lngItem
    If .MultiPage1.Value = 0 Then
      .cmdApply.Enabled = lngPreds > 0
    ElseIf .MultiPage1.Value = 1 Then
      .cmdApply.Enabled = lngSuccs > 0
    ElseIf .MultiPage1.Value = 2 Then
      .cmdApply.Enabled = (lngPreds + lngSuccs) > 0
    End If
  End With
End Sub

Sub cptBulkLogicApply(ByRef myBulkLogic_frm As cptBulkLogic_frm)
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  Dim oLink As MSProject.TaskDependency
  'strings
  'longs
  Dim lngFrom As Long
  Dim lngFromUID As Long
  Dim lngTo As Long
  Dim lngToUID As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSpeed True
  
  With myBulkLogic_frm
    If .MultiPage1.Value = 0 Then 'preds
      For lngFrom = 0 To .lboFrom.ListCount - 1
        If .lboFrom.Selected(lngFrom) Then 'only link selected pred
          lngFromUID = .lboFrom.List(lngFrom, 0)
          For lngTo = 0 To .lboTo.ListCount - 1
            lngToUID = .lboTo.List(lngTo, 0)
            ActiveProject.Tasks.UniqueID(lngToUID).TaskDependencies.Add ActiveProject.Tasks.UniqueID(lngFromUID), .cboLinkType.Value, .txtLag.Value
          Next lngTo
        End If
      Next lngFrom
    ElseIf .MultiPage1.Value = 1 Then 'succs
      For lngFrom = 0 To .lboFrom.ListCount - 1
        lngFromUID = .lboFrom.List(lngFrom, 0)
        For lngTo = 0 To .lboTo.ListCount - 1
          If .lboTo.Selected(lngFrom) Then 'only link selected succ
            lngToUID = .lboTo.List(lngTo, 0)
            ActiveProject.Tasks.UniqueID(lngToUID).TaskDependencies.Add ActiveProject.Tasks.UniqueID(lngFromUID), .cboLinkType.Value, .txtLag.Value
          End If
        Next lngTo
      Next lngFrom
    ElseIf .MultiPage1.Value = 2 Then 'remove
      Set oTasks = ActiveSelection.Tasks
      For lngFrom = .lboFrom.ListCount - 1 To 0 Step -1
        If .lboFrom.Selected(lngFrom) Then
          lngFromUID = .lboFrom.List(lngFrom, 0)
          For Each oTask In oTasks
            For Each oLink In oTask.TaskDependencies
              If oLink.To = oTask And oLink.From.UniqueID = lngFromUID Then
                oLink.Delete
              End If
            Next oLink
          Next oTask
          .lboFrom.RemoveItem lngFrom
        End If
      Next lngFrom
      For lngTo = .lboTo.ListCount - 1 To 0 Step -1
        If .lboTo.Selected(lngTo) Then
          lngToUID = .lboTo.List(lngTo, 0)
          For Each oTask In oTasks
            For Each oLink In oTask.TaskDependencies
              If oLink.From = oTask And oLink.To.UniqueID = lngToUID Then
                oLink.Delete
              End If
            Next oLink
          Next oTask
        End If
        .lboTo.RemoveItem (lngTo)
      Next lngTo
    End If
  End With

exit_here:
  On Error Resume Next
  cptSpeed False
  
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptBulkLogicApply", Err, Erl)
  Resume exit_here

End Sub
