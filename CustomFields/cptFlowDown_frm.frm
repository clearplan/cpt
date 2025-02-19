VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFlowDown_frm 
   Caption         =   "LCF FlowDown"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
   OleObjectBlob   =   "cptFlowDown_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptFlowDown_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.0</cpt_version>
Option Explicit
Private oLCF As ADODB.Recordset

Private Sub cptUpdateLCFList()
  Dim strDataTypes As String
  Dim oControl As Control
  Dim lngItem As Long
  For lngItem = 0 To Me.lboSubprojects.ListCount - 1
    Me.lboSubprojects.List(lngItem, 2) = "..."
  Next lngItem
  For Each oControl In Me.Controls
    If Left(oControl.Name, 3) = "chk" Then
      If oControl.Name = "chkFilter" Then GoTo next_control
      If Me.Controls(oControl.Name).Value = True Then
        If oControl.Name = "chkOutlineCode" Then
          strDataTypes = strDataTypes & "Outline Code,"
        Else
          strDataTypes = strDataTypes & Replace(oControl.Name, "chk", "") & ","
        End If
      End If
    End If
next_control:
  Next oControl
  If Len(strDataTypes) > 0 Then
    strDataTypes = Left(strDataTypes, Len(strDataTypes) - 1)
    Me.lboMasterLCF.List = cptGetCustomFields("t", strDataTypes, "c,fn,cfn")
  Else
    Me.lboMasterLCF.Clear
  End If
End Sub

Private Sub cboMoveTo_Change()
  If Me.ActiveControl.Name = "cboMoveTo" Then cptUpdateLCFMoveToPreview Me
End Sub

Private Sub chkCost_Click()
  cptUpdateLCFList
End Sub
Private Sub chkDate_Click()
  cptUpdateLCFList
End Sub
Private Sub chkDuration_Click()
  cptUpdateLCFList
End Sub
Private Sub chkFlag_Click()
  cptUpdateLCFList
End Sub
Private Sub chkNumber_Click()
  cptUpdateLCFList
End Sub
Private Sub chkOutlineCode_Click()
  cptUpdateLCFList
End Sub
Private Sub chkText_Click()
  cptUpdateLCFList
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdLoad_Click()
  Dim oProject As MSProject.Project
  Dim oSubproject As MSProject.Subproject
  Dim vLCF As Variant
  Dim lngRow As Long
  Set oProject = Projects(Me.cboMasterProject.Value)
  If Not oLCF Is Nothing Then oLCF.Close
  Set oLCF = CreateObject("ADODB.Recordset")
  oLCF.Fields.Append "SP", adInteger
  oLCF.Fields.Append "FC", adInteger
  oLCF.Fields.Append "FN", adVarChar, 50
  oLCF.Fields.Append "CFN", adVarChar, 100
  oLCF.Open
  For Each oSubproject In oProject.Subprojects
    If Not oSubproject.IsLoaded Then FileOpenEx oSubproject.Path, False
    Projects(oSubproject.Path).Activate
    vLCF = cptGetCustomFields("t", "Text", "c,fn,cfn")
    For lngRow = 0 To UBound(vLCF, 1)
      oLCF.AddNew Array(0, 1, 2, 3), Array(oSubproject.Index, CLng(vLCF(lngRow, 0)), vLCF(lngRow, 1), vLCF(lngRow, 2))
      'todo: progress %
    Next lngRow
    oLCF.Update
  Next oSubproject
  Me.lboSubprojects.Enabled = True
  oProject.Activate
  WindowArrangeAll
  Set oSubproject = Nothing
  Set oProject = Nothing
End Sub

Private Sub cmdMoveLCF_Click()
  cptMoveLCF Me, oLCF
End Sub

Private Sub cmdTake_Click()
  'objects
  Dim oMP As MSProject.Project
  Dim oSP As MSProject.Project
  Dim oTask As MSProject.Task
  'strings
  Dim strFormula As String
  Dim strLCF As String
  'longs
  Dim lngLCF As Long
  Dim lngItem As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  MsgBox "isolate this to only the selected subproject", vbExclamation + vbOKOnly, "todo" 'todo
  GoTo exit_here
  lngLCF = Me.lboMasterLCF.Value
  If lngLCF = 0 Then Exit Sub
  Set oMP = Projects(Me.cboMasterProject.Value)
  oMP.Activate
  strLCF = CustomFieldGetName(lngLCF)
  Application.Calculation = pjManual
  Application.ScreenUpdating = False
  For lngItem = 0 To Me.lboSubprojects.ListCount - 1
    Me.lboSubprojects.List(lngItem, 2) = "..."
    DoEvents
    'activate the subproject
    Set oSP = Projects(Me.lboSubprojects.List(lngItem, 1))
    lngTask = 0
    lngTasks = oSP.Tasks.Count
    oSP.Activate
    'delete the LCF
    CustomFieldDelete lngLCF
    'rename the LCF
    CustomFieldRename lngLCF, strLCF
    'todo: formula
    oMP.Activate
    strFormula = Application.CustomFieldGetFormula(lngLCF)
    If Len(strFormula) > 0 Then
      oSP.Activate
      Application.CustomFieldSetFormula lngLCF, strFormula
    End If
    'todo: pick list
    'todo: how to 'get' what funciton is currently being used...?
    'todo: outline code
    SetAutoFilter strLCF, pjAutoFilterIn, "does not equal", ""
    If MsgBox("Delete task data?", vbQuestion + vbYesNo, Me.lboSubprojects.List(lngItem, 1)) = vbYes Then
      'now delete the data
      For Each oTask In oSP.Tasks
        If oTask Is Nothing Then GoTo next_task
        If oTask.ExternalTask Then GoTo next_task
        If oTask.GetField(lngLCF) <> "" Then oTask.SetField lngLCF, ""
next_task:
        lngTask = lngTask + 1
        Me.lboSubprojects.List(lngItem, 2) = "(" & Format(lngTask / lngTasks, "0%") & ")"
        DoEvents
        'todo: maybe export, map, import?
      Next oTask
    End If
    'update listbox
    Me.lboSubprojects.List(lngItem, 2) = strLCF
    DoEvents
  Next lngItem

exit_here:
  On Error Resume Next
  Application.ScreenUpdating = True
  Application.Calculation = pjAutomatic
  Set oTask = Nothing
  Set oSP = Nothing
  Set oMP = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_frm", "cmdTake_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdUnload_Click()
  Dim oProject As MSProject.Project
  Dim oSubproject As MSProject.Subproject
  Set oProject = Projects(Me.cboMasterProject.Value)
  For Each oSubproject In oProject.Subprojects
    If oSubproject.IsLoaded Then
      Projects(Replace(Dir(oSubproject.Path), ".mpp", "")).Activate
      FileCloseEx pjDoNotSave
    End If
  Next oSubproject
  Me.lboSubprojects.Enabled = False
  oProject.Activate
  Set oSubproject = Nothing
  Set oSubproject = Nothing
  
End Sub

Private Sub lboMasterLCF_AfterUpdate()
  'objects
  Dim oProject As MSProject.Project
  Dim oSubproject As MSProject.Subproject
  'strings
  Dim strPicklist As String
  Dim strFN As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnFilter As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'ScreenUpdating = False
  'Calculation = pjManual

  'clear subproject lookup
  For lngItem = 0 To Me.lboSubprojects.ListCount - 1
    Me.lboSubprojects.List(lngItem, 2) = "..."
    DoEvents
  Next lngItem
  Me.lboSubprojects.ListIndex = -1
  Me.cboMoveTo.Clear
  Me.lboMoveLCF.Clear
  blnFilter = Me.chkFilter
  'update Table to show only selected LCF
  If blnFilter Then
    Projects(Me.cboMasterProject.Value).Activate
    Application.OptionsViewEx DisplayNameIndent:=False, DisplaySummaryTasks:=False
    ViewApply "Task Entry", True
    TableEditEx "LCF FlowDown Table", True, True, True, , "Unique ID"
    TableEditEx "LCF FlowDown Table", True, False, False, , , Me.lboMasterLCF.List(Me.lboMasterLCF.ListIndex, 1)
    TableApply "LCF FlowDown Table"
    ColumnBestFit Column:=1
    ColumnBestFit Column:=2
    SetSplitBar 3
    DoEvents
    SetSplitBar 2
    DoEvents
    FilterClear
    strFN = FieldConstantToFieldName(Me.lboMasterLCF.Value)
    If InStr(strFN, "Cost") > 0 Then
      SetAutoFilter strFN, Criteria1:=">0"
    ElseIf InStr(strFN, "Date") > 0 Then
      SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:="NA"
    ElseIf InStr(strFN, "Duration") > 0 Then
      SetAutoFilter strFN, Criteria1:=">0"
    ElseIf InStr(strFN, "Flag") > 0 Then
      SetAutoFilter strFN, FilterType:=pjAutoFilterFlagYes
    ElseIf InStr(strFN, "Number") > 0 Then
      SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:=0
    Else
      SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:=""
    End If
    SelectBeginning
    DoEvents
  End If
  Me.lblFormula.Visible = Len(CustomFieldGetFormula(Me.lboMasterLCF.Value)) > 0
  strPicklist = ""
  Dim oOutlineCode As MSProject.OutlineCode
  On Error Resume Next
  
  If InStr(FieldConstantToFieldName(Me.lboMasterLCF.Value), "Outline Code") > 0 Then
    Set oOutlineCode = Nothing
    Set oOutlineCode = ActiveProject.OutlineCodes(CustomFieldGetName(Me.lboMasterLCF.Value))
    If Not oOutlineCode Is Nothing Then
      If oOutlineCode.LookupTable.Count > 0 Then
        strPicklist = "has items"
      End If
    Else
      strPicklist = ""
    End If
  Else
    On Error Resume Next
    strPicklist = CustomFieldValueListGetItem(Me.lboMasterLCF.Value, pjValueListValue, 1)
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Len(strPicklist) > 1 Then
    Me.lblPickList.Visible = True
  Else
    Me.lblPickList.Visible = False
  End If

  'populate subproject LCF CustomFieldName
  Set oProject = Projects(Me.cboMasterProject.Value)
  For lngItem = 0 To Me.lboSubprojects.ListCount - 1
    oLCF.Filter = "SP=" & lngItem + 1 & " AND FC=" & Me.lboMasterLCF.Value
    If Not oLCF.EOF Then
      Me.lboSubprojects.List(lngItem, 2) = oLCF(3)
      DoEvents
    End If
    oLCF.Filter = 0
    If blnFilter Then
      Projects(oProject.Subprojects(lngItem + 1).Path).Activate
      Application.OptionsViewEx DisplayNameIndent:=False, DisplaySummaryTasks:=False
      ViewApply "Gantt Chart", True
      TableEditEx "LCF FlowDown Table", True, True, True, , "Unique ID"
      TableEditEx "LCF FlowDown Table", True, False, False, , , Me.lboMasterLCF.List(Me.lboMasterLCF.ListIndex, 1)
      TableApply "LCF FlowDown Table"
      ColumnBestFit Column:=1
      ColumnBestFit Column:=2
      SetSplitBar 3
      DoEvents
'      SetSplitBar 2
'      DoEvents
      FilterClear
      If InStr(strFN, "Cost") > 0 Then
        SetAutoFilter strFN, Criteria1:=">0"
      ElseIf InStr(strFN, "Date") > 0 Then
        SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:="NA"
      ElseIf InStr(strFN, "Duration") > 0 Then
        SetAutoFilter strFN, Criteria1:=">0"
      ElseIf InStr(strFN, "Flag") > 0 Then
        SetAutoFilter strFN, FilterType:=pjAutoFilterFlagYes
      ElseIf InStr(strFN, "Number") > 0 Then
        SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:=0
      Else
        SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:=""
      End If
      SelectBeginning
      DoEvents
    End If
  Next lngItem
  Projects(Me.cboMasterProject.Value).Activate

exit_here:
  On Error Resume Next
  ScreenUpdating = True
  Calculation = pjAutomatic
  Set oProject = Nothing
  Set oSubproject = Nothing
  Set oOutlineCode = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_frm", "lboMasterLCF_AfterUpdate", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboSubprojects_AfterUpdate()
  
  If Not IsNull(Me.lboSubprojects.ListIndex) Then
    Projects(Me.lboSubprojects.List(Me.lboSubprojects.ListIndex, 1)).Activate
    Me.tglMove.Enabled = True
    cptUpdateMoveToLCF Me, oLCF
    Me.lboMoveLCF.Clear
  Else
    Me.tglMove.Enabled = False
  End If
  
End Sub

Private Sub stxtSearch_Change()
  Dim oProject As MSProject.Project
  Dim oSubproject As MSProject.Subproject
  Dim strText As String
  strText = Me.stxtSearch.Text
  Me.lboSubprojects.Clear
  Set oProject = Projects(Me.cboMasterProject.Value)
  If Len(Me.stxtSearch.Text) > 0 Then
    Me.lboMasterLCF.ListIndex = -1
    Me.lboMasterLCF.Enabled = False
    Me.tglMove = False
    Me.chkCost.Enabled = False
    Me.chkDate.Enabled = False
    Me.chkDuration.Enabled = False
    Me.chkFlag.Enabled = False
    Me.chkNumber.Enabled = False
    Me.chkOutlineCode.Enabled = False
    Me.chkText.Enabled = False
    oLCF.Filter = "CFN Like '%" & strText & "%'"
    If Not oLCF.EOF Then
      oLCF.MoveFirst
      Do While Not oLCF.EOF
        If oLCF(0) > 0 And oLCF(3) <> oLCF(2) Then
          Me.lboSubprojects.AddItem
          Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 0) = oLCF(0)
          Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 1) = oProject.Subprojects(oLCF(0)).InsertedProjectSummary.Name
          Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 2) = oLCF(2) & " (" & oLCF(3) & ")"
        End If
        oLCF.MoveNext
      Loop
    End If
  Else
    Me.lboMasterLCF.Enabled = True
    Me.chkCost.Enabled = True
    Me.chkDate.Enabled = True
    Me.chkDuration.Enabled = True
    Me.chkFlag.Enabled = True
    Me.chkNumber.Enabled = True
    Me.chkOutlineCode.Enabled = True
    Me.chkText.Enabled = True
    For Each oSubproject In oProject.Subprojects
      Me.lboSubprojects.AddItem
      Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 0) = oSubproject.Index
      Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 1) = oSubproject.InsertedProjectSummary.Name
      Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 2) = "..."
    Next oSubproject
  End If
  oLCF.Filter = 0
  Set oSubproject = Nothing
  Set oProject = Nothing
End Sub

Private Sub tglMove_Click()
  Me.lblMoveTo.Visible = Me.tglMove
  Me.lboMoveLCF.Visible = Me.tglMove
  Me.cboMoveTo.Visible = Me.tglMove
  Me.cmdMoveLCF.Visible = Me.tglMove
  If Me.tglMove Then
    Me.lboSubprojects.Height = 108.75
    Me.lboMoveLCF.Height = 95
    If Not IsNull(Me.lboSubprojects.Value) Then cptUpdateMoveToLCF Me, oLCF
  Else
    Me.lboSubprojects.Height = Me.lboMasterLCF.Height
  End If
End Sub

Private Sub UserForm_Initialize()
  'objects
  Dim oProject As MSProject.Project
  Dim oSubproject As MSProject.Subproject
  'strings
  Dim strCFN As String
  'longs
  Dim lngRow As Long
  Dim lngLCF As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vLCF As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  For Each oProject In Projects
    If oProject.Subprojects.Count > 0 Then
      Me.cboMasterProject.AddItem oProject.Name
      Me.cboMasterProject.Value = oProject.Name
      'adjust the view
      oProject.Activate
      ActiveWindow.TopPane.Activate
      FilterClear
      GroupClear
      Sort "ID", , , , , , False, True
      OptionsViewEx DisplaySummaryTasks:=True
      OutlineShowAllTasks
      'add master
      Me.cboMasterProject.AddItem oProject.Name
      Me.cboMasterProject.Value = oProject.Name
      Me.cboMasterProject.Locked = True
      'create database of LCFs
      If Not oLCF Is Nothing Then oLCF.Close
      Set oLCF = CreateObject("ADODB.Recordset")
      oLCF.Fields.Append "SP", adInteger
      oLCF.Fields.Append "FC", adInteger
      oLCF.Fields.Append "FN", adVarChar, 50
      oLCF.Fields.Append "CFN", adVarChar, 100
      oLCF.Open
      'get master LCFs
      vLCF = cptGetCustomFields("t", "Cost,Date,Duration,Flag,Number,Outline Code,Text", "c,fn,cfn")
      For lngRow = 0 To UBound(vLCF, 1)
        oLCF.AddNew Array(0, 1, 2, 3), Array(0, CLng(vLCF(lngRow, 0)), vLCF(lngRow, 1), vLCF(lngRow, 2))
      Next lngRow
      'get subprojects
      For Each oSubproject In oProject.Subprojects
        If oSubproject.IsLoaded Then
          'adjust the view
          On Error Resume Next
          Projects(oSubproject.SourceProject.Name).Activate
          ActiveWindow.TopPane.Activate
          FilterClear
          GroupClear
          Sort "ID", , , , , , False, True
          OptionsViewEx DisplaySummaryTasks:=True
          OutlineShowAllTasks
          'add to lboSubprojects
          Me.lboSubprojects.AddItem
          Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 0) = oSubproject.Index
          Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 1) = oSubproject.InsertedProjectSummary.Name
          If Not oSubproject.IsLoaded Then FileOpenEx oSubproject.Path, False
          vLCF = cptGetCustomFields("t", "Cost,Date,Duration,Flag,Number,Outline Code,Text", "c,fn,cfn")
          For lngRow = 0 To UBound(vLCF, 1)
            oLCF.AddNew Array(0, 1, 2, 3), Array(oSubproject.Index, CLng(vLCF(lngRow, 0)), vLCF(lngRow, 1), vLCF(lngRow, 2))
          Next lngRow
          oLCF.Update
        End If
      Next oSubproject
    End If
    oProject.Activate
    Exit For
  Next oProject
  WindowArrangeAll
  
  'todo: only show the form if a master project is loaded
  'todo: save selections?
  Me.chkCost.Value = False
  Me.chkDate.Value = False
  Me.chkDuration.Value = False
  Me.chkFlag.Value = False
  Me.chkNumber.Value = False
  Me.chkOutlineCode.Value = False
  Me.chkText.Value = True
  Me.lblFormula.Visible = False
  Me.lblFormula.Top = Me.lboMasterLCF.Top + Me.lboMasterLCF.Height + 1
  Me.lblPickList.Visible = False
  Me.lblPickList.Top = Me.lboMasterLCF.Top + Me.lboMasterLCF.Height + 1

exit_here:
  On Error Resume Next
  Set oSubproject = Nothing
  Set oProject = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_frm", "UserForm_Initialize", Err, Erl)
  Resume exit_here
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Set oLCF = Nothing
  If Not IsNull(Me.cboMasterProject) Then
    Projects(Me.cboMasterProject.Value).Activate
  End If
  Unload Me
  'Cancel = True
'  If CloseMode = VbQueryClose.vbFormControlMenu Then
'    Unload Me
'    Cancel = True
'  ElseIf CloseMode = VbQueryClose.vbFormCode Then
'
'  End If
End Sub
