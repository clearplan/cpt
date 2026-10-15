VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFlowDown_frm 
   Caption         =   "LCF FlowDown"
   ClientHeight    =   6540
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

'<cpt_version>v0.0.1</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptFlowDown_frm"
Private oLCF As ADODB.Recordset

Private Sub cboMoveTo_Change()
  If Me.ActiveControl.Name = "cboMoveTo" Then cptUpdateLCFMoveToPreview Me
End Sub

Private Sub chkCost_Click()
  cptUpdateLCFList Me
End Sub
Private Sub chkDate_Click()
  cptUpdateLCFList Me
End Sub
Private Sub chkDuration_Click()
  cptUpdateLCFList Me
End Sub

Private Sub chkFinish_Click()
  cptUpdateLCFList Me
End Sub

Private Sub chkFlag_Click()
  cptUpdateLCFList Me
End Sub
Private Sub chkNumber_Click()
  cptUpdateLCFList Me
End Sub
Private Sub chkOutlineCode_Click()
  cptUpdateLCFList Me
End Sub

Private Sub chkStart_Click()
  cptUpdateLCFList Me
End Sub

Private Sub chkText_Click()
  cptUpdateLCFList Me
End Sub

Private Sub cmdAssignmentRolldownFalse_Click()
  cptAssignmentRolldown Me, False
  'todo: also clear task-level data? if so use SetTaskField/SetField
End Sub

Private Sub cmdAssignmentRolldownTrue_Click()
  cptAssignmentRolldown Me, True
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdDone_Click()
  Me.Hide
End Sub

Private Sub cmdLoad_Click()
  Dim oProject As MSProject.Project
  Dim oSubProject As MSProject.SubProject
  Dim vLCF As Variant
  Dim lngRow As Long
  Set oProject = Projects(Me.cboMasterProject.Value)
  If Not oLCF Is Nothing Then oLCF.Close
  Set oLCF = CreateObject("ADODB.Recordset")
  oLCF.Fields.Append "SP", adInteger 'subproject index
  oLCF.Fields.Append "FC", adInteger 'field constant
  oLCF.Fields.Append "FN", adVarChar, 50 'field name (native)
  oLCF.Fields.Append "CFN", adVarChar, 100 'custom field name
  oLCF.Open
  For Each oSubProject In oProject.Subprojects
    Debug.Print oSubProject.SourceProject.Tasks.Count 'forces load
    FileOpenEx oSubProject.Path, False
    Projects(oSubProject.Path).Activate
    vLCF = cptGetCustomFields("t", "Text", "c,fn,cfn")
    For lngRow = 0 To UBound(vLCF, 1)
      oLCF.AddNew Array(0, 1, 2, 3), Array(oSubProject.Index, CLng(vLCF(lngRow, 0)), vLCF(lngRow, 1), vLCF(lngRow, 2))
      'todo: progress %
    Next lngRow
    oLCF.Update
  Next oSubProject
  Me.lboSubprojects.Enabled = True
  oProject.Activate
  WindowArrangeAll
  Set oSubProject = Nothing
  Set oProject = Nothing
  'keep oLCF open for life of the userform
End Sub

Private Sub cmdMoveLCF_Click()
  cptMoveLCF Me, oLCF
End Sub

Private Sub cmdTake_Click()
  cptTakeLCF Me, oLCF
End Sub

Private Sub cmdUnload_Click()
  Dim oProject As MSProject.Project
  Dim oSubProject As MSProject.SubProject
  Set oProject = Projects(Me.cboMasterProject.Value)
  For Each oSubProject In oProject.Subprojects
    On Error GoTo next_subproject
    oSubProject.SourceProject.Activate
    FileCloseEx pjPromptSave
next_subproject:
    Err.Clear
  Next oSubProject
  Me.lboSubprojects.Enabled = False
  oProject.Activate
  ActiveWindow.WindowState = pjMaximized
  
exit_here:
  On Error Resume Next
  Set oSubProject = Nothing
  Set oProject = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_frm", "cmdUnload_Click", Err)
  Resume exit_here
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then
    CreateObject("WScript.Shell").Run "https://www.ClearPlanConsulting.com"
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "lblURL", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboMasterLCF_AfterUpdate()
  'objects
  Dim oProject As MSProject.Project
  Dim oOutlineCode As MSProject.OutlineCode
  'strings
  Dim strLookup As String
  Dim strCFN As String
  Dim strFN As String
  'longs
  Dim lngLCF As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnHasLookup As Boolean
  Dim blnHasFormula As Boolean
  Dim blnFilter As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'cptSpeed True
  If IsNull(Me.lboMasterLCF.Value) Then Exit Sub
  lngLCF = CLng(Me.lboMasterLCF.Value)
  strFN = FieldConstantToFieldName(lngLCF)
  strCFN = CustomFieldGetName(lngLCF)
  'clear subproject lookup
  For lngItem = 0 To Me.lboSubprojects.ListCount - 1
    Me.lboSubprojects.List(lngItem, 2) = "..."
    Me.lboSubprojects.List(lngItem, 3) = ""
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
    TableEditEx "LCF FlowDown Table", True, False, False, , , strFN
    TableApply "LCF FlowDown Table"
    ColumnBestFit Column:=1
    ColumnBestFit Column:=2
    SetSplitBar 3
    DoEvents
    SetSplitBar 2
    DoEvents
    FilterClear
    If InStr(strFN, "Cost") > 0 Then
      SetAutoFilter strFN, Criteria1:=">0"
    ElseIf InStr(strFN, "Date") > 0 Then
      SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:="NA"
    ElseIf InStr(strFN, "Duration") > 0 Then
      SetAutoFilter strFN, Criteria1:=">0"
    ElseIf InStr(strFN, "Finish") > 0 Then
      SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:="NA"
    ElseIf InStr(strFN, "Flag") > 0 Then
      SetAutoFilter strFN, FilterType:=pjAutoFilterFlagYes
    ElseIf InStr(strFN, "Number") > 0 Then
      SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:=0
    ElseIf InStr(strFN, "Start") > 0 Then
      SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:="NA"
    Else
      SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:=""
    End If
    SelectBeginning
    DoEvents
  End If
  blnHasFormula = cptHasFormula(ActiveProject, lngLCF)
  Me.lblFormula.Visible = blnHasFormula
  blnHasLookup = cptHasLookup(ActiveProject, lngLCF)
  Me.lblLookup.Visible = blnHasLookup
  If blnHasFormula And blnHasLookup Then
    Me.optNone.Value = True
    MsgBox "LCF '" & strCFN & "' has a formula AND a lookup - please choose which you want to work with in the 'options' section of the form.", vbExclamation + vbOKOnly, "Ambiguity Uncovered"
  ElseIf blnHasFormula Then
    Me.optFormula.Value = True
  ElseIf blnHasLookup Then
    Me.optLookup.Value = True
  Else
    Me.optNone.Value = True
  End If

  'populate subproject LCF CustomFieldName
  Set oProject = Projects(Me.cboMasterProject.Value)
  For lngItem = 0 To Me.lboSubprojects.ListCount - 1
    oLCF.Filter = "SP=" & lngItem + 1 & " AND FC=" & lngLCF
    If Not oLCF.EOF Then
      Me.lboSubprojects.List(lngItem, 2) = oLCF(3)
      If oLCF(3) <> strCFN Then
        Me.lboSubprojects.List(lngItem, 3) = "<!>"
      End If
      DoEvents
    End If
    oLCF.Filter = 0
    If blnFilter Then
      Projects(oProject.Subprojects(lngItem + 1).Path).Activate
      Application.OptionsViewEx DisplayNameIndent:=False, DisplaySummaryTasks:=False
      ViewApply "Gantt Chart", True
      TableEditEx "LCF FlowDown Table", True, True, True, , "Unique ID"
      TableEditEx "LCF FlowDown Table", True, False, False, , , strFN
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
      ElseIf InStr(strFN, "Finish") > 0 Then
        SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:="NA"
      ElseIf InStr(strFN, "Flag") > 0 Then
        SetAutoFilter strFN, FilterType:=pjAutoFilterFlagYes
      ElseIf InStr(strFN, "Number") > 0 Then
        SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:=0
      ElseIf InStr(strFN, "Start") > 0 Then
        SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:="NA"
      Else 'Text, Outline Code
        SetAutoFilter strFN, FilterType:=pjAutoFilterCustom, Test1:="does not equal", Criteria1:=""
      End If
      SelectBeginning
      DoEvents
    End If
  Next lngItem
  Me.cmdTake.Enabled = Not IsNull(Me.lboMasterLCF.Value)
  
  Projects(Me.cboMasterProject.Value).Activate

exit_here:
  On Error Resume Next
  ScreenUpdating = True
  Calculation = pjAutomatic
  Set oProject = Nothing
  Set oOutlineCode = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_frm", "lboMasterLCF_AfterUpdate", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboSubprojects_AfterUpdate()
  Dim oProject As MSProject.Project
  Dim lngLCF As Long
  Dim blnHasFormula As Boolean
  Dim blnHasLookup As Boolean
  Dim strCFN As String
  
  Me.lblFormula1.Visible = False
  Me.lblFormula1.Left = Me.lblLookup2.Left
  Me.lblLookup2.Visible = False
  If IsNull(Me.lboMasterLCF) Then Exit Sub
  If Not IsNull(Me.lboSubprojects.ListIndex) Then
    Set oProject = Projects(Me.lboSubprojects.List(Me.lboSubprojects.ListIndex, 1))
    oProject.Activate
    Me.tglMove.Enabled = True
    cptUpdateMoveToLCF Me, oLCF
    Me.lboMoveLCF.Clear
    strCFN = Me.lboSubprojects.List(Me.lboSubprojects.ListIndex, 2)
    lngLCF = FieldNameToFieldConstant(strCFN)
    blnHasFormula = cptHasFormula(oProject, lngLCF)
    blnHasLookup = cptHasLookup(oProject, lngLCF)
    If blnHasFormula And blnHasLookup Then
      Me.lblFormula1.Left = Me.cmdMoveLCF.Left
      Me.optNone.Value = True
      MsgBox "LCF '" & strCFN & "' has a formula AND a lookup - please choose which you want to work with in the 'options' section of the form.", vbExclamation + vbOKOnly, "Ambiguity Uncovered"
    ElseIf blnHasFormula Then
      Me.optFormula.Value = True
    ElseIf blnHasLookup Then
      Me.optLookup.Value = True
    Else
      Me.optNone.Value = True
    End If
    Me.lblFormula1.Visible = blnHasFormula
    Me.lblLookup2.Visible = blnHasLookup
  Else
    Me.tglMove.Enabled = False
  End If
  Set oProject = Nothing
End Sub

Private Sub stxtSearch_Change()
  Dim oProject As MSProject.Project
  Dim oSubProject As MSProject.SubProject
  Dim strText As String
  strText = Me.stxtSearch.Text
  Me.lboSubprojects.Clear
  Set oProject = Projects(Me.cboMasterProject.Value)
  If Len(Me.stxtSearch.Text) > 0 Then
    Me.lboMasterLCF.ListIndex = -1
    Me.lboMasterLCF.Enabled = False
    Me.cmdTake.Enabled = False
    Me.tglMove = False
    Me.tglMove.Enabled = False
    Me.chkCost.Enabled = False
    Me.chkDate.Enabled = False
    Me.chkDuration.Enabled = False
    Me.chkFinish.Enabled = False
    Me.chkFlag.Enabled = False
    Me.chkNumber.Enabled = False
    Me.chkOutlineCode.Enabled = False
    Me.chkStart.Enabled = False
    Me.chkText.Enabled = False
    Me.optFormula.Enabled = False
    Me.optLookup.Enabled = False
    Me.optNone.Enabled = False
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
    Me.cmdTake.Enabled = Not IsNull(Me.lboMasterLCF.Value)
    Me.chkCost.Enabled = True
    Me.chkDate.Enabled = True
    Me.chkDuration.Enabled = True
    Me.chkFinish.Enabled = True
    Me.chkFlag.Enabled = True
    Me.chkNumber.Enabled = True
    Me.chkOutlineCode.Enabled = True
    Me.chkStart.Enabled = True
    Me.chkText.Enabled = True
    Me.optFormula.Enabled = True
    Me.optLookup.Enabled = True
    Me.optNone.Enabled = True
    For Each oSubProject In oProject.Subprojects
      Me.lboSubprojects.AddItem
      Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 0) = oSubProject.Index
      Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 1) = oSubProject.InsertedProjectSummary.Name
      Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 2) = "..."
    Next oSubProject
  End If
  oLCF.Filter = 0
  Set oSubProject = Nothing
  Set oProject = Nothing
End Sub

Private Sub tglMove_Click()
  Me.lblMoveTo.Visible = Me.tglMove
  Me.lboHeader.Visible = Me.tglMove
  Me.lboMoveLCF.Visible = Me.tglMove
  Me.cboMoveTo.Visible = Me.tglMove
  Me.cmdMoveLCF.Visible = Me.tglMove
  If Me.tglMove Then
    Me.cboMoveTo.Clear
    Me.lboMoveLCF.Clear
    Me.lboHeader.Clear
    Me.lboHeader.AddItem
    Me.lboHeader.List(0, 0) = "UID"
    Me.lboHeader.List(0, 1) = "FROM"
    Me.lboHeader.List(0, 2) = "> TO"
    Me.lboMoveLCF.Top = Me.lboHeader.Top + Me.lboHeader.Height + 1
    Me.lboSubprojects.Height = 108.75
    Me.lboSubprojects.IntegralHeight = True
    Me.lboMoveLCF.Height = Me.lboMasterLCF.Height - (Me.lboMoveLCF.Top - Me.lboMasterLCF.Top)
    Me.lboMoveLCF.IntegralHeight = True
    If Not IsNull(Me.lboSubprojects.Value) Then cptUpdateMoveToLCF Me, oLCF
  Else
    Me.lboSubprojects.Height = Me.lboMasterLCF.Height
    Me.lboSubprojects.IntegralHeight = True
  End If
End Sub

Private Sub UserForm_Initialize()
  'objects
  Dim oProject As MSProject.Project
  Dim oSubProject As MSProject.SubProject
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
      vLCF = cptGetCustomFields("t", "Cost,Date,Duration,Finish,Flag,Number,Outline Code,Start,Text", "c,fn,cfn")
      For lngRow = 0 To UBound(vLCF, 1)
        oLCF.AddNew Array(0, 1, 2, 3), Array(0, CLng(vLCF(lngRow, 0)), vLCF(lngRow, 1), vLCF(lngRow, 2))
      Next lngRow
      'get subprojects
      For Each oSubProject In oProject.Subprojects
        Debug.Print oSubProject.SourceProject.Tasks.Count 'forces load
        If oSubProject.IsLoaded = True Then
          'adjust the view
          On Error Resume Next
          FileOpenEx oSubProject.Path, False
          Projects(oSubProject.SourceProject.Name).Activate
          ActiveWindow.TopPane.Activate
          FilterClear
          GroupClear
          Sort "ID", , , , , , False, True
          OptionsViewEx DisplaySummaryTasks:=True
          OutlineShowAllTasks
          'add to lboSubprojects
          Me.lboSubprojects.AddItem
          Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 0) = oSubProject.Index
          Me.lboSubprojects.List(Me.lboSubprojects.ListCount - 1, 1) = oSubProject.InsertedProjectSummary.Name
          If oSubProject.IsLoaded = False Then FileOpenEx oSubProject.Path, False
          vLCF = cptGetCustomFields("t", "Cost,Date,Duration,Finish,Flag,Number,Outline Code,Start,Text", "c,fn,cfn")
          For lngRow = 0 To UBound(vLCF, 1)
            oLCF.AddNew Array(0, 1, 2, 3), Array(oSubProject.Index, CLng(vLCF(lngRow, 0)), vLCF(lngRow, 1), vLCF(lngRow, 2))
          Next lngRow
          oLCF.Update
        End If
      Next oSubProject
      oProject.Activate
      WindowArrangeAll
    End If
  Next oProject
  
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
  Me.lblLookup.Visible = False
  Me.lblLookup.Top = Me.lboMasterLCF.Top + Me.lboMasterLCF.Height + 1
  Me.cmdTake.Enabled = False
  
exit_here:
  On Error Resume Next
  Set oSubProject = Nothing
  Set oProject = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_frm", "UserForm_Initialize", Err, Erl)
  Resume exit_here
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Set oLCF = Nothing
  If Not IsNull(Me.cboMasterProject.Value) Then
    If Me.cboMasterProject.Value <> "" Then Projects(Me.cboMasterProject.Value).Activate
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
