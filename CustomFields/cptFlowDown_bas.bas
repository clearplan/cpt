Attribute VB_Name = "cptFlowDown_bas"
'<cpt_version>v0.0.0</cpt_version>
Option Explicit

Sub cptShowLCFFlowDown_frm()
  Dim oProject As MSProject.Project
  Dim myFlowDown_frm As cptFlowDown_frm
  Set myFlowDown_frm = New cptFlowDown_frm
  With myFlowDown_frm
    For Each oProject In Projects
      If oProject.Subprojects.Count > 0 Then
        .cboMasterProject.AddItem oProject.Name
      End If
    Next
    .lblFormula.Visible = False
    .lblLookup.Visible = False
    .lboSubprojects.Height = .lboMasterLCF.Height
    .lblMoveTo.Visible = False
    .cboMoveTo.Visible = False
    .lboMoveLCF.Visible = False
    .cmdMoveLCF.Visible = False
    .tglMove = False
    .Show
  End With
  Set oProject = Nothing
  Set myFlowDown_frm = Nothing
End Sub

Sub cptUpdateMoveToLCF(ByRef myFlowDown_frm As cptFlowDown_frm, ByRef oLCF As ADODB.Recordset)
  Dim strType As String
  Dim strCFN As String
  Dim strFN As String
  'todo: move Cost -> Cost
  'todo: move Date -> Date
  'todo: move Duration -> Duration
  'todo: move Finish -> Finish
  'todo: move Flag -> Flag
  'todo: move Number -> Number
  'todo: move Outline Code -> Outline Code
  'todo: move Start -> Start
  'todo: move Text -> Text
  With myFlowDown_frm
    If IsNull(.lboMasterLCF.Value) Then Exit Sub
    If .cboMoveTo.Visible Then
      strType = cptRegEx(FieldConstantToFieldName(.lboMasterLCF.Value), "[A-z]*")
      .cboMoveTo.Clear
      .cboMoveTo.ColumnCount = 2
      .cboMoveTo.ColumnWidths = "0;"
      .lboHeader.Clear
      .lboHeader.AddItem
      .lboHeader.List(0, 0) = "UID"
      .lboHeader.List(0, 1) = "FROM"
      .lboHeader.List(0, 2) = "> TO"
      oLCF.Filter = 0
      oLCF.Filter = "SP=" & .lboSubprojects.Value & " AND FN Like '" & strType & "*'"
      If Not oLCF.EOF Then
        oLCF.MoveFirst
        Do While Not oLCF.EOF
          If oLCF(1) <> CLng(.lboMasterLCF.Value) Then
            .cboMoveTo.AddItem
            .cboMoveTo.List(.cboMoveTo.ListCount - 1, 0) = oLCF(1)
            strCFN = FieldConstantToFieldName(oLCF(1))
            If oLCF(2) <> oLCF(3) Then
              .cboMoveTo.List(.cboMoveTo.ListCount - 1, 1) = oLCF(2) & " (" & oLCF(3) & ")"
            Else
              .cboMoveTo.List(.cboMoveTo.ListCount - 1, 1) = oLCF(2)
            End If
          End If
          oLCF.MoveNext
        Loop
      End If
      oLCF.Filter = 0
    End If
  End With
End Sub

Sub cptUpdateLCFMoveToPreview(ByRef myFlowDown_frm As cptFlowDown_frm)
  Dim oProject As MSProject.Project
  Dim oTask As MSProject.Task
  Dim lngFrom As Long
  Dim lngTo As Long
  With myFlowDown_frm
    If IsNull(.cboMoveTo.Value) Then Exit Sub
    Set oProject = Projects(.lboSubprojects.List(.lboSubprojects.ListIndex, 1))
    oProject.Activate
    
    lngFrom = FieldNameToFieldConstant(.lboSubprojects.List(.lboSubprojects.ListIndex, 2))
    lngTo = .cboMoveTo.Value
    
    If .chkFilter Then
      TableEditEx "LCF FlowDown Table", True, True, True, , "Unique ID"
      TableEditEx "LCF FlowDown Table", True, False, False, , , FieldConstantToFieldName(lngFrom)
      TableEditEx "LCF FlowDown Table", True, False, False, , , FieldConstantToFieldName(lngTo)
      TableApply "LCF FlowDown Table"
      ColumnBestFit Column:=1
      ColumnBestFit Column:=2
      ColumnBestFit Column:=3
      SetSplitBar 4
      DoEvents
    End If
    
    .lboHeader.Clear
    .lboHeader.AddItem
    .lboHeader.List(.lboHeader.ListCount - 1, 0) = "UID"
    .lboHeader.List(.lboHeader.ListCount - 1, 1) = FieldConstantToFieldName(lngFrom)
    .lboHeader.List(.lboHeader.ListCount - 1, 2) = "> " & FieldConstantToFieldName(lngTo)
    .lboMoveLCF.Clear
    .lboMoveLCF.ColumnCount = 3
    For Each oTask In oProject.Tasks
      If oTask Is Nothing Then GoTo next_task
      If Not oTask.Active Then GoTo next_task
      If oTask.ExternalTask Then GoTo next_task
      .lboMoveLCF.AddItem
      .lboMoveLCF.List(.lboMoveLCF.ListCount - 1, 0) = oTask.UniqueID
      .lboMoveLCF.List(.lboMoveLCF.ListCount - 1, 1) = oTask.GetField(lngFrom)
      .lboMoveLCF.List(.lboMoveLCF.ListCount - 1, 2) = "> " & oTask.GetField(lngTo)
next_task:
    Next oTask
    If Not IsNull(.lboSubprojects) And Not IsNull(.cboMoveTo) Then
      .cmdMoveLCF.Enabled = True
    Else
      .cmdMoveLCF.Enabled = False
    End If
    .lboMoveLCF.IntegralHeight = True
  End With
  Set oTask = Nothing
  Set oProject = Nothing
End Sub

Sub cptMoveLCF(ByRef myFlowDown_frm As cptFlowDown_frm, ByRef oLCF As ADODB.Recordset)
  Dim oProject As MSProject.Project
  Dim oTask As MSProject.Task
  Dim lngFrom As Long
  Dim strFromType As String
  Dim lngTo As Long
  Dim strToType As String
  Dim lngItem As Long
  Dim lngUID As Long
  Dim strCFN As String
  'todo: if lookup
  'todo: if outlinecode
  'todo: if formula
  '   > check all other fields for references to lngFrom and replace with lngTo
  'todo: if formula AND lookup
  'todo: move Cost -> Cost,Number,Text?
  'todo: move Date -> Date,Start,Finish
  'todo: move Duration -> Duration,Number
  '   > if converting, then convert data accordingly
  'todo: move Finish -> Finish,Date,Start?
  'todo: move Flag -> Flag,Number(0,1),Text
  '   > if converting, then convert data accordinlgy
  'todo: move Number -> Number
  'todo: move Outline Code -> Outline Code
  'todo: move Start -> Start,Date,Finish?
  'todo: move Text -> Text
  
  cptSpeed True
  With myFlowDown_frm
    Set oProject = Projects(.lboSubprojects.List(.lboSubprojects.ListIndex, 1))
    oProject.Activate
    lngFrom = FieldNameToFieldConstant(.lboSubprojects.List(.lboSubprojects.ListIndex, 2))
    strFromType = cptRegEx(FieldConstantToFieldName(lngFrom), "(Cost|Date|Duration|Finish|Flag|Number|Outline Code|Start|Text)")
    lngTo = .cboMoveTo.Value
    strToType = cptRegEx(FieldConstantToFieldName(lngTo), "(Cost|Date|Duration|Finish|Flag|Number|Outline Code|Start|Text)")
    strCFN = CustomFieldGetName(lngFrom)
    CustomFieldRename lngFrom, strCFN & "_OLD"
    CustomFieldDelete lngTo
    CustomFieldRename lngTo, strCFN
    'todo: move formula
    'todo: move lookup
    'todo: move Outline Code
    For lngItem = 0 To .lboMoveLCF.ListCount - 1
      lngUID = .lboMoveLCF.List(lngItem, 0)
      Set oTask = oProject.Tasks.UniqueID(lngUID)
      If Len(.lboMoveLCF.List(lngItem, 1)) > 0 Then
        oTask.SetField lngTo, oTask.GetField(lngFrom)
        .lboMoveLCF.List(lngItem, 2) = .lboMoveLCF.List(lngItem, 1)
      Else
        oTask.SetField lngTo, ""
        .lboMoveLCF.List(lngItem, 2) = ""
      End If
      .lboMoveLCF.List(lngItem, 1) = ""
      oTask.SetField lngFrom, ""
next_task:
      Application.StatusBar = "Moving LCF...(" & Format(lngItem / .lboMoveLCF.ListCount - 1, "0%") & ")"
      DoEvents
    Next lngItem
    'update oLCF
    oLCF.Filter = 0
    oLCF.Filter = "SP=" & .lboSubprojects.Value & " AND FC=" & lngFrom
    oLCF(3) = FieldConstantToFieldName(lngFrom)
    oLCF.Filter = "SP=" & .lboSubprojects.Value & " AND FC=" & lngTo
    oLCF(3) = strCFN
    oLCF.Filter = 0
    'update lboSubprojects
    .lboSubprojects.List(.lboSubprojects.ListIndex, 2) = FieldConstantToFieldName(lngFrom)
    'update cboMoveTo
    For lngItem = 0 To .cboMoveTo.ListCount - 1
      If .cboMoveTo.List(lngItem, 0) = lngFrom Then
        .cboMoveTo.List(lngItem, 1) = FieldConstantToFieldName(lngFrom)
      ElseIf .cboMoveTo.List(lngItem, 0) = lngTo Then
        .cboMoveTo.List(lngItem, 1) = FieldConstantToFieldName(lngTo) & " (" & strCFN & ")"
      End If
    Next lngItem
    DoEvents
  End With
  cptSpeed False
  Application.StatusBar = "Moving LCF...done."
  CustomFieldDelete lngFrom
  Set oTask = Nothing
  Set oProject = Nothing
End Sub

Sub cptUpdateLCFList(ByRef myFlowDown_frm As cptFlowDown_frm)
  Dim strDataTypes As String
  Dim oControl As Control
  Dim lngItem As Long
  With myFlowDown_frm
    For lngItem = 0 To .lboSubprojects.ListCount - 1
      .lboSubprojects.List(lngItem, 2) = "..."
    Next lngItem
    For Each oControl In .Controls
      If Left(oControl.Name, 3) = "chk" Then
        If oControl.Name = "chkFilter" Then GoTo next_control
        If .Controls(oControl.Name).Value = True Then
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
      .lboMasterLCF.List = cptGetCustomFields("t", strDataTypes, "c,fn,cfn")
    Else
      .lboMasterLCF.Clear
    End If
  End With
End Sub
