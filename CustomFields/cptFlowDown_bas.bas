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
    .lblPickList.Visible = False
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
  With myFlowDown_frm
    If IsNull(.lboMasterLCF.Value) Then Exit Sub
    If .cboMoveTo.Visible Then
      strType = cptRegEx(FieldConstantToFieldName(.lboMasterLCF.Value), "[A-z]*")
      .cboMoveTo.Clear
      .cboMoveTo.ColumnCount = 2
      .cboMoveTo.ColumnWidths = "0;"
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
    
    .lboMoveLCF.Clear
    .lboMoveLCF.ColumnCount = 3
    .lboMoveLCF.AddItem
    .lboMoveLCF.List(.lboMoveLCF.ListCount - 1, 0) = "UID"
    .lboMoveLCF.List(.lboMoveLCF.ListCount - 1, 1) = FieldConstantToFieldName(lngFrom)
    .lboMoveLCF.List(.lboMoveLCF.ListCount - 1, 2) = "> " & FieldConstantToFieldName(lngTo)
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
  End With
  Set oTask = Nothing
  Set oProject = Nothing
End Sub

Sub cptMoveLCF(ByRef myFlowDown_frm As cptFlowDown_frm, ByRef oLCF As ADODB.Recordset)
  Dim oProject As MSProject.Project
  Dim oTask As MSProject.Task
  Dim lngFrom As Long
  Dim lngTo As Long
  Dim lngItem As Long
  Dim lngUID As Long
  Dim strCFN As String
  'todo: readonly?
  'todo: if picklist
  'todo: if outlinecode
  'todo: if formula
  
  cptSpeed True
  With myFlowDown_frm
    Set oProject = Projects(.lboSubprojects.List(.lboSubprojects.ListIndex, 1))
    oProject.Activate
    lngFrom = FieldNameToFieldConstant(.lboSubprojects.List(.lboSubprojects.ListIndex, 2))
    lngTo = .cboMoveTo.Value
    strCFN = CustomFieldGetName(lngFrom)
    CustomFieldRename lngFrom, strCFN & "_OLD"
    CustomFieldDelete lngTo
    CustomFieldRename lngTo, strCFN
    For lngItem = 1 To .lboMoveLCF.ListCount - 1
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
