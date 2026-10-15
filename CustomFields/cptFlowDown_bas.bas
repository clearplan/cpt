Attribute VB_Name = "cptFlowDown_bas"
'<cpt_version>v0.0.1</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptFlowDown_bas"

Sub cptShowFlowDown_frm()
  Dim oProject As MSProject.Project
  Dim myFlowDown_frm As cptFlowDown_frm
  Dim blnMaster As Boolean
  
  'check for an update
  'If Not cptProceedOnUpdate(THIS_MODULE) Then GoTo exit_here
  
  'prevent spawning
  If Not cptGetUserForm("cptFiscal_frm") Is Nothing Then Exit Sub
  
  'the boolean should never actually come into play since
  'the ribbon menu only appears if the activeproject has subprojects
  'but just in case someone triggers this manually...
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  Set myFlowDown_frm = New cptFlowDown_frm
  blnMaster = False
  With myFlowDown_frm
    .Caption = "LCF FlowDown (" & cptGetVersion("cptFlowDown_bas") & ") - ClearPlan Toolbar"
    For Each oProject In Projects
      If oProject.Subprojects.Count > 0 Then
        .cboMasterProject.AddItem oProject.Name
        blnMaster = True
      End If
    Next
    If Not blnMaster Then GoTo exit_here
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
exit_here:
  On Error Resume Next
  Set oProject = Nothing
  Set myFlowDown_frm = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_bas", "cptShowLCFFlowDown_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptAssignmentRolldown(myFlowDown_frm As cptFlowDown_frm, blnRolldown As Boolean)
  'objects
  Dim oProject As MSProject.Project
  'longs
  Dim lngItem As Long
  Dim lngLCF As Long
  'strings
  Dim strProject As String
  
  With myFlowDown_frm
    If IsNull(.lboMasterLCF) Then GoTo exit_here
    For lngItem = 0 To .lboSubprojects.ListCount - 1
      strProject = .lboSubprojects.List(lngItem, 1)
      lngLCF = .lboMasterLCF
      Application.Projects(strProject).Activate
      CustomFieldPropertiesEx FieldID:=lngLCF, AutomaticallyRolldownToAssn:=blnRolldown
    Next lngItem
  End With
  
exit_here:
  On Error Resume Next
  Set oProject = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_bas", "cptAssignmentRolldown", Err)
  Resume exit_here

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
  'objects
  Dim oLookupTableTo As MSProject.LookupTable
  Dim oLookupTableFrom As MSProject.LookupTable
  Dim oOutlineCodeFrom As MSProject.OutlineCode
  Dim oOutlineCodeTo As MSProject.OutlineCode
  Dim oSubProject As MSProject.Project
  Dim oTask As MSProject.Task
  'strings
  Dim strValue As String
  Dim strDescription As String
  Dim strFormula As String
  Dim strCFN As String
  Dim strFromType As String
  Dim strToType As String
  'longs
  Dim lngFrom As Long
  Dim lngTo As Long
  Dim lngItem As Long
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  Dim blnHasFormula As Boolean
  Dim blnHasLookup As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    
  'todo: if formula
  '   > check all other fields for references to lngFrom and replace with lngTo
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
    
    lngFrom = FieldNameToFieldConstant(.lboSubprojects.List(.lboSubprojects.ListIndex, 2))
    strFromType = cptRegEx(FieldConstantToFieldName(lngFrom), "(Cost|Date|Duration|Finish|Flag|Number|Outline Code|Start|Text)")
    lngTo = .cboMoveTo.Value
    strToType = cptRegEx(FieldConstantToFieldName(lngTo), "(Cost|Date|Duration|Finish|Flag|Number|Outline Code|Start|Text)")
    
    Set oSubProject = Projects(.lboSubprojects.List(.lboSubprojects.ListIndex, 1))
    oSubProject.Activate
    
    blnHasFormula = cptHasFormula(oSubProject, lngFrom)
    If blnHasFormula Then 'get it
      strFormula = CustomFieldGetFormula(lngFrom)
    End If
    'todo: if strFormula contains [Cost|Date|Duration|Finish|Flag|Number|Outline Code|Start]
    'todo: has the other custom field been moved already?
    'todo: replace formula with CFN and prompt?
    blnHasLookup = cptHasLookup(oSubProject, lngFrom)
    
    strCFN = CustomFieldGetName(lngFrom)
    CustomFieldRename lngFrom, strCFN & "_OLD"
    CustomFieldDelete lngTo
    CustomFieldRename lngTo, strCFN
    If blnHasFormula Then
      CustomFieldDelete lngFrom
      CustomFieldSetFormula lngTo, strFormula
    End If
    If .optFormula Then CustomFieldPropertiesEx lngTo, pjFieldAttributeFormula
    
    'move lookup
    If blnHasLookup Then
      On Error Resume Next
      Set oOutlineCodeFrom = oSubProject.OutlineCodes(strCFN & "_OLD")
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If Not oOutlineCodeFrom Is Nothing Then
        'first copy the codemask
        Set oOutlineCodeTo = oSubProject.OutlineCodes(strCFN)
        For lngItem = 1 To oOutlineCodeFrom.CodeMask.Count
          oOutlineCodeTo.CodeMask.Add oOutlineCodeFrom.CodeMask(lngItem).Sequence, oOutlineCodeFrom.CodeMask(lngItem).Length, oOutlineCodeFrom.CodeMask(lngItem).Separator
        Next lngItem
        Set oLookupTableFrom = oOutlineCodeFrom.LookupTable
        Set oLookupTableTo = oOutlineCodeTo.LookupTable
        'load from bottom to top
        For lngItem = oLookupTableFrom.Count To 1 Step -1
          oLookupTableTo.AddChild oLookupTableFrom.Item(lngItem).Name
        Next lngItem
        'indent from top to bottom
        For lngItem = 1 To oLookupTableFrom.Count
          oLookupTableTo.Item(lngItem).Level = oLookupTableFrom.Item(lngItem).Level
          oLookupTableTo.Item(lngItem).Description = oLookupTableFrom.Item(lngItem).Description
        Next lngItem
        If Len(oOutlineCodeFrom.DefaultValue) > 0 Then
          oOutlineCodeTo.DefaultValue = oOutlineCodeFrom.DefaultValue
        End If
        'todo: some settings must be made manually
        oOutlineCodeTo.OnlyLookUpTableCodes = oOutlineCodeFrom.OnlyLookUpTableCodes
        oOutlineCodeTo.OnlyCompleteCodes = oOutlineCodeFrom.OnlyCompleteCodes
        If oOutlineCodeFrom.OnlyCompleteCodes Then
          oOutlineCodeTo.OnlyLeaves = oOutlineCodeFrom.OnlyLeaves
          oOutlineCodeTo.RequiredCode = oOutlineCodeFrom.RequiredCode
        End If
        oOutlineCodeTo.SortOrder = oOutlineCodeFrom.SortOrder
      Else 'picklist
        'todo: some settings must be made manually
        On Error Resume Next
        For lngItem = 1 To 1000 'surely 1000 items is enough?
          strValue = CustomFieldValueListGetItem(lngFrom, pjValueListValue, lngItem)
          If Len(strValue) = 0 Then Exit For
          strDescription = CustomFieldValueListGetItem(lngFrom, pjValueListDescription, lngItem)
          CustomFieldValueListAdd lngTo, strValue, strDescription
          strValue = ""
          strDescription = ""
        Next lngItem
      End If
    End If
    If .optLookup Then CustomFieldPropertiesEx lngTo, pjFieldAttributeValueList
    
    'now move task-level data (if not a formula
    For lngItem = 0 To .lboMoveLCF.ListCount - 1
      lngUID = .lboMoveLCF.List(lngItem, 0)
      Set oTask = oSubProject.Tasks.UniqueID(lngUID)
      If .optFormula Then
        oTask.SetField lngFrom, ""
      Else
        If Len(.lboMoveLCF.List(lngItem, 1)) > 0 Then
          'todo: does this string-only operation move Cost and Duration accurately?
          oTask.SetField lngTo, oTask.GetField(lngFrom)
          .lboMoveLCF.List(lngItem, 2) = .lboMoveLCF.List(lngItem, 1)
        Else
          oTask.SetField lngTo, ""
          .lboMoveLCF.List(lngItem, 2) = ""
        End If
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
  Application.StatusBar = "Moving LCF...done."
  CustomFieldDelete lngFrom
  myFlowDown_frm.lblFormula1.Visible = False
  myFlowDown_frm.lblLookup2.Visible = False

exit_here:
  On Error Resume Next
  cptSpeed False
  Application.StatusBar = ""
  Set oLookupTableTo = Nothing
  Set oLookupTableFrom = Nothing
  Set oTask = Nothing
  Set oOutlineCodeFrom = Nothing
  Set oOutlineCodeTo = Nothing
  Set oSubProject = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_bas", "cptMoveLCF", Err, Erl)
  Resume exit_here
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

Sub cptTakeLCF(ByRef myFlowDown_frm As cptFlowDown_frm, ByRef oLCF As ADODB.Recordset)
  'objects
  Dim oDict As Scripting.Dictionary
  Dim oOutlineCodeFrom As MSProject.OutlineCode
  Dim oOutlineCodeTo As MSProject.OutlineCode
  Dim oMP As MSProject.Project
  Dim oSP As MSProject.Project
  Dim oTask As MSProject.Task
  'strings
  Dim strDescription As String
  Dim strValue As String
  Dim strFN As String
  Dim strMsg As String
  Dim strFormula As String
  Dim strCFN As String
  'longs
  Dim lngIndex As Long
  Dim lngLookupItem As Long
  Dim lngLCF As Long
  Dim lngItem As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnHasLookup As Boolean
  Dim blnHasFormula As Boolean
  Dim blnSingle As Boolean
  Dim blnClearTaskData As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  Select Case MsgBox("Take LCF for ALL subprojects?", vbExclamation + vbYesNoCancel, "Take how much?")
    Case vbYes
      blnSingle = False
    Case vbNo
      blnSingle = True
    Case vbCancel
      GoTo exit_here
  End Select
  If IsNull(myFlowDown_frm.lboMasterLCF.Value) Then GoTo exit_here
  lngLCF = myFlowDown_frm.lboMasterLCF.Value
  If lngLCF = 0 Then Exit Sub
  Set oMP = Projects(myFlowDown_frm.cboMasterProject.Value)
  oMP.Activate
  strFN = FieldConstantToFieldName(lngLCF)
  strCFN = CustomFieldGetName(lngLCF)
  If Not myFlowDown_frm.optFormula Then
    Select Case MsgBox("Clear subproject task data in LCF '" & strCFN & "' (" & strFN & ")?", vbQuestion + vbYesNoCancel, "Overwrite?")
      Case vbYes
        blnClearTaskData = True
      Case vbNo
        blnClearTaskData = False
      Case vbCancel
        GoTo exit_here
    End Select
  End If
  cptSpeed True
  For lngItem = 0 To myFlowDown_frm.lboSubprojects.ListCount - 1
    If blnSingle And Not myFlowDown_frm.lboSubprojects.Selected(lngItem) Then GoTo next_subproject
    myFlowDown_frm.lboSubprojects.List(lngItem, 2) = "..."
    DoEvents
    Set oSP = Projects(myFlowDown_frm.lboSubprojects.List(lngItem, 1)) 'name
    lngIndex = myFlowDown_frm.lboSubprojects.List(lngItem, 0) 'index
    lngTask = 0
    lngTasks = oSP.Tasks.Count
    oSP.Activate
    'delete the LCF
    CustomFieldDelete lngLCF
    'rename the LCF
    oLCF.Filter = 0
    oLCF.Filter = "SP=" & lngIndex & " AND CFN='" & strCFN & "' AND FC<>" & lngLCF
    If oLCF.EOF Then
      CustomFieldRename lngLCF, strCFN
      Stop
      oLCF.Filter = 0
      oLCF.Filter = "SP=" & lngIndex & " AND FC=" & lngLCF
      If Not oLCF.EOF Then
        oLCF(3) = strCFN
      End If
      oLCF.Filter = 0
      myFlowDown_frm.lboSubprojects.List(lngItem, 2) = strCFN
      myFlowDown_frm.lboSubprojects.List(lngItem, 3) = ""
    Else 'already exists in this subproject
      strMsg = "Subproject '" & oMP.Subprojects(oLCF(0)).InsertedProjectSummary.Name & "' already has a '" & oLCF(3) & "' in " & oLCF(2) & "!" & vbCrLf & vbCrLf
      strMsg = strMsg & "Rename " & oLCF(2) & " '" & oLCF(3) & "_OLD' and proceed?" & vbCrLf & vbCrLf
      If MsgBox(strMsg, vbExclamation + vbOKCancel, "Custom Field Name Already Exists") = vbCancel Then
        oLCF.Filter = 0
        GoTo exit_here
      Else
        CustomFieldRename oLCF(1), strCFN & "_OLD"
        DoEvents
        'update oLCF for old
        oLCF(3) = strCFN & "_OLD"
        'update oLCF for new
        oLCF.Filter = 0
        oLCF.Filter = "SP=" & lngIndex & " AND FC=" & lngLCF
        oLCF(3) = strCFN
        CustomFieldRename lngLCF, strCFN
        'update listbox
        myFlowDown_frm.lboSubprojects.List(lngItem, 2) = strCFN
        myFlowDown_frm.lboSubprojects.List(lngItem, 3) = ""
      End If
    End If
    oLCF.Filter = 0
    'formula
    blnHasFormula = cptHasFormula(oMP, lngLCF)
    If blnHasFormula Then
      'todo: formula dependencies (other fields)
      oMP.Activate
      strFormula = CustomFieldGetFormula(lngLCF)
      oSP.Activate
      CustomFieldSetFormula lngLCF, strFormula
    End If
    blnHasLookup = cptHasLookup(oMP, lngLCF)
    If blnHasLookup Then
      'move lookup
      If InStr(strFN, "Outline Code") > 0 Then
        oMP.Activate
        Set oOutlineCodeFrom = oMP.OutlineCodes(strCFN)
        oSP.Activate
        Set oOutlineCodeTo = oSP.OutlineCodes(strCFN)
        'first copy the code mask
        For lngLookupItem = 1 To oOutlineCodeFrom.CodeMask.Count
          oOutlineCodeTo.CodeMask.Add oOutlineCodeFrom.CodeMask(lngLookupItem).Sequence, oOutlineCodeFrom.CodeMask(lngLookupItem).Length, oOutlineCodeFrom.CodeMask.Item(lngLookupItem).Separator
        Next lngLookupItem
        'then copy the lookup table items (bottom-up)
        For lngLookupItem = oOutlineCodeFrom.LookupTable.Count To 1 Step -1
          oOutlineCodeTo.LookupTable.AddChild oOutlineCodeFrom.LookupTable.Item(lngLookupItem).Name
        Next lngLookupItem
        'then indent the lookup table structure (top-down)
        For lngLookupItem = 1 To oOutlineCodeTo.LookupTable.Count
          oOutlineCodeTo.LookupTable.Item(lngLookupItem).Level = oOutlineCodeFrom.LookupTable.Item(lngLookupItem).Level
          oOutlineCodeTo.LookupTable.Item(lngLookupItem).Description = oOutlineCodeFrom.LookupTable.Item(lngLookupItem).Description
        Next lngLookupItem
        If Len(oOutlineCodeFrom.DefaultValue) > 0 Then
          oOutlineCodeTo.DefaultValue = oOutlineCodeFrom.DefaultValue
        End If
        'some settings must be made manually
        oOutlineCodeTo.OnlyLookUpTableCodes = oOutlineCodeFrom.OnlyLookUpTableCodes
        oOutlineCodeTo.OnlyCompleteCodes = oOutlineCodeFrom.OnlyCompleteCodes
        If oOutlineCodeFrom.OnlyCompleteCodes Then
          oOutlineCodeTo.OnlyLeaves = oOutlineCodeFrom.OnlyLeaves
          oOutlineCodeTo.RequiredCode = oOutlineCodeFrom.RequiredCode
        End If
        oOutlineCodeTo.SortOrder = oOutlineCodeFrom.SortOrder
      Else 'it's a value list
        Set oDict = CreateObject("Scripting.Dictionary")
        'get lookup
        oMP.Activate
        For lngLookupItem = 1 To 1000 'surely that's enough?
          strValue = ""
          strDescription = ""
          On Error Resume Next
          strValue = CustomFieldValueListGetItem(lngLCF, pjValueListValue, lngLookupItem)
          If Len(strValue) = 0 Then Exit For
          strDescription = CustomFieldValueListGetItem(lngLCF, pjValueListDescription, lngLookupItem)
          If Not oDict.Exists(strValue) Then
            oDict.Add strValue, strDescription
          End If
        Next lngLookupItem
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        'set lookup
        oSP.Activate
        For lngLookupItem = 0 To oDict.Count - 1
          CustomFieldValueListAdd lngLCF, oDict.Keys(lngLookupItem), oDict.Items(lngLookupItem)
        Next lngLookupItem
      End If
    End If
    If InStr(strFN, "Outline Code") = 0 Then
      If blnHasLookup And blnHasFormula Then
        oSP.Activate
        strMsg = "LCF '" & strCFN & "' (" & strFN & ") has both a Lookup and a Formula." & vbCrLf
        strMsg = strMsg & "Both will be imported, but you must decide which one should be applied." & vbCrLf & vbCrLf
        strMsg = strMsg & "Click Yes to use the Lookup." & vbCrLf
        strMsg = strMsg & "Click No to use the Formula." & vbCrLf
        strMsg = strMsg & "Click Cancel to use niether (option: None)."
        Select Case MsgBox(strMsg, vbQuestion + vbYesNoCancel, "Please choose:")
          Case vbYes
            CustomFieldPropertiesEx lngLCF, pjFieldAttributeValueList
          Case vbNo
            CustomFieldPropertiesEx lngLCF, pjFieldAttributeFormula
          Case vbCancel
            CustomFieldPropertiesEx lngLCF, pjFieldAttributeNone
        End Select
      ElseIf blnHasLookup Then
        CustomFieldPropertiesEx lngLCF, pjFieldAttributeValueList
      ElseIf blnHasFormula Then
        CustomFieldPropertiesEx lngLCF, pjFieldAttributeFormula
      Else
        CustomFieldPropertiesEx lngLCF, pjFieldAttributeNone
      End If
    End If
    SetAutoFilter strCFN, pjAutoFilterIn, "does not equal", ""
    'todo: use SetField it's much faster
    'todo: or maybe SetTaskField
    If Not myFlowDown_frm.optFormula Then
      If blnClearTaskData Then
        'now delete the data
        For Each oTask In oSP.Tasks
          If oTask Is Nothing Then GoTo next_task
          If oTask.ExternalTask Then GoTo next_task
          If oTask.GetField(lngLCF) <> "" Then oTask.SetField lngLCF, ""
next_task:
          lngTask = lngTask + 1
          myFlowDown_frm.lboSubprojects.List(lngItem, 2) = "(" & Format(lngTask / lngTasks, "0%") & ")"
          DoEvents
        Next oTask
      End If
    End If
    DoEvents
next_subproject:
  Next lngItem

  strMsg = "Certain LCF settings cannot be read programmatically and so must be transferred/applied manually, e.g.:" & vbCrLf & vbCrLf
  strMsg = strMsg & "- Lookups & Outline Codes: display indenting in lookup table (default: true)" & vbCrLf
  strMsg = strMsg & "- Lookups: default value (default: none)" & vbCrLf
  strMsg = strMsg & "- Lookups: allow new values (default: false)" & vbCrLf
  strMsg = strMsg & "- Lookups: display order for lookup table (default: by row number)" & vbCrLf
  strMsg = strMsg & "- Lookups: data entry options (allow new?) (default: false)" & vbCrLf
  strMsg = strMsg & "- Lookups & Outline Codes: data entry options (lowest level only?) (default: false)" & vbCrLf
  strMsg = strMsg & "- calculation for task and group summary rows (default: none)" & vbCrLf
  strMsg = strMsg & "- calculation for assignment rows (default: none)" & vbCrLf
  strMsg = strMsg & "- values to display (data or graphical indicators) (default: data)" & vbCrLf & vbCrLf
  strMsg = strMsg & "Please review and apply these settings manually." & vbCrLf & vbCrLf
  strMsg = strMsg & "Hint: take screenshots of the Custom Fields dialogs between master and sub(s) and compare."
  MsgBox strMsg, vbInformation + vbOKOnly, "One more step..."

exit_here:
  On Error Resume Next
  cptSpeed False
  Set oDict = Nothing
  Set oOutlineCodeFrom = Nothing
  Set oOutlineCodeTo = Nothing
  Set oTask = Nothing
  Set oSP = Nothing
  Set oMP = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFlowDown_bas", "cptTakeLCF", Err, Erl)
  Resume exit_here
End Sub
