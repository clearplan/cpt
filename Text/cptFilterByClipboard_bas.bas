Attribute VB_Name = "cptFilterByClipboard_bas"
'<cpt_version>v1.4.0</cpt_version>
Option Explicit

Sub cptShowFilterByClipboard_frm()
  'objects
  Dim myFilterByClipboard_frm As cptFilterByClipboard_frm
  'strings
  Dim strMsg As String
  Dim strFreeField As String
  'longs
  Dim lngFreeField As Long
  'integers
  'doubles
  'booleans
  Dim blnMaster As Boolean
  'variants
  'dates
  
  'prevent spawning
  If Not cptGetUserForm("cptFilterByClipboard_frm") Is Nothing Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If ActiveProject.Tasks.Count = 0 Then GoTo exit_here
  blnMaster = ActiveProject.Subprojects.Count > 0
  
  Set myFilterByClipboard_frm = New cptFilterByClipboard_frm
  
  With myFilterByClipboard_frm
    .Caption = "Filter By Clipboard (" & cptGetVersion("cptFilterByClipboard_frm") & ")"
    .tglEdit = False
    .lboHeader.Height = 12.5
    .lboHeader.Clear
    .lboHeader.AddItem
    .optUID = True
    .optID.Enabled = Not blnMaster
    .lboHeader.List(.lboHeader.ListCount - 1, 0) = "UID"
    .lboHeader.List(.lboHeader.ListCount - 1, 1) = "Task Name"
    .lboHeader.Width = .lboFilter.Width
    .lboHeader.ColumnCount = 2
    .lboHeader.ColumnWidths = 45
    .lboFilter.Top = .lboHeader.Top + .lboHeader.Height
    .lboFilter.ColumnCount = 2
    .lboFilter.ColumnWidths = 45
    .txtFilter.Top = .lboFilter.Top
    .txtFilter.Width = .lboFilter.Width
    .txtFilter.Height = .lboFilter.Height
    .txtFilter.Visible = True
    .lboFilter.Visible = False
    .chkFilter = True
  End With
  
  strFreeField = cptGetSetting("FilterByClipboard", "cboFreeField")
  If Len(strFreeField) > 0 Then
    'is it named cptFilterByClipboard?
    If CustomFieldGetName(CLng(strFreeField)) = "cptFilterByClipboard" Then
      lngFreeField = CLng(strFreeField)
    Else 'remove it
      cptDeleteSetting "FilterByClipboard", "cboFreeField"
      lngFreeField = cptGetFreeField("Number")
    End If
  Else
    lngFreeField = cptGetFreeField("Number")
  End If

  If lngFreeField > 0 Then
    With myFilterByClipboard_frm.cboFreeField
      .Clear
      .AddItem
      .List(0, 0) = lngFreeField
      .List(0, 1) = FieldConstantToFieldName(lngFreeField)
      .Value = lngFreeField
      .Locked = True
    End With
  ElseIf lngFreeField = 0 Then
    strMsg = "Since there are no custom task number fields available*, filtered tasks will not appear in the same order as pasted."
    strMsg = strMsg & vbCrLf & vbCrLf
    strMsg = strMsg & "* 'available' means:" & vbCrLf
    strMsg = strMsg & "> no custom field name" & vbCrLf
    strMsg = strMsg & "> no formula" & vbCrLf
    strMsg = strMsg & "> no pick list" & vbCrLf
    strMsg = strMsg & "> no data on any task"
    If MsgBox(strMsg, vbInformation + vbOKCancel, "No Room at the Inn") = vbCancel Then GoTo exit_here
    With myFilterByClipboard_frm.cboFreeField
      .Clear
      .AddItem 0
      .List(.ListCount - 1, 1) = "Not Available"
      .Value = 0
      .Locked = True
    End With
  ElseIf lngFreeField = -1 Then 'user hit cancel
    GoTo exit_here
  End If
  
  myFilterByClipboard_frm.Show False
  
exit_here:
  On Error Resume Next
  Set myFilterByClipboard_frm = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_bas", "cptShowFilterByClipboard_frm", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptClipboardJump(ByRef myFilterByClipboard_frm As cptFilterByClipboard_frm)
  'objects
  'strings
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vList As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(myFilterByClipboard_frm.txtFilter.Text) = 0 Then Exit Sub
  vList = Split(myFilterByClipboard_frm.txtFilter.Text, ",")
  If UBound(vList) > 0 Then
    If myFilterByClipboard_frm.txtFilter.SelStart = Len(myFilterByClipboard_frm.txtFilter.Text) Then Exit Sub
    lngUID = vList(Len(Left(myFilterByClipboard_frm.txtFilter.Text, IIf(myFilterByClipboard_frm.txtFilter.SelStart = 0, 1, myFilterByClipboard_frm.txtFilter.SelStart))) - Len(Replace(Left(myFilterByClipboard_frm.txtFilter.Text, IIf(myFilterByClipboard_frm.txtFilter.SelStart = 0, 1, myFilterByClipboard_frm.txtFilter.SelStart)), ",", "")))
  Else
    lngUID = vList(0)
  End If
  If myFilterByClipboard_frm.lboFilter.ListCount > 0 Then myFilterByClipboard_frm.lboFilter.Value = lngUID

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_bas", "cptCliipboardJump", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateClipboard(ByRef myFilterByClipboard_frm As cptFilterByClipboard_frm)
  Dim oTask As MSProject.Task
  Dim oAssignment As MSProject.Assignment
  'strings
  Dim strMatchingTable As String
  Dim strFilter As String
  'longs
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngFreeField As Long
  Dim lngItems As Long
  Dim lngItem As Long
  Dim lngUID As Long
  Dim lngFactor As Long
  Dim lngAssignmentUID As Long
  'integers
  'doubles
  'booleans
  Dim blnAssignments As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vUID As Variant
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSpeed True
  
  myFilterByClipboard_frm.lboFilter.Clear
  strFilter = myFilterByClipboard_frm.txtFilter.Text
  ActiveWindow.TopPane.Activate
  FilterClear
  ScreenUpdating = False
  OptionsViewEx DisplaySummaryTasks:=True
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    Sort "ID", , , , , , False, True
    OutlineShowAllTasks
  End If
  SelectAll
  If Not IsNull(myFilterByClipboard_frm.cboFreeField.Value) Then
    lngFreeField = myFilterByClipboard_frm.cboFreeField
    SetField FieldConstantToFieldName(lngFreeField), 0
  End If
  If Len(strFilter) = 0 Then
    GoTo exit_here
  End If
  
  Application.StatusBar = ""
  
  vUID = Split(strFilter, ",")
  strFilter = ""
  blnAssignments = False
  If IsEmpty(vUID) Then GoTo exit_here
  For lngItem = 0 To UBound(vUID)
    If vUID(lngItem) = "" Then GoTo next_item

    If Not IsNumeric(vUID(lngItem)) Then GoTo next_item
    lngUID = vUID(lngItem)
    myFilterByClipboard_frm.lboFilter.AddItem lngUID
    
    'validate task (or assignment) exists
    On Error Resume Next
    Set oTask = Nothing
    Set oAssignment = Nothing
    If myFilterByClipboard_frm.optUID Then
      Set oTask = ActiveProject.Tasks.UniqueID(lngUID)
      If oTask Is Nothing Then
        Set oAssignment = ActiveProject.Tasks(1).Assignments.UniqueID(lngUID)
      End If
    Else
      Set oTask = ActiveProject.Tasks(lngUID)
      If oTask Is Nothing Then
        Set oAssignment = ActiveProject.Tasks(1).Assignments.UniqueID(lngUID)
      End If
    End If
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oAssignment Is Nothing Then blnAssignments = True
    If Not oTask Is Nothing Or Not oAssignment Is Nothing Then
      'add to autofilter
      strFilter = strFilter & lngUID & vbTab
      If Not oTask Is Nothing Then
        myFilterByClipboard_frm.lboFilter.List(myFilterByClipboard_frm.lboFilter.ListCount - 1, 1) = oTask.Name
        If lngFreeField > 0 Then oTask.SetField lngFreeField, CStr(lngItem + 1)
      ElseIf Not oAssignment Is Nothing Then
        myFilterByClipboard_frm.lboFilter.List(myFilterByClipboard_frm.lboFilter.ListCount - 1, 1) = oAssignment.Task.Name & " | " & oAssignment.ResourceName
        If lngFreeField > 0 Then SetMatchingField FieldConstantToFieldName(lngFreeField), CStr(lngItem + 1), "Unique ID", lngUID ' oTask.SetField lngFreeField, CStr(lngItem + 1)
      End If
    Else
      myFilterByClipboard_frm.lboFilter.List(lngItem, 1) = "< not found >"
    End If
next_item:
    Application.StatusBar = "Applying filter...(" & Format(lngItem / IIf(UBound(vUID) = 0, 1, UBound(vUID)), "0%") & ")"
    DoEvents
  Next lngItem
  
  If Not myFilterByClipboard_frm.tglEdit Then
    myFilterByClipboard_frm.lboFilter.Visible = True
    myFilterByClipboard_frm.txtFilter.Visible = False
  End If
  
  If blnAssignments Then
    myFilterByClipboard_frm.lboHeader.List(0, 1) = "Task Name | Resource Name"
    If MsgBox("Filter criteria includes Assignments!" & vbCrLf & vbCrLf & "Switch to Task Usage View?", vbQuestion + vbYesNo, "Switch View?") = vbYes Then
      strMatchingTable = ActiveProject.CurrentTable
      ActiveWindow.TopPane.Activate
      ScreenUpdating = False
      ViewApplyEx "Task Usage"
      FilterClear
      GroupClear
      OptionsViewEx DisplaySummaryTasks:=True
      On Error Resume Next
      If Not OutlineShowAllTasks Then
        Sort "ID", , , , , , False, True
        OutlineShowAllTasks
      End If
      If ActiveProject.CurrentTable <> strMatchingTable Then
        If MsgBox("Task Usage Table is '" & ActiveProject.CurrentTable & "'" & vbCrLf & vbCrLf & "Switch to Table '" & strMatchingTable & "' to match previous view?", vbQuestion + vbYesNo, "Switch Table?") = vbYes Then
          TableApply strMatchingTable
        End If
      End If
    End If
  Else
    myFilterByClipboard_frm.lboHeader.List(0, 1) = "Task Name"
  End If
  
  If Len(strFilter) > 0 And myFilterByClipboard_frm.chkFilter Then
    ActiveWindow.TopPane.Activate
    ScreenUpdating = False
    OptionsViewEx DisplaySummaryTasks:=True
    On Error Resume Next
    If Not OutlineShowAllTasks Then
      Sort "ID", , , , , , False, True
      OutlineShowAllTasks
    End If
    SelectAll
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    SelectBeginning
    strFilter = Left(strFilter, Len(strFilter) - 1)
    If myFilterByClipboard_frm.optUID Then
      myFilterByClipboard_frm.lboHeader.List(0, 0) = "UID"
      SetAutoFilter "Unique ID", FilterType:=pjAutoFilterIn, Criteria1:=strFilter
    ElseIf myFilterByClipboard_frm.optID Then
      myFilterByClipboard_frm.lboHeader.List(0, 0) = "ID"
      SetAutoFilter "Unique ID", FilterType:=pjAutoFilterIn, Criteria1:=strFilter
    End If
    OptionsViewEx ProjectSummary:=False, DisplayOutlineNumber:=False, DisplayNameIndent:=False, DisplaySummaryTasks:=False
    If lngFreeField > 0 Then Sort FieldConstantToFieldName(lngFreeField)
  End If
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Set oTask = Nothing
  Set oAssignment = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_bas", "cptUpdateClipboard", Err, Erl)
  Resume exit_here
End Sub

Function cptGuessDelimiter(ByRef vData As Variant, strRegEx As String) As Long
  'objects
  Dim dScores As Scripting.Dictionary
  Dim RE As Object
  Dim REMatches As Object
  'strings
  'longs
  Dim lngMax As Long
  Dim lngMatch As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vRecords As Variant
  Dim REMatch As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Set RE = CreateObject("vbscript.regexp")
  With RE
    .MultiLine = True
    .Global = True
    .IgnoreCase = True
    .Pattern = strRegEx
  End With
  
  Set dScores = CreateObject("Scripting.Dictionary")
  
  'check all "^([^\t\,\;]*[\t\,\;])"
  RE.Pattern = "^([^\t\,\;]*[\t\,\;])"
  For lngItem = 0 To UBound(vData)
    Set REMatches = RE.Execute(CStr(vData(lngItem)))
    For Each REMatch In REMatches
      lngMatch = Asc(Right(REMatch, 1))
      If dScores.Exists(lngMatch) Then
        'add a point
        dScores.Item(lngMatch) = dScores.Item(lngMatch) + 1
        If dScores.Item(lngMatch) > lngMax Then lngMax = dScores.Item(lngMatch)
      Else
        dScores.Add lngMatch, 1
      End If
    Next
  Next lngItem
  
  'check only valid "^([0-9]{1,}[\t\,\;])"
  RE.Pattern = "^([0-9]{1,}[\t\,\;])+"
  For lngItem = 0 To UBound(vData)
    On Error GoTo skip_it
    Set REMatches = RE.Execute(CStr(vData(lngItem)))
    For Each REMatch In REMatches
      lngMatch = Asc(Right(REMatch, 1))
      If dScores.Exists(lngMatch) Then
        'add a point
        dScores.Item(lngMatch) = dScores.Item(lngMatch) + 1
        If dScores.Item(lngMatch) > lngMax Then lngMax = dScores.Item(lngMatch)
      Else
        dScores.Add lngMatch, 1
      End If
    Next
skip_it:
  Next lngItem
  Err.Clear
  
  On Error Resume Next
  'which delimiter got the most points?
  'todo: this doesn't work if there is a tie
  For lngItem = 0 To dScores.Count - 1
    If dScores.Items(lngItem) = lngMax Then
      lngMatch = dScores.Keys(lngItem)
      Exit For
    End If
  Next lngItem
  If Err.Number > 0 Then
    cptGuessDelimiter = 0
  Else
    cptGuessDelimiter = lngMatch
  End If

exit_here:
  On Error Resume Next
  Set dScores = Nothing
  Set RE = Nothing
  Set REMatches = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptFilterByClipboard_bas", "cptGuessDelimiter", Err, Erl)
  If Err.Number = 5 Then
    cptGuessDelimiter = 0
    Err.Clear
  End If
  Resume exit_here
End Function

Function cptGetFreeField(strDataType As String, Optional lngType As Long) As Long
  'objects
  Dim dTypes As Object 'Scripting.Dictionary
  Dim rstFree As Object 'ADODB.Recordset
  Dim oTask As MSProject.Task
  Dim oAssignment As MSProject.Assignment
  'strings
  Dim strFreeField As String
  Dim strNum As String
  'longs
  Dim lngResponse As Long
  Dim lngFreeField As Long
  Dim lngField As Long
  Dim lngItems As Long
  Dim lngItem As Long
  Dim lngAssignmentUID As Long
  Dim lngFactor As Long
  'integers
  'doubles
  'booleans
  Dim blnFree As Boolean
  Dim blnMaster As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  lngFreeField = cptCustomFieldExists("cptFilterByClipboard")
  If lngFreeField > 0 Then
    cptSaveSetting "FilterByClipboard", "cboFreeField", lngFreeField
    cptGetFreeField = lngFreeField
    GoTo exit_here
  End If

  Calculation = pjManual
  
  'field type
  If lngType = 0 Then lngType = pjTask      'for reference
  If lngType = 1 Then lngType = pjResource  'for reference
  If lngType = 2 Then lngType = pjProject   'for refernce
  If lngType > 2 Then
    Err.Raise 9999, Description:="Invalid lngType: must be <=2"
  End If
  'data type
  'todo: Outline Code not acceptable?
  If InStr("Cost|Date|Duration|Finish|Flag|Number|OutlineCode|Outline Code|Start|Text", strDataType) = 0 Then
    Err.Raise 9999, Description:="Invalid strDataType: must be 'Cost' or 'Date' or 'Number' or 'Text' etc."
  End If
  
  'hash of local custom field counts
  Set dTypes = CreateObject("Scripting.Dictionary")
  dTypes.Add "Flag", 20
  dTypes.Add "Number", 20
  dTypes.Add "Text", 30
  If dTypes.Exists(strDataType) Then lngItems = dTypes(strDataType) Else lngItems = 10
  
  'prep to capture free fields
  Set rstFree = CreateObject("ADODB.Recordset")
  rstFree.Fields.Append "FieldConstant", adBigInt
  rstFree.Fields.Append "Available", adBoolean
  rstFree.Open
  
  'start with custom fields witout custom field names, examine last to first
  For lngItem = lngItems To 1 Step -1
    lngField = FieldNameToFieldConstant(strDataType & lngItem, lngType)
    If CustomFieldGetName(lngField) = "" Then
      'then ensure no formula
      If CustomFieldGetFormula(lngField) <> "" Then GoTo next_field 'skip it
      'then ensure no pick list (brute force)
      If strDataType <> "Flag" Then
        strNum = ActiveProject.Tasks(1).GetField(lngField)
        On Error Resume Next
        If InStr("Date|Start|Finish", strDataType) > 0 Then
          ActiveProject.Tasks(1).SetField lngField, #1/1/1984# 'what are the odds
        Else
          ActiveProject.Tasks(1).SetField lngField, 3.14285714285714 'what are the odds?
        End If
        If Err.Number > 0 Then
          Err.Clear
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo next_field
          GoTo next_field 'skip it
        Else
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo next_field
          ActiveProject.Tasks(1).SetField lngField, strNum
        End If
      End If
      'if we made it this far, then it's a candidate
      rstFree.AddNew Array(0, 1), Array(lngField, True)
    End If
next_field:
  Next lngItem
  'if all custom fields are named, then none available
  If rstFree.RecordCount = 0 Then
    lngFreeField = 0
    GoTo return_value
  End If
  
  'next ensure there is no data in that field on the tasks
  'note: use of ActiveProject.Tasks ensures all subprojects included
  'todo: does not catch when assignment field has a value
  'todo: would have to select each one to 'get' its data with CheckField
  If cptTableExists("cpt-temp-table") Then ActiveProject.TaskTables("cpt-temp-table").Delete
  TableEdit "cpt-temp-table", True, True, True, , "Unique ID"
  If cptViewExists("cpt-temp-view") Then ActiveProject.Views("cpt-temp-view").Delete
  ViewEditSingle "cpt-temp-view", True, , pjTaskUsage, False, False, "cpt-temp-table", "All Tasks", "No Group"
  Application.WindowNewWindow ActiveProject, "cpt-temp-view"
  FilterClear
  GroupClear
  SelectAll
  Sort "ID", , , , , , , True
  OptionsViewEx DisplaySummaryTasks:=True 'won't work without Sort first
  OutlineShowAllTasks 'this won't work unless summary tasks are showing
  rstFree.MoveFirst
  Do While Not rstFree.EOF
    If rstFree(1) Then
      Select Case strDataType
        Case "Cost"
          blnFree = Not Find(FieldConstantToFieldName(rstFree(0)), "does not equal", 0)
        Case "Date"
          blnFree = Not Find(FieldConstantToFieldName(rstFree(0)), "does not equal", "NA")
        Case "Duration"
          blnFree = Not Find(FieldConstantToFieldName(rstFree(0)), "does not equal", 0)
        Case "Finish"
          blnFree = Not Find(FieldConstantToFieldName(rstFree(0)), "does not equal", "NA")
        Case "Flag"
          blnFree = Not Find(FieldConstantToFieldName(rstFree(0)), "does not equal", False)
        Case "Number"
          blnFree = Not Find(FieldConstantToFieldName(rstFree(0)), "does not equal", 0)
        Case "Outline Code"
          blnFree = Not Find(FieldConstantToFieldName(rstFree(0)), "does not equal", "")
        Case "Start"
          blnFree = Not Find(FieldConstantToFieldName(rstFree(0)), "does not equal", "NA")
        Case "Text"
          blnFree = Not Find(FieldConstantToFieldName(rstFree(0)), "does not equal", "")
      End Select
      If blnFree Then
        rstFree(1) = blnFree
        lngResponse = MsgBox("Looks like " & FieldConstantToFieldName(rstFree(0)) & " isn't in use." & vbCrLf & vbCrLf & "OK to temporarily borrow it for this?", vbQuestion + vbYesNoCancel, "Wanted: Custom " & StrConv(strDataType, vbProperCase) & " Field")
        If lngResponse = vbYes Then
          lngFreeField = rstFree(0)
          Exit Do
        ElseIf lngResponse = vbCancel Then
          lngFreeField = -1
          Exit Do
        Else
          lngFreeField = 0
        End If
      End If
    End If
    rstFree.MoveNext
  Loop
  Application.ActiveWindow.Close
  ActiveProject.Views("cpt-temp-view").Delete
  ActiveProject.TaskTables("cpt-temp-table").Delete
  
return_value:
  If lngFreeField <> 0 Then
    cptGetFreeField = lngFreeField
  Else
    cptGetFreeField = 0
  End If

exit_here:
  On Error Resume Next
  Set dTypes = Nothing
  cptSpeed False
  If rstFree.State Then rstFree.Close
  Set rstFree = Nothing
  Set oTask = Nothing
  Set oAssignment = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptFilterByClipboard_bas", "cptGetFreeField", Err)
  Resume exit_here
End Function

Sub cptClearFreeField(ByRef myFilterByClipboard_frm As cptFilterByClipboard_frm, Optional blnPromptToSave As Boolean = False)
  'objects
  Dim oTask As MSProject.Task
  Dim oAssignment As MSProject.Assignment
  'strings
  Dim strMsg As String
  'longs
  Dim lngFreeField As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  Dim lngFactor As Long
  Dim lngAssignmentUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  Calculation = pjManual
  ScreenUpdating = False
  If IsNull(myFilterByClipboard_frm.cboFreeField) Then GoTo exit_here
  If myFilterByClipboard_frm.cboFreeField = "" Then GoTo exit_here
  lngFreeField = myFilterByClipboard_frm.cboFreeField.Value
  If lngFreeField > 0 And blnPromptToSave Then
    strMsg = "Save '" & FieldConstantToFieldName(lngFreeField) & "' for next time?" & vbCrLf & vbCrLf
    If ActiveProject.Subprojects.Count > 0 Then
      strMsg = strMsg & "Local Custom Field '" & FieldConstantToFieldName(lngFreeField) & "' will be renamed 'cptFilterByClipboard' in '" & ActiveProject.Name & "' only."
    Else
      strMsg = strMsg & "Local Custom Field '" & FieldConstantToFieldName(lngFreeField) & "' will be renamed 'cptFilterByClipboard'."
    End If
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Save Local Custom Number Field?") = vbYes Then
      CustomFieldRename lngFreeField, "cptFilterByClipboard"
      cptSaveSetting "FilterByClipboard", "cboFreeField", lngFreeField
    Else
      cptDeleteSetting "FilterByClipboard", "cboFreeField"
    End If
    SelectAll
    SetField FieldConstantToFieldName(lngFreeField), 0
    SelectBeginning
  End If
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Calculation = pjAutomatic
  ScreenUpdating = True

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_bas", "cptClearFreeField", Err, Erl)
  Resume exit_here
  
End Sub
