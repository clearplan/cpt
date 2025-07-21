Attribute VB_Name = "cptText_bas"
'<cpt_version>v1.5.5</cpt_version>
Option Explicit

Sub cptBulkAppend()
  'objects
  Dim oTasks As MSProject.Tasks, oTask As MSProject.Task
  'strings
  Dim strAppend As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strAppend = InputBox("Append what text to selected tasks?", "Append Text")
  
  If Len(strAppend) = 0 Then Exit Sub
  
  Application.OpenUndoTransaction "Bulk Append"
  
  For Each oTask In oTasks
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask Is Nothing Then
      oTask.Name = Trim(oTask.Name) & " " & Trim(strAppend)
    End If
next_task:
  Next oTask
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set oTask = Nothing
  Set oTasks = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptBulkAppend", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptBulkPrepend()
  'objects
  Dim oTasks As MSProject.Tasks, oTask As MSProject.Task
  'strings
  Dim strPrepend As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strPrepend = InputBox("Prepend what text to selected tasks?", "Prepend Text")
  
  If Len(strPrepend) = 0 Then Exit Sub
  
  Application.OpenUndoTransaction "Bulk Prepend"
  
  For Each oTask In oTasks
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask Is Nothing Then
       oTask.Name = Trim(strPrepend) & " " & Trim(oTask.Name)
    End If
next_task:
  Next oTask
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set oTask = Nothing
  Set oTasks = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptBulkPrepend", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptEnumerate()
  'objects
  Dim oTasks As MSProject.Tasks, oTask As MSProject.Task
  'strings
  'longs
  Dim lngDigits As Long
  Dim lngEnumerate As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vResponse As Variant
  'dates

  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  vResponse = InputBox("How many digits (number input only)?", "Format Enumeration", 3)
  If StrPtr(vResponse) = 0 Then
    'user hit cancel
    GoTo exit_here
  ElseIf vResponse = vbNullString Then
    'user entered null value
    GoTo exit_here
  End If
  lngDigits = CLng(vResponse)
  
  vResponse = InputBox("Start at what number (number input only)?", "Format Enumeration", 1)
  If StrPtr(vResponse) = 0 Then
    'user hit cancel
    GoTo exit_here
  ElseIf vResponse = vbNullString Then
    'user entered null value
    GoTo exit_here
  End If
  lngEnumerate = CLng(vResponse)
  
  cptSpeed True
  
  Application.OpenUndoTransaction "Enumeration"
  
  If oTasks.Count > 2 Then
    For Each oTask In oTasks
      If oTask.ExternalTask Then GoTo next_task
      If Not oTask Is Nothing Then
        oTask.Name = oTask.Name & " (" & Format(lngEnumerate, String(lngDigits, "0")) & ")"
        lngEnumerate = lngEnumerate + 1
      End If
next_task:
    Next oTask
  End If
  cptSpeed False

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set oTask = Nothing
  Set oTasks = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptEnumerate", Err, Erl)
  Resume exit_here

End Sub

Sub cptMyReplace()
  'fields affected: Marked, Task Name, Text Fields, Outline Code Fields
  'objects
  Dim rstReplaced As Object 'ADODB.Recordset
  Dim oTasks As MSProject.Tasks, oTask As MSProject.Task
  'strings
  Dim strMsg As String
  'longs
  Dim lngItem As Long
  Dim lngFound As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vField As Variant, vFind As Variant, vReplace As Variant
  'dates

  On Error Resume Next
  cptSpeed True
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'get string to find
  vFind = InputBox("Find what text:", "Replace")
  If StrPtr(vFind) = 0 Then GoTo exit_here 'user hit cancel
  vFind = Trim(vFind)
  
  'get string to replace it with
  vReplace = InputBox("Replace '" & CStr(vFind) & "' with what text:", "Replace")
  If StrPtr(vReplace) = 0 Then GoTo exit_here 'user hit cancel
  vReplace = Trim(vReplace)
  
  Application.OpenUndoTransaction "MyReplace"

  Set rstReplaced = CreateObject("ADODB.Recordset")
  rstReplaced.Fields.Append "UID", adBigInt
  rstReplaced.Open

  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    For Each vField In ActiveSelection.FieldIDList
      'limit to text fields
      If Len(cptRegEx(FieldConstantToFieldName(vField), "Text|Name")) > 0 Then
        If InStr(oTask.GetField(vField), CStr(vFind)) > 0 Then
          oTask.SetField vField, Replace(oTask.GetField(vField), CStr(vFind), CStr(vReplace))
          rstReplaced.AddNew Array("UID"), Array(oTask.UniqueID)
          rstReplaced.Update
          lngFound = lngFound + 1
        End If
      End If
    Next vField
next_task:
  Next oTask

  If lngFound = 0 Then
    MsgBox "No instances of '" & CStr(vFind) & "' found in selected cells.", vbExclamation + vbOKOnly, "MyReplace"
  Else
    rstReplaced.MoveFirst
    FilterEdit "cptMyReplace", True, True, True, False, , "Unique ID", , "equals", rstReplaced(0), "Or", True
    Do While Not rstReplaced.EOF
      FilterEdit "cptMyReplace", TaskFilter:=True, FieldName:="", NewFieldName:="Unique ID", test:="equals", Value:=rstReplaced(0), operation:="Or", ShowInMenu:=True
      rstReplaced.MoveNext
    Loop
    FilterApply "cptMyReplace", True
    rstReplaced.MoveFirst
    Application.Find "Unique ID", "equals", rstReplaced(0)
    cptSpeed False
    strMsg = "Replaced " & Format(lngFound, "#,##0") & " instance" & IIf(lngFound = 1, "", "s") & " of '" & CStr(vFind) & "' with '" & CStr(vReplace) & "'" & vbCrLf & vbCrLf
    strMsg = strMsg & "Keep highlighted?"
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Replace") = vbNo Then
      cptSpeed True
      FilterApply "All Tasks", True
      Application.Find "Unique ID", "equals", rstReplaced(0)
      cptSpeed False
    End If
  End If
  
exit_here:
  On Error Resume Next
  If rstReplaced.State Then rstReplaced.Close
  Set rstReplaced = Nothing
  Application.CloseUndoTransaction
  cptSpeed False
  Set oTasks = Nothing
  Set oTask = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptMyReplace", Err, Erl)
  Resume exit_here

End Sub

Sub cptFindDuplicateTaskNames()
  'objects
  Dim oTask As MSProject.Task
  Dim oDict As Scripting.Dictionary
  Dim oSubproject As MSProject.SubProject
  Dim oShell As Object
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oRange As Excel.Range
  Dim oListObject As ListObject
  'strings
  Dim strRestoreFilter As String
  Dim strFilter As String
  Dim strMap As String
  Dim strFileName As String
  'longs
  Dim lngItem As Long
  Dim lngItems As Long
  Dim lngNameCol As Long
  Dim lngDuplicateNames As Long
  Dim lngInstances As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnMaster As Boolean
  Dim blnIgnoreExternal As Boolean
  Dim blnIgnoreSummaryTasks As Boolean
  'variants
  'dates

  If ActiveProject.Tasks.Count = 0 Then GoTo exit_here
  If ActiveProject.Subprojects.Count > 0 Then blnMaster = True
  If Not cptCheckReference("Excel") Then GoTo exit_here
  blnErrorTrapping = cptErrorTrapping 'evaluate once per routine
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'first, expand all subprojects
  ActiveWindow.TopPane.Activate
  GroupClear
  FilterClear
  Sort "ID", , , , , , False, True
  
  If blnMaster Then
    If MsgBox("Load and expand all subprojects?", vbQuestion + vbYesNo, "Please confirm:") = vbNo Then GoTo exit_here
    lngItems = ActiveProject.Subprojects.Count
    For Each oSubproject In ActiveProject.Subprojects
      lngItem = lngItem + 1
      Application.StatusBar = "Loading " & oSubproject.InsertedProjectSummary.Name & "..."
      EditGoTo oSubproject.InsertedProjectSummary.ID
      Application.OutlineShowSubTasks
      Application.StatusBar = "Loading " & oSubproject.InsertedProjectSummary.Name & "...(" & Format(lngItem / lngItems, "0%") & ")"
    Next oSubproject
    Application.StatusBar = "Loading subprojects...done."
    Set oSubproject = Nothing
  End If
  
  blnIgnoreSummaryTasks = MsgBox("Ignore Summary Tasks?", vbQuestion + vbYesNo, "Please Confrim") = vbYes
  'first, build the filter
  strFilter = "cptDuplicateTaskNames Filter"
  If Edition = pjEditionProfessional Then
    FilterEdit Name:=strFilter, TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Active", test:="equals", Value:="Yes"
  ElseIf Edition = pjEditionStandard Then
    FilterEdit Name:=strFilter, TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Name", test:="does not contain", Value:="absolute_gibberish"
  End If
  If blnIgnoreSummaryTasks Then
    FilterEdit Name:=strFilter, TaskFilter:=True, NewFieldName:="Summary", test:="equals", Value:="No", operation:="And", Parenthesis:=False
  End If
  strRestoreFilter = ActiveProject.CurrentFilter
  FilterApply strFilter
  DoEvents
  blnIgnoreExternal = True
  'only prompt if external tasks exist
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo analyze_next
    If Not oTask.Active Then GoTo analyze_next
    If blnIgnoreSummaryTasks And oTask.Summary Then GoTo analyze_next
    If oTask.ExternalTask Then
      blnIgnoreExternal = MsgBox("Ignore External Tasks?", vbQuestion + vbYesNo, "Please Confirm") = vbYes
      Exit For
    End If
analyze_next:
  Next oTask
  If blnIgnoreExternal Then
    FilterEdit strFilter, TaskFilter:=True, NewFieldName:="External Task", test:="equals", Value:="No", operation:="And", Parenthesis:=False
  End If
  FilterEdit strFilter, TaskFilter:=True, ShowSummaryTasks:=Not blnIgnoreSummaryTasks
  FilterApply strFilter
  DoEvents
  
  'then build the map
  strMap = "cptDuplicateTaskNames Map"
  MapEdit Name:=strMap, Create:=True, OverwriteExisting:=True, datacategory:=0, categoryenabled:=True, TableName:="cptDuplicateTaskNames", FieldName:="Unique ID", ExternalFieldName:="Unique_ID", ExportFilter:=strFilter, ImportMethod:=0, HeaderRow:=True, AssignmentData:=False, TextDelimiter:=Chr$(9), TextFileOrigin:=0, UseHtmlTemplate:=False, IncludeImage:=False
  If blnMaster Then
    MapEdit Name:=strMap, datacategory:=0, FieldName:="Project", ExternalFieldName:="Project"
  End If
  If Not blnIgnoreExternal Then
    MapEdit Name:=strMap, datacategory:=0, FieldName:="External Task", ExternalFieldName:="External Task"
  End If
  If Not blnIgnoreSummaryTasks Then
    MapEdit Name:=strMap, datacategory:=0, FieldName:="Summary", ExternalFieldName:="Summary"
  End If
  MapEdit Name:=strMap, datacategory:=0, FieldName:="Name", ExternalFieldName:="Name", ExportFilter:=strFilter, ImportMethod:=0, HeaderRow:=True, AssignmentData:=False, TextDelimiter:=Chr$(9), TextFileOrigin:=0, UseHtmlTemplate:=False, IncludeImage:=False
  
  'perform the analysis
  Set oDict = CreateObject("Scripting.Dictionary")
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If blnIgnoreExternal And oTask.ExternalTask Then GoTo next_task
    If blnIgnoreSummaryTasks And oTask.Summary Then GoTo next_task
    If Not oDict.Exists(oTask.Name) Then
      oDict.Add oTask.Name, 1
    Else
      lngDuplicateNames = lngDuplicateNames + 1
      oDict.Item(oTask.Name) = oDict(oTask.Name) + 1
    End If
next_task:
  Next oTask
  
  If lngDuplicateNames > 0 Then
    lngDuplicateNames = 0
    For lngItem = 0 To oDict.Count - 1
      If oDict.Items(lngItem) > 1 Then
        lngDuplicateNames = lngDuplicateNames + 1
        lngInstances = lngInstances + oDict.Items(lngItem)
      End If
    Next lngItem
    If MsgBox(Format(lngDuplicateNames, "#,##0") & " duplicate task names found." & vbCrLf & Format(lngInstances, "#,##") & " instances of duplication." & vbCrLf & vbCrLf & "Open in Excel?", vbExclamation + vbYesNo, "Duplicate Task Names Found") = vbYes Then
      Set oShell = CreateObject("WScript.Shell")
      strFileName = oShell.SpecialFolders("Desktop") & "\DuplicateTaskNames_" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".xlsx"
      Application.StatusBar = "Exporting..."
      Application.FileSaveAs Name:=strFileName, FormatID:="MSProject.ACE", Map:=strMap
      Application.StatusBar = "Exporting...done."
      Application.StatusBar = "Formatting..."
      On Error Resume Next
      Set oExcel = GetObject(, "Excel.Application")
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oExcel Is Nothing Then
        Set oExcel = CreateObject("Excel.Application")
      Else
        Application.ActivateMicrosoftApp pjMicrosoftExcel
      End If
      oExcel.Visible = True
      oExcel.WindowState = xlNormal
      Set oWorkbook = oExcel.Workbooks.Open(strFileName)
      Set oWorksheet = oWorkbook.Sheets(1)
      Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(-4121)), , xlYes)
      oExcel.ActiveWindow.Zoom = 85
      oListObject.Range.Columns.AutoFit
      oListObject.TableStyle = ""
      Set oRange = oWorksheet.Range("Table1[Name]")
      oRange.FormatConditions.AddUniqueValues
      oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
      oRange.FormatConditions(1).DupeUnique = xlDuplicate
      With oRange.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
      End With
      With oRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
      End With
      oRange.FormatConditions(1).StopIfTrue = False
      'filter for duplicates
      lngNameCol = oWorksheet.Rows(1).Find("Name", lookat:=xlWhole).Column
      oListObject.Range.AutoFilter Field:=lngNameCol, Criteria1:=RGB(255, 199, 206), Operator:=xlFilterCellColor
      oListObject.Range.Columns.AutoFit
      'add a count formula
      oWorksheet.Cells(1, oWorksheet.[A1].End(xlToRight).Column + 1).Value = "Count"
      oListObject.ListColumns("Count").DataBodyRange.FormulaR1C1 = "=COUNTIFS([Name],[@Name])"
      'sort by task name (to put duplicates together)
      oListObject.Sort.SortFields.Clear
      oListObject.Sort.SortFields.Add2 Key:=oWorksheet.Range("Table1[Count]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
      oListObject.Sort.SortFields.Add2 Key:=oWorksheet.Range("Table1[Name]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      With oListObject.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
      End With
      oWorksheet.[A1].Select
      oWorkbook.Activate
      Application.StatusBar = "Formatting...done."
      Application.StatusBar = "Complete."
    End If
  Else
    MsgBox "No duplicate task names found.", vbInformation + vbOKOnly, "Well Done"
  End If
  
  ActiveWindow.TopPane.Activate
  GroupClear
  FilterClear
  Sort "ID", , , , , , False, True
  OptionsViewEx DisplaySummaryTasks:=True
  FilterApply strRestoreFilter
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oDict = Nothing
  Set oTask = Nothing
  Set oSubproject = Nothing
  Set oShell = Nothing
  Set oWorkbook = Nothing
  Set oWorksheet = Nothing
  Set oRange = Nothing
  If Not oExcel Is Nothing Then oExcel.Visible = True
  Set oExcel = Nothing
  Set oListObject = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptFindDuplicateTaskNames", Err, Erl)
  Resume exit_here

End Sub

Sub cptTrimTaskNames()
  'objects
  Dim oTask As MSProject.Task
  'strings
  'longs
  Dim lngSubproject As Long
  Dim lngSubprojects As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  Dim lngBefore As Long
  Dim lngAfter As Long
  Dim lngCount As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True

  Application.OpenUndoTransaction "Trim Task Names"
  
  lngSubprojects = ActiveProject.Subprojects.Count
  If lngSubprojects > 0 Then
    lngTasks = ActiveProject.Tasks.Count
    For lngSubproject = 1 To lngSubprojects
      lngTasks = lngTasks + ActiveProject.Subprojects(lngSubproject).SourceProject.Tasks.Count
    Next lngSubproject
  Else
    lngTasks = ActiveProject.Tasks.Count
  End If
  lngTask = 0
  For Each oTask In ActiveProject.Tasks
    lngTask = lngTask + 1
    If Not oTask Is Nothing Then
      If oTask.ExternalTask Then GoTo next_task
      If ActiveProject.Subprojects.Count > 0 Then
        Application.StatusBar = "Trimming Task UID " & oTask.UniqueID & " (" & Format(lngTask / lngTasks, "0%") & ")"
      Else
        Application.StatusBar = "Trimming Task ID " & oTask.ID & " (" & Format(lngTask / lngTasks, "0%") & ")"
      End If
      DoEvents
      lngBefore = Len(oTask.Name)
      'replace multi-spaces with single space
      oTask.Name = Replace(oTask.Name, cptRegEx(oTask.Name, "\s{2,}"), " ")
      'trim leading and trailing spaces
      oTask.Name = Trim(oTask.Name)
      lngAfter = Len(oTask.Name)
      If lngBefore > lngAfter Then lngCount = lngCount + 1
    End If
next_task:
  Next oTask

  Application.StatusBar = Format(lngCount, "#,##0") & " task names trimmed."

  MsgBox Format(lngCount, "#,##0") & " task names trimmed.", vbInformation + vbOKOnly, "Trim Task Names"

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  cptSpeed False
  Application.StatusBar = ""
  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptTrimTaskNames", Err, Erl)
  Resume exit_here

End Sub

Sub cptShowText_frm()
  'objects
  Dim myText_frm As cptText_frm
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  'strings
  Dim strCustomFieldName As String
  'longs
  Dim lngItem As Long
  'integers
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not cptModuleExists("cptText_frm") Then GoTo exit_here
  
  Set myText_frm = New cptText_frm
  With myText_frm.cboScope
    .AddItem
    .List(0, 0) = FieldNameToFieldConstant("Name", pjTask)
    .List(0, 1) = "Task Name"
    'todo: others?
    For lngItem = 1 To 30
'      If Len(CustomFieldGetFormula(FieldNameToFieldConstant("Text" & lngItem))) = 0 Then
        .AddItem
        .List(.ListCount - 1, 0) = FieldNameToFieldConstant("Text" & lngItem)
        strCustomFieldName = CustomFieldGetName(FieldNameToFieldConstant("Text" & lngItem))
        If Len(strCustomFieldName) > 0 Then
          .List(.ListCount - 1, 1) = strCustomFieldName & " (Text" & lngItem & ")"
        Else
          .List(.ListCount - 1, 1) = "Text" & lngItem
        End If
'      End If
    Next lngItem
    .Value = FieldNameToFieldConstant("Name", pjTask)
  End With
  
  lngItem = 0
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oTasks Is Nothing Then
    myText_frm.lboOutput.Clear
    For Each oTask In oTasks
      myText_frm.lboOutput.AddItem
      myText_frm.lboOutput.List(lngItem, 0) = oTask.UniqueID
      myText_frm.lboOutput.List(lngItem, 1) = oTask.Name
      lngItem = lngItem + 1
    Next oTask
  End If
  myText_frm.Caption = "Advanced Text Tools (" & cptGetVersion("cptText_frm") & ")"
  Call cptStartEvents
  myText_frm.Show False 'VBA.FormShowConstants.vbModeless
  
exit_here:
  On Error Resume Next
  Set myText_frm = Nothing
  Set oTask = Nothing
  Set oTasks = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptShowText_frm", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptUpdatePreview(ByRef myText_frm As cptText_frm, Optional strPrepend As String, Optional strAppend As String, Optional strPrefix As String, Optional lngCharacters As Long, Optional lngStartAt As Long, _
                  Optional lngCountBy As Long, Optional strSuffix As String, Optional strReplaceWhat As String, Optional strReplaceWith As String)
  'objects
  Dim oTask As MSProject.Task
  'strings
  Dim strTaskName As String
  Dim strEnumerate As String
  'longs
  Dim lngScope As Long
  Dim lngItem As Long
  Dim lngEnumerate As Long
  'integers
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If myText_frm.Visible Then
    lngScope = myText_frm.cboScope.Value
  Else
    lngScope = FieldNameToFieldConstant("Name", pjTask)
  End If

  For lngItem = 0 To myText_frm.lboOutput.ListCount - 1
    If IsNull(myText_frm.lboOutput.List(lngItem, 0)) Then GoTo exit_here
    On Error Resume Next
    Set oTask = ActiveProject.Tasks.UniqueID(myText_frm.lboOutput.List(lngItem, 0))
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oTask Is Nothing Then
      If MsgBox("UID " & myText_frm.lboOutput.List(lngItem, 0) & " not found in " & UCase(ActiveProject.Name) & "! Proceed?", vbCritical + vbYesNo, "Task Not Found") = vbNo Then
        Err.Clear
        GoTo exit_here
      Else
        GoTo next_item
      End If
    End If
    
    'start with the task name
    strTaskName = oTask.GetField(lngScope) 'Name
    
    If Len(strPrepend) > 0 Then
      strTaskName = Trim(strPrepend) & " " & strTaskName
    ElseIf Len(myText_frm.txtPrepend.Value) > 0 Then
      strTaskName = Trim(myText_frm.txtPrepend.Value) & " " & strTaskName
    End If
    If Len(strAppend) > 0 Then
      strTaskName = Trim(strTaskName) & " " & Trim(strAppend)
    ElseIf Len(myText_frm.txtAppend.Value) > 0 Then
      strTaskName = Trim(strTaskName) & " " & Trim(myText_frm.txtAppend.Value)
    End If
    myText_frm.chkIsDirty = myText_frm.CheckDirty
    If myText_frm.chkIsDirty Then
      strEnumerate = IIf(Len(strPrefix) > 0, strPrefix, myText_frm.txtPrefix.Value)

      If lngStartAt = 0 Then
        If myText_frm.txtStartAt.Value = "" Then
          lngStartAt = 1
          'myText_frm.txtStartAt.Value = 1
        Else
          lngStartAt = CLng(myText_frm.txtStartAt.Value)
        End If
      End If

      If lngCountBy = 0 Then
        If myText_frm.txtCountBy.Value = "" Then
          lngCountBy = 1
          'myText_frm.txtCountBy.Value = 1
        Else
          lngCountBy = CLng(myText_frm.txtCountBy.Value)
        End If
      End If

      lngEnumerate = lngStartAt + (lngItem * lngCountBy)

      If lngCharacters = 0 Then
        If myText_frm.txtCharacters.Value = "" Then
          lngCharacters = 1
          'myText_frm.txtCharacters.Value = 1
        Else
          lngCharacters = CLng(myText_frm.txtCharacters.Value)
        End If
      End If

      strEnumerate = strEnumerate & Format(lngEnumerate, String(lngCharacters, "0"))
      strEnumerate = strEnumerate & IIf(Len(strSuffix) > 0, strSuffix, myText_frm.txtSuffix.Value)
      myText_frm.lboOutput.List(lngItem, 1) = strTaskName & " " & strEnumerate
    Else
      myText_frm.lboOutput.List(lngItem, 1) = strTaskName
    End If
    
    'replace
    '<issue27> added
    If Len(strReplaceWhat) = 0 Then strReplaceWhat = myText_frm.txtReplaceWhat.Value
    If Len(strReplaceWith) = 0 Then strReplaceWith = myText_frm.txtReplaceWith.Value
    If Len(strReplaceWhat) > 0 And Len(strReplaceWith) > 0 Then
      strTaskName = Replace(strTaskName, strReplaceWhat, strReplaceWith)
      myText_frm.lboOutput.List(lngItem, 1) = strTaskName & " " & strEnumerate '</issue27>
    End If
next_item:
  Next lngItem

exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptUpdatePreview", Err, Erl)
  Resume exit_here

End Sub

Sub cptResetRowHeight()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "ResetRowHeight"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  SetRowHeight 1, "all"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptResetRowHeight", Err, Erl)
  Resume exit_here
End Sub

Sub cptCheckAnnoyances()
  'objects
  Dim oDict As Scripting.Dictionary
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  'strings
  Dim strInactive As String
  Dim strInactiveList As String
  Dim strElapsed As String
  Dim strElapsedList As String
  Dim strFileName As String
  Dim strTimes As String
  Dim strTimesList As String
  Dim strDurations As String
  Dim strDurationsList As String
  Dim strFilter As String
  'longs
  Dim lngItem As Long
  Dim lngInactive As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngFile As Long
  Dim lngCount As Long
  'integers
  'doubles
  'booleans
  Dim blnProjectStandard As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnProjectStandard = Application.Edition = pjEditionStandard
  Set oDict = CreateObject("Scripting.Dictionary")
  If ActiveProject.Subprojects.Count > 0 Then
    ActiveWindow.TopPane.Activate
    FilterClear
    GroupClear
    Sort "ID", , , , , , False, True
    OptionsViewEx DisplaySummaryTasks:=True
    SelectAll
    OutlineShowAllTasks
    SelectAll
    lngTasks = ActiveSelection.Tasks.Count
  Else
    lngTasks = ActiveProject.Tasks.Count
  End If
  
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then
      If blnProjectStandard Then
        lngInactive = lngInactive + 1
        strInactiveList = strInactiveList & oTask.UniqueID & vbTab
        strInactive = strInactive & oTask.UniqueID & ",INACTIVE,"
        If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
        GoTo next_task
      Else
        GoTo next_task
      End If
    End If
    If IsDate(oTask.ActualFinish) Then GoTo next_task
    If oTask.Duration = 0 Then GoTo next_task 'todo: skip milestones or not?
    If TimeValue(oTask.Finish) <> "5:00:00 PM" Or TimeValue(oTask.Start) <> "8:00:00 AM" Then
      strTimesList = strTimesList & oTask.UniqueID & vbTab
      strTimes = strTimes & oTask.UniqueID & "," & TimeValue(oTask.Start) & "," & TimeValue(oTask.Finish) & vbCrLf
      If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
    End If
    If InStr(oTask.DurationText, ".") > 0 Then
      strDurationsList = strDurationsList & oTask.UniqueID & vbTab
      strDurations = strDurations & oTask.UniqueID & "," & oTask.DurationText & vbCrLf
      If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
    End If
    If Left(cptRegEx(oTask.DurationText, "[A-z]{1,}"), 1) = "e" Then
      strElapsedList = strElapsedList & oTask.UniqueID & vbTab
      strElapsed = strElapsed & oTask.UniqueID & "," & oTask.DurationText & vbCrLf
      If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
    End If
    'todo: Assignments without work
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Checking...(" & Format(lngTask / lngTasks, "0%") & ")"
  Next oTask
  
  strFilter = strTimesList & strDurationsList & strElapsedList & strInactiveList
  If Len(strFilter) = 0 Then
    MsgBox "No annoyances!", vbInformation + vbOKOnly, "Well Done"
  Else
    strFilter = Left(strFilter, Len(strFilter) - 1) 'hack off last tab
    ActiveWindow.TopPane.Activate
    GroupClear
    FilterClear
    OptionsViewEx DisplaySummaryTasks:=True
    OutlineShowAllTasks
    SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strFilter
    SelectBeginning
    SelectAll
    On Error Resume Next
    Set oTasks = ActiveSelection.Tasks
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oTasks.Count = 0 Then
      MsgBox "No annoyances.", vbInformation + vbOKOnly, "Well Done"
    Else
      For Each oTask In oTasks
        If Not oTask.Summary Then lngCount = lngCount + 1
      Next oTask
    End If
    MsgBox Format(lngCount, "#,##0") & " annoyance" & IIf(lngCount > 1, "s", "") & " found!", vbInformation + vbOKOnly, "Annoyances"
    If MsgBox("View report?", vbQuestion + vbYesNo, "Annoyances") = vbYes Then
      lngFile = FreeFile
      strFileName = Environ("tmp") & "\annoyances.txt"
      Open strFileName For Output As lngFile
      If Len(strTimes) > 0 Then
        Print #lngFile, "===== ODD TIMES ARE ANNOYING ====="
        Print #lngFile, "UID,START,FINISH"
        Print #lngFile, strTimes
        Print #lngFile, "UID LIST: " & Replace(strTimesList, vbTab, ",")
        Print #lngFile, vbCrLf
      End If
      If Len(strDurations) > 0 Then
        Print #lngFile, "===== FRACTIONAL DURATIONS ARE ANNOYING ====="
        Print #lngFile, "UID,DURATION"
        Print #lngFile, strDurations
        Print #lngFile, "UID LIST: " & Replace(strDurationsList, vbTab, ",")
        Print #lngFile, vbCrLf
      End If
      If Len(strElapsed) > 0 Then
        Print #lngFile, "===== ELAPSED DURATIONS ARE ANNOYING ====="
        Print #lngFile, "UID,DURATION"
        Print #lngFile, strElapsed
        Print #lngFile, "UID LIST: " & Replace(strElapsedList, vbTab, ",")
        Print #lngFile, vbCrLf
      End If
      If blnProjectStandard And Len(strInactive) > 0 Then
        Print #lngFile, "===== INACTIVE TASKS IN PROJECT STANDARD ARE ANNOYING ====="
        Print #lngFile, "===== AND WILL PROBABLY LEAD TO ANOMALOUS CALCULATIONS ===="
        Print #lngFile, "UID,ACTIVE"
        Print #lngFile, strInactive
        Print #lngFile, "UID LIST: " & Replace(strInactiveList, vbTab, ",")
        Print #lngFile, vbCrLf
      End If
      strFilter = ""
      For lngItem = 0 To oDict.Count - 1
        strFilter = strFilter & oDict.Keys(lngItem) & ","
      Next lngItem
      Print #lngFile, "COMBINED UID LIST: " & Replace(strFilter, vbTab, ",")
      Close #lngFile
      Shell "notepad.exe """ & strFileName & """", vbNormalFocus
    End If
  End If
  
exit_here:
  On Error Resume Next
  Set oDict = Nothing
  Set oTasks = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptCheckAnnoyances", Err, Erl)
  Resume exit_here
End Sub

