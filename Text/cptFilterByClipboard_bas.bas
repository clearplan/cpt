Attribute VB_Name = "cptFilterByClipboard_bas"
'<cpt_version>v1.2.3</cpt_version>
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
      lngFreeField = cptGetFreeField(myFilterByClipboard_frm, "Number")
    End If
  Else
    lngFreeField = cptGetFreeField(myFilterByClipboard_frm, "Number")
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
  'strings
  Dim strFilter As String
  'longs
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngFreeField As Long
  Dim lngItems As Long
  Dim lngItem As Long
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  Dim blnMaster As Boolean
  'variants
  Dim vUID As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnMaster = ActiveProject.Subprojects.Count > 0
  
  cptSpeed True
  
  myFilterByClipboard_frm.lboFilter.Clear
  strFilter = myFilterByClipboard_frm.txtFilter.Text
  If Len(strFilter) = 0 Then
    ActiveWindow.TopPane.Activate
    FilterClear
    GoTo exit_here
  End If
  
  lngTasks = ActiveProject.Tasks.Count
  If Not IsNull(myFilterByClipboard_frm.cboFreeField.Value) Then
    lngFreeField = myFilterByClipboard_frm.cboFreeField
    For Each oTask In ActiveProject.Tasks
      If oTask Is Nothing Then GoTo next_task
      If oTask.ExternalTask Then GoTo next_task
      If lngFreeField > 0 Then
        If CLng(oTask.GetField(lngFreeField)) > 0 Then
          oTask.SetField lngFreeField, 0
        End If
      End If
next_task:
      lngTask = lngTask + 1
      Application.StatusBar = "Resetting number field...(" & Format(lngTask / IIf(lngTasks = 0, 1, lngTasks), "0%") & ")"
      DoEvents
    Next oTask
  Else
    lngFreeField = 0
  End If
  Application.StatusBar = ""
  
  vUID = Split(strFilter, ",")
  strFilter = ""
  If IsEmpty(vUID) Then GoTo exit_here
  For lngItem = 0 To UBound(vUID)
    If vUID(lngItem) = "" Then GoTo next_item

    If Not IsNumeric(vUID(lngItem)) Then GoTo next_item
    lngUID = vUID(lngItem)
    myFilterByClipboard_frm.lboFilter.AddItem lngUID
    
    'validate task exists
    On Error Resume Next
    If myFilterByClipboard_frm.optUID Then
      Set oTask = ActiveProject.Tasks.UniqueID(lngUID)
    Else
      Set oTask = ActiveProject.Tasks(lngUID)
    End If
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oTask Is Nothing Then
      'add to autofilter
      strFilter = strFilter & lngUID & vbTab
      myFilterByClipboard_frm.lboFilter.List(myFilterByClipboard_frm.lboFilter.ListCount - 1, 1) = oTask.Name
      If lngFreeField > 0 Then oTask.SetField lngFreeField, CStr(lngItem + 1)
      Set oTask = Nothing
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
  
  If Len(strFilter) > 0 And myFilterByClipboard_frm.chkFilter Then
    ActiveWindow.TopPane.Activate
    ScreenUpdating = False
    OptionsViewEx DisplaySummaryTasks:=True
    SelectAll
    On Error Resume Next
    If Not OutlineShowAllTasks Then
      Sort "ID", , , , , , False, True
      OutlineShowAllTasks
    End If
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
        dScores.item(lngMatch) = dScores.item(lngMatch) + 1
        If dScores.item(lngMatch) > lngMax Then lngMax = dScores.item(lngMatch)
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
        dScores.item(lngMatch) = dScores.item(lngMatch) + 1
        If dScores.item(lngMatch) > lngMax Then lngMax = dScores.item(lngMatch)
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
    If dScores.items(lngItem) = lngMax Then
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

Function cptGetFreeField(ByRef myFilterByClipboard_frm As cptFilterByClipboard_frm, strDataType As String, Optional lngType As Long) As Long
  'objects
  Dim dTypes As Object 'Scripting.Dictionary
  Dim rstFree As Object 'ADODB.Recordset
  Dim oTask As MSProject.Task
  'strings
  Dim strFreeField As String
  Dim strNum As String
  'longs
  Dim lngResponse As Long
  Dim lngFreeField As Long
  Dim lngField As Long
  Dim lngItems As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnFree As Boolean
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  lngFreeField = cptCustomFieldExists("cptFilterByClipboard")
  If lngFreeField > 0 Then
    cptSaveSetting "FilterByClipboard", "cboFreeField", lngFreeField
    cptGetFreeField = lngFreeField
    GoTo exit_here
  End If

  Calculation = pjManual
  
  'field type
  If lngType = 0 Then lngType = pjTask
  
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
      If CustomFieldGetFormula(lngField) <> "" Then GoTo next_field
      'then ensure no pick list (brute force)
      strNum = ActiveProject.Tasks(1).GetField(lngField)
      On Error Resume Next
      ActiveProject.Tasks(1).SetField lngField, 3.14285714285714 'what are the odds?
      If Err.Number > 0 Then
        Err.Clear
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        GoTo next_field
      Else
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        ActiveProject.Tasks(1).SetField lngField, strNum
      End If
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
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    rstFree.MoveFirst
    Do While Not rstFree.EOF
      blnFree = True
      If Val(oTask.GetField(rstFree(0))) > 0 Then
        blnFree = False
        rstFree.Update Array(1), Array(blnFree)
        Exit For
      End If
      rstFree.MoveNext
    Loop
next_task:
  Next oTask

  rstFree.MoveFirst
  Do While Not rstFree.EOF
    If rstFree(1) = True Then
      lngResponse = MsgBox("Looks like " & FieldConstantToFieldName(rstFree(0)) & " isn't in use." & vbCrLf & vbCrLf & "OK to temporarily borrow it for this?", vbQuestion + vbYesNoCancel, "Wanted: Custom Number Field")
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
    rstFree.MoveNext
  Loop
  rstFree.Close
  
return_value:
  If lngFreeField <> 0 Then
    cptGetFreeField = lngFreeField
  Else
    cptGetFreeField = 0
  End If

exit_here:
  On Error Resume Next
  Set dTypes = Nothing
  Calculation = pjAutomatic
  If rstFree.State Then rstFree.Close
  Set rstFree = Nothing
  Set oTask = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptFilterByClipboard", "cptGetFreeField", Err)
  Resume exit_here
End Function

Sub cptClearFreeField(ByRef myFilterByClipboard_frm As cptFilterByClipboard_frm, Optional blnPromptToSave As Boolean = False)
  'objects
  Dim oTask As MSProject.Task
  'strings
  Dim strMsg As String
  'longs
  Dim lngFreeField As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  'integers
  'doubles
  'booleans
  Dim blnMaster As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  blnMaster = ActiveProject.Subprojects.Count > 0
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
    lngTasks = ActiveProject.Tasks.Count
    For Each oTask In ActiveProject.Tasks
      If Not oTask Is Nothing Then
        If CLng(oTask.GetField(lngFreeField)) > 0 Then
          oTask.SetField lngFreeField, 0
        End If
      End If
      lngTask = lngTask + 1
      Application.StatusBar = "Clearing " & FieldConstantToFieldName(lngFreeField) & "...(" & Format(lngTask / lngTasks, "0%") & ")"
    Next oTask
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
