Attribute VB_Name = "cptDynamicFilter_bas"
'<cpt_version>v1.6.4</cpt_version>
Option Explicit
Private pCachedRegexes As Scripting.Dictionary

Sub cptShowDynamicFilter_frm()
  'objects
  Dim myDynamicFilter_frm As cptDynamicFilter_frm
  'strings
  Dim strCustomFields As String
  Dim strCustomFieldName As String
  'longs
  Dim lngFieldConstant As Long
  Dim lngItem As Long
  'integers
  'booleans
  'variants
  Dim vArray As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Dynamic Filter"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  Set myDynamicFilter_frm = New cptDynamicFilter_frm
  With myDynamicFilter_frm
    .Caption = "Dynamic Filter (" & cptGetVersion("cptDynamicFilter_frm") & ")"
    .txtFilter = ""
    With .cboField
      .Clear
      'get local custom text fields
      For lngItem = 1 To 30
        lngFieldConstant = FieldNameToFieldConstant("Text" & lngItem)
        strCustomFieldName = CustomFieldGetName(lngFieldConstant)
        If Len(strCustomFieldName) > 0 Then
          strCustomFields = strCustomFields & strCustomFieldName & ","
        End If
      Next lngItem
      
      'get enterprise custom fields
      For lngItem = 188776000 To 188778000 '2000 should do it for now
        If Application.FieldConstantToFieldName(lngItem) <> "<Unavailable>" Then
          strCustomFields = strCustomFields & FieldConstantToFieldName(lngItem) & ","
        End If
      Next lngItem
      
      'remove terminal comma, reducing array size by one
      If Len(strCustomFields) > 0 Then
        strCustomFields = Left(strCustomFields, Len(strCustomFields) - 1)
        vArray = Split(strCustomFields, ",")
        cptQuickSort vArray, 0, UBound(vArray)
        'join vArray into string, prepend 'Task Name', split into array
        .List = Split("Task Name," & Join(vArray, ","), ",")
      Else
        .List = Array("Task Name")
      End If
    End With
    With .cboOperator
      .Clear
      .AddItem "equals"
      .AddItem "does not equal"
      .AddItem "contains"
      .AddItem "does not contain"
    End With
    .cboField = "Task Name"
    .chkKeepSelected = cptGetSetting("DynamicFilter", "KeepSelected") = "1"
    .chkHideSummaries = cptGetSetting("DynamicFilter", "IncludeSummaries") = "1"
    .chkShowRelatedSummaries = cptGetSetting("DynamicFilter", "RelatedSummaries") = "1"
    .chkHighlight = cptGetSetting("DynamicFilter", "Highlight") = "1"
    .tglRegEx = cptGetSetting("DynamicFilter", "geekMode") = "1"
    .chkHighlight.Visible = Not .tglRegEx
    .cboOperator.Value = cptGetSetting("DynamicFilter", "Operator")
    If .cboOperator.Value = "" Then
      If .tglRegEx Then .cboOperator.Value = "matches" Else .cboOperator = "contains"
    End If
    If Application.Edition = pjEditionProfessional Then
      .chkActiveOnly = cptGetSetting("DynamicFilter", "ActiveOnly") = "1"
      .chkActiveOnly.Enabled = True
    Else
      .chkActiveOnly = True
      .chkActiveOnly.Enabled = False
    End If
    .Show False
    .txtFilter.SetFocus
  End With
  
exit_here:
  On Error Resume Next
  Set myDynamicFilter_frm = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_bas", "cptShowDynamicFilter_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptGoRegEx(ByRef myDynamicFilter_frm As cptDynamicFilter_frm, strRegEx As String)
  'objects
  Dim oTask As MSProject.Task
  'strings
  'longs
  Dim lngFieldConstant As Long
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True
  
  If Len(strRegEx) = 0 Then
    SetAutoFilter "Marked"
    GoTo exit_here
  End If
  
  lngUID = 0
  If myDynamicFilter_frm.chkKeepSelected Then
    On Error Resume Next
    Set oTask = ActiveSelection.Tasks(1)
    lngUID = oTask.UniqueID
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  End If
  If myDynamicFilter_frm.cboField.Value = "Task Name" Then
    lngFieldConstant = FieldNameToFieldConstant("Name", pjTask)
  Else
    lngFieldConstant = FieldNameToFieldConstant(myDynamicFilter_frm.cboField.Value)
  End If
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Marked Then oTask.Marked = False
    If myDynamicFilter_frm.chkActiveOnly And Not oTask.Active Then GoTo next_task
    If Len(oTask.GetField(lngFieldConstant)) = 0 Then GoTo next_task
    If myDynamicFilter_frm.chkHideSummaries And oTask.Summary Then
      If Len(cptRxMatch(oTask.GetField(lngFieldConstant), strRegEx)) > 0 Then oTask.Marked = True
    ElseIf Not oTask.Summary Then
      If Len(cptRxMatch(oTask.GetField(lngFieldConstant), strRegEx)) > 0 Then oTask.Marked = True
    End If
    
next_task:
  Next oTask
  
  If lngUID > 0 Then ActiveProject.Tasks.UniqueID(lngUID).Marked = True
  
  FilterClear 'in case Dynamic Filter is applied
  OptionsViewEx DisplaySummaryTasks:=True
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    Sort "ID", , , , , , False, True
    OutlineShowAllTasks
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  OptionsViewEx DisplaySummaryTasks:=myDynamicFilter_frm.chkShowRelatedSummaries
  
  SetAutoFilter "Marked", pjAutoFilterFlagYes
  'todo: allow user-selected Flag or Marked
  'todo: if Dependency Browser is visible then do not allow use of Marked

exit_here:
  On Error Resume Next
  cptSpeed False
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_bas", "cptGoRegEx", Err, Erl)
  Resume exit_here
End Sub

'===============================================================================
'attribution: https://bytecomb.com/increasing-performance-of-regular-expressions-in-vba/
'Private pCachedRegexes As Dictionary 'moved to top of module - AG
 
Public Function cptGetRegex( _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As RegExp
      
    ' Ensure the dictionary has been initialized
    If pCachedRegexes Is Nothing Then Set pCachedRegexes = CreateObject("Scripting.Dictionary")
    
    ' Build the unique key for the regex: a combination
    ' of the boolean properties and the pattern itself
    Dim rxKey As String
    rxKey = IIf(IgnoreCase, "1", "0") & _
            IIf(MultiLine, "1", "0") & _
            IIf(MatchGlobal, "1", "0") & _
            Pattern
            
    ' If the RegExp object doesn't already exist, create it
    If Not pCachedRegexes.Exists(rxKey) Then
        Dim oRegExp As New RegExp
        With oRegExp
            .Pattern = Pattern
            .IgnoreCase = IgnoreCase
            .MultiLine = MultiLine
            .Global = MatchGlobal
        End With
        Set pCachedRegexes(rxKey) = oRegExp
    End If
 
    ' Fetch and return the pre-compiled RegExp object
    Set cptGetRegex = pCachedRegexes(rxKey)
End Function

Public Function cptRxTest( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Boolean
    
    ' Wow, that was easy:
    cptRxTest = cptGetRegex(Pattern, IgnoreCase, MultiLine, False).test(SourceString)
    
End Function

Public Function cptRxMatch( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant
    
    Dim oMatches As MatchCollection
    
    With cptGetRegex(Pattern, IgnoreCase, MultiLine, False)
        Set oMatches = .Execute(SourceString)
        If oMatches.Count > 0 Then
            cptRxMatch = oMatches(0).Value
        Else
            cptRxMatch = Null
        End If
    End With
 
End Function

Public Function cptRxMatches( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As Variant
 
    Dim oMatch As Match
    Dim arrMatches
    Dim lngCount As Long
    
    arrMatches = Array()
    With cptGetRegex(Pattern, IgnoreCase, MultiLine, MatchGlobal)
        For Each oMatch In .Execute(SourceString)
            ReDim Preserve arrMatches(lngCount)
            arrMatches(lngCount) = oMatch.Value
            lngCount = lngCount + 1
        Next
    End With
 
    cptRxMatches = arrMatches
End Function
