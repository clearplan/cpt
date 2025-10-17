Attribute VB_Name = "cptSaveLocal_bas"
'<cpt_version>v1.3.1</cpt_version>
Option Explicit
Public strStartView As String
Public strStartTable As String
Public strStartFilter As String
Public strStartGroup As String

Sub cptShowSaveLocal_frm()
  'objects
  Dim mySaveLocal_frm As cptSaveLocal_frm
  Dim oRange As Excel.Range
  Dim oListObject As ListObject
  Dim rstProjects As Object 'ADODB.Recordset
  Dim oSubProject As MSProject.SubProject
  Dim oMasterProject As MSProject.Project
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  Dim oTask As MSProject.Task
  Dim rstSavedMap As Object 'ADODB.Recordset
  Dim dTypes As Scripting.Dictionary
  Dim rstECF As Object 'ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strURL As String
  Dim strSaved As String
  Dim strEntity As String
  Dim strGUID As String
  Dim strECF As String
  'longs
  Dim lngECF As Long
  Dim lngLCF As Long
  Dim lngMismatchCount As Long
  Dim lngLastRow As Long
  Dim lngSubProject As Long
  Dim lngProject As Long
  Dim lngSubProjectCount As Long
  Dim lngField As Long
  Dim lngFields As Long
  Dim lngType As Long
  Dim lngECFCount As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vLine As Variant
  Dim vEntity As Variant
  Dim vType As Variant
  'dates

  'prevent spawning
  If Not cptGetUserForm("cptSaveLocal_frm") Is Nothing Then Exit Sub

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  'get server URL
  If Projects.Count = 0 Then GoTo exit_here
  If Len(ActiveProject.ServerURL) = 0 Then
    MsgBox "You are not connected to a Project Server.", vbCritical + vbOKOnly, "No Server"
    GoTo exit_here
  Else
    strURL = ActiveProject.ServerURL
  End If
  
  'setup array of types/counts
  Set dTypes = CreateObject("Scripting.Dictionary")
  'record: field type, number of available custom fields
  For Each vType In Array("Cost", "Date", "Duration", "Finish", "Start", "Outline Code")
    dTypes.Add vType, 10
  Next
  dTypes.Add "Flag", 20
  dTypes.Add "Number", 20
  dTypes.Add "Text", 30
  
  'if master/sub then ensure LCFs match
  lngSubProjectCount = ActiveProject.Subprojects.Count
  If lngSubProjectCount > 0 Then
    If MsgBox(lngSubProjectCount & " subproject(s) found." & vbCrLf & vbCrLf & "It is highly recommended that you analyze master/sub LCF matches." & vbCrLf & vbCrLf & "Do it now?", vbExclamation + vbYesNo, "Master/Sub Detected") = vbNo Then GoTo skip_it
    Set oMasterProject = ActiveProject
    Application.StatusBar = "Setting up Excel..."
    'set up Excel
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = False
    Set oWorkbook = oExcel.Workbooks.Add
    oExcel.ScreenUpdating = False
    oExcel.Calculation = xlCalculationManual
    Set oWorksheet = oWorkbook.Sheets(1)
    oExcel.ActiveWindow.Zoom = 85
    oExcel.ActiveWindow.SplitRow = 1
    oExcel.ActiveWindow.SplitColumn = 4
    oExcel.ActiveWindow.FreezePanes = True
    oWorksheet.Name = "Sync"
    'set up headers
    oWorksheet.[A1:D1] = Array("ENTITY", "TYPE", "CONSTANT", "NAME")
    'capture master and subproject names
    oWorksheet.Cells(1, 5) = oMasterProject.Name
    oWorksheet.Columns.AutoFit
    cptSpeed True
    Application.StatusBar = "Opening subprojects..."
    Set rstProjects = CreateObject("ADODB.Recordset")
    rstProjects.Fields.Append "PROJECT", adVarChar, 200
    rstProjects.Open
    rstProjects.AddNew Array(0), Array(oMasterProject.Name)
    For Each oSubProject In oMasterProject.Subprojects
      FileOpenEx oSubProject.SourceProject.FullName, True
      rstProjects.AddNew Array(0), Array(ActiveProject.Name)
    Next oSubProject
    rstProjects.MoveFirst
    Do While Not rstProjects.EOF
      Application.StatusBar = "Analyzing " & rstProjects(0) & "..."
      DoEvents
      lngLastRow = 1
      Projects(CStr(rstProjects(0))).Activate
      oWorksheet.Cells(1, 5 + CLng(rstProjects.AbsolutePosition) - 1) = rstProjects(0)
      For Each vEntity In Array(pjTask, pjResource)
        For lngType = 0 To dTypes.Count - 1
          For lngField = 1 To dTypes.Items(lngType)
            lngLastRow = lngLastRow + 1
            If lngProject = 0 Then
              oWorksheet.Cells(lngLastRow, 1) = Choose(vEntity + 1, "Task", "Resource")
              oWorksheet.Cells(lngLastRow, 2) = dTypes.Keys(lngType)
            End If
            lngLCF = FieldNameToFieldConstant(dTypes.Keys(lngType) & lngField, vEntity)
            If lngProject = 0 Then
              oWorksheet.Cells(lngLastRow, 3) = lngLCF
              oWorksheet.Cells(lngLastRow, 4) = FieldConstantToFieldName(lngLCF)
            End If
            oExcel.ActiveWindow.ScrollRow = lngLastRow
            oWorksheet.Cells(lngLastRow, 5 + CLng(rstProjects.AbsolutePosition) - 1) = CustomFieldGetName(lngLCF)
            oWorksheet.Cells.Columns.AutoFit
          Next lngField
        Next lngType
      Next vEntity
      rstProjects.MoveNext
    Loop
    'add a formula
    oWorksheet.Cells(1, 5 + rstProjects.RecordCount) = "MATCH"
    oWorksheet.Range(oWorksheet.Cells(2, 5 + rstProjects.RecordCount), oWorksheet.Cells(lngLastRow, 5 + rstProjects.RecordCount)).FormulaR1C1 = "=AND(EXACT(RC[-5],RC[-4]),EXACT(RC[-4],RC[-3]),EXACT(RC[-3],RC[-2]),EXACT(RC[-2],RC[-1]))"
    Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)), , xlYes)
    oListObject.TableStyle = ""
    oListObject.HeaderRowRange.Font.Bold = True
    'throw some shade
    With oListObject.HeaderRowRange.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -0.149998474074526
      .PatternTintAndShade = 0
    End With

    oListObject.Range.Borders(xlDiagonalDown).LineStyle = xlNone
    oListObject.Range.Borders(xlDiagonalUp).LineStyle = xlNone
    For Each vLine In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
      With oListObject.Range.Borders(vLine)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
      End With
    Next vLine
    oExcel.Calculation = xlCalculationAutomatic
    'add conditional formatting
    Set oRange = oListObject.ListColumns("MATCH").DataBodyRange
    oRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE"
    oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
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
    oRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=TRUE"
    oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
    With oRange.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With oRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    oRange.FormatConditions(1).StopIfTrue = False
    'autofilter it
    oListObject.Range.AutoFilter oRange.Column, False
    oWorksheet.Columns.AutoFit
    oExcel.ActiveWindow.ScrollRow = 1
    oMasterProject.Activate
    rstProjects.MoveFirst
    Do While Not rstProjects.EOF
      If CStr(rstProjects(0)) <> oMasterProject.Name Then
        Projects(CStr(rstProjects(0))).Activate
        Application.FileCloseEx pjDoNotSave
      End If
      rstProjects.MoveNext
    Loop
    cptSpeed False
    oExcel.ScreenUpdating = True
    oExcel.Visible = True
    oExcel.WindowState = xlMaximized
    lngMismatchCount = oRange.SpecialCells(xlCellTypeVisible).Count
    If lngMismatchCount > 0 Then
      oExcel.ActivateMicrosoftApp xlMicrosoftProject
      MsgBox lngMismatchCount & " Local Custom Fields do not match between Master and all Subprojects!", vbCritical + vbOKOnly, "Warning"
      Application.ActivateMicrosoftApp pjMicrosoftExcel
      GoTo exit_here
    Else
      oWorkbook.Close False
      oExcel.Quit
    End If
    rstProjects.Close
  End If
  
skip_it:

  'get project guid
  If CLng(Left(Application.Build, 2)) < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'capture starting view/table/filter/group
  ActiveWindow.TopPane.Activate
  strStartView = ActiveProject.CurrentView
  strStartTable = ActiveProject.CurrentTable
  strStartFilter = ActiveProject.CurrentFilter
  strStartGroup = ActiveProject.CurrentGroup
  
  'apply the ECF to LCF view
  cptUpdateSaveLocalView mySaveLocal_frm
  
  'prepare to capture all ECFs
  Set rstECF = CreateObject("ADODB.Recordset")
  rstECF.Fields.Append "URL", adVarChar, 255
  rstECF.Fields.Append "GUID", adGUID
  rstECF.Fields.Append "pjType", adInteger
  rstECF.Fields.Append "ENTITY", adVarChar, 50
  rstECF.Fields.Append "ECF", adInteger
  rstECF.Fields.Append "ECF_Name", adVarChar, 120
  rstECF.Fields.Append "LCF", adInteger
  rstECF.Fields.Append "LCF_Name", adVarChar, 120
  rstECF.Open
  
  'create a dummy task to interrogate the ECFs
  Set oTask = ActiveProject.Tasks.Add("<dummy for cpt-save-local>")
  Application.CalculateProject
  
  'populate field types
  With mySaveLocal_frm
    .Caption = "Save Local (" & cptGetVersion("cptSaveLocal_frm") & ")"
    .cboLCF.Clear
    .cboECF.Clear
    .cboECF.AddItem "All Types"
    For lngType = 0 To dTypes.Count - 1
      If dTypes.Keys(lngType) = "Start" Or dTypes.Keys(lngType) = "Finish" Then GoTo next_type
      .cboLCF.AddItem
      .cboLCF.List(.cboLCF.ListCount - 1, 0) = dTypes.Keys(lngType)
      .cboLCF.List(.cboLCF.ListCount - 1, 1) = dTypes.Items(lngType)
      .cboECF.AddItem
      .cboECF.List(.cboECF.ListCount - 1, 0) = dTypes.Keys(lngType)
      .cboECF.List(.cboECF.ListCount - 1, 1) = dTypes.Items(lngType)
next_type:
    Next lngType
    .cboECF.AddItem "undetermined"
    .cmdAutoMap.Enabled = False
    .tglAutoMap = False
    .txtAutoMap.Visible = False
    .chkAutoSwitch = True
    .cboLCF.Value = "Text"
    .optTasks = True
    
    .Show False
  
    'get enterprise custom task fields
    For lngField = 188776000 To 188778000 '2000 should do it for now
      If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
        strECF = FieldConstantToFieldName(lngField)
        strEntity = cptInterrogateECF(oTask, lngField)
        rstECF.AddNew Array("URL", "GUID", "pjType", "ENTITY", "ECF", "ECF_Name"), Array(strURL, strGUID, pjTask, strEntity, lngField, FieldConstantToFieldName(lngField))
        lngECFCount = lngECFCount + 1
      End If
      .lblStatus.Caption = "Analyzing Task ECFs...(" & Format(((lngField - 188776000) / (188778000 - 188776000)), "0%") & ")"
      .lblProgress.Width = ((lngField - 188776000) / (188778000 - 188776000)) * .lblStatus.Width
      DoEvents
    Next lngField

'    'get enterprise custom resource fields - skipped for now
'    For lngField = 205553664 To 205555664 '2000 should do it for now
'      If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
'        strECF = FieldConstantToFieldName(lngField)
'        strEntity = cptInterrogateECF(oTask, lngField)
'        rstECF.AddNew Array("URL", "GUID", "pjType", "ENTITY", "ECF", "ECF_Name"), Array(strURL, strGUID, pjResource, strEntity, lngField, FieldConstantToFieldName(lngField))
'        lngECFCount = lngECFCount + 1
'      End If
'      .lblStatus.Caption = "Analyzing Resource ECFs...(" & Format((lngField - 205553664) / (205555664 - 205553664), "0%") & ")"
'      .lblProgress.Width = ((lngField - 205553664) / (205555664 - 205553664)) * .lblStatus.Width
'      DoEvents
'    Next lngField
  
    oTask.Delete
  
    If Dir(strDir & "\settings\cpt-ecf.adtg") <> vbNullString Then
      Kill strDir & "\settings\cpt-ecf.adtg"
    End If
    rstECF.Sort = "ECF_Name"
    rstECF.Save strDir & "\settings\cpt-ecf.adtg"
    rstECF.Close
  
    'trigger lboECF refresh
    .cboECF.Value = "All Types"
    
    'update the table
    For lngField = 0 To .lboECF.ListCount - 1
      If Not IsNull(.lboECF.List(lngField, 3)) Then
        lngECF = .lboECF.List(lngField, 0)
        lngLCF = .lboECF.List(lngField, 3)
        cptUpdateSaveLocalView mySaveLocal_frm, lngECF, lngLCF
      End If
    Next lngField
        
    If cptErrorTrapping Then
      .Hide
      cptSpeed False
      .Show 'modal to control changes to custom fields
    Else
      cptSpeed False
      .Show False
    End If
  End With

exit_here:
  On Error Resume Next
  Set mySaveLocal_frm = Nothing
  Set oRange = Nothing
  Set oListObject = Nothing
  Set rstProjects = Nothing
  Set oSubProject = Nothing
  Set oMasterProject = Nothing
  If rstProjects.State Then rstProjects.Close
  Set rstProjects = Nothing
  Application.StatusBar = ""
  Set oWorksheet = Nothing
  oExcel.Calculation = xlCalculationAutomatic
  oExcel.ScreenUpdating = True
  oExcel.Visible = True
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  cptSpeed False
  oTask.Delete
  Set oTask = Nothing
  Set rstSavedMap = Nothing
  Set vType = Nothing
  Set dTypes = Nothing
  If rstECF.State Then rstECF.Close
  Set rstECF = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptShowSaveLocal_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptSaveLocal(ByRef mySaveLocal_frm As cptSaveLocal_frm)
  'objects
  Dim rstSavedMap As Object 'ADODB.Recordset
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  Dim oResources As Resources
  Dim oResource As Resource
  'strings
  Dim strErrors As String
  Dim strGUID As String
  Dim strSavedMap As String
  Dim strMsg As String
  'longs
  Dim lngType As Long
  Dim lngItems As Long
  Dim lngLCF As Long
  Dim lngECF As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'get project guid
  If CLng(Left(Application.Build, 2)) < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'save map
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) = vbNullString Then 'create it
    rstSavedMap.Fields.Append "URL", adVarChar, 255
    rstSavedMap.Fields.Append "GUID", adGUID
    rstSavedMap.Fields.Append "ECF", adBigInt
    rstSavedMap.Fields.Append "LCF", adBigInt
    rstSavedMap.Open
  Else
    'todo: allow multiple maps per user
    'replace existing saved map
    rstSavedMap.Filter = "GUID<>'" & strGUID & "'"
    rstSavedMap.Open strSavedMap
    rstSavedMap.Save strSavedMap, adPersistADTG
  End If
  'get total task count
  ActiveWindow.TopPane.Activate
  'determine whether we're working with tasks or resources
  If mySaveLocal_frm.optTasks Then
    lngType = pjTask
    If ActiveProject.CurrentView <> ".cptSaveLocal Task View" Then
      ViewApply ".cptSaveLocal Task View"
    End If
  ElseIf mySaveLocal_frm.optResources Then
    lngType = pjResource
    If ActiveProject.CurrencyCode <> ".cptSaveLocal Resource View" Then
      ViewApply ".cptSaveLocal Resource View"
    End If
  End If
  FilterClear
  GroupClear
  On Error Resume Next
  If lngType = pjTask Then
    On Error Resume Next
    If Not OutlineShowAllTasks Then
      Sort "ID", , , , , , False, True
      OutlineShowAllTasks
    End If
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    SelectAll
    Set oTasks = ActiveSelection.Tasks
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oTasks Is Nothing Then
      lngItems = oTasks.Count
    Else
      MsgBox "There are no tasks in this schedule.", vbCritical + vbOKOnly, "No Tasks"
      GoTo exit_here
    End If
  ElseIf lngType = pjResource Then
    SelectAll
    Set oResources = ActiveSelection.Resources
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oResources Is Nothing Then
      lngItems = oResources.Count
    Else
      MsgBox "There are no resources in this schedule.", vbCritical + vbOKOnly, "No Resources"
      GoTo exit_here
    End If
  End If
  
  With mySaveLocal_frm
    If lngType = pjTask Then
      For Each oTask In ActiveProject.Tasks
        If oTask Is Nothing Then GoTo next_task
        If oTask.ExternalTask Then GoTo next_task
        On Error Resume Next
        For lngItem = 0 To .lboECF.ListCount - 1
          If .lboECF.List(lngItem, 3) > 0 Then
            lngECF = .lboECF.List(lngItem, 0)
            lngLCF = .lboECF.List(lngItem, 3)
            'no duplicates
            'todo: does Filter = X AND (Y OR Z) work?
  '          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & lngECF
  '          If rstSavedMap.RecordCount = 1 Then
              'overwrite it
  '            rstSavedMap.Delete adAffectCurrent
  '          End If
  '          rstSavedMap.Filter = ""
  '          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND LCF=" & lngLCF
  '          If rstSavedMap.RecordCount = 1 Then
              'overwrite it
  '            rstSavedMap.Delete adAffectCurrent
  '          End If
  '          rstSavedMap.Filter = ""
            'add the new record
  '          rstSavedMap.AddNew Array(0, 1, 2), Array(strGUID, lngECF, lngLCF)
            'first clear the values
            If Len(oTask.GetField(lngLCF)) > 0 Then oTask.SetField lngLCF, ""
            'if ECF is formula, then skip it
            If Len(CustomFieldGetFormula(lngECF)) > 0 Then GoTo next_task_mapping
            If Len(oTask.GetField(lngECF)) > 0 Then
              oTask.SetField lngLCF, CStr(oTask.GetField(lngECF))
              If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
              If oTask.GetField(lngLCF) <> CStr(oTask.GetField(lngECF)) Then
                If MsgBox("There was an error copying from ECF " & CustomFieldGetName(lngECF) & " to LCF " & CustomFieldGetName(lngLCF) & " on Task UID " & oTask.UniqueID & "." & vbCrLf & vbCrLf & "Please validate data type mapping." & vbCrLf & vbCrLf & "Proceed anyway?", vbExclamation + vbYesNo, "Failed!") = vbNo Then
                  GoTo exit_here
                End If
              End If
            End If
          End If
next_task_mapping:
        Next lngItem
next_task:
      Next oTask
    ElseIf lngType = pjResource Then
      'todo: how to handle enterprise resources?
      'todo: - copy down to master project?
      'todo: prompt to copy down to each subproject, and
      'todo: ResourceAssignment Resources:="pgm mgmt", Operation:=pjReplace, With:="Agile SW Eng"
      'todo: -- note! does not move baseline work; resource must exist in the subproject; prompts user on EACH for completed work; if same name then pjReplace defaults to Enterprise
      'todo: issue: cannot write to resource until it's local; cannot read from resource if no longer enterprise
      'todo: - create copy of enterprise resource
      For Each oResource In oResources
        'cannot write to local custom fields on an enterprise resource
        If oResource.Enterprise Then
          EditGoTo oResource.ID
          MsgBox "'" & oResource.Name & "' is an enterprise resource:" & vbCrLf & vbCrLf & "Cannot write to local custom fields of an enterprise resource." & vbCrLf & vbCrLf & "This resource will be skipped.", vbExclamation + vbOKOnly, "Not a Local Resource"
'          If MsgBox("'" & oResource.Name & "' is an enterprise resource:" & vbCrLf & vbCrLf & "Cannot write to local custom fields of an enterprise resource." & vbCrLf & vbCrLf & "Save to local resource pool?", vbExclamation + vbOKOnly, "Not a Local Resource") = vbYes Then
'            'todo: where to save to master/sub?
'          End If
          GoTo next_resource
        End If
        On Error Resume Next
        For lngItem = 0 To .lboECF.ListCount - 1
          If .lboECF.List(lngItem, 3) > 0 Then
            lngECF = .lboECF.List(lngItem, 0)
            lngLCF = .lboECF.List(lngItem, 3)
            'first clear the values
            If Len(oResource.GetField(lngLCF)) > 0 Then oResource.SetField lngLCF, ""
            'if ECF is formula, then skip it
            If Len(CustomFieldGetFormula(lngECF)) > 0 Then GoTo next_resource_mapping
            If Len(oResource.GetField(lngECF)) > 0 Then
              oResource.SetField lngLCF, CStr(oResource.GetField(lngECF))
              If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
              If oResource.GetField(lngLCF) <> CStr(oResource.GetField(lngECF)) Then
                strMsg = "There was an error copying..." & vbCrLf
                strMsg = strMsg & "- from ECF: '" & CustomFieldGetName(lngECF) & "'" & vbCrLf
                strMsg = strMsg & "- to LCF: '" & CustomFieldGetName(lngLCF) & "'" & vbCrLf
                strMsg = strMsg & "- on Resource: '" & oResource.Name & "':" & vbCrLf
                strMsg = strMsg & vbCrLf & "Please validate data type mapping, and verify subproject custom fields match master project custom fields."
                strMsg = strMsg & vbCrLf & vbCrLf & "Click 'OK' to continue; 'Cancel' to cancel"
                If MsgBox(strMsg, vbExclamation + vbOKCancel, "Failed!") = vbCancel Then
                  GoTo exit_here
                End If
              End If
            End If
          End If
next_resource_mapping:
        Next lngItem
next_resource:
      Next oResource
    End If
  End With

'  rstSavedMap.Save strSavedMap, adPersistADTG

  MsgBox "Enteprise Custom Fields saved locally.", vbInformation + vbOKOnly, "Complete"

exit_here:
  On Error Resume Next
  Set oResource = Nothing
  Set oResources = Nothing
  Set oTasks = Nothing
  rstSavedMap.Close
  Set rstSavedMap = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptSaveLocal", Err, Erl)
  Resume exit_here
End Sub

Function cptInterrogateECF(ByRef oTask As MSProject.Task, lngField As Long)
  'objects
  Dim oOutlineCode As OutlineCode
  'strings
  Dim strPattern As String
  Dim strVal As String
  'longs
  Dim lngItem As Long
  Dim lngVal As Long
  'integers
  'doubles
  'booleans
  Dim blnVal As Boolean
  'variants
  'dates
  Dim dtVal As Date
  
  On Error Resume Next
  
  'check for outlinecode requirement (has parent-child structure)
  Set oOutlineCode = Application.GlobalOutlineCodes(FieldConstantToFieldName(lngField))
  If Not oOutlineCode Is Nothing Then
    If oOutlineCode.CodeMask.Count > 1 Then
      cptInterrogateECF = "Outline Code"
      GoTo exit_here
    Else
      If oOutlineCode.CodeMask(1).Sequence = 4 Then
        cptInterrogateECF = "Date"
      ElseIf oOutlineCode.CodeMask(1).Sequence = 5 Then
        cptInterrogateECF = "Cost"
      ElseIf oOutlineCode.CodeMask(1).Sequence = 7 Then
        cptInterrogateECF = "Number"
      Else
        cptInterrogateECF = "Text"
      End If
      GoTo exit_here
    End If
  End If
   
  oTask.SetField lngField, "xxx"

  If Err.Description = "This field only supports positive numbers." Then
    cptInterrogateECF = "Cost"
  ElseIf Err.Description = "The date you entered isn't supported for this field." Then
    cptInterrogateECF = "Date"
  ElseIf Err.Description = "The duration you entered isn't supported for this field." Then
    cptInterrogateECF = "Duration"
  ElseIf Err.Description = "Select either Yes or No from the list." Then
    cptInterrogateECF = "Flag"
  ElseIf Err.Description = "This field only supports numbers." Then
    cptInterrogateECF = "Number"
  ElseIf Err.Description = "This is not a valid lookup table value." Or Err.Description = "The value you entered does not exist in the lookup table of this code" Then
    'select the first value and check it
    oTask.SetField lngField, oOutlineCode.LookupTable(1).Name
    strVal = oTask.GetField(lngField)
    GoTo enhanced_interrogation
  ElseIf Err.Description = "The argument value is not valid." Then
    'figure out formula
    If Len(CustomFieldGetFormula(lngField)) > 0 Then
      strVal = oTask.GetField(lngField)
      GoTo enhanced_interrogation
    End If
  ElseIf Err.Description = "" Then
    cptInterrogateECF = "Text"
  Else
    GoTo enhanced_interrogation
  End If
  
  GoTo exit_here
  
enhanced_interrogation:
  
  Err.Clear
  
  'check for cost
  If InStr(strVal, ActiveProject.CurrencySymbol) > 0 Then
    cptInterrogateECF = "Cost"
    GoTo exit_here
  End If
  
  'check for number
  On Error Resume Next
  lngVal = oTask.GetField(lngField)
  If Err.Number = 0 And Len(oTask.GetField(lngField)) = Len(CStr(lngVal)) Then
    cptInterrogateECF = "Number"
    GoTo exit_here
  End If
  
  'check for date
  On Error Resume Next
  dtVal = oTask.GetField(lngField)
  If Err.Number = 0 Then
    cptInterrogateECF = "Date"
    GoTo exit_here
  End If
  
  'could be flag
  If Len(cptRegEx(strVal, "Yes|No")) > 0 Then
    On Error Resume Next
    Set oOutlineCode = GlobalOutlineCodes(FieldConstantToFieldName(lngField))
    If oOutlineCode Is Nothing Then
      cptInterrogateECF = "MaybeFlag"
    Else
      cptInterrogateECF = "Text"
    End If
    GoTo exit_here
  End If
  
  On Error Resume Next
  strVal = oTask.GetField(lngField)
  'could be duration
  If strVal = DurationFormat(DurationValue(strVal), ActiveProject.DefaultDurationUnits) Then
    If Err.Number = 0 Then
      cptInterrogateECF = "Duration"
      GoTo exit_here
    End If
  End If
  
  'otherwise, it's most likely text
  cptInterrogateECF = "Text"

exit_here:
  On Error Resume Next
  Set oOutlineCode = Nothing
  
  Exit Function
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptInterrogateECF", Err, Erl)
  Resume exit_here
End Function

Sub cptGetAllFields(lngFrom As Long, lngTo As Long)
  'objects
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim rst As Object 'ADODB.Recordset
  Dim oExcel As Excel.Application
  'strings
  Dim strCustomName As String
  Dim strDir As String
  Dim strFieldName As String
  'longs
  'Dim lngTo As Long
  'Dim lngFrom As Long
  Dim lngFile As Long
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  GoTo exit_here
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set rst = CreateObject("ADODB.Recordset")
  rst.Fields.Append "Constant", adBigInt
  rst.Fields.Append "Name", adVarChar, 155
  rst.Fields.Append "CustomName", adVarChar, 155
  rst.Open
  
  '184549399 = lowest = ID
  '184550803 start of <Unavailable>
  '188744879 might be last of built-ins
  '188750001 start of ecfs?
  '218103807 highest and enterprise
  
  'restart at 188800001
  'lngFrom = 215000001
  'lngTo = 218103807
  
  For lngField = lngFrom To lngTo
    strFieldName = FieldConstantToFieldName(lngField)
    If Len(strFieldName) > 0 And strFieldName <> "<Unavailable>" Then
      strCustomName = CustomFieldGetName(lngField)
      rst.AddNew Array(0, 1, 2), Array(lngField, strFieldName, strCustomName)
    End If
    Debug.Print "Processing " & Format(lngField, "###,###,##0") & " of " & Format(lngTo, "###,###,##0") & " (" & Format(lngField / lngTo, "0%") & ")"
  Next lngField

  If rst.RecordCount > 0 Then
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.[A1].CopyFromRecordset rst
  Else
    MsgBox "No fields found between " & lngFrom & " and " & lngTo & ".", vbInformation + vbOKOnly, "No results."
  End If
exit_here:
  On Error Resume Next
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  rst.Close
  Set rst = Nothing
  Set oExcel = Nothing
  Close #lngFile
  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptGetAllFields", Err, Erl)
  Resume exit_here
End Sub

Sub cptAnalyzeAutoMap(ByRef mySaveLocal_frm As cptSaveLocal_frm)
  'objects
  Dim rstAvailable As Object 'ADODB.Recordset
  Dim dTypes As Scripting.Dictionary
  'strings
  Dim strMsg As String
  'longs
  Dim lngItem2 As Long
  Dim lngAvailable As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vType As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Set rstAvailable = CreateObject("ADODB.Recordset")
  With rstAvailable
    .Fields.Append "TYPE", adVarChar, 50
    .Fields.Append "ECF", adInteger
    .Fields.Append "LCF", adInteger
    .Open

    Set dTypes = CreateObject("Scripting.Dictionary")
    'record: field type, number of available custom fields
    For Each vType In Array("Cost", "Date", "Duration", "Finish", "Start", "Outline Code")
      dTypes.Add vType, 10
      .AddNew Array(0, 1, 2), Array(vType, 0, 10)
    Next
    dTypes.Add "Flag", 20
    .AddNew Array(0, 1, 2), Array("Flag", 0, 20)
    dTypes.Add "Number", 20
    .AddNew Array(0, 1, 2), Array("Number", 0, 20)
    dTypes.Add "Text", 30
    .AddNew Array(0, 1, 2), Array("Text", 0, 30)
    .Update
    .Sort = "TYPE"
    
    'todo: start->date;finish->date;date->date
    
    'get available LCF
    For lngItem = 0 To dTypes.Count - 1
      For lngItem2 = 1 To dTypes.Items(lngItem)
        'todo: account for both pjTask and pjResource
        If Len(CustomFieldGetName(FieldNameToFieldConstant(dTypes.Keys(lngItem) & lngItem2))) > 0 Then
          .MoveFirst
          .Find "TYPE='" & dTypes.Keys(lngItem) & "'"
          If Not .EOF Then
            .Fields(2) = .Fields(2) - 1
          End If
        End If
      Next lngItem2
    Next lngItem
    
    'get total ECF
    For lngItem = 0 To mySaveLocal_frm.lboECF.ListCount - 1
      If mySaveLocal_frm.lboECF.Selected(lngItem) Then
        .MoveFirst
        .Find "TYPE='" & Replace(mySaveLocal_frm.lboECF.List(lngItem, 2), "Maybe", "") & "'"
        If Not .EOF Then
          If IsNull(mySaveLocal_frm.lboECF.List(mySaveLocal_frm.lboECF.ListIndex, 3)) Then
            'only count unmapped
            .Fields(1) = .Fields(1) + 1
          End If
        End If
      End If
    Next lngItem
    
    'return result
    strMsg = strMsg & String(34, "-") & vbCrLf
    strMsg = strMsg & "| " & "TYPE" & String(10, " ") & "|"
    strMsg = strMsg & " ECF |"
    strMsg = strMsg & " LCF |"
    strMsg = strMsg & " <> |" & vbCrLf
    strMsg = strMsg & String(34, "-") & vbCrLf
    .MoveFirst
    Do While Not .EOF
      strMsg = strMsg & "| " & rstAvailable(0) & String(14 - Len(rstAvailable(0)), " ") & "|"
      If rstAvailable(0) = "Start" Or rstAvailable(0) = "Finish" Then
        strMsg = strMsg & "   - |"
      Else
        strMsg = strMsg & String(4 - Len(CStr(rstAvailable(1))), " ") & rstAvailable(1) & " |"
      End If
      strMsg = strMsg & String(4 - Len(CStr(rstAvailable(2))), " ") & rstAvailable(2) & " |"
      strMsg = strMsg & IIf(rstAvailable(2) >= rstAvailable(1), " ok ", "  X ") & "|" & vbCrLf
      .MoveNext
    Loop
    strMsg = strMsg & String(34, "-") & vbCrLf
    mySaveLocal_frm.cmdAutoMap.Enabled = False
    If InStr(strMsg, "  X ") > 0 Then
      strMsg = strMsg & "AutoMap is NOT available." & vbCrLf
      strMsg = strMsg & "Free up some fields and try again."
    Else
      strMsg = strMsg & "AutoMap IS available." & vbCrLf
      strMsg = strMsg & "Click GO! to AutoMap now."
      If mySaveLocal_frm.tglAutoMap Then
        mySaveLocal_frm.cmdAutoMap.Enabled = True
      Else
        mySaveLocal_frm.cmdAutoMap.Enabled = False
      End If
    End If
    
    mySaveLocal_frm.txtAutoMap.Value = strMsg
    
    .Close
    
  End With
  
exit_here:
  On Error Resume Next
  If rstAvailable.State Then rstAvailable.Close
  Set rstAvailable = Nothing
  Set dTypes = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptAutoMap", Err, Erl)
  Resume exit_here
End Sub

Sub cptAutoMap(ByRef mySaveLocal_frm As cptSaveLocal_frm)
  'objects
  'strings
  'longs
  Dim lngECF As Long
  Dim lngLCF As Long
  Dim lngLCFs As Long
  Dim lngECFs As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'todo: unselect after complete - if fails, leave selected
  'todo: hide analysis after AutoMap; cptRefreshLCF
  'todo: for dates > 10 cycle through date, start, finish

  With mySaveLocal_frm
    .lblStatus.Caption = "AutoMapping..."
    'loop through ECFs looking for selected ECFs to map
    For lngECFs = 0 To .lboECF.ListCount - 1
      If .lboECF.Selected(lngECFs) Then
        lngECF = .lboECF.List(lngECFs, 0)
        'switch cbo types to get list of lngLCFs
        If .cboLCF <> .lboECF.List(lngECFs, 2) Then .cboLCF = Replace(.lboECF.List(lngECFs, 2), "Maybe", "")
        'loop through LCFs looking for one available
        For lngLCFs = 0 To .lboLCF.ListCount - 1
          lngLCF = .lboLCF.List(lngLCFs, 0)
          If Len(CustomFieldGetName(lngLCF)) = 0 Then
            Call cptMapECFtoLCF(mySaveLocal_frm, lngECF, lngLCF)
            .lboECF.List(lngECFs, 3) = lngLCF
            .lboECF.List(lngECFs, 4) = CustomFieldGetName(lngLCF)
            cptAnalyzeAutoMap mySaveLocal_frm
            Exit For
          End If
        Next lngLCFs
      End If
      .lblProgress.Width = (lngECFs / (.lboECF.ListCount - 1)) * .lblStatus.Width
    Next lngECFs
    .lblStatus.Caption = "AutoMap complete."
    .lblProgress.Width = .lblStatus.Width
    
    If MsgBox("Fields AutoMapped. Import field data now?", vbQuestion + vbYesNo, "Save Local") = vbYes Then
      .cmdSaveLocal.SetFocus
    End If
    
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptAutoMap", Err, Erl)
  Resume exit_here
End Sub

Sub cptMapECFtoLCF(ByRef mySaveLocal_frm As cptSaveLocal_frm, lngECF As Long, lngLCF As Long)
  'objects
  Dim rstSavedMap As Object 'ADODB.Recordset
  Dim oLookupTableEntry As LookupTableEntry
  Dim oOutlineCodeLocal As OutlineCode
  Dim oOutlineCode As OutlineCode
  'strings
  Dim strUnmapped As String
  Dim strFields As String
  Dim strField As String
  Dim strCustomFormula As String
  Dim strURL As String
  Dim strGUID As String
  Dim strSavedMap As String
  Dim strECF As String
  'longs
  Dim lngSourceField As Long
  Dim lngItem As Long
  Dim lngDown As Long
  Dim lngCodeNumber As Long
  'integers
  'doubles
  'booleans
  Dim blnProceed As Boolean
  'variants
  Dim vField As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'todo: deletecustomfield if overwriting

  With mySaveLocal_frm
    'if already mapped then prompt with ECF name and ask to remap
    For lngItem = 0 To .lboECF.ListCount - 1
      If .lboECF.List(lngItem, 3) = lngLCF Then
        If .lboECF.List(lngItem, 0) = lngECF Then GoTo exit_here 'already mapped
        If MsgBox(FieldConstantToFieldName(lngLCF) & " is already mapped to " & .lboECF.List(lngItem, 1) & " - reassign it?", vbExclamation + vbYesNo, "Already Mapped") = vbYes Then '
          .lboECF.List(lngItem, 3) = ""
          .lboECF.List(lngItem, 4) = ""
        Else
          GoTo exit_here
        End If
      End If
next_ecf:
    Next lngItem
    
    CustomFieldDelete lngLCF
    
    'capture rename
    If CustomFieldGetName(lngECF) = Replace(CustomFieldGetName(lngLCF), " (" & FieldConstantToFieldName(lngLCF) & ")", "") Then GoTo skip_rename
    If Len(CustomFieldGetName(lngLCF)) > 0 Then
      'todo: skip if name already matches
      If MsgBox("Rename " & FieldConstantToFieldName(lngLCF) & " to " & FieldConstantToFieldName(lngECF) & "?", vbQuestion + vbYesNo, "Please confirm") = vbYes Then
        'rename it
        CustomFieldRename CLng(lngLCF), CustomFieldGetName(lngECF) & " (" & FieldConstantToFieldName(lngLCF) & ")"
        'rename in lboLCF
        If Not .tglAutoMap Then .lboLCF.List(.lboLCF.ListIndex, 1) = FieldConstantToFieldName(lngLCF) & " (" & CustomFieldGetName(lngLCF) & ")"
      Else
        GoTo exit_here
      End If
    Else
      ActiveWindow.TopPane.Activate
      'rename it in msp
      CustomFieldRename lngLCF, CustomFieldGetName(lngECF) & " (" & FieldConstantToFieldName(lngLCF) & ")"
      'rename it in lboLCF
      If Not .tglAutoMap Then .lboLCF.List(.lboLCF.ListIndex, 1) = FieldConstantToFieldName(.lboLCF) & " (" & CustomFieldGetName(.lboLCF) & ")"
    End If
    
skip_rename:
    
    'does ECF have a formula?
    strCustomFormula = CustomFieldGetFormula(lngECF)
    If Len(strCustomFormula) > 0 Then
      'does ECF rely on other fields?
      strField = ""
      Do While Len(cptRegEx(Replace(strCustomFormula, strField, ""), "\[[^\]]*\]")) > 0
        strField = cptRegEx(Replace(strCustomFormula, strField, ""), "\[[^\]]*\]")
        If InStr(strFields, strField) = 0 Then
          strFields = strFields & strField & "|"
        End If
        strCustomFormula = Replace(strCustomFormula, strField, "")
      Loop
      blnProceed = True
      'reset the temporary formula string
      strCustomFormula = CustomFieldGetFormula(lngECF)
      For Each vField In Split(strFields, "|")
        'is input field an ECF?
        If Len(vField) = 0 Then Exit For
        'skip status date no need to map this field
        If vField = "[Status Date]" Then GoTo next_formula_field
        lngSourceField = FieldNameToFieldConstant(Replace(Replace(vField, "[", ""), "]", ""))
        If lngSourceField > 188776000 Then
          'has input field been mapped?
          For lngItem = 0 To .lboECF.ListCount - 1
            If .lboECF.List(lngItem, 0) = lngSourceField Then
              If IsNull(.lboECF.List(lngItem, 3)) Then
                strUnmapped = strUnmapped & "- " & FieldConstantToFieldName(lngSourceField) & vbCrLf
                blnProceed = False
                Exit For
              Else
                strCustomFormula = Replace(strCustomFormula, "[" & FieldConstantToFieldName(lngSourceField) & "]", "[" & CustomFieldGetName(.lboECF.List(lngItem, 3)) & "]")
                Exit For
              End If
            End If
          Next lngItem
        End If
next_formula_field:
      Next vField
      If Not blnProceed Then
        MsgBox "ECF '" & FieldConstantToFieldName(lngECF) & "' contains a formula:" & vbCrLf & vbCrLf & strCustomFormula & vbCrLf & vbCrLf & "Please map dependent ECFs first:" & vbCrLf & strUnmapped, vbInformation + vbOKOnly, "Not yet..."
        CustomFieldDelete lngLCF
        cptUpdateLCF mySaveLocal_frm
        GoTo exit_here
      Else
        CustomFieldPropertiesEx lngLCF, pjFieldAttributeFormula
        If Not CustomFieldSetFormula(lngLCF, strCustomFormula) Then  'CustomFieldGetFormula(lngECF)
          MsgBox "Problem importing formula; please validate.", vbCritical + vbOKOnly, "Formula Error"
        End If
      End If
    End If
    
    'get indicators
    'todo: warn user these are not exposed/available
    
    'get pick list
    strECF = CustomFieldGetName(lngECF)
    On Error Resume Next
    Set oOutlineCode = GlobalOutlineCodes(strECF)
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oOutlineCode Is Nothing Then
      'make it a picklist
      CustomFieldPropertiesEx lngLCF, pjFieldAttributeValueList
      If oOutlineCode.CodeMask.Count > 1 Then 'import outline code and all settings
        'capture code mask
        With oOutlineCode.CodeMask
          For lngItem = 1 To .Count
            CustomOutlineCodeEditEx lngLCF, .Item(lngItem).Level, .Item(lngItem).Sequence, .Item(lngItem).Length, .Item(lngItem).Separator
          Next lngItem
        End With
        'capture picklist
        Set oOutlineCodeLocal = ActiveProject.OutlineCodes(CustomFieldGetName(lngLCF))
        With oOutlineCode.LookupTable
          'load items bottom to top
          For lngItem = .Count To 1 Step -1
            Set oLookupTableEntry = oOutlineCodeLocal.LookupTable.AddChild(.Item(lngItem).Name)
            If Len(.Item(lngItem).Description) > 0 Then
              oLookupTableEntry.Description = .Item(lngItem).Description
            End If
          Next lngItem
          'indent top to bottom
          For lngItem = 1 To .Count
            oOutlineCodeLocal.LookupTable.Item(lngItem).Level = .Item(lngItem).Level
          Next lngItem
        End With
        'capture other options
        CustomOutlineCodeEditEx FieldID:=lngLCF, OnlyLookUpTableCodes:=oOutlineCode.OnlyLookUpTableCodes, OnlyCompleteCodes:=oOutlineCode.OnlyCompleteCodes
        CustomOutlineCodeEditEx FieldID:=lngLCF, OnlyLeaves:=oOutlineCode.OnlyLeaves
        'todo: next line for RequiredCode throws an error, not sure why
        'CustomOutlineCodeEditEx FieldID:=lngLCF, RequiredCode:=oOutlineCode.RequiredCode
        If oOutlineCode.DefaultValue <> "" Then CustomOutlineCodeEditEx FieldID:=lngLCF, DefaultValue:=oOutlineCode.DefaultValue
        CustomOutlineCodeEditEx FieldID:=lngLCF, SortOrder:=oOutlineCode.SortOrder
      Else 'import just the pick list
        For lngItem = 1 To oOutlineCode.LookupTable.Count
          CustomFieldValueListAdd lngLCF, oOutlineCode.LookupTable(lngItem).Name, oOutlineCode.LookupTable(lngItem).Description
        Next lngItem
        
      End If
    End If
    If Not .tglAutoMap Then
      .lboECF.List(.lboECF.ListIndex, 3) = lngLCF
      .lboECF.List(.lboECF.ListIndex, 4) = CustomFieldGetName(lngLCF)
    End If
  End With
  
  strURL = ActiveProject.ServerURL
  
  'update rstSavedMap
  If CLng(Left(Application.Build, 2)) < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) <> vbNullString Then
    rstSavedMap.Open strSavedMap
    rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & lngECF
    If Not rstSavedMap.EOF Then
      rstSavedMap.Fields(2) = lngLCF
    Else
      rstSavedMap.AddNew Array(0, 1, 2, 3), Array(strURL, strGUID, lngECF, lngLCF)
    End If
    rstSavedMap.Filter = ""
    rstSavedMap.Save strSavedMap, adPersistADTG
  Else 'create it
    rstSavedMap.Fields.Append "URL", adVarChar, 255
    rstSavedMap.Fields.Append "GUID", adGUID
    rstSavedMap.Fields.Append "ECF", adInteger
    rstSavedMap.Fields.Append "LCF", adInteger
    rstSavedMap.Open
    rstSavedMap.AddNew Array(0, 1, 2, 3), Array(strURL, strGUID, lngECF, lngLCF)
    rstSavedMap.Save strSavedMap, adPersistADTG
  End If
  rstSavedMap.Close
  
  'update the table
  cptUpdateSaveLocalView mySaveLocal_frm, lngECF, lngLCF
  
exit_here:
  On Error Resume Next
  Set rstSavedMap = Nothing
  Set oLookupTableEntry = Nothing
  Set oOutlineCodeLocal = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptMapECFtoLCF", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCFMap()
  'objects
  Dim rstSavedMap As Object 'ADODB.Recordset
  'strings
  Dim strMsg As String
  Dim strSavedMapExport As String
  Dim strGUID As String
  Dim strSavedMap As String
  'longs
  Dim lngFile As Long
  Dim lngProjectCount As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strMsg = "Your maps are only valid for other users on this server:" & vbCrLf & vbCrLf
  strMsg = strMsg & ActiveProject.ServerURL & vbCrLf & vbCrLf
  strMsg = strMsg & "Proceed with export?"
  If MsgBox(strMsg, vbExclamation + vbYesNo, "Important") = vbNo Then GoTo exit_here
  
  'ensure file exists
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) = vbNullString Then
    MsgBox "You have no saved map for this project.", vbExclamation + vbOKOnly, "No Map"
    GoTo exit_here
  End If
  
  If CLng(Left(Application.Build, 2)) < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'prepare an export csv file
  lngFile = FreeFile
  strSavedMapExport = Environ("USERPROFILE") & "\Downloads\"
  If Dir(strSavedMapExport, vbDirectory) = vbNullString Then
    strSavedMapExport = Environ("USERPROFILE")
  End If
  strSavedMapExport = strSavedMapExport & "cpt-saved-map.csv"
  If Dir(strSavedMapExport) <> vbNullString Then Kill strSavedMapExport
  Open strSavedMapExport For Output As #lngFile
  'open the filtered recordset and export it
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  With rstSavedMap
    .Open strSavedMap, "Provider=MSPersist", , , adCmdFile
    .Filter = "GUID='" & UCase(strGUID) & "'"
    If .RecordCount = 0 Then
      MsgBox "You have no saved map for this project.", vbExclamation + vbOKOnly, "No Map"
    Else
      Print #lngFile, .GetString(adClipString, , ",", vbCrLf, vbNullString)
    End If
    Close #lngFile
    .Filter = ""
    .Close
  End With
  
  MsgBox "Map saved to '" & strSavedMapExport & "'", vbInformation + vbOKOnly, "Export Complete"
    
exit_here:
  On Error Resume Next
  rstSavedMap.Close
  Set rstSavedMap = Nothing
  Close #lngFile
  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptSaveLocal_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportCFMap(ByRef mySaveLocal_frm As cptSaveLocal_frm)
  'objects
  Dim rstSavedMap As Object 'ADODB.Recordset
  Dim oStream As Scripting.TextStream
  Dim oFile As Scripting.File
  Dim oFSO As Scripting.FileSystemObject
  Dim oExcel As Excel.Application
  Dim oFileDialog As Object 'FileDialog
  'strings
  Dim strDir As String
  Dim strGUID As String
  Dim strConn As String
  Dim strSavedMapImport As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  Dim aLine As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  'get guid
  If CLng(Left(Application.Build, 2)) < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
    
  'borrow Excel's FileDialogFilePicker
  Set oExcel = CreateObject("Excel.Application")
  Set oFileDialog = oExcel.FileDialog(msoFileDialogFilePicker)
  With oFileDialog
    .AllowMultiSelect = False
    .ButtonName = "Import"
    .InitialView = msoFileDialogViewDetails
    .InitialFileName = Environ("USERPROFILE") & "\Downloads\"
    .Title = "Select cpt-saved-map.csv:"
    .Filters.Add "Comma Separated Values (csv)", "*.csv"
    If .Show = -1 Then
      strSavedMapImport = .SelectedItems(1)
    End If
  End With
  'close Excel, with thanks...
  oExcel.Quit
  
  'stream the csv
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFile = oFSO.GetFile(strSavedMapImport)
  Set oStream = oFile.OpenAsTextStream(ForReading)
  
  'open user's saved map
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  With rstSavedMap
    If Dir(strDir & "\settings\cpt-save-local.adtg") = vbNullString Then 'create it
      .Fields.Append "ServerURL", adVarChar, 255
      .Fields.Append "GUID", adGUID
      .Fields.Append "ECF", adInteger
      .Fields.Append "LCF", adInteger
      .Save strSavedMapImport, adPersistADTG
    End If
    .Open strDir & "\settings\cpt-save-local.adtg", "Provider=MSPersist", , , adCmdFile
    
    Do Until oStream.AtEndOfStream
      aLine = Split(oStream.ReadLine, ",")
      If UBound(aLine) > 0 Then
        'ensure same server
        If aLine(0) <> ActiveProject.ServerURL Then
          MsgBox "Maps can only be imported from this server:" & vbCrLf & vbCrLf & ActiveProject.ServerURL, vbCritical + vbOKOnly, "Invalid Map"
          GoTo exit_here
        Else
          'ensure ecf exists
          If FieldConstantToFieldName(CLng(aLine(2))) = "<Unavailable>" Then
            'todo: strECF
            MsgBox "The imported ECF (" & CLng(aLine(2)) & ") is <Unavailable> on this server.", vbCritical + vbOKOnly, "Invalid ECF"
          Else
            'todo: ensure same ECF settings?
            mySaveLocal_frm.lboECF.Value = CLng(aLine(2))
            mySaveLocal_frm.lboLCF.Value = CLng(aLine(3))
            Call cptMapECFtoLCF(mySaveLocal_frm, CLng(aLine(2)), CLng(aLine(3)))
          End If
        End If
      End If
    Loop
    
  End With

exit_here:
  On Error Resume Next
  Set rstSavedMap = Nothing
  Set oStream = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing
  oExcel.Quit
  Set oExcel = Nothing
  Set oFileDialog = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptImportCFMap", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateECF(ByRef mySaveLocal_frm As cptSaveLocal_frm, Optional strFilter As String)
  'objects
  Dim rstECF As Object 'ADODB.Recordset
  Dim rstSavedMap As Object 'ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strGUID As String
  Dim strSavedMap As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnExists As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  
  'get project guid
  If CLng(Left(Application.Build, 2)) < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'open the ecf recordset
  If Dir(strDir & "\settings\cpt-ecf.adtg") = vbNullString Then GoTo exit_here
  Set rstECF = CreateObject("ADODB.Recordset")
  rstECF.Open strDir & "\settings\cpt-ecf.adtg"
    
  'check for saved map
  strSavedMap = strDir & "\settings\cpt-save-local.adtg"
  blnExists = Dir(strSavedMap) <> vbNullString
  If blnExists Then
    Set rstSavedMap = CreateObject("ADODB.Recordset")
    rstSavedMap.Open strSavedMap
  End If
  
  'populate the form - defaults to task ECFs, text
  With mySaveLocal_frm
    'populate map
    .lboECF.Clear
    If rstECF.RecordCount = 0 Then
      rstECF.Close
      MsgBox "No Enterprise Custom Fields available in this file.", vbExclamation + vbOKOnly, "No ECFs found"
      GoTo exit_here
    End If
    'apply
    If .cboECF <> "All Types" Then
      If .cboECF = "undetermined" Then
        rstECF.Filter = "ENTITY='undetermined' OR ENTITY LIKE '%Maybe%'"
      Else
        rstECF.Filter = "ENTITY='" & .cboECF & "' OR ENTITY='undetermined'"
      End If
    End If
    If rstECF.RecordCount = 0 Then GoTo no_records
    rstECF.MoveFirst
    Do While Not rstECF.EOF
      If UCase(rstECF("GUID")) = UCase(strGUID) And rstECF("pjType") = IIf(.optTasks, pjTask, pjResource) Then
        If strFilter <> "" Then
          If InStr(UCase(rstECF("ECF_Name")), UCase(strFilter)) = 0 Then GoTo next_field
        End If
        .lboECF.AddItem
        .lboECF.List(.lboECF.ListCount - 1, 0) = rstECF("ECF")
        .lboECF.List(.lboECF.ListCount - 1, 1) = rstECF("ECF_Name")
        .lboECF.List(.lboECF.ListCount - 1, 2) = rstECF("ENTITY")
        If blnExists Then
          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & rstECF("ECF") '& " AND ENTITY=" & IIf(.optTasks, pjTask, pjResource)
          If Not rstSavedMap.EOF Then
            .lboECF.List(.lboECF.ListCount - 1, 3) = rstSavedMap("LCF")
            If Len(CustomFieldGetName(rstSavedMap("LCF"))) > 0 Then
              .lboECF.List(.lboECF.ListCount - 1, 4) = CustomFieldGetName(rstSavedMap("LCF"))
            Else
              .lboECF.List(.lboECF.ListCount - 1, 4) = FieldConstantToFieldName(rstSavedMap("LCF"))
            End If
          End If
          rstSavedMap.Filter = ""
        End If
      End If
next_field:
      rstECF.MoveNext
    Loop
no_records:
    .lblStatus.Caption = Format(.lboECF.ListCount, "#,##0") & " enterprise custom field(s)."

  End With

exit_here:
  On Error Resume Next
  If rstECF.State Then rstECF.Close
  Set rstECF = Nothing
  If rstSavedMap.State Then rstSavedMap.Close
  Set rstSavedMap = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptUpdateECF", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateLCF(ByRef mySaveLocal_frm As cptSaveLocal_frm, Optional strFilter As String)
  'objects
  'strings
  Dim strCustomName As String
  Dim strFieldName As String
  'longs
  Dim lngFieldID As Long
  Dim lngFields As Long
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  With mySaveLocal_frm
    .lboLCF.Clear
    lngFields = .cboLCF.Column(1)
    For lngField = 1 To lngFields
      strFieldName = .cboLCF.Column(0) & lngField
      lngFieldID = FieldNameToFieldConstant(strFieldName, IIf(.optTasks, pjTask, pjResource))
      strCustomName = CustomFieldGetName(FieldNameToFieldConstant(.cboLCF.Column(0) & lngField, IIf(.optTasks, pjTask, pjResource)))
      If strFilter <> "" Then
        If InStr(UCase(strCustomName), UCase(strFilter)) = 0 Or Len(strCustomName) = 0 Then GoTo next_field
      End If
      If Len(strCustomName) > 0 Then
        .lboLCF.AddItem
        .lboLCF.List(.lboLCF.ListCount - 1, 0) = lngFieldID
        .lboLCF.List(.lboLCF.ListCount - 1, 1) = strFieldName & " (" & CustomFieldGetName(lngFieldID) & ")" 'Me.lboLCF.List(Me.lboLCF.ListCount - 1, 0) = CustomFieldGetName(FieldNameToFieldConstant(Me.cboLCF.Column(0) & lngField))
      Else
        .lboLCF.AddItem
        .lboLCF.List(.lboLCF.ListCount - 1, 0) = lngFieldID
        .lboLCF.List(.lboLCF.ListCount - 1, 1) = strFieldName
      End If
next_field:
    Next lngField
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cboLCF_Change", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateSaveLocalView(ByRef mySaveLocal_frm As cptSaveLocal_frm, Optional lngECF As Long, Optional lngLCF As Long)
  'objects
  'strings
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  With mySaveLocal_frm
    If Not .Visible Then
      'create/overwrite the tables
      cptSpeed True
      ViewApply "Gantt Chart"
      TableEditEx ".cptSaveLocal Task Table", True, True, True, , "ID", , , , , , True, , , , , , , , False
      TableEditEx ".cptSaveLocal Task Table", True, False, , , , "Unique ID", "UID", , , , True
      TableEditEx ".cptSaveLocal Resource Table", False, True, True, , "ID", , , , , , True, , , , , , , , False
      TableEditEx ".cptSaveLocal Resource Table", False, False, , , , "Unique ID", "UID", , , , True
      If ActiveProject.Subprojects.Count > 0 Then
        TableEditEx ".cptSaveLocal Task Table", True, False, , , , "Project", , 40, , , True
        TableEditEx ".cptSaveLocal Resource Table", False, False, , , , "Enterprise", , 15, pjCenter, , True
      End If
      TableEditEx ".cptSaveLocal Resource Table", False, False, , , , "Name", , 30, , , True
      On Error Resume Next
      ActiveProject.Views(".cptSaveLocal Task View").Delete
      ActiveProject.Views(".cptSaveLocal Resource View").Delete
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      ViewEditSingle ".cptSaveLocal Task View", True, , pjTaskSheet, , , ".cptSaveLocal Task Table", "All Tasks", "No Group"
      ViewEditSingle ".cptSaveLocal Resource View", True, , pjResourceSheet, , , ".cptSaveLocal Resource Table", "All Resources", "No Group"
      'update the table
      For lngField = 0 To .lboECF.ListCount - 1
        If Not IsNull(.lboECF.List(lngField, 3)) Then
          lngECF = .lboECF.List(lngField, 0)
          lngLCF = .lboECF.List(lngField, 3)
          TableEditEx ".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngECF), , , , , True, , , , , , , , False
          TableEditEx ".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngLCF), , , , , True, , , , , , , , False
        End If
      Next lngField
      ViewApply ".cptSaveLocal Task View"
      cptSpeed False
    Else
      If .optTasks Then
        If ActiveProject.CurrentView <> ".cptSaveLocal Task View" Then ViewApply ".cptSaveLocal Task View"
        If lngECF > 0 And lngLCF > 0 Then
          If Not TableEditEx(".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngECF), , , , , True, , , , , , , , False) Then
            MsgBox "Failed to add column " & FieldConstantToFieldName(lngECF) & "!", vbExclamation + vbOKOnly, "Fail"
          End If
          If Not TableEditEx(".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngLCF), , , , , True, , , , , , , , False) Then
            MsgBox "Failed to add column " & FieldConstantToFieldName(lngECF) & "!", vbExclamation + vbOKOnly, "Fail"
          End If
          TableApply ".cptSaveLocal Task Table"
        End If
      ElseIf .optResources Then
        If ActiveProject.CurrentView <> ".cptSaveLocal Resource View" Then ViewApply ".cptSaveLocal Resource View"
        If lngECF > 0 And lngLCF > 0 Then
          If Not TableEditEx(".cptSaveLocal Resource Table", False, False, False, , , FieldConstantToFieldName(lngECF), , , , , True, , , , , , , , False) Then
            MsgBox "Failed to add column " & FieldConstantToFieldName(lngECF) & "!", vbExclamation + vbOKOnly, "Fail"
          End If
          If Not TableEditEx(".cptSaveLocal Resource Table", False, False, False, , , FieldConstantToFieldName(lngLCF), , , , , True, , , , , , , , False) Then
            MsgBox "Failed to add column " & FieldConstantToFieldName(lngECF) & "!", vbExclamation + vbOKOnly, "Fail"
          End If
          TableApply ".cptSaveLocal Resource Table"
        End If
      End If
    End If
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptUpdateSaveLocalView", Err, Erl)
  Resume exit_here
End Sub

Function cptLocalCustomFieldsMatch() As Boolean
  'objects
  Dim oRange As Excel.Range
  Dim oListObject As Excel.ListObject
  Dim dTypes As Scripting.Dictionary
  Dim oMasterProject As MSProject.Project
  Dim oSubProject As MSProject.SubProject
  Dim rstProjects As ADODB.Recordset
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  'strings
  Dim strFormula As String
  'longs
  Dim lngCol As Long
  Dim lngMismatchCount As Long
  Dim lngLCF As Long
  Dim lngProject As Long
  Dim lngField As Long
  Dim lngType As Long
  Dim lngLastRow As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vLine As Variant
  Dim vType As Variant
  Dim vEntity As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If ActiveProject.Subprojects.Count = 0 Then
    cptLocalCustomFieldsMatch = True
    GoTo exit_here
  End If
  
  'setup array of types/counts
  Set dTypes = CreateObject("Scripting.Dictionary")
  'record: field type, number of available custom fields
  For Each vType In Array("Cost", "Date", "Duration", "Finish", "Start", "Outline Code")
    dTypes.Add vType, 10
  Next
  dTypes.Add "Flag", 20
  dTypes.Add "Number", 20
  dTypes.Add "Text", 30
  
  Set oMasterProject = ActiveProject
  Application.StatusBar = "Setting up Excel..."
  'set up Excel
  Set oExcel = CreateObject("Excel.Application")
  oExcel.Visible = False
  Set oWorkbook = oExcel.Workbooks.Add
  oExcel.ScreenUpdating = False
  oExcel.Calculation = xlCalculationManual
  Set oWorksheet = oWorkbook.Sheets(1)
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 4
  oExcel.ActiveWindow.FreezePanes = True
  oWorksheet.Name = "Sync"
  'set up headers
  oWorksheet.[A1:D1] = Array("ENTITY", "TYPE", "CONSTANT", "NAME")
  'capture master and subproject names
  oWorksheet.Cells(1, 5) = oMasterProject.Name
  oWorksheet.Columns.AutoFit
  cptSpeed True
  Application.StatusBar = "Opening subprojects..."
  Set rstProjects = CreateObject("ADODB.Recordset")
  rstProjects.Fields.Append "PROJECT", adVarChar, 200
  rstProjects.Open
  rstProjects.AddNew Array(0), Array(oMasterProject.Name)
  For Each oSubProject In oMasterProject.Subprojects
    FileOpenEx oSubProject.SourceProject.FullName, True
    rstProjects.AddNew Array(0), Array(ActiveProject.Name)
  Next oSubProject
  rstProjects.MoveFirst
  Do While Not rstProjects.EOF
    Application.StatusBar = "Analyzing " & rstProjects(0) & "..."
    DoEvents
    lngLastRow = 1
    Projects(CStr(rstProjects(0))).Activate
    oWorksheet.Cells(1, 5 + CLng(rstProjects.AbsolutePosition) - 1) = rstProjects(0)
    For Each vEntity In Array(pjTask, pjResource)
      For lngType = 0 To dTypes.Count - 1
        For lngField = 1 To dTypes.Items(lngType)
          lngLastRow = lngLastRow + 1
          If lngProject = 0 Then
            oWorksheet.Cells(lngLastRow, 1) = Choose(vEntity + 1, "Task", "Resource")
            oWorksheet.Cells(lngLastRow, 2) = dTypes.Keys(lngType)
          End If
          lngLCF = FieldNameToFieldConstant(dTypes.Keys(lngType) & lngField, vEntity)
          If lngProject = 0 Then
            oWorksheet.Cells(lngLastRow, 3) = lngLCF
            oWorksheet.Cells(lngLastRow, 4) = FieldConstantToFieldName(lngLCF)
          End If
          oExcel.ActiveWindow.ScrollRow = lngLastRow
          lngCol = 5 + CLng(rstProjects.AbsolutePosition) - 1
          oWorksheet.Cells(lngLastRow, lngCol).Value = CustomFieldGetName(lngLCF)
          If Len(CustomFieldGetName(lngLCF)) > 0 Then
            strFormula = CustomFieldGetFormula(lngLCF)
            If Len(strFormula) > 0 Then
              On Error Resume Next
              Dim oComment As Excel.Comment
              Set oComment = oWorksheet.Cells(lngLastRow, lngCol).Comment
              If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
              If oComment Is Nothing Then Set oComment = oWorksheet.Cells(lngLastRow, lngCol).AddComment("<formula>" & strFormula & "</formula>" & vbCrLf)
              oComment.Shape.TextFrame.Characters.Font.Bold = False
              oComment.Shape.TextFrame.Characters.Font.Name = "Consolas"
              oComment.Shape.TextFrame.Characters.Font.Size = 11
              oComment.Shape.TextFrame.AutoSize = True
              Set oComment = Nothing
            End If
            'lookuptables
            
          End If
          oWorksheet.Cells.Columns.AutoFit
        Next lngField
      Next lngType
    Next vEntity
    rstProjects.MoveNext
  Loop
  
  'todo: this should populate an ADODB recordset and just run queries against it
  'entity(task,resource);type(text,cost...);constant;name;custom_name;formula;pick_list;otherproperties?
  
  'todo: compare formulae
  Dim lngRow As Long
  lngRow = oWorksheet.Comments(1).Parent.Row
  oWorksheet.Cells(lngRow, lngCol + 2) = oWorksheet.Comments(1).Text
  For Each oComment In oWorksheet.Comments
    If oComment.Parent.Row = lngRow Then
      If oComment.Text <> oWorksheet.Cells(lngRow, lngCol + 2) Then
        oWorksheet.Cells(lngRow, lngCol + 2).Value = "MISMATCHED FORMULA"
      End If
    Else
      lngRow = oWorksheet.Comments(1).Parent.Row
      oWorksheet.Cells(lngRow, lngCol + 2) = oWorksheet.Comments(1).Text
    End If
  Next
  
  'add a formula
  oWorksheet.Cells(1, 5 + rstProjects.RecordCount) = "MATCH"
  oWorksheet.Range(oWorksheet.Cells(2, 5 + rstProjects.RecordCount), oWorksheet.Cells(lngLastRow, 5 + rstProjects.RecordCount)).FormulaR1C1 = "=AND(EXACT(RC[-5],RC[-4]),EXACT(RC[-4],RC[-3]),EXACT(RC[-3],RC[-2]),EXACT(RC[-2],RC[-1]))"
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)), , xlYes)
  oListObject.TableStyle = ""
  oListObject.HeaderRowRange.Font.Bold = True
  'throw some shade
  With oListObject.HeaderRowRange.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = -0.149998474074526
    .PatternTintAndShade = 0
  End With

  oListObject.Range.Borders(xlDiagonalDown).LineStyle = xlNone
  oListObject.Range.Borders(xlDiagonalUp).LineStyle = xlNone
  For Each vLine In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
    With oListObject.Range.Borders(vLine)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.249946592608417
      .Weight = xlThin
    End With
  Next vLine
  oExcel.Calculation = xlCalculationAutomatic
  'add conditional formatting
  Set oRange = oListObject.ListColumns("MATCH").DataBodyRange
  oRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE"
  oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
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
  oRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=TRUE"
  oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
  With oRange.FormatConditions(1).Font
      .Color = -16752384
      .TintAndShade = 0
  End With
  With oRange.FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 13561798
      .TintAndShade = 0
  End With
  oRange.FormatConditions(1).StopIfTrue = False
  'autofilter it
  oListObject.Range.AutoFilter oRange.Column, False
  oWorksheet.Columns.AutoFit
  oExcel.ActiveWindow.ScrollRow = 1
  oMasterProject.Activate
  rstProjects.MoveFirst
  Do While Not rstProjects.EOF
    If CStr(rstProjects(0)) <> oMasterProject.Name Then
      Projects(CStr(rstProjects(0))).Activate
      Application.FileCloseEx pjDoNotSave
    End If
    rstProjects.MoveNext
  Loop
  cptSpeed False
  oExcel.ScreenUpdating = True
  oExcel.Visible = True
  oExcel.WindowState = xlMaximized
  lngMismatchCount = oRange.SpecialCells(xlCellTypeVisible).Count
  If lngMismatchCount > 0 Then
    oExcel.ActivateMicrosoftApp xlMicrosoftProject
    MsgBox lngMismatchCount & " Local Custom Fields do not match between Master and all Subprojects!", vbCritical + vbOKOnly, "Warning"
    Application.ActivateMicrosoftApp pjMicrosoftExcel
    GoTo exit_here
    cptLocalCustomFieldsMatch = False
  Else
    oWorkbook.Close False
    oExcel.Quit
    cptLocalCustomFieldsMatch = True
  End If

exit_here:
  On Error Resume Next
  Set oRange = Nothing
  Set oListObject = Nothing
  Set dTypes = Nothing
  Set oMasterProject = Nothing
  Set oSubProject = Nothing
  If rstProjects.State Then rstProjects.Close
  Set rstProjects = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptLocalCustomFieldsMatch", Err, Erl)
  Resume exit_here
End Function
