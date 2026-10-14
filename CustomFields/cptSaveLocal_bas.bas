Attribute VB_Name = "cptSaveLocal_bas"
'<cpt_version>v1.4.0</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptSaveLocal_bas"
Private Const TABLE_NAME As String = ".cptSaveLocal Table"
Private Const VIEW_NAME As String = ".cptSaveLocal View"
Public strStartView As String
Public strStartTable As String
Public strStartFilter As String
Public strStartGroup As String

Sub cptShowSaveLocal_frm()
  'objects
  Dim oSavedMap As ADODB.Recordset
  Dim mySaveLocal_frm As cptSaveLocal_frm
  Dim oTask As MSProject.Task
  Dim dTypes As Scripting.Dictionary
  Dim rstECF As ADODB.Recordset
  'strings
  Dim strSetting As String
  Dim strDir As String
  Dim strSavedMap As String
  Dim strServerURL As String
  Dim strDataType As String
  Dim strProgramAcronym As String
  Dim strECF As String
  Dim strFileName As String
  'longs
  Dim lngFile As Long
  Dim lngECF As Long
  Dim lngLCF As Long
  Dim lngSubprojectCount As Long
  Dim lngField As Long
  Dim lngType As Long
  Dim lngECFCount As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vType As Variant
  'dates

  'prevent spawning
  If Not cptGetUserForm("cptSaveLocal_frm") Is Nothing Then Exit Sub
  
  'check for an update
  'If Not cptProceedOnUpdate(THIS_MODULE) Then GoTo exit_here

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Projects.Count = 0 Then GoTo exit_here
  
  'get server URL
  strServerURL = Application.Profiles.ActiveProfile.Server
  If Len(strServerURL) = 0 Then
    MsgBox "You are not connected to a Project Server.", vbCritical + vbOKOnly, "ECF to LCF"
    GoTo exit_here
  End If
  
  'get program acronym
  strProgramAcronym = cptGetProgramAcronym
  If Len(strProgramAcronym) = 0 Then
    MsgBox "Program Acronym is required. Exiting.", vbCritical + vbOKOnly, "ECF to LCF"
    GoTo exit_here
  End If
  
  cptSpeed True
  
  'check for saved map
  strDir = cptDir
  strSavedMap = strDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) <> vbNullString Then
    'warn about upgrade, kick out existing
    Set oSavedMap = CreateObject("ADODB.Recordset")
    oSavedMap.Open strSavedMap
    If oSavedMap.RecordCount > 0 Then
      If oSavedMap.Fields(1).Name = "GUID" Then
        strFileName = Environ("USERPROFILE") & "\cpt-ecf-to-lcf-old-map.txt"
        lngFile = FreeFile
        MsgBox "Database Structure has changed for saved mapping." & vbCrLf & "Please re-map your fields." & vbCrLf & "Existing map will be exported for reference.", vbCritical + vbOKOnly, "ECF to LCF"
        Open strFileName For Output As #lngFile
        Print #lngFile, "SERVER: [FROM ECF] >> [TO LCF]"
        oSavedMap.MoveFirst
        Do While Not oSavedMap.EOF
          Print #lngFile, oSavedMap(0) & ": [" & FieldConstantToFieldName(oSavedMap(2)) & "] >> [" & FieldConstantToFieldName(oSavedMap(3)) & "]"
          oSavedMap.MoveNext
        Loop
        Close #lngFile
        cptShellExecute 0, "open", "notepad.exe", strFileName, vbNullString, 1
        oSavedMap.Close
        Kill strSavedMap
      End If
    End If
  Else 'create it
    Set oSavedMap = CreateObject("ADODB.Recordset")
    With oSavedMap
      .Fields.Append "URL", adVarChar, 255
      .Fields.Append "PROGRAM_ACRONYM", adVarChar, 255
      .Fields.Append "ECF", adInteger
      .Fields.Append "LCF", adInteger
      .Open
      .Save strSavedMap, adPersistADTG
      .Close
    End With
  End If
  
  'if master/sub then ensure LCFs match
  lngSubprojectCount = ActiveProject.Subprojects.Count
  If lngSubprojectCount > 0 Then
    If MsgBox(lngSubprojectCount & " subproject(s) found." & vbCrLf & vbCrLf & "It is highly recommended that you analyze master/sub LCF matches." & vbCrLf & vbCrLf & "Do it now?", vbExclamation + vbYesNo, "Master/Sub Detected") = vbYes Then
      cptAlignLCF ActiveProject
    End If
  End If
  
  'capture starting view/table/filter/group
  ActiveWindow.TopPane.Activate
  strStartView = ActiveProject.CurrentView
  strStartTable = ActiveProject.CurrentTable
  strStartFilter = ActiveProject.CurrentFilter
  strStartGroup = ActiveProject.CurrentGroup
    
  'prepare to capture all ECFs
  Set rstECF = CreateObject("ADODB.Recordset")
  rstECF.Fields.Append "URL", adVarChar, 255
  rstECF.Fields.Append "PROGRAM_ACRONYM", adVarChar, 120
  rstECF.Fields.Append "DATA_TYPE", adVarChar, 50
  rstECF.Fields.Append "ECF", adInteger
  rstECF.Fields.Append "ECF_Name", adVarChar, 120
  rstECF.Fields.Append "LCF", adInteger
  rstECF.Fields.Append "LCF_Name", adVarChar, 120
  rstECF.Open
  
  'create a dummy task to interrogate the ECFs
  Set oTask = ActiveProject.Tasks.Add("<dummy for cpt-save-local>")
  Application.CalculateProject
  
  'setup array of types/counts
  Set dTypes = CreateObject("Scripting.Dictionary")
  'record: field type, number of available custom fields
  For Each vType In Array("Cost", "Date", "Duration", "Finish", "Start", "Outline Code")
    dTypes.Add vType, 10
  Next
  dTypes.Add "Flag", 20
  dTypes.Add "Number", 20
  dTypes.Add "Text", 30
  
  'populate field types
  Set mySaveLocal_frm = New cptSaveLocal_frm
  With mySaveLocal_frm
    .Caption = "Save Local (" & cptGetVersion("cptSaveLocal_frm") & ")"
    .txtServerURL = strServerURL
    .lboECF.Clear
    .lboLCF.Clear
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
    strSetting = cptGetSetting("SaveLocal", "txtRenamingConvention")
    If Len(strSetting) > 0 Then
      .txtRenamingConvention = strSetting
    Else
      .txtRenamingConvention = "[ECF] ([LCF])" 'default
    End If
    
    'cptUpdateECF mySaveLocal_frm
    'cptUpdateLCF mySaveLocal_frm
    'cptUpdateSaveLocalView mySaveLocal_frm
    
    'get enterprise custom task fields
    For lngField = 188776000 To 188778000 '2000 should do it for now
      If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
        strECF = FieldConstantToFieldName(lngField)
        strDataType = cptInterrogateECF(oTask, lngField)
        rstECF.AddNew Array("URL", "PROGRAM_ACRONYM", "DATA_TYPE", "ECF", "ECF_Name"), Array(strServerURL, strProgramAcronym, strDataType, lngField, FieldConstantToFieldName(lngField))
        lngECFCount = lngECFCount + 1
      End If
      .lblStatus.Caption = "Analyzing Task ECFs...(" & Format(((lngField - 188776000) / (188778000 - 188776000)), "0%") & ")"
      .lblProgress.Width = ((lngField - 188776000) / (188778000 - 188776000)) * .lblStatus.Width
      DoEvents
    Next lngField
  
    oTask.Delete
  
    If Dir(strDir & "\settings\cpt-ecf.adtg") <> vbNullString Then
      Kill strDir & "\settings\cpt-ecf.adtg"
    End If
    rstECF.Sort = "ECF_Name"
    rstECF.Save strDir & "\settings\cpt-ecf.adtg"
    rstECF.Close
  
    'trigger lboECF refresh
    .cboECF.Value = "All Types" 'this loads saved map
    
    'update the table
    cptUpdateSaveLocalView mySaveLocal_frm
    
    DoEvents
    
    cptSpeed False
    .Show 'modal to control changes to custom fields
  End With

exit_here:
  On Error Resume Next
  Set mySaveLocal_frm = Nothing
  Application.StatusBar = ""
  cptSpeed False
  oTask.Delete
  Set oTask = Nothing
  Set vType = Nothing
  Set dTypes = Nothing
  If rstECF.State Then rstECF.Close
  Set rstECF = Nothing
  If oSavedMap.State Then oSavedMap.Close
  Set oSavedMap = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptShowSaveLocal_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptSaveLocal(ByRef mySaveLocal_frm As cptSaveLocal_frm)
  'objects
  Dim oSavedMap As ADODB.Recordset
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  'strings
  Dim strExportMapName As String
  Dim strImportMapName As String
  Dim strErrors As String
  Dim strFileName As String
  Dim strSavedMap As String
  Dim strMsg As String
  Dim strServerURL As String
  Dim strProgramAcronym As String
  Dim strLCF As String
  Dim strECF As String
  'longs
  Dim lngMapped As Long
  Dim lngMappedCount As Long
  Dim lngFile As Long
  Dim lngItems As Long
  Dim lngLCF As Long
  Dim lngECF As Long
  Dim lngItem As Long
  Dim lngTask As Long
  Dim lngTaskCount As Long
  'integers
  'doubles
  'booleans
  Dim blnMaster As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vMap As Variant
  'dates
  Dim tmr As Date
  
  tmr = Now
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnMaster = ActiveProject.Subprojects.Count > 0
  
  strServerURL = Application.Profiles.ActiveProfile.Server
    
  'save map
  Set oSavedMap = CreateObject("ADODB.Recordset")
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) = vbNullString Then 'create it
    oSavedMap.Fields.Append "URL", adVarChar, 255
    oSavedMap.Fields.Append "PROGRAM_ACRONYM", adVarChar, 255 'was "GUID", adGUID
    oSavedMap.Fields.Append "ECF", adBigInt
    oSavedMap.Fields.Append "LCF", adBigInt
    oSavedMap.Open
  Else
    oSavedMap.Open strSavedMap
  End If
  'get total task count
  ActiveWindow.TopPane.Activate
  If ActiveProject.CurrentView <> ".cptSaveLocal Task View" Then
    ViewApply ".cptSaveLocal Task View"
  End If
  FilterClear
  GroupClear

  On Error Resume Next
  If Not OutlineShowAllTasks Then
    OptionsViewEx DisplaySummaryTasks:=True
    OutlineShowAllTasks
    Sort "ID", , , , , , False, True
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  SelectAll
  Set oTasks = ActiveSelection.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oTasks Is Nothing Then
    lngTaskCount = oTasks.Count
  Else
    MsgBox "There are no tasks in this schedule.", vbCritical + vbOKOnly, "No Tasks"
    GoTo exit_here
  End If
  
  cptSpeed True 'speed it up
  
  OptionsViewEx ProjectSummary:=False, DisplayNameIndent:=False, DisplaySummaryTasks:=False, DisplayExternalSuccessors:=False, DisplayExternalPredecessors:=False
  
  With mySaveLocal_frm
    'get count of mapped
    lngItems = .lboECF.ListCount - 1
    For lngItem = 0 To lngItems
      If Len(.lboECF.List(lngItem, 3)) > 0 Then
        lngMappedCount = lngMappedCount + 1
      End If
    Next lngItem
    If blnMaster Then
      For lngItem = 0 To lngItems
        If .lboECF.List(lngItem, 3) > 0 Then
          lngECF = .lboECF.List(lngItem, 0)
          lngLCF = .lboECF.List(lngItem, 3)
          'if ECF is formula, then skip it - already built it with cmdMap
          If Len(CustomFieldGetFormula(lngECF)) > 0 Then GoTo next_mapped_master
          FilterClear
          SelectAll
          SetField FieldConstantToFieldName(lngLCF), vbNullString
          'todo: filter out external?
          SetAutoFilter FieldConstantToFieldName(lngECF), pjAutoFilterCustom, "does not equal", ""
          SelectAll
          On Error Resume Next
          Set oTasks = ActiveSelection.Tasks
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oTasks Is Nothing Then
            'error - ECF had no values?
            GoTo next_mapped_master
          End If
          lngTask = 0
          On Error Resume Next
          lngTaskCount = oTasks.Count
          For Each oTask In oTasks
            If oTask Is Nothing Then GoTo next_task_master
            If oTask.ExternalTask Then GoTo next_task_master
            strECF = CStr(oTask.GetField(lngECF))
            On Error Resume Next
            oTask.SetField lngLCF, strECF
            If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
            strLCF = CStr(oTask.GetField(lngLCF))
            If strLCF <> strECF Then
              Find "Unique ID", "equals", oTask.UniqueID
              strErrors = strErrors & oTask.Project & "," & oTask.UniqueID & "," & CustomFieldGetName(lngECF) & "," & CustomFieldGetName(lngLCF) & "," & FieldConstantToFieldName(lngLCF) & vbCrLf
              If MsgBox("PROJECT: " & oTask.Project & vbCrLf & "TASK: " & oTask.UniqueID & vbCrLf & vbCrLf & "Error copying from ECF '" & CustomFieldGetName(lngECF) & "' to LCF '" & CustomFieldGetName(lngLCF) & "' (" & FieldConstantToFieldName(lngLCF) & ")" & vbCrLf & vbCrLf & "Please validate data type mapping." & vbCrLf & vbCrLf & "Proceed anyway?" & vbCrLf & vbCrLf & "Yes: Skip this LCF and keep going" & vbCrLf & "No: Exit this process", vbExclamation + vbYesNo, "Failed!") = vbYes Then
                GoTo next_mapped_master
              Else
                GoTo exit_here
              End If
            End If
next_task_master:
            lngTask = lngTask + 1
            Application.StatusBar = "Processing '" & CustomFieldGetName(lngECF) & "'... " & Format(lngMapped, "#,##0") & " / " & Format(lngMappedCount, "#,##0") & " (" & Format(lngMapped / lngMappedCount, "0%") & ") | Task " & Format(lngTask, "#,##0") & " / " & Format(lngTaskCount, "#,##0") & " (" & Format(lngTask / lngTaskCount, "0%") & ") | " & Format(Now - tmr, "hh:nn:ss")
            DoEvents
          Next oTask
          lngMapped = lngMapped + 1
          
  '        'no duplicates - doesn't this happen real-time via form controls?
  '        oSavedMap.Filter = "PROGRAM_ACRONYM='" & strProgramAcronym & "' AND ECF=" & lngECF
  '        If oSavedMap.RecordCount = 1 Then 'overwrite it
  '          oSavedMap.Delete adAffectCurrent
  '        End If
  '        oSavedMap.Filter = ""
  '        oSavedMap.Filter = "PROGRAM_ACRONYM='" & strProgramAcronym & "' AND LCF=" & lngLCF
  '        If oSavedMap.RecordCount = 1 Then 'overwrite it
  '          oSavedMap.Delete adAffectCurrent
  '        End If
  '        oSavedMap.Filter = ""
  '        'add the new record
  '        oSavedMap.AddNew Array(0, 1, 2, 3), Array(strServerURL, strProgramAcronym, lngECF, lngLCF)
        End If
next_mapped_master:
        Application.StatusBar = "Processing... " & Format(lngMapped, "#,##0") & " / " & Format(lngMappedCount, "#,##0") & " (" & Format(lngMapped / lngMappedCount, "0%") & ") | " & Format(Now - tmr, "hh:nn:ss")
        DoEvents
      Next lngItem
      FilterClear
    Else
      'create map for export
      strExportMapName = "cpt-ecf-to-lcf-export-map"
      Application.MapEdit Name:=strExportMapName, Create:=True, OverwriteExisting:=True, _
                          DataCategory:=pjMapTasks, CategoryEnabled:=True, TableName:="Entry", _
                          FieldName:="Unique ID", ExternalFieldName:="Unique ID", _
                          ExportFilter:="All Tasks", MergeKey:="Unique ID", ImportMethod:=pjImportMerge, _
                          HeaderRow:=True, AssignmentData:=False, TextDelimiter:=Chr$(44), TextFileOrigin:=0, _
                          UseHtmlTemplate:=False, IncludeImage:=False
      
      'create map for import
      strImportMapName = "cpt-ecf-to-lcf-import-map"
      Application.MapEdit Name:=strImportMapName, Create:=True, OverwriteExisting:=True, _
                          DataCategory:=pjMapTasks, CategoryEnabled:=True, TableName:="Entry", _
                          FieldName:="Unique ID", ExternalFieldName:="Unique ID", _
                          ExportFilter:="All Tasks", MergeKey:="Unique ID", ImportMethod:=pjImportMerge, _
                          HeaderRow:=True, AssignmentData:=False, TextDelimiter:=Chr$(44), TextFileOrigin:=0, _
                          UseHtmlTemplate:=False, IncludeImage:=False
      
      'loop through map and add
      oSavedMap.Filter = 0
      oSavedMap.Filter = "PROGRAM_ACRONYM='" & strProgramAcronym & "'"
      If oSavedMap.EOF Then GoTo exit_here
      oSavedMap.MoveFirst
      Do While Not oSavedMap.EOF
        strECF = FieldConstantToFieldName(oSavedMap(2))
        strLCF = FieldConstantToFieldName(oSavedMap(3))
        Application.MapEdit Name:=strExportMapName, OverwriteExisting:=True, DataCategory:=pjMapTasks, CategoryEnabled:=True, TableName:="Entry", FieldName:=strECF, ExternalFieldName:=strLCF, ExportFilter:="All Tasks", MergeKey:="Unique ID", ImportMethod:=pjImportMerge, HeaderRow:=True, AssignmentData:=False, TextDelimiter:=Chr$(44), TextFileOrigin:=0, UseHtmlTemplate:=False, IncludeImage:=False
        Application.MapEdit Name:=strImportMapName, OverwriteExisting:=True, DataCategory:=pjMapTasks, CategoryEnabled:=True, TableName:="Entry", FieldName:=strLCF, ExternalFieldName:=strLCF, ExportFilter:="All Tasks", ImportMethod:=pjImportMerge, MergeKey:="Unique ID", HeaderRow:=True, AssignmentData:=False, TextDelimiter:=Chr$(44), TextFileOrigin:=0, UseHtmlTemplate:=False, IncludeImage:=False
        oSavedMap.MoveNext
      Loop
      'export to csv
      strFileName = Environ("TMP") & "\cpt-ecf-to-lcf.csv"
      ActiveProject.SaveAs strFileName, Format:=pjCSV, FormatID:="MSProject.csv", Map:=strExportMapName
      
      'import the csv
      FileOpenEx Name:=strFileName, ReadOnly:=False, Merge:=pjMerge, Map:=strImportMapName
      
      'delete the maps
      For Each vMap In Array(strExportMapName, strImportMapName)
        If Not OrganizerDeleteItem(Type:=pjMaps, FileName:="Global (+ non-cached Enterprise)", Name:=CStr(vMap)) Then
          MsgBox "Could not delete the map '" & vMap & "' - please delete manually using the Organizer.", vbExclamation + vbOKOnly, "ECF to LCF"
        End If
      Next vMap
      
      'delete the csv
      Kill strFileName
      
    End If
    mySaveLocal_frm.lblStatus.Caption = "Saving ECF to LCF...done."
    mySaveLocal_frm.lblProgress.Width = mySaveLocal_frm.lblStatus.Width
    DoEvents
    
  End With

  oSavedMap.Filter = 0
  oSavedMap.Save strSavedMap, adPersistADTG

  MsgBox "Enteprise Custom Fields saved locally." & vbCrLf & vbCrLf & "Time: " & Format(Now - tmr, "hh:nn:ss"), vbInformation + vbOKOnly, "Complete"
  mySaveLocal_frm.lblStatus.Caption = "Ready..."

exit_here:
  On Error Resume Next
  cptSpeed False
  If Len(strErrors) > 0 Then
    lngFile = FreeFile
    strFileName = Environ("TMP") & "\cpt-ecf-to-lcf-errors.txt"
    Open strFileName For Output As #lngFile
    Print #lngFile, "PROJECT,TASK_UID,ECF,LCF,FIELD_NAME"
    Print #lngFile, strErrors
    Close #lngFile
    cptShellExecute 0, "open", "notepad.exe", strFileName, vbNullString, 1
  End If
  Set oTasks = Nothing
  If oSavedMap.State Then oSavedMap.Close
  Set oSavedMap = Nothing
  Set oTask = Nothing
  Set oTasks = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptSaveLocal", Err, Erl)
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
  Call cptHandleErr(THIS_MODULE, "cptInterrogateECF", Err, Erl)
  Resume exit_here
End Function

Sub cptAnalyzeAutoMap(ByRef mySaveLocal_frm As cptSaveLocal_frm)
  'objects
  Dim rstAvailable As ADODB.Recordset
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
    Next vType
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
          If IsNull(mySaveLocal_frm.lboECF.List(lngItem, 3)) Then
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
  Call cptHandleErr(THIS_MODULE, "cptAutoMap", Err, Erl)
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
  Call cptHandleErr(THIS_MODULE, "cptAutoMap", Err, Erl)
  Resume exit_here
End Sub

Sub cptMapECFtoLCF(ByRef mySaveLocal_frm As cptSaveLocal_frm, lngECF As Long, lngLCF As Long)
  'objects
  Dim rstSavedMap As ADODB.Recordset
  Dim oLookupTableEntry As LookupTableEntry
  Dim oOutlineCodeLocal As OutlineCode
  Dim oOutlineCode As OutlineCode
  'strings
  Dim strRenamingConvention
  Dim strUnmapped As String
  Dim strFields As String
  Dim strField As String
  Dim strCustomFormula As String
  Dim strServerURL As String
  Dim strProgramAcronym As String
  Dim strSavedMap As String
  Dim strECF As String
  Dim strFN As String
  Dim strCFN As String
  'longs
  Dim lngSourceField As Long
  Dim lngItem As Long
  Dim lngDown As Long
  Dim lngCodeNumber As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnProceed As Boolean
  'variants
  Dim vField As Variant
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'todo: use cptCustomFieldExists
  'todo: ensure not already mapped
  
  With mySaveLocal_frm
    'if already mapped then prompt with ECF name and ask to remap
    For lngItem = 0 To .lboECF.ListCount - 1
      If .lboECF.List(lngItem, 3) = lngLCF Then
        If .lboECF.List(lngItem, 0) = lngECF Then
          'just rename
          Dim blnRenameOnly As Boolean
          blnRenameOnly = True
          GoTo rename_it 'already mapped
        End If
        If MsgBox(FieldConstantToFieldName(lngLCF) & " is already mapped to " & .lboECF.List(lngItem, 1) & " - reassign it?", vbExclamation + vbYesNo, "Already Mapped") = vbYes Then '
          .lboECF.List(lngItem, 3) = ""
          .lboECF.List(lngItem, 4) = ""
        Else
          GoTo exit_here
        End If
      End If
next_ecf:
    Next lngItem
    
    'todo: reset .txtRenamingConvention to normal border color
    'todo: save & load renaming convention setting
    'todo: include renaming convention in export
    If .txtRenamingConvention = "[ECF]" Or Len(.txtRenamingConvention) = 0 Then
      'todo: color border red
      MsgBox "You must define a renaming convention.", vbExclamation + vbOKOnly, "ECF to LCF"
      GoTo exit_here
    End If
    
    If Not CustomFieldDelete(lngLCF) Then
      MsgBox "Could not delete LCF '" & FieldConstantToFieldName(lngLCF) & "'!", vbCritical + vbOKOnly, "Error"
      GoTo exit_here
    End If
    
rename_it:
    
    'capture rename
    strRenamingConvention = Replace(.txtRenamingConvention, "[ECF]", CustomFieldGetName(lngECF)) 'todo: this might need FieldConstantToFieldName() instead - test it
    strRenamingConvention = Replace(strRenamingConvention, "[LCF]", FieldConstantToFieldName(lngLCF))
    If CustomFieldGetName(lngLCF) = strRenamingConvention Then GoTo skip_rename
    strFN = FieldConstantToFieldName(lngLCF)
    strCFN = CustomFieldGetName(lngLCF)
    If Len(strCFN) > 0 Then 'CFN exists confirm overwrite
      If MsgBox("Rename " & strCFN & " (" & strFN & ") to " & strRenamingConvention & "?", vbQuestion + vbYesNo, "Please confirm") = vbYes Then
        'rename it
        CustomFieldRename CLng(lngLCF), strRenamingConvention 'CustomFieldGetName(lngECF) & " (" & FieldConstantToFieldName(lngLCF) & ")"
        'rename in lboLCF
        If Not .tglAutoMap Then
          .lboLCF.List(.lboLCF.ListIndex, 1) = FieldConstantToFieldName(lngLCF) & " (" & CustomFieldGetName(lngLCF) & ")"
          .lboECF.List(.lboECF.ListIndex, 4) = CustomFieldGetName(lngLCF)
        End If
      Else
        GoTo exit_here
      End If
    Else
      ActiveWindow.TopPane.Activate
      'rename it in msp
      CustomFieldRename lngLCF, strRenamingConvention 'CustomFieldGetName(lngECF) & " (" & FieldConstantToFieldName(lngLCF) & ")"
      'rename it in lboLCF
      If Not .tglAutoMap Then
        .lboLCF.List(.lboLCF.ListIndex, 1) = FieldConstantToFieldName(.lboLCF) & " (" & CustomFieldGetName(.lboLCF) & ")"
        .lboECF.Value = lngECF
        .lboECF.List(.lboECF.ListIndex, 4) = CustomFieldGetName(lngLCF)
      End If
    End If
    If blnRenameOnly Then GoTo update_table
    
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
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
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
        'todo: try all together; try serietem
        CustomOutlineCodeEditEx FieldID:=lngLCF, OnlyLookUpTableCodes:=oOutlineCode.OnlyLookUpTableCodes, OnlyCompleteCodes:=oOutlineCode.OnlyCompleteCodes
        CustomOutlineCodeEditEx FieldID:=lngLCF, OnlyLeaves:=oOutlineCode.OnlyLeaves
        'todo: next line for RequiredCode throws an error, not sure why
        'todo: does LCF Outline Code have items? is a calcualtion needed?
        'todo: maybe blnRequiredCode=oOutlineCode.RequiredCode; RequiredCode:=blnRequiredCode?
        'todo: maybe just RequiredCode:=CBool(oOutlineCode.RequiredCode)?
        'todo: maybe use If . = False then warn user
        'CustomOutlineCodeEditEx FieldID:=lngLCF, RequiredCode:=oOutlineCode.RequiredCode
        If oOutlineCode.DefaultValue <> "" Then CustomOutlineCodeEditEx FieldID:=lngLCF, DefaultValue:=oOutlineCode.DefaultValue
        CustomOutlineCodeEditEx FieldID:=lngLCF, SortOrder:=oOutlineCode.SortOrder
      Else 'import just the pick list
        For lngItem = 1 To oOutlineCode.LookupTable.Count
          CustomFieldValueListAdd lngLCF, oOutlineCode.LookupTable(lngItem).Name, oOutlineCode.LookupTable(lngItem).Description
        Next lngItem
      End If
    End If
    If Not .tglAutoMap Then 'why?
      .lboECF.List(.lboECF.ListIndex, 3) = lngLCF
      .lboECF.List(.lboECF.ListIndex, 4) = CustomFieldGetName(lngLCF)
    End If
  End With
  
  'update rstSavedMap
  strProgramAcronym = cptGetProgramAcronym
  strServerURL = Application.Profiles.ActiveProfile.Server
  
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) <> vbNullString Then
    rstSavedMap.Open strSavedMap
    rstSavedMap.Filter = "PROGRAM_ACRONYM='" & strProgramAcronym & "' AND ECF=" & lngECF
    If Not rstSavedMap.EOF Then
      rstSavedMap.Fields(3) = lngLCF
    Else
      rstSavedMap.AddNew Array(0, 1, 2, 3), Array(strServerURL, strProgramAcronym, lngECF, lngLCF)
    End If
    rstSavedMap.Filter = ""
    rstSavedMap.Save strSavedMap, adPersistADTG
  Else 'create it
    rstSavedMap.Fields.Append "URL", adVarChar, 255
    rstSavedMap.Fields.Append "PROGRAM_ACRONYM", adVarChar, 255 'was "GUID", adGUID
    rstSavedMap.Fields.Append "ECF", adInteger
    rstSavedMap.Fields.Append "LCF", adInteger
    rstSavedMap.Open
    rstSavedMap.AddNew Array(0, 1, 2, 3), Array(strServerURL, strProgramAcronym, lngECF, lngLCF)
    rstSavedMap.Save strSavedMap, adPersistADTG
  End If
  rstSavedMap.Close
update_table:
  'update the table
  cptUpdateSaveLocalView mySaveLocal_frm, "map", lngECF & "|" & lngLCF
  
exit_here:
  On Error Resume Next
  Set rstSavedMap = Nothing
  Set oLookupTableEntry = Nothing
  Set oOutlineCodeLocal = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptMapECFtoLCF", Err, Erl)
  Resume exit_here
End Sub

Sub cptDeleteCFMap(mySaveLocal_frm As cptSaveLocal_frm)
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strSavedMap As String
  Dim strDir As String
  Dim strProgramAcronym As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If MsgBox("Would you like to export your map first?" & vbCrLf & vbCrLf & "(This cannot be undone.)", vbQuestion + vbYesNo, "Backup?") = vbYes Then
    cptExportCFMap
  End If
  
  strProgramAcronym = cptGetProgramAcronym
  
  'open recordset, delete records
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strSavedMap
  oRecordset.Filter = "PROGRAM_ACRONYM='" & strProgramAcronym & "'"
  If oRecordset.RecordCount > 0 Then
    oRecordset.MoveFirst
    Do While Not oRecordset.EOF
      oRecordset.Delete adAffectCurrent
      oRecordset.MoveNext
    Loop
    oRecordset.Filter = 0
    oRecordset.Save strSavedMap, adPersistADTG
  End If
  
  'todo: delete the LCF?
  'todo: if so: delete the task data?
  
  With mySaveLocal_frm
    For lngItem = .lboECF.ListCount - 1 To 0 Step -1
      If Len(.lboECF.List(lngItem, 3)) > 0 Then
        CustomFieldDelete .lboECF.List(lngItem, 3)
        .lboECF.List(lngItem, 3) = ""
        .lboECF.List(lngItem, 4) = ""
      End If
    Next lngItem
    'todo: status bar
    cptUpdateLCF mySaveLocal_frm
  End With
  
  cptUpdateSaveLocalView mySaveLocal_frm, blnReset:=True

exit_here:
  On Error Resume Next
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptDeleteCFMap", Err, Erl)
  Resume exit_here
End Sub

Sub cptDeleteLCFAndTaskData(ByRef mySaveLocal_frm As cptSaveLocal_frm)
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  'strings
  Dim strFieldName As String
  Dim strCustomFieldName As String
  'longs
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngLCF As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnDeleteTaskData As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  With mySaveLocal_frm
    If .lboLCF.ListCount = 0 Then GoTo exit_here
    If IsNull(.lboLCF) Then GoTo exit_here
    For lngItem = 0 To .lboLCF.ListCount - 1
      If .lboLCF.Selected(lngItem) Then
        lngLCF = CLng(.lboLCF.List(lngItem, 0))
        strCustomFieldName = CustomFieldGetName(lngLCF)
        If Len(strCustomFieldName) > 0 Then
          strFieldName = strCustomFieldName
        Else
          strFieldName = FieldConstantToFieldName(lngLCF)
        End If
        blnDeleteTaskData = MsgBox("Delete Task Data for LCF '" & strFieldName & "'?", vbQuestion + vbYesNo, "Please Confirm") = vbYes
        If CustomFieldDelete(lngLCF) Then
          .lboLCF.List(lngItem, 1) = FieldConstantToFieldName(lngLCF)
          If blnDeleteTaskData Then
            If ActiveProject.Subprojects.Count > 0 Then
              ActiveWindow.TopPane.Activate
              If ActiveWindow.TopPane.View.Type <> pjTaskItem Then ViewApply "Gantt Chart"
              GroupClear
              FilterClear
              Sort "ID", , , , , , False, True
              OptionsViewEx DisplaySummaryTasks:=True
              OutlineShowTasks pjTaskOutlineShowLevelMax
            End If
            SelectAll
            Set oTasks = ActiveSelection.Tasks
            lngTasks = oTasks.Count
            Application.Calculation = pjManual
            For Each oTask In oTasks
              If oTask Is Nothing Then GoTo next_task
              If oTask.ExternalTask Then GoTo next_task
              If Not oTask.Active Then GoTo next_task
              If Len(oTask.GetField(lngLCF)) > 0 Then oTask.SetField lngLCF, vbNullString
next_task:
              lngTask = lngTask + 1
              .lboLCF.List(lngItem, 1) = "deleting...(" & Format(lngTask / lngTasks, "0%") & ")"
              Application.StatusBar = "Deleting task data for '" & strFieldName & "'...(" & Format(lngTask / lngTasks, "0%") & ")"
              DoEvents
            Next oTask
            Application.Calculation = pjAutomatic
          End If
          .lboLCF.List(lngItem, 1) = FieldConstantToFieldName(lngLCF)
        Else
          MsgBox "Could not delete LCF '" & strFieldName & "'!" & vbCrLf & vbCrLf & "Please do so manually using the CustomFields button on this form.", vbExclamation + vbOKOnly
        End If
      End If
    Next lngItem
  End With
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oTasks = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptDeleteLCFAndTaskData", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCFMap()
  'objects
  Dim rstSavedMap As ADODB.Recordset
  'strings
  Dim strMsg As String
  Dim strSavedMapExport As String
  Dim strSavedMap As String
  Dim strProgramAcronym As String
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
  strMsg = strMsg & Application.Profiles.ActiveProfile.Server & vbCrLf & vbCrLf
  strMsg = strMsg & "Proceed with export?"
  If MsgBox(strMsg, vbExclamation + vbYesNo, "Important") = vbNo Then GoTo exit_here
  
  'ensure file exists
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) = vbNullString Then
    MsgBox "Saved map file 'cpt-save-local.adtg' not found!", vbExclamation + vbOKOnly, "No Saved Maps"
    GoTo exit_here
  End If

  strProgramAcronym = cptGetProgramAcronym
  
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
    .Filter = "PROGRAM_ACRONYM='" & strProgramAcronym & "'"
    If .RecordCount = 0 Then
      MsgBox "You have no saved map for program acronym '" & strProgramAcronym & "'.", vbExclamation + vbOKOnly, "No Map"
      GoTo exit_here
    Else
      .MoveFirst
      Print #lngFile, .GetString(adClipString, , ",", vbCrLf, vbNullString)
      Close #lngFile
      MsgBox "Map saved to '" & strSavedMapExport & "'", vbInformation + vbOKOnly, "Export Complete"
      cptShellExecute 0, "open", "notepad.exe", strSavedMapExport, vbNullString, 1
    End If
    Close #lngFile
    .Filter = 0
    .Close
  End With
      
exit_here:
  On Error Resume Next
  rstSavedMap.Filter = 0
  rstSavedMap.Close
  Set rstSavedMap = Nothing
  Close #lngFile
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptSaveLocal_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportCFMap(ByRef mySaveLocal_frm As cptSaveLocal_frm)
  'objects
  Dim rstSavedMap As ADODB.Recordset
  Dim oStream As Scripting.TextStream
  Dim oFile As Scripting.File
  Dim oFSO As Scripting.FileSystemObject
  Dim oFiles As Collection
  'strings
  Dim strServerURL As String
  Dim strMsg As String
  Dim strDir As String
  Dim strProgramAcronym As String
  Dim strConn As String
  Dim strSavedMap As String
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
  
  'get server
  strServerURL = Application.Profiles.ActiveProfile.Server
  
  'get program
  strProgramAcronym = cptGetProgramAcronym
  If Len(strProgramAcronym) = 0 Then
    MsgBox "Program Acronym is required. Exiting.", vbExclamation + vbOKOnly, "cptImportMap"
    GoTo exit_here
  End If
  
  Set oFiles = cptGetFiles("Select cpt-saved-map.csv", Environ("USERPROFILE") & "\Downloads\", "Comma Separated Values (*.csv)" & vbNullChar & "*.csv")
  If oFiles.Count > 0 Then
    strSavedMapImport = oFiles(1)
  Else
    MsgBox "No file selected. Exiting.", vbExclamation + vbOKOnly, "cptImportMap"
    GoTo exit_here
  End If
  
  'stream the csv
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFile = oFSO.GetFile(strSavedMapImport)
  Set oStream = oFile.OpenAsTextStream(ForReading)
  
  'open user's saved map
  strSavedMap = strDir & "\settings\cpt-save-local.adtg"
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  With rstSavedMap
    If Dir(strSavedMap) = vbNullString Then 'create it
      .Fields.Append "URL", adVarChar, 255
      .Fields.Append "PROGRAM_ACRONYM", adVarChar, 255 'was "GUID", adGUID
      .Fields.Append "ECF", adInteger
      .Fields.Append "LCF", adInteger
      .Open
      .Save strSavedMap, adPersistADTG
    Else
      .Open strSavedMap, "Provider=MSPersist", , , adCmdFile
    End If

    Do Until oStream.AtEndOfStream
      'todo: prompt to override program acronyms don't match?
      aLine = Split(oStream.ReadLine, ",")
      If UBound(aLine) > 0 Then
        'ensure same server
        If aLine(0) <> strServerURL Then
          strMsg = "You are attempting to import a mapping from server:" & vbCrLf & aLine(0) & vbCrLf & vbCrLf
          strMsg = strMsg & "Maps can only be imported from this server:" & vbCrLf & strServerURL
          MsgBox strMsg, vbCritical + vbOKOnly, "Invalid Map"
          GoTo exit_here
        Else
          'ensure ecf exists
          If FieldConstantToFieldName(CLng(aLine(2))) = "<Unavailable>" Then
            'todo: strECF
            MsgBox "The imported ECF (" & CLng(aLine(2)) & ") is <Unavailable> on this server.", vbCritical + vbOKOnly, "Invalid ECF"
          Else
            'todo: ensure same ECF settings?
            'mySaveLocal_frm.lboECF.Value = CLng(aLine(2))
            mySaveLocal_frm.cboLCF.Value = cptRegEx(FieldConstantToFieldName(CLng(aLine(3))), "^[A-Za-z ]{1,}")
            mySaveLocal_frm.cboLCF_Change
            mySaveLocal_frm.lboLCF.Value = CLng(aLine(3))
            Call cptMapECFtoLCF(mySaveLocal_frm, CLng(aLine(2)), CLng(aLine(3)))
          End If
        End If
      End If
    Loop
    
  End With

  cptUpdateSaveLocalView mySaveLocal_frm

exit_here:
  On Error Resume Next
  Set rstSavedMap = Nothing
  Set oStream = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing
  Set oFiles = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptImportCFMap", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateECF(ByRef mySaveLocal_frm As cptSaveLocal_frm, Optional strFilter As String)
  'objects
  Dim rstECF As ADODB.Recordset
  Dim rstSavedMap As ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strServerURL As String
  Dim strProgramAcronym As String
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
  
  strServerURL = Application.Profiles.ActiveProfile.Server
  
  strProgramAcronym = cptGetProgramAcronym
  If Len(strProgramAcronym) = 0 Then
    MsgBox "Program Acronym is required. Exiting.", vbCritical + vbOKOnly, "ECF to LCF"
    GoTo exit_here
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
    'why what are you doing here
    'catch legacy adtg here and fix
    If rstSavedMap.Fields(1).Name = "GUID" Then
      Stop
    End If
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
        rstECF.Filter = "DATA_TYPE='undetermined' OR DATA_TYPE LIKE '%Maybe%'"
      Else
        rstECF.Filter = "DATA_TYPE='" & .cboECF & "' OR DATA_TYPE='undetermined'"
      End If
    End If
    If rstECF.RecordCount = 0 Then GoTo no_records
    rstECF.MoveFirst
    Do While Not rstECF.EOF
      If rstECF("PROGRAM_ACRONYM") = strProgramAcronym Then
        If strFilter <> "" Then
          If InStr(UCase(rstECF("ECF_Name")), UCase(strFilter)) = 0 Then GoTo next_field
        End If
        .lboECF.AddItem
        .lboECF.List(.lboECF.ListCount - 1, 0) = rstECF("ECF")
        .lboECF.List(.lboECF.ListCount - 1, 1) = rstECF("ECF_Name")
        .lboECF.List(.lboECF.ListCount - 1, 2) = rstECF("DATA_TYPE")
        If blnExists Then
          rstSavedMap.Filter = "PROGRAM_ACRONYM='" & strProgramAcronym & "' AND ECF=" & rstECF("ECF")
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
  Call cptHandleErr(THIS_MODULE, "cptUpdateECF", Err, Erl)
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
      lngFieldID = FieldNameToFieldConstant(strFieldName, pjTask)
      strCustomName = CustomFieldGetName(FieldNameToFieldConstant(.cboLCF.Column(0) & lngField, pjTask))
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
  Call cptHandleErr(THIS_MODULE, "cboLCF_Change", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateSaveLocalView(ByRef mySaveLocal_frm As cptSaveLocal_frm, Optional strAction As String, Optional strMap As String, Optional blnReset As Boolean = False)
  'objects
  'strings
  'longs
  Dim lngLCF As Long
  Dim lngECF As Long
  Dim lngField As Long
  Dim lngCols As Long
  Dim lngCol As Long
  'integers
  'doubles
  'booleans
  Dim blnExists As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  Debug.Print "triggered!" 'todo: remove
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'create/overwrite the tables
  If Not cptTableExists(TABLE_NAME) Then
    TableEditEx TABLE_NAME, True, True, True, , "ID", , , , , , True, , , , , , , , False
    TableEditEx TABLE_NAME, True, False, , , , "Unique ID", "UID", , , , True
    If ActiveProject.Subprojects.Count > 0 Then
      TableEditEx TABLE_NAME, True, False, , , , "Project", , 40, , , True
    End If
  End If
  If Not cptViewExists(VIEW_NAME) Then
    ViewEditSingle VIEW_NAME, True, , pjTaskSheet, , , TABLE_NAME, "All Tasks", "No Group"
  End If

  With mySaveLocal_frm
    If .Visible Then
      'update if add/remove mapping
      If Len(strAction) > 0 And Len(strMap) > 0 Then
        lngECF = CLng(Split(strMap, "|")(0))
        lngLCF = CLng(Split(strMap, "|")(1))
        lngCols = ActiveProject.TaskTables(TABLE_NAME).TableFields.Count
        If strAction = "map" Then
          blnExists = False 'move this and line above
          For lngCol = lngCols To 1 Step -1
            If ActiveProject.TaskTables(TABLE_NAME).TableFields(lngCol).Field = lngLCF Then
              blnExists = True
            ElseIf ActiveProject.TaskTables(TABLE_NAME).TableFields(lngCol).Field = lngECF Then
              blnExists = True
            End If
          Next lngCol
          If Not blnExists Then
            TableEditEx TABLE_NAME, True, False, False, , , FieldConstantToFieldName(lngECF), , , , , True, , , , , , , , False
            TableEditEx TABLE_NAME, True, False, False, , , FieldConstantToFieldName(lngLCF), , , , , True, , , , , , , , False
          End If
        ElseIf strAction = "unmap" Then
          'count down and remove
          For lngCol = lngCols To 1 Step -1
            Stop 'todo: remove
            If ActiveProject.TaskTables(TABLE_NAME).TableFields(lngCol).Field = lngLCF Then
              ActiveProject.TaskTables(TABLE_NAME).TableFields(lngCol).Delete
            ElseIf ActiveProject.TaskTables(TABLE_NAME).TableFields(lngCol).Field = lngECF Then
              ActiveProject.TaskTables(TABLE_NAME).TableFields(lngCol).Delete
            End If
          Next lngCol
        End If
      End If
    Else
      'ensure view is applied
      Stop 'todo: remove
      ActiveWindow.TopPane.Activate
      If ActiveProject.CurrentView <> VIEW_NAME Then
        ViewApply VIEW_NAME
      End If
      'add existing mapping
      For lngField = 0 To .lboECF.ListCount - 1
        If Not IsNull(.lboECF.List(lngField, 3)) Then
          lngECF = .lboECF.List(lngField, 0)
          lngLCF = .lboECF.List(lngField, 3)
          blnExists = False
          lngCols = ActiveProject.TaskTables(TABLE_NAME).TableFields.Count
          For lngCol = lngCols To 1 Step -1
            If ActiveProject.TaskTables(TABLE_NAME).TableFields(lngCol).Field = lngLCF Then
              blnExists = True
            ElseIf ActiveProject.TaskTables(TABLE_NAME).TableFields(lngCol).Field = lngECF Then
              blnExists = True
            End If
          Next lngCol
          If Not blnExists Then
            TableEditEx TABLE_NAME, True, False, False, , , FieldConstantToFieldName(lngECF), , , , , True, , , , , , , , False
            DoEvents
            TableEditEx TABLE_NAME, True, False, False, , , FieldConstantToFieldName(lngLCF), , , , , True, , , , , , , , False
            DoEvents
          End If
        End If
      Next lngField
      TableApply TABLE_NAME
    End If
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptUpdateSaveLocalView", Err, Erl)
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
    For lngType = 0 To dTypes.Count - 1
      For lngField = 1 To dTypes.Items(lngType)
        lngLastRow = lngLastRow + 1
        If lngProject = 0 Then
          oWorksheet.Cells(lngLastRow, 1) = "Task"
          oWorksheet.Cells(lngLastRow, 2) = dTypes.Keys(lngType)
        End If
        lngLCF = FieldNameToFieldConstant(dTypes.Keys(lngType) & lngField, pjTask)
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
  Call cptHandleErr(THIS_MODULE, "cptLocalCustomFieldsMatch", Err, Erl)
  Resume exit_here
End Function
