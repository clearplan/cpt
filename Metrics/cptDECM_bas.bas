Attribute VB_Name = "cptDECM_bas"
'<cpt_version>v7.0.1</cpt_version>
Option Explicit
Private strWBS As String
Private strOBS As String
Private strCA As String
Private strCAM As String
Private strWP As String
Private strEVT As String
Private strEVP As String
Private strPass As String
Private strFail As String
Private lngWBS As Long
Private lngOBS As Long
Private lngCA As Long
Private lngCAM As Long
Private lngWP As Long
Private lngEVT As Long
Private lngEVP As Long
Private lngY As Long
Private blnResourceLoaded As Boolean
Public oDECM As Scripting.Dictionary
Private oSubMap As Scripting.Dictionary

Sub cptDECM_GET_DATA()
  'Optional blnIncompleteOnly As Boolean = True, Optional blnDiscreteOnly As Boolean = True
  'objects
  Dim oSubproject As MSProject.SubProject
  Dim myDECM_frm As cptDECM_frm
  Dim oException As MSProject.Exception
  Dim oTasks As MSProject.Tasks
  Dim oCell As Excel.Range
  Dim oListObject As Excel.ListObject
  Dim oFile As Scripting.TextStream
  Dim oFSO As Scripting.FileSystemObject
  Dim oAssignment As MSProject.Assignment
  Dim oDict As Scripting.Dictionary
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oRecordset As ADODB.Recordset
  Dim oLink As MSProject.TaskDependency
  Dim oTask As MSProject.Task
  'strings
  Dim strProject As String
  Dim strMetric As String
  Dim strRollingWaveDate As String
  Dim strUpdateView As String
  Dim strProgramAcronym As String
  Dim strRequiredFields As String
  Dim strLinks  As String
  Dim strPE As String
  Dim strSE As String
  Dim strRecord As String
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  Dim strFile As String
  Dim strLOE As String
  Dim strList As String
  'longs
  Dim lngDefaultDateFormat As Long
  Dim lngToUID As Long
  Dim lngFactor As Long
  Dim lngFromUID As Long
  Dim lngTargetFile As Long
  Dim lngTaskName As Long
  Dim lngTargetUID As Long
  Dim lngAssignmentFile As Long
  Dim lngTS As Long
  Dim lngConst As Long
  Dim lngX As Long
  'Dim lngY As Long
  Dim lngSummary As Long
  Dim lngDur As Long
  Dim lngBDur As Long
  Dim lngAF As Long
  Dim lngAS As Long
  Dim lngBLF As Long
  Dim lngBLS As Long
  Dim lngFF As Long
  Dim lngFS As Long
  Dim lngUID As Long
  Dim lngTask As Long
  Dim lngLinkFile As Long
  Dim lngTaskFile As Long
  Dim lngWBS As Long
  Dim lngOBS As Long
  Dim lngCA As Long
  Dim lngCAM As Long
  Dim lngWP As Long
  Dim lngWPM As Long
  Dim lngEVT As Long
  Dim lngEVP As Long
  Dim lngFile As Long
  Dim lngTasks As Long
  Dim lngItem As Long
  Dim lngBLW As Long
  Dim lngBLC As Long
  Dim lngRW As Long
  Dim lngRC As Long
  'integers
  'doubles
  Dim dblScore As Double
  'booleans
  Dim blnMaster As Boolean
  Dim blnLimitToPMB As Boolean
  Dim blnTaskHistoryExists As Boolean
  Dim blnFiscalExists As Boolean
  Dim blnDumpToExcel As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vFile As Variant
  Dim vHeader As Variant
  Dim vField As Variant
  'dates
  Dim dtRollingWaveDate As Date
  Dim dtPrevious As Date
  Dim dtCurrent As Date
  Dim dtStatus As Date
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Not IsDate(ActiveProject.StatusDate) Then
    If Not ChangeStatusDate Then
      MsgBox "Status Date is required. Exiting.", vbCritical + vbOKOnly, "Status Date Required"
      GoTo exit_here
    End If
  End If
  
  dtStatus = ActiveProject.StatusDate 'GetField returns mm/dd/yyyy hh:nn AMPM
  'todo: if blnSubprojects then ensure status dates all in sync?
  
  strRequiredFields = "WBS,OBS,CA,CAM,WP,EVT,LOE,PP,EVP"
  If Not cptValidMap(strRequiredFields:=strRequiredFields, blnConfirmationRequired:=True) Then GoTo exit_here
  
  strProgramAcronym = cptGetProgramAcronym
  
  cptSpeed True
  
  lngFile = FreeFile
  strDir = Environ("tmp")
  strFile = strDir & "\Schema.ini"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngFile
  Print #lngFile, "[tasks.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=UID integer"
  Print #lngFile, "Col2=WBS text"
  Print #lngFile, "Col3=OBS text"
  Print #lngFile, "Col4=CA text"
  Print #lngFile, "Col5=CAM text"
  Print #lngFile, "Col6=WP text"
  Print #lngFile, "Col7=WPM text" 'todo: why is WPM required for the DECM?
  Print #lngFile, "Col8=EVT text"
  Print #lngFile, "Col9=EVP integer"
  Print #lngFile, "Col10=FS date"
  Print #lngFile, "Col11=FF date"
  Print #lngFile, "Col12=BLS date"
  Print #lngFile, "Col13=BLF date"
  Print #lngFile, "Col14=AS date"
  Print #lngFile, "Col15=AF date"
  Print #lngFile, "Col16=BDUR integer"
  Print #lngFile, "Col17=DUR integer"
  Print #lngFile, "Col18=SUMMARY text" 'Yes/No
  Print #lngFile, "Col19=CONST text"
  Print #lngFile, "Col20=TS integer"
  Print #lngFile, "Col21=TASK_NAME text"
  Print #lngFile, "Col22=BLW double"
  Print #lngFile, "Col23=BLC double"
  Print #lngFile, "Col24=RW double"
  Print #lngFile, "Col25=RC double"
  Print #lngFile, "[links.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=FROM integer"
  Print #lngFile, "Col2=TO integer"
  Print #lngFile, "Col3=TYPE text"
  Print #lngFile, "Col4=LAG integer"
  Print #lngFile, "Col5=LAG_TYPE integer"
  Print #lngFile, "[assignments.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=TASK_UID integer"
  Print #lngFile, "Col2=RESOURCE_UID integer"
  Print #lngFile, "Col3=BLW Double"
  Print #lngFile, "Col4=BLC Double"
  Print #lngFile, "Col5=RW Double"
  Print #lngFile, "Col6=RC Double"
  For Each vFile In Split("wp-ims.csv,wp-ev.csv,wp-not-in-ims.csv,wp-not-in-ev.csv,10A302b-x.csv,10A303a-x.csv", ",")
    Print #lngFile, "[" & vFile & "]"
    Print #lngFile, "Format=CSVDelimited"
    Print #lngFile, "ColNameHeader=False"
    Print #lngFile, "Col1=WP text"
  Next vFile
  Print #lngFile, "[06A506c-x.csv]"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "Col1=UID integer"
  Print #lngFile, "Col2=P1_TASK_FINISH date"
  Print #lngFile, "Col3=P1_STATUS_DATE date"
  Print #lngFile, "Col4=P1_DELTA Double"
  Print #lngFile, "Col5=P2_TASK_FINISH date"
  Print #lngFile, "Col6=P2_STATUS_DATE date"
  Print #lngFile, "Col7=P2_DELTA Double"
  Print #lngFile, "[fiscal.csv]"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "Col1=FISCAL_END date"
  Print #lngFile, "Col2=LABEL text"
  Print #lngFile, "[targets.csv]"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "Col1=UID integer"
  Print #lngFile, "Col2=TASK_NAME text"
  Print #lngFile, "[segregated.csv]"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "Col1=CA text"
  Print #lngFile, "Col2=WP text"
  Print #lngFile, "Col3=WP_BLW double"
  Print #lngFile, "[itemized.csv]"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "Col1=CA text"
  Print #lngFile, "Col2=CA_BAC double"
  Print #lngFile, "Col3=WP_BAC double"
  Print #lngFile, "Col4=discrepancy double"
  Print #lngFile, "[06A504a.csv]"
  Print #lngFile, "ColNameHeader=False"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "Col1=UID integer"
  Print #lngFile, "Col2=AS_WAS date"
  Print #lngFile, "Col3=AS_IS date"
  Print #lngFile, "[06A504b.csv]"
  Print #lngFile, "ColNameHeader=False"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "Col1=UID integer"
  Print #lngFile, "Col2=AF_WAS date"
  Print #lngFile, "Col3=AF_IS date"
  Close #lngFile
  
  lngTaskFile = FreeFile
  strFile = strDir & "\tasks.csv"
  
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngTaskFile
  
  lngLinkFile = FreeFile
  strFile = strDir & "\links.csv"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngLinkFile
  
  lngAssignmentFile = FreeFile
  strFile = strDir & "\assignments.csv"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngAssignmentFile
  
  lngTargetFile = FreeFile
  strFile = strDir & "\targets.csv"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngTargetFile
  
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  
  'todo: blnMaster
  blnMaster = ActiveProject.Subprojects.Count > 0
  If blnMaster Then
    'todo: run LCF sync?
    
    'set up mapping
    If oSubMap Is Nothing Then
      Set oSubMap = CreateObject("Scripting.Dictionary")
    Else
      oSubMap.RemoveAll
    End If
    For Each oSubproject In ActiveProject.Subprojects
      If Left(oSubproject.Path, 2) <> "<>" Then 'offline
        oSubMap.Add Replace(Dir(oSubproject.Path), ".mpp", ""), 0
      ElseIf Left(oSubproject.Path, 2) = "<>" Then 'online
        oSubMap.Add oSubproject.Path, 0
      End If
    Next oSubproject
    For Each oTask In ActiveProject.Tasks
      If oTask Is Nothing Then GoTo next_mapping_task
      If Not oTask.Active Then GoTo next_mapping_task
      If oSubMap.Exists(oTask.Project) Then
        If oSubMap(oTask.Project) > 0 Then GoTo next_mapping_task
        If Not oTask.Summary Then
          oSubMap.Item(oTask.Project) = CLng(oTask.UniqueID / 4194304)
        End If
      End If
next_mapping_task:
      If oTask.Active Then lngTasks = lngTasks + 1
    Next oTask
    
  Else
    lngTasks = ActiveProject.Tasks.Count
  End If

  'get settings
  lngUID = FieldNameToFieldConstant("Unique ID")
  lngWBS = CLng(Split(cptGetSetting("Integration", "WBS"), "|")(0))
  lngOBS = CLng(Split(cptGetSetting("Integration", "OBS"), "|")(0))
  lngCA = CLng(Split(cptGetSetting("Integration", "CA"), "|")(0))
  lngCAM = CLng(Split(cptGetSetting("Integration", "CAM"), "|")(0))
  lngWP = CLng(Split(cptGetSetting("Integration", "WP"), "|")(0))
  'lngWPM = CLng(Split(cptGetSetting("Integration", "WPM"), "|")(0)) 'note: WPM is not required for DECM and is skipped
  lngEVT = CLng(Split(cptGetSetting("Integration", "EVT"), "|")(0))
  strLOE = cptGetSetting("Integration", "LOE")
  lngEVP = CLng(Split(cptGetSetting("Integration", "EVP"), "|")(0))
  'todo: clean EVP: remove %; normalize to whole number?
  lngFS = FieldNameToFieldConstant("Start")
  lngFF = FieldNameToFieldConstant("Finish")
  lngBLS = FieldNameToFieldConstant("Baseline Start")
  lngBLF = FieldNameToFieldConstant("Baseline Finish")
  lngAS = FieldNameToFieldConstant("Actual Start")
  lngAF = FieldNameToFieldConstant("Actual Finish")
  lngBDur = FieldNameToFieldConstant("Baseline Duration")
  lngDur = FieldNameToFieldConstant("Duration")
  lngSummary = FieldNameToFieldConstant("Summary")
  lngConst = FieldNameToFieldConstant("Constraint Type")
  lngTS = FieldNameToFieldConstant("Total Slack")
  lngTaskName = FieldNameToFieldConstant("Name", pjTask)
  lngBLW = FieldNameToFieldConstant("Baseline Work", pjTask)
  lngBLC = FieldNameToFieldConstant("Baseline Cost", pjTask)
  lngRW = FieldNameToFieldConstant("Remaining Work", pjTask)
  lngRC = FieldNameToFieldConstant("Remaining Cost", pjTask)
  
  'headers
  Print #lngTaskFile, "UID,WBS,OBS,CA,CAM,WP,WPM,EVT,EVP,FS,FF,BLS,BLF,AS,AF,BDUR,DUR,SUMMARY,CONST,TS,TASK_NAME,BLW,BLC,RW,RC," 'note: WPM is not required for DECM and is skipped
  Print #lngLinkFile, "FROM,TO,TYPE,LAG,"
  Print #lngAssignmentFile, "TASK_UID,RESOURCE_UID,BLW,BLC,RW,RC,"
  Print #lngTargetFile, "UID,TASK_NAME,"
  
  Set myDECM_frm = New cptDECM_frm
  With myDECM_frm
    .Caption = "DECM v7.0 (cpt " & cptGetVersion("cptDECM_bas") & ")"
    .lboOOS.Visible = False
    lngItem = 0
    .lboHeader.Clear
    .lboHeader.AddItem
    For Each vHeader In Split("METRIC,TITLE,TARGET,X,Y,SCORE,ICON,DESCRIPTION,TBD", ",")
      .lboHeader.List(0, lngItem) = vHeader
      lngItem = lngItem + 1
    Next vHeader
    .lboMetrics.Clear
    strUpdateView = cptGetSetting("Integration", "chkUpdateView")
    If Len(strUpdateView) > 0 Then
      .chkUpdateView = CBool(strUpdateView)
    Else
      .chkUpdateView = True 'default
    End If
    .cmdExport.Enabled = False
    .cmdDone.Enabled = False
    .Show False
  End With
  
  blnDumpToExcel = False 'for debug
  blnLimitToPMB = False 'don't do this
  lngDefaultDateFormat = Application.DefaultDateFormat
  If lngDefaultDateFormat <> pjDate_mm_dd_yyyy Then
    Application.DefaultDateFormat = pjDate_mm_dd_yyyy
  End If
  blnResourceLoaded = False
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If blnLimitToPMB Then
      If oTask.Assignments.Count = 0 Then GoTo next_task
      If oTask.BaselineWork + oTask.BaselineCost = 0 Then GoTo next_task
    End If

    'todo: what about NOT blnMaster AND External Tasks?
    'If oTask.Summary Then GoTo next_task
'    If blnIncompleteOnly Then If IsDate(oTask.ActualFinish) Then GoTo next_task 'todo: what was this for?
'    If blnDiscreteOnly Then If oTask.GetField(lngEVT) = "A" Then GoTo next_task 'todo: what else is non-discrete? apportioned?
    
    For Each vField In Array(lngUID, lngWBS, lngOBS, lngCA, lngCAM, lngWP, lngWPM, lngEVT, lngEVP, lngFS, lngFF, lngBLS, lngBLF, lngAS, lngAF, lngBDur, lngDur, lngSummary, lngConst, lngTS, lngTaskName, lngBLW, lngBLC, lngRW, lngRC)
      If vField = 0 Then
        strRecord = strRecord & "," 'account for empty WPM
        GoTo next_field
      End If
      If vField = FieldNameToFieldConstant("Physical % Complete") Then
        strRecord = strRecord & cptRegEx(oTask.GetField(vField), "[0-9]{1,}") & ","
      ElseIf vField = FieldNameToFieldConstant("% Complete") Then
        strRecord = strRecord & cptRegEx(oTask.GetField(vField), "[0-9]{1,}") & ","
      ElseIf vField = lngBDur Then
        strRecord = strRecord & oTask.BaselineDuration & ","
      ElseIf vField = lngDur Then
        strRecord = strRecord & oTask.Duration & ","
      ElseIf vField = lngConst Then
        strRecord = strRecord & Choose(oTask.ConstraintType + 1, "ASAP", "ALAP", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & ","
      ElseIf vField = lngTS Then
        strRecord = strRecord & oTask.TotalSlack & "," 'todo: convert to days?
      ElseIf Len(cptRegEx(FieldConstantToFieldName(vField), "Start|Finish")) > 0 And IsDate(oTask.GetField(vField)) Then 'convert text to date if field name as 'start' or 'finish')
        strRecord = strRecord & FormatDateTime(oTask.GetField(CLng(vField)), vbShortDate) & ","
      ElseIf vField = lngTaskName Then
        strRecord = strRecord & Replace(Replace(oTask.Name, Chr(34), "'"), ",", "-") & ","
      ElseIf FieldConstantToFieldName(vField) = "Baseline Work" Then
        'strRecord = strRecord & Replace(cptRegEx(oTask.GetField(vField), "[0-9.,]{1,}"), ",", "") & ","
        If oTask.Summary Then
          strRecord = strRecord & "0,"
        Else
          strRecord = strRecord & oTask.BaselineWork & ","
        End If
      ElseIf FieldConstantToFieldName(vField) = "Baseline Cost" Then
        'strRecord = strRecord & Replace(cptRegEx(oTask.GetField(vField), "[0-9.,]{1,}"), ",", "") & ","
        If oTask.Summary Then
          strRecord = strRecord & "0,"
        Else
          strRecord = strRecord & oTask.BaselineCost & ","
        End If
      ElseIf FieldConstantToFieldName(vField) = "Remaining Work" Then
        'strRecord = strRecord & Replace(cptRegEx(oTask.GetField(vField), "[0-9.,]{1,}"), ",", "") & ","
        If oTask.Summary Then
          strRecord = strRecord & "0,"
        Else
          strRecord = strRecord & oTask.RemainingWork & ","
        End If
      ElseIf FieldConstantToFieldName(vField) = "Remaining Cost" Then
        'strRecord = strRecord & Replace(cptRegEx(oTask.GetField(vField), "[0-9.,]{1,}"), ",", "") & ","
        If oTask.Summary Then
          strRecord = strRecord & "0,"
        Else
          strRecord = strRecord & Replace(cptRegEx(oTask.GetField(vField), "[0-9.,]{1,}"), ",", "") & ","
        End If
      Else
        strRecord = strRecord & oTask.GetField(CLng(vField)) & ","
      End If
next_field:
    Next vField
    strRecord = Left(strRecord, Len(strRecord) - 1) 'hack off last comma
    Print #lngTaskFile, strRecord
    For Each oLink In oTask.TaskDependencies
      'todo: convert lag to effective days?
      If oTask.Guid = oLink.To.Guid Then 'get predecessors
        lngToUID = oTask.UniqueID
        If blnMaster And oLink.From.ExternalTask Then
          'fix the pred UID if master-sub
          lngFromUID = oLink.From.GetField(185073906) Mod 4194304
          strProject = oLink.From.Project
          If InStr(oLink.From.Project, "\") > 0 Then
            strProject = Replace(strProject, ".mpp", "")
            strProject = Mid(strProject, InStrRev(strProject, "\") + 1)
          End If
          lngFactor = oSubMap(strProject)
          lngFromUID = (lngFactor * 4194304) + lngFromUID
        Else
          If blnMaster Then
            lngFactor = Round(oTask.UniqueID / 4194304, 0)
            lngFromUID = (lngFactor * 4194304) + oLink.From.UniqueID
          Else
            lngFromUID = oLink.From.UniqueID
          End If
        End If
      ElseIf oTask.Guid = oLink.From.Guid Then 'get successors
        lngFromUID = oTask.UniqueID
        If blnMaster And oLink.To.ExternalTask Then
          lngToUID = oLink.To.GetField(185073906) Mod 4194304
          strProject = oLink.To.Project
          If InStr(strProject, "\") > 0 Then
            strProject = Replace(strProject, ".mpp", "")
            strProject = Mid(strProject, InStrRev(strProject, "\") + 1)
          End If
          lngFactor = oSubMap(strProject)
          lngToUID = (lngFactor * 4194304) + lngToUID
        Else
          If blnMaster Then
            lngFactor = Round(oTask.UniqueID / 4194304, 0)
            lngToUID = (lngFactor * 4194304) + oLink.To.UniqueID
          Else
            lngToUID = oLink.To.UniqueID
          End If
        End If
      End If
      Print #lngLinkFile, lngFromUID & "," & lngToUID & "," & Choose(oLink.Type + 1, "FF", "FS", "SF", "SS") & "," & oLink.Lag & ","
    Next oLink
    For Each oAssignment In oTask.Assignments
      blnResourceLoaded = True
      Print #lngAssignmentFile, Join(Array(oTask.UniqueID, oAssignment.ResourceUniqueID, oAssignment.BaselineWork, oAssignment.BaselineCost, oAssignment.RemainingWork, oAssignment.RemainingCost), ",")
    Next
    'only capture incomplete milestones as targets
    If (oTask.Duration = 0 Or oTask.Milestone) And Not oTask.ExternalTask And Not IsDate(oTask.ActualFinish) Then
      Print #lngTargetFile, Join(Array(oTask.UniqueID, Replace(Replace(oTask.Name, ",", ""), Chr(34), "'")), ",")
    End If
next_task:
    strRecord = ""
    lngTask = lngTask + 1
    Application.StatusBar = "Loading Data...(" & Format(lngTask / lngTasks, "0%") & ")"
    myDECM_frm.lblStatus.Caption = "Loading Data...(" & Format(lngTask / lngTasks, "0%") & ")"
    myDECM_frm.lblProgress.Width = (lngTask / lngTasks) * myDECM_frm.lblStatus.Width
    DoEvents
  Next oTask
  
  Close #lngTaskFile
  Close #lngLinkFile
  Close #lngAssignmentFile
  Close #lngTargetFile
  
  myDECM_frm.lblStatus.Caption = "Loading...done."
  Application.StatusBar = "Loading...done."
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  
  'lboMetrics: METRIC,TITLE,THRESHOLD,X,Y,SCORE,DESCRIPTION,?sql
  strPass = "[+]"
  strFail = "<!>"
  'strMore = "..."
  'strQuestion = "<?>"
  'strError = "<!>"
  
  'confirm fiscal calendar exists and export it
  'needed for: 10A103a; 06A504a; 06A504b;
  blnFiscalExists = cptCalendarExists("cptFiscalCalendar")
  If blnFiscalExists Then
    'export fiscal.csv
    lngFile = FreeFile
    strFile = strDir & "\fiscal.csv"
    Open strFile For Output As #lngFile
    Print #1, "FISCAL_END,LABEL,"
    For Each oException In ActiveProject.BaseCalendars("cptFiscalCalendar").Exceptions
      Print #1, oException.Finish & " 5:00 PM," & oException.Name & ","
    Next oException
    Close #lngFile
  End If
  
  'ad hoc todo: destroy on form close
  Set oDECM = CreateObject("Scripting.Dictionary")
  
  'check for missing metadata
  If Not DECM_CPT01(oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel) Then GoTo exit_here 'bonus - missing metadata
  '===== EVMS =====
  DECM_05A101a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '05A101a - 1 CA : 1 OBS
  DECM_05A102a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '05A102a - 1 CA : 1 CAM
  DECM_05A103a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '05A103a - 1 CA : 1 WBS
  DECM_CPT02 oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel 'bonus - 1 WP : 1 CA
  DECM_10A102a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '10A102a - 1 WP : 1 EVT
  DECM_10A103a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel, blnFiscalExists '10A103a - 0/100 EVTs in one fiscal period
  DECM_10A109b oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '10A109b - all WPs have budget
  DECM_10A302b oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '10A302b - PPs with progress
  DECM_10A303a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '10A303a - all PPs have duration?
  DECM_11A101a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '11A101a - CA BAC = SUM(WP BAC)?
  '===== SCHEDULE =====
  DECM_06A101a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A101a - WPs Missing between IMS vs EV
  DECM_06A204b oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A204b - Dangling Logic
  DECM_06A205a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A205a - Lags (what about leads?)
  DECM_CPT03 oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel  'bonus - leads
  DECM_06A208a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A208a - summary tasks with logic
  DECM_06A209a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A209a - hard constraints
  DECM_06A210a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A210a - LOE Driving Discrete
  DECM_06A211a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A211a - High Float
  DECM_06A212a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A212a - out of sequence
  DECM_06A401a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A401a - critical path (constraint method)
  DECM_06A501a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A501a - baselines
    
  'confirm cpt-cei.adtg exists and convert it to csv for easy querying
  'needed for: 06A504a; 06A504b;
  myDECM_frm.lblStatus.Caption = "Task history file..."
  Application.StatusBar = "Task history file..."
  DoEvents
  strFile = cptDir & "\settings\cpt-cei.adtg"
  If Dir(strFile) <> vbNullString Then
    myDECM_frm.lblStatus.Caption = "Task history file found. Querying..."
    Application.StatusBar = "Task history file found. Querying..."
    DoEvents
    
    'copy cpt-cei.adtg to tmp dir
    FileCopy strFile, strDir & "\cpt-cei.adtg"
    
    'convert to csv for sql query...
    strFile = strDir & "\cpt-cei.adtg"
    Set oRecordset = CreateObject("ADODB.Recordset")
    oRecordset.Open strFile
    'limit to program
    oRecordset.Filter = "PROJECT='" & strProgramAcronym & "'"
    oRecordset.Save strFile, adPersistADTG
    oRecordset.Close
    
    'clean it up (remove commas)
    'todo: any other fields that might have a comma?
    oRecordset.Open strFile
    oRecordset.Filter = "TASK_NAME LIKE '%,%'"
    If Not oRecordset.EOF Then
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
        oRecordset.Fields("TASK_NAME") = Replace(oRecordset.Fields("TASK_NAME"), ",", "-")
        oRecordset.MoveNext
      Loop
    End If
    oRecordset.Filter = 0
    oRecordset.Save strFile
    
    If oRecordset.RecordCount > 0 Then
      'capture field names
      strList = ""
      For lngItem = 0 To oRecordset.Fields.Count - 1
        strList = strList & oRecordset.Fields(lngItem).Name & ","
      Next lngItem
      'save as csv
      Set oFSO = CreateObject("Scripting.FileSystemObject")
      Set oFile = oFSO.CreateTextFile(strDir & "\cpt-cei.csv", True)
      oFile.Write strList & vbCrLf
      oRecordset.MoveFirst
      oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
      oRecordset.Close
      Kill strDir & "\cpt-cei.adtg"
      oFile.Close
    End If
  Else
    myDECM_frm.lblStatus.Caption = "Task history file...not found."
    Application.StatusBar = "Task history...not found."
    DoEvents
  End If
  
  If Dir(strDir & "\cpt-cei.csv") <> vbNullString And blnFiscalExists Then
    'is there more than one period of history?
    Set oRecordset = CreateObject("ADODB.Recordset")
    strSQL = "SELECT DISTINCT STATUS_DATE "
    strSQL = strSQL & "FROM [cpt-cei.csv] "
    strSQL = strSQL & "WHERE PROJECT ='" & strProgramAcronym & "'"
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If oRecordset.EOF Then
      blnTaskHistoryExists = False
    ElseIf oRecordset.RecordCount < 2 Then
      blnTaskHistoryExists = False
    Else
      blnTaskHistoryExists = True
    End If
    oRecordset.Close
    
    If blnTaskHistoryExists Then 'get previous
      Set oRecordset = CreateObject("ADODB.Recordset")
      strSQL = "SELECT "
      strSQL = strSQL & "    MAX(STATUS_DATE) "
      strSQL = strSQL & "FROM "
      strSQL = strSQL & "    ( "
      strSQL = strSQL & "        SELECT "
      strSQL = strSQL & "            T1.FISCAL_END, "
      strSQL = strSQL & "            T2.STATUS_DATE "
      strSQL = strSQL & "        FROM "
      strSQL = strSQL & "            [fiscal.csv] T1 "
      strSQL = strSQL & "            LEFT JOIN ( "
      strSQL = strSQL & "                SELECT "
      strSQL = strSQL & "                    DISTINCT STATUS_DATE "
      strSQL = strSQL & "                FROM "
      strSQL = strSQL & "                    [cpt-cei.csv] "
      strSQL = strSQL & "                WHERE "
      strSQL = strSQL & "                    PROJECT = '" & strProgramAcronym & "' "
      strSQL = strSQL & "            ) T2 ON T2.[STATUS_DATE] = T1.FISCAL_END "
      strSQL = strSQL & "    )"
      oRecordset.Open strSQL, strCon, adOpenKeyset
      If oRecordset.EOF Then
        oRecordset.Close
        blnTaskHistoryExists = False
      End If
      dtCurrent = oRecordset(0)
      oRecordset.Close
      'get 2nd most recent fiscal end/status date
      strSQL = strSQL & " WHERE STATUS_DATE<#" & dtCurrent & "#"
      oRecordset.Open strSQL, strCon, adOpenKeyset
      If oRecordset.EOF Then
        oRecordset.Close
        blnTaskHistoryExists = False
      End If
      dtPrevious = oRecordset(0)
      oRecordset.Close
    End If
  End If
  
  If blnTaskHistoryExists Then
    myDECM_frm.lblStatus.Caption = "Previous period found."
    Application.StatusBar = "Previous period found."
  Else
    myDECM_frm.lblStatus.Caption = "Previous period not found."
    Application.StatusBar = "Previous period not found."
  End If
  DoEvents
  
  '06A504a - AS changed - only if task history otherwise notify to 'use capture period'
  strMetric = "06A504a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A505a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Changed Actual Start"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 10%"
  DoEvents
  
  'X = Count of tasks/activities & milestones where actual start date does not equal previously reported actual start date
  'Y = Total count of tasks/activities & milestones with actual start dates
  
  If blnTaskHistoryExists Then
  
    'todo: requires cptFiscal
    'todo: confirm previous month end [list of status dates in cpt-cei where project, order desc]
    'todo: confirm current month end against fiscal
  
    'get Y
    strSQL = "SELECT TASK_UID,TASK_AS FROM [cpt-cei.csv] "
    strSQL = strSQL & "WHERE PROJECT='" & strProgramAcronym & "' AND TASK_AS IS NOT NULL "
    strSQL = strSQL & "AND STATUS_DATE = #" & dtPrevious & "#"
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If Not oRecordset.EOF Then lngY = oRecordset.RecordCount Else lngY = 1
    oRecordset.Close
  
    'get X
    strSQL = "SELECT "
    strSQL = strSQL & "    t1.TASK_UID, "
    strSQL = strSQL & "    t1.TASK_AS AS TASK_AS_IS, "
    strSQL = strSQL & "    t2.TASK_AS_WAS "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    [cpt-cei.csv] AS t1 "
    strSQL = strSQL & "    INNER JOIN ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            TASK_UID, "
    strSQL = strSQL & "            TASK_AS AS TASK_AS_WAS "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            [cpt-cei.csv] "
    strSQL = strSQL & "        WHERE "
    strSQL = strSQL & "            PROJECT = '" & strProgramAcronym & "' "
    strSQL = strSQL & "            AND TASK_AS IS NOT NULL "
    strSQL = strSQL & "            AND STATUS_DATE = #" & dtPrevious & "#"
    strSQL = strSQL & "    ) AS T2 ON T2.TASK_UID = T1.TASK_UID "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    PROJECT = '" & strProgramAcronym & "' "
    strSQL = strSQL & "    AND TASK_AS IS NOT NULL "
    strSQL = strSQL & "    AND STATUS_DATE = #" & dtCurrent & "# "
    strSQL = strSQL & "    AND T1.TASK_AS <> T2.TASK_AS_WAS; "
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If Not oRecordset.EOF Then lngX = oRecordset.RecordCount Else lngX = 0
    strList = ""
    If lngX > 0 Then
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
        strList = strList & oRecordset(0) & ","
        oRecordset.MoveNext
      Loop
      'save results for later
      Set oFile = oFSO.CreateTextFile(strDir & "\06A504a.csv", True)
      oRecordset.MoveFirst
      oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
      oFile.Close
    End If
    oRecordset.Close
    'X/Y <= 10%
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
    dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
    If dblScore <= 0.1 Then
      myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
    Else
      myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
    End If
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
    'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
    oDECM.Add strMetric, strList
    myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
    Application.StatusBar = "Getting " & strMetric & "...done."
    DoEvents
    
  Else 'Not blnTaskHistoryExists
  
    'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = Null 'X
    'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = Null 'Y
    'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Null 'SCORE
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = "-" 'ICON
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = "No Task History found. Please run ClearPlan > Schedule > Status > Capture Week before and after each Status Period to capture Task History."
    myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...skipped."
    Application.StatusBar = "Getting " & strMetric & "...skipped."
    DoEvents
    
  End If 'blnTaskHistoryExists
  
  '06A504b - AF changed - only if task history
  strMetric = "06A504b"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A505a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Changed Actual Finish"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 10%"
  DoEvents

  'X = Count of tasks/activities & milestones where actual finish date does not equal previously reported actual finish date
  'Y = Total count of tasks/activities & milestones with actual finish dates

  If blnTaskHistoryExists Then
  
    'get Y
    strSQL = "SELECT TASK_UID,TASK_AF FROM [cpt-cei.csv] "
    strSQL = strSQL & "WHERE PROJECT='" & strProgramAcronym & "' AND TASK_AF IS NOT NULL "
    strSQL = strSQL & "AND STATUS_DATE = #" & dtPrevious & "#"
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If Not oRecordset.EOF Then lngY = oRecordset.RecordCount Else lngY = 1
    oRecordset.Close
    
    'get X
    strSQL = "SELECT "
    strSQL = strSQL & "    t1.TASK_UID, "
    strSQL = strSQL & "    t1.TASK_AF AS TASK_AF_IS, "
    strSQL = strSQL & "    t2.TASK_AF_WAS "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    [cpt-cei.csv] AS t1 "
    strSQL = strSQL & "    INNER JOIN ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            TASK_UID, "
    strSQL = strSQL & "            TASK_AF AS TASK_AF_WAS "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            [cpt-cei.csv] "
    strSQL = strSQL & "        WHERE "
    strSQL = strSQL & "            PROJECT = '" & strProgramAcronym & "' "
    strSQL = strSQL & "            AND TASK_AF IS NOT NULL "
    strSQL = strSQL & "            AND STATUS_DATE = #" & dtPrevious & "#"
    strSQL = strSQL & "    ) AS T2 ON T2.TASK_UID = T1.TASK_UID "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    PROJECT = '" & strProgramAcronym & "' "
    strSQL = strSQL & "    AND TASK_AF IS NOT NULL "
    strSQL = strSQL & "    AND STATUS_DATE = #" & dtCurrent & "# "
    strSQL = strSQL & "    AND T1.TASK_AF <> T2.TASK_AF_WAS; "
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If Not oRecordset.EOF Then lngX = oRecordset.RecordCount Else lngX = 0
    strList = ""
    If lngX > 0 Then
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
        strList = strList & oRecordset(0) & ","
        oRecordset.MoveNext
      Loop
      'save results for later
      Set oFile = oFSO.CreateTextFile(strDir & "\06A504b.csv", True)
      oRecordset.MoveFirst
      oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
      oFile.Close
    End If
    oRecordset.Close
    
    'X/Y <= 10%
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
    dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
    If dblScore <= 0.1 Then
      myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
    Else
      myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
    End If
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
    'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
    oDECM.Add strMetric, strList
    myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
    Application.StatusBar = "Getting " & strMetric & "...done."
    DoEvents
    
  Else 'blnTaskHistoryExists
  
    'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = Null 'X
    'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = Null 'Y
    'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Null 'SCORE
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = "-" 'ICON
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = "No Task History found. Please run ClearPlan > Schedule > Status > Capture Week before and after each Status Period to capture Task History."
    myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...skipped."
    Application.StatusBar = "Getting " & strMetric & "...skipped."
    DoEvents
    
  End If
  
  DECM_06A505a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A505a - In-Progress Tasks Have AS
  DECM_06A505b oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel '06A505b - Complete Tasks Have AF
  DECM_06A506a oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel, dtStatus '06A506a - bogus actuals
  DECM_06A506b oDECM, myDECM_frm, strCon, oRecordset, blnDumpToExcel, dtStatus  '06A506b - invalid forecast
  
  'todo: allow user to refresh analysis on a one-by-one basis?
  
  '06A506c - riding status date
  strMetric = "06A506c"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A506b"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Riding the Status Date"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 1%"
  DoEvents
  
  'X = Count of incomplete tasks/activities & milestones with either forecast start or forecast finish date riding the status date
  'Y = Total count of incomplete tasks/activities & milestones
  
  If blnTaskHistoryExists Then
  
    strSQL = "SELECT UID FROM [tasks.csv] "
    strSQL = strSQL & "WHERE SUMMARY='No' AND [AF] IS NULL "
    With oRecordset
      .Open strSQL, strCon, adOpenKeyset
      lngY = oRecordset.RecordCount
      If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
      .Close
    End With
      
    'get list of incomplete UID,finish1,status1,delta1,finish2,status2,delta2 where delta2=delta1
    strSQL = "SELECT "
    strSQL = strSQL & "    P1.*, "
    strSQL = strSQL & "    P2.TASK_FINISH, "
    strSQL = strSQL & "    P2.STATUS_DATE, "
    strSQL = strSQL & "    P2.DELTA "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            TASK_UID, "
    strSQL = strSQL & "            TASK_FINISH, "
    strSQL = strSQL & "            STATUS_DATE, "
    strSQL = strSQL & "            TASK_FINISH - STATUS_DATE AS [DELTA] "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            [cpt-cei.csv] "
    strSQL = strSQL & "        WHERE "
    strSQL = strSQL & "            PROJECT = 'IBL' "
    strSQL = strSQL & "            AND IS_LOE = FALSE "
    strSQL = strSQL & "            AND TASK_AF IS NULL "
    strSQL = strSQL & "            AND TASK_FINISH >= #" & dtCurrent & "# "
    strSQL = strSQL & "            AND STATUS_DATE = #" & dtCurrent & "# "
    strSQL = strSQL & "    ) P1 "
    strSQL = strSQL & "    INNER JOIN ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            TASK_UID, "
    strSQL = strSQL & "            TASK_FINISH, "
    strSQL = strSQL & "            STATUS_DATE, "
    strSQL = strSQL & "            TASK_FINISH - STATUS_DATE AS [DELTA] "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            [cpt-cei.csv] "
    strSQL = strSQL & "        WHERE "
    strSQL = strSQL & "            PROJECT = 'IBL' "
    strSQL = strSQL & "            AND IS_LOE = FALSE "
    strSQL = strSQL & "            AND TASK_AF IS NULL "
    strSQL = strSQL & "            AND TASK_FINISH >= #" & dtPrevious & "# "
    strSQL = strSQL & "            AND STATUS_DATE = #" & dtPrevious & "# "
    strSQL = strSQL & "    ) P2 ON P2.TASK_UID = P1.TASK_UID "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    P1.DELTA = P2.DELTA "
    Set oRecordset = CreateObject("ADODB.Recordset")
    With oRecordset
      .Open strSQL, strCon, adOpenKeyset
      If .EOF Then
        lngX = 0
      Else
        lngX = .RecordCount
        Set oFile = oFSO.CreateTextFile(strDir & "\06A506c-x.csv", True)
        oRecordset.MoveFirst
        oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
        oFile.Close
      End If
      strList = ""
      If lngX > 0 Then
        .MoveFirst
        Do While Not .EOF
          strList = strList & .Fields("TASK_UID") & ","
          .MoveNext
        Loop
      End If
      If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
      .Close
    End With
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
    dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
    If dblScore < 0.01 Then
      myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
    Else
      myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
    End If
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
    'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
    oDECM.Add strMetric, strList
    myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
    Application.StatusBar = "Getting " & strMetric & "...done."
    DoEvents
  Else 'blnTaskHistoryExists
    'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = Null 'X
    'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = Null 'Y
    'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Null 'SCORE
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = "-" 'ICON
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = "No Task History found. Please run ClearPlan > Schedule > Status > Capture Week before and after each Status Period to capture Task History."
    myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...skipped."
    Application.StatusBar = "Getting " & strMetric & "...skipped."
    DoEvents
  End If 'blnTaskHistoryExists
  
  '06I201a - SVTs todo: capture task names with "^SVT" ; allow alternative
  strMetric = "06I201a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06I201a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Schedule Visibility Tasks (SVTs)"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = Count of incomplete tasks/activities and milestones that are not properly identified and controlled as SVTs in the IMS
  'X = 0
        
  ActiveWindow.TopPane.Activate
  DoEvents
  OpenUndoTransaction "cpt DECM 06I201a"
  DoEvents
  FilterClear
  GroupClear
  OptionsViewEx DisplaySummaryTasks:=True
  OutlineShowAllTasks
  FilterEdit "cpt DECM Filter - 06I201a", True, True, True, , , "Actual Finish", , "equals", "NA"
  If Application.Edition = pjEditionProfessional Then
    FilterEdit "cpt DECM Filter - 06I201a", True, , , , , , "Active", "equals", "Yes"
  End If
  FilterEdit "cpt DECM Filter - 06I201a", True, , , , , , "Resource Names", "does not equal", ""
  FilterEdit "cpt DECM Filter - 06I201a", True, , , , , , "Name", "contains", "SVT", , , False
  FilterApply "cpt DECM Filter - 06I201a"
  SelectAll
  CloseUndoTransaction
  DoEvents
  Set oTasks = Nothing
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strList = ""
  If Not oTasks Is Nothing Then
    lngX = ActiveSelection.Tasks.Count
    For Each oTask In oTasks
      strList = strList & oTask.UniqueID & ","
    Next oTask
  Else
    lngX = 0
  End If
  FilterClear

  If GetUndoListCount > 0 Then
    If GetUndoListItem(1) = "cpt DECM 06I201a" Then Undo 'todo: why isn't label 'taking'?
  End If
  
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 4) = lngY there is no Y
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If lngX = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
  
  '29A601a - rolling wave period is detail planned
  strMetric = "29A601a"
  strRollingWaveDate = cptGetSetting("Integration", "RollingWaveDate")
  If Len(strRollingWaveDate) > 0 Then
    dtRollingWaveDate = CDate(strRollingWaveDate)
    
    myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
    Application.StatusBar = "Getting " & strMetric & "..."
    myDECM_frm.lboMetrics.AddItem
    myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
    'myDECM_Frm.lboMetrics.Value = "29A601a"
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Rolling Wave Planning"
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 10%"
    DoEvents
    'X = Count of PPs/SLPPs where baseline start precedes the next rolling wave cycle
    'Y = Total count of PPs/SLPPs
    
    strSQL = "SELECT DISTINCT WP "
    strSQL = strSQL & "FROM [tasks.csv] "
    strSQL = strSQL & "WHERE EVT='K'" 'K = PP and SLPP
    Set oRecordset = CreateObject("ADODB.Recordset")
    strList = ""
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If oRecordset.EOF Then
      lngY = 0
    Else
      If oRecordset.RecordCount > 0 Then
        lngY = oRecordset.RecordCount
      Else
        lngY = 0
      End If
    End If
    oRecordset.Close
    
    strSQL = "SELECT DISTINCT WP FROM [tasks.csv] WHERE EVT='K' AND BLS <= #" & FormatDateTime(dtRollingWaveDate, vbGeneralDate) & " 5:00 PM#"
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If oRecordset.EOF Then
      lngX = 0
    Else
      If oRecordset.RecordCount > 0 Then
        lngX = oRecordset.RecordCount
        oRecordset.MoveFirst
        Do While Not oRecordset.EOF
          strList = strList & oRecordset("WP") & ","
          oRecordset.MoveNext
        Loop
      Else
        lngX = 0
      End If
    End If
    oRecordset.Close
    strSQL = "SELECT DISTINCT WP FROM [tasks.csv] WHERE EVT='K' and BLS <= #" & FormatDateTime(dtRollingWaveDate, vbGeneralDate) & " 5:00 PM#"
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If Not oRecordset.EOF Then
      lngY = oRecordset.RecordCount
    Else
      lngY = 1
    End If
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
    dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
    If (lngX / lngY) <= 0.1 Then
      myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
    Else
      myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
    End If
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
    'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
    oDECM.Add strMetric, strList
    myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
    Application.StatusBar = "Getting " & strMetric & "...done."
    DoEvents
    
  End If 'Len(strRollingWaveDate) > 0
  
  myDECM_frm.lboMetrics.ListIndex = 0
  myDECM_frm.cmdExport.Enabled = True
  myDECM_frm.cmdDone.Enabled = True
  
  Application.StatusBar = "DECM Scoring Complete"
  myDECM_frm.lblStatus.Caption = "DECM Scoring Complete"
  DoEvents
  
exit_here:
  On Error Resume Next
  If lngDefaultDateFormat > 0 And Application.DefaultDateFormat <> lngDefaultDateFormat Then
    Application.DefaultDateFormat = lngDefaultDateFormat
  End If
  Set myDECM_frm = Nothing
  Set oException = Nothing
  Set oTasks = Nothing
  Set oCell = Nothing
  Set oListObject = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing
  Set oAssignment = Nothing
  Set oDict = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  cptSpeed False
  Application.StatusBar = ""
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  Reset 'closes all active files opened by the Open statement and writes the contents of all file buffers to disk.
  Set oLink = Nothing
  Set oTask = Nothing
  
  Exit Sub
err_here:
 On Error Resume Next
 Call cptHandleErr("cptDECM", "cptDECM_GET_DATA", Err, Erl)
 Resume exit_here
End Sub

Function DECM_CPT01(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean) As Boolean
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  Dim blnProceed As Boolean
  
  'missing metadata
  myDECM_frm.lblStatus.Caption = "Checking for missing metadata..."
  Application.StatusBar = "Checking for missing metadata..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = "CPT01"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "MISSING METADATA"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  If blnResourceLoaded Then
    strSQL = "SELECT T1.UID,"
    strSQL = strSQL & "IIF(ISNULL(WBS),'MISSING',WBS) AS [WBS],"
    strSQL = strSQL & "IIF(ISNULL(OBS),'MISSING',OBS) AS [OBS],"
    strSQL = strSQL & "IIF(ISNULL(CA),'MISSING',CA) AS [CA],"
    strSQL = strSQL & "IIF(ISNULL(CAM),'MISSING',CAM) AS [CAM],"
    strSQL = strSQL & "IIF(ISNULL(WP),'MISSING',WP) AS [WP],"
    strSQL = strSQL & "IIF(ISNULL(EVT),'MISSING',EVT) AS [EVT],"
    strSQL = strSQL & "EVP,BLS,BLF,[AS],AF,SUM(T2.BLW)/60 AS [BLW],SUM(T2.BLC) AS [BLC] "
    strSQL = strSQL & "FROM [tasks.csv] T1 INNER JOIN [assignments.csv] T2 ON T2.TASK_UID=T1.UID "
    strSQL = strSQL & "WHERE WBS IS NULL "
    strSQL = strSQL & "OR OBS IS NULL "
    strSQL = strSQL & "OR CA IS NULL "
    strSQL = strSQL & "OR CAM IS NULL "
    strSQL = strSQL & "OR WP IS NULL "
    strSQL = strSQL & "OR EVT IS NULL "
    strSQL = strSQL & "GROUP BY T1.UID,WBS,OBS,CA,CAM,WP,EVT,EVP,BLS,BLF,[AS],AF "
    strSQL = strSQL & "HAVING SUM(T2.BLW)>0 OR SUM(T2.BLC)>0 "
  Else
    strSQL = "SELECT UID,"
    strSQL = strSQL & "IIF(ISNULL(WBS),'MISSING',WBS) AS [WBS],"
    strSQL = strSQL & "IIF(ISNULL(OBS),'MISSING',OBS) AS [OBS],"
    strSQL = strSQL & "IIF(ISNULL(CA),'MISSING',CA) AS [CA],"
    strSQL = strSQL & "IIF(ISNULL(CAM),'MISSING',CAM) AS [CAM],"
    strSQL = strSQL & "IIF(ISNULL(WP),'MISSING',WP) AS [WP],"
    strSQL = strSQL & "IIF(ISNULL(EVT),'MISSING',EVT) AS [EVT], "
    strSQL = strSQL & "BLW/60 AS [BLW],BLC,'Yes' AS PMB "
    strSQL = strSQL & "FROM [tasks.csv] "
    strSQL = strSQL & "WHERE SUMMARY='No' "
    strSQL = strSQL & "AND (BLW>0 OR BLC>0) "
    strSQL = strSQL & "AND (WBS IS NULL "
    strSQL = strSQL & "OR OBS IS NULL "
    strSQL = strSQL & "OR CA IS NULL "
    strSQL = strSQL & "OR CAM IS NULL "
    strSQL = strSQL & "OR WP IS NULL "
    strSQL = strSQL & "OR EVT IS NULL) "
  End If
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    '.Close
  End With
  'lngY = 100
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If dblScore = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = "MISSING METADATA"
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add "CPT01", strList
  myDECM_frm.lblStatus.Caption = "Checking for missing metadata...done."
  Application.StatusBar = "Checking for missing metadata...done."
  
  If lngX > 0 Then
    If Dir(Environ("tmp") & "\decm-cpt01.adtg") <> vbNullString Then Kill Environ("tmp") & "\decm-cpt01.adtg"
    oRecordset.Save Environ("tmp") & "\decm-cpt01.adtg", adPersistADTG 'maybe just save as a csv instead
    cptDECM_UPDATE_VIEW "CPT01", strList
    If MsgBox(Format(lngX, "#,##0") & " PMB task(s) have missing metadata!" & vbCrLf & vbCrLf & "Proceed anyway?", vbCritical + vbYesNo, "Missing Metadata") = vbNo Then
      myDECM_frm.cmdDone.Enabled = True
      myDECM_frm.cmdExport.Enabled = True
      blnProceed = False
      DumpRecordsetToExcel oRecordset
      GoTo exit_here
    Else
      blnProceed = True
      DumpRecordsetToExcel oRecordset
    End If
  Else
    blnProceed = True
  End If
  
exit_here:
  On Error Resume Next
  oRecordset.Close
  DECM_CPT01 = blnProceed
  DoEvents
  Exit Function
err_here:
  Call cptHandleErr("cptDECM_bas", "DECM_CPT01", Err, Erl)
  Resume exit_here
  
End Function

Sub DECM_05A101a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '05A101a - 1 CA : 1 OBS
  strMetric = "05A101a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting EVMS: 05A101a..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "1 CA : 1 OBS"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  'X = Count of CAs with more than one OBS element or no OBS elements assigned
  'Y = Total count of CAs
  'X/Y = 0%
  strSQL = "SELECT DISTINCT CA FROM tasks.csv WHERE CA IS NOT NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = .RecordCount
    .Close
  End With
  strSQL = "SELECT CA,COUNT(OBS) AS CountOfOBS "
  strSQL = strSQL & "FROM (SELECT DISTINCT CA,OBS FROM [tasks.csv]) "
  strSQL = strSQL & "WHERE CA IS NOT NULL "
  strSQL = strSQL & "GROUP BY CA "
  strSQL = strSQL & "HAVING COUNT(OBS)>1"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("CA") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_05A102a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  'Dim lngY As Long
  Dim dblScore As Double
  
  '05A102a - 1 CA : 1 CAM
  strMetric = "05A102a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "1 CA : 1 CAM"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = Count of CAs that have more than one CAM or no CAM assigned
  'Y = Total count of CAs
  'X/Y <= 5%
   strSQL = "SELECT DISTINCT CA FROM tasks.csv WHERE CA IS NOT NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = .RecordCount
    .Close
  End With
  strSQL = "SELECT CA,COUNT(CAM) AS CountOfCAM "
  strSQL = strSQL & "FROM (SELECT DISTINCT CA,CAM FROM [tasks.csv]) "
  strSQL = strSQL & "WHERE CA IS NOT NULL "
  strSQL = strSQL & "GROUP BY CA "
  strSQL = strSQL & "HAVING COUNT(CAM)>1"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("CA") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_05A103a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  'Dim lngY As Long
  Dim dblScore As Double
  
  '05A103a - 1 CA : 1 WBS
  strMetric = "05A103a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting EVMS: 05A103a..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "1 CA : 1 WBS"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  'X = Count of CAs with more than one WBS element or no WBS elements assigned
  'Y = Total count of CAs
  'X/Y = 0%
  strSQL = "SELECT DISTINCT CA FROM tasks.csv WHERE CA IS NOT NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = .RecordCount
    .Close
  End With
  strSQL = "SELECT CA,COUNT(WBS) AS CountOfWBS "
  strSQL = strSQL & "FROM (SELECT DISTINCT CA,WBS FROM [tasks.csv]) "
  strSQL = strSQL & "WHERE CA IS NOT NULL "
  strSQL = strSQL & "GROUP BY CA "
  strSQL = strSQL & "HAVING COUNT(WBS)>1"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("CA") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_CPT02(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  'bonus: 1 WP : 1 CA
  myDECM_frm.lblStatus.Caption = "Getting bonus metric CPT02..."
  Application.StatusBar = "Getting bonus metric CPT02..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = "CPT02"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "1 WP : 1 CA"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = count of incomplete WPs that have more than one CA or no CA assigned
  'Y = count of incomplete WPs
  strSQL = "SELECT WP,COUNT(CA) AS CountOfCA "
  strSQL = strSQL & "FROM (SELECT DISTINCT WP,CA FROM [tasks.csv] WHERE AF IS NULL) "
  strSQL = strSQL & "GROUP BY WP "
  strSQL = strSQL & "HAVING COUNT(CA)>1"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("WP") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT DISTINCT WP "
  strSQL = strSQL & "FROM [tasks.csv] "
  strSQL = strSQL & "WHERE WP IS NOT NULL AND AF IS NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = .RecordCount
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = lngX 'Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = lngX 'Format(dblScore, "0%")
  If dblScore = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription("CPT02")
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add "CPT02", strList
  myDECM_frm.lblStatus.Caption = "Getting bonus metric CPT02...done."
  Application.StatusBar = "Getting bonus metric CPT02...done."
  DoEvents
End Sub

Sub DECM_10A102a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strPP As String
  'Dim lngY As Long
  Dim lngX As Long
  Dim dblScore As Double
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oListObject As Excel.ListObject
  
  strPP = cptGetSetting("Integration", "PP")
  
  '10A102a - 1 WP : 1 EVT
  strMetric = "10A102a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting EVMS: 05A103a..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "1 WP : 1 EVT"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = count of incomplete WPs that have more than one EVT or no EVT assigned
  'Y = count of incomplete WPs
  'X/Y <= 5%

  'limit to incomplete WPs with PMB and either mixed or missing EVTs
  'discrete WPs are complete if BAC and BCWP are within $100 (or 1h)
  'LOE WPs are complete if BAC and BCWP are within $100 (or 1h) AND ETC < $100 (or 1h)
  'PPs and SLPPs are not included todo: unless a WP has a K and something else
  strSQL = "SELECT DISTINCT WP "
  strSQL = strSQL & "FROM("
  strSQL = strSQL & "    SELECT WP, Count(EVT) AS CountOfEVT" 'WP has mixed EVTs
  strSQL = strSQL & "    FROM ("
  strSQL = strSQL & "        SELECT T.WP, Iif(Isnull(T.EVT),'',T.EVT) AS EVT, SUM(T.BLW+T.BLC) AS BAC "
  strSQL = strSQL & "        FROM [tasks.csv] AS T "
  strSQL = strSQL & "        WHERE T.WP IS NOT NULL AND T.AF IS NULL AND T.EVT<>'" & strPP & "' "
  strSQL = strSQL & "        GROUP BY T.WP, T.EVT "
  strSQL = strSQL & "        HAVING SUM(T.BLW+T.BLC)>0 "
  strSQL = strSQL & "    )  AS PMB"
  strSQL = strSQL & "    GROUP BY WP"
  strSQL = strSQL & "    HAVING Count(EVT)>1"
  strSQL = strSQL & "    UNION"
  strSQL = strSQL & "    SELECT WP,Count(EVT) " 'WP has no EVTs
  strSQL = strSQL & "    FROM [tasks.csv] AS T "
  strSQL = strSQL & "    WHERE WP IS NOT NULL AND T.EVT IS NULL"
  strSQL = strSQL & "    GROUP BY WP"
  strSQL = strSQL & ") AS [10A102a]"

  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    If Not .EOF Then
      lngX = .RecordCount
      strList = ""
      If lngX > 0 Then
        'create report workbook
        On Error Resume Next
        Set oExcel = GetObject(, "Excel.Application")
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If oExcel Is Nothing Then Set oExcel = CreateObject("Excel.Application")
        Set oWorkbook = oExcel.Workbooks.Add
        Set oWorksheet = oWorkbook.Sheets(1)
        oWorksheet.Name = strMetric
        .MoveFirst
        Do While Not .EOF
          strList = strList & .Fields("WP") & ","
          .MoveNext
        Loop
        oWorksheet.[A1] = "X"
        oWorksheet.[A2] = lngX
        oWorksheet.[A3] = "WP(X)"
        oRecordset.MoveFirst
        oWorksheet.[A4].CopyFromRecordset oRecordset
      End If
      If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    End If
    .Close
  End With
  'limit to incomplete WPs with PMB
  'discrete WPs are complete if BAC and BCWP are within $100
  'LOE WPs are complete if BAC and BCWP are within $100 AND ETC < $100
  'PPs and SLPPs are not included
  strSQL = "SELECT T.WP,SUM(T.BLW+T.BLC) AS BAC "
  strSQL = strSQL & "FROM [tasks.csv] AS T "
  strSQL = strSQL & "WHERE T.WP IS NOT NULL AND T.AF IS NULL AND (T.EVT<>'" & strPP & "' OR T.EVT IS NULL) "
  strSQL = strSQL & "GROUP BY T.WP "
  strSQL = strSQL & "HAVING SUM(T.BLW+T.BLC)>0"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = .RecordCount
    If Not oExcel Is Nothing Then
      oWorksheet.[B1] = "Y"
      oWorksheet.[B2] = lngY
      oWorksheet.[B3] = "WP(Y)"
      oWorksheet.[B4].CopyFromRecordset oRecordset
      oWorksheet.Columns(3).Clear
      oWorksheet.[C2].FormulaR1C1 = "=R2C1/R2C2"
      oWorksheet.[C2].Style = "percent"
      oWorksheet.[A1:C2].HorizontalAlignment = xlCenter
      If Dir(Environ("tmp") & "\" & strMetric & ".xlsx") <> vbNullString Then Kill Environ("tmp") & "\" & strMetric & ".xlsx"
      oWorkbook.SaveAs Environ("tmp") & "\" & strMetric & ".xlsx"
      oWorkbook.Close True
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore < 0.05 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
  
exit_here:
  On Error Resume Next
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Exit Sub
err_here:
  cptHandleErr "cptDECM_bas", "DECM_10A102a", Err, Erl
  Resume exit_here
  
End Sub

Sub DECM_10A103a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean, blnFiscalExists As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oListObject As Excel.ListObject
  Dim oCell As Excel.Range
  
  '10A103a - 0/100 EVTs in one fiscal period
  strMetric = "10A103a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "0/100 EVTs in >1 period"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  
  'X = Count of 0-100 EVT incomplete WPs with more than one accounting period of budget
  'Y = Total count of 0-100 EVT incomplete WPs
  strSQL = "SELECT DISTINCT WP FROM [tasks.csv] "
  strSQL = strSQL & "WHERE WP IS NOT NULL "
  strSQL = strSQL & "AND AF IS NULL "
  strSQL = strSQL & "AND EVT='F'" 'todo: what about other tools or values?
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    lngY = .RecordCount
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  If lngY > 0 Then
    If blnFiscalExists Then
      Set oWorkbook = cptGetEVTAnalysis
      Set oWorksheet = oWorkbook.Sheets(1)
      Set oListObject = oWorksheet.ListObjects(1)
      lngY = oListObject.DataBodyRange.Rows.Count
      lngX = oWorksheet.Evaluate("COUNTIFS(Table1[FiscalPeriods],"">1"")")
      strList = ""
      If lngX > 0 Then
        For Each oCell In oListObject.ListColumns("WP").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells
          If InStr(strList, oCell.Value) = 0 Then strList = strList & oCell.Value & vbTab
        Next oCell
      End If
      oWorkbook.Close True
    Else 'blnFiscalExists
      lngX = lngY 'triggers failure
    End If
  Else
    lngX = 0
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore < 0.05 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
  'todo: bonus metrics if EVT_MS is used:
  'todo: EVT = "B" AND (EVT_MS MISSING OR EVT_MS NOT IN ({discrete})
  'todo: EVT = "B" AND EVT_MS = 0/100 AND FiscalPeriods > 1
  'todo: EVT = "B" AND EVT_MS = 50/50 AND FiscalPeriods > 2
  
  Set oWorkbook = Nothing
  Set oWorksheet = Nothing
  Set oListObject = Nothing
  Set oCell = Nothing
  
End Sub

Sub DECM_10A109b(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '10A109b - all WPs have budget
  strMetric = "10A109b"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "WPs w/o Budgets"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = Count of WPs/PPs/SLPPs with BAC = 0
  'Y = Total count of WPs/PPs/SLPPs
  strSQL = "SELECT DISTINCT WP FROM [tasks.csv] WHERE WP IS NOT NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = .RecordCount
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT WP,SUM(BLW) AS [BLW],SUM(BLC) AS [BLC] FROM [tasks.csv] "
  strSQL = strSQL & "WHERE WP IS NOT NULL "
  strSQL = strSQL & "GROUP BY WP "
  strSQL = strSQL & "HAVING SUM(BLW)=0 AND SUM(BLC)=0"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("WP") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore < 0.05 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_10A302b(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strDir As String
  Dim lngX As Long
  Dim dblScore As Double
  Dim oFSO As Scripting.FileSystemObject
  Dim oFile As Scripting.TextStream
  
  '10A302b - PPs with progress
  strMetric = "10A302b"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "PPs w/EVP > 0"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 2%"
  DoEvents
  strSQL = "SELECT DISTINCT WP FROM [tasks.csv] "
  strSQL = strSQL & "WHERE EVT='K' " 'todo: what about other values/tools
  oRecordset.Open strSQL, strCon, adOpenStatic, adLockReadOnly
  If oRecordset.EOF Then
    lngY = 0
    lngX = 0
    dblScore = 0
    strList = ""
  Else
    lngY = oRecordset.RecordCount
    oRecordset.Close
    strSQL = "SELECT WP,IIF(SUM(EVP)>0,TRUE,FALSE) AS HasProgress  FROM [tasks.csv] "
    strSQL = strSQL & "WHERE EVT='K' "
    strSQL = strSQL & "GROUP BY WP "
    strSQL = strSQL & "HAVING SUM(EVP)>0"
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If oRecordset.EOF Then
      lngX = 0
    Else
      lngX = oRecordset.RecordCount
    End If
    dblScore = Round(lngX / lngY, 2)
    strList = ""
    With oRecordset
      If oRecordset.RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          strList = strList & oRecordset(0) & ","
          .MoveNext
        Loop
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        strDir = Environ("tmp")
        Set oFile = oFSO.CreateTextFile(strDir & "\10A302b-x.csv", True)
        oRecordset.MoveFirst
        oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
        oFile.Close
      End If
    End With
    
  End If
  oRecordset.Close
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.02 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
  
  Set oFSO = Nothing
  Set oFile = Nothing
  
End Sub

Sub DECM_10A303a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strDir As String
  Dim lngX As Long
  Dim dblScore As Double
  Dim oFSO As Scripting.FileSystemObject
  Dim oFile As Scripting.TextStream
  
  '10A303a - all PPs have duration?
  strMetric = "10A303a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "PPs duration = 0"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 10%"
  DoEvents
  'we already have lngY
  If lngY = 0 Then
    lngX = 0
    dblScore = 0
    strList = ""
  Else
    strSQL = "SELECT WP,SUM(DUR) FROM [tasks.csv] "
    strSQL = strSQL & "WHERE EVT='K' "
    strSQL = strSQL & "GROUP BY WP "
    strSQL = strSQL & "HAVING SUM(DUR)=0"
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    strList = ""
    If oRecordset.EOF Then
      lngX = 0
    Else
      lngX = oRecordset.RecordCount
      With oRecordset
        .MoveFirst
        Do While Not .EOF
          strList = strList & oRecordset(0) & ","
          .MoveNext
        Loop
      End With
      Set oFSO = CreateObject("Scripting.FileSystemObject")
      strDir = Environ("tmp")
      Set oFile = oFSO.CreateTextFile(strDir & "\10A303a-x.csv", True)
      oRecordset.MoveFirst
      oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
      oFile.Close
    End If
    dblScore = Round(lngX / lngY, 2)
    oRecordset.Close
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.02 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
  
  Set oFSO = Nothing
  Set oFile = Nothing
  
End Sub

Sub DECM_11A101a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strDir As String
  Dim strFile As String
  Dim lngX As Long
  Dim lngFile As Long
  Dim dblScore As Double
  
  '11A101a - CA BAC = SUM(WP BAC)?
  strMetric = "11A101a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  'X = Sum of the absolute values of (CA BAC - the sum of its WP and PP budgets)
  'Y = Total program BAC
  'create segregated.csv
  If blnResourceLoaded Then
    strSQL = "SELECT "
    strSQL = strSQL & "    T1.CA, "
    strSQL = strSQL & "    T1.WP, "
    strSQL = strSQL & "    SUM(T3.[WP BLW]) AS [WP_BLW] "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    ( "
    strSQL = strSQL & "        [tasks.csv] T1 "
    strSQL = strSQL & "        INNER JOIN assignments.csv T2 ON T2.task_uid = T1.uid "
    strSQL = strSQL & "    ) "
    strSQL = strSQL & "    INNER JOIN ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            task_uid, "
    strSQL = strSQL & "            sum(blw / 60) AS [wp blw] "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            assignments.csv "
    strSQL = strSQL & "        GROUP BY "
    strSQL = strSQL & "            task_uid "
    strSQL = strSQL & "    ) AS t3 ON t3.task_uid = t1.uid "
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "    t1.ca, "
    strSQL = strSQL & "    t1.wp "
  Else
    strSQL = "SELECT CA,WP,SUM(BLW/60) AS [WP_BLW] "
    strSQL = strSQL & "FROM [tasks.csv] "
    strSQL = strSQL & "GROUP BY CA,WP"
  End If
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  If oRecordset.EOF Then
    myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...skipped."
    Application.StatusBar = "Getting " & strMetric & "...skipped."
    oRecordset.Close
    Exit Sub
  Else
    myDECM_frm.lboMetrics.AddItem
    myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "CA BAC = Sum(WP BAC)"
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 1%"
    DoEvents
  End If
  lngFile = FreeFile
  strFile = Environ("tmp") & "\segregated.csv"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngFile
  Print #lngFile, "CA,WP,WP_BLW,"
  oRecordset.MoveFirst
  Print #lngFile, oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
  Close #lngFile
  oRecordset.Close
  
  'create itemized.csv
  strSQL = "SELECT "
  strSQL = strSQL & "    t1.ca, "
  strSQL = strSQL & "    t2.[ca_bac], "
  strSQL = strSQL & "    sum(t3.[wp_bac]) AS [WP_BAC], "
  strSQL = strSQL & "    t2.[ca_bac] - sum(t3.[wp_bac]) AS [discrepancy] "
  strSQL = strSQL & "FROM "
  strSQL = strSQL & "    ( "
  strSQL = strSQL & "        segregated.csv t1 "
  strSQL = strSQL & "        LEFT JOIN ( "
  strSQL = strSQL & "            SELECT "
  strSQL = strSQL & "                ca, "
  strSQL = strSQL & "                sum([wp_blw]) AS [ca_bac] "
  strSQL = strSQL & "            FROM "
  strSQL = strSQL & "                segregated.csv "
  strSQL = strSQL & "            GROUP BY "
  strSQL = strSQL & "                ca "
  strSQL = strSQL & "        ) AS t2 ON t2.ca = t1.ca "
  strSQL = strSQL & "    ) "
  strSQL = strSQL & "    LEFT JOIN ( "
  strSQL = strSQL & "        SELECT "
  strSQL = strSQL & "            wp, "
  strSQL = strSQL & "            sum([wp_blw]) AS [wp_bac] "
  strSQL = strSQL & "        FROM "
  strSQL = strSQL & "            segregated.csv "
  strSQL = strSQL & "        GROUP BY "
  strSQL = strSQL & "            wp "
  strSQL = strSQL & "    ) AS t3 ON t3.wp = t1.wp "
  strSQL = strSQL & "GROUP BY "
  strSQL = strSQL & "    t1.ca, "
  strSQL = strSQL & "    t2.[ca_bac] "
  strSQL = strSQL & "HAVING "
  strSQL = strSQL & "    t2.[ca_bac] - sum(t3.[wp_bac]) <> 0 "
  
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  lngFile = FreeFile
  strFile = Environ("tmp") & "\itemized.csv"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngFile
  Print #lngFile, "CA,CA_BAC,WP_BAC,discrepancy,"
  If oRecordset.RecordCount > 0 Then
    oRecordset.MoveFirst
    Print #lngFile, oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
    Close #lngFile
    oRecordset.Close
    'get list of ca offenders
    strSQL = "SELECT DISTINCT CA FROM itemized.csv"
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    oRecordset.MoveFirst
    strList = ""
    Do While Not oRecordset.EOF
      If Not IsNull(oRecordset("CA")) Then
        strList = strList & oRecordset("CA") & ","
      End If
      oRecordset.MoveNext
    Loop
    oRecordset.Close
    strList = Left(strList, Len(strList) - 1) 'remove trailing comma
    strList = strList & ";" 'add separator between CAs and WPs
    'get list of wp offenders
    strSQL = "SELECT "
    strSQL = strSQL & "    wp, "
    strSQL = strSQL & "    count(ca) "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            DISTINCT wp, "
    strSQL = strSQL & "            ca "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            tasks.csv "
    strSQL = strSQL & "    ) "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    wp IS NOT NULL "
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "    wp "
    strSQL = strSQL & "HAVING "
    strSQL = strSQL & "    count(ca) > 1 "
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If oRecordset.RecordCount > 0 Then
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
        strList = strList & oRecordset("wp") & ","
        oRecordset.MoveNext
      Loop
    End If
  End If
  Close #lngFile
  oRecordset.Close
  
  'get delta as X
  strSQL = "SELECT sum(abs([discrepancy])) as [DELTA] from itemized.csv"
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  If oRecordset.RecordCount > 0 Then
    If Not IsNull(oRecordset("DELTA")) Then lngX = Round(oRecordset("DELTA"), 0) Else lngX = 0
  Else
    lngX = 0
  End If
  oRecordset.Close
  
  'get total as Y
  If blnResourceLoaded Then
    strSQL = "SELECT SUM(BLW/60) FROM assignments.csv"
  Else
    strSQL = "SELECT SUM(BLW/60) FROM tasks.csv"
  End If
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  If oRecordset.RecordCount > 0 Then
    lngY = Round(oRecordset(0), 0)
  Else
    lngX = 0
  End If
  oRecordset.Close
  
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.01 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A101a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  'LIMIT TO WHERE BAC>0
  'BCWP-BAC +/- 1h or $100 then 'complete'
  'note: due to the complications in calculating BCWP this metric relies on EVP<100
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strDir As String
  Dim strLOE As String
  Dim lngX As Long
  Dim dblScore As Double
  Dim oFSO As Scripting.FileSystemObject
  Dim oFile As Scripting.TextStream
  
  '06A101a - WPs Missing between IMS vs EV
  strLOE = cptGetSetting("Integration", "LOE")
  strMetric = "06A101a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "WPs IMS vs EV Tool"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  strSQL = "SELECT WP "
  strSQL = strSQL & "FROM [tasks.csv] "
  strSQL = strSQL & "WHERE EVP<100 AND EVT<>'" & strLOE & "' AND SUMMARY='No' "
  strSQL = strSQL & "GROUP BY WP "
  strSQL = strSQL & "HAVING Sum(BLW + BLC) > 0"
  oRecordset.Open strSQL, strCon, adOpenKeyset
  lngX = oRecordset.RecordCount 'pending upload
  lngY = oRecordset.RecordCount 'pending upload
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  strDir = Environ("tmp")
  Set oFile = oFSO.CreateTextFile(strDir & "\wp-ims.csv", True)
  If Not oRecordset.EOF Then oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
  oRecordset.Close
  oFile.Close
  FileCopy strDir & "\wp-ims.csv", strDir & "\wp-not-in-ev.csv"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
  
  Set oFSO = Nothing
  Set oFile = Nothing
  
End Sub

Sub DECM_06A204b(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strDir As String
  Dim strLOE As String
  Dim strLinks As String
  Dim lngX As Long
  Dim lngItem As Long
  Dim dblScore As Double
  Dim oDict As Scripting.Dictionary
  Dim vField As Variant
  
  '06A204b - Dangling Logic
  strMetric = "06A204b"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Dangling Logic"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  'myDECM_Frm.lboMetrics.Value = "06A204b"
  DoEvents
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 1) = DECM("06A204b")
  'Y = count incomplete Non-LOE tasks/activities & milestones
  'X = count of tasks with open starts or finishes
  'X/Y = 0%
  strSQL = "SELECT * FROM [tasks.csv] "
  strSQL = strSQL & "WHERE AF IS NULL "
  strSQL = strSQL & "AND (EVT<>'" & strLOE & "' OR EVT IS NULL) "
  strSQL = strSQL & "AND SUMMARY='No'"
  oRecordset.Open strSQL, strCon, adOpenKeyset
  lngY = oRecordset.RecordCount
  'start with this list - guilty until proven innocent
  Set oDict = CreateObject("Scripting.Dictionary")
  With oRecordset
    .MoveFirst
    Do While Not .EOF
      oDict.Add CStr(oRecordset("UID")), CStr(oRecordset("UID"))
      .MoveNext
    Loop
  End With
  oRecordset.Close
  
  'now do predecessors - guilty until proven innocent
  strSQL = "SELECT t.UID,t.DUR,p.[TYPE],p.[FROM] FROM [tasks.csv] t "
  strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT * FROM [links.csv]) p ON p.TO=t.UID "
  strSQL = strSQL & "WHERE t.SUMMARY='No' AND t.AF IS NULL AND (t.EVT<>'" & strLOE & "' OR t.EVT IS NULL)"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    .MoveFirst
    Do While Not .EOF
      If oRecordset("UID") <> "" And (oRecordset("TYPE") = "SS" Or oRecordset("TYPE") = "FS") And oRecordset("DUR") > 0 Then
        If oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Remove (CStr(oRecordset("UID")))
      ElseIf oRecordset("UID") <> "" And oRecordset("DUR") = 0 And Not IsNull(oRecordset("TYPE")) Then
        If oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Remove (CStr(oRecordset("UID")))
      End If
      .MoveNext
    Loop
  End With
  oRecordset.Close
  'extract the guilty to a string for later consolidation
  For lngItem = 0 To oDict.Count - 1
    strLinks = strLinks & oDict.Items(lngItem) & "," 'keep trailing comma, we're going to build on it
  Next lngItem
    
  'now do successors - guilty until proven innocent
  oDict.RemoveAll
  strSQL = "SELECT * FROM [tasks.csv] WHERE AF IS NULL AND (EVT<>'" & strLOE & "' OR EVT IS NULL) AND SUMMARY='No'"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    .MoveFirst
    Do While Not .EOF
      oDict.Add CStr(oRecordset("UID")), CStr(oRecordset("UID"))
      .MoveNext
    Loop
    .Close
  End With
  
  strSQL = "SELECT t.UID,t.DUR,s.[TYPE],s.[TO] FROM [tasks.csv] t "
  strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT * FROM [links.csv]) s ON s.[FROM]=t.UID "
  strSQL = strSQL & "WHERE t.SUMMARY='No' AND t.AF IS NULL AND (t.EVT<>'" & strLOE & "' OR t.EVT IS NULL)"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    .MoveFirst
    Do While Not .EOF
      If oRecordset("UID") <> "" And (oRecordset("TYPE") = "FF" Or oRecordset("TYPE") = "FS") And oRecordset("DUR") > 0 Then
        If oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Remove (CStr(oRecordset("UID")))
      ElseIf oRecordset("UID") <> "" And oRecordset("DUR") = 0 And Not IsNull(oRecordset("TYPE")) Then
        If oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Remove (CStr(oRecordset("UID")))
      End If
      .MoveNext
    Loop
    .Close
  End With
  
  'account for earliest/latest
  strSQL = "SELECT UID,BLS,BLF "
  strSQL = strSQL & "FROM [tasks.csv] "
  strSQL = strSQL & "WHERE AF IS NULL " 'incomplete
  strSQL = strSQL & "AND SUMMARY='No' " 'non-summary
  strSQL = strSQL & "AND (EVT <> '" & strLOE & "' OR EVT IS NULL) " 'non-LOE
  strSQL = strSQL & "AND (BLS = (SELECT MIN(BLS) FROM [tasks.csv]) " 'earliest BLS
  strSQL = strSQL & "OR BLF = (SELECT MAX(BLF) FROM [tasks.csv])) " 'latest BLF
  strSQL = strSQL & "ORDER BY BLS"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    .MoveFirst
    Do While Not .EOF
      If oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Remove (CStr(oRecordset("UID")))
      .MoveNext
    Loop
    .Close
  End With
  
  'extract the guilty to a string for later consolidation
  For lngItem = 0 To oDict.Count - 1
    strLinks = strLinks & oDict.Items(lngItem) & ","
  Next lngItem
  If Len(strLinks) > 0 Then strLinks = Left(strLinks, Len(strLinks) - 1)
  oDict.RemoveAll
  strList = ""
  For Each vField In Split(strLinks, ",")
    If Len(vField) > 0 And Not oDict.Exists(vField) Then
      oDict.Add vField, vField
      strList = strList & vField & ","
    End If
  Next vField
  lngX = oDict.Count
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
  
  Set oDict = Nothing
  
End Sub

Sub DECM_06A205a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strDir As String
  Dim strLOE As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '06A205a - Lags (what about leads?)
  strLOE = cptGetSetting("Integration", "LOE")
  strMetric = "06A205a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A205a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Lags"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 10%"
  DoEvents
  'X = count of incomplete tasks/activities & milestones with at least one lag in the pred logic
  'Y = count of incomplete tasks/activities & milestones in the IMS
  'X/Y <=10%
  'we already have lngY
  strSQL = "SELECT t.UID FROM [tasks.csv] t "
  strSQL = strSQL & "INNER JOIN (SELECT DISTINCT TO FROM [links.csv] WHERE LAG>0) p ON p.TO=t.UID " 'todo
  'todo: to include leads replace above with strSQL = strSQL & "INNER JOIN (SELECT DISTINCT TO FROM [links.csv] WHERE LAG<>0) p ON p.TO=t.UID "
  strSQL = strSQL & "WHERE t.SUMMARY='No' AND t.AF IS NULL AND (t.EVT<>'" & strLOE & "' OR t.EVT IS NULL) "
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = oRecordset.RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.1 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_CPT03(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strLOE As String
  Dim lngX As Long
  Dim dblScore As Double
  
  'CPT03 - leads
  strMetric = "CPT03"
  strLOE = cptGetSetting("Integration", "LOE")
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A205a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Leads"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = count of incomplete tasks/activities & milestones with at least one lead (negative lag) in the pred logic
  'Y = not used
  strSQL = "SELECT t.UID FROM [tasks.csv] t "
  strSQL = strSQL & "INNER JOIN (SELECT DISTINCT TO FROM [links.csv] WHERE LAG<0) p ON p.TO=t.UID " 'todo
  strSQL = strSQL & "WHERE t.SUMMARY='No' AND t.AF IS NULL AND (t.EVT<>'" & strLOE & "' OR t.EVT IS NULL) "
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = oRecordset.RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  'myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If lngX = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A208a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '06A208a - summary tasks with logic
  strMetric = "06A208a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A208a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Summary Logic"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = Count of summary tasks/activities with logic applied (# predecessors > 0 or # successors > 0)
  'X = 0
  strSQL = "SELECT DISTINCT T1.UID FROM [tasks.csv] T1 "
  strSQL = strSQL & "INNER JOIN [links.csv] T2 ON T2.FROM=T1.UID "
  strSQL = strSQL & "WHERE T1.SUMMARY='Yes'"
  lngX = 0
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
      lngX = .RecordCount
      strList = ""
      If lngX > 0 Then
        .MoveFirst
        Do While Not .EOF
          strList = strList & .Fields("UID") & ","
          .MoveNext
        Loop
      End If
    End If
    .Close
  End With
  'summary with successors
  strSQL = "SELECT DISTINCT T1.UID FROM [tasks.csv] T1 "
  strSQL = strSQL & "INNER JOIN [links.csv] T2 ON T2.TO=T1.UID "
  strSQL = strSQL & "WHERE T1.SUMMARY='Yes'"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
      lngX = lngX + .RecordCount
      If lngX > 0 Then
        .MoveFirst
        Do While Not .EOF
          strList = strList & .Fields("UID") & ","
          .MoveNext
        Loop
      End If
    End If
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If lngX = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A209a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strLOE As String
  Dim lngX As Long
  Dim dblScore As Double

  strLOE = cptGetSetting("Integration", "LOE")
  '06A209a - hard constraints
  strMetric = "06A209a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A209a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Hard Constraints"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  'X = count of incomplete tasks/activities & milestones with hard constraints
  'Y = count of incomplete tasks/activities & milestones
  'X/Y = 0%
  'we already have lngY
  strSQL = "SELECT UID FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND AF IS NULL " 'AND (EVT<>'" & strLOE & "' OR EVT IS NULL) "
  strSQL = strSQL & "AND (CONST='SNLT' OR CONST='FNLT' OR CONST='MSO' OR CONST='MFO')"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.1 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A210a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strLOE As String
  Dim lngX As Long
  Dim dblScore As Double
  Dim oDict As Scripting.Dictionary
  
  '06A210a - LOE Driving Discrete
  strLOE = cptGetSetting("Integration", "LOE")
  strMetric = "06A210a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A210a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "LOE Driving Discrete"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  'X = count of incomplete LOE tasks/activities in the IMS with at least one Non-LOE successor
  'Y = count of incomplete LOE tasks/activities in the IMS
  'X/Y = 0%
  'get Y
  strSQL = "SELECT UID FROM [tasks.csv] "
  strSQL = strSQL & "WHERE AF IS NULL "
  strSQL = strSQL & "AND SUMMARY='No' "
  strSQL = strSQL & "AND EVT='" & strLOE & "' "
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    If Not oRecordset.EOF Then
      lngY = oRecordset.RecordCount
    Else
      lngY = 0
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  'get X
  strSQL = "SELECT DISTINCT [FROM], PRED.EVT, [TO], SUCC.EVT "
  strSQL = strSQL & "FROM ([links.csv] LINKS "
  strSQL = strSQL & "    INNER JOIN (SELECT * FROM [tasks.csv]  "
  strSQL = strSQL & "        WHERE AF IS NULL  " 'incomplete
  strSQL = strSQL & "        AND SUMMARY = 'No'  " 'non-summary
  strSQL = strSQL & "        AND BLF <(SELECT MAX(BLF) FROM [tasks.csv] WHERE SUMMARY = 'No') " 'not last task/milestone/deliverable
  strSQL = strSQL & "        ) AS PRED ON PRED.UID = LINKS.FROM) "
  strSQL = strSQL & "    INNER JOIN (SELECT * FROM [tasks.csv] "
  strSQL = strSQL & "        WHERE AF IS NULL " 'incomplete
  strSQL = strSQL & "        AND SUMMARY = 'No' " 'non-summary
  strSQL = strSQL & "        AND BLF <(SELECT MAX(BLF) FROM [tasks.csv] WHERE SUMMARY = 'No') " 'not last task/milestone/deliverable
  strSQL = strSQL & "        ) AS SUCC ON SUCC.UID = LINKS.TO "
  'strSQL = strSQL & "WHERE (PRED.EVT = '" & strLOE & "' OR PRED.EVT IS NULL) " 'LOE/Null Pred EVT
  strSQL = strSQL & "WHERE PRED.EVT = '" & strLOE & "' "
  strSQL = strSQL & "AND (SUCC.EVT <> '" & strLOE & "' OR SUCC.EVT IS NULL) " 'LOE/Null Succ EVT
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    Set oDict = CreateObject("Scripting.Dictionary")
    strList = ""
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        If Not oDict.Exists(CStr(oRecordset("FROM"))) Then oDict.Add CStr(oRecordset("FROM")), CStr(oRecordset("FROM"))
        strList = strList & .Fields("FROM") & "," & .Fields("TO") & "," 'includes guilty successors
        .MoveNext
      Loop
    End If
    lngX = oDict.Count
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList 'todo: need guilty link too
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
  
  Set oDict = Nothing
  
End Sub

Sub DECM_06A211a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim strLOE As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '06A211a - High Float
  '06A211a - High Float todo: refine TS into effective days (elapsed, etc)
  '06A211a - High Float todo: need rationale; user can mark 'acceptable'
  '06A211a - High Float todo: allow user input for lngX
  strLOE = cptGetSetting("Integration", "LOE")
  strMetric = "06A211a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A211a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "High Float"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 20%"
  DoEvents
'  X = count of high total float Non-LOE tasks/activities & milestones sampled with inadequate rationale
'  Y = count of high total float Non-LOE tasks/activities & milestones sampled
'  X/Y <= 20%
  strSQL = "SELECT UID,ROUND(TS/" & CLng(60 * ActiveProject.HoursPerDay) & ",2) AS HTF "
  strSQL = strSQL & "FROM [tasks.csv] "
  strSQL = strSQL & "WHERE EVT<>'" & strLOE & "' "
  strSQL = strSQL & "GROUP BY UID,ROUND(TS/480,2) "
  strSQL = strSQL & "HAVING ROUND(TS/480,2)>44 "
  strList = ""
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = oRecordset.RecordCount
    lngY = oRecordset.RecordCount
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.2 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A212a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '06A212a - out of sequence
  'todo: don't do Excel; cut title lbo in half and
  'todo: add list of pairs - user can pull up whatever add'l info is wanted
  strMetric = "06A212a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Out of Sequence"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = Count of out of sequence conditions
  strList = cptGetOutOfSequence(myDECM_frm) 'function returns lngX|uid vbtab uid vbtab uid
  lngX = CLng(Split(strList, "|")(0))
  strList = Split(strList, "|")(1)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 4) = ""
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If lngX = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  ElseIf lngX > 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric) 'todo: see workbook
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A401a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim lngTargetUID As Long
  Dim dblScore As Double
  Dim oTask As MSProject.Task
  Dim dtConstraint As Date
  Dim lngConstraintType As Long
  Dim lngTargetTotalSlack As Long
  
  '06A401a - critical path (constraint method)
  strMetric = "06A401a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Critical Path"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents

  'get targetUID
  lngTargetUID = cptDECMGetTargetUID()
  If lngTargetUID = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Critical Path - SKIPPED"
    'todo: remove it? give a bad score?
    Exit Sub 'was GoTo skip_06A401a
  Else
    Set oTask = ActiveProject.Tasks.UniqueID(lngTargetUID)
    'save existing constraint
    If IsDate(oTask.ConstraintDate) Then dtConstraint = oTask.ConstraintDate
    lngConstraintType = oTask.ConstraintType
    'replace with DateSubtract("yyyy",-10,finish)
    oTask.ConstraintType = pjMFO
    oTask.ConstraintDate = DateAdd("yyyy", -10, oTask.Finish)
    'get total slack
    lngTargetTotalSlack = oTask.TotalSlack
    'get list of primary driving path UIDs
    strList = ""
    For Each oTask In ActiveProject.Tasks
      If oTask Is Nothing Then GoTo next_critical_task
      If oTask.Summary Then GoTo next_critical_task
      If Not oTask.Active Then GoTo next_critical_task
      If oTask.TotalSlack = lngTargetTotalSlack Then
        strList = strList & oTask.UniqueID & ","
      End If
next_critical_task:
    Next oTask
    'restore constraint
    If dtConstraint > 0 Then
      ActiveProject.Tasks.UniqueID(lngTargetUID).ConstraintDate = dtConstraint
    Else
      ActiveProject.Tasks.UniqueID(lngTargetUID).ConstraintDate = "NA"
    End If
    ActiveProject.Tasks.UniqueID(lngTargetUID).ConstraintType = lngConstraintType
    If Len(strList) > 0 Then
      lngX = UBound(Split(strList, ","))
    Else
      lngX = 0
    End If
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If dblScore = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, lngTargetUID & "|" & strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
  Set oTask = Nothing
End Sub

Sub DECM_06A501a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '06A501a - baselines
  strMetric = "06A501a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A501a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Baselines"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = Count of tasks/activities & milestones without baseline dates
  'Y = Total count of tasks/activities & milestones
  'X/Y <= 5%
  strSQL = "SELECT UID,BLS,BLF FROM [tasks.csv] WHERE SUMMARY='No'"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = oRecordset.RecordCount
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT UID,BLS,BLF FROM [tasks.csv] WHERE SUMMARY='No' AND (BLS IS NULL OR BLF IS NULL)"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = oRecordset.RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A505a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '06A505a - In-Progress Tasks Have AS
  strMetric = "06A505a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A505a"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "In-Progress Tasks w/o Actual Start"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = count of in-progress tasks/activities & milestones with no actual start date
  'Y = count of in-progress tasks/activities & milestones
  'X/Y <= 5%
  strSQL = "SELECT UID,EVP,[AS] FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND EVP<100 AND EVP>0 "
  With oRecordset
    If .State Then .Close
    .Open strSQL, strCon, adOpenKeyset
    lngY = oRecordset.RecordCount
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT UID,EVP,[AS] FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND EVP<100 AND EVP>0 "
  strSQL = strSQL & "AND [AS] IS NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A505b(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '06A505b - Complete Tasks Have AF
  strMetric = "06A505b"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A505b"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Complete Tasks w/o Actual Finish"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = count of complete tasks/activities & milestones with no actual finish date
  'Y = count of complete tasks/activities & milestones
  'X/Y <= 5%
  strSQL = "SELECT UID,EVP,AF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND EVP=100 "
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = oRecordset.RecordCount
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT UID,EVP,AF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND EVP=100 AND AF IS NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = oRecordset.RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A506a(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean, dtStatus As Date)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '06A506a - bogus actuals
  strMetric = "06A506a"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Bogus Actuals"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  'myDECM_Frm.lboMetrics.Value = "06A506a"
  DoEvents
  'X = count of tasks/activities & milestones with either actual start or actual finish after status date
  'Y = count of tasks/activities & milestones with an actual start date
  'X/Y <= 5%
  strSQL = "SELECT UID,[AS],AF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE [AS] IS NOT NULL OR AF IS NOT NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = oRecordset.RecordCount
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT UID,[AS],AF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE ([AS]>#" & dtStatus & "# OR AF>#" & dtStatus & "#)"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Sub DECM_06A506b(ByRef oDECM As Scripting.Dictionary, ByRef myDECM_frm As cptDECM_frm, strCon As String, ByRef oRecordset As ADODB.Recordset, blnDumpToExcel As Boolean, dtStatus As Date)
  Dim strMetric As String
  Dim strSQL As String
  Dim strList As String
  Dim lngX As Long
  Dim dblScore As Double
  
  '06A506b - invalid forecast
  strMetric = "06A506b"
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "..."
  Application.StatusBar = "Getting " & strMetric & "..."
  myDECM_frm.lboMetrics.AddItem
  myDECM_frm.lboMetrics.TopIndex = myDECM_frm.lboMetrics.ListCount - 1
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 0) = strMetric
  'myDECM_Frm.lboMetrics.Value = "06A506b"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 1) = "Invalid Forecast"
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = Count of incomplete tasks/activities & milestones with either forecast start or forecast finish before the status date
  'X = 0
  strSQL = "SELECT UID,FS,FF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE ((FS<#" & dtStatus & "# AND [AS] IS NULL) "
  strSQL = strSQL & "OR (FF<#" & dtStatus & "# AND AF IS NULL))"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 4) = lngX there is no Y
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If lngX = 0 Then
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  myDECM_frm.lboMetrics.List(myDECM_frm.lboMetrics.ListCount - 1, 7) = cptGetDECMDescription(strMetric)
  'myDECM_Frm.lboMetrics.List(myDECM_Frm.lboMetrics.ListCount - 1, 8) = strList
  oDECM.Add strMetric, strList
  myDECM_frm.lblStatus.Caption = "Getting " & strMetric & "...done."
  Application.StatusBar = "Getting " & strMetric & "...done."
  DoEvents
End Sub

Function DECM(ByRef myDECM_frm As cptDECM_frm, strDECM As String, Optional blnNotify As Boolean = False) As Double
  Dim oTask As MSProject.Task
  Dim oLinks As Scripting.Dictionary
  Dim oLink As TaskDependency
  Dim lngX As Long
  Dim strLinks As String
  
  'If Not cptValidMap Then GoTo exit_here
  
  Select Case strDECM
    Case "06A204b" 'dangling logic
      
      ActiveWindow.TopPane.Activate
      FilterClear
      GroupClear
      OptionsViewEx DisplaySummaryTasks:=True
      OutlineShowAllTasks
    
      Set oLinks = CreateObject("Scripting.Dictionary")
      oLinks.Add 0, "FF"
      oLinks.Add 1, "FS"
      oLinks.Add 2, "SF"
      oLinks.Add 3, "SS"
      
      cptSpeed True
      
      lngEVT = CLng(Split(cptGetSetting("Integration", "EVT"), "|")(0))
      For Each oTask In ActiveProject.Tasks
        If oTask Is Nothing Then GoTo next_task
        'todo: external?
        If Not oTask.Active Then GoTo next_task
        oTask.Marked = False
        If oTask.Summary Then GoTo next_task
        If oTask.GetField(lngEVT) = "A" Then GoTo next_task
        If IsDate(oTask.ActualFinish) Then GoTo next_task
        lngY = lngY + 1
        'whether task or milestone, must have PE and SE
        If oTask.PredecessorTasks.Count = 0 Then
          lngX = lngX + 1
          oTask.Marked = True
          GoTo next_task
        End If
        If oTask.SuccessorTasks.Count = 0 Then
          lngX = lngX + 1
          oTask.Marked = True
          GoTo next_task
        End If
        If oTask.Duration > 0 Then 'examine tasks
          strLinks = ""
          For Each oLink In oTask.TaskDependencies
            If oLink.To = oTask Then 'examine predecessors
              strLinks = strLinks & oLinks(oLink.Type) & ","
            End If
          Next oLink
          If InStr(strLinks, "FS") = 0 And InStr(strLinks, "SS") = 0 Then
            lngX = lngX + 1
            oTask.Marked = True
            GoTo next_task
          End If
          strLinks = ""
          For Each oLink In oTask.TaskDependencies
            If oLink.From = oTask Then 'examine successors
              strLinks = strLinks & oLinks(oLink.Type) & ","
            End If
          Next oLink
          If InStr(strLinks, "FS") = 0 And InStr(strLinks, "FF") = 0 Then
            lngX = lngX + 1
            oTask.Marked = True
            GoTo next_task
          End If
        End If
next_task:
      Next oTask
      
      cptSpeed False
      
      ActiveWindow.TopPane.Activate
      OptionsViewEx DisplaySummaryTasks:=True
      OutlineShowAllTasks
      OptionsViewEx DisplaySummaryTasks:=False
      SetAutoFilter "Marked", pjAutoFilterFlagYes
      
      If blnNotify Then MsgBox "X: " & lngX & vbCrLf & "Y: " & lngY & vbCrLf & "X/Y = " & Format(lngX / lngY, "0%"), vbInformation + vbOKOnly, "06A204b"
      myDECM_frm.txtTitle = "X: " & lngX & vbCrLf & "Y: " & lngY
      'X = count of incomplete Non-LOE tasks/activities & milestones in the IMS WITH open starts or finishes
      'Y = count of incomplete Non-LOE tasks/activities & milestones in the IMS
      DECM = Round(lngX / lngY, 2)
  End Select

exit_here:
  Set oTask = Nothing
  Set oLink = Nothing
  Set oLinks = Nothing
  Exit Function

End Function

Private Sub DumpRecordsetToExcel(ByRef oRecordset As ADODB.Recordset)
  'objects
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  'strings
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.APplication")
  End If
  oExcel.Visible = True
  
  Set oExcel = GetObject(, "Excel.Application")
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "AD HOC"
  For lngItem = 0 To oRecordset.Fields.Count - 1
    oWorksheet.Cells(1, lngItem + 1) = oRecordset.Fields(lngItem).Name
  Next lngItem
  oWorksheet.[A2].Select
  oRecordset.MoveFirst
  oWorksheet.[A2].CopyFromRecordset oRecordset
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True
  oWorksheet.[A1].AutoFilter
  oWorksheet.Columns.AutoFit

exit_here:
  On Error Resume Next
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("basDECM_bas", "DumpRecordsetToExcel", Err, Erl)
  Resume exit_here

End Sub

Sub opencsv(strFileName)
  Shell "notepad.exe """ & Environ("tmp") & "\" & strFileName & """", vbNormalFocus
End Sub
Sub cptDECM_EXPORT(ByRef myDECM_frm As cptDECM_frm, Optional blnDetail As Boolean = False)
  'objects
  Dim oShading As Object
  Dim oBorders As Object
  Dim o06A101a As Excel.Workbook
  Dim o10A103a As Excel.Workbook
  Dim oRecordset As ADODB.Recordset
  Dim oTasks As MSProject.Tasks
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oRange As Excel.Range
  Dim oCell As Excel.Range
  'strings
  Dim strMetric As String
  Dim strResult As String
  Dim strDir As String
  Dim strCon As String
  Dim strSQL As String
  Dim strLOE As String
  'longs
  Dim lngItem As Long
  Dim lngField As Long
  Dim lngFirstRow As Long
  Dim lngLastRow As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnResourceAssignments As Boolean
  'variants
  Dim vSetting As Variant
  'dates
  
  cptSpeed True
  
  blnDetail = MsgBox("Include Details?", vbQuestion + vbYesNo, "Detailed Results") = vbYes

  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Application.StatusBar = "Opening Excel..."
    Set oExcel = CreateObject("Excel.Application")
    Application.StatusBar = ""
  End If
  
  Set oWorkbook = oExcel.Workbooks.Add
  oExcel.Visible = True 'just in case
  oExcel.WindowState = xlMinimized 'xlMaximized
  oWorkbook.Activate
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Activate
  oWorksheet.Name = "DECM Dashboard"
  oWorksheet.[A1:I1] = myDECM_frm.lboHeader.List
  oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A2].Offset(myDECM_frm.lboMetrics.ListCount - 1, myDECM_frm.lboMetrics.ColumnCount - 1)) = myDECM_frm.lboMetrics.List
  oExcel.ActiveWindow.Zoom = 85
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).HorizontalAlignment = xlLeft
  oWorksheet.[G1].Value = "RESULT"
  With oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1048576].End(xlUp))
    .Font.Name = "Calibri"
    .Font.Size = 11
    .HorizontalAlignment = xlCenter
    .Columns.AutoFit
  End With
  
  oWorksheet.Columns(1).HorizontalAlignment = xlLeft
  oWorksheet.Columns(2).HorizontalAlignment = xlLeft
  'oWorksheet.Columns(8).HorizontalAlignment = xlLeft
  oWorksheet.Columns("H:I").Delete
  With oWorksheet.Range(oWorksheet.[G2], oWorksheet.[G1048576].End(xlUp))
    .Replace what:=strPass, Replacement:="2", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    .Replace what:=strFail, Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    .FormatConditions.AddIconSetCondition
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1)
      .ReverseOrder = False
      .ShowIconOnly = True
      .IconSet = oWorkbook.IconSets(xl3Symbols)
    End With
    With .FormatConditions(1).IconCriteria(2)
      .Type = xlConditionValueNumber
      .Value = 1
      .Operator = 7
    End With
    With .FormatConditions(1).IconCriteria(3)
      .Type = xlConditionValueNumber
      .Value = 2
      .Operator = 7
    End With
  End With
  
  If blnDetail And Not myDECM_frm.chkUpdateView Then myDECM_frm.chkUpdateView = True
  strDir = Environ("tmp")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  Set oRecordset = CreateObject("ADODB.Recordset")
  If blnDetail Then
    With myDECM_frm
      For lngItem = 0 To .lboMetrics.ListCount - 1
        .lboMetrics.Value = .lboMetrics.List(lngItem)
        .lboMetrics.Selected(lngItem) = True
        .lboMetrics.ListIndex = lngItem
        strMetric = .lboMetrics.List(lngItem)
        strResult = .lboMetrics.List(lngItem, 6)
        Application.StatusBar = "Exporting " & strMetric & "..."
        .lblStatus.Caption = "Exporting " & strMetric & "..."
        .lblProgress.Width = (lngItem / .lboMetrics.ListCount) * .lblStatus.Width
        DoEvents
        Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
        oWorksheet.Activate
        oWorksheet.Name = strMetric
        oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
        oWorksheet.[A2] = strMetric & ": " & .lboMetrics.List(lngItem, 1)
        oWorksheet.Cells.Font.Name = "Calibri"
        oWorksheet.Cells.Font.Size = 11
        oWorksheet.Cells.WrapText = False
        If strResult = strFail Then
          oWorksheet.Tab.Color = 192
        Else
          oWorksheet.Tab.Color = 5287936
        End If
        oExcel.ActiveWindow.Zoom = 85
        oExcel.ActiveWindow.DisplayGridlines = False
        strLOE = cptGetSetting("Integration", "LOE")
        If strMetric = "CPT01" Then 'missing metadata
          If strResult = strFail Then
            If Dir(Environ("tmp") & "\decm-cpt01.adtg") <> vbNullString Then
              oRecordset.Open Environ("tmp") & "\decm-cpt01.adtg"
              For lngField = 0 To oRecordset.Fields.Count - 1
                oWorksheet.Cells(3, lngField + 1) = oRecordset.Fields(lngField).Name
              Next lngField
              oWorksheet.[A4].CopyFromRecordset oRecordset
              oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
              oRecordset.Close
              oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
              oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).FormatConditions.Add xlCellValue, xlEqual, "=""MISSING"""
              oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).FormatConditions(1).SetFirstPriority
              oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).FormatConditions(1).Font.Color = -16383844
              oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).FormatConditions(1).Font.TintAndShade = 0
              oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
              oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).FormatConditions(1).Interior.Color = 13551615
              oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).FormatConditions(1).Interior.TintAndShade = 0
              cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
              cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
              cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            End If
          End If
        ElseIf strMetric = "05A101a" Then '1 CA : 1 OBS
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT CA,CAM,OBS FROM [tasks.csv] "
            strSQL = strSQL & "WHERE CA IN (" & Chr(34) & Replace(oDECM(strMetric), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ") "
            strSQL = strSQL & "ORDER BY CA"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:C3] = Split("CA,CAM,OBS", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = .lboMetrics.List(lngItem, 3)
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "05A102a" Then '1 CA : 1 CAM
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT CA,CAM FROM [tasks.csv] "
            strSQL = strSQL & "WHERE CA IN (" & Chr(34) & Replace(oDECM(strMetric), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ") "
            strSQL = strSQL & "ORDER BY CA"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:B3] = Split("CA,CAM", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = .lboMetrics.List(lngItem, 3)
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "05A103a" Then '1 CA : 1 WBS
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT CA,CAM,WBS FROM [tasks.csv] "
            strSQL = strSQL & "WHERE CA IN (" & Chr(34) & Replace(oDECM(strMetric), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ") "
            strSQL = strSQL & "ORDER BY CA"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:C3] = Split("CA,CAM,WBS", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = .lboMetrics.List(lngItem, 3)
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "CPT02" Then '1 WP : 1 CA
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT DISTINCT WP,CA,CAM "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE WP IN (" & Chr(34) & Replace(oDECM(strMetric), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ") "
            strSQL = strSQL & "ORDER BY WP"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:C3] = Split("WP,CA,CAM", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = .lboMetrics.List(lngItem, 3)
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "10A102a" Then '1 WP : 1 EVT
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT DISTINCT CAM,WP,EVT "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE WP IN (" & Chr(34) & Replace(oDECM(strMetric), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:C3] = Split("CAM,WP,EVT", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = .lboMetrics.List(lngItem, 3)
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "10A103a" Then '0/100 in >1 fiscal period
          If Len(oDECM(strMetric)) > 0 Then
            On Error Resume Next
            Set o10A103a = oExcel.Workbooks.Open(strDir & "\10A103a.xlsx")
            If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
            If o10A103a Is Nothing Then
              'todo: what if it doesn't exist?
              strSQL = "SELECT CAM,WP,EVT,MIN(BLS),MAX(BLF) FROM [tasks.csv] WHERE WP IN () ORDER BY CAM"
            Else
              If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
              'replace current worksheet with worksheet from saved workbook
              oExcel.DisplayAlerts = False
              oWorksheet.Delete
              oExcel.DisplayAlerts = True
              o10A103a.Sheets(1).Copy After:=oWorkbook.Sheets(oWorkbook.Sheets.Count)
              o10A103a.Close True
              Set oWorksheet = oWorkbook.Sheets("10A103a")
              oWorksheet.Rows("1:2").Insert
              oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
              oWorksheet.[A2].Value = strMetric & ": 0/100 WPs in more than 1 fiscal period"
              oExcel.ActiveWindow.Zoom = 85
              oExcel.ActiveWindow.DisplayGridlines = False
              cptAddBorders oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight))
              cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
              cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
              cptAddBorders oWorksheet.Range(oWorksheet.[H3].End(xlDown), oWorksheet.[H3].End(xlToRight))
              cptAddBorders oWorksheet.Range(oWorksheet.[H3], oWorksheet.[H3].End(xlToRight))
              cptAddShading oWorksheet.Range(oWorksheet.[H3], oWorksheet.[H3].End(xlToRight))
              If strResult = strFail Then
                oWorksheet.Tab.Color = 192
              Else
                oWorksheet.Tab.Color = 5287936
              End If
            End If
          End If
        ElseIf strMetric = "10A109b" Then 'WPs w/o budgets
          If Len(oDECM(strMetric)) > 0 Then
            'todo: add BLW,BLC,RW,RC to tasks.csv
            'todo: FilterByClipboard add fields
            strSQL = "SELECT DISTINCT CAM,WP,0 AS BLW,0 AS BLC FROM [tasks.csv] "
            strSQL = strSQL & "WHERE WP IN (" & Chr(34) & Replace(oDECM(strMetric), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ") "
            strSQL = strSQL & "ORDER BY CAM,WP"
            oRecordset.Open strSQL, strCon, adOpenKeyset
            oWorksheet.[A3:D3] = Split("CAM,WP,BLW,BLC", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "10A302b" Then 'PPs w/EVP > 0
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT CAM,WP,EVT,EVP "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE WP IN (" & Chr(34) & Replace(oDECM(strMetric), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:D3] = Split("CAM,WP,EVT,EVP", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "10A303a" Then 'PPs duration = 0
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT CAM,WP,EVT,DUR "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE WP IN (" & Chr(34) & Replace(oDECM(strMetric), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:D3] = Split("CAM,WP,EVT,DUR", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            oExcel.ActiveWindow.DisplayGridlines = False
          End If
        ElseIf strMetric = "11A101a" Then 'CA BAC = Sum(WP BAC)
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT CAM,CA,SUM(BLW)/60,SUM(BLC) FROM [tasks.csv] "
            strSQL = strSQL & "WHERE CA IN(" & Chr(34) & Replace(Split(oDECM(strMetric), ";")(0), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ")"
            strSQL = strSQL & "GROUP BY CAM,CA "
            strSQL = strSQL & "ORDER BY CAM,CA"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:D3] = Split("CAM,CA,BLW,BLC", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            'todo: highlight duplicate CA
            oWorksheet.Columns(5).ColumnWidth = 1
            strSQL = "SELECT CAM,CA,WP,SUM(BLW)/60,SUM(BLC) FROM [tasks.csv] "
            strSQL = strSQL & "WHERE CA IN(" & Chr(34) & Replace(Split(oDECM(strMetric), ";")(0), ",", Chr(34) & "," & Chr(34)) & Chr(34) & ")"
            strSQL = strSQL & "GROUP BY CAM,CA,WP "
            strSQL = strSQL & "ORDER BY CAM,CA,WP"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[F3:J3] = Split("CAM,CA,WP,BLW,BLC", ",")
            oWorksheet.[F4].CopyFromRecordset oRecordset
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[F3], oWorksheet.[F3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[F3], oWorksheet.[F3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[F3], oWorksheet.[F3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[F3], oWorksheet.[F3].End(xlToRight))
          End If
        ElseIf strMetric = "06A101a" Then 'WPs IMS vs EV Tool
          Set o06A101a = Nothing
          On Error Resume Next
          Set o06A101a = oExcel.Workbooks.Open(strDir & "\06A101a.xlsx")
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If o06A101a Is Nothing Then
            'run queries
            oWorksheet.[A3].Value = "NOT IN IMS:"
            If Dir(strDir & "\wp-not-in-ims.csv") <> vbNullString Then
              strSQL = "SELECT * FROM [wp-not-in-ims.csv]"
              oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
              oWorksheet.[A4].CopyFromRecordset oRecordset
              oRecordset.Close
            End If
            oWorksheet.[C3].Value = "NOT IN EV TOOL:"
            If Dir(strDir & "\wp-not-in-ev.csv") <> vbNullString Then
              strSQL = "SELECT * FROM [wp-not-in-ev.csv]"
              oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
              oWorksheet.[C4].CopyFromRecordset oRecordset
              oRecordset.Close
            End If
            oWorksheet.Columns.AutoFit
          Else
            If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
            'replace current worksheet with worksheet from saved workbook
            oExcel.DisplayAlerts = False
            oWorksheet.Delete
            oExcel.DisplayAlerts = True
            o06A101a.Sheets(1).Copy After:=oWorkbook.Sheets(oWorkbook.Sheets.Count)
            o06A101a.Close True
            Set oWorksheet = oWorkbook.Sheets("06A101a")
            oWorksheet.Rows("1:2").Insert
            oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
            oWorksheet.[A2].Value = strMetric & ": WPs in IMS vs EV Tool"
          End If
        ElseIf strMetric = "06A204b" Then 'Dangling Logic
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT UID,CAM,TASK_NAME FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset
            oWorksheet.[A3:C3] = Split("UID,CAM,TASK_NAME", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A205a" Then 'Lags
          If Len(oDECM(strMetric)) > 0 Then 'returns successor UID
            strSQL = "SELECT DISTINCT T2.CAM,T1.[FROM],T2.TASK_NAME,T1.LAG/480,T1.TO,T3.TASK_NAME "
            strSQL = strSQL & "FROM ([links.csv] T1 "
            strSQL = strSQL & "LEFT JOIN [tasks.csv] T2 ON T2.UID=T1.[FROM]) "
            strSQL = strSQL & "LEFT JOIN [tasks.csv] T3 ON T3.UID=T1.TO "
            strSQL = strSQL & "WHERE T1.To IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY T2.CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:F3] = Split("CAM,FROM UID,FROM TASK,LAG,TO UID,TO TASK", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            oExcel.ActiveWindow.DisplayGridlines = False
          End If
        ElseIf strMetric = "CPT03" Then 'Leads
          If Len(oDECM(strMetric)) > 0 Then 'returns successor UID
            strSQL = "SELECT DISTINCT T2.CAM,T1.[FROM],T2.TASK_NAME,T1.LAG/480,T1.TO,T3.TASK_NAME "
            strSQL = strSQL & "FROM ([links.csv] T1 "
            strSQL = strSQL & "LEFT JOIN [tasks.csv] T2 ON T2.UID=T1.[FROM]) "
            strSQL = strSQL & "LEFT JOIN [tasks.csv] T3 ON T3.UID=T1.TO "
            strSQL = strSQL & "WHERE T1.To IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY T2.CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:F3] = Split("CAM,FROM UID,FROM TASK,LAG,TO UID,TO TASK", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A208a" Then 'Summary Logic
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT UID,CAM,TASK_NAME,SUMMARY,'','' FROM [tasks.csv] WHERE UID IN (" & oDECM(strMetric) & ") "
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:F3] = Split("UID,CAM,TASK NAME,SUMMARY,UID PREDS,UID SUCCS", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.MoveFirst
            Do While Not oRecordset.EOF
              oWorksheet.Cells(oRecordset.AbsolutePosition + 3, 5) = Replace(ActiveProject.Tasks.UniqueID(oRecordset(0)).UniqueIDPredecessors, ",", ";")
              oWorksheet.Cells(oRecordset.AbsolutePosition + 3, 6) = Replace(ActiveProject.Tasks.UniqueID(oRecordset(0)).UniqueIDSuccessors, ",", ";")
              oRecordset.MoveNext
            Loop
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3].End(xlToRight), oWorksheet.[A3].End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A209a" Then 'Hard Constraints
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT UID,CAM,TASK_NAME,CONST "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:D3] = Split("UID,CAM,TASK NAME,CONSTRAINT", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3].End(xlToRight), oWorksheet.[A3].End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A210a" Then 'LOE Driving Discrete
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT DISTINCT T1.[FROM],T2.TASK_NAME,T2.EVT,T1.TO,T3.TASK_NAME,T3.EVT "
            strSQL = strSQL & "FROM ([links.csv] T1 "
            strSQL = strSQL & "LEFT JOIN [tasks.csv] T2 ON T2.UID=T1.[FROM]) "
            strSQL = strSQL & "LEFT JOIN [tasks.csv] T3 ON T3.UID=T1.TO "
            strSQL = strSQL & "WHERE [FROM] IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "AND T2.EVT='" & strLOE & "' "
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:F3] = Split("FROM UID,FROM TASK NAME,FROM EVT,TO UID,TO TASK NAME,TO EVT", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3].End(xlToRight), oWorksheet.[A3].End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
          'todo: dashboard X should be formulae so users can refine/correct
        ElseIf strMetric = "06A211a" Then 'High Float
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT UID,CAM,TASK_NAME,TS/480 FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset
            oWorksheet.[A3:D3] = Split("UID,CAM,TASK_NAME,TOTAL SLACK", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A212a" Then 'Out of Sequence
          If Len(oDECM(strMetric)) > 0 Then
            oWorksheet.[A3:I3] = Split("CAM,TYPE,LAG,UID,TASK NAME,FORECAST START,ACTUAL START,FORECAST FINISH,ACTUAL FINISH", ",")
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = UBound(Split(oDECM(strMetric), ";"))
            cptAddBorders oWorksheet.[A3:I3]
            cptAddShading oWorksheet.[A3:I3]
            Dim vLink As Variant
            For Each vLink In Split(oDECM(strMetric), ";")
              If vLink <> "" Then
                strSQL = "SELECT CAM,'','',UID,TASK_NAME,FS,[AS],FF,AF FROM [tasks.csv] WHERE UID=" & Split(vLink, ",")(0)
                oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
                lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
                oWorksheet.Cells(lngLastRow, 1).CopyFromRecordset oRecordset
                oRecordset.Close
                cptAddShading oWorksheet.Range(oWorksheet.Cells(lngLastRow, 2), oWorksheet.Cells(lngLastRow, 3)), True
                strSQL = "SELECT CAM,'','',UID,TASK_NAME,FS,[AS],FF,AF FROM [tasks.csv] WHERE UID=" & Split(vLink, ",")(1)
                oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
                lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
                oWorksheet.Cells(lngLastRow, 1).CopyFromRecordset oRecordset
                oRecordset.Close
                strSQL = "SELECT DISTINCT TYPE,LAG/480 FROM [links.csv] WHERE [FROM]=" & Split(vLink, ",")(0) & " AND TO=" & Split(vLink, ",")(1)
                oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
                oWorksheet.Cells(lngLastRow, 2).CopyFromRecordset oRecordset
                oRecordset.Close
                'todo: highlight dates in conflict
                'todo: if FS then Pred FF/AF and Succ AS/FS, etc.
                cptAddBorders oWorksheet.Range(oWorksheet.Cells(lngLastRow - 1, 1), oWorksheet.Cells(lngLastRow, 9))
              End If
            Next vLink
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
          End If
        ElseIf strMetric = "06A401a" Then 'Critical path
          'todo: add AD,RD to tasks.csv
          If Len(oDECM(strMetric)) > 0 Then
            oWorksheet.[A3].Value = "TARGET:"
            oWorksheet.[B3].Value = Split(oDECM(strMetric), "|")(0)
            oWorksheet.[C3].Value = ActiveProject.Tasks.UniqueID(Split(oDECM(strMetric), "|")(0)).Name
            strSQL = "SELECT UID,CAM,TASK_NAME,TS/480,FF FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & Split(oDECM(strMetric), "|")(1) & ") "
            strSQL = strSQL & "ORDER BY FF,DUR DESC"
            oRecordset.Open strSQL, strCon, adOpenKeyset
            oWorksheet.[A4:E4] = Split("UID,CAM,TASK_NAME,TOTAL SLACK,FORECAST FINISH", ",")
            oWorksheet.[A5].CopyFromRecordset oRecordset
            oWorksheet.[A4].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A4].End(xlDown), oWorksheet.[A4].End(xlToRight)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A4].End(xlDown), oWorksheet.[A4].End(xlToRight))
            cptAddBorders oWorksheet.Range(oWorksheet.[A4], oWorksheet.[A4].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A4], oWorksheet.[A4].End(xlToRight))
          End If
        ElseIf strMetric = "06A501a" Then 'Baselines
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT UID,CAM,WP,TASK_NAME,BLS,BLF "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:F3] = Split("UID,CAM,WP,TASK NAME,BLS,BLF", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A504a" Then 'Changed Actual Start
          If Dir(strDir & "\06A504a.csv") <> vbNullString Then
            strSQL = "SELECT T1.UID,T2.CAM,T2.TASK_NAME,T1.AS_WAS,T1.AS_IS "
            strSQL = strSQL & "FROM [06A504a.csv] T1 "
            strSQL = strSQL & "INNER JOIN [tasks.csv] T2 ON T2.UID=T1.UID "
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:E3] = Split("UID,CAM,TASK NAME,ACTUAL START WAS,ACTUAL START IS", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A504b" Then 'Changed Actual Finish
          If Dir(strDir & "\06A504b.csv") <> vbNullString Then
            strSQL = "SELECT T1.UID,T2.CAM,T2.TASK_NAME,T1.AF_WAS,T1.AF_IS "
            strSQL = strSQL & "FROM [06A504b.csv] T1 "
            strSQL = strSQL & "INNER JOIN [tasks.csv] T2 ON T2.UID=T1.UID "
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:E3] = Split("UID,CAM,TASK NAME,ACTUAL FINISH WAS,ACTUAL FINISH IS", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A505a" Then 'In-Progress Tasks w/o Actual Start
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT UID,CAM,TASK_NAME,FS,[AS],EVP "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:F3] = Split("UID,CAM,TASK NAME,FORECAST START,ACTUAL START,EV%", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A505b" Then 'Complete Tasks w/o Actual Finish
          'todo: UID always first
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT UID,CAM,TASK_NAME,FF,AF,EVP "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A3:F3] = Split("UID,CAM,TASK NAME,FORECAST FINISH,ACTUAL FINISH,EV%", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06A506a" Then 'Bogus Actuals
          If Len(oDECM(strMetric)) > 0 Then
            oWorksheet.[A3] = "STATUS DATE:"
            oWorksheet.[B3] = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
            strSQL = "SELECT UID,CAM,TASK_NAME,[AS],AF "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A4:E4] = Split("UID,CAM,TASK NAME,ACTUAL START,ACTUAL FINISH", ",")
            oWorksheet.[A5].CopyFromRecordset oRecordset
            oWorksheet.[A4].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A4].End(xlDown), oWorksheet.[A4].End(xlToRight)).Columns.AutoFit
            oWorksheet.[A3:B3].Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A4].End(xlDown), oWorksheet.[A4].End(xlToRight))
            cptAddBorders oWorksheet.Range(oWorksheet.[A4], oWorksheet.[A4].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A4], oWorksheet.[A4].End(xlToRight))
          End If
        ElseIf strMetric = "06A506b" Then 'Invalid Forecast
          If Len(oDECM(strMetric)) > 0 Then
            oWorksheet.[A3].Value = "Status Date:"
            oWorksheet.[B3].Value = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
            strSQL = "SELECT UID,CAM,TASK_NAME,FS,FF "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[A4:E4] = Split("UID,CAM,TASK NAME,FORECAST START,FORECAST FINISH", ",")
            oWorksheet.[A5].CopyFromRecordset oRecordset
            oWorksheet.[A4].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A4].End(xlDown), oWorksheet.[A4].End(xlToRight)).Columns.AutoFit
            oWorksheet.[A3:B3].Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A4].End(xlDown), oWorksheet.[A4].End(xlToRight))
            cptAddBorders oWorksheet.Range(oWorksheet.[A4], oWorksheet.[A4].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A4], oWorksheet.[A4].End(xlToRight))
          End If
        ElseIf strMetric = "06A506c" Then 'Riding the Status Date
          If Dir(strDir & "\06A506c-x.csv") <> vbNullString Then
            strSQL = "SELECT * FROM [06A506c-x.csv]"
            oRecordset.Open strSQL, strCon, adOpenKeyset
            For lngField = 0 To oRecordset.Fields.Count - 1
              oWorksheet.Cells(3, lngField + 1).Value = oRecordset.Fields(lngField).Name
            Next lngField
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight).End(xlDown))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        ElseIf strMetric = "06I201a" Then 'SVTs
          If Len(oDECM(strMetric)) > 0 Then
            strSQL = "SELECT UID,CAM,TASK_NAME,SUM(BLW),SUM(BLC) "
            strSQL = strSQL & "FROM [tasks.csv] "
            strSQL = strSQL & "WHERE UID IN (" & oDECM(strMetric) & ") "
            strSQL = strSQL & "ORDER BY CAM"
            oRecordset.Open strSQL, strCon, adOpenKeyset
            oWorksheet.[A3:E3] = Split("UID,CAM,TASK NAME,BLW,BLC", ",")
            oWorksheet.[A4].CopyFromRecordset oRecordset
            oWorksheet.[A3].End(xlToRight).Offset(-1, 0) = oRecordset.RecordCount
            oRecordset.Close
            oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight)).Columns.AutoFit
            cptAddBorders oWorksheet.Range(oWorksheet.[A3].End(xlDown), oWorksheet.[A3].End(xlToRight))
            cptAddBorders oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
            cptAddShading oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight))
          End If
        Else
          Debug.Print "MISSING: " & strMetric & ": " & oDECM(strMetric)
        End If
        Set oTasks = Nothing
        Application.StatusBar = "Exporting " & strMetric & "...done."
        .lblStatus.Caption = "Exporting " & strMetric & "...done."
        .lblProgress.Width = ((lngItem + 1) / .lboMetrics.ListCount) * .lblStatus.Width
        DoEvents
      Next lngItem
      .lblStatus.Caption = "Export Complete."
      .lboMetrics.Value = .lboMetrics.List(0)
      .lboMetrics.Selected(0) = True
      .lboMetrics.ListIndex = 0
    End With
    'create hyperlinks
    Set oWorksheet = oWorkbook.Sheets("DECM Dashboard")
    oWorksheet.Activate
    Set oRange = oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A1048576].End(xlUp))
    For Each oCell In oRange.Cells
      oWorksheet.Hyperlinks.Add Anchor:=oCell, Address:="", SubAddress:="'" & CStr(oCell.Value) & "'!A1", TextToDisplay:=CStr(oCell.Value), ScreenTip:="Jump to " & CStr(oCell.Value)
    Next oCell
  End If
  
  Set oBorders = oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A1].End(xlToRight).End(xlDown))
  Set oShading = oWorksheet.[A1:G1]
  
  'get general stats
  oExcel.WindowState = xlNormal
  oWorkbook.Activate
  Set oRecordset = CreateObject("ADODB.Recordset")
  'count of complete, incomplete, total CA, by CAM
  'limit to PMB tasks
  'determine if resource assignments
  strSQL = "SELECT * FROM [assignments.csv]"
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  blnResourceAssignments = Not oRecordset.EOF
  oRecordset.Close
  If blnResourceAssignments Then
    strSQL = "SELECT T1.CAM,SUM(INCOMPLETE) AS [_INCOMPLETE],SUM(COMPLETE) AS [_COMPLETE] "
    strSQL = strSQL & "FROM ("
    strSQL = strSQL & "SELECT T1.CAM,T1.CA,IIF(AVG(T1.EVP)<100,1,0) AS [INCOMPLETE],IIF(AVG(T1.EVP)=100,1,0) AS [COMPLETE] "
    strSQL = strSQL & "FROM [tasks.csv] T1 "
    strSQL = strSQL & "INNER JOIN [assignments.csv] T2 ON T2.TASK_UID=T1.UID "
    strSQL = strSQL & "GROUP BY T1.CAM,T1.CA "
    strSQL = strSQL & "HAVING SUM(T2.BLW)>0 OR SUM(T2.BLC)>0) GROUP BY T1.CAM "
  Else
    strSQL = "SELECT T1.CAM,SUM(INCOMPLETE) AS [_INCOMPLETE],SUM(COMPLETE) AS [_COMPLETE] "
    strSQL = strSQL & "FROM ("
    strSQL = strSQL & "SELECT T1.CAM,T1.CA,IIF(AVG(T1.EVP)<100,1,0) AS [INCOMPLETE],IIF(AVG(T1.EVP)=100,1,0) AS [COMPLETE] "
    strSQL = strSQL & "FROM [tasks.csv] T1 "
    strSQL = strSQL & "GROUP BY T1.CAM,T1.CA "
    strSQL = strSQL & "HAVING SUM(T1.BLW)>0 OR SUM(T1.BLC)>0 "
    strSQL = strSQL & ") GROUP BY T1.CAM "
  End If
  oRecordset.Open strSQL, strCon, adOpenKeyset
  If Not oRecordset.EOF Then
    oWorksheet.[I2:L2].Merge True
    oWorksheet.[I2] = "CONTROL ACCOUNTS"
    oWorksheet.[I2].HorizontalAlignment = xlCenter
    oWorksheet.[I3:L3] = Split("CAM,INCOMPLETE,COMPLETE,TOTAL", ",")
    oWorksheet.[I4].CopyFromRecordset oRecordset
    lngFirstRow = oWorksheet.[L1048576].End(xlUp).Row + 1
    lngLastRow = oWorksheet.[I1048576].End(xlUp).Row + 1
    oWorksheet.Range(oWorksheet.Cells(lngFirstRow, 12), oWorksheet.Cells(lngLastRow - 1, 12)).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    oWorksheet.Cells(lngLastRow, 9) = "TOTAL:"
    oWorksheet.Cells(lngLastRow, 9).HorizontalAlignment = xlRight
    oWorksheet.Range(oWorksheet.Cells(lngLastRow, 10), oWorksheet.Cells(lngLastRow, 12)).FormulaR1C1 = "=SUM(R" & lngFirstRow & "C:R" & lngLastRow - 1 & "C)"
    oWorksheet.Range(oWorksheet.Cells(lngFirstRow, 10), oWorksheet.Cells(lngLastRow, 12)).NumberFormat = "#,##0"
    'section header
    oWorksheet.[I2].Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.[I2])
    'column header
    oWorksheet.[I3:L3].Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.[I3:L3])
    'total
    oWorksheet.Range(oWorksheet.[I1048576].End(xlUp), oWorksheet.[I1048576].End(xlUp).Offset(0, 3)).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.Range(oWorksheet.[I1048576].End(xlUp), oWorksheet.[I1048576].End(xlUp).Offset(0, 3)))
    Set oBorders = oExcel.Union(oBorders, oWorksheet.Range(oWorksheet.[I2], oWorksheet.[I1048576].End(xlUp).Offset(0, 3)))
  End If
  oRecordset.Close
'  'todo: CA checksum
'  strSQL = "SELECT T1.CA, T1.CAM,SUM(T2.BLW+T2.BLC) AS BAC "
'  strSQL = strSQL & "FROM [tasks.csv] AS T1 INNER JOIN [assignments.csv] AS T2 ON T2.TASK_UID=T1.UID "
'  strSQL = strSQL & "GROUP BY T1.CA,T1.CAM "
'  strSQL = strSQL & "HAVING SUM(T2.BLW+T2.BLC)<=0"
'  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
'  If Not oRecordset.EOF Then
'    oWorksheet.[L1048576].End(xlUp).Offset(0, 1) = oRecordset.RecordCount
'    'todo: gumball red if >0
'    'todo: ALL CAs, or ONLY CAs
'  End If
'  oRecordset.Close
   
  'count of complete, incomplete, total WP, by CAM *only includes WPs in the IMS
  'first try to limit by PMB tasks (assuming resource assignments)
  If blnResourceLoaded Then
    strSQL = "SELECT T1.CAM,SUM(INCOMPLETE) AS [_INCOMPLETE],SUM(COMPLETE) AS [_COMPLETE] "
    strSQL = strSQL & "FROM ("
    strSQL = strSQL & "SELECT T1.CAM,T1.WP,IIF(AVG(T1.EVP)<100,1,0) AS [INCOMPLETE],IIF(AVG(T1.EVP)=100,1,0) AS [COMPLETE] "
    strSQL = strSQL & "FROM [tasks.csv] T1 "
    strSQL = strSQL & "INNER JOIN [assignments.csv] T2 ON T2.TASK_UID=T1.UID "
    strSQL = strSQL & "WHERE T1.EVT<>'" & strLOE & "' "
    strSQL = strSQL & "GROUP BY T1.CAM,T1.WP "
    strSQL = strSQL & "HAVING SUM(T2.BLW)>0 OR SUM(T2.BLC)>0) "
    strSQL = strSQL & "GROUP BY T1.CAM"
  Else
    strSQL = "SELECT T1.CAM,SUM(INCOMPLETE) AS [_INCOMPLETE],SUM(COMPLETE) AS [_COMPLETE] "
    strSQL = strSQL & "FROM ("
    strSQL = strSQL & "SELECT T1.CAM,T1.WP,IIF(AVG(T1.EVP)<100,1,0) AS [INCOMPLETE],IIF(AVG(T1.EVP)=100,1,0) AS [COMPLETE] "
    strSQL = strSQL & "FROM [tasks.csv] T1 "
    strSQL = strSQL & "WHERE T1.EVT<>'" & strLOE & "' "
    'strSQL = strSQL & "AND IsDate(T1.BLS) AND IsDate(T1.BLF) "
    strSQL = strSQL & "GROUP BY T1.CAM,T1.WP) "
    strSQL = strSQL & "GROUP BY T1.CAM"
  End If
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  If Not oRecordset.EOF Then
    oWorksheet.[N2:Q2].Merge True
    oWorksheet.[N2].Value = "DISCRETE WORK PACKAGES"
    oWorksheet.[N2].HorizontalAlignment = xlCenter
    oWorksheet.[N3:Q3] = Split("CAM,INCOMPLETE,COMPLETE,TOTAL", ",")
    oWorksheet.[N4].CopyFromRecordset oRecordset
    lngFirstRow = oWorksheet.[Q1048576].End(xlUp).Row + 1
    lngLastRow = oWorksheet.[N1048576].End(xlUp).Row + 1
    oWorksheet.Range(oWorksheet.Cells(lngFirstRow, 17), oWorksheet.Cells(lngLastRow - 1, 17)).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    oWorksheet.Cells(lngLastRow, 14) = "TOTAL:"
    oWorksheet.Cells(lngLastRow, 14).HorizontalAlignment = xlRight
    oWorksheet.Range(oWorksheet.Cells(lngLastRow, 15), oWorksheet.Cells(lngLastRow, 17)).FormulaR1C1 = "=SUM(R" & lngFirstRow & "C:R" & lngLastRow - 1 & "C)"
    oWorksheet.Range(oWorksheet.Cells(lngFirstRow, 15), oWorksheet.Cells(lngLastRow, 17)).NumberFormat = "#,##0"
    'section header
    oWorksheet.[N1048576].End(xlUp).End(xlUp).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.[N1048576].End(xlUp).End(xlUp))
    'column header
    oWorksheet.Range(oWorksheet.[N1048576].End(xlUp).End(xlUp).Offset(1, 0), oWorksheet.[N1048576].End(xlUp).End(xlUp).Offset(1, 0).Offset(0, 3)).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.Range(oWorksheet.[N1048576].End(xlUp).End(xlUp).Offset(1, 0), oWorksheet.[N1048576].End(xlUp).End(xlUp).Offset(1, 0).Offset(0, 3)))
    'total
    oWorksheet.Range(oWorksheet.[N1048576].End(xlUp), oWorksheet.[N1048576].End(xlUp).Offset(0, 3)).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.Range(oWorksheet.[N1048576].End(xlUp), oWorksheet.[N1048576].End(xlUp).Offset(0, 3)))
    Set oBorders = oExcel.Union(oBorders, oWorksheet.Range(oWorksheet.[N1048576].End(xlUp), oWorksheet.[N1048576].End(xlUp).Offset(0, 3).End(xlUp).Offset(-1, 0)))
    'todo: formula=ABS(Nx-Lx)
    'todo: gumball: reverse order; icon only; green when <=0; etc.
  End If
  oRecordset.Close
'  'todo: WP checksum
'  strSQL = "SELECT T1.CAM, T1.WP,SUM(T2.BLW+T2.BLC) AS BAC "
'  strSQL = strSQL & "FROM [tasks.csv] AS T1 INNER JOIN [assignments.csv] AS T2 ON T2.TASK_UID=T1.UID "
'  strSQL = strSQL & "GROUP BY T1.CAM,T1.WP "
'  strSQL = strSQL & "HAVING SUM(T2.BLW+T2.BLC)<=0"
'  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
'  If Not oRecordset.EOF Then
'    oWorksheet.[L1048576].End(xlUp).Offset(0, 1) = oRecordset.RecordCount
'    'todo: red gumball if >0
'  End If
'  oRecordset.Close
  
  'count of complete, incomplete, total PMB tasks, by CAM
  'first try to limit by PMB tasks (assumes resource assignments)
  If blnResourceLoaded Then
    strSQL = "SELECT CAM,SUM(INCOMPLETE) AS [_INCOMPLETE],SUM(COMPLETE) AS [_COMPLETE] "
    strSQL = strSQL & "FROM [tasks.csv] t INNER JOIN "
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT T1.UID,IIF(T1.AF IS NULL,1,0) AS [INCOMPLETE],IIF(T1.AF IS NOT NULL,1,0) AS [COMPLETE], SUM(T2.BLW),SUM(T2.BLC) "
    strSQL = strSQL & "FROM [tasks.csv] T1 "
    strSQL = strSQL & "INNER JOIN [assignments.csv] T2 ON T2.TASK_UID=T1.UID "
    strSQL = strSQL & "GROUP BY T1.UID,T1.AF "
    strSQL = strSQL & "HAVING SUM(T2.BLW)>0 OR SUM(T2.BLC)>0 ) AS s ON s.UID=t.UID "
    strSQL = strSQL & "WHERE t.EVT IS NOT NULL AND t.EVT<>'" & strLOE & "' "
    strSQL = strSQL & "GROUP BY CAM"
  Else
    strSQL = "SELECT CAM,SUM(INCOMPLETE) AS [_INCOMPLETE],SUM(COMPLETE) AS [_COMPLETE] "
    strSQL = strSQL & "FROM [tasks.csv] T INNER JOIN "
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT T1.UID,IIF(T1.AF IS NULL,1,0) AS [INCOMPLETE],IIF(T1.AF IS NOT NULL,1,0) AS [COMPLETE] "
    strSQL = strSQL & "FROM [tasks.csv] T1 "
    strSQL = strSQL & ") AS S ON S.UID=T.UID "
    strSQL = strSQL & "WHERE T.EVT IS NOT NULL AND T.EVT<>'" & strLOE & "' "
    strSQL = strSQL & "GROUP BY CAM"
  End If
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  If Not oRecordset.EOF Then
    lngLastRow = oWorksheet.[I1048576].End(xlUp).Row + 2
    oWorksheet.Range(oWorksheet.Cells(lngLastRow, 9), oWorksheet.Cells(lngLastRow, 12)).Merge True
    oWorksheet.Cells(lngLastRow, 9).Value = "DISCRETE PMB TASKS"
    oWorksheet.Cells(lngLastRow, 9).HorizontalAlignment = xlCenter
    oWorksheet.Range(oWorksheet.Cells(lngLastRow + 1, 9), oWorksheet.Cells(lngLastRow + 1, 12)) = Split("CAM,INCOMPLETE,COMPLETE,TOTAL", ",")
    oWorksheet.Cells(lngLastRow + 2, 9).CopyFromRecordset oRecordset
    'get total
    lngFirstRow = oWorksheet.[L1048576].End(xlUp).Row + 1
    lngLastRow = oWorksheet.[I1048576].End(xlUp).Row + 1
    oWorksheet.Range(oWorksheet.Cells(lngFirstRow, 12), oWorksheet.Cells(lngLastRow - 1, 12)).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    oWorksheet.Cells(lngLastRow, 9) = "TOTAL:"
    oWorksheet.Cells(lngLastRow, 9).HorizontalAlignment = xlRight
    oWorksheet.Range(oWorksheet.Cells(lngLastRow, 10), oWorksheet.Cells(lngLastRow, 12)).FormulaR1C1 = "=SUM(R" & lngFirstRow & "C:R" & lngLastRow - 1 & "C)"
    oWorksheet.Range(oWorksheet.Cells(lngFirstRow, 10), oWorksheet.Cells(lngLastRow, 12)).NumberFormat = "#,##0"
    'section header
    oWorksheet.[I1048576].End(xlUp).End(xlUp).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.[I1048576].End(xlUp).End(xlUp))
    'column header
    oWorksheet.Range(oWorksheet.[I1048576].End(xlUp).End(xlUp).Offset(1, 0), oWorksheet.[I1048576].End(xlUp).End(xlUp).Offset(1, 0).Offset(0, 3)).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.Range(oWorksheet.[I1048576].End(xlUp).End(xlUp).Offset(1, 0), oWorksheet.[I1048576].End(xlUp).End(xlUp).Offset(1, 0).Offset(0, 3)))
    'total
    oWorksheet.Range(oWorksheet.[I1048576].End(xlUp), oWorksheet.[I1048576].End(xlUp).End(xlToRight)).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.Range(oWorksheet.[I1048576].End(xlUp), oWorksheet.[I1048576].End(xlUp).End(xlToRight)))
    Set oBorders = oExcel.Union(oBorders, oWorksheet.Range(oWorksheet.[I1048576].End(xlUp), oWorksheet.[I1048576].End(xlUp).Offset(0, 3).End(xlUp).Offset(-1, 0)))
  End If
  oRecordset.Close
  
  'count of relationship FS, SS, FF, SF
  strSQL = "SELECT TYPE,COUNT(TYPE) FROM [links.csv] GROUP BY TYPE"
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  If Not oRecordset.EOF Then
    lngLastRow = oWorksheet.[N1048576].End(xlUp).Row + 2
    oWorksheet.Range(oWorksheet.Cells(lngLastRow, 14), oWorksheet.Cells(lngLastRow, 16)).Merge True
    oWorksheet.Cells(lngLastRow, 14).Value = "LOGIC RELATIONSHIPS"
    oWorksheet.Cells(lngLastRow, 14).HorizontalAlignment = xlCenter
    oWorksheet.Range(oWorksheet.Cells(lngLastRow + 1, 14), oWorksheet.Cells(lngLastRow + 1, 16)) = Split("TYPE,COUNT,PERCENT", ",")
    oWorksheet.Cells(lngLastRow + 2, 14).CopyFromRecordset oRecordset
    'get total
    lngLastRow = oWorksheet.[O1048576].End(xlUp).Row + 1
    oWorksheet.Cells(lngLastRow, 14).Value = "TOTAL:"
    oWorksheet.Cells(lngLastRow, 14).HorizontalAlignment = xlRight
    lngFirstRow = oWorksheet.[P1048576].End(xlUp).Row + 1
    oWorksheet.Cells(lngLastRow, 15).FormulaR1C1 = "=SUM(R" & lngFirstRow & "C:R[-1]C"
    oWorksheet.Range(oWorksheet.Cells(lngFirstRow, 15), oWorksheet.Cells(lngLastRow, 15)).NumberFormat = "#,##0"
    'get percentages
    oWorksheet.Range(oWorksheet.Cells(lngFirstRow, 16), oWorksheet.Cells(lngLastRow, 16)).FormulaR1C1 = "=RC[-1]/R" & lngLastRow & "C[-1]"
    oWorksheet.Range(oWorksheet.Cells(lngFirstRow, 16), oWorksheet.Cells(lngLastRow, 16)).NumberFormat = "0%"
    oWorksheet.Columns("I:Q").AutoFit
    'section header
    oWorksheet.[N1048576].End(xlUp).End(xlUp).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.[N1048576].End(xlUp).End(xlUp))
    'column header
    oWorksheet.Range(oWorksheet.[N1048576].End(xlUp).End(xlUp).Offset(1, 0), oWorksheet.[N1048576].End(xlUp).End(xlUp).Offset(1, 0).Offset(0, 2)).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.Range(oWorksheet.[N1048576].End(xlUp).End(xlUp).Offset(1, 0), oWorksheet.[N1048576].End(xlUp).End(xlUp).Offset(1, 0).Offset(0, 2)))
    'total
    oWorksheet.Range(oWorksheet.[N1048576].End(xlUp), oWorksheet.[N1048576].End(xlUp).End(xlToRight)).Font.Bold = True
    Set oShading = oExcel.Union(oShading, oWorksheet.Range(oWorksheet.[N1048576].End(xlUp), oWorksheet.[N1048576].End(xlUp).End(xlToRight)))
    Set oBorders = oExcel.Union(oBorders, oWorksheet.Range(oWorksheet.[N1048576].End(xlUp), oWorksheet.[N1048576].End(xlUp).Offset(0, 2).End(xlUp).Offset(-1, 0)))
  End If
  oRecordset.Close
  
  cptAddBorders oBorders
  cptAddBorders oShading
  cptAddShading oShading
  
  oWorksheet.[A1:G1].Insert xlShiftDown
  oWorksheet.[A1:A2].EntireRow.Insert xlShiftDown
  If InStr(ActiveProject.Name, "/") > 0 Then
    oWorksheet.[A1].Value = cptRegEx(ActiveProject.Name, "[^/]*.mpp")
  ElseIf InStr(ActiveProject.Name, ":") > 0 Then
    oWorksheet.[A1].Value = cptRegEx(ActiveProject.Name, "[^\\]*.mpp")
  ElseIf InStr(ActiveProject.Name, "<>") > 0 Then
    oWorksheet.[A1].Value = Replace(ActiveProject.Name, "<>\", "")
  Else
    oWorksheet.[A1].Value = ActiveProject.Name
  End If
  oWorksheet.[A1].Font.Size = 18
  oWorksheet.[A1].Font.Bold = True
  oWorksheet.[A2].Value = "Date:"
  oWorksheet.[B2].Value = Now
  oWorksheet.[B2].NumberFormat = "[$-en-US]m/d/yyyy h:mm AM/PM;@"
  oWorksheet.[A3].Value = "Analyst:"
  oWorksheet.[B3].Value = Application.UserName
  
  'dump out the integration settings used
  oWorksheet.[S4:U4].Merge True
  oWorksheet.[S4].Value = "INTEGRATION SETTINGS"
  oWorksheet.[S4].HorizontalAlignment = xlCenter
  oWorksheet.[S4].Font.Bold = True
  
  For Each vSetting In Split("WBS,OBS,CA,CAM,WP,EVP,EVT,LOE,PP", ",")
    lngLastRow = oWorksheet.[S1048576].End(xlUp).Row + 1
    oWorksheet.Cells(lngLastRow, 19).Value = vSetting
    If vSetting = "LOE" Then
      oWorksheet.Cells(lngLastRow, 20) = FieldConstantToFieldName(Split(cptGetSetting("Integration", "EVT"), "|")(0))
      oWorksheet.Cells(lngLastRow, 21) = "EVT='" & cptGetSetting("Integration", CStr(vSetting)) & "'"
    ElseIf vSetting = "PP" Then
      oWorksheet.Cells(lngLastRow, 20) = FieldConstantToFieldName(Split(cptGetSetting("Integration", "EVT"), "|")(0))
      oWorksheet.Cells(lngLastRow, 21) = "EVT='" & cptGetSetting("Integration", CStr(vSetting)) & "'"
    Else
      lngField = CLng(Split(cptGetSetting("Integration", CStr(vSetting)), "|")(0))
      oWorksheet.Cells(lngLastRow, 20).Value = FieldConstantToFieldName(lngField)
      If Len(CustomFieldGetName(lngField)) > 0 Then
        oWorksheet.Cells(lngLastRow, 21).Value = CustomFieldGetName(lngField)
      Else
        oWorksheet.Cells(lngLastRow, 21).Value = FieldConstantToFieldName(lngField)
      End If
    End If
  Next vSetting
  cptAddShading oWorksheet.[S4]
  cptAddBorders oWorksheet.Range(oWorksheet.[S4], oWorksheet.[S4].End(xlDown).Offset(0, 2))
  cptAddBorders oWorksheet.[S4:U4]
  oWorksheet.Range(oWorksheet.[S4], oWorksheet.[S4].End(xlDown).Offset(0, 2)).Columns.AutoFit
  
  oExcel.WindowState = xlMaximized
  oExcel.ActiveWindow.DisplayGridlines = False
  Application.ActivateMicrosoftApp pjMicrosoftExcel

exit_here:
  On Error Resume Next
  Set oShading = Nothing
  Set oBorders = Nothing
  Set o06A101a = Nothing
  Set o10A103a = Nothing
  Set oRecordset = Nothing
  Set oTasks = Nothing
  cptSpeed False
  Set oCell = Nothing
  Set oRange = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDECM_bas", "cptDECM_EXPORT", Err, Erl)
  Resume exit_here
End Sub

Sub cptDECM_UPDATE_VIEW(strMetric As String, Optional strList As String)
  Dim strGroup As String
  ScreenUpdating = False
  ActiveWindow.TopPane.Activate
  FilterClear
  GroupClear
  Sort "ID", renumber:=False, Outline:=True
  OptionsViewEx DisplaySummaryTasks:=True
  OutlineShowAllTasks
  If strMetric <> "06A208a" Then OptionsViewEx DisplaySummaryTasks:=False

  Select Case strMetric
    Case "05A101a" '1 CA : 1 OBS
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "CA"), "|")(0)), pjAutoFilterIn, "equals", strList
        'group by CA,OBS
        strGroup = "cpt 05A101a 1 CA : 1 OBS"
        If cptGroupExists(strGroup) Then ActiveProject.TaskGroups2(strGroup).Delete
        ActiveProject.TaskGroups.Add strGroup, FieldConstantToFieldName(Split(cptGetSetting("Integration", "CA"), "|")(0))
        ActiveProject.TaskGroups(strGroup).GroupCriteria.Add FieldConstantToFieldName(Split(cptGetSetting("Integration", "OBS"), "|")(0))
        GroupApply Name:=strGroup
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
    
    Case "05A102a" '1 CA : 1 CAM
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "CA"), "|")(0)), pjAutoFilterIn, "equals", strList
        'group by CA,CAM
        strGroup = "cpt 05A102a 1 CA : 1 CAM"
        If cptGroupExists(strGroup) Then ActiveProject.TaskGroups2(strGroup).Delete
        ActiveProject.TaskGroups.Add strGroup, FieldConstantToFieldName(Split(cptGetSetting("Integration", "CA"), "|")(0))
        ActiveProject.TaskGroups(strGroup).GroupCriteria.Add FieldConstantToFieldName(Split(cptGetSetting("Integration", "CAM"), "|")(0))
        GroupApply Name:=strGroup
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
    
    Case "05A103a" '1 CA : 1 WBS
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "CA"), "|")(0)), pjAutoFilterIn, "equals", strList
        'group by CA,WBS
        strGroup = "cpt 05A103a 1 CA : 1 WBS"
        If cptGroupExists(strGroup) Then ActiveProject.TaskGroups2(strGroup).Delete
        ActiveProject.TaskGroups.Add strGroup, FieldConstantToFieldName(Split(cptGetSetting("Integration", "CA"), "|")(0))
        ActiveProject.TaskGroups(strGroup).GroupCriteria.Add FieldConstantToFieldName(Split(cptGetSetting("Integration", "WBS"), "|")(0))
        GroupApply Name:=strGroup
        
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
    
    Case "06A101a" 'WP mismatches
      'todo: do what?
    
    Case "06A210a" 'LOE driving discrete
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        ScreenUpdating = True
        SetAutoFilter "Unique ID", pjAutoFilterIn, "equals", strList
        strEVT = Split(cptGetSetting("Integration", "EVT"), "|")(1)
        strGroup = "cpt 06A210a LOE driving Discrete"
        If cptGroupExists(strGroup) Then ActiveProject.TaskGroups2(strGroup).Delete
        ActiveProject.TaskGroups2.Add strGroup, strEVT
        GroupApply Name:=strGroup
        SelectAll
        On Error Resume Next
        ActiveWindow.BottomPane.Activate
        If Err.Number > 0 Then
          Application.WindowSplit
          ActiveWindow.TopPane.Activate
          SelectAll
          Err.Clear
        End If
        ActiveWindow.BottomPane.Activate
        ViewApply "Network Diagram"
        
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
        
    Case "CPT02"
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0)), pjAutoFilterIn, "equals", strList
        'group by WP, CA
        strGroup = "cpt 1wp_1ca"
        If cptGroupExists(strGroup) Then ActiveProject.TaskGroups2(strGroup).Delete
        ActiveProject.TaskGroups.Add strGroup, FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0))
        ActiveProject.TaskGroups(strGroup).GroupCriteria.Add FieldConstantToFieldName(Split(cptGetSetting("Integration", "CA"), "|")(0))
        GroupApply Name:=strGroup
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
      
    Case "06A212a" 'out of sequence
      If Len(strList) > 0 Then
        SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
    
    Case "06A401a" 'critical path
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strList
        Sort Key1:="Finish", ascending1:=True, Key2:="Duration", ascending2:=False, renumber:=False, Outline:=False
        SelectBeginning
        EditGoTo Date:=ActiveSelection.Tasks(1).Finish
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
    
    Case "10A102a" '1 WP : 1 EVT
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0)), pjAutoFilterIn, "equals", strList
        'group by WP,EVT
        strGroup = "cpt 10A102a 1 WP : 1 EVT"
        If cptGroupExists(strGroup) Then ActiveProject.TaskGroups2(strGroup).Delete
        ActiveProject.TaskGroups.Add strGroup, FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0))
        ActiveProject.TaskGroups(strGroup).GroupCriteria.Add FieldConstantToFieldName(Split(cptGetSetting("Integration", "EVT"), "|")(0))
        GroupApply Name:=strGroup
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
    
    Case "10A103a" '0/100 >1 fiscal periods
      If Len(strList) > 0 Then
        strList = Left(strList, Len(strList) - 1) 'remove last tab
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0)), pjAutoFilterIn, "equals", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
      
    Case "10A109b" 'WP with no budget
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0)), pjAutoFilterIn, "equals", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
    Case "10A302a" 'same as 29A601a
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0)), pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
    
    Case "10A302b" 'same as 29A601a
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0)), pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If

    Case "10A303a"
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0)), pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
      
    Case "11A101a" 'CA BAC = Sum(WP BAC)
      If Len(strList) > 0 Then
        Dim strCAList As String
        Dim strWPList As String
        strList = Left(strList, Len(strList) - 1) 'remove last comma
        strCAList = Split(strList, ";")(0)
        strWPList = Split(strList, ";")(1)
        
        If Len(strCAList) > 0 Then
          strCAList = Replace(strCAList, ",", vbTab)
          SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "CA"), "|")(0)), pjAutoFilterIn, "equals", strCAList
        End If
        If Len(strWPList) > 0 Then
          strWPList = Replace(strWPList, ",", vbTab)
          SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0)), pjAutoFilterIn, "equals", strWPList
        End If
        strGroup = "cpt 11A101a CA BAC = SUM(WP BAC)"
        If cptGroupExists(strGroup) Then ActiveProject.TaskGroups2(strGroup).Delete
        ActiveProject.TaskGroups.Add strGroup, FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0))
        ActiveProject.TaskGroups(strGroup).GroupCriteria.Add FieldConstantToFieldName(Split(cptGetSetting("Integration", "CA"), "|")(0))
        GroupApply Name:=strGroup
        OptionsViewEx DisplaySummaryTasks:=True
        OutlineShowTasks 2
        'collapse to 2nd level
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
      
    Case "29A601a" 'PPs within Rolling Wave Period
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WP"), "|")(0)), pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
    
    Case Else
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter "Unique ID", pjAutoFilterIn, "equals", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "equals", "<< zero results >>"
      End If
      
  End Select
  SelectBeginning
  ScreenUpdating = True
  
End Sub

Function cptGetOutOfSequence(ByRef myDECM_frm As cptDECM_frm) As String
  'objects
  Dim oAssignment As MSProject.Assignment
  Dim oOOS As Scripting.Dictionary
  Dim oCalendar As MSProject.Calendar
  Dim oSubproject As MSProject.SubProject
  'Dim oSubMap As Scripting.Dictionary
  Dim oTask As MSProject.Task
  Dim oLink As MSProject.TaskDependency
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  'strings
  Dim strOOS As String
  Dim strEarliest As String
  Dim strProject As String
  Dim strMacro As String
  Dim strMsg As String
  Dim strProjectNumber As String
  Dim strProjectName As String
  Dim strDir As String
  Dim strFile As String
  'longs
  Dim lngItem As Long
  Dim lngLagType As Long
  Dim lngLag As Long
  Dim lngFactor As Long
  Dim lngToUID As Long
  Dim lngFromUID As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngLastRow As Long
  Dim lngOOS As Long
  'integers
  'doubles
  'booleans
  Dim blnElapsed As Boolean
  Dim blnMaster As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  Dim dtDate As Date
  
  If cptErrorTrapping Then
    On Error GoTo err_here
    cptSpeed True
  Else
    On Error GoTo 0
  End If
  
  blnMaster = ActiveProject.Subprojects.Count > 0
  If blnMaster Then
'    'set up mapping
'    If oSubMap Is Nothing Then
'      Set oSubMap = CreateObject("Scripting.Dictionary")
'    Else
'      oSubMap.RemoveAll
'    End If
'    For Each oSubproject In ActiveProject.Subprojects
'      If Left(oSubproject.Path, 2) <> "<>" Then 'offline
'        oSubMap.Add Replace(Dir(oSubproject.Path), ".mpp", ""), 0
'      ElseIf Left(oSubproject.Path, 2) = "<>" Then 'online
'        oSubMap.Add oSubproject.Path, 0
'      End If
'    Next oSubproject
    For Each oTask In ActiveProject.Tasks
      If oTask Is Nothing Then GoTo next_mapping_task
'      If oSubMap.Exists(oTask.Project) Then
'        If oSubMap(oTask.Project) > 0 Then GoTo next_mapping_task
'        If Not oTask.Summary Then
'          oSubMap.Item(oTask.Project) = CLng(oTask.UniqueID / 4194304)
'        End If
'      End If
next_mapping_task:
      If oTask.Active Then lngTasks = lngTasks + 1
    Next oTask
  Else
    lngTasks = ActiveProject.Tasks.Count
  End If
  
  Set oExcel = CreateObject("Excel.Application")
  On Error Resume Next
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "06A212a"
  oWorksheet.[A2:K2] = Split("UID,ID,TASK,DATE,TYPE,LAG,UID,ID,TASK,DATE,COMMENT", ",")
  oWorksheet.[A1:D1].Merge
  oWorksheet.[A1].Value = "FROM"
  oWorksheet.[A1].HorizontalAlignment = xlCenter
  oWorksheet.[G1:J1].Merge
  oWorksheet.[G1].Value = "TO"
  oWorksheet.[G1].HorizontalAlignment = xlCenter
  oWorksheet.[A1:K2].Font.Bold = True
  oExcel.EnableEvents = False
    
  lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
  lngTask = 0
  
  Set oOOS = CreateObject("Scripting.Dictionary")
  
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task 'skip blank lines
    If oTask.Summary Then GoTo next_task 'skip summary tasks
    If Not oTask.Active Then GoTo next_task 'skip inactive tasks
    If IsDate(oTask.ActualFinish) Then GoTo next_task 'incomplete predecessors only
    For Each oLink In oTask.TaskDependencies
      If oLink.From.Guid = oTask.Guid Then   'predecessors only
        If Not oLink.To.Active Then GoTo next_link
        lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
        If blnMaster And oLink.From.ExternalTask Then
          'fix the pred UID if master-sub
          lngFromUID = oLink.From.GetField(185073906) Mod 4194304
          strProject = oLink.From.Project
          If InStr(oLink.From.Project, "\") > 0 Then
            strProject = Replace(strProject, ".mpp", "")
            strProject = Mid(strProject, InStrRev(strProject, "\") + 1)
          End If
          lngFactor = oSubMap(strProject)
          lngFromUID = (lngFactor * 4194304) + lngFromUID
        Else
          If blnMaster Then
            lngFactor = Round(oTask.UniqueID / 4194304, 0)
            lngFromUID = (lngFactor * 4194304) + oLink.From.UniqueID
          Else
            lngFromUID = oLink.From.UniqueID
          End If
        End If
        If blnMaster And oLink.To.ExternalTask Then
          lngToUID = oLink.To.GetField(185073906) Mod 4194304
          strProject = oLink.To.Project
          If InStr(strProject, "\") > 0 Then
            strProject = Replace(strProject, ".mpp", "")
            strProject = Mid(strProject, InStrRev(strProject, "\") + 1)
          End If
          lngFactor = oSubMap(strProject)
          lngToUID = (lngFactor * 4194304) + lngToUID
        Else
          If blnMaster Then
            lngFactor = Round(oTask.UniqueID / 4194304, 0)
            lngToUID = (lngFactor * 4194304) + oLink.To.UniqueID
          Else
            lngToUID = oLink.To.UniqueID
          End If
        End If
        lngLag = 0
        If oLink.Lag <> 0 Then
          'elapsed lagType properties are even
          'see https://learn.microsoft.com/en-us/office/vba/api/project.pjformatunit
          blnElapsed = False
          If (oLink.LagType Mod 2) = 0 Then
            blnElapsed = True
          End If
          lngLag = oLink.Lag
        End If
        If oLink.To.Calendar = "None" Then
          Set oCalendar = ActiveProject.Calendar
        Else
          Set oCalendar = oLink.To.CalendarObject
        End If
        Select Case oLink.Type
          Case pjFinishToFinish
            'get target successor finish date
            'account for lag
            If blnElapsed Then
              dtDate = DateAdd("n", lngLag, oLink.From.Finish)
            Else
              dtDate = Application.DateAdd(oLink.From.Finish, lngLag, oCalendar)
            End If
            If oLink.To.Finish < dtDate Or IsDate(oLink.To.ActualFinish) Then
              lngOOS = lngOOS + 1
              oOOS.Add oOOS.Count, lngFromUID & "," & lngToUID
              oWorksheet.Cells(lngLastRow, 1) = lngFromUID
              oWorksheet.Cells(lngLastRow, 2) = IIf(blnMaster, "-", oLink.From.ID)
              oWorksheet.Cells(lngLastRow, 3) = oLink.From.Name
              oWorksheet.Cells(lngLastRow, 4) = oLink.From.Finish
              oWorksheet.Cells(lngLastRow, 5) = "FF"
              oWorksheet.Cells(lngLastRow, 6) = oLink.Lag / (8 * 60)
              oWorksheet.Cells(lngLastRow, 7) = lngToUID
              oWorksheet.Cells(lngLastRow, 8) = IIf(blnMaster, "-", oLink.To.ID)
              oWorksheet.Cells(lngLastRow, 9) = oLink.To.Name
              oWorksheet.Cells(lngLastRow, 10) = oLink.To.Finish
              If IsDate(oLink.To.ActualFinish) Then
                oWorksheet.Cells(lngLastRow, 11) = "Successor has Actual Finish"
              Else
                If IsDate(oLink.To.ConstraintDate) And ActiveProject.HonorConstraints Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Finish (has " & Choose(oLink.To.ConstraintType, "", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & " constraint)"
                ElseIf lngLag < 0 Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Finish (has " & Format(lngLag / (60 * 8), "#0d") & " lead)"
                ElseIf IsDate(oLink.To.Deadline) Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Finish (has deadline)"
                Else
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Finish" 'what else would cause this?
                End If
              End If
            End If
          Case pjFinishToStart
            'get target successor start date
            'account for lag
            If blnElapsed Then
              dtDate = DateAdd("n", lngLag, oLink.From.Finish)
            Else
              dtDate = Application.DateAdd(oLink.From.Finish, lngLag, oCalendar)
            End If
            'compare and report
            If oLink.To.Start < dtDate Or IsDate(oLink.To.ActualStart) Then
              lngOOS = lngOOS + 1
              oOOS.Add oOOS.Count, lngFromUID & "," & lngToUID
              oWorksheet.Cells(lngLastRow, 1) = lngFromUID
              oWorksheet.Cells(lngLastRow, 2) = IIf(blnMaster, "-", oLink.From.ID)
              oWorksheet.Cells(lngLastRow, 3) = oLink.From.Name
              oWorksheet.Cells(lngLastRow, 4) = oLink.From.Finish
              oWorksheet.Cells(lngLastRow, 5) = "FS"
              oWorksheet.Cells(lngLastRow, 6) = oLink.Lag / (8 * 60)
              oWorksheet.Cells(lngLastRow, 7) = lngToUID
              oWorksheet.Cells(lngLastRow, 8) = IIf(blnMaster, "-", oLink.To.ID)
              oWorksheet.Cells(lngLastRow, 9) = oLink.To.Name
              oWorksheet.Cells(lngLastRow, 10) = oLink.To.Start
              If IsDate(oLink.To.ActualStart) Then
                oWorksheet.Cells(lngLastRow, 11) = "Successor has Actual Start"
              Else
                If IsDate(oLink.To.ConstraintDate) And ActiveProject.HonorConstraints Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Finish (has " & Choose(oLink.To.ConstraintType, "", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & " constraint)"
                ElseIf lngLag < 0 Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Finish (has " & Format(lngLag / (60 * 8), "#0d") & " lead)"
                ElseIf IsDate(oLink.To.Deadline) Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Finish (has deadline)"
                Else
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Finish" 'what else could cause this?
                End If
              End If
            End If
          Case pjStartToStart
            'get target successor start date
            'account for lag
            If blnElapsed Then
              dtDate = DateAdd("n", lngLag, oLink.From.Start)
            Else
              dtDate = Application.DateAdd(oLink.From.Start, lngLag, oCalendar)
            End If
            'compare and report
            If IsDate(oLink.To.ActualStart) Or oLink.To.Start < dtDate Then 'should not be an issue if both have actual starts
              lngOOS = lngOOS + 1
              oOOS.Add oOOS.Count, lngFromUID & "," & lngToUID
              oWorksheet.Cells(lngLastRow, 1) = lngFromUID
              oWorksheet.Cells(lngLastRow, 2) = IIf(blnMaster, "-", oLink.From.ID)
              oWorksheet.Cells(lngLastRow, 3) = oLink.From.Name
              oWorksheet.Cells(lngLastRow, 4) = oLink.From.Start
              oWorksheet.Cells(lngLastRow, 5) = "SS"
              oWorksheet.Cells(lngLastRow, 6) = oLink.Lag / (8 * 60)
              oWorksheet.Cells(lngLastRow, 7) = lngToUID
              oWorksheet.Cells(lngLastRow, 8) = IIf(blnMaster, "-", oLink.To.ID)
              oWorksheet.Cells(lngLastRow, 9) = oLink.To.Name
              oWorksheet.Cells(lngLastRow, 10) = oLink.To.Start
              If IsDate(oLink.To.ActualStart) Then
                oWorksheet.Cells(lngLastRow, 11) = "Successor has Actual Start"
              Else
                If IsDate(oLink.To.ConstraintDate) And ActiveProject.HonorConstraints Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Start (has " & Choose(oLink.To.ConstraintType, "", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & " constraint)"
                ElseIf lngLag < 0 Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Start (has " & Format(lngLag / (60 * 8), "#0d") & " lead)"
                ElseIf IsDate(oLink.To.Deadline) Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Start (has deadline lead)"
                Else
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Start" 'what else could cause this?
                End If
              End If
            End If
          Case pjStartToFinish
            'this should never happen
            'get target finish
            If blnElapsed Then
              dtDate = DateAdd("n", lngLag, oLink.From.Start)
            Else
              dtDate = Application.DateAdd(oLink.From.Start, lngLag, oCalendar)
            End If
            'compare and report
            If IsDate(oLink.To.ActualFinish) Or oLink.To.Finish < oLink.From.Start Then
              lngOOS = lngOOS + 1
              oOOS.Add oOOS.Count, lngFromUID & "," & lngToUID
              oWorksheet.Cells(lngLastRow, 1) = lngFromUID
              oWorksheet.Cells(lngLastRow, 2) = IIf(blnMaster, "-", oLink.From.ID)
              oWorksheet.Cells(lngLastRow, 3) = oLink.From.Name
              oWorksheet.Cells(lngLastRow, 4) = oLink.From.Start
              oWorksheet.Cells(lngLastRow, 5) = "SF"
              oWorksheet.Cells(lngLastRow, 6) = oLink.Lag / (8 * 60)
              oWorksheet.Cells(lngLastRow, 7) = lngToUID
              oWorksheet.Cells(lngLastRow, 8) = IIf(blnMaster, "-", oLink.To.ID)
              oWorksheet.Cells(lngLastRow, 9) = oLink.To.Name
              oWorksheet.Cells(lngLastRow, 10) = oLink.To.Finish
              If IsDate(oLink.To.ActualFinish) Then
                oWorksheet.Cells(lngLastRow, 11) = "Successor has Actual Finish"
              Else
                If IsDate(oLink.To.ConstraintDate) And ActiveProject.HonorConstraints Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Start (has " & Choose(oLink.To.ConstraintType, "", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & " constraint)"
                ElseIf lngLag < 0 Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Start (has " & Format(lngLag / (60 * 8), "#0d") & " lead)"
                ElseIf IsDate(oLink.To.Deadline) Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Start (has deadline)"
                Else
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Start"
                End If
              End If
            End If
        End Select
      End If
next_link:
    Next oLink
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "[06A212a] Analyzing Out of Sequence Status...(" & Format(lngTask / lngTasks, "0%") & ")" & IIf(lngOOS > 0, " | " & lngOOS & " found", "")
    If myDECM_frm.Visible Then
      myDECM_frm.lblProgress.Width = (lngTask / lngTasks) * myDECM_frm.lblStatus.Width
    End If
    DoEvents
  Next oTask
  
  'only open workbook if OOS oTasks found
  lngOOS = oOOS.Count
  If lngOOS = 0 Then
    oWorkbook.Close False
    GoTo return_val
  Else
    'strOOS = Join(oOOS.Keys, vbTab)
    For lngItem = 0 To lngOOS - 1
      strOOS = strOOS & oOOS.Items(lngItem) & ";"
    Next lngItem
    'todo: delete tmp\06A212a.xlsx on form close
  End If
    
  With oExcel.ActiveWindow
    .Zoom = 85
    .SplitRow = 2
    .SplitColumn = 0
    .FreezePanes = True
  End With
  oWorksheet.Columns.AutoFit
  
  'add a macro
  strMacro = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf
  strMacro = strMacro & " Dim MSPROJ As Object, Task As Object" & vbCrLf
  strMacro = strMacro & " Dim strFrom As String, strTo As String" & vbCrLf
  strMacro = strMacro & "" & vbCrLf
  strMacro = strMacro & " On Error GoTo exit_here" & vbCrLf
  strMacro = strMacro & "" & vbCrLf
  strMacro = strMacro & " If Target.Cells.Count <> 1 Then Exit Sub" & vbCrLf
  strMacro = strMacro & "  If Target.Row < 3 Then Exit Sub" & vbCrLf
  strMacro = strMacro & "  If Target.Column > Me.[A2].End(xlToRight).Column Then Exit Sub" & vbCrLf
  strMacro = strMacro & "  If Target.Row > Me.[A1048576].End(xlUp).Row Then Exit Sub" & vbCrLf
  strMacro = strMacro & "  Set MSPROJ = GetObject(, ""MSProject.Application"")" & vbCrLf
  strMacro = strMacro & "  MSPROJ.ActiveWindow.TopPane.Activate" & vbCrLf
  strMacro = strMacro & "  MSPROJ.ScreenUpdating = False" & vbCrLf
  strMacro = strMacro & "  MSPROJ.FilterClear" & vbCrLf
  strMacro = strMacro & "  MSPROJ.OptionsViewEx DisplaySummaryTasks:=True" & vbCrLf
  strMacro = strMacro & "  MSPROJ.OutlineShowAllTasks" & vbCrLf
  strMacro = strMacro & "  strFrom = Me.Cells(Target.Row, 1).Value" & vbCrLf
  strMacro = strMacro & "  strTo = Me.Cells(Target.Row, 7).Value" & vbCrLf
  strMacro = strMacro & "  MSPROJ.SetAutoFilter FieldName:=""Unique ID"", FilterType:=2, Criteria1:=strFrom & Chr$(9) & strTo '1=pjAutoFilterIn" & vbCrLf
  strMacro = strMacro & "  MSPROJ.Find ""Unique ID"", ""equals"", CLng(strFrom)" & vbCrLf
  strMacro = strMacro & "  MSPROJ.EditGoto Date:=MSPROJ.ActiveProject.Tasks.UniqueID(CLng(strFrom)).Start" & vbCrLf & vbCrLf
  strMacro = strMacro & "exit_here:" & vbCrLf
  strMacro = strMacro & "  MSPROJ.ScreenUpdating = True" & vbCrLf
  strMacro = strMacro & "  Set MSPROJ = Nothing" & vbCrLf
  strMacro = strMacro & "End Sub"
  
  oWorkbook.VBProject.VBComponents("Sheet1").CodeModule.AddFromString strMacro
  
  oExcel.Visible = True
  oExcel.WindowState = xlMinimized
  'todo: 06A212a capture if open and close
  If Dir(Environ("tmp") & "\06A212a.xlsm") <> vbNull Then Kill Environ("tmp") & "\06A212a.xlsm"
  oWorkbook.SaveAs Environ("tmp") & "\06A212a.xlsm", 52 'xlOpenXMLWorkbookMacroEnabled
  oWorkbook.Close
  
return_val:
  cptGetOutOfSequence = CStr(lngOOS) & "|" & strOOS
  
exit_here:
  On Error Resume Next
  'Set myDECM_frm = Nothing 'don't do this
  Set oAssignment = Nothing
  oOOS.RemoveAll
  Set oOOS = Nothing
  Set oCalendar = Nothing
  Set oSubproject = Nothing
  Set oSubMap = Nothing
  Application.StatusBar = ""
  oExcel.EnableEvents = True
  cptSpeed False
  Set oTask = Nothing
  Set oLink = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Exit Function
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptFindOutOfSequence", Err, Erl)
  Resume exit_here
  
End Function

Private Function cptGetEVTAnalysis() As Excel.Workbook
  'objects
  Dim oListObject As Excel.ListObject
  Dim oRange As Excel.Range
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  Dim rst As ADODB.Recordset
  Dim oException As MSProject.Exception
  Dim oCalendar As MSProject.Calendar
  Dim oProject As MSProject.Project
  Dim oTask As MSProject.Task
  'strings
  Dim strMissingBaselines As String
  Dim strLOE As String
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  Dim strFile As String
  'longs
  Dim lngWP As Long
  Dim lngFiscalPeriodsCol As Long
  Dim lngFiscalEndCol As Long
  Dim lngLastRow As Long
  Dim lngFile As Long
  Dim lngEVT As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnExists As Boolean
  'variants
  Dim vbResponse As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oProject = ActiveProject
  
  'ensure project is baselined
  If Not IsDate(oProject.BaselineSavedDate(pjBaseline)) Then
    MsgBox "This project is not yet baselined.", vbCritical + vbOKOnly, "No Baseline"
    GoTo exit_here
  End If
  
'  'ensure fiscal calendar is still loaded
'  If Not cptCalendarExists("cptFiscalCalendar") Then
'    MsgBox "The Fiscal Calendar (cptFiscalCalendar) is missing! Please reset it and try again.", vbCritical + vbOKOnly, "What happened?"
'    GoTo exit_here
'  End If
'
'  'export the calendar
'  Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
'  lngFile = FreeFile
'  strFile = Environ("tmp") & "\fiscal.csv"
'  Open strFile For Output As #lngFile
'  Print #lngFile, "fisc_end,label,"
'  For Each oException In oCalendar.Exceptions
'    Print #lngFile, oException.Finish & "," & oException.Name
'  Next oException
'  Close #lngFile
  
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  oExcel.Visible = True
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "10A103a"
  oWorksheet.[A1:F1] = Split("UID,WP,BLS,BLF,EVT,FiscalPeriods", ",")
  
  Set rst = CreateObject("ADODB.Recordset")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("tmp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT '-' AS [UID],t1.WP,t2.BLS,t2.BLF,t1.EVT "
  strSQL = strSQL & "FROM [tasks.csv] AS t1 "
  strSQL = strSQL & "LEFT JOIN (SELECT WP,MIN(BLS) AS [BLS],MAX(BLF) AS [BLF] FROM [tasks.csv] GROUP BY WP) AS t2 ON T2.WP=T1.WP "
  strSQL = strSQL & "WHERE t1.WP IS NOT NULL "
  strSQL = strSQL & "AND t1.AF IS NULL "
  strSQL = strSQL & "AND t1.EVT='F'" 'todo: what about other ways of identifying 0/100?
  
  rst.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  oWorksheet.[A2].CopyFromRecordset rst
  rst.Close
  
  strSQL = "SELECT * FROM [fiscal.csv]"
  rst.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  oWorksheet.[H1:I1] = Split("fisc_end,label", ",")
  oWorksheet.[H2].CopyFromRecordset rst
  rst.Close
  
  Set oRange = oWorksheet.Range(oWorksheet.[A1].End(xlToRight).Offset(1, 0), oWorksheet.[A1].End(xlDown).Offset(0, 5))
  lngFiscalEndCol = oWorksheet.Rows(1).Find(what:="fisc_end").Column
  lngLastRow = oWorksheet.Cells(2, lngFiscalEndCol).End(xlDown).Row
  'Excel 2016 compatibility
  'oRange.FormulaR1C1 = "=COUNTIFS(R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & ","">=""&RC[-3],R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & ",""<""&RC[-2])+1"
  '=SUMPRODUCT(--($G$2:$G$109>=B15)*--($G$2:$G$109<C15)*1)+1
  oRange.FormulaR1C1 = "=SUMPRODUCT(--(R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & ">=RC[-3])*--(R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & "<RC[-2])*1)+1"
  lngFiscalPeriodsCol = oWorksheet.Rows(1).Find(what:="FiscalPeriods").Column
  oWorksheet.Columns(lngFiscalPeriodsCol).NumberFormat = "#0"
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True
  
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)), , xlYes)
  oListObject.TableStyle = ""
  oWorksheet.[A1].AutoFilter
  oWorksheet.Columns.AutoFit
  On Error Resume Next
  oExcel.Windows("10A103a.xlsx").Close False
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Dir(Environ("tmp") & "\10A103a.xlsx") <> vbNullString Then Kill Environ("tmp") & "\10A103a.xlsx"
  oWorkbook.SaveAs Environ("tmp") & "\10A103a.xlsx", 51
  Set cptGetEVTAnalysis = oWorkbook
  
exit_here:
  On Error Resume Next
  Set oListObject = Nothing
  cptSpeed False
  Set oRange = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Reset 'closes all active files opened by the Open statement and writes the contents of all file buffers to disk.
  If rst.State = 1 Then rst.Close
  Set rst = Nothing
  Set oException = Nothing
  Set oCalendar = Nothing
  Set oTask = Nothing
  Set oProject = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptDECM_bas", "cptGetEVTAnalysis", Err, Erl)
  Resume exit_here
End Function

Private Function cptDECMGetTargetUID() As Long
  'objects
  Dim myDECMTargetUID_frm As cptDECMTargetUID_frm
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strCon As String
  Dim strSQL As String
  'longs
  Dim lngTargetUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  strDir = Environ("tmp")
  If Dir(strDir & "\targets.csv") = vbNullString Then
    cptDECMGetTargetUID = 0
    GoTo exit_here
  End If
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT * FROM [targets.csv] "
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  If oRecordset.RecordCount = 0 Then 'user has no zero-day duration tasks nor milestones
    cptDECMGetTargetUID = 0
    GoTo exit_here
  End If
  oRecordset.MoveFirst
  
  Set myDECMTargetUID_frm = New cptDECMTargetUID_frm
  With myDECMTargetUID_frm
    .lboHeader.Clear
    .lboHeader.AddItem
    .lboHeader.List(0, 0) = "UID"
    .lboHeader.List(0, 1) = "TASK NAME"
    .lboTasks.Clear
    Do While Not oRecordset.EOF
      .lboTasks.AddItem
      .lboTasks.List(.lboTasks.ListCount - 1, 0) = oRecordset("UID")
      .lboTasks.List(.lboTasks.ListCount - 1, 1) = oRecordset("TASK_NAME")
      oRecordset.MoveNext
    Loop
    oRecordset.Close
    .cmdSubmit.Enabled = False
    .Show
    lngTargetUID = myDECMTargetUID_frm.lngTargetTaskUID
  End With
  
  cptDECMGetTargetUID = lngTargetUID
  
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing
  Unload myDECMTargetUID_frm
  Set myDECMTargetUID_frm = Nothing
  
  Exit Function
err_here:
  Call cptHandleErr("cptDECM_bas", "cptDECMGetTargetUID", Err, Erl)
  Resume exit_here
End Function

Function cptGetDECMDescription(strDECM As String) As String
  'macro to create this macro is in "DCMA EVMS Compliance Metrics v6.0 20221205.xlsm"
  Dim strDescription As String
  
  Select Case strDECM
    Case "05A101a"
      strDescription = "Does each control account have exactly one responsible organizational element assigned?" & vbCrLf
      strDescription = strDescription & "X = Count of CAs with more than one OBS element or no OBS elements assigned" & vbCrLf
      strDescription = strDescription & "Y = Total count of CAs"
    
    Case "05A102a"
      strDescription = "Is each control account assigned to a single Control Account Manager (CAM)?" & vbCrLf
      strDescription = strDescription & "X = Count of CAs that have more than one CAM or no CAM assigned" & vbCrLf
      strDescription = strDescription & "Y = Total count of CAs"
    
    Case "05A103a"
      strDescription = "Does each control account have exactly one WBS element assigned?" & vbCrLf
      strDescription = strDescription & "X = Count of CAs with more than one WBS element or no WBS elements assigned" & vbCrLf
      strDescription = strDescription & "Y = Total count of CAs"
    
    Case "06A101a"
      strDescription = "Does each discrete WP, PP, SLPP have task(s) represented in the IMS and EV Cost Tool?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete discrete WPs, PPs, SLPPs in the EV Cost Tool that are not in the IMS + Count of incomplete discrete WPs, PPs, SLPPs in the IMS that are not in the EV Cost Tool" & vbCrLf
      strDescription = strDescription & "Y = Total count of all incomplete discrete WPs, PPs, SLPPs in either the IMS or the EV Cost Tool"
    
    Case "06A204b"
      strDescription = "Are there open starts or finishes ('dangling logic') in the schedule?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete Non-LOE tasks/activities & milestones with open starts or finishes" & vbCrLf
      strDescription = strDescription & "Y = Total count of incomplete Non-LOE tasks/activities & milestones"
    
    Case "06A205a"
      strDescription = "Are lags used in the schedule?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete tasks/activities & milestones with at least one lag in the predecessor logic in the IMS" & vbCrLf
      strDescription = strDescription & "Y = Total count of incomplete tasks/activities & milestones in the IMS"
    
    Case "06A208a"
      strDescription = "Do summary tasks/activities in the schedule have logic applied?" & vbCrLf
      strDescription = strDescription & "X = Count of summary tasks/activities with logic applied (# predecessors > 0 or # successors > 0)"
    
    Case "06A209a"
      strDescription = "Are schedule network constraints limited?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete tasks/activities & milestones with hard constraints" & vbCrLf
      strDescription = strDescription & "Y = Total count of incomplete tasks/activities & milestones"
    
    Case "06A210a"
      strDescription = "Do LOE tasks/activities have discrete successors?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete LOE tasks/activities in the IMS with at least one Non-LOE successor" & vbCrLf
      strDescription = strDescription & "Y = Total count of incomplete LOE tasks/activities in the IMS"
    
    Case "06A211a"
      strDescription = "Is high total float rationale/justification acceptable?" & vbCrLf
      strDescription = "NOTE: X must be determined manually." & vbCrLf
      strDescription = strDescription & "X = Count of high total float (>44 days) non-LOE tasks/activities & milestones sampled with inadequate rationale" & vbCrLf
      strDescription = strDescription & "Y = Total count of high total float non-LOE tasks/activities & milestones sampled"
    
    Case "06A212a"
      strDescription = "Are there out of sequence tasks/activities & milestones?" & vbCrLf
      strDescription = strDescription & "X = Count of out of sequence conditions"
    
    Case "06A401a"
      strDescription = "Does the schedule tool produce a critical path that represents the longest total duration with the least amount of total float?" & vbCrLf
      strDescription = strDescription & "X = Count of tasks/activities & milestones on the constraint method critical path that are not on the contractor's critical path"
    
    Case "06A501a"
      strDescription = "In the IMS, do all of the tasks/activities & milestones have baseline start and baseline finish dates?" & vbCrLf
      strDescription = strDescription & "X = Count of tasks/activities & milestones without baseline dates" & vbCrLf
      strDescription = strDescription & "Y = Total count of tasks/activities & milestones"
    
    Case "06A504a"
      strDescription = "Are actual start dates changed after first reported?" & vbCrLf
      strDescription = strDescription & "X = Count of tasks/activities & milestones where actual start date does not equal previously reported actual start date" & vbCrLf
      strDescription = strDescription & "Y = Total count of tasks/activities & milestones with actual start dates"
    
    Case "06A504b"
      strDescription = "Are actual finish dates changed after first reported?" & vbCrLf
      strDescription = strDescription & "X = Count of tasks/activities & milestones where actual finish date does not equal previously reported actual finish date" & vbCrLf
      strDescription = strDescription & "Y = Total count of tasks/activities & milestones with actual finish dates"
    
    Case "06A505a"
      strDescription = "Do all in progress tasks/activities & milestones have actual start dates?" & vbCrLf
      strDescription = strDescription & "X = Count of in progress tasks/activities & milestones with no actual start date" & vbCrLf
      strDescription = strDescription & "Y = Total count of in progress tasks/activities & milestones"
    
    Case "06A505b"
      strDescription = "Do all complete tasks/activities & milestones have actual finish dates?" & vbCrLf
      strDescription = strDescription & "X = Count of complete tasks/activities & milestones with no actual finish date" & vbCrLf
      strDescription = strDescription & "Y = Total count of complete tasks/activities & milestones"
    
    Case "06A506a"
      strDescription = "Are actual start and actual finish dates valid for all tasks/activities & milestones in the IMS?" & vbCrLf
      strDescription = strDescription & "X = Count of tasks/activities & milestones with either actual start or actual finish after status date" & vbCrLf
      strDescription = strDescription & "Y = Total count of tasks/activities & milestones with an actual start date"
    
    Case "06A506b"
      strDescription = "Are forecast start and finish dates valid for all tasks/activities & milestones in the IMS?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete tasks/activities & milestones with either forecast start or forecast finish before the status date"
    
    Case "06A506c"
      strDescription = "Are forecast start/finish dates riding the status date of the IMS for two consecutive months?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete tasks/activities & milestones with either forecast start or forecast finish date riding the status date" & vbCrLf
      strDescription = strDescription & "Y = Total count of incomplete tasks/activities & milestones"
    
    Case "06I201a"
      strDescription = "Are Schedule Visibility Tasks (SVTs) identified and controlled in the IMS?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete tasks/activities and [milestones] that are not properly identified and controlled as SVTs in the IMS"
    
    Case "10A102a"
      strDescription = "NOTIONAL ONLY: RUN IN EV COST TOOL" & vbCrLf
      strDescription = strDescription & "Is each Work Package assigned a single EVT?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete WPs that have more than one EVT or no EVT assigned" & vbCrLf
      strDescription = strDescription & "Y = Total count of incomplete WPs"
    
    Case "10A103a"
      strDescription = "NOTIONAL ONLY: RUN IN EV COST TOOL" & vbCrLf
      strDescription = strDescription & "Are 0-100 EVTs applied to incomplete WPs with one accounting period of budget?" & vbCrLf
      strDescription = strDescription & "X = Count of 0-100 EVT incomplete WPs with more than one accounting period of budget" & vbCrLf
      strDescription = strDescription & "Y = Total count of 0-100 EVT incomplete WPs"
    
    Case "10A109b"
      strDescription = "Does each WP/PP/SLPPs have an assigned budget?" & vbCrLf
      strDescription = strDescription & "X = Count of WPs/PPs/SLPPs with BAC = 0" & vbCrLf
      strDescription = strDescription & "Y = Total count of WPs/PPs/SLPPs"
    
    Case "10A302b"
      strDescription = "Have PPs earned performance?" & vbCrLf
      strDescription = strDescription & "X = Count of PPs with BCWPCUM" & vbCrLf
      strDescription = strDescription & "Y = Total count of PPs"
    
    Case "10A303a"
      strDescription = "Do all PPs have duration?" & vbCrLf
      strDescription = strDescription & "X = Count of PPs (tasks/activities & milestones level) with baseline duration less than or equal to one day" & vbCrLf
      strDescription = strDescription & "Y = Total count of PPs (tasks/activities & milestones level)"
    
    Case "11A101a"
      strDescription = "For all CAs, does the BAC value for the CA equate to the sum of the WP and PP budgets within the CA?" & vbCrLf
      strDescription = strDescription & "X = Sum of the absolute values of (CA BAC - the sum of its WP and PP budgets)" & vbCrLf
      strDescription = strDescription & "Y = Total program BAC"
    
    Case "29A601a"
      strDescription = "Is all effort detailed planned within the current rolling wave/freeze period?" & vbCrLf
      strDescription = strDescription & "X = Count of PPs/SLPPs where baseline start precedes the next rolling wave cycle" & vbCrLf
      strDescription = strDescription & "Y = Total count of PPs/SLPPs"
    
    Case "CPT01"
      strDescription = "Do all PMB Tasks (where Baseline Work > 0h or Baseline Cost > $0) have appropriate metadata?" & vbCrLf
      strDescription = strDescription & "X = Count of PMB Tasks with missing WBS, OBS, CA, CAM, WP, or EVT" & vbCrLf
      strDescription = strDescription & "Y = (not used)"
      
    Case "CPT02"
      strDescription = "Is each Work Package assigned a single Control Account?" & vbCrLf
      strDescription = strDescription & "X = Count of incomplete WPs that are assigned to more than one CA or no CA assigned" & vbCrLf
      strDescription = strDescription & "Y = Total count incomplete WPs"
    
    Case "CPT03"
      strDescription = "Are there any leads (negative lags)?" & vbCrLf
      strDescription = strDescription & "X = Count of dependencies with leads (negative lag)"
      
    Case Else
      strDescription = "No Description provided."
      
  End Select
  
  cptGetDECMDescription = strDescription
  
End Function

Function cptDECMDatabaseExists() 'todo: make private
  cptDECMDatabaseExists = Dir(cptDir & "\decm\decm-v.6.0.csv") <> vbNullString
End Function

Sub cptWriteDECMDataBase()
  'todo: confirm decm directory
  'todo: write Schema:
  'ID
  'DEFINITION
  'NUMERATOR
  'DENOMINATOR
  'THRESHOLD
  'todo: write csv
  'todo: enable diff (added, removed, changed)
End Sub
