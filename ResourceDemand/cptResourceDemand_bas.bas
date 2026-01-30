Attribute VB_Name = "cptResourceDemand_bas"
'<cpt_version>v1.6.0</cpt_version>
Option Explicit
Private Const MODULE_NAME = "cptResourceDemand_bas"

Sub cptExportResourceDemand(ByRef myResourceDemand_frm As cptResourceDemand_frm, Optional lngTaskCount As Long)
  'objects
  Dim oEstimates As Scripting.Dictionary
  Dim oCalendar As MSProject.Calendar
  Dim oRecordset As ADODB.Recordset
  Dim oException As MSProject.Exception
  Dim oSettings As Object
  Dim oListObject As Excel.ListObject
  Dim oSubproject As MSProject.SubProject
  Dim oTask As MSProject.Task
  Dim oResource As MSProject.Resource
  Dim oAssignment As MSProject.Assignment
  Dim oTSV As TimeScaleValue
  Dim oTSVS_BCWS As TimeScaleValues
  Dim oTSVS_WORK As TimeScaleValues
  Dim oTSVS_AW As TimeScaleValues
  Dim oTSVS_COST As TimeScaleValues
  Dim oTSVS_AC As TimeScaleValues
  Dim oCostRateTable As CostRateTable
  Dim oPayRate As PayRate
  Dim oExcel As Excel.Application 'Object
  Dim oWorksheet As Excel.Worksheet 'Object
  Dim oWorkbook As Excel.Workbook 'Object
  Dim oRange As Excel.Range 'Object
  Dim oPivotTable As Excel.PivotTable 'Object
  Dim oPivotChartTable As Excel.PivotTable
  Dim oChart As Excel.Chart
  'dates
  Dim dtWeek As Date
  Dim dtStart As Date
  Dim dtFinish As Date
  'doubles
  Dim dblWork As Double
  Dim dblCost As Double
  'strings
  Dim strCFN As String
  Dim strTask As String
  Dim strFields As String
  Dim strRateSets As String
  Dim strMsg As String
  Dim strSettings As String
  Dim strKey As String
  Dim strView As String
  Dim strFileName As String
  Dim strRange As String
  Dim strHeader As String
  'longs
  Dim lngLastCol As Long
  Dim lngItem As Long
  Dim lngCols As Long
  Dim lngLastRow As Long
  Dim lngFiscalMonthCol As Long
  Dim lngHoursCol As Long
  Dim lngRateSets As Long
  Dim lngCol As Long
  Dim lngOriginalRateSet As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  Dim lngWeekCol As Long
  Dim lngExport As Long
  Dim lngField As Long
  Dim lngRateSet As Long
  Dim lngRow As Long
  'variants
  Dim vParts As Variant
  Dim aResult() As Variant
  Dim vKey As Variant
  Dim vData As Variant
  Dim vChk As Variant
  Dim vRateSet As Variant
  Dim aUserFields() As Variant
  Dim vFiscalCalendar As Variant
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnFiscal As Boolean
  Dim blnExportAssociatedBaseline As Boolean
  Dim blnExportFullBaseline As Boolean
  Dim blnIncludeCosts As Boolean
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Application.StatusBar = "Confirming Status Date..."
  myResourceDemand_frm.lblStatus.Caption = "Confirming Status Date..."
  
  If IsDate(ActiveProject.StatusDate) Then
    dtStart = ActiveProject.StatusDate
    If ActiveProject.ProjectStart > dtStart Then dtStart = ActiveProject.ProjectStart
  Else
    Application.StatusBar = "Please enter a Status Date."
    MsgBox "Please enter a Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
    Application.StatusBar = ""
    GoTo exit_here
  End If

  'save settings, build header
  strHeader = "PROJECT,"
  With myResourceDemand_frm
    Application.StatusBar = "Saving user settings..."
    aUserFields = .lboExport.List()
    For lngExport = 0 To UBound(aUserFields, 1)
      lngField = aUserFields(lngExport, 0)
      strCFN = CustomFieldGetName(lngField)
      If Len(strCFN) > 0 Then
        strHeader = strHeader & UCase(strCFN) & ","
      Else
        strHeader = strHeader & UCase(FieldConstantToFieldName(lngField)) & ","
      End If
    Next lngExport
    strHeader = strHeader & "[UID] TASK,RESOURCE_NAME,CLASS,"
    .lblStatus.Caption = "Saving user settings..."
    cptSaveSetting "ResourceDemand", "cboMonths", .cboMonths.Value
    blnFiscal = .cboMonths.Value = 1
    cptSaveSetting "ResourceDemand", "cboWeeks", .cboWeeks.Value
    cptSaveSetting "ResourceDemand", "cboWeekday", .cboWeekday.Value
    cptSaveSetting "ResourceDemand", "chkCosts", IIf(.chkCosts, 1, 0)
    blnIncludeCosts = .chkCosts
    If blnIncludeCosts Then
      lngItem = 0
      For Each vChk In Split("A,B,C,D,E", ",")
        strRateSets = strRateSets & IIf(.Controls("chk" & vChk), lngItem & ",", "")
        lngItem = lngItem + 1
      Next
      If Len(strRateSets) > 0 Then strRateSets = Left(strRateSets, Len(strRateSets) - 1)
      lngRateSets = UBound(Split(strRateSets, ",")) + 1
      cptSaveSetting "ResourceDemand", "CostSets", strRateSets
      strHeader = strHeader & "RATE_TABLE,ACTIVE,"
    End If
    If blnFiscal Then
      strHeader = strHeader & "FISCAL_MONTH,"
    Else
      strHeader = strHeader & "WEEK,MONTH,"
    End If
    strHeader = strHeader & "HOURS"
    If blnIncludeCosts Then
      strHeader = strHeader & ",COST"
    End If
    cptDeleteSetting "ResourceDemand", "chkBaseline"
    blnExportAssociatedBaseline = .chkAssociatedBaseline = True
    cptSaveSetting "ResourceDemand", "chkAssociatedBaseline", IIf(blnExportAssociatedBaseline, 1, 0)
    blnExportFullBaseline = .chkFullBaseline = True
    cptSaveSetting "ResourceDemand", "chkFullBaseline", IIf(blnExportFullBaseline, 1, 0)
    cptDeleteSetting "ResourceDemand", "chkNonLabor"
  End With
  
  strFileName = cptDir & "\settings\cpt-export-resource-userfields.adtg."
  Set oSettings = CreateObject("ADODB.Recordset")
  With oSettings
    .Fields.Append "Field Constant", adVarChar, 255
    .Fields.Append "Custom Field Name", adVarChar, 255
    .Open
    strSettings = "Week=" & myResourceDemand_frm.cboWeeks & ";"
    strSettings = strSettings & "Weekday=" & myResourceDemand_frm.cboWeekday & ";"
    strSettings = strSettings & "Costs=" & myResourceDemand_frm.chkCosts & ";"
    strSettings = strSettings & "AssociatedBaseline=" & blnExportAssociatedBaseline & ";"
    strSettings = strSettings & "FullBaseline=" & blnExportFullBaseline & ";"
    strSettings = strSettings & "RateSets="
    For Each vChk In Split("A,B,C,D,E", ",")
      strFields = strFields & IIf(myResourceDemand_frm.Controls("chk" & vChk), vChk & ",", "")
    Next vChk
    .AddNew Array(0, 1), Array("settings", strSettings)
    .Update
    'save userfields
    For lngExport = 0 To UBound(aUserFields, 1)
      .AddNew Array(0, 1), Array(aUserFields(lngExport, 0), aUserFields(lngExport, 1))
      .Update
    Next lngExport
    If Dir(strFileName) <> vbNullString Then Kill strFileName
    .Save strFileName, adPersistADTG
    .Close
  End With
  
  Application.StatusBar = "Preparing to export..."
  myResourceDemand_frm.lblStatus.Caption = "Preparing to export..."
  
  If ActiveProject.Subprojects.Count = 0 Then
    lngTasks = ActiveProject.Tasks.Count
  Else
    cptSpeed True
    strView = ActiveWindow.TopPane.View.Name
    ViewApply "Gantt Chart"
    FilterClear
    GroupClear
    SelectAll
    OptionsViewEx DisplaySummaryTasks:=True
    On Error Resume Next
    If Not OutlineShowAllTasks Then
      Sort "ID", , , , , , False, True
      OutlineShowAllTasks
    End If
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    SelectAll
    lngTasks = ActiveSelection.Tasks.Count
    ViewApply strView
    cptSpeed False
  End If
  
  If blnFiscal Then 'get the fiscal calendar
    Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
    ReDim vFiscalCalendar(0 To 1, 0 To oCalendar.Exceptions.Count)
    For Each oException In oCalendar.Exceptions
      vFiscalCalendar(0, oException.Index) = oException.Start
      vFiscalCalendar(1, oException.Index) = oException.Name
    Next oException
  End If
  
  'Key=PROJECT|{USER_FIELD}|[UID] TASK|RESOURCE_NAME|CLASS|COST_SET|ACTIVE|MONTH
  'Value=HOURS|COST
  
  'iterate over tasks
  Application.StatusBar = "Getting Excel..."
  myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  'set reference to Excel
'  On Error Resume Next
'  Set oExcel = GetObject(, "Excel.Application")
'  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
'  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
'  End If
  
  Set oEstimates = CreateObject("Scripting.Dictionary")
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task 'skip blank lines
    If oTask.ExternalTask Then GoTo next_task 'skip external tasks
    If oTask.Summary Then GoTo next_task 'skip summary task
    If Not oTask.Active Then GoTo next_task 'skip inactive tasks
    If Not blnExportAssociatedBaseline And Not blnExportFullBaseline Then
      If oTask.RemainingDuration = 0 Then GoTo next_task
    End If
    
    'capture oTask data common to all oAssignments
    strTask = oTask.Project
    
    'get custom field values
    For lngExport = 0 To UBound(aUserFields, 1) 'myResourceDemand_frm.lboExport.ListCount - 1
      lngField = aUserFields(lngExport, 0)
      strTask = strTask & "|" & Trim(Replace(oTask.GetField(lngField), "|", "-"))
    Next lngExport
    
    strTask = strTask & "|[" & oTask.UniqueID & "] " & Replace(Replace(oTask.Name, "|", "-"), Chr(34), Chr(39))
    
    'examine every oAssignment on the task
    For Each oAssignment In oTask.Assignments
      
      'capture original rate set
      lngOriginalRateSet = oAssignment.CostRateTable
      
      'skip non-labor entirely
      If oAssignment.ResourceType <> pjResourceTypeWork Then GoTo next_assignment 'skip non-labor entirely
      
      'skip completed tasks for ETC
      If IsDate(oTask.ActualFinish) Then GoTo export_baseline 'NOT Exit For
      
      
      'capture remaining work (ETC)
      If IsDate(oTask.Stop) Then 'capture the unstatused / remaining portion
        dtStart = oTask.Resume
      Else 'capture the entire unstarted task
        dtStart = oTask.Start
      End If
      dtFinish = oTask.Finish
      
      If blnFiscal Then
        'Set oTSVS_WORK = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleDays, 1)
        Set oTSVS_WORK = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
      Else
        Set oTSVS_WORK = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
      End If
      
      For Each oTSV In oTSVS_WORK
        If Val(oTSV.Value) = 0 Then GoTo next_tsv_etc
        'capture common oAssignment data
        strKey = strTask & "|" & oAssignment.ResourceName & "|ETC" 'keep this here
        'capture (and subtract) actual work, leaving ETC/Remaining Work
        If blnFiscal Then
          'Set oTSVS_AW = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleDays, 1)
          Set oTSVS_AW = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
        Else
          Set oTSVS_AW = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
        End If
        dblWork = (Val(oTSV.Value) - Val(oTSVS_AW(1))) / 60
        If dblWork = 0 Then GoTo next_tsv_etc
        
        If blnIncludeCosts Then
          strKey = strKey & "|" & Choose(oAssignment.CostRateTable + 1, "A", "B", "C", "D", "E")
          strKey = strKey & "|TRUE"
        End If
        
        If blnFiscal Then
          strKey = strKey & "|" & cptGetFiscalMonthOfDay(oTSV.StartDate, vFiscalCalendar)
        Else
          'apply user settings for week identification
          With myResourceDemand_frm
            If .cboWeeks = "Beginning" Then
              If .cboWeekday = "Monday" Then
                dtWeek = DateAdd("d", 2 - Weekday(oTSV.StartDate), oTSV.StartDate)
              End If
            ElseIf .cboWeeks = "Ending" Then
              If .cboWeekday = "Friday" Then
                dtWeek = DateAdd("d", 6 - Weekday(oTSV.StartDate), oTSV.StartDate)
              ElseIf .cboWeekday = "Saturday" Then
                dtWeek = DateAdd("d", 7 - Weekday(oTSV.StartDate), oTSV.StartDate)
              End If
            End If
          End With
          strKey = strKey & "|" & dtWeek & "|" & Format(dtWeek, "yyyymm")
        End If
        
        'add work without cost yet
        If oEstimates.Exists(strKey) Then
          If blnIncludeCosts Then
            dblWork = dblWork + Split(oEstimates(strKey), "|")(0) 'add
            dblCost = Split(oEstimates(strKey), "|")(1) 'keep
            oEstimates(strKey) = dblWork & "|" & dblCost
          Else
            oEstimates(strKey) = oEstimates(strKey) + dblWork
          End If
        Else
          If blnIncludeCosts Then
            oEstimates.Add strKey, dblWork & "|" & 0 'dblCost
          Else
            oEstimates.Add strKey, dblWork
          End If
        End If
        
        'get default costs
        If blnIncludeCosts Then
          'get active cost
          If blnFiscal Then
            'Set oTSVS_COST = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleDays, 1)
            Set oTSVS_COST = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
            'get actual cost
            'Set oTSVS_AC = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleDays, 1)
            Set oTSVS_AC = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
          Else
            Set oTSVS_COST = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
            'get actual cost
            Set oTSVS_AC = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
          End If
          'subtract actual cost from cost to get remaining cost
          dblCost = Val(oTSVS_COST(1).Value) - Val(oTSVS_AC(1))
          
          'add cost without work
          If oEstimates.Exists(strKey) Then
            If blnIncludeCosts Then
              dblWork = Split(oEstimates(strKey), "|")(0) 'keep
              dblCost = dblCost + Split(oEstimates(strKey), "|")(1) 'add
              oEstimates(strKey) = dblWork & "|" & dblCost
            'Else
              'oEstimates(strKey) = oEstimates(strKey) + dblWork
            End If
          Else
            If blnIncludeCosts Then
              'Stop 'uh oh
              oEstimates.Add strKey, 0 & "|" & dblCost 'this should never happen
            'Else
              'oEstimates.Add strKey, dblWork
            End If
          End If
        End If
          
next_tsv_etc:
      Next oTSV
      
      If lngRateSets > 0 Then
        'silly to have to repeat it, but changing cost rate tables is expensive
        'better to do it once per rate table, per assignment
        'than to do it once per rate table, per assignment, per timescalevalue
        For Each vRateSet In Split(strRateSets, ",")
          If CLng(vRateSet) = lngOriginalRateSet Then GoTo next_rate_set
          
          For Each oTSV In oTSVS_WORK
            'capture common oAssignment data
            strKey = strTask & "|" & oAssignment.ResourceName & "|ETC" 'keep this here
            'capture (and subtract) actual work, leaving ETC/Remaining Work
            If blnFiscal Then
              'Set oTSVS_AW = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleDays, 1)
              Set oTSVS_AW = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
            Else
              Set oTSVS_AW = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
            End If
            dblWork = (Val(oTSV.Value) - Val(oTSVS_AW(1))) / 60
            If dblWork = 0 Then GoTo next_tsv_rs

            If blnIncludeCosts Then
              strKey = strKey & "|" & Choose(CLng(vRateSet) + 1, "A", "B", "C", "D", "E")
              strKey = strKey & "|FALSE"
            End If
            
            If blnFiscal Then
              strKey = strKey & "|" & cptGetFiscalMonthOfDay(oTSV.StartDate, vFiscalCalendar)
            Else
              'apply user settings for week identification
              With myResourceDemand_frm
                If .cboWeeks = "Beginning" Then
                  If .cboWeekday = "Monday" Then
                    dtWeek = DateAdd("d", 2 - Weekday(oTSV.StartDate), oTSV.StartDate)
                  End If
                ElseIf .cboWeeks = "Ending" Then
                  If .cboWeekday = "Friday" Then
                    dtWeek = DateAdd("d", 6 - Weekday(oTSV.StartDate), oTSV.StartDate)
                  ElseIf .cboWeekday = "Saturday" Then
                    dtWeek = DateAdd("d", 7 - Weekday(oTSV.StartDate), oTSV.StartDate)
                  End If
                End If
              End With
              strKey = strKey & "|" & dtWeek & "|" & Format(dtWeek, "yyyymm")
            End If

            'add work without cost yet
            If oEstimates.Exists(strKey) Then
              If blnIncludeCosts Then
                dblWork = dblWork + Split(oEstimates(strKey), "|")(0) 'add
                dblCost = Split(oEstimates(strKey), "|")(1) 'keep
                oEstimates(strKey) = dblWork & "|" & dblCost
              Else
                oEstimates(strKey) = oEstimates(strKey) + dblWork
              End If
            Else
              If blnIncludeCosts Then
                oEstimates.Add strKey, dblWork & "|" & 0 'dblCost
              Else
                oEstimates.Add strKey, dblWork
              End If
            End If
        
            'get active cost
            If oAssignment.CostRateTable <> CLng(vRateSet) Then oAssignment.CostRateTable = CLng(vRateSet) 'very expensive
            If blnFiscal Then
              'Set oTSVS_COST = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleDays, 1)
              Set oTSVS_COST = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
              'get actual cost
              'Set oTSVS_AC = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleDays, 1)
              Set oTSVS_AC = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
            Else
              Set oTSVS_COST = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
              'get actual cost
              Set oTSVS_AC = oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
            End If
            'subtract actual cost from cost to get remaining cost
            dblCost = Val(oTSVS_COST(1).Value) - Val(oTSVS_AC(1))
            
            'add cost without work
            If oEstimates.Exists(strKey) Then
              If blnIncludeCosts Then
                dblWork = Split(oEstimates(strKey), "|")(0) 'keep
                dblCost = dblCost + Split(oEstimates(strKey), "|")(1) 'add
                oEstimates(strKey) = dblWork & "|" & dblCost
              End If
            Else
              'this should never happen
            End If
        
next_tsv_rs:
          Next oTSV
next_rate_set:
        Next vRateSet
        If oAssignment.CostRateTable <> lngOriginalRateSet Then oAssignment.CostRateTable = lngOriginalRateSet
      End If
      
export_baseline:
      If blnExportAssociatedBaseline Or blnExportFullBaseline Then
        dtStart = oExcel.WorksheetFunction.Min(oTask.Start, IIf(oTask.BaselineStart = "NA", oTask.Start, oTask.BaselineStart)) 'works with forecast, actual, and baseline start
        dtFinish = oExcel.WorksheetFunction.Max(oTask.Finish, IIf(oTask.BaselineFinish = "NA", oTask.Finish, oTask.BaselineFinish)) 'works with forecast, actual, and baseline finish
        'Set oTSVS_BCWS = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleDays, 1)
        Set oTSVS_BCWS = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks, 1)
        For Each oTSV In oTSVS_BCWS
          If Val(oTSV.Value) = 0 Then GoTo next_tsv_bcws
          strKey = strTask & "|" & oAssignment.ResourceName & "|BCWS" 'keep this here
          dblWork = Val(oTSV.Value) / 60
          If blnIncludeCosts Then
            strKey = strKey & "|BASELINED|TRUE"
            'dblCost = Val(oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledBaselineCost, pjTimescaleDays, 1)(1).Value)
            dblCost = Val(oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledBaselineCost, pjTimescaleWeeks, 1)(1).Value)
          End If
          If blnFiscal Then
            'get fiscal month of day
            strKey = strKey & "|" & cptGetFiscalMonthOfDay(oTSV.StartDate, vFiscalCalendar)
          Else
            'apply user settings for week identification
            With myResourceDemand_frm
              If .cboWeeks = "Beginning" Then
                If .cboWeekday = "Monday" Then
                  dtWeek = DateAdd("d", 2 - Weekday(oTSV.StartDate), oTSV.StartDate)
                End If
              ElseIf .cboWeeks = "Ending" Then
                If .cboWeekday = "Friday" Then
                  dtWeek = DateAdd("d", 6 - Weekday(oTSV.StartDate), oTSV.StartDate)
                ElseIf .cboWeekday = "Saturday" Then
                  dtWeek = DateAdd("d", 7 - Weekday(oTSV.StartDate), oTSV.StartDate)
                End If
              End If
            End With
            strKey = strKey & "|" & dtWeek & "|" & Format(dtWeek, "yyyymm")
          End If
          If oEstimates.Exists(strKey) Then
            If blnIncludeCosts Then
              dblWork = dblWork + Split(oEstimates(strKey), "|")(0)
              dblCost = dblCost + Split(oEstimates(strKey), "|")(1)
              oEstimates(strKey) = dblWork & "|" & dblCost
            Else
              dblWork = dblWork + Split(oEstimates(strKey), "|")(0)
              oEstimates(strKey) = dblWork
            End If
          Else
            If blnIncludeCosts Then
              oEstimates.Add strKey, dblWork & "|" & dblCost
            Else
              oEstimates.Add strKey, dblWork
            End If
          End If
next_tsv_bcws:
        Next oTSV
      End If
next_assignment:
      'restore original rate set
      If oAssignment.CostRateTable <> lngOriginalRateSet Then oAssignment.CostRateTable = lngOriginalRateSet
    Next oAssignment
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Exporting " & Format(lngTask, "#,##0") & " of " & Format(lngTasks, "#,##0") & "...(" & Format(lngTask / lngTasks, "0%") & ")"
    myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
    myResourceDemand_frm.lblProgress.Width = (lngTask / lngTasks) * myResourceDemand_frm.lblStatus.Width
    DoEvents
  Next oTask

  If oEstimates.Count > 0 Then
    Application.StatusBar = "Creating Workbook..."
    myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    'header
    oWorksheet.[A1].Resize(1, UBound(Split(strHeader, ",")) + 1) = Split(strHeader, ",")
    'data
    ReDim aResult(1 To oEstimates.Count, 1 To UBound(Split(strHeader, ",")) + 1)
    oWorksheet.[A1].AutoFilter
    lngRow = 1
    lngCols = UBound(Split(strHeader, ",")) + 1
    For Each vKey In oEstimates.Keys
      vParts = Split(vKey, "|")
      For lngCol = 1 To (UBound(vParts, 1) + 1)
        aResult(lngRow, lngCol) = vParts(lngCol - 1)
      Next lngCol
      If blnIncludeCosts Then
        aResult(lngRow, lngCols - 1) = Split(oEstimates(vKey), "|")(0)
        aResult(lngRow, lngCols) = Split(oEstimates(vKey), "|")(1)
      Else
        aResult(lngRow, lngCols) = oEstimates(vKey)
      End If
      lngRow = lngRow + 1
    Next vKey
    oWorksheet.[A2].Resize(UBound(aResult, 1), UBound(aResult, 2)).Value = aResult
  End If
  
'  'is previous run still open?
'  On Error Resume Next
'  strFileName = Environ("TEMP") & "\ExportResourceDemand.xlsx"
'  Set oWorkbook = oExcel.oWorkbooks(strFileName)
'  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
'  If Not oWorkbook Is Nothing Then oWorkbook.Close False
'  On Error Resume Next
'  Set oWorkbook = oExcel.Workbooks(Environ("TEMP") & "\ExportResourceDemand.xlsx")
'  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
'  If Not oWorkbook Is Nothing Then 'add timestamp to existing file
'    If oWorkbook.Application.Visible = False Then oWorkbook.Application.Visible = True
'    strMsg = "'" & strFileName & "' already exists and is open."
'    strFileName = Replace(strFileName, ".xlsx", "_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".xlsx")
'    strMsg = strMsg & "Your new file will be saved as:" & vbCrLf & strFileName
'    MsgBox strMsg, vbExclamation + vbOKOnly, "File Exists and is Open"
'  End If
  
  Application.StatusBar = "Saving workbook..."
  myResourceDemand_frm.lblStatus.Caption = Application.StatusBar

  On Error Resume Next
  If oWorkbook Is Nothing Then GoTo exit_here 'todo
  If Dir(Environ("TEMP") & "\ExportResourceDemand.xlsx") <> vbNullString Then Kill Environ("TEMP") & "\ExportResourceDemand.xlsx"
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  MsgBox "If your company requires security classifications, please make them from within the Excel Window.", vbExclamation + vbOKOnly, "Heads up"
  oExcel.Visible = True
  oExcel.WindowState = xlNormal
  If Dir(Environ("TEMP") & "\ExportResourceDemand.xlsx") <> vbNullString Then 'kill failed, rename it
    oWorkbook.SaveAs Environ("TEMP") & "\ExportResourceDemand_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".xlsx", 51
  Else
    oWorkbook.SaveAs Environ("TEMP") & "\ExportResourceDemand.xlsx", 51
  End If
  oExcel.Visible = False
    
  If blnFiscal Then
    Application.StatusBar = "Extracting Fiscal Periods..."
    myResourceDemand_frm.lblStatus.Caption = "Extracting Fiscal Periods..."
    Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
    With oWorksheet
      .Name = "FiscalPeriods"
      .[A1:B1] = Array("fisc_end", "label")
      .[A2].Resize(UBound(vFiscalCalendar, 2), UBound(vFiscalCalendar, 1) + 1).Value = oExcel.WorksheetFunction.Transpose(vFiscalCalendar)
      Set oListObject = .ListObjects.Add(xlSrcRange, .Range(.[A1].End(xlToRight), .[A1].End(xlDown)), , xlYes)
      oListObject.Name = "FISCAL"
      'add Holidays table
      .[E1] = "EXCEPTIONS"
    End With
    'just...go get the exceptions
    Set oCalendar = ActiveProject.Calendar
    If oCalendar.Exceptions.Count > 0 Then
      Set oWorksheet = oWorkbook.Worksheets.Add(After:=oWorkbook.Worksheets(oWorkbook.Worksheets.Count))
      oWorksheet.Name = "Exceptions"
      Set oWorksheet = oWorkbook.Worksheets.Add(After:=oWorkbook.Worksheets(oWorkbook.Worksheets.Count))
      oWorksheet.Name = "WorkWeeks"
      cptExportCalendarExceptions oWorkbook, oCalendar, True
      With oWorksheet
        .Activate
        oExcel.ActiveWindow.Zoom = 85
        .Columns.AutoFit
      End With
      Set oWorksheet = oWorkbook.Worksheets("Exceptions")
      With oWorksheet
        .Activate
        oExcel.ActiveWindow.Zoom = 85
        .Columns.AutoFit
        .Outline.ShowLevels Rowlevels:=1
      End With
      Set oWorksheet = oWorkbook.Worksheets("FiscalPeriods")
      With oWorksheet
        .Activate
        .[E2].Formula2 = "=UNIQUE(Exceptions!" & oWorkbook.Sheets("Exceptions").Range(oWorkbook.Sheets("Exceptions").[C2], oWorkbook.Sheets("Exceptions").[C2].End(xlDown)).Address & ")"
        .Range(.[E2], .[E2].End(xlDown)).NumberFormat = "m/d/YYYY"
        vData = .Range(.[E2], .[E2].End(xlDown))
        .Range(.[E2], .[E2].End(xlDown)) = vData
        'convert to a table
        Set oListObject = .ListObjects.Add(xlSrcRange, .Range(.[E1], .[E2].End(xlDown)), , xlYes)
        'reset oCalendar
        Set oCalendar = ActiveProject.Calendar
        .Columns(6).ColumnWidth = 1
        .[G3] = "Fiscal periods imported from 'cptFiscalCalendar'"
        With .[G3:L3]
          .Merge
          .HorizontalAlignment = xlCenter
          .Style = "Note"
        End With
        .[G4] = "Exceptions imported from '" & oCalendar.Name & "'"
        With .[G4:L4]
          .Merge
          .HorizontalAlignment = xlCenter
          .Style = "Note"
        End With
      End With
    Else
      'convert to a table
      Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[E1], oWorksheet.[E2]), , xlYes)
    End If
    oExcel.ActiveWindow.DisplayGridlines = False
    oExcel.ActiveWindow.Zoom = 85
    oListObject.Name = "EXCEPTIONS"
    'add efficiency factor entry
    With oWorksheet.[G1]
      .Value = "Efficiency:"
      .EntireColumn.AutoFit
    End With
    With oWorksheet
      With .[H1]
        .Value = 1
        .Style = "Percent"
        .Style = "Input"
      End With
      .Names.Add "efficiency_factor", .[H1]
    End With
    'add HPM formula
    Application.StatusBar = "Calculating HPM..."
    myResourceDemand_frm.lblStatus.Caption = "Calculating HPM..."
    oWorksheet.[C1].Value = "hpm"
    oWorksheet.[C3].Formula = "=IFERROR(NETWORKDAYS(A2+1,[@[fisc_end]],EXCEPTIONS)*(8*efficiency_factor),0)"
  End If
  
  Set oWorksheet = oWorkbook.Sheets(1)
  With oWorksheet
    .Name = "SourceData"
    lngHoursCol = .Rows(1).Find("HOURS", lookat:=1).Column '1=xlWhole
    If Not blnFiscal Then
      lngWeekCol = oWorksheet.Rows(1).Find("WEEK", lookat:=1).Column '1=xlWhole
    End If
    lngLastCol = .[A2].End(xlToRight).Column
    'number formats
    For lngCol = 1 To lngLastCol
      If cptRxTest(.Columns(lngCol), "(HOURS|COST|FTE)") Then
        .Columns(lngCol).NumberFormat = "_(* #,##0.00000000_);_(* (#,##0.00000000);_(* ""-""??_);_(@_)"
      End If
    Next lngCol
    'add note on CostRateTable column
    If blnIncludeCosts Then
      lngCol = .Rows(1).Find("RATE_TABLE", lookat:=1).Column
      .Cells(1, lngCol).AddComment "Rate Table Applied in the Project"
    End If
    'add fte for non-fiscal
    If Not blnFiscal Then
      'create FTE_WEEK column
      Set oRange = .[A1].End(xlToRight).End(xlDown).Offset(0, 1)
      Set oRange = .Range(oRange, .[A1].End(xlToRight).Offset(1, 1))
      If blnFiscal Then 'fiscal
        'get fiscal_month column
        lngFiscalMonthCol = .Rows(1).Find(what:="FISCAL_MONTH", lookat:=xlWhole).Column
        oRange.FormulaR1C1 = "=RC" & lngHoursCol & "/NETWORKDAYS(RC" & lngWeekCol & "-7,RC" & lngWeekCol & ",EXCEPTIONS)"
      Else
        oRange.FormulaR1C1 = "=RC" & lngHoursCol & "/40"
      End If
      .[A1].End(xlToRight).Offset(0, 1).Value = "FTE_WEEK"
    End If
    'create FTE/FTE_MONTH column
    Set oRange = .[A1].End(xlToRight).Offset(1, 1)
    Set oRange = .Range(oRange, .Cells(.UsedRange.Rows.Count, oRange.Column))
    lngHoursCol = .Rows(1).Find("HOURS", lookat:=1).Column '1=xlWhole
    If blnFiscal Then
      lngFiscalMonthCol = .Rows(1).Find("FISCAL_MONTH", lookat:=1).Column '1=xlWhole
      oRange.FormulaR1C1 = "=RC" & lngHoursCol & "/LOOKUP(RC" & lngFiscalMonthCol & ",FISCAL[label],FISCAL[hpm])"
      .[A1].End(xlToRight).Offset(0, 1).Value = "FTE"
    Else
      lngWeekCol = .Rows(1).Find("WEEK", lookat:=1).Column
      oRange.FormulaR1C1 = "=RC" & lngHoursCol & "/160" 'todo: can we do something smarter?
      .[A1].End(xlToRight).Offset(0, 1).Value = "FTE_MONTH"
    End If
    'capture the range of data to feed as variable to PivotTable
    Set oRange = .Range(.[A1].End(xlDown), .[A1].End(xlToRight))
    strRange = .Name & "!" & Replace(oRange.Address, "$", "")
  End With 'SourceData Worksheet
  
  'add a new Worksheet for the oPivotTable
  Set oWorksheet = oWorkbook.Sheets.Add(Before:=oWorkbook.Sheets("SourceData"))
  'rename the new Worksheet
  oWorksheet.Name = "ResourceDemand"

  Application.StatusBar = "Creating PivotTable..."
  myResourceDemand_frm.lblStatus.Caption = Application.StatusBar

  'create the PivotTable
  oWorkbook.PivotCaches.Create(SourceType:=1, _
        SourceData:=strRange, Version:= _
        3).CreatePivotTable TableDestination:="ResourceDemand!R3C1", TableName:="RESOURCE_DEMAND", DefaultVersion:=3
  Set oPivotTable = oWorksheet.PivotTables(1)
  If blnFiscal Then
    oPivotTable.AddFields Array("RESOURCE_NAME", "[UID] TASK"), Array("FISCAL_MONTH")
    oPivotTable.AddDataField oPivotTable.PivotFields("FTE"), "FTE ", -4157
  Else
    If ActiveProject.Subprojects.Count > 0 Then
      oPivotTable.AddFields Array("RESOURCE_NAME", "PROJECT", "[UID] TASK"), Array("WEEK")
    Else
      oPivotTable.AddFields Array("RESOURCE_NAME", "[UID] TASK"), Array("WEEK")
    End If
    oPivotTable.AddDataField oPivotTable.PivotFields("FTE_WEEK"), "FTE_WEEK ", -4157
  End If
  
  'set default to ETC
  If blnExportAssociatedBaseline Or blnExportFullBaseline Then
    With oPivotTable
      With .PivotFields("CLASS")
        .Orientation = xlPageField
        .Position = 1
        .ClearAllFilters
        .CurrentPage = "ETC"
      End With
      If lngRateSets > 0 Then
        With .PivotFields("ACTIVE")
          .Orientation = xlPageField
          .Position = 1
          .ClearAllFilters
          .CurrentPage = "TRUE"
        End With
      End If
    End With
  End If
  
  'format the oPivotTable
  With oPivotTable
    .ShowDrillIndicators = True
    .EnableDrilldown = True
    .PivotCache.MissingItemsLimit = xlMissingItemsNone
    .PivotFields("RESOURCE_NAME").ShowDetail = False
    .TableStyle2 = "PivotStyleLight16"
    .PivotSelect "", 2, True
  End With
  oExcel.Selection.Style = "Comma"
  With oExcel.Selection
    With .FormatConditions
      .Delete
      .AddColorScale ColorScaleType:=2
    End With
    With .FormatConditions(1)
      .SetFirstPriority
      With .ColorScaleCriteria(1)
        .Type = 1 '1=xlConditionValueLowestValue
        .FormatColor.Color = 10285055
        .FormatColor.TintAndShade = 0
      End With
      With .ColorScaleCriteria(2)
        .Type = 2 '2=xlConditionValueHighestValue
        .FormatColor.Color = 2650623
        .FormatColor.TintAndShade = 0
      End With
      .ScopeType = 1 '1=xlFieldsScope
    End With
  End With
  
  Application.StatusBar = "Building header..."
  myResourceDemand_frm.lblStatus = Application.StatusBar

  'add a title
  With oWorksheet
    .Rows("1:3").EntireRow.Insert
    .[A2] = "Status Date: " & FormatDateTime(ActiveProject.StatusDate, vbShortDate)
    .[A2].EntireColumn.AutoFit
    With .[A1]
      .Value = "REMAINING WORK IN IMS: " & cptRegEx(ActiveProject.Name, "[^\\/]{1,}$")
      With .Font
        .Bold = True
        .Italic = True
        .Size = 14
      End With
    End With
    .[A1:F1].Merge
    'revise according to user options
    If blnFiscal Then
      .[B2] = "FTE by Fiscal Month"
    Else
      .[B2] = "FTE by Weeks " & myResourceDemand_frm.cboWeeks.Value & " " & myResourceDemand_frm.cboWeekday.Value
    End If
    oPivotTable.DataBodyRange.Select
    oExcel.ActiveWindow.FreezePanes = True
    .[A2].Select
    'make it nice
    oExcel.ActiveWindow.Zoom = 85
  End With

  Application.StatusBar = "Creating PivotChart..."
  myResourceDemand_frm.lblStatus.Caption = Application.StatusBar

  'create a PivotChart
  Set oWorksheet = oWorkbook.Sheets("SourceData")
  With oWorksheet
    .Activate
    .[A2].Select
    .[A2].EntireColumn.AutoFit
    oExcel.ActiveWindow.Zoom = 85
    oExcel.ActiveWindow.FreezePanes = True
    .Cells.EntireColumn.AutoFit
  End With
  Set oWorksheet = oWorkbook.Sheets.Add
  With oWorksheet
    .Name = "PivotChart_Source"
    Set oPivotTable = oWorkbook.Worksheets("ResourceDemand").PivotTables("RESOURCE_DEMAND")
    oPivotTable.PivotCache.CreatePivotTable TableDestination:="PivotChart_Source!R1C1", TableName:="PivotTable1", DefaultVersion:=3
    .Activate
    .[A1].Select
    Set oChart = .Shapes.AddChart2.Chart
    Set oRange = .Range(.[A1].End(-4161), .[A1].End(-4121))
  End With
  oChart.SetSourceData Source:=oRange
  oWorkbook.ShowPivotChartActiveFields = True
  oChart.ChartType = 76 'xlAreaStacked
  Set oPivotChartTable = oChart.PivotLayout.PivotTable
  If blnFiscal Then
    With oPivotChartTable.PivotFields("FISCAL_MONTH")
      .Orientation = 1 'xlRowField
      .Position = 1
    End With
  Else
    With oPivotChartTable.PivotFields("WEEK")
      .Orientation = 1 'xlRowField
      .Position = 1
    End With
  End If
  oPivotChartTable.AddDataField oPivotChartTable.PivotFields("HOURS"), "Sum of HOURS", -4157
  With oPivotChartTable.PivotFields("RESOURCE_NAME")
    .Orientation = 2 'xlColumnField
    .Position = 1
  End With
  If blnExportAssociatedBaseline Or blnExportFullBaseline Then
    'set default to ETC
    With oPivotChartTable.PivotFields("CLASS")
      .Orientation = xlPageField
      .Position = 1
      .ClearAllFilters
      .CurrentPage = "ETC"
    End With
  Else
    If Not blnFiscal Then
      oPivotTable.PivotFields("WEEK").PivotFilters.Add Type:=33, Value1:=ActiveProject.StatusDate '33 = xlAfter
    End If
  End If
  With oChart
    .ClearToMatchStyle
    .ChartStyle = 34
    .ClearToMatchStyle
    .SetElement (msoElementChartTitleAboveChart)
    .ChartTitle.Text = "Resource Demand"
    .Location 1, "PivotChart" 'xlLocationAsNewSheet = 1
  End With
  Set oWorksheet = oWorkbook.Sheets("PivotChart_Source")
  oWorksheet.Visible = False

  'add legend
  oExcel.ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
  oExcel.ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "HOURS"
  
  'export selected cost rate tables to oWorksheet
  If blnIncludeCosts Then
    Application.StatusBar = "Exporting Cost Rate Tables..."
    myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
    Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets("SourceData"))
    oWorksheet.Name = "Cost Rate Tables"
    oWorksheet.[A1:I1].Value = Array("PROJECT", "RESOURCE_NAME", "RESOURCE_TYPE", "ENTERPRISE", "RATE_TABLE", "EFFECTIVE_DATE", "STANDARD_RATE", "OVERTIME_RATE", "PER_USE_COST")
    lngRow = 2
    'make compatible with master/sub projects
    If ActiveProject.ResourceCount > 0 Then
      For Each oResource In ActiveProject.Resources
        With oWorksheet
          .Cells(lngRow, 1) = oResource.Name
          For Each oCostRateTable In oResource.CostRateTables
            If myResourceDemand_frm.Controls(Choose(oCostRateTable.Index, "chkA", "chkB", "chkC", "chkD", "chkE")).Value = True Then
              For Each oPayRate In oCostRateTable.PayRates
                .Cells(lngRow, 1) = cptRegEx(ActiveProject.Name, "[^\\/]{1,}$")
                .Cells(lngRow, 2) = oResource.Name
                .Cells(lngRow, 3) = Choose(oResource.Type + 1, "Work", "Material", "Cost")
                .Cells(lngRow, 4) = oResource.Enterprise
                .Cells(lngRow, 5) = oCostRateTable.Name
                .Cells(lngRow, 6) = FormatDateTime(oPayRate.EffectiveDate, vbShortDate)
                .Cells(lngRow, 7) = Replace(oPayRate.StandardRate, "/h", "")
                .Cells(lngRow, 8) = Replace(oPayRate.OvertimeRate, "/h", "")
                .Cells(lngRow, 9) = oPayRate.CostPerUse
                lngRow = .Cells(.Rows.Count, 1).End(-4162).Row + 1
              Next oPayRate
            End If
          Next oCostRateTable
        End With
      Next oResource
    ElseIf ActiveProject.Subprojects.Count > 0 Then
      For Each oSubproject In ActiveProject.Subprojects
        For Each oResource In oSubproject.SourceProject.Resources
          With oWorksheet
            .Cells(lngRow, 1) = oResource.Name
            For Each oCostRateTable In oResource.CostRateTables
              If myResourceDemand_frm.Controls(Choose(oCostRateTable.Index, "chkA", "chkB", "chkC", "chkD", "chkE")).Value = True Then
                For Each oPayRate In oCostRateTable.PayRates
                  .Cells(lngRow, 1) = cptRegEx(oSubproject.SourceProject.Name, "[^\\/]{1,}$")
                  .Cells(lngRow, 2) = oResource.Name
                  .Cells(lngRow, 3) = Choose(oResource.Type + 1, "Work", "Material", "Cost")
                  .Cells(lngRow, 4) = oResource.Enterprise
                  .Cells(lngRow, 5) = oCostRateTable.Name
                  .Cells(lngRow, 6) = FormatDateTime(oPayRate.EffectiveDate, vbShortDate)
                  .Cells(lngRow, 7) = Replace(oPayRate.StandardRate, "/h", "")
                  .Cells(lngRow, 8) = Replace(oPayRate.OvertimeRate, "/h", "")
                  .Cells(lngRow, 9) = oPayRate.CostPerUse
                  lngRow = .Cells(.Rows.Count, 1).End(-4162).Row + 1
                Next oPayRate
              End If
            Next oCostRateTable
          End With
        Next oResource
      Next oSubproject
    End If
  
    'make it a oListObject
    Set oListObject = oWorksheet.ListObjects.Add(1, oWorksheet.Range(oWorksheet.[A1].End(-4161), oWorksheet.[A1].End(-4121)), , 1)
    oListObject.Name = "CostRateTables"
    oListObject.TableStyle = ""
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.[A2].Select
    oExcel.ActiveWindow.FreezePanes = True
    oWorksheet.Columns.AutoFit
    
  End If
    
  'PivotTable Worksheet active by default
  oWorkbook.Sheets("ResourceDemand").Activate
  
  'provide user feedback
  Application.StatusBar = "Saving the Workbook..."
  myResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  
'  'save the file
'  '<issue49> - file exists in location
'  strFileName = oShell.SpecialFolders("Desktop") & "\" & Replace(oWorkbook.Name, ".xlsx", "_" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".xlsx") '<issue49>
'  If Dir(strFileName) <> vbNullString Then '<issue49>
'    If MsgBox("A file named '" & strFileName & "' already exists in this location. Replace?", vbYesNo + vbExclamation, "Overwrite?") = vbYes Then '<issue49>
'      Kill strFileName '<issue49>
'      oWorkbook.SaveAs strFileName, 51 '<issue49>
'      MsgBox "Saved to your Desktop:" & vbCrLf & vbCrLf & Dir(strFileName), vbInformation + vbOKOnly, "Resource Demand Exported" '<issue49>
'    End If '<issue49>
'  Else '<issue49>
'    oWorkbook.SaveAs strFileName, 51  '<issue49>
'  End If '</issue49>
  
  If blnFiscal Then
    strMsg = "Apply an efficiency factor in cell H1 of the FiscalPeriods worksheet (e.g., 1 FTE = 85%)." & vbCrLf & vbCrLf
    strMsg = strMsg & "To account for calendar exceptions:" & vbCrLf
    strMsg = strMsg & "- use Calendar Details feature to export calendar exceptions;" & vbCrLf
    strMsg = strMsg & "- for recurring exceptions, be sure to select 'detailed';" & vbCrLf
    strMsg = strMsg & "- expand recurring exceptions to show full list of Start dates;" & vbCrLf
    strMsg = strMsg & "- copy list of 'Start' dates and paste into Exceptions List on FiscalPeriods sheet;" & vbCrLf
    strMsg = strMsg & "- activate ResourceDemand or PivotChart sheet and Refresh Pivot data" & vbCrLf & vbCrLf
    strMsg = strMsg & "(Take a screen shot of these instructions, if needed.)"
    MsgBox strMsg, vbInformation + vbOKOnly, "Next Actions:"
    oWorkbook.Sheets("FiscalPeriods").Activate
    oWorkbook.Sheets("FiscalPeriods").[E2].Select
  End If
  
  MsgBox "Export Complete", vbInformation + vbOKOnly, "Staffing Profile"
  
  Application.StatusBar = "Complete."
  myResourceDemand_frm.lblStatus.Caption = Application.StatusBar

  oExcel.Visible = True
  Application.ActivateMicrosoftApp pjMicrosoftExcel
  
exit_here:
  On Error Resume Next
  If Not oExcel Is Nothing Then oExcel.Visible = True
  Application.StatusBar = ""
  myResourceDemand_frm.lblStatus.Caption = "Ready..."
  cptSpeed False
  Set oAssignment = Nothing
  Set oCalendar = Nothing
  Set oChart = Nothing
  Set oCostRateTable = Nothing
  Set oEstimates = Nothing
  Set oExcel = Nothing
  Set oException = Nothing
  Set oListObject = Nothing
  Set oPayRate = Nothing
  Set oPivotChartTable = Nothing
  Set oPivotTable = Nothing
  Set oRange = Nothing
  Set oRecordset = Nothing
  Set oResource = Nothing
  Set oSettings = Nothing
  Set oSubproject = Nothing
  Set oTask = Nothing
  Set oTSV = Nothing
  Set oTSVS_AC = Nothing
  Set oTSVS_AW = Nothing
  Set oTSVS_BCWS = Nothing
  Set oTSVS_COST = Nothing
  Set oTSVS_WORK = Nothing
  Set oWorkbook = Nothing
  Set oWorksheet = Nothing

  If Not oWorkbook Is Nothing Then oWorkbook.Close False
  If Not oExcel Is Nothing Then oExcel.Quit
  Exit Sub
err_here:
  Call cptHandleErr(MODULE_NAME, "cptExportResourceDemand", Err, Erl)
  On Error Resume Next
  Resume exit_here

End Sub

Sub cptShowExportResourceDemand_frm()
  'objects
  Dim myResourceDemand_frm As cptResourceDemand_frm
  Dim rst As ADODB.Recordset
  Dim rstResources As Object 'ADODB.Recordset
  Dim objProject As Object
  Dim rstFields As Object 'ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strNonLabor As String
  Dim strBaseline As String
  Dim strCostSets As String
  Dim strCosts As String
  Dim strWeeks As String
  Dim strMonths As String
  Dim strWeekday As String
  Dim strMissing As String
  Dim strFieldName As String
  Dim strFileName As String
  'longs
  Dim lngFile As Long
  Dim lngResourceCount As Long
  Dim lngResource As Long
  Dim lngField As Long
  Dim lngItem As Long
  'integers
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnFiscalCalendarExists As Boolean
  'variants
  Dim vField As Variant
  Dim vCostSet As Variant
  Dim vCostSets As Variant
  Dim vFieldType As Variant
  'dates

  'prevent spawning
  If Not cptGetUserForm("cptResourceDemand_frm") Is Nothing Then Exit Sub
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  
  'requires ms excel
  If Not cptCheckReference("Excel") Then
    MsgBox "This feature requires MS Excel.", vbCritical + vbOKOnly, "Resource Demand"
    GoTo exit_here
  End If
  If ActiveProject.Subprojects.Count = 0 And ActiveProject.ResourceCount = 0 Then
    MsgBox "This project has no resources to export.", vbExclamation + vbOKOnly, "No Resources"
    GoTo exit_here
  Else
    cptSpeed True
    lngResourceCount = ActiveProject.ResourceCount
    Set rstResources = CreateObject("ADODB.Recordset")
    rstResources.Fields.Append "RESOURCE_NAME", adVarChar, 200
    rstResources.Open
    For lngItem = 1 To ActiveProject.Subprojects.Count
      Set objProject = ActiveProject.Subprojects(lngItem).SourceProject
      Application.StatusBar = "Loading " & objProject.Name & "..."
      For lngResource = 1 To objProject.Resources.Count
        With rstResources
          .Filter = "[RESOURCE_NAME]='" & objProject.Resources(lngResource).Name & "'"
          If rstResources.RecordCount = 0 Then
            .AddNew Array(0), Array("'" & objProject.Resources(lngResource).Name & "'")
          Else
            Debug.Print "duplicate found"
          End If
          .Filter = ""
        End With
      Next lngResource
      Set objProject = Nothing
    Next lngItem
    rstResources.Close 'todo: save for later?
    Application.StatusBar = ""
    cptSpeed False
  End If
  
  'instantiate the form
  Set myResourceDemand_frm = New cptResourceDemand_frm
  myResourceDemand_frm.lboFields.Clear
  myResourceDemand_frm.lboExport.Clear

  Set rstFields = CreateObject("ADODB.Recordset")
  rstFields.Fields.Append "CONSTANT", adInteger
  rstFields.Fields.Append "CUSTOM_NAME", adVarChar, 255
  rstFields.Open
  
  'add the 'Critical' field
  rstFields.AddNew Array(0, 1), Array(FieldNameToFieldConstant("Critical"), "Critical")
    
  For Each vFieldType In Array("Text", "Outline Code")
    On Error GoTo err_here
    For lngItem = 1 To 30
      lngField = FieldNameToFieldConstant(vFieldType & lngItem) ',lngFieldType)
      strFieldName = CustomFieldGetName(lngField)
      If Len(strFieldName) > 0 Then
        'todo: handle duplicates if master/subprojects
        rstFields.AddNew Array(0, 1), Array(lngField, strFieldName)
        rstFields.Update
      End If
next_field:
    Next lngItem
  Next vFieldType

  'get enterprise custom fields
  For lngField = 188776000 To 188778000
    If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      strFieldName = Application.FieldConstantToFieldName(lngField)
      'todo: avoid conflicts between local and custom fields?
      'If rstFields.Contains(strFieldName) Then
      '  MsgBox "An Enterprise Field named '" & strFieldName & "' conflicts with a local custom field of the same name. The local field will be ignored.", vbExclamation + vbOKOnly, "Conflict"
        'rstFields.Remove Application.FieldConstantToFieldName(lngField)
      'End If
      rstFields.AddNew Array(0, 1), Array(lngField, strFieldName)
      rstFields.Update
    End If
next_field1:
  Next lngField

  'add fields to listbox
  rstFields.Sort = "CUSTOM_NAME"
  rstFields.MoveFirst
  lngItem = 0
  Do While Not rstFields.EOF
    myResourceDemand_frm.lboFields.AddItem
    myResourceDemand_frm.lboFields.List(lngItem, 0) = rstFields(0)
    If rstFields(0) > 188776000 Then
      myResourceDemand_frm.lboFields.List(lngItem, 1) = rstFields(1) & " (Enterprise)"
    Else
      myResourceDemand_frm.lboFields.List(lngItem, 1) = rstFields(1) & " (" & FieldConstantToFieldName(rstFields(0)) & ")"
    End If
    rstFields.MoveNext
    lngItem = lngItem + 1
  Loop

  'save the fields to a file for fast searching
  If rstFields.RecordCount > 0 Then
    strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
    If Dir(strFileName) <> vbNullString Then Kill strFileName
    rstFields.Save strFileName, adPersistADTG
  End If
  rstFields.Close
  
  'populate options and set defaults
  With myResourceDemand_frm
    .cboWeeks.AddItem "Beginning"
    .cboWeeks.AddItem "Ending"
    'allow to trigger, it populates the form
    .cboWeeks.Value = "Beginning"
    .cboWeekday = "Monday"
    .chkA.Value = False
    .chkB.Value = False
    .chkC.Value = False
    .chkD.Value = False
    .chkE.Value = False
    .chkCosts.Value = False
    .chkAssociatedBaseline = False
    .chkFullBaseline = False
    .cboMonths.Clear
    .cboMonths.AddItem
    .cboMonths.List(.cboMonths.ListCount - 1, 0) = 0
    .cboMonths.List(.cboMonths.ListCount - 1, 1) = "Calendar (Default Excel Grouping)"
    .cboMonths.Value = 0
    blnFiscalCalendarExists = cptCalendarExists("cptFiscalCalendar")
    If blnFiscalCalendarExists Then
      .cboMonths.AddItem
      .cboMonths.List(.cboMonths.ListCount - 1, 0) = 1
      .cboMonths.List(.cboMonths.ListCount - 1, 1) = "Fiscal (cptFiscalCalendar)"
    Else
      .cboMonths.Enabled = False
      .cboMonths.Locked = True
    End If
  End With
  
  'import saved fields if exists
  strFileName = strDir & "\settings\cpt-export-resource-userfields.adtg"
  If Dir(strFileName) <> vbNullString Then
    Set rst = CreateObject("ADODB.Recordset")
    With rst
      .Open strFileName, , adOpenKeyset, adLockReadOnly
      .MoveFirst
      lngItem = 0
      Do While Not .EOF
        If .Fields(0) = "settings" Then
          'don't use it - obsolete
        Else
          If .Fields(0) >= 188776000 Then 'check enterprise field
            If FieldConstantToFieldName(.Fields(0)) <> Replace(.Fields(1), cptRegEx(.Fields(1), " \([A-z0-9]*\)$"), "") Then
              strMissing = strMissing & "- " & .Fields(1) & vbCrLf
              GoTo next_saved_field
            End If
          Else 'check local field
            If CustomFieldGetName(.Fields(0)) <> Trim(Replace(.Fields(1), cptRegEx(.Fields(1), "\([^\(].*\)$"), "")) Then
              'limit this check to Custom Fields
              If IsNumeric(Right(FieldConstantToFieldName(.Fields(0)), 1)) Then
                strMissing = strMissing & "- " & .Fields(1) & vbCrLf
                GoTo next_saved_field
              End If
            End If
          End If
          myResourceDemand_frm.lboExport.AddItem
          myResourceDemand_frm.lboExport.List(lngItem, 0) = .Fields(0) 'Field Constant
          myResourceDemand_frm.lboExport.List(lngItem, 1) = .Fields(1) 'Custom Field Name
          lngItem = lngItem + 1
        End If
next_saved_field:
        .MoveNext
      Loop
      .Close
    End With
  End If
  
  'import saved settings
  With myResourceDemand_frm
    If Dir(strDir & "\settings\cpt-settings.ini") <> vbNullString Then
      cptDeleteSetting "ResourceDemand", "chkBaseline"
      cptDeleteSetting "ResourceDemand", "lboExport"
      'month
      strMonths = cptGetSetting("ResourceDemand", "cboMonths")
      If Len(strMonths) > 0 Then
        If CLng(strMonths) = 1 And blnFiscalCalendarExists Then
          .cboMonths.Value = CLng(strMonths)
        Else
          .cboMonths.Value = 0
        End If
      End If
      'week
      strWeeks = cptGetSetting("ResourceDemand", "cboWeeks")
      If Len(strWeeks) > 0 Then
        .cboWeeks.Value = strWeeks
      End If
      'weekday
      strWeekday = cptGetSetting("ResourceDemand", "cboWeekday")
      If Len(strWeekday) > 0 Then
        .cboWeekday.Value = strWeekday
      End If
      'costs
      strCosts = cptGetSetting("ResourceDemand", "chkCosts")
      If Len(strCosts) > 0 Then
        .chkCosts = CBool(strCosts)
      End If
      If .chkCosts Then
        strCostSets = cptGetSetting("ResourceDemand", "CostSets")
        If Len(strCostSets) > 0 Then
          If Right(strCostSets, 1) = "," Then strCostSets = Left(strCostSets, Len(strCostSets) - 1)
          For Each vCostSet In Split(strCostSets, ",")
            .Controls("chk" & Choose(CLng(vCostSet + 1), "A", "B", "C", "D", "E")).Value = True
          Next vCostSet
        End If
      Else
        For Each vCostSet In Split("A,B,C,D,E", ",")
          .Controls("chk" & vCostSet).Value = False
          .Controls("chk" & vCostSet).Enabled = False
        Next vCostSet
      End If
      'baseline
      strBaseline = cptGetSetting("ResourceDemand", "chkAssociatedBaseline")
      If Len(strBaseline) > 0 Then
        .chkAssociatedBaseline = CBool(strBaseline)
      End If
      strBaseline = cptGetSetting("ResourceDemand", "chkFullBaseline")
      If Len(strBaseline) > 0 Then
        .chkFullBaseline = CBool(strBaseline)
      End If
      'non-labor
      cptDeleteSetting "ResourceDemand", "chkNonLabor" 'obsolete setting
    End If
    .Caption = "Export Resource Demand (" & cptGetVersion(MODULE_NAME) & ")"
    If Len(strMissing) > 0 Then
      If UBound(Split(strMissing, vbCrLf)) > 10 Then
        lngFile = FreeFile
        strFileName = Environ("tmp") & "\cpt-resourcedemand-missing-fields.txt"
        Open strFileName For Output As #lngFile
        Print #lngFile, "The following saved fields do not exist in this project:"
        Print #lngFile, strMissing
        Close #lngFile
        ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
        MsgBox "There are " & UBound(Split(strMissing, vbCrLf)) & " saved fields that do not exist in this project.", vbCritical + vbOKOnly, "Saved Settings"
      Else
        MsgBox "The following saved fields do not exist in this project:" & vbCrLf & strMissing, vbInformation + vbOKOnly, "Saved Settings"
      End If
    End If
    .Show 'False
  End With

exit_here:
  On Error Resume Next
  Unload myResourceDemand_frm
  Set myResourceDemand_frm = Nothing
  Set rst = Nothing
  If rstResources.State Then rstResources.Close
  Set rstResources = Nothing
  Set objProject = Nothing
  If rstFields.State Then rstFields.Close
  Set rstFields = Nothing
  Exit Sub

err_here:
  If Err.Number = 1101 Or Err.Number = 1004 Then
    Err.Clear
    Resume next_field
  Else
    Call cptHandleErr(MODULE_NAME, "cptShowExportResourceDemand_frm", Err, Erl)
    Resume exit_here
  End If

End Sub

Function cptGetFiscalMonthOfDay(dtDate As Date, vFiscal As Variant)
  Dim lngItem As Long
  For lngItem = 0 To UBound(vFiscal, 2)
    If vFiscal(0, lngItem) >= dtDate Then
      cptGetFiscalMonthOfDay = vFiscal(1, lngItem)
      Exit Function
    End If
  Next lngItem
  cptGetFiscalMonthOfDay = ""
End Function
