Attribute VB_Name = "cptMetrics_bas"
'<cpt_version>v1.0.7</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

'add disclaimer: unburdened hours - not meant to be precise - generally within +/- 1%

Sub cptExportMetricsExcel()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  MsgBox "Stay tuned...", vbInformation + vbOKOnly, "Under Construction..."

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptExportMetricsExcel", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetBAC()
  MsgBox Format(cptGetMetric("bac"), "#,##0.00h"), vbInformation + vbOKOnly, "Budget at Complete (BAC) - hours"
End Sub

Sub cptGetETC()
  MsgBox Format(cptGetMetric("etc"), "#,##0.00h"), vbInformation + vbOKOnly, "Estimate to Complete (ETC) - hours"
End Sub

Sub cptGetBCWS()

  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    MsgBox Format(cptGetMetric("bcws"), "#,##0.00"), vbInformation + vbOKOnly, "Budgeted Cost of Work Scheduled (BCWS) - hours"
  End If

End Sub

Sub cptGetBCWP()
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    MsgBox Format(cptGetMetric("bcwp"), "#,##0.00"), vbInformation + vbOKOnly, "Budgeted Cost of Work Performed (BCWP) - hours"
  End If
  
End Sub

Sub cptGetSPI()
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("SPI")
  End If
  
End Sub

Sub cptGetBEI()

  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("BEI")
  End If
  
End Sub

Sub cptGetCEI()
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("CEI")
  End If
  
End Sub

Sub cptGetSV()
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("SV")
  End If

End Sub

Sub cptGetCPLI()
'objects
Dim oTasks As Tasks
Dim oPred As Task
Dim oTask As Task
'strings
Dim strMsg As String
Dim strTitle As String
'longs
Dim lngConstraintType As Long
Dim lngTS As Long
Dim lngMargin As Long
Dim lngCPL As Long
'integers
'doubles
'booleans
'variants
'dates
Dim dtStart As Date, dtFinish As Date
Dim dtConstraintDate As Date

  strTitle = "Critical Path Length Index (CPLI)"

  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then
    MsgBox "No Target Task selected.", vbExclamation + vbOKOnly, "Oops"
    GoTo exit_here
  End If

  'confirm a single, target oTask is selected
  If oTasks.Count <> 1 Then
    MsgBox "Please select a single, active, and non-summary target oTask.", vbExclamation + vbOKOnly, strTitle
    GoTo exit_here
  End If
  
  Set oTask = oTasks(1)
  
  'use MFO or MSO constraint
  If oTask.ConstraintType <> pjMFO And oTask.ConstraintType <> pjMSO Then
    strMsg = "No MSO/MFO constraint found; temporarily using Deadline..." & vbCrLf
    'if no MFO then use deadline as MFO
    If IsDate(oTask.Deadline) Then
      If IsDate(oTask.ConstraintDate) Then dtConstraintDate = oTask.ConstraintDate
      lngConstraintType = oTask.ConstraintType
      oTask.ConstraintDate = oTask.Deadline
      oTask.ConstraintType = pjMFO
      lngTS = oTask.TotalSlack
      dtFinish = oTask.Finish
      If CLng(dtConstraintDate) > 0 Then oTask.ConstraintDate = dtConstraintDate
      oTask.ConstraintType = lngConstraintType
    Else
      strMsg = strMsg & "No Deadline found; temporarily using Baseline Finish..." & vbCrLf
      If Not IsDate(oTask.BaselineFinish) Then
        strMsg = strMsg & "No Baseline Finish found." & vbCrLf & vbCrLf
        strMsg = strMsg & "In order to calculate the CPLI, the target Task should be (at least temporarily) constrained with a MFO or Deadline." & vbCrLf & vbCrLf
        strMsg = strMsg & "Please constrain the Target Task and try again."
        MsgBox strMsg, vbExclamation + vbOKOnly, strTitle
        GoTo exit_here
      Else
        If IsDate(oTask.ConstraintDate) Then dtConstraintDate = oTask.ConstraintDate
        lngConstraintType = oTask.ConstraintType
        oTask.ConstraintDate = oTask.BaselineFinish
        oTask.ConstraintType = pjMFO
        lngTS = oTask.TotalSlack
        dtFinish = oTask.Finish
        If CLng(dtConstraintDate) > 0 Then oTask.ConstraintDate = dtConstraintDate
        oTask.ConstraintType = lngConstraintType
      End If
    End If
  Else
    lngTS = oTask.TotalSlack
    dtFinish = oTask.Finish
  End If
      
  'use status date if exists
  If IsDate(ActiveProject.StatusDate) Then
    dtStart = ActiveProject.StatusDate
  Else
    dtStart = FormatDateTime(Now(), vbShortDate) & " 08:00 AM"
  End If
  
  'use earliest start date
  'NOTE: cannot account for schedule margin due to possibility
  'of dual paths, one with and one without, a particular SM Task
  
  If oTask Is Nothing Then GoTo exit_here
  If oTask.Summary Then GoTo exit_here
  If Not oTask.Active Then GoTo exit_here
  HighlightDrivingPredecessors Set:=True
  For Each oPred In ActiveProject.Tasks
    If oPred.PathDrivingPredecessor Then
      If IsDate(oPred.ActualStart) Then
        If oPred.Stop < dtStart Then dtStart = oPred.Stop
      Else
        If oPred.Start < dtStart Then dtStart = oPred.Start
      End If
    End If
  Next oPred
  'calculate the CPL
  lngCPL = Application.DateDifference(dtStart, dtFinish)
  'convert values to days
  lngCPL = lngCPL / 480
  lngTS = lngTS / 480
  'notify user
  strMsg = strMsg & vbCrLf & "CPL = Critical Path Length" & vbCrLf
'  strMsg = strMsg & "CPL = Target Finish - Timenow (or CP start)" & vbCrLf
'  strMsg = strMsg & "CPL = " & FormatDateTime(dtFinish, vbShortDate) & " - " & FormatDateTime(dtStart, vbShortDate) & vbCrLf
'  strMsg = strMsg & "CPL = " & lngCPL & vbCrLf
  strMsg = strMsg & "TS = Total Slack" & vbCrLf & vbCrLf
  strMsg = strMsg & "CPLI = ( CPL + TS ) / CPL" & vbCrLf
  strMsg = strMsg & "CPLI = ( " & lngCPL & " + " & lngTS & " ) / " & lngCPL & vbCrLf & vbCrLf
  strMsg = strMsg & "CPLI = " & Round((lngCPL + lngTS) / lngCPL, 3) & vbCrLf & vbCrLf
  strMsg = strMsg & "Note: Schedule Margin Tasks are not considered."
  
  MsgBox strMsg, vbInformation + vbOKOnly, "Critical Path Length Index (CPLI)"
    
exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  Set oPred = Nothing
  Application.CloseUndoTransaction
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetCPLI", Err, Erl)
  Resume exit_here
End Sub

Sub cptGET(strWhat As String)
'todo: need to store weekly bcwp, etc data somewhere
'objects
Dim oRecordset As ADODB.Recordset
'strings
Dim strMsg As String, strProgram As String
'longs
Dim lngBEI_AF As Long
Dim lngBEI_BF As Long
'integers
'doubles
Dim dblBCWS As Double
Dim dblBCWP As Double
Dim dblResult As Double
'booleans
'variants
'dates
Dim dtStatus As Date, dtPrevious As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'validate tasks exist
  If ActiveProject.Tasks.Count = 0 Then
    MsgBox "This project has no tasks.", vbExclamation + vbOKOnly, "No Tasks"
    GoTo exit_here
  End If
  
  strProgram = cptGetProgramAcronym
  If Len(strProgram) = 0 Then GoTo exit_here
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project requires a Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
    GoTo exit_here
  Else
    dtStatus = ActiveProject.StatusDate
  End If
    
  Select Case strWhat
    Case "BEI"
      lngBEI_BF = CLng(cptGetMetric("bei_bf"))
      If lngBEI_BF = 0 Then
        MsgBox "No baseline finishes found.", vbExclamation + vbOKOnly, "No BEI"
        GoTo exit_here
      End If
      lngBEI_AF = CLng(cptGetMetric("bei_af"))
      strMsg = "BEI = # Actual Finishes / # Planned Finishes" & vbCrLf
      strMsg = strMsg & "BEI = " & Format(lngBEI_AF, "#,##0") & " / " & Format(lngBEI_BF, "#,##0") & vbCrLf & vbCrLf
      strMsg = strMsg & "BEI = " & Round(lngBEI_AF / lngBEI_BF, 2)
      cptCaptureMetric strProgram, dtStatus, "BEI", Round(lngBEI_AF / lngBEI_BF, 2)
      MsgBox strMsg, vbInformation + vbOKOnly, "Baseline Execution Index (BEI)"
      
    Case "CEI"
      'does cpt-cei.adtg exist?
      If Dir(cptDir & "\settings\cpt-cei.adtg") = vbNullString Then
        MsgBox "No data file found. You must 'Capture Week' on previous period's file before you can run CEI on current period's statused IMS.", vbExclamation + vbOKOnly, "File Not Found"
        GoTo exit_here
      End If
      'get program acronym
      strProgram = cptGetProgramAcronym
      If Len(strProgram) = 0 Then GoTo exit_here
      'connect to data source
      Set oRecordset = CreateObject("ADODB.Recordset")
      'get list of tasks & count
      oRecordset.Open cptDir & "\settings\cpt-cei.adtg"
      dtStatus = ActiveProject.StatusDate
      With oRecordset
        .MoveFirst
        'get most previous week_ending
        dtPrevious = .Fields("STATUS_DATE")
        Do While Not .EOF
          If .Fields("PROJECT") = strProgram Then
            If .Fields("STATUS_DATE") > dtPrevious And .Fields("STATUS_DATE") < dtStatus Then
              dtPrevious = .Fields("STATUS_DATE")
            End If
          End If
          .MoveNext
        Loop
        'test each one to see if complete and get count
        .MoveFirst
        Do While Not .EOF
          If CBool(.Fields("IS_LOE")) Then GoTo next_record
          If .Fields("PROJECT") = strProgram And .Fields("STATUS_DATE") = dtPrevious Then
            If .Fields("TASK_FINISH") > dtPrevious And .Fields("TASK_FINISH") <= dtStatus Then
              Dim lngFF As Long
              lngFF = lngFF + 1
              On Error Resume Next
              Dim oTask As Task
              If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
              Set oTask = ActiveProject.Tasks.UniqueID(.Fields(1))
              If Not oTask Is Nothing Then
                If IsDate(oTask.ActualFinish) Then
                  Dim lngAF As Long
                  lngAF = lngAF + 1
                End If
              End If
            End If
          End If
next_record:
          .MoveNext
        Loop
        'todo: notify user; prompt for list of FFs?
        strMsg = "CEI = Tasks completed in current period / Tasks forecasted to complete in current period" & vbCrLf & vbCrLf
        strMsg = strMsg & "CEI = " & lngAF & " / " & lngFF & vbCrLf
        strMsg = strMsg & "CEI = " & Round(lngAF / lngFF, 2) & vbCrLf & vbCrLf
        strMsg = strMsg & "- Does not include LOE tasks." & vbCrLf
        strMsg = strMsg & "- Does not include tasks completed in current period but not forecasted to complete in current period." & vbCrLf
        strMsg = strMsg & "- See NDIA Predictive Measures Guide for more information."
        Call cptCaptureMetric(strProgram, dtStatus, "CEI", Round(lngAF / lngFF, 2))
        MsgBox strMsg, vbInformation + vbOKOnly, "Current Execution Index"
        .Close
      End With
      
    Case "SPI"
      dblBCWP = cptGetMetric("bcwp")
      dblBCWS = cptGetMetric("bcws")
      If dblBCWS = 0 Then
        MsgBox "No BCWS found.", vbExclamation + vbOKOnly, "Schedule Performance Index (SPI) - Hours"
        GoTo exit_here
      End If
      strMsg = "SPI = BCWP / BCWS" & vbCrLf
      strMsg = strMsg & "SPI = " & Format(dblBCWP, "#,##0h") & " / " & Format(dblBCWS, "#,##0h") & vbCrLf & vbCrLf
      strMsg = strMsg & "SPI = ~" & Round(dblBCWP / dblBCWS, 2) '& vbCrLf & vbCrLf
      'strMsg = strMsg & "(Assumes EV% in Physical % Complete.)"
      cptCaptureMetric strProgram, dtStatus, "SPI", Round(dblBCWP / dblBCWS, 2)
      MsgBox strMsg, vbInformation + vbOKOnly, "Schedule Performance Index (SPI) - Hours"
      
    Case "SV"
      dblBCWP = cptGetMetric("bcwp")
      dblBCWS = cptGetMetric("bcws")
      If dblBCWS = 0 Then
        MsgBox "No BCWS found.", vbExclamation + vbOKOnly, "Schedule Variance (SV) - Hours"
        GoTo exit_here
      End If
      strMsg = strMsg & "Schedule Variance (SV)" & vbCrLf
      strMsg = strMsg & "SV = BCWP - BCWS" & vbCrLf
      strMsg = strMsg & "SV = " & Format(dblBCWP, "#,##0h") & " - " & Format(dblBCWS, "#,##0h") & vbCrLf
      strMsg = strMsg & "SV = ~" & Format(dblBCWP - dblBCWS, "#,##0.0h") & vbCrLf & vbCrLf
      strMsg = strMsg & "Schedule Variance % (SV%)" & vbCrLf
      strMsg = strMsg & "SV% = ( SV / BCWS ) * 100" & vbCrLf
      strMsg = strMsg & "SV% = ( " & Format((dblBCWP - dblBCWS), "#,##0.0h") & " / " & Format(dblBCWS, "#,##0.0h") & " ) * 100" & vbCrLf
      strMsg = strMsg & "SV% = " & Format(((dblBCWP - dblBCWS) / dblBCWS), "0.00%") '& vbCrLf & vbCrLf
      'strMsg = strMsg & "(Assumes EV% in Physical % Complete.)"
      cptCaptureMetric strProgram, dtStatus, "SV", Round((dblBCWP - dblBCWS) / dblBCWS, 2)
      MsgBox strMsg, vbInformation + vbOKOnly, "Schedule Variance (SV) - Hours"
      
    Case "es" 'earned schedule
      'todo: earned schedule
      'todo: update cptCaptureMetric for date
    
  End Select
  
  
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMerics_Bas", "cptGet", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetHitTask()
'objects
Dim oTask As Task
'strings
Dim strMsg As String
'longs
Dim lngAF As Long
Dim lngBLF As Long
'integers
'doubles
'booleans
'variants
'dates
Dim dtStatus As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    GoTo exit_here
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  'find it
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If IsDate(oTask.BaselineFinish) Then
      'was task baselined to finish before status date?
      If oTask.BaselineFinish <= dtStatus Then
        lngBLF = lngBLF + 1
        'did it?
        If IsDate(oTask.ActualFinish) Then
          If oTask.ActualFinish <= oTask.BaselineFinish Then
            lngAF = lngAF + 1
          End If
        End If
      End If
    End If
next_task:
  Next oTask

  strMsg = "BF = # Tasks Baselined to Finish ON or before Status Date" & vbCrLf
  strMsg = strMsg & "AF = # BF that Actually Finished ON or before Baseline Finish" & vbCrLf & vbCrLf
  strMsg = strMsg & "Hit Task % = (AF / BF) / 100" & vbCrLf
  strMsg = strMsg & "Hit Task % = (" & Format(lngAF, "#,##0") & " / " & Format(lngBLF, "#,##0") & ") / 100" & vbCrLf & vbCrLf
  strMsg = strMsg & "Hit Task % = " & Format((lngAF / lngBLF), "0%")
  MsgBox strMsg, vbInformation + vbOKOnly, "Hit Task %"

exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetHitTask", Err, Erl)
  Resume exit_here
End Sub

Function cptGetMetric(strGet As String) As Double
'todo: no screen changes!
'objects
Dim oAssignment As Assignment
Dim tsv As TimeScaleValue
Dim tsvs As TimeScaleValues
Dim oTasks As Tasks
Dim oTask As Task
'strings
Dim strLOE As String
'longs
Dim lngLOEField As Long
Dim lngEVP As Long
Dim lngYears As Long
'integers
'doubles
Dim dblResult As Double
'booleans
'variants
'dates
Dim dtStatus As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  lngYears = Year(ActiveProject.ProjectFinish) - Year(ActiveProject.ProjectStart) + 1
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    GoTo exit_here
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  cptSpeed True
  FilterClear
  GroupClear
  OptionsViewEx displaysummarytasks:=True, displaynameindent:=True
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    Sort "ID", , , , , , False, True
    OutlineShowAllTasks
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  SelectAll
  Set oTasks = ActiveSelection.Tasks
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.BaselineWork > 0 Then 'idea here was to limit tasks to PMB tasks only
                                  'but won't work for non-resource loaded schedules
      Select Case strGet
        Case "bac"
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              dblResult = dblResult + (oAssignment.BaselineWork / 60)
            End If
          Next oAssignment
          
        Case "etc"
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              dblResult = dblResult + (oAssignment.RemainingWork / 60)
            End If
          Next oAssignment

          
        Case "bcws"
          If oTask.BaselineStart < dtStatus Then
            For Each oAssignment In oTask.Assignments
              If oAssignment.ResourceType = pjResourceTypeWork Then
                Set tsvs = oAssignment.TimeScaleData(oTask.BaselineStart, dtStatus, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                For Each tsv In tsvs
                  dblResult = dblResult + (IIf(tsv.Value = "", 0, tsv.Value) / 60)
                Next
              End If
            Next oAssignment
          End If
          
        Case "bcwp"
          lngEVP = CLng(cptGetSetting("Metrics", "cboEVP"))
          lngLOEField = CLng(cptGetSetting("Metrics", "cboLOEField"))
          strLOE = cptGetSetting("Metrics", "txtLOE")
          
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              If oTask.GetField(lngLOEField) = strLOE Then
                If oTask.BaselineStart < dtStatus Then
                  Set tsvs = oAssignment.TimeScaleData(oTask.BaselineStart, dtStatus, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks, 1)
                  For Each tsv In tsvs
                    dblResult = dblResult + (IIf(tsv.Value = "", 0, tsv.Value) / 60)
                  Next
                End If
              Else
                dblResult = dblResult + ((oAssignment.BaselineWork / 60) * (CLng(cptRegEx(oTask.GetField(lngEVP), "[0-9]*")) / 100))
              End If
            End If
          Next oAssignment
      End Select
    End If 'bac>0
    Select Case strGet
    
      Case "bei_bf"
        dblResult = dblResult + IIf(oTask.BaselineFinish <= dtStatus, 1, 0)
          
      Case "bei_af"
        dblResult = dblResult + IIf(oTask.ActualFinish <= dtStatus, 1, 0)

    End Select
next_task:
    Application.StatusBar = "Calculating " & UCase(strGet) & "..."
  Next

  cptGetMetric = dblResult

exit_here:
  On Error Resume Next
  Set oAssignment = Nothing
  Application.StatusBar = ""
  cptSpeed False
  Set tsv = Nothing
  Set tsvs = Nothing
  Set oTasks = Nothing
  Set oTask = Nothing

  Exit Function
err_here:
  'Debug.Print Task.UniqueID & ": " & Task.Name
  Call cptHandleErr("cptMetrics_bas", "cptGetMetric", Err, Erl)
  Resume exit_here

End Function

Sub cptShowMetricsSettings_frm(Optional blnModal As Boolean = False)
  'objects
  'strings
  Dim strCustomName As String
  Dim strLOE As String
  Dim strLOEField As String
  Dim strEVP As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  With cptMetricsSettings_frm
  
    .cboEVP.Clear
    .cboEVP.AddItem
    .cboEVP.List(.cboEVP.ListCount - 1, 0) = FieldNameToFieldConstant("Physical % Complete")
    .cboEVP.List(.cboEVP.ListCount - 1, 1) = "Physical % Complete"
    For lngItem = 1 To 20
      .cboEVP.AddItem
      .cboEVP.List(.cboEVP.ListCount - 1, 0) = FieldNameToFieldConstant("Number" & lngItem)
      .cboEVP.List(.cboEVP.ListCount - 1, 1) = "Number" & lngItem
      strCustomName = CustomFieldGetName(FieldNameToFieldConstant("Number" & lngItem))
      If Len(strCustomName) > 0 Then
        .cboEVP.List(.cboEVP.ListCount - 1, 1) = strCustomName & " (Number" & lngItem & ")"
      End If
    Next lngItem
    
    .cboLOEField.Clear
    For lngItem = 1 To 30
      .cboLOEField.AddItem
      .cboLOEField.List(.cboLOEField.ListCount - 1, 0) = FieldNameToFieldConstant("Text" & lngItem)
      .cboLOEField.List(.cboLOEField.ListCount - 1, 1) = "Text" & lngItem
      strCustomName = CustomFieldGetName(FieldNameToFieldConstant("Text" & lngItem))
      If Len(strCustomName) > 0 Then
        .cboLOEField.List(.cboLOEField.ListCount - 1, 1) = strCustomName & " (Text" & lngItem & ")"
      End If
    Next lngItem
    
    strEVP = cptGetSetting("Metrics", "cboEVP")
    If Len(strEVP) > 0 Then .cboEVP.Value = CLng(strEVP)
    strLOEField = cptGetSetting("Metrics", "cboLOEField")
    If Len(strLOEField) > 0 Then .cboLOEField.Value = CLng(strLOEField)
    strLOE = cptGetSetting("Metrics", "txtLOE")
    If Len(strLOE) > 0 Then .txtLOE = strLOE
    If blnModal Then
      .Show
    Else
      .Show False
    End If
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  'Call HandleErr("cptMetrics_bas", "cptShowMetricsSettings_frm", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Function cptMetricsSettingsExist() As Boolean
  'objects
  'strings
  Dim strLOE As String
  Dim strLOEField As String
  Dim strEVP As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
    
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strEVP = cptGetSetting("Metrics", "cboEVP")
  strLOEField = cptGetSetting("Metrics", "cboLOEField")
  strLOE = cptGetSetting("Metrics", "txtLOE")
  
  If Len(strEVP) = 0 Or Len(strLOEField) = 0 Or Len(strLOE) = 0 Then
    cptMetricsSettingsExist = False
  Else
    cptMetricsSettingsExist = True
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptMetricsSettingsExist", Err, Erl)
  Resume exit_here
End Function


Sub cptCaptureWeek()
  'objects
  Dim oTasks As Tasks
  Dim oTask As Task
  Dim rst As ADODB.Recordset
  'strings
  Dim strLOE As String
  Dim strEVT As String
  Dim strProject As String
  Dim strFile As String
  Dim strDir As String
  'longs
  Dim lngEVT As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtStatus As Date
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'ensure program acronym
  strProject = cptGetProgramAcronym
  If Len(strProject) = 0 Then
    MsgBox "Program Acronym is required for this feature.", vbExclamation + vbOKOnly, "Program Acronym Needed"
    GoTo exit_here
  End If
    
  Set rst = CreateObject("ADODB.Recordset")
  strFile = cptDir & "\settings\cpt-cei.adtg"
  If Dir(strFile) = vbNullString Then
    rst.Fields.Append "PROJECT", adVarChar, 25      '0
    rst.Fields.Append "TASK_UID", adInteger         '1
    rst.Fields.Append "TASK_NAME", adVarChar, 255   '2
    rst.Fields.Append "IS_LOE", adInteger           '3
    rst.Fields.Append "TASK_BLS", adDate            '4
    rst.Fields.Append "TASK_BLD", adInteger         '5
    rst.Fields.Append "TASK_BLF", adDate            '6
    rst.Fields.Append "TASK_AS", adDate             '7
    rst.Fields.Append "TASK_AD", adInteger          '8
    rst.Fields.Append "TASK_AF", adDate             '9
    rst.Fields.Append "TASK_START", adDate          '10
    rst.Fields.Append "TASK_RD", adInteger          '11
    rst.Fields.Append "TASK_FINISH", adDate         '12
    rst.Fields.Append "STATUS_DATE", adDate         '13
    rst.Open
  Else
    rst.Open strFile
  End If
  
  dtStatus = ActiveProject.StatusDate
  If rst.RecordCount > 0 Then
    rst.Find "STATUS_DATE=#" & FormatDateTime(dtStatus, vbGeneralDate) & "# AND PROJECT='" & strProject & "'"
    If Not rst.EOF Then
      If MsgBox("Status Already Imported for WE " & FormatDateTime(dtStatus, vbShortDate) & "." & vbCrLf & vbCrLf & "Overwrite it?", vbExclamation + vbYesNo, "Overwrite?") = vbYes Then
        rst.MoveFirst
        Do While Not rst.EOF
          If rst("PROJECT") = strProject And rst("STATUS_DATE") = FormatDateTime(dtStatus, vbGeneralDate) Then rst.Delete adAffectCurrent
          rst.MoveNext
        Loop
      End If
    End If
  End If
  
  strEVT = cptGetSetting("Metrics", "cboLOEField")
  If Len(strEVT) > 0 Then
    lngEVT = CLng(strEVT)
  Else
    'todo: settings needed
  End If
  strLOE = cptGetSetting("Metrics", "txtLOE")
  If Len(strLOE) = 0 Then
    'todo: settings needed
  End If
  
  Set oTasks = ActiveProject.Tasks
  lngTasks = oTasks.Count
  'include all discrete, LOE, milestones, and all SVTs
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task 'skip blank lines
    If Not oTask.Active Then GoTo next_task 'skip inactive
    If oTask.ExternalTask Then GoTo next_task 'skip external
    If oTask.Summary Then GoTo next_task 'skip summaries
    'If oTask.Milestone Then GoTo next_task 'skip milestones
    If oTask.Resources.Count > 0 Or InStr(oTask.Name, "SVT") > 0 Then
      rst.AddNew
      rst(0) = strProject
      rst(1) = oTask.UniqueID
      rst(2) = oTask.Name
      rst(3) = IIf(oTask.GetField(lngEVT) = strLOE, 1, 0)
      If IsDate(oTask.BaselineStart) Then
        rst(4) = FormatDateTime(oTask.BaselineStart, vbGeneralDate)
        rst(5) = Round(oTask.BaselineDuration / (60 * 8), 0)
      End If
      If IsDate(oTask.BaselineFinish) Then
        rst(6) = FormatDateTime(oTask.BaselineFinish, vbGeneralDate)
      End If
      If IsDate(oTask.ActualStart) Then
        rst(7) = FormatDateTime(oTask.ActualStart, vbGeneralDate)
        rst(8) = Round(oTask.ActualDuration / (60 * 8), 0)
      End If
      If IsDate(oTask.ActualFinish) Then
        rst(9) = FormatDateTime(oTask.ActualFinish, vbGeneralDate)
      End If
      rst(10) = FormatDateTime(oTask.Start, vbGeneralDate)
      rst(11) = Round(oTask.RemainingDuration / (60 * 8), 0)
      rst(12) = FormatDateTime(oTask.Finish, vbGeneralDate)
      rst(13) = FormatDateTime(ActiveProject.StatusDate, vbGeneralDate)
      rst.Update
    End If
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = lngTask & " / " & lngTasks & " (" & Format(lngTask / lngTasks, "0%")
  Next oTask
  
  rst.Save strFile, adPersistADTG
  rst.Close
  Application.StatusBar = "Complete."
  MsgBox "Current Schedule as of " & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & " captured.", vbInformation + vbOKOnly, "Complete"
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oTasks = Nothing
  Set oTask = Nothing
  If rst.State = 1 Then rst.Close
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("focptMetrics_bas", "cptCaptureWeek", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub


Sub cptLateStartsFinishes()
  'objects
  Dim oSeries As Excel.Series
  Dim oChart As Excel.ChartObject
  Dim oShape As Excel.Shape
  Dim oOutlook As Outlook.Application
  Dim oMailItem As Outlook.MailItem
  Dim oDocument As Word.Document
  Dim oWord As Word.Application
  Dim oSelection As Word.Selection
  Dim oEmailTemplate As Word.Template
  Dim oWorksheet As Excel.Worksheet
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oListObject As Excel.ListObject
  Dim oRange As Excel.Range
  Dim oCell As Excel.Range
  Dim oAssignment As MSProject.Assignment
  Dim oTask As Task
  'strings
  Dim strLOE As String
  Dim strLOEField As String
  Dim strCC As String
  Dim strTo As String
  Dim strProject As String
  Dim strFile As String
  Dim strDir As String
  'longs
  Dim lngLOEField As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngForecastCount As Long
  Dim lngBaselineCount As Long
  Dim lngLastRow As Long
  Dim lngCol As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vCol As Variant
  'dates
  Dim dtDate As Date
  Dim dtMax As Date
  Dim dtMin As Date
  Dim dtStatus As Date
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  dtStatus = ActiveProject.StatusDate
  
  If MsgBox("not ready for prime time!" & vbCrLf & vbCrLf & "proceed anyway?", vbCritical + vbYesNo, "note to self") = vbNo Then GoTo exit_here
  
  strProject = cptGetProgramAcronym
  
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then Set oExcel = CreateObject("Excel.Application")
  'oExcel.Visible = True
  Set oWorkbook = oExcel.Workbooks.Add
  oExcel.Calculation = xlCalculationManual
  oExcel.ScreenUpdating = False
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "DETAILS"
  
  'todo: user-defined fields to include
  strLOEField = cptGetSetting("Metrics", "lngLOEField")
  If Len(strLOEField) > 0 Then
    lngLOEField = CLng(strLOEField)
  Else
    'todo: error
  End If
  strLOE = cptGetSetting("Metrics", "strLOE")
  If Len(strLOE) = 0 Then
    'todo: error
  End If
  
  oWorksheet.[A1:P1] = Array("UID", "WPCN", "WPM", "NAME", "TOTAL SLACK", "REMAINING DURATION", "REMAINING WORK", "BASELINE START", "START VARIANCE", "ACTUAL START", "START", "BASELINE FINISH", "FINISH VARIANCE", "ACTUAL FINISH", "FINISH", "STATUS")
  
  lngTasks = ActiveProject.Tasks.Count
  
  For Each oTask In ActiveProject.Tasks
    If Not oTask Is Nothing Then
      oTask.Marked = False
      'skip inactive tasks
      If Not oTask.Active Then GoTo next_task
      'skip summaries
      If oTask.Summary Then GoTo next_task
      'only check for tasks with assignments
      If oTask.Resources.Count = 0 Then GoTo next_task
      'only check for discrete tasks
      If oTask.GetField(FieldNameToFieldConstant("EVT")) = "A" Then GoTo next_task
      'skip unassigned (currently material/odc/tvl)
      If oTask.GetField(FieldNameToFieldConstant("WPM")) = "" Then GoTo next_task
      'only report early/late starts/finishes
      If oTask.StartVariance <> 0 Or oTask.FinishVariance <> 0 Then
        lngLastRow = oWorksheet.Cells(1048576, 1).End(xlUp).Row + 1
        oWorksheet.Cells(lngLastRow, 1) = oTask.UniqueID
        oWorksheet.Cells(lngLastRow, 2) = oTask.GetField(FieldNameToFieldConstant("WPCN"))
        oWorksheet.Cells(lngLastRow, 3) = oTask.GetField(FieldNameToFieldConstant("WPM"))
        oWorksheet.Cells(lngLastRow, 4) = oTask.Name
        oWorksheet.Cells(lngLastRow, 5) = Round(oTask.TotalSlack / (8 * 60), 0)
        oWorksheet.Cells(lngLastRow, 6) = oTask.RemainingDuration / (8 * 60)
        oWorksheet.Cells(lngLastRow, 7) = Round(oTask.RemainingWork / 60, 0)
        
        oWorksheet.Cells(lngLastRow, 8) = FormatDateTime(oTask.BaselineStart, vbShortDate)
        oWorksheet.Cells(lngLastRow, 9) = Round(oTask.StartVariance / (8 * 60), 0)
        If IsDate(oTask.ActualStart) Then
          oWorksheet.Cells(lngLastRow, 10) = FormatDateTime(oTask.ActualStart, vbShortDate)
        End If
        oWorksheet.Cells(lngLastRow, 11) = FormatDateTime(oTask.Start, vbShortDate)
        
        oWorksheet.Cells(lngLastRow, 12) = FormatDateTime(oTask.BaselineFinish, vbShortDate)
        oWorksheet.Cells(lngLastRow, 13) = Round(oTask.FinishVariance / (8 * 60), 0)
        If IsDate(oTask.ActualFinish) Then
          oWorksheet.Cells(lngLastRow, 14) = FormatDateTime(oTask.ActualFinish, vbShortDate)
        End If
        oWorksheet.Cells(lngLastRow, 15) = FormatDateTime(oTask.Finish, vbShortDate)
        
        oWorksheet.Cells(lngLastRow, 16) = oTask.GetField(FieldNameToFieldConstant("Task Status"))
        
      End If
    End If
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Exporting BEI... " & Format(lngTask, "#,##0") & " / " & Format(lngTasks, "#,##0") & " (" & Format(lngTask / lngTasks, "0%") & ")"
    DoEvents
  Next oTask

  Application.StatusBar = "Analyzing..."

  oWorksheet.Cells(1, oWorksheet.Rows(1).Find("START", lookat:=xlWhole).Column).Value = "CURRENT START"
  oWorksheet.Cells(1, oWorksheet.Rows(1).Find("FINISH", lookat:=xlWhole).Column).Value = "CURRENT FINISH"
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True

  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True

  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)).Address, , xlYes)
  
  oListObject.HeaderRowRange.WrapText = True
  oListObject.TableStyle = ""
  oWorksheet.Columns.AutoFit
  oListObject.ListColumns(5).Range.ColumnWidth = 10
  For lngCol = 6 To 15
    oListObject.ListColumns(lngCol).Range.ColumnWidth = 12
  Next lngCol
  oListObject.HeaderRowRange.EntireRow.AutoFit
  
  'create summary worksheet
  Set oWorksheet = oWorkbook.Sheets.Add(oWorkbook.Sheets(1))
  oWorksheet.Name = "SUMMARY"
  oWorksheet.[A1] = strProject & " IMS - Early/Late Starts/Finishes"
  oWorksheet.[A1].Font.Bold = True
  oWorksheet.[A1].Font.Size = 14
  oWorksheet.[A2].Value = FormatDateTime(dtStatus, vbShortDate)
  oWorksheet.Names.Add "STATUS_DATE", oWorksheet.[A2].Address
  
  oListObject.ListColumns("WPM").Range.Copy oWorksheet.[A5]
  oWorksheet.Range(oWorksheet.[A6], oWorksheet.[A1048576]).RemoveDuplicates Columns:=1, Header:=xlNo
  oWorksheet.Sort.SortFields.Clear
  oWorksheet.Sort.SortFields.Add2 key:=oWorksheet.Range(oWorksheet.[A6], oWorksheet.[A6].End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With oWorksheet.Sort
    .SetRange oWorksheet.Range(oWorksheet.[A6], oWorksheet.[A1048576].End(xlUp))
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With
  
  'todo: nuance for Critical/Driving?
  oWorksheet.[B4].Value = "ACTUAL"
  oWorksheet.[B4:H4].Merge True
  oWorksheet.[B4:H4].HorizontalAlignment = xlCenter
  oWorksheet.[B4:H4].Font.Bold = True
  'todo: interior
  oWorksheet.[B5:H5] = Array("ES", "EF", "LS", "LF", "# BLF", "# AF", "BEI (Finishes)")
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A5].End(xlToRight), oWorksheet.[A5].End(xlDown)), , xlYes)
  oListObject.TableStyle = "TableStyleMedium2"
  oListObject.Name = "BEI"
  
  'ACTUAL
  oListObject.ListColumns("ES").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[WPM],RC1,Table1[START VARIANCE],""<0"",Table1[ACTUAL START],""<>"")"
  oListObject.ListColumns("EF").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[WPM],RC1,Table1[FINISH VARIANCE],""<0"",Table1[ACTUAL FINISH],""<>"")"
  oListObject.ListColumns("LS").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[WPM],RC1,Table1[START VARIANCE],"">0"",Table1[ACTUAL START],""<>"")"
  oListObject.ListColumns("LF").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[WPM],RC1,Table1[FINISH VARIANCE],"">0"",Table1[ACTUAL FINISH],""<>"")"
  oListObject.ListColumns("# BLF").DataBodyRange.FormulaR1C1 = "=COUNTIFS(Table1[WPM],[@WPM],Table1[BASELINE FINISH],""<=" & Format(dtStatus, "mm/dd/yyyy") & """)"
  oListObject.ListColumns("# AF").DataBodyRange.FormulaR1C1 = "=COUNTIFS(Table1[WPM],[@WPM],Table1[ACTUAL FINISH],""<>"")"
  oListObject.ListColumns("BEI (Finishes)").DataBodyRange.FormulaR1C1 = "=[@['# AF]]/IF([@['# BLF]]=0,1,[@['# BLF]])"
  oListObject.ListColumns("BEI (Finishes)").DataBodyRange.Style = "Comma"
  oListObject.ShowTotals = True
  oListObject.ListColumns("ES").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("EF").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("LS").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("LF").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("# BLF").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("# AF").TotalsCalculation = xlTotalsCalculationSum
  oListObject.TotalsRowRange(ColumnIndex:=oListObject.ListColumns("BEI (Finishes)").DataBodyRange.Column).FormulaR1C1 = "=BEI[[#Totals],['# AF]]/BEI[[#Totals],['# BLF]]"
  oListObject.ListColumns("BEI (Finishes)").DataBodyRange.Style = "Comma"
  oListObject.TotalsRowRange(ColumnIndex:=oListObject.ListColumns("BEI (Finishes)").DataBodyRange.Column).Style = "Comma"

  'PROJECTED
  lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 2
  oWorksheet.Cells(lngLastRow, 2).Value = "PROJECTED"
  oWorksheet.Range(oWorksheet.Cells(lngLastRow, 2), oWorksheet.Cells(lngLastRow, 2).Offset(0, 5)).Merge True
  oWorksheet.Cells(lngLastRow, 2).HorizontalAlignment = xlCenter
  oWorksheet.Cells(lngLastRow, 2).Font.Bold = True
  oListObject.Range.Copy oWorksheet.Cells(oWorksheet.[A1048576].End(xlUp).Row + 3, 1)
  Set oListObject = oWorksheet.ListObjects(2)
  oListObject.Name = "PROJECTED"
  oListObject.ListColumns("ES").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[WPM],RC1,Table1[START VARIANCE],""<0"",Table1[ACTUAL START],""="")"
  oListObject.ListColumns("EF").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[WPM],RC1,Table1[FINISH VARIANCE],""<0"",Table1[ACTUAL FINISH],""="")"
  oListObject.ListColumns("LS").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[WPM],RC1,Table1[START VARIANCE],"">0"",Table1[ACTUAL START],""="")"
  oListObject.ListColumns("LF").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[WPM],RC1,Table1[FINISH VARIANCE],"">0"",Table1[ACTUAL FINISH],""="")"
  oListObject.ListColumns("# BLF").Name = "# TOTAL"
  oListObject.ListColumns("# TOTAL").DataBodyRange.FormulaR1C1 = "=COUNTIFS(Table1[WPM],[@WPM])"
  oListObject.ListColumns("# AF").Name = "% TOTAL"
  oListObject.ListColumns("% TOTAL").DataBodyRange.FormulaR1C1 = "=[@[LF]]/IF([@['# TOTAL]]=0,1,[@['# TOTAL]])"
  oListObject.ListColumns("% TOTAL").DataBodyRange.Style = "Comma"
  oListObject.TotalsRowRange(ColumnIndex:=oListObject.ListColumns("% TOTAL").DataBodyRange.Column).FormulaR1C1 = "=PROJECTED[[#Totals],[LF]]/IF(PROJECTED[[#Totals],['# TOTAL]]=0,1,PROJECTED[[#Totals],['# TOTAL]])"
  oListObject.TotalsRowRange(ColumnIndex:=oListObject.ListColumns("% TOTAL").DataBodyRange.Column).Style = "Comma"
  oListObject.ListColumns("BEI (Finishes)").Delete
    
  oExcel.ActiveWindow.DisplayGridLines = False
  oExcel.ActiveWindow.Zoom = 85
  oListObject.Range.Columns.AutoFit
  
  'week,BLF,AF,CF (BEI/S-chart)
  'get earliest start and latest finish
  Set oListObject = oWorkbook.Sheets("DETAILS").ListObjects(1)
  dtMin = oExcel.WorksheetFunction.Min(oListObject.ListColumns("Baseline Start").DataBodyRange)
  dtMin = oExcel.WorksheetFunction.Min(dtMin, oListObject.ListColumns("Actual Start").DataBodyRange)
  dtMin = oExcel.WorksheetFunction.Min(dtMin, oListObject.ListColumns("Current Start").DataBodyRange)
  'convert to WE Friday
  dtMin = DateAdd("d", 6 - Weekday(dtMin), dtMin)
  dtMax = oExcel.WorksheetFunction.Max(oListObject.ListColumns("Baseline Finish").DataBodyRange)
  dtMax = oExcel.WorksheetFunction.Max(dtMax, oListObject.ListColumns("Actual Finish").DataBodyRange)
  dtMax = oExcel.WorksheetFunction.Max(dtMax, oListObject.ListColumns("Current Finish").DataBodyRange)
  dtMax = DateAdd("d", 6 - Weekday(dtMax), dtMax)
  
  Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
  oWorksheet.Name = "Chart_Data"
  oWorksheet.[A1:D1] = Array("WEEK", "BLF", "AF", "FF")
  lngLastRow = 2
  dtDate = dtMin & " 5:00 PM"
  oWorksheet.Cells(lngLastRow, 1) = dtDate
  Do While dtDate <= dtMax
    dtDate = DateAdd("d", 7, dtDate)
    lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
    oWorksheet.Cells(lngLastRow, 1) = dtDate
  Loop
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)))
  oListObject.Name = "ChartData"
  oListObject.ListColumns("BLF").DataBodyRange.Formula2R1C1 = "=SUMPRODUCT((--Table1[BASELINE FINISH]<=[@WEEK])*--(Table1[BASELINE FINISH]>R[-1]C[-1])*1)"
  oListObject.ListColumns("AF").DataBodyRange.Formula2R1C1 = "=SUMPRODUCT((--Table1[ACTUAL FINISH]<=[@WEEK])*--(Table1[ACTUAL FINISH]>R[-1]C[-2])*1)"
  oListObject.ListColumns("FF").DataBodyRange.Formula2R1C1 = "=SUMPRODUCT((--Table1[CURRENT FINISH]<=[@WEEK])*--(Table1[CURRENT FINISH]>R[-1]C[-3])*--(Table1[ACTUAL FINISH]="""")*1)"
  oWorksheet.[I1] = dtStatus
  oWorksheet.[E1] = "BLF_CUM"
  oListObject.ListColumns("BLF_CUM").DataBodyRange.FormulaR1C1 = "=IF(ROW(R[-1]C)=1,[@BLF],R[-1]C+[@BLF])"
  oWorksheet.[F1] = "AF_CUM"
  oListObject.ListColumns("AF_CUM").DataBodyRange.FormulaR1C1 = "=IF(ROW(R[-1]C)=1,[@AF],IF([@WEEK]<=R1C9,R[-1]C+[@AF],""""))"
  oWorksheet.[G1] = "FF_CUM"
  oListObject.ListColumns("FF_CUM").DataBodyRange.FormulaR1C1 = "=IF([@WEEK]=R1C9,[@[AF_CUM]],IF([@WEEK]>R1C9,R[-1]C+[@FF],""""))"
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True
  oListObject.Range.Columns.AutoFit
  oListObject.DataBodyRange.Copy
  oListObject.DataBodyRange.PasteSpecial xlPasteValuesAndNumberFormats
  lngLastRow = oWorksheet.Columns(1).Find(dtStatus).Row
  oWorksheet.Range(oWorksheet.Cells(2, 7), oWorksheet.Cells(lngLastRow - 1, 7)).ClearContents
  oWorksheet.Range(oWorksheet.Cells(lngLastRow + 1, 6), oWorksheet.Cells(1048576, 6)).ClearContents
  oWorksheet.[I1].Select
  oWorksheet.Shapes.AddChart2 227, xlLine
  Set oChart = oWorksheet.ChartObjects(oWorksheet.ChartObjects.Count)
  oChart.Chart.FullSeriesCollection(1).Delete
  oChart.Chart.SeriesCollection.NewSeries
  oChart.Chart.FullSeriesCollection(1).Name = "=Chart_Data!$E$1"
  oChart.Chart.FullSeriesCollection(1).Values = "=Chart_Data!" & oListObject.ListColumns("BLF_CUM").DataBodyRange.Address(True)
  oChart.Chart.FullSeriesCollection(1).XValues = "=Chart_Data!" & oListObject.ListColumns("WEEK").DataBodyRange.Address(True)
  oChart.Chart.SeriesCollection.NewSeries
  oChart.Chart.FullSeriesCollection(2).Name = "=Chart_Data!$F$1"
  oChart.Chart.FullSeriesCollection(2).Values = "=Chart_Data!" & oListObject.ListColumns("AF_CUM").DataBodyRange.Address(True)
  oChart.Chart.SeriesCollection.NewSeries
  oChart.Chart.FullSeriesCollection(3).Name = "=Chart_Data!$G$1"
  oChart.Chart.FullSeriesCollection(3).Values = "=Chart_Data!" & oListObject.ListColumns("FF_CUM").DataBodyRange.Address(True)
  oChart.Chart.SetElement (msoElementChartTitleAboveChart)
  oChart.Chart.SetElement (msoElementLegendBottom)
  oChart.Chart.ChartTitle.Text = strProject & " IMS - Task Completion" & Chr(10) & Format(dtStatus, "mm/dd/yyyy")
  oChart.Chart.ChartTitle.Characters(1, 25).Font.Bold = True
  oChart.Chart.Location Where:=xlLocationAsObject, Name:="SUMMARY"
  'must reset the object after move
  oWorksheet.Visible = xlSheetHidden
  Set oWorksheet = oWorkbook.Sheets("SUMMARY")
  Set oShape = oWorksheet.Shapes(oWorksheet.Shapes.Count)
  oShape.Top = oWorksheet.[J5].Top
  oShape.Left = oWorksheet.[J5].Left
  oShape.ScaleWidth 1.6663381968, msoFalse, msoScaleFromTopLeft
  oShape.ScaleHeight 1.8082112132, msoFalse, msoScaleFromTopLeft
  Set oChart = oWorksheet.ChartObjects(1)
  Set oSeries = oChart.Chart.SeriesCollection(1)
  With oSeries.Format.Line
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorText1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
  End With
  Set oSeries = oChart.Chart.FullSeriesCollection(3)
  With oSeries.Format.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(0, 112, 192)
    .Transparency = 0
    .DashStyle = msoLineDash
  End With
  oChart.Chart.Axes(xlCategory).CategoryType = xlTimeScale
  oChart.Chart.Axes(xlCategory).TickLabels.NumberFormat = "m/d/yyyy"

  Set oWorksheet = oWorkbook.Worksheets("Chart_Data")
  lngBaselineCount = oWorksheet.[E1048576].End(xlUp).Value
  lngForecastCount = oWorksheet.[G1048576].End(xlUp).Value
  If lngForecastCount < lngBaselineCount Then
    oWorkbook.Sheets("Summary").[J31] = "There are " & lngBaselineCount - lngForecastCount & " unstatused tasks in the current IMS."
    oWorkbook.Sheets("Summary").[J31].Font.Italic = True
    With oWorkbook.Sheets("Summary").[J31].Font
      .Color = -16777024
      .TintAndShade = 0
    End With
  End If
  
  oWorkbook.Sheets("Summary").Activate
  oWorkbook.Sheets("Summary").[A2].Select
  
  'save the file
  'todo: user-defined locations for metrics output
  strDir = ActiveProject.Path & "\Metrics\"
  strDir = strDir & Format(dtStatus, "yyyy-mm-dd") & "\"
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  strFile = strDir & Replace(strProject, " ", "_") & "_IMS_EarlyLateStartsFinishes_" & Format(ActiveProject.StatusDate, "yyyy-mm-dd") & ".xlsx"
  If Dir(strFile) <> vbNullString Then Kill strFile
  oExcel.Calculation = xlCalculationAutomatic
  oExcel.ScreenUpdating = True
  oWorkbook.SaveAs strFile, 51
  oWorkbook.Close True
'  'send the file
'  On Error Resume Next
'  Set oOutlook = GetObject(, "Outlook.Application")
'  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
'  If oOutlook Is Nothing Then
'    Set oOutlook = CreateObject("Outlook.Application")
'  End If
'  Set oMailItem = oOutlook.CreateItem(olMailItem)
'  oMailItem.Display
'  oMailItem.Subject = strProject & " IMS - Early/Late Starts/Finishes WE " & FormatDateTime(ActiveProject.StatusDate, vbShortDate)
'  oMailItem.Attachments.Add strFile
'  oMailItem.To = strTo
'  oMailItem.CC = strCC
'  oMailItem.HTMLBody = strProject & " IMS Early/Late Starts/Finishes for week ending " & Format(dtStatus, "mm/dd/yyyy") & " attached." & oMailItem.HTMLBody
  
  Application.StatusBar = "Complete."
  
  If MsgBox("Complete. Open for review?", vbInformation + vbYesNo, "Late Starts and Finishes") = vbYes Then
    oExcel.Workbooks.Open strFile
    oExcel.Visible = True
    oExcel.ScreenUpdating = True
    Application.ActivateMicrosoftApp pjMicrosoftExcel
  End If

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oSeries = Nothing
  Set oChart = Nothing
  Set oShape = Nothing
  Set oOutlook = Nothing
  Set oMailItem = Nothing
  Set oDocument = Nothing
  Set oWord = Nothing
  Set oSelection = Nothing
  Set oEmailTemplate = Nothing
  Set oWorksheet = Nothing
  Set oCell = Nothing
  Set oRange = Nothing
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oTask = Nothing
  Set oShape = Nothing
  Set oChart = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics", "cptLateStartsFinishes", Err, Erl)
  Resume exit_here
End Sub

Sub cptCaptureMetric(strProgram As String, dtStatus As Date, strMetric As String, vMetric As Variant)
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set oRecordset = CreateObject("ADODB.Recordset")
  strFile = cptDir & "\settings\cpt-metrics.adtg"
  With oRecordset
    If Dir(strFile) = vbNullString Then
      'create it
      .Fields.Append "PROGRAM", adVarChar, 50
      .Fields.Append "STATUS_DATE", adDate
      .Fields.Append "SPI", adDouble
      .Fields.Append "SV", adDouble
      .Fields.Append "BEI", adDouble
      .Fields.Append "CPLI", adDouble
      .Fields.Append "CEI", adDouble
      .Fields.Append "TFCI", adDouble
      'others needed for ES?
      .Fields.Append "ES", adDate
      .Open
    Else
      .Open strFile
    End If
    .Filter = "PROGRAM='" & strProgram & "' AND STATUS_DATE=#" & dtStatus & "#"
    'todo: mechanism for correcting if .recordcount > 1
    If Not .EOF Then
      .Update Array(strMetric), Array(CDbl(vMetric))
    Else
      .AddNew Array("PROGRAM", "STATUS_DATE", strMetric), Array(strProgram, dtStatus, CDbl(vMetric))
    End If
    .Filter = ""
    .Save strFile, adPersistADTG
    .Close
  End With


exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptCaptureMetric", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetTrend(strMetric As String, Optional dtStatus As Date)
  'objects
  Dim oRecordset As ADODB.Recordset
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  'strings
  Dim strProgram As String
  Dim strFile As String
  'longs
  Dim lngLastRow As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  strFile = cptDir & "\settings\cpt-metrics.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox strFile & " not found.", vbExclamation + vbOKOnly, "File Not Found"
    GoTo exit_here
  Else
    'get program
    strProgram = cptGetProgramAcronym
    If Len(strProgram) = 0 Then GoTo exit_here
    'get status date
    If dtStatus = 0 Then
      If Not IsDate(ActiveProject.StatusDate) Then
        MsgBox "This project requires a Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
        GoTo exit_here
      End If
      dtStatus = ActiveProject.StatusDate
    End If
    'get excel
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If oExcel Is Nothing Then
      Set oExcel = CreateObject("Excel.Application")
    End If
    oExcel.Visible = True
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.Name = strMetric & " TREND"
    Set oRecordset = CreateObject("ADODB.Recordset")
    With oRecordset
      .Open strFile
      .Sort = "STATUS_DATE"
      .Filter = "PROGRAM='" & strProgram & "' AND STATUS_DATE<=#" & dtStatus & "#"
      If .RecordCount = 0 Then
        'handle that
      End If
      .MoveFirst
      oWorksheet.Cells(1, 1) = strProgram
      oWorksheet.Cells(3, 1) = "STATUS_DATE"
      oWorksheet.Cells(3, 2) = strMetric
      'todo: banding
      Do While Not .EOF
        lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
        oWorksheet.Cells(lngLastRow, 1) = FormatDateTime(CDate(.Fields("STATUS_DATE")), vbShortDate)
        oWorksheet.Cells(lngLastRow, 2) = .Fields(strMetric)
        'todo: banding
        .MoveNext
      Loop
      .Close
    End With
  End If

  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True
  oWorksheet.Columns.AutoFit

exit_here:
  On Error Resume Next
  Set oRecordset = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  'Call HandleErr("cptMetrics_bas", "cptGetTrend", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub
