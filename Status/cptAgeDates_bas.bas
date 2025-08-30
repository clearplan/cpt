Attribute VB_Name = "cptAgeDates_bas"
'<cpt_version>v1.1.1</cpt_version>
Option Explicit

Sub cptShowAgeDates_frm()
  'objects
  Dim myAgeDates_frm As cptAgeDates_frm
  'strings
  Dim strSetting As String
  'longs
  Dim lngFF As Long
  Dim lngFS As Long
  Dim lngAF As Long
  Dim lngAS As Long
  Dim lngControl As Long
  Dim lngWeek As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  'prevent spawning
  If Not cptGetUserForm("cptAgeDates_frm") Is Nothing Then Exit Sub

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Status Date required.", vbExclamation + vbOKOnly, "Age Dates"
    Application.ChangeStatusDate
  End If
  
  'todo: create and apply table
  'todo: create and apply filter?
  'todo: create and apply view
  'todo: update table dynamically
  'todo: multiple settings for multiple projects >> convert to ADODB
  
  'todo: avoid cpt conflicts with status sheet import
  strSetting = cptGetSetting("StatusSheetImport", "cboAS")
  If Len(strSetting) > 0 Then lngAS = CLng(strSetting) Else lngAS = 0
  strSetting = cptGetSetting("StatusSheetImport", "cboAF")
  If Len(strSetting) > 0 Then lngAF = CLng(strSetting) Else lngAF = 0
  strSetting = cptGetSetting("StatusSheetImport", "cboFS")
  If Len(strSetting) > 0 Then lngFS = CLng(strSetting) Else lngFS = 0
  strSetting = cptGetSetting("StatusSheetImport", "cboFF")
  If Len(strSetting) > 0 Then lngFF = CLng(strSetting) Else lngFF = 0
  
  Set myAgeDates_frm = New cptAgeDates_frm
  With myAgeDates_frm
    .Caption = "Age Dates (" & cptGetVersion("cptAgeDates_frm") & ")"
    .lblStatus = "(" & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & ")"
    .cboWeeks.Clear
    For lngWeek = 1 To 10
      .cboWeeks.AddItem
      .cboWeeks.List(.cboWeeks.ListCount - 1, 0) = lngWeek & IIf(lngWeek = 1, " week", " weeks")
      For lngControl = 1 To 10
        With .Controls("cboWeek" & lngControl)
          .AddItem
          .List(.ListCount - 1, 0) = lngWeek + 1
          .List(.ListCount - 1, 1) = lngWeek
          .List(.ListCount - 1, 2) = "Start" & lngWeek & "/Finish" & lngWeek
        End With
      Next lngControl
    Next lngWeek
    strSetting = cptGetSetting("AgeDates", "cboWeeks")
    If Len(strSetting) > 0 Then
      .cboWeeks.Value = strSetting
    Else
      .cboWeeks.Value = "3 weeks"
    End If
    For lngControl = 1 To 10
      strSetting = cptGetSetting("AgeDates", "cboWeek" & lngControl)
      If Len(strSetting) > 0 Then
        .Controls("cboWeek" & lngControl).Value = cptGetSetting("AgeDates", "cboWeek" & lngControl)
      End If
    Next lngControl
    strSetting = cptGetSetting("AgeDates", "chkIncludeDurations")
    If Len(strSetting) > 0 Then
      .chkIncludeDurations = CBool(strSetting)
    Else
      .chkIncludeDurations = True
    End If
    strSetting = cptGetSetting("AgeDates", "chkUpdateCustomFieldNames")
    If Len(strSetting) > 0 Then
      .chkUpdateCustomFieldNames = CBool(strSetting)
    Else
      .chkUpdateCustomFieldNames = True
    End If
    
    .Show False
  End With
  
exit_here:
  On Error Resume Next
  Set myAgeDates_frm = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptAgeDates_bas", "cptShowAgeDates_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptAgeDates(ByRef myAgeDates_frm As cptAgeDates_frm)
  'run this immediately prior to a status meeting
  'objects
  Dim oTask As MSProject.Task
  'strings
  Dim strCustom As String
  Dim strStatus As String
  'longs
  Dim lngTo As Long
  Dim lngFrom As Long
  Dim lngTest As Long
  Dim lngControl As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Application.Calculation = pjManual
  Application.OpenUndoTransaction "Age Dates"
  dtStatus = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  
  On Error Resume Next
  lngTest = FieldNameToFieldConstant("Start (" & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & ")")
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If lngTest > 0 Then
    MsgBox "Dates already aged for status date " & FormatDateTime(dtStatus, vbShortDate) & ".", vbExclamation + vbOKOnly, "Age Dates"
    GoTo exit_here
  End If

  With myAgeDates_frm
    
    For lngControl = 10 To 1 Step -1
      If .Controls("cboWeek" & lngControl).Enabled Then
        If lngControl = 1 Then
          lngFrom = 0
        Else
          lngFrom = .Controls("cboWeek" & lngControl - 1).List(.Controls("cboWeek" & lngControl - 1).ListIndex, 0)
        End If
        lngTo = .Controls("cboWeek" & lngControl).List(.Controls("cboWeek" & lngControl).ListIndex, 1)
        BaselineSave True, lngFrom, lngTo
        'update custom field names
        If .chkUpdateCustomFieldNames Then
          If lngControl = 1 Then
            strCustom = "Start (" & FormatDateTime(dtStatus, vbShortDate) & ")"
            CustomFieldRename FieldNameToFieldConstant("Start" & lngControl), strCustom
            strCustom = "Finish (" & FormatDateTime(dtStatus, vbShortDate) & ")"
            CustomFieldRename FieldNameToFieldConstant("Finish" & lngControl), strCustom
            If .chkIncludeDurations Then
              strCustom = "Duration (" & FormatDateTime(dtStatus, vbShortDate) & ")"
              CustomFieldRename FieldNameToFieldConstant("Duration" & lngControl), strCustom
            End If
          Else
            strCustom = CustomFieldGetName(FieldNameToFieldConstant("Start" & lngControl - 1, pjTask))
            CustomFieldRename FieldNameToFieldConstant("Start" & lngControl - 1), ""
            CustomFieldRename FieldNameToFieldConstant("Start" & lngControl), strCustom
            strCustom = CustomFieldGetName(FieldNameToFieldConstant("Finish" & lngControl - 1, pjTask))
            CustomFieldRename FieldNameToFieldConstant("Finish" & lngControl - 1), ""
            CustomFieldRename FieldNameToFieldConstant("Finish" & lngControl), strCustom
            If .chkIncludeDurations Then
              strCustom = CustomFieldGetName(FieldNameToFieldConstant("Duration" & lngControl - 1, pjTask))
              CustomFieldRename FieldNameToFieldConstant("Duration" & lngControl - 1), ""
              CustomFieldRename FieldNameToFieldConstant("Duration" & lngControl), strCustom
            End If
          End If
        End If
      End If
    Next lngControl
    
    If .chkIncludeDurations Then
      For Each oTask In ActiveProject.Tasks
        If oTask Is Nothing Then GoTo next_task
        If Not oTask.Active Then GoTo next_task
        If oTask.ExternalTask Then GoTo next_task
        For lngControl = 10 To 1 Step -1
          If .Controls("cboWeek" & lngControl).Enabled Then
            lngTo = cptRegEx(.Controls("cboWeek" & lngControl).List(.Controls("cboWeek" & lngControl).ListIndex, 1), "[0-9]")
            If lngControl = 1 Then
              oTask.SetField FieldNameToFieldConstant("Duration" & lngTo), oTask.DurationText
            Else
              lngFrom = cptRegEx(.Controls("cboWeek" & lngControl - 1).List(.Controls("cboWeek" & lngControl - 1).ListIndex, 2), "[0-9]")
              oTask.SetField FieldNameToFieldConstant("Duration" & lngTo), oTask.GetField(FieldNameToFieldConstant("Duration" & lngFrom))
            End If
          End If
        Next lngControl
next_task:
      Next oTask
    End If
  End With
        
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Application.Calculation = pjAutomatic
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAgeDates_bas", "cptAgeDates", Err, Erl)
  Resume exit_here
End Sub

Sub cptBlameReport()
  'objects
  Dim oRange As Excel.Range
  Dim oListObject As Excel.ListObject
  Dim oTask As MSProject.Task
  Dim oWorksheet As Excel.Worksheet 'Object
  Dim oWorkbook As Excel.Workbook 'Object
  Dim oExcel As Excel.Application 'Object
  'strings
  Dim strMyHeaders As String
  Dim strHeaders As String
  Dim strWeek1 As String
  'longs
  Dim lngMin As Long
  Dim lngResponse As Long
  Dim lngMyHeaders As Long
  Dim lngCol As Long
  Dim lngLastRow As Long
  Dim lngLastCol As Long
  Dim lngFinish1 As Long
  Dim lngDuration1 As Long
  Dim lngField As Long
  Dim lngStart1 As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vMyHeader As Variant
  Dim vBorder As Variant
  Dim vColumn As Variant
  Dim vRow As Variant
  'dates
  Dim dtCurr As Date
  Dim dtPrev As Date
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Application.StatusBar = "Gettings saved settings..."
  DoEvents
  
  'todo: allow user to designate which previous data to use
  strWeek1 = cptGetSetting("AgeDates", "cboWeek1")
  If Len(strWeek1) = 0 Then
    MsgBox "Please designate fields and Age Dates before running The Blame Report.", vbExclamation + vbOKOnly, "The Blame Report"
    Call cptShowAgeDates_frm
    GoTo exit_here
  End If
  
  Application.StatusBar = "Getting custom fields..."
  DoEvents
  
  strMyHeaders = cptGetMyHeaders("The Blame Report")
    
  cptSpeed True
  
  lngField = CLng(cptRegEx(strWeek1, "[0-9]{1,}"))
  lngStart1 = FieldNameToFieldConstant("Start" & lngField)
  lngDuration1 = FieldNameToFieldConstant("Duration" & lngField)
  lngFinish1 = FieldNameToFieldConstant("Finish" & lngField)
  
  Application.StatusBar = "Getting Microsoft Excel..."
  DoEvents
  
  On Error Resume Next
  'Set oExcel = GetObject(, "Excel.Application")
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  
  Application.StatusBar = "Creating Workbook..."
  DoEvents
  
  Set oWorkbook = oExcel.Workbooks.Add
  oExcel.Calculation = xlCalculationManual
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "Blame"
  strHeaders = "UID," & strMyHeaders
  strHeaders = strHeaders & "TASK,PREVIOUS START,CURRENT START,START DELTA,PREVIOUS DURATION,CURRENT DURATION,DURATION DELTA,PREVIOUS FINISH,CURRENT FINISH,FINISH DELTA"
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].Offset(0, UBound(Split(strHeaders, ",")))) = Split(strHeaders, ",") 'Array("UID", "TASK", "PREVIOUS START", "CURRENT START", "START DELTA", "PREVIOUS DURATION", "CURRENT DURATION", "DURATION DELTA", "PREVIOUS FINISH", "CURRENT FINISH", "FINISH DELTA")
  lngLastCol = oWorksheet.[A1].End(xlToRight).Column
  lngTasks = ActiveProject.Tasks.Count
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If IsDate(oTask.ActualFinish) Then GoTo next_task
    lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
    vRow = oWorksheet.Range(oWorksheet.Cells(lngLastRow, 1), oWorksheet.Cells(lngLastRow, lngLastCol))
    vRow(1, 1) = oTask.UniqueID
    lngCol = 1
    For Each vMyHeader In Split(strMyHeaders, ",")
      If vMyHeader = "" Then Exit For
      lngCol = lngCol + 1
      vRow(1, lngCol) = oTask.GetField(FieldNameToFieldConstant(vMyHeader))
    Next
    lngMyHeaders = UBound(Split(strMyHeaders, ","))
    vRow(1, 2 + lngMyHeaders) = oTask.Name
    If IsDate(oTask.ActualStart) Then
      vRow(1, 3 + lngMyHeaders) = Null
      With oWorksheet.Cells(lngLastRow, 4 + lngMyHeaders).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
      End With
      vRow(1, 5 + lngMyHeaders) = Null
    ElseIf IsDate(oTask.GetField(lngStart1)) Then
      dtPrev = CDate(FormatDateTime(oTask.GetField(lngStart1), vbShortDate))
      dtCurr = CDate(FormatDateTime(oTask.Start, vbShortDate))
      vRow(1, 3 + lngMyHeaders) = dtPrev
      If dtCurr >= dtPrev Then 'slipped
        vRow(1, 5 + lngMyHeaders) = -(Application.DateDifference(dtPrev, dtCurr) / (60 * 8))
      Else 'pulled left
        vRow(1, 5 + lngMyHeaders) = Application.DateDifference(dtCurr, dtPrev) / (60 * 8)
      End If
    Else
      vRow(1, 3 + lngMyHeaders) = "NA"
      vRow(1, 5 + lngMyHeaders) = Null
    End If
    vRow(1, 4 + lngMyHeaders) = CDate(FormatDateTime(oTask.Start, vbShortDate))
    
    vRow(1, 6 + lngMyHeaders) = CLng(cptRegEx(oTask.GetField(lngDuration1), "[0-9]{1,}"))
    vRow(1, 7 + lngMyHeaders) = cptRegEx(oTask.Duration, "[0-9]{1,}") / (60 * 8)
    vRow(1, 8 + lngMyHeaders) = "=RC[-2]-RC[-1]"
    
    If IsDate(oTask.GetField(lngFinish1)) Then
      dtPrev = CDate(FormatDateTime(oTask.GetField(lngFinish1), vbShortDate))
      dtCurr = CDate(FormatDateTime(oTask.Finish, vbShortDate))
      vRow(1, 9 + lngMyHeaders) = dtPrev
      If dtCurr >= dtPrev Then 'slipped
        vRow(1, 11 + lngMyHeaders) = -(Application.DateDifference(dtPrev, dtCurr) / (60 * 8))
      Else 'pulled left
        vRow(1, 11 + lngMyHeaders) = Application.DateDifference(dtCurr, dtPrev) / (60 * 8)
      End If
    Else
      vRow(1, 9 + lngMyHeaders) = "NA"
      vRow(1, 11 + lngMyHeaders) = Null
    End If
    vRow(1, 10 + lngMyHeaders) = CDate(FormatDateTime(oTask.Finish, vbShortDate))
    oWorksheet.Range(oWorksheet.Cells(lngLastRow, 1), oWorksheet.Cells(lngLastRow, lngLastCol)) = vRow
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Analyzing..." & Format(lngTask, "#,##0") & " / " & Format(lngTasks, "#,##0") & "...(" & Format(lngTask / lngTasks, "0%") & ")"
    DoEvents
  Next oTask
  
  Application.StatusBar = "Formatting report..."
  DoEvents
  
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlDown), oWorksheet.[A1].End(xlToRight)), , xlYes)
  oListObject.HeaderRowRange.Font.Bold = True
  oListObject.HeaderRowRange.HorizontalAlignment = xlLeft
  oListObject.TableStyle = ""
  With oListObject
    .TableStyle = ""
    .DataBodyRange.Borders(xlDiagonalDown).LineStyle = xlNone
    .DataBodyRange.Borders(xlDiagonalUp).LineStyle = xlNone
    For Each vBorder In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
      With .DataBodyRange.Borders(vBorder)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
      End With
      With .HeaderRowRange.Borders(vBorder)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
      End With
    Next vBorder
    For Each vBorder In Array(xlInsideVertical, xlInsideHorizontal)
      With .DataBodyRange.Borders(vBorder)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.14996795556505
        .Weight = xlThin
      End With
      With .HeaderRowRange.Borders(vBorder)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.14996795556505
        .Weight = xlThin
      End With
    Next
  End With
  
  With oListObject.HeaderRowRange.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = -0.149998474074526
    .PatternTintAndShade = 0
  End With
  
  'conditional formatting
  For Each vColumn In Array("START DELTA", "DURATION DELTA", "FINISH DELTA")
    Set oRange = oListObject.ListColumns(vColumn).DataBodyRange
    oRange.FormatConditions.AddColorScale ColorScaleType:=3
    oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
    oRange.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    With oRange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
      .Color = 7039480
      .TintAndShade = 0
    End With
    oRange.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
    oRange.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With oRange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
      .Color = 16776444
      .TintAndShade = 0
    End With
    oRange.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    With oRange.FormatConditions(1).ColorScaleCriteria(3).FormatColor
      .Color = 8109667
      .TintAndShade = 0
    End With
  Next vColumn
  
  For Each vColumn In Array("UID", "PREVIOUS START", "CURRENT START", "PREVIOUS FINISH", "CURRENT FINISH")
    oListObject.ListColumns(vColumn).DataBodyRange.HorizontalAlignment = xlCenter
  Next vColumn
  
  For Each vColumn In Array("START DELTA", "PREVIOUS DURATION", "CURRENT DURATION", "DURATION DELTA", "FINISH DELTA")
    oListObject.ListColumns(vColumn).DataBodyRange.Style = "Comma"
  Next vColumn
  
  oWorksheet.Columns.AutoFit
  
  Application.StatusBar = "Creating Header..."
  DoEvents
  
  oWorksheet.Rows(1).Insert
  oWorksheet.Rows(1).Insert
  oWorksheet.Rows(1).Insert
  oWorksheet.Rows(1).Insert
  oWorksheet.[A1].Value = "The Blame Report"
  oWorksheet.[A1].Font.Bold = True
  oWorksheet.[A1].Font.Size = 24
  oWorksheet.[A2] = ActiveProject.Name
  oWorksheet.[A3] = "Status Date: " & FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  oExcel.ActiveWindow.DisplayGridlines = False
  oExcel.Calculation = xlCalculationAutomatic
  
  Application.StatusBar = "Determining worst offender..."
  DoEvents
  
  lngMin = oExcel.WorksheetFunction.Min(oListObject.ListColumns("START DELTA").DataBodyRange, oListObject.ListColumns("DURATION DELTA").DataBodyRange, oListObject.ListColumns("FINISH DELTA").DataBodyRange)
  For Each vColumn In Array("START DELTA", "DURATION DELTA", "FINISH DELTA")
    If lngMin = oExcel.WorksheetFunction.Min(oListObject.ListColumns(vColumn).DataBodyRange) Then
      oListObject.Sort.SortFields.Clear
      oListObject.Sort.SortFields.Add Key:=oListObject.ListColumns(vColumn).DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      oListObject.Sort.SortFields.Add Key:=oListObject.ListColumns("CURRENT " & Replace(vColumn, " DELTA", "")).DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      With oListObject.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
      End With
      oWorksheet.Cells(oListObject.HeaderRowRange.Row, oListObject.ListColumns(vColumn).Range.Column).Select
      Exit For
    End If
  Next vColumn
  
  Application.StatusBar = "Opening The Blame Report..."
  DoEvents
  
  oExcel.Visible = True
  oExcel.WindowState = xlMaximized
  Application.StatusBar = "Complete."
  DoEvents
  Application.ActivateMicrosoftApp pjMicrosoftExcel
  
exit_here:
  On Error Resume Next
  Set oRange = Nothing
  Set oListObject = Nothing
  cptSpeed False
  Application.StatusBar = ""
  Set oTask = Nothing
  oExcel.Calculation = xlCalculationAutomatic
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAgeDates_bas", "cptBlameReport", Err, Erl)
  Resume exit_here
End Sub
