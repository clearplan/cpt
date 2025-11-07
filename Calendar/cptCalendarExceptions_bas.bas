Attribute VB_Name = "cptCalendarExceptions_bas"
'<cpt_version>v1.2.0</cpt_version>
Option Explicit

Sub cptShowCalendarExceptions_frm()
  'objects
  Dim myCalendarExceptions_frm As cptCalendarExceptions_frm
  Dim oResource As Resource
  'strings
  Dim strCalendar As String
  Dim strDetail As String
  Dim strErrors As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnDetail As Boolean
  'variants
  'dates
  
  'prevent spawning
  If Not cptGetUserForm("cptCalendarExceptions_frm") Is Nothing Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set myCalendarExceptions_frm = New cptCalendarExceptions_frm
  With myCalendarExceptions_frm
    .Caption = "Calendar Details (" & cptGetVersion("cptCalendarExceptions_frm") & ")"
    .cboCalendars.Clear
    .cboCalendars.AddItem "All Calendars"
    For lngItem = 1 To ActiveProject.BaseCalendars.Count
      .cboCalendars.AddItem ActiveProject.BaseCalendars(lngItem).Name
    Next lngItem
    For Each oResource In ActiveProject.Resources
      If oResource Is Nothing Then GoTo next_resource
      If oResource.Type <> pjResourceTypeWork Then GoTo next_resource
      If oResource.Name <> oResource.Calendar.Name Then Debug.Print oResource.Name & ": " & oResource.Calendar.Name
      If Not oResource.Calendar Is Nothing Then
        If oResource.Calendar.Exceptions.Count > 0 Or oResource.Calendar.WorkWeeks.Count > 0 Then
          If Len(oResource.Calendar.Name) > 0 Then
            'todo: only add a calendar once
            'todo: what does it mean when a resource's calendar is named differently than the resource itself?
            .cboCalendars.AddItem oResource.Calendar.Name
          Else
            strErrors = strErrors & "Resource UID " & oResource.UniqueID & " is unnamed." & vbCrLf
          End If
        End If
      End If
next_resource:
    Next oResource
    .cboCalendars.Value = ActiveProject.Calendar.Name ' "All Calendars"
    strDetail = cptGetSetting("CalendarDetails", "optDetailed")
    If strDetail <> "" Then
      blnDetail = CBool(strDetail)
    Else
      blnDetail = False
    End If
    .optSummary = Not blnDetail
    .optDetailed = blnDetail
    If Len(strErrors) = 0 Then
      Application.StatusBar = "Ready..."
      .Show ' False
    Else
      MsgBox strErrors & vbCrLf & vbCrLf & "Please name these resources then try again.", vbCritical + vbOKOnly, "Error"
    End If
  End With

exit_here:
  On Error Resume Next
  Set oResource = Nothing
  Unload myCalendarExceptions_frm
  
  Exit Sub
err_here:
  Call cptHandleErr("cptCalendarExceptions_bas", "cptShowCalendarExceptions_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCalendarExceptionsMain(ByRef myCalendarExceptions_frm As cptCalendarExceptions_frm, Optional blnDetail As Boolean = False)
  'objects
  Dim oListObject As ListObject
  Dim oResource As Resource
  Dim oCalendar As MSProject.Calendar
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  'strings
  'longs
  Dim lngStartCol As Long
  Dim lngLastRow As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vBorder As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'set up file
  Application.StatusBar = "Setting up Excel Workbook..."
  DoEvents
  'get new instance of Excel
  Set oExcel = CreateObject("Excel.Application")
  Set oWorkbook = oExcel.Workbooks.Add
  cptSpeed True
  oWorkbook.Sheets(1).Name = "Exceptions"
  If oWorkbook.Sheets.Count < 2 Then oWorkbook.Sheets.Add After:=oWorkbook.Sheets(1)
  oWorkbook.Sheets(2).Name = "WorkWeeks"
  If oWorkbook.Sheets.Count < 3 Then oWorkbook.Sheets.Add After:=oWorkbook.Sheets(2)
  oWorkbook.Sheets(3).Name = "Settings"
  
  'export base calendars
  For Each oCalendar In ActiveProject.BaseCalendars
    If myCalendarExceptions_frm.cboCalendars = "All Calendars" Then
      cptExportCalendarExceptions oWorkbook, oCalendar, blnDetail
    ElseIf oCalendar.Name = myCalendarExceptions_frm.cboCalendars.Value Then
      cptExportCalendarExceptions oWorkbook, oCalendar, blnDetail
    End If
  Next oCalendar
  'export resource calendars
  For Each oResource In ActiveProject.Resources
    If oResource Is Nothing Then GoTo next_resource
    If Not oResource.Calendar Is Nothing Then
      If myCalendarExceptions_frm.cboCalendars.Value = "All Calendars" Then
        cptExportCalendarExceptions oWorkbook, oResource.Calendar, blnDetail
      ElseIf oResource.Calendar.Name = myCalendarExceptions_frm.cboCalendars.Value Then
        cptExportCalendarExceptions oWorkbook, oResource.Calendar, blnDetail
      End If
    End If
next_resource:
  Next oResource
  'export application settings
  Application.StatusBar = "Formatting Worksheet..."
  DoEvents
  Set oWorksheet = oWorkbook.Sheets("Settings")
  oWorksheet.[A1:A6].Value = oExcel.WorksheetFunction.Transpose(Array("Setting", "DefaultStartTime:", "DefaultFinishTime:", "HoursPerDay:", "HoursPerWeek:", "DaysPerMonth:"))
  oWorksheet.[A1:B1].Font.Bold = True
  oWorksheet.[B1] = "Value"
  oWorksheet.[B2] = ActiveProject.DefaultStartTime
  oWorksheet.[B3] = ActiveProject.DefaultFinishTime
  oWorksheet.[B4] = ActiveProject.HoursPerDay
  oWorksheet.[B5] = ActiveProject.HoursPerWeek
  oWorksheet.[B6] = ActiveProject.DaysPerMonth
  
  'clean up the worksheet
  For Each oWorksheet In oWorkbook.Worksheets
    oWorksheet.Activate
    oExcel.ActiveWindow.Zoom = 85
    oExcel.ActiveWindow.SplitRow = 1
    oExcel.ActiveWindow.SplitColumn = 0
    oExcel.ActiveWindow.FreezePanes = True
    If oWorksheet.Name <> "Settings" Then oWorksheet.[A1].AutoFilter
    oWorksheet.Columns.AutoFit
    If oWorksheet.Name = "Exceptions" Then
      lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
      oWorksheet.Range(oWorksheet.Cells(2, 3), oWorksheet.Cells(lngLastRow, 10)).HorizontalAlignment = xlCenter
      oWorksheet.Range(oWorksheet.Cells(2, 5), oWorksheet.Cells(lngLastRow, 5)).HorizontalAlignment = xlLeft
      oWorksheet.Range(oWorksheet.Cells(2, 13), oWorksheet.Cells(lngLastRow, 13)).HorizontalAlignment = xlCenter
      'make it a listobject
      Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)), , xlYes)
      With oListObject
        .Name = "_" & UCase(oWorksheet.Name)
        .TableStyle = ""
        'format the table
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
        'highlight duplicate start dates
        lngStartCol = oWorksheet.Rows(1).Find(what:="START", lookat:=xlWhole).Column
        .ListColumns(lngStartCol).Range.FormatConditions.Delete
        .ListColumns(lngStartCol).Range.FormatConditions.AddUniqueValues
        With .ListColumns(lngStartCol).Range.FormatConditions(1)
          .DupeUnique = xlDuplicate
          .Font.Color = -16383844
          .Font.TintAndShade = 0
          .Interior.Color = 13551615
          .Interior.TintAndShade = 0
          .StopIfTrue = False
        End With
        'collapse to level one if details
        If blnDetail Then oWorksheet.Outline.ShowLevels RowLevels:=1
      End With
    End If
  Next oWorksheet
  
  Application.StatusBar = "Complete."
  MsgBox "Export complete.", vbInformation + vbOKOnly, "Fiscal Export"
  DoEvents
  
  'show the user what you've done for them
  oExcel.Visible = True
  oExcel.WindowState = xlMaximized
  oExcel.Windows(oExcel.Windows.Count).Activate
  oWorkbook.Sheets("Exceptions").Activate

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oListObject = Nothing
  Set oResource = Nothing
  Set oCalendar = Nothing
  If Not oExcel Is Nothing Then oExcel.Visible = True
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCalendarExceptions_bas", "cptExportCalendarExceptionsMain", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCalendarExceptions(ByRef oWorkbook As Excel.Workbook, ByRef oCalendar As MSProject.Calendar, Optional blnDetail As Boolean = False)
  'objects
  Dim oExcel As Excel.Application
  Dim oWorksheet As Excel.Worksheet
  Dim oWeekDay As WorkWeekDay
  Dim oWorkWeek As WorkWeek
  Dim oException As MSProject.Exception
  'strings
  Dim strDaysOfWeekRev As String
  Dim strDaysOfWeek As String
  Dim strRecord As String
  Dim strException As String
  'longs
  Dim lngDay As Long
  Dim lngWeekDay As Long
  Dim lngCalendarCol As Long
  Dim lngPeriodCol As Long
  Dim lngMonthPositionCol As Long
  Dim lngMonthItemCol As Long
  Dim lngMonthDayCol As Long
  Dim lngMonthCol As Long
  Dim lngOccurrencesCol As Long
  Dim lngDaysCol As Long
  Dim lngTypeCol As Long
  Dim lngHoursCol As Long
  Dim lngFinishCol As Long
  Dim lngStartCol As Long
  Dim lngNameCol As Long
  Dim lngLastRow As Long
  Dim lngDaysOfWeek As Long
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vDayOfWeek As Variant
  Dim vDaysOfWeek As Variant
  'dates
  Dim dtDate As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'exceptions
  Set oWorksheet = oWorkbook.Sheets("Exceptions")
  Set oExcel = oWorksheet.Application
  oWorksheet.Activate
  oWorksheet.[A1:M1] = Array("CALENDAR", "EXCEPTION", "START", "FINISH", "TYPE", "WORKING_HOURS", "DAYS_OF_WEEK", "OCCURRENCES", "MONTH", "MONTH_DAY", "MONTH_ITEM", "MONTH_POSITION", "PERIOD")
  lngCalendarCol = oWorksheet.Rows(1).Find(what:="CALENDAR", lookat:=xlWhole).Column
  lngNameCol = oWorksheet.Rows(1).Find(what:="EXCEPTION", lookat:=xlWhole).Column
  lngStartCol = oWorksheet.Rows(1).Find(what:="START", lookat:=xlWhole).Column
  lngFinishCol = oWorksheet.Rows(1).Find(what:="FINISH", lookat:=xlWhole).Column
  lngHoursCol = oWorksheet.Rows(1).Find(what:="WORKING_HOURS", lookat:=xlWhole).Column
  lngTypeCol = oWorksheet.Rows(1).Find(what:="TYPE", lookat:=xlWhole).Column
  lngDaysCol = oWorksheet.Rows(1).Find(what:="DAYS_OF_WEEK", lookat:=xlWhole).Column
  lngOccurrencesCol = oWorksheet.Rows(1).Find(what:="OCCURRENCES", lookat:=xlWhole).Column
  lngMonthCol = oWorksheet.Rows(1).Find(what:="MONTH", lookat:=xlWhole).Column
  lngMonthDayCol = oWorksheet.Rows(1).Find(what:="MONTH_DAY", lookat:=xlWhole).Column
  lngMonthItemCol = oWorksheet.Rows(1).Find(what:="MONTH_ITEM", lookat:=xlWhole).Column
  lngMonthPositionCol = oWorksheet.Rows(1).Find(what:="MONTH_POSITION", lookat:=xlWhole).Column
  lngPeriodCol = oWorksheet.Rows(1).Find(what:="PERIOD", lookat:=xlWhole).Column
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True
  
  If oCalendar.Exceptions.Count > 0 Then
    For Each oException In oCalendar.Exceptions
      If Len(oException.Name) = 0 Then strException = "[Unnamed]" Else strException = oException.Name
      Application.StatusBar = "Processing Calendar '" & oCalendar.Name & "' Exception '" & strException & "'..."
      DoEvents
      lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
      With oException
        oWorksheet.Cells(lngLastRow, lngCalendarCol) = oCalendar.Name
        oWorksheet.Cells(lngLastRow, lngNameCol) = strException
        oWorksheet.Cells(lngLastRow, lngStartCol) = .Start
        oWorksheet.Cells(lngLastRow, lngFinishCol) = .Finish
        oWorksheet.Cells(lngLastRow, lngHoursCol) = cptGetShifts(oException)
        oWorksheet.Cells(lngLastRow, lngTypeCol) = .Type & " - " & Choose(.Type, "pjDaily", "pjYearlyMonthDay", "pjYearlyPositional", "pjMonthlyMonthDay", "pjMonthlyPositional", "pjWeekly", "pjDayCount", "pjWeekDayCount")
        '1 pjDaily: The exception recurrence pattern is daily.
        '2 pjYearlyMonthDay: The exception recurrence pattern is yearly on a specified day of a month, for example on December 24.
        '3 pjYearlyPositional: The exception recurrence pattern is yearly on a specified position of a day in a month, for example the fourth Friday of every month.
        '4 pjMonthlyMonthDay: The exception recurrence pattern is monthly on a specified day, for example the 24th of the month.
        '5 pjMonthlyPositional: The exception recurrence pattern is monthly on a specified position of a day in a month, for example the fourth Friday.
        '6 pjWeekly: The exception recurrence pattern is weekly.
        '7 pjDayCount: The exception daily recurrence ends after a specified number of occurrences.
        '8 pjWeekDayCount: The exception recurrence ends after a specified number of weekday occurrences.
        
        'where Exception.Type = pjWeekly(6) only
        'Sunday=1, Monday=2, Tuesday=4, Wednesday=8, Thursday=16, Friday=32, Saturday=64
        lngDaysOfWeek = .DaysOfWeek
        strDaysOfWeek = ""
        If lngDaysOfWeek > 0 Then
          If lngDaysOfWeek >= 64 Then
            strDaysOfWeek = "Sa,"
            lngDaysOfWeek = lngDaysOfWeek - 64
          End If
          If lngDaysOfWeek >= 32 Then
            strDaysOfWeek = "F," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 32
          End If
          If lngDaysOfWeek >= 16 Then
            strDaysOfWeek = "Th," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 16
          End If
          If lngDaysOfWeek >= 8 Then
            strDaysOfWeek = "W," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 8
          End If
          If lngDaysOfWeek >= 4 Then
            strDaysOfWeek = "T," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 4
          End If
          If lngDaysOfWeek >= 2 Then
            strDaysOfWeek = "M," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 2
          End If
          If lngDaysOfWeek >= 1 Then
            strDaysOfWeek = "Su," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 1
          End If
          strDaysOfWeek = Left(strDaysOfWeek, Len(strDaysOfWeek) - 1)
          oWorksheet.Cells(lngLastRow, lngDaysCol) = strDaysOfWeek
        End If
        
        oWorksheet.Cells(lngLastRow, lngOccurrencesCol) = .Occurrences
        
        If .Month > 0 Then
          oWorksheet.Cells(lngLastRow, lngMonthCol) = .Month  '& " - " & Choose(.Month, "pjJanuary", "pjFebruary", "pjMarch", "pjApril", "pjMay", "pjJune", "pjJuly", "pjAugust", "pjSeptember", "pjOctober", "pjNovember", "pjDecember") & ","
          'type = pjYearlyMonthDay/pjYearlyPositional
          'pjJanuary = 1...pjDecember = 12
        End If
        
        If .MonthDay > 0 Then
          oWorksheet.Cells(lngLastRow, lngMonthDayCol) = .MonthDay
          'type = pjMonthlyMonthDay/pjMonthlyPositional
          'day of the month 1-31
        End If
        
        If .MonthItem > 0 Then
          oWorksheet.Cells(lngLastRow, lngMonthItemCol) = .MonthItem & " - " & Choose(.MonthItem, "", "", "pjItemSunday", "pjItemMonday", "pjItemTuesday", "pjItemWednesday", "pjItemThursday", "pjItemFriday", "pjItemSaturday")
          'type = pjMonthlyMonthDay/pjMonthlyPositional
          'pjItemSunday = 3...pjItemSaturday = 9
        End If
        
        If .Type = pjMonthlyPositional Or .Type = pjYearlyPositional Then
          oWorksheet.Cells(lngLastRow, lngMonthPositionCol) = .MonthPosition & " - " & Choose(.MonthPosition + 1, "pjFirst", "pjSecond", "pjThird", "pjFourth", "pjLast")
          'type = pjMonthlyMonthDay/pjMonthlyPositional/pjYearlyMontyDay/pjYearlyPositional
          'pjFirst = 0...pjLast =4
        End If
        
        If .Period > 0 Then
          oWorksheet.Cells(lngLastRow, lngPeriodCol) = .Period
          'e.g., every X [Days, Weeks, Months, Years]
        End If
        
        If blnDetail Then
          With oWorksheet.Outline
            .AutomaticStyles = False
            .SummaryRow = xlAbove
            .SummaryColumn = xlRight
          End With
        
          Select Case .Type
            Case pjDaily '1
              dtDate = .Start
              Do While dtDate < .Finish
                dtDate = DateAdd("d", 1, dtDate)
                lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
                oWorksheet.Cells(lngLastRow, lngCalendarCol) = oCalendar.Name
                oWorksheet.Cells(lngLastRow, lngNameCol) = .Name
                oWorksheet.Cells(lngLastRow, lngStartCol) = dtDate
                oWorksheet.Cells(lngLastRow, lngHoursCol) = cptGetShifts(oException)
                Call cptFormatExceptionDetail(oWorksheet.Cells(lngLastRow, 1))
              Loop
            Case pjYearlyMonthDay '2
              If .Start <> DateValue(.Month & "/" & .MonthDay & "/" & Year(.Start)) Then
                dtDate = CDate(DateValue(.Month & "/" & .MonthDay & "/" & Year(.Start)))
                'fix the start date - no, don't: notify user it's wrong?
                'oWorksheet.Cells(lngLastRow, lngStartCol) = dtDate
              Else
                dtDate = .Start
              End If
              Do While dtDate <= .Finish
                dtDate = DateAdd("yyyy", 1, dtDate)
                If dtDate <= .Finish Then
                  lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
                  oWorksheet.Cells(lngLastRow, lngCalendarCol) = oCalendar.Name
                  oWorksheet.Cells(lngLastRow, lngNameCol) = .Name
                  oWorksheet.Cells(lngLastRow, lngStartCol) = dtDate
                  oWorksheet.Cells(lngLastRow, lngHoursCol) = cptGetShifts(oException)
                  Call cptFormatExceptionDetail(oWorksheet.Cells(lngLastRow, 1))
                End If
              Loop
            Case pjYearlyPositional '3
              dtDate = .Start
              Do While dtDate < .Finish
                dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, 11) + 1
                'find first .MonthItem
                If Weekday(dtDate) <= .MonthItem - 2 Then
                  dtDate = DateAdd("d", .MonthItem - 2 - Weekday(dtDate), dtDate)
                Else
                  dtDate = DateAdd("d", 7 + .MonthItem - 2 - Weekday(dtDate), dtDate)
                End If
                'adjust for position
                If .MonthPosition < pjLast Then
                  dtDate = DateAdd("d", 7 * .MonthPosition, dtDate)
                ElseIf .MonthPosition = pjLast Then
                  'find last day of month
                  dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, 0)
                  If Weekday(dtDate) <> (.MonthItem - 2) Then
                    For lngDay = Day(dtDate) To 1 Step -1
                      If Weekday(dtDate) = .MonthItem - 2 Then
                        'bingo
                        Exit For
                      Else
                        dtDate = DateAdd("d", -1, dtDate)
                      End If
                    Next lngDay
                  End If
                End If
                If dtDate <= .Finish Then
                  lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
                  oWorksheet.Cells(lngLastRow, lngCalendarCol) = oCalendar.Name
                  oWorksheet.Cells(lngLastRow, lngNameCol) = .Name
                  oWorksheet.Cells(lngLastRow, lngStartCol) = dtDate
                  oWorksheet.Cells(lngLastRow, lngHoursCol) = cptGetShifts(oException)
                  Call cptFormatExceptionDetail(oWorksheet.Cells(lngLastRow, 1))
                End If
              Loop
            Case pjMonthlyMonthDay '4
              dtDate = .Start
              Do While dtDate <= .Finish
                dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, .Period - 1) + .MonthDay
                If dtDate <= .Finish Then
                  lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
                  oWorksheet.Cells(lngLastRow, lngCalendarCol) = oCalendar.Name
                  oWorksheet.Cells(lngLastRow, lngNameCol) = .Name
                  oWorksheet.Cells(lngLastRow, lngStartCol) = dtDate
                  oWorksheet.Cells(lngLastRow, lngHoursCol) = cptGetShifts(oException)
                  Call cptFormatExceptionDetail(oWorksheet.Cells(lngLastRow, 1))
                End If
              Loop
            Case pjMonthlyPositional '5
              dtDate = .Start
              Do While dtDate <= .Finish
                'find first day of next month
                dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, .Period - 1) + 1
                'find first .MonthItem
                If Weekday(dtDate) <= .MonthItem - 2 Then
                  dtDate = DateAdd("d", .MonthItem - 2 - Weekday(dtDate), dtDate)
                Else
                  dtDate = DateAdd("d", 7 + .MonthItem - 2 - Weekday(dtDate), dtDate)
                End If
                'adjust for position
                If .MonthPosition < pjLast Then
                  dtDate = DateAdd("d", 7 * .MonthPosition, dtDate)
                Else
                  'find last day of month
                  dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, 0)
                  If Weekday(dtDate) <> (.MonthItem - 2) Then
                    For lngDay = Day(dtDate) To 1 Step -1
                      If Weekday(dtDate) = .MonthItem - 2 Then
                        'bingo
                        Exit For
                      Else
                        dtDate = DateAdd("d", -1, dtDate)
                      End If
                    Next lngDay
                  End If
                End If
                If dtDate <= .Finish Then
                  lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
                  oWorksheet.Cells(lngLastRow, lngCalendarCol) = oCalendar.Name
                  oWorksheet.Cells(lngLastRow, lngNameCol) = .Name
                  oWorksheet.Cells(lngLastRow, lngStartCol) = dtDate
                  oWorksheet.Cells(lngLastRow, lngHoursCol) = cptGetShifts(oException)
                  Call cptFormatExceptionDetail(oWorksheet.Cells(lngLastRow, 1))
                End If
              Loop
            Case pjWeekly '6
              If .DaysOfWeek > 0 Then
                lngDaysOfWeek = .DaysOfWeek
                strDaysOfWeek = ""
                If lngDaysOfWeek >= 64 Then
                  strDaysOfWeek = "7,"
                  lngDaysOfWeek = lngDaysOfWeek - 64
                End If
                If lngDaysOfWeek >= 32 Then
                  strDaysOfWeek = "6," & strDaysOfWeek
                  lngDaysOfWeek = lngDaysOfWeek - 32
                End If
                If lngDaysOfWeek >= 16 Then
                  strDaysOfWeek = "5," & strDaysOfWeek
                  lngDaysOfWeek = lngDaysOfWeek - 16
                End If
                If lngDaysOfWeek >= 8 Then
                  strDaysOfWeek = "4," & strDaysOfWeek
                  lngDaysOfWeek = lngDaysOfWeek - 8
                End If
                If lngDaysOfWeek >= 4 Then
                  strDaysOfWeek = "3," & strDaysOfWeek
                  lngDaysOfWeek = lngDaysOfWeek - 4
                End If
                If lngDaysOfWeek >= 2 Then
                  strDaysOfWeek = "2," & strDaysOfWeek
                  lngDaysOfWeek = lngDaysOfWeek - 2
                End If
                If lngDaysOfWeek >= 1 Then
                  strDaysOfWeek = "1," & strDaysOfWeek
                End If
                For Each vDayOfWeek In Split(strDaysOfWeek, ",")
                  If Len(vDayOfWeek) = 0 Then Exit For
                  If Weekday(.Start) = CLng(vDayOfWeek) Then
                    dtDate = .Start
                  Else
                    dtDate = DateAdd("d", CLng(vDayOfWeek) - Weekday(.Start), .Start)
                    lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
                    oWorksheet.Cells(lngLastRow, lngCalendarCol) = oCalendar.Name
                    oWorksheet.Cells(lngLastRow, 1) = .Name
                    oWorksheet.Cells(lngLastRow, 2) = dtDate
                    Call cptFormatExceptionDetail(oWorksheet.Cells(lngLastRow, 1))
                  End If
                  Do While dtDate <= .Finish
                    dtDate = DateAdd("d", 7 * .Period, dtDate)
                    If dtDate <= .Finish Then
                      lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
                      oWorksheet.Cells(lngLastRow, lngCalendarCol) = oCalendar.Name
                      oWorksheet.Cells(lngLastRow, lngNameCol) = .Name
                      oWorksheet.Cells(lngLastRow, lngStartCol) = dtDate
                      oWorksheet.Cells(lngLastRow, lngHoursCol) = cptGetShifts(oException)
                      Call cptFormatExceptionDetail(oWorksheet.Cells(lngLastRow, 1))
                    End If
                  Loop
                Next vDayOfWeek
                oWorksheet.Range(oWorksheet.Cells(lngLastRow, 2), oWorksheet.Cells(lngLastRow - .Occurrences + 1, 2)).Sort oWorksheet.Cells(lngLastRow, 2)
              End If 'daysofweek>0
            
          End Select 'case .Type
        End If 'blnDetail
      End With 'oException
    Next oException
  End If 'oExceptions.Count > 0
  
  'get work weeks
  Set oWorksheet = oWorkbook.Sheets("WorkWeeks")
  oWorksheet.Activate
  oWorksheet.[A1:J1] = Array("CALENDAR", "WORK WEEK", "START", "FINISH", "DAY", "WORKING", "SHIFT", "SHIFT START", "SHIFT FINISH", "SHIFT HOURS")
  oWorksheet.[A1:J1].Font.Bold = True
  'get default only once (skip resources)
  If oCalendar.ResourceGuid = "00000000-0000-0000-0000-000000000000" Then
    Application.StatusBar = "Processing Calendar '" & oCalendar.Name & "' WorkWeek [Default]..."
    DoEvents
    For lngWeekDay = 1 To oCalendar.WeekDays.Count
      lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
      oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
      oWorksheet.Cells(lngLastRow, 2) = "[Default]"
      oWorksheet.Cells(lngLastRow, 3) = "-"
      oWorksheet.Cells(lngLastRow, 4) = "-"
      oWorksheet.Cells(lngLastRow, 5) = oCalendar.WeekDays(lngWeekDay).Name
      oWorksheet.Cells(lngLastRow, 6) = oCalendar.WeekDays(lngWeekDay).Working
      If oCalendar.WeekDays(lngWeekDay).Working Then
        If oCalendar.WeekDays(lngWeekDay).Shift1.Start > 0 Then
          oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
          oWorksheet.Cells(lngLastRow, 2) = "[Default]"
          oWorksheet.Cells(lngLastRow, 3) = "-"
          oWorksheet.Cells(lngLastRow, 4) = "-"
          oWorksheet.Cells(lngLastRow, 5) = oCalendar.WeekDays(lngWeekDay).Name
          oWorksheet.Cells(lngLastRow, 6) = oCalendar.WeekDays(lngWeekDay).Working
          oWorksheet.Cells(lngLastRow, 7) = 1
          oWorksheet.Cells(lngLastRow, 8) = oCalendar.WeekDays(lngWeekDay).Shift1.Start
          oWorksheet.Cells(lngLastRow, 9) = oCalendar.WeekDays(lngWeekDay).Shift1.Finish
          oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
        End If
        If oCalendar.WeekDays(lngWeekDay).Shift2.Start > 0 Then
          lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
          oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
          oWorksheet.Cells(lngLastRow, 2) = "[Default]"
          oWorksheet.Cells(lngLastRow, 3) = "-"
          oWorksheet.Cells(lngLastRow, 4) = "-"
          oWorksheet.Cells(lngLastRow, 5) = oCalendar.WeekDays(lngWeekDay).Name
          oWorksheet.Cells(lngLastRow, 6) = oCalendar.WeekDays(lngWeekDay).Working
          oWorksheet.Cells(lngLastRow, 7) = 2
          oWorksheet.Cells(lngLastRow, 8) = oCalendar.WeekDays(lngWeekDay).Shift2.Start
          oWorksheet.Cells(lngLastRow, 9) = oCalendar.WeekDays(lngWeekDay).Shift2.Finish
          oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
        End If
        If oCalendar.WeekDays(lngWeekDay).Shift3.Start > 0 Then
          lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
          oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
          oWorksheet.Cells(lngLastRow, 2) = "[Default]"
          oWorksheet.Cells(lngLastRow, 3) = "-"
          oWorksheet.Cells(lngLastRow, 4) = "-"
          oWorksheet.Cells(lngLastRow, 5) = oCalendar.WeekDays(lngWeekDay).Name
          oWorksheet.Cells(lngLastRow, 6) = oCalendar.WeekDays(lngWeekDay).Working
          oWorksheet.Cells(lngLastRow, 7) = 3
          oWorksheet.Cells(lngLastRow, 8) = oCalendar.WeekDays(lngWeekDay).Shift3.Start
          oWorksheet.Cells(lngLastRow, 9) = oCalendar.WeekDays(lngWeekDay).Shift3.Finish
          oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
        End If
        If oCalendar.WeekDays(lngWeekDay).Shift4.Start > 0 Then
          lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
          oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
          oWorksheet.Cells(lngLastRow, 2) = "[Default]"
          oWorksheet.Cells(lngLastRow, 3) = "-"
          oWorksheet.Cells(lngLastRow, 4) = "-"
          oWorksheet.Cells(lngLastRow, 5) = oCalendar.WeekDays(lngWeekDay).Name
          oWorksheet.Cells(lngLastRow, 6) = oCalendar.WeekDays(lngWeekDay).Working
          oWorksheet.Cells(lngLastRow, 7) = 4
          oWorksheet.Cells(lngLastRow, 8) = oCalendar.WeekDays(lngWeekDay).Shift4.Start
          oWorksheet.Cells(lngLastRow, 9) = oCalendar.WeekDays(lngWeekDay).Shift4.Finish
          oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
        End If
        If oCalendar.WeekDays(lngWeekDay).Shift5.Start > 0 Then
          lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
          oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
          oWorksheet.Cells(lngLastRow, 2) = "[Default]"
          oWorksheet.Cells(lngLastRow, 3) = "-"
          oWorksheet.Cells(lngLastRow, 4) = "-"
          oWorksheet.Cells(lngLastRow, 5) = oCalendar.WeekDays(lngWeekDay).Name
          oWorksheet.Cells(lngLastRow, 6) = oCalendar.WeekDays(lngWeekDay).Working
          oWorksheet.Cells(lngLastRow, 7) = 5
          oWorksheet.Cells(lngLastRow, 8) = oCalendar.WeekDays(lngWeekDay).Shift5.Start
          oWorksheet.Cells(lngLastRow, 9) = oCalendar.WeekDays(lngWeekDay).Shift5.Finish
          oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
        End If
      Else
        
      End If
    Next lngWeekDay
  End If
  If oCalendar.WorkWeeks.Count > 0 Then
    For Each oWorkWeek In oCalendar.WorkWeeks
      Application.StatusBar = "Processing Calendar '" & oCalendar.Name & "' WorkWeek '" & oWorkWeek.Name & "'..."
      DoEvents
      For Each oWeekDay In oWorkWeek.WeekDays
        lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
        oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
        oWorksheet.Cells(lngLastRow, 2) = oWorkWeek.Name
        oWorksheet.Cells(lngLastRow, 3) = oWorkWeek.Start
        oWorksheet.Cells(lngLastRow, 4) = oWorkWeek.Finish
        oWorksheet.Cells(lngLastRow, 5) = oWeekDay.Name
        oWorksheet.Cells(lngLastRow, 6) = oWeekDay.Working
        If oWeekDay.Working Then
          If oWeekDay.Shift1.Start > 0 Then
            oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
            oWorksheet.Cells(lngLastRow, 2) = oWorkWeek.Name
            oWorksheet.Cells(lngLastRow, 3) = oWorkWeek.Start
            oWorksheet.Cells(lngLastRow, 4) = oWorkWeek.Finish
            oWorksheet.Cells(lngLastRow, 5) = oWeekDay.Name
            oWorksheet.Cells(lngLastRow, 6) = oWeekDay.Working
            oWorksheet.Cells(lngLastRow, 7) = 1
            oWorksheet.Cells(lngLastRow, 8) = oWeekDay.Shift1.Start
            oWorksheet.Cells(lngLastRow, 9) = oWeekDay.Shift1.Finish
            oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
          End If
          If oWeekDay.Shift2.Start > 0 Then
            lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
            oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
            oWorksheet.Cells(lngLastRow, 2) = oWorkWeek.Name
            oWorksheet.Cells(lngLastRow, 3) = oWorkWeek.Start
            oWorksheet.Cells(lngLastRow, 4) = oWorkWeek.Finish
            oWorksheet.Cells(lngLastRow, 5) = oWeekDay.Name
            oWorksheet.Cells(lngLastRow, 6) = oWeekDay.Working
            oWorksheet.Cells(lngLastRow, 7) = 2
            oWorksheet.Cells(lngLastRow, 8) = oWeekDay.Shift2.Start
            oWorksheet.Cells(lngLastRow, 9) = oWeekDay.Shift2.Finish
            oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
          End If
          If oWeekDay.Shift3.Start > 0 Then
            lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
            oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
            oWorksheet.Cells(lngLastRow, 2) = oWorkWeek.Name
            oWorksheet.Cells(lngLastRow, 3) = oWorkWeek.Start
            oWorksheet.Cells(lngLastRow, 4) = oWorkWeek.Finish
            oWorksheet.Cells(lngLastRow, 5) = oWeekDay.Name
            oWorksheet.Cells(lngLastRow, 6) = oWeekDay.Working
            oWorksheet.Cells(lngLastRow, 7) = 3
            oWorksheet.Cells(lngLastRow, 8) = oWeekDay.Shift3.Start
            oWorksheet.Cells(lngLastRow, 9) = oWeekDay.Shift3.Finish
            oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
          End If
          If oWeekDay.Shift4.Start > 0 Then
            lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
            oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
            oWorksheet.Cells(lngLastRow, 2) = oWorkWeek.Name
            oWorksheet.Cells(lngLastRow, 3) = oWorkWeek.Start
            oWorksheet.Cells(lngLastRow, 4) = oWorkWeek.Finish
            oWorksheet.Cells(lngLastRow, 5) = oWeekDay.Name
            oWorksheet.Cells(lngLastRow, 6) = oWeekDay.Working
            oWorksheet.Cells(lngLastRow, 7) = 4
            oWorksheet.Cells(lngLastRow, 8) = oWeekDay.Shift4.Start
            oWorksheet.Cells(lngLastRow, 9) = oWeekDay.Shift4.Finish
            oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
          End If
          If oWeekDay.Shift5.Start > 0 Then
            lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
            oWorksheet.Cells(lngLastRow, 1) = oCalendar.Name
            oWorksheet.Cells(lngLastRow, 2) = oWorkWeek.Name
            oWorksheet.Cells(lngLastRow, 3) = oWorkWeek.Start
            oWorksheet.Cells(lngLastRow, 4) = oWorkWeek.Finish
            oWorksheet.Cells(lngLastRow, 5) = oWeekDay.Name
            oWorksheet.Cells(lngLastRow, 6) = oWeekDay.Working
            oWorksheet.Cells(lngLastRow, 7) = 5
            oWorksheet.Cells(lngLastRow, 8) = oWeekDay.Shift5.Start
            oWorksheet.Cells(lngLastRow, 9) = oWeekDay.Shift5.Finish
            oWorksheet.Cells(lngLastRow, 10).FormulaR1C1 = "=IF(RC[-1]<RC[-2],24-HOUR(RC[-2])-HOUR(RC[-1]),IF(RC[-2]=RC[-1],24,HOUR(RC[-1])-HOUR(RC[-2])))"
          End If
        Else
          'oWorksheet.Cells(lngLastRow, 7) = 0
        End If
      Next oWeekDay
    Next oWorkWeek
  End If
  
exit_here:
  On Error Resume Next
  Set oExcel = Nothing
  Set oWorksheet = Nothing
  cptSpeed False
  Set oWeekDay = Nothing
  Set oWorkWeek = Nothing
  Set oException = Nothing
  Set oCalendar = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCalendarExceptions_bas", "cptExportCalendarExceptions", Err, Erl)
  Resume exit_here
End Sub

Sub cptFormatExceptionDetail(ByRef rng As Excel.Range)
  rng.IndentLevel = 1
  rng.EntireRow.OutlineLevel = 2
  Set rng = rng.Resize(, rng.Worksheet.[A1].End(xlToRight).Column - 1)
  With rng.Font
    .Italic = True
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0.499984740745262
  End With
End Sub

Function cptGetShifts(ByRef oException As MSProject.Exception) As Double
  'objects
  'strings
  'longs
  'integers
  'doubles
  Dim dblShift1 As Double
  Dim dblShift2 As Double
  Dim dblShift3 As Double
  Dim dblShift4 As Double
  Dim dblShift5 As Double
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  With oException
    dblShift1 = Hour(.Shift1.Finish) - Hour(.Shift1.Start)
    dblShift2 = Hour(.Shift2.Finish) - Hour(.Shift2.Start)
    dblShift3 = Hour(.Shift3.Finish) - Hour(.Shift3.Start)
    dblShift4 = Hour(.Shift4.Finish) - Hour(.Shift4.Start)
    dblShift5 = Hour(.Shift5.Finish) - Hour(.Shift5.Start)
  End With

  cptGetShifts = dblShift1 + dblShift2 + dblShift3 + dblShift4 + dblShift5

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCalendarExceptions_bas", "cptGetShifts", Err, Erl)
  Resume exit_here
End Function

Sub cptCalendarCompareMain()
  'objects
  Dim oListObject As Excel.ListObject
  Dim oRecordset As ADODB.Recordset
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  Dim oMaster As MSProject.Project
  Dim oException As MSProject.Exception
  Dim oSubproject As MSProject.SubProject
  Dim oCalendar As MSProject.Calendar
  'strings
  'longs
  Dim lngMismatched As Long
  Dim lngLastRow As Long
  Dim lngLastCol As Long
  Dim lngItems As Long
  Dim lngItem As Long
  Dim lngSubproject As Long
  Dim lngSubprojects As Long
  Dim lngCalendar As Long
  Dim lngCalendars As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  If ActiveProject.Subprojects.Count = 0 Then
    MsgBox "This feature is meant for Master/Sub scenarios.", vbExclamation + vbOKOnly, "Calendar Compare"
    GoTo exit_here
  End If
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptSpeed True
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Fields.Append "PROJECT", adVarChar, 255
  oRecordset.Fields.Append "CALENDAR", adVarChar, 255
  oRecordset.Fields.Append "EXCEPTION", adVarChar, 255
  oRecordset.Fields.Append "DATE", adDate
  oRecordset.Open
  
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    
  'first export the master
  Set oMaster = ActiveProject
  Application.StatusBar = "Exporting Master project..."
  oRecordset.AddNew Array(0), Array(oMaster.Name)
  lngCalendars = oMaster.BaseCalendars.Count
  lngCalendar = 0
  For Each oCalendar In oMaster.BaseCalendars
    If oCalendar.Name <> "cptFiscalCalendar" Then
      For Each oException In oCalendar.Exceptions
        If oException.Occurrences > 1 Then
          cptCalendarCompareDetail oRecordset, oExcel, oException, oMaster.Name
        Else
          oRecordset.AddNew Array(0, 1, 2, 3), Array(oMaster.Name, oException.Parent.Name, oException.Name, oException.Start)
        End If
      Next oException
    End If
    Application.StatusBar = "Exporting Master project | Calendar " & lngCalendar & " of " & lngCalendars & " (" & Format(lngCalendar / lngCalendars, "0%") & ")"
    DoEvents
  Next oCalendar
  'then export the subs
  lngSubprojects = oMaster.Subprojects.Count
  lngSubproject = 0
  For Each oSubproject In oMaster.Subprojects
    lngSubproject = lngSubproject + 1
    Application.StatusBar = "Exporting Subproject " & lngSubproject & " of " & lngSubprojects & ")..."
    oRecordset.AddNew Array(0), Array(oSubproject.SourceProject.Name)
    lngCalendars = oSubproject.SourceProject.BaseCalendars.Count
    lngCalendar = 0
    For Each oCalendar In oSubproject.SourceProject.BaseCalendars
      lngCalendar = lngCalendar + 1
      Application.StatusBar = "Exporting Subproject " & lngSubproject & " of " & lngSubprojects & " | Calendar " & lngCalendar & " of " & lngCalendars & " (" & Format(lngCalendar / lngCalendars, "0%") & ")"
      If oCalendar.Name <> "cptFiscalCalendar" Then
        For Each oException In oCalendar.Exceptions
          If oException.Occurrences > 1 Then
            cptCalendarCompareDetail oRecordset, oExcel, oException, oSubproject.SourceProject.Name
          Else
            oRecordset.AddNew Array(0, 1, 2, 3), Array(oSubproject.SourceProject.Name, oException.Parent.Name, oException.Name, oException.Start)
          End If
        Next oException
      End If
    Next oCalendar
  Next oSubproject
  
  If oRecordset.RecordCount > 0 Then
    Application.StatusBar = "Creating report..."
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.Name = "DATA"
    For lngItem = 0 To oRecordset.Fields.Count - 1
      oWorksheet.Cells(1, lngItem + 1) = oRecordset.Fields(lngItem).Name
    Next lngItem
    oWorksheet.[A2].CopyFromRecordset oRecordset
    
    Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)), , xlYes)
    oListObject.Name = "DATA"
    oListObject.TableStyle = ""
    
    oExcel.ActiveWindow.Zoom = 85
    oExcel.ActiveWindow.DisplayGridlines = False
    oExcel.ActiveWindow.SplitRow = 1
    oExcel.ActiveWindow.SplitColumn = 0
    oExcel.ActiveWindow.FreezePanes = True
    cptAddBorders oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown))
    cptAddBorders oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight))
    oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True
    oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)).EntireColumn.AutoFit
    cptAddShading oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)), True
    
    Set oWorksheet = oWorkbook.Worksheets.Add(oWorkbook.Sheets("Data"))
    oWorksheet.Name = "Compare"
    oWorksheet.[A1:C1] = Split("CALENDAR,EXCEPTION,DATE", ",")
    If [COUNTA(DATA[EXCEPTION])] > 0 Then
      oWorksheet.[A2].Formula2 = "=UNIQUE(CHOOSECOLS(FILTER(DATA,LEN(DATA[EXCEPTION])>0),2,3,4))"
    Else
      oWorksheet.[A2] = "No Exceptions"
    End If
    oWorksheet.[D1].Formula2 = "=TOROW(SORT(UNIQUE(DATA[PROJECT])))"
    lngLastRow = oWorksheet.[A1].End(xlDown).Row
    lngLastCol = oWorksheet.[A1].End(xlToRight).Column
    oWorksheet.Range(oWorksheet.Cells(2, 3), oWorksheet.Cells(lngLastRow, 3)).NumberFormat = "m/d/yyyy"
    If oWorksheet.[A2] = "No Exceptions" Then
      oWorksheet.Range(oWorksheet.Cells(2, 4), oWorksheet.Cells(lngLastRow, lngLastCol)).FormulaR1C1 = 2
    Else
      oWorksheet.Range(oWorksheet.Cells(2, 4), oWorksheet.Cells(lngLastRow, lngLastCol)).FormulaR1C1 = "=IF(COUNTIFS(DATA[CALENDAR],RC1,DATA[EXCEPTION],RC2,DATA[DATE],RC3,DATA[PROJECT],R1C)>=1,2,0)"
    End If
    With oWorksheet.Range(oWorksheet.Cells(2, 4), oWorksheet.Cells(lngLastRow, lngLastCol))
      .FormatConditions.AddIconSetCondition
      .FormatConditions(.FormatConditions.Count).SetFirstPriority
      With .FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = True
        .IconSet = oWorkbook.IconSets(xl3Symbols)
      End With
      With .FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 5
      End With
      With .FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 2
        .Operator = 7
      End With
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
    End With
    
    oWorksheet.Range(oWorksheet.[A1], oWorksheet.[C2].End(xlDown)).EntireColumn.AutoFit
    
    With oWorksheet.Range(oWorksheet.Cells(1, 4), oWorksheet.Cells(1, lngLastCol))
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 45
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
    End With
    
    oExcel.ActiveWindow.Zoom = 85
    oExcel.ActiveWindow.DisplayGridlines = False
    oExcel.ActiveWindow.SplitRow = 1
    oExcel.ActiveWindow.SplitColumn = 0
    oExcel.ActiveWindow.FreezePanes = True
    cptAddBorders oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown))
    cptAddBorders oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight))
    oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True
    oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)).EntireColumn.AutoFit
    cptAddShading oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)), True
    
    'count mismatches & notify
    oWorksheet.Range(oWorksheet.Cells(2, lngLastCol + 2), oWorksheet.Cells(lngLastRow, lngLastCol + 2)).FormulaR1C1 = "=IF(COUNTIF(RC3:RC[-2],0)>0,1,0)"
    lngMismatched = oExcel.WorksheetFunction.Sum(oWorksheet.Range(oWorksheet.Cells(2, lngLastCol + 2), oWorksheet.Cells(lngLastRow, lngLastCol + 2)))
    oWorksheet.Range(oWorksheet.Cells(2, lngLastCol + 2), oWorksheet.Cells(lngLastRow, lngLastCol + 2)).ClearContents
    Application.StatusBar = Format(lngMismatched, "#,##0") & " mismatched exception" & IIf(lngMismatched <> 1, "s", "") & "."
    MsgBox Format(lngMismatched, "#,##0") & " mismatched exception" & IIf(lngMismatched <> 1, "s", "") & ".", vbInformation + vbOKOnly, "Calendar Compare"
    oExcel.Visible = True
    oExcel.ScreenUpdating = True
  Else
    MsgBox "No Calendar Exceptions found.", vbInformation + vbOKOnly, "Calendar Compare"
  End If
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  cptSpeed False
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oMaster = Nothing
  Set oException = Nothing
  Set oSubproject = Nothing
  Set oCalendar = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCalendarExceptions_bas", "cptCalendarCompareMain", Err, Erl)
  Resume exit_here
End Sub

Private Sub cptCalendarCompareDetail(ByRef oRecordset As ADODB.Recordset, ByRef oExcel As Excel.Application, ByRef oException As MSProject.Exception, strProjectName As String)
  'objects
  'strings
  Dim strExceptionName As String
  Dim strDaysOfWeek As String
  'longs
  Dim lngDay As Long
  Dim lngDaysOfWeek As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vDayOfWeek As Variant
  'dates
  Dim dtDate As Date
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  With oException
    If Len(oException.Name) = 0 Then strExceptionName = "[unnamed]" Else strExceptionName = oException.Name
    Select Case .Type
      Case pjDaily '1
        dtDate = .Start
        Do While dtDate < .Finish
          dtDate = DateAdd("d", 1, dtDate)
          oRecordset.AddNew Array(0, 1, 2, 3), Array(strProjectName, oException.Parent.Name, strExceptionName, dtDate)
        Loop
      Case pjYearlyMonthDay '2
        If .Start <> DateValue(.Month & "/" & .MonthDay & "/" & Year(.Start)) Then
          dtDate = CDate(DateValue(.Month & "/" & .MonthDay & "/" & Year(.Start)))
        Else
          dtDate = .Start
        End If
        Do While dtDate <= .Finish
          dtDate = DateAdd("yyyy", 1, dtDate)
          If dtDate <= .Finish Then
            oRecordset.AddNew Array(0, 1, 2, 3), Array(strProjectName, oException.Parent.Name, strExceptionName, dtDate)
          End If
        Loop
      Case pjYearlyPositional '3
        dtDate = .Start
        Do While dtDate < .Finish
          dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, 11) + 1
          'find first .MonthItem
          If Weekday(dtDate) <= .MonthItem - 2 Then
            dtDate = DateAdd("d", .MonthItem - 2 - Weekday(dtDate), dtDate)
          Else
            dtDate = DateAdd("d", 7 + .MonthItem - 2 - Weekday(dtDate), dtDate)
          End If
          'adjust for position
          If .MonthPosition < pjLast Then
            dtDate = DateAdd("d", 7 * .MonthPosition, dtDate)
          ElseIf .MonthPosition = pjLast Then
            'find last day of month
            dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, 0)
            If Weekday(dtDate) <> (.MonthItem - 2) Then
              For lngDay = Day(dtDate) To 1 Step -1
                If Weekday(dtDate) = .MonthItem - 2 Then
                  'bingo
                  Exit For
                Else
                  dtDate = DateAdd("d", -1, dtDate)
                End If
              Next lngDay
            End If
          End If
          If dtDate <= .Finish Then
            oRecordset.AddNew Array(0, 1, 2, 3), Array(strProjectName, oException.Parent.Name, strExceptionName, dtDate)
          End If
        Loop
      Case pjMonthlyMonthDay '4
        dtDate = .Start
        Do While dtDate <= .Finish
          dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, .Period - 1) + .MonthDay
          If dtDate <= .Finish Then
            oRecordset.AddNew Array(0, 1, 2, 3), Array(strProjectName, oException.Parent.Name, strExceptionName, dtDate)
          End If
        Loop
      Case pjMonthlyPositional '5
        dtDate = .Start
        Do While dtDate <= .Finish
          'find first day of next month
          dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, .Period - 1) + 1
          'find first .MonthItem
          If Weekday(dtDate) <= .MonthItem - 2 Then
            dtDate = DateAdd("d", .MonthItem - 2 - Weekday(dtDate), dtDate)
          Else
            dtDate = DateAdd("d", 7 + .MonthItem - 2 - Weekday(dtDate), dtDate)
          End If
          'adjust for position
          If .MonthPosition < pjLast Then
            dtDate = DateAdd("d", 7 * .MonthPosition, dtDate)
          Else
            'find last day of month
            dtDate = oExcel.WorksheetFunction.EoMonth(dtDate, 0)
            If Weekday(dtDate) <> (.MonthItem - 2) Then
              For lngDay = Day(dtDate) To 1 Step -1
                If Weekday(dtDate) = .MonthItem - 2 Then
                  'bingo
                  Exit For
                Else
                  dtDate = DateAdd("d", -1, dtDate)
                End If
              Next lngDay
            End If
          End If
          If dtDate <= .Finish Then
            oRecordset.AddNew Array(0, 1, 2, 3), Array(strProjectName, oException.Parent.Name, strExceptionName, dtDate)
          End If
        Loop
      Case pjWeekly '6
        If .DaysOfWeek > 0 Then
          lngDaysOfWeek = .DaysOfWeek
          strDaysOfWeek = ""
          If lngDaysOfWeek >= 64 Then
            strDaysOfWeek = "7,"
            lngDaysOfWeek = lngDaysOfWeek - 64
          End If
          If lngDaysOfWeek >= 32 Then
            strDaysOfWeek = "6," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 32
          End If
          If lngDaysOfWeek >= 16 Then
            strDaysOfWeek = "5," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 16
          End If
          If lngDaysOfWeek >= 8 Then
            strDaysOfWeek = "4," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 8
          End If
          If lngDaysOfWeek >= 4 Then
            strDaysOfWeek = "3," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 4
          End If
          If lngDaysOfWeek >= 2 Then
            strDaysOfWeek = "2," & strDaysOfWeek
            lngDaysOfWeek = lngDaysOfWeek - 2
          End If
          If lngDaysOfWeek >= 1 Then
            strDaysOfWeek = "1," & strDaysOfWeek
          End If
          For Each vDayOfWeek In Split(strDaysOfWeek, ",")
            If Len(vDayOfWeek) = 0 Then Exit For
            If Weekday(.Start) = CLng(vDayOfWeek) Then
              dtDate = .Start
            Else
              dtDate = DateAdd("d", CLng(vDayOfWeek) - Weekday(.Start), .Start)
              oRecordset.AddNew Array(0, 1, 2, 3), Array(strProjectName, oException.Parent.Name, strExceptionName, dtDate)
            End If
            Do While dtDate <= .Finish
              dtDate = DateAdd("d", 7 * .Period, dtDate)
              If dtDate <= .Finish Then
                oRecordset.AddNew Array(0, 1, 2, 3), Array(strProjectName, oException.Parent.Name, strExceptionName, dtDate)
              End If
            Loop
          Next vDayOfWeek
        End If 'daysofweek>0
      End Select 'case .Type
    End With
    
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptCalendarExceptions_bas", "cptCalendarCompareDetail", Err, Erl)
  Resume exit_here
End Sub

