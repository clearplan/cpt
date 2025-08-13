Attribute VB_Name = "cptTaskHistory_bas"
'<cpt_version>v1.1.2</cpt_version>
Option Explicit
Public oTaskHistory As ADODB.Recordset

Sub cptShowTaskHistory_frm()
  'objects
  Dim myTaskHistory_frm As cptTaskHistory_frm
  'strings
  Dim strProgramAcronym As String
  Dim strFile As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnFieldExists As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'ensure program acroynm
  strProgramAcronym = cptGetProgramAcronym
  
  'ensure file exists
  strFile = cptDir & "\settings\cpt-cei.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox "Please run 'capture week' for each week's file that you want recorded.", vbCritical + vbOKOnly, "File Not Found"
    GoTo exit_here
  End If
  Set oTaskHistory = CreateObject("ADODB.Recordset")
  oTaskHistory.Open strFile
  oTaskHistory.MoveFirst
  
  'ensure NOTE field exists
  blnFieldExists = True
  On Error Resume Next
  Debug.Print oTaskHistory.Fields("NOTE") 'keep
  If Err.Number = 3265 Then blnFieldExists = False
  If Not blnFieldExists Then
    oTaskHistory.Close
    cptAppendColumn strFile, "NOTE", 203, 500 '203=adLongVarWChar
    oTaskHistory.Open strFile
  End If
  Set myTaskHistory_frm = New cptTaskHistory_frm
  cptUpdateTaskHistory myTaskHistory_frm
  With myTaskHistory_frm
    .Caption = "Task History (" & cptGetVersion("cptTaskHistory_bas") & ")"
    .Show False
  End With
  
exit_here:
  On Error Resume Next
  Set myTaskHistory_frm = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHistory_bas", "cptShowTaskHistory_frm", Err, Erl)
  Resume exit_here

End Sub

Sub cptUpdateTaskHistory(ByRef myTaskHistory_frm As cptTaskHistory_frm)
  'objects
  Dim oTasks As Object
  'strings
  Dim strNote As String
  Dim strFile As String
  Dim strProgramAcronym As String
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtFinish As Date
  Dim dtStart As Date
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  myTaskHistory_frm.lblWarning.Visible = False
  If oTasks Is Nothing Then
    myTaskHistory_frm.lngTaskHistoryUID = 0
    myTaskHistory_frm.lboTaskHistory.Clear
    myTaskHistory_frm.txtVariance = ""
    myTaskHistory_frm.lblWarning.Caption = "Nothing selected"
    myTaskHistory_frm.lblWarning.Visible = True
    GoTo exit_here
  End If
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strProgramAcronym = cptGetProgramAcronym
  
  With myTaskHistory_frm
    .lboHeader.Clear
    .lboHeader.ColumnCount = 7
    .lboHeader.AddItem
    .lboHeader.List(0, 0) = "STATUS_DATE"
    .lboHeader.List(0, 1) = "STATUS DATE"
    .lboHeader.List(0, 2) = "START"
    .lboHeader.List(0, 3) = "DUR"
    .lboHeader.List(0, 4) = "FINISH"
    .lboHeader.List(0, 5) = "RDur"
    '.lboHeader.List(0, 6) = "RWork"
    .lboHeader.List(0, 6) = "STATUS NOTE"
    .lboTaskHistory.Clear
    .lboTaskHistory.ColumnCount = 7
    .txtVariance = ""
    .lblWarning.Visible = False
    If oTasks.Count <> 1 Then
      .lngTaskHistoryUID = 0
      .tglExport.Enabled = False
      If oTasks.Count > 1 Then
        .lboTaskHistory.ColumnCount = 2
        .lboTaskHistory.AddItem
        .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 1) = "One at a time, please..."
      Else
        .lboTaskHistory.Clear
      End If
    Else
      If .lngTaskHistoryUID > 0 Then .tglExport.Enabled = True
      If oTasks(1).Summary Then
        .lblWarning.Caption = "History not captured on summary tasks."
        .lblWarning.Visible = True
        GoTo exit_here
      ElseIf Not oTasks(1).Active Then
        .lblWarning.Caption = "History not captured on inactive tasks."
        .lblWarning.Visible = True
        GoTo exit_here
      ElseIf oTasks(1).ExternalTask Then
        .lblWarning.Caption = "History not captured on external tasks."
        .lblWarning.Visible = True
        GoTo exit_here
      End If
      .lngTaskHistoryUID = oTasks(1).UniqueID
      lngUID = .lngTaskHistoryUID
      If oTaskHistory.RecordCount > 0 Then
        oTaskHistory.Filter = "PROJECT='" & strProgramAcronym & "' AND TASK_UID=" & CLng(lngUID)
        oTaskHistory.Sort = "STATUS_DATE desc"
        If oTaskHistory.EOF Then
          .lblWarning.Caption = "No history for UID " & lngUID & "."
          .lblWarning.Visible = True
          GoTo exit_here
        Else
          oTaskHistory.MoveFirst
          Do While Not oTaskHistory.EOF
            .lboTaskHistory.AddItem
            .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 0) = oTaskHistory("STATUS_DATE")
            .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 1) = FormatDateTime(oTaskHistory("STATUS_DATE"), vbShortDate) & IIf(Len(oTaskHistory("NOTE")) > 0, "*", "")
            If oTaskHistory("TASK_AS") > 0 Then
              dtStart = oTaskHistory("TASK_AS")
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 2) = "[" & FormatDateTime(dtStart, vbShortDate) & "]"
            Else
              dtStart = oTaskHistory("TASK_START")
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 2) = FormatDateTime(dtStart, vbShortDate)
            End If
          
            If oTaskHistory("TASK_AF") > 0 Then
              dtFinish = oTaskHistory("TASK_AF")
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 4) = "[" & FormatDateTime(dtFinish, vbShortDate) & "]"
            Else
              dtFinish = oTaskHistory("TASK_FINISH")
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 4) = FormatDateTime(dtFinish, vbShortDate)
            End If
            .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 3) = Application.DateDifference(dtStart, dtFinish) / 480 & "d"
            .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 5) = oTaskHistory("TASK_RD") & "d"
            'todo: add RWork to cpt-cei.adtg and cptCaptureWeek
            'todo: add EVP to cpt-cei.adtg and cptCaptureWeek
            'todo: add BAC to cpt-cei.adtg and cptCaptureWeek?
            'todo: remove NOTE from lboTaskHistory
            strNote = oTaskHistory("NOTE")
            If Len(strNote) = 0 Then
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 6) = ""
            ElseIf Len(strNote) > 20 Then
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 6) = Left(strNote, 17) & "..."
            Else
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 6) = strNote
            End If
            If Len(strNote) > 0 Then
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 1) = FormatDateTime(oTaskHistory("STATUS_DATE"), vbShortDate) & "*"
            Else
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 1) = FormatDateTime(oTaskHistory("STATUS_DATE"), vbShortDate)
            End If
            oTaskHistory.MoveNext
          Loop
          oTaskHistory.Filter = 0
        End If
      Else
        .lblWarning.Caption = "No records found."
        .lblWarning.Visible = True
      End If
    End If
  End With

exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  oTaskHistory.Filter = 0

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHistory_bas", "cptUpdateTaskHistory", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateTaskHistoryNote(ByRef myTaskHistory_frm As cptTaskHistory_frm, lngUID As Long, dtStatus As Date, strVariance As String)
  'objects
  'strings
  Dim strFile As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnFieldExists As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  oTaskHistory.MoveFirst
  oTaskHistory.Filter = "PROJECT='" & cptGetProgramAcronym & "' AND TASK_UID=" & CLng(lngUID) & " AND STATUS_DATE=#" & FormatDateTime(dtStatus, vbGeneralDate) & "#"
  If Not oTaskHistory.EOF Then
    oTaskHistory.Update Array("NOTE"), Array(strVariance)
  Else
    'todo: do what if Filter > .EOF=TRUE
  End If
  
  With myTaskHistory_frm
    If Len(strVariance) = 0 Then
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 6) = ""
    ElseIf Len(strVariance) > 10 Then
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 6) = Left(strVariance, 17) & "..."
    Else
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 6) = strVariance
    End If
    If Len(strVariance) > 0 Then
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 1) = FormatDateTime(dtStatus, vbShortDate) & "*"
    Else
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 1) = FormatDateTime(dtStatus, vbShortDate)
    End If
  End With
  
exit_here:
  On Error Resume Next
  oTaskHistory.Filter = 0
  
  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHistory_bas", "cptUpdateTaskHistoryNote", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetTaskHistoryNote(ByRef myTaskHistory_frm As cptTaskHistory_frm, dtStatus As Date, lngUID As Long)
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  With myTaskHistory_frm
    .txtVariance = ""
    oTaskHistory.MoveFirst
    oTaskHistory.Filter = "PROJECT='" & cptGetProgramAcronym & "' AND TASK_UID=" & lngUID & " AND STATUS_DATE=#" & FormatDateTime(dtStatus, vbGeneralDate) & "#"
    If Not oTaskHistory.EOF Then
      .txtVariance = oTaskHistory("NOTE")
    End If
  End With
    
exit_here:
  On Error Resume Next
  oTaskHistory.Filter = 0

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHistory_bas", "cptGetTaskHistoryNote", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportTaskHistory(ByRef myTaskHistory_frm As cptTaskHistory_frm, Optional lngUID As Long, Optional blnNotesOnly As Boolean = False)
  'objects
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  'strings
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vRow As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Application.StatusBar = "Getting data..."
  oTaskHistory.MoveFirst
  If lngUID > 0 Then
    Application.StatusBar = "Filtering for UID " & lngUID & "..."
    oTaskHistory.Filter = "PROJECT='" & cptGetProgramAcronym & "' AND TASK_UID=" & lngUID
  ElseIf blnNotesOnly Then
    Application.StatusBar = "Fitlering for notes in the current period..."
    oTaskHistory.Filter = "PROJECT='" & cptGetProgramAcronym & "' AND STATUS_DATE=#" & FormatDateTime(ActiveProject.StatusDate, vbGeneralDate) & "# AND NOTE <>''"
  Else
    Application.StatusBar = "Filtering..."
    oTaskHistory.Filter = "PROJECT='" & cptGetProgramAcronym & "'"
  End If
  Application.StatusBar = "Sorting by STATUS_DATE descending..."
  oTaskHistory.Sort = "STATUS_DATE desc"
  If Not oTaskHistory.EOF Then
    Application.StatusBar = "Getting Excel..."
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oExcel Is Nothing Then
      Set oExcel = CreateObject("Excel.Application")
    End If
    oExcel.ScreenUpdating = False
    Application.StatusBar = "Setting up workbook..."
    Set oWorkbook = oExcel.Workbooks.Add
    oExcel.Calculation = xlCalculationManual
    Set oWorksheet = oWorkbook.Sheets(1)
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.Name = "Task History"
    oWorksheet.[A1].Value = ActiveProject.Name
    oWorksheet.[A1:E1].Merge
    Application.StatusBar = "Building header..."
    vRow = oWorksheet.Range(oWorksheet.[A3], oWorksheet.Cells(3, oTaskHistory.Fields.Count)).Value
    For lngItem = 0 To oTaskHistory.Fields.Count - 1
      vRow(1, lngItem + 1) = oTaskHistory.Fields(lngItem).Name
    Next lngItem
    oWorksheet.Range(oWorksheet.[A3], oWorksheet.Cells(3, oTaskHistory.Fields.Count)) = vRow
    Application.StatusBar = "Importing data..."
    oWorksheet.[A4].CopyFromRecordset cptConvertFilteredRecordset(oTaskHistory)
    Application.StatusBar = "Formatting..."
    oWorksheet.[A3].AutoFilter
    oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight)).Font.Bold = True
    oWorksheet.Columns.AutoFit
    oWorksheet.[A4].Select
    MsgBox "Export complete.", vbInformation + vbOKOnly, "Task History"
    oExcel.Visible = True
    Application.ActivateMicrosoftApp pjMicrosoftExcel
    oExcel.ScreenUpdating = True
    oExcel.ActiveWindow.FreezePanes = True
    Application.StatusBar = "...done."
  Else
    myTaskHistory_frm.lblWarning.Caption = "No records found."
    myTaskHistory_frm.lblWarning.Visible = True
  End If
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oWorksheet = Nothing
  oExcel.Calculation = xlCalculationAutomatic
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  oTaskHistory.Filter = 0

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHistory_bas", "cptExportTaskHistory", Err, Erl)
  Resume exit_here
End Sub

