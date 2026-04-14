Attribute VB_Name = "cptCriticalPathTools_bas"
'<cpt_version>v1.3.0</cpt_version>
Option Explicit
#If Win64 And VBA7 Then
  Declare PtrSafe Sub keybd_event Lib "user32" ( _
      ByVal bVk As Byte, _
      ByVal bScan As Byte, _
      ByVal dwFlags As Long, _
      ByVal dwExtraInfo As Long)
#Else
Declare Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long)
#End If

Const VK_PAGEDOWN As Long = &H22
Const KEYEVENTF_KEYUP As Long = &H2

Sub cptExportCriticalPath(ByRef oProject As MSProject.Project, Optional blnSendEmail As Boolean = False, Optional blnKeepOpen As Boolean = False, Optional ByRef oTargetTask As MSProject.Task)
  'objects
  Dim oDrivingPaths As Scripting.Dictionary
  Dim oShell As Object
  Dim pptExists As PowerPoint.Presentation
  Dim oTask As MSProject.Task
  Dim oTasks As MSProject.Tasks
  Dim oPowerPoint As PowerPoint.Application
  Dim oPresentation As PowerPoint.Presentation
  Dim oSlide As PowerPoint.Slide
  'Dim Shape As PowerPoint.Shape
  'Dim ShapeRange As PowerPoint.ShapeRange
  'strings
  Dim strTitle As String
  Dim strDrivingPaths As String
  Dim strFileName As String
  Dim strProjectName As String
  Dim strDir As String
  'longs
  Dim lngItem As Long
  Dim lngDrivingPath As Long
  Dim lngDrivingPathField As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngSlide As Long
  Dim lngFromRow As Long
  Dim lngToRow As Long
  'dates
  Dim dtFrom As Date
  Dim dtTo As Date
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vPath As Variant

  blnErrorTrapping = True
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not cptModuleExists("cptCriticalPath_bas") Then
    MsgBox "Please install the ClearPlan Critical Path Module.", vbCritical + vbOKOnly, "CP Toolbar"
    GoTo exit_here
  End If
  
  cptSpeed True
  
  export_to_PPT = True
  Call DrivingPaths
  export_to_PPT = False
  
  'get path count
  SelectAll
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then GoTo exit_here
  If oTasks.Count = 0 Then GoTo exit_here
  Set oDrivingPaths = CreateObject("Scripting.Dictionary")
  lngDrivingPathField = CLng(cptGetSetting("Driving Path Group", "GUID"))
  For Each oTask In oTasks
    If Not oTask.GroupBySummary Then
      lngDrivingPath = CLng(oTask.GetField(lngDrivingPathField))
      If lngDrivingPath > 0 Then
        If Not oDrivingPaths.Exists(lngDrivingPath) Then oDrivingPaths.Add lngDrivingPath, lngDrivingPath
      End If
    End If
  Next oTask
  For lngItem = 0 To oDrivingPaths.Count - 1
    strDrivingPaths = strDrivingPaths & oDrivingPaths.Items(lngItem) & ","
  Next lngItem
  If Right(strDrivingPaths, 1) = "," Then
    strDrivingPaths = Left(strDrivingPaths, Len(strDrivingPaths) - 1)
  End If
  Set oTasks = Nothing
  Set oDrivingPaths = Nothing
  
  're-select the target task
  Find "Unique ID", "equals", oTargetTask.UniqueID
  
  If Not IsDate(oProject.StatusDate) Then
    dtFrom = DateAdd("d", -14, oProject.ProjectStart)
  Else
    dtFrom = DateAdd("d", -14, oProject.StatusDate)
  End If
  dtTo = DateAdd("d", 30, oTargetTask.Finish)

  EditGoTo Date:=dtFrom
  
  Set oPowerPoint = CreateObject("PowerPoint.Application")
  oPowerPoint.Visible = True
  Set oPresentation = oPowerPoint.Presentations.Add(msoCTrue)
  
  'ensure directory
  Set oShell = CreateObject("WScript.Shell")
  strDir = oShell.SpecialFolders("Desktop") & "\"
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  'build filename
  strFileName = cptRegEx(ActiveProject.Name, "[^\\/]{1,}$")
  strFileName = Replace(strFileName, ".mpp", "")
  strFileName = Replace(strFileName, " ", "_")
  strFileName = strDir & cptGetProgramAcronym & "-DrivingPathAnalysis-" & Format(Now, "yyyy-mm-dd") & ".pptx"
  On Error Resume Next
  Set pptExists = oPowerPoint.Presentations(strFileName)
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not pptExists Is Nothing Then 'add timestamp to this file
    pptExists.Save
    pptExists.Close
  End If
  'might exist but be closed
  If Dir(strFileName) <> vbNullString Then
    If MsgBox("A file with this name already exists:" & vbCrLf & vbCrLf & strFileName & vbCrLf & vbCrLf & "OK to overwrite?", vbExclamation + vbYesNo, "File Exists") = vbYes Then
      Kill strFileName
    Else
      MsgBox "The presentation you are creating will have a time stamp in the filename to prevent overwriting.", vbInformation + vbOKOnly, "File Name Changed"
      strFileName = Replace(strFileName, ".mpp", "-" & Format(Now, "hh-nn-ss") & ".mpp")
    End If
  Else
    
  End If
  oPresentation.SaveAs strFileName
  'make a title slide
  Set oSlide = oPresentation.Slides.Add(1, ppLayoutCustom)
  oSlide.Layout = ppLayoutTitle
  strProjectName = Replace(cptRegEx(ActiveProject.Name, "[^\\/]{1,}$"), ".mpp", "")
  oSlide.Shapes(1).TextFrame.TextRange.Text = strProjectName & vbCrLf & "Driving Path Analysis"
  oSlide.Shapes(2).TextFrame.TextRange.Text = cptGetUserFullName & vbCrLf & FormatDateTime(Now, vbShortDate)
  
  'close timeline view / bottom pane if open
  On Error Resume Next
  ActiveWindow.BottomPane.Close
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  SelectTaskColumn "Name"
  WrapText
 
  'for each primary,secondary,tertiary > make a slide
  For Each vPath In Split(strDrivingPaths, ",")
    'copy the picture
    'SetAutoFilter FieldName:="CP Driving Paths", FilterType:=pjAutoFilterCustom, Test1:="contains", Criteria1:=CStr(vPath)
    SetAutoFilter FieldName:="CP Driving Path Group ID", FilterType:=pjAutoFilterIn, Criteria1:=CStr(vPath)
    Sort Key1:="Finish", Key2:="Duration", Ascending2:=False, Renumber:=False
    TimescaleEdit MajorUnits:=0, MinorUnits:=2, MajorLabel:=0, MinorLabel:=10, MinorTicks:=True, Separator:=True, TierCount:=2
    SelectBeginning
    SelectAll
    'account for when a path is somehow not found
    On Error Resume Next
    Set oTasks = ActiveSelection.Tasks
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oTasks Is Nothing Then GoTo next_path
    'account for when task count exceeds easily visible range on powerpoint slide
    'also account for very long task names (wraptext)
    ActiveWindow.Activate
    SelectBeginning
    lngSlide = 0
    Do
      ActiveWindow.Activate
      lngSlide = lngSlide + 1
      SelectBeginning
      DoEvents
      If lngFromRow > 0 Then
        SelectEnd
        DoEvents
        SelectRow lngFromRow - oTasks.Count, True
        DoEvents
      End If
      'PageDown here
      keybd_event VK_PAGEDOWN, 0, 0, 0
      keybd_event VK_PAGEDOWN, 0, KEYEVENTF_KEYUP, 0
      DoEvents
      SelectCellUp
      DoEvents
      On Error Resume Next
      Set oTask = Nothing
      Set oTask = ActiveSelection.Tasks(1)
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      TimescaleEdit MajorUnits:=0, MinorUnits:=2, MajorLabel:=0, MinorLabel:=10, MinorTicks:=True, Separator:=True, TierCount:=2
      DoEvents
      If Not oTask Is Nothing Then
        SelectBeginning True
        DoEvents
        lngToRow = ActiveSelection.Tasks.Count
        SelectBeginning
        DoEvents
        SelectRow lngFromRow
        DoEvents
        SelectRow lngToRow - lngFromRow, True, , True
        DoEvents
        EditCopyPicture Object:=False, ForPrinter:=0, SelectedRows:=1, FromDate:=Format(dtFrom, "m/d/yy hh:nn AMPM"), ToDate:=Format(dtTo, "m/d/yy hh:mm ampm"), ScaleOption:=pjCopyPictureTimescale, MaxImageHeight:=-1#, MaxImageWidth:=-1#, MeasurementUnits:=2 'pjCopyPictureShowOptions
        DoEvents
        oPresentation.Slides.Add oPresentation.Slides.Count + 1, ppLayoutCustom
        Set oSlide = oPresentation.Slides(oPresentation.Slides.Count)
        oSlide.Layout = ppLayoutChart
        strTitle = "Driving Path #" & vPath
        If CLng(vPath) <= 3 Then
          strTitle = strTitle & " (" & Choose(vPath, "Primary/Critical", "Secondary", "Tertiary") & ")"
        End If
        oSlide.Shapes(1).TextFrame.TextRange.Text = strTitle & IIf(lngSlide > 1, " (cont'd)", "")
        oSlide.Shapes(2).Delete
        oSlide.Shapes.Paste
        oSlide.Shapes(oSlide.Shapes.Count).Width = oSlide.Master.Width * 0.9
        oSlide.Shapes(oSlide.Shapes.Count).Left = (oSlide.Master.Width / 2) - (oSlide.Shapes(oSlide.Shapes.Count).Width / 2)
        If oSlide.Shapes(oSlide.Shapes.Count).Top <> 108 Then oSlide.Shapes(oSlide.Shapes.Count).Top = 108
        lngFromRow = lngToRow
      Else
        SelectBeginning
        DoEvents
        SelectRow lngFromRow
        DoEvents
        SelectEnd True
        DoEvents
        EditCopyPicture Object:=False, ForPrinter:=0, SelectedRows:=1, FromDate:=Format(dtFrom, "m/d/yy hh:nn AMPM"), ToDate:=Format(dtTo, "m/d/yy hh:mm ampm"), ScaleOption:=pjCopyPictureTimescale, MaxImageHeight:=-1#, MaxImageWidth:=-1#, MeasurementUnits:=2 'pjCopyPictureShowOptions
        DoEvents
        oPresentation.Slides.Add oPresentation.Slides.Count + 1, ppLayoutCustom
        Set oSlide = oPresentation.Slides(oPresentation.Slides.Count)
        oSlide.Layout = ppLayoutChart
        strTitle = "Driving Path #" & vPath
        If CLng(vPath) <= 3 Then
          strTitle = strTitle & " (" & Choose(vPath, "Primary/Critical", "Secondary", "Tertiary") & ")"
        End If
        oSlide.Shapes(1).TextFrame.TextRange.Text = strTitle & IIf(lngSlide > 1, " (cont'd)", "")
        oSlide.Shapes(2).Delete
        oSlide.Shapes.Paste
        oSlide.Shapes(oSlide.Shapes.Count).Width = oSlide.Master.Width * 0.9
        oSlide.Shapes(oSlide.Shapes.Count).Left = (oSlide.Master.Width / 2) - (oSlide.Shapes(oSlide.Shapes.Count).Width / 2)
        If oSlide.Shapes(oSlide.Shapes.Count).Top <> 108 Then oSlide.Shapes(oSlide.Shapes.Count).Top = 108
        Exit Do
      End If
    Loop
    oPresentation.Save
next_path:
    Set oTasks = Nothing
  Next vPath
  SetAutoFilter "CP Driving Path Group ID"
  SelectBeginning
  're-select target task
  Find "Unique ID", "equals", oTargetTask.UniqueID
  If Not oPresentation.Saved Then oPresentation.Save
  
  MsgBox "Critical Path slides created.", vbInformation + vbOKOnly, "Complete"
  
  oPowerPoint.Activate
  
exit_here:
  On Error Resume Next
  Set oDrivingPaths = Nothing
  Set oShell = Nothing
  cptSpeed False
  Set pptExists = Nothing
  Set oTargetTask = Nothing
  Set oTask = Nothing
  Set oTasks = Nothing
  Set oPowerPoint = Nothing
  Set oPresentation = Nothing
  Set oSlide = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptCriticalPathTools_bas", "cptExportCriticalPath", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCriticalPathSelected()
  'objects
  Dim oTasks As MSProject.Tasks
  'booleans
  Dim blnErrorTrapping As Boolean
  
  blnErrorTrapping = cptErrorTrapping
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oTasks Is Nothing Then
    MsgBox "Please select a target task.", vbExclamation + vbOKOnly, "Driving Paths"
    GoTo exit_here
  End If
  If oTasks.Count <> 1 Then
    MsgBox "Please select a single target task.", vbExclamation + vbOKOnly, "Driving Paths"
    GoTo exit_here
  End If

  Call cptExportCriticalPath(ActiveProject, blnKeepOpen:=True, oTargetTask:=oTasks(1))
  
exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  Exit Sub
err_here:
  If Err.Number = 1101 Then
    MsgBox "Please a a single (non-summary, active, and incomplete) 'Target' task.", vbExclamation + vbOKOnly, "Trace Tools - Error"
  Else
    Call cptHandleErr("cptCriticalPathTools_bas", "cptExportCriticalPathSelected", Err, Erl)
  End If
  Resume exit_here
End Sub

Sub cptDrivingPath()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  singlePath = True
  Call DrivingPaths

exit_here:
  On Error Resume Next
  singlePath = False
  Exit Sub
err_here:
  Call cptHandleErr("cptCriticalPathTools_bas", "cptDrivingPath", Err, Erl)
  Resume exit_here
End Sub
