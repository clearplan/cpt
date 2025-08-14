Attribute VB_Name = "cptCriticalPathTools_bas"
'<cpt_version>v1.1.0</cpt_version>
Option Explicit

Sub cptExportCriticalPath(ByRef oProject As MSProject.Project, Optional blnSendEmail As Boolean = False, Optional blnKeepOpen As Boolean = False, Optional ByRef oTargetTask As MSProject.Task)
  'objects
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
  Dim strFileName As String, strProjectName As String, strDir As String
  'longs
  Dim lngTask As Long, lngTasks As Long, lngSlide As Long
  'dates
  Dim dtFrom As Date, dtTo As Date
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
  strFileName = strDir & cptGetProgramAcronym & "-CriticalPathAnalysis-" & Format(Now, "yyyy-mm-dd") & ".pptx"
  On Error Resume Next
  Set pptExists = oPowerPoint.Presentations(strFileName)
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not pptExists Is Nothing Then 'add timestamp to this file
    pptExists.Save
    pptExists.Close
  End If
  'might exist but be closed
  If Dir(strFileName) <> vbNullString Then
    If MsgBox(strFileName & " exists. Overwrite?", vbExclamation + vbYesNo, "File Exists") = vbYes Then
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
  oSlide.Shapes(1).TextFrame.TextRange.Text = strProjectName & vbCrLf & "Critical Path Analysis"
  oSlide.Shapes(2).TextFrame.TextRange.Text = cptGetUserFullName & vbCrLf & FormatDateTime(Now, vbShortDate)
  
  'for each primary,secondary,tertiary > make a slide
  For Each vPath In Split("1,2,3,4,5", ",")
    'copy the picture
    'SetAutoFilter FieldName:="CP Driving Paths", FilterType:=pjAutoFilterCustom, Test1:="contains", Criteria1:=CStr(vPath)
    SetAutoFilter FieldName:="CP Driving Path Group ID", FilterType:=pjAutoFilterIn, Criteria1:=CStr(vPath)

    Sort Key1:="Finish", Key2:="Duration", ascending2:=False, renumber:=False
    TimescaleEdit MajorUnits:=0, MinorUnits:=2, MajorLabel:=0, MinorLabel:=10, MinorTicks:=True, Separator:=True, TierCount:=2
    SelectBeginning
    Debug.Print vPath & ": " & FormatDateTime(dtFrom, vbShortDate) & " - " & FormatDateTime(dtTo, vbShortDate)
    SelectAll
    'account for when a path is somehow not found
    On Error Resume Next
    Set oTasks = ActiveSelection.Tasks
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oTasks Is Nothing Then GoTo next_path
    'account for when task count exceeds easily visible range
    'on powerpoint slide
    lngTasks = oTasks.Count
    lngSlide = 0
    lngTask = 0
    Do While lngTask <= lngTasks
      lngTask = lngTask + 20
      lngSlide = lngSlide + 1
      SelectBeginning
      SelectTaskField Row:=lngTask - 20, Column:="Name", Height:=20, Extend:=False
      EditCopyPicture Object:=False, ForPrinter:=0, SelectedRows:=1, FromDate:=Format(dtFrom, "m/d/yy hh:nn AMPM"), ToDate:=Format(dtTo, "m/d/yy hh:mm ampm"), ScaleOption:=pjCopyPictureTimescale, MaxImageHeight:=-1#, MaxImageWidth:=-1#, MeasurementUnits:=2  'pjCopyPictureShowOptions
      'paste the picture
      oPresentation.Slides.Add oPresentation.Slides.Count + 1, ppLayoutCustom
      Set oSlide = oPresentation.Slides(oPresentation.Slides.Count)
      oSlide.Layout = ppLayoutChart
      oSlide.Shapes(1).TextFrame.TextRange.Text = Choose(vPath, "Primary", "Secondary", "Tertiary", "Quaternary", "Quinary") & " Critical Path" & IIf(lngSlide > 1, " (cont'd)", "")
      oSlide.Shapes(2).Delete
      oSlide.Shapes.Paste
      oSlide.Shapes(oSlide.Shapes.Count).Width = oSlide.Master.Width * 0.9
      oSlide.Shapes(oSlide.Shapes.Count).Left = (oSlide.Master.Width / 2) - (oSlide.Shapes(oSlide.Shapes.Count).Width / 2)
      If oSlide.Shapes(oSlide.Shapes.Count).Top <> 108 Then oSlide.Shapes(oSlide.Shapes.Count).Top = 108
    Loop
    oPresentation.Save
next_path:
    Set oTasks = Nothing
  Next vPath
  SetAutoFilter "CP Driving Path Group ID"
  SelectBeginning
  If Not oPresentation.Saved Then oPresentation.Save
  
  MsgBox "Critical Path slides created.", vbInformation + vbOKOnly, "Complete"
  
  oPowerPoint.Activate
  
exit_here:
  On Error Resume Next
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
  Dim oTargetTask As MSProject.Task

  On Error GoTo err_here

  Set oTargetTask = ActiveCell.Task

  Call cptExportCriticalPath(ActiveProject, blnKeepOpen:=True, oTargetTask:=ActiveSelection.Tasks(1))
  
exit_here:
  On Error Resume Next
  Set oTargetTask = Nothing
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
