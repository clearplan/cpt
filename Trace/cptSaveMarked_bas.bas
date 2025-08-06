Attribute VB_Name = "cptSaveMarked_bas"
'<cpt_version>v1.0.9</cpt_version>
Option Explicit

Sub cptShowSaveMarked_frm()
  'objects
  Dim mySaveMarked_frm As cptSaveMarked_frm
  'strings
  Dim strFileName As String
  Dim strApplyFilter As String
  Dim strProgram As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set mySaveMarked_frm = New cptSaveMarked_frm
  cptUpdateMarked mySaveMarked_frm
  With mySaveMarked_frm
    .Caption = "Import Marked (" & cptGetVersion("cptSaveMarked_frm") & ")"
    strApplyFilter = cptGetSetting("SaveMarked", "chkApplyFilter")
    If Len(strApplyFilter) > 0 Then
      .chkApplyFilter = CBool(strApplyFilter)
    Else
      .chkApplyFilter = False
    End If
    strProgram = cptGetProgramAcronym
    If Len(strProgram) > 0 Then
      .cboProjects.AddItem strProgram
      .cboProjects.Value = strProgram
      .cboProjects.Locked = True
      .cboProjects.Enabled = False
    End If
    .lblDir.Caption = "%UserProfile%\cpt-backup\settings\cpt-marked.adtg; cpt-marked-details.adtg"
    .Show False
  End With

exit_here:
  On Error Resume Next
  Set mySaveMarked_frm = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_bas", "cptShowSaveMarked_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateMarked(ByRef mySaveMarked_frm As cptSaveMarked_frm, Optional strFilter As String)
  'objects
  Dim rstMarked As Object 'ADODB.Recordset 'Object
  'strings
  Dim strDir As String
  Dim strProject As String
  Dim strMarked As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  strProject = cptGetProgramAcronym
  If Len(strProject) = 0 Then
    MsgBox "Program Acronym is required for this feature.", vbExclamation + vbOKOnly, "Program Acronym Needed"
    GoTo exit_here
  End If
  
  'clear listboxes and reset headers
  With mySaveMarked_frm
    .cboProjects.Clear
    .lboMarked.Clear
    .lboMarked.AddItem
    .lboMarked.List(.lboMarked.ListCount - 1, 0) = "TIMESTAMP"
    .lboMarked.List(.lboMarked.ListCount - 1, 1) = "PROJECT"
    .lboMarked.List(.lboMarked.ListCount - 1, 2) = "DESCRIPTION"
    .lboMarked.List(.lboMarked.ListCount - 1, 3) = "COUNT"
    .lboDetails.Clear
    .lboDetails.AddItem
    .lboDetails.List(.lboDetails.ListCount - 1, 0) = "UID"
    .lboDetails.List(.lboDetails.ListCount - 1, 1) = "TASK"
  End With
  
  'get list of marked sets
  'todo: filter for where PROJECT=cptGetProgramAcronym?
  strMarked = strDir & "\cpt-marked.adtg"
  If Dir(strMarked) = vbNullString Then
    MsgBox "No marked tasks saved.", vbCritical + vbOKOnly, "Nada"
    GoTo exit_here
  End If
  Set rstMarked = CreateObject("ADODB.Recordset")
  With rstMarked
    .Open strMarked
    .Sort = "TSTAMP DESC"
    If Len(strFilter) > 0 Then
      .Filter = "DESCRIPTION Like '*" & strFilter & "*' AND PROJECT_ID='" & strProject & "'"
    Else
      .Filter = "PROJECT_ID='" & strProject & "'"
    End If
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        With mySaveMarked_frm
          .lboMarked.AddItem
          .lboMarked.List(.lboMarked.ListCount - 1, 0) = rstMarked(1)
          .lboMarked.List(.lboMarked.ListCount - 1, 1) = rstMarked(2)
          .lboMarked.List(.lboMarked.ListCount - 1, 2) = rstMarked(3)
        End With
        .MoveNext
      Loop
    End If
    .Filter = 0
    .Close
    
    'get marked task count
    strMarked = strDir & "\cpt-marked-details.adtg"
    rstMarked.Open strMarked
    With mySaveMarked_frm
      For lngItem = 1 To .lboMarked.ListCount - 1
        rstMarked.Filter = "TSTAMP=#" & CDate(.lboMarked.List(lngItem, 0)) & "#"
        .lboMarked.List(lngItem, 3) = rstMarked.RecordCount
        rstMarked.Filter = 0
      Next lngItem
    End With
  End With

exit_here:
  On Error Resume Next
  If rstMarked.State = 1 Then rstMarked.Close
  Set rstMarked = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_bas", "cptUpdateMarked", Err, Erl)
  Resume exit_here
End Sub

Sub cptSaveMarked()
  'objects
  Dim mySaveMarked_frm As cptSaveMarked_frm
  Dim oTask As MSProject.Task
  Dim rstMarked As Object 'ADODB.Recordset
  'strings
  Dim strProject As String
  Dim strGUID As String
  Dim strDescription As String
  Dim strMarked As String
  'longs
  Dim lngMarked As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtTimestamp As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_one
    If oTask.ExternalTask Then GoTo next_one
    If Not oTask.Active Then GoTo next_one
    If oTask.Marked Then lngMarked = lngMarked + 1
next_one:
  Next oTask
  
  If lngMarked = 0 Then
    MsgBox "There are no Marked Tasks.", vbExclamation + vbOKOnly, "What happened?"
    GoTo exit_here
  End If
  
  Set rstMarked = CreateObject("ADODB.Recordset")
  strMarked = cptDir & "\cpt-marked.adtg"
  If Dir(strMarked) = vbNullString Then
    rstMarked.Fields.Append "GUID", adGUID
    rstMarked.Fields.Append "TSTAMP", adDBTimeStamp
    rstMarked.Fields.Append "PROJECT_ID", adVarChar, 255
    rstMarked.Fields.Append "Description", adVarChar, 255
    rstMarked.Open
    rstMarked.Save strMarked, adPersistADTG
    rstMarked.Close
  End If
  If rstMarked.State <> 1 Then rstMarked.Open strMarked
  
  strProject = cptGetProgramAcronym
  If Len(strProject) = 0 Then
    MsgBox "You must set a program acronym to use this feature.", vbCritical + vbOKOnly, "Program Acronym Needed"
    GoTo exit_here
  End If

  If ActiveProject.Subprojects.Count > 0 Then
    MsgBox "Saved set will use UIDs keyed to this master project. You must, therefore, import this saved set back to this master project. Importing to the standalone subproject will fail.", vbExclamation + vbOKOnly, "Nota Bene"
  End If

  strDescription = InputBox("Describe this capture:", "Save Marked")
  If Len(strDescription) = 0 Then
    MsgBox "No description; nothing saved.", vbExclamation + vbOKOnly
    GoTo exit_here
  End If
  dtTimestamp = Now()
  rstMarked.AddNew Array(1, 2, 3), Array(dtTimestamp, strProject, strDescription)
  rstMarked.Update
  rstMarked.Save strMarked, adPersistADTG
  rstMarked.Close
  
  Set rstMarked = CreateObject("ADODB.Recordset")
  strMarked = cptDir & "\cpt-marked-details.adtg"
  If Dir(strMarked) = vbNullString Then
    rstMarked.Fields.Append "TSTAMP", adDBTimeStamp
    rstMarked.Fields.Append "UID", adInteger
    rstMarked.Open
  Else
    rstMarked.Open strMarked
  End If
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Marked Then
      rstMarked.AddNew Array(0, 1), Array(dtTimestamp, oTask.UniqueID)
      rstMarked.Update
    End If
next_task:
  Next oTask
  rstMarked.Save strMarked, adPersistADTG
  rstMarked.Close
  
  dtTimestamp = 0
  If Not cptGetUserForm("cptSaveMarked_frm") Is Nothing Then
    Set mySaveMarked_frm = cptGetUserForm("cptSaveMarked_frm")
    If Not IsNull(mySaveMarked_frm.lboMarked.Value) Then dtTimestamp = mySaveMarked_frm.lboMarked.Value
    cptUpdateMarked mySaveMarked_frm
    If dtTimestamp > 0 Then mySaveMarked_frm.lboMarked.Value = dtTimestamp
  End If

exit_here:
  On Error Resume Next
  Set oTask = Nothing
  If rstMarked.State = 1 Then rstMarked.Close
  Set rstMarked = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_bas", "cptSaveMarked", Err, Erl)
  Resume exit_here
End Sub
