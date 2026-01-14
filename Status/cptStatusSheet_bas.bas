Attribute VB_Name = "cptStatusSheet_bas"
'<cpt_version>v1.7.2</cpt_version>
Option Explicit
Private Const adVarChar As Long = 200
Private strStartingViewTopPane As String
Private strStartingViewBottomPane As String
Private strStartingTable As String
Private strStartingFilter As String
Private strStartingGroup As String
Private oAssignmentRange As Excel.Range
Private oNumberValidationRange As Excel.Range
Private oETCValidationRange As Excel.Range
Private oInputRange As Excel.Range
Private oUnlockedRange As Excel.Range
Public oEVTs As Scripting.Dictionary
Private Const lngForeColorValid As Long = -2147483630
Private Const lngBorderColorValid As Long = 8421504 '-2147483642

Sub cptShowStatusSheet_frm()
  'populate all outline codes, text, and number fields
  'populate UID,[user selections],Task Name,Duration,Forecast Start,Forecast Finish,Total Slack,[EVT],EV%,New EV%,BLW,Remaining Work,New ETC,BLS,BLF,Reason/Impact/Action
  'add pick list for EV% or default to Physical % Complete
  'objects
  Dim myStatusSheet_frm As cptStatusSheet_frm
  Dim oRecordset As ADODB.Recordset 'Object
  Dim oShell As Object
  Dim oTasks As MSProject.Tasks
  Dim rstFields As ADODB.Recordset 'Object
  Dim rstEVT As ADODB.Recordset 'Object
  Dim rstEVP As ADODB.Recordset 'Object
  'longs
  Dim lngField As Long
  Dim lngItem As Long
  Dim lngSelectedItems As Long
  'integers
  Dim intField As Integer
  'strings
  Dim strCptDir As String
  Dim strNewCustomFieldName As String
  Dim strLOE As String
  Dim strIgnoreLOE As String
  Dim strLookahead As String
  Dim strLookaheadDays As String
  Dim strAssignments As String
  Dim strKeepOpen As String
  Dim strExportNotes As String
  Dim strAllowAssignmentNotes As String
  Dim strNotesColTitle As String
  Dim strFileNamingConvention As String
  Dim strDir As String
  Dim strAllItems As String
  Dim strAppendStatusDate As String
  Dim strQuickPart As String
  Dim strCC As String
  Dim strSubject As String
  Dim strProtect As String
  Dim strDataValidation As String
  Dim strConditionalFormatting As String
  Dim strEmail As String
  Dim strEach As String
  Dim strCostTool As String
  Dim strHide As String
  Dim strCreate As String
  Dim strEVP As String
  Dim strEVT As String
  Dim strFieldName As String
  Dim strFileName As String
  'booleans
  Dim blnErrorTrapping As Boolean
  'dates
  Dim dtStatus As Date
  'variants
  Dim vFieldType As Variant

  'prevent spawning
  If Not cptGetUserForm("cptStatusSheet_frm") Is Nothing Then Exit Sub

  'confirm existence of tasks to export
  On Error Resume Next
  Set oTasks = ActiveProject.Tasks
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strCptDir = cptDir
  If oTasks Is Nothing Then
    MsgBox "This Project has no Tasks.", vbExclamation + vbOKOnly, "No Tasks"
    GoTo exit_here
  ElseIf oTasks.Count = 0 Then
    MsgBox "This Project has no Tasks.", vbExclamation + vbOKOnly, "No Tasks"
    GoTo exit_here
  End If

  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Please enter a Status Date.", vbExclamation + vbOKOnly, "No Status Date"
    If Not Application.ChangeStatusDate Then
      MsgBox "No Status Date. Exiting.", vbCritical + vbOKOnly, "No Status Date"
      GoTo exit_here
    End If
  End If
    
  'requires metrics settings
  If Not cptValidMap("EVP,EVT,LOE", blnConfirmationRequired:=True) Then
    MsgBox "No settings saved; cannot proceed.", vbExclamation + vbOKOnly, "Settings Required"
    GoTo exit_here
  End If
  
  'requires ms excel
  Application.StatusBar = "Validating OLE references..."
  DoEvents
  If Not cptCheckReference("Excel") Then
    #If Win64 Then
      MsgBox "A reference to Microsoft Excel (64-bit) could not be set.", vbExclamation + vbOKOnly, "Excel Required"
      GoTo exit_here
    #Else
      MsgBox "A reference to Microsoft Excel (32-bit) could not be set.", vbExclamation + vbOKOnly, "Excel Required"
    #End If
  End If
  'requires scripting (cptRegEx)
  If Not cptCheckReference("Scripting") Then GoTo exit_here
  
  'reset options
  Application.StatusBar = "Loading default settings..."
  DoEvents
  Set myStatusSheet_frm = New cptStatusSheet_frm
  With myStatusSheet_frm
    .Caption = "Create Status Sheets (" & cptGetVersion("cptStatusSheet_frm") & ")"
    .lboFields.Clear
    .lboExport.Clear
    .cboCostTool.Clear
    .cboCostTool.AddItem "COBRA"
    .cboCostTool.AddItem "MPM"
    .cboCostTool.AddItem "<none>"
    For lngItem = 0 To 2
      .cboCreate.AddItem
      .cboCreate.List(lngItem, 0) = lngItem
      .cboCreate.List(lngItem, 1) = Choose(lngItem + 1, "Single Workbook", "Worksheet for each", "Workbook for each")
    Next lngItem
    .chkSendEmails.Enabled = cptCheckReference("Outlook")
    .chkHide = True
    .chkConditionalFormatting = True
    .chkValidation = True
    .chkProtect = True
    .chkAllItems = False
    If Left(ActiveProject.Path, 2) = "<>" Or Left(ActiveProject.Path, 4) = "http" Then 'it is a server project: default to Desktop
      Set oShell = CreateObject("WScript.Shell")
      .txtDir = oShell.SpecialFolders("Desktop") & "\" 'Status Requests\" & IIf(.chkAppendStatusDate, "[yyyy-mm-dd]\", "")
    Else  'not a server project: use ActiveProject.Path
      .txtDir = ActiveProject.Path & "\" 'Status Requests\" & IIf(.chkAppendStatusDate, "[yyyy-mm-dd]\", "")
    End If
    .txtFileName.ForeColor = -2147483630 'lngForeColorValid
    .txtFileName = "StatusRequest_[yyyy-mm-dd]"
    .lblPathLength.Visible = False
  End With

  'set up arrays to capture values
  Application.StatusBar = "Getting local custom fields..."
  DoEvents
  Set rstFields = CreateObject("ADODB.Recordset")
  rstFields.Fields.Append "CONSTANT", adBigInt
  rstFields.Fields.Append "NAME", adVarChar, 200
  rstFields.Fields.Append "TYPE", adVarChar, 50
  rstFields.Open
  
  'cycle through and add custom fields (text, outline code, number only)
  For Each vFieldType In Array("Text|30", "Outline Code|10", "Number|20")
    Dim strFieldType As String
    Dim lngFieldCount As Long
    strFieldType = Split(vFieldType, "|")(0)
    lngFieldCount = Split(vFieldType, "|")(1)
    For intField = 1 To lngFieldCount
      lngField = FieldNameToFieldConstant(strFieldType & intField, pjTask)
      strFieldName = CustomFieldGetName(lngField)
      If Len(strFieldName) > 0 Then
        If strFieldType = "Number" Then
          rstFields.AddNew Array(0, 1, 2), Array(lngField, strFieldName, "Number")
        Else
          rstFields.AddNew Array(0, 1, 2), Array(lngField, strFieldName, "Text")
        End If
      End If
    Next intField
  Next vFieldType
  
  'add Physical % Complete
  rstFields.AddNew Array(0, 1, 2), Array(FieldNameToFieldConstant("Physical % Complete"), "Physical % Complete", "Number")
  
  'add Contact field
  rstFields.AddNew Array(0, 1, 2), Array(FieldNameToFieldConstant("Contact"), "Contact", "Text")
  
  'get enterprise custom fields
  Application.StatusBar = "Getting Enterprise custom fields..."
  DoEvents
  For lngField = 188776000 To 188778000 '2000 should do it for now
    If Len(FieldConstantToFieldName(lngField)) > 0 And FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      rstFields.AddNew Array(0, 1, 2), Array(lngField, FieldConstantToFieldName(lngField), "Enterprise")
    End If
  Next lngField

  'add custom fields
  Application.StatusBar = "Populating Export Field list box..."
  DoEvents
  rstFields.Sort = "NAME"
  If rstFields.RecordCount > 0 Then
    rstFields.MoveFirst
    With myStatusSheet_frm
      Do While Not rstFields.EOF
        If rstFields(1) = "Physical % Complete" Then GoTo skip_fields
        .lboFields.AddItem
        .lboFields.List(.lboFields.ListCount - 1, 0) = rstFields(0)
        .lboFields.List(.lboFields.ListCount - 1, 1) = rstFields(1)
        If rstFields(1) = "Resources" Then
          .lboFields.List(.lboFields.ListCount - 1, 2) = FieldConstantToFieldName(rstFields(0))
        ElseIf FieldNameToFieldConstant(rstFields(1), pjTask) >= 188776000 Then
          .lboFields.List(.lboFields.ListCount - 1, 2) = "Enterprise"
        Else
          .lboFields.List(.lboFields.ListCount - 1, 2) = FieldConstantToFieldName(rstFields(0))
        End If
skip_fields:
        'add to Each
        If rstFields(1) <> "Physical % Complete" Then .cboEach.AddItem rstFields(1)
        rstFields.MoveNext
      Loop
    End With
  Else
    MsgBox "No Custom Fields have been set up in this file.", vbInformation + vbOKOnly, "No Fields Found"
    GoTo exit_here
  End If
  
  'convert saved settings if they exist
  strFileName = strCptDir & "\settings\cpt-status-sheet.adtg"
  If Dir(strFileName) <> vbNullString Then
    Application.StatusBar = "Converting saved settings..."
    DoEvents
    With CreateObject("ADODB.Recordset")
      .Open strFileName
      If Not .EOF Then
        .MoveFirst
        cptSaveSetting "StatusSheet", "cboEVT", .Fields(0)
        cptSaveSetting "StatusSheet", "cboEVP", .Fields(1)
        cptSaveSetting "StatusSheet", "cboCreate", .Fields(2) - 1
        cptSaveSetting "StatusSheet", "chkHide", .Fields(3)
        If .Fields.Count >= 5 Then
          cptSaveSetting "StatusSheet", "cboCostTool", .Fields(4)
        End If
        If .Fields.Count >= 6 Then
          cptSaveSetting "StatusSheet", "cboEach", .Fields(5)
        End If
      End If
      .Close
      Kill strFileName
    End With
  End If
  
  'import saved settings
  With myStatusSheet_frm
    Application.StatusBar = "Getting saved settings..."
    DoEvents
    strCreate = cptGetSetting("StatusSheet", "cboCreate")
    If strCreate <> "" Then .cboCreate.Value = CLng(strCreate)
    strHide = cptGetSetting("StatusSheet", "chkHide")
    If strHide <> "" Then
      .chkHide = CBool(strHide)
    Else
      .chkHide = False
    End If
    .txtHideCompleteBefore.Enabled = .chkHide
    strCostTool = cptGetSetting("StatusSheet", "cboCostTool")
    If strCostTool <> "" Then .cboCostTool.Value = strCostTool
    If .cboCreate <> 0 Then
      strEach = cptGetSetting("StatusSheet", "cboEach")
      If strEach <> "" Then
        If rstFields.RecordCount > 0 Then
          rstFields.MoveFirst
          rstFields.Find "NAME='" & strEach & "'"
          If Not rstFields.EOF Then
            On Error Resume Next
            .cboEach.Value = strEach '<none> would not be found
            .txtFileName = "StatusRequest_[item]_[yyyy-mm-dd]"
            If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
            If Err.Number > 0 Then
              MsgBox "Unable to set 'For Each' Field to '" & rstFields(1) & "' - contact help@ClearPlanConsulting.com if you need assistance.", vbExclamation + vbOKOnly, "Cannot assign For Each"
              Err.Clear
            End If
          End If
        End If
      End If
      ActiveWindow.TopPane.Activate
      FilterClear
      strAllItems = cptGetSetting("StatusSheet", "chkAllItems")
      If strAllItems <> "" Then
        .chkAllItems = CBool(strAllItems)
      Else
        .chkAllItems = False
      End If
    Else
      ActiveWindow.TopPane.Activate
      FilterClear
    End If
    strDir = cptGetSetting("StatusSheet", "txtDir")
    If strDir <> "" Then .txtDir = strDir
    strFileNamingConvention = cptGetSetting("StatusSheet", "txtFileName")
    If strFileNamingConvention <> "" Then .txtFileName = strFileNamingConvention
    
    strEmail = cptGetSetting("StatusSheet", "chkEmail")
    If strEmail <> "" Then
      .chkSendEmails.Value = CBool(strEmail)  'this refreshes the quickparts list
      If .chkSendEmails Then
        .chkKeepOpen.Value = False
        .chkKeepOpen.Enabled = False
      End If
    Else
      .chkSendEmails.Value = False
      .chkKeepOpen.Enabled = True
    End If
    If .chkSendEmails Then
      strSubject = cptGetSetting("StatusSheet", "txtSubject")
      If strSubject <> "" Then
        .txtSubject.Value = strSubject
      Else
        .txtSubject = "Status Request WE [yyyy-mm-dd]"
      End If
      strCC = cptGetSetting("StatusSheet", "txtCC")
      If strCC <> "" Then .txtCC.Value = strCC
      'cboQuickParts updated when .chkSendEmails = true
    End If
    
    strConditionalFormatting = cptGetSetting("StatusSheet", "chkConditionalFormatting")
    If strConditionalFormatting <> "" Then
      .chkConditionalFormatting.Value = CBool(strConditionalFormatting)
    Else
      .chkConditionalFormatting.Value = False
    End If
    
    strDataValidation = cptGetSetting("StatusSheet", "chkDataValidation")
    If strDataValidation <> "" Then
      .chkValidation = CBool(strDataValidation)
    Else
      .chkValidation = True
    End If
    
    strProtect = cptGetSetting("StatusSheet", "chkLocked")
    If strProtect <> "" Then
      .chkProtect.Value = CBool(strProtect)
      cptSaveSetting "StatusSheet", "chkProtect", strProtect
    Else
      .chkProtect.Value = True
    End If
    cptDeleteSetting "StatusSheet", "chkLocked"
    
    strNotesColTitle = cptGetSetting("StatusSheet", "txtNotesColTitle")
    If Len(strNotesColTitle) > 0 Then
      .txtNotesColTitle.Value = strNotesColTitle
    Else
      .txtNotesColTitle = "Reason / Action / Impact"
    End If
    
    strExportNotes = cptGetSetting("StatusSheet", "chkExportNotes")
    If strExportNotes <> "" Then
      .chkExportNotes.Value = CBool(strExportNotes)
    Else
      .chkExportNotes.Value = False
    End If
        
    strKeepOpen = cptGetSetting("StatusSheet", "chkKeepOpen")
    If strKeepOpen <> "" Then
      .chkKeepOpen.Value = CBool(strKeepOpen)
      If .chkKeepOpen Then
        .chkSendEmails.Value = False
        .chkSendEmails.Enabled = False
      Else
        .chkSendEmails.Enabled = True
      End If
    Else
      .chkKeepOpen.Value = False
    End If
    
    strAssignments = cptGetSetting("StatusSheet", "chkAssignments")
    If strAssignments <> "" Then
      .chkAssignments.Value = CBool(strAssignments)
    Else
      .chkAssignments.Value = True 'default
    End If
    
    If .chkAssignments Then
      .chkAllowAssignmentNotes.Enabled = True
      strAllowAssignmentNotes = cptGetSetting("StatusSheet", "chkAllowAssignmentNotes")
      If strAllowAssignmentNotes <> "" Then
        .chkAllowAssignmentNotes.Value = CBool(strAllowAssignmentNotes)
      Else
        .chkAllowAssignmentNotes.Value = False 'default
      End If
    Else
      .chkAllowAssignmentNotes.Value = False
      .chkAllowAssignmentNotes.Enabled = False
    End If
    
    .txtLookaheadDays.Enabled = False
    .txtLookaheadDate.Enabled = False
    .lblLookaheadWeekday.Visible = False
    strLookahead = cptGetSetting("StatusSheet", "chkLookahead")
    If Len(strLookahead) > 0 Then
      .chkLookahead = CBool(strLookahead)
    Else
      .chkLookahead = False 'default
    End If
    
    If .chkLookahead Then
      .txtLookaheadDays.Enabled = True
      strLookaheadDays = cptGetSetting("StatusSheet", "txtLookaheadDays")
      If Len(strLookaheadDays) > 0 Then
        .txtLookaheadDays = CLng(strLookaheadDays)
        .lblLookaheadWeekday.Visible = True
      End If
      .txtLookaheadDate.Enabled = True
    End If
    
    .chkIgnoreLOE.Enabled = False
    strEVT = Split(cptGetSetting("Integration", "EVT"), "|")(1)
    strLOE = cptGetSetting("Integration", "LOE")
    If Len(strEVT) > 0 And Len(strLOE) > 0 Then
      .chkIgnoreLOE.Enabled = True
      .chkIgnoreLOE.ControlTipText = "Limit to tasks where " & strEVT & " <> " & strLOE
      strIgnoreLOE = cptGetSetting("StatusSheet", "chkIgnoreLOE")
      If Len(strIgnoreLOE) > 0 Then
        .chkIgnoreLOE = CBool(strIgnoreLOE)
      Else
        .chkIgnoreLOE = False
      End If
    End If
    
  End With

  'add saved export fields if they exist
  strFileName = strCptDir & "\settings\cpt-status-sheet-userfields.adtg"
  If Dir(strFileName) <> vbNullString Then
    Set oRecordset = CreateObject("ADODB.Recordset")
    With oRecordset
      .Open strFileName
      'todo: add program acronym field and filter for it?
      If .RecordCount > 0 Then
        .MoveFirst
        lngItem = 0
        Do While Not .EOF
          myStatusSheet_frm.lboExport.AddItem
          myStatusSheet_frm.lboExport.List(lngItem, 0) = .Fields(0) 'Field Constant
          myStatusSheet_frm.lboExport.List(lngItem, 1) = .Fields(1) 'Custom Field Name
          myStatusSheet_frm.lboExport.List(lngItem, 2) = .Fields(2) 'Local Field Name
          'todo: what was this for? no FieldConstantToFieldName(constant) returns "Custom"?
          'todo: was this for filtering out enterprise fields since CFGN = FCFN?
          'If cptRegEx(FieldConstantToFieldName(.Fields(0)), "[0-9]{1,}$") = "" Then GoTo next_item
          'If InStr("Custom", FieldConstantToFieldName(FieldNameToFieldConstant(.Fields(2)))) = 0 Then GoTo next_item
          If CustomFieldGetName(.Fields(0)) <> CStr(.Fields(1)) Then
            If FieldConstantToFieldName(.Fields(0)) = CStr(.Fields(1)) Then GoTo next_item
            If Len(CustomFieldGetName(.Fields(0))) > 0 Then
              strNewCustomFieldName = CustomFieldGetName(.Fields(0))
            Else
              strNewCustomFieldName = "<unnamed>"
            End If
            'prompt user to accept changed name or remove from list
            If MsgBox("Saved field '" & .Fields(1) & "' has been renamed to '" & strNewCustomFieldName & "'." & vbCrLf & vbCrLf & "Click Yes to accept the name change." & vbCrLf & "Click No to remove from export list.", vbExclamation + vbYesNo, "Confirm Export Field") = vbYes Then
              'update export list
              myStatusSheet_frm.lboExport.List(lngItem, 1) = CustomFieldGetName(.Fields(0))
              'update the adtg
              .Fields(1) = CustomFieldGetName(.Fields(0))
              .Update
            Else
              'remove from export list
              myStatusSheet_frm.lboExport.RemoveItem (lngItem)
              'remove from adtg
              .Delete adAffectCurrent
              .Update
              lngItem = lngItem - 1
            End If
          End If
next_item:
          lngItem = lngItem + 1
          .MoveNext
        Loop
      End If
      .Filter = 0
      'overwrite in case of field name changes
      .Save strFileName, adPersistADTG
      .Close
    End With
  End If
    
  'set the status date / hide complete
  If ActiveProject.StatusDate = "NA" Then
    myStatusSheet_frm.txtStatusDate.Value = FormatDateTime(DateAdd("d", 6 - Weekday(Now), Now), vbShortDate)
  Else
    myStatusSheet_frm.txtStatusDate.Value = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  End If
  dtStatus = CDate(myStatusSheet_frm.txtStatusDate.Value)
  'default to one week prior to status date
  myStatusSheet_frm.txtHideCompleteBefore.Value = DateAdd("d", -7, dtStatus)

  strAppendStatusDate = cptGetSetting("StatusSheet", "chkAppendStatusDate")
  If strAppendStatusDate <> "" Then
    myStatusSheet_frm.chkAppendStatusDate = CBool(strAppendStatusDate)
  Else
    myStatusSheet_frm.chkAppendStatusDate = False 'default
  End If

  'delete pre-existing search file
  strFileName = strCptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName

  'set up the view/table/filter
  Application.StatusBar = "Preparing View/Table/Filter..."
  DoEvents
  ActiveWindow.TopPane.Activate
  strStartingViewTopPane = ActiveWindow.TopPane.View.Name
  If Not ActiveWindow.BottomPane Is Nothing Then
    strStartingViewBottomPane = ActiveWindow.BottomPane.View.Name
    ActiveWindow.BottomPane.Activate
    Application.PaneClose
  Else
    strStartingViewBottomPane = "None"
  End If
  strStartingTable = ActiveProject.CurrentTable
  strStartingFilter = ActiveProject.CurrentFilter
  If ActiveProject.CurrentGroup = "Custom Group" Then
    MsgBox "An ad hoc Autofilter Group cannot be used." & vbCrLf & vbCrLf & "Please save the group and name it, or select another saved Group, before you proceed.", vbInformation + vbOKOnly, "Invalid Group"
    GoTo exit_here
  Else
    strStartingGroup = ActiveProject.CurrentGroup
  End If
  
  cptSpeed True
  ActiveWindow.TopPane.Activate
  If ActiveWindow.TopPane.View.Type <> pjTaskItem Then
    ViewApply "Gantt Chart"
    If ActiveProject.CurrentGroup <> "No Group" Then GroupApply "No Group" 'if not a task view, then group is irrelevant
  Else
    If strStartingGroup = "No Group" Then
      'no fake Group Summary UIDs will be used
    Else
      If CBool(cptGetSetting("StatusSheet", "chkAssignments")) Then
        If Not strStartingViewTopPane = "Task Usage" Then ViewApply "Task Usage"
      Else
        If Not strStartingViewTopPane = "Gantt Chart" Then ViewApply "Gantt Chart"
      End If
      'task usage view avoids fake Group Summary UIDs
      If ActiveProject.CurrentGroup <> strStartingGroup Then GroupApply strStartingGroup
    End If
  End If
  DoEvents
  
  OptionsViewEx DisplaySummaryTasks:=True, DisplayNameIndent:=True
  If strStartingGroup = "No Group" Then
    Sort "ID", , , , , , False, True 'OutlineShowAllTasks won't work without this
  Else
    If ActiveProject.CurrentGroup <> strStartingGroup Then
      On Error Resume Next
      GroupApply strStartingGroup
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    End If
  End If
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    Sort "ID", , , , , , False, True
    OutlineShowAllTasks
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptRefreshStatusTable myStatusSheet_frm, True 'this only runs when form is visible
  FilterClear 'added 9/28/2021
  FilterApply "cptStatusSheet Filter"
  If Len(strCreate) > 0 And Len(strEach) > 0 Then
    On Error Resume Next
    SetAutoFilter strEach, pjAutoFilterClear
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    DoEvents
  End If
  If strStartingGroup <> "No Group" Then
    On Error Resume Next
    GroupApply strStartingGroup
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  End If
  DoEvents
  Application.StatusBar = "Ready..."
  DoEvents
  myStatusSheet_frm.txtStatusDate.SetFocus
  cptSpeed True
  myStatusSheet_frm.Show 'Modal = True! Keep!
  
  'after user closes form, then:
  Application.StatusBar = "Restoring your view/table/filter/group..."
  DoEvents
  cptSpeed True
  ActiveWindow.TopPane.Activate
  ViewApply strStartingViewTopPane
  If strStartingViewBottomPane <> "None" Then
    If strStartingViewBottomPane = "Timeline" Then
      ViewApplyEx Name:="Timeline", applyto:=1
    Else
      PaneCreate
      ViewApplyEx strStartingViewBottomPane, applyto:=1
      ActiveWindow.TopPane.Activate
    End If
  End If
  If ActiveProject.CurrentTable <> strStartingTable Then TableApply strStartingTable
  SetSplitBar ShowColumns:=ActiveProject.TaskTables(ActiveProject.CurrentTable).TableFields.Count
  If ActiveProject.CurrentFilter <> strStartingFilter Then FilterApply strStartingFilter
  If ActiveProject.CurrentGroup <> strStartingGroup Then GroupApply strStartingGroup
  
exit_here:
  On Error Resume Next
  Unload myStatusSheet_frm
  Set myStatusSheet_frm = Nothing
  Set oRecordset = Nothing
  Set oShell = Nothing
  Application.StatusBar = ""
  cptSpeed False
  Set oTasks = Nothing
  If rstFields.State Then rstFields.Close
  Set rstFields = Nothing
  Exit Sub

err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptShowStatusSheet_frm", Err, Erl)
  Resume exit_here

End Sub

Sub cptCreateStatusSheet(ByRef myStatusSheet_frm As cptStatusSheet_frm)
  'objects
  Dim oListObject As Excel.ListObject
  Dim oTasks As MSProject.Tasks, oTask As MSProject.Task, oAssignment As MSProject.Assignment
  Dim oExcel As Excel.Application, oWorkbook As Excel.Workbook, oWorksheet As Excel.Worksheet, rng As Excel.Range
  Dim rSummaryTasks As Excel.Range, rMilestones As Excel.Range, rNormal As Excel.Range, rAssignments As Excel.Range
  Dim rDates As Excel.Range, rWork As Excel.Range, rMedium As Excel.Range, rCentered As Excel.Range, rEntry As Excel.Range
  Dim xlCells As Excel.Range, rngAll As Excel.Range
  Dim aSummaries As Object, aMilestones As Object, aNormal As Object, aAssignments As Object
  Dim rstEach As ADODB.Recordset, aTaskRow As Object, rstColumns As ADODB.Recordset
  'longs
  Dim lngSelectedItems As Long
  Dim lngFormatCondition As Long
  Dim lngConditionalFormats As Long
  Dim lngDayLabelDisplay As Long
  Dim lngTaskRow As Long
  Dim lngLastRow As Long
  Dim lngDateFormat As Long
  Dim lngGroups As Long
  Dim lngLastCol As Long
  Dim lngTaskCount As Long, lngTask As Long, lngHeaderRow As Long
  Dim lngRow As Long, lngCol As Long, lngField As Long
  Dim lngNameCol As Long, lngBaselineWorkCol As Long, lngRemainingWorkCol As Long, lngEach As Long
  Dim lngNotesCol As Long, lngColumnWidth As Long
  Dim lngASCol As Long, lngAFCol As Long, lngETCCol As Long, lngEVPCol As Long
  #If Win64 And VBA7 Then '<issue53>
    Dim t As LongPtr, tTotal As LongPtr '<issue53>
  #Else '<issue53>
    Dim t As Long, tTotal As Long '<issue53>
  #End If '<issue53>
  Dim lngItem As Long
  'strings
  Dim strStatusDate As String
  Dim strCriteria As String
  Dim strFieldName As String
  Dim strMsg As String
  Dim strEVT As String, strEVP As String, strDir As String, strFileName As String
  Dim strFirstCell As String
  Dim strItem As String
  'dates
  Dim dtStatus As Date
  'variants
  Dim vBorder As Variant
  Dim vArray As Variant
  Dim vHeader As Variant
  Dim vCol As Variant
  Dim vUserFields As Variant
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnConditionalFormattingLegend As Boolean
  Dim blnKeepOpen As Boolean
  Dim blnProtect As Boolean
  Dim blnValidation As Boolean
  Dim blnConditionalFormatting As Boolean
  Dim blnPerformanceTest As Boolean
  Dim blnSpace As Boolean
  Dim blnEmail As Boolean

  'check reference
  If Not cptCheckReference("Excel") Then
    MsgBox "Reference to Microsoft Excel not found.", vbCritical + vbOKOnly, "Is Excel installed?"
    GoTo exit_here
  End If

  'ensure required module exists
  If Not cptModuleExists("cptCore_bas") Then
    MsgBox "Please install the ClearPlan 'cptCore_bas' module.", vbExclamation + vbOKOnly, "Missing Module"
    GoTo exit_here
  End If

  'this boolean spits out a speed test to the immediate window
  blnPerformanceTest = False
  If blnPerformanceTest Then tTotal = GetTickCount

  On Error Resume Next
  Set oTasks = ActiveProject.Tasks
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'ensure project has tasks
  If oTasks Is Nothing Then
    MsgBox "This project has no tasks.", vbExclamation + vbOKOnly, "Create Status Sheet"
    GoTo exit_here
  End If
    
  myStatusSheet_frm.lblStatus.Caption = " Analyzing project..."
  Application.StatusBar = "Analyzing project..."
  DoEvents
  blnValidation = myStatusSheet_frm.chkValidation = True
  blnProtect = myStatusSheet_frm.chkProtect = True
  blnConditionalFormatting = myStatusSheet_frm.chkConditionalFormatting = True
  blnConditionalFormattingLegend = myStatusSheet_frm.chkConditionalFormattingLegend = True
  blnEmail = myStatusSheet_frm.chkSendEmails = True
  blnKeepOpen = myStatusSheet_frm.chkKeepOpen
  'get task count
  If blnPerformanceTest Then t = GetTickCount
  SelectAll
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then
    MsgBox "There are no incomplete tasks in this schedule.", vbExclamation + vbOKOnly, "No Tasks Found"
    GoTo exit_here
  End If
  lngTaskCount = oTasks.Count
  If blnPerformanceTest Then Debug.Print "<=====PERFORMANCE TEST " & Now() & "=====>"

  myStatusSheet_frm.lblStatus.Caption = " Setting up Workbook..."
  Application.StatusBar = "Setting up Workbook..."
  DoEvents
  'set up an excel Workbook
  If blnPerformanceTest Then t = GetTickCount
  Set oExcel = CreateObject("Excel.Application") 'do not use GetObject
  'oExcel.Visible = False
  oExcel.WindowState = xlMinimized
  '/=== debug ==\
  If Not blnErrorTrapping Then oExcel.Visible = True
  '\=== debug ===/
  
  If blnPerformanceTest Then Debug.Print "set up excel Workbook: " & (GetTickCount - t) / 1000

  'get status date
  If ActiveProject.StatusDate = "NA" Then
    dtStatus = Now()
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  'cptCopyData task data and applies formatting, validation, protection
  'cptCopyData also extracts existing assignment data and applies formatting, validation, protection
  'obective is to only loop through tasks once
  
  'copy/paste the data
  lngHeaderRow = 8
  With myStatusSheet_frm

    'ensure proper task-assignment rolldown
    'only matters if we're applying a filter before export
    'so it only matters if 'worksheet for each' or 'workbook for each'
    If .cboCreate.Value > 0 And .chkAssignments Then
      If Not cptValidateAssignmentRolldown(.cboEach.Value) Then
        If MsgBox("Ignore Assignment roll-down error(s) and proceed anyway (not recommended)?", vbQuestion + vbYesNo, "Danger, Will Robinson!") = vbNo Then
          GoTo exit_here
        Else
          Msgbox "...good choice.", vbInformation + vbOKOnly, ""
        End If
      End if
      Dim strStartingGroup as String
      strStartingGroup = ActiveProject.CurrentGroup
      cptRefreshStatusTable myStatusSheet_frm, blnFilterOnly:=True
      If strStartingGroup <> "No Group" Then GroupApply strStartingGroup
    End If

    If .cboCreate.Value = "0" Then 'single workbook
      
      SelectAll
      On Error Resume Next
      Set oTasks = Nothing
      Set oTasks = ActiveSelection.Tasks
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oTasks Is Nothing Then
        .lblStatus.Caption = "No incomplete tasks ...skipped"
        Application.StatusBar = .lblStatus.Caption
        GoTo exit_here
      End If
      
      'get excel
      If oExcel Is Nothing Then Set oExcel = CreateObject("Excel.Application")
      Set oWorkbook = oExcel.Workbooks.Add
      oExcel.Calculation = xlCalculationManual
      oExcel.ScreenUpdating = False
      Set oWorksheet = oWorkbook.Sheets(1)
      oWorksheet.Name = "Status Sheet"
      
      'copy data
      If blnPerformanceTest Then t = GetTickCount
      .lblStatus.Caption = "Creating Workbook..."
      Application.StatusBar = .lblStatus.Caption
      DoEvents
      cptCopyData myStatusSheet_frm, oWorksheet, lngHeaderRow
      If blnPerformanceTest Then Debug.Print "copy data: " & (GetTickCount - t) / 1000
      
      'add legend
      If blnPerformanceTest Then t = GetTickCount
      .lblStatus.Caption = "Building legend..."
      Application.StatusBar = .lblStatus.Caption
      cptAddLegend oWorksheet, dtStatus
      .lblStatus.Caption = "Building legend...done."
      Application.StatusBar = .lblStatus.Caption
      If blnPerformanceTest Then Debug.Print "set up legend: " & (GetTickCount - t) / 1000
      
      'final formatting
      .lblStatus.Caption = "Formatting..."
      Application.StatusBar = .lblStatus.Caption
      cptFinalFormats oWorksheet
      .lblStatus.Caption = "Formatting...done."
      Application.StatusBar = .lblStatus.Caption
            
      oWorksheet.Calculate
      
      If blnProtect Then 'protect the sheet
        .lblStatus.Caption = "Protecting..."
        Application.StatusBar = .lblStatus.Caption
        oWorksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
        oWorksheet.EnableSelection = xlNoRestrictions
        .lblStatus.Caption = "Protecting...done."
        Application.StatusBar = .lblStatus.Caption
      End If
      
      Set oInputRange = Nothing
      Set oNumberValidationRange = Nothing
      Set oETCValidationRange = Nothing
      Set oUnlockedRange = Nothing
      Set oAssignmentRange = Nothing
      
      If blnConditionalFormatting And blnConditionalFormattingLegend Then
        .lblStatus.Caption = "Adding Conditional Formatting Legend..."
        Application.StatusBar = .lblStatus.Caption
        cptAddConditionalFormattingLegend oWorkbook
        .lblStatus.Caption = "Adding Conditional Formatting Legend...done."
        Application.StatusBar = .lblStatus.Caption
        oWorkbook.Sheets(1).Activate
      End If
                  
      'save the workbook
      .lblStatus.Caption = "Saving Workbook..."
      Application.StatusBar = .lblStatus.Caption
      strFileName = cptSaveStatusSheet(myStatusSheet_frm, oWorkbook)
      If Len(strFileName) = 0 Then
        .lblStatus.Caption = "Saving Workbook...error!"
      Else
        .lblStatus.Caption = "Saving Workbook...done."
      End If
      Application.StatusBar = .lblStatus.Caption
      DoEvents
      
      oExcel.Calculation = xlCalculationAutomatic
      oExcel.ScreenUpdating = True
      
      If blnEmail And Len(strFileName) > 0 Then 'send workbook
        .lblStatus.Caption = "Creating Email..."
        Application.StatusBar = .lblStatus.Caption
        DoEvents
        'must close before attaching to email
        oWorkbook.Close True
        oExcel.Wait Now + TimeValue("00:00:02")
        oExcel.Quit
        Set oExcel = Nothing
        cptSendStatusSheet myStatusSheet_frm, strFileName
        .lblStatus.Caption = "Creating Email...done."
        Application.StatusBar = .lblStatus.Caption
        DoEvents
      Else
        If blnKeepOpen Then
          oExcel.Visible = True
          oWorkbook.Activate
        Else
          oWorkbook.Close True
          oExcel.Wait Now + TimeValue("00:00:002")
        End If
      End If 'blnEmail
      
    ElseIf .cboCreate.Value = "1" Then  'worksheet for each
      
      Set oWorkbook = oExcel.Workbooks.Add
      oExcel.Calculation = xlCalculationManual
      oExcel.ScreenUpdating = False
      For lngItem = 0 To .lboItems.ListCount - 1
        If .lboItems.Selected(lngItem) Then
          strItem = .lboItems.List(lngItem, 0)
          SetAutoFilter .cboEach.Value, pjAutoFilterCustom, "equals", strItem
          SelectAll
          Set oTasks = Nothing
          On Error Resume Next
          Set oTasks = ActiveSelection.Tasks
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oTasks Is Nothing Then
            .lblStatus.Caption = "No incomplete tasks for " & strItem & "...skipped"
            Application.StatusBar = .lblStatus.Caption
            GoTo next_worksheet
          End If
          
          'create worksheet
          Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
          oWorksheet.Name = strItem

          'copy data
          If blnPerformanceTest Then t = GetTickCount
          .lblStatus.Caption = "Creating Worksheet for " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          DoEvents
          cptCopyData myStatusSheet_frm, oWorksheet, lngHeaderRow, strItem
          If blnPerformanceTest Then Debug.Print "copy data: " & (GetTickCount - t) / 1000

          'add legend
          If blnPerformanceTest Then t = GetTickCount
          .lblStatus.Caption = "Building legend for " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          cptAddLegend oWorksheet, dtStatus
          .lblStatus.Caption = "Building legend for " & strItem & "...done."
          Application.StatusBar = .lblStatus.Caption
          If blnPerformanceTest Then Debug.Print "set up legend: " & (GetTickCount - t) / 1000
          
          'final formatting
          .lblStatus.Caption = "Formatting " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          cptFinalFormats oWorksheet
          .lblStatus.Caption = "Formatting " & strItem & "...done."
          Application.StatusBar = .lblStatus.Caption
          
          oWorksheet.Calculate
          
          If blnProtect Then 'protect the sheet
            .lblStatus.Caption = "Protecting " & strItem & "..."
            Application.StatusBar = .lblStatus.Caption
            oWorksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
            oWorksheet.EnableSelection = xlNoRestrictions
            .lblStatus.Caption = "Protecting " & strItem & "...done."
            Application.StatusBar = .lblStatus.Caption
          End If
          
          Set oInputRange = Nothing
          Set oNumberValidationRange = Nothing
          Set oETCValidationRange = Nothing
          Set oUnlockedRange = Nothing
          Set oAssignmentRange = Nothing
          
next_worksheet:
          .lblStatus.Caption = "Creating Worksheet for " & strItem & "...done"
          Application.StatusBar = .lblStatus.Caption
          DoEvents
          
        End If
      Next lngItem
      
      Set oWorksheet = Nothing
      On Error Resume Next
      Set oWorksheet = oWorkbook.Sheets("Sheet1")
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If Not oWorksheet Is Nothing Then oWorksheet.Delete
      
      If blnConditionalFormatting And blnConditionalFormattingLegend Then
        .lblStatus.Caption = "Adding Conditional Formatting Legend..."
        Application.StatusBar = .lblStatus.Caption
        cptAddConditionalFormattingLegend oWorkbook
        .lblStatus.Caption = "Adding Conditional Formatting Legend...done."
        Application.StatusBar = .lblStatus.Caption
      End If
      oWorkbook.Sheets(1).Activate

      'save the workbook
      .lblStatus.Caption = "Saving Workbook..."
      Application.StatusBar = .lblStatus.Caption
      strFileName = cptSaveStatusSheet(myStatusSheet_frm, oWorkbook)
      If Len(strFileName) = 0 Then
        .lblStatus.Caption = "Saving Workbook...error!"
      Else
        .lblStatus.Caption = "Saving Workbook...done."
      End If
      Application.StatusBar = .lblStatus.Caption
      DoEvents
      
      oExcel.Calculation = xlCalculationAutomatic
      oExcel.ScreenUpdating = True
      
      If blnEmail And Len(strFileName) > 0 Then 'send workbook
        .lblStatus.Caption = "Creating Email..."
        Application.StatusBar = .lblStatus.Caption
        DoEvents
        'must close before attaching
        oWorkbook.Close True
        oExcel.Wait Now + TimeValue("00:00:02")
        oExcel.Quit
        Set oExcel = Nothing
        cptSendStatusSheet myStatusSheet_frm, strFileName
        .lblStatus.Caption = "Creating Email...done."
        Application.StatusBar = .lblStatus.Caption
        DoEvents
      Else
        If blnKeepOpen Then
          oExcel.Visible = True
          oWorkbook.Activate
        Else
          oWorkbook.Close True
          oExcel.Wait Now + TimeValue("00:00:002")
        End If
      End If 'blnEmail
      
    ElseIf .cboCreate.Value = "2" Then  'workbook for each
      
      For lngItem = 0 To .lboItems.ListCount - 1
        If .lboItems.Selected(lngItem) Then
          strItem = .lboItems.List(lngItem, 0)
          SetAutoFilter .cboEach.Value, pjAutoFilterCustom, "equals", strItem
          SelectAll
          On Error Resume Next
          Set oTasks = Nothing
          Set oTasks = ActiveSelection.Tasks
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oTasks Is Nothing Then
            .lblStatus.Caption = "No incomplete tasks for " & strItem & "...skipped"
            Application.StatusBar = .lblStatus.Caption
            GoTo next_workbook
          End If
          
          'get excel
          If oExcel Is Nothing Then Set oExcel = CreateObject("Excel.Application")
          Set oWorkbook = oExcel.Workbooks.Add
          oExcel.Calculation = xlCalculationManual
          oExcel.ScreenUpdating = False
          Set oWorksheet = oWorkbook.Sheets(1)
          oWorksheet.Name = "Status Request"
          
          'copy data
          If blnPerformanceTest Then t = GetTickCount
          .lblStatus.Caption = "Creating Workbook for " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          DoEvents
          cptCopyData myStatusSheet_frm, oWorksheet, lngHeaderRow, strItem
          If blnPerformanceTest Then Debug.Print "copy data: " & (GetTickCount - t) / 1000
          
          'add legend
          If blnPerformanceTest Then t = GetTickCount
          .lblStatus.Caption = "Building legend for " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          cptAddLegend oWorksheet, dtStatus
          .lblStatus.Caption = "Building legend for " & strItem & "...done."
          Application.StatusBar = .lblStatus.Caption
          If blnPerformanceTest Then Debug.Print "set up legend: " & (GetTickCount - t) / 1000
          
          'final formatting
          .lblStatus.Caption = "Formatting " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          cptFinalFormats oWorksheet
          .lblStatus.Caption = "Formatting " & strItem & "...done."
          Application.StatusBar = .lblStatus.Caption
          
          oWorksheet.Calculate
          
          If blnProtect Then 'protect the sheet
            .lblStatus.Caption = "Protecting " & strItem & "..."
            Application.StatusBar = .lblStatus.Caption
            oWorksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
            oWorksheet.EnableSelection = xlNoRestrictions
            .lblStatus.Caption = "Protecting " & strItem & "...done."
            Application.StatusBar = .lblStatus.Caption
          End If
          
          Set oInputRange = Nothing
          Set oNumberValidationRange = Nothing
          Set oETCValidationRange = Nothing
          Set oUnlockedRange = Nothing
          Set oAssignmentRange = Nothing
                    
          If blnConditionalFormatting And blnConditionalFormattingLegend Then
            .lblStatus.Caption = "Adding Conditional Formatting Legend (" & strItem & ")..."
            Application.StatusBar = .lblStatus.Caption
            cptAddConditionalFormattingLegend oWorkbook
            .lblStatus.Caption = "Adding Conditional Formatting Legend (" & strItem & ")...done."
            Application.StatusBar = .lblStatus.Caption
            oWorkbook.Sheets(1).Activate
          End If
                    
          'save the workbook
          .lblStatus.Caption = "Saving Workbook for " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          strFileName = cptSaveStatusSheet(myStatusSheet_frm, oWorkbook, strItem)
          If Len(strFileName) = 0 Then
            .lblStatus.Caption = "Saving Workbook for " & strItem & "...error!"
          Else
            .lblStatus.Caption = "Saving Workbook for " & strItem & "...done."
          End If
          Application.StatusBar = .lblStatus.Caption
          DoEvents
          
          oExcel.Calculation = xlCalculationAutomatic
          'oExcel.ScreenUpdating = True 'not yet
          
          If blnEmail And Len(strFileName) > 0 Then 'send workbook
            .lblStatus.Caption = "Creating Email for " & strItem & "..."
            Application.StatusBar = .lblStatus.Caption
            DoEvents
            'must close before attaching to email
            oWorkbook.Close True
            oExcel.Wait Now + TimeValue("00:00:02")
            oExcel.Quit
            Set oExcel = Nothing
            cptSendStatusSheet myStatusSheet_frm, strFileName, strItem
            .lblStatus.Caption = "Creating Email for " & strItem & "...done"
            Application.StatusBar = .lblStatus.Caption
            DoEvents
          Else
            If blnKeepOpen Then
              oExcel.Visible = True
              oWorkbook.Activate
            Else
              oWorkbook.Close True
              oExcel.Wait Now + TimeValue("00:00:002")
            End If
          End If 'blnEmail
        End If '.lboItems.Selected(lngItem)
                
next_workbook:
        
      Next lngItem
      
      If Not blnEmail Then
        oExcel.ScreenUpdating = True
        oExcel.Visible = True
      End If
      
    End If 'cboCreate
    lngSelectedItems = 0
    For lngItem = 0 To .lboItems.ListCount - 1
      If .lboItems.Selected(lngItem) Then
        lngSelectedItems = lngSelectedItems + 1
      End If
    Next lngItem
    If CLng(.cboCreate) = 2 And lngSelectedItems > 1 Then 'workbook for each
      .lblStatus.Caption = "Workbooks complete"
      Application.StatusBar = .lblStatus.Caption
      MsgBox "Workbooks complete.", vbInformation + vbOKOnly, "Create Status Sheet(s)"
    Else 'single workbook
      .lblStatus.Caption = "Workbook complete"
      Application.StatusBar = .lblStatus.Caption
      MsgBox "Workbook complete", vbInformation + vbOKOnly, "Create Status Sheet(s)"
    End If
    DoEvents
  End With

exit_here:
  On Error Resume Next
  Set oListObject = Nothing
  If oExcel.Workbooks.Count > 0 Then oExcel.Calculation = xlAutomatic
  oExcel.ScreenUpdating = True
  oExcel.EnableEvents = True
  Application.StatusBar = ""
  cptSpeed False
  Set oTasks = Nothing
  Set oTask = Nothing
  Set oAssignment = Nothing
  If blnEmail Then oExcel.Quit
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set rng = Nothing
  Set rSummaryTasks = Nothing
  Set rMilestones = Nothing
  Set rNormal = Nothing
  Set rAssignments = Nothing
  Set rDates = Nothing
  Set rWork = Nothing
  Set rMedium = Nothing
  Set rCentered = Nothing
  Set rEntry = Nothing
  If rstEach.State Then rstEach.Close
  Set rstEach = Nothing
  Set aTaskRow = Nothing
  If rstColumns.State Then rstColumns.Close
  Set rstColumns = Nothing
  Set xlCells = Nothing
  Exit Sub

err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptCreateStatusSheet", Err, Erl)
  If Not oExcel Is Nothing Then
    If Not oWorkbook Is Nothing Then oWorkbook.Close False
    oExcel.Quit
  End If
  Resume exit_here

End Sub

Sub cptRefreshStatusTable(ByRef myStatusSheet_frm As cptStatusSheet_frm, Optional blnOverride As Boolean = False, Optional blnFilterOnly As Boolean = False)
  'objects
  'strings
  Dim strLOE As String
  Dim strEVT As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtLookahead As Date

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not myStatusSheet_frm.Visible And blnOverride = False Then GoTo exit_here

  If Not blnOverride Then cptSpeed True
  If blnFilterOnly Then GoTo filter_only
  
  'reset the view
  Application.StatusBar = "Resetting the cptStatusSheet View..."
  Application.ActiveWindow.TopPane.Activate
  If myStatusSheet_frm.chkAssignments Then
    ViewApply "Task Usage"
  Else
    ViewApply "Gantt Chart"
  End If
  OptionsViewEx DisplaySummaryTasks:=True
  
  'reset the group
  Application.StatusBar = "Resetting the cptStatusSheet Group..."
  If ActiveProject.CurrentGroup <> "No Group" Then
    'strStartingGroup = ActiveProject.CurrentGroup
    GroupApply "No Group"
  End If
  
  'reset the table
  Application.StatusBar = "Resetting the cptStatusSheet Table..."
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=10, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Unique ID", Title:="UID", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  lngItem = 0
  If myStatusSheet_frm.lboExport.ListCount > 0 Then
    For lngItem = 0 To myStatusSheet_frm.lboExport.ListCount - 1
      If Not IsNull(myStatusSheet_frm.lboExport.List(lngItem, 0)) Then
        TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(myStatusSheet_frm.lboExport.List(lngItem, 0)), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
      End If
    Next lngItem
  End If
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Name", Title:="Task Name / Scope", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Remaining Duration", Title:="", Width:=12, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Total Slack", Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Baseline Start", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Baseline Finish", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False, ShowAddNewColumn:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Start", Title:="Forecast Start", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Finish", Title:="Forecast Finish", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Actual Start", Title:="New Forecast/ Actual Start", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Actual Finish", Title:="New Forecast/ Actual Finish", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=Split(cptGetSetting("Integration", "EVT"), "|")(1), Title:="EVT", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=Split(cptGetSetting("Integration", "EVP"), "|")(1), Title:="EV%", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=Split(cptGetSetting("Integration", "EVP"), "|")(1), Title:="New EV%", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Baseline Work", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Remaining Work", Title:="ETC", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Remaining Work", Title:="New ETC", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableApply Name:="cptStatusSheet Table"
  SetSplitBar ShowColumns:=ActiveProject.TaskTables(ActiveProject.CurrentTable).TableFields.Count

filter_only:
  'reset the filter
  Application.StatusBar = "Resetting the cptStatusSheet Filter..."
  FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Actual Finish", test:="equals", Value:="NA", ShowInMenu:=False, ShowSummaryTasks:=True
  If myStatusSheet_frm.chkHide And IsDate(myStatusSheet_frm.txtHideCompleteBefore) Then
    FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", NewFieldName:="Actual Finish", test:="is greater than or equal to", Value:=myStatusSheet_frm.txtHideCompleteBefore, operation:="Or", ShowSummaryTasks:=True
  End If
  If Edition = pjEditionProfessional Then
    FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", NewFieldName:="Active", test:="equals", Value:="Yes", ShowInMenu:=False, ShowSummaryTasks:=True, Parenthesis:=True
  End If
  With myStatusSheet_frm
    If .chkLookahead And .txtLookaheadDate.BorderColor <> 192 Then
      dtLookahead = CDate(.txtLookaheadDate) & " 5:00 PM"
      FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", NewFieldName:="Start", test:="is less than or equal to", Value:=dtLookahead, operation:="And", Parenthesis:=False
    End If
    If .chkIgnoreLOE Then
      strEVT = Split(cptGetSetting("Integration", "EVT"), "|")(1)
      strLOE = cptGetSetting("Integration", "LOE")
      FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", NewFieldName:=strEVT, test:="does not equal", Value:=strLOE, operation:="And", Parenthesis:=False
    End If
  End With
  FilterApply "cptStatusSheet Filter"
  
  If Len(strStartingGroup) > 0 Then
    Application.StatusBar = "Restoring the cptStatusSheet Group..."
    GroupApply strStartingGroup
  End If
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  If Not blnOverride Then cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptRefreshStatusTable", Err, Erl)
  Err.Clear
  Resume exit_here
End Sub

Private Sub cptAddLegend(ByRef oWorksheet As Excel.Worksheet, dtStatus As Date)
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    
  oWorksheet.Cells(1, 1).Value = "Status Date:"
  oWorksheet.Cells(1, 1).Font.Bold = True
  oWorksheet.Cells(1, 2) = FormatDateTime(dtStatus, vbShortDate)
  oWorksheet.Names.Add "STATUS_DATE", oWorksheet.[B1]
  oWorksheet.Cells(1, 2).Font.Bold = True
  oWorksheet.Cells(1, 2).Font.Size = 14
  oWorksheet.Cells(1, 2).HorizontalAlignment = xlCenter
  oWorksheet.Cells(1, 2).Style = "Note"
  oWorksheet.Cells(1, 2).Columns.AutoFit
  'current
  oWorksheet.Cells(3, 1).Style = "Input" '<issue58>
  oWorksheet.Cells(3, 2) = "Task is active or within current status window. Cell requires update."
  'within two weeks
  oWorksheet.Cells(4, 1).Style = "Neutral" '<issue58>
  oWorksheet.Cells(4, 1).BorderAround xlContinuous, xlThin, , -8421505
  oWorksheet.Cells(4, 2) = "Task is within two week look-ahead. Please review forecast dates."
  'complete
  oWorksheet.Cells(5, 1).Style = "Explanatory Text"
  oWorksheet.Cells(5, 1) = "AaBbCc"
  oWorksheet.Cells(5, 2) = "Task is complete."
  'summary
  oWorksheet.Cells(6, 1) = "AaBbCc"
  oWorksheet.Cells(6, 1).Font.Bold = True
  oWorksheet.Cells(6, 1).Interior.ThemeColor = xlThemeColorDark1
  oWorksheet.Cells(6, 1).Interior.TintAndShade = -0.149998474074526
  oWorksheet.Cells(6, 2) = "MS Project Summary Task (Rollup).  No update required."

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptAddLegend", Err, Erl)
  Resume exit_here

End Sub

Private Sub cptCopyData(ByRef myStatusSheet_frm As cptStatusSheet_frm, ByRef oWorksheet As Excel.Worksheet, lngHeaderRow As Long, Optional strItem As String)
  'objects
  Dim oAssignmentETCRange As Excel.Range
  Dim oAssignment As MSProject.Assignment
  Dim oFormatRange As Object
  Dim oDict As Scripting.Dictionary
  Dim oRecordset As ADODB.Recordset
  Dim oFirstCell As Excel.Range
  Dim oETCRange As Excel.Range
  Dim oEVPRange As Excel.Range
  Dim oNFRange As Excel.Range
  Dim oNSRange As Excel.Range
  Dim oComment As Excel.Comment
  Dim oEVTRange As Excel.Range
  Dim oCompleted As Excel.Range
  Dim oMilestoneRange As Excel.Range
  Dim oClearRange As Excel.Range
  Dim oSummaryRange As Excel.Range
  Dim oDateValidationRange As Excel.Range
  Dim oTwoWeekWindowRange As Excel.Range
  Dim oTask As MSProject.Task
  'strings
  Dim strCETC As String
  Dim strCEVP As String
  Dim strCF As String
  Dim strCS As String
  Dim strFormula As String
  Dim strETC As String
  Dim strEVP As String
  Dim strNF As String
  Dim strAF As String
  Dim strFF As String
  Dim strNS As String
  Dim strAS As String
  Dim strFS As String
  Dim strNotesColTitle As String
  Dim strLOE As String
  Dim strEVT As String
  Dim strEVTList As String
  'longs
  Dim lngCEVPCol As Long
  Dim lngCFCol As Long
  Dim lngCETCCol As Long
  Dim lngCSCol As Long
  Dim lngNFCol As Long
  Dim lngNSCol As Long
  Dim lngLastRow As Long
  Dim lngFormatCondition As Long
  Dim lngFormatConditions As Long
  Dim lngEVT As Long
  Dim lngEVTCol As Long
  Dim lngLastCol As Long
  Dim lngBLWCol As Long
  Dim lngETCCol As Long
  Dim lngTask As Long
  Dim lngRow As Long
  Dim lngNameCol As Long
  Dim lngTasks As Long
  Dim lngCol As Long
  Dim lngBLSCol As Long
  Dim lngBLFCol As Long
  Dim lngASCol As Long
  Dim lngAFCol As Long
  Dim lngEVPCol As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnAssignments As Boolean
  Dim blnAlerts As Boolean
  Dim blnLOE As Boolean
  Dim blnProtect As Boolean
  Dim blnValidation As Boolean
  Dim blnConditionalFormats As Boolean
  Dim blnMilestones As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  Dim dtEarliestStart As Date
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  dtStatus = ActiveProject.StatusDate
  blnValidation = myStatusSheet_frm.chkValidation = True
  blnConditionalFormats = myStatusSheet_frm.chkConditionalFormatting = True
  blnProtect = myStatusSheet_frm.chkProtect = True
  ActiveWindow.TopPane.Activate
try_again:
  SelectAll
  EditCopy
  DoEvents
  oWorksheet.Application.Wait 5000
  On Error Resume Next
  oWorksheet.Paste oWorksheet.Cells(lngHeaderRow, 1), False
  If Err.Number = 1004 Then 'try again
    EditCopy
    oWorksheet.Application.Wait 5000
    oWorksheet.Paste oWorksheet.Cells(lngHeaderRow, 1), False
    oWorksheet.Application.Wait 5000
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  oWorksheet.Application.Wait 5000
  oWorksheet.Cells.WrapText = False
  oWorksheet.Application.ActiveWindow.Zoom = 85
  oWorksheet.Cells.Font.Name = "Calibri"
  oWorksheet.Cells.Font.Size = 11
  oWorksheet.Rows(lngHeaderRow).Font.Bold = True
  oWorksheet.Columns.AutoFit
  'format the colums
  blnAlerts = oWorksheet.Application.DisplayAlerts
  If myStatusSheet_frm.cboCreate <> 2 Then
    strItem = ""
  'Else
  '  strItem = myStatusSheet_frm.lboItems.List(myStatusSheet_frm.lboItems.ListIndex, 0)
  End If
  If blnAlerts Then oWorksheet.Application.DisplayAlerts = False
  For lngCol = 1 To ActiveSelection.FieldIDList.Count
    oWorksheet.Columns(lngCol).ColumnWidth = ActiveProject.TaskTables("cptStatusSheet Table").TableFields(lngCol + 1).Width + 2
    oWorksheet.Cells(lngHeaderRow, lngCol).WrapText = True
    If InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "Start") > 0 Then
      oWorksheet.Columns(lngCol).Replace "NA", ""
      oWorksheet.Columns(lngCol).NumberFormat = "m/d/yyyy"
    ElseIf InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "Finish") > 0 Then
      oWorksheet.Columns(lngCol).Replace "NA", ""
      oWorksheet.Columns(lngCol).NumberFormat = "m/d/yyyy"
    ElseIf InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "Work") > 0 Or InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "ETC") > 0 Then
      oWorksheet.Columns(lngCol).Style = "Comma"
    End If
  Next lngCol
  oWorksheet.Application.DisplayAlerts = blnAlerts
  
  'format the header
  lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
  If lngLastCol > ActiveProject.TaskTables("cptStatusSheet Table").TableFields.Count + 10 Then GoTo try_again
  oWorksheet.Columns(lngLastCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
  oWorksheet.Columns(lngLastCol + 1).ColumnWidth = 40
  strNotesColTitle = cptGetSetting("StatusSheet", "txtNotesColTitle")
  If Len(strNotesColTitle) > 0 Then
    oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Offset(0, 1).Value = strNotesColTitle
  Else
    oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Offset(0, 1).Value = "Reason / Action / Impact" 'default
  End If
  With oWorksheet.Cells(lngHeaderRow, 1).Resize(, ActiveProject.TaskTables(ActiveProject.CurrentTable).TableFields.Count)
    .Interior.ThemeColor = xlThemeColorLight2
    .Interior.TintAndShade = 0
    .Font.ThemeColor = xlThemeColorDark1
    .Font.TintAndShade = 0
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
  End With
  
  'get LOE settings
  strEVT = cptGetSetting("Integration", "EVT")
  If Len(strEVT) > 0 Then
    lngEVT = CLng(Split(strEVT, "|")(0))
  End If
  strLOE = cptGetSetting("Integration", "LOE")
  
  'format the data rows
  lngNameCol = oWorksheet.Rows(lngHeaderRow).Find("Task Name / Scope", lookat:=xlWhole).Column
  lngASCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Start", lookat:=xlPart).Column
  lngAFCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Finish", lookat:=xlPart).Column
  lngEVPCol = oWorksheet.Rows(lngHeaderRow).Find("New EV%", lookat:=xlWhole).Column
  lngEVTCol = oWorksheet.Rows(lngHeaderRow).Find("EVT", lookat:=xlWhole).Column
  'todo: add Milestone EVT
  lngETCCol = oWorksheet.Rows(lngHeaderRow).Find("New ETC", lookat:=xlWhole).Column
  lngBLWCol = oWorksheet.Rows(lngHeaderRow).Find("Baseline Work", lookat:=xlWhole).Column
  lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
  lngTasks = ActiveSelection.Tasks.Count
  lngTask = 0
  For Each oTask In ActiveSelection.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    'find the row of the current task
    On Error Resume Next
    lngRow = 0
    lngRow = oWorksheet.Columns(1).Find(oTask.UniqueID, lookat:=xlWhole).Row
    If Err.Number = 91 Then
      MsgBox "UID " & oTask.UniqueID & " not found on worksheet!" & vbCrLf & vbCrLf & "You may need to re-run...", vbExclamation + vbOKOnly, "ERROR"
      GoTo next_task
    End If
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    'capture if task is LOE
    blnLOE = oTask.GetField(lngEVT) = strLOE
    If oTask.Summary Then
      If oSummaryRange Is Nothing Then
        Set oSummaryRange = oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol))
      Else
        Set oSummaryRange = oWorksheet.Application.Union(oSummaryRange, oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol)))
      End If
      If oClearRange Is Nothing Then
        Set oClearRange = oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol))
      Else
        Set oClearRange = oWorksheet.Application.Union(oClearRange, oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol)))
      End If
      GoTo next_task
    End If
    If oTask.Milestone Then
      If oMilestoneRange Is Nothing Then
        Set oMilestoneRange = oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol))
      Else
        Set oMilestoneRange = oWorksheet.Application.Union(oMilestoneRange, oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol)))
      End If
'      If oClearRange Is Nothing Then
'        Set oClearRange = oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol))
'      Else
'        Set oClearRange = oWorksheet.Application.Union(oClearRange, oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol)))
'      End If
'      GoTo next_task 'don't skip - need to unlock foreceast dates for milestones, too
    End If
    If blnLOE Then
      oWorksheet.Cells(lngRow, lngEVPCol - 1) = "'-"
      oWorksheet.Cells(lngRow, lngEVPCol) = "'-"
    End If
    'format completed
    If IsDate(oTask.ActualFinish) Then
      If oCompleted Is Nothing Then
        Set oCompleted = oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol))
      Else
        Set oCompleted = oWorksheet.Application.Union(oCompleted, oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol)))
      End If
      GoTo get_assignments
    End If
    'we know now that it is incomplete
    'unlock new finish
    If oUnlockedRange Is Nothing Then
      Set oUnlockedRange = oWorksheet.Cells(lngRow, lngAFCol)
    Else
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngAFCol))
    End If
    'unlock new EV (discrete only)
    If Not blnLOE Then Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngEVPCol))
    'unlock new start if not started
    If Not IsDate(oTask.ActualStart) Then
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngASCol))
    End If
    'capture status formating:
    'tasks requiring status:
    If oTask.Start < dtStatus And Not IsDate(oTask.ActualStart) Then 'should have started
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
      If Not blnLOE Then Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngEVPCol))
    End If
    If oTask.Finish <= dtStatus And Not IsDate(oTask.ActualFinish) Then 'should have finished
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
      If Not blnLOE Then Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngEVPCol))
    End If
    If IsDate(oTask.ActualStart) And Not IsDate(oTask.ActualFinish) Then 'in progress
      'highlight EVP for discrete only
      If Not blnLOE Then
        If oInputRange Is Nothing Then
          Set oInputRange = oWorksheet.Cells(lngRow, lngEVPCol)
        Else
          Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngEVPCol))
        End If
      End If
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
    End If
    'two week window
    If oTask.Start > dtStatus And oTask.Start <= DateAdd("d", 14, dtStatus) Then
      If oTwoWeekWindowRange Is Nothing Then
        Set oTwoWeekWindowRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oTwoWeekWindowRange = oWorksheet.Application.Union(oTwoWeekWindowRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
    End If
    If oTask.Finish > dtStatus And oTask.Finish <= DateAdd("d", 14, dtStatus) Then
      If oTwoWeekWindowRange Is Nothing Then
        Set oTwoWeekWindowRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oTwoWeekWindowRange = oWorksheet.Application.Union(oTwoWeekWindowRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
    End If
    
    'capture data validation
    If blnValidation Then
      If Not IsDate(oTask.ActualStart) Then
        If oDateValidationRange Is Nothing Then
          Set oDateValidationRange = oWorksheet.Cells(lngRow, lngASCol)
        Else
          Set oDateValidationRange = oWorksheet.Application.Union(oDateValidationRange, oWorksheet.Cells(lngRow, lngASCol))
        End If
      End If
      If Not IsDate(oTask.ActualFinish) Then
        If oDateValidationRange Is Nothing Then
          Set oDateValidationRange = oWorksheet.Cells(lngRow, lngAFCol)
        Else
          Set oDateValidationRange = oWorksheet.Application.Union(oDateValidationRange, oWorksheet.Cells(lngRow, lngAFCol))
        End If
        'allow incomplete tasks to have EVP updated
        If Not blnLOE Then
          If oNumberValidationRange Is Nothing Then
            Set oNumberValidationRange = oWorksheet.Cells(lngRow, lngEVPCol)
          Else
            Set oNumberValidationRange = oWorksheet.Application.Union(oNumberValidationRange, oWorksheet.Cells(lngRow, lngEVPCol))
          End If
        End If
      End If
    End If 'blnValidation
    
    'capture conditional formatting ranges
    blnConditionalFormats = myStatusSheet_frm.chkConditionalFormatting
    If Not blnLOE And blnConditionalFormats Then 'todo: include LOE?
      If oNSRange Is Nothing Then
        Set oNSRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oNSRange = oWorksheet.Application.Union(oNSRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
      If oNFRange Is Nothing Then
        Set oNFRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oNFRange = oWorksheet.Application.Union(oNFRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
      If oEVPRange Is Nothing Then
        Set oEVPRange = oWorksheet.Cells(lngRow, lngEVPCol)
      Else
        Set oEVPRange = oWorksheet.Application.Union(oEVPRange, oWorksheet.Cells(lngRow, lngEVPCol))
      End If
      If oEVTRange Is Nothing Then
        Set oEVTRange = oWorksheet.Cells(lngRow, lngEVTCol)
      Else
        Set oEVTRange = oWorksheet.Application.Union(oEVTRange, oWorksheet.Cells(lngRow, lngEVTCol))
      End If
      If oETCRange Is Nothing Then
        Set oETCRange = oWorksheet.Cells(lngRow, lngETCCol)
      Else
        Set oETCRange = oWorksheet.Application.Union(oETCRange, oWorksheet.Cells(lngRow, lngETCCol))
      End If
    End If
        
''    'add EVT comment - this is slow, and often fails
'    oWorksheet.Application.ScreenUpdating = True
'    Set oComment = oWorksheet.Cells(lngRow, lngEVTCol).AddComment(oEVTs.Item(oTask.GetField(FieldNameToFieldConstant(myStatusSheet_frm.cboEVT.Value))))
'    oComment.Shape.TextFrame.Characters.Font.Bold = False
'    oComment.Shape.TextFrame.AutoSize = True
'    oWorksheet.Application.ScreenUpdating = False
    
    'unlock comment column
    If oUnlockedRange Is Nothing Then
      Set oUnlockedRange = oWorksheet.Cells(lngRow, lngLastCol)
    Else
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngLastCol))
    End If
    
    'export notes
    If myStatusSheet_frm.chkExportNotes Then
      oWorksheet.Cells(lngRow, lngLastCol) = Trim(Replace(oTask.Notes, vbCr, vbLf))
    End If
    
    'format comments column
    oWorksheet.Cells(lngRow, lngLastCol).HorizontalAlignment = xlLeft
    oWorksheet.Cells(lngRow, lngLastCol).NumberFormat = "General"
    oWorksheet.Cells(lngRow, lngLastCol).WrapText = True
    
get_assignments:
    blnAssignments = CBool(cptGetSetting("StatusSheet", "chkAssignments"))
    If blnAssignments Then
      If oTask.Assignments.Count > 0 And Not IsDate(oTask.ActualFinish) Then
        cptGetAssignmentData myStatusSheet_frm, oTask, oWorksheet, lngRow, lngHeaderRow, lngNameCol, lngETCCol - 1
      ElseIf IsDate(oTask.ActualFinish) Then
        For Each oAssignment In oTask.Assignments
          Set oAssignment = Nothing
          On Error Resume Next
          Set oAssignment = oTask.Assignments.UniqueID(oWorksheet.Cells(lngRow + 1, 1).Value)
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If Not oAssignment Is Nothing Then
            oWorksheet.Rows(lngRow + 1).EntireRow.Delete
          End If
        Next oAssignment
        Set oAssignment = Nothing
      End If
    Else
      oWorksheet.Cells(lngRow, lngETCCol) = oTask.RemainingWork / 60
      oWorksheet.Cells(lngRow, lngETCCol - 1) = oTask.RemainingWork / 60
      oWorksheet.Cells(lngRow, lngBLWCol) = oTask.BaselineWork / 60
      'add ETC to inputrange
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngETCCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngETCCol))
      End If
      'add to ETC Validation Range
      If oETCValidationRange Is Nothing Then
        Set oETCValidationRange = oWorksheet.Cells(lngRow, lngETCCol)
      Else
        Set oETCValidationRange = oWorksheet.Application.Union(oETCValidationRange, oWorksheet.Cells(lngRow, lngETCCol))
      End If
    End If
        
    oWorksheet.Columns(1).AutoFit
    oWorksheet.Rows(lngRow).AutoFit

next_task:
    lngTask = lngTask + 1
    myStatusSheet_frm.lblProgress.Width = (lngTask / lngTasks) * myStatusSheet_frm.lblStatus.Width
  Next oTask
  
  'clear out group summary stuff
  If ActiveProject.CurrentGroup <> "No Group" Then
    lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row
    For lngRow = lngHeaderRow + 1 To lngLastRow
      If Not blnAssignments Then
        'remove UID on Group Summaries
        Set oTask = Nothing
        On Error Resume Next
        Set oTask = ActiveProject.Tasks.UniqueID(oWorksheet.Cells(lngRow, 1))
        If oTask Is Nothing Then
          oWorksheet.Cells(lngRow, 1).Value = ""
        Else
          If Trim(oTask.Name) <> Trim(oWorksheet.Cells(lngRow, lngNameCol).Value) Then
            oWorksheet.Cells(lngRow, 1).Value = ""
          End If
        End If
      End If
      If Len(oWorksheet.Cells(lngRow, 1)) = 0 Then
        If oSummaryRange Is Nothing Then
          Set oSummaryRange = oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol))
        Else
          Set oSummaryRange = oWorksheet.Application.Union(oSummaryRange, oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol)))
        End If
        If oClearRange Is Nothing Then
          Set oClearRange = oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol))
        Else
          Set oClearRange = oWorksheet.Application.Union(oClearRange, oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol)))
        End If
      End If
    Next lngRow
  End If
  
  If Not oClearRange Is Nothing Then oClearRange.Value = ""
  If Not oSummaryRange Is Nothing Then
    oSummaryRange.Interior.ThemeColor = xlThemeColorDark1
    oSummaryRange.Interior.TintAndShade = -0.149998474074526
    oSummaryRange.Font.Bold = True
  End If
  If Not oMilestoneRange Is Nothing Then
    oMilestoneRange.Font.ThemeColor = xlThemeColorAccent6
    oMilestoneRange.Font.TintAndShade = -0.249977111117893
  End If
  If Not oCompleted Is Nothing Then
    oCompleted.Font.Italic = True
    oCompleted.Font.ColorIndex = 16
  End If
  If blnValidation And Not oDateValidationRange Is Nothing Then
    'date validation range
    If ActiveProject.Subprojects.Count > 0 Then
      dtEarliestStart = cptGetEarliestStart
    Else
      dtEarliestStart = ActiveProject.ProjectStart
    End If
    With oDateValidationRange.Validation
      .Delete
      oWorksheet.Application.WindowState = xlNormal
      .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=FormatDateTime(dtEarliestStart, vbShortDate), Formula2:="12/31/2149"
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = "Date Only"
      .ErrorTitle = "Date Only"
      .InputMessage = "Please enter a date between " & FormatDateTime(dtEarliestStart, vbShortDate) & " and 12/31/2149 in 'm/d/yyyy' format."
      .ErrorMessage = "Please enter a date between " & FormatDateTime(dtEarliestStart, vbShortDate) & " and 12/31/2149 in 'm/d/yyyy' format."
      .ShowInput = True
      .ShowError = True
    End With
  End If
  If blnValidation And Not oNumberValidationRange Is Nothing Then
    'number validation range (contains EV% only)
    With oNumberValidationRange.Validation
      .Delete
      .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1"
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = "Number Only"
      .ErrorTitle = "Number Only"
      .InputMessage = "Please enter a percentage between 0% and 100%."
      .ErrorMessage = "Please enter a percentage between 0% and 100%."
      .ShowInput = True
      .ShowError = True
    End With
  End If
  If blnValidation And Not oETCValidationRange Is Nothing Then
    'ETC validation range (contains ETC only)
    With oETCValidationRange.Validation
      .Delete
      .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = "Number Only"
      .ErrorTitle = "Number Only"
      .InputMessage = "Please enter a number greater than, or equal to, zero (0)."
      .ErrorMessage = "Please enter a number greater than, or equal to, zero (0)"
      .ShowInput = True
      .ShowError = True
    End With
  End If
  If Not blnAssignments And Not oETCValidationRange Is Nothing Then
    oETCValidationRange.Locked = False
    oETCValidationRange.HorizontalAlignment = xlCenter
  End If
  
  'format the Assignment Rows
  If Not oAssignmentRange Is Nothing Then
    With oAssignmentRange.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -4.99893185216834E-02
      .PatternTintAndShade = 0
    End With
  End If
  'unlock the input cells
  If Not oInputRange Is Nothing Then oInputRange.Locked = False
  'unlock cells whether blnProtect = True or False
  If Not oUnlockedRange Is Nothing Then oUnlockedRange.Locked = False
  If Not oTwoWeekWindowRange Is Nothing Then oTwoWeekWindowRange.Locked = False
  'add EVT gloassary - test comment
  If Not oEVTRange Is Nothing Then
    If myStatusSheet_frm.cboCostTool = "COBRA" Then
      strEVTList = "A - Level of Effort,"
      strEVTList = strEVTList & "B - Milestones,"
      strEVTList = strEVTList & "C - % Complete,"
      strEVTList = strEVTList & "D - Units Complete,"
      strEVTList = strEVTList & "E - 50-50,"
      strEVTList = strEVTList & "F - 0-100,"
      strEVTList = strEVTList & "G - 100-0,"
      strEVTList = strEVTList & "H - User Defined,"
      strEVTList = strEVTList & "J - Apportioned,"
      strEVTList = strEVTList & "K - Planning Package,"
      strEVTList = strEVTList & "L - Assignment % Complete,"
      strEVTList = strEVTList & "M - Calculated Apportionment,"
      strEVTList = strEVTList & "N - Steps,"
      strEVTList = strEVTList & "O - Earned As Spent,"
      strEVTList = strEVTList & "P - % Complete Manual Entry,"
    ElseIf myStatusSheet_frm.cboCostTool = "MPM" Then
      strEVTList = "0 - No EVM required,"
      strEVTList = strEVTList & "1 - 0/100,"
      strEVTList = strEVTList & "'2 - 25/75,"
      strEVTList = strEVTList & "'3 - 40/60,"
      strEVTList = strEVTList & "'4 - 50/50,"
      strEVTList = strEVTList & "5 - % Complete,"
      strEVTList = strEVTList & "6 - LOE,"
      strEVTList = strEVTList & "7 - Earned Standards,"
      strEVTList = strEVTList & "8 - Milestone Weights,"
      strEVTList = strEVTList & "9 - BCWP Entry,"
      strEVTList = strEVTList & "A - Apportioned,"
      strEVTList = strEVTList & "P - Milestone Weights with % Complete,"
      strEVTList = strEVTList & "K - Key Event,"
    End If
    If Len(strEVTList) > 0 Then
      oWorksheet.Cells(lngHeaderRow, lngLastCol + 2).Value = "Earned Value Techniques (EVT)"
      oWorksheet.Cells(lngHeaderRow, lngLastCol).Copy
      oWorksheet.Cells(lngHeaderRow, lngLastCol + 2).PasteSpecial xlPasteFormats
      oWorksheet.Range(oWorksheet.Cells(lngHeaderRow + 1, lngLastCol + 2), oWorksheet.Cells(lngHeaderRow + 1, lngLastCol + 2).Offset(UBound(Split(strEVTList, ",")), 0)).Value = oWorksheet.Application.Transpose(Split(strEVTList, ","))
      oWorksheet.Columns(lngLastCol + 2).AutoFit
    End If
    
  End If
  
  If blnConditionalFormats And Not oNSRange Is Nothing Then
    
    'entry required cells:
    'green by nature or green as last condition
    'if empty then "input"
    'if invalid then "red"
    'if valid then "green"
    
    oNSRange.Select
    Set oFirstCell = oWorksheet.Application.ActiveCell
    oFirstCell.Select
    strNS = oFirstCell.Address(False, True)
    lngNSCol = lngASCol  'new start
    lngNFCol = lngAFCol  'new finish
    lngCSCol = oWorksheet.Cells(lngHeaderRow).Find(what:="Forecast Start", lookat:=xlWhole).Column
    lngCFCol = oWorksheet.Cells(lngHeaderRow).Find(what:="Forecast Finish", lookat:=xlWhole).Column
    lngCEVPCol = oWorksheet.Cells(lngHeaderRow).Find(what:="EV%", lookat:=xlWhole).Column
    lngCETCCol = oWorksheet.Cells(lngHeaderRow).Find(what:="ETC", lookat:=xlWhole).Column
    strCS = oWorksheet.Cells(oFirstCell.Row, lngCSCol).Address(False, True)
    strCF = oWorksheet.Cells(oFirstCell.Row, lngCFCol).Address(False, True)
    strNF = oWorksheet.Cells(oFirstCell.Row, lngNFCol).Address(False, True)
    strEVP = oWorksheet.Cells(oFirstCell.Row, lngEVPCol).Address(False, True)
    strCEVP = oWorksheet.Cells(oFirstCell.Row, lngCEVPCol).Address(False, True)
    strETC = oWorksheet.Cells(oFirstCell.Row, lngETCCol).Address(False, True)
    strCETC = oWorksheet.Cells(oFirstCell.Row, lngCETCCol).Address(False, True)
    strEVT = oWorksheet.Cells(oFirstCell.Row, lngEVTCol).Address(False, True)
    'set up derived addresses for ease of formula writing
    'AS = (NS>0,NS<=SD)
    strAS = strNS & ">0," & strNS & "<=STATUS_DATE"
    'AF = (NF>0,NF<=SD)
    strAF = strNF & ">0," & strNF & "<=STATUS_DATE"
    'FS = (NS>0,NS>SD)
    strFS = strNS & ">0," & strNS & ">STATUS_DATE"
    'FF = (NF>0,NF>SD)
    strFF = strNF & ">0," & strNF & ">STATUS_DATE"
    
    'create map of ranges
    Set oDict = CreateObject("Scripting.Dictionary")
    Set oDict.item("NS") = oNSRange
    oNSRange.FormatConditions.Delete
    Set oDict.item("NF") = oNFRange
    oNFRange.FormatConditions.Delete
    Set oDict.item("EVP") = oEVPRange
    oEVPRange.FormatConditions.Delete
    Set oDict.item("EVT") = oEVTRange
    oEVTRange.FormatConditions.Delete
    Set oDict.item("ETC") = oETCRange
    oETCRange.FormatConditions.Delete
    If blnAssignments And Not oAssignmentRange Is Nothing Then
      Set oAssignmentETCRange = oWorksheet.Application.Intersect(oAssignmentRange, oWorksheet.Columns(lngETCCol))
      Set oDict.item("AssignmentETC") = oAssignmentETCRange
      oAssignmentETCRange.FormatConditions.Delete
    End If
    
    'capture list of formulae
    Set oRecordset = CreateObject("ADODB.Recordset")
    oRecordset.Fields.Append "RANGE", adVarChar, 13
    oRecordset.Fields.Append "FORMULA", adVarChar, 255
    oRecordset.Fields.Append "FORMAT", adVarChar, 10
    oRecordset.Fields.Append "STOP", adInteger
    oRecordset.Open
    
    '<cpt-breadcrumbs:format-conditions>
    'VARIABLES:
    'SD = Status Date
    'CS = Current Start
    'CF = Current Finish
    'NS = New Start
    'NF = New Finish
    'EVT = Earned Value Technique
    'EVP = Earned Value Percent
    'ETC = Estimate to Complete
    '
    'DERIVED VARIABLES:
    'AS = (NS>0,NS<=SD)
    'AF = (NF>0,NF<=SD)
    'FS = (NS>0,NS>SD)
    'FF = (NF>0,NF>SD)
    '
    'NS:AND(NS>0,NS<=SD) -> COMPLETE
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strAS & ")", "COMPLETE")
    'NS:AND(NS>0,NS>SD) -> GOOD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & ">0," & strNS & ">STATUS_DATE)", "GOOD")
    'NS:AND(CS<=(SD+14),NS=0) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strCS & "<=(STATUS_DATE+14)," & strNS & "=0)", "NEUTRAL")
    'NS:AND(CS<=SD,NS=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strCS & "<=STATUS_DATE," & strNS & "=0)", "BAD") 'should have started
    'NS:AND(NS>0,NF>0,NS>NF) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & ">0," & strNF & ">0," & strNS & ">" & strNF & ")", "BAD")
    'todo: oRecordset.AddNew Array(0, 1, 2), Array("NS", "=IF(""NS>0,NF>0,NS>NF,"",""BAD"",AND(" & strNS & ">0," & strNF & ">0," & strNS & ">" & strNF & "))","BAD")
    'NS:AND(NS=0,AF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & "=0," & strAF & ")", "BAD")
    'NS:AND(FS>0,EVP>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strFS & "," & strEVP & ">0)", "BAD")
    'NS:AND(NS=0,EVP>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & "=0," & strEVP & ">0)", "BAD")
    'NS:AND(FS>0,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strFS & "," & strETC & "=0)", "BAD")
    'NS:AND(AS=0,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & "=0," & strETC & "=0)", "BAD")
    
    'NF:AND(NF>0,NF<=SD) -> COMPLETE
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAF & ")", "COMPLETE")
    'NF:AND(NF>0,NF>SD) -> GOOD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strNF & ">0," & strNF & ">STATUS_DATE)", "GOOD")
    'NF:AND(CF<=(SD+14),NF=0) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strCF & "<=(STATUS_DATE+14)," & strNF & "=0)", "NEUTRAL")
    'NF:AND(CF<=SD,NF=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strCF & "<=STATUS_DATE," & strNF & "=0)", "BAD") 'should have finished
    'NF:AND(AS,NF=0) -> INPUT
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAS & "," & strNF & "=0)", "INPUT") 'in progress
    'NF:AND(NF>0,NS>0,NF<NS) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strNF & ">0," & strNS & ">0," & strNS & ">" & strNF & ")", "BAD")
    'NF:AND(AF>0,NS=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAF & "," & strNS & "=0)", "BAD")
    'NF:AND(FF>0,EVP=1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strFF & "," & strEVP & "=1)", "BAD")
    'NF:AND(AF,EVP<1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAF & "," & strEVP & "<1)", "BAD")
    'NF:AND(NF=0,EVP=1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strNF & "=0," & strEVP & "=1)", "BAD")
    'NF:AND(FF>0,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strFF & "," & strETC & "=0)", "BAD")
    'NF:AND(AF>0,ETC>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAF & "," & strETC & ">0)", "BAD")
    'NF:AND(NF=0,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strNF & "=0," & strETC & "=0)", "BAD")
    
    'EVP:AND(FF,NEW EVP>EVP) -> GOOD
    'todo: keeping FF forces FF before all good
    'todo: remove FF; add last and add stop if true to isolate this good update
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strFF & "," & strEVP & ">" & strCEVP & ")", "GOOD")
    'EVP:AND(AF,EVP=1) -> GOOD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strAF & "," & strEVP & "=1)", "GOOD")
    'EVP:AND(FF,EVP=PREVIOUS) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strFF & "," & strEVP & "=" & strCEVP & ")", "NEUTRAL")
    'EVP:AND(AS,NF=0) -> INPUT
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strAS & "," & strNF & "=0)", "INPUT") 'in progress
    'EVP:AND(EVP>0,FS>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & ">0," & strFS & ")", "BAD")
    'EVP:AND(EVP=1,FF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "=1," & strFF & ")", "BAD")
    'EVP:AND(EVP=1,NF=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "=1," & strNF & "=0)", "BAD")
    'EVP:AND(EVP<1,AF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "<1," & strAF & ")", "BAD")
    'EVP:AND(EVP=1,ETC>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "=1," & strETC & ">0)", "BAD")
    'EVP:AND(EVP<1,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "<1," & strETC & "=0)", "BAD")
    'EVP:AND(EVP>0,AS=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & ">0," & strNS & "=0)", "BAD")
    'EVP:AND(EVP<PREVIOUS) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "<" & strCEVP & ")", "BAD")
    
    'ETC:AND(FF,NEW ETC<>ETC) -> GOOD
    'todo: keeping FF forces FF before all good
    'todo: remove FF; add last and add stop if true to isolate this good update
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strFF & "," & strETC & "<>" & strCETC & ")", "GOOD")
    'ETC:AND(AF,ETC=0) -> GOOD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strAF & "," & strETC & "=0," & strEVP & "=1)", "GOOD")
    'ETC:AND(FF,ETC=PREVIOUS) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strFF & "," & strETC & "=" & strCETC & ")", "NEUTRAL")
    'ETC:AND(FS,ETC=PREVIOUS) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strFS & "," & strETC & "=" & strCETC & ")", "NEUTRAL")
    If Not blnAssignments Then
      'ETC:AND(AS,NF=0) -> INPUT
      oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strAS & "," & strNF & "=0)", "INPUT") 'in progress
    End If
    'ETC:AND(ETC=0,FS>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strFS & ")", "BAD")
    'ETC:AND(ETC=0,FF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strFF & ")", "BAD")
    'ETC:AND(ETC>0,AF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & ">0," & strAF & ")", "BAD")
    'ETC:AND(ETC>0,EVP=1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & ">0," & strEVP & "=1)", "BAD")
    'ETC:AND(ETC=0,EVP<1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strEVP & "<1)", "BAD")
    'ETC:AND(ETC=0,EVP=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strEVP & "=0)", "BAD")
    'ETC:AND(ETC=0,AF=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strNF & "=0)", "BAD")
    'ETC:AND(ETC=0,AS=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strNS & "=0)", "BAD")
    
    If blnAssignments And Not oETCValidationRange Is Nothing Then
      'need to reset certain variables for Assignment ranges
      oETCValidationRange.Select
      Set oFirstCell = oWorksheet.Application.ActiveCell
      oFirstCell.Select
      strNS = oWorksheet.Cells(oFirstCell.Row, lngASCol).Address(False, True)
      strNF = oWorksheet.Cells(oFirstCell.Row, lngAFCol).Address(False, True)
      strEVP = oWorksheet.Cells(oFirstCell.Row, lngEVPCol).Address(False, True)
      strETC = oWorksheet.Cells(oFirstCell.Row, lngETCCol).Address(False, True)
      strCETC = oWorksheet.Cells(oFirstCell.Row, lngCETCCol).Address(False, True)
      'set up derived addresses for ease of formula writing
      'AS = (NS>0,NS<=SD)
      strAS = strNS & ">0," & strNS & "<=STATUS_DATE"
      'AF = (NF>0,NF<=SD)
      strAF = strNF & ">0," & strNF & "<=STATUS_DATE"
      'FS = (NS>0,NS>SD)
      strFS = strNS & ">0," & strNS & ">STATUS_DATE"
      'FF = (NF>0,NF>SD)
      strFF = strNF & ">0," & strNF & ">STATUS_DATE"
      'ETC:AND(FF,ETC=PREVIOUS) -> NEUTRAL (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strFF & "," & strETC & "=" & strCETC & ")", "NEUTRAL")
      'ETC:AND(FS,ETC=PREVIOUS) -> NEUTRAL (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strFS & "," & strETC & "=" & strCETC & ")", "NEUTRAL")
      'ETC:AND(FF,ETC=0) -> NEUTRAL (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strFF & "," & strETC & "=0)", "NEUTRAL")
      'ETC:AND(FS,ETC=0) -> NEUTRAL (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strFS & "," & strETC & "=0)", "NEUTRAL")
      'ETC:AND(AS,NF=0) -> INPUT (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strAS & "," & strNF & "=0)", "INPUT") 'in progress
      'ETC:AND(ETC>0,AF>0) -> BAD (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strETC & ">0," & strAF & ")", "BAD")
      'ETC:AND(ETC>0,EVP=1) -> BAD (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strETC & ">0," & strEVP & "=1)", "BAD")
    End If
    
    If blnMilestones Then 'assumes COBRA and field values = COBRA codes
      'todo: AS>0,EVT='E',EVP<>50
      'todo: oRecordset.AddNew Array(0,1),Array("NS", "=AND(" & strAS & "," & strEVT & "='E'," & strEVP & "<>.5)")
      'todo: AS>0,EVT='F',EVP>0
      'todo: oRecordset.AddNew Array(0,1),Array("NS", "=AND(" & strAS & "," & strEVT & "='F'," & strEVP & ">0)")
      'todo: AS>0,EVT='G',EVP<>1
      'todo: EVP<>.5,EVT='E",AS>0
      'todo: oRecordset.AddNew Array(0,1),Array("EVP", "=AND(" & strEVP & "<>.5," & strEVT & "='E'," & strAS & ">0)")
      'todo: EVP>0,EVT='F',AS>0
      'todo: oRecordset.AddNew Array(0,1),Array("EVP", "=AND(" & strEVP & ">0," & strEVT & "='F'," & strAS & ">0)")
      'todo: EVP<1,EVT='G',AS>0

    End If
    '</cpt-breadcrumbs:format-conditions>
skip_working:
    lngFormatCondition = 0
    With oRecordset
      'for the progress bar
      lngFormatConditions = .RecordCount
      .MoveFirst
      Do While Not .EOF
        'race is on
        lngFormatCondition = lngFormatCondition + 1
        If strItem <> "" Then
          myStatusSheet_frm.lblStatus.Caption = "Applying Conditional Formatting [" & strItem & "]...(" & Format(lngFormatCondition / lngFormatConditions, "0%") & ")"
          Application.StatusBar = "Applying Conditional Formatting [" & strItem & "]...(" & Format(lngFormatCondition / lngFormatConditions, "0%") & ")"
        Else
          myStatusSheet_frm.lblStatus.Caption = "Applying Conditional Formatting...(" & Format(lngFormatCondition / lngFormatConditions, "0%") & ")"
          Application.StatusBar = "Applying Conditional Formatting...(" & Format(lngFormatCondition / lngFormatConditions, "0%") & ")"
        End If
        myStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngFormatConditions) * myStatusSheet_frm.lblStatus.Width
        Set oFormatRange = oDict.item(CStr(.Fields(0)))
        oFormatRange.Select
        oFormatRange.FormatConditions.Add Type:=xlExpression, Formula1:=CStr(.Fields(1))
        oFormatRange.FormatConditions(oFormatRange.FormatConditions.Count).SetFirstPriority
        If .Fields(2) = "BAD" Then
          oFormatRange.FormatConditions(1).Font.Color = -16383844
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
          oFormatRange.FormatConditions(1).Interior.Color = 13551615
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
        ElseIf .Fields(2) = "CALCULATION" Then
          oFormatRange.FormatConditions(1).Font.Color = 32250
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = -4105
          oFormatRange.FormatConditions(1).Interior.Color = 15921906
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
        ElseIf .Fields(2) = "COMPLETE" Then
          oFormatRange.FormatConditions(1).Font.Color = 8355711
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = -4105
          oFormatRange.FormatConditions(1).Interior.Color = 15921906
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
        ElseIf .Fields(2) = "GOOD" Then
          oFormatRange.FormatConditions(1).Font.Color = -16752384
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
          oFormatRange.FormatConditions(1).Interior.Color = 13561798
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
        ElseIf .Fields(2) = "INPUT" Then
          oFormatRange.FormatConditions(1).Font.Color = 7749439
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = -4105
          oFormatRange.FormatConditions(1).Interior.Color = 10079487
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
          'oFormatRange.FormatConditions(1).BorderAround xlContinuous, xlThin, , Color:=RGB(127, 127, 127)
        ElseIf .Fields(2) = "NEUTRAL" Then
          oFormatRange.FormatConditions(1).Font.Color = -16754788
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
          oFormatRange.FormatConditions(1).Interior.Color = 10284031
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
          'oFormatRange.FormatConditions(1).BorderAround xlContinuous, xlThin, , Color:=RGB(127, 127, 127)
        End If
        oFormatRange.FormatConditions(1).StopIfTrue = False 'CBool(oRecordset(3))
        .MoveNext
      Loop
      'race is over - notify
      If strItem <> "" Then
        myStatusSheet_frm.lblStatus.Caption = "Applying Conditional Formatting [" & strItem & "]...done."
      Else
        myStatusSheet_frm.lblStatus.Caption = "Applying Conditional Formatting...done."
      End If
      myStatusSheet_frm.lblProgress.Width = myStatusSheet_frm.lblStatus.Width
      Application.StatusBar = myStatusSheet_frm.lblStatus.Caption
      oDict.RemoveAll
      .Close
    End With
  Else 'blnConditionalFormats=false
    If Not oInputRange Is Nothing Then
      oInputRange.Style = "Input"
    End If
    If Not oTwoWeekWindowRange Is Nothing Then
      oTwoWeekWindowRange.Style = "Neutral"
    End If
    If Not oUnlockedRange Is Nothing Then
      oUnlockedRange.Locked = False
    End If
  End If
  
exit_here:
  On Error Resume Next
  Set oAssignment = Nothing
  Set oFormatRange = Nothing
  Set oDict = Nothing
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  Set oFirstCell = Nothing
  Set oETCRange = Nothing
  Set oEVPRange = Nothing
  Set oNFRange = Nothing
  Set oNSRange = Nothing
  Set oComment = Nothing
  Set oUnlockedRange = Nothing
  Set oAssignmentETCRange = Nothing
  Set oComment = Nothing
  Set oEVTRange = Nothing
  Set oCompleted = Nothing
  Set oMilestoneRange = Nothing
  Set oClearRange = Nothing
  Set oSummaryRange = Nothing
  Set oNumberValidationRange = Nothing
  Set oDateValidationRange = Nothing
  Set oTwoWeekWindowRange = Nothing
  Set oInputRange = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptCopyData", Err, Erl)
  Resume exit_here
End Sub

Private Sub cptGetAssignmentData(ByRef myStatusSheet_frm As cptStatusSheet_frm, ByRef oTask As MSProject.Task, ByRef oWorksheet As Excel.Worksheet, lngRow As Long, lngHeaderRow As Long, lngNameCol As Long, lngRemainingWorkCol As Long)
  'objects
  Dim oAssignment As Assignment
  'strings
  Dim strAllowAssignmentNotes As String
  Dim strProtect As String
  Dim strDataValidation As String
  'longs
  Dim lngEVPCol As Long
  Dim lngEVTCol  As Long
  Dim lngNFCol As Long
  Dim lngNSCol As Long
  Dim lngFFCol As Long
  Dim lngFSCol As Long
  Dim lngBaselineCostCol As Long
  Dim lngBaselineWorkCol As Long
  Dim lngIndent As Long
  Dim lngItem As Long
  Dim lngLastCol As Long
  Dim lngLastRow As Long
  Dim lngCol As Long
  'integers
  'doubles
  'booleans
  Dim blnAllowAssignmentNotes As Boolean
  'variants
  Dim vCol As Variant
  Dim vAssignment As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  lngIndent = Len(cptRegEx(oWorksheet.Cells(lngRow, lngNameCol).Value, "^\s*"))
  lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
  lngLastRow = oWorksheet.Cells(1048576, 1).End(xlUp).Row
  'get column for FS,FF,NS,NF,EVT,EVP
  lngFSCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Forecast Start", lookat:=xlWhole).Column
  lngFFCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Forecast Finish", lookat:=xlWhole).Column
  lngNSCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Actual Start", lookat:=xlPart).Column
  lngNFCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Actual Finish", lookat:=xlPart).Column
  'todo: lngEVTCol = oWorksheet.Rows(lngHeaderRow).Find(what:="EVT", lookat:=xlWhole).Column - Milestone EVT?
  lngEVPCol = oWorksheet.Rows(lngHeaderRow).Find(what:="New EV%", lookat:=xlWhole).Column
  
  lngItem = 0
  For Each oAssignment In oTask.Assignments
    lngItem = lngItem + 1
    'if custom field used to filter the ims is not filled down to assignment...problems.
    'if the next row is NOT the assignment, then make room for it
    If (oWorksheet.Cells(lngRow + lngItem, 1).Value Mod 4194304) <> (oAssignment.UniqueID - (2 ^ 20)) Then
      oWorksheet.Rows(lngRow + lngItem).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
      oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Font.ColorIndex = xlAutomatic
    Else 'clear it out and rebuild
      oWorksheet.Rows(lngRow + lngItem).Value = ""
    End If
    'this fills down task custom fields to assignments
    For lngCol = 2 To lngNameCol
      If lngCol <> lngNameCol Then oWorksheet.Cells(lngRow + lngItem, lngCol) = oWorksheet.Cells(lngRow, lngCol)
    Next lngCol
    
    oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Font.Italic = True
    vAssignment = oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Value
    vAssignment(1, 1) = oAssignment.UniqueID 'import assumes this is oAssignment.UniqueID
    vAssignment(1, lngNameCol) = String(lngIndent + 3, " ") & oAssignment.ResourceName
    If oAssignment.ResourceType = pjWork Then
      lngBaselineWorkCol = oWorksheet.Rows(lngHeaderRow).Find("Baseline Work", lookat:=xlWhole).Column
      vAssignment(1, lngBaselineWorkCol) = oAssignment.BaselineWork / 60
      vAssignment(1, lngRemainingWorkCol) = oAssignment.RemainingWork / 60
      vAssignment(1, lngRemainingWorkCol + 1) = oAssignment.RemainingWork / 60
    Else
      lngBaselineCostCol = oWorksheet.Rows(lngHeaderRow).Find("Baseline Work", lookat:=xlWhole).Column
      vAssignment(1, lngBaselineCostCol) = oAssignment.BaselineCost
      vAssignment(1, lngRemainingWorkCol) = oAssignment.RemainingCost
      vAssignment(1, lngRemainingWorkCol + 1) = oAssignment.RemainingCost
    End If

    'fill down NS,NF,EVP
    For Each vCol In Array(lngFSCol, lngFFCol, lngNSCol, lngNFCol, lngEVPCol)
      vAssignment(1, vCol) = "=" & oWorksheet.Cells(lngRow, vCol).AddressLocal(False, True)
      oWorksheet.Cells(lngRow + lngItem, vCol).Font.ThemeColor = xlThemeColorDark1
      oWorksheet.Cells(lngRow + lngItem, vCol).Font.TintAndShade = -4.99893185216834E-02
      If vCol = lngEVPCol Then oWorksheet.Cells(lngRow + lngItem, lngEVPCol).NumberFormat = "0%"
    Next vCol

    'add validation
    If oETCValidationRange Is Nothing Then
      Set oETCValidationRange = oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1)
    Else
      Set oETCValidationRange = oWorksheet.Application.Union(oETCValidationRange, oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1))
    End If
    'allow input on ETC if task is unstarted or incomplete - i.e., in progress
    If (Not IsDate(oTask.ActualStart) And oTask.Start <= ActiveProject.StatusDate) Or (IsDate(oTask.ActualStart) And Not IsDate(oTask.ActualFinish)) Then
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1))
      End If
    End If
    'add protection
    If oUnlockedRange Is Nothing Then
      Set oUnlockedRange = oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1)
    Else
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1))
    End If
    
    'export assignment notes
    If myStatusSheet_frm.chkExportNotes And Len(oAssignment.Notes) > 0 Then
      vAssignment(1, lngLastCol) = Trim(Replace(oAssignment.Notes, vbCr, vbLf))
    End If
    'allow notes at the assignment level?
    strAllowAssignmentNotes = cptGetSetting("StatusSheet", "chkAllowAssignmentNotes")
    If strAllowAssignmentNotes <> "" Then
      blnAllowAssignmentNotes = CBool(strAllowAssignmentNotes)
    Else
      blnAllowAssignmentNotes = False
    End If
    If blnAllowAssignmentNotes Then
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow + lngItem, lngLastCol))
    End If
    oWorksheet.Cells(lngRow + lngItem, lngLastCol).HorizontalAlignment = xlLeft
    oWorksheet.Cells(lngRow + lngItem, lngLastCol).NumberFormat = "General"
    oWorksheet.Cells(lngRow + lngItem, lngLastCol).WrapText = True
    
    'enter the values
    oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Value = vAssignment
    
    If oAssignmentRange Is Nothing Then
      Set oAssignmentRange = oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol))
    Else
      Set oAssignmentRange = oWorksheet.Application.Union(oAssignmentRange, oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)))
    End If
    oWorksheet.Rows(lngRow + lngItem).AutoFit
  Next oAssignment
  'add formulae
  If oTask.Assignments.Count > 0 Then
    'baseline work
    oWorksheet.Cells(lngRow, lngRemainingWorkCol - 1).FormulaR1C1 = "=SUM(R" & lngRow + 1 & "C" & lngRemainingWorkCol - 1 & ":R" & lngRow + lngItem & "C" & lngRemainingWorkCol - 1 & ")"
    'prev etc
    oWorksheet.Cells(lngRow, lngRemainingWorkCol).FormulaR1C1 = "=SUM(R" & lngRow + 1 & "C" & lngRemainingWorkCol & ":R" & lngRow + lngItem & "C" & lngRemainingWorkCol & ")"
    'new etc
    oWorksheet.Cells(lngRow, lngRemainingWorkCol + 1).FormulaR1C1 = "=SUM(R" & lngRow + 1 & "C" & lngRemainingWorkCol + 1 & ":R" & lngRow + lngItem & "C" & lngRemainingWorkCol + 1 & ")"
  End If

exit_here:
  On Error Resume Next
  Set oAssignment = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptGetAssignmentData", Err, Erl)
  Resume exit_here
End Sub

Sub cptFinalFormats(ByRef oWorksheet As Excel.Worksheet)
  Dim lngHeaderRow As Long
  Dim lngLastRow As Long
  Dim lngNameCol As Long
  Dim vBorder As Variant
  
  lngHeaderRow = 8
  oWorksheet.Cells(lngHeaderRow, 1).AutoFilter
  oWorksheet.Columns(1).AutoFit
  lngNameCol = WorksheetFunction.Match("Task Name / Scope", oWorksheet.Rows(lngHeaderRow), 0)
  lngLastRow = oWorksheet.Cells(1048576, lngNameCol).End(xlUp).Row
  oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, lngNameCol), oWorksheet.Cells(lngLastRow, lngNameCol)).Columns.AutoFit
  With oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight), oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp))
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    For Each vBorder In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
      With .Borders(vBorder)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
      End With
    Next vBorder
  End With
  oWorksheet.Application.WindowState = xlNormal 'cannot apply certain settings below if window is minimized...like data validation
  oWorksheet.Application.Calculation = xlCalculationAutomatic
  oWorksheet.Application.ScreenUpdating = True
  oWorksheet.Application.ActiveWindow.DisplayGridlines = False
  oWorksheet.[B1].Select
  oWorksheet.Application.ActiveWindow.SplitRow = 8
  oWorksheet.Application.ActiveWindow.SplitColumn = 0
  'note: if user's Excel 'normal' window size is puny then FreezePanes might fail;
  'note: have them adjust manually - alternative is to change the line above to xlMaximized
  'note: and that's a terrible UX with all the screen flashes
  oWorksheet.Application.ActiveWindow.FreezePanes = True
  oWorksheet.Application.ActiveWindow.DisplayHorizontalScrollBar = True
  oWorksheet.Application.ActiveWindow.DisplayVerticalScrollBar = True
  oWorksheet.Application.WindowState = xlMinimized
End Sub

Sub cptListQuickParts(ByRef myStatusSheet_frm As cptStatusSheet_frm, Optional blnRefreshOutlook As Boolean = False)
  'objects
  Dim oOutlook As Object 'Outlook.Application
  Dim oMailItem As Object 'Outlook.MailItem
  Dim oDocument As Word.Document
  Dim oWord As Word.Application
  Dim oTemplate As Word.Template
  Dim oBuildingBlockEntries As Word.BuildingBlockEntries
  Dim oBuildingBlock As Word.BuildingBlock
  'strings
  Dim strQuickPartList As String
  Dim strSQL As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vQuickPart As Variant
  Dim vQuickParts As Variant
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If blnRefreshOutlook Then
    'refresh QuickParts in Outlook
    myStatusSheet_frm.cboQuickParts.Clear
    'get Outlook
    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application") 'this works even if Outlook isn't open
    If oOutlook Is Nothing Then
      Set oOutlook = CreateObject("Outlook.Application")
    End If
    If oOutlook Is Nothing Then GoTo exit_here
    'create MailItem, insert quickparts, update links, dates
    Set oMailItem = oOutlook.CreateItem(0) '0=olMailItem
    If oMailItem.BodyFormat <> 2 Then oMailItem.BodyFormat = 2 '2=olFormatHTML
    If Err.Number > 0 Then
      MsgBox "Outlook QuickParts are inaccessible.", vbExclamation + vbOKOnly, "Blocked"
      myStatusSheet_frm.cboQuickParts.AddItem "[blocked]"
      myStatusSheet_frm.cboQuickParts.Value = "[blocked]"
      myStatusSheet_frm.cboQuickParts.Enabled = False
      Err.Clear
      GoTo exit_here
    End If
    Set oDocument = oMailItem.GetInspector.WordEditor
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oDocument Is Nothing Then
      'try again with MailItem displayed
      oMailItem.Display False
      On Error Resume Next
      Set oDocument = oMailItem.GetInspector.WordEditor
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oDocument Is Nothing Then
        'todo: try again by accessing Word directly
        MsgBox "Outlook QuickParts are inaccessible.", vbExclamation + vbOKOnly, "Blocked"
        myStatusSheet_frm.cboQuickParts.AddItem "[blocked]"
        myStatusSheet_frm.cboQuickParts.Value = "[blocked]"
        myStatusSheet_frm.cboQuickParts.Enabled = False
        oMailItem.Close 1 'SaveMode:=1 (olDiscard)
        GoTo exit_here
      Else
        oMailItem.GetInspector.WindowState = 1 '1=olMinimized
      End If
    End If
    Set oWord = oDocument.Application
    Set oTemplate = oWord.Templates(1)
    Set oBuildingBlockEntries = oTemplate.BuildingBlockEntries
    'loop through them
    For lngItem = 1 To oBuildingBlockEntries.Count
      Set oBuildingBlock = oBuildingBlockEntries(lngItem)
      If oBuildingBlock.Type.Name = "Quick Parts" Then
        strQuickPartList = strQuickPartList & oBuildingBlock.Name & ","
      End If
    Next lngItem
    'sort them
    If Len(strQuickPartList) > 0 Then
      strQuickPartList = Left(strQuickPartList, Len(strQuickPartList) - 1)
      vQuickParts = Split(strQuickPartList, ",")
      cptQuickSort vQuickParts, 0, UBound(vQuickParts)
      For Each vQuickPart In vQuickParts
        myStatusSheet_frm.cboQuickParts.AddItem vQuickPart
      Next vQuickPart
    End If
    oMailItem.Close 1 'SaveMode:=1 (olDiscard)
  End If
    
exit_here:
  On Error Resume Next
  Set oWord = Nothing
  Set oOutlook = Nothing
  Set oMailItem = Nothing
  Set oDocument = Nothing
  Set oWord = Nothing
  Set oTemplate = Nothing
  Set oBuildingBlockEntries = Nothing
  Set oBuildingBlock = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptListQuickParts", Err)
  Resume exit_here
End Sub

Function cptSaveStatusSheet(ByRef myStatusSheet_frm As cptStatusSheet_frm, ByRef oWorkbook As Excel.Workbook, Optional strItem As String) As String
  'objects
  'strings
  Dim strValidPath As String
  Dim strMsg As String
  Dim strFileName As String
  Dim strDir As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  dtStatus = ActiveProject.StatusDate

  With myStatusSheet_frm
    strDir = .lblDirSample.Caption
    'create the status date directory
    If Dir(strDir, vbDirectory) = vbNullString Then
      MkDir strDir
      oWorkbook.Application.Wait Now + TimeValue("00:00:03")
    End If
    strFileName = .txtFileName.Value & ".xlsx"
    strFileName = Replace(strFileName, "[yyyy-mm-dd]", Format(dtStatus, "yyyy-mm-dd"))
    strFileName = Replace(strFileName, "[program]", cptGetProgramAcronym)
    If Len(strItem) > 0 Then
      strFileName = Replace(strFileName, "[item]", strItem)
    End If
    strFileName = cptRemoveIllegalCharacters(strFileName, "-")
    If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
    strValidPath = cptValidPath(strDir & strFileName)
    If CBool(Split(strValidPath, ":")(0)) = False Then
      If InStr(strValidPath, "exceeds") > 0 Then
        strDir = cptGetShortPath(strDir)
      Else
        MsgBox "There was an error saving this workbook:" & Replace(strValidPath, "0:", ""), vbExclamation + vbOKOnly, "Path/Filename error"
        cptSaveStatusSheet = ""
        GoTo exit_here
      End If
    End If
    On Error Resume Next
    If Dir(strDir & strFileName) <> vbNullString Then
      Kill strDir & strFileName
      oWorkbook.Application.Wait Now + TimeValue("00:00:02")
      DoEvents
    End If
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    'account for if the file exists and is open in the background
    If Dir(strDir & strFileName) <> vbNullString Then  'delete failed, rename with timestamp
      strMsg = "'" & strFileName & "' already exists, and is likely open." & vbCrLf
      strFileName = Replace(strFileName, ".xlsx", "_" & Format(Now, "hh-nn-ss") & ".xlsx")
      strMsg = strMsg & vbCrLf & "The file you are now creating will be named '" & strFileName & "'"
      MsgBox strMsg, vbExclamation + vbOKOnly, "NOTA BENE"
      oWorkbook.SaveAs strDir & strFileName, 51
      oWorkbook.Application.Wait Now + TimeValue("00:00:02")
    Else
      oWorkbook.SaveAs strDir & strFileName, 51
      oWorkbook.Application.Wait Now + TimeValue("00:00:02")
    End If
  End With

  cptSaveStatusSheet = strDir & strFileName

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptSaveStatusSheet", Err, Erl)
  Resume exit_here
End Function

Sub cptSendStatusSheet(ByRef myStatusSheet_frm As cptStatusSheet_frm, strFullName As String, Optional strItem As String)
  'objects
  Dim oOutlook As Object 'Outlook.Application
  Dim oInspector As Object 'Outlook.Inspector
  Dim oMailItem As Object 'Outlook.MailItem
  Dim oWord As Word.Application
  Dim oDocument As Word.Document
  Dim oBuildingBlock As Word.BuildingBlock
  Dim oSelection As Word.Selection
  Dim oEmailTemplate As Word.Template
  'strings
  Dim strTempItem As String
  Dim strSubject As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  On Error Resume Next
  Set oOutlook = GetObject(, "Outlook.Application")
  If oOutlook Is Nothing Then
    Set oOutlook = CreateObject("Outlook.Application")
  End If
  If oOutlook Is Nothing Then GoTo exit_here
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Set oMailItem = oOutlook.CreateItem(0) '0=olMailItem
  oMailItem.Display False
  oMailItem.Attachments.Add strFullName
  With myStatusSheet_frm
    strSubject = .txtSubject
    strSubject = Replace(strSubject, cptRegEx(strSubject, "\[status\_date\]"), FormatDateTime(ActiveProject.StatusDate, vbShortDate))
    strSubject = Replace(strSubject, cptRegEx(strSubject, "\[yyyy\-mm\-dd\]"), Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
    strSubject = Replace(strSubject, cptRegEx(strSubject, "\[item\]"), strItem)
    strSubject = Replace(strSubject, cptRegEx(strSubject, "\[program\]"), cptGetProgramAcronym)
    oMailItem.Subject = strSubject
    oMailItem.CC = .txtCC
    If oMailItem.BodyFormat <> 2 Then oMailItem.BodyFormat = 2 '2=olFormatHTML
    If Not IsNull(.cboQuickParts.Value) And .cboQuickParts.Enabled Then
      If Len(.cboQuickParts.Value) = 0 Then GoTo skip_QuickPart
      Set oDocument = oMailItem.GetInspector.WordEditor
      Set oWord = oDocument.Application
      Set oSelection = oDocument.Windows(1).Selection
      Set oEmailTemplate = oWord.Templates(1)
      On Error Resume Next
      Set oBuildingBlock = oEmailTemplate.BuildingBlockEntries(.cboQuickParts)
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oBuildingBlock Is Nothing Then
        MsgBox "Quick Part '" & .cboQuickParts & "' not found!", vbExclamation + vbOKOnly, "Missing Quick Part"
      Else
        oBuildingBlock.Insert oSelection.Range, True
      End If
      'only do replacements if QuickPart is used
      'clean date format
      strTempItem = cptRegEx(oMailItem.HTMLBody, "\[(Y|y){1,}-(M|m){1,}-(D|d){1,}\]")
      If Len(strTempItem) > 0 Then
        oMailItem.HTMLBody = Replace(oMailItem.HTMLBody, strTempItem, Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
      End If
      'clean status date
      strTempItem = cptRegEx(oMailItem.HTMLBody, "\[(S|s)(T|t)(A|a)(T|t)(U|u)(S|s).(D|d)(A|a)(T|t)(E|e)\]")
      If Len(strTempItem) > 0 Then
        oMailItem.HTMLBody = Replace(oMailItem.HTMLBody, strTempItem, FormatDateTime(ActiveProject.StatusDate, vbShortDate))
      End If
      'clean program
      strTempItem = cptRegEx(oMailItem.HTMLBody, "\[(P|p)(R|r)(O|o)(G|g)(R|r)(A|a)(M|m)\]")
      If Len(strTempItem) > 0 Then
        oMailItem.HTMLBody = Replace(oMailItem.HTMLBody, strTempItem, cptGetProgramAcronym)
      End If
      'clean item
      strTempItem = cptRegEx(oMailItem.HTMLBody, "\[(I|i)(T|t)(E|e)(M|m)\]")
      If Len(strTempItem) > 0 Then
        oMailItem.HTMLBody = Replace(oMailItem.HTMLBody, strTempItem, strItem)
      End If
    End If
skip_QuickPart:
    On Error Resume Next
    Set oInspector = oMailItem.GetInspector
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oInspector Is Nothing Then
      oInspector.WindowState = 1 '1=olMinimized
    End If
      
  End With
  
exit_here:
  On Error Resume Next
  Set oInspector = Nothing
  Set oBuildingBlock = Nothing
  Set oOutlook = Nothing
  Set oMailItem = Nothing
  Set oDocument = Nothing
  Set oWord = Nothing
  Set oSelection = Nothing
  Set oEmailTemplate = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptSendStatusSheet", Err, Erl)
  Resume exit_here

End Sub

Sub cptSaveStatusSheetSettings(ByRef myStatusSheet_frm As cptStatusSheet_frm)
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFileName As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
    
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  With myStatusSheet_frm
    'save settings
    cptDeleteSetting "StatusSheet", "cboEVP" 'moved to Integration
    cptSaveSetting "StatusSheet", "cboCostTool", .cboCostTool.Value
    cptDeleteSetting "StatusSheet", "cboEVT" 'moved to Integration
    cptSaveSetting "StatusSheet", "chkHide", IIf(.chkHide, 1, 0)
    If Not IsNull(.cboCreate) Then
      cptSaveSetting "StatusSheet", "cboCreate", .cboCreate
    End If
    cptSaveSetting "StatusSheet", "txtDir", .txtDir
    cptSaveSetting "StatusSheet", "chkAppendStatusDate", IIf(.chkAppendStatusDate, 1, 0)
    If .cboEach.Value <> "" Then
      cptSaveSetting "StatusSheet", "cboEach", .cboEach.Value
    Else
      cptSaveSetting "StatusSheet", "cboEach", ""
    End If
    cptSaveSetting "StatusSheet", "txtFileName", .txtFileName
    cptSaveSetting "StatusSheet", "chkAllItems", IIf(.chkAllItems, 1, 0)
    cptSaveSetting "StatusSheet", "chkDataValidation", IIf(.chkValidation, 1, 0)
    cptSaveSetting "StatusSheet", "chkProtect", IIf(.chkProtect, 1, 0)
    cptDeleteSetting "StatusSheet", "chkLocked"
    cptSaveSetting "StatusSheet", "chkAssignments", IIf(.chkAssignments, 1, 0)
    cptSaveSetting "StatusSheet", "chkConditionalFormatting", IIf(.chkConditionalFormatting, 1, 0)
    cptSaveSetting "StatusSheet", "chkConditionalFormattingLegend", IIf(.chkConditionalFormattingLegend, 1, 0)
    cptSaveSetting "StatusSheet", "chkEmail", IIf(.chkSendEmails, 1, 0)
    If .chkSendEmails Then
      cptSaveSetting "StatusSheet", "txtSubject", .txtSubject
      cptSaveSetting "StatusSheet", "txtCC", .txtCC
      If Not IsNull(.cboQuickParts.Value) Then
        If Not .cboCostTool.Value = "[blocked]" Then
          cptSaveSetting "StatusSheet", "cboQuickPart", .cboQuickParts.Value
        Else
          cptDeleteSetting "StatusSheet", "cboQuickPart"
        End If
      End If
    End If
    If Len(.txtNotesColTitle.Value) > 0 Then
      cptSaveSetting "StatusSheet", "txtNotesColTitle", .txtNotesColTitle.Value
    Else
      cptSaveSetting "StatusSheet", "txtNotesColTitle", "Reason / Action / Impact"
    End If
    cptSaveSetting "StatusSheet", "chkExportNotes", IIf(.chkExportNotes, 1, 0)
    cptSaveSetting "StatusSheet", "chkAllowAssignmentNotes", IIf(.chkAllowAssignmentNotes, 1, 0)
    cptSaveSetting "StatusSheet", "chkKeepOpen", IIf(.chkKeepOpen, 1, 0)
    cptSaveSetting "StatusSheet", "chkConditionalFormatting", IIf(.chkConditionalFormatting, 1, 0)
    cptSaveSetting "StatusSheet", "chkLookahead", IIf(.chkLookahead, 1, 0)
    If .chkLookahead And Len(.txtLookaheadDays) > 0 Then
      cptSaveSetting "StatusSheet", "txtLookaheadDays", CLng(.txtLookaheadDays.Value)
    End If
    cptSaveSetting "StatusSheet", "chkIgnoreLOE", IIf(.chkIgnoreLOE, 1, 0)
    'save user fields - overwrite
    strFileName = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
    Set oRecordset = CreateObject("ADODB.Recordset")
    oRecordset.Fields.Append "Field Constant", adInteger
    oRecordset.Fields.Append "Custom Field Name", adVarChar, 255
    oRecordset.Fields.Append "Local Field Name", adVarChar, 100
    oRecordset.Open
    If .lboExport.ListCount > 0 Then
      For lngItem = 0 To .lboExport.ListCount - 1
        oRecordset.AddNew Array(0, 1, 2), Array(.lboExport.List(lngItem, 0), .lboExport.List(lngItem, 1), .lboExport.List(lngItem, 2))
      Next lngItem
    End If
    If Dir(strFileName) <> vbNullString Then Kill strFileName
    oRecordset.Save strFileName, adPersistADTG
    oRecordset.Close

  End With

exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptSaveStatusSheetSettings", Err, Erl)
  Resume exit_here
  
End Sub


Sub cptAdvanceStatusDate()
  'objects
  Dim oMaster As MSProject.Project
  Dim oSubproject As MSProject.SubProject
  'strings
  Dim strMsg As String
  'longs
  Dim lngCount As Long
  Dim lngMismatch As Long
  Dim lngFixed As Long
  'integers
  'doubles
  'booleans
  Dim blnMatch As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Application.ChangeStatusDate
  
  lngCount = ActiveProject.Subprojects.Count
  If lngCount > 0 Then
    Set oMaster = ActiveProject
    If MsgBox("Would you like to review/edit Subproject Status Dates also?", vbQuestion + vbYesNo, "Status Date FlowDown") = vbYes Then
      Application.StatusBar = "Checking Subproject Status Dates...(" & Format(0 / lngCount, "0%") & ")"
      blnMatch = True 'default
      strMsg = "[ok] " & FormatDateTime(oMaster.StatusDate, vbShortDate) & " - " & oMaster.ProjectSummaryTask.Name & vbCrLf
      For Each oSubproject In oMaster.Subprojects
        If IsDate(oSubproject.SourceProject.StatusDate) Then
          If oSubproject.SourceProject.StatusDate <> oMaster.StatusDate Then
            strMsg = strMsg & vbCrLf & "<!> " & FormatDateTime(oSubproject.SourceProject.StatusDate, vbShortDate) & " - " & oSubproject.InsertedProjectSummary.Name
            blnMatch = False
            lngMismatch = lngMismatch + 1
          Else
            strMsg = strMsg & vbCrLf & "[ok] " & FormatDateTime(oSubproject.SourceProject.StatusDate, vbShortDate) & " - " & oSubproject.InsertedProjectSummary.Name
          End If
        Else
          strMsg = strMsg & vbCrLf & "<!> [No Status Date] - " & oSubproject.InsertedProjectSummary.Name
          blnMatch = False
          lngMismatch = lngMismatch + 1
        End If
        Application.StatusBar = "Checking Subproject Status Dates...(" & Format(oSubproject.Index / lngCount, "0%") & ")"
      Next oSubproject
      If blnMatch Then
        MsgBox strMsg & vbCrLf & vbCrLf & "Subproject Status Dates match!", vbInformation + vbOKOnly, "Status Date FlowDown"
      Else
        strMsg = strMsg & vbCrLf & vbCrLf & "Subproject Status Dates do not match. Flow Down now?"
        If Len(strMsg) > 255 Then
          strMsg = Format(lngMismatch, "#,##0") & " SubProject Status Dates do not match. Flow Down now?"
        End If
        If MsgBox(strMsg, vbExclamation + vbYesNo, "Status Date FlowDown") = vbYes Then
          Application.OpenUndoTransaction "cpt - Status Date Flow Down"
          strMsg = "[ok] " & FormatDateTime(oMaster.StatusDate, vbShortDate) & " - " & oMaster.ProjectSummaryTask.Name & vbCrLf
          For Each oSubproject In ActiveProject.Subprojects
            blnMatch = True
            If Not IsDate(oSubproject.SourceProject.StatusDate) Then
              lngFixed = lngFixed + 1
              blnMatch = False
            Else
              If oSubproject.SourceProject.StatusDate <> oMaster.StatusDate Then
                lngFixed = lngFixed + 1
                blnMatch = False
              End If
            End If
            If Not blnMatch Then
              oSubproject.SourceProject.StatusDate = oMaster.StatusDate
            End If
            Application.StatusBar = "Flowing Down Status Dates...(" & Format(oSubproject.Index / lngCount, "0%") & ")"
            strMsg = strMsg & vbCrLf & "[ok] " & FormatDateTime(oSubproject.SourceProject.StatusDate, vbShortDate) & " - " & oSubproject.InsertedProjectSummary.Name
          Next oSubproject
          strMsg = strMsg & "Complete."
          If Len(strMsg) > 255 Then
            MsgBox Format(lngFixed, "#,##0") & " SubProject Status Dates aligned.", vbInformation + vbOKOnly, "Status Date FlowDown"
          Else
            MsgBox strMsg, vbInformation + vbOKOnly, "Status Date FlowDown"
          End If
        End If
      End If
    End If
  End If
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Application.StatusBar = ""
  Set oMaster = Nothing
  Set oSubproject = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptAdvanceStatusDate", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptCaptureJournal()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strProgram As String
  Dim strFileName As String
  'longs
  Dim lngTask As Long
  Dim lngTasks As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtStatus As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strProgram = cptGetProgramAcronym
  
  dtStatus = FormatDateTime(ActiveProject.StatusDate, vbGeneralDate)
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  
  strFileName = cptDir & "\settings\cpt-journal.adtg"
  If Dir(strFileName) = vbNullString Then
    With oRecordset
      .Fields.Append "PROGRAM", adVarChar, 50
      .Fields.Append "STATUS_DATE", adDate
      .Fields.Append "TASK_UID", adInteger
      .Fields.Append "TASK_NOTE", adVarChar, 255
      .Fields.Append "ASSIGNMENT_UID", adInteger
      .Fields.Append "ASSIGNMENT_NOTE", adVarChar, 255
      .Open
    End With
  Else
    oRecordset.Open strFileName
  End If
  Dim oTask As MSProject.Task, oTasks As MSProject.Tasks
  Set oTasks = ActiveProject.Tasks
  lngTasks = oTasks.Count
  lngTask = 0
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Len(oTask.Notes) > 0 Then
      oRecordset.AddNew Array(0, 1, 2, 3), Array(strProgram, dtStatus, oTask.UniqueID, Chr(34) & oTask.Notes & Chr(34))
    End If
    Dim oAssignment As Assignment
    For Each oAssignment In oTask.Assignments
      If Len(oAssignment.Notes) > 0 Then
        oRecordset.AddNew Array(0, 1, 2, 3, 4, 5), Array(strProgram, dtStatus, oTask.UniqueID, oTask.Notes, oAssignment.UniqueID, oAssignment.Notes)
      End If
    Next
next_task:
    lngTask = lngTask + 1
    Debug.Print Format(lngTask / lngTasks, "0%")
  Next oTask
  
  oRecordset.Save strFileName
  oRecordset.Close
  
exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptCaptureJournal", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCompletedWork()
  'objects
  Dim oCDP As DocumentProperty
  Dim oAssignment As Assignment
  Dim oWorksheet As Object 'Excel.Worksheet
  Dim oWorkbook As Object 'Excel.Workbook
  Dim oExcel As Object 'Excel.Application
  Dim oRecordset As Object 'ADODB.Recordset
  Dim oTask As MSProject.Task
  'strings
  Dim strSetting As String
  Dim strCA As String
  Dim strEVP As String
  Dim strEVT As String
  Dim strLC As String
  Dim strWPM As String
  Dim strWP As String
  Dim strCAM As String
  Dim strOBS As String
  Dim strWBS As String
  Dim strProgram As String
  Dim strRecord As String
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  Dim strFileName As String
  'longs
  Dim lngWPMCol as Long
  Dim lngEVPCol As Long
  Dim lngCA As Long
  Dim lngLC As Long
  Dim lngEVP As Long
  Dim lngEVT As Long
  Dim lngItem As Long
  Dim lngWPM As Long
  Dim lngWP As Long
  Dim lngCAM As Long
  Dim lngOBS As Long
  Dim lngWBS As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  Dim blnHasWPM As Boolean
  Dim blnErrorTrapping As Boolean
  Dim blnMissing As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  Dim dtAF As Date
  
  If Not cptValidMap("WBS,OBS,CA,CAM,WP,[WPM],EVT,EVP", False, False, True) Then 'todo: WPM is not really required...
    MsgBox "Settings required. Exiting.", vbExclamation + vbOKOnly, "Invalid Settings"
    GoTo exit_here
  End If
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  lngWBS = Split(cptGetSetting("Integration", "WBS"), "|")(0)
  strWBS = CustomFieldGetName(lngWBS)
  lngOBS = Split(cptGetSetting("Integration", "OBS"), "|")(0)
  strOBS = CustomFieldGetName(lngOBS)
  lngCA = Split(cptGetSetting("Integration", "CA"), "|")(0)
  strCA = CustomFieldGetName(lngCA)
  lngCAM = Split(cptGetSetting("Integration", "CAM"), "|")(0)
  strCAM = CustomFieldGetName(lngCAM)
  lngWP = Split(cptGetSetting("Integration", "WP"), "|")(0)
  strWP = CustomFieldGetName(lngWP)
  strSetting = cptGetSetting("Integration", "WPM")
  blnHasWPM = Len(strSetting) > 0
  If blnHasWPM Then
    lngWPM = Split(cptGetSetting("Integration", "WPM"), "|")(0)
    strWPM = CustomFieldGetName(lngWPM)
  End If
  On Error Resume Next
  Set oCDP = ActiveProject.CustomDocumentProperties("fResID")
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oCDP Is Nothing Then
    strLC = ActiveProject.CustomDocumentProperties("fResID")
    lngLC = FieldNameToFieldConstant(strLC, pjResource)
  End If
  lngEVT = Split(cptGetSetting("Integration", "EVT"), "|")(0)
  strEVT = CustomFieldGetName(lngEVT)
  lngEVP = Split(cptGetSetting("Integration", "EVP"), "|")(0)
  strEVP = CustomFieldGetName(lngEVP)
  
  'create Schema
  strFileName = Environ("tmp") & "\Schema.ini"
  lngFile = FreeFile
  Open strFileName For Output As #lngFile
  Print #lngFile, "[wp.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=UID Long"
  Print #lngFile, "Col2=WBS Text"
  Print #lngFile, "Col3=OBS Text"
  Print #lngFile, "Col4=CA Text"
  Print #lngFile, "Col5=CAM Text"
  Print #lngFile, "Col6=WP Text"
  Print #lngFile, "Col7=WPM Text"
  Print #lngFile, "Col8=LC Text"
  Print #lngFile, "Col9=AF DateTime"
  Print #lngFile, "Col10=PercentComplete Long"
  Close #lngFile
  
  strFileName = Environ("tmp") & "\wp.csv"
  lngFile = FreeFile
  Open strFileName For Output As #lngFile
  Print #lngFile, "UID,WBS,OBS,CA,CAM,WP,WPM,LC,AF,PercentComplete,"
  
  lngTasks = ActiveProject.Tasks.Count
  If ActiveProject.Subprojects.Count > 0 Then
    ActiveWindow.TopPane.Activate
    If ActiveWindow.ActivePane.View.Type <> pjTaskItem Then ViewApply "Gantt Chart"
    OptionsViewEx DisplaySummaryTasks:=True
    OutlineShowTasks pjTaskOutlineShowLevel2
  End If
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    For Each oAssignment In oTask.Assignments
      strRecord = oTask.UniqueID & ","
      strRecord = strRecord & oTask.GetField(lngWBS) & ","
      strRecord = strRecord & oTask.GetField(lngOBS) & ","
      strRecord = strRecord & oTask.GetField(lngCA) & ","
      strRecord = strRecord & oTask.GetField(lngCAM) & ","
      strRecord = strRecord & oTask.GetField(lngWP) & ","
      If blnHasWPM Then
        strRecord = strRecord & oTask.GetField(lngWPM) & ","
      Else
        strRecord = strRecord & ","
      End If
      If lngLC > 0 Then
        strRecord = strRecord & oAssignment.Resource.GetField(lngLC) & ","
      Else
        strRecord = strRecord & oAssignment.ResourceName & ","
      End If
      If IsDate(oTask.ActualFinish) Then
        dtAF = FormatDateTime(oTask.ActualFinish, vbShortDate)
      Else
        dtAF = #1/1/1984#
      End If
      strRecord = strRecord & dtAF & ","
      strRecord = strRecord & CLng(Replace(oTask.GetField(lngEVP), "%", "")) & "," 'assumes 100
      Print #lngFile, strRecord
    Next oAssignment
    
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = Format(lngTask, "#,##0") & " / " & Format(lngTasks, "#,##0") & "...(" & Format(lngTask / lngTasks, "0%") & ")"
    DoEvents
  Next oTask
  Close #lngFile
  
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("tmp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT WP,MAX(AF),AVG(PercentComplete) AS EV "
  strSQL = strSQL & "FROM [wp.csv] "
  strSQL = strSQL & "GROUP BY WP "
  strSQL = strSQL & "HAVING AVG(PercentComplete)=100 "
  strSQL = strSQL & "ORDER BY MAX(AF) Desc "
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strSQL, strCon, 1, 1 '1=adOpenKeyset, 1=adLockReadOnly
  If oRecordset.RecordCount > 0 Then
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oExcel Is Nothing Then
      Set oExcel = CreateObject("Excel.Application")
    End If
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.Name = "COMPLETED WPs"
    oWorksheet.[A1:C1] = Array("WP", "AF", "EV")
    oWorksheet.[A1:C1].Font.Bold = True
    oWorksheet.[A2].CopyFromRecordset oRecordset
    oRecordset.Close
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.Columns(2).HorizontalAlignment = -4108 'xlCenter
    oWorksheet.Rows(1).HorizontalAlignment = -4131 'xlLeft
    oWorksheet.[A1].AutoFilter
    oWorksheet.Columns.AutoFit
    oExcel.ActiveWindow.SplitRow = 1
    oExcel.ActiveWindow.SplitColumn = 0
    oExcel.ActiveWindow.FreezePanes = True
    'get details
    If oWorkbook.Sheets.Count >= 2 Then
      Set oWorksheet = oWorkbook.Sheets(2)
    Else
      Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
    End If
    oWorksheet.Name = "DETAILS"
    strSQL = "SELECT * FROM [wp.csv] ORDER BY WP,PercentComplete"
    oRecordset.Open strSQL, strCon, 1, 1 '1=adOpenKeyset, 1=adLockReadOnly
    For lngItem = 0 To oRecordset.Fields.Count - 1
      oWorksheet.Cells(1, lngItem + 1) = oRecordset.Fields(lngItem).Name
    Next lngItem
    oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(-4161)).Font.Bold = True 'xlToRight
    oWorksheet.[A2].CopyFromRecordset oRecordset
    oWorksheet.Columns(8).Replace #1/1/1984#, "NA"
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.Columns(8).HorizontalAlignment = -4108 'xlCenter
    oWorksheet.Rows(1).HorizontalAlignment = -4131 'xlLeft
    oWorksheet.[A1].AutoFilter
    oWorksheet.Columns.AutoFit
    oExcel.ActiveWindow.SplitRow = 1
    oExcel.ActiveWindow.SplitColumn = 0
    oExcel.ActiveWindow.FreezePanes = True
    lngWPMCol = oWorksheet.Rows(1).Find("WPM", lookat:=xlWhole).Column
    If Not blnHasWPM Then
      oWorksheet.Columns(lngWPMCol).Delete
    End If
    lngEVPCol = oWorksheet.Rows(1).Find("PercentComplete", lookat:=xlWhole).Column
    oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)).AutoFilter Field:=lngEVPCol, Criteria1:="100"
    oRecordset.Close
    oWorkbook.Sheets("COMPLETED WPs").Activate
    oExcel.Visible = True
    oExcel.ActiveWindow.WindowState = -4143 'xlNormal
    Application.ActivateMicrosoftApp pjMicrosoftExcel
  Else
    MsgBox "No records found!", vbExclamation + vbOKOnly, "Completed Work"
  End If
    
exit_here:
  On Error Resume Next
  Set oCDP = Nothing
  Set oAssignment = Nothing
  Application.StatusBar = ""
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  If oRecordset.State = 1 Then oRecordset.Close
  Set oRecordset = Nothing
  Kill Environ("tmp") & "\Schema.ini"
  Kill Environ("tmp") & "\wp.csv"
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptCompletedWork", Err, Erl)
  Resume exit_here
End Sub

Sub cptFindUnstatusedTasks()
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  'strings
  Dim strMsg As String
  Dim strUnstatused As String
  'longs
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngUnstatused As Long
  'integers
  'doubles
  'booleans
  Dim blnAutoCalcCosts As Boolean
  Dim blnAutoTrack As Boolean
  Dim blnChangeSettings As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  
  On Error Resume Next
  Set oTasks = ActiveProject.Tasks
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oTasks Is Nothing Then
    MsgBox "This project has no tasks.", vbCritical + vbOKOnly, "No tasks"
    GoTo exit_here
  End If
  
  If oTasks.Count = 0 Then
    MsgBox "This project has no tasks.", vbCritical + vbOKOnly, "No tasks"
    GoTo exit_here
  End If
  
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Status Date is required.", vbCritical + vbOKOnly, "No Status Date"
    If Not Application.ChangeStatusDate Then
      GoTo exit_here
    Else
      dtStatus = ActiveProject.StatusDate
    End If
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  If Not IsDate(dtStatus) Then GoTo exit_here
  
  'catch and fix a non-working day
  If Not ActiveProject.Calendar.Period(dtStatus).Working Then
    'align status date to the most previous workday
    'to compare against a task's stop date
    Do While ActiveProject.Calendar.Period(dtStatus).Working = False
      dtStatus = DateAdd("h", -1, dtStatus)
    Loop
    'align the time to default finish time
    dtStatus = CDate(FormatDateTime(dtStatus, vbShortDate) & " " & ActiveProject.DefaultFinishTime)
  End If
  
  'Updating Task status updates resource status
  blnAutoTrack = ActiveProject.AutoTrack
  
  'Actual costs are always calculated by Project
  blnAutoCalcCosts = ActiveProject.AutoCalcCosts
  
  'prompt user to apply recommended settings
  If Not blnAutoTrack Or Not blnAutoCalcCosts Then
    strMsg = "Recommended settings:" & vbCrLf
    strMsg = strMsg & "> Updating Task status updates resource status = True" & vbCrLf 'AutoTrack
    strMsg = strMsg & "> Actual costs are always calculated by Project = True" & vbCrLf & vbCrLf 'AutoCalcCosts
    strMsg = strMsg & "Your settings:" & vbCrLf
    strMsg = strMsg & "> Updating Task status updates resource status = " & blnAutoTrack & vbCrLf
    strMsg = strMsg & "> Actual costs are always calculated by Project = " & blnAutoCalcCosts & vbCrLf & vbCrLf
    strMsg = strMsg & "(These options are found under File > Options > Schedule > Calculation options for this project:)" & vbCrLf & vbCrLf
    strMsg = strMsg & "Would you like to apply these recommended settings?"
    blnChangeSettings = MsgBox(strMsg, vbInformation + vbYesNo, "Apply Recommended Settings?") = vbYes
    If blnChangeSettings Then
      blnAutoTrack = ActiveProject.AutoTrack
      ActiveProject.AutoTrack = True
      blnAutoCalcCosts = ActiveProject.AutoCalcCosts
      ActiveProject.AutoCalcCosts = True
    End If
  End If
  
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True
  ActiveWindow.TopPane.Activate
  FilterClear
  GroupClear
  OptionsViewEx DisplaySummaryTasks:=True
  Sort "ID", , , , , , False, True
  OutlineShowAllTasks
  TimescaleEdit MajorUnits:=3, MinorUnits:=4, MajorCount:=1, MinorCount:=1, TierCount:=2
  EditGoTo Date:=dtStatus
  
  lngTasks = oTasks.Count
  
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task 'skip blank lines
    If oTask.Summary Then GoTo next_task 'skip summary tasks
    If oTask.ExternalTask Then GoTo next_task 'skip external tasks
    If Not oTask.Active Then GoTo next_task 'skip inactive tasks
    If oTask.Start < dtStatus And Not IsDate(oTask.ActualStart) Then 'unstarted
      strUnstatused = strUnstatused & oTask.UniqueID & vbTab
    ElseIf oTask.Finish <= dtStatus And Not IsDate(oTask.ActualFinish) Then 'unfinished
      strUnstatused = strUnstatused & oTask.UniqueID & vbTab
    ElseIf IsDate(oTask.ActualStart) And Not IsDate(oTask.ActualFinish) Then
      If oTask.Stop <> dtStatus Then  'unstatused
        strUnstatused = strUnstatused & oTask.UniqueID & vbTab
      End If
    End If
next_task:
    lngTask = lngTask + 1
    lngUnstatused = UBound(Split(strUnstatused, vbTab))
    Application.StatusBar = "Processing..." & Format(lngTask, "#,##0") & " of " & Format(lngTasks, "#,##0") & " (" & Format(lngTask / lngTasks, "0%") & ") " & IIf(lngUnstatused > 0, "| " & Format(lngUnstatused, "#,##0") & " found", "")
    DoEvents
  Next oTask
  'report results
  If lngUnstatused > 0 Then
    strUnstatused = Left(strUnstatused, Len(strUnstatused) - 1) 'hack off trailing tab
    ActiveWindow.TopPane.Activate
    FilterClear
    GroupClear
    Sort "ID", , , , , , False, True
    OptionsViewEx DisplaySummaryTasks:=True
    OutlineShowAllTasks
    OptionsViewEx DisplaySummaryTasks:=False
    SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strUnstatused
    SelectAll
    SetRowHeight "1"
    SelectBeginning
    cptSpeed False
    strMsg = ""
    If dtStatus <> ActiveProject.StatusDate Then
      strMsg = "NOTE: For purposes of this analysis, your Status Date was adjusted to the prior working day:" & vbCrLf
      strMsg = strMsg & ActiveProject.StatusDate & " > " & dtStatus & vbCrLf & vbCrLf
    End If
    strMsg = strMsg & "Given a Status Date of " & dtStatus & ":" & vbCrLf & vbCrLf
    strMsg = strMsg & "You have " & Format(lngUnstatused, "#,##0") & " unstatused task" & IIf(lngUnstatused = 1, ".", "s.") & vbCrLf & vbCrLf
    strMsg = strMsg & "Unstatused means:" & vbCrLf
    strMsg = strMsg & "> Forecast Start prior to Status Date" & vbCrLf
    strMsg = strMsg & "> Forecast Finish prior to Status Date" & vbCrLf
    strMsg = strMsg & "> In progress but not statused through 'Time Now' (see task field [Stop] for details)."
    MsgBox strMsg, vbExclamation + vbOKOnly, "You Have Unstatused Tasks!"
  Else
    MsgBox "No unstatused tasks.", vbInformation + vbOKOnly, "Well Done"
  End If
  
  'prompt to restore settings
  If blnChangeSettings Then
    strMsg = "Click YES to keep recommended settings:" & vbCrLf
    strMsg = strMsg & "> Updating Task status updates resource status = True" & vbCrLf 'AutoTrack
    strMsg = strMsg & "> Actual costs are always calculated by Project = True" & vbCrLf & vbCrLf 'AutoCalcCosts
    strMsg = strMsg & "Click NO to restore your settings:" & vbCrLf
    If Not blnAutoTrack Then
      strMsg = strMsg & "> Updating Task status updates resource status = " & blnAutoTrack & vbCrLf
    End If
    If Not blnAutoCalcCosts Then
      strMsg = strMsg & "> Actual costs are always calculated by Project = " & blnAutoCalcCosts & vbCrLf
    End If
    strMsg = strMsg & vbCrLf & "Would you like to keep these recommended settings?"
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Keep Recommended Settings?") = vbNo Then
      ActiveProject.AutoTrack = blnAutoTrack
      ActiveProject.AutoCalcCosts = blnAutoCalcCosts
    End If
  End If
  
  Application.StatusBar = "Complete."

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oTasks = Nothing
  Set oTask = Nothing
  cptSpeed False
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptFindUnstatusedTasks", Err, Erl)
  Resume exit_here
End Sub

Sub cptAddConditionalFormattingLegend(ByRef oWorkbook As Excel.Workbook)
  'objects
  Dim oListObject As Object
  Dim oWorksheet As Object
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  Dim vBorder As Variant
  Dim vArray As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
  oWorksheet.Activate
  oWorksheet.Name = "Conditional Formatting"
  vArray = Split(cptGetBreadcrumbs("cptStatusSheet_bas", "cptCopyData", "format-conditions"), vbCrLf)
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].Offset(UBound(vArray) - 1)) = oWorkbook.Application.WorksheetFunction.Transpose(vArray)
  oWorksheet.Range(oWorksheet.[A1048576].End(xlUp), oWorksheet.[A1048576].End(xlUp).End(xlUp)).Replace ":", ";", lookat:=xlPart
  oWorksheet.Range(oWorksheet.[A1048576].End(xlUp), oWorksheet.[A1048576].End(xlUp).End(xlUp)).Replace " -> ", ";", lookat:=xlPart
  oWorksheet.[C1:E1] = Split("COLUMN,CONDITION,FORMAT", ",")
  oWorksheet.Range(oWorksheet.[A1048576].End(xlUp), oWorksheet.[A1048576].End(xlUp).End(xlUp)).Cut oWorksheet.[C2]
  oWorksheet.Range(oWorksheet.[C2], oWorksheet.[C2].End(xlDown)).TextToColumns DataType:=xlDelimited, SemiColon:=True
  oWorksheet.[A1].Font.Bold = True
  oWorksheet.[A11].Font.Bold = True
  oWorksheet.[C1:E1].Font.Bold = True
  'make it a list
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[C1].End(xlToRight), oWorksheet.[C1].End(xlDown)), , xlYes)
  'borders and shading
  oListObject.TableStyle = ""
  oListObject.Range.Borders(xlDiagonalDown).LineStyle = xlNone
  oListObject.Range.Borders(xlDiagonalUp).LineStyle = xlNone
  For Each vBorder In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    With oListObject.HeaderRowRange.Borders(vBorder)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
    End With
    With oListObject.DataBodyRange.Borders(vBorder)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
    End With
  Next vBorder
  'inside borders
  For Each vBorder In Array(xlInsideVertical, xlInsideHorizontal)
    With oListObject.HeaderRowRange.Borders(vBorder)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.249946592608417
      .Weight = xlThin
    End With
    With oListObject.DataBodyRange.Borders(vBorder)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.249946592608417
      .Weight = xlThin
    End With
  Next vBorder
  oListObject.HeaderRowRange.Font.Bold = True
  With oListObject.HeaderRowRange.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = -0.149998474074526
    .PatternTintAndShade = 0
  End With
  
  'conditional formatting
  With oListObject.ListColumns("FORMAT").DataBodyRange
    .FormatConditions.Delete
    
    .FormatConditions.Add xlTextString, String:="=""BAD""", TextOperator:=xlContains
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Font
      .Color = -16383844
      .TintAndShade = 0
    End With
    With .FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 13551615
      .TintAndShade = 0
    End With
    .FormatConditions(1).StopIfTrue = False
    
    .FormatConditions.Add xlTextString, String:="=""NEUTRAL""", TextOperator:=xlContains
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Font
      .Color = -16754788
      .TintAndShade = 0
    End With
    With .FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 10284031
      .TintAndShade = 0
    End With
    .FormatConditions(1).StopIfTrue = False
  
    .FormatConditions.Add xlCellValue, xlEqual, Formula1:="=""GOOD"""
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Font
      .Color = -16752384
      .TintAndShade = 0
    End With
    With .FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 13561798
      .TintAndShade = 0
    End With
    .FormatConditions(1).StopIfTrue = False
  
    .FormatConditions.Add xlTextString, String:="=""COMPLETE""", TextOperator:=xlContains
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Font
      .Color = 8355711
      .TintAndShade = 0
    End With
    With .FormatConditions(1).Interior
      .PatternColorIndex = -4105
      .Color = 15921906
      .TintAndShade = 0
    End With
    .FormatConditions(1).StopIfTrue = False

    .FormatConditions.Add xlTextString, String:="=""INPUT""", TextOperator:=xlContains
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Font
      .Color = 7749439
      .TintAndShade = 0
    End With
    With .FormatConditions(1).Interior
      .PatternColorIndex = -4105
      .Color = 10079487
      .TintAndShade = 0
    End With
    .FormatConditions(1).StopIfTrue = False
  End With
  oWorksheet.Application.ActiveWindow.Zoom = 85
  oWorksheet.[A1].Select
  oWorksheet.Columns.AutoFit
  oWorksheet.Columns(2).ColumnWidth = 1
  oWorksheet.Application.ActiveWindow.SplitRow = 1
  oWorksheet.Application.ActiveWindow.SplitColumn = 0
  oWorksheet.Application.WindowState = xlNormal
  oWorksheet.Application.ActiveWindow.FreezePanes = True
  oWorksheet.Application.WindowState = xlMinimized
  oWorksheet.Application.ActiveWindow.DisplayGridlines = False
        
exit_here:
  On Error Resume Next
  Set oListObject = Nothing
  Set oWorksheet = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptAddConditionalFormattingLegend", Err, Erl)
  Resume exit_here
End Sub

Sub cptFindCompleteThrough()
  'objects
  Dim oTSV As TimeScaleValue
  Dim oTSVS As TimeScaleValues
  Dim oTask As MSProject.Task
  'strings
  Dim strReport As String
  Dim strFileName As String
  'longs
  Dim lngFile As Long
  Dim lngComplete As Long
  Dim lngWorkdays As Long
  Dim lngSplitPart As Long
  'integers
  'doubles
  Dim dblPercentComplete As Double
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  lngFile = FreeFile
  strFileName = Environ("tmp") & "\completeThrough.txt"
  Open strFileName For Output As #lngFile
  
  On Error Resume Next
  Set oTask = ActiveSelection.Tasks(1)
  If oTask Is Nothing Then GoTo exit_here
  If Not IsDate(oTask.ActualStart) Then 'ignore unstarted
    MsgBox "Task has not started yet (no Actual Start).", vbExclamation + vbOKCancel, "Invalid"
    GoTo exit_here
  End If
  If IsDate(oTask.ActualFinish) Then  'ignore completed
    MsgBox "Task is already complete (has Actual Finish).", vbExclamation + vbOKOnly, "Invalid"
    GoTo exit_here
  End If
  
  'todo: do splits matter if no resources?
  Print #lngFile, "Task UID: " & oTask.UniqueID
  Print #lngFile, "Task Name: " & oTask.Name
  Print #lngFile, "Task Type: " & Choose(oTask.Type + 1, "Fixed Units", "Fixed Duration", "Fixed Work")
  Print #lngFile, "Actual Start: " & FormatDateTime(oTask.ActualStart, vbShortDate)
  Print #lngFile, "Stop: " & FormatDateTime(oTask.Stop, vbShortDate) & " (vs. Status Date: " & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & ")"
  Print #lngFile, "Resume: " & FormatDateTime(oTask.Resume, vbShortDate)
  Print #lngFile, "Finish: " & FormatDateTime(oTask.Finish, vbShortDate)
  Print #lngFile, String(20, "-")
  Print #lngFile, "DETERMINE DURATION PERCENT COMPLETE:"
  Print #lngFile, "actual duration: " & oTask.ActualDuration / 480 & " days"
  Print #lngFile, "total duration: " & oTask.Duration / 480 & " days"
  Print #lngFile, "duration % complete: " & Format(oTask.ActualDuration / oTask.Duration, "0%")
  Print #lngFile, String(20, "-")
  Print #lngFile, "DETERMINE WORK % COMPLETE:"
  Print #lngFile, "actual work: " & oTask.ActualWork / 60 & " hrs"
  Print #lngFile, "total work: " & oTask.Work / 60 & " hrs"
  Print #lngFile, "work % complete " & Format(oTask.ActualWork / oTask.Work, "0%")
  Print #lngFile, String(20, "-")
  Print #lngFile, "DETERMINE TOTAL RESOURCED WORKDAYS:"
  dblPercentComplete = oTask.ActualDuration / oTask.Duration
  Print #lngFile, "SplitParts: " & oTask.SplitParts.Count
  For lngSplitPart = 1 To oTask.SplitParts.Count
    Print #lngFile, "SplitPart" & lngSplitPart & ": " & oTask.SplitParts(lngSplitPart).Start & " -  " & oTask.SplitParts(lngSplitPart).Finish & " (" & Application.DateDifference(oTask.SplitParts(lngSplitPart).Start, oTask.SplitParts(lngSplitPart).Finish) / 480 & "d)"
    lngWorkdays = lngWorkdays + Application.DateDifference(oTask.SplitParts(lngSplitPart).Start, oTask.SplitParts(lngSplitPart).Finish)
  Next lngSplitPart
  Print #lngFile, "= " & lngWorkdays / 480 & " total resourced workdays"
  Print #lngFile, String(20, "-")
  Print #lngFile, "DETERMINE COMPLETED RESOURCED WORKDAYS:"
  Print #lngFile, (lngWorkdays / 480) & " total resourced workdays * " & Format(dblPercentComplete, "0%") & " duration % complete = " & (oTask.PercentComplete / 100) * (lngWorkdays / 480) & " completed resourced workdays"
  Print #lngFile, String(20, "-")
  Print #lngFile, "DETERMINE CompleteThrough:"
  
  Set oTSVS = oTask.TimeScaleData(oTask.Start, oTask.Finish, pjTaskTimescaledWork, pjTimescaleDays)
  
  lngComplete = 0
  For Each oTSV In oTSVS
    If Val(oTSV.Value) > 0 Then
      'note: workdays may not begin at 8:00 AM and end at 5:00 PM
      If lngComplete + Application.DateDifference(oTSV.StartDate, oTSV.EndDate) < (lngWorkdays * dblPercentComplete) Then
        lngComplete = lngComplete + Application.DateDifference(oTSV.StartDate, oTSV.EndDate)
        Print #lngFile, oTSV.StartDate & " 8:00 AM - " & oTSV.StartDate & " 5:00 PM (" & lngComplete / 480 & " workday" & IIf((lngComplete / 480) = 1, "", "s") & " complete = " & Format(lngComplete / lngWorkdays, "0%") & ")"
      Else
        Print #lngFile, oTSV.StartDate & " 8:00 AM - " & Application.DateAdd(oTSV.StartDate, (lngWorkdays * dblPercentComplete) - lngComplete) & " (" & Round(dblPercentComplete * (lngWorkdays / 480), 1) & " workdays complete = " & Format(dblPercentComplete, "0%") & ") <-- CompleteThrough"
        Exit For
      End If
    Else
      Print #lngFile, oTSV.StartDate & " excluded (no work)"
    End If
  Next oTSV

  Print #lngFile, strReport
  Print #lngFile, String(20, "-")
'  Print #lngFile, "TIP: if the graphical progress bar is causing confusion, set it to run from [Actual Start] through [Stop] instead of from [Actual Start] through [CompleteThrough]."
'  Print #lngFile, String(20, "-")
  Close #lngFile
  ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
  
exit_here:
  On Error Resume Next
  Reset 'closes all active files opened by the Open statement and writes the contents of all file buffers to disk.
  Set oTask = Nothing
  Set oTSV = Nothing
  Set oTSVS = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptFindCompleteThrough", Err, Erl)
  Resume exit_here

End Sub

Function cptGetEarliestStart() As Date
  Dim dtStart As Date
  Dim oSubproject As SubProject
  dtStart = #12/31/2149#
  For Each oSubproject In ActiveProject.Subprojects
    If oSubproject.InsertedProjectSummary.Start < dtStart Then
      dtStart = oSubproject.InsertedProjectSummary.Start
    End If
  Next oSubproject
  cptGetEarliestStart = dtStart
End Function

Sub cptFindAssignmentsWithoutWork()
  'objects
  Dim oDict As Scripting.Dictionary
  Dim oTasks As MSProject.Tasks
  Dim oAssignment As MSProject.Assignment
  Dim oTask As MSProject.Task
  'strings
  Dim strResultUID As String
  Dim strResult As String
  Dim strFileName As String
  Dim strMissingForecastWork As String
  'longs
  Dim lngCount As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnDelete As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnDelete = MsgBox("Delete whatever I find?" & vbCrLf & vbCrLf & "(Baseline data will not be touched.)" & vbCrLf & vbCrLf & "Note: click No for a dry-run and review what clicking Yes might do.", vbQuestion + vbYesNo, "Find Assignments Without ETC") = vbYes
  If blnDelete Then
    Application.OpenUndoTransaction "Delete Assignments with Zero Remaining Work"
    Application.Calculation = pjManual
    Application.ScreenUpdating = False
  End If
  If ActiveProject.Subprojects.Count > 0 Then
    ActiveWindow.TopPane.Activate
    FilterClear
    GroupClear
    Sort "ID", , , , , , False, True
    SelectAll
    OutlineShowAllTasks
    SelectAll
    Set oTasks = ActiveSelection.Tasks
  Else
    Set oTasks = ActiveProject.Tasks
  End If
  Set oDict = CreateObject("Scripting.Dictionary")
  lngTasks = oTasks.Count
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    For Each oAssignment In oTask.Assignments
      If oAssignment.Work + oAssignment.Cost = 0 Then
        If oAssignment.BaselineWork + oAssignment.BaselineCost = 0 Then
          lngCount = lngCount + 1 'counting assignments, not tasks
          'add to list for autofilter / filter by clipboard
          If Not oDict.Exists(oTask.UniqueID) Then
            oDict.Add oTask.UniqueID, oTask.UniqueID
          End If
          If blnDelete Then
            strResult = strResult & oTask.UniqueID & "," & oAssignment.ResourceUniqueID & "," & oAssignment.ResourceName & ",0,0, Assignment had zero Baseline Work/Cost and zero Remaining Work (ETC) and has been deleted." & vbCrLf
            oAssignment.Delete
          Else
            strResult = strResult & oTask.UniqueID & "," & oAssignment.ResourceUniqueID & "," & oAssignment.ResourceName & ",0,0,Assignment has zero Baseline Work/Cost and zero Remaining Work (ETC) and can be deleted." & vbCrLf
          End If
        End If
      End If
    Next oAssignment
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Analyzing...(" & Format(lngTask / lngTasks, "0%") & ") | " & Format(lngCount, "#,##0") & " found"
  Next oTask
  If lngCount > 0 Then
    strFileName = Environ("tmp") & "\cpt-assignments-without-work_" & Format(Now, "yyyy-mm-dd_hh-nn-ss") & ".txt"
    lngFile = FreeFile
    Open strFileName For Output As #lngFile
    Print #lngFile, "FILE: " & ActiveProject.FullName
    Print #lngFile, "DATE: " & FormatDateTime(Now, vbGeneralDate) & vbCrLf
    Print #lngFile, "'ASSIGNMENTS WITHOUT WORK' MEANS:"
    Print #lngFile, "Assignment Baseline Work/Cost = 0 AND Assignment Remaining Work/Cost (ETC) = 0"
    Print #lngFile, "WHERE:"
    Print #lngFile, "-> [Assignment Work] = (Assignment Actual Work + Assignment Remaining Work)"
    Print #lngFile, "-> [Assignment Cost] = (Assignment Actual Cost + Assignment Remaining Cost)"
    Print #lngFile, "-> [Assignment Work] + [Assignment Cost] = 0"
    Print #lngFile, "-> Assignment Baseline Work + Assignment Baseline Cost = 0" & vbCrLf
    Print #lngFile, String(80, "-")
    Print #lngFile, "TASK UID,RESOURCE UID,RESOURCE NAME,BASELINE WORK/COST,REMAINING WORK/COST,COMMENT"
    Print #lngFile, Left(strResult, Len(strResult) - 1)
    Print #lngFile, Format(lngCount, "#,##0") & " ASSIGNMENT" & IIf(lngCount = 1, ":", "S") & " FOUND."
    Print #lngFile, String(80, "-")
    Print #lngFile, "Paste this into ClearPlan > Text > FilterByClipboard:"
    Print #lngFile, Join(oDict.Keys, ",") & ","
    Print #lngFile, String(80, "-")
    If blnDelete = False Then
      Print #lngFile, "NOTE: To delete these assignments, run this macro again. At the prompt ('Delete what I find?'), click Yes. Undo is enabled."
    End If
    Print #lngFile, "NOTE: Resources can have the same name in MS Project. Confirm Resource Unique ID before deleting."
    Close #lngFile
    ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
    SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", Join(oDict.Keys, vbTab)
  Else
    MsgBox "There are ZERO assignments without remaining work!", vbInformation + vbOKOnly, "Well Done"
  End If

exit_here:
  On Error Resume Next
  Set oDict = Nothing
  Application.CloseUndoTransaction
  Set oTasks = Nothing
  Application.Calculation = pjAutomatic
  Application.ScreenUpdating = True
  Set oAssignment = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  cptHandleErr "cptStatusSheet_bas", "cptAssignmentsWithoutWork", Err, Erl
  Resume exit_here
End Sub

Sub cptRespreadAssignmentWork()
  'purpose: to spread assignment finish dates to task finish dates
  'use: run macro
  'objects
  Dim oRemainingWork As Scripting.Dictionary
  Dim oAssignment As MSProject.Assignment
  Dim oTask As MSProject.Task
  Dim oTasks As MSProject.Tasks
  'strings
  'longs
  Dim lngMismatched As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngItem As Long
  Dim lngRemainingDuration As Long
  Dim lngRemainingWork As Long
  Dim lngTaskType As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnMismatch As Boolean
  Dim blnEffortDriven As Boolean
  'variants
  'dates

  Application.OpenUndoTransaction "cptRespreadAssignments"

  blnErrorTrapping = cptErrorTrapping
  
  On Error Resume Next
  'OPTIONAL: Change 'ActiveSelection' to 'ActiveProject' in the next line to execute on ALL tasks
  Set oTasks = ActiveSelection.Tasks
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then
    MsgBox "No tasks selected.", vbCritical + vbOKOnly, "Error"
    GoTo exit_here
  End If
  If oTasks.Count = 0 Then
    MsgBox "No tasks selected.", vbCritical + vbOKOnly, "Error"
    GoTo exit_here
  End If
  
  'provide user feedback in StatusBar
  lngTasks = oTasks.Count
  lngTask = 0

  Set oRemainingWork = CreateObject("Scripting.Dictionary")
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task 'skip blank task lines
    If oTask.Summary Then GoTo next_task 'skip summary tasks
    If Not oTask.Active Then GoTo next_task 'skip inactive tasks
    If oTask.ExternalTask Then GoTo next_task 'skip external tasks
    If oTask.Assignments.Count = 0 Then GoTo next_task 'skip SVTs, Milestones, Schedule Margin, etc.
    blnMismatch = False
    For Each oAssignment In oTask.Assignments
      'todo: test on work, material, and cost
      'todo: if cost, then account for AccrueAt
      If oAssignment.Finish <> oTask.Finish Then
        blnMismatch = True
        Exit For
      End If
    Next oAssignment
    If Not blnMismatch Then GoTo next_task
    lngMismatched = lngMismatched + 1
    'capture task settings
    lngTaskType = oTask.Type
    blnEffortDriven = oTask.EffortDriven
    'capture remaining duration
    lngRemainingDuration = oTask.RemainingDuration
    lngRemainingWork = oTask.RemainingWork
    'clear the dictionary before capturing task assignments
    If oRemainingWork.Count > 0 Then oRemainingWork.RemoveAll
    'capture remaining work
    For Each oAssignment In oTask.Assignments
      oRemainingWork.Add oAssignment.ResourceName, oAssignment.RemainingWork
    Next oAssignment
    'set to fixed duration
    oTask.Type = pjFixedDuration
    oTask.EffortDriven = False
    'set remaining duration to 0
    oTask.RemainingDuration = 0
    'restore remaining duration
    oTask.RemainingDuration = lngRemainingDuration
    'restore remaining work
    For Each oAssignment In oTask.Assignments
      oAssignment.RemainingWork = oRemainingWork(oAssignment.ResourceName)
    Next oAssignment
    'restore task settings
    If lngTaskType <> pjFixedDuration Then oTask.Type = lngTaskType
    If oTask.Type <> pjFixedWork Then oTask.EffortDriven = blnEffortDriven
next_task:
    'provide user feedback
    Application.StatusBar = "Fixing tasks...(" & Format(lngTask / lngTasks, "0%") & ")"
    DoEvents
  Next oTask

  If lngMismatched > 0 Then
    MsgBox Format(lngMismatched, "#,##0") & " mismatched task(s) respread.", vbInformation + vbOKOnly, "Complete"
  Else
    MsgBox "No mismatched task/assignment finish dates found.", vbInformation + vbOKOnly, "Complete"
  End If

  'provide user feedback
  Application.StatusBar = "Complete."

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Application.StatusBar = ""
  Set oRemainingWork = Nothing
  Set oAssignment = Nothing
  Set oTask = Nothing
  Set oTasks = Nothing

  Exit Sub
err_here:
  cptHandleErr "cptStatusSheet_bas", "cptRespreadAssignmentWork", Err, Erl
  Resume exit_here
End Sub

Sub cptMarkOnTrackRetainETC(Optional blnOpenUndoTransaction As Boolean = True)
  'objects
  Dim oDict As Scripting.Dictionary
  Dim oAssignment As MSProject.Assignment
  Dim oTask As MSProject.Task
  Dim oTasks As MSProject.Tasks
  'strings
  Dim strMsg As String
  Dim strChangeApplicationSettings As String
  Dim strChangeProjectSettings As String
  'longs
  Dim lngTaskType As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  'integers
  'doubles
  Dim dblRemainingDuration  As Double
  Dim dblRemainingWork As Double
  'booleans
  Dim blnSpreadCostsToStatusDate As Boolean
  Dim blnShowTaskWarnings As Boolean
  Dim blnShowTaskSuggestions As Boolean
  Dim blnSpreadPercentCompleteToStatusDate As Boolean
  Dim blnAndMoveRemaining As Boolean
  Dim blnMoveCompleted As Boolean
  Dim blnAndMoveCompleted As Boolean
  Dim blnMoveRemaining As Boolean
  Dim blnAutoCalcCosts As Boolean
  Dim blnAutoTrack As Boolean
  Dim blnDisplayWizardUsage As Boolean
  Dim blnDisplayWizardScheduling As Boolean
  Dim blnDisplayWizardErrors As Boolean
  Dim blnDisplayScheduleMessages As Boolean
  Dim blnDisplayAlerts As Boolean
  Dim blnEffortDriven As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  Dim dtFinish As Date
  Dim dtStatus As Date
  
  'ensure status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Please enter a Status Date.", vbExclamation + vbOKOnly, "Status Date Required"
    If Not Application.ChangeStatusDate Then
      GoTo exit_here
    Else
      dtStatus = ActiveProject.StatusDate
    End If
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  If Not IsDate(dtStatus) Then GoTo exit_here
  
  cptSpeed True
  
  Application.StatusBar = "Capturing settings..."
  DoEvents
  'Application Settings:
  blnDisplayAlerts = Application.DisplayAlerts
  If blnDisplayAlerts Then
    strChangeApplicationSettings = strChangeApplicationSettings & "> Display Alerts: True -> False" & vbCrLf
  End If
  blnDisplayScheduleMessages = Application.DisplayScheduleMessages
  If blnDisplayScheduleMessages Then
    strChangeApplicationSettings = strChangeApplicationSettings & "> Display Schedule Messages: True -> False" & vbCrLf
  End If
  blnDisplayWizardErrors = Application.DisplayWizardErrors
  If blnDisplayWizardErrors Then
    strChangeApplicationSettings = strChangeApplicationSettings & "> Display Wizard Errors: True -> False" & vbCrLf
  End If
  blnDisplayWizardScheduling = Application.DisplayWizardScheduling
  If blnDisplayWizardScheduling Then
    strChangeApplicationSettings = strChangeApplicationSettings & "> Display Wizard Scheduling: True -> False" & vbCrLf
  End If
  blnDisplayWizardUsage = Application.DisplayWizardUsage
  If blnDisplayWizardUsage Then
    strChangeApplicationSettings = strChangeApplicationSettings & "> Display Wizard Usage: True -> False" & vbCrLf
  End If
  If Len(strChangeApplicationSettings) > 0 Then
    strChangeApplicationSettings = "Application Settings:" & vbCrLf & strChangeApplicationSettings
  End If
  
  'Current Project Settings
  'Schedule Settings:
  blnAutoTrack = ActiveProject.AutoTrack
  If blnAutoTrack = False Then
    strChangeProjectSettings = strChangeProjectSettings & "> Schedule.AutoTrack: False -> True" & vbCrLf
  End If
  blnAutoCalcCosts = ActiveProject.AutoCalcCosts
  If blnAutoCalcCosts = False Then
    strChangeProjectSettings = strChangeProjectSettings & "> Schedule.AutoCalcCosts: False -> True" & vbCrLf
  End If
  'Advanced Settings:
  blnMoveRemaining = ActiveProject.MoveRemaining
  If blnMoveRemaining = False Then
    strChangeProjectSettings = strChangeProjectSettings & "> Advanced.MoveRemaining: False -> True" & vbCrLf
  End If
  blnAndMoveCompleted = ActiveProject.AndMoveCompleted
  If blnAndMoveCompleted = False Then
    strChangeProjectSettings = strChangeProjectSettings & "> Advanced.AndMoveCompleted: False -> True" & vbCrLf
  End If
  blnMoveCompleted = ActiveProject.MoveCompleted
  If blnMoveCompleted = False Then
    strChangeProjectSettings = strChangeProjectSettings & "> Advanced.MoveCompleted: False -> True" & vbCrLf
  End If
  blnAndMoveRemaining = ActiveProject.AndMoveRemaining
  If blnAndMoveRemaining = False Then
    strChangeProjectSettings = strChangeProjectSettings & "> Advanced.AndMoveRemaining: False -> True" & vbCrLf
  End If
  blnSpreadPercentCompleteToStatusDate = ActiveProject.SpreadPercentCompleteToStatusDate
  If blnSpreadPercentCompleteToStatusDate = False Then
    strChangeProjectSettings = strChangeProjectSettings & "> Advanced.SpreadPercentCompleteToStatusDate: False -> True" & vbCrLf
  End If
  blnShowTaskSuggestions = ActiveProject.ShowTaskSuggestions
  If blnShowTaskSuggestions Then
    strChangeProjectSettings = strChangeProjectSettings & "> Advanced.ShowTaskSuggestions: True -> False" & vbCrLf
  End If
  blnShowTaskWarnings = ActiveProject.ShowTaskWarnings
  If blnShowTaskWarnings Then
    strChangeProjectSettings = strChangeProjectSettings & "> Advanced.ShowTaskWarnings: True -> False" & vbCrLf
  End If
  blnSpreadCostsToStatusDate = ActiveProject.SpreadCostsToStatusDate
  If blnSpreadCostsToStatusDate = False Then
    strChangeProjectSettings = strChangeProjectSettings & "> Advanced.SpreadCostsToStatusDate: True -> False" & vbCrLf
  End If
  If Len(strChangeProjectSettings) > 0 Then
    strChangeProjectSettings = "Settings for this file (" & ActiveProject.Name & "):" & vbCrLf & strChangeProjectSettings
  End If
  
  If Len(strChangeApplicationSettings) > 0 Or Len(strChangeProjectSettings) > 0 Then
    'do something?
  End If
  
  'maybe I don't care, I'll just do it, eh?
  Application.StatusBar = "Applying temporary settings..."
  DoEvents
  Application.DisplayAlerts = False
  Application.DisplayScheduleMessages = False
  Application.DisplayWizardErrors = False
  Application.DisplayWizardScheduling = False
  Application.DisplayWizardUsage = False
  ActiveProject.AutoTrack = True
  ActiveProject.AutoCalcCosts = True
  ActiveProject.MoveRemaining = True
  ActiveProject.AndMoveCompleted = True
  ActiveProject.MoveCompleted = True
  ActiveProject.AndMoveRemaining = True
  ActiveProject.SpreadPercentCompleteToStatusDate = True
  ActiveProject.ShowTaskSuggestions = False
  ActiveProject.ShowTaskWarnings = False
  ActiveProject.SpreadCostsToStatusDate = True

  'catch and fix a non-working day
  If Not ActiveProject.Calendar.Period(dtStatus).Working Then
    'align status date to the most previous workday
    'to compare against a task's stop date
    Do While ActiveProject.Calendar.Period(dtStatus).Working = False
      dtStatus = DateAdd("h", -1, dtStatus)
    Loop
    'align the time to default finish time
    dtStatus = CDate(FormatDateTime(dtStatus, vbShortDate) & " " & ActiveProject.DefaultFinishTime)
  End If
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then
    MsgBox "Please select a task (or tasks).", vbExclamation + vbOKOnly, "No Task(s) Selected"
    GoTo exit_here
  End If
  If oTasks.Count = 0 Then GoTo exit_here
  
  'prep to capture assignment remaining work
  Set oDict = CreateObject("Scripting.Dictionary")
  
  If blnOpenUndoTransaction Then
    Application.OpenUndoTransaction "cpt Mark On Track - Retain ETC"
  End If
  
  lngTasks = oTasks.Count
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If IsDate(oTask.ActualFinish) Then GoTo next_task
    ActiveWindow.TopPane.Activate
    EditGoTo , dtStatus
    
    'mark complete if should have finished
    If oTask.Finish <= dtStatus Then
      If MsgBox("Mark Complete?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        oTask.ActualFinish = oTask.Finish
      Else
        GoTo next_task
      End If
    End If
    
    'mark started if should have started
    Application.ScreenUpdating = True
    If oTask.Start < dtStatus And Not IsDate(oTask.ActualStart) Then
      UpdateProject All:=False, UpdateDate:=dtStatus, Action:=0 'AS or AF only
    End If
    
    'capture forecast finish
    dtFinish = oTask.Finish
    dblRemainingWork = oTask.RemainingWork
    dblRemainingDuration = oTask.RemainingDuration
    
    'capture task settings
    lngTaskType = oTask.Type
    blnEffortDriven = oTask.EffortDriven
    oDict.RemoveAll 'start fresh with each task
    For Each oAssignment In oTask.Assignments
      Application.StatusBar = "Capturing " & oAssignment.ResourceName & "..."
      DoEvents
      If oAssignment.WorkContour <> pjFlat Then
        strMsg = "Task UID " & oTask.UniqueID & " - " & oTask.Name & vbCrLf
        strMsg = strMsg & "Resource Assignment '" & oAssignment.ResourceName & "' has a non-standard Work Contour (" & cptGetConstantName("WorkContour", oAssignment.WorkContour) & ")." & vbCrLf & vbCrLf
        strMsg = strMsg & "> Click YES to override manual edits" & vbCrLf
        strMsg = strMsg & "> Click NO to skip this Assignment"
        If MsgBox(strMsg, vbQuestion + vbYesNo, "Override Manual Work Contour?") = vbYes Then
          Calculation = pjAutomatic
          oAssignment.WorkContour = pjFlat
          CalculateProject
          Calculation = pjManual
        Else
          GoTo next_assignment
        End If
      End If
      If oAssignment.ResourceType <> pjResourceTypeCost Then
        oDict.Add oAssignment.UniqueID, oAssignment.RemainingWork
      Else
        If oAssignment.Resource.AccrueAt = pjStart Then
          Calculation = pjAutomatic
          UpdateProject All:=False, UpdateDate:=dtStatus, Action:=1
          CalculateProject
          Calculation = pjManual
        ElseIf oAssignment.Resource.AccrueAt = pjEnd Then
          Calculation = pjAutomatic
          UpdateProject All:=False, UpdateDate:=dtStatus, Action:=1
          CalculateProject
          Calculation = pjManual
        ElseIf oAssignment.Resource.AccrueAt = pjProrated Then
          oDict.Add oAssignment.UniqueID, oAssignment.RemainingCost
        End If
      End If
      'todo: deal with misaligned dates
next_assignment:
    Next oAssignment
    'change task type to remaining duration
    'Application.ScreenUpdating = True
    If oTask.Type <> pjFixedDuration Then oTask.Type = pjFixedDuration
    If oTask.EffortDriven Then oTask.EffortDriven = False
    'update as scheduled
    If oTask.Stop <> dtStatus Then 'rebuild the task
      Application.StatusBar = "Rebuilding UID " & oTask.UniqueID & "..."
      DoEvents
      Calculation = pjAutomatic
      oTask.ActualFinish = ActiveProject.StatusDate 'dtStatus
      oTask.RemainingDuration = Application.DateDifference(dtStatus, dtFinish, ActiveProject.Calendar)
      Calculation = pjManual
    End If
    'retain ETC
    For Each oAssignment In oTask.Assignments
      If oDict.Exists(oAssignment.UniqueID) Then
        Application.StatusBar = "Restoring " & oAssignment.ResourceName & "..."
        DoEvents
        If oAssignment.ResourceType = pjResourceTypeWork Then
          Do While oAssignment.RemainingWork <> oDict(oAssignment.UniqueID)
            oAssignment.RemainingWork = 0
            oAssignment.RemainingWork = oDict(oAssignment.UniqueID)
          Loop
        ElseIf oAssignment.ResourceType = pjResourceTypeMaterial Then
          oAssignment.RemainingWork = oDict(oAssignment.UniqueID) * 60
        ElseIf oAssignment.ResourceType = pjResourceTypeCost Then
          If oAssignment.Resource.AccrueAt = pjStart Then
            'take no action
          ElseIf oAssignment.Resource.AccrueAt = pjEnd Then
            'take no action
          ElseIf oAssignment.Resource.AccrueAt = pjProrated Then
            oAssignment.Cost = oAssignment.ActualCost + oDict(oAssignment.UniqueID)
          End If
        End If
      Else
        Application.StatusBar = "Skipping " & oAssignment.ResourceName & "..."
        DoEvents
      End If
    Next oAssignment
    'restore task settings
    oTask.Type = lngTaskType
    If lngTaskType <> pjFixedWork Then oTask.EffortDriven = blnEffortDriven
    'todo: ok, what could go wrong?
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Marking on track (retaining ETC)...(" & Format(lngTask / lngTasks, "0%") & ")"
    DoEvents
  Next oTask
  
  If blnOpenUndoTransaction Then
    Application.CloseUndoTransaction
  End If
  
  Application.StatusBar = "Restoring settings..."
  DoEvents
  
  'restore application/project settings
  Application.DisplayAlerts = blnDisplayAlerts
  Application.DisplayScheduleMessages = blnDisplayScheduleMessages
  If blnDisplayScheduleMessages Then
    Application.DisplayWizardErrors = blnDisplayWizardErrors
    Application.DisplayWizardScheduling = blnDisplayWizardScheduling
    Application.DisplayWizardUsage = blnDisplayWizardUsage
  End If
  ActiveProject.AutoTrack = blnAutoTrack
  ActiveProject.AutoCalcCosts = blnAutoCalcCosts
  ActiveProject.MoveRemaining = blnMoveRemaining
  ActiveProject.AndMoveCompleted = blnAndMoveCompleted
  ActiveProject.MoveCompleted = blnMoveCompleted
  ActiveProject.AndMoveRemaining = blnAndMoveRemaining
  ActiveProject.SpreadPercentCompleteToStatusDate = blnSpreadPercentCompleteToStatusDate
  ActiveProject.ShowTaskSuggestions = blnShowTaskSuggestions
  ActiveProject.ShowTaskWarnings = blnShowTaskWarnings
  ActiveProject.SpreadCostsToStatusDate = blnSpreadCostsToStatusDate

  Application.StatusBar = "...complete."
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  cptSpeed False
  If blnOpenUndoTransaction Then Application.CloseUndoTransaction
  Set oAssignment = Nothing
  Set oDict = Nothing
  Set oTask = Nothing
  Set oTasks = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptMarkOnTrackRetainETC", Err, Erl)
  Resume exit_here
End Sub

Function cptValidateAssignmentRolldown(strCFN As String) As Boolean
  'objects
  Dim oTask As MSProject.Task
  Dim oAssignment As MSProject.Assignment
  Dim oToFix As Scripting.Dictionary
  'strings
  Dim strNotFixed As String
  Dim strValue As String  'task-level field value
  Dim strKey As String    'field name (native)
  Dim strMsg As String
  'longs
  Dim lngNotFixed As Long
  Dim lngFieldID As Long  'field constant
  Dim lngMismatched As Long
  Dim lngFactor As Long   'for deriving Unique IDs
  Dim lngAssignmentUID As Long
  Dim lngItem As Long
  Dim lngTasks As Long
  Dim lngAssignments As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnValid As Boolean
  Dim blnLCF As Boolean
  Dim blnMaster As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  
  blnMaster = ActiveProject.Subprojects.Count > 0
  'todo: assumes LCF sync (oof)
  
  lngFieldID = FieldNameToFieldConstant(strCFN)
  blnLCF = lngFieldID < 188776000
  strKey = FieldConstantToFieldName(lngFieldID)
  
  cptSpeed True
  
  'capture mismatches
  Set oToFix = CreateObject("Scripting.Dictionary")
  SelectAll
  For Each oTask In ActiveSelection.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    strValue = oTask.GetField(lngFieldID)
    For Each oAssignment In oTask.Assignments
      lngAssignmentUID = oAssignment.UniqueID - (2 ^ 20)
      If blnMaster Then
        lngFactor = (oTask.UniqueID - (oTask.UniqueID Mod 4194304)) / 4194304
        lngAssignmentUID = lngAssignmentUID + (lngFactor * 4194304)
      End If
      If blnLCF Then
        Select Case strKey
          Case "Cost1"
            If oAssignment.Cost1 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Cost2"
            If oAssignment.Cost2 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Cost3"
            If oAssignment.Cost3 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Cost4"
            If oAssignment.Cost4 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Cost5"
            If oAssignment.Cost5 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Cost6"
            If oAssignment.Cost6 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Cost7"
            If oAssignment.Cost7 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Cost8"
            If oAssignment.Cost8 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Cost9"
            If oAssignment.Cost9 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Cost10"
            If oAssignment.Cost10 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date1"
            If oAssignment.Date1 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date2"
            If oAssignment.Date2 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date3"
            If oAssignment.Date3 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date4"
            If oAssignment.Date4 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date5"
            If oAssignment.Date5 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date6"
            If oAssignment.Date6 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date7"
            If oAssignment.Date7 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date8"
            If oAssignment.Date8 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date9"
            If oAssignment.Date9 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Date10"
            If oAssignment.Date10 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration1"
            If oAssignment.Duration1 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration2"
            If oAssignment.Duration2 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration3"
            If oAssignment.Duration3 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration4"
            If oAssignment.Duration4 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration5"
            If oAssignment.Duration5 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration6"
            If oAssignment.Duration6 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration7"
            If oAssignment.Duration7 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration8"
            If oAssignment.Duration8 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration9"
            If oAssignment.Duration9 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Duration10"
            If oAssignment.Duration10 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish1"
            If oAssignment.Finish1 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish2"
            If oAssignment.Finish2 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish3"
            If oAssignment.Finish3 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish4"
            If oAssignment.Finish4 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish5"
            If oAssignment.Finish5 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish6"
            If oAssignment.Finish6 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish7"
            If oAssignment.Finish7 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish8"
            If oAssignment.Finish8 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish9"
            If oAssignment.Finish9 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Finish10"
            If oAssignment.Finish10 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag1"
            If oAssignment.Flag1 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag2"
            If oAssignment.Flag2 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag3"
            If oAssignment.Flag3 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag4"
            If oAssignment.Flag4 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag5"
            If oAssignment.Flag5 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag6"
            If oAssignment.Flag6 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag7"
            If oAssignment.Flag7 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag8"
            If oAssignment.Flag8 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag9"
            If oAssignment.Flag9 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag10"
            If oAssignment.Flag10 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag11"
            If oAssignment.Flag11 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag12"
            If oAssignment.Flag12 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag13"
            If oAssignment.Flag13 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag14"
            If oAssignment.Flag14 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag15"
            If oAssignment.Flag15 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag16"
            If oAssignment.Flag16 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag17"
            If oAssignment.Flag17 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag18"
            If oAssignment.Flag18 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag19"
            If oAssignment.Flag19 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Flag20"
            If oAssignment.Flag20 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number1"
            If oAssignment.Number1 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number2"
            If oAssignment.Number2 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number3"
            If oAssignment.Number3 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number4"
            If oAssignment.Number4 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number5"
            If oAssignment.Number5 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number6"
            If oAssignment.Number6 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number7"
            If oAssignment.Number7 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number8"
            If oAssignment.Number8 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number9"
            If oAssignment.Number9 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number10"
            If oAssignment.Number10 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number11"
            If oAssignment.Number11 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number12"
            If oAssignment.Number12 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number13"
            If oAssignment.Number13 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number14"
            If oAssignment.Number14 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number15"
            If oAssignment.Number15 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number16"
            If oAssignment.Number16 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number17"
            If oAssignment.Number17 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number18"
            If oAssignment.Number18 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number19"
            If oAssignment.Number19 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Number20"
            If oAssignment.Number20 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode1"
            If oAssignment.OutlineCode1 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode2"
            If oAssignment.OutlineCode2 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode3"
            If oAssignment.OutlineCode3 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode4"
            If oAssignment.OutlineCode4 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode5"
            If oAssignment.OutlineCode5 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode6"
            If oAssignment.OutlineCode6 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode7"
            If oAssignment.OutlineCode7 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode8"
            If oAssignment.OutlineCode8 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode9"
            If oAssignment.OutlineCode9 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "OutlineCode10"
            If oAssignment.OutlineCode10 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start1"
            If oAssignment.Start1 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start2"
            If oAssignment.Start2 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start3"
            If oAssignment.Start3 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start4"
            If oAssignment.Start4 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start5"
            If oAssignment.Start5 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start6"
            If oAssignment.Start6 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start7"
            If oAssignment.Start7 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start8"
            If oAssignment.Start8 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start9"
            If oAssignment.Start9 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Start10"
            If oAssignment.Start10 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text1"
            If oAssignment.Text1 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text2"
            If oAssignment.Text2 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text3"
            If oAssignment.Text3 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text4"
            If oAssignment.Text4 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text5"
            If oAssignment.Text5 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text6"
            If oAssignment.Text6 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text7"
            If oAssignment.Text7 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text8"
            If oAssignment.Text8 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text9"
            If oAssignment.Text9 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text10"
            If oAssignment.Text10 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text11"
            If oAssignment.Text11 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text12"
            If oAssignment.Text12 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text13"
            If oAssignment.Text13 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text14"
            If oAssignment.Text14 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text15"
            If oAssignment.Text15 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text16"
            If oAssignment.Text16 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text17"
            If oAssignment.Text17 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text18"
            If oAssignment.Text18 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text19"
            If oAssignment.Text19 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text20"
            If oAssignment.Text20 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text21"
            If oAssignment.Text21 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text22"
            If oAssignment.Text22 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text23"
            If oAssignment.Text23 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text24"
            If oAssignment.Text24 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text25"
            If oAssignment.Text25 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text26"
            If oAssignment.Text26 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text27"
            If oAssignment.Text27 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text28"
            If oAssignment.Text28 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text29"
            If oAssignment.Text29 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
          Case "Text30"
            If oAssignment.Text30 <> strValue Then
              oToFix.Add lngAssignmentUID, strValue
              lngMismatched = lngMismatched + 1
            End If
        End Select
      Else 'ECF
        'ViewApply "Task Usage" 'status sheet already applies this
        FilterClear 'status sheet already runs this
        'GroupClear 'status sheet already runs this and the next line (I think)
        Application.OptionsViewEx DisplaySummaryTasks:=True, DisplayExternalSuccessors:=False, DisplayExternalPredecessors:=False
        Sort "ID", , , , , , False, True 'OutlineShowAllTasks won't work without this
        SelectAll 'is this necessary?
        OutlineShowAllTasks 'doesn't status sheet already do this?
        'todo: revert view/table/ 'might not be necessary
        'Assignment.GetField() doesn't exist, so we have to use this...
        If Find("Unique ID", "equals", lngAssignmentUID) Then
          'need to select it to use CheckField: must be expanded
          If Not CheckField(strKey, strValue, "equals") Then
            lngMismatched = lngMismatched + 1
            oToFix.Add lngAssignmentUID, strValue
          End If
        End If
      End If
      lngAssignments = lngAssignments + 1
    Next oAssignment
next_task:
    lngTasks = lngTasks + 1
    'todo: if userform is showing then update the status/progress
    'todo: running tally of mismatches in status bar, with percent progress
    Application.StatusBar = "Analyzing: " & Format(lngTasks, "#,##0") & " task(s); " & Format(lngAssignments, "#,##0") & " assignment(s) (" & Format(lngMismatched, "#,##0") & " mismatched)"
  Next oTask
  Application.StatusBar = "Complete: " & Format(lngMismatched, "#,##0") & " mismatch(es)."
  If lngMismatched > 0 Then
    strMsg = "The selected For Each field '" & strCFN & " (" & strKey & ")' has not properly been rolled down from Tasks to Assignments." & vbCrLf & vbCrLf
    If lngMismatched = 1 Then
      strMsg = strMsg & Format(lngMismatched, "#,##0") & " assignment does not match its parent task--this will cause issues in the export." & vbCrLf & vbCrLf
    Else
      strMsg = strMsg & Format(lngMismatched, "#,##0") & " assignments do not match their parent tasks--this will cause issues in the export." & vbCrLf & vbCrLf
    End If
    strMsg = strMsg & "Would you like to roll-down '" & strCFN & " (" & strKey & ")' from Tasks to Assignments now?"
    If MsgBox(strMsg, vbQuestion + vbYesNo, "ForEach Rolldown Error") = vbYes Then
      If blnLCF Then
        If MsgBox("Set Custom Field '" & strCFN & "' to automatically roll-down values in the future?", vbQuestion + vbYesNo, "Recommendation: Yes") = vbYes Then
          Application.CustomFieldPropertiesEx FieldID:=lngFieldID, AutomaticallyRolldownToAssn:=True
        End If
      Else 'ECF
        MsgBox "Consider asking your PWA administrator to set '" & strCFN & "' to automatically roll-down values in the future.", vbInformation + vbOKOnly, "Recommendation:"
      End If
      For lngItem = 0 To oToFix.Count - 1
        If Not SetMatchingField(strKey, oToFix.Items(lngItem), "Unique ID", oToFix.Keys(lngItem)) Then
          lngNotFixed = lngNotFixed + 1
          strNotFixed = strNotFixed & oToFix.Keys(lngItem) & ","
        Else
          Application.StatusBar = "Fixing Assignments: " & Format((lngItem + 1 - lngNotFixed) / oToFix.Count, "0%")
        End If
      Next lngItem
      If lngNotFixed > 0 Then
        strMsg = "Error(s) encountered attempting to roll-down Task-Level values to Assignment-Level values in the field '" & strCFN & "'!" & vbCrLf & vbCrLf
        If lngNotFixed <= 6 Then 'just show an Input box (hijacked to be an export box)
          strMsg = strMsg & "Review the following ASSIGNMENT UNIQUE IDs:"
          InputBox strMsg, "ForEach Rolldown", Left(strNotFixed, Len(strNotFixed) - 1)
        Else
          strMsg = strMsg & "Review the ASSIGNMENT UNIQUE IDs in the error log that will appear next."
          MsgBox strMsg, vbExclamation + vbOKOnly, "ForEach Rolldown"
          Dim lngFile As Long
          Dim strFileName As String
          lngFile = FreeFile
          strFileName = Environ("tmp") & "\cpt-rolldown-errors.txt"
          Open strFileName For Output As #lngFile
          Print #lngFile, "ASSIGNMENT UNIQUE IDs: " & Left(strNotFixed, Len(strNotFixed) - 1)
          Print #lngFile, vbCrLf & "HINT: Use FilterByClipboard to find and fix (try it with the 'Filter' option unchecked)"
          Close #lngFile
          ShellExecute 0, "open", "notepad.exe", strFileName, vbNullString, 1
        End If
        blnValid = False
      Else
        Application.StatusBar = "Complete: " & Format(oToFix.Count, "#,##0") & " assignment(s) fixed."
        'capture alert settings
        Dim blnDisplayAlerts As Boolean
        Dim blnDisplayScheduleMessages As Boolean
        blnDisplayAlerts = Application.DisplayAlerts
        blnDisplayScheduleMessages = Application.DisplayScheduleMessages
        'turn off alerts
        Application.DisplayAlerts = False
        Application.DisplayScheduleMessages = False
        Application.CalculateProject 'triggers refresh so fixes can be seen
        'restore alert settings
        Application.DisplayAlerts = blnDisplayAlerts
        Application.DisplayScheduleMessages = blnDisplayScheduleMessages
        MsgBox "The selected For Each field '" & strCFN & " (" & strKey & ")' has been successfully rolled down from Tasks to Assignments.", vbInformation + vbOKOnly, "ForEach Rolldown"
        blnValid = True
      End If
    Else
      blnValid = False
    End If
  Else
    blnValid = True
  End If
  
  cptValidateAssignmentRolldown = blnValid
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oToFix = Nothing
  cptSpeed False
  Set oAssignment = Nothing
  Set oTask = Nothing
  Reset
  Exit Function
  
err_here:
  On Error Resume Next
  cptValidateAssignmentRolldown = False
  Call cptHandleErr("cptStatusSheet_bas", "cptValidateAssignmentRolldown", Err, Erl)
  Resume exit_here
End Function

