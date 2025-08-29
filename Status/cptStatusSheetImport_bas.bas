Attribute VB_Name = "cptStatusSheetImport_bas"
'<cpt_version>v1.3.3</cpt_version>
Option Explicit
Private rBad As Excel.Range
Private oBad As Scripting.Dictionary

Sub cptShowStatusSheetImport_frm()
  'objects
  Dim myStatusSheetImport_frm As cptStatusSheetImport_frm
  Dim rst As Object 'ADODB.Recordset
  'strings
  Dim strKickoutReport As String
  Dim strImportLog As String
  Dim strAppend As String
  Dim strTaskUsage As String
  Dim strAppendTo As String
  Dim strETC As String
  Dim strEVP As String
  Dim strFF As String
  Dim strFS As String
  Dim strAF As String
  Dim strAS As String
  Dim strGUID As String
  Dim strSettings As String
  Dim strCustomFieldName As String
  'longs
  Dim lngETC As Long
  Dim lngEVP As Long
  Dim lngFF As Long
  Dim lngFS As Long
  Dim lngAF As Long
  Dim lngAS As Long
  Dim lngField As Long
  'integers
  Dim intField As Integer
  'doubles
  'booleans
  Dim blnRename As Boolean
  Dim blnErrorTrapping As Boolean
  Dim blnTaskUsageBelow As Boolean
  'variants
  Dim vField As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'todo: start Excel in the background if not already open; close on form close
  
  'ensure settings
  If Not cptValidMap("EVP,EVT,LOE") Then
    MsgBox "Settings not saved; cannot proceed.", vbExclamation + vbOKOnly, "Settings Required"
    GoTo exit_here
  End If
  
  'populate comboboxes
  Set myStatusSheetImport_frm = New cptStatusSheetImport_frm
  With myStatusSheetImport_frm
    .Caption = "Import Status Sheets (" & cptGetVersion("cptStatusSheetImport_frm") & ")"
    .cboAppendTo.Clear
    .cboAppendTo.AddItem "Bottom of Task Note"
    .cboAppendTo.AddItem "Top of Task Note"
    
    'start
    For Each vField In Array("Start", "Date")
      For intField = 1 To 10
        lngField = FieldNameToFieldConstant(vField & intField, pjTask)
        strCustomFieldName = CustomFieldGetName(lngField)
        .cboAS.AddItem
        .cboAS.List(.cboAS.ListCount - 1, 0) = lngField
        .cboAS.List(.cboAS.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
        .cboFS.AddItem
        .cboFS.List(.cboFS.ListCount - 1, 0) = lngField
        .cboFS.List(.cboFS.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
      Next intField
    Next vField
    'direct import to Actual Start removed in v1.3.0
    
    'finish
    For Each vField In Array("Finish", "Date")
      For intField = 1 To 10
        lngField = FieldNameToFieldConstant(vField & intField, pjTask)
        strCustomFieldName = CustomFieldGetName(lngField)
        .cboAF.AddItem
        .cboAF.List(.cboAF.ListCount - 1, 0) = lngField
        .cboAF.List(.cboAF.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
        .cboFF.AddItem
        .cboFF.List(.cboFF.ListCount - 1, 0) = lngField
        .cboFF.List(.cboFF.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
      Next intField
    Next vField
    'direct import to Actual Finish removed in v1.3.0
    
    'ev% and etc
    For Each vField In Array("Number")
      For intField = 1 To 20
        lngField = FieldNameToFieldConstant(vField & intField, pjTask)
        strCustomFieldName = CustomFieldGetName(lngField)
        .cboEV.AddItem
        .cboEV.List(.cboEV.ListCount - 1, 0) = lngField
        .cboEV.List(.cboEV.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
        .cboETC.AddItem
        .cboETC.List(.cboETC.ListCount - 1, 0) = lngField
        .cboETC.List(.cboETC.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
      Next intField
    Next vField
        
    'add enterprise custom fields? -> no
        
    'convert legacy user settings
    strSettings = cptDir & "\settings\cpt-status-sheet-import.adtg"
    If Dir(strSettings) <> vbNullString Then
      'import user settings
      Set rst = CreateObject("ADODB.Recordset")
      rst.Open strSettings
      If Not rst.EOF Then
        cptSaveSetting "StatusSheetImport", "cboAS", CStr(rst("AS"))
        cptSaveSetting "StatusSheetImport", "cboAF", CStr(rst("AF"))
        cptSaveSetting "StatusSheetImport", "cboFS", CStr(rst("FS"))
        cptSaveSetting "StatusSheetImport", "cboFF", CStr(rst("FF"))
        cptSaveSetting "StatusSheetImport", "cboEV", CStr(rst("EV"))
        cptSaveSetting "StatusSheetImport", "cboETC", CStr(rst("ETC"))
        cptSaveSetting "StatusSheetImport", "chkAppend", CStr(rst("Append"))
        If rst("AppendTo") <> "" Then
          cptSaveSetting "StatusSheetImport", "cboAppendTo", CStr(rst("AppendTo"))
        Else
          cptSaveSetting "StatusSheetImport", "cboAppendTo", "Top of Task Note"
        End If
      End If
      Kill strSettings
    Else
      'default settings
      .cboAppendTo.Value = "Top of Task Note"
    End If

    'import user settings
    blnRename = False
    strAS = cptGetSetting("StatusSheetImport", "cboAS")
    If Len(strAS) > 0 Then
      If strAS = CStr(FieldNameToFieldConstant("Actual Start")) Then
        MsgBox "Direct import to Actual Start is no longer supported. Please select a different field.", vbExclamation + vbOKOnly, "Actual Start"
      Else
        lngAS = CLng(strAS)
        .cboAS.Value = lngAS
      End If
      If CustomFieldGetName(lngAS) = "" Then blnRename = True
    Else
      blnRename = True
    End If
    
    strAF = cptGetSetting("StatusSheetImport", "cboAF")
    If Len(strAF) > 0 Then
      If strAF = CStr(FieldNameToFieldConstant("Actual Finish")) Then
        MsgBox "Direct import to Actual Finish is no longer supported. Please select a different field.", vbExclamation + vbOKOnly, "Actual Finish"
      Else
        lngAF = CLng(strAF)
        .cboAF.Value = lngAF
      End If
      If CustomFieldGetName(lngAF) = "" Then blnRename = True
    Else
      blnRename = True
    End If
    
    strFS = cptGetSetting("StatusSheetImport", "cboFS")
    If Len(strFS) > 0 Then
      lngFS = CLng(strFS)
      .cboFS.Value = lngFS
      If CustomFieldGetName(lngFS) = "" Then blnRename = True
    Else
      blnRename = True
    End If
    
    strFF = cptGetSetting("StatusSheetImport", "cboFF")
    If Len(strFF) > 0 Then
      lngFF = CLng(strFF)
      .cboFF.Value = lngFF
      If CustomFieldGetName(lngFS) = "" Then blnRename = True
    Else
      blnRename = True
    End If
    
    strEVP = cptGetSetting("StatusSheetImport", "cboEV")
    If Len(strEVP) > 0 Then
      lngEVP = CLng(strEVP)
      .cboEV.Value = lngEVP
      If CustomFieldGetName(lngEVP) = "" Then blnRename = True
    Else
      blnRename = True
    End If
    
    strETC = cptGetSetting("StatusSheetImport", "cboETC")
    If Len(strETC) > 0 Then
      lngETC = CLng(strETC)
      .cboETC.Value = lngETC
      If CustomFieldGetName(lngEVP) = "" Then blnRename = True
    Else
      blnRename = True
    End If
    
    strAppend = cptGetSetting("StatusSheetImport", "chkAppend")
    If Len(strAppend) > 0 Then .chkAppend = CBool(strAppend)
    strAppendTo = cptGetSetting("StatusSheetImport", "cboAppendTo")
    If Len(strAppendTo) > 0 Then .cboAppendTo.Value = strAppendTo
    strImportLog = cptGetSetting("StatusSheetImport", "chkImportLog")
    If Len(strImportLog) > 0 Then
      .chkImportLog = CBool(strImportLog)
    Else
      .chkImportLog = True 'default
    End If
    strKickoutReport = cptGetSetting("StatusSheetImport", "chkKickoutReport")
    If Len(strKickoutReport) > 0 Then
      .chkKickoutReport = CBool(strKickoutReport)
    Else
      .chkKickoutReport = True 'default
    End If
    'refresh which view
    strTaskUsage = cptGetSetting("StatusSheetImport", "optTaskUsage")
    If Len(strTaskUsage) > 0 Then
      If strTaskUsage = "above" Then
        .optAbove = True
        blnTaskUsageBelow = False
      ElseIf strTaskUsage = "below" Then
        .optBelow = True
        blnTaskUsageBelow = True
      End If
    Else
      .optBelow = True
      blnTaskUsageBelow = True
    End If
    .cmdRename.Visible = blnRename
    .cmdRemove.Enabled = False
    'show the form
    .Show (False)
    Call cptRefreshStatusImportTable(myStatusSheetImport_frm, blnTaskUsageBelow)

  End With
  
'  ActiveWindow.TopPane.Activate
'  If blnTaskUsageBelow Then
'    ViewApply "Task Entry"
'  Else
'    ViewApply "Task Usage"
'  End If
'  Call cptRefreshStatusImportTable(blnTaskUsageBelow)
  

exit_here:
  On Error Resume Next
  Set rst = Nothing
  Set myStatusSheetImport_frm = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_bas", "cptShowStatusSheetImport_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptStatusSheetImport(ByRef myStatusSheetImport_frm As cptStatusSheetImport_frm)
  'objects
  Dim oInspector As Outlook.Inspector
  Dim oOutlook As Outlook.Application
  Dim oMailItem As Outlook.MailItem
  Dim oDocument As Word.Document
  Dim oWord As Word.Application
  Dim oSelection As Word.Selection
  Dim oEmailTemplate As Word.Template
  Dim oDict As Scripting.Dictionary
  Dim oShell As Object
  Dim oRecordset As ADODB.Recordset
  Dim oSubproject As MSProject.SubProject
  Dim oTask As MSProject.Task
  Dim oResource As MSProject.Resource
  Dim oAssignment As MSProject.Assignment
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oListObject As Excel.ListObject
  Dim oRange As Excel.Range
  Dim oComboBox As MSForms.ComboBox
  Dim rst As ADODB.Recordset
  'strings
  Dim strAS As String
  Dim strAF As String
  Dim strNS As String
  Dim strNF As String
  Dim strETC As String
  Dim strUIDList As String
  Dim strLOE As String
  Dim strEVT As String
  Dim strHeader As String
  Dim strCon As String
  Dim strSQL As String
  Dim strDeconflictionFile As String
  Dim strSchema As String
  Dim strEVP As String
  Dim strFileName As String
  Dim strNotesColTitle As String
  Dim strImportLog As String
  Dim strAppendTo As String
  Dim strSettings As String
  Dim strGUID As String
  'longs
  Dim lngUID As Long
  Dim lngBadItem As Long
  Dim lngEVT As Long 'EVT LCF
  Dim lngMultiplier As Long
  Dim lngDeconflictionFile As Long
  Dim lngEVP As Long
  Dim lngTaskNameCol As Long
  Dim lngEVTCol As Long
  Dim lngEVPCol As Long
  Dim lngUIDCol As Long
  Dim lngFile As Long
  Dim lngRow As Long
  Dim lngCommentsCol As Long
  Dim lngETCCol As Long
  Dim lngNFCol As Long 'new finish
  Dim lngNSCol As Long 'new start
  Dim lngHeaderRow As Long
  Dim lngLastRow As Long
  Dim lngItem As Long
  Dim lngETC As Long
  Dim lngEV As Long
  Dim lngFF As Long
  Dim lngFS As Long
  Dim lngAF As Long
  Dim lngAS As Long
  'integers
  'doubles
  Dim dblWas As Double
  Dim dblETC As Double
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnKickoutReport As Boolean
  Dim blnImportLog As Boolean
  Dim blnAppend As Boolean
  Dim blnTask As Boolean
  Dim blnValid As Boolean
  'variants
  Dim vKey As Variant
  Dim vField As Variant
  Dim vControl As Variant
  'dates
  Dim dtStart As Date
  Dim dtNewDate As Date
  Dim dtStatus As Date

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'validate choices for all
  With myStatusSheetImport_frm
    
    blnValid = True
    
    'ensure file(s) are added to import list
    If .lboStatusSheets.ListCount = 0 Then
      MsgBox "Please select one or more files to import.", vbInformation + vbOKOnly, "No Files Found"
      blnValid = False
      GoTo exit_here
    End If
    
    'ensure import fields are selected
    For Each vControl In Array("cboAS", "cboAF", "cboFS", "cboFF", "cboEV", "cboETC", "cboAppendTo")
      'reset border color
      Set oComboBox = .Controls(vControl)
      oComboBox.BorderColor = -2147483642
      If IsNull(oComboBox) And oComboBox.Enabled Then
        oComboBox.BorderColor = 192 'red
        blnValid = False
      End If
    Next vControl
    
    'warn if invalid
    If Not blnValid Then
      MsgBox "Please select import fields.", vbExclamation + vbOKOnly, "Invalid Import Fields"
      GoTo exit_here
    End If
    
  End With
  
  'speed up
  cptSpeed True
  
  'capture import fields and settings
  With myStatusSheetImport_frm
    lngAS = .cboAS.Value
    lngAF = .cboAF.Value
    lngFS = .cboFS.Value
    lngFF = .cboFF.Value
    lngEV = .cboEV.Value
    lngETC = .cboETC.Value
    blnAppend = .chkAppend
    strAppendTo = .cboAppendTo
    blnImportLog = .chkImportLog
    blnKickoutReport = .chkKickoutReport
  End With
  
  'save user settings
  cptSaveSetting "StatusSheetImport", "cboAS", CStr(lngAS)
  cptSaveSetting "StatusSheetImport", "cboAF", CStr(lngAF)
  cptSaveSetting "StatusSheetImport", "cboFS", CStr(lngFS)
  cptSaveSetting "StatusSheetImport", "cboFF", CStr(lngFF)
  cptSaveSetting "StatusSheetImport", "cboEV", CStr(lngEV)
  cptSaveSetting "StatusSheetImport", "cboETC", CStr(lngETC)
  cptSaveSetting "StatusSheetImport", "chkAppend", IIf(blnAppend, 1, 0)
  cptSaveSetting "StatusSheetImport", "cboAppendTo", strAppendTo
    
  'get LOE settings
  strEVT = cptGetSetting("Integration", "EVT")
  strLOE = cptGetSetting("Integration", "LOE")
    
  'set up import log file
  If Left(ActiveProject.Path, 2) = "<>" Or Left(ActiveProject.Path, 4) = "http" Then 'server project: default to Desktop
    Set oShell = CreateObject("WScript.Shell")
    strImportLog = oShell.SpecialFolders("Desktop") & "\cpt-import-log-" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".txt"
  Else  'not a server project: use ActiveProject.Path
    strImportLog = ActiveProject.Path & "\cpt-import-log-" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".txt"
  End If
  
  lngFile = FreeFile
  Open strImportLog For Output As #lngFile
  'log action
  Print #lngFile, "STATUS SHEET IMPORT LOG"
  dtStart = Now
  Print #lngFile, "START: " & FormatDateTime(dtStart, vbGeneralDate)
  'set up deconfliction db
  strSchema = Environ("temp") & "\Schema.ini"
  lngDeconflictionFile = FreeFile
  Open strSchema For Output As lngDeconflictionFile
  Print #lngDeconflictionFile, "[imported.csv]"
  Print #lngDeconflictionFile, "Format=CSVDelimited"
  Print #lngDeconflictionFile, "ColNameHeader=True"
  Print #lngDeconflictionFile, "Col1=FILE Text Width 255"
  Print #lngDeconflictionFile, "Col2=TASK_UID Integer"
  Print #lngDeconflictionFile, "Col3=FIELD Text Width 100"
  Print #lngDeconflictionFile, "Col4=RESOURCE_NAME Text Width 150"
  Print #lngDeconflictionFile, "Col5=WAS Text Width 50"
  Print #lngDeconflictionFile, "Col6=IS Text Width 50"
  Close #lngDeconflictionFile
  strDeconflictionFile = Environ("temp") & "\imported.csv"
  lngDeconflictionFile = FreeFile
  Open strDeconflictionFile For Output As #lngDeconflictionFile
  Print #lngDeconflictionFile, "FILE,TASK_UID,FIELD,RESOURCE_NAME,WAS,IS"
  
  'clear existing values from selected import fields
  myStatusSheetImport_frm.lblStatus = "Clearing previous values..."
  'clear out previous values tasks
  strAS = FieldConstantToFieldName(lngAS)
  strAF = FieldConstantToFieldName(lngAF)
  strNS = FieldConstantToFieldName(lngFS)
  strNF = FieldConstantToFieldName(lngFF)
  strEVP = FieldConstantToFieldName(lngEVP)
  strETC = FieldConstantToFieldName(lngETC)
  
  ActiveWindow.TopPane.Activate
  ViewApply "Task Usage"
  FilterClear
  'GroupClear 'todo: capture group name
  OptionsViewEx DisplaySummaryTasks:=True
  SelectAll
  OutlineShowAllTasks
  Sort "ID", Renumber:=False, Outline:=True
  SelectAll
  SetField strAS, ""
  SetField strAF, ""
  SetField strNS, ""
  SetField strNF, ""
  SetField strEVP, ""
  SetField strETC, ""
  myStatusSheetImport_frm.lblStatus.Caption = "Clearing previous values...done."
  ViewApply "cptStatusSheetImport View"
  
  'set up array of updated  UIDs
  Set oDict = CreateObject("Scripting.Dictionary")
  Set oBad = CreateObject("Scripting.Dictionary")
  'set up excel
  Set oExcel = CreateObject("Excel.Application")
  With myStatusSheetImport_frm
    .lblStatus.Caption = "Importing..."
    For lngItem = 0 To .lboStatusSheets.ListCount - 1
      blnValid = True
      strFileName = .lboStatusSheets.List(lngItem, 0) & .lboStatusSheets.List(lngItem, 1)
      Set oWorkbook = oExcel.Workbooks.Open(strFileName, ReadOnly:=True)
      .lboStatusSheets.Selected(lngItem) = True
      DoEvents
      Print #lngFile, String(25, "=")
      Print #lngFile, "IMPORTING Workbook: " & strFileName & " (" & oWorkbook.Sheets.Count & " Worksheets)"
      Print #lngFile, String(25, "-")
      For Each oWorksheet In oWorkbook.Sheets
        Set rBad = Nothing
        oBad.RemoveAll
        If oWorksheet.Name = "Conditional Formatting" Then
          Print #lngFile, "SKIPPING Worksheet: " & oWorksheet.Name
          GoTo next_worksheet
        End If
        
        Print #lngFile, "IMPORTING Worksheet: " & oWorksheet.Name
        DoEvents
        If oBad.Count > 0 Then oBad.RemoveAll
        'unhide columns and rows (sort is blocked by sheet protection...)
        oWorksheet.Columns.Hidden = False
        oWorksheet.Rows.Hidden = False
        
        'get status date
        On Error Resume Next
        dtStatus = oWorksheet.Range("STATUS_DATE")
        If Err.Number = 1004 Then 'invalid oWorkbook
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          Print #lngFile, "INVALID Worksheet - range 'STATUS_DATE' not found"
          GoTo next_worksheet
        End If
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        'get header row
        lngUIDCol = 1
        On Error Resume Next
        lngHeaderRow = WorksheetFunction.Match("UID", oWorksheet.Columns(lngUIDCol), 0)
        If Err.Number = 1004 Then 'invalid oWorkbook
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          Print #lngFile, "INVALID Worksheet - UID column not found"
          GoTo next_worksheet
        End If
        'get header columns
        lngTaskNameCol = WorksheetFunction.Match("Task Name / Scope", oWorksheet.Rows(lngHeaderRow))
        lngEVTCol = WorksheetFunction.Match("EVT", oWorksheet.Rows(lngHeaderRow), 0)
        lngNSCol = WorksheetFunction.Match("New Forecast/ Actual Start", oWorksheet.Rows(lngHeaderRow), 0)
        lngNFCol = WorksheetFunction.Match("New Forecast/ Actual Finish", oWorksheet.Rows(lngHeaderRow), 0)
        lngEVPCol = WorksheetFunction.Match("New EV%", oWorksheet.Rows(lngHeaderRow), 0)
        lngETCCol = WorksheetFunction.Match("ETC", oWorksheet.Rows(lngHeaderRow), 0) + 1
        strNotesColTitle = cptGetSetting("StatusSheet", "txtNotesColTitle")
        On Error Resume Next
        If Len(strNotesColTitle) > 0 Then
          lngCommentsCol = WorksheetFunction.Match(strNotesColTitle, oWorksheet.Rows(lngHeaderRow), 0)
        Else
          lngCommentsCol = WorksheetFunction.Match("Reason / Action / Impact", oWorksheet.Rows(lngHeaderRow), 0)
        End If
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If lngCommentsCol = 0 Then
          lngCommentsCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
        End If
        'get last row
        lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row
        'pull in the data
        For lngRow = lngHeaderRow + 1 To lngLastRow
          'summary lines and group summaries have UID = 0
          If oWorksheet.Cells(lngRow, lngUIDCol).Value = 0 Then GoTo next_row
          
          'is this an assignment row?
          Set oAssignment = Nothing
          On Error Resume Next
          Set oAssignment = ActiveProject.Tasks(1).Assignments.UniqueID(oWorksheet.Cells(lngRow, lngUIDCol).Value)
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If Not oAssignment Is Nothing Then
            Set oResource = Nothing
            On Error Resume Next
            Set oResource = ActiveProject.Resources(oWorksheet.Cells(lngRow, lngTaskNameCol).Value)
            If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
            blnTask = False
            GoTo do_stuff
          End If
          
          'is this a task row?
          Set oTask = Nothing
          On Error Resume Next
          Set oTask = ActiveProject.Tasks.UniqueID(oWorksheet.Cells(lngRow, lngUIDCol).Value)
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oTask Is Nothing Then 'it's nothing: notify user
            Print #lngFile, "UID " & oWorksheet.Cells(lngRow, lngUIDCol) & " in Status Sheet not found in IMS! Note: could be a missing Task or a missing Resource/Assignment."
            cptAddRange rBad, oWorksheet.Cells(lngRow, lngUIDCol)
            GoTo next_row
          Else
            blnTask = True
          End If
          
do_stuff:
          'set Task
          If blnTask Then
            lngUID = oTask.UniqueID
            'skip completed tasks (which are also italicized)
            If IsDate(oTask.ActualFinish) Then
              If FormatDateTime(oTask.ActualFinish, vbShortDate) = FormatDateTime(oWorksheet.Cells(lngRow, lngNFCol).Value, vbShortDate) Then GoTo next_row
            End If

            'new start date
            If Not oWorksheet.Cells(lngRow, lngNSCol).Locked Then
              If oWorksheet.Cells(lngRow, lngNSCol).DisplayFormat.Interior.Color = 13551615 Then '13551615 = Style = "Bad"
                Print #lngFile, "UID " & lngUID & " - Should Have Started " & String(10, "<")
                cptAddRange rBad, oWorksheet.Cells(lngRow, lngUIDCol)
                blnValid = False
                GoTo skip_ns
              End If
              If oWorksheet.Cells(lngRow, lngNSCol).DisplayFormat.Interior.Color = 10079487 Then '10079487 = Style = "Input"
                If Len(oWorksheet.Cells(lngRow, lngNSCol)) = 0 Then
                  Print #lngFile, "UID " & lngUID & " = invalid New Start Date " & String(10, "<")
                  cptAddRange rBad, oWorksheet.Cells(lngRow, lngUIDCol)
                  blnValid = False
                  GoTo skip_ns
                End If
              End If
              oWorksheet.Cells(lngRow, lngNSCol).NumberFormat = "0.00" 'work around overflow issue
              If oWorksheet.Cells(lngRow, lngNSCol).Value > 91312 Then 'dates greater than 12/31/2149 are invalid
                cptAddRange rBad, oWorksheet.Cells(lngRow, lngUIDCol)
                Print #lngFile, "UID " & lngUID & " - invalid New Start Date " & String(10, "<")
              Else
                oWorksheet.Cells(lngRow, lngNSCol).NumberFormat = "m/d/yyyy" 'restore date format
                If Len(oWorksheet.Cells(lngRow, lngNSCol).Value) > 0 And Not IsDate(oWorksheet.Cells(lngRow, lngNSCol).Value) Then
                  Print #lngFile, "UID " & lngUID & " - invalid New Start Date " & String(10, "<")
                ElseIf oWorksheet.Cells(lngRow, lngNSCol).Value > 0 Then
                  dtNewDate = FormatDateTime(CDate(oWorksheet.Cells(lngRow, lngNSCol).Value), vbShortDate)
                  'determine actual or forecast
                  If dtNewDate <= FormatDateTime(dtStatus, vbShortDate) Then 'actual start
                    If IsDate(oTask.ActualStart) Then
                      If FormatDateTime(oTask.ActualStart, vbShortDate) <> dtNewDate Then oTask.SetField lngAS, CDate(dtNewDate & " 08:00 AM")
                    Else
                      oTask.SetField lngAS, CDate(dtNewDate & " 08:00 AM")
                    End If
                    Print #lngFile, "UID " & lngUID & " AS > " & FormatDateTime(dtNewDate, vbShortDate)
                    If Not oDict.Exists(lngUID) Then oDict.Add lngUID, lngUID
                  ElseIf dtNewDate > dtStatus Then 'forecast start
                    If FormatDateTime(oTask.Start, vbShortDate) <> dtNewDate Then
                      If dtNewDate > #12/31/2149# Then
                        Print #lngFile, "UID " & lngUID & " FS > ERROR (" & FormatDateTime(dtNewDate, vbShortDate) & " is outside allowable range)"
                      Else
                        oTask.SetField lngFS, CDate(dtNewDate & " 08:00 AM")
                        Print #lngFile, "UID " & lngUID & " FS > " & FormatDateTime(dtNewDate, vbShortDate)
                        If Not oDict.Exists(lngUID) Then oDict.Add lngUID, lngUID
                      End If
                    End If
                  End If
                  If FormatDateTime(dtNewDate, vbShortDate) <> FormatDateTime(oTask.Start, vbShortDate) Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, "START", "", CStr(FormatDateTime(oTask.Start, vbShortDate)), CStr(FormatDateTime(dtNewDate, vbShortDate))), ",")
                  End If
                End If
              End If
skip_ns:
              oWorksheet.Cells(lngRow, lngNSCol).NumberFormat = "m/d/yyyy" 'restore date format
            End If
            
            'new finish date
            If Not oWorksheet.Cells(lngRow, lngNFCol).Locked Then
              If oWorksheet.Cells(lngRow, lngNFCol).DisplayFormat.Interior.Color = 13551615 Then 'red = invalid
                Print #lngFile, "UID " & lngUID & " = invalid New Finish Date " & String(10, "<")
                cptAddRange rBad, oWorksheet.Cells(lngRow, lngUIDCol)
                blnValid = False
                GoTo skip_nf
              End If
              If oWorksheet.Cells(lngRow, lngNFCol).DisplayFormat.Interior.Color = 10079487 Then 'orange = input
                If Len(oWorksheet.Cells(lngRow, lngNFCol)) = 0 Then
                  Print #lngFile, "UID " & lngUID & " = invalid New Finish Date " & String(10, "<")
                  cptAddRange rBad, oWorksheet.Cells(lngRow, lngUIDCol)
                  blnValid = False
                  GoTo skip_nf
                End If
              End If
              
              oWorksheet.Cells(lngRow, lngNFCol).NumberFormat = "0.00" 'work around overflow issue
              If oWorksheet.Cells(lngRow, lngNFCol).Value > 91312 Then 'invalid
                Print #lngFile, "UID " & lngUID & " - invalid New Finish Date " & String(10, "<")
                cptAddRange rBad, oWorksheet.Cells(lngRow, lngUIDCol)
                blnValid = False
              Else
                oWorksheet.Cells(lngRow, lngNFCol).NumberFormat = "m/d/yyyy" 'restore date format
                If Len(oWorksheet.Cells(lngRow, lngNFCol).Value) > 0 And Not IsDate(oWorksheet.Cells(lngRow, lngNFCol).Value) Then
                  Print #lngFile, "UID " & lngUID & " - invalid New Finish Date " & String(10, "<")
                  cptAddRange rBad, oWorksheet.Cells(lngRow, lngNFCol)
                ElseIf oWorksheet.Cells(lngRow, lngNFCol).Value > 0 Then
                  dtNewDate = FormatDateTime(CDate(oWorksheet.Cells(lngRow, lngNFCol)), vbShortDate)
                  'determine actual or forecast
                  If dtNewDate <= dtStatus Then 'actual finish
                    If IsDate(oTask.ActualFinish) Then
                      If FormatDateTime(oTask.ActualFinish, vbShortDate) <> dtNewDate Then oTask.SetField lngAF, CDate(dtNewDate & " 05:00 PM")
                    Else
                      oTask.SetField lngAF, CDate(dtNewDate & " 05:00 PM")
                    End If
                    Print #lngFile, "UID " & lngUID & " AF > " & FormatDateTime(dtNewDate, vbShortDate)
                    If Not oDict.Exists(lngUID) Then oDict.Add lngUID, lngUID
                  ElseIf dtNewDate > dtStatus Then 'forecast finish
                    If FormatDateTime(oTask.Finish, vbShortDate) <> FormatDateTime(dtNewDate, vbShortDate) Then
                      If dtNewDate > #12/31/2149# Then
                        Print #lngFile, "UID " & lngUID & " FF > ERROR (" & FormatDateTime(dtNewDate, vbShortDate) & " is outside allowable range) " & String(10, "<")
                      Else
                        oTask.SetField lngFF, CDate(dtNewDate & " 05:00 PM")
                        Print #lngFile, "UID " & lngUID & " FF > " & FormatDateTime(dtNewDate, vbShortDate)
                        If Not oDict.Exists(lngUID) Then oDict.Add lngUID, lngUID
                      End If
                    End If
                  End If
                  If FormatDateTime(dtNewDate, vbShortDate) <> FormatDateTime(oTask.Finish, vbShortDate) Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, "FINISH", "", CStr(FormatDateTime(oTask.Finish, vbShortDate)), CStr(FormatDateTime(dtNewDate, vbShortDate))), ",")
                  End If
                End If
              End If
skip_nf:
              oWorksheet.Cells(lngRow, lngNFCol).NumberFormat = "m/d/yyyy"
            End If
            
            'evp
            'skip LOE
            If Len(strEVT) > 0 And Len(strLOE) > 0 Then
              lngEVT = CLng(Split(strEVT, "|")(0))
              If oTask.GetField(lngEVT) = strLOE Then GoTo skip_evp
            End If
            'secondary catch to skip LOE
            If oWorksheet.Cells(lngRow, lngEVPCol).Value <> "-" Then
              If oWorksheet.Cells(lngRow, lngEVPCol).DisplayFormat.Interior.Color = 13551615 Then 'invalid EV
                Print #lngFile, "UID " & lngUID & " - Invalid EV " & String(10, "<")
                cptAddRange rBad, oWorksheet.Cells(lngRow, lngUIDCol)
                blnValid = False
                GoTo skip_evp
              End If
              
              lngEVP = Round(oWorksheet.Cells(lngRow, lngEVPCol).Value * 100, 0)
              strEVP = cptGetSetting("Integration", "EVP")
              If Len(strEVP) > 0 Then 'compare
                If CLng(cptRegEx(oTask.GetField(Split(strEVP, "|")(0)), "[0-9]{1,}")) <> lngEVP Then
                  oTask.SetField lngEV, lngEVP
                  Print #lngFile, "UID " & lngUID & " EV% > " & lngEVP & "%"
                  If Not oDict.Exists(lngUID) Then oDict.Add lngUID, lngUID
                  Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, Split(strEVP, "|")(1), "", cptRegEx(oTask.GetField(Split(strEVP, "|")(0)), "[0-9]{1,}"), CStr(lngEVP)), ",")
                End If
              Else 'log
                oTask.SetField lngEV, lngEVP
                Print #lngFile, "UID " & lngUID & " EV% > " & lngEVP & "%"
                If Not oDict.Exists(lngUID) Then oDict.Add lngUID, lngUID
                Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, "EV%", "", "<unknown>", CStr(lngEVP)), ",")
              End If
            End If
            
skip_evp:
            'comments todo: only import if different (diff depends on vbCr and vbLf etc.)
            If .chkAppend And oWorksheet.Cells(lngRow, lngCommentsCol).Value <> "" Then
              If .cboAppendTo = "Top of Task Note" Then
                oTask.Notes = FormatDateTime(dtStatus, vbShortDate) & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf & String(25, "-") & vbCrLf & vbCrLf & oTask.Notes
              'todo: replace task note
              ElseIf .cboAppendTo = "Overwrite Note" Then
                oTask.Notes = FormatDateTime(dtStatus, vbShortDate) & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf
              ElseIf .cboAppendTo = "Bottom of Task Note" Then
                oTask.AppendNotes vbCrLf & String(25, "-") & vbCrLf & FormatDateTime(dtStatus, vbShortDate) & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf
              End If
            End If
            
          ElseIf Not blnTask Then 'it's an Assignment
            If oAssignment Is Nothing Then
              Print #lngFile, "MISSING: TASK UID: [" & lngUID & "] ASSIGNMENT UID: [" & oWorksheet.Cells(lngRow, lngUIDCol).Value & "] - " & oWorksheet.Cells(lngRow, lngTaskNameCol).Value
            Else
              Set oTask = oAssignment.Task
              If Not oWorksheet.Cells(lngRow, lngETCCol).Locked Then
                If oWorksheet.Cells(lngRow, lngETCCol).DisplayFormat.Interior.Color = 13551615 Then 'invalid ETC
                  Print #lngFile, "UID " & lngUID & " - Invalid ETC for " & oAssignment.ResourceName & " " & String(10, "<")
                  cptAddRange rBad, oWorksheet.Cells(lngRow, lngUIDCol)
                  cptAddRange rBad, oWorksheet.Cells(oWorksheet.Evaluate("MATCH(" & lngUID & ",A:A,0)"), lngUIDCol)
                  blnValid = False
                  GoTo next_row
                End If
                dblETC = oWorksheet.Cells(lngRow, lngETCCol).Value 'get the new value
                dblWas = 0 'reset was
                If oAssignment.ResourceType = pjResourceTypeWork Then
                  dblWas = Val(oAssignment.RemainingWork) / 60
                Else
                  dblWas = Val(oAssignment.RemainingCost)
                End If
                'only import if updated
                If Round(dblWas, 2) <> Round(dblETC, 2) Then
                  If lngETC = pjTaskNumber1 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number1 = dblETC
                    oTask.Number1 = oTask.Number1 + dblETC
                  ElseIf lngETC = pjTaskNumber2 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number2 = dblETC
                    oTask.Number2 = oTask.Number2 + dblETC
                  ElseIf lngETC = pjTaskNumber3 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number3 = dblETC
                    oTask.Number3 = oTask.Number3 + dblETC
                  ElseIf lngETC = pjTaskNumber4 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number4 = dblETC
                    oTask.Number4 = oTask.Number4 + dblETC
                  ElseIf lngETC = pjTaskNumber5 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number5 = dblETC
                    oTask.Number5 = oTask.Number5 + dblETC
                  ElseIf lngETC = pjTaskNumber6 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number6 = dblETC
                    oTask.Number6 = oTask.Number6 + dblETC
                  ElseIf lngETC = pjTaskNumber7 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number7 = dblETC
                    oTask.Number7 = oTask.Number7 + dblETC
                  ElseIf lngETC = pjTaskNumber8 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number8 = dblETC
                    oTask.Number8 = oTask.Number8 + dblETC
                  ElseIf lngETC = pjTaskNumber9 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number9 = dblETC
                    oTask.Number9 = oTask.Number9 + dblETC
                  ElseIf lngETC = pjTaskNumber10 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number10 = dblETC
                    oTask.Number10 = oTask.Number10 + dblETC
                  ElseIf lngETC = pjTaskNumber11 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number11 = dblETC
                    oTask.Number11 = oTask.Number11 + dblETC
                  ElseIf lngETC = pjTaskNumber12 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number12 = dblETC
                    oTask.Number12 = oTask.Number12 + dblETC
                  ElseIf lngETC = pjTaskNumber13 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number13 = dblETC
                    oTask.Number13 = oTask.Number13 + dblETC
                  ElseIf lngETC = pjTaskNumber14 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number14 = dblETC
                    oTask.Number14 = oTask.Number14 + dblETC
                  ElseIf lngETC = pjTaskNumber15 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number15 = dblETC
                    oTask.Number15 = oTask.Number15 + dblETC
                  ElseIf lngETC = pjTaskNumber16 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number16 = dblETC
                    oTask.Number16 = oTask.Number16 + dblETC
                  ElseIf lngETC = pjTaskNumber17 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number17 = dblETC
                    oTask.Number17 = oTask.Number17 + dblETC
                  ElseIf lngETC = pjTaskNumber18 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number18 = dblETC
                    oTask.Number18 = oTask.Number18 + dblETC
                  ElseIf lngETC = pjTaskNumber19 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number19 = dblETC
                    oTask.Number19 = oTask.Number19 + dblETC
                  ElseIf lngETC = pjTaskNumber20 Then
                    Print #lngDeconflictionFile, Join(Array(strFileName, lngUID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number20 = dblETC
                    oTask.Number20 = oTask.Number20 + dblETC
                  End If
                  Print #lngFile, "UID " & lngUID & " [" & oAssignment.Resource.Name & "] ETC > " & dblETC
                  If Not oDict.Exists(lngUID) Then oDict.Add lngUID, lngUID
                End If
                If .chkAppend And Len(oWorksheet.Cells(lngRow, lngCommentsCol)) > 0 Then
                  If .cboAppendTo = "Top of Task Note" Then
                    oAssignment.Notes = FormatDateTime(dtStatus, vbShortDate) & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf & String(25, "-") & vbCrLf & vbCrLf & oAssignment.Notes
                  'todo: replace assignment note
                  ElseIf .cboAppendTo = "Overwrite Note" Then
                    oAssignment.Notes = FormatDateTime(dtStatus, vbShortDate) & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf
                  ElseIf .cboAppendTo = "Bottom of Task Note" Then
                    oAssignment.AppendNotes vbCrLf & String(25, "-") & vbCrLf & FormatDateTime(dtStatus, vbShortDate) & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf
                  End If
                End If
              End If
              'todo: consolidate Assignment Notes into Task Notes?
              Set oAssignment = Nothing
            End If
          End If
next_row:
          myStatusSheetImport_frm.lblStatus.Caption = "Importing...(" & Format(lngRow / lngLastRow, "0%") & ")"
          myStatusSheetImport_frm.lblProgress.Width = (lngRow / lngLastRow) * myStatusSheetImport_frm.lblStatus.Width
          DoEvents
        Next lngRow
next_worksheet:
        If Not rBad Is Nothing Then rBad.Style = "Bad"
        If oBad.Count > 0 Then
          For lngBadItem = 0 To oBad.Count - 1
            oBad.Items(lngBadItem).Style = "Bad"
          Next lngBadItem
        End If
        Print #lngFile, String(25, "-")
      Next oWorksheet
next_file:

      If Not blnValid And blnKickoutReport Then
        Application.StatusBar = "Kickouts! Generating email..."
        'get outlook
        On Error Resume Next
        Set oOutlook = GetObject(, "Outlook.Application")
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If oOutlook Is Nothing Then
          Set oOutlook = CreateObject("Outlook.Application")
        End If
        'create email
        Set oMailItem = oOutlook.CreateItem(0) '0=olMailItem
        oMailItem.Display
        'add subject
        oMailItem.Subject = "ACTION REQUIRED: " & cptGetProgramAcronym & " - Invalid Status - " & Format(ActiveProject.StatusDate, "yyyy-mm-dd")
        If oMailItem.BodyFormat <> olFormatHTML Then oMailItem.BodyFormat = olFormatHTML
        'add some words
        On Error Resume Next
        Set oInspector = oMailItem.GetInspector
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If Not oInspector Is Nothing Then
          oInspector.WindowState = 1 '1=olMinimized
        End If
        Set oDocument = oMailItem.GetInspector.WordEditor
        Set oWord = oDocument.Application
        Set oSelection = oDocument.Windows(1).Selection
        oSelection.Text = "[NAME]: " & vbCrLf & vbCrLf & "Please correct the following invalid status entries and return to me ASAP:" & vbCrLf & vbCrLf
        For Each oWorksheet In oWorkbook.Sheets
          If oWorksheet.Name = "Conditional Formatting" Then GoTo next_worksheet1
          'validate status request worksheet
          Set oRange = Nothing
          On Error Resume Next
          Set oRange = oWorksheet.Range("STATUS_DATE")
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oRange Is Nothing Then GoTo next_worksheet1
          oWorksheet.Activate
          'show all rows and columns (if scheduler did not 'protect')
          If Not oWorksheet.AutoFilterMode Then
            oWorksheet.Rows.Hidden = False
            oWorksheet.Columns.Hidden = False
            oWorksheet.Cells(lngHeaderRow, lngUIDCol).Select
            oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, lngUIDCol), oWorksheet.Cells(lngLastRow, lngCommentsCol)).AutoFilter
          End If
          'unprotect sheet
          oWorksheet.UnProtect "NoTouching!" 'keep it secret; keep it safe
          'filter the list
          oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, lngUIDCol), oWorksheet.Cells(lngLastRow, lngCommentsCol)).AutoFilter Field:=lngUIDCol, Criteria1:=393372, Operator:=xlFilterFontColor
          'copy
          oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, lngUIDCol), oWorksheet.Cells(lngLastRow, lngCommentsCol)).SpecialCells(xlVisible).Copy
          'paste picture, resize it
          oSelection.MoveRight
          If oWorkbook.Sheets.Count > 0 Then
            oSelection.TypeText "Worksheet: " & oWorksheet.Name
            oSelection.MoveDown
          End If
          oSelection.Range.PasteAndFormat wdChartPicture
          oDocument.InlineShapes(1).LockAspectRatio = msoTrue
          oDocument.InlineShapes(1).Width = 1296
          're-protect sheet
          oWorksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
next_worksheet1:
        Next oWorksheet
        'save a copy
        strFileName = Replace(strFileName, ".xlsx", "_invalid.xlsx")
        If Dir(strFileName) <> vbNullString Then Kill strFileName
        oWorkbook.SaveCopyAs strFileName
        'attach it
        oMailItem.Attachments.Add strFileName
        'show it
        oInspector.WindowState = 2 'olNormalWindow
      End If
      
      .lblStatus.Caption = "Importing...(" & lngItem + 1 & " of " & .lboStatusSheets.ListCount & ")"
      .lblProgress.Width = ((lngItem + 1) / .lboStatusSheets.ListCount) * .lblStatus.Width
      .lboStatusSheets.Selected(lngItem) = False
      oWorkbook.Close False
      DoEvents
    Next lngItem
  End With 'myStatusSheetImport_frm
  
  'were there any conflicts?
  Application.StatusBar = "Checking for conflicting updates..."
  Close #lngDeconflictionFile
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("temp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  
  strSQL = "SELECT T1.TASK_UID,T1.RESOURCE_NAME,T1.FIELD,T2.WAS,T2.[IS],T2.FILE "
  strSQL = strSQL & "FROM ((SELECT TASK_UID,RESOURCE_NAME,FIELD,COUNT(FILE) FROM [imported.csv] GROUP BY TASK_UID,RESOURCE_NAME,FIELD HAVING COUNT(FILE)>1) AS T1) "
  strSQL = strSQL & "LEFT JOIN [imported.csv] AS T2 ON T2.TASK_UID=T1.TASK_UID AND T2.FIELD=T1.FIELD  " 'AND T2.RESOURCE_NAME=T1.RESOURCE_NAME
  strSQL = strSQL & "ORDER BY T1.TASK_UID,T1.FIELD" 'todo: refine query
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  If oRecordset.RecordCount > 0 Then
    Application.StatusBar = "Conflicts found! Generating report..."
    Print #lngFile, ">>> " & oRecordset.RecordCount & " POTENTIAL CONFLICTS IDENTIFIED <<<"
    If MsgBox("Potential conflicts found!" & vbCrLf & vbCrLf & "Review in Excel?", vbExclamation + vbYesNo, "Please Review") = vbYes Then
      Application.StatusBar = "Conflicts found! Generating report..."
      oExcel.Visible = True
      Set oWorkbook = oExcel.Workbooks.Add
      Set oWorksheet = oWorkbook.Sheets(1)
      For lngItem = 1 To oRecordset.Fields.Count
        oWorksheet.Cells(1, lngItem).Value = oRecordset.Fields(lngItem - 1).Name
      Next lngItem
      oWorksheet.[A2].CopyFromRecordset oRecordset
      oExcel.ActiveWindow.Zoom = 85
      oWorksheet.Columns.AutoFit
    Else
      Application.StatusBar = "Conflicts found! Skipping report..."
      For lngItem = 0 To oRecordset.Fields.Count - 1
        strHeader = strHeader & oRecordset.Fields(lngItem).Name & ","
      Next lngItem
      Print #lngFile, strHeader
      Print #lngFile, oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
      Print #lngFile, "...conflicts not reviewed."
    End If
  End If
  
  'get the string of UIDs updated
  If oDict.Count > 0 Then
    strUIDList = Join(oDict.Keys(), ",")
  End If
  
exit_here:
  On Error Resume Next
  Set oBad = Nothing
  Set oInspector = Nothing
  Set oOutlook = Nothing
  Set oMailItem = Nothing
  Set oDocument = Nothing
  Set oWord = Nothing
  Set oSelection = Nothing
  Set oEmailTemplate = Nothing
  Set oDict = Nothing
  Set oShell = Nothing
  If oRecordset.State = 1 Then oRecordset.Close
  Set oRecordset = Nothing
  Set oSubproject = Nothing
  myStatusSheetImport_frm.lblStatus.Caption = "Import Complete."
  myStatusSheetImport_frm.lblProgress.Width = myStatusSheetImport_frm.lblStatus.Width
  DoEvents
  'close log for output
  Print #lngFile, String(25, "=")
  Print #lngFile, "COMPLETE: " & FormatDateTime(Now, vbGeneralDate) & " [" & Format(Now - dtStart, "hh:nn:ss") & "]" & vbCrLf
  Print #lngFile, "UPDATED UIDs:"
  If Len(strUIDList) = 0 Then
    Print #lngFile, "< no updates >"
  Else
    Print #lngFile, strUIDList
  End If
  myStatusSheetImport_frm.lblStatus.Caption = "Ready..."
  Application.StatusBar = "Ready..."
  cptSpeed False
  Set oAssignment = Nothing
  Set oResource = Nothing
  Set oTask = Nothing
  Reset 'closes all active files opened by the Open statement and writes the contents of all file buffers to disk.
  If Dir(strImportLog) <> vbNullString And blnImportLog Then 'open log in notepad
    ShellExecute 0, "open", strImportLog, vbNullString, vbNullString, 1
  End If
  If Dir(Environ("tmp") & "\Schema.ini") <> vbNullString Then Kill Environ("tmp") & "\Schema.ini"
  If Dir(Environ("tmp") & "\imported.csv") <> vbNullString Then Kill Environ("tmp") & "\imported.csv"
  Set oRange = Nothing
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  'If Not oWorkbook Is Nothing Then oWorkbook.Close False
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oComboBox = Nothing
  If Not rst Is Nothing Then
    If rst.State = 1 Then rst.Close
  End If
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_bas", "cptStatusSheetImport", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshStatusImportTable(ByRef myStatusSheetImport_frm As cptStatusSheetImport_frm, Optional blnUsageBelow As Boolean = False)
  'objects
  Dim rst As Object 'ADODB.Recordset 'Object
  'strings
  Dim strDir As String
  Dim strBottomPaneViewName As String
  Dim strEVT As String
  Dim strEVP As String
  Dim strSettings As String
  'longs
  Dim lngEVT As Long
  Dim lngETC As Long
  Dim lngEVP As Long
  Dim lngNewEVP As Long
  Dim lngFF As Long
  Dim lngAF As Long
  Dim lngFS As Long
  Dim lngAS As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  If Not myStatusSheetImport_frm.Visible Then GoTo exit_here

  cptSpeed True
  
  'get saved settings
  'get EVP and EVT
  strSettings = strDir & "\settings\cpt-status-sheet.adtg" 'todo: keep for a few more versions
  If Dir(strSettings) <> vbNullString Then
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open strSettings
    If Not rst.EOF Then
      'todo: does field name still match?
      strEVP = rst("cboEVP")
      strEVT = rst("cboEVT")
    End If
    rst.Close
    'convert to ini
    cptSaveSetting "StatusSheet", "cboEVP", strEVP
    cptSaveSetting "StatusSheet", "cboEVT", strEVT
    'todo: don't kill the file here, kill it on Status Sheet Creation
  End If
  
  strEVP = cptGetSetting("Integration", "EVP")
  strEVT = cptGetSetting("Integration", "EVT")
  
  'reset the table
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=10, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Unique ID", Title:="UID", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  
  'import user fields
  strSettings = strDir & "\settings\cpt-status-sheet-userfields.adtg"
  If Dir(strSettings) <> vbNullString Then
    'import user settings
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open strSettings
    If Not rst.EOF Then
      rst.MoveFirst
      Do While Not rst.EOF
        'does field name still match?
        If CustomFieldGetName(rst(0)) = rst(1) Or FieldConstantToFieldName(rst(0)) = rst(1) Then
          TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=rst(1), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
        Else
          If CustomFieldGetName(rst(0)) = "" Then
            MsgBox "Saved field '" & rst(1) & "' has been renamed to '" & FieldConstantToFieldName(rst(0)) & "' - you may want to remove it from your list.", vbInformation + vbOKOnly, "Saved Field Changed"
          Else
            MsgBox "Saved field '" & rst(1) & "' has been renamed to '" & CustomFieldGetName(rst(0)) & "' - you may want to remove it from your list.", vbInformation + vbOKOnly, "Saved Field Changed"
          End If
        End If
        rst.MoveNext
      Loop
    End If
    rst.Close
  End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Name", Title:="", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Total Slack", Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Actual Start", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If Not IsNull(myStatusSheetImport_frm.cboAS.Value) Then
    lngAS = myStatusSheetImport_frm.cboAS.Value
    If lngAS <> FieldNameToFieldConstant("Actual Start") Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngAS), Title:="New Actual Start", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Start", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If Not IsNull(myStatusSheetImport_frm.cboFS.Value) Then
    lngFS = myStatusSheetImport_frm.cboFS.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngFS), Title:="New Forecast Start", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Actual Finish", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If Not IsNull(myStatusSheetImport_frm.cboAF.Value) Then
    lngAF = myStatusSheetImport_frm.cboAF.Value
    If lngAF <> FieldNameToFieldConstant("Actual Finish") Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngAF), Title:="New Actual Finish", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Remaining Duration", Title:="", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Finish", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  If Not IsNull(myStatusSheetImport_frm.cboFF.Value) Then
    lngFF = myStatusSheetImport_frm.cboFF.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngFF), Title:="New Forecast Finish", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  'EVT (WP)
  If Len(strEVT) > 0 Then
    On Error Resume Next
    lngEVT = Split(strEVT, "|")(0) 'FieldNameToFieldConstant(strEVT)
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If lngEVT > 0 Then
      If cptCustomFieldExists(CStr(Split(strEVT, "|")(1))) Then
        TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=Split(strEVT, "|")(1), Title:="EVT", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
      End If
    End If
  End If




  'existing EV%
  If Len(strEVP) > 0 Then
    'does field still exist?
    On Error Resume Next
    lngEVP = Split(strEVP, "|")(0) 'FieldNameToFieldConstant (strEVP)
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If lngEVP > 0 Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=Split(strEVP, "|")(1), Title:="EV%", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  End If
  'imported EV
  If Not IsNull(myStatusSheetImport_frm.cboEV.Value) Then
    lngNewEVP = myStatusSheetImport_frm.cboEV.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngNewEVP), Title:="New EV%", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  'keep these here so user can filter on changes above, make edits below
  'Type
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Type", Width:=17, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  'Effort Driven
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Effort Driven", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  'existing ETC (remaining work)
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Remaining Work", Title:="ETC", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  'imported ETC
  If Not IsNull(myStatusSheetImport_frm.cboETC.Value) Then
    lngETC = myStatusSheetImport_frm.cboETC.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngETC), Title:="New ETC", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  
  If blnUsageBelow Then
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=10, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Unique ID", Title:="UID", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    
    'import user fields
    strSettings = strDir & "\settings\cpt-status-sheet-userfields.adtg"
    If Dir(strSettings) <> vbNullString Then
      'import user settings
      Set rst = CreateObject("ADODB.Recordset")
      rst.Open strSettings
      If Not rst.EOF Then
        rst.MoveFirst
        Do While Not rst.EOF
          'does field name still match?
          If CustomFieldGetName(rst(0)) = rst(1) Then
            TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:=rst(1), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
          End If
          rst.MoveNext
        Loop
      End If
      rst.Close
    End If
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Name", Title:="", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    'Type
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Type", Width:=17, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    'Effort Driven
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Effort Driven", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    'existing ETC (remaining work)
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Remaining Work", Title:="ETC", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    'imported ETC
    If Not IsNull(myStatusSheetImport_frm.cboETC.Value) Then
      lngETC = myStatusSheetImport_frm.cboETC.Value
      TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngETC), Title:="New ETC", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Notes", Title:="", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    ActiveWindow.TopPane.Activate
    'If ActiveProject.CurrentView <> "Gantt Chart" Then ViewApply Name:="Gantt Chart"
    ViewApply Name:="Gantt Chart"
    'If ActiveProject.CurrentTable <> "cptStatusSheetImport Table" Then TableApply Name:="cptStatusSheetImport Table"
    TableApply Name:="cptStatusSheetImport Table"
    SetSplitBar ShowColumns:=ActiveProject.TaskTables(ActiveProject.CurrentTable).TableFields.Count
    'todo: reapply group?
    
    On Error Resume Next
    strBottomPaneViewName = ActiveWindow.BottomPane.View.Name
    If Err.Number = 91 Then
      Err.Clear
      Application.FormViewShow
      'Application.ToggleTaskDetails
    End If
    ActiveWindow.BottomPane.Activate
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    'If ActiveProject.CurrentView <> "Task Usage" Then ViewApply "Task Usage"
    ViewApply "Task Usage"
    'If ActiveProject.CurrentTable <> "cptStatusSheetImportDetails Table" Then TableApply "cptStatusSheetImportDetails Table"
    TableApply "cptStatusSheetImportDetails Table"
    ActiveWindow.TopPane.Activate
    SetSplitBar ShowColumns:=ActiveProject.TaskTables(ActiveProject.CurrentTable).TableFields.Count
  Else
    ActiveWindow.TopPane.Activate
    'If ActiveProject.CurrentView <> "Task Usage" Then ViewApply "Task Usage"
    ViewApply "Task Usage"
    DoEvents
    'If ActiveProject.CurrentTable <> "cptStatusSheetImport Table" Then TableApply Name:="cptStatusSheetImport Table"
    TableApply Name:="cptStatusSheetImport Table"
    SetSplitBar ShowColumns:=ActiveProject.TaskTables(ActiveProject.CurrentTable).TableFields.Count
    On Error Resume Next
    strBottomPaneViewName = ActiveWindow.BottomPane.View.Name
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Len(strBottomPaneViewName) > 0 Then
      DetailsPaneToggle
    End If
    'todo: reapply group?
  End If
  
  'reset the filter
'  FilterEdit Name:="cptStatusSheetImport Filter", Taskfilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Actual Finish", test:="equals", Value:="NA", ShowInMenu:=False, ShowSummaryTasks:=True
'  If myStatusSheetImport_frm.chkHide And IsDate(myStatusSheetImport_frm.txtHideCompleteBefore) Then
'    FilterEdit Name:="cptStatusSheetImport Filter", Taskfilter:=True, FieldName:="", newfieldname:="Actual Finish", test:="is greater than or equal to", Value:=myStatusSheetImport_frm.txtHideCompleteBefore, operation:="Or", ShowSummaryTasks:=True
'  End If
'  FilterApply "cptStatusSheetImport Filter"

exit_here:
  On Error Resume Next
  If rst.State = 1 Then rst.Close
  Set rst = Nothing
  cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_bas", "cptRefreshStatusImportTable", Err, Erl)
  Err.Clear
  Resume exit_here
End Sub

Sub cptAddRange(ByRef rBad As Excel.Range, ByRef oRange As Excel.Range)
  If rBad Is Nothing Then
    Set rBad = oRange
  Else
    Set rBad = oRange.Application.Union(rBad, oRange)
  End If
  If Len(rBad.Address) > 200 Then
    oBad.Add rBad, rBad
    Set rBad = Nothing
  End If
End Sub
