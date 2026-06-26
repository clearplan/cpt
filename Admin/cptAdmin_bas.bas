Attribute VB_Name = "cptAdmin_bas"
'>no cpt version - not for release<
Option Explicit

Function cptBuildAdminRibbon() As String
  Dim ribbonXML As String
  
  ribbonXML = ribbonXML + vbCrLf & "<mso:tab id=""tCPTAdmin"" label=""CPT ADMIN"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCPTAdmin"" label=""Admin"" visible=""true"">"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLoadFromPath"" label=""Create Release Asset"" imageMso=""RefreshWebView"" onAction=""cptLoadModulesFromPath"" size=""large"" supertip=""Create release asset; save it to /releases; and load modules from repo branch."" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:tab>"
  cptBuildAdminRibbon = ribbonXML
  
End Function

Sub cptDocument()
  'objects
  Dim vbComponent As vbComponent
  Dim oExcel As Object
  Dim oWorkbook As Object
  Dim oWorksheet As Object
  'strings
  Dim strModule As String
  Dim strProcName As String
  'longs
  Dim lngSLOC As Long
  Dim lngLines As Long
  Dim lngLine As Long
  Dim lngRow As Long
  Dim lngCountDecl As Long
  'integers
  'booleans
  'variants
  Dim arrHeader As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'get excel
  Set oExcel = CreateObject("Excel.Application")
  oExcel.Visible = True
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)

  oExcel.ActiveWindow.Zoom = 85
  oWorksheet.[A2].Select
  oExcel.ActiveWindow.FreezePanes = True

  'set the header
  arrHeader = Array("Ribbon Group", "Module", "SLOC", "Procedure", "SLOC", "Directory", "HelpDoc", "Author")
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].Offset(0, UBound(arrHeader))) = arrHeader
  oWorksheet.Columns.AutoFit

  lngRow = 2

  For Each vbComponent In ThisProject.VBProject.VBComponents
    strModule = vbComponent.Name
    Debug.Print "working on " & strModule & "..."
    If strModule = "ThisProject" Or Left(strModule, 3) = "cpt" Then
      With vbComponent.CodeModule
        lngCountDecl = .CountOfDeclarationLines
        lngLines = .CountOfLines
        oWorksheet.Cells(lngRow, 2) = .Name
        oWorksheet.Cells(lngRow, 3) = .CountOfLines
        strProcName = .ProcOfLine(lngCountDecl + 1, 0) '0 = vbext_pk_Proc
        oWorksheet.Cells(lngRow, 4) = strProcName
        oWorksheet.Cells(lngRow, 5) = .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
        lngSLOC = lngSLOC + .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
        oWorksheet.Columns.AutoFit
        For lngLine = lngCountDecl + 1 To lngLines
          If .ProcOfLine(lngLine, 0) <> strProcName Then '0 = vbext_pk_Proc
            strProcName = .ProcOfLine(lngLine, 0) '0 = vbext_pk_Proc
            lngRow = lngRow + 1
            oWorksheet.Cells(lngRow, 2) = strModule
            oWorksheet.Cells(lngRow, 4) = strProcName
            oWorksheet.Cells(lngRow, 5) = .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
            lngSLOC = lngSLOC + .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
            oWorksheet.Columns.AutoFit
            If lngRow > 10 Then oExcel.ActiveWindow.ScrollRow = lngRow - 10
          End If
        Next
      End With
      lngRow = lngRow + 1
      If lngRow > 10 Then oExcel.ActiveWindow.ScrollRow = lngRow - 10
    End If
  Next vbComponent

  oExcel.ActiveWindow.ScrollRow = 2

  MsgBox "Documented." & vbCrLf & vbCrLf & "(" & Format(lngSLOC, "#,##0") & " SLOC)", vbInformation + vbOKOnly, "Documenter"

exit_here:
  On Error Resume Next
  Set vbComponent = Nothing
  Set oExcel = Nothing
  Set oWorkbook = Nothing
  Set oWorksheet = Nothing
  Set oExcel = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptAdmin_bas", "Document", Err)
  Resume exit_here
End Sub

Sub cptCheckAllVersions()
Dim vbComponent As vbComponent

  For Each vbComponent In ThisProject.VBProject.VBComponents
    If Left(vbComponent.Name, 3) = "cpt" Then
      Debug.Print vbComponent.Name & ": " & Replace(Replace(cptRegEx(vbComponent.CodeModule.Lines(1, 10), "<cpt_version>.*</cpt_version>"), "<cpt_version>", ""), "</cpt_version>", "")
    End If
  Next vbComponent
  Set vbComponent = Nothing

End Sub

Function cptSetDirectory(strComponentName As String) As String
'strings
Dim strDirectory As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strDirectory = cptRegEx(strComponentName, "[^(cpt)](.*)(?=(_frm|_bas|_cls))", , False)
  
  Select Case strDirectory
    Case "About"
      strDirectory = "Core"
    Case "Adjustment"
      strDirectory = "Integration"
    Case "AdvancedFilter"
      strDirectory = "Text"
    Case "AdvancedFilterEdit"
      strDirectory = "Text"
    Case "AgeDates"
      strDirectory = "Status"
    Case "BrowseFolder"
      strDirectory = "Core"
    Case "BulkLogic"
      strDirectory = "Trace"
    Case "CalendarExceptions"
      strDirectory = "Calendar"
    Case "CountTasks"
      strDirectory = "Count"
    Case "CriticalPath"
      strDirectory = "Trace"
    Case "CriticalPathTools"
      strDirectory = "Trace"
    Case "CritPathFields"
      strDirectory = "Trace"
    Case "CustomFieldUsage"
      strDirectory = "CustomFields"
    Case "CheckAssignments"
      strDirectory = "Integration"
    Case "CommonFieldMap"
      strDirectory = "Core"
    Case "DataDictionary"
      strDirectory = "CustomFields"
    Case "DECM"
      strDirectory = "Metrics"
    Case "DECMTargetUID"
      strDirectory = "Metrics"
    Case "DynamicFilter"
      strDirectory = "Text"
    Case "Events"
      strDirectory = "Core"
    Case "ExIm"
      strDirectory = "Text"
    Case "FieldBuilder"
      strDirectory = "CustomFields"
    Case "FilterByClipboard"
      strDirectory = "Text"
    Case "FilterItem"
      strDirectory = "Text"
    Case "Fiscal"
      strDirectory = "Calendar"
    Case "FlowDown"
      strDirectory = "CustomFields"
    Case "Graphics"
      strDirectory = "Metrics"
    Case "IMSCobraExport"
      strDirectory = "Integration"
    Case "IPMDAR"
      strDirectory = "Status"
    Case "IPMDARMapping"
      strDirectory = "Status"
    Case "ListBox"
      strDirectory = "Core"
    Case "MetricsData"
      strDirectory = "Metrics"
    Case "MetricsSettings"
      strDirectory = "Metrics"
    Case "NetworkBrowser"
      strDirectory = "Trace"
    Case "Patch"
      strDirectory = ""
    Case "QBD"
      strDirectory = "Status"
    Case "ResetAll"
      strDirectory = "Core"
    Case "SaveLocal"
      strDirectory = "CustomFields"
    Case "SaveMarked"
      strDirectory = "Trace"
    Case "Settings"
      strDirectory = "Core"
    Case "Setup"
      strDirectory = ""
    Case "SmartDuration"
      strDirectory = "Status"
    Case "StatusSheet"
      strDirectory = "Status"
    Case "StatusSheetImport"
      strDirectory = "Status"
    Case "TaskHistory"
      strDirectory = "Status"
    Case "TaskTypeMapping"
      strDirectory = "Status"
    Case "ThisProject"
      strDirectory = "Core"
    Case "Upgrades"
      strDirectory = "Core"
    Case Else
      'use module name as directory

  End Select

  cptSetDirectory = strDirectory & "\"

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptAdmin_bas", "cptSetDirectory()", Err)
  Resume exit_here

End Function

Sub cptSQL(strFileName As String, Optional blnExport As Boolean = False)
  'objects
  Dim oListObject As Excel.ListObject
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strRecord As String
  Dim strFields As String
  Dim strCon As String, strDir As String, strSQL As String
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  'cpt-data-dictionary.adtg
  'cpt-export-resource-userfields.adtg
  'cpt-qbd.adtg
  'cpt-status-sheet.adtg
  'cpt-status-sheet-userfields.adtg
  'git-vba-repo.adtg
  'vba-backup-modules.adtg

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = cptDir & strFileName

  If Dir(strFileName) = vbNullString Then
    Debug.Print "Invalid file: " & strFileName
    GoTo exit_here
  End If
  
  If blnExport Then
    Set oExcel = CreateObject("Excel.Application")
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
  End If
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strFileName
    If .EOF Then
      MsgBox "No records.", vbInformation + vbOKOnly, "cptSQL"
      .Close
      GoTo exit_here
    Else
      blnExport = MsgBox(Format(.RecordCount, "#,##0") & " record(s). Export to Excel?", vbQuestion + vbYesNo, "cptSQL") = vbYes
    End If
    'get field names
    For lngField = 0 To .Fields.Count - 1
      If blnExport Then
        oWorksheet.Cells(1, lngField + 1).Value = .Fields(lngField).Name
      Else
        strFields = strFields & .Fields(lngField).Name & " | "
      End If
    Next lngField
    If Not blnExport Then Debug.Print strFields
    'get records
    If Not .EOF Then .MoveFirst
    If blnExport Then
      oWorksheet.[A2].CopyFromRecordset oRecordset
    Else
      Do While Not .EOF
        strRecord = ""
        For lngField = 0 To .Fields.Count - 1
          strRecord = strRecord & .Fields(lngField) & " | "
        Next lngField
        Debug.Print strRecord
        .MoveNext
      Loop
      .Close
    End If
  End With

  If blnExport Then
    oExcel.Visible = True
    oExcel.WindowState = xlMaximized
    With oExcel.ActiveWindow
      .Zoom = 85
      .SplitRow = 1
      .SplitColumn = 0
      .FreezePanes = True
    End With
    Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)), , xlYes)
    oListObject.TableStyle = ""
    oListObject.HeaderRowRange.Font.Bold = True
    oWorksheet.Columns.AutoFit
  End If

exit_here:
  On Error Resume Next
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptAdmin_bas", "cptSQL", Err, Erl)
  Resume exit_here
End Sub

Sub cptLoadModulesFromPath()
  'objects
  Dim oSubFolder As Object
  Dim oFSO As Scripting.FileSystemObject
  Dim oFolder As Scripting.Folder
  Dim oFile As Scripting.File
  Dim oVBProject As VBProject
  'strings
  Dim strVersion As String
  Dim strBranch As String
  Dim strDir As String
  'longs
  Dim lngProject As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vResponse As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  vResponse = InputBox("version:", "Create Release Asset")
  If StrPtr(vResponse) = 0 Then 'user hit cancel
    GoTo exit_here
  ElseIf Len(vResponse) = 0 Then 'no entry
    GoTo exit_here
  ElseIf Len(vResponse) > 0 Then
    strVersion = CStr(vResponse)
  End If
  
  Application.FileNew
  strDir = Environ("USERPROFILE") & "\GitHub\cpt"
  Application.FileSaveAs strDir & "\releases\cpt_" & strVersion & ".mpp"
  Set oVBProject = ActiveProject.VBProject

  If oVBProject Is Nothing Then GoTo exit_here
  
  strBranch = Replace(GitCommand("c:/users/arongahagan/GitHub/cpt", "rev-parse --abbrev-ref HEAD"), Chr(10), "")
  If MsgBox("Load modules from branch '" & strBranch & "' into cpt_" & strVersion & ".mpp?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
    FileClose pjDoNotSave
    Kill strDir & "\releases\cpt_" & strVersion & ".mpp"
  End If
  
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFolder = oFSO.GetFolder(strDir)
  For Each oFile In oFolder.Files
    If Len(cptRegEx(oFile.Name, "bas$|frm$|cls$")) > 0 Then
      Application.StatusBar = "Importing " & oFile.Name & "..."
      oVBProject.VBComponents.Import oFile.Path
    End If
  Next oFile
  For Each oSubFolder In oFolder.SubFolders
    If oSubFolder.Name = "Admin" Then GoTo next_subfolder
    For Each oFile In oSubFolder.Files
      If Len(cptRegEx(oFile.Name, "bas$|frm$|cls$")) > 0 Then
        Application.StatusBar = "Importing " & oFile.Name & "..."
        oVBProject.VBComponents.Import oFile.Path
      End If
    Next oFile
next_subfolder:
  Next oSubFolder
  
  Application.StatusBar = "Complete."
  
  MsgBox "Run cptSetReferences in newly created file; and" & vbCrLf & vbCrLf & "...compile it!", vbExclamation + vbOKOnly, "Don't Forget:"
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oSubFolder = Nothing
  Set oFolder = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing
  Set oVBProject = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptAdmin_bas", "cptLoadModulesFromPath", Err, Erl)
  Resume exit_here
End Sub

Function cptGetAllSettings(strSection)
  Dim vSettings As Variant
  Dim intSetting As Integer
  vSettings = GetAllSettings("ClearPlanToolbar", strSection)
  For intSetting = LBound(vSettings, 1) To UBound(vSettings, 1)
    Debug.Print vSettings(intSetting, 0) & "=" & vSettings(intSetting, 1)
  Next
End Function

Function cptGetLongestLine() As Long
  Dim vbComponent As vbComponent, lngLine As Long, lngMax As Long, lngLineLength As Long
  For Each vbComponent In ThisProject.VBProject.VBComponents
    For lngLine = 1 To vbComponent.CodeModule.CountOfLines
      lngLineLength = Len(vbComponent.CodeModule.Lines(lngLine, 1))
      If lngLineLength > lngMax Then
        lngMax = lngLineLength
      End If
    Next lngLine
  Next vbComponent
  cptGetLongestLine = lngMax
  Set vbComponent = Nothing
End Function

Sub ImportKEODataToCPT()
  'objects
  Dim oCPT As ADODB.Recordset
  Dim oKEO As ADODB.Recordset
  'strings
  Dim strCPTFile As String
  Dim strKEOFile As String
  'longs
  Dim lngItem As Long
  Dim lngImported As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Set oKEO = CreateObject("ADODB.Recordset")
  Set oCPT = CreateObject("ADODB.Recordset")

  strKEOFile = Environ("USERPROFILE") & "\OneDrive - ClearPlan LLC\Clients\L3H\KEO\BrightLights\Metrics\keo-tasks.adtg"
  strCPTFile = cptDir & "\settings\cpt-cei.adtg"
  
  oKEO.Open strKEOFile
  oCPT.Open strCPTFile
  
  oKEO.MoveFirst
  oCPT.MoveFirst
  
  Do While Not oKEO.EOF
    oCPT.Filter = "PROJECT='" & oKEO("PROJECT") & "' AND TASK_UID=" & oKEO("TASK_UID") & " AND STATUS_DATE=#" & FormatDateTime(oKEO("STATUS_DATE"), vbGeneralDate) & "#"
    If oCPT.EOF Then 'import it
      oCPT.AddNew
      For lngItem = 0 To oKEO.Fields.Count - 1
        oCPT(oKEO.Fields(lngItem).Name) = oKEO(lngItem)
      Next lngItem
      oCPT.Update
      lngImported = lngImported + 1
    End If
    oCPT.Filter = 0
    Debug.Print oKEO.AbsolutePosition & " / " & oKEO.RecordCount & "...(" & Format(oKEO.AbsolutePosition / oKEO.RecordCount, "0%") & ")"
    oKEO.MoveNext
  Loop
  
  Debug.Print Format(lngImported, "#,##0") & " records imported."
  
  oKEO.Close
  oCPT.Save strCPTFile, adPersistADTG
  oCPT.Close
  
exit_here:
  On Error Resume Next
  Set oCPT = Nothing
  Set oKEO = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAdmin_bas", "ImportKEODataToCPT", Err, Erl)
  Resume exit_here
End Sub

Sub LoadMilStd881()
  'copy/paste from latest Mil-Std-881 into vim/csv and clean up
  'objects
  Dim oRecordsOut As ADODB.Recordset
  Dim oRecordsIn As ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strFileName As String
  Dim strCon As String
  Dim strSQL As String
  'longs
  Dim lngField As Long
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'create schema for the csv
  strDir = Environ("USERPROFILE") & "\GitHub\cpt"
  strFileName = strDir & "\Schema.ini"
  lngFile = FreeFile
  Open strFileName For Output As #lngFile
  Print #lngFile, "[mil-std-881.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=VERSION text width 1"
  Print #lngFile, "Col2=APPENDIX text width 4"
  Print #lngFile, "Col3=CODE text width 16"
  Print #lngFile, "Col4=DESCRIPTION text width 255"
  Close #lngFile
  
  'create a connection string
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  
  'create the sql query
  strSQL = "SELECT * FROM [mil-std-881.csv]"
  
  'create recordset: csv (in)
  Set oRecordsIn = CreateObject("ADODB.Recordset")
  'open recordset: csv (in)
  oRecordsIn.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  
  'create recordset: adtg (out)
  Set oRecordsOut = CreateObject("ADODB.Recordset")
  For lngField = 0 To oRecordsIn.Fields.Count - 1
    oRecordsOut.Fields.Append oRecordsIn.Fields(lngField).Name, oRecordsIn.Fields(lngField).Type, oRecordsIn.Fields(lngField).DefinedSize
  Next lngField
  'open recordset: adtg (out)
  oRecordsOut.Open
  
  'load csv->adtg
  Do While Not oRecordsIn.EOF
    oRecordsOut.AddNew Array(0, 1, 2, 3), Array(oRecordsIn(0), oRecordsIn(1), oRecordsIn(2), Replace(oRecordsIn(3), "…", "..."))
    oRecordsIn.MoveNext
  Loop
  
  'save the adtg
  strFileName = cptDir & "\cpt-mil-std-881.adtg"
  'provide user feedback
  Debug.Print Format(oRecordsOut.RecordCount, "#,##0") & " records imported."
  'overwrite file if it exists
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  oRecordsOut.Save strFileName, adPersistADTG

exit_here:
  On Error Resume Next
  'kill the Schema.ini
  If Not Dir(strDir & "\Schema.ini") = vbNullString Then
    Kill strDir & "\Schema.ini"
  End If
  'close the recordsets
  If oRecordsOut.State Then oRecordsOut.Close
  If oRecordsIn.State Then oRecordsIn.Close
  Set oRecordsOut = Nothing
  Set oRecordsIn = Nothing
  Reset 'close any open text files

  Exit Sub
err_here:
  Call cptHandleErr("foo", "bar", Err, Erl)
  Resume exit_here
End Sub

Sub UpdateVersionsFromFile()
  'run ./get-current-versions.sh >> CurrentVersions.txt
  'run this
  Dim oFSO As Scripting.FileSystemObject
  Dim oStream As Scripting.TextStream
  Dim oCodeModule As VBIDE.CodeModule
  Dim strFileName As String
  Dim strLine As String
  Dim strModule As String
  Dim strVersion As String
  Dim strVersionWas As String
  Dim lngLine As Long
  Dim lngResponse As Long
  Dim lngUpdated As Long
  Dim lngSkipped As Long
  
  On Error GoTo 0
  
  'get new versions from branch
  Set oFSO = CreateObject("Scripting.FileSystemOBject")
  strFileName = Environ("userprofile") & "\GitHub\cpt\CurrentVersions.txt"
  Set oStream = oFSO.OpenTextFile(strFileName, ForReading)
  Do While Not oStream.AtEndOfStream
    strLine = oStream.ReadLine
    strModule = Replace(cptRxMatch(Split(strLine, ",")(0), "[A-z_]+\.", True, False), ".", "")
    strVersion = Split(strLine, ",")(1)
    Set oCodeModule = Nothing
    On Error Resume Next
    Set oCodeModule = ThisProject.VBProject.VBComponents(strModule).CodeModule
    If oCodeModule Is Nothing Then
      Debug.Print strModule & ": not found! <<<<<<<<<<"
      lngSkipped = lngSkipped + 1
      GoTo next_module
    End If
    If Not oCodeModule.Find("<cpt_version>" & strVersion & "</cpt_version>", 1, 1, 50, 80, True, True, False) Then
      For lngLine = 1 To oCodeModule.CountOfLines
        If oCodeModule.Find("</cpt_version>", lngLine, 1, lngLine, 80, False, True, False) Then
          strVersionWas = cptRxMatch(oCodeModule.Lines(lngLine, 1), "<cpt_version>.*</cpt_version>")
          strVersionWas = Replace(Replace(strVersionWas, "<cpt_version>", ""), "</cpt_version>", "")
          Debug.Print oCodeModule.Lines(lngLine, 1)
          lngResponse = MsgBox(strModule & ":" & vbCrLf & strVersionWas & " > " & strVersion & "?" & vbCrLf & vbCrLf & "Replace?", vbQuestion + vbYesNoCancel, "Please Confirm")
          If lngResponse = vbYes Then
            oCodeModule.ReplaceLine lngLine, "'<cpt_version>" & strVersion & "</cpt_version>"
            lngUpdated = lngUpdated + 1
            Debug.Print strModule & ": " & strVersionWas & " > " & strVersion
            Exit For
          ElseIf lngResponse = vbNo Then
            Debug.Print strModule & ": skipped"
            lngSkipped = lngSkipped + 1
            Exit For
          ElseIf lngResponse = vbCancel Then
            Debug.Print "...process terminated."
            Exit Do
          End If
        End If
      Next lngLine
    End If
next_module:
  Loop
  Debug.Print "updated: " & lngUpdated & vbCrLf & "skipped: " & lngSkipped
  oStream.Close
  Set oCodeModule = Nothing
  Set oStream = Nothing
  Set oFSO = Nothing
End Sub

