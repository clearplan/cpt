Attribute VB_Name = "cptBackbone_bas"
'<cpt_version>v1.4.1</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptBackbone_bas"

Sub cptImportCWBSFromExcel(ByRef myBackbone_frm As cptBackbone_frm, lngOutlineCode As Long)
  'objects
  Dim oValid As Scripting.Dictionary
  Dim oInvalid As Scripting.Dictionary
  Dim oTask As MSProject.Task
  Dim oLookupTable As MSProject.LookupTable
  Dim oOutlineCode As MSProject.OutlineCode
  Dim c As Excel.Range
  Dim oRange As Excel.Range
  Dim oFileDialog As Office.FileDialog
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  'strings
  Dim strMsg As String
  Dim strOutlineCode As String
  Dim strValue As String
  Dim strDescription As String
  'longs
  Dim lngItems As Long
  Dim lngOutlineLevel As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnValid As Boolean
  'variants
  'dates
    
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    
  If MsgBox("Expected fields/column headers, in range [A1:C1], are CODE,LEVEL,DESCRIPTION and there should be no blank rows. CODE should be unique and sorted properly." & vbCrLf & vbCrLf & "Proceed?", vbQuestion + vbYesNo, "Confirm CWBS Import") = vbNo Then
    'export a sample template
    If MsgBox("Would you like an example?", vbQuestion + vbYesNo, "A little help") = vbYes Then Call cptExportTemplate
  Else
    strOutlineCode = CustomFieldGetName(lngOutlineCode)
    Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
    
    Set oExcel = CreateObject("Excel.Application")
    'allow user to select excel file and import it to chosen
    Set oFileDialog = oExcel.FileDialog(msoFileDialogFilePicker)
    With oFileDialog
      .AllowMultiSelect = False
      .ButtonName = "Import"
      .InitialView = msoFileDialogViewDetails
      .InitialFileName = Environ("USERPROFILE") & "\"
      .Title = "Select " & strOutlineCode & " source file:"
      .Filters.Add "Microsoft Excel Workbook (xlsx)", "*.xlsx"
      .Filters.Add "Comma Separated Values (csv)", "*.csv"
      If .Show = -1 Then
      
        Application.OpenUndoTransaction "Import " & strOutlineCode & " from Excel Workbook"
      
        cptSpeed True
                
        'open the workbook
        Set oWorkbook = oExcel.Workbooks.Open(oFileDialog.SelectedItems(1))
        'find the sheet
        For Each oWorksheet In oWorkbook.Sheets
          If UCase(Trim(oWorksheet.[A1].Value)) = "CODE" And UCase(Trim(oWorksheet.[B1].Value)) = "LEVEL" And UCase(Trim(oWorksheet.[C1].Value)) = "DESCRIPTION" Then
            strOutlineCode = CustomFieldGetName(lngOutlineCode)
            'build the code mask
            lngOutlineLevel = oWorksheet.Evaluate("MAX(B:B)")
            For lngItem = 1 To lngOutlineLevel
              CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=lngItem, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
            Next lngItem
            CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
            
            Set oRange = oWorksheet.Range(oWorksheet.[A2], oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(-4162)) '-4162 = xlUp
            'validate unique Codes (esp. 1.1 and 1.10 when excel hacks off trailing zeros
            Set oValid = CreateObject("Scripting.Dictionary")
            Set oInvalid = CreateObject("Scripting.Dictionary")
            For Each c In oRange.Cells
              If Not oValid.Exists(c.Value) Then
                oValid.Add c.Value, c.Offset(0, 2).Value
              Else
                oInvalid.Add c.Value, c.Offset(0, 2).Value
              End If
            Next c
            If oInvalid.Count > 0 Then
              blnValid = False
              strMsg = "Duplicate Codes found!" & vbCrLf
              For lngItem = 0 To oInvalid.Count - 1
                strMsg = strMsg & "- " & oInvalid.Keys(lngItem) & vbCrLf
              Next lngItem
              strMsg = strMsg & vbCrLf & "Code must be unique."
              'indicate duplicate values
              oRange.FormatConditions.AddUniqueValues
              oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
              oRange.FormatConditions(1).DupeUnique = xlDuplicate
              With oRange.FormatConditions(1).Font
                .Color = -16383844
                .TintAndShade = 0
              End With
              With oRange.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615
                .TintAndShade = 0
              End With
              oRange.FormatConditions(1).StopIfTrue = False
              oWorkbook.Save
              MsgBox strMsg, vbExclamation + vbOKOnly, "INVALID CODE"
              GoTo exit_here
            Else
              blnValid = True
            End If
            lngItems = oRange.Cells.Count
            lngItem = 0
            For Each c In oRange.Cells
              lngItem = lngItem + 1
              strValue = Trim(c.Value)
              strDescription = Trim(c.Offset(0, 2).Value)
              If Len(strDescription) > 0 Then
                Set oTask = ActiveProject.Tasks.Add(Left(strDescription, 255))
              Else
                Set oTask = ActiveProject.Tasks.Add("DELETE - PLACEHOLDER")
              End If
              oTask.OutlineLevel = 1
              oTask.SetField lngOutlineCode, strValue
              If oOutlineCode Is Nothing Then Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
              If oLookupTable Is Nothing Then Set oLookupTable = oOutlineCode.LookupTable
              If Len(strDescription) > 0 Then
                oLookupTable.Item(lngItem).Description = Left(strDescription, 255)
              End If
              If myBackbone_frm.chkAlsoCreateTasks Then
                lngOutlineLevel = Len(strValue) - Len(Replace(strValue, ".", ""))
                If lngOutlineLevel > 0 Then
                  oTask.OutlineLevel = lngOutlineLevel + 1
                End If
              Else
                oTask.Delete
              End If
              myBackbone_frm.lblStatus.Caption = "Importing " & lngItem & " of " & lngItems & " (" & Format(lngItem / lngItems, "0%") & ")..."
              myBackbone_frm.lblProgress.Width = (lngItem / lngItems) * myBackbone_frm.lblStatus.Width
            Next c
            myBackbone_frm.lblStatus.Caption = "Ready..."
            myBackbone_frm.lblProgress.Width = myBackbone_frm.lblStatus.Width
            'reset outline code to disallow new entries
            CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=True, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
            'refresh the form
            myBackbone_frm.cboOutlineCodes.List(myBackbone_frm.cboOutlineCodes.ListIndex, 1) = FieldConstantToFieldName(lngOutlineCode) & " (" & strOutlineCode & ")"
            'prevent importing multiple sheets
            Exit For
          Else
            MsgBox "No worksheet found where [A1:C1] contains CODE, LEVEL, DESCRIPTION.", vbExclamation + vbOKOnly, "Invalid Workbook"
          End If
        Next oWorksheet
      End If 'proper headers found
    End With
  End If 'proceed

exit_here:
  On Error Resume Next
  Set oValid = Nothing
  Set oInvalid = Nothing
  cptSpeed False
  Application.CloseUndoTransaction
  Set oTask = Nothing
  Set oOutlineCode = Nothing
  Set oLookupTable = Nothing
  Set c = Nothing
  Set oRange = Nothing
  Set oFileDialog = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  If blnValid Then
    oWorkbook.Close False
    oExcel.Quit
  Else
    oExcel.Visible = True
    oWorkbook.Activate
  End If
  Set oExcel = Nothing
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptImportCWBSFromExcel", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportCWBSFromServer(ByRef myBackbone_frm As cptBackbone_frm, lngOutlineCode As Long)
  'objects
  Dim c As Excel.Range
  Dim oTask As MSProject.Task
  Dim oRange As Excel.Range
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oLookupTable As MSProject.LookupTable
  Dim oOutlineCode As MSProject.OutlineCode
  Dim oFileDialog As Office.FileDialog
  Dim oExcel As Excel.Application
  'strings
  Dim strDescription As String
  Dim strCode As String
  Dim strOutlineCode As String
  'longs
  Dim lngItems As Long
  Dim lngOutlineLevel As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If MsgBox("Expected fields/column headers, in range [A1:C1], are LEVEL,VALUE,DESCRIPTION and there should be no blank rows." & vbCrLf & vbCrLf & "Proceed?", vbQuestion + vbYesNo, "Confirm CWBS Import") = vbYes Then
    strOutlineCode = CustomFieldGetName(lngOutlineCode)
    Set oExcel = CreateObject("Excel.Application")
    'allow user to select excel file and import it to chosen
    Set oFileDialog = oExcel.FileDialog(msoFileDialogFilePicker)
    With oFileDialog
      .AllowMultiSelect = False
      .ButtonName = "Import"
      .InitialView = 2 'msoFileDialogViewDetails
      .InitialFileName = Environ("USERPROFILE") & "\"
      .Title = "Select " & strOutlineCode & " source file:"
      .Filters.Add "Microsoft Excel Workbook (xlsx)", "*.xlsx"
      .Filters.Add "Comma Separated Values (csv)", "*.csv"
      If .Show = -1 Then
      
        Application.OpenUndoTransaction "Import " & strOutlineCode & " from MSP Server Outline Code Export"
      
        cptSpeed True
      
        'set up the outline code field
        For lngItem = 1 To 10
          CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=lngItem, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        Next lngItem
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
        
        Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
        'open the workbook
        Set oWorkbook = oExcel.Workbooks.Open(oFileDialog.SelectedItems(1))
        'find the sheet
        For Each oWorksheet In oWorkbook.Sheets
          If UCase(oWorksheet.[A1].Value) = "LEVEL" And UCase(oWorksheet.[B1].Value) = "VALUE" And UCase(oWorksheet.[C1].Value) = "DESCRIPTION" Then
            strOutlineCode = CustomFieldGetName(lngOutlineCode)
            Set oRange = oWorksheet.Range(oWorksheet.[A2], oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(-4162)) '-4162 = xlUp
            lngItems = oRange.Cells.Count
            lngItem = 0
            For Each c In oRange.Cells
              lngItem = lngItem + 1
              If oOutlineCode Is Nothing Then Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
              If oLookupTable Is Nothing Then Set oLookupTable = oOutlineCode.LookupTable
              Set oTask = ActiveProject.Tasks.Add(c.Offset(0, 2).Value)
              strCode = Left(c.Offset(0, 2), InStr(c.Offset(0, 2), " ") - 1)
              strDescription = Replace(c.Offset(0, 2), strCode & " - ", "")
              oTask.SetField lngOutlineCode, strCode
              oLookupTable.Item(lngItem).Description = strDescription
              myBackbone_frm.lblStatus.Caption = "Importing " & lngItem & " of " & lngItems & "(" & Format(lngItem / lngItems, "0%") & ")..."
              myBackbone_frm.lblProgress.Width = (lngItem / lngItems) * myBackbone_frm.lblStatus.Width
            Next c
            myBackbone_frm.lblStatus.Caption = "Ready..."
            myBackbone_frm.lblProgress.Width = myBackbone_frm.lblStatus.Width
            'reset outline code to disallow new entries
            CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=True, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
            'refresh the form
            myBackbone_frm.cboOutlineCodes.List(myBackbone_frm.cboOutlineCodes.ListIndex, 1) = FieldConstantToFieldName(lngOutlineCode) & " (" & strOutlineCode & ")"
            Exit For
          End If
        Next oWorksheet
      Else
        MsgBox "No worksheet found where [A1:C1] contains LEVEL, VALUE, DESCRIPTION.", vbExclamation + vbOKOnly, "Invalid Workbook"
      End If 'proper headers found
    End With
  End If 'proceed

exit_here:
  On Error Resume Next
  cptSpeed False
  Set c = Nothing
  Set oTask = Nothing
  Set oRange = Nothing
  Set oWorksheet = Nothing
  oWorkbook.Close False
  Set oWorkbook = Nothing
  Set oLookupTable = Nothing
  Set oOutlineCode = Nothing
  Set oFileDialog = Nothing
  oExcel.Quit
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptImportCWBSFromServer", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportAppendix(ByRef myBackbone_frm As cptBackbone_frm, lngOutlineCode As Long)
  'objects
  Dim oRecordset As ADODB.Recordset
  Dim oTaskTable As Object 'TaskTable
  Dim oTask As MSProject.Task
  'strings
  Dim strAppendix As String
  Dim strVersion As String
  Dim strFileName As String
  Dim strDir As String
  Dim strMsg As String
  'longs
  Dim lngItem As Long
  Dim lngField As Long
  Dim lngOutlineLevel As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  
  strVersion = Right(myBackbone_frm.cboImport, 1)
  strAppendix = myBackbone_frm.cboAppendix.Value
  
  Application.OpenUndoTransaction "Import MIL-STD-881" & strVersion & " Appendix " & strAppendix
  
  For lngItem = 1 To 10
    CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=lngItem, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  Next lngItem
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=False, LookupDefault:=False, SortOrder:=0
    
  strFileName = strDir & "/cpt-mil-std-881.adtg"
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strFileName, , adOpenKeyset, adLockReadOnly
  oRecordset.Filter = "VERSION='" & strVersion & "' AND APPENDIX='" & strAppendix & "'"
  lngItem = 0
  Do While Not oRecordset.EOF
    lngItem = lngItem + 1
    Set oTask = ActiveProject.Tasks.Add(oRecordset.Fields(3).Value)
    oTask.SetField lngOutlineCode, oRecordset.Fields(2).Value
    ActiveProject.OutlineCodes(CustomFieldGetName(lngOutlineCode)).LookupTable.Item(lngItem).Description = oRecordset.Fields(3).Value

    lngOutlineLevel = Len(oRecordset.Fields(2).Value) - Len(Replace(oRecordset.Fields(2).Value, ".", ""))
    If lngOutlineLevel > 0 Then
      oTask.OutlineLevel = lngOutlineLevel + 1
    End If
    
    oRecordset.MoveNext
  Loop
  oRecordset.Close
  
  'pretty up the task table
  If Len(ActiveProject.CurrentTable) > 0 Then
    SelectBeginning
    SetRowHeight 1, "all"
    Set oTaskTable = ActiveProject.TaskTables(ActiveProject.CurrentTable)
    For lngField = 1 To oTaskTable.TableFields.Count
      If FieldConstantToFieldName(oTaskTable.TableFields(lngField).Field) = "Name" Then
        ColumnBestFit lngField
        Exit For
      End If
    Next lngField
  End If
  
  'reset outline code to disallow new entries
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=True, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
  Call cptRefreshOutlineCodePreview(myBackbone_frm, CustomFieldGetName(lngOutlineCode))

exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  Application.CloseUndoTransaction
  Set oTaskTable = Nothing
  Set oTask = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptImportAppendix", Err, Erl)
  Resume exit_here
  
  
End Sub

Sub cptExportOutlineCodeToExcel(ByRef myBackbone_frm As cptBackbone_frm, lngOutlineCode As Long)
  'objects
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oListObject As Excel.ListObject
  Dim oLookupTable As MSProject.LookupTable
  Dim oOutlineCode As MSProject.OutlineCode
  'strings
  Dim strOutlineCode As String
  'longs
  Dim lngLastRow As Long
  Dim lngLookupItems As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strOutlineCode = CustomFieldGetName(lngOutlineCode)
  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oLookupTable Is Nothing Then
    MsgBox "There is no LookupTable associated with " & FieldConstantToFieldName(lngOutlineCode) & IIf(Len(strOutlineCode) > 0, " (" & strOutlineCode & ")", "") & ".", vbCritical + vbOKOnly, "No Code Defined"
    GoTo exit_here
  End If
  Application.StatusBar = "Exporting Outline Code '" & strOutlineCode & "'..."
  myBackbone_frm.lblStatus.Caption = Application.StatusBar
  
  'get excel
  Application.StatusBar = "Setting up Excel..."
  myBackbone_frm.lblStatus.Caption = Application.StatusBar
  Set oExcel = CreateObject("Excel.Application")
  Set oWorkbook = oExcel.Workbooks.Add
  oExcel.Calculation = -4135 'xlCalculationManual
  oExcel.ScreenUpdating = False
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Outline.SummaryRow = 0 'xlSummaryAbove
  oWorksheet.[A1:C1] = Array("CODE", "LEVEL", "DESCRIPTION")
  
  'export the codes
  For lngLookupItems = 1 To oLookupTable.Count
    lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(-4162).Row + 1 '-4162 = xlUp
    oWorksheet.Cells(lngLastRow, 1).Value = "'" & oLookupTable.Item(lngLookupItems).FullName
    oWorksheet.Cells(lngLastRow, 2).Value = oLookupTable.Item(lngLookupItems).Level
    oWorksheet.Cells(lngLastRow, 3).Value = oLookupTable.Item(lngLookupItems).Description
    oWorksheet.Cells(lngLastRow, 3).IndentLevel = oLookupTable.Item(lngLookupItems).Level - 1
    If oLookupTable.Item(lngLookupItems).Level > 8 Then
      oWorksheet.Rows(lngLastRow).OutlineLevel = 8
      oWorksheet.Cells(lngLastRow, 2).AddComment "Excel grouping limited to 8 levels"
    Else
      oWorksheet.Rows(lngLastRow).OutlineLevel = oLookupTable.Item(lngLookupItems).Level
    End If
    myBackbone_frm.lblProgress.Width = (lngLookupItems / oLookupTable.Count) * myBackbone_frm.lblStatus.Width
    myBackbone_frm.lblStatus.Caption = "Exporting...(" & Format(lngLookupItems / oLookupTable.Count, "0%") & ")"
  Next lngLookupItems
  
  Application.StatusBar = "Formatting Worksheet..."
  myBackbone_frm.lblStatus.Caption = Application.StatusBar
  
  'format the table
  oExcel.ActiveWindow.Zoom = 85
  'Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)), , xlYes)
  Set oListObject = oWorksheet.ListObjects.Add(1, oWorksheet.Range(oWorksheet.[A1].End(-4161), oWorksheet.[A1].End(-4121)), , 1)
  oListObject.Name = strOutlineCode
  oListObject.TableStyle = ""
  oListObject.HeaderRowRange.Font.Bold = True
  oListObject.Range.Borders(5).LineStyle = -4142 'xlDiagonalDown = xlNone
  oListObject.Range.Borders(6).LineStyle = -4142 'xlDiagonalUp = xlNone
  With oListObject.Range.Borders(7) 'xlEdgeLeft
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(8) 'xlEdgeTop
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(9) 'xlEdgeBottom
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(10) 'xlEdgeRight
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(11) 'xlInsideVertical
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946592608417
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(12) 'xlInsideHorizontal
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946592608417
    .Weight = 2 'xlThin
  End With
  With oListObject.HeaderRowRange.Interior
    .Pattern = 1 'xlSolid
    .PatternColorIndex = -4105 'xlAutomatic
    .ThemeColor = 1 'xlThemeColorDark1
    .TintAndShade = -0.149998474074526
    .PatternTintAndShade = 0
  End With
  oWorksheet.Name = strOutlineCode
  oWorksheet.[A2].Select
  oExcel.ActiveWindow.FreezePanes = True
  oWorksheet.Columns.AutoFit
    
exit_here:
  On Error Resume Next
  Set oLookupTable = Nothing
  Application.StatusBar = "Ready..."
  myBackbone_frm.lblStatus.Caption = Application.StatusBar
  myBackbone_frm.lblProgress.Width = myBackbone_frm.lblStatus.Width
  oExcel.Visible = True
  oExcel.ScreenUpdating = True
  oExcel.Calculation = -4105 'xlCalculationAutomatic
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oOutlineCode = Nothing

  Exit Sub
  
err_here:
  Call cptHandleErr(THIS_MODULE, "ExportOutlineCode", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptExport81334D(ByRef myBackbone_frm As cptBackbone_frm, lngOutlineCode As Long)
  'objects
  Dim oMailItem As Outlook.MailItem
  Dim oOutlook As Outlook.Application
  Dim oLookupTable As MSProject.LookupTable
  Dim oOutlineCode As MSProject.OutlineCode
  Dim wsDictionary As Excel.Worksheet
  Dim wsIndex As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  Dim oStream As ADODB.Stream
  Dim oXMLHttpDoc As Object
  Dim oShell As Object
  'strings
  Dim strOutlineCode As String
  Dim strURL As String
  Dim strTemplateDir As String
  Dim strTemplate As String
  'longs
  Dim lngBorder As Long
  Dim lngRow As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'get outline code name and export it
  myBackbone_frm.lblStatus.Caption = "Exporting..."
  strOutlineCode = CustomFieldGetName(lngOutlineCode)
  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oLookupTable Is Nothing Then
    MsgBox "There is no LookupTable associated with " & FieldConstantToFieldName(lngOutlineCode) & IIf(Len(strOutlineCode) > 0, " (" & strOutlineCode & ")", "") & ".", vbCritical + vbOKOnly, "No Code Defined"
    GoTo exit_here
  Else
  
    'first determine if user has the template installed
    Set oShell = CreateObject("WScript.Shell")
    strTemplateDir = oShell.SpecialFolders("Templates")
    strTemplate = "81334D_CWBS_TEMPLATE.xltm"
    
    If Dir(strTemplateDir & "\" & strTemplate) = vbNullString Then
      'provide user feedback
      myBackbone_frm.lblStatus.Caption = "Downloading template..."
      Set oXMLHttpDoc = CreateObject("Microsoft.XMLHTTP")
      strURL = strGitHub & "Templates/" & strTemplate
      oXMLHttpDoc.Open "GET", strURL, False
      oXMLHttpDoc.Send
      If oXMLHttpDoc.Status = 200 And oXMLHttpDoc.readyState = 4 Then
        'success: save it to templates directory
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1 'adTypeBinary
        oStream.Write oXMLHttpDoc.responseBody
        oStream.SaveToFile strTemplateDir & "\" & strTemplate
        oStream.Close
      Else
        myBackbone_frm.lblStatus.Caption = "Download failed."
        'fail: prompt to request by email
        If MsgBox("Unable to download template. Request via email?", vbExclamation + vbYesNo, "No Connection") = vbYes Then
          MsgBox "When the template arrives, please save to:" & vbCrLf & vbCrLf & strTemplateDir, vbOKOnly + vbInformation, "Save Location"
          Shell "explorer.exe " & strTemplateDir, vbNormalFocus
          On Error Resume Next
          Set oOutlook = GetObject(, "Outlook.Application")
          If oOutlook Is Nothing Then
            Set oOutlook = CreateObject("Outlook.Application")
          End If
          If oOutlook Is Nothing Then
            MsgBox "Outlook is not available.", vbCritical + vbOKOnly, "Request 81334D"
            GoTo exit_here
          End If
          Set oMailItem = oOutlook.CreateItem(0) '0 = olMailItem
          oMailItem.To = "help@ClearPlanConsulting.com"
          oMailItem.Importance = 2 '2=olImportanceHigh
          oMailItem.Subject = "Template Request: " & strTemplate
          oMailItem.HTMLBody = "Please forward the subject-referenced template. Thank you." & oMailItem.HTMLBody
          oMailItem.Display False
        End If
        GoTo exit_here
      End If
      
    End If
  
    'open excel and create template
    Set oExcel = CreateObject("Excel.Application")
    Set oWorkbook = oExcel.Workbooks.Add(strTemplateDir & "\" & strTemplate)
    oExcel.Calculation = -4135 'xlManual
    oExcel.ScreenUpdating = False
    Set wsIndex = oWorkbook.Sheets("CWBS Index")
    wsIndex.Outline.SummaryRow = 0 'xlSummaryAbove
    Set wsDictionary = oWorkbook.Sheets("CWBS Dictionary")
    wsDictionary.Outline.SummaryRow = 0 'xlSummaryAbove
    lngRow = 7
    For lngItem = 1 To oLookupTable.Count
      'index: code=col1; name=col9
      wsIndex.Cells(lngRow, 1).Value = "'" & oLookupTable.Item(lngItem).FullName
      wsIndex.Cells(lngRow, 10).Value = oLookupTable.Item(lngItem).Description
      wsIndex.Cells(lngRow, 10).HorizontalAlignment = -4131 'xlLeft
      wsIndex.Cells(lngRow, 10).IndentLevel = Len(CStr(oLookupTable.Item(lngItem).FullName)) - Len(Replace(CStr(oLookupTable.Item(lngItem).FullName), ".", ""))
      wsIndex.Rows(lngRow).OutlineLevel = Len(CStr(oLookupTable.Item(lngItem).FullName)) - Len(Replace(CStr(oLookupTable.Item(lngItem).FullName), ".", "")) + 1
      If lngRow >= 8 Then
        wsIndex.Range(wsIndex.Cells(lngRow, 10), wsIndex.Cells(lngRow, 19)).Merge
      End If
      'dictionary: code=col1; name=col2
      wsDictionary.Cells(lngRow, 1).Value = "'" & oLookupTable.Item(lngItem).FullName
      wsDictionary.Cells(lngRow, 2).Value = oLookupTable.Item(lngItem).Description
      wsDictionary.Cells(lngRow, 2).HorizontalAlignment = -4131 'xlLeft
      wsDictionary.Cells(lngRow, 2).IndentLevel = wsIndex.Cells(lngRow, 10).IndentLevel
      wsDictionary.Rows(lngRow).OutlineLevel = wsIndex.Rows(lngRow).OutlineLevel
      If lngRow >= 8 Then
        wsDictionary.Range(wsDictionary.Cells(lngRow, 2), wsDictionary.Cells(lngRow, 3)).Merge
        wsDictionary.Range(wsDictionary.Cells(lngRow, 4), wsDictionary.Cells(lngRow, 11)).Merge
      End If
      myBackbone_frm.lblStatus.Caption = "Exporting...(" & Format(lngItem / oLookupTable.Count, "0%") & ")"
      myBackbone_frm.lblProgress.Width = (lngItem / oLookupTable.Count) * myBackbone_frm.lblStatus.Width
      lngRow = lngRow + 1
    Next
  End If
  
  'format it
  '-4121=-4121; -4161=xlToRight; 1=xlContinuous; 2=xlThin; -4105=xlColorIndexAutomatic
  wsIndex.[B8:I8].AutoFill Destination:=wsIndex.Range(wsIndex.Cells(8, 2), wsIndex.Cells(7 + oLookupTable.Count - 1, 9))
  For lngBorder = 7 To 12 'left,top,bottom,right,insidevertical,insidehorizontal
    With wsIndex.Range(wsIndex.[A7].End(-4121), wsIndex.Cells(7, 19)).Borders(lngBorder)
      .LineStyle = 1
      .Weight = 2
      .ColorIndex = -4105
    End With
    With wsDictionary.Range(wsDictionary.[A7].End(-4121), wsDictionary.Cells(7, 11)).Borders(lngBorder)
      .LineStyle = 1
      .Weight = 2
      .ColorIndex = -4105
    End With
  Next lngBorder
  wsDictionary.Range(wsDictionary.[A7].End(-4121), wsDictionary.[A7].End(-4161)).BorderAround 1, 2, -4105
  
  'freeze panes
  wsDictionary.Activate
  wsDictionary.[A7].Select
  oExcel.ActiveWindow.FreezePanes = True
  wsIndex.Activate
  wsIndex.[A7].Select
  oExcel.ActiveWindow.FreezePanes = True
  oExcel.Visible = True
  
  'provide user feedback
  myBackbone_frm.lblStatus.Caption = "Complete."
  
exit_here:
  On Error Resume Next
  Set oMailItem = Nothing
  myBackbone_frm.lblStatus.Caption = "Ready..."
  myBackbone_frm.lblProgress.Width = myBackbone_frm.lblStatus.Width
  Set oLookupTable = Nothing
  Set oOutlineCode = Nothing
  Set wsDictionary = Nothing
  Set wsIndex = Nothing
  oExcel.Calculation = -4105 'xlAutomatic
  oExcel.ScreenUpdating = True
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oStream = Nothing
  Set oXMLHttpDoc = Nothing
  Set oShell = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptExport81334D", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportTemplate()
  'objects
  Dim oWorksheet As Object
  Dim oWorkbook As Object
  Dim oExcel As Object
  'strings
  Dim strMsg As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strMsg = "Instructions:" & vbCrLf
  strMsg = strMsg & "1. Do not add, edit, move, or remove columns." & vbCrLf
  strMsg = strMsg & "2. No empty rows from row 2 to the end of your Code." & vbCrLf
  strMsg = strMsg & "3. Save and import when done." & vbCrLf & vbCrLf
  strMsg = strMsg & "- CWBS SUGGESTION: Include down to Control Account levels, suffixed with ' CA'" & vbCrLf
  strMsg = strMsg & "- IMP SUGGESTION: Include down to an accomplishment criteria milestone." & vbCrLf & vbCrLf
  strMsg = strMsg & "Proceed?"
  If MsgBox(strMsg, vbInformation + vbYesNo, "Instructions:") = vbYes Then
    Set oExcel = CreateObject("Excel.Application")
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.Name = "CWBS"
    oWorksheet.[A1:C1] = Array("CODE", "LEVEL", "DESCRIPTION")
    oWorksheet.[A1:C1].Font.Bold = True
    oWorksheet.[A2].Select
    oWorksheet.Columns(1).ColumnWidth = 10
    oWorksheet.Columns(2).ColumnWidth = 5.2
    oWorksheet.Columns(3).ColumnWidth = 59.14
    oExcel.ActiveWindow.FreezePanes = True
    oExcel.ActiveWindow.Zoom = 85
    oExcel.Visible = True
    Application.ActivateMicrosoftApp pjMicrosoftExcel
  End If
  
exit_here:
  On Error Resume Next
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptExportTemplate", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowBackbone_frm()
  'objects
  Dim xmlHttpDoc As Object
  Dim oStream As Object 'ADODB.Stream
  Dim oRecordset As ADODB.Recordset
  Dim oDict As Scripting.Dictionary
  Dim myBackbone_frm As cptBackbone_frm
  'longs
  Dim lngCode As Long
  Dim lngOutlineCode As Long
  Dim lngItem As Long
  'strings
  Dim strOutlineCode As String
  Dim strOutlineCodeName As String
  Dim strFileName As String
  Dim strURL As String
  Dim strMsg As String

  'prevent spawning
  If Not cptGetUserForm("cptBackbone_frm") Is Nothing Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set myBackbone_frm = New cptBackbone_frm
  With myBackbone_frm.cboOutlineCodes
    .Clear
    'populate the listbox/combobox
    For lngCode = 1 To 10
      strOutlineCode = "Outline Code" & lngCode
      lngOutlineCode = Application.FieldNameToFieldConstant(strOutlineCode)
      strOutlineCodeName = Application.CustomFieldGetName(lngOutlineCode)
      .AddItem
      If Len(strOutlineCodeName) > 0 Then
        strOutlineCode = strOutlineCode & " (" & strOutlineCodeName & ")"
      End If
      .List(lngCode - 1, 0) = lngOutlineCode
      .List(lngCode - 1, 1) = strOutlineCode
    Next lngCode
  End With
  
  'get latest 881 dictionary version (if possible)
  strFileName = cptDir & "\cpt-mil-std-881.adtg"
  Set xmlHttpDoc = CreateObject("Microsoft.XMLHTTP")
  strURL = "https://raw.githubusercontent.com/clearplan/cpt/master/Backbone/cpt-mil-std-881.adtg"
  xmlHttpDoc.Open "GET", strURL, False
  xmlHttpDoc.Send
  If xmlHttpDoc.Status = 200 And xmlHttpDoc.readyState = 4 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1 'adTypeBinary
    oStream.Write xmlHttpDoc.responseBody
    If Dir(strFileName) <> vbNullString Then Kill strFileName
    oStream.SaveToFile strFileName
    oStream.Close
  End If
    
  'add Import Actions
  With myBackbone_frm.cboImport
    .Clear
    .AddItem "From Excel Workbook"
    .AddItem "From MSP Server Outline Code Export"
    .AddItem "From Existing Tasks"
    'add options for mil-std-881
    If Dir(strFileName) <> vbNullString Then
      Set oDict = CreateObject("Scripting.Dictionary")
      Set oRecordset = CreateObject("ADODB.Recordset")
      oRecordset.Open strFileName, , adOpenKeyset, adLockReadOnly
      oRecordset.Filter = "CODE='1'"
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
        If Not oDict.Exists(oRecordset(0).Value) Then
          oDict.Add oRecordset(0).Value, oRecordset(0).Value
        End If
        oRecordset.MoveNext
      Loop
      oRecordset.Close
      Set oRecordset = Nothing
      For lngItem = oDict.Count - 1 To 0 Step -1 'in reverse so that latest versions are at top
        .AddItem "From MIL-STD-881" & oDict.Keys(lngItem)
      Next lngItem
      Set oDict = Nothing
    Else
      'not sure if we can send an *.adtg file via email?
      strMsg = "Could not download the MIL-STD-881 dictionary." & vbCrLf & vbCrLf
      strMsg = strMsg & "If you would like to import from the MIL-STD-881 dictionary, please:" & vbCrLf
      strMsg = strMsg & "1. download the MIL-STD-881 dictionary" & vbCrLf
      strMsg = strMsg & "2. install to: " & cptDir & "\" & vbCrLf & vbCrLf
      strMsg = strMsg & "Alternatively, request it from help@ClearPlanConsulting.com" & vbCrLf & vbCrLf
      strMsg = strMsg & "Download from URL:"
      InputBox strMsg, "File Not Found!", strURL
    End If
    
  End With
  
  myBackbone_frm.cboAppendix.Enabled = False
  
  'add Export Actions
  With myBackbone_frm.cboExport
    .Clear
    .AddItem "To Excel Workbook"
    .AddItem "To CSV for MPM"
    .AddItem "To CSV for COBRA"
    .AddItem "To DI-MGMT-81334D Template"
  End With
  
  'pre-select Outline Code 1
  With myBackbone_frm
    '.cboOutlineCodes.ListIndex = 0
    '.txtNameIt = CustomFieldGetName(.cboOutlineCodes.List(0, 0))
    .Caption = "Backbone (" & cptGetVersion("cptBackbone_frm") & ")"
    .cboOutlineCodes.SetFocus
    cptBackboneHideControls myBackbone_frm
    .Show '(False)
  End With

exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  Set xmlHttpDoc = Nothing
  Set oStream = Nothing
  Set oDict = Nothing
  Unload myBackbone_frm
  Set myBackbone_frm = Nothing
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptShowBackbone_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptCreateCode(ByRef myBackbone_frm As cptBackbone_frm, lngOutlineCode As Long)
  'objects
  Dim objOutlineCode As OutlineCode
  Dim objLookupTable As LookupTable
  Dim objLookupTableEntry As LookupTableEntry
  Dim oTask As MSProject.Task
  'strings
  Dim strWBS As String
  Dim strParent As String
  Dim strChild As String
  'longs
  Dim lngUID As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  Dim lngItem As Long
  'variants
  Dim aOutlineCode As Variant
  'dates
  Dim tmr As Date

  tmr = Now
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'ensure name doesn't already exist - trust form formatting
  If myBackbone_frm.txtNameIt.BorderColor = 255 Then GoTo exit_here

  'first name the field and create the code mask
  For lngItem = 1 To 10
    CustomOutlineCodeEditEx lngOutlineCode, Level:=lngItem, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  Next lngItem
  CustomOutlineCodeEditEx lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=False, LookupDefault:=False, SortOrder:=0
  Set objOutlineCode = ActiveProject.OutlineCodes(CustomFieldGetName(lngOutlineCode))
  Set objLookupTable = objOutlineCode.LookupTable
  
  lngTasks = ActiveProject.Tasks.Count
  
  For Each oTask In ActiveProject.Tasks
    If Not oTask Is Nothing Then
      lngTask = lngTask + 1
      oTask.SetField lngOutlineCode, oTask.WBS
      objLookupTable.Item(lngTask).Description = oTask.Name
      myBackbone_frm.lblProgress.Width = ((lngTask - 1) / lngTasks) * myBackbone_frm.lblStatus.Width
      myBackbone_frm.lblStatus.Caption = Format(lngTask - 1, "#,##0") & " / " & Format(lngTasks, "#,##0") & " (" & Format((lngTask - 1) / lngTasks, "0%") & ") [" & Format(Now - tmr, "hh:nn:ss") & "]"
    End If 'task is nothing
  Next oTask
  CustomOutlineCodeEditEx lngOutlineCode, OnlyLeaves:=True, OnlyLookUpTableCodes:=True
  myBackbone_frm.lblStatus.Caption = "Complete."
  Application.StatusBar = "Complete."
  myBackbone_frm.cmdCancel.Caption = "Done"
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  cptSpeed False
  Set objOutlineCode = Nothing
  Set objLookupTable = Nothing
  Set objLookupTableEntry = Nothing
  Set oTask = Nothing
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptCreateCode", Err, Erl)
  Resume exit_here
End Sub

Sub cptRenameInsideOutlineCode(ByRef myBackbone_frm As cptBackbone_frm, strOutlineCode As String, strFind As String, strReplace As String)
  'usage: Call RenameOutlineCode("CWBS","BOSS","IBRS")
  'objects
  Dim oOutlineCode As OutlineCode, oLookupTable As LookupTable, oLookupTableEntry As LookupTableEntry
  'longs
  Dim lngEntry As Long
  Dim lngReplaced As Long
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  Set oLookupTable = oOutlineCode.LookupTable
  For lngEntry = 1 To oLookupTable.Count
    If InStr(oLookupTable(lngEntry).Description, strFind) > 0 Then
      oLookupTable(lngEntry).Description = Replace(oLookupTable(lngEntry).Description, strFind, strReplace)
      lngReplaced = lngReplaced + 1
    End If
  Next lngEntry
  
  myBackbone_frm.lblFeedback.Caption = Format(lngReplaced, "#,##0") & " replaced"
  
exit_here:
  On Error Resume Next
  Set oOutlineCode = Nothing
  Set oLookupTable = Nothing
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptRenameInsideOutlineCode", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshOutlineCodePreview(ByRef myBackbone_frm As cptBackbone_frm, strOutlineCode As String)
  'objects
  Dim oOutlineCode As OutlineCode, oLookupTable As LookupTable, oLookupTableEntry As LookupTableEntry
  Dim oNode As Object 'Node
  'strings
  'longs
  Dim lngEntries As Long
  Dim lngEntry As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strOutlineCode = Replace(Replace(strOutlineCode, cptRegEx(strOutlineCode, "Outline Code[0-9]{1,}") & " (", ""), ")", "")
  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oLookupTable Is Nothing Then
    If oLookupTable.Count > 0 Then
      lngEntries = oLookupTable.Count
      myBackbone_frm.lboOutlineCode.Clear
      For lngEntry = 1 To oLookupTable.Count
        With myBackbone_frm
          With .lboOutlineCode
            .AddItem
            .List(.ListCount - 1, 0) = oLookupTable(lngEntry).UniqueID
            .List(.ListCount - 1, 1) = oLookupTable(lngEntry).Level
            .List(.ListCount - 1, 2) = oLookupTable(lngEntry).FullName & " - " & oLookupTable(lngEntry).Description
          End With
          .lblStatus.Caption = "Loading...(" & Format(lngEntry / lngEntries, "0%") & ")"
          .lblProgress.Width = (lngEntry / lngEntries) * .lblStatus.Width
          If .Visible Then DoEvents
          Application.StatusBar = "Adding: " & oLookupTable(lngEntry).FullName & " - " & oLookupTable(lngEntry).Description
        End With
      Next lngEntry
    End If 'lookuptable.count > 0
  End If 'lookuptable is nothing
  With myBackbone_frm
    .lblProgress.Width = .lblStatus.Width
    .lblStatus.Caption = "Ready..."
  End With
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oOutlineCode = Nothing
  Set oLookupTable = Nothing
  Set oLookupTableEntry = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptRefreshOutlineCodePreview", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptExportOutlineCodeForMPM(ByRef myBackbone_frm As cptBackbone_frm, lngOutlineCode As Long)
  'exports local Outline Code to CSV for MPM Upload
  'objects
  Dim oOutlineCode As OutlineCode
  Dim oLookupTable As LookupTable
  'longs
  Dim lngItem As Long, lngFile As Long
  'strings
  Dim strHeader As String
  Dim strMsg As String
  Dim strCode As String, strDescription As String, strParent As String
  Dim strDir As String, strFileName As String, strOutlineCode As String
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnCA As Boolean
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'confirm lookuptable exists
  Set oOutlineCode = ActiveProject.OutlineCodes(CustomFieldGetName(lngOutlineCode))
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oLookupTable Is Nothing Then
    strOutlineCode = CustomFieldGetName(lngOutlineCode)
    MsgBox "There is no LookupTable associated with " & FieldConstantToFieldName(lngOutlineCode) & IIf(Len(strOutlineCode) > 0, " (" & strOutlineCode & ")", "") & ".", vbExclamation + vbOKOnly, "No LookupTable"
    GoTo exit_here
  End If
  
  'set directory
  strDir = Environ("TEMP") & "\"
  strFileName = "WBS_DESCRIPTIVE_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".csv"
  If Dir(strDir & strFileName) <> vbNullString Then Kill strDir & strFileName

  lngFile = FreeFile
  Open strDir & strFileName For Output As #lngFile
  
  If myBackbone_frm.chkIncludeHeaders Then
    strHeader = "WBS ID,"
    strHeader = strHeader & "WBS Description,"
    strHeader = strHeader & "Alias,"
    strHeader = strHeader & "XREF1,"
    strHeader = strHeader & "XREF2,"
    strHeader = strHeader & "XREF3,"
    strHeader = strHeader & "XREF4,"
    strHeader = strHeader & "XREF5,"
    strHeader = strHeader & "XREF6,"
    strHeader = strHeader & "XREF7,"
    strHeader = strHeader & "XREF8,"
    strHeader = strHeader & "XREF9,"
    strHeader = strHeader & "XREF10,"
    strHeader = strHeader & "Manager,"
    strHeader = strHeader & "Charge Number,"
    strHeader = strHeader & "Performing Department,"
    strHeader = strHeader & "Responsible Department,"
    strHeader = strHeader & "Element Type,"
    strHeader = strHeader & "Earned Value Method,"
    strHeader = strHeader & "CLIN,"
    strHeader = strHeader & "Recurring or non-recurring,"
    strHeader = strHeader & "Fee %,"
    strHeader = strHeader & "Fee Limit Amount,"
    strHeader = strHeader & "BCWP Base Unit,"
    strHeader = strHeader & "Parent WBS ID,"
    strHeader = strHeader & "Base WBS,"
    Print #lngFile, strHeader
  End If
  
  'output top level
  Print #lngFile, "*" & "," & Chr(34) & ActiveProject.ProjectSummaryTask.Name & Chr(34) & String(25, ",")
  For lngItem = 1 To oLookupTable.Count
    strCode = oLookupTable(lngItem).FullName
    strDescription = oLookupTable(lngItem).Description
    If Not oLookupTable(lngItem).IsValid Then
      MsgBox "Invalid Code Found! See " & strCode & " : " & strDescription, vbCritical + vbOKOnly, "Error"
      GoTo kill_file
    End If
    blnCA = Right(strDescription, 3) = " CA"
    If Len(strCode) = 1 Then
      strParent = "*"
    Else
      strParent = Left(strCode, InStrRev(strCode, ".") - 1)
    End If
    myBackbone_frm.lblStatus.Caption = "Exporting...(" & Format(lngItem / oLookupTable.Count, "0%") & ")"
    myBackbone_frm.lblProgress.Width = (lngItem / oLookupTable.Count) * myBackbone_frm.lblStatus.Width
    Print #lngFile, strCode & "," & Chr(34) & strDescription & Chr(34) & String(16, ",") & IIf(blnCA, "C", "") & String(7, ",") & strParent & ",,"
  Next lngItem
  
  Close #lngFile
  
  'open it in notepad
  cptShellExecute 0, "open", "notepad.exe", strDir & strFileName, vbNullString, 1
  
exit_here:
  On Error Resume Next
  Set oLookupTable = Nothing
  Set oOutlineCode = Nothing
  Reset 'closes all active files opened by the Open statement and writes the contents of all file buffers to disk.
  Exit Sub
  
kill_file:
  On Error Resume Next
  Close #lngFile
  Kill strDir & strFileName
  Resume exit_here
  
err_here:
  Call cptHandleErr(THIS_MODULE, "cptExportOutlineCodeForMPM", Err, Erl)
  Resume exit_here

End Sub

Sub cptBackboneHideControls(ByRef myBackbone_frm As cptBackbone_frm)

  With myBackbone_frm
    'Replace
    .lblFeedback.Visible = .optReplace
    .txtReplace.Enabled = .optReplace
    .txtReplacement.Enabled = .optReplace
    .cmdReplace.Enabled = .optReplace
    'Import
    .txtNameIt.Enabled = .optImport
    .cboImport.Enabled = .optImport
    .chkAlsoCreateTasks.Enabled = .optImport
    .cmdExportTemplate.Visible = False
    .cmdImport.Enabled = .optImport
    'Export
    .cboExport.Enabled = .optExport
    .chkIncludeHeaders.Enabled = .optExport
    .chkIncludeThresholds.Enabled = .optExport
    .cmdExport.Enabled = .optExport
  End With

End Sub

Sub cptExportOutlineCodeForCOBRA(ByRef myBackbone_frm As cptBackbone_frm, lngOutlineCode)
  'objects
  Dim oLookupTable As LookupTable
  Dim oOutlineCode As OutlineCode
  'strings
  Dim strOutlineCode As String
  Dim strDescription As String
  Dim strCode As String
  Dim strFileName As String
  Dim strHeader As String
  'longs
  Dim lngItem As Long
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnIncludeThresholds As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'confirm lookuptable exists
  Set oOutlineCode = ActiveProject.OutlineCodes(CustomFieldGetName(lngOutlineCode))
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oLookupTable Is Nothing Then
    strOutlineCode = CustomFieldGetName(lngOutlineCode)
    MsgBox "There is no LookupTable associated with " & FieldConstantToFieldName(lngOutlineCode) & IIf(Len(strOutlineCode) > 0, " (" & strOutlineCode & ")", "") & ".", vbExclamation + vbOKOnly, "No LookupTable"
    GoTo exit_here
  End If
  
  'setup the export file
  strFileName = Environ("TEMP") & "\CODE_FILE_WBS.csv"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  lngFile = FreeFile
  Open strFileName For Output As #lngFile
  
  'export header
  strHeader = "Code,"
  strHeader = strHeader & "Description,"
  blnIncludeThresholds = myBackbone_frm.chkIncludeThresholds
  If blnIncludeThresholds Then
    strHeader = strHeader & "Threshold SV Value Current Period Favorable,"
    strHeader = strHeader & "Threshold SV Value Current Period Unfavorable,"
    strHeader = strHeader & "Threshold SV % Current Period Favorable,"
    strHeader = strHeader & "Threshold SV % Current Period Unfavorable,"
    strHeader = strHeader & "Threshold SV Value Cumulative Favorable,"
    strHeader = strHeader & "Threshold SV Value Cumulative Unfavorable,"
    strHeader = strHeader & "Threshold SV % Cumulative Favorable,"
    strHeader = strHeader & "Threshold SV % Cumulative Unfavorable,"
    strHeader = strHeader & "Threshold CV Value Current Period Favorable,"
    strHeader = strHeader & "Threshold CV Value Current Period Unfavorable,"
    strHeader = strHeader & "Threshold CV % Current Period Favorable,"
    strHeader = strHeader & "Threshold CV % Current Period Unfavorable,"
    strHeader = strHeader & "Threshold CV Value Cumulative Favorable,"
    strHeader = strHeader & "Threshold CV Value Cumulative Unfavorable,"
    strHeader = strHeader & "Threshold CV % Cumulative Favorable,"
    strHeader = strHeader & "Threshold CV % Cumulative Unfavorable,"
    strHeader = strHeader & "Threshold At Complete Value Favorable,"
    strHeader = strHeader & "Threshold At Complete Value Unfavorable,"
    strHeader = strHeader & "Threshold At Complete % Favorable,"
    strHeader = strHeader & "Threshold At Complete % Unfavorable"
  End If
  
  Print #lngFile, strHeader
  
  'export outline code
  For lngItem = 1 To oLookupTable.Count
    strCode = oLookupTable(lngItem).FullName
    strDescription = oLookupTable(lngItem).Description
    If Not oLookupTable(lngItem).IsValid Then
      MsgBox "Invalid Code Found! See " & strCode & " : " & strDescription, vbCritical + vbOKOnly, "Error"
      GoTo kill_file
    End If
    myBackbone_frm.lblStatus.Caption = "Exporting...(" & Format(lngItem / oLookupTable.Count, "0%") & ")"
    myBackbone_frm.lblProgress.Width = (lngItem / oLookupTable.Count) * myBackbone_frm.lblStatus.Width
    Print #lngFile, strCode & "," & Chr(34) & strDescription & Chr(34) & IIf(blnIncludeThresholds, String(20, ","), ",") '2 fields OR 22 fields
  Next lngItem

  Close #lngFile
  
  cptShellExecute 0, "open", "notepad.exe", strFileName, vbNullString, 1

exit_here:
  On Error Resume Next
  Set oLookupTable = Nothing
  Set oOutlineCode = Nothing
  Reset 'closes all active files opened by the Open statement and writes the contents of all file buffers to disk.
  Exit Sub
  
kill_file:
  On Error Resume Next
  Close #lngFile
  Kill strFileName
  Resume exit_here
  
err_here:
  Call cptHandleErr(THIS_MODULE, "cptExportOutlineCodeForCOBRA", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptExportAllCodes()
  'exports all lookups from all LCFs (Flags have no lookups)
  'does not interrogate ECFs
  'objects
  Dim oFieldCounts As Scripting.Dictionary
  Dim oCodes As Scripting.Dictionary
  Dim oOutlineCode As OutlineCode
  Dim oLookupTable As LookupTable
  Dim oLookupTableEntry As LookupTableEntry
  'strings
  Dim strDescription As String
  Dim strValue As String
  Dim strFileName As String
  Dim strFN As String
  Dim strCFN As String
  'longs
  Dim lngCodes As Long
  Dim lngFile As Long
  Dim lngCF As Long
  Dim lngListItem As Long
  Dim lngItem As Long
  Dim lngItems As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vFieldType As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSpeed True
  
  'first do outline codes, because they're different
  For Each oOutlineCode In ActiveProject.OutlineCodes
    lngCF = oOutlineCode.FieldID
    strCFN = cptRemoveIllegalCharacters(CustomFieldGetName(lngCF))
    Set oLookupTable = oOutlineCode.LookupTable
    lngItems = oLookupTable.Count
    If lngItems > 0 Then
      Application.StatusBar = "Exporting Code file for " & strCFN & "..."
      lngFile = FreeFile
      strFileName = Environ("tmp") & "\" & Replace(strCFN, " ", "_") & ".csv"
      Open strFileName For Output As #lngFile
      Print #lngFile, "CODE,DESCRIPTION,PARENT"
      For lngItem = 1 To oLookupTable.Count
        Set oLookupTableEntry = oLookupTable(lngItem)
        If oLookupTableEntry.Level = 1 Then
          Print #lngFile, oLookupTableEntry.FullName & "," & Chr(34) & oLookupTableEntry.Description & Chr(34) & ",*****"
        Else
          Print #lngFile, oLookupTableEntry.FullName & "," & Chr(34) & oLookupTableEntry.Description & Chr(34) & "," & oLookupTableEntry.ParentEntry.FullName
        End If
      Next lngItem
      Close #lngFile
      cptShellExecute 0, "open", "notepad.exe", strFileName, vbNullString, 1
      Application.StatusBar = "Exporting Code file for " & strCFN & "...done."
      lngCodes = lngCodes + 1
    End If
  Next oOutlineCode
  
  Set oFieldCounts = CreateObject("Scripting.Dictionary")
  oFieldCounts.Add "Text", 30
  oFieldCounts.Add "Number", 20
  
  Set oCodes = CreateObject("Scripting.Dictionary")
  
  For Each vFieldType In Array("Cost", "Date", "Duration", "Finish", "Number", "Start", "Text") 'Flag has no picklist
    If oFieldCounts.Exists(vFieldType) Then
      lngItems = oFieldCounts(vFieldType)
    Else
      lngItems = 10
    End If
    For lngItem = 1 To lngItems
      strFN = vFieldType & lngItem
      lngCF = FieldNameToFieldConstant(strFN)
      strCFN = CustomFieldGetName(lngCF)
      If Len(strCFN) > 0 Then
        strCFN = cptRemoveIllegalCharacters(CustomFieldGetName(lngCF))
      Else
        GoTo next_cf
      End If
      On Error Resume Next
      If Len(CustomFieldValueListGetItem(lngCF, pjValueListValue, 1)) = 0 Then GoTo next_cf
      For lngListItem = 1 To 1000 'capped at 1000, hopefully that's enough...
        strValue = CustomFieldValueListGetItem(lngCF, pjValueListValue, lngListItem)
        strDescription = CustomFieldValueListGetItem(lngCF, pjValueListDescription, lngListItem)
        If strValue <> "" Then
          oCodes.Add strValue, strDescription
        Else
          Exit For
        End If
        strValue = ""
        strDescription = ""
      Next lngListItem
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oCodes.Count > 0 Then
        lngFile = FreeFile
        strFileName = Environ("tmp") & "\" & Replace(strCFN, " ", "_") & ".csv"
        Open strFileName For Output As #lngFile
        Print #lngFile, "CODE,DESCRIPTION"
        For lngListItem = 0 To oCodes.Count - 1
          Print #lngFile, oCodes.Keys(lngListItem) & "," & Chr(34) & oCodes.Items(lngListItem) & Chr(34)
        Next lngListItem
        Close #lngFile
        cptShellExecute 0, "open", "notepad.exe", strFileName, vbNullString, 1
        lngCodes = lngCodes + 1
      End If
      oCodes.RemoveAll
next_cf:
    Next lngItem
  Next vFieldType
  
  Application.StatusBar = lngCodes & " codes exported."
  MsgBox lngCodes & " codes exported.", vbInformation + vbOKOnly, "Code Export"

exit_here:
  On Error Resume Next
  Set oFieldCounts = Nothing
  Reset
  Application.StatusBar = ""
  cptSpeed False
  Set oCodes = Nothing
  Set oOutlineCode = Nothing
  Set oLookupTable = Nothing
  Set oLookupTableEntry = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptExportAllCodes", Err, Erl)
  Resume exit_here
End Sub
