Attribute VB_Name = "cptDataDictionary_bas"
'<cpt_version>v1.5.1</cpt_version>
Option Explicit

Sub cptExportDataDictionary(ByRef myDataDictionary_frm As cptDataDictionary_frm)
  'objects
  Dim wsLookups As Excel.Worksheet
  Dim dFields As Scripting.Dictionary
  Dim oLookupTable As MSProject.LookupTable
  Dim rstDictionary As ADODB.Recordset
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oRange As Excel.Range
  'strings
  Dim strDir As String
  Dim strProject As String
  Dim strFullName As String
  Dim strDescription As String
  Dim strValue As String
  Dim strAttributes As String
  Dim strFieldName As String
  'longs
  Dim lngLookupCol As Long
  Dim lngItem As Long
  Dim lngItems As Long
  Dim lngListItem As Long
  Dim lngListItems As Long
  Dim lngCol As Long
  Dim lngHeaderRow As Long
  Dim lngRow  As Long
  Dim lngField As Long
  'integers
  Dim intField As Integer
  'doubles
  'booleans
  Dim blnHasFormula As Boolean
  Dim blnHasPickList As Boolean
  Dim blnLookups As Boolean
  Dim blnExists As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vArray() As Variant
  Dim vColumns As Variant
  Dim vFieldType As Variant
  Dim vFieldScope As Variant
  'dates
  
  'prevent spawning
  If Not cptGetUserForm("cptDataDictionary_frm") Is Nothing Then Exit Sub

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  'ensure project name
  strProject = cptGetProgramAcronym
  
  blnLookups = MsgBox("Replicate Pick Lists in Excel?", vbQuestion + vbYesNo, "Data Dictionary") = vbYes
  
  'set up a workbook/worksheet
  myDataDictionary_frm.lblStatus.Caption = "Creating Excel Workbook..."
  Set oExcel = CreateObject("Excel.Application")
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Worksheets(1)
  oWorksheet.Name = "Data Dictionary"
  
  'create lookups worksheet (optional)
  If blnLookups Then
    Set wsLookups = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(1))
    wsLookups.Name = "LOOKUPS"
    wsLookups.Activate
    oExcel.ActiveWindow.Zoom = 83
    oWorksheet.Activate
  End If
  
  'IMS Title
  oWorksheet.[A1].Value = "IMS Data Dictionary"
  oWorksheet.[A1].Font.Size = 18
  oWorksheet.[A1].Font.Bold = True
  'subtitle
  oWorksheet.[A2].Value = ActiveProject.Name
  oWorksheet.[A2].Font.Size = 14
  oWorksheet.[A2].Font.Bold = True
  'date
  oWorksheet.[A3].Value = FormatDateTime(Now, vbLongDate)
  
  'set header row
  lngHeaderRow = 5
  
  'set up columns
  vColumns = Array("Enterprise", "Scope", "Type", "Field", "Custom Name", "Attributes", "Description")
  oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, 1), oWorksheet.Cells(lngHeaderRow, 1).Offset(0, UBound(vColumns))) = vColumns
  
  'freezepanes
  oWorksheet.Cells(lngHeaderRow + 1, 1).Select
  oExcel.ActiveWindow.FreezePanes = True
  oExcel.ActiveWindow.Zoom = 85

  myDataDictionary_frm.lblStatus.Caption = "Exporting local custom fields..."
  
  blnExists = Dir(strDir & "\settings\cpt-data-dictionary.adtg") <> vbNullString

  If blnExists Then
    Set rstDictionary = CreateObject("ADODB.Recordset")
    rstDictionary.Open strDir & "\settings\cpt-data-dictionary.adtg"
  End If
  
  'count of custom fields = local + enterprise
  lngItems = 260 + (188778000 - 188776000)
  
  Set dFields = CreateObject("Scripting.Dictionary")
  dFields.Add "Cost", 10
  dFields.Add "Date", 10
  dFields.Add "Duration", 10
  dFields.Add "Flag", 20
  dFields.Add "Finish", 10
  dFields.Add "Outline Code", 10
  dFields.Add "Number", 20
  dFields.Add "Start", 10
  dFields.Add "Text", 30
  
  'prep for data dump
  lngRow = lngHeaderRow
  'export local custom fields
  For Each vFieldScope In Array(0, 1) '0 = pjTask; 1 = pjResource; 2 = pjProject
    For Each vFieldType In Array("Cost", "Date", "Duration", "Flag", "Finish", "Outline Code", "Number", "Start", "Text")
      For intField = 1 To dFields.Item(vFieldType)
        lngField = FieldNameToFieldConstant(vFieldType & intField, vFieldScope)
        If blnExists Then 'check if field is ignored
          rstDictionary.Filter = "PROJECT_NAME='" & strProject & "' AND FIELD_ID=" & lngField
          If Not rstDictionary.EOF Then
            If rstDictionary.Fields("IGNORE") Then
              rstDictionary.Filter = 0
              GoTo next_field
            End If
          End If
        End If
        strFieldName = CustomFieldGetName(lngField)
        
        blnHasFormula = Len(CustomFieldGetFormula(lngField)) > 0
        blnHasPickList = False 'default/reset
        If Not blnHasFormula Then 'does it have a picklist?
          On Error Resume Next
          blnHasPickList = Len(CStr(CustomFieldValueListGetItem(lngField, pjValueListValue, 1))) > 0
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        End If
        'only export if has custom field name OR a formula OR a picklist
        If Len(strFieldName) > 0 Or blnHasFormula Or blnHasPickList Then
          lngRow = lngRow + 1
          oWorksheet.Cells(lngRow, 1).Value = False
          oWorksheet.Cells(lngRow, 2).Value = Choose(CInt(vFieldScope) + 1, "Task", "Resource", "Project")
          oWorksheet.Cells(lngRow, 3).Value = CStr(vFieldType)
          oWorksheet.Cells(lngRow, 4).Value = FieldConstantToFieldName(lngField)
          oWorksheet.Cells(lngRow, 5).Value = strFieldName
          'get attributes
          strAttributes = ""
          If blnHasFormula Then
            oWorksheet.Cells(lngRow, 6).Value = CustomFieldGetFormula(lngField)
          ElseIf blnHasPickList Then
            If blnLookups Then
              lngLookupCol = wsLookups.[XFD2].End(-4159).Column
              If wsLookups.Cells(1, lngLookupCol) <> "" Then lngLookupCol = lngLookupCol + 2
              wsLookups.Cells(1, lngLookupCol) = UCase(strFieldName)
              wsLookups.Cells(2, lngLookupCol) = UCase(strFieldName) & " LOOKUP:"
              'outline codes
              If vFieldType = "Outline Code" Then
                Set oLookupTable = ActiveProject.OutlineCodes(CustomFieldGetName(lngField)).LookupTable
                lngListItems = oLookupTable.Count
                For lngListItem = 0 To oLookupTable.Count - 1
                  ReDim Preserve vArray(0 To lngListItem)
                  strDescription = ""
                  strFullName = ""
                  strDescription = oLookupTable(lngListItem + 1).Description
                  strFullName = oLookupTable(lngListItem + 1).FullName
                  If Len(strDescription) > 0 Then
                    If Left(strDescription, Len(strFullName)) = strFullName Then
                      vArray(lngListItem) = strDescription
                    Else
                      vArray(lngListItem) = strFullName & " - " & strDescription
                    End If
                  Else
                    vArray(lngListItem) = strFullName
                  End If
                  If lngListItem > 50 Then
                    myDataDictionary_frm.lboCustomFields.Value = lngField
                    myDataDictionary_frm.lblStatus.Caption = strFieldName & ": " & Format(lngListItem + 1, "#,##0") & "/" & Format(lngListItems, "#,##0") & " (" & Format((lngListItem + 1) / lngListItems, "0%") & ")"
                    myDataDictionary_frm.lblProgress.Width = ((lngListItem + 1) / lngListItems) * myDataDictionary_frm.lblStatus.Width
                    DoEvents
                  End If
                Next lngListItem
                If blnLookups And UBound(vArray) > 0 Then
                  wsLookups.Range(wsLookups.Cells(3, lngLookupCol), wsLookups.Cells(2 + lngListItem, lngLookupCol)) = WorksheetFunction.Transpose(vArray)
                End If
              Else
                'pick lists
                lngListItems = 1000
                For lngListItem = 0 To 999
                  ReDim Preserve vArray(0 To lngListItem)
                  strValue = ""
                  On Error Resume Next
                  strValue = CStr(CustomFieldValueListGetItem(lngField, pjValueListValue, lngListItem + 1))
                  If Err.Number = 1101 Then Exit For
                  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
                  strDescription = ""
                  strDescription = CustomFieldValueListGetItem(lngField, pjValueListDescription, lngListItem + 1)
                  If Len(strDescription) > 0 Then
                    vArray(lngListItem) = strDescription
                  Else
                    vArray(lngListItem) = strValue
                  End If
                  If lngListItem > 50 Then
                    myDataDictionary_frm.lboCustomFields.Value = lngField
                    myDataDictionary_frm.lblStatus.Caption = strFieldName & ": " & Format(lngListItem + 1, "#,##0") & "/<=" & Format(lngListItems, "#,##0") & "? (" & Format((lngListItem + 1) / lngListItems, "0%") & ")"
                    myDataDictionary_frm.lblProgress.Width = ((lngListItem + 1) / lngListItems) * myDataDictionary_frm.lblStatus.Width
                    DoEvents
                  End If
                Next lngListItem
                If UBound(vArray) > 0 Then
                  wsLookups.Range(wsLookups.Cells(3, lngLookupCol), wsLookups.Cells(2 + lngListItem, lngLookupCol)) = WorksheetFunction.Transpose(vArray)
                End If
              End If
            
              'name the range
              wsLookups.ListObjects.Add(SourceType:=1, Source:=wsLookups.Range(wsLookups.Cells(1, lngLookupCol), wsLookups.Cells(2 + lngListItem, lngLookupCol)), xlListObjectHasHeaders:=1).Name = UCase(Replace(Replace(FieldConstantToFieldName(lngField), " ", "_"), "-", "_"))
              wsLookups.Columns(lngLookupCol).AutoFit
              wsLookups.Columns(lngLookupCol + 1).ColumnWidth = 2
              'todo: how to keep pick list with cell when ListObject gets sorted?
              With oWorksheet.Cells(lngRow, 6).Validation
                 .Delete
                 .Add Type:=3, AlertStyle:=1, Operator:= _
                 1, Formula1:="=INDIRECT(""" & UCase(Replace(Replace(FieldConstantToFieldName(lngField), " ", "_"), "-", "_")) & """)"
                 .IgnoreBlank = True
                 .InCellDropdown = True
                 .InputTitle = ""
                 .ErrorTitle = ""
                 .InputMessage = ""
                 .ErrorMessage = ""
                 .ShowInput = True
                 .ShowError = True
               End With
               oWorksheet.Cells(lngRow, 6).Value = UCase(strFieldName) & " LOOKUP:"
            Else 'don't
              oWorksheet.Cells(lngRow, 6).Value = "[HAS PICKLIST/LOOKUP]"
            End If 'blnLookups
            
          End If 'Not LookupTable Is Nothing Then
          
        End If 'Len(strFieldName) > 0

        If blnExists Then
          rstDictionary.Filter = "PROJECT_NAME='" & strProject & "' AND FIELD_ID=" & lngField
          If Not rstDictionary.EOF Then oWorksheet.Cells(lngRow, 7).Value = rstDictionary("DESCRIPTION")
          rstDictionary.Filter = 0
        End If
        
next_field:
        Erase vArray
        lngItem = lngItem + 1
        myDataDictionary_frm.lboCustomFields.ListIndex = 0
        myDataDictionary_frm.lboCustomFields.Value = Null
        myDataDictionary_frm.lblStatus.Caption = "Exporting Local Custom Fields..." & lngItem & "/" & lngItems & " (" & Format(lngItem / lngItems, "0%") & ")"
        myDataDictionary_frm.lblProgress.Width = (lngItem / lngItems) * myDataDictionary_frm.lblStatus.Width
      Next intField
    Next vFieldType
  Next vFieldScope
  
  myDataDictionary_frm.lblStatus.Caption = "Exporting Enterprise Custom Fields..."
  
  'get enterprise custom fields
  For lngField = 188776000 To 188778000
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      lngRow = lngRow + 1
      oWorksheet.Cells(lngRow, 1).Value = True
      oWorksheet.Cells(lngRow, 2).Value = "Enterprise"
      oWorksheet.Cells(lngRow, 3).Value = "Enterprise"
      oWorksheet.Cells(lngRow, 4).Value = FieldConstantToFieldName(lngField)
      oWorksheet.Cells(lngRow, 5).Value = FieldConstantToFieldName(lngField)
      If Not Application.IsOffline Then
        If Len(CustomFieldGetFormula(lngField)) > 0 Then
          oWorksheet.Cells(lngRow, 6).Value = CustomFieldGetFormula(lngField)
        Else
          strAttributes = ""
          Set oLookupTable = Nothing
          On Error Resume Next
          Set oLookupTable = GlobalOutlineCodes(FieldConstantToFieldName(lngField)).LookupTable
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If Not oLookupTable Is Nothing Then
            If blnLookups Then
              lngLookupCol = wsLookups.[XFD2].End(-4159).Column
              If wsLookups.Cells(1, lngLookupCol) <> "" Then lngLookupCol = lngLookupCol + 2
              wsLookups.Cells(1, lngLookupCol) = UCase(FieldConstantToFieldName(lngField))
              wsLookups.Cells(2, lngLookupCol) = UCase(FieldConstantToFieldName(lngField)) & " LOOKUP:"
            End If
            For lngListItem = 1 To oLookupTable.Count
              If Len(oLookupTable(lngListItem).Description) > 0 Then
                If Left(oLookupTable(lngListItem).Description, Len(oLookupTable(lngListItem).FullName)) = oLookupTable(lngListItem).FullName Then
                  strAttributes = strAttributes & vbCrLf & oLookupTable(lngListItem).Description
                Else
                  strAttributes = strAttributes & vbCrLf & oLookupTable(lngListItem).FullName & " - " & oLookupTable(lngListItem).Description
                End If
              Else
                strAttributes = strAttributes & vbCrLf & oLookupTable(lngListItem).FullName
              End If
            Next lngListItem
            If blnLookups And Len(strAttributes) > 0 Then
              wsLookups.Range(wsLookups.Cells(3, lngLookupCol), wsLookups.Cells(2 + lngListItem, lngLookupCol)) = WorksheetFunction.Transpose(Split(Right(strAttributes, Len(strAttributes) - 2), vbCrLf))
            End If

            If blnLookups Then 'use validation
              'name the range
              wsLookups.ListObjects.Add(SourceType:=1, Source:=wsLookups.Range(wsLookups.Cells(1, lngLookupCol), wsLookups.Cells(2 + oLookupTable.Count, lngLookupCol)), xlListObjectHasHeaders:=1).Name = UCase(Replace(Replace(FieldConstantToFieldName(lngField), " ", "_"), "-", "_"))
              wsLookups.Columns(lngLookupCol).AutoFit
              wsLookups.Columns(lngLookupCol + 1).ColumnWidth = 2
              With oWorksheet.Cells(lngRow, 6).Validation
                .Delete
                .Add Type:=3, AlertStyle:=1, Operator:= _
                1, Formula1:="=INDIRECT(""" & UCase(Replace(Replace(FieldConstantToFieldName(lngField), " ", "_"), "-", "_")) & """)"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = True
              End With
              oWorksheet.Cells(lngRow, 6).Value = UCase(FieldConstantToFieldName(lngField)) & " LOOKUP:"
            Else
              If Len(strAttributes) > 0 Then oWorksheet.Cells(lngRow, 6).Value = "Lookup Values:" & strAttributes
            End If 'blnLookups
          End If 'has formula
        End If 'Not LookupTable Is Nothing Then
      End If 'is offline

      If blnExists Then
        rstDictionary.Filter = "PROJECT_NAME='" & strProject & "' AND FIELD_ID=" & lngField
        If Not rstDictionary.EOF Then oWorksheet.Cells(lngRow, 7).Value = rstDictionary("DESCRIPTION")
        rstDictionary.Filter = ""
      End If

    End If
    lngItem = lngItem + 1
    myDataDictionary_frm.lblStatus.Caption = "Exporting Enterprise Custom Fields..." & lngItem & "/" & lngItems & " (" & Format(lngItem / lngItems, "0%") & ")"
    myDataDictionary_frm.lblProgress.Width = (lngItem / lngItems) * myDataDictionary_frm.lblStatus.Width
  Next lngField
    
  myDataDictionary_frm.lblStatus.Caption = "Formatting..."
  
  'make it nice
  If blnLookups Then
    wsLookups.Activate
    oExcel.ActiveWindow.Zoom = 85
    wsLookups.[A2].Select
    oExcel.ActiveWindow.FreezePanes = True
    wsLookups.[A3].Select
    wsLookups.Rows(2).Hidden = True
  End If
  
  'convert to table / format it
  oWorksheet.Activate
  oExcel.ActiveWindow.ScrollRow = lngHeaderRow
  Set oRange = oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, 1).End(-4161), oWorksheet.Cells(lngHeaderRow, 1).End(-4121))
  oWorksheet.ListObjects.Add(1, oRange, , 1).Name = "DATA_DICTIONARY"
  'autofit
  oWorksheet.Range("DATA_DICTIONARY[#All]").Select
  oRange.Columns.AutoFit
  oRange.Rows.AutoFit
  oRange.VerticalAlignment = xlCenter
  lngCol = oWorksheet.Rows(lngHeaderRow).Find("Attributes", lookat:=1).Column
  oWorksheet.Columns(lngCol).ColumnWidth = 100
  oWorksheet.Columns(lngCol).WrapText = True
  lngCol = oWorksheet.Rows(lngHeaderRow).Find("Description", lookat:=1).Column
  oWorksheet.Columns(lngCol).ColumnWidth = 100
  oWorksheet.Cells(lngHeaderRow + 1, 1).Select
  
  myDataDictionary_frm.lblStatus.Caption = "Opening..."
  
exit_here:
  On Error Resume Next
  Set wsLookups = Nothing
  If rstDictionary.State Then rstDictionary.Close
  Set rstDictionary = Nothing
  Set oLookupTable = Nothing
  myDataDictionary_frm.lblStatus.Caption = "Ready..."
  If Not oExcel Is Nothing Then oExcel.Visible = True
  Set oRange = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
  
err_here:
  If Err.Number = 1101 Or Err.Number = 1004 Then
    Err.Clear
    Resume next_field
  Else
    Call cptHandleErr("cptExportCustomFields_bas", "cptExportDataDictionary", Err, Erl)
  End If
  
End Sub

Sub cptShowDataDictionary_frm()
  'objects
  Dim myDataDictionary_frm As cptDataDictionary_frm
  Dim cn As ADODB.Recordset
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  Dim rst As ADODB.Recordset 'Object
  'strings
  Dim strDir As String
  Dim strProgram As String
  Dim strMsg As String
  Dim strFileName As String
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  strProgram = cptGetProgramAcronym
  
  'update legacy adtg for ProjectName field
  'todo: ensure IMPORT Feature Supports PROJECT_NAME, but allows for legacy files as well
  'todo: should there *still* be a 'recover' option? maybe repeat this procedure also if there are blanks in the PROJECT_NAME?
  'todo: to remove an entry, either put 'REMOVE' in the PROJECT_NAME, or delete the worksheet row
  strFileName = strDir & "\settings\cpt-data-dictionary.adtg"
  'does the file exist?
  If Dir(strFileName) <> vbNullString Then
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open strFileName
    'does it have any records?
    If rst.RecordCount > 0 Then
      rst.MoveFirst
      'has it been upgraded yet?
      If rst.Fields.Count < 6 Then
        'add new field
        'rst.Fields.Append "ProjectName", adVarChar, 255
        'rst.Update
        'prompt the user
        strMsg = "The structure of the saved data for this feature has been upgraded to rely on a user-defined PROJECT_NAME instead of a GUID. " & vbCrLf & vbCrLf
        strMsg = strMsg & "Existing entries in your saved data must be updated accordingly for this feature to work as expected. Please update each row of the workbook you're about to see with a unique PROJECT_NAME for your Project, save it, keep it open, then Import that workbook. (Grab a screen shot of these instructions if you'd like.)" & vbCrLf & vbCrLf
        strMsg = strMsg & "Note: copying an .mpp file changes its GUID and thus some data dictionary entries may have been 'orphaned' - however, all of your saved entries will now be in the imminent workbook and may be recovered."
        MsgBox strMsg, vbInformation + vbOKOnly, "Action Required"
        'create the workbook
        On Error Resume Next
        Set oExcel = GetObject(, "Excel.Application")
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If oExcel Is Nothing Then Set oExcel = CreateObject("Excel.Application")
        oExcel.Visible = True
        Set oWorkbook = oExcel.Workbooks.Add
        'todo: handle if cpt-data-dictionary-reset.xlsx already exists
        oWorkbook.SaveAs strDir & "\settings\cpt-data-dictionary-reset.xlsx", 51
        Set oWorksheet = oWorkbook.Sheets(1)
        oWorksheet.Name = "Data Dictionary"
        'dump the columns
        For lngField = 0 To rst.Fields.Count - 1
          oWorksheet.Cells(1, lngField + 1).Value = rst.Fields(lngField).Name
        Next lngField
        oWorksheet.Cells(2, 1).CopyFromRecordset rst
        rst.Close
        oWorksheet.Columns(2).Insert
        oWorksheet.Cells(1, 2).Value = "PROJECT_NAME"
        oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True
        oWorksheet.[A1].AutoFilter
        With oExcel.ActiveWindow
          .Zoom = 85
          .SplitRow = 1
          .SplitColumn = 0
          .FreezePanes = True
        End With
        oWorksheet.Columns.AutoFit
      End If
    End If
  End If
  
  Set myDataDictionary_frm = New cptDataDictionary_frm
  With myDataDictionary_frm
    .lboCustomFields.Clear
    cptRefreshDictionary myDataDictionary_frm
    .txtFilter.SetFocus
    .Caption = "IMS Data Dictionary (" & cptGetVersion("cptDataDictionary_frm") & ")"
    .txtDescription.Enabled = False
    .chkIgnore.Enabled = False
    .Show
  End With
  
exit_here:
  On Error Resume Next
  Unload myDataDictionary_frm
  Set myDataDictionary_frm = Nothing
  Set cn = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  If rst.State Then rst.Close
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "cptShowDataDictionary_frm()", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshDictionary(ByRef myDataDictionary_frm As cptDataDictionary_frm)
  'objects
  Dim dTypes As Object 'Scripting.Dictionary
  Dim rstSaved As Object 'ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strProject As String
  Dim strFieldName As String
  Dim strCustomName As String
  Dim strGUID As String
  'longs
  Dim lngFieldCount As Long
  Dim lngItem As Long
  Dim lngField As Long
  Dim lngMax As Long
  'integers
  Dim intField As Integer
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnHasPickList As Boolean
  Dim blnHasFormula As Boolean
  Dim blnCreate As Boolean
  Dim blnExists As Boolean
  'variants
  Dim vFieldType As Variant
  Dim vFieldScope As Variant
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  
  'clear the form if it's visible
  If myDataDictionary_frm.Visible Then
    With myDataDictionary_frm
      .lboCustomFields.Clear
      .txtFilter.Text = ""
      .txtDescription.Value = ""
      .txtDescription.Enabled = False
      .chkIgnore.Enabled = False
      .lblAlert.Caption = "-'"
      .lblAlert.Visible = False
      .lboCustomFields.Height = 128.25
    End With
  End If
  
  'ensure project acronym
  strProject = cptGetProgramAcronym
  
  'if data file exists then use it else create it
  Set rstSaved = CreateObject("ADODB.Recordset")
  With rstSaved
    If Dir(strDir & "\settings\cpt-data-dictionary.adtg") = vbNullString Then
      blnCreate = True
      .Fields.Append "PROJECT_ID", 200, 50 'adVarChar
      .Fields.Append "FIELD_ID", 3 'adInteger = Long
      .Fields.Append "FIELD_NAME", 200, 50
      .Fields.Append "CUSTOM_NAME", 200, 50
      .Fields.Append "DESCRIPTION", 203, 500 'adLongVarWChar
      .Fields.Append "PROJECT_NAME", 200, 50
      .Fields.Append "IGNORE", 11 'adBoolean
      .Open
    Else
      blnCreate = False
      .Open strDir & "\settings\cpt-data-dictionary.adtg"
      'todo: has it been upgraded yet?
      .Filter = "PROJECT_NAME='" & strProject & "'"
      'has it been upgraded with IGNORE?
      On Error Resume Next
      If Err.Number = 3265 Then 'it has not
        Err.Clear
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        .Filter = 0
        .Close
        cptAppendColumn strDir & "\settings\cpt-data-dictionary.adtg", "IGNORE", adBoolean, , 0
        .Open strDir & "\settings\cpt-data-dictionary.adtg"
        .Filter = "PROJECT_NAME='" & strProject & "'"
      End If
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    End If
    
    'get local custom fields
    'export local custom fields
    
    Set dTypes = CreateObject("Scripting.Dictionary")
    dTypes.Add "Cost", 10
    dTypes.Add "Date", 10
    dTypes.Add "Duration", 10
    dTypes.Add "Flag", 20
    dTypes.Add "Finish", 10
    dTypes.Add "Number", 20
    dTypes.Add "Start", 10
    dTypes.Add "Text", 30
    dTypes.Add "Outline Code", 10
        
    For Each vFieldScope In Array(0, 1) '0 = pjTask; 1 = pjResource; 2 = pjProject
      For Each vFieldType In Array("Cost", "Date", "Duration", "Flag", "Finish", "Number", "Start", "Text", "Outline Code")
        For intField = 1 To dTypes(vFieldType) 'lngMax
          lngField = FieldNameToFieldConstant(vFieldType & intField, vFieldScope)
          strFieldName = FieldConstantToFieldName(lngField)
          strCustomName = CustomFieldGetName(lngField)
          If Len(strCustomName) = 0 Then 'include unnamed but has formula or pick list 'todo: what about has data?
            If Len(CustomFieldGetFormula(lngField)) > 0 Then blnHasFormula = True Else blnHasFormula = False
            blnHasPickList = False
            On Error Resume Next
            blnHasPickList = Len(CustomFieldValueListGetItem(lngField, pjValueListValue, 1)) > 0
            If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          End If
          If Len(strCustomName) > 0 Or blnHasFormula Or blnHasPickList Then
            If blnCreate Then
              'add to data store
              .AddNew Array("PROJECT_ID", "FIELD_ID", "FIELD_NAME", "CUSTOM_NAME", "DESCRIPTION", "PROJECT_NAME", "IGNORE"), Array(strGUID, lngField, strFieldName, strCustomName, "<missing>", strProject, CBool(0))
            Else
              'does it exist?
              .Filter = "PROJECT_NAME='" & strProject & "' AND FIELD_ID=" & CLng(lngField)
              'todo: update CUSTOM_NAME if changed; wipe out old definition?
              'if not then add it
              If .EOF Then
                .AddNew Array("PROJECT_ID", "FIELD_ID", "FIELD_NAME", "CUSTOM_NAME", "DESCRIPTION", "PROJECT_NAME", "IGNORE"), Array(strGUID, lngField, strFieldName, strCustomName, "<missing>", strProject, CBool(0))
              Else
                .Fields("CUSTOM_NAME") = strCustomName
                .Update
              End If
              .Filter = ""
            End If
          Else
            .Filter = "PROJECT_NAME='" & strProject & "' AND FIELD_ID=" & CLng(lngField)
            If Not .EOF Then .Delete adAffectCurrent
            .Filter = 0
          End If
          lngFieldCount = lngFieldCount + 1
          myDataDictionary_frm.lblStatus.Caption = "Refreshing...(" & Format(lngFieldCount / 2261, "0%") & ")"
          myDataDictionary_frm.lblProgress.Width = (lngFieldCount / 2261) * myDataDictionary_frm.lblStatus.Width
          DoEvents
        Next intField
      Next vFieldType
    Next vFieldScope
    
    'get enterprise custom fields
    For lngField = 188776000 To 188778000
      If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
        strFieldName = "Enterprise"
        strCustomName = FieldConstantToFieldName(lngField)
        If blnCreate Then
          .AddNew Array("PROJECT_ID", "FIELD_ID", "FIELD_NAME", "CUSTOM_NAME", "DESCRIPTION", "PROJECT_NAME", "IGNORE"), Array(strGUID, lngField, strFieldName, strCustomName, "<missing>", strProject, CBool(0))
        Else
          'does it exist?
          .Filter = "PROJECT_NAME='" & strProject & "' AND FIELD_ID=" & lngField
          'todo: update CUSTOM_NAME if changed; wipe out old definition?
          'if not, then add it
          If .EOF Then
            .AddNew Array("PROJECT_ID", "FIELD_ID", "FIELD_NAME", "CUSTOM_NAME", "DESCRIPTION", "PROJECT_NAME", "IGNORE"), Array(strGUID, lngField, strFieldName, strCustomName, "<missing>", strProject, CBool(0))
          End If
          .Filter = ""
        End If
      End If
      lngFieldCount = lngFieldCount + 1
      myDataDictionary_frm.lblStatus.Caption = "Refreshing...(" & Format(lngFieldCount / 2261, "0%") & ")"
      myDataDictionary_frm.lblProgress.Width = (lngFieldCount / 2261) * myDataDictionary_frm.lblStatus.Width
      DoEvents
    Next lngField
    'todo: remove PROJECT_ID
    
    'save the data
    .Filter = 0
    .Save strDir & "\settings\cpt-data-dictionary.adtg", 0 '0=adPersistADTG
    
    'populate the list
    If Not .EOF Then
      .Filter = "PROJECT_NAME='" & strProject & "'"
      If .RecordCount = 0 Then
        .Filter = 0
        GoTo exit_here
      End If
      .Sort = "CUSTOM_NAME"
      .MoveFirst
      lngItem = 0
      Do While Not .EOF
        myDataDictionary_frm.lboCustomFields.AddItem
        myDataDictionary_frm.lboCustomFields.List(lngItem, 0) = .Fields("FIELD_ID")
        If .Fields("FIELD_ID") >= 188776000 Then
          myDataDictionary_frm.lboCustomFields.List(lngItem, 1) = .Fields("CUSTOM_NAME") & " (Enterprise)"
        Else
          myDataDictionary_frm.lboCustomFields.List(lngItem, 1) = .Fields("CUSTOM_NAME") & " (" & .Fields("FIELD_NAME") & ")"
        End If
        If Len(.Fields("CUSTOM_NAME")) = 0 Then
          If Len(CustomFieldGetFormula(.Fields("FIELD_ID"))) > 0 Then
            myDataDictionary_frm.lboCustomFields.List(lngItem, 4) = "f"
          End If
          blnHasPickList = False
          On Error Resume Next
          blnHasPickList = Len(CustomFieldValueListGetItem(lngField, pjValueListValue, 1)) > 0
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        End If
        myDataDictionary_frm.lboCustomFields.List(lngItem, 2) = .Fields("DESCRIPTION")
        myDataDictionary_frm.lboCustomFields.List(lngItem, 3) = CBool(.Fields("IGNORE"))
        .MoveNext
        lngItem = lngItem + 1
      Loop
      .Close
    End If
  End With
  
exit_here:
  On Error Resume Next
  myDataDictionary_frm.lblStatus.Caption = "Ready..."
  Set dTypes = Nothing
  If rstSaved.State = 1 Then rstSaved.Close
  Set rstSaved = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "cptRefreshDictionary", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportDataDictionary(ByRef myDataDictionary_frm As cptDataDictionary_frm, Optional strFileName As String)
  'objects
  Dim rstSaved As ADODB.Recordset
  Dim oExcel As Object
  Dim oWorkbook As Object
  Dim oWorksheet As Object
  Dim oRange As Object
  'strings
  Dim strGUID As String
  Dim strDescription As String
  Dim strFieldName As String
  Dim strProject As String
  Dim strMsg As String
  Dim strNewDescription As String
  Dim strCustomName As String
  Dim strFieldsNotFound As String
  Dim strSavedSettings As String
  'longs
  Dim lngSkipped As Long
  Dim lngFieldID As Long
  Dim lngNotFound As Long
  Dim lngNew As Long
  Dim lngTooLong As Long
  Dim lngLastCol As Long
  Dim lngRow As Long
  Dim lngField As Long
  Dim lngLastRow As Long
  Dim lngDescriptionCol As Long
  Dim lngNameCol As Long
  Dim lngHeaderRow As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnSkip As Boolean
  Dim blnClose As Boolean
  'variants
  Dim vFile As Variant
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'prompt user to select a file
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
    blnClose = True
  Else
    'oExcel.WindowState = xlMaximized
    'Application.ActivateMicrosoftApp pjMicrosoftExcel
    blnClose = False
  End If
  If Len(strFileName) > 0 Then
    vFile = strFileName
    GoTo skip_that
  Else
    vFile = oExcel.GetOpenFilename(FileFilter:="Microsoft Excel *.xls* (*.xls*),", _
                                  Title:="Select a Populated IMS Data Dictionary:", _
                                  ButtonText:="Import", MultiSelect:=False)
    If vFile = False Then GoTo exit_here
  End If
  oExcel.ActivateMicrosoftApp 6 'xlMicrosoftProject
  
skip_that:
  
  'ensure a project acronym is defined
  strProject = cptGetProgramAcronym
  
  'validate the file by its headers
  On Error Resume Next
  oExcel.ScreenUpdating = False
  Set oWorkbook = oExcel.Workbooks(vFile)
  If oWorkbook Is Nothing Then Set oWorkbook = oExcel.Workbooks.Open(vFile)
  Set oWorksheet = oWorkbook.Sheets("Data Dictionary")
  If oWorksheet Is Nothing Then
    MsgBox strFileName & " does not appear to be a valid IMS Data Dictionary workbook. The wheet named 'Data Dictionary' not found.", vbExclamation + vbOKOnly, "Invalid Workbook"
    GoTo exit_here
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set rstSaved = CreateObject("ADODB.Recordset")
  strSavedSettings = cptDir & "\settings\cpt-data-dictionary.adtg"
  
  'handle a reset
  If oWorkbook.Name = "cpt-data-dictionary-reset.xlsx" Then
    lngHeaderRow = oWorksheet.Columns(1).Find("PROJECT_ID", lookat:=1).Row
    lngLastRow = oWorksheet.Cells(lngHeaderRow, 1).End(-4121).Row
    With rstSaved
      .Fields.Append "PROJECT_ID", 200, 50 'adVarChar
      .Fields.Append "FIELD_ID", 3 'adInteger = Long
      .Fields.Append "FIELD_NAME", 200, 50
      .Fields.Append "CUSTOM_NAME", 200, 50
      .Fields.Append "DESCRIPTION", 203, 500 'adLongVarWChar
      .Fields.Append "PROJECT_NAME", 200, 50
      .Open
      For lngRow = 2 To lngLastRow
        blnSkip = False
        strGUID = oWorksheet.Cells(lngRow, 1).Value
        strProject = oWorksheet.Cells(lngRow, 2).Value
        If Len(strProject) = 0 Then
          oWorksheet.Cells(lngRow, 2).Style = "Bad"
          blnSkip = True
        Else
          If oWorksheet.Cells(lngRow, 2).Style = "Bad" Then
            oWorksheet.Cells(lngRow, 2).Style = "Good"
          End If
        End If
        lngFieldID = CLng(oWorksheet.Cells(lngRow, 3).Value)
        strFieldName = oWorksheet.Cells(lngRow, 4).Value
        strCustomName = oWorksheet.Cells(lngRow, 5).Value
        strDescription = oWorksheet.Cells(lngRow, 6).Value
        If blnSkip Then
          lngSkipped = lngSkipped + 1
        Else
          .AddNew Array(0, 1, 2, 3, 4, 5), Array(strGUID, lngFieldID, strFieldName, strCustomName, strDescription, strProject)
        End If
        With myDataDictionary_frm
          .lblStatus.Caption = "Resetting...(" & Format((lngRow - 1) / (lngLastRow - 1), "0%") & ")"
          .lblProgress.Width = .lblStatus.Width * ((lngRow - 1) / (lngLastRow - 1))
          DoEvents
        End With
      Next lngRow
      If lngSkipped > 0 Then
        MsgBox Format(lngSkipped, "#,##0") & " rows had missing project names and were skipped and marked red. Correct the file and import again, if desired.", vbExclamation + vbOKOnly, "Records Skipped"
      End If
      If Dir(strSavedSettings) <> vbNullString Then Kill strSavedSettings
      .Save strSavedSettings, adPersistADTG
      .Close
      myDataDictionary_frm.lblStatus.Caption = "Data Dictionary Reset."
      cptRefreshDictionary myDataDictionary_frm
      GoTo exit_here
    End With
  End If
  
  lngHeaderRow = oWorksheet.Columns(1).Find("Enterprise", lookat:=1).Row
  lngLastRow = oWorksheet.Cells(lngHeaderRow, 1).End(-4121).Row
  lngNameCol = oWorksheet.Rows(lngHeaderRow).Find("Custom Name", lookat:=1).Column
  lngLastCol = oWorksheet.Rows(lngHeaderRow).End(-4161).Column
  lngDescriptionCol = oWorksheet.Rows(lngHeaderRow).Find("Description", lookat:=1).Column
  
  'get saved dictionary settings
  If Dir(strSavedSettings) = vbNullString Then Call cptRefreshDictionary(myDataDictionary_frm)
  
  With rstSaved
    .Open strSavedSettings
    .Filter = "PROJECT_NAME='" & strProject & "'"
    .MoveFirst
    For lngRow = lngHeaderRow + 1 To lngLastRow
      'reset validation on description cell
      oWorksheet.Cells(lngRow, lngDescriptionCol).Style = "Normal"
      oWorksheet.Cells(lngRow, lngDescriptionCol).HorizontalAlignment = xlCenter
      strCustomName = oWorksheet.Cells(lngRow, lngNameCol).Value
      strNewDescription = oWorksheet.Cells(lngRow, lngDescriptionCol).Value
      If Len(oWorksheet.Cells(lngRow, lngDescriptionCol).Value) > 500 Then
        lngTooLong = lngTooLong + 1
        oWorksheet.Cells(lngRow, lngDescriptionCol).Style = "Neutral"
      End If
      If strNewDescription = "<missing>" Then GoTo next_row
      .Find "CUSTOM_NAME='" & strCustomName & "'", , 1 'adSearchForward
      If Not .EOF Then
        If rstSaved("DESCRIPTION").Value <> strNewDescription Then
          oWorksheet.Cells(lngRow, lngDescriptionCol).Style = "Good"
          lngNew = lngNew + 1
        End If
      Else
        oWorksheet.Cells(lngRow, lngDescriptionCol).Style = "Bad"
        oWorksheet.Cells(lngRow, lngLastCol + 1).Value = "NOT FOUND"
        oWorksheet.Cells(lngRow, lngLastCol + 1).Style = "Bad"
        lngNotFound = lngNotFound + 1
      End If
next_row:
    Next lngRow
    
    'show the selected file
    If Not oExcel.Visible Then oExcel.Visible = True
    oExcel.ScreenUpdating = True
    oExcel.WindowState = -4140 'xlMinimized
    
    'notify of missing custom fields
    If lngNotFound > 0 Then
      strMsg = Format(lngNotFound, "#,##0") & " fields were found in the selected Excel Workbook but were not found in this Project File." & vbCrLf
      strMsg = strMsg & "These have been marked with cell style 'Bad'." & vbCrLf
      strMsg = strMsg & "Please correct the Excel workbook before importing if necessary. Otherwise, click 'Yes' to proceed." & vbCrLf & vbCrLf
      strMsg = strMsg & "Do you wish to proceed?"
      If MsgBox(strMsg, vbExclamation + vbYesNo, "Missing Custom Fields!") = vbNo Then GoTo exit_here
    End If
    
    'notify if character limit is exceeded
    If lngTooLong > 0 Then
      strMsg = Format(lngTooLong, "#,##0") & " descriptions exceed the 500 character limit; these have been marked with cell style 'Neutral'." & vbCrLf
      strMsg = strMsg & "If you proceed without correcting them, they will be concatenanted (though your spreadsheet will remain unchanged." & vbCrLf & vbCrLf
      strMsg = strMsg & "Do you wish to proceed?"
      If MsgBox(strMsg, vbExclamation + vbYesNo, "Character Limitation Exceeded") = vbNo Then GoTo exit_here
    End If
    
    'confirm updates
    If lngNew > 0 Then
      strMsg = Format(lngNew, "#,##0") & " updated IMS Data Dictionary entries found." & vbCrLf
      strMsg = strMsg & "Please take a moment to indicate unwanted changes in the Excel workbook by marking the cell style as 'Bad.'" & vbCrLf
      strMsg = strMsg & "When new entries have been validated, click 'Yes' below, or hit 'No' to cancel this import." & vbCrLf & vbCrLf
      strMsg = strMsg & "Do you wish to proceed?"
    Else
      MsgBox "No updated entries found.", vbInformation + vbOKOnly, "Import Skipped"
      'Workbook.Close False
      GoTo exit_here
    End If
    
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Confirm Import") = vbNo Then
      'Workbook.Close True
      GoTo exit_here
    Else
      lngNew = 0 'reset the counter
      .MoveFirst
      For lngRow = lngHeaderRow + 1 To lngLastRow
        If oWorksheet.Cells(lngRow, lngDescriptionCol).Style <> "Bad" Then
          strCustomName = oWorksheet.Cells(lngRow, lngNameCol).Value
          strNewDescription = oWorksheet.Cells(lngRow, lngDescriptionCol).Value
          If strNewDescription = "<missing>" Then GoTo next_row2
          'todo: should this match constant AND custom name?
          .Find "CUSTOM_NAME='" & strCustomName & "'", , 1 'adSearchForward
          If Not .EOF Then
            If rstSaved("DESCRIPTION").Value <> strNewDescription Then
              If rstSaved("DESCRIPTION") <> "<missing>" Then
                oWorksheet.Cells(lngRow, lngLastCol + 1).Value = "DESCRIPTION WAS: " & rstSaved("DESCRIPTION").Value
              End If
              rstSaved("DESCRIPTION").Value = strNewDescription
              .Update
              lngNew = lngNew + 1
            End If
          End If
        Else
          oWorksheet.Cells(lngRow, lngLastCol + 1).Value = oWorksheet.Cells(lngRow, lngLastCol + 1).Value & " - SKIPPED"
        End If
next_row2:
      Next
    End If
    .Filter = 0
    .Save strSavedSettings, adPersistADTG
    .Close
  End With
  
  If lngNew > 0 Then
    cptRefreshDictionary myDataDictionary_frm
    MsgBox Format(lngNew, "#,##0") & " entries updated.", vbInformation + vbOKOnly, "Import Complete"
  Else
    MsgBox "No updates found.", vbInformation + vbOKOnly, "Import Skipped"
  End If
  
exit_here:
  On Error Resume Next
  myDataDictionary_frm.lblStatus.Caption = "Ready..."
  If rstSaved.State Then rstSaved.Close
  Set rstSaved = Nothing
  Set oRange = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  If Not oExcel.Visible Then oExcel.Visible = True
  oExcel.ScreenUpdating = True
  Set oExcel = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "cptImportDataDictionary", Err, Erl)
  Resume exit_here
End Sub

Function WhatTheActual() As String
  Dim dFieldCounts As Scripting.Dictionary
  Dim strResult As String
  Dim blnHasFormula As Boolean
  Dim blnHasPickList As Boolean
  Dim lngItem As Long
  Dim lngItems As Long
  Dim lngField As Long
  Dim vType As Variant
  
  Set dFieldCounts = CreateObject("Scripting.Dictionary")
  dFieldCounts.Add "Text", 30
  dFieldCounts.Add "Number", 20
  dFieldCounts.Add "Flag", 20
  
  For Each vType In Split("Cost,Date,Duration,Flag,Finish,Number,Outline Code,Start,Text", ",")
    If dFieldCounts.Exists(vType) Then
      lngItems = dFieldCounts(vType)
    Else
      lngItems = 10
    End If
    For lngItem = 1 To lngItems
      lngField = FieldNameToFieldConstant(vType & lngItem)
      blnHasFormula = False
      blnHasPickList = False
      On Error Resume Next
      blnHasFormula = Len(CustomFieldGetFormula(lngField)) > 0
      'If Not blnHasFormula Then
        blnHasPickList = Len(CustomFieldValueListGetItem(lngField, pjValueListValue, 1)) > 0
      'End If
      On Error GoTo 0
      strResult = strResult & lngField & "," & FieldConstantToFieldName(lngField) & "," & CustomFieldGetName(lngField) & "," & blnHasFormula & "," & blnHasPickList & vbCrLf
    Next lngItem
  Next vType
  
  WhatTheActual = strResult
  
  Set dFieldCounts = Nothing
  
End Function
