Attribute VB_Name = "cptCostRateTables_bas"
'<cpt_version>v1.2.0</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptCostRateTables_bas"

Sub cptShowCostRateTables_frm()
  'objects
  Dim myCostRateTables_frm As cptCostRateTables_frm
  'strings
  Dim strStatusField As String
  Dim strOverwrite As String
  Dim strAddNew As String
  Dim strCustomFieldName As String
  'longs
  Dim lngCustomField As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  'check for an update
  'If Not cptProceedOnUpdate(THIS_MODULE) Then GoTo exit_here
  
  'prevent spawning
  If Not cptGetUserForm("cptCostRateTables_frm") Is Nothing Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set myCostRateTables_frm = New cptCostRateTables_frm
  With myCostRateTables_frm
    .Caption = "Cost Rate Tables (" & cptGetVersion("cptCostRateTables_frm") & ")"
    .lblProgress.Width = .lblStatus.Width
    .lblStatus.Caption = "Ready..."
    With .cboStatusField
      .Clear
      For lngItem = 1 To 30
        lngCustomField = FieldNameToFieldConstant("Text" & lngItem, pjResource)
        strCustomFieldName = CustomFieldGetName(lngCustomField)
        If Len(strCustomFieldName) > 0 Then
          .AddItem
          .List(lngItem - 1, 0) = lngCustomField
          .List(lngItem - 1, 1) = "Text" & lngItem & " (" & strCustomFieldName & ")"
        Else
          .AddItem
          .List(lngItem - 1, 0) = lngCustomField
          .List(lngItem - 1, 1) = "Text" & lngItem
        End If
      Next lngItem
      .AddItem
      .List(.ListCount - 1, 0) = 0
      .List(.ListCount - 1, 1) = "TO CSV"
    End With
    If ActiveProject.ResourceCount > 0 Then
      .tglExport = True
    Else
      .tglImport = True
    End If
    strStatusField = cptGetSetting("CostRateTables", "cboStatusField")
    If Len(strStatusField) > 0 Then
      .cboStatusField.Value = CLng(strStatusField)
    End If
    strOverwrite = cptGetSetting("CostRateTables", "chkOverwrite")
    If Len(strOverwrite) > 0 Then
      .chkOverwrite = CBool(strOverwrite)
    Else
      .chkOverwrite = True 'default
    End If
    strAddNew = cptGetSetting("CostRateTables", "chkAddNew")
    If Len(strAddNew) > 0 Then
      .chkAddNew = CBool(strAddNew)
    Else
      .chkAddNew = True 'default
    End If
    .Show
  End With

exit_here:
  On Error Resume Next
  Unload myCostRateTables_frm
  
  Exit Sub
err_here:
  Call cptHandleErr("cptCostRateTables_bas", "cptShowCostRateTables_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCostRateTables(ByRef myCostRateTables_frm As cptCostRateTables_frm, strCostRateTables As String)
  'objects
  Dim oPayRate As PayRate
  Dim oCostRateTable As CostRateTable
  Dim oResource As Resource
  Dim oExcel As Object 'Excel.Application
  Dim oWorkbook As Object 'Excel.Workbook
  Dim oWorksheet As Object 'Excel.Worksheet
  'strings
  Dim strType As String
  Dim strRateTable As String
  Dim strResource As String
  'longs
  Dim lngCostRateTable As Long
  Dim lngLastRow As Long
  Dim lngResource As Long
  Dim lngResourceCount As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vCostRateTable As Variant
  'dates
  
  myCostRateTables_frm.lblStatus.Caption = "Getting Excel..."
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  Set oWorkbook = oExcel.Workbooks.Add
  oExcel.Calculation = xlCalculationManual
  oExcel.ScreenUpdating = False
  Set oWorksheet = oWorkbook.Sheets(1)
  myCostRateTables_frm.lblStatus.Caption = "Creating Header..."
  oWorksheet.[A1:G1] = Split(("RESOURCE,TYPE,RATE TABLE,EFFECTIVE DATE,STANDARD RATE,OVERTIME RATE,COST PER USE"), ",")
  
  lngResourceCount = ActiveProject.ResourceCount
  lngResource = 0
  For Each oResource In ActiveProject.Resources
    lngResource = lngResource + 1
    strResource = oResource.Name
    For Each vCostRateTable In Split(strCostRateTables, ",")
      If vCostRateTable = "" Then GoTo next_cost_rate_table
      lngCostRateTable = Switch(vCostRateTable = "A", 1, vCostRateTable = "B", 2, vCostRateTable = "C", 3, vCostRateTable = "D", 4, vCostRateTable = "E", 5)
      Set oCostRateTable = oResource.CostRateTables(lngCostRateTable)
      strType = Choose(oResource.Type + 1, "WORK", "MATERIAL", "COST")
      For Each oPayRate In oCostRateTable.PayRates
        lngLastRow = oWorksheet.[A1048576].End(-4162).Row + 1 '-4162 = xlUp
        oWorksheet.Cells(lngLastRow, 1) = strResource
        oWorksheet.Cells(lngLastRow, 2) = strType
        oWorksheet.Cells(lngLastRow, 3) = CStr(vCostRateTable)
        oWorksheet.Cells(lngLastRow, 4) = FormatDateTime(oPayRate.EffectiveDate, vbShortDate)
        oWorksheet.Cells(lngLastRow, 5) = oPayRate.StandardRate
        oWorksheet.Cells(lngLastRow, 6) = oPayRate.OvertimeRate
        oWorksheet.Cells(lngLastRow, 7) = oPayRate.CostPerUse
      Next oPayRate
next_cost_rate_table:
    Next vCostRateTable
    Application.StatusBar = Format(lngResource, "#,##0") & "/" & Format(lngResourceCount, "#,##0") & "...(" & Format(lngResource / lngResourceCount, "0%") & ")"
    myCostRateTables_frm.lblStatus.Caption = Format(lngResource, "#,##0") & "/" & Format(lngResourceCount, "#,##0") & "...(" & Format(lngResource / lngResourceCount, "0%") & ")"
    myCostRateTables_frm.lblProgress.Width = (lngResource / lngResourceCount) * myCostRateTables_frm.lblStatus.Width
    DoEvents
  Next oResource

  With myCostRateTables_frm
    .lblProgress.Width = .lblStatus.Width
    .lblStatus = "Complete."
  End With
  Application.StatusBar = "Complete."
  
  oExcel.Visible = True
  With oExcel.ActiveWindow
    .Zoom = 85
    .SplitRow = 1
    .SplitColumn = 0
    .FreezePanes = True
  End With
  oWorksheet.Columns.AutoFit

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oPayRate = Nothing
  Set oCostRateTable = Nothing
  Set oResource = Nothing
  Set oWorksheet = Nothing
  oExcel.Visible = True
  oExcel.ScreenUpdating = True
  oExcel.Calculation = xlCalculationAutomatic
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("basCostRateTables_bas", "cptExportCostRateTables", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportCostRateTables(ByRef myCostRateTables_frm As cptCostRateTables_frm, lngField As Long)
  'objects
  Dim oUpdated As Scripting.Dictionary
  Dim oPayRate As MSProject.PayRate
  Dim oCostRateTable As MSProject.CostRateTable
  Dim oResource As MSProject.Resource
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oComment As Excel.Comment
  'strings
  Dim strResourceName As String
  Dim strFileName As String
  Dim strOverwrite As String
  Dim strAddResources As String
  Dim strCostRateTable As String
  Dim strType As String
  Dim strWorkbook As String
  'longs
  Dim lngItem As Long
  Dim lngFile As Long
  Dim lngCostRateTable As Long
  Dim lngResourceNameCol As Long
  Dim lngResourceTypeCol As Long
  Dim lngRateTableCol As Long
  Dim lngEffectiveDateCol As Long
  Dim lngStandardRateCol As Long
  Dim lngOvertimeRateCol As Long
  Dim lngCostPerUseCol As Long
  Dim lngRow As Long
  Dim lngLastRow As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnImportStatus As Boolean
  Dim blnOverwrite As Boolean
  Dim blnAddResources As Boolean
  'variants
  Dim vCostPerUse As Variant
  Dim vOvtRate As Variant
  Dim vStdRate As Variant
  Dim vEffectiveDate As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Application.ActiveWindow.TopPane.Activate
  ViewApply "Resource Sheet"
  FilterClear
  
  'clear out the field
  Application.ScreenUpdating = True
  Application.Calculation = pjAutomatic
  If lngField > 0 Then
    myCostRateTables_frm.lblStatus.Caption = "Clearing " & FieldConstantToFieldName(lngField) & "..."
    ActiveWindow.TopPane.Activate
    FilterClear
    SetField FieldConstantToFieldName(lngField), ""
    DoEvents
  End If
  
  myCostRateTables_frm.lblStatus.Caption = "Getting Excel..."
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  With oExcel.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .ButtonName = "Import"
    .Title = "Import Cost Rate Tables:"
    .Filters.Clear
    .Filters.Add "Microsoft Excel", "*.xls*"
    .Show
    If .SelectedItems.Count > 0 Then
      strWorkbook = .SelectedItems(1)
    Else
      myCostRateTables_frm.lblStatus.Caption = "Ready..."
      GoTo exit_here
    End If
  End With
  
  'get / create saved settings
  strOverwrite = cptGetSetting("CostRateTables", "chkOverwrite")
  If Len(strOverwrite) > 0 Then
    blnOverwrite = CBool(strOverwrite)
  Else
    blnOverwrite = MsgBox("Overwrite existing Cost Rate Tables?", vbQuestion + vbYesNo, "Confirm Overwrite Cost Rate Tables") = vbYes
    cptSaveSetting "CostRateTables", "chkOverwrite", CBool(blnOverwrite)
  End If
  strAddResources = cptGetSetting("CostRateTables", "chkAddNew")
  If Len(strAddResources) > 0 Then
    blnAddResources = CBool(strAddResources)
  Else
    blnAddResources = MsgBox("Add Resources in Workbook but not in this project?", vbQuestion + vbYesNo, "Confirm Add New Resources") = vbYes
    cptSaveSetting "CostRateTables", "chkAddNew", CBool(blnAddResources)
  End If
  blnImportStatus = lngField > 0
  Set oUpdated = CreateObject("Scripting.Dictionary")
  
  Application.Calculation = pjManual
  Application.ScreenUpdating = False
  myCostRateTables_frm.lblStatus.Caption = "Opening Workbook..."
  Set oWorkbook = oExcel.Workbooks.Open(strWorkbook)
  Set oWorksheet = oWorkbook.Sheets(1)
  lngResourceNameCol = oWorksheet.Rows(1).Find("RESOURCE", lookat:=xlWhole).Column
  lngResourceTypeCol = oWorksheet.Rows(1).Find("TYPE", lookat:=xlWhole).Column
  lngRateTableCol = oWorksheet.Rows(1).Find("RATE TABLE", lookat:=xlWhole).Column
  lngEffectiveDateCol = oWorksheet.Rows(1).Find("EFFECTIVE DATE", lookat:=xlWhole).Column
  lngStandardRateCol = oWorksheet.Rows(1).Find("STANDARD RATE", lookat:=xlWhole).Column
  lngOvertimeRateCol = oWorksheet.Rows(1).Find("OVERTIME RATE", lookat:=xlWhole).Column
  lngCostPerUseCol = oWorksheet.Rows(1).Find("COST PER USE", lookat:=xlWhole).Column
  'sort for efficiency: RESOURCE,RATE TABLE,EFFECTIVE DATE
  If oWorksheet.AutoFilterMode = False Then oWorksheet.[A1].AutoFilter
  oWorksheet.AutoFilter.Sort.SortFields.Clear
  oWorksheet.AutoFilter.Sort.SortFields.Add2 Key:= _
      oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A2].End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
      xlSortNormal
  oWorksheet.AutoFilter.Sort.SortFields.Add2 Key:= _
      oWorksheet.Range(oWorksheet.[C2], oWorksheet.[C2].End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
      xlSortNormal
  oWorksheet.AutoFilter.Sort.SortFields.Add2 Key:= _
      oWorksheet.Range(oWorksheet.[D2], oWorksheet.[D2].End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
      xlSortNormal
  With oWorksheet.AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With
  
  lngLastRow = oWorksheet.[A1048576].End(-4162).Row '-4162 = xlUp
  For lngRow = 2 To lngLastRow
    strResourceName = Trim(oWorksheet.Cells(lngRow, lngResourceNameCol))
    'get/add resource
    If lngRow = 2 Then
      On Error Resume Next
      Set oResource = ActiveProject.Resources(strResourceName)
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    Else
      If oResource Is Nothing Then
        On Error Resume Next
        Set oResource = ActiveProject.Resources(strResourceName)
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      ElseIf Trim(oResource.Name) <> strResourceName Then
        Set oResource = Nothing
        On Error Resume Next
        Set oResource = ActiveProject.Resources(strResourceName)
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      Else
        GoTo cost_rate_tables
      End If
    End If
    If oResource Is Nothing Then
      If blnAddResources Then
        Set oResource = ActiveProject.Resources.Add(strResourceName)
        If Not oUpdated.Exists(strResourceName) Then
          oUpdated.Add oResource.UniqueID & "|" & strResourceName, "ADDED"
        End If
        strType = oWorksheet.Cells(lngRow, lngResourceTypeCol).Value
        oResource.Type = Switch(strType = "WORK", pjResourceTypeWork, strType = "COST", pjResourceTypeCost, strType = "MATERIAL", pjResourceTypeMaterial)
        If blnImportStatus Then
          oResource.SetField lngField, "ADDED"
        End If
        GoTo cost_rate_tables
      Else
        oWorksheet.Cells(lngRow, lngResourceNameCol).Style = "BAD"
        On Error Resume Next
        Set oComment = oWorksheet.Cells(lngRow, lngResourceNameCol).Comment
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If oComment Is Nothing Then
          oWorksheet.Cells(lngRow, lngResourceNameCol).AddComment "RESOURCE NOT FOUND"
        Else
          If InStr("RESOURCE NOT FOUND", oComment.Text) = 0 Then
            oComment.Text "RESOURCE NOT FOUND" & vbCrLf & oComment.Text
          End If
        End If
        If Not oUpdated.Exists(0 & "|" & strResourceName) Then
          oUpdated.Add 0 & "|" & strResourceName, "RESOURCE NOT FOUND"
        Else
          oUpdated(0 & "|" & strResourceName) = "RESOURCE NOT FOUND"
        End If
        GoTo next_row
      End If
    Else
      If Not oUpdated.Exists(oResource.UniqueID & "|" & strResourceName) Then
        oUpdated.Add oResource.UniqueID & "|" & strResourceName, "UPDATED: "
      Else
        oUpdated(oResource.UniqueID & "|" & strResourceName) = "UPDATED: "
      End If
      If blnImportStatus Then
        oResource.SetField lngField, "UPDATED: "
      End If
    End If
        
    'get cost rate table
cost_rate_tables:
    strCostRateTable = oWorksheet.Cells(lngRow, lngRateTableCol).Value
    lngCostRateTable = Switch(strCostRateTable = "A", 1, strCostRateTable = "B", 2, strCostRateTable = "C", 3, strCostRateTable = "D", 4, strCostRateTable = "E", 5)
    Set oCostRateTable = oResource.CostRateTables(lngCostRateTable)
    If InStr(oUpdated(oResource.UniqueID & "|" & strResourceName), "ADDED") = 0 Then
      If InStr(Split(oUpdated(oResource.UniqueID & "|" & strResourceName), ": ")(1), strCostRateTable) = 0 Then 'cost rate table not wiped yet
        If blnOverwrite Then
          For Each oPayRate In oCostRateTable.PayRates
            If oPayRate.Index = 1 Then
              oPayRate.StandardRate = 0
              oPayRate.OvertimeRate = 0
              oPayRate.CostPerUse = 0
            Else
              oPayRate.Delete
            End If
          Next oPayRate
          oUpdated(oResource.UniqueID & "|" & strResourceName) = oUpdated(oResource.UniqueID & "|" & strResourceName) & strCostRateTable & IIf(strCostRateTable <> "E", ",", "")
          If blnImportStatus Then
            oResource.SetField lngField, oResource.GetField(lngField) & strCostRateTable & IIf(strCostRateTable <> "E", ",", "")
          End If
        Else
          'todo: allow append vs overwrite?
        End If
      End If
    End If
    vEffectiveDate = oWorksheet.Cells(lngRow, lngEffectiveDateCol).Value
    vStdRate = oWorksheet.Cells(lngRow, lngStandardRateCol).Value
    vOvtRate = oWorksheet.Cells(lngRow, lngOvertimeRateCol).Value
    vCostPerUse = oWorksheet.Cells(lngRow, lngCostPerUseCol).Value
    Set oPayRate = Nothing
    If vEffectiveDate < #1/1/1984# Then
      oWorksheet.Cells(lngRow, lngEffectiveDateCol).Style = "Bad"
    ElseIf vEffectiveDate >= #12/31/2149# Then
      oWorksheet.Cells(lngRow, lngEffectiveDateCol).Style = "Bad"
    Else
      If blnOverwrite Then
        If oCostRateTable.PayRates.Count = 1 Then
          If cptRegEx(oCostRateTable.PayRates(1).StandardRate, "[0-9]{1,}\.[0-9]{1,}") = 0 Then
            Set oPayRate = oCostRateTable.PayRates(1)
          Else
            oCostRateTable.PayRates.Add vEffectiveDate
            Set oPayRate = oCostRateTable.PayRates(oCostRateTable.PayRates.Count)
          End If
        Else
          oCostRateTable.PayRates.Add vEffectiveDate
          Set oPayRate = oCostRateTable.PayRates(oCostRateTable.PayRates.Count)
        End If
      Else
        Set oPayRate = Nothing
        For Each oPayRate In oCostRateTable.PayRates
          If oPayRate.EffectiveDate = vEffectiveDate Then Exit For
        Next oPayRate
        If oPayRate Is Nothing Then
          oCostRateTable.PayRates.Add vEffectiveDate
          Set oPayRate = oCostRateTable.PayRates(oCostRateTable.PayRates.Count)
        End If
      End If
    End If
    oPayRate.StandardRate = vStdRate
    If Not IsEmpty(vOvtRate) And oResource.Type = pjResourceTypeWork Then oPayRate.OvertimeRate = vOvtRate
    If Not IsEmpty(vCostPerUse) Then oPayRate.CostPerUse = vCostPerUse
next_row:
    Application.StatusBar = Format(lngRow, "#,##0") & "/" & Format(lngLastRow, "#,##0") & "...(" & Format(lngRow / lngLastRow, "0%") & ")"
    myCostRateTables_frm.lblStatus.Caption = Format(lngRow, "#,##0") & "/" & Format(lngLastRow, "#,##0") & "...(" & Format(lngRow / lngLastRow, "0%") & ")"
    myCostRateTables_frm.lblProgress.Width = (lngRow / lngLastRow) * myCostRateTables_frm.lblStatus.Width
    DoEvents
  Next lngRow
  
  If Not blnImportStatus Then
    lngFile = FreeFile
    strFileName = Environ("tmp") & "\cpt-CostRateTableImportStatus.csv"
    Open strFileName For Output As #lngFile
    Print #lngFile, "UID,RESOURCE,STATUS_NOTE"
    For lngItem = 0 To oUpdated.Count - 1
      Print #lngFile, Split(oUpdated.Keys(lngItem), "|")(0) & "," & Chr(34) & Split(oUpdated.Keys(lngItem), "|")(1) & Chr(34) & "," & Chr(34) & oUpdated.Items(lngItem) & Chr(34)
    Next lngItem
    Close #lngFile
    ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
  End If
  
  With myCostRateTables_frm
    .lblProgress.Width = .lblStatus.Width
    .lblStatus.Caption = "Complete."
  End With
  Application.StatusBar = "Complete."
  
  oWorkbook.Close False
  
exit_here:
  On Error Resume Next
  Set oUpdated = Nothing
  Reset
  Application.StatusBar = ""
  Application.ScreenUpdating = True
  Application.Calculation = pjAutomatic
  Set oPayRate = Nothing
  Set oCostRateTable = Nothing
  Set oResource = Nothing
  Set oComment = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptCostRateTables_bas", "cptImportCostRateTables", Err, Erl)
  Resume exit_here
End Sub
