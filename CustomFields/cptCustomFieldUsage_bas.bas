Attribute VB_Name = "cptCustomFieldUsage_bas"
'<cpt_version>v0.1.0</cpt_version>
Option Explicit
Public oCustomFields As Scripting.Dictionary
Public strCustomFieldUsageStartingView As String
Public strCustomFieldUsageStartingTable As String
Public strCustomFieldUsageStartingGroup As String
Public strCustomFieldUsageStartingFilter As String

Function GetTaskViewList(Optional strDelimiter As String = ";") As String
  Dim oView As MSProject.View
  Dim strViewList As String
  
  On Error GoTo err_here
  
  For Each oView In ActiveProject.Views
    If oView.Type = pjTaskItem Then
      If oView.Screen = pjGantt Then
        If Len(strViewList) = 0 Then
          strViewList = oView.Name
        Else
          strViewList = strViewList & strDelimiter & oView.Name
        End If
      End If
    End If
  Next oView
exit_here:
  On Error Resume Next
  GetTaskViewList = strViewList
  Set oView = Nothing
  Exit Function
err_here:
  MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
  Resume exit_here
End Function

Sub cptShowCustomFieldUsage_frm()
  'objects
  Dim myCustomFieldUsage_frm As cptCustomFieldUsage_frm
  'strings
  'longs
  Dim lngLCF As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  Dim lngField As Long
  Dim lngFields As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vFieldType As Variant
  'dates
  
  If Not cptGetUserForm("cptCustomFieldUsage_frm") Is Nothing Then Exit Sub 'prevent spawning
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'capture view/table/group/filter
  ActiveWindow.TopPane.Activate
  If ActiveWindow.ActivePane.View.Type <> pjTaskItem Then ViewApply "Gantt Chart"
  strCustomFieldUsageStartingView = ActiveProject.CurrentView
  If strCustomFieldUsageStartingView = "cptCustomFieldUsage View" Then
    strCustomFieldUsageStartingView = "Gantt Chart"
    ViewApply strCustomFieldUsageStartingView
  End If
  strCustomFieldUsageStartingTable = ActiveProject.CurrentTable
  strCustomFieldUsageStartingFilter = ActiveProject.CurrentFilter
  strCustomFieldUsageStartingGroup = ActiveProject.CurrentGroup
  
  'create view/table (no group, no filter)
  cptUpdateCustomFieldUsageView 0
  If cptViewExists("cptCustomFieldUsage View") Then ActiveProject.Views("cptCustomFieldUsage View").Delete
  ViewEditSingle "cptCustomFieldUsage View", True, , pjGantt, False, False, "cptCustomFieldUsage Table", "All Tasks", "No Group"
  ViewApply "cptCustomFieldUsage View"
  OptionsViewEx Displaysummarytasks:=True, DisplayNameIndent:=True, DisplayOutlineSymbols:=True
  OutlineShowAllTasks
  SelectBeginning
  
  Set oCustomFields = CreateObject("Scripting.Dictionary")
  oCustomFields.Add "Flag", 20
  oCustomFields.Add "Number", 20
  oCustomFields.Add "Text", 30
  
  Set myCustomFieldUsage_frm = New cptCustomFieldUsage_frm
  With myCustomFieldUsage_frm
    .Caption = "Local Custom Field Usage (" & cptGetVersion("cptCustomFieldUsage_bas") & ")"
    .lboFieldTypeHeader.AddItem "Data Type"
    .lboCustomFieldsHeader.AddItem
    .lboCustomFieldsHeader.List(0, 0) = "Constant"
    .lboCustomFieldsHeader.List(0, 1) = "Field Name"
    .lboCustomFieldsHeader.List(0, 2) = "Custom Field Name"
    .lboCustomFieldsHeader.List(0, 3) = "Tasks"
    .lboCustomFieldsHeader.ColumnWidths = "0 pt;55 pt;120 pt;15 pt"
    .lboFieldTypeHeader.Height = .lboCustomFieldsHeader.Height
    .lboFieldTypeHeader.Locked = True
    .lboCustomFieldsHeader.Locked = True
    For Each vFieldType In Split("Cost,Date,Duration,Finish,Flag,Number,Outline Code,Start,Text", ",")
      .lboFieldTypes.AddItem vFieldType
    Next vFieldType
    .chkIncludeSummaryTasks.Value = True
    .lblFormula.Visible = False
    .lblLookup.Visible = False
    .lblProgress.Visible = False
    .lblStatus.Visible = False
    .lboCustomFields.Height = .lboFieldTypes.Height
    .tglAll.Top = .lboCustomFields.Top + .lboCustomFields.Height - .tglAll.Height
    .tglAll.Value = False 'user setting?
    .Show (False)
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("foo", "bar", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateCustomFieldUsageView(lngLCF As Long, Optional strFieldType As String, Optional blnFilter As Boolean = False)
  FilterClear
  If lngLCF > 0 Then
    TableEditEx "cptCustomFieldUsage Table", True, True, True, , "ID", , , , , False, True, , , , , False, False, False, False
    TableEditEx "cptCustomFieldUsage Table", True, , , , , "Unique ID", "UID", , , , True
    TableEditEx "cptCustomFieldUsage Table", True, , , , , "Name", , 85, , , True
    TableEditEx "cptCustomFieldUsage Table", True, , , , , FieldConstantToFieldName(lngLCF), , , , , True
    If blnFilter Then
      Select Case strFieldType
        Case "Cost"
          SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterCustom, "does not equal", "$0.00"
        Case "Date"
          SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterCustom, "does not equal", "NA"
        Case "Duration"
          SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterCustom, "does not equal", 0
        Case "Finish"
          SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterCustom, "does not equal", "NA"
        Case "Flag"
          SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterFlagYes
        Case "Number"
          SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterCustom, "does not equal", 0
        Case "Outline Code"
          SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterCustom, "does not equal", ""
        Case "Start"
          SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterCustom, "does not equal", "NA"
        Case "Text"
          SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterCustom, "does not equal", ""
      End Select
    End If
  Else
    TableEditEx "cptCustomFieldUsage Table", True, True, True, , "ID", , , , , False, True, , , , , False, False, False, True
    TableEditEx "cptCustomFieldUsage Table", True, , , , , "Unique ID", "UID", , , , True
    TableEditEx "cptCustomFieldUsage Table", True, , , , , "Name", , 85, , , True
  End If
  TableApply "cptCustomFieldUsage Table"
  SetSplitBar 4
End Sub
