VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptCustomFieldUsage_frm 
   Caption         =   "Custom Field Usage"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   OleObjectBlob   =   "cptCustomFieldUsage_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptCustomFieldUsage_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.1.0</cpt_version>
Option Explicit

Private Sub chkIncludeSummaryTasks_Click()
  Dim blnSummaries As Boolean
  Dim oTasks As MSProject.Tasks
  Dim lngSelected As Long
  blnSummaries = Me.chkIncludeSummaryTasks
  OptionsViewEx DisplaySummaryTasks:=blnSummaries, DisplayNameIndent:=blnSummaries
  If Not Me.Visible Then Exit Sub
  If Not IsNull(Me.lboFieldTypes.Value) Then
    If Not IsNull(Me.lboCustomFields.Value) Then lngSelected = Me.lboCustomFields.ListIndex
    Me.lboFieldTypes_Click
    Me.lboCustomFields.Selected(lngSelected) = True
  End If
  Set oTasks = Nothing
End Sub

Private Sub cmdClear_Click()
  Dim blnMaster As Boolean
  Dim lngLCF As Long
  
  blnMaster = ActiveProject.Subprojects.Count > 0
  If Not IsNull(Me.lboCustomFields.Value) Then
    lngLCF = Me.lboCustomFields.Value
    If MsgBox("Really clear data from '" & FieldConstantToFieldName(lngLCF) & "'?" & vbCrLf & vbCrLf & "Undo *should* be available...but be careful!", vbQuestion + vbYesNo, "Please confirm") = vbNo Then Exit Sub
    Application.OpenUndoTransaction "cpt - Clear " & FieldConstantToFieldName(lngLCF)
    SelectTaskColumn FieldConstantToFieldName(lngLCF)
    If Me.lboFieldTypes.Value = "Flag" Then
      SetField FieldConstantToFieldName(lngLCF), "No"
    Else
      SetField FieldConstantToFieldName(lngLCF), ""
    End If
    Me.lboCustomFields.List(Me.lboCustomFields.ListIndex, 3) = 0
    Application.CloseUndoTransaction
  End If
End Sub

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub cmdRename_Click()
  Dim lngLCF As Long
  Dim strCFN As String
  If Not IsNull(Me.lboCustomFields.Value) Then
    lngLCF = Me.lboCustomFields.Value
    strCFN = InputBox("Rename " & FieldConstantToFieldName(lngLCF) & " to what:", "Custom Field Name")
    If cptCustomFieldExists(strCFN) > 0 Then
      MsgBox FieldConstantToFieldName(FieldNameToFieldConstant(strCFN)) & " is already named '" & strCFN & "'!", vbExclamation + vbOKOnly, "No Duplicates"
    Else
      CustomFieldRename lngLCF, strCFN
      Me.lboCustomFields.List(Me.lboCustomFields.ListIndex, 2) = strCFN
    End If
  End If
End Sub

Private Sub lblURL_Click()
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptCustomFieldUsage_frm", "lblURL_Click()", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboCustomFields_Click()
  If Not Me.tglAll Then
    cptUpdateCustomFieldUsageView Me.lboCustomFields.Value, Me.lboFieldTypes.Value, True
  End If
  SelectTaskColumn FieldConstantToFieldName(Me.lboCustomFields.Value)
  Me.lblFormula.Visible = cptHasFormula(ActiveProject, Me.lboCustomFields.Value)
  Me.lblLookup.Visible = cptHasLookup(ActiveProject, Me.lboCustomFields.Value)
  If Not Me.lblFormula.Visible And Me.lblLookup.Visible Then
    Me.lblLookup.Top = Me.lblFormula.Top
  Else
    Me.lblLookup.Top = 102
  End If
End Sub

Sub lboFieldTypes_Click()
  Dim oTask As MSProject.Task
  Dim oTasks As MSProject.Tasks
  Dim strFieldType As String
  Dim strCustomFieldName As String
  Dim lngLCF As Long
  Dim lngField As Long
  Dim lngCustomFieldCount As Long
  Dim lngTaskCount As Long
  
  cptSpeed True
  
  strFieldType = Me.lboFieldTypes.Value
  If oCustomFields.Exists(strFieldType) Then
    lngCustomFieldCount = oCustomFields(strFieldType)
  Else
    lngCustomFieldCount = 10
  End If
  If strFieldType = "Outline Code" Then
    Me.lboCustomFields.ColumnWidths = "0 pt;75 pt;120 pt;15 pt"
    Me.lboCustomFieldsHeader.ColumnWidths = "0 pt;75 pt;120 pt;15 pt"
  Else
    Me.lboCustomFields.ColumnWidths = "0 pt;55 pt;120 pt;15 pt"
    Me.lboCustomFieldsHeader.ColumnWidths = "0 pt;55 pt;120 pt;15 pt"
  End If
  Me.lboCustomFields.Clear
  Me.lblFormula.Visible = False
  Me.lblLookup.Visible = False
  Me.lblStatus.Top = 84
  Me.lblStatus.Caption = Format(lngField / lngCustomFieldCount, "0%")
  Me.lblStatus.Visible = True
  Me.lblProgress.Top = 84
  Me.lblProgress.Visible = True
  'Me.Repaint
  For lngField = 1 To lngCustomFieldCount
    lngLCF = FieldNameToFieldConstant(strFieldType & lngField, pjTask)
    strCustomFieldName = CustomFieldGetName(lngLCF)
    Me.lboCustomFields.AddItem
    Me.lboCustomFields.List(Me.lboCustomFields.ListCount - 1, 0) = lngLCF
    Me.lboCustomFields.List(Me.lboCustomFields.ListCount - 1, 1) = strFieldType & lngField
    Me.lboCustomFields.List(Me.lboCustomFields.ListCount - 1, 2) = strCustomFieldName
    cptUpdateCustomFieldUsageView lngLCF, Me.lboFieldTypes.Value, True
    SelectAll
    Set oTasks = Nothing
    On Error Resume Next
    Set oTasks = ActiveSelection.Tasks
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oTasks Is Nothing Then
      Me.lboCustomFields.List(Me.lboCustomFields.ListCount - 1, 3) = 0
    Else
      If Me.chkIncludeSummaryTasks Then
        lngTaskCount = 0
        For Each oTask In oTasks
          If Not oTask Is Nothing Then
            If Not oTask.Summary Then 'todo: what if the summary task has a value, dummy?
              lngTaskCount = lngTaskCount + 1
            End If
          End If
        Next oTask
        Me.lboCustomFields.List(Me.lboCustomFields.ListCount - 1, 3) = Format(lngTaskCount, "#,##0")
      Else
        Me.lboCustomFields.List(Me.lboCustomFields.ListCount - 1, 3) = Format(oTasks.Count, "#,##0")
      End If
    End If
    SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterClear
    SelectBeginning
    Me.lblStatus.Caption = Format(lngField / lngCustomFieldCount, "0%")
    Me.lblProgress.Width = (lngField / lngCustomFieldCount) * Me.lblStatus.Width
    Me.Repaint
  Next lngField
  cptUpdateCustomFieldUsageView 0
  If Me.tglAll Then
    Me.tglAll.Value = False
    Me.tglAll.Value = True
  End If
exit_here:
  On Error Resume Next
  Me.lblStatus.Visible = False
  Me.lblProgress.Visible = False
  cptSpeed False
  Set oTask = Nothing
  Set oTasks = Nothing
  Exit Sub
err_here:
  cptHandleErr "cptCustomFieldUsage_frm", "lboFieldTypes_Click", Err, Erl
  Resume exit_here
End Sub

Private Sub tglAll_Click()
  Dim lngItem As Long
'  Me.cmdClear.Enabled = Not Me.tglAll
'  Me.cmdRename.Enabled = Not Me.tglAll
  If Me.tglAll Then
    Me.lblFormula.Visible = False
    Me.lblLookup.Visible = False
    Me.lboCustomFields.ListIndex = -1
    'Me.lboCustomFields.Enabled = False
    For lngItem = 0 To Me.lboCustomFields.ListCount - 1
      If lngItem = 0 Then
        cptUpdateCustomFieldUsageView Me.lboCustomFields.List(lngItem, 0)
      Else
        TableEditEx "cptCustomFieldUsage Table", True, , , , , FieldConstantToFieldName(Me.lboCustomFields.List(lngItem, 0)), , , , , True
      End If
    Next lngItem
    TableApply "cptCustomFieldUsage Table"
    SetSplitBar 3 + Me.lboCustomFields.ListCount
  Else
    Me.lboCustomFields.Enabled = True
    Me.lboCustomFields.Value = Me.lboCustomFields.List(0, 0)
  End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Set oCustomFields = Nothing
  Me.Hide
  ViewApply strCustomFieldUsageStartingView
  TableApply strCustomFieldUsageStartingTable
  GroupApply strCustomFieldUsageStartingGroup
  FilterApply strCustomFieldUsageStartingFilter
  On Error Resume Next 'fails if view is in use, e.g., on another project or another window of same project
  If cptViewExists("cptCustomFieldUsage View") Then ActiveProject.Views("cptCustomFieldUsage View").Delete
  If cptTableExists("cptCustomFieldUsage Table") Then ActiveProject.TaskTables("cptCustomFieldUsage Table").Delete
End Sub
