VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptCustomFieldUsage_frm 
   Caption         =   "Custom Field Usage"
   ClientHeight    =   3975
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
  OptionsViewEx Displaysummarytasks:=blnSummaries, DisplayNameIndent:=blnSummaries
  If Not Me.Visible Then Exit Sub
  If Not IsNull(Me.lboFieldTypes.Value) Then
    If Not IsNull(Me.lboCustomFields.Value) Then lngSelected = Me.lboCustomFields.ListIndex
    Me.lboFieldTypes_Click
    Me.lboCustomFields.Selected(lngSelected) = True
  End If
  Set oTasks = Nothing
End Sub

Private Sub cmdClear_Click()
  If Not IsNull(Me.lboCustomFields.Value) Then
    Application.OpenUndoTransaction "Clear " & FieldConstantToFieldName(Me.lboCustomFields.Value)
    SelectColumn 4
    SetField FieldConstantToFieldName(Me.lboCustomFields.Value), ""
    SelectBeginning
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

Private Sub lboCustomFields_Click()
  Dim oTasks As MSProject.Tasks
  cptUpdateCustomFieldUsageView Me.lboCustomFields.Value, Me.lboFieldTypes.Value, True
  SelectColumn 4
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If Not oTasks Is Nothing Then
    Me.lboCustomFields.List(Me.lboCustomFields.ListIndex, 3) = Format(oTasks.Count, "#,##0")
  Else
    Me.lboCustomFields.List(Me.lboCustomFields.ListIndex, 3) = 0
  End If
  Me.lblFormula.Visible = cptHasFormula(ActiveProject, Me.lboCustomFields.Value)
  Me.lblLookup.Visible = cptHasLookup(ActiveProject, Me.lboCustomFields.Value)
  If Not Me.lblFormula.Visible And Me.lblLookup.Visible Then
    Me.lblLookup.Top = Me.lblFormula.Top
  Else
    Me.lblLookup.Top = 102
  End If
  Set oTasks = Nothing
End Sub

Sub lboFieldTypes_Click()
  Dim oTasks As MSProject.Tasks
  Dim strFieldType As String
  Dim strCustomFieldName As String
  Dim lngLCF As Long
  Dim lngField As Long
  Dim lngCustomFieldCount As Long
  
  cptSpeed True
  
  strFieldType = Me.lboFieldTypes.Value
  If oCustomFields.Exists(strFieldType) Then
    lngCustomFieldCount = oCustomFields(strFieldType)
  Else
    lngCustomFieldCount = 10
  End If
  If strFieldType = "Outline Code" Then
    Me.lboCustomFields.ColumnWidths = "0 pt;75 pt;120 pt;15 pt"
  Else
    Me.lboCustomFields.ColumnWidths = "0 pt;55 pt;120 pt;15 pt"
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
      Me.lboCustomFields.List(Me.lboCustomFields.ListCount - 1, 3) = Format(oTasks.Count, "#,##0")
    End If
    SetAutoFilter FieldConstantToFieldName(lngLCF), pjAutoFilterClear
    SelectBeginning
    Me.lblStatus.Caption = Format(lngField / lngCustomFieldCount, "0%")
    Me.lblProgress.Width = (lngField / lngCustomFieldCount) * Me.lblStatus.Width
    Me.Repaint
  Next lngField
  cptUpdateCustomFieldUsageView 0
exit_here:
  On Error Resume Next
  Me.lblStatus.Visible = False
  Me.lblProgress.Visible = False
  cptSpeed False
  Set oTasks = Nothing
  Exit Sub
err_here:
  cptHandleErr "cptCustomFieldUsage_frm", "lboFieldTypes_Click", Err, Erl
  Resume exit_here
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Set oCustomFields = Nothing
  Me.Hide
  ViewApply strCustomFieldUsageStartingView
  TableApply strCustomFieldUsageStartingTable
  GroupApply strCustomFieldUsageStartingGroup
  FilterApply strCustomFieldUsageStartingFilter
  If cptViewExists("cptCustomFieldUsage View") Then ActiveProject.Views("cptCustomFieldUsage View").Delete
  If cptTableExists("cptCustomFieldUsage Table") Then ActiveProject.TaskTables("cptCustomFieldUsage Table").Delete
End Sub
