VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSaveLocal_frm 
   Caption         =   "Save ECF to LCF"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   OleObjectBlob   =   "cptSaveLocal_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptSaveLocal_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.4.0</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptSaveLocal_frm"
Private Const TABLE_NAME As String = ".cptSaveLocal Table"
Private Const VIEW_NAME As String = ".cptSaveLocal View"

Private Sub cboECF_Change()
  Call cptUpdateECF(Me, Me.txtFilterECF)
End Sub

Sub cboLCF_Change()
  Call cptUpdateLCF(Me, Me.txtFilterLCF)
End Sub

Private Sub chkAutoSwitch_Click()
  If Not Me.Visible Then Exit Sub
  If Not IsNull(Me.lboLCF.Value) Then
    If Me.cboLCF <> Me.lboECF.List(Me.lboECF.ListIndex, 2) Then
      Me.cboLCF = Me.lboECF.List(Me.lboECF.ListIndex, 2)
    End If
  End If
End Sub

Private Sub cmdAutoMap_Click()
  cptAutoMap Me
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdCustomFields_Click()
  'long
  Dim lngSelected As Long
  'string
  Dim strDescription As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  ActiveWindow.TopPane.Activate
  Application.CustomizeField
  Me.cboLCF_Change
  'todo: update LCF name
  'todo: if mapped field is renamed, prompt to unmap
  'todo: if unmapped, prompt to clear out data
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cmdCustomFields_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdDeleteMap_Click()
  cptDeleteCFMap Me
End Sub

Private Sub cmdDeleteTaskData_Click()
  cptDeleteLCFAndTaskData Me
End Sub

Private Sub cmdExportMap_Click()
  cptExportCFMap
End Sub

Private Sub cmdImportMap_Click()
  cptImportCFMap Me
End Sub

Private Sub cmdMap_Click()
  If Not IsNull(Me.lboECF) And Not IsNull(Me.lboLCF) Then
    Call cptMapECFtoLCF(Me, Me.lboECF, Me.lboLCF)
  End If
End Sub

Private Sub cmdSaveLocal_Click()
  cptSaveSetting "SaveLocal", "txtRenamingConvention", Me.txtRenamingConvention.Value
  cptSaveLocal Me
End Sub

Private Sub cmdUnmap_Click()
  'objects
  Dim oTableField As Object
  Dim rstSavedMap As Object 'ADODB.Recordset
  'strings
  Dim strProgramAcronym As String
  Dim strSavedMap As String
  'longs
  Dim lngECF As Long
  Dim lngItem As Long
  Dim lngLCF As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Me.lboECF.ListIndex < 0 Then GoTo exit_here
  If IsNull(Me.lboECF.List(Me.lboECF.ListIndex, 3)) Or Me.lboECF.List(Me.lboECF.ListIndex, 3) = "" Then GoTo exit_here

  If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Please Confirm") = vbNo Then GoTo exit_here
  
  'get the ECF and LCF
  lngECF = Me.lboECF.List(Me.lboECF.ListIndex, 0)
  lngLCF = Me.lboECF.List(Me.lboECF.ListIndex, 3)
  
  'delete it from LCF
  CustomFieldDelete lngLCF
  
  strProgramAcronym = cptGetProgramAcronym
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) <> vbNullString Then
    Set rstSavedMap = CreateObject("ADODB.Recordset")
    rstSavedMap.Open strSavedMap
    rstSavedMap.Filter = "PROGRAM_ACRONYM='" & strProgramAcronym & "' AND ECF=" & lngECF & " AND LCF=" & lngLCF
    If Not rstSavedMap.EOF Then
      rstSavedMap.Delete adAffectCurrent
    End If
    rstSavedMap.Filter = ""
    rstSavedMap.Save strSavedMap, adPersistADTG
  End If
  
  'remove from lboECF
  Me.lboECF.List(Me.lboECF.ListIndex, 3) = ""
  Me.lboECF.List(Me.lboECF.ListIndex, 4) = ""
  
  'rename in lboLocal
  For lngItem = 0 To Me.lboLCF.ListCount - 1
    If Me.lboLCF.List(lngItem, 0) = lngLCF Then
      Me.lboLCF.List(lngItem, 1) = FieldConstantToFieldName(lngLCF)
    End If
  Next lngItem

  'remove from cptSaveLocal table
  For Each oTableField In ActiveProject.TaskTables(TABLE_NAME).TableFields
    If oTableField.Field = lngECF Or oTableField.Field = lngLCF Then
      oTableField.Delete
    End If
  Next oTableField
  TableApply TABLE_NAME
  
exit_here:
  On Error Resume Next
  Set oTableField = Nothing
  Set rstSavedMap = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cmdUnmap", Err, Erl)
  Resume exit_here
End Sub

Private Sub lblSelectAll_Click()
  Dim lngItem As Long
  For lngItem = 0 To Me.lboECF.ListCount - 1
    If IsNull(Me.lboECF.List(lngItem, 3)) Then
      Me.lboECF.Selected(lngItem) = True
    End If
  Next lngItem
  cptAnalyzeAutoMap Me
End Sub

Private Sub lblSelectNone_Click()
  Dim lngItem As Long
  For lngItem = 0 To Me.lboECF.ListCount - 1
    Me.lboECF.Selected(lngItem) = False
  Next lngItem
  cptAnalyzeAutoMap Me
End Sub

Private Sub lblShowFormula_Click()
  If Me.lboECF.ListIndex >= 0 Then
    MsgBox CustomFieldGetFormula(Me.lboECF.List(Me.lboECF.ListIndex, 0)), vbInformation + vbOKOnly, "Formula:"
  End If
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then
    CreateObject("WScript.Shell").Run "https://www.ClearPlanConsulting.com"
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "lblURL", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboECF_Change()
  'objects
  'strings
  'longs
  Dim lngItems As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  Me.lblShowFormula.Visible = False
  If Me.Visible Then
    If Me.ActiveControl.Name = "lboECF" Or Me.ActiveControl.Name = "cmdImportMap" Then
      If Me.tglAutoMap Then
        For lngItem = 0 To Me.lboECF.ListCount - 1
          If Me.lboECF.Selected(lngItem) Then lngItems = lngItems + 1
        Next lngItem
        Me.lblStatus.Caption = lngItems & " ECFs selected."
        cptAnalyzeAutoMap Me
      Else
        If Me.chkAutoSwitch Then
          If Me.lboECF.List(Me.lboECF.ListIndex, 2) = "" Then
            Me.cboLCF = "Text"
          Else
            Me.cboLCF = Replace(Me.lboECF.List(Me.lboECF.ListIndex, 2), "Maybe", "")
          End If
          For lngItem = 0 To Me.lboLCF.ListCount - 1
            If Not IsNull(Me.lboECF.List(Me.lboECF.ListIndex, 3)) Then
              If Me.lboLCF.List(lngItem, 0) = Me.lboECF.List(Me.lboECF.ListIndex, 3) Then
                Me.lboLCF.Selected(lngItem) = True
                Exit For
              End If
            End If
            If CustomFieldGetName(Me.lboLCF.List(lngItem)) = "" Then
              Me.lboLCF.Selected(lngItem) = True
              Exit For
            End If
          Next lngItem
        End If
      End If
    End If
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "lboECF_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboECF_Click()
  'objects
  'strings
  Dim strSwitch As String
  Dim strECF As String
  'longs
  Dim lngItem As Long
  Dim lngECF As Long
  Dim lngLCF As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Me.lboECF.MultiSelect = fmMultiSelectSingle Then
    Me.lblShowFormula.Visible = False
    Me.lblStatus.Caption = "Analyzing..."
  
    lngECF = Me.lboECF.List(Me.lboECF.ListIndex, 0)
    strECF = Me.lboECF.List(Me.lboECF.ListIndex, 1)
    
    Select Case Me.lboECF.List(Me.lboECF.ListIndex, 2)
      Case "Cost"
        Me.lblStatus.Caption = "This is likely a Cost field."
        strSwitch = "Cost"
      Case "Date"
        Me.lblStatus.Caption = "This is likely a Date field."
        strSwitch = "Date"
      Case "Duration"
        Me.lblStatus.Caption = "This is likely a Duration field."
        strSwitch = "Duration"
      Case "Flag"
        Me.lblStatus.Caption = "This is likely a Flag field."
        strSwitch = "Flag"
      Case "MaybeFlag"
        Me.lblStatus.Caption = "This is likely a Flag field."
        strSwitch = "Flag"
      Case "Number"
        Me.lblStatus.Caption = "This is likely a Number field."
        strSwitch = "Number"
      Case "Outline Code"
        Me.lblStatus.Caption = "This field requires an Outline Code."
        strSwitch = "Outline Code"
      Case "MaybeText"
        Me.lblStatus.Caption = "This is likely a Text field."
        strSwitch = "Text"
      Case "Text"
        Me.lblStatus.Caption = "This is likely a Text field."
        strSwitch = "Text"
      Case Else
        Me.lblStatus.Caption = "Undetermined: confirm manually."
        strSwitch = "Text"
    End Select
    
    If Me.chkAutoSwitch Then
      If Len(Me.lboECF.List(Me.lboECF.ListIndex, 3)) > 0 Then
        lngLCF = Me.lboECF.List(Me.lboECF.ListIndex, 3)
        If Me.lboLCF.Value <> lngLCF Then
          Me.cboLCF.Value = cptRxMatch(FieldConstantToFieldName(lngLCF), "Cost|Date|Duration|Finish|Number|Outline Code|Start|Text")
          Me.lboLCF.Value = lngLCF
        End If
      Else
        Me.cboLCF.Value = strSwitch
        'Me.cboLCF.Value = strSwitch 'why twice? supposed to be lboECF?
      End If
    End If
    
    If Len(CustomFieldGetFormula(Me.lboECF)) > 0 Then
      Me.lblShowFormula.Visible = True
    End If
  Else
    'todo: anything here?
  End If
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "lboECF_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboECF_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  'todo: if mapped, then unmap
  'todo: if unmapped, then map
  If Not IsNull(Me.lboECF) And Not IsNull(Me.lboLCF) Then
    Call cptMapECFtoLCF(Me, Me.lboECF, Me.lboLCF)
  End If
End Sub

Private Sub lboLCF_Click()
  'objects
  'strings
  'longs
  Dim lngItem As Long
  Dim lngLCF As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Me.ActiveControl.Name = "lboECF" Then GoTo exit_here
  If Len(Me.lboLCF.List(Me.lboLCF.ListIndex, 0)) > 0 Then
    lngLCF = Me.lboLCF.Value
    For lngItem = 0 To Me.lboECF.ListCount - 1
      If Me.lboECF.List(lngItem, 3) = lngLCF Then
        Me.lboECF.Value = Me.lboECF.List(lngItem, 0)
        Exit For
      End If
    Next lngItem
  End If
  
'  If Me.lboLCF.MultiSelect = fmMultiSelectSingle Then
'    Me.lblShowFormula.Visible = False
'    Me.lblStatus.Caption = "Analyzing..."
'
'    lngECF = Me.lboLCF.List(Me.lboLCF.ListIndex, 0)
'    strECF = Me.lboLCF.List(Me.lboLCF.ListIndex, 1)
'
'    Select Case Me.lboECF.List(Me.lboECF.ListIndex, 2)
'      Case "Cost"
'        Me.lblStatus.Caption = "This is likely a Cost field."
'        strSwitch = "Cost"
'      Case "Date"
'        Me.lblStatus.Caption = "This is likely a Date field."
'        strSwitch = "Date"
'      Case "Duration"
'        Me.lblStatus.Caption = "This is likely a Duration field."
'        strSwitch = "Duration"
'      Case "Flag"
'        Me.lblStatus.Caption = "This is likely a Flag field."
'        strSwitch = "Flag"
'      Case "MaybeFlag"
'        Me.lblStatus.Caption = "This is likely a Flag field."
'        strSwitch = "Flag"
'      Case "Number"
'        Me.lblStatus.Caption = "This is likely a Number field."
'        strSwitch = "Number"
'      Case "Outline Code"
'        Me.lblStatus.Caption = "This field requires an Outline Code."
'        strSwitch = "Outline Code"
'      Case "MaybeText"
'        Me.lblStatus.Caption = "This is likely a Text field."
'        strSwitch = "Text"
'      Case "Text"
'        Me.lblStatus.Caption = "This is likely a Text field."
'        strSwitch = "Text"
'      Case Else
'        Me.lblStatus.Caption = "Undetermined: confirm manually."
'        strSwitch = "Text"
'    End Select
'
'    If Me.chkAutoSwitch And Me.cboLCF.Value <> strSwitch Then
'      If Len(Me.lboECF.List(Me.lboECF.ListIndex, 3)) > 0 Then
'        Me.cboLCF.Value = cptRxMatch(FieldConstantToFieldName(Me.lboECF.List(Me.lboECF.ListIndex, 3)), "Cost|Date|Duration|Finish|Number|Outline Code|Start|Text")
'        Me.lboLCF.Value = Me.lboECF.List(Me.lboECF.ListIndex, 3)
'      Else
'        Me.cboLCF.Value = strSwitch
'      End If
'    End If
'
'    If Len(CustomFieldGetFormula(Me.lboECF)) > 0 Then
'      Me.lblShowFormula.Visible = True
'    End If
'  Else
'    'todo: anything here?
'  End If
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "lboLCF_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub tglAutoMap_Click()
  'objects
  'strings
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Me.tglAutoMap Then
    Me.lboECF.MultiSelect = fmMultiSelectMulti
    For lngItem = 0 To Me.lboECF.ListCount - 1
      If IsNull(Me.lboECF.List(lngItem, 3)) Then
        Me.lboECF.Selected(lngItem) = True
      End If
    Next lngItem
    cptAnalyzeAutoMap Me
    Me.lboLCF.Visible = False
    Me.txtAutoMap.Visible = True
    Me.lblStatus.Caption = Me.lboECF.ListCount & " ECFs selected."
    Me.lblSelectAll.Visible = True
    Me.lblSelectNone.Visible = True
    Me.cmdMap.Enabled = False
    Me.lblLCF.ForeColor = 8421504
    Me.lblTypesLCF.ForeColor = 8421504
    Me.txtFilterLCF.Enabled = False
    Me.cboLCF.Enabled = False
    Me.txtAutoMap.Height = Me.lboECF.Height
  Else
    Me.lboECF.MultiSelect = fmMultiSelectSingle
    Me.lboECF.ListIndex = 0
    Me.txtAutoMap.Visible = False
    Me.lboLCF.Visible = True
    Me.cmdAutoMap.Enabled = False
    Me.lblSelectAll.Visible = False
    Me.lblSelectNone.Visible = False
    Me.cmdMap.Enabled = True
    Me.lblLCF.ForeColor = -2147483630
    Me.lblTypesLCF.ForeColor = -2147483630
    Me.txtFilterLCF.Enabled = True
    Me.cboLCF.Enabled = True
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "tglAutoMap_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtFilterECF_Change()
  Call cptUpdateECF(Me, Me.txtFilterECF.Text)
End Sub

Private Sub txtFilterLCF_Change()
  Call cptUpdateLCF(Me, Me.txtFilterLCF.Text)
End Sub

Private Sub txtRenamingConvention_Change()
  Dim strPreview As String
  strPreview = Replace(Me.txtRenamingConvention.Text, "[ECF]", "Enterprise Custom Field Name")
  strPreview = Replace(Me.txtRenamingConvention.Text, "[LCF]", "Text1")
  If Me.Visible Then
    If Me.ActiveControl.Name = "txtRenamingConvention" Then
      Me.lblRenamingConvention.Caption = strPreview
    End If
  End If
End Sub

Private Sub txtRenamingConvention_Enter()
  Dim strPreview As String
  strPreview = Replace(Me.txtRenamingConvention.Text, "[ECF]", "Enterprise Custom Field Name")
  strPreview = Replace(Me.txtRenamingConvention.Text, "[LCF]", "Text1")
  With Me.lblRenamingConvention
    .BorderStyle = fmBorderStyleSingle
    .BorderColor = 12632256
    .ForeColor = 8421504
    .TextAlign = fmTextAlignCenter
    .Caption = strPreview
  End With
End Sub

Private Sub txtRenamingConvention_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  With Me.lblRenamingConvention
    .BorderStyle = fmBorderStyleNone
    .ForeColor = -2147483630
    .TextAlign = fmTextAlignRight
    .Caption = "LCF Renaming Convention:"
  End With
  cptSaveSetting "SaveLocal", "txtRenamingConvention", Me.txtRenamingConvention.Value
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub UserForm_Terminate()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  On Error Resume Next
  If MsgBox("Keep Current View?" & vbCrLf & vbCrLf & "(Click 'No' to restore previous.)", vbQuestion + vbYesNo, "ECF to LCF") = vbNo Then
    If Len(strStartView) > 0 Then
      If Not ViewApply(strStartView, True) Then ViewApply "Gantt Chart"
    End If
    If Len(strStartTable) > 0 Then
      If strStartTable = TABLE_NAME Then
        TableApply "Entry"
      Else
        If Not TableApply(strStartTable) Then TableApply "Entry"
      End If
    End If
    If Len(strStartFilter) > 0 Then
      If Not FilterApply(strStartFilter) Then FilterApply "All Tasks"
    End If
    If Len(strStartGroup) > 0 Then
      If Not GroupApply(strStartGroup) Then GroupApply "No Group"
    End If
    
    If ActiveProject.CurrentView = VIEW_NAME Then ViewApply "Gantt Chart"
    On Error Resume Next
    ActiveProject.Views(VIEW_NAME).Delete
    ActiveProject.TaskTables(TABLE_NAME).Delete
  End If
  
exit_here:
  On Error Resume Next
  Unload Me
  
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "UserForm_Terminate", Err, Erl)
  Resume exit_here
End Sub
