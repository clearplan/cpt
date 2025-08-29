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
'<cpt_version>v1.3.0</cpt_version>
Option Explicit

Private Sub cboECF_Change()
  Call cptUpdateECF(Me.txtFilterECF)
End Sub

Sub cboLCF_Change()
  Call cptUpdateLCF(Me.txtFilterLCF)
End Sub

Private Sub chkAutoSwitch_Click()
  If Not Me.Visible Then Exit Sub
  If Me.cboLCF <> Me.lboECF.List(Me.lboECF.ListIndex, 2) Then
    Me.cboLCF = Me.lboECF.List(Me.lboECF.ListIndex, 2)
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
  Call cptHandleErr("cptSaveLocal_frm", "cmdCustomFields_Click", Err, Erl)
  Resume exit_here
  
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
  cptSaveLocal Me
End Sub

Private Sub cmdUnmap_Click()
  'objects
  Dim oTableField As Object
  Dim rstSavedMap As Object 'ADODB.Recordset
  'strings
  Dim strGUID As String
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
  
  'delete it from saved map
  If CLng(Left(Application.Build, 2)) < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) <> vbNullString Then
    Set rstSavedMap = CreateObject("ADODB.Recordset")
    rstSavedMap.Open strSavedMap
    rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & lngECF & " AND LCF=" & lngLCF
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
  If Me.optTasks Then
    For Each oTableField In ActiveProject.TaskTables(".cptSaveLocal Task Table").TableFields
      If oTableField.Field = lngECF Or oTableField.Field = lngLCF Then
        oTableField.Delete
      End If
    Next oTableField
    TableApply ".cptSaveLocal Task Table"
  ElseIf Me.optResources Then
    For Each oTableField In ActiveProject.ResourceTables(".cptSaveLocal Resource Table").TableFields
      If oTableField.Field = lngECF Or oTableField.Field = lngLCF Then
        oTableField.Delete
      End If
    Next oTableField
    TableApply ".cptSaveLocal Resource Table"
  End If
  
exit_here:
  On Error Resume Next
  Set oTableField = Nothing
  Set rstSavedMap = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "cmdUnmap", Err, Erl)
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

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "lblURL_Click", Err, Erl)
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
    If Me.ActiveControl.Name = "lboECF" Then
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
  Call cptHandleErr("cptSaveLocal_frm", "lboECF_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboECF_Click()
  'objects
  Dim oLookupTable As LookupTable
  'strings
  Dim strSwitch As String
  Dim strECF As String
  'longs
  Dim lngItems As Long
  Dim lngItem As Long
  Dim lngMax As Long
  Dim lngECF As Long
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
    
    If Me.chkAutoSwitch And Me.cboLCF.Value <> strSwitch Then
      Me.cboLCF.Value = strSwitch
      If Len(Me.lboECF.List(Me.lboECF.ListIndex, 3)) > 0 Then
        Me.lboLCF.Value = Me.lboECF.List(Me.lboECF.ListIndex, 3)
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
  Set oLookupTable = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "lboECF_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub optResources_Click()
  cptUpdateECF Me.txtFilterECF
  cptUpdateLCF Me.txtFilterLCF
  cptUpdateSaveLocalView Me
End Sub

Private Sub optTasks_Click()
  cptUpdateECF Me.txtFilterECF
  cptUpdateLCF Me.txtFilterLCF
  cptUpdateSaveLocalView Me
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
  Else
    Me.lboECF.MultiSelect = fmMultiSelectSingle
    Me.lboECF.ListIndex = 0
    Me.txtAutoMap.Visible = False
    Me.lboLCF.Visible = True
    Me.cmdAutoMap.Enabled = False
    Me.lblSelectAll.Visible = False
    Me.lblSelectNone.Visible = False
    Me.cmdMap.Enabled = True
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "tglAutoMap_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtFilterECF_Change()
  Call cptUpdateECF(Me, Me.txtFilterECF.Text)
End Sub

Private Sub txtFilterLCF_Change()
  Call cptUpdateLCF(Me, Me.txtFilterLCF.Text)
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
  If Len(strStartView) > 0 Then
    If Not ViewApply(strStartView, True) Then ViewApply "Gantt Chart"
  End If
  If Len(strStartTable) > 0 Then
    If strStartTable = ".cptSaveLocal Task Table" Then
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
  
  If ActiveProject.CurrentView = ".cptSaveLocal Task View" Then ViewApply "Gantt Chart"
  On Error Resume Next
  ActiveProject.Views(".cptSaveLocal Task View").Delete
  ActiveProject.TaskTables(".cptSaveLocal Task Table").Delete
  ActiveProject.Views(".cptSaveLocal Resource View").Delete
  ActiveProject.TaskTables(".cptSaveLocal Resource Table").Delete
  
exit_here:
  On Error Resume Next
  Unload Me
  
  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "UserForm_Terminate", Err, Erl)
  Resume exit_here
End Sub
