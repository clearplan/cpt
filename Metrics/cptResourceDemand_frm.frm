VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptResourceDemand_frm 
   Caption         =   "Export Staffing Profile"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11865
   OleObjectBlob   =   "cptResourceDemand_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptResourceDemand_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.6.3</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptResourceDemand_frm"
Private Const adVarChar As Long = 200

Private Sub cboMonths_Change()
  If Me.Visible Then cptSaveSetting "ResourceDemand", "cboMonths", Me.cboMonths.Value
End Sub

Private Sub cboPrograms_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strFileName As String
  Dim strFieldName As String
  Dim strProgramAcronym As String
  'longs
  Dim lngItem As Long
  Dim lngConstant As Long
  Dim lngResponse As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnOverwrite As Boolean
  'variants
  'dates
  
  If Not Me.Visible Then GoTo exit_here
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Me.lboExport.ListCount > 0 Then
    lngResponse = MsgBox("Overwrite existing list?" & vbCrLf & vbCrLf & "(Click No to append.)", vbQuestion + vbYesNoCancel, "Do what now?")
    If lngResponse = vbCancel Then
      GoTo exit_here
    ElseIf lngResponse = vbYes Then
      blnOverwrite = True
    ElseIf lngResponse = vbNo Then
      blnOverwrite = False
    End If
  End If
  strDir = cptDir
  strFileName = strDir & "\settings\cpt-export-resource-userfields.adtg"
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strFileName, , adOpenKeyset, adLockOptimistic
  strProgramAcronym = cptGetProgramAcronym
  With oRecordset
    If blnOverwrite Then
      'clear listbox
      Me.lboExport.Clear
    End If
    'don't filter that's weird
    .Filter = 0
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        If .Fields(2) = Me.cboPrograms.Value Then
          lngConstant = .Fields(0)
          strFieldName = .Fields(1)
          Dim strImportedFieldName As String
          Dim strCustomFieldName As String
          If lngConstant >= 188776000 Then 'enterprise
            strCustomFieldName = FieldConstantToFieldName(lngConstant) & " (Enterprise)"
          Else
            strCustomFieldName = CustomFieldGetName(lngConstant)
            If Len(strCustomFieldName) = 0 Then strCustomFieldName = FieldConstantToFieldName(lngConstant)
            strCustomFieldName = strCustomFieldName & " (" & FieldConstantToFieldName(lngConstant) & ")"
          End If
          If strCustomFieldName <> strFieldName Then
            Dim strMsg As String
            strMsg = "In '" & oRecordset(2) & "'" & vbCrLf
            strMsg = strMsg & FieldConstantToFieldName(lngConstant) & " is named [" & strFieldName & "]" & vbCrLf & vbCrLf
            strMsg = strMsg & "In '" & strProgramAcronym & "'" & vbCrLf
            strMsg = strMsg & FieldConstantToFieldName(lngConstant) & " is named [" & strCustomFieldName & "]" & vbCrLf & vbCrLf
            strMsg = strMsg & "Import [" & FieldConstantToFieldName(lngConstant) & "] anyway?" & vbCrLf & vbCrLf
            strMsg = strMsg & "(It will not be renamed.)"
            If MsgBox(strMsg, vbExclamation + vbYesNo, "Fields Don't Match!") = vbNo Then GoTo next_import_field
          End If
          'ensure not already exists
          For lngItem = 0 To Me.lboExport.ListCount - 1
            If Me.lboExport.List(lngItem, 0) = .Fields(0) Then GoTo next_import_field
          Next lngItem
          'add to lbo
          Me.lboExport.AddItem
          Me.lboExport.List(Me.lboExport.ListCount - 1, 0) = lngConstant
          Me.lboExport.List(Me.lboExport.ListCount - 1, 1) = FieldConstantToFieldName(lngConstant) 'overwrite with local CFN
          'note: can't addnew here bc it resets the index
        End If
next_import_field:
        .MoveNext
      Loop
    End If
    'filter for current program
    .Filter = "PROGRAM='" & strProgramAcronym & "'"
    If Not .EOF Then
      'delete them
      .MoveFirst
      Do While Not .EOF
        .Delete adAffectCurrent
        .MoveNext
      Loop
    End If
    'reset the filter
    .Filter = 0
    For lngItem = 0 To Me.lboExport.ListCount - 1
      lngConstant = Me.lboExport.List(lngItem, 0)
      strFieldName = Me.lboExport.List(lngItem, 1)
      .AddNew Array(0, 1, 2), Array(lngConstant, strFieldName, strProgramAcronym) 'resets the index
    Next lngItem
  End With
     
exit_here:
  On Error Resume Next
  oRecordset.Filter = 0
  oRecordset.Save strFileName, adPersistADTG
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cboPrograms_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboWeekday_Change()
  If Me.Visible And Me.cboWeekday.Value <> "" Then cptSaveSetting "ResourceDemand", "cboWeekday", Me.cboWeekday.Value
End Sub

Private Sub cboWeeks_Change()
  Me.cboWeekday.Clear
  Select Case Me.cboWeeks
    Case "Beginning"
      Me.cboWeekday.AddItem "Sunday"
      Me.cboWeekday.AddItem "Monday"
      Me.cboWeekday.Value = "Monday"
    Case "Ending"
      Me.cboWeekday.AddItem "Friday"
      Me.cboWeekday.AddItem "Saturday"
      Me.cboWeekday.Value = "Friday"
  End Select
  If Me.Visible Then cptSaveSetting "ResourceDemand", "cboWeeks", Me.cboWeeks.Value
End Sub

Private Sub chkAssociatedBaseline_Click()
  If Me.Visible Then
    cptSaveSetting "ResourceDemand", "chkAssociatedBaseline", IIf(Me.chkAssociatedBaseline, 1, 0)
    If Me.chkAssociatedBaseline And Me.chkFullBaseline Then Me.chkFullBaseline = False
  End If
End Sub

Private Sub chkCosts_AfterUpdate()
  Me.chkA.Enabled = CBool(Me.chkCosts.Value)
  Me.chkB.Enabled = CBool(Me.chkCosts.Value)
  Me.chkC.Enabled = CBool(Me.chkCosts.Value)
  Me.chkD.Enabled = CBool(Me.chkCosts.Value)
  Me.chkE.Enabled = CBool(Me.chkCosts.Value)
  If Me.Visible Then cptSaveSetting "ResourceDemand", "chkCosts", IIf(Me.chkCosts, 1, 0)
End Sub

Private Sub chkFullBaseline_Click()
  If Me.Visible Then
    cptSaveSetting "ResourceDemand", "chkFullBaseline", IIf(Me.chkFullBaseline, 1, 0)
    If Me.chkAssociatedBaseline And Me.chkFullBaseline Then Me.chkAssociatedBaseline = False
  End If
End Sub

Private Sub cmdAdd_Click()
  Dim lngField As Long, lngExport As Long, lngExists As Long
  Dim blnExists As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngField = 0 To Me.lboFields.ListCount - 1
    If Me.lboFields.Selected(lngField) Then
      'ensure doesn't already exist
      blnExists = False
      For lngExists = 0 To Me.lboExport.ListCount - 1
        If Me.lboExport.List(lngExists, 0) = Me.lboFields.List(lngField) Then
          GoTo next_item
        End If
      Next lngExists
      Me.lboExport.AddItem
      lngExport = Me.lboExport.ListCount - 1
      Me.lboExport.List(lngExport, 0) = Me.lboFields.List(lngField, 0)
      Me.lboExport.List(lngExport, 1) = Me.lboFields.List(lngField, 1)
    End If
next_item:
  Next lngField

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cmdAdd_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdCancel_Click()
  Dim strFileName As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  Unload Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cmdCancel_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdExport_Click()
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptExportResourceDemand Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cmdExport_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdRemove_Click()
  Dim lngExport As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If Me.lboExport.Selected(lngExport) Then
      Me.lboExport.RemoveItem lngExport
    End If
  Next lngExport

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cmdRemove_Click", Err, Erl)
  Resume exit_here
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

Private Sub lboExport_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFileName As String
  Dim strFieldName As String
  Dim strProgramAcronym As String
  'longs
  Dim lngExport As Long
  Dim lngConstant As Long
  'integers
  'doubles
  'booleans
  Dim blnExists As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strProgramAcronym = cptGetProgramAcronym
  Set oRecordset = CreateObject("ADODB.Recordset")
  strFileName = cptDir & "/settings/cpt-export-resource-userfields.adtg"
  oRecordset.Open strFileName, , adOpenKeyset, adLockOptimistic

  If Not IsNull(Me.lboExport.Value) Then
    lngConstant = Me.lboExport.List(Me.lboExport.ListIndex, 0)
    strFieldName = Me.lboExport.List(Me.lboExport.ListIndex, 1)
    oRecordset.Filter = 0
    oRecordset.Filter = "PROGRAM='" & strProgramAcronym & "' AND [Field Constant]='" & CStr(lngConstant) & "'"
    If Not oRecordset.EOF Then
      oRecordset.Delete adAffectCurrent
    End If
    For lngExport = Me.lboExport.ListCount - 1 To 0 Step -1
      If Me.lboExport.Selected(lngExport) Then
        Me.lboExport.RemoveItem lngExport
        Exit For
      End If
    Next lngExport
  End If
  
  oRecordset.Filter = 0
  oRecordset.Save strFileName, adPersistADTG

exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "lboExport_DblClick", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboFields_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFileName As String
  Dim strFieldName As String
  Dim strProgramAcronym As String
  'longs
  'Dim lngField As Long
  Dim lngExport As Long
  Dim lngConstant As Long
  'integers
  'doubles
  'booleans
  Dim blnExists As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strProgramAcronym = cptGetProgramAcronym
  Set oRecordset = CreateObject("ADODB.Recordset")
  strFileName = cptDir & "/settings/cpt-export-resource-userfields.adtg"
  oRecordset.Open strFileName, , adOpenKeyset, adLockOptimistic

  If Not IsNull(Me.lboFields.Value) Then
    lngConstant = Me.lboFields.List(Me.lboFields.ListIndex, 0)
    strFieldName = Me.lboFields.List(Me.lboFields.ListIndex, 1)
    'ensure doesn't already exist
    blnExists = False
    For lngExport = 0 To Me.lboExport.ListCount - 1
      If Me.lboExport.List(lngExport, 0) = lngConstant Then
        blnExists = True
      End If
    Next lngExport
    If Not blnExists Then
      Me.lboExport.AddItem
      Me.lboExport.List(lngExport, 0) = lngConstant
      Me.lboExport.List(lngExport, 1) = strFieldName
      oRecordset.AddNew Array(0, 1, 2), Array(lngConstant, strFieldName, strProgramAcronym)
    End If
  End If
  oRecordset.Filter = 0
  oRecordset.Save strFileName, adPersistADTG

exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "lboFields_DblClick", Err, Erl)
  Resume exit_here

End Sub

Private Sub stxtSearch_Change()
  'objects
  'strings
  Dim strFileName As String
  'longs
  Dim lngItem As Long
  'integers
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboFields.Clear

  strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
  With CreateObject("ADODB.Recordset")
    .Open strFileName
    If Len(Me.stxtSearch.Text) > 0 Then
      .Filter = "CUSTOM_NAME LIKE '*" & cptRemoveIllegalCharacters(Me.stxtSearch.Text) & "*'"
    Else
      .Filter = 0
    End If
    If .RecordCount > 0 Then .MoveFirst
    lngItem = 0
    Do While Not .EOF
      Me.lboFields.AddItem
      Me.lboFields.List(lngItem, 0) = .Fields(0)
      If .Fields(0) >= 188776000 Then 'enterprise
        Me.lboFields.List(lngItem, 1) = .Fields(1) & " (Enterprise)"
      Else
        Me.lboFields.List(lngItem, 1) = .Fields(1) & " (" & FieldConstantToFieldName(.Fields(0)) & ")"
      End If
      .MoveNext
      lngItem = lngItem + 1
    Loop
    .Close
    Me.lblStatus.Caption = lngItem & " record" & IIf(lngItem = 1, "", "s") & " found."
  End With
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "stxtSearch_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub
