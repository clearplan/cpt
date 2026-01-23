VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptDataDictionary_frm 
   Caption         =   "IMS Data Dictionary"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
   OleObjectBlob   =   "cptDataDictionary_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptDataDictionary_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'<cpt_version>v1.5.2</cpt_version>
Option Explicit

Private Sub cboOpenWorkbooks_Change()
  If Not Me.cboOpenWorkbooks.Visible Then Exit Sub
  Select Case Me.cboOpenWorkbooks
    Case "Cancel"
      Me.cmdImport.SetFocus
      Me.cboOpenWorkbooks.Visible = False
    Case "Open another workbook..."
      Me.cmdImport.SetFocus
      Me.cboOpenWorkbooks.Visible = False
      cptImportDataDictionary Me
    Case "------------"
      Me.cmdImport.SetFocus
      Me.cboOpenWorkbooks.Visible = False
    Case Else
      Me.cmdImport.SetFocus
      Me.cboOpenWorkbooks.Visible = False
      cptImportDataDictionary Me, Me.cboOpenWorkbooks.Value
  End Select

End Sub

Private Sub chkIgnore_Click()
  'objects
  'strings
  Dim strProject As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If Me.ActiveControl.Name <> "chkIgnore" Then GoTo exit_here

  If IsNull(Me.lboCustomFields.Value) Then GoTo exit_here

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.lblStatus.Caption = "Saving..."

  strProject = cptGetProgramAcronym
  
  'find and update the record
  With CreateObject("ADODB.Recordset")
    .Open cptDir & "\settings\cpt-data-dictionary.adtg"
    .Filter = "PROJECT_NAME='" & strProject & "' AND FIELD_ID=" & CLng(Me.lboCustomFields.Value)
    If Not .EOF Then
      .Fields("IGNORE") = CBool(Me.chkIgnore)
      .Update
      Me.lboCustomFields.Column(3) = CBool(Me.chkIgnore)
    End If
    .Filter = ""
    .Save
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Me.lblStatus.Caption = "Ready..."
  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "chkIgnore_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdClose_Click()
  Me.Hide
End Sub

Private Sub cmdCustomFields_Click()
  'long
  Dim lngSelected As Long
  Dim lngItem As Long
  'string
  Dim strDescription As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not IsNull(Me.lboCustomFields.Value) Then
    lngSelected = Me.lboCustomFields.Value
  End If
  Me.lblStatus = "Refreshing..."
  Application.CustomizeField
  DoEvents
  Me.lboCustomFields.ListIndex = Null
  cptRefreshDictionary Me
  If lngSelected > 0 Then
    For lngItem = 0 To Me.lboCustomFields.ListCount - 1
      If Me.lboCustomFields.List(lngItem, 0) = lngSelected Then
        Me.lboCustomFields.SetFocus
        Me.lboCustomFields.Value = lngSelected
        Exit For
      End If
    Next lngItem
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "cmdCustomFields_Click()", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdExport_Click()
  cptExportDataDictionary Me
  'todo: refresh cboOpenWorkbooks
End Sub

Private Sub cmdFormGrow_Click()
  Me.cmdFormGrow.Visible = False
  Me.imgLogo.Visible = True
  Me.Height = 300
  Me.Width = 485.25
End Sub

Private Sub cmdFormShrink_Click()
  Me.cmdFormGrow.Visible = True
  Me.imgLogo.Visible = False
  Me.Height = 65
  Me.Width = 50
End Sub

Private Sub cmdImport_Click()
  'objects
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  'strings
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oExcel Is Nothing Then
    Me.cboOpenWorkbooks.Clear
    For lngItem = 1 To oExcel.Workbooks.Count
      Me.cboOpenWorkbooks.AddItem oExcel.Workbooks(lngItem).Name
    Next
    Me.cboOpenWorkbooks.AddItem "------------"
    Me.cboOpenWorkbooks.AddItem "Open another Workbook..."
    Me.cboOpenWorkbooks.AddItem "Cancel"
    
    Me.cboOpenWorkbooks.Visible = True
    Me.cboOpenWorkbooks.DropDown
  Else
    cptImportDataDictionary Me
  End If

exit_here:
  On Error Resume Next
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "cmdImport_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub lblRecover_Click()
  'todo: notify user of the issue
  'todo: create the workbook
  'todo: notify user to update and use 'import' to restore
End Sub

Private Sub lblURL_Click()
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "lblURL_Click()", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboCustomFields_AfterUpdate()
  Dim strMsg As String
  
  'todo: allow user to multi-select
  'todo: and ignore in bulk
  'todo: if selecteditems.count>1
  'todo: txtDescription = "< multiple fields selected >"
  'todo: and txtDescription.Enabled=False
  'todo: else proceed
  
  If Not IsNull(Me.lboCustomFields.Value) Then
    Me.txtDescription.Enabled = True
    Me.txtDescription = Me.lboCustomFields.Column(2)
    Me.chkIgnore.Enabled = True
    Me.chkIgnore = CBool(Me.lboCustomFields.Column(3))
    If Len(Me.lboCustomFields.Column(4)) > 0 Then
      Select Case Me.lboCustomFields.Column(4)
        Case "f"
          strMsg = "a formula"
        Case "p"
          strMsg = "a pick list"
        Case "fp"
          strMsg = "a formula and a pick list"
      End Select
      strMsg = FieldConstantToFieldName(Me.lboCustomFields.Column(0)) & " has " & strMsg & "."
      Me.lboCustomFields.Height = 108.75
      Me.lblAlert.Visible = True
      Me.lblAlert.Caption = strMsg
    Else
      Me.lblAlert.Caption = "-"
      Me.lblAlert.Visible = False
      Me.lboCustomFields.Height = 128.25
    End If
  Else
    Me.txtDescription.Value = ""
    Me.txtDescription.Enabled = False
    Me.chkIgnore.Enabled = False
  End If
End Sub

Private Sub txtDescription_AfterUpdate()
  'objects
  'strings
  Dim strProject As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If IsNull(Me.lboCustomFields.Value) Then
    Me.lblStatus.Caption = "Select a Custom Field"
    GoTo exit_here
  End If

  Me.lblStatus.Caption = "Saving..."

  strProject = cptGetProgramAcronym
  
  'find and update the record
  With CreateObject("ADODB.Recordset")
    .Open cptDir & "\settings\cpt-data-dictionary.adtg"
    .Filter = "PROJECT_NAME='" & strProject & "' AND FIELD_ID=" & CLng(Me.lboCustomFields.Value)
    If Not .EOF Then
      .Fields("DESCRIPTION") = Me.txtDescription.Text
      .Update
      Me.lboCustomFields.Column(2) = Me.txtDescription.Text
    End If
    .Filter = ""
    .Save
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Me.lblStatus.Caption = "Ready..."
  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "txtDescription_Change()", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtDescription_Change()
  Me.lblCharacterCount.Caption = 500 - Len(Me.txtDescription.Text)
End Sub

Private Sub txtDescription_Enter()
  Me.lblCharacterCount.Caption = 500 - Len(Me.txtDescription.Text)
  Me.lblCharacterCount.Visible = True
End Sub

Private Sub txtDescription_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  Me.lblCharacterCount.Visible = False
End Sub

Private Sub txtFilter_Change()
  'objects
  Dim rst As Object 'ADODB.Recordset
  'strings
  Dim strDictionary As String, strFilter As String, strText As String, strProject As String
  'longs
  Dim lngSelected As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strProject = cptGetProgramAcronym
  strDictionary = cptDir & "\settings\cpt-data-dictionary.adtg"
  
  If Not IsNull(Me.lboCustomFields.Value) Then
    lngSelected = Me.lboCustomFields.Value
  Else
    Me.txtDescription.Enabled = False
    Me.chkIgnore.Enabled = False
  End If
  
  If Dir(strDictionary) <> vbNullString Then
    Me.lboCustomFields.Clear
    Me.txtDescription = ""
    With CreateObject("ADODB.Recordset")
      .Open strDictionary
      .Sort = "CUSTOM_NAME"
      If Len(Me.txtFilter.Text) > 0 Then
        strText = cptRemoveIllegalCharacters(Me.txtFilter.Text)
        strFilter = "(CUSTOM_NAME LIKE '*" & strText & "*' AND PROJECT_NAME='" & strProject & "') "
        strFilter = strFilter & "OR "
        strFilter = strFilter & "(FIELD_NAME LIKE '*" & strText & "*' AND PROJECT_NAME='" & strProject & "') "
      Else
        strFilter = "PROJECT_NAME='" & strProject & "' "
      End If
      .Filter = strFilter
      If Not .EOF Then .MoveFirst
      lngItem = 0
      Do While Not .EOF
        Me.lboCustomFields.AddItem
        Me.lboCustomFields.Column(0, lngItem) = .Fields("FIELD_ID")
        Me.lboCustomFields.Column(1, lngItem) = .Fields("CUSTOM_NAME") & " (" & .Fields("FIELD_NAME") & ")"
        Me.lboCustomFields.Column(2, lngItem) = .Fields("DESCRIPTION")
        Me.lboCustomFields.Column(3, lngItem) = CBool(.Fields("IGNORE"))
        
        If Len(CustomFieldGetFormula(.Fields("FIELD_ID"))) > 0 Then
          Me.lboCustomFields.List(lngItem, 4) = "f"
        End If
        On Error Resume Next
        If Len(CustomFieldValueListGetItem(.Fields("FIELD_ID"), pjValueListValue, 1)) > 0 Then
          If Err.Number <> 1101 Then
            Me.lboCustomFields.List(lngItem, 4) = Me.lboCustomFields.List(lngItem, 4) & "p"
          End If
        End If
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        
        .MoveNext
        lngItem = lngItem + 1
      Loop
      .Close
    End With
  Else
    MsgBox "IMS Data Dictionary file not found!" & vbCrLf & "Please close and re-open the form to reset.", vbExclamation + vbOKOnly, "Error"
    GoTo exit_here
  End If
  
  For lngItem = 0 To Me.lboCustomFields.ListCount - 1
    If Me.lboCustomFields.List(lngItem, 0) = lngSelected Then
      Me.lboCustomFields.SetFocus
      Me.lboCustomFields.Value = lngSelected
      Me.txtFilter.SetFocus
      Exit For
    End If
  Next lngItem
  
  If IsNull(Me.lboCustomFields.Value) Then
    Me.txtDescription = ""
    Me.txtDescription.Enabled = False
    Me.chkIgnore.Value = False
    Me.chkIgnore.Enabled = False
  End If
  
  Me.lblStatus.Caption = Me.lboCustomFields.ListCount & " result" & IIf(Me.lboCustomFields.ListCount = 1, "", "s")

exit_here:
  On Error Resume Next
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "txtFilter_Change()", Err, Erl)
  Resume exit_here
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub UserForm_Terminate()
  Unload Me
End Sub
