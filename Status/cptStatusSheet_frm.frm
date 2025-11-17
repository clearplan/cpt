VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptStatusSheet_frm 
   Caption         =   "Create Status Sheets"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12435
   OleObjectBlob   =   "cptStatusSheet_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptStatusSheet_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.7.0</cpt_version>
Option Explicit
Private Const lngForeColorValid As Long = -2147483630
Private Const lngBorderColorValid As Long = 8421504 '-2147483642
Private Const lngForeColorInvalid As Long = 192
Private Const lngBorderColorInvalid As Long = 192

Private Sub cboCostTool_Change()
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'setup EVT dictionary
  Set oEVTs = CreateObject("Scripting.Dictionary")
  If Not oEVTs Is Nothing Then
    If oEVTs.Count > 0 Then oEVTs.RemoveAll
  End If
  If Not IsNull(Me.cboCostTool.Value) Then
    If Me.cboCostTool.Value = "COBRA" Then
      oEVTs.Add "A", "Level of Effort"
      oEVTs.Add "B", "Milestones"
      oEVTs.Add "C", "% Complete"
      oEVTs.Add "D", "Units Complete"
      oEVTs.Add "E", "50-50"
      oEVTs.Add "F", "0-100"
      oEVTs.Add "G", "100-0"
      oEVTs.Add "H", "User Defined"
      oEVTs.Add "J", "Apportioned"
      oEVTs.Add "K", "Planning Package"
      oEVTs.Add "L", "Assignment % Complete"
      oEVTs.Add "M", "Calculated Apportionment"
      oEVTs.Add "N", "Steps"
      oEVTs.Add "O", "Earned As Spent"
      oEVTs.Add "P", "% Complete Manual Entry"
    ElseIf Me.cboCostTool.Value = "MPM" Then
      oEVTs.Add "0", "No EVM required"
      oEVTs.Add "1", "0/100"
      oEVTs.Add "2", "25/75"
      oEVTs.Add "3", "40/60"
      oEVTs.Add "4", "50/50"
      oEVTs.Add "5", "% Complete"
      oEVTs.Add "6", "LOE"
      oEVTs.Add "7", "Earned Standards"
      oEVTs.Add "8", "Milestone Weights"
      oEVTs.Add "9", "BCWP Entry"
      oEVTs.Add "A", "Apportioned"
      oEVTs.Add "P", "Milestone Weights with % Complete"
      oEVTs.Add "K", "Key Event"
    End If
  End If
  
  If Not IsNull(Me.cboCostTool) Then
    Me.lblCostTool.ForeColor = lngForeColorValid
    Me.cboCostTool.BorderColor = lngBorderColorValid
  End If
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboCostTool_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboCreate_Change()
  'objects
  'strings
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Select Case Me.cboCreate
    Case 0 'A single workbook
      Me.lboItems.ForeColor = lngForeColorValid
      Me.chkSendEmails.Caption = "Create Email"
      'Me.chkLocked.Caption = "Protect Workbook"
      'Me.lblForEach.Visible = False
      Me.cboEach.Enabled = False
      Me.lboItems.Enabled = False
      Me.chkAllItems = False
      Me.chkAllItems.Enabled = False
      FilterClear
      If Not cptFilterExists("cptStatusSheet Filter") Then
        cptRefreshStatusTable Me
      End If
      
    Case 1 'A worksheet for each
      Me.lboItems.ForeColor = lngForeColorValid
      Me.chkSendEmails.Caption = "Create Email"
      'Me.chkLocked.Caption = "Protect Worksheets"
      'Me.lblForEach.Visible = True
      Me.cboEach.Enabled = True
      Me.lboItems.Enabled = True
      Me.chkAllItems.Enabled = True
      If Me.Visible Then Me.cboEach.DropDown

    Case 2 'A workbook for each
      Me.lboItems.ForeColor = lngForeColorValid
      Me.chkSendEmails.Caption = "Create Emails"
      'Me.chkLocked.Caption = "Protect Workbooks"
      'Me.lblForEach.Visible = True
      Me.cboEach.Enabled = True
      Me.lboItems.Enabled = True
      Me.chkAllItems.Enabled = True
      If Me.Visible Then Me.cboEach.DropDown

    End Select
        
    If Not IsNull(Me.cboCreate) Then
      Me.lblCreate.ForeColor = lngForeColorValid
      Me.cboCreate.BorderColor = lngBorderColorValid
      If CLng(Me.cboCreate.Value) > 0 Then
        If Me.cboEach = "" Then
          Me.lblForEach.ForeColor = lngForeColorInvalid
          Me.cboEach.BorderColor = lngBorderColorInvalid
        Else
          Me.lblForEach.ForeColor = lngForeColorValid
          Me.cboEach.BorderColor = lngBorderColorValid
        End If
      Else
        Me.lblForEach.ForeColor = lngForeColorValid
        Me.cboEach.BorderColor = lngBorderColorValid
      End If
    End If
        
exit_here:
  On Error Resume Next
  Me.Repaint
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboCreate_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEach_Change()
  'objects
  Dim rstItems As Object 'ADODB.Recordset
  Dim oTask As MSProject.Task
  'strings
  Dim strFieldName As String
  'longs
  Dim lngItem As Long
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboItems.Clear
  Me.lboItems.ForeColor = lngForeColorValid
  If Me.Visible Then
    ActiveWindow.TopPane.Activate
    FilterApply "cptStatusSheet Filter"
  End If
  
  On Error Resume Next
  lngField = FieldNameToFieldConstant(Me.cboEach)
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If lngField > 0 Then
    Set rstItems = CreateObject("ADODB.Recordset")
    With rstItems
      .Fields.Append Me.cboEach.Value, 200, 255 '200=adVarChar
      .Open
      For Each oTask In ActiveProject.Tasks
        If oTask Is Nothing Then GoTo next_task
        If Not oTask.Active Then GoTo next_task
        If oTask.ExternalTask Then GoTo next_task
        If oTask.Summary Then GoTo next_task
        If IsDate(oTask.ActualFinish) Then GoTo next_task
        If Len(oTask.GetField(lngField)) > 0 Then
          If .RecordCount > 0 Then .MoveFirst
          .Find "[" & Me.cboEach.Value & "]='" & oTask.GetField(lngField) & "'"
          If .EOF Then
            .AddNew Array(0), Array(oTask.GetField(lngField))
          End If
        End If
next_task:
      Next oTask
      'validate field has items
      If .RecordCount = 0 Then
        If Len(CustomFieldGetName(lngField)) > 0 Then
          strFieldName = CustomFieldGetName(lngField)
        Else
          strFieldName = FieldConstantToFieldName(lngField)
        End If
        MsgBox "The field '" & strFieldName & "' contains no values.", vbExclamation + vbOKOnly, "Invalid Selection"
      Else
        .Sort = "[" & Me.cboEach.Value & "]"
        .MoveFirst
        Do While Not .EOF
          Me.lboItems.AddItem .Fields(0)
          .MoveNext
        Loop
      End If
      Me.txtFileName_Change
    End With
  End If 'lngField > 0
  
  If Me.cboEach.Value <> "" Then
    Me.lblForEach.ForeColor = lngForeColorValid
    Me.cboEach.BorderColor = lngBorderColorValid
  End If
  
exit_here:
  On Error Resume Next
  Me.Repaint
  If rstItems.State = 1 Then rstItems.Close
  Set rstItems = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboEach_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub chkAllItems_Click()
  Dim lngItem As Long
  Dim strCriteria As String
  Dim strFieldName As String
  
  If IsNull(Me.cboEach) Or Me.lboItems.ListCount = 0 Then Exit Sub
  strFieldName = Me.cboEach.Value
  If Me.chkAllItems Then
    For lngItem = 0 To Me.lboItems.ListCount - 1
      Me.lboItems.Selected(lngItem) = True
      strCriteria = strCriteria & Me.lboItems.List(lngItem) & Chr$(9)
    Next lngItem
    If Len(strCriteria) = 0 Then Exit Sub
    strCriteria = Left(strCriteria, Len(strCriteria) - 1)
    SetAutoFilter FieldName:=strFieldName, FilterType:=pjAutoFilterIn, Criteria1:=strCriteria
  Else
    If Me.ActiveControl.Name = Me.chkAllItems.Name Then
      For lngItem = 0 To Me.lboItems.ListCount - 1
        Me.lboItems.Selected(lngItem) = False
      Next lngItem
      SetAutoFilter FieldName:=strFieldName, FilterType:=pjAutoFilterClear
    End If
  End If
  
End Sub

Private Sub chkAppendStatusDate_Click()
  If Me.chkAppendStatusDate Then
    Me.lblDirSample.Caption = Me.txtDir & Format(CDate(Me.txtStatusDate), "yyyy-mm-dd") & "\"
  Else
    Me.lblDirSample.Caption = Me.txtDir
  End If
End Sub

Private Sub chkAssignments_Click()
  Dim strAllowAssignmentNotes As String
  
  If Not Me.Visible Then GoTo exit_here
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Me.chkAssignments Then
    Me.chkAllowAssignmentNotes.Enabled = True
    strAllowAssignmentNotes = cptGetSetting("StatusSheet", "chkAllowAssignmentNotes")
    If strAllowAssignmentNotes <> "" Then
      Me.chkAllowAssignmentNotes.Value = CBool(strAllowAssignmentNotes)
    Else
      Me.chkAllowAssignmentNotes.Value = False
    End If
  Else
    Me.chkAllowAssignmentNotes.Value = False
    Me.chkAllowAssignmentNotes.Enabled = False
  End If
  
  Call cptRefreshStatusTable(Me, True)
  
exit_here:
  On Error Resume Next
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "chkAssignments_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub chkConditionalFormatting_Click()
  Dim strConditionalFormattingLegend As String
  If Me.chkConditionalFormatting Then
    Me.chkConditionalFormattingLegend.Enabled = True
    strConditionalFormattingLegend = cptGetSetting("StatusSheet", "chkConditionalFormattingLegend")
    If Len(strConditionalFormattingLegend) > 0 Then
      Me.chkConditionalFormattingLegend = CBool(strConditionalFormattingLegend)
    Else
      Me.chkConditionalFormattingLegend = True 'default
    End If
  Else
    Me.chkConditionalFormattingLegend = False
    Me.chkConditionalFormattingLegend.Enabled = False
  End If
End Sub

Private Sub chkHide_Click()

  If Not Me.Visible Then GoTo exit_here
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.txtHideCompleteBefore.Enabled = Me.chkHide
  Call cptRefreshStatusTable(Me, False, True)
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "chkHide_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub chkIgnoreLOE_Click()
  Call cptRefreshStatusTable(Me, False, True)
End Sub

Private Sub chkKeepOpen_Click()
  If Me.chkKeepOpen Then
    Me.chkSendEmails = False
    Me.chkSendEmails.Enabled = False
  Else
    Me.chkSendEmails.Enabled = True
  End If
End Sub

Private Sub chkLookahead_Click()
  If Me.chkLookahead Then
    Me.txtLookaheadDays.Enabled = True
    Me.txtLookaheadDate.Enabled = True
    Me.txtLookaheadDays.SetFocus
  Else
    Me.txtLookaheadDays = ""
    Me.txtLookaheadDays.Enabled = False
    Me.txtLookaheadDate = ""
    Me.txtLookaheadDate.Enabled = False
    Call cptRefreshStatusTable(Me, False, True)
  End If
End Sub

Private Sub chkSendEmails_Click()
  Dim strQuickPart As String, strSubject As String, strCC As String, strKeepOpen As String
  Dim blnExists As Boolean
  Dim lngItem As Long

  Me.txtSubject.Enabled = Me.chkSendEmails
  Me.txtCC.Enabled = Me.chkSendEmails
  Me.cboQuickParts.Clear
  Me.cboQuickParts.Enabled = Me.chkSendEmails
  If Me.chkSendEmails Then
    Me.chkKeepOpen = False
    Me.chkKeepOpen.Enabled = False
    strSubject = cptGetSetting("StatusSheet", "txtSubject")
    If Len(strSubject) > 0 Then
      Me.txtSubject = strSubject
    End If
    strCC = cptGetSetting("StatusSheet", "txtCC")
    If Len(strCC) > 0 Then
      Me.txtCC = strCC
    End If
    Call cptListQuickParts(Me, True)
    strQuickPart = cptGetSetting("StatusSheet", "cboQuickPart")
    If Len(strQuickPart) > 0 And Me.cboQuickParts.ListCount > 0 Then
      blnExists = False
      For lngItem = 0 To Me.cboQuickParts.ListCount - 1
        If Me.cboQuickParts.List(lngItem, 0) = strQuickPart Then
          Me.cboQuickParts.Value = strQuickPart
          blnExists = True
          Exit For
        End If
      Next lngItem
      If Not blnExists Then
        MsgBox "QuickPart '" & strQuickPart & "' not found.", vbExclamation + vbOKOnly, "Stored Setting Invalid"
      End If
    End If
  Else
    Me.chkKeepOpen.Enabled = True
  End If

End Sub

Sub cmdAdd_Click()
  'objects
  'strings
  Dim strEVT As String
  Dim strEVP As String
  'longs
  Dim lngField As Long
  Dim lngExport As Long
  Dim lngExists As Long
  Dim lngEVT As Long
  Dim lngEVP As Long
  'integers
  'doubles
  'booleans
  Dim blnExists As Boolean
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'do not allow addition of EVT or EVP - get settings
  lngEVT = CLng(Split(cptGetSetting("Integration", "EVT"), "|")(0))
  If Len(CustomFieldGetName(lngEVT)) > 0 Then
    strEVT = CustomFieldGetName(lngEVT)
  Else
    strEVT = FieldConstantToFieldName(lngEVT)
  End If
  lngEVP = CLng(Split(cptGetSetting("Integration", "EVP"), "|")(0))
  If Len(CustomFieldGetName(lngEVP)) > 0 Then
    strEVP = CustomFieldGetName(lngEVP)
  Else
    strEVP = FieldConstantToFieldName(lngEVP)
  End If
  
  For lngField = 0 To Me.lboFields.ListCount - 1
    If Me.lboFields.Selected(lngField) Then
      'do not allow EVT
      If CLng(Me.lboFields.List(lngField)) = lngEVT Then
        MsgBox "The EVT Field ('" & strEVT & "') is automatically included.", vbInformation + vbOKOnly, "EVT Rejected"
        GoTo next_item
      End If
      'do not allow EVP
      If CLng(Me.lboFields.List(lngField)) = lngEVP Then
        MsgBox "The EVP Field ('" & strEVP & "') is automatically included.", vbInformation + vbOKOnly, "EVP Rejected"
        GoTo next_item
      End If
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
      Me.lboExport.List(lngExport, 2) = Me.lboFields.List(lngField, 2)
    End If
next_item:
  Next lngField

  cptRefreshStatusTable Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdAdd_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdAddAll_Click()
  Dim lngField As Long, lngExport As Long, lngExists As Long
  Dim blnExists As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngField = 0 To Me.lboFields.ListCount - 1
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
    Me.lboExport.List(lngExport, 2) = Me.lboFields.List(lngField, 2)
next_item:
  Next lngField

  cptRefreshStatusTable Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdAddAll_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdCancel_Click()
  Dim strFileName As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = cptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  If Not oEVTs Is Nothing Then oEVTs.RemoveAll
  Set oEVTs = Nothing
  Me.Hide

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdCancel_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdDown_Click()
  Dim lngExport As Long
  Dim lngField As Long, strField As String, strField2 As String
  Dim blnSelected As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnSelected = False
  For lngExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If lngExport < Me.lboExport.ListCount - 1 Then
      If Me.lboExport.Selected(lngExport) Then
        blnSelected = True
        'capture values
        lngField = Me.lboExport.List(lngExport + 1, 0)
        strField = Me.lboExport.List(lngExport + 1, 1)
        strField2 = Me.lboExport.List(lngExport + 1, 2)
        'move selected values
        Me.lboExport.List(lngExport + 1, 0) = Me.lboExport.List(lngExport, 0)
        Me.lboExport.List(lngExport + 1, 1) = Me.lboExport.List(lngExport, 1)
        Me.lboExport.List(lngExport + 1, 2) = Me.lboExport.List(lngExport, 2)
        Me.lboExport.Selected(lngExport + 1) = True
        Me.lboExport.List(lngExport, 0) = lngField
        Me.lboExport.List(lngExport, 1) = strField
        Me.lboExport.List(lngExport, 2) = strField2
        Me.lboExport.Selected(lngExport) = False
      End If
    End If
  Next lngExport

  If blnSelected Then cptRefreshStatusTable Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdDown_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdDn_Click()
  Dim lngExport As Long
  Dim lngField As Long, strField As String, strField2 As String
  Dim blnSelected As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnSelected = False
  For lngExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If lngExport < Me.lboExport.ListCount - 1 Then
      If Me.lboExport.Selected(lngExport) Then
        blnSelected = True
        'capture values
        lngField = Me.lboExport.List(lngExport + 1, 0)
        strField = Me.lboExport.List(lngExport + 1, 1)
        strField2 = Me.lboExport.List(lngExport + 1, 2)
        'move selected values
        Me.lboExport.List(lngExport + 1, 0) = Me.lboExport.List(lngExport, 0)
        Me.lboExport.List(lngExport + 1, 1) = Me.lboExport.List(lngExport, 1)
        Me.lboExport.List(lngExport + 1, 2) = Me.lboExport.List(lngExport, 2)
        Me.lboExport.Selected(lngExport + 1) = True
        Me.lboExport.List(lngExport, 0) = lngField
        Me.lboExport.List(lngExport, 1) = strField
        Me.lboExport.List(lngExport, 2) = strField2
        Me.lboExport.Selected(lngExport) = False
      End If
    End If
  Next lngExport

  If blnSelected Then cptRefreshStatusTable Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdDn_Click", Err, Erl)
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

  cptRefreshStatusTable Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdRemove_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdRemoveAll_Click()
  Dim lngExport As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngExport = Me.lboExport.ListCount - 1 To 0 Step -1
    Me.lboExport.RemoveItem lngExport
  Next lngExport

  cptRefreshStatusTable Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdRemoveAll_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdRun_Click()
  'objects
  'strings
  Dim strTempDir As String
  Dim strMsg As String
  'longs
  Dim lngResponse As Long
  Dim lngEVP As Long
  Dim lngEVT As Long
  Dim lngDateFormat As Long
  Dim lngSelectedItems As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnError As Boolean
  Dim blnIncluded As Boolean
  'variants
  Dim vResponse As Variant
  Dim vDir As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  lngDateFormat = Application.DefaultDateFormat
  blnError = False

  Me.lblStatusDate.ForeColor = lngForeColorValid
  Me.txtStatusDate.BorderColor = lngBorderColorValid
  Me.chkHide.ForeColor = lngForeColorValid
  Me.lblStatus.ForeColor = lngForeColorValid
  Me.lblCostTool.ForeColor = lngForeColorValid
  Me.cboCostTool.BorderColor = lngBorderColorValid
  Me.lblCreate.ForeColor = lngForeColorValid
  Me.cboCreate.BorderColor = lngBorderColorValid
  Me.lblForEach.ForeColor = lngForeColorValid
  Me.cboEach.BorderColor = lngBorderColorValid
  Me.lblDirectory.ForeColor = lngForeColorValid
  Me.txtDir.BorderColor = lngBorderColorValid
  Me.lblNamingConvention.ForeColor = lngForeColorValid
  Me.txtFileName.BorderColor = lngBorderColorValid
  Me.lboItems.BorderColor = lngBorderColorValid
  Me.lblIncludeItems.ForeColor = lngForeColorValid
  
  lngEVT = CLng(Split(cptGetSetting("Integration", "EVT"), "|")(0))
  lngEVP = CLng(Split(cptGetSetting("Integration", "EVP"), "|")(0))
  
  'validation
  If Not IsDate(Me.txtStatusDate.Value) Then
    Me.lblStatusDate.ForeColor = lngForeColorInvalid
    Me.txtStatusDate.BorderColor = lngBorderColorInvalid
    blnError = True
  ElseIf IsDate(Me.txtStatusDate.Value) Then
    If CDate(Me.txtStatusDate.Value) < #1/1/1984# Then
      Me.lblStatusDate.ForeColor = lngForeColorInvalid
      Me.txtStatusDate.BorderColor = lngBorderColorInvalid
      blnError = True
    End If
  End If
  If Me.chkHide.Value = True Then
    If Not IsDate(Me.txtHideCompleteBefore.Value) Then
      Me.chkHide.ForeColor = lngForeColorInvalid
      blnError = True
    ElseIf IsDate(Me.txtHideCompleteBefore.Value) Then
      If CDate(Me.txtHideCompleteBefore.Value) < #1/1/1984# Then
        Me.chkHide.ForeColor = lngForeColorInvalid
        blnError = True
      End If
    End If
  End If
  If Len(Me.cboCostTool.Value) = 0 Then
    Me.lblCostTool.ForeColor = lngForeColorInvalid
    Me.cboCostTool.BorderColor = lngBorderColorInvalid
    blnError = True
  End If
  'hide complete before must be prior to or equal to status date
  If IsDate(Me.txtStatusDate.Value) And IsDate(Me.txtHideCompleteBefore.Value) Then
    If CDate(Me.txtHideCompleteBefore.Value) > CDate(Me.txtStatusDate.Value) Then
      MsgBox "'Hide Complete Before' date must be prior to, or equal to, status date.", vbExclamation + vbOKOnly, "Invalid Hide Complete Before Date"
      Me.chkHide.ForeColor = lngForeColorInvalid
      blnError = True
    End If
  End If
  'ensure selections
  If IsNull(Me.cboCreate) Then
    Me.cboCreate.BorderColor = lngBorderColorInvalid
    Me.lblCreate.ForeColor = lngForeColorInvalid
    blnError = True
  Else
    'ensure each if necessary
    If Me.cboCreate.Value <> "0" Then
      'a limiting field must be selected
      If Me.cboEach.Value = 0 Or Me.cboEach.Value = "" Then
        Me.cboEach.BorderColor = lngBorderColorInvalid
        Me.lblForEach.ForeColor = lngForeColorInvalid
        blnError = True
      Else
        'at least one item selected
        For lngItem = 0 To Me.lboItems.ListCount - 1
          If Me.lboItems.Selected(lngItem) Then lngSelectedItems = lngSelectedItems + 1
        Next lngItem
        If lngSelectedItems = 0 Then
          Me.lboItems.BorderColor = lngBorderColorInvalid
          Me.lblIncludeItems.ForeColor = lngForeColorInvalid
          blnError = True
        End If
        'the limiting field should be included in the export list
        blnIncluded = False
        For lngItem = 0 To Me.lboExport.ListCount - 1
          If Me.lboExport.List(lngItem, 1) = Me.cboEach Then blnIncluded = True
        Next lngItem
        If Not blnIncluded Then
          If MsgBox("The For Each field '" & Me.cboEach & "' is not included in the export list." & vbCrLf & vbCrLf & "Include it?", vbYesNo + vbQuestion, "Include For Each Field?") = vbYes Then
            For lngItem = 0 To Me.lboFields.ListCount - 1
              Me.lboFields.Selected(lngItem) = Me.lboFields.List(lngItem, 1) = Me.cboEach
            Next lngItem
            Me.cmdAdd_Click
          End If
        End If
      End If
    End If
    'ensure unique filenames
    If Me.cboCreate.Value = "0" Then 'one workbook
      If InStr(Me.txtFileName, "[item]") > 0 Then
        Me.lblNamingConvention.ForeColor = lngForeColorInvalid
        Me.txtFileName.BorderColor = lngForeColorInvalid
        MsgBox "Cannot use '[item]' in naming convention when creating a single workbook.", vbExclamation + vbOKOnly, "Invalid Naming Convention"
        blnError = True
      End If
    ElseIf Me.cboCreate.Value = "1" Then 'worksheet for each
      If InStr(Me.txtFileName, "[item]") > 0 Then
        Me.lblNamingConvention.ForeColor = lngForeColorInvalid
        Me.txtFileName.BorderColor = lngForeColorInvalid
        MsgBox "Cannot use '[item]' in naming convention when creating worksheet for each.", vbExclamation + vbOKOnly, "Invalid Naming Convention"
        blnError = True
      End If
    ElseIf Me.cboCreate.Value = "2" Then 'workbook for each
      If InStr(Me.txtFileName, "[item]") = 0 Then
        Me.lblNamingConvention.ForeColor = lngForeColorInvalid
        Me.txtFileName.BorderColor = lngForeColorInvalid
        If Len(Me.txtFileName) > 0 Then
          strMsg = Me.txtFileName & "_[item]"
        Else
          strMsg = cptGetProgramAcronym & "_Status_[YYYY-MM-DD]_[item]"
        End If
        strMsg = InputBox("Must include '[item]' in naming convention when creating workbook for each." & vbCrLf & vbCrLf & "Example (click 'OK' to accept):", "Invalid Naming Convention", strMsg)
        If StrPtr(strMsg) = 0 Then 'user hit cancel
          blnError = True
        Else
          Me.txtFileName.Value = strMsg
        End If
      End If
    End If
  End If
  'ensure directory exists
  If Dir(Me.txtDir, vbDirectory) = vbNullString Then
    For Each vDir In Split(Me.txtDir, "\")
      strTempDir = strTempDir & "\" & vDir
      If vDir = "C:" Then GoTo next_dir
      If Dir(strTempDir, vbDirectory) = vbNullString Then
        vResponse = MsgBox("The directory at:" & vbCrLf & vbCrLf & strTempDir & vbCrLf & vbCrLf & "...does not exit. Create it now?", vbExclamation + vbYesNoCancel)
        If vResponse = vbYes Then
          MkDir strTempDir
        Else
          Me.lblDirectory.ForeColor = lngForeColorInvalid
          Me.txtDir.BorderColor = lngBorderColorInvalid
          blnError = True
        End If
      End If
next_dir:
    Next vDir
  End If
  'prevent duplication of EVT and EV%
  If Me.lboExport.ListCount > 0 Then
    For lngItem = Me.lboExport.ListCount - 1 To 0 Step -1
      If CLng(Me.lboExport.List(lngItem, 0)) = lngEVP Or CLng(Me.lboExport.List(lngItem, 0)) = lngEVT Then
        MsgBox "'" & Me.lboExport.List(lngItem, 1) & "' is included by default; removing from export list.", vbInformation + vbOKOnly, "Duplicate"
        Me.lboExport.RemoveItem lngItem
      End If
    Next lngItem
  End If
  'todo: ensure notes column title is unique in the columns
  If blnError Then
    Me.lblStatus.ForeColor = lngForeColorInvalid
    Me.lblStatus.Caption = " Please complete all required fields."
    Me.Repaint
  Else
    'save settings
    cptSaveStatusSheetSettings Me
    'create the sheet
    Application.DefaultDateFormat = pjDate_mm_dd_yyyy
    cptCreateStatusSheet Me
  End If

exit_here:
  On Error Resume Next
  Me.Repaint
  Application.DefaultDateFormat = lngDateFormat
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdRun_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdUp_Click()
  Dim lngExport As Long
  Dim lngField As Long, strField As String, strField2 As String
  Dim blnSelected As Boolean
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnSelected = False
  For lngExport = 0 To Me.lboExport.ListCount - 1
    If lngExport > 0 Then
      If Me.lboExport.Selected(lngExport) Then
        blnSelected = True
        'capture values
        lngField = Me.lboExport.List(lngExport - 1, 0)
        strField = Me.lboExport.List(lngExport - 1, 1)
        strField2 = Me.lboExport.List(lngExport - 1, 2)
        'move selected values
        Me.lboExport.List(lngExport - 1, 0) = Me.lboExport.List(lngExport, 0)
        Me.lboExport.List(lngExport - 1, 1) = Me.lboExport.List(lngExport, 1)
        Me.lboExport.List(lngExport - 1, 2) = Me.lboExport.List(lngExport, 2)
        Me.lboExport.Selected(lngExport - 1) = True
        Me.lboExport.List(lngExport, 0) = lngField
        Me.lboExport.List(lngExport, 1) = strField
        Me.lboExport.List(lngExport, 2) = strField2
        Me.lboExport.Selected(lngExport) = False
      End If
    End If
  Next lngExport
  
  If blnSelected Then cptRefreshStatusTable Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdUp_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdSave_Click()
  Me.Hide
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "lblURL", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboItems_Change()
  'strings
  Dim strCriteria As String
  Dim strFieldName As String
  'longs
  Dim lngSelectedItems As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Not Me.Visible Then Exit Sub
  If Me.ActiveControl.Name <> Me.lboItems.Name Then Exit Sub
  
  If Application.Calculation = pjAutomatic Then cptSpeed True
  
  strFieldName = Me.cboEach.Value

  For lngItem = 0 To Me.lboItems.ListCount - 1
    If Me.lboItems.Selected(lngItem) Then
      strCriteria = strCriteria & Me.lboItems.List(lngItem) & Chr$(9)
      lngSelectedItems = lngSelectedItems + 1
    End If
  Next lngItem
  
  If Len(strCriteria) = 0 Then
    FilterClear
    FilterApply "cptStatusSheet Filter"
  Else
    FilterClear
    FilterApply "cptStatusSheet Filter"
    strCriteria = Left(strCriteria, Len(strCriteria) - 1)
    SetAutoFilter FieldName:=strFieldName, FilterType:=pjAutoFilterIn, Criteria1:=strCriteria
  End If
  
  If Me.Visible Then
    If Me.ActiveControl.Name = Me.lboItems.Name Then
      Me.chkAllItems.Value = lngSelectedItems = Me.lboItems.ListCount
    End If
  End If
  
  If lngSelectedItems > 0 Then
    Me.lblIncludeItems.ForeColor = lngForeColorValid
    Me.lboItems.BorderColor = lngBorderColorValid
  End If
  
  ActiveWindow.TopPane.Activate
  SelectBeginning
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "lboItems_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub SpinButton1_SpinDown()
  Dim lngExport As Long
  Dim lngField As Long, strField As String, strField2 As String
  Dim blnSelected As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnSelected = False
  For lngExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If lngExport < Me.lboExport.ListCount - 1 Then
      If Me.lboExport.Selected(lngExport) Then
        blnSelected = True
        'capture values
        lngField = Me.lboExport.List(lngExport + 1, 0)
        strField = Me.lboExport.List(lngExport + 1, 1)
        strField2 = Me.lboExport.List(lngExport + 1, 2)
        'move selected values
        Me.lboExport.List(lngExport + 1, 0) = Me.lboExport.List(lngExport, 0)
        Me.lboExport.List(lngExport + 1, 1) = Me.lboExport.List(lngExport, 1)
        Me.lboExport.List(lngExport + 1, 2) = Me.lboExport.List(lngExport, 2)
        Me.lboExport.Selected(lngExport + 1) = True
        Me.lboExport.List(lngExport, 0) = lngField
        Me.lboExport.List(lngExport, 1) = strField
        Me.lboExport.List(lngExport, 2) = strField2
        Me.lboExport.Selected(lngExport) = False
      End If
    End If
  Next lngExport

  If blnSelected Then cptRefreshStatusTable Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdDown_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub SpinButton1_SpinUp()
  Dim lngExport As Long
  Dim lngField As Long, strField As String, strField2 As String
  Dim blnSelected As Boolean
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnSelected = False
  For lngExport = 0 To Me.lboExport.ListCount - 1
    If lngExport > 0 Then
      If Me.lboExport.Selected(lngExport) Then
        blnSelected = True
        'capture values
        lngField = Me.lboExport.List(lngExport - 1, 0)
        strField = Me.lboExport.List(lngExport - 1, 1)
        strField2 = Me.lboExport.List(lngExport - 1, 2)
        'move selected values
        Me.lboExport.List(lngExport - 1, 0) = Me.lboExport.List(lngExport, 0)
        Me.lboExport.List(lngExport - 1, 1) = Me.lboExport.List(lngExport, 1)
        Me.lboExport.List(lngExport - 1, 2) = Me.lboExport.List(lngExport, 2)
        Me.lboExport.Selected(lngExport - 1) = True
        Me.lboExport.List(lngExport, 0) = lngField
        Me.lboExport.List(lngExport, 1) = strField
        Me.lboExport.List(lngExport, 2) = strField2
        Me.lboExport.Selected(lngExport) = False
      End If
    End If
  Next lngExport
  
  If blnSelected Then cptRefreshStatusTable Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "SpinButton1_SpinUp", Err, Erl)
  Resume exit_here

End Sub

Private Sub stxtSearch_Change()
  Dim lngItem As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboFields.Clear
  
  With CreateObject("ADODB.REcordset")
    .Open cptDir & "\settings\cpt-status-sheet-search.adtg"
    If Len(Me.stxtSearch.Text) > 0 Then
      .Filter = "[Custom Field Name] LIKE '*" & cptRemoveIllegalCharacters(Me.stxtSearch.Text) & "*'"
    Else
      .Filter = 0
    End If
    If .RecordCount > 0 Then .MoveFirst
    lngItem = 0
    Do While Not .EOF
      Me.lboFields.AddItem
      Me.lboFields.List(lngItem, 0) = .Fields(0)
      Me.lboFields.List(lngItem, 1) = .Fields(1)
      Me.lboFields.List(lngItem, 2) = .Fields(2)
      .MoveNext
      lngItem = lngItem + 1
    Loop
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "stxtSearch_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub stxtSearch_Enter()
  Dim lngField As Long, strFileName As String
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = cptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Exit Sub
  With CreateObject("ADODB.Recordset")
    .Fields.Append "Field Constant", 200, 100 '200=adVarChar
    .Fields.Append "Custom Field Name", 200, 100 '200=adVarChar
    .Fields.Append "Local Field Name", 200, 100 '200=adVarChar
    .Open
    For lngField = 0 To Me.lboFields.ListCount - 1
      .AddNew Array(0, 1, 2), Array(Me.lboFields.List(lngField, 0), Me.lboFields.List(lngField, 1), Me.lboFields.List(lngField, 2))
    Next lngField
    If Not .EOF Then .Update
    .Save strFileName
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "stxtSearch_Enter", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub txtCC_DropButtonClick()
  Dim strHint As String
  strHint = "List of email addresses separated by a simi-colon (';')."
  MsgBox strHint, vbInformation + vbOKOnly, "CC hint"
End Sub

Private Sub txtDir_Change()
  Dim strDir As String
  Dim strNamingConvention As String
  
  strDir = Me.txtDir.Text
  If InStr(strDir, "[yyyy-mm-dd]") > 0 Then
    strDir = Replace(strDir, "[yyyy-mm-dd]", Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
  End If
  If Right(strDir, 1) <> "\" Then
    strDir = strDir & "\"
  End If
  'todo: apply chkAppendStatusDate
  Me.lblDirSample.Caption = strDir
  
  If Dir(strDir, vbDirectory) = vbNullString Then
    Me.txtDir.BorderColor = lngBorderColorInvalid
    Me.txtDir.ForeColor = lngForeColorInvalid
    Me.lblDirectory.ForeColor = lngForeColorInvalid
    Me.lblDirSample.ForeColor = lngForeColorInvalid
  Else
    Me.txtDir.BorderColor = lngBorderColorValid
    Me.txtDir.ForeColor = lngForeColorValid
    Me.lblDirectory.ForeColor = lngForeColorValid
    Me.lblDirSample.ForeColor = 8421504
  End If
  Me.Repaint

End Sub

Private Sub txtDir_DropButtonClick()
  'objects
  Dim oShell As Object
  Dim oFileDialog As Object 'FileDialog
  Dim oExcel As Excel.Application
  'strings
  Dim strValidPath As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oExcel = CreateObject("Excel.Application")
  Set oFileDialog = oExcel.FileDialog(msoFileDialogFolderPicker)
  With oFileDialog
    .AllowMultiSelect = False
    If Left(ActiveProject.Path, 2) = "<>" Or Left(ActiveProject.Path, 4) = "http" Then 'server project: default to Desktop
      Set oShell = CreateObject("WScript.Shell")
      .InitialFileName = oShell.SpecialFolders("Desktop")
    Else 'not a server project
      .InitialFileName = ActiveProject.Path
    End If
    If .Show Then
      strValidPath = cptValidPath(.SelectedItems(1))
      If Not CBool(Split(strValidPath, ":")(0)) Then
        MsgBox Replace(strValidPath, "0:", "Reason: "), vbCritical + vbOKOnly, "Invalid Path"
      Else
        Me.txtDir = .SelectedItems(1) & "\" & IIf(Me.chkAppendStatusDate, Format(ActiveProject.StatusDate, "yyyy-mm-dd") & "\", "")
      End If
    End If
  End With
  
exit_here:
  On Error Resume Next
  Set oShell = Nothing
  Set oFileDialog = Nothing
  Set oExcel = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdDir_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtDir_Enter()
  Me.lblDirSample.BackStyle = fmBackStyleOpaque
  Me.lblDirSample.Visible = True
  Me.lblStatusDate.Visible = False
  Me.lblCreate.Visible = False
  Me.lblDirectory.Visible = False
  Me.chkAppendStatusDate.Visible = False
  Me.lblPathLength.Visible = True
End Sub

Private Sub txtDir_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  Me.lblDirSample.BackStyle = fmBackStyleOpaque
  Me.lblDirSample.Visible = False
  Me.lblStatusDate.Visible = True
  Me.lblCreate.Visible = True
  Me.lblDirectory.Visible = True
  Me.chkAppendStatusDate.Visible = True
  Me.lblPathLength.Visible = False
End Sub

Sub txtFileName_Change()
  Dim strFileName As String
  Dim strNamingConvention As String
  Dim strTempItem As String
  Dim blnValid As Boolean
  
  Me.txtFileName.Text = cptRemoveIllegalCharacters(Me.txtFileName.Text, "", True)
  strFileName = Me.txtFileName.Text
  strNamingConvention = strFileName
  blnValid = True
  
  'clean date format
  strTempItem = cptRegEx(strFileName, "\[(Y|y){1,}-(M|m){1,}-(D|d){1,}\]")
  If Len(strTempItem) > 0 Then
    strNamingConvention = Replace(strFileName, strTempItem, "[yyyy-mm-dd]")
    strFileName = Replace(strFileName, "[yyyy-mm-dd]", Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
  End If
  
  'clean program
  strTempItem = cptRegEx(strFileName, "\[(P|p)(R|r)(O|o)(G|g)(R|r)(A|a)(M|m)\]")
  If Len(strTempItem) > 0 Then
    strFileName = Replace(strFileName, strTempItem, cptGetProgramAcronym)
    strNamingConvention = Replace(strNamingConvention, strTempItem, "[program]")
  End If
  
  'make [item] case insensitive
  strTempItem = cptRegEx(strNamingConvention, "\[(I|i)(T|t)(E|e)(M|m)\]")
  If Len(strTempItem) > 0 And strTempItem <> "[item]" Then
    strNamingConvention = Replace(strNamingConvention, strTempItem, "[item]")
  End If
  
  If Me.cboCreate.Value = 2 Then 'for each
    'replace [item] with an example
    If InStr(strFileName, "[item]") > 0 Then
      If Me.lboItems.ListCount > 0 Then
        Me.lblFileNameSample.Caption = Replace(strFileName, "[item]", Me.lboItems.List(0, 0)) & ".xlsx"
      Else
        Me.lblFileNameSample.Caption = "< no item found >"
      End If
    Else
      blnValid = False
      Me.lblFileNameSample.Caption = "'" & Me.cboCreate.List(Me.cboCreate.Value, 1) & "' requires use of '[item]'"
      Me.lblFileNameSample.ForeColor = lngForeColorInvalid
    End If
  Else
    Me.lblFileNameSample.Caption = strFileName & ".xlsx"
  End If
  If Me.txtFileName.Text <> strNamingConvention Then
    Me.txtFileName.Text = strNamingConvention
  End If
  
  If Not blnValid Then
    Me.txtFileName.BorderColor = lngBorderColorInvalid
    Me.lblNamingConvention.ForeColor = lngForeColorInvalid
    Me.lblFileNameSample.ForeColor = lngForeColorInvalid
    Me.lblPathLength.Caption = "<!>"
    Me.lblPathLength.BorderColor = lngBorderColorInvalid
    Me.lblPathLength.ForeColor = lngForeColorInvalid
  Else
    Me.txtFileName.BorderColor = lngBorderColorValid
    Me.lblNamingConvention.ForeColor = lngForeColorValid
    Me.lblFileNameSample.ForeColor = 8421504
    Dim lngPathLength As Long
    Dim lngLongest As Long
    Dim lngItem As Long
    lngPathLength = Len(Me.txtDir.Value) + Len(Replace(strNamingConvention, "[item]", "")) + Len(".xlsx")
    If Me.chkAppendStatusDate Then lngPathLength = lngPathLength + Len("YYYY-MM-DD\")
    If Me.lboItems.ListCount > 0 Then
      For lngItem = 0 To Me.lboItems.ListCount - 1
        If Len(Me.lboItems.List(lngItem)) > lngLongest Then lngLongest = Len(Me.lboItems.List(lngItem))
      Next lngItem
      lngPathLength = lngPathLength + lngLongest
    End If
    Me.lblPathLength.Caption = lngPathLength
    If lngPathLength < 218 Then
      Me.lblPathLength.BorderColor = lngBorderColorValid
      Me.lblPathLength.ForeColor = 8421504
    Else
      Me.lblPathLength.BorderColor = lngBorderColorInvalid
      Me.lblPathLength.ForeColor = lngForeColorInvalid
    End If
  End If
  Me.Repaint
  
End Sub

Private Sub txtFileName_DropButtonClick()
  Dim strHint As String
  strHint = "Fields Available:" & vbCrLf
  strHint = strHint & "> [status_date] = '" & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & "'" & vbCrLf
  strHint = strHint & "> [yyyy-mm-dd] = '" & Format(ActiveProject.StatusDate, "yyyy-mm-dd") & "'" & vbCrLf
  strHint = strHint & "> [program] = '" & cptGetProgramAcronym & "'" & vbCrLf
  strHint = strHint & "> [item] = selected item in 'for each' list (required for both 'for each' options)"
  MsgBox strHint, vbInformation + vbOKOnly, "File Naming Convention hints"
  
End Sub

Private Sub txtFileName_Enter()
  Me.lblFileNameSample.BackStyle = fmBackStyleOpaque
  Me.lblFileNameSample.Visible = True
  Me.lblNamingConvention.Visible = False
  Me.lblPathLength.Visible = True
End Sub

Private Sub txtFileName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  Me.lblFileNameSample.Visible = False
  Me.lblFileNameSample.BackStyle = fmBackStyleTransparent
  Me.lblNamingConvention.Visible = True
  Me.lblPathLength.Visible = False
End Sub

Private Sub txtHideCompleteBefore_AfterUpdate()
  If IsDate(Me.txtHideCompleteBefore) Then
    Me.txtHideCompleteBefore = FormatDateTime(Me.txtHideCompleteBefore, vbShortDate)
  End If
End Sub

Private Sub txtHideCompleteBefore_Change()
  Dim stxt As String
  Dim dtLookahead As Date
  Dim dtStatus As Date
  Dim blnValid As Boolean
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not Me.Visible Then GoTo exit_here
  'If Me.ActiveControl.Name <> "txtHideCompleteBefore" Then GoTo exit_here
  stxt = cptRegEx(Me.txtHideCompleteBefore.Text, "[0-9\/]*")
  Me.txtHideCompleteBefore.Text = stxt
  blnValid = False
  If Len(Me.txtHideCompleteBefore.Text) > 0 Then
    If IsDate(Me.txtHideCompleteBefore.Text) Then
      dtLookahead = FormatDateTime(Me.txtHideCompleteBefore, vbShortDate)
      dtStatus = FormatDateTime(ActiveProject.StatusDate)
      If dtLookahead > #1/1/1984# And dtLookahead < dtStatus Then
        blnValid = True
      End If
    End If
  End If
  If blnValid Then
    Me.chkHide.ForeColor = lngForeColorValid
    Me.txtHideCompleteBefore.ForeColor = lngForeColorValid
    Me.txtHideCompleteBefore.BorderColor = lngBorderColorValid
  Else
    Me.chkHide.ForeColor = lngForeColorInvalid
    Me.txtHideCompleteBefore.ForeColor = lngForeColorInvalid
    Me.txtHideCompleteBefore.BorderColor = lngBorderColorInvalid
  End If
  Me.Repaint
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "txtHideCompleteBefore", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtHideCompleteBefore_DropButtonClick()
  Dim strMsg As String
  MsgBox "Use this to include tasks that, e.g., were completed in the last status cycle.", vbInformation + vbOKOnly, "Show Completed After"
End Sub

Private Sub txtLookaheadDate_AfterUpdate()
  If Len(Me.txtLookaheadDate) > 0 And IsDate(Me.txtLookaheadDate) Then
    Me.txtLookaheadDate = CDate(FormatDateTime(Me.txtLookaheadDate.Value, vbShortDate))
  End If
End Sub

Private Sub txtLookaheadDate_Change()
  Dim dtDate As Date
  
  If Not Me.Visible Then Exit Sub
  If Not IsDate(ActiveProject.StatusDate) Then Exit Sub
  'If Not Me.ActiveControl.Name = Me.txtLookaheadDate.Name Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Len(Me.txtLookaheadDate.Text) > 0 Then
    Me.txtLookaheadDate.Text = cptRegEx(Me.txtLookaheadDate.Text, "[0-9\/]{1,}")
    'limit to a dates only
    If Not IsDate(Me.txtLookaheadDate.Text) Then
      Me.txtLookaheadDays = ""
      Me.txtLookaheadDate.BorderColor = lngBorderColorInvalid
      Me.txtLookaheadDate.ForeColor = lngForeColorInvalid
      Me.lblLookaheadWeekday.Visible = False
      Me.Repaint
    Else
      'limit to dates after the start date
      dtDate = CDate(FormatDateTime(Me.txtLookaheadDate.Text, vbShortDate) & " 5:00 PM")
      If dtDate < ActiveProject.StatusDate Then
        Me.txtLookaheadDate.BorderColor = lngBorderColorInvalid
        Me.txtLookaheadDate.ForeColor = lngForeColorInvalid
        Me.lblLookaheadWeekday.Visible = False
      Else
        Me.txtLookaheadDays = CLng(Application.DateDifference(ActiveProject.StatusDate, dtDate) / 480)
        Me.txtLookaheadDate.BorderColor = lngBorderColorValid
        Me.txtLookaheadDate.ForeColor = lngForeColorValid
        Me.lblLookaheadWeekday.Visible = True
        Me.lblLookaheadWeekday.Caption = Format(CDate(Me.txtLookaheadDate.Text), "dddd")
        Me.Repaint
        Call cptRefreshStatusTable(Me, False, True)
      End If
      Me.Repaint
    End If
  Else
    Me.txtLookaheadDays = ""
  End If

exit_here:
  On Error Resume Next
  Me.Repaint
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "txtLookaheadDate_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtLookaheadDate_Enter()
  Me.lblLookaheadWeekday.BackStyle = fmBackStyleOpaque
  Me.lblLookaheadWeekday.Visible = True
  Me.lblSearch.Visible = False
End Sub

Private Sub txtLookaheadDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  Me.lblLookaheadWeekday.Visible = False
  Me.lblLookaheadWeekday.BackStyle = fmBackStyleTransparent
  Me.lblSearch.Visible = True
End Sub

Private Sub txtLookaheadDays_Change()
  Dim lngDays As Long
  
  'If Not Me.Visible Then Exit Sub
  If Not IsDate(ActiveProject.StatusDate) Then Exit Sub
  If Not Me.ActiveControl.Name = Me.txtLookaheadDays.Name Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Len(Me.txtLookaheadDays.Text) > 0 Then
    lngDays = CLng(cptRegEx(Me.txtLookaheadDays, "[0-9]{1,}"))
    Me.txtLookaheadDays.Text = lngDays
    Me.txtLookaheadDate.Value = FormatDateTime(Application.DateAdd(ActiveProject.StatusDate, lngDays * 480), vbShortDate)
    Me.txtLookaheadDate.BorderColor = lngBorderColorValid
    Me.txtLookaheadDate.ForeColor = lngForeColorValid
    Me.lblLookaheadWeekday.Visible = True
    Me.lblLookaheadWeekday.Caption = Format(Me.txtLookaheadDate, "dddd")
    Call cptRefreshStatusTable(Me, False, True)
  Else
    Me.txtLookaheadDate = ""
    Me.txtLookaheadDate.BorderColor = lngBorderColorValid
    Me.txtLookaheadDate.ForeColor = lngForeColorValid
    Me.lblLookaheadWeekday.Visible = False
  End If

exit_here:
  On Error Resume Next
  Me.Repaint
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "txtLookaheadDays_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtLookaheadDays_Enter()
  Me.lblLookaheadWeekday.BackStyle = fmBackStyleOpaque
  Me.lblLookaheadWeekday.Visible = True
  Me.lblSearch.Visible = False
End Sub

Private Sub txtLookaheadDays_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  Me.lblLookaheadWeekday.Visible = False
  Me.lblLookaheadWeekday.BackStyle = fmBackStyleTransparent
  Me.lblSearch.Visible = True
End Sub

Private Sub txtNotesColTitle_Change()
  If Not Me.Visible Then Exit Sub
  'todo: ensure column name uniqueness

End Sub

Private Sub txtNotesColTitle_DropButtonClick()
  Dim strMsg As String
  strMsg = "Define a custom name for the exported Task Notes column." & vbCrLf
  strMsg = strMsg & "Default is 'Reason / Action / Impact'" & vbCrLf
  strMsg = strMsg & "Suggestions include: 'Basis of Estimate'; 'Status Notes'; 'Duration Uncertainty (H,M,L)', etc." & vbCrLf & vbCrLf
  strMsg = strMsg & "Task Notes can be exported (and imported) at the Task and/or the Assignment level."
  MsgBox strMsg, vbInformation + vbOKOnly, "Note Column Hints"
  
End Sub

Private Sub txtStatusDate_Change()
  Dim dtStatus As Date
  Dim lngDiff As Long
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not Me.Visible Then GoTo exit_here
  
  'form won't open without a status date and is modal
  dtStatus = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  Me.txtStatusDate.Text = dtStatus
  If Me.chkHide Then
    lngDiff = Application.DateDifference(CDate(Me.txtHideCompleteBefore), dtStatus)
  End If
  
  'update hide complete before
  If Me.chkHide Then
    Me.txtHideCompleteBefore = FormatDateTime(Application.DateSubtract(dtStatus, lngDiff), vbShortDate)
    If IsDate(CDate(Me.txtHideCompleteBefore)) Then
      If CDate(Me.txtHideCompleteBefore) >= dtStatus Then
        Me.chkHide.ForeColor = lngForeColorInvalid
        Me.txtHideCompleteBefore.BorderColor = lngBorderColorInvalid
        Me.txtHideCompleteBefore.ForeColor = lngForeColorInvalid
      Else
        Me.chkHide.ForeColor = lngForeColorValid
        Me.txtHideCompleteBefore.BorderColor = lngBorderColorValid
        Me.txtHideCompleteBefore.ForeColor = lngForeColorValid
      End If
    Else
      Me.txtHideCompleteBefore = FormatDateTime(Application.DateSubtract(dtStatus, 5 * 480), vbShortDate)
    End If
  End If
  
  'update lookahead date based on lookahead days
  If Me.chkLookahead Then
    Me.txtLookaheadDate = FormatDateTime(Application.DateAdd(ActiveProject.StatusDate, CLng(Me.txtLookaheadDays) * 480), vbShortDate)
  End If
  
exit_here:
  On Error Resume Next
  Me.Repaint
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "txtStatusDate_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub txtStatusDate_DropButtonClick()
  Dim lngDiff As Long
  
  If Me.chkHide Then lngDiff = VBA.DateDiff("d", CDate(Me.txtHideCompleteBefore), CDate(Me.txtStatusDate))
  If lngDiff = 0 Then lngDiff = 5
  
  If ChangeStatusDate Then
    Me.txtStatusDate = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
    If Me.chkHide Then Me.txtHideCompleteBefore = FormatDateTime(VBA.DateAdd("d", -lngDiff, ActiveProject.StatusDate), vbShortDate)
  End If
  
End Sub

Private Sub txtSubject_Change()
  Dim strSubjectPattern As String
  Dim strSubjectHint As String
  Dim strTempItem As String
  
  strSubjectPattern = Me.txtSubject.Text
  strSubjectHint = strSubjectPattern
    
  'clean date format
  strTempItem = cptRegEx(strSubjectPattern, "\[(Y|y){1,}-(M|m){1,}-(D|d){1,}\]")
  If Len(strTempItem) > 0 Then
    strSubjectPattern = Replace(strSubjectPattern, strTempItem, "[yyyy-mm-dd]")
    strSubjectHint = Replace(strSubjectPattern, "[yyyy-mm-dd]", Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
  End If
  
  'clean status date
  strTempItem = cptRegEx(strSubjectPattern, "\[(S|s)(T|t)(A|a)(T|t)(U|u)(S|s).(D|d)(A|a)(T|t)(E|e)\]")
  If Len(strTempItem) > 0 Then
    strSubjectPattern = Replace(strSubjectPattern, strTempItem, "[status_date]")
    strSubjectHint = Replace(strSubjectHint, strTempItem, FormatDateTime(ActiveProject.StatusDate, "m/d/yyyy"))
  End If
  
  'clean program
  strTempItem = cptRegEx(strSubjectPattern, "\[(P|p)(R|r)(O|o)(G|g)(R|r)(A|a)(M|m)\]")
  If Len(strTempItem) > 0 Then
    strSubjectPattern = Replace(strSubjectPattern, strTempItem, "[program]")
    strSubjectHint = Replace(strSubjectHint, strTempItem, cptGetProgramAcronym)
  End If

  'clean item
  strTempItem = cptRegEx(strSubjectPattern, "\[(I|i)(T|t)(E|e)(M|m)\]")
  If Len(strTempItem) > 0 And strTempItem <> "[item]" Then
    strSubjectPattern = Replace(strSubjectPattern, strTempItem, "[item]")
  End If

  If Me.cboCreate > 0 And Me.lboItems.ListCount > 0 Then
    strSubjectHint = Replace(strSubjectHint, "[item]", Me.lboItems.List(0, 0))
  Else
    strSubjectHint = Trim(Replace(strSubjectHint, "[item]", ""))
  End If
  Me.lblSubjectPreview.Caption = strSubjectHint
  If Me.txtSubject.Text <> strSubjectPattern Then
    Me.txtSubject.Text = strSubjectPattern
  End If
  
End Sub

Private Sub txtSubject_DropButtonClick()
  Dim strHint As String

  strHint = "The following fields are available for auto replacement in the subject line and in your Email Template (a.k.a., 'Quick Part'):" & vbCrLf & vbCrLf
  strHint = strHint & "> [status_date] = '" & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & "'" & vbCrLf
  strHint = strHint & "> [yyyy-mm-dd] = '" & Format(ActiveProject.StatusDate, "yyyy-mm-dd") & "'" & vbCrLf
  strHint = strHint & "> [program] = '" & cptGetProgramAcronym & "'" & vbCrLf
  strHint = strHint & "> [item] = selected item in 'for each' list (required for both 'for each' options)"
  MsgBox strHint, vbInformation + vbOKOnly, "Email Subject hints"

End Sub

Private Sub txtSubject_Enter()
  Me.chkKeepOpen.Visible = False
  Me.chkSendEmails.Visible = False
  Me.lblSubjectPreview.Visible = True
End Sub

Private Sub txtSubject_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  Me.chkKeepOpen.Visible = True
  Me.chkSendEmails.Visible = True
  Me.lblSubjectPreview.Visible = False
End Sub

Private Sub UserForm_Initialize()
  Me.txtStatusDate.DropButtonStyle = fmDropButtonStyleEllipsis
  Me.txtStatusDate.ShowDropButtonWhen = fmShowDropButtonWhenAlways
  Me.txtDir.DropButtonStyle = fmDropButtonStyleEllipsis
  Me.txtDir.ShowDropButtonWhen = fmShowDropButtonWhenAlways
  Me.txtSubject.DropButtonStyle = fmDropButtonStyleEllipsis
  Me.txtSubject.ShowDropButtonWhen = fmShowDropButtonWhenAlways
  Me.txtCC.DropButtonStyle = fmDropButtonStyleEllipsis
  Me.txtCC.ShowDropButtonWhen = fmShowDropButtonWhenAlways
  Me.txtFileName.DropButtonStyle = fmDropButtonStyleEllipsis
  Me.txtFileName.ShowDropButtonWhen = fmShowDropButtonWhenAlways
  Me.txtNotesColTitle.DropButtonStyle = fmDropButtonStyleEllipsis
  Me.txtNotesColTitle.ShowDropButtonWhen = fmShowDropButtonWhenAlways
  Me.txtHideCompleteBefore.DropButtonStyle = fmDropButtonStyleEllipsis
  Me.txtHideCompleteBefore.ShowDropButtonWhen = fmShowDropButtonWhenAlways
  Me.lblDirSample.Visible = False
  Me.lblFileNameSample.Visible = False
  Me.lblSubjectPreview.Visible = False
  Me.lblLookaheadWeekday.Visible = False
  Me.txtStatusDate.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  ElseIf CloseMode = VbQueryClose.vbFormCode Then
    If Me.ActiveControl.Name = "cmdSave" Then
      cptSaveStatusSheetSettings Me
    ElseIf Me.ActiveControl.Name = "cmdRun" Then
      cptSaveStatusSheetSettings Me
    End If
  End If
End Sub
