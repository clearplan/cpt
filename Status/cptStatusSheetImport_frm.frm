VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptStatusSheetImport_frm 
   Caption         =   "Import Status Sheets"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810.001
   OleObjectBlob   =   "cptStatusSheetImport_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptStatusSheetImport_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.3.1</cpt_version>
Option Explicit

Private Sub cboAF_Change()
  If Me.Visible Then
    If Me.ActiveControl.Name = "cboAF" Then Call cptRefreshStatusImportTable(Me)
  End If
End Sub

Private Sub cboAS_Change()
  If Me.Visible Then
    If Me.ActiveControl.Name = "cboAS" Then Call cptRefreshStatusImportTable(Me)
  End If
End Sub

Private Sub cboETC_Change()
  If Me.Visible Then
    If Me.ActiveControl.Name = "cboETC" Then
      'ensure user didn't select the same field for New EVP and New ETC
      '-2147483630 = Black
      Me.lblEV.ForeColor = -2147483630
      Me.lblETC.ForeColor = -2147483630
      If Me.cboEV.Value = Me.cboETC.Value Then
        Me.lblEV.ForeColor = 192
        Me.lblETC.ForeColor = 192
      Else
        Call cptRefreshStatusImportTable(Me)
      End If
    End If
  End If
End Sub

Private Sub cboEV_Change()
  If Me.Visible Then
    If Me.ActiveControl.Name = "cboEV" Then
      'ensure user didn't select the same field for New EVP and New ETC
      '-2147483630 = Black
      Me.lblEV.ForeColor = -2147483630
      Me.lblETC.ForeColor = -2147483630
      If Me.cboEV.Value = Me.cboETC.Value Then
        Me.lblEV.ForeColor = 192
        Me.lblETC.ForeColor = 192
      Else
        Call cptRefreshStatusImportTable(Me)
      End If
    End If
  End If
End Sub

Private Sub cboFF_Change()
  If Me.Visible Then
    If Me.ActiveControl.Name = "cboFF" Then Call cptRefreshStatusImportTable(Me)
  End If
End Sub

Private Sub cboFS_Change()
  If Me.Visible Then
    If Me.ActiveControl.Name = "cboFS" Then Call cptRefreshStatusImportTable(Me)
  End If
End Sub

Private Sub chkAppend_Click()
  Me.cboAppendTo.Enabled = Me.chkAppend
End Sub

Private Sub cmdDone_Click()
  Me.Hide
End Sub

Private Sub cmdImport_Click()
  
  'ensure user didn't select the same field for New EVP and New ETC
  '-2147483630 = Black
  Me.lblEV.ForeColor = -2147483630
  Me.lblETC.ForeColor = -2147483630
  If Me.cboEV.Value = Me.cboETC.Value Then
    Me.lblEV.ForeColor = 192
    Me.lblETC.ForeColor = 192
    MsgBox "Cannot import EVP and ETC to the same field.", vbExclamation + vbOKOnly, "Invalid Selections"
  Else
    'capture user settings
    cptSaveSetting "StatusSheetImport", "cboAS", Me.cboAS.Value
    cptSaveSetting "StatusSheetImport", "cboAF", Me.cboAF.Value
    cptSaveSetting "StatusSheetImport", "cboFS", Me.cboFS.Value
    cptSaveSetting "StatusSheetImport", "cboFF", Me.cboFF.Value
    cptSaveSetting "StatusSheetImport", "cboEVP", Me.cboEV.Value
    cptSaveSetting "StatusSheetImport", "cboETC", Me.cboETC.Value
    cptSaveSetting "StatusSheetImport", "chkNotes", CStr(Me.chkAppend)
    If Me.chkAppend Then
      cptSaveSetting "StatusSheetImport", "cboAppendTo", Me.cboAppendTo.Value
    Else
      cptSaveSetting "StatusSheetImport", "cboAppendTo", ""
    End If
    cptSaveSetting "StatusSheetImport", "optTaskUsage", IIf(Me.optAbove, "above", "below")
    
    cptStatusSheetImport Me
  End If
  
End Sub

Private Sub cmdRemove_Click()
  'objects
  'strings
  Dim strRemove As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vRemove As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngItem = Me.lboStatusSheets.ListCount - 1 To 0 Step -1
    If Me.lboStatusSheets.Selected(lngItem) Then
      strRemove = strRemove & lngItem & ","
      Me.lboStatusSheets.Selected(lngItem) = False
    End If
  Next lngItem
  
  For Each vRemove In Split(strRemove, ",")
    If vRemove = "" Then Exit For
    Me.lboStatusSheets.RemoveItem (CLng(vRemove))
  Next vRemove
  
  Me.cmdRemove.Enabled = False
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_frm", "cmdRemove_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdRename_Click()
  'objects
  Dim oDict As Scripting.Dictionary
  Dim oComboBox As MSForms.ComboBox
  'strings
  Dim strFieldName As String
  'longs
  Dim lngItem As Long
  Dim lngLCF As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSpeed True
  
  Set oDict = CreateObject("Scripting.Dictionary")
  oDict.Add "cboAS", "New Actual Start"
  oDict.Add "cboAF", "New Actual Finish"
  oDict.Add "cboFS", "New Forecast Start"
  oDict.Add "cboFF", "New Forecast Finish"
  oDict.Add "cboEV", "New EV%"
  oDict.Add "cboETC", "New ETC"
  
  For lngItem = 0 To oDict.Count - 1
    Set oComboBox = Me.Controls(oDict.Keys(lngItem))
    If oComboBox.Value > 0 Then
      lngLCF = Me.Controls(oDict.Keys(lngItem)).Value
      strFieldName = FieldConstantToFieldName(lngLCF)
      CustomFieldRename lngLCF, oDict.items(lngItem)
      oComboBox.List(oComboBox.ListIndex, 1) = strFieldName & " (" & oDict.items(lngItem) & ")"
    End If
  Next lngItem
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Set oDict = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_frm", "cmdRename_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdSelectFiles_Click()
  'objects
  Dim oShell As Object
  Dim oFileDialog As Object 'FileDialog
  Dim oExcel As Excel.Application
  'strings
  Dim strFile As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnQuit As Boolean
  'variants
  'dates

  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
    blnQuit = True
  Else
    blnQuit = False
  End If
  Set oFileDialog = oExcel.FileDialog(msoFileDialogFilePicker)
  With oFileDialog
    .AllowMultiSelect = True
    .ButtonName = "Import"
    .InitialView = 2 'msoFileDialogViewDetails
    If Left(ActiveProject.Path, 2) = "<>" Or Left(ActiveProject.Path, 4) = "http" Then 'server: default to Desktop
      Set oShell = CreateObject("WScript.Shell")
      .InitialFileName = oShell.SpecialFolders("Desktop") & "\"
    Else 'not a server project: use ActiveProject.Path
      .InitialFileName = ActiveProject.Path & "\"
    End If
    .Title = "Select Returned Status Sheet(s):"
    .Filters.Add "Microsoft Excel Workbook (xlsx)", "*.xlsx"
    If .Show = -1 Then
      If .SelectedItems.Count > 0 Then
        For lngItem = 1 To .SelectedItems.Count
          strFile = .SelectedItems(lngItem)
          If Dir(strFile) <> vbNullString Then
            Me.lboStatusSheets.AddItem
            Me.lboStatusSheets.List(Me.lboStatusSheets.ListCount - 1, 0) = Replace(strFile, Dir(strFile), "")
            Me.lboStatusSheets.List(Me.lboStatusSheets.ListCount - 1, 1) = Dir(strFile)
          End If
        Next lngItem
      End If
    End If
  End With

exit_here:
  On Error Resume Next
  Set oShell = Nothing
  Set oFileDialog = Nothing
  If blnQuit Then oExcel.Quit
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_frm", "cmdSelectFiles_Click", Err, Erl)
  Resume exit_here
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

Private Sub lboStatusSheets_Change()
  If InStr(Me.lblStatus.Caption, "Importing") > 0 Then Exit Sub
  If Me.lboStatusSheets.ListCount > 0 Then
    If Not IsNull(Me.lboStatusSheets.ListIndex) Then
      Me.cmdRemove.Enabled = True
    End If
  Else
    Me.cmdRemove.Enabled = False
  End If
End Sub

Private Sub lboStatusSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  'objects
  Dim oExcel As Object
  'strings
  Dim strPath As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Me.lboStatusSheets.ListCount > 0 Then
    For lngItem = 0 To Me.lboStatusSheets.ListCount - 1
      If Me.lboStatusSheets.Selected(lngItem) Then
        strPath = Me.lboStatusSheets.List(lngItem, 0) & Me.lboStatusSheets.List(lngItem, 1)
        If Dir(strPath) <> vbNullString Then
          On Error Resume Next
          Set oExcel = GetObject(, "Excel.Application")
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oExcel Is Nothing Then Set oExcel = CreateObject("Excel.Application")
          oExcel.Workbooks.Open strPath
          oExcel.Visible = True
          Application.ActivateMicrosoftApp pjMicrosoftExcel
        End If
        Exit For
      End If
    Next lngItem
  End If
  
exit_here:
  On Error Resume Next
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_frm", "lboStatusSheets_DblClick", Err, Erl)
  Resume exit_here
End Sub

Private Sub optAbove_Click()
  If Me.Visible Then Call cptRefreshStatusImportTable(Me, Me.optBelow)
End Sub

Private Sub optBelow_Click()
  If Me.Visible Then Call cptRefreshStatusImportTable(Me, Me.optBelow)
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
