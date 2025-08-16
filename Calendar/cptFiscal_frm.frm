VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFiscal_frm 
   Caption         =   "Fiscal Calendar"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840.001
   OleObjectBlob   =   "cptFiscal_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptFiscal_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.2.0</cpt_version>
Option Explicit

Private Sub cboImportField_Change()
  If Not IsNull(Me.cboImportField.Value) Then
    Me.lblProgress.Width = Me.lblStatus.Width
    Me.lblProgress.Visible = True
    Me.lblStatus.Visible = True
    Me.cmdAnalyzeEVT.SetFocus
  Else
    Me.lblProgress.Width = Me.lblStatus.Width
    Me.lblProgress.Visible = False
    Me.lblStatus.Visible = False
  End If
End Sub

Private Sub chkImportResults_Click()
  Dim lngItem As Long, lngField As Long
  Dim strField As String
  Dim oDict As Scripting.Dictionary
  Dim vType As Variant

  On Error GoTo 0
  If chkImportResults Then
    Me.lblAvailableFields.Visible = True
    Me.cboImportField.Visible = True
    Me.cboImportField.Clear
    Set oDict = New Scripting.Dictionary
    oDict.Add "Number", 20
    oDict.Add "Text", 30
    For Each vType In oDict.Keys
      For lngItem = 1 To oDict.Item(vType)
        strField = vType & lngItem
        lngField = FieldNameToFieldConstant(strField)
        If Len(CustomFieldGetName(lngField)) > 0 Then
          'strField = CustomFieldGetName(lngField) & " (" & strField & ")"
          'Me.cboImportField.List(Me.cboImportField.ListCount - 1, 1) = strField
        Else
          Me.cboImportField.AddItem
          Me.cboImportField.List(Me.cboImportField.ListCount - 1, 0) = lngField
          Me.cboImportField.List(Me.cboImportField.ListCount - 1, 1) = strField
        End If
      Next lngItem
    Next vType
    If Me.cboImportField.ListCount = 0 Then
      MsgBox "You have no local custom number or text fields available.", vbExclamation + vbOKOnly, "No Room"
      Me.chkImportResults = False
    Else
      Me.cboImportField.SetFocus
      Me.cboImportField.DropDown
    End If
  Else
    Me.cboImportField.Clear
    Me.cboImportField.Visible = False
    Me.lblStatus.Visible = False
    Me.lblProgress.Visible = False
    Me.lblAvailableFields.Visible = False
  End If
  
  Set oDict = Nothing
End Sub

Private Sub cmdAnalyzeEVT_Click()
  If Me.chkImportResults Then
    If IsNull(Me.cboImportField.Value) Then
      MsgBox "Please select an avaialable local custom number or text field and try again.", vbCritical + vbOKOnly, "Import...where?"
      Exit Sub
    Else
      Call cptAnalyzeEVT(Me.cboImportField.Value)
    End If
  Else
    cptAnalyzeEVT Me
  End If
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdDelete_Click()
  If MsgBox("Are you sure?", vbExclamation + vbYesNo, "Confirm Delete") = vbYes Then
    If cptCalendarExists("cptFiscalCalendar") Then ActiveProject.BaseCalendars("cptFiscalCalendar").Delete
    Me.lboExceptions.Clear
    Me.txtExceptions = ""
    Me.lboExceptions.Visible = False
    Me.txtExceptions.Visible = True
    Me.lblCount.Caption = "0 exceptions."
  End If
End Sub

Private Sub cmdExport_Click()
  cptExportFiscalCalendar Me
End Sub

Private Sub cmdImport_Click()
  cptImportCalendarExceptions Me
End Sub

Private Sub cmdTemplate_Click()
  Call cptExportExceptionsTemplate
  Me.cmdImport.Enabled = True
  Me.cmdImport.ControlTipText = "Import a populated template"
End Sub

Private Sub txtExceptions_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
  'objects
  Dim oException As MSProject.Exception
  Dim oCalendar As MSProject.Calendar
  'strings
  Dim strLabel As String
  Dim strExceptions As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnLabels As Boolean
  'variants
  Dim vException As Variant
  Dim vExceptions As Variant
  'dates

  On Error Resume Next
  Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oCalendar Is Nothing Then
    BaseCalendarCreate Name:="cptFiscalCalendar", FromName:="Standard" ' [" & ActiveProject.Name & "]"
    Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
    If oCalendar.Exceptions.Count > 0 Then
      For Each oException In oCalendar.Exceptions
        oException.Delete
      Next oException
    End If
  End If
  
  vExceptions = Split(Data.GetText, vbCrLf)
  Me.lboExceptions.Clear
  For lngItem = 0 To UBound(vExceptions)
    vException = Split(vExceptions(lngItem), vbTab)
    If Len(Join(vException, vbTab)) = 0 Then GoTo next_item
    If UBound(vException) > 0 Then 'labels included
      blnLabels = True
      If IsDate(vException(0)) Then
        strExceptions = strExceptions & vException(0) & vbTab
        Me.lboExceptions.AddItem vException(0)
        strExceptions = strExceptions & vException(1) & vbCrLf
        Me.lboExceptions.List(Me.lboExceptions.ListCount - 1, 1) = vException(1)
        oCalendar.Exceptions.Add pjDaily, CStr(vException(0)), CStr(vException(0)), , CStr(vException(1))
      End If
    Else 'labels not included, guess them...
      blnLabels = False
      If IsDate(vExceptions(lngItem)) Then
        strExceptions = strExceptions & vExceptions(lngItem)
        Me.lboExceptions.AddItem vExceptions(lngItem)
        If Me.lboExceptions.ListCount = 1 Then
          strExceptions = strExceptions & vbTab & Format(vExceptions(lngItem), "yyyymm") & vbCrLf
          Me.lboExceptions.List(Me.lboExceptions.ListCount - 1, 1) = Format(vExceptions(lngItem), "yyyymm")
          oCalendar.Exceptions.Add pjDaily, CStr(vException(0)), CStr(vException(0)), , Format(vExceptions(lngItem), "yyyymm")
        Else
          strLabel = Me.lboExceptions.List(Me.lboExceptions.ListCount - 2, 1)
          If Right(strLabel, 2) = 12 Then
            strLabel = CStr(CLng(Left(strLabel, 4) + 1) & "01")
          Else
            strLabel = Left(strLabel, 4) & Format(CLng(Right(strLabel, 2) + 1), "00")
          End If
          strExceptions = strExceptions & vbTab & strLabel & vbCrLf
          Me.lboExceptions.List(Me.lboExceptions.ListCount - 1, 1) = strLabel
          oCalendar.Exceptions.Add pjDaily, CStr(vException(0)), CStr(vException(0)), , strLabel
        End If
      End If
    End If
next_item:
    Me.lblCount.Caption = oCalendar.Exceptions.Count & " exception" & IIf(oCalendar.Exceptions.Count = 1, "", "s") & "."
  Next lngItem

  If Not blnLabels Then
    MsgBox "Labels are required. We have attempted to guess them, but you can revise them in the 'cptFiscalCalendar' under Project > Change Working Time.", vbInformation + vbOKOnly, "No Labels Detected"
  End If
  
  Cancel = True
  If Len(strExceptions) > 0 Then
    Me.txtExceptions.Text = strExceptions
    Me.lboExceptions.Visible = True
    Me.txtExceptions.Visible = False
  Else
    Me.txtExceptions.Visible = True
    Me.lboExceptions.Visible = False
  End If
  
exit_here:
  On Error Resume Next
  Set oException = Nothing
  Set oCalendar = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_frm", "txtExceptions_BeforeDropOrPaste", Err, Erl)
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
