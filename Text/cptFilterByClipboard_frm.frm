VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFilterByClipboard_frm 
   Caption         =   "Filter By Clipboard"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "cptFilterByClipboard_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptFilterByClipboard_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.3.0</cpt_version>
Option Explicit

Private Sub chkFilter_Click()
  Dim strFilter As String
  strFilter = Me.txtFilter.Text
  Me.txtFilter.Text = ""
  Me.txtFilter.Value = strFilter
End Sub

Private Sub cmdClear_Click()
  Me.lboHeader.Clear
  Me.lboHeader.ColumnCount = 2
  Me.lboHeader.AddItem "UID"
  Me.lboHeader.Column(1, 0) = "Task Name"
  Me.lboFilter.ColumnCount = 2
  Me.txtFilter.Text = ""
  Me.txtFilter.Visible = True
  Me.lboFilter.Visible = False
  cptClearFreeField Me
  cptUpdateClipboard Me
  Me.txtFilter.SetFocus
End Sub

Private Sub cmdClose_Click()
  Me.Hide
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink ("http://" & Me.lblURL.Caption)

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_frm", "lblURL", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboFilter_Click()
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  'strings
  Dim strField As String
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  Dim dtGoTo As Date

  On Error Resume Next
  If Me.optUID Then
    lngUID = CLng(Me.lboFilter.Value)
    Set oTask = ActiveProject.Tasks.UniqueID(lngUID)
    strField = "Unique ID"
  ElseIf Me.optID Then
    lngUID = CLng(Me.lboFilter.Value)
    Set oTask = ActiveProject.Tasks.Item(lngUID)
    strField = "ID"
  End If
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oTask Is Nothing Then
    
    If IsDate(oTask.Stop) Then
      dtGoTo = oTask.Stop
    Else
      dtGoTo = oTask.Start
    End If
    If ActiveWindow.ActivePane <> ActiveWindow.TopPane Then ActiveWindow.TopPane.Activate
  
    If ActiveProject.Subprojects.Count = 0 Then 'use EditGoto
      On Error Resume Next
      If Not EditGoTo(oTask.ID, dtGoTo) Then
        If MsgBox("Task " & strField & " " & lngUID & " is currently hidden. Would you like to remove all filters, show summary tasks, and show all tasks in order to find it?", vbQuestion + vbYesNo, "Reset View?") = vbYes Then
          ScreenUpdating = False
          FilterClear
          OptionsViewEx DisplaySummaryTasks:=True
          SelectAll
          On Error Resume Next
          If Not OutlineShowAllTasks Then
            Sort "ID", , , , , , False, True
            OutlineShowAllTasks
          End If
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          ScreenUpdating = True
          If Not EditGoTo(oTask.ID, dtGoTo) Then
            MsgBox "An unknown error has occured--can't find it!", vbCritical + vbOKOnly, "Still can't find it"
          End If
        End If
      End If
    
    ElseIf ActiveProject.Subprojects.Count > 0 Then 'use Find
      On Error Resume Next
      If Not FindEx(strField, "equals", lngUID) Then
        If MsgBox("Task " & strField & " " & lngUID & " is currently hidden. Would you like to remove all filters, show summary tasks, and show all tasks in order to find it?", vbQuestion + vbYesNo, "Reset View?") = vbYes Then
          ScreenUpdating = False
          FilterClear
          OptionsViewEx DisplaySummaryTasks:=True
          SelectAll
          On Error Resume Next
          If Not OutlineShowAllTasks Then
            Sort "ID", , , , , , False, True
            OutlineShowAllTasks
          End If
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          ScreenUpdating = True
          If Not FindEx(strField, "equals", lngUID) Then
            MsgBox "An unknown error has occured--can't find it!", vbCritical + vbOKOnly, "Still can't find it"
          End If
        End If
      End If
      
    End If 'ActiveProject.Subprojects.Count = 0
  Else
    MsgBox "Task " & strField & " " & lngUID & " not found in this project.", vbExclamation + vbOKOnly, strField & " not found"
  End If 'Not oTask Is Nothing
  
exit_here:
  On Error Resume Next
  ScreenUpdating = True
  Set oTasks = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_frm", "lboFilter_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub optID_Click()
  Dim strFilter As String
  strFilter = Me.txtFilter.Text
  Me.txtFilter.Text = ""
  Me.txtFilter.Value = strFilter
  Me.lboHeader.List(0, 0) = "ID"
  FilterClear
  cptUpdateClipboard Me
End Sub

Private Sub optUID_Click()
  Dim strFilter As String
  strFilter = Me.txtFilter.Text
  Me.txtFilter.Text = ""
  Me.txtFilter.Value = strFilter
  Me.lboHeader.List(0, 0) = "UID"
  ActiveWindow.TopPane.Activate
  FilterClear
  cptUpdateClipboard Me
End Sub

Private Sub tglEdit_Click()
  If Me.tglEdit Then
    Me.txtFilter.Visible = True
    Me.lboFilter.Visible = False
    Me.txtFilter.SetFocus
  Else
    If Len(Me.txtFilter.Value) = 0 Then
      Me.txtFilter.Visible = True
      Me.lboFilter.Visible = False
      Me.txtFilter.SetFocus
    Else
      Me.txtFilter.Visible = False
      Me.lboFilter.Visible = True
    End If
  End If
  Me.txtFilter.Height = Me.lboFilter.Height
End Sub

Private Sub txtFilter_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
  'objects
  'strings
  Dim strDelimiter As String
  Dim strNewRange As String
  Dim strRange As String
  Dim strFilter As String
  Dim strItem As String
  'longs
  Dim lngInsert As Long
  Dim lngTo As Long
  Dim lngFrom As Long
  Dim lngDelimiter As Long
  Dim lngField As Long
  Dim lngRecord As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vRecord As Variant
  Dim strNewList As Variant
  Dim vData As Variant
  'dates

  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'scrub the incoming data
  vData = Split(Data.GetText, vbCrLf)

  'populate lboFilter
  If UBound(vData) > 1 Then 'user pasted a column of data
    Me.lboFilter.Clear
    For lngRecord = 0 To UBound(vData)
      If vData(lngRecord) = "" Then GoTo next_record
        strItem = cptRegEx(CStr(vData(lngRecord)), "[0-9]*")
        If Len(strItem) > 0 Then
          'ignore UID 0
          If CLng(strItem) = 0 Then GoTo next_record
          'remove duplicates
          If cptRegEx(CStr(strNewList), "\b" & strItem & "\b") = "" Then
            strNewList = strNewList & CLng(strItem) & ","
          End If
        End If
next_record:
    Next lngRecord
    
  Else 'user pasted single line of delimited values, possibly including ranges

    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    
    strFilter = Data.GetText
       
    'guess the delimiter
    lngDelimiter = cptGuessDelimiter(vData, "^([^\t\,\;]*[\t\,\;])+")
    If lngDelimiter = -1 Then 'couldn't figure it out
      strDelimiter = InputBox("Please enter delimiter, without apostrophes (',' or ';' or '\t' for tab):", "Delimiter Undetermined", ",")
      If strDelimiter = "\t" Then
        lngDelimiter = 32
      Else
        lngDelimiter = Asc(strDelimiter)
      End If
    End If
       
    vRecord = Split(strFilter, Chr(lngDelimiter))
        
    If IsEmpty(vRecord) Then GoTo exit_here
            
    'expand ranges
    For lngItem = 0 To UBound(vRecord)
      strRange = cptRegEx(CStr(vRecord(lngItem)), "[0-9]{1,}-[0-9]{1,}")
      If Len(strRange) > 0 Then
        strNewRange = ""
        lngFrom = CLng(Left(strRange, InStr(strRange, "-") - 1))
        lngTo = CLng(Mid(strRange, InStr(strRange, "-") + 1))
        If lngTo = lngFrom Then
          strFilter = Replace(strFilter, strRange, lngTo & Chr(lngDelimiter))
          GoTo skip_it
        ElseIf lngTo < lngFrom Then
          'switch it
          lngInsert = lngFrom
          lngFrom = lngTo
          lngTo = lngInsert
        End If
        For lngInsert = lngFrom To lngTo
          strNewRange = strNewRange & lngInsert & Chr(lngDelimiter)
        Next
        strFilter = Replace(strFilter, strRange, strNewRange)
      End If
skip_it:
    Next lngItem
    
    vRecord = Split(strFilter, Chr(lngDelimiter))
    
    For lngItem = 0 To UBound(vRecord)
      strItem = cptRegEx(CStr(vRecord(lngItem)), "[0-9]{1,}")
      If Len(strItem) > 0 Then
        'ignore UID 0
        If CLng(strItem) = 0 Then GoTo next_item
        'remove duplicates
        If cptRegEx(CStr(strNewList), "\b" & strItem & "\b") = "" Then
          strNewList = strNewList & CLng(strItem) & ","
        End If
      End If
next_item:
    Next lngItem
    
  End If
  
  Cancel = True
  Me.txtFilter.Visible = True
  Me.txtFilter.Text = strNewList
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_frm", "txtFilter_BeforeDropOrPaste", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtFilter_Change()
  If Me.txtFilter.Visible Then
    Me.lboFilter.Clear
    If Right(Me.txtFilter.Text, 1) = "," Then cptUpdateClipboard Me
  End If
End Sub

Private Sub txtFilter_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  If Len(Me.txtFilter.Text) > 0 Then cptClipboardJump Me
End Sub

Private Sub txtFilter_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Len(Me.txtFilter.Text) > 0 Then cptClipboardJump Me
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If ActiveProject.Subprojects.Count > 0 Then
    Me.optID = False
    Me.optUID = True
    Me.optID.Locked = True
    Me.optID.ControlTipText = "Unavailable for Master/Subproject files"
    Me.optID.Enabled = False
  Else
    Me.optID.Enabled = True
    Me.optID.Locked = False
    Me.optID.ControlTipText = ""
  End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub UserForm_Terminate()
  cptClearFreeField Me, True
  Unload Me
End Sub
