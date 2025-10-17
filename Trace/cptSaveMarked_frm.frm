VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSaveMarked_frm 
   Caption         =   "Import Marked"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   OleObjectBlob   =   "cptSaveMarked_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptSaveMarked_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.10</cpt_version>
Option Explicit

Private Sub cmdDone_Click()
  Me.Hide
End Sub

Private Sub cmdImport_Click()
  'objects
  Dim oTask As MSProject.Task
  'strings
  'longs
  Dim lngResponse As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnApplyFilter As Boolean
  'variants
  'dates
  Dim dtTimestamp As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If IsNull(Me.lboMarked.Value) Then GoTo exit_here
  If Me.lboDetails.ListCount <= 1 Then
    MsgBox "No tasks were saved.", vbExclamation + vbOKOnly, "What happened?"
    GoTo exit_here
  End If
  
  'confirm overwrite
  lngResponse = MsgBox("Save current marked tasks before import?", vbQuestion + vbYesNoCancel, "Confirm Overwrite")
  If lngResponse = vbCancel Then
    GoTo exit_here
  ElseIf lngResponse = vbYes Then
    dtTimestamp = Me.lboMarked.Value
    cptSaveMarked
    Me.lboMarked.AddItem , 0
    cptUpdateMarked Me
    Me.lboMarked.Value = dtTimestamp
  End If
  
  'clear currently marked
  Application.StatusBar = "Clearing existing marked..."
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Marked Then oTask.Marked = False
next_task:
  Next oTask
  
  'import the set
  Application.StatusBar = "Marking saved set..."
  For lngItem = 1 To Me.lboDetails.ListCount - 1
    If Me.lboDetails.List(lngItem, 1) <> "< task not found >" Then
      ActiveProject.Tasks.UniqueID(CLng(Me.lboDetails.List(lngItem, 0))).Marked = True
    End If
    Application.StatusBar = "Marking saved set...(" & Format(lngItem / (Me.lboDetails.ListCount - 1), "0%") & ")"
  Next lngItem
  Application.StatusBar = "Marking saved set...done."
    
  'apply the filter
  blnApplyFilter = Me.chkApplyFilter
  If blnApplyFilter Then
    ActiveWindow.TopPane.Activate
    OptionsViewEx DisplaySummaryTasks:=True
    SelectAll
    OutlineShowAllTasks
    On Error Resume Next
    If Not FilterApply("Marked") Then
      FilterEdit "Marked", True, True, True, , , "Marked", , "equals", "Yes", , "Yes", False 'keep related summaries FALSE for compatibility with Network Browser
      FilterApply "Marked"
    End If
  End If
  cptSaveSetting "SaveMarked", "chkApplyFilter", IIf(blnApplyFilter, "1", "0")
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_frm", "cmdImport_Click", Err, Erl)
  Resume exit_here
    
End Sub

Private Sub cmdRemove_Click()
  'objects
  Dim rstMarked As Object 'ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strMarked As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtTimestamp As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  If Me.lboMarked.Value = "TIMESTAMP" Then GoTo exit_here
  If Me.lboMarked.ListIndex < 1 Then GoTo exit_here
  
  If MsgBox("This cannot be undone." & vbCrLf & vbCrLf & "Are you sure?", vbQuestion + vbYesNo, "Remove Saved Set") = vbNo Then GoTo exit_here
  
  'capture tstamp
  dtTimestamp = Me.lboMarked.Value
  Set rstMarked = CreateObject("ADODB.Recordset")
  'remove from marked
  strMarked = strDir & "\cpt-marked.adtg"
  rstMarked.Open strMarked
  rstMarked.Filter = "TSTAMP<>#" & dtTimestamp & "#"
  rstMarked.Save strMarked, adPersistADTG
  rstMarked.Close
  'remove from marked details
  strMarked = strDir & "\cpt-marked-details.adtg"
  rstMarked.Open strMarked
  rstMarked.Filter = "TSTAMP<>#" & dtTimestamp & "#"
  rstMarked.Save strMarked, adPersistADTG
  rstMarked.Close
  
  cptUpdateMarked Me
  Me.lboMarked.ListIndex = -1
  
exit_here:
  On Error Resume Next
  If rstMarked.State = 1 Then rstMarked.Close
  Set rstMarked = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_frm", "cmdRemove", Err, Erl)
  Resume exit_here
End Sub


Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_frm", "lblURL_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboMarked_Click()
  'objects
  Dim oTask As MSProject.Task
  Dim rstMarked As ADODB.Recordset 'Object
  'strings
  Dim strDir As String
  Dim strMarked As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  Dim dtTimestamp As Date
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  If Me.lboMarked.Value = "TIMESTAMP" Then
    Me.lboMarked.Value = Null
    GoTo exit_here
  End If

  With Me.lboDetails
    .Clear
    .AddItem
    .List(.ListCount - 1, 0) = "UID"
    .List(.ListCount - 1, 1) = "TASK NAME"
  End With

  strMarked = strDir & "\cpt-marked-details.adtg"
  If Dir(strMarked) = vbNullString Then
    MsgBox "Save Marked Details not found!", vbCritical + vbOKOnly, "Nada"
    GoTo exit_here
  End If
  
  dtTimestamp = CDate(Me.lboMarked.Value)

  Set rstMarked = CreateObject("ADODB.Recordset")
  With rstMarked
    .Open strMarked
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        If rstMarked(0) = dtTimestamp Then
          Me.lboDetails.AddItem
          Me.lboDetails.List(Me.lboDetails.ListCount - 1, 0) = rstMarked(1)
          Set oTask = Nothing
          On Error Resume Next
          Set oTask = ActiveProject.Tasks.UniqueID(CLng(rstMarked(1)))
          If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oTask Is Nothing Then
            Me.lboDetails.List(Me.lboDetails.ListCount - 1, 1) = "< task not found >"
          Else
            Me.lboDetails.List(Me.lboDetails.ListCount - 1, 1) = oTask.Name
          End If
        End If
        .MoveNext
      Loop
    End If
  End With
  
exit_here:
  On Error Resume Next
  Set oTask = Nothing
  Set rstMarked = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_frm", "lboMarked_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtFilter_Change()
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtFilter.Text) > 0 Then
    Call cptUpdateMarked(Me, Me.txtFilter.Text)
  Else
    Call cptUpdateMarked(Me)
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_frm", "txtFilter_Change", Err, Erl)
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
