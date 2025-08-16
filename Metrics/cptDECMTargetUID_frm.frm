VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptDECMTargetUID_frm 
   Caption         =   "Target UID for Critical Path:"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "cptDECMTargetUID_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptDECMTargetUID_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v7.1.0</cpt_version>
Option Explicit
Public lngTargetTaskUID As Long

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdSubmit_Click()
  If IsNull(Me.lboTasks.Value) Then
    Me.lngTargetTaskUID = 0
  Else
    Me.lngTargetTaskUID = CLng(Me.lboTasks.Value)
  End If
  Me.Hide
End Sub

Private Sub lboTasks_AfterUpdate()
  If IsNull(Me.lboTasks.Value) Then Exit Sub
  If Me.lboTasks.Value > 0 Then
    Me.cmdSubmit.Caption = "Use " & Me.lboTasks.Value
    Me.cmdSubmit.Enabled = True
  Else
    Me.cmdSubmit.Caption = "Use This"
    Me.cmdSubmit.Enabled = False
  End If
End Sub

Private Sub txtTaskName_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.lboTasks.Clear
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  strDir = Environ("tmp")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT UID,TASK_NAME FROM [targets.csv] "
  If Len(Me.txtTaskName.Text) > 0 Then
    strSQL = strSQL & "WHERE LCASE(TASK_NAME) Like '%" & LCase(Me.txtTaskName.Text) & "%'"
  End If
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    Do While Not .EOF
      Me.lboTasks.AddItem
      Me.lboTasks.List(Me.lboTasks.ListCount - 1, 0) = oRecordset("UID")
      Me.lboTasks.List(Me.lboTasks.ListCount - 1, 1) = oRecordset("TASK_NAME")
      .MoveNext
    Loop
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDECMTargetUID_frm", "txtTaskName_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtUID_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.lboTasks.Clear
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  strDir = Environ("tmp")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT UID,TASK_NAME FROM [targets.csv] "
  If Len(Me.txtUID.Text) > 0 Then
    'first, clean the input
    lngUID = cptRegEx(Me.txtUID.Text, "[0-9]{1,}")
    Me.txtUID.Text = lngUID
    strSQL = strSQL & "WHERE UID Like '%" & lngUID & "%'"
  End If
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    Do While Not .EOF
      Me.lboTasks.AddItem
      Me.lboTasks.List(Me.lboTasks.ListCount - 1, 0) = oRecordset("UID")
      Me.lboTasks.List(Me.lboTasks.ListCount - 1, 1) = oRecordset("TASK_NAME")
      .MoveNext
    Loop
    .Close
  End With

exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDECMTargetUID_frm", "txtUID_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If Cancel = 1 Then Me.lngTargetTaskUID = 0
  If CloseMode = 0 Then Me.lngTargetTaskUID = 0
End Sub
