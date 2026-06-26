VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptBulkLogic_frm 
   Caption         =   "cptBulkLogic"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   OleObjectBlob   =   "cptBulkLogic_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptBulkLogic_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.1</cpt_version>
Private Const THIS_MODULE As String = "cptBulkLogic_frm"
Option Explicit

Private Sub chkID_Click()
  cptUpdateBulkLogicForm Me
  cptSaveSetting "BulkLogic", "chkID", IIf(Me.chkID, 1, 0)
End Sub

Private Sub cmdApply_Click()
  cptBulkLogicApply Me
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
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
  Call cptHandleErr(THIS_MODULE, "lblURL_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboFrom_Change()
  cptBulkLogicUpdateCommand Me
End Sub

Private Sub lboTo_Change()
  cptBulkLogicUpdateCommand Me
End Sub

Private Sub MultiPage1_Change()
  Me.txtFilterFrom = ""
  Me.txtFilterTo = ""
  cptUpdateBulkLogicForm Me
End Sub

Private Sub txtFilterFrom_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strCon As String
  Dim strSQL As String
  Dim strFilter As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboFrom.Clear
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("tmp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT UID,ID,TASK_NAME FROM [cpt-blt.csv]"
  strFilter = Me.txtFilterFrom.Text
  If Len(strFilter) > 0 Then
    If cptRxTest(strFilter, "^\d*$") Then 'no alpha, use UID/ID
      If Me.chkID Then
        strSQL = strSQL & " WHERE ID LIKE '%" & strFilter & "%'"
      Else
        strSQL = strSQL & " WHERE UID LIKE '%" & strFilter & "%'"
      End If
    Else 'has alpha
      strSQL = strSQL & " WHERE TASK_NAME LIKE '%" & strFilter & "%'"
    End If
  End If
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
      Do While Not .EOF
        Me.lboFrom.AddItem
        Me.lboFrom.List(Me.lboFrom.ListCount - 1, 0) = oRecordset(0)
        Me.lboFrom.List(Me.lboFrom.ListCount - 1, 1) = oRecordset(1)
        Me.lboFrom.List(Me.lboFrom.ListCount - 1, 2) = oRecordset(2)
        .MoveNext
      Loop
    End If
  End With
  
  'todo: if a single result, then select it
  
exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "txtFilterFrom_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub txtFilterTo_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strCon As String
  Dim strSQL As String
  Dim strFilter As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboTo.Clear
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("tmp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT UID,ID,TASK_NAME FROM [cpt-blt.csv]"
  strFilter = Me.txtFilterTo.Text
  If Len(strFilter) > 0 Then
    If cptRxTest(strFilter, "^\d*$") Then 'no alpha, use UID/ID
      If Me.chkID Then
        strSQL = strSQL & " WHERE ID LIKE '%" & strFilter & "%'"
      Else
        strSQL = strSQL & " WHERE UID LIKE '%" & strFilter & "%'"
      End If
    Else 'has alpha
      strSQL = strSQL & " WHERE TASK_NAME LIKE '%" & strFilter & "%'"
    End If
  End If
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
      Do While Not .EOF
        Me.lboTo.AddItem
        Me.lboTo.List(Me.lboTo.ListCount - 1, 0) = oRecordset(0)
        Me.lboTo.List(Me.lboTo.ListCount - 1, 1) = oRecordset(1)
        Me.lboTo.List(Me.lboTo.ListCount - 1, 2) = oRecordset(2)
        .MoveNext
      Loop
    End If
  End With
  
  'todo: if a single result, then select it
  
exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "txtFilterTo_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtLag_Change()
  Dim strLag As String
  strLag = Me.txtLag.Text
  If Len(strLag) > 0 Then
    strLag = cptRxReplace(strLag, "\D+", "")
    Me.txtLag.Text = strLag
  End If
End Sub

Private Sub UserForm_Terminate()
  If Dir(Environ("tmp") & "\Schema.ini") <> vbNullString Then Kill Environ("tmp") & "\Schema.ini"
  If Dir(Environ("tmp") & "\cpt-bulk-logic.csv") <> vbNullString Then Kill Environ("tmp") & "\cpt-bulk-logic.csv"
End Sub
