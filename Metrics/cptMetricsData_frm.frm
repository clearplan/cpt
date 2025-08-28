VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptMetricsData_frm 
   Caption         =   "cpt Metrics Data"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   OleObjectBlob   =   "cptMetricsData_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptMetricsData_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.1.0</cpt_version>
Option Explicit

Private Sub cmdDelete_Click()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strProgram As String
  Dim strFileName As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtStatus As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  For lngItem = 0 To Me.lboMetricsData.ListCount - 1
    If Me.lboMetricsData.Selected(lngItem) Then
      strProgram = Me.lboMetricsData.List(lngItem, 0)
      dtStatus = CDate(Me.lboMetricsData.List(lngItem, 1))
      If MsgBox("Permanently delete record for " & strProgram & " - " & dtStatus & "?", vbExclamation + vbYesNo, "Confirm") = vbYes Then
        strFileName = cptDir & "\settings\cpt-metrics.adtg"
        If Dir(strFileName) = vbNullString Then
          MsgBox "File suddenly disappeared!", vbCritical + vbOKOnly, "File Not Found"
          GoTo exit_here
        End If
        Set oRecordset = CreateObject("ADODB.Recordset")
        oRecordset.Open strFileName
        If oRecordset.RecordCount = 0 Then
          MsgBox "No Records", vbExclamation + vbOKOnly, "No Data"
          oRecordset.Close
          GoTo exit_here
        End If
        oRecordset.MoveFirst
        oRecordset.Filter = "PROGRAM='" & strProgram & "' AND STATUS_DATE=#" & dtStatus & "#"
        If Not oRecordset.EOF Then
          oRecordset.Delete adAffectCurrent
          oRecordset.Filter = 0
          oRecordset.Save strFileName, adPersistADTG
        Else
          MsgBox "This record cannot be found.", vbExclamation + vbOKOnly, "Record Not Found"
          oRecordset.Close
          GoTo exit_here
        End If
        Me.lboMetricsData.Clear
        oRecordset.MoveFirst
        oRecordset.Sort = "STATUS_DATE DESC"
        oRecordset.Filter = "PROGRAM='" & strProgram & "'"
        Do While Not oRecordset.EOF
          Me.lboMetricsData.AddItem
          Me.lboMetricsData.List(Me.lboMetricsData.ListCount - 1, 0) = oRecordset.Fields(0)
          Me.lboMetricsData.List(Me.lboMetricsData.ListCount - 1, 1) = oRecordset.Fields(1)
          Me.lboMetricsData.List(Me.lboMetricsData.ListCount - 1, 2) = oRecordset.Fields(2)
          Me.lboMetricsData.List(Me.lboMetricsData.ListCount - 1, 3) = oRecordset.Fields(3)
          Me.lboMetricsData.List(Me.lboMetricsData.ListCount - 1, 4) = oRecordset.Fields(4)
          Me.lboMetricsData.List(Me.lboMetricsData.ListCount - 1, 5) = oRecordset.Fields(5)
          Me.lboMetricsData.List(Me.lboMetricsData.ListCount - 1, 6) = oRecordset.Fields(6)
          Me.lboMetricsData.List(Me.lboMetricsData.ListCount - 1, 7) = oRecordset.Fields(7)
          Me.lboMetricsData.List(Me.lboMetricsData.ListCount - 1, 8) = IIf(CLng(oRecordset.Fields(8)) = 0, "-", oRecordset.Fields(8))
          oRecordset.MoveNext
        Loop
      End If
    End If
  Next lngItem

exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetricsData_frm", "cmdDelete_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub cmdExport_Click()
  cptExportMetricsData
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub
