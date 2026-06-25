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
'<cpt_version>v1.2.0</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptMetricsData_frm"

Private Sub cmdDelete_Click()
  cptDeleteMetricsData Me
End Sub

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub cmdExport_Click()
  cptExportMetricsData
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
  Call cptHandleErr(THIS_MODULE, "lblURL_Click()", Err, Erl)
  Resume exit_here

End Sub

Private Sub optDetail_Click()
  Me.lblWait.Visible = True
  DoEvents
  cptUpdateMetricsDataForm Me
  Me.lblWait.Visible = False
End Sub

Private Sub optSummary_Click()
  cptUpdateMetricsDataForm Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub
