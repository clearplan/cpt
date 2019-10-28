VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptStatusSheetImport_frm 
   Caption         =   "Import Status Sheets"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810
   OleObjectBlob   =   "cptStatusSheetImport_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptStatusSheetImport_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboAF_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub cboAS_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub cboETC_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub cboEV_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub cboFF_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub cboFS_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub chkAppend_Click()
    Me.cboAppendTo.Enabled = Me.chkAppend
  If Me.chkAppend Then
    Me.cboAppendTo.SetFocus
    Me.cboAppendTo.DropDown
  End If
End Sub

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub cmdImport_Click()
  Call cptStatusSheetImport
End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "lblURL", err, Erl)
  Resume exit_here
End Sub

Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, _
                                  Effect As Long, _
                                  Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
  Call cptAddFiles(Data)
End Sub

Private Sub UserForm_Initialize()
  Me.TreeView1.OLEDropMode = ccOLEDropManual
End Sub