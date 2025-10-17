VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptAdjustment_frm 
   Caption         =   "ETC Adjustment"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6330
   OleObjectBlob   =   "cptAdjustment_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptAdjustment_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.4</cpt_version>
Option Explicit

Private Sub cboResources_Change()
  
  If Me.Visible Then Call cptRefreshAdjustment(Me)
  
End Sub

Private Sub cmdApply_Click()

  'require an amount
  Me.txtAmount.BorderColor = -2147483642
  If IsNull(Me.txtAmount) Or Len(Me.txtAmount) = 0 Then
    Me.txtAmount.BorderColor = 192
    Exit Sub
  End If
  
  cptApplyAdjustment Me
  cptRefreshAdjustment Me
    
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdUndo_Click()
  Application.Undo
  cptRefreshAdjustment Me
End Sub

Private Sub optDelta_Click()
  Me.txtAmount.ControlTipText = "Add/Reduce by set number of hours"
  'stick to apportioning by remaining work
  cptRefreshAdjustment Me
End Sub

Private Sub optPercent_Click()
  Me.txtAmount.ControlTipText = "Please use decimal format"
  'stick to apportioning by remaining work
  cptRefreshAdjustment Me
End Sub

Private Sub optTarget_Click()
  Me.txtAmount.ControlTipText = "Apportion to hit Target"
  'stick to apportioning by remaining work
  If Me.Visible Then cptRefreshAdjustment Me
End Sub

Private Sub txtAmount_Change()
  'objects
  'strings
  Dim strAmount As String
  'longs
  'integers
  'doubles
  Dim dblAmount As Double
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strAmount = Me.txtAmount.Text
  If strAmount = "-" Then GoTo exit_here 'be patient
  If strAmount = "." Or Right(strAmount, 1) = "." Then GoTo exit_here 'be patient
  If strAmount = "-." Then GoTo exit_here 'be patient
  If Len(strAmount) > 0 Then
    If Me.optTarget Then
      dblAmount = cptRegEx(strAmount, "(-)?([0-9]{1,})?(\.[0-9]{1,})?")
      If dblAmount < 0 Then
        dblAmount = Abs(dblAmount)
        Me.txtAmount = dblAmount
      End If
    ElseIf Me.optPercent Then
      If Right(strAmount, 1) = "%" Then
        strAmount = Replace(strAmount, "%", "")
        dblAmount = cptRegEx(strAmount, "(-)?([0-9]{1,})?(\.[0-9]{1,})?")
        Me.txtAmount.Text = dblAmount / 100
      Else
        Me.txtAmount.Text = cptRegEx(strAmount, "(-)?([0-9]{1,})?(\.[0-9]{1,})?")
      End If
    Else
      Me.txtAmount.Text = cptRegEx(strAmount, "(-)?([0-9]{1,})?(\.[0-9]{1,})?")
    End If
  End If
  cptRefreshAdjustment Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_frm", "txtAmount_Change", Err, Erl)
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
