VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptAdvancedFilterEdit_frm 
   Caption         =   "UserForm1"
   ClientHeight    =   1644
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   4815
   OleObjectBlob   =   "cptAdvancedFilterEdit_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptAdvancedFilterEdit_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'<cpt_version>v0.3.2</cpt_version>
Private Sub closeBtn_Click()
    Me.Tag = "Close"
    Me.Hide
End Sub

Private Sub editBtn_Click()
    If Trim(Me.itemValue_TextBox.Value = "") Then
        MsgBox "Please enter a value."
        Exit Sub
    End If
    Me.Tag = "Edit"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    closeBtn_Click
  End If
End Sub
