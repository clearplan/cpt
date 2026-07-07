VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptListBox_frm 
   Caption         =   "UserForm6"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "cptListBox_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptListBox_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.1</cpt_version>
Option Explicit
Private oListBoxData As ADODB.Recordset
Private Sub chkAll_Click()
  Dim lngItem As Long
  For lngItem = 0 To Me.lboListBox.ListCount - 1
    Me.lboListBox.Selected(lngItem) = Me.chkAll
  Next lngItem
End Sub

Private Sub cmdGo_Click()
  Dim lngItem As Long
  Dim lngSelected As Long
  For lngItem = 0 To Me.lboListBox.ListCount - 1
    If Me.lboListBox.Selected(lngItem) Then lngSelected = lngSelected + 1
  Next lngItem
  If lngSelected = 0 Then 'confirm
    If MsgBox("No selection(s) made." & vbCrLf & vbCrLf & "Proceed?", vbExclamation + vbYesNo, "No selection?") = vbNo Then
      Exit Sub
    End If
  End If
  Me.Hide
End Sub

Private Sub txtFilter_Change()
  Me.lboListBox.Clear
  If Len(Me.txtFilter.Text) > 0 Then
    oListBoxData.Filter = "ANCHOR_TEXT Like '%" & Me.txtFilter.Text & "%'"
  Else
    oListBoxData.Filter = 0
  End If
  If Not oListBoxData.EOF Then
    oListBoxData.MoveFirst
    Do While Not oListBoxData.EOF
      Me.lboListBox.AddItem oListBoxData(0)
      oListBoxData.MoveNext
    Loop
  End If
  oListBoxData.Filter = 0
End Sub

Private Sub UserForm_Activate()
  Dim lngItem As Long
  'todo: what if it has multiple columns?
  'todo: better plan is to make a recordset from the calling function
  'todo: so you can define columnwidths and searchcolumn
  If Me.lboListBox.ListCount > 0 And oListBoxData Is Nothing Then
    Set oListBoxData = CreateObject("ADODB.Recordset")
    oListBoxData.Fields.Append "ANCHOR_TEXT", adVarChar, 100
    oListBoxData.Open
    For lngItem = 0 To Me.lboListBox.ListCount - 1
      oListBoxData.AddNew Array(0), Array(Me.lboListBox.List(lngItem, 0))
    Next lngItem
  End If
End Sub

Private Sub UserForm_Terminate()
  Set oListBoxData = Nothing
End Sub
