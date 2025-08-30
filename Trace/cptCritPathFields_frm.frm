VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptCritPathFields_frm 
   Caption         =   "cpt Driving Paths"
   ClientHeight    =   3060
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4068
   OleObjectBlob   =   "cptCritPathFields_frm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cptCritPathFields_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v3.4.1</cpt_version>
Option Explicit
Private Const MODULE_NAME As String = "cptCritPathFields_frm"

Private Sub RunBtn_Click()
        
    If PathField_Combobox.Text = "" Or GroupField_Combobox.Text = "" Then
        MsgBox "Please complete the required field mapping."
        Exit Sub
    End If
    
    If Not IsNumeric(pathCnt_txtBox.Value) Then
        MsgBox "Please enter a valid Path Count number."
        Exit Sub
    End If
    
    cptStoreCustomFieldName "Driving Paths", "CP Driving Paths", FieldNameToFieldConstant(PathField_Combobox.Text)
    cptStoreCustomFieldName "Driving Path Group", "CP Driving Path Group ID", FieldNameToFieldConstant(GroupField_Combobox.Text)
    
    'Store Field Names
    On Error GoTo Driving_FieldExists
    CustomFieldRename FieldID:=FieldNameToFieldConstant(PathField_Combobox.Text), NewName:="CP Driving Paths"
    
Group_Field_Rename:
    
    On Error GoTo Group_FieldExists
    CustomFieldRename FieldID:=FieldNameToFieldConstant(GroupField_Combobox.Text), NewName:="CP Driving Path Group ID"
    
End_Field_Rename:
    
    Me.Tag = "run"
    Me.Hide
    
    Exit Sub
    
Driving_FieldExists:

    CustomFieldRename FieldID:=FieldNameToFieldConstant("CP Driving Paths"), NewName:="CP Driving Paths_" & FieldNameToFieldConstant("CP Driving Paths")
    CustomFieldRename FieldID:=FieldNameToFieldConstant(PathField_Combobox.Text), NewName:="CP Driving Paths"
    
    Resume Group_Field_Rename
    
Group_FieldExists:

    CustomFieldRename FieldID:=FieldNameToFieldConstant("CP Driving Path Group ID"), NewName:="CP Driving Path Group ID_" & FieldNameToFieldConstant("CP Driving Path Group ID")
    CustomFieldRename FieldID:=FieldNameToFieldConstant(GroupField_Combobox.Text), NewName:="CP Driving Path Group ID"

    Resume End_Field_Rename
    
End Sub

Private Sub UserForm_Initialize()

    Dim drivingPathField As String
    Dim groupPathField As String

    drivingPathField = cptGetCustomFieldName("Driving Paths")
    groupPathField = cptGetCustomFieldName("Driving Path Group")
    
    DisplayUserCustomFields drivingPathField, groupPathField
    
End Sub

Private Sub DisplayUserCustomFields(ByVal drivingPathField As String, ByVal groupPathField As String)

    Dim nameTest As Long
    
    nameTest = 0
    
    On Error Resume Next
    
    nameTest = FieldNameToFieldConstant(drivingPathField)
    
    If nameTest <> 0 Then
        PathField_Combobox.Value = drivingPathField
    End If

    nameTest = FieldNameToFieldConstant(groupPathField)
    
    If nameTest <> 0 Then
        GroupField_Combobox.Value = groupPathField
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    Me.Tag = "cancel"
    Me.Hide
  End If
End Sub


