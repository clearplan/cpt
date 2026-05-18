VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptCritPathFields_frm 
   Caption         =   "cpt Driving Paths"
   ClientHeight    =   3810
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4065
   OleObjectBlob   =   "cptCritPathFields_frm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cptCritPathFields_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v3.5.0</cpt_version>
Option Explicit
Private Const MODULE_NAME As String = "cptCritPathFields_frm"
Private Const cptSettingFeature As String = "Driving Paths" 'v3.5.0
Private Const cptViewSetting As String = "User View" 'v3.5.0
Private Const cptGanttSetting As String = "Gantt Format" 'v3.5.0

Private Sub GroupField_Combobox_Change()
    If checkDuplicate(GroupField_Combobox) = True Then
        MsgBox "Please select a unique IMS Field."
        GroupField_Combobox.ListIndex = 0
        Exit Sub
    End If
End Sub

Private Sub PathField_Combobox_Change()
    If checkDuplicate(PathField_Combobox) = True Then
        MsgBox "Please select a unique IMS Field."
        PathField_Combobox.ListIndex = 0
        Exit Sub
    End If
End Sub

Private Function checkDuplicate(ByVal cBoxTest As MSForms.ComboBox) As Boolean 'v3.5.0

    If cBoxTest.value = "" Then
    
        checkDuplicate = False
        Exit Function
    
    End If

    Dim cBoxOther As MSForms.ComboBox 'v3.3.8
    Dim formObj As MSForms.Control 'v3.3.8
    
    For Each formObj In Me.Controls
    
        If TypeName(formObj) = "ComboBox" Then
        
            Set cBoxOther = formObj
            
            If cBoxOther.Name <> cBoxTest.Name Then
            
                If cBoxOther.value = cBoxTest.value Then
                
                    checkDuplicate = True
                    Exit Function
                
                End If
            
            End If
        
        End If
    
    Next formObj
    
    checkDuplicate = False

End Function

Private Sub RunBtn_Click()
        
    If PathField_Combobox.Text = "" Or GroupField_Combobox.Text = "" Then
        MsgBox "Please complete the required field mapping."
        Exit Sub
    End If
    
    If Not IsNumeric(pathCnt_txtBox.value) Then
        MsgBox "Please enter a valid Path Count number."
        Exit Sub
    End If
    
    cptStoreCustomFieldName "Driving Paths", "CP Driving Paths", FieldNameToFieldConstant(PathField_Combobox.Text)
    cptStoreCustomFieldName "Driving Path Group", "CP Driving Path Group ID", FieldNameToFieldConstant(GroupField_Combobox.Text)
    
    'Store Field Names
    cptSaveSetting cptSettingFeature, cptViewSetting, UserView_Combobox.Text
    cptSaveSetting cptSettingFeature, cptGanttSetting, ganttFormatCheckBox.value
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

Private Sub UserForm_Activate()

    Dim settingTest As String
    settingTest = cptGetSetting(cptSettingFeature, cptGanttSetting)
    
    If settingTest <> "" Then
        ganttFormatCheckBox.value = CBool(settingTest)
    Else
        ganttFormatCheckBox.value = True
    End If
    
    settingTest = cptGetSetting(cptSettingFeature, cptViewSetting)
    
    If settingTest <> "" Then
        Me.UserView_Combobox.value = settingTest
    End If

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
        PathField_Combobox.value = drivingPathField
    End If

    nameTest = FieldNameToFieldConstant(groupPathField)
    
    If nameTest <> 0 Then
        GroupField_Combobox.value = groupPathField
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    Me.Tag = "cancel"
    Me.Hide
  End If
End Sub


