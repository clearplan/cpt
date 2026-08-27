VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptCritPathFields_frm 
   Caption         =   "cpt Driving Paths"
   ClientHeight    =   4968
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
'<cpt_version>v3.5.3</cpt_version>
Option Explicit
Private Const MODULE_NAME As String = "cptCritPathFields_frm"
Private Const cptSettingFeature As String = "Driving Paths" 'v3.5.0
Private Const cptViewSetting As String = "User View" 'v3.5.0
Private Const cptGanttSetting As String = "Gantt Format" 'v3.5.0
Private Const cptSubPathSetting As String = "SubPaths" 'v3.5.1
Private Const cptPathCountSetting As String = "Path Count" 'v3.5.2

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

Private Function The_Aron_Double_Check(ByVal fieldName As String) As Boolean

    If cptCustomFieldExists(fieldName) = 0 Then
        The_Aron_Double_Check = False
        Exit Function
    Else
        The_Aron_Double_Check = True
        Exit Function
    End If

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
    
    If The_Aron_Double_Check(PathField_Combobox.Text) = True Then
        cptStoreCustomFieldName "Driving Paths", PathField_Combobox.Text, FieldNameToFieldConstant(PathField_Combobox.Text)
    Else
        MsgBox "The selected Driving Paths field no longer exists." & vbCr & vbCr & "Please close and reopen the Critical Path tool."
        Exit Sub
    End If
    
    If The_Aron_Double_Check(GroupField_Combobox.Text) = True Then
        cptStoreCustomFieldName "Driving Path Group", GroupField_Combobox.Text, FieldNameToFieldConstant(GroupField_Combobox.Text)
    Else
        MsgBox "The selected Driving Path Group field no longer exists." & vbCr & vbCr & "Please close and reopen the Critical Path tool."
        Exit Sub
    End If
    
    If SubPath_Checkbox Then
        If The_Aron_Double_Check(SubPath_Combobox.Text) = True Then
        cptStoreCustomFieldName "SubPath Group", SubPath_Combobox.Text, FieldNameToFieldConstant(SubPath_Combobox.Text)
        Else
            MsgBox "The selected SubPath Group field no longer exists." & vbCr & vbCr & "Please close and reopen the Critical Path tool."
            Exit Sub
        End If
    End If
    
    'Store Field Names
    cptSaveSetting cptSettingFeature, cptViewSetting, UserView_Combobox.Text
    cptSaveSetting cptSettingFeature, cptGanttSetting, ganttFormatCheckBox.value
    cptSaveSetting cptSettingFeature, cptSubPathSetting, SubPath_Checkbox.value
    cptSaveSetting cptSettingFeature, cptPathCountSetting, pathCnt_txtBox.value
    
    Me.Tag = "run"
    Me.Hide
    
    Exit Sub
    
End Sub

Private Sub SubPath_Checkbox_Click()
    SubPath_Combobox.Enabled = SubPath_Checkbox.value
End Sub

Private Sub SubPath_Combobox_Change()
    If checkDuplicate(SubPath_Combobox) = True Then
        MsgBox "Please select a unique IMS Field."
        SubPath_Combobox.ListIndex = 0
        Exit Sub
    End If
End Sub

Private Sub UserForm_Activate()

    Dim settingTest As String
    
    settingTest = cptGetSetting(cptSettingFeature, cptSubPathSetting)
    
    If settingTest <> "" Then
        SubPath_Checkbox.value = CBool(settingTest)
        SubPath_Combobox.Enabled = CBool(settingTest)
    Else
        SubPath_Checkbox.value = False
        SubPath_Combobox.Enabled = False
    End If
    
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
    
    settingTest = cptGetSetting(cptSettingFeature, cptPathCountSetting)
    
    If settingTest <> "" Then
        pathCnt_txtBox.value = settingTest
    End If

End Sub

Private Sub UserForm_Initialize()

    Dim drivingPathField As String
    Dim groupPathField As String
    Dim SubPathField As String

    drivingPathField = cptGetCustomFieldName("Driving Paths")
    groupPathField = cptGetCustomFieldName("Driving Path Group")
    SubPathField = cptGetCustomFieldName("SubPath Group")
    
    DisplayUserCustomFields drivingPathField, groupPathField, SubPathField
    
End Sub

Private Sub DisplayUserCustomFields(ByVal drivingPathField As String, ByVal groupPathField As String, ByVal SubPathField As String)

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
    
    nameTest = FieldNameToFieldConstant(SubPathField)
    
    If nameTest <> 0 Then
        SubPath_Combobox.value = SubPathField
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    Me.Tag = "cancel"
    Me.Hide
  End If
End Sub
