VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptCritPathFields_frm 
   Caption         =   "cpt Driving Paths"
   ClientHeight    =   3525
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6405
   OleObjectBlob   =   "cptCritPathFields_frm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cptCritPathFields_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v3.5.5</cpt_version>
Option Explicit
Private Const MODULE_NAME As String = "cptCritPathFields_frm"
Private Const lngBorderColorValid As Long = 8421504 '-2147483642
Private Const lngBorderColorInvalid As Long = 192

Private Sub ganttFormatCheckBox_Click()
  If Not Me.Visible Then Exit Sub
  If Me.ganttFormatCheckBox And Me.UserView_Combobox.Value <> "<Default>" Then
    MsgBox "Gradient coloring on paths is only available with the <Default> view.", vbInformation + vbOKOnly, "Driving Path"
  End If
  Me.UserView_Combobox.Enabled = Not Me.ganttFormatCheckBox
  cptSaveSetting "DrivingPath", "chkGanttFormat", IIf(Me.ganttFormatCheckBox, 1, 0)
End Sub

Private Sub GroupField_Combobox_Change()
  Dim strSetting As String
  If Not Me.Visible Then Exit Sub
  If checkDuplicate(GroupField_Combobox) = True Then
    Me.GroupField_Combobox.BorderColor = lngBorderColorInvalid
    MsgBox "Please select a unique IMS Field.", vbExclamation + vbOKOnly, "Driving Path"
    Exit Sub
  Else
    strSetting = Me.GroupField_Combobox.Value
    Me.GroupField_Combobox.BorderColor = lngBorderColorValid
    If cptCustomFieldExists(strSetting) Then
      cptSaveSetting "DrivingPath", "cboGroup", FieldNameToFieldConstant(strSetting) & "|" & strSetting
    End If
  End If
End Sub

Private Sub pathCnt_txtBox_Change()
  Dim strPathCount As String
  Dim lngPathCount As Long
  If Not Me.Visible Then Exit Sub
  strPathCount = Me.pathCnt_txtBox.Text
  If Len(strPathCount) > 0 Then
    If cptRxMatch(strPathCount, "[0-9]+") Then
      lngPathCount = cptRxMatch(Me.pathCnt_txtBox.Text, "[0-9]+")
      If lngPathCount = 0 Then
        MsgBox "Path Count must be greater than zero.", vbCritical + vbOKOnly, "Driving Path"
        Me.pathCnt_txtBox.Text = 1
      Else
        Me.pathCnt_txtBox.Text = lngPathCount
      End If
    Else
      Me.pathCnt_txtBox.Text = 1
    End If
    cptSaveSetting "DrivingPath", "txtPathCount", lngPathCount
  End If
End Sub

Private Sub PathField_Combobox_Change()
  Dim strSetting As String
  If Not Me.Visible Then Exit Sub
  If checkDuplicate(PathField_Combobox) = True Then
    Me.GroupField_Combobox.BorderColor = lngBorderColorInvalid
    MsgBox "Please select a unique IMS Field.", vbExclamation + vbOKOnly, "Driving Path"
    Exit Sub
  Else
    Me.GroupField_Combobox.BorderColor = lngBorderColorValid
    strSetting = Me.PathField_Combobox.Value
    If cptCustomFieldExists(strSetting) Then
      cptSaveSetting "DrivingPath", "cboPath", FieldNameToFieldConstant(strSetting) & "|" & strSetting
    End If
  End If
End Sub

Private Function checkDuplicate(ByVal cBoxTest As MSForms.ComboBox) As Boolean 'v3.5.0

    cBoxTest.BorderColor = lngBorderColorValid
    If cBoxTest.Value = "" Then
    
        checkDuplicate = False
        Exit Function
    
    End If

    Dim cBoxOther As MSForms.ComboBox 'v3.3.8
    Dim formObj As MSForms.Control 'v3.3.8
    
    For Each formObj In Me.Controls
    
        If TypeName(formObj) = "ComboBox" Then
        
            Set cBoxOther = formObj
            
            If cBoxOther.Name <> cBoxTest.Name Then
            
                If cBoxOther.Value = cBoxTest.Value Then

                    cBoxTest.BorderColor = lngBorderColorInvalid
                    checkDuplicate = True
                    Exit Function
                
                End If
            
            End If
        
        End If
    
    Next formObj
    
    checkDuplicate = False

End Function

Private Sub RunBtn_Click()
    
    Dim lngBorderColorValid As Long
    Dim lngBorderColorInvalid As Long
    
    'validate selections
    If PathField_Combobox.Text = "" Or GroupField_Combobox.Text = "" _
        Or (Me.SubPath_Checkbox = True And Me.SubPath_Combobox = "") Then
        MsgBox "Please complete the required field mapping."
        Exit Sub
    End If
    
    If Me.PathField_Combobox.BorderColor = lngBorderColorInvalid Then Exit Sub
    If Me.GroupField_Combobox.BorderColor = lngBorderColorInvalid Then Exit Sub
    If Me.SubPath_Combobox.BorderColor = lngBorderColorInvalid Then Exit Sub
    
    If Not IsNumeric(pathCnt_txtBox.Value) Then
        MsgBox "Please enter a valid Path Count number."
        Exit Sub
    End If
        
    On Error GoTo Driving_FieldExists
    CustomFieldRename FieldID:=FieldNameToFieldConstant(PathField_Combobox.Text), NewName:="CP Driving Paths"
    
Group_Field_Rename:
    
    On Error GoTo Group_FieldExists
    CustomFieldRename FieldID:=FieldNameToFieldConstant(GroupField_Combobox.Text), NewName:="CP Driving Path Group ID"
    
End_Field_Rename:

    On Error GoTo SubPath_FieldExists
    If SubPath_Checkbox Then
        CustomFieldRename FieldID:=FieldNameToFieldConstant(SubPath_Combobox.Text), NewName:="CP SubPath Group ID"
    End If
    
SubPath_Rename:
    
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
    
SubPath_FieldExists:

    CustomFieldRename FieldID:=FieldNameToFieldConstant("CP SubPath Group ID"), NewName:="CP SubPath Group ID_" & FieldNameToFieldConstant("CP SubPath Group ID")
    CustomFieldRename FieldID:=FieldNameToFieldConstant(SubPath_Combobox.Text), NewName:="CP SubPath Group ID"
    
    Resume SubPath_Rename
    
End Sub

Private Sub SubPath_Checkbox_Click()
  If Not Me.Visible Then Exit Sub
  If SubPath_Checkbox.Value Then
    SubPath_Checkbox.Caption = "Code Sub-Paths in field:"
  Else
    SubPath_Checkbox.Caption = "Code Sub-Paths"
  End If
  SubPath_Combobox.Enabled = SubPath_Checkbox.Value
  cptSaveSetting "DrivingPath", "chkSubPaths", IIf(Me.SubPath_Checkbox, 1, 0)
End Sub

Private Sub SubPath_Combobox_Change()
  Dim strSetting As String
  If Not Me.Visible Then Exit Sub
  If checkDuplicate(SubPath_Combobox) = True Then
    Me.SubPath_Combobox.BorderColor = lngBorderColorInvalid
    MsgBox "Please select a unique IMS Field."
    Exit Sub
  Else
    Me.SubPath_Combobox.BorderColor = lngBorderColorValid
    strSetting = Me.SubPath_Combobox.Value
    If cptCustomFieldExists(strSetting) Then
      cptSaveSetting "DrivingPath", "cboSubPath", FieldNameToFieldConstant(strSetting) & "|" & strSetting
    End If
  End If
End Sub

Private Sub UserForm_Activate()
'  Dim blnValidSettings As Boolean
'
'  blnValidSettings = True
'
'  'confirm cboPath
'  With Me.PathField_Combobox
'    If cptCustomFieldExists(.Value) = 0 Then
'      blnValidSettings = False
'      .BorderColor = lngBorderColorInvalid
'    Else
'      .BorderColor = lngBorderColorValid
'    End If
'  End With
'
'  'validate cboGroup
'  With Me.GroupField_Combobox
'    If cptCustomFieldExists(.Value) = 0 Then
'      blnValidSettings = False
'      .BorderColor = lngBorderColorInvalid
'    Else
'      .BorderColor = lngBorderColorValid
'    End If
'  End With
'
'  'validate cboSubPath
'  If Me.SubPath_Checkbox Then
'    With Me.SubPath_Combobox
'      If cptCustomFieldExists(.Value) = 0 Then
'        blnValidSettings = False
'        .BorderColor = lngBorderColorInvalid
'      Else
'        .BorderColor = lngBorderColorValid
'      End If
'    End With
'  End If
'
'  'validate cboUserView
'  If Me.ganttFormatCheckBox Then
'    With Me.UserView_Combobox
'      If .Value <> "<Default>" Then
'        If Not cptViewExists(.Value) Then
'          blnValidSettings = False
'          .BorderColor = lngBorderColorInvalid
'        Else
'          .BorderColor = lngBorderColorValid
'        End If
'      End If
'    End With
'  End If
'
'  'notify
'  If Not blnValidSettings Then
'    MsgBox "Some of your settings are no longer valid." & vbCrLf & vbCrLf & vbCrLf & "Please review and correct before proceeding.", vbExclamation + vbOKOnly, "Driving Paths"
'  End If

End Sub

Private Sub UserForm_Initialize()

'    Dim drivingPathField As String
'    Dim groupPathField As String
'    Dim SubPathField As String
'
'    drivingPathField = cptGetCustomFieldName("Driving Paths")
'    groupPathField = cptGetCustomFieldName("Driving Path Group")
'    SubPathField = cptGetCustomFieldName("SubPath Group")
'
'    DisplayUserCustomFields drivingPathField, groupPathField, SubPathField
    
End Sub

Private Sub DisplayUserCustomFields(ByVal drivingPathField As String, ByVal groupPathField As String, ByVal SubPathField As String)

'    Dim nameTest As Long
'    Dim fieldsRenamed As Boolean
'
'    nameTest = 0
'    fieldsRenamed = False
'
'    On Error Resume Next
'
'    If Len(drivingPathField) = 0 Then GoTo CheckGroupPathField
'
'    nameTest = cptCustomFieldExists(drivingPathField)
'
'    If nameTest <> 0 Then
'        If CustomFieldGetName(nameTest) = drivingPathField Then
'            PathField_Combobox.Value = drivingPathField
'            GoTo CheckGroupPathField
'        End If
'    End If
'
'    fieldsRenamed = True
'
'CheckGroupPathField:
'
'    If Len(groupPathField) = 0 Then GoTo CheckSubPathField
'
'    nameTest = cptCustomFieldExists(groupPathField)
'
'    If nameTest <> 0 Then
'        If CustomFieldGetName(nameTest) = groupPathField Then
'            GroupField_Combobox.Value = groupPathField
'            GoTo CheckSubPathField
'        End If
'    End If
'
'    fieldsRenamed = True
'
'CheckSubPathField:
'
'  Dim blnSubPaths As Boolean
'  blnSubPaths = cptGetSetting("Driving Paths", "SubPaths") = True
'  If blnSubPaths Then
'
'    If Len(SubPathField) = 0 Then GoTo CheckFieldsRenamed
'
'    nameTest = cptCustomFieldExists(SubPathField)
'
'    If nameTest <> 0 Then
'        If CustomFieldGetName(nameTest) = SubPathField Then
'            SubPath_Combobox.Value = SubPathField
'            GoTo CheckFieldsRenamed
'        End If
'    End If
'
'    fieldsRenamed = True
'
'  End If
'
'CheckFieldsRenamed:
'
'    If fieldsRenamed Then
'        MsgBox "Some custom fields have been renamed." & vbCr & vbCr & "Please review and update the field mapping."
'    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    Me.Tag = "cancel"
    Me.Hide
  End If
End Sub

Private Sub UserView_Combobox_Change()
  Dim strSetting As String
  If Not Me.Visible Then Exit Sub
  strSetting = Me.UserView_Combobox.Value
  Me.UserView_Combobox.BorderColor = lngBorderColorValid
  If strSetting <> "<Default>" Then
    If cptViewExists(strSetting) Then
      cptSaveSetting "DrivingPath", "cboUserView", strSetting
    Else
      Me.UserView_Combobox.BorderColor = lngBorderColorInvalid
      MsgBox "The view named '" & strSetting & "' no longer exists.", vbExclamation + vbOKOnly, "Driving Path"
    End If
  Else
    cptSaveSetting "DrivingPath", "cboUserView", strSetting
  End If
End Sub
