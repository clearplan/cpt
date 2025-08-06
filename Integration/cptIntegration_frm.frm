VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptIntegration_frm 
   Caption         =   "Integration"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7020
   OleObjectBlob   =   "cptIntegration_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptIntegration_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.1.0</cpt_version>
Option Explicit
Public blnValidIntegrationMap As Boolean

Private Sub cboCA_Change()
  If Not Me.Visible Then Exit Sub
  'DO NOT SYNC WBS, OBS, CA with COBRA Export Tool because
  'CA1 is not always "WBS" (and it usually means CA anyway; and
  'CA2 is not always "OBS"; and
  'CA3 is not always CA; therefore
  'the DECM requires the WBS,OBS, and CA to be different fields.
  UpdateIntegrationSettings
End Sub

Private Sub cboCAM_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub cboEVTMS_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub cboEVP_Change()
  'objects
  Dim oCDP As DocumentProperty
  'strings
  Dim strCFN As String
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  If Not Me.Visible Then Exit Sub
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Me.chkSyncSettings Then
    
    On Error Resume Next
    Set oCDP = ActiveProject.CustomDocumentProperties("fPCNT")
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oCDP Is Nothing Then
      strCFN = ActiveProject.CustomDocumentProperties("fPCNT")
      lngField = FieldNameToFieldConstant(strCFN)
      If Me.cboEVP.Value <> lngField Then
        If MsgBox("COBRA Export Tool setting is '" & strCFN & "' - use this instead?", vbQuestion + vbYesNo, "Synchronize?") = vbYes Then
          Me.cboEVP.Value = lngField
        End If
      End If
    Else
      If Not IsNull(Me.cboEVP) Then
        Set oCDP = ActiveProject.CustomDocumentProperties.Add("fPCNT", False, msoPropertyTypeString, FieldConstantToFieldName(Me.cboEVP.Value))
      End If
    End If
  End If
  UpdateIntegrationSettings

exit_here:
  On Error Resume Next
  Set oCDP = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "cboEVT_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cboEVT_Change()
  'objects
  Dim oSubproject As MSProject.SubProject
  Dim oCDP As DocumentProperty
  Dim oDict As Scripting.Dictionary
  Dim oTask As MSProject.Task
  'strings
  Dim strValue As String
  'longs
  Dim lngField As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'If Not Me.Visible Then Exit Sub
  If IsNull(Me.cboEVT.Value) Then GoTo exit_here
  UpdateIntegrationSettings 'todo: can this be moved to only the end?
  Me.cboLOE.Value = ""
  Me.cboLOE.Clear
  Me.cboPP.Value = ""
  Me.cboPP.Clear
  Set oDict = CreateObject("Scripting.Dictionary")
  If ActiveProject.Subprojects.Count > 0 Then
    For Each oSubproject In ActiveProject.Subprojects
      For Each oTask In oSubproject.SourceProject.Tasks
        If oTask Is Nothing Then GoTo next_task_master
        If Not oTask.Active Then GoTo next_task_master
        strValue = oTask.GetField(Me.cboEVT.Value)
        If Len(strValue) > 0 Then
          If Not oDict.Exists(strValue) Then oDict.Add strValue, strValue
        End If
next_task_master:
      Next oTask
    Next oSubproject
  Else
    For Each oTask In ActiveProject.Tasks
      If oTask Is Nothing Then GoTo next_task_single
      If Not oTask.Active Then GoTo next_task_single
      strValue = oTask.GetField(Me.cboEVT.Value)
      If Len(strValue) > 0 Then
        If Not oDict.Exists(strValue) Then oDict.Add strValue, strValue
      End If
next_task_single:
    Next oTask
  End If
  Me.cboLOE.AddItem "<unused>"
  Me.cboPP.AddItem "<unused>"
  For lngItem = 0 To oDict.Count - 1
    Me.cboLOE.AddItem oDict.Items(lngItem)
    Me.cboPP.AddItem oDict.Items(lngItem)
  Next lngItem
  If Me.chkSyncSettings Then
    
    On Error Resume Next
    Set oCDP = ActiveProject.CustomDocumentProperties("fEVT")
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oCDP Is Nothing Then
      lngField = FieldNameToFieldConstant(ActiveProject.CustomDocumentProperties("fEVT"))
      If Me.cboEVT.Value <> lngField Then
        If MsgBox("COBRA Export Tool setting is '" & CustomFieldGetName(lngField) & "' - use this instead?", vbQuestion + vbYesNo, "Synchronize?") = vbYes Then
          Me.cboEVT.Value = lngField
        End If
      End If
    Else
      Set oCDP = ActiveProject.CustomDocumentProperties.Add("fEVT", False, msoPropertyTypeString, FieldConstantToFieldName(Me.cboEVT.Value))
    End If
  End If
  UpdateIntegrationSettings
  
exit_here:
  On Error Resume Next
  Set oSubproject = Nothing
  Set oDict = Nothing
  Set oCDP = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "cboEVT_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cboLOE_Change()
  If Not Me.Visible Then Exit Sub
  If Me.cboLOE.Value = "" Then
    Me.cboLOE.BorderColor = 192
  Else
    Me.cboLOE.BorderColor = -2147483642
    cptSaveSetting "Metrics", "txtLOE", Me.cboLOE.Value
    cptSaveSetting "Integration", "LOE", Me.cboLOE.Value
  End If
End Sub

Private Sub cboOBS_Change()
  If Not Me.Visible Then Exit Sub
  'DO NOT SYNC WBS, OBS, CA with COBRA Export Tool because
  'CA1 is not always "WBS" (and it usually means CA anyway; and
  'CA2 is not always "OBS"; and
  'CA3 is not always CA; therefore
  'the DECM requires the WBS,OBS, and CA to be different fields.
  UpdateIntegrationSettings
End Sub

Private Sub cboPP_Change()
  If Not Me.Visible Then Exit Sub
  If Me.cboPP.Value = "" Then
    Me.cboPP.BorderColor = 192
  Else
    Me.cboPP.BorderColor = -2147483642
    cptSaveSetting "Integration", "PP", Me.cboPP.Value
  End If
End Sub

Private Sub cboWBS_Change()
  If Not Me.Visible Then Exit Sub
  'DO NOT SYNC WBS, OBS, CA with COBRA Export Tool because
  'CA1 is not always "WBS" (and it usually means CA anyway; and
  'CA2 is not always "OBS"; and
  'CA3 is not always CA; therefore
  'the DECM requires the WBS,OBS, and CA to be different fields.
  UpdateIntegrationSettings
End Sub

Private Sub cboWP_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub cboWPM_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub chkECF_Click()
  'objects
  Dim oComboBox As MSForms.ComboBox
  'strings
  Dim strLOE As String
  Dim strPP As String
  'longs
  Dim lngItem As Long
  Dim lngKeep As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  Dim blnECF As Boolean
  'variants
  Dim vAddField As Variant
  Dim vFields As Variant
  Dim vControl As Variant
  'dates
  
  If Not Me.Visible Then GoTo exit_here
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  blnECF = CBool(Me.chkECF.Value)
  'first, save the changed setting
  cptSaveSetting "Integration", "chkECF", IIf(Me.chkECF, "1", "0")
  
  'update WBS,OBS,CA,WP,EVT,EVTMS
  If Me.cboLOE <> "" Then strLOE = Me.cboLOE.Value
  If Me.cboPP <> "" Then strPP = Me.cboPP.Value
  vFields = cptSortedArray(cptGetCustomFields("t", "Outline Code,Text", "c,cfn,loc", blnECF), 1)
  For Each vControl In Split("WBS,OBS,CA,WP,EVT,EVTMS", ",")
    Set oComboBox = Me.Controls("cbo" & vControl)
    lngKeep = 0
    If Not IsNull(oComboBox.Value) Then lngKeep = oComboBox.Value
    oComboBox.Clear
    If blnECF Then
      oComboBox.ColumnCount = 3
      oComboBox.ColumnWidths = "0 pt;105 pt;10 pt"
      oComboBox.ListWidth = 140
    Else
      oComboBox.ColumnCount = 2
      oComboBox.ColumnWidths = "0 pt"
      oComboBox.ListWidth = oComboBox.Width
    End If
    For lngItem = 0 To UBound(vFields)
      oComboBox.AddItem
      oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
      oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
      If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = vFields(lngItem, 2)
    Next lngItem
    If lngKeep > 0 Then oComboBox.Value = lngKeep
    If vControl = "EVT" Then
      On Error Resume Next
      Me.cboLOE.Value = strLOE
      Me.cboPP.Value = strPP
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    End If
  Next vControl
  
  'update CAM,WPM
  vFields = cptSortedArray(cptGetCustomFields("t", "Text,Outline Code", "c,cfn,loc", blnECF), 1)
  For Each vControl In Split("CAM,WPM", ",")
    Set oComboBox = Me.Controls("cbo" & vControl)
    lngKeep = 0
    If Not IsNull(oComboBox.Value) Then lngKeep = oComboBox.Value
    oComboBox.Clear
    If blnECF Then
      oComboBox.ColumnCount = 3
      oComboBox.ColumnWidths = "0 pt;105 pt;10 pt"
      oComboBox.ListWidth = 140
    Else
      oComboBox.ColumnCount = 2
      oComboBox.ColumnWidths = "0 pt"
      oComboBox.ListWidth = oComboBox.Width
    End If
    oComboBox.AddItem
    oComboBox.List(oComboBox.ListCount - 1, 0) = FieldNameToFieldConstant("Contact", pjTask)
    oComboBox.List(oComboBox.ListCount - 1, 1) = "Contact"
    If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = "LCF"
    For lngItem = 0 To UBound(vFields)
      oComboBox.AddItem
      oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
      oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
      If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = vFields(lngItem, 2)
    Next lngItem
    If lngKeep > 0 Then oComboBox.Value = lngKeep
  Next vControl
  
  'update EV%
  Set oComboBox = Me.cboEVP
  lngKeep = 0
  If Not IsNull(oComboBox) Then lngKeep = oComboBox.Value
  oComboBox.Clear
  If blnECF Then
    oComboBox.ColumnCount = 3
    oComboBox.ColumnWidths = "0 pt;105 pt;10 pt"
    oComboBox.ListWidth = 140
  Else
    oComboBox.ColumnCount = 2
    oComboBox.ColumnWidths = "0 pt"
    oComboBox.ListWidth = oComboBox.Width
  End If
  For Each vAddField In Split("Physical % Complete,% Complete", ",")
    oComboBox.AddItem
    oComboBox.List(oComboBox.ListCount - 1, 0) = FieldNameToFieldConstant(vAddField)
    oComboBox.List(oComboBox.ListCount - 1, 1) = vAddField
    If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = "LCF"
  Next vAddField
  vFields = cptSortedArray(cptGetCustomFields("t", "Number", "c,cfn,loc", False), 1)
  For lngItem = 0 To UBound(vFields)
    oComboBox.AddItem
    oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
    oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
    If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = vFields(lngItem, 2)
  Next lngItem
  If lngKeep > 0 Then Me.cboEVP.Value = lngKeep
  
exit_here:
  On Error Resume Next
  Set oComboBox = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "chkECF_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub chkSyncSettings_AfterUpdate()
  'objects
  Dim oCDP As DocumentProperty
  Dim oComboBox As MSForms.ComboBox
  Dim oDict As Scripting.Dictionary
  'strings
  Dim strFields As String
  Dim strCFN As String
  'longs
  Dim lngCFC As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vControl As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSaveSetting "Integration", "chkSyncSettings", IIf(Me.chkSyncSettings, "1", "0")
    
  strFields = "CAM,WP,EVT,EVP"
  'map integration settings with COBRA Export Tool settings
  Set oDict = CreateObject("Scripting.Dictionary")
  'DO NOT SYNC WBS, OBS, CA with COBRA Export Tool because
  'CA1 is not always "WBS" (and it usually means CA anyway; and
  'CA2 is not always "OBS"; and
  'CA3 is not always CA; therefore
  'the DECM requires the WBS,OBS, and CA to be different fields.
  oDict.Add "CAM", "fCAM"
  oDict.Add "WP", "fWP"
  oDict.Add "EVT", "fEVT"
  oDict.Add "EVP", "fPCNT"
  
  For Each vControl In Split(strFields, ",")
    Set oComboBox = Me.Controls("cbo" & vControl)
    If Me.chkSyncSettings Then
      If Not oComboBox.Enabled Then GoTo next_control
      On Error Resume Next
      Set oCDP = Nothing
      Set oCDP = ActiveProject.CustomDocumentProperties(oDict(vControl))
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If Not oCDP Is Nothing Then
        'does saved COBRA Export field still exist?
        strCFN = "" 'reset
        strCFN = oCDP.Value
        lngCFC = 0 'reset
        On Error Resume Next
        lngCFC = FieldNameToFieldConstant(strCFN)
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If lngCFC = 0 Then
          If MsgBox("The field '" & oCDP.Value & "' in the COBRA settings no longer exists in this file. Update with '" & strCFN & "'?", vbCritical + vbYesNo, "Mapping Invalid") = vbYes Then
            oCDP = strCFN
            oComboBox.BorderColor = -2147483642
          Else
            oComboBox.BorderColor = 192
            oCDP.Delete
          End If
        Else
          If IsNull(oComboBox) Then 'import from COBRA Export Tool setting
            oComboBox.Value = FieldNameToFieldConstant(oCDP.Value)
            cptSaveSetting "Integration", CStr(vControl), oComboBox.Value & "|" & oCDP.Value
            oComboBox.BorderColor = -2147483642
          Else 'notify discrepancy
            If FieldNameToFieldConstant(oCDP.Value) <> oComboBox.Value Then
              oComboBox.BorderColor = 192
            Else
              oComboBox.BorderColor = -2147483642
            End If
          End If
        End If
      Else 'CDP for COBRA Export setting doesn't exist yet
        oComboBox.BorderColor = 192
        lngCFC = 0
        lngCFC = oComboBox.Value
        strCFN = ""
        strCFN = CustomFieldGetName(lngCFC)
        If Len(strCFN) = 0 Then strCFN = FieldConstantToFieldName(lngCFC)
        If MsgBox("COBRA Export Setting for '" & IIf(vControl = "EVP", "EV%", vControl) & "' ('" & oDict(vControl) & "') does not exist - add it?", vbExclamation + vbYesNo, "Add Setting?") = vbYes Then
          Set oCDP = ActiveProject.CustomDocumentProperties.Add(oDict(vControl), False, msoPropertyTypeString, strCFN)
          oComboBox.BorderColor = -2147483642
        End If
        cptSaveSetting "Integration", CStr(vControl), lngCFC & "|" & strCFN
      End If
next_control:
      Set oCDP = Nothing
    Else
      oComboBox.BorderColor = -2147483642
    End If
  Next vControl
  
exit_here:
  Set oCDP = Nothing
  Set oComboBox = Nothing
  Set oDict = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "chkSyncSettings_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdCancel_Click()
  Me.blnValidIntegrationMap = False
  Me.Hide
End Sub

Private Sub cmdConfirm_Click()
  Dim blnValid As Boolean
  Dim oControl As MSForms.Control
  Dim strConstants As String
  Dim strKey As String
  Dim strValue As String
  
  blnValid = True
  For Each oControl In Me.Controls
    If Left(oControl.Name, 3) = "cmd" Then GoTo next_control
    If Left(oControl.Name, 3) = "chk" Then GoTo next_control
    If oControl.BorderColor = 192 Then
      blnValid = False
      Exit For
    Else
      If Left(oControl.Name, 3) = "cbo" Then
        If IsNull(oControl.Value) Then GoTo next_control
        If oControl.Value = "" Then GoTo next_control
        strKey = Replace(oControl.Name, "cbo", "")
        If oControl.Name = "cboLOE" Then
          strValue = oControl.Value
        ElseIf oControl.Name = "cboPP" Then
          strValue = oControl.Value
        Else
          strValue = CustomFieldGetName(oControl.Value)
          If Len(strValue) = 0 Then
            strValue = FieldConstantToFieldName(oControl.Value)
          End If
          strValue = oControl.Value & "|" & strValue
        End If
        cptSaveSetting "Integration", strKey, strValue
      End If
    End If
next_control:
    strKey = ""
    strValue = ""
  Next oControl
  'todo: if blnValid and me.chkSyncWithCOBRA and master/sub -> flowdown oCDPs?
'  todo: PMB missing data quicklook?
'  If blnValid Then
'    For Each oControl In Me.Controls
'      If Left(oControl.Name, 3) = "cbo" And oControl.Name <> "cboLOE" And oControl.Name <> "cboPP" Then
'        If Not IsNull(Me.Controls(oControl.Name).Value) And Me.Controls(oControl.Name).Enabled Then
'          strConstants = strConstants & Replace(oControl.Name, "cbo", "") & "|" & CustomFieldGetName(Me.Controls(oControl.Name)) & "|" & Me.Controls(oControl.Name).Value & ","
'        End If
'      End If
'    Next oControl
'    cptCheckMetadata strConstants, "strMissing"
'  End If
  Me.blnValidIntegrationMap = blnValid
  If Not blnValid Then
    If MsgBox("Cancel?", vbExclamation + vbYesNo, "Mapping Incomplete") = vbNo Then
      Exit Sub
    Else
      Me.Hide
    End If
  Else
    Me.Hide
  End If
End Sub

Private Sub UpdateIntegrationSettings()
  'objects
  Dim oCDP As Office.DocumentProperty
  'strings
  Dim strControl As String
  Dim strField As String
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vControl As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not Me.Visible Then Exit Sub
  strControl = Me.ActiveControl.Name
  If Left(strControl, 3) <> "cbo" Then GoTo exit_here
  If IsNull(Me.Controls(strControl).Value) Then GoTo exit_here
  lngField = Me.Controls(strControl).Value
  Me.Controls(strControl).BorderColor = -2147483642
  strControl = Replace(strControl, "cbo", "")
  strField = CustomFieldGetName(lngField)
  If Len(strField) = 0 Then strField = FieldConstantToFieldName(lngField)
  cptSaveSetting "Integration", strControl, lngField & "|" & strField
  'sync metrics settings - todo: is this still needed?
  If strControl = "EVT" Then
    cptSaveSetting "Metrics", "cboLOEField", lngField
  End If
  If strControl = "LOE" Then
    cptSaveSetting "Metrics", "txtLOE", Me.cboLOE.Value
  End If
  If strControl = "EVP" Then
    cptSaveSetting "Metrics", "cboEVP", lngField
  End If
  'validate against COBRA Export Tool
  'todo: does sync COBRA Export Tool conflict with strRequiredFields?
  If Me.chkSyncSettings Then
    For Each vControl In Split("CAM,WP,EVT,EVP", ",")
      strControl = CStr(vControl)
      If strControl = "EVP" Then strControl = "PCNT"
      On Error Resume Next
      Set oCDP = ActiveProject.CustomDocumentProperties("f" & strControl)
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oCDP Is Nothing Then GoTo next_control
      If IsNull(Me.Controls("cbo" & vControl).Value) Then 'populate it
        If Me.Controls("cbo" & vControl).Enabled Then Me.Controls("cbo" & vControl).Value = ActiveProject.CustomDocumentProperties("f" & strControl)
      ElseIf CustomFieldGetName(Me.Controls("cbo" & vControl).Value) <> ActiveProject.CustomDocumentProperties("f" & strControl) Then
        If FieldConstantToFieldName(Me.Controls("cbo" & vControl).Value) <> ActiveProject.CustomDocumentProperties("f" & strControl) Then 'catch if not a custom field (e.g., Physical % Complete)
          If Me.Controls("cbo" & vControl).Enabled Then Me.Controls("cbo" & vControl).BorderColor = 192
        Else
          If Me.Controls("cbo" & vControl).Enabled Then Me.Controls("cbo" & vControl).BorderColor = -2147483642
        End If
      Else
        If Me.Controls("cbo" & vControl).Enabled Then Me.Controls("cbo" & vControl).BorderColor = -2147483642
      End If
next_control:
      Set oCDP = Nothing
    Next vControl
  End If
  
'  For Each oControl In Me.Controls
'    'get constants, do metadatacheck, get lngCount
'    'if lngCount>0 then visible,indicate,populate listbox
'    'if tgl then visible lbo, form height, etc
'  Next oControl
  
exit_here:
  On Error Resume Next
  Set oCDP = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "UpdateIntegrationSettings", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtRollingWave_AfterUpdate()
  Dim dtRollingWave As Date
  If Len(Me.txtRollingWave) > 0 Then
    If IsDate(Me.txtRollingWave) Then
      dtRollingWave = CDate(Me.txtRollingWave.Value)
      Me.txtRollingWave = FormatDateTime(dtRollingWave, vbShortDate)
      cptSaveSetting "Integration", "RollingWaveDate", dtRollingWave
    Else
      Me.txtRollingWave = ""
    End If
  Else
    cptDeleteSetting "Integration", "RollingWaveDate"
  End If
  Me.lblWeekday.Visible = False
End Sub

Private Sub txtRollingWave_Change()
  Dim dtDate As Date
  
  If Not Me.Visible Then Exit Sub
  If Not Me.ActiveControl.Name = Me.txtRollingWave.Name Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.lblWeekday.Visible = False
  If Len(Me.txtRollingWave.Text) > 0 Then
    Me.txtRollingWave.Text = cptRegEx(Me.txtRollingWave.Text, "[0-9\/]{1,}")
    'limit to a date only
    If Not IsDate(Me.txtRollingWave.Text) Then
      Me.txtRollingWave.BorderColor = 192
      Me.Repaint
    Else
      'limit to dates after the start date
      dtDate = CDate(Format(Me.txtRollingWave.Text, "mm/dd/yyyy") & " 5:00 PM")
      If dtDate < ActiveProject.StatusDate Then
        Me.txtRollingWave.BorderColor = 192
        Me.Repaint
      Else
        Me.txtRollingWave.BorderColor = -2147483642
        Me.lblWeekday.Caption = Format(dtDate, "dddd")
        Me.lblWeekday.Visible = True
        Me.Repaint
      End If
    End If
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "txtRollingWave_Change", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtRollingWave_Enter()
  If IsDate(Me.txtRollingWave) Then
    Me.lblWeekday.Caption = Format(CDate(Me.txtRollingWave), "dddd")
    Me.lblWeekday.Visible = True
  End If
End Sub

Private Sub txtRollingWave_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  Me.lblWeekday.Visible = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.blnValidIntegrationMap = False
    Me.Hide
    Cancel = True
  End If
End Sub
