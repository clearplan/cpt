VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSettings_frm 
   Caption         =   "cpt Settings"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   OleObjectBlob   =   "cptSettings_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptSettings_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.4.0</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptSettings_frm"

Private Sub cboCheckForUpdates_Change()
  If Not IsNull(Me.cboCheckForUpdates.Value) Then
    cptSaveSetting "General", "cboCheckForUpdates", Me.cboCheckForUpdates.Value
    cptUpdateSetting "General", "cboCheckForUpdates", Me.cboCheckForUpdates.Value
  End If
End Sub

Private Sub chkUpdateIndicator_Click()
  cptSaveSetting "General", "chkUpdateIndicator", IIf(Me.chkUpdateIndicator, 1, 0)
  'hacky way to trigger a ribbon refresh
  Dim oProject As MSProject.Project
  Set oProject = ActiveProject
  cptSpeed True
  FileNew
  DoEvents
  FileClose pjDoNotSave
  DoEvents
  oProject.Activate
  Set oProject = Nothing
End Sub

Private Sub cmdDone_Click()
  Me.Hide
End Sub

Private Sub cmdEdit_Click()
  Dim strMsg As String
  Dim strFileName As String
  strMsg = "...unless you *really* know what you're doing." & vbCrLf & vbCrLf
  strMsg = strMsg & "Contact help@ClearPlanConsulting.com if you need help." & vbCrLf & vbCrLf
  strMsg = strMsg & "Do you still wish to venture forth?"
  If MsgBox(strMsg, vbCritical + vbYesNo, "Do Not Attempt This...") = vbYes Then
    MsgBox "...you've been warned.", vbInformation + vbOKOnly, "OK"
    strFileName = Replace(Me.lblDir.Caption, "%USERPROFILE%", Environ("USERPROFILE")) 'cptDir & "\settings\cpt-settings.ini"
    cptShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
  End If
End Sub

Private Sub cmdSetProgramAcronym_Click()
  'objects
  Dim oRecordset As ADODB.Recordset
  Dim oDocProp As DocumentProperty
  Dim oDocProps As DocumentProperties
  Dim oSettingsFiles As Scripting.Dictionary
  'strings
  Dim strDir As String
  Dim strUpdated As String
  Dim strFileName As String
  Dim strFieldName As String
  Dim strOld As String
  Dim strNew As String
  'longs
  Dim lngItem As Long
  Dim lngUpdated As Long
  Dim lngResponse As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  Set oDocProps = ActiveProject.CustomDocumentProperties
  On Error Resume Next
  Set oDocProp = oDocProps("cptProgramAcronym")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  If oDocProp Is Nothing Then
    oDocProps.Add "cptProgramAcronym", False, msoPropertyTypeString, Me.txtProgramAcronym.Value, False
   Else
    strOld = oDocProp.Value
    strNew = Me.txtProgramAcronym
    oDocProp.Value = strNew
    lngResponse = MsgBox("Also replace all instances of '" & strOld & "' with '" & strNew & "' in stored program data and settings (CEI, Data Dictionary, Marked tasks, Metrics, etc.)?" & vbCrLf & vbCrLf & "Hit Cancel to ignore all updates.", vbQuestion + vbYesNoCancel, "Confirm Overwrite")
    Select Case lngResponse
      Case vbYes
        Set oSettingsFiles = CreateObject("Scripting.Dictionary")
        oSettingsFiles.Add "\settings\cpt-cei.adtg", "PROJECT"
        oSettingsFiles.Add "\settings\cpt-data-dictionary.adtg", "PROJECT_NAME"
        oSettingsFiles.Add "\settings\cpt-export-resource-userfields.adtg", "PROGRAM"
        cptMigrateSaveMarked
        oSettingsFiles.Add "\settings\cpt-marked.adtg", "PROJECT_ID"
        oSettingsFiles.Add "\settings\cpt-metrics.adtg", "PROGRAM"
        oSettingsFiles.Add "\settings\cpt-qbd.adtg", "PROGRAM"
        'oSettingsFiles.Add "\settings\cpt-sanitize.adtg", "PROGRAM"
        oSettingsFiles.Add "\settings\cpt-save-local.adtg", "PROGRAM_ACRONYM"
        oSettingsFiles.Add "\settings\cpt-status-sheet-userfields.adtg", "PROGRAM"
        
        'update all data sources
        For lngItem = 0 To oSettingsFiles.Count - 1
          strFileName = strDir & oSettingsFiles.Keys(lngItem)
          strFieldName = oSettingsFiles.Items(lngItem)
          If Dir(strFileName) <> vbNullString Then
            lngUpdated = 0
            Set oRecordset = CreateObject("ADODB.Recordset")
            With oRecordset
              .Open strFileName, , adOpenKeyset, adLockOptimistic
              .MoveFirst
              Do While Not .EOF
                If .Fields(strFieldName) = strOld Then
                  .Fields(strFieldName) = strNew
                  lngUpdated = lngUpdated + 1
                End If
                .MoveNext
              Loop
              .Save strFileName, adPersistADTG
              .Close
            End With
            strUpdated = strUpdated & Dir(strFileName) & ": " & Format(lngUpdated, "#,##0") & " record(s) updated." & vbCrLf
          Else
            MsgBox strFileName, vbExclamation + vbOKOnly, "File Not Found:"
          End If
        Next lngItem
        
        '\settings\cpt-settings-[PROGRAM].ini
        strOld = strDir & "\settings\cpt-settings-" & strOld & ".ini"
        strNew = strDir & "\settings\cpt-settings-" & strNew & ".ini"
        If Dir(strOld) <> vbNullString Then
          Name strOld As strNew
        Else
          MsgBox "Could not find:" & strOld & vbCrLf & vbCrLf & "Rename your settings file manually to:" & vbCrLf & strNew, vbCritical + vbOKOnly, "File Not Found"
        End If
        'todo: lblDir.Caption
        MsgBox strUpdated, vbInformation + vbOKOnly, "Data Files Updated"
      Case vbNo
        If MsgBox("CEI, Data Dictionary, Marked tasks, and Metrics entries associated with '" & strOld & "' will be disconnected from this project file. Are you sure you wish to proceed?", vbQuestion + vbYesNo) = vbNo Then
          oDocProp.Value = strOld
          Me.txtProgramAcronym.Value = strOld
        End If
      Case vbCancel
        oDocProp.Value = strOld
        Me.txtProgramAcronym.Value = strOld
    End Select
  End If
  
exit_here:
  On Error Resume Next
  If oRecordset.State = 1 Then oRecordset.Close
  Set oRecordset = Nothing
  Set oDocProp = Nothing
  Set oDocProps = Nothing
  Set oSettingsFiles = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cmdSetProgramAcronym_Click", Err, Erl)
  Resume exit_here
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
  Call cptHandleErr(THIS_MODULE, "lblURL_Click", Err, Erl)
  Resume exit_here
  
End Sub

Sub lboFeatures_AfterUpdate()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFeature As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboSettings.Clear
  If IsNull(Me.lboFeatures.Value) Then GoTo exit_here
  If Len(Me.lboFeatures.Value) = 0 Then GoTo exit_here
  strFeature = Me.lboFeatures.Value
  strFeature = Left(strFeature, Len(strFeature) - 4)
  Me.lblDir.Caption = "%USERPROFILE%\cpt-backup\settings\" & cptGetSettingsFile(strFeature)
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open cptDir & "\settings\cpt-settings.adtg"
    .Sort = "Setting"
    .MoveFirst
    Do While Not .EOF
      If .Fields(0) = strFeature And .Fields(1) <> "" Then
        Me.lboSettings.AddItem .Fields(1)
      End If
      .MoveNext
    Loop
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "lboFeatures_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Sub tglErrorTrapping_Click()
  If Me.tglErrorTrapping Then
    Me.tglErrorTrapping.Caption = "OFF"
    Me.tglErrorTrapping.BackColor = 192 'red
    If Me.Visible Then
      cptSaveSetting "General", "ErrorTrapping", "0"
      cptUpdateSetting "General", "ErrorTrapping", "0"
    End If
  Else
    Me.tglErrorTrapping.Caption = "ON"
    Me.tglErrorTrapping.BackColor = 49152 'green
    If Me.Visible Then
      cptSaveSetting "General", "ErrorTrapping", "1"
      cptUpdateSetting "General", "ErrorTrapping", "1"
    End If
  End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub UserForm_Terminate()
  Dim strFileName  As String
  strFileName = cptDir & "\settings\cpt-settings.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  Unload Me
End Sub

Sub cptUpdateSetting(strFeature As String, strKey As String, strVal As String)
  'objects
  Dim rst As ADODB.Recordset
  'strings
  Dim strFeatures As String
  Dim strFileName  As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vFeatures As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = cptDir & "\settings\cpt-settings.adtg"
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    .Open strFileName
    .MoveFirst
    .Find "Setting like '" & strKey & "%'"
    If Not .EOF Then
      .Fields(1) = strKey & "=" & strVal
    Else
      .AddNew Array(0, 1), Array(strFeature, strKey & "=" & strVal)
      .MoveFirst
      .Sort = "Feature,Setting"
      'add it alphabetized
      For lngItem = 0 To Me.lboFeatures.ListCount - 1
        strFeatures = strFeatures & Me.lboFeatures.List(lngItem) & ","
      Next lngItem
      strFeatures = strFeatures & "General"
      vFeatures = Split(strFeatures, ",")
      cptQuickSort vFeatures, 0, UBound(vFeatures)
      Me.lboFeatures.List = vFeatures
      Me.lboFeatures.Value = "General"
    End If
    .Save strFileName, adPersistADTG
    .Close
  End With
  
  Me.lboFeatures_AfterUpdate
  
exit_here:
  On Error Resume Next
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "cptUpdateSettings", Err, Erl)
  Resume exit_here
End Sub
