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
'<cpt_version>v1.3.2</cpt_version>
Option Explicit

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
    strFileName = cptDir & "\settings\cpt-settings.ini"
    ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
  End If
End Sub

Private Sub cmdSetProgramAcronym_Click()
  'objects
  Dim oRecordset As ADODB.Recordset
  Dim oDocProp As DocumentProperty
  Dim oDocProps As DocumentProperties
  'strings
  Dim strDir As String
  Dim strUpdated As String
  Dim strFileName  As String
  Dim strOld As String, strNew As String
  'longs
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
    lngResponse = MsgBox("Also replace all instances of '" & strOld & "' with '" & strNew & "' in stored program data and settings (CEI, Data Dictionary, Marked tasks, and Metrics)?" & vbCrLf & vbCrLf & "Hit Cancel to ignore all updates.", vbQuestion + vbYesNoCancel, "Confirm Overwrite")
    Select Case lngResponse
      Case vbYes
        Set oRecordset = CreateObject("ADODB.Recordset")
        '\settings\cpt-cei.adtg
        strFileName = strDir & "\settings\cpt-cei.adtg"
        lngUpdated = 0
        If Dir(strFileName) <> vbNullString Then
          With oRecordset
            .Open strFileName
            .MoveFirst
            Do While Not .EOF
              If .Fields("PROJECT") = strOld Then
                .Fields("PROJECT") = strNew
                lngUpdated = lngUpdated + 1
              End If
              .MoveNext
            Loop
            .Save strFileName, adPersistADTG
            .Close
          End With
          strUpdated = Format(lngUpdated, "#,##0") & " record(s) updated in CEI data." & vbCrLf
        End If
        'settings\cpt-data-dictionary.adtg
        strFileName = strDir & "\settings\cpt-data-dictionary.adtg"
        lngUpdated = 0
        If Dir(strFileName) <> vbNullString Then
          With oRecordset
            .Open strFileName
            .MoveFirst
            Do While Not .EOF
              If .Fields("PROJECT_NAME") = strOld Then
                .Fields("PROJECT_NAME") = strNew
                lngUpdated = lngUpdated + 1
              End If
              .MoveNext
            Loop
            .Save strFileName, adPersistADTG
            .Close
          End With
          strUpdated = strUpdated & Format(lngUpdated, "#,##0") & " record(s) updated in Data Dictionary." & vbCrLf
        End If
        '\cpt-marked.adtg
        strFileName = strDir & "\cpt-marked.adtg"
        lngUpdated = 0
        If Dir(strFileName) <> vbNullString Then
          With oRecordset
            .Open strFileName
            .MoveFirst
            Do While Not .EOF
              If .Fields("PROJECT_ID") = strOld Then
                .Fields("PROJECT_ID") = strNew
                lngUpdated = lngUpdated + 1
              End If
              .MoveNext
            Loop
            .Save strFileName, adPersistADTG
            .Close
          End With
          strUpdated = strUpdated & Format(lngUpdated, "#,##0") & " record(s) updated in Marked tasks data." & vbCrLf
        End If
        '\settings\cpt-metrics.adtg
        strFileName = strDir & "\settings\cpt-metrics.adtg"
        lngUpdated = 0
        If Dir(strFileName) <> vbNullString Then
          With oRecordset
            .Open strFileName
            .MoveFirst
            Do While Not .EOF
              If .Fields("PROGRAM") = strOld Then
                .Fields("PROGRAM") = strNew
                lngUpdated = lngUpdated + 1
              End If
              .MoveNext
            Loop
            .Save strFileName, adPersistADTG
            .Close
          End With
          strUpdated = strUpdated & Format(lngUpdated, "#,##0") & " record(s) updated in Metrics data." & vbCrLf
        End If
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

  Exit Sub
err_here:
  Call cptHandleErr("cptSettings_frm", "cmdSetProgramAcronym_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub lblURL_Click()
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_frm", "lblURL_Click", Err, Erl)
  Resume exit_here
End Sub

Sub lboFeatures_AfterUpdate()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboSettings.Clear
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open cptDir & "\settings\cpt-settings.adtg"
    .Sort = "Setting"
    .MoveFirst
    Do While Not .EOF
      If .Fields(0) = Me.lboFeatures.Value And .Fields(1) <> "" Then
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
  Call cptHandleErr("cptSettings_frm", "lboFeatures_AfterUpdate", Err, Erl)
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
  Call cptHandleErr("cptSettings_frm", "cptUpdateSettings", Err, Erl)
  Resume exit_here
End Sub
