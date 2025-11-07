VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptDECM_frm 
   Caption         =   "DECM v7.0"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760.001
   OleObjectBlob   =   "cptDECM_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptDECM_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v8.1.1</cpt_version>
Option Explicit

Private Sub chkUpdateView_Click()
  Dim blnUpdateView As Boolean
  If Not Me.Visible Then Exit Sub
  blnUpdateView = Me.chkUpdateView
  cptSaveSetting "Integration", "chkUpdateView", IIf(blnUpdateView, "1", "0")
  If blnUpdateView Then
    If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 8)) Then
      cptDECM_UPDATE_VIEW Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0), Me.lboMetrics.List(Me.lboMetrics.ListIndex, 8)
    Else
      cptDECM_UPDATE_VIEW Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0), ""
    End If
  End If
End Sub

Private Sub cmdDone_Click()
  Dim oExcel As Excel.Application
  Dim vFile As Variant
  Dim strFileName As String
  Dim vGroup As Variant
  Dim strGroups As String
  Dim blnErrorTrapping As Boolean
  
  blnErrorTrapping = cptErrorTrapping
  Me.Hide
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  'then clean up after yourself
  For Each vFile In Split("Schema.ini,tasks.csv,targets.csv,assignments.csv,links.csv,wp-ims.csv,wp-ev.csv,wp-not-in-ims.csv,wp-not-in-ev.csv,10A302b-x.csv,decm-cpt01.adtg,10A303a-x.csv,fiscal.csv,cpt-cei.csv,06A506c-x.csv,06A504a.csv,06A504b.csv,segregated.csv,itemized.csv,06A101a.xlsx,06A212a.xlsm,10A102a.xlsx,10A103a.xlsx", ",")
    strFileName = Environ("tmp") & "\" & vFile
    If Not oExcel Is Nothing Then
      On Error Resume Next
      oExcel.Windows(vFile).Close False
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    End If
    If Dir(strFileName) <> vbNullString Then Kill strFileName
  Next vFile
  cptResetAll

  'git grep 'strGroup =' | grep -v "grep" | awk -F"strGroup = " '{ print $2}' | sed 's/"//g' | tr -s '\n' ','
  strGroups = "cpt 05A101a 1 CA : 1 OBS,cpt 05A102a 1 CA : 1 CAM,cpt 05A103a 1 CA : 1 WBS,cpt 1wp_1ca,cpt 10A102a 1 WP : 1 EVT,cpt 11A101a CA BAC = SUM(WP BAC),cpt 06A210a LOE driving Discrete"
  For Each vGroup In Split(strGroups, ",")
    If cptGroupExists(CStr(vGroup)) Then ActiveProject.TaskGroups2(vGroup).Delete
  Next vGroup
  
exit_here:
  On Error Resume Next
  Set oExcel = Nothing
  Set oDECM = Nothing
  Set oSubMap = Nothing
  Exit Sub
err_here:
  cptHandleErr "cptDECM_frm", "cmdDone_Click", Err, Erl
  Resume exit_here
  
  
End Sub

Private Sub cmdExport_Click()
  cptDECM_EXPORT Me
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDECM_frm", "lblURL_Click", Err, Erl)
  Resume exit_here

End Sub

Public Sub lboMetrics_AfterUpdate()
  'objects
  Dim oFile As Scripting.TextStream  'Object
  Dim oFSO As Scripting.FileSystemObject  'Object
  'strings
  Dim strFileName As String
  Dim strRollingWaveDate As String
  Dim strMsg As String
  Dim strDir As String
  Dim strDescription As String
  Dim strMetric As String
  Dim strTitle As String
  Dim strTarget As String
  Dim strScore As String
  'longs
  Dim lngX As Long
  Dim lngY As Long
  Dim lngItem As Long
  Dim lngResponse As Long
  'integers
  'doubles
  Dim dblScore As Double
  'booleans
  Dim blnUpdateView As Boolean
  'variants
  Dim vOOS As Variant
  'dates
  
  Me.lboOOS.Visible = False
  Me.txtTitle.Height = 128.25 + Me.lboHeader.Height + 1
  If Me.lboMetrics.ListIndex = -1 Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strMetric = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0)
  strTitle = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 1)
  strTarget = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 2)
  If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 3)) Then
    lngX = CLng(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 3))
  Else
    lngX = 0
  End If
  If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 4)) Then 'todo: not all metrics have Y
    lngY = CLng(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 4))
  Else
    lngY = 0
  End If
  If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)) Then
    strScore = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)
  Else
    strScore = "-"
  End If
  strDescription = strMetric & vbCrLf
  strDescription = strDescription & strTitle & vbCrLf & vbCrLf
  strDescription = strDescription & "TARGET: " & strTarget & vbCrLf
  strDescription = strDescription & "X: " & lngX & vbCrLf
  strDescription = strDescription & "Y: " & lngY & vbCrLf
  
  Select Case strMetric
    Case "06A101a"
      If Dir(Environ("tmp") & "\wp-ev.csv") = vbNullString Then
        strDescription = "needed: wp-ims.csv [+]" & vbCrLf
        strDescription = strDescription & "needed: wp-ev.csv  <?>" & vbCrLf
        Me.txtTitle.Value = strDescription
        lngResponse = MsgBox("Has the EV Analyst sent you the list of discrete, incomplete WPs in the EV Tool?", vbQuestion + vbYesNoCancel, "06A101a - WP Mismatches")
        If lngResponse = vbNo Then
          MsgBox "Please send the following query to your EV Analyst...", vbOKOnly + vbInformation, "Data Needed"
          Set oFSO = CreateObject("Scripting.FileSystemObject")
          strDir = Environ("tmp")
          Set oFile = oFSO.CreateTextFile(strDir & "\wp-ev-sql.txt", True)
          strMsg = "Hi [name]," & vbCrLf & vbCrLf
          strMsg = strMsg & "I'm running DECM metric 06A101a which compares the list of discrete, incomplete WPs in the IMS vs what's in the EV Tool. " & vbCrLf
          strMsg = strMsg & "Could you please provide the list of discrete, incomplete WPs currently in the EV Tool?" & vbCrLf & vbCrLf
          strMsg = strMsg & "An example query for COBRA would be:" & vbCrLf
          strMsg = strMsg & String(25, "-") & vbCrLf
          
          strMsg = strMsg & "DECLARE @MyProj VARCHAR(MAX) " & vbCrLf
          strMsg = strMsg & "SET @MyProj=inputbox('Project Name:') " & vbCrLf
          strMsg = strMsg & "SELECT DISTINCT WP " & vbCrLf
          strMsg = strMsg & "FROM CAWP " & vbCrLf
          strMsg = strMsg & "WHERE PROGRAM=@MyProj " & vbCrLf
          strMsg = strMsg & "AND WP<>'' " & vbCrLf
          strMsg = strMsg & "AND PMT NOT IN ('A','J','M') " & vbCrLf
          strMsg = strMsg & "AND BCWP<(BAC-100) " & vbCrLf
          
          strMsg = strMsg & String(25, "-") & vbCrLf & vbCrLf
          strMsg = strMsg & "I appreciate your assistance. Please let me know if you have any questions."
          oFile.Write strMsg
          oFile.Close
          strFileName = strDir & "\wp-ev-sql.txt"
          ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
          GoTo exit_here
        ElseIf lngResponse = vbYes Then
          Me.txtTitle.Value = Me.txtTitle.Text & vbCrLf & "please paste data here (w/o headers):" & vbCrLf
          Me.txtTitle.SetFocus
          Me.txtTitle.SelStart = 0
          Me.txtTitle.CurLine = Me.txtTitle.LineCount - 2
          Me.txtTitle.SelLength = 65535
          Me.txtTitle.CurLine = Me.txtTitle.LineCount - 3
          Me.txtTitle.SelLength = 65535
          
          GoTo exit_here
        ElseIf lngResponse = vbCancel Then
          GoTo exit_here
        End If
      Else
        strDescription = strDescription & "SCORE: " & strScore
        strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
      End If
    Case "06I201a"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "Task Name contains 'SVT' and has resource assignments"
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "06A205a"
      strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "NOTE: metric does not address leads (negative lags)."
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "06A208a"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "06A210a"
      strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "NOTE: filter shows both LOE pred and Non-LOE successor."
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "06A212a"
      strDescription = strMetric & vbCrLf
      strDescription = strDescription & strTitle & vbCrLf & vbCrLf
      strDescription = strDescription & "TARGET: " & strTarget & vbCrLf
      strDescription = strDescription & "X: " & lngX & vbCrLf & vbCrLf
      strDescription = strDescription & Me.lboMetrics.List(Me.lboMetrics.ListIndex, 7)
      If Len(oDECM(strMetric)) > 0 Then
        Me.txtTitle.Height = (Me.lboMetrics.Height / 2) + 5
        Me.lboOOS.Top = Me.txtTitle.Top + Me.txtTitle.Height + 1
        Me.lboOOS.Height = (Me.lboMetrics.Height / 2) + Me.lboHeader.Height
        Me.lboOOS.Width = Me.txtTitle.Width
        Me.lboOOS.ColumnWidths = "15;65;65"
        Me.lboOOS.Clear
        Me.lboOOS.Visible = True
        Me.Repaint
        Me.lboOOS.AddItem
        Me.lboOOS.List(Me.lboOOS.ListCount - 1, 0) = "#"
        Me.lboOOS.List(Me.lboOOS.ListCount - 1, 1) = "FROM"
        Me.lboOOS.List(Me.lboOOS.ListCount - 1, 2) = "TO"
        lngItem = 0
        For Each vOOS In Split(oDECM(strMetric), ";")
          If Len(vOOS) > 0 Then
            Me.lboOOS.AddItem
            Me.lboOOS.List(Me.lboOOS.ListCount - 1, 0) = lngItem + 1
            Me.lboOOS.List(Me.lboOOS.ListCount - 1, 1) = Split(vOOS, ",")(0)
            Me.lboOOS.List(Me.lboOOS.ListCount - 1, 2) = Split(vOOS, ",")(1)
            lngItem = lngItem + 1
          End If
        Next vOOS
        strDescription = strDescription & vbCrLf & vbCrLf & "...double-click to open 06A211a.xlsm."
        strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
      End If
    Case "06A401a" 'critical path
      strDescription = strMetric & vbCrLf
      strDescription = strDescription & strTitle & vbCrLf & vbCrLf
      strDescription = strDescription & "TARGET: " & strTarget & vbCrLf
      strDescription = strDescription & "X: " & lngX & vbCrLf
      strDescription = strDescription & "SCORE: " & lngX & vbCrLf & vbCrLf
      strDescription = strDescription & "UID Targeted: " & Split(oDECM(strMetric), "|")(0) & vbCrLf
      strDescription = strDescription & "NOTE: subtract # of tasks that *are* on the this schedule's critical path."
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "06A504a"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "...requires CPT > Status > Capture Week, two periods"
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "06A504b"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "...requires CPT > Status > Capture Week, two periods"
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "06A506b"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "06A506c"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "...requires CPT > Status > Capture Week, two periods"
     strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "06A212a"
      strDescription = strDescription & vbCrLf & "...pairs exported to Excel" & vbCrLf & "...select to filter"
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "10A103a"
      strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
      If lngX > 0 Then strDescription = strDescription & vbCrLf & vbCrLf & "...details exported to Excel."
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "11A101a"
      strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "NOTE: analysis done on Baseline Work only."
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "29A601a"
      strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
      strRollingWaveDate = cptGetSetting("Integration", "RollingWaveDate")
      If Len(strRollingWaveDate) > 0 Then
        strDescription = strDescription & vbCrLf & vbCrLf & "Rolling Wave Date: " & FormatDateTime(CDate(strRollingWaveDate), vbShortDate)
      End If
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "CPT01"
      strDescription = strDescription & "SCORE: " & lngX
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case "CPT02"
      strDescription = strDescription & "SCORE: " & lngX
      strDescription = strDescription & vbCrLf & vbCrLf & cptGetDECMDescription(strMetric)
    Case Else
      strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore & vbCrLf & vbCrLf
      strDescription = strDescription & cptGetDECMDescription(strMetric)
  End Select
  
  Me.txtTitle.Value = strDescription
  blnUpdateView = Me.chkUpdateView
  If blnUpdateView Then
    If Len(oDECM(strMetric)) > 0 Then
      If strMetric = "06A212a" Then
        cptDECM_UPDATE_VIEW strMetric, Replace(Replace(Left(oDECM(strMetric), Len(oDECM(strMetric)) - 1), ",", vbTab), ";", vbTab)
      ElseIf strMetric = "06A401a" Then
        cptDECM_UPDATE_VIEW strMetric, CStr(Split(oDECM(strMetric), "|")(1))
      Else
        cptDECM_UPDATE_VIEW strMetric, oDECM(strMetric)
      End If
    Else
      cptDECM_UPDATE_VIEW Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0), oDECM(strMetric) 'todo: huh?
    End If
    AppActivate Me.Caption, False
  End If
  
exit_here:
  On Error Resume Next
  Set oFile = Nothing
  Set oFSO = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDECM_frm", "lboMetrics_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboMetrics_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim strMetric As String
  Dim strFileName As String
  If Not IsNull(Me.lboMetrics.Value) Then
    strMetric = Me.lboMetrics.Value
    Select Case strMetric
      Case "06A101a"
        strFileName = Environ("tmp") & "\" & strMetric & ".xlsx"
        If Dir(strFileName) <> vbNullString Then
          If MsgBox("Open " & strMetric & ".xslx?", vbQuestion + vbYesNo, strMetric & " Details") = vbYes Then
            Shell "excel.exe """ & strFileName & """", vbMaximizedFocus
          End If
        End If
      Case Else
        strFileName = Environ("tmp") & "\" & strMetric & ".xlsm"
        If Dir(strFileName) <> vbNullString Then
          If MsgBox("Open " & strMetric & ".xslm?", vbQuestion + vbYesNo, strMetric & " Details") = vbYes Then
            Shell "excel.exe """ & strFileName & """", vbMaximizedFocus
          End If
        End If
        
    End Select
    
  End If
End Sub

Private Sub lboOOS_AfterUpdate()
  If Me.lboOOS.ListIndex = 0 Then
    Me.lboOOS.ListIndex = 1
    Me.lboOOS.Selected(0) = False
    Me.lboOOS.Selected(1) = True
    Exit Sub
  End If
End Sub

Private Sub lboOOS_Click()
  If Me.lboOOS.ListIndex = 0 Then Exit Sub
  If Me.lboOOS.ListCount > 1 Then
    Dim strFilter As String
    ScreenUpdating = False
    ActiveWindow.TopPane.Activate
    FilterClear
    GroupClear
    Sort "ID", , , , , , False
    OptionsViewEx DisplaySummaryTasks:=True
    OutlineShowAllTasks
    OptionsViewEx DisplaySummaryTasks:=False
    strFilter = Me.lboOOS.List(Me.lboOOS.ListIndex, 1) & vbTab
    strFilter = strFilter & Me.lboOOS.List(Me.lboOOS.ListIndex, 2)
    SetAutoFilter "Unique ID", pjAutoFilterIn, "equals", strFilter
    ScreenUpdating = True
  End If
End Sub

Private Sub txtTitle_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
  'objects
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oRecordset As ADODB.Recordset
  Dim oFile As Scripting.TextStream
  Dim oFSO As Scripting.FileSystemObject
  'strings
  Dim strFileName As String
  Dim strDescription As String
  Dim strPass As String
  Dim strFail As String
  Dim strScore As String
  Dim strTarget As String
  Dim strTitle As String
  Dim strMetric As String
  Dim strDir As String
  Dim strCon As String
  Dim strSQL As String
  'longs
  Dim lngX As Long
  Dim lngY As Long
  Dim lngRecord As Long
  'integers
  'doubles
  Dim dblScore As Double
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vData As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strMetric = Me.lboMetrics.Value
  If strMetric <> "06A101a" Then Exit Sub
  
  Me.txtTitle.Text = Me.txtTitle.Text & "validating..." & vbCrLf
  vData = Split(Data.GetText, vbCrLf)
  
  strDir = Environ("tmp")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  Set oRecordset = CreateObject("ADODB.Recordset")

  'populate lboFilter
  If UBound(vData) > 1 Then 'user pasted a column of data
    Cancel = True 'cancel paste operation
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFSO.CreateTextFile(strDir & "\wp-ev.csv", True)
    oFile.Write Join(vData, vbCrLf)
    oFile.Close
    'ensure distinct
    strSQL = "SELECT DISTINCT WP FROM [wp-ev.csv]"
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    lngY = oRecordset.RecordCount
    Me.lboMetrics.List(Me.lboMetrics.ListCount - 1, 4) = lngY
    Set oFile = oFSO.CreateTextFile(strDir & "\wp-ev-distinct.csv", True)
    oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
    oFile.Close
    oRecordset.Close
    Kill strDir & "\wp-ev.csv"
    Name strDir & "\wp-ev-distinct.csv" As strDir & "\wp-ev.csv"
  Else
    'user pasted nothing
  End If
  
  Me.txtTitle.Text = Replace(Me.txtTitle.Text, "validating...", "validating...ok")
  Me.txtTitle.Text = Me.txtTitle.Text & lngY & " WPs loaded..." & vbCrLf
  Me.txtTitle.Text = Replace(Me.txtTitle.Text, "<?>", "[+]")
  Me.txtTitle.Text = Me.txtTitle.Text & "analyzing..." & vbCrLf
  
  'get count of WPs in IMS but not in EV
  strSQL = "SELECT wp_ims.WP,wp_ev.WP "
  strSQL = strSQL & "FROM [wp-ims.csv] AS wp_ims "
  strSQL = strSQL & "LEFT JOIN [wp-ev.csv] AS wp_ev ON wp_ev.WP=wp_ims.WP "
  strSQL = strSQL & "WHERE wp_ev.WP IS NULL"
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  lngX = oRecordset.RecordCount
  Me.txtTitle.Text = Me.txtTitle.Text & lngX & " in IMS, not in EV" & vbCrLf
  Set oFile = oFSO.CreateTextFile(strDir & "\wp-not-in-ev.csv", True)
  If Not oRecordset.EOF Then
    oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
  End If
  oFile.Close
  oRecordset.Close
  
  'get count of WPs in EV but not in IMS
  strSQL = "SELECT wp_ev.WP,wp_ims.WP "
  strSQL = strSQL & "FROM [wp-ev.csv] AS wp_ev "
  strSQL = strSQL & "LEFT JOIN [wp-ims.csv] AS wp_ims ON wp_ims.WP=wp_ev.WP "
  strSQL = strSQL & "WHERE wp_ims.WP IS NULL"
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  Me.txtTitle.Text = Me.txtTitle.Text & oRecordset.RecordCount & " in EV, not in IMS" & vbCrLf
  lngX = lngX + oRecordset.RecordCount
  Set oFile = oFSO.CreateTextFile(strDir & "\wp-not-in-ims.csv", True)
  If Not oRecordset.EOF Then
    oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
  End If
  oFile.Close
  oRecordset.Close
  Me.txtTitle.Text = Me.txtTitle.Text & lngX & " total mismatches"
  MsgBox "Analysis complete.", vbInformation + vbOKOnly, strMetric
  
  'create a report
  strSQL = "SELECT DISTINCT A.WP AS [ALL],IMS.WP AS [IMS],EV.WP AS [EV] "
  strSQL = strSQL & "FROM (( "
  strSQL = strSQL & "    SELECT * FROM [wp-ims.csv] "
  strSQL = strSQL & "    UNION "
  strSQL = strSQL & "    SELECT * FROM [wp-ev.csv] "
  strSQL = strSQL & "    ) AS A "
  strSQL = strSQL & "        LEFT JOIN [wp-ims.csv] AS IMS ON IMS.WP = A.WP "
  strSQL = strSQL & "    ) "
  strSQL = strSQL & "    LEFT JOIN [wp-ev.csv] AS EV ON EV.WP = A.WP "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & "    A.WP IS NOT NULL; "
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If Not .EOF Then
      On Error Resume Next
      Set oExcel = GetObject(, "Excel.Application")
      If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oExcel Is Nothing Then
        Set oExcel = CreateObject("Excel.Application")
      End If
      oExcel.Visible = True
      Set oWorkbook = oExcel.Workbooks.Add
      Set oWorksheet = oWorkbook.Sheets(1)
      oWorksheet.Name = "06A101a"
      oWorksheet.[A2:C2] = Split("ALL,IMS,EV", ",")
      oWorksheet.[A1:D2].Font.Bold = True
      oWorksheet.[A1:D2].HorizontalAlignment = xlCenter
      'shade header
      With oWorksheet.[A2:C2].Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
      End With
      oWorksheet.[A3].CopyFromRecordset oRecordset
      oExcel.ActiveWindow.Zoom = 85
      oExcel.ActiveWindow.SplitRow = 2
      'oExcel.ActiveWindow.SplitColumn = 0
      oExcel.ActiveWindow.FreezePanes = True
      oWorksheet.Columns.AutoFit
      Dim oListObject As Excel.ListObject
      Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A2].End(xlDown), oWorksheet.[A2].End(xlToRight)), , xlYes)
      oListObject.TableStyle = ""
      oWorksheet.[A1].FormulaR1C1 = "=COUNTA(Table1[ALL])"
      lngY = oWorksheet.[A1].Value
      oWorksheet.[B1].FormulaR1C1 = "=COUNTBLANK(Table1[IMS])"
      oWorksheet.[C1].FormulaR1C1 = "=COUNTBLANK(Table1[EV])"
      lngX = oWorksheet.[B1].Value + oWorksheet.[C1].Value
      oWorksheet.[D1].FormulaR1C1 = "=SUM(RC[-2]:RC[-1])/RC[-3]"
      oWorksheet.[D1].Style = "Percent"
      cptAddBorders oWorksheet.[A1:D1]
      cptAddBorders oListObject.Range
      'conditional formatting for blanks
      With oWorksheet.Range("Table1[[IMS]:[EV]]")
        .FormatConditions.Add Type:=xlExpression, Operator:=xlEqual, Formula1:="=ISBLANK(B3)"
        With .FormatConditions(1)
          .SetFirstPriority
          .Font.Color = -16383844
          .Font.TintAndShade = 0
          .Interior.PatternColorIndex = xlAutomatic
          .Interior.Color = 13551615
          .Interior.TintAndShade = 0
          .StopIfTrue = False
        End With
      End With
      'conditional formatting for score
      With oWorksheet.[D1]
        .FormatConditions.AddIconSetCondition
        With .FormatConditions(1)
          .SetFirstPriority
          .ReverseOrder = True
          .ShowIconOnly = False
          .IconSet = oWorkbook.IconSets(xl3Symbols)
          With .IconCriteria(2)
            .Type = xlConditionValueNumber
            .Value = 0
            .Operator = 5
          End With
          With .IconCriteria(3)
            .Type = xlConditionValueNumber
            .Value = 0
            .Operator = 5
          End With
        End With
      End With
      strFileName = Environ("tmp") & "\" & strMetric & ".xlsx"
      If Dir(strFileName) <> vbNullString Then Kill strFileName
      oWorkbook.SaveAs strFileName, 51
      'oWorkbook.Close 'todo: keep open
      .Close
    End If
  End With
  
  'update the listbox
  Me.lboMetrics.List(Me.lboMetrics.ListIndex, 3) = lngX
  Me.lboMetrics.List(Me.lboMetrics.ListIndex, 4) = lngY
  dblScore = Round(lngX / lngY, 2)
  Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5) = Format(dblScore, "0%")
  strPass = "[+]"
  strFail = "<!>"
  If dblScore = 0 Then
    Me.lboMetrics.List(Me.lboMetrics.ListIndex, 6) = strPass
  Else
    Me.lboMetrics.List(Me.lboMetrics.ListIndex, 6) = strFail
  End If

  'update the description
  strMetric = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0)
  strTitle = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 1)
  strTarget = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 2)
  If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)) Then
    strScore = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)
  Else
    strScore = "-"
  End If
  strDescription = strMetric & vbCrLf
  strDescription = strDescription & strTitle & vbCrLf & vbCrLf
  strDescription = strDescription & "TARGET: " & strTarget & vbCrLf
  strDescription = strDescription & "X: " & lngX & vbCrLf
  strDescription = strDescription & "Y: " & lngY & vbCrLf
  strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
  strDescription = strDescription & vbCrLf & vbCrLf & Me.lboMetrics.List(Me.lboMetrics.ListIndex, 7)
  
  Me.txtTitle.Text = strDescription
  If oDECM.Exists("06A101a") Then 'oDECM.Remove "06A101a"
    oDECM("06A101a") = lngX & "|" & lngY
  Else
    oDECM.Add "06A101a", lngX & "|" & lngY
  End If
  'todo: user can re-paste data
    
exit_here:
  On Error Resume Next
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oRecordset = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDECM_frm", "txtTitle_BeforeDropOrPaste", Err, Erl)
  Resume exit_here
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Set oDECM = Nothing
End Sub
