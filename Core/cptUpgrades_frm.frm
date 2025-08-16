VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptUpgrades_frm 
   Caption         =   "Installation Status"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
   OleObjectBlob   =   "cptUpgrades_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptUpgrades_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.5.5</cpt_version>
'/===== IMPORTANT =====\
'ALL CODE IN THIS MODULE MUST BE SELF-CONTAINED
'DO NOT BUMP THE VERSION WHILE DEVELOPING OR IT WILL GET OVERWRITTEN NEXT TIME YOU OPEN THE FORM
'\===== IMPORTANT =====/
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboBranches_Change()
  'objects
  Dim FindRecord As Object
  Dim vbComponent As Object
  Dim rstStatus As Object 'ADODB.Recordset
  Dim xmlNode As Object
  Dim xmlDoc As Object
  'strings
  Dim strInstVer As String
  Dim strCurVer As String
  Dim strVersion As String
  Dim strURL As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnUpdatesAreAvailable As Boolean
  'variants
  Dim vCol As Variant
  'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not Me.Visible Then Exit Sub
  'set up the recordset
  Set rstStatus = CreateObject("ADODB.Recordset")
  rstStatus.Fields.Append "Module", 200, 200
  rstStatus.Fields.Append "Directory", 200, 200
  rstStatus.Fields.Append "Current", 200, 200
  rstStatus.Fields.Append "Installed", 200, 200
  rstStatus.Fields.Append "Status", 200, 200
  rstStatus.Open
  
  'get current versions
  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.SetProperty "SelectionLanguage", "XPath"
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
  strURL = strGitHub & "CurrentVersions.xml"
  If Me.cboBranches.Value <> "master" Then
    strURL = Replace(strURL, "master", Me.cboBranches.Value)
  End If
  If Not xmlDoc.Load(strURL) Then
    MsgBox xmlDoc.parseError.errorcode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
    GoTo exit_here
  Else
    For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
      rstStatus.AddNew
      rstStatus(0) = xmlNode.SelectSingleNode("Name").Text
      rstStatus(1) = xmlNode.SelectSingleNode("Directory").Text
      rstStatus(2) = xmlNode.SelectSingleNode("Version").Text
      rstStatus.Update
    Next
  End If

  'get installed versions
  blnUpdatesAreAvailable = False
  For Each vbComponent In ThisProject.VBProject.VBComponents
    'is the vbComponent one of ours?
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = cptRegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
      rstStatus.MoveFirst
      rstStatus.Find "Module='" & vbComponent.Name & "'", , 1
      If Not rstStatus.EOF Then
        rstStatus(3) = strVersion
        rstStatus(4) = cptVersionStatus(rstStatus(2), strVersion)
        rstStatus.Update
      End If
    End If
  Next vbComponent
  Set vbComponent = Nothing

  'populate the listbox header
  lngItem = 0
  Me.lboHeader.AddItem
  For Each vCol In Array("Module", "Directory", "Current", "Installed", "Status", "Type")
    Me.lboHeader.List(0, lngItem) = vCol
    lngItem = lngItem + 1
  Next vCol
  Me.lboHeader.Height = 16

  'populate the listbox
  Me.lboModules.Clear
  rstStatus.Sort = "Module"
  rstStatus.MoveFirst
  lngItem = 0
  Do While Not rstStatus.EOF
    strCurVer = rstStatus(2)
    If Not IsNull(rstStatus(3)) Then
      strInstVer = rstStatus(3)
    Else
      strInstVer = "< missing >"
    End If
    Me.lboModules.AddItem
    Me.lboModules.List(lngItem, 0) = rstStatus(0) 'module name
    Me.lboModules.List(lngItem, 1) = rstStatus(1) 'directory
    Me.lboModules.List(lngItem, 2) = strCurVer 'current version
    Me.lboModules.List(lngItem, 3) = strInstVer 'installed version
    
    Select Case strInstVer
      Case Is = strCurVer
        Me.lboModules.List(lngItem, 4) = "< ok >"
      Case Is = "< missing >"
        Me.lboModules.List(lngItem, 4) = "< install >"
      Case Is <> strCurVer
        Me.lboModules.List(lngItem, 4) = "< " & cptVersionStatus(strInstVer, strCurVer) & " >"
    End Select
    'capture the type while we're at it - could have just pulled the FileName
    Set FindRecord = xmlDoc.SelectSingleNode("//Name[text()='" + Me.lboModules.List(lngItem, 0) + "']").ParentNode.SelectSingleNode("Type")
    Me.lboModules.List(lngItem, 5) = FindRecord.Text
next_lngItem:
    lngItem = lngItem + 1
    rstStatus.MoveNext
  Loop

exit_here:
  On Error Resume Next
  If rstStatus.State Then rstStatus.Close
  Set rstStatus = Nothing
  Set FindRecord = Nothing
  Set vbComponent = Nothing
  Set xmlNode = Nothing
  Set xmlDoc = Nothing

  Exit Sub
err_here:
  Call cptHandleErrUpgrade("cptUpgrades_frm", "cboBranches_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdUpgradeAll_Click()
  Dim lngItem As Long

  For lngItem = 0 To Me.lboModules.ListCount - 1
    If Me.lboModules.List(lngItem, 2) <> Me.lboModules.List(lngItem, 3) Then
      Me.lboModules.Selected(lngItem) = True
    End If
  Next lngItem
  
  Call cmdUpgradeSelected_Click
  
End Sub

Private Sub cmdUpgradeSelected_Click()
  'do not call out of this module
  'objects
  Dim rstCode As Object 'ADODB.Recordset
  Dim cmCptThisProject As Object 'VBCodeModule
  Dim cmThisProject As Object 'VBCodeModule
  Dim Project As Object
  Dim vbComponent As Object 'vbComponent
  Dim xmlHttpDoc As Object
  Dim oStream As Object 'ADODB.Stream
  'strings
  Dim strFileType As String
  Dim lngEvent As String
  Dim strVersion As String
  Dim strMsg As String
  Dim strCptFileName As String
  Dim strDirectory As String
  Dim strModule As String, strFileName As String, strURL As String
  'longs
  Dim lngLine As Long
  Dim lngItem As Long
  'integers
  'booleans
  'variants
  Dim vEvent As Variant
  'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For lngItem = 0 To Me.lboModules.ListCount - 1
    'scroll to selected
    If lngItem > 1 Then
      Me.lboModules.TopIndex = lngItem - 2
    Else
      Me.lboModules.TopIndex = lngItem
    End If
    If Me.lboModules.Selected(lngItem) Then
      '<issue33> trap invalid use of null error?
      If IsNull(Me.lboModules.List(lngItem, 0)) Then
        MsgBox "Unable to download upgrades.", vbExclamation + vbOKOnly, "Can't Connect"
        GoTo exit_here
      End If '</issue33>
      
      Me.lboModules.ListIndex = lngItem
      strModule = Me.lboModules.List(lngItem, 0)
      If strModule = "cptUpgrades_frm" Then
        Me.lboModules.List(lngItem, 4) = "<skipped>"
        GoTo next_module
      Else
        Me.lboModules.List(lngItem, 4) = "<installing...>"
      End If
      Application.StatusBar = "installing " & strModule & "..."

      'get the module name
      'get the repo directory
      'get the filename
      Set xmlHttpDoc = CreateObject("Microsoft.XMLHTTP")
      strDirectory = Me.lboModules.List(lngItem, 1)
      strFileType = Me.lboModules.List(lngItem, 5)
      strFileName = strModule & Switch(strFileType = "1", ".bas", _
                                  strFileType = "2", ".cls", _
                                  strFileType = "3", ".frm", _
                                  strFileType = "100", ".cls")
      strDirectory = Me.lboModules.List(lngItem, 1)
get_frx:
      strURL = strGitHub
      If Me.cboBranches <> "master" Then
        strURL = Replace(strURL, "master", Me.cboBranches.Value)
      End If
      strURL = strURL & strDirectory & "/" & strFileName
      xmlHttpDoc.Open "GET", strURL, False
      xmlHttpDoc.Send
      
      'strURL = xmlHttpDoc.responseBody
      If xmlHttpDoc.Status = 200 And xmlHttpDoc.readyState = 4 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1 'adTypeBinary
        oStream.Write xmlHttpDoc.responseBody
        If Dir(cptDirUpgrade & "\" & strFileName) <> vbNullString Then Kill cptDirUpgrade & "\" & strFileName
        oStream.SaveToFile cptDirUpgrade & "\" & strFileName
        oStream.Close
      Else
        MsgBox "Download failed. Please contact help@ClearPlanConsulting.com for help.", vbCritical + vbOKOnly, "XML Error"
        Me.lboModules.List(lngItem, 3) = "<failed>"
        GoTo exit_here
      End If
      If Right(strFileName, 4) = ".frm" Then
        strFileName = Replace(strFileName, ".frm", ".frx")
        GoTo get_frx
      ElseIf Right(strFileName, 4) = ".frx" Then
        strFileName = Replace(strFileName, ".frx", ".frm")
      End If

      '<issue15> added
      If strModule = "ThisProject" Then GoTo next_module 'handle separately </issue25>

      If cptModuleExistsUpgrade(strModule) Then
        '<issue19>
        Set vbComponent = ThisProject.VBProject.VBComponents(strModule)
        Dim lngCounter As Long
        lngCounter = lngCounter + 1
        Dim strComponentName As String
        strComponentName = "remove" & lngCounter & Format(Now, "hhnnss")
        vbComponent.Name = strComponentName
        DoEvents
        ThisProject.VBProject.VBComponents.Remove vbComponent 'ThisProject.VBProject.VBComponents(strModule)
        DoEvents '</issue19>
      End If
      ThisProject.VBProject.VBComponents.Import cptDirUpgrade & "\" & strFileName
      
      '<issue24> remove the whitespace added by VBE import/export
      With ThisProject.VBProject.VBComponents(strModule).CodeModule
        For lngLine = .CountOfDeclarationLines To 1 Step -1
          If Len(.Lines(lngLine, 1)) = 0 Then .DeleteLines lngLine, 1
        Next lngLine
      End With '</issue24>
      
      Me.lboModules.List(lngItem, 3) = Me.lboModules.List(lngItem, 2)
      Me.lboModules.List(lngItem, 4) = "<updated>"
    End If
next_module:     '</issue25>
  Next lngItem

  '<issue25> added
  'update ThisProject
  strFileName = cptDirUpgrade & "\ThisProject.cls"
  If Dir(strFileName) <> vbNullString Then 'the file was downloaded, proceed

    'notify user that modifications are about to be made to the ThisProject module
    strMsg = "This upgrade requires a revision to your ThisProject module. "
    strMsg = strMsg & "If you have made modifications, your code will not be lost, but it may need to be rearranged." & vbCrLf & vbCrLf
    strMsg = strMsg & "Please contact help@ClearPlanConsulting.com if you require assistance."
    MsgBox strMsg, vbInformation + vbOKOnly, "Notice"
    'ideally this would prompt user to proceed or rollback...

    'clear out existing lines of cpt-related code
    Set cmThisProject = ThisProject.VBProject.VBComponents("ThisProject").CodeModule
    For lngLine = cmThisProject.CountOfLines To 1 Step -1
      'cover both '</cpt_version> and '</cpt>
      If InStr(cmThisProject.Lines(lngLine, 1), "</cpt") > 0 Then
        cmThisProject.DeleteLines lngLine, 1
        DoEvents
      End If
    Next lngLine

    'rename file and import it
    strCptFileName = Replace(strFileName, "ThisProject", "cptThisProject")
    Name strFileName As strCptFileName
    Set cmCptThisProject = ThisProject.VBProject.VBComponents.Import(strCptFileName).CodeModule
    'grab and insert the updated version
    strVersion = cptRegEx(cmCptThisProject.Lines(1, cmCptThisProject.CountOfLines), "<cpt_version>.*</cpt_version>")
    cmThisProject.InsertLines 1, "'" & strVersion

    'grab the imported code
    Set rstCode = CreateObject("ADODB.Recordset")
    rstCode.Fields.Append "Event", 200, 120
    rstCode.Fields.Append "SLOC", 203, 5000
    rstCode.Open
    With cmCptThisProject
      For Each vEvent In Array("Project_Activate", "Project_Open")
        rstCode.AddNew
        rstCode(0) = CStr(vEvent)
        rstCode(1) = .Lines(.ProcStartLine(CStr(vEvent), 0) + 2, .ProcCountLines(CStr(vEvent), 0) - 3) '0=vbext_pl_Proc
        rstCode.Update
      Next vEvent
    End With
    ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents(cmCptThisProject.Parent.Name)
    '<issue19> added
    DoEvents '</issue19>

    'add the events, or insert new text
    'three cases: empty or not (code exists or not)
    For Each vEvent In Array("Project_Activate", "Project_Open")
      
      'find the record
      rstCode.MoveFirst
      rstCode.Find "Event='" & CStr(vEvent) & "'", , 1
      
      'if event exists then insert code else create new event handler
      With cmThisProject
        If .CountOfLines > .CountOfDeclarationLines Then 'complications
          If .Find("Sub " & CStr(vEvent), 1, 1, .CountOfLines, 1000) = True Then
          'find its line number
            lngEvent = .ProcBodyLine(CStr(vEvent), 0)  '= vbext_pk_Proc
            'import them if they *as a group* don't exist
            If .Find(rstCode(1), .ProcStartLine(CStr(vEvent), 0), 1, .ProcCountLines(CStr(vEvent), 0), 1000) = False Then 'vbext_pk_Proc
              .InsertLines lngEvent + 1, rstCode(1)
            Else
              'Debug.Print CStr(vEvent) & " code exists."
            End If
          Else 'create it
            'create it, returning its line number
            lngEvent = .CreateEventProc(Replace(CStr(vEvent), "Project_", ""), "Project")
            'insert cpt code after line number
            .InsertLines lngEvent + 1, rstCode(1)
          End If
        Else 'easy
          'create it, returning its line number
          lngEvent = .CreateEventProc(Replace(CStr(vEvent), "Project_", ""), "Project")
          'insert cpt code after line number
          .InsertLines lngEvent + 1, rstCode(1)
        End If 'lines exist
      End With 'thisproject.codemodule

      'leave no trace
      If Dir(strCptFileName) <> vbNullString Then Kill strCptFileName

    Next vEvent
  End If '</issue25>

  'reset the ribbon
  strMsg = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
  strMsg = strMsg + "<mso:customUI "
  strMsg = strMsg + "xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >"
  strMsg = strMsg + vbCrLf & "<mso:ribbon startFromScratch=""false"" >"
  strMsg = strMsg + vbCrLf & "<mso:tabs>"
  strMsg = strMsg + cptBuildRibbonTab()
  'todo: need to get any additions made also...
  strMsg = strMsg + vbCrLf & "</mso:tabs>"
  strMsg = strMsg + vbCrLf & "</mso:ribbon>"
  strMsg = strMsg + vbCrLf & "</mso:customUI>"
  ActiveProject.SetCustomUI (strMsg)
  
exit_here:
  On Error Resume Next
  If rstCode.State Then rstCode.Close
  Set rstCode = Nothing
  Set cmCptThisProject = Nothing
  Set cmThisProject = Nothing
  Application.ScreenUpdating = True
  Set Project = Nothing
  Set vbComponent = Nothing
  Application.StatusBar = ""
  Set xmlHttpDoc = Nothing
  Set oStream = Nothing
  Exit Sub
err_here:
  Call cptHandleErrUpgrade("cptUpgrades_frm", "cmdUpdate_Click", Err, Erl)
  Me.lboModules.List(lngItem, 3) = "<error>" '</issue25>
  Resume exit_here

End Sub

Private Sub lblTitle_Click()
  Me.txtDevMode.Value = Val(Me.txtDevMode.Value) + 1
  If Val(Me.txtDevMode) > 5 Then
    Me.txtDevMode.Value = 0
    Me.cboBranches.Visible = False
  ElseIf Val(Me.txtDevMode.Value) = 5 Then
    Me.cboBranches.Visible = True
  Else
    Me.cboBranches.Visible = False
  End If
End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErrUpgrade("cptUpgrades_frm", "lblURL_Click", Err, Erl)
  Resume exit_here
End Sub

Private Function cptDirUpgrade() As String
Dim strPath As String

  'confirm existence of cpt settings and backup modules file

  'strPath = ThisProject.FullName
  'strPath = Left(strPath, InStrRev(strPath, "MS Project\") - 1 + Len("MS Project\"))

  strPath = Environ("USERPROFILE")
  strPath = strPath & "\cpt-backup"
  If Dir(strPath, vbDirectory) = vbNullString Then
    MkDir strPath
  End If
  If Dir(strPath & "\settings", vbDirectory) = vbNullString Then
    MkDir strPath & "\settings"
  End If
  If Dir(strPath & "\modules", vbDirectory) = vbNullString Then
    MkDir strPath & "\modules"
  End If
  cptDirUpgrade = strPath

End Function

Private Sub cptHandleErrUpgrade(strModule As String, strProcedure As String, objErr As ErrObject, Optional lngErl As Long)
  'common error handling prompt
  Dim strMsg As String

  strMsg = "Please contact help@ClearPlanConsulting.com for assistance if needed." & vbCrLf & vbCrLf
  strMsg = strMsg & "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf
  strMsg = strMsg & "Source: " & strModule & "." & strProcedure
  If lngErl > 0 Then
    strMsg = strMsg & ":" & lngErl
  End If
  MsgBox strMsg, vbExclamation + vbOKOnly, "Error"
  Dim strFileName As String
  Dim lngFile As Long
  strFileName = Environ("tmp") & "\cptUpgradeError.txt"
  Open strFileName For Output As #lngFile
  Print #lngFile, "Please send the following text to help@ClearPlanConsulting.com:"
  Print #lngFile, "Error: " & Err.Number & ": " & Err.Description
  Print #lngFile, "Source: " & strModule & "." & strProcedure
  If lngErl > 0 Then
    Print #lngFile, "Line: " & lngErl
  End If
  Close #lngFile
  ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1

End Sub

Private Function cptModuleExistsUpgrade(strModule As String)
  'objects
  Dim vbComponent As Object
  'booleans
  Dim blnExists As Boolean
  'strings
  Dim strError As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  blnExists = False
  For Each vbComponent In ThisProject.VBProject.VBComponents
    If UCase(vbComponent.Name) = UCase(strModule) Then
      blnExists = True
      Exit For
    End If
  Next vbComponent

  cptModuleExistsUpgrade = blnExists

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErrUpgrade("cptUpgrades_frm", "cptModuleExistsUpgrade", Err, Erl)
  Resume exit_here

End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub
