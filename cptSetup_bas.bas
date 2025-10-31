Attribute VB_Name = "cptSetup_bas"
'<cpt_version>v1.11.0</cpt_version>
Option Explicit
Public Const strGitHub = "https://raw.githubusercontent.com/clearplan/cpt/master/"
Private Const BLN_TRAP_ERRORS As Boolean = True 'keep this: cptErrorTrapping() lives in cptCore_bas
#If Win64 And VBA7 Then
  Private Declare PtrSafe Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As LongPtr, _
                                                                        ByVal lpszConnectionName As String, _
                                                                        ByVal dwNameLen As Integer, _
                                                                        ByVal dwReserved As LongPtr) As LongPtr
#Else
  Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, _
                                                                        ByVal lpszConnectionName As String, _
                                                                        ByVal dwNameLen As Integer, _
                                                                        ByVal dwReserved As Long) As Long
#End If
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub cptSetup()
  'setup only needs to be run once
  'objects
  Dim Project As Object
  Dim vbComponent As Object 'vbComponent
  Dim rstCode As Object 'ADODB.Recordset
  Dim cmThisProject As Object 'CodeModule
  Dim cmCptThisProject As Object 'CodeModule
  Dim oStream As Object 'ADODB.Stream
  Dim xmlHttpDoc As Object
  Dim xmlNode As Object
  Dim xmlDoc As Object
  Dim rstCore As Object 'ADODB.Recordset
  'strings
  Dim strDir As String
  Dim strMsg As String
  Dim strError As String
  Dim strCptFileName As String
  Dim strVersion As String
  Dim strFileName As String
  Dim strModule As String
  Dim strURL As String
  'longs
  Dim lngLine As Long
  Dim lngEvent As Long
  'integers
  'booleans
  Dim blnImportModule As Boolean
  Dim blnExists As Boolean
  'variants
  Dim vEvent As Variant
  'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  strDir = cptDir
  
  '<issue61> ensure proper installation
  If InStr(ThisProject.FullName, "Global") = 0 Then
    strMsg = "The CPT can only be installed in one of the following:" & vbCrLf
    strMsg = strMsg & "> Global.MPT" & vbCrLf
    strMsg = strMsg & "> Global (+ non-cached Enterprise) temporarily and for testing purposes only" & vbCrLf
    strMsg = strMsg & "> Checked-out Enterprise Global (when ready to release to Enterprise user base)" & vbCrLf & vbCrLf
    strMsg = strMsg & "(Do not install to a *.mpp file.)"
    MsgBox strMsg, vbCritical + vbOKOnly, "Faulty Installation"
    GoTo exit_here
  End If '</issue61>

  'prompt user for setup instructions
  strMsg = "NOTE: This procedure should only be run once." & vbCrLf & vbCrLf
  strMsg = strMsg & "Before proceeding:" & vbCrLf
  strMsg = strMsg & "1. Setup your Global.MPT: File > Options > Trust Center > Trust Center Settings..." & vbCrLf
  strMsg = strMsg & vbTab & "a. Macro Settings > Enable all macros" & vbCrLf
  strMsg = strMsg & vbTab & "b. Legacy Formats > Allow loading files with legacy or non-default file formats" & vbCrLf
  strMsg = strMsg & "2. Completely exit, then re-open, MS Project (this makes the settings above 'stick')" & vbCrLf
  strMsg = strMsg & "Have you completed the above steps?" & vbCrLf & vbCrLf
  strMsg = strMsg & "(Yes = Proceed; No = Cancel and Close)"
  If MsgBox(strMsg, vbQuestion + vbYesNo, "Before you proceed...") = vbNo Then GoTo exit_here

  'capture list of files to download
  'why?
  On Error Resume Next
  Set rstCore = CreateObject("ADODB.Recordset")
  rstCore.Fields.Append "FileName", 200, 255 '200=adVarChar
  rstCore.Fields.Append "FileType", 3 '3=adInteger
  rstCore.Open
  
  Application.StatusBar = "Identifying latest core CPT modules..."
  'get CurrentVersions.xml
  'get file list in cpt\Core
  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.SetProperty "SelectionLanguage", "XPath"
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
  strURL = strGitHub & "CurrentVersions.xml"
  If Not xmlDoc.Load(strURL) Then
    If xmlDoc.parseError.errorcode = -2146697210 Or -xmlDoc.parseError.errorcode = -2146697208 Then '</issue35>
      MsgBox "Please check your internet connection.", vbCritical + vbOKOnly, "Can't Connect"
    Else
      strMsg = "We're having trouble downloading modules:" & vbCrLf & vbCrLf  '</issue35>
      strMsg = strMsg & xmlDoc.parseError.errorcode & ": " & xmlDoc.parseError.reason & vbCrLf & vbCrLf '</issue35>
      strMsg = strMsg & "Please try the manual installation method instead." & vbCrLf & vbCrLf & "Would you like to open the online instructions now?"  '</issue35>
      If MsgBox(strMsg, vbExclamation + vbYesNo, "XML Error") = vbYes Then
        If Not Application.FollowHyperlink("https://github.com/clearplan/cpt#installation") Then
          MsgBox "Your organization appears to have blocked this url." & vbCrLf & vbCrLf & "Please contact help@ClearPlanConsulting.com for further assistance.", vbCritical + vbOKOnly, "Apologies!"
        End If
      End If
    End If
    GoTo this_project
  Else
    'download cpt/core/*.* to user's tmp directory
    For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
      If xmlNode.SelectSingleNode("Directory").Text = "Core" Then
        Application.StatusBar = "Fetching " & xmlNode.SelectSingleNode("Name").Text & "..."
        rstCore.AddNew Array(0, 1), Array(xmlNode.SelectSingleNode("FileName").Text, xmlNode.SelectSingleNode("Type").Text)
        rstCore.Update
        'get ThisProject status for later
        If xmlNode.SelectSingleNode("FileName").Text = "ThisProject.cls" Then
          strVersion = xmlNode.SelectSingleNode("Version").Text
        End If
        'build the url of the download
        strURL = strGitHub
        If Len(xmlNode.SelectSingleNode("Directory").Text) > 0 Then
          strURL = strURL & xmlNode.SelectSingleNode("Directory").Text & "/"
        End If
        strFileName = xmlNode.SelectSingleNode("FileName").Text
        strURL = strURL & strFileName
frx:
        Set xmlHttpDoc = CreateObject("Microsoft.XMLHTTP")
        xmlHttpDoc.Open "GET", strURL, False
        xmlHttpDoc.Send
        If xmlHttpDoc.Status = 200 Then
          Set oStream = CreateObject("ADODB.Stream")
          oStream.Open
          oStream.Type = 1 'adTypeBinary
          oStream.Write xmlHttpDoc.responseBody
          If Dir(strDir & "\" & strFileName) <> vbNullString Then Kill strDir & "\" & strFileName
          oStream.SaveToFile strDir & "\" & strFileName
          oStream.Close
          'need to fetch the .frx first
          If Right(strURL, 4) = ".frm" Then
            strURL = Replace(strURL, ".frm", ".frx")
            strFileName = Replace(strFileName, ".frm", ".frx")
            GoTo frx
          ElseIf Right(strURL, 4) = ".frx" Then
            strURL = Replace(strURL, ".frx", ".frm")
            strFileName = Replace(strFileName, ".frx", ".frm")
          End If
        Else
          strError = strError & "- " & strFileName & vbCrLf
          GoTo next_xmlNode
        End If

        'remove if exists
        strModule = Left(strFileName, InStr(strFileName, ".") - 1)
        If strModule = "ThisProject" Then GoTo next_xmlNode
        blnExists = False
        For Each vbComponent In ThisProject.VBProject.VBComponents
          If vbComponent.Name = strModule Then
            Application.StatusBar = "Removing obsolete version of " & vbComponent.Name
            'Debug.Print Application.StatusBar
            '<issue19> revised
            vbComponent.Name = vbComponent.Name & "_" & Format(Now, "hhnnss")
            DoEvents
            ThisProject.VBProject.VBComponents.Remove vbComponent 'ThisProject.VBProject.VBComponents(CStr(vbComponent.Name))
            DoEvents '</issue19>
            Exit For
          End If
        Next vbComponent

        'import the module - skip ThisProject which needs special handling
        If strModule <> "ThisProject" Then
          Application.StatusBar = "Importing " & strFileName & "..."
          'Debug.Print Application.StatusBar
          ThisProject.VBProject.VBComponents.Import strDir & "\" & strFileName
          '<issue19> added
          DoEvents '</issue19>

          '<issue24>remove the whitespace added by VBE import/export
          With ThisProject.VBProject.VBComponents(strModule).CodeModule
            For lngLine = .CountOfDeclarationLines To 1 Step -1
              If Len(.Lines(lngLine, 1)) = 0 Then .DeleteLines lngLine, 1
              DoEvents
            Next lngLine
          End With '</issue24>

        End If

      End If
next_xmlNode:
    Next xmlNode
  End If

  Application.StatusBar = "CPT Modules imported."

this_project:

  '<issue35>
  'update user's ThisProject - if it downloaded correctly, or was copied in correctly
  strFileName = strDir & "\ThisProject.cls"
  If Dir(strFileName) <> vbNullString Then 'it was downloaded, import it
    'rename the file and import it
    strCptFileName = Replace(strFileName, "ThisProject", "cptThisProject_cls")
    If Dir(strCptFileName) <> vbNullString Then Kill strCptFileName
    Name strFileName As strCptFileName
    'import the module
    If cptModuleExists("cptThisProject_cls") Then
      ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents("cptThisProject_cls")
      DoEvents
    End If
    Set cmCptThisProject = ThisProject.VBProject.VBComponents.Import(strCptFileName).CodeModule
  ElseIf cptModuleExists("cptThisProject_cls") Then 'it was copied in
    Set cmCptThisProject = ThisProject.VBProject.VBComponents("cptThisProject_cls").CodeModule
  Else 'ThisProject not imported or downloaded, so skip
    GoTo skip_import
  End If '</issue35>

  'avoid messy overwrites of ThisProject
  Set cmThisProject = ThisProject.VBProject.VBComponents("ThisProject").CodeModule
  '<issue10> revised
  'If cmThisProject.Find("<cpt_version>", 1, 1, cmThisProject.CountOfLines, 1000, True, True) = True Then
  If cmThisProject.Find("<cpt_version>", 1, 1, cmThisProject.CountOfLines, 1000, False, True) = True Then
  '</issue10>
    strMsg = "Your 'ThisProject' module has already been updated to work with the ClearPlan toolbar." & vbCrLf & vbCrLf
    strMsg = strMsg & "Would you like to reset it? This will only overwrite CodeModule lines appended with '</cpt>'" & vbCrLf & vbCrLf
    strMsg = strMsg & "(Please note: if you have made modifications to your ThisProject module, you may need to review them if you proceed.)"
    If MsgBox(strMsg, vbExclamation + vbYesNo, "Danger, Will Robinson!") = vbYes Then
      For lngLine = cmThisProject.CountOfLines To 1 Step -1
        If InStr(cmThisProject.Lines(lngLine, 1), "'</cpt>") > 0 Then
          cmThisProject.DeleteLines lngLine
          DoEvents
        End If
      Next lngLine
    Else
      GoTo skip_import
    End If
  End If

  'grab the imported code
  '<issue35>
  If Len(strVersion) = 0 Then 'grab the version
    strVersion = cptRegEx(ThisProject.VBProject.VBComponents("cptThisProject_cls").CodeModule.Lines(1, 1000), "<cpt_version>.*</cpt_version>")
    strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
  End If '</issue35>
  Set rstCode = CreateObject("ADODB.Recordset")
  rstCode.Fields.Append "EVENT", 200, 255 '200=adVarChar
  rstCode.Fields.Append "LINES", 201, 1 '201=adLongVarChar;1=adParamInput
  rstCode.Open
  With cmCptThisProject
    For Each vEvent In Array("Project_Activate", "Project_Open")
      rstCode.AddNew Array(0, 1), Array(CStr(vEvent), .Lines(.ProcStartLine(CStr(vEvent), 0) + 2, .ProcCountLines(CStr(vEvent), 0) - 3)) '0 = vbext_pk_Proc
      rstCode.Update
    Next vEvent
  End With
  ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents(cmCptThisProject.Parent.Name)
  DoEvents
  If cptModuleExists("ThisProject1") Then
    ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents("ThisProject1")
  End If
  If cptModuleExists("cptThisProject_cls") Then
    ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents("cptThisProject_cls")
  End If
  '<issue19> added
  DoEvents '</issue19>

  'add the events, or insert new text
  'three cases: empty or not empty (code exists or not)
  For Each vEvent In Array("Project_Activate", "Project_Open")

    'if event exists then insert code else create new event handler
    With cmThisProject
      If .CountOfLines > .CountOfDeclarationLines Then 'complications
        rstCode.MoveFirst
        rstCode.Find "EVENT='" & vEvent & "'"
        If .Find("Sub " & CStr(vEvent), 1, 1, .CountOfLines, 1000) = True Then
          'find its line number
          lngEvent = .ProcBodyLine(CStr(vEvent), 0) '= vbext_pk_Proc
          'import them if they *as a group* don't exist
          If .Find(rstCode(1), .ProcStartLine(CStr(vEvent), 0), 1, .ProcCountLines(CStr(vEvent), 0), 1000) = False Then  'vbext_pk_Proc
            .InsertLines lngEvent + 1, rstCode(1)
            DoEvents
          End If
        Else 'create it
          'create it, returning its line number
          lngEvent = .CreateEventProc(Replace(CStr(vEvent), "Project_", ""), "Project")
          'insert cpt code after line number
          .InsertLines lngEvent + 1, rstCode(1)
          DoEvents
        End If
      Else 'easy
        rstCode.MoveFirst
        rstCode.Find "EVENT='" & vEvent & "'"
        'create it, returning its line number
        lngEvent = .CreateEventProc(Replace(CStr(vEvent), "Project_", ""), "Project")
        'insert cpt code after line number
        .InsertLines lngEvent + 1, rstCode(1)
        DoEvents
      End If 'lines exist
    End With 'thisproject.codemodule

    'add version if not exists
    With cmThisProject
      If .Find("<cpt_version>", 1, 1, .CountOfLines, 1000) = False Then
        .InsertLines 1, "'<cpt_version>" & strVersion & "</cpt_version>" & vbCrLf
        DoEvents
      End If
    End With
  Next vEvent

  'leave no trace
  'If Dir(strCptFileName, vbNormal) <> vbNullString Then Kill strCptFileName

skip_import:

  If Len(strError) > 0 Then
    strError = "The following modules did not download correctly:" & vbCrLf & strError & vbCrLf & vbCrLf & "Please contact help@ClearPlanConsulting.com for assistance."
    MsgBox strError, vbCritical + vbOKOnly, "Unknown Error"
    'Debug.Print strError
  End If
    
  'reset the toolbar
  Application.ScreenUpdating = False
  Application.FileNew
  DoEvents
  Application.FileCloseEx pjDoNotSave
  Application.ScreenUpdating = True
  GoTo exit_here
  
  strMsg = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
  strMsg = strMsg + "<mso:customUI "
  strMsg = strMsg + "xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >"
  strMsg = strMsg + vbCrLf & "<mso:ribbon startFromScratch=""false"" >"
  strMsg = strMsg + vbCrLf & "<mso:tabs>"
  strMsg = strMsg + cptBuildRibbonTab()
  strMsg = strMsg + vbCrLf & "</mso:tabs>"
  strMsg = strMsg + vbCrLf & "</mso:ribbon>"
  strMsg = strMsg + vbCrLf & "</mso:customUI>"
  ActiveProject.SetCustomUI (strMsg)

exit_here:
  On Error Resume Next
  Set Project = Nothing
  '<issue19> added
  Application.StatusBar = "" '</issue19>
  '<issue23> added
  Application.ScreenUpdating = True '</issue23>
  Set vbComponent = Nothing
  Set rstCode = Nothing
  Set cmThisProject = Nothing
  Set cmCptThisProject = Nothing
  Set oStream = Nothing
  Set xmlHttpDoc = Nothing
  Set xmlNode = Nothing
  Set xmlDoc = Nothing
  Set rstCore = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptSetup_bas", "cptSetup", Err, Erl)
  Resume exit_here
End Sub

Public Function cptBuildRibbonTab()
  Dim ribbonXML As String
  Dim lngCleanUp As Long

  'build ClearPlan Ribbon Tab XML
  ribbonXML = ribbonXML + vbCrLf & "<mso:tab id=""tCommon"" label=""ClearPlan"" >" 'insertBeforeQ=""mso:TabTask"">"

  'common tools
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""custom_view"" label=""View"" visible=""true"">"

  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:OutlineSymbolsShow"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:SummaryTasks"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:NameIndent"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:AutoFilterProject"" visible=""true""/>"
  
  ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mDateFormat"" label=""Date Format"" imageMso=""TimelineDateFormat"" visible=""true"" >" 'size=""large""
  
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mm_dd_yy"" label=""" & Format(Now, "m/d/yy") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mm_dd_yy"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mm_dd_yy_hh_mmAM"" label=""" & Format(Now, "m/d/yy hh:mm ampm") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mm_dd_yy_hh_mmAM"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_dd"" label=""" & Format(Now, "dd") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_dd"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_ddd_dd"" label=""" & Format(Now, "ddd dd") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_ddd_dd"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_ddd_hh_mmAM"" label=""" & Format(Now, "ddd hh:mm ampm") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_ddd_hh_mmAM"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_ddd_mm_dd"" label=""" & Format(Now, "ddd mm/dd") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_ddd_mm_dd"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_ddd_mm_dd_yy"" label=""" & Format(Now, "ddd mm/dd/yy") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_ddd_mm_dd_yy"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_ddd_mm_dd_yy_hh_mmAM"" label=""" & Format(Now, "ddd mm/dd/yy hh:mm ampm") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_ddd_mm_dd_yy_hh_mmAM"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_ddd_mmm_dd"" label=""" & Format(Now, "ddd mmm dd") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_ddd_mmm_dd"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_ddd_mmm_dd_yyy"" label=""" & Format(Now, "ddd mmm dd 'yy") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_ddd_mmm_dd_yyy"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_hh_mmAM"" label=""" & Format(Now, "hh:mm ampm") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_hh_mmAM"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mm_dd"" label=""" & Format(Now, "m/d") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mm_dd"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mm_dd_yyyy"" label=""" & Format(Now, "m/d/yyyy") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mm_dd_yyyy"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mmm_dd"" label=""" & Format(Now, "mmm dd") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mmm_dd"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mmm_dd_hh_mmAM"" label=""" & Format(Now, "mmm dd hh:mm ampm") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mmm_dd_hh_mmAM"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mmm_dd_yyy"" label=""" & Format(Now, "mmm dd 'yy") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mmm_dd_yyy"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mmmm_dd"" label=""" & Format(Now, "mmmm dd") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mmmm_dd"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mmmm_dd_yyyy"" label=""" & Format(Now, "mmmm dd, yyyy") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mmmm_dd_yyyy"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_mmmm_dd_yyyy_hh_mmAM"" label=""" & Format(Now, "mmmm dd, yyyy hh:mm ampm") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_mmmm_dd_yyyy_hh_mmAM"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_Www_dd"" label=""" & "W" & Format(Now, "ww/dd") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_Www_dd"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""pjDate_Www_dd_yy_hh_mmAM"" label=""" & "W" & Format(Now, "ww/dd/yy hh:mm ampm") & """ imageMso=""TimelineDateFormat"" onAction=""cptDate_Www_dd_yy_hh_mmAM"" />"
  
  ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:FilterClear"" visible=""true""/>"
  
  ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  
  ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbResetAll"" size=""large"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetAll"" label=""Reset All"" imageMso=""FilterClear"" onAction=""cptResetAll"" screentip=""Reset All"" supertip=""Reset the current view based on your saved settings.""/>"  'in basCore_bas 'visible=""true""
  ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mResetAll"">"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Settings"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetAllSettings"" label=""Settings"" imageMso=""AdministrationHome"" onAction=""cptShowResetAll_frm"" visible=""true"" screentip=""Reset All Settings"" supertip=""Tweak your saved 'Reset All' settings.""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
  
  ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbWrapItUp"" >" 'size=""large""
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bWrapItUp"" label=""WrapItUp"" imageMso=""CollapseAll"" onAction=""cptWrapItUp"" supertip=""Collapse summary tasks starting from lowest level up to level 2. Defaults to your saved setting from Reset All or 2 if you don't have a saved setting yet."" />"   'in basCore_bas;visible=""true"" size=""large""
  ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mWrapItUp"">"
  'ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""WrapItUp to Level:"" />"
  'ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:OutlineShowAllTasks"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevelAll"" label=""All Subtasks"" imageMso=""OutlineTasksShowAll"" onAction=""cptWrapItUpAll"" visible=""true"" screentip=""Show All Subtasks"" supertip=""Show All Subtasks""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel1"" label=""Level 1"" imageMso=""_1"" onAction=""cptWrapItUp1"" visible=""true"" screentip=""WrapItUp to Level 1"" supertip=""WrapItUp to Level 1""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel2"" label=""Level 2"" imageMso=""_2"" onAction=""cptWrapItUp2"" visible=""true"" screentip=""WrapItUp to Level 2"" supertip=""WrapItUp to Level 2""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel3"" label=""Level 3"" imageMso=""_3"" onAction=""cptWrapItUp3"" visible=""true"" screentip=""WrapItUp to Level 3"" supertip=""WrapItUp to Level 3""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel4"" label=""Level 4"" imageMso=""_4"" onAction=""cptWrapItUp4"" visible=""true"" screentip=""WrapItUp to Level 4"" supertip=""WrapItUp to Level 4""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel5"" label=""Level 5"" imageMso=""_5"" onAction=""cptWrapItUp5"" visible=""true"" screentip=""WrapItUp to Level 5"" supertip=""WrapItUp to Level 5""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel6"" label=""Level 6"" imageMso=""_6"" onAction=""cptWrapItUp6"" visible=""true"" screentip=""WrapItUp to Level 6"" supertip=""WrapItUp to Level 6""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel7"" label=""Level 7"" imageMso=""_7"" onAction=""cptWrapItUp7"" visible=""true"" screentip=""WrapItUp to Level 7"" supertip=""WrapItUp to Level 7""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel8"" label=""Level 8"" imageMso=""_8"" onAction=""cptWrapItUp8"" visible=""true"" screentip=""WrapItUp to Level 8"" supertip=""WrapItUp to Level 8""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel9"" label=""Level 9"" imageMso=""_9"" onAction=""cptWrapItUp9"" visible=""true"" screentip=""WrapItUp to Level 9"" supertip=""WrapItUp to Level 9""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
  
  'ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bGroupReapply""  label=""ReGroup"" imageMso=""RefreshWebView"" onAction=""cptGroupReapply"" visible=""true"" supertip=""Reapply Group"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFilterReapply"" label=""ReFilter"" imageMso=""RefreshWebView"" onAction=""cptFilterReapply"" visible=""true"" supertip=""Reapply Filter"" />"
'  ribbonXML = ribbonXML + vbCrLf & "<mso:dialogBoxLauncher>"
'  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""dbl-reset"" screentip=""Settings for Reset All"" onAction=""cptShowResetAll_frm"" />"
'  ribbonXML = ribbonXML + vbCrLf & "</mso:dialogBoxLauncher>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  'task counters
  If cptModuleExists("cptCountTasks_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCount"" label=""Count"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountSelected"" label=""Selected"" imageMso=""NumberInsert"" onAction=""cptCountTasksSelected"" visible=""true""/>" 'SelectTaskCell
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountVisible"" label=""Visible"" imageMso=""NumberInsert"" onAction=""cptCountTasksVisible"" visible=""true""/>" 'SelectRows
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountAll"" label=""All"" imageMso=""NumberInsert"" onAction=""cptCountTasksAll"" visible=""true""/>" 'SelectWholeLayout
    ribbonXML = ribbonXML + vbCrLf & "<mso:dialogBoxLauncher>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""dbl-count"" screentip=""Status Bar Count Option"" onAction=""cptSetShowStatusBarTaskCount"" />"
    ribbonXML = ribbonXML + vbCrLf & "</mso:dialogBoxLauncher>"
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'text tools
  If cptModuleExists("cptText_bas") Then
        
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gTextTools"" label=""Text"" visible=""true"" >"
    If cptModuleExists("cptFilterByClipboard_bas") And cptModuleExists("cptFilterByClipboard_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bClipboard"" label=""Filter by Clipboard"" imageMso=""PasteOption"" onAction=""cptShowFilterByClipboard_frm"" visible=""true"" supertip=""Paste a list of Unique IDs or IDs from text, email, Excel, etc. to filter the current schedule. Accepts strings delimited by commas, tabs, or semicolons--or even tables, as long as the Unique ID (or ID) is the left-most column."" />"
    End If
    If cptModuleExists("cptAdvancedFilter_bas") And cptModuleExists("cptAdvancedFilter_frm") And cptModuleExists("cptAdvancedFilterEdit_frm") And cptModuleExists("cptFilterItem_cls") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAdvancedFilter"" label=""Advanced Filter"" imageMso=""PasteOption"" onAction=""cptAdvancedFilter"" visible=""true"" supertip=""Paste a list of Unique IDs or IDs from text, email, Excel, etc. to filter the current schedule. Accepts strings delimited by commas, tabs, or semicolons--or even tables, as long as the Unique ID (or ID) is the left-most column."" />"
    End If
    If cptModuleExists("cptDynamicFilter_bas") And cptModuleExists("cptDynamicFilter_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bDynamicFilter"" label=""Dynamic Filter"" imageMso=""FilterBySelection"" onAction=""cptShowDynamicFilter_frm"" visible=""true"" supertip=""Find-as-you-type. Example: Keep Selected task, filter the rest of the schedule for a predecessor, add a link, CTRL+BACKSPACE to return to task you kept. Then do the next one. Filter or Highlight filter, include summaries in the search, or include related summaries. Oh, and you can Undo. Pure awesomeness."" />"
    End If
    If cptModuleExists("cptText_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbText"" size=""large"" >"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAdvancedTextTools"" label=""Bulk Edit"" imageMso=""AdvancedFilterDialog"" onAction=""cptShowText_frm"" supertip=""Bulk prefix, append, real find/replace, enumeration, everyting you could want. Oh, and Undo. Go ahead, give it a try."" />" 'visible=""true""
      ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mText"">"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Utilities"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPrepend"" label=""Bulk Prepend"" imageMso=""RightArrow2"" onAction=""cptBulkPrepend"" visible=""true"" supertip=""Just what it sounds like."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAppend"" label=""Bulk Append"" imageMso=""LeftArrow2"" onAction=""cptBulkAppend"" visible=""true"" supertip=""Just what it sounds like."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bMyReplace"" label=""MyReplace"" imageMso=""ReplaceDialog"" onAction=""cptMyReplace"" visible=""true"" supertip=""Find/Replace only on selected tasks, in the selected field."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bEnumerate"" label=""Enumerate"" imageMso=""NumberingRestart"" onAction=""cptEnumerate"" visible=""true"" supertip=""Select a group of tasks, and then enumerate them."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrimText"" label=""Trim Task Names"" imageMso=""TextEffectsClear"" onAction=""cptTrimTaskNames"" visible=""true"" supertip=""For the 'Type A' folks out there, this trims leading and trailing spaces (and multiple spaces) in your task names (e.g., after pasting them in from Excel--cool, right?). Note: this applies to all non-external tasks in the project."" />"
      'ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bReplicateProcess"" label=""Replicate A Process (WIP)"" imageMso=""DuplicateSelectedSlides"" onAction=""cptReplicateProcess"" visible=""true"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFindDuplicates"" label=""Find Duplicate Task Names"" imageMso=""RemoveDuplicates"" onAction=""cptFindDuplicateTaskNames"" visible=""true"" supertip=""Clearly worded tasks represent well-defined tasks and are important for estimating and providing status. Click to find duplicate task names and create a report in Excel. Remember: Noun and Verb!"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetRowHeight"" label=""Reset Row Height"" imageMso=""RowHeight"" onAction=""cptResetRowHeight"" visible=""true"" supertip=""Another one for our fellow 'Type A' folks out there--reset all row heights after they get all jacked up. Give it a go; you'll like it."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAnnoyances"" label=""Annoyances"" imageMso=""SnapToRulerSubdivisions"" onAction=""cptCheckAnnoyances"" visible=""true"" supertip=""Yet another 'Type A' friendly utility--check for start times not equal to 8:00 AM or finish times not equal to 5:00 PM or fractional durations. Have another idea? Let us know help@ClearPlanConsulting.com."" />"
      ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
    Else
      ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mTextTools"" label=""Tools"" imageMso=""TextBoxInsert"" visible=""true"" >"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPrepend"" label=""Bulk Prepend"" imageMso=""RightArrow2"" onAction=""cptBulkPrepend"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAppend"" label=""Bulk Append"" imageMso=""LeftArrow2"" onAction=""cptBulkAppend"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bMyReplace"" label=""MyReplace"" imageMso=""ReplaceDialog"" onAction=""cptMyReplace"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bEnumerate"" label=""Enumerate"" imageMso=""NumberingRestart"" onAction=""cptEnumerate"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrimText"" label=""Trim Task Names"" imageMso=""TextEffectsClear"" onAction=""cptTrimTaskNames"" visible=""true""/>"
      'ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bReplicateProcess"" label=""Replicate A Process"" imageMso=""DuplicateSelectedSlides"" onAction=""cptReplicateProcess"" visible=""true"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFindDuplicates"" label=""Find Duplicate Task Names"" imageMso=""RemoveDuplicates"" onAction=""cptFindDuplicateTaskNames"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetRowHeight"" label=""Reset Row Height"" imageMso=""RowHeight"" onAction=""cptResetRowHeight"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'trace tools
  If cptModuleExists("cptCriticalPathTools_bas") Or cptModuleExists("cptCriticalPath_bas") Or cptModuleExists("cptNetworkBrowser_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCPA"" label=""Trace and Mark"" visible=""true"">"
    If cptModuleExists("cptCriticalPathTools_bas") And cptModuleExists("cptCriticalPath_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbTrace"" size=""large"" >"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrace"" imageMso=""TaskDrivers"" label=""Driving Paths"" onAction=""DrivingPaths"" supertip=""Select a target task, get the primary, secondary, and tertiary driving paths to that task."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mTrace"">"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Export"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPowerPoint"" label="">> PowerPoint"" imageMso=""SlideNew"" onAction=""cptExportCriticalPathSelected"" supertip=""Select a target task, get the primary, secondary, and tertiary driving paths to that task--and export them to PowerPoint."" />"
      ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSinglePath"" label=""Driving Path"" imageMso=""TaskDrivers"" onAction=""cptDrivingPath"" visible=""true"" size=""large"" supertip=""Select a target task, get the driving path."" />"
    Else
      If cptModuleExists("cptCriticalPath_bas") Then
        ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrace"" label=""Driving Path"" imageMso=""TaskDrivers"" onAction=""DrivingPaths"" visible=""true"" size=""large"" />"
      End If
      If cptModuleExists("cptCriticalPathTools_bas") Then
        ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bExport"" label="">> PowerPoint"" imageMso=""SlideNew"" onAction=""cptExportCriticalPathSelected"" visible=""true"" size=""large"" />"
      End If
    End If
    If cptModuleExists("cptNetworkBrowser_bas") And cptModuleExists("cptNetworkBrowser_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bNetworkBrowser"" label=""Network Browser"" imageMso=""ViewPredecessorsSuccessorsShow"" onAction=""cptShowNetworkBrowser_frm"" visible=""true"" size=""large"" supertip=""Jump to, and/or trace, predecessors and successors using the Network Diagram view in full screen or in the details pane."" />"
    End If
    If cptModuleExists("cptSaveMarked_bas") And cptModuleExists("cptSaveMarked_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bMarkSelected"" label=""Mark"" imageMso=""ApproveApprovalRequest"" onAction=""cptMarkSelected"" visible=""true"" supertip=""Mark selected task(s)"" />" 'size=""large""
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSaveMarked"" label=""Save"" imageMso=""Archive"" onAction=""cptSaveMarked"" visible=""true"" supertip=""Save currently marked tasks for later import."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bImportMarked"" label=""Import"" imageMso=""ApproveApprovalRequest"" onAction=""cptShowSaveMarked_frm"" visible=""true"" supertip=""Import saved sets of marked tasks."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bMarkedApply"" label=""Filter Marked"" imageMso=""FilterToggleFilter"" onAction=""cptMarked"" visible=""true"" supertip=""Filter Marked Tasks"" />" 'size=""large""
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUnmarkAll"" label=""Unmark All"" imageMso=""RejectApprovalRequest"" onAction=""cptClearMarked"" visible=""true"" supertip=""Unmark all currently marked tasks."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUnmark"" label=""Unmark"" imageMso=""RejectApprovalRequest"" onAction=""cptUnmarkSelected"" visible=""true"" supertip=""Unmark selected task(s)"" />" 'size=""large""
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
'  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gIntegration2"" visible=""true"" >"
'  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bIntegrationSettings3"" label=""Field Mapping"" imageMso=""Settings"" screentip=""Integration Settings"" onAction=""cptGetValidMap"" size=""large"" />"
'  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  
  'schedule
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gStatus"" label=""Schedule"" visible=""true"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mHealth"" label=""Health"" imageMso=""CheckWorkflow"" visible=""true"" size=""large"" >"
   ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""DCMA EVMS Compliance Metric (DECM)"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
   ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bDECM"" label=""DECM Dashboard (v8.0)"" imageMso=""CheckWorkflow"" onAction=""cptDECM_GET_DATA"" visible=""true"" supertip=""DECM Dashboard (v8.0)"" />"
   ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
   ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bIntegrationSettings1"" label=""Integration Settings"" imageMso=""Settings"" onAction=""cptGetValidMap"" visible=""true"" supertip=""Set, edit, and confirm Integration Settings"" />"
'  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""DCMA 14"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
'  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bDCMA14"" label=""DCMA 01"" imageMso=""CheckWorkflow"" onAction=""cptDCMA14_GET_DATA"" visible=""true"" supertip=""DCMA 14-pt Analysis"" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mStatus"" label=""Status"" imageMso=""UpdateAsScheduled"" visible=""true"" size=""large"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""General"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  If cptModuleExists("cptQBD_frm") And cptModuleExists("cptQBD_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bQBD"" label=""Quantifiable Backup Data (QBD)"" imageMso=""ExportExcel"" onAction=""cptShowQBD_frm"" visible=""true"" supertip=""Yes, Quantifiable Backup Data."" />"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Before Status"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cpt_bAdvanceStatusDate"" label=""Advance Status Date"" imageMso=""CalendarToolSelectDate"" onAction=""cptAdvanceStatusDate"" visible=""true"" supertip=""Advance the Status Date prior to kicking off a status cycle."" />"
  If cptModuleExists("cptAgeDates_bas") And cptModuleExists("cptAgeDates_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cpt_bAgeDates"" label=""Age Dates"" imageMso=""CalendarToolSelectDate"" onAction=""cptShowAgeDates_frm"" visible=""true"" supertip=""Keep a rolling history of the current schedule.""  />"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCaptureWeek1"" label=""Capture Week"" imageMso=""RefreshWebView"" onAction=""cptCaptureWeek"" visible=""true"" supertip=""OPTIONAL: Capture the Current Schedule before updates if you want to record task-level notes for the current status date."" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Status Export &amp;&amp; Import"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  If cptModuleExists("cptStatusSheet_bas") And cptModuleExists("cptStatusSheet_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bStatusSheet"" label=""Create Status Sheet(s)"" imageMso=""ExportExcel"" onAction=""cptShowStatusSheet_frm"" visible=""true"" supertip=""Just what it sounds like. Include any fields you like. Settings are saved between sessions."" />" 'DateAndTimeInsertOneNote
  End If
  If cptModuleExists("cptStatusSheetImport_bas") And cptModuleExists("cptStatusSheetImport_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bStatusSheetImport"" label=""Import Status Sheet(s)"" imageMso=""ImportExcel"" onAction=""cptShowStatusSheetImport_frm"" visible=""true"" supertip=""Just what it sounds like. (Note: Assignment ETC is at the Assignment level, so use the Task Usage view to review after import.)"" />"
  End If
  
  ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mScheduleUtilities"" label=""Utilities"" imageMso=""CheckWorkflow"" visible=""true"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Status Utilities"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  
  If cptModuleExists("cptSmartDuration_frm") And cptModuleExists("cptSmartDuration_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSmartDuration"" label=""Smart Duration"" imageMso=""CalendarToolSelectDate"" onAction=""cptShowSmartDuration_frm"" visible=""true"" supertip=""We've all been there: how many days between Time Now and the finish date the CAM just gave me? No more guess work: click here and improve your life."" />"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUnstatused"" label=""Find Unstatused"" imageMso=""UpdateAsScheduled"" onAction=""cptFindUnstatusedTasks"" visible=""true"" supertip=""Find tasks not statused through 'Time Now'."" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bRetainETC"" label=""Mark On Track (Retain ETC)"" imageMso=""UpdateAsScheduled"" onAction=""cptMarkOnTrackRetainETC"" visible=""true"" supertip=""Mark on Track and Retain Remaining Work (ETC)."" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bRespread"" label=""Respread Assignments"" imageMso=""UpdateAsScheduled"" onAction=""cptRespreadAssignmentWork"" visible=""true"" supertip=""Respread Assignment Work/Cost so that Assignment Finish Dates match Task Finish Date."" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAssignmentsNoETC"" label=""Find Assignments without ETC"" imageMso=""UpdateAsScheduled"" onAction=""cptFindAssignmentsWithoutWork"" supertip=""Find Assignments on incomplete tasks with zero remaining work. These can lead to odd task start times."" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""After Status"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBlameReport"" label=""The Blame Report"" imageMso=""ContactProperties"" onAction=""cptBlameReport"" visible=""true"" supertip=""Find out which tasks slipped from last period."" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCaptureWeek2"" label=""Capture Week"" imageMso=""RefreshWebView"" onAction=""cptCaptureWeek"" visible=""true"" supertip=""Capture the Current Schedule after updates to compare against past and future weeks during execution. This is required for certain metrics (e.g., CEI, all Trending) to run properly."" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCompletedWork"" label=""Export Completed WPCNs"" imageMso=""DisconnectFromServer"" onAction=""cptExportCompletedWork"" visible=""true"" supertip=""Export Completed WPCNs for closure in the time card system."" />"
  If cptModuleExists("cptTaskHistory_bas") And cptModuleExists("cptTaskHistory_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTaskHistory"" label=""Task History"" imageMso=""Archive"" onAction=""cptShowTaskHistory_frm"" visible=""true"" supertip=""Explore selected task history, take notes, export history, etc. Requires consistent use of Capture Week."" />"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Analyze"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFindCompleteThrough"" label=""Analyze CompleteThrough"" imageMso=""UpdateAsScheduled"" onAction=""cptFindCompleteThrough"" visible=""true"" supertip=""Analyze unexpected 'CompleteThrough' progress bar in Gantt Chart."" />"
  
  'todo: account for EV Tool in cptValidateEVP
  'ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bValidateEVT"" enabled=""false"" label=""Validate EVT"" imageMso=""RefreshWebView"" onAction=""cptAnalyzeEVT"" visible=""true"" supertip=""Validate EVT - e.g., ensure incomplete 50/50 tasks with Actual Start are marked as 50% EV % complete."" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  
  'metrics
  If cptModuleExists("cptMetrics_bas") Then
'    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gMetrics"" label=""Metrics"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mSchedule"" label=""Metrics"" imageMso=""ChartTypeLineInsertGallery"" visible=""true"" size=""large"" >" 'UpdateAsScheduled
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Schedule Metrics"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptSPI"" label=""Schedule Performance Index (SPI) in hours"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetSPI"" visible=""true"" supertip=""SPI (in hours) relies on timephased baseline work and EV% stored in Physical % Complete."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptSV"" label=""Schedule Variance (SV) in hours"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetSV"" visible=""true"" supertip=""SV (in hours) relies on timephased baseline work and EV% stored in Physical % Complete."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptBEI"" label=""Baseline Execution Index (BEI)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetBEI"" visible=""true"" supertip=""Just what it sounds like..."" />"
    If CLng(Left(Application.Build, 2)) >= 12 Then 'CPLI only available in versions after 2010
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptCPLI"" label=""Critical Path Length Index (CPLI)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetCPLI"" visible=""true"" supertip=""Select a target task, click to get the CPLI. Raw calculation based on time now and total slack."" />"
    Else
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptCPLI"" enabled=""false"" label=""Critical Path Length Index (CPLI)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetCPLI"" visible=""true"" supertip=""Select a target task, click to get the CPLI. Raw calculation based on time now and total slack. (Feature not available in this version of MS Project)"" />"
    End If
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptCEI"" label=""Current Execution Index (CEI)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetCEI"" visible=""true"" supertip=""Tracks forecast accuracy between periods. Be sure to 'Capture Week' in previous period's file under Schedule > Status > Capture Week."" />"
'    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptTFCI"" enabled=""false"" label=""Total Float Consumption Index (TFCI)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetCEI"" visible=""true"" supertip=""Measures forecast accuracy between reporting periods"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptES"" label=""Earned Schedule"" imageMso=""CalendarToolSelectDate"" onAction=""cptGetEarnedSchedule"" visible=""true"" supertip=""Just what it sounds like. See the NDIA Predictive Measures Guide for more information."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptCaptureAllMetrics"" label=""Capture All Metrics"" imageMso=""DataViewDetailsView"" onAction=""cptCaptureAllMetrics"" visible=""true"" supertip=""Capture all metrics above for this program for this period."" />"
    
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Schedule Metrics Trends"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptSPItrend"" label=""SPI Trend in hours"" imageMso=""ChartTypeLineInsertGallery"" onAction=""cptGetTrend_SPI"" visible=""true"" supertip=""Relies on timephased baseline work and your metrics settings for EV % Complete."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptBEItrend"" label=""BEI Trend"" imageMso=""ChartTypeLineInsertGallery"" onAction=""cptGetTrend_BEI"" visible=""true"" supertip=""Just what it sounds like..."" />"
    If CLng(Left(Application.Build, 2)) >= 12 Then 'CPLI only available in versions after 2010
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptCPLItrend"" label=""CPLI Trend"" imageMso=""ChartTypeLineInsertGallery"" onAction=""cptGetTrend_CPLI"" visible=""true"" supertip=""Create a chart of CPLI Trend."" />"
    Else
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptCPLItrend"" enabled=""false"" label=""CPLI Trend"" imageMso=""ChartTypeLineInsertGallery"" onAction=""cptGetTrend_CPLI"" visible=""true"" supertip=""Create a chart of CPLI Trend. (Feature not available in this version of MS Project)"" />"
    End If
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptCEItrend"" label=""CEI Trend"" imageMso=""ChartTypeLineInsertGallery"" onAction=""cptGetTrend_CEI"" visible=""true"" supertip=""Just what it sounds like..."" />"
    'todo: TFCI Trend
    If cptModuleExists("cptResourceDemand_bas") And cptModuleExists("cptResourceDemand_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Staffing Metrics"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptResourceDemandExcel"" label=""Staffing Profile"" imageMso=""Chart3DColumnChart"" onAction=""cptShowExportResourceDemand_frm"" visible=""true"" supertip=""Export timephased assignment remaining work, baseline work, costs (any or all rate sets), and your choice of extra fields. Settings are saved between sessions."" />" 'size=""large""
    End If
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Other"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptBCWS"" label=""Budgeted Cost of Work Scheduled (BCWS) in hours"" imageMso=""NumberInsert"" onAction=""cptGetBCWS"" visible=""true"" supertip=""Timephased BCWS/PV (in hours)."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptBCWP"" label=""Budgeted Cost of Work Performed (BCWP) in hours"" imageMso=""NumberInsert"" onAction=""cptGetBCWP"" visible=""true"" supertip=""Timephased BCWP/EV (in hours)--relies on baseline work and Physical % Complete."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptBCWR"" label=""Budgeted Cost of Work Remaining (BCWR) in hours"" imageMso=""NumberInsert"" onAction=""cptGetBCWR"" visible=""true"" supertip=""Budgeted Cost of Work Remaining = (BAC - BCWP)"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptBAC"" label=""Budget at Complete (BAC) in hours"" imageMso=""NumberInsert"" onAction=""cptGetBAC"" visible=""true"" supertip=""Budget at Complete (BAC) in hours"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptETC"" label=""Estimate to Complete (ETC) in hours"" imageMso=""NumberInsert"" onAction=""cptGetETC"" visible=""true"" supertip=""Estimate to Complete (ETC) in hours"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptLSLF"" label=""Late Starts and Finishes"" imageMso=""ChartTypeLineInsertGallery"" onAction=""cptLateStartsFinishes"" visible=""true"" supertip=""Late Starts and Finishes Chart"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptHitTask"" label=""Hit Task %"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetHitTask"" visible=""true"" supertip=""Because it's (still) on the Gold Card."" />"

    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Data &amp;&amp; Settings"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
    If cptModuleExists("cptMetricsData_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptEditMetricsData"" label=""View &amp;&amp; Edit Metrics Data"" imageMso=""DataValidation"" onAction=""cptShowMetricsData_frm"" visible=""true"" supertip=""Review and delete metrics records for this program."" />"
    End If
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptExportAllData"" label=""Export Metrics Data"" imageMso=""ExportExcel"" onAction=""cptExportMetricsData"" visible=""true"" supertip=""Export stored metrics data for this program to Excel."" />"
    If cptModuleExists("cptIntegration_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptMetricsSettings"" label=""Metrics Settings"" imageMso=""Settings"" onAction=""cptGetValidMap"" visible=""true"" supertip=""Settings required for some EV-ish metrics."" />"
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
'    ribbonXML = ribbonXML + vbCrLf & "<mso:dialogBoxLauncher>"
'    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""test"" screentip=""Concept of operations"" onAction=""cptShowSettings_frm"" />"
'    ribbonXML = ribbonXML + vbCrLf & "</mso:dialogBoxLauncher>"
  End If
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
    
  'integration
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gIntegration"" label=""Integration"" visible=""true"" >"
  'outline codes
  If cptModuleExists("cptBackbone_frm") And cptModuleExists("cptBackbone_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBackbone"" label=""Outline Codes"" imageMso=""OrganizationChartLayoutRightHanging"" onAction=""cptShowBackbone_frm"" visible=""true"" size=""large"" supertip=""Quickly create or edit Outline Codes (CWBS, IMP, etc.); import and/or export; create DI-MGMT-81334D, etc."" />"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  If cptModuleExists("cptIMSCobraExport_bas") And cptModuleExists("cptIMSCobraExport_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCOBRA"" label=""COBRA Export Tool"" imageMso=""Export"" onAction=""Export_IMS"" visible=""true"" supertip=""Validate that your IMS is ready for integration; create CSV transaction files for COBRA. Baseline, forecast, status, etc."" />"
  End If
  If cptModuleExists("cptCheckAssignments_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCheckAssignments"" label=""Check Assignments"" imageMso=""SynchronizationStatus"" onAction=""cptCheckAssignments"" visible=""true"" supertip=""Reconcile task vs assignment work, baselines, etc."" />"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bExportCodes"" label=""Export All Codes"" imageMso=""Export"" onAction=""cptExportAllCodes"" visible=""true"" supertip=""Export all Local Custom Fields with lookups to *.csv"" />"
  If cptModuleExists("cptAdjustment_bas") And cptModuleExists("cptAdjustment_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAdjustment"" label=""ETC Adjustments"" imageMso=""SynchronizationStatus"" onAction=""cptShowAdjustment_frm"" visible=""true"" supertip=""Bulk adjust ETCs by resource, to given target, by percentage, or by a given amount."" />"
  End If
  
  If cptModuleExists("cptCostRateTables_bas") And cptModuleExists("cptCostRateTables_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCostRateTables"" label=""Cost Rate Tables"" imageMso=""DataTypeCurrency"" onAction=""cptShowCostRateTables_frm""  size=""large"" visible=""true"" supertip=""Easily Import and Export Cost Rate Tables."" />"
  End If
  
  'integration settings
  If cptModuleExists("cptIntegration_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:dialogBoxLauncher>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bIntegrationSettings2"" screentip=""Integration Settings"" onAction=""cptGetValidMap"" />"
    ribbonXML = ribbonXML + vbCrLf & "</mso:dialogBoxLauncher>"
  End If
  
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  
  'calendars
  If (cptModuleExists("cptFiscal_frm") And cptModuleExists("cptFiscal_bas")) Or (cptModuleExists("cptCalendarExceptions_frm") And cptModuleExists("cptCalendarExceptions_bas")) Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCalendars"" label=""Calendars"" visible=""true"" >"
    If cptModuleExists("cptFiscal_frm") And cptModuleExists("cptFiscal_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFiscal"" label=""Fiscal"" imageMso=""MonthlyView"" onAction=""cptShowFiscal_frm"" visible=""true"" supertip=""Maintain a fiscal calendar for various reports."" />"
    End If
    If cptModuleExists("cptCalendarExceptions_frm") And cptModuleExists("cptCalendarExceptions_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCalDetails"" label=""Details"" imageMso=""MonthlyView"" onAction=""cptShowCalendarExceptions_frm"" visible=""true"" supertip=""Export Calendar Exceptions, WorkWeeks, and settings."" />"
      If ActiveProject.Subprojects.Count > 0 Then
        ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCalCompare"" label=""Compare"" imageMso=""MonthlyView"" onAction=""cptCalendarCompareMain"" visible=""true"" supertip=""Compare Calendar Exceptions between Master Project and Subprojects."" />"
      End If
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'custom fields
  If (cptModuleExists("cptDataDictionary_frm") And cptModuleExists("cptDataDictionary_bas")) Or (cptModuleExists("cptSaveLocal_bas") And cptModuleExists("cptSaveLocal_frm")) Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCustomFields"" label=""Custom Fields"" visible=""true"">"
    If cptModuleExists("cptDataDictionary_frm") And cptModuleExists("cptDataDictionary_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bDataDictionary"" imageMso=""ReadingMode"" label=""Data Dictionary"" onAction=""cptShowDataDictionary_frm"" supertip=""Provide a description of each custom field; create export in Excel for deliverables; share dictionary. Settings are saved between sessions."" />" 'size=""large""
    End If
    If cptModuleExists("cptFieldBuilder_bas") And cptModuleExists("cptFieldBuilder_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBuilder"" imageMso=""CustomFieldDialog"" label=""Field Builder"" onAction=""cptShowFieldBuilder_frm"" supertip=""A little help building common custom field pick lists, etc."" />" 'size=""large""
    End If
    If cptModuleExists("cptSaveLocal_bas") And cptModuleExists("cptSaveLocal_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bECFtoLCF"" imageMso=""CustomFieldDialog"" label=""ECF to LCF"" onAction=""cptShowSaveLocal_frm"" supertip=""Save Enterprise Custom Field (ECF) settings (and, optionally, task-level data) to Local Custom Fields (LCF). Settings are saved (by project) between sessions."" />" 'size=""large""
    End If
    If cptModuleExists("cptCustomFieldUsage_bas") And cptModuleExists("cptCustomFieldUsage_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLCFUsage"" imageMso=""CustomFieldDialog"" label=""LCF Usage"" onAction=""cptShowCustomFieldUsage_frm"" supertip=""Quickly Analyze what Local Custom Fields (LCF) are in use."" />" 'size=""large""
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'about
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gHelp"" label=""Help"" visible=""true"" >"
  If cptInternetIsConnected Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mHelp"" label=""Help"" imageMso=""Help"" visible=""true"" size=""large"" supertip=""Submit a bug report, feature request, or general feedback. Upgrade modules from the InterWebs."" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Upgrades"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUpdate"" label=""Check for Upgrades"" imageMso=""PreviousUnread"" onAction=""cptShowUpgrades_frm"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Contribute"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bIssue"" label=""Submit an Issue"" imageMso=""SubmitFormInfoPath"" onAction=""cptSubmitIssue"" visible=""true"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bRequest"" label=""Submit a Feature Request"" imageMso=""SubmitFormInfoPath"" onAction=""cptSubmitRequest"" visible=""true"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFeedback"" label=""Submit Other Feedback"" imageMso=""SubmitFormInfoPath"" onAction=""cptSubmitFeedback"" visible=""true"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Settings"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSettingsCPT"" label=""View CPT Settings"" imageMso=""Settings"" onAction=""cptShowSettings_frm"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Uninstall"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUninstall"" label=""Uninstall ClearPlan Toolbar"" imageMso=""TasksUnlink"" onAction=""cptUninstall"" visible=""true"" />"
    ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAbout"" label=""About"" imageMso=""Info"" onAction=""cptShowAbout_frm""  size=""large"" visible=""true"" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  ribbonXML = ribbonXML + vbCrLf & "</mso:tab>"
    
  'Debug.Print "<mso:customUI ""xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >" & ribbonXML
  cptBuildRibbonTab = ribbonXML

End Function

Function cptGetLatest(strModule As String) As String
  'objects
  Dim xmlDoc As Object
  Dim xmlNode As Object
  Dim oFile As Scripting.File
  Dim oFSO As Scripting.FileSystemObject
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strLatest As String
  Dim strURL As String
  Dim strFileName As String
  Dim strDir As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnStale As Boolean
  'variants
  'dates
  
  'todo: prompt the user no more than once a week if an update is available
  'todo: use it as a basis for the upgrades form?
  
  If Not cptInternetIsConnected Then GoTo exit_here
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  
  strFileName = strDir & "\cpt-latest.adtg"
  blnStale = True
  'does file exist?
  If Dir(strFileName) <> vbNullString Then
    'is it stale?
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFSO.GetFile(strFileName)
    'todo: use DateModified? DateCreated? notified as date or boolean?
    'todo: what if a file is installed that gets removed from the package?
    'todo: if a module gets removed then <cpt_version>remove</cpt_version>
    If oFile.DateCreated > DateAdd("d", -7, Now()) Then
      blnStale = False
    End If
    Set oFile = Nothing
    Set oFSO = Nothing
    If blnStale Then
      Kill strFileName
    End If
  End If
  
  'todo: if update is available then notify user and mark 'notified=true
  
  If blnStale Then
    'set up the recordset
    Set oRecordset = CreateObject("ADODB.Recordset")
    oRecordset.Fields.Append "Module", 200, 200 '200=adVarChar
    oRecordset.Fields.Append "Directory", 200, 200 '200=adVarChar
    oRecordset.Fields.Append "Current", 200, 200 '200=adVarChar
    oRecordset.Fields.Append "Notified", 11 '11=adBoolean
    oRecordset.Fields.Append "Installed", 200, 200 '200=adVarChar
    oRecordset.Fields.Append "Status", 200, 200 '200=adVarChar
    oRecordset.Open

    'get current versions
    Application.StatusBar = "Fetching latest versions..."
    DoEvents
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.SetProperty "SelectionLanguage", "XPath"
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
    strURL = strGitHub & "CurrentVersions.xml"
    If Not xmlDoc.Load(strURL) Then
      'MsgBox xmlDoc.parseError.errorcode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
      cptGetLatest = "error"
      GoTo exit_here
    Else
      For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
        oRecordset.AddNew
        oRecordset(0) = xmlNode.SelectSingleNode("Name").Text '>Module
        oRecordset(1) = xmlNode.SelectSingleNode("Directory").Text '>Directory
        oRecordset(2) = xmlNode.SelectSingleNode("Version").Text '>Latest
        oRecordset(3) = False
        If cptModuleExists(xmlNode.SelectSingleNode("Name").Text) Then
          oRecordset(4) = cptGetVersion(xmlNode.SelectSingleNode("Name").Text)
          oRecordset(5) = cptVersionStatus(cptGetVersion(xmlNode.SelectSingleNode("Name").Text), xmlNode.SelectSingleNode("Version").Text)
        End If
        oRecordset.Update
      Next xmlNode
    End If
    oRecordset.Save strFileName, adPersistADTG
    oRecordset.Close
  End If
  
  'now check latest version
  Set oRecordset = CreateObject("ADODB.REcordset")
  oRecordset.Open strFileName
  oRecordset.Filter = "Module='" & strModule & "'"
  If Not oRecordset.EOF Then
    strLatest = oRecordset(2)
    'note: updating a record will affect LastModified...
  End If
  oRecordset.Filter = 0
  oRecordset.Close
  
  cptGetLatest = strLatest

exit_here:
  On Error Resume Next
  Set xmlNode = Nothing
  Set xmlDoc = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptSetup_bas", "cptGetLatest", Err, Erl)
  Resume exit_here
End Function

Function cptGetVersion(strModule As String) As String
  Dim vbComponent As Object, strVersion As String
  If Not cptModuleExists(strModule) Then
    cptGetVersion = "<uninstalled>"
  Else
    Set vbComponent = ThisProject.VBProject.VBComponents(strModule)
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = cptRegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
    End If
    cptGetVersion = strVersion
  End If
  
End Function

Sub cptHandleErr(strModule As String, strProcedure As String, objErr As ErrObject, Optional lngErl As Long, Optional strProcessStep As String)
  'common error handling prompt
  'objects
  Dim oLink As MSProject.TaskDependency
  Dim oTask As MSProject.Task
  Dim oSubproject As MSProject.SubProject
  Dim oProfile As MSProject.Profile
  Dim oShell As Object
  'strings
  Dim strFileName As String
  Dim strErrNumber As String
  Dim strErrDescription As String
  Dim strErrSource As String
  Dim strInstalled As String
  Dim strLatest As String
  Dim strStatus As String
  Dim strMsg As String
  Dim strKey As String
  'longs
  Dim lngTotalResources As Long
  Dim lngTotalTasks As Long
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  Dim blnMaster As Boolean
  Dim blnInactive As Boolean
  Dim blnCPL As Boolean
  Dim blnResourceLoaded As Boolean
  Dim blnBeta As Boolean
  'variants
  'dates
    
  strErrNumber = CStr(Err.Number)
  strErrDescription = Err.Description
  strErrSource = Err.Source

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'todo: has the code been modified?
  'todo: not until CurrentVersions.xml is revised
  
  'is an upgrade available?
  strInstalled = cptGetVersion(strModule)
  strLatest = cptGetLatest(strModule)
  If strLatest = "error" Or strLatest = "" Then
    strStatus = "unavailable"
    strLatest = "unavailable"
  Else
    strStatus = cptVersionStatus(strInstalled, strLatest)
  End If
  blnBeta = False
  If InStr(strStatus, "upgrade") > 0 Then 'please upgrade, try again
    strMsg = "An error has occurred in module '" & strModule & "-" & strInstalled & "." & vbCrLf
    strMsg = strMsg & "An upgrade is available: " & strModule & "-" & strLatest & "." & vbCrLf & vbCrLf
    strMsg = strMsg & strErrNumber & ": " & strErrDescription & vbCrLf
    strMsg = strMsg & "Procedure: " & strProcedure & vbCrLf
    If Len(strProcessStep) > 0 Then
      strMsg = strMsg & "Process Step: " & strProcessStep & vbCrLf
    End If
    If lngErl > 0 Then
      strMsg = strMsg & "Line: " & ThisProject.VBProject.VBComponents(strModule).CodeModule.Lines(lngErl, 1) & vbCrLf
    End If
    strMsg = strMsg & vbCrLf
    strMsg = strMsg & "Would you like to upgrade to '" & strModule & "' (" & strLatest & ") now?"
    If MsgBox(strMsg, vbInformation + vbYesNo, "CPT Error") = vbYes Then
      cptShowUpgrades_frm
    Else
      Exit Sub
    End If
  ElseIf InStr(strStatus, "downgrade") > 0 Then 'beta tester, tell us, thank you
    blnBeta = True
    strMsg = "Thank you for being a beta tester!" & vbCrLf & vbCrLf
    strMsg = strMsg & "An error has occurred in module " & strModule & "-" & strInstalled & "." & vbCrLf
    strMsg = strMsg & "Latest release version is " & strModule & "-" & strLatest & "." & vbCrLf & vbCrLf
    strMsg = strMsg & strErrNumber & ": " & strErrDescription & vbCrLf
    strMsg = strMsg & "Procedure: " & strProcedure & vbCrLf
    If Len(strProcessStep) > 0 Then
      strMsg = strMsg & "Process Step: " & strProcessStep & vbCrLf
    End If
    If lngErl > 0 Then
      strMsg = strMsg & "Line: " & ThisProject.VBProject.VBComponents(strModule).CodeModule.Lines(lngErl, 1) & vbCrLf
    End If
    strMsg = strMsg & vbCrLf
    strMsg = strMsg & "Would you like to submit a ticket? (Click 'No' to keep testing...)"
    If MsgBox(strMsg, vbInformation + vbYesNo, "CPT Error") = vbNo Then GoTo exit_here
  Else 'strStatus = "unavailable" or "ok" or "error" then submit ticket
    strMsg = "An error has occurred in module " & strModule & "-" & strInstalled & "." & vbCrLf
    If strStatus = "ok" Then
      strMsg = strMsg & "You are running the latest version of this module." & vbCrLf & vbCrLf
    Else
      strMsg = strMsg & "Unable to query github.com for latest version of this module." & vbCrLf & vbCrLf
    End If
    strMsg = strMsg & strErrNumber & ": " & strErrDescription & vbCrLf
    strMsg = strMsg & "Procedure: " & strProcedure & vbCrLf
    If Len(strProcessStep) > 0 Then
      strMsg = strMsg & "Process Step: " & strProcessStep & vbCrLf
    End If
    If lngErl > 0 Then
      strMsg = strMsg & "Line: " & ThisProject.VBProject.VBComponents(strModule).CodeModule.Lines(lngErl, 1) & vbCrLf
    End If
    strMsg = strMsg & vbCrLf
    strMsg = strMsg & "Would you like to submit a support ticket?"
    If MsgBox(strMsg, vbExclamation + vbYesNo, "CPT Error") = vbNo Then GoTo exit_here
  End If
  
  Set oShell = CreateObject("WScript.Shell")
  
  'get error and system details
  Application.StatusBar = "Collecting anonymous error information..."
  strMsg = "-> Your browser will open to our ticketing system: https://clearplan.happyfox.com/new" & vbCrLf & vbCrLf
  strMsg = strMsg & "-> The below information will help you submit a support ticket..." & vbCrLf & vbCrLf
  strMsg = strMsg & "-> === TICKET DETAILS ===" & vbCrLf
  strMsg = strMsg & "-> CATEGORY: ClearPlan Toolbar CPT" & vbCrLf
  strMsg = strMsg & "-> CPT Issue Type: (select one: installation; use; other)" & vbCrLf
  strMsg = strMsg & "-> SUBJECT: " & strModule & "-" & strInstalled & IIf(blnBeta, " (beta)", "") & " (Error: " & strErrNumber & ")" & vbCrLf
  strMsg = strMsg & "-> MESSAGE: (COPY & PASTE EVERYTHING BETWEEN THE TWO LONG DOTTED LINES and DO NOT EDIT; SEE FINAL STEPS AT BOTTOM)" & vbCrLf
  strMsg = strMsg & String(80, "-") & vbCrLf
  strMsg = strMsg & "[Please REPLACE THIS LINE with any notes or comments you'd like to add.]" & vbCrLf
  strMsg = strMsg & "EXAMPLE: I'm trying to run Status Sheets and I keep getting this error..." & vbCrLf
  strMsg = strMsg & "A screenshot with your dialog selections (e.g., Status Sheet) would also be helpful (see ticket page where it says 'drag and drop file')." & vbCrLf & vbCrLf
  
  'FILE
  blnMaster = ActiveProject.Subprojects.Count > 0
  strMsg = strMsg & "--- FILE ---" & vbCrLf 'most likely culprit
  If Left(ActiveProject.Path, 2) = "<>" Or Left(ActiveProject.Path, 4) = "http" Then
    strMsg = strMsg & "Type: Cloud/PWA" & vbCrLf
  Else
    strMsg = strMsg & "Type: Local or Network (*.mpp)" & vbCrLf
  End If
  If blnMaster Then
    strMsg = strMsg & "Subprojects: " & ActiveProject.Subprojects.Count & vbCrLf
    strMsg = strMsg & "Master Project: " & Format(ActiveProject.Tasks.Count, "#,##0") & " tasks; " & Format(ActiveProject.ResourceCount, "#,##0") & " resources" & vbCrLf
    lngTotalTasks = ActiveProject.Tasks.Count
    lngTotalResources = ActiveProject.ResourceCount
    For Each oSubproject In ActiveProject.Subprojects
      strMsg = strMsg & "- Subproject " & oSubproject.Index & ": " & Format(oSubproject.SourceProject.Tasks.Count, "#,##0") & " tasks; " & Format(oSubproject.SourceProject.ResourceCount, "#,##0") & " resources" & vbCrLf
      lngTotalTasks = lngTotalTasks + oSubproject.SourceProject.Tasks.Count
      lngTotalResources = lngTotalResources + oSubproject.SourceProject.ResourceCount
    Next oSubproject
    strMsg = strMsg & "Total: " & Format(lngTotalTasks, "#,##0") & " tasks; " & Format(lngTotalResources, "#,##0") & " resources" & vbCrLf
  Else
    strMsg = strMsg & "Tasks: " & Format(ActiveProject.Tasks.Count, "#,##0") & vbCrLf
    strMsg = strMsg & "Resources: " & Format(ActiveProject.ResourceCount, "#,##0") & vbCrLf
  End If
  strMsg = strMsg & "Baselined: " & IsDate(ActiveProject.BaselineSavedDate(pjBaseline)) & vbCrLf
  blnResourceLoaded = False
  If blnMaster Then
    For Each oSubproject In ActiveProject.Subprojects
      For Each oTask In oSubproject.SourceProject.Tasks
        If oTask Is Nothing Then GoTo next_task_master
        If oTask.Assignments.Count > 0 Then
          blnResourceLoaded = True
          Exit For
        End If
next_task_master:
      Next oTask
    Next oSubproject
  Else
    If ActiveProject.ResourceCount > 0 Then
      For Each oTask In ActiveProject.Tasks
        If oTask Is Nothing Then GoTo next_task_single
        If oTask.Assignments.Count > 0 Then
          blnResourceLoaded = True
          Exit For
        End If
next_task_single:
      Next oTask
    End If
  End If
  strMsg = strMsg & "Resource Loaded: " & blnResourceLoaded & vbCrLf
  blnCPL = False
  If blnMaster Then
    For Each oSubproject In ActiveProject.Subprojects
      For Each oTask In oSubproject.SourceProject.Tasks
        If Not oTask Is Nothing Then
          For Each oLink In oTask.TaskDependencies
            If oLink.To = oTask Then
              If oLink.From.ExternalTask Then
                blnCPL = True
                Exit For
              End If
            End If
            If oLink.From = oTask Then
              If oLink.To.ExternalTask Then
                blnCPL = True
                Exit For
              End If
            End If
          Next oLink
          If blnCPL Then Exit For
        End If
      Next oTask
      If blnCPL Then Exit For
    Next oSubproject
  Else
    For Each oTask In ActiveProject.Tasks
      If Not oTask Is Nothing Then
        For Each oLink In oTask.TaskDependencies
          If oLink.To = oTask Then
            If oLink.From.ExternalTask Then
              blnCPL = True
              Exit For
            End If
          End If
          If oLink.From = oTask Then
            If oLink.To.ExternalTask Then
              blnCPL = True
              Exit For
            End If
          End If
        Next oLink
        If blnCPL Then Exit For
      End If
    Next oTask
  End If
  strMsg = strMsg & "Cross-Project Links: " & blnCPL & vbCrLf
  
  'CPT
  strMsg = strMsg & "--- CPT ---" & vbCrLf 'slightly less likely culprit
  strMsg = strMsg & "Installed to: " & ThisProject.Name & vbCrLf
  strMsg = strMsg & "Error Number: " & strErrNumber & vbCrLf
  strMsg = strMsg & "Error Description: " & strErrDescription & vbCrLf
  strMsg = strMsg & "Error Source: " & strErrSource & vbCrLf
  If Len(strProcessStep) > 0 Then
    strMsg = strMsg & "Process Step: " & strProcessStep & vbCrLf
  End If
  If lngErl > 0 Then
    strMsg = strMsg & "Error Line: " & lngErl & vbCrLf
    'todo: go get the error line...LOL
  End If
  strMsg = strMsg & "Installed: " & strModule & "-" & strInstalled
  If blnBeta Then
    strMsg = strMsg & " -> NOTE: USER IS RUNNING A BETA VERSION" & vbCrLf
  Else
    strMsg = strMsg & vbCrLf
  End If
  strMsg = strMsg & "Latest: " & strModule & "-" & strLatest & vbCrLf
  strMsg = strMsg & "Procedure: " & strProcedure & vbCrLf
  If lngErl > 0 Then
    strMsg = strMsg & "Line: " & lngErl & vbCrLf
  End If
  
  'PROFILE
  strMsg = strMsg & "--- PROFILE ---" & vbCrLf 'even less likely culprit
  Set oProfile = Application.Profiles.ActiveProfile
  With oProfile
    strMsg = strMsg & "Name: " & .Name & vbCrLf
    strMsg = strMsg & "Type: " & Choose(.Type + 1, "pjLocalProfile", "pjServerProfile") & " (" & .Type & ")" & vbCrLf
    strMsg = strMsg & "LoginType: " & Choose(.LoginType, "pjProjectServerLogin", "pjWindowsLogin") & " (" & .LoginType & ")" & vbCrLf
    If .Type = pjServerProfile Then
      strMsg = strMsg & "ConnectionState: " & Choose(.ConnectionState + 1, "pjProfileOffline", "pjProfileOnline") & " (" & .ConnectionState & ")" & vbCrLf
      strMsg = strMsg & "Server: " & .Server & vbCrLf
      strMsg = strMsg & "SiteId: " & .SiteId & vbCrLf
      'strMsg = strMsg & "Username: " & .UserName & vbCrLf 'keep it anonymous, though they'll type in their own name and email...
      strMsg = strMsg & "IsOffline: " & Application.IsOffline & vbCrLf
    End If
  End With
  
  'APPLICATION
  strMsg = strMsg & "--- APPLICATION ---" & vbCrLf 'even less likely culprit
  strMsg = strMsg & "Name: " & Application.Name & " " & Choose(Application.Edition + 1, "Standard", "Professional") & " (" & Application.Build & ")" & vbCrLf
  blnInactive = False
  If Application.Edition = pjEditionProfessional Then
    If Application.IsOffline Then
      strMsg = strMsg & "PWA Status: Offline" & vbCrLf
    Else
      strMsg = strMsg & "PWA Status: Connected" & vbCrLf
    End If
  ElseIf Application.Edition = pjEditionStandard Then 'check for inactive tasks
    If blnMaster Then
      For Each oSubproject In ActiveProject.Subprojects
        For Each oTask In oSubproject.SourceProject.Tasks
          If Not oTask Is Nothing Then
            If Not oTask.Active Then
              blnInactive = True
              Exit For
            End If
          End If
        Next oTask
      Next oSubproject
    Else
      If ActiveProject.ResourceCount > 0 Then
        For Each oTask In ActiveProject.Tasks
          If Not oTask Is Nothing Then
            If Not oTask.Active Then
              blnInactive = True
              Exit For
            End If
          End If
        Next oTask
      End If
    End If
    If blnInactive Then strMsg = strMsg & "WARNING: INACTIVE TASKS FOUND" & vbCrLf
  End If
  strMsg = strMsg & "Calculation: " & IIf(Application.Calculation = pjAutomatic, "On", "Off") & vbCrLf
  #If Win64 Then
    strMsg = strMsg & "Win64: True" & vbCrLf
  #Else
    strMsg = strMsg & "Win64: False" & vbCrLf
  #End If
  #If VBA7 Then
    strMsg = strMsg & "VBA7: True" & vbCrLf
  #Else
    strMsg = strMsg & "VBA7: False" & vbCrLf
  #End If
  strMsg = strMsg & "VBE Version: " & VBE.Version & vbCrLf
  
  'OS
  strMsg = strMsg & "--- OS ---" & vbCrLf 'least likely culprit, but might be releveant
  strKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName" 'Windows 10 Pro
  strMsg = strMsg & "Name: " & oShell.RegRead(strKey)
  If Len(Environ("ProgramW6432")) > 0 Then
    strMsg = strMsg & " (64-bit)" & vbCrLf
  Else
    strMsg = strMsg & " (32-bit)" & vbCrLf
  End If
  strKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DisplayVersion" '22H2
  strMsg = strMsg & "Version: " & oShell.RegRead(strKey) & " (OS Build "
  strKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentMajorVersionNumber" '10
  strMsg = strMsg & oShell.RegRead(strKey) & "."
  strKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentMinorVersionNumber" '0
  strMsg = strMsg & oShell.RegRead(strKey) & "."
  strKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuildNumber" '19045
  strMsg = strMsg & oShell.RegRead(strKey) & "."
  strKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\UBR" '3930
  strMsg = strMsg & oShell.RegRead(strKey) & ")" & vbCrLf
  strKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\EditionID" 'Professional
  strMsg = strMsg & "Edition: " & oShell.RegRead(strKey) & vbCrLf
  
  strMsg = strMsg & String(80, "-") & vbCrLf
  strMsg = strMsg & "=== FINAL STEPS ===" & vbCrLf
  strMsg = strMsg & "-> Enter your name (required)." & vbCrLf
  strMsg = strMsg & "-> Enter your email address (required)." & vbCrLf
  strMsg = strMsg & "-> Enter your phone number (optional). (It is more likely we'll set up a Teams meeting.)" & vbCrLf
  strMsg = strMsg & "-> Enter an alternate email address (optional)." & vbCrLf
  strMsg = strMsg & "-> Pass the [I AM NOT ROBOT] test." & vbCrLf
  strMsg = strMsg & "-> Click [Create Ticket]."
  strMsg = strMsg & "-> Thank you! We will contact you ASAP..."
  lngFile = FreeFile
  strFileName = Environ("tmp") & "\cpt-err-" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"
  Open strFileName For Output As #lngFile
  Print #lngFile, "-> The location of this file is " & strFileName & vbCrLf
  Print #lngFile, strMsg
  Close #lngFile
  ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
  Application.StatusBar = "Opening https://clearplan.happyfox.com/new..."
  Application.FollowHyperlink "https://clearplan.happyfox.com/new/"
  
exit_here:
  On Error Resume Next
  Set oLink = Nothing
  Application.StatusBar = ""
  Set oTask = Nothing
  Set oSubproject = Nothing
  Set oShell = Nothing
  Set oProfile = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSetup_bas", "cptHandleErr", Err, Erl)
  Resume exit_here
End Sub

Function cptIncrement(ByRef lngCleanUp As Long) As Long
  lngCleanUp = lngCleanUp + 1
  cptIncrement = lngCleanUp
End Function

Public Function cptInternetIsConnected() As Boolean

  cptInternetIsConnected = InternetGetConnectedStateEx(0, "", 254, 0)

End Function

Function cptRegEx(strText As String, strRegEx As String, Optional blnMultiLine As Boolean = False, Optional blnIgnoreCase As Boolean = True) As String
  Dim RE As Object, REMatch As Variant, REMatches As Object
  Dim strMatch As String

  On Error GoTo err_here

  Set RE = CreateObject("vbscript.regexp")
  With RE
    .MultiLine = blnMultiLine
    .Global = True
    .IgnoreCase = blnIgnoreCase
    .Pattern = strRegEx
  End With

  Set REMatches = RE.Execute(strText)
  For Each REMatch In REMatches
    strMatch = REMatch
    Exit For
  Next
  cptRegEx = strMatch

exit_here:
  On Error Resume Next
  Set RE = Nothing
  Set REMatches = Nothing
  Exit Function
err_here:
  If Err.Number = 5 Then
    cptRegEx = ""
    Err.Clear
  End If
  Resume exit_here
End Function

Function cptDir() As String
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
  cptDir = strPath

End Function

Function cptModuleExists(strModule As String) As Boolean
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

  cptModuleExists = blnExists

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptSetup_bas", "cptModuleExists", Err, Erl)
  Resume exit_here

End Function

Sub cptUninstall()
  'objects
  Dim vEvent As Object
  Dim Project As Object
  Dim vbComponent As Object
  Dim cmThisProject As Object
  'strings
  Dim strDir As String
  Dim strMsg As String
  'longs
  Dim lngLine As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir

  If MsgBox("Are you sure?", vbCritical + vbYesNo, "Uninstall CPT") = vbNo Then GoTo exit_here

  strMsg = "1. Please delete the module 'cptSetup_bas' manually after this process completes." & vbCrLf & vbCrLf
  strMsg = strMsg & "2. If you have made modifications to the 'ThisProject' module, you may need to review it." & vbCrLf & vbCrLf
  strMsg = strMsg & "3. This process will NOT delete your historical data and saved settings which are stored in '" & strDir & "'. Please do this manually - and fully aware that this cannot be undone." & vbCrLf & vbCrLf
  strMsg = strMsg & "Alternatively, if you would like to reinstall, re-run cptSetup() and then install updates."
  If MsgBox(strMsg, vbInformation + vbOKCancel, "Thank You!") = vbCancel Then GoTo exit_here

  'remove cpt-related lines from ThisProject
  Set cmThisProject = ThisProject.VBProject.VBComponents("ThisProject").CodeModule
  With cmThisProject
    'delete the version
    For lngLine = .CountOfDeclarationLines To 1 Step -1
      If InStr(.Lines(lngLine, 1), "<cpt_version>") > 0 Then
        .DeleteLines lngLine, 1
        DoEvents
      End If
    Next lngLine
    For lngLine = .CountOfLines To 1 Step -1
      'comment out existing lines to avoid immediate errors
      If InStr(.Lines(lngLine, 1), "Sub") > 0 Then
        'do nothing
      ElseIf InStr(.Lines(lngLine, 1), "</cpt") > 0 Then
        If .ProcOfLine(lngLine, 1) = "Project_Activate" Then
          'holding next line in case we decide to comment out instead of delete
          '.ReplaceLine lngLine, "'" & .Lines(lngLine, 1)
          .DeleteLines lngLine, 1
          DoEvents
        ElseIf .ProcOfLine(lngLine, 1) = "Project_Open" Then
          .DeleteLines lngLine, 1
          DoEvents
        End If
      End If
    Next lngLine
  End With

  'reset the toolbar
  ActiveProject.SetCustomUI "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui""><mso:ribbon></mso:ribbon></mso:customUI>"

  'remove all cpt modules
  For Each vbComponent In ThisProject.VBProject.VBComponents
    If Left(vbComponent.Name, 3) = "cpt" And vbComponent.Name <> "cptSetup_bas" Then
      If vbComponent.Name = "cptAdmin_bas" Then GoTo next_component
      Application.StatusBar = "Purging module " & vbComponent.Name & "..."
      If Dir(strDir & "\modules\", vbDirectory) = vbNullString Then MkDir strDir & "\modules"
      vbComponent.Export strDir & "\modules\" & vbComponent.Name
      ThisProject.VBProject.VBComponents.Remove vbComponent
    End If
next_component:
  Next vbComponent

  MsgBox "Thank you for using the ClearPlan Toolbar.", vbInformation + vbOKOnly, "Uninstall Complete"

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set vEvent = Nothing
  Set Project = Nothing
  Set vbComponent = Nothing
  Set cmThisProject = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptSetup_bas", "cptUninstall", Err, Erl)
  Resume exit_here
End Sub

Sub cptDate_dd()
  DefaultDateFormat = pjDate_dd
End Sub
Sub cptDate_ddd_dd()
  DefaultDateFormat = pjDate_ddd_dd
End Sub
Sub cptDate_ddd_hh_mmAM()
  DefaultDateFormat = pjDate_ddd_hh_mmAM
End Sub
Sub cptDate_ddd_mm_dd()
  DefaultDateFormat = pjDate_ddd_mm_dd
End Sub
Sub cptDate_ddd_mm_dd_yy()
  DefaultDateFormat = pjDate_ddd_mm_dd_yy
End Sub
Sub cptDate_ddd_mm_dd_yy_hh_mmAM()
  DefaultDateFormat = pjDate_ddd_mm_dd_yy_hh_mmAM
End Sub
Sub cptDate_ddd_mmm_dd()
  DefaultDateFormat = pjDate_ddd_mmm_dd
End Sub
Sub cptDate_ddd_mmm_dd_yyy()
  DefaultDateFormat = pjDate_ddd_mmm_dd_yyy
End Sub
Sub cptDate_hh_mmAM()
  DefaultDateFormat = pjDate_hh_mmAM
End Sub
Sub cptDate_mm_dd()
  DefaultDateFormat = pjDate_mm_dd
End Sub
Sub cptDate_mm_dd_yy()
  DefaultDateFormat = pjDate_mm_dd_yy
End Sub
Sub cptDate_mm_dd_yy_hh_mmAM()
  DefaultDateFormat = pjDate_mm_dd_yy_hh_mmAM
End Sub
Sub cptDate_mm_dd_yyyy()
  DefaultDateFormat = pjDate_mm_dd_yyyy
End Sub
Sub cptDate_mmm_dd()
  DefaultDateFormat = pjDate_mmm_dd
End Sub
Sub cptDate_mmm_dd_hh_mmAM()
  DefaultDateFormat = pjDate_mmm_dd_hh_mmAM
End Sub
Sub cptDate_mmm_dd_yyy()
  DefaultDateFormat = pjDate_mmm_dd_yyy
End Sub
Sub cptDate_mmmm_dd()
  DefaultDateFormat = pjDate_mmmm_dd
End Sub
Sub cptDate_mmmm_dd_yyyy()
  DefaultDateFormat = pjDate_mmmm_dd_yyyy
End Sub
Sub cptDate_mmmm_dd_yyyy_hh_mmAM()
  DefaultDateFormat = pjDate_mmmm_dd_yyyy_hh_mmAM
End Sub
Sub cptDate_Www_dd()
  DefaultDateFormat = pjDate_Www_dd
End Sub
Sub cptDate_Www_dd_yy_hh_mmAM()
  DefaultDateFormat = pjDate_Www_dd_yy_hh_mmAM
End Sub

Sub cptValidateXML(strXML As String)
  'objects
  Dim oXML As MSXML2.DOMDocument30
  'strings
  Dim strFileName As String
  'longs
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  strFileName = Environ("tmp") & "\cpt-validate.xml"
  lngFile = FreeFile
  Open strFileName For Output As #lngFile
  Print #lngFile, strXML
  Close #lngFile
  
  Set oXML = New MSXML2.DOMDocument30
  If oXML.Load(strFileName) Then
    MsgBox "cpt ribbon xml validated", vbInformation + vbOKOnly, "success"
  Else
    MsgBox "cpt ribbon xml validation failed", vbCritical + vbOKOnly, "failure"
    If oXML.parseError.errorcode <> 0 Then
      MsgBox oXML.parseError.reason, vbInformation + vbOKOnly, "reason:"
    End If
  End If

  Kill strFileName

exit_here:
  On Error Resume Next
  Set oXML = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSetup_bas", "cptValidateXML", Err, Erl)
  Resume exit_here
End Sub
