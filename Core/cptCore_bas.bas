Attribute VB_Name = "cptCore_bas"
'<cpt_version>v1.16.0</cpt_version>
Option Explicit
Private oMSPEvents As cptEvents_cls
#If Win64 And VBA7 Then
  Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  Private Declare PtrSafe Function SetPrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
  Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr '<issue53>
  Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
  Private Declare Function GetPrivateProfileString lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  Private Declare Function SetPrivateProfileString lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
  Public Declare Function GetTickCount Lib "kernel32" () As Long '<issue53>
  Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Sub cptStartEvents()
  Set oMSPEvents = New cptEvents_cls
End Sub

Sub cptStopEvents()
  Set oMSPEvents = Nothing
End Sub

Sub cptSpeed(blnOn As Boolean)

  Application.Calculation = pjAutomatic = Not blnOn
  Application.ScreenUpdating = Not blnOn

End Sub

Function cptGetUserForm(strModuleName As String) As MSForms.UserForm
  'NOTE: this only works if the form is loaded
  'objects
  Dim oUserForm As Object
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For Each oUserForm In VBA.UserForms
    If oUserForm.Name = strModuleName Then
      Set cptGetUserForm = oUserForm
      Exit For
    End If
  Next

exit_here:
  On Error Resume Next
  Set oUserForm = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetUserForm", Err, Erl)
  Resume exit_here
End Function

Function cptGetControl(ByRef cptForm_frm As MSForms.UserForm, strControlName As String) As MSForms.Control
  'NOTE: this only works for loaded forms

  Set cptGetControl = cptForm_frm.Controls(strControlName)

End Function

Function cptGetUserFullName()
  'used to add user's name to PowerPoint title slide
  Dim objAllNames As Object, objIndName As Object

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  On Error Resume Next
  Set objAllNames = GetObject("Winmgmts:").instancesof("win32_networkloginprofile")
  For Each objIndName In objAllNames
    If Len(objIndName.FullName) > 0 Then
      cptGetUserFullName = objIndName.FullName
      Exit For
    End If
  Next

exit_here:
  On Error Resume Next
  Set objAllNames = Nothing
  Set objIndName = Nothing
  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetUserFullName", Err, Erl)
  Resume exit_here

End Function

Function cptGetBreadcrumbs(strModule As String, strProcedure As String, strBreadcrumb As String) As Variant
  'usage: cptGetBreadcrumbs("cptStatusSheet_bas","cptCopyData","<cpt-breadcrumbs:format-conditions>")
  Dim vbComponent As Object 'vbComponent
  Dim strResult As String
  Dim vbCodeModule As Object 'CodeModule
  Dim lngStart As Long
  Dim lngCount As Long
  Dim lngLine As Long
  Dim blnResult As Boolean
  Dim strLine As String
  
  If Not cptModuleExists(strModule) Then
    cptGetBreadcrumbs = "...module '" & strModule & "' not found"
    Exit Function
  Else
    Set vbComponent = ThisProject.VBProject.VBComponents(strModule)
    Set vbCodeModule = vbComponent.CodeModule
    If Not vbCodeModule.Find(strProcedure, 1, 1, vbCodeModule.CountOfLines, 100, True) Then
      cptGetBreadcrumbs = " ...procedure '" & strProcedure & "' not found"
      Exit Function
    End If
    lngStart = vbCodeModule.ProcBodyLine(strProcedure, 0) '0=vbext_pk_Proc
    lngCount = lngStart + vbCodeModule.ProcCountLines(strProcedure, 0) '0=vbext_pk_Proc
    If vbCodeModule.Find("<cpt-breadcrumbs:" & strBreadcrumb & ">", lngStart, 1, lngStart + lngCount, 100) = True Then
      For lngLine = lngStart To (lngStart + lngCount)
        If vbCodeModule.Find("<cpt-breadcrumbs:" & strBreadcrumb & ">", lngLine, 1, lngLine, 100) = True Then
          blnResult = True 'start capturing
        ElseIf vbCodeModule.Find("</cpt-breadcrumbs:" & strBreadcrumb & ">", lngLine, 1, lngLine, 100) = True Then
          Exit For 'stop capturing
        Else
          If blnResult Then
            strLine = Trim(vbCodeModule.Lines(lngLine, 1))
            If Left(strLine, 1) = "'" Then
              If InStr(strLine, "todo") = 0 And InStr(strLine, "</skip-breadcrumb>") = 0 Then
                strResult = strResult & Right(strLine, Len(strLine) - 1) & vbCrLf 'comments only, sans apostrophe
              End If
            End If
          End If
        End If
      Next lngLine
      cptGetBreadcrumbs = Left(strResult, Len(strResult) - 1) 'hack off trailing comma
    End If
  End If
End Function

Function cptGetVersions() As String
  'requires reference: Microsoft Scripting Runtime
  Dim vbComponent As Object, strVersion As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For Each vbComponent In ThisProject.VBProject.VBComponents
    'is the vbComponent one of ours?
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = cptRegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
      cptGetVersions = cptGetVersions & vbComponent.Name & ": " & strVersion & vbCrLf
    End If
  Next vbComponent

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetVersions", Err, Erl)
  Resume exit_here

End Function

'<issue31>
Sub cptUpgrade(Optional strFileName As String)
  'objects
  Dim oStream As Object
  Dim xmlHttpDoc As Object
  'strings
  Dim strDir As String
  Dim strNewFileName As String
  Dim strModule As String
  Dim strError As String
  Dim strURL As String
  'longs
  Dim lngLine As Long
  'integers
  'doubles
  'booleans
  Dim blnExists As Boolean
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  If Len(strFileName) = 0 Then strFileName = "Core/cptUpgrades_frm.frm"

  'go get it
  strURL = strGitHub
  strURL = strURL & strFileName
  strFileName = Replace(cptRegEx(strFileName, "\/.*\.[A-z]{3}"), "/", "")
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
  End If

  'remove if exists
  strModule = Left(strFileName, InStr(strFileName, ".") - 1)
  blnExists = cptModuleExists(strModule)
  If blnExists Then
    'Set vbComponent = ThisProject.VBProject.VBComponents("cptUpgrades_frm")
    Application.StatusBar = "Removing obsolete version of " & strModule
    strNewFileName = strModule & "_" & Format(Now, "hhnnss")
    ThisProject.VBProject.VBComponents(strModule).Name = strNewFileName
    DoEvents
    ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents(strNewFileName)
    cptCore_bas.cptStartEvents
    DoEvents
  End If

  'import the module
  Application.StatusBar = "Importing " & strFileName & "..."
  ThisProject.VBProject.VBComponents.Import strDir & "\" & strFileName
  DoEvents
  
  '<issue24> remove the whitespace added by VBE import/export
  With ThisProject.VBProject.VBComponents(strModule).CodeModule
    For lngLine = .CountOfDeclarationLines To 1 Step -1
      If Len(.Lines(lngLine, 1)) = 0 Then .DeleteLines lngLine, 1
    Next lngLine
  End With '</issue24>

  'MsgBox "The Upgrade Form was itself just upgraded. Please repeat your click.", vbInformation + vbOKOnly, "Upgraded the Upgrader"

exit_here:
  On Error Resume Next
  Set oStream = Nothing
  Set xmlHttpDoc = Nothing
  Application.StatusBar = ""
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptUpgrade", Err, Erl)
  Resume exit_here

End Sub '<issue31>

Sub cptShowAbout_frm()
  'objects
  Dim myAbout_frm As cptAbout_frm
  'strings
  Dim strAbout As String
  'longs
  'integers
  'booleans
  'variants
  'dates
  
  'prevent spawning
  If Not cptGetUserForm("cptAbout_frm") Is Nothing Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not cptModuleExists("cptAbout_frm") Then '<issue19>
    MsgBox "Please re-run cptSetup() to restore the missing About form.", vbCritical + vbOKOnly, "Missing Form"
    GoTo exit_here
  End If '</issue19>

  'contact and license
  strAbout = vbCrLf & "The ClearPlan Toolbar" & vbCrLf
  strAbout = strAbout & "by ClearPlan Consulting, LLC" & vbCrLf & vbCrLf
  strAbout = strAbout & "This software is provided free of charge," & vbCrLf
  strAbout = strAbout & "AS IS and without warranty." & vbCrLf
  strAbout = strAbout & "It is free to use, free to distribute with prior written consent from the developers/copyright holders and without modification." & vbCrLf & vbCrLf
  strAbout = strAbout & "All rights reserved." & vbCrLf & "Copyright " & Chr(169) & " " & Year(Now()) & ", ClearPlan Consulting, LLC"
  Set myAbout_frm = New cptAbout_frm
  myAbout_frm.txtAbout.Value = strAbout  '<issue19>

  'follow the project
  strAbout = vbCrLf & vbCrLf & "Follow the Project:" & vbCrLf & vbCrLf
  strAbout = strAbout & "https://github.com/clearplan/cpt" & vbCrLf & vbCrLf
  myAbout_frm.txtGitHub.Value = strAbout '<issue19>

  'show/hide
  'myAbout_frm.lblScoreBoard.Visible = IIf(Now <= #10/25/2019#, False, True) '<issue19>
  'myAbout_frm.lblScoreBoard.Caption = "t0 : b1" EWR > MSY '3/22/19 = 1
  'myAbout_frm.lblScoreBoard.Caption = "t0 : b2" MSY > EWR '3/24/19 = 2
  'myAbout_frm.lblScoreBoard.Caption = "t0 : b3" 'EWR > SAN '10/25/19 = 3
  'myAbout_frm.lblScoreBoard.Caption = "t0 : b4" 'SAN > EWR '10/27/19 = 4
  'myAbout_frm.lblScoreBoard.Caption = "t0 : b5" 'EWR > NAS '2/17/20 = 5
  'myAbout_frm.lblScoreBoard.Caption = "t0 : b6" 'NAS > EWR '2/20/20 = 6
  'myAbout_frm.lblScoreBoard.Caption = "t0 : b7" 'EWR > SAV '6/3/22 = 7
  'myAbout_frm.lblScoreBoard.Caption = "t0 : b8" 'EWR > SAV '6/5/22 = 8
  'myAbout_frm.lblScoreBoard.Caption = "t0 : b9" 'EWR > DFW '5/16/25 = 9     v1.9.0
  myAbout_frm.lblScoreBoard.Caption = "t0 : b10" 'DFW > EWR '5/18/25 = 10  v1.9.1
  
  myAbout_frm.Caption = "The ClearPlan Toolbar - " & cptGetVersion("cptAbout_frm")
  myAbout_frm.Show '<issue19>

exit_here:
  On Error Resume Next
  Unload myAbout_frm
  Set myAbout_frm = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptShowAbout_frm", Err, Erl)
  Resume exit_here '</issue19>

End Sub

Public Sub cptQuickSort(vArray As Variant, inLow As Long, inHi As Long)
  'public domain: https://stackoverflow.com/questions/152319/vba-array-sort-function
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then cptQuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then cptQuickSort vArray, tmpLow, inHi
End Sub

Function cptReferenceExists(strReference As String) As Boolean
  'used to ensure a reference exists, returns boolean
  Dim Ref As Object, blnExists As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  blnExists = False

  For Each Ref In ThisProject.VBProject.References
    If Ref.Name = strReference Then
      blnExists = True
      Exit For
    End If
  Next Ref

  cptReferenceExists = blnExists

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptReferenceExists", Err, Erl)
  Resume exit_here
End Function

Sub cptGetReferences()
  'prints the current uesr's selected references
  'this would be used to troubleshoot with users real-time
  'although simply running cptSetReferences should fix it
  Dim oRef As Object 'Reference
  Dim lngFile As Long
  Dim strFileName As String
  Dim strRef As String
  Dim lngRefs As Long
  Dim lngRef As Long
  
  lngFile = FreeFile
  strFileName = Environ("tmp") & "\cpt-references.csv"
  Open strFileName For Output As #lngFile
  
  Print #lngFile, "NAME,DESCRIPTION,FULL_PATH,GUID,MAJOR,MINOR,BUILT_IN,IS_BROKEN,TYPE,"
  lngRefs = ThisProject.VBProject.References.Count
  lngRef = 0
  For Each oRef In ThisProject.VBProject.References
    lngRef = lngRef + 1
    Debug.Print lngRef & "/" & lngRefs & " " & String(25, "=")
    Debug.Print oRef.Name
    Debug.Print "-- " & oRef.Description
    Debug.Print "-- " & oRef.FullPath
    Debug.Print "-- " & oRef.Guid & " | " & oRef.Major & " | " & oRef.Minor
    Debug.Print "-- BuiltIn: " & oRef.BuiltIn
    Debug.Print "-- IsBroken: " & oRef.IsBroken
    Debug.Print "-- Type: " & oRef.Type
    strRef = Join(Array(oRef.Name, oRef.Description, oRef.FullPath, oRef.Guid, oRef.Major, oRef.Minor, oRef.BuiltIn, oRef.IsBroken, oRef.Type), ",")
    Print #lngFile, strRef & ","
  Next oRef
  Reset
  
  ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
  
End Sub

Function cptGetDirectory(strModule As String) As String
  'this function retrieves the directory of the module from CurrentVersions.xml on gitHub
  'objects
  Dim xmlDoc As Object
  Dim xmlNode As Object
  'strings
  Dim strDirectory As String
  Dim strURL As String
  'longs
  'integers
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'the calling subroutine should catch the Not cptInternetIsConnected function before calling this

  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.SetProperty "SelectionLanguage", "XPath"
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
  strURL = strGitHub & "CurrentVersions.xml"
  If Not xmlDoc.Load(strURL) Then
    MsgBox xmlDoc.parseError.errorcode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
  Else
    Set xmlNode = xmlDoc.SelectSingleNode("//Name[text()='" + strModule + "']").ParentNode.SelectSingleNode("Directory")
    strDirectory = xmlNode.Text
  End If

  cptGetDirectory = strDirectory

exit_here:
  On Error Resume Next
  Set xmlDoc = Nothing
  Set xmlNode = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetDirectory()", Err, Erl)
  Resume exit_here
End Function

Sub cptGetEnviron()
  'list the environment variables and their associated values
  Dim lngIndex As Long
  Dim strFileName As String
  Dim lngFile As Long
  
  strFileName = Environ("tmp") & "\current_environment.txt"
  lngFile = FreeFile
  Open strFileName For Output As #lngFile

  For lngIndex = 1 To 200
    Print #lngFile, lngIndex & ": " & Environ(lngIndex)
  Next
  Close #lngFile
  Reset
  ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
  
End Sub
Function cptCheckReference(strReference As String) As Boolean
  'this routine will be called ahead of any subroutine requiring a reference
  'returns boolean and subroutine only proceeds if true
  Dim strDir As String
  Dim strRegEx As String

  On Error GoTo err_here

  cptCheckReference = True

  Select Case strReference
    'CommonProgramFiles
    Case "Office"
      If Not cptReferenceExists("Office") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\Microsoft Shared\OFFICE16\MSO.DLL"
      End If
    Case "VBIDE"
      If Not cptReferenceExists("VBIDE") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
      End If
    Case "VBA"
      If Not cptReferenceExists("VBA") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
      End If
    Case "ADODB"
      If Not cptReferenceExists("ADODB") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\System\ado\msado15.dll"
      End If
  End Select
    
  'Office Applications
  strDir = cptRegEx(Environ("PATH"), "C:\\[^;]*Office[0-9]{1,}\\")
  If Len(strDir) = 0 Then GoTo windows_common
  Select Case strReference
    Case "Excel"
      If Not cptReferenceExists("Excel") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\EXCEL.EXE"
      End If
    Case "Outlook"
      If Not cptReferenceExists("Outlook") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\MSOUTL.OLB"
      End If
    Case "PowerPoint"
      If Not cptReferenceExists("PowerPoint") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\MSPPT.OLB"
      End If
    Case "MSProject"
      If Not cptReferenceExists("MSProject") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\MSPRJ.OLB"
      End If
    Case "Word"
      If Not cptReferenceExists("Word") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\MSWORD.OLB"
      End If

    'Windows Common
windows_common:
    Case "MSForms"
      If Not cptReferenceExists("MSForms") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\FM20.DLL"
      End If
    Case "Scripting"
      If Not cptReferenceExists("Scripting") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\scrrun.dll"
      End If
    Case "stdole"
      If Not cptReferenceExists("stdole") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\stdole2.tlb"
      End If
    Case "mscorlib"
      If Not cptReferenceExists("mscorlib") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb"
      End If
    Case "MSXML2"
      If Not cptReferenceExists("MSXML2") Then '</issue33>
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\msxml3.dll"
      End If
    Case "MSComctlLib"
      If Not cptReferenceExists("MSComctlLib") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\MSCOMCTL.OCX"
      End If
    Case Else
      cptCheckReference = False

  End Select

  If Not cptCheckReference Then
    MsgBox "Missing Reference: " & strReference, vbExclamation + vbOKOnly, "CP Tool Bar"
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  cptCheckReference = False
  Resume exit_here

End Function

Sub cptResetAll()
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  Dim rstSettings As Object 'ADODB.Recordset
  'strings
  Dim strDefaultView As String
  Dim strFilter As String
  Dim strOutlineLevel As String
  Dim strSettings As String
  Dim strFileName As String
  Dim strOutlineParents As String
  'longs
  Dim lngUID As Long
  Dim lngSettings As Long
  Dim lngOutlineLevel As Long
  Dim lngLevel As Long
  'integers
  'doubles
  'booleans
  Dim blnFilter As Boolean
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vOutlineParent As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Reset All"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===

  cptSpeed True
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If Not oTasks Is Nothing Then
    If oTasks.Count > 0 Then
      lngUID = oTasks(1).UniqueID
    End If
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strFileName = cptDir & "\settings\cpt-reset-all.adtg"
  If Dir(strFileName) <> vbNullString Then
    Set rstSettings = CreateObject("ADODB.Recordset")
    rstSettings.Open strFileName
    rstSettings.MoveFirst
    lngSettings = rstSettings(0)
    cptSaveSetting "ResetAll", "Settings", CStr(lngSettings)
    lngOutlineLevel = rstSettings(1)
    cptSaveSetting "ResetAll", "OutlineLevel", CStr(lngOutlineLevel)
    rstSettings.Close
    Kill strFileName
  Else
    strSettings = cptGetSetting("ResetAll", "Settings")
    If Len(strSettings) > 0 Then lngSettings = CLng(strSettings)
    strOutlineLevel = cptGetSetting("ResetAll", "OutlineLevel")
    If Len(strOutlineLevel) > 0 Then lngOutlineLevel = CLng(strOutlineLevel)
  End If
  
  strDefaultView = cptGetSetting("ResetAll", "DefaultView")
  If Len(strDefaultView) > 0 And cptViewExists(strDefaultView) Then
    ActiveWindow.TopPane.Activate
    ViewApply strDefaultView
    SetSplitBar ShowColumns:=ActiveProject.TaskTables(ActiveProject.CurrentTable).TableFields.Count
  End If
  
  If lngSettings > 0 Then
    'parse and apply
    If lngSettings >= 128 Then 'outline symbols
      OptionsViewEx DisplayOutlineSymbols:=True
      lngSettings = lngSettings - 128
    End If
    If lngSettings >= 64 Then 'display name indent
      OptionsViewEx DisplayNameIndent:=True
      lngSettings = lngSettings - 64
    End If
    If lngSettings >= 32 Then 'clear filter
      blnFilter = True
      FilterClear
      lngSettings = lngSettings - 32
    Else
      blnFilter = False
    End If
    If lngSettings >= 16 Then 'sort by ID
      Sort "ID", , , , , , False, True
      lngSettings = lngSettings - 16
    End If
    If lngSettings >= 8 Then 'expand all tasks
      OptionsViewEx DisplaySummaryTasks:=True
      On Error Resume Next
      If Not OutlineShowAllTasks Then
        If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
          Sort "ID", , , , , , False, True
          OutlineShowAllTasks
        Else
          SelectBeginning
          GoTo exit_here
        End If
      End If
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If ActiveProject.Subprojects.Count > 0 Then
        OptionsViewEx DisplaySummaryTasks:=True
        If Not blnFilter Then
          strFilter = ActiveProject.CurrentFilter
        End If
        FilterClear
        SelectAll
        OutlineShowAllTasks
        If Len(strFilter) > 0 Then FilterApply strFilter
      End If
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      lngSettings = lngSettings - 8
    Else 'expand to specific level
      OptionsViewEx DisplaySummaryTasks:=True
      On Error Resume Next
      If Not OutlineShowAllTasks Then
        If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
          Sort "ID", , , , , , False, True
          OutlineShowAllTasks
        Else
          SelectBeginning
          GoTo exit_here
        End If
      End If
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      OutlineShowTasks pjTaskOutlineShowLevelMax
      For lngLevel = 20 To lngOutlineLevel Step -1
        OutlineShowTasks lngLevel
      Next lngLevel
    End If
    If lngSettings >= 4 Then 'show summaries
      OptionsViewEx DisplaySummaryTasks:=True
      lngSettings = lngSettings - 4
    Else
      OptionsViewEx DisplaySummaryTasks:=False
    End If
    If lngSettings >= 2 Then 'clear group
      GroupClear
      lngSettings = lngSettings - 2
    End If
    If lngSettings >= 1 Then 'hide inactive
      If Edition = pjEditionProfessional Then SetAutoFilter "Active", pjAutoFilterFlagYes
    End If
  Else 'prompt for defaults
    Call cptShowResetAll_frm
  End If
  
  strSettings = cptGetSetting("ResetAll", "KeepPosition")
  If Len(strSettings) = 0 Then
    strSettings = 0
  End If
  If lngUID > 0 And CBool(strSettings) Then
    cptSpeed False
    On Error Resume Next
    If Not EditGoTo(ActiveProject.Tasks.UniqueID(lngUID).ID) Then
      If MsgBox("Return to UID " & lngUID & "?", vbQuestion + vbYesNo, "Keep Position?") = vbYes Then
        strOutlineParents = cptGetOutlineParents(ActiveProject.Tasks.UniqueID(lngUID))
        If Len(strOutlineParents) > 0 Then
          For Each vOutlineParent In Split(strOutlineParents, ",")
            EditGoTo ActiveProject.Tasks.UniqueID(vOutlineParent).ID
            Application.OutlineShowSubTasks
          Next
        End If
        EditGoTo ActiveProject.Tasks.UniqueID(lngUID).ID
      End If
    End If
  Else
    SelectBeginning
  End If

exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  If rstSettings.State Then rstSettings.Close
  Set rstSettings = Nothing
  cptSpeed False

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptResetAll", Err, Erl)
  Resume exit_here

End Sub

Function cptGetOutlineParents(oTask As MSProject.Task) As String
  'objects
  Dim oCurrentTask As MSProject.Task
  Dim oOutlineParents As Collection
  'strings
  Dim strOutlineParents As String
  'longs
  Dim lngCount As Long
    
  Set oOutlineParents = New Collection
  Set oCurrentTask = oTask
  Do While oCurrentTask.OutlineParent.OutlineLevel > 0
    Set oCurrentTask = oCurrentTask.OutlineParent
    oOutlineParents.Add oCurrentTask.UniqueID
  Loop
  For lngCount = oOutlineParents.Count To 1 Step -1
    If lngCount = oOutlineParents.Count Then
      strOutlineParents = CStr(oOutlineParents(lngCount))
    Else
      strOutlineParents = strOutlineParents & "," & CStr(oOutlineParents(lngCount))
    End If
  Next lngCount
  cptGetOutlineParents = strOutlineParents
End Function

Sub cptShowResetAll_frm()
  'objects
  Dim myResetAll_frm As cptResetAll_frm
  Dim oView As MSProject.View
  Dim rstSettings As Object 'ADODB.Recordset
  'strings
  Dim strKeepPosition As String
  Dim strViewList As String
  Dim strDefaultView As String
  Dim strOutlineLevel As String
  Dim strSettings As String
  Dim strFileName As String
  'longs
  Dim lngSettings As Long
  Dim lngOutlineLevel As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vViewList As Variant
  'dates
  
  'prevent spawning
  If Not cptGetUserForm("cptResetAll_frm") Is Nothing Then Exit Sub
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Reset All"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  'instantiate the form
  Set myResetAll_frm = New cptResetAll_frm
  'populate the outline level picklist
  For lngOutlineLevel = 1 To 20
    myResetAll_frm.cboOutlineLevel.AddItem lngOutlineLevel
  Next lngOutlineLevel
  'default to 2
  myResetAll_frm.cboOutlineLevel.Value = 2
  
  'populate cboViews
  myResetAll_frm.cboViews.Clear
  For Each oView In ActiveProject.Views
    If oView.Type = pjTaskItem Then
      If oView.Screen = 1 Or oView.Screen = 14 Then
        strViewList = strViewList & oView.Name & ","
'        Debug.Print oView.Type & vbTab & Choose(oView.Type + 1, "pjTaskItem", "pjResourceItem", "pjOtherItem") & vbTab & oView.Name
'        Debug.Print oView.Screen & vbTab & Choose(oView.Screen, "pjGantt", "pjNetworkDiagram", "pjRelationshipDiagram", "pjTaskForm", "pjTaskSheet", "pjResourceForm", "pjResourceSheet", "pjResourceGraph", "pjRSVDoNotUse", "pjTaskDetailsForm", "pjTaskNameForm", "pjResourceNameForm", "pjCalendar", "pjTaskUsage", "pjResourceUsage", "pjTimeline", "pjResourceScheduling")
      End If
    End If
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  Next oView
  Set oView = Nothing
  strViewList = Left(strViewList, Len(strViewList) - 1)
  
  vViewList = Split(strViewList, ",")
  cptQuickSort vViewList, 0, UBound(vViewList)
  myResetAll_frm.cboViews.List = Split("<None>," & Join(vViewList, ","), ",")
  
  strFileName = cptDir & "\settings\cpt-reset-all.adtg"
  If Dir(strFileName) <> vbNullString Then
    'get saved settings
    Set rstSettings = CreateObject("ADODB.Recordset")
    rstSettings.Open strFileName
    rstSettings.MoveFirst
    lngSettings = rstSettings(0)
    cptSaveSetting "ResetAll", "Settings", CStr(lngSettings)
    lngOutlineLevel = rstSettings(1)
    cptSaveSetting "ResetAll", "OutlineLevel", CStr(lngOutlineLevel)
    rstSettings.Close
    Kill strFileName
  Else
    strSettings = cptGetSetting("ResetAll", "Settings")
    If Len(strSettings) > 0 Then lngSettings = CLng(strSettings)
    strOutlineLevel = cptGetSetting("ResetAll", "OutlineLevel")
    If Len(strOutlineLevel) > 0 Then lngOutlineLevel = CLng(strOutlineLevel)
    strDefaultView = cptGetSetting("ResetAll", "DefaultView")
    If Len(strDefaultView) > 0 Then
      If strDefaultView <> "<None>" Then
        If Not cptViewExists(strDefaultView) Then
          MsgBox "Your default view '" & strDefaultView & "' does not exist.", vbExclamation + vbOKOnly, "Saved View Not Found"
        Else
          myResetAll_frm.cboViews.Value = strDefaultView
        End If
      Else
        myResetAll_frm.cboViews.Value = "<None>"
      End If
    Else
      cptSaveSetting "ResetAll", "DefaultView", "<None>"
      myResetAll_frm.cboViews.Value = "<None>"
    End If
  End If
    
  If lngSettings > 0 Then
    'parse and update the form
    With myResetAll_frm
      If lngSettings >= 128 Then
        .chkOutlineSymbols = True
        lngSettings = lngSettings - 128
      End If
      If lngSettings >= 64 Then
        .chkIndent = True
        lngSettings = lngSettings - 64
      End If
      If lngSettings >= 32 Then
        .chkFilter = True
        lngSettings = lngSettings - 32
      End If
      If lngSettings >= 16 Then
        .chkSort = True
        lngSettings = lngSettings - 16
      End If
      If lngSettings >= 8 Then
        .optShowAllTasks = True
        lngSettings = lngSettings - 8
      Else
        .optOutlineLevel = True
        .cboOutlineLevel.Value = IIf(lngOutlineLevel = 0, 2, lngOutlineLevel)
      End If
      If lngSettings >= 5 Then
        .chkSummaries = True
        lngSettings = lngSettings - 4
      End If
      If lngSettings >= 2 Then
        .chkGroup = True
        lngSettings = lngSettings - 2
      End If
      If Edition = pjEditionProfessional Then
        .chkActiveOnly.Enabled = True
        If lngSettings >= 1 Then
          .chkActiveOnly = True
        End If
      ElseIf Edition = pjEditionStandard Then
        .chkActiveOnly = False
        .chkActiveOnly.Enabled = False
      End If
    End With
  End If
  strKeepPosition = cptGetSetting("ResetAll", "KeepPosition")
  If Len(strKeepPosition) > 0 Then
    myResetAll_frm.chkKeepPosition = CBool(strKeepPosition)
  Else
    myResetAll_frm.chkKeepPosition.Value = False
  End If
  myResetAll_frm.Caption = "How would you like to Reset All? (" & cptGetVersion("cptResetAll_frm") & ")"
  myResetAll_frm.Show False
  
exit_here:
  On Error Resume Next
  Set oView = Nothing
  Set rstSettings = Nothing
  Set myResetAll_frm = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptShowResetAll_frm", Err, Erl)
  Resume exit_here

End Sub

Sub cptShowUpgrades_frm()
  'objects
  Dim oUserForm As Object
  Dim myUpgrades_frm As cptUpgrades_frm
  Dim REMatch As Object
  Dim REMatches As Object
  Dim RE As Object
  Dim oStream As Object
  Dim xmlHttpDoc As Object
  Dim rstStatus As Object 'ADODB.Recordset
  Dim vbComponent As Object
  Dim xmlDoc As Object
  Dim xmlNode As Object
  Dim FindRecord As Object
  'long
  Dim lngItem As Long
  'strings
  Dim strOpenForms As String
  Dim strBranch As String
  Dim strFileName As String
  Dim strInstVer As String
  Dim strCurVer As String
  Dim strURL As String
  Dim strVersion As String
  'booleans
  Dim blnOpenForms As Boolean
  Dim blnUpdatesAreAvailable As Boolean
  'variants
  Dim vCol As Variant
  
  'prevent spawning
  If Not cptGetUserForm("cptUpgrades_frm") Is Nothing Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    
  'ensure all cpt forms are closed/unloaded
  blnOpenForms = False
  For Each oUserForm In VBA.UserForms
    If Left(oUserForm.Name, 3) = "cpt" Then
      blnOpenForms = True
      strOpenForms = strOpenForms & "> " & oUserForm.Caption & vbCrLf
    End If
  Next oUserForm
  If blnOpenForms Then
    If MsgBox("We need to close all open cpt forms before we can proceed:" & vbCrLf & vbCrLf & strOpenForms & vbCrLf & "Proceed?", vbQuestion + vbYesNo, "Upgrade CPT") = vbNo Then
      MsgBox "Upgrade cancelled.", vbExclamation + vbOKOnly, "Upgrade CPT"
      GoTo exit_here
    End If
    For Each oUserForm In VBA.UserForms
      If Left(oUserForm.Name, 3) = "cpt" Then
        Unload oUserForm
      End If
    Next oUserForm
  End If
    
  'update references needed before downloading updates
  Application.StatusBar = "Updating VBA references..."
  DoEvents
  Call cptSetReferences
    
  'todo: user should still be able to check currently installed versions
  Application.StatusBar = "Confirming access to GitHub.com..."
  DoEvents
  If Not cptInternetIsConnected Then
    MsgBox "You must be connected to the internet to perform updates.", vbInformation + vbOKOnly, "No Connection"
    GoTo exit_here
  End If

  'set up the recordset
  Set rstStatus = CreateObject("ADODB.Recordset")
  rstStatus.Fields.Append "Module", 200, 200
  rstStatus.Fields.Append "Directory", 200, 200
  rstStatus.Fields.Append "Current", 200, 200
  rstStatus.Fields.Append "Installed", 200, 200
  rstStatus.Fields.Append "Status", 200, 200
  rstStatus.Open
  
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
    MsgBox xmlDoc.parseError.errorcode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
    GoTo exit_here
  Else
    For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
      rstStatus.AddNew
      rstStatus(0) = xmlNode.SelectSingleNode("Name").Text
      rstStatus(1) = xmlNode.SelectSingleNode("Directory").Text
      rstStatus(2) = xmlNode.SelectSingleNode("Version").Text
      rstStatus.Update
    Next xmlNode
  End If

  'get installed versions
  Application.StatusBar = "Comparing installed versions..."
  DoEvents
  blnUpdatesAreAvailable = False
  For Each vbComponent In ThisProject.VBProject.VBComponents
    'is the vbComponent one of ours?
    If vbComponent.CodeModule.Find("'<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = cptRegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>", True)
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

  'if cptUpgrade_frm is updated, install it automatically

  rstStatus.MoveFirst
  rstStatus.Find "Module='cptUpgrades_frm'", , 1
  If cptVersionStatus(rstStatus(3), rstStatus(2)) <> "ok" Then
    Application.StatusBar = "Automatically updating cptUpgrades_frm..."
    DoEvents
    Call cptUpgrade(rstStatus(1) & "/cptUpgrades_frm.frm")
    rstStatus(3) = rstStatus(2)
    rstStatus.Update
  End If
  
  'cannot auto upgrade cptCore_bas because this is the cptCore_bas module so use cptPatch_bas
  'if cptPatch_bas is updated, install it automatically and run it
  rstStatus.MoveFirst
  rstStatus.Find "Module='cptPatch_bas'", , 1
  If cptVersionStatus(rstStatus(3), rstStatus(2)) <> "ok" Then
    Call cptUpgrade(rstStatus(1) & "/cptPatch_bas.bas")
    rstStatus(3) = rstStatus(2)
    rstStatus.Update
    
    '/=== temp fix to cptPatch_bas private/public issue ===\
    'patch code goes here
    Application.StatusBar = "Applying patch 21.04.10..."
    If Not cptReferenceExists("VBScript_RegExp_55") Then
      ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\System32\vbscript.dll\3"
    End If
    '\=== temp fix to cptPatch_bas private/public issue ===/
    
  End If



  'populate the listbox header
  Application.StatusBar = "Preparing form..."
  DoEvents
  lngItem = 0
  Set myUpgrades_frm = New cptUpgrades_frm
  myUpgrades_frm.lboHeader.AddItem
  For Each vCol In Array("Module", "Directory", "Current", "Installed", "Status", "Type")
    myUpgrades_frm.lboHeader.List(0, lngItem) = vCol
    lngItem = lngItem + 1
  Next vCol
  myUpgrades_frm.lboHeader.Height = 16

  'populate the listbox
  myUpgrades_frm.lboModules.Clear
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
    myUpgrades_frm.lboModules.AddItem
    myUpgrades_frm.lboModules.List(lngItem, 0) = rstStatus(0) 'module name
    myUpgrades_frm.lboModules.List(lngItem, 1) = rstStatus(1) 'directory
    myUpgrades_frm.lboModules.List(lngItem, 2) = strCurVer 'current version
    myUpgrades_frm.lboModules.List(lngItem, 3) = strInstVer 'installed version
    
    Select Case strInstVer
      Case Is = strCurVer
        myUpgrades_frm.lboModules.List(lngItem, 4) = "< ok >"
      Case Is = "< missing >"
        myUpgrades_frm.lboModules.List(lngItem, 4) = "< install >"
      Case Is <> strCurVer
        myUpgrades_frm.lboModules.List(lngItem, 4) = "< " & cptVersionStatus(strInstVer, strCurVer) & " >"
    End Select
    'capture the type while we're at it - could have just pulled the FileName
    Set FindRecord = xmlDoc.SelectSingleNode("//Name[text()='" + myUpgrades_frm.lboModules.List(lngItem, 0) + "']").ParentNode.SelectSingleNode("Type")
    myUpgrades_frm.lboModules.List(lngItem, 5) = FindRecord.Text
next_lngItem:
    lngItem = lngItem + 1
    Application.StatusBar = Application.StatusBar = "Preparing form...(" & Format(lngItem / rstStatus.RecordCount, "0%") & ")"
    DoEvents
    rstStatus.MoveNext
  Loop
  
  'populate branches
  Set xmlHttpDoc = CreateObject("MSXML2.XMLHTTP.6.0")
  strURL = "https://api.github.com/repos/clearplan/cpt/branches"
  xmlHttpDoc.Open "GET", strURL, False
  xmlHttpDoc.setRequestHeader "Content-Type", "application/json"
  xmlHttpDoc.setRequestHeader "Accept", "application/json"
  xmlHttpDoc.Send
  If xmlHttpDoc.Status = 200 And xmlHttpDoc.readyState = 4 Then
    Set RE = CreateObject("vbscript.regexp")
    With RE
      .MultiLine = False
      .Global = True
      .IgnoreCase = True
      '.Pattern = Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "[A-z0-9\-]*"
      .Pattern = Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "[A-z0-9-.]*"
    End With
    Set REMatches = RE.Execute(xmlHttpDoc.responseText)
    myUpgrades_frm.cboBranches.Clear
    For Each REMatch In REMatches
      myUpgrades_frm.cboBranches.AddItem Replace(REMatch, Chr(34) & "name" & Chr(34) & ":" & Chr(34), "")
    Next
    myUpgrades_frm.cboBranches.Value = "master"
  Else
    myUpgrades_frm.cboBranches.Clear
    myUpgrades_frm.cboBranches.AddItem "<unavailable>"
  End If
  myUpgrades_frm.Caption = "Installation Status (" & cptGetVersion("cptUpgrades_frm") & ")"
  Application.StatusBar = "Ready for user input..."
  DoEvents
  myUpgrades_frm.Show

exit_here:
  On Error Resume Next
  Set oUserForm = Nothing
  Unload myUpgrades_frm
  Set myUpgrades_frm = Nothing
  Application.StatusBar = ""
  If rstStatus.State Then rstStatus.Close
  Set rstStatus = Nothing
  Set REMatch = Nothing
  Set REMatches = Nothing
  Set RE = Nothing
  Set oStream = Nothing
  Set xmlHttpDoc = Nothing
  Application.StatusBar = ""
  Set vbComponent = Nothing
  Set xmlDoc = Nothing
  Set xmlNode = Nothing
  Set FindRecord = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptShowUpgrades_frm", Err, Erl)
  Resume exit_here

End Sub

Sub cptSetReferences()
  'this is a one-time shot to set all references currently required by the cp toolbar
  Dim oExcel As Object
  Dim strDir As String
  Dim strRegEx As String
  Dim vPath As Variant
  Dim vApp As Variant

  On Error Resume Next

  'CommonProgramFiles
  strDir = Environ("CommonProgramFiles")
  If Not cptReferenceExists("Office") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\OFFICE16\MSO.DLL"
  End If
  If Not cptReferenceExists("VBIDE") Then
    #If Not Win64 Then
      ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
    #Else
      ThisProject.VBProject.References.AddFromFile "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
    #End If
  End If
  If Not cptReferenceExists("VBA") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
  End If
  If Not cptReferenceExists("ADODB") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\System\ado\msado15.dll"
  End If

  'office applications
  For Each vApp In Split("EXCEL.EXE,MSOUTL.OLB,MSPPT.OLB,MSWORD.OLB", ",")
    strDir = cptGetOfficeDir2(CStr(vApp))
    If Len(strDir) > 0 Then
      ThisProject.VBProject.References.AddFromFile strDir & "\" & CStr(vApp)
    End If
  Next vApp

windows_common:
  If Not cptReferenceExists("MSForms") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\SysWOW64\FM20.DLL"
  End If
  If Not cptReferenceExists("Scripting") Then
    ThisProject.VBProject.References.AddFromFile "C:\Windows\SysWOW64\scrrun.dll"
  End If
  If Not cptReferenceExists("stdole") Then
    ThisProject.VBProject.References.AddFromFile "C:\Windows\SysWOW64\stdole2.tlb"
  End If
  If Not cptReferenceExists("mscorlib") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb"
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\mscorlib.dll"
  End If
  If Not cptReferenceExists("MSComctlLib") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\SysWOW64\MSCOMCTL.OCX"
  End If
  If Not cptReferenceExists("MSXML2") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\SysWOW64\msxml3.dll"
  End If
  If Not cptReferenceExists("VBScript_RegExp_55") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\System32\vbscript.dll\3"
  End If
  
exit_here:
  On Error Resume Next
  If Not oExcel Is Nothing Then oExcel.Quit
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptSetReferences", Err, Erl)
  Resume exit_here

End Sub

Function cptGetOfficeDir(strApp As String) As String
  Dim strDir As String
  Dim vPath As Variant
  
  strDir = ""
  For Each vPath In Split(Environ("PATH"), ";")
    If InStr(vPath, "Office") > 0 Then
      If Dir(CStr(vPath) & strApp) <> vbNullString Then
        strDir = vPath
        Exit For
      End If
    End If
  Next vPath

  If Len(strDir) > 0 Then
    cptGetOfficeDir = strDir
  ElseIf Len(strDir) = 0 Then 'weird installation or Excel not installed
    cptGetOfficeDir = strDir
    MsgBox "Microsoft Office installation is not detetcted. Some features may not operate as expected." & vbCrLf & vbCrLf & "Please contact help@ClearPlanConsulting.com for specialized assistance.", vbCritical + vbOKOnly, "Microsoft Office Compatibility"
  End If
  
End Function

Function cptGetOfficeDir2(strApp As String) As String
  Dim strDir As String
  Dim vPath As Variant
  Dim oFSO As Object 'Scripting.FileSystemObject
  Dim oFolder As Object 'Scripting.Folder
  strDir = ""
  For Each vPath In Split(Environ("PATH"), ";")
    strDir = cptRegEx(CStr(vPath), ".*Microsoft Office\\")
    If Len(strDir) > 0 Then
      Set oFSO = CreateObject("Scripting.FileSystemObject")
      Set oFolder = oFSO.GetFolder(strDir)
      cptGetOfficeDir2 = cptGetAppDir(oFolder, strApp)
    End If
  Next vPath
  
  Set oFSO = Nothing
  Set oFolder = Nothing
End Function

Function cptGetAppDir(oFolder As Object, strApp As String) As String
  Dim f As Object 'Scripting.File
  Dim sf As Object 'Scripting.Folder
  
  If Dir(oFolder.Path & "\" & strApp) <> vbNullString Then
    cptGetAppDir = oFolder.Path
  Else
    For Each sf In oFolder.SubFolders
      If Len(cptGetAppDir) > 0 Then Exit Function
      cptGetAppDir = cptGetAppDir(sf, strApp)
    Next sf
  End If
  Set f = Nothing
  Set sf = Nothing
End Function

Sub cptSubmitIssue()
  Dim strMsg As String
  Dim strSource As String
  Dim strDescription As String
  
  strSource = InputBox("What feature are you having an issue with?", "Creating a Ticket...", "e.g., Status Sheet")
  If Len(strSource) = 0 Then
    MsgBox "Nothing entered.", vbInformation + vbOKOnly, "Ticket cancelled"
    Exit Sub
  End If
  strDescription = InputBox("Please summarize the issue you're having:" & vbCrLf & vbCrLf & "(You will be able to provide details in a moment).", "Creating a Ticket...", "e.g., assignments won't export")
  If Len(strDescription) = 0 Then
    MsgBox "Nothing entered.", vbInformation + vbOKOnly, "Ticket cancelled"
    Exit Sub
  End If
  strMsg = "Thank you." & vbCrLf & vbCrLf
  strMsg = strMsg & "Feature: " & strSource & vbCrLf
  strMsg = strMsg & "Description: " & strDescription & vbCrLf & vbCrLf
  strMsg = strMsg & "Note: you may wish to edit the suggested 'SUBJECT' in the text file that is about to appear, and feel free to add as much or as little detail as you would like where it says REPLACE THIS LINE." & vbCrLf & vbCrLf
  strMsg = strMsg & "Please click 'Yes' on the next prompt to submit your support request."
  If MsgBox(strMsg, vbInformation + vbOKCancel, "Creating a Ticket...") = vbCancel Then
    MsgBox "Ticket cancelled.", vbInformation + vbOKOnly, "Ticket cancelled"
    Exit Sub
  End If
  On Error Resume Next
  Err.Raise 1, strSource, strDescription
  cptHandleErr "user-initiated", "user-initiated", Err
  
  'If Not Application.FollowHyperlink("https://clearplan.happyfox.com/new", , , True) Then
  '  Call cptSendMail("Issue")
  'End If
End Sub

Sub cptSubmitRequest()
  If Not Application.FollowHyperlink("https://clearplan.happyfox.com/new", , , True) Then
    Call cptSendMail("Request")
  End If
End Sub

Sub cptSubmitFeedback()
  If Not Application.FollowHyperlink("https://clearplan.happyfox.com/new", , , True) Then
    Call cptSendMail("Feedback")
  End If
End Sub

Sub cptSendMail(strCategory As String)
  'objects
  Dim oOutlook As Object 'Outlook.Application
  Dim oMailItem As Object 'MailItem
  'strings
  Dim strHTML As String
  Dim strURL As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  'get outlook
  On Error Resume Next
  Set oOutlook = GetObject(, "Outlook.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oOutlook Is Nothing Then
    Set oOutlook = CreateObject("Outlook.Application")
  End If
  If oOutlook Is Nothing Then
    MsgBox "Outlook is not available.", vbCritical + vbOKOnly, "SendMail"
    GoTo exit_here
  End If

  'create the email and set generic settings
  Set oMailItem = oOutlook.CreateItem(0) '0=olMailItem
  oMailItem.To = "help@ClearPlanConsulting.com"
  oMailItem.Importance = 2 '2=olImportanceHigh
  oMailItem.Display

  'get strURL and message body
  Select Case strCategory
    Case "Issue"
      oMailItem.Subject = "Issue: <enter brief summary of the issue>"
      On Error Resume Next
      Err.Raise 1, "User", "User-submitted Issue"
      cptHandleErr "cptCore_bas", "cptSendMail", Err
      strHTML = "<h3>Please Describe Your Environment:</h3><p>"
      strHTML = strHTML & "<i>Operating System</i>: [operating system]<p>"
      strHTML = strHTML & "<i>Microsoft Project Version</i>: [Standard / Professional] [Year]<p>"
      strHTML = strHTML & "<i>Do you have unfettered internet access (try opening <a href=""https://github.com/clearplan/cpt/blob/master/README.md"">this page</a>)?</i> [Yes/No]<p>"
      strHTML = strHTML & "<h3>Please Describe the Issue:</h3><p><p>"
      strHTML = strHTML & "<i>Please be as detailed as possible: what were you trying to do, what selections did you make, describe the file you are working on, etc.</i><p>"
      strHTML = strHTML & "<h3>Please Include Screenshot(s):</h3><p>Please include any screenshot(s) of any error messages or anything else that might help us troubleshoot this issue for you.<p><p>"
      strHTML = strHTML & "<i>Thank you for helping us improve the ClearPlan Toolbar!</i>"
      oMailItem.HTMLBody = strHTML & oMailItem.HTMLBody
      
    Case "Request"
      oMailItem.Subject = "Feature Request: <enter brief description of the feature>"
      strHTML = "<h3>Please Describe the Feature you are Requesting:</h3><p>&nbsp;<p>&nbsp;"
      strHTML = strHTML & "<i>Thank you for contributing to the ClearPlan Toolbar project!</i>"
      oMailItem.HTMLBody = strHTML & oMailItem.HTMLBody
      
    Case "Feedback"
      oMailItem.Subject = "Feedback: <enter summary of feedback>"
      strHTML = "<h3>Feedback:</h3><p>&nbsp;<p>&nbsp;<i>We sincerely appreciate any and all constructive feedback. Thank you for contributing!</i>"
      oMailItem.HTMLBody = strHTML & oMailItem.HTMLBody
      
  End Select
  
exit_here:
  On Error Resume Next
  Set oOutlook = Nothing
  Set oMailItem = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptSendMail", Err, Erl)
  Resume exit_here
End Sub

Function cptRemoveIllegalCharacters(strText As String, Optional strReplacement As String, Optional blnTrim As Boolean = False) As String
  Dim vBad As Variant
  If Len(cptRegEx(strReplacement, "[\/:*?<>|[]")) > 0 Or InStr(strReplacement, Chr(34)) > 0 Then
    MsgBox "Function cptRemoveIllegalCharacters() returned an error:" & vbCrLf & vbCrLf & "Replacement '" & strReplacement & "' contains illegal characters.", vbCritical + vbOKOnly, "Invalid"
    Debug.Print "Function cptRemoveIllegalCharacters() returned an error: replacement '" & strReplacement & "' contains illegal characters."
    cptRemoveIllegalCharacters = vbNullString
    Exit Function
  End If
  For Each vBad In Split("\,/,:,*,?,"",<,>,|", ",")
    strText = Replace(strText, vBad, strReplacement)
  Next vBad
  If blnTrim Then
    strText = Trim(strText)
    strText = Replace(strText, cptRegEx(strText, "\s{2,}"), " ")
  End If
  cptRemoveIllegalCharacters = strText
End Function

Sub cptWrapItUp(Optional lngOutlineLevel As Long)
  'objects
  'strings
  Dim strDir As String
  'longs
  Dim lngLevel As Long
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  
  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "WrapItUp"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  If lngOutlineLevel = 0 Then
    'check for a saved setting
    If Dir(strDir & "\settings\cpt-reset-all.adtg") <> vbNullString Then
      With CreateObject("ADODB.Recordset")
        .Open strDir & "\settings\cpt-reset-all.adtg"
        .MoveFirst
        lngOutlineLevel = .Fields(1)
        .Close
      End With
    Else
      lngOutlineLevel = 2
    End If
  End If
  
  cptSpeed True 'speed up
  Application.OpenUndoTransaction "WrapItUp"
  'FilterClear 'do not reset, keep autofilters
  'GroupClear 'do not reset, applies to groups to
  OptionsViewEx DisplaySummaryTasks:=True
  SelectAll
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
      Sort "ID", , , , , , False, True
      OutlineShowAllTasks
    Else
      SelectBeginning
      GoTo exit_here
    End If
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  OutlineShowTasks pjTaskOutlineShowLevelMax
  'pjTaskOutlineShowLevelMax = 65,535 = do not use
  For lngLevel = 20 To lngOutlineLevel Step -1
    OutlineShowTasks lngLevel
  Next lngLevel
  SelectBeginning

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  cptSpeed False
  Exit Sub

'no_tasks:
'  MsgBox "This project has no tasks to collapse.", vbExclamation + vbOKOnly, "WrapItUp"
'  GoTo exit_here

err_here:
  Call cptHandleErr("cptCore_bas", "cptWrapItUp", Err, Erl)
  Resume exit_here
End Sub

Sub cptWrapItUpAll()
  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "WrapItUp"
    Exit Sub
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  OptionsViewEx DisplaySummaryTasks:=True
  On Error Resume Next
  If ActiveProject.Subprojects.Count > 0 Then
    FilterClear
    SelectAll
  End If
  If Not OutlineShowAllTasks Then
    If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
      Sort "ID", , , , , , False, True
      OutlineShowAllTasks
    Else
      SelectBeginning
    End If
  End If

End Sub
Sub cptWrapItUp1()
  Call cptWrapItUp(1)
End Sub
Sub cptWrapItUp2()
  Call cptWrapItUp(2)
End Sub
Sub cptWrapItUp3()
  Call cptWrapItUp(3)
End Sub
Sub cptWrapItUp4()
  Call cptWrapItUp(4)
End Sub
Sub cptWrapItUp5()
  Call cptWrapItUp(5)
End Sub
Sub cptWrapItUp6()
  Call cptWrapItUp(6)
End Sub
Sub cptWrapItUp7()
  Call cptWrapItUp(7)
End Sub
Sub cptWrapItUp8()
  Call cptWrapItUp(8)
End Sub
Sub cptWrapItUp9()
  Call cptWrapItUp(9)
End Sub

Function cptVersionStatus(strInstalled As String, strCurrent As String) As String
  'objects
  'strings
  'longs
  Dim lngVersion As Long
  'integers
  'booleans
  'variants
  Dim aCurrent As Variant
  Dim aInstalled As Variant
  Dim vVersion As Variant
  Dim vLevel As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'clean the versions - include all three levels
  For Each vVersion In Array(strInstalled, strCurrent)
    'following line doesn't remove non-numeric characters
    vVersion = cptRegEx(CStr(vVersion), "([0-9].*.?){1,3}")
    If Len(vVersion) - Len(Replace(vVersion, ".", "")) < 1 Then
      vVersion = vVersion & ".0"
    End If
    If Len(vVersion) - Len(Replace(vVersion, ".", "")) < 2 Then
      vVersion = vVersion & ".0"
    End If
    If lngVersion = 0 Then
      aInstalled = Split(vVersion, ".")
    Else
      aCurrent = Split(vVersion, ".")
    End If
    lngVersion = lngVersion + 1
  Next

  'figure out the things
'  If strInstalled = "<missing>" Then
'    cptVersionStatus = "install"
'  Else
    For Each vLevel In Array(0, 1, 2)
      If aCurrent(vLevel) <> aInstalled(vLevel) Then
        cptVersionStatus = Choose(vLevel + 1, "major", "minor", "patch")
        If Len(aInstalled(vLevel)) = 0 Then
          cptVersionStatus = "install " & cptVersionStatus
        ElseIf CLng(aCurrent(vLevel)) > CLng(aInstalled(vLevel)) Then '<issue62>
          cptVersionStatus = cptVersionStatus & " upgrade"
        Else
          cptVersionStatus = cptVersionStatus & " downgrade"
        End If
        Exit For
      End If
    Next vLevel
'  End If
  
  If cptVersionStatus = "" Then
    cptVersionStatus = "ok"
  ElseIf cptVersionStatus = "install" Then
    cptVersionStatus = "install"
  Else
    cptVersionStatus = "install " & cptVersionStatus
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptVersionStatus", Err, Erl)
  Resume exit_here

End Function

Sub cptFilterReapply()

  ActiveWindow.TopPane.Activate
  GroupApply ActiveProject.CurrentGroup
  Exit Sub

End Sub

Sub cptGroupReapply()
  Dim strCurrentGroup As String
  Dim lngUID As Long
  lngUID = 0
  On Error Resume Next
  lngUID = ActiveSelection.Tasks(1).UniqueID
  strCurrentGroup = ActiveProject.CurrentGroup
  ScreenUpdating = False
  ActiveWindow.TopPane.Activate
  GroupApply "No Group"
  'todo: how to reapply an AutoFilter group?
  On Error Resume Next
  If Not GroupApply(strCurrentGroup) Then
    MsgBox "Cannot reapply a Custom AutoFilter Group", vbInformation + vbOKCancel, "Reapply Group"
  End If
  ScreenUpdating = True
  If lngUID > 0 Then Find "Unique ID", "equals", lngUID
End Sub

Function cptSaveSetting(strFeature As String, strKey As String, strValue As Variant) As Boolean
  Dim strSettingsFile As String, lngWorked As Long
  strSettingsFile = cptDir & "\settings\cpt-settings.ini"
  lngWorked = SetPrivateProfileString(strFeature, strKey, CStr(strValue), strSettingsFile)
  If lngWorked Then
    cptSaveSetting = True
  Else
    cptSaveSetting = False
  End If
End Function

Function cptGetSetting(strFeature As String, strKey As String) As String
  Dim strSettingsFile As String, strReturned As String, lngSize As Long, lngWorked As Long
  strSettingsFile = cptDir & "\settings\cpt-settings.ini"
  strReturned = Space(255) 'this determines the length of the returned value, not the length of the stored value
  lngSize = Len(strReturned)
  lngWorked = GetPrivateProfileString(strFeature, strKey, "", strReturned, lngSize, strSettingsFile)
  If lngWorked Then
    cptGetSetting = Left$(strReturned, lngWorked)
  Else
    cptGetSetting = ""
  End If
End Function

Function cptRenameSetting(strFeature As String, strOldKey As String, strNewKey As String) As Boolean
  Dim strValue As String
  strValue = cptGetSetting(strFeature, strOldKey)
  If Len(strValue) > 0 Then
    cptSaveSetting strFeature, strNewKey, strValue
    cptDeleteSetting strFeature, strOldKey
    cptRenameSetting = True
  Else
    cptRenameSetting = False
  End If
End Function

Function cptDeleteSetting(strFeature As String, strKey As String) As Boolean
  Dim strSettingsFile As String, lngWorked As Long
  strSettingsFile = cptDir & "\settings\cpt-settings.ini"
  lngWorked = SetPrivateProfileString(strFeature, strKey, CLng(0), strSettingsFile)
  If lngWorked Then
    cptDeleteSetting = True
  Else
    cptDeleteSetting = False
  End If
End Function

Function cptViewExists(strView As String) As Boolean
  'objects
  Dim oView As MSProject.View

  On Error Resume Next
  Set oView = ActiveProject.Views(strView)
  If oView Is Nothing Then
    Set oView = Application.GlobalViews(strView)
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptViewExists = Not oView Is Nothing
  
exit_here:
  On Error Resume Next
  Set oView = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptViewExists", Err, Erl)
  Resume exit_here
End Function

Function cptTableExists(strTable As String, Optional blnResourceTable As Boolean = False) As Boolean
  'objects
  Dim oTable As MSProject.Table

  On Error Resume Next
  If blnResourceTable Then
    Set oTable = ActiveProject.ResourceTables(strTable)
    If oTable Is Nothing Then
      Set oTable = Application.GlobalResourceTables(strTable)
    End If
  Else
    Set oTable = ActiveProject.TaskTables(strTable)
    If oTable Is Nothing Then
      Set oTable = Application.GlobalTaskTables(strTable)
    End If
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptTableExists = Not oTable Is Nothing
  
exit_here:
  On Error Resume Next
  Set oTable = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptTableExists", Err, Erl)
  Resume exit_here
End Function

Function cptFilterExists(strFilter As String, Optional blnResourceFilter As Boolean = False) As Boolean
  'objects
  Dim oFilter As MSProject.Filter

  On Error Resume Next
  If blnResourceFilter Then
    Set oFilter = ActiveProject.ResourceFilters(strFilter)
    If oFilter Is Nothing Then
      Set oFilter = Application.GlobalResourceFilters(strFilter)
    End If
  Else
    Set oFilter = ActiveProject.TaskFilters(strFilter)
    If oFilter Is Nothing Then
      Set oFilter = Application.GlobalTaskFilters(strFilter)
    End If
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptFilterExists = Not oFilter Is Nothing
  
exit_here:
  On Error Resume Next
  Set oFilter = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptFilterExists", Err, Erl)
  Resume exit_here
End Function

Function cptGroupExists(strGroup As String, Optional blnResourceGroup As Boolean = False) As Boolean
  'objects
  Dim oGroup As MSProject.Group

  On Error Resume Next
  If blnResourceGroup Then
    Set oGroup = ActiveProject.ResourceGroups(strGroup)
    If oGroup Is Nothing Then
      Set oGroup = Application.ResourceGroups(strGroup)
    End If
  Else
    Set oGroup = ActiveProject.TaskGroups(strGroup)
    If oGroup Is Nothing Then
      Set oGroup = Application.TaskGroups(strGroup)
    End If
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptGroupExists = Not oGroup Is Nothing
  
exit_here:
  On Error Resume Next
  Set oGroup = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGroupExists", Err, Erl)
  Resume exit_here
End Function

Sub cptCreateFilter(strFilter As String)
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Select Case strFilter
    Case "Marked"
      FilterEdit Name:="Marked", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Marked", test:="equals", Value:="Yes", ShowInMenu:=True, ShowSummaryTasks:=False
      
  End Select
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptCreateFilter", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowSettings_frm()
  'objects
  Dim mySettings_frm As cptSettings_frm
  Dim oRecordset As ADODB.Recordset
  Dim oStream As Scripting.TextStream
  Dim oFSO As Scripting.FileSystemObject
  'strings
  Dim strDir As String
  Dim strErrorTrapping As String
  Dim strSettingsFileNew As String
  Dim strSettingsFile As String
  Dim strProgramAcronym As String
  Dim strFeature As String
  Dim strLine As String
  'longs
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  'prevent spawning
  If Not cptGetUserForm("cptSettings_frm") Is Nothing Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  Set mySettings_frm = New cptSettings_frm
  mySettings_frm.Caption = "ClearPlan Toolbar Settings (" & cptGetVersion("cptSettings_frm") & ")"
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Fields.Append "Feature", adVarChar, 100
    .Fields.Append "Setting", adVarChar, 255
    .Open
  End With
  
  strSettingsFile = strDir & "\settings\cpt-settings.ini"
  strSettingsFileNew = strDir & "\settings\cpt-settings-temp.ini"
  lngFile = FreeFile
  Open strSettingsFileNew For Output As #lngFile
  
  With mySettings_frm
    .lboFeatures.Clear
    .lboSettings.Clear
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.OpenTextFile(strDir & "\settings\cpt-settings.ini")
    Do While Not oStream.AtEndOfStream
      strLine = oStream.ReadLine
      If Left(strLine, 1) = "[" Then
        strFeature = Replace(Replace(strLine, "[", ""), "]", "")
        'todo: If [Driving Path Group] OR [Driving Path] Then skip
        'todo: If Count Then StatusBar: StatusBar.blnSummarizeSelection; StatusBar.blnShowStatusBarTaskCount
        oRecordset.AddNew Array(0), Array(strFeature)
      Else
        'todo: If [Metrics] Then remove cboLOEField > Integration.EVT
        'todo: If [Metrics] Then remove txtLOE > Integration.LOE
        'todo: If [Driving Path Group] OR [Driving Path] Then skip
        If strFeature = "Integration" Then
          If Left(strLine, 5) = "CWBS=" Then
            strLine = Replace(strLine, "CWBS=", "WBS=")
          ElseIf Left(strLine, 5) = "WPCN=" Then
            strLine = Replace(strLine, "WPCN=", "WP=")
          End If
        End If
        oRecordset.AddNew Array(0, 1), Array(strFeature, strLine)
      End If
    Loop
    oStream.Close
    oRecordset.Sort = "Feature,Setting"
    oRecordset.MoveFirst
    Do While Not oRecordset.EOF
      If oRecordset(1) = "" Then
        .lboFeatures.AddItem oRecordset(0)
        Print #lngFile, "[" & oRecordset(0) & "]"
      Else
        Print #lngFile, oRecordset(1)
      End If
      oRecordset.MoveNext
    Loop
    Close #lngFile
    If Dir(strSettingsFile) <> vbNullString Then Kill strSettingsFile
    Name strSettingsFileNew As strSettingsFile
    If Dir(strSettingsFileNew) <> vbNullString Then Kill strSettingsFileNew
    If Dir(strDir & "\settings\cpt-settings.adtg") <> vbNullString Then Kill strDir & "\settings\cpt-settings.adtg"
    oRecordset.Save strDir & "\settings\cpt-settings.adtg", adPersistADTG
    oRecordset.Close
    '.lblDir = strSettingsFile
    strProgramAcronym = cptGetProgramAcronym
    .txtProgramAcronym = strProgramAcronym
    If .lboFeatures.ListCount > 0 Then
      .lboFeatures.Value = .lboFeatures.List(0, 0)
    End If
    'get error-trapping on/off and set toggle button
    strErrorTrapping = cptGetSetting("General", "ErrorTrapping")
    If Len(strErrorTrapping) > 0 Then
      .tglErrorTrapping = strErrorTrapping = 0
    Else
      .tglErrorTrapping = False
    End If
    .Show
  End With
  
exit_here:
  On Error Resume Next
  Reset
  Set oRecordset = Nothing
  Set oStream = Nothing
  Set oFSO = Nothing
  Unload mySettings_frm
  Set mySettings_frm = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptShowSettings_frm", Err, Erl)
  Resume exit_here
End Sub

Function cptGetProgramAcronym() As String
  'objects
  Dim oCustomDocumentProperty As DocumentProperty
  'strings
  Dim strMsg As String
  Dim strProgramAcronym As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  Dim vResponse As Variant
  'dates
  
  On Error Resume Next
  Set oCustomDocumentProperty = ActiveProject.CustomDocumentProperties("cptProgramAcronym")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oCustomDocumentProperty Is Nothing Then
    strMsg = "For some features, a unique program acronym is required to capture data (locally)." & vbCrLf & vbCrLf
    strMsg = strMsg & "This program acronym is saved in a custom document property named 'cptProgramAcronym'." & vbCrLf & vbCrLf
    strMsg = strMsg & "Please enter a program acronym for this file:"
    vResponse = InputBox(strMsg, "Program Acronym")
    If StrPtr(vResponse) = 0 Then
      MsgBox "No Program Acronym saved.", vbCritical + vbOKOnly, "Invalid Response"
      cptGetProgramAcronym = ""
    ElseIf vResponse = vbNullString Then
      MsgBox "No Program Acronym saved.", vbCritical + vbOKOnly, "Invalid Response"
      cptGetProgramAcronym = ""
    Else
      Set oCustomDocumentProperty = ActiveProject.CustomDocumentProperties.Add("cptProgramAcronym", False, msoPropertyTypeString, CStr(vResponse))
      cptGetProgramAcronym = CStr(vResponse)
      MsgBox "Program Acronym '" & CStr(vResponse) & "' saved!", vbInformation + vbOKOnly, "Success"
    End If
  Else
    cptGetProgramAcronym = ActiveProject.CustomDocumentProperties("cptProgramAcronym").Value
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetProgramAcronym", Err, Erl)
  Resume exit_here
End Function

Sub cptOpenSettingsFile()
  Dim strFileName As String
  strFileName = cptDir & "\settings\cpt-settings.ini"
  ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
End Sub

Function cptGetMyHeaders(strTitle As String, Optional blnRequired As Boolean = False) As String
  'objects
  'strings
  Dim strMyHeaders As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vResponse As Variant
  Dim vMyHeader As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

try_again:
  'get other fields
  strMyHeaders = cptGetSetting("Metrics", "txtMyHeaders")
  If Len(strMyHeaders) = 0 Then strMyHeaders = "CAM,WP,"
  If blnRequired Then
    vResponse = InputBox("At least one custom field is required." & vbCrLf & vbCrLf & "Enter a comma-separated list:", strTitle, strMyHeaders)
  Else
    vResponse = InputBox("Include other Custom Fields (e.g., 'CAM,WP,')?" & vbCrLf & vbCrLf & "Enter a comma-separated list of Custom Field Names (or leave blank for none):", strTitle, strMyHeaders)
  End If
  
  If StrPtr(vResponse) = 0 Then 'user hit cancel
    strMyHeaders = ""
    GoTo exit_here
  ElseIf vResponse = "" Or Len(Replace(vResponse, ",", "")) = 0 Then 'user entered zero-value
    If blnRequired Then
      'nothing selected
      If MsgBox("You must select at least one custom field. Try again?", vbQuestion + vbYesNo, "Field Required") = vbYes Then
        GoTo try_again
      Else
        strMyHeaders = ""
        GoTo exit_here
      End If
    Else
      strMyHeaders = ""
      GoTo exit_here
    End If
  ElseIf Len(vResponse) > 0 Then
    strMyHeaders = CStr(vResponse)
  End If

  Application.StatusBar = "Validating custom fields..."
  DoEvents
  If Right(Trim(strMyHeaders), 1) <> "," Then strMyHeaders = Trim(strMyHeaders) & ","
  'validate strMyHeaders
  On Error Resume Next
  For Each vMyHeader In Split(strMyHeaders, ",")
    If vMyHeader = "" Then Exit For
    If cptCustomFieldExists(CStr(vMyHeader)) = 0 Then
      vResponse = MsgBox("Custom Field '" & vMyHeader & "' not found!" & vbCrLf & vbCrLf & "OK = skip; Cancel = try again", vbExclamation + vbOKCancel, "Invalid Field")
      If vResponse = vbCancel Then
        GoTo try_again
      Else
        strMyHeaders = Replace(strMyHeaders, vMyHeader & ",", "")
      End If
    End If
  Next vMyHeader
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSaveSetting "Metrics", "txtMyHeaders", strMyHeaders

  cptGetMyHeaders = strMyHeaders

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetMyHeaders", Err, Erl)
  Resume exit_here

End Function

Function cptMasterUIDToSubUID(lngMasterUID As Long) As Long
  If ActiveProject.Subprojects.Count = 0 Then Exit Function
  cptMasterUIDToSubUID = lngMasterUID Mod 4194304
End Function

Function cptGetSubprojectUID(lngMasterUID As Long) As Long
  If ActiveProject.Subprojects.Count = 0 Then Exit Function
  cptGetSubprojectUID = Round(lngMasterUID / 4194304, 0) - 1
End Function

Function cptSubUIDToMasterUID(lngSubProjectUID As Long, lngSubUID As Long) As Long
  If ActiveProject.Subprojects.Count = 0 Then Exit Function
  cptSubUIDToMasterUID = ((lngSubProjectUID + 1) * 4194304) + lngSubUID
End Function

Function cptConvertToMasterUIDs(oTask As MSProject.Task, strReturn As String) As String
  'strReturn variable expects either "p" for predecessors or "s" for successors
  Dim oSubprojects As MSProject.Subprojects
  Dim strProject As String, strList As String, strLinkProject As String, strConvertedList As String
  Dim lngUID As Long, lngLinkUID As Long
  Dim vLink As Variant
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oSubprojects = ActiveProject.Subprojects
  If oSubprojects.Count = 0 Then GoTo exit_here
  
  strProject = oTask.Project
  If strReturn = "p" Then
    strList = oTask.UniqueIDPredecessors
  ElseIf strReturn = "s" Then
    strList = oTask.UniqueIDSuccessors
  End If
  For Each vLink In Split(strList, ",")
    If InStr(vLink, "\") > 0 Then 'handle offline and server paths
      lngUID = CLng(Mid(vLink, InStrRev(vLink, "\") + 1))   'extract source task UID
      strLinkProject = Replace(vLink, "\" & lngUID, "")     'strip source task UID
      strLinkProject = Replace(strLinkProject, ".mpp", "")  'strip file extension
      strLinkProject = Mid(strLinkProject, InStrRev(strLinkProject, "\") + 1) 'strip path
      lngLinkUID = oSubprojects(strLinkProject).InsertedProjectSummary.UniqueID + 1 'get master task UID seed
    Else
      lngUID = vLink
      lngLinkUID = oSubprojects(strProject).InsertedProjectSummary.UniqueID + 1 'get master task UID seed
    End If
    lngLinkUID = (lngLinkUID * 4194304) + lngUID            'derive master task UID
    strConvertedList = strConvertedList & lngLinkUID & ","  'build return string
  Next vLink
  strConvertedList = Left(strConvertedList, Len(strConvertedList) - 1) 'strip last comma avoiding null value
  cptConvertToMasterUIDs = strConvertedList

exit_here:
  On Error Resume Next
  Set oSubprojects = Nothing
  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptConvertToMasterUIDs", Err, Erl)
  Resume exit_here
End Function

Function cptGetShowStatusBarCountFirstRun() As Boolean
  'objects
  'strings
  Dim strShow As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnShow As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strShow = cptGetSetting("Count", "blnShowStatusBarTaskCount")
  If Len(strShow) > 0 Then
    blnShow = CBool(strShow)
  Else
    Call cptSaveSetting("Count", "blnShowStatusBarTaskCount", "1") 'default is true
    blnShow = True
  End If

  cptGetShowStatusBarCountFirstRun = blnShow

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetShowStatusBarCountFirstRun", Err, Erl)
  Resume exit_here
End Function

Sub cptAppendColumn(strFileName As String, strColumn As String, lngType As Long, Optional lngLength As Long, Optional vDefault As Variant)
  'objects
  Dim oRecordsetNew As Object 'ADODB.Recordset
  Dim oRecordset As Object 'ADODB.Recordset
  'strings
  Dim strDir As String
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strDir = cptDir
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  Set oRecordsetNew = CreateObject("ADODB.Recordset")
  If InStr(strFileName, strDir) = 0 Then strFileName = strDir & strFileName
  oRecordset.Open strFileName, , adOpenKeyset, adLockReadOnly
  On Error Resume Next
  Debug.Print oRecordset.Fields(strColumn)
  If Err.Number = 0 Then 'field already exists
    Err.Clear
    If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    GoTo exit_here
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  'first rebuld the existing fields
  For lngField = 0 To oRecordset.Fields.Count - 1
    With oRecordsetNew
      If oRecordset.Fields(lngField).DefinedSize > 0 Then
        .Fields.Append oRecordset.Fields(lngField).Name, oRecordset.Fields(lngField).Type, oRecordset.Fields(lngField).DefinedSize
      Else
        .Fields.Append oRecordset.Fields(lngField).Name, oRecordset.Fields(lngField).Type
      End If
    End With
  Next lngField
  'next add the new field
  If lngLength > 0 Then
    oRecordsetNew.Fields.Append strColumn, lngType, lngLength
  Else
    oRecordsetNew.Fields.Append strColumn, lngType
  End If
  oRecordsetNew.Open
  'next move the existing data over
  If Not oRecordset.EOF Then oRecordset.MoveFirst
  Do While Not oRecordset.EOF
    oRecordsetNew.AddNew
    For lngField = 0 To oRecordset.Fields.Count - 1
      oRecordsetNew.Fields(lngField) = oRecordset.Fields(lngField)
    Next lngField
    oRecordsetNew.Fields(strColumn) = vDefault
    oRecordset.MoveNext
  Loop
  oRecordset.Close
  Name strFileName As Replace(strFileName, ".adtg", "-backup_" & Format(Now, "yyyy-mm-dd-HH-nn-ss") & ".adtg")
  oRecordsetNew.Save strFileName, adPersistADTG
  oRecordsetNew.Close
  
exit_here:
  On Error Resume Next
  Set oRecordsetNew = Nothing
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptAppendColumn", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetSums(ByRef oTasks As MSProject.Tasks, lngFieldID As Long)
  'objects
  Dim oTask As MSProject.Task
  'strings
  Dim strCustomFieldName As String
  Dim strFieldName As String
  'longs
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngDuration As Long
  'integers
  'doubles
  Dim dblWork As Double
  Dim dblCost As Double
  Dim dblNumber As Double
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  dblCost = 0#
  lngDuration = 0
  dblNumber = 0#
  dblWork = 0#
  lngTasks = oTasks.Count
  
  strFieldName = FieldConstantToFieldName(lngFieldID)
  
  If Len(cptRegEx(strFieldName, "Cost|Duration|Number|Work")) > 0 Then
    lngTask = 0
    For Each oTask In oTasks
      If oTask Is Nothing Then GoTo next_task
      If Not oTask.Active Then GoTo next_task
      If oTask.GetField(lngFieldID) = "#ERROR" Then GoTo next_task
      'do not ignore external tasks
      'do not ignore summary tasks
      Select Case strFieldName
        Case "Actual Cost"
          dblCost = dblCost + oTask.ActualCost
        Case "Cost"
          dblCost = dblCost + oTask.Cost
        Case "Remaining Cost"
          dblCost = dblCost + oTask.RemainingCost
        Case "Cost1"
          dblCost = dblCost + oTask.Cost1
        Case "Cost2"
          dblCost = dblCost + oTask.Cost2
        Case "Cost3"
          dblCost = dblCost + oTask.Cost3
        Case "Cost4"
          dblCost = dblCost + oTask.Cost4
        Case "Cost5"
          dblCost = dblCost + oTask.Cost5
        Case "Cost6"
          dblCost = dblCost + oTask.Cost6
        Case "Cost7"
          dblCost = dblCost + oTask.Cost7
        Case "Cost8"
          dblCost = dblCost + oTask.Cost8
        Case "Cost9"
          dblCost = dblCost + oTask.Cost9
        Case "Cost10"
          dblCost = dblCost + oTask.Cost10
          
        Case "Baseline Cost"
          dblCost = dblCost + Val(oTask.BaselineCost)
        Case "Baseline1 Cost"
          dblCost = dblCost + Val(oTask.Baseline1Cost)
        Case "Baseline2 Cost"
          dblCost = dblCost + Val(oTask.Baseline2Cost)
        Case "Baseline3 Cost"
          dblCost = dblCost + Val(oTask.Baseline3Cost)
        Case "Baseline4 Cost"
          dblCost = dblCost + Val(oTask.Baseline4Cost)
        Case "Baseline5 Cost"
          dblCost = dblCost + Val(oTask.Baseline5Cost)
        Case "Baseline6 Cost"
          dblCost = dblCost + Val(oTask.Baseline6Cost)
        Case "Baseline7 Cost"
          dblCost = dblCost + Val(oTask.Baseline7Cost)
        Case "Baseline8 Cost"
          dblCost = dblCost + Val(oTask.Baseline8Cost)
        Case "Baseline9 Cost"
          dblCost = dblCost + Val(oTask.Baseline9Cost)
        Case "Baseline10 Cost"
          dblCost = dblCost + Val(oTask.Baseline10Cost)
          
        Case "Actual Duration"
          lngDuration = lngDuration + oTask.ActualDuration
        Case "Duration"
          lngDuration = lngDuration + oTask.Duration
        Case "Remaining Duration"
          lngDuration = lngDuration + oTask.RemainingDuration
        Case "Duration1"
          lngDuration = lngDuration + oTask.Duration1
        Case "Duration2"
          lngDuration = lngDuration + oTask.Duration2
        Case "Duration3"
          lngDuration = lngDuration + oTask.Duration3
        Case "Duration4"
          lngDuration = lngDuration + oTask.Duration4
        Case "Duration5"
          lngDuration = lngDuration + oTask.Duration5
        Case "Duration6"
          lngDuration = lngDuration + oTask.Duration6
        Case "Duration7"
          lngDuration = lngDuration + oTask.Duration7
        Case "Duration8"
          lngDuration = lngDuration + oTask.Duration8
        Case "Duration9"
          lngDuration = lngDuration + oTask.Duration9
        Case "Duration10"
          lngDuration = lngDuration + oTask.Duration10
          
        Case "Baseline Duration"
          lngDuration = lngDuration + oTask.BaselineDuration
        Case "Baseline1 Duration"
          lngDuration = lngDuration + oTask.Baseline1Duration
        Case "Baseline2 Duration"
          lngDuration = lngDuration + oTask.Baseline2Duration
        Case "Baseline3 Duration"
          lngDuration = lngDuration + oTask.Baseline3Duration
        Case "Baseline4 Duration"
          lngDuration = lngDuration + oTask.Baseline4Duration
        Case "Baseline5 Duration"
          lngDuration = lngDuration + oTask.Baseline5Duration
        Case "Baseline6 Duration"
          lngDuration = lngDuration + oTask.Baseline6Duration
        Case "Baseline7 Duration"
          lngDuration = lngDuration + oTask.Baseline7Duration
        Case "Baseline8 Duration"
          lngDuration = lngDuration + oTask.Baseline8Duration
        Case "Baseline9 Duration"
          lngDuration = lngDuration + oTask.Baseline9Duration
        Case "Baseline10 Duration"
          lngDuration = lngDuration + oTask.Baseline10Duration
                    
        Case "Number"
          dblNumber = dblNumber + oTask.Number
        Case "Number1"
          dblNumber = dblNumber + oTask.Number1
        Case "Number2"
          dblNumber = dblNumber + oTask.Number2
        Case "Number3"
          dblNumber = dblNumber + oTask.Number3
        Case "Number4"
          dblNumber = dblNumber + oTask.Number4
        Case "Number5"
          dblNumber = dblNumber + oTask.Number5
        Case "Number6"
          dblNumber = dblNumber + oTask.Number6
        Case "Number7"
          dblNumber = dblNumber + oTask.Number7
        Case "Number8"
          dblNumber = dblNumber + oTask.Number8
        Case "Number9"
          dblNumber = dblNumber + oTask.Number9
        Case "Number10"
          dblNumber = dblNumber + oTask.Number10
        Case "Number11"
          dblNumber = dblNumber + oTask.Number11
        Case "Number12"
          dblNumber = dblNumber + oTask.Number12
        Case "Number13"
          dblNumber = dblNumber + oTask.Number13
        Case "Number14"
          dblNumber = dblNumber + oTask.Number14
        Case "Number15"
          dblNumber = dblNumber + oTask.Number15
        Case "Number16"
          dblNumber = dblNumber + oTask.Number16
        Case "Number17"
          dblNumber = dblNumber + oTask.Number17
        Case "Number18"
          dblNumber = dblNumber + oTask.Number18
        Case "Number19"
          dblNumber = dblNumber + oTask.Number19
        Case "Number20"
          dblNumber = dblNumber + oTask.Number20
          
        Case "Actual Work"
          dblWork = dblWork + oTask.ActualWork
        Case "Work"
          dblWork = dblWork + oTask.Work
        Case "Remaining Work"
          dblWork = dblWork + oTask.RemainingWork
        Case "Baseline Work"
          dblWork = dblWork + oTask.BaselineWork
        Case "Baseline1 Work"
          dblWork = dblWork + oTask.Baseline1Work
        Case "Baseline2 Work"
          dblWork = dblWork + oTask.Baseline2Work
        Case "Baseline3 Work"
          dblWork = dblWork + oTask.Baseline3Work
        Case "Baseline4 Work"
          dblWork = dblWork + oTask.Baseline4Work
        Case "Baseline5 Work"
          dblWork = dblWork + oTask.Baseline5Work
        Case "Baseline6 Work"
          dblWork = dblWork + oTask.Baseline6Work
        Case "Baseline7 Work"
          dblWork = dblWork + oTask.Baseline7Work
        Case "Baseline8 Work"
          dblWork = dblWork + oTask.Baseline8Work
        Case "Baseline9 Work"
          dblWork = dblWork + oTask.Baseline9Work
        Case "Baseline10 Work"
          dblWork = dblWork + oTask.Baseline10Work
      End Select

next_task:
      lngTask = lngTask + 1
      Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | Calculating...(" & Format(lngTask / lngTasks, "0%") & ")"
    Next oTask
  End If
  
  strCustomFieldName = CustomFieldGetName(lngFieldID)
  If Len(strCustomFieldName) > 0 Then strFieldName = strCustomFieldName & " (" & strFieldName & ")"

  If dblCost > 0 Then
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | " & strFieldName & ": " & Format(dblCost, "$#,###,##0.00")
  ElseIf lngDuration > 0 Then
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | " & strFieldName & ": " & Format(lngDuration / 480, "#,###,##0d")
  ElseIf dblNumber > 0 Then
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | " & strFieldName & ": " & Format(dblNumber, "#,###,##0.00")
  ElseIf dblWork > 0 Then
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | " & strFieldName & ": " & Format(dblWork / 60, "#,###,##0.00h")
  Else
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected"
  End If
  
exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetSums", Err, Erl)
  Resume exit_here
End Sub

Function cptGetCustomFields(strFieldTypes As String, strDataTypes As String, strInclude As String, Optional blnIncludeEnterprise As Boolean = False) As Variant
  'strFieldTypes  := comma-separated list of any of "p,t,r" [project,task,resource]
  'strDataTypes   := comma-separated list of any of "Cost,Date,Duration,Flag,Finish,Number,Outline Code,Start,Text"
  'strInclude     := comma-separated list of any of "c,fn,cfn,loc" [constant,fieldname,customfieldname,location(LCF|ECF)]
  'blnIncludeEnterprise := self-explanatory
  'objects
  Dim oFieldTypes As Object 'Scripting.Dictionary
  'strings
  Dim strFieldName As String
  Dim strCustomFieldName As String
  Dim strResult As String
  Dim strRow As String
  'longs
  Dim lngInclude As Long
  Dim lngFieldCount As Long
  Dim lngFieldType As Long
  Dim lngField As Long
  Dim lngConstant As Long
  Dim lngResultCount As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vInclude As Variant
  Dim vFieldType As Variant
  Dim vField As Variant
  Dim vRow() As Variant
  Dim vResult() As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oFieldTypes = CreateObject("Scripting.Dictionary")
  oFieldTypes.Add "p", pjProject
  oFieldTypes.Add "t", pjTask
  oFieldTypes.Add "r", pjResource
  
  vInclude = Split(strInclude, ",")
  ReDim vRow(0 To UBound(vInclude))
  For Each vFieldType In Split(strFieldTypes, ",")
    lngFieldType = oFieldTypes(vFieldType)
    For Each vField In Split(strDataTypes, ",")
      lngFieldCount = 10
      If vField = "Flag" Then lngFieldCount = 20
      If vField = "Number" Then lngFieldCount = 20
      If vField = "Text" Then lngFieldCount = 30
      For lngField = 1 To lngFieldCount
        lngConstant = FieldNameToFieldConstant(vField & lngField, lngFieldType)
        'Debug.Print lngConstant; FieldConstantToFieldName(lngConstant)
        For lngInclude = 0 To UBound(vInclude)
          If vInclude(lngInclude) = "c" Then
            vRow(lngInclude) = lngConstant
          ElseIf vInclude(lngInclude) = "fn" Then
            vRow(lngInclude) = FieldConstantToFieldName(lngConstant)
          ElseIf vInclude(lngInclude) = "cfn" Then
            If Len(CustomFieldGetName(lngConstant)) > 0 Then
              vRow(lngInclude) = CustomFieldGetName(lngConstant)
            Else
              vRow(lngInclude) = FieldConstantToFieldName(lngConstant)
              'todo: if blnHideCFN then vRow(lngInclude) = ""
            End If
          ElseIf vInclude(lngInclude) = "loc" Then
            vRow(lngInclude) = "LCF"
          End If
        Next lngInclude
        strResult = strResult & Join(vRow, ",") & vbCrLf
        lngResultCount = lngResultCount + 1
      Next lngField
    Next vField
  Next vFieldType
  
  'get enterprise custom fields
  If blnIncludeEnterprise Then
    For lngConstant = 188776000 To 188778000 '2000 should do it for now
      If FieldConstantToFieldName(lngConstant) <> "<Unavailable>" Then
        For lngInclude = 0 To UBound(vInclude)
          If vInclude(lngInclude) = "c" Then
            vRow(lngInclude) = lngConstant
          ElseIf vInclude(lngInclude) = "fn" Then
            vRow(lngInclude) = FieldConstantToFieldName(lngConstant)
          ElseIf vInclude(lngInclude) = "cfn" Then
            If Len(CustomFieldGetName(lngConstant)) > 0 Then
              vRow(lngInclude) = CustomFieldGetName(lngConstant)
            Else
              vRow(lngInclude) = FieldConstantToFieldName(lngConstant)
            End If
          ElseIf vInclude(lngInclude) = "loc" Then
            vRow(lngInclude) = "ECF"
          End If
        Next lngInclude
        strResult = strResult & Join(vRow, ",") & vbCrLf
        lngResultCount = lngResultCount + 1
      End If
    Next lngConstant
  End If
  
  ReDim vResult(0 To UBound(Split(strResult, vbCrLf)) - 1, 0 To UBound(vInclude))
  For lngField = 0 To UBound(Split(strResult, vbCrLf)) - 1
    For lngInclude = 0 To UBound(vInclude)
      vResult(lngField, lngInclude) = Split(Split(strResult, vbCrLf)(lngField), ",")(lngInclude)
    Next lngInclude
  Next lngField
    
  'alphabetization is handled by cptSortedArray on a case-by-case
  cptGetCustomFields = vResult
  
exit_here:
  On Error Resume Next
  Set oFieldTypes = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetCustomFields", Err, Erl)
  Resume exit_here
  
End Function

Sub cptGetValidMap(Optional strRequiredFields As String)
  Dim blnValidMap As Boolean
  Dim strMsg As String
  Dim lngResponse As Long

  If cptModuleExists("cptDECM_bas") And cptModuleExists("cptIntegration_frm") Then
    blnValidMap = cptValidMap(strRequiredFields, False, False, True)
  Else
    strMsg = "Please install the modules 'cptDECM_bas' and 'cptIntegration_frm' from the latest release."
    strMsg = strMsg & "Go to GitHub now?"
    If MsgBox(strMsg, vbExclamation + vbYesNo, "Missing Modules") = vbYes Then
      Application.FollowHyperlink "https://www.GitHub.com/clearplan/cpt/releases/latest"
    End If
  End If
End Sub

Function cptValidMap(Optional strRequiredFields As String, Optional blnFiscalRequired As Boolean = False, Optional blnRollingWaveDateRequired As Boolean = False, Optional blnConfirmationRequired As Boolean = False) As Boolean
  'objects
  Dim myIntegration_frm As cptIntegration_frm
  Dim oRequiredFields As Object 'Scripting.Dictionary
  Dim oComboBox As MSForms.ComboBox
  'strings
  Dim strPP As String
  Dim strDefaultFields As String
  Dim strLOE As String
  Dim strSetting As String
  'longs
  Dim lngItem  As Long
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  Dim blnECF As Boolean
  Dim blnErrorTrapping As Boolean
  Dim blnUseDefault As Boolean
  Dim blnValid As Boolean
  'variants
  Dim vRequired As Variant
  Dim vAddField  As Variant
  Dim vFields As Variant
  Dim vControl As Variant
  'dates
  
  'prevent spawning
  If Not cptGetUserForm("cptIntegration_frm") Is Nothing Then Exit Function
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnValid = True
  
  strDefaultFields = "WBS,OBS,CA,CAM,WP,WPM,EVT,LOE,PP,EVTMS,EVP"
  'todo: distinguish between default,enabled,required
  blnUseDefault = Len(strRequiredFields) = 0
  Set oRequiredFields = CreateObject("Scripting.Dictionary")
  For Each vRequired In Split(strDefaultFields, ",")
    oRequiredFields.Add vRequired, blnUseDefault
  Next vRequired
  oRequiredFields("WPM") = False
  oRequiredFields("EVTMS") = False
  oRequiredFields("PP") = False
  For Each vRequired In Split(strRequiredFields, ",")
    If Left(vRequired, 1) <> "[" And Right(vRequired, 1) <> "]" Then
      oRequiredFields(vRequired) = True
    End If
  Next vRequired
  'todo: LOE and PP must have selection, even if it's just '<unused>'
  
  blnECF = False 'default
  For Each vControl In Split(strDefaultFields, ",")
    If vControl <> "LOE" And vControl <> "PP" Then
      strSetting = cptGetSetting("Integration", CStr(vControl))
      If Len(strSetting) > 0 And strSetting <> "<unused>" Then
        If CLng(Split(strSetting, "|")(0)) >= 188776000 Then
          blnECF = True 'at least one ECF is mapped
          Exit For
        End If
      End If
    End If
  Next vControl
  
  Set myIntegration_frm = New cptIntegration_frm
  With myIntegration_frm
    
    .Caption = "Integration (" & cptGetVersion("cptIntegration_frm") & ")"
    .lblMasterOnly.Visible = ActiveProject.Subprojects.Count > 0
      
    .chkLCF.Enabled = False 'always
    .chkLCF.Value = True 'always
    If blnECF Then
      .chkECF.Enabled = True
      .chkECF.Value = True
      .chkECF.Locked = True
    Else
      If Application.Edition = pjEditionProfessional Then
        .chkECF.Enabled = True
        strSetting = cptGetSetting("Integration", "chkECF")
        If Len(strSetting) > 0 Then
          blnECF = CBool(strSetting)
          .chkECF.Value = blnECF
        End If
      Else
        .chkECF.Value = False
        .chkECF.Enabled = False
      End If
    End If
    'convert legacy saved settings
    strSetting = cptGetSetting("Integration", "CWBS")
    If Len(strSetting) > 0 Then
      cptSaveSetting "Integration", "WBS", strSetting
      cptDeleteSetting "Integration", "CWBS"
    End If
    strSetting = cptGetSetting("Integration", "WPCN")
    If Len(strSetting) > 0 Then
      cptSaveSetting "Integration", "WP", strSetting
      cptDeleteSetting "Integration", "WPCN"
    End If
    cptDeleteSetting "Integration", "EOC" 'no longer required
    
    For Each vControl In Split(strDefaultFields, ",")
      strSetting = cptGetSetting("Integration", CStr(vControl))
      If Len(strSetting) = 0 Then 'pull from Metrics
        If vControl = "EVP" Then
          strSetting = cptGetSetting("Metrics", "cboEVP")
          If Len(strSetting) = 0 Then
            If oRequiredFields(vControl) Then blnValid = False
          Else
            strSetting = strSetting & "|" & FieldConstantToFieldName(strSetting)
            cptSaveSetting "Integration", "EVP", strSetting
            cptDeleteSetting "Metrics", "cboEVP"
          End If
        ElseIf vControl = "EVT" Or vControl = "EVTMS" Then
          strSetting = cptGetSetting("Metrics", "cboLOEField")
          If Len(strSetting) = 0 Then
            If oRequiredFields(vControl) Then blnValid = False
          Else
            If vControl = "EVT" Then
              strSetting = strSetting & "|" & FieldConstantToFieldName(strSetting)
              cptSaveSetting "Integration", CStr(vControl), strSetting
              cptDeleteSetting "Metrics", "cboEVT"
            Else
              strSetting = cptGetSetting("Integration", CStr(vControl))
            End If
          End If
        ElseIf vControl = "LOE" Then
          strSetting = cptGetSetting("Metrics", "txtLOE")
          If Len(strSetting) = 0 Then
            If oRequiredFields(vControl) Then blnValid = False
          Else
            cptSaveSetting "Integration", "LOE", strSetting
            cptDeleteSetting "Metrics", "txtLOE"
          End If
        End If
      End If
      Set oComboBox = .Controls("cbo" & vControl)
      'default to black borders
      oComboBox.BorderColor = -2147483642
      If Len(strSetting) = 0 Then
        lngField = 0
      Else
        If vControl = "LOE" Then
          strLOE = strSetting
        ElseIf vControl = "PP" Then
          strPP = strSetting
        Else
          lngField = CLng(Split(strSetting, "|")(0))
        End If
      End If
      If vControl <> "LOE" And vControl <> "PP" Then
        If blnECF Then
          oComboBox.ColumnCount = 3
          oComboBox.ColumnWidths = "0 pt;105 pt;10 pt"
          oComboBox.ListWidth = 140
        Else
          oComboBox.ColumnCount = 2
          oComboBox.ColumnWidths = "0 pt"
          oComboBox.ListWidth = oComboBox.Width
        End If
      End If
      If vControl = "WBS" Then
        vFields = cptSortedArray(cptGetCustomFields("t", "Outline Code,Text", "c,cfn,loc", blnECF), 1)
        For lngItem = 0 To UBound(vFields)
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
          If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = vFields(lngItem, 2)
        Next lngItem
        If IsEmpty(oComboBox.List(oComboBox.ListCount - 1, 0)) Then oComboBox.RemoveItem (oComboBox.ListCount - 1)
      ElseIf vControl = "CAM" Or vControl = "WPM" Then
        For Each vAddField In Split("Contact", ",")
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = FieldNameToFieldConstant(vAddField)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vAddField
          If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = "LCF"
        Next vAddField
        vFields = cptSortedArray(cptGetCustomFields("t", "Text,Outline Code", "c,cfn,loc", blnECF), 1)
        For lngItem = 0 To UBound(vFields)
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
          If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = vFields(lngItem, 2)
        Next lngItem
      ElseIf vControl = "EVP" Then
        For Each vAddField In Split("Physical % Complete,% Complete", ",")
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = FieldNameToFieldConstant(vAddField)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vAddField
          If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = "LCF"
        Next vAddField
        vFields = cptSortedArray(cptGetCustomFields("t", "Number", "c,cfn,loc", blnECF), 1)
        For lngItem = 0 To UBound(vFields)
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
          If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = vFields(lngItem, 2)
        Next lngItem
      ElseIf vControl = "LOE" Then
        On Error Resume Next
        If .cboLOE.ListCount > 0 Then .cboLOE.Value = strLOE
        If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        GoTo next_control
      ElseIf vControl = "PP" Then
        On Error Resume Next
        If .cboPP.ListCount > 0 Then .cboPP.Value = strPP
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        GoTo next_control
      Else 'WP,EVT,EVTMS
        vFields = cptSortedArray(cptGetCustomFields("t", "Text,Outline Code", "c,cfn,loc", blnECF), 1)
        For lngItem = 0 To UBound(vFields)
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
          If blnECF Then oComboBox.List(oComboBox.ListCount - 1, 2) = vFields(lngItem, 2)
        Next lngItem
        If IsEmpty(oComboBox.List(oComboBox.ListCount - 1, 0)) Then oComboBox.RemoveItem (oComboBox.ListCount - 1)
      End If
      If lngField > 0 And InStr(strSetting, "|") > 0 Then
        If lngField > 188776000 And FieldConstantToFieldName(lngField) = "<Unavailable>" Then 'this happens when it's an ECF; but offline
          MsgBox "The saved mapping field for element '" & vControl & "' is an Enterprise Custom Field (ECF) named '" & Split(strSetting, "|")(1) & "' but ECFs are only available when connected to PWA." & vbCrLf & vbCrLf & "Import of saved field mapping for '" & vControl & "' will be skipped.", vbExclamation + vbOKOnly, "Integration: " & vControl
        ElseIf InStr(strSetting, "|") > 0 Then
          If LCase(FieldConstantToFieldName(lngField)) <> LCase(Split(strSetting, "|")(1)) Then
            If LCase(CustomFieldGetName(lngField)) <> LCase(Split(strSetting, "|")(1)) Then
              If MsgBox("The saved mapping field for element '" & vControl & "' is named '" & Split(strSetting, "|")(1) & "' but in this file that field is named '" & CustomFieldGetName(lngField) & "'." & vbCrLf & vbCrLf & "Import saved mapping field anyway?", vbQuestion + vbYesNo, "Integration: " & vControl) = vbYes Then
                oComboBox.Value = lngField
              End If
            Else
              oComboBox.Value = lngField
            End If
          Else
            oComboBox.Value = lngField
          End If
        End If
      End If
next_control:
      If blnUseDefault Then
        oComboBox.Enabled = True
      Else
        oComboBox.Enabled = cptGetPosition(Replace(Replace(strRequiredFields, "[", ""), "]", ""), vControl, ",") > 0
      End If
      If oRequiredFields(vControl) Then
        If vControl = "LOE" Or vControl = "PP" Then
          If oComboBox.Value = "" Then
            oComboBox.BorderColor = 192
            blnValid = False
          End If
        Else
          If IsNull(oComboBox.Value) Then
            oComboBox.BorderColor = 192
            blnValid = False
          End If
        End If
      End If
      Set oComboBox = Nothing
    Next vControl
    
    .txtFiscalCalendar.Enabled = False
    .txtFiscalCalendar.Locked = True
    If cptCalendarExists("cptFiscalCalendar") Then
      .txtFiscalCalendar = "cptFiscalCalendar"
    Else
      If blnFiscalRequired Then
        .txtFiscalCalendar.BorderColor = 192
        blnValid = False
      End If
    End If
    
    strSetting = cptGetSetting("Integration", "RollingWaveDate")
    If Len(strSetting) > 0 Then
      .txtRollingWave = FormatDateTime(strSetting, vbShortDate)
      .lblWeekday.Caption = Format(CDate(.txtRollingWave), "dddd")
      .lblWeekday.Visible = True
    End If
    
    If blnRollingWaveDateRequired Then
      .txtRollingWave.Enabled = True
      .txtRollingWave.BorderColor = 192
      If IsDate(.txtRollingWave) Then
        .txtRollingWave.BorderColor = -2147483642
      Else
        blnValid = False
      End If
    Else
      .txtRollingWave.Enabled = True
    End If
    
    If cptModuleExists("cptIMSCobraExport_bas") And cptModuleExists("cptIMSCobraExport_frm") Then
      .chkSyncSettings.Enabled = True
      strSetting = cptGetSetting("Integration", "chkSyncSettings")
      If Len(strSetting) > 0 Then
        .chkSyncSettings = CBool(strSetting)
      Else
        .chkSyncSettings = True 'default
      End If
    Else
      .chkSyncSettings.Enabled = False
    End If
        
    If Not blnValid Or blnConfirmationRequired Then
      .Show
      cptValidMap = .blnValidIntegrationMap
    Else
      cptValidMap = blnValid
    End If
    
  End With

exit_here:
  On Error Resume Next
  Set oRequiredFields = Nothing
  Set oComboBox = Nothing
  Unload myIntegration_frm
  Set myIntegration_frm = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptValidMap", Err, Erl)
  Resume exit_here
    
End Function

Function cptValidPath(strFullName As String) As String
  'usage: strValidPath = cptValidPath(strFullName)
  'usage: if not cbool(Split(strValidPath,":")(0)) then
  'usage:   msgbox replace(strValidPath,"0:","False:")
  'usage:   goto exit_here
  'usage: end if
  'objects
  Dim oFSO As Scripting.FileSystemObject
  Dim oFolder As Scripting.Folder
  'strings
  Dim strDir As String
  Dim strFileName As String
  Dim strInvalidCharacters As String
  Dim strReason As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnValid As Boolean
  'variants
  Dim v As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'ensure folder exists
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  On Error Resume Next
  Set oFolder = oFSO.GetFolder(oFSO.GetParentFolderName(strFullName))
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oFolder Is Nothing Then
    strReason = "folder does not exist"
    GoTo exit_here
  End If
'  'ensure folder is not read-only
  'todo: this returns false positive on some systems...(?)
'  If (oFolder.Attributes And ReadOnly) = 1 Then
'    strReason = "folder is read-only"
'    GoTo exit_here
'  End If
  'test for illegal characters in filename
  strFileName = Replace(strFullName, oFSO.GetParentFolderName(strFullName), "")
  For Each v In Split("<,>,?,[,],:,|,*", ",")
    If Len(cptRegEx(strFileName, "[\" & v & "]")) > 0 Then
      strInvalidCharacters = strInvalidCharacters & v & " "
    End If
  Next v
  If Len(strInvalidCharacters) > 0 Then
    If UBound(Split(strInvalidCharacters, " ")) > 1 Then
      strReason = "invalid characters in filename: " & strInvalidCharacters
    Else
      strReason = "invalid character in filename: " & Split(strInvalidCharacters, " ")(0)
    End If
    GoTo exit_here
  End If
  'test for length
  If Len(strFullName) > 218 Then
    strReason = "exceeds 218 characters"
    'todo: then cptGetShortPath?
    GoTo exit_here
  End If
  
exit_here:
  On Error Resume Next
  If strReason <> "" Then
    cptValidPath = "0: " & strReason
  Else
    cptValidPath = "1: "
  End If
  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptValidPath()", Err, Erl)
  Resume exit_here
End Function

Function cptGetShortPath(strLongPath As String) As String
  'if strLongPath file exists: shortpath + shortfilename
  'if strLongPath file not exists: shortpath + longfilename
  'if strLongPath has no file: shortpath
  'objects
  Dim oFSO As Scripting.FileSystemObject
  Dim oFile As Scripting.File
  'strings
  Dim strShortPath As String
  Dim strFileName As String
  Dim strFolderName As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
    
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  
  On Error Resume Next
  Set oFile = oFSO.GetFile(strLongPath)
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oFile Is Nothing Then
    strFileName = oFile.ShortName
    strLongPath = oFSO.GetParentFolderName(strLongPath)
    strShortPath = oFSO.GetFolder(strLongPath).ShortPath
    strShortPath = strShortPath & "\" & oFile.ShortName
  Else
    If Len(oFSO.GetExtensionName(strLongPath)) > 0 Then
      strFileName = oFSO.GetFileName(strLongPath)
      strShortPath = oFSO.GetParentFolderName(strLongPath)
      strShortPath = oFSO.GetFolder(strShortPath).ShortPath
      strShortPath = strShortPath & "\" & strFileName
    Else
      strShortPath = oFSO.GetFolder(strLongPath).ShortPath
    End If
  End If
    
exit_here:
  On Error Resume Next
  Set oFSO = Nothing
  Set oFile = Nothing
  cptGetShortPath = strShortPath
  Exit Function
  
err_here:
  On Error Resume Next
  cptHandleErr "cptCore_bas", "cptGetShortPath", Err, Erl
  strShortPath = ""
  Resume exit_here

End Function

Function cptErrorTrapping() As Boolean
  'objects
  'strings
  Dim strErrorTrapping As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  On Error GoTo err_here 'some users experiencing error on recursive call

  strErrorTrapping = cptGetSetting("General", "ErrorTrapping")
  If Len(strErrorTrapping) > 0 Then
    cptErrorTrapping = CBool(strErrorTrapping)
  Else
    cptSaveSetting "General", "ErrorTrapping", "1"
    cptErrorTrapping = True
  End If

exit_here:
  On Error Resume Next
  Call cptCore_bas.cptStartEvents
  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptErrorTrapping", Err, Erl)
  Resume exit_here
End Function

Function cptConvertFilteredRecordset(oRecordset As Object) As Object 'ADODB.Recordset,ADODB.Recordset
  Dim oStream As Object 'ADODB.Stream
  Dim oFiltered As Object 'ADODB.Recordset
  Set oStream = CreateObject("ADODB.Stream")
  oRecordset.Save oStream, adPersistXML
  Set oFiltered = CreateObject("ADODB.Recordset")
  oFiltered.Open oStream
  Set cptConvertFilteredRecordset = oFiltered
  oStream.Close
  Set oStream = Nothing
End Function

Function cptSortedArray(vArray As Variant, lngSortKey As Long) As Variant
  'strFieldTypes  := comma-separated list of any of "p,t,r" [project,task,resource]
  'strDataTypes   := comma-separated list of any of "Cost,Date,Duration,Flag,Finish,Number,Start,Text,Outline Code"
  'strInclude     := comma-separated list of any of "c,fn,cfn" [constant,fieldname,customfieldname]
  'blnIncludeEnterprise := self-explanatory

  'read array into oDict using key as key, row as row
  'sort the sort keys
  'build sorted array off of keys using lookup
  Dim oDict As Object 'Scripting.Dictionary
  Dim lngRows As Long, lngCols As Long
  Dim lngRow As Long, lngCol As Long
  Dim strRow As String
  Dim vTemp() As Variant
  Dim vSorted As Variant
  
  Set oDict = CreateObject("Scripting.Dictionary")
  lngRows = UBound(vArray)
  lngCols = UBound(vArray, 2)
  For lngRow = 0 To lngRows
    strRow = ""
    For lngCol = 0 To lngCols
      strRow = strRow & Replace(vArray(lngRow, lngCol), ",", "<cpt>") & "," 'replace commas with unique string
    Next lngCol
    strRow = Left(strRow, Len(strRow) - 1)
    If Not oDict.Exists(vArray(lngRow, lngSortKey)) Then
      oDict.Add vArray(lngRow, lngSortKey), strRow
    Else
      lngRows = lngRows - 1
    End If
  Next lngRow
    
  vSorted = oDict.Keys()
  cptQuickSort vSorted, 0, UBound(vSorted)

  ReDim vTemp(0 To lngRows, 0 To lngCols)
  For lngRow = 0 To lngRows
    For lngCol = 0 To lngCols
      vTemp(lngRow, lngCol) = Replace(Split(oDict(vSorted(lngRow)), ",")(lngCol), "<cpt>", ",") 'replace unique strings with comma
    Next lngCol
  Next lngRow
  
  cptSortedArray = vTemp
  
  Set oDict = Nothing
End Function

Function cptGetPosition(vList As Variant, vValue As Variant, Optional strDelimiter As String) As Long
  'find position of a value in a list
  'accepts comma-separated lists or arrays
  'if vList is comma-separated then strDelimiter is required
  'boolean type is not supported
  Dim lngPosition As Long
  Dim lngTemp As Long
  Dim vPosition As Variant
  
  If IsArray(vList) Then
    For lngPosition = 1 To UBound(vList) + 1
      Select Case TypeName(vValue)
        Case "Date"
          If vValue = CDate(vList(lngPosition - 1)) Then
            lngTemp = lngPosition
          End If
        Case "Double"
          If vValue = CDbl(vList(lngPosition - 1)) Then
            lngTemp = lngPosition
          End If
        Case "Integer"
          If vValue = CInt(vList(lngPosition - 1)) Then
            lngTemp = lngPosition
          End If
        Case "Long"
          If vValue = CLng(vList(lngPosition - 1)) Then
            lngTemp = lngPosition
          End If
        Case "String"
          If vValue = vList(lngPosition - 1) Then
            lngTemp = lngPosition
          End If
        Case "Null"
          If vList(lngPosition) = "" Then
            lngTemp = lngPosition
          End If
      End Select
    Next lngPosition
  Else 'delimited string
    lngPosition = 1
    For Each vPosition In Split(vList, strDelimiter)
      If vValue = vPosition Then
        lngTemp = lngPosition
        Exit For
      End If
      lngPosition = lngPosition + 1
    Next vPosition
  End If
  cptGetPosition = lngTemp
End Function

Function cptGetDate(dtDate As Date, Optional strFormat As String)
  Dim vFormat As Variant
  For Each vFormat In Array(vbGeneralDate, vbLongDate, vbShortDate, vbLongTime, vbShortTime)
    Debug.Print vFormat & ": " & FormatDateTime(dtDate, vFormat)
  Next vFormat
  If Len(strFormat) > 0 Then
    Debug.Print "custom: " & Format(dtDate, strFormat)
  End If
End Function

Function cptCustomFieldExists(strCustomFieldName As String, Optional blnResourceField As Boolean = False) As Variant
  'returns 0 if false; constant if true
  Dim lngCFC As Long
  On Error Resume Next
  If blnResourceField Then
    lngCFC = FieldNameToFieldConstant(strCustomFieldName, pjResource)
  Else
    lngCFC = FieldNameToFieldConstant(strCustomFieldName, pjTask)
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptCustomFieldExists = lngCFC
err_here:
End Function

Function cptCalendarExists(strCalendar As String) As Boolean
  Dim oCalendar As MSProject.Calendar
  Dim strMsg As String
  
  On Error Resume Next
  Set oCalendar = ActiveProject.BaseCalendars(strCalendar)
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oCalendar Is Nothing Then
    cptCalendarExists = False
  Else
    If oCalendar.Exceptions.Count = 0 Then
      strMsg = "cptFiscalCalendar exists but has no exceptions." & vbCrLf & vbCrLf
      strMsg = strMsg & "Please rebuild it (ClearPlan > Calendars > Fiscal)."
      MsgBox strMsg, vbCritical + vbOKOnly, "No Exceptions"
      oCalendar.Delete
      cptCalendarExists = False
    Else
      cptCalendarExists = True
    End If
  End If
  
exit_here:
  On Error Resume Next
  Set oCalendar = Nothing
  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptCalendarExists", Err, Erl)
  Resume exit_here
End Function

Sub cptCheckMetadata(strConstants As String, strReturn As String)
  'objects
  Dim oTask As MSProject.Task
  Dim oDict As Scripting.Dictionary
  'strings
  Dim strMissing As String
  'longs
  Dim lngCount As Long
  'integers
  'doubles
  'booleans
  Dim blnErrorTrapping As Boolean
  'variants
  Dim vCF As Variant
  'dates
  
  blnErrorTrapping = cptErrorTrapping
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'todo: should strConstants include things like BLS,BLF,BLD?
  'todo: should strConstants include assignments?
  
  Set oDict = CreateObject("Scripting.Dictionary")
  For Each vCF In Split(strConstants, ",")
    oDict.Add vCF, "0|"
  Next vCF
  
  'assume alignment between master and subs
  'if they're not, this will make it clear
  If ActiveProject.Subprojects.Count > 0 Then
    ActiveWindow.TopPane.Activate
    If ActiveWindow.TopPane.View.Type <> pjTaskItem Then ViewApply "Gantt Chart"
    OptionsViewEx DisplaySummaryTasks:=True
    OutlineShowTasks pjTaskOutlineShowLevel2
  End If
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If oTask.Assignments.Count > 0 And (oTask.BaselineWork > 0 Or oTask.BaselineCost > 0) Then
      For Each vCF In Split(strConstants, ",")
        If vCF = "" Then GoTo next_cf
        If Len(oTask.GetField(Split(vCF, "|")(2))) = 0 Then
          lngCount = CLng(Split(oDict(vCF), "|")(0))
          strMissing = Split(oDict(vCF), "|")(1)
          lngCount = lngCount + 1
          strMissing = strMissing & oTask.UniqueID & ","
          oDict(vCF) = lngCount & "|" & strMissing
        End If
next_cf:
      Next vCF
    End If
next_task:
  Next oTask
  
  lngCount = 0
  strMissing = "Results:" & vbCrLf
  For Each vCF In Split(strConstants, ",")
    If vCF <> "" Then
      lngCount = lngCount + CLng(Split(oDict(vCF), "|")(0))
      strMissing = strMissing & "[" & Split(vCF, "|")(0) & "] " & CustomFieldGetName(Split(vCF, "|")(2)) & ": " & Split(oDict(vCF), "|")(0) & " missing" & vbCrLf
    End If
  Next vCF
  If strReturn = "lngCount" Then
    MsgBox lngCount, vbExclamation + vbOKOnly, "PMB Task Metadata Check"
  ElseIf strReturn = "strMissing" Then
    MsgBox strMissing, vbExclamation + vbOKOnly, "PMB Task Metadata Check"
  End If
  
exit_here:
  On Error Resume Next
  Set oTask = Nothing
  Set oDict = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptCheckMetadata", Err, Erl)
  Resume exit_here
End Sub

Sub cptAddBorders(ByRef rng As Object, Optional blnHorizontal As Boolean = True)
  rng.Borders(xlDiagonalDown).LineStyle = xlNone
  rng.Borders(xlDiagonalUp).LineStyle = xlNone
  With rng.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ThemeColor = 2
    .TintAndShade = 0.499984740745262
    .Weight = xlThin
  End With
  With rng.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ThemeColor = 2
    .TintAndShade = 0.499984740745262
    .Weight = xlThin
  End With
  With rng.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ThemeColor = 2
    .TintAndShade = 0.499984740745262
    .Weight = xlThin
  End With
  With rng.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ThemeColor = 2
    .TintAndShade = 0.499984740745262
    .Weight = xlThin
  End With
  With rng.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946592608417
    .Weight = xlThin
  End With
  'optional horizontal lines
  If blnHorizontal Then
    rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rng.Borders(xlInsideHorizontal).ThemeColor = 1
    rng.Borders(xlInsideHorizontal).TintAndShade = -0.249946592608417
    rng.Borders(xlInsideHorizontal).Weight = xlThin
  Else
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
  End If
End Sub

Sub cptAddShading(ByRef oRange As Object, Optional blnLight = False)
  If blnLight Then
    With oRange.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -4.99893185216834E-02
      .PatternTintAndShade = 0
    End With
  Else
    With oRange.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -0.149998474074526
      .PatternTintAndShade = 0
    End With
  End If
End Sub

Function cptGetConstantName(strProperty As String, lngValue As Long) As String
  Dim strConstantName As String
  Select Case strProperty
    Case "CurrencySymbolPosition"
      strConstantName = Choose(lngValue + 1, _
                  "0 - pjBefore", _
                  "1 - pjAfter", _
                  "2 - pjBeforeWithSpace", _
                  "3 - pjAfterWithSpace")
    Case "DefaultDurationUnits"
      Select Case lngValue
        Case 3
          strConstantName = "3 - pjMinute"
        Case 5
          strConstantName = "5 - pjHour"
        Case 7
          strConstantName = "7 - pjDay"
        Case 9
          strConstantName = "9 - pjWeek"
        Case 11
          strConstantName = "11 - pjMonthUnit"
      End Select
    Case "EarnedValueBaseline"
      strConstantName = lngValue & " - pjBaseline" & IIf(lngValue > 0, lngValue, "")
    Case "TrackingMethod"
      strConstantName = Choose(lngValue + 1, _
                  "0 - pjTrackingMethodDefault", _
                  "1 - pjTrackingMethodSpecifyHours", _
                  "2 - pjTrackingMethodPercentComplete", _
                  "3 - pjTrackingMethodTotalAndRemaining")
    Case "Type"
      strConstantName = Choose(lngValue + 1, _
                  "0 - pjProjectTypeUnsaved", _
                  "1 - pjProjectTypeNonEnterprise", _
                  "2 - pjProjectTypeEnterpriseCheckedOut", _
                  "3 - pjProjectTypeEnterpriseReadOnly", _
                  "4 - pjProjectTypeEnterpriseGlobalCheckedOut", _
                  "5 - pjProjectTypeEnterpriseGlobalInMemory", _
                  "6 - pjProjectTypeEnterpriseGlobalLocal", _
                  "7 - pjProjectTypeEnterpriseResourcesCheckedOut", _
                  "8 - pjProjectTypeBasicProjectSite")
    Case "DefaultDateFormat"
      If lngValue = 255 Then
        strConstantName = "255 - pjDateDefault"
      Else
        strConstantName = Choose(lngValue + 1, _
                    "0 - pjDate_mm_dd_yy_hh_mmAM", _
                    "1 - pjDate_mm_dd_yy", _
                    "2 - pjDate_mmmm_dd_yyyy_hh_mmAM", _
                    "3 - pjDate_mmmm_dd_yyyy", _
                    "4 - pjDate_mmm_dd_hh_mmAM", _
                    "5 - pjDate_mmm_dd_yyy", _
                    "6 - pjDate_mmmm_dd", _
                    "7 - pjDate_mmm_dd", _
                    "8 - pjDate_ddd_mm_dd_yy_hh_mmAM", _
                    "9 - pjDate_ddd_mm_dd_yy", _
                    "10 - pjDate_ddd_mmm_dd_yyy", _
                    "11 - pjDate_ddd_hh_mmAM", _
                    "12 - pjDate_mm_dd", _
                    "13 - pjDate_dd", _
                    "14 - pjDate_hh_mmAM", _
                    "15 - pjDate_ddd_mmm_dd", _
                    "16 - pjDate_ddd_mm_dd", _
                    "17 - pjDate_ddd_dd", _
                    "18 - pjDate_Www_dd", _
                    "19 - pjDate_Www_dd_yy_hh_mmAM", _
                    "20 - pjDate_mm_dd_yyyy")
      End If
    Case "LevelPeriodBasis"
      strConstantName = Choose(lngValue + 1, _
                  "0 - pjMinuteByMinute", _
                  "1 - pjHourByHour", _
                  "2 - pjDayByDay", _
                  "3 - pjWeekByWeek", _
                  "4 - pjMonthByMonth")
    Case "StartWeekOn"
      strConstantName = Choose(lngValue, _
                  "1 - pjSunday", _
                  "2 - pjMonday", _
                  "3 - pjTuesday", _
                  "4 - pjWednesday", _
                  "5 - pjThursday", _
                  "6 - pjFriday", _
                  "7 - pjSaturday")
                  
    Case "StartYearIn"
      strConstantName = Choose(lngValue, _
                  "1 - pjJanuary", _
                  "2 - pjFebruary", _
                  "3 - pjMarch", _
                  "4 - pjApril", _
                  "5 - pjMay", _
                  "6 - pjJune", _
                  "7 - pjJuly", _
                  "8 - pjAugust", _
                  "9 - pjSeptember", _
                  "10 - pjOctober", _
                  "11 - pjNovember", _
                  "12 - pjDecember")
                  
    Case "WorkContour"
      strConstantName = Choose(lngValue + 1, _
                  "0 - pjFlat", _
                  "1 - pjBackLoaded", _
                  "2 - pjFrontLoaded", _
                  "3 - pjDoublePeak", _
                  "4 - pjEarlyPeak", _
                  "5 - pjLatePeak", _
                  "6 - pjBell", _
                  "7 - pjTurtle", _
                  "8 - pjContour")
    Case Else
      strConstantName = ""
  End Select
  cptGetConstantName = strConstantName
End Function

Function cptHasFormula(ByRef oProject As MSProject.Project, lngCFC As Long) As Boolean
  oProject.Activate
  cptHasFormula = Len(CustomFieldGetFormula(lngCFC)) > 0
End Function

Function cptHasLookup(ByRef oProject As MSProject.Project, lngCFC As Long) As Boolean
  Dim strFN As String
  Dim strCFN As String
  Dim blnHasLookup As Boolean
  Dim blnErrorTrapping As Boolean
  Dim oOutlineCode As MSProject.OutlineCode
  
  blnErrorTrapping = cptErrorTrapping
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  oProject.Activate
  strFN = FieldConstantToFieldName(lngCFC)
  strCFN = CustomFieldGetName(lngCFC)
  If InStr(strFN, "Outline Code") > 0 Then
    On Error Resume Next
    Set oOutlineCode = oProject.OutlineCodes(strCFN)
    blnHasLookup = Not oOutlineCode Is Nothing
  Else
    On Error Resume Next
    blnHasLookup = Len(CustomFieldValueListGetItem(lngCFC, pjValueListValue, 1)) > 0
  End If
  If blnErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
exit_here:
  On Error Resume Next
  cptHasLookup = blnHasLookup
  Set oOutlineCode = Nothing
  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptHasLookup()", Err, Erl)
  Resume exit_here
End Function

