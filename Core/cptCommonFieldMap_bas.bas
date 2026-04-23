Attribute VB_Name = "cptCommonFieldMap_bas"
'<cpt_version>v1.0.2</cpt_version>
Option Explicit

#If VBA7 Then

    Private Declare PtrSafe Function GetPrivateProfileStringA Lib "Kernel32" (ByVal strSection As String, _
        ByVal strKey As String, ByVal strDefault As String, _
        ByVal strReturnedString As String, _
        ByVal lngSize As Long, ByVal strFileNameName As String) As Long
        
    Private Declare PtrSafe Function WritePrivateProfileStringA Lib _
        "Kernel32" (ByVal strSection As String, _
        ByVal strKey As String, ByVal strString As String, _
        ByVal strFileNameName As String) As Long

#Else

    Private Declare Function GetPrivateProfileStringA Lib "Kernel32" (ByVal strSection As String, _
        ByVal strKey As String, ByVal strDefault As String, _
        ByVal strReturnedString As String, _
        ByVal lngSize As Long, ByVal strFileNameName As String) As Long
        
    Private Declare Function WritePrivateProfileStringA Lib _
        "Kernel32" (ByVal strSection As String, _
        ByVal strKey As String, ByVal strString As String, _
        ByVal strFileNameName As String) As Long
    
#End If

Function cptGetCustomFieldName(ByVal FieldName As String) As String

    Dim settingsFile As String
    
    settingsFile = GetSettingsFile

    cptGetCustomFieldName = GetPrivateProfileString(settingsFile, FieldName, "Name")
    
End Function

Function cptGetCustomFieldGUID(ByVal FieldName As String) As String

    Dim settingsFile As String
    
    settingsFile = GetSettingsFile

    cptGetCustomFieldGUID = GetPrivateProfileString(settingsFile, FieldName, "GUID")
    
End Function

Sub cptStoreCustomFieldName(ByVal FieldName As String, MSP_FieldName As String, MSP_FieldGUID As String)

    On Error GoTo ErrorHandler

    Dim settingsFile As String
    
    settingsFile = GetSettingsFile
    
    If Not (WritePrivateProfileString(settingsFile, FieldName, "Name", MSP_FieldName)) Then
        err.Raise 1
    End If
    
    If Not (WritePrivateProfileString(settingsFile, FieldName, "GUID", MSP_FieldGUID)) Then
        err.Raise 2
    End If
    
    Exit Sub
    
ErrorHandler:

    Select Case err.Number
    
        Case 1
            err.Description = "Error setting Custom Field Name value."
        Case 2
            err.Description = "Error setting Custom Field GUID value."
        Case Else
            err.Description = "Error storing Custom Field information."
    
    End Select
    
End Sub

Private Function GetSettingsFile() As String

    Dim cptSettingsFilePath As String
    Dim UserProfilePath As String
    Dim cptSettingsFolderPath As String
    
    UserProfilePath = Environ$("USERPROFILE")
    
    cptSettingsFolderPath = UserProfilePath & "\cpt-backup\settings\"
    cptSettingsFilePath = cptSettingsFolderPath & "cpt-settings.ini"
    
    If VBA.FileSystem.Dir(cptSettingsFolderPath, vbDirectory) = "" Then
        CreateSettingsDirectory (cptSettingsFolderPath)
        CreateSettingsFile (cptSettingsFilePath)
    ElseIf VBA.FileSystem.Dir(cptSettingsFilePath) = "" Then
        CreateSettingsFile (cptSettingsFilePath)
    End If
    
    GetSettingsFile = cptSettingsFilePath

End Function

Private Sub CreateSettingsDirectory(ByVal directoryToMake As String)
    
    MkDir directoryToMake
    
End Sub

Private Sub CreateSettingsFile(ByVal fileToMake As String)
    Dim fs As Object
    Dim a As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(fileToMake)
    a.Close
    
    Set fs = Nothing
    Set a = Nothing

End Sub

Private Function WritePrivateProfileString(ByVal strFileName As String, _
    ByVal strSection As String, ByVal strKey As String, _
    ByVal strValue As String) As Boolean
    
    Dim lngValid As Long
    On Error Resume Next
    lngValid = WritePrivateProfileStringA(strSection, strKey, _
        strValue, strFileName)
    If lngValid > 0 Then WritePrivateProfileString = True
    
    On Error GoTo 0
    
End Function

Private Function GetPrivateProfileString(ByVal strFileName As String, _
    ByVal strSection As String, ByVal strKey As String, _
    Optional strDefault) As String
    
    Dim strReturnString As String, lngSize As Long, lngValid As Long
    On Error Resume Next
    If IsMissing(strDefault) Then strDefault = ""
    strReturnString = Space(1024)
    lngSize = Len(strReturnString)
    lngValid = GetPrivateProfileStringA(strSection, strKey, _
        strDefault, strReturnString, lngSize, strFileName)
    GetPrivateProfileString = Left(strReturnString, lngValid)
    
    On Error GoTo 0
    
End Function
