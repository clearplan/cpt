Attribute VB_Name = "cptBrowseFolder_bas"
'<cpt_version>v2.0.0</cpt_version>
Option Explicit

#If VBA7 Then

Private Type OPENFILENAMEW
    lStructSize       As Long
    hwndOwner         As LongPtr
    hInstance         As LongPtr
    lpstrFilter       As LongPtr
    lpstrCustomFilter As LongPtr
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As LongPtr
    nMaxFile          As Long
    lpstrFileTitle    As LongPtr
    nMaxFileTitle     As Long
    lpstrInitialDir   As LongPtr
    lpstrTitle        As LongPtr
    Flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As LongPtr
    lCustData         As LongPtr
    lpfnHook          As LongPtr
    lpTemplateName    As LongPtr
    pvReserved        As LongPtr
    dwReserved        As Long
    FlagsEx           As Long
End Type

Private Declare PtrSafe Function cptGetFilesW _
    Lib "comdlg32.dll" _
    Alias "GetOpenFileNameW" ( _
        ByRef pOpenfilename As OPENFILENAMEW) As Long

#Else

Private Type OPENFILENAMEW
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As Long
    lpstrCustomFilter As Long
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As Long
    nMaxFile          As Long
    lpstrFileTitle    As Long
    nMaxFileTitle     As Long
    lpstrInitialDir   As Long
    lpstrTitle        As Long
    Flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As Long
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As Long
    pvReserved        As Long
    dwReserved        As Long
    FlagsEx           As Long
End Type

Private Declare Function cptGetFilesW _
    Lib "comdlg32.dll" _
    Alias "GetOpenFileNameW" ( _
        ByRef pOpenfilename As OPENFILENAMEW) As Long

#End If

Private Const OFN_ALLOWMULTISELECT As Long = &H200&
Private Const OFN_FILEMUSTEXIST    As Long = &H1000&
Private Const OFN_PATHMUSTEXIST    As Long = &H800&
Private Const OFN_EXPLORER         As Long = &H80000
Private Const OFN_HIDEREADONLY     As Long = &H4&
Private Const OFN_NOCHANGEDIR      As Long = &H8&

Public Function cptGetFiles( _
    Optional ByVal szDialogTitle As String = "Select file", _
    Optional ByVal szInitialFolder As String = vbNullString, _
    Optional ByVal szFilter As String = "All Files (*.*)" & vbNullChar & "*.*", _
    Optional ByVal blnAllowMultiSelect As Boolean = False, _
    Optional ByVal hwndOwner As LongPtr = 0 _
) As Collection

  'usage:
  Const BUFFER_CHARS As Long = 65536

  Dim ofn As OPENFILENAMEW
  Dim szBuffer As String
  Dim szResults As String
  Dim astrParts() As String
  Dim colResults As New Collection

  Dim lngResult As Long
  Dim lngEnd As Long
  Dim lngIndex As Long
  Dim szFolder As String
  'sz=zero-terminated string for Win32 API calls; different than VBA String "proper"

  szBuffer = String$(BUFFER_CHARS, vbNullChar)

  ' GetOpenFileNameW requires:
  ' description + NUL + pattern + NUL + final NUL
  'example: Filter = _
  '    "Supported Files (*.csv;*.xlsx;*.xlsm)" & vbNullChar & _
  '    "*.csv;*.xlsx;*.xlsm" & vbNullChar & _
  '    "CSV Files (*.csv)" & vbNullChar & _
  '    "*.csv" & vbNullChar & _
  '    "Excel Files (*.xlsx;*.xlsm)" & vbNullChar & _
  '    "*.xlsx;*.xlsm" & vbNullChar & _
  '    "All Files (*.*)" & vbNullChar & _
  '    "*.*" & vbNullChar & _
  '    vbNullChar
  
  szFilter = szFilter & vbNullChar & vbNullChar

  With ofn
    .lStructSize = LenB(ofn)
    .hwndOwner = hwndOwner
    .lpstrFilter = StrPtr(szFilter)
    .nFilterIndex = 1
    .lpstrFile = StrPtr(szBuffer)
    .nMaxFile = Len(szBuffer)
    .lpstrTitle = StrPtr(szDialogTitle)

    If LenB(szInitialFolder) <> 0 Then
      .lpstrInitialDir = StrPtr(szInitialFolder)
    End If

    .Flags = OFN_EXPLORER Or _
             OFN_FILEMUSTEXIST Or _
             OFN_PATHMUSTEXIST Or _
             OFN_HIDEREADONLY Or _
             OFN_NOCHANGEDIR

    If blnAllowMultiSelect Then
      .Flags = .Flags Or OFN_ALLOWMULTISELECT
    End If
  End With

  lngResult = cptGetFilesW(ofn)

  If lngResult = 0 Then
    Set cptGetFiles = New Collection
    Exit Function
  End If

  lngEnd = InStr(1, szBuffer, vbNullChar & vbNullChar, vbBinaryCompare)

  If lngEnd > 0 Then
    szResults = Left$(szBuffer, lngEnd - 1)
  Else
    szResults = Left$(szBuffer, InStr(1, szBuffer, vbNullChar) - 1)
  End If

  astrParts = Split(szResults, vbNullChar)

  If UBound(astrParts) = 0 Then
    colResults.Add astrParts(0)
  Else
    szFolder = astrParts(0)
    If Right$(szFolder, 1) <> "\" Then
      szFolder = szFolder & "\"
    End If
    For lngIndex = 1 To UBound(astrParts)
      If LenB(astrParts(lngIndex)) <> 0 Then
        colResults.Add szFolder & astrParts(lngIndex)
      End If
    Next lngIndex
  End If

  Set cptGetFiles = colResults

End Function

Public Function cptGetFolder(strTitle As String, strDefaultPath As String) As String
  'usage:
  'If LenB(BrowseForFolder()) Then...
  Dim oFolder As Object
  Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0, "[title]", 0, strDefaultPath)

  If Not oFolder Is Nothing Then
    On Error Resume Next
    cptGetFolder = oFolder.Self.Path
    On Error GoTo 0
  End If

  If Len(cptGetFolder) > 0 Then
    If Left$(cptGetFolder, 2) = "\\" Or Mid$(cptGetFolder, 2, 1) = ":" Then
      Exit Function
    End If
    cptGetFolder = vbNullString
  End If
End Function
