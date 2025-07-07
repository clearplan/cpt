Attribute VB_Name = "cptAdvancedFilter_bas"
'<cpt_version>v0.2.0</cpt_version>
Option Explicit

Private Const MODULE_NAME As String = "cptAdvancedFilter_bas"
Private filterForm As cptAdvancedFilter_frm
Private curProj As Project
Private CustTextFields() As String
Private EntFields() As String
Private CustNumFields() As String
Private CustOLCodeFields() As String

Sub cptAdvancedFilter()
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Set curProj = ActiveProject
    
    filterReadCustomFields curProj

    Set filterForm = New cptAdvancedFilter_frm
    
    With filterForm
    
        Dim vArray As Variant
        vArray = Split(Join(CustTextFields, ",") & "," & Join(CustNumFields, ",") & "," & Join(CustOLCodeFields, ",") & "," & Join(EntFields, ","), ",")
        If vArray(UBound(vArray)) = "" Then ReDim Preserve vArray(UBound(vArray) - 1)
        cptQuickSort vArray, 0, UBound(vArray)
        .fltrField.List = Split("UniqueID,ID,Name," & Join(vArray, ","), ",")
        .fltrField.ListIndex = 0
        .versionLbl = "Advanced Clipboard Filter"
        .Caption = .versionLbl.Caption & " " & cptGetVersion("cptAdvancedFilter_bas")
        .summaryCheckBox = curProj.Application.SummaryTasksShow
    
        .Show
    
    End With
    
    Exit Sub
    
ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "cptAdvancedFilter", err, Erl, "Error initializing Clipboard Filter")
    'MsgBox "Error initializing Clipboard Filter: " & err.Description, vbCritical, "Clipboard Filter Error"

End Sub

Public Sub cptExitAdvancedFilter()
    curProj.Application.FilterApply Name:="All Tasks"
    curProj.Application.FilterClear
    Set filterForm = Nothing
    Set curProj = Nothing
    Exit Sub
End Sub

Public Sub setFilter(ByRef filterItemsList As Collection, ByVal caseSensitive As Boolean)
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0

    If curProj.AutoFilter = False Then curProj.AutoFilter = True
    curProj.Application.FilterApply Name:="All Tasks"
    curProj.Application.FilterClear

    Dim t As Task
    Dim tempFilterString As String
    Dim cntr As Integer
    Dim testValue As String
    
    For Each t In curProj.Tasks
    
        If Not t Is Nothing Then
        
            testValue = t.GetField(FieldNameToFieldConstant(filterForm.fltrField.Text))
            
            For cntr = 1 To filterItemsList.Count
            
                If filterItemsList(cntr).Method = "Equals" Then
                    If caseSensitive Then
                        If filterItemsList(cntr).Value = testValue Then
                            If tempFilterString = "" Then
                                tempFilterString = t.UniqueID
                            Else
                                tempFilterString = tempFilterString & Chr$(9) & t.UniqueID
                            End If
                            
                            GoTo NextTask
                            
                        End If
                    Else
                        If LCase(filterItemsList(cntr).Value) = LCase(testValue) Then
                            If tempFilterString = "" Then
                                tempFilterString = t.UniqueID
                            Else
                                tempFilterString = tempFilterString & Chr$(9) & t.UniqueID
                            End If
                            
                            GoTo NextTask
                            
                        End If
                    End If
                Else
                    If caseSensitive Then
                        If InStr(1, testValue, filterItemsList(cntr).Value, vbBinaryCompare) > 0 Then
                            If tempFilterString = "" Then
                                tempFilterString = t.UniqueID
                            Else
                                tempFilterString = tempFilterString & Chr$(9) & t.UniqueID
                            End If
                            
                            GoTo NextTask
                        
                        End If
                    Else
                        If InStr(1, testValue, filterItemsList(cntr).Value, vbTextCompare) > 0 Then
                            If tempFilterString = "" Then
                                tempFilterString = t.UniqueID
                            Else
                                tempFilterString = tempFilterString & Chr$(9) & t.UniqueID
                            End If
                            
                            GoTo NextTask
                        
                        End If
                    End If
                End If
            
            Next cntr
        
        End If
        
NextTask:

    Next t
    
    If tempFilterString <> "" Then
    
        Application.SetAutoFilter FieldName:="Unique ID", FilterType:=pjAutoFilterIn, Criteria1:=tempFilterString
    Else
        
        MsgBox "There are no matching results.", vbOKOnly + vbInformation, "No Results Found"
    
    End If
    
    SelectBeginning 'select first cell in table (top left)
    
    Exit Sub
    
ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "setFilter", err, Erl, "Error setting AutoFilter")
    'MsgBox "Error setting AutoFilter: " & err.Description, vbCritical, "AutoFilter Error"

End Sub

Public Function GetClipboardText() As String
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Dim dataObj As Object
    Dim clipText As String
    
    ' Try using MSForms DataObject first (more reliable)
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.GetFromClipboard
    
    If dataObj.GetFormat(1) Then ' CF_TEXT format
        clipText = dataObj.GetText(1)
    ElseIf dataObj.GetFormat(13) Then ' CF_UNICODETEXT format
        clipText = dataObj.GetText(13)
    Else
        clipText = ""
    End If
    
    ' Clean up line endings
    clipText = Replace(Replace(clipText, vbCrLf, vbLf), vbCr, vbLf)
    
    GetClipboardText = clipText
    Exit Function
    
ErrorHandler:
    ' Fallback method
    On Error Resume Next
    Set dataObj = CreateObject("HTMLFile")
    Set dataObj = Nothing
    Set dataObj = CreateObject("Forms.DataObject.1")
    dataObj.GetFromClipboard
    If dataObj.GetFormat(1) Then
        GetClipboardText = dataObj.GetText(1)
    Else
        GetClipboardText = ""
    End If
    
    If err.Number <> 0 Then
        GetClipboardText = ""
    End If
    On Error GoTo 0
End Function

Public Function ParseClipboardData(clipText As String) As Collection
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Dim items As New Collection
    Dim lines() As String
    Dim i As Integer
    Dim item As cptFilterItem
    
    ' Split by line breaks
    lines = Split(Replace(Replace(clipText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    
    For i = 0 To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Set item = New cptFilterItem
            item.Value = Trim(lines(i))
            item.Method = "Equals"
            items.Add item
        End If
    Next i
    
    Set ParseClipboardData = items
    Exit Function
    
ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "ParseClipboardData", err, Erl, "Error parsing clipboard data")
    'MsgBox "Error parsing clipboard data: " & err.Description, vbExclamation, "Parse Error"
    Set ParseClipboardData = New Collection
End Function

Private Sub filterReadCustomFields(ByVal curProj As Project)

    Dim i As Integer
    Dim fID As Long

    'Read local Custom Text Fields
    For i = 1 To 30

        If Len(curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Text" & i))) > 0 Then
            ReDim Preserve CustTextFields(1 To i)
            CustTextFields(i) = curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Text" & i))
        Else
            ReDim Preserve CustTextFields(1 To i)
            CustTextFields(i) = "Text" & i
        End If

    Next i
    
    'Read local Custom Number Fields
    For i = 1 To 20

        If Len(curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Number" & i))) > 0 Then
            ReDim Preserve CustNumFields(1 To i)
            CustNumFields(i) = curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Number" & i))
        Else
            ReDim Preserve CustNumFields(1 To i)
            CustNumFields(i) = "Number" & i
        End If

    Next i

    'Read local Custom Outline Code Fields
    For i = 1 To 10

        If Len(curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("OutlineCode" & i))) > 0 Then
            ReDim Preserve CustOLCodeFields(1 To i)
            CustOLCodeFields(i) = curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("OutlineCode" & i))
        Else
            ReDim Preserve CustOLCodeFields(1 To i)
            CustOLCodeFields(i) = "OutlineCode" & i
        End If

    Next i
    
    'Read Enterprise Custom Fields
    i = 1

    For fID = 188776000 To 188778000

        On Error GoTo fID_Error

        If Application.CustomFieldGetName(fID) <> "" Then
            ReDim Preserve EntFields(1 To i)
            EntFields(i) = Application.CustomFieldGetName(fID)
            i = i + 1
        End If

next_fID:

    Next fID
    
    Exit Sub

fID_Error:

    Resume next_fID

End Sub

Public Sub updateSummaries(ByVal checkboxValue As Boolean)

    curProj.Application.SummaryTasksShow (checkboxValue)

End Sub
