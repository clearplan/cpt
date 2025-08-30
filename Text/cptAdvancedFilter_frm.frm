VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptAdvancedFilter_frm 
   Caption         =   "UserForm1"
   ClientHeight    =   6492
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   6900
   OleObjectBlob   =   "cptAdvancedFilter_frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cptAdvancedFilter_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.3.2</cpt_version>
Option Explicit
Private Const MODULE_NAME As String = "cptAdvancedFilter_frm"
Private oFilterItems As Collection
Public blnDisableChangeEvents As Boolean

Private Sub sortField_Change()

    If blnDisableChangeEvents Then Exit Sub
    
    Dim usrResponse As Integer
    
    usrResponse = MsgBox("This action will clear all settings and values from the selected field." _
    & vbCr & vbCr & "Are you sure you wish to proceed?", vbYesNoCancel + vbExclamation, "Sort Field")
    
    If usrResponse = vbYes Then
        Exit Sub
    Else
        blnDisableChangeEvents = True
        sortField.ListIndex = 0
        blnDisableChangeEvents = False
    End If
    
End Sub

Private Sub UserForm_Initialize()
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Set oFilterItems = New Collection

    With Me.clipboardListHeaders
        .AddItem
        .List(0, 0) = "Row"
        .List(0, 1) = "Value"
        .List(0, 2) = "Filter Type"
        .List(0, 3) = "Count"
    End With
    
    Exit Sub

ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "UserForm_Initialize", Err, Erl, "Error initializing Advanced Filter form")
    'MsgBox "Error initializing Advanced Filter form: " & err.Description, vbCritical, "Initialization Error"
    
End Sub

Private Sub btnApply_Click()
    setFilter oFilterItems, Me.caseCheckbox.Value
End Sub

Private Sub btnPaste_Click()
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Dim clipText As String
    Dim oNewItems As Collection
    Dim appendResponse As VbMsgBoxResult
    Dim oItem As cptFilterItem_cls
    
    clipText = GetClipboardText()
    
    If clipText = "" Then
        MsgBox "Clipboard is empty or does not contain text data.", vbInformation, "No Data"
        Exit Sub
    End If
    
    If oFilterItems.Count > 0 Then
        appendResponse = MsgBox("Append items?", vbYesNoCancel, "Confirmation")
    
        Select Case appendResponse
        
            Case vbYes
            
                Set oNewItems = ParseClipboardData(clipText)
                For Each oItem In oNewItems
                    oFilterItems.Add oItem
                Next oItem
                RefreshItemsList
            
            Case vbNo
            
                Set oNewItems = ParseClipboardData(clipText)
                Set oFilterItems = oNewItems
                RefreshItemsList
            
            Case Else
            
                Exit Sub
        
        End Select
    Else
        Set oNewItems = ParseClipboardData(clipText)
        Set oFilterItems = oNewItems
        RefreshItemsList
    End If

    Exit Sub
    
ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "btnPaste_Click", Err, Erl, "Error loading clipboard data")
    MsgBox "Error loading clipboard data: " & Err.Description, vbExclamation, "Clipboard Error"
    
End Sub

Private Sub btnAdd_Click()
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Dim strResponse As String
    Dim editFrm As cptAdvancedFilterEdit_frm
        
    Set editFrm = New cptAdvancedFilterEdit_frm
    
    With editFrm
        .itemFilter_ComboBox.List = Split("Equals,Contains", ",")
        .itemFilter_ComboBox.ListIndex = 0
        .Caption = "Add Item"
        .Show vbModal
    End With
    
    strResponse = editFrm.Tag
    
    If strResponse = "Edit" Then
        Dim oNewItem As cptFilterItem_cls
        Set oNewItem = New cptFilterItem_cls
        oNewItem.Value = editFrm.itemValue_TextBox.Value
        oNewItem.Method = editFrm.itemFilter_ComboBox.Text
        oNewItem.Count = 0
        
        oFilterItems.Add oNewItem
        RefreshItemsList
    End If
    
    Unload editFrm
    Set editFrm = Nothing
    
    Exit Sub
    
ErrorHandler:
    If Not editFrm Is Nothing Then
        Unload editFrm
        Set editFrm = Nothing
    End If
    Call cptHandleErr(MODULE_NAME, "btnAdd_Click", Err, Erl, "Error adding item")
    'MsgBox "Error adding item: " & err.Description, vbExclamation, "Add Error"
End Sub

Private Sub btnDelete_Click()
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    If Me.clipboardList.ListIndex >= 0 Then
        oFilterItems.Remove Me.clipboardList.ListIndex + 1
        RefreshItemsList
    End If
    
    Exit Sub
    
ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "btnDelete_Click", Err, Erl, "Error deleting item")
    'MsgBox "Error deleting item: " & err.Description, vbExclamation, "Delete Error"
End Sub

Private Sub btnClear_Click()
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Set oFilterItems = New Collection
    RefreshItemsList
    
    Exit Sub
    
ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "btnClear_Click", Err, Erl, "Error clearing data")
    'MsgBox "Error clearing data: " & err.Description, vbExclamation, "Clear Error"
End Sub

Private Sub btnEquals_Click()
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Dim cntr As Integer
        
    Me.clipboardList.Clear
    
    For cntr = 1 To oFilterItems.Count
        oFilterItems(cntr).Method = "Equals"
    Next cntr
    
    RefreshItemsList
    
    Exit Sub

ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "btnEquals_Click", Err, Erl, "Error setting all to Equals")
    'MsgBox "Error setting all to Equals: " & err.Description, vbExclamation, "Equals Error"
End Sub

Private Sub btnContains_Click()
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Dim cntr As Integer
        
    Me.clipboardList.Clear
    
    For cntr = 1 To oFilterItems.Count
        oFilterItems(cntr).Method = "Contains"
    Next cntr
    
    RefreshItemsList
    
    Exit Sub

ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "btnContains_Click", Err, Erl, "Error setting all to Contains")
    'MsgBox "Error setting all to Contains: " & err.Description, vbExclamation, "Contains Error"
End Sub

Private Sub RefreshItemsList()
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
        Dim cntr As Integer
        
        Me.clipboardList.Clear
        
        For cntr = 1 To oFilterItems.Count
            With Me.clipboardList
                .AddItem
                .List(cntr - 1, 0) = cntr 'Row Id
                .List(cntr - 1, 1) = oFilterItems(cntr).Value 'Value
                .List(cntr - 1, 2) = oFilterItems(cntr).Method 'Type
                .List(cntr - 1, 3) = oFilterItems(cntr).Count 'count
            End With
        Next cntr
    
    Exit Sub
    
ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "RefreshItemsList", Err, Erl, "Error refreshing items list")
    'MsgBox "Error refreshing items list: " & err.Description, vbExclamation, "Refresh Error"

End Sub

Private Sub clipboardList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    If clipboardList.ListIndex > -1 Then
        edititem clipboardList.ListIndex + 1
    End If
    
    Exit Sub
    
ErrorHandler:
    Call cptHandleErr(MODULE_NAME, "clipboardList_DblClick", Err, Erl, "Error editing item")
    'MsgBox "Error editing item: " & err.Description, vbExclamation, "Edit Error"
End Sub

Private Sub edititem(itemIndex As Integer)
    If cptErrorTrapping Then On Error GoTo ErrorHandler Else On Error GoTo 0
    
    Dim strResponse As String
    Dim oCurrentItem As cptFilterItem_cls
    Dim editFrm As cptAdvancedFilterEdit_frm
    
    If itemIndex > 0 And itemIndex < oFilterItems.Count + 1 Then
        Set oCurrentItem = oFilterItems(itemIndex)
        
        Set editFrm = New cptAdvancedFilterEdit_frm
        
        With editFrm
            .itemValue_TextBox = oCurrentItem.Value
            .itemFilter_ComboBox.List = Split("Equals,Contains", ",")
            .itemFilter_ComboBox = oCurrentItem.Method
            .Caption = "Edit Item: " & oCurrentItem.Value
            .Show vbModal
        End With
        
        strResponse = editFrm.Tag
        
        If strResponse = "Edit" Then
            oCurrentItem.Value = editFrm.itemValue_TextBox.Value
            oCurrentItem.Method = editFrm.itemFilter_ComboBox.Text
            oCurrentItem.Count = 0
            RefreshItemsList
        End If
        
        Unload editFrm
        Set editFrm = Nothing
        
    End If
        
    Exit Sub
    
ErrorHandler:
    If Not editFrm Is Nothing Then
        Unload editFrm
        Set editFrm = Nothing
    End If
    Call cptHandleErr(MODULE_NAME, "edititem", Err, Erl, "Error editing item")
    'MsgBox "Error editing item: " & err.Description, vbExclamation, "Edit Error"
End Sub

Private Sub summaryCheckBox_Click()
    updateSummaries Me.summaryCheckBox.Value
    If isSorted Then
        Application.Sort Me.sortField.Value
    End If
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Handle Ctrl+V for paste anywhere on the form
    If KeyCode = vbKeyV And Shift = 2 Then
        ' Clear placeholder text if present
        btnPaste_Click
        KeyCode = 0 ' Prevent default handling
    End If
End Sub

Private Sub clipboardList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Handle Ctrl+V specifically in the text box
    If KeyCode = vbKeyV And Shift = 2 Then
        ' Clear placeholder text if present
        btnPaste_Click
        KeyCode = 0 ' Prevent default handling
    End If
End Sub

Private Sub btnClose_Click()
    Me.Tag = "Close"
    Me.Hide
    cptExitAdvancedFilter
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    btnClose_Click
  End If
End Sub
