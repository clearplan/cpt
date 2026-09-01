VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SCT_Form 
   Caption         =   "UserForm1"
   ClientHeight    =   1596
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5196
   OleObjectBlob   =   "SCT_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SCT_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelBtn_Click()
    Me.Tag = False
    Me.Hide
End Sub
Private Sub FileSaveTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim fileName As String
    Dim xlApp As Excel.Application
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False

    fileName = xlApp.GetSaveAsFilename(InitialFileName:=ActiveProject.Name, FileFilter:="Project Files,*.mpp,All Files,*.*", Title:="Save As File Name")
    
    xlApp.Quit
    Set xlApp = Nothing
    
    If fileName = "" Then Exit Sub
    
    If Right(fileName, 4) <> ".mpp" Then
        fileName = fileName & ".mpp"
    End If
    
    FileSaveTextBox.Text = fileName
    
    FileSaveTextBox.TextAlign = fmTextAlignLeft
    
    OKBtn.SetFocus
    
End Sub

Private Sub OKBtn_Click()
    If FileSaveTextBox.Text <> "<Click to select save location>" Then
        Me.Tag = True
        Me.Hide
    Else
        MsgBox "Please select a save location for the consolidated master."
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Integer
    
    Me.Tag = False
    
    FileSaveTextBox.Locked = True
    
    With FileUIDCombobox
        .Clear
        
        For i = 1 To 30
        
            .AddItem "Text" & i
            
        Next i
        
        .Text = .List(0)
        
    End With
    
    ResourceCheckbox.Value = False
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 Then
        Cancel = True
        CancelBtn_Click
    End If
End Sub
