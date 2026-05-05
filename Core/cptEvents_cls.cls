VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cptEvents_cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<cpt_version>v1.5.0</cpt_version>
Option Explicit
Private WithEvents MyMSPApplication As Application
Attribute MyMSPApplication.VB_VarHelpID = -1

Private Sub Class_Initialize()

    Set MyMSPApplication = Application

End Sub

Private Sub MyMSPApplication_WindowSelectionChange(ByVal Window As Window, ByVal sel As Selection, ByVal selType As Variant)
'objects
Dim oUserForm As Object 'UserForm
Dim oTasks As MSProject.Tasks
Dim oTask As MSProject.Task
Dim oResources As MSProject.Resources
'strings
'longs
Dim lngItem As Long
Dim lngItems As Long
'integers
'booleans
'variants
'dates

  On Error Resume Next
  
  'Validate users selected view type
  If ActiveWindow.ActivePane.View.Type = pjTaskItem Then
    Set oTasks = ActiveSelection.Tasks
    If Not oTasks Is Nothing Then
      Set oUserForm = cptGetUserForm("cptText_frm") 'returns Nothing if form is not loaded
      If Not oUserForm Is Nothing Then
        oUserForm.Controls("lboOutput").Clear
        For Each oTask In ActiveSelection.Tasks
          oUserForm.Controls("lboOutput").AddItem
          oUserForm.Controls("lboOutput").List(lngItem, 0) = oTask.UniqueID
          oUserForm.Controls("lboOutput").List(lngItem, 1) = oTask.Name
          lngItem = lngItem + 1
        Next oTask
        Application.Run "cptUpdatePreview", oUserForm
      End If
    End If
  End If
  
  'Validate users selected view type
  If cptGetShowStatusBarCountFirstRun Then
    If ActiveWindow.ActivePane.View.Type = pjTaskItem Then
      Set oTasks = ActiveSelection.Tasks
      If Not oTasks Is Nothing Then
        lngItems = oTasks.Count
        If ActiveSelection.FieldIDList.Count = 1 And lngItems > 1 Then
          If ActiveSelection.FieldIDList(1) <> -1 Then
            cptGetSums oTasks, ActiveSelection.FieldIDList(1)
          Else
            Application.StatusBar = ""
          End If
        Else
          Application.StatusBar = Format(lngItems, "#,##0") & " task" & IIf(lngItems = 1, "", "s") & " selected"
        End If
      Else
        Application.StatusBar = "0 tasks selected"
      End If
    ElseIf ActiveWindow.ActivePane.View.Type = pjResourceItem Then
      Set oResources = ActiveSelection.Resources
      If Not oResources Is Nothing Then
        lngItems = oResources.Count
        Application.StatusBar = Format(lngItems, "#,##0") & " resource" & IIf(lngItems = 1, "", "s") & " selected"
      Else
        Application.StatusBar = "0 resources selected"
      End If
    End If
  End If
  
  Set oUserForm = cptGetUserForm("cptNetworkBrowser_frm") 'returns Nothing if form is not loaded
  If Not oUserForm Is Nothing Then Application.Run "cptShowPreds", oUserForm
  
  Set oUserForm = cptGetUserForm("cptAdjustment_frm") 'returns Nothing if form is not loaded
  If Not oUserForm Is Nothing Then Application.Run "cptRefreshAdjustment", oUserForm
  
  Set oUserForm = cptGetUserForm("cptTaskHistory_frm") 'returns Nothing if form is not loaded
  If Not oUserForm Is Nothing Then Application.Run "cptUpdateTaskHistory", oUserForm
  
  Set oUserForm = cptGetUserForm("cptQBD_frm") 'returns Nothing if form is not loaded
  If Not oUserForm Is Nothing Then Application.Run "cptUpdateQBDForm", oUserForm
  
  Set oUserForm = cptGetUserForm("cptSmartDuration_frm") 'returns Nothing if form is not loaded
  If Not oUserForm Is Nothing Then Application.Run "cptUpdateSmartDurationForm", oUserForm
  
exit_here:
  On Error Resume Next
  Set oUserForm = Nothing
  Set oTask = Nothing
  Set oTasks = Nothing
  Set oResources = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptEvents_cls", "MyMSApplication_WindowSelectionChange", Err, Erl)
  Resume exit_here
  
End Sub
