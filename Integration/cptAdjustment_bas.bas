Attribute VB_Name = "cptAdjustment_bas"
'<cpt_version>v0.0.4</cpt_version>
Option Explicit

Sub cptShowAdjustment_frm()
  'objects
  Dim myAdjustment_frm As cptAdjustment_frm
  Dim oResources As Object 'Scripting.Dictionary
  Dim oResource As MSProject.Resource
  'strings
  Dim strIgnoreTaskType As String
  Dim strResources As String
  'variants
  Dim vResources As Variant
  Dim vResource As Variant
 
  'prevent spawning
  If Not cptGetUserForm("cptAdjustment_frm") Is Nothing Then Exit Sub

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If ActiveProject.ResourceCount = 0 Then
    MsgBox "This project has no resources.", vbExclamation + vbOKOnly, "No Resources"
    GoTo exit_here
  End If
  Set oResources = CreateObject("Scripting.Dictionary")
  For Each oResource In ActiveProject.Resources
    If oResource.Type = pjResourceTypeWork Then
      strResources = strResources & oResource.Name & ","
      oResources.Add oResource.Name, oResource.UniqueID
    End If
  Next
  If Len(strResources) = 0 Then
    MsgBox "No Labor [Resource Type=Work] resources found.", vbExclamation + vbOKOnly, "Adjustments"
    GoTo exit_here
  End If
  strResources = Left(strResources, Len(strResources) - 1)
  vResources = Split(strResources, ",")
  
  cptQuickSort vResources, 0, UBound(vResources)
  
  'vResources = Split("All Resources," & Join(vResources, ","), ",")
  
  cptStartEvents
  
  Set myAdjustment_frm = New cptAdjustment_frm
  With myAdjustment_frm
    .Caption = "ETC Adjustment (" & cptGetVersion("cptAdjustment_frm") & ")"
    .lboHeader.Clear
    .lboHeader.AddItem
    .lboHeader.List(.lboHeader.ListCount - 1, 0) = "UID"
    .lboHeader.List(.lboHeader.ListCount - 1, 1) = "RESOURCE"
    .lboHeader.List(.lboHeader.ListCount - 1, 2) = "ETC"
    .lboHeader.List(.lboHeader.ListCount - 1, 3) = "NEW ETC"
    .lboTotal.Clear
    .lboTotal.AddItem
    .lboTotal.List(.lboTotal.ListCount - 1, 1) = "TOTAL"
    .cboResources.Clear
    .cboResources.AddItem
    .cboResources.List(0, 0) = "0"
    .cboResources.List(0, 1) = "All Resources"
    For Each vResource In vResources
      .cboResources.AddItem
      .cboResources.List(.cboResources.ListCount - 1, 0) = CStr(oResources(vResource))
      .cboResources.List(.cboResources.ListCount - 1, 1) = vResource
    Next vResource
    .cboResources.Value = "0" '"All Resources"
    .optTarget = True
    strIgnoreTaskType = cptGetSetting("ETCAdjustment", "chkIgnoreTaskType")
    If Len(strIgnoreTaskType) > 0 Then
      .chkIgnoreTaskType = CBool(strIgnoreTaskType)
    Else
      .chkIgnoreTaskType = True 'defaults to true
    End If
    cptRefreshAdjustment myAdjustment_frm
    .Show False
  End With
  
exit_here:
  On Error Resume Next
  Set oResources = Nothing
  Set myAdjustment_frm = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptShowAdjustment_frm", Err)
  Resume exit_here

End Sub

Sub cptApplyAdjustment(ByRef myAdjustment_frm As cptAdjustment_frm)
  'objects
  Dim oAssignment As MSProject.Assignment
  Dim oTask As MSProject.Task
  'strings
  Dim strStatusBar As String
  'longs
  Dim lngItem As Long
  Dim lngType As Long
  'integers
  'doubles
  'booleans
  Dim blnRefreshStatusBar As Boolean
  Dim blnIgnoreTaskType As Boolean
  Dim blnEffortDriven As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptSpeed True
  Application.OpenUndoTransaction "cptAdjustment"
  
  blnRefreshStatusBar = False
  If InStr(Application.StatusBar, "Remaining Work") > 0 Then
    blnRefreshStatusBar = True
    strStatusBar = Replace(Application.StatusBar, myAdjustment_frm.lboTotal.List(0, 2), myAdjustment_frm.lboTotal.List(0, 3))
  End If
  
  With myAdjustment_frm
    blnIgnoreTaskType = .chkIgnoreTaskType
    cptSaveSetting "ETCAdjustment", "chkIgnoreTaskType", IIf(.chkIgnoreTaskType, "1", "0")
    If Not blnIgnoreTaskType Then
      If MsgBox("Adjustment may alter remaining duration on selected tasks." & vbCrLf & vbCrLf & "Continue?", vbExclamation + vbOKCancel, "Please Confirm") = vbCancel Then GoTo exit_here
    End If
    With .lboAdjustmentPreview
      If .ListCount = 0 Then GoTo exit_here
      For lngItem = 0 To .ListCount - 1
        Set oTask = ActiveProject.Tasks.UniqueID(.List(lngItem, 0))
        For Each oAssignment In oTask.Assignments
          If oAssignment.ResourceUniqueID = CLng(.List(lngItem, 1)) Then
            If blnIgnoreTaskType Then 'capture settings
              lngType = oTask.Type
              blnEffortDriven = oTask.EffortDriven
              'change type to apply
              oTask.Type = pjFixedDuration
              oTask.EffortDriven = False
            End If
            'apply adjustment
            oAssignment.RemainingWork = CDbl(.List(lngItem, 4)) * 60
            If blnIgnoreTaskType Then
              'restore settings
              oTask.Type = lngType
              If lngType <> pjFixedWork Then oTask.EffortDriven = blnEffortDriven
            End If
          End If
        Next oAssignment
      Next lngItem
    End With
  End With
  
  If blnRefreshStatusBar Then
    Application.StatusBar = strStatusBar
  End If
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  cptSpeed False
  Set oAssignment = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptApplyAdjustment", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshAdjustment(ByRef myAdjustment_frm As cptAdjustment_frm)
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oAssignment As MSProject.Assignment
  Dim oTask As MSProject.Task
  'strings
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  Dim dblPercent As Double
  Dim dblETC As Double
  Dim dblNewETC As Double
  Dim dblNewTotal As Double
  'booleans
  'variants
  'dates
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  With myAdjustment_frm
    .lboAdjustmentPreview.Clear
    For Each oTask In ActiveSelection.Tasks
      If oTask Is Nothing Then GoTo next_task
      If oTask.ExternalTask Then GoTo next_task
      If Not oTask.Active Then GoTo next_task
      If IsDate(oTask.ActualFinish) Then GoTo next_task
      If oTask.Assignments.Count > 0 Then
        For Each oAssignment In oTask.Assignments
          If oAssignment.ResourceType <> pjResourceTypeWork Then GoTo next_assignment
          If .cboResources.Value <> "0" And oAssignment.ResourceUniqueID <> CLng(.cboResources.Value) Then GoTo next_assignment
          .lboAdjustmentPreview.AddItem
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 0) = oTask.UniqueID
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 1) = oAssignment.ResourceUniqueID
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 2) = oAssignment.ResourceName 'todo: necessary if filtered?
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 3) = Format(Round((oAssignment.RemainingWork / 60), 2), "#,###,##0.00")
          dblETC = dblETC + oAssignment.RemainingWork
next_assignment:
        Next oAssignment
      End If
next_task:
    Next oTask

    If .lboAdjustmentPreview.ListCount > 0 Then
      If Len(.txtAmount) > 0 Then
        If .optDelta Then
          dblNewETC = (dblETC / 60) + CDbl(.txtAmount.Value)
          If dblNewETC < 0 Then
            dblNewETC = -(dblETC / 60) + 0.5
            .txtAmount.Value = Format(dblNewETC, "#,##0.00")
          End If
        ElseIf .optTarget Then
          dblNewETC = CDbl(.txtAmount.Value)
          If dblNewETC < 0 Then
            dblNewETC = 0.5
            .txtAmount.Value = Format(dblNewETC, "#,##0.00")
          End If
        ElseIf .optPercent Then
          dblPercent = CDbl(.txtAmount.Value)
          If dblPercent = -1 Then
            dblPercent = 0.99
            .txtAmount.Value = "0.99"
          End If
        End If
      End If
      
      For lngItem = 0 To .lboAdjustmentPreview.ListCount - 1
        If CDbl(.lboAdjustmentPreview.List(lngItem, 3)) = 0 Then
          .lboAdjustmentPreview.List(lngItem, 4) = 0
          GoTo next_item
        Else
          If dblNewETC <> 0 Then
            .lboAdjustmentPreview.List(lngItem, 4) = Format(((.lboAdjustmentPreview.List(lngItem, 3) * 60) / dblETC) * dblNewETC, "#,##0.00")
            dblNewTotal = dblNewTotal + ((.lboAdjustmentPreview.List(lngItem, 3) * 60) / dblETC) * dblNewETC
          ElseIf dblPercent <> 0 Then
            If dblPercent < 0 Then dblPercent = 1 + dblPercent
            .lboAdjustmentPreview.List(lngItem, 4) = Format((.lboAdjustmentPreview.List(lngItem, 3)) * dblPercent, "#,##0.00")
            dblNewTotal = dblNewTotal + (.lboAdjustmentPreview.List(lngItem, 3) * dblPercent)
          Else
            .lboAdjustmentPreview.List(lngItem, 4) = Format(.lboAdjustmentPreview.List(lngItem, 3), "#,##0.00")
            dblNewTotal = dblNewTotal + .lboAdjustmentPreview.List(lngItem, 3)
          End If
        End If
next_item:
      Next lngItem
      
    End If
    
    'update totals
    .lboTotal.List(0, 2) = Format((dblETC / 60), "#,##0.00")
    .lboTotal.List(0, 3) = Format(dblNewTotal, "#,##0.00")
    .lboTotal.Top = .lboHeader.Top + .lboHeader.Height + .lboAdjustmentPreview.Height
    
  End With

exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  Set oAssignment = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptRefreshAdjustment", Err, Erl)
  Resume exit_here
End Sub

