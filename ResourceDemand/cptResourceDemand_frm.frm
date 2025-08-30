VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptResourceDemand_frm 
   Caption         =   "Export Resource Demand"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   OleObjectBlob   =   "cptResourceDemand_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptResourceDemand_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.5.1</cpt_version>
Option Explicit
Private Const adVarChar As Long = 200

Private Sub cboMonths_Change()
  If Me.cboMonths.Value = 0 Then
    Me.cboWeeks.Enabled = True
    Me.cboWeeks.Locked = False
    Me.cboWeekday.Enabled = True
    Me.cboWeekday.Locked = False
  ElseIf Me.cboMonths.Value = 1 Then 'fiscal
    Me.cboWeeks.Value = "Ending"
    Me.cboWeeks.Enabled = False
    Me.cboWeeks.Locked = True
    Me.cboWeekday.Value = "Friday"
    Me.cboWeekday.Enabled = False
    Me.cboWeekday.Locked = True
  End If
End Sub

Private Sub cboWeeks_Change()
  Me.cboWeekday.Clear
  Select Case Me.cboWeeks
    Case "Beginning"
      Me.cboWeekday.AddItem "Sunday"
      Me.cboWeekday.AddItem "Monday"
      Me.cboWeekday.Value = "Monday"
    Case "Ending"
      Me.cboWeekday.AddItem "Friday"
      Me.cboWeekday.AddItem "Saturday"
      Me.cboWeekday.Value = "Friday"
  End Select
End Sub

Private Sub chkCosts_AfterUpdate()
  Me.chkA.Enabled = Me.chkCosts
  Me.chkB.Enabled = Me.chkCosts
  Me.chkC.Enabled = Me.chkCosts
  Me.chkD.Enabled = Me.chkCosts
  Me.chkE.Enabled = Me.chkCosts
End Sub

Private Sub cmdAdd_Click()
Dim lngField As Long, lngExport As Long, lngExists As Long
Dim blnExists As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngField = 0 To Me.lboFields.ListCount - 1
    If Me.lboFields.Selected(lngField) Then
      'ensure doesn't already exist
      blnExists = False
      For lngExists = 0 To Me.lboExport.ListCount - 1
        If Me.lboExport.List(lngExists, 0) = Me.lboFields.List(lngField) Then
          GoTo next_item
        End If
      Next lngExists
      Me.lboExport.AddItem
      lngExport = Me.lboExport.ListCount - 1
      Me.lboExport.List(lngExport, 0) = Me.lboFields.List(lngField, 0)
      Me.lboExport.List(lngExport, 1) = Me.lboFields.List(lngField, 1)
    End If
next_item:
  Next lngField

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "cmdAdd_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdCancel_Click()
Dim strFileName As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  Unload Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "cmdCancel_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdExport_Click()
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptExportResourceDemand Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "cmdExport_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdRemove_Click()
Dim lgExport As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lgExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If Me.lboExport.Selected(lgExport) Then
      Me.lboExport.RemoveItem lgExport
    End If
  Next lgExport

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "cmdRemove_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "lblURL_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub stxtSearch_Change()
'objects
'strings
Dim strFileName As String
'longs
Dim lngItem As Long
'integers
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboFields.Clear

  strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
  With CreateObject("ADODB.Recordset")
    .Open strFileName
    If Len(Me.stxtSearch.Text) > 0 Then
      .Filter = "CUSTOM_NAME LIKE '*" & cptRemoveIllegalCharacters(Me.stxtSearch.Text) & "*'"
    Else
      .Filter = 0
    End If
    If .RecordCount > 0 Then .MoveFirst
    lngItem = 0
    Do While Not .EOF
      Me.lboFields.AddItem
      Me.lboFields.List(lngItem, 0) = .Fields(0)
      If .Fields(0) >= 188776000 Then 'enterprise
        Me.lboFields.List(lngItem, 1) = .Fields(1) & " (Enterprise)"
      Else
        Me.lboFields.List(lngItem, 1) = .Fields(1) & " (" & FieldConstantToFieldName(.Fields(0)) & ")"
      End If
      .MoveNext
      lngItem = lngItem + 1
    Loop
    .Close
    Me.lblStatus.Caption = lngItem & " record" & IIf(lngItem = 1, "", "s") & " found."
  End With
  
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "stxtSearch_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub
