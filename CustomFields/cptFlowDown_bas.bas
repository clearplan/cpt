Attribute VB_Name = "cptFlowDown_bas"
'<cpt_version>v0.0.0</cpt_version>
Option Explicit

Sub cptShowLCFFlowDown_frm()
  Dim myLCFFlowDown_frm As cptLCFFlowDown_frm
  Set myLCFFlowDown_frm = New cptLCFFlowDown_frm
  Dim oProject As MSProject.Project
  For Each oProject In Projects
    If oProject.Subprojects.Count > 0 Then
      myLCFFlowDown_frm.cboMasterProject.AddItem oProject.Name
    End If
  Next
  myLCFFlowDown_frm.Show
  Set myLCFFlowDown_frm = Nothing
End Sub
