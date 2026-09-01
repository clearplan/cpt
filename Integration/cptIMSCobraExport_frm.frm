VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptIMSCobraExport_frm 
   Caption         =   "IMS Export Utility"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   15396
   OleObjectBlob   =   "cptIMSCobraExport_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptIMSCobraExport_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v3.5.3</cpt_version>
Option Explicit
Private Const THIS_MODULE As String = "cptIMSCobraExport_frm"

Private Sub AsgnPcntBox_Change() 'v3.3.1
    
    If isIMSfield(AsgnPcntBox.value) = False And AsgnPcntBox.value <> "" And AsgnPcntBox.value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        AsgnPcntBox.value = "" 'v3.4.3
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fAssignPcnt").value = Me.AsgnPcntBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fAssignPcnt", False, msoPropertyTypeString, Me.AsgnPcntBox.value
    Resume PropFound
End Sub

Private Sub bcrBox_Change()

    If checkDuplicate(bcrBox) = True Then
        MsgBox "Please select a unique IMS Field."
        bcrBox.value = ""
        Exit Sub
    End If
    
    If isIMSfield(bcrBox.value) = False And bcrBox.value <> "" And bcrBox.value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        bcrBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fBCR").value = Me.bcrBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fBCR", False, msoPropertyTypeString, Me.bcrBox.value
    Resume PropFound
End Sub

Private Function checkDuplicate(ByVal cBoxTest As MSForms.ComboBox) As Boolean 'v3.3.8

    If cBoxTest.value = "<None>" Or cBoxTest.value = "" Then
    
        checkDuplicate = False
        Exit Function
    
    End If

    Dim cBoxOther As MSForms.ComboBox 'v3.3.8
    Dim formObj As MSForms.Control 'v3.3.8
    
    For Each formObj In Me.TabButtons.Pages(1).Controls
    
        If TypeName(formObj) = "ComboBox" Then
        
            Set cBoxOther = formObj
            
            If cBoxOther.Name <> cBoxTest.Name Then
            
                If cBoxOther.value = cBoxTest.value Then
                
                    checkDuplicate = True
                    Exit Function
                
                End If
            
            End If
        
        End If
    
    Next formObj
    
    checkDuplicate = False

End Function

Private Sub BcrBtn_Change()

    If BcrBtn = True Then
        Me.BCR_ID_TextBox.Enabled = True
        Exit Sub
    Else
        Me.BCR_ID_TextBox.Enabled = False
        Exit Sub
    End If

End Sub

Private Sub BcrBtn_Click()
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing

PropFound:

    If docProps("fBCR").value <> "<None>" Then
        Exit Sub
    End If
    
PropMissing:

    MsgBox "Please map a BCR Field before using the BCR Export function."
    Me.BcrBtn = False
    Me.TotalProjBtn = True
    Me.BCR_ID_TextBox.Enabled = False
    Exit Sub
End Sub

Private Sub BCWS_Checkbox_Change()

    If BCWS_Checkbox.value = True Then
        Me.TotalProjBtn.Enabled = True
        Me.BcrBtn.Enabled = True
        If BcrBtn = True Then
            BCR_ID_TextBox.Enabled = True
        End If
        Me.exportDescCheckBox.Enabled = True
        Me.exportTPhaseCheckBox.Enabled = True
        Me.Milestone_CheckBox.Enabled = True 'v3.4.1
    Else
        If Me.WhatIf_CheckBox.value = False Then 'v3.3.15
            Me.BcrBtn.Enabled = False
            Me.TotalProjBtn.Enabled = False
            Me.BCR_ID_TextBox.Enabled = False
        End If
        Me.exportDescCheckBox.Enabled = False
        If Me.ETC_Checkbox.value = False And Me.WhatIf_CheckBox.value = False Then
            Me.exportTPhaseCheckBox.Enabled = False
        End If
        Me.Milestone_CheckBox.Enabled = False 'v3.4.1
    End If

End Sub


Private Sub caID1Box_Change()

    If checkDuplicate(caID1Box) = True Then
        MsgBox "Please select a unique IMS Field."
        caID1Box.value = ""
        Exit Sub
    End If
    
    If isIMSfield(caID1Box.value) = False And caID1Box.value <> "" Then
        MsgBox "Please select a valid IMS Field."
        caID1Box.value = ""
        CAID1TxtBox.value = ""
        Exit Sub
    End If

    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID1").value = Me.caID1Box.value
    If Me.Tag = "Loaded" And Me.CAID1TxtBox.value = "" Then Me.CAID1TxtBox.value = Me.caID1Box.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID1", False, msoPropertyTypeString, Me.caID1Box.value
    Resume PropFound

End Sub

Private Sub CAID1TxtBox_Change()
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID1t").value = Me.CAID1TxtBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID1t", False, msoPropertyTypeString, Me.CAID1TxtBox.value
    Resume PropFound
End Sub

Private Sub CAID1TxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID1t").value = Me.CAID1TxtBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID1t", False, msoPropertyTypeString, Me.CAID1TxtBox.value
    Resume PropFound
End Sub

Private Sub caID2Box_Change()

    If checkDuplicate(caID2Box) = True Then
        MsgBox "Please select a unique IMS Field."
        caID2Box.value = ""
        Exit Sub
    End If
    
    If isIMSfield(caID2Box.value) = False And caID2Box.value <> "" And caID2Box.value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        caID2Box.value = ""
        CAID2TxtBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID2").value = Me.caID2Box.value
    If Me.Tag = "Loaded" And Me.CAID2TxtBox.value = "" Then Me.CAID2TxtBox.value = Me.caID2Box.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    If Me.caID2Box.value = "<None>" Then
        Me.CAID2TxtBox.Enabled = False
        Me.CAID2TxtBox.Visible = False
    Else
        Me.CAID2TxtBox.Enabled = True
        Me.CAID2TxtBox.Visible = True
        If Me.Tag = "Loaded" And Me.CAID2TxtBox.value = "" Then Me.CAID2TxtBox.value = Me.caID2Box.value
    End If
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID2", False, msoPropertyTypeString, Me.caID2Box.value
    Resume PropFound
End Sub

Private Sub CAID2TxtBox_Change()
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID2t").value = Me.CAID2TxtBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID2t", False, msoPropertyTypeString, Me.CAID2TxtBox.value
    Resume PropFound
End Sub

Private Sub CAID2TxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID2t").value = Me.CAID2TxtBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID2t", False, msoPropertyTypeString, Me.CAID2TxtBox.value
    Resume PropFound
End Sub

Private Sub caID3Box_Change()

    If checkDuplicate(caID3Box) = True Then
        MsgBox "Please select a unique IMS Field."
        caID3Box.value = ""
        Exit Sub
    End If
    
    If isIMSfield(caID3Box.value) = False And caID3Box.value <> "" And caID3Box.value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        caID3Box.value = ""
        CAID3TxtBox = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID3").value = Me.caID3Box.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    If Me.caID3Box.value = "<None>" Then
        Me.CAID3TxtBox.Enabled = False
        Me.CAID3TxtBox.Visible = False
    Else
        Me.CAID3TxtBox.Enabled = True
        Me.CAID3TxtBox.Visible = True
        If Me.Tag = "Loaded" And Me.CAID3TxtBox.value = "" Then Me.CAID3TxtBox.value = Me.caID3Box.value
    End If
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID3", False, msoPropertyTypeString, Me.caID3Box.value
    Resume PropFound
End Sub

Private Sub CAID3TxtBox_Change()
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID3t").value = Me.CAID3TxtBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID3t", False, msoPropertyTypeString, Me.CAID3TxtBox.value
    Resume PropFound
End Sub

Private Sub CAID3TxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID3t").value = Me.CAID3TxtBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID3t", False, msoPropertyTypeString, Me.CAID3TxtBox.value
    Resume PropFound
End Sub

Private Sub camBox_Change()

    If checkDuplicate(camBox) = True Then
        MsgBox "Please select a unique IMS Field."
        camBox.value = ""
        Exit Sub
    End If
    
    If isIMSfield(camBox.value) = False And camBox.value <> "" Then
        MsgBox "Please select a valid IMS Field."
        camBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAM").value = Me.camBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAM", False, msoPropertyTypeString, Me.camBox.value
    Resume PropFound
End Sub

Private Sub CancelBtn_Click()
    Me.Tag = "Cancel"
    Me.Hide
End Sub

Private Sub cptLinkLabel_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then
    CreateObject("WScript.Shell").Run "https://www.ClearPlanConsulting.com"
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr(THIS_MODULE, "lblURL", err, Erl)
  Resume exit_here
End Sub

Private Sub CSVBtn_Change()

    If CSVBtn.value = True Then
        Me.BCWS_Checkbox.Enabled = True
        Me.BCWP_Checkbox.Enabled = True
        Me.ETC_Checkbox.Enabled = True
        Me.WhatIf_CheckBox.Enabled = True 'v3.2
        Me.ResExportCheckbox.Enabled = True
        Me.Milestone_CheckBox.Enabled = True 'v3.4
        If Me.ResExportCheckbox.value = True Then
            Me.exportTPhaseCheckBox.Enabled = True
            If Me.exportTPhaseCheckBox.value = True Then 'v3.4
                Me.ScaleCombobox.Enabled = True
                Me.ScaleLabel.Enabled = True
                If Me.ScaleCombobox.value = "Weekly" Then
                    Me.WeekStartCombobox.Enabled = True
                    Me.WeekStartLabel.Enabled = True
                End If
            End If
        Else
            Me.exportTPhaseCheckBox.Enabled = False
        End If
        If Me.BCWS_Checkbox = True Then
            Me.TotalProjBtn.Enabled = True
            Me.BcrBtn.Enabled = True
            Me.exportDescCheckBox.Enabled = True
            If Me.BcrBtn = True Then Me.BCR_ID_TextBox.Enabled = True
        End If
    Else
        Me.BCWS_Checkbox.Enabled = False
        Me.BCWP_Checkbox.Enabled = False
        Me.ETC_Checkbox.Enabled = False
        Me.WhatIf_CheckBox.Enabled = False 'v3.2
        Me.TotalProjBtn.Enabled = False
        Me.ResExportCheckbox.Enabled = False
        Me.exportTPhaseCheckBox.Enabled = False
        Me.BcrBtn.Enabled = False
        Me.BCR_ID_TextBox.Enabled = False
        Me.exportDescCheckBox.Enabled = False
        Me.Milestone_CheckBox.Enabled = False 'v3.4
        Me.WeekStartCombobox.Enabled = False
        Me.ScaleCombobox.Enabled = False 'v3.4
        Me.ScaleLabel.Enabled = False 'v3.4
        Me.WeekStartLabel.Enabled = False 'v3.4
    End If
    
    If BCWS_Checkbox.value = False And BCWP_Checkbox.value = False And ETC_Checkbox.value = False And WhatIf_CheckBox.value = False Then 'v3.2
    
        BCWS_Checkbox.value = True
        Me.TotalProjBtn.Enabled = True
        Me.BcrBtn.Enabled = True
        If BcrBtn = True Then
            BCR_ID_TextBox.Enabled = True
        End If
    
    End If

End Sub

Private Sub DateFormat_Combobox_Change() 'v3.3.5

    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("dateFmt").value = Me.DateFormat_Combobox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "dateFmt", False, msoPropertyTypeString, Me.DateFormat_Combobox.value
    Resume PropFound

End Sub

Private Sub ETC_Checkbox_Click()
    If Me.ETC_Checkbox = True Then
        Me.exportTPhaseCheckBox.Enabled = True
    Else
        If Me.BCWS_Checkbox.value = False And Me.WhatIf_CheckBox.value = False Then
            Me.exportTPhaseCheckBox.Enabled = False
        End If
    End If
End Sub

Private Sub evtBox_Change()

    If checkDuplicate(evtBox) = True Then
        MsgBox "Please select a unique IMS Field."
        evtBox.value = ""
        Exit Sub
    End If
    
    If isIMSfield(evtBox.value) = False And evtBox.value <> "" Then
        MsgBox "Please select a valid IMS Field."
        evtBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fEVT").value = Me.evtBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fEVT", False, msoPropertyTypeString, Me.evtBox.value
    Resume PropFound
End Sub

Private Sub ExportBtn_Click()

    If CSVBtn.value = True And BCWS_Checkbox.value = False And BCWP_Checkbox.value = False And ETC_Checkbox.value = False And WhatIf_CheckBox.value = False Then 'v3.2
    
        MsgBox "You must select at least one CSV export file type."
        Exit Sub
        
    End If
    
    If BCR_ID_TextBox.Enabled = True Then
        If BCR_ID_TextBox.value = "Enter BCR ID" Or BCR_ID_TextBox.value = "" Then
            MsgBox "You must enter a valid BCR ID."
            BCR_ID_TextBox.value = "Enter BCR ID"
            Exit Sub
        End If
        If Me.bcrBox.value = "<None>" Then
            MsgBox "You must map a BCR ID Field."
            Exit Sub
        End If
    End If
    
    If WhatIf_CheckBox.value = True Then 'v3.2
        If Me.whatifBox.value = "<None>" Then
            MsgBox "You must map a What-If Field."
            Exit Sub
        End If
    End If

    Me.Tag = "Export"
    Me.Hide
    
End Sub

Private Sub exportTPhaseCheckBox_Click() 'v3.3.6
    If exportTPhaseCheckBox.value = True Then
        'if exporting timescaled data
        'increase visibility of MSP's week start day
        ScaleLabel.Enabled = True
        ScaleCombobox.Enabled = True
    Else
        WeekStartLabel.Enabled = False
        WeekStartCombobox.Enabled = False
        ScaleLabel.Enabled = False 'v3.4
        ScaleCombobox.Enabled = False 'v3.4
    End If
End Sub

Private Sub msidBox_Change()

    If checkDuplicate(msidBox) = True Then
        MsgBox "Please select a unique IMS Field."
        msidBox.value = ""
        Exit Sub
    End If
    
    If isIMSfield(msidBox.value) = False And msidBox.value <> "" And msidBox.value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        msidBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fMSID").value = Me.msidBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fMSID", False, msoPropertyTypeString, Me.msidBox.value
    Resume PropFound
End Sub

Private Sub mswBox_Change()

    If checkDuplicate(mswBox) = True Then
        MsgBox "Please select a unique IMS Field."
        mswBox.value = ""
        Exit Sub
    End If
    
    If isIMSfield(mswBox.value) = False And mswBox.value <> "" And mswBox.value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        mswBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fMSW").value = Me.mswBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fMSW", False, msoPropertyTypeString, Me.mswBox.value
    Resume PropFound
End Sub

Private Sub PercentBox_Change()

    If checkDuplicate(PercentBox) = True Then
        MsgBox "Please select a unique IMS Field."
        PercentBox.value = ""
        Exit Sub
    End If
    
    If isIMSfield(PercentBox.value) = False And PercentBox.value <> "" Then
        MsgBox "Please select a valid IMS Field."
        PercentBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fPCNT").value = Me.PercentBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fPCNT", False, msoPropertyTypeString, Me.PercentBox.value
    Resume PropFound
End Sub

Private Sub projBox_Change()

    If checkDuplicate(projBox) = True Then
        MsgBox "Please select a unique IMS Field."
        projBox.value = ""
        Exit Sub
    End If
    
    If isIMSfield(projBox.value) = False And projBox.value <> "" And projBox.value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        projBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fProject").value = Me.projBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fProject", False, msoPropertyTypeString, Me.projBox.value
    Resume PropFound

End Sub

Private Sub resBox_Change()
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fResID").value = Me.resBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fResID", False, msoPropertyTypeString, Me.resBox.value
    Resume PropFound
    
End Sub

Private Sub ResExportCheckbox_Click()

    If ResExportCheckbox.value = True Then
        exportTPhaseCheckBox.Enabled = True
    Else
        exportTPhaseCheckBox.Enabled = False
    End If

End Sub

Private Sub RunDataBtn_Click()
    
    Me.Tag = "DataCheck"
    Me.Hide
    
End Sub

Private Sub ScaleCombobox_Change() 'v3.4

    If ScaleCombobox.value = "Weekly" Then
        WeekStartCombobox.Enabled = True
        WeekStartLabel.Enabled = True
    Else
        WeekStartCombobox.Enabled = False
        WeekStartLabel.Enabled = False
    End If

End Sub

Private Sub TabButtons_Click(ByVal Index As Long)
    If Index <> 1 And Me.TabButtons(1).Tag = False Then
        Me.TabButtons.value = 1
        Exit Sub
    End If
    If Index <> 1 And VerifyTitles = False Then
        Me.TabButtons.value = 1
        MsgBox "Complete CA ID Titles"
        Exit Sub
    End If
End Sub

Private Sub UserForm_Activate()

    If Me.TabButtons(1).Tag = False Then
        Me.TabButtons.value = 1
        MsgBox "Please complete the Custom Field Configuration"
    End If

End Sub

Private Sub UserForm_Initialize()

    Me.MPPBtn.value = True
    Me.TabButtons.value = 0
    Me.ExportBtn.SetFocus
    
    If CSVBtn.value = True Then
        Me.BCWS_Checkbox.Enabled = True
        Me.BCWP_Checkbox.Enabled = True
        Me.ETC_Checkbox.Enabled = True
        Me.WhatIf_CheckBox.Enabled = True 'v3.2
        Me.ResExportCheckbox.Enabled = True
        Me.Milestone_CheckBox.Enabled = True 'v3.4
        If Me.ResExportCheckbox.value = True Then
            Me.exportTPhaseCheckBox.Enabled = True
            If Me.exportTPhaseCheckBox.value = True Then 'v3.4
                Me.ScaleCombobox.Enabled = True
                Me.ScaleLabel.Enabled = True
                If Me.ScaleCombobox.value = "Weekly" Then
                    Me.WeekStartCombobox.Enabled = True
                    Me.WeekStartLabel.Enabled = True
                End If
            End If
        Else
            Me.exportTPhaseCheckBox.Enabled = False
        End If
        If Me.BCWS_Checkbox = True Then
            Me.TotalProjBtn.Enabled = True
            Me.BcrBtn.Enabled = True
            Me.exportDescCheckBox.Enabled = True
            If Me.BcrBtn = True Then Me.BCR_ID_TextBox.Enabled = True
        End If
    Else
        Me.BCWS_Checkbox.Enabled = False
        Me.BCWP_Checkbox.Enabled = False
        Me.ETC_Checkbox.Enabled = False
        Me.WhatIf_CheckBox.Enabled = False 'v3.2
        Me.TotalProjBtn.Enabled = False
        Me.ResExportCheckbox.Enabled = False
        Me.exportTPhaseCheckBox.Enabled = False
        Me.BcrBtn.Enabled = False
        Me.BCR_ID_TextBox.Enabled = False
        Me.exportDescCheckBox.Enabled = False
        Me.Milestone_CheckBox.Enabled = False 'v3.4
        Me.WeekStartCombobox.Enabled = False
        Me.WeekStartLabel.Enabled = False 'v3.4
        Me.ScaleCombobox.Enabled = False 'v3.4
        Me.ScaleLabel.Enabled = False 'v3.4
    End If
    
    If Me.TotalProjBtn = False And Me.BcrBtn = False Then
        Me.TotalProjBtn = True
    End If
    
    Me.Tag = "Loading"
    Me.TabButtons(1).Tag = PopulateCustFieldUsage
    Me.Tag = "Loaded"

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    CancelBtn_Click
  End If
End Sub
Private Function VerifyCustFieldUsage() As Boolean

    Dim fCAID1, fCAID2, fCAID3, fWP, fCAM, fEVT, fPCNT, fResID, dateFmt As Boolean 'v3.3.5
    
    If Me.caID1Box.value <> "" Then fCAID1 = True
    If CAID2TxtBox.value <> "<None>" Then
        If Me.caID2Box.value <> "" Then fCAID2 = True
    Else
        fCAID2 = False
    End If
    If CAID3TxtBox.value <> "<None>" Then
        If Me.caID3Box.value <> "" Then fCAID3 = True
    Else
        fCAID3 = False
    End If
    If Me.resBox.value <> "" Then fResID = True Else fResID = False 'v3.2.2
    If Me.wpBox.value <> "" Then fWP = True Else fWP = False 'v3.2.2
    If Me.camBox.value <> "" Then fCAM = True Else fCAM = False 'v3.2.2
    If Me.evtBox.value <> "" Then fEVT = True Else fEVT = False 'v3.2.2
    If Me.PercentBox.value <> "" Then fPCNT = True Else fPCNT = False 'v3.2.2
    If Me.DateFormat_Combobox.value <> "" Then dateFmt = True Else dateFmt = False 'v3.3.5
    
    If fCAID1 And fCAID2 And fCAID3 And fWP And fCAM And fEVT And fPCNT And fResID And dateFmt Then 'v3.3.5
    
        VerifyCustFieldUsage = True
    
    Else
    
        VerifyCustFieldUsage = False
    
    End If

End Function

Private Function VerifyTitles() As Boolean

    Dim TitlesComplete As Boolean
    
    TitlesComplete = True
    
    If Me.CAID1TxtBox.value = "" Then
        Me.CAID1TxtBox.BackColor = RGB(255, 255, 0)
        TitlesComplete = False
    Else
        Me.CAID1TxtBox.BackColor = RGB(255, 255, 255)
    End If
    
    If Me.caID2Box.value <> "<None>" Then
        If Me.CAID2TxtBox.value = "" Then
            Me.CAID2TxtBox.BackColor = RGB(255, 255, 0)
            TitlesComplete = False
        Else
            Me.CAID2TxtBox.BackColor = RGB(255, 255, 255)
        End If
    End If
    
    If Me.caID3Box.value <> "<None>" Then
        If Me.CAID3TxtBox.value = "" Then
            Me.CAID3TxtBox.BackColor = RGB(255, 255, 0)
            TitlesComplete = False
        Else
            Me.CAID3TxtBox.BackColor = RGB(255, 255, 255)
        End If
    End If
    
    VerifyTitles = TitlesComplete

End Function
Private Function PopulateCustFieldUsage() As Boolean

    Dim curProj As Project
    Dim docProp As DocumentProperty
    Dim docProps As DocumentProperties
    Dim fProject, fCAID1, fCAID1t, fCAID3, fCAID3t, fWP, fCAM, fEVT, fCAID2, fCAID2t, fMSID, fMSW, fBCR, fWhatIf, fPCNT, fAssignPcnt, fResID, dateFmt As Boolean 'v3.3.0, v3.3.5, v3.4.3
    Dim nameTest As Double
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo DocPropNameChange
    
    For Each docProp In docProps
    
        Select Case docProp.Name
        
            Case "dateFmt" 'v3.3.5
            
                dateFmt = True
                Me.DateFormat_Combobox.value = docProp.value
        
            Case "fAssignPcnt"
            
                If docProp.value = "<None>" Then 'v3.3.3 - testing for "None"
                    fAssignPcnt = True
                    Me.AsgnPcntBox.value = docProp.value
                Else
                    nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                    fAssignPcnt = True
                    Me.AsgnPcntBox.value = docProp.value
                End If
        
            Case "fCAID1"
            
                nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                fCAID1 = True
                Me.caID1Box.value = docProp.value
                
            Case "fCAID1t"
            
                fCAID1t = True
                Me.CAID1TxtBox.value = docProp.value
                
            Case "fCAID3"
                
                If docProp.value = "<None>" Then
                    fCAID3 = True
                    Me.caID3Box.value = docProp.value
                Else
                    nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                    fCAID3 = True
                    Me.caID3Box.value = docProp.value
                End If
                
            Case "fCAID3t"
                
                fCAID3t = True
                Me.CAID3TxtBox.value = docProp.value
                
            Case "fCAID2"
            
                If docProp.value = "<None>" Then
                    fCAID2 = True
                    Me.caID2Box.value = docProp.value
                Else
                    nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                    fCAID2 = True
                    Me.caID2Box.value = docProp.value
                End If
                
            Case "fCAID2t"
            
                fCAID2t = True
                Me.CAID2TxtBox.value = docProp.value
                
            Case "fWP"
                
                nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                fWP = True
                Me.wpBox.value = docProp.value
                
            Case "fCAM"
                
                nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                fCAM = True
                Me.camBox.value = docProp.value
                
            Case "fEVT"
                
                nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                fEVT = True
                Me.evtBox.value = docProp.value
                
            Case "fCAID2"
            
                nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                fCAID2 = True
                Me.caID2Box.value = docProp.value
                
            Case "fCAID2t"
            
                fCAID2t = True
                Me.CAID2TxtBox.value = docProp.value
                
            Case "fMSID"
                
                If docProp.value = "<None>" Then
                    fMSID = True
                    Me.msidBox.value = docProp.value
                Else
                    nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                    fMSID = True
                    Me.msidBox.value = docProp.value
                End If
                
            Case "fMSW"
                
                If docProp.value = "<None>" Then
                    fMSW = True
                    Me.mswBox.value = docProp.value
                Else
                    nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                    fMSW = True
                    Me.mswBox.value = docProp.value
                End If
                
            Case "fBCR"
            
                If docProp.value = "<None>" Then
                    fBCR = True
                    Me.bcrBox.value = docProp.value
                Else
                    nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                    fBCR = True
                    Me.bcrBox.value = docProp.value
                End If
                
            Case "fProject"
            
                If docProp.value = "<None>" Then
                    fProject = True
                    Me.projBox.value = docProp.value
                Else
                    nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                    fProject = True
                    Me.projBox.value = docProp.value
                End If
                
            Case "fWhatIf" 'v3.2
            
                If docProp.value = "<None>" Then
                    fWhatIf = True
                    Me.whatifBox.value = docProp.value
                Else
                    nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                    fWhatIf = True
                    Me.whatifBox.value = docProp.value
                End If
                
            Case "fPCNT"
            
                nameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.value)
                fPCNT = True
                Me.PercentBox.value = docProp.value
                
            Case "fResID"
            
                fResID = True
                Me.resBox.value = docProp.value
            
            Case Else
        
        End Select
    
NextDocProp:
    
    Next docProp
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    If fCAID1 And fCAID2 And fWP And fCAM And fEVT And fCAID3 And fPCNT And fResID And dateFmt Then 'v3.2.6, v3.3.5
    
        PopulateCustFieldUsage = True
    
    Else
    
        PopulateCustFieldUsage = False
    
    End If
    
    Exit Function
    
DocPropNameChange:

    Resume NextDocProp

End Function

Private Sub WeekStartCombobox_Change() 'v3.3.6
    'sets project "week starts on" value
    Dim curProj As Project
    Set curProj = ActiveProject
    curProj.StartWeekOn = WeekStartCombobox.ListIndex + 1
    
    Set curProj = Nothing 'v3.4
    
End Sub

Private Sub WhatIf_CheckBox_Click() 'v3.2
    If Me.WhatIf_CheckBox.value = True Then
        Me.exportTPhaseCheckBox.Enabled = True
        Me.BcrBtn.Enabled = True 'v3.3.15
        Me.TotalProjBtn.Enabled = True 'v3.3.15
        Me.Milestone_CheckBox.Enabled = True 'v3.4.1
    Else
        If Me.BCWS_Checkbox.value = False Then
            Me.exportTPhaseCheckBox.Enabled = False
            Me.BcrBtn.Enabled = False 'v3.3.15
            Me.TotalProjBtn.Enabled = False 'v3.3.15
            Me.BCR_ID_TextBox.Enabled = False 'v3.3.15
            Me.Milestone_CheckBox.Enabled = False 'v3.4.1
        End If
    End If
End Sub

Private Sub whatifBox_Change() 'v3.2
    If checkDuplicate(whatifBox) = True Then
        MsgBox "Please select a unique IMS Field."
        whatifBox.value = ""
        Exit Sub
    End If
    
    If isIMSfield(whatifBox.value) = False And whatifBox.value <> "" And whatifBox.value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        whatifBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fWhatIf").value = Me.whatifBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fWhatIf", False, msoPropertyTypeString, Me.whatifBox.value
    Resume PropFound
End Sub

Private Sub wpBox_Change()

    If checkDuplicate(wpBox) = True Then
        MsgBox "Please select a unique IMS Field."
        wpBox.value = ""
        Exit Sub
    End If
    
    If isIMSfield(wpBox.value) = False And wpBox.value <> "" Then
        MsgBox "Please select a valid IMS Field."
        wpBox.value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fWP").value = Me.wpBox.value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fWP", False, msoPropertyTypeString, Me.wpBox.value
    Resume PropFound
End Sub
Private Function isIMSfield(ByVal mspField As String) As Boolean

    On Error GoTo fieldMissing
    
    Dim curProj As Project
    Set curProj = ActiveProject
    
    If curProj.Application.FieldNameToFieldConstant(mspField) Then
    
        isIMSfield = True
        Set curProj = Nothing
        Exit Function
    
    End If
    
fieldMissing:

    isIMSfield = False
    Set curProj = Nothing

End Function
