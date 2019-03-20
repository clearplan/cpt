Attribute VB_Name = "cptIMSCobraExport_bas"
'<cpt_version>v3.0.1</cpt_version>

Global destFolder As String
Global BCWSxport As Boolean
Global BCWPxport As Boolean
Global ETCxport As Boolean
Global ResourceLoaded As Boolean
Global MasterProject As Boolean
Global ACTfilename As String
Global RESfilename As String
Global BCR_WP() As String
Global BCR_ID As String
Global BCRxport As Boolean
Global BCR_Error As Boolean
Global fCAID1, fCAID1t, fCAID3, fCAID3t, fWP, fCAM, fPCNT, fEVT, fCAID2, fCAID2t, fMilestone, fMilestoneWeight, fBCR, fResID As String
Global CustTextFields() As String
Global EntFields() As String
Global CustNumFields() As String
Global CustOLCodeFields() As String
Global ActFound As Boolean
Global CAID3_Used As Boolean
Global CAID2_Used As Boolean
Global TimeScaleExport As Boolean

Private Type ACTrowWP

    CAID1 As String
    CAID3 As String
    CAID2 As String
    CAM As String
    WP As String
    ID As String
    BStart As Date
    BFinish As Date
    FStart As Date
    FFinish As Date
    AStart As Date
    AFinish As Date
    EVT As String
    sumBCWS As Double
    sumBCWP As Double
    Prog As Integer

End Type

Private Type WPDataCheck

    WP_ID As String
    ID_Test As String
    EVT_Test As String
    WP_DupError As Boolean
    EVT_Error As Boolean

End Type

Private Type CAMDataCheck
    
    ID_str As String '**CAID1/CAID2/CAID3**
    CAM_Test As String
    CAM_Error As Boolean
    
End Type

Private Type TaskDataCheck

    UID As String
    WP As String
    CAID1 As String
    CAID2 As String
    CAID3 As String
    CAM As String
    BStart As String
    BFinish As String
    BWork As Double
    BCost As Double
    AssignmentBStart As String
    AssignmentBFinish As String
    AssignmentBWork As Double
    AssignmentBCost As Double
    EVT As String
    MSID As String
    MSWeight As String
    AssignmentCount As Integer

End Type

Sub Export_IMS()

    Dim xportFrm As cptIMSCobraExport_frm
    Dim xportFormat As String
    Dim curProj As Project
    
    On Error GoTo CleanUp
    
    Set curProj = ActiveProject
    
    curProj.Application.Calculation = pjManual
    curProj.Application.DisplayAlerts = False
    
    If curProj.Subprojects.count > 0 And InStr(curProj.FullName, "<>") > 0 And curProj.ReadOnly <> True Then
        MsgBox "Master Project Files with Subprojects must be opened Read Only"
        GoTo Quick_Exit
    End If
    
    If curProj.Subprojects.count > 0 Then
        MasterProject = True
    Else
        MasterProject = False
    End If
    
    ReadCustomFields curProj
    
    Set xportFrm = New cptIMSCobraExport_frm
    
    With cptIMSCobraExport_frm
    
        On Error Resume Next
        
        .msidBox.AddItem "UniqueID"
        .mswBox.AddItem "BaselineWork"
        .mswBox.AddItem "BaselineCost"
        .mswBox.AddItem "Work"
        .mswBox.AddItem "Cost"
        .camBox.AddItem "Contact"
        .caID1Box.AddItem "WBS"
        .caID2Box.AddItem "<None>"
        .caID3Box.AddItem "<None>"
        .bcrBox.AddItem "<None>"
        .PercentBox.AddItem "Physical % Complete"
        .PercentBox.AddItem "% Complete"
        
        .resBox.AddItem "Name"
        .resBox.AddItem "Code"
        .resBox.AddItem "Initials"
        
        For i = 1 To UBound(CustNumFields)
            .msidBox.AddItem CustNumFields(i)
            .mswBox.AddItem CustNumFields(i)
            .PercentBox.AddItem CustNumFields(i)
        Next i
        For i = 1 To UBound(EntFields)
            .caID1Box.AddItem EntFields(i)
            .caID3Box.AddItem EntFields(i)
            .wpBox.AddItem EntFields(i)
            .camBox.AddItem EntFields(i)
            .evtBox.AddItem EntFields(i)
            .caID2Box.AddItem EntFields(i)
            .msidBox.AddItem EntFields(i)
            .mswBox.AddItem EntFields(i)
            .bcrBox.AddItem EntFields(i)
        Next i
        For i = 1 To UBound(CustTextFields)
            .caID1Box.AddItem CustTextFields(i)
            .caID3Box.AddItem CustTextFields(i)
            .wpBox.AddItem CustTextFields(i)
            .camBox.AddItem CustTextFields(i)
            .evtBox.AddItem CustTextFields(i)
            .caID2Box.AddItem CustTextFields(i)
            .msidBox.AddItem CustTextFields(i)
            .mswBox.AddItem CustTextFields(i)
            .bcrBox.AddItem CustTextFields(i)
        Next i
        For i = 1 To UBound(CustOLCodeFields)
            .caID1Box.AddItem CustOLCodeFields(i)
            .caID3Box.AddItem CustOLCodeFields(i)
            .wpBox.AddItem CustOLCodeFields(i)
            .camBox.AddItem CustOLCodeFields(i)
            .caID2Box.AddItem CustOLCodeFields(i)
            .evtBox.AddItem CustOLCodeFields(i)
            .caID2Box.AddItem CustOLCodeFields(i)
            .msidBox.AddItem CustOLCodeFields(i)
            .bcrBox.AddItem CustOLCodeFields(i)
        Next i
        
        On Error GoTo CleanUp
        'On Error GoTo 0 '**Used for Debugging ONLY**
        
        .Show
        
        If .Tag = "Cancel" Then
            Set cptIMSCobraExport_frm = Nothing
            Set curProj = Nothing
            Exit Sub
        End If
        
        If .Tag = "DataCheck" Then
            CAID3_Used = .CAID3TxtBox.Enabled
            CAID2_Used = .CAID2TxtBox.Enabled
            DataChecks curProj
            Set curProj = Nothing
            Set cptIMSCobraExport_frm = Nothing
            Exit Sub
        End If
        
        If .MPPBtn.Value = True Then
            Set cptIMSCobraExport_frm = Nothing
            xportFormat = "MPP"
        ElseIf .XMLBtn.Value = True Then
            Set cptIMSCobraExport_frm = Nothing
            xportFormat = "XML"
        ElseIf .CSVBtn.Value = True Then
            BCWSxport = .BCWS_Checkbox.Value
            BCWPxport = .BCWP_Checkbox.Value
            ETCxport = .ETC_Checkbox.Value
            BCRxport = .BcrBtn.Value
            BCR_ID = .BCR_ID_TextBox
            ResourceLoaded = .ResExportCheckbox
            TimeScaleExport = .exportTPhaseCheckBox
            Set cptIMSCobraExport_frm = Nothing
            xportFormat = "CSV"
            CAID3_Used = .CAID3TxtBox.Enabled
            CAID2_Used = .CAID2TxtBox.Enabled
        End If
        
    End With
    
    Select Case xportFormat
    
        Case "MPP"
        
            MPP_Export curProj
        
        Case "XML"
        
            XML_Export curProj
        
        Case "CSV"
        
            CSV_Export curProj
        
        Case Else
        
    End Select
    
    If BCR_Error = False Then
        MsgBox "IMS Export saved to " & destFolder
        Shell "explorer.exe" & " " & destFolder, vbNormalFocus
    End If
    
    curProj.Application.Calculation = pjAutomatic
    curProj.Application.DisplayAlerts = True
    Set curProj = Nothing
    
    Exit Sub
    
CleanUp:

    If ACTfilename <> "" Then Reset

    curProj.Application.Calculation = pjAutomatic
    curProj.Application.DisplayAlerts = True
    Set curProj = Nothing
    MsgBox "An error was encountered." & vbCr & vbCr & "Please try again, or contact the developer if this message repeats."
    Exit Sub
    
Quick_Exit:
    
    curProj.Application.Calculation = pjAutomatic
    curProj.Application.DisplayAlerts = True
    Set curProj = Nothing
    
    Exit Sub
    
End Sub

Private Sub DataChecks(ByVal curProj As Project)

    Dim WPChecks() As WPDataCheck
    Dim wpFound As Boolean
    Dim CAMChecks() As CAMDataCheck
    Dim CAfound As Boolean
    Dim TaskChecks() As TaskDataCheck
    Dim taskFound As Boolean
    Dim t As Task
    Dim tAss As Assignment
    Dim tAsses As Assignments
    Dim tAssBStart As String
    Dim tAssBFin As String
    Dim tAssBWork As String
    Dim tempID As String
    Dim subProj As Subproject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim wpCount As Integer
    Dim camCount As Integer
    Dim TaskCount As Integer
    Dim x As Integer
    Dim i As Integer
    Dim errorStr As String
    Dim ErrorCounter As Integer
    Dim tempBValue As Double
    Dim tempBWork As Double
    
    Dim docProps As DocumentProperties
    
    Set docProps = curProj.CustomDocumentProperties
    
    fCAID1 = docProps("fCAID1").Value
    fCAID1t = docProps("fCAID1t").Value
    If BCRxport = True Then
        fBCR = docProps("fBCR").Value
    End If
    If CAID3_Used = True Then
        fCAID3 = docProps("fCAID3").Value
        fCAID3t = docProps("fCAID3t").Value
    End If
    fWP = docProps("fWP").Value
    fCAM = docProps("fCAM").Value
    fEVT = docProps("fEVT").Value
    fMilestone = docProps("fMSID").Value
    If CAID2_Used = True Then
        fCAID2 = docProps("fCAID2").Value
        fCAID2t = docProps("fCAID2t").Value
    End If
    fMilestoneWeight = docProps("fMSW").Value
    fPCNT = docProps("fPCNT").Value
    
    destFolder = SetDirectory(curProj.ProjectSummaryTask.Name)
    
    TaskCount = 0
    taskFound = False
    
    '**Scan Task Data**
    
    If curProj.Subprojects.count > 0 Then
    
        Set subProjs = curProj.Subprojects
        
        For Each subProj In subProjs
        
            FileOpen Name:=subProj.Path, ReadOnly:=True
            
            Set curSProj = ActiveProject
            
            For Each t In curSProj.Tasks
            
                If Not t Is Nothing Then
                
                    If t.Summary = False And t.Active = True And t.ExternalTask = False Then
                    
                        TaskCount = TaskCount + 1
                        taskFound = True
                        ReDim Preserve TaskChecks(1 To TaskCount)
                        
                        With TaskChecks(TaskCount)
                        
                            .UID = t.UniqueID
                            .WP = t.GetField(FieldNameToFieldConstant(fWP))
                            .CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID2_Used = True Then
                                .CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            If CAID3_Used = True Then
                                .CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            .EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            .MSID = t.GetField(FieldNameToFieldConstant(fMilestone))
                            .MSWeight = t.GetField(FieldNameToFieldConstant(fMilestoneWeight))
                            .BWork = t.BaselineWork / 60 / curSProj.HoursPerDay
                            .BCost = t.BaselineCost
                            .CAM = t.GetField(FieldNameToFieldConstant(fCAM))
                            .AssignmentBStart = "NA"
                            .AssignmentBFinish = "NA"
                            .AssignmentBCost = 0
                            .AssignmentBWork = 0
                            .BStart = t.BaselineStart
                            .BFinish = t.BaselineFinish
                            
                            Set tAsses = t.Assignments
                            .AssignmentCount = tAsses.count
                            
                            For Each tAss In tAsses
                            
                                If tAss.BaselineStart <> "NA" Then
                                    If .AssignmentBStart = "NA" Then
                                        .AssignmentBStart = tAss.BaselineStart
                                    Else
                                        If tAss.BaselineStart < .AssignmentBStart Then
                                            .AssignmentBStart = tAss.BaselineStart
                                        End If
                                    End If
                                End If
                                
                                If tAss.BaselineFinish <> "NA" Then
                                    If .AssignmentBFinish = "NA" Then
                                        .AssignmentBFinish = tAss.BaselineFinish
                                    Else
                                        If tAss.BaselineFinish > .AssignmentBFinish Then
                                            .AssignmentBFinish = tAss.BaselineFinish
                                        End If
                                    End If
                                End If
                                
                                .AssignmentBCost = .AssignmentBCost + tAss.BaselineCost
                                If tAss.BaselineWork = "" Then
                                    tempBWork = 0
                                Else
                                    tempBWork = tAss.BaselineWork
                                End If
                                .AssignmentBWork = .AssignmentBWork + tempBWork / 60 / curSProj.HoursPerDay
                            
                            Next tAss
                            
                        End With
                        
                    End If
                    
                End If
                
            Next t
                
            FileClose pjDoNotSave
        
        Next subProj
    
    Else
    
        For Each t In curProj.Tasks
        
            If Not t Is Nothing Then
            
                If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                
                    TaskCount = TaskCount + 1
                    taskFound = True
                    ReDim Preserve TaskChecks(1 To TaskCount)
                    
                    With TaskChecks(TaskCount)
                    
                        .UID = t.UniqueID
                        .WP = t.GetField(FieldNameToFieldConstant(fWP))
                        .CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                        If CAID2_Used = True Then
                            .CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                        End If
                        If CAID3_Used = True Then
                            .CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                        End If
                        .EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                        .MSID = t.GetField(FieldNameToFieldConstant(fMilestone))
                        .MSWeight = t.GetField(FieldNameToFieldConstant(fMilestoneWeight))
                        .BWork = t.BaselineWork / 60 / curProj.HoursPerDay
                        .BCost = t.BaselineCost
                        .CAM = t.GetField(FieldNameToFieldConstant(fCAM))
                        .AssignmentBStart = "NA"
                        .AssignmentBFinish = "NA"
                        .AssignmentBCost = 0
                        .AssignmentBWork = 0
                        .BStart = t.BaselineStart
                        .BFinish = t.BaselineFinish
                        
                        Set tAsses = t.Assignments
                        .AssignmentCount = tAsses.count
                        
                        For Each tAss In tAsses
                        
                            If tAss.BaselineStart <> "NA" Then
                                If .AssignmentBStart = "NA" Then
                                    .AssignmentBStart = tAss.BaselineStart
                                Else
                                    If tAss.BaselineStart < .AssignmentBStart Then
                                        .AssignmentBStart = tAss.BaselineStart
                                    End If
                                End If
                            End If
                            
                            If tAss.BaselineFinish <> "NA" Then
                                If .AssignmentBFinish = "NA" Then
                                    .AssignmentBFinish = tAss.BaselineFinish
                                Else
                                    If tAss.BaselineFinish > .AssignmentBFinish Then
                                        .AssignmentBFinish = tAss.BaselineFinish
                                    End If
                                End If
                            End If
                            
                            .AssignmentBCost = .AssignmentBCost + tAss.BaselineCost
                            If tAss.BaselineWork = "" Then
                                tempBWork = 0
                            Else
                                tempBWork = tAss.BaselineWork
                            End If
                            .AssignmentBWork = .AssignmentBWork + tempBWork / 60 / curProj.HoursPerDay
                        
                        Next tAss
                        
                    End With
                
                End If
            
            End If
            
        Next t
    
    End If
    
    ACTfilename = destFolder & "\DataChecks_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

    Open ACTfilename For Output As #1
    
    Print #1, "Tasks Missing Data - The following tasks are assumed to be EV Relevant Activities due to User population of one or more of the following data values: Work Package ID; Earned Value Technique; Baseline Work; Baseline Cost"
    
    If CAID3_Used = True And CAID2_Used = True Then
        Print #1, vbCrLf & "UID," & fCAID1t & "," & fCAID2t & "," & fCAID3t & ",CAM,WP,EVT,Baseline Value,Baseline Start,Baseline Finish,Milestone ID (As Req),Milestone Weight (As Req)"
    End If
    If CAID3_Used = False And CAID2_Used = True Then
        Print #1, vbCrLf & "UID," & fCAID1t & "," & fCAID2t & ",CAM,WP,EVT,Baseline Value,Baseline Start,Baseline Finish,Milestone ID (As Req),Milestone Weight (As Req)"
    End If
    If CAID3_Used = False And CAID2_Used = False Then
        Print #1, vbCrLf & "UID," & fCAID1t & ",CAM,WP,EVT,Baseline Value,Baseline Start,Baseline Finish,Milestone ID (As Req),Milestone Weight (As Req)"
    End If
    
    '**Evaluate WP and CAM data**
    wpCount = 0
    camCount = 0
    
    ErrorCounter = 0
    
    For x = 1 To TaskCount
    
        If CAID3_Used = True And CAID2_Used = True Then
            tempID = TaskChecks(x).CAID1 & "/" & TaskChecks(x).CAID2 & "/" & TaskChecks(x).CAID3
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            tempID = TaskChecks(x).CAID1 & "/" & TaskChecks(x).CAID2
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            tempID = TaskChecks(x).CAID1
        End If
        
        If TaskChecks(x).CAM <> "" And TaskChecks(x).WP <> "" Then
        
            CAfound = False
            
            If camCount = 0 Then
            
                camCount = 1
                
                ReDim CAMChecks(1 To camCount)
                
                With CAMChecks(camCount)
                
                    .ID_str = tempID
                    .CAM_Test = TaskChecks(x).CAM
                    .CAM_Error = False
                
                End With
            
            Else
            
                For i = 1 To camCount
                
                    If CAMChecks(i).ID_str = tempID Then
                    
                        CAfound = True
                        
                        If TaskChecks(x).CAM <> CAMChecks(i).CAM_Test Then
                            CAMChecks(i).CAM_Error = True
                        End If
                        
                        GoTo next_task
                    
                    End If
                
                Next i
                
                If CAfound = False Then
                
                    camCount = camCount + 1
                
                    ReDim Preserve CAMChecks(1 To camCount)
                    
                    With CAMChecks(camCount)
                    
                        .ID_str = tempID
                        .CAM_Test = TaskChecks(x).CAM
                        .CAM_Error = False
                    
                    End With
                
                End If
            
            End If
        
        End If
        
        If TaskChecks(x).WP <> "" Then
        
            wpFound = False
            
            If wpCount = 0 Then
            
                wpCount = 1
                
                ReDim WPChecks(1 To wpCount)
            
                With WPChecks(wpCount)
                
                    .ID_Test = tempID
                    .WP_ID = TaskChecks(x).WP
                    .EVT_Test = TaskChecks(x).EVT
                    .WP_DupError = False
                    .EVT_Error = False
                
                End With
                
            Else
                
                For i = 1 To wpCount
                
                    If WPChecks(i).WP_ID = TaskChecks(x).WP Then
                    
                        wpFound = True
                        
                        If tempID <> WPChecks(i).ID_Test Then
                        
                            WPChecks(i).WP_DupError = True

                        End If
                        
                        If TaskChecks(x).EVT <> WPChecks(i).EVT_Test Then
                        
                            WPChecks(i).EVT_Error = True
                            
                        End If
                        
                        GoTo next_task

                    End If
                
                Next i
                
                If wpFound = False Then
                
                    wpCount = wpCount + 1
                    
                    ReDim Preserve WPChecks(1 To wpCount)
                    
                    With WPChecks(wpCount)
                    
                        .ID_Test = tempID
                        .WP_ID = TaskChecks(x).WP
                        .EVT_Test = TaskChecks(x).EVT
                        .WP_DupError = False
                        .EVT_Error = False
                        
                    End With
                    
                End If
            
            End If
        
        End If
        
next_task:

        '**Report Tasks Missing Metadata**
        
        If TaskChecks(x).WP <> "" Or TaskChecks(x).EVT <> "" Or TaskChecks(x).BCost <> 0 Or TaskChecks(x).BWork <> 0 Then
        
            If TaskChecks(x).BWork = 0 Then tempBValue = TaskChecks(x).BCost Else tempBValue = TaskChecks(x).BWork
        
            If TaskChecks(x).WP = "" Or TaskChecks(x).EVT = "" Or TaskChecks(x).BStart = "NA" Or TaskChecks(x).BFinish = "NA" Or tempBValue = 0 Then
                
                ErrorCounter = ErrorCounter + 1
            
                If CAID3_Used = True And CAID2_Used = True Then
                
                    With TaskChecks(x)
                    
                        errorStr = .UID & ","
                        If .CAID1 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID1 & ","
                        End If
                        If .CAID2 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID2 & ","
                        End If
                        If .CAID3 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID3 & ","
                        End If
                        If .CAM = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAM & ","
                        End If
                        If .WP = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .WP & ","
                        End If
                        If .EVT = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .EVT & ","
                        End If
                        If .BCost = 0 And .BWork = 0 Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            If .BWork > 0 Then
                                errorStr = errorStr & .BWork & ","
                            Else
                                errorStr = errorStr & .BCost & ","
                            End If
                        End If
                        If .BStart = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BStart & ","
                        End If
                        If .BFinish = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BFinish & ","
                        End If
                        If .EVT = "B" Or .EVT = "B Milestone" Then
                            If .MSID = "" Then
                                errorStr = errorStr & "MISSING,"
                            Else
                                errorStr = errorStr & .MSID & ","
                            End If
                            
                            If .MSWeight = "" Then
                                errorStr = errorStr & "MISSING"
                            Else
                                errorStr = errorStr & .MSWeight
                            End If
                        Else
                            errorStr = errorStr & "N/A,N/A"
                        End If
                        
                    End With
                End If
                
                If CAID3_Used = False And CAID2_Used = True Then
                
                    With TaskChecks(x)
                    
                        errorStr = .UID & ","
                        If .CAID1 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID1 & ","
                        End If
                        If .CAID2 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID2 & ","
                        End If
                        If .CAM = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAM & ","
                        End If
                        If .WP = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .WP & ","
                        End If
                        If .EVT = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .EVT & ","
                        End If
                        If .BCost = 0 And .BWork = 0 Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            If .BWork > 0 Then
                                errorStr = errorStr & .BWork & ","
                            Else
                                errorStr = errorStr & .BCost & ","
                            End If
                        End If
                        If .BStart = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BStart & ","
                        End If
                        If .BFinish = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BFinish & ","
                        End If
                        If .EVT = "B" Or .EVT = "B Milestones" Then
                            If .MSID = "" Then
                                errorStr = errorStr & "MISSING,"
                            Else
                                errorStr = errorStr & .MSID & ","
                            End If
                            
                            If .MSWeight = "" Then
                                errorStr = errorStr & "MISSING"
                            Else
                                errorStr = errorStr & .MSWeight
                            End If
                        Else
                            errorStr = errorStr & "N/A,N/A"
                        End If
                        
                    End With
                    
                End If
                
                If CAID3_Used = False And CAID2_Used = False Then
                
                    With TaskChecks(x)
                    
                        errorStr = .UID & ","
                        If .CAID1 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID1 & ","
                        End If
                        If .CAM = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAM & ","
                        End If
                        If .WP = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .WP & ","
                        End If
                        If .EVT = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .EVT & ","
                        End If
                        If .BCost = 0 And .BWork = 0 Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            If .BWork > 0 Then
                                errorStr = errorStr & .BWork & ","
                            Else
                                errorStr = errorStr & .BCost & ","
                            End If
                        End If
                        If .BStart = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BStart & ","
                        End If
                        If .BFinish = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BFinish & ","
                        End If
                        If .EVT = "B" Or .EVT = "B Milestone" Then
                            If .MSID = "" Then
                                errorStr = errorStr & "MISSING,"
                            Else
                                errorStr = errorStr & .MSID & ","
                            End If
                            
                            If .MSWeight = "" Then
                                errorStr = errorStr & "MISSING"
                            Else
                                errorStr = errorStr & .MSWeight
                            End If
                        Else
                            errorStr = errorStr & "N/A,N/A"
                        End If
                        
                    End With
                    
                End If
                                 
                Print #1, errorStr
        
                errorStr = ""
            
            End If
            
        End If

    Next x
    
    Print #1, vbCrLf & "Total Task Errors Found: " & ErrorCounter
    
    '**Report Multiple CAM Assignments**
    
    Print #1, vbCrLf & vbCrLf & "CAM Errors - The following items reflect multiple CAM assignments per Control Account as interpreted by Cobra (based on a unique CA record ID constructed from Concatenated CA ID Values"
    
    Print #1, vbCrLf & "CA ID String"
    
    ErrorCounter = 0
    
    For x = 1 To camCount
    
        If CAMChecks(x).CAM_Error = True Then
        
            ErrorCounter = ErrorCounter + 1
        
            Print #1, CAMChecks(x).ID_str
        
        End If
    
    Next x
    
    Print #1, vbCrLf & "Total CAM Errors Found: " & ErrorCounter
    
    '**Report Duplicate WP IDs & Multiple EVTs**
    
    Print #1, vbCrLf & vbCrLf & "Work Package Errors - The following Work Package IDs are duplicated across multiple CA ID values and/or are assigned multiple EVTs"
    
    Print #1, vbCrLf & "Work Package,Duplicate WP ID,Multiple EVTs"
    
    ErrorCounter = 0
    
    For x = 1 To wpCount
    
        If WPChecks(x).WP_DupError = True Or WPChecks(x).EVT_Error = True Then
            
            ErrorCounter = ErrorCounter + 1
            
            With WPChecks(x)
                errorStr = .WP_ID & "," & .WP_DupError & "," & .EVT_Error
            End With
            
            Print #1, errorStr
            
            errorStr = ""
        End If
    
    Next x
    
    Print #1, vbCrLf & "Total Work Package Errors Found: " & ErrorCounter
    
    '**Reporting Assignment Baseline Issues (Values and Dates)**
    
    Print #1, vbCrLf & vbCrLf & "Task Assignment Discrepancies - The following Tasks have vertical traceability errors with their Assignment Baseline Values and/or Baseline Dates"
    
    Print #1, vbCrLf & "UID,Task Baseline Work,Assignment Baseline Work,Task Baseline Cost,Assignment Baseline Cost,Task Baseline Start,Assignment Baseline Start,Task Baseline Finish,Assignment Baseline Finish"
    
    ErrorCounter = 0
    
    For x = 1 To TaskCount
    
        With TaskChecks(x)
    
            If TaskChecks(x).AssignmentCount > 0 Then
    
                If Round(.BCost, 2) <> Round(.AssignmentBCost, 2) Or Round(.BWork, 2) <> Round(.AssignmentBWork, 2) Or .BStart <> .AssignmentBStart Or .BFinish <> .AssignmentBFinish Then
                
                    ErrorCounter = ErrorCounter + 1
                    
                    errorStr = .UID & ","
                    errorStr = errorStr & .BWork & ","
                    errorStr = errorStr & .AssignmentBWork & ","
                    errorStr = errorStr & .BCost & ","
                    errorStr = errorStr & .AssignmentBCost & ","
                    errorStr = errorStr & .BStart & ","
                    errorStr = errorStr & .AssignmentBStart & ","
                    errorStr = errorStr & .BFinish & ","
                    errorStr = errorStr & .AssignmentBFinish
                
                    Print #1, errorStr
                    errorStr = ""
                    
                End If
                
            End If
            
        End With
    
    Next x
    
    Print #1, vbCrLf & "Total Work Package Errors Found: " & ErrorCounter
    
    MsgBox "Data Check Report saved to " & destFolder
    
    Shell "explorer.exe" & " " & destFolder, vbNormalFocus
    
    Close #1

End Sub

Private Sub MPP_Export(ByVal curProj As Project)

    Dim subProj As Subproject
    Dim subProjs As Subprojects
    
    destFolder = SetDirectory(curProj.ProjectSummaryTask.Name)
    
    If curProj.Subprojects.count > 0 Then
        
        Set subProjs = curProj.Subprojects
        
        For Each subProj In subProjs
        
            subProj.SourceProject.SaveAs Name:=destFolder & "\" & subProj.SourceProject.Name
            curProj.Subprojects(subProj.Index).SourceProject = destFolder & "\" & subProj.SourceProject.Name
        
        Next subProj
        
        curProj.SaveAs Name:=destFolder & "\" & curProj.ProjectSummaryTask.Name
    
    Else
    
        curProj.SaveAs Name:=destFolder & "\" & curProj.ProjectSummaryTask.Name
    
    End If

End Sub
Private Sub XML_Export(ByVal curProj As Project)

    Dim subProj As Subproject
    Dim subProjs As Subprojects
    
    destFolder = SetDirectory(curProj.ProjectSummaryTask.Name)
    
    If curProj.Subprojects.count > 0 Then
        
        Set subProjs = curProj.Subprojects
        
        For Each subProj In subProjs
        
            subProj.SourceProject.SaveAs Name:=destFolder & "\" & subProj.SourceProject.Name, FormatID:="MSProject.XML"
        
        Next subProj

    
    Else
    
        curProj.SaveAs Name:=destFolder & "\" & curProj.ProjectSummaryTask.Name, FormatID:="MSProject.XML"
        
    End If

End Sub

Private Sub CSV_Export(ByVal curProj As Project)

    Dim docProps As DocumentProperties
    
    Set docProps = curProj.CustomDocumentProperties
    
    fCAID1 = docProps("fCAID1").Value
    fCAID1t = docProps("fCAID1t").Value
    If BCRxport = True Then
        fBCR = docProps("fBCR").Value
    End If
    If CAID3_Used = True Then
        fCAID3 = docProps("fCAID3").Value
        fCAID3t = docProps("fCAID3t").Value
    End If
    If CAID2_Used = True Then
        fCAID2 = docProps("fCAID2").Value
        fCAID2t = docProps("fCAID2t").Value
    End If
    fWP = docProps("fWP").Value
    fCAM = docProps("fCAM").Value
    fEVT = docProps("fEVT").Value
    fMilestone = docProps("fMSID").Value
    fMilestoneWeight = docProps("fMSW").Value
    fPCNT = docProps("fPCNT").Value
    fResID = docProps("fResID").Value
    
    BCR_Error = False
    
    destFolder = SetDirectory(curProj.ProjectSummaryTask.Name)
    
    '*******************
    '****BCR Review*****
    '*******************
    
    If BCRxport = True Then
        If Find_BCRs(curProj, fWP, fBCR, BCR_ID) = 0 Then
            MsgBox "BCR ID " & Chr(34) & BCR_ID & Chr(34) & " was not found in the IMS." & vbCr & vbCr & "Please try again."
            BCR_Error = True
            GoTo BCR_Error
        End If
    End If
    
    '*******************
    '****BCWS Export****
    '*******************
    
    If BCWSxport = True Then
    
        BCWS_Export curProj
    
    End If
    
    '*******************
    '****ETC Export****
    '*******************
    
    If ETCxport = True Then
    
        ETC_Export curProj
        
    End If
    
    '*******************
    '****BCWP Export****
    '*******************
    
    If BCWPxport = True Then
    
        BCWP_Export curProj
        
    End If
    
    Exit Sub
    
BCR_Error:

    If BCR_Error = True Then
        DeleteDirectory (destFolder)
    End If

End Sub

Private Sub BCWP_Export(ByVal curProj As Project)

    '*******************
    '****BCWP Export****
    '*******************

    Dim t As Task
    Dim tAss As Assignments
    Dim tAssign As Assignment
    Dim CAID1, CAID3, WP, CAM, EVT, UID, CAID2, MSWeight, ID, PCNT As String
    Dim Milestone As String
    Dim subProj As Subproject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim x As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String

    If ResourceLoaded = False Then
    
        ACTfilename = destFolder & "\BCWP ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
    
        Open ACTfilename For Output As #1
        
        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete"
        End If
    
        x = 1
        ActFound = False
        
        If curProj.Subprojects.count > 0 Then
            
            Set subProjs = curProj.Subprojects
            
            For Each subProj In subProjs
            
                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject
                
                For Each t In curSProj.Tasks
                
                    If Not t Is Nothing Then
                        
                        If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                        
                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then
                        
                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                
                                If EVT = "B" Or EVT = "N" Or EVT = "B Milestone" Or EVT = "N Earning Rules" Then
                                    
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                
                                ElseIf EVT = "C" Or EVT = "C % Work Complete" Then
                                
                                    'store ACT info
                                    'WP Data
                                    If x = 1 Then
                                        
                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To x)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(x).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(x).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(x).CAID3 = CAID3
                                        End If
                                        ACTarray(x).CAM = CAM
                                        ACTarray(x).ID = ID
                                        ACTarray(x).CAID1 = CAID1
                                        ACTarray(x).EVT = EVT
                                        ACTarray(x).FFinish = t.Finish
                                        ACTarray(x).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(x).CAID2 = CAID2
                                        End If
                                        ACTarray(x).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(x).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(x).AFinish = t.ActualFinish
                                        End If
                                        ACTarray(x).WP = WP
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(x).sumBCWS = 1
                                            ACTarray(x).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        Else
                                            ACTarray(x).sumBCWS = 1
                                            ACTarray(x).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        End If
                                        ACTarray(x).Prog = ACTarray(x).sumBCWP / ACTarray(x).sumBCWS * 100
                                        
                                        x = x + 1
                                        ActFound = True
                                        
                                        GoTo nrBCWP_WP_Match_A
                                    
                                    End If
                                        
                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = ID Then
                                            'Found an existing matching WP line
                                            If t.BaselineStart <> "NA" Then
                                                If ACTarray(i).BStart = 0 Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                Else
                                                    If ACTarray(i).BStart > t.BaselineStart Then
                                                        ACTarray(i).BStart = t.BaselineStart
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineFinish <> "NA" Then
                                                If ACTarray(i).BFinish = 0 Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                Else
                                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                                        ACTarray(i).BFinish = t.BaselineFinish
                                                    End If
                                                End If
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                            If t.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                Else
                                                    If t.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = t.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If t.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                Else
                                                    If t.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = t.ActualFinish
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineWork <> 0 Then
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + 1
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                            Else
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + 1
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                            End If
                                            
                                            ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100
                                            
                                            GoTo nrBCWP_WP_Match_A
                                        End If
                                    Next i
                                    
                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(i).sumBCWS = 1
                                        ACTarray(i).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(i).sumBCWS = 1
                                        ACTarray(i).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(x).Prog = ACTarray(x).sumBCWP / ACTarray(x).sumBCWS * 100
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                ElseIf EVT = "E" Or EVT = "F" Or EVT = "G" Or EVT = "H" Or EVT = "E 50/50" Or EVT = "F 0/100" Or EVT = "G 100/0" Or EVT = "H User Defined" Then
                                
                                    'store ACT info
                                    'WP Data
                                    If x = 1 Then
                                        
                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To x)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(x).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(x).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(x).CAID3 = CAID3
                                        End If
                                        ACTarray(x).CAM = CAM
                                        ACTarray(x).ID = ID
                                        ACTarray(x).CAID1 = CAID1
                                        ACTarray(x).EVT = EVT
                                        ACTarray(x).FFinish = t.Finish
                                        ACTarray(x).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(x).CAID2 = CAID2
                                        End If
                                        ACTarray(x).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(x).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(x).AFinish = t.ActualFinish
                                        End If
                                        ACTarray(x).WP = WP
                                        
                                        x = x + 1
                                        ActFound = True
                                        
                                        GoTo nrBCWP_WP_Match_A
                                    
                                    End If
                                        
                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = ID Then
                                            'Found an existing matching WP line
                                            If t.BaselineStart <> "NA" Then
                                                If ACTarray(i).BStart = 0 Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                Else
                                                    If ACTarray(i).BStart > t.BaselineStart Then
                                                        ACTarray(i).BStart = t.BaselineStart
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineFinish <> "NA" Then
                                                If ACTarray(i).BFinish = 0 Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                Else
                                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                                        ACTarray(i).BFinish = t.BaselineFinish
                                                    End If
                                                End If
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                            If t.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                Else
                                                    If t.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = t.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If t.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                Else
                                                    If t.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = t.ActualFinish
                                                    End If
                                                End If
                                            End If
                                            
                                            GoTo nrBCWP_WP_Match_A
                                        End If
                                    Next i
                                    
                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    ACTarray(x).WP = WP
                                    
                                    x = x + 1
                                    ActFound = True
                                
                                End If
                                
                            End If
                            
                        End If
                        
                    End If
                
nrBCWP_WP_Match_A:
                
                Next t
                
                FileClose pjDoNotSave
            
            Next subProj
        
        Else
        
            For Each t In curProj.Tasks
            
                If Not t Is Nothing Then
                    
                    If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                    
                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then
                    
                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            
                            If EVT = "B" Or EVT = "B Milestone" Or EVT = "N" Or EVT = "N Earning Rules" Then
                            
                                If CAID3_Used = True And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    Print #1, CAID1 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                
                            ElseIf EVT = "C" Or EVT = "C % Work Complete" Then
                            
                                'store ACT info
                                'WP Data
                                If x = 1 Then
                                        
                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(x).sumBCWS = 1
                                        ACTarray(x).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(x).sumBCWS = 1
                                        ACTarray(x).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(x).Prog = ACTarray(x).sumBCWP / ACTarray(x).sumBCWS * 100
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                    GoTo nrBCWP_WP_Match_B
                                
                                End If
                                    
                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + 1
                                            ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                        Else
                                            ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + 1
                                            ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                        End If
                                        ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100
                                        
                                        GoTo nrBCWP_WP_Match_B
                                    End If
                                Next i
                                
                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To x)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(x).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(x).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).ID = ID
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(x).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(x).AFinish = t.ActualFinish
                                End If
                                ACTarray(x).WP = WP
                                If t.BaselineWork <> 0 Then
                                    ACTarray(x).sumBCWS = 1
                                    ACTarray(x).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                Else
                                    ACTarray(x).sumBCWS = 1
                                    ACTarray(x).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                End If
                                ACTarray(x).Prog = ACTarray(x).sumBCWP / ACTarray(x).sumBCWS * 100
                                
                                x = x + 1
                                ActFound = True
                                
                            ElseIf EVT = "E" Or EVT = "E 50/50" Or EVT = "F" Or EVT = "F 0/100" Or EVT = "G" Or EVT = "G 100/0" Or EVT = "H" Or EVT = "H User Defined" Then
                                
                                'store ACT info
                                'WP Data
                                If x = 1 Then
                                    
                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    ACTarray(x).WP = WP
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                    GoTo nrBCWP_WP_Match_B
                                
                                End If
                                    
                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        
                                        GoTo nrBCWP_WP_Match_B
                                    End If
                                Next i
                                
                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To x)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(x).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(x).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).ID = ID
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(x).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(x).AFinish = t.ActualFinish
                                End If
                                ACTarray(x).WP = WP
                                
                                x = x + 1
                                ActFound = True
                                
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
nrBCWP_WP_Match_B:

            Next t
    
        End If
        
        If ActFound = True Then
            For i = 1 To UBound(ACTarray)
            
                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, "M/D/YYYY")
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, "M/D/YYYY")
                
                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog
                End If
                
            Next i
        End If
        
        Close #1
    
    Else '**Resource Loaded**

        ACTfilename = destFolder & "\BCWP ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
    
        Open ACTfilename For Output As #1
        
        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete"
        End If
    
        x = 1
        ActFound = False
        
        If curProj.Subprojects.count > 0 Then
            
            Set subProjs = curProj.Subprojects
            
            For Each subProj In subProjs
            
                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject
                
                For Each t In curSProj.Tasks
                
                    If Not t Is Nothing Then
                        
                        If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                        
                            If t.BaselineWork > 0 Or t.BaselineCost > 0 Then
                        
                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                
                                If EVT = "B" Or EVT = "B Milestone" Or EVT = "N" Or EVT = "N Earned Rules" Then
                                    
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                
                                ElseIf EVT = "C" Or EVT = "C % Work Complete" Then
                                
                                    'store ACT info
                                    'WP Data
                                    If x = 1 Then
                                        
                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To x)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(x).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(x).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(x).CAID3 = CAID3
                                        End If
                                        ACTarray(x).CAM = CAM
                                        ACTarray(x).ID = ID
                                        ACTarray(x).CAID1 = CAID1
                                        ACTarray(x).EVT = EVT
                                        ACTarray(x).FFinish = t.Finish
                                        ACTarray(x).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(x).CAID2 = CAID2
                                        End If
                                        ACTarray(x).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(x).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(x).AFinish = t.ActualFinish
                                        End If
                                        ACTarray(x).WP = WP
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(x).sumBCWS = t.BaselineWork / 60
                                            ACTarray(x).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        Else
                                            ACTarray(x).sumBCWS = t.BaselineCost
                                            ACTarray(x).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        End If
                                        ACTarray(x).Prog = ACTarray(x).sumBCWP / ACTarray(x).sumBCWS * 100
                                        
                                        x = x + 1
                                        ActFound = True
                                        
                                        GoTo BCWP_WP_Match_A
                                    
                                    End If
                                        
                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = ID Then
                                            'Found an existing matching WP line
                                            If t.BaselineStart <> "NA" Then
                                                If ACTarray(i).BStart = 0 Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                Else
                                                    If ACTarray(i).BStart > t.BaselineStart Then
                                                        ACTarray(i).BStart = t.BaselineStart
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineFinish <> "NA" Then
                                                If ACTarray(i).BFinish = 0 Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                Else
                                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                                        ACTarray(i).BFinish = t.BaselineFinish
                                                    End If
                                                End If
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                            If t.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                Else
                                                    If t.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = t.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If t.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                Else
                                                    If t.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = t.ActualFinish
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineWork <> 0 Then
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + t.BaselineWork / 60
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                            Else
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + t.BaselineCost
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                            End If
                                            
                                            ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100
                                            
                                            GoTo BCWP_WP_Match_A
                                        End If
                                    Next i
                                    
                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(i).sumBCWS = t.BaselineWork / 60
                                        ACTarray(i).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(i).sumBCWS = t.BaselineCost
                                        ACTarray(i).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(x).Prog = ACTarray(x).sumBCWP / ACTarray(x).sumBCWS * 100
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                ElseIf EVT = "E" Or EVT = "E 50/50" Or EVT = "F" Or EVT = "F 0/100" Or EVT = "G" Or EVT = "G 100/0" Or EVT = "H" Or EVT = "H User Defined" Then
                                
                                    'store ACT info
                                    'WP Data
                                    If x = 1 Then
                                        
                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To x)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(x).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(x).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(x).CAID3 = CAID3
                                        End If
                                        ACTarray(x).CAM = CAM
                                        ACTarray(x).ID = ID
                                        ACTarray(x).CAID1 = CAID1
                                        ACTarray(x).EVT = EVT
                                        ACTarray(x).FFinish = t.Finish
                                        ACTarray(x).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(x).CAID2 = CAID2
                                        End If
                                        ACTarray(x).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(x).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(x).AFinish = t.ActualFinish
                                        End If
                                        ACTarray(x).WP = WP
                                        
                                        x = x + 1
                                        ActFound = True
                                        
                                        GoTo BCWP_WP_Match_A
                                    
                                    End If
                                        
                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = ID Then
                                            'Found an existing matching WP line
                                            If t.BaselineStart <> "NA" Then
                                                If ACTarray(i).BStart = 0 Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                Else
                                                    If ACTarray(i).BStart > t.BaselineStart Then
                                                        ACTarray(i).BStart = t.BaselineStart
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineFinish <> "NA" Then
                                                If ACTarray(i).BFinish = 0 Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                Else
                                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                                        ACTarray(i).BFinish = t.BaselineFinish
                                                    End If
                                                End If
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                            If t.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                Else
                                                    If t.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = t.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If t.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                Else
                                                    If t.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = t.ActualFinish
                                                    End If
                                                End If
                                            End If
                                            
                                            GoTo BCWP_WP_Match_A
                                        End If
                                    Next i
                                    
                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    ACTarray(x).WP = w
                                    
                                    x = x + 1
                                    ActFound = True
                                
                                End If
                                
                            End If
                                
                        End If
                
                    End If
                
BCWP_WP_Match_A:
                
                Next t
                
                FileClose pjDoNotSave
            
            Next subProj
        
        Else
        
            For Each t In curProj.Tasks
            
                If Not t Is Nothing Then
                    
                    If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                    
                        If t.BaselineWork > 0 Or t.BaselineCost > 0 Then
                    
                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            
                            If EVT = "B" Or EVT = "B Milestone" Or EVT = "N" Or EVT = "N Earned Rules" Then
                            
                                If CAID3_Used = True And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    Print #1, CAID1 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                
                            ElseIf EVT = "C" Or EVT = "C % Work Complete" Then
                            
                                'store ACT info
                                'WP Data
                                If x = 1 Then
                                        
                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(x).sumBCWS = t.BaselineWork / 60
                                        ACTarray(x).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(x).sumBCWS = t.BaselineCost
                                        ACTarray(x).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(x).Prog = ACTarray(x).sumBCWP / ACTarray(x).sumBCWS * 100
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                    GoTo BCWP_WP_Match_B
                                
                                End If
                                    
                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + t.BaselineWork / 60
                                            ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                        Else
                                            ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + t.BaselineCost
                                            ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                        End If
                                        ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100
                                        
                                        GoTo BCWP_WP_Match_B
                                    End If
                                Next i
                                
                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To x)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(x).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(x).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).ID = ID
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(x).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(x).AFinish = t.ActualFinish
                                End If
                                ACTarray(x).WP = WP
                                If t.BaselineWork <> 0 Then
                                    ACTarray(x).sumBCWS = t.BaselineWork / 60
                                    ACTarray(x).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                Else
                                    ACTarray(x).sumBCWS = t.BaselineCost
                                    ACTarray(x).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                End If
                                ACTarray(x).Prog = ACTarray(x).sumBCWP / ACTarray(x).sumBCWS * 100
                                
                                x = x + 1
                                ActFound = True
                                
                            ElseIf EVT = "E" Or EVT = "E 50/50" Or EVT = "F" Or EVT = "F 0/100" Or EVT = "G" Or EVT = "G 100/0" Or EVT = "H" Or EVT = "H User Defined" Then
                                
                                'store ACT info
                                'WP Data
                                If x = 1 Then
                                    
                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    ACTarray(x).WP = WP
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                    GoTo BCWP_WP_Match_B
                                
                                End If
                                    
                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        
                                        GoTo BCWP_WP_Match_B
                                    End If
                                Next i
                                
                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To x)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(x).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(x).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).ID = ID
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(x).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(x).AFinish = t.ActualFinish
                                End If
                                ACTarray(x).WP = w
                                
                                x = x + 1
                                ActFound = True
                                
                            End If
                        
                        End If
                        
                    End If
                    
                End If
                
BCWP_WP_Match_B:

            Next t
    
        End If
        
        If ActFound = True Then
            For i = 1 To UBound(ACTarray)
            
                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, "M/D/YYYY")
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, "M/D/YYYY")
                
                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog
                End If
                
            Next i
        End If
        
        Close #1
        
    End If

End Sub

Private Sub ETC_Export(ByVal curProj As Project)

    Dim t As Task
    Dim tAss As Assignments
    Dim tAssign As Assignment
    Dim CAID1, CAID3, WP, CAM, EVT, UID, CAID2, MSWeight, ID, PCNT As String
    Dim Milestone As String
    Dim subProj As Subproject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim x As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String
    
    '*******************
    '****ETC Export****
    '*******************
    
    If ResourceLoaded = False Then
    
        ACTfilename = destFolder & "\ETC ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
    
        Open ACTfilename For Output As #1
        
        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
    
        x = 1
        ActFound = False
    
        If curProj.Subprojects.count > 0 Then
            
            Set subProjs = curProj.Subprojects
            
            For Each subProj In subProjs
            
                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject
                
                For Each t In curSProj.Tasks
                
                    If Not t Is Nothing Then
                    
                        If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                        
                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then
                        
                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                
                                'store ACT info
                                'WP Data
                                If x = 1 Then
                                
                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                    GoTo nrETC_WP_Match
                                
                                End If
                                
                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        GoTo nrETC_WP_Match
                                    End If
                                Next i
                                
                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To x)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(x).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(x).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).ID = ID
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(x).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(x).AFinish = t.ActualFinish
                                End If
                                
                                x = x + 1
                                ActFound = True
                            
                                'Milestone Data
nrETC_WP_Match:
        

                                
                            End If
                            
                        End If
                        
                    End If
                
                Next t
                
                FileClose pjDoNotSave
            
            Next subProj
        
        Else
        
            For Each t In curProj.Tasks
            
                If Not t Is Nothing Then
                
                    If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                    
                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then
                    
                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            
                            'store ACT info
                            'WP Data
                            If x = 1 Then
                                
                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To x)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(x).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(x).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).CAM = CAM
                                ACTarray(x).ID = ID
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(x).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(x).AFinish = t.ActualFinish
                                End If
    
                                x = x + 1
                                ActFound = True
                                
                                GoTo nrETC_WP_Match_B
                            
                            End If
                                
                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If t.BaselineStart <> "NA" Then
                                        If ACTarray(i).BStart = 0 Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        Else
                                            If ACTarray(i).BStart > t.BaselineStart Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            End If
                                        End If
                                    End If
                                    If t.BaselineFinish <> "NA" Then
                                        If ACTarray(i).BFinish = 0 Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        Else
                                            If ACTarray(i).BFinish < t.BaselineFinish Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            End If
                                        End If
                                    End If
                                    If ACTarray(i).FStart > t.Start Then
                                        ACTarray(i).FStart = t.Start
                                    End If
                                    If ACTarray(i).FFinish < t.Finish Then
                                        ACTarray(i).FFinish = t.Finish
                                    End If
                                    If t.ActualStart <> "NA" Then
                                        If ACTarray(i).AStart = 0 Then
                                            ACTarray(i).AStart = t.ActualStart
                                        Else
                                            If t.ActualStart < ACTarray(i).AStart Then
                                                ACTarray(i).AStart = t.ActualStart
                                            End If
                                        End If
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        If ACTarray(i).AFinish = 0 Then
                                            ACTarray(i).AFinish = t.ActualFinish
                                        Else
                                            If t.ActualFinish > ACTarray(i).AFinish Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            End If
                                        End If
                                    End If
    
                                    GoTo nrETC_WP_Match_B
                                End If
                            Next i
                            
                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To x)
                            If t.BaselineFinish <> "NA" Then
                                ACTarray(x).BFinish = t.BaselineFinish
                            End If
                            If t.BaselineStart <> "NA" Then
                                ACTarray(x).BStart = t.BaselineStart
                            End If
                            If CAID3_Used = True Then
                                ACTarray(x).CAID3 = CAID3
                            End If
                            ACTarray(x).CAM = CAM
                            ACTarray(x).ID = ID
                            ACTarray(x).CAID1 = CAID1
                            ACTarray(x).EVT = EVT
                            ACTarray(x).FFinish = t.Finish
                            ACTarray(x).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(x).CAID2 = CAID2
                            End If
                            ACTarray(x).WP = WP
                            If t.ActualStart <> "NA" Then
                                ACTarray(x).AStart = t.ActualStart
                            End If
                            If t.ActualFinish <> "NA" Then
                                ACTarray(x).AFinish = t.ActualFinish
                            End If
                            
                            x = x + 1
                            ActFound = True

nrETC_WP_Match_B:
        
                        End If
                    
                    End If
                    
                End If
            
            Next t
    
        End If
        
        If ActFound = True Then
            For i = 1 To UBound(ACTarray)
            
                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, "M/D/YYYY")
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, "M/D/YYYY")
            
                If aFinishString = "NA" Then
                    If CAID3_Used = True And CAID2_Used = True Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                    If CAID3_Used = False And CAID2_Used = True Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                    If CAID3_Used = False And CAID2_Used = False Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                End If
                
            Next i
        End If
        
        Close #1
    
    Else '**Resource Loaded**

        ACTfilename = destFolder & "\ETC ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
        RESfilename = destFolder & "\ETC RES_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
    
        Open ACTfilename For Output As #1
        Open RESfilename For Output As #2
        
        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        Print #2, "Cobra ID,Resource,Amount,From Date,To Date"
    
        x = 1
        ActFound = False
    
        If curProj.Subprojects.count > 0 Then
            
            Set subProjs = curProj.Subprojects
            
            For Each subProj In subProjs
            
                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject
                
                For Each t In curSProj.Tasks
                
                    If Not t Is Nothing Then
                    
                        If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                        
                            If t.Work > 0 Or t.Cost > 0 Then
                        
                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                
                                'store ACT info
                                'WP Data
                                If x = 1 Then
                                
                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To x)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(x).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(x).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(x).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(x).AFinish = t.ActualFinish
                                    End If
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                    GoTo ETC_WP_Match
                                
                                End If
                                
                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        GoTo ETC_WP_Match
                                    End If
                                Next i
                                
                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To x)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(x).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(x).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).ID = ID
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(x).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(x).AFinish = t.ActualFinish
                                End If
                                
                                x = x + 1
                                ActFound = True
                            
                                'Milestone Data
ETC_WP_Match:
        
                                
                                Set tAss = t.Assignments
                                
                                For Each tAssign In tAss
                                
                                    If TimeScaleExport = True Then
                                    
                                        ExportTimeScaleResources ID, t, tAssign, 2, "ETC"
                                        
                                    Else
                                
                                        Select Case tAssign.ResourceType
                                        
                                            Case pjResourceTypeWork
                                    
                                            If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                        
                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork / 60 & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                                
                                            ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                            
                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                            End If
                                            
                                        Case pjResourceTypeCost
                                        
                                            If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                        
                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingCost & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                                
                                            ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                            
                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                            End If
                                        
                                        Case pjResourceTypeMaterial
                                        
                                            If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                        
                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                                
                                            ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                            
                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                            End If
                                        
                                        End Select
                                        
                                    End If
                                                            
                                Next tAssign
                                
                            End If
                            
                        End If
                        
                    End If
                
                Next t
                
                FileClose pjDoNotSave
            
            Next subProj
        
        Else
        
            For Each t In curProj.Tasks
            
                If Not t Is Nothing Then
                
                    If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                    
                        If t.Work > 0 Or t.Cost > 0 Then
                    
                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            
                            'store ACT info
                            'WP Data
                            If x = 1 Then
                                
                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To x)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(x).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(x).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).CAM = CAM
                                ACTarray(x).ID = ID
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(x).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(x).AFinish = t.ActualFinish
                                End If
    
                                x = x + 1
                                ActFound = True
                                
                                GoTo ETC_WP_Match_B
                            
                            End If
                                
                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If t.BaselineStart <> "NA" Then
                                        If ACTarray(i).BStart = 0 Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        Else
                                            If ACTarray(i).BStart > t.BaselineStart Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            End If
                                        End If
                                    End If
                                    If t.BaselineFinish <> "NA" Then
                                        If ACTarray(i).BFinish = 0 Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        Else
                                            If ACTarray(i).BFinish < t.BaselineFinish Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            End If
                                        End If
                                    End If
                                    If ACTarray(i).FStart > t.Start Then
                                        ACTarray(i).FStart = t.Start
                                    End If
                                    If ACTarray(i).FFinish < t.Finish Then
                                        ACTarray(i).FFinish = t.Finish
                                    End If
                                    If t.ActualStart <> "NA" Then
                                        If ACTarray(i).AStart = 0 Then
                                            ACTarray(i).AStart = t.ActualStart
                                        Else
                                            If t.ActualStart < ACTarray(i).AStart Then
                                                ACTarray(i).AStart = t.ActualStart
                                            End If
                                        End If
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        If ACTarray(i).AFinish = 0 Then
                                            ACTarray(i).AFinish = t.ActualFinish
                                        Else
                                            If t.ActualFinish > ACTarray(i).AFinish Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            End If
                                        End If
                                    End If
    
                                    GoTo ETC_WP_Match_B
                                End If
                            Next i
                            
                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To x)
                            If t.BaselineFinish <> "NA" Then
                                ACTarray(x).BFinish = t.BaselineFinish
                            End If
                            If t.BaselineStart <> "NA" Then
                                ACTarray(x).BStart = t.BaselineStart
                            End If
                            If CAID3_Used = True Then
                                ACTarray(x).CAID3 = CAID3
                            End If
                            ACTarray(x).CAM = CAM
                            ACTarray(x).ID = ID
                            ACTarray(x).CAID1 = CAID1
                            ACTarray(x).EVT = EVT
                            ACTarray(x).FFinish = t.Finish
                            ACTarray(x).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(x).CAID2 = CAID2
                            End If
                            ACTarray(x).WP = WP
                            If t.ActualStart <> "NA" Then
                                ACTarray(x).AStart = t.ActualStart
                            End If
                            If t.ActualFinish <> "NA" Then
                                ACTarray(x).AFinish = t.ActualFinish
                            End If
    
                            
                            x = x + 1
                            ActFound = True
                            
                            'Milestone Data
ETC_WP_Match_B:
        
                            
                            Set tAss = t.Assignments
                            
                            For Each tAssign In tAss
                            
                                If TimeScaleExport = True Then
                                    
                                    ExportTimeScaleResources ID, t, tAssign, 2, "ETC"
                                        
                                Else
                                
                                    Select Case tAssign.ResourceType
                                    
                                        Case pjResourceTypeWork
                                
                                        If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                    
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork / 60 & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                        ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                        
                                        End If
                                        
                                    Case pjResourceTypeCost
                                    
                                        If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                    
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingCost & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                        ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                        
                                        End If
                                    
                                    Case pjResourceTypeMaterial
                                    
                                        If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                    
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                        ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                        
                                        End If
                                    
                                    End Select
                                    
                                End If
                                                        
                            Next tAssign
                            
                        End If
                    
                    End If
            
                End If
            
            Next t
    
        End If
        
        If ActFound = True Then
            For i = 1 To UBound(ACTarray)
            
                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, "M/D/YYYY")
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, "M/D/YYYY")
            
                If aFinishString = "NA" Then
                    If CAID3_Used = True And CAID2_Used = True Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                    If CAID3_Used = False And CAID2_Used = True Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                    If CAID3_Used = False And CAID2_Used = False Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                End If
                
            Next i
        End If
        
        Close #1
        Close #2
        
    End If

End Sub
Private Sub BCWS_Export(ByVal curProj As Project)

    Dim t As Task
    Dim tAss As Assignments
    Dim tAssign As Assignment
    Dim CAID1, CAID3, WP, CAM, EVT, UID, CAID2, MSWeight, ID, PCNT As String
    Dim Milestone As String
    Dim subProj As Subproject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim x As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String

    '*******************
    '****BCWS Export****
    '*******************
    
    If ResourceLoaded = False Then
    
        ACTfilename = destFolder & "\BCWS ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
    
        Open ACTfilename For Output As #1
        
        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
    
        x = 1
        ActFound = False
    
        If curProj.Subprojects.count > 0 Then
            
            Set subProjs = curProj.Subprojects
            
            For Each subProj In subProjs
            
                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject
                
                For Each t In curSProj.Tasks
                
                    If Not t Is Nothing Then
                    
                        If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                        
                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then
                        
                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                    
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                
                                If BCRxport = True Then
                                    If IsInArray(BCR_WP, WP) = False Then
                                        GoTo Next_nrSProj_Task
                                    End If
                                End If
                                
                                'store ACT info
                                'WP Data
                                If x = 1 Then
                                
                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To x)
                                    ACTarray(x).BFinish = t.BaselineFinish
                                    ACTarray(x).BStart = t.BaselineStart
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    ACTarray(x).CAID2 = CAID2
                                    ACTarray(x).WP = WP
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                    GoTo nrWP_Match
                                
                                End If
                                
                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If ACTarray(i).BStart > t.BaselineStart Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        End If
                                        If ACTarray(i).BFinish < t.BaselineFinish Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        
                                        GoTo nrWP_Match
                                    End If
                                Next i
                                
                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To x)
                                ACTarray(x).BFinish = t.BaselineFinish
                                ACTarray(x).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).ID = ID
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                ACTarray(x).CAID2 = CAID2
                                ACTarray(x).WP = WP
                                
                                x = x + 1
                                ActFound = True
                            
                                'Milestone Data
nrWP_Match:
    
                                If EVT = "B" Or EVT = "B Milestone" Then
                                
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                    End If
Next_nrSProj_Task:

                Next t
                
                FileClose pjDoNotSave
            
            Next subProj
        
        Else
        
            For Each t In curProj.Tasks
            
                If Not t Is Nothing Then
                
                    If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                    
                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then
                    
                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            
                            If BCRxport = True Then
                                If IsInArray(BCR_WP, WP) = False Then
                                    GoTo Next_nrTask
                                End If
                            End If
                            
                            'store ACT info
                            'WP Data
                            If x = 1 Then
                                
                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To x)
                                ACTarray(x).BFinish = t.BaselineFinish
                                ACTarray(x).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).ID = ID
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                ACTarray(x).CAID2 = CAID2
                                ACTarray(x).WP = WP
                                
                                x = x + 1
                                ActFound = True
                                
                                GoTo nrWP_Match_B
                            
                            End If
                                
                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If ACTarray(i).BStart > t.BaselineStart Then
                                        ACTarray(i).BStart = t.BaselineStart
                                    End If
                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                        ACTarray(i).BFinish = t.BaselineFinish
                                    End If
                                    If ACTarray(i).FStart > t.Start Then
                                        ACTarray(i).FStart = t.Start
                                    End If
                                    If ACTarray(i).FFinish < t.Finish Then
                                        ACTarray(i).FFinish = t.Finish
                                    End If
                                    GoTo nrWP_Match_B
                                End If
                            Next i
                            
                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To x)
                            ACTarray(x).BFinish = t.BaselineFinish
                            ACTarray(x).BStart = t.BaselineStart
                            If CAID3_Used = True Then
                                ACTarray(x).CAID3 = CAID3
                            End If
                            If CAID2_Used = True Then
                                ACTarray(x).CAID2 = CAID2
                            End If
                            ACTarray(x).CAM = CAM
                            ACTarray(x).CAID1 = CAID1
                            ACTarray(x).EVT = EVT
                            ACTarray(x).ID = ID
                            ACTarray(x).FFinish = t.Finish
                            ACTarray(x).FStart = t.Start
                            ACTarray(x).CAID2 = CAID2
                            ACTarray(x).WP = WP
                            
                            x = x + 1
                            ActFound = True
                            
                            'Milestone Data
nrWP_Match_B:
    
                            If EVT = "B" Or EVT = "B Milestone" Then
                                If CAID3_Used = True And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                            End If
                            
                        End If
                    
                    End If
                    
                End If
Next_nrTask:

            Next t
    
        End If
        
        If ActFound = True Then
            For i = 1 To UBound(ACTarray)
                
                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                
            Next i
        End If
        
        Close #1
    
    Else

        ACTfilename = destFolder & "\BCWS ACT_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
        RESfilename = destFolder & "\BCWS RES_" & RemoveIllegalCharacters(curProj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
    
        Open ACTfilename For Output As #1
        Open RESfilename For Output As #2
        
        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        Print #2, "Cobra ID,Resource,Amount,From Date,To Date"
    
        x = 1
        ActFound = False
    
        If curProj.Subprojects.count > 0 Then
            
            Set subProjs = curProj.Subprojects
            
            For Each subProj In subProjs
            
                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject
                
                For Each t In curSProj.Tasks
                
                    If Not t Is Nothing Then
                    
                        If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                        
                            If t.BaselineWork > 0 Or t.BaselineCost > 0 Then
                        
                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                
                                If BCRxport = True Then
                                    If IsInArray(BCR_WP, WP) = False Then
                                        GoTo Next_SProj_Task
                                    End If
                                End If
                                
                                'store ACT info
                                'WP Data
                                If x = 1 Then
                                
                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To x)
                                    ACTarray(x).BFinish = t.BaselineFinish
                                    ACTarray(x).BStart = t.BaselineStart
                                    If CAID3_Used = True Then
                                        ACTarray(x).CAID3 = CAID3
                                    End If
                                    ACTarray(x).ID = ID
                                    ACTarray(x).CAM = CAM
                                    ACTarray(x).CAID1 = CAID1
                                    ACTarray(x).EVT = EVT
                                    ACTarray(x).FFinish = t.Finish
                                    ACTarray(x).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(x).CAID2 = CAID2
                                    End If
                                    ACTarray(x).WP = WP
                                    
                                    x = x + 1
                                    ActFound = True
                                    
                                    GoTo WP_Match
                                
                                End If
                                
                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If ACTarray(i).BStart > t.BaselineStart Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        End If
                                        If ACTarray(i).BFinish < t.BaselineFinish Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        
                                        GoTo WP_Match
                                    End If
                                Next i
                                
                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To x)
                                ACTarray(x).BFinish = t.BaselineFinish
                                ACTarray(x).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).ID = ID
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                
                                x = x + 1
                                ActFound = True
                            
                                'Milestone Data
WP_Match:
    
                                If EVT = "B" Or EVT = "B Milestone" Then
                                
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    
                                End If
                                
                                Set tAss = t.Assignments
                                
                                For Each tAssign In tAss
                                
                                    If tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0 Then
                                    
                                        If TimeScaleExport = True Then
                                    
                                            ExportTimeScaleResources ID, t, tAssign, 2, "BCWS"
                                            
                                        Else
                                
                                            Select Case tAssign.ResourceType
                                            
                                                Case pjResourceTypeWork
                                        
                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                                                                    
                                                Case pjResourceTypeCost
                                            
                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                                            
                                                Case pjResourceTypeMaterial
                                                
                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                                            
                                            End Select
                                            
                                        End If
                                        
                                    End If
                                
                                Next tAssign
                                
                            End If
                            
                        End If
                        
                    End If
Next_SProj_Task:

                Next t
                
                FileClose pjDoNotSave
            
            Next subProj
        
        Else
        
            For Each t In curProj.Tasks
            
                If Not t Is Nothing Then
                
                    If t.Active = True And t.Summary = False And t.ExternalTask = False Then
                    
                        If t.BaselineWork > 0 Or t.BaselineCost > 0 Then
                    
                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            
                            If BCRxport = True Then
                                If IsInArray(BCR_WP, WP) = False Then
                                    GoTo next_task
                                End If
                            End If
                            
                            'store ACT info
                            'WP Data
                            If x = 1 Then
                                
                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To x)
                                ACTarray(x).BFinish = t.BaselineFinish
                                ACTarray(x).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(x).CAID3 = CAID3
                                End If
                                ACTarray(x).ID = ID
                                ACTarray(x).CAM = CAM
                                ACTarray(x).CAID1 = CAID1
                                ACTarray(x).EVT = EVT
                                ACTarray(x).FFinish = t.Finish
                                ACTarray(x).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(x).CAID2 = CAID2
                                End If
                                ACTarray(x).WP = WP
                                
                                x = x + 1
                                ActFound = True
                                
                                GoTo WP_Match_B
                            
                            End If
                                
                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If ACTarray(i).BStart > t.BaselineStart Then
                                        ACTarray(i).BStart = t.BaselineStart
                                    End If
                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                        ACTarray(i).BFinish = t.BaselineFinish
                                    End If
                                    If ACTarray(i).FStart > t.Start Then
                                        ACTarray(i).FStart = t.Start
                                    End If
                                    If ACTarray(i).FFinish < t.Finish Then
                                        ACTarray(i).FFinish = t.Finish
                                    End If
                                    GoTo WP_Match_B
                                End If
                            Next i
                            
                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To x)
                            ACTarray(x).BFinish = t.BaselineFinish
                            ACTarray(x).BStart = t.BaselineStart
                            If CAID3_Used = True Then
                                ACTarray(x).CAID3 = CAID3
                            End If
                            ACTarray(x).CAM = CAM
                            ACTarray(x).CAID1 = CAID1
                            ACTarray(x).EVT = EVT
                            ACTarray(x).ID = ID
                            ACTarray(x).FFinish = t.Finish
                            ACTarray(x).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(x).CAID2 = CAID2
                            End If
                            ACTarray(x).WP = WP
                            
                            x = x + 1
                            ActFound = True
                            
                            'Milestone Data
WP_Match_B:
    
                            If EVT = "B" Or EVT = "B Milestone" Then
                                If CAID3_Used = True And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                            End If
            
                            Set tAss = t.Assignments
                            
                            For Each tAssign In tAss
                            
                                If tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0 Then
                                
                                    If TimeScaleExport = True Then
                                    
                                        ExportTimeScaleResources ID, t, tAssign, 2, "BCWS"
                                        
                                    Else
                                
                                        Select Case tAssign.ResourceType
                                        
                                            Case pjResourceTypeWork
                                        
                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                                                                
                                            Case pjResourceTypeCost
                                        
                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                                        
                                            Case pjResourceTypeMaterial
                                            
                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                                        
                                        End Select
                                        
                                    End If
                                    
                                End If
                                
                            Next tAssign
                            
                        End If
                    
                    End If
                    
                End If
next_task:

            Next t
    
        End If
        
        If ActFound = True Then
            For i = 1 To UBound(ACTarray)
                
                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                
            Next i
        End If
        
        Close #1
        Close #2
        
    End If
        
End Sub
Private Function SetDirectory(ByVal ProjName As String) As String
    Dim newDir As String
    Dim pathDesktop As String
    
    pathDesktop = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    newDir = pathDesktop & "\" & RemoveIllegalCharacters(ProjName) & "_" & Format(Now, "YYYYMMDD HHMMSS")
    
    MkDir newDir
    SetDirectory = newDir
    Exit Function
    
End Function
Private Sub DeleteDirectory(ByVal DirName As String)

    RmDir DirName & "\"
    Exit Sub
    
End Sub

Private Function CleanCamName(ByVal CAM As String) As String

    If InStr(CAM, ".") > 0 Then
        CleanCamName = Right(CAM, Len(CAM) - InStr(CAM, "."))
    Else
        CleanCamName = CAM
        Exit Function
    End If

End Function

Private Function Find_BCRs(ByVal curProj As Project, ByVal fWP As String, ByVal fBCR As String, ByVal BCRnum As String) As Integer

    Dim t As Task
    Dim i As Integer
    Dim x As Integer
    Dim tempBCRstr As String
    Dim subProjs As Subprojects
    Dim subProj As Subproject
    Dim curSProj As Project
    
    i = 0
    
    If curProj.Subprojects.count > 0 Then
    
        Set subProjs = curProj.Subprojects
        
        For Each subProj In subProjs
        
            FileOpen Name:=subProj.Path, ReadOnly:=True
            
            Set curSProj = ActiveProject
    
            For Each t In curSProj.Tasks
            
                If Not t Is Nothing Then
                
                    tempBCRstr = t.GetField(FieldNameToFieldConstant(fBCR))
                    
                    If InStr(tempBCRstr, BCRnum) > 0 Then
                        
                        If i = 0 Then
                            i = 1
                        End If
                        
                        If i = 1 Then
                        
                            ReDim BCR_WP(1 To i)
                            BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                            i = i + 1
                            
                        Else
                        
                            For x = 1 To UBound(BCR_WP)
                                If BCR_WP(x) = t.GetField(FieldNameToFieldConstant(fWP)) Then
                                    GoTo Next_SubProj_WPtask
                                End If
                            Next x
                            
                            ReDim Preserve BCR_WP(1 To i)
                            BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                            i = i + 1
            
                        End If
                        
                    End If
                End If
Next_SubProj_WPtask:
        
            Next t
            
        Next subProj
        
    Else
    
        For Each t In curProj.Tasks
            
            If Not t Is Nothing Then
                tempBCRstr = t.GetField(FieldNameToFieldConstant(fBCR))
                
                If InStr(tempBCRstr, BCRnum) > 0 Then
                    
                    If i = 0 Then
                        i = 1
                    End If
                    
                    If i = 1 Then
                    
                        ReDim BCR_WP(1 To i)
                        BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                        i = i + 1
                        
                    Else
                    
                        For x = 1 To UBound(BCR_WP)
                            If BCR_WP(x) = t.GetField(FieldNameToFieldConstant(fWP)) Then
                                GoTo Next_WPtask
                            End If
                        Next x
                        
                        ReDim Preserve BCR_WP(1 To i)
                        BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                        i = i + 1
        
                    End If
                    
                End If
                
            End If
            
Next_WPtask:
    
        Next t
        
    End If

    Find_BCRs = i

End Function

Private Function IsInArray(ByVal arr As Variant, ByVal WP_ID As String) As Boolean

    Dim i As Integer
    
    i = 1
    
    For i = 1 To UBound(arr)
        If arr(i) = WP_ID Then
            IsInArray = True
            Exit Function
        End If
    Next i
    
    IsInArray = False

End Function

Private Function RemoveIllegalCharacters(ByVal strText As String) As String

    Const cstrIllegals As String = "\,/,:,*,?,"",<,>,|"
    
    Dim lngCounter As Long
    Dim astrChars() As String
    
    astrChars() = Split(cstrIllegals, ",")
    
    For lngCounter = LBound(astrChars()) To UBound(astrChars())
        strText = Replace(strText, astrChars(lngCounter), vbNullString)
    Next lngCounter
    
    RemoveIllegalCharacters = strText

End Function

Private Sub ReadCustomFields(ByVal curProj As Project)
    
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

Private Function CleanNumber(ByVal NumStr As String) As String

    Dim i As Integer
    Dim newNumStr As String
    
    For i = 1 To Len(NumStr)
        
        If Mid(NumStr, i, 1) = "." Or IsNumeric(Mid(NumStr, i, 1)) Then
        
            newNumStr = newNumStr & Mid(NumStr, i, 1)
        
        End If
        
    Next
    
    CleanNumber = newNumStr

End Function
Private Function PercentfromString(ByVal inputStr As String) As Double

    Dim tempDbl As String

    'Test for % String
    If InStr(inputStr, "%") > 0 Then
    
        tempDbl = Left(inputStr, Len(inputStr) - 1)
        
    Else
    
        tempDbl = inputStr
    
    End If
    
    PercentfromString = CDbl(tempDbl)

End Function
Private Sub ExportTimeScaleResources(ByVal ID As String, ByVal t As Task, ByVal tAssign As Assignment, ResFile As Integer, exportType As String)

    Dim TSV As TimeScaleValue
    Dim TSVS As TimeScaleValues
    Dim tsvsa As TimeScaleValues
    Dim tsva As TimeScaleValue
    Dim tempWork As Double

    Select Case exportType
    
        Case "ETC"
        
            Select Case tAssign.ResourceType
            
                Case pjResourceTypeWork
                                    
                    If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then
                    
                        Set TSVS = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        Set tsvsa = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledActualWork, pjTimescaleWeeks)
                        
                        For Each TSV In TSVS
                        
                            Set tsva = tsvsa(TSV.Index)
                            
                            tempWork = 0
                            
                            If tsva <> "" Then
                                tempWork = CDbl(TSV.Value) - CLng(tsva.Value)
                            ElseIf TSV.Value <> "" Then
                                tempWork = CDbl(TSV.Value)
                            End If
                        
                            If tempWork <> 0 Then
                            
                                If TSVS.count = 1 Then
                                
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                
                                Else
                                
                                    Select Case TSV.Index
                                    
                                        Case 1
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                    
                                        Case TSVS.count
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                        Case Else
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                    
                                    End Select
                                    
                                End If
    
                            End If
                        
                        Next TSV
                        
                        Exit Sub
                                    
                    ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then
                    
                        Set TSVS = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        For Each TSV In TSVS
                        
                            If TSV.Value <> "" Then
                            
                                If TSVS.count = 1 Then
                                
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                Else
                                
                                    Select Case TSV.Index
                                    
                                        Case 1
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                    
                                        Case TSVS.count
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value / 60 & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                        Case Else
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value / 60 & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                    
                                    End Select
                                    
                                End If
                                
                            End If
                            
                        Next TSV
                        
                        Exit Sub
                    
                    End If
                
            Case pjResourceTypeCost
                    
                If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then
                
                    Set TSVS = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledCost, pjTimescaleWeeks)
                    Set tsvsa = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledActualCost, pjTimescaleWeeks)
                    
                    For Each TSV In TSVS
                    
                        Set tsva = tsvsa(TSV.Index)
                        
                        tempWork = 0
                        
                        If tsva <> "" Then
                            tempWork = CDbl(TSV.Value) - CLng(tsva.Value)
                        ElseIf TSV.Value <> "" Then
                            tempWork = CDbl(TSV.Value)
                        End If
                    
                        If tempWork <> 0 Then
                        
                            If TSVS.count = 1 Then
                            
                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                            
                            Else
                            
                                Select Case TSV.Index
                                
                                    Case 1
                                    
                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                
                                    Case TSVS.count
                                    
                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                        
                                    Case Else
                                    
                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                
                                End Select
                                
                            End If

                        End If
                    
                    Next TSV
                    
                    Exit Sub
                               
                ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then
                
                    Set TSVS = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledCost, pjTimescaleWeeks)
                    For Each TSV In TSVS
                    
                        If TSV.Value <> "" Then
                        
                            If TSVS.count = 1 Then
                            
                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                            
                            Else
                            
                                Select Case TSV.Index
                                
                                    Case 1
                                    
                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                
                                    Case TSVS.count
                                    
                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                        
                                    Case Else
                                    
                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                
                                End Select
                                
                            End If
                            
                        End If
                        
                    Next TSV
                    
                    Exit Sub
                
                End If
            
            Case pjResourceTypeMaterial
            
                If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then
                    
                        Set TSVS = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        Set tsvsa = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledActualWork, pjTimescaleWeeks)
                        
                        For Each TSV In TSVS
                        
                            Set tsva = tsvsa(TSV.Index)
                            
                            tempWork = 0
                            
                            If tsva <> "" Then
                                tempWork = CDbl(TSV.Value) - CLng(tsva.Value)
                            ElseIf TSV.Value <> "" Then
                                tempWork = CDbl(TSV.Value)
                            End If
                        
                            If tempWork <> 0 Then
                            
                                If TSVS.count = 1 Then
                                    
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                
                                Else
                                
                                    Select Case TSV.Index
                                    
                                        Case 1
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                    
                                        Case TSVS.count
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                        Case Else
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                    
                                    End Select
                                    
                                End If
    
                            End If
                        
                        Next TSV
                        
                        Exit Sub
                                    
                    ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then
                    
                        Set TSVS = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        For Each TSV In TSVS
                        
                            If TSV.Value <> "" Then
                            
                                If TSVS.count = 1 Then
                                
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                
                                Else
                                
                                    Select Case TSV.Index
                                    
                                        Case 1
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                    
                                        Case TSVS.count
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
                                            
                                        Case Else
                                        
                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                    
                                    End Select
                                    
                                End If
                                
                            End If
                            
                        Next TSV
                        
                        Exit Sub
                    
                    End If
            
            End Select
        
        Case "BCWS"
        
            Select Case tAssign.ResourceType
                
                Case pjResourceTypeWork
                                             
                    Set TSVS = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                    For Each TSV In TSVS
                    
                        If TSV.Value <> "" Then
                        
                            If TSVS.count = 1 Then
                            
                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                            
                            Else
                            
                                Select Case TSV.Index
                                
                                    Case 1
                                    
                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                
                                    Case TSVS.count
                                    
                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value / 60 & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                                        
                                    Case Else
                                    
                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value / 60 & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                                
                                End Select
                                
                            End If
                            
                        End If
                        
                    Next TSV
                    
                    Exit Sub
                
            Case pjResourceTypeCost
                
                Set TSVS = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineCost, pjTimescaleWeeks)
                For Each TSV In TSVS
                
                    If TSV.Value <> "" Then
                    
                        If TSVS.count = 1 Then
                        
                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                        
                        Else
                        
                            Select Case TSV.Index
                            
                                Case 1
                                
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                            
                                Case TSVS.count
                                
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                                    
                                Case Else
                                
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                            
                            End Select
                            
                        End If
                        
                    End If
                    
                Next TSV
                
                Exit Sub
            
            Case pjResourceTypeMaterial
            
                Set TSVS = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                For Each TSV In TSVS
                
                    If TSV.Value <> "" Then
                    
                        If TSVS.count = 1 Then
                        
                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                        
                        Else
                        
                            Select Case TSV.Index
                            
                                Case 1
                                
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                            
                                Case TSVS.count
                                
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")
                                    
                                Case Else
                                
                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & TSV.Value & "," & Format(TSV.startDate, "M/D/YYYY") & "," & Format(TSV.EndDate - 1, "M/D/YYYY")
                            
                            End Select
                        
                        End If
                        
                    End If
                    
                Next TSV
                
                Exit Sub
            
            End Select
        
        Case Else
        
            Exit Sub
        
    End Select

End Sub